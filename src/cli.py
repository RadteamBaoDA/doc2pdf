import atexit
import ctypes
import msvcrt
import os
import shutil
import sys
import threading
import time
import zipfile
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from dataclasses import dataclass, replace
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Optional, Tuple

# pythoncom needed for COM in threads
import pythoncom
import typer
from rich.console import Console
from rich.live import Live
from rich.panel import Panel
from rich.progress import (
    BarColumn,
    MofNCompleteColumn,
    Progress,
    SpinnerColumn,
    TaskProgressColumn,
    TextColumn,
    TimeElapsedColumn,
    TimeRemainingColumn,
)
from rich.table import Table
from rich.text import Text

from .config import (
    FileType,
    get_config_path,
    get_logging_config,
    get_parallel_config,
    get_pdf_handling_config,
    get_pdf_settings,
    get_post_processing_config,
    get_reporting_config,
    get_suffix_config,
    get_timeout_config,
    set_config_path,
)
from .core.job_runner import run_excel_job
from .core.macro_converter import SUPPORTED_FORMATS, MacroConverter
from .core.pdf_processor import PDFProcessor
from .core.powerpoint_converter import PowerPointConverter
from .core.word_converter import WordConverter
from .utils.logger import logger, setup_logger
from .utils.process_manager import ProcessRegistry
from .utils.timeout import TimeoutError, run_with_timeout
from .version import __version__


class RealtimeReportWriter:
    """
    Writes conversion reports in realtime.
    - Errors are written immediately when they occur
    - Successful files are tracked and written as they complete
    """
    
    def __init__(self, reports_dir: Path, input_path: Path, output_path: Path, timestamp: str):
        self.reports_dir = reports_dir
        self.input_path = input_path
        self.output_path = output_path
        self.timestamp = timestamp
        self.error_count = 0
        self.success_count = 0
        self.skipped_count = 0
        self._lock = threading.Lock()
        
        # Create reports directory
        self.reports_dir.mkdir(parents=True, exist_ok=True)
        
        # Initialize report files with headers
        self._init_error_report()
        self._init_summary_report()
    
    def _init_error_report(self):
        """Initialize error report file with header."""
        self.error_path = self.reports_dir / f"error_{self.timestamp}.txt"
        with open(self.error_path, "w", encoding="utf-8") as f:
            f.write(f"doc2pdf Error Report (Realtime)\n")
            f.write(f"{'='*50}\n")
            f.write(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Input: {self.input_path}\n")
            f.write(f"Output: {self.output_path}\n\n")
            f.write(f"Errors:\n")
            f.write(f"{'-'*50}\n")
    
    def _init_summary_report(self):
        """Initialize summary report file with header."""
        self.summary_path = self.reports_dir / f"summary_{self.timestamp}.txt"
        with open(self.summary_path, "w", encoding="utf-8") as f:
            f.write(f"doc2pdf Conversion Summary (Realtime)\n")
            f.write(f"{'='*50}\n")
            f.write(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Input: {self.input_path}\n")
            f.write(f"Output: {self.output_path}\n\n")
            f.write(f"Successfully Converted Files:\n")
            f.write(f"{'-'*50}\n")
    
    def write_error(self, input_file: Path, output_file: Path, error_msg: str):
        """Write an error entry immediately to the error report."""
        with self._lock:
            self.error_count += 1
            with open(self.error_path, "a", encoding="utf-8") as f:
                f.write(f"\n[{self.error_count}] {input_file.name}\n")
                f.write(f"    Time:   {datetime.now().strftime('%H:%M:%S')}\n")
                f.write(f"    Input:  {input_file}\n")
                f.write(f"    Output: {output_file}\n")
                f.write(f"    Error:  {error_msg}\n")
    
    def write_success(
        self,
        input_file: Path,
        output_file: Path,
        file_type: str,
        duration_seconds: Optional[float] = None,
    ):
        """Write a success entry immediately to the summary report."""
        with self._lock:
            self.success_count += 1
            with open(self.summary_path, "a", encoding="utf-8") as f:
                f.write(f"[{self.success_count}] {input_file.name}\n")
                f.write(f"    Time:   {datetime.now().strftime('%H:%M:%S')}\n")
                f.write(f"    Type:   {file_type}\n")
                if duration_seconds is not None:
                    f.write(f"    Duration: {duration_seconds:.3f}s\n")
                f.write(f"    Input:  {input_file}\n")
                f.write(f"    Output: {output_file}\n\n")
    
    def write_skipped(self, input_file: Path, reason: str):
        """Track skipped file."""
        with self._lock:
            self.skipped_count += 1

    def write_excel_scheduling(
        self,
        *,
        file_count: int,
        resolved_workers: int,
        configured_workers: int | str,
        worker_cap: int,
        logical_cpus: Optional[int],
        available_memory_mb: Optional[int],
    ) -> None:
        """Record the adaptive scheduling decision in the batch summary."""
        with self._lock:
            with open(self.summary_path, "a", encoding="utf-8") as stream:
                stream.write("Excel Scheduling:\n")
                stream.write(f"    Files: {file_count}\n")
                stream.write(f"    Resolved workers: {resolved_workers}\n")
                stream.write(f"    Configured workers: {configured_workers}\n")
                stream.write(f"    Worker cap: {worker_cap}\n")
                stream.write(f"    Logical CPUs: {logical_cpus or 'unknown'}\n")
                stream.write(
                    f"    Available memory MB: "
                    f"{available_memory_mb if available_memory_mb is not None else 'unknown'}\n\n"
                )
    
    def finalize(self, total_files: int):
        """Write final summary statistics to both reports."""
        end_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Finalize summary report
        with open(self.summary_path, "a", encoding="utf-8") as f:
            f.write(f"\n{'-'*50}\n")
            f.write(f"Completed: {end_time}\n\n")
            f.write(f"Final Results:\n")
            f.write(f"  Success: {self.success_count}\n")
            f.write(f"  Failed:  {self.error_count}\n")
            f.write(f"  Skipped: {self.skipped_count}\n")
            f.write(f"  Total:   {total_files}\n")
        
        # Finalize error report
        with open(self.error_path, "a", encoding="utf-8") as f:
            f.write(f"\n{'-'*50}\n")
            f.write(f"Completed: {end_time}\n")
            f.write(f"Total Errors: {self.error_count}\n")
        
        # Remove error report if no errors occurred
        if self.error_count == 0:
            try:
                self.error_path.unlink()
            except:
                pass
            return None
        
        return self.error_path

app = typer.Typer(
    name="doc2pdf",
    help="""
    [bold]doc2pdf[/bold] - Convert Microsoft Office documents to PDF.
    
    [bold]Features:[/bold]
    - Batch conversion of folders
    - Support for Word, Excel, and PowerPoint (Configuration)
    - Configurable settings via pattern matching
    - Detailed logging to file and console
    
    [bold]Logging:[/bold]
    Logs are written to the console and to files in the `logs/` directory.
    Check `config.yml` for log rotation settings.
    """,
    add_completion=False,
)
console = Console()


@dataclass(frozen=True)
class FileConversionResult:
    input_file: Path
    output_file: Optional[Path]
    file_type: str
    status: str
    error: Optional[str] = None
    duration_seconds: Optional[float] = None


def get_available_memory_mb() -> Optional[int]:
    """Return currently available physical memory without adding a dependency."""
    if os.name != "nt":
        return None

    class MemoryStatus(ctypes.Structure):
        _fields_ = [
            ("length", ctypes.c_ulong),
            ("memory_load", ctypes.c_ulong),
            ("total_physical", ctypes.c_ulonglong),
            ("available_physical", ctypes.c_ulonglong),
            ("total_page_file", ctypes.c_ulonglong),
            ("available_page_file", ctypes.c_ulonglong),
            ("total_virtual", ctypes.c_ulonglong),
            ("available_virtual", ctypes.c_ulonglong),
            ("available_extended_virtual", ctypes.c_ulonglong),
        ]

    status = MemoryStatus()
    status.length = ctypes.sizeof(status)
    try:
        if not ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
            return None
    except (AttributeError, OSError):
        return None
    return int(status.available_physical // (1024 * 1024))


def estimate_excel_work(path: Path) -> int:
    """Estimate workbook cost cheaply so long jobs enter the queue first."""
    try:
        size = path.stat().st_size
    except OSError:
        size = 0
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return size
    try:
        with zipfile.ZipFile(path) as package:
            worksheet_bytes = sum(
                item.file_size
                for item in package.infolist()
                if item.filename.startswith("xl/worksheets/")
                and item.filename.endswith(".xml")
            )
            return size + worksheet_bytes
    except (OSError, zipfile.BadZipFile):
        return size


def is_transient_excel_failure(result: FileConversionResult) -> bool:
    """Classify retryable infrastructure failures, never quality failures."""
    if result.status != "failed" or not result.error:
        return False
    message = result.error.casefold()
    deterministic = (
        "postflight failed",
        "quality floor",
        "missing sentinel",
        "unexpectedly blank",
        "page count",
        "invalid authored",
        "no printable content",
        "unsupported",
    )
    if any(marker in message for marker in deterministic):
        return False
    transient = (
        "call was rejected by callee",
        "rpc server",
        "disconnected from its clients",
        "printer is busy",
        "server busy",
        "not enough memory",
        "out of memory",
        "resource temporarily unavailable",
    )
    return any(marker in message for marker in transient)


def run_parallel_excel_jobs(
    files: List[Path],
    worker: Callable[[Path], FileConversionResult],
    *,
    max_workers: int,
    cancel_event: threading.Event,
    on_complete: Callable[[FileConversionResult], None],
) -> None:
    """Run isolated Excel jobs concurrently and report results centrally."""
    if not files:
        return
    ordered_files = sorted(
        files,
        key=lambda path: (-estimate_excel_work(path), str(path).casefold()),
    )
    if max_workers == 1:
        for file_path in ordered_files:
            if cancel_event.is_set():
                break
            on_complete(worker(file_path))
        return

    executor = ThreadPoolExecutor(
        max_workers=max_workers,
        thread_name_prefix="excel-job",
    )
    futures = {}
    try:
        futures = {executor.submit(worker, path): path for path in ordered_files}
        pending = set(futures)
        transient_retries: List[Path] = []
        while pending:
            if cancel_event.is_set():
                for future in pending:
                    future.cancel()
                break
            done, pending = wait(pending, timeout=0.1, return_when=FIRST_COMPLETED)
            for future in done:
                result = future.result()
                if is_transient_excel_failure(result):
                    transient_retries.append(futures[future])
                    logger.warning(
                        f"Retrying transient Excel failure serially: "
                        f"{futures[future].name}"
                    )
                else:
                    on_complete(result)
    finally:
        executor.shutdown(wait=True, cancel_futures=True)
    for file_path in transient_retries:
        if cancel_event.is_set():
            break
        on_complete(worker(file_path))



def version_callback(value: bool):
    if value:
        console.print(f"[bold green]doc2pdf[/bold green] version: {__version__}")
        raise typer.Exit()

@app.callback(invoke_without_command=True)
def main(
    ctx: typer.Context,
    version: Optional[bool] = typer.Option(
        None,
        "--version",
        "-v",
        help="Show the application version and exit.",
        callback=version_callback,
        is_eager=True,
    ),
):
    """
    doc2pdf - Convert your documents to PDF with ease.
    """
    if ctx.invoked_subcommand is None:
        console.print(ctx.get_help())

def get_files(path: Path) -> List[Path]:
    if path.is_file():
        return [path]
    
    extensions = {
        "*.docx", "*.doc", 
        "*.xlsx", "*.xls", "*.xlsm", "*.xlsb",
        "*.pptx", "*.ppt",
        "*.pdf"
    }
    
    files = []
    for ext in extensions:
        files.extend(list(path.rglob(ext)))
    return sorted(files)

def get_file_type(path: Path) -> FileType:
    ext = path.suffix.lower()
    if ext in [".docx", ".doc"]:
        return "word"
    elif ext in [".xlsx", ".xls", ".xlsm", ".xlsb"]:
        return "excel"
    elif ext in [".pptx", ".ppt"]:
        return "powerpoint"
    elif ext == ".pdf":
        return "pdf"
    return "word" # Default fallback


@app.command("convert-macros")
def convert_macros(
    input_path: Path = typer.Argument(..., help="Macro-enabled Office file or directory", exists=True),
    output_path: Path = typer.Option(Path("output"), "--output", "-o", help="Output file or directory"),
):
    """Convert .docm/.pptm/.xlsm files to .docx/.pptx/.xlsx (remove macros)."""
    files = (
        [input_path]
        if input_path.is_file()
        else sorted(
            path
            for path in input_path.rglob("*")
            if path.is_file() and path.suffix.lower() in SUPPORTED_FORMATS
        )
    )
    if input_path.is_file() and input_path.suffix.lower() not in SUPPORTED_FORMATS:
        raise typer.BadParameter("Input must be a .docm, .pptm, or .xlsm file")
    if not files:
        console.print(f"[yellow]No .docm, .pptm, or .xlsm files found in {input_path}.[/yellow]")
        raise typer.Exit()

    converter = MacroConverter()
    failures = 0
    for source in files:
        target_extension = SUPPORTED_FORMATS[source.suffix.lower()][0]
        if input_path.is_dir():
            target = output_path / source.relative_to(input_path).with_suffix(target_extension)
        elif output_path.suffix.lower() == target_extension:
            target = output_path
        else:
            target = output_path / source.with_suffix(target_extension).name
        try:
            result = converter.convert(source, target)
            console.print(f"[green]Converted:[/green] {source} -> {result}")
        except Exception as error:
            failures += 1
            console.print(f"[red]Failed:[/red] {source}: {error}")

    if failures:
        raise typer.Exit(code=1)

@app.command()
def convert(
    input_path: Path = typer.Argument(Path("input"), help="Path to the input file or directory", exists=True),
    output_path: Optional[Path] = typer.Option(Path("output"), "--output", "-o", help="Path to the output PDF or Directory"),
    config_path: Optional[Path] = typer.Option(None, "--config", "-c", help="Path to configuration file", exists=True, dir_okay=False),
    verbose: bool = typer.Option(False, "--verbose", help="Enable verbose logging"),
    trim: Optional[bool] = typer.Option(None, "--trim/--no-trim", help="Trim whitespace from output PDF (overrides config.yml)"),
    trim_margin: Optional[float] = typer.Option(None, "--trim-margin", help="Margin in points when trimming (default: 10)"),
):
    """
    Convert a document or a directory of documents to PDF.
    
    Defaults:
    - Input: ./input
    - Output: ./output
    
    Supports Word (.doc, .docx), Excel (.xls, .xlsx), and PowerPoint (.ppt, .pptx).
    """
    # Resolve a concrete destination even for programmatic invocations that pass
    # None, so suffix and post-processing rules never operate on a missing path.
    output_path = output_path or Path("output")
    # Register cleanup on exit
    atexit.register(ProcessRegistry.kill_all)
    
    # Configure config path if provided
    if config_path:
        set_config_path(config_path)

    # Load config (refresh in case path changed)
    config = get_logging_config()


    # Configure verbose logging
    current_config = config.copy()
    if verbose:
        current_config["level"] = "DEBUG"
    
    # Capture console handler ID to remove it later during TUI to prevent flashing
    console_handler_id = setup_logger(current_config)

    # Log config path
    logger.info(f"Using configuration file: {get_config_path().resolve()}")

    files = get_files(input_path)
    
    if not files:
        console.print(f"[yellow]No supported Office documents found in {input_path}.[/yellow]")
        raise typer.Exit()

    # Get post-processing settings (CLI overrides config)
    post_proc_config = get_post_processing_config()
    should_trim = trim if trim is not None else post_proc_config.trim_whitespace.enabled
    trim_margin_value = trim_margin if trim_margin is not None else post_proc_config.trim_whitespace.margin
    
    # Get timeout settings
    timeout_config = get_timeout_config()
    document_timeout = timeout_config.document_parsing
    excel_trim_timeout = timeout_config.excel_trim
    parallel_config = get_parallel_config()

    # TUI Setup
    from .tui import LogBuffer, TUIContext
    
    log_buffer = LogBuffer()
    tui_ctx = TUIContext(log_buffer)
    
    # Redirect Logger to TUI Buffer
    def tui_sink(message):
        record = message.record
        level_name = record['level'].name
        colors = { "INFO": "green", "WARNING": "yellow", "ERROR": "bold red", "CRITICAL": "bold white on red", "DEBUG": "cyan" }
        color = colors.get(level_name, "white")
        log_msg = f"[{color}]{record['time'].strftime('%H:%M:%S')} | {level_name: <8} | {record['message']}[/{color}]"
        log_buffer.write(log_msg)
    
    try:
        tui_level = current_config.get("level", "INFO")
        logger.add(tui_sink, format="{message}", level=tui_level)
        if console_handler_id is not None:
            logger.remove(console_handler_id)
    except Exception:
        pass

    # Initialize Progress (passive)
    progress = Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
        TimeRemainingColumn(),
        expand=True
    )

    cancel_event = threading.Event()

    # Define worker function for threading
    def conversion_worker():
        # COM initialization for thread
        pythoncom.CoInitialize()
        try:
            # Initialize converters inside thread to ensure COM affinity
            word_converter = WordConverter()
            ppt_converter = PowerPointConverter()
            pdf_processor = PDFProcessor()
            
            def convert_one(file_path: Path) -> FileConversionResult:
                file_type = get_file_type(file_path)

                # Get settings
                # base_path is the root input directory (either a folder or file's parent)
                base_path = input_path if input_path.is_dir() else input_path.parent
                settings = get_pdf_settings(input_path=file_path, file_type=file_type, base_path=base_path)
                suffix_config = get_suffix_config()
                suffix = suffix_config.get(file_type, "")
                
                # Determine output
                if output_path:
                    if input_path.is_dir():
                        rel_path = file_path.relative_to(input_path)
                        base_name = rel_path.stem + suffix + ".pdf"
                        target_file = output_path / rel_path.parent / base_name
                        target_file.parent.mkdir(parents=True, exist_ok=True)
                    else:
                        if output_path.suffix.lower() == ".pdf":
                            target_file = output_path
                        else:
                            base_name = file_path.stem + suffix + ".pdf"
                            target_file = output_path / base_name
                            target_file.parent.mkdir(parents=True, exist_ok=True)
                else:
                    target_file = None 

                try:
                    def progress_callback(amount: float):
                        # Per-workbook progress arrives on executor threads. The
                        # main TUI renders logs and aggregate completion only.
                        return None

                    def child_log(level_name: str, text: str):
                        # Re-emit in the parent so it reaches the file log and
                        # the TUI's thread-safe log buffer.
                        logger.log(level_name, text)
                    
                    converted_pdf = None
                    
                    if file_type == "word":
                        converted_pdf = run_with_timeout(
                            word_converter.convert,
                            document_timeout,
                            file_path, target_file, settings, base_path=base_path
                        )
                    elif file_type == "powerpoint":
                        converted_pdf = run_with_timeout(
                            ppt_converter.convert,
                            document_timeout,
                            file_path, target_file, settings, base_path=base_path
                        )
                    elif file_type == "excel":
                        if target_file is None:
                            target_file = file_path.with_name(
                                file_path.stem + suffix + ".pdf"
                            )
                        trim_this_file = should_trim and (
                            trim is True or file_type in post_proc_config.trim_whitespace.include
                        )
                        trim_cfg = post_proc_config.trim_whitespace
                        trim_options = None
                        if trim_this_file:
                            trim_options = {
                                "margin": trim_margin_value,
                                "box_mode": trim_cfg.box_mode,
                                "render_dpi": trim_cfg.render_dpi,
                                "max_render_pixels": trim_cfg.max_render_pixels,
                                "background_tolerance": trim_cfg.background_tolerance,
                                "include_annotations": trim_cfg.include_annotations,
                                "allow_signature_invalidation": trim_cfg.allow_signature_invalidation,
                            }
                        combined_timeout = None
                        if document_timeout or (trim_this_file and excel_trim_timeout):
                            combined_timeout = (document_timeout or 0) + (
                                (excel_trim_timeout or 0) if trim_this_file else 0
                            )
                        converted_pdf = run_excel_job(
                            file_path, target_file, settings,
                            trim_options=trim_options,
                            timeout_seconds=combined_timeout,
                            base_path=base_path,
                            on_log=child_log,
                            on_progress=progress_callback,
                            cancel_event=cancel_event,
                            runtime_evidence={
                                "resolved_excel_workers": resolved_excel_workers,
                                "configured_excel_workers": parallel_config.excel_workers,
                                "excel_worker_cap": parallel_config.excel_worker_cap,
                                "logical_cpus": logical_cpus,
                                "available_memory_mb": available_memory_mb,
                            },
                        )
                    elif file_type == "pdf":
                        # Log full path
                        logger.info(f"Input PDF found: {file_path}")
                        
                        pdf_handling = get_pdf_handling_config()
                        if pdf_handling.copy_to_output and target_file:
                             # Logic to copy
                             shutil.copy2(file_path, target_file)
                             logger.info(f"Copied PDF to: {target_file}")
                             converted_pdf = target_file
                        else:
                             # Just skip or count as success? 
                             # If we don't copy, we essentially "skipped" processing it, but it was "handled".
                             # But let's count as skipped if not copied, or success if we just wanted to log it?
                             # Requirement: "when input have pdf, write input full path of this pdf file."
                             # So we always do that.
                             # If copy is disabled, we effectively did nothing else.
                             # Let's count as skipped-by-policy or success? 
                             # Let's count as success because we "handled" it as per config (logging).
                             # But "skipped" might be more valuable for user stats.
                             if not pdf_handling.copy_to_output:
                                 logger.debug(f"PDF copy disabled. Skipping copy for {file_path.name}")
                                 return FileConversionResult(
                                     file_path, target_file, file_type, "skipped"
                                 )
                             else:
                                 # This branch is for when target_file is None (dry run?) or copy succeeded
                                 pass 

                    else:
                        logger.warning(f"Conversion for {file_type} not supported. Skipping {file_path.name}")
                        return FileConversionResult(
                            file_path, target_file, file_type, "skipped"
                        )
                    
                    if converted_pdf and file_type != "excel" and should_trim and converted_pdf.exists():
                        # Check if file type is included in trim settings
                        if trim is True or file_type in post_proc_config.trim_whitespace.include:
                            try:
                                # Apply timeout specifically for Excel trimming
                                trim_timeout = excel_trim_timeout if file_type == "excel" else None
                                trim_cfg = post_proc_config.trim_whitespace
                                run_with_timeout(
                                    pdf_processor.trim_whitespace, trim_timeout,
                                    converted_pdf, margin=trim_margin_value,
                                    box_mode=trim_cfg.box_mode,
                                    render_dpi=trim_cfg.render_dpi,
                                    max_render_pixels=trim_cfg.max_render_pixels,
                                    background_tolerance=trim_cfg.background_tolerance,
                                    include_annotations=trim_cfg.include_annotations,
                                    allow_signature_invalidation=trim_cfg.allow_signature_invalidation,
                                )
                            except TimeoutError:
                                raise
                            except Exception as trim_err:
                                raise RuntimeError(
                                    f"PDF trimming failed for {converted_pdf.name}: {trim_err}"
                                ) from trim_err
                        else:
                            logger.debug(f"Skipping trim for {file_type} file: {file_path.name}")

                    if converted_pdf:
                        return FileConversionResult(
                            file_path, converted_pdf, file_type, "success"
                        )
                    return FileConversionResult(
                        file_path, target_file, file_type, "skipped"
                    )
                        
                except TimeoutError as timeout_err:
                    error_msg = f"Conversion timed out: {timeout_err}"
                    logger.error(f"Failed to convert {file_path.name}: {error_msg}")
                    return FileConversionResult(
                        file_path, target_file, file_type, "failed", error_msg
                    )
                except Exception as e:
                    logger.error(f"Failed to convert {file_path}: {e}")
                    return FileConversionResult(
                        file_path, target_file, file_type, "failed", str(e)
                    )

            def safe_convert_one(file_path: Path) -> FileConversionResult:
                started = time.perf_counter()
                try:
                    result = convert_one(file_path)
                except Exception as exc:
                    file_type = get_file_type(file_path)
                    logger.error(f"Failed to prepare conversion for {file_path}: {exc}")
                    result = FileConversionResult(
                        file_path, None, file_type, "failed", str(exc)
                    )
                return replace(
                    result,
                    duration_seconds=time.perf_counter() - started,
                )

            def record_result(result: FileConversionResult) -> None:
                nonlocal success_count, fail_count, skipped_count
                if result.status == "success":
                    success_count += 1
                    if report_writer and result.output_file is not None:
                        report_writer.write_success(
                            result.input_file,
                            result.output_file,
                            result.file_type,
                            result.duration_seconds,
                        )
                elif result.status == "failed":
                    fail_count += 1
                    error_msg = result.error or "Unknown conversion failure"
                    failed_files.append(
                        (result.input_file, result.output_file, error_msg)
                    )
                    if report_writer:
                        report_writer.write_error(
                            result.input_file, result.output_file, error_msg
                        )
                else:
                    skipped_count += 1
                    if report_writer:
                        report_writer.write_skipped(
                            result.input_file, "Skipped by conversion policy"
                        )
                progress.advance(task_id, advance=1)
                tui_ctx.update_progress(progress)

            # Other Office applications retain their existing serial COM lane.
            serial_files = [path for path in files if get_file_type(path) != "excel"]
            excel_files = [path for path in files if get_file_type(path) == "excel"]
            logical_cpus = os.cpu_count()
            available_memory_mb = get_available_memory_mb()
            resolved_excel_workers = parallel_config.resolve_excel_workers(
                len(excel_files),
                logical_cpus=logical_cpus,
                available_memory_mb=available_memory_mb,
            )
            if report_writer:
                report_writer.write_excel_scheduling(
                    file_count=len(excel_files),
                    resolved_workers=resolved_excel_workers,
                    configured_workers=parallel_config.excel_workers,
                    worker_cap=parallel_config.excel_worker_cap,
                    logical_cpus=logical_cpus,
                    available_memory_mb=available_memory_mb,
                )
            for file_path in serial_files:
                if cancel_event.is_set():
                    break
                progress.update(
                    task_id,
                    description=(
                        f"[cyan]Converting ({get_file_type(file_path)}): "
                        f"{file_path.name}"
                    ),
                )
                record_result(safe_convert_one(file_path))

            if not cancel_event.is_set():
                progress.update(
                    task_id,
                    description=(
                        f"[cyan]Converting {len(excel_files)} Excel file(s) "
                        f"with {resolved_excel_workers} worker(s)"
                    ),
                )
                logger.info(
                    f"Converting {len(excel_files)} Excel file(s) with "
                    f"{resolved_excel_workers} worker(s) "
                    f"(configured={parallel_config.excel_workers!r}, "
                    f"cap={parallel_config.excel_worker_cap})"
                )
                run_parallel_excel_jobs(
                    excel_files,
                    safe_convert_one,
                    max_workers=resolved_excel_workers,
                    cancel_event=cancel_event,
                    on_complete=record_result,
                )

        finally:
            pythoncom.CoUninitialize()

    # Initialize realtime report writer
    reporting_config = get_reporting_config()
    report_writer = None
    if reporting_config.enabled:
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        reports_dir = Path(reporting_config.reports_dir)
        report_writer = RealtimeReportWriter(reports_dir, input_path, output_path, timestamp)
    
    try:
        with Live(tui_ctx.layout, refresh_per_second=10, screen=True) as live:
            task_id = progress.add_task(f"[cyan]Converting {len(files)} files...", total=len(files))
            
            success_count = 0
            fail_count = 0
            skipped_count = 0
            failed_files: List[Tuple[Path, Path, str]] = []  # (input, output, error)
            
            tui_ctx.update_progress(progress)
            
            # Start Worker Thread
            worker_thread = threading.Thread(target=conversion_worker, daemon=True)
            worker_thread.start()
            
            # Main Loop: Handle Inputs and TUI Refresh
            while worker_thread.is_alive():
                # Rich Live must be updated from the owning/main thread. The
                # logging sink only appends to LogBuffer; this loop renders
                # newly-arrived logs at most 20 times per second.
                if log_buffer.consume_changed():
                    tui_ctx.update_logs()
                # Check for key press (Windows only)
                if msvcrt.kbhit():
                    key = msvcrt.getch()
                    if key == b'\xe0': # Special key prefix
                        code = msvcrt.getch()
                        if code == b'H': # Up Arrow
                             log_buffer.scroll_up()
                             tui_ctx.update_logs()
                        elif code == b'P': # Down Arrow
                             log_buffer.scroll_down()
                             tui_ctx.update_logs()
                
                # Small sleep to prevent CPU spinning
                time.sleep(0.05)
                # live.refresh() # handled by refresh_per_second, but explicit update helps responsiveness
            
            # Thread finished
            worker_thread.join()
            # Render events emitted immediately before the worker completed.
            if log_buffer.consume_changed():
                tui_ctx.update_logs()

    except KeyboardInterrupt:
        logger.warning("Conversion cancelled by user.")
        cancel_event.set()
        if 'worker_thread' in locals() and worker_thread.is_alive():
            worker_thread.join(timeout=15)
        ProcessRegistry.kill_all()
        console.print("[bold red]Conversion cancelled by user.[/bold red]")
        sys.exit(130)
            
    # Remove TUI Sink cleanup (optional, but good practice)
    # logger.remove(sink_id) # Hard to get ID without return value from add.
    
    # Summary
    # Check if console is safe to clear (might have been closed by Live context exit)
    # Live context restores terminal.
    console.clear() 
    console.print(Panel(Text(" Conversion Completed ", style="bold green"), style="green"))

    table = Table(title="Conversion Summary")
    table.add_column("Status", style="bold")
    table.add_column("Count")
    
    table.add_row("[green]Success[/green]", str(success_count))
    table.add_row("[red]Failed[/red]", str(fail_count))
    table.add_row("[yellow]Skipped[/yellow]", str(skipped_count))
    table.add_row("Total", str(len(files)))
    
    console.print(table)
    console.print(f"Logs available in: [bold]{current_config['file'].get('path', 'logs/')}[/bold]")
    
    # Finalize realtime reports
    if report_writer:
        report_writer.finalize(len(files))
        console.print(f"Summary report: [bold]{report_writer.summary_path}[/bold]")
        if report_writer.error_count > 0:
            console.print(f"Error report: [bold]{report_writer.error_path}[/bold]")
    
    # Copy error files to separate folder (preserving input folder structure)
    if reporting_config.enabled and reporting_config.copy_error_files.enabled and failed_files:
        errors_dir = output_path / reporting_config.copy_error_files.target_dir
        errors_dir.mkdir(parents=True, exist_ok=True)
        for input_file, _, _ in failed_files:
            try:
                # Preserve folder structure relative to input_path
                if input_path.is_dir():
                    rel_path = input_file.relative_to(input_path)
                    dest = errors_dir / rel_path
                    dest.parent.mkdir(parents=True, exist_ok=True)
                else:
                    dest = errors_dir / input_file.name
                shutil.copy2(input_file, dest)
            except Exception as copy_err:
                logger.warning(f"Could not copy error file {input_file.name}: {copy_err}")
        console.print(f"Error files copied to: [bold]{errors_dir}[/bold]")
    
if __name__ == "__main__":
    app()
