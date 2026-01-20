import typer
import sys
import time
import shutil
import threading
import msvcrt
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Tuple, Dict
# pythoncom needed for COM in threads
import pythoncom
from rich.console import Console
from rich.panel import Panel
from rich.progress import (
    Progress, SpinnerColumn, TextColumn, BarColumn, 
    TaskProgressColumn, TimeElapsedColumn, TimeRemainingColumn,
    MofNCompleteColumn
)
from rich.logging import RichHandler
from rich.layout import Layout
from rich.live import Live
from rich.table import Table
from rich.text import Text

from .version import __version__
from .core.word_converter import WordConverter
from .core.powerpoint_converter import PowerPointConverter
from .core.excel_converter import ExcelConverter
from .core.pdf_processor import PDFProcessor
from .utils.logger import setup_logger, logger
from .config import get_logging_config, get_pdf_settings, get_suffix_config, get_reporting_config, get_post_processing_config, FileType

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

config = get_logging_config()

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
        "*.pptx", "*.ppt"
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
    return "word" # Default fallback

@app.command()
def convert(
    input_path: Path = typer.Argument(Path("input"), help="Path to the input file or directory", exists=True),
    output_path: Optional[Path] = typer.Option(Path("output"), "--output", "-o", help="Path to the output PDF or Directory"),
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
    
    # Configure verbose logging
    current_config = config.copy()
    if verbose:
        current_config["level"] = "DEBUG"
    
    # Capture console handler ID to remove it later during TUI to prevent flashing
    console_handler_id = setup_logger(current_config)

    files = get_files(input_path)
    
    if not files:
        console.print(f"[yellow]No supported Office documents found in {input_path}.[/yellow]")
        raise typer.Exit()

    # Initialize converters
    word_converter = WordConverter()
    ppt_converter = PowerPointConverter()
    excel_converter = ExcelConverter()
    
    # Get post-processing settings (CLI overrides config)
    post_proc_config = get_post_processing_config()
    should_trim = trim if trim is not None else post_proc_config.trim_whitespace.enabled
    trim_margin_value = trim_margin if trim_margin is not None else post_proc_config.trim_whitespace.margin

    # TUI Setup
    from .tui import LogBuffer, TUIContext
    
    log_buffer = LogBuffer()
    tui_ctx = TUIContext(log_buffer)
    
    # Redirect Logger to TUI Buffer
    def tui_sink(message):
        record = message.record
        level_name = record['level'].name
        colors = { "INFO": "green", "WARNING": "yellow", "ERROR": "bold red", "CRITICAL": "bold white on red", "DEBUG": "dim blue" }
        color = colors.get(level_name, "white")
        log_msg = f"[{color}]{record['time'].strftime('%H:%M:%S')} | {level_name: <8} | {record['message']}[/{color}]"
        log_buffer.write(log_msg)
        tui_ctx.update_logs()
    
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

    # Define worker function for threading
    def conversion_worker():
        # COM initialization for thread
        pythoncom.CoInitialize()
        try:
            # Initialize converters inside thread to ensure COM affinity
            word_converter = WordConverter()
            ppt_converter = PowerPointConverter()
            excel_converter = ExcelConverter()
            pdf_processor = PDFProcessor()
            
            nonlocal success_count, fail_count, skipped_count
            
            for file_path in files:
                file_type = get_file_type(file_path)
                progress.update(task_id, description=f"[cyan]Converting ({file_type}): {file_path.name}")
                # Note: tui_ctx.update_progress is called by main loop or here? 
                # Ideally main loop updates TUI. 
                # But we need real-time log updates.
                # logs trigger update_logs automatically now via sink.
                # We should update progress here too? 
                # No, main loop refreshes Live context. 
                # But if we want instant feedback on "Converting..." text change, we can force update.
                tui_ctx.update_progress(progress)
                
                # Get settings
                settings = get_pdf_settings(input_path=file_path, file_type=file_type)
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
                        progress.advance(task_id, advance=amount)
                        tui_ctx.update_progress(progress)
                    
                    converted_pdf = None
                    
                    if file_type == "word":
                        word_converter.convert(file_path, target_file, settings)
                        converted_pdf = target_file
                        success_count += 1
                        progress.advance(task_id, advance=1)
                    elif file_type == "powerpoint":
                        ppt_converter.convert(file_path, target_file, settings)
                        converted_pdf = target_file
                        success_count += 1
                        progress.advance(task_id, advance=1)
                    elif file_type == "excel":
                        excel_converter.convert(file_path, target_file, settings, on_progress=progress_callback)
                        converted_pdf = target_file
                        success_count += 1
                    else:
                        logger.warning(f"Conversion for {file_type} not supported. Skipping {file_path.name}")
                        skipped_count += 1
                        progress.advance(task_id, advance=1)
                    
                    if converted_pdf and should_trim and converted_pdf.exists():
                        # Check if file type is included in trim settings
                        if file_type in post_proc_config.trim_whitespace.include:
                            try:
                                pdf_processor.trim_whitespace(converted_pdf, margin=trim_margin_value)
                            except Exception as trim_err:
                                logger.warning(f"Failed to trim whitespace from {converted_pdf.name}: {trim_err}")
                        else:
                            logger.debug(f"Skipping trim for {file_type} file: {file_path.name}")
                        
                except Exception as e:
                    # Thread-safe counter update
                    fail_count += 1
                    failed_files.append((file_path, target_file, str(e)))
                    logger.error(f"Failed to convert {file_path}: {e}")
                    progress.advance(task_id, advance=1)
                
                # Sync progress bar
                current_completed = progress.tasks[task_id].completed
                target_completed = (files.index(file_path) + 1)
                remaining = target_completed - current_completed
                if remaining > 0:
                    progress.advance(task_id, remaining)
                
                tui_ctx.update_progress(progress)

        finally:
            pythoncom.CoUninitialize()

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

    except KeyboardInterrupt:
        logger.warning("Conversion cancelled by user.")
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
    
    # Generate reports if enabled
    reporting_config = get_reporting_config()
    if reporting_config.enabled:
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        reports_dir = Path(reporting_config.reports_dir)
        reports_dir.mkdir(parents=True, exist_ok=True)
        
        # Summary report
        if reporting_config.summary.enabled:
            summary_filename = reporting_config.summary.format.replace("{timestamp}", timestamp)
            summary_path = reports_dir / summary_filename
            with open(summary_path, "w", encoding="utf-8") as f:
                f.write(f"doc2pdf Conversion Summary\n")
                f.write(f"{'='*50}\n")
                f.write(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Input: {input_path}\n")
                f.write(f"Output: {output_path}\n\n")
                f.write(f"Results:\n")
                f.write(f"  Success: {success_count}\n")
                f.write(f"  Failed:  {fail_count}\n")
                f.write(f"  Skipped: {skipped_count}\n")
                f.write(f"  Total:   {len(files)}\n")
            console.print(f"Summary report: [bold]{summary_path}[/bold]")
        
        # Error log with file paths
        if reporting_config.error_log.enabled and failed_files:
            error_filename = reporting_config.error_log.format.replace("{timestamp}", timestamp)
            error_path = reports_dir / error_filename
            with open(error_path, "w", encoding="utf-8") as f:
                f.write(f"doc2pdf Error Report\n")
                f.write(f"{'='*50}\n")
                f.write(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                for i, (input_file, output_file, error_msg) in enumerate(failed_files, 1):
                    f.write(f"[{i}] {input_file.name}\n")
                    f.write(f"    Input:  {input_file}\n")
                    f.write(f"    Output: {output_file}\n")
                    f.write(f"    Error:  {error_msg}\n\n")
            console.print(f"Error report: [bold]{error_path}[/bold]")
        
        # Copy error files to separate folder
        if reporting_config.copy_error_files.enabled and failed_files:
            errors_dir = output_path / reporting_config.copy_error_files.target_dir
            errors_dir.mkdir(parents=True, exist_ok=True)
            for input_file, _, _ in failed_files:
                try:
                    dest = errors_dir / input_file.name
                    shutil.copy2(input_file, dest)
                except Exception as copy_err:
                    logger.warning(f"Could not copy error file {input_file.name}: {copy_err}")
            console.print(f"Error files copied to: [bold]{errors_dir}[/bold]")
    
if __name__ == "__main__":
    app()
