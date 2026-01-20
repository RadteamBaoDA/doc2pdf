import typer
import sys
import time
from pathlib import Path
from typing import Optional, List
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
from .utils.logger import setup_logger, logger
from .config import get_logging_config, get_pdf_settings, get_suffix_config, FileType

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
    setup_logger(current_config)

    files = get_files(input_path)
    
    if not files:
        console.print(f"[yellow]No supported Office documents found in {input_path}.[/yellow]")
        raise typer.Exit()

    # Initialize converters
    word_converter = WordConverter()
    ppt_converter = PowerPointConverter()
    excel_converter = ExcelConverter()

    # TUI Setup
    from .tui import LogBuffer, TUIContext
    
    log_buffer = LogBuffer()
    
    # Redirect Logger to TUI Buffer
    def tui_sink(message):
        record = message.record
        level_name = record['level'].name
        
        # Color mapping
        colors = {
            "INFO": "green",
            "WARNING": "yellow",
            "ERROR": "bold red",
            "CRITICAL": "bold white on red",
            "DEBUG": "dim blue"
        }
        color = colors.get(level_name, "white")
        
        log_msg = f"[{color}]{record['time'].strftime('%H:%M:%S')} | {level_name: <8} | {record['message']}[/{color}]"
        log_buffer.write(log_msg)
    
    # Add TUI sink
    try:
        logger.add(tui_sink, format="{message}", level="INFO")
    except Exception:
        pass 

    tui_ctx = TUIContext(log_buffer)

    # Initialize Progress (passive)
    progress = Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        MofNCompleteColumn(), # Added File Counter (M of N)
        TimeElapsedColumn(),
        TimeRemainingColumn(),
        expand=True
    )

    try:
        with Live(tui_ctx.layout, refresh_per_second=10, screen=True):
            task_id = progress.add_task(f"[cyan]Converting {len(files)} files...", total=len(files))
            
            success_count = 0
            fail_count = 0
            skipped_count = 0
            
            tui_ctx.update_progress(progress) # Initial render
            
            for file_path in files:
                file_type = get_file_type(file_path)
                progress.update(task_id, description=f"[cyan]Converting ({file_type}): {file_path.name}")
                tui_ctx.update_progress(progress)
                tui_ctx.update_logs()
                
                # Get settings based on file type and pattern overrides
                settings = get_pdf_settings(input_path=file_path, file_type=file_type)
                
                # Get suffix for this file type
                suffix_config = get_suffix_config()
                suffix = suffix_config.get(file_type, "")
                
                # Determine output with suffix
                if output_path:
                    if input_path.is_dir():
                        # Calculate relative path to maintain structure
                        rel_path = file_path.relative_to(input_path)
                        # Apply suffix: filename_suffix.pdf
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
                    
                    if file_type == "word":
                        word_converter.convert(file_path, target_file, settings)
                        progress.advance(task_id, advance=1)
                    elif file_type == "powerpoint":
                        ppt_converter.convert(file_path, target_file, settings)
                        progress.advance(task_id, advance=1)
                    elif file_type == "excel":
                        # Excel converter handles its own progress if callback provided
                        # But we need to make sure we don't over-advance if it doesn't chunk.
                        # Actually our ExcelConverter logic only calls callback if chunking.
                        # Let's trust it. If it doesn't call back (no chunking), we advance 1 at end.
                        # Wait, convert method advances partials. We need to track it?
                        # Simplified: converter reports 0.xxx. We accumulate or just advance.
                        # If converter logic is: "Chunk 1 done" -> advance(0.2).
                        # We need to distinguish between "Self-reporting" and "Manual".
                        # Let's modify logic: Convert method is void.
                        # If we pass callback, it uses it.
                        # We should only advance manually if callback wasn't used fully?
                        # Or safer: Pass callback. If callback was called, good.
                        # But simpler: converter advances 1.0 TOTAL.
                        excel_converter.convert(file_path, target_file, settings, on_progress=progress_callback)
                        # If the converter didn't chunk, it wouldn't have called progress_callback.
                        # We should check if we need to force complete.
                        # But 'convert' is blocking. When it returns, the file is DONE.
                        # So we can just ensure task is advanced to next integer?
                        # progress.update(task_id, completed=files_processed_count)
                        
                        success_count += 1
                    else:
                        logger.warning(f"Conversion for {file_type} not supported. Skipping {file_path.name}")
                        skipped_count += 1
                        progress.advance(task_id, advance=1)
                        
                except Exception as e:
                    fail_count += 1
                    logger.error(f"Failed to convert {file_path}: {e}")
                    # If failed, we still need to advance to keep counter correct?
                    # Yes, file is "processed" (failed).
                    progress.advance(task_id, advance=1) # May overfill if partials were reported?
                
                # Ensure we are exactly at the next integer step (M of N count relies on completed tasks)
                # If we used partials, we might be at 3.99.
                # Rich's advance adds to completed.
                # We can just set completed explicitly?
                # progress.update(task_id, completed=success_count + fail_count + skipped_count)
                # But success_count etc are local.
                
                # Safe approach:
                # We assume 1 unit per file.
                # If excel converter advanced 0.99, we need 0.01.
                # If excel converter didn't call callback, we need 1.0.
                # Let's make ExcelConverter responsible for 100% via callback if passed?
                # Or just update to integer index.
                current_completed = progress.tasks[task_id].completed
                target_completed = (files.index(file_path) + 1)
                remaining = target_completed - current_completed
                if remaining > 0:
                    progress.advance(task_id, remaining)
                
                tui_ctx.update_progress(progress)
                tui_ctx.update_logs()

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
    
if __name__ == "__main__":
    app()
