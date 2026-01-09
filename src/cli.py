import typer
import sys
import time
from pathlib import Path
from typing import Optional, List
from rich.console import Console
from rich.panel import Panel
from rich.progress import (
    Progress, SpinnerColumn, TextColumn, BarColumn, 
    TaskProgressColumn, TimeElapsedColumn, TimeRemainingColumn
)
from rich.logging import RichHandler
from rich.layout import Layout
from rich.live import Live
from rich.table import Table

from .version import __version__
from .core.word_converter import WordConverter
from .utils.logger import setup_logger, logger
from .config import get_logging_config, get_pdf_settings, FileType

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
    pass

def get_files(path: Path) -> List[Path]:
    if path.is_file():
        return [path]
    
    extensions = {
        "*.docx", "*.doc", 
        "*.xlsx", "*.xls",
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
    elif ext in [".xlsx", ".xls"]:
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
    # excel_converter = ExcelConverter() # TODO: Implement
    # ppt_converter = PowerPointConverter() # TODO: Implement

    # TUI Setup
    progress = Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        TimeElapsedColumn(),
        TimeRemainingColumn(),
        console=console
    )

    with progress:
        task_id = progress.add_task(f"[cyan]Converting {len(files)} files...", total=len(files))
        
        success_count = 0
        fail_count = 0
        skipped_count = 0
        
        for file_path in files:
            file_type = get_file_type(file_path)
            progress.update(task_id, description=f"[cyan]Converting ({file_type}): {file_path.name}")
            
            # Get settings based on file type and pattern overrides
            settings = get_pdf_settings(input_path=file_path, file_type=file_type)
            
            # Determine output
            if output_path:
                if input_path.is_dir():
                    # Calculate relative path to maintain structure
                    rel_path = file_path.relative_to(input_path)
                    target_file = output_path / rel_path.with_suffix(".pdf")
                    target_file.parent.mkdir(parents=True, exist_ok=True)
                else:
                    if output_path.suffix.lower() == ".pdf":
                        target_file = output_path
                    else:
                        target_file = output_path / file_path.with_suffix(".pdf").name
                        target_file.parent.mkdir(parents=True, exist_ok=True)
            else:
                target_file = None 

            try:
                if file_type == "word":
                    word_converter.convert(file_path, target_file, settings)
                    success_count += 1
                else:
                    # Placeholder for other types
                    logger.warning(f"Conversion for {file_type} not yet implemented. Skipping {file_path.name}")
                    skipped_count += 1
                    
            except Exception as e:
                fail_count += 1
                logger.error(f"Failed to convert {file_path}: {e}")
            
            progress.advance(task_id)

    # Summary
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
