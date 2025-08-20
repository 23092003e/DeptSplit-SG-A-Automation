"""
Logging configuration using Rich for beautiful console output.
"""

import logging
import sys
from typing import Optional

from rich.console import Console
from rich.logging import RichHandler
from rich.table import Table
from rich.text import Text


def setup_logging(verbose: bool = False) -> None:
    """
    Configure logging with Rich handler.
    
    Args:
        verbose: Whether to enable debug-level logging
    """
    # Create console for Rich output
    console = Console(stderr=True)
    
    # Configure logging level
    level = logging.DEBUG if verbose else logging.INFO
    
    # Create Rich handler
    rich_handler = RichHandler(
        console=console,
        show_time=True,
        show_path=False,
        markup=True,
        rich_tracebacks=True
    )
    
    # Configure logging
    logging.basicConfig(
        level=level,
        format="%(message)s",
        handlers=[rich_handler]
    )
    
    # Reduce noise from other libraries
    logging.getLogger("openpyxl").setLevel(logging.WARNING)
    logging.getLogger("xlsxwriter").setLevel(logging.WARNING)


def print_summary_table(summary: dict, console: Optional[Console] = None) -> None:
    """
    Print a formatted summary table of the split operation.
    
    Args:
        summary: Summary dictionary from split_workbook
        console: Rich console instance (creates new one if None)
    """
    if console is None:
        console = Console()
    
    # Create summary table
    table = Table(title="SG&A Split Summary", show_header=True, header_style="bold magenta")
    table.add_column("Property", style="cyan", no_wrap=True)
    table.add_column("Value", style="white")
    
    # Add summary rows
    table.add_row("Input File", summary.get('input_file', 'Unknown'))
    table.add_row("Sheet Used", summary.get('sheet_used', 'Unknown'))
    table.add_row("Header Row", str(summary.get('header_row', 'Unknown')))
    table.add_row("Department/Project Column", summary.get('dp_column', 'Unknown'))
    table.add_row("Total Input Rows", str(summary.get('total_rows', 0)))
    table.add_row("Groups Found", str(summary.get('groups_found', 0)))
    table.add_row("Files Created", str(summary.get('files_created', 0)))
    table.add_row("Export Mode", summary.get('mode', 'Unknown'))
    table.add_row("Output Directory", summary.get('output_dir', 'Unknown'))
    
    console.print()
    console.print(table)


def print_manifest_table(manifest_entries: list, console: Optional[Console] = None) -> None:
    """
    Print a formatted table of created files.
    
    Args:
        manifest_entries: List of manifest entry dictionaries
        console: Rich console instance (creates new one if None)
    """
    if console is None:
        console = Console()
    
    if not manifest_entries:
        console.print("[yellow]No files were created.[/yellow]")
        return
    
    # Create files table
    table = Table(title="Created Files", show_header=True, header_style="bold green")
    table.add_column("Department/Project", style="cyan", no_wrap=False)
    table.add_column("Rows", justify="right", style="magenta")
    table.add_column("Mode", justify="center", style="blue")
    table.add_column("File Path", style="white", no_wrap=False)
    
    # Add file rows
    for entry in manifest_entries:
        dept_project = entry.get('Department/Project', 'Unknown')
        row_count = str(entry.get('row_count', 0))
        mode = entry.get('mode', 'unknown')
        file_path = entry.get('output_path', 'Unknown')
        
        # Truncate long department names for better display
        if len(dept_project) > 30:
            dept_project = dept_project[:27] + "..."
        
        # Truncate long file paths
        if len(file_path) > 50:
            file_path = "..." + file_path[-47:]
        
        table.add_row(dept_project, row_count, mode, file_path)
    
    console.print()
    console.print(table)


def print_success_message(files_created: int, output_dir: str, console: Optional[Console] = None) -> None:
    """
    Print a success message with file count and output directory.
    
    Args:
        files_created: Number of files created
        output_dir: Output directory path
        console: Rich console instance (creates new one if None)
    """
    if console is None:
        console = Console()
    
    if files_created > 0:
        console.print()
        console.print(f"✅ [bold green]Successfully created {files_created} files in:[/bold green]")
        console.print(f"   [cyan]{output_dir}[/cyan]")
    else:
        console.print()
        console.print("⚠️ [bold yellow]No files were created. Check your input data and filters.[/bold yellow]")


def print_error_message(error: str, console: Optional[Console] = None) -> None:
    """
    Print an error message with Rich formatting.
    
    Args:
        error: Error message to display
        console: Rich console instance (creates new one if None)
    """
    if console is None:
        console = Console()
    
    console.print()
    console.print(f"❌ [bold red]Error:[/bold red] {error}")


def print_warning_message(warning: str, console: Optional[Console] = None) -> None:
    """
    Print a warning message with Rich formatting.
    
    Args:
        warning: Warning message to display
        console: Rich console instance (creates new one if None)
    """
    if console is None:
        console = Console()
    
    console.print(f"⚠️ [bold yellow]Warning:[/bold yellow] {warning}")


def print_progress_step(step: str, console: Optional[Console] = None) -> None:
    """
    Print a progress step message.
    
    Args:
        step: Description of the current step
        console: Rich console instance (creates new one if None)
    """
    if console is None:
        console = Console()
    
    console.print(f"🔄 [bold blue]{step}[/bold blue]")