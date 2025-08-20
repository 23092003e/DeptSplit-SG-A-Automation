"""
Command-line interface for SG&A splitter using Typer.
"""

from pathlib import Path
from typing import Optional, Annotated
import sys

import typer
from rich.console import Console

from .core import split_workbook, split_workbook_multi_sheet, validate_inputs
from .logging_utils import (
    setup_logging,
    print_summary_table,
    print_manifest_table,
    print_success_message,
    print_error_message,
    print_progress_step
)

app = typer.Typer(
    name="sga-split",
    help="Split Excel workbook SG&A Summary Sheet by Department/Project",
    add_completion=False
)

console = Console()


@app.command()
def main(
    input_file: Annotated[
        Path,
        typer.Option("--input", "-i", help="Path to input Excel file", exists=True, file_okay=True, dir_okay=False)
    ],
    sheet: Annotated[
        Optional[str],
        typer.Option("--sheet", "-s", help="Sheet name to process (default: auto-detect)")
    ] = None,
    by: Annotated[
        Optional[str],
        typer.Option("--by", help="Department/Project column name (default: auto-detect)")
    ] = None,
    out: Annotated[
        Path,
        typer.Option("--out", "-o", help="Output directory")
    ] = Path("./SGA_Splits"),
    mode: Annotated[
        str,
        typer.Option("--mode", "-m", help="Export mode: 'fast' (pandas+xlsxwriter) or 'clone' (openpyxl+styles)")
    ] = "fast",
    skip_totals: Annotated[
        bool,
        typer.Option("--skip-totals/--keep-totals", help="Skip rows containing 'total'")
    ] = True,
    case_insensitive: Annotated[
        bool,
        typer.Option("--case-insensitive", help="Case-insensitive group matching")
    ] = False,
    fuzzy_sheet: Annotated[
        bool,
        typer.Option("--fuzzy-sheet", help="Enable fuzzy sheet name matching")
    ] = False,
    make_index: Annotated[
        bool,
        typer.Option("--make-index", help="Generate HTML index file")
    ] = False,
    manifest: Annotated[
        Optional[Path],
        typer.Option("--manifest", help="Path for manifest CSV file")
    ] = None,
    include_empty: Annotated[
        bool,
        typer.Option("--include-empty", help="Include groups that would be empty")
    ] = False,
    verbose: Annotated[
        bool,
        typer.Option("--verbose", "-v", help="Enable verbose logging")
    ] = False,
) -> None:
    """
    Split an Excel workbook's SG&A Summary Sheet into separate files by Department/Project.
    
    The tool will automatically detect the sheet and Department/Project column if not specified.
    Two export modes are available:
    
    - fast: Uses pandas + xlsxwriter for speed (no style preservation)
    - clone: Uses openpyxl to preserve styles and formulas (slower)
    
    Examples:
    
        # Basic usage with auto-detection
        sga-split --input "budget.xlsx"
        
        # Specify sheet and output directory
        sga-split --input "budget.xlsx" --sheet "SG&A Summary" --out ./splits
        
        # Use clone mode to preserve formatting
        sga-split --input "budget.xlsx" --mode clone --make-index
        
        # Enable fuzzy matching and generate manifest
        sga-split --input "budget.xlsx" --fuzzy-sheet --manifest results.csv
    """
    # Setup logging
    setup_logging(verbose)
    
    try:
        # Validate inputs
        validate_inputs(input_file, out, mode)
        
        # Show progress
        print_progress_step("Validating input file and parameters...")
        
        # Determine manifest path
        manifest_path = manifest
        if manifest_path is None and make_index:
            manifest_path = out / "manifest.csv"
        
        # Run the split operation
        print_progress_step("Analyzing workbook structure...")
        
        result = split_workbook(
            input_path=input_file,
            sheet_name=sheet,
            dp_header=by,
            mode=mode,
            out_dir=out,
            skip_totals=skip_totals,
            case_insensitive=case_insensitive,
            fuzzy_sheet=fuzzy_sheet,
            make_index=make_index,
            manifest_path=manifest_path,
            include_empty=include_empty
        )
        
        # Print results
        print_summary_table(result, console)
        
        if result['manifest_entries']:
            print_manifest_table(result['manifest_entries'], console)
        
        print_success_message(result['files_created'], result['output_dir'], console)
        
        # Additional output info
        if make_index:
            index_path = out / "index.html"
            console.print(f"📄 [bold blue]HTML index:[/bold blue] {index_path}")
        
        if manifest_path:
            console.print(f"📋 [bold blue]Manifest:[/bold blue] {manifest_path}")
        
    except Exception as e:
        print_error_message(str(e), console)
        raise typer.Exit(1)


@app.command()
def version() -> None:
    """Show version information."""
    console.print("sga-split version 0.1.0")


@app.command()
def info() -> None:
    """Show information about the tool."""
    console.print("""
[bold blue]SG&A Splitter[/bold blue]

A tool to split Excel workbook SG&A Summary Sheets by Department/Project values.

[bold]Features:[/bold]
• Automatic sheet and column detection
• Two export modes: fast (pandas) and clone (style-preserving)
• Fuzzy sheet name matching
• HTML index and CSV manifest generation
• Comprehensive logging and error handling

[bold]Supported file formats:[/bold]
• .xlsx (Excel 2007+)
• .xlsm (Excel with macros)

[bold]Export modes:[/bold]
• fast: Quick export using pandas + xlsxwriter (no formatting)
• clone: Full fidelity export using openpyxl (preserves styles/formulas)

Use --help for detailed usage information.
""")


@app.command()
def multi_sheet(
    input_file: Annotated[
        Path,
        typer.Option("--input", "-i", help="Path to input Excel file", exists=True, file_okay=True, dir_okay=False)
    ],
    out: Annotated[
        Path,
        typer.Option("--out", "-o", help="Output directory")
    ] = Path("./SGA_Splits"),
    skip_totals: Annotated[
        bool,
        typer.Option("--skip-totals/--keep-totals", help="Skip rows containing 'total'")
    ] = True,
    case_insensitive: Annotated[
        bool,
        typer.Option("--case-insensitive", help="Case-insensitive group matching")
    ] = False,
    include_empty: Annotated[
        bool,
        typer.Option("--include-empty", help="Include groups that would be empty")
    ] = False,
    remove_columns: Annotated[
        Optional[str],
        typer.Option("--remove-columns", help="Comma-separated list of column patterns to remove")
    ] = None,
    verbose: Annotated[
        bool,
        typer.Option("--verbose", "-v", help="Enable verbose logging")
    ] = False,
) -> None:
    """
    Split all sheets in a workbook with automatic sheet-specific processing:
    - Sheet 1: Split by Project
    - Sheet 2 & 3: Split by Department
    - Remove unwanted columns (Unnamed, Project/Department by default)
    - Preserve original formatting and layout
    
    This mode automatically processes multi-sheet workbooks and applies the correct
    splitting logic to each sheet while maintaining visual consistency.
    
    Examples:
    
        # Process all sheets with default column removal
        sga-split multi-sheet --input "workbook.xlsx"
        
        # Custom output directory and column removal
        sga-split multi-sheet --input "workbook.xlsx" --out ./MultiSheet_Results --remove-columns "unnamed,temp,notes"
        
        # Include total rows and empty groups
        sga-split multi-sheet --input "workbook.xlsx" --keep-totals --include-empty
    """
    # Setup logging
    setup_logging(verbose)
    
    try:
        # Parse remove_columns parameter
        remove_col_list = None
        if remove_columns:
            remove_col_list = [col.strip() for col in remove_columns.split(',')]
        
        # Show progress
        print_progress_step("Starting multi-sheet processing...")
        
        # Run the multi-sheet split operation
        result = split_workbook_multi_sheet(
            input_path=input_file,
            out_dir=out,
            skip_totals=skip_totals,
            case_insensitive=case_insensitive,
            include_empty=include_empty,
            remove_columns=remove_col_list
        )
        
        # Print results
        _print_multi_sheet_summary(result, console)
        
        if result['manifest_entries']:
            print_manifest_table(result['manifest_entries'], console)
        
        print_success_message(result['total_files_created'], result['output_dir'], console)
        
        # Additional output info
        manifest_path = Path(result['output_dir']) / "manifest.csv"
        index_path = Path(result['output_dir']) / "index.html"
        
        if manifest_path.exists():
            console.print(f"📋 [bold blue]Manifest:[/bold blue] {manifest_path}")
        
        if index_path.exists():
            console.print(f"📄 [bold blue]HTML index:[/bold blue] {index_path}")
        
    except Exception as e:
        print_error_message(str(e), console)
        raise typer.Exit(1)


def _print_multi_sheet_summary(summary: dict, console: Console) -> None:
    """
    Print a formatted summary table for multi-sheet processing.
    
    Args:
        summary: Summary dictionary from split_workbook_multi_sheet
        console: Rich console instance
    """
    from rich.table import Table
    
    # Create summary table
    table = Table(title="Multi-Sheet SG&A Split Summary", show_header=True, header_style="bold magenta")
    table.add_column("Property", style="cyan", no_wrap=True)
    table.add_column("Value", style="white")
    
    # Add summary rows
    table.add_row("Input File", summary.get('input_file', 'Unknown'))
    table.add_row("Sheets Processed", str(len(summary.get('sheets_processed', []))))
    table.add_row("Total Files Created", str(summary.get('total_files_created', 0)))
    table.add_row("Output Directory", summary.get('output_dir', 'Unknown'))
    
    console.print()
    console.print(table)
    
    # Create detailed sheet table
    if summary.get('sheets_processed'):
        sheet_table = Table(title="Sheet Processing Details", show_header=True, header_style="bold green")
        sheet_table.add_column("Sheet", style="cyan")
        sheet_table.add_column("Split By", style="blue")
        sheet_table.add_column("Column", style="yellow")
        sheet_table.add_column("Groups", justify="right", style="magenta")
        sheet_table.add_column("Files", justify="right", style="green")
        
        for sheet_info in summary['sheets_processed']:
            sheet_table.add_row(
                sheet_info.get('sheet_name', 'Unknown'),
                sheet_info.get('split_by', 'Unknown'),
                sheet_info.get('split_column', 'Unknown'),
                str(sheet_info.get('groups_found', 0)),
                str(sheet_info.get('files_created', 0))
            )
        
        console.print()
        console.print(sheet_table)


if __name__ == "__main__":
    app()