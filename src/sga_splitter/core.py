"""
Core orchestration logic for SG&A workbook splitting.
"""

import re
from pathlib import Path
from typing import Optional, Literal, Dict, Any, List
import logging

import pandas as pd

from .detect import find_target_sheet_name, detect_header_and_column
from .io_utils import (
    load_workbook_safe, 
    read_sheet_as_dataframe, 
    ensure_out_dir,
    write_manifest_csv,
    write_html_index
)
from .exporters import export_fast, export_clone, export_clone_multi_sheet

logger = logging.getLogger(__name__)


def split_workbook(
    input_path: Path,
    sheet_name: Optional[str],
    dp_header: Optional[str],
    mode: Literal["fast", "clone"],
    out_dir: Path,
    skip_totals: bool,
    case_insensitive: bool,
    fuzzy_sheet: bool,
    make_index: bool,
    manifest_path: Optional[Path],
    include_empty: bool
) -> Dict[str, Any]:
    """
    Main orchestration function to split an Excel workbook by Department/Project.
    
    Args:
        input_path: Path to input Excel file
        sheet_name: Specific sheet name to use (None for auto-detection)
        dp_header: Specific Department/Project column name (None for auto-detection)
        mode: Export mode ('fast' or 'clone')
        out_dir: Output directory for split files
        skip_totals: Whether to skip rows containing 'total'
        case_insensitive: Whether group matching should be case insensitive
        fuzzy_sheet: Whether to enable fuzzy sheet name matching
        make_index: Whether to generate HTML index file
        manifest_path: Path for manifest CSV file (None to skip)
        include_empty: Whether to include groups that would be empty
        
    Returns:
        Summary dictionary with results
    """
    logger.info(f"Starting workbook split: {input_path}")
    
    # Ensure output directory exists
    ensure_out_dir(out_dir)
    
    # Load workbook
    wb = load_workbook_safe(input_path)
    logger.info(f"Loaded workbook with {len(wb.sheetnames)} sheets")
    
    # Find target sheet
    target_sheet = find_target_sheet_name(wb, sheet_name, fuzzy_sheet)
    logger.info(f"Using sheet: '{target_sheet}'")
    
    # Detect header row and Department/Project column
    ws = wb[target_sheet]
    header_row_idx, dp_col_idx = detect_header_and_column(ws)
    logger.info(f"Found header at row {header_row_idx + 1}, Department/Project column at index {dp_col_idx}")
    
    # Read sheet as DataFrame
    df = read_sheet_as_dataframe(input_path, target_sheet, header_row_idx)
    logger.info(f"Loaded {len(df)} rows of data")
    
    # Get Department/Project column name
    dp_col_name = df.columns[dp_col_idx]
    logger.info(f"Department/Project column: '{dp_col_name}'")
    
    # Collect unique groups
    groups = collect_groups(
        df, 
        dp_col_name, 
        skip_totals=skip_totals, 
        case_insensitive=case_insensitive,
        include_empty=include_empty
    )
    
    logger.info(f"Found {len(groups)} unique groups: {groups[:5]}{'...' if len(groups) > 5 else ''}")
    
    # Export files based on mode
    if mode == "fast":
        manifest_entries = export_fast(df, groups, dp_col_name, target_sheet, out_dir)
    else:  # clone mode
        manifest_entries = export_clone(
            input_path, groups, target_sheet, header_row_idx, dp_col_idx, out_dir
        )
    
    # Write manifest CSV if requested
    if manifest_path and manifest_entries:
        write_manifest_csv(manifest_entries, manifest_path)
        logger.info(f"Wrote manifest to: {manifest_path}")
    
    # Write HTML index if requested
    if make_index and manifest_entries:
        index_path = out_dir / "index.html"
        write_html_index(manifest_entries, index_path)
        logger.info(f"Wrote HTML index to: {index_path}")
    
    # Close workbook
    wb.close()
    
    # Return summary
    return {
        'input_file': str(input_path),
        'sheet_used': target_sheet,
        'header_row': header_row_idx + 1,  # Convert to 1-based for display
        'dp_column': dp_col_name,
        'total_rows': len(df),
        'groups_found': len(groups),
        'files_created': len(manifest_entries),
        'mode': mode,
        'output_dir': str(out_dir),
        'manifest_entries': manifest_entries
    }


def collect_groups(
    df: pd.DataFrame,
    dp_col_name: str,
    skip_totals: bool = True,
    case_insensitive: bool = False,
    include_empty: bool = False
) -> List[str]:
    """
    Collect unique Department/Project groups from the DataFrame.
    
    Args:
        df: Source DataFrame
        dp_col_name: Name of the Department/Project column
        skip_totals: Whether to skip groups containing 'total'
        case_insensitive: Whether to treat groups case-insensitively
        include_empty: Whether to include empty/null groups
        
    Returns:
        List of unique group names
    """
    if dp_col_name not in df.columns:
        raise ValueError(f"Column '{dp_col_name}' not found in DataFrame")
    
    # Get unique values, handling NaN
    series = df[dp_col_name].dropna() if not include_empty else df[dp_col_name]
    unique_values = series.unique()
    
    # Convert to strings and strip whitespace
    groups = []
    for value in unique_values:
        if pd.isna(value):
            if include_empty:
                groups.append("")
            continue
        
        group_name = str(value).strip()
        
        # Skip empty strings unless explicitly including them
        if not group_name and not include_empty:
            continue
        
        # Skip totals if requested
        if skip_totals and _is_total_row(group_name):
            logger.info(f"Skipping total row: '{group_name}'")
            continue
        
        groups.append(group_name)
    
    # Remove duplicates while preserving order
    unique_groups = []
    seen = set()
    
    for group in groups:
        # Handle case insensitive comparison for deduplication
        compare_key = group.lower() if case_insensitive else group
        
        if compare_key not in seen:
            seen.add(compare_key)
            unique_groups.append(group)
    
    return sorted(unique_groups)


def _is_total_row(value: str) -> bool:
    """
    Check if a value represents a total/summary row.
    
    Args:
        value: Value to check
        
    Returns:
        True if value appears to be a total row
    """
    if not value:
        return False
    
    # Look for 'total' as a whole word (case insensitive)
    pattern = r'\btotal\b'
    return bool(re.search(pattern, value.lower()))


def validate_inputs(
    input_path: Path,
    out_dir: Path,
    mode: str
) -> None:
    """
    Validate input parameters.
    
    Args:
        input_path: Path to input file
        out_dir: Output directory
        mode: Export mode
        
    Raises:
        ValueError: If validation fails
    """
    if not input_path.exists():
        raise ValueError(f"Input file does not exist: {input_path}")
    
    if not input_path.suffix.lower() in ['.xlsx', '.xlsm']:
        raise ValueError(f"Input file must be Excel format (.xlsx or .xlsm): {input_path}")
    
    if mode not in ['fast', 'clone']:
        raise ValueError(f"Mode must be 'fast' or 'clone', got: {mode}")
    
    try:
        ensure_out_dir(out_dir)
    except Exception as e:
        raise ValueError(f"Cannot create output directory {out_dir}: {e}")


def split_workbook_multi_sheet(
    input_path: Path,
    out_dir: Path,
    skip_totals: bool = True,
    case_insensitive: bool = False,
    include_empty: bool = False,
    remove_columns: Optional[List[str]] = None
) -> Dict[str, Any]:
    """
    Process all three sheets in a workbook and create a single output file containing all sheets:
    - Sheet 1: Split by Project
    - Sheet 2 & 3: Split by Department
    - Remove specified columns from output
    - Preserve formatting in clone mode
    
    Args:
        input_path: Path to input Excel file
        out_dir: Output directory for the single output file
        skip_totals: Whether to skip rows containing 'total'
        case_insensitive: Whether group matching should be case insensitive
        include_empty: Whether to include groups that would be empty
        remove_columns: List of column names/patterns to remove from output
        
    Returns:
        Summary dictionary with results
    """
    logger.info(f"Starting multi-sheet workbook processing: {input_path}")
    
    # Default columns to remove
    if remove_columns is None:
        remove_columns = ["unnamed", "project/department"]
    
    # Ensure output directory exists
    ensure_out_dir(out_dir)
    
    # Load workbook
    wb = load_workbook_safe(input_path)
    sheet_names = wb.sheetnames
    logger.info(f"Loaded workbook with {len(sheet_names)} sheets: {sheet_names}")
    
    if len(sheet_names) < 3:
        raise ValueError(f"Workbook must have at least 3 sheets, found {len(sheet_names)}")
    
    # Define processing configuration for each sheet
    sheet_configs = [
        {
            'sheet_name': sheet_names[0],
            'split_by': 'project',
            'column_patterns': ['project', 'proj'],
        },
        {
            'sheet_name': sheet_names[1], 
            'split_by': 'department',
            'column_patterns': ['department', 'dept'],
        },
        {
            'sheet_name': sheet_names[2],
            'split_by': 'department', 
            'column_patterns': ['department', 'dept'],
        }
    ]
    
    # First pass: Process all sheets and collect their data
    processed_sheets = []
    all_groups = set()
    
    for i, config in enumerate(sheet_configs):
        try:
            logger.info(f"Processing sheet {i+1}: '{config['sheet_name']}' - split by {config['split_by']}")
            
            # Process this sheet and collect its information
            sheet_data = _analyze_sheet_for_multi_processing(
                wb=wb,
                input_path=input_path,
                sheet_config=config,
                skip_totals=skip_totals,
                case_insensitive=case_insensitive,
                include_empty=include_empty,
                remove_columns=remove_columns
            )
            
            processed_sheets.append(sheet_data)
            all_groups.update(sheet_data['groups'])
            
        except Exception as e:
            logger.error(f"Failed to process sheet '{config['sheet_name']}': {e}")
            continue
    
    if not processed_sheets:
        raise ValueError("No sheets could be processed")
    
    logger.info(f"Found {len(all_groups)} unique groups across all sheets: {list(all_groups)}")
    
    # Second pass: Create one file per group with all 3 sheets
    from .exporters import create_multi_sheet_files_per_group
    
    manifest_entries = create_multi_sheet_files_per_group(
        input_path=input_path,
        groups=list(all_groups),
        processed_sheets=processed_sheets,
        out_dir=out_dir,
        remove_columns=remove_columns
    )
    
    # Write manifest
    manifest_path = out_dir / "manifest.csv"
    if manifest_entries:
        write_manifest_csv(manifest_entries, manifest_path)
        logger.info(f"Wrote manifest to: {manifest_path}")
    
    # Write HTML index
    index_path = out_dir / "index.html"
    if manifest_entries:
        write_html_index(manifest_entries, index_path, "Multi-Sheet SG&A Split Results")
        logger.info(f"Wrote HTML index to: {index_path}")
    
    wb.close()
    
    summary_data = {
        'input_file': str(input_path),
        'sheets_processed': [sheet['config']['sheet_name'] for sheet in processed_sheets],
        'total_groups': len(all_groups),
        'files_created': len(manifest_entries),
        'output_dir': str(out_dir),
        'manifest_entries': manifest_entries
    }
    
    logger.info(f"Multi-sheet processing complete: {len(manifest_entries)} files created for {len(all_groups)} groups")
    return summary_data




def _analyze_sheet_for_multi_processing(
    wb,
    input_path: Path,
    sheet_config: Dict[str, Any],
    skip_totals: bool,
    case_insensitive: bool,
    include_empty: bool,
    remove_columns: List[str]
) -> Dict[str, Any]:
    """
    Analyze a single sheet for multi-sheet processing - collects data without creating files.
    
    Args:
        wb: Loaded workbook
        input_path: Path to input file
        sheet_config: Configuration for this sheet
        skip_totals: Whether to skip total rows
        case_insensitive: Case insensitive matching
        include_empty: Include empty groups
        remove_columns: Columns to remove
        
    Returns:
        Dictionary with sheet analysis data
    """
    sheet_name = sheet_config['sheet_name']
    split_by = sheet_config['split_by']
    column_patterns = sheet_config['column_patterns']
    
    # Get worksheet
    ws = wb[sheet_name]
    
    # Detect header and split column using flexible patterns
    header_row_idx, split_col_idx = _detect_header_and_split_column(ws, column_patterns)
    logger.info(f"Sheet '{sheet_name}': header at row {header_row_idx + 1}, {split_by} column at index {split_col_idx}")
    
    # Read sheet as DataFrame
    df = read_sheet_as_dataframe(input_path, sheet_name, header_row_idx)
    
    # Get split column name
    split_col_name = df.columns[split_col_idx]
    logger.info(f"Sheet '{sheet_name}': splitting by column '{split_col_name}'")
    
    # Remove unwanted columns from DataFrame (but preserve the split column)
    df_cleaned = _remove_unwanted_columns(df, remove_columns, preserve_column=split_col_name)
    
    # Adjust split column index after column removal
    if split_col_name in df_cleaned.columns:
        new_split_col_idx = df_cleaned.columns.get_loc(split_col_name)
    else:
        raise ValueError(f"Split column '{split_col_name}' was removed by column filtering")
    
    # Collect groups
    groups = collect_groups(
        df_cleaned,
        split_col_name,
        skip_totals=skip_totals,
        case_insensitive=case_insensitive,
        include_empty=include_empty
    )
    
    logger.info(f"Sheet '{sheet_name}': found {len(groups)} groups")
    
    return {
        'config': sheet_config,
        'header_row': header_row_idx,
        'split_col_idx': split_col_idx,
        'split_col_name': split_col_name,
        'new_split_col_idx': new_split_col_idx,
        'groups': groups,
        'df': df_cleaned
    }


def _detect_header_and_split_column(ws, column_patterns: List[str]):
    """
    Detect header row and split column using flexible patterns.
    
    Args:
        ws: Worksheet to analyze
        column_patterns: List of patterns to match for split column
        
    Returns:
        Tuple of (header_row_index, split_column_index) (0-based)
    """
    from .detect import candidate_name_matches
    
    max_rows_to_scan = min(50, ws.max_row)
    
    # Look for a proper header row - one that has multiple column headers
    for row_idx in range(max_rows_to_scan):
        row_values = []
        for col_idx in range(min(20, ws.max_column)):
            cell = ws.cell(row=row_idx + 1, column=col_idx + 1)
            value = str(cell.value or "").strip()
            row_values.append(value)
        
        # Check if this looks like a header row (has multiple non-empty cells)
        non_empty_count = sum(1 for val in row_values if val)
        if non_empty_count >= 2:  # At least 2 columns with headers
            # Find the best match for our patterns, preferring exact matches
            best_match = None
            best_score = 0
            
            for col_idx, header in enumerate(row_values):
                if _matches_any_pattern(header, column_patterns):
                    # Calculate match quality - prefer exact matches and shorter headers
                    match_score = 0
                    header_lower = header.lower()
                    
                    # Exact match gets highest score
                    for pattern in column_patterns:
                        if header_lower == pattern.lower():
                            match_score = 100
                            break
                        elif header_lower.startswith(pattern.lower()) and '/' not in header_lower:
                            match_score = 80  # High score for clean prefix matches
                        elif header_lower.startswith(pattern.lower()):
                            match_score = 40  # Lower score if prefix match but has separators
                        elif header_lower.endswith(pattern.lower()) and '/' not in header_lower:
                            match_score = 70  # Good score for clean suffix matches like "B_Project"
                        elif pattern.lower() in header_lower and '/' not in header_lower:
                            match_score = max(match_score, 60)  # Good score if no separators
                        elif pattern.lower() in header_lower:
                            match_score = max(match_score, 20)  # Low score if has separators
                    
                    # Heavily penalize compound headers like "Project/Department" that will be removed
                    if '/' in header_lower or '_' in header_lower or 'department' in header_lower:
                        match_score -= 50
                    
                    # Extra penalty if this column name appears in removal patterns
                    for remove_pattern in ['project/department', 'department/project', 'unnamed']:
                        if remove_pattern.lower() in header_lower:
                            match_score -= 100
                    
                    # Prefer shorter headers (more specific)
                    match_score += max(0, 20 - len(header))
                    
                    if match_score > best_score:
                        best_score = match_score
                        best_match = (row_idx, col_idx)
            
            if best_match and non_empty_count >= 3:
                return best_match
    
    # If no specific pattern found, try the original detection
    from .detect import detect_header_and_column
    try:
        return detect_header_and_column(ws)
    except ValueError:
        # If that fails, look for first row with multiple columns that looks like a header
        for row_idx in range(max_rows_to_scan):
            row_values = [str(ws.cell(row=row_idx + 1, column=col_idx + 1).value or "").strip() 
                         for col_idx in range(min(10, ws.max_column))]
            non_empty_count = sum(1 for val in row_values if val)
            
            if non_empty_count >= 3:  # At least 3 columns suggest a header row
                # Find first column that looks like it could contain groups
                for col_idx, header in enumerate(row_values):
                    if header and len(header) > 0:
                        return row_idx, col_idx
        
        raise ValueError(f"Could not detect header and split column for patterns: {column_patterns}")


def _matches_any_pattern(header: str, patterns: List[str]) -> bool:
    """
    Check if header matches any of the given patterns.
    
    Args:
        header: Header text to check
        patterns: List of patterns to match against
        
    Returns:
        True if header matches any pattern
    """
    if not header:
        return False
    
    normalized = header.lower().strip()
    
    for pattern in patterns:
        if pattern.lower() in normalized:
            return True
    
    return False


def _remove_unwanted_columns(df: pd.DataFrame, remove_patterns: List[str], preserve_column: str = None) -> pd.DataFrame:
    """
    Remove columns that match unwanted patterns, but preserve the specified column.
    
    Args:
        df: Source DataFrame
        remove_patterns: List of patterns to match for removal
        preserve_column: Column name to preserve even if it matches removal patterns
        
    Returns:
        DataFrame with unwanted columns removed
    """
    if not remove_patterns:
        return df
    
    columns_to_keep = []
    
    for col in df.columns:
        col_lower = str(col).lower().strip()
        
        # Always preserve the split column
        if preserve_column and str(col) == preserve_column:
            columns_to_keep.append(col)
            continue
        
        should_remove = False
        
        for pattern in remove_patterns:
            pattern_lower = pattern.lower()
            if (pattern_lower in col_lower or 
                col_lower.startswith('unnamed') or
                'project/department' in col_lower or
                'department/project' in col_lower):
                should_remove = True
                break
        
        if not should_remove:
            columns_to_keep.append(col)
    
    logger.info(f"Keeping {len(columns_to_keep)} columns, removed {len(df.columns) - len(columns_to_keep)} columns")
    if preserve_column:
        logger.info(f"Preserved split column: '{preserve_column}'")
    return df[columns_to_keep]