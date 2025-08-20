"""
Export utilities for generating split Excel files in fast and clone modes.
"""

from pathlib import Path
from typing import List, Dict, Any
import logging

import pandas as pd
import xlsxwriter
import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .io_utils import sanitize_filename, generate_unique_filename

logger = logging.getLogger(__name__)


def export_fast(
    df: pd.DataFrame,
    groups: List[str],
    dp_col: str,
    sheet_name: str,
    out_dir: Path
) -> List[Dict[str, Any]]:
    """
    Export groups using fast mode (pandas + xlsxwriter).
    
    Args:
        df: Source DataFrame
        groups: List of Department/Project values to export
        dp_col: Name of the Department/Project column
        sheet_name: Name for the output sheet
        out_dir: Output directory
        
    Returns:
        List of manifest entries for created files
    """
    manifest_entries = []
    
    for group in groups:
        # Filter data for this group
        group_df = df[df[dp_col] == group].copy()
        
        if group_df.empty:
            logger.warning(f"No data found for group: {group}")
            continue
        
        # Generate output filename
        sanitized_name = sanitize_filename(group)
        base_filename = f"{sanitized_name} - SG&A Summary"
        output_path = generate_unique_filename(out_dir / base_filename, ".xlsx")
        
        try:
            # Write using xlsxwriter for performance
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                group_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Get the workbook and worksheet for formatting
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]
                
                # Freeze the header row
                worksheet.freeze_panes(1, 0)
                
                # Add autofilter
                if not group_df.empty:
                    last_col = xlsxwriter.utility.xl_col_to_name(len(group_df.columns) - 1)
                    worksheet.autofilter(f'A1:{last_col}{len(group_df) + 1}')
                
                # Format header row
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#D9E1F2',
                    'border': 1
                })
                
                for col_idx, col_name in enumerate(group_df.columns):
                    worksheet.write(0, col_idx, col_name, header_format)
                    # Auto-adjust column width
                    worksheet.set_column(col_idx, col_idx, len(str(col_name)) + 2)
            
            manifest_entries.append({
                'Department/Project': group,
                'output_path': str(output_path),
                'row_count': len(group_df),
                'mode': 'fast'
            })
            
            logger.info(f"Created {output_path} with {len(group_df)} rows")
            
        except Exception as e:
            logger.error(f"Failed to create file for group '{group}': {e}")
            continue
    
    return manifest_entries


def export_clone(
    input_path: Path,
    groups: List[str],
    sheet_name: str,
    header_row: int,
    dp_col_idx: int,
    out_dir: Path
) -> List[Dict[str, Any]]:
    """
    Export groups using clone mode (openpyxl with style preservation).
    
    Args:
        input_path: Path to source Excel file
        groups: List of Department/Project values to export
        sheet_name: Name of the sheet to clone
        header_row: 0-based index of header row
        dp_col_idx: 0-based index of Department/Project column
        out_dir: Output directory
        
    Returns:
        List of manifest entries for created files
    """
    manifest_entries = []
    
    for group in groups:
        try:
            # Load fresh copy of workbook for each group
            wb = openpyxl.load_workbook(input_path)
            
            if sheet_name not in wb.sheetnames:
                logger.error(f"Sheet '{sheet_name}' not found in workbook")
                continue
            
            ws = wb[sheet_name]
            
            # Find rows to keep (header + matching data rows)
            rows_to_keep = set()
            rows_to_keep.update(range(1, header_row + 2))  # Keep everything up to and including header
            
            row_count = 0
            for row_idx in range(header_row + 2, ws.max_row + 1):
                cell_value = ws.cell(row=row_idx, column=dp_col_idx + 1).value
                if cell_value is not None and str(cell_value).strip() == group:
                    rows_to_keep.add(row_idx)
                    row_count += 1
            
            if row_count == 0:
                logger.warning(f"No data rows found for group: {group}")
                wb.close()
                continue
            
            # Delete rows that don't match (from bottom to top to preserve indices)
            rows_to_delete = []
            for row_idx in range(header_row + 2, ws.max_row + 1):
                if row_idx not in rows_to_keep:
                    rows_to_delete.append(row_idx)
            
            # Delete from bottom to top
            for row_idx in sorted(rows_to_delete, reverse=True):
                ws.delete_rows(row_idx)
            
            # Update autofilter if it exists
            if ws.auto_filter:
                ws.auto_filter.ref = None
            
            # Reapply autofilter to the remaining data
            if ws.max_row > header_row + 1:  # If we have data rows
                end_col = openpyxl.utils.get_column_letter(ws.max_column)
                ws.auto_filter.ref = f"A{header_row + 1}:{end_col}{ws.max_row}"
            
            # Generate output filename
            sanitized_name = sanitize_filename(group)
            base_filename = f"{sanitized_name} - SG&A Summary"
            output_path = generate_unique_filename(out_dir / base_filename, ".xlsx")
            
            # Save the modified workbook
            wb.save(output_path)
            wb.close()
            
            manifest_entries.append({
                'Department/Project': group,
                'output_path': str(output_path),
                'row_count': row_count,
                'mode': 'clone'
            })
            
            logger.info(f"Created {output_path} with {row_count} rows (clone mode)")
            
        except Exception as e:
            logger.error(f"Failed to create file for group '{group}' in clone mode: {e}")
            continue
    
    return manifest_entries


def _should_skip_row_for_group(ws: Worksheet, row_idx: int, dp_col_idx: int, target_group: str) -> bool:
    """
    Check if a row should be skipped (doesn't match the target group).
    
    Args:
        ws: Worksheet to check
        row_idx: 1-based row index
        dp_col_idx: 0-based column index for Department/Project
        target_group: Target group value to match
        
    Returns:
        True if row should be skipped
    """
    cell_value = ws.cell(row=row_idx, column=dp_col_idx + 1).value
    
    if cell_value is None:
        return True
    
    return str(cell_value).strip() != target_group


def export_clone_multi_sheet(
    input_path: Path,
    groups: List[str],
    sheet_name: str,
    header_row: int,
    split_col_idx: int,
    out_dir: Path,
    remove_columns: List[str],
    original_split_col_idx: int
) -> List[Dict[str, Any]]:
    """
    Export groups using clone mode with column removal and enhanced formatting preservation.
    
    Args:
        input_path: Path to source Excel file
        groups: List of values to export
        sheet_name: Name of the sheet to clone
        header_row: 0-based index of header row
        split_col_idx: 0-based index of split column (after column removal)
        out_dir: Output directory
        remove_columns: List of column patterns to remove
        original_split_col_idx: Original split column index before column removal
        
    Returns:
        List of manifest entries for created files
    """
    manifest_entries = []
    
    for group in groups:
        try:
            # Load fresh copy of workbook for each group
            wb = openpyxl.load_workbook(input_path)
            
            if sheet_name not in wb.sheetnames:
                logger.error(f"Sheet '{sheet_name}' not found in workbook")
                continue
            
            ws = wb[sheet_name]
            
            # Remove unwanted columns first (but preserve the split column)
            columns_to_remove = _identify_columns_to_remove(ws, header_row, remove_columns, preserve_col_idx=original_split_col_idx)
            
            # Adjust split column index if columns before it are removed
            adjusted_split_col_idx = original_split_col_idx
            for col_idx in sorted(columns_to_remove):
                if col_idx < original_split_col_idx:
                    adjusted_split_col_idx -= 1
            
            # Remove columns from right to left to preserve indices
            for col_idx in sorted(columns_to_remove, reverse=True):
                ws.delete_cols(col_idx + 1)  # openpyxl uses 1-based indexing
            
            # Find rows to keep (header + matching data rows)
            rows_to_keep = set()
            rows_to_keep.update(range(1, header_row + 2))  # Keep everything up to and including header
            
            row_count = 0
            for row_idx in range(header_row + 2, ws.max_row + 1):
                cell_value = ws.cell(row=row_idx, column=adjusted_split_col_idx + 1).value
                if cell_value is not None and str(cell_value).strip() == group:
                    rows_to_keep.add(row_idx)
                    row_count += 1
            
            if row_count == 0:
                logger.warning(f"No data rows found for group: {group}")
                wb.close()
                continue
            
            # Delete rows that don't match (from bottom to top to preserve indices)
            rows_to_delete = []
            for row_idx in range(header_row + 2, ws.max_row + 1):
                if row_idx not in rows_to_keep:
                    rows_to_delete.append(row_idx)
            
            # Delete from bottom to top
            for row_idx in sorted(rows_to_delete, reverse=True):
                ws.delete_rows(row_idx)
            
            # Preserve and enhance formatting
            _preserve_formatting(ws, header_row)
            
            # Update autofilter if it exists
            if ws.auto_filter:
                ws.auto_filter.ref = None
            
            # Reapply autofilter to the remaining data
            if ws.max_row > header_row + 1:  # If we have data rows
                end_col = openpyxl.utils.get_column_letter(ws.max_column)
                ws.auto_filter.ref = f"A{header_row + 1}:{end_col}{ws.max_row}"
            
            # Generate output filename
            sanitized_name = sanitize_filename(group)
            base_filename = f"{sanitized_name} - {sheet_name}"
            output_path = generate_unique_filename(out_dir / base_filename, ".xlsx")
            
            # Save the modified workbook
            wb.save(output_path)
            wb.close()
            
            manifest_entries.append({
                'Sheet': sheet_name,
                'Group': group,
                'output_path': str(output_path),
                'row_count': row_count,
                'mode': 'clone_multi'
            })
            
            logger.info(f"Created {output_path} with {row_count} rows (multi-sheet clone mode)")
            
        except Exception as e:
            logger.error(f"Failed to create file for group '{group}' in multi-sheet clone mode: {e}")
            continue
    
    return manifest_entries


def _identify_columns_to_remove(ws: Worksheet, header_row: int, remove_patterns: List[str], preserve_col_idx: int = None) -> List[int]:
    """
    Identify column indices to remove based on patterns, but preserve specified column.
    
    Args:
        ws: Worksheet to analyze
        header_row: 0-based header row index
        remove_patterns: List of patterns to match for removal
        preserve_col_idx: 0-based column index to preserve even if it matches patterns
        
    Returns:
        List of 0-based column indices to remove
    """
    columns_to_remove = []
    
    if not remove_patterns:
        return columns_to_remove
    
    # Get header row values
    max_col = ws.max_column
    for col_idx in range(max_col):
        # Always preserve the split column
        if preserve_col_idx is not None and col_idx == preserve_col_idx:
            logger.info(f"Preserving split column at index {col_idx + 1}")
            continue
            
        header_cell = ws.cell(row=header_row + 1, column=col_idx + 1)
        header_value = str(header_cell.value or "").strip().lower()
        
        should_remove = False
        for pattern in remove_patterns:
            pattern_lower = pattern.lower()
            if (pattern_lower in header_value or 
                header_value.startswith('unnamed') or
                'project/department' in header_value or
                'department/project' in header_value):
                should_remove = True
                break
        
        if should_remove:
            columns_to_remove.append(col_idx)
            logger.info(f"Marking column {col_idx + 1} ('{header_value}') for removal")
    
    return columns_to_remove


def _preserve_formatting(ws: Worksheet, header_row: int):
    """
    Preserve and enhance worksheet formatting.
    
    Args:
        ws: Worksheet to format
        header_row: 0-based header row index
    """
    try:
        # Preserve column widths by copying from original
        # (This is already preserved by openpyxl when not explicitly changed)
        
        # Ensure header row has proper formatting
        for cell in ws[header_row + 1]:  # openpyxl uses 1-based row indexing
            if cell.value:
                # Keep existing formatting but ensure visibility
                if not cell.font.bold:
                    from openpyxl.styles import Font
                    cell.font = Font(bold=True)
        
        # Auto-adjust column widths if they seem too narrow
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            
            # Set a reasonable width (minimum 10, maximum 50)
            adjusted_width = min(max(max_length + 2, 10), 50)
            ws.column_dimensions[column_letter].width = adjusted_width
            
    except Exception as e:
        logger.warning(f"Could not fully preserve formatting: {e}")
        # Continue processing even if formatting fails