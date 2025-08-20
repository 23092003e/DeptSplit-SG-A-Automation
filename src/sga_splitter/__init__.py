"""
SG&A Splitter - CLI tool to split Excel workbook SG&A Summary Sheet by Department/Project.

This package provides functionality to automatically detect and split Excel workbooks
containing SG&A (Selling, General & Administrative) summary data by department or project.
"""

__version__ = "0.1.0"
__author__ = "SGA Splitter"
__email__ = "noreply@example.com"

from .core import split_workbook
from .detect import find_target_sheet_name, detect_header_and_column
from .exporters import export_fast, export_clone
from .io_utils import load_workbook_safe, sanitize_filename

__all__ = [
    "split_workbook",
    "find_target_sheet_name", 
    "detect_header_and_column",
    "export_fast",
    "export_clone", 
    "load_workbook_safe",
    "sanitize_filename"
]