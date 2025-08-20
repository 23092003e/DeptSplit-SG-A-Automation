"""
I/O utilities for file operations and data handling.
"""

import re
from pathlib import Path
from typing import Optional

import pandas as pd
import openpyxl
from openpyxl.workbook import Workbook


def load_workbook_safe(path: Path) -> Workbook:
    """
    Safely load an Excel workbook.
    
    Args:
        path: Path to the Excel file
        
    Returns:
        Loaded openpyxl workbook
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file is not a valid Excel file
    """
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    
    if not path.suffix.lower() in ['.xlsx', '.xlsm']:
        raise ValueError(f"File must be an Excel file (.xlsx or .xlsm): {path}")
    
    try:
        return openpyxl.load_workbook(path, data_only=False)
    except Exception as e:
        raise ValueError(f"Failed to load Excel file {path}: {e}")


def read_sheet_as_dataframe(
    path: Path, 
    sheet_name: str, 
    header_row: int
) -> pd.DataFrame:
    """
    Read a specific sheet as a pandas DataFrame.
    
    Args:
        path: Path to the Excel file
        sheet_name: Name of the sheet to read
        header_row: 0-based index of the header row
        
    Returns:
        DataFrame with the sheet data
    """
    try:
        # Read the sheet with the specified header row
        df = pd.read_excel(
            path,
            sheet_name=sheet_name,
            header=header_row,
            engine='openpyxl'
        )
        
        # Drop completely empty rows
        df = df.dropna(how='all')
        
        # Clean up column names - handle unnamed columns
        new_columns = []
        for i, col in enumerate(df.columns):
            if pd.isna(col) or str(col).startswith('Unnamed:'):
                new_columns.append(f"Unnamed_{i}")
            else:
                new_columns.append(str(col).strip())
        df.columns = new_columns
        
        return df
        
    except Exception as e:
        raise ValueError(f"Failed to read sheet '{sheet_name}' from {path}: {e}")


def sanitize_filename(text: str) -> str:
    """
    Sanitize text for use as a filename.
    
    Args:
        text: Text to sanitize
        
    Returns:
        Sanitized filename-safe text
    """
    if not text or not text.strip():
        return "Unknown"
    
    # Remove or replace invalid characters for Windows/macOS/Linux
    # Invalid chars: < > : " | ? * \ /
    sanitized = re.sub(r'[<>:"|?*\\/]', '_', text.strip())
    
    # Replace multiple spaces with single space
    sanitized = re.sub(r'\s+', ' ', sanitized)
    
    # Remove leading/trailing dots and spaces (problematic on Windows)
    sanitized = sanitized.strip('. ')
    
    # Limit length to avoid filesystem issues
    if len(sanitized) > 200:
        sanitized = sanitized[:200].strip()
    
    # Ensure we don't end up with an empty string
    if not sanitized:
        return "Unknown"
    
    return sanitized


def ensure_out_dir(path: Path) -> Path:
    """
    Ensure output directory exists.
    
    Args:
        path: Directory path to create
        
    Returns:
        The created directory path
    """
    path.mkdir(parents=True, exist_ok=True)
    return path


def generate_unique_filename(base_path: Path, extension: str = ".xlsx") -> Path:
    """
    Generate a unique filename by adding numeric suffix if needed.
    
    Args:
        base_path: Base path without extension
        extension: File extension to use
        
    Returns:
        Unique file path
    """
    full_path = base_path.with_suffix(extension)
    
    if not full_path.exists():
        return full_path
    
    counter = 2
    while True:
        new_name = f"{base_path.stem} #{counter}"
        new_path = base_path.with_name(new_name).with_suffix(extension)
        if not new_path.exists():
            return new_path
        counter += 1


def write_manifest_csv(manifest_data: list[dict], output_path: Path) -> None:
    """
    Write manifest data to CSV file.
    
    Args:
        manifest_data: List of dictionaries with manifest information
        output_path: Path where to write the CSV file
    """
    if not manifest_data:
        return
    
    df = pd.DataFrame(manifest_data)
    df.to_csv(output_path, index=False, encoding='utf-8')


def write_html_index(manifest_data: list[dict], output_path: Path, title: str = "SG&A Split Results") -> None:
    """
    Write HTML index file with download links.
    
    Args:
        manifest_data: List of dictionaries with manifest information
        output_path: Path where to write the HTML file
        title: Title for the HTML page
    """
    if not manifest_data:
        return
    
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 40px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 800px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        h1 {{
            color: #333;
            border-bottom: 2px solid #007acc;
            padding-bottom: 10px;
        }}
        .file-list {{
            list-style: none;
            padding: 0;
        }}
        .file-item {{
            margin: 10px 0;
            padding: 15px;
            background: #f8f9fa;
            border-radius: 5px;
            border-left: 4px solid #007acc;
        }}
        .file-link {{
            text-decoration: none;
            color: #007acc;
            font-weight: bold;
            font-size: 16px;
        }}
        .file-link:hover {{
            text-decoration: underline;
        }}
        .file-info {{
            color: #666;
            font-size: 14px;
            margin-top: 5px;
        }}
        .summary {{
            background: #e7f3ff;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>{title}</h1>
        
        <div class="summary">
            <strong>Summary:</strong> Generated {len(manifest_data)} files
        </div>
        
        <ul class="file-list">
"""
    
    for item in manifest_data:
        dept_project = item.get('Department/Project', 'Unknown')
        output_path_str = item.get('output_path', '')
        row_count = item.get('row_count', 0)
        mode = item.get('mode', 'unknown')
        
        # Convert absolute path to relative for HTML links
        file_path = Path(output_path_str)
        relative_path = file_path.name  # Just use filename for local links
        
        html_content += f"""
            <li class="file-item">
                <a href="{relative_path}" class="file-link">{dept_project}</a>
                <div class="file-info">
                    {row_count} rows • {mode} mode • {file_path.name}
                </div>
            </li>
"""
    
    html_content += """
        </ul>
    </div>
</body>
</html>"""
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)