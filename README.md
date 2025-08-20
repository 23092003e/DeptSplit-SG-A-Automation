# SG&A Splitter

A production-ready CLI tool to split Excel workbook SG&A Summary Sheets into separate files by Department/Project values.

## Features

- **Multi-Sheet Processing**: Automatically processes all sheets with sheet-specific splitting logic:
  - Sheet 1: Split by Project
  - Sheet 2 & 3: Split by Department
- **Column Removal**: Automatically removes unwanted columns (Unnamed, Project/Department) from output
- **Formatting Preservation**: Maintains original cell styles, colors, fonts, borders, and column widths
- **Automatic Detection**: Intelligently finds SG&A summary sheets and Department/Project columns
- **Two Export Modes**:
  - **Fast**: Quick export using pandas + xlsxwriter (no style preservation)
  - **Clone**: Full fidelity export using openpyxl (preserves styles, formulas, and formatting)
- **Fuzzy Matching**: Flexible sheet name detection with configurable matching
- **Rich Output**: Generates manifest CSV and optional HTML index with download links
- **Robust Handling**: Sanitizes filenames, handles duplicates, and provides comprehensive error reporting
- **Beautiful Logging**: Rich console output with progress tracking and summary tables

## Installation

### From Source

```bash
# Clone the repository
git clone <repository-url>
cd sga-splitter

# Install in development mode
pip install -e .
```

### From PyPI (when published)

```bash
pip install sga-splitter
```

## Requirements

- Python 3.11+
- Dependencies are automatically installed:
  - pandas >= 2.0.0
  - openpyxl >= 3.1.0
  - xlsxwriter >= 3.1.0
  - typer[all] >= 0.9.0
  - python-slugify >= 8.0.0
  - rich >= 13.0.0

## Quick Start

### Multi-Sheet Processing (Recommended)

For workbooks with multiple sheets where you want automatic processing:

```bash
# Process all sheets with automatic logic
# Sheet 1: Split by Project, Sheets 2 & 3: Split by Department
# Removes unwanted columns and preserves formatting
sga-split multi-sheet --input "BWID SG&A Budget report_250818.xlsx"

# Custom output directory and column removal
sga-split multi-sheet --input "workbook.xlsx" --out ./Results --remove-columns "unnamed,temp,notes"
```

### Single Sheet Processing

For processing individual sheets with manual control:

```bash
# Split a workbook with auto-detection
sga-split main --input "BWID SG&A Budget report_250818.xlsx"

# Specify output directory
sga-split main --input "budget.xlsx" --out ./SGA_Splits
```

### Advanced Usage

```bash
# Use clone mode to preserve formatting
sga-split --input "budget.xlsx" --mode clone --make-index

# Enable fuzzy sheet matching and generate full reports
sga-split --input "budget.xlsx" \
  --fuzzy-sheet \
  --make-index \
  --manifest results.csv \
  --verbose

# Specify exact sheet and column names
sga-split --input "budget.xlsx" \
  --sheet "SG&A Summary Sheet" \
  --by "Department/Project" \
  --mode clone
```

## Command Line Options

### Multi-Sheet Command (`sga-split multi-sheet`)

| Option | Short | Description | Default |
|--------|-------|-------------|---------|
| `--input` | `-i` | Path to input Excel file | Required |
| `--out` | `-o` | Output directory | `./SGA_Splits` |
| `--skip-totals` / `--keep-totals` |  | Skip/include rows containing 'total' | Skip |
| `--case-insensitive` |  | Case-insensitive group matching | `False` |
| `--include-empty` |  | Include groups that would be empty | `False` |
| `--remove-columns` |  | Comma-separated column patterns to remove | "unnamed,project/department" |
| `--verbose` | `-v` | Enable verbose logging | `False` |

### Single Sheet Command (`sga-split main`)

| Option | Short | Description | Default |
|--------|-------|-------------|---------|
| `--input` | `-i` | Path to input Excel file | Required |
| `--sheet` | `-s` | Sheet name to process | Auto-detect |
| `--by` |  | Department/Project column name | Auto-detect |
| `--out` | `-o` | Output directory | `./SGA_Splits` |
| `--mode` | `-m` | Export mode: 'fast' or 'clone' | `fast` |
| `--skip-totals` / `--keep-totals` |  | Skip/include rows containing 'total' | Skip |
| `--case-insensitive` |  | Case-insensitive group matching | `False` |
| `--fuzzy-sheet` |  | Enable fuzzy sheet name matching | `False` |
| `--make-index` |  | Generate HTML index file | `False` |
| `--manifest` |  | Path for manifest CSV file | None |
| `--include-empty` |  | Include groups that would be empty | `False` |
| `--verbose` | `-v` | Enable verbose logging | `False` |

## Export Modes

### Fast Mode (Default)

- Uses pandas + xlsxwriter for maximum speed
- Applies basic formatting (frozen headers, autofilter)
- **Best for**: Large files where speed is priority
- **Trade-off**: Original cell styles and formulas are not preserved

### Clone Mode

- Uses openpyxl to create exact copies of the original workbook
- Preserves all formatting, styles, formulas, and cell properties
- Updates autofilter ranges appropriately
- **Best for**: Files where formatting must be preserved
- **Trade-off**: Slower processing, especially for large files

## Output Structure

When you run the tool, it creates:

```
SGA_Splits/
├── Sales - SG&A Summary.xlsx
├── Marketing - SG&A Summary.xlsx
├── IT - SG&A Summary.xlsx
├── manifest.csv (if --manifest specified)
└── index.html (if --make-index specified)
```

### Manifest CSV

Contains summary information about each created file:

```csv
Department/Project,output_path,row_count,mode
Sales,./SGA_Splits/Sales - SG&A Summary.xlsx,15,fast
Marketing,./SGA_Splits/Marketing - SG&A Summary.xlsx,8,fast
IT,./SGA_Splits/IT - SG&A Summary.xlsx,12,fast
```

### HTML Index

Provides a user-friendly interface with download links for all generated files.

## Auto-Detection Logic

### Sheet Detection

The tool automatically finds sheets containing SG&A data by looking for:

1. Exact match (if `--sheet` specified)
2. Fuzzy matching (if `--fuzzy-sheet` enabled):
   - Sheets containing "SG&A" or "SGA" keywords
   - Sheets containing "summary" keywords
   - Best similarity match to requested name

### Column Detection

Automatically detects Department/Project columns matching these patterns:

- `Department/Project`
- `Department - Project`
- `Dept/Project` 
- `Project/Department`
- `Department`
- `Project`
- `Dept`

Detection is case-insensitive and handles extra whitespace.

## Multi-Sheet Processing Features

### Automatic Sheet-Specific Logic

The `multi-sheet` command automatically applies the correct splitting logic based on sheet position:

- **Sheet 1**: Always splits by Project (looks for "Project", "Proj" columns)
- **Sheet 2 & 3**: Split by Department (looks for "Department", "Dept" columns)

### Column Removal

Automatically removes unwanted columns from output files:

- **Default removal patterns**: "Unnamed", "Project/Department" 
- **Custom patterns**: Use `--remove-columns` to specify your own patterns
- **Pattern matching**: Case-insensitive, partial matching

### Formatting Preservation

The clone mode preserves all original formatting:

- **Cell Styles**: Colors, fonts, borders, fills
- **Column Widths**: Maintains original column sizing
- **Table Layout**: Preserves title positions and structure
- **Formulas**: Keeps all Excel formulas intact

## Examples

### Example 1: Multi-Sheet Processing (Recommended)

```bash
sga-split multi-sheet --input "Q4_SGA_Budget.xlsx"
```

**Output**: 
- Separate directories for each sheet
- Files split by Project (Sheet 1) and Department (Sheets 2 & 3)
- Unwanted columns removed
- Original formatting preserved

### Example 2: Custom Column Removal

```bash
sga-split multi-sheet --input "Budget.xlsx" --remove-columns "unnamed,temp,notes,draft"
```

**Output**: Removes additional columns matching the specified patterns

### Example 3: Single Sheet Processing

```bash
sga-split main --input "Styled_Budget.xlsx" --mode clone --out ./Formatted_Splits
```

**Output**: Creates formatted files preserving all original styles and formulas

### Example 4: Complete Report Generation

```bash
sga-split main --input "Annual_Budget.xlsx" \
  --mode clone \
  --make-index \
  --manifest summary.csv \
  --fuzzy-sheet \
  --verbose
```

**Output**: 
- Individual department files
- `summary.csv` with processing details
- `index.html` with download links
- Detailed console logging

### Example 4: Custom Configuration

```bash
sga-split --input "Custom_Report.xlsx" \
  --sheet "Budget Summary" \
  --by "Cost Center" \
  --out ./Custom_Output \
  --keep-totals \
  --include-empty
```

**Output**: Processes specific sheet and column, includes total rows and empty groups

## Error Handling

The tool provides comprehensive error handling:

- **File not found**: Clear error message with file path
- **Invalid Excel file**: Validation of file format
- **Sheet not found**: Lists available sheets
- **Column not found**: Shows sample headers to help identify correct column
- **Permission errors**: Guidance on file access issues

## Development

### Running Tests

```bash
# Install development dependencies
pip install -e ".[dev]"

# Run tests
pytest

# Run tests with coverage
pytest --cov=sga_splitter

# Run specific test file
pytest tests/test_detect.py -v
```

### Code Quality

```bash
# Format code
black src/ tests/

# Sort imports
isort src/ tests/

# Lint code  
flake8 src/ tests/

# Type checking
mypy src/
```

## Troubleshooting

### Common Issues

1. **"Sheet not found"**: Use `--fuzzy-sheet` or check exact sheet name
2. **"Column not found"**: The tool shows sample headers to help identify the correct column pattern
3. **Empty output**: Check if `--skip-totals` is filtering out your data
4. **Permission denied**: Ensure the output directory is writable and input file is not open

### Getting Help

```bash
# Show help
sga-split --help

# Show version
sga-split version

# Show tool information
sga-split info
```

## License

MIT License - see LICENSE file for details.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Add tests for new functionality
4. Ensure all tests pass
5. Submit a pull request

## Changelog

### v0.1.0
- Initial release
- Fast and clone export modes
- Automatic sheet and column detection
- Manifest and HTML index generation
- Comprehensive test suite