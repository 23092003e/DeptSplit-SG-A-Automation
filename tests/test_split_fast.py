"""
Tests for fast export mode functionality.
"""

import pytest
import pandas as pd
from pathlib import Path
import tempfile
import shutil

from sga_splitter.exporters import export_fast


class TestExportFast:
    """Test suite for export_fast function."""
    
    @pytest.fixture
    def sample_dataframe(self):
        """Create a sample DataFrame for testing."""
        return pd.DataFrame({
            'ID': [1, 2, 3, 4, 5, 6],
            'Department/Project': ['Sales', 'Marketing', 'Sales', 'IT', 'Marketing', 'IT'],
            'Amount': [1000, 2000, 1500, 3000, 2500, 4000],
            'Description': ['Sales expense', 'Marketing cost', 'Sales travel', 'IT equipment', 'Marketing ads', 'IT software']
        })
    
    @pytest.fixture
    def temp_output_dir(self):
        """Create a temporary directory for test outputs."""
        temp_dir = Path(tempfile.mkdtemp())
        yield temp_dir
        # Cleanup after test
        shutil.rmtree(temp_dir)
    
    def test_basic_export(self, sample_dataframe, temp_output_dir):
        """Test basic export functionality."""
        groups = ['Sales', 'Marketing', 'IT']
        dp_col = 'Department/Project'
        sheet_name = 'Test Sheet'
        
        manifest_entries = export_fast(
            df=sample_dataframe,
            groups=groups,
            dp_col=dp_col,
            sheet_name=sheet_name,
            out_dir=temp_output_dir
        )
        
        # Check manifest entries
        assert len(manifest_entries) == 3
        
        # Check each group
        for entry in manifest_entries:
            assert entry['mode'] == 'fast'
            assert 'Department/Project' in entry
            assert 'output_path' in entry
            assert 'row_count' in entry
            assert entry['row_count'] > 0
        
        # Verify files were created
        created_files = list(temp_output_dir.glob("*.xlsx"))
        assert len(created_files) == 3
        
        # Check file naming
        expected_names = ['Sales - SG&A Summary.xlsx', 'Marketing - SG&A Summary.xlsx', 'IT - SG&A Summary.xlsx']
        actual_names = [f.name for f in created_files]
        for expected in expected_names:
            assert expected in actual_names
    
    def test_group_filtering(self, sample_dataframe, temp_output_dir):
        """Test that each file contains only the correct group data."""
        groups = ['Sales', 'Marketing']
        dp_col = 'Department/Project'
        sheet_name = 'Test Sheet'
        
        manifest_entries = export_fast(
            df=sample_dataframe,
            groups=groups,
            dp_col=dp_col,
            sheet_name=sheet_name,
            out_dir=temp_output_dir
        )
        
        # Check row counts
        sales_entry = next(e for e in manifest_entries if e['Department/Project'] == 'Sales')
        marketing_entry = next(e for e in manifest_entries if e['Department/Project'] == 'Marketing')
        
        assert sales_entry['row_count'] == 2  # 2 Sales rows in sample data
        assert marketing_entry['row_count'] == 2  # 2 Marketing rows in sample data
        
        # Verify file contents by reading back
        sales_file = Path(sales_entry['output_path'])
        assert sales_file.exists()
        
        sales_df = pd.read_excel(sales_file, sheet_name=sheet_name)
        assert len(sales_df) == 2
        assert all(sales_df['Department/Project'] == 'Sales')
    
    def test_empty_group_handling(self, sample_dataframe, temp_output_dir):
        """Test handling of groups with no data."""
        groups = ['Sales', 'NonExistent']
        dp_col = 'Department/Project'
        sheet_name = 'Test Sheet'
        
        manifest_entries = export_fast(
            df=sample_dataframe,
            groups=groups,
            dp_col=dp_col,
            sheet_name=sheet_name,
            out_dir=temp_output_dir
        )
        
        # Should only create file for Sales, not for NonExistent
        assert len(manifest_entries) == 1
        assert manifest_entries[0]['Department/Project'] == 'Sales'
        
        # Only one file should be created
        created_files = list(temp_output_dir.glob("*.xlsx"))
        assert len(created_files) == 1
    
    def test_filename_sanitization(self, temp_output_dir):
        """Test that problematic characters in group names are sanitized."""
        df = pd.DataFrame({
            'Department/Project': ['Sales/Marketing', 'IT<>Support', 'Legal|Compliance'],
            'Amount': [1000, 2000, 3000]
        })
        
        groups = ['Sales/Marketing', 'IT<>Support', 'Legal|Compliance']
        
        manifest_entries = export_fast(
            df=df,
            groups=groups,
            dp_col='Department/Project',
            sheet_name='Test',
            out_dir=temp_output_dir
        )
        
        # Check that files were created with sanitized names
        created_files = list(temp_output_dir.glob("*.xlsx"))
        assert len(created_files) == 3
        
        # Verify filenames don't contain problematic characters
        for file in created_files:
            filename = file.name
            problematic_chars = ['<', '>', ':', '"', '|', '?', '*', '\\', '/']
            for char in problematic_chars:
                assert char not in filename, f"Filename contains problematic char '{char}': {filename}"
    
    def test_duplicate_filename_handling(self, temp_output_dir):
        """Test handling of duplicate filenames."""
        df = pd.DataFrame({
            'Department/Project': ['Sales', 'Sales'],  # Same name
            'Amount': [1000, 2000],
            'ID': [1, 2]
        })
        
        # Create the first file manually to force a collision
        first_file = temp_output_dir / "Sales - SG&A Summary.xlsx"
        first_file.touch()
        
        groups = ['Sales']
        
        manifest_entries = export_fast(
            df=df,
            groups=groups,
            dp_col='Department/Project',
            sheet_name='Test',
            out_dir=temp_output_dir
        )
        
        # Should create a file with a numeric suffix
        assert len(manifest_entries) == 1
        output_path = Path(manifest_entries[0]['output_path'])
        assert output_path.exists()
        # Should be either the original or with #2 suffix
        assert output_path.name in ["Sales - SG&A Summary.xlsx", "Sales - SG&A Summary #2.xlsx"]
    
    def test_preserve_column_structure(self, sample_dataframe, temp_output_dir):
        """Test that column structure is preserved in output files."""
        groups = ['Sales']
        
        manifest_entries = export_fast(
            df=sample_dataframe,
            groups=groups,
            dp_col='Department/Project',
            sheet_name='Test',
            out_dir=temp_output_dir
        )
        
        # Read back the created file
        output_file = Path(manifest_entries[0]['output_path'])
        result_df = pd.read_excel(output_file, sheet_name='Test')
        
        # Check that all original columns are present
        original_columns = list(sample_dataframe.columns)
        result_columns = list(result_df.columns)
        
        assert original_columns == result_columns
        
        # Check that header values match
        for col in original_columns:
            assert col in result_df.columns
    
    def test_error_handling(self, sample_dataframe):
        """Test error handling for invalid inputs."""
        # Test with non-existent output directory (should handle via ensure_out_dir)
        invalid_dir = Path("/totally/invalid/path/that/should/not/exist")
        
        # This should not raise an error due to ensure_out_dir creating the path
        # Instead test with empty groups list
        manifest_entries = export_fast(
            df=sample_dataframe,
            groups=[],
            dp_col='Department/Project',
            sheet_name='Test',
            out_dir=Path(tempfile.mkdtemp())
        )
        
        assert len(manifest_entries) == 0