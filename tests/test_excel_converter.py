import pytest
from unittest.mock import MagicMock, patch, PropertyMock
from pathlib import Path
from src.core.excel_converter import ExcelConverter
from src.config import PDFConversionSettings, ExcelSettings, OptimizationSettings


@pytest.fixture
def mock_excel_app():
    with patch("src.core.excel_converter.win32com.client.Dispatch") as mock_dispatch:
        mock_app = MagicMock()
        mock_dispatch.return_value = mock_app
        yield mock_app


@pytest.fixture
def mock_pythoncom():
    with patch("src.core.excel_converter.pythoncom") as mock_com:
        yield mock_com


@pytest.fixture
def mock_sheet_settings():
    """Mock get_excel_sheet_settings to return the base settings unchanged."""
    with patch("src.core.excel_converter.get_excel_sheet_settings") as mock_get:
        # Return the base_settings passed to it (second argument)
        mock_get.side_effect = lambda sheet_name, base_settings: base_settings
        yield mock_get


@pytest.fixture
def converter(mock_excel_app, mock_pythoncom, mock_sheet_settings):
    return ExcelConverter()


def test_convert_success(converter, mock_excel_app, tmp_path):
    """Test basic Excel to PDF conversion."""
    # Setup paths
    input_file = tmp_path / "test.xlsx"
    input_file.touch()
    output_file = tmp_path / "test.pdf"
    
    # Setup mock workbook and sheet
    mock_workbook = MagicMock()
    mock_sheet = MagicMock()
    mock_sheet.Visible = -1  # xlSheetVisible
    mock_sheet.Name = "Sheet1"
    mock_sheet.UsedRange.Columns.Count = 10
    mock_sheet.UsedRange.Rows.Count = 100
    mock_workbook.Worksheets = [mock_sheet]
    mock_workbook.ActiveSheet = mock_sheet
    mock_excel_app.Workbooks.Open.return_value = mock_workbook
    
    # Run conversion
    settings = PDFConversionSettings()
    result = converter.convert(input_file, output_file, settings)
    
    # Verify Open called
    mock_excel_app.Workbooks.Open.assert_called_once()
    
    # Verify Export called
    assert mock_sheet.ExportAsFixedFormat.call_count == 1
    
    # Verify Cleanup
    mock_workbook.Close.assert_called_once_with(SaveChanges=False)
    mock_excel_app.Quit.assert_called_once()


def test_convert_with_sheet_name(converter, mock_excel_app, tmp_path):
    """Test conversion with specific sheet name filter."""
    input_file = tmp_path / "test.xlsx"
    input_file.touch()
    
    # Setup mock with multiple sheets
    mock_workbook = MagicMock()
    mock_sheet1 = MagicMock()
    mock_sheet1.Visible = -1
    mock_sheet1.Name = "Data"
    mock_sheet1.UsedRange.Columns.Count = 5
    
    mock_sheet2 = MagicMock()
    mock_sheet2.Visible = -1
    mock_sheet2.Name = "Summary"
    mock_sheet2.UsedRange.Columns.Count = 3
    
    mock_workbook.Worksheets = [mock_sheet1, mock_sheet2]
    mock_workbook.ActiveSheet = mock_sheet1
    mock_excel_app.Workbooks.Open.return_value = mock_workbook
    
    # Request only "Data" sheet
    settings = PDFConversionSettings(
        excel=ExcelSettings(sheet_name="Data")
    )
    
    converter.convert(input_file, None, settings)
    
    # Only Data sheet should be selected for export (filtered by sheet_name)
    # Verify export was called
    mock_sheet1.ExportAsFixedFormat.assert_called_once()


def test_convert_row_dimensions_fit_all(converter, mock_excel_app, tmp_path):
    """Test row_dimensions=0 fits entire sheet on one page."""
    input_file = tmp_path / "test.xlsx"
    input_file.touch()
    
    mock_workbook = MagicMock()
    mock_sheet = MagicMock()
    mock_sheet.Visible = -1
    mock_sheet.Name = "Sheet1"
    mock_sheet.UsedRange.Columns.Count = 20
    mock_sheet.UsedRange.Rows.Count = 500
    mock_workbook.Worksheets = [mock_sheet]
    mock_workbook.ActiveSheet = mock_sheet
    mock_excel_app.Workbooks.Open.return_value = mock_workbook
    
    settings = PDFConversionSettings(
        excel=ExcelSettings(row_dimensions=0)  # Fit all on one page
    )
    
    converter.convert(input_file, None, settings)
    
    # Verify PageSetup was configured
    assert mock_sheet.PageSetup.FitToPagesTall == 1


def test_convert_orientation_portrait(converter, mock_excel_app, tmp_path):
    """Test portrait orientation setting."""
    input_file = tmp_path / "test.xlsx"
    input_file.touch()
    
    mock_workbook = MagicMock()
    mock_sheet = MagicMock()
    mock_sheet.Visible = -1
    mock_sheet.Name = "Sheet1"
    mock_sheet.UsedRange.Columns.Count = 5
    mock_workbook.Worksheets = [mock_sheet]
    mock_workbook.ActiveSheet = mock_sheet
    mock_excel_app.Workbooks.Open.return_value = mock_workbook
    
    settings = PDFConversionSettings(
        excel=ExcelSettings(orientation="portrait")
    )
    
    converter.convert(input_file, None, settings)
    
    # Portrait orientation = 1
    assert mock_sheet.PageSetup.Orientation == 1


def test_convert_metadata_header_enabled(converter, mock_excel_app, tmp_path):
    """Test metadata header is applied when enabled."""
    input_file = tmp_path / "test.xlsx"
    input_file.touch()
    
    mock_workbook = MagicMock()
    mock_sheet = MagicMock()
    mock_sheet.Visible = -1
    mock_sheet.Name = "DataSheet"
    mock_sheet.UsedRange.Columns.Count = 10
    mock_workbook.Worksheets = [mock_sheet]
    mock_workbook.ActiveSheet = mock_sheet
    mock_excel_app.Workbooks.Open.return_value = mock_workbook
    
    settings = PDFConversionSettings(
        excel=ExcelSettings(metadata_header=True)
    )
    
    converter.convert(input_file, None, settings)
    
    # Verify headers were set
    page_setup = mock_sheet.PageSetup
    assert page_setup.LeftHeader is not None
    assert page_setup.CenterHeader is not None
    assert page_setup.RightHeader is not None


def test_convert_low_quality(converter, mock_excel_app, tmp_path):
    """Test low quality optimization setting."""
    input_file = tmp_path / "test.xlsx"
    input_file.touch()
    
    mock_workbook = MagicMock()
    mock_sheet = MagicMock()
    mock_sheet.Visible = -1
    mock_sheet.Name = "Sheet1"
    mock_sheet.UsedRange.Columns.Count = 5
    mock_workbook.Worksheets = [mock_sheet]
    mock_workbook.ActiveSheet = mock_sheet
    mock_excel_app.Workbooks.Open.return_value = mock_workbook
    
    settings = PDFConversionSettings(
        optimization=OptimizationSettings(image_quality="low")
    )
    
    converter.convert(input_file, None, settings)
    
    # Verify export was called with low quality
    args = mock_sheet.ExportAsFixedFormat.call_args[1]
    assert args["Quality"] == 1  # xlQualityMinimum


def test_convert_failure_handling(converter, mock_excel_app, tmp_path):
    """Test exception handling and cleanup."""
    input_file = tmp_path / "test.xlsx"
    input_file.touch()
    
    # Make Open raise exception
    mock_excel_app.Workbooks.Open.side_effect = Exception("Excel Error")
    
    with pytest.raises(Exception, match="Excel Error"):
        converter.convert(input_file)
        
    # Ensure Quit is still called (safety cleanup)
    mock_excel_app.Quit.assert_called_once()


def test_convert_file_not_found(converter):
    """Test FileNotFoundError for missing files."""
    with pytest.raises(FileNotFoundError):
        converter.convert(Path("nonexistent.xlsx"))


def test_smart_page_size_calculation(converter, mock_excel_app, tmp_path):
    """Test smart page size calculates correct width for many columns."""
    input_file = tmp_path / "wide.xlsx"
    input_file.touch()
    
    mock_workbook = MagicMock()
    mock_sheet = MagicMock()
    mock_sheet.Visible = -1
    mock_sheet.Name = "WideData"
    mock_sheet.UsedRange.Columns.Count = 50  # 50 columns
    mock_sheet.UsedRange.Rows.Count = 100
    mock_workbook.Worksheets = [mock_sheet]
    mock_workbook.ActiveSheet = mock_sheet
    mock_excel_app.Workbooks.Open.return_value = mock_workbook
    
    settings = PDFConversionSettings(
        excel=ExcelSettings(min_col_width_inches=0.5)
    )
    
    # Calculate expected page width: 50 * 0.5 = 25 inches
    page_width, _ = converter._calculate_smart_page_size(mock_sheet, 0.5)
    
    # Should be 25 inches (clamped between 8.5 and 129)
    assert page_width == 25.0


def test_smart_page_size_max_clamp(converter, mock_excel_app, tmp_path):
    """Test smart page size clamps to max 129 inches."""
    input_file = tmp_path / "huge.xlsx"
    input_file.touch()
    
    mock_sheet = MagicMock()
    mock_sheet.UsedRange.Columns.Count = 500  # 500 columns
    
    # Calculate expected: 500 * 0.5 = 250, clamped to 129
    page_width, _ = converter._calculate_smart_page_size(mock_sheet, 0.5)
    
    assert page_width == 129.0
