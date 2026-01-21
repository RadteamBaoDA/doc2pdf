"""
Excel to PDF Converter using pywin32 COM.

Features:
- Smart Page Size for OCR optimization
- Dynamic page width based on column count
- Configurable row dimensions for vertical pagination
- Metadata headers (sheet name, row range, filename)
"""
import sys
from pathlib import Path
from typing import Optional, List, Tuple, Literal, Callable
import win32com.client
import win32print
import pythoncom
import dataclasses
from contextlib import contextmanager

from .base import Converter
from ..config import PDFConversionSettings, ExcelSettings, get_excel_sheet_settings
from ..utils.logger import logger
from ..utils.process_manager import ProcessRegistry

# Excel constants from Object Model
xlTypePDF = 0
xlQualityStandard = 0
xlQualityMinimum = 1
xlLandscape = 2
xlPortrait = 1
xlPaperLetter = 1
xlPaperA4 = 9
xlPaperA3 = 8
xlPaperA2 = 66  # 16.5x23.4 in
xlPaperTabloid = 3  # 11x17 in
xlPaperLegal = 5  # 8.5x14 in
xlPaperLedger = 4  # 17x11 in (Tabloid rotated)
xlPaperB4 = 12  # 9.84x13.9 in (JIS B4)
xlPaperB3 = 13  # 13.9x19.7 in (JIS B3)
# Architecture sizes
xlPaperC = 24  # 17x22 in (Arch C)
xlPaperD = 25  # 22x34 in (Arch D)
xlPaperE = 26  # 34x44 in (Arch E)

# Page Setup constants
xlFitToPage = 2
xlPrintNoComments = -4142

# Worksheet visibility
xlSheetVisible = -1


class ExcelConverter(Converter):
    """
    Converter for Excel documents (.xlsx, .xls, .xlsm, .xlsb) to PDF.
    
    Features Smart Page Size for OCR optimization - ensures all columns
    are readable by OCR tools like miner U, Deepseek OCR, RAGFlow deepdoc.
    """
    
    # Maximum paper dimensions in Excel (inches)
    MAX_PAGE_WIDTH_INCHES = 129.0
    MIN_PAGE_WIDTH_INCHES = 8.5
    DEFAULT_PAGE_HEIGHT_INCHES = 11.0
    POINTS_PER_INCH = 72
    
    # Search Constants
    xlByRows = 1
    xlByColumns = 2
    xlPrevious = 2
    
    def convert(self, input_file: Path, output_file: Path, settings: PDFConversionSettings, on_progress: Optional[Callable[[float], None]] = None) -> None:
        """
        Convert Excel file to PDF using COM automation.
        
        Args:
            input_file: Path to source Excel file
            output_file: Path to destination PDF file
            settings: PDFConversionSettings object containing conversion configuration
            on_progress: Optional callback for partial progress (0.0 to 1.0)
        """    
        input_file = input_file.resolve()
        if not input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")
            
        if output_file:
            out_file = output_file.resolve()
        else:
            out_file = input_file.with_suffix(".pdf")
            
        # Ensure output directory exists
        out_file.parent.mkdir(parents=True, exist_ok=True)
            
        # settings is PDFConversionSettings
        excel_settings = settings.excel or ExcelSettings()
        
        logger.info(f"Converting '{input_file.name}' to PDF...")
        logger.debug(f"Settings: {settings}")

        # Ensure CoInitialize is called for this thread
        pythoncom.CoInitialize()
        
        try:
            with self._excel_application() as excel:
                workbook = None
                temp_sheets_to_delete = []
                final_sheets_to_process = []
                
                try:
                    # Open Workbook (ReadOnly for safety)
                    workbook = excel.Workbooks.Open(
                        str(input_file), 
                        ReadOnly=True,
                        UpdateLinks=False,
                        IgnoreReadOnlyRecommended=True
                    )
                    
                    # Get sheets to process
                    sheets_to_export = self._get_sheets_to_export(workbook, excel_settings)
                    
                    if not sheets_to_export:
                        logger.warning(f"No visible sheets found in {input_file.name}")
                        raise ValueError(f"No visible sheets to export in {input_file.name}")
                    
                    # Calculate progress weight per sheet
                    total_sheets = len(sheets_to_export)
                    sheet_weight = 1.0 / total_sheets if total_sheets > 0 else 0
                    
                    # Apply page setup and process chunks
                    for sheet in sheets_to_export:
                        # Get sheet-specific settings
                        # Note: Arguments are (sheet_name, base_settings, input_path)
                        sheet_settings = get_excel_sheet_settings(sheet.Name, settings, input_file)
                        sheet_excel_settings = sheet_settings.excel or excel_settings
                        
                        logger.debug(f"Sheet '{sheet.Name}' settings: row_dimensions={sheet_excel_settings.row_dimensions}")
                        
                        # Determine True Bounds (Last Col/Row with actual data)
                        # This avoids issues where UsedRange is huge due to formatting
                        last_row, last_col = self._get_true_used_bounds(sheet)
                        last_col_alpha = self._col_num_to_letter(last_col)
                        
                        # Check for Chunking
                        row_lim = sheet_excel_settings.row_dimensions
                        if row_lim and row_lim > 0:
                            # Chunking Mode
                            # Use true last_row instead of UsedRange
                            if last_row == 0:
                                # Empty sheet
                                if on_progress: on_progress(sheet_weight)
                                continue
                                
                            chunks = (last_row + row_lim - 1) // row_lim
                            logger.info(f"Splitting sheet '{sheet.Name}' into {chunks} chunks (Rows: {row_lim})")
                            
                            # Weight for each chunk
                            chunk_weight = sheet_weight / chunks
                            
                            for i in range(chunks):
                                start_row = i * row_lim + 1
                                end_row = min((i + 1) * row_lim, last_row)
                                
                                # Copy sheet to end
                                last_sheet = workbook.Sheets(workbook.Sheets.Count)
                                sheet.Copy(None, last_sheet)
                                new_sheet = workbook.Sheets(workbook.Sheets.Count)
                                
                                temp_sheets_to_delete.append(new_sheet)
                                
                                # Set Print Area explicitly to True content columns
                                # Format: A{start}:{LastColAlpha}{end} e.g. "A1:Z50"
                                new_sheet.PageSetup.PrintArea = f"$A${start_row}:${last_col_alpha}${end_row}"
                                
                                # Create chunk settings
                                chunk_settings = ExcelSettings(**dataclasses.asdict(sheet_excel_settings))
                                chunk_settings.row_dimensions = 0 # Force 1 page tall
                                
                                self._apply_page_setup(new_sheet, chunk_settings, input_file.name, last_col)

                                if on_progress:
                                    on_progress(chunk_weight)
                                
                                if sheet_excel_settings.metadata_header:
                                    center_text = f"{start_row}-{end_row}"
                                    self._apply_metadata_header(new_sheet, sheet_excel_settings, input_file.name, center_text, left_text=sheet.Name)
                                    
                                final_sheets_to_process.append(new_sheet)
                        else:
                            # Standard Mode
                            # Set print area to avoid printing 1000 blank pages of formatting
                            if last_row > 0:
                                sheet.PageSetup.PrintArea = f"$A$1:${last_col_alpha}${last_row}"
                            
                            self._apply_page_setup(sheet, sheet_excel_settings, input_file.name, last_col)
                            if sheet_excel_settings.metadata_header:
                                self._apply_metadata_header(sheet, sheet_excel_settings, input_file.name, center_text="")
                            final_sheets_to_process.append(sheet)
                            
                            if on_progress:
                                on_progress(sheet_weight)
                    
                    # Export to PDF
                    if final_sheets_to_process:
                        self._export_to_pdf(workbook, final_sheets_to_process, str(out_file), settings)
                        logger.success(f"Successfully converted: {out_file}")
                    else:
                        logger.warning("No content to export.")
                    
                except Exception as e:
                    logger.error(f"Failed to convert {input_file.name}: {e}")
                    raise
                finally:
                    # Cleanup temps
                    if temp_sheets_to_delete:
                        excel.DisplayAlerts = False
                        for t in temp_sheets_to_delete:
                            try:
                                t.Delete()
                            except:
                                pass
                    
                    if workbook:
                        workbook.Close(SaveChanges=False)
        finally:
            pythoncom.CoUninitialize()
            
        return out_file

    @contextmanager
    def _excel_application(self):
        """Context manager for Excel COM application lifecycle."""
        excel = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            
            # Try to set optimal printer
            self._set_optimal_printer(excel)
            
            ProcessRegistry.register(excel)
            yield excel
        except Exception as e:
            logger.critical(f"Failed to initialize Microsoft Excel: {e}")
            raise
        finally:
            if excel:
                ProcessRegistry.unregister(excel)
                excel.Quit()

    def _set_optimal_printer(self, excel) -> None:
        """
        Attempt to set ActivePrinter to 'Microsoft Print to PDF' for better paper size support.
        Uses win32print API for reliable port detection, with brute-force fallback.
        """
        target_name = "Microsoft Print to PDF"
        
        try:
            # Check if already active
            current = excel.ActivePrinter
            if target_name in current:
                logger.debug(f"'{target_name}' is already the active printer.")
                return
        except:
            pass

        # Strategy 1: Use win32print API for reliable port detection
        port_name = None
        try:
            handle = win32print.OpenPrinter(target_name)
            try:
                # Level 5 is lightweight, contains pPortName
                info = win32print.GetPrinter(handle, 5)
                port_name = info.get('pPortName', '')
                
                # Fallback to Level 2 if Level 5 didn't have port
                if not port_name:
                    info = win32print.GetPrinter(handle, 2)
                    port_name = info.get('pPortName', '')
            finally:
                win32print.ClosePrinter(handle)
        except Exception as e:
            logger.debug(f"OpenPrinter/GetPrinter API failed for '{target_name}': {e}")

        # If we got a port name from the API and it's not PORTPROMPT, try it first
        candidates = []
        if port_name and port_name.upper() != 'PORTPROMPT:':
            candidates.append(f"{target_name} on {port_name}")
        
        # Strategy 2: Brute force Ne00-Ne99 as fallback (expanded range)
        for i in range(100):
            candidates.append(f"{target_name} on Ne{i:02d}:")
            
        # Strategy 3: Naked name (rare)
        candidates.append(target_name)
        
        success = False
        for candidate in candidates:
            try:
                excel.ActivePrinter = candidate
                logger.info(f"Successfully switched ActivePrinter to: '{candidate}'")
                success = True
                break
            except Exception as e:
                # Only log first few failures to avoid spam
                if candidates.index(candidate) < 5:
                    logger.debug(f"Failed to set ActivePrinter to '{candidate}': {e}")
                
        if not success:
            logger.warning(
                f"Could not set ActivePrinter to '{target_name}'. "
                f"Using default: '{excel.ActivePrinter}'. "
                "Large paper sizes (A3) may rely on the default printer's capabilities."
            )

    def _get_sheets_to_export(self, workbook, excel_settings: ExcelSettings) -> List:
        """Get list of sheets to export based on settings."""
        sheets = []
        
        for sheet in workbook.Worksheets:
            # Only process visible sheets
            if sheet.Visible != xlSheetVisible:
                continue
                
            # Filter by sheet name if specified
            if excel_settings.sheet_name:
                if sheet.Name != excel_settings.sheet_name:
                    continue
            
            sheets.append(sheet)
            logger.debug(f"Will export sheet: {sheet.Name}")
        
        return sheets

    def _calculate_smart_page_size(
        self, 
        sheet, 
        last_col_index: int
    ) -> Tuple[float, float]:
        """
        Calculate page width based on actual column widths of used range.
        
        Args:
            sheet: Excel Worksheet object
            last_col_index: The 1-based index of the last used column (e.g. 5 for Column E)
            
        Returns:
            Tuple of (page_width_inches, page_height_inches)
        """
        try:
            if last_col_index < 1:
                return self.MIN_PAGE_WIDTH_INCHES, self.DEFAULT_PAGE_HEIGHT_INCHES
                
            # Measure width of Range(A:LastCol)
            # content_width_points = range.Width
            
            first_col_char = "A"
            last_col_char = self._col_num_to_letter(last_col_index)
            col_range = sheet.Range(f"{first_col_char}1:{last_col_char}1")
            
            # .Width corresponds to the width in points of the range
            content_width_points = col_range.Width
            content_width_inches = content_width_points / self.POINTS_PER_INCH
            
            # Add a small buffer for margins (0.5 inch total)
            page_width = content_width_inches + 0.5
            
            # Page height - defaults
            page_height = self.DEFAULT_PAGE_HEIGHT_INCHES
            
            logger.debug(
                f"Sheet '{sheet.Name}' (Cols 1-{last_col_index}): "
                f"Content Width: {content_width_inches:.2f}\" -> Page Width (w/ margins): {page_width:.2f}\""
            )
            
            return page_width, page_height
            
        except Exception as e:
            logger.warning(f"Could not calculate smart page size: {e}")
            return self.MIN_PAGE_WIDTH_INCHES, self.DEFAULT_PAGE_HEIGHT_INCHES

    def _apply_page_setup(
        self, 
        sheet, 
        excel_settings: ExcelSettings,
        filename: str,
        last_col: int
    ) -> None:
        """
        Apply page setup settings for OCR-optimized PDF output.
        
        Args:
            sheet: Excel Worksheet object
            excel_settings: Excel-specific settings
            filename: Original filename for header
            last_col: Last used column index for width calculation
        """
        try:
            page_setup = sheet.PageSetup
            
            # Calculate smart page size
            page_width, page_height = self._calculate_smart_page_size(
                sheet, 
                last_col
            )
            
            # Set orientation
            # If width > 8.5 OR explicitly set to landscape, try landscape.
            # But normally we auto-detect orientation based on width ratio?
            # For now, stick to settings or default to portrait if narrow, landscape if wide?
            # User requirement: "base on current size of sheet choose page size"
            
            # Force orientation based on content?
            # If content is wider than 8.5 but less than 11, Landscape Letter is better than Portrait Letter?
            # Let's trust the settings or default to Portrait unless it's very wide.
            
            if excel_settings.orientation.lower() == "landscape":
                page_setup.Orientation = xlLandscape
            elif excel_settings.orientation.lower() == "portrait":
                page_setup.Orientation = xlPortrait
            else:
                 # Auto orientation
                 if page_width > 8.5:
                     page_setup.Orientation = xlLandscape
                 else:
                     page_setup.Orientation = xlPortrait

            is_landscape = (page_setup.Orientation == xlLandscape)
            
            # Define ladder of supported sizes
            # Format: (Enum, WidthLimit (inches), Name)
            # WidthLimit: The maximum content width this paper size can effectively hold (considering margins/orientation)
            
            # Standard sizes only first? Or mix?
            # Lets define physically available constraints.
            
            if is_landscape:
                # Width is the longer dimension
                 paper_ladder = [
                    (xlPaperLetter, 11.0, "Letter"),        # 11 wide
                    (xlPaperLegal, 14.0, "Legal"),          # 14 wide
                    (xlPaperA3, 16.54, "A3"),               # 16.54 wide
                    (xlPaperTabloid, 17.0, "Tabloid"),      # 17 wide
                    (xlPaperA2, 23.39, "A2"),               # 23.39 wide
                    (xlPaperD, 34.0, "Arch D"),             # 34 wide
                    (xlPaperE, 44.0, "Arch E"),             # 44 wide
                ]
            else:
                # Width is the shorter dimension
                paper_ladder = [
                    (xlPaperLetter, 8.5, "Letter"),         # 8.5 wide
                    (xlPaperLegal, 8.5, "Legal"),           # 8.5 wide (Legal is just taller)
                    (xlPaperA3, 11.69, "A3"),               # 11.69 wide
                    (xlPaperTabloid, 11.0, "Tabloid"),      # 11 wide
                    (xlPaperA2, 16.54, "A2"),               # 16.54 wide
                    (xlPaperD, 22.0, "Arch D"),             # 22, wide
                    (xlPaperE, 34.0, "Arch E"),             # 34 wide
                ]

            selected_paper = None
            selected_name = None
            limit_width = 8.5
            
            # 1. Find the Smallest Fit and fallback candidates
            candidates = []
            best_fit_index = -1
            
            # Find the index of the first size that fits
            for i, (enum_val, width_limit, name) in enumerate(paper_ladder):
                if width_limit >= page_width:
                    best_fit_index = i
                    break
            
            if best_fit_index != -1:
                # Try all sizes from best fit upwards
                candidates = paper_ladder[best_fit_index:]
            else:
                # Content exceeds all standard sizes, try the largest one
                candidates = [paper_ladder[-1]]
                oversized = True
            
            # 2. Try to set valid paper size
            for (enum_to_try, limit_to_try, name_to_try) in candidates:
                try:
                    page_setup.PaperSize = enum_to_try
                    # Verify
                    if page_setup.PaperSize == enum_to_try:
                        selected_paper = enum_to_try
                        selected_name = name_to_try
                        limit_width = limit_to_try
                        logger.info(f"Selected paper size: {selected_name} for width {page_width:.2f}\"")
                        paper_set_success = True
                        break
                    else:
                        logger.debug(f"Printer rejected paper size {name_to_try} (Enum {enum_to_try}). Trying next larger size...")
                except Exception as e:
                    logger.debug(f"Failed to set paper size to {name_to_try}: {e}")
                    continue
            
            if not paper_set_success:
                logger.warning(f"Could not set any appropriate paper size for width {page_width:.2f}\". Printer may lack support for large sizes.")
                # We failed to set any size. We are stuck with whatever default is.
                # Check what we have.
                try:
                    current_size = page_setup.PaperSize
                    # Try to map back to name? 
                    # Just warn and proceed, or error?
                    # If we are oversized, we need check threshold against the *attempted* largest?
                    pass
                except:
                    pass

            # 3. Validation and Error (The "Make this file error" requirement)
            # If still oversized despite using largest paper, OR if we failed to set it and standard one is too small.
            
            # Re-read what we actually have
            # current_paper_width? Not easily accessible directly without a map.
            # We assume best effort was made.
            
            if oversized and not paper_set_success:
                 # Check threshold against the largest size we *tried* to set (Arch E)
                 limit_to_try = paper_ladder[-1][1]
                 name_to_try = paper_ladder[-1][2]
                 
                 shrink_factor = limit_to_try / page_width
                 if shrink_factor < 0.8:
                     err_msg = (
                        f"Content is too wide ({page_width:.2f}\") for the largest supported paper '{name_to_try}' ({limit_to_try:.2f}\"). "
                        f"Shrink factor {shrink_factor:.2f} is below 0.8 threshold. Cannot convert safely."
                     )
                     logger.error(err_msg)
                     raise ValueError(err_msg)
                 else:
                     logger.warning(f"Content slightly larger than {name_to_try}. Shrinking to fit (Factor: {shrink_factor:.2f})")
            elif paper_set_success and oversized:
                 # We successfully set the largest size, but content is still bigger than it
                 # Check threshold
                 shrink_factor = limit_width / page_width
                 if shrink_factor < 0.8:
                     err_msg = (
                        f"Content is too wide ({page_width:.2f}\") for selected paper '{selected_name}' ({limit_width:.2f}\"). "
                        f"Shrink factor {shrink_factor:.2f} is below 0.8 threshold. Cannot convert safely."
                     )
                     logger.error(err_msg)
                     raise ValueError(err_msg)

            # 4. Final Setup
            page_setup.Zoom = False
            page_setup.FitToPagesWide = 1
            self._apply_row_dimensions(sheet, page_setup, excel_settings)
            
            # Margins
            margin_points = 36 # 0.5 inch
            top_margin = 72 if excel_settings.metadata_header else 36
            page_setup.LeftMargin = margin_points
            page_setup.RightMargin = margin_points
            page_setup.TopMargin = top_margin
            page_setup.BottomMargin = margin_points
            
        except ValueError:
            raise
        except Exception as e:
            logger.warning(f"Could not apply some page setup settings for '{sheet.Name}': {e}")

    def _apply_row_dimensions(self, sheet, page_setup, excel_settings: ExcelSettings) -> None:
        """Apply vertical pagination based on row_dimensions."""
        if excel_settings.row_dimensions == 0:
            # Fit entire sheet on one page
            page_setup.FitToPagesTall = 1
        elif excel_settings.row_dimensions:
            # Multiple pages based on row count
            used_rows = sheet.UsedRange.Rows.Count
            pages_tall = max(1, (used_rows + excel_settings.row_dimensions - 1) // excel_settings.row_dimensions)
            page_setup.FitToPagesTall = pages_tall
        else:
            # Auto - let Excel decide
            page_setup.FitToPagesTall = False

    def _apply_metadata_header(
        self, 
        sheet, 
        excel_settings: ExcelSettings,
        filename: str,
        center_text: str = "",
        left_text: str = None
    ) -> None:
        """
        Set header text: sheet name | row range | filename
        """
        try:
            page_setup = sheet.PageSetup
            
            # Left header: Sheet name (or custom)
            page_setup.LeftHeader = left_text if left_text else "&A"
            
            # Center header: Custom text (Row range) or empty
            page_setup.CenterHeader = center_text
            
            # Right header: Filename + Page Number
            # Format: filename (Page X)
            page_setup.RightHeader = f"{filename} (Page &P)"
            
            # Clear footers to avoid clutter and potential crop issues
            page_setup.RightFooter = ""
            page_setup.CenterFooter = ""
            page_setup.LeftFooter = ""
            
            logger.debug(f"Applied metadata header for sheet '{sheet.Name}' (Center: '{center_text}')")
            
        except Exception as e:
            logger.warning(f"Could not apply metadata header for '{sheet.Name}': {e}")
            
        except Exception as e:
            logger.warning(f"Could not apply metadata header for '{sheet.Name}': {e}")

    def _get_true_used_bounds(self, sheet) -> Tuple[int, int]:
        """
        Find the last row and column with actual data (ignoring empty formatting).
        Returns (last_row_index, last_col_index).
        """
        try:
            # Strategies for finding last cell
            # 1. UsedRange (often overestimated)
            # 2. Cells.Find("*", SearchOrder=xlByRows/xlByColumns, SearchDirection=xlPrevious)
            
            # Find Last Row
            try:
                # Find any non-empty cell, searching backwards from A1 (wrapping to end)
                last_row_cell = sheet.Cells.Find(
                    What="*",
                    After=sheet.Range("A1"),
                    LookIn=win32com.client.constants.xlValues if hasattr(win32com.client.constants, 'xlValues') else -4163, # xlValues
                    LookAt=2, # xlPart
                    SearchOrder=self.xlByRows,
                    SearchDirection=self.xlPrevious
                )
                last_row = last_row_cell.Row if last_row_cell else 0
            except Exception:
                # Fallback to UsedRange if Find fails (e.g. empty sheet)
                last_row = sheet.UsedRange.Rows.Count
            
            # Find Last Column
            try:
                last_col_cell = sheet.Cells.Find(
                    What="*",
                    After=sheet.Range("A1"),
                    LookIn=win32com.client.constants.xlValues if hasattr(win32com.client.constants, 'xlValues') else -4163,
                    LookAt=2,
                    SearchOrder=self.xlByColumns,
                    SearchDirection=self.xlPrevious
                )
                last_col = last_col_cell.Column if last_col_cell else 0
            except Exception:
                last_col = sheet.UsedRange.Columns.Count

            # If find returned 0 but UsedRange has data, compare?
            # Usually Find("*") returns None if sheet is purely empty.
            if last_row == 0 and last_col == 0:
                # Double check UsedRange count for safety? NO, if Find("*") fails, it's empty.
                return 0, 0
                
            # Check for Shapes/Objects (Images, Charts, etc.)
            try:
                for shape in sheet.Shapes:
                    try:
                        # Get bottom-right cell of the shape
                        br_cell = shape.BottomRightCell
                        if br_cell:
                            last_row = max(last_row, br_cell.Row)
                            last_col = max(last_col, br_cell.Column)
                            logger.debug(f"Shape '{shape.Name}' extends bounds to Row {br_cell.Row}, Col {br_cell.Column}")
                    except Exception:
                        # Some shapes might not avail BottomRightCell (e.g. comments sometimes)
                        continue
            except Exception as e:
                logger.warning(f"Failed to check shapes for bounds detection: {e}")

            return last_row, last_col
            
        except Exception as e:
            logger.warning(f"Error determining true bounds, falling back to UsedRange: {e}")
            ur = sheet.UsedRange
            return ur.Rows.Count, ur.Columns.Count

    def _col_num_to_letter(self, n: int) -> str:
        """Convert 1-based column number to Excel column letter (e.g. 1->A, 27->AA)."""
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

    def _export_to_pdf(
        self, 
        workbook, 
        sheets: List,
        output_path: str,
        settings: PDFConversionSettings
    ) -> None:
        """
        Export sheets to PDF.
        
        Args:
            workbook: Excel Workbook object
            sheets: List of sheets to export
            output_path: Path for output PDF
            settings: PDF conversion settings
        """
        try:
            # Determine quality
            quality = xlQualityStandard
            if settings.optimization.image_quality == "low":
                quality = xlQualityMinimum

            if len(sheets) == 1:
                # Export single sheet directly
                sheets[0].ExportAsFixedFormat(
                    Type=xlTypePDF,
                    Filename=output_path,
                    Quality=quality,
                    IncludeDocProperties=settings.metadata.include_properties,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False
                )
            else:
                # Multiple sheets: Copy to new temporary workbook iteratively
                logger.debug(f"Preparing to copy {len(sheets)} sheets to new workbook.")
                
                # Copy first sheet -> Creates new Workbook
                sheets[0].Copy()
                temp_wb = workbook.Application.ActiveWorkbook
                
                logger.debug(f"Created temp WB. Sheets count: {temp_wb.Sheets.Count}")
                
                # Copy remaining sheets into the new workbook
                for idx, s in enumerate(sheets[1:], start=2):
                    try:
                        last_sheet = temp_wb.Sheets(temp_wb.Sheets.Count)
                        # Use positional arguments for Copy: Copy(Before, After)
                        # s.Copy(None, last_sheet) -> Copy After last_sheet
                        s.Copy(None, last_sheet)
                        logger.debug(f"Copied sheet {idx}/{len(sheets)}. New count: {temp_wb.Sheets.Count}")
                    except Exception as copy_err:
                        logger.error(f"Failed to copy sheet {idx}: {copy_err}")
                
                try:
                    # Select all sheets in the new workbook (Explicit)
                    count = temp_wb.Sheets.Count
                    logger.debug(f"Exporting created workbook with {count} sheets.")
                    
                    if count > 1:
                        temp_wb.Sheets(1).Select(Replace=True)
                        for i in range(2, count + 1):
                            temp_wb.Sheets(i).Select(Replace=False)
                            
                    sel_count = temp_wb.Windows(1).SelectedSheets.Count
                    logger.debug(f"Selected {sel_count} sheets for export.")
                            
                    temp_wb.ExportAsFixedFormat(
                        Type=xlTypePDF,
                        Filename=output_path,
                        Quality=quality,
                        IncludeDocProperties=settings.metadata.include_properties,
                        IgnorePrintAreas=False,
                        OpenAfterPublish=False
                    )
                finally:
                    temp_wb.Close(SaveChanges=False)

            
        except Exception as e:
            logger.error(f"Failed to export to PDF: {e}")
            raise
