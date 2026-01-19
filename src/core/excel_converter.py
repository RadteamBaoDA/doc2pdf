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
from typing import Optional, List, Tuple
import win32com.client
import win32print
import pythoncom
import dataclasses
from contextlib import contextmanager

from .base import Converter
from ..config import PDFConversionSettings, ExcelSettings, get_excel_sheet_settings
from ..utils.logger import logger

# Excel constants from Object Model
xlTypePDF = 0
xlQualityStandard = 0
xlQualityMinimum = 1
xlLandscape = 2
xlPortrait = 1
xlPaperLetter = 1
xlPaperA4 = 9
xlPaperA3 = 8
xlPaperTabloid = 3
xlPaperLegal = 5
# Architecture sizes (approximate enum values, varies by driver but standard for PDF printers)
xlPaperC = 24  # 17x22 in
xlPaperD = 25  # 22x34 in
xlPaperE = 26  # 34x44 in

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
    
    def convert(
        self, 
        input_path: Path, 
        output_path: Optional[Path] = None, 
        settings: Optional[PDFConversionSettings] = None
    ) -> Path:
        """
        Convert an Excel document to PDF.
        
        Args:
            input_path: Path to the source Excel file.
            output_path: Optional path for the output PDF.
            settings: PDF conversion settings.
            
        Returns:
            Path to the generated PDF file.
        """
        input_file = input_path.resolve()
        if not input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")
            
        if output_path:
            out_file = output_path.resolve()
        else:
            out_file = input_file.with_suffix(".pdf")
            
        # Ensure output directory exists
        out_file.parent.mkdir(parents=True, exist_ok=True)
            
        settings = settings or PDFConversionSettings()
        excel_settings = settings.excel or ExcelSettings()
        
        logger.info(f"Converting '{input_file.name}' to PDF...")
        logger.debug(f"Settings: {settings}")
        logger.debug(f"Excel settings: {excel_settings}")

        # Ensure CoInitialize is called for this thread
        pythoncom.CoInitialize()
        
        try:
            with self._excel_application() as excel:
                workbook = None
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
                    
                    final_sheets_to_process = []
                    temp_sheets_to_delete = []
                    
                    # Apply page setup and process chunks
                    for sheet in sheets_to_export:
                        # Get sheet-specific settings
                        # Note: Arguments are (sheet_name, settings)
                        sheet_settings = get_excel_sheet_settings(sheet.Name, settings)
                        sheet_excel_settings = sheet_settings.excel or excel_settings
                        
                        logger.debug(f"Sheet '{sheet.Name}' settings: row_dimensions={sheet_excel_settings.row_dimensions}")
                        
                        # Check for Chunking
                        row_lim = sheet_excel_settings.row_dimensions
                        if row_lim and row_lim > 0:
                            # Chunking Mode
                            used_rows = sheet.UsedRange.Rows.Count
                            if used_rows == 0:
                                continue
                                
                            chunks = (used_rows + row_lim - 1) // row_lim
                            logger.info(f"Splitting sheet '{sheet.Name}' into {chunks} chunks (Rows: {row_lim})")
                            
                            for i in range(chunks):
                                start_row = i * row_lim + 1
                                end_row = min((i + 1) * row_lim, used_rows)
                                
                                # Copy sheet to end
                                # Use positional args for Copy: Copy(Before, After)
                                last_sheet = workbook.Sheets(workbook.Sheets.Count)
                                sheet.Copy(None, last_sheet)
                                new_sheet = workbook.Sheets(workbook.Sheets.Count)
                                
                                # Rename removed to prevent errors
                                
                                temp_sheets_to_delete.append(new_sheet)
                                
                                # Set Print Area
                                new_sheet.PageSetup.PrintArea = f"${start_row}:${end_row}"
                                
                                # Create chunk settings (Force fit to 1 page tall for this chunk)
                                # Cloning dataclass
                                chunk_settings = ExcelSettings(**dataclasses.asdict(sheet_excel_settings))
                                chunk_settings.row_dimensions = 0 # Force 1 page tall
                                
                                self._apply_page_setup(new_sheet, chunk_settings, input_file.name)
                                
                                # Header
                                if sheet_excel_settings.metadata_header:
                                    center_text = f"{start_row}-{end_row}"
                                    self._apply_metadata_header(new_sheet, sheet_excel_settings, input_file.name, center_text, left_text=sheet.Name)
                                    
                                final_sheets_to_process.append(new_sheet)
                        else:
                            # Standard Mode
                            self._apply_page_setup(sheet, sheet_excel_settings, input_file.name)
                            if sheet_excel_settings.metadata_header:
                                self._apply_metadata_header(sheet, sheet_excel_settings, input_file.name, center_text="")
                            final_sheets_to_process.append(sheet)
                    
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
            
            yield excel
        except Exception as e:
            logger.critical(f"Failed to initialize Microsoft Excel: {e}")
            raise
        finally:
            if excel:
                excel.Quit()

    def _set_optimal_printer(self, excel) -> None:
        """
        Attempt to set ActivePrinter to 'Microsoft Print to PDF' for better paper size support.
        Tries detailed port detection and fallback strategies.
        """
        target_name = "Microsoft Print to PDF"
        
        try:
            # Check if already active
            current = excel.ActivePrinter
            if target_name in current:
                return
        except:
            pass

        found_printer_info = None
        
        try:
             # Find correct printer string with port via win32print
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            printers = win32print.EnumPrinters(flags, None, 2)
            
            for p in printers:
                p_name = p.get('pPrinterName', '')
                if p_name == target_name:
                    found_printer_info = p
                    break
        except Exception as e:
            logger.debug(f"Printer enumeration failed: {e}")
            
        if not found_printer_info:
            logger.debug(f"Printer '{target_name}' not found in system enumeration.")
            return

        # Strategy 1: Use detected port
        p_name = found_printer_info.get('pPrinterName')
        p_port = found_printer_info.get('pPortName', '')
        
        candidates = []
        if p_port:
             candidates.append(f"{p_name} on {p_port}")
        
        # Strategy 2: Brute force Ne00-Ne09 (common for network/virtual printers in Excel)
        # This handles cases where pPortName is 'PORTPROMPT:' or similar non-connectable strings
        for i in range(10):
            candidates.append(f"{p_name} on Ne{i:02d}:")
            
        # Strategy 3: Naked name (rare but possible)
        candidates.append(p_name)
        
        success = False
        for candidate in candidates:
            try:
                excel.ActivePrinter = candidate
                logger.info(f"Successfully switched ActivePrinter to: '{candidate}'")
                success = True
                break
            except Exception:
                continue
                
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
        min_col_width_inches: float = 0.5
    ) -> Tuple[float, float]:
        """
        Calculate page width to ensure OCR-readable text (14pt minimum).
        
        Args:
            sheet: Excel Worksheet object
            min_col_width_inches: Minimum width per column for readable text
            
        Returns:
            Tuple of (page_width_inches, page_height_inches)
        """
        try:
            used_range = sheet.UsedRange
            col_count = used_range.Columns.Count
            
            # Calculate minimum page width for readable text
            min_page_width = col_count * min_col_width_inches
            
            # Clamp to valid range
            page_width = max(min_page_width, self.MIN_PAGE_WIDTH_INCHES)
            page_width = min(page_width, self.MAX_PAGE_WIDTH_INCHES)
            
            # Page height - use default or calculate based on row dimensions
            page_height = self.DEFAULT_PAGE_HEIGHT_INCHES
            
            logger.debug(
                f"Sheet '{sheet.Name}': {col_count} columns, "
                f"page size: {page_width:.1f}\" x {page_height:.1f}\""
            )
            
            return page_width, page_height
            
        except Exception as e:
            logger.warning(f"Could not calculate smart page size: {e}")
            return self.MIN_PAGE_WIDTH_INCHES, self.DEFAULT_PAGE_HEIGHT_INCHES

    def _apply_page_setup(
        self, 
        sheet, 
        excel_settings: ExcelSettings,
        filename: str
    ) -> None:
        """
        Apply page setup settings for OCR-optimized PDF output.
        
        Args:
            sheet: Excel Worksheet object
            excel_settings: Excel-specific settings
            filename: Original filename for header
        """
        try:
            page_setup = sheet.PageSetup
            
            # Calculate smart page size
            page_width, page_height = self._calculate_smart_page_size(
                sheet, 
                excel_settings.min_col_width_inches
            )
            
            # Set orientation
            if excel_settings.orientation.lower() == "portrait":
                page_setup.Orientation = xlPortrait
            else:
                page_setup.Orientation = xlLandscape
            
            # Smart Paper Size Selection
            is_landscape = (page_setup.Orientation == xlLandscape)
            
            # Define ladder of supported sizes (Enum, Landscape Width inches, Name)
            # Widths based on standard dimensions. Portrait widths would be Height.
            # We focus on Width for fitting columns.
            paper_ladder = [
                (xlPaperLetter, 11.0, "Letter"),
                (xlPaperA3, 16.54, "A3"),
                (xlPaperC, 22.0, "Arch C"),
                (xlPaperD, 34.0, "Arch D"),
                (xlPaperE, 44.0, "Arch E")
            ]
            
            if not is_landscape:
                # Approximate portrait widths
                paper_ladder = [
                    (xlPaperLetter, 8.5, "Letter"),
                    (xlPaperA3, 11.69, "A3"),
                    (xlPaperC, 17.0, "Arch C"),
                    (xlPaperD, 22.0, "Arch D"),
                    (xlPaperE, 34.0, "Arch E")
                ]

            selected_paper = xlPaperLetter
            available_width = 11.0
            selected_name = "Letter"
            
            # Find smallest paper that fits content OR largest available
            for (enum_val, width, name) in paper_ladder:
                selected_paper = enum_val
                available_width = width
                selected_name = name
                if page_width <= width:
                    break
            
            # Apply selection
            try:
                page_setup.PaperSize = selected_paper
                # Verify if applied
                current_size = page_setup.PaperSize
                if current_size != selected_paper:
                    logger.warning(f"Printer rejected paper size {selected_name} (Enum {selected_paper}). Got Enum {current_size}.")
                    # If rejected, we must revert available_width to estimated actual width
                    # Mapping generic enums back to width is hard, but usually it reverts to Letter/A4
                    if current_size == xlPaperLetter:
                         available_width = 11.0 if is_landscape else 8.5
                         selected_name = "Letter (Fallback)"
                    elif current_size == xlPaperA4:
                         available_width = 11.69 if is_landscape else 8.27
                         selected_name = "A4 (Fallback)"
                    elif current_size == xlPaperA3:
                         available_width = 16.54 if is_landscape else 11.69
                         selected_name = "A3 (Fallback)"
                    else:
                         # Unknown reset - assume worst case (Letter)
                         available_width = 11.0 if is_landscape else 8.5
                         selected_name = f"Unknown-{current_size} (Fallback)"
            except Exception as e:
                logger.warning(f"Failed to set paper size: {e}")
                available_width = 11.0 if is_landscape else 8.5
                selected_name = "Letter (Error)"

            if selected_name != "Letter":
                logger.info(f"Using {selected_name} paper for sheet '{sheet.Name}' (Content: {page_width:.1f}\" <= Paper: {available_width:.1f}\")")
            
            # Check for Microscopic Text
            if page_width > available_width:
                shrink_factor = available_width / page_width
                
                if shrink_factor < 0.5:
                    # Error condition as per user request
                    err_msg = (
                        f"Sheet '{sheet.Name}' is too wide ({page_width:.1f}\") for largest supported paper {selected_name} ({available_width:.1f}\"). "
                        f"Shrink factor {shrink_factor:.2f} results in microscopic text. "
                        f"Max supported behavior is Arch E paper."
                    )
                    logger.error(err_msg)
                    raise ValueError(err_msg)
                
                # Otherwise, fit to page (Standard Shrink)
                page_setup.Zoom = False
                page_setup.FitToPagesWide = 1
                self._apply_row_dimensions(sheet, page_setup, excel_settings)
            else:
                 # Fits natively
                 page_setup.Zoom = False
                 page_setup.FitToPagesWide = 1
                 self._apply_row_dimensions(sheet, page_setup, excel_settings)
            
            # Set margins (narrow for more content)
            margin_points = 36  # 0.5 inch
            
            # Increase Top Margin if metadata header is enabled to avoid overlap
            top_margin = 72 if excel_settings.metadata_header else 36 # 1.0 inch vs 0.5 inch
            
            page_setup.LeftMargin = margin_points
            page_setup.RightMargin = margin_points
            page_setup.TopMargin = top_margin
            page_setup.BottomMargin = margin_points
            
            logger.debug(f"Applied page setup for sheet '{sheet.Name}'")
            
        except ValueError:
            raise  # Re-raise explicit validation errors
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
            
            # Right header: Filename
            page_setup.RightHeader = filename
            
            # Right Footer: Page Number
            page_setup.RightFooter = "Page &P"
            
            # Clear other footers to avoid clutter
            page_setup.CenterFooter = ""
            page_setup.LeftFooter = ""
            
            logger.debug(f"Applied metadata header for sheet '{sheet.Name}' (Center: '{center_text}')")
            
        except Exception as e:
            logger.warning(f"Could not apply metadata header for '{sheet.Name}': {e}")
            
        except Exception as e:
            logger.warning(f"Could not apply metadata header for '{sheet.Name}': {e}")

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
