import sys
from pathlib import Path
from typing import Optional
import win32com.client
import pythoncom
from contextlib import contextmanager

from .base import Converter
from ..config import PDFConversionSettings, LayoutSettings
from ..utils.logger import logger

# Constants from Word Object Model
wdExportFormatPDF = 17
wdExportOptimizeForPrint = 0
wdExportOptimizeForOnScreen = 1
wdExportAllDocument = 0
wdExportSelection = 1
wdExportFromTo = 3
wdExportCreateNoBookmarks = 0
wdExportCreateHeadingBookmarks = 1
wdExportCreateWordBookmarks = 2
wdOrientPortrait = 0
wdOrientLandscape = 1
wdDoNotSaveChanges = 0

class WordConverter(Converter):
    def convert(self, input_path: Path, output_path: Optional[Path] = None, settings: Optional[PDFConversionSettings] = None) -> Path:
        input_file = input_path.resolve()
        if not input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")
            
        if output_path:
            out_file = output_path.resolve()
        else:
            out_file = input_file.with_suffix(".pdf")
            
        settings = settings or PDFConversionSettings()
        
        logger.info(f"Converting '{input_file.name}' to PDF...")
        logger.debug(f"Settings: {settings}")

        # Ensure CoInitialize is called for this thread
        pythoncom.CoInitialize()
        
        try:
            with self._word_application() as word:
                doc = None
                try:
                    # Open Document (ReadOnly to be safe)
                    doc = word.Documents.Open(str(input_file), ReadOnly=True, Visible=False)
                    
                    # Apply temporary layout settings if needed (careful with ReadOnly, might need to change ReadOnly=False if we want to change layout before print?)
                    # Actually, changing layout on a ReadOnly doc changes the view in memory, but we can't save. 
                    # ExportAsFixedFormat should respect current view settings.
                    self._apply_page_setup(doc, settings.layout)
                    
                    # Prepare Export Arguments
                    export_args = self._map_settings(settings, str(out_file))
                    
                    # Export
                    doc.ExportAsFixedFormat(**export_args)
                    
                    logger.success(f"Successfully converted: {out_file}")
                    
                except Exception as e:
                    logger.error(f"Failed to convert {input_file.name}: {e}")
                    raise
                finally:
                    if doc:
                        doc.Close(SaveChanges=wdDoNotSaveChanges)
        finally:
            pythoncom.CoUninitialize()
            
        return out_file

    @contextmanager
    def _word_application(self):
        word = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            yield word
        except Exception as e:
            logger.critical(f"Failed to initialize Microsoft Word: {e}")
            raise
        finally:
            if word:
                # We generally don't want to kill Word if it was already open effectively, 
                # but Dispatch logic usually attaches. 
                # Ideally we check if we created it or not, but strictly quitting is safer for batch processing CLI.
                # However, Dispatch creates a new connection. `DispatchEx` enforces new instance. 
                # Using standard Dispatch.
                word.Quit()

    def _apply_page_setup(self, doc, layout: LayoutSettings):
        """
        Apply page setup settings. 
        Note: Modifying page setup usually requires the document to be editable.
        If ReadOnly=True, this might raise error or fail silently. 
        """
        try:
            # Orientation
            if layout.orientation.lower() == "landscape":
                doc.PageSetup.Orientation = wdOrientLandscape
            else:
                doc.PageSetup.Orientation = wdOrientPortrait
                
            # Margins logic is complex to map generic "normal/narrow" to points.
            # Using simple heuristic or skipping for now if too complex for generic COM without specific points.
            # But the user asked for it. 
            # 1 inch = 72 points.
            
            if layout.margins == "narrow":
                 margin = 36 # 0.5 inch
                 doc.PageSetup.LeftMargin = margin
                 doc.PageSetup.RightMargin = margin
                 doc.PageSetup.TopMargin = margin
                 doc.PageSetup.BottomMargin = margin
            
            # Pages per sheet is usually a print driver setting, not directly settable easily in ExportAsFixedFormat 
            # unless using PrintOut. ExportAsFixedFormat is standard PDF export which is 1:1.
            # If user heavily requested "Pages per sheet", we might need PrintOut method, but that requires selecting a printer.
            # ExportAsFixedFormat is more reliable for "Save as PDF". 
            # We will log a warning if Pages per sheet > 1 is requested as it's not supported in standard ExportAsFixedFormat.
            if layout.pages_per_sheet > 1:
                logger.warning("Pages per sheet setting is not supported in direct PDF export mode. Ignoring.")

        except Exception as e:
            logger.warning(f"Could not apply some page setup settings: {e}")

    def _map_settings(self, settings: PDFConversionSettings, output_path: str) -> dict:
        """
        Map settings to ExportAsFixedFormat arguments.
        """
        # OptimizeFor
        optimize_for = wdExportOptimizeForOnScreen if settings.optimization.image_quality == "low" else wdExportOptimizeForPrint
        
        # Range
        export_range = wdExportAllDocument
        if settings.scope == "selection":
            export_range = wdExportSelection
        # 'range' requires From/To, which is not fully captured in our simple config yet unless we add fields.
        
        # Bookmarks
        create_bookmarks = wdExportCreateNoBookmarks
        if settings.bookmarks == "headings":
            create_bookmarks = wdExportCreateHeadingBookmarks
        elif settings.bookmarks == "bookmarks":
            create_bookmarks = wdExportCreateWordBookmarks
            
        return {
            "OutputFileName": output_path,
            "ExportFormat": wdExportFormatPDF,
            "OpenAfterExport": False,
            "OptimizeFor": optimize_for,
            "Range": export_range,
            "IncludeDocProps": settings.metadata.include_properties,
            "KeepIRM": True, 
            "CreateBookmarks": create_bookmarks,
            "DocStructureTags": settings.metadata.include_tags, # Accessibility tags
            "BitmapMissingFonts": settings.optimization.bitmap_text,
            "UseISO19005_1": (settings.compliance == "pdfa")
        }
