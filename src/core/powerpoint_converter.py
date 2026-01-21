"""
PowerPoint to PDF Converter using pywin32 COM.
"""
import sys
from pathlib import Path
from typing import Optional
import win32com.client
import pythoncom
from contextlib import contextmanager

from .base import Converter
from ..config import PDFConversionSettings, PowerPointSettings
from ..utils.logger import logger
from ..utils.process_manager import ProcessRegistry

# Constants from PowerPoint Object Model
ppFixedFormatTypePDF = 2
ppFixedFormatIntentPrint = 2
ppFixedFormatIntentScreen = 1

# Print Range Type
ppPrintAll = 1
ppPrintSelection = 2
ppPrintSlideRange = 4

# Print Color Type
ppPrintColor = 1
ppPrintBlackAndWhite = 2
ppPrintPureBlackAndWhite = 3

# Save/Close constants
ppSaveChanges = 1
ppDoNotSaveChanges = 2


class PowerPointConverter(Converter):
    """
    Converter for PowerPoint documents (.ppt, .pptx) to PDF.
    """
    
    def convert(
        self, 
        input_path: Path, 
        output_path: Optional[Path] = None, 
        settings: Optional[PDFConversionSettings] = None
    ) -> Path:
        """
        Convert a PowerPoint document to PDF.
        
        Args:
            input_path: Path to the source PowerPoint file.
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
        
        logger.info(f"Converting '{input_file.name}' to PDF...")
        logger.debug(f"Settings: {settings}")

        # Ensure CoInitialize is called for this thread
        pythoncom.CoInitialize()
        
        try:
            with self._powerpoint_application() as ppt:
                presentation = None
                try:
                    # Open Presentation (ReadOnly for safety)
                    presentation = ppt.Presentations.Open(
                        str(input_file), 
                        ReadOnly=True, 
                        Untitled=False, 
                        WithWindow=False
                    )
                    
                    # Prepare Export Arguments
                    export_args = self._map_settings(settings, str(out_file))
                    
                    # Export to PDF
                    presentation.ExportAsFixedFormat(**export_args)
                    
                    logger.success(f"Successfully converted: {out_file}")
                    
                except Exception as e:
                    logger.error(f"Failed to convert {input_file.name}: {e}")
                    raise
                finally:
                    if presentation:
                        presentation.Close()
        finally:
            pythoncom.CoUninitialize()
            
        return out_file

    @contextmanager
    def _powerpoint_application(self):
        """
        Context manager for PowerPoint COM application lifecycle.
        """
        ppt = None
        try:
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            # Note: PowerPoint doesn't have a Visible property like Word
            # but setting DisplayAlerts to false can help
            ppt.DisplayAlerts = False
            ProcessRegistry.register(ppt)
            yield ppt
        except Exception as e:
            logger.critical(f"Failed to initialize Microsoft PowerPoint: {e}")
            raise
        finally:
            if ppt:
                ProcessRegistry.unregister(ppt)
                ppt.Quit()

    def _map_settings(self, settings: PDFConversionSettings, output_path: str) -> dict:
        """
        Map PDFConversionSettings to ExportAsFixedFormat arguments.
        """
        # Get PowerPoint-specific settings
        ppt_settings = settings.powerpoint or PowerPointSettings()
        
        # Print Range Type
        range_type = ppPrintAll
        from_slide = 1
        to_slide = -1  # Will be set to presentation length by COM
        
        if settings.scope == "range" and ppt_settings.slide_from:
            range_type = ppPrintSlideRange
            from_slide = ppt_settings.slide_from
            to_slide = ppt_settings.slide_to or from_slide
        
        # Color Mode
        color_mode = ppPrintColor
        if ppt_settings.color_mode == "grayscale":
            color_mode = ppPrintBlackAndWhite
        elif ppt_settings.color_mode == "bw":
            color_mode = ppPrintPureBlackAndWhite
        
        # Intent (quality)
        intent = ppFixedFormatIntentPrint
        if settings.optimization.image_quality == "low":
            intent = ppFixedFormatIntentScreen
        
        export_args = {
            "Path": output_path,
            "FixedFormatType": ppFixedFormatTypePDF,
            "Intent": intent,
            "PrintRange": None,  # Use RangeType instead
            "RangeType": range_type,
            "FrameSlides": False,
            "HandoutOrder": 1,  # ppPrintHandoutVerticalFirst
            "OutputType": 1,  # ppPrintOutputSlides
            "IncludeDocProps": settings.metadata.include_properties,
            "KeepIRMSettings": True,
            "DocStructureTags": settings.metadata.include_tags,
            "BitmapMissingFonts": settings.optimization.bitmap_text,
            "UseISO19005_1": (settings.compliance == "pdfa"),
        }
        
        # Add slide range if specified
        if range_type == ppPrintSlideRange:
            export_args["SlideShowName"] = ""
            # Note: For range, we need to create a PrintRange object
            # This is handled differently - we'll use From/To parameters
            logger.debug(f"Exporting slides {from_slide} to {to_slide}")
        
        return export_args
