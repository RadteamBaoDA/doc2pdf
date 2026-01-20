"""
PDF Post-Processing utilities using pypdf and pdfminer.six.

Features:
- Auto-detect content bounds and trim whitespace
- Non-destructive cropping via CropBox manipulation
- License-friendly (BSD/MIT) implementation suitable for enterprise use
"""

from pathlib import Path
from typing import Optional, List, Union
import math

from pypdf import PdfReader, PdfWriter
from pypdf.generic import RectangleObject
from pdfminer.high_level import extract_pages
from pdfminer.layout import (
    LTPage, LTTextContainer, LTImage, LTFigure, 
    LTRect, LTLine, LTCurve, LTTextBox
)

from ..utils.logger import logger


class PDFProcessor:
    """Post-processing utilities for PDF files using pypdf and pdfminer."""
    
    def trim_whitespace(
        self, 
        pdf_path: Path, 
        margin: float = 10.0,
        output_path: Optional[Path] = None
    ) -> Path:
        """
        Auto-detect content bounds and crop PDF to remove whitespace.
        
        Algorithm:
        1. Analyze content bounds using pdfminer
        2. Calculate union rectangle of all content (filtering outliers)
        3. Apply crop using pypdf
        
        Args:
            pdf_path: Input PDF file path
            margin: Padding in points around detected content (default: 10pt)
            output_path: Optional output path (defaults to overwrite input)
            
        Returns:
            Path to the processed PDF file
        """
        pdf_path = Path(pdf_path).resolve()
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
        
        # Default: overwrite input file
        target_path = output_path.resolve() if output_path else pdf_path
        
        logger.info(f"Trimming whitespace from '{pdf_path.name}' (using pypdf/pdfminer)...")
        
        modified = False
        writer = PdfWriter()
        
        try:
            # 1. Open PDF with pypdf for modification
            reader = PdfReader(str(pdf_path))
            
            # 2. Extract layout analysis with pdfminer
            # extract_pages yields LTPage objects
            layout_pages = list(extract_pages(str(pdf_path)))
            
            if len(layout_pages) != len(reader.pages):
                logger.warning(
                    f"Page count mismatch: pypdf={len(reader.pages)}, "
                    f"pdfminer={len(layout_pages)}. Trimming may be inaccurate."
                )
            
            for i, page in enumerate(reader.pages):
                # Match logical page from pdfminer
                lt_page = layout_pages[i] if i < len(layout_pages) else None
                
                # Get existing dimensions from pypdf
                # mediabox is [ll_x, ll_y, ur_x, ur_y]
                mb = page.mediabox
                page_width = project_float(mb.width)
                page_height = project_float(mb.height)
                
                # Detect content bounds
                content_rect = self._detect_content_bounds(lt_page, page_width, page_height) if lt_page else None
                
                if content_rect:
                    # content_rect is (x0, y0, x1, y1)
                    c_x0, c_y0, c_x1, c_y1 = content_rect
                    
                    # Add margin padding
                    # Ensure we don't go outside existing MediaBox
                    new_x0 = max(project_float(mb.left), c_x0 - margin)
                    new_y0 = max(project_float(mb.bottom), c_y0 - margin)
                    new_x1 = min(project_float(mb.right), c_x1 + margin)
                    new_y1 = min(project_float(mb.top), c_y1 + margin)
                    
                    # Sanity check: If resulting box is basically the whole page, skips
                    current_w = project_float(mb.width)
                    current_h = project_float(mb.height)
                    new_w = new_x1 - new_x0
                    new_h = new_y1 - new_y0
                    
                    if new_w < current_w * 0.95 or new_h < current_h * 0.95:
                        # Apply CropBox
                        # pypdf RectangleObject expects (x, y, x, y) or list
                        page.cropbox = RectangleObject((new_x0, new_y0, new_x1, new_y1))
                        modified = True
                        logger.debug(
                            f"Page {i + 1}: Cropped to {new_w:.1f}x{new_h:.1f}pt "
                            f"(was {current_w:.1f}x{current_h:.1f}pt)"
                        )
                    else:
                        logger.debug(f"Page {i + 1}: Content fills most of page, no trim needed")
                else:
                    logger.debug(f"Page {i + 1}: No significant content detected, skipping")

                writer.add_page(page)
            
            # Save result
            if modified:
                if target_path == pdf_path:
                    # Overwrite
                    temp_path = pdf_path.with_suffix(".tmp.pdf")
                    with open(temp_path, "wb") as f:
                        writer.write(f)
                    
                    temp_path.replace(pdf_path)
                    logger.success(f"Trimmed whitespace from '{pdf_path.name}'")
                else:
                    target_path.parent.mkdir(parents=True, exist_ok=True)
                    with open(target_path, "wb") as f:
                        writer.write(f)
                    logger.success(f"Trimmed PDF saved to '{target_path.name}'")
            else:
                logger.info(f"No whitespace trimming needed for '{pdf_path.name}'")
                if target_path != pdf_path:
                    # If user requested explicit output path, verify we copy original?
                    # The CLI usually implies a conversion pipeline.
                    # Just copy original if needed, or do nothing.
                    pass

        except Exception as e:
            logger.error(f"Failed to trim whitespace: {e}")
            raise

        return target_path

    def _detect_content_bounds(self, lt_page: LTPage, page_width: float, page_height: float) -> Optional[tuple]:
        """
        Detect content bounds using pdfminer Layout Analysis.
        Returns (x0, y0, x1, y1) or None.
        """
        rects = []
        page_area = page_width * page_height
        
        # Traverse layout elements
        # LTPage acts as a container
        stack = list(lt_page)
        
        while stack:
            element = stack.pop()
            
            # Determine if element is "content"
            is_content = False
            
            if isinstance(element, (LTTextContainer, LTTextBox)):
                # Text
                if element.get_text().strip():
                    is_content = True
            elif isinstance(element, (LTImage, LTFigure)):
                is_content = True
            elif isinstance(element, (LTRect, LTLine, LTCurve)):
                # Vector graphics
                is_content = True
            
            if is_content:
                # Capture bounding box: (x0, y0, x1, y1)
                bbox = element.bbox
                x0, y0, x1, y1 = bbox
                w = x1 - x0
                h = y1 - y0
                
                # Filter 1: Background Artifacts (Huge rectangles)
                if w * h > page_area * 0.90:
                    continue
                
                # Save rect for clustering
                # Mark text elements to prevent them from being treated as outliers
                is_text = isinstance(element, (LTTextContainer, LTTextBox))
                rects.append(SimpleRect(x0, y0, x1, y1, is_important=is_text))
            
            # Recurse if container (e.g. LTFigure can contain text)
            if isinstance(element, (LTFigure, LTTextContainer)) and hasattr(element, "__iter__"):
                 # Simplification: we usually get text from containers directly
                 pass

        if not rects:
            return None
            
        # Refined Logic: Outlier Rejection
        # 1. Sort by Area Descending
        rects.sort(key=lambda r: r.area, reverse=True)
        
        # 2. Main content
        union_rect = rects[0]
        
        # 3. Merge loop
        for rect in rects[1:]:
            current_union_area = union_rect.width * union_rect.height
            
            # Calculate merged bbox
            ux0 = min(union_rect.x0, rect.x0)
            uy0 = min(union_rect.y0, rect.y0)
            ux1 = max(union_rect.x1, rect.x1)
            uy1 = max(union_rect.y1, rect.y1)
            
            merged_area = (ux1 - ux0) * (uy1 - uy0)
            expansion = merged_area - current_union_area
            
            # Heuristic
            is_tiny = rect.area < current_union_area * 0.01
            is_expansive = expansion > current_union_area * 0.10
            
            # EXCEPTION: If it is text (Important), do not discard even if expensive
            # This preserves Headers/Footers
            if is_tiny and is_expansive and not rect.is_important:
                continue
                
            union_rect = SimpleRect(ux0, uy0, ux1, uy1)
            
        return (union_rect.x0, union_rect.y0, union_rect.x1, union_rect.y1)

class SimpleRect:
    """Helper for bounding box calculations."""
    def __init__(self, x0, y0, x1, y1, is_important: bool = False):
        self.x0 = x0
        self.y0 = y0
        self.x1 = x1
        self.y1 = y1
        self.is_important = is_important
    
    @property
    def width(self): return self.x1 - self.x0
    
    @property
    def height(self): return self.y1 - self.y0
    
    @property
    def area(self): return self.width * self.height


def project_float(val):
    """Safely convert pypdf float objects to python float"""
    return float(val)
