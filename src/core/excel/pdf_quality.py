"""Conservative, non-rasterizing postflight for Excel-produced PDFs."""

from __future__ import annotations

import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Optional, Tuple

from pypdf import PdfReader
from pypdf.generic import ContentStream

from .models import PdfPostflightResult


@dataclass(frozen=True)
class PdfQualityExpectation:
    expected_pages: Optional[int] = None
    sentinels: Tuple[str, ...] = ()
    allowed_boxes_points: Tuple[Tuple[float, float], ...] = ()
    min_font_pt: float = 0.0
    min_image_dpi: float = 0.0
    max_dimension_in: float = math.inf
    max_area_in2: float = math.inf
    require_searchable_text: bool = True


def _resources(page: Any) -> Mapping[str, Any]:
    """Document this Excel pipeline operation and its side effects."""
    resources = page.get("/Resources", {})
    return resources.get_object() if hasattr(resources, "get_object") else resources


def _xobjects(page: Any) -> Mapping[str, Any]:
    """Document this Excel pipeline operation and its side effects."""
    objects = _resources(page).get("/XObject", {})
    return objects.get_object() if hasattr(objects, "get_object") else objects


class PdfQualityPostflight:
    def validate(
        self, path: Path, expectation: PdfQualityExpectation,
    ) -> PdfPostflightResult:
        """Document this Excel pipeline operation and its side effects."""
        failures: list[str] = []
        warnings: list[str] = []
        page_evidence: list[Mapping[str, Any]] = []
        try:
            reader = PdfReader(str(path))
        except Exception as exc:
            return PdfPostflightResult(
                False, 0, {"readable": False}, (f"unreadable PDF: {exc}",)
            )
        pages = list(reader.pages)
        # Page-level checks stay independent so failures identify the affected export unit.
        if not pages:
            failures.append("PDF has no pages")
        if expectation.expected_pages is not None and len(pages) != expectation.expected_pages:
            failures.append(
                f"page count {len(pages)} does not equal expected {expectation.expected_pages}"
            )
        all_text: list[str] = []
        all_fonts: list[float] = []
        all_image_dpi: list[float] = []
        any_searchable = False
        any_nonblank = False
        for page_number, page in enumerate(pages, 1):
            # Inspect text and graphics together to detect blank or rasterized pages.
            width = abs(float(page.mediabox.right) - float(page.mediabox.left))
            height = abs(float(page.mediabox.top) - float(page.mediabox.bottom))
            rotation = int(page.get("/Rotate", 0) or 0) % 360
            text_fragments: list[str] = []
            font_sizes: list[float] = []

            clipped_text = [False]
            page_ref = page

            def visitor(
                text: str, cm, tm, _font, size: float,
                _texts=text_fragments, _fonts=font_sizes,
                _page=page_ref, _clipped=clipped_text,
            ) -> None:
                """Document this Excel pipeline operation and its side effects."""
                if text:
                    _texts.append(text)
                    try:
                        vertical_scale = math.hypot(float(tm[2]), float(tm[3])) or 1.0
                        _fonts.append(abs(float(size)) * vertical_scale)
                    except (TypeError, ValueError, IndexError):
                        pass
                    try:
                        x = float(tm[4]) * float(cm[0]) + float(tm[5]) * float(cm[2]) + float(cm[4])
                        y = float(tm[4]) * float(cm[1]) + float(tm[5]) * float(cm[3]) + float(cm[5])
                        if x < float(_page.cropbox.left) - 2 or x > float(_page.cropbox.right) + 2:
                            _clipped[0] = True
                        if y < float(_page.cropbox.bottom) - 2 or y > float(_page.cropbox.top) + 2:
                            _clipped[0] = True
                    except (TypeError, ValueError, IndexError):
                        pass

            try:
                extracted = page.extract_text(visitor_text=visitor) or ""
            except Exception as exc:
                extracted = ""
                warnings.append(f"page {page_number} text extraction failed: {exc}")
            if extracted.strip():
                any_searchable = True
            all_text.append(extracted)
            all_fonts.extend(value for value in font_sizes if value > 0)
            image_count = 0
            full_page_rasters = 0
            image_dpi: list[float] = []
            try:
                for reference in _xobjects(page).values():
                    obj = reference.get_object()
                    if obj.get("/Subtype") == "/Image":
                        image_count += 1
                        pixel_width = float(obj.get("/Width", 0) or 0)
                        pixel_height = float(obj.get("/Height", 0) or 0)
                        # Without interpreting every nested Form CTM, conservatively
                        # flag only a bitmap whose pixel aspect and page coverage are
                        # characteristic of a page raster.
                        if pixel_width > 0 and pixel_height > 0:
                            page_aspect = width / height if height else 0
                            image_aspect = pixel_width / pixel_height
                            if page_aspect and abs(page_aspect - image_aspect) / page_aspect < 0.03:
                                full_page_rasters += 1
                image_dpi = self._image_placement_dpi(page)
                all_image_dpi.extend(image_dpi)
            except Exception as exc:
                warnings.append(f"page {page_number} image inventory failed: {exc}")
            content = page.get_contents()
            try:
                has_content_stream = bool(
                    content is not None and content.get_data().strip()
                )
            except Exception:
                has_content_stream = content is not None
            nonblank = bool(extracted.strip() or image_count or has_content_stream)
            any_nonblank = any_nonblank or nonblank
            if not nonblank:
                failures.append(f"page {page_number} is unexpectedly blank")
            if max(width, height) / 72.0 > expectation.max_dimension_in + 1e-6:
                failures.append(f"page {page_number} exceeds maximum dimension")
            if (width * height) / (72.0 * 72.0) > expectation.max_area_in2 + 1e-6:
                failures.append(f"page {page_number} exceeds maximum area")
            if expectation.allowed_boxes_points and not any(
                math.isclose(width, box[0], abs_tol=1.0)
                and math.isclose(height, box[1], abs_tol=1.0)
                or math.isclose(width, box[1], abs_tol=1.0)
                and math.isclose(height, box[0], abs_tol=1.0)
                for box in expectation.allowed_boxes_points
            ):
                failures.append(f"page {page_number} has an unexpected MediaBox")
            if full_page_rasters and not extracted.strip():
                failures.append(f"page {page_number} appears fully rasterized")
            if clipped_text[0]:
                failures.append(f"page {page_number} contains text outside its CropBox")
            if expectation.min_image_dpi and image_count:
                if not image_dpi:
                    failures.append(f"page {page_number} image DPI could not be verified")
                elif min(image_dpi) + 1e-6 < expectation.min_image_dpi:
                    failures.append(
                        f"page {page_number} image DPI {min(image_dpi):.1f} is below "
                        f"{expectation.min_image_dpi:.1f}"
                    )
            page_evidence.append({
                "page": page_number,
                "width_points": width,
                "height_points": height,
                "rotation": rotation,
                "searchable_characters": len(extracted.strip()),
                "font_sizes_pt": sorted(font_sizes),
                "image_count": image_count,
                "image_dpi": sorted(image_dpi),
                "full_page_raster_candidates": full_page_rasters,
                "nonblank": nonblank,
            })
        joined = "\n".join(all_text)
        for sentinel in expectation.sentinels:
            # Boundary sentinels prove that edge content survived PDF export.
            if sentinel and sentinel not in joined:
                failures.append(f"missing sentinel text: {sentinel!r}")
        if expectation.require_searchable_text and not any_searchable:
            failures.append("PDF contains no searchable text")
        if expectation.min_font_pt:
            if not all_fonts:
                failures.append("effective font size could not be verified")
            else:
                minimum = min(all_fonts)
                percentile = sorted(all_fonts)[max(0, int(len(all_fonts) * 0.1) - 1)]
                if minimum + 1e-6 < expectation.min_font_pt:
                    failures.append(
                        f"minimum effective font {minimum:.2f}pt is below "
                        f"{expectation.min_font_pt:.2f}pt"
                    )
                if percentile + 1e-6 < expectation.min_font_pt:
                    failures.append("10th-percentile effective font is below the quality floor")
        checks = {
            "readable": True,
            "page_count": expectation.expected_pages is None or len(pages) == expectation.expected_pages,
            "nonblank": any_nonblank and not any("unexpectedly blank" in item for item in failures),
            "searchable_text": (not expectation.require_searchable_text) or any_searchable,
            "sentinels": not any("missing sentinel" in item for item in failures),
            "page_geometry": not any("maximum" in item or "MediaBox" in item for item in failures),
            "fonts": not any("font" in item for item in failures),
            "image_dpi": not any("image DPI" in item for item in failures),
            "clipping": not any("outside its CropBox" in item for item in failures),
            "rasterization": not any("rasterized" in item for item in failures),
        }
        return PdfPostflightResult(
            not failures, len(pages), checks, tuple(failures), tuple(warnings),
            tuple(page_evidence),
        )

    @staticmethod
    def _image_placement_dpi(page: Any) -> list[float]:
        """Inspect top-level image Do operations and their active CTM."""
        xobjects = _xobjects(page)
        if not xobjects or page.get_contents() is None:
            return []
        stream = ContentStream(page.get_contents(), page.pdf)
        current = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
        stack: list[Tuple[float, float, float, float, float, float]] = []
        result: list[float] = []
        for operands, operator in stream.operations:
            if operator == b"q":
                stack.append(current)
            elif operator == b"Q":
                current = stack.pop() if stack else (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
            elif operator == b"cm" and len(operands) == 6:
                matrix = tuple(float(value) for value in operands)
                a, b, c, d, e, f = matrix
                a0, b0, c0, d0, e0, f0 = current
                current = (
                    a * a0 + b * c0,
                    a * b0 + b * d0,
                    c * a0 + d * c0,
                    c * b0 + d * d0,
                    e * a0 + f * c0 + e0,
                    e * b0 + f * d0 + f0,
                )
            elif operator == b"Do" and operands:
                reference = xobjects.get(operands[0])
                if reference is None:
                    continue
                obj = reference.get_object()
                if obj.get("/Subtype") != "/Image":
                    continue
                display_width = math.hypot(current[0], current[1]) / 72.0
                display_height = math.hypot(current[2], current[3]) / 72.0
                pixel_width = float(obj.get("/Width", 0) or 0)
                pixel_height = float(obj.get("/Height", 0) or 0)
                if display_width > 0 and display_height > 0:
                    result.append(min(
                        pixel_width / display_width,
                        pixel_height / display_height,
                    ))
        return result
