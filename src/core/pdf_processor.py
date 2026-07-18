"""Conservative, pixel-based PDF whitespace trimming.

PDFium renders exactly what a viewer sees.  pypdf is then used only to clone the
document and update page boxes, so outlines, metadata, attachments and catalog
structures survive the operation.
"""

from __future__ import annotations

import math
import os
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Tuple

import pypdfium2 as pdfium
from PIL import Image, ImageChops
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, RectangleObject

from ..utils.logger import logger

Box = Tuple[float, float, float, float]


class PDFTrimError(RuntimeError):
    """Raised when trimming cannot be completed without risking the input."""


@dataclass(frozen=True)
class _RenderInfo:
    page_box: Box
    scale: float
    pixel_width: int
    pixel_height: int


def _box(box: object) -> Box:
    return tuple(float(value) for value in box)  # type: ignore[return-value]


def _intersection(first: Box, second: Box) -> Optional[Box]:
    result = (
        max(first[0], second[0]),
        max(first[1], second[1]),
        min(first[2], second[2]),
        min(first[3], second[3]),
    )
    return result if result[2] > result[0] and result[3] > result[1] else None


class PDFProcessor:
    """Trim PDFs using rendered pixels while preserving existing visibility."""

    def __init__(self, max_workers: Optional[int] = None):
        # Kept for API compatibility. PDFium is not thread-safe; pages are rendered
        # sequentially and process-level parallelism belongs to job orchestration.
        if max_workers is not None and max_workers <= 0:
            raise ValueError("max_workers must be > 0")
        self.max_workers = max_workers or 1

    def trim_whitespace(
        self,
        pdf_path: Path,
        margin: float = 10.0,
        output_path: Optional[Path] = None,
        *,
        box_mode: str = "physical",
        render_dpi: int = 72,
        max_render_pixels: int = 20_000_000,
        background_tolerance: int = 8,
        include_annotations: bool = True,
        allow_signature_invalidation: bool = False,
        password: Optional[str] = None,
    ) -> Path:
        """Trim visible whitespace and atomically commit the validated result.

        Existing CropBoxes are always respected. Blank pages are left unchanged.
        Explicit outputs are created even when no page needs trimming.
        """
        source = Path(pdf_path).resolve()
        target = Path(output_path).resolve() if output_path else source
        self._validate_options(
            source, margin, box_mode, render_dpi, max_render_pixels,
            background_tolerance,
        )
        target.parent.mkdir(parents=True, exist_ok=True)

        reader = PdfReader(str(source), password=password)
        if reader.is_encrypted and password is None:
            raise PDFTrimError("Encrypted PDF requires valid credentials")
        if self._has_signatures(reader) and not allow_signature_invalidation:
            raise PDFTrimError(
                "Signed PDF refused because changing page boxes invalidates signatures"
            )

        writer = PdfWriter()
        writer.clone_document_from_reader(reader)
        if reader.trailer.get("/ID") is not None:
            writer._ID = reader.trailer["/ID"]
        document = pdfium.PdfDocument(str(source), password=password)
        if include_annotations:
            try:
                document.init_forms()
            except Exception as exc:
                logger.debug(f"PDF has no initializable form environment: {exc}")
        if len(document) != len(writer.pages):
            raise PDFTrimError("PDFium and pypdf disagree on the page count")

        changed = False
        try:
            for index, output_page in enumerate(writer.pages):
                visible = _intersection(_box(output_page.mediabox), _box(output_page.cropbox))
                if visible is None:
                    raise PDFTrimError(f"Page {index + 1} has invalid MediaBox/CropBox")
                bounds = self._detect_page_bounds(
                    document[index], visible, render_dpi, max_render_pixels,
                    background_tolerance, include_annotations,
                )
                if bounds is None:
                    logger.debug(f"Page {index + 1}: blank; left unchanged")
                    continue
                proposed = self._margin_and_clamp(bounds, visible, margin)
                proposed = self._verify_and_expand(proposed, bounds, visible, margin)
                if proposed is None:
                    logger.warning(
                        f"Page {index + 1}: protected-margin verification failed; "
                        "page left unchanged"
                    )
                    continue
                if not self._saves_space(visible, proposed):
                    continue
                self._apply_boxes(output_page, proposed, box_mode)
                changed = True
        finally:
            document.close()

        self._atomic_write(writer, target)
        logger.info(
            f"Whitespace trim {'updated' if changed else 'validated'} '{target.name}'"
        )
        return target

    @staticmethod
    def _validate_options(
        source: Path, margin: float, box_mode: str, render_dpi: int,
        max_render_pixels: int, tolerance: int,
    ) -> None:
        if not source.is_file():
            raise FileNotFoundError(f"PDF file not found: {source}")
        if margin < 0:
            raise ValueError("margin must be >= 0")
        if box_mode not in {"physical", "cropbox"}:
            raise ValueError("box_mode must be physical or cropbox")
        if not 18 <= render_dpi <= 600:
            raise ValueError("render_dpi must be between 18 and 600")
        if max_render_pixels <= 0:
            raise ValueError("max_render_pixels must be > 0")
        if not 0 <= tolerance <= 255:
            raise ValueError("background_tolerance must be within 0..255")

    @staticmethod
    def _has_signatures(reader: PdfReader) -> bool:
        root = reader.trailer.get("/Root", {})
        if root.get("/Perms"):
            return True
        acroform = root.get("/AcroForm")
        if not acroform:
            return False
        for field_ref in acroform.get_object().get("/Fields", []):
            field = field_ref.get_object()
            value = field.get("/V")
            value = value.get_object() if hasattr(value, "get_object") else value
            if field.get("/FT") == "/Sig" or (
                hasattr(value, "get") and value.get("/Type") == "/Sig"
            ):
                return True
        return False

    def _detect_page_bounds(
        self, page, visible: Box, dpi: int, pixel_cap: int,
        tolerance: int, include_annotations: bool,
    ) -> Optional[Box]:
        width = visible[2] - visible[0]
        height = visible[3] - visible[1]
        scale = dpi / 72.0
        pixels = max(1, math.ceil(width * scale)) * max(1, math.ceil(height * scale))
        if pixels > pixel_cap:
            scale = max(18 / 72.0, scale * math.sqrt(pixel_cap / pixels))
        pixel_width = max(1, math.ceil(width * scale))
        pixel_height = max(1, math.ceil(height * scale))
        info = _RenderInfo(visible, scale, pixel_width, pixel_height)

        if pixel_width * pixel_height <= pixel_cap:
            image, converter = self._render_visible(page, visible, scale, include_annotations)
            bitmap_bounds = self._ink_bounds(image, tolerance)
            mapped = (
                self._device_to_page(converter, bitmap_bounds, info)
                if bitmap_bounds else None
            )
            image.close()
            return mapped

        # Even 18 DPI exceeds the cap: render independent capped tiles and union
        # their page-coordinate ink bounds instead of allocating a huge bitmap.
        union: Optional[Box] = None
        tile_edge = max(1, int(math.sqrt(pixel_cap)))
        for x0 in range(0, pixel_width, tile_edge):
            for y0 in range(0, pixel_height, tile_edge):
                x1 = min(pixel_width, x0 + tile_edge)
                y1 = min(pixel_height, y0 + tile_edge)
                tile_box = self._pixel_window_to_page((x0, y0, x1, y1), info)
                image, converter = self._render_visible(page, tile_box, scale, include_annotations)
                detected = self._ink_bounds(image, tolerance)
                tile_info = _RenderInfo(tile_box, scale, image.width, image.height)
                image.close()
                if detected:
                    page_bounds = self._device_to_page(converter, detected, tile_info)
                    union = page_bounds if union is None else (
                        min(union[0], page_bounds[0]), min(union[1], page_bounds[1]),
                        max(union[2], page_bounds[2]), max(union[3], page_bounds[3]),
                    )
        return union

    @staticmethod
    def _render_visible(page, visible: Box, scale: float, annotations: bool):
        try:
            effective = tuple(float(value) for value in page.get_bbox())
        except Exception:
            page_width, page_height = page.get_size()
            effective = (0.0, 0.0, float(page_width), float(page_height))
        crop = (
            max(0.0, visible[0] - effective[0]),
            max(0.0, visible[1] - effective[1]),
            max(0.0, effective[2] - visible[2]),
            max(0.0, effective[3] - visible[3]),
        )
        bitmap = page.render(
            scale=scale,
            crop=crop,
            rotation=0,
            draw_annots=annotations,
            fill_color=(255, 255, 255, 255),
        )
        image = bitmap.to_pil().convert("RGB")
        try:
            converter = bitmap.get_posconv(page)
        except Exception:
            converter = None
        bitmap.close()
        return image, converter

    @staticmethod
    def _ink_bounds(image: Image.Image, tolerance: int) -> Optional[Tuple[int, int, int, int]]:
        white = Image.new("RGB", image.size, "white")
        difference = ImageChops.difference(image, white)
        # Max-channel distance > tolerance. Point() produces a compact mask.
        mask = difference.convert("RGB").point(
            lambda value: 255 if value > tolerance else 0
        ).convert("L")
        bounds = mask.getbbox()
        mask.close()
        white.close()
        if bounds is None:
            return None
        # One-pixel antialias guard.
        return (
            max(0, bounds[0] - 1), max(0, bounds[1] - 1),
            min(image.width, bounds[2] + 1), min(image.height, bounds[3] + 1),
        )

    @staticmethod
    def _pixels_to_page(bounds: Tuple[int, int, int, int], info: _RenderInfo) -> Box:
        x0, y0, x1, y1 = bounds
        page = info.page_box
        return (
            page[0] + x0 / info.scale,
            page[3] - y1 / info.scale,
            page[0] + x1 / info.scale,
            page[3] - y0 / info.scale,
        )

    @staticmethod
    def _device_to_page(converter, bounds, fallback: _RenderInfo) -> Box:
        """Use PDFium's device-to-page transform, including rotation and origins."""
        try:
            if converter is None:
                raise RuntimeError("PdfBitmap.get_posconv() is unavailable")
            points = [
                converter.to_page(bounds[0], bounds[1]),
                converter.to_page(bounds[2], bounds[1]),
                converter.to_page(bounds[0], bounds[3]),
                converter.to_page(bounds[2], bounds[3]),
            ]
            xs = [float(point[0]) for point in points]
            ys = [float(point[1]) for point in points]
            transformed = (min(xs), min(ys), max(xs), max(ys))
            return _intersection(transformed, fallback.page_box) or fallback.page_box
        except Exception as exc:
            logger.debug(f"PDFium position conversion unavailable; using fallback: {exc}")
            return PDFProcessor._pixels_to_page(bounds, fallback)

    @staticmethod
    def _pixel_window_to_page(bounds: Tuple[int, int, int, int], info: _RenderInfo) -> Box:
        return PDFProcessor._pixels_to_page(bounds, info)

    @staticmethod
    def _margin_and_clamp(bounds: Box, visible: Box, margin: float) -> Box:
        return (
            max(visible[0], bounds[0] - margin),
            max(visible[1], bounds[1] - margin),
            min(visible[2], bounds[2] + margin),
            min(visible[3], bounds[3] + margin),
        )

    @staticmethod
    def _verify_and_expand(
        proposed: Box, ink: Box, visible: Box, margin: float,
    ) -> Optional[Box]:
        # This is the page-coordinate equivalent of re-render edge verification:
        # every protected edge must contain the requested margin unless clamped by
        # the pre-existing visible box. Expand once by one point if rounding ate it.
        expanded = list(proposed)
        required = (
            ink[0] - margin, ink[1] - margin, ink[2] + margin, ink[3] + margin,
        )
        if expanded[0] > max(visible[0], required[0]):
            expanded[0] = max(visible[0], required[0] - 1)
        if expanded[1] > max(visible[1], required[1]):
            expanded[1] = max(visible[1], required[1] - 1)
        if expanded[2] < min(visible[2], required[2]):
            expanded[2] = min(visible[2], required[2] + 1)
        if expanded[3] < min(visible[3], required[3]):
            expanded[3] = min(visible[3], required[3] + 1)
        result = tuple(expanded)  # type: ignore[assignment]
        if result[0] > ink[0] or result[1] > ink[1] or result[2] < ink[2] or result[3] < ink[3]:
            return None
        return result  # type: ignore[return-value]

    @staticmethod
    def _saves_space(old: Box, new: Box) -> bool:
        return any((new[0] - old[0], new[1] - old[1], old[2] - new[2], old[3] - new[3])) and any(
            saved > 2.0
            for saved in (new[0] - old[0], new[1] - old[1], old[2] - new[2], old[3] - new[3])
        )

    @staticmethod
    def _apply_boxes(page, tight: Box, mode: str) -> None:
        rectangle = RectangleObject(tight)
        page.cropbox = rectangle
        if mode == "physical":
            page.mediabox = RectangleObject(tight)
            for name in ("/TrimBox", "/BleedBox", "/ArtBox"):
                existing = page.get(name)
                if existing is None:
                    continue
                clipped = _intersection(_box(existing), tight)
                if clipped:
                    page[NameObject(name)] = RectangleObject(clipped)

    @staticmethod
    def _atomic_write(writer: PdfWriter, target: Path) -> None:
        fd, temporary_name = tempfile.mkstemp(
            prefix=f".{target.name}.", suffix=".tmp.pdf", dir=str(target.parent)
        )
        temporary = Path(temporary_name)
        try:
            with os.fdopen(fd, "wb") as stream:
                writer.write(stream)
                stream.flush()
                os.fsync(stream.fileno())
            validation = PdfReader(str(temporary))
            if len(validation.pages) != len(writer.pages) or temporary.stat().st_size == 0:
                raise PDFTrimError("Temporary output failed PDF validation")
            os.replace(temporary, target)
        finally:
            temporary.unlink(missing_ok=True)
