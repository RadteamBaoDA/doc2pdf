from pathlib import Path

import pytest
from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, NameObject, RectangleObject

from src.core.pdf_processor import PDFProcessor, PDFTrimError


def _write_fixture(path: Path, *, cropbox=None, rotation=0, blank=False) -> None:
    writer = PdfWriter()
    page = writer.add_blank_page(width=600, height=800)
    if cropbox:
        page.cropbox = RectangleObject(cropbox)
    if rotation:
        page.rotate(rotation)
    if not blank:
        stream = DecodedStreamObject()
        stream.set_data(b"q 0 0 0 rg 100 200 50 30 re f Q")
        page[NameObject("/Contents")] = stream
    writer.add_metadata({"/Title": "trim fixture"})
    writer.add_outline_item("Content", 0)
    writer.add_attachment("note.txt", b"retained")
    with path.open("wb") as output:
        writer.write(output)


def test_trim_physical_boxes_and_preserves_metadata(tmp_path):
    source = tmp_path / "source.pdf"
    target = tmp_path / "target.pdf"
    _write_fixture(source)

    PDFProcessor().trim_whitespace(source, margin=10, output_path=target)

    result = PdfReader(str(target))
    assert result.metadata.title == "trim fixture"
    assert result.outline
    assert result.attachments["note.txt"] == [b"retained"]
    page = result.pages[0]
    assert float(page.mediabox.width) < 100
    assert float(page.mediabox.height) < 100
    assert tuple(page.cropbox) == tuple(page.mediabox)


def test_cropbox_mode_never_reveals_hidden_content(tmp_path):
    source = tmp_path / "cropped.pdf"
    _write_fixture(source, cropbox=(80, 180, 180, 260))

    PDFProcessor().trim_whitespace(source, margin=5, box_mode="cropbox")

    page = PdfReader(str(source)).pages[0]
    assert tuple(page.mediabox) == (0, 0, 600, 800)
    assert float(page.cropbox.left) >= 80
    assert float(page.cropbox.bottom) >= 180
    assert float(page.cropbox.right) <= 180
    assert float(page.cropbox.top) <= 260


@pytest.mark.parametrize("rotation", [0, 90, 180, 270])
def test_rotated_pages_remain_readable(tmp_path, rotation):
    source = tmp_path / f"rotation-{rotation}.pdf"
    _write_fixture(source, rotation=rotation)
    PDFProcessor().trim_whitespace(source, margin=8)
    assert len(PdfReader(str(source)).pages) == 1


def test_blank_explicit_output_is_still_written(tmp_path):
    source = tmp_path / "blank.pdf"
    target = tmp_path / "blank-output.pdf"
    _write_fixture(source, blank=True)
    PDFProcessor().trim_whitespace(source, output_path=target)
    assert target.is_file()
    assert len(PdfReader(str(target)).pages) == 1


def test_signed_pdf_is_refused(tmp_path):
    source = tmp_path / "signed.pdf"
    writer = PdfWriter()
    writer.add_blank_page(width=100, height=100)
    writer._root_object[NameObject("/Perms")] = writer._root_object.__class__()
    with source.open("wb") as output:
        writer.write(output)
    with pytest.raises(PDFTrimError, match="Signed PDF"):
        PDFProcessor().trim_whitespace(source)


def test_invalid_options_are_fatal(tmp_path):
    source = tmp_path / "blank.pdf"
    _write_fixture(source, blank=True)
    with pytest.raises(ValueError):
        PDFProcessor().trim_whitespace(source, render_dpi=17)
