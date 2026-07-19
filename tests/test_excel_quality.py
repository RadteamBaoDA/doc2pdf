from types import SimpleNamespace
from unittest.mock import patch
from zipfile import ZipFile

import pytest
from pypdf import PdfWriter
from pypdf.generic import DecodedStreamObject, NameObject

from src.config import ExcelSettings, PDFConversionSettings, get_excel_sheet_settings
from src.core.excel.chunking import SafeChunkPlanner
from src.core.excel.extensions import (
    UnsupportedExcelExtensionError,
    require_supported_extensions,
)
from src.core.excel.layout import AuthoredLayoutInspector, persisted_print_sheets
from src.core.excel.models import (
    ConversionManifest,
    LayoutConstraints,
    LayoutDecision,
    PrintableObject,
    QualityLayoutCandidate,
    ResolvedRegion,
    stable_id,
)
from src.core.excel.planner import choose_candidate
from src.core.excel.pagination import ExcelPaginationProbe
from src.core.excel.printer import PrinterCapabilityProvider
from src.core.excel.pdf_quality import PdfQualityExpectation, PdfQualityPostflight


def test_strict_and_balanced_profiles_expand_before_explicit_values():
    strict = PDFConversionSettings.from_dict({
        "excel": {"quality_profile": "strict", "min_shrink_factor": 0.95}
    }).excel
    balanced = PDFConversionSettings.from_dict({
        "excel": {"quality_profile": "balanced"}
    }).excel
    assert strict is not None and balanced is not None
    assert strict.page_size_scope == "sheet"
    assert strict.postflight_policy == "strict"
    assert strict.min_shrink_factor == pytest.approx(0.95)
    assert balanced.print_area_policy == "expand_visible_objects"
    assert balanced.trim_policy == "cropbox"


def test_direct_strict_profile_uses_strict_defaults():
    settings = ExcelSettings(quality_profile="strict")
    assert settings.layout_policy == "preserve_authored"
    assert settings.printer_policy == "required"
    assert settings.postflight_policy == "strict"


def test_sheet_rule_reexpands_raw_profile_without_default_leakage():
    base = PDFConversionSettings.from_dict({
        "excel": {"quality_profile": "legacy", "orientation": "portrait"}
    })
    config = {"pdf_settings": {"excel": [{
        "sheet_name": "Data", "priority": 10,
        "settings": {"excel": {"quality_profile": "strict"}},
    }]}}
    with patch("src.config.load_config", return_value=config):
        result = get_excel_sheet_settings("Data", base_settings=base)
    assert result.excel is not None
    assert result.excel.quality_profile == "strict"
    assert result.excel.page_size_scope == "sheet"
    assert result.excel.orientation == "portrait"


def test_invalid_strict_cross_field_settings_fail():
    with pytest.raises(ValueError, match="cannot disable postflight"):
        ExcelSettings(quality_profile="strict", postflight_policy="disabled")
    with pytest.raises(ValueError, match="cannot refresh"):
        ExcelSettings(quality_profile="strict", external_link_policy="refresh_allowed")
    with pytest.raises(ValueError, match="only in the legacy"):
        ExcelSettings(quality_profile="balanced", page_size_scope="chunk")


def test_stable_decision_ids_ignore_construction_order():
    assert stable_id({"b": 2, "a": 1}) == stable_id({"a": 1, "b": 2})
    decision = LayoutDecision(
        workbook="book.xlsx", sheet="Data", sheet_index=1,
        mode="smart", region_ids=("r1",), chosen=None,
    )
    assert decision.decision_id.startswith("exq-")


def test_manifest_is_deterministic_and_atomic(tmp_path):
    decision = LayoutDecision(
        workbook="book.xlsx", sheet="Data", sheet_index=1,
        mode="smart", region_ids=("r1",), chosen=None,
    )
    manifest = ConversionManifest(
        workbook="book.xlsx", output="book.pdf", profile="strict",
        decisions=(decision,),
    )
    first = tmp_path / "first.json"
    second = tmp_path / "second.json"
    manifest.write_atomic(first)
    manifest.write_atomic(second)
    assert first.read_bytes() == second.read_bytes()


def test_manifest_v2_records_timings_without_changing_decision_identity(tmp_path):
    first = LayoutDecision(
        workbook="book.xlsx", sheet="Data", sheet_index=1,
        mode="smart", region_ids=("r1",), chosen=None, schema_version=1,
    )
    second = LayoutDecision(
        workbook="book.xlsx", sheet="Data", sheet_index=1,
        mode="smart", region_ids=("r1",), chosen=None, schema_version=2,
    )
    assert first.decision_id == second.decision_id

    manifest = ConversionManifest(
        workbook="book.xlsx", output="book.pdf", profile="strict",
        decisions=(second,), timings_ms={"export": 12.5},
        runtime_evidence={"resolved_excel_workers": 2},
    )
    data = manifest.to_dict()
    assert data["schema_version"] == 2
    assert data["timings_ms"] == {"export": 12.5}
    assert data["runtime_evidence"]["resolved_excel_workers"] == 2


def test_safe_chunks_move_boundaries_out_of_atomic_objects():
    region = ResolvedRegion(0, 1, 1, 20, 5)
    shape = PrintableObject("chart", "chart", 9, 1, 13, 5)
    forbidden = SafeChunkPlanner.forbidden_row_boundaries([], [shape])
    chunks = SafeChunkPlanner().chunks([region], 10, forbidden)
    assert chunks[0].last_row == 8
    assert chunks[0].moved_boundary is True
    assert all(not (chunk.first_row <= 9 <= chunk.last_row < 13) for chunk in chunks)


def test_candidate_scoring_favors_horizontal_splits_then_preferred_paper():
    constraints = LayoutConstraints(90, 10, 150, 24, 300, ("A4", "A3"))
    common = dict(
        orientation=1, usable_width_inches=8, usable_height_inches=10,
        width_scale=1, height_scale=1, effective_scale=1, zoom=90,
        pages_tall=2, effective_font_pt=10, effective_image_dpi=150,
    )
    wide = QualityLayoutCandidate(
        paper_enum=8, paper_name="A3", pages_wide=2,
        preferred_rank=1, **common,
    )
    narrow = QualityLayoutCandidate(
        paper_enum=9, paper_name="A4", pages_wide=1,
        preferred_rank=0, **common,
    )
    chosen, rejected = choose_candidate([wide, narrow], constraints)
    assert chosen == narrow
    assert rejected == ()


def test_candidate_rejects_page_and_readability_violations():
    constraints = LayoutConstraints(90, 10, 150, 24, 300, ("A4",))
    candidate = QualityLayoutCandidate(
        9, "A4", 1, 25, 10, 1, 1, 0.89, 89, 1, 1,
        effective_font_pt=9, effective_image_dpi=100,
    )
    chosen, rejected = choose_candidate([candidate], constraints)
    assert chosen is None
    assert len(rejected[0].rejection_reasons) >= 4


def test_printer_hard_geometry_rejects_impossible_values():
    caps = {
        88: 300, 90: 300, 110: 2550, 111: 3300,
        112: 75, 113: 90, 8: 2400, 10: 3120,
    }
    dc = SimpleNamespace(GetDeviceCaps=lambda key: caps[key])
    width, height, margins = PrinterCapabilityProvider.imageable_geometry(dc)
    assert width == pytest.approx(8)
    assert height == pytest.approx(10.4)
    assert margins == pytest.approx((0.25, 0.25, 0.3, 0.3))
    caps[8] = -1
    with pytest.raises(ValueError, match="impossible"):
        PrinterCapabilityProvider.imageable_geometry(dc)


def test_pagination_probe_records_manual_and_automatic_breaks():
    def collection(values):
        items = [
            SimpleNamespace(
                Location=SimpleNamespace(Row=row, Column=column), Type=kind
            ) for row, column, kind in values
        ]
        return SimpleNamespace(
            Count=len(items), Item=lambda index: items[index - 1]
        )

    sheet = SimpleNamespace(
        Application=SimpleNamespace(PrintCommunication=True),
        HPageBreaks=collection([(10, 1, -4135), (20, 1, 0)]),
        VPageBreaks=collection([(1, 5, 0)]),
        DisplayPageBreaks=False,
    )
    evidence = ExcelPaginationProbe().probe(sheet, "preserve")
    assert (evidence.pages_wide, evidence.pages_tall) == (2, 3)
    assert evidence.manual_horizontal == (10,)


def test_extension_strategies_fail_explicitly():
    settings = ExcelSettings(
        quality_profile="balanced", horizontal_overflow_strategy="vector_stitch"
    )
    with pytest.raises(UnsupportedExcelExtensionError, match="M6"):
        require_supported_extensions(settings, "standard")
    with pytest.raises(UnsupportedExcelExtensionError, match="PDF/A"):
        require_supported_extensions(ExcelSettings(quality_profile="strict"), "pdfa")


def test_ooxml_authored_print_metadata_is_detected(tmp_path):
    path = tmp_path / "authored.xlsx"
    with ZipFile(path, "w") as package:
        package.writestr(
            "xl/workbook.xml",
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
        )
        package.writestr(
            "xl/_rels/workbook.xml.rels",
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
        )
        package.writestr(
            "xl/worksheets/sheet1.xml",
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5"/>'
            '<pageSetup orientation="landscape"/></worksheet>',
        )
    assert persisted_print_sheets(path) == {"Data"}


def test_ooxml_default_margins_alone_are_not_authored(tmp_path):
    path = tmp_path / "defaults.xlsx"
    with ZipFile(path, "w") as package:
        package.writestr(
            "xl/workbook.xml",
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
        )
        package.writestr(
            "xl/_rels/workbook.xml.rels",
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
        )
        package.writestr(
            "xl/worksheets/sheet1.xml",
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" '
            'header="0.3" footer="0.3"/></worksheet>',
        )
    assert persisted_print_sheets(path) == set()


def test_legacy_authored_detection_uses_positive_com_signals(tmp_path):
    setup = SimpleNamespace(
        PrintArea="$A$1:$B$2", PrintTitleRows="", PrintTitleColumns="",
        LeftHeader="", CenterHeader="", RightHeader="",
        LeftFooter="", CenterFooter="", RightFooter="",
        PaperSize=9, Orientation=1, LeftMargin=1, RightMargin=1,
        TopMargin=1, BottomMargin=1, HeaderMargin=1, FooterMargin=1,
        Zoom=100, FitToPagesWide=False, FitToPagesTall=False,
        Order=1, BlackAndWhite=False, Draft=False,
    )
    sheet = SimpleNamespace(
        Name="Data", PageSetup=setup,
        Range=lambda _value: object(),
    )
    snapshot = AuthoredLayoutInspector(tmp_path / "book.xls").inspect(sheet)
    assert snapshot.classification == "authored"
    assert snapshot.source == "com-signals"


def test_pdf_postflight_rejects_blank_and_accepts_vector_page(tmp_path):
    blank = tmp_path / "blank.pdf"
    vector = tmp_path / "vector.pdf"
    writer = PdfWriter()
    writer.add_blank_page(width=600, height=800)
    with blank.open("wb") as stream:
        writer.write(stream)
    result = PdfQualityPostflight().validate(
        blank, PdfQualityExpectation(require_searchable_text=False)
    )
    assert result.passed is False

    writer = PdfWriter()
    page = writer.add_blank_page(width=600, height=800)
    content = DecodedStreamObject()
    content.set_data(b"q 0 0 0 rg 100 100 50 50 re f Q")
    page[NameObject("/Contents")] = content
    with vector.open("wb") as stream:
        writer.write(stream)
    result = PdfQualityPostflight().validate(
        vector,
        PdfQualityExpectation(
            expected_pages=1, require_searchable_text=False,
            max_dimension_in=12, max_area_in2=200,
        ),
    )
    assert result.passed is True
