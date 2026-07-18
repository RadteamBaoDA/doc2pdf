from contextlib import nullcontext
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import pytest

from src.config import ExcelSettings, PDFConversionSettings
from src.core.excel_converter import (
    ExcelConverter,
    OversizedSheetError,
    PaperForm,
    SheetRegion,
)


class StrictPageSetup:
    allowed = {
        "Application", "Orientation", "PaperSize", "LeftMargin", "RightMargin",
        "TopMargin", "BottomMargin", "Zoom", "FitToPagesWide",
        "FitToPagesTall", "BlackAndWhite", "PrintArea",
        "PrintTitleRows", "PrintTitleColumns",
    }

    def __init__(self, app):
        object.__setattr__(self, "Application", app)
        object.__setattr__(self, "Orientation", 1)
        object.__setattr__(self, "PaperSize", 1)
        object.__setattr__(self, "LeftMargin", 0)
        object.__setattr__(self, "RightMargin", 0)
        object.__setattr__(self, "TopMargin", 0)
        object.__setattr__(self, "BottomMargin", 0)
        object.__setattr__(self, "Zoom", False)
        object.__setattr__(self, "FitToPagesWide", 1)
        object.__setattr__(self, "FitToPagesTall", 1)
        object.__setattr__(self, "BlackAndWhite", False)
        object.__setattr__(self, "PrintArea", "")
        object.__setattr__(self, "PrintTitleRows", "")
        object.__setattr__(self, "PrintTitleColumns", "")

    def __setattr__(self, name, value):
        if name not in self.allowed:
            raise AttributeError(f"unknown PageSetup property {name}")
        object.__setattr__(self, name, value)


def _strict_sheet(width_inches=7.0):
    app = SimpleNamespace(
        DisplayAlerts=False, Interactive=False, PrintCommunication=True,
        ActivePrinter="Microsoft Print to PDF on PORTPROMPT:",
    )
    setup = StrictPageSetup(app)
    def range_for_title(address):
        if any(char.isdigit() for char in address):
            return SimpleNamespace(
                Address=address, Width=width_inches * 72, Height=15
            )
        return SimpleNamespace(Address=address, Width=72, Height=5 * 72)

    sheet = SimpleNamespace(
        Name="Data", Application=app, PageSetup=setup, Range=range_for_title
    )
    return sheet, width_inches * 72


TEST_PAPER_FORMS = (
    PaperForm(9, "A4", 8.27, 11.69),
    PaperForm(8, "A3", 11.69, 16.54),
    PaperForm(26, "E", 34.0, 44.0),
)


def test_layout_candidate_accounts_for_height_when_one_page_tall_is_required():
    candidate = ExcelConverter._build_layout_candidate(
        PaperForm(100, "Square", 10.0, 10.0),
        orientation=1,
        content_width_inches=5.0,
        content_height_inches=20.0,
        fit_tall=True,
        margins_points=(0.0, 0.0, 0.0, 0.0),
        quality_zoom=90,
    )

    assert candidate.width_scale == pytest.approx(2.0)
    assert candidate.height_scale == pytest.approx(0.5)
    assert candidate.effective_scale == pytest.approx(0.5)
    assert candidate.max_zoom == 50
    assert candidate.limiting_axis == "height"


def test_layout_candidate_allows_vertical_pagination_without_shrinking_height():
    candidate = ExcelConverter._build_layout_candidate(
        PaperForm(100, "Square", 10.0, 10.0),
        orientation=1,
        content_width_inches=5.0,
        content_height_inches=20.0,
        fit_tall=False,
        margins_points=(0.0, 0.0, 0.0, 0.0),
        quality_zoom=90,
    )

    assert candidate.effective_scale == pytest.approx(1.0)
    assert candidate.max_zoom == 100
    assert candidate.pages_wide == 1
    assert candidate.pages_tall == 2
    assert candidate.page_count == 2


def test_error_mode_with_null_row_dimensions_still_requires_one_page_high():
    converter = ExcelConverter()
    sheet, _ = _strict_sheet()
    forms = (PaperForm(115, "TenByTen", 10.0, 10.0),)

    with patch.object(converter, "_get_printer_paper_forms", return_value=forms):
        with pytest.raises(ValueError, match="below min_shrink_factor"):
            converter._apply_page_setup(
                sheet,
                ExcelSettings(
                    orientation="portrait",
                    row_dimensions=None,
                    metadata_header=False,
                    min_shrink_factor=0.90,
                    oversized_action="error",
                ),
                "book.xlsx",
                5,
                content_width_points=5 * 72,
                content_height_points=20 * 72,
            )


def test_fit_candidate_accepts_exact_quality_floor_and_rejects_just_below():
    exact = ExcelConverter._build_layout_candidate(
        PaperForm(101, "Exact", 9.0, 20.0), 1, 10.0, 10.0, True,
        (0.0, 0.0, 0.0, 0.0), 90,
    )
    below = ExcelConverter._build_layout_candidate(
        PaperForm(102, "Below", 8.99, 20.0), 1, 10.0, 10.0, True,
        (0.0, 0.0, 0.0, 0.0), 90,
    )

    assert exact.max_zoom == 90
    assert below.max_zoom == 89
    assert ExcelConverter._select_fit_candidate([below], 90) is None
    assert ExcelConverter._select_fit_candidate([below, exact], 90) == exact


def test_auto_orientation_selects_better_scale_while_forced_set_does_not_expand():
    form = PaperForm(103, "EightByTwelve", 8.0, 12.0)
    portrait = ExcelConverter._build_layout_candidate(
        form, 1, 10.0, 6.0, True, (0.0, 0.0, 0.0, 0.0), 90,
    )
    landscape = ExcelConverter._build_layout_candidate(
        form, 2, 10.0, 6.0, True, (0.0, 0.0, 0.0, 0.0), 90,
    )

    assert ExcelConverter._select_fit_candidate([portrait], 90) is None
    assert ExcelConverter._select_fit_candidate([portrait, landscape], 90) == landscape


def test_fit_candidate_tie_breaks_on_smaller_paper_independent_of_input_order():
    small = ExcelConverter._build_layout_candidate(
        PaperForm(104, "Small", 10.0, 10.0), 1, 5.0, 5.0, True,
        (0.0, 0.0, 0.0, 0.0), 90,
    )
    large = ExcelConverter._build_layout_candidate(
        PaperForm(105, "Large", 20.0, 20.0), 1, 5.0, 5.0, True,
        (0.0, 0.0, 0.0, 0.0), 90,
    )

    assert ExcelConverter._select_fit_candidate([large, small], 90) == small
    assert ExcelConverter._select_fit_candidate([small, large], 90) == small


def test_paginated_candidate_minimizes_pages_then_paper_area():
    small_two_pages = ExcelConverter._build_layout_candidate(
        PaperForm(106, "Small", 10.0, 10.0), 1, 18.0, 5.0, True,
        (0.0, 0.0, 0.0, 0.0), 90,
    )
    large_two_pages = ExcelConverter._build_layout_candidate(
        PaperForm(107, "Large", 9.0, 20.0), 1, 18.0, 5.0, True,
        (0.0, 0.0, 0.0, 0.0), 90,
    )
    one_page = ExcelConverter._build_layout_candidate(
        PaperForm(108, "OnePage", 20.0, 20.0), 1, 18.0, 5.0, True,
        (0.0, 0.0, 0.0, 0.0), 90,
    )

    assert small_two_pages.page_count == large_two_pages.page_count == 2
    assert ExcelConverter._select_paginated_candidate(
        [large_two_pages, small_two_pages]
    ) == small_two_pages
    assert ExcelConverter._select_paginated_candidate(
        [small_two_pages, one_page]
    ) == one_page


def test_paginated_candidate_accounts_for_repeated_print_titles():
    without_titles = ExcelConverter._build_layout_candidate(
        PaperForm(109, "Square", 10.0, 10.0), 1, 5.0, 20.0, False,
        (0.0, 0.0, 0.0, 0.0), 100,
    )
    with_titles = ExcelConverter._build_layout_candidate(
        PaperForm(109, "Square", 10.0, 10.0), 1, 5.0, 20.0, False,
        (0.0, 0.0, 0.0, 0.0), 100,
        title_height_inches=2.0,
    )

    assert without_titles.pages_tall == 2
    assert with_titles.pages_tall == 3


def test_print_title_measurement_adds_only_the_portion_outside_print_area():
    print_area = SimpleNamespace(Width=4 * 72, Height=10 * 15)
    title_columns = SimpleNamespace(Width=72, Height=10 * 15)
    title_rows = SimpleNamespace(Width=4 * 72, Height=15)
    column_overlap = None
    row_overlap = SimpleNamespace(Width=4 * 72, Height=15)

    def intersect(_print_area, title_range):
        return row_overlap if title_range is title_rows else column_overlap

    app = SimpleNamespace(Intersect=intersect)
    setup = SimpleNamespace(
        PrintArea="$C$1:$F$10",
        PrintTitleColumns="$A:$A",
        PrintTitleRows="$1:$1",
    )
    ranges = {
        setup.PrintArea: print_area,
        setup.PrintTitleColumns: title_columns,
        setup.PrintTitleRows: title_rows,
    }
    sheet = SimpleNamespace(
        Application=app,
        PageSetup=setup,
        Range=lambda address: ranges[address],
    )

    title_width, title_height, extra_width, extra_height = (
        ExcelConverter._measure_print_titles(sheet)
    )

    assert title_width == pytest.approx(1.0)
    assert title_height == pytest.approx(15 / 72)
    assert extra_width == pytest.approx(1.0)
    assert extra_height == pytest.approx(0.0)


def test_nonempty_print_title_that_cannot_be_resolved_is_fatal():
    setup = SimpleNamespace(
        PrintArea="",
        PrintTitleColumns="",
        PrintTitleRows="$1:$1",
    )
    sheet = SimpleNamespace(
        Name="Data",
        PageSetup=setup,
        Range=MagicMock(side_effect=RuntimeError("bad range")),
    )

    with pytest.raises(ValueError, match="cannot measure PrintTitleRows"):
        ExcelConverter._measure_print_titles(sheet)


def test_page_setup_uses_numeric_zoom_and_only_documented_properties():
    converter = ExcelConverter()
    sheet, width = _strict_sheet()
    with patch.object(
        converter, "_get_printer_paper_forms", return_value=TEST_PAPER_FORMS
    ):
        converter._apply_page_setup(
            sheet, ExcelSettings(row_dimensions=None), "book.xlsx", 8,
            content_width_points=width,
            content_height_points=5 * 72,
        )
    assert isinstance(sheet.PageSetup.Zoom, int)
    assert sheet.PageSetup.Zoom == 100
    assert sheet.PageSetup.FitToPagesWide is False
    assert sheet.PageSetup.FitToPagesTall is False
    assert not hasattr(sheet.PageSetup, "PaperWidth")


def test_page_setup_honors_forced_orientation_and_auto_chooses_better_scale():
    converter = ExcelConverter()
    form = (PaperForm(110, "EightByTwelve", 8.0, 12.0),)
    portrait_sheet, _ = _strict_sheet()
    auto_sheet, _ = _strict_sheet()

    with patch.object(converter, "_get_printer_paper_forms", return_value=form):
        converter._apply_page_setup(
            portrait_sheet,
            ExcelSettings(
                orientation="portrait",
                min_shrink_factor=0.50,
                metadata_header=False,
            ),
            "book.xlsx",
            10,
            content_width_points=10 * 72,
            content_height_points=5 * 72,
        )
        converter._apply_page_setup(
            auto_sheet,
            ExcelSettings(min_shrink_factor=0.50, metadata_header=False),
            "book.xlsx",
            10,
            content_width_points=10 * 72,
            content_height_points=5 * 72,
        )

    assert portrait_sheet.PageSetup.Orientation == 1
    assert portrait_sheet.PageSetup.Zoom == 70
    assert auto_sheet.PageSetup.Orientation == 2
    assert auto_sheet.PageSetup.Zoom == 100


def test_page_setup_does_not_count_vertical_margins_twice():
    converter = ExcelConverter()
    sheet, _ = _strict_sheet()
    form = (PaperForm(111, "TenByTen", 10.0, 10.0),)

    with patch.object(converter, "_get_printer_paper_forms", return_value=form):
        converter._apply_page_setup(
            sheet,
            ExcelSettings(
                orientation="portrait",
                row_dimensions=0,
                metadata_header=False,
                min_shrink_factor=0.90,
                oversized_action="error",
            ),
            "book.xlsx",
            5,
            content_width_points=5 * 72,
            content_height_points=10 * 72,
        )

    assert sheet.PageSetup.Zoom == 90


@pytest.mark.parametrize(
    ("min_shrink_factor", "expected_zoom"),
    [(0.90, 90), (0.901, 91)],
)
def test_paginate_keeps_ceil_quality_zoom_and_disables_fit_flags(
    min_shrink_factor, expected_zoom
):
    converter = ExcelConverter()
    sheet, _ = _strict_sheet()
    with patch.object(
        converter, "_get_printer_paper_forms", return_value=TEST_PAPER_FORMS
    ):
        converter._apply_page_setup(
            sheet,
            ExcelSettings(
                row_dimensions=0,
                min_shrink_factor=min_shrink_factor,
                oversized_action="paginate",
            ),
            "book.xlsx",
            200,
            content_width_points=100 * 72,
            content_height_points=100 * 72,
        )
    assert sheet.PageSetup.Zoom == expected_zoom
    assert sheet.PageSetup.FitToPagesWide is False
    assert sheet.PageSetup.FitToPagesTall is False


def test_print_titles_are_preserved_by_default_and_overridden_when_configured():
    converter = ExcelConverter()
    preserved_sheet, width = _strict_sheet()
    configured_sheet, _ = _strict_sheet()
    preserved_sheet.PageSetup.PrintTitleRows = "$1:$1"
    preserved_sheet.PageSetup.PrintTitleColumns = "$A:$A"

    with patch.object(
        converter, "_get_printer_paper_forms", return_value=TEST_PAPER_FORMS
    ):
        converter._apply_page_setup(
            preserved_sheet,
            ExcelSettings(),
            "book.xlsx",
            8,
            content_width_points=width,
            content_height_points=5 * 72,
        )
        converter._apply_page_setup(
            configured_sheet,
            ExcelSettings(
                print_title_rows="$1:$2",
                print_title_columns="$A:$B",
            ),
            "book.xlsx",
            8,
            content_width_points=width,
            content_height_points=5 * 72,
        )

    assert preserved_sheet.PageSetup.PrintTitleRows == "$1:$1"
    assert preserved_sheet.PageSetup.PrintTitleColumns == "$A:$A"
    assert configured_sheet.PageSetup.PrintTitleRows == "$1:$2"
    assert configured_sheet.PageSetup.PrintTitleColumns == "$A:$B"


@pytest.mark.parametrize("advertised", [False, True])
def test_a2_is_available_in_fallback_only_when_advertised(advertised):
    converter = ExcelConverter()
    app = SimpleNamespace(
        ActivePrinter="Test PDF Printer on TEST:",
    )

    with patch.object(
        converter, "_printer_advertises_a2", return_value=advertised
    ), patch(
        "src.core.excel_converter.win32print.OpenPrinter",
        side_effect=OSError("no test printer"),
    ):
        forms = converter._get_printer_paper_forms(app)

    assert (66 in {form.paper_enum for form in forms}) is advertised


def test_printer_advertised_dimensions_override_fallback_catalog():
    converter = ExcelConverter()
    app = SimpleNamespace(ActivePrinter="Test PDF Printer on TEST:")

    def device_capabilities(_printer, _port, capability):
        return {
            2: [9],
            3: [(3000, 4000)],
            16: ["Driver A4"],
        }[capability]

    with patch(
        "src.core.excel_converter.win32print.OpenPrinter", return_value=123
    ), patch(
        "src.core.excel_converter.win32print.GetPrinter",
        return_value={"pPortName": "TEST:"},
    ), patch(
        "src.core.excel_converter.win32print.DeviceCapabilities",
        side_effect=device_capabilities,
    ), patch("src.core.excel_converter.win32print.ClosePrinter") as close:
        forms = converter._get_printer_paper_forms(app)

    assert len(forms) == 1
    assert forms[0].paper_enum == 9
    assert forms[0].name == "Driver A4"
    assert forms[0].width_inches == pytest.approx(3000 / 254)
    assert forms[0].height_inches == pytest.approx(4000 / 254)
    close.assert_called_once_with(123)


def test_rejected_paper_forms_are_not_selected():
    converter = ExcelConverter()
    sheet, width = _strict_sheet()

    def accept_only_a4(page_setup, paper_enum, _paper_name, timeout_seconds=3):
        if paper_enum != 9:
            return False
        page_setup.PaperSize = paper_enum
        return True

    with (
        patch.object(
            converter,
            "_get_printer_paper_forms",
            return_value=TEST_PAPER_FORMS,
        ),
        patch.object(
            converter, "_try_set_paper_size", side_effect=accept_only_a4
        ),
    ):
        converter._apply_page_setup(
            sheet,
            ExcelSettings(),
            "book.xlsx",
            8,
            content_width_points=width,
            content_height_points=5 * 72,
        )

    assert sheet.PageSetup.PaperSize == 9


def test_orientation_probe_rejects_a_silently_ignored_orientation():
    converter = ExcelConverter()
    app = SimpleNamespace(
        DisplayAlerts=False,
        Interactive=False,
        PrintCommunication=True,
        ActivePrinter="Test PDF Printer on TEST:",
    )

    class RejectingLandscape(StrictPageSetup):
        def __setattr__(self, name, value):
            if name == "Orientation" and value == 2:
                return
            super().__setattr__(name, value)

    setup = RejectingLandscape(app)
    sheet = SimpleNamespace(Name="Data", Application=app, PageSetup=setup)
    form = (PaperForm(112, "EightByTwelve", 8.0, 12.0),)

    with patch.object(converter, "_get_printer_paper_forms", return_value=form):
        converter._apply_page_setup(
            sheet,
            ExcelSettings(min_shrink_factor=0.50, metadata_header=False),
            "book.xlsx",
            10,
            content_width_points=10 * 72,
            content_height_points=5 * 72,
        )

    assert sheet.PageSetup.Orientation == 1
    assert sheet.PageSetup.Zoom == 70


def test_page_setup_reuses_printer_normalized_margins_from_selected_candidate():
    converter = ExcelConverter()
    app = SimpleNamespace(
        DisplayAlerts=False,
        Interactive=False,
        PrintCommunication=True,
        ActivePrinter="Test PDF Printer on TEST:",
    )

    class NormalizingMargins(StrictPageSetup):
        def __setattr__(self, name, value):
            if name in {"LeftMargin", "RightMargin", "BottomMargin"}:
                value = max(float(value), 54.0)
            super().__setattr__(name, value)

    setup = NormalizingMargins(app)
    sheet = SimpleNamespace(Name="Data", Application=app, PageSetup=setup)
    form = (PaperForm(113, "TenByTen", 10.0, 10.0),)

    with patch.object(converter, "_get_printer_paper_forms", return_value=form):
        converter._apply_page_setup(
            sheet,
            ExcelSettings(orientation="portrait"),
            "book.xlsx",
            5,
            content_width_points=5 * 72,
            content_height_points=5 * 72,
        )

    assert sheet.PageSetup.LeftMargin == 54.0
    assert sheet.PageSetup.RightMargin == 54.0
    assert sheet.PageSetup.TopMargin == 72.0
    assert sheet.PageSetup.BottomMargin == 54.0


def test_readback_failure_is_fatal():
    converter = ExcelConverter()
    app = SimpleNamespace(PrintCommunication=True)

    class Rejecting(StrictPageSetup):
        def __setattr__(self, name, value):
            if name == "Zoom":
                return
            super().__setattr__(name, value)

    setup = Rejecting(app)
    with pytest.raises(ValueError, match="did not retain"):
        converter._required_set_page_property(setup, "Zoom", 90)


def test_final_readback_detects_a_later_property_resetting_paper_size():
    converter = ExcelConverter()
    app = SimpleNamespace(
        DisplayAlerts=False,
        Interactive=False,
        PrintCommunication=True,
        ActivePrinter="Test PDF Printer on TEST:",
    )

    class ResetsPaperOnNumericZoom(StrictPageSetup):
        def __setattr__(self, name, value):
            super().__setattr__(name, value)
            if name == "Zoom" and type(value) is int:
                object.__setattr__(self, "PaperSize", 999)

    setup = ResetsPaperOnNumericZoom(app)
    sheet = SimpleNamespace(Name="Data", Application=app, PageSetup=setup)
    forms = (PaperForm(114, "TenByTen", 10.0, 10.0),)

    with patch.object(converter, "_get_printer_paper_forms", return_value=forms):
        with pytest.raises(ValueError, match="final PageSetup.PaperSize"):
            converter._apply_page_setup(
                sheet,
                ExcelSettings(orientation="portrait", metadata_header=False),
                "book.xlsx",
                5,
                content_width_points=5 * 72,
                content_height_points=5 * 72,
            )


def test_final_readback_detects_a_preserved_print_title_being_cleared():
    converter = ExcelConverter()
    app = SimpleNamespace(
        DisplayAlerts=False,
        Interactive=False,
        PrintCommunication=True,
        ActivePrinter="Test PDF Printer on TEST:",
    )

    class ClearsTitleOnNumericZoom(StrictPageSetup):
        def __setattr__(self, name, value):
            super().__setattr__(name, value)
            if name == "Zoom" and type(value) is int:
                object.__setattr__(self, "PrintTitleRows", "")

    setup = ClearsTitleOnNumericZoom(app)
    setup.PrintTitleRows = "$1:$1"
    sheet = SimpleNamespace(
        Name="Data",
        Application=app,
        PageSetup=setup,
        Range=lambda address: SimpleNamespace(
            Address=address, Width=5 * 72, Height=15
        ),
    )
    forms = (PaperForm(116, "TenByTen", 10.0, 10.0),)

    with patch.object(converter, "_get_printer_paper_forms", return_value=forms):
        with pytest.raises(ValueError, match="PrintTitleRows"):
            converter._apply_page_setup(
                sheet,
                ExcelSettings(orientation="portrait", metadata_header=False),
                "book.xlsx",
                5,
                content_width_points=5 * 72,
                content_height_points=5 * 72,
            )


def test_preserve_multiple_non_a1_print_areas():
    converter = ExcelConverter()
    area1 = SimpleNamespace(
        Row=3,
        Column=2,
        Rows=SimpleNamespace(Count=4),
        Columns=SimpleNamespace(Count=5),
    )
    area2 = SimpleNamespace(
        Row=20,
        Column=8,
        Rows=SimpleNamespace(Count=2),
        Columns=SimpleNamespace(Count=3),
    )
    areas = MagicMock(Count=2)
    areas.side_effect = lambda index: [area1, area2][index - 1]
    range_object = SimpleNamespace(Areas=areas)
    sheet = MagicMock()
    sheet.PageSetup.PrintArea = "$B$3:$F$6,$H$20:$J$21"
    sheet.Range.return_value = range_object

    assert converter._resolve_sheet_regions(sheet, "preserve") == [
        SheetRegion(3, 2, 6, 6), SheetRegion(20, 8, 21, 10)
    ]


def test_oversized_action_error_and_skip():
    converter = ExcelConverter()
    error_sheet, _ = _strict_sheet()
    skip_sheet, _ = _strict_sheet()
    with patch.object(
        converter, "_get_printer_paper_forms", return_value=TEST_PAPER_FORMS
    ):
        with pytest.raises(ValueError, match="below min_shrink_factor"):
            converter._apply_page_setup(
                error_sheet,
                ExcelSettings(
                    min_shrink_factor=0.90,
                    oversized_action="error",
                    row_dimensions=0,
                ),
                "book.xlsx",
                100,
                content_width_points=100 * 72,
                content_height_points=100 * 72,
            )
        with pytest.raises(OversizedSheetError):
            converter._apply_page_setup(
                skip_sheet,
                ExcelSettings(
                    min_shrink_factor=0.90,
                    oversized_action="skip",
                    row_dimensions=0,
                ),
                "book.xlsx",
                100,
                content_width_points=100 * 72,
                content_height_points=100 * 72,
            )


def test_oversized_action_warn_keeps_legacy_fit_mode():
    converter = ExcelConverter()
    sheet, _ = _strict_sheet()

    with (
        patch.object(
            converter, "_get_printer_paper_forms", return_value=TEST_PAPER_FORMS
        ),
        patch("src.core.excel_converter.logger.warning") as warning,
    ):
        converter._apply_page_setup(
            sheet,
            ExcelSettings(
                min_shrink_factor=0.90,
                oversized_action="warn",
                row_dimensions=0,
            ),
            "book.xlsx",
            100,
            content_width_points=100 * 72,
            content_height_points=100 * 72,
        )

    warning.assert_called()
    assert sheet.PageSetup.Zoom is False
    assert sheet.PageSetup.FitToPagesWide == 1
    assert sheet.PageSetup.FitToPagesTall == 1


def test_skip_rolls_back_previously_staged_chunks_and_page_count(tmp_path):
    converter = ExcelConverter()
    input_path = tmp_path / "book.xlsx"
    output_path = tmp_path / "book.pdf"
    input_path.touch()

    app = MagicMock()
    app.Version = "16"
    workbook = MagicMock()
    workbook.Application = app
    app.Workbooks.Open.return_value = workbook

    skipped_source = MagicMock(Name="Skipped")
    valid_source = MagicMock(Name="Valid")
    skipped_first = MagicMock(Name="Skipped-1")
    skipped_second = MagicMock(Name="Skipped-2")
    valid_copy = MagicMock(Name="Valid-1")
    for copied in (skipped_first, skipped_second, valid_copy):
        copied.Range.return_value = SimpleNamespace(Width=5 * 72, Height=5 * 72)

    skipped_regions = [
        SheetRegion(1, 1, 10, 5),
        SheetRegion(11, 1, 20, 5),
    ]
    valid_regions = [SheetRegion(1, 1, 10, 5)]
    settings = PDFConversionSettings(
        excel=ExcelSettings(
            row_dimensions=0,
            oversized_action="skip",
            metadata_header=False,
            print_area_policy="auto",
        )
    )
    exported_sheets = []

    def resolve_regions(sheet, _policy):
        return skipped_regions if sheet is skipped_source else valid_regions

    def export_pdf(_workbook, sheets, stage_path, _settings):
        exported_sheets.extend(sheets)
        Path(stage_path).write_bytes(b"staged-pdf")

    with (
        patch.object(
            converter, "_excel_application", return_value=nullcontext(app)
        ),
        patch.object(
            converter,
            "_get_sheets_to_export",
            return_value=[skipped_source, valid_source],
        ),
        patch(
            "src.core.excel_converter.get_excel_sheet_settings",
            return_value=settings,
        ),
        patch.object(
            converter, "_resolve_sheet_regions", side_effect=resolve_regions
        ),
        patch.object(
            converter,
            "_copy_region_sheet",
            side_effect=[skipped_first, skipped_second, valid_copy],
        ),
        patch.object(
            converter,
            "_apply_page_setup",
            side_effect=[None, OversizedSheetError("too large"), None],
        ),
        patch.object(converter, "_export_to_pdf", side_effect=export_pdf),
        patch(
            "src.core.excel_converter.pythoncom.CoInitialize"
        ),
        patch(
            "src.core.excel_converter.pythoncom.CoUninitialize"
        ),
        patch("pypdf.PdfReader", return_value=SimpleNamespace(pages=[object()])),
    ):
        result = converter.convert(input_path, output_path, settings)

    assert result == output_path.resolve()
    assert exported_sheets == [valid_copy]
    assert output_path.read_bytes() == b"staged-pdf"


def test_multi_sheet_temp_workbook_copy_failure_is_fatal():
    converter = ExcelConverter()
    app = MagicMock()
    app.Version = "16"
    workbook = MagicMock()
    workbook.Application = app
    temp_workbook = MagicMock()
    temp_workbook.Sheets.Count = 1
    app.ActiveWorkbook = temp_workbook

    first = MagicMock(Name="First")
    second = MagicMock(Name="Second")
    second.Copy.side_effect = RuntimeError("copy failed")

    with pytest.raises(ValueError):
        converter._export_to_pdf(
            workbook,
            [first, second],
            "out.pdf",
            PDFConversionSettings(),
        )

    temp_workbook.ExportAsFixedFormat.assert_not_called()
    temp_workbook.Close.assert_called_once_with(SaveChanges=False)


def test_excel_application_uses_dispatch_ex():
    converter = ExcelConverter()
    app = MagicMock()
    app.Version = "16"
    with patch(
        "src.core.excel_converter.win32com.client.DispatchEx", return_value=app
    ) as dispatch, patch.object(
        converter, "_set_optimal_printer"
    ), patch(
        "src.core.excel_converter.ProcessRegistry"
    ):
        with converter._excel_application() as actual:
            assert actual is app
    dispatch.assert_called_once_with("Excel.Application")
