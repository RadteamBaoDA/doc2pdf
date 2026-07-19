"""Create Excel workbooks that exercise the active config.yml Excel rules."""

from pathlib import Path

import xlsxwriter


OUTPUT_DIR = Path(__file__).parent / "config_examples"


def add_table(workbook, sheet, rows: int, columns: int, *, title: str) -> None:
    title_format = workbook.add_format(
        {"bold": True, "font_size": 16, "font_color": "#FFFFFF", "bg_color": "#1F4E78", "align": "center"}
    )
    header_format = workbook.add_format(
        {"bold": True, "font_color": "#FFFFFF", "bg_color": "#5B9BD5", "border": 1, "align": "center"}
    )
    text_format = workbook.add_format({"border": 1})
    number_format = workbook.add_format({"border": 1, "num_format": "#,##0.00"})
    sheet.merge_range(0, 0, 0, columns - 1, title, title_format)
    sheet.write_row(2, 0, ["LEFT_SENTINEL", *[f"Metric {n}" for n in range(2, columns)], "RIGHT_SENTINEL"], header_format)
    for row in range(rows):
        sheet.write(row + 3, 0, f"row-{row + 1}", text_format)
        for column in range(1, columns - 1):
            sheet.write(row + 3, column, (row + 1) * (column + 1), number_format)
        sheet.write(row + 3, columns - 1, f"right-{row + 1}", text_format)
    sheet.set_row(0, 28)
    sheet.set_row(2, 24)
    sheet.set_column(0, columns - 1, 14)
    sheet.freeze_panes(3, 1)
    sheet.autofilter(2, 0, rows + 2, columns - 1)
    sheet.repeat_rows(0, 2)
    sheet.repeat_columns(0, 0)


def create_auto_layout_workbook() -> None:
    path = OUTPUT_DIR / "config_auto_layout.xlsx"
    with xlsxwriter.Workbook(path) as workbook:
        summary = workbook.add_worksheet("Sales Summary")
        add_table(workbook, summary, 8, 6, title="Summary rule: one-page attempt and file-path label")
        summary.set_landscape()
        summary.fit_to_pages(1, 1)
        summary.print_area(0, 0, 10, 5)

        data = workbook.add_worksheet("Wide Data")
        add_table(workbook, data, 35, 30, title="Data rule: ten-row chunks with auto orientation")
        data.set_landscape()
        data.fit_to_pages(1, 0)
        data.print_area(0, 0, 37, 29)

        tall = workbook.add_worksheet("Tall Data")
        add_table(workbook, tall, 75, 10, title="Data rule: vertical chunking and repeating titles")
        tall.set_portrait()
        tall.fit_to_pages(1, 0)
        tall.print_area(0, 0, 77, 9)


def create_print_trim_workbook() -> None:
    path = OUTPUT_DIR / "config_print_area_trim.xlsx"
    with xlsxwriter.Workbook(path) as workbook:
        sheet = workbook.add_worksheet("Print Area Data")
        add_table(workbook, sheet, 18, 16, title="Preserved print area with whitespace for PDF trimming")
        note = workbook.add_format({"italic": True, "font_color": "#666666"})
        sheet.write("A25", "This row intentionally remains outside the print area.", note)
        sheet.set_landscape()
        sheet.set_margins(left=0.9, right=0.9, top=1.0, bottom=1.0)
        sheet.fit_to_pages(1, 0)
        # The explicit area exercises print_area_policy: preserve; margins leave
        # whitespace for post_processing.trim_whitespace.
        sheet.print_area(0, 0, 20, 15)


def create_appendix_workbook() -> None:
    path = OUTPUT_DIR / "appendix_oversized.xlsx"
    with xlsxwriter.Workbook(path) as workbook:
        normal = workbook.add_worksheet("Appendix Summary")
        add_table(workbook, normal, 8, 8, title="Normal sheet before oversized appendix data")
        normal.print_area(0, 0, 10, 7)

        wide = workbook.add_worksheet("Appendix Data")
        add_table(workbook, wide, 12, 80, title="Path rule: appendix requires oversized_action=error")
        wide.set_column(0, 79, 18)
        wide.set_landscape()
        wide.print_area(0, 0, 14, 79)


def create_empty_workbook() -> None:
    path = OUTPUT_DIR / "config_empty_visible.xlsx"
    with xlsxwriter.Workbook(path) as workbook:
        workbook.add_worksheet("EmptyVisible")
        hidden = workbook.add_worksheet("Hidden Data")
        hidden.write("A1", "This sheet is hidden and should not be exported")
        hidden.hide()


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    create_auto_layout_workbook()
    create_print_trim_workbook()
    create_appendix_workbook()
    create_empty_workbook()
    print(f"Created config examples in {OUTPUT_DIR.resolve()}")


if __name__ == "__main__":
    main()
