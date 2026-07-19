"""Printable content inventory for Excel worksheets."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, List, Tuple

from .models import PrintableObject, ResolvedRegion, WorkbookSheetInfo


def _range_bounds(value: Any) -> Tuple[int, int, int, int]:
    """Document this Excel pipeline operation and its side effects."""
    return (
        int(value.Row), int(value.Column),
        int(value.Row) + int(value.Rows.Count) - 1,
        int(value.Column) + int(value.Columns.Count) - 1,
    )


@dataclass(frozen=True)
class ContentInventory:
    regions: Tuple[ResolvedRegion, ...]
    objects: Tuple[PrintableObject, ...]
    errors: Tuple[str, ...] = ()

    @property
    def certain(self) -> bool:
        """Document this Excel pipeline operation and its side effects."""
        return not self.errors and all(item.confidence == "certain" for item in self.regions)


class PrintableContentResolver:
    def inventory_sheets(self, workbook: Any, selected_name: str | None = None) -> Tuple[WorkbookSheetInfo, ...]:
        """Document this Excel pipeline operation and its side effects."""
        result = []
        sheets = workbook.Sheets
        # Sheets includes chart sheets; preserving this collection order preserves workbook order.
        for index in range(1, int(sheets.Count) + 1):
            sheet = sheets.Item(index)
            name = str(sheet.Name)
            if selected_name and name != selected_name:
                continue
            visible = getattr(sheet, "Visible", -1) == -1
            type_value = getattr(sheet, "Type", None)
            kind = "chart" if type_value == -4109 or not hasattr(sheet, "Cells") else "worksheet"
            result.append(WorkbookSheetInfo(index, name, kind, visible))
        return tuple(result)

    def resolve(self, sheet: Any, policy: str, strict: bool = False) -> ContentInventory:
        """Document this Excel pipeline operation and its side effects."""
        errors: List[str] = []
        objects = self._objects(sheet, errors)
        print_area = ""
        try:
            print_area = str(sheet.PageSetup.PrintArea or "")
        except Exception as exc:
            errors.append(f"PrintArea: {exc}")
        regions: List[ResolvedRegion] = []
        if print_area and policy in {"preserve", "preserve_strict", "expand_visible_objects"}:
            # Authored multi-area PrintAreas are retained as separate regions.
            try:
                area_set = sheet.Range(print_area).Areas
                for index in range(1, int(area_set.Count) + 1):
                    first_row, first_col, last_row, last_col = _range_bounds(area_set(index))
                    regions.append(ResolvedRegion(
                        index - 1, first_row, first_col, last_row, last_col,
                        ("print_area",), "certain",
                    ))
            except Exception as exc:
                errors.append(f"invalid authored PrintArea: {exc}")
                if policy == "preserve_strict" or strict:
                    return ContentInventory((), tuple(objects), tuple(errors))
        if not regions:
            try:
                used = sheet.UsedRange
                first_row, first_col, last_row, last_col = _range_bounds(used)
                regions.append(ResolvedRegion(
                    0, first_row, first_col, last_row, last_col,
                    ("used_range",), "certain",
                ))
            except Exception as exc:
                errors.append(f"UsedRange: {exc}")
        if objects and (not regions or policy in {"auto", "expand_visible_objects"}):
            # Expand discovery to include printable objects outside UsedRange.
            min_row = min([item.first_row for item in objects] + [r.first_row for r in regions])
            min_col = min([item.first_col for item in objects] + [r.first_col for r in regions])
            max_row = max([item.last_row for item in objects] + [r.last_row for r in regions])
            max_col = max([item.last_col for item in objects] + [r.last_col for r in regions])
            regions = [ResolvedRegion(
                0, min_row, min_col, max_row, max_col,
                tuple(sorted({source for region in regions for source in region.sources} | {"objects"})),
                "uncertain" if errors else "certain", tuple(errors),
            )]
        if errors:
            regions = [ResolvedRegion(
                region.order, region.first_row, region.first_col,
                region.last_row, region.last_col, region.sources,
                "uncertain", tuple(errors),
            ) for region in regions]
        return ContentInventory(tuple(regions), tuple(objects), tuple(errors))

    @staticmethod
    def _objects(sheet: Any, errors: List[str]) -> List[PrintableObject]:
        """Document this Excel pipeline operation and its side effects."""
        result: List[PrintableObject] = []
        try:
            shapes = sheet.Shapes
            count = int(shapes.Count)
        except Exception as exc:
            # Chart sheets and simple test doubles need not expose Shapes.
            if hasattr(sheet, "Cells"):
                errors.append(f"Shapes: {exc}")
            return result
        for index in range(1, count + 1):
            try:
                shape = shapes.Item(index)
                printable = bool(getattr(shape, "PrintObject", True))
                if not printable:
                    continue
                top_left = shape.TopLeftCell
                bottom_right = shape.BottomRightCell
                result.append(PrintableObject(
                    str(getattr(shape, "Name", index)),
                    str(getattr(shape, "Type", "shape")),
                    int(top_left.Row), int(top_left.Column),
                    int(bottom_right.Row), int(bottom_right.Column), True,
                ))
            except Exception as exc:
                errors.append(f"Shape[{index}]: {exc}")
        return result
