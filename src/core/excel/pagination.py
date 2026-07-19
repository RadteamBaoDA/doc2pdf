"""Actual Excel page-break probing after PageSetup is committed."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Tuple


XL_PAGE_BREAK_MANUAL = -4135
XL_DOWN_THEN_OVER = 1


@dataclass(frozen=True)
class PaginationEvidence:
    pages_wide: int
    pages_tall: int
    horizontal_breaks: Tuple[int, ...]
    vertical_breaks: Tuple[int, ...]
    manual_horizontal: Tuple[int, ...]
    manual_vertical: Tuple[int, ...]


def _breaks(collection: Any, axis: str) -> Tuple[Tuple[int, ...], Tuple[int, ...]]:
    """Document this Excel pipeline operation and its side effects."""
    all_breaks = []
    manual = []
    for index in range(1, int(collection.Count) + 1):
        item = collection.Item(index)
        location = item.Location
        value = int(location.Row if axis == "row" else location.Column)
        all_breaks.append(value)
        if getattr(item, "Type", None) == XL_PAGE_BREAK_MANUAL:
            manual.append(value)
    return tuple(all_breaks), tuple(manual)


class ExcelPaginationProbe:
    def probe(self, sheet: Any, manual_policy: str) -> PaginationEvidence:
        """Document this Excel pipeline operation and its side effects."""
        if manual_policy == "reset":
            sheet.ResetAllPageBreaks()
        application = sheet.Application
        try:
            application.PrintCommunication = True
        except Exception:
            pass
        # DisplayPageBreaks forces Excel to materialize automatic breaks.
        try:
            sheet.DisplayPageBreaks = True
        except Exception:
            pass
        horizontal, manual_horizontal = _breaks(sheet.HPageBreaks, "row")
        vertical, manual_vertical = _breaks(sheet.VPageBreaks, "column")
        # Counts are independent from predicted pagination and provide readback evidence.
        return PaginationEvidence(
            len(vertical) + 1, len(horizontal) + 1,
            horizontal, vertical, manual_horizontal, manual_vertical,
        )
