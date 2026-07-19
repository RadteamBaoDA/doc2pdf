"""Pure safe-boundary planning for Excel logical chunks."""

from __future__ import annotations

from typing import Iterable, Set, Tuple

from .models import PrintableObject, ResolvedRegion, SafeChunk


class SafeChunkPlanner:
    @staticmethod
    def forbidden_row_boundaries(
        atomic_ranges: Iterable[Tuple[int, int]],
        objects: Iterable[PrintableObject] = (),
    ) -> Set[int]:
        """Document the behavior and contract of this Excel pipeline operation."""
        forbidden: Set[int] = set()
        # A boundary inside a merged/table/object span would orphan content.
        for first, last in atomic_ranges:
            forbidden.update(range(int(first), int(last)))
        for item in objects:
            forbidden.update(range(item.first_row, item.last_row))
        return forbidden

    def chunks(
        self,
        regions: Iterable[ResolvedRegion],
        row_limit: int | None,
        forbidden_boundaries: Set[int],
    ) -> Tuple[SafeChunk, ...]:
        """Document the behavior and contract of this Excel pipeline operation."""
        result = []
        order = 0
        for region in regions:
            if row_limit is None or row_limit == 0:
                result.append(SafeChunk(
                    order, region.order, region.first_row, region.last_row,
                    region.first_col, region.last_col,
                ))
                order += 1
                continue
            first = region.first_row
            while first <= region.last_row:
                # Start with the soft row limit, then move to a safe boundary.
                requested = min(region.last_row, first + row_limit - 1)
                boundary = requested
                moved = False
                if boundary < region.last_row and boundary in forbidden_boundaries:
                    lower = boundary
                    while lower >= first and lower in forbidden_boundaries:
                        lower -= 1
                    upper = boundary
                    while upper < region.last_row and upper in forbidden_boundaries:
                        upper += 1
                    if lower >= first:
                        # Prefer the nearest safe boundary that stays in this chunk.
                        boundary = lower
                    elif upper < region.last_row:
                        boundary = upper
                    else:
                        boundary = region.last_row
                    moved = boundary != requested
                result.append(SafeChunk(
                    order, region.order, first, boundary,
                    region.first_col, region.last_col, moved,
                ))
                order += 1
                first = boundary + 1
        return tuple(result)
