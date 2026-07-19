"""Deterministic quality-first candidate validation and scoring."""

from __future__ import annotations

import math
from typing import Iterable, Optional, Tuple

from .models import LayoutConstraints, QualityLayoutCandidate


def evaluate_candidate(
    candidate: QualityLayoutCandidate,
    constraints: LayoutConstraints,
) -> QualityLayoutCandidate:
    """Document this Excel pipeline operation and its side effects."""
    reasons = list(candidate.rejection_reasons)
    # Reject impossible geometry before comparing candidates; clamping hides driver errors.
    width = candidate.usable_width_inches
    height = candidate.usable_height_inches
    if not all(math.isfinite(value) and value > 0 for value in (width, height)):
        reasons.append("invalid usable geometry")
    if max(width, height) > constraints.max_page_dimension_in:
        reasons.append("maximum page dimension exceeded")
    if width * height > constraints.max_page_area_in2:
        reasons.append("maximum page area exceeded")
    if candidate.zoom < constraints.quality_zoom:
        reasons.append("quality scale floor violated")
    if (
        candidate.effective_font_pt is not None
        and candidate.effective_font_pt < constraints.min_font_pt
    ):
        reasons.append("minimum effective font size violated")
    if (
        candidate.effective_image_dpi is not None
        and candidate.effective_image_dpi < constraints.min_image_dpi
    ):
        reasons.append("minimum effective image DPI violated")
    return QualityLayoutCandidate(
        **{
            **candidate.__dict__,
            "rejection_reasons": tuple(dict.fromkeys(reasons)),
        }
    )


def candidate_sort_key(candidate: QualityLayoutCandidate) -> Tuple[object, ...]:
    """Document this Excel pipeline operation and its side effects."""
    return (
        candidate.pages_wide,
        candidate.pages_wide * candidate.pages_tall,
        0 if candidate.repeated_titles else 1,
        candidate.preferred_rank,
        candidate.whitespace_area_in2,
        candidate.usable_width_inches * candidate.usable_height_inches,
        candidate.paper_name.casefold(),
        candidate.orientation,
        candidate.paper_enum,
    )


def choose_candidate(
    candidates: Iterable[QualityLayoutCandidate],
    constraints: LayoutConstraints,
) -> Tuple[Optional[QualityLayoutCandidate], Tuple[QualityLayoutCandidate, ...]]:
    """Document this Excel pipeline operation and its side effects."""
    evaluated = tuple(evaluate_candidate(item, constraints) for item in candidates)
    # Keep rejected candidates as evidence so callers can explain the decision.
    accepted = [item for item in evaluated if item.accepted]
    chosen = min(accepted, key=candidate_sort_key) if accepted else None
    rejected = tuple(item for item in evaluated if not item.accepted)
    return chosen, rejected
