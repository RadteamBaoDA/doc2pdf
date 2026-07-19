"""Pure-data contracts for the Excel quality-first conversion pipeline."""

from __future__ import annotations

import dataclasses
import hashlib
import json
import math
import os
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, Mapping, Optional, Tuple


SCHEMA_VERSION = 2


def _json_value(value: Any) -> Any:
    """Document this Excel pipeline operation and its side effects."""
    if dataclasses.is_dataclass(value):
        return {
            item.name: _json_value(getattr(value, item.name))
            for item in dataclasses.fields(value)
        }
    if isinstance(value, Mapping):
        return {str(key): _json_value(value[key]) for key in sorted(value)}
    if isinstance(value, (tuple, list, set, frozenset)):
        return [_json_value(item) for item in value]
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, float) and not math.isfinite(value):
        raise ValueError("decision evidence cannot contain a non-finite number")
    return value


def stable_id(value: Any, prefix: str = "exq") -> str:
    """Return a deterministic identifier for normalized decision inputs."""
    encoded = json.dumps(
        _json_value(value), sort_keys=True, ensure_ascii=False,
        separators=(",", ":"),
    ).encode("utf-8")
    return f"{prefix}-{hashlib.sha256(encoded).hexdigest()[:20]}"


@dataclass(frozen=True)
class WorkbookSheetInfo:
    index: int
    name: str
    kind: str
    visible: bool


@dataclass(frozen=True)
class AuthoredLayoutSnapshot:
    classification: str
    confidence: str
    source: str
    print_area: str = ""
    paper_size: Optional[int] = None
    orientation: Optional[int] = None
    margins_points: Tuple[float, ...] = ()
    zoom: Any = None
    fit_to_pages_wide: Any = None
    fit_to_pages_tall: Any = None
    print_title_rows: str = ""
    print_title_columns: str = ""
    page_order: Any = None
    manual_row_breaks: Tuple[int, ...] = ()
    manual_column_breaks: Tuple[int, ...] = ()
    headers: Tuple[str, ...] = ()
    footers: Tuple[str, ...] = ()
    black_and_white: Any = None
    draft: Any = None
    errors: Tuple[str, ...] = ()

    @property
    def fingerprint(self) -> str:
        """Document this Excel pipeline operation and its side effects."""
        return stable_id(self, "layout")


@dataclass(frozen=True)
class PrintableObject:
    identifier: str
    kind: str
    first_row: int
    first_col: int
    last_row: int
    last_col: int
    printable: bool = True


@dataclass(frozen=True)
class ResolvedRegion:
    order: int
    first_row: int
    first_col: int
    last_row: int
    last_col: int
    sources: Tuple[str, ...] = ()
    confidence: str = "certain"
    errors: Tuple[str, ...] = ()

    @property
    def is_empty(self) -> bool:
        """Document this Excel pipeline operation and its side effects."""
        return self.last_row < self.first_row or self.last_col < self.first_col


@dataclass(frozen=True)
class SafeChunk:
    order: int
    region_order: int
    first_row: int
    last_row: int
    first_col: int
    last_col: int
    moved_boundary: bool = False


@dataclass(frozen=True)
class PrinterFormCapability:
    paper_enum: int
    name: str
    width_inches: float
    height_inches: float
    imageable_width_inches: Optional[float] = None
    imageable_height_inches: Optional[float] = None
    hard_margins_inches: Tuple[float, float, float, float] = ()


@dataclass(frozen=True)
class PrinterCapability:
    name: str
    driver: str = ""
    driver_version: str = ""
    port: str = ""
    forms: Tuple[PrinterFormCapability, ...] = ()
    fallback: Optional[str] = None
    errors: Tuple[str, ...] = ()


@dataclass(frozen=True)
class LayoutConstraints:
    quality_zoom: int
    min_font_pt: float
    min_image_dpi: float
    max_page_dimension_in: float
    max_page_area_in2: float
    preferred_papers: Tuple[str, ...]


@dataclass(frozen=True)
class QualityLayoutCandidate:
    paper_enum: int
    paper_name: str
    orientation: int
    usable_width_inches: float
    usable_height_inches: float
    width_scale: float
    height_scale: float
    effective_scale: float
    zoom: int
    pages_wide: int
    pages_tall: int
    effective_font_pt: Optional[float] = None
    effective_image_dpi: Optional[float] = None
    whitespace_area_in2: float = 0.0
    preferred_rank: int = 1_000_000
    repeated_titles: bool = False
    rejection_reasons: Tuple[str, ...] = ()

    @property
    def accepted(self) -> bool:
        """Document this Excel pipeline operation and its side effects."""
        return not self.rejection_reasons


@dataclass(frozen=True)
class LayoutDecision:
    workbook: str
    sheet: str
    sheet_index: int
    mode: str
    region_ids: Tuple[str, ...]
    chosen: Optional[QualityLayoutCandidate]
    rejected: Tuple[QualityLayoutCandidate, ...] = ()
    predicted_grid: Tuple[int, int] = (0, 0)
    actual_grid: Tuple[int, int] = (0, 0)
    printer: Optional[PrinterCapability] = None
    authored_fingerprint: Optional[str] = None
    calculation_policy: str = "saved_cache"
    metadata_policy: str = "preserve"
    manual_break_policy: str = "preserve"
    warnings: Tuple[str, ...] = ()
    failures: Tuple[str, ...] = ()
    schema_version: int = SCHEMA_VERSION
    decision_id: str = field(default="", compare=False)

    def __post_init__(self) -> None:
        """Document this Excel pipeline operation and its side effects."""
        if not self.decision_id:
            values = {
                field_info.name: getattr(self, field_info.name)
                for field_info in dataclasses.fields(self)
                if field_info.name not in {"decision_id", "schema_version"}
            }
            object.__setattr__(self, "decision_id", stable_id(values))


@dataclass(frozen=True)
class ExportedSheetArtifact:
    decision_id: str
    sheet: str
    path: str
    first_page: int
    last_page: int


@dataclass(frozen=True)
class PdfPostflightResult:
    passed: bool
    page_count: int
    checks: Mapping[str, bool]
    failures: Tuple[str, ...] = ()
    warnings: Tuple[str, ...] = ()
    page_evidence: Tuple[Mapping[str, Any], ...] = ()


@dataclass(frozen=True)
class ConversionManifest:
    workbook: str
    output: str
    profile: str
    decisions: Tuple[LayoutDecision, ...]
    artifacts: Tuple[ExportedSheetArtifact, ...] = ()
    postflight: Optional[PdfPostflightResult] = None
    skipped_sheets: Tuple[str, ...] = ()
    failures: Tuple[str, ...] = ()
    timings_ms: Mapping[str, float] = field(default_factory=dict)
    runtime_evidence: Mapping[str, Any] = field(default_factory=dict)
    schema_version: int = SCHEMA_VERSION

    def to_dict(self) -> Dict[str, Any]:
        """Document this Excel pipeline operation and its side effects."""
        return _json_value(self)

    def write_atomic(self, path: Path) -> Path:
        """Document this Excel pipeline operation and its side effects."""
        path = path.resolve()
        path.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary_name = tempfile.mkstemp(
            prefix=f".{path.name}.", suffix=".tmp", dir=str(path.parent)
        )
        temporary = Path(temporary_name)
        try:
            with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as stream:
                json.dump(
                    self.to_dict(), stream, indent=2, sort_keys=True,
                    ensure_ascii=False,
                )
                stream.write("\n")
            os.replace(temporary, path)
        finally:
            temporary.unlink(missing_ok=True)
        return path


def manifest_name(workbook: Path, decisions: Iterable[LayoutDecision]) -> str:
    """Document this Excel pipeline operation and its side effects."""
    identity = stable_id(
        {"workbook": workbook.name, "decisions": tuple(decisions)}, "manifest"
    )
    return f"{workbook.stem}.{identity}.json"
