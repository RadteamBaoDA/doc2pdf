import dataclasses
import fnmatch
import warnings
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Literal, Optional

import yaml

_CONFIG_PATH: Path = Path("config.yml")

def set_config_path(path: Path) -> None:
    """Set the global configuration file path."""
    global _CONFIG_PATH
    _CONFIG_PATH = path

def get_config_path() -> Path:
    """Get the current configuration file path."""
    return _CONFIG_PATH



FileType = Literal["word", "excel", "powerpoint"]


@dataclass
class TimeoutSettings:
    """Settings for operation timeouts."""
    document_parsing: Optional[int] = 3600  # seconds (1 hour default)
    excel_trim: Optional[int] = 3600  # seconds (1 hour default)

    def __post_init__(self) -> None:
        for name in ("document_parsing", "excel_trim"):
            value = getattr(self, name)
            if value is not None and (isinstance(value, bool) or value <= 0):
                raise ValueError(f"timeouts.{name} must be null or > 0")
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "TimeoutSettings":
        """Create TimeoutSettings from dictionary."""
        return cls(
            document_parsing=data.get("document_parsing", 3600),
            excel_trim=data.get("excel_trim", 3600)
        )


@dataclass
class ParallelSettings:
    """Bounded concurrency settings for batch conversion."""

    excel_workers: int | Literal["auto"] = "auto"
    excel_worker_cap: int = 4

    def __post_init__(self) -> None:
        if self.excel_workers != "auto" and (
            isinstance(self.excel_workers, bool)
            or not isinstance(self.excel_workers, int)
            or not 1 <= self.excel_workers <= 8
        ):
            raise ValueError(
                "parallel.excel_workers must be 'auto' or an integer between 1 and 8"
            )
        if (
            isinstance(self.excel_worker_cap, bool)
            or not isinstance(self.excel_worker_cap, int)
            or not 1 <= self.excel_worker_cap <= 8
        ):
            raise ValueError(
                "parallel.excel_worker_cap must be an integer between 1 and 8"
            )

    def resolve_excel_workers(
        self,
        file_count: int,
        *,
        logical_cpus: Optional[int],
        available_memory_mb: Optional[int],
    ) -> int:
        """Resolve adaptive Excel concurrency without overcommitting the host."""
        if file_count <= 0:
            return 0
        if isinstance(self.excel_workers, int):
            return min(file_count, self.excel_workers)
        cpu_limit = max(1, (logical_cpus or 2) // 2)
        if available_memory_mb is None:
            memory_limit = 2
        else:
            memory_limit = max(1, (available_memory_mb - 2048) // 1536)
        return max(
            1,
            min(file_count, self.excel_worker_cap, cpu_limit, memory_limit),
        )

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "ParallelSettings":
        return cls(
            excel_workers=data.get("excel_workers", "auto"),
            excel_worker_cap=data.get("excel_worker_cap", 4),
        )


@dataclass
class TrimWhitespaceSettings:
    """Settings for PDF whitespace trimming."""
    enabled: bool = False
    margin: float = 10.0  # Points (1/72 inch) of padding around content
    include: List[str] = field(default_factory=lambda: ["word", "excel", "powerpoint"])
    box_mode: str = "physical"
    render_dpi: int = 72
    max_render_pixels: int = 20_000_000
    background_tolerance: int = 8
    include_annotations: bool = True
    allow_signature_invalidation: bool = False

    def __post_init__(self) -> None:
        if isinstance(self.margin, bool) or self.margin < 0:
            raise ValueError("trim_whitespace.margin must be >= 0")
        if self.box_mode not in {"physical", "cropbox"}:
            raise ValueError("trim_whitespace.box_mode must be physical or cropbox")
        if not 18 <= self.render_dpi <= 600:
            raise ValueError("trim_whitespace.render_dpi must be between 18 and 600")
        if self.max_render_pixels <= 0:
            raise ValueError("trim_whitespace.max_render_pixels must be > 0")
        if not 0 <= self.background_tolerance <= 255:
            raise ValueError("trim_whitespace.background_tolerance must be within 0..255")
        valid_types = {"word", "excel", "powerpoint", "pdf", "ocr"}
        if not isinstance(self.include, list) or not set(self.include) <= valid_types:
            raise ValueError("trim_whitespace.include contains an invalid file type")


@dataclass
class PostProcessingSettings:
    """PDF post-processing settings."""
    trim_whitespace: TrimWhitespaceSettings = field(default_factory=TrimWhitespaceSettings)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "PostProcessingSettings":
        """Create PostProcessingSettings from dictionary."""
        trim_data = data.get("trim_whitespace", {})
        return cls(
            trim_whitespace=TrimWhitespaceSettings(**trim_data) if trim_data else TrimWhitespaceSettings()
        )


@dataclass
class LayoutSettings:
    orientation: str = "portrait"
    pages_per_sheet: int = 1
    margins: str = "normal"

@dataclass
class MetadataSettings:
    include_properties: bool = True
    include_tags: bool = True

@dataclass
class OptimizationSettings:
    image_quality: str = "high"
    bitmap_text: bool = False

@dataclass
class PowerPointSettings:
    """PowerPoint-specific PDF conversion settings."""
    color_mode: str = "color"  # color, grayscale, bw
    slide_from: Optional[int] = None  # For range scope
    slide_to: Optional[int] = None

@dataclass
class ExcelSettings:
    """Effective Excel conversion policy.

    ``source_values`` retains only values explicitly supplied by configuration.
    It is deliberately carried with the effective object so sheet rules can be
    merged before profile defaults are expanded again.
    """
    quality_profile: str = "strict"  # strict, balanced, legacy
    layout_policy: Optional[str] = None
    page_size_scope: Optional[str] = None
    sheet_name: Optional[str] = None  # Target specific sheet, None = all visible sheets
    orientation: str = "auto"  # auto, portrait, landscape
    row_dimensions: Optional[int] = None  # None=auto, 0=try whole sheet, N=max rows per chunk
    metadata_header: bool = True  # Print header: sheet name | row range | filename
    min_shrink_factor: float = 0.90  # Minimum effective 2D scale (default 0.90 = 90%)
    ocr_sheet_name_label: bool = False  # Insert sheet name as large text in row 1 for OCR
    is_write_file_path: bool = False  # Insert file path row before last row
    oversized_action: str = "paginate"  # paginate, error, skip, or warn
    horizontal_overflow_strategy: str = "paginate"
    print_area_policy: Optional[str] = None  # profile-specific
    manual_page_break_policy: str = "preserve"
    print_title_rows: Optional[str] = None  # None preserves the workbook setting
    print_title_columns: Optional[str] = None  # None preserves the workbook setting
    preferred_papers: List[str] = field(default_factory=lambda: ["A4", "A3"])
    allowed_papers: Optional[List[str]] = None
    max_page_dimension_in: float = 24.0
    max_page_area_in2: float = 300.0
    avoid_horizontal_pagination: bool = True
    min_effective_font_pt: float = 10.0
    min_effective_image_dpi: float = 150.0
    print_quality: str = "standard"
    draft_mode: bool = False
    color_policy: str = "preserve"
    metadata_header_policy: str = "preserve"
    calculation_policy: str = "saved_cache"
    external_link_policy: str = "never_refresh"
    printer_policy: Optional[str] = None
    printer_name: str = "Microsoft Print to PDF"
    postflight_policy: Optional[str] = None
    trim_policy: Optional[str] = None
    # The extension provider fields are intentionally unavailable in M0-M5.
    vector_stitch_provider: Optional[str] = None
    pdfa_provider: Optional[str] = None
    source_values: Dict[str, Any] = field(default_factory=dict, repr=False, compare=False)

    def __post_init__(self) -> None:
        # Optional policy fields make direct ``ExcelSettings(quality_profile=...)``
        # select the same defaults as YAML profile expansion while retaining the
        # ability to reject an explicitly incompatible value.
        profile_seed = EXCEL_PROFILE_DEFAULTS.get(self.quality_profile, {})
        for name in (
            "layout_policy", "page_size_scope", "print_area_policy",
            "printer_policy", "postflight_policy", "trim_policy",
        ):
            if getattr(self, name) is None:
                setattr(self, name, profile_seed.get(name))
        enum_values = {
            "quality_profile": {"strict", "balanced", "legacy"},
            "layout_policy": {"preserve_authored", "optimize_missing", "force_optimize"},
            "page_size_scope": {"sheet", "chunk"},
            "orientation": {"portrait", "landscape", "auto"},
            "oversized_action": {"paginate", "error", "skip", "warn"},
            "horizontal_overflow_strategy": {
                "paginate", "error", "one_logical_page", "vector_stitch"
            },
            "print_area_policy": {
                "preserve", "preserve_strict", "expand_visible_objects", "auto"
            },
            "manual_page_break_policy": {"preserve", "reset"},
            "print_quality": {"standard"},
            "color_policy": {"preserve", "force_color", "black_and_white"},
            "metadata_header_policy": {"preserve", "append", "replace"},
            "calculation_policy": {"saved_cache", "calculate", "full_rebuild"},
            "external_link_policy": {"never_refresh", "refresh_allowed"},
            "printer_policy": {"required", "configured_fallback", "system_default"},
            "postflight_policy": {"strict", "warn", "disabled"},
            "trim_policy": {"disabled", "cropbox", "physical"},
        }
        for name, allowed in enum_values.items():
            if getattr(self, name) not in allowed:
                choices = ", ".join(sorted(allowed))
                raise ValueError(f"excel.{name} must be one of: {choices}")
        if self.orientation not in {"portrait", "landscape", "auto"}:
            raise ValueError("excel.orientation must be portrait, landscape, or auto")
        if self.row_dimensions is not None and (
            isinstance(self.row_dimensions, bool)
            or not isinstance(self.row_dimensions, int)
            or self.row_dimensions < 0
        ):
            raise ValueError("excel.row_dimensions must be null or an integer >= 0")
        if (
            isinstance(self.min_shrink_factor, bool)
            or not isinstance(self.min_shrink_factor, (int, float))
            or not 0 < float(self.min_shrink_factor) <= 1
        ):
            raise ValueError("excel.min_shrink_factor must be within (0, 1]")
        for name in ("print_title_rows", "print_title_columns"):
            value = getattr(self, name)
            if value is not None and (not isinstance(value, str) or not value.strip()):
                raise ValueError(f"excel.{name} must be null or a non-empty A1-style range")
        for name in (
            "metadata_header", "ocr_sheet_name_label", "is_write_file_path",
            "avoid_horizontal_pagination", "draft_mode",
        ):
            if not isinstance(getattr(self, name), bool):
                raise ValueError(f"excel.{name} must be a boolean")
        for name in (
            "max_page_dimension_in", "max_page_area_in2",
            "min_effective_font_pt", "min_effective_image_dpi",
        ):
            value = getattr(self, name)
            if isinstance(value, bool) or not isinstance(value, (int, float)) or value <= 0:
                raise ValueError(f"excel.{name} must be > 0")
        for name in ("preferred_papers", "allowed_papers"):
            value = getattr(self, name)
            if value is None and name == "allowed_papers":
                continue
            if not isinstance(value, list) or not value or any(
                not isinstance(item, str) or not item.strip() for item in value
            ):
                raise ValueError(f"excel.{name} must be a non-empty list of paper names")
        if not isinstance(self.printer_name, str) or not self.printer_name.strip():
            raise ValueError("excel.printer_name must be a non-empty string")
        if self.quality_profile == "strict":
            if self.postflight_policy == "disabled":
                raise ValueError("strict Excel profile cannot disable postflight")
            if self.external_link_policy == "refresh_allowed":
                raise ValueError("strict Excel profile cannot refresh external links")
            if self.draft_mode or self.print_quality != "standard":
                raise ValueError("strict Excel profile requires standard quality and Draft=False")
        if self.quality_profile != "legacy" and self.page_size_scope == "chunk":
            raise ValueError(
                "excel.page_size_scope=chunk is available only in the legacy profile"
            )
        if self.quality_profile != "legacy" and self.postflight_policy == "disabled":
            raise ValueError(
                "excel.postflight_policy=disabled is available only in legacy"
            )


EXCEL_PROFILE_DEFAULTS: Dict[str, Dict[str, Any]] = {
    "strict": {
        "quality_profile": "strict",
        "layout_policy": "preserve_authored",
        "page_size_scope": "sheet",
        "orientation": "auto",
        "min_shrink_factor": 0.90,
        "oversized_action": "paginate",
        "horizontal_overflow_strategy": "paginate",
        "row_dimensions": None,
        "print_area_policy": "preserve_strict",
        "manual_page_break_policy": "preserve",
        "preferred_papers": ["A4", "A3"],
        "allowed_papers": None,
        "max_page_dimension_in": 24.0,
        "max_page_area_in2": 300.0,
        "avoid_horizontal_pagination": True,
        "min_effective_font_pt": 10.0,
        "min_effective_image_dpi": 150.0,
        "print_quality": "standard",
        "draft_mode": False,
        "color_policy": "preserve",
        "metadata_header_policy": "preserve",
        "metadata_header": True,
        "ocr_sheet_name_label": False,
        "is_write_file_path": False,
        "calculation_policy": "saved_cache",
        "external_link_policy": "never_refresh",
        "printer_policy": "required",
        "printer_name": "Microsoft Print to PDF",
        "postflight_policy": "strict",
        "trim_policy": "disabled",
    },
    "balanced": {
        "quality_profile": "balanced",
        "layout_policy": "preserve_authored",
        "page_size_scope": "sheet",
        "print_area_policy": "expand_visible_objects",
        "postflight_policy": "warn",
        "trim_policy": "cropbox",
        "printer_policy": "configured_fallback",
    },
    "legacy": {
        "quality_profile": "legacy",
        "layout_policy": "optimize_missing",
        "page_size_scope": "chunk",
        "print_area_policy": "preserve",
        "postflight_policy": "disabled",
        "trim_policy": "disabled",
        "printer_policy": "system_default",
    },
}
EXCEL_PROFILE_DEFAULTS["balanced"] = {
    **EXCEL_PROFILE_DEFAULTS["strict"],
    **EXCEL_PROFILE_DEFAULTS["balanced"],
}


def _excel_settings_from_mapping(data: Dict[str, Any]) -> ExcelSettings:
    """Expand an Excel profile, preserving the supplied keys for later merges."""
    supplied = dict(data)
    nested_source = supplied.pop("source_values", None)
    if isinstance(nested_source, dict):
        # This path is used when reconstructing an already-resolved settings object.
        supplied = _merge_dict(nested_source, supplied)
    profile = supplied.get("quality_profile", "strict")
    if profile not in EXCEL_PROFILE_DEFAULTS:
        raise ValueError("excel.quality_profile must be strict, balanced, or legacy")
    effective = _merge_dict(EXCEL_PROFILE_DEFAULTS[profile], supplied)
    effective["source_values"] = dict(supplied)
    return ExcelSettings(**effective)
    
@dataclass
class SummaryReportSettings:
    """Summary report settings."""
    enabled: bool = True
    format: str = "summary_{timestamp}.txt"

@dataclass
class ErrorLogSettings:
    """Error log file settings."""
    enabled: bool = True
    format: str = "error_{timestamp}.txt"

@dataclass
class CopyErrorFilesSettings:
    """Settings for copying failed source files."""
    enabled: bool = True
    target_dir: str = "errors"

@dataclass
class ReportingSettings:
    """Reporting and error handling settings."""
    enabled: bool = True
    reports_dir: str = "reports"
    summary: SummaryReportSettings = field(default_factory=SummaryReportSettings)
    error_log: ErrorLogSettings = field(default_factory=ErrorLogSettings)
    copy_error_files: CopyErrorFilesSettings = field(default_factory=CopyErrorFilesSettings)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "ReportingSettings":
        """Create ReportingSettings from dictionary."""
        summary_data = data.get("summary", {})
        error_log_data = data.get("error_log", {})
        copy_error_data = data.get("copy_error_files", {})
        
        return cls(
            enabled=data.get("enabled", True),
            reports_dir=data.get("reports_dir", "reports"),
            summary=SummaryReportSettings(**summary_data) if summary_data else SummaryReportSettings(),
            error_log=ErrorLogSettings(**error_log_data) if error_log_data else ErrorLogSettings(),
            copy_error_files=CopyErrorFilesSettings(**copy_error_data) if copy_error_data else CopyErrorFilesSettings(),
        )

@dataclass
class PdfHandlingSettings:
    """Settings for handling existing PDF files."""
    copy_to_output: bool = False

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "PdfHandlingSettings":
        return cls(
            copy_to_output=data.get("copy_to_output", False)
        )


@dataclass
class PDFConversionSettings:
    scope: str = "all"
    layout: LayoutSettings = field(default_factory=LayoutSettings)
    metadata: MetadataSettings = field(default_factory=MetadataSettings)
    bookmarks: str = "headings"
    compliance: str = "standard"
    optimization: OptimizationSettings = field(default_factory=OptimizationSettings)
    powerpoint: Optional[PowerPointSettings] = None
    excel: Optional[ExcelSettings] = None

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "PDFConversionSettings":
        """Recursively create settings from dictionary."""
        layout_data = data.get("layout", {})
        metadata_data = data.get("metadata", {})
        opt_data = data.get("optimization", {})
        ppt_data = data.get("powerpoint", {})
        
        # Excel settings can be in nested 'excel' key OR at top level
        excel_data = dict(data.get("excel", {}) or {})
        # Also check for top-level excel settings (flat structure)
        top_level_excel_keys = [
            field_info.name for field_info in dataclasses.fields(ExcelSettings)
            if field_info.name != "source_values"
        ] + ["page_shrink_threshold"]
        for key in top_level_excel_keys:
            if key in data:
                # Top-level (flat) settings override nested 'excel' settings
                excel_data[key] = data[key]
        if "page_shrink_threshold" in excel_data:
            warnings.warn(
                "excel.page_shrink_threshold is deprecated and ignored; remove it "
                "from the configuration",
                FutureWarning,
                stacklevel=2,
            )
            excel_data.pop("page_shrink_threshold")
        
        return cls(
            scope=data.get("scope", "all"),
            layout=LayoutSettings(**layout_data) if layout_data else LayoutSettings(),
            metadata=MetadataSettings(**metadata_data) if metadata_data else MetadataSettings(),
            bookmarks=data.get("bookmarks", "headings"),
            compliance=data.get("compliance", "standard"),
            optimization=OptimizationSettings(**opt_data) if opt_data else OptimizationSettings(),
            powerpoint=PowerPointSettings(**ppt_data) if ppt_data else None,
            excel=_excel_settings_from_mapping(excel_data) if excel_data else None,
        )

def load_config(path: Optional[Path] = None) -> Dict[str, Any]:
    """
    Load configuration from a YAML file.
    """
    if path is None:
        path = _CONFIG_PATH

    if not path.exists():
        return {}
        
    try:
        with open(path, "r", encoding="utf-8") as f:
            return yaml.safe_load(f) or {}
    except yaml.YAMLError as e:
        raise ValueError(f"Malformed YAML configuration {path}: {e}") from e
    except OSError as e:
        raise ValueError(f"Cannot read configuration {path}: {e}") from e

def get_logging_config() -> Dict[str, Any]:
    """
    Get logging configuration with defaults.
    """
    config = load_config()
    logging_config = config.get("logging", {})
    
    # Defaults
    defaults = {
        "level": "INFO",
        "console": True,
        "file": {
            "enabled": True,
            "path": "logs/doc2pdf_{time:YYYYMMDDHHmmss}.log",
            "rotation": "10 MB",
            "retention": "10 days"
        }
    }
    
    # Merge defaults
    merged = defaults.copy()
    merged.update(logging_config)
    
    if "file" in logging_config:
        file_defaults = defaults["file"]
        file_config = logging_config["file"]
        merged_file = file_defaults.copy()
        merged_file.update(file_config)
        merged["file"] = merged_file
        
    return merged


def get_suffix_config() -> Dict[str, str]:
    """
    Get PDF filename suffix configuration per document type.
    
    Returns:
        Dict with keys 'word', 'powerpoint', 'excel' and their suffix values.
    """
    config = load_config()
    suffix_config = config.get("suffix", {})
    
    # Defaults (empty suffix)
    defaults = {
        "word": "",
        "powerpoint": "",
        "excel": ""
    }
    
    defaults.update(suffix_config)
    return defaults

def get_reporting_config() -> ReportingSettings:
    """
    Get reporting configuration with defaults.
    
    Returns:
        ReportingSettings object with summary, error_log, and copy_error_files settings.
    """
    config = load_config()
    reporting_data = config.get("reporting", {})
    
    if not reporting_data:
        return ReportingSettings()
    
    return ReportingSettings.from_dict(reporting_data)


def get_post_processing_config() -> PostProcessingSettings:
    """
    Get post-processing configuration with defaults.
    
    Returns:
        PostProcessingSettings object with trim_whitespace settings.
    """
    config = load_config()
    post_proc_data = config.get("post_processing", {})
    
    if not post_proc_data:
        return PostProcessingSettings()
    
    return PostProcessingSettings.from_dict(post_proc_data)


def get_pdf_handling_config() -> PdfHandlingSettings:
    """
    Get PDF handling configuration with defaults.
    """
    config = load_config()
    pdf_handling_data = config.get("pdf_handling", {})
    
    if not pdf_handling_data:
        return PdfHandlingSettings()
    
    return PdfHandlingSettings.from_dict(pdf_handling_data)


def get_timeout_config() -> TimeoutSettings:
    """
    Get timeout configuration with defaults.
    
    Returns:
        TimeoutSettings object with document_parsing and excel_trim timeout values.
    """
    config = load_config()
    timeout_data = config.get("timeout", {})
    
    if not timeout_data:
        return TimeoutSettings()
    
    return TimeoutSettings.from_dict(timeout_data)


def get_parallel_config() -> ParallelSettings:
    """Get bounded batch-concurrency settings with validated defaults."""
    config = load_config()
    parallel_data = config.get("parallel", {})
    if not isinstance(parallel_data, dict):
        raise ValueError("parallel must be a mapping")
    return ParallelSettings.from_dict(parallel_data)


def _merge_dict(base: Dict[str, Any], update: Dict[str, Any]) -> Dict[str, Any]:
    """Deep merge two dictionaries."""
    merged = base.copy()
    for key, value in update.items():
        if isinstance(value, dict) and key in merged and isinstance(merged[key], dict):
            merged[key] = _merge_dict(merged[key], value)
        else:
            merged[key] = value
    return merged

def get_pdf_settings(input_path: Path, file_type: str, base_path: Optional[Path] = None) -> PDFConversionSettings:
    """
    Get PDF settings by applying Pattern-Priority rules.
    
    1. Fetch list of rules for `file_type`.
    2. Filter rules where `input_path` matches `pattern`.
       Matches against:
       - Relative path (if base_path provided)
       - Full absolute path
       - Filename only
    3. Sort matching rules by `priority` (ascending).
    4. Merge settings sequentially.
    
    Args:
        input_path: The file path to check against rule patterns.
        file_type: The type of document ("word", "excel", "powerpoint").
        base_path: Optional root directory to calculate relative paths for matching.
    """
    config = load_config()
    pdf_section = config.get("pdf_settings", {})
    
    # Get rules list for the specific file type
    rules = pdf_section.get(file_type, [])
    if not isinstance(rules, list):
         print(f"Warning: Config for {file_type} is not a list of rules. Using defaults.")
         return PDFConversionSettings()

    # Determine paths for matching
    path_str = input_path.as_posix() if input_path else ""
    rel_path_str = ""
    if input_path and base_path:
        try:
            # Ensure both are absolute for reliable relative_to
            abs_input = input_path.resolve()
            abs_base = base_path.resolve()
            rel_path = abs_input.relative_to(abs_base)
            rel_path_str = rel_path.as_posix()
        except ValueError:
            # Not under base_path
            rel_path_str = ""
    
    # Filter matching rules
    matching_rules = []
    for rule in rules:
        pattern = rule.get("pattern", "*") # Default to match all if missing
        
        if pattern == "*":
            matching_rules.append(rule)
            continue

        # Try to match against relative path, absolute path, or filename
        matches = False
        if rel_path_str:
            matches = fnmatch.fnmatch(rel_path_str, pattern) or fnmatch.fnmatch(rel_path_str, f"*{pattern}*")
        
        if not matches and path_str:
            matches = fnmatch.fnmatch(path_str, pattern) or fnmatch.fnmatch(path_str, f"*{pattern}*")
            
        if not matches and input_path:
            matches = fnmatch.fnmatch(input_path.name, pattern)
            
        if matches:
            matching_rules.append(rule)
    
    # Sort by priority (Ascending means later overrides earlier? "Higher numbers override lower numbers" -> implies Sort Ascending)
    # If Priority 10 is default, Priority 100 is override. 
    # We want to apply 10 THEN 100. So 100 overwrites 10.
    # So we sort by priority ASCENDING.
    matching_rules.sort(key=lambda x: x.get("priority", 0))
    
    # Merge settings
    final_settings_dict = {}
    for rule in matching_rules:
        # We need a base to merge into. 
        # If the first rule is partial, we might need a hardcoded "system default".
        # But usually the "*" rule with priority 10 acts as the base.
        # We'll rely on the user config providing a base rule. 
        # But we should start with empty dict or Pydantic defaults?
        # Pydantic defaults are 'good', but nested objects might need care.
        # Let's start with empty dict and trust _merge_dict + from_dict to handle missing keys by using dataclass defaults.
        rule_settings = rule.get("settings", {})
        final_settings_dict = _merge_dict(final_settings_dict, rule_settings)
        
    return PDFConversionSettings.from_dict(final_settings_dict)


def get_excel_sheet_settings(sheet_name: str, base_settings: Optional[PDFConversionSettings] = None, input_path: Optional[Path] = None, base_path: Optional[Path] = None) -> PDFConversionSettings:
    """
    Get Excel PDF settings by applying sheet_name-based Pattern-Priority rules.
    
    For Excel, rules are matched by sheet_name pattern instead of file path.
    
    1. Fetch list of rules for 'excel'.
    2. Filter rules where `sheet_name` matches `sheet_name` pattern.
    3. Sort matching rules by `priority` (ascending).
    4. Merge settings sequentially.
    
    Args:
        sheet_name: The Excel sheet name to check against rule patterns.
        base_settings: Optional base settings to merge into.
        input_path: Optional file path to check against rule patterns.
        base_path: Optional root directory to calculate relative paths for matching.
    
    Returns:
        PDFConversionSettings with merged sheet-specific settings.
    """
    config = load_config()
    pdf_section = config.get("pdf_settings", {})
    
    rules = pdf_section.get("excel", [])
    if not isinstance(rules, list):
        return base_settings or PDFConversionSettings()

    # Filter matching rules
    matching_rules = []
    
    # Determine path details for matching
    path_str = input_path.as_posix() if input_path else ""
    rel_path_str = ""
    if input_path and base_path:
        try:
            abs_input = input_path.resolve()
            abs_base = base_path.resolve()
            rel_path = abs_input.relative_to(abs_base)
            rel_path_str = rel_path.as_posix()
        except ValueError:
            rel_path_str = ""

    for rule in rules:
        # Check Sheet Name Pattern
        sheet_pattern = rule.get("sheet_name", "*") 
        
        # Check File Path Pattern
        file_pattern = rule.get("pattern", "*")
        
        # 1. Check Sheet Name Match
        sheet_match = fnmatch.fnmatch(sheet_name, sheet_pattern)
        
        # 2. Check File Path Match
        file_match = True
        if input_path:
             if file_pattern != "*":
                # Match against relative path, absolute path, or filename
                file_match = False
                if rel_path_str:
                    file_match = fnmatch.fnmatch(rel_path_str, file_pattern) or fnmatch.fnmatch(rel_path_str, f"*{file_pattern}*")
                
                if not file_match:
                    file_match = fnmatch.fnmatch(path_str, file_pattern) or fnmatch.fnmatch(path_str, f"*{file_pattern}*")
                
                if not file_match:
                    file_match = fnmatch.fnmatch(input_path.name, file_pattern)
        else:
             # If no input path provided, only match if pattern is universal
             if file_pattern != "*":
                 file_match = False

        if sheet_match and file_match:
            matching_rules.append(rule)
    
    # Sort by priority ascending (higher priority overrides)
    matching_rules.sort(key=lambda x: x.get("priority", 0))
    
    # Start with base settings dict or empty
    if base_settings:
        final_settings_dict = asdict(base_settings)
        # Profile expansion must happen after sheet rules.  Reuse the original
        # Excel keys rather than feeding every effective default back as if it
        # had been explicitly configured.
        if base_settings.excel is not None:
            raw_excel = base_settings.excel.source_values
            if not raw_excel:
                raw_excel = {
                    key: value for key, value in asdict(base_settings.excel).items()
                    if key != "source_values"
                }
            final_settings_dict["excel"] = dict(raw_excel)
    else:
        final_settings_dict = {}
    
    # Merge matching rules
    for rule in matching_rules:
        rule_settings = rule.get("settings", {})
        final_settings_dict = _merge_dict(final_settings_dict, rule_settings)
        
    return PDFConversionSettings.from_dict(final_settings_dict)
