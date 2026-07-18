import yaml
from pathlib import Path
from typing import Any, Dict, Optional, List, Literal
from dataclasses import dataclass, field, asdict
import fnmatch
import warnings

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
    """Excel-specific PDF conversion settings for OCR-optimized output."""
    sheet_name: Optional[str] = None  # Target specific sheet, None = all visible sheets
    orientation: str = "auto"  # auto, portrait, landscape
    row_dimensions: Optional[int] = None  # None=auto, 0=try whole sheet, N=max rows per chunk
    metadata_header: bool = True  # Print header: sheet name | row range | filename
    min_shrink_factor: float = 0.90  # Minimum effective 2D scale (default 0.90 = 90%)
    ocr_sheet_name_label: bool = False  # Insert sheet name as large text in row 1 for OCR
    is_write_file_path: bool = False  # Insert file path row before last row
    oversized_action: str = "paginate"  # paginate, error, skip, or warn
    print_area_policy: str = "preserve"  # preserve existing areas, or auto-detect
    print_title_rows: Optional[str] = None  # None preserves the workbook setting
    print_title_columns: Optional[str] = None  # None preserves the workbook setting

    def __post_init__(self) -> None:
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
        if self.oversized_action not in {"paginate", "error", "skip", "warn"}:
            raise ValueError("excel.oversized_action must be paginate, error, skip, or warn")
        if self.print_area_policy not in {"preserve", "auto"}:
            raise ValueError("excel.print_area_policy must be preserve or auto")
        for name in ("print_title_rows", "print_title_columns"):
            value = getattr(self, name)
            if value is not None and (not isinstance(value, str) or not value.strip()):
                raise ValueError(f"excel.{name} must be null or a non-empty A1-style range")
    
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
    compliance: str = "pdfa"
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
        top_level_excel_keys = ["orientation", "row_dimensions", "metadata_header", "sheet_name", "min_shrink_factor", "ocr_sheet_name_label", "is_write_file_path", "oversized_action", "page_shrink_threshold", "print_area_policy", "print_title_rows", "print_title_columns"]
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
            compliance=data.get("compliance", "pdfa"),
            optimization=OptimizationSettings(**opt_data) if opt_data else OptimizationSettings(),
            powerpoint=PowerPointSettings(**ppt_data) if ppt_data else None,
            excel=ExcelSettings(**excel_data) if excel_data else None,
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
        # Preserve every current and future dataclass field during sheet overrides.
        final_settings_dict = asdict(base_settings)
    else:
        final_settings_dict = {}
    
    # Merge matching rules
    for rule in matching_rules:
        rule_settings = rule.get("settings", {})
        final_settings_dict = _merge_dict(final_settings_dict, rule_settings)
        
    return PDFConversionSettings.from_dict(final_settings_dict)
