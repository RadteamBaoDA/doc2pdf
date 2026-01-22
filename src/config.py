import yaml
from pathlib import Path
from typing import Any, Dict, Optional, List, Literal
from dataclasses import dataclass, field, asdict
import fnmatch

CONFIG_FILE = Path("config.yml")

# Supported file types
FileType = Literal["word", "excel", "powerpoint"]


@dataclass
class TrimWhitespaceSettings:
    """Settings for PDF whitespace trimming."""
    enabled: bool = False
    margin: float = 10.0  # Points (1/72 inch) of padding around content
    include: List[str] = field(default_factory=lambda: ["word", "excel", "powerpoint"])


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
    orientation: str = "landscape"  # portrait, landscape
    row_dimensions: Optional[int] = None  # Rows per page: None=auto, 0=fit all on one page, N=fixed rows
    metadata_header: bool = True  # Print header: sheet name | row range | filename
    min_shrink_factor: float = 0.8  # Minimum allowed scaling factor before error (default 0.8 = 80%)

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
        excel_data = data.get("excel", {})
        # Also check for top-level excel settings (flat structure)
        top_level_excel_keys = ["orientation", "row_dimensions", "metadata_header", "sheet_name", "min_shrink_factor"]
        for key in top_level_excel_keys:
            if key in data:
                # Top-level (flat) settings override nested 'excel' settings
                excel_data[key] = data[key]
        
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

def load_config(path: Path = CONFIG_FILE) -> Dict[str, Any]:
    """
    Load configuration from a YAML file.
    """
    if not path.exists():
        return {}
        
    try:
        with open(path, "r", encoding="utf-8") as f:
            return yaml.safe_load(f) or {}
    except Exception as e:
        print(f"Warning: Failed to load configuration from {path}: {e}")
        return {}

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



def _merge_dict(base: Dict[str, Any], update: Dict[str, Any]) -> Dict[str, Any]:
    """Deep merge two dictionaries."""
    merged = base.copy()
    for key, value in update.items():
        if isinstance(value, dict) and key in merged and isinstance(merged[key], dict):
            merged[key] = _merge_dict(merged[key], value)
        else:
            merged[key] = value
    return merged

def get_pdf_settings(input_path: Optional[Path] = None, file_type: FileType = "word") -> PDFConversionSettings:
    """
    Get PDF settings by applying Pattern-Priority rules.
    
    1. Fetch list of rules for `file_type`.
    2. Filter rules where `input_path` matches `pattern`.
    3. Sort matching rules by `priority` (ascending).
    4. Merge settings sequentially.
    
    Args:
        input_path: The file path to check against rule patterns.
        file_type: The type of document ("word", "excel", "powerpoint").
    """
    config = load_config()
    pdf_section = config.get("pdf_settings", {})
    
    # Get rules list for the specific file type
    # If it's a dict (legacy/error fallback), treat as single item? No, enforcing list.
    rules = pdf_section.get(file_type, [])
    if not isinstance(rules, list):
         # Logic to handle if config is malformed or old structure: return default or try to parse
         print(f"Warning: Config for {file_type} is not a list of rules. Using defaults.")
         return PDFConversionSettings()

    # Determine path string for matching
    path_str = input_path.as_posix() if input_path else ""
    
    # Filter matching rules
    matching_rules = []
    for rule in rules:
        pattern = rule.get("pattern", "*") # Default to match all if missing
        priority = rule.get("priority", 0)
        
        # If input_path is None (e.g. unknown context), we only match "*"
        if input_path is None:
            if pattern == "*":
                matching_rules.append(rule)
        else:
            # Use Path.match for glob support including **
            # Ensure pattern is compatible with current OS path separator if needed, 
            # but usually Path.match handles forward slashes in pattern well.
            if input_path.match(pattern):
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


def get_excel_sheet_settings(sheet_name: str, base_settings: Optional[PDFConversionSettings] = None, input_path: Optional[Path] = None) -> PDFConversionSettings:
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
    
    # Determine path string for matching
    path_str = input_path.as_posix() if input_path else ""

    for rule in rules:
        # Check Sheet Name Pattern
        # Default to "*" (match all sheets) if missing, but usually Excel rules are sheet-based
        sheet_pattern = rule.get("sheet_name", "*") # Backward compat or default
        
        # Check File Path Pattern (New)
        # Default to "*" (match all files) if missing
        file_pattern = rule.get("pattern", "*")
        
        # 1. Check Sheet Name Match
        sheet_match = fnmatch.fnmatch(sheet_name, sheet_pattern)
        
        # 2. Check File Path Match
        file_match = True
        if input_path:
             if file_pattern != "*":
                file_match = input_path.match(file_pattern)
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
        # Convert base_settings to dict for merging
        from dataclasses import asdict
        final_settings_dict = {
            "scope": base_settings.scope,
            "bookmarks": base_settings.bookmarks,
            "compliance": base_settings.compliance,
            "layout": {
                "orientation": base_settings.layout.orientation,
                "pages_per_sheet": base_settings.layout.pages_per_sheet,
                "margins": base_settings.layout.margins,
            },
            "metadata": {
                "include_properties": base_settings.metadata.include_properties,
                "include_tags": base_settings.metadata.include_tags,
            },
            "optimization": {
                "image_quality": base_settings.optimization.image_quality,
                "bitmap_text": base_settings.optimization.bitmap_text,
            },
        }
        if base_settings.excel:
            final_settings_dict["excel"] = {
                "sheet_name": base_settings.excel.sheet_name,
                "orientation": base_settings.excel.orientation,
                "row_dimensions": base_settings.excel.row_dimensions,
                "metadata_header": base_settings.excel.metadata_header,
                "min_shrink_factor": base_settings.excel.min_shrink_factor,
            }
    else:
        final_settings_dict = {}
    
    # Merge matching rules
    for rule in matching_rules:
        rule_settings = rule.get("settings", {})
        final_settings_dict = _merge_dict(final_settings_dict, rule_settings)
        
    return PDFConversionSettings.from_dict(final_settings_dict)
