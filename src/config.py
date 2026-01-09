import yaml
from pathlib import Path
from typing import Any, Dict, Optional, List, Literal
from dataclasses import dataclass, field, asdict
import fnmatch

CONFIG_FILE = Path("config.yml")

# Supported file types
FileType = Literal["word", "excel", "powerpoint"]

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
class PDFConversionSettings:
    scope: str = "all"
    layout: LayoutSettings = field(default_factory=LayoutSettings)
    metadata: MetadataSettings = field(default_factory=MetadataSettings)
    bookmarks: str = "headings"
    compliance: str = "pdfa"
    optimization: OptimizationSettings = field(default_factory=OptimizationSettings)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "PDFConversionSettings":
        """Recursively create settings from dictionary."""
        layout_data = data.get("layout", {})
        metadata_data = data.get("metadata", {})
        opt_data = data.get("optimization", {})
        
        return cls(
            scope=data.get("scope", "all"),
            layout=LayoutSettings(**layout_data) if layout_data else LayoutSettings(),
            metadata=MetadataSettings(**metadata_data) if metadata_data else MetadataSettings(),
            bookmarks=data.get("bookmarks", "headings"),
            compliance=data.get("compliance", "pdfa"),
            optimization=OptimizationSettings(**opt_data) if opt_data else OptimizationSettings()
        )

def load_config(path: Path = CONFIG_FILE) -> Dict[str, Any]:
    """
    Load configuration from a YAML file.
    """
    if not path.exists():
        return {}
        
    try:
        with open(path, "r") as f:
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
            if fnmatch.fnmatch(path_str, pattern):
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
