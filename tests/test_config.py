from pathlib import Path
from unittest.mock import patch

import pytest

from src.config import (
    ExcelSettings,
    ParallelSettings,
    PDFConversionSettings,
    TrimWhitespaceSettings,
    _merge_dict,
    get_excel_sheet_settings,
    get_parallel_config,
    get_pdf_handling_config,
    get_pdf_settings,
    load_config,
)

# Mock config data
MOCK_CONFIG = {
    "pdf_settings": {
        "word": [
            {
                "pattern": "*",
                "priority": 10,
                "settings": {
                    "scope": "all",
                    "layout": {"orientation": "portrait"},
                    "compliance": "pdfa"
                }
            },
            {
                "pattern": "**/CONFIDENTIAL/**",
                "priority": 100,
                "settings": {
                    "metadata": {"include_properties": False}
                }
            }
        ],
        "excel": [
            {
                "pattern": "*",
                "priority": 10,
                "settings": {
                    "layout": {"orientation": "landscape"}
                }
            }
        ]
    }
}

@pytest.fixture
def mock_load_config():
    with patch("src.config.load_config", return_value=MOCK_CONFIG) as mock:
        yield mock

def test_merge_dict():
    base = {"a": 1, "b": {"x": 10, "y": 20}}
    update = {"a": 2, "b": {"x": 30}, "c": 3}
    
    result = _merge_dict(base, update)
    
    assert result["a"] == 2
    assert result["b"]["x"] == 30
    assert result["b"]["y"] == 20
    assert result["c"] == 3
    # Base should not be mutated
    assert base["a"] == 1


def test_parallel_settings_default_to_adaptive_workers():
    settings = ParallelSettings()
    assert settings.excel_workers == "auto"
    assert settings.excel_worker_cap == 4


def test_parallel_settings_resolves_cpu_memory_and_fallback_limits():
    settings = ParallelSettings()
    assert settings.resolve_excel_workers(
        10, logical_cpus=16, available_memory_mb=16_384
    ) == 4
    assert settings.resolve_excel_workers(
        10, logical_cpus=16, available_memory_mb=3_000
    ) == 1
    assert settings.resolve_excel_workers(
        10, logical_cpus=16, available_memory_mb=None
    ) == 2


@pytest.mark.parametrize("workers", [1, 2, 8])
def test_parallel_settings_accept_supported_worker_counts(workers):
    assert ParallelSettings(excel_workers=workers).excel_workers == workers


@pytest.mark.parametrize("workers", [True, False, 0, 9, -1, 1.5, "2"])
def test_parallel_settings_reject_invalid_worker_counts(workers):
    with pytest.raises(ValueError, match="parallel.excel_workers"):
        ParallelSettings(excel_workers=workers)


def test_get_parallel_config_loads_explicit_serial_mode():
    with patch("src.config.load_config", return_value={"parallel": {"excel_workers": 1}}):
        assert get_parallel_config().excel_workers == 1


def test_get_parallel_config_rejects_non_mapping_section():
    with patch("src.config.load_config", return_value={"parallel": 2}):
        with pytest.raises(ValueError, match="parallel must be a mapping"):
            get_parallel_config()

def test_get_pdf_settings_default_word(mock_load_config):
    # Test getting defaults for a standard file
    settings = get_pdf_settings(input_path=Path("input/doc.docx"), file_type="word")
    
    assert settings.scope == "all"
    assert settings.layout.orientation == "portrait"
    assert settings.compliance == "pdfa"
    # Metadata should follow class defaults if not specified, 
    # but here defaults are True in dataclass.
    assert settings.metadata.include_properties is True


def test_programmatic_defaults_use_strict_excel_and_standard_pdf():
    assert PDFConversionSettings().compliance == "standard"
    assert ExcelSettings().quality_profile == "strict"

def test_get_pdf_settings_pattern_override(mock_load_config):
    # Test override logic
    settings = get_pdf_settings(input_path=Path("input/CONFIDENTIAL/secret.docx"), file_type="word")
    
    # Base settings should be present
    assert settings.scope == "all"
    assert settings.layout.orientation == "portrait"
    # Override should be applied
    assert settings.metadata.include_properties is False

def test_get_pdf_settings_excel(mock_load_config):
    settings = get_pdf_settings(input_path=Path("sheet.xlsx"), file_type="excel")
    assert settings.layout.orientation == "landscape"

def test_get_pdf_settings_no_match(mock_load_config):
    # If no pattern matches (which is hard with "*"), but let's assume empty config logic
    pass 

def test_priority_sorting(mock_load_config):
    # Create a complex scenario with 3 priorities
    complex_config = {
        "pdf_settings": {
            "word": [
                {"pattern": "*", "priority": 10, "settings": {"bookmarks": "none"}},
                {"pattern": "*important*", "priority": 50, "settings": {"bookmarks": "headings"}},
                {"pattern": "**/CLIENT/**", "priority": 100, "settings": {"bookmarks": "bookmarks"}}
            ]
        }
    }
    
    with patch("src.config.load_config", return_value=complex_config):
        # Case 1: Just *
        s1 = get_pdf_settings(Path("doc.docx"), "word")
        assert s1.bookmarks == "none"
        
        # Case 2: * + *important*
        s2 = get_pdf_settings(Path("very_important_doc.docx"), "word")
        assert s2.bookmarks == "headings"
        
        # Case 3: * + *important* + **/CLIENT/** (Highest priority wins)
        s3 = get_pdf_settings(Path("input/CLIENT/very_important_doc.docx"), "word")
        assert s3.bookmarks == "bookmarks"

def test_get_pdf_handling_config(mock_load_config):
    # Test loading PDF handling config
    with patch("src.config.load_config", return_value={"pdf_handling": {"copy_to_output": True}}):
        config = get_pdf_handling_config()
        assert config.copy_to_output is True

    # Test default
    with patch("src.config.load_config", return_value={}):
        config = get_pdf_handling_config()
        assert config.copy_to_output is False


def test_excel_settings_quality_first_defaults():
    settings = ExcelSettings()

    assert settings.quality_profile == "strict"
    assert settings.orientation == "auto"
    assert settings.min_shrink_factor == pytest.approx(0.90)
    assert settings.oversized_action == "paginate"
    assert settings.print_title_rows is None
    assert settings.print_title_columns is None


def test_excel_settings_accept_print_titles_and_paginate_action():
    settings = ExcelSettings(
        oversized_action="paginate",
        print_title_rows="$1:$2",
        print_title_columns="$A:$B",
    )

    assert settings.print_title_rows == "$1:$2"
    assert settings.print_title_columns == "$A:$B"


@pytest.mark.parametrize(
    "field,value",
    [
        ("print_title_rows", ""),
        ("print_title_rows", 1),
        ("print_title_columns", "   "),
        ("print_title_columns", False),
    ],
)
def test_excel_settings_reject_invalid_print_titles(field, value):
    with pytest.raises(ValueError, match=field):
        ExcelSettings(**{field: value})


def test_deprecated_page_shrink_threshold_warns_and_is_removed():
    with pytest.warns(FutureWarning, match="page_shrink_threshold"):
        settings = PDFConversionSettings.from_dict(
            {
                "excel": {
                    "orientation": "portrait",
                    "page_shrink_threshold": 0.30,
                }
            }
        )

    assert settings.excel is not None
    assert settings.excel.orientation == "portrait"
    assert not hasattr(settings.excel, "page_shrink_threshold")


def test_flat_excel_print_title_settings_are_loaded():
    settings = PDFConversionSettings.from_dict(
        {
            "print_title_rows": "$1:$3",
            "print_title_columns": "$A:$A",
        }
    )

    assert settings.excel is not None
    assert settings.excel.print_title_rows == "$1:$3"
    assert settings.excel.print_title_columns == "$A:$A"


def test_sheet_specific_merge_preserves_all_excel_base_settings():
    base = PDFConversionSettings(
        excel=ExcelSettings(
            oversized_action="warn",
            print_title_rows="$1:$2",
            print_title_columns="$A:$B",
        )
    )

    with patch("src.config.load_config", return_value={"pdf_settings": {"excel": []}}):
        settings = get_excel_sheet_settings("Data", base_settings=base)

    assert settings.excel is not None
    assert settings.excel.oversized_action == "warn"
    assert settings.excel.print_title_rows == "$1:$2"
    assert settings.excel.print_title_columns == "$A:$B"


@pytest.mark.parametrize(
    "kwargs",
    [
        {"orientation": "diagonal"},
        {"row_dimensions": -1},
        {"row_dimensions": 1.5},
        {"row_dimensions": True},
        {"min_shrink_factor": 0},
        {"min_shrink_factor": 1.1},
        {"min_shrink_factor": "0.9"},
        {"min_shrink_factor": True},
        {"oversized_action": "maybe"},
        {"print_area_policy": "replace"},
    ],
)
def test_invalid_excel_settings_are_fatal(kwargs):
    with pytest.raises(ValueError):
        ExcelSettings(**kwargs)


@pytest.mark.parametrize(
    "kwargs",
    [
        {"margin": -1},
        {"box_mode": "media"},
        {"render_dpi": 17},
        {"max_render_pixels": 0},
        {"background_tolerance": 256},
    ],
)
def test_invalid_trim_settings_are_fatal(kwargs):
    with pytest.raises(ValueError):
        TrimWhitespaceSettings(**kwargs)


def test_malformed_yaml_is_fatal(tmp_path):
    malformed = tmp_path / "bad.yml"
    malformed.write_text("settings: [unterminated", encoding="utf-8")
    with pytest.raises(ValueError, match="Malformed YAML"):
        load_config(malformed)

