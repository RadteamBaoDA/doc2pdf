"""Core converters package."""

import sys

from .word_converter import WordConverter
from .powerpoint_converter import PowerPointConverter
from .excel import COMDisconnectedError, ExcelConverter
from .macro_converter import MacroConverter

# Backwards-compatible module aliases. Excel implementation now lives under
# ``src.core.excel`` while integrations using the historical import paths keep
# resolving to the real modules (including monkeypatch targets).
from .excel import chunking as _excel_chunking
from .excel import content as _excel_content
from .excel import converter as _excel_converter
from .excel import extensions as _excel_extensions
from .excel import layout as _excel_layout
from .excel import models as _excel_layout_models
from .excel import pagination as _excel_pagination
from .excel import pdf_quality as _pdf_quality
from .excel import planner as _excel_layout_planner
from .excel import printer as _excel_printer

_COMPAT_MODULES = {
    "excel_chunking": _excel_chunking,
    "excel_content": _excel_content,
    "excel_converter": _excel_converter,
    "excel_extensions": _excel_extensions,
    "excel_layout": _excel_layout,
    "excel_layout_models": _excel_layout_models,
    "excel_layout_planner": _excel_layout_planner,
    "excel_pagination": _excel_pagination,
    "excel_printer": _excel_printer,
    "pdf_quality": _pdf_quality,
}
for _legacy_name, _module in _COMPAT_MODULES.items():
    sys.modules.setdefault(f"{__name__}.{_legacy_name}", _module)

__all__ = ["WordConverter", "PowerPointConverter", "ExcelConverter", "COMDisconnectedError"]
