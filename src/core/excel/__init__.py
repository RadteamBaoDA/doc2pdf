"""Excel conversion business logic.

This package owns discovery, layout planning, printer integration, pagination,
quality evidence, export, and postflight for Excel workbooks.
"""

from .converter import COMDisconnectedError, ExcelConverter

__all__ = ["ExcelConverter", "COMDisconnectedError"]
