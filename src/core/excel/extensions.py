"""Extension contracts deliberately left unimplemented for M0-M5."""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Sequence


class UnsupportedExcelExtensionError(RuntimeError):
    pass


class VectorStitchProvider(ABC):
    @abstractmethod
    def stitch(self, tiles: Sequence[Path], output: Path) -> Path:
        """Document this Excel pipeline operation and its side effects."""
        raise NotImplementedError


class PdfAProvider(ABC):
    @abstractmethod
    def convert_and_validate(self, source: Path, output: Path) -> Path:
        """Document this Excel pipeline operation and its side effects."""
        raise NotImplementedError


def require_supported_extensions(settings, compliance: str) -> None:
    """Document this Excel pipeline operation and its side effects."""
    strategy = settings.horizontal_overflow_strategy
    if strategy in {"one_logical_page", "vector_stitch"}:
        raise UnsupportedExcelExtensionError(
            f"Excel horizontal strategy {strategy!r} requires an M6 extension provider"
        )
    if settings.quality_profile == "strict" and str(compliance).lower() == "pdfa":
        raise UnsupportedExcelExtensionError(
            "Strict PDF/A output requires a configured conversion and validation provider"
        )
