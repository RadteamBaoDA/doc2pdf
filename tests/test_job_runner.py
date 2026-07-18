from pathlib import Path

import pytest

from src.config import PDFConversionSettings
from src.core.job_runner import run_excel_job


def test_worker_failure_leaves_existing_target_untouched(tmp_path):
    source = tmp_path / "missing.xlsx"
    target = tmp_path / "existing.pdf"
    original = b"existing destination"
    target.write_bytes(original)

    with pytest.raises(RuntimeError):
        run_excel_job(
            source, target, PDFConversionSettings(), timeout_seconds=20,
            base_path=tmp_path,
        )

    assert target.read_bytes() == original
    assert not list(tmp_path.glob("*.job.pdf"))
