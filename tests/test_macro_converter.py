from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from src.core.macro_converter import MacroConverter


@pytest.mark.parametrize(
    ("filename", "program_id", "expected_suffix"),
    [
        ("sample.docm", "Word.Application", ".docx"),
        ("sample.pptm", "PowerPoint.Application", ".pptx"),
        ("sample.xlsm", "Excel.Application", ".xlsx"),
    ],
)
def test_convert_macro_file(tmp_path, filename, program_id, expected_suffix):
    source = tmp_path / filename
    source.touch()
    application = MagicMock()

    with patch("src.core.macro_converter.win32com.client.Dispatch", return_value=application), patch(
        "src.core.macro_converter.pythoncom"
    ):
        result = MacroConverter().convert(source)

    assert result.suffix == expected_suffix
    application.Quit.assert_called_once()
    if filename.endswith(".docm"):
        application.Documents.Open.return_value.SaveAs2.assert_called_once()
    elif filename.endswith(".pptm"):
        application.Presentations.Open.return_value.SaveAs.assert_called_once()
    else:
        application.Workbooks.Open.return_value.SaveAs.assert_called_once()


def test_rejects_unsupported_file(tmp_path):
    source = tmp_path / "sample.docx"
    source.touch()
    with pytest.raises(ValueError, match="Unsupported macro file"):
        MacroConverter().convert(source)
