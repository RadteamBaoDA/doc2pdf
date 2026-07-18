"""Convert macro-enabled Office files to macro-free Open XML formats."""

from contextlib import contextmanager
from pathlib import Path

import pythoncom
import win32com.client


SUPPORTED_FORMATS = {
    ".docm": (".docx", "Word.Application", 16),
    ".pptm": (".pptx", "PowerPoint.Application", 24),
    ".xlsm": (".xlsx", "Excel.Application", 51),
}


class MacroConverter:
    """Remove VBA projects by saving files in macro-free Office formats."""

    def convert(self, input_path: Path, output_path: Path | None = None) -> Path:
        input_file = input_path.resolve()
        if not input_file.is_file():
            raise FileNotFoundError(f"Input file not found: {input_file}")

        extension = input_file.suffix.lower()
        if extension not in SUPPORTED_FORMATS:
            supported = ", ".join(SUPPORTED_FORMATS)
            raise ValueError(f"Unsupported macro file '{extension}'. Expected: {supported}")

        target_extension, program_id, file_format = SUPPORTED_FORMATS[extension]
        output_file = (output_path or input_file.with_suffix(target_extension)).resolve()
        if output_file.suffix.lower() != target_extension:
            raise ValueError(f"Output for {extension} must use the {target_extension} extension")
        if output_file == input_file:
            raise ValueError("Output path must be different from the input path")

        output_file.parent.mkdir(parents=True, exist_ok=True)
        pythoncom.CoInitialize()
        try:
            with self._office_application(program_id) as application:
                if extension == ".docm":
                    document = application.Documents.Open(
                        str(input_file), ConfirmConversions=False, ReadOnly=True,
                        AddToRecentFiles=False, Visible=False,
                    )
                    try:
                        document.SaveAs2(str(output_file), FileFormat=file_format)
                    finally:
                        document.Close(SaveChanges=0)
                elif extension == ".pptm":
                    presentation = application.Presentations.Open(
                        str(input_file), ReadOnly=-1, Untitled=0, WithWindow=0,
                    )
                    try:
                        presentation.SaveAs(str(output_file), file_format)
                    finally:
                        presentation.Close()
                else:
                    workbook = application.Workbooks.Open(
                        str(input_file), UpdateLinks=0, ReadOnly=True,
                        IgnoreReadOnlyRecommended=True, AddToMru=False,
                    )
                    try:
                        workbook.SaveAs(str(output_file), FileFormat=file_format)
                    finally:
                        workbook.Close(SaveChanges=False)
        finally:
            pythoncom.CoUninitialize()

        return output_file

    @contextmanager
    def _office_application(self, program_id: str):
        application = win32com.client.Dispatch(program_id)
        try:
            application.Visible = False
            application.DisplayAlerts = 0
            application.AutomationSecurity = 3
            yield application
        finally:
            application.Quit()
