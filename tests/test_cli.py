import time
from pathlib import Path
from threading import Event, Lock
from unittest.mock import MagicMock, patch

import pytest
from typer.testing import CliRunner

from src.cli import (
    FileConversionResult,
    app,
    get_file_type,
    run_parallel_excel_jobs,
)
from src.version import __version__

runner = CliRunner()


@pytest.fixture(autouse=True)
def mock_console_clear():
    with patch("src.cli.console.clear"):
        yield

@pytest.fixture(autouse=True)
def mock_atexit():
    with patch("src.cli.atexit.register"):
        yield

@pytest.fixture(autouse=True)
def mock_process_registry():
    with patch("src.cli.ProcessRegistry.kill_all"):
        yield

def test_version():

    result = runner.invoke(app, ["--version"])
    assert result.exit_code == 0
    assert f"version: {__version__}" in result.stdout

def test_get_file_type():
    assert get_file_type(Path("test.docx")) == "word"
    assert get_file_type(Path("test.DOC")) == "word"
    assert get_file_type(Path("test.xlsx")) == "excel"
    assert get_file_type(Path("test.ppt")) == "powerpoint"
    assert get_file_type(Path("unknown.txt")) == "word" # Fallback


def test_parallel_excel_scheduler_respects_worker_limit():
    files = [Path(f"book-{index}.xlsx") for index in range(5)]
    lock = Lock()
    active = 0
    maximum_active = 0
    completed = []

    def worker(path):
        nonlocal active, maximum_active
        with lock:
            active += 1
            maximum_active = max(maximum_active, active)
        time.sleep(0.05)
        with lock:
            active -= 1
        return FileConversionResult(path, path.with_suffix(".pdf"), "excel", "success")

    run_parallel_excel_jobs(
        files,
        worker,
        max_workers=2,
        cancel_event=Event(),
        on_complete=completed.append,
    )

    assert maximum_active == 2
    assert {result.input_file for result in completed} == set(files)


def test_parallel_excel_scheduler_supports_serial_mode():
    files = [Path("one.xlsx"), Path("two.xlsx")]
    order = []

    def worker(path):
        order.append(("start", path))
        order.append(("end", path))
        return FileConversionResult(path, path.with_suffix(".pdf"), "excel", "success")

    run_parallel_excel_jobs(
        files,
        worker,
        max_workers=1,
        cancel_event=Event(),
        on_complete=lambda result: order.append(("complete", result.input_file)),
    )

    assert order == [
        ("start", files[0]), ("end", files[0]), ("complete", files[0]),
        ("start", files[1]), ("end", files[1]), ("complete", files[1]),
    ]


def test_parallel_excel_scheduler_starts_heaviest_file_first(tmp_path):
    small = tmp_path / "small.xls"
    large = tmp_path / "large.xls"
    small.write_bytes(b"x")
    large.write_bytes(b"x" * 100)
    order = []

    def worker(path):
        order.append(path)
        return FileConversionResult(
            path, path.with_suffix(".pdf"), "excel", "success"
        )

    run_parallel_excel_jobs(
        [small, large],
        worker,
        max_workers=1,
        cancel_event=Event(),
        on_complete=lambda _result: None,
    )

    assert order == [large, small]


def test_parallel_excel_scheduler_retries_only_transient_failure():
    transient = Path("transient.xlsx")
    quality = Path("quality.xlsx")
    attempts = {transient: 0, quality: 0}
    completed = []

    def worker(path):
        attempts[path] += 1
        if path == transient and attempts[path] == 1:
            return FileConversionResult(
                path, None, "excel", "failed", "RPC server is unavailable"
            )
        if path == quality:
            return FileConversionResult(
                path, None, "excel", "failed", "PDF postflight failed: blank"
            )
        return FileConversionResult(
            path, path.with_suffix(".pdf"), "excel", "success"
        )

    run_parallel_excel_jobs(
        [transient, quality],
        worker,
        max_workers=2,
        cancel_event=Event(),
        on_complete=completed.append,
    )

    assert attempts == {transient: 2, quality: 1}
    assert len(completed) == 2
    assert next(item for item in completed if item.input_file == transient).status == "success"
    assert next(item for item in completed if item.input_file == quality).status == "failed"

@patch("src.cli.WordConverter")
@patch("src.cli.get_files")
def test_convert_success_mock(mock_get_files, mock_converter_cls):
    # Setup mocks
    mock_instance = mock_converter_cls.return_value
    mock_get_files.return_value = [Path("test.docx")]
    
    with runner.isolated_filesystem():
        # Create dummy input
        Path("test.docx").touch()
        
        result = runner.invoke(app, ["convert", "test.docx"])
        
        assert result.exit_code == 0
        # assert "Converting" in result.stdout # TUI hides this
        assert "Conversion Completed" in result.stdout
        assert "Success" in result.stdout
        
        # Verify converter called
        mock_instance.convert.assert_called_once()


@patch("src.cli.WordConverter")
def test_convert_directory(mock_converter_cls):
    mock_instance = mock_converter_cls.return_value
    
    with runner.isolated_filesystem():
        # Setup inputs
        input_dir = Path("input")
        input_dir.mkdir()
        (input_dir / "doc1.docx").touch()
        (input_dir / "doc2.doc").touch()
        
        output_dir = Path("output")
        
        result = runner.invoke(app, ["convert", "input", "--output", "output"])
        
        assert result.exit_code == 0
        # Should convert 2 files
        assert mock_instance.convert.call_count == 2


@patch("src.cli.run_excel_job")
@patch("src.cli.WordConverter")
def test_convert_mixed_batch_parallelizes_only_excel(
    mock_word_converter, mock_run_excel_job
):
    lock = Lock()
    active_excel = 0
    maximum_excel = 0

    def convert_word(source, target, settings, base_path=None):
        with lock:
            assert active_excel == 0
        target.write_bytes(b"word pdf")
        return target

    def convert_excel(source, target, settings, **kwargs):
        nonlocal active_excel, maximum_excel
        with lock:
            active_excel += 1
            maximum_excel = max(maximum_excel, active_excel)
        time.sleep(0.05)
        target.write_bytes(b"excel pdf")
        with lock:
            active_excel -= 1
        return target

    mock_word_converter.return_value.convert.side_effect = convert_word
    mock_run_excel_job.side_effect = convert_excel

    with runner.isolated_filesystem():
        input_dir = Path("input")
        input_dir.mkdir()
        (input_dir / "document.docx").touch()
        (input_dir / "one.xlsx").touch()
        (input_dir / "two.xlsx").touch()

        result = runner.invoke(app, ["convert", "input", "--output", "output"])

        summary = next(Path("reports").glob("summary_*.txt")).read_text(
            encoding="utf-8"
        )

    assert result.exit_code == 0, result.output
    assert maximum_excel == 2
    assert mock_word_converter.return_value.convert.call_count == 1
    assert "Success: 3" in summary
    assert "Failed:  0" in summary
    assert "Skipped: 0" in summary

def test_convert_missing_input():
    # We do NOT mock filesystem here to test generic Typer check, 
    # but Typer checks existence before calling logic if argument has `exists=True`.
    # So we expect fail.
    result = runner.invoke(app, ["convert", "non_existent.docx"])
    assert result.exit_code != 0
    # Typer/Click prints validation errors to output/stderr
    assert "does not exist" in result.output or "Invalid value" in result.output


@patch("src.cli.MacroConverter")
def test_convert_macros_directory(mock_converter_cls):
    mock_converter_cls.return_value.convert.side_effect = lambda source, target: target
    with runner.isolated_filesystem():
        input_dir = Path("input")
        (input_dir / "nested").mkdir(parents=True)
        (input_dir / "a.docm").touch()
        (input_dir / "nested" / "b.pptm").touch()
        (input_dir / "nested" / "c.xlsm").touch()

        result = runner.invoke(app, ["convert-macros", "input", "--output", "clean"])

    assert result.exit_code == 0
    targets = [call.args[1] for call in mock_converter_cls.return_value.convert.call_args_list]
    assert targets == [Path("clean/a.docx"), Path("clean/nested/b.pptx"), Path("clean/nested/c.xlsx")]

@patch("src.cli.get_pdf_handling_config")
@patch("src.cli.shutil.copy2")
def test_convert_pdf_copy(mock_copy, mock_get_config):
    # Setup mock config
    mock_config = MagicMock()
    mock_config.copy_to_output = True
    mock_get_config.return_value = mock_config

    with runner.isolated_filesystem():
        # Setup inputs
        input_dir = Path("input")
        input_dir.mkdir()
        pdf_file = input_dir / "doc.pdf"
        pdf_file.touch()
        
        output_dir = Path("output")
        
        # Test
        result = runner.invoke(app, ["convert", "input", "--output", "output"])
        
        assert result.exit_code == 0
        assert "Conversion Completed" in result.stdout
        # assert "doc.pdf" in result.stdout # In TUI logs
        
        # Verify copy called
        mock_copy.assert_called()

        
@patch("src.cli.get_pdf_handling_config")
@patch("src.cli.shutil.copy2")
def test_convert_pdf_no_copy(mock_copy, mock_get_config):
    # Setup mock config
    mock_config = MagicMock()
    mock_config.copy_to_output = False
    mock_get_config.return_value = mock_config

    with runner.isolated_filesystem():
        # Setup inputs
        input_dir = Path("input")
        input_dir.mkdir()
        pdf_file = input_dir / "doc.pdf"
        pdf_file.touch()
        
        output_dir = Path("output")
        
        # Test
        result = runner.invoke(app, ["convert", "input", "--output", "output"])
        
        assert result.exit_code == 0
        
        # Verify copy NOT called
        mock_copy.assert_not_called()
        # Should be counted as success but logged
        assert "Success" in result.stdout


