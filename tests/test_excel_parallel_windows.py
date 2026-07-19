"""Opt-in acceptance and timing check for real desktop Excel.

Run with ``DOC2PDF_RUN_EXCEL_INTEGRATION=1`` on a Windows host with Excel.
"""

import os
import threading
import time
from pathlib import Path

import pytest
from pypdf import PdfReader

from src.cli import FileConversionResult, run_parallel_excel_jobs
from src.config import ExcelSettings, PDFConversionSettings
from src.core.job_runner import run_excel_job

pytestmark = pytest.mark.skipif(
    os.name != "nt" or os.environ.get("DOC2PDF_RUN_EXCEL_INTEGRATION") != "1",
    reason="requires desktop Excel and DOC2PDF_RUN_EXCEL_INTEGRATION=1",
)


def _create_workbook(path: Path, label: str) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Add()
        sheet = workbook.Worksheets(1)
        sheet.Cells(1, 1).Value = label
        for row in range(2, 502):
            sheet.Cells(row, 1).Value = row - 1
            sheet.Cells(row, 2).Formula = f"=A{row}*2"
        workbook.SaveAs(str(path.resolve()), FileFormat=51)
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def _excel_process_has_exited(process_id: int) -> bool:
    import pywintypes
    import win32api
    import win32con
    import win32event

    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, process_id)
    except pywintypes.error:
        return True
    try:
        return win32event.WaitForSingleObject(handle, 5000) == win32event.WAIT_OBJECT_0
    finally:
        handle.Close()


def test_parallel_excel_matches_serial_output_and_leaves_no_processes(
    tmp_path, record_property
):
    sources = [tmp_path / "one.xlsx", tmp_path / "two.xlsx"]
    for index, source in enumerate(sources, start=1):
        _create_workbook(source, f"Parallel Excel fixture {index}")

    settings = PDFConversionSettings(excel=ExcelSettings())

    def run_batch(max_workers: int, output_dir: Path):
        output_dir.mkdir()
        observed_pids = []
        completed = []
        started = time.perf_counter()

        def worker(source: Path) -> FileConversionResult:
            target = output_dir / source.with_suffix(".pdf").name
            result = run_excel_job(
                source,
                target,
                settings,
                timeout_seconds=300,
                base_path=tmp_path,
                on_office_pid=observed_pids.append,
            )
            return FileConversionResult(source, result, "excel", "success")

        run_parallel_excel_jobs(
            sources,
            worker,
            max_workers=max_workers,
            cancel_event=threading.Event(),
            on_complete=completed.append,
        )
        elapsed = time.perf_counter() - started
        page_counts = {
            result.input_file.name: len(PdfReader(str(result.output_file)).pages)
            for result in completed
        }
        return elapsed, page_counts, observed_pids

    serial_seconds, serial_pages, serial_pids = run_batch(1, tmp_path / "serial")
    parallel_seconds, parallel_pages, parallel_pids = run_batch(
        2, tmp_path / "parallel"
    )

    record_property("serial_seconds", serial_seconds)
    record_property("parallel_seconds", parallel_seconds)
    record_property("parallel_speedup", serial_seconds / parallel_seconds)

    assert serial_pages == parallel_pages
    assert all(page_count > 0 for page_count in parallel_pages.values())
    assert len(set(parallel_pids)) == len(sources)
    assert all(
        _excel_process_has_exited(process_id)
        for process_id in serial_pids + parallel_pids
    )
