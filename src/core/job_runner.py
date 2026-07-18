"""Spawned Excel conversion jobs with parent-owned atomic commit."""

from __future__ import annotations

import multiprocessing as mp
import os
import queue
import tempfile
import traceback
from pathlib import Path
from typing import Any, Dict, Optional

from pypdf import PdfReader

from ..config import PDFConversionSettings
from ..utils.timeout import TimeoutError as ConversionTimeoutError
from .excel_converter import ExcelConverter
from .pdf_processor import PDFProcessor


class JobTimeoutError(ConversionTimeoutError):
    pass


def _excel_job_worker(
    events, source: str, stage: str, settings: PDFConversionSettings,
    trim_options: Optional[Dict[str, Any]], base_path: Optional[str],
) -> None:
    try:
        def record_process(process_id: int) -> None:
            events.put(("office_pid", process_id))

        converter = ExcelConverter(process_recorder=record_process)
        result = converter.convert(
            Path(source), Path(stage), settings,
            base_path=Path(base_path) if base_path else None,
        )
        if trim_options is not None:
            options = dict(trim_options)
            margin = options.pop("margin", 10.0)
            PDFProcessor().trim_whitespace(result, margin=margin, **options)
        validation = PdfReader(str(result))
        if not validation.pages or Path(result).stat().st_size == 0:
            raise ValueError("Job output is empty or unreadable")
        events.put(("success", str(result)))
    except BaseException as exc:
        events.put(("error", f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}"))


def _terminate_recorded_process(process_id: Optional[int]) -> None:
    if not process_id or os.name != "nt":
        return
    try:
        import win32api
        import win32con
        import win32process

        handle = win32api.OpenProcess(
            win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE, False, process_id
        )
        try:
            win32process.TerminateProcess(handle, 1)
        finally:
            handle.Close()
    except Exception:
        pass


def run_excel_job(
    source: Path,
    target: Path,
    settings: PDFConversionSettings,
    *,
    trim_options: Optional[Dict[str, Any]] = None,
    timeout_seconds: Optional[int] = None,
    base_path: Optional[Path] = None,
) -> Path:
    """Run conversion plus trim in a spawned process and commit in the parent."""
    source = source.resolve()
    target = target.resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, stage_name = tempfile.mkstemp(
        prefix=f".{target.name}.", suffix=".job.pdf", dir=str(target.parent)
    )
    os.close(fd)
    stage = Path(stage_name)
    stage.unlink(missing_ok=True)
    context = mp.get_context("spawn")
    events = context.Queue()
    worker = context.Process(
        target=_excel_job_worker,
        args=(events, str(source), str(stage), settings, trim_options,
              str(base_path) if base_path else None),
        daemon=False,
    )
    office_pid: Optional[int] = None
    terminal = None
    try:
        worker.start()
        deadline = None
        if timeout_seconds:
            import time
            deadline = time.monotonic() + timeout_seconds
        while worker.is_alive() or terminal is None:
            if deadline is not None:
                import time
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    _terminate_recorded_process(office_pid)
                    worker.terminate()
                    worker.join(10)
                    if worker.is_alive():
                        worker.kill()
                        worker.join()
                    raise JobTimeoutError(
                        f"Excel conversion and trim exceeded {timeout_seconds} seconds"
                    )
            else:
                remaining = 0.25
            try:
                event = events.get(timeout=min(0.25, remaining) if deadline else 0.25)
            except queue.Empty:
                if not worker.is_alive() and terminal is None:
                    break
                continue
            if event[0] == "office_pid":
                office_pid = int(event[1])
            else:
                terminal = event
                break
        worker.join(10)
        if terminal is None:
            raise RuntimeError(f"Excel worker exited with code {worker.exitcode} without a result")
        if terminal[0] == "error":
            raise RuntimeError(terminal[1])
        output = Path(terminal[1])
        validation = PdfReader(str(output))
        if not validation.pages or output.stat().st_size == 0:
            raise RuntimeError("Parent validation rejected worker output")
        os.replace(output, target)
        return target
    finally:
        if worker.is_alive():
            _terminate_recorded_process(office_pid)
            worker.terminate()
            worker.join()
        stage.unlink(missing_ok=True)
        events.close()
