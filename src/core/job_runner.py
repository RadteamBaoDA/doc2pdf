"""Spawned Excel conversion jobs with parent-owned atomic commit."""

from __future__ import annotations

import multiprocessing as mp
import os
import queue
import shutil
import tempfile
import time
import traceback
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from pypdf import PdfReader

from ..config import PDFConversionSettings
from ..utils.logger import logger
from ..utils.timeout import TimeoutError as ConversionTimeoutError
from .excel import ExcelConverter
from .excel.pdf_quality import PdfQualityExpectation, PdfQualityPostflight
from .pdf_processor import PDFProcessor


class JobTimeoutError(ConversionTimeoutError):
    pass


class JobCancelledError(RuntimeError):
    """Raised when the parent batch cancels an Excel conversion job."""


def _excel_job_worker(
    events, source: str, stage: str, settings: PDFConversionSettings,
    trim_options: Optional[Dict[str, Any]], base_path: Optional[str],
    runtime_evidence: Optional[Dict[str, Any]] = None,
) -> None:
    log_handler_id: Optional[int] = None
    try:
        # Excel runs in a spawned process. Forward its Loguru records through
        # the job queue so the parent TUI can display them while conversion is
        # still in progress.
        logger.remove()

        def forward_log(message) -> None:
            record = message.record
            events.put(("log", record["level"].name, record["message"]))

        log_handler_id = logger.add(forward_log, level="DEBUG", format="{message}")

        def record_process(process_id: int) -> None:
            events.put(("office_pid", process_id))

        def report_progress(amount: float) -> None:
            events.put(("progress", float(amount)))

        converter = ExcelConverter(process_recorder=record_process)
        convert_kwargs = {
            "on_progress": report_progress,
            "base_path": Path(base_path) if base_path else None,
        }
        if runtime_evidence is not None:
            convert_kwargs["runtime_evidence"] = runtime_evidence
        result = converter.convert(Path(source), Path(stage), settings, **convert_kwargs)
        excel_settings = settings.excel
        quality_profile = (
            excel_settings.quality_profile if excel_settings is not None else "legacy"
        )
        trim_policy = (
            excel_settings.trim_policy if excel_settings is not None else "disabled"
        )
        should_trim = trim_options is not None
        if quality_profile != "legacy":
            should_trim = trim_policy != "disabled"
        postflight = None
        postprocess_timings: Dict[str, float] = {}
        if should_trim:
            trim_started = time.perf_counter()
            options = dict(trim_options or {})
            margin = options.pop("margin", 10.0)
            if quality_profile != "legacy":
                options["box_mode"] = trim_policy
                options["render_dpi"] = max(150, int(options.get("render_dpi", 150)))
                options.setdefault("max_render_pixels", 20_000_000)
                options.setdefault("background_tolerance", 8)
                options.setdefault("include_annotations", True)
                options.setdefault("allow_signature_invalidation", False)
            pretrim = Path(str(result) + ".pretrim.pdf")
            try:
                shutil.copyfile(result, pretrim)
                processor = PDFProcessor()
                processor.trim_whitespace(result, margin=margin, **options)
                if quality_profile != "legacy":
                    processor.verify_preserved_content(
                        pretrim, result,
                        render_dpi=int(options.get("render_dpi", 150)),
                        max_render_pixels=int(options.get("max_render_pixels", 20_000_000)),
                        background_tolerance=int(options.get("background_tolerance", 8)),
                        include_annotations=bool(options.get("include_annotations", True)),
                    )
            finally:
                pretrim.unlink(missing_ok=True)
            postprocess_timings["trim"] = time.perf_counter() - trim_started
        # The converter already performed authoritative sheet and final-document
        # postflight. Repeat it only when post-processing changed the bytes.
        if (
            should_trim
            and quality_profile != "legacy"
            and excel_settings is not None
        ):
            postflight_started = time.perf_counter()
            postflight = PdfQualityPostflight().validate(
                result,
                PdfQualityExpectation(
                    min_font_pt=excel_settings.min_effective_font_pt,
                    min_image_dpi=excel_settings.min_effective_image_dpi,
                    max_dimension_in=excel_settings.max_page_dimension_in,
                    max_area_in2=excel_settings.max_page_area_in2,
                    require_searchable_text=False,
                ),
            )
            postprocess_timings["postflight"] = (
                time.perf_counter() - postflight_started
            )
            if not postflight.passed and excel_settings.postflight_policy == "strict":
                raise ValueError(
                    "Final post-trim postflight failed: "
                    + "; ".join(postflight.failures)
                )
        if should_trim and quality_profile != "legacy":
            converter.finalize_postprocess_evidence(
                postflight,
                postprocess_timings,
            )
        if quality_profile == "legacy":
            validation = PdfReader(str(result))
            if not validation.pages or Path(result).stat().st_size == 0:
                raise ValueError("Job output is empty or unreadable")
        events.put(("success", str(result)))
    except BaseException as exc:
        events.put(("error", f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}"))
    finally:
        if log_handler_id is not None:
            logger.remove(log_handler_id)


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
    on_log: Optional[Callable[[str, str], None]] = None,
    on_progress: Optional[Callable[[float], None]] = None,
    on_office_pid: Optional[Callable[[int], None]] = None,
    cancel_event: Optional[Any] = None,
    runtime_evidence: Optional[Dict[str, Any]] = None,
) -> Path:
    """Run conversion plus trim in a spawned process and commit in the parent."""
    if cancel_event is not None and cancel_event.is_set():
        raise JobCancelledError("Excel conversion cancelled before it started")
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
              str(base_path) if base_path else None, runtime_evidence),
        daemon=False,
    )
    office_pid: Optional[int] = None
    terminal = None

    def handle_event(event) -> None:
        nonlocal office_pid, terminal
        if event[0] == "office_pid":
            office_pid = int(event[1])
            if on_office_pid is not None:
                on_office_pid(office_pid)
        elif event[0] == "log":
            if on_log is not None:
                on_log(str(event[1]), str(event[2]))
        elif event[0] == "progress":
            if on_progress is not None:
                on_progress(float(event[1]))
        else:
            terminal = event

    def capture_starting_office_pid() -> None:
        """Allow a just-started worker's PID event to reach the parent."""
        deadline = time.monotonic() + 0.5
        while office_pid is None and worker.is_alive() and time.monotonic() < deadline:
            try:
                event = events.get(timeout=0.05)
            except queue.Empty:
                continue
            handle_event(event)

    try:
        worker.start()
        deadline = None
        if timeout_seconds:
            deadline = time.monotonic() + timeout_seconds
        while worker.is_alive() or terminal is None:
            if cancel_event is not None and cancel_event.is_set():
                capture_starting_office_pid()
                _terminate_recorded_process(office_pid)
                worker.terminate()
                worker.join(10)
                if worker.is_alive():
                    worker.kill()
                    worker.join()
                raise JobCancelledError("Excel conversion cancelled")
            if deadline is not None:
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
            handle_event(event)
            if terminal is not None:
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
