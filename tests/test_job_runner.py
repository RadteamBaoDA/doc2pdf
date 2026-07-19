from queue import Queue
from threading import Event
from unittest.mock import patch

import pytest
from pypdf import PdfWriter

from src.config import ExcelSettings, PDFConversionSettings
from src.core import job_runner
from src.core.job_runner import JobCancelledError, run_excel_job
from src.utils.logger import logger


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


def test_excel_worker_forwards_logs_to_parent_event_queue(tmp_path, monkeypatch):
    class FailingConverter:
        def __init__(self, process_recorder):
            self.process_recorder = process_recorder

        def convert(self, source, stage, settings, on_progress, base_path):
            logger.info("Opening workbook in child process")
            on_progress(0.5)
            raise ValueError("intentional conversion failure")

    monkeypatch.setattr(job_runner, "ExcelConverter", FailingConverter)
    events = Queue()

    job_runner._excel_job_worker(
        events,
        str(tmp_path / "source.xlsx"),
        str(tmp_path / "stage.pdf"),
        PDFConversionSettings(),
        None,
        str(tmp_path),
    )

    received = []
    while not events.empty():
        received.append(events.get_nowait())

    assert ("log", "INFO", "Opening workbook in child process") in received
    assert ("progress", 0.5) in received
    assert received[-1][0] == "error"


def test_strict_worker_does_not_repeat_untrimmed_final_postflight(
    tmp_path, monkeypatch
):
    class VerifiedConverter:
        def __init__(self, process_recorder):
            self.process_recorder = process_recorder

        def convert(self, _source, stage, _settings, **_kwargs):
            writer = PdfWriter()
            writer.add_blank_page(width=100, height=100)
            with stage.open("wb") as stream:
                writer.write(stream)
            return stage

        def finalize_postprocess_evidence(self, _postflight, _timings):
            raise AssertionError("untrimmed output must not be finalized twice")

    monkeypatch.setattr(job_runner, "ExcelConverter", VerifiedConverter)
    events = Queue()
    with patch.object(
        job_runner.PdfQualityPostflight,
        "validate",
        side_effect=AssertionError("duplicate postflight"),
    ):
        job_runner._excel_job_worker(
            events,
            str(tmp_path / "source.xlsx"),
            str(tmp_path / "stage.pdf"),
            PDFConversionSettings(excel=ExcelSettings()),
            None,
            str(tmp_path),
        )

    received = []
    while not events.empty():
        received.append(events.get_nowait())
    assert received[-1][0] == "success"


def test_cancelled_excel_job_does_not_spawn_or_touch_target(tmp_path, monkeypatch):
    source = tmp_path / "source.xlsx"
    target = tmp_path / "existing.pdf"
    target.write_bytes(b"existing destination")
    cancelled = Event()
    cancelled.set()

    get_context_called = False

    def fail_if_context_requested(*args, **kwargs):
        nonlocal get_context_called
        get_context_called = True
        raise AssertionError("cancelled jobs must not spawn")

    monkeypatch.setattr(job_runner.mp, "get_context", fail_if_context_requested)

    with pytest.raises(JobCancelledError, match="before it started"):
        run_excel_job(
            source,
            target,
            PDFConversionSettings(),
            cancel_event=cancelled,
        )

    assert get_context_called is False
    assert target.read_bytes() == b"existing destination"
    assert not list(tmp_path.glob("*.job.pdf"))


def test_active_excel_job_cancellation_terminates_worker_and_cleans_stage(
    tmp_path, monkeypatch
):
    source = tmp_path / "source.xlsx"
    target = tmp_path / "existing.pdf"
    target.write_bytes(b"existing destination")
    cancelled = Event()

    class FakeQueue(Queue):
        def close(self):
            return None

    events = FakeQueue()

    class FakeProcess:
        exitcode = None

        def __init__(self):
            self.alive = False
            self.terminated = False

        def start(self):
            self.alive = True
            events.put(("office_pid", 4321))
            cancelled.set()

        def is_alive(self):
            return self.alive

        def terminate(self):
            self.terminated = True
            self.alive = False

        def kill(self):
            self.alive = False

        def join(self, timeout=None):
            return None

    fake_process = FakeProcess()

    class FakeContext:
        def Queue(self):
            return events

        def Process(self, **kwargs):
            return fake_process

    monkeypatch.setattr(job_runner.mp, "get_context", lambda method: FakeContext())
    terminated_office_pids = []
    observed_office_pids = []
    monkeypatch.setattr(
        job_runner,
        "_terminate_recorded_process",
        terminated_office_pids.append,
    )

    with pytest.raises(JobCancelledError, match="cancelled"):
        run_excel_job(
            source,
            target,
            PDFConversionSettings(),
            on_office_pid=observed_office_pids.append,
            cancel_event=cancelled,
        )

    assert fake_process.terminated is True
    assert observed_office_pids == [4321]
    assert terminated_office_pids == [4321]
    assert target.read_bytes() == b"existing destination"
    assert not list(tmp_path.glob("*.job.pdf"))
