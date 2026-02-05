from __future__ import annotations

import logging

from PySide6 import QtCore

from app.app_context import AppContext
from app.controller.job_manager import JobManager
from app.model.print_job import PrintJob, FileType, JobStatus
from app.backend.pdf_backend import PdfBackend
from app.backend.word_backend import WordBackend
from app.backend.excel_backend import ExcelBackend
from app.backend.ppt_backend import PptBackend


class JobExecutor(QtCore.QThread):
    job_status = QtCore.Signal(str, object, str)
    progress = QtCore.Signal(int, int, str)
    job_failed = QtCore.Signal(str)
    finished_all = QtCore.Signal(bool)

    def __init__(
        self,
        context: AppContext,
        job_manager: JobManager,
        jobs_override: list[PrintJob] | None = None,
    ) -> None:
        super().__init__()
        self._context = context
        self._job_manager = job_manager
        self._jobs_override = list(jobs_override) if jobs_override is not None else None
        self._cancel_requested = False
        self._logger = logging.getLogger(__name__)

    def request_cancel(self) -> None:
        self._cancel_requested = True

    def run(self) -> None:
        jobs = self._jobs_override if self._jobs_override is not None else [
            job for job in self._job_manager.jobs() if job.enabled
        ]
        total = len(jobs)
        completed = 0
        cancelled = False

        for job in jobs:
            if self._cancel_requested:
                cancelled = True
                self._mark_remaining_as_cancelled(job, jobs)
                break
            self.job_status.emit(job.id, JobStatus.PRINTING, "")
            self.progress.emit(completed, total, job.file_name)

            try:
                self._logger.info(
                    "Printing %s | printer=%s copies=%s duplex=%s",
                    job.file_path,
                    job.printer_name or "default",
                    job.copies,
                    job.duplex.value,
                )
                if job.paper_size:
                    self._logger.info("Paper size: %s", job.paper_size)
                backend = self._resolve_backend(job)
                backend.print(job)
                self.job_status.emit(job.id, JobStatus.SUCCESS, "")
            except Exception as exc:
                self._logger.exception("Print failed for %s", job.file_path)
                message = str(exc) or "予期しないエラーが発生しました。"
                self.job_status.emit(job.id, JobStatus.FAILED, message)
                self.job_failed.emit(job.id)
            finally:
                completed += 1
                self.progress.emit(completed, total, job.file_name)

        self.finished_all.emit(cancelled)

    def _resolve_backend(self, job: PrintJob):
        if job.file_type == FileType.PDF:
            return PdfBackend(self._context)
        if job.file_type == FileType.WORD:
            return WordBackend(self._context)
        if job.file_type == FileType.EXCEL:
            return ExcelBackend(self._context)
        if job.file_type == FileType.PPT:
            return PptBackend(self._context)
        raise RuntimeError("Unsupported file type")

    def _mark_remaining_as_cancelled(self, current_job: PrintJob, jobs: list[PrintJob]) -> None:
        mark = False
        for job in jobs:
            if job.id == current_job.id:
                mark = True
            if mark:
                self.job_status.emit(job.id, JobStatus.CANCELLED, "キャンセルしました")
