from __future__ import annotations

from pathlib import Path
from typing import List

from PySide6 import QtCore

from app.app_context import AppContext
from app.controller.rules_engine import RulesEngine
from app.model.print_job import PrintJob, FileType, JobStatus
from app.backend.printer_utils import get_default_printer_name
from app.backend.excel_backend import ExcelBackend


class JobManager(QtCore.QObject):
    jobs_changed = QtCore.Signal()
    job_updated = QtCore.Signal(str)

    def __init__(self, context: AppContext) -> None:
        super().__init__()
        self._context = context
        self._rules = RulesEngine(context)
        self._jobs: List[PrintJob] = []

    def jobs(self) -> List[PrintJob]:
        return list(self._jobs)

    def job_count(self) -> int:
        return len(self._jobs)

    def get_job(self, index: int) -> PrintJob:
        return self._jobs[index]

    def find_job_by_id(self, job_id: str) -> PrintJob | None:
        for job in self._jobs:
            if job.id == job_id:
                return job
        return None

    def clear(self) -> None:
        self._jobs.clear()
        self.jobs_changed.emit()

    def add_files(self, file_paths: List[str]) -> None:
        added = False
        for path in file_paths:
            normalized = str(Path(path))
            if any(job.file_path == normalized for job in self._jobs):
                continue
            file_type = self._detect_file_type(normalized)
            if file_type == FileType.UNKNOWN:
                continue
            printer = self._resolve_printer_for_path(normalized)
            job = PrintJob(
                file_path=normalized,
                file_type=file_type,
                printer_name=printer,
                copies=self._context.settings.copies,
                duplex=self._context.settings.duplex,
                paper_size=self._context.settings.paper_size,
            )
            self._jobs.append(job)
            added = True
        if added:
            self.jobs_changed.emit()

    def add_folder(self, folder_path: str, recursive: bool = True) -> None:
        folder = Path(folder_path)
        if not folder.exists() or not folder.is_dir():
            return
        pattern = "**/*" if recursive else "*"
        files = [str(path) for path in folder.glob(pattern) if path.is_file()]
        self.add_files(files)

    def move_job(self, from_index: int, to_index: int) -> None:
        if from_index < 0 or from_index >= len(self._jobs):
            return
        if to_index < 0:
            to_index = 0
        if to_index >= len(self._jobs):
            to_index = len(self._jobs) - 1
        if from_index == to_index:
            return
        job = self._jobs.pop(from_index)
        self._jobs.insert(to_index, job)
        self.jobs_changed.emit()

    def set_job_enabled(self, job_id: str, enabled: bool) -> None:
        job = self.find_job_by_id(job_id)
        if not job:
            return
        job.enabled = enabled
        self.job_updated.emit(job_id)

    def set_jobs_enabled(self, job_ids: list[str], enabled: bool) -> None:
        updated = False
        ids = set(job_ids)
        for job in self._jobs:
            if job.id in ids:
                job.enabled = enabled
                updated = True
        if updated:
            self.jobs_changed.emit()

    def set_jobs_printer(self, job_ids: list[str], printer_name: str) -> None:
        updated = False
        ids = set(job_ids)
        for job in self._jobs:
            if job.id in ids:
                job.printer_name = printer_name
                job.manual_printer = True
                updated = True
        if updated:
            self.jobs_changed.emit()

    def remove_jobs(self, job_ids: list[str]) -> None:
        if not job_ids:
            return
        ids = set(job_ids)
        before = len(self._jobs)
        self._jobs = [job for job in self._jobs if job.id not in ids]
        if len(self._jobs) != before:
            self.jobs_changed.emit()

    def set_job_printer(self, job_id: str, printer_name: str) -> None:
        job = self.find_job_by_id(job_id)
        if not job:
            return
        job.printer_name = printer_name
        job.manual_printer = True
        self.job_updated.emit(job_id)

    def set_job_sheets(self, job_id: str, sheet_names: list[str]) -> None:
        job = self.find_job_by_id(job_id)
        if not job:
            return
        job.excel_sheets = list(sheet_names)
        self.job_updated.emit(job_id)

    def list_excel_sheets(self, file_path: str) -> list[str]:
        backend = ExcelBackend(self._context)
        return backend.list_sheets(file_path)

    def set_job_status(self, job_id: str, status: JobStatus, message: str = "") -> None:
        job = self.find_job_by_id(job_id)
        if not job:
            return
        job.status = status
        job.message = message
        if status == JobStatus.FAILED:
            job.summary = self._summarize_message(message)
        else:
            job.summary = ""
        self.job_updated.emit(job_id)

    def get_failed_jobs(self) -> List[PrintJob]:
        return [job for job in self._jobs if job.enabled and job.status == JobStatus.FAILED]

    def get_enabled_jobs(self) -> List[PrintJob]:
        return [job for job in self._jobs if job.enabled]

    def reset_failed_jobs(self) -> None:
        updated = False
        for job in self._jobs:
            if job.status == JobStatus.FAILED:
                job.status = JobStatus.WAITING
                job.message = ""
                job.summary = ""
                updated = True
        if updated:
            self.jobs_changed.emit()

    @staticmethod
    def _summarize_message(message: str) -> str:
        text = message.lower()
        if ("pymupdf" in text or "fitz" in text) and ("必要" in message or "required" in text):
            return "PDF 印刷に必要な部品が見つかりません。"
        if "pywin32" in text or "win32com" in text:
            return "Office 印刷に必要な部品が見つかりません。"
        if "ファイルが見つかりません" in message or "file not found" in text:
            return "ファイルが見つかりません。"
        if "printer" in text or "プリンター" in message:
            return "プリンターに接続できません。"
        if "access is denied" in text or "権限" in message:
            return "権限が足りません。"
        return "印刷に失敗しました。"

    def apply_rules(self, force: bool = False) -> None:
        for job in self._jobs:
            if force or not job.manual_printer:
                job.printer_name = self._resolve_printer_for_path(job.file_path)
                if force:
                    job.manual_printer = False
        self.jobs_changed.emit()

    def apply_settings_to_jobs(self) -> None:
        for job in self._jobs:
            job.copies = self._context.settings.copies
            job.duplex = self._context.settings.duplex
            job.paper_size = self._context.settings.paper_size
        self.jobs_changed.emit()

    def reset_statuses(self) -> None:
        for job in self._jobs:
            job.status = JobStatus.WAITING
            job.message = ""
            job.summary = ""
        self.jobs_changed.emit()

    def reset_statuses_for(self, job_ids: list[str]) -> None:
        ids = set(job_ids)
        updated = False
        for job in self._jobs:
            if job.id in ids:
                job.status = JobStatus.WAITING
                job.message = ""
                job.summary = ""
                updated = True
        if updated:
            self.jobs_changed.emit()

    def sort_jobs(self, column: int, descending: bool = False) -> None:
        def key(job: PrintJob):
            if column == 0:
                return 1 if job.enabled else 0
            if column == 1:
                return job.file_name.lower()
            if column == 2:
                return job.file_path.lower()
            if column == 3:
                return job.file_type.value.lower()
            if column == 4:
                return job.auto_label().lower()
            if column == 5:
                return job.display_sheets().lower()
            if column == 6:
                return job.display_printer().lower()
            if column == 7:
                return job.status.value.lower()
            return ""

        self._jobs.sort(key=key, reverse=descending)
        self.jobs_changed.emit()

    def _resolve_printer_for_path(self, file_path: str) -> str:
        if self._context.settings.use_default_printer:
            try:
                default_printer = get_default_printer_name()
            except Exception:
                default_printer = ""
        else:
            default_printer = self._context.settings.selected_printer
        return self._rules.resolve_printer(file_path, default_printer)

    @staticmethod
    def _detect_file_type(file_path: str) -> FileType:
        ext = Path(file_path).suffix.lower()
        if ext == ".pdf":
            return FileType.PDF
        if ext in (".doc", ".docx"):
            return FileType.WORD
        if ext in (".xls", ".xlsx", ".xlsm"):
            return FileType.EXCEL
        if ext in (".ppt", ".pptx"):
            return FileType.PPT
        return FileType.UNKNOWN
