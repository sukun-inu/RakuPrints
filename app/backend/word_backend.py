from __future__ import annotations

import gc
import time

from app.app_context import AppContext
from app.model.print_job import PrintJob


class WordBackend:
    def __init__(self, context: AppContext) -> None:
        self._context = context

    def print(self, job: PrintJob) -> None:
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except Exception as exc:
            raise RuntimeError("Word 印刷には pywin32 が必要です。") from exc

        pythoncom.CoInitialize()
        app = None
        doc = None
        try:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            if hasattr(app, "DisplayAlerts"):
                app.DisplayAlerts = False
            doc = app.Documents.Open(job.file_path, ReadOnly=True)
            if job.printer_name:
                app.ActivePrinter = job.printer_name
            paper_const = _word_paper_constant(job.paper_size, win32com.client.constants)
            if paper_const is not None:
                doc.PageSetup.PaperSize = paper_const
            doc.PrintOut(Copies=job.copies, Background=False)
            _wait_for_print_queue(app)
        finally:
            if doc is not None:
                doc.Close(False)
            if app is not None:
                app.Quit()
            if doc is not None:
                del doc
            if app is not None:
                del app
            gc.collect()
            pythoncom.CoUninitialize()


def _wait_for_print_queue(app) -> None:
    if hasattr(app, "BackgroundPrintingStatus"):
        for _ in range(300):
            if app.BackgroundPrintingStatus == 0:
                break
            time.sleep(0.1)


def _word_paper_constant(name: str, constants) -> int | None:
    if not name:
        return None
    normalized = name.replace(" ", "").replace("-", "").upper()
    mapping = {
        "A3": ["wdPaperA3"],
        "A4": ["wdPaperA4"],
        "A5": ["wdPaperA5"],
        "B4": ["wdPaperB4"],
        "B5": ["wdPaperB5", "wdPaperB5JIS"],
        "LETTER": ["wdPaperLetter"],
        "LEGAL": ["wdPaperLegal"],
    }
    for key, candidates in mapping.items():
        if key in normalized:
            for candidate in candidates:
                value = getattr(constants, candidate, None)
                if value is not None:
                    return value
    return None
