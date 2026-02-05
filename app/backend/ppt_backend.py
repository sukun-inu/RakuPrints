from __future__ import annotations

import gc
import time

from app.app_context import AppContext
from app.model.print_job import PrintJob


class PptBackend:
    def __init__(self, context: AppContext) -> None:
        self._context = context

    def print(self, job: PrintJob) -> None:
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except Exception as exc:
            raise RuntimeError("PowerPoint 印刷には pywin32 が必要です。") from exc

        pythoncom.CoInitialize()
        app = None
        presentation = None
        try:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            app.Visible = False
            if hasattr(app, "DisplayAlerts"):
                app.DisplayAlerts = False
            presentation = app.Presentations.Open(job.file_path, WithWindow=False)
            if job.printer_name:
                app.ActivePrinter = job.printer_name
            presentation.PrintOut(Copies=job.copies)
            _wait_for_print_queue(app)
        finally:
            if presentation is not None:
                presentation.Close()
            if app is not None:
                app.Quit()
            if presentation is not None:
                del presentation
            if app is not None:
                del app
            gc.collect()
            pythoncom.CoUninitialize()


def _wait_for_print_queue(app) -> None:
    if hasattr(app, "PrintStatus"):
        for _ in range(300):
            if app.PrintStatus == 0:
                break
            time.sleep(0.1)
