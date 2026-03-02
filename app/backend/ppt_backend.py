from __future__ import annotations

import gc
import time

from app.app_context import AppContext
from app.backend.paper_utils import is_supported_name, normalize_name, size_key_from_points
from app.backend.printer_utils import get_default_printer_name, list_paper_sizes
from app.model.print_job import PrintJob


def _resolve_printer_name(job: PrintJob) -> str:
    if job.printer_name:
        return job.printer_name
    try:
        return get_default_printer_name()
    except Exception:
        return ""


def _assert_printer_paper_supported(printer_name: str, paper_size: str) -> None:
    if not printer_name or not paper_size:
        return
    try:
        sizes = list_paper_sizes(printer_name)
    except Exception:
        return
    if sizes and not is_supported_name(paper_size, sizes):
        raise RuntimeError("プリンターが選択した用紙サイズに対応していません。")


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
            printer_name = _resolve_printer_name(job)
            _assert_printer_paper_supported(printer_name, job.paper_size)
            if job.paper_size:
                target_key = normalize_name(job.paper_size)
                if target_key:
                    width_pt = float(presentation.PageSetup.SlideWidth)
                    height_pt = float(presentation.PageSetup.SlideHeight)
                    slide_key = size_key_from_points(width_pt, height_pt)
                    if slide_key and slide_key != target_key:
                        raise RuntimeError("PowerPoint の用紙サイズが選択したサイズと一致しません。")
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
