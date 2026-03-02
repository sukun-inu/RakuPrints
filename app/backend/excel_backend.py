from __future__ import annotations

import gc
import time

from app.app_context import AppContext
from app.backend.paper_utils import is_supported_name
from app.backend.printer_utils import (
    resolve_excel_printer_name,
    get_default_printer_name,
    list_paper_sizes,
    set_default_printer,
)
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


class ExcelBackend:
    def __init__(self, context: AppContext) -> None:
        self._context = context

    def list_sheets(self, file_path: str) -> list[str]:
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except Exception as exc:
            raise RuntimeError("Excel 印刷には pywin32 が必要です。") from exc

        pythoncom.CoInitialize()
        app = None
        workbook = None
        try:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            if hasattr(app, "DisplayAlerts"):
                app.DisplayAlerts = False
            workbook = app.Workbooks.Open(file_path, ReadOnly=True)
            sheets = [sheet.Name for sheet in workbook.Worksheets]
            return sheets
        finally:
            if workbook is not None:
                workbook.Close(False)
            if app is not None:
                app.Quit()
            if workbook is not None:
                del workbook
            if app is not None:
                del app
            gc.collect()
            pythoncom.CoUninitialize()

    def print(self, job: PrintJob) -> None:
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except Exception as exc:
            raise RuntimeError("Excel 印刷には pywin32 が必要です。") from exc

        pythoncom.CoInitialize()
        app = None
        workbook = None
        default_before = ""
        default_changed = False
        try:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            if hasattr(app, "DisplayAlerts"):
                app.DisplayAlerts = False
            workbook = app.Workbooks.Open(job.file_path, ReadOnly=True)
            printer_for_check = _resolve_printer_name(job)
            _assert_printer_paper_supported(printer_for_check, job.paper_size)
            if job.printer_name:
                default_printer = get_default_printer_name()
                if self._context.settings.use_default_printer and job.printer_name == default_printer:
                    job_printer = ""
                else:
                    job_printer = job.printer_name
            else:
                job_printer = ""

            if job_printer:
                resolved_name = resolve_excel_printer_name(job.printer_name)
                try:
                    app.ActivePrinter = resolved_name
                except Exception:
                    try:
                        app.ActivePrinter = job_printer
                    except Exception:
                        resolved_name = ""
                        default_before = get_default_printer_name()
                        if set_default_printer(job_printer):
                            default_changed = True
                        else:
                            raise RuntimeError(
                                "Excel がプリンターを指定できません。Excel の既定プリンターに切り替えると印刷できます。"
                            )
            if job_printer and not resolved_name and not default_changed:
                raise RuntimeError(
                    "Excel がプリンターを指定できません。Excel の既定プリンターに切り替えると印刷できます。"
                )
            paper_const = _excel_paper_constant(job.paper_size, win32com.client.constants)
            _assert_excel_paper_matches(workbook, job.excel_sheets, paper_const)
            auto_mode = self._context.settings.excel_orientation_mode
            auto_orientation = auto_mode == "auto" or (auto_mode == "ask" and job.excel_auto_orientation)
            if job.excel_sheets:
                for name in job.excel_sheets:
                    try:
                        sheet = workbook.Worksheets(name)
                    except Exception as exc:
                        raise RuntimeError(f"シートが見つかりません: {name}") from exc
                    if paper_const is not None:
                        sheet.PageSetup.PaperSize = paper_const
                    if auto_orientation:
                        sheet.PageSetup.Orientation = _suggest_sheet_orientation(sheet, win32com.client.constants)
                    _apply_sheet_centering(sheet)
                    sheet.PrintOut(Copies=job.copies)
            else:
                for sheet in workbook.Worksheets:
                    if paper_const is not None:
                        sheet.PageSetup.PaperSize = paper_const
                    if auto_orientation:
                        sheet.PageSetup.Orientation = _suggest_sheet_orientation(sheet, win32com.client.constants)
                    _apply_sheet_centering(sheet)
                workbook.PrintOut(Copies=job.copies)
            _wait_for_print_queue(app)
        finally:
            if default_changed and default_before:
                set_default_printer(default_before)
            if workbook is not None:
                workbook.Close(False)
            if app is not None:
                app.Quit()
            if workbook is not None:
                del workbook
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


def _assert_excel_paper_matches(workbook, sheet_names: list[str], paper_const: int | None) -> None:
    if paper_const is None:
        return
    sheets = sheet_names or [sheet.Name for sheet in workbook.Worksheets]
    for name in sheets:
        try:
            sheet = workbook.Worksheets(name)
        except Exception as exc:
            raise RuntimeError(f"シートが見つかりません: {name}") from exc
        try:
            current = int(sheet.PageSetup.PaperSize)
        except Exception as exc:
            raise RuntimeError("Excel の用紙サイズを確認できません。") from exc
        if current != paper_const:
            raise RuntimeError(f"Excel の用紙サイズが選択したサイズと一致しません。 (シート: {name})")


def _excel_paper_constant(name: str, constants) -> int | None:
    if not name:
        return None
    normalized = name.replace(" ", "").replace("-", "").upper()
    mapping = {
        "A3": ["xlPaperA3"],
        "A4": ["xlPaperA4"],
        "A5": ["xlPaperA5"],
        "B4": ["xlPaperB4"],
        "B5": ["xlPaperB5", "xlPaperB5JIS"],
        "LETTER": ["xlPaperLetter"],
        "LEGAL": ["xlPaperLegal"],
    }
    for key, candidates in mapping.items():
        if key in normalized:
            for candidate in candidates:
                value = getattr(constants, candidate, None)
                if value is not None:
                    return value
    return None


def _apply_sheet_centering(sheet) -> None:
    try:
        sheet.PageSetup.CenterHorizontally = True
    except Exception:
        pass
    try:
        sheet.PageSetup.CenterVertically = True
    except Exception:
        pass


def _suggest_sheet_orientation(sheet, constants) -> int:
    used = sheet.UsedRange
    width = float(used.Width)
    height = float(used.Height)
    if width > height * 1.05:
        return getattr(constants, "xlLandscape", 2)
    return getattr(constants, "xlPortrait", 1)
