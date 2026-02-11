from __future__ import annotations

import gc
from dataclasses import dataclass

from PySide6 import QtCore

from app.model.print_job import FileType


@dataclass(frozen=True)
class PageCountRequest:
    job_id: str
    file_path: str
    file_type: FileType
    excel_sheets: list[str]
    revision: int


class PageCountWorker(QtCore.QThread):
    page_count_ready = QtCore.Signal(str, int, int)
    page_count_failed = QtCore.Signal(str, int)

    def __init__(self, requests: list[PageCountRequest]) -> None:
        super().__init__()
        self._requests = list(requests)

    def run(self) -> None:
        com_needed = any(
            req.file_type in (FileType.WORD, FileType.EXCEL, FileType.PPT) for req in self._requests
        )
        com_ready = False
        pythoncom = None
        win32 = None
        constants = None
        if com_needed:
            try:
                import pythoncom  # type: ignore
                import win32com.client  # type: ignore

                com_ready = True
            except Exception:
                com_ready = False

        if com_ready:
            pythoncom.CoInitialize()

        word_app = None
        excel_app = None
        ppt_app = None
        try:
            if com_ready:
                win32 = win32com.client
                constants = win32com.client.constants

            for req in self._requests:
                try:
                    if req.file_type == FileType.PDF:
                        count = _count_pdf(req.file_path)
                    elif req.file_type == FileType.WORD:
                        if not com_ready:
                            raise RuntimeError("pywin32 is required.")
                        if word_app is None:
                            word_app = win32.DispatchEx("Word.Application")
                            word_app.Visible = False
                            if hasattr(word_app, "DisplayAlerts"):
                                word_app.DisplayAlerts = False
                        count = _count_word(req.file_path, word_app, constants)
                    elif req.file_type == FileType.EXCEL:
                        if not com_ready:
                            raise RuntimeError("pywin32 is required.")
                        if excel_app is None:
                            excel_app = win32.DispatchEx("Excel.Application")
                            excel_app.Visible = False
                            if hasattr(excel_app, "DisplayAlerts"):
                                excel_app.DisplayAlerts = False
                        count = _count_excel(req.file_path, excel_app, req.excel_sheets)
                    elif req.file_type == FileType.PPT:
                        if not com_ready:
                            raise RuntimeError("pywin32 is required.")
                        if ppt_app is None:
                            ppt_app = win32.DispatchEx("PowerPoint.Application")
                            ppt_app.Visible = False
                            if hasattr(ppt_app, "DisplayAlerts"):
                                ppt_app.DisplayAlerts = False
                        count = _count_ppt(req.file_path, ppt_app)
                    else:
                        raise RuntimeError("Unsupported file type.")
                except Exception:
                    self.page_count_failed.emit(req.job_id, req.revision)
                    continue

                self.page_count_ready.emit(req.job_id, count, req.revision)
        finally:
            if word_app is not None:
                try:
                    word_app.Quit()
                except Exception:
                    pass
                try:
                    del word_app
                except Exception:
                    pass
            if excel_app is not None:
                try:
                    excel_app.Quit()
                except Exception:
                    pass
                try:
                    del excel_app
                except Exception:
                    pass
            if ppt_app is not None:
                try:
                    ppt_app.Quit()
                except Exception:
                    pass
                try:
                    del ppt_app
                except Exception:
                    pass
            gc.collect()
            if com_ready:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


def _count_pdf(file_path: str) -> int:
    try:
        import fitz  # type: ignore
    except Exception as exc:
        raise RuntimeError("PyMuPDF is required.") from exc

    doc = None
    try:
        doc = fitz.open(file_path)
        page_count = int(getattr(doc, "page_count", 0) or 0)
    except Exception as exc:
        raise RuntimeError("Failed to read PDF.") from exc
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass
    if page_count <= 0:
        raise RuntimeError("No pages found.")
    return page_count


def _count_word(file_path: str, app, constants) -> int:
    doc = None
    try:
        doc = app.Documents.Open(file_path, ReadOnly=True)
        statistic = getattr(constants, "wdStatisticPages", 2)
        count = int(doc.ComputeStatistics(statistic))
    except Exception as exc:
        raise RuntimeError("Failed to read Word document.") from exc
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
    if count <= 0:
        raise RuntimeError("No pages found.")
    return count


def _count_excel(file_path: str, app, sheet_names: list[str]) -> int:
    workbook = None
    try:
        workbook = app.Workbooks.Open(file_path, ReadOnly=True)
        sheets = sheet_names or [sheet.Name for sheet in workbook.Worksheets]
        total = 0
        for name in sheets:
            sheet = workbook.Worksheets(name)
            count = _excel_sheet_pages(sheet)
            if count is None or count <= 0:
                raise RuntimeError("Failed to read Excel pages.")
            total += count
    except Exception as exc:
        raise RuntimeError("Failed to read Excel workbook.") from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
    if total <= 0:
        raise RuntimeError("No pages found.")
    return total


def _excel_sheet_pages(sheet) -> int | None:
    try:
        sheet.DisplayPageBreaks = True
    except Exception:
        pass
    try:
        pages = sheet.PageSetup.Pages
        count = int(pages.Count)
        if count > 0:
            return count
    except Exception:
        pass
    try:
        h_breaks = int(sheet.HPageBreaks.Count)
        v_breaks = int(sheet.VPageBreaks.Count)
        return max(1, (h_breaks + 1) * (v_breaks + 1))
    except Exception:
        return None


def _count_ppt(file_path: str, app) -> int:
    presentation = None
    try:
        presentation = app.Presentations.Open(file_path, WithWindow=False)
        count = int(presentation.Slides.Count)
    except Exception as exc:
        raise RuntimeError("Failed to read PowerPoint.") from exc
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
    if count <= 0:
        raise RuntimeError("No pages found.")
    return count
