from __future__ import annotations

import json
import sys
from pathlib import Path

from PySide6 import QtCore, QtGui, QtPrintSupport, QtWidgets


def _apply_paper_size(printer: QtPrintSupport.QPrinter, name: str) -> None:
    if not name:
        return
    normalized = name.replace(" ", "").replace("-", "").upper()
    mapping = {
        "A3": QtGui.QPageSize.A3,
        "A4": QtGui.QPageSize.A4,
        "A5": QtGui.QPageSize.A5,
        "B4": QtGui.QPageSize.B4,
        "B5": QtGui.QPageSize.B5,
        "LETTER": QtGui.QPageSize.Letter,
        "LEGAL": QtGui.QPageSize.Legal,
    }
    for key, page_id in mapping.items():
        if key in normalized:
            printer.setPageSize(QtGui.QPageSize(page_id))
            break


def _target_rect(printer: QtPrintSupport.QPrinter, use_paper_rect: bool) -> QtCore.QRect:
    if use_paper_rect:
        rect = printer.paperRect(QtPrintSupport.QPrinter.DevicePixel)
        if rect.isNull():
            rect = printer.pageRect(QtPrintSupport.QPrinter.DevicePixel)
    else:
        rect = printer.pageRect(QtPrintSupport.QPrinter.DevicePixel)
        if rect.isNull():
            rect = printer.paperRect(QtPrintSupport.QPrinter.DevicePixel)
    return rect


def _scale_factor(mode: str, source_w: int, source_h: int, target_w: int, target_h: int) -> float:
    if source_w <= 0 or source_h <= 0 or target_w <= 0 or target_h <= 0:
        return 1.0
    fit = min(target_w / source_w, target_h / source_h)
    if mode == "fit":
        return fit
    if mode in ("shrink", "auto"):
        return min(1.0, fit)
    return 1.0


def _read_payload() -> dict:
    raw = sys.stdin.read()
    if not raw:
        return {}
    return json.loads(raw)


def main() -> int:
    payload = _read_payload()
    file_path = payload.get("file_path", "")
    printer_name = payload.get("printer_name", "")
    copies = int(payload.get("copies", 1))
    duplex = payload.get("duplex", "")
    paper_size = payload.get("paper_size", "")
    page_start = int(payload.get("page_start", 0) or 0)
    page_end_raw = payload.get("page_end", None)
    pdf_scale_mode = str(payload.get("pdf_scale_mode", "fit") or "fit")
    try:
        pdf_scale_percent = int(payload.get("pdf_scale_percent", 100))
    except (TypeError, ValueError):
        pdf_scale_percent = 100
    pdf_auto_rotate = bool(payload.get("pdf_auto_rotate", False))
    pdf_center = bool(payload.get("pdf_center", True))

    if not file_path:
        print("File path is required", file=sys.stderr)
        return 2
    if not Path(file_path).exists():
        print("File not found", file=sys.stderr)
        return 2

    try:
        import fitz  # type: ignore
    except Exception:
        print("PyMuPDF is required", file=sys.stderr)
        return 3

    app = QtWidgets.QApplication([])
    doc = None
    painter = QtGui.QPainter()
    try:
        doc = fitz.open(file_path)
        if doc.page_count == 0:
            print("No pages in PDF", file=sys.stderr)
            return 4

        if page_end_raw is None:
            page_end = doc.page_count - 1
        else:
            page_end = int(page_end_raw)

        if page_start < 0:
            page_start = 0
        if page_end < 0:
            page_end = 0
        if page_start >= doc.page_count:
            page_start = doc.page_count - 1
        if page_end >= doc.page_count:
            page_end = doc.page_count - 1
        if page_start > page_end:
            print("Invalid page range", file=sys.stderr)
            return 6

        printer = QtPrintSupport.QPrinter(QtPrintSupport.QPrinter.HighResolution)
        if printer_name:
            printer.setPrinterName(printer_name)
        if copies > 0:
            printer.setCopyCount(copies)
        if duplex == "長辺とじ":
            printer.setDuplex(QtPrintSupport.QPrinter.DuplexLongSide)
        elif duplex == "短辺とじ":
            printer.setDuplex(QtPrintSupport.QPrinter.DuplexShortSide)
        else:
            printer.setDuplex(QtPrintSupport.QPrinter.DuplexNone)
        if paper_size:
            _apply_paper_size(printer, paper_size)

        if not painter.begin(printer):
            print("Failed to initialize printer", file=sys.stderr)
            return 5

        dpi_raw = payload.get("dpi", None)
        try:
            target_dpi = int(dpi_raw) if dpi_raw is not None else 0
        except (TypeError, ValueError):
            target_dpi = 0
        if target_dpi <= 0:
            try:
                target_dpi = int(getattr(printer, "resolution")() or 0)
            except Exception:
                target_dpi = 0
        if target_dpi <= 0:
            target_dpi = 600
        scale = target_dpi / 72.0
        if pdf_scale_percent < 10:
            pdf_scale_percent = 10
        if pdf_scale_percent > 200:
            pdf_scale_percent = 200
        scale_percent = pdf_scale_percent / 100.0
        if pdf_scale_mode not in ("auto", "fit", "shrink", "none"):
            pdf_scale_mode = "auto"
        use_paper_rect = bool(payload.get("pdf_use_paper_rect", False))
        target_rect = _target_rect(printer, use_paper_rect)
        safe_rect = target_rect
        if safe_rect.width() > 2 and safe_rect.height() > 2:
            safe_rect = safe_rect.adjusted(1, 1, -1, -1)
        for page_index in range(page_start, page_end + 1):
            if page_index > page_start:
                printer.newPage()
            page = doc.load_page(page_index)
            rotate = False
            if pdf_auto_rotate:
                page_rect = page.rect
                page_landscape = page_rect.width > page_rect.height
                target_landscape = safe_rect.width() > safe_rect.height()
                rotate = page_landscape != target_landscape
            matrix = fitz.Matrix(scale, scale)
            if rotate:
                matrix = matrix.prerotate(90)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            image = QtGui.QImage(
                pix.samples,
                pix.width,
                pix.height,
                pix.stride,
                QtGui.QImage.Format_RGB888,
            ).copy()
            max_width = int(max(1, safe_rect.width()))
            max_height = int(max(1, safe_rect.height()))
            base_factor = _scale_factor(pdf_scale_mode, image.width(), image.height(), max_width, max_height)
            final_factor = base_factor * scale_percent
            if pdf_scale_mode == "auto":
                max_factor = _scale_factor("fit", image.width(), image.height(), max_width, max_height)
                if final_factor > max_factor:
                    final_factor = max_factor
            scaled_w = max(1, int(image.width() * final_factor))
            scaled_h = max(1, int(image.height() * final_factor))
            scaled = image.scaled(
                scaled_w,
                scaled_h,
                QtCore.Qt.IgnoreAspectRatio,
                QtCore.Qt.SmoothTransformation,
            )
            if pdf_center:
                x = safe_rect.x() + (safe_rect.width() - scaled.width()) // 2
                y = safe_rect.y() + (safe_rect.height() - scaled.height()) // 2
            else:
                x = safe_rect.x()
                y = safe_rect.y()
            painter.drawImage(QtCore.QPoint(x, y), scaled)
    except Exception as exc:
        print(str(exc) or "PDF print failed", file=sys.stderr)
        return 10
    finally:
        if painter.isActive():
            painter.end()
        if doc is not None:
            doc.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
