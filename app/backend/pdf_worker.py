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

        target_dpi = int(payload.get("dpi", 600))
        scale = target_dpi / 72.0
        for page_index in range(doc.page_count):
            if page_index > 0:
                printer.newPage()
            page = doc.load_page(page_index)
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            image = QtGui.QImage(
                pix.samples,
                pix.width,
                pix.height,
                pix.stride,
                QtGui.QImage.Format_RGB888,
            ).copy()

            page_rect = printer.pageRect(QtPrintSupport.QPrinter.DevicePixel)
            target_size = page_rect.size().toSize() if hasattr(page_rect.size(), "toSize") else page_rect.size()
            scaled = image.scaled(
                target_size,
                QtCore.Qt.KeepAspectRatio,
                QtCore.Qt.SmoothTransformation,
            )
            x = page_rect.x() + (page_rect.width() - scaled.width()) // 2
            y = page_rect.y() + (page_rect.height() - scaled.height()) // 2
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
