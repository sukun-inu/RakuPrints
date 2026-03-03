from __future__ import annotations

from PySide6 import QtGui, QtPrintSupport


def apply_qt_paper_size(printer: QtPrintSupport.QPrinter, name: str) -> None:
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
