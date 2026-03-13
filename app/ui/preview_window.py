from __future__ import annotations

from PySide6 import QtCore, QtGui, QtWidgets


class PreviewWindow(QtWidgets.QMainWindow):
    closed = QtCore.Signal()

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.resize(860, 640)
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose, False)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        self.closed.emit()
        super().closeEvent(event)

