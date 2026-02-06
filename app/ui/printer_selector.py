from __future__ import annotations

from PySide6 import QtCore, QtWidgets

from app.backend.printer_utils import list_printers
from app.i18n import t


class PrinterSelectorDialog(QtWidgets.QDialog):
    def __init__(self, current_printer: str = "", parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.resize(360, 400)

        layout = QtWidgets.QVBoxLayout(self)
        self.printer_list = QtWidgets.QListWidget()
        layout.addWidget(self.printer_list)

        button_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        layout.addWidget(button_box)

        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        self._load_printers(current_printer)
        self.retranslate()

    def retranslate(self) -> None:
        self.setWindowTitle(t("printer_select_title"))

    def _load_printers(self, current_printer: str) -> None:
        try:
            printers = list_printers()
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, t("title_printer"), str(exc))
            printers = []

        for name in printers:
            self.printer_list.addItem(name)
        if current_printer:
            matches = self.printer_list.findItems(current_printer, QtCore.Qt.MatchExactly)
            if matches:
                self.printer_list.setCurrentItem(matches[0])

    def selected_printer(self) -> str:
        item = self.printer_list.currentItem()
        return item.text() if item else ""

    @staticmethod
    def get_printer(parent: QtWidgets.QWidget | None = None, current_printer: str = ""):
        dialog = PrinterSelectorDialog(current_printer=current_printer, parent=parent)
        result = dialog.exec()
        return dialog.selected_printer(), result == QtWidgets.QDialog.Accepted
