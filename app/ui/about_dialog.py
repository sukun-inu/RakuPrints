from __future__ import annotations

from PySide6 import QtWidgets

from app.i18n import t
from app.version import __version__


class AboutDialog(QtWidgets.QMessageBox):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setStandardButtons(QtWidgets.QMessageBox.Ok)
        self.retranslate()

    def retranslate(self) -> None:
        self.setWindowTitle(t("about_title"))
        self.setText(t("about_text"))
        info = t("about_info")
        info += f"\n\nVersion: {__version__}\nLicense: MIT"
        self.setInformativeText(info)
