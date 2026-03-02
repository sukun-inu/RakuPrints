from __future__ import annotations

from PySide6 import QtWidgets

from app.i18n import t
from app.version import __version__
from app.ui.release_notes_dialog import ReleaseNotesDialog


class AboutDialog(QtWidgets.QMessageBox):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setStandardButtons(QtWidgets.QMessageBox.Ok)
        self._release_button = self.addButton(t("btn_release_notes"), QtWidgets.QMessageBox.ActionRole)
        self._release_button.clicked.connect(self._open_release_notes)
        self.retranslate()

    def retranslate(self) -> None:
        self.setWindowTitle(t("about_title"))
        self.setText(t("about_text"))
        info = t("about_info")
        info += f"\n\nVersion: {__version__}\nLicense: MIT"
        self.setInformativeText(info)
        if self._release_button:
            self._release_button.setText(t("btn_release_notes"))

    def _open_release_notes(self) -> None:
        dialog = ReleaseNotesDialog(self)
        dialog.exec()
