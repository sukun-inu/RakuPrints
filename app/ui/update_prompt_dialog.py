from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from PySide6 import QtCore, QtGui, QtWidgets

from app.i18n import t

if TYPE_CHECKING:
    from app.controller.update_manager import UpdateInfo


@dataclass(frozen=True)
class UpdatePromptResult:
    update_now: bool
    snooze: bool


class UpdatePromptDialog(QtWidgets.QDialog):
    def __init__(self, info: UpdateInfo, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._info = info
        self._update_now = False
        self.setWindowTitle(t("title_update"))
        self.setModal(True)
        self.resize(740, 520)

        layout = QtWidgets.QVBoxLayout(self)

        header = QtWidgets.QLabel(t("msg_update_available_fmt", version=info.version))
        header.setWordWrap(True)
        layout.addWidget(header)

        notes_label = QtWidgets.QLabel(t("update_release_notes"))
        layout.addWidget(notes_label)

        self.notes_view = QtWidgets.QTextBrowser()
        self.notes_view.setOpenExternalLinks(True)
        if info.notes:
            try:
                self.notes_view.setMarkdown(info.notes)
            except Exception:
                self.notes_view.setPlainText(info.notes)
        else:
            self.notes_view.setText(t("release_notes_empty"))
        layout.addWidget(self.notes_view, 1)

        self.snooze_checkbox = QtWidgets.QCheckBox(t("update_ignore_7days"))
        layout.addWidget(self.snooze_checkbox)

        button_row = QtWidgets.QHBoxLayout()
        button_row.addStretch(1)
        self.open_browser_button = QtWidgets.QPushButton(t("btn_open_browser"))
        self.open_browser_button.setEnabled(bool(info.url))
        self.open_browser_button.clicked.connect(self._open_in_browser)
        button_row.addWidget(self.open_browser_button)

        self.update_button = QtWidgets.QPushButton(t("btn_update_now"))
        self.update_button.clicked.connect(self._on_update_now)
        button_row.addWidget(self.update_button)

        self.cancel_button = QtWidgets.QPushButton(t("btn_update_no"))
        self.cancel_button.clicked.connect(self.reject)
        button_row.addWidget(self.cancel_button)
        layout.addLayout(button_row)

    def _on_update_now(self) -> None:
        self._update_now = True
        self.accept()

    def _open_in_browser(self) -> None:
        if not self._info.url:
            return
        QtGui.QDesktopServices.openUrl(QtCore.QUrl(self._info.url))

    def result_state(self) -> UpdatePromptResult:
        return UpdatePromptResult(
            update_now=self._update_now,
            snooze=self.snooze_checkbox.isChecked(),
        )

    @staticmethod
    def get_result(info: UpdateInfo, parent: QtWidgets.QWidget | None = None) -> UpdatePromptResult:
        dialog = UpdatePromptDialog(info, parent)
        dialog.exec()
        return dialog.result_state()
