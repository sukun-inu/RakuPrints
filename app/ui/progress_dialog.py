from __future__ import annotations

from PySide6 import QtCore, QtWidgets

from app.i18n import t


class ProgressDialog(QtWidgets.QDialog):
    cancel_requested = QtCore.Signal()

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setModal(True)
        self.resize(420, 160)

        layout = QtWidgets.QVBoxLayout(self)
        self.status_label = QtWidgets.QLabel()
        layout.addWidget(self.status_label)

        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setMinimum(0)
        layout.addWidget(self.progress_bar)

        self.cancel_button = QtWidgets.QPushButton()
        self.cancel_button.clicked.connect(self.cancel_requested)
        layout.addWidget(self.cancel_button)

        self.retranslate()

    def retranslate(self) -> None:
        self.setWindowTitle(t("progress_title"))
        self.status_label.setText(t("progress_preparing"))
        self.cancel_button.setText(t("progress_cancel"))

    def set_total(self, total: int) -> None:
        self.progress_bar.setMaximum(total)

    def update_progress(self, completed: int, total: int, current: str) -> None:
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(completed)
        if current:
            self.status_label.setText(t("progress_printing_fmt", current=current))
        else:
            self.status_label.setText(t("progress_printing"))

    def set_finished(self, cancelled: bool) -> None:
        if cancelled:
            self.status_label.setText(t("progress_cancelled"))
        else:
            self.status_label.setText(t("progress_done"))
        self.cancel_button.setEnabled(False)
