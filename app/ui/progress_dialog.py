from __future__ import annotations

from PySide6 import QtCore, QtWidgets


class ProgressDialog(QtWidgets.QDialog):
    cancel_requested = QtCore.Signal()

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("印刷中")
        self.setModal(True)
        self.resize(420, 160)

        layout = QtWidgets.QVBoxLayout(self)
        self.status_label = QtWidgets.QLabel("準備しています...")
        layout.addWidget(self.status_label)

        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setMinimum(0)
        layout.addWidget(self.progress_bar)

        self.cancel_button = QtWidgets.QPushButton("キャンセル")
        self.cancel_button.clicked.connect(self.cancel_requested)
        layout.addWidget(self.cancel_button)

    def set_total(self, total: int) -> None:
        self.progress_bar.setMaximum(total)

    def update_progress(self, completed: int, total: int, current: str) -> None:
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(completed)
        if current:
            self.status_label.setText(f"印刷中: {current}")
        else:
            self.status_label.setText("印刷中...")

    def set_finished(self, cancelled: bool) -> None:
        if cancelled:
            self.status_label.setText("キャンセルしました")
        else:
            self.status_label.setText("完了しました")
        self.cancel_button.setEnabled(False)
