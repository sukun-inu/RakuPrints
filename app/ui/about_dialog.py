from __future__ import annotations

from PySide6 import QtWidgets


class AboutDialog(QtWidgets.QMessageBox):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("このアプリについて")
        self.setText("らーく印刷\n印刷ジョブ管理ツール")
        self.setInformativeText(
            "Windows + Office 環境でまとめて印刷できるように作られています。\n"
            "Copyright © 2026 Hibiki Suzuki"
        )
        self.setStandardButtons(QtWidgets.QMessageBox.Ok)
