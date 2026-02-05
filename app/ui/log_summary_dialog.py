from __future__ import annotations

from typing import Iterable, Tuple

from PySide6 import QtCore, QtGui, QtWidgets


class LogSummaryDialog(QtWidgets.QDialog):
    def __init__(
        self,
        log_path: str,
        items: Iterable[Tuple[str, str, str]],
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._log_path = log_path
        self.setWindowTitle("ログの要約")
        self.resize(720, 420)

        layout = QtWidgets.QVBoxLayout(self)
        title = QtWidgets.QLabel("印刷エラーの要約")
        title.setStyleSheet("font-weight: bold;")
        layout.addWidget(title)

        self.table = QtWidgets.QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["ファイル名", "要約", "詳細"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(28)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        layout.addWidget(self.table)

        self._set_items(items)

        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addStretch(1)
        self.open_log_button = QtWidgets.QPushButton("詳細ログを開く")
        self.close_button = QtWidgets.QPushButton("閉じる")
        button_layout.addWidget(self.open_log_button)
        button_layout.addWidget(self.close_button)
        layout.addLayout(button_layout)

        self.open_log_button.clicked.connect(self._open_log)
        self.close_button.clicked.connect(self.accept)

    def _set_items(self, items: Iterable[Tuple[str, str, str]]) -> None:
        rows = list(items)
        self.table.setRowCount(0)
        if not rows:
            self.table.setRowCount(1)
            item = QtWidgets.QTableWidgetItem("失敗した印刷はありません。")
            item.setFlags(item.flags() & ~QtCore.Qt.ItemIsSelectable)
            self.table.setItem(0, 0, item)
            self.table.setSpan(0, 0, 1, 3)
            return

        for row_index, (file_name, summary, detail) in enumerate(rows):
            self.table.insertRow(row_index)
            self.table.setItem(row_index, 0, QtWidgets.QTableWidgetItem(file_name))
            self.table.setItem(row_index, 1, QtWidgets.QTableWidgetItem(summary))
            self.table.setItem(row_index, 2, QtWidgets.QTableWidgetItem(detail))

    def _open_log(self) -> None:
        QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(self._log_path))
