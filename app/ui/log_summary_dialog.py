from __future__ import annotations

from typing import Iterable, Tuple

from PySide6 import QtCore, QtGui, QtWidgets

from app.i18n import t


class LogSummaryDialog(QtWidgets.QDialog):
    def __init__(
        self,
        log_path: str,
        items: Iterable[Tuple[str, str, str]],
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._log_path = log_path
        self.resize(720, 420)

        layout = QtWidgets.QVBoxLayout(self)
        self.title_label = QtWidgets.QLabel()
        self.title_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.title_label)

        self.table = QtWidgets.QTableWidget(0, 3)
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
        self.open_log_button = QtWidgets.QPushButton()
        self.close_button = QtWidgets.QPushButton()
        button_layout.addWidget(self.open_log_button)
        button_layout.addWidget(self.close_button)
        layout.addLayout(button_layout)

        self.open_log_button.clicked.connect(self._open_log)
        self.close_button.clicked.connect(self.accept)

        self.retranslate()

    def retranslate(self) -> None:
        self.setWindowTitle(t("log_title"))
        self.title_label.setText(t("log_heading"))
        self.table.setHorizontalHeaderLabels([t("log_file"), t("log_summary"), t("log_detail")])
        self.open_log_button.setText(t("btn_open_log"))
        self.close_button.setText(t("btn_close"))

    def _set_items(self, items: Iterable[Tuple[str, str, str]]) -> None:
        rows = list(items)
        self.table.setRowCount(0)
        if not rows:
            self.table.setRowCount(1)
            item = QtWidgets.QTableWidgetItem(t("log_no_errors"))
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
