from __future__ import annotations

from typing import Iterable, Tuple

from PySide6 import QtCore, QtWidgets


class ExcelOrientationDialog(QtWidgets.QDialog):
    def __init__(
        self,
        items: Iterable[Tuple[str, str, str, str]],
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Excelの向き確認")
        self.resize(520, 360)

        self._items = list(items)

        layout = QtWidgets.QVBoxLayout(self)
        label = QtWidgets.QLabel("Excel の向き推奨を確認してください。\n印刷時に自動で向きを調整しますか？")
        label.setWordWrap(True)
        layout.addWidget(label)

        self.table = QtWidgets.QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["ファイル", "推奨", "理由", "適用"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.table)

        self._populate()

        button_row = QtWidgets.QHBoxLayout()
        self.select_all_button = QtWidgets.QPushButton("すべて適用")
        self.clear_button = QtWidgets.QPushButton("すべて解除")
        button_row.addWidget(self.select_all_button)
        button_row.addWidget(self.clear_button)
        button_row.addStretch(1)
        layout.addLayout(button_row)

        self.select_all_button.clicked.connect(self._select_all)
        self.clear_button.clicked.connect(self._clear_all)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        layout.addWidget(buttons)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

    def _populate(self) -> None:
        self.table.setRowCount(0)
        for row, (_job_id, file_name, recommendation, reason) in enumerate(self._items):
            self.table.insertRow(row)
            self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(file_name))
            self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(recommendation))
            self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(reason))
            checkbox = QtWidgets.QTableWidgetItem("")
            checkbox.setFlags(checkbox.flags() | QtCore.Qt.ItemIsUserCheckable)
            checkbox.setCheckState(QtCore.Qt.Checked)
            self.table.setItem(row, 3, checkbox)

    def _select_all(self) -> None:
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 2)
            if item:
                item.setCheckState(QtCore.Qt.Checked)

    def _clear_all(self) -> None:
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 2)
            if item:
                item.setCheckState(QtCore.Qt.Unchecked)

    def selected_job_ids(self) -> list[str]:
        selected: list[str] = []
        for row, (job_id, _file, _rec, _reason) in enumerate(self._items):
            item = self.table.item(row, 3)
            if item and item.checkState() == QtCore.Qt.Checked:
                selected.append(job_id)
        return selected

    @staticmethod
    def get_selection(
        items: Iterable[Tuple[str, str, str, str]],
        parent: QtWidgets.QWidget | None = None,
    ) -> tuple[list[str], bool]:
        dialog = ExcelOrientationDialog(items, parent)
        result = dialog.exec()
        return dialog.selected_job_ids(), result == QtWidgets.QDialog.Accepted
