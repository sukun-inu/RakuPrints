from __future__ import annotations

from typing import Iterable, Tuple

from PySide6 import QtCore, QtWidgets

from app.i18n import t


class ExcelOrientationDialog(QtWidgets.QDialog):
    def __init__(
        self,
        items: Iterable[Tuple[str, str, str, str]],
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.resize(520, 360)

        self._items = list(items)

        layout = QtWidgets.QVBoxLayout(self)
        self.label = QtWidgets.QLabel()
        self.label.setWordWrap(True)
        layout.addWidget(self.label)

        self.table = QtWidgets.QTableWidget(0, 4)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.table)

        self._populate()

        button_row = QtWidgets.QHBoxLayout()
        self.select_all_button = QtWidgets.QPushButton()
        self.clear_button = QtWidgets.QPushButton()
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

        self.retranslate()

    def retranslate(self) -> None:
        self.setWindowTitle(t("excel_orientation_title"))
        self.label.setText(t("excel_orientation_prompt"))
        self.table.setHorizontalHeaderLabels(
            [
                t("excel_orientation_file"),
                t("excel_orientation_recommend"),
                t("excel_orientation_reason"),
                t("excel_orientation_apply"),
            ]
        )
        self.select_all_button.setText(t("excel_orientation_select_all"))
        self.clear_button.setText(t("excel_orientation_clear"))

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
            item = self.table.item(row, 3)
            if item:
                item.setCheckState(QtCore.Qt.Checked)

    def _clear_all(self) -> None:
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 3)
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
