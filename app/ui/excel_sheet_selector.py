from __future__ import annotations

from PySide6 import QtCore, QtWidgets

from app.i18n import t


class ExcelSheetSelectorDialog(QtWidgets.QDialog):
    def __init__(
        self,
        file_name: str,
        sheet_names: list[str],
        selected: list[str] | None = None,
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.resize(420, 420)

        self._selected: set[str] = set(selected or [])

        layout = QtWidgets.QVBoxLayout(self)
        self.title = QtWidgets.QLabel()
        self.title.setWordWrap(True)
        layout.addWidget(self.title)

        self.list_widget = QtWidgets.QListWidget()
        self.list_widget.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        layout.addWidget(self.list_widget)

        for name in sheet_names:
            item = QtWidgets.QListWidgetItem(name)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
            if not self._selected or name in self._selected:
                item.setCheckState(QtCore.Qt.Checked)
            else:
                item.setCheckState(QtCore.Qt.Unchecked)
            self.list_widget.addItem(item)

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
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)

        self._file_name = file_name
        self.retranslate()

    def retranslate(self) -> None:
        self.setWindowTitle(t("excel_sheet_title"))
        self.title.setText(t("excel_sheet_prompt_fmt", file=self._file_name))
        self.select_all_button.setText(t("excel_sheet_select_all"))
        self.clear_button.setText(t("excel_sheet_clear"))

    def _select_all(self) -> None:
        for index in range(self.list_widget.count()):
            self.list_widget.item(index).setCheckState(QtCore.Qt.Checked)

    def _clear_all(self) -> None:
        for index in range(self.list_widget.count()):
            self.list_widget.item(index).setCheckState(QtCore.Qt.Unchecked)

    def _on_accept(self) -> None:
        selected = self.selected_sheets()
        if not selected:
            QtWidgets.QMessageBox.information(self, t("title_excel"), t("msg_excel_sheet_required"))
            return
        self.accept()

    def selected_sheets(self) -> list[str]:
        sheets: list[str] = []
        for index in range(self.list_widget.count()):
            item = self.list_widget.item(index)
            if item.checkState() == QtCore.Qt.Checked:
                sheets.append(item.text())
        return sheets

    @staticmethod
    def get_sheets(
        file_name: str,
        sheet_names: list[str],
        selected: list[str] | None = None,
        parent: QtWidgets.QWidget | None = None,
    ) -> tuple[list[str], bool]:
        dialog = ExcelSheetSelectorDialog(file_name, sheet_names, selected, parent)
        result = dialog.exec()
        return dialog.selected_sheets(), result == QtWidgets.QDialog.Accepted
