from __future__ import annotations

from typing import Dict

from PySide6 import QtCore, QtWidgets

from app.model.print_job import DuplexMode


class SettingsPanel(QtWidgets.QWidget):
    use_default_changed = QtCore.Signal(bool)
    select_printer_clicked = QtCore.Signal()
    printer_settings_clicked = QtCore.Signal()
    copies_changed = QtCore.Signal(int)
    duplex_changed = QtCore.Signal(str)
    excel_orientation_mode_changed = QtCore.Signal(str)
    rule_printer_changed = QtCore.Signal(str, str)
    rule_add_requested = QtCore.Signal(str)
    rule_remove_requested = QtCore.Signal(list)
    theme_changed = QtCore.Signal(str)
    paper_size_changed = QtCore.Signal(str)

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)

        self.printer_group = QtWidgets.QGroupBox("プリンター設定")
        printer_layout = QtWidgets.QVBoxLayout(self.printer_group)

        self.use_default_radio = QtWidgets.QRadioButton("既定のプリンターを使う")
        self.select_printer_radio = QtWidgets.QRadioButton("プリンターを指定する")
        self.select_printer_button = QtWidgets.QPushButton("プリンターを選ぶ")
        self.select_printer_button.setEnabled(False)
        self.printer_settings_button = QtWidgets.QPushButton("プリンターの詳細設定")

        printer_layout.addWidget(self.use_default_radio)
        printer_layout.addWidget(self.select_printer_radio)
        printer_layout.addWidget(self.select_printer_button)
        printer_layout.addWidget(self.printer_settings_button)

        copies_layout = QtWidgets.QHBoxLayout()
        copies_label = QtWidgets.QLabel("部数:")
        self.copies_spin = QtWidgets.QSpinBox()
        self.copies_spin.setMinimum(1)
        self.copies_spin.setMaximum(999)
        copies_layout.addWidget(copies_label)
        copies_layout.addWidget(self.copies_spin)
        copies_layout.addStretch(1)

        duplex_layout = QtWidgets.QHBoxLayout()
        duplex_label = QtWidgets.QLabel("両面印刷:")
        self.duplex_combo = QtWidgets.QComboBox()
        self.duplex_combo.addItems([mode.value for mode in DuplexMode])
        duplex_layout.addWidget(duplex_label)
        duplex_layout.addWidget(self.duplex_combo)
        duplex_layout.addStretch(1)

        paper_layout = QtWidgets.QHBoxLayout()
        paper_label = QtWidgets.QLabel("用紙サイズ:")
        self.paper_combo = QtWidgets.QComboBox()
        paper_layout.addWidget(paper_label)
        paper_layout.addWidget(self.paper_combo)
        paper_layout.addStretch(1)

        excel_layout = QtWidgets.QHBoxLayout()
        excel_label = QtWidgets.QLabel("Excelの向き:")
        self.excel_orientation_combo = QtWidgets.QComboBox()
        self.excel_orientation_combo.addItem("自動で調整", "auto")
        self.excel_orientation_combo.addItem("確認してから", "ask")
        self.excel_orientation_combo.addItem("変更しない", "none")
        excel_layout.addWidget(excel_label)
        excel_layout.addWidget(self.excel_orientation_combo)
        excel_layout.addStretch(1)

        printer_layout.addLayout(copies_layout)
        printer_layout.addLayout(duplex_layout)
        printer_layout.addLayout(paper_layout)
        printer_layout.addLayout(excel_layout)

        self.rules_group = QtWidgets.QGroupBox("ファイルの種類ごとのプリンター")
        rules_layout = QtWidgets.QVBoxLayout(self.rules_group)
        self.rules_table = QtWidgets.QTableWidget(0, 2)
        self.rules_table.setHorizontalHeaderLabels(["種類（PDFなど）", "プリンター"])
        self.rules_table.horizontalHeader().setStretchLastSection(True)
        self.rules_table.verticalHeader().setVisible(False)
        self.rules_table.verticalHeader().setDefaultSectionSize(28)
        self.rules_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.rules_table.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.rules_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.rules_table.setAlternatingRowColors(True)
        rules_layout.addWidget(self.rules_table)

        rules_button_row = QtWidgets.QHBoxLayout()
        self.rule_add_button = QtWidgets.QPushButton("追加")
        self.rule_remove_button = QtWidgets.QPushButton("削除")
        rules_button_row.addWidget(self.rule_add_button)
        rules_button_row.addWidget(self.rule_remove_button)
        rules_button_row.addStretch(1)
        rules_layout.addLayout(rules_button_row)

        self.theme_group = QtWidgets.QGroupBox("表示")
        theme_layout = QtWidgets.QHBoxLayout(self.theme_group)
        theme_label = QtWidgets.QLabel("テーマ:")
        self.theme_combo = QtWidgets.QComboBox()
        self.theme_combo.addItem("OSに合わせる", "system")
        self.theme_combo.addItem("ライト", "light")
        self.theme_combo.addItem("ダーク", "dark")
        theme_layout.addWidget(theme_label)
        theme_layout.addWidget(self.theme_combo)
        theme_layout.addStretch(1)

        layout.addWidget(self.printer_group)
        layout.addWidget(self.rules_group)
        layout.addWidget(self.theme_group)
        layout.addStretch(1)

        self._printers: list[str] = []

        self.use_default_radio.toggled.connect(self._on_default_toggled)
        self.select_printer_radio.toggled.connect(self._on_select_toggled)
        self.select_printer_button.clicked.connect(self.select_printer_clicked)
        self.printer_settings_button.clicked.connect(self.printer_settings_clicked)
        self.copies_spin.valueChanged.connect(self.copies_changed)
        self.duplex_combo.currentTextChanged.connect(self.duplex_changed)
        self.paper_combo.currentIndexChanged.connect(self._on_paper_changed)
        self.excel_orientation_combo.currentIndexChanged.connect(self._on_excel_orientation_changed)
        self.theme_combo.currentIndexChanged.connect(self._on_theme_changed)
        self.rules_table.doubleClicked.connect(self._on_rule_double_clicked)
        self.rule_add_button.clicked.connect(self._on_rule_add)
        self.rule_remove_button.clicked.connect(self._on_rule_remove)

    def set_settings(
        self,
        use_default: bool,
        copies: int,
        duplex: str,
        theme_mode: str,
        selected_printer: str,
        default_printer: str,
        excel_orientation_mode: str,
    ) -> None:
        with QtCore.QSignalBlocker(self.use_default_radio), QtCore.QSignalBlocker(self.select_printer_radio), \
                QtCore.QSignalBlocker(self.copies_spin), QtCore.QSignalBlocker(self.duplex_combo), \
                QtCore.QSignalBlocker(self.theme_combo):
            self.use_default_radio.setChecked(use_default)
            self.select_printer_radio.setChecked(not use_default)
            self.select_printer_button.setEnabled(not use_default)
            self._update_printer_button(use_default, selected_printer, default_printer)
            self.copies_spin.setValue(copies)
            index = self.duplex_combo.findText(duplex)
            if index >= 0:
                self.duplex_combo.setCurrentIndex(index)
            theme_index = self.theme_combo.findData(theme_mode)
            if theme_index >= 0:
                self.theme_combo.setCurrentIndex(theme_index)
            excel_index = self.excel_orientation_combo.findData(excel_orientation_mode)
            if excel_index >= 0:
                self.excel_orientation_combo.setCurrentIndex(excel_index)

    def set_paper_sizes(self, sizes: list[str], current: str, enabled: bool, tooltip: str = "") -> None:
        with QtCore.QSignalBlocker(self.paper_combo):
            self.paper_combo.clear()
            self.paper_combo.addItem("プリンターの既定", "")
            for name in sizes:
                self.paper_combo.addItem(name, name)
            index = self.paper_combo.findData(current)
            if index >= 0:
                self.paper_combo.setCurrentIndex(index)
            else:
                self.paper_combo.setCurrentIndex(0)
        self.paper_combo.setEnabled(enabled)
        self.paper_combo.setToolTip(tooltip)

    def set_printers(self, printers: list[str]) -> None:
        self._printers = list(printers)

    def set_rules(self, rules: Dict[str, dict], auto_printer_label: str = "") -> None:
        self.rules_table.setRowCount(0)
        for row, (ext, rule) in enumerate(sorted(rules.items())):
            self.rules_table.insertRow(row)
            ext_item = QtWidgets.QTableWidgetItem(ext)
            ext_item.setFlags(ext_item.flags() & ~QtCore.Qt.ItemIsEditable)
            self.rules_table.setItem(row, 0, ext_item)
            combo = QtWidgets.QComboBox()
            combo.addItem("自動（既定）", "")
            for name in self._printers:
                combo.addItem(name, name)
            current_printer = rule.get("printer", "")
            index = combo.findData(current_printer)
            if index == -1 and current_printer:
                combo.addItem(current_printer, current_printer)
                index = combo.findData(current_printer)
            with QtCore.QSignalBlocker(combo):
                if index >= 0:
                    combo.setCurrentIndex(index)
            if auto_printer_label:
                combo.setToolTip(f"自動（既定）: {auto_printer_label}")
            combo.currentTextChanged.connect(
                lambda _text, ext=ext, combo=combo: self.rule_printer_changed.emit(
                    ext, str(combo.currentData() or "")
                )
            )
            self.rules_table.setCellWidget(row, 1, combo)

    def _on_default_toggled(self, checked: bool) -> None:
        if checked:
            self.select_printer_button.setEnabled(False)
            self.use_default_changed.emit(True)

    def _on_select_toggled(self, checked: bool) -> None:
        if checked:
            self.select_printer_button.setEnabled(True)
            self.use_default_changed.emit(False)

    def _on_theme_changed(self) -> None:
        mode = str(self.theme_combo.currentData() or "system")
        self.theme_changed.emit(mode)

    def _update_printer_button(self, use_default: bool, selected_printer: str, default_printer: str) -> None:
        if use_default:
            label = default_printer or "未検出"
            self.select_printer_button.setText(f"プリンターを選ぶ（既定: {label}）")
        else:
            label = selected_printer or "未選択"
            self.select_printer_button.setText(f"プリンターを選ぶ（現在: {label}）")

    def _on_paper_changed(self) -> None:
        value = str(self.paper_combo.currentData() or "")
        self.paper_size_changed.emit(value)

    def _on_excel_orientation_changed(self) -> None:
        mode = str(self.excel_orientation_combo.currentData() or "auto")
        self.excel_orientation_mode_changed.emit(mode)

    def _on_rule_double_clicked(self, index: QtCore.QModelIndex) -> None:
        if index.column() != 1:
            return
        ext_item = self.rules_table.item(index.row(), 0)
        if not ext_item:
            return
        from app.ui.printer_selector import PrinterSelectorDialog

        selected, ok = PrinterSelectorDialog.get_printer(self, "")
        if ok:
            self.rule_printer_changed.emit(ext_item.text(), selected)

    def _on_rule_add(self) -> None:
        text, ok = QtWidgets.QInputDialog.getText(
            self,
            "拡張子を追加",
            "拡張子を入力してください（例: .pdf）",
        )
        if not ok:
            return
        ext = text.strip().lower()
        if not ext:
            return
        if not ext.startswith("."):
            ext = f".{ext}"
        self.rule_add_requested.emit(ext)

    def _on_rule_remove(self) -> None:
        rows = [idx.row() for idx in self.rules_table.selectionModel().selectedRows()]
        if not rows:
            return
        extensions = []
        for row in rows:
            item = self.rules_table.item(row, 0)
            if item:
                extensions.append(item.text())
        if extensions:
            self.rule_remove_requested.emit(extensions)
