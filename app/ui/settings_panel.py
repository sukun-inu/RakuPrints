from __future__ import annotations

from typing import Dict

from PySide6 import QtCore, QtWidgets

from app.model.print_job import DuplexMode
from app.i18n import t, language_label, LANGUAGES


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
    language_changed = QtCore.Signal(str)
    update_check_changed = QtCore.Signal(bool)
    auto_update_changed = QtCore.Signal(bool)

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)

        self.printer_group = QtWidgets.QGroupBox()
        printer_layout = QtWidgets.QVBoxLayout(self.printer_group)

        self.use_default_radio = QtWidgets.QRadioButton()
        self.select_printer_radio = QtWidgets.QRadioButton()
        self.select_printer_button = QtWidgets.QPushButton()
        self.select_printer_button.setEnabled(False)
        self.printer_settings_button = QtWidgets.QPushButton()

        printer_layout.addWidget(self.use_default_radio)
        printer_layout.addWidget(self.select_printer_radio)
        printer_layout.addWidget(self.select_printer_button)
        printer_layout.addWidget(self.printer_settings_button)

        copies_layout = QtWidgets.QHBoxLayout()
        self.copies_label = QtWidgets.QLabel()
        self.copies_spin = QtWidgets.QSpinBox()
        self.copies_spin.setMinimum(1)
        self.copies_spin.setMaximum(999)
        copies_layout.addWidget(self.copies_label)
        copies_layout.addWidget(self.copies_spin)
        copies_layout.addStretch(1)

        duplex_layout = QtWidgets.QHBoxLayout()
        self.duplex_label = QtWidgets.QLabel()
        self.duplex_combo = QtWidgets.QComboBox()
        duplex_layout.addWidget(self.duplex_label)
        duplex_layout.addWidget(self.duplex_combo)
        duplex_layout.addStretch(1)

        paper_layout = QtWidgets.QHBoxLayout()
        self.paper_label = QtWidgets.QLabel()
        self.paper_combo = QtWidgets.QComboBox()
        paper_layout.addWidget(self.paper_label)
        paper_layout.addWidget(self.paper_combo)
        paper_layout.addStretch(1)

        excel_layout = QtWidgets.QHBoxLayout()
        self.excel_label = QtWidgets.QLabel()
        self.excel_orientation_combo = QtWidgets.QComboBox()
        excel_layout.addWidget(self.excel_label)
        excel_layout.addWidget(self.excel_orientation_combo)
        excel_layout.addStretch(1)

        printer_layout.addLayout(copies_layout)
        printer_layout.addLayout(duplex_layout)
        printer_layout.addLayout(paper_layout)
        printer_layout.addLayout(excel_layout)

        self.rules_group = QtWidgets.QGroupBox()
        rules_layout = QtWidgets.QVBoxLayout(self.rules_group)
        self.rules_table = QtWidgets.QTableWidget(0, 2)
        self.rules_table.horizontalHeader().setStretchLastSection(True)
        self.rules_table.verticalHeader().setVisible(False)
        self.rules_table.verticalHeader().setDefaultSectionSize(28)
        self.rules_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.rules_table.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.rules_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.rules_table.setAlternatingRowColors(True)
        rules_layout.addWidget(self.rules_table)

        rules_button_row = QtWidgets.QHBoxLayout()
        self.rule_add_button = QtWidgets.QPushButton()
        self.rule_remove_button = QtWidgets.QPushButton()
        rules_button_row.addWidget(self.rule_add_button)
        rules_button_row.addWidget(self.rule_remove_button)
        rules_button_row.addStretch(1)
        rules_layout.addLayout(rules_button_row)

        self.theme_group = QtWidgets.QGroupBox()
        theme_layout = QtWidgets.QVBoxLayout(self.theme_group)

        theme_row = QtWidgets.QHBoxLayout()
        self.theme_label = QtWidgets.QLabel()
        self.theme_combo = QtWidgets.QComboBox()
        theme_row.addWidget(self.theme_label)
        theme_row.addWidget(self.theme_combo)
        theme_row.addStretch(1)
        theme_layout.addLayout(theme_row)

        language_row = QtWidgets.QHBoxLayout()
        self.language_label = QtWidgets.QLabel()
        self.language_combo = QtWidgets.QComboBox()
        language_row.addWidget(self.language_label)
        language_row.addWidget(self.language_combo)
        language_row.addStretch(1)
        theme_layout.addLayout(language_row)

        self.update_group = QtWidgets.QGroupBox()
        update_layout = QtWidgets.QVBoxLayout(self.update_group)
        self.update_check_box = QtWidgets.QCheckBox()
        self.auto_update_box = QtWidgets.QCheckBox()
        update_layout.addWidget(self.update_check_box)
        update_layout.addWidget(self.auto_update_box)

        layout.addWidget(self.printer_group)
        layout.addWidget(self.rules_group)
        layout.addWidget(self.theme_group)
        layout.addWidget(self.update_group)
        layout.addStretch(1)

        self._printers: list[str] = []

        self.use_default_radio.toggled.connect(self._on_default_toggled)
        self.select_printer_radio.toggled.connect(self._on_select_toggled)
        self.select_printer_button.clicked.connect(self.select_printer_clicked)
        self.printer_settings_button.clicked.connect(self.printer_settings_clicked)
        self.copies_spin.valueChanged.connect(self.copies_changed)
        self.duplex_combo.currentIndexChanged.connect(self._on_duplex_changed)
        self.paper_combo.currentIndexChanged.connect(self._on_paper_changed)
        self.excel_orientation_combo.currentIndexChanged.connect(self._on_excel_orientation_changed)
        self.theme_combo.currentIndexChanged.connect(self._on_theme_changed)
        self.language_combo.currentIndexChanged.connect(self.on_language_changed)
        self.update_check_box.toggled.connect(self.update_check_changed)
        self.auto_update_box.toggled.connect(self.auto_update_changed)
        self.rules_table.doubleClicked.connect(self._on_rule_double_clicked)
        self.rule_add_button.clicked.connect(self._on_rule_add)
        self.rule_remove_button.clicked.connect(self._on_rule_remove)

        self.retranslate()

    def retranslate(self) -> None:
        self.printer_group.setTitle(t("settings_printer_group"))
        self.rules_group.setTitle(t("settings_rules_group"))
        self.theme_group.setTitle(t("settings_display_group"))
        self.language_label.setText(t("settings_language"))
        self.update_group.setTitle(t("settings_update_group"))

        self.use_default_radio.setText(t("settings_use_default"))
        self.select_printer_radio.setText(t("settings_select_printer"))
        self.select_printer_button.setText(t("settings_select_printer_button"))
        self.printer_settings_button.setText(t("settings_printer_settings_button"))

        self.copies_label.setText(t("settings_copies"))
        self.duplex_label.setText(t("settings_duplex"))
        self.paper_label.setText(t("settings_paper_size"))
        self.excel_label.setText(t("settings_excel_orientation"))
        self.theme_label.setText(t("settings_theme"))
        self.language_label.setText(t("settings_language"))
        self.update_check_box.setText(t("settings_update_check"))
        self.auto_update_box.setText(t("settings_update_auto"))

        self._refresh_duplex_items()
        self._refresh_excel_orientation_items()
        self._refresh_theme_items()
        self._refresh_language_items()

        self.rules_table.setHorizontalHeaderLabels([t("settings_rules_type"), t("settings_rules_printer")])
        self.rule_add_button.setText(t("settings_add"))
        self.rule_remove_button.setText(t("settings_remove"))

    def _refresh_duplex_items(self) -> None:
        current = self.duplex_combo.currentData()
        with QtCore.QSignalBlocker(self.duplex_combo):
            self.duplex_combo.clear()
            self.duplex_combo.addItem(t("duplex_off"), DuplexMode.OFF.value)
            self.duplex_combo.addItem(t("duplex_long"), DuplexMode.LONG_EDGE.value)
            self.duplex_combo.addItem(t("duplex_short"), DuplexMode.SHORT_EDGE.value)
        if current:
            index = self.duplex_combo.findData(current)
            if index >= 0:
                self.duplex_combo.setCurrentIndex(index)

    def _refresh_excel_orientation_items(self) -> None:
        current = self.excel_orientation_combo.currentData()
        with QtCore.QSignalBlocker(self.excel_orientation_combo):
            self.excel_orientation_combo.clear()
            self.excel_orientation_combo.addItem(t("settings_excel_auto"), "auto")
            self.excel_orientation_combo.addItem(t("settings_excel_ask"), "ask")
            self.excel_orientation_combo.addItem(t("settings_excel_none"), "none")
        if current:
            index = self.excel_orientation_combo.findData(current)
            if index >= 0:
                self.excel_orientation_combo.setCurrentIndex(index)

    def _refresh_theme_items(self) -> None:
        current = self.theme_combo.currentData()
        with QtCore.QSignalBlocker(self.theme_combo):
            self.theme_combo.clear()
            self.theme_combo.addItem(t("settings_theme_system"), "system")
            self.theme_combo.addItem(t("settings_theme_light"), "light")
            self.theme_combo.addItem(t("settings_theme_dark"), "dark")
        if current:
            index = self.theme_combo.findData(current)
            if index >= 0:
                self.theme_combo.setCurrentIndex(index)

    def _refresh_language_items(self) -> None:
        current = self.language_combo.currentData()
        with QtCore.QSignalBlocker(self.language_combo):
            self.language_combo.clear()
            for code in ["system", "ja", "en", "ko", "zh"]:
                self.language_combo.addItem(language_label(code), code)
        if current:
            index = self.language_combo.findData(current)
            if index >= 0:
                self.language_combo.setCurrentIndex(index)

    def set_settings(
        self,
        use_default: bool,
        copies: int,
        duplex: str,
        theme_mode: str,
        selected_printer: str,
        default_printer: str,
        excel_orientation_mode: str,
        language_mode: str,
        update_check_enabled: bool,
        auto_update_enabled: bool,
    ) -> None:
        with QtCore.QSignalBlocker(self.use_default_radio), QtCore.QSignalBlocker(self.select_printer_radio), \
                QtCore.QSignalBlocker(self.copies_spin), QtCore.QSignalBlocker(self.duplex_combo), \
                QtCore.QSignalBlocker(self.theme_combo), QtCore.QSignalBlocker(self.language_combo), \
                QtCore.QSignalBlocker(self.update_check_box), QtCore.QSignalBlocker(self.auto_update_box):
            self.use_default_radio.setChecked(use_default)
            self.select_printer_radio.setChecked(not use_default)
            self.select_printer_button.setEnabled(not use_default)
            self._update_printer_button(use_default, selected_printer, default_printer)
            self.copies_spin.setValue(copies)

            duplex_index = self.duplex_combo.findData(duplex)
            if duplex_index >= 0:
                self.duplex_combo.setCurrentIndex(duplex_index)

            theme_index = self.theme_combo.findData(theme_mode)
            if theme_index >= 0:
                self.theme_combo.setCurrentIndex(theme_index)

            excel_index = self.excel_orientation_combo.findData(excel_orientation_mode)
            if excel_index >= 0:
                self.excel_orientation_combo.setCurrentIndex(excel_index)

            language_index = self.language_combo.findData(language_mode)
            if language_index >= 0:
                self.language_combo.setCurrentIndex(language_index)

            self.update_check_box.setChecked(update_check_enabled)
            self.auto_update_box.setChecked(auto_update_enabled)

    def set_paper_sizes(self, sizes: list[str], current: str, enabled: bool, tooltip: str = "") -> None:
        with QtCore.QSignalBlocker(self.paper_combo):
            self.paper_combo.clear()
            self.paper_combo.addItem(t("paper_size_default"), "")
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
            combo.addItem(t("label_auto"), "")
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
                combo.setToolTip(t("rules_auto_tooltip_fmt", printer=auto_printer_label))
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

    def on_language_changed(self) -> None:
        """Handle language change event."""
        mode = str(self.language_combo.currentData() or "system")
        self.language_changed.emit(mode)

    def _update_printer_button(self, use_default: bool, selected_printer: str, default_printer: str) -> None:
        if use_default:
            label = default_printer or "-"
            self.select_printer_button.setText(t("printer_button_default_fmt", label=label))
        else:
            label = selected_printer or "-"
            self.select_printer_button.setText(t("printer_button_current_fmt", label=label))

    def _on_paper_changed(self) -> None:
        value = str(self.paper_combo.currentData() or "")
        self.paper_size_changed.emit(value)

    def _on_duplex_changed(self) -> None:
        value = str(self.duplex_combo.currentData() or DuplexMode.OFF.value)
        self.duplex_changed.emit(value)

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
            t("rules_add_title"),
            t("rules_add_prompt"),
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



