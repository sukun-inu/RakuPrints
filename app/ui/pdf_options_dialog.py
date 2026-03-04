from __future__ import annotations

from dataclasses import dataclass

from PySide6 import QtCore, QtWidgets

from app.i18n import t


@dataclass(frozen=True)
class PdfOptionSelection:
    scale_mode: str
    scale_percent_mode: str
    scale_percent: int
    auto_rotate: str
    center: str


class PdfOptionsDialog(QtWidgets.QDialog):
    def __init__(self, job_count: int, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(t("pdf_options_title"))
        self.setModal(True)

        layout = QtWidgets.QVBoxLayout(self)

        info_label = QtWidgets.QLabel(t("pdf_options_target_fmt", count=job_count))
        layout.addWidget(info_label)

        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignLeft)
        form.setFormAlignment(QtCore.Qt.AlignTop)

        self.scale_mode_combo = QtWidgets.QComboBox()
        self.scale_mode_combo.addItem(t("pdf_options_keep"), "keep")
        self.scale_mode_combo.addItem(t("pdf_options_default"), "default")
        self.scale_mode_combo.addItem(t("pdf_scale_auto"), "auto")
        self.scale_mode_combo.addItem(t("pdf_scale_fit"), "fit")
        self.scale_mode_combo.addItem(t("pdf_scale_shrink"), "shrink")
        self.scale_mode_combo.addItem(t("pdf_scale_none"), "none")
        form.addRow(t("pdf_options_scale_mode"), self.scale_mode_combo)

        scale_percent_row = QtWidgets.QHBoxLayout()
        self.scale_percent_combo = QtWidgets.QComboBox()
        self.scale_percent_combo.addItem(t("pdf_options_keep"), "keep")
        self.scale_percent_combo.addItem(t("pdf_options_default"), "default")
        self.scale_percent_combo.addItem(t("pdf_scale_custom"), "custom")
        self.scale_percent_spin = QtWidgets.QSpinBox()
        self.scale_percent_spin.setMinimum(10)
        self.scale_percent_spin.setMaximum(200)
        self.scale_percent_spin.setSuffix("%")
        self.scale_percent_spin.setValue(100)
        scale_percent_row.addWidget(self.scale_percent_combo)
        scale_percent_row.addWidget(self.scale_percent_spin)
        scale_percent_row.addStretch(1)
        scale_percent_widget = QtWidgets.QWidget()
        scale_percent_widget.setLayout(scale_percent_row)
        form.addRow(t("pdf_options_scale_percent"), scale_percent_widget)

        self.auto_rotate_combo = QtWidgets.QComboBox()
        self.auto_rotate_combo.addItem(t("pdf_options_keep"), "keep")
        self.auto_rotate_combo.addItem(t("pdf_options_default"), "default")
        self.auto_rotate_combo.addItem(t("pdf_options_on"), "on")
        self.auto_rotate_combo.addItem(t("pdf_options_off"), "off")
        form.addRow(t("pdf_options_auto_rotate"), self.auto_rotate_combo)

        self.center_combo = QtWidgets.QComboBox()
        self.center_combo.addItem(t("pdf_options_keep"), "keep")
        self.center_combo.addItem(t("pdf_options_default"), "default")
        self.center_combo.addItem(t("pdf_options_on"), "on")
        self.center_combo.addItem(t("pdf_options_off"), "off")
        form.addRow(t("pdf_options_center"), self.center_combo)

        layout.addLayout(form)

        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.scale_percent_combo.currentIndexChanged.connect(self._on_scale_percent_mode_changed)
        self._on_scale_percent_mode_changed()

    def _on_scale_percent_mode_changed(self) -> None:
        mode = str(self.scale_percent_combo.currentData() or "keep")
        self.scale_percent_spin.setEnabled(mode == "custom")

    def selection(self) -> PdfOptionSelection:
        return PdfOptionSelection(
            scale_mode=str(self.scale_mode_combo.currentData() or "keep"),
            scale_percent_mode=str(self.scale_percent_combo.currentData() or "keep"),
            scale_percent=int(self.scale_percent_spin.value()),
            auto_rotate=str(self.auto_rotate_combo.currentData() or "keep"),
            center=str(self.center_combo.currentData() or "keep"),
        )

    @staticmethod
    def get_options(job_count: int, parent: QtWidgets.QWidget | None = None) -> tuple[PdfOptionSelection, bool]:
        dialog = PdfOptionsDialog(job_count, parent)
        ok = dialog.exec() == QtWidgets.QDialog.Accepted
        return dialog.selection(), ok
