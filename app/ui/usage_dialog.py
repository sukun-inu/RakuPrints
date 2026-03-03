from __future__ import annotations

from PySide6 import QtWidgets

from app.i18n import t


class UsageDialog(QtWidgets.QDialog):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.resize(760, 560)

        layout = QtWidgets.QVBoxLayout(self)
        self.body_view = QtWidgets.QTextBrowser()
        self.body_view.setOpenExternalLinks(True)
        layout.addWidget(self.body_view, 1)

        button_row = QtWidgets.QHBoxLayout()
        button_row.addStretch(1)
        self.close_button = QtWidgets.QPushButton()
        self.close_button.clicked.connect(self.accept)
        button_row.addWidget(self.close_button)
        layout.addLayout(button_row)

        self.retranslate()

    def retranslate(self) -> None:
        self.setWindowTitle(t("usage_title"))
        self.body_view.setMarkdown(t("usage_body"))
        self.close_button.setText(t("btn_close"))
