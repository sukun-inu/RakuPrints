from __future__ import annotations

import json
import urllib.request
from dataclasses import dataclass
from typing import List

from PySide6 import QtCore, QtGui, QtWidgets

from app.i18n import t


REPO_SLUG = "sukun-inu/RakuPrints"
GITHUB_RELEASES_API = "https://api.github.com/repos/{slug}/releases?per_page=30"


@dataclass
class ReleaseNote:
    tag: str
    name: str
    published_at: str
    body: str
    url: str


class ReleaseNotesFetcher(QtCore.QThread):
    finished_fetch = QtCore.Signal(list, str)

    def run(self) -> None:
        try:
            url = GITHUB_RELEASES_API.format(slug=REPO_SLUG)
            request = urllib.request.Request(
                url,
                headers={
                    "Accept": "application/vnd.github+json",
                    "User-Agent": "RakuPrints",
                },
            )
            with urllib.request.urlopen(request, timeout=12) as response:
                payload = json.loads(response.read().decode("utf-8"))
            notes: list[ReleaseNote] = []
            if isinstance(payload, list):
                for item in payload:
                    tag = str(item.get("tag_name") or "")
                    name = str(item.get("name") or "")
                    published_at = str(item.get("published_at") or "")
                    body = str(item.get("body") or "")
                    url = str(item.get("html_url") or "")
                    notes.append(
                        ReleaseNote(
                            tag=tag,
                            name=name,
                            published_at=published_at,
                            body=body,
                            url=url,
                        )
                    )
            self.finished_fetch.emit(notes, "")
        except Exception as exc:
            self.finished_fetch.emit([], str(exc))


class ReleaseNotesDialog(QtWidgets.QDialog):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.resize(860, 520)
        self._fetcher: ReleaseNotesFetcher | None = None
        self._notes: List[ReleaseNote] = []

        layout = QtWidgets.QVBoxLayout(self)
        self.status_label = QtWidgets.QLabel()
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        self.splitter = QtWidgets.QSplitter()
        self.splitter.setOrientation(QtCore.Qt.Horizontal)

        self.list_widget = QtWidgets.QListWidget()
        self.list_widget.setMinimumWidth(240)
        self.list_widget.currentRowChanged.connect(self._on_select_release)
        self.splitter.addWidget(self.list_widget)

        detail_container = QtWidgets.QWidget()
        detail_layout = QtWidgets.QVBoxLayout(detail_container)
        detail_layout.setContentsMargins(0, 0, 0, 0)

        self.title_label = QtWidgets.QLabel()
        self.title_label.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        detail_layout.addWidget(self.title_label)

        self.date_label = QtWidgets.QLabel()
        self.date_label.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        detail_layout.addWidget(self.date_label)

        self.body_view = QtWidgets.QTextBrowser()
        self.body_view.setOpenExternalLinks(True)
        detail_layout.addWidget(self.body_view, 1)

        self.splitter.addWidget(detail_container)
        self.splitter.setStretchFactor(1, 1)
        layout.addWidget(self.splitter, 1)

        button_row = QtWidgets.QHBoxLayout()
        button_row.addStretch(1)
        self.open_browser_button = QtWidgets.QPushButton()
        self.open_browser_button.setEnabled(False)
        self.open_browser_button.clicked.connect(self._open_in_browser)
        button_row.addWidget(self.open_browser_button)
        self.close_button = QtWidgets.QPushButton()
        self.close_button.clicked.connect(self.accept)
        button_row.addWidget(self.close_button)
        layout.addLayout(button_row)

        self.retranslate()
        self._start_fetch()

    def retranslate(self) -> None:
        self.setWindowTitle(t("release_notes_title"))
        self.status_label.setText(t("release_notes_loading"))
        self.title_label.setText("")
        self.date_label.setText("")
        self.body_view.setText("")
        self.open_browser_button.setText(t("btn_open_browser"))
        self.close_button.setText(t("btn_close"))

    def _start_fetch(self) -> None:
        if self._fetcher and self._fetcher.isRunning():
            return
        self._fetcher = ReleaseNotesFetcher()
        self._fetcher.finished_fetch.connect(self._on_fetched)
        self._fetcher.start()

    def _on_fetched(self, notes: list[ReleaseNote], error: str) -> None:
        if error:
            self.status_label.setText(t("release_notes_error_fmt", error=error))
            return
        self._notes = notes
        if not notes:
            self.status_label.setText(t("release_notes_empty"))
            return
        self.status_label.setText("")
        self.list_widget.clear()
        for note in notes:
            label = note.tag or note.name or "-"
            if note.published_at:
                label = f"{label} ({note.published_at[:10]})"
            self.list_widget.addItem(label)
        self.list_widget.setCurrentRow(0)

    def _on_select_release(self, row: int) -> None:
        if row < 0 or row >= len(self._notes):
            self.open_browser_button.setEnabled(False)
            self.title_label.setText("")
            self.date_label.setText("")
            self.body_view.setText("")
            return
        note = self._notes[row]
        title = note.name or note.tag or "-"
        self.title_label.setText(title)
        self.date_label.setText(note.published_at)
        if note.body:
            try:
                self.body_view.setMarkdown(note.body)
            except Exception:
                self.body_view.setPlainText(note.body)
        else:
            self.body_view.setText(t("release_notes_empty"))
        self.open_browser_button.setEnabled(bool(note.url))

    def _open_in_browser(self) -> None:
        row = self.list_widget.currentRow()
        if row < 0 or row >= len(self._notes):
            return
        url = self._notes[row].url
        if not url:
            return
        QtGui.QDesktopServices.openUrl(QtCore.QUrl(url))
