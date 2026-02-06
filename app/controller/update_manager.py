from __future__ import annotations

import datetime as dt
import json
import os
import re
import subprocess
import sys
import tempfile
import urllib.request
from dataclasses import dataclass
from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets

from app.app_context import AppContext
from app.i18n import t
from app.version import __version__

REPO_SLUG = "sukun-inu/RakuPrints"
GITHUB_API = "https://api.github.com/repos/{slug}/releases/latest"


@dataclass
class UpdateInfo:
    tag: str
    version: str
    url: str
    assets: list[dict]
    notes: str


def _parse_version(text: str) -> tuple[int, ...]:
    parts = re.findall(r"\d+", text)
    if not parts:
        return (0,)
    return tuple(int(p) for p in parts)


def _is_newer(tag: str, current: str) -> bool:
    return _parse_version(tag) > _parse_version(current)


def _iso_now() -> str:
    return dt.datetime.utcnow().isoformat(timespec="seconds") + "Z"


def _iso_plus_days(days: int) -> str:
    return (dt.datetime.utcnow() + dt.timedelta(days=days)).isoformat(timespec="seconds") + "Z"


def _iso_parse(value: str) -> dt.datetime | None:
    if not value:
        return None
    try:
        return dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except Exception:
        return None


class UpdateChecker(QtCore.QThread):
    checked = QtCore.Signal(object, str)

    def __init__(self, parent: QtCore.QObject | None = None) -> None:
        super().__init__(parent)

    def run(self) -> None:
        try:
            url = GITHUB_API.format(slug=REPO_SLUG)
            with urllib.request.urlopen(url, timeout=10) as response:
                payload = json.loads(response.read().decode("utf-8"))
            tag = str(payload.get("tag_name") or "")
            html_url = str(payload.get("html_url") or "")
            assets = payload.get("assets") or []
            notes = str(payload.get("body") or "")
            info = UpdateInfo(tag=tag, version=tag.lstrip("v"), url=html_url, assets=list(assets), notes=notes)
            self.checked.emit(info, "")
        except Exception as exc:
            self.checked.emit(None, str(exc))


class UpdateDownloader(QtCore.QThread):
    finished_download = QtCore.Signal(bool, str, str)

    def __init__(self, info: UpdateInfo, parent: QtCore.QObject | None = None) -> None:
        super().__init__(parent)
        self._info = info

    def run(self) -> None:
        try:
            # Find installer EXE (Setup_*.exe or *_Setup_*.exe)
            asset = None
            for item in self._info.assets:
                name = str(item.get("name") or "").lower()
                if name.endswith(".exe") and "setup" in name:
                    asset = item
                    break
            # Fallback to any .exe
            if asset is None:
                for item in self._info.assets:
                    name = str(item.get("name") or "").lower()
                    if name.endswith(".exe"):
                        asset = item
                        break
            if not asset:
                self.finished_download.emit(False, "", "No installer found")
                return
            download_url = str(asset.get("browser_download_url") or "")
            if not download_url:
                self.finished_download.emit(False, "", "No download URL")
                return

            temp_dir = Path(tempfile.mkdtemp(prefix="rakuprint_update_"))
            file_name = str(asset.get("name") or "RakuPrint_Setup.exe")
            file_path = temp_dir / file_name
            with urllib.request.urlopen(download_url, timeout=120) as response:
                file_path.write_bytes(response.read())

            self.finished_download.emit(True, str(file_path), "")
        except Exception as exc:
            self.finished_download.emit(False, "", str(exc))


class UpdateManager(QtCore.QObject):
    def __init__(self, context: AppContext, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._context = context
        self._parent = parent
        self._checker: UpdateChecker | None = None
        self._downloader: UpdateDownloader | None = None
        self._progress: QtWidgets.QProgressDialog | None = None
        self.last_notified = self._load_last_notified()

    def _load_last_notified(self) -> dt.datetime | None:
        """Load the last notification timestamp from a file."""
        path = Path(tempfile.gettempdir()) / "last_update_notify.txt"
        if path.exists():
            with path.open("r", encoding="utf-8") as file:
                return _iso_parse(file.read().strip())
        return None

    def _save_last_notified(self, timestamp: dt.datetime) -> None:
        """Save the last notification timestamp to a file."""
        path = Path(tempfile.gettempdir()) / "last_update_notify.txt"
        with path.open("w", encoding="utf-8") as file:
            file.write(timestamp.isoformat())

    def should_notify(self) -> bool:
        """Check if the user should be notified about updates."""
        if not self.last_notified:
            return True
        return dt.datetime.utcnow() >= self.last_notified + dt.timedelta(days=14)

    def notify_user(self, update_info: UpdateInfo) -> None:
        """Notify the user about the update."""
        if self.should_notify():
            # Emit a signal or show a dialog to notify the user
            self._save_last_notified(dt.datetime.utcnow())

    def check_on_startup(self) -> None:
        settings = self._context.settings
        if not settings.update_check_enabled:
            return
        snooze_until = _iso_parse(settings.update_snooze_until)
        now = dt.datetime.utcnow().replace(tzinfo=dt.timezone.utc)
        if snooze_until and snooze_until > now:
            return
        self.check_for_updates(manual=False)

    def check_for_updates(self, manual: bool = False) -> None:
        if self._checker and self._checker.isRunning():
            return
        self._checker = UpdateChecker(self)
        self._checker.checked.connect(lambda info, err: self._on_checked(info, err, manual))
        self._checker.start()
        self._context.settings.last_update_check = _iso_now()
        self._context.save_settings()

    def _on_checked(self, info: UpdateInfo | None, error: str, manual: bool) -> None:
        if error:
            if manual:
                QtWidgets.QMessageBox.warning(
                    self._parent,
                    t("title_update"),
                    t("msg_update_error_fmt", error=error),
                )
            return
        if info is None or not info.tag:
            if manual:
                QtWidgets.QMessageBox.information(
                    self._parent,
                    t("title_update"),
                    t("msg_update_none"),
                )
            return
        if not _is_newer(info.tag, __version__):
            if manual:
                QtWidgets.QMessageBox.information(
                    self._parent,
                    t("title_update"),
                    t("msg_update_none"),
                )
            return
        if self._context.settings.auto_update_enabled and not manual:
            self._download_update(info)
            return
        self._prompt_update(info)

    def _prompt_update(self, info: UpdateInfo) -> None:
        dialog = QtWidgets.QMessageBox(self._parent)
        dialog.setWindowTitle(t("title_update"))
        dialog.setIcon(QtWidgets.QMessageBox.Information)
        dialog.setText(t("msg_update_available_fmt", version=info.version))
        update_now = dialog.addButton(t("btn_update_now"), QtWidgets.QMessageBox.AcceptRole)
        remind = dialog.addButton(t("btn_update_later"), QtWidgets.QMessageBox.ActionRole)
        skip = dialog.addButton(t("btn_update_skip"), QtWidgets.QMessageBox.RejectRole)
        dialog.exec()

        clicked = dialog.clickedButton()
        if clicked == update_now:
            self._download_update(info)
        elif clicked == remind:
            self._context.update_setting(update_snooze_until=_iso_plus_days(14))
        else:
            return

    def _download_update(self, info: UpdateInfo) -> None:
        if self._downloader and self._downloader.isRunning():
            return
        self._progress = QtWidgets.QProgressDialog(t("msg_update_downloading"), "", 0, 0, self._parent)
        self._progress.setWindowTitle(t("title_update"))
        self._progress.setCancelButton(None)
        self._progress.setWindowModality(QtCore.Qt.ApplicationModal)
        self._progress.show()

        self._downloader = UpdateDownloader(info, self)
        self._downloader.finished_download.connect(self._on_download_finished)
        self._downloader.start()

    def _on_download_finished(self, success: bool, installer_path: str, error: str) -> None:
        if self._progress:
            self._progress.close()
            self._progress = None
        if not success:
            QtWidgets.QMessageBox.warning(
                self._parent,
                t("title_update"),
                t("msg_update_error_fmt", error=error or "Download failed"),
            )
            return
        if not installer_path or not Path(installer_path).exists():
            QtWidgets.QMessageBox.warning(
                self._parent,
                t("title_update"),
                t("msg_update_error_fmt", error="Installer not found"),
            )
            return

        # Ask user to run installer
        result = QtWidgets.QMessageBox.question(
            self._parent,
            t("title_update"),
            t("msg_update_apply"),
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
        )
        if result != QtWidgets.QMessageBox.Yes:
            return

        try:
            # Run installer and quit app
            subprocess.Popen([installer_path], shell=True)
            QtCore.QTimer.singleShot(500, QtWidgets.QApplication.quit)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(
                self._parent,
                t("title_update"),
                t("msg_update_error_fmt", error=str(exc)),
            )

