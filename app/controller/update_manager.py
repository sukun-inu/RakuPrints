from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets

from app.app_context import AppContext
from app.i18n import t
from app.net.http_utils import urlopen
from app.version import __version__
from app.ui.update_prompt_dialog import UpdatePromptDialog

REPO_SLUG = "sukun-inu/RakuPrints"
GITHUB_API = "https://api.github.com/repos/{slug}/releases/latest"


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest().lower()


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
            with urlopen(
                url,
                timeout=10,
                headers={
                    "Accept": "application/vnd.github+json",
                    "User-Agent": "RakuPrints",
                },
            ) as response:
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
    # hash_status: "verified" | "no_checksum" | "mismatch"
    finished_download = QtCore.Signal(bool, str, str, str)

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
                self.finished_download.emit(False, "", "No installer found", "")
                return
            download_url = str(asset.get("browser_download_url") or "")
            if not download_url:
                self.finished_download.emit(False, "", "No download URL", "")
                return

            temp_dir = Path(tempfile.mkdtemp(prefix="rakuprint_update_"))
            file_name = str(asset.get("name") or "RakuPrint_Setup.exe")
            file_path = temp_dir / file_name
            with urlopen(download_url, timeout=120, headers={"User-Agent": "RakuPrints"}) as response:
                file_path.write_bytes(response.read())

            hash_status = self._verify_hash(file_path)
            self.finished_download.emit(True, str(file_path), "", hash_status)
        except Exception as exc:
            self.finished_download.emit(False, "", str(exc), "")

    def _verify_hash(self, file_path: Path) -> str:
        """Returns "verified", "mismatch", or "no_checksum"."""
        checksum_asset = None
        for item in self._info.assets:
            name = str(item.get("name") or "").upper()
            if "SHA256" in name or name in ("SHA256SUMS", "SHA256SUMS.TXT", "CHECKSUMS.TXT"):
                checksum_asset = item
                break
        if not checksum_asset:
            return "no_checksum"
        checksum_url = str(checksum_asset.get("browser_download_url") or "")
        if not checksum_url:
            return "no_checksum"
        try:
            with urlopen(checksum_url, timeout=30, headers={"User-Agent": "RakuPrints"}) as r:
                checksum_text = r.read().decode("utf-8")
        except Exception:
            return "no_checksum"
        installer_name = file_path.name.lower()
        expected_hash = None
        for line in checksum_text.splitlines():
            parts = line.strip().split()
            if len(parts) >= 2 and parts[1].lower().lstrip("*") == installer_name:
                expected_hash = parts[0].lower()
                break
        if not expected_hash:
            return "no_checksum"
        actual_hash = _sha256_file(file_path)
        return "verified" if actual_hash == expected_hash else "mismatch"


class UpdateManager(QtCore.QObject):
    def __init__(self, context: AppContext, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._context = context
        self._parent = parent
        self._checker: UpdateChecker | None = None
        self._downloader: UpdateDownloader | None = None
        self._progress: QtWidgets.QProgressDialog | None = None
        self.last_notified = self._load_last_notified()

    def _notify_path(self) -> Path:
        return self._context.data_dir / "last_update_notify.txt"

    def _load_last_notified(self) -> dt.datetime | None:
        path = self._notify_path()
        if path.exists():
            try:
                with path.open("r", encoding="utf-8") as file:
                    return _iso_parse(file.read().strip())
            except Exception:
                return None
        return None

    def _save_last_notified(self, timestamp: dt.datetime) -> None:
        path = self._notify_path()
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open("w", encoding="utf-8") as file:
                file.write(timestamp.isoformat())
        except Exception:
            pass

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
        result = UpdatePromptDialog.get_result(info, self._parent)
        if result.update_now:
            self._download_update(info)
            return
        if result.snooze:
            self._context.update_setting(update_snooze_until=_iso_plus_days(7))

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

    def _on_download_finished(self, success: bool, installer_path: str, error: str, hash_status: str) -> None:
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

        if hash_status == "mismatch":
            QtWidgets.QMessageBox.critical(
                self._parent,
                t("title_update"),
                t("msg_update_hash_mismatch"),
            )
            return

        if hash_status == "no_checksum":
            answer = QtWidgets.QMessageBox.warning(
                self._parent,
                t("title_update"),
                t("msg_update_hash_no_checksum"),
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            )
            if answer != QtWidgets.QMessageBox.Yes:
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
            subprocess.Popen([installer_path])
            QtCore.QTimer.singleShot(500, QtWidgets.QApplication.quit)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(
                self._parent,
                t("title_update"),
                t("msg_update_error_fmt", error=str(exc)),
            )
