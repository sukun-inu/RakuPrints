from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import json

from PySide6 import QtCore

from app.model.print_job import DuplexMode
from app.i18n import resolve_language


DEFAULT_RULES = {
    ".pdf": {"printer": ""},
    ".doc": {"printer": ""},
    ".docx": {"printer": ""},
    ".xls": {"printer": ""},
    ".xlsx": {"printer": ""},
    ".xlsm": {"printer": ""},
    ".ppt": {"printer": ""},
    ".pptx": {"printer": ""},
}


@dataclass
class UserSettings:
    use_default_printer: bool = True
    selected_printer: str = ""
    copies: int = 1
    duplex: DuplexMode = DuplexMode.OFF
    theme_mode: str = "system"
    paper_size: str = ""
    excel_orientation_mode: str = "auto"
    pdf_scale_mode: str = "fit"
    pdf_scale_percent: int = 100
    pdf_auto_rotate: bool = True
    pdf_center: bool = True
    pdf_warn_if_clip: bool = True
    pdf_printer_scale: dict[str, int] = field(default_factory=dict)
    language_mode: str = "system"
    update_check_enabled: bool = True
    auto_update_enabled: bool = False
    update_snooze_until: str = ""
    last_update_check: str = ""
    file_list_column_widths: list[int] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "use_default_printer": self.use_default_printer,
            "selected_printer": self.selected_printer,
            "copies": self.copies,
            "duplex": self.duplex.value,
            "theme_mode": self.theme_mode,
            "paper_size": self.paper_size,
            "excel_orientation_mode": self.excel_orientation_mode,
            "pdf_scale_mode": self.pdf_scale_mode,
            "pdf_scale_percent": self.pdf_scale_percent,
            "pdf_auto_rotate": self.pdf_auto_rotate,
            "pdf_center": self.pdf_center,
            "pdf_warn_if_clip": self.pdf_warn_if_clip,
            "pdf_printer_scale": self.pdf_printer_scale,
            "language_mode": self.language_mode,
            "update_check_enabled": self.update_check_enabled,
            "auto_update_enabled": self.auto_update_enabled,
            "update_snooze_until": self.update_snooze_until,
            "last_update_check": self.last_update_check,
            "file_list_column_widths": self.file_list_column_widths,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "UserSettings":
        duplex_value = data.get("duplex", DuplexMode.OFF.value)
        legacy_map = {
            "Off": DuplexMode.OFF.value,
            "Long Edge": DuplexMode.LONG_EDGE.value,
            "Short Edge": DuplexMode.SHORT_EDGE.value,
        }
        if duplex_value in legacy_map:
            duplex_value = legacy_map[duplex_value]
        try:
            duplex = DuplexMode(duplex_value)
        except ValueError:
            duplex = DuplexMode.OFF
        language_mode = str(data.get("language_mode", "system"))
        if language_mode == "system":
            language_mode = "system"
        else:
            language_mode = resolve_language(language_mode)
        widths_raw = data.get("file_list_column_widths", [])
        widths: list[int] = []
        if isinstance(widths_raw, list):
            for item in widths_raw:
                try:
                    width = int(item)
                except (TypeError, ValueError):
                    continue
                if width > 0:
                    widths.append(width)
        scale_raw = data.get("pdf_scale_percent", 100)
        try:
            scale_percent = int(scale_raw)
        except (TypeError, ValueError):
            scale_percent = 100
        if scale_percent <= 0:
            scale_percent = 100
        printer_scale_raw = data.get("pdf_printer_scale", {})
        printer_scale: dict[str, int] = {}
        if isinstance(printer_scale_raw, dict):
            for key, value in printer_scale_raw.items():
                if not isinstance(key, str):
                    continue
                try:
                    pct = int(value)
                except (TypeError, ValueError):
                    continue
                if 10 <= pct <= 200:
                    printer_scale[key] = pct
        try:
            copies = int(data.get("copies", 1))
        except (TypeError, ValueError):
            copies = 1
        if copies < 1:
            copies = 1
        return cls(
            use_default_printer=bool(data.get("use_default_printer", True)),
            selected_printer=str(data.get("selected_printer", "")),
            copies=copies,
            duplex=duplex,
            theme_mode=str(data.get("theme_mode", "system")),
            paper_size=str(data.get("paper_size", "")),
            excel_orientation_mode=str(data.get("excel_orientation_mode", "auto")),
            pdf_scale_mode=str(data.get("pdf_scale_mode", "fit")),
            pdf_scale_percent=scale_percent,
            pdf_auto_rotate=bool(data.get("pdf_auto_rotate", True)),
            pdf_center=bool(data.get("pdf_center", True)),
            pdf_warn_if_clip=bool(data.get("pdf_warn_if_clip", True)),
            pdf_printer_scale=printer_scale,
            language_mode=language_mode,
            update_check_enabled=bool(data.get("update_check_enabled", True)),
            auto_update_enabled=bool(data.get("auto_update_enabled", False)),
            update_snooze_until=str(data.get("update_snooze_until", "")),
            last_update_check=str(data.get("last_update_check", "")),
            file_list_column_widths=widths,
        )


def _get_app_data_dir() -> Path:
    """Get the appropriate directory for storing app data."""
    import os
    import sys
    # Use AppData/Local for installed app, or local dir for development
    if getattr(sys, 'frozen', False):
        # Running as compiled exe
        app_data = Path(os.environ.get('LOCALAPPDATA', Path.home() / 'AppData' / 'Local'))
        return app_data / 'RakuPrint'
    else:
        # Running in development
        return Path(__file__).resolve().parent


class AppContext(QtCore.QObject):
    settings_changed = QtCore.Signal()
    rules_changed = QtCore.Signal()

    def __init__(self) -> None:
        super().__init__()
        base_dir = _get_app_data_dir()
        self._config_dir = base_dir / "config"
        self._log_dir = base_dir / "logging"
        self._config_dir.mkdir(parents=True, exist_ok=True)
        self._log_dir.mkdir(parents=True, exist_ok=True)

        self.settings_path = self._config_dir / "user_settings.json"
        self.rules_path = self._config_dir / "rules.json"
        self.log_path = self._log_dir / "app.log"

        self.settings = self._load_settings()
        self.rules = self._load_rules()

    def _load_settings(self) -> UserSettings:
        if self.settings_path.exists():
            try:
                with open(str(self.settings_path), "r", encoding="utf-8") as f:
                    data = json.loads(f.read())
                return UserSettings.from_dict(data)
            except Exception:
                return UserSettings()
        return UserSettings()

    def save_settings(self) -> None:
        try:
            # Use standard file writing to avoid potential recursion issues
            with open(str(self.settings_path), "w", encoding="utf-8") as f:
                f.write(json.dumps(self.settings.to_dict(), indent=2))
            self.settings_changed.emit()
        except Exception:
            # Silently fail to avoid recursion loops
            pass

    def _load_rules(self) -> dict:
        if self.rules_path.exists():
            try:
                with open(str(self.rules_path), "r", encoding="utf-8") as f:
                    data = json.loads(f.read())
                if isinstance(data, dict):
                    return data
            except Exception:
                pass
        rules = dict(DEFAULT_RULES)
        try:
            with open(str(self.rules_path), "w", encoding="utf-8") as f:
                f.write(json.dumps(rules, indent=2))
        except Exception:
            pass
        return rules

    def save_rules(self) -> None:
        try:
            with open(str(self.rules_path), "w", encoding="utf-8") as f:
                f.write(json.dumps(self.rules, indent=2))
            self.rules_changed.emit()
        except Exception:
            # Silently fail to avoid recursion loops
            pass

    def update_setting(self, **kwargs) -> None:
        for key, value in kwargs.items():
            if hasattr(self.settings, key):
                setattr(self.settings, key, value)
        self.save_settings()

    def update_rule(self, extension: str, printer: str) -> None:
        self.rules[extension.lower()] = {"printer": printer}
        self.save_rules()

    def remove_rule(self, extension: str) -> None:
        key = extension.lower()
        if key in self.rules:
            self.rules.pop(key, None)
            self.save_rules()
