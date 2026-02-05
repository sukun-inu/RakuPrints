from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import json

from PySide6 import QtCore

from app.model.print_job import DuplexMode


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

    def to_dict(self) -> dict:
        return {
            "use_default_printer": self.use_default_printer,
            "selected_printer": self.selected_printer,
            "copies": self.copies,
            "duplex": self.duplex.value,
            "theme_mode": self.theme_mode,
            "paper_size": self.paper_size,
            "excel_orientation_mode": self.excel_orientation_mode,
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
        return cls(
            use_default_printer=bool(data.get("use_default_printer", True)),
            selected_printer=str(data.get("selected_printer", "")),
            copies=int(data.get("copies", 1)),
            duplex=duplex,
            theme_mode=str(data.get("theme_mode", "system")),
            paper_size=str(data.get("paper_size", "")),
            excel_orientation_mode=str(data.get("excel_orientation_mode", "auto")),
        )


class AppContext(QtCore.QObject):
    settings_changed = QtCore.Signal()
    rules_changed = QtCore.Signal()

    def __init__(self) -> None:
        super().__init__()
        base_dir = Path(__file__).resolve().parent
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
                data = json.loads(self.settings_path.read_text(encoding="utf-8"))
                return UserSettings.from_dict(data)
            except Exception:
                return UserSettings()
        return UserSettings()

    def save_settings(self) -> None:
        self.settings_path.write_text(
            json.dumps(self.settings.to_dict(), indent=2),
            encoding="utf-8",
        )
        self.settings_changed.emit()

    def _load_rules(self) -> dict:
        if self.rules_path.exists():
            try:
                data = json.loads(self.rules_path.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    return data
            except Exception:
                pass
        rules = dict(DEFAULT_RULES)
        try:
            self.rules_path.write_text(
                json.dumps(rules, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass
        return rules

    def save_rules(self) -> None:
        self.rules_path.write_text(
            json.dumps(self.rules, indent=2),
            encoding="utf-8",
        )
        self.rules_changed.emit()

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
