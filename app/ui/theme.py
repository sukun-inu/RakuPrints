from __future__ import annotations

import sys

from PySide6 import QtGui, QtWidgets


def detect_system_theme() -> str:
    if sys.platform != "win32":
        return "light"
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        ) as key:
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return "light" if value == 1 else "dark"
    except Exception:
        return "light"


def resolve_theme(mode: str) -> str:
    if mode == "system":
        return detect_system_theme()
    if mode in ("light", "dark"):
        return mode
    return "light"


def _get_windows_accent_color() -> QtGui.QColor | None:
    if sys.platform != "win32":
        return None
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\DWM",
        ) as key:
            try:
                value, _ = winreg.QueryValueEx(key, "AccentColor")
            except FileNotFoundError:
                value, _ = winreg.QueryValueEx(key, "ColorizationColor")
        value = int(value) & 0xFFFFFFFF
        r = value & 0xFF
        g = (value >> 8) & 0xFF
        b = (value >> 16) & 0xFF
        return QtGui.QColor(r, g, b)
    except Exception:
        return None


def _accent_hex() -> str:
    accent = _get_windows_accent_color()
    if accent is None:
        return "#0078D4"
    return accent.name()


def apply_theme(app: QtWidgets.QApplication, mode: str) -> None:
    theme = resolve_theme(mode)
    app.setStyle("Fusion")
    app.setFont(QtGui.QFont("Yu Gothic UI", 11))
    accent = _accent_hex()

    if theme == "dark":
        colors = {
            "window": "#201F1E",
            "text": "#F3F2F1",
            "base": "#292827",
            "alternate": "#323130",
            "button": "#292827",
            "highlight": accent,
            "highlight_text": "#FFFFFF",
            "grid": "#3B3A39",
            "header": "#2D2C2B",
            "menu_hover": "#2F2E2D",
            "link": accent,
            "disabled_text": "#8A8886",
            "disabled_bg": "#292827",
            "card": "#252423",
        }
    else:
        colors = {
            "window": "#F3F2F1",
            "text": "#323130",
            "base": "#FFFFFF",
            "alternate": "#FAF9F8",
            "button": "#FFFFFF",
            "highlight": accent,
            "highlight_text": "#FFFFFF",
            "grid": "#E1DFDD",
            "header": "#F8F8F8",
            "menu_hover": "#EDEBE9",
            "link": accent,
            "disabled_text": "#A19F9D",
            "disabled_bg": "#F3F2F1",
            "card": "#FFFFFF",
        }

    palette = QtGui.QPalette()
    palette.setColor(QtGui.QPalette.Window, QtGui.QColor(colors["window"]))
    palette.setColor(QtGui.QPalette.WindowText, QtGui.QColor(colors["text"]))
    palette.setColor(QtGui.QPalette.Base, QtGui.QColor(colors["base"]))
    palette.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor(colors["alternate"]))
    palette.setColor(QtGui.QPalette.Text, QtGui.QColor(colors["text"]))
    palette.setColor(QtGui.QPalette.Button, QtGui.QColor(colors["button"]))
    palette.setColor(QtGui.QPalette.ButtonText, QtGui.QColor(colors["text"]))
    palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor(colors["highlight"]))
    palette.setColor(QtGui.QPalette.HighlightedText, QtGui.QColor(colors["highlight_text"]))
    palette.setColor(QtGui.QPalette.ToolTipBase, QtGui.QColor("#FFFFE1"))
    palette.setColor(QtGui.QPalette.ToolTipText, QtGui.QColor(colors["text"]))
    palette.setColor(QtGui.QPalette.Link, QtGui.QColor(colors["link"]))
    palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Text, QtGui.QColor(colors["disabled_text"]))
    palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, QtGui.QColor(colors["disabled_text"]))
    palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, QtGui.QColor(colors["disabled_text"]))
    palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Base, QtGui.QColor(colors["disabled_bg"]))
    palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Button, QtGui.QColor(colors["disabled_bg"]))
    app.setPalette(palette)

    app.setStyleSheet(
        f"""
        QMainWindow {{ background: {colors['window']}; }}
        QGroupBox {{
            font-weight: bold;
            margin-top: 12px;
            border: 1px solid {colors['grid']};
            border-radius: 8px;
            padding: 8px;
            background: {colors['card']};
        }}
        QGroupBox::title {{ subcontrol-origin: margin; left: 10px; padding: 0 6px; }}
        QTableView {{
            gridline-color: {colors['grid']};
            selection-background-color: {colors['highlight']};
            selection-color: {colors['highlight_text']};
            background: {colors['base']};
            border: 1px solid {colors['grid']};
        }}
        QHeaderView::section {{
            background: {colors['header']};
            padding: 6px;
            border: 1px solid {colors['grid']};
            font-weight: 600;
        }}
        QLineEdit, QComboBox, QSpinBox {{
            padding: 4px 8px;
            border: 1px solid {colors['grid']};
            border-radius: 6px;
            background: {colors['base']};
        }}
        QComboBox::drop-down {{
            subcontrol-origin: padding;
            width: 24px;
            border-left: 1px solid {colors['grid']};
        }}
        QPushButton {{
            padding: 6px 12px;
            border: 1px solid {colors['grid']};
            border-radius: 6px;
            background: {colors['button']};
        }}
        QPushButton:hover {{ border-color: {colors['highlight']}; }}
        *:disabled {{ color: {colors['disabled_text']}; }}
        QPushButton:disabled {{ background: {colors['disabled_bg']}; }}
        QMenuBar {{ background: {colors['window']}; }}
        QMenuBar::item:selected {{ background: {colors['menu_hover']}; }}
        QStatusBar {{ color: {colors['text']}; background: {colors['window']}; }}
        """
    )
