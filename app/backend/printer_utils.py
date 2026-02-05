from __future__ import annotations

from typing import List

import subprocess


def _require_win32print():
    try:
        import win32print  # type: ignore
    except Exception as exc:
        raise RuntimeError("Windows でプリンター一覧を取得するには pywin32 が必要です。") from exc
    return win32print


def list_printers() -> List[str]:
    win32print = _require_win32print()
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    printers = win32print.EnumPrinters(flags)
    names = [printer[2] for printer in printers]
    return sorted(names)


def get_default_printer_name() -> str:
    win32print = _require_win32print()
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return ""


def open_printer_properties(printer_name: str) -> None:
    if not printer_name:
        raise RuntimeError("プリンターが見つかりません。")
    cmd = ["rundll32.exe", "printui.dll,PrintUIEntry", "/p", "/n", printer_name]
    creationflags = 0
    if hasattr(subprocess, "CREATE_NO_WINDOW"):
        creationflags = subprocess.CREATE_NO_WINDOW
    subprocess.Popen(cmd, creationflags=creationflags)


def list_paper_sizes(printer_name: str) -> list[str]:
    if not printer_name:
        return []
    win32print = _require_win32print()
    try:
        names = win32print.DeviceCapabilities(printer_name, None, win32print.DC_PAPERNAMES)
        if not names:
            return []
        cleaned = [name.strip() for name in names if name and str(name).strip()]
        return sorted(set(cleaned))
    except Exception:
        return []


def resolve_excel_printer_name(printer_name: str) -> str:
    if not printer_name:
        return ""
    if " on " in printer_name:
        return printer_name
    win32print = _require_win32print()
    handle = None
    try:
        handle = win32print.OpenPrinter(printer_name)
        info = win32print.GetPrinter(handle, 2)
        port = info.get("pPortName", "") if info else ""
        if port:
            port = port.split(",")[0].strip()
            return f"{printer_name} on {port}"
    except Exception:
        return printer_name
    finally:
        if handle:
            win32print.ClosePrinter(handle)
    return printer_name


def set_default_printer(printer_name: str) -> bool:
    if not printer_name:
        return False
    win32print = _require_win32print()
    try:
        win32print.SetDefaultPrinter(printer_name)
        return True
    except Exception:
        if " on " in printer_name:
            base = printer_name.split(" on ")[0].strip()
            if base:
                try:
                    win32print.SetDefaultPrinter(base)
                    return True
                except Exception:
                    return False
        return False
