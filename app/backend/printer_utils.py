from __future__ import annotations

from typing import List

import json
import os
from pathlib import Path
import subprocess
import sys
import time


def _require_win32print():
    try:
        import win32print  # type: ignore
    except Exception as exc:
        raise RuntimeError("Windows でプリンター一覧を取得するには pywin32 が必要です。") from exc
    return win32print


_CACHE_TTL_SECONDS = 60 * 60
_CACHE_FILE_NAME = "printer_paper_cache.json"


def _get_cache_path() -> Path:
    if getattr(sys, "frozen", False):
        base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")) / "RakuPrint"
    else:
        base = Path(__file__).resolve().parents[1]
    config_dir = base / "config"
    config_dir.mkdir(parents=True, exist_ok=True)
    return config_dir / _CACHE_FILE_NAME


def _load_paper_cache() -> dict:
    path = _get_cache_path()
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8") as handle:
            data = json.loads(handle.read())
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_paper_cache(data: dict) -> None:
    path = _get_cache_path()
    try:
        with path.open("w", encoding="utf-8") as handle:
            handle.write(json.dumps(data, indent=2))
    except Exception:
        return


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
    cache = _load_paper_cache()
    entry = cache.get(printer_name) if isinstance(cache, dict) else None
    cached_sizes = None
    cached_ts = None
    if isinstance(entry, dict):
        cached_sizes = entry.get("sizes")
        cached_ts = entry.get("timestamp")
    now = time.time()
    if isinstance(cached_sizes, list) and isinstance(cached_ts, (int, float)):
        if now - cached_ts < _CACHE_TTL_SECONDS:
            return cached_sizes

    try:
        win32print = _require_win32print()
        names = win32print.DeviceCapabilities(printer_name, None, win32print.DC_PAPERNAMES)
        if not names:
            return cached_sizes or []
        cleaned = [name.strip() for name in names if name and str(name).strip()]
        sizes = sorted(set(cleaned))
        cache[printer_name] = {"timestamp": now, "sizes": sizes}
        _save_paper_cache(cache)
        return sizes
    except Exception:
        return cached_sizes or []


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
