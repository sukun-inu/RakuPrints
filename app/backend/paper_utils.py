from __future__ import annotations

from typing import Iterable


STANDARD_SIZES_MM: dict[str, tuple[float, float]] = {
    "A3": (297.0, 420.0),
    "A4": (210.0, 297.0),
    "A5": (148.0, 210.0),
    "B4": (257.0, 364.0),
    "B5": (182.0, 257.0),
    "LETTER": (215.9, 279.4),
    "LEGAL": (215.9, 355.6),
}


def normalize_name(name: str) -> str:
    if not name:
        return ""
    normalized = name.replace(" ", "").replace("-", "").upper()
    for key in ("A3", "A4", "A5", "B4", "B5", "LETTER", "LEGAL"):
        if key in normalized:
            return key
    return ""


def size_key_from_mm(width_mm: float, height_mm: float, tolerance_mm: float = 3.0) -> str:
    if width_mm <= 0 or height_mm <= 0:
        return ""
    for key, (w, h) in STANDARD_SIZES_MM.items():
        if abs(width_mm - w) <= tolerance_mm and abs(height_mm - h) <= tolerance_mm:
            return key
        if abs(width_mm - h) <= tolerance_mm and abs(height_mm - w) <= tolerance_mm:
            return key
    return ""


def size_key_from_points(width_pt: float, height_pt: float, tolerance_mm: float = 3.0) -> str:
    mm_per_point = 25.4 / 72.0
    return size_key_from_mm(width_pt * mm_per_point, height_pt * mm_per_point, tolerance_mm=tolerance_mm)


def is_supported_name(target_name: str, names: Iterable[str]) -> bool:
    if not target_name:
        return False
    if target_name in names:
        return True
    target_key = normalize_name(target_name)
    if not target_key:
        return False
    for name in names:
        if normalize_name(name) == target_key:
            return True
    return False
