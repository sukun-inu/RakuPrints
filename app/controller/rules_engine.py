from __future__ import annotations

from pathlib import Path

from app.app_context import AppContext


class RulesEngine:
    def __init__(self, context: AppContext) -> None:
        self._context = context

    def resolve_printer(self, file_path: str, default_printer: str) -> str:
        ext = Path(file_path).suffix.lower()
        rule = self._context.rules.get(ext)
        if isinstance(rule, dict):
            printer = rule.get("printer", "")
            if printer:
                return printer
        return default_printer
