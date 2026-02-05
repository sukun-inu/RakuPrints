from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
import uuid
from typing import List


class FileType(str, Enum):
    PDF = "PDF"
    WORD = "Word"
    EXCEL = "Excel"
    PPT = "PowerPoint"
    UNKNOWN = "不明"


class DuplexMode(str, Enum):
    OFF = "しない"
    LONG_EDGE = "長辺とじ"
    SHORT_EDGE = "短辺とじ"


class JobStatus(str, Enum):
    WAITING = "待機"
    PRINTING = "印刷中"
    SUCCESS = "完了"
    FAILED = "失敗"
    CANCELLED = "キャンセル"
    SKIPPED = "スキップ"


@dataclass
class PrintJob:
    file_path: str
    file_type: FileType
    printer_name: str
    copies: int
    duplex: DuplexMode
    manual_printer: bool = False
    enabled: bool = True
    status: JobStatus = JobStatus.WAITING
    message: str = ""
    summary: str = ""
    excel_sheets: List[str] = field(default_factory=list)
    excel_auto_orientation: bool = False
    paper_size: str = ""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))

    @property
    def file_name(self) -> str:
        return Path(self.file_path).name

    @property
    def extension(self) -> str:
        return Path(self.file_path).suffix.lower()

    def display_printer(self) -> str:
        return self.printer_name or "自動"

    def auto_label(self) -> str:
        ext = self.extension
        if ext == ".pdf":
            return "PDF資料"
        if ext in (".doc", ".docx"):
            return "Word文書"
        if ext in (".xls", ".xlsx", ".xlsm"):
            return "Excel表"
        if ext in (".ppt", ".pptx"):
            return "PowerPoint資料"
        return "不明"

    def display_sheets(self) -> str:
        if self.file_type != FileType.EXCEL:
            return ""
        if not self.excel_sheets:
            return "全シート"
        joined = ", ".join(self.excel_sheets)
        return joined if len(joined) <= 20 else f"{joined[:17]}..."
