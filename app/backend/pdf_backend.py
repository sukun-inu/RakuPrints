from __future__ import annotations

from pathlib import Path

import json
import subprocess
import sys

from app.app_context import AppContext
from app.model.print_job import PrintJob


class PdfBackend:
    def __init__(self, context: AppContext) -> None:
        self._context = context

    def print(self, job: PrintJob) -> None:
        if not Path(job.file_path).exists():
            raise RuntimeError("ファイルが見つかりません。")

        try:
            import fitz  # type: ignore
        except Exception as exc:
            raise RuntimeError("PDF 印刷には PyMuPDF が必要です。") from exc
        payload = {
            "file_path": job.file_path,
            "printer_name": job.printer_name,
            "copies": job.copies,
            "duplex": job.duplex.value,
            "paper_size": job.paper_size,
            "dpi": 600,
        }
        try:
            if getattr(sys, "frozen", False):
                cmd = [sys.executable, "--pdf-worker"]
            else:
                cmd = [sys.executable, "-m", "app.backend.pdf_worker"]
            result = subprocess.run(
                cmd,
                input=json.dumps(payload),
                text=True,
                capture_output=True,
                timeout=300,
            )
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError("PDF 印刷がタイムアウトしました。") from exc

        if result.returncode != 0:
            message = result.stderr.strip() or result.stdout.strip() or "PDF の印刷に失敗しました。"
            raise RuntimeError(message)
