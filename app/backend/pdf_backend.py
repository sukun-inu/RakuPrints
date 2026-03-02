from __future__ import annotations

from pathlib import Path

import json
import subprocess
import sys

from app.app_context import AppContext
from app.model.print_job import PrintJob
from app.backend.paper_utils import is_supported_name, normalize_name, size_key_from_points
from app.backend.printer_utils import get_default_printer_name, list_paper_sizes


def _resolve_printer_name(job: PrintJob) -> str:
    if job.printer_name:
        return job.printer_name
    try:
        return get_default_printer_name()
    except Exception:
        return ""


def _assert_printer_paper_supported(printer_name: str, paper_size: str) -> None:
    if not printer_name or not paper_size:
        return
    try:
        sizes = list_paper_sizes(printer_name)
    except Exception:
        return
    if sizes and not is_supported_name(paper_size, sizes):
        raise RuntimeError("プリンターが選択した用紙サイズに対応していません。")


def _assert_pdf_paper_matches(doc, paper_size: str) -> None:
    if not paper_size:
        return
    target_key = normalize_name(paper_size)
    if not target_key:
        return
    for page_index in range(doc.page_count):
        page = doc.load_page(page_index)
        rect = page.rect
        page_key = size_key_from_points(rect.width, rect.height)
        if page_key and page_key != target_key:
            raise RuntimeError(
                f"PDF の用紙サイズが選択したサイズと一致しません。"
                f" (ページ {page_index + 1}: {page_key}, 選択: {target_key})"
            )


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

        try:
            doc = fitz.open(job.file_path)
            page_count = int(getattr(doc, "page_count", 0) or 0)
        except Exception as exc:
            raise RuntimeError("PDF の読み込みに失敗しました。") from exc
        try:
            if page_count <= 0:
                raise RuntimeError("PDF のページが見つかりません。")
            printer_name = _resolve_printer_name(job)
            _assert_printer_paper_supported(printer_name, job.paper_size)
            _assert_pdf_paper_matches(doc, job.paper_size)
        finally:
            try:
                doc.close()
            except Exception:
                pass

        payload = {
            "file_path": job.file_path,
            "printer_name": job.printer_name,
            "copies": job.copies,
            "duplex": job.duplex.value,
            "paper_size": job.paper_size,
            "dpi": 600,
        }
        max_pages = 10
        if page_count <= max_pages:
            self._run_worker(payload)
            return

        for start in range(0, page_count, max_pages):
            end = min(page_count - 1, start + max_pages - 1)
            chunk_payload = dict(payload, page_start=start, page_end=end)
            self._run_worker(chunk_payload)

    @staticmethod
    def _run_worker(payload: dict) -> None:
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
