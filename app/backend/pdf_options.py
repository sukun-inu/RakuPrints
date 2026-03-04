from __future__ import annotations

from dataclasses import dataclass

from app.app_context import UserSettings
from app.model.print_job import PrintJob


@dataclass(frozen=True)
class PdfPrintOptions:
    scale_mode: str
    scale_percent: int
    auto_rotate: bool
    center: bool


def resolve_pdf_options(job: PrintJob, settings: UserSettings) -> PdfPrintOptions:
    scale_mode = job.pdf_scale_mode or settings.pdf_scale_mode or "auto"
    auto_rotate = settings.pdf_auto_rotate if job.pdf_auto_rotate is None else bool(job.pdf_auto_rotate)
    center = settings.pdf_center if job.pdf_center is None else bool(job.pdf_center)

    if job.pdf_scale_percent is not None:
        scale_percent = int(job.pdf_scale_percent)
    else:
        printer_name = job.printer_name or ""
        scale_percent = int(settings.pdf_printer_scale.get(printer_name, settings.pdf_scale_percent))

    if scale_percent < 10:
        scale_percent = 10
    if scale_percent > 200:
        scale_percent = 200

    if scale_mode not in ("auto", "fit", "shrink", "none"):
        scale_mode = "auto"

    return PdfPrintOptions(
        scale_mode=scale_mode,
        scale_percent=scale_percent,
        auto_rotate=auto_rotate,
        center=center,
    )
