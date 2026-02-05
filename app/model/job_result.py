from __future__ import annotations

from dataclasses import dataclass


@dataclass
class JobResult:
    job_id: str
    success: bool
    message: str
    error_detail: str = ""
