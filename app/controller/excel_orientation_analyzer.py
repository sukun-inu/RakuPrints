from __future__ import annotations

import gc
from dataclasses import dataclass

from PySide6 import QtCore

from app.model.print_job import PrintJob


@dataclass
class ExcelOrientationResult:
    job_id: str
    file_name: str
    recommendation: str
    reason: str


class ExcelOrientationAnalyzer(QtCore.QThread):
    completed = QtCore.Signal(list)
    failed = QtCore.Signal(str)

    def __init__(self, jobs: list[PrintJob]) -> None:
        super().__init__()
        self._jobs = list(jobs)

    def run(self) -> None:
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except Exception as exc:
            self.failed.emit("Excel の解析には pywin32 が必要です。")
            return

        results: list[ExcelOrientationResult] = []
        pythoncom.CoInitialize()
        try:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            if hasattr(app, "DisplayAlerts"):
                app.DisplayAlerts = False

            for job in self._jobs:
                workbook = None
                try:
                    workbook = app.Workbooks.Open(job.file_path, ReadOnly=True)
                    sheets = job.excel_sheets or [sheet.Name for sheet in workbook.Worksheets]
                    landscape = 0
                    portrait = 0
                    for name in sheets:
                        sheet = workbook.Worksheets(name)
                        used = sheet.UsedRange
                        width = float(used.Width)
                        height = float(used.Height)
                    if width > height * 1.05:
                        landscape += 1
                    else:
                        portrait += 1
                    recommendation = "横向き" if landscape > portrait else "縦向き"
                    reason = "横幅 > 縦幅" if recommendation == "横向き" else "縦幅 >= 横幅"
                    results.append(
                        ExcelOrientationResult(
                            job_id=job.id,
                            file_name=job.file_name,
                            recommendation=recommendation,
                            reason=reason,
                        )
                    )
                except Exception:
                    results.append(
                        ExcelOrientationResult(
                            job_id=job.id,
                            file_name=job.file_name,
                            recommendation="判定できません",
                            reason="解析できませんでした",
                        )
                    )
                finally:
                    if workbook is not None:
                        workbook.Close(False)
            app.Quit()
            del app
        except Exception:
            self.failed.emit("Excel の向きを解析できませんでした。")
            return
        finally:
            gc.collect()
            pythoncom.CoUninitialize()

        self.completed.emit(results)
