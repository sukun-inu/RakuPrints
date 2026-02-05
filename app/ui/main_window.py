from __future__ import annotations

from PySide6 import QtCore, QtGui, QtWidgets

try:
    from PySide6.QtWinExtras import QWinTaskbarButton  # type: ignore
except Exception:  # pragma: no cover - optional on non-Windows
    QWinTaskbarButton = None

from app.app_context import AppContext
from app.controller.job_manager import JobManager
from app.controller.job_executor import JobExecutor
from app.model.print_job import DuplexMode, JobStatus, FileType
from app.backend.printer_utils import (
    list_printers,
    get_default_printer_name,
    open_printer_properties,
    list_paper_sizes,
)
from app.ui.file_list_view import FileListView
from app.ui.settings_panel import SettingsPanel
from app.ui.printer_selector import PrinterSelectorDialog
from app.ui.progress_dialog import ProgressDialog
from app.ui.about_dialog import AboutDialog
from app.ui.log_summary_dialog import LogSummaryDialog
from app.ui.theme import apply_theme
from app.ui.excel_sheet_selector import ExcelSheetSelectorDialog
from app.ui.excel_orientation_dialog import ExcelOrientationDialog
from app.controller.excel_orientation_analyzer import ExcelOrientationAnalyzer


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, context: AppContext, job_manager: JobManager) -> None:
        super().__init__()
        self._context = context
        self._job_manager = job_manager
        self._executor: JobExecutor | None = None
        self._progress_dialog: ProgressDialog | None = None
        self._printers: list[str] = []
        self._orientation_analyzer: ExcelOrientationAnalyzer | None = None
        self._orientation_progress: QtWidgets.QProgressDialog | None = None
        self._pending_jobs: list = []
        self._taskbar_button = None
        self._taskbar_progress = None

        self.setWindowTitle("らーく印刷")
        self.resize(1100, 700)

        self._build_menu()
        self._build_layout()
        self._bind_signals()

        self._refresh_settings()
        self._load_printers()
        self._refresh_rules()
        self._update_status()

    def _build_menu(self) -> None:
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("ファイル")
        add_files_action = QtGui.QAction("ファイル追加", self)
        add_folder_action = QtGui.QAction("フォルダー追加", self)
        apply_rules_action = QtGui.QAction("ルールで自動設定（手動は保護）", self)
        apply_rules_force_action = QtGui.QAction("ルールで全部上書き", self)
        exit_action = QtGui.QAction("終了", self)

        add_files_action.triggered.connect(self._on_add_files)
        add_folder_action.triggered.connect(self._on_add_folder)
        apply_rules_action.triggered.connect(lambda: self._job_manager.apply_rules(force=False))
        apply_rules_force_action.triggered.connect(lambda: self._job_manager.apply_rules(force=True))
        exit_action.triggered.connect(self.close)

        file_menu.addAction(add_files_action)
        file_menu.addAction(add_folder_action)
        file_menu.addAction(apply_rules_action)
        file_menu.addAction(apply_rules_force_action)
        file_menu.addSeparator()
        file_menu.addAction(exit_action)

        help_menu = menu_bar.addMenu("ヘルプ")
        about_action = QtGui.QAction("このアプリについて", self)
        log_summary_action = QtGui.QAction("ログの要約", self)
        about_action.triggered.connect(self._on_about)
        help_menu.addAction(about_action)
        log_summary_action.triggered.connect(self._on_log_summary)
        help_menu.addAction(log_summary_action)

    def _build_layout(self) -> None:
        splitter = QtWidgets.QSplitter()
        splitter.setOrientation(QtCore.Qt.Horizontal)

        self.file_list = FileListView(self._job_manager)
        self.settings_panel = SettingsPanel()

        splitter.addWidget(self.file_list)
        splitter.addWidget(self.settings_panel)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)

        central = QtWidgets.QWidget()
        central_layout = QtWidgets.QVBoxLayout(central)
        central_layout.addWidget(splitter)

        bottom_layout = QtWidgets.QHBoxLayout()
        self.start_button = QtWidgets.QPushButton("印刷開始")
        self.retry_button = QtWidgets.QPushButton("失敗だけ再印刷")
        self.retry_button.setEnabled(False)
        bottom_layout.addWidget(self.start_button)
        bottom_layout.addWidget(self.retry_button)
        bottom_layout.addStretch(1)
        central_layout.addLayout(bottom_layout)

        self.setCentralWidget(central)
        self.statusBar().showMessage("準備完了")

    def _bind_signals(self) -> None:
        self.file_list.files_dropped.connect(self._job_manager.add_files)
        self.file_list.printer_requested.connect(self._on_job_printer_select)
        self.file_list.excel_sheets_requested.connect(self._on_excel_sheets_select)
        self.file_list.print_selected_requested.connect(self._on_print_selected)
        self.file_list.printer_selected_requested.connect(self._on_printer_selected)

        self.settings_panel.use_default_changed.connect(self._on_use_default_changed)
        self.settings_panel.select_printer_clicked.connect(self._on_global_printer_select)
        self.settings_panel.printer_settings_clicked.connect(self._on_open_printer_settings)
        self.settings_panel.copies_changed.connect(self._on_copies_changed)
        self.settings_panel.duplex_changed.connect(self._on_duplex_changed)
        self.settings_panel.excel_orientation_mode_changed.connect(self._on_excel_orientation_mode_changed)
        self.settings_panel.rule_printer_changed.connect(self._on_rule_printer_changed)
        self.settings_panel.rule_add_requested.connect(self._on_rule_add)
        self.settings_panel.rule_remove_requested.connect(self._on_rule_remove)
        self.settings_panel.theme_changed.connect(self._on_theme_changed)
        self.settings_panel.paper_size_changed.connect(self._on_paper_size_changed)

        self.start_button.clicked.connect(self._on_start_printing)
        self.retry_button.clicked.connect(self._on_retry_failed)

        self._job_manager.jobs_changed.connect(self._update_status)
        self._job_manager.job_updated.connect(self._update_status)

        self._context.rules_changed.connect(self._refresh_rules)
        self._context.settings_changed.connect(self._refresh_settings)

    def _refresh_settings(self) -> None:
        settings = self._context.settings
        default_printer = self._get_default_printer_name()
        self.settings_panel.set_settings(
            use_default=settings.use_default_printer,
            copies=settings.copies,
            duplex=settings.duplex.value,
            theme_mode=settings.theme_mode,
            selected_printer=settings.selected_printer,
            default_printer=default_printer,
            excel_orientation_mode=settings.excel_orientation_mode,
        )
        self._refresh_paper_sizes()
        self._refresh_rules()

    def _refresh_rules(self) -> None:
        self.settings_panel.set_printers(self._printers)
        auto_label = self._get_effective_printer_name()
        self.settings_panel.set_rules(self._context.rules, auto_label)

    def _update_status(self, *_args) -> None:
        total = self._job_manager.job_count()
        failed = sum(1 for job in self._job_manager.jobs() if job.status == JobStatus.FAILED)
        completed = sum(1 for job in self._job_manager.jobs() if job.status == JobStatus.SUCCESS)
        self.statusBar().showMessage(f"ジョブ数: {total} | 完了: {completed} | 失敗: {failed}")
        self.retry_button.setEnabled(failed > 0 and not (self._executor and self._executor.isRunning()))

    def _on_add_files(self) -> None:
        filter_text = "印刷できるファイル (*.pdf *.doc *.docx *.xls *.xlsx *.xlsm *.ppt *.pptx)"
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "ファイルを追加", "", filter_text)
        if files:
            self._job_manager.add_files(files)

    def _on_add_folder(self) -> None:
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "フォルダーを追加")
        if folder:
            self._job_manager.add_folder(folder, recursive=True)

    def _on_about(self) -> None:
        dialog = AboutDialog(self)
        dialog.exec()

    def _on_job_printer_select(self, job_id: str) -> None:
        job = self._job_manager.find_job_by_id(job_id)
        if not job:
            return
        selected, ok = PrinterSelectorDialog.get_printer(self, job.printer_name)
        if ok and selected:
            self._job_manager.set_job_printer(job_id, selected)

    def _on_excel_sheets_select(self, job_id: str) -> None:
        job = self._job_manager.find_job_by_id(job_id)
        if not job:
            return
        try:
            sheet_names = self._job_manager.list_excel_sheets(job.file_path)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Excel", str(exc))
            return
        selected, ok = ExcelSheetSelectorDialog.get_sheets(
            job.file_name,
            sheet_names,
            selected=job.excel_sheets,
            parent=self,
        )
        if ok:
            self._job_manager.set_job_sheets(job_id, selected)

    def _on_global_printer_select(self) -> None:
        current = self._context.settings.selected_printer
        selected, ok = PrinterSelectorDialog.get_printer(self, current)
        if ok:
            self._context.update_setting(selected_printer=selected)
            if not self._context.settings.use_default_printer:
                self._job_manager.apply_rules()
            self._refresh_paper_sizes()

    def _on_use_default_changed(self, use_default: bool) -> None:
        self._context.update_setting(use_default_printer=use_default)
        if not use_default and not self._context.settings.selected_printer:
            self._on_global_printer_select()
        self._job_manager.apply_rules()
        self._refresh_paper_sizes()

    def _on_copies_changed(self, value: int) -> None:
        self._context.update_setting(copies=value)
        self._job_manager.apply_settings_to_jobs()

    def _on_duplex_changed(self, value: str) -> None:
        try:
            duplex = DuplexMode(value)
        except ValueError:
            duplex = DuplexMode.OFF
        self._context.update_setting(duplex=duplex)
        self._job_manager.apply_settings_to_jobs()

    def _on_paper_size_changed(self, value: str) -> None:
        self._context.update_setting(paper_size=value)
        self._job_manager.apply_settings_to_jobs()

    def _on_excel_orientation_mode_changed(self, mode: str) -> None:
        self._context.update_setting(excel_orientation_mode=mode)

    def _on_theme_changed(self, mode: str) -> None:
        self._context.update_setting(theme_mode=mode)
        app = QtWidgets.QApplication.instance()
        if app:
            apply_theme(app, mode)

    def _on_open_printer_settings(self) -> None:
        settings = self._context.settings
        printer_name = ""
        if settings.use_default_printer:
            printer_name = self._get_default_printer_name()
        else:
            printer_name = settings.selected_printer
        if not printer_name:
            QtWidgets.QMessageBox.information(self, "プリンター", "プリンターが選ばれていません。")
            return
        try:
            open_printer_properties(printer_name)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "プリンター", str(exc))

    def _on_rule_printer_changed(self, extension: str, printer: str) -> None:
        self._context.update_rule(extension, printer)
        self._job_manager.apply_rules()

    def _on_rule_add(self, extension: str) -> None:
        if extension in self._context.rules:
            QtWidgets.QMessageBox.information(self, "ルール", "すでに追加されています。")
            return
        self._context.update_rule(extension, "")
        self._job_manager.apply_rules()

    def _on_rule_remove(self, extensions: list[str]) -> None:
        if not extensions:
            return
        for ext in extensions:
            self._context.remove_rule(ext)
        self._job_manager.apply_rules(force=True)

    def _on_start_printing(self) -> None:
        if self._executor and self._executor.isRunning():
            return
        if self._job_manager.job_count() == 0:
            QtWidgets.QMessageBox.information(self, "印刷", "印刷するファイルがありません。")
            return
        enabled_jobs = self._job_manager.get_enabled_jobs()
        if not enabled_jobs:
            QtWidgets.QMessageBox.information(self, "印刷", "印刷するチェックが入ったファイルがありません。")
            return

        if self._context.settings.excel_orientation_mode == "ask":
            excel_jobs = [job for job in enabled_jobs if job.file_type == FileType.EXCEL]
            if excel_jobs:
                self._pending_jobs = enabled_jobs
                self._start_excel_orientation_analysis(excel_jobs)
                return

        self._lock_ui(True)
        self._job_manager.reset_statuses()
        self._start_executor(enabled_jobs)

    def _on_print_selected(self, job_ids: list[str]) -> None:
        if self._executor and self._executor.isRunning():
            return
        jobs = [job for job in self._job_manager.jobs() if job.id in set(job_ids)]
        if not jobs:
            QtWidgets.QMessageBox.information(self, "印刷", "選択された行がありません。")
            return
        if self._context.settings.excel_orientation_mode == "ask":
            excel_jobs = [job for job in jobs if job.file_type == FileType.EXCEL]
            if excel_jobs:
                self._pending_jobs = jobs
                self._start_excel_orientation_analysis(excel_jobs)
                return
        self._lock_ui(True)
        self._job_manager.reset_statuses_for(job_ids)
        self._start_executor(jobs)

    def _on_printer_selected(self, job_ids: list[str]) -> None:
        if not job_ids:
            return
        selected, ok = PrinterSelectorDialog.get_printer(self, "")
        if ok and selected:
            self._job_manager.set_jobs_printer(job_ids, selected)

    def _on_retry_failed(self) -> None:
        if self._executor and self._executor.isRunning():
            return
        failed_jobs = self._job_manager.get_failed_jobs()
        if not failed_jobs:
            QtWidgets.QMessageBox.information(self, "再印刷", "失敗した印刷はありません。")
            return
        self._lock_ui(True)
        self._job_manager.reset_failed_jobs()
        self._start_executor(failed_jobs)

    def _start_executor(self, jobs) -> None:
        self._executor = JobExecutor(self._context, self._job_manager, jobs)
        self._executor.job_status.connect(self._job_manager.set_job_status)
        self._executor.progress.connect(self._on_progress)
        self._executor.job_failed.connect(self._on_job_failed)
        self._executor.finished_all.connect(self._on_finished)

        self._progress_dialog = ProgressDialog(self)
        self._progress_dialog.cancel_requested.connect(self._executor.request_cancel)
        self._progress_dialog.set_total(len(jobs))
        self._progress_dialog.show()
        self._set_taskbar_total(len(jobs))

        self._executor.start()

    def _on_progress(self, completed: int, total: int, current: str) -> None:
        if self._progress_dialog:
            self._progress_dialog.update_progress(completed, total, current)
        self._update_taskbar_progress(completed, total)

    def _on_job_failed(self, job_id: str) -> None:
        job = self._job_manager.find_job_by_id(job_id)
        file_name = job.file_name if job else "不明なファイル"
        summary = job.summary if job and job.summary else "印刷に失敗しました。"
        message = job.message if job and job.message else ""
        dialog = QtWidgets.QMessageBox(self)
        dialog.setIcon(QtWidgets.QMessageBox.Warning)
        dialog.setWindowTitle("印刷に失敗しました")
        if message:
            dialog.setText(f"{file_name}\n{summary}\n\n詳細:\n{message}")
        else:
            dialog.setText(f"{file_name}\n{summary}")
        log_button = dialog.addButton("詳細ログを開く", QtWidgets.QMessageBox.ActionRole)
        dialog.addButton(QtWidgets.QMessageBox.Ok)
        dialog.exec()
        if dialog.clickedButton() == log_button:
            QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(str(self._context.log_path)))

    def _on_finished(self, cancelled: bool) -> None:
        if self._progress_dialog:
            self._progress_dialog.set_finished(cancelled)
            QtCore.QTimer.singleShot(800, self._progress_dialog.accept)
            self._progress_dialog = None
        self._clear_taskbar_progress()
        self._lock_ui(False)
        self._executor = None
        self._update_status()

    def _lock_ui(self, locked: bool) -> None:
        self.file_list.setEnabled(not locked)
        self.settings_panel.setEnabled(not locked)
        self.start_button.setEnabled(not locked)
        self.retry_button.setEnabled(not locked)
        self.menuBar().setEnabled(not locked)

    def _get_default_printer_name(self) -> str:
        try:
            return get_default_printer_name()
        except Exception:
            return ""

    def _get_effective_printer_name(self) -> str:
        settings = self._context.settings
        if settings.use_default_printer:
            return self._get_default_printer_name() or "未検出"
        return settings.selected_printer or "未選択"

    def _load_printers(self) -> None:
        try:
            self._printers = list_printers()
        except Exception:
            self._printers = []

    def _refresh_paper_sizes(self) -> None:
        settings = self._context.settings
        printer_name = (
            self._get_default_printer_name()
            if settings.use_default_printer
            else settings.selected_printer
        )
        sizes = list_paper_sizes(printer_name) if printer_name else []
        if sizes:
            enabled = True
            tooltip = ""
        else:
            enabled = False
            tooltip = "このプリンターから用紙サイズを取得できません。"
        self.settings_panel.set_paper_sizes(sizes, settings.paper_size, enabled, tooltip)

    def _on_log_summary(self) -> None:
        items = []
        for job in self._job_manager.get_failed_jobs():
            summary = job.summary or "印刷に失敗しました。"
            detail = job.message or "詳細情報がありません。"
            items.append((job.file_name, summary, detail))
        dialog = LogSummaryDialog(str(self._context.log_path), items, self)
        dialog.exec()

    def _start_excel_orientation_analysis(self, jobs) -> None:
        if self._orientation_analyzer and self._orientation_analyzer.isRunning():
            return
        self._orientation_analyzer = ExcelOrientationAnalyzer(jobs)
        self._orientation_analyzer.completed.connect(self._on_excel_orientation_ready)
        self._orientation_analyzer.failed.connect(self._on_excel_orientation_failed)
        self._orientation_progress = QtWidgets.QProgressDialog(
            "Excel の向きを確認中です...", "", 0, 0, self
        )
        self._orientation_progress.setWindowTitle("確認中")
        self._orientation_progress.setCancelButton(None)
        self._orientation_progress.setWindowModality(QtCore.Qt.ApplicationModal)
        self._orientation_progress.show()
        self._orientation_analyzer.start()

    def _on_excel_orientation_ready(self, results) -> None:
        if self._orientation_progress:
            self._orientation_progress.close()
            self._orientation_progress = None
        items = [
            (result.job_id, result.file_name, result.recommendation, result.reason)
            for result in results
            if result.recommendation in ("横向き", "縦向き")
        ]
        if items:
            selected_ids, ok = ExcelOrientationDialog.get_selection(items, self)
            if not ok:
                self._pending_jobs = []
                return
            pending_ids = {job.id for job in self._pending_jobs}
            for job in self._job_manager.jobs():
                if job.id in pending_ids and job.file_type == FileType.EXCEL:
                    job.excel_auto_orientation = job.id in selected_ids
        self._lock_ui(True)
        pending_ids = [job.id for job in self._pending_jobs]
        self._job_manager.reset_statuses_for(pending_ids)
        jobs = self._pending_jobs
        self._pending_jobs = []
        self._start_executor(jobs)

    def _on_excel_orientation_failed(self, message: str) -> None:
        if self._orientation_progress:
            self._orientation_progress.close()
            self._orientation_progress = None
        QtWidgets.QMessageBox.warning(self, "Excel", message)
        self._lock_ui(True)
        pending_ids = [job.id for job in self._pending_jobs]
        self._job_manager.reset_statuses_for(pending_ids)
        jobs = self._pending_jobs
        self._pending_jobs = []
        self._start_executor(jobs)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        if self._executor and self._executor.isRunning():
            QtWidgets.QMessageBox.information(
                self,
                "印刷中",
                "印刷中です。完了を待つか、キャンセルしてください。",
            )
            event.ignore()
            return
        super().closeEvent(event)

    def _set_taskbar_total(self, total: int) -> None:
        if not QWinTaskbarButton:
            return
        if total <= 0:
            self._clear_taskbar_progress()
            return
        self._ensure_taskbar_progress()
        if self._taskbar_progress:
            self._taskbar_progress.setMaximum(total)
            self._taskbar_progress.setValue(0)
            self._taskbar_progress.setVisible(True)

    def _update_taskbar_progress(self, value: int, total: int) -> None:
        if not QWinTaskbarButton:
            return
        if not self._taskbar_progress:
            self._set_taskbar_total(total)
        if self._taskbar_progress:
            if total > 0:
                self._taskbar_progress.setMaximum(total)
            self._taskbar_progress.setValue(value)

    def _clear_taskbar_progress(self) -> None:
        if not QWinTaskbarButton:
            return
        if self._taskbar_progress:
            self._taskbar_progress.setVisible(False)
            self._taskbar_progress.setValue(0)

    def _ensure_taskbar_progress(self) -> None:
        if self._taskbar_button or not QWinTaskbarButton:
            return
        self._taskbar_button = QWinTaskbarButton(self)
        if self.windowHandle() is None:
            self.createWinId()
        window = self.windowHandle()
        if window is not None:
            self._taskbar_button.setWindow(window)
            self._taskbar_progress = self._taskbar_button.progress()
            self._taskbar_progress.setMinimum(0)
