from __future__ import annotations

from PySide6 import QtCore, QtGui, QtWidgets, QtPrintSupport

try:
    from PySide6.QtWinExtras import QWinTaskbarButton  # type: ignore
except Exception:  # pragma: no cover - optional on non-Windows
    QWinTaskbarButton = None

from app.app_context import AppContext
from app.controller.job_manager import JobManager
from app.controller.job_executor import JobExecutor
from app.controller.update_manager import UpdateManager
from app.model.print_job import DuplexMode, JobStatus, FileType
from app.backend.printer_utils import (
    list_printers,
    get_default_printer_name,
    open_printer_properties,
    list_paper_sizes,
)
from app.backend.qt_print_utils import apply_qt_paper_size
from app.backend.pdf_options import resolve_pdf_options
from app.backend.paper_utils import is_supported_name
from app.ui.file_list_view import FileListView
from app.ui.settings_panel import SettingsPanel
from app.ui.printer_selector import PrinterSelectorDialog
from app.ui.progress_dialog import ProgressDialog
from app.ui.about_dialog import AboutDialog
from app.ui.log_summary_dialog import LogSummaryDialog
from app.ui.theme import apply_theme
from app.ui.excel_sheet_selector import ExcelSheetSelectorDialog
from app.ui.excel_orientation_dialog import ExcelOrientationDialog
from app.ui.pdf_options_dialog import PdfOptionsDialog
from app.ui.usage_dialog import UsageDialog
from app.ui.document_preview_panel import DocumentPreviewPanel
from app.ui.preview_window import PreviewWindow
from app.controller.excel_orientation_analyzer import ExcelOrientationAnalyzer
from app.i18n import t, set_language, resolve_language


class MainWindow(QtWidgets.QMainWindow):
    _SPLITTER_SIZES_DEFAULT = [980, 360]

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
        self._update_manager = UpdateManager(context, self)
        self._preview_window: PreviewWindow | None = None
        self._ignore_preview_window_close = False
        self._closing_main_window = False

        # Prefer a 16:9 friendly starting size and require at least 1366x768
        self.resize(1366, 768)
        self.setMinimumSize(1366, 768)

        self._build_menu()
        self._build_layout()
        self._bind_signals()

        # Apply language first to ensure UI elements are properly initialized
        self._apply_language()

        # Always start with preview window hidden.
        self._context.settings.preview_visible = False

        self._refresh_settings()
        self._load_printers()
        self._refresh_rules()
        self._update_status()
        self._update_selection_actions()
        self._refresh_preview_selection()

        QtCore.QTimer.singleShot(600, self._update_manager.check_on_startup)

    def _build_menu(self) -> None:
        self.menu_bar = self.menuBar()
        self.file_menu = self.menu_bar.addMenu("")
        self.help_menu = self.menu_bar.addMenu("")

        self.add_files_action = QtGui.QAction(self)
        self.add_folder_action = QtGui.QAction(self)
        self.apply_rules_action = QtGui.QAction(self)
        self.apply_rules_force_action = QtGui.QAction(self)
        self.exit_action = QtGui.QAction(self)
        self.start_print_action = QtGui.QAction(self)
        self.retry_failed_action = QtGui.QAction(self)
        self.print_selected_action = QtGui.QAction(self)
        self.printer_selected_action = QtGui.QAction(self)
        self.pdf_options_action = QtGui.QAction(self)
        self.excel_sheets_action = QtGui.QAction(self)
        self.move_up_action = QtGui.QAction(self)
        self.move_down_action = QtGui.QAction(self)
        self.delete_action = QtGui.QAction(self)
        self.preview_toggle_action = QtGui.QAction(self)
        self.preview_toggle_action.setCheckable(True)

        self.usage_action = QtGui.QAction(self)
        self.about_action = QtGui.QAction(self)
        self.log_summary_action = QtGui.QAction(self)
        self.check_updates_action = QtGui.QAction(self)

        self.add_files_action.setShortcut(QtGui.QKeySequence.Open)
        self.add_folder_action.setShortcut(QtGui.QKeySequence("Ctrl+Shift+O"))
        self.start_print_action.setShortcut(QtGui.QKeySequence("Ctrl+Return"))
        self.retry_failed_action.setShortcut(QtGui.QKeySequence("Ctrl+R"))
        self.print_selected_action.setShortcut(QtGui.QKeySequence.Print)
        self.printer_selected_action.setShortcut(QtGui.QKeySequence("Ctrl+Shift+P"))
        self.pdf_options_action.setShortcut(QtGui.QKeySequence("Ctrl+Alt+P"))
        self.excel_sheets_action.setShortcut(QtGui.QKeySequence("Ctrl+Alt+E"))
        self.preview_toggle_action.setShortcut(QtGui.QKeySequence("Ctrl+Shift+V"))
        self.usage_action.setShortcut(QtGui.QKeySequence.HelpContents)

        self.add_files_action.triggered.connect(self._on_add_files)
        self.add_folder_action.triggered.connect(self._on_add_folder)
        self.apply_rules_action.triggered.connect(lambda: self._job_manager.apply_rules(force=False))
        self.apply_rules_force_action.triggered.connect(lambda: self._job_manager.apply_rules(force=True))
        self.exit_action.triggered.connect(self.close)
        self.start_print_action.triggered.connect(self._on_start_printing)
        self.retry_failed_action.triggered.connect(self._on_retry_failed)
        self.print_selected_action.triggered.connect(self._on_print_selected_action)
        self.printer_selected_action.triggered.connect(self._on_printer_selected_action)
        self.pdf_options_action.triggered.connect(self._on_pdf_options_action)
        self.excel_sheets_action.triggered.connect(self._on_excel_sheets_action)
        self.move_up_action.triggered.connect(lambda: self.file_list.move_selected(-1))
        self.move_down_action.triggered.connect(lambda: self.file_list.move_selected(1))
        self.delete_action.triggered.connect(lambda: self.file_list.confirm_and_remove_selected())
        self.preview_toggle_action.triggered.connect(self._on_preview_toggle_changed)

        self.usage_action.triggered.connect(self._on_usage)
        self.about_action.triggered.connect(self._on_about)
        self.log_summary_action.triggered.connect(self._on_log_summary)
        self.check_updates_action.triggered.connect(self._on_check_updates)

        self.file_menu.addAction(self.add_files_action)
        self.file_menu.addAction(self.add_folder_action)
        self.file_menu.addAction(self.apply_rules_action)
        self.file_menu.addAction(self.apply_rules_force_action)
        self.file_menu.addAction(self.preview_toggle_action)
        self.file_menu.addSeparator()
        self.file_menu.addAction(self.exit_action)

        self.help_menu.addAction(self.usage_action)
        self.help_menu.addSeparator()
        self.help_menu.addAction(self.about_action)
        self.help_menu.addAction(self.log_summary_action)
        self.help_menu.addAction(self.check_updates_action)

    def _build_layout(self) -> None:
        self.main_splitter = QtWidgets.QSplitter()
        self.main_splitter.setOrientation(QtCore.Qt.Horizontal)

        self.file_list = FileListView(self._job_manager)
        self.file_list.apply_column_widths(self._context.settings.file_list_column_widths)
        self.preview_panel = DocumentPreviewPanel(
            preview_zoom_mode=self._context.settings.preview_zoom_mode,
            preview_zoom_percent=self._context.settings.preview_zoom_percent,
        )
        self.settings_panel = SettingsPanel()

        self.main_splitter.addWidget(self.file_list)
        self.main_splitter.addWidget(self.settings_panel)
        self.main_splitter.setStretchFactor(0, 7)
        self.main_splitter.setStretchFactor(1, 3)
        self.main_splitter.setSizes(list(self._SPLITTER_SIZES_DEFAULT))
        self.main_splitter.setHandleWidth(1)

        central = QtWidgets.QWidget()
        central_layout = QtWidgets.QVBoxLayout(central)
        self.command_bar = QtWidgets.QToolBar()
        self.command_bar.setMovable(False)
        self.command_bar.setFloatable(False)
        self.command_bar.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
        self.command_bar.addAction(self.add_files_action)
        self.command_bar.addAction(self.add_folder_action)
        self.command_bar.addAction(self.preview_toggle_action)
        self.command_bar.addSeparator()
        self.command_bar.addAction(self.start_print_action)
        self.command_bar.addAction(self.print_selected_action)
        self.command_bar.addAction(self.retry_failed_action)
        self.command_bar.addSeparator()
        self.command_bar.addAction(self.move_up_action)
        self.command_bar.addAction(self.move_down_action)
        self.command_bar.addAction(self.delete_action)
        self.command_bar.addSeparator()
        self.command_bar.addAction(self.printer_selected_action)
        self.command_bar.addAction(self.pdf_options_action)
        self.command_bar.addAction(self.excel_sheets_action)
        central_layout.addWidget(self.command_bar)
        central_layout.addWidget(self.main_splitter)

        self.setCentralWidget(central)
        self._apply_preview_state(
            visible=self._context.settings.preview_visible,
            persist=False,
        )

    def _bind_signals(self) -> None:
        self.file_list.files_dropped.connect(self._job_manager.add_files)
        self.file_list.printer_requested.connect(self._on_job_printer_select)
        self.file_list.excel_sheets_requested.connect(self._on_excel_sheets_select)
        self.file_list.print_selected_requested.connect(self._on_print_selected)
        self.file_list.printer_selected_requested.connect(self._on_printer_selected)
        self.file_list.pdf_options_requested.connect(self._on_pdf_options_requested)
        self.file_list.column_widths_changed.connect(self._on_column_widths_changed)
        self.file_list.selectionModel().selectionChanged.connect(self._on_file_selection_changed)
        self.preview_panel.zoom_mode_changed.connect(self._on_preview_zoom_mode_changed)
        self.preview_panel.zoom_percent_changed.connect(self._on_preview_zoom_percent_changed)

        self.settings_panel.use_default_changed.connect(self._on_use_default_changed)
        self.settings_panel.select_printer_clicked.connect(self._on_global_printer_select)
        self.settings_panel.printer_settings_clicked.connect(self._on_open_printer_settings)
        self.settings_panel.copies_changed.connect(self._on_copies_changed)
        self.settings_panel.duplex_changed.connect(self._on_duplex_changed)
        self.settings_panel.excel_orientation_mode_changed.connect(self._on_excel_orientation_mode_changed)
        self.settings_panel.pdf_auto_rotate_changed.connect(self._on_pdf_auto_rotate_changed)
        self.settings_panel.pdf_center_changed.connect(self._on_pdf_center_changed)
        self.settings_panel.pdf_scale_mode_changed.connect(self._on_pdf_scale_mode_changed)
        self.settings_panel.pdf_scale_percent_changed.connect(self._on_pdf_scale_percent_changed)
        self.settings_panel.pdf_scale_reset_requested.connect(self._on_pdf_scale_reset_requested)
        self.settings_panel.pdf_warn_clip_changed.connect(self._on_pdf_warn_clip_changed)
        self.settings_panel.pdf_test_page_requested.connect(self._on_pdf_test_page_requested)
        self.settings_panel.rule_printer_changed.connect(self._on_rule_printer_changed)
        self.settings_panel.rule_add_requested.connect(self._on_rule_add)
        self.settings_panel.rule_remove_requested.connect(self._on_rule_remove)
        self.settings_panel.theme_changed.connect(self._on_theme_changed)
        self.settings_panel.paper_size_changed.connect(self._on_paper_size_changed)
        self.settings_panel.language_changed.connect(self._on_language_changed)
        self.settings_panel.update_check_changed.connect(self._on_update_check_changed)
        self.settings_panel.auto_update_changed.connect(self._on_auto_update_changed)

        self._job_manager.jobs_changed.connect(self._update_status)
        self._job_manager.job_updated.connect(self._update_status)
        self._job_manager.jobs_changed.connect(self._update_selection_actions)
        self._job_manager.jobs_changed.connect(self._refresh_preview_selection)

        self._context.rules_changed.connect(self._refresh_rules)
        self._context.settings_changed.connect(self._refresh_settings)

    def _apply_language(self) -> None:
        self.setWindowTitle(t("app_title"))
        self.file_menu.setTitle(t("menu_file"))
        self.help_menu.setTitle(t("menu_help"))

        self.add_files_action.setText(t("action_add_files"))
        self.add_folder_action.setText(t("action_add_folder"))
        self.apply_rules_action.setText(t("action_apply_rules"))
        self.apply_rules_force_action.setText(t("action_apply_rules_force"))
        self.exit_action.setText(t("action_exit"))

        self.start_print_action.setText(t("button_start_printing"))
        self.retry_failed_action.setText(t("button_retry_failed"))
        self.print_selected_action.setText(t("action_print_selected"))
        self.printer_selected_action.setText(t("action_set_printer"))
        self.pdf_options_action.setText(t("action_pdf_options"))
        self.excel_sheets_action.setText(t("action_excel_sheets"))
        self.move_up_action.setText(t("action_move_up"))
        self.move_down_action.setText(t("action_move_down"))
        self.delete_action.setText(t("action_delete_selected"))
        self.preview_toggle_action.setText(t("action_toggle_preview"))

        self.usage_action.setText(t("action_usage"))
        self.about_action.setText(t("action_about"))
        self.log_summary_action.setText(t("action_log_summary"))
        self.check_updates_action.setText(t("action_check_updates"))

        self.settings_panel.retranslate()
        self.file_list.retranslate()
        self.preview_panel.retranslate()
        if self._preview_window:
            self._preview_window.setWindowTitle(t("preview_title"))
        self._refresh_rules()
        self._refresh_paper_sizes()
        if self._progress_dialog:
            self._progress_dialog.retranslate()
        self._update_status()

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
            pdf_scale_mode=settings.pdf_scale_mode,
            pdf_auto_rotate=settings.pdf_auto_rotate,
            pdf_center=settings.pdf_center,
            pdf_warn_clip=settings.pdf_warn_if_clip,
            language_mode=settings.language_mode,
            update_check_enabled=settings.update_check_enabled,
            auto_update_enabled=settings.auto_update_enabled,
        )
        self.preview_panel.set_zoom(settings.preview_zoom_mode, settings.preview_zoom_percent)
        self._apply_preview_state(
            visible=settings.preview_visible,
            persist=False,
        )
        self._refresh_paper_sizes()
        self._refresh_rules()
        self._refresh_preview_selection()

    def _refresh_rules(self) -> None:
        self.settings_panel.set_printers(self._printers)
        auto_label = self._get_effective_printer_name()
        self.settings_panel.set_rules(self._context.rules, auto_label)

    def _update_status(self, *_args) -> None:
        total = self._job_manager.job_count()
        failed = sum(1 for job in self._job_manager.jobs() if job.status == JobStatus.FAILED)
        completed = sum(1 for job in self._job_manager.jobs() if job.status == JobStatus.SUCCESS)
        self.statusBar().showMessage(t("status_jobs_fmt", total=total, completed=completed, failed=failed))
        self.retry_failed_action.setEnabled(failed > 0 and not (self._executor and self._executor.isRunning()))

    def _on_file_selection_changed(self, *_args) -> None:
        self._update_selection_actions()
        self._refresh_preview_selection()

    def _refresh_preview_selection(self, *_args) -> None:
        if not self.preview_toggle_action.isChecked():
            return
        self.preview_panel.set_jobs(self.file_list.selected_jobs())

    def _update_selection_actions(self) -> None:
        if not self.file_list.isEnabled():
            self.print_selected_action.setEnabled(False)
            self.printer_selected_action.setEnabled(False)
            self.pdf_options_action.setEnabled(False)
            self.excel_sheets_action.setEnabled(False)
            self.move_up_action.setEnabled(False)
            self.move_down_action.setEnabled(False)
            self.delete_action.setEnabled(False)
            return

        jobs = self.file_list.selected_jobs()
        has_selection = bool(jobs)
        self.print_selected_action.setEnabled(has_selection)
        self.printer_selected_action.setEnabled(has_selection)
        self.delete_action.setEnabled(has_selection)
        self.pdf_options_action.setEnabled(any(job.file_type == FileType.PDF for job in jobs))
        self.excel_sheets_action.setEnabled(len(jobs) == 1 and jobs[0].file_type == FileType.EXCEL)
        can_move_up, can_move_down = self.file_list.can_move_selected()
        self.move_up_action.setEnabled(can_move_up)
        self.move_down_action.setEnabled(can_move_down)

    def _on_add_files(self) -> None:
        filter_text = "印刷できるファイル (*.pdf *.doc *.docx *.xls *.xlsx *.xlsm *.ppt *.pptx)"
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(self, t("action_add_files"), "", filter_text)
        if files:
            self._job_manager.add_files(files)

    def _on_add_folder(self) -> None:
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, t("action_add_folder"))
        if folder:
            self._job_manager.add_folder(folder, recursive=True)

    def _on_about(self) -> None:
        dialog = AboutDialog(self)
        dialog.exec()

    def _on_check_updates(self) -> None:
        self._update_manager.check_for_updates(manual=True)

    def _on_usage(self) -> None:
        dialog = UsageDialog(self)
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
            QtWidgets.QMessageBox.warning(self, t("title_excel"), str(exc))
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

    def _on_pdf_auto_rotate_changed(self, enabled: bool) -> None:
        self._context.update_setting(pdf_auto_rotate=enabled)

    def _on_pdf_center_changed(self, enabled: bool) -> None:
        self._context.update_setting(pdf_center=enabled)

    def _on_pdf_scale_mode_changed(self, mode: str) -> None:
        self._context.update_setting(pdf_scale_mode=mode)

    def _on_pdf_warn_clip_changed(self, enabled: bool) -> None:
        self._context.update_setting(pdf_warn_if_clip=enabled)

    def _on_pdf_scale_percent_changed(self, value: int) -> None:
        printer_name = self._current_printer_name()
        settings = self._context.settings
        if printer_name:
            updated = dict(settings.pdf_printer_scale)
            updated[printer_name] = value
            self._context.update_setting(pdf_printer_scale=updated)
        else:
            self._context.update_setting(pdf_scale_percent=value)

    def _on_pdf_scale_reset_requested(self) -> None:
        printer_name = self._current_printer_name()
        if not printer_name:
            return
        settings = self._context.settings
        if printer_name not in settings.pdf_printer_scale:
            return
        updated = dict(settings.pdf_printer_scale)
        updated.pop(printer_name, None)
        self._context.update_setting(pdf_printer_scale=updated)

    def _on_pdf_test_page_requested(self) -> None:
        printer_name = self._current_printer_name()
        if not printer_name:
            QtWidgets.QMessageBox.information(self, t("title_printer"), t("msg_printer_not_selected"))
            return
        printer = QtPrintSupport.QPrinter(QtPrintSupport.QPrinter.HighResolution)
        printer.setPrinterName(printer_name)
        paper_size = self._context.settings.paper_size
        if paper_size:
            apply_qt_paper_size(printer, paper_size)
        painter = QtGui.QPainter()
        if not painter.begin(printer):
            QtWidgets.QMessageBox.warning(self, t("title_print"), t("msg_pdf_test_failed"))
            return
        try:
            paper_rect = printer.paperRect(QtPrintSupport.QPrinter.DevicePixel)
            page_rect = printer.pageRect(QtPrintSupport.QPrinter.DevicePixel)
            paper_pen = QtGui.QPen(QtGui.QColor("#9CA3AF"), 2, QtCore.Qt.DashLine)
            page_pen = QtGui.QPen(QtGui.QColor("#DC2626"), 2)
            painter.setPen(paper_pen)
            painter.drawRect(paper_rect)
            painter.setPen(page_pen)
            painter.drawRect(page_rect)
            center = page_rect.center()
            painter.drawLine(center.x() - 20, center.y(), center.x() + 20, center.y())
            painter.drawLine(center.x(), center.y() - 20, center.x(), center.y() + 20)
            painter.setPen(QtGui.QColor("#111827"))
            font = painter.font()
            font.setPointSize(10)
            painter.setFont(font)
            painter.drawText(page_rect.adjusted(12, 12, -12, -12), QtCore.Qt.AlignTop | QtCore.Qt.AlignLeft,
                             t("msg_pdf_test_label"))
        finally:
            painter.end()

    def _on_pdf_options_requested(self, job_ids: list[str]) -> None:
        ids = set(job_ids)
        jobs = [job for job in self._job_manager.jobs() if job.id in ids and job.file_type == FileType.PDF]
        if not jobs:
            QtWidgets.QMessageBox.information(self, t("title_print"), t("msg_no_selected_rows"))
            return
        options, ok = PdfOptionsDialog.get_options(len(jobs), self)
        if not ok:
            return
        changed = False
        for job in jobs:
            if options.scale_mode == "default":
                job.pdf_scale_mode = None
                changed = True
            elif options.scale_mode in ("auto", "fit", "shrink", "none"):
                job.pdf_scale_mode = options.scale_mode
                changed = True

            if options.scale_percent_mode == "default":
                job.pdf_scale_percent = None
                changed = True
            elif options.scale_percent_mode == "custom":
                job.pdf_scale_percent = options.scale_percent
                changed = True

            if options.auto_rotate == "default":
                job.pdf_auto_rotate = None
                changed = True
            elif options.auto_rotate == "on":
                job.pdf_auto_rotate = True
                changed = True
            elif options.auto_rotate == "off":
                job.pdf_auto_rotate = False
                changed = True

            if options.center == "default":
                job.pdf_center = None
                changed = True
            elif options.center == "on":
                job.pdf_center = True
                changed = True
            elif options.center == "off":
                job.pdf_center = False
                changed = True
        if changed:
            self._job_manager.jobs_changed.emit()

    def _on_theme_changed(self, mode: str) -> None:
        self._context.update_setting(theme_mode=mode)
        app = QtWidgets.QApplication.instance()
        if app:
            apply_theme(app, mode)

    def _on_language_changed(self, mode: str) -> None:
        if mode == self._context.settings.language_mode:
            return
        # Avoid re-entrancy via settings_changed while language combo is updating.
        with QtCore.QSignalBlocker(self._context):
            self._context.update_setting(language_mode=mode)
        set_language(resolve_language(mode))
        self._apply_language()
        self._refresh_settings()

    def _on_update_check_changed(self, enabled: bool) -> None:
        self._context.update_setting(update_check_enabled=enabled)

    def _on_auto_update_changed(self, enabled: bool) -> None:
        self._context.update_setting(auto_update_enabled=enabled)

    def _on_preview_toggle_changed(self, checked: bool) -> None:
        self._apply_preview_state(visible=checked, persist=True)

    def _apply_preview_state(self, visible: bool, persist: bool) -> None:
        visible = bool(visible)

        with QtCore.QSignalBlocker(self.preview_toggle_action):
            self.preview_toggle_action.setChecked(visible)

        if visible:
            self._move_preview_to_window()
            if self._preview_window:
                self._preview_window.show()
                self._preview_window.raise_()
                self._preview_window.activateWindow()
            self._refresh_preview_selection()
        else:
            if self._preview_window:
                self._preview_window.hide()
            self.preview_panel.set_jobs([])

        if not persist:
            return
        updates: dict[str, object] = {}
        settings = self._context.settings
        if settings.preview_visible != visible:
            updates["preview_visible"] = visible
        if updates:
            self._context.update_setting(**updates)

    def _move_preview_to_window(self) -> None:
        self._ensure_preview_window()
        if not self._preview_window:
            return
        if self._preview_window.centralWidget() is self.preview_panel:
            return
        self._detach_preview_panel_from_parent()
        self._preview_window.setCentralWidget(self.preview_panel)
        self._preview_window.setWindowTitle(t("preview_title"))
        self.preview_panel.show()

    def _detach_preview_panel_from_parent(self) -> None:
        parent = self.preview_panel.parentWidget()
        if parent is None:
            return
        parent_layout = parent.layout()
        if parent_layout is not None:
            parent_layout.removeWidget(self.preview_panel)
        self.preview_panel.setParent(None)

    def _ensure_preview_window(self) -> None:
        if self._preview_window:
            return
        self._preview_window = PreviewWindow(self)
        self._preview_window.setWindowTitle(t("preview_title"))
        self._preview_window.closed.connect(self._on_preview_window_closed)

    def _on_preview_window_closed(self) -> None:
        if self._closing_main_window or self._ignore_preview_window_close:
            return
        if not self.preview_toggle_action.isChecked():
            return
        self._apply_preview_state(visible=False, persist=True)

    def _on_preview_zoom_mode_changed(self, mode: str) -> None:
        if mode == self._context.settings.preview_zoom_mode:
            return
        self._context.update_setting(preview_zoom_mode=mode)

    def _on_preview_zoom_percent_changed(self, value: int) -> None:
        if value == self._context.settings.preview_zoom_percent:
            return
        self._context.update_setting(preview_zoom_percent=value)

    def _on_open_printer_settings(self) -> None:
        settings = self._context.settings
        printer_name = ""
        if settings.use_default_printer:
            printer_name = self._get_default_printer_name()
        else:
            printer_name = settings.selected_printer
        if not printer_name:
            QtWidgets.QMessageBox.information(self, t("title_printer"), t("msg_printer_not_selected"))
            return
        try:
            open_printer_properties(printer_name)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, t("title_printer"), str(exc))

    def _on_column_widths_changed(self, widths: list[int]) -> None:
        if widths == self._context.settings.file_list_column_widths:
            return
        self._context.update_setting(file_list_column_widths=list(widths))

    def _on_rule_printer_changed(self, extension: str, printer: str) -> None:
        self._context.update_rule(extension, printer)
        self._job_manager.apply_rules()

    def _on_rule_add(self, extension: str) -> None:
        if extension in self._context.rules:
            QtWidgets.QMessageBox.information(self, t("title_rules"), t("msg_rule_exists"))
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
            QtWidgets.QMessageBox.information(self, t("title_print"), t("msg_no_files"))
            return
        enabled_jobs = self._job_manager.get_enabled_jobs()
        if not enabled_jobs:
            QtWidgets.QMessageBox.information(self, t("title_print"), t("msg_no_checked"))
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
            QtWidgets.QMessageBox.information(self, t("title_print"), t("msg_no_selected_rows"))
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

    def _on_print_selected_action(self) -> None:
        self._on_print_selected(self.file_list.selected_job_ids())

    def _on_printer_selected_action(self) -> None:
        job_ids = self.file_list.selected_job_ids()
        if not job_ids:
            return
        self._on_printer_selected(job_ids)

    def _on_pdf_options_action(self) -> None:
        job_ids = self.file_list.selected_job_ids()
        if not job_ids:
            return
        self._on_pdf_options_requested(job_ids)

    def _on_excel_sheets_action(self) -> None:
        jobs = self.file_list.selected_jobs()
        if len(jobs) != 1 or jobs[0].file_type != FileType.EXCEL:
            return
        self._on_excel_sheets_select(jobs[0].id)

    def _on_retry_failed(self) -> None:
        if self._executor and self._executor.isRunning():
            return
        failed_jobs = self._job_manager.get_failed_jobs()
        if not failed_jobs:
            QtWidgets.QMessageBox.information(self, t("title_retry"), t("msg_no_failed"))
            return
        self._lock_ui(True)
        self._job_manager.reset_failed_jobs()
        self._start_executor(failed_jobs)

    def _start_executor(self, jobs) -> None:
        if not self._confirm_pdf_clipping(jobs):
            self._lock_ui(False)
            return
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

    def _confirm_pdf_clipping(self, jobs) -> bool:
        settings = self._context.settings
        if not settings.pdf_warn_if_clip:
            return True
        pdf_jobs = [job for job in jobs if job.file_type == FileType.PDF]
        if not pdf_jobs:
            return True
        warned = False
        for job in pdf_jobs:
            effective_printer_name = job.printer_name or self._get_default_printer_name()
            options = resolve_pdf_options(
                job,
                settings,
                effective_printer_name=effective_printer_name,
            )
            if options.scale_mode != "none":
                continue
            if self._printer_has_margins(effective_printer_name, job.paper_size):
                warned = True
                break
        if not warned:
            return True
        result = QtWidgets.QMessageBox.warning(
            self,
            t("title_print"),
            t("msg_pdf_clip_warning"),
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
        )
        return result == QtWidgets.QMessageBox.Yes

    def _printer_has_margins(self, printer_name: str, paper_size: str) -> bool:
        if not printer_name:
            printer_name = self._get_default_printer_name()
        if not printer_name:
            return False
        printer = QtPrintSupport.QPrinter(QtPrintSupport.QPrinter.HighResolution)
        printer.setPrinterName(printer_name)
        if paper_size:
            apply_qt_paper_size(printer, paper_size)
        page_rect = printer.pageRect(QtPrintSupport.QPrinter.DevicePixel)
        paper_rect = printer.paperRect(QtPrintSupport.QPrinter.DevicePixel)
        if page_rect.isNull() or paper_rect.isNull():
            return False
        return (paper_rect.width() - page_rect.width() > 2) or (paper_rect.height() - page_rect.height() > 2)

    def _on_progress(self, completed: int, total: int, current: str) -> None:
        if self._progress_dialog:
            self._progress_dialog.update_progress(completed, total, current)
        self._update_taskbar_progress(completed, total)

    def _on_job_failed(self, job_id: str) -> None:
        job = self._job_manager.find_job_by_id(job_id)
        file_name = job.file_name if job else "-"
        summary = job.summary if job and job.summary else t("msg_print_failed")
        message = job.message if job and job.message else ""
        dialog = QtWidgets.QMessageBox(self)
        dialog.setIcon(QtWidgets.QMessageBox.Warning)
        dialog.setWindowTitle(t("title_print_failed"))
        if message:
            dialog.setText(f"{file_name}\n{summary}\n\n{message}")
        else:
            dialog.setText(f"{file_name}\n{summary}")
        log_button = dialog.addButton(t("btn_open_log"), QtWidgets.QMessageBox.ActionRole)
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
        self.preview_panel.setEnabled(not locked)
        self.settings_panel.setEnabled(not locked)
        self.command_bar.setEnabled(not locked)
        self.menuBar().setEnabled(not locked)
        self._update_selection_actions()

    def _get_default_printer_name(self) -> str:
        try:
            return get_default_printer_name()
        except Exception:
            return ""

    def _get_effective_printer_name(self) -> str:
        settings = self._context.settings
        if settings.use_default_printer:
            return self._get_default_printer_name() or "-"
        return settings.selected_printer or "-"

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
        current = settings.paper_size
        if sizes and current and not is_supported_name(current, sizes):
            with QtCore.QSignalBlocker(self._context):
                self._context.update_setting(paper_size="")
            self._job_manager.apply_settings_to_jobs()
            current = ""
        if sizes:
            enabled = True
            tooltip = ""
        else:
            enabled = False
            tooltip = t("paper_size_unavailable")
        self.settings_panel.set_paper_sizes(sizes, current, enabled, tooltip)
        self._refresh_pdf_scale()

    def _current_printer_name(self) -> str:
        settings = self._context.settings
        if settings.use_default_printer:
            return self._get_default_printer_name()
        return settings.selected_printer

    def _refresh_pdf_scale(self) -> None:
        settings = self._context.settings
        printer_name = self._current_printer_name()
        if printer_name:
            percent = settings.pdf_printer_scale.get(printer_name, settings.pdf_scale_percent)
            has_override = printer_name in settings.pdf_printer_scale
            tooltip = t("settings_pdf_scale_tooltip_fmt", printer=printer_name)
            enabled = True
        else:
            percent = settings.pdf_scale_percent
            has_override = False
            tooltip = t("msg_printer_not_selected")
            enabled = False
        self.settings_panel.set_pdf_scale(percent, enabled, has_override, tooltip)

    def _on_log_summary(self) -> None:
        items = []
        for job in self._job_manager.get_failed_jobs():
            summary = job.summary or t("msg_print_failed")
            detail = job.message or ""
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
            t("msg_excel_orientation_checking"), "", 0, 0, self
        )
        self._orientation_progress.setWindowTitle(t("title_checking"))
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
        QtWidgets.QMessageBox.warning(self, t("title_excel"), message)
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
                t("title_print"),
                t("msg_printing_in_progress"),
            )
            event.ignore()
            return
        self._closing_main_window = True
        if self._preview_window:
            self._ignore_preview_window_close = True
            self._preview_window.close()
            self._ignore_preview_window_close = False
        self.preview_panel.release_cached_previews(clear_shared_blob_cache=True)
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
