from __future__ import annotations

from typing import List

from PySide6 import QtCore, QtGui, QtWidgets

from app.controller.job_manager import JobManager
from app.model.print_job import JobStatus, FileType
from app.i18n import t


class JobTableModel(QtCore.QAbstractTableModel):
    def __init__(self, job_manager: JobManager) -> None:
        super().__init__()
        self._job_manager = job_manager
        self._job_manager.jobs_changed.connect(self._on_jobs_changed)
        self._job_manager.job_updated.connect(self._on_job_updated)
        self._status_icons = self._build_status_icons()
        self._status_colors = {
            JobStatus.WAITING: QtGui.QColor("#9AA0A6"),
            JobStatus.PRINTING: QtGui.QColor("#3B82F6"),
            JobStatus.SUCCESS: QtGui.QColor("#16A34A"),
            JobStatus.FAILED: QtGui.QColor("#DC2626"),
            JobStatus.CANCELLED: QtGui.QColor("#6B7280"),
            JobStatus.SKIPPED: QtGui.QColor("#6B7280"),
        }

    def rowCount(self, parent: QtCore.QModelIndex = QtCore.QModelIndex()) -> int:
        return self._job_manager.job_count()

    def columnCount(self, parent: QtCore.QModelIndex = QtCore.QModelIndex()) -> int:
        return len(self._headers())

    def headerData(self, section: int, orientation: QtCore.Qt.Orientation, role: int = QtCore.Qt.DisplayRole):
        if role == QtCore.Qt.DisplayRole and orientation == QtCore.Qt.Horizontal:
            return self._headers()[section]
        return None

    def data(self, index: QtCore.QModelIndex, role: int = QtCore.Qt.DisplayRole):
        if not index.isValid():
            return None
        job = self._job_manager.get_job(index.row())
        column = index.column()

        if role == QtCore.Qt.CheckStateRole and column == 0:
            return QtCore.Qt.Checked if job.enabled else QtCore.Qt.Unchecked

        if role == QtCore.Qt.DisplayRole:
            if column == 1:
                return job.file_name
            if column == 2:
                return job.file_path
            if column == 3:
                return self._format_file_type(job.file_type)
            if column == 4:
                return self._format_label(job)
            if column == 5:
                return self._format_sheets(job)
            if column == 6:
                return self._format_printer(job)
            if column == 7:
                return self._format_status(job)

        if role == QtCore.Qt.DecorationRole and column == 7:
            return self._status_icons.get(job.status)

        if role == QtCore.Qt.ForegroundRole and column == 7:
            return self._status_colors.get(job.status)

        if role == QtCore.Qt.ToolTipRole:
            if column == 5 and job.excel_sheets:
                return "\n".join(job.excel_sheets)
            if job.message:
                return job.message

        return None

    def flags(self, index: QtCore.QModelIndex):
        if not index.isValid():
            return QtCore.Qt.ItemIsDropEnabled
        flags = (
            QtCore.Qt.ItemIsSelectable
            | QtCore.Qt.ItemIsEnabled
            | QtCore.Qt.ItemIsDragEnabled
            | QtCore.Qt.ItemIsDropEnabled
        )
        if index.column() == 0:
            flags |= QtCore.Qt.ItemIsUserCheckable
        return flags

    def setData(self, index: QtCore.QModelIndex, value, role: int = QtCore.Qt.EditRole) -> bool:
        if not index.isValid():
            return False
        if index.column() == 0 and role == QtCore.Qt.CheckStateRole:
            job = self._job_manager.get_job(index.row())
            enabled = value == QtCore.Qt.Checked
            self._job_manager.set_job_enabled(job.id, enabled)
            return True
        return False

    def supportedDropActions(self) -> QtCore.Qt.DropActions:
        return QtCore.Qt.MoveAction

    def mimeTypes(self) -> list[str]:
        return ["application/x-oss-print-row"]

    def mimeData(self, indexes: list[QtCore.QModelIndex]) -> QtCore.QMimeData:
        mime_data = QtCore.QMimeData()
        rows = sorted({index.row() for index in indexes})
        if rows:
            mime_data.setData(self.mimeTypes()[0], str(rows[0]).encode("utf-8"))
        return mime_data

    def dropMimeData(
        self,
        data: QtCore.QMimeData,
        action: QtCore.Qt.DropAction,
        row: int,
        column: int,
        parent: QtCore.QModelIndex,
    ) -> bool:
        if action == QtCore.Qt.IgnoreAction:
            return True
        if not data.hasFormat(self.mimeTypes()[0]):
            return False
        source_row = int(bytes(data.data(self.mimeTypes()[0])).decode("utf-8"))
        destination_row = row
        if destination_row == -1:
            destination_row = self.rowCount()
        if destination_row > source_row:
            destination_row -= 1
        if destination_row == source_row:
            return False
        self._job_manager.move_job(source_row, destination_row)
        return True

    def sort(self, column: int, order: QtCore.Qt.SortOrder = QtCore.Qt.AscendingOrder) -> None:
        descending = order == QtCore.Qt.DescendingOrder
        self._job_manager.sort_jobs(column, descending)

    def retranslate(self) -> None:
        self.headerDataChanged.emit(QtCore.Qt.Horizontal, 0, self.columnCount() - 1)
        self.layoutChanged.emit()

    def _headers(self) -> list[str]:
        return [
            t("table_print"),
            t("table_file"),
            t("table_path"),
            t("table_type"),
            t("table_label"),
            t("table_sheet"),
            t("table_printer"),
            t("table_status"),
        ]

    def _format_file_type(self, file_type: FileType) -> str:
        if file_type == FileType.PDF:
            return "PDF"
        if file_type == FileType.WORD:
            return "Word"
        if file_type == FileType.EXCEL:
            return "Excel"
        if file_type == FileType.PPT:
            return "PowerPoint"
        return t("label_unknown")

    def _format_label(self, job) -> str:
        ext = job.extension
        if ext == ".pdf":
            return t("label_pdf")
        if ext in (".doc", ".docx"):
            return t("label_word")
        if ext in (".xls", ".xlsx", ".xlsm"):
            return t("label_excel")
        if ext in (".ppt", ".pptx"):
            return t("label_ppt")
        return t("label_unknown")

    def _format_sheets(self, job) -> str:
        if job.file_type != FileType.EXCEL:
            return ""
        if not job.excel_sheets:
            return t("label_all_sheets")
        joined = ", ".join(job.excel_sheets)
        return joined if len(joined) <= 20 else f"{joined[:17]}..."

    def _format_printer(self, job) -> str:
        return job.printer_name or t("label_auto")

    def _format_status(self, job) -> str:
        if job.status == JobStatus.FAILED and job.message:
            summary = job.summary or job.message
            return f"{self._status_text(job.status)}: {summary}"
        return self._status_text(job.status)

    def _status_text(self, status: JobStatus) -> str:
        mapping = {
            JobStatus.WAITING: t("status_waiting"),
            JobStatus.PRINTING: t("status_printing"),
            JobStatus.SUCCESS: t("status_success"),
            JobStatus.FAILED: t("status_failed"),
            JobStatus.CANCELLED: t("status_cancelled"),
            JobStatus.SKIPPED: t("status_skipped"),
        }
        return mapping.get(status, status.value)

    def _on_jobs_changed(self) -> None:
        self.beginResetModel()
        self.endResetModel()

    def _on_job_updated(self, job_id: str) -> None:
        for row in range(self._job_manager.job_count()):
            if self._job_manager.get_job(row).id == job_id:
                top_left = self.index(row, 0)
                bottom_right = self.index(row, self.columnCount() - 1)
                self.dataChanged.emit(top_left, bottom_right, [])
                break

    def _build_status_icons(self) -> dict[JobStatus, QtGui.QIcon]:
        return {
            JobStatus.WAITING: self._dot_icon(QtGui.QColor("#9AA0A6")),
            JobStatus.PRINTING: self._spinner_icon(QtGui.QColor("#3B82F6")),
            JobStatus.SUCCESS: self._dot_icon(QtGui.QColor("#16A34A")),
            JobStatus.FAILED: self._dot_icon(QtGui.QColor("#DC2626")),
            JobStatus.CANCELLED: self._dot_icon(QtGui.QColor("#6B7280")),
            JobStatus.SKIPPED: self._dot_icon(QtGui.QColor("#6B7280")),
        }

    @staticmethod
    def _dot_icon(color: QtGui.QColor) -> QtGui.QIcon:
        pix = QtGui.QPixmap(10, 10)
        pix.fill(QtCore.Qt.transparent)
        painter = QtGui.QPainter(pix)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        painter.setBrush(color)
        painter.setPen(QtCore.Qt.NoPen)
        painter.drawEllipse(1, 1, 8, 8)
        painter.end()
        return QtGui.QIcon(pix)

    @staticmethod
    def _spinner_icon(color: QtGui.QColor) -> QtGui.QIcon:
        pix = QtGui.QPixmap(12, 12)
        pix.fill(QtCore.Qt.transparent)
        painter = QtGui.QPainter(pix)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        pen = QtGui.QPen(color, 2)
        painter.setPen(pen)
        painter.drawArc(1, 1, 10, 10, 30 * 16, 300 * 16)
        painter.end()
        return QtGui.QIcon(pix)


class FileListView(QtWidgets.QTableView):
    files_dropped = QtCore.Signal(list)
    printer_requested = QtCore.Signal(str)
    excel_sheets_requested = QtCore.Signal(str)
    print_selected_requested = QtCore.Signal(list)
    printer_selected_requested = QtCore.Signal(list)

    def __init__(self, job_manager: JobManager, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._job_manager = job_manager
        self.setAcceptDrops(True)
        self.setDragEnabled(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QtWidgets.QAbstractItemView.DragDrop)
        self.setDefaultDropAction(QtCore.Qt.MoveAction)
        self.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.setAlternatingRowColors(True)
        self.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.setEditTriggers(
            QtWidgets.QAbstractItemView.SelectedClicked
            | QtWidgets.QAbstractItemView.DoubleClicked
        )

        self._model = JobTableModel(job_manager)
        self.setModel(self._model)
        self.setSortingEnabled(True)
        self.horizontalHeader().setStretchLastSection(True)
        self.horizontalHeader().setSectionsClickable(True)
        self.verticalHeader().setDefaultSectionSize(28)
        self.verticalHeader().setVisible(False)
        self.setShowGrid(True)
        self.doubleClicked.connect(self._on_double_click)
        self.customContextMenuRequested.connect(self._on_context_menu)

    def retranslate(self) -> None:
        self._model.retranslate()

    def dragEnterEvent(self, event: QtGui.QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event: QtGui.QDragMoveEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event: QtGui.QDropEvent) -> None:
        if event.mimeData().hasUrls():
            paths: List[str] = []
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    paths.append(url.toLocalFile())
            if paths:
                self.files_dropped.emit(paths)
            event.acceptProposedAction()
        else:
            super().dropEvent(event)

    def _on_double_click(self, index: QtCore.QModelIndex) -> None:
        if not index.isValid():
            return
        if index.column() == 6:
            job = self._job_manager.get_job(index.row())
            self.printer_requested.emit(job.id)

    def mousePressEvent(self, event: QtGui.QMouseEvent) -> None:
        if event.button() == QtCore.Qt.LeftButton:
            pos = event.position().toPoint() if hasattr(event, "position") else event.pos()
            index = self.indexAt(pos)
            if index.isValid() and index.column() == 0:
                current = self.model().data(index, QtCore.Qt.CheckStateRole)
                new_state = QtCore.Qt.Unchecked if current == QtCore.Qt.Checked else QtCore.Qt.Checked
                rows = [idx.row() for idx in self.selectionModel().selectedRows()]
                if index.row() not in rows:
                    rows = [index.row()]
                job_ids = [self._job_manager.get_job(row).id for row in rows]
                self._job_manager.set_jobs_enabled(job_ids, new_state == QtCore.Qt.Checked)
                self.selectionModel().select(
                    index,
                    QtCore.QItemSelectionModel.ClearAndSelect | QtCore.QItemSelectionModel.Rows,
                )
                return
        super().mousePressEvent(event)

    def keyPressEvent(self, event: QtGui.QKeyEvent) -> None:
        if event.matches(QtGui.QKeySequence.SelectAll):
            self.selectAll()
            return
        if event.key() == QtCore.Qt.Key_Space:
            rows = [idx.row() for idx in self.selectionModel().selectedRows()]
            if rows:
                jobs = [self._job_manager.get_job(row) for row in rows]
                enable = not all(job.enabled for job in jobs)
                self._job_manager.set_jobs_enabled([job.id for job in jobs], enable)
                return
        if event.key() == QtCore.Qt.Key_Delete:
            self._confirm_and_remove_selected()
            return
        super().keyPressEvent(event)

    def _on_context_menu(self, pos: QtCore.QPoint) -> None:
        selected_ids = self._selected_job_ids()
        excel_job = self._selected_excel_job()
        menu = QtWidgets.QMenu(self)
        print_selected_action = menu.addAction(t("context_print_selected"))
        printer_selected_action = menu.addAction(t("context_printer_select"))
        sheet_action = menu.addAction(t("context_excel_sheets"))
        enable_action = menu.addAction(t("context_enable"))
        disable_action = menu.addAction(t("context_disable"))
        delete_action = menu.addAction(t("context_delete"))
        if not selected_ids:
            print_selected_action.setEnabled(False)
            printer_selected_action.setEnabled(False)
            sheet_action.setEnabled(False)
            enable_action.setEnabled(False)
            disable_action.setEnabled(False)
            delete_action.setEnabled(False)
        if excel_job is None:
            sheet_action.setEnabled(False)
        action = menu.exec(self.viewport().mapToGlobal(pos))
        if action == print_selected_action:
            self.print_selected_requested.emit(selected_ids)
        elif action == printer_selected_action:
            self.printer_selected_requested.emit(selected_ids)
        if action == sheet_action and excel_job is not None:
            self.excel_sheets_requested.emit(excel_job.id)
        if action == enable_action:
            self._job_manager.set_jobs_enabled(selected_ids, True)
        elif action == disable_action:
            self._job_manager.set_jobs_enabled(selected_ids, False)
        elif action == delete_action:
            self._confirm_and_remove_selected()

    def _selected_job_ids(self) -> list[str]:
        rows = [idx.row() for idx in self.selectionModel().selectedRows()]
        return [self._job_manager.get_job(row).id for row in rows]

    def _selected_excel_job(self):
        rows = [idx.row() for idx in self.selectionModel().selectedRows()]
        if len(rows) != 1:
            return None
        job = self._job_manager.get_job(rows[0])
        return job if job.file_type == FileType.EXCEL else None

    def _confirm_and_remove_selected(self) -> None:
        selected_ids = self._selected_job_ids()
        if not selected_ids:
            return
        result = QtWidgets.QMessageBox.question(
            self,
            t("title_delete"),
            t("msg_delete_confirm_fmt", count=len(selected_ids)),
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
        )
        if result == QtWidgets.QMessageBox.Yes:
            self._job_manager.remove_jobs(selected_ids)
