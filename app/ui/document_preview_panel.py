from __future__ import annotations

from dataclasses import dataclass
import gc
import os
from pathlib import Path
import signal
import tempfile
import threading
import time

from PySide6 import QtCore, QtGui, QtWidgets

from app.i18n import t
from app.model.print_job import FileType, PrintJob


_IMAGE_CACHE_LIMIT = 120
_FULLSCREEN_CACHE_LIMIT = 48
_SCALED_ICON_CACHE_LIMIT = 520
_BLOB_CACHE_LIMIT_ITEMS = 192
_BLOB_CACHE_LIMIT_BYTES = 160 * 1024 * 1024
_BLOB_IMAGE_FORMAT = "PNG"


class _FirstPageBlobCache:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._cache: dict[tuple[str, int, int, int], bytes] = {}
        self._order: list[tuple[str, int, int, int]] = []
        self._total_bytes = 0

    @staticmethod
    def _edge_bucket(target_long_edge: int) -> int:
        try:
            edge = int(target_long_edge)
        except Exception:
            edge = 640
        edge = max(320, min(4096, edge))
        return int(round(edge / 64.0) * 64)

    @classmethod
    def _cache_key(
        cls,
        signature: tuple[str, int, int],
        target_long_edge: int,
    ) -> tuple[str, int, int, int]:
        return (
            str(signature[0]),
            int(signature[1]),
            int(signature[2]),
            cls._edge_bucket(target_long_edge),
        )

    def _touch_unlocked(self, key: tuple[str, int, int, int]) -> None:
        try:
            self._order.remove(key)
        except ValueError:
            pass
        self._order.append(key)

    def load_image(
        self,
        signature: tuple[str, int, int] | None,
        target_long_edge: int,
    ) -> QtGui.QImage | None:
        if signature is None:
            return None
        key = self._cache_key(signature, target_long_edge)
        with self._lock:
            blob = self._cache.get(key)
            if blob is None:
                return None
            self._touch_unlocked(key)
            payload = bytes(blob)
        image = QtGui.QImage.fromData(payload, _BLOB_IMAGE_FORMAT)
        if image.isNull():
            self.remove(signature, target_long_edge)
            return None
        return image

    def remove(
        self,
        signature: tuple[str, int, int],
        target_long_edge: int,
    ) -> None:
        key = self._cache_key(signature, target_long_edge)
        with self._lock:
            blob = self._cache.pop(key, None)
            if blob is not None:
                self._total_bytes = max(0, self._total_bytes - len(blob))
            try:
                self._order.remove(key)
            except ValueError:
                pass

    def store_image(
        self,
        signature: tuple[str, int, int] | None,
        target_long_edge: int,
        image: QtGui.QImage | None,
    ) -> None:
        if signature is None or image is None or image.isNull():
            return
        encoded = QtCore.QByteArray()
        buffer = QtCore.QBuffer(encoded)
        if not buffer.open(QtCore.QIODevice.WriteOnly):
            return
        try:
            if not image.save(buffer, _BLOB_IMAGE_FORMAT):
                return
        finally:
            buffer.close()
        payload = bytes(encoded)
        if not payload:
            return
        key = self._cache_key(signature, target_long_edge)
        with self._lock:
            previous = self._cache.pop(key, None)
            if previous is not None:
                self._total_bytes = max(0, self._total_bytes - len(previous))
            self._cache[key] = payload
            self._total_bytes += len(payload)
            self._touch_unlocked(key)
            while self._order and (
                len(self._order) > _BLOB_CACHE_LIMIT_ITEMS
                or self._total_bytes > _BLOB_CACHE_LIMIT_BYTES
            ):
                oldest = self._order.pop(0)
                oldest_blob = self._cache.pop(oldest, None)
                if oldest_blob is not None:
                    self._total_bytes = max(0, self._total_bytes - len(oldest_blob))

    def clear(self) -> None:
        with self._lock:
            self._cache.clear()
            self._order.clear()
            self._total_bytes = 0


_FIRST_PAGE_BLOB_CACHE = _FirstPageBlobCache()


def clear_preview_blob_cache() -> None:
    _FIRST_PAGE_BLOB_CACHE.clear()


@dataclass(frozen=True)
class PreviewRequest:
    job_id: str
    file_path: str
    file_type: FileType
    signature: tuple[str, int, int] | None


class PdfPreviewWorker(QtCore.QThread):
    thumbnail_ready = QtCore.Signal(str, object, object, int)
    batch_finished = QtCore.Signal(int)

    def __init__(
        self,
        requests: list[PreviewRequest],
        token: int,
        target_long_edge: int = 640,
        parent: QtCore.QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._requests = list(requests)
        self._token = token
        self._target_long_edge = target_long_edge

    def run(self) -> None:
        try:
            import fitz  # type: ignore
        except Exception:
            self.batch_finished.emit(self._token)
            return

        for req in self._requests:
            if self.isInterruptionRequested():
                break
            image = _render_request_preview(req, self._target_long_edge, fitz)
            self.thumbnail_ready.emit(req.job_id, req.signature, image, self._token)

        self.batch_finished.emit(self._token)


class SinglePdfPreviewWorker(QtCore.QThread):
    image_ready = QtCore.Signal(str, object, object)

    def __init__(
        self,
        job_id: str,
        file_path: str,
        file_type: FileType,
        signature: tuple[str, int, int] | None,
        target_long_edge: int,
        parent: QtCore.QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._job_id = job_id
        self._file_path = file_path
        self._file_type = file_type
        self._signature = signature
        self._target_long_edge = target_long_edge

    def run(self) -> None:
        try:
            import fitz  # type: ignore
        except Exception:
            self.image_ready.emit(self._job_id, self._signature, None)
            return
        if self.isInterruptionRequested():
            return
        req = PreviewRequest(
            job_id=self._job_id,
            file_path=self._file_path,
            file_type=self._file_type,
            signature=self._signature,
        )
        image = _render_request_preview(req, self._target_long_edge, fitz)
        self.image_ready.emit(self._job_id, self._signature, image)


def _render_pdf_preview(file_path: str, target_long_edge: int, fitz_module) -> QtGui.QImage | None:
    doc = None
    try:
        doc = fitz_module.open(file_path)
        if int(getattr(doc, "page_count", 0) or 0) <= 0:
            return None
        page = doc.load_page(0)
        rect = page.rect
        if rect.width <= 0 or rect.height <= 0:
            return None
        base = max(float(rect.width), float(rect.height))
        zoom = max(0.2, min(5.0, float(target_long_edge) / base))
        pix = page.get_pixmap(matrix=fitz_module.Matrix(zoom, zoom), alpha=False)

        if pix.n == 1:
            image_format = QtGui.QImage.Format_Grayscale8
        elif pix.n >= 4:
            image_format = QtGui.QImage.Format_RGBA8888
        else:
            image_format = QtGui.QImage.Format_RGB888

        image = QtGui.QImage(pix.samples, pix.width, pix.height, pix.stride, image_format)
        if image.isNull():
            return None
        return image.copy()
    except Exception:
        return None
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass


def _render_request_preview(req: PreviewRequest, target_long_edge: int, fitz_module) -> QtGui.QImage | None:
    cached = _FIRST_PAGE_BLOB_CACHE.load_image(req.signature, target_long_edge)
    if cached is not None and not cached.isNull():
        return cached

    rendered_image: QtGui.QImage | None = None
    if req.file_type == FileType.PDF:
        rendered_image = _render_pdf_preview(req.file_path, target_long_edge, fitz_module)
    elif req.file_type in {FileType.WORD, FileType.EXCEL, FileType.PPT}:
        exported_pdf = _export_office_to_pdf(req.file_path, req.file_type)
        if not exported_pdf:
            return None
        try:
            rendered_image = _render_pdf_preview(exported_pdf, target_long_edge, fitz_module)
        finally:
            try:
                os.remove(exported_pdf)
            except Exception:
                pass
    else:
        return None

    if rendered_image is not None and not rendered_image.isNull():
        _FIRST_PAGE_BLOB_CACHE.store_image(req.signature, target_long_edge, rendered_image)
    return rendered_image


def _office_process_pid(app, win32process_module) -> int | None:
    if app is None or win32process_module is None:
        return None
    hwnd = None
    for attr in ("Hwnd", "HWND"):
        try:
            value = getattr(app, attr)
        except Exception:
            value = None
        if value:
            try:
                hwnd = int(value)
            except Exception:
                hwnd = None
            if hwnd and hwnd > 0:
                break
    if not hwnd:
        return None
    try:
        _, pid = win32process_module.GetWindowThreadProcessId(hwnd)
    except Exception:
        return None
    try:
        pid_int = int(pid)
    except Exception:
        return None
    return pid_int if pid_int > 0 else None


def _is_process_alive(pid: int | None) -> bool:
    if not pid or pid <= 0:
        return False
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False
    except Exception:
        return True
    return True


def _wait_office_exit_or_terminate(pid: int | None, timeout_seconds: float = 1.5) -> None:
    if not pid or pid <= 0:
        return
    deadline = time.monotonic() + max(0.2, float(timeout_seconds))
    while time.monotonic() < deadline:
        if not _is_process_alive(pid):
            return
        time.sleep(0.05)
    if not _is_process_alive(pid):
        return
    sigterm = getattr(signal, "SIGTERM", None)
    if sigterm is None:
        return
    try:
        os.kill(pid, sigterm)
    except Exception:
        pass


def _export_office_to_pdf(file_path: str, file_type: FileType) -> str | None:
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception:
        return None
    try:
        import win32process  # type: ignore
    except Exception:
        win32process = None

    tmp = tempfile.NamedTemporaryFile(prefix="raku_preview_", suffix=".pdf", delete=False)
    pdf_path = tmp.name
    tmp.close()

    app = None
    doc = None
    workbook = None
    presentation = None
    exported = False
    app_pid: int | None = None

    pythoncom.CoInitialize()
    try:
        if file_type == FileType.WORD:
            app = win32com.client.DispatchEx("Word.Application")
            app_pid = _office_process_pid(app, win32process)
            app.Visible = False
            if hasattr(app, "DisplayAlerts"):
                app.DisplayAlerts = False
            doc = app.Documents.Open(file_path, ReadOnly=True)
            # 17 = wdExportFormatPDF
            doc.ExportAsFixedFormat(pdf_path, 17)
            exported = True

        elif file_type == FileType.EXCEL:
            app = win32com.client.DispatchEx("Excel.Application")
            app_pid = _office_process_pid(app, win32process)
            app.Visible = False
            if hasattr(app, "DisplayAlerts"):
                app.DisplayAlerts = False
            workbook = app.Workbooks.Open(file_path, ReadOnly=True)
            # 0 = xlTypePDF
            workbook.ExportAsFixedFormat(0, pdf_path)
            exported = True

        elif file_type == FileType.PPT:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            app_pid = _office_process_pid(app, win32process)
            app.Visible = False
            if hasattr(app, "DisplayAlerts"):
                app.DisplayAlerts = False
            presentation = app.Presentations.Open(file_path, WithWindow=False)
            try:
                # 2 = ppFixedFormatTypePDF
                presentation.ExportAsFixedFormat(pdf_path, 2)
            except Exception:
                # 32 = ppSaveAsPDF
                presentation.SaveAs(pdf_path, 32)
            exported = True
        else:
            return None
    except Exception:
        exported = False
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if workbook is not None:
                workbook.Close(False)
        except Exception:
            pass
        try:
            if presentation is not None:
                presentation.Close()
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        try:
            if doc is not None:
                del doc
        except Exception:
            pass
        try:
            if workbook is not None:
                del workbook
        except Exception:
            pass
        try:
            if presentation is not None:
                del presentation
        except Exception:
            pass
        try:
            if app is not None:
                del app
        except Exception:
            pass
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        _wait_office_exit_or_terminate(app_pid)

    try:
        if exported and Path(pdf_path).exists() and Path(pdf_path).stat().st_size > 0:
            return pdf_path
    except Exception:
        pass

    try:
        os.remove(pdf_path)
    except Exception:
        pass
    return None


class FullscreenPreviewDialog(QtWidgets.QDialog):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowFlags(QtCore.Qt.Window | QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
        self._source_pixmap: QtGui.QPixmap | None = None

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(28, 20, 28, 20)
        layout.setSpacing(8)

        self.title_label = QtWidgets.QLabel("")
        self.title_label.setObjectName("fullscreenTitle")
        title_font = self.title_label.font()
        title_font.setPointSize(max(12, title_font.pointSize() + 2))
        title_font.setBold(True)
        self.title_label.setFont(title_font)
        self.title_label.setAlignment(QtCore.Qt.AlignCenter)

        self.subtitle_label = QtWidgets.QLabel("")
        self.subtitle_label.setObjectName("fullscreenSubtitle")
        self.subtitle_label.setAlignment(QtCore.Qt.AlignCenter)
        self.subtitle_label.setWordWrap(True)

        self.preview_label = QtWidgets.QLabel("")
        self.preview_label.setObjectName("fullscreenPreview")
        self.preview_label.setAlignment(QtCore.Qt.AlignCenter)
        self.preview_label.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)

        self.status_label = QtWidgets.QLabel("")
        self.status_label.setObjectName("fullscreenStatus")
        self.status_label.setAlignment(QtCore.Qt.AlignCenter)

        self.hint_label = QtWidgets.QLabel("")
        self.hint_label.setObjectName("fullscreenHint")
        self.hint_label.setAlignment(QtCore.Qt.AlignCenter)

        layout.addWidget(self.title_label)
        layout.addWidget(self.subtitle_label)
        layout.addWidget(self.preview_label, 1)
        layout.addWidget(self.status_label)
        layout.addWidget(self.hint_label)

        for widget in (
            self.title_label,
            self.subtitle_label,
            self.preview_label,
            self.status_label,
            self.hint_label,
        ):
            widget.setAttribute(QtCore.Qt.WA_TransparentForMouseEvents, True)

        self._apply_theme_style()

    def set_title(self, title: str, subtitle: str) -> None:
        self.title_label.setText(title)
        self.subtitle_label.setText(subtitle)

    def set_hint(self, hint: str) -> None:
        self.hint_label.setText(hint)

    def set_status(self, status: str) -> None:
        self.status_label.setText(status)

    def set_image(self, image: QtGui.QImage) -> None:
        if image.isNull():
            return
        self._source_pixmap = QtGui.QPixmap.fromImage(image)
        self._update_scaled_preview()

    def set_icon(self, icon: QtGui.QIcon) -> None:
        pixmap = icon.pixmap(QtCore.QSize(480, 480))
        if pixmap.isNull():
            pixmap = self.style().standardIcon(QtWidgets.QStyle.SP_FileIcon).pixmap(480, 480)
        self._source_pixmap = pixmap
        self._update_scaled_preview()

    def resizeEvent(self, event: QtGui.QResizeEvent) -> None:
        super().resizeEvent(event)
        self._update_scaled_preview()

    def mousePressEvent(self, event: QtGui.QMouseEvent) -> None:
        self.accept()
        event.accept()

    def keyPressEvent(self, event: QtGui.QKeyEvent) -> None:
        if event.key() in (
            QtCore.Qt.Key_Escape,
            QtCore.Qt.Key_Return,
            QtCore.Qt.Key_Enter,
            QtCore.Qt.Key_Space,
        ):
            self.accept()
            event.accept()
            return
        super().keyPressEvent(event)

    def _update_scaled_preview(self) -> None:
        if self._source_pixmap is None or self._source_pixmap.isNull():
            self.preview_label.clear()
            return
        area = self.preview_label.size()
        if area.width() <= 10 or area.height() <= 10:
            return
        scaled = self._source_pixmap.scaled(
            area,
            QtCore.Qt.KeepAspectRatio,
            QtCore.Qt.SmoothTransformation,
        )
        self.preview_label.setPixmap(scaled)

    @staticmethod
    def _rgba(color: QtGui.QColor, alpha: int) -> str:
        tone = QtGui.QColor(color)
        tone.setAlpha(max(0, min(255, int(alpha))))
        return tone.name(QtGui.QColor.HexArgb)

    def _apply_theme_style(self) -> None:
        background_color = QtGui.QColor(0, 0, 0)
        text_color = QtGui.QColor("#FFFFFF")
        self.setStyleSheet(
            f"""
            FullscreenPreviewDialog {{
                background: {self._rgba(background_color, 236)};
                color: {text_color.name()};
            }}
            QLabel#fullscreenTitle {{
                color: {text_color.name()};
            }}
            QLabel#fullscreenSubtitle {{
                color: {text_color.name()};
            }}
            QLabel#fullscreenHint {{
                color: {text_color.name()};
            }}
            QLabel#fullscreenStatus {{
                color: {text_color.name()};
                font-weight: 600;
            }}
            """
        )


class DocumentPreviewPanel(QtWidgets.QWidget):
    zoom_mode_changed = QtCore.Signal(str)
    zoom_percent_changed = QtCore.Signal(int)

    def __init__(
        self,
        preview_zoom_mode: str = "fit",
        preview_zoom_percent: int = 100,
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._max_items = 8
        self._selected_total = 0
        self._request_token = 0
        self._worker: PdfPreviewWorker | None = None
        self._workers: set[PdfPreviewWorker] = set()
        self._pending_jobs: list[PrintJob] | None = None
        self._fullscreen_worker: SinglePdfPreviewWorker | None = None
        self._fullscreen_dialog: FullscreenPreviewDialog | None = None
        self._fullscreen_job_id = ""

        self._jobs_by_id: dict[str, PrintJob] = {}
        self._items_by_job: dict[str, QtWidgets.QListWidgetItem] = {}
        self._images_by_job: dict[str, QtGui.QImage] = {}
        self._image_cache: dict[tuple[str, int, int], QtGui.QImage] = {}
        self._image_cache_order: list[tuple[str, int, int]] = []
        self._fullscreen_cache: dict[tuple[str, int, int], QtGui.QImage] = {}
        self._fullscreen_cache_order: list[tuple[str, int, int]] = []
        self._scaled_icon_cache: dict[tuple[str, int, int, int], QtGui.QIcon] = {}
        self._scaled_icon_order: list[tuple[str, int, int, int]] = []
        self._icon_provider = QtWidgets.QFileIconProvider()

        self._layout_timer = QtCore.QTimer(self)
        self._layout_timer.setSingleShot(True)
        self._layout_timer.setInterval(32)
        self._layout_timer.timeout.connect(self._update_icon_layout)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        header_row = QtWidgets.QHBoxLayout()
        self.title_label = QtWidgets.QLabel()
        title_font = self.title_label.font()
        title_font.setBold(True)
        self.title_label.setFont(title_font)
        header_row.addWidget(self.title_label)
        header_row.addStretch(1)
        self.summary_label = QtWidgets.QLabel("")
        self.summary_label.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        header_row.addWidget(self.summary_label)
        layout.addLayout(header_row)

        control_row = QtWidgets.QHBoxLayout()
        self.scale_label = QtWidgets.QLabel()
        control_row.addWidget(self.scale_label)
        self.mode_combo = QtWidgets.QComboBox()
        control_row.addWidget(self.mode_combo)
        self.zoom_spin = QtWidgets.QSpinBox()
        self.zoom_spin.setRange(25, 400)
        self.zoom_spin.setSingleStep(5)
        self.zoom_spin.setSuffix("%")
        control_row.addWidget(self.zoom_spin)
        control_row.addStretch(1)
        self.interaction_hint_label = QtWidgets.QLabel("")
        self.interaction_hint_label.setStyleSheet("color: palette(windowText);")
        control_row.addWidget(self.interaction_hint_label)
        self.loading_label = QtWidgets.QLabel("")
        self.loading_label.setStyleSheet("color: palette(highlight);")
        control_row.addWidget(self.loading_label)
        layout.addLayout(control_row)

        self.stack = QtWidgets.QStackedLayout()
        self.empty_label = QtWidgets.QLabel()
        self.empty_label.setAlignment(QtCore.Qt.AlignCenter)
        self.empty_label.setWordWrap(True)
        self.stack.addWidget(self.empty_label)

        self.preview_list = QtWidgets.QListWidget()
        self.preview_list.setViewMode(QtWidgets.QListView.IconMode)
        self.preview_list.setResizeMode(QtWidgets.QListView.Adjust)
        self.preview_list.setMovement(QtWidgets.QListView.Static)
        self.preview_list.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.preview_list.setWordWrap(True)
        self.preview_list.setSpacing(12)
        self.preview_list.setUniformItemSizes(True)
        self.preview_list.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.preview_list.setVerticalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        self.preview_list.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.preview_list.setTextElideMode(QtCore.Qt.ElideMiddle)
        self.stack.addWidget(self.preview_list)
        layout.addLayout(self.stack, 1)

        self._apply_styles()

        self.mode_combo.currentIndexChanged.connect(self._on_mode_changed)
        self.zoom_spin.valueChanged.connect(self._on_zoom_changed)
        self.preview_list.itemDoubleClicked.connect(self._on_item_double_clicked)

        self.retranslate()
        self.set_zoom(preview_zoom_mode, preview_zoom_percent)
        self._show_empty()

    def retranslate(self) -> None:
        self.title_label.setText(t("preview_title"))
        self.scale_label.setText(t("preview_scale_mode"))
        self.interaction_hint_label.setText(t("preview_dblclick_hint"))
        current_mode = self.zoom_mode()
        with QtCore.QSignalBlocker(self.mode_combo):
            self.mode_combo.clear()
            self.mode_combo.addItem(t("preview_scale_fit"), "fit")
            self.mode_combo.addItem(t("preview_scale_custom"), "custom")
            index = self.mode_combo.findData(current_mode)
            self.mode_combo.setCurrentIndex(index if index >= 0 else 0)
        self.empty_label.setText(t("preview_empty"))
        self._update_summary()

    @staticmethod
    def _can_render_preview(file_type: FileType) -> bool:
        return file_type in {FileType.PDF, FileType.WORD, FileType.EXCEL, FileType.PPT}

    def set_zoom(self, mode: str, percent: int) -> None:
        valid_mode = mode if mode in {"fit", "custom"} else "fit"
        try:
            valid_percent = int(percent)
        except (TypeError, ValueError):
            valid_percent = 100
        valid_percent = max(25, min(400, valid_percent))

        with QtCore.QSignalBlocker(self.mode_combo), QtCore.QSignalBlocker(self.zoom_spin):
            index = self.mode_combo.findData(valid_mode)
            self.mode_combo.setCurrentIndex(index if index >= 0 else 0)
            self.zoom_spin.setValue(valid_percent)
        self.zoom_spin.setEnabled(valid_mode == "custom")
        self._schedule_layout_refresh(immediate=True)

    def zoom_mode(self) -> str:
        value = str(self.mode_combo.currentData() or "fit")
        return value if value in {"fit", "custom"} else "fit"

    def zoom_percent(self) -> int:
        return int(self.zoom_spin.value())

    def release_cached_previews(self, clear_shared_blob_cache: bool = False) -> None:
        self._request_token += 1
        self._pending_jobs = None
        self.loading_label.setText("")
        self._cancel_list_worker()

        for worker in list(self._workers):
            if worker.isRunning():
                worker.requestInterruption()
        for worker in list(self._workers):
            if worker.isRunning():
                worker.wait(1200)
        self._workers.clear()
        self._worker = None

        self._cancel_fullscreen_worker()
        if self._fullscreen_dialog is not None:
            self._fullscreen_dialog.close()
            self._fullscreen_dialog = None
        self._fullscreen_job_id = ""

        self._jobs_by_id.clear()
        self._items_by_job.clear()
        self._images_by_job.clear()
        self._image_cache.clear()
        self._image_cache_order.clear()
        self._fullscreen_cache.clear()
        self._fullscreen_cache_order.clear()
        self._scaled_icon_cache.clear()
        self._scaled_icon_order.clear()

        if clear_shared_blob_cache:
            clear_preview_blob_cache()

    def set_jobs(self, jobs: list[PrintJob]) -> None:
        requested_jobs = list(jobs)
        worker_running = self._cancel_list_worker()
        self._selected_total = len(jobs)
        self._request_token += 1
        current_token = self._request_token

        self.preview_list.clear()
        self._jobs_by_id.clear()
        self._items_by_job.clear()
        self._images_by_job.clear()
        self.loading_label.setText("")
        self._update_summary()

        if not requested_jobs:
            self._pending_jobs = None
            self._show_empty()
            return

        self.stack.setCurrentWidget(self.preview_list)
        render_requests: list[PreviewRequest] = []
        shown_jobs = requested_jobs[: self._max_items]

        for job in shown_jobs:
            self._jobs_by_id[job.id] = job
            item = QtWidgets.QListWidgetItem(self._item_text(job))
            item.setData(QtCore.Qt.UserRole, job.id)
            item.setTextAlignment(QtCore.Qt.AlignHCenter)
            item.setToolTip(job.file_path)
            self.preview_list.addItem(item)
            self._items_by_job[job.id] = item

            if not self._can_render_preview(job.file_type):
                continue

            signature = self._file_signature(job.file_path)
            cached_image = self._image_cache.get(signature) if signature else None
            if cached_image is not None and not cached_image.isNull():
                self._images_by_job[job.id] = cached_image
                continue

            render_requests.append(
                PreviewRequest(
                    job_id=job.id,
                    file_path=job.file_path,
                    file_type=job.file_type,
                    signature=signature,
                )
            )

        self._update_summary()
        self._schedule_layout_refresh(immediate=True)

        if not render_requests:
            self._pending_jobs = None
            return

        if worker_running:
            self._pending_jobs = requested_jobs
            self.loading_label.setText(t("preview_loading"))
            return

        self._pending_jobs = None
        self.loading_label.setText(t("preview_loading"))
        worker = PdfPreviewWorker(render_requests, current_token)
        self._workers.add(worker)
        worker.thumbnail_ready.connect(self._on_thumbnail_ready)
        worker.batch_finished.connect(self._on_batch_finished)
        worker.finished.connect(lambda w=worker: self._on_list_worker_finished(w))
        self._worker = worker
        worker.start()

    def resizeEvent(self, event: QtGui.QResizeEvent) -> None:
        super().resizeEvent(event)
        self._schedule_layout_refresh(immediate=False)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        self.release_cached_previews(clear_shared_blob_cache=False)
        super().closeEvent(event)

    def _on_mode_changed(self) -> None:
        mode = self.zoom_mode()
        self.zoom_spin.setEnabled(mode == "custom")
        self.zoom_mode_changed.emit(mode)
        self._schedule_layout_refresh(immediate=False)

    def _on_zoom_changed(self, value: int) -> None:
        self.zoom_percent_changed.emit(value)
        if self.zoom_mode() == "custom":
            self._schedule_layout_refresh(immediate=False)

    def _show_empty(self) -> None:
        self.stack.setCurrentWidget(self.empty_label)
        self.loading_label.setText("")
        self.preview_list.clear()

    def _update_summary(self) -> None:
        shown = self.preview_list.count()
        if self._selected_total <= 0:
            self.summary_label.setText("")
            return
        self.summary_label.setText(t("preview_selection_fmt", selected=self._selected_total, shown=shown))

    def _schedule_layout_refresh(self, immediate: bool) -> None:
        if immediate:
            self._layout_timer.stop()
            self._update_icon_layout()
            return
        self._layout_timer.start()

    def _update_icon_layout(self) -> None:
        count = self.preview_list.count()
        if count <= 0:
            return

        icon_size = self._resolve_icon_size()
        self.preview_list.setIconSize(icon_size)
        item_size = QtCore.QSize(icon_size.width() + 24, icon_size.height() + 66)

        for idx in range(count):
            item = self.preview_list.item(idx)
            item.setSizeHint(item_size)

        for job_id in list(self._items_by_job):
            self._apply_job_icon(job_id)

    def _resolve_icon_size(self) -> QtCore.QSize:
        viewport_width = max(260, self.preview_list.viewport().width())
        if self.zoom_mode() == "fit":
            columns = max(1, min(6, viewport_width // 250))
            spacing = max(10, self.preview_list.spacing())
            usable_width = max(140, viewport_width - spacing * (columns + 1))
            icon_width = max(150, min(360, usable_width // columns))
        else:
            icon_width = max(150, min(520, int(220 * self.zoom_percent() / 100)))
        icon_height = max(190, int(icon_width * 1.34))
        return QtCore.QSize(icon_width, icon_height)

    def _apply_job_icon(self, job_id: str) -> None:
        item = self._items_by_job.get(job_id)
        job = self._jobs_by_id.get(job_id)
        if item is None or job is None:
            return

        size = self.preview_list.iconSize()
        image = self._images_by_job.get(job_id)
        if image is not None and not image.isNull():
            source_key = int(image.cacheKey())
        else:
            source_key = hash((job.file_type.value, job.extension))
        cache_key = (job_id, size.width(), size.height(), source_key)

        icon = self._scaled_icon_cache.get(cache_key)
        if icon is None:
            icon = self._build_card_icon(job, image, size)
            self._remember_scaled_icon(cache_key, icon)
        item.setIcon(icon)

    def _build_card_icon(
        self,
        job: PrintJob,
        image: QtGui.QImage | None,
        size: QtCore.QSize,
    ) -> QtGui.QIcon:
        pixmap = QtGui.QPixmap(size)
        pixmap.fill(QtCore.Qt.transparent)

        painter = QtGui.QPainter(pixmap)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        painter.setRenderHint(QtGui.QPainter.SmoothPixmapTransform)

        palette = self.palette()
        card_rect = QtCore.QRectF(2, 2, max(1, size.width() - 4), max(1, size.height() - 4))
        card_color = palette.color(QtGui.QPalette.Base)
        card_border = palette.color(QtGui.QPalette.Mid)
        view_bg = palette.color(QtGui.QPalette.AlternateBase)

        painter.setPen(QtGui.QPen(card_border, 1))
        painter.setBrush(card_color)
        painter.drawRoundedRect(card_rect, 12, 12)

        viewport_rect = card_rect.adjusted(10, 10, -10, -10).toRect()
        painter.setPen(QtGui.QPen(card_border, 1))
        painter.setBrush(view_bg)
        painter.drawRoundedRect(QtCore.QRectF(viewport_rect), 8, 8)

        if image is not None and not image.isNull():
            source = QtGui.QPixmap.fromImage(image)
            scaled = source.scaled(
                viewport_rect.size(),
                QtCore.Qt.KeepAspectRatio,
                QtCore.Qt.SmoothTransformation,
            )
            target = QtCore.QRect(
                viewport_rect.x() + (viewport_rect.width() - scaled.width()) // 2,
                viewport_rect.y() + (viewport_rect.height() - scaled.height()) // 2,
                scaled.width(),
                scaled.height(),
            )
            painter.setPen(QtGui.QPen(card_border, 1))
            painter.setBrush(QtCore.Qt.NoBrush)
            painter.drawRoundedRect(QtCore.QRectF(target.adjusted(-1, -1, 1, 1)), 6, 6)
            painter.drawPixmap(target, scaled)
        else:
            icon = self._fallback_icon(job.file_path)
            side = max(64, min(120, int(min(viewport_rect.width(), viewport_rect.height()) * 0.48)))
            icon_pix = icon.pixmap(QtCore.QSize(side, side))
            target = QtCore.QRect(
                viewport_rect.x() + (viewport_rect.width() - side) // 2,
                viewport_rect.y() + (viewport_rect.height() - side) // 2,
                side,
                side,
            )
            painter.drawPixmap(target, icon_pix)

        painter.end()
        return QtGui.QIcon(pixmap)

    def _remember_scaled_icon(self, key: tuple[str, int, int, int], icon: QtGui.QIcon) -> None:
        if key in self._scaled_icon_cache:
            return
        self._scaled_icon_cache[key] = icon
        self._scaled_icon_order.append(key)
        if len(self._scaled_icon_order) <= _SCALED_ICON_CACHE_LIMIT:
            return
        oldest = self._scaled_icon_order.pop(0)
        self._scaled_icon_cache.pop(oldest, None)

    def _fallback_icon(self, file_path: str) -> QtGui.QIcon:
        icon = self._icon_provider.icon(QtCore.QFileInfo(file_path))
        if icon and not icon.isNull():
            return icon
        return self.style().standardIcon(QtWidgets.QStyle.SP_FileIcon)

    def _item_text(self, job: PrintJob) -> str:
        type_label = self._type_label(job.file_type)
        if job.page_count is not None:
            detail = f"{type_label} / {job.page_count}"
        else:
            detail = type_label
        return f"{job.file_name}\n{detail}"

    @staticmethod
    def _type_label(file_type: FileType) -> str:
        if file_type == FileType.PDF:
            return t("label_pdf")
        if file_type == FileType.WORD:
            return t("label_word")
        if file_type == FileType.EXCEL:
            return t("label_excel")
        if file_type == FileType.PPT:
            return t("label_ppt")
        return t("label_unknown")

    @staticmethod
    def _file_signature(file_path: str) -> tuple[str, int, int] | None:
        try:
            path = Path(file_path)
            stat = path.stat()
            return (str(path), int(stat.st_mtime_ns), int(stat.st_size))
        except Exception:
            return None

    def _remember_image_cache(self, signature: tuple[str, int, int], image: QtGui.QImage) -> None:
        if signature in self._image_cache:
            return
        self._image_cache[signature] = image
        self._image_cache_order.append(signature)
        if len(self._image_cache_order) <= _IMAGE_CACHE_LIMIT:
            return
        oldest = self._image_cache_order.pop(0)
        self._image_cache.pop(oldest, None)

    def _remember_fullscreen_cache(self, signature: tuple[str, int, int], image: QtGui.QImage) -> None:
        if signature in self._fullscreen_cache:
            return
        self._fullscreen_cache[signature] = image
        self._fullscreen_cache_order.append(signature)
        if len(self._fullscreen_cache_order) <= _FULLSCREEN_CACHE_LIMIT:
            return
        oldest = self._fullscreen_cache_order.pop(0)
        self._fullscreen_cache.pop(oldest, None)

    def _on_thumbnail_ready(self, job_id: str, signature, image, token: int) -> None:
        if token != self._request_token:
            return
        if not isinstance(image, QtGui.QImage) or image.isNull():
            return
        self._images_by_job[job_id] = image
        if isinstance(signature, tuple):
            self._remember_image_cache(signature, image)
        self._apply_job_icon(job_id)

    def _on_batch_finished(self, token: int) -> None:
        if token != self._request_token:
            return
        self.loading_label.setText("")
        if self._worker and not self._worker.isRunning():
            self._worker = None

    def _cancel_list_worker(self) -> bool:
        if self._worker is None:
            return False
        if self._worker.isRunning():
            self._worker.requestInterruption()
            return True
        self._worker = None
        return False

    def _on_list_worker_finished(self, worker: PdfPreviewWorker) -> None:
        self._workers.discard(worker)
        if self._worker is worker:
            self._worker = None
        worker.deleteLater()
        if self._worker is not None:
            return
        pending_jobs = self._pending_jobs
        if pending_jobs is None:
            return
        self._pending_jobs = None
        self.set_jobs(list(pending_jobs))

    def _on_item_double_clicked(self, item: QtWidgets.QListWidgetItem) -> None:
        data = item.data(QtCore.Qt.UserRole)
        job_id = str(data) if data else ""
        if not job_id:
            return
        self._open_fullscreen_preview(job_id)

    def _open_fullscreen_preview(self, job_id: str) -> None:
        job = self._jobs_by_id.get(job_id)
        if job is None:
            return

        if self._fullscreen_dialog is not None:
            self._fullscreen_dialog.close()

        dialog = FullscreenPreviewDialog(self.window())
        dialog.set_title(job.file_name, job.file_path)
        dialog.set_hint(t("preview_fullscreen_click_close"))
        dialog.finished.connect(self._on_fullscreen_closed)
        self._fullscreen_dialog = dialog
        self._fullscreen_job_id = job_id

        image = self._images_by_job.get(job_id)
        if image is not None and not image.isNull():
            dialog.set_image(image)
        else:
            dialog.set_icon(self._fallback_icon(job.file_path))

        if self._can_render_preview(job.file_type):
            dialog.set_status(t("preview_fullscreen_loading"))
            self._start_fullscreen_render(job)
        else:
            dialog.set_status("")

        dialog.showFullScreen()
        dialog.raise_()
        dialog.activateWindow()

    def _start_fullscreen_render(self, job: PrintJob) -> None:
        signature = self._file_signature(job.file_path)
        if signature is not None:
            cached = self._fullscreen_cache.get(signature)
            if cached is not None and not cached.isNull():
                if self._fullscreen_dialog and self._fullscreen_job_id == job.id:
                    self._fullscreen_dialog.set_image(cached)
                    self._fullscreen_dialog.set_status("")
                return

        if not self._cancel_fullscreen_worker():
            return
        edge = self._fullscreen_target_edge()
        worker = SinglePdfPreviewWorker(
            job_id=job.id,
            file_path=job.file_path,
            file_type=job.file_type,
            signature=signature,
            target_long_edge=edge,
            parent=self,
        )
        worker.image_ready.connect(self._on_fullscreen_image_ready)
        worker.finished.connect(lambda w=worker: self._on_fullscreen_worker_finished(w))
        self._fullscreen_worker = worker
        worker.start()

    def _fullscreen_target_edge(self) -> int:
        if self._fullscreen_dialog:
            screen = self._fullscreen_dialog.screen()
        else:
            screen = QtGui.QGuiApplication.primaryScreen()
        if screen is None:
            return 1600
        size = screen.size()
        return max(1400, int(max(size.width(), size.height()) * 1.15))

    def _on_fullscreen_image_ready(self, job_id: str, signature, image) -> None:
        if not isinstance(image, QtGui.QImage) or image.isNull():
            if self._fullscreen_dialog and self._fullscreen_job_id == job_id:
                self._fullscreen_dialog.set_status("")
            return

        if isinstance(signature, tuple):
            self._remember_fullscreen_cache(signature, image)

        if self._fullscreen_dialog and self._fullscreen_job_id == job_id:
            self._fullscreen_dialog.set_image(image)
            self._fullscreen_dialog.set_status("")

    def _cancel_fullscreen_worker(self) -> bool:
        if self._fullscreen_worker is None:
            return True
        worker = self._fullscreen_worker
        self._fullscreen_worker = None
        if worker.isRunning():
            worker.requestInterruption()
            if not worker.wait(2000):
                # Keep ownership until it actually finishes to avoid deleting a running QThread.
                self._fullscreen_worker = worker
                return False
        try:
            worker.finished.disconnect()
        except Exception:
            pass
        try:
            worker.image_ready.disconnect(self._on_fullscreen_image_ready)
        except Exception:
            pass
        worker.deleteLater()
        return True

    def _on_fullscreen_worker_finished(self, worker: SinglePdfPreviewWorker) -> None:
        if self._fullscreen_worker is worker:
            self._fullscreen_worker = None
        worker.deleteLater()

    def _on_fullscreen_closed(self, _result: int) -> None:
        self._fullscreen_dialog = None
        self._fullscreen_job_id = ""
        self._cancel_fullscreen_worker()

    def _apply_styles(self) -> None:
        self.setObjectName("previewPanel")
        self.empty_label.setObjectName("previewEmpty")
        self.preview_list.setObjectName("previewList")
        self.setStyleSheet(
            """
            QWidget#previewPanel {
                border: 1px solid palette(mid);
                border-radius: 10px;
                background: palette(base);
            }
            QLabel#previewEmpty {
                padding: 20px;
            }
            QListWidget#previewList {
                border: none;
                background: transparent;
                padding: 4px;
            }
            QListWidget#previewList::item {
                border: 1px solid palette(mid);
                border-radius: 10px;
                padding: 8px;
                margin: 2px;
                background: palette(base);
            }
            QListWidget#previewList::item:hover {
                border-color: palette(highlight);
                background: palette(alternate-base);
            }
            """
        )
