from __future__ import annotations

import logging
import os
import sys
from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets

try:
    import app
    print("app module imported successfully.")
except ModuleNotFoundError as e:
    print("ModuleNotFoundError:", e)

from app.app_context import AppContext
from app.controller.job_manager import JobManager
from app.ui.main_window import MainWindow
from app.ui.theme import apply_theme
from app.ui.icon_data import ICON_PNG_BASE64
from app.backend import pdf_worker
from app.i18n import set_language, resolve_language
from app.updater import apply_update


# Ensure the project root is in sys.path
project_root = Path(__file__).resolve().parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

print("sys.path:", sys.path)
print("Current working directory:", os.getcwd())


def configure_logging(log_path: str) -> None:
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)


def _apply_window_icon(app: QtWidgets.QApplication) -> None:
    base = Path(__file__).resolve().parent.parent
    ico = base / "icon.ico"
    png = base / "icon.png"
    icon_path = ico if ico.exists() else png if png.exists() else None
    if icon_path:
        app.setWindowIcon(QtGui.QIcon(str(icon_path)))
        return
    try:
        data = QtCore.QByteArray.fromBase64(ICON_PNG_BASE64.encode("ascii"))
        image = QtGui.QImage.fromData(data, "PNG")
        if not image.isNull():
            app.setWindowIcon(QtGui.QIcon(QtGui.QPixmap.fromImage(image)))
    except Exception:
        return


def main() -> int:
    if "--pdf-worker" in sys.argv:
        return pdf_worker.main()

    if "--apply-update" in sys.argv:
        idx = sys.argv.index("--apply-update")
        return apply_update(sys.argv[idx + 1 :])

    app = QtWidgets.QApplication(sys.argv)
    context = AppContext()
    set_language(resolve_language(context.settings.language_mode))
    _apply_window_icon(app)
    apply_theme(app, context.settings.theme_mode)
    configure_logging(str(context.log_path))

    job_manager = JobManager(context)
    window = MainWindow(context, job_manager)
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
