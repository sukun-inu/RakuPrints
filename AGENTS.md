# AGENTS.md

This file provides guidance to WARP (warp.dev) when working with code in this repository.

## Project Overview

**らーく印刷 (RakuPrints)** is a Windows desktop application for batch printing multiple document formats (PDF, Excel, Word, PowerPoint). Built with Python and PySide6 (Qt), it uses pywin32/COM automation for Office documents and PyMuPDF for PDFs.

## Build & Run Commands

```powershell
# Create and activate virtual environment
python -m venv .venv
.venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt

# Run the application
python -m app.main

# Build executable with PyInstaller
python -m PyInstaller RakuPrint.spec --noconfirm

# Build with installer (requires Inno Setup 6)
.\scripts\build.ps1
```

The built executable outputs to `dist/RakuPrint/`.

## Architecture

### Entry Points
- `app/main.py` - Main entry point. Also handles `--pdf-worker` subprocess mode for PDF printing and `--apply-update` for self-updates.

### Core Layers

**Model (`app/model/`)**
- `print_job.py` - `PrintJob` dataclass representing a print task with status, printer settings, and file metadata. Uses `FileType`, `JobStatus`, `DuplexMode` enums.

**Controller (`app/controller/`)**
- `job_manager.py` - `JobManager` manages the job queue (add/remove/reorder), applies printer rules, and emits Qt signals on changes.
- `job_executor.py` - `JobExecutor` (QThread) iterates through enabled jobs and dispatches to appropriate backend. Emits progress/status signals.
- `rules_engine.py` - Resolves which printer to use based on file extension rules defined in `app/config/rules.json`.
- `excel_orientation_analyzer.py` - Background analyzer that checks Excel sheet dimensions to recommend portrait/landscape orientation.

**Backend (`app/backend/`)**
- `pdf_backend.py` - Spawns a subprocess (`--pdf-worker`) to render PDF pages via PyMuPDF and send to printer via GDI.
- `excel_backend.py` - Uses COM automation (`win32com.client`) to open Excel and call `PrintOut()`. Handles sheet selection and auto-orientation.
- `word_backend.py` - COM automation for Word documents.
- `ppt_backend.py` - COM automation for PowerPoint files.
- `printer_utils.py` - Windows printer enumeration and management via `win32print`.

**UI (`app/ui/`)**
- `main_window.py` - Main application window with file list, settings panel, and menu bar.
- `file_list_view.py` - Table view showing print jobs with drag-drop support.
- `settings_panel.py` - Right panel for printer settings, rules, theme, and language.
- `theme.py` - Dark/light theme application using Qt stylesheets.

**Application State (`app/app_context.py`)**
- `AppContext` - Singleton holding `UserSettings` and extension-to-printer rules. Persists to `app/config/user_settings.json` and `app/config/rules.json`. Emits signals when settings change.

**Internationalization (`app/i18n.py`)**
- `t(key, **kwargs)` function for translations. Supports Japanese (default), English, Korean, and Chinese. Language detection from Windows UI language or locale.

### Data Flow
1. User adds files → `JobManager.add_files()` creates `PrintJob` objects
2. User clicks "Start Printing" → `MainWindow` creates `JobExecutor` thread
3. `JobExecutor.run()` iterates jobs, calls `backend.print(job)` for each
4. Backend uses COM/subprocess to print, signals success/failure
5. `MainWindow` updates UI via Qt signals

### Key Patterns
- **Qt Signals**: Cross-thread and component communication via `jobs_changed`, `job_updated`, `settings_changed`, etc.
- **COM Initialization**: Office backends call `pythoncom.CoInitialize()` before COM calls and clean up with `CoUninitialize()`.
- **Subprocess for PDF**: PDF printing runs in a separate process to isolate PyMuPDF memory usage.

## Configuration Files

- `app/config/user_settings.json` - User preferences (printer, copies, duplex, theme, language)
- `app/config/rules.json` - Per-extension printer rules (e.g., `.pdf` → specific printer)
- `app/logging/app.log` - Application log file

## Language Support

All UI strings use `t("key")` from `app/i18n.py`. When adding new UI text:
1. Add the key to all language dictionaries in `TRANSLATIONS`
2. Use `t("key")` or `t("key", param=value)` for interpolation
