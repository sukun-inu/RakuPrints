# RakuPrints (らーく印刷)

A batch printing application for Windows. Print PDF, Excel, Word and PowerPoint files in
one go, with the printer and settings chosen automatically per file type.

> **Impact —** Built to replace a manual Adobe Acrobat workflow at a vehicle inspection
> shop: ~170 documents a day taking 50+ minutes, now about 5 minutes. In daily
> production use.

日本語版: [README.ja.md](README.ja.md)

**Microsoft Office is required to print Excel, Word and PowerPoint files.** PDF printing
works without it.

## Features

- Batch printing across PDF / Excel / Word / PowerPoint
- Automatic printer selection and per-file-type settings, including duplex
- Automatic page-orientation detection for Excel sheets, to avoid wasted pages
- Print job management: progress display, error handling, log output
- Pre-print preview, to catch misprints and wrong settings before paper is used
- Drag and drop
- Multilingual UI (Japanese / English) and theme switching

## Install

Download `RakuPrint_Setup_x.x.x.exe` from
[Releases](https://github.com/sukun-inu/RakuPrints/releases) and run it.

## Documentation

| | |
|---|---|
| [docs/USAGE.ja.md](docs/USAGE.ja.md) | Worked examples and FAQ |
| [docs/ARCHITECTURE.ja.md](docs/ARCHITECTURE.ja.md) | Feature overview, architecture, data flow |
| [docs/DEVELOPMENT.ja.md](docs/DEVELOPMENT.ja.md) | Building and running, config files, contributing |

Detailed documentation is Japanese-only for now.

## Licence

MIT.
