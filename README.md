# RakuPrints (らーく印刷)

<img src="assets/icon.png" width="80">

A Windows app for printing a lot of files at once.

A folder with PDFs, Excel sheets, Word documents and PowerPoint decks mixed together
does not have to be opened one file at a time. Drag the whole lot in and press print.

It **switches printer and duplex settings automatically per file type**, and works out
page orientation for Excel sheets by looking at the content. A preview step before
anything reaches paper keeps a wrong setting from becoming a hundred wasted pages.

> **Impact —** Built to replace a manual Adobe Acrobat workflow at a vehicle inspection
> shop: ~170 documents a day taking 50+ minutes, now about 5 minutes. In daily
> production use.

日本語版: [README.ja.md](README.ja.md)

## What it does

- Batch printing across PDF, Excel, Word and PowerPoint
- Automatic printer selection and per-file-type settings, including duplex
- Automatic page-orientation detection for Excel sheets
- Pre-print preview, to catch a wrong setting before paper is used
- Progress and result display, with error logging
- Drag and drop, Japanese/English UI, dark and light themes

## Using it

Download `RakuPrint_Setup_x.x.x.exe` from
[Releases](https://github.com/sukun-inu/RakuPrints/releases) and run it.

1. Start the app
2. Drag in the files you want printed
3. Adjust printer or duplex settings if you need to
4. Press print

**Microsoft Office is required to print Excel, Word and PowerPoint files.** PDF-only use
does not need it.

## Documentation

| | |
|---|---|
| [docs/USAGE.ja.md](docs/USAGE.ja.md) | Worked examples and FAQ |
| [docs/ARCHITECTURE.ja.md](docs/ARCHITECTURE.ja.md) | Feature overview, architecture, data flow |
| [docs/DEVELOPMENT.ja.md](docs/DEVELOPMENT.ja.md) | Building and running, config files, contributing |

Detailed documentation is Japanese-only for now.

## Licence

MIT. Copyright (c) 2026 Hibiki Suzuki. See [LICENSE](LICENSE).
