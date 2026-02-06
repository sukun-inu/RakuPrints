# らーく印刷 (RakuPrint)

<img src="assets/icon.png" width="80">

Windows向けの一括印刷アプリケーション。PDF、Excel、Word、PowerPointファイルをまとめて印刷できます。

## 📥 インストール

### ユーザー向け
[Releases](https://github.com/sukun-inu/RakuPrints/releases) から `RakuPrint_Setup_x.x.x.exe` をダウンロードして実行してください。

### 開発者向け
```powershell
# リポジトリをクローン
git clone https://github.com/sukun-inu/RakuPrints.git
cd RakuPrints

# 仮想環境を作成・有効化
python -m venv .venv
.venv\Scripts\Activate.ps1

# 依存関係をインストール
pip install -r requirements.txt

# アプリを起動
python -m app.main
```

## ✨ 機能

- **対応形式**: PDF, Excel (.xlsx/.xls), Word (.docx/.doc), PowerPoint (.pptx/.ppt)
- **ドラッグ&ドロップ**: ファイルを画面にドロップするだけで追加
- **プリンター設定**: ファイル形式ごとにプリンターを指定可能
- **多言語対応**: 日本語、英語、韓国語、中国語
- **テーマ**: ライト/ダーク/システム連動
- **自動アップデート**: 新バージョンを自動検出

## 📁 プロジェクト構成

```
RakuPrints/
├── app/                    # アプリケーション本体
│   ├── backend/            # 印刷処理 (PDF, Excel, Word, PPT)
│   ├── controller/         # ビジネスロジック
│   ├── model/              # データモデル
│   ├── ui/                 # ユーザーインターフェース (PySide6)
│   ├── config/             # 設定ファイル (自動生成)
│   ├── logging/            # ログファイル (自動生成)
│   ├── app_context.py      # アプリ設定管理
│   ├── i18n.py             # 多言語対応
│   ├── main.py             # エントリーポイント
│   └── version.py          # バージョン情報
├── scripts/                # ビルドスクリプト
│   └── build.ps1           # ビルド自動化
├── assets/                 # アイコン等のリソース
├── RakuPrint.spec          # PyInstaller設定
├── installer.iss           # Inno Setupスクリプト
├── requirements.txt        # Python依存関係
└── README.md               # このファイル
```

## 🔨 ビルド

### EXEのみ作成
```powershell
python -m PyInstaller RakuPrint.spec --noconfirm
```
出力: `dist/RakuPrint/RakuPrint.exe`

### インストーラー作成 (Inno Setup 6 必要)
```powershell
.\scripts\build.ps1
```
出力:
- `dist/RakuPrint/` - スタンドアロン版
- `dist/RakuPrint_vX.X.X.zip` - GitHub Release用
- `installer_output/RakuPrint_Setup_X.X.X.exe` - インストーラー

## 🔄 自動アップデート

GitHub Releasesに `.zip` ファイルを添付すると、アプリが自動的に検出してアップデートを提案します。

## 📝 ライセンス

Copyright © 2026 Hibiki Suzuki. All rights reserved.
