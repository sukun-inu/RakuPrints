# 開発・ビルド

[← README に戻る](../README.md)

---

## ビルド・実行方法

```powershell
# 仮想環境作成・有効化
python -m venv .venv
.venv\Scripts\Activate.ps1

# 依存パッケージインストール
pip install -r requirements.txt

# アプリ起動
python -m app.main

# PyInstallerでexe化
python -m PyInstaller RakuPrint.spec --noconfirm

# インストーラー作成（Inno Setup 6必要）
.\scripts\build.ps1
```

---

## 設定ファイル
- app/config/user_settings.json：ユーザー設定（プリンター、部数、両面、テーマ、言語）
- app/config/rules.json：拡張子ごとのプリンター自動選択ルール
- app/logging/app.log：アプリログ

---

---

## 開発・貢献
Pull Request歓迎。Issueで要望・バグ報告も受付。
