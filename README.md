# らーく印刷

## 概要
らーく印刷は、Excel、Word、PDF、PowerPointなどのドキュメントを効率的に印刷するためのアプリケーションです。ユーザーは、印刷ジョブを簡単に管理し、設定をカスタマイズすることができます。

## 主な機能
- **Excel印刷**: Excelシートの選択や向きの分析をサポート。
- **Word印刷**: Wordドキュメントの印刷設定を簡単に管理。
- **PDF印刷**: PDFファイルの印刷ジョブを効率的に処理。
- **PowerPoint印刷**: スライドの印刷設定をサポート。
- **プリンター選択**: 使用可能なプリンターを簡単に選択可能。
- **ログ管理**: 印刷ジョブのログを確認可能。

## ディレクトリ構成
```
app/
  backend/       # 各種ドキュメントの印刷処理を担当
  config/        # 設定ファイル (ルールやユーザー設定)
  controller/    # 印刷ジョブの管理と実行
  model/         # データモデル (印刷ジョブや結果)
  ui/            # ユーザーインターフェース
build/           # ビルド成果物
```

## 必要条件
- Python 3.9以上
- 必要なPythonライブラリは`requirements.txt`に記載されています。

## セットアップ
1. 仮想環境を作成し、アクティベートします。
   ```powershell
   python -m venv .venv
   & .venv\Scripts\Activate.ps1
   ```
2. 必要なライブラリをインストールします。
   ```powershell
   pip install -r requirements.txt
   ```

## 実行方法
1. アプリケーションを起動します。
   ```powershell
   python app/main.py
   ```

## ビルド方法
PyInstallerを使用してアプリケーションをビルドします。
```powershell
pyinstaller らーく印刷.spec
```

## ライセンス
このプロジェクトのライセンスについては、別途記載されている場合を除き、すべての権利を留保します。