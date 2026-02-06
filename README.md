
# らーく印刷 (RakuPrint)

<img src="assets/icon.png" width="80">

**※ Excel・Word・PowerPointファイルの印刷には Microsoft Office が必須です（PDF印刷は不要）**

Windows向けの一括印刷アプリケーション。PDF、Excel、Word、PowerPointファイルをまとめて印刷できます。

## 📥 インストール

### ユーザー向け
[Releases](https://github.com/sukun-inu/RakuPrints/releases) から `RakuPrint_Setup_x.x.x.exe` をダウンロードして実行してください。


## らーく印刷 (RakuPrints)

### こんな方におすすめ
- 仕事や学校で大量のPDFやOfficeファイルをまとめて印刷したい
- ファイルごとにプリンターや両面印刷設定を自動で切り替えたい
- Excelの用紙向き（縦横）を自動判定して無駄な印刷を減らしたい
- ドラッグ＆ドロップで簡単に操作したい
- 日本語・英語など多言語UIで使いたい

### 概要
「らーく印刷」は、PDF・Excel・Word・PowerPointなど様々なファイルを一括で印刷できるWindowsアプリです。ファイルを画面にドラッグ＆ドロップするだけで、印刷ジョブが自動で作成されます。拡張子ごとにプリンターや両面印刷などのルールを設定でき、Excelは自動で用紙向きを判定。印刷進捗や結果も画面で確認できます。

### 主なメリット
- 手作業で1つずつ開いて印刷する必要なし
- プリンターや設定の自動切り替えでミス防止
- 印刷状況やエラーも一目でわかる
- ダーク/ライトテーマや多言語対応

### 基本操作
1. 🚀 アプリを起動
2. 📂 印刷したいファイルをドラッグ＆ドロップ
3. 🖨️ 必要に応じてプリンターや両面印刷など設定
4. ▶️ 「印刷開始」ボタンを押すだけ

---

### 操作例

#### 例1：PDFとExcelをまとめて印刷
1. アプリを起動
2. PDFとExcelファイルをまとめてドラッグ＆ドロップ
3. プリンターや両面印刷設定を確認
4. 「印刷開始」をクリック
→ ファイルごとに自動でプリンターや設定が切り替わり、順番に印刷されます

#### 例2：Excelの用紙向きを自動判定
1. Excelファイルを追加
2. 「自動判定」機能で縦横を自動選択
3. 印刷設定を確認し「印刷開始」

#### 例3：拡張子ごとにプリンターを指定
1. 設定画面で「.pdfはAプリンター」「.xlsxはBプリンター」などルールを登録
2. ファイルを追加し「印刷開始」
→ ルールに従って自動でプリンターが切り替わります

---

### ❓ FAQ（よくある質問）

**Q. 印刷ジョブの順番を変えたい**
→ ファイルリストでドラッグして並び替えできます

**Q. 印刷エラーが出た場合は？**
→ エラー内容が画面に表示されます。ログ（app/logging/app.log）も確認できます

**Q. 設定やルールは保存される？**
→ アプリ終了時に自動保存され、次回起動時に復元されます

**Q. 多言語やテーマはどこで切り替える？**
→ 設定パネルからいつでも変更可能です

**Q. Officeがインストールされていないと使えますか？**
→ Excel/Word/PPT印刷はMicrosoft Officeが必要です。PDF印刷は不要です

---

### 機能概要
- 対応ファイル形式：PDF, Excel, Word, PowerPoint
- バッチ印刷：複数ファイルを一括で印刷
- 印刷設定：プリンター選択、部数、両面印刷、用紙向き
- 拡張子ごとのプリンター自動選択ルール
- Excelシートの自動用紙向き判定
- 印刷ジョブ管理（追加・削除・並び替え・状態表示）
- 印刷進捗・結果のUI表示
- ログ記録（app/logging/app.log）
- 多言語対応（日本語・英語・韓国語・中国語）
- ダーク/ライトテーマ切替
- 設定保存（app/config/user_settings.json, rules.json）

---

### アーキテクチャ

```
app/
	main.py           # エントリーポイント
	app_context.py    # 設定・ルール管理
	i18n.py           # 多言語対応
	backend/          # 印刷処理（PDF, Office, プリンター管理）
	controller/       # ジョブ管理・実行・ルール判定
	model/            # ジョブデータ構造
	ui/               # UI各種（メイン画面、設定、ダイアログ等）
	config/           # 設定・ルールファイル
	logging/          # ログ出力
```

---

### データフロー
1. ファイル追加 → PrintJob生成・ジョブリスト更新
2. 「印刷開始」クリック → JobExecutorスレッド起動
3. ジョブごとにBackendで印刷処理
4. 成否・進捗をUIへシグナル通知
5. 設定・ルール変更時はAppContextが永続化

---

### ビルド・実行方法

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

### 設定ファイル
- app/config/user_settings.json：ユーザー設定（プリンター、部数、両面、テーマ、言語）
- app/config/rules.json：拡張子ごとのプリンター自動選択ルール
- app/logging/app.log：アプリログ

---

### ライセンス・著作権
（必要に応じて記載）

---

### 開発・貢献
Pull Request歓迎。Issueで要望・バグ報告も受付。

---

### 作者
Kawasaki
