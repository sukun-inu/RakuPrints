# 機能と内部構成

[← README.ja.md に戻る](../README.ja.md)

---

## 機能概要
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

## アーキテクチャ

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

## データフロー
1. ファイル追加 → PrintJob生成・ジョブリスト更新
2. 「印刷開始」クリック → JobExecutorスレッド起動
3. ジョブごとにBackendで印刷処理
4. 成否・進捗をUIへシグナル通知
5. 設定・ルール変更時はAppContextが永続化

---
