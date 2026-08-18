# らーく印刷 (RakuPrints)

<img src="assets/icon.png" width="80">

たくさんのファイルをまとめて印刷するための Windows アプリです。

PDF・Excel・Word・PowerPoint が混ざっていても、1つずつ開いて印刷する必要はありません。
ファイルをまとめてドラッグ＆ドロップすれば、あとは印刷開始を押すだけです。

ファイルの種類ごとに**使うプリンターや両面印刷の設定を自動で切り替え**、Excel は
用紙の向き（縦・横）も中身を見て判断します。印刷する前に仕上がりを確認できるので、
設定を間違えたまま大量に刷ってしまう事故も防げます。

> **導入実績 —** 自動車整備店の車検証等の印刷業務に導入。Adobe Acrobat による手作業
> （約170枚/日・50分以上）を**約5分**に短縮し、日次運用中。

English: [README.md](README.md)

## できること

- PDF・Excel・Word・PowerPoint の一括印刷
- ファイル種別ごとのプリンター自動切替、両面印刷などの設定
- Excel の用紙向き（縦・横）の自動判定
- 印刷前プレビューによる刷り間違いの防止
- 印刷の進捗・結果の表示、エラーのログ出力
- ドラッグ＆ドロップ操作、日本語／英語UI、ダーク／ライトテーマ

## 使うには

[Releases](https://github.com/sukun-inu/RakuPrints/releases) から
`RakuPrint_Setup_x.x.x.exe` をダウンロードして実行してください。

1. アプリを起動する
2. 印刷したいファイルをドラッグ＆ドロップする
3. 必要ならプリンターや両面印刷を設定する
4. 「印刷開始」を押す

**Excel・Word・PowerPoint の印刷には Microsoft Office が必要です。**
PDF だけなら Office は要りません。

## ドキュメント

| | |
|---|---|
| [docs/USAGE.ja.md](docs/USAGE.ja.md) | 操作例、FAQ |
| [docs/ARCHITECTURE.ja.md](docs/ARCHITECTURE.ja.md) | 機能概要、アーキテクチャ、データフロー |
| [docs/DEVELOPMENT.ja.md](docs/DEVELOPMENT.ja.md) | ビルド・実行方法、設定ファイル、開発・貢献 |

## ライセンス

MIT License. Copyright (c) 2026 Hibiki Suzuki
詳しくは [LICENSE](LICENSE) を参照してください。
