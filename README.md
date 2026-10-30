# inspection-data-gateway（UF検査判定システムの Python Web 移植）

Excel VBA の「UF検査判定システム」（`UF検査判定.xlsm`、梱包課のライン検査）を、
python-web-tools の基盤に載せた Web システムへ移植するためのリポジトリです。

**現在の状態: 解析と移植計画まで（実装前）。** 判断事項のご回答をいただいてから実装に入ります。

## ドキュメント

| ファイル | 内容 |
|---|---|
| [docs/01_VBA解析.md](docs/01_VBA解析.md) | このVBAは何をしているか、モジュール構成、業務フロー、判定ロジック、外部依存、出力、エラー処理 |
| [docs/02_python-web-tools調査.md](docs/02_python-web-tools調査.md) | 既存基盤の構成と、再利用できる部品（db-viewer の移植部品を含む） |
| [docs/03_移植計画.md](docs/03_移植計画.md) | 方針、VBA→Web の機能対応表、Access→SQLite・複数PC・長時間放置への方針、使用パスの設定、工程、検証計画 |
| [docs/04_不明点・判断事項.md](docs/04_不明点・判断事項.md) | 判断事項、実機確認、業務確認、不足資料、既存の不具合一覧 |
| [docs/05_AccessDB調査.md](docs/05_AccessDB調査.md) | 受領した Access 5 ファイルの表・列と VBA の参照の突き合わせ |
| [docs/analysis/](docs/analysis/) | モジュール群ごとの詳細な解析レポート（行番号つき） |
| [docs/db_schema/](docs/db_schema/) | Access のスキーマ（表・列・型のみ。データは含めない） |
| [vba_src/](vba_src/) | 抽出した VBA の原文（93 モジュール。コメント中の氏名は伏せ字） |

## 前提

- 対象はPCのみ（ラインPC、Windows）。
- ラインPCの Python には追加ライブラリを入れられない（標準ライブラリ: sqlite3 / tkinter / csv / logging / subprocess）。
- Access(.accdb) へは `cscript.exe → VBScript → ADODB → ACE OLE DB` で接続。将来 SQLite に切り替える。
