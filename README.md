# UF検査判定(Web 版)

Excel VBA の「UF検査判定システム」(`UF検査判定.xlsm`、梱包課のライン検査)を、
python-web-tools の基盤に載せた Web システムへ移植しています。

**現在の状態(2026-10-30)**: 解析・移植計画・判断事項の回答まで完了。実装中です。

| できていること | 内容 |
|---|---|
| 基盤 | 起動・停止・多重起動防止・自動終了・ログ(python-web-tools / db-viewer から複製。`FOUNDATION_ORIGIN.md`) |
| 設定画面 | この端末(設備名・号機)・モード(オンライン/オフライン)・**使用パス**(確認ボタン)・権限・不具合の扱い |
| 取り込み画面 | Access(.accdb)→ SQLite(.sqlite3)。cscript → ADODB → ACE(追加ライブラリ不要) |
| 業務ロジックの土台 | VBA 互換(Single・Val・CStr・Format・Round)、ADO Filter と同じ絞り込み、不具合スイッチ(B1〜B38) |
| 判定ロジック | 移植中(判定・公差外・品質基準照合・ロット・入力先・Aインプット・保存・帳票・特殊検査) |
| 検査画面ほか | これから(A案「計測コックピット」) |

## 起動

ほかの業務ツールと同じく、業務ツール統合ランチャーに `Start.vbs` を登録して起動します(ポート 8800)。
必要なもの: Python 3.12、`pip install -r requirements.txt`(Flask・waitress。判断事項 D1)。
Access からの取り込みには Microsoft Access Database Engine(ACE OLE DB 12.0)が要ります。

初めて開いたら **設定 → この端末** で設備名(ライン)を決めます(VBA の初期設定 SF3 と同じ)。

## 試験

    python -m pytest -q

`tests/fixtures/kensa/` は受領した Access から作った基準データです(名簿は氏名・カード番号を伏せた。仕掛は108ロット分)。
作り直すときは `scripts/dev_accdb_to_tsv.py`(開発用。mdbtools が要る)。

## ドキュメント

| ファイル | 内容 |
|---|---|
| [docs/01_VBA解析.md](docs/01_VBA解析.md) | VBA が何をしているか、モジュール構成、業務フロー、判定ロジック、外部依存 |
| [docs/02_python-web-tools調査.md](docs/02_python-web-tools調査.md) | 既存基盤と再利用する部品 |
| [docs/03_移植計画.md](docs/03_移植計画.md) | 方針、機能対応表、Access→SQLite・複数PC・長時間放置への方針、使用パス、工程、検証計画 |
| [docs/04_不明点・判断事項.md](docs/04_不明点・判断事項.md) | **決定事項(D1〜D13)**、実機確認、業務確認、不足資料、既存不具合 B1〜B38 |
| [docs/05_AccessDB調査.md](docs/05_AccessDB調査.md) | 受領した Access 5 ファイルの表・列と VBA の参照の突き合わせ |
| [docs/analysis/](docs/analysis/) | モジュール群ごとの詳細な解析レポート(行番号つき) |
| [docs/foundation/](docs/foundation/) | 基盤の規約(複製元) |
| [vba_src/](vba_src/) | VBA の原文(93 モジュール) |

## 構成

    app/                 画面(Flask。ルート・テンプレート・JS・CSS)
    ufkensa/             基盤(起動・ログ・権限・設定・取り込み)
    ufkensa/kensa/       業務ロジック(VBA の移植。Python だけ)
    tests/               試験
