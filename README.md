# UF検査判定(Web 版)

Excel VBA の「UF検査判定システム」(`UF検査判定.xlsm`、梱包課のライン検査)を、
python-web-tools の基盤に載せた Web システムへ移植しています。

**現在の状態(2026-11-02)**: 業務ロジックの移植と、全画面(検査・キャスク・全ヵ所/H494・Aインプット・校正・品質基準・履歴)の結合が済みました。実機での確認(U1〜U12)と業務確認(W1〜W15)の回答待ちです。

| できていること | 内容 |
|---|---|
| 基盤 | 起動・停止・多重起動防止・自動終了(python-web-tools / db-viewer から複製。`FOUNDATION_ORIGIN.md`) |
| 設定画面 | この端末(設備名・号機)・モード・**使用パス**(確認ボタン)・権限・ログ・不具合の扱い |
| 取り込み画面 | Access(.accdb)→ SQLite、各ラインの xlsm から1回だけ取り込み(D5) |
| 業務ロジック | 判定・公差外・品質基準照合・ロット・測定機器・入力先・Min/Max・Aインプット・作業者・校正・保存・帳票・特殊検査(`ufkensa/kensa/`。約2.5万行・試験 700 件超) |
| 帳票 | 検査成績票・不良要因票を原紙どおりの HTML に。Edge のヘッドレス印刷で PDF(D7) |
| ログ | all-tools と同じ決まり(`docs/07_ログ.md`)。処理の開始・終了と処理時間、エラーの詳細 |
| 画面 | 検査(A案)・キャスク・全ヵ所/H494・Aインプット・校正・品質基準・履歴。どれもメニューから開ける |
| これから | 実機確認(U1〜U12)・業務確認(W1〜W15)の反映、直す不具合の反映(B1〜B83)、Windows 実機での Edge 印刷の確認 |

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
| [docs/04_不明点・判断事項.md](docs/04_不明点・判断事項.md) | **決定事項(D1〜D17)**、実機確認、業務確認、不足資料、既存不具合 B1〜B83 |
| [docs/05_AccessDB調査.md](docs/05_AccessDB調査.md) | 受領した Access 5 ファイルの表・列と VBA の参照の突き合わせ |
| [docs/06_意図した差異.md](docs/06_意図した差異.md) | VBA と意図して変えたところ |
| [docs/07_ログ.md](docs/07_ログ.md) | ログの置き場所・名前・記録する内容 |
| [tools/vba_verify/](tools/vba_verify/) | 実機確認マクロ(U1〜U12。Excel で1回動かして結果を送ってもらう) |
| [docs/analysis/](docs/analysis/) | モジュール群ごとの詳細な解析レポート(行番号つき) |
| [docs/foundation/](docs/foundation/) | 基盤の規約(複製元) |
| [vba_src/](vba_src/) | VBA の原文(93 モジュール) |

## 構成

    app/                 画面(Flask。ルート・テンプレート・JS・CSS)
    ufkensa/             基盤(起動・ログ・権限・設定・取り込み)
    ufkensa/kensa/       業務ロジック(VBA の移植。Python だけ)
    tests/               試験
