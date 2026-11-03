# 02 既存基盤の調査（python-web-tools と、その上の移植例 db-viewer）

詳細: [analysis/G7_python_web_tools.md](analysis/G7_python_web_tools.md)（python-web-tools VER4.8.123 を読んだ結果）。

## 結論

1. **python-web-tools は「梱包資材総合ツール」**（Flask + waitress + SQLite、Windows のラインPCで `Start.vbs` から起動、専用ウィンドウで表示）。検査判定と同じ機能（測定値入力・公差判定・検査成績票）は**ありません**。重複して作る心配はなく、再利用できるのは**基盤と共通部品**です。
2. **社内では「基盤を複製して、業務ごとに別アプリにする」形が定着しています。** db-viewer（VBA+Access の作業標準管理システムの移植）と packaging-*-python-web は、python-web-tools から起動・停止・外枠・共通部品を複製し、`FOUNDATION_ORIGIN.md` に複製元のコミットを記録しています。複数のツールは「業務ツール統合ランチャー」（python-business-tools-launcher）から起動します。
3. **VBA+Access からの移植部品は db-viewer にすでにあります。** UF検査の `【UFKen_15】SQL_追加`（adoSQL / GetRecordsArr）は db-viewer が移植した VBA 共通モジュールと同じ系統で、そのまま流用できる見込みです。

| db-viewer の部品 | UF検査での使い道 |
|---|---|
| `dbviewer/accdb_bridge/bridge.vbs` + `importer.py` | **cscript → VBScript → ADODB → ACE** で .accdb を読み、TSV(UTF-16) で Python に渡す。ご指定の経路そのもの |
| `dbviewer/ado_sql.py` | VBA `adoSQL`（ADDNEW/UPDATE の自動切替、全値を文字列で書く、1列目を鍵）の移植 |
| `dbviewer/vba_compat.py` | VBA の `CDate`/`IsDate`/`Format`/四捨五入などの再現。「直したくなっても直さない」方針も同じ |
| `dbviewer/settings_file.py` / `local_paths.py` / 設定画面 | **パスを画面から設定する仕組み**（接続先・参照フォルダ・出力先） |
| `dbviewer/shared_db.py` / `db.py` | 共有 `.sqlite3` を直接読み書き（`journal_mode=DELETE`、busy_timeout 10秒） |
| `docs/03_移植計画.md` の「追加確認事項への設計方針」 | Access→SQLite、複数PC、長時間放置への方針の前例 |

## python-web-tools から再利用するもの

| 区分 | 部品 |
|---|---|
| 起動・停止 | `Start.vbs` / `start.bat` / `stop.bat` / `start_app.py` / `launch_guard.py` / `server.py` / `boot_server.py` / `process_manager.py` / `idle_exit.py`（多重起動防止、待機画面、タブを閉じたら自動終了） |
| 設定・ログ | `app_config.py` + `config/app.json`、`user_settings.py`、`logging_utils.py` / `log_names.py` / `trace_log.py`（操作番号・エラー番号） |
| DB | `sqlite_toolkit.py`（業務非依存。ロック時 1→2→3 秒の再試行、WAL）、`source_db.py`（共有 sqlite3 の安全な読み書き、`BEGIN IMMEDIATE`）、`outbox_sync.py`（手元に書いてから共有へ送る。送信IDの一意索引で二重登録を防ぐ） |
| 権限 | `access_control.py`（Windows ログインID × PC名）、`line_names.py`（L-1・LVC・HVC… の社内共通の読み方）、`admin_password.py` |
| 長い処理 | `jobs.py` / `stall_watch.py` |
| 帳票 | `printing.py`（HTML + 印刷用 CSS をブラウザで印刷） |
| 画面 | `base.html` + `shell.py`（外枠）、`tokens.css`（ライト/ダーク/OS追従）、`api.js` `toast.js` `busy.js` `unsaved.js` `drafts.js` `askbox.js` `tabs.js` `errorlog.js` `health.js` |
| 試験 | unittest + Flask test_client、`tests/_isolation.py`、Windows CI |
| ファイル出力 | CSV は `utf-8-sig`、サーバ側のフォルダへ書き出し（ダウンロード方式ではない） |

## 合わないところ・足りないところ

| 項目 | 状況 | 対応案 |
|---|---|---|
| **追加ライブラリ** | python-web-tools も db-viewer も **Flask + waitress が必須**（`start_app.py` が起動時に確認）。「ラインPCの Python に追加ライブラリを入れられない」制約と衝突 | 判断事項 D1。純 Python なので**リポジトリに同梱（vendor）して `sys.path` に足せば pip 不要**で動く（`vendor/accdb_converter/_libs` と同じ方式） |
| Access 接続 | python-web-tools には cscript/ADODB 経路が無い | db-viewer の `bridge.vbs` を拡張（読みに加えて、1行追加・削除） |
| ラベルプリンタ | 無い | UF検査にもラベル印刷は無い（G6）。不要 |
| PDF 保存・パスワード付き xlsx | 無い（標準ライブラリでは作れない） | 判断事項 D7 |
| 作業者 | 身元は Windows ログインIDのみ | 社員カード13桁 → 名簿（伝送用ファイル）を移植 |
| 1PC = 1ライン | python-web-tools は「1PC=1人・プロセスに1状態」 | UF検査の運用と同じ（設備名で1ライン）。そのまま合う |
