# G7: 既存 `python-web-tools`（梱包資材総合ツール）構造調査 — UF検査判定システム移植のための再利用性評価

- 対象: `/home/user/vivio-hellon/python-web-tools`（shallow clone、HEAD `25f661f` = VER4.8.123、2026-10-30）
- 調査方法: 読むだけ（変更なし）。ファイル数 約450（`tests/` 150、`packaging_tool/` 約150モジュール）
- 表記: **【事実】**=コード/文書で確認、**【推測】**=根拠からの推定

---

## 0. 結論サマリ（最重要ポイント）

1. **【事実】このツールは「標準ライブラリのみ」ではない。** `requirements.txt` は `Flask~=3.0` と `waitress~=3.0` の2つ。`app/` 全体が Flask（Blueprint・Jinja2・`g`・`before_request`）で書かれ、`start_app.py` の `REQUIRED_PACKAGES = (("flask","Flask"),("waitress","waitress"))` で起動時に有無を検査し、無ければ「必要なパッケージが入っていません」で止まる。デスクトップ版（Tauri exe + `bridge.py`）でも Flask は必須（waitress だけ不要）。
   → **ラインPCに追加ライブラリを入れられない制約とは正面から衝突する。** 既存基盤の上に「新しい画面」として載せるなら、Flask/waitress をラインPCで使えるようにする（同梱 or 例外承認）か、Web層を作り直すかの判断が最初に必要。
2. **【事実】Access への接続は cscript/VBScript/ADODB ではない。** リポジトリ全体に `cscript` / `ADODB` / `ACE.OLEDB` は1件も無い。取り込み元は共有フォルダ上の **`.sqlite3` ファイル**（Access から変換済み）で、`source_db.py` が標準 `sqlite3` で読む。唯一 Access を直接読む経路は設定画面の「表を持ってくる」（`table_bring.py`）で、同梱の `vendor/accdb_converter`（`pyodbc` + ACEドライバ、無ければ同梱の純Python `access_parser`）を**別プロセス**（`subprocess.Popen`）で動かして sqlite3 に変換する。
   → 「cscript→VBScript→ADODB→ACE」の部品は**新規に作る必要がある**。ただし「別プロセスで動かして結果を受け取る」枠組み（`table_bring._run_child` / `_convert` / `kill_children`）は設計の手本になる。
3. **【事実】業務と無関係な共通基盤が厚い。** SQLite ヘルパ（`sqlite_toolkit.py`）、共有sqlite3 の安全な読み書き（`source_db.py`）、二重送信防止の書き戻しエンジン（`outbox_sync.py`）、ログ/エラー番号（`logging_utils.py`・`trace_log.py`）、権限（`access_control.py`・`line_names.py`）、HTML帳票（`printing.py`）、長時間処理（`jobs.py`）、フロントの共通JS部品（`api.js`/`toast.js`/`busy.js`/`askbox.js`/`drafts.js`…）は**業務非依存**と明記されており、検査判定でもそのまま使える。
4. **【事実】「プラグイン機構」は無い。** 画面の追加は `app/__init__.py::_register_routes` に Blueprint を1行足し、`app/shell.py` の `FIELD_NAV`/`MATERIAL_NAV` と `READY_SCREENS` に1行足し、`app/templates/<画面>.html` と `app/static/js/views/<画面>.js` を作る、という**手作業の定型**。
5. **【推測】`db.py` の docstring に「VBA `UFdaily` 標準モジュールのADOラッパー群の移植」とある**（`adoConnection`→`get_connection`、`ExecuteSQLWithRetry`→`execute_with_retry`、`IsLockError`→`is_lock_error`、`NowDBString`→`now_db_string`、`SanitizeForDB`→`sanitize_for_db` 等）。「UF検査判定システム」が同じ `UF*` 系の VBA 共通モジュールを使っているなら、`sqlite_toolkit.py` はほぼ1対1の移植先になる。要確認。

---

## 1. 全体構成

### 1.1 起動経路（【事実】`docs/設計.md` §2 / §2.5）
```
Start.vbs ─▶ pythonw start_app.py ─▶ launch_guard.py(多重起動判定・ポート選択) ─▶ server.py(waitress)
                                                         ├─ boot_server.BootApp(待機画面だけ・標準ライブラリのみ)
                                                         └─ Handover.install() で app.create_app()(Flask) に差し替え
start.bat ─▶ python start_app.py %*（診断用。--check / --mode material）
stop.bat  ─▶ process_manager.py ─▶ POST /api/shutdown → 応答なければ PID で停止
画面を閉じる ─▶ 心拍途絶 ─▶ packaging_tool/idle_exit.py ─▶ 同じ停止経路
デスクトップ版: 梱包資材総合ツール.exe(src-tauri, Rust) ─標準入出力─▶ bridge.py ─▶ Flask(WSGIとして呼ぶだけ、ポート無し)
```
- `Start.vbs` / `start.bat` は **CP932 + CRLF 固定**（`tests/test_launch_files.py` が検査）。`Start.vbs` は `python --version` / `pythonw --version` を確認してから `pythonw start_app.py` を投げる。引数はわざと渡さない（ランチャー連携の都合、`docs/ランチャー連携.md`）。
- `boot_server.py`: Flask を import する前に待機画面を返す最小 WSGI（`BootApp`）。`/`・`/api/health`・`/api/shutdown`（トークン必須）だけ答え、他は 503「起動しています」。`Handover` が WSGI の呼び先を差し替えるので待ち受けを開き直さない。
- `server.py::AppServer`: waitress `create_server(threads=8, host=127.0.0.1)`。`stop()` は二重呼び出し安全・keep-alive 接続も `wasyncore.close_all` で閉じる。
- `config/app.json`: `app_id`=`nlm.packaging-tool`、`display_name`、`version`=4.8.123、ポート（field 8713 / material 8723）、`monitoring`（health 15秒、job 300ms）。読み手は `packaging_tool/app_config.py` のみ。

### 1.2 Webサーバ実装と requirements（【事実】）
- **Flask + waitress**（自前 `http.server` ではない）。`requirements.txt` のコメント「依存はこの2つだけに抑える」「npm/webpack は持ち込まない」。
- 標準ライブラリだけで書かれているのは起動前段（`boot_server.py`、`packaging_tool/boot_screen.py`、`app_config.py`）と業務ロジック層（`packaging_tool/*`。Flask を import しない）。**Flask 依存は `app/` 配下と `server.py`/`bridge.py` に閉じている**。
- 前例: `vendor/accdb_converter/_libs/` に純Pythonライブラリ（access_parser・construct・tabulate）を**同梱**して `sys.path` 末尾に足す運用がある（`table_bring.BUNDLED_LIBS`）。【推測】Flask/Werkzeug/Jinja2/MarkupSafe/itsdangerous/click/blinker と waitress はすべて純Pythonで動くため、同じ「vendor 同梱」方式でラインPCの pip 不要化は技術的には可能。ただし社内規程上「追加ライブラリ」に当たるかは要確認。

### 1.3 ディレクトリ構成（【事実】README「ディレクトリ構成」）
```
Start.vbs / start.bat / stop.bat / start_app.py / launch_guard.py / server.py / boot_server.py / process_manager.py / bridge.py
app/                    Flask層
  __init__.py           create_app(mode)・セキュリティ・書き込み錠・DB接続・Blueprint登録
  shell.py              外枠(レール/リボン)の値。NavItem / FIELD_NAV / MATERIAL_NAV / READY_SCREENS
  routes/*.py           Blueprint(catalog, health, inventory, layout, log, lot, manual, master, pallet_fit, pending, selection, settings, spec_sheet, trace, warehouse)
  templates/*.html      Jinja2(base.html を継承)
  static/css/           tokens.css(色の唯一の出どころ) base.css components.css layout.css lotmodal.css
  static/js/            共通部品 + views/<画面>.js(素の ES modules)
packaging_tool/         業務ロジック(Flask非依存)
  presenters/           ビューモデル組み立て・状態機械(「押せるか」等の判断は全部ここ)
  *_session.py          画面をまたぐ作業状態(プロセスに1つ・保存しない)
  *_service.py          業務ロジック
  db.py / sqlite_toolkit.py / source_db.py / outbox_sync.py / data_sync.py + sync_*.py  データ層
  schema.sql            手元SQLiteのスキーマ(41表)
config/app.json         アプリ固有値
scripts/                init_db / import_source / dedupe_writeback / compare_ui / check_contrast / check_encoding / make_dist …
tests/                  unittest 141本 + golden/ + js/(.mjs 3本)
vendor/accdb_converter/ Access→sqlite3 変換部品(同梱)
src-tauri/              デスクトップ版の外枠(Rust/Tauri)
docs/                   設計.md / UIUX設計指針.md / はじめに読む_起動と使い方.md / ランチャー連携.md / 変更履歴.md(870KB) …
```

### 1.4 ルーティングと画面追加の手順（【事実】）
- `app/__init__.py::_register_routes()` が全 Blueprint を**ハードコードで**登録。モード（現場 `field` / 資材 `material`、`packaging_tool/modes.py`）や権限で登録/拒否を分ける（現在は全部登録し `before_request` で要求ごとに403/404）。
- レールは `app/shell.py::FIELD_NAV` / `MATERIAL_NAV`（`(key, label, url, note)` のタプル）。`READY_SCREENS` に無いキーは `app/routes/pending.py::register()` が自動で「準備中」画面を生やす。
- `catalog.html`（`/__catalog`, `app/routes/catalog.py`）は**部品カタログ（デザインシステム見本帳）**で、プラグイン機構ではない。コントラスト検査結果（`scripts/check_contrast.py`）と配置図サンプルを表示。
- **新しい画面の追加手順（実質の定型）**:
  1. `app/routes/<new>.py` に `bp = Blueprint(...)`、画面 `GET /<new>` と `GET/POST /api/<new>/...`
  2. `app/__init__.py::_register_routes` に `app.register_blueprint(<new>.bp)`
  3. `app/shell.py` の NAV 表と `READY_SCREENS` に1行
  4. `app/templates/<new>.html`（`{% extends "base.html" %}`、`shell_context(key)` を渡す）
  5. `app/static/js/views/<new>.js`（`nav.js` の `start()`/`onLeave()` 規約に従う）
  6. 業務データを返す経路は `/api/` か `/report/` 配下に置く（`TOKEN_REQUIRED_PREFIXES`）
  7. `packaging_tool/<new>_service.py` + `presenters/<new>.py` に判断を置く（JSに判断を置かない＝設計.md §1-1）

### 1.5 守っている約束（【事実】`docs/設計.md` §1、違反はバグ扱い）
業務判断はすべて Python／APIは更新後のビューモデル一式を返す／HTTP ステータスで断りを区別（400 形の誤り・422 業務の断り・403 資格なし・409 先に変えられた）／断りの種類は `reason` 定数／同じ事実を2か所に持たない／一覧に無いものは選べない（サーバでも検証）／色だけで意味を運ばない（色は `tokens.css` のみ）。
エラー応答形式: `{"error":{"code":"...","message":"現場向け日本語"}}`（`app/__init__.py::_error`）。

### 1.6 セキュリティ（【事実】`app/__init__.py::_register_security`）
127.0.0.1 バインドのみ／起動ごとのトークン（`X-Tool-Token` or `?t=`、`/api/*` と `/report/*` に必須。`/api/health`・`/api/alive` は免除）／Host ヘッダ検証／`Sec-Fetch-Site` で同一オリジン確認／CORS ヘッダ無し。

---

## 2. フロントエンド

### 2.1 テンプレート（【事実】）
- Jinja2。`app/templates/base.html` が外枠（接続断の帯 `#offline`、長処理帯 `#busyband`、リボン（Lot/パレット/製品/拠点/共有）、モード切替、説明書・テーマ切替・終了ボタン、左レール `<nav class="rail">`、`<main>`、トースト置き場）。各画面は `{% extends %}`。
- **SPAではない**が、`static/js/nav.js` が行き先HTMLを fetch して `<main>`/リボン/レール/画面ごとの `<style>`/`<script>` だけ差し替える（白画面を挟まない）。ビルド工程なし（素の ES modules）。
- CSS/JS は `url_for('static', ...)` に `?v=<版>` を自動付与（`_register_static_version`）。

### 2.2 CSS / デザイントークン / ダークモード（【事実】`app/static/css/tokens.css` 306行）
- 色・書体・寸法の唯一の出どころ。`:root`（ライト）／`@media (prefers-color-scheme: dark)`（明示指定なし時）／`[data-theme="dark"|"light"]`（利用者が明示）の3状態。
- 意味別トークン: `--ground/--surface/--ink/--muted/--rule/--accent`、操作色 `--act-nav/run/print/select/find/danger/commit`、状態 `--state-ok/warn/error/info`、`--warn-bg`/`--ok-bg` 等。骨格色はティール系。社内 SymfoNavi Data Hub の配色に合わせたと記載。
- テーマは `user_settings.get_theme()/set_theme()` で端末に保存し、`<html data-theme>` に反映。
- `scripts/check_contrast.py` が3テーマ×約300組の WCAG コントラスト比を検査。
- **検査判定の OK/NG 表示**: `--state-ok`/`--state-error`/`--ok-bg`/`--warn-bg` がそのまま使える（色＋文言＋形の併用ルールあり）。

### 2.3 共通JS部品（【事実】`app/static/js/`）
| ファイル | 役割 | 検査で再利用 |
|---|---|---|
| `api.js` | fetch ラッパ。トークン・`X-Screen-Id` 付与、`ApiError`（`{"error"}` 形式と 422 の最上位 `message` の両方を拾う）、`tokenUrl()` | ◎ |
| `toast.js` | 作業を止めない通知（ok/info/warn/error、errorは自動で消えない、アクション付き） | ◎ |
| `busy.js` | 押したボタン自身を 250ms 後に待機表示・1.2秒後に経過秒・二重押し防止（`press()` / `watchClicks()`） | ◎（保存・印刷の二度押し防止） |
| `unsaved.js` | 保存していない編集があるときタブを閉じさせない（`mark(key,dirty)` / サーバの `/api/health` の `unsaved` と連動） | ◎（入力途中の測定値） |
| `drafts.js` | 打ちかけ入力を `localStorage` に自動退避し、再表示時に復元（`bind`/`done`、元値が変わっていたら捨てる、14日で失効） | ◎（測定値の打ちかけ保護） |
| `askbox.js` | 自前の確認/入力ダイアログ（`confirmBox`/`promptBox`。Tauri で `window.confirm` が素通りする問題の対策） | ◎ |
| `tabs.js` | タブ切替（キーボード操作、バッジ、`sessionStorage` に選択を記憶） | ○ |
| `errorlog.js` | 画面のJSエラーを `/api/client-error` へ送り、エラー番号を表示（1画面10件まで） | ◎ |
| `jobs.js` / `running.js` | `/api/jobs` を1か所で見張り（走行中0.7秒・平常15秒）、帯に進捗表示 | ○（取り込み・一括送信） |
| `health.js` | 生存監視・心拍（`/api/alive` 20秒ごと）、接続断の帯、権限/マスタ世代の変化検知 | ◎（外枠ごと） |
| `nav.js` | 画面差し替え遷移、`pageSignal()`/`onLeave()` でリスナ後始末 | ◎（外枠ごと） |
| `screen.js` | 2枚目のタブを開いたら古いタブに覆い（`screen_lock.py` と対） | ◎ |
| `authsync.js` / `datasync.js` | 管理者認証/マスタの「世代」が変わったら画面へイベント（BroadcastChannel 併用） | ○ |
| `toggles.js` | 択一ボタン（`aria-pressed`）の見た目統一 | ○ |
| `halfwidth.js` | 全角→半角の比較用関数（サーバの `text_input.py` と同じ規則） | ○ |
| `desktop.js` | Tauri 版で別窓を開く（帳票ウィンドウ）／ブラウザ版は `window.open` | ◎（印刷窓） |
| `svgplan.js`/`mapedit.js`/`lotdetail.js`/`pallet_fit.js` 等 | 梱包資材固有 | × |

---

## 3. DBアクセス

### 3.1 手元SQLite（作業用DB）（【事実】`packaging_tool/db.py`）
- 場所: `config.DB_PATH` = `<アプリフォルダ>/data/packaging_tool.db`（環境変数 `PACKAGING_TOOL_DB_PATH` で変更可）。**端末ごと**（配布はフォルダごとコピー、`data/` は追跡しない）。
- `get_connection()`: `sqlite3.connect(path, timeout=10)`、`row_factory=sqlite3.Row`、`PRAGMA foreign_keys=ON`、`PRAGMA busy_timeout=10000`、`sqlite_toolkit.enable_wal()`（`journal_mode=WAL`、失敗時は警告して既定のまま）。`isolation_level` は既定（DEFERRED）で `with conn:` をトランザクション単位に使う。
- **接続はリクエストごとに開閉**（`app/__init__.py::get_db()` が `g.db` に1本、`teardown_appcontext` で close）。理由: waitress のスレッドプール＋`check_same_thread`。
- **書き込みはプロセス内で直列化**: `_WRITE_LOCK = threading.RLock()`、GET/HEAD/OPTIONS 以外を1本ずつ通す（方法だけで判定）。待ち上限 `WRITE_WAIT_SEC=30` 超過で 503「前の操作がまだ終わっていません」。例外は `/api/jobs`・`/api/mode`・`/api/log`・`/api/alive` 等。
- 起動時 `db.prepare()`: `schema.sql` 適用（冪等）、壊れていれば `quarantine()` で `-wal/-shm` ごと退避して作り直し（消さない）。`_migrate_before_schema`/`_migrate_after_schema` で列追加等の移行。
- `db.connect()` コンテキストマネージャ、`execute_with_retry`/`fetch_all`/`fetch_one`/`update_record`/`insert_record` を `sqlite_toolkit` から再エクスポート。

### 3.2 `sqlite_toolkit.py`（【事実】業務非依存と明記、「他のVBA移行プロジェクトでもそのまま使い回せる」）
- `is_lock_error(exc)`、`enable_wal(conn)`、`execute_with_retry(...)`、`fetch_all(...)`、`fetch_one(...)`（ロック時のみ `retry_wait_sec*retry_count` 待って最大3回）、`update_record(conn, table, key_field, key_value, values, where_clause=...)`（存在確認→UPDATE→影響0なら `not_found`）、`insert_record(...)`、`now_db_string()`、`sanitize_for_db()`。
- 戻り値は `ExecResult` / `UpdateResult`（`ok`/`reason`/`rowcount`/`error`）。

### 3.3 取り込み元（共有フォルダ上の sqlite3）（【事実】`packaging_tool/source_db.py` 1200行超）
- 読み取りは `file:...?mode=ro` → 素のパス → **手元への写し**（`-wal/-shm` ごとコピー）の順に試す（`_try_uri`/`_try_plain`/`_try_copy`）。SMB上のWALファイルは読むだけでも開けないため。開けたかは `SELECT count(*) FROM sqlite_master` で確認してから返す。
- `probe(path)`: サイズ・マジック・付き添いファイル・試した順と理由を返す（設定画面が「次にすること」を表示）。`sniff_encoding`/`make_text_factory`: CP932 混入の文字化け対策。
- 書き込みは `SourceConnection`（`connect(path)`）: `query`/`execute`（1文ごとcommit）/`insert`/`update`/`delete`（**WHERE空は拒否**）、`transaction()` は `BEGIN IMMEDIATE`＋失敗で全ロールバック（`SourceTransaction`）。`BUSY_TIMEOUT_MS=10_000`。
- 例外分類: `is_lock_error`/`is_duplicate_error`/`is_already_exists_error`/`is_missing_error`、識別子は `quote_identifier()`（日本語列名対応）。

### 3.4 Access(.accdb) との関係（【事実】）
- **通常運用では Access を読まない。** 取り込み元は `梱包資材マスタ.sqlite3`・`看板マスタ.sqlite3`・`PalletThresholdMaster.sqlite3`・仕掛台帳 `SIKALOT/SIKAHIKI/SIKAODR.sqlite3`（上流で Access から変換済み）。`source_db.py` 冒頭:「以前は Access(.accdb)…ODBCドライバの有無で動いたり動かなかったり…sqlite3 になってその分岐は丸ごと不要」。
- 取り込み処理: `data_sync.py`（外向けの窓）→ `sync_sources.py`（ファイル探索 `find_material_db`/`find_lot_dbs` 等）、`sync_import.py`（`import_tables`/`import_master`/`import_lot_ledger`。**総入れ替え**、列の対応は `import_specs.py`、鍵列が無ければ取り込まない）、`sync_auto.py`（起動時自動取り込み、ファイルの更新印で要否判定）、`sync_refresh.py`（発注まわりの再取り込み）、`sync_writeback.py`（書き戻し）。
- `table_bring.py`（設定画面「表を持ってくる」）だけが .accdb/.mdb を扱う: `_convert()` が `_pythons()`（自分の Python → `py -3` → `python`）で `vendor/accdb_converter` を子プロセス実行（`-c _CONVERT_SCRIPT`、`CREATE_NO_WINDOW`、タイムアウト、`_children` 管理・`kill_children()`）。読み取りエンジンは `pyodbc`+ACE、無ければ同梱 `access_parser`。**cscript/VBScript/ADODB は未使用。**

### 3.5 スキーマ（【事実】`packaging_tool/schema.sql` 41表）
PalletMaster/BoardMaster 等の資材マスタ、資材パレット注文管理・パレット入出庫履歴・発注コメント・切断依頼・予定ロット等のツールが書く表、看板_*、アクセス権限、仕掛ロット/仕掛引当/仕掛受注、パレット閾値7表。冒頭コメントに Access→SQLite の型対応（Yes/No→INTEGER 0/1、日付→ISO8601 TEXT、オートナンバー→INTEGER PRIMARY KEY AUTOINCREMENT）。`master_schema.py` は「マスタ管理画面から取り込み元に表を作る/作り直す/列を足す」機能（DDL生成）。

---

## 4. 共有フォルダ上の SQLite 運用と複数PC同時書き込み対策

### 4.1 基本アーキテクチャ（【事実】`docs/設計.md` §8、`はじめに読む` §3「複数の端末で同時に使うとき」）
```
各PC: 手元DB(data/packaging_tool.db, WAL, 端末専用) ←取り込み(総入れ替え)── 共有 *.sqlite3(取り込み元)
                                      ──書き戻し(outbox)──▶ 共有 *.sqlite3
```
- **共有ファイルを直接の作業DBにしない**。各PCは手元DBで作業し、共有へは「足す行」と「印（確認済み等）」だけを送る。共有の sqlite3 は DELETE ジャーナル想定（WAL だと SMB で読めない→`probe` が `PRAGMA journal_mode=DELETE;` を案内）。

### 4.2 `outbox_sync.py`（【事実】「特定のテーブル名や業務に一切依存しない」書き戻しエンジン）
- `WriteBackSpec`（`sqlite_table`, `access_table`, `key_column`, `op_id_column`=「送信ID」, `use_op_id_guard`, `mark_columns`, `mark_pending`, `source_key`, `mark_blockers`, `unique_column`, `match_columns`, `on_insert`, `optional_columns`, `number_column`, `identity_columns`）を渡して `write_back(conn, access, specs)` を呼ぶだけ。
- 仕組み: 手元の同期記録表 `Access同期記録` に「送信中」で**予約（claim）**→送れたら「済」。送信中のまま10分（`STALE_CLAIM_MINUTES`）放置されたら拾い直し、送信IDは引き継ぐ。送り先に送信ID列＋一意索引を自動作成（`ensure_op_id_column`/`_create_op_id_index`）し、重複INSERTは「既に届いていた」扱い。`identity_columns` で業務上の不変キーにも一意索引（`ensure_identity_index`）。`number_column` で送り先に `MAX+1` 採番。1行の失敗で全体を止めない。印の後送り（`push_marks`、`UNMARK_PREFIX="外す:"`）。行の出どころ列（`ORIGIN_LOCAL="手元"`/`ORIGIN_IMPORTED="取込"`）で取り込んだ行を送り返さない。取り込み直後に `mark_all_sent()` で再送を防ぐ。
- `sync_writeback.py::WRITEBACK_SPECS` が業務の対応表（資材パレット注文管理・パレット入出庫履歴・ボード使用実績・発注コメント・…）。「新しいテーブルを足したくなったらここに1行足すだけ」。`write_back_in_background()` / `retry_unsent_in_background()`。ロック時は1→2→3秒で3回再試行、駄目なら次の機会。
- `shared_link.py`: 手元DBが「どの共有マスタの写しか」を表 `つながっている共有` に1行記憶。置き場所が変わったのに未送信行があれば**送らない**（本番⇔Test環境の混入防止）。`bind(conn, path)` → `Link.blocked`/`why()`。
- 同時更新検知: `UPDATE ... WHERE 管理番号=? AND 更新日時=?` の compare-and-swap（README「VBA版との主な違い」）。マスタ編集（`master_admin._write_row`）も読んだ時の値 `seen` と比較して409。
- 重複検出: `data_sync.duplicate_count()`、`dedupe.py`、`scripts/dedupe_writeback.py`。

### 4.3 検査判定への適用（評価）
- 検査結果（判定履歴）は「足すだけの行」なので **`outbox_sync` の適用が最も素直**（送信ID・`identity_columns`=例:検査ID で二重登録不可、オフライン時は手元に溜めて後送り）。
- ただし送り先が **Access(.accdb) のまま**なら、`outbox_sync` は `source_db.SourceConnection` 互換の口（`execute`/`insert`/`query`/`transaction`/`path`）を期待するので、**cscript+ADODB 版の `SourceConnection` 互換アダプタ**を作れば流用できる【推測：インタフェース上は可能だが、`ensure_op_id_column` の ALTER/CREATE INDEX は Access SQL 方言差があり要修正】。

---

## 5. マスタ管理の汎用性

（【事実】`master_admin.py` docstring「このツール固有なのは `master_common` の `MANAGED` と `VIEW_ONLY_WHY` だけ。他のツールへ持っていくときは、そこと `import_specs` を差し替えれば残りはそのまま動く」）

| モジュール | 役割 |
|---|---|
| `master_common.py` | `MANAGED`（直せる表の一覧: PalletMaster/BoardMaster/…/アクセス権限/閾値7表）、`VIEW_ONLY_WHY`、`can_edit()`（資材モード権限＋管理者認証、最初の1行は塞がない）、`source_for(table)`、書いた後の取り込み直し `_follow()` |
| `master_browse.py` | `tables(path)`（取り込み元の**全表**を列挙。MANAGED以外も「見るだけ」で出す）、`page(path, table, query, sort...)`（200行/ページ、検索・並べ替え） |
| `master_columns.py` | 列の形（宣言型→入力種別、NOT NULL 必須、更新日時の自動補完） |
| `master_schema.py` | 表の作成/作り直し/列追加/削除（取り込み元へDDL） |
| `master_admin.py` | 行の保存/追加/削除（取り込み元の `rowid` で行を指す、楽観ロック） |
| `master_state.py` | マスタ世代（変わったら画面へ通知） |
| `app/routes/master.py` | `/api/master/browse`・`/access`・`/row/save|add|delete`・`/table/create|rebuild|drop`・`/column/add` |
| `app/static/js/views/master.js` | 汎用テーブルエディタUI（設定画面内） |

- **任意のテーブルの「閲覧」は可能**（取り込み元 sqlite3 の全表）。**編集は `MANAGED` に登録した表＋「表を持ってくる」で足した表（`ツールで足した表` 登録簿）**。書き先は常に取り込み元の sqlite3（手元DBではない）。
- 制約: 取り込み元が **sqlite3 であること**が前提（`source_db` 経由）。Access を直接編集する機能は無い。また `import_specs`（列対応表）との整合が編集可能列の条件（`IMPORT_SPECS ∩ 手元スキーマ ∩ 取り込み元の列`）。
- 検査の「公差マスタ」「検査項目マスタ」を共有 sqlite3 に置くなら、`MANAGED` に1行足すだけで閲覧・編集UIが手に入る。Access に置いたままなら使えない（要アダプタ）。

---

## 6. 共通部品

### 6.1 設定管理（【事実】）
- `app_config.py`: `config/app.json` の読み手（app_id/display_name/version/port/host/monitor/local_dir）。ローカル領域 `%LOCALAPPDATA%\PackagingTool\{runtime,logs,pycache,cache,work,backup,data}`（`local_dir(name)`）。
- `config.py`: 業務定数（テーブル名 `TBL_*`、共有フォルダ既定 `DEFAULT_MASTER_DB_DIR` 等、環境変数で上書き可、`resolve_dir`（相対/絶対）、`master_db_dir()` 等は `user_settings` の保存値を優先）。`ADMIN_PASSWORD` 既定値あり（`PACKAGING_TOOL_ADMIN_PASSWORD`）。
- `user_settings.py`: `data/user_config.json`（端末ごと、非追跡）に `get`/`save`/`load_all`、読めない時は `SettingsUnreadable`（「たまたま読めない」と「壊れている」を区別、VER4.8.123）。拠点・テーマ等。VBA `GetSetting/SaveSetting` の代替。
- `distribution.py`: 1台で決めた設定を `配布設定\` に書き出し、配布先が起動時に読み込む（端末で変えた値は上書きしない）。

### 6.2 ログ（【事実】）
- `logging_utils.py`: 標準 `logging`。`get_logger(name)`。日次ファイル `<出力先>\packagingtool\packagingtool_<PC>_YYYYMMDD.log`（`DailyFileHandler`、非同期書き込み `_AsyncWriter`）。**操作番号**（1要求ごと、全行頭に `[番号]`）、ERROR以上に**エラー番号**を振り `packagingtool_エラー記録_<PC>_YYYYMM.jsonl` に直前の行・トレースバック・端末・利用者・版を記録。
- `log_names.py`: 命名規則（ツールID英小文字フォルダ＋`<ツールID>_種類_<PC>_日付`。「日報複合ツールに合わせた決まり」＝社内共通ルール）。
- `trace_log.py`: 出力先決定（環境変数＞設定画面＞ローカル）、書けなければローカルへ退避、エラー一覧/1件の詳細（同じ操作番号の行を集める）、**なぜなぜ分析用テキスト**出力、90日/24か月で片付け。設定画面「ログ」の面。
- `user_log.py`: 利用者向けの判断ログ（VBA `UserLog` の移植、画面 `/log` が差分取得）。`trouble.py`: 画面の記録に何を出すかの規則（続いた失敗だけ1件等）。

### 6.3 認証/権限（【事実】）
- `access_control.py`: **Windows ログインID（`USERNAME`）× PC名（`COMPUTERNAME`）** で `アクセス権限` マスタ（取り込み元に置く）を引く。`current_identity()`、`load_rules()`、`resolve()`→`Grant`（`allows_mode`）、該当行の和、両方空の行は無効、該当なしは最も狭い現場モード。権限コードは `PERMISSIONS` に足すだけで拡張（列追加不要）。ライン関連: `lines()`、`is_leader()`、`line_state()`、`registered_lines()`、`can_order()`、`is_administrator()`。
- `admin_password.py`: 誤操作ガード用パスワード（PBKDF2、画面で変更可）。`verify`/`change`/`reset`。
- `auth_state.py`: 管理者認証の「世代」番号（画面へ即時反映用）。
- **画面ロック**（`screen_lock.py` + `screen.js`）: 最後に開いたタブだけ操作可（`X-Screen-Id`）。※「離席時の画面ロック（パスワード再入力）」の機能は無い。
- `idle_exit.py`: 心拍途絶90秒／閉じた合図8秒でプロセス終了（処理中・未保存時は落とさない）。**長時間放置でログアウトする機能ではない**（窓が無いアプリのプロセス残留対策）。
- `stall_watch.py`: 20秒以上返らない要求のスタックをログに1回出す（フリーズ調査用）。
- `jobs.py`: `JobRegistry`（プロセスに1つ、同時に1つだけ、`start(kind,label,work)`、進捗 `_progress`、状態をファイル保存、再起動時は「中断」に倒す）＋`/api/jobs`。`work_status.py`: 裏処理の進み具合を帯に出す（`with work_status.work(...)`）。
- エラー表示: サーバ側 `{"error":{code,message}}` + `api.js::ApiError` + `toast.toastError`、JS例外は `errorlog.js`→`/api/client-error`（`app/routes/trace.py`）。

---

## 7. 印刷処理（【事実】）
- `printing.py`: **HTML + 印刷用CSS**（Excel 非依存）。`PageSetup`（向き・余白、`@page` に対応）、`Report`（`.sheet` 1つ=1ページ）、`render_html(report, edit_url, send_url)`、`table()`、`label_pairs()`、`print_bar()`（画面上の「印刷する/閉じる」帯、紙には出ない）、`editable()`（印刷前に点線欄を直せる）。文字列を返すだけ。
- `reports.py`: Lot貼付用（`build_label_report`、A4横・余白0）、切断依頼書（`build_cut_request_report`、A4縦）を VBA のセル結合を colspan/rowspan で再現。
- 経路: `/report/label` 等が HTML を直接返す → 画面は `desktop.openWindow()`（Tauri）/`window.open()` → `onload` で `window.print()`。**出せないときもHTMLを返す**方針。
- **ラベルプリンタ専用対応は無い**（ZPL/ESC-P等の直接送信なし）。README は「現物と寸法を合わせる帳票は実機確認」「Excelテンプレートの見た目を厳密に保ちたい場合は .xlsx を Excel に印刷させる選択肢もある」と記すが、現行コードにその経路は無い（`printing.py`:「tkinter版の OS 印刷動詞経路はもう誰も通らないので落とした」。README の `print_report(direct=True)` 記述は古い）。
- `spec_sheet.py`: 包装仕様書の図面を社内閲覧システムから `urllib` で裏取得・キャッシュ（印刷ではない）。

## 8. ファイル入出力（【事実】）
- CSV: `order_export.py::export()` — 標準 `csv`、`utf-8-sig`（BOM付き）、先頭0のコードは `="058901"` 形式（`excel_text`）、ファイル名に日時（上書きしない）。**サーバ側で設定フォルダ（`config.order_csv_dir()`）へ書き出し、画面に出力先パスを表示**する方式（ブラウザのダウンロードではない）。`board_usage.py` も同様。
- ダウンロード方式は取り込み記録 `/report/import-diag?save=1` のみ（`Content-Disposition: attachment`）。
- Excel(.xlsx) 出力: アプリ本体には無い。`vendor/accdb_converter/xlsx_writer.py`/`xlsx_reader.py`（純Python）が同梱されているが変換ツール用で、本体からは未使用。
- アップロード: `table_bring.save_upload()`（`/api/settings/table-bring/upload`）。
- サーバ側フォルダ参照: `/api/fs/list`（`presenters/fs_browse.py`、フォルダ名と取り込み元ファイル名だけ返す）。

## 9. ライン名・作業者・ロットの既存モデル（【事実】）
| 概念 | 既存実装 | 検査との共通化 |
|---|---|---|
| ライン | `line_names.py::OFFICIAL = ("L-1","LVC","HVC","機側","AIM","コイル","トット","バランサー","NS1")`、`official()`（全角半角・前後空白のみ正規化、`L1`/`LS` は読まない）、`LEADER="作業長"`、`ADMINISTRATOR`。docstring に「ほかのツール（看板・日報）と同じ読み方（`kanban/line_names.py`）」 | ◎ 社内共通規則。梱包課ラインと一致するか要確認 |
| 端末のライン | `access_control.lines(conn)`（アクセス権限マスタの「権限」列にライン名を書く） | ◎ 「このPCはどのラインか」をそのまま使える |
| 作業者 | **専用の作業者マスタは無い**。身元は Windows ログインID（`access_control.Identity.login_id`）。帳票の「担当者」は手入力欄 | △ 検査で作業者コード/氏名が要るなら新規 |
| ロット | `lot_service.py`（7桁 `LOT_NO_LENGTH=7`、`search_lot(conn, lot_no)`→`LotSearchResult`(LotInfo/HikiRow/OdrInfo)、仕掛台帳 `仕掛ロット/仕掛引当/仕掛受注` を手元へ取り込んで引く）、`lot_query.py`（条件チップ式一覧検索、列ホワイトリスト）、`lot_browse_session.py`、`lot_plans.py`（作業長がロットをラインに割当→ライン端末へ通知） | ◎ ロット引当の入口として流用可能性大（検査の「ロット引当」が同じ仕掛台帳を見るなら）【推測】 |
| 作業コンテキスト | `work_context.py::WorkContext`（lot_no・製品サイズ・パレット等。プロセスに1つ・保存しない、リボンに表示） | ○ 形は流用、項目は検査用に別クラスが妥当 |
| `models.py` | `BoardModel`/`PlacedBoardModel`（梱包資材固有） | × |
| 文字正規化 | `text_input.py`（全角→半角・大文字化・×統一、VBA `ToNarrowAlphaNumeric` 移植） | ◎ |

## 10. テストとCI（【事実】）
- `tests/` に `unittest` 141モジュール（「試験 3,000 件余り」`bridge.py` docstring）、Flask `test_client` を使うAPI試験11本、`tests/golden/`（選定・配置の41通り）、`tests/js/*.mjs` 3本（Node）。実行: `python3 -m unittest discover -s tests -t . -v`（`-t .` 必須。`tests/__init__.py`/`_isolation.py` がログ・設定を一時フォルダへ隔離）。表示環境不要。
- 運用系の試験: `test_launch_files.py`（.vbs/.bat の CP932+CRLF）、`test_ci_workflow.py`、`test_app_config.py`（版と `docs/変更履歴.md` の対応）、`test_templates.py` 等。
- CI `.github/workflows/desktop-windows.yml`: push 毎、`windows-latest`、Python 3.12、`pip install -r requirements.txt pytest` → `pytest -q`（unittest を pytest で回す）→ Rust `cargo test` → `cargo build --release` → `scripts/desktop_smoke.py` で exe 起動確認 → `PackagingTool.exe` を artifact 化。
- 補助: `scripts/compare_ui.py`（ゴールデン比較）、`scripts/check_contrast.py`、`scripts/check_encoding.py`。`docs/設計.md` §7 は Playwright による画面試験にも言及（リポジトリ内の実体は1ファイルの参照のみ）。

## 11. README・マニュアルから読み取れる運用ルール（【事実】）
- 利用者がダブルクリックするのは **`Start.vbs` だけ**（デスクトップ版は exe）。不調時は `start.bat`。停止は `stop.bat` または画面の「終了」。
- 1台の PC = 1人が使う前提。作業状態はプロセスに1つ。タブ2枚は最後の1枚だけ操作可。
- 配布は**フォルダごとコピー**。`data\`（DB・設定・配置図）は端末ごとで配布物に含めない。版入替時は `data\` を持っていく。`scripts\make_dist.bat` で配布フォルダ作成、`配布設定\` で設定を配る。
- ログは `%LOCALAPPDATA%\PackagingTool\logs\packagingtool\`（設定で共有フォルダ指定可、PC名入りで混ざらない）。
- 版は `config/app.json` が唯一の出どころ。上げたら `docs/変更履歴.md` に節が必須（試験で検査）。
- 共有フォルダの取り込み元を本番⇔Test で取り違えないよう、リボンに「共有」フォルダ名を常時表示。
- 権限はアクセス権限マスタ（ログインID×PC名×権限コード）で管理。
- 印刷は初回に実機で1枚テスト（ブラウザ余白でmm単位のずれ）。
- `docs/UIUX設計指針.md`（56KB）に画面設計の判断基準（認知心理学・情報構造・配色）。`docs/決め打ちの値一覧.md` は許容値をコードから自動生成（`scripts/make_constants_doc.py`）── **公差等の定数を文書化する仕組みとして参考になる**。
- 業務ツール統合ランチャー（`python-business-tools-launcher`）との連携規約（`docs/ランチャー連携.md`）: 起動ファイル・`config/app.json`・`stop.bat`・`/api/health` だけでランチャーが扱う。社内に「看板」「日報」「all-tools」等の兄弟ツール群がある【事実：docstring/文書中の言及】。

---

## 12. 「UF検査判定」移植時の提案一覧

前提の分岐（最初に決める）:
- **A案: 本ツール（Flask基盤）に画面を足す** → ラインPCに Flask/waitress が必要（同梱 vendor 化 or 例外承認）。共通部品をほぼ全部使える。
- **B案: 標準ライブラリだけの別アプリ** → Web層（Flask依存の `app/`）は使えない。`packaging_tool/` の Flask 非依存モジュールと JS/CSS 資産を「コピーして」使う（`boot_server.py` のような素の WSGI + `wsgiref`/`http.server` を自作）。【推測】`boot_server.BootApp` が「標準ライブラリだけで答える WSGI」の実例として下敷きになる。

### 12.1 そのまま再利用できるもの
| 部品 | ファイル | 備考 |
|---|---|---|
| SQLite 汎用ヘルパ（リトライ・CRUD・WAL・ロック判定） | `packaging_tool/sqlite_toolkit.py` | 業務非依存・標準ライブラリのみ。B案でもコピー可 |
| 手元DBの開き方・壊れたDBの退避・スキーマ適用 | `packaging_tool/db.py`（`get_connection`/`prepare`/`quarantine`/`apply_schema`） | スキーマとマイグレーション部分は差し替え |
| 共有 sqlite3 の安全な読み書き | `packaging_tool/source_db.py` | 共有マスタを sqlite3 化するなら必須級 |
| 書き戻しエンジン（予約・送信ID・一意索引・再送・印） | `packaging_tool/outbox_sync.py` | 検査結果の共有送信に最適。`WriteBackSpec` を1行書くだけ（送り先が sqlite3 の場合） |
| 共有先の取り違え防止 | `packaging_tool/shared_link.py` | |
| ログ・操作番号・エラー番号・なぜなぜ出力 | `logging_utils.py` / `log_names.py` / `trace_log.py` / `trouble.py` | ツールIDを変えるだけ（`log_names.TOOL_ID`） |
| アプリ固有値・ローカル領域 | `app_config.py` + `config/app.json` | |
| 端末設定 | `user_settings.py` | |
| 権限（ログインID×PC名） | `access_control.py` | 権限コードを足すだけ（例 `inspect:judge`、`inspect:master`） |
| ライン名の正規化 | `line_names.py` | 社内共通規則 |
| 管理者パスワード（誤操作ガード） | `admin_password.py` / `auth_state.py` | 公差マスタ編集・判定の取消などに |
| 長時間処理・進捗 | `jobs.py` / `work_status.py` / `stall_watch.py` | 取り込み・一括送信 |
| 起動・停止・多重起動・自動終了 | `Start.vbs`/`start.bat`/`stop.bat`/`start_app.py`/`launch_guard.py`/`server.py`/`boot_server.py`/`process_manager.py`/`idle_exit.py`/`screen_lock.py` | A案なら全部そのまま |
| HTML帳票の土台 | `printing.py`（`Report`/`PageSetup`/`render_html`/`table`/`label_pairs`/`print_bar`/`editable`） | 検査成績書・A4帳票はこれで可 |
| 文字正規化 | `text_input.py`、`halfwidth.js` | 全角数字の測定値入力対策 |
| CSV出力の作法 | `order_export.py`（`utf-8-sig`・`excel_text`・日時入りファイル名・出力先パス表示） | 関数は業務依存なので**作法を流用** |
| フロント共通部品 | `api.js` `toast.js` `busy.js` `unsaved.js` `drafts.js` `askbox.js` `tabs.js` `errorlog.js` `health.js` `nav.js` `screen.js` `jobs.js` `running.js` `desktop.js` `toggles.js` | 判断はサーバ・JSは描くだけの規約込み |
| デザイントークン・ダークモード・コントラスト検査 | `tokens.css` / `base.css` / `components.css` / `layout.css` / `scripts/check_contrast.py` | OK/NG は `--state-ok/--state-error` 系 |
| 外枠テンプレート | `app/templates/base.html` + `app/shell.py` | A案 |
| 試験の隔離・CI | `tests/__init__.py` / `tests/_isolation.py` / `.github/workflows/desktop-windows.yml` | |

### 12.2 拡張が必要なもの
| 対象 | 何を拡張するか |
|---|---|
| `app/shell.py`（FIELD_NAV/MATERIAL_NAV/READY_SCREENS）、`app/__init__.py::_register_routes` | 検査画面のレール項目と Blueprint 登録。【推測】検査は「現場/資材」とは別の利用者なので、`modes.py` に新モード（例 `inspect`、ポート別）を足すのが自然（`modes.ALL` に1行＋`config/app.json` の `roles` にポート） |
| `modes.py` / `config/app.json` / `access_control.permissions()` | 検査モード・検査権限コード |
| `master_common.MANAGED` / `import_specs.py` / `schema.sql` | 検査項目・公差マスタを共有 sqlite3 に置く場合、表を足すだけでマスタ管理UIが使える |
| `sync_writeback.WRITEBACK_SPECS` | 検査結果ヘッダ/明細の書き戻し定義（ヘッダ＋明細は `SourceTransaction` で同一トランザクション） |
| `lot_service.py` / `lot_query.py` | 検査の「ロット引当」が仕掛台帳と同じなら流用、引当ロジックが違えば検査用に別関数（`search_lot` の結果を入力に使う） |
| `work_context.py` / リボン（`base.html` の `ribbon__slots`） | 検査中ロット・ライン・作業者を常時表示する項目（Lot/パレット/製品の固定スロットなので、モード別に切替が必要） |
| `printing.py` | ラベルプリンタ対応（用紙サイズ指定 `@page size: 100mm 50mm` 等、余白0、バーコード/QRが要るなら純Python生成）【推測】。ブラウザ印刷の余白依存問題は既知の「分かっている穴」 |
| `source_db.SourceConnection` 互換の **Access(ADODB) アダプタ** | `outbox_sync`/`master_admin` を Access 宛てに使うには、`execute`/`insert`/`query`/`transaction`/`path` を cscript 経由で実装した互換クラスが要る |
| `user_log.py` | 判定理由（どの項目が公差外か）を利用者向けログとして残す |

### 12.3 新規に作るべきもの
| 対象 | 内容 |
|---|---|
| **cscript→VBScript→ADODB→ACE OLE DB ブリッジ** | 既存に一切無い。`subprocess`（`CREATE_NO_WINDOW`・タイムアウト・子プロセス管理は `table_bring._run_child`/`kill_children` を手本）で `cscript //nologo //U x.vbs` を呼び、結果を UTF-16/JSON/CSV で受け取る部品。SELECT/INSERT/UPDATE、パラメータ（ADODB.Command）、ロック・再試行、文字コード、32/64bit ACE の違い（`%windir%\SysWOW64\cscript.exe`）【推測：要実機検証】 |
| 検査判定ロジック | 測定値×公差（上限/下限/基準±、片側公差、境界の含む/含まない、丸め・有効桁）→ OK/NG。`presenters/` に状態機械、判定は Python のみ（設計約束 §1）。**ゴールデン試験**（`scripts/compare_ui.py` 方式）で VBA と突き合わせ |
| 検査セッション | `inspection_session.py`（ロット引当→測定入力→判定→保存→印刷 の段遷移、プロセスに1つ、`selection_session` の作り方を踏襲） |
| 画面 | `app/routes/inspection.py`、`app/templates/inspection.html`、`app/static/js/views/inspection.js`（測定値入力グリッド、Enter で次欄、NG 即時表示、`drafts.js` で打ちかけ保護、`unsaved.js`） |
| 検査結果のスキーマ | 手元 `schema.sql` に検査ヘッダ/明細（検査ID=不変キー、ロット、ライン、作業者、判定、測定値、公差の写し、日時、送信ID列は outbox が付与） |
| 作業者マスタ/作業者選択 | 既存は Windows ログインIDのみ。ライン作業者を選ぶ必要があれば新規 |
| ラベル/帳票テンプレート | 検査ラベル・検査成績書（`reports.py` に倣い `build_*_report`→`printing.Report`） |
| 公差等の決め打ち値の文書化 | `scripts/make_constants_doc.py` の対象に検査定数を追加、または検査版を新設 |

### 12.4 リスク・確認事項
1. **【事実】Flask/waitress 依存 vs 標準ライブラリ制約**。本ツールの現場端末には Flask が入っている前提（`はじめに読む` §1「Flask と waitress: 初回に1度だけ pip install」）。UF検査のラインPCが同じ運用を許せないなら A案は不可。【推測】vendor 同梱（`vendor/accdb_converter/_libs` と同方式）が落としどころ候補。
2. **【事実】共有データの前提が「sqlite3」**。UF検査が Access(.accdb) を共有DBとして読み書きし続けるなら、既存の取り込み/書き戻し/マスタ管理は**直接は使えない**（アダプタ前提）。Access 側を sqlite3 化できるなら再利用度が大幅に上がる。
3. **【事実】1PC=1人・プロセスに1状態**。1台で複数ラインを並行検査する運用だと前提が崩れる。
4. **【事実】印刷はブラウザ印刷のみ**。ラベルプリンタの用紙・ドライバ・余白で寸法がずれる既知課題（設計.md §9）。実機検証必須。
5. 【推測】`UFdaily` 由来の VBA ラッパ群と UF検査判定の VBA 共通モジュールが同系統なら、`sqlite_toolkit` の対応表を移植ガイドとして使える。VBA ソースとの突き合わせを推奨。
