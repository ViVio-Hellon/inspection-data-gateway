"""DB の読み方(VBA `【UFKen_15】SQL_追加` の GetRecordsArr / GetFieldsArr を SQLite で)

VBA は次のように読んでいた(SQL_追加 L292-300, L400-470):
1. `SELECT * FROM 表` を全件取る(ORDER BY なし = Access の返却順)
2. `Recordset.Filter` で絞る(→ `ado_filter`)
3. 1 件も無ければ Empty を返す(GetFieldsArr も Empty。「表が空」と「該当なし」は区別しない)
4. 値は `GetString` で**すべて文字列**。Null は ""

ここでも同じ形で返す:
- 行の順は取り込んだ順(`rowid` の昇順)。Access から取り込むときに元の順で入れる(importer)
- 値は文字列。整数はそのまま、実数は VBA の CStr と同じ有効 15 桁、None は ""
- 該当が無ければ `None`(VBA の Empty)

【ファイルが無いとき】VBA は `End`(全部リセット)。D10 で「止めるが状態は保つ」に決まったので、
`DbMissing` を投げ、画面は状態を保ったままメッセージを出す。
**存在しないパスに sqlite3.connect すると空の DB を作ってしまう**ので、必ず読み取り専用(`mode=ro`)で開く。
"""
from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Iterable, Optional, Union
from urllib.parse import quote as _urlquote

from . import ado_filter, paths, vbx
from ..logging_utils import get_logger

log = get_logger("kensa.datasource")

BUSY_TIMEOUT_MS = 10_000


class DbMissing(RuntimeError):
    """DB ファイルが無い・開けない。文言はそのまま画面に出す。"""


class TableMissing(RuntimeError):
    pass


def resolve(db: Union[str, Path]) -> Path:
    """`"品質基準"` のような名前、または Windows のパス文字列 → この OS の Path。"""
    if isinstance(db, Path):
        return db
    if db in paths.DB_FILES:
        return paths.to_local(paths.db_file(db))
    return paths.to_local(db)


class _ClosingConnection(sqlite3.Connection):
    """`with open_ro(...) as conn:` を抜けたら閉じる接続。

    sqlite3 の標準の `with` は commit / rollback だけで**閉じない**。閉じないと Windows ではファイルを掴んだまま
    になり、仕掛 DB の手元の写しを差し替えられない・共有フォルダのファイルが使用中になる。
    """

    def __exit__(self, *exc):  # type: ignore[override]
        try:
            return super().__exit__(*exc)
        finally:
            self.close()


def sqlite_uri(path: Union[str, Path], mode: str = "ro") -> str:
    """SQLite の URI。共有フォルダ(\\\\server\\share)は file:////server/share にする
    (file://server は「authority」と読まれて「invalid uri authority」で開けない)。"""
    text = str(path).replace("\\", "/")
    if text.startswith("//"):
        text = "//" + text
    return "file:" + _urlquote(text, safe="/:") + "?mode=" + mode


def open_ro(db: Union[str, Path]) -> sqlite3.Connection:
    p = resolve(db)
    if not p.is_file():
        raise DbMissing(f"DB ファイルが見つかりません: {p}")
    if paths.is_copied(p):
        # D26: 仕掛 DB は共有を直接開かず、手元へ写した写しを読む。写せなければ読まない
        from . import lot_copy
        c = lot_copy.local_copy(str(p))
        if c is None:
            raise DbMissing(f"DB ファイルを手元へ写せないので読みません: {p}")
        p = c
    uri = sqlite_uri(p.resolve(), "ro")
    try:
        conn = sqlite3.connect(uri, uri=True, timeout=BUSY_TIMEOUT_MS / 1000, factory=_ClosingConnection)
    except sqlite3.Error as e:
        raise DbMissing(f"DB を開けません: {p}({e})") from e
    conn.execute(f"PRAGMA busy_timeout = {BUSY_TIMEOUT_MS}")
    return conn


def as_text(value: object) -> str:
    """ADO `GetString` が書く文字列。"""
    if value is None:
        return ""
    if isinstance(value, bytes):
        return ""
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return vbx.cstr(value)
    return str(value)


def _q(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def read_table(conn: sqlite3.Connection, table: str) -> tuple[list[str], list[list[str]]]:
    """表を取り込み順で全部読む(列名, 行)。読込先・件数・処理時間をログに残す(ログの決まり)。"""
    from .. import trace
    path = _db_path(conn)
    trace.note_used(path, "read")
    with trace.timed("db.read", f"{table} を読む", path=path, table=table) as info:
        try:
            cur = conn.execute(f"SELECT * FROM {_q(table)} ORDER BY rowid")
        except sqlite3.OperationalError as e:
            if "no such table" in str(e):
                raise TableMissing(f"表 {table} がありません") from e
            raise
        cols = [d[0] for d in cur.description]
        rows = [[as_text(v) for v in r] for r in cur.fetchall()]
        info["rows"] = len(rows)
    return cols, rows


def _db_path(conn: sqlite3.Connection) -> str:
    try:
        return str(conn.execute("PRAGMA database_list").fetchone()[2])
    except sqlite3.Error:
        return ""


def _filtered(cols: list[str], rows: list[list[str]], expr: str) -> list[list[str]]:
    if not expr:
        return rows
    f = ado_filter.compile_filter(expr)
    return [r for r in rows if f(dict(zip(cols, r)))]


#: 読んだ表の控え。(実際に読んだファイル, 表) → (ファイルの印, 列名, 行)。ファイルが変われば読み直す。
#: ロット登録 1 回で同じ表(仕掛台帳・品質基準など)を 40 回以上読んでいたので、変わっていなければ読み直さない
#: (2026-11-06 重さの見直し)。行は呼び手が変えても控えが壊れないよう、写しを返す
_table_cache: dict[tuple[str, str], tuple[tuple[int, int], list[str], list[list[str]]]] = {}
_TABLE_CACHE_MAX = 24


def _actual_file(db: Union[str, Path]) -> Optional[Path]:
    """実際に開くファイル(仕掛台帳は手元の写し)。無ければ None。"""
    try:
        p = resolve(db)
        if paths.is_copied(p):
            from . import lot_copy
            return lot_copy.local_copy(str(p))
        return p if p.is_file() else None
    except OSError:
        return None


def _read_uncached(db: Union[str, Path], table: str) -> tuple[list[str], list[list[str]]]:
    with open_ro(db) as conn:
        return read_table(conn, table)


def _read(db: Union[str, Path], table: str) -> tuple[list[str], list[list[str]]]:
    """表を読む。仕掛台帳(手元へ写して読むファイル)は、列名の書き方の違い(2つ目の名前のファイル。
    半角/全角カナなど)を第1の名前に直す(`lot_columns.canonical`)。"""
    actual = _actual_file(db)
    stamp = None
    if actual is not None:
        try:
            st = actual.stat()
            stamp = (st.st_size, st.st_mtime_ns)
        except OSError:
            stamp = None
    key = (str(actual), table)
    hit = _table_cache.get(key) if stamp is not None else None
    if hit is not None and hit[0] == stamp:
        cols, rows = list(hit[1]), [list(r) for r in hit[2]]
    else:
        cols, rows = _read_uncached(db, table)
        if stamp is not None:
            if len(_table_cache) >= _TABLE_CACHE_MAX:
                _table_cache.pop(next(iter(_table_cache)))
            _table_cache[key] = (stamp, list(cols), [list(r) for r in rows])
    try:
        lot = paths.is_copied(resolve(db))
    except (OSError, ValueError):
        lot = False
    if lot:
        from . import lot_columns
        cols = lot_columns.canonical(cols)
        n0 = len(cols)
        cols, rows = lot_columns.pad(cols, rows)
        if len(cols) > n0:
            rows = fill_from_first(db, cols, rows, cols[n0:])
    return cols, rows


#: 第1の仕掛台帳の「ロットの鍵 → 行」(ファイルの場所と更新で覚える)
_first_index: dict = {}


def fill_from_first(db: Union[str, Path], cols: list[str], rows: list, added: list[str]) -> list:
    """2026-11-09(ご依頼): 第2(控え)の仕掛台帳に無い列(公差・品質ｸﾞﾚｰﾄﾞなど)を、第1の同じロットの行で補う。

    第2は列が少ない(SIKALOT の控えは 46 列)。無い列は空として足していたので、公差が 0・歪等級が空になっていた。
    第1が読めない・第1にそのロットが無いときは空のまま。第1そのものを読んでいるときは何もしない。
    鍵は ﾛｯﾄ番号 + 検査番号 + ｵｰﾀﾞｰ番号(合わなければ ﾛｯﾄ番号 だけ)。SIKALOT だけ。
    """
    from . import lot_columns
    if lot_columns.which_db(cols) != "SIKALOTNOW" or not rows:
        return rows
    try:
        base = paths.slot_value(paths.LOT, paths.FIRST)
        first_text = paths.join(base, paths.first_name("SIKALOTNOW"))
        if resolve(db).resolve() == resolve(first_text).resolve():
            return rows
        actual = _actual_file(first_text)                # 第1も手元へ写してから読む(共有を直接読まない)
        if actual is None:
            return rows
        st = actual.stat()
        stamp = (str(actual), st.st_size, st.st_mtime_ns)
        hit = _first_index.get("SIKALOTNOW")
        if hit is None or hit[0] != stamp:
            fcols, frows = _read(first_text, "仕掛")    # 第1は列がそろっているので、ここで補う処理は走らない
            idx = {c: i for i, c in enumerate(fcols)}
            full, by_lot = {}, {}
            for r in frows:
                k = tuple(str(r[idx[c]] or "").strip() if c in idx and idx[c] < len(r) else ""
                          for c in ("ﾛｯﾄ番号", "検査番号", "ｵｰﾀﾞｰ番号"))
                full.setdefault(k, r)
                by_lot.setdefault(k[0], r)
            hit = (stamp, idx, full, by_lot)
            _first_index["SIKALOTNOW"] = hit
    except (OSError, sqlite3.Error, DbMissing, TableMissing) as e:
        log.info("第1の仕掛台帳で補えませんでした: %s", e)
        return rows
    _stamp, idx, full, by_lot = hit
    me = {c: i for i, c in enumerate(cols)}
    take = [(me[c], idx[c]) for c in added if c in idx]
    if not take:
        return rows
    out = []
    n = 0
    for r in rows:
        k = tuple(str(r[me[c]] or "").strip() if c in me and me[c] < len(r) else ""
                  for c in ("ﾛｯﾄ番号", "検査番号", "ｵｰﾀﾞｰ番号"))
        src = full.get(k) or by_lot.get(k[0])
        if src is not None:
            r = list(r)
            for i, j in take:
                if j < len(src) and (r[i] is None or r[i] == ""):
                    r[i] = "" if src[j] is None else as_text(src[j])
            n += 1
        out.append(r)
    log.info("第2の仕掛台帳に無い列 %d 個を、第1の同じロットで補いました(%d 行)", len(take), n)
    return out


def get_records(db: Union[str, Path], table: str, sql_filter: str = "") -> Optional[list[list[str]]]:
    """VBA `GetRecordsArr`。0 件なら None。"""
    cols, rows = _read(db, table)
    out = _filtered(cols, rows, sql_filter)
    return out or None


def get_fields(db: Union[str, Path], table: str, sql_filter: str = "") -> Optional[list[str]]:
    """VBA `GetFieldsArr`。**表が空・該当なしでも None**(VBA の EOF チェックがフィールド取得より先)。"""
    cols, rows = _read(db, table)
    return cols if _filtered(cols, rows, sql_filter) else None


def get_table(db: Union[str, Path], table: str, sql_filter: str = "") -> tuple[list[str], list[list[str]]]:
    """列名と行を一度に(列名は 0 件でも返す。呼び手が VBA の Empty を再現するときは上の2つを使う)。"""
    cols, rows = _read(db, table)
    return cols, _filtered(cols, rows, sql_filter)


def column_index(fields: Iterable[str], name: str) -> int:
    """VBA の `For i = 1 To UBound(F, 2): Select Case F(1, i): Case "列名": C_列 = i` に当たる(0 始まり)。

    見つからなければ -1(VBA の `C_列 = 0` のまま、に当たる)。
    """
    for i, f in enumerate(fields):
        if f == name:
            return i
    return -1


def exists(db: Union[str, Path]) -> bool:
    """ファイルがあるか。仕掛 DB は手元へ写せたか(写せなければ読まないので「無い」。D26)。"""
    try:
        p = resolve(db)
        if paths.is_copied(p):
            return p.is_file() and paths.lot_usable(str(p))
        return p.is_file()
    except OSError:
        return False
