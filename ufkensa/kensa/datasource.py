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


def open_ro(db: Union[str, Path]) -> sqlite3.Connection:
    p = resolve(db)
    if not p.is_file():
        raise DbMissing(f"DB ファイルが見つかりません: {p}")
    uri = "file:" + _urlquote(str(p.resolve()).replace("\\", "/"), safe="/:") + "?mode=ro"
    try:
        conn = sqlite3.connect(uri, uri=True, timeout=BUSY_TIMEOUT_MS / 1000)
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


def get_records(db: Union[str, Path], table: str, sql_filter: str = "") -> Optional[list[list[str]]]:
    """VBA `GetRecordsArr`。0 件なら None。"""
    with open_ro(db) as conn:
        cols, rows = read_table(conn, table)
    out = _filtered(cols, rows, sql_filter)
    return out or None


def get_fields(db: Union[str, Path], table: str, sql_filter: str = "") -> Optional[list[str]]:
    """VBA `GetFieldsArr`。**表が空・該当なしでも None**(VBA の EOF チェックがフィールド取得より先)。"""
    with open_ro(db) as conn:
        cols, rows = read_table(conn, table)
    return cols if _filtered(cols, rows, sql_filter) else None


def get_table(db: Union[str, Path], table: str, sql_filter: str = "") -> tuple[list[str], list[list[str]]]:
    """列名と行を一度に(列名は 0 件でも返す。呼び手が VBA の Empty を再現するときは上の2つを使う)。"""
    with open_ro(db) as conn:
        cols, rows = read_table(conn, table)
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
    try:
        return resolve(db).is_file()
    except OSError:
        return False
