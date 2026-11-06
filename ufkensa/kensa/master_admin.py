"""マスタ管理 ── このアプリが読む DB の中身を見る / 直す(python-web-tools の master_admin / master_browse の移植)

【どの DB か】(SOURCES)
- 品質基準.sqlite3・伝送用ファイル.sqlite3(使用パス「品質基準」のフォルダ) … 見る・直す
- 梱包資材マスタ.sqlite3 の「アクセス権限」表だけ(使用パス「アクセス権限マスタ」) … 見る・直す
- 仕掛 SIKALOT / SIKAHIKI / SIKAODR … **見るだけ**(上流が作る DB。ここからは決して書かない)

【python-web-tools と同じ所】
- 行は sqlite3 の暗黙の `rowid` で指す。ただし rowid は付き直ることがあるので、画面が見ていた
  その行の値(`seen`)がいまも同じときだけ書く(確かめと書き込みは 1 つのトランザクション)
- 絞り込みは全部の列を LIKE。並び替えは実在する列だけ(数字は数の大きさで、空は最後)
- 出す行数に上限(ROW_LIMIT)。切ったときは黙らず「ほか N 件」と言う
- 値の形は列の型で確かめる(整数・数・文字)。NOT NULL の列は空にできない

【ここで足したもの】
- 書くときは `BEGIN IMMEDIATE` + busy_timeout(ほかの端末が書いている間は待つ)
- 書いたら trace に「前 → 後」を残す(ev = master.change)。変更履歴の窓はそれを読む
- 書けるのは `auth.master_edit_ok()`(管理者パスワード + アクセス権限の管理者)。判断は画面の入口(routes)で
"""
from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Optional
from urllib.parse import quote as _urlquote

from .. import access_rights, trace
from . import datasource, paths

ROW_KEY = "__rowid"
ROW_LIMIT = 500
BUSY_TIMEOUT_MS = datasource.BUSY_TIMEOUT_MS
CHANGE_EV = "master.change"

STALE_ROW = ("その行は、一覧を出したあとにほかで変わっています(消された・中身が入れ替わった など)。"
             "一覧を出し直してから、もう一度開いてください。")
SIKA_WHY = "仕掛の DB は上流のシステムが作るものなので、ここからは書きません(見るだけ)。"


class MasterError(Exception):
    """断る理由(文言はそのまま画面に出す)。code は画面の出し分け用。"""

    def __init__(self, message: str, code: str = "bad_value", field: str = "", status: int = 400):
        super().__init__(message)
        self.message, self.code, self.field, self.status = message, code, field, status


@dataclass(frozen=True)
class Source:
    key: str
    label: str
    locate: Callable[[], Optional[Path]]
    editable: bool
    tables: tuple[str, ...] = ()      # 空 = ファイルの全部の表
    why: str = ""                     # 直せない理由
    copied: bool = False              # 仕掛(datasource が手元の写しを読む)


def _db(name: str) -> Callable[[], Optional[Path]]:
    return lambda: datasource.resolve(name)


SOURCES: tuple[Source, ...] = (
    Source("quality", "品質基準", _db("品質基準"), True),
    Source("denso", "伝送用ファイル", _db("伝送用ファイル"), True),
    Source("access", "梱包資材マスタ(アクセス権限)", access_rights.master_path, True, ("アクセス権限",)),
    Source("SIKALOTNOW", "仕掛 SIKALOT", _db("SIKALOTNOW"), False, why=SIKA_WHY, copied=True),
    Source("SIKAHIKINOW", "仕掛 SIKAHIKI", _db("SIKAHIKINOW"), False, why=SIKA_WHY, copied=True),
    Source("SIKAODRNOW", "仕掛 SIKAODR", _db("SIKAODRNOW"), False, why=SIKA_WHY, copied=True),
)
BY_KEY = {s.key: s for s in SOURCES}


def _q(name: str) -> str:
    return '"' + str(name).replace('"', '""') + '"'


def source(key: str) -> Source:
    s = BY_KEY.get(key)
    if s is None:
        raise MasterError(f"知らない DB です: {key}", "not_found", status=404)
    return s


def _path(s: Source) -> Path:
    p = s.locate()
    if p is None:
        raise MasterError(f"{s.label} の場所が決まっていません(設定「使用パス」で入れてください)。", "no_source", status=404)
    if not Path(p).is_file():
        raise MasterError(f"{s.label} が見つかりません: {p}", "no_source", status=404)
    return Path(p)


def _open_read(s: Source) -> sqlite3.Connection:
    """読むだけで開く(仕掛は datasource が手元の写しを開く。D26)。"""
    try:
        return datasource.open_ro(_path(s))
    except datasource.DbMissing as e:
        raise MasterError(str(e), "no_source", status=404) from e


def _table_names(conn: sqlite3.Connection, s: Source) -> list[str]:
    names = [r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name")]
    # 「~TMPCLP…」は Access が作る作業用の表(貼り付け・クエリの残り)。取り込みで写っただけなので出さない
    names = [n for n in names if not n.startswith(("~", "MSys"))]
    if s.tables:
        names = [n for n in names if n in s.tables]
    return names


def _check_table(conn: sqlite3.Connection, s: Source, table: str) -> None:
    if table not in _table_names(conn, s):
        raise MasterError(f"{s.label} に表「{table}」がありません。", "not_found", status=404)


# ------------------------------------------------------------------
# 見る
# ------------------------------------------------------------------
def list_sources() -> list[dict[str, Any]]:
    out = []
    for s in SOURCES:
        try:
            p = s.locate()
        except Exception as e:                       # noqa: BLE001 - 場所の設定が壊れていても一覧は出す
            p, err = None, str(e)
        else:
            err = ""
        out.append({"key": s.key, "label": s.label, "path": str(p) if p else "", "exists": bool(p and Path(p).is_file()),
                    "editable": s.editable, "why": s.why, "error": err})
    return out


def tables(key: str) -> dict[str, Any]:
    s = source(key)
    with _open_read(s) as conn:
        out = []
        for n in _table_names(conn, s):
            try:
                count = conn.execute(f"SELECT COUNT(*) FROM {_q(n)}").fetchone()[0]
            except sqlite3.Error:
                count = None
            out.append({"name": n, "rows": count})
    missing = [t for t in s.tables if t not in {x["name"] for x in out}]
    note = f"{s.label} に表「{'・'.join(missing)}」がありません。" if missing else ""
    return {"source": key, "label": s.label, "path": str(_path(s)), "editable": s.editable, "why": s.why,
            "tables": out, "note": note}


def _kind(decl: str) -> str:
    """列の型(SQLite の型の決め方と同じ順)→ int / real / text。"""
    d = (decl or "").upper()
    if "INT" in d:
        return "int"
    if any(x in d for x in ("CHAR", "CLOB", "TEXT")):
        return "text"
    if any(x in d for x in ("REAL", "FLOA", "DOUB", "NUM", "DEC")):
        return "real"
    return "text"


KIND_LABEL = {"int": "整数", "real": "数", "text": "文字"}

#: Access から移した表は列に型が無い(文字)。VBA が Val で数として読む列は、ここで数だけを許す
#: (数でない文字を入れると Val が 0 にして、黙って別の行に当たる)。空は許す
NUMBER_TEXT: dict[tuple[str, str], frozenset[str]] = {
    ("denso", "測定機器ID"): frozenset({"厚下", "厚上"}),
}


def _hint(key: str, table: str, cols: list[dict[str, Any]]) -> list[dict[str, Any]]:
    nums = NUMBER_TEXT.get((key, table), frozenset())
    for c in cols:
        if c["name"] in nums and c["kind"] == "text":
            c["number_text"] = True
            c["kind_label"] = "数(文字で保存)"
    return cols


def columns(conn: sqlite3.Connection, table: str) -> list[dict[str, Any]]:
    out = []
    for cid, name, decl, notnull, default, pk in conn.execute(f"PRAGMA table_info({_q(table)})"):
        kind = _kind(decl)
        auto = bool(pk) and kind == "int" and "INT" in (decl or "").upper()
        out.append({"name": name, "decl": decl or "", "kind": kind, "kind_label": KIND_LABEL[kind],
                    "required": bool(notnull) and default is None and not auto, "pk": bool(pk), "auto": auto})
    return out


def _order(names: list[str], sort: str, sort_dir: str) -> tuple[str, str]:
    """見出しの並び替え(python-web-tools master_browse._order と同じ)。実在しない列は既定(rowid)へ黙って戻す。"""
    if sort and sort in names:
        direction = "DESC" if sort_dir == "desc" else "ASC"
        col = _q(sort)
        text = f"TRIM(CAST({col} AS TEXT))"
        unsigned = f"LTRIM({text}, '+-')"
        is_number = (f"(typeof({col}) IN ('integer', 'real') OR ("
                     f"{unsigned} GLOB '[0-9]*' AND {unsigned} NOT GLOB '*[^0-9.]*'"
                     f" AND LENGTH({text}) - LENGTH(REPLACE({text}, '.', '')) <= 1"
                     f" AND LENGTH({text}) - LENGTH({unsigned}) <= 1))")
        blank = f"({col} IS NULL OR {text} = '')"
        return (f"ORDER BY {blank} ASC, {is_number} DESC,"
                f" CASE WHEN {is_number} THEN CAST({text} AS REAL) END {direction},"
                f" {col} {direction}, rowid ASC"), sort
    return "ORDER BY rowid ASC", ""


def _filter(names: list[str], query: str) -> tuple[str, list[Any]]:
    """絞り込み。どの列でもよい(全部の列を見る)。"""
    text = (query or "").strip()
    if not text:
        return "", []
    conds = " OR ".join(f"CAST({_q(n)} AS TEXT) LIKE ?" for n in names)
    return f" WHERE ({conds})", [f"%{text}%"] * len(names)


def _json_value(v: Any) -> Any:
    if isinstance(v, bytes):
        return "(バイナリ)"
    return v


def page(key: str, table: str, *, query: str = "", sort: str = "", sort_dir: str = "asc",
         limit: int = ROW_LIMIT) -> dict[str, Any]:
    """表の 1 ページぶん。**ファイルから直に読む**(直したあと本当に入ったかをここで確かめられる)。"""
    s = source(key)
    with _open_read(s) as conn:
        _check_table(conn, s, table)
        cols = _hint(key, table, columns(conn, table))
        names = [c["name"] for c in cols]
        where, params = _filter(names, query)
        order, sort_col = _order(names, sort, sort_dir)
        total = conn.execute(f"SELECT COUNT(*) FROM {_q(table)}{where}", params).fetchone()[0]
        cur = conn.execute(f'SELECT rowid AS "{ROW_KEY}", * FROM {_q(table)}{where} {order} LIMIT ?',
                           [*params, max(1, int(limit))])
        heads = [d[0] for d in cur.description]
        rows = [{h: _json_value(v) for h, v in zip(heads, r)} for r in cur.fetchall()]
    note = ""
    hidden = total - len(rows)
    if hidden > 0:
        note = (f"{total}件のうち {len(rows)}件を出しています(ほか {hidden}件)。絞り込むと目当ての行が出ます。"
                "見出しを押すと、出ていない行も含めた全件で並び替えます。")
    return {"source": key, "label": s.label, "table": table, "columns": cols, "rows": rows, "total": total,
            "note": note, "sort": sort_col, "sort_dir": "desc" if sort_dir == "desc" else "asc",
            "editable": s.editable, "why": s.why, "query": (query or "").strip()}


# ------------------------------------------------------------------
# 直す・足す・消す
# ------------------------------------------------------------------
def _clean(cols: list[dict[str, Any]], values: dict[str, Any], *, filling: bool) -> dict[str, Any]:
    """画面から来た値を、列の型に合わせる。filling(足す)なら送られてこない必須の列も見る。"""
    out: dict[str, Any] = {}
    known = {c["name"] for c in cols}
    unknown = [k for k in values if k not in known and k != ROW_KEY]
    if unknown:
        raise MasterError(f"表に無い列です: {'・'.join(unknown)}(一覧を出し直してください)", field=unknown[0])
    for c in cols:
        name = c["name"]
        if name not in values:
            if filling and c["required"]:
                raise MasterError(f"「{name}」は空にできません。", field=name)
            continue
        raw = "" if values[name] is None else str(values[name]).strip()
        if raw == "":
            if c["auto"] and filling:
                continue                                   # INTEGER PRIMARY KEY: 番号は SQLite が振る
            if c["required"]:
                raise MasterError(f"「{name}」は空にできません。", field=name)
            out[name] = None
            continue
        if c["kind"] == "int":
            try:
                f = float(raw)
                if f != int(f):
                    raise ValueError
                out[name] = int(f)
            except (ValueError, OverflowError):
                raise MasterError(f"「{name}」は整数で入れてください(いま: {raw})。", field=name) from None
        elif c["kind"] == "real":
            try:
                out[name] = float(raw)
            except ValueError:
                raise MasterError(f"「{name}」は数で入れてください(いま: {raw})。", field=name) from None
        else:
            if c.get("number_text"):
                try:
                    float(raw)
                except ValueError:
                    raise MasterError(f"「{name}」は数で入れてください(いま: {raw})。検査画面は数として読みます。",
                                      field=name) from None
            out[name] = raw
    return out


def _same_value(now: Any, was: Any) -> bool:
    """python-web-tools master_admin._same_value。1 と 1.0 は同じと見る(JSON で区別が消える)。"""
    if now is None or was is None:
        return (now is None or now == "") and (was is None or was == "")
    if isinstance(now, (int, float)) and isinstance(was, (int, float)) and not isinstance(was, bool):
        return float(now) == float(was)
    return str(now) == str(was)


def _same_row(now: dict[str, Any], seen: dict[str, Any]) -> bool:
    return all(_same_value(now[k], v) for k, v in seen.items() if k in now)


def _connect_rw(path: Path) -> sqlite3.Connection:
    """書くために開く。**無いファイルは作らない**(mode=rw)。BEGIN は自分で出す(isolation_level=None)。"""
    uri = datasource.sqlite_uri(path.resolve(), "rw")
    try:
        conn = sqlite3.connect(uri, uri=True, timeout=BUSY_TIMEOUT_MS / 1000, isolation_level=None)
    except sqlite3.Error as e:
        raise MasterError(f"書くために開けません: {path}({e})", "write_failed", status=409) from e
    conn.execute(f"PRAGMA busy_timeout = {BUSY_TIMEOUT_MS}")
    return conn


def _row(conn: sqlite3.Connection, table: str, rowid: Any) -> Optional[dict[str, Any]]:
    cur = conn.execute(f"SELECT * FROM {_q(table)} WHERE rowid = ?", [rowid])
    r = cur.fetchone()
    return None if r is None else {d[0]: _json_value(v) for d, v in zip(cur.description, r)}


def _write(key: str, table: str, action: str, work: Callable[[sqlite3.Connection, list[dict[str, Any]]], dict[str, Any]]
           ) -> dict[str, Any]:
    """書く 1 回ぶん: BEGIN IMMEDIATE → work → COMMIT。失敗したら ROLLBACK。前後を trace に残す。"""
    s = source(key)
    if not s.editable:
        raise MasterError(s.why or "直せない DB です。", "not_editable", status=403)
    if s.tables and table not in s.tables:
        raise MasterError(f"{s.label} では表「{table}」は直せません。", "not_editable", status=403)
    path = _path(s)
    with trace.timed("db.save", f"マスタ管理: {s.label}/{table} を{action}", expected=(MasterError,),
                     path=str(path), table=table, action=action) as info:
        conn = _connect_rw(path)
        try:
            conn.execute("BEGIN IMMEDIATE")
            try:
                if table not in [r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")]:
                    raise MasterError(f"表「{table}」がありません(ほかで消されました)。一覧を出し直してください。",
                                      "not_found", status=404)
                out = work(conn, _hint(key, table, columns(conn, table)))
                conn.execute("COMMIT")
            except BaseException:
                conn.execute("ROLLBACK")
                raise
        except sqlite3.OperationalError as e:
            raise MasterError(f"書けませんでした({e})。ほかの端末が使っていれば、少し待ってもう一度。",
                              "write_failed", status=409) from e
        finally:
            conn.close()
        info["rowid"] = out.get("rowid")
    trace.note_used(path, "save")
    trace.event(CHANGE_EV, f"マスタ管理: {s.label}/{table} を{action}", data={
        "source": key, "label": s.label, "table": table, "action": action, "rowid": out.get("rowid"),
        "before": out.get("before"), "after": out.get("after")})
    _after_write(key)
    return out


def _after_write(key: str) -> None:
    """読む側の覚えを捨てる(品質基準の表のキャッシュ・アクセス権限の判定)。"""
    if key == "quality":
        from . import quality
        quality.clear_caches()


def save_row(key: str, table: str, rowid: Any, values: dict[str, Any], seen: Optional[dict[str, Any]] = None
             ) -> dict[str, Any]:
    """1 行を書き換える。送られてきた列だけ触る。`seen` = 画面が見ていたその行。"""
    def work(conn: sqlite3.Connection, cols: list[dict[str, Any]]) -> dict[str, Any]:
        clean = _clean(cols, {k: v for k, v in values.items() if k != ROW_KEY}, filling=False)
        before = _row(conn, table, rowid)
        if before is None or (seen is not None and not _same_row(before, seen)):
            raise MasterError(STALE_ROW, "stale", status=409)
        clean = {k: v for k, v in clean.items() if not _same_value(before.get(k), v)}
        if not clean:
            raise MasterError("変わった値がありません。", "no_change")
        sets = ", ".join(f"{_q(k)} = ?" for k in clean)
        conn.execute(f"UPDATE {_q(table)} SET {sets} WHERE rowid = ?", [*clean.values(), rowid])
        return {"rowid": rowid, "before": {k: before.get(k) for k in clean}, "after": clean,
                "message": f"{table} の 1 行を直しました({'・'.join(clean)})"}
    return _write(key, table, "直す", work)


def add_row(key: str, table: str, values: dict[str, Any]) -> dict[str, Any]:
    def work(conn: sqlite3.Connection, cols: list[dict[str, Any]]) -> dict[str, Any]:
        clean = _clean(cols, {k: v for k, v in values.items() if k != ROW_KEY}, filling=True)
        clean = {k: v for k, v in clean.items() if v is not None}
        if not clean:
            raise MasterError("入れる値がありません。")
        names = ", ".join(_q(k) for k in clean)
        try:
            cur = conn.execute(f"INSERT INTO {_q(table)} ({names}) VALUES ({', '.join('?' * len(clean))})",
                               list(clean.values()))
        except sqlite3.IntegrityError as e:
            raise MasterError(f"足せませんでした(同じ値の行がある・必須の列が空 など): {e}") from e
        return {"rowid": cur.lastrowid, "before": None, "after": clean, "message": f"{table} に 1 行足しました"}
    return _write(key, table, "足す", work)


def delete_row(key: str, table: str, rowid: Any, seen: Optional[dict[str, Any]] = None) -> dict[str, Any]:
    def work(conn: sqlite3.Connection, cols: list[dict[str, Any]]) -> dict[str, Any]:
        before = _row(conn, table, rowid)
        if before is None or (seen is not None and not _same_row(before, seen)):
            raise MasterError(STALE_ROW, "stale", status=409)
        conn.execute(f"DELETE FROM {_q(table)} WHERE rowid = ?", [rowid])
        return {"rowid": rowid, "before": before, "after": None, "message": f"{table} の 1 行を消しました"}
    return _write(key, table, "消す", work)


# ------------------------------------------------------------------
# 変更履歴(trace の master.change を新しい順に)
# ------------------------------------------------------------------
def history(*, source_key: str = "", table: str = "", days: int = 30, limit: int = 100) -> list[dict[str, Any]]:
    out = []
    for r in trace.read(days, q=CHANGE_EV, limit=5000):
        if r.get("ev") != CHANGE_EV:
            continue
        d = r.get("data") or {}
        if not isinstance(d, dict):
            continue
        if source_key and d.get("source") != source_key:
            continue
        if table and d.get("table") != table:
            continue
        out.append({"t": r.get("t", ""), "user": r.get("user", ""), "pc": r.get("pc", ""), "label": d.get("label", ""),
                    "table": d.get("table", ""), "action": d.get("action", ""), "rowid": d.get("rowid"),
                    "before": d.get("before"), "after": d.get("after")})
        if len(out) >= limit:
            break
    return out
