"""Aインプットのライン → アイカス(NIS)のコード(2026-11-14 のご依頼)。伝送用ファイル.sqlite3 の表「アイカスラインコード」。

VBA は Input.LinConv(Input:1320-1333)にライン名とコードを直書きしていた。ここでは共有の表で持ち、
保守 → アイカスのライン で直せるようにする(全 PC で同じになるよう 伝送用ファイル に置く)。

    列: ライン(VBA の LineName)・コード・説明

- **既定は VBA と同じ**(L-1=05、LVC=03、HVC=04、ﾊﾞﾗﾝｻｰ/ﾄｯﾄ=01、LS4=06、NS1=07、Test=AA、ほか = Maru)
- 「(その他)」の行が、表に無いラインのコード(VBA の Case Else = "Maru")
- 検査中に読むときは書かない(表が無ければ既定)。読めないときは最後に読めた値(無ければ既定)
"""
from __future__ import annotations

import sqlite3
import threading
import time
from pathlib import Path
from typing import Any, Optional

from .. import safe_io
from ..logging_utils import get_logger
from . import datasource

log = get_logger("kensa.icas_lines")

TABLE = "アイカスラインコード"
DB = "伝送用ファイル"
OTHER = "(その他)"
CHANGE_EV = "icas_lines.change"
TTL_SEC = 5.0
MAX_LEN = 4                      # 電文は固定長。VBA の "Maru" が 4 文字なので 4 まで

#: VBA の LinConv(並びもそのまま)。(ライン, コード, 説明)
DEFAULTS: tuple[tuple[str, str, str], ...] = (
    ("L-1", "05", ""), ("LVC", "03", ""), ("HVC", "04", ""), ("ﾊﾞﾗﾝｻｰ", "01", "大板"), ("ﾄｯﾄ", "01", "大板"),
    ("LS4", "06", ""), ("NS1", "07", ""), ("Test", "AA", "試験"), (OTHER, "Maru", "表に無いライン(丸徳など)"),
)

_lock = threading.Lock()
_last: dict[str, tuple[float, list[tuple[str, str, str]]]] = {}


class Invalid(ValueError):
    def __init__(self, line: str, message: str):
        super().__init__(message)
        self.line = line


def _db_path() -> Path:
    return datasource.resolve(DB)


def _has_table(con: sqlite3.Connection) -> bool:
    return con.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (TABLE,)).fetchone() is not None


def check(rows: list[tuple[str, str, str]]) -> list[tuple[str, str, str]]:
    """確かめて整える。ラインは空・重複なし、コードは 1〜4 文字の英数字。(その他)は必ず 1 行。"""
    out: list[tuple[str, str, str]] = []
    seen: set[str] = set()
    for line, code, note in rows:
        line, code, note = str(line or "").strip(), str(code or "").strip(), str(note or "").strip()
        if not line and not code:
            continue
        if not line:
            raise Invalid(line, f"コード {code} のラインが空です")
        if line in seen:
            raise Invalid(line, f"ライン {line} が 2 回あります")
        if not code or len(code) > MAX_LEN or not code.isascii() or not code.isalnum():
            raise Invalid(line, f"{line} のコードは 1〜{MAX_LEN} 文字の英数字にしてください(いま「{code}」)")
        seen.add(line)
        out.append((line, code, note))
    if OTHER not in seen:
        out.append(next(r for r in DEFAULTS if r[0] == OTHER))
    return out


def _read(p: Path) -> Optional[list[tuple[str, str, str]]]:
    try:
        if not p.is_file():
            return None
        con = safe_io.connect_existing(p, timeout=5)
        try:
            if not _has_table(con):
                return list(DEFAULTS)
            recs = con.execute(f'SELECT "ライン", "コード", "説明" FROM "{TABLE}" ORDER BY rowid').fetchall()
        finally:
            con.close()
    except (OSError, sqlite3.Error):
        return None
    try:
        return check([(r[0], r[1], r[2]) for r in recs])
    except Invalid as e:
        log.warning("アイカスラインコードの表が悪いので既定を使います: %s", e)
        return list(DEFAULTS)


def rows(path: Optional[Path] = None, *, fresh: bool = False) -> list[tuple[str, str, str]]:
    p = path or _db_path()
    key, now = str(p), time.monotonic()
    with _lock:
        hit = _last.get(key)
    if hit and not fresh and now - hit[0] < TTL_SEC:
        return list(hit[1])
    got = _read(p)
    if got is None:
        return list(hit[1]) if hit else list(DEFAULTS)
    with _lock:
        _last[key] = (now, got)
    return list(got)


def code_for(line: str) -> str:
    """Input.LinConv。空のラインは ""、表に無いラインは (その他) のコード。"""
    if line == "":
        return ""
    table = rows()
    for ln, code, _ in table:
        if ln == line:
            return code
    return next((c for ln, c, _ in table if ln == OTHER), "Maru")


def ensure_table(path: Optional[Path] = None) -> bool:
    """表が無ければ作って VBA の既定を入れる(保守からだけ呼ぶ)。"""
    p = path or _db_path()
    if not p.is_file():
        return False
    try:
        con = safe_io.connect_existing(p, timeout=5)
        try:
            with con:
                if not _has_table(con):
                    con.execute(f'CREATE TABLE "{TABLE}" ("ライン" TEXT PRIMARY KEY, "コード" TEXT NOT NULL, '
                                f'"説明" TEXT NOT NULL DEFAULT \'\')')
                    con.executemany(f'INSERT INTO "{TABLE}" VALUES (?, ?, ?)', DEFAULTS)
                    log.info("アイカスラインコードの表を作りました: %s", p)
            return True
        finally:
            con.close()
    except sqlite3.Error as e:
        log.warning("アイカスラインコードの表を作れませんでした(%s): %s", p, e)
        return False


def save(new_rows: list[Any], path: Optional[Path] = None) -> list[tuple[str, str, str]]:
    """全行を置き換える。悪ければ Invalid(何も書かない)。前 → 後 を trace に残す。"""
    p = path or _db_path()
    vals = check([(r.get("line"), r.get("code"), r.get("note")) if isinstance(r, dict) else tuple(r) for r in new_rows])
    before = rows(p, fresh=True)
    if not ensure_table(p):
        raise OSError(f"伝送用ファイルに書けません: {p}")
    con = safe_io.connect_existing(p, timeout=10)
    try:
        with con:
            con.execute(f'DELETE FROM "{TABLE}"')
            con.executemany(f'INSERT INTO "{TABLE}" VALUES (?, ?, ?)', vals)
    finally:
        con.close()
    with _lock:
        _last[str(p)] = (time.monotonic(), list(vals))
    b, a = {r[0]: r[1] for r in before}, {r[0]: r[1] for r in vals}
    changed = {k: [b.get(k, ""), a.get(k, "")] for k in sorted(set(b) | set(a)) if b.get(k) != a.get(k)}
    log.info("アイカスラインコードを保存しました: %s", changed)
    if changed:
        try:
            from .. import trace
            trace.event(CHANGE_EV, "保守: アイカスのラインコードを変更", data={"changed": changed})
        except Exception as e:  # noqa: BLE001 - 履歴が残せなくても保存は済んでいる
            log.warning("アイカスラインコードの変更履歴を残せませんでした: %s", e)
    return vals


def reset_cache() -> None:
    with _lock:
        _last.clear()


def history(days: int = 90, limit: int = 30) -> list[dict[str, Any]]:
    """保守で変えた履歴(trace の icas_lines.change)。新しい順。"""
    from .. import trace
    try:
        trace.flush()
    except Exception:  # noqa: BLE001
        pass
    out = []
    for r in trace.read(days, q=CHANGE_EV, limit=2000):
        if r.get("ev") != CHANGE_EV:
            continue
        ch = (r.get("data") or {}).get("changed") if isinstance(r.get("data"), dict) else None
        if not isinstance(ch, dict):
            continue
        out.append({"t": r.get("t", ""), "user": r.get("user", ""), "pc": r.get("pc", ""),
                    "changes": [{"line": k, "before": v[0], "after": v[1]} for k, v in ch.items()
                                if isinstance(v, list) and len(v) == 2]})
        if len(out) >= limit:
            break
    return out
