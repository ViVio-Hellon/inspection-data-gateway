"""公差外の確認(UFConfi)を出す行のマスタ(2026-11-11 のご依頼)。伝送用ファイル.sqlite3 の表「公差外の確認行」。

VBA は `OutTolerance.Out_Tolerance` の Select Case に行番号を直書きしていた
(薄板 2,4,6,8,10 / スラブ 2,4,6,8,10 / 厚板・少枚数 4,10 / それ以外 4,7,10,12,14)。
ここでは共有の表で持ち、保守 → 公差外の確認 で直せるようにする(全 PC で同じになるよう 伝送用ファイル に置く)。

    列: 区分(薄板・スラブ・厚板・それ以外)・行(「2,4,6,8,10」のようなカンマ区切り)

- 検査中に読むときは **書かない**(表が無ければ VBA と同じ既定)。表は保守で開いたとき・保存したときに作る
- 読めないときは最後に読めた値(無ければ VBA の既定)。画面は止めない
"""
from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Optional

from .. import safe_io
from ..logging_utils import get_logger
from . import cache_registry, datasource

log = get_logger("kensa.confi_rows")

TABLE = "公差外の確認行"
DB = "伝送用ファイル"
COLUMNS = ("区分", "行")
KINDS = ("薄板", "スラブ", "厚板", "それ以外")
LABELS = {"薄板": "薄板", "スラブ": "スラブ", "厚板": "厚板・少枚数", "それ以外": "それ以外"}
MAX_ROW = {"薄板": 10, "スラブ": 10, "厚板": 10, "それ以外": 14}
DEFAULTS: dict[str, tuple[int, ...]] = {
    "薄板": (2, 4, 6, 8, 10),
    "スラブ": (2, 4, 6, 8, 10),
    "厚板": (4, 10),
    "それ以外": (4, 7, 10, 12, 14),
}

_last_known: dict[str, dict[str, tuple[int, ...]]] = {}
cache_registry.register_dict("confi_rows._last_known", _last_known)


class Invalid(ValueError):
    def __init__(self, kind: str, message: str):
        super().__init__(message)
        self.kind = kind


def kind_of(work: str) -> str:
    """作業区分 → 表の区分(Out_Tolerance の分け方と同じ)。"""
    if work in ("薄板", "スラブ"):
        return work
    if work in ("厚板", "少枚数"):
        return "厚板"
    return "それ以外"


def parse(kind: str, text: object) -> tuple[int, ...]:
    """「2, 4,6」→ (2, 4, 6)。範囲外・数字でない・重複は Invalid。空は 確認を出さない(())。"""
    if kind not in KINDS:
        raise Invalid(kind, f"区分が違います: {kind}")
    s = str(text if text is not None else "").replace("、", ",").replace("，", ",").strip()
    if s == "":
        return ()
    out: list[int] = []
    hi = MAX_ROW[kind]
    for part in s.split(","):
        p = part.strip()
        if not p.isdigit():
            raise Invalid(kind, f"{LABELS[kind]}: 「{p or '(空)'}」は行番号ではありません(1〜{hi} をカンマ区切り)")
        n = int(p)
        if not 1 <= n <= hi:
            raise Invalid(kind, f"{LABELS[kind]}: {n} は範囲外です(1〜{hi})")
        if n in out:
            raise Invalid(kind, f"{LABELS[kind]}: {n} が重なっています")
        out.append(n)
    return tuple(sorted(out))


def fmt(rows: tuple[int, ...]) -> str:
    return ",".join(str(n) for n in rows)


def _db_path() -> Path:
    return datasource.resolve(DB)


def ensure_table(path: Optional[Path] = None) -> bool:
    """表が無ければ作って VBA の既定を入れる(保守からだけ呼ぶ)。"""
    p = path or _db_path()
    if not p.is_file():
        return False
    try:
        con = safe_io.connect_existing(p, timeout=5)
        try:
            have = con.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?", (TABLE,)).fetchone()
            if have:
                return True
            with con:
                con.execute(f'CREATE TABLE "{TABLE}" ("区分" TEXT NOT NULL, "行" TEXT NOT NULL)')
                con.executemany(f'INSERT INTO "{TABLE}" VALUES (?, ?)', [(k, fmt(v)) for k, v in DEFAULTS.items()])
            log.info("公差外の確認行のマスタを作りました: %s", p)
            return True
        finally:
            con.close()
    except sqlite3.Error as e:
        log.warning("公差外の確認行のマスタを作れませんでした(%s): %s", p, e)
        return False


def rows_map(path: Optional[Path] = None) -> dict[str, tuple[int, ...]]:
    """区分 → 行。表が無ければ既定(作らない)。読めなければ最後に読めた値。悪い値の区分は既定。"""
    p = path or _db_path()
    key = str(p)
    base = dict(_last_known.get(key) or DEFAULTS)
    try:
        if not p.is_file():
            return base
        con = safe_io.connect_existing(p, timeout=5)      # 表が無いだけなら既定(読むだけ。作らない)
        try:
            have = con.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?", (TABLE,)).fetchone()
        finally:
            con.close()
        if not have:
            _last_known.pop(key, None)
            return dict(DEFAULTS)
        cols, recs = datasource.get_table(str(p), TABLE)
    except datasource.TableMissing:
        return dict(DEFAULTS)
    except (datasource.DbMissing, OSError, sqlite3.Error):
        log.warning("公差外の確認行のマスタを読めないので、最後に読めた値(または既定)を使います: %s", p)
        return base
    idx = {c: i for i, c in enumerate(cols)}
    if not all(c in idx for c in COLUMNS):
        return dict(DEFAULTS)
    got = dict(DEFAULTS)
    for r in recs:
        k = str(r[idx["区分"]]).strip()
        if k not in KINDS:
            continue
        try:
            got[k] = parse(k, r[idx["行"]])
        except Invalid as e:
            log.warning("公差外の確認行の値が悪いので既定を使います: %s", e)
            got[k] = DEFAULTS[k]
    _last_known[key] = got
    return got


def save(values: dict, path: Optional[Path] = None) -> dict[str, tuple[int, ...]]:
    """区分 → 文字(カンマ区切り)を確かめて保存。悪ければ Invalid(何も書かない)。"""
    parsed = {k: parse(k, (values or {}).get(k, fmt(DEFAULTS[k]))) for k in KINDS}
    p = path or _db_path()
    if not ensure_table(p):
        raise OSError(f"伝送用ファイルに書けません: {p}")
    con = safe_io.connect_existing(p, timeout=10)
    try:
        with con:
            con.execute(f'DELETE FROM "{TABLE}"')
            con.executemany(f'INSERT INTO "{TABLE}" VALUES (?, ?)', [(k, fmt(v)) for k, v in parsed.items()])
    finally:
        con.close()
    _last_known[str(p)] = dict(parsed)
    log.info("公差外の確認行を保存しました: %s", {k: fmt(v) for k, v in parsed.items()})
    return parsed
