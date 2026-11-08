"""使用パスを変えたあとの読み直し(2026-11-08)。

読んだ DB・表・公差表・仕掛の写しは「ファイルの場所と更新時刻」で覚えているので、使用パスを変えれば次に読むときは
新しい場所から読む。ただし**すでに画面に取り込んだ値**(検査中のロットの公差・規格など)は、ロットを読み直すまで
前の場所のままになる。

- clear_caches(): 覚えているものを全部捨てる(使用パスを保存したとき・「読み込む」で呼ぶ)
- reload_now(): 捨てたうえで、検査中のロットがあれば読み直す(設定 → 使用パス の「読み込む」)
"""
from __future__ import annotations

from typing import Any

from .. import trace
from ..logging_utils import get_logger

log = get_logger("kensa.reload_data")


def clear_caches() -> None:
    from . import datasource, judge, lot_copy, quality
    datasource._table_cache.clear()                 # noqa: SLF001 - 表の読み込み
    with judge._tol_lock:                           # noqa: SLF001 - 公差表(対角・ラテラル)
        judge._tol_cache.clear()                    # noqa: SLF001
    for c in (quality._gaichu_cache, quality._gauge_cache, quality._pack_cache):   # noqa: SLF001
        c.clear()
    with lot_copy._lock:                            # noqa: SLF001 - 仕掛の写し(次に読むとき元を確かめ直す)
        lot_copy._checked.clear()                   # noqa: SLF001
    trace.event("settings.reload", "読んだ DB・表・公差表の覚えを捨てました")


def reload_now() -> dict[str, Any]:
    """覚えを捨て、検査中のロットがあれば読み直す。返り値: {"lot": …, "messages": [...]}。"""
    clear_caches()
    from . import session
    s = session.get()
    out: dict[str, Any] = {"lot": "", "messages": []}
    if not getattr(s, "started", False) or s.state.lot == "":
        out["messages"].append({"text": "読み直しました(検査中のロットはありません)", "level": "info"})
        return out
    out["lot"] = s.state.lot
    r = s.reload_lot()
    if r.stop is not None:
        out["messages"].append({"text": r.stop.text, "level": "error"})
    out["messages"] += [m.to_dict() for m in r.messages if m.text]
    return out
