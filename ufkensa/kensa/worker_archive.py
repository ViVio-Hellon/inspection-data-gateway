"""測定値ごとの検査員と時刻の控え(2026-11-08)。

WorkerTime_<L>(VBA の Time_Posting と同じシート)は、いま検査中のロットの分しか持たない(ロットが替わると作り直す)。
そこで**印刷して保存するとき**に、そのロットの分を手元DB の表「測定者履歴」に写して残す。履歴画面はここから読む。

    測定者履歴(検査開始終了・ロット・ライン・欄・検査員・時刻・パケット)
    検査開始終了 = 保存用の A 列(履歴画面の行と結びつける鍵)。同じ鍵・ロット・ラインの行は入れ替える(再印刷)
"""
from __future__ import annotations

from typing import Any

from . import grid
from .store import LocalStore, cell_text, line_sheet

TABLE = "測定者履歴"
WORKER_TIME = "WorkerTime"
_COLS = ("検査開始終了", "ロット", "ライン", "欄", "検査員", "時刻", "パケット")


def read_sheet(store: LocalStore, lname: str, lot: str) -> dict[str, dict[str, str]]:
    """WorkerTime_<L> のいまのロットの分(欄の名前 → {"who", "when", "pake"})。A1 が別のロットなら空。"""
    name = line_sheet(WORKER_TIME, lname)
    if lot == "" or not store.exists(name):
        return {}
    cells = {(r, c): v for r, c, v in store.sheet(name).used()}
    if cell_text(cells.get((1, 1))) != lot:
        return {}
    out: dict[str, dict[str, str]] = {}
    for i, p in enumerate(grid.POSITIONS, start=1):
        for x in range(1, grid.ROW_COUNT[p] + 1):
            who, when = cell_text(cells.get((i + 2, x))), cell_text(cells.get((i + 8, x)))
            if who or when:
                out[f"{p}{x}"] = {"who": who, "when": when, "pake": cell_text(cells.get((i + 14, x)))}
    return out


def _ensure(store: LocalStore) -> None:
    store.conn.execute(f'CREATE TABLE IF NOT EXISTS "{TABLE}" (' + ", ".join(f'"{c}" TEXT' for c in _COLS) + ")")


def archive(store: LocalStore, lname: str, lot: str, period: str) -> int:
    """印刷・保存のときに、いまのロットの分を控えに写す。写した欄の数を返す。"""
    data = read_sheet(store, lname, lot)
    with store.transaction():
        with store._write():                                    # noqa: SLF001
            _ensure(store)
            store.conn.execute(f'DELETE FROM "{TABLE}" WHERE "検査開始終了" = ? AND "ロット" = ? AND "ライン" = ?',
                               (period, lot, lname))
            store.conn.executemany(f'INSERT INTO "{TABLE}" VALUES (?, ?, ?, ?, ?, ?, ?)',
                                   [(period, lot, lname, k, w["who"], w["when"], w["pake"]) for k, w in data.items()])
    return len(data)


def lookup(store: LocalStore, period: str, lot: str) -> dict[str, dict[str, Any]]:
    """履歴画面の 1 件分(欄の名前 → {"who", "when", "pake"})。控えが無ければ空。"""
    try:
        rows = store.conn.execute(f'SELECT "欄", "検査員", "時刻", "パケット" FROM "{TABLE}" '
                                  f'WHERE "検査開始終了" = ? AND "ロット" = ?', (period, lot)).fetchall()
    except Exception:                                           # noqa: BLE001 - 表がまだ無い
        return {}
    return {k: {"who": w or "", "when": t or "", "pake": pk or ""} for k, w, t, pk in rows}
