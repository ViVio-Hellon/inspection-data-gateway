"""VBA との突き合わせ(工程 10): 保存用に残っている公差と、Python で登録し直した公差を比べる

保存用(極厚は耳用)の 1 行には、印刷したときに VBA が出した 厚・巾・丈 と その公差(-/+)が残っている
(MainDataSave。3〜11 列 = 厚, 巾, 丈, 厚-, 厚+, 巾-, 巾+, 丈-, 丈+)。
仕掛(SIKALOTNOW)にまだ残っているロットを Python のロット登録(Session.register_lot)にかけ、同じ値になるかを見る。

- 本番の手元DB・測定DB・共有フォルダには書かない。登録し直しは一時フォルダの手元DB で行い、旧ログも書かない
- 公差は登録した日の品質基準で決まるので、品質基準が変わったロットは違って当然。結果の「理由」はこちらで決めつけない
- 設備仕様のチェックはこの端末の設備名(設定 → この端末)で行う。別ラインのブックから取り込んだ行は対象外にできる
"""
from __future__ import annotations

import csv
import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional

from . import datasource, store as st

#: 保存用・耳用の 3〜11 列(MainDataSave の hiki)
FIELDS = ("厚", "巾", "丈", "厚-", "厚+", "巾-", "巾+", "丈-", "丈+")
STATE_ATTRS = ("atu", "hab", "tak", "am", "ap", "hm", "hp", "tm", "tp")


@dataclass
class Row:
    sheet: str
    row: int
    when: str
    lot: str
    saved: tuple[str, ...]


@dataclass
class Outcome:
    row: Row
    status: str                          # 一致 / 違う / 仕掛に無い / 登録で止まった
    python: tuple[str, ...] = ()
    diffs: list[str] = field(default_factory=list)
    note: str = ""


def saved_rows(store: st.LocalStore, *, latest_only: bool = True) -> list[Row]:
    """保存用・耳用の行(ロットが同じなら新しい行だけ)。"""
    out: list[Row] = []
    for name in (st.HOZON, st.MIMI):
        if not store.exists(name):
            continue
        sh = store.sheet(name)
        last = sh.last_row(2)
        if last < 2:
            continue
        for i, r in enumerate(sh.read_range(2, 1, last, 11), start=2):
            lot = st.cell_text(r[1])
            if not lot:
                continue
            out.append(Row(name, i, st.cell_text(r[0]), lot, tuple(st.cell_text(x) for x in r[2:11])))
    if latest_only:
        seen: dict[str, Row] = {}
        for r in out:
            seen[r.lot] = r                  # 下の行ほど新しい
        out = list(seen.values())
    return out


def in_shikake(lot: str) -> bool:
    try:
        return datasource.get_records("SIKALOTNOW", "仕掛", f"[ﾛｯﾄ番号] = '{lot}'") is not None
    except datasource.DbMissing:
        return False


def _same(a: str, b: str) -> bool:
    """文字が同じか、数として同じ(VBA の保存はセルの値なので "2" と "2.0" がありうる)。"""
    if a == b:
        return True
    try:
        return float(a) == float(b)
    except ValueError:
        return False


def compare(rows: Iterable[Row], *, measure_db: Optional[Path] = None,
            work_dir: Optional[Path] = None) -> list[Outcome]:
    """各行のロットを一時の手元DB で登録し直し、公差を比べる。"""
    from . import session
    tmp = Path(work_dir) if work_dir else Path(tempfile.mkdtemp(prefix="ufkensa_verify_"))
    outs: list[Outcome] = []
    try:
        for i, row in enumerate(rows):
            if not in_shikake(row.lot):
                outs.append(Outcome(row, "仕掛に無い"))
                continue
            local = st.LocalStore(tmp / f"verify_{i}.sqlite3")
            try:
                s = session.Session(local, write_logs=False, calibration_skip=lambda: True, measure_db=measure_db)
                s.entry_gate = False          # 突き合わせは検査員・校正なしで登録し直す
                s.start()
                res = s.register_lot(row.lot, {"lot_register": True, "lot_recent": True, "lot_split": False})
                if res.stop is not None:
                    outs.append(Outcome(row, "登録で止まった", note=res.stop.text.replace("\r\n", " ").replace("\n", " ")))
                    continue
                py = tuple(str(getattr(s.state, a)) for a in STATE_ATTRS)
                diffs = [f"{n}: VBA {a} / Python {b}" for n, a, b in zip(FIELDS, row.saved, py) if not _same(a, b)]
                outs.append(Outcome(row, "違う" if diffs else "一致", py, diffs))
            finally:
                local.close()
    finally:
        if work_dir is None:
            shutil.rmtree(tmp, ignore_errors=True)
    return outs


def write_report(outs: list[Outcome], path: Path) -> Path:
    """結果を CSV(cp932・Excel でそのまま開ける)に書く。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="cp932", errors="replace", newline="") as f:
        w = csv.writer(f)
        w.writerow(["シート", "行", "検査開始～終了", "LotNo", "結果", "違い", "メモ"]
                   + [f"VBA {n}" for n in FIELDS] + [f"Python {n}" for n in FIELDS])
        for o in outs:
            r = o.row
            w.writerow([r.sheet, r.row, r.when, r.lot, o.status, " / ".join(o.diffs), o.note]
                       + list(r.saved) + list(o.python or ("",) * len(FIELDS)))
    return path


def summary(outs: list[Outcome]) -> dict[str, int]:
    c: dict[str, int] = {}
    for o in outs:
        c[o.status] = c.get(o.status, 0) + 1
    return c
