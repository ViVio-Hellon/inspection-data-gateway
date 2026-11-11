"""照合ルール 第3段(測定機器ID = MachineTransfer.DbReturnMicro / 設備仕様 = QP.CheckEquipmentDb)の特性テストの入力と、
いまの結果の記録(2026-11-11)。

記録は `tests/fixtures/match_rules/characterization_equip.json`。作り直すとき(いまのコードの結果を正とする):

    python -m tests.match_rules_equip_cases --write

測定機器ID(db_return_micro):
- 基準データの表 + 試験用の表(用途の ≠X / *X / X* / 完全一致・カンマ区切り・空の厚の行・用途の列が全部空のライン)
- ライン(VBA の名前・正規の名前・大板・無いライン)× 厚(仕掛の全ロットの厚 + 表の厚下・厚上ちょうど / ±0.01)
  × 用途(仕掛の全ロットの用途 + 表の用途 + 含む・含まない・空)× 返す列(出力テキスト・機器名・ID)、デプス(dep)
- B52(_bounds): 厚下・厚上は Val(Double)、厚は Val(Single)。28.01 などの境界ちょうど
設備仕様(check_equipment_db):
- 基準データの表 + 試験用の表(空の MIN・MAX、空の材質、重なる行、作業可否が空)
- ライン × 仕掛の全ロットの (材質, オーダー厚・幅・丈, 単重) + 表の各行の MIN・MAX ちょうど / ±0.01・±1、
  条件を 1 つずつ外した組
"""
from __future__ import annotations

import json
import sqlite3
import sys
from pathlib import Path
from typing import Any, Iterator

RECORD = Path(__file__).resolve().parent / "fixtures" / "match_rules" / "characterization_equip.json"

MICRO_COLS = ("ライン", "機器名", "ID", "出力テキスト", "厚下", "厚上", "用途", "備考", "更新日", "番号")
#: 試験用の測定機器ID(用途の書き方・空の厚・用途が全部空のライン・重なる行)
MICRO_ROWS: tuple[tuple[str, ...], ...] = (
    ("SYN", "M1", "S01", "20001", "0", "10", "≠H191", "", "", "1"),
    ("SYN", "M2", "S02", "20002", "0", "10", "*K2", "", "", "2"),
    ("SYN", "M3", "S03", "20003", "0", "10", "A5*", "", "", "3"),
    ("SYN", "M4", "S04", "20004", "0", "10", "H191,H207,*Z9,Q1*", "", "", "4"),
    ("SYN", "M5", "S05", "20005", "", "", "XX", "", "", "5"),
    ("SYN", "M6", "S06", "20006", "10.01", "20", "", "", "", "6"),
    ("SYN", "M7", "S07", "20007", "", "", "", "", "", "7"),
    ("SYN", "デプス", "S08", "20008", "", "", "", "", "", "8"),
    ("SYN", "M8", "S09", "20009", "0", "0", "", "", "", "9"),
    ("SYN", "M9", "S10", "20010", "abc", "5x", "", "", "", "10"),
    ("SYN2", "M1", "T01", "30001", "0", "24.99", "", "", "", "11"),
    ("SYN2", "M2", "T02", "30002", "24.99", "28.01", "", "", "", "12"),
    ("SYN2", "M3", "T03", "30003", "", "", "", "", "", "13"),
    ("ｼﾝ", "M1", "U01", "40001", "0", "99", "≠", "", "", "14"),
    ("シン", "M2", "U02", "40002", "0", "99", "*", "", "", "15"),
)
MICRO_LINES = ("L-1", "LVC", "HVC", "ﾊﾞﾗﾝｻｰ", "バランサー", "ﾄｯﾄ", "トット", "大板", "Test", "Test_", "丸徳", "無い",
               "SYN", "SYN2", "ｼﾝ", "シン", "")
MICRO_YOUTO_EXTRA = ("", " ", "H191", "H19", "XH191Y", "H207", "K2", "AK2B", "A5", "A50", "BA5", "Z9", "Q1", "Q",
                     "XX", "≠H191", "*")
#: 用途を全部ためす厚(ほかの厚は FEW_YOUTO だけ)
FULL_ATU = (0.0, 5.0, 10.0, 15.0, 24.99, 25.0, 28.01, 60.0)
FEW_YOUTO = ("", "H191", "K2", "A5")
EQUIP_COLS = ("管理番号", "設備名", "製造材質", "オーダー厚MIN", "オーダー厚MAX", "オーダー幅MIN", "オーダー幅MAX",
              "オーダー丈MIN", "オーダー丈MAX", "単重MIN", "単重MAX", "作業可否", "備考")
#: 試験用の設備仕様(空の MIN・MAX・材質、重なる行、作業可否が空・否)
EQUIP_ROWS: tuple[tuple[str, ...], ...] = (
    ("101", "SYN", "52S,A52*,≠1100", "0.5", "6", "500", "1500", "1000", "3000", "0", "100", "可", ""),
    ("102", "SYN", "", "6.01", "20", "500", "1500", "1000", "3000", "0", "100", "否", ""),
    ("103", "SYN", "*83*", "", "", "", "", "", "", "", "", "", ""),
    ("104", "SYN", "", "0", "600", "0", "2000", "0", "6000", "0", "3000", "", ""),
    ("105", "SYN2", "5052", "1", "2", "1000", "1000", "2000", "2000", "10", "10", "条件付", ""),
    ("106", "ｼﾝ", "", "1", "2", "1000", "1250", "2000", "2500", "0", "50", "可", ""),
)


def _float(t: Any) -> Any:
    try:
        return float(str(t).strip())
    except (TypeError, ValueError):
        return None


def _around(vals: Any, step: float) -> list[float]:
    out: list[float] = []
    for t in vals:
        v = _float(t)
        if v is None:
            continue
        out += [round(v - step, 4), v, round(v + step, 4)]
    return out


def _lot_values() -> list[dict]:
    """仕掛の全ロットの画面の値(厚・用途)と設備仕様の引数。"""
    from tests.match_rules_cases import fixture_lots, lot_form
    out = []
    for lot in fixture_lots():
        lv = lot_form(lot)
        args = getattr(lv.reflect, "equipment_args", None) if lv.reflect else None
        out.append({"lot": lot, "atu": lv.form("ATU"), "youto": lv.form("Youto"),
                    "equip": list(args[1:]) if args else None})
    return out


def _table(db: Path, table: str) -> tuple[list[str], list[list[str]]]:
    con = sqlite3.connect(db)
    try:
        cur = con.execute(f'SELECT * FROM "{table}"')
        return [d[0] for d in cur.description], [["" if v is None else str(v) for v in r] for r in cur.fetchall()]
    finally:
        con.close()


def micro_cases(db: Path) -> list[tuple]:
    """(ライン, 厚, 用途, n, dep)。"""
    fields, rows = _table(db, "測定機器ID")
    c = {n: fields.index(n) for n in fields}
    lots = _lot_values()
    atus: list[float] = [0.0, 0.01, 0.6, 24.99, 25.0, 28.0, 28.005, 28.01, 28.02, 300.0, 1000.0, -1.0]
    for lv in lots:
        v = _float(lv["atu"])
        if v is not None:
            atus.append(v)
    atus += _around([r[c["厚下"]] for r in rows] + [r[c["厚上"]] for r in rows], 0.01)
    atus = sorted(set(atus))
    youtos = list(dict.fromkeys([lv["youto"] for lv in lots] + list(MICRO_YOUTO_EXTRA)
                                + [y for r in rows for y in r[c["用途"]].split(",") if r[c["用途"]]]))
    out: list[tuple] = []
    for line in MICRO_LINES:
        for a in atus:
            for y in (youtos if a in FULL_ATU else FEW_YOUTO):
                out.append((line, a, y, "", ""))
            out.append((line, a, "H191", "機器名", ""))
            out.append((line, a, "", "ID", ""))
        for dep in ("デプス", "M1", "M9", "無い"):
            out.append((line, 0.0, " ", "", dep))
    return out


def equip_cases(db: Path) -> list[tuple]:
    """(ライン, 材質, 厚, 幅, 丈, 単重)。"""
    fields, rows = _table(db, "設備仕様")
    c = {n: fields.index(n) for n in fields}
    out: list[tuple] = []
    lots = [lv["equip"] for lv in _lot_values() if lv["equip"]]
    for line in ("LVC", "HVC", "SVC", "トット", "ﾄｯﾄ", "Test", "L-1", "SYN", "SYN2", "シン", "無い", ""):
        for e in lots:
            out.append((line, *e))
    for r in rows:
        line = r[c["設備名"]]
        mats = [m.strip().strip("*").lstrip("≠") for m in r[c["製造材質"]].split(",") if m.strip()] or ["5052"]
        zai = mats[0] or "5052"

        def mid(lo: str, hi: str, dflt: float) -> float:
            a, b = _float(r[c[lo]]), _float(r[c[hi]])
            return dflt if a is None or b is None else round((a + b) / 2, 2)
        base = [zai, mid("オーダー厚MIN", "オーダー厚MAX", 1.0), mid("オーダー幅MIN", "オーダー幅MAX", 1000.0),
                mid("オーダー丈MIN", "オーダー丈MAX", 2000.0), mid("単重MIN", "単重MAX", 10.0)]
        out.append((line, *base))
        for m in mats[:4] + ["ZZ外れ", "", "1100"]:
            out.append((line, m, *base[1:]))
        for k, (lo, hi, step) in enumerate((("オーダー厚MIN", "オーダー厚MAX", 0.01), ("オーダー幅MIN", "オーダー幅MAX", 1.0),
                                            ("オーダー丈MIN", "オーダー丈MAX", 1.0), ("単重MIN", "単重MAX", 1.0)), start=1):
            for v in _around([r[c[lo]], r[c[hi]]], step):
                b = list(base)
                b[k] = v
                out.append((line, *b))
    seen, uniq = set(), []
    for x in out:
        k = json.dumps(x, ensure_ascii=False)
        if k not in seen:
            seen.add(k)
            uniq.append(x)
    return uniq


def run_micro(case: tuple) -> str:
    from ufkensa.kensa import instruments
    line, atu, y, n, dep = case
    try:
        return instruments.db_return_micro(line, atu, y, n, dep)
    except Exception as e:  # noqa: BLE001 - VBA なら実行時エラー。種類を記録する
        return "ERR:" + type(e).__name__


def run_equip(case: tuple) -> str:
    from ufkensa.kensa import quality
    try:
        return quality.check_equipment_db(*case)
    except Exception as e:  # noqa: BLE001
        return "ERR:" + type(e).__name__


def synthetic_db(tmp: Path) -> Path:
    """基準データを写して、測定機器ID・設備仕様に試験用の行を足す。"""
    from tests import kensa_fixture
    folder = kensa_fixture.copy_to(tmp)
    con = sqlite3.connect(folder / "伝送用ファイル.sqlite3")
    with con:
        con.executemany('INSERT INTO "測定機器ID" VALUES (' + ",".join("?" * len(MICRO_COLS)) + ")", MICRO_ROWS)
        con.executemany('INSERT INTO "設備仕様" VALUES (' + ",".join("?" * len(EQUIP_COLS)) + ")", EQUIP_ROWS)
    con.close()
    return folder


def _key(c: tuple) -> str:
    return json.dumps(c, ensure_ascii=False)


def each(tmp: Path) -> Iterator[tuple[str, str, Any, tuple]]:
    """(表, 種類, 結果を出す関数, 入力)。表 = base(基準データ)/ syn(試験用の行を足した表)。"""
    from tests import kensa_fixture
    from ufkensa.kensa import cache_registry
    base = kensa_fixture.use_offline()
    syn = synthetic_db(tmp / "syn")
    for name, folder in (("base", base), ("syn", syn)):
        kensa_fixture.use_offline(folder)
        cache_registry.clear_all()
        db = folder / "伝送用ファイル.sqlite3"
        for c in micro_cases(db):
            yield name, "micro", run_micro, c
        for c in equip_cases(db):
            yield name, "equip", run_equip, c
    kensa_fixture.use_offline()
    cache_registry.clear_all()


def record(tmp: Path) -> dict:
    out: dict = {"micro": {}, "equip": {}}
    for name, kind, fn, c in each(tmp):
        out[kind][name + " " + _key(c)] = fn(c)
    return out


def main() -> None:
    import tempfile
    from tests import _isolation  # noqa: F401 - 設定を一時フォルダへ
    with tempfile.TemporaryDirectory() as t:
        got = record(Path(t))
    if "--write" in sys.argv:
        RECORD.write_text(json.dumps(got, ensure_ascii=False, indent=0, sort_keys=True), encoding="utf-8")
        print(f"書きました: {RECORD} (測定機器ID {len(got['micro'])} 件・設備仕様 {len(got['equip'])} 件)")
    else:
        from collections import Counter
        for k in ("micro", "equip"):
            print(k, len(got[k]), Counter(got[k].values()).most_common(12))


if __name__ == "__main__":
    main()
