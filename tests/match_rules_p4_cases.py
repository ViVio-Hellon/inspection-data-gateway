"""照合ルール 第4段(外観_H = Q_S_QA_124_別表2 / 外観_S = SRF_128_別表1 / ラテラル = DIM_179 日車・白銅滑川)の特性テストの入力と、いまの結果の記録(2026-11-11)。

記録は `tests/fixtures/match_rules/characterization_p4.json`。作り直すとき(いまのコードの結果を正とする):

    python -m tests.match_rules_p4_cases --write

- 仕掛の全ロット(そのまま / 外観_H はオーダー厚 10 にして厚板の道へ / 外観_S はオーダー厚 2.5 にして薄板の道へ)
- いくつかのロットで項目を 1 つずつ差し替えた組: 用途・材質・調質・納入先・表面処理グレード(表のセルの語 + 含む・
  含まない・空)、寸法(表の下限・上限ちょうど / ±0.01、空 = 型エラー)
- 表: base(基準データ)/ syn(試験用の行を足した表。124 別表2 には 調質 の列も足す。空のセル・≠・*X*・
  厚フラグなし・同じ調質の重なり・Null の材質)
"""
from __future__ import annotations

import json
import sqlite3
import sys
from pathlib import Path
from typing import Any, Iterator

RECORD = Path(__file__).resolve().parent / "fixtures" / "match_rules" / "characterization_p4.json"

H_BASE = ("H8401Z0", "L726M50", "L9150S0")       # 外観_H・外観_S・ラテラルで差し替えに使うロット
S_BUTTONS = ("トビコミ", "キズ,スリキズ", "無いボタン")
LATERAL = ("日車", "白銅滑川")


def _float(t: Any) -> Any:
    try:
        return float(str(t).strip())
    except (TypeError, ValueError):
        return None


def _around(vals: Any, step: float = 0.01) -> list[str]:
    out: list[str] = []
    for t in vals:
        v = _float(t)
        if v is None:
            continue
        out += [str(round(v - step, 4)), str(v), str(round(v + step, 4))]
    return list(dict.fromkeys(out))


def _tokens(vals: Any) -> list[str]:
    out: list[str] = []
    for v in vals:
        for t in str(v or "").split(","):
            t = t.strip()
            out += [t, t.strip("*").lstrip("≠")]
    return [x for x in dict.fromkeys(out)]


def _tab(table: str) -> tuple[list[str], list[list[Any]]]:
    from ufkensa.kensa import datasource, quality as q
    db = q._quality_db()
    return datasource.columns(db, table), datasource.get_records(db, table) or []


def _col(table: str, name: str) -> list[Any]:
    f, rows = _tab(table)
    if name not in f:
        return []
    i = f.index(name)
    return [r[i] for r in rows]


def _lots() -> list[str]:
    from tests.match_rules_cases import fixture_lots
    return fixture_lots()


def lot_with(lot_no: str, changes: dict[str, str]) -> Any:
    from ufkensa.kensa import quality as q
    ld = q.read_lot(lot_no)
    if not ld.records or not ld.fields:
        return ld
    rec = list(ld.records[0])
    for k, v in changes.items():
        rec[ld.fields.index(k)] = v
    return q.LotData(ld.fields, [rec])


def h_cases() -> list[tuple[str, dict]]:
    t = "Q_S_QA_124_別表2"
    out: list[tuple[str, dict]] = [("無いロット", {})]
    for lot in _lots():
        out += [(lot, {}), (lot, {"ｵｰﾀﾞｰ板厚": "10"})]
    var: dict[str, list[str]] = {
        "用途ｺｰﾄﾞ": _tokens(_col(t, "用途コード")) + ["", "H1", "XH191Y", "*"],
        "製造材質": _tokens(_col(t, "材質")) + ["", "A5052S", "1100", "ZZ"],
        "製造調質": ["", "H112", "O", "T6"],
        "製造板厚": _around(_col(t, "板厚下") + _col(t, "板厚上")) + ["", "0"],
        "ｵｰﾀﾞｰ板幅": _around(_col(t, "板幅下") + _col(t, "板幅上"), 1) + [""],
        "ｵｰﾀﾞｰ板丈": _around(_col(t, "板丈下") + _col(t, "板丈上"), 1),
        "納入先名": _tokens(_col(t, "納入先")) + ["", "ﾆｯｼｬ", "X"],
    }
    for lot in H_BASE:
        for k, vals in var.items():
            for v in vals:
                out.append((lot, {"ｵｰﾀﾞｰ板厚": "10", k: v}))
    return out


def s_cases() -> list[tuple[str, dict, str]]:
    t = "Q_S_QA_SRF_128_別表1"
    out: list[tuple[str, dict, str]] = [("無いロット", {}, "トビコミ")]
    for lot in _lots():
        out += [(lot, {}, b) for b in S_BUTTONS[:2]]
    var: dict[str, list[str]] = {
        "用途ｺｰﾄﾞ": _tokens(_col(t, "用途コード")) + ["", "H1", "*"],
        "製造材質": _tokens(_col(t, "材質")) + ["", "A5052S", "1100"],
        "品質ｸﾞﾚｰﾄﾞ_表面処理": _tokens(_col(t, "表面処理コード")) + ["", "Z", "B,H", "*"],
        "ｵｰﾀﾞｰ板厚": _around(_col(t, "板厚1")) + ["0", "7.99", "8"],
    }
    for lot in H_BASE:
        for k, vals in var.items():
            for v in vals:
                ch = {"ｵｰﾀﾞｰ板厚": "2.5", k: v} if k != "ｵｰﾀﾞｰ板厚" else {k: v}
                for b in S_BUTTONS:
                    out.append((lot, ch, b))
    return out


def lateral_cases() -> list[tuple[str, dict, str]]:
    out: list[tuple[str, dict, str]] = [("無いロット", {}, "日車"), (H_BASE[0], {}, ""), (H_BASE[0], {}, "x")]
    for lot in _lots():
        out += [(lot, {}, w) for w in LATERAL]
    for w in LATERAL:
        t = "Q_S_QA_DIM_179_" + w
        f, _rows = _tab(t)
        nums = [c for c in f if _float(c) is not None]
        var: dict[str, list[str]] = {
            "製造材質": _tokens(_col(t, "材質")) + ["", "A5052S", "ZZ"],
            "製造板厚": _around(_col(t, "板厚下") + _col(t, "板厚上")) + [""],
            "製造板幅": _around(_col(t, "板幅下") + _col(t, "板幅上"), 0.1),
            "製造板丈": _around(nums, 1) + ["0", "99999"],
        }
        for lot in H_BASE:
            for k, vals in var.items():
                for v in vals:
                    out.append((lot, {"製造板厚": "3", "製造板幅": "1000", "製造材質": "A5052S", k: v}, w))
    return out


def _err(e: Exception) -> str:
    return "ERR:" + type(e).__name__


def _msgs(ms: Any) -> list[str]:
    return [m.text for m in ms]


def run_h(case: tuple) -> Any:
    from ufkensa.kensa import quality_view as qv
    lot, ch = case
    try:
        r = qv.display_appearance_h(lot_with(lot, ch), qv.QualityForm())
    except Exception as e:  # noqa: BLE001 - VBA なら実行時エラー。種類を記録する
        return _err(e)
    f = r.form
    return {"g": r.grade_array, "name": f.AbnoName, "k": f.KousinL1, "lab": f.AbnoLabel, "gr": f.gradeTB,
            "tb": f.AbnoTB, "tabs": f.tabs, "m": _msgs(r.messages)}


def run_s(case: tuple) -> Any:
    from ufkensa.kensa import quality_view as qv
    lot, ch, b = case
    try:
        r = qv.display_appearance_s(lot_with(lot, ch), b, qv.QualityForm())
    except Exception as e:  # noqa: BLE001
        return _err(e)
    f = r.form
    return {"p": f.Upater, "sel": [f.SelectO, f.SelectS, f.SelectK, f.SelectA], "k": f.KousinL4, "s": f.Sansyou1,
            "l2": f.L2_visible, "m": _msgs(r.messages)}


def run_lateral(case: tuple) -> Any:
    from ufkensa.kensa import quality as q
    lot, ch, w = case
    msgs: list = []
    try:
        v = q.dbdisplay_lateral(lot_with(lot, ch), w, messages=msgs)
    except Exception as e:  # noqa: BLE001
        return _err(e)
    return [v, _msgs(msgs)]


#: 試験用に足す行(表 → (列, 行))
SYN_ROWS: dict[str, tuple[tuple[str, ...], tuple[tuple[Any, ...], ...]]] = {
    "Q_S_QA_124_別表2": (
        ("管理番号", "材質", "用途コード", "板厚下", "板厚上", "厚", "飛込み", "キズ/スリキズ/スリ跡", "調質", "納入先",
         "板幅下", "板幅上", "幅", "板丈下", "板丈上", "丈", "更新日"),
        ((901, "", "", "", "", "", "1", "2", "", "", "", "", "", "", "", "", ""),
         (902, "≠*52S", "H191", "8", "200", "", "5", "5", "H112", "", "", "", "", "", "", "", ""),
         (903, "*52*", "H1", "8", "20", "", "3", "3", "≠O", "ﾆｯｼｬ,X*", "1000", "2000", "", "", "", "", ""),
         (904, None, None, "", "", "x", "2", "2", None, None, "", "", "", "2000", "4000", "", ""))),
    "Q_S_QA_SRF_128_別表1": (
        ("管理番号", "表面処理コード", "材質", "用途コード", "厚フラグ", "板厚1", "パターン1", "パターン2", "更新日"),
        ((901, "", "", "", "", "", "3", "", "syn"),
         (902, "B,H", "≠*52S", "H1", "レ", "2.5", "1", "2", ""),
         (903, None, None, None, None, None, "4", "4", None))),
    "Q_S_QA_DIM_179_日車": (
        ("管理番号", "材質", "板厚下", "板厚上", "厚", "板幅下", "板幅上", "幅", "750", "1525", "2300", "3050", "4500",
         "5200"),
        ((901, "", "", "", "", "", "", "", "9", "9", "9", "9", "9", "9"),
         (902, "≠*52S", "0", "10", "", "0", "2000", "", "7", "7", "7", "7", "7", "7"))),
    "Q_S_QA_DIM_179_白銅滑川": (
        ("管理番号", "材質", "板厚下", "板厚上", "厚", "板幅下", "板幅上", "幅", "1500", "2500", "9999"),
        ((901, "ZZ*", "0", "500", "", "0", "9999", "", "8", "8", "8"),)),
}


def synthetic_db(tmp: Path) -> Path:
    from tests import kensa_fixture
    folder = kensa_fixture.copy_to(tmp)
    con = sqlite3.connect(folder / "品質基準.sqlite3")
    with con:
        con.execute('ALTER TABLE "Q_S_QA_124_別表2" ADD COLUMN "調質" TEXT')
        for table, (cols, rows) in SYN_ROWS.items():
            con.executemany(f'INSERT INTO "{table}" ({",".join(chr(34) + c + chr(34) for c in cols)}) VALUES '
                            f'({",".join("?" * len(cols))})',
                            [tuple(None if v is None else str(v) for v in r) for r in rows])
    con.close()
    return folder


def _key(c: tuple) -> str:
    return json.dumps(c, ensure_ascii=False, sort_keys=True)


KINDS = ("h", "s", "lateral")


def each(tmp: Path) -> Iterator[tuple[str, str, Any, tuple]]:
    """(表, 種類, 結果を出す関数, 入力)。"""
    from tests import kensa_fixture
    from ufkensa.kensa import cache_registry
    base = kensa_fixture.use_offline()
    syn = synthetic_db(tmp / "syn")
    for name, folder in (("base", base), ("syn", syn)):
        kensa_fixture.use_offline(folder)
        cache_registry.clear_all()
        for kind, cases, fn in (("h", h_cases, run_h), ("s", s_cases, run_s), ("lateral", lateral_cases, run_lateral)):
            for c in cases():
                yield name, kind, fn, c
    kensa_fixture.use_offline()
    cache_registry.clear_all()


def record(tmp: Path) -> dict:
    out: dict = {k: {} for k in KINDS}
    for name, kind, fn, c in each(tmp):
        out[kind][name + " " + _key(c)] = json.loads(json.dumps(fn(c), ensure_ascii=False, default=str))
    return out


def main() -> None:
    import tempfile
    from tests import _isolation  # noqa: F401 - 設定を一時フォルダへ
    with tempfile.TemporaryDirectory() as t:
        got = record(Path(t))
    if "--write" in sys.argv:
        RECORD.write_text(json.dumps(got, ensure_ascii=False, indent=0, sort_keys=True), encoding="utf-8")
        print(f"書きました: {RECORD} (" + "・".join(f"{k} {len(got[k])} 件" for k in KINDS) + ")")
    else:
        from collections import Counter
        for k in KINDS:
            print(k, len(got[k]), Counter(json.dumps(v, ensure_ascii=False)[:60] for v in got[k].values()).most_common(8))


if __name__ == "__main__":
    main()
