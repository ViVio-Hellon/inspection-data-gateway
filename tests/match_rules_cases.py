"""照合ルール(第1段: ラテフリ直角・梱包保護材)の特性テストの入力と、いまの結果の記録(2026-11-11)。

記録は `tests/fixtures/match_rules/characterization.json`。作り直すとき(いまのコードの結果を正とする):

    python -m tests.match_rules_cases --write

- ラテフリ直角(QP.DbReturnItem): 基準データの仕掛の全ロットの引当(SIKAHIKINOW + 仕掛の製造寸法)と、
  表の各行の境界(厚・幅・丈の下限・上限ちょうど / ±0.01・±1)、条件を 1 つずつ外した組
- 梱包保護材(QP.MaterialsUpperPart): 受領していないので、ここで作る小さな表 + 仕掛の全ロットから組んだ
  Lot情報履歴 + 境界の組。当板用か(check_bikou)の両方
保守画面の「保存の前の確かめ」も、ここと同じ入力(cases_*)を使う。
"""
from __future__ import annotations

import json
import sqlite3
import sys
from pathlib import Path
from typing import Any

RECORD = Path(__file__).resolve().parent / "fixtures" / "match_rules" / "characterization.json"

PACK_COLS = ("管理番号", "包装仕様書", "用途コード", "材質", "調質", "板厚下", "板厚上", "板厚", "板幅下", "板幅上",
             "板幅", "板丈下", "板丈上", "板丈", "使用保護材", "備考")
#: 試験用の梱包保護材(条件式の書き方・空の列・当板用・重なる行をひと通り)
PACK_ROWS: tuple[tuple[Any, ...], ...] = (
    (1, "1P0001", "", "52S,A52N", "", "0.6", "3.0", "", "", "", "", "", "", "", "アングル", ""),
    (2, "1P0001", "", "52S,A52N", "", "0.6", "3.0", "", "", "", "", "", "", "", "上部当板A", "当板用"),
    (3, "1P0104,1P0110", "≠K21S", "", "H*", "3.01", "6.0", "", "1000", "1250", "", "", "", "", "ダンボール", ""),
    (4, "1P0104,1P0110", "", "*52*", "", "", "", "", "", "", "", "2000", "3054", "", "発泡", ""),
    (5, "", "", "", "", "", "", "", "", "", "", "", "", "", "共通保護材", "当板用"),
    (6, "1P0115,1P0116,1P0118", "K*", "", "O", "6.01", "28.0", "", "", "", "", "", "", "", "木枠", ""),
    (7, "1P0119,1P0120", "", "≠*83*", "", "0.6", "0.6", "0.6", "1219", "1220", "", "2438", "2438", "", "ちょうど", ""),
)
HIST_HDR = ("ﾛｯﾄNo", "ｵｰﾀﾞｰNo", "引当NO", "包装仕様NO", "厚", "巾", "丈", "材質", "調質", "用途ｺｰﾄﾞ")
TARGETS = ("1P0001", "1P0104", "1P0110", "1P0115", "1P0116", "1P0118", "1P0119", "1P0120", "9X9999")


def make_pack_db(path: Path) -> Path:
    con = sqlite3.connect(path)
    with con:
        con.execute('CREATE TABLE "梱包保護材" (' + ", ".join(f'"{c}" TEXT' for c in PACK_COLS) + ")")
        con.executemany('INSERT INTO "梱包保護材" VALUES (' + ",".join("?" * len(PACK_COLS)) + ")",
                        [tuple("" if v == "" else str(v) for v in r) for r in PACK_ROWS])
    con.close()
    return path


_forms: dict = {}


def lot_form(lot: str) -> Any:
    """ロットの画面の値(AccReflec)。同じロットは 1 回だけ読む。"""
    from ufkensa.kensa import quality_session as qs
    if lot not in _forms:
        _forms[lot] = qs.read_lot(lot)
    return _forms[lot]


def fixture_lots() -> list[str]:
    from ufkensa.kensa import maint
    return maint.lot_numbers(plate_only=False)


def _num(v: Any) -> float:
    from ufkensa.kensa import quality
    try:
        return quality.csng(v)
    except Exception:  # noqa: BLE001 - 仕掛に寸法が無いロット
        return 0.0


def _bounds(lo: str, hi: str, step: float) -> list[float]:
    out: list[float] = []
    for t in (lo, hi):
        try:
            v = float(t)
        except (TypeError, ValueError):
            continue
        out += [v - step, v, v + step]
    return out


def ratefuri_cases() -> list[tuple]:
    """(nou, tori, youtc, zai, tyou, atu, haba, take, lot) の組。lot は lateral_lot に使うロット("" は無し)。"""
    from ufkensa.kensa import datasource, quality, quality_session as qs
    cases: list[tuple] = []
    for lot in fixture_lots():
        lv = lot_form(lot)
        f = lv.form
        for n, t in quality.hiki_arr_create(lot) or []:
            cases.append((qs._s(n), qs._s(t), f("Youto"), f("Zai"), f("Tyu"), _num(f("ATU")), _num(f("Hab")),
                          _num(f("Tak")), lot))
    db = quality._quality_db()
    fields = datasource.get_fields(db, "ラテフリ直角") or []
    for r in datasource.get_records(db, "ラテフリ直角") or []:
        g = {n: quality.vstr(r[i]) for i, n in enumerate(fields)}

        def first(name: str) -> str:
            return g[name].split(",")[0].strip().strip("*").lstrip("≠")

        base = [first("納入先"), first("取引先"), first("用途コード"), first("製造材質"), first("製造調質")]
        dims = [g["厚下"] or "1", g["幅下"] or "1000", g["丈下"] or "2000"]
        num = [float(x) if x.replace(".", "", 1).isdigit() else 1.0 for x in dims]
        cases.append((*base, *num, ""))
        for k, (lo, hi, step) in enumerate(((g["厚下"], g["厚上"], 0.01), (g["幅下"], g["幅上"], 1.0),
                                            (g["丈下"], g["丈上"], 1.0))):
            for v in _bounds(lo, hi, step):
                d = list(num)
                d[k] = v
                cases.append((*base, *d, ""))
        for k in range(5):            # 条件を 1 つずつ外す(該当なしの文言は最後の行の Conditions で決まる)
            b = list(base)
            b[k] = "ZZ外れ"
            cases.append((*b, *num, ""))
        for nou in ("ﾆﾎﾝｼﾔﾘﾖｳｾｲｿﾞｳ(ｶ", "K.T.N. CO. LTD", "ﾆﾎﾝﾊﾂｼﾞﾖｳ(ｶ"):
            cases.append((nou, "ﾅﾒｶﾜｱﾙﾐ(ｶ", *base[2:], *num, ""))
            for k in range(3):
                d = list(num)
                d[k] = 99999.0
                cases.append((nou, base[1], *base[2:], *d, ""))
    return list(dict.fromkeys(cases))


def run_ratefuri(case: tuple, lots: dict) -> dict:
    from ufkensa.kensa import quality, quality_session as qs, vbx
    *args, lot = case
    if lot and lot not in lots:
        lots[lot] = lot_form(lot).data
    try:
        r = quality.db_return_item(*args, lateral_lot=lots.get(lot) if lot else None, lateral_quiet=True)
    except Exception as e:  # noqa: BLE001 - VBA なら実行時エラー。種類を記録する
        return {"error": type(e).__name__}
    return {"items": r.items, "uf1": {k: vbx.cstr(v) for k, v in r.uf1.items()},
            "cb_add_atu": r.cb_add_atu, "messages": [m.text for m in r.messages]}


def hist_rows() -> list[tuple]:
    from ufkensa.kensa import quality_session as qs
    rows: list[tuple] = [HIST_HDR]
    for i, lot in enumerate(fixture_lots()):
        f = lot_form(lot).form
        rows.append((lot, "O" + lot, "1", TARGETS[i % len(TARGETS)], f("ATU"), f("Hab"), f("Tak"), f("Zai"), f("Tyu"),
                     f("Youto")))
    n = 0
    # 条件ごとの値を 1 つずつ回す(全部の掛け合わせは多すぎるので、厚の境界 × 包装仕様 を軸に他をずらす)
    zais, tyos, yous = ("52S", "A52N", "F52S", "A5083", "A5183", ""), ("H18", "O", "H24", ""), ("K21S", "K22", "F975", "")
    dims = (("1000", "2000"), ("1219", "2438"), ("1220", "2438"), ("1250", "3054"), ("1251", "3055"), ("999", "1999"))
    for pac in TARGETS:
        for atu in ("0.59", "0.6", "0.61", "3.0", "3.01", "6.0", "6.01", "28.0", "28.01", ""):
            for k in range(6):
                n += 1
                hab, tak = dims[(n + k) % len(dims)]
                rows.append((f"S{n}", f"S{n}", "1", pac, atu, hab, tak, zais[k], tyos[(n + k) % 4], yous[(n // 3) % 4]))
    return rows


def run_pack(hist: list, ord_no: str, db: Path) -> list:
    from ufkensa.kensa import quality
    out = []
    for b in (False, True):
        try:
            out.append(quality.materials_upper_part(ord_no, hist, b, db=db))
        except Exception as e:  # noqa: BLE001
            out.append("ERR:" + type(e).__name__)
    return out


def record(tmp: Path) -> dict:
    lots: dict = {}
    rate = {json.dumps(c, ensure_ascii=False): run_ratefuri(c, lots) for c in ratefuri_cases()}
    db = make_pack_db(tmp / "梱包資材マスタ.sqlite3")
    hist = hist_rows()
    pack = {r[1]: run_pack(hist, r[1], db) for r in hist[1:]}
    return {"ratefuri": rate, "pack": pack}


def main() -> None:
    import tempfile
    from tests import _isolation  # noqa: F401 - 設定を一時フォルダへ
    from tests import kensa_fixture
    kensa_fixture.use_offline()
    with tempfile.TemporaryDirectory() as t:
        got = record(Path(t))
    if "--write" in sys.argv:
        RECORD.write_text(json.dumps(got, ensure_ascii=False, indent=0, sort_keys=True), encoding="utf-8")
        print(f"書きました: {RECORD} (ラテフリ直角 {len(got['ratefuri'])} 件・梱包保護材 {len(got['pack'])} 件)")
    else:
        print(len(got["ratefuri"]), len(got["pack"]))


if __name__ == "__main__":
    main()
