"""照合ルール 第2段(平坦度一覧QP追加 = QP.DBQP の照合 15 条件)の特性テストの入力と、いまの結果の記録(2026-11-11)。

記録は `tests/fixtures/match_rules/characterization_qp.json`。作り直すとき(いまのコードの結果を正とする):

    python -m tests.match_rules_qp_cases --write

- 基準データの仕掛の全ロット × 引当(納入先・取引先)ごと(+ 引当なしの組)
- 境界の組: 表の各行の 板厚・板幅・板丈 の下限・上限ちょうど / ±0.01・±1 をオーダー寸法に入れ、
  用途・材質・調質をその行の最初の値に合わせたもの、条件を 1 つずつ外したもの
結果は 歪公差(Kousa)・知らせと、画面に書く値・trace(手順・文言・合否・残った行)の sha1。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

RECORD = Path(__file__).resolve().parent / "fixtures" / "match_rules" / "characterization_qp.json"
TABLE = "平坦度一覧QP追加"


def _cases() -> list[tuple[str, str, str, str, dict]]:
    """(ロット, 検番, 納入先, 取引先, 仕掛の行の上書き) の組。"""
    from tests.match_rules_cases import fixture_lots, lot_form
    from ufkensa.kensa import datasource, quality
    out: list[tuple[str, str, str, str, dict]] = []
    lots = fixture_lots()
    for lot in lots:
        lv = lot_form(lot)
        hiki = quality.hiki_arr_create(lot) or []
        pairs = list(dict.fromkeys((quality.vstr(n), quality.vstr(t)) for n, t in hiki)) or [("", "")]
        for n, t in pairs:
            out.append((lot, lv.ken_no or "", n, t, {}))
    base = next((x for x in lots if lot_form(x).data is not None and lot_form(x).data.records), None)
    if base is None:
        return out
    db = quality._quality_db()
    fields = datasource.get_fields(db, TABLE) or []
    ken = lot_form(base).ken_no or ""
    for r in datasource.get_records(db, TABLE) or []:
        g = {n: quality.vstr(r[i]) for i, n in enumerate(fields)}

        def first(name: str) -> str:
            return g.get(name, "").split(",")[0].strip().strip("*").lstrip("≠")

        head = first("検査番号の頭")
        kn = (head[:1] + ken[1:]) if head else ken
        over = {"用途ｺｰﾄﾞ": first("用途コード"), "製造材質": first("製造材質"), "ｵｰﾀﾞｰ材質": first("オーダー材質"),
                "製造調質": first("製造調質"), "ｵｰﾀﾞｰ調質": first("オーダー調質"),
                "ｵｰﾀﾞｰ板厚": g.get("板厚下") or "1", "ｵｰﾀﾞｰ板幅": g.get("板幅下") or "1000",
                "ｵｰﾀﾞｰ板丈": g.get("板丈下") or "2000"}
        nou, tori = first("納入先"), first("取引先")
        out.append((base, kn, nou, tori, dict(over)))
        for key, lo, hi, step in (("ｵｰﾀﾞｰ板厚", "板厚下", "板厚上", 0.01), ("ｵｰﾀﾞｰ板幅", "板幅下", "板幅上", 1.0),
                                  ("ｵｰﾀﾞｰ板丈", "板丈下", "板丈上", 1.0)):
            for t in (g.get(lo, ""), g.get(hi, "")):
                try:
                    v = float(t)
                except ValueError:
                    continue
                for d in (v - step, v, v + step):
                    out.append((base, kn, nou, tori, {**over, key: f"{d:g}"}))
        for key in ("用途ｺｰﾄﾞ", "製造材質", "ｵｰﾀﾞｰ材質", "製造調質", "ｵｰﾀﾞｰ調質"):
            out.append((base, kn, nou, tori, {**over, key: "ZZ外れ"}))
        out.append((base, kn, "ZZ外れ", tori, over))
        out.append((base, kn, nou, "ZZ外れ", over))
    seen, uniq = set(), []
    for c in out:
        k = json.dumps(c, ensure_ascii=False, sort_keys=True)
        if k not in seen:
            seen.add(k)
            uniq.append(c)
    return uniq


def _lot_data(lot: str, over: dict) -> Any:
    from tests.match_rules_cases import lot_form
    from ufkensa.kensa import quality
    d = lot_form(lot).data
    if not over or d is None or not d.records:
        return d
    rec = list(d.records[0])
    for k, v in over.items():
        if k in d.fields:
            rec[d.fields.index(k)] = v
    return quality.LotData(list(d.fields), [rec] + [list(x) for x in d.records[1:]])


def _trace(tr: list) -> list:
    import re
    def norm(x: Any) -> Any:          # 試験の DB の一時フォルダは毎回ちがう
        return re.sub(r"\S*[/\\]([^/\\\s]+\.sqlite3)", r"<db>/\1", x) if isinstance(x, str) else x
    return [[t.get("step"), norm(t.get("detail")), t.get("ok"), t.get("kept"), t.get("before"),
             [h.get("keys") for h in t.get("hits") or []], t.get("rows") or [], t.get("chosen")] for t in tr]


def run(case: tuple) -> dict:
    from ufkensa.kensa import quality, vbx
    lot, ken, nou, tori, over = case
    try:
        r = quality.dbqp(lot, ken, nou, tori, _lot_data(lot, over), execution="Test")
    except Exception as e:  # noqa: BLE001 - VBA なら実行時エラー。種類を記録する
        return {"error": type(e).__name__}
    return {"kousa": [[vbx.cstr(v) if v is not None else None for v in row] for row in r.kousa or []]
            if r.kousa is not None else None,
            "form": {k: vbx.cstr(v) for k, v in r.form.items()}, "uf1": {k: vbx.cstr(v) for k, v in r.uf1.items()},
            "messages": [m.text for m in r.messages], "trace": _trace(r.trace)}


def _compact(full: dict) -> dict:
    """記録を小さくする: 歪公差と知らせはそのまま、画面の値と trace は sha1 で持つ(違えば中身を run で見る)。"""
    import hashlib
    if "error" in full:
        return full
    rest = json.dumps({k: full[k] for k in ("form", "uf1", "trace")}, ensure_ascii=False, sort_keys=True)
    return {"kousa": full["kousa"], "messages": full["messages"],
            "digest": hashlib.sha1(rest.encode("utf-8")).hexdigest()[:16]}


def record() -> dict:
    return {json.dumps(c, ensure_ascii=False, sort_keys=True): _compact(run(c)) for c in _cases()}


def main() -> None:
    from tests import _isolation  # noqa: F401 - 設定を一時フォルダへ
    from tests import kensa_fixture
    kensa_fixture.use_offline()
    got = record()
    if "--write" in sys.argv:
        RECORD.write_text(json.dumps(got, ensure_ascii=False, indent=0, sort_keys=True), encoding="utf-8")
        print(f"書きました: {RECORD} ({len(got)} 件)")
    else:
        print(len(got), sum(1 for v in got.values() if "error" in v))


if __name__ == "__main__":
    main()
