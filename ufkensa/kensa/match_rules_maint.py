"""保守 → 照合ルール の組み立て(2026-11-11)。表は読むだけ。保存は match_rules.save。

- 一覧: ルール群ごとのルールと、選べる表の列(その表に在る列だけ)・ロットの項目
- 試し照合: ロットNo を入れると、引当(ラテフリ直角)/ Lot情報履歴 の受注(梱包保護材)ごとに、
  ルールの順に「残った行の数・どの行(管理番号・行)」と最後の結果。保存済みと編集中を並べる
- 保存の前の確かめ: 仕掛の全ロット(ラテフリ直角)/ Lot情報履歴 の全受注(梱包保護材)を
  保存済みと編集中のルールで通し、結果が変わるものを並べる
"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Optional, Sequence

from . import datasource, match_rules as mr, quality as q, quality_session as qs, wiring
from .store import LOT_HISTORY, LocalStore

ROW_CAP = 30               # 1 ステップで並べる行の数
LOT_CAP = 3000             # 保存の前の確かめで通すロットの数


def _table(g: mr.Group) -> tuple[Optional[Path], list[str], list[list[Any]]]:
    if g.name == mr.PACK:
        path = q._packing_db("梱包資材マスタ.sqlite3")
    else:
        path = q._quality_db()
    try:
        f = datasource.get_fields(path, g.table) or []
        r = datasource.get_records(path, g.table) or []
    except Exception:  # noqa: BLE001 - 表が無い・読めない
        return path, [], []
    return path, list(f), [list(x) for x in r]


def state() -> dict[str, Any]:
    mr.ensure_table()
    out = []
    for name, g in mr.groups(fresh=True).items():
        path, fields, rows = _table(g)
        out.append({**mr.to_payload(g), "table_fields": fields, "table_rows": len(rows), "table_path": str(path or ""),
                    "default": mr.to_payload(mr.BY_NAME[name])})
    return {"groups": out, "compares": list(mr.COMPARES), "preps": list(mr.PREPS), "blanks": list(mr.BLANKS),
            "use_rows": list(mr.USE_ROWS), "history": mr.history()}


def edited(name: str, payload: Optional[dict]) -> mr.Group:
    """画面の編集中のルール(確かめ済み)。payload が無ければ保存済み。"""
    if not payload:
        return mr.groups(fresh=True)[name]
    g = mr.from_payload(name, payload)
    mr.check(g, _table(g)[1])
    return g


# ------------------------------------------------------------------
# 1 件の照合と結果
# ------------------------------------------------------------------
def _key_col(fields: Sequence[str]) -> int:
    return q.col_of(fields, "管理番号")


def _label(rows: list[list[Any]], fields: Sequence[str], i: int) -> dict[str, Any]:
    c = _key_col(fields)
    return {"row": i + 1, "no": q.vstr(rows[i][c - 1]) if c else ""}


def _result(g: mr.Group, rows: list[list[Any]], fields: Sequence[str], res: mr.MatchResult, ctx: dict) -> str:
    if g.name == mr.PACK:
        c = q.col_of(fields, "使用保護材")
        return q.vstr(rows[res.rows[0]][c - 1]) if res.rows and c else "一致なし"
    if not res.rows:
        return "該当なし"
    nos = [q.clng(q.cell(rows[i], _key_col(fields))) for i in res.rows]
    hit = q._hit_rows(rows, _key_col(fields), nos)
    items = []
    for h in hit:
        for n in ("フリー反り", "ラテラルボー", "直角度"):
            if q.vstr(q.cell(h, q.col_of(fields, n))) == "要":
                items.append(n)
    kousa = q.vstr(q.cell(hit[0], q.col_of(fields, "公差")))
    return ("・".join(dict.fromkeys(items)) or "要なし") + f"(公差 {kousa})"


def run(g: mr.Group, rows: list[list[Any]], fields: Sequence[str], ctx: dict) -> dict[str, Any]:
    try:
        res = mr.match(g, rows, fields, ctx)
    except Exception as e:  # noqa: BLE001 - VBA なら実行時エラー
        return {"error": f"{type(e).__name__}: {getattr(e, 'text', '') or e}", "steps": [], "result": "エラー"}
    steps = [{"order": s.rule.order, "rule": f"{','.join(s.rule.columns)} ← {s.rule.field} : {s.rule.compare}",
              "value": q.vstr(ctx.get(s.rule.field, "")), "count": len(s.remain),
              "rows": [_label(rows, fields, i) for i in s.remain[:ROW_CAP]]} for s in res.steps]
    return {"steps": steps, "matched": [_label(rows, fields, i) for i in res.matched[:ROW_CAP]],
            "used": [_label(rows, fields, i) for i in res.rows[:ROW_CAP]], "result": _result(g, rows, fields, res, ctx)}


# ------------------------------------------------------------------
# ロットから照合の入力(ctx)を作る
# ------------------------------------------------------------------
def _ratefuri_inputs(store: LocalStore, lot: str) -> list[tuple[str, dict]]:
    lv = qs.read_lot(lot)
    _src, rows = qs.allocations(store, lv)
    out, seen = [], set()
    for row in rows:
        if qs._s(row[0]) == "":
            continue
        try:
            dims = [q.csng(row[k]) for k in (5, 6, 7)]
        except Exception:  # noqa: BLE001 - 寸法が無い
            continue
        args = (qs._s(row[0]), qs._s(row[1]), qs._s(row[2]), qs._s(row[3]), qs._s(row[4]), *dims)
        if args in seen:
            continue
        seen.add(args)
        out.append((f"{args[0]} / {args[1]}", q.ratefuri_ctx(args)))
    return out


def _hist(store: LocalStore) -> list[list[Any]]:
    return wiring.sheet_rows(store, LOT_HISTORY) or []


def _pack_inputs(hist: list[list[Any]], lot: Optional[str]) -> list[tuple[str, dict]]:
    if not hist:
        return []
    hdr = hist[0]
    col = {q.vstr(h): i for i, h in enumerate(hdr)}
    out, seen = [], set()
    for a in hist[1:]:
        if lot is not None and q.vstr(a[col.get("ﾛｯﾄNo", 0)]) != lot:
            continue
        ordno = q.vstr(a[col["ｵｰﾀﾞｰNo"]]) if "ｵｰﾀﾞｰNo" in col else ""
        if ordno in seen or not q.is_target_hoso_no(q.vstr(a[col["包装仕様NO"]]) if "包装仕様NO" in col else ""):
            continue
        seen.add(ordno)
        for b in (False, True):
            try:
                ctx = q.pack_ctx(hdr, a, b)
            except Exception:  # noqa: BLE001 - 厚などが数でない(VBA なら型エラー)
                continue
            out.append((f"受注 {ordno}{'(当板用)' if b else ''}", ctx))
    return out


def trial(name: str, lot: str, payload: Optional[dict], store: LocalStore) -> dict[str, Any]:
    saved = mr.groups(fresh=True)[name]
    new = edited(name, payload)
    _path, fields, rows = _table(saved)
    inputs = _ratefuri_inputs(store, lot) if name == mr.RATEFURI else _pack_inputs(_hist(store), lot)
    cases = []
    for label, ctx in inputs:
        a, b = run(saved, rows, fields, ctx), run(new, rows, fields, ctx)
        cases.append({"label": label, "saved": a, "edited": b, "changed": a["result"] != b["result"]})
    note = "" if inputs else ("引当がありません(仕掛・Lot情報履歴)" if name == mr.RATEFURI
                              else "Lot情報履歴 にこのロットの対象の包装仕様(蓋アングル)の行がありません")
    return {"lot": lot, "cases": cases, "note": note, "table_rows": len(rows)}


def regression(name: str, payload: Optional[dict], store: LocalStore) -> dict[str, Any]:
    """保存の前の確かめ。結果が変わる入力を並べる。"""
    saved = mr.groups(fresh=True)[name]
    new = edited(name, payload)
    _path, fields, rows = _table(saved)
    if name == mr.RATEFURI:
        from . import maint
        lots = maint.lot_numbers(plate_only=False)[:LOT_CAP]
        inputs = [(lot, lab, ctx) for lot in lots for lab, ctx in _ratefuri_inputs(store, lot)]
    else:
        hist = _hist(store)
        inputs = []
        for lab, ctx in _pack_inputs(hist, None):
            inputs.append((q.vstr(ctx.get("ﾛｯﾄNo", "")), lab, ctx))
    changed = []
    for lot, lab, ctx in inputs:
        a, b = run(saved, rows, fields, ctx)["result"], run(new, rows, fields, ctx)["result"]
        if a != b:
            changed.append({"lot": lot, "label": lab, "before": a, "after": b})
    return {"checked": len(inputs), "changed": changed[:200], "changed_count": len(changed)}
