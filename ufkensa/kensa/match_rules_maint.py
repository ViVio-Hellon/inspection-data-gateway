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

from . import datasource, match_rules as mr, quality as q, quality_session as qs, vbx, wiring
from .store import LOT_HISTORY, LocalStore

#: 試し照合でロットの代わりに直接入れられる値(画面の入力欄)
INPUTS = {mr.MICRO: ["ライン", "厚", "用途"], mr.EQUIP: ["ライン", "オーダー厚"], mr.BH: ["材質", "調質"]}
#: 第4段: ロットの行(仕掛の先頭行)をそのまま入力にする群
SIKA_GROUPS = (mr.APPEAR_H, mr.APPEAR_S, mr.LATERAL_N, mr.LATERAL_H)
ROW_CAP = 30               # 1 ステップで並べる行の数
LOT_CAP = 3000             # 保存の前の確かめで通すロットの数


def _table(g: mr.Group) -> tuple[Optional[Path], list[str], list[list[Any]]]:
    if g.name == mr.PACK:
        path = q._packing_db("梱包資材マスタ.sqlite3")
    elif g.name in (mr.MICRO, mr.EQUIP):
        path = q._transfer_db()
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
                    "default": mr.to_payload(mr.BY_NAME[name]),
                    "field_sources": {f: mr.field_source(f, name) for f in g.fields},
                    "inputs": INPUTS.get(name, []),
                    "dim_swap": [list(x) for x in mr.DIM_SWAP.get(name, ())]})
    return {"groups": out, "help_switch": mr.HELP_SWITCH, "compares": list(mr.COMPARES), "preps": list(mr.PREPS), "blanks": list(mr.BLANKS),
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
    return q.col_of(fields, "管理番号") or q.col_of(fields, "番号")     # 測定機器ID は 番号


def _label(rows: list[list[Any]], fields: Sequence[str], i: int) -> dict[str, Any]:
    c = _key_col(fields)
    return {"row": i + 1, "no": q.vstr(rows[i][c - 1]) if c else ""}


def _p4_result(g: mr.Group, rows: list[list[Any]], fields: Sequence[str], res: mr.MatchResult, ctx: dict) -> str:
    """第4段の群の結果(コードに残した後処理を同じように当てる)。"""
    def c(n: str) -> int:
        return q.col_of(fields, n)
    if g.name == mr.BH:
        out = 0
        for i in res.rows:
            out = q.clng(q.cell(rows[i], c("バーコール下限値")))
        return f"下限値 {out}" if out else "0(BH 不要)"
    if not res.rows:
        return "該当用途ｺｰﾄﾞなし" if g.name == mr.APPEAR_S and not res.evaluated else "該当なし"
    nos = [q.vstr(q.cell(rows[i], c("管理番号"))) for i in res.rows]
    if g.name == mr.APPEAR_H:
        return f"管理番号 {', '.join(nos)}(後の番号の欠陥グレードで上書き)"
    if g.name == mr.APPEAR_S:
        r = rows[res.rows[0]]
        o_atu = q.csng(ctx.get("ｵｰﾀﾞｰ板厚", 0) or 0)
        if q.vstr(q.cell(r, c("厚フラグ"))) != "":
            k = "パターン1" if vbx.val(q.cell(r, c("板厚1"))) >= vbx.val_single(o_atu) else "パターン2"
        else:
            k = "パターン1"
        return f"管理番号 {nos[0]} → {k} = {q.vstr(q.cell(r, c(k)))}"
    hit = q._hit_rows(rows, c("管理番号"), [q.clng(n) for n in nos])
    v = q.loop_array(hit, list(fields), q.csng(ctx.get("製造板丈", 0) or 0), q.csng(ctx.get("製造板厚", 0) or 0))
    return f"管理番号 {', '.join(nos)} → ラテラルボー {vbx.fmt(v, '0.0')}"


def _result(g: mr.Group, rows: list[list[Any]], fields: Sequence[str], res: mr.MatchResult, ctx: dict) -> str:
    if g.name in SIKA_GROUPS or g.name == mr.BH:
        return _p4_result(g, rows, fields, res, ctx)
    if g.name == mr.QP:
        if not res.rows:
            return "合う行なし(等級表へ)"
        nos = [q.vstr(q.cell(rows[i], _key_col(fields))) for i in res.rows]
        lim = q.vstr(q.cell(rows[res.rows[0]], q.col_of(fields, "1m当たり平坦度上限")))
        return f"管理番号 {', '.join(nos)}(先頭の 1m当たり平坦度上限 {lim or '(空)'})"
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


def _micro_result(g: mr.Group, rows: list[list[Any]], fields: Sequence[str], ctx: dict) -> tuple[str, list[int]]:
    from . import instruments
    temp, hit = instruments._micro_new(g, fields, rows, q.vstr(ctx.get("ライン", "")), q.csng(ctx.get("厚", 0) or 0),
                                       q.vstr(ctx.get("用途", "")))
    if not temp:
        return "このラインの行なし(機器名は空)", []
    if hit is None:
        return "合う行なし(機器名は空)", []
    c = {k: q.col_of(fields, k) for k in ("機器名", "出力テキスト")}
    return f"{q.vstr(q.cell(rows[hit], c['機器名']))}(出力テキスト {q.vstr(q.cell(rows[hit], c['出力テキスト']))})", [hit]


def _equip_result(g: mr.Group, rows: list[list[Any]], fields: Sequence[str], ctx: dict) -> tuple[str, list[int]]:
    C = {n: q.col_of(fields, n) for n in ("管理番号", "作業可否")}
    a = [ctx.get(k, "") for k in mr.EQUIP_FIELDS]
    sub_rows, matching, text = q._equipment_new(g, rows, fields, C, q.vstr(a[0]), q.vstr(a[1]),
                                                *(q.csng(x or 0) for x in a[2:]))
    if not sub_rows:
        return "可(このラインの行なし)", []
    if not matching:
        return text, []
    out = ""
    for h in q._hit_rows(sub_rows, C["管理番号"], matching):
        if q.vstr(q.cell(h, C["作業可否"])) != "":
            out = q.vstr(q.cell(h, C["作業可否"]))
    used = [i for i, r in enumerate(rows) if any(r is x for x in sub_rows)]
    return out or "(作業可否が空)", used


def run(g: mr.Group, rows: list[list[Any]], fields: Sequence[str], ctx: dict) -> dict[str, Any]:
    special = {mr.MICRO: _micro_result, mr.EQUIP: _equip_result}.get(g.name)
    try:
        res = mr.match(g, rows, fields, ctx)
        if special is not None:
            text, used = special(g, rows, fields, ctx)
            if g.name == mr.EQUIP:
                used = [i for i in res.rows if i in used] if used else []
            res.rows = used
            res.result_text = text
    except Exception as e:  # noqa: BLE001 - VBA なら実行時エラー
        return {"error": f"{type(e).__name__}: {getattr(e, 'text', '') or e}", "steps": [], "result": "エラー"}
    steps = [{"order": s.rule.order, "rule": f"{','.join(s.rule.columns)} ← {s.rule.field} : {s.rule.compare}",
              "value": q.vstr(ctx.get(s.rule.field, "")), "count": len(s.remain),
              "rows": [_label(rows, fields, i) for i in s.remain[:ROW_CAP]]} for s in res.steps]
    return {"steps": steps, "matched": [_label(rows, fields, i) for i in res.matched[:ROW_CAP]],
            "used": [_label(rows, fields, i) for i in res.rows[:ROW_CAP]],
            "result": getattr(res, "result_text", None) or _result(g, rows, fields, res, ctx)}


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
        out.append((f"{args[0]} / {args[1]}", q.ratefuri_ctx(args, getattr(lv, "data", None))))
    return out


def _qp_rows(rows: list[list[Any]], fields: Sequence[str], head: str) -> list[int]:
    """DBQP の先の絞り込み: 検査番号の頭 LIKE '*頭*' OR = ''(表の行の位置)。"""
    c = q.col_of(fields, "検査番号の頭")
    out = []
    for i, r in enumerate(rows):
        v = q.vstr(q.cell(r, c)) if c else ""
        if v == "" or (head and head in v):
            out.append(i)
    return out


def _qp_inputs(lot: str) -> list[tuple[str, dict, str]]:
    """(見出し, 照合の入力, 検番の頭)。引当ごと(納入先・取引先)。"""
    lv = qs.read_lot(lot)
    data = getattr(lv, "data", None)
    if data is None or not data.records or not data.fields:
        return []
    head = q.qp_ken_head(lv.ken_no or "", data)
    pairs = list(dict.fromkeys((q.vstr(n), q.vstr(t)) for n, t in (q.hiki_arr_create(lot) or []))) or [("", "")]
    out = []
    for n, t in pairs:
        try:
            out.append((f"{n or '(納入先なし)'} / {t}(検番の頭 {head or '-'})", q.qp_ctx(data, n, t), head))
        except Exception:  # noqa: BLE001 - 寸法が数でない
            continue
    return out


def _table_lines(rows: list[list[Any]], fields: Sequence[str], g: mr.Group) -> list[str]:
    """表に在るライン(先に外すルール ライン の列の値。重なりなし)。"""
    out: list[str] = []
    for r in g.rules:
        if r.compare == "ライン":
            c = q.col_of(fields, r.columns[0])
            out += [q.vstr(q.cell(x, c)) for x in rows] if c else []
    return [x for x in dict.fromkeys(out) if x]


def _num_or(v: Any, dflt: Any) -> Any:
    s = q.vstr(v).strip()
    if s == "":
        return dflt
    try:
        return q.csng(s)
    except Exception:  # noqa: BLE001
        raise mr.Invalid(f"「{s}」は数ではありません") from None


def _micro_inputs(lot: str, lines_: list[str], over: Optional[dict] = None) -> list[tuple[str, dict]]:
    over = over or {}
    atu: Any = ""
    youto = ""
    if lot:
        lv = qs.read_lot(lot)
        atu, youto = lv.form("ATU"), lv.form("Youto")
    atu = _num_or(over.get("厚"), None) if q.vstr(over.get("厚")).strip() else _num_or(atu, None)
    if atu is None:
        return []
    if q.vstr(over.get("用途")).strip():
        youto = q.vstr(over.get("用途")).strip()
    lines_ = [over["ライン"]] if q.vstr(over.get("ライン")).strip() else lines_
    return [(f"{ln} / 厚 {vbx.cstr(atu, is_single=True)} / 用途 {youto or '(空)'}", {"ライン": ln, "厚": atu, "用途": youto})
            for ln in lines_]


def _equip_inputs(lot: str, lines_: list[str], over: Optional[dict] = None) -> list[tuple[str, dict]]:
    over = over or {}
    if not lot:
        return []
    lv = qs.read_lot(lot)
    args = getattr(lv.reflect, "equipment_args", None) if lv.reflect else None
    if not args:
        return []
    zai, atu, hab, tak, wei = args[1:]
    if q.vstr(over.get("オーダー厚")).strip():
        atu = _num_or(over.get("オーダー厚"), atu)
    lines_ = [over["ライン"]] if q.vstr(over.get("ライン")).strip() else lines_
    def sv(x: Any) -> str:
        return vbx.cstr(x, is_single=True) if isinstance(x, float) else q.vstr(x)
    return [(f"{ln} / {zai} {sv(atu)}×{sv(hab)}×{sv(tak)} 単重 {sv(wei)}",
             q.equipment_ctx(ln, zai, atu, hab, tak, wei)) for ln in lines_]


def _sika_inputs(lot: str) -> list[tuple[str, dict]]:
    """外観_厚板・外観_薄板・ラテラル: 仕掛の先頭行。"""
    lv = qs.read_lot(lot)
    data = getattr(lv, "data", None)
    if data is None or not data.records or not data.fields:
        return []
    ctx = q.sika_ctx(data)
    lab = " / ".join(q.vstr(ctx.get(k, "")) for k in ("用途ｺｰﾄﾞ", "製造材質", "製造調質"))
    dims = "×".join(q.vstr(ctx.get(k, "")) for k in ("製造板厚", "製造板幅", "製造板丈"))
    return [(f"{lab} 製造 {dims}(オーダー厚 {q.vstr(ctx.get('ｵｰﾀﾞｰ板厚', ''))})", ctx)]


def _bh_inputs(lot: str, over: Optional[dict] = None) -> list[tuple[str, dict]]:
    over = over or {}
    zai = tyo = ""
    if lot:
        lv = qs.read_lot(lot)
        zai, tyo = lv.form("Zai"), lv.form("Tyu")
    if q.vstr(over.get("材質")).strip():
        zai = q.vstr(over.get("材質")).strip()
    if q.vstr(over.get("調質")).strip():
        tyo = q.vstr(over.get("調質")).strip()
    if zai == "":
        return []                       # 材質が空なら BhSearch を呼ばない(wiring:1004)
    return [(f"材質 {zai} / 調質 {tyo or '(空)'}", {"材質": zai, "調質": tyo})]


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


def trial(name: str, lot: str, payload: Optional[dict], store: LocalStore,
          inputs: Optional[dict] = None) -> dict[str, Any]:
    saved = mr.groups(fresh=True)[name]
    new = edited(name, payload)
    _path, fields, rows = _table(saved)
    if name in (mr.MICRO, mr.EQUIP):
        fn = _micro_inputs if name == mr.MICRO else _equip_inputs
        inputs3 = [(lab, ctx, None) for lab, ctx in fn(lot, _table_lines(rows, fields, saved), inputs)]
    elif name == mr.QP:
        inputs3 = _qp_inputs(lot)
    elif name in SIKA_GROUPS:
        inputs3 = [(lab, ctx, None) for lab, ctx in _sika_inputs(lot)]
    elif name == mr.BH:
        inputs3 = [(lab, ctx, None) for lab, ctx in _bh_inputs(lot, inputs)]
    else:
        inputs3 = [(lab, ctx, None) for lab, ctx in
                   (_ratefuri_inputs(store, lot) if name == mr.RATEFURI else _pack_inputs(_hist(store), lot))]
    inputs = inputs3
    cases = []
    for label, ctx, head in inputs3:
        sub = rows if head is None else [rows[i] for i in _qp_rows(rows, fields, head)]
        a, b = run(saved, sub, fields, ctx), run(new, sub, fields, ctx)
        cases.append({"label": label, "saved": a, "edited": b, "changed": a["result"] != b["result"]})
    note = "" if inputs else ("厚が分かりません(ロットNo か 厚 を入れてください)" if name == mr.MICRO
                              else "仕掛にこのロットがありません(設備仕様はロットNo が要ります)" if name == mr.EQUIP
                              else "材質が分かりません(ロットNo か 材質 を入れてください)" if name == mr.BH
                              else "仕掛にこのロットがありません" if name == mr.QP or name in SIKA_GROUPS
                              else "引当がありません(仕掛・Lot情報履歴)" if name == mr.RATEFURI
                              else "Lot情報履歴 にこのロットの対象の包装仕様(蓋アングル)の行がありません")
    return {"lot": lot, "cases": cases, "note": note, "table_rows": len(rows)}


def regression(name: str, payload: Optional[dict], store: LocalStore) -> dict[str, Any]:
    """保存の前の確かめ。結果が変わる入力を並べる。"""
    saved = mr.groups(fresh=True)[name]
    new = edited(name, payload)
    _path, fields, rows = _table(saved)
    heads: dict[int, Optional[str]] = {}
    if name == mr.QP:
        from . import maint
        inputs = []
        for lot in maint.lot_numbers(plate_only=False)[:LOT_CAP]:
            for lab, ctx, head in _qp_inputs(lot):
                heads[len(inputs)] = head
                inputs.append((lot, lab, ctx))
    elif name in (mr.MICRO, mr.EQUIP):
        from . import maint
        fn = _micro_inputs if name == mr.MICRO else _equip_inputs
        lines_ = list(dict.fromkeys(_table_lines(rows, fields, saved) + _table_lines(rows, fields, new)))
        inputs = [(lot, lab, ctx) for lot in maint.lot_numbers(plate_only=False)[:LOT_CAP]
                  for lab, ctx in fn(lot, lines_)]
    elif name in SIKA_GROUPS or name == mr.BH:
        from . import maint
        inputs = [(lot, lab, ctx) for lot in maint.lot_numbers(plate_only=False)[:LOT_CAP]
                  for lab, ctx in (_sika_inputs(lot) if name in SIKA_GROUPS else _bh_inputs(lot))]
    elif name == mr.RATEFURI:
        from . import maint
        lots = maint.lot_numbers(plate_only=False)[:LOT_CAP]
        inputs = [(lot, lab, ctx) for lot in lots for lab, ctx in _ratefuri_inputs(store, lot)]
    else:
        hist = _hist(store)
        inputs = []
        for lab, ctx in _pack_inputs(hist, None):
            inputs.append((q.vstr(ctx.get("ﾛｯﾄNo", "")), lab, ctx))
    changed = []
    for n, (lot, lab, ctx) in enumerate(inputs):
        head = heads.get(n)
        sub = rows if head is None else [rows[i] for i in _qp_rows(rows, fields, head)]
        a, b = run(saved, sub, fields, ctx)["result"], run(new, sub, fields, ctx)["result"]
        if a != b:
            changed.append({"lot": lot, "label": lab, "before": a, "after": b})
    return {"checked": len(inputs), "changed": changed[:200], "changed_count": len(changed)}
