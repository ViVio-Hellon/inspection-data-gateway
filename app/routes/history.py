"""履歴画面(VBA の UF管理)

この端末の手元DB の「保存用」(耳付は「耳用」)に貯まった検査を一覧にし、選んだ検査の中身を表示する。
VBA の UF管理 と同じ関数(records.kanri_lots / kanri_view)を使う。検査画面への復元(UF管理 の CommandButton1)は
検査画面のセッションにつなぐ(結合 工程 8c)。
"""
from __future__ import annotations

from flask import Blueprint, jsonify, render_template, request

from ufkensa import trace
from ufkensa.kensa import records, store as st
from ufkensa.kensa.records import FACES
from ufkensa.logging_utils import get_logger

from ..shell import shell_context

log = get_logger("app.routes.history")

bp = Blueprint("history", __name__)


def _store() -> st.LocalStore:
    return st.LocalStore()


def _mimi() -> bool:
    return request.args.get("mimi") == "1"


@bp.get("/history")
def page():
    return render_template("history.html", **shell_context("history"))


@bp.get("/api/history/lots")
def lots():
    """一覧: 行(2〜)・検査開始終了(A 列)・LOT(B 列)。新しいものが上。"""
    s = _store()
    try:
        name = st.MIMI if _mimi() else st.HOZON
        if not s.exists(name):
            return jsonify({"rows": [], "sheet": name})
        sh = s.sheet(name)
        last = sh.last_row(2)
        with trace.timed("db.read", f"{name} の一覧を読む", path=str(s.path)) as info:
            raw = sh.read_range(2, 1, last, 5) if last >= 2 else []
            info["rows"] = len(raw)
        rows = [{"row": i + 2, "time": st.cell_text(r[0]), "lot": st.cell_text(r[1]),
                 "atu": st.cell_text(r[2]), "hab": st.cell_text(r[3]), "tak": st.cell_text(r[4])}
                for i, r in enumerate(raw)]
        return jsonify({"rows": list(reversed(rows)), "sheet": name})
    finally:
        s.close()


@bp.get("/api/history/view")
def view():
    """選んだ行の中身(UF管理 の CB1_Change と同じ並べ方)。"""
    try:
        row = int(request.args.get("row", "0"))
    except ValueError:
        row = 0
    if row < 2:
        return jsonify({"error": {"code": "bad_value", "message": "行を選んでください"}}), 400
    s = _store()
    try:
        mimi = _mimi()
        lots_ = records.kanri_lots(s, mimi=mimi)
        if row - 2 >= len(lots_):
            return jsonify({"error": {"code": "not_found", "message": "その行はありません"}}), 404
        lot = lots_[row - 2]
        kv = records.kanri_view(s, row - 2, lots_, lot, mimi=mimi, cb1_tag="ループ")
        v = kv.values
        grid = [{"no": i, "cells": [v.get(f"{f}{i}", "") for f in FACES]} for i in range(1, 15)
                if any(v.get(f"{f}{i}", "") for f in FACES)]
        extras = {
            "ラテラル": [v.get(f"{f}late", "") for f in FACES],
            "直角度": [v.get(f"{f}_Angle", "") for f in FACES],
            "BH": [v.get(f"{f}_BH", "") for f in FACES],
            "定盤歪 A": [v.get(f"{f}ZH", "") for f in FACES],
            "ゲージ歪 A": [v.get(f"{f}GH", "") for f in FACES],
            "定盤歪 B": [v.get(f"{f}ZH_B", "") for f in FACES],
            "ゲージ歪 B": [v.get(f"{f}GH_B", "") for f in FACES],
            "外観 A": [v.get(f"A{f.lower()}_Vis", "") for f in FACES],
            "外観 B": [v.get(f"B{f.lower()}_Vis", "") for f in FACES],
        }
        extras = {k: x for k, x in extras.items() if any(x)}
        known = {f"{f}{i}" for f in FACES for i in range(1, 15)}
        others = {k: x for k, x in v.items() if k not in known and x and not k.endswith(("late", "_Angle", "_BH", "_Vis"))
                  and not k[1:] in ("ZH", "GH", "ZH_B", "GH_B")}
        return jsonify({"lot": lot, "row": row, "faces": list(FACES), "grid": grid, "extras": extras, "others": others,
                        "messages": [m.to_dict() for m in kv.messages]})
    finally:
        s.close()
