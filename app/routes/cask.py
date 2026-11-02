"""キャスク検査の画面(VBA の UFCask + キャスク検査シート)

画面の操作 1 つ = `cask_session.CaskSession` のメソッド 1 つ。返り値は `Result`(view・messages・question・stop)。
判断はすべてセッション(部品 modes/cask.py・modes/sheet.py)が行い、ここは受け渡すだけ。
"""
from __future__ import annotations

import re
from pathlib import Path

from flask import Blueprint, jsonify, render_template, request, send_file

from ufkensa import trace
from ufkensa.kensa import cask_session
from ufkensa.kensa.flow import Stop
from ufkensa.logging_utils import get_logger

from ..shell import shell_context

log = get_logger("app.routes.cask")

bp = Blueprint("cask", __name__)

#: 画面の操作 → (セッションのメソッド, 受け取る引数)
OPS: dict[str, tuple[str, tuple[str, ...]]] = {
    "open": ("open", ()),
    "lot": ("set_lot", ("lot",)),
    "field": ("set_field", ("name", "value")),
    "zero": ("zero", ("text",)),
    "receive": ("receive", ("text",)),
    "select": ("select", ("row", "col")),
    "transfer": ("transfer", ()),
    "print": ("print_out", ("answers",)),
    "save": ("save", ()),
    "new": ("new_sheet", ("answers",)),
    "cancel": ("cancel", ()),
    "inspection": ("toggle_inspection", ()),
    "resume": ("resume", ()),
    "proofread": ("proofread", ()),
    "close": ("close", ()),
}


@bp.get("/cask")
def page():
    return render_template("cask.html", **shell_context("cask"))


@bp.get("/api/cask/state")
def state():
    return jsonify(cask_session.get().state().to_dict())


@bp.get("/api/cask/report/<path:name>")
def report(name: str):
    """印刷した帳票(D8: ブラウザで開いて印刷する HTML)。手元のデータフォルダの print にあるものだけ。"""
    s = cask_session.get()
    safe = Path(name).name
    if not re.fullmatch(r"[^\\/]+\.html", safe):
        return jsonify({"error": {"code": "bad_value", "message": "帳票の名前が正しくありません"}}), 400
    f = s.store.path.parent / "print" / safe
    if not f.is_file():
        return jsonify({"error": {"code": "not_found", "message": f"帳票がありません: {safe}"}}), 404
    return send_file(f, mimetype="text/html", max_age=0)


@bp.post("/api/cask/<op>")
def operate(op: str):
    if op not in OPS:
        return jsonify({"error": {"code": "not_found", "message": "その操作はありません"}}), 404
    method, names = OPS[op]
    body = request.get_json(silent=True) or {}
    kwargs = {n: body[n] for n in names if n in body}
    s = cask_session.get()
    lot = ""
    if s.book is not None:
        lot = s.book.cask.lot_box or None
    missing = [n for n in names if n != "answers" and n not in kwargs]
    if missing:
        return jsonify({"error": {"code": "bad_value", "message": "足りない値があります: " + ", ".join(missing)}}), 400
    with trace.timed(f"cask.{op}", f"キャスク検査: {op}", expected=(Stop,), lot=lot or None):
        result = getattr(s, method)(**kwargs)
    return jsonify(result.to_dict())
