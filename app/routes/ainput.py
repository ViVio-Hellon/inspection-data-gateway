"""Aインプット画面(VBA の UFAinput。梱包高さチェックとアイカス送信)

検査画面の「Aインプット」ボタン(/ainput へのリンク)から開く。いま検査中のロット・Min/Max・検査員は
検査画面のセッションから読む(`ainput_session`)。

画面の操作 1 つ = `ainput_session.AinputSession` のメソッド 1 つ。返り値は検査画面と同じ Result
(view・messages・question・stop)。判断はすべてセッション(業務ロジック `ainput.py`)が行い、ここは受け渡すだけ。
"""
from __future__ import annotations

from flask import Blueprint, jsonify, render_template, request

from ufkensa import trace
from ufkensa.kensa import ainput_session
from ufkensa.kensa.flow import Stop
from ufkensa.logging_utils import get_logger

from ..shell import shell_context

log = get_logger("app.routes.ainput")

bp = Blueprint("ainput", __name__)

#: 画面の操作 → (セッションのメソッド, 受け取る引数, ログの説明)
OPS: dict[str, tuple[str, tuple[str, ...], str]] = {
    "open": ("open", ("ken_m", "source"), "開く(Ainput_Startup。source = 呼び元 Cask / Mitsubishi / H494)"),
    "value": ("set_value", ("name", "value"), "入力欄"),
    "spin": ("spin", ("name", "dir"), "段・風袋の欄を増減"),
    "mount": ("set_mount", ("value",), "山数"),
    "receive": ("receive", ("text",), "山高さの伝送"),
    "clear-height": ("clear_height", ("n",), "山高さを消す"),
    "press": ("press", ("answers",), "Icas出力"),
    "unlock": ("unlock", (), "外注出荷"),
    "reset": ("reset", ("answers",), "Reset(hostsend の書き直し)"),
}


@bp.get("/ainput")
def page():
    return render_template("ainput.html", **shell_context("ainput"))


@bp.get("/api/ainput/state")
def state():
    s = ainput_session.get()
    return jsonify({"view": s.view(), "messages": [], "question": None, "stop": None})


@bp.post("/api/ainput/<op>")
def operate(op: str):
    if op not in OPS:
        return jsonify({"error": {"code": "not_found", "message": "その操作はありません"}}), 404
    method, names, desc = OPS[op]
    body = request.get_json(silent=True) or {}
    kwargs = {n: body[n] for n in names if n in body}
    s = ainput_session.get()
    lot = s.form.lot if s.form is not None else None
    with trace.timed(f"ainput.{op}", f"Aインプット: {desc}", expected=(Stop,), lot=lot or None):
        result = getattr(s, method)(**kwargs)
    return jsonify(result.to_dict())
