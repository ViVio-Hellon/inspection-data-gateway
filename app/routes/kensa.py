"""検査画面(A案「計測コックピット」。D6)

画面の操作 1 つ = `session.Session` のメソッド 1 つ。返り値は `Result`(view・messages・question・stop)。
判断はすべてセッション(業務ロジック)が行い、ここは受け渡すだけ。
"""
from __future__ import annotations

from flask import Blueprint, jsonify, render_template, request

from ufkensa.kensa import session, view
from ufkensa.kensa.flow import Stop
from ufkensa.logging_utils import get_logger

from ..shell import shell_context

log = get_logger("app.routes.kensa")

bp = Blueprint("kensa", __name__)

#: 画面の操作 → (セッションのメソッド, 受け取る引数)
OPS: dict[str, tuple[str, tuple[str, ...]]] = {
    "start": ("start", ()),
    "worker": ("register_worker", ("card", "answers")),
    "lot": ("register_lot", ("lot", "answers")),
    "work": ("select_work", ("work",)),
    "yama": ("set_yama", ("yama",)),
    "at": ("toggle_at", ()),
    "inspe": ("toggle_inspe", ()),
    "receive": ("receive", ("text",)),
    "click": ("click", ("name",)),
    "dblclick": ("dblclick", ("name", "answers", "reason")),
    "value": ("set_value", ("name", "value")),
    "minmax": ("min_max", ()),
    "print-check": ("print_check", ("answers",)),
    "print-commit": ("print_commit", ("answers",)),
    "new": ("new_inspection", ("answers",)),
}

MSG_NOT_READY = "この操作はまだ結合中です(工程 8c)。いまは VBA の UF検査判定.xlsm を使ってください。"


@bp.get("/kensa")
def page():
    return render_template("kensa.html", **shell_context("kensa"))


@bp.get("/api/kensa/state")
def state():
    s = session.get()
    return jsonify({"view": view.build(s.state), "messages": [], "question": None, "stop": None})


@bp.post("/api/kensa/<op>")
def operate(op: str):
    if op not in OPS:
        return jsonify({"error": {"code": "not_found", "message": "その操作はありません"}}), 404
    method, names = OPS[op]
    body = request.get_json(silent=True) or {}
    kwargs = {n: body[n] for n in names if n in body}
    s = session.get()
    try:
        result = getattr(s, method)(**kwargs)
    except NotImplementedError:
        return jsonify({"error": {"code": "not_ready", "message": MSG_NOT_READY}}), 501
    except Stop as e:                                   # 念のため: セッションが Result に包まずに投げたとき
        return jsonify({"view": view.build(s.state), "messages": [m.to_dict() for m in e.messages],
                        "question": None, "stop": e.to_dict()["stop"]})
    return jsonify(result.to_dict())
