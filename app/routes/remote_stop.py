"""管理者以外停止モード(`remote_stop.py` / `ufkensa/remote_stop_link.py`)

- GET  /api/remote_stop/state         いまのお願いと、各 PC の様子(誰でも見られる。業務のデータは含まない)
- POST /api/remote_stop/issue {reason} お願いを出す。管理者パスワード(保守の鍵)**かつ** Administrator
- POST /api/remote_stop/clear         解除する(同じ条件)
- POST /api/remote_stop/drafts {page, drafts:[{label, value}]}
                                      止めると決まった画面が、打ちかけを値ごと送る → 手元の控え

鍵が閉じていれば 403 {code: "locked"}(画面は鍵の窓を出して呼び直す。admin_unlock.js)。
Administrator が無ければ 403 {code: "not_administrator"}(呼び直さない)。
"""
from __future__ import annotations

from flask import Blueprint, jsonify, request

from ufkensa import remote_stop_link as link
from ufkensa.logging_utils import get_logger

log = get_logger("app.routes.remote_stop")

bp = Blueprint("remote_stop", __name__)


def _err(code: str, message: str, status: int):
    return jsonify({"error": {"code": code, "message": message}}), status


@bp.get("/api/remote_stop/state")
def state():
    return jsonify(link.overview())


def _guarded(action):
    w = link.who()
    if not w["password_ok"]:
        return _err("locked", "管理者以外停止には、管理者パスワードが要ります", 403)
    result = action()
    if not result.ok:
        status = 403 if result.reason in ("wrong_password", "not_administrator") else 409
        code = result.reason or "refused"
        return _err(code, result.message, status)
    return jsonify({**link.note_for(result), "state": link.overview()})


@bp.post("/api/remote_stop/issue")
def issue():
    reason = str((request.get_json(silent=True) or {}).get("reason") or "")
    return _guarded(lambda: link.issue(reason))


@bp.post("/api/remote_stop/clear")
def clear():
    return _guarded(link.clear)


@bp.post("/api/remote_stop/drafts")
def drafts():
    body = request.get_json(silent=True) or {}
    n = link.save_drafts(str(body.get("page") or ""), body.get("drafts"))
    return jsonify({"ok": True, "kept": n})
