"""マスタ編集の認証(管理者パスワード)。2026-11-02 の決め

マスタ管理・保守の画面は誰でも開けるが、**編集は 管理者パスワード と アクセス権限の「管理者」の両方**が要る
(python-web-tools の 設定「パスワード」→「マスタ編集の認証」と同じ考え方)。本体は `ufkensa.auth`
(unlock_master / lock_master / master_edit_ok / master_edit_reason)と `ufkensa.admin_password`。

- POST /api/admin/unlock   {password}                 → {ok, reason, state}(違っても 200。ok=false)
- POST /api/admin/lock                                → {state}
- GET  /api/admin/state                               → {unlocked, can_edit, admin, reason, custom, min_length}
- POST /api/admin/password {current, new, confirm}    → {ok, message, state}(変えるには いまのパスワード が要る)

**パスワードの値はログに残さない**(trace にも log にも渡さない)。続けて間違えたら少し待たせる。
"""
from __future__ import annotations

import threading
import time

from flask import Blueprint, g, jsonify, request

from ufkensa import admin_password, auth, trace
from ufkensa.logging_utils import get_logger

log = get_logger("app.routes.admin_auth")

bp = Blueprint("admin_auth", __name__)

#: 続けて間違えたときの待ち(admin_password.Session と同じ値)
_miss_lock = threading.Lock()
_misses = 0
_wait_until = 0.0


def state_dict() -> dict:
    return {"unlocked": auth.master_unlocked(), "can_edit": auth.master_edit_ok(), "admin": auth.can_edit(),
            "reason": auth.master_edit_reason(), "custom": admin_password.is_custom(),
            "min_length": admin_password.MIN_LENGTH}


def _body() -> dict:
    """送られた値。記録(リクエストのログ)には値を伏せたものを残す(`g.trace_sent`)。"""
    b = request.get_json(silent=True) or {}
    g.trace_sent = {k: "***" for k in b} if isinstance(b, dict) else "***"
    return b if isinstance(b, dict) else {}


def _err(code: str, message: str, status: int = 400, field: str = ""):
    e = {"code": code, "message": message}
    if field:
        e["field"] = field
    return jsonify({"error": e, "state": state_dict()}), status


def reset_misses() -> None:
    """試験用: 間違いの数え直し。"""
    global _misses, _wait_until
    with _miss_lock:
        _misses, _wait_until = 0, 0.0


@bp.get("/api/admin/state")
def state():
    return jsonify(state_dict())


@bp.post("/api/admin/unlock")
def unlock():
    global _misses, _wait_until
    password = str(_body().get("password", "") or "")
    with _miss_lock:
        now = time.monotonic()
        if now < _wait_until:
            wait = int(_wait_until - now) + 1
            return jsonify({"ok": False, "reason": f"続けて間違えたので、{wait} 秒待ってください", "state": state_dict()})
    with trace.timed("admin.unlock", "管理者: マスタ編集の認証") as info:
        ok, reason = auth.unlock_master(password)
        info.update(ok=ok, reason=reason or None)
    with _miss_lock:
        if ok:
            _misses = 0
        elif auth.can_edit():                       # 権限が無くて断ったときは数えない
            _misses += 1
            if _misses >= admin_password.MAX_MISSES:
                _misses, _wait_until = 0, time.monotonic() + admin_password.COOL_DOWN_SEC
    if not ok:
        log.info("マスタ編集の認証を断りました: %s", reason)
    return jsonify({"ok": ok, "reason": reason, "state": state_dict()})


@bp.post("/api/admin/lock")
def lock():
    with trace.timed("admin.lock", "管理者: マスタ編集の鍵を閉じる"):
        auth.lock_master()
    return jsonify({"state": state_dict()})


@bp.post("/api/admin/password")
def change_password():
    """管理者パスワードを変える。アクセス権限の「管理者」と、いまのパスワードが要る。"""
    if not auth.can_edit():
        return _err("not_admin", "パスワードを変えるには、アクセス権限の「管理者」が要ります", 403)
    b = _body()
    with trace.timed("admin.password", "管理者: 管理者パスワードを変える") as info:
        res = admin_password.change(str(b.get("current", "") or ""), str(b.get("new", "") or ""),
                                    str(b.get("confirm", "") or ""))
        info.update(ok=res.ok, reason=res.reason or None)
    if not res.ok:
        field = {"wrong_password": "current", "too_short": "new", "same": "new", "mismatch": "confirm"}.get(res.reason, "")
        return _err(res.reason or "bad_value", res.message, 400, field)
    log.info("管理者パスワードを変えました")
    return jsonify({"ok": True, "message": res.message, "state": state_dict()})
