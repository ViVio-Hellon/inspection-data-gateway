"""保守 → 照合ルール(2026-11-11)。品質基準の照合のルール(伝送用ファイル「照合ルール」)を見る・試す・直す。

見る・試し照合・保存の前の確かめ は誰でも。保存は編集の鍵(`auth.master_edit_ok()`)。
組み立ては ufkensa/kensa/match_rules_maint.py。
"""
from __future__ import annotations

import re
import sqlite3

from flask import Blueprint, jsonify, request

from ufkensa import auth
from ufkensa.kensa import match_rules as mr, match_rules_maint as mm, store as st
from ufkensa.logging_utils import get_logger

log = get_logger("app.routes.match_rules")

bp = Blueprint("match_rules", __name__)

_LOT_RE = re.compile(r"^[0-9A-Za-z]{1,20}$")


def _err(code: str, message: str, status: int = 400, **extra):
    return jsonify({"error": {"code": code, "message": message, **extra}}), status


def _body() -> dict:
    return request.get_json(silent=True) or {}


def _group(b: dict):
    name = str(b.get("group", ""))
    if name not in mr.BY_NAME:
        return None, _err("bad_request", "ルール群がありません")
    return name, None


@bp.get("/api/maint/match_rules")
def get_state():
    return jsonify(mm.state())


@bp.post("/api/maint/match_rules/trial")
def trial():
    b = _body()
    name, e = _group(b)
    if e:
        return e
    lot = str(b.get("lot", "")).strip().upper()
    if not _LOT_RE.match(lot):
        return _err("bad_lot", "ロットNo は英数字 20 文字までです", field="lot")
    try:
        return jsonify(mm.trial(name, lot, b.get("edited"), st.LocalStore()))
    except mr.Invalid as x:
        return _err("invalid", str(x), order=x.order)


@bp.post("/api/maint/match_rules/regression")
def regression():
    b = _body()
    name, e = _group(b)
    if e:
        return e
    try:
        return jsonify(mm.regression(name, b.get("edited"), st.LocalStore()))
    except mr.Invalid as x:
        return _err("invalid", str(x), order=x.order)


@bp.post("/api/maint/match_rules")
def save():
    if not auth.master_edit_ok():
        return _err("locked", auth.master_edit_reason() or auth.MSG_NO_LEVEL, 403)
    b = _body()
    name, e = _group(b)
    if e:
        return e
    reason = str(b.get("reason", "") or "").strip()[:200]
    try:
        g = mm.edited(name, b.get("edited"))
        reg = mm.regression(name, b.get("edited"), st.LocalStore())
        if reg["changed_count"] and not reason:
            return _err("need_reason", f"結果が変わるロットが {reg['changed_count']} 件あります。保存の理由を書いてください",
                        field="reason")
        mr.save(g, reason=reason)
    except mr.Invalid as x:
        return _err("invalid", str(x), order=x.order)
    except (OSError, sqlite3.Error) as x:
        return _err("not_saved", f"保存できませんでした: {x}", 500)
    log.info("保守: 照合ルール「%s」を保存(理由: %s)", name, reason)
    return jsonify(mm.state())
