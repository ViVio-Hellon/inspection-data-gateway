"""保守画面(管理者だけ。VBA の Test フォームのうち業務で要る 2 つと、初期設定.bas の端末設定)

- 歪公差チェック: 品質基準の表を変えたあとの確認(Test.CommandButton34)。ロット 1 つ / 仕掛の全ロット。
  別の 品質基準.sqlite3 を指定すると、いまの表と並べて違いに印を付ける。ソリ公差(CommandButton24)・
  設備仕様(CommandButton38)の表引きも。**DB は読むだけ**
- 端末設定のセット: 初期設定.AIM / HVC / DelSet(レジストリ → user_settings)
組み立ては `ufkensa/kensa/maint.py`。判断はすべてサーバ。
"""
from __future__ import annotations

import re

from flask import Blueprint, jsonify, render_template, request

from ufkensa import auth, trace
from ufkensa.kensa import maint, terminal
from ufkensa.kensa.flow import Stop
from ufkensa.logging_utils import get_logger

from ..shell import shell_context

log = get_logger("app.routes.maint")

bp = Blueprint("maint", __name__)

_LOT_RE = re.compile(r"^[0-9A-Za-z]{1,20}$")
_PREFIX_RE = re.compile(r"^[0-9A-Za-z]{0,20}$")


def _err(code: str, message: str, status: int = 400, field: str = ""):
    e = {"code": code, "message": message}
    if field:
        e["field"] = field
    return jsonify({"error": e}), status


def _deny():
    """管理者でなければ 403(D12)。"""
    if not auth.can_edit():
        return _err("not_admin", auth.MSG_NO_LEVEL, 403)
    return None


def _body() -> dict:
    return request.get_json(silent=True) or {}


def _alt(body: dict):
    """別の品質基準ファイル(空なら None)。指定が悪ければ ValueError。"""
    return maint.resolve_alt(str(body.get("alt", "") or ""))


@bp.get("/maint")
def page():
    return render_template("maint.html", **shell_context("maint"))


@bp.get("/api/maint/state")
def state():
    return jsonify({"admin": auth.can_edit(), "terminal": maint.terminal_state(),
                    "quality_db": str(maint.current_quality_db()), "facilities": list(terminal.FACILITIES),
                    "labels": maint.DIST_LABELS, "keys": list(maint.DIST_KEYS)})


@bp.post("/api/maint/dist")
def dist():
    """歪公差チェック(ロット 1 つ)。Test.CommandButton34_Click。"""
    if (d := _deny()) is not None:
        return d
    body = _body()
    lot = str(body.get("lot", "") or "").strip()
    if not _LOT_RE.match(lot):
        return _err("bad_value", "ロットNo は英数字で入れてください", field="lot")
    try:
        alt = _alt(body)
    except ValueError as e:
        return _err("bad_alt", str(e), field="alt")
    with trace.timed("maint.dist", "保守: 歪公差チェック(1ロット)", expected=(Stop,),
                     lot=lot.upper(), alt=str(alt) if alt else None) as info:
        row = maint.dist_check(lot, alt)
        info["status"] = row["current"]["status"]
        info["diff"] = row["diff"] or None
    return jsonify({"row": row, "quality_db": str(maint.current_quality_db()), "alt": str(alt) if alt else ""})


@bp.post("/api/maint/dist_all")
def dist_all():
    """歪公差チェック(仕掛の全ロット)。表を変えたあとの確認。"""
    if (d := _deny()) is not None:
        return d
    body = _body()
    prefix = str(body.get("prefix", "") or "").strip()
    if not _PREFIX_RE.match(prefix):
        return _err("bad_value", "ロットNo の頭は英数字で入れてください", field="prefix")
    try:
        alt = _alt(body)
    except ValueError as e:
        return _err("bad_alt", str(e), field="alt")
    with trace.timed("maint.dist_all", "保守: 歪公差チェック(仕掛の全ロット)", expected=(Stop,),
                     prefix=prefix or None, alt=str(alt) if alt else None) as info:
        out = maint.dist_check_all(alt, plate_only=bool(body.get("plate_only", True)), prefix=prefix)
        info.update(count=out["count"], diffs=out["diffs"], by_status=out["by_status"])
    out["quality_db"] = str(maint.current_quality_db())
    return jsonify(out)


@bp.post("/api/maint/sori")
def sori():
    """ソリ公差(ソリ4級)。Test.CommandButton24_Click。"""
    if (d := _deny()) is not None:
        return d
    body = _body()
    try:
        alt = _alt(body)
        with trace.timed("maint.sori", "保守: ソリ公差チェック", expected=(Stop, ValueError)):
            cur = maint.sori_check(body.get("atu", ""), body.get("haba", ""), body.get("take", ""))
            other = maint.sori_check(body.get("atu", ""), body.get("haba", ""), body.get("take", ""), alt) if alt else None
    except ValueError as e:
        return _err("bad_value", str(e))
    return jsonify({"current": cur, "alt": other})


@bp.post("/api/maint/equipment")
def equipment():
    """設備仕様(伝送用ファイル)。Test.CommandButton38_Click。"""
    if (d := _deny()) is not None:
        return d
    b = _body()
    try:
        with trace.timed("maint.equipment", "保守: 設備仕様チェック", expected=(Stop, ValueError)):
            out = maint.equipment_check(str(b.get("line", "")), str(b.get("zai", "")), b.get("atu", ""),
                                        b.get("haba", ""), b.get("take", ""), b.get("wei", ""))
    except ValueError as e:
        return _err("bad_value", str(e))
    return jsonify(out)


@bp.post("/api/maint/terminal")
def set_terminal():
    """端末設定のセット(初期設定.AIM / HVC / DelSet)。"""
    if (d := _deny()) is not None:
        return d
    action = str(_body().get("action", ""))
    before = maint.terminal_state()
    desc = {"AIM": "AIM のセット", "HVC": "HVC のセット", "clear": "端末設定を消す(DelSet)"}.get(action)
    if desc is None:
        return _err("bad_value", f"知らない操作です: {action}")
    with trace.timed("maint.terminal", f"保守: {desc}", expected=(ValueError,), action=action,
                     before={"設備名": before["facility_text"], "号機": before["line_no"], "モード": before["mode_text"]}) as info:
        after = maint.clear_terminal() if action == "clear" else maint.apply_preset(action)
        info["after"] = {"設備名": after["facility_text"], "号機": after["line_no"], "モード": after["mode_text"]}
    log.info("保守: %s 設備名 %s → %s、モード %s → %s", desc, before["facility_text"], after["facility_text"],
             before["mode_text"], after["mode_text"])
    return jsonify({"terminal": after})
