"""設定画面(この端末・使用パス・権限・不具合の扱い)

- この端末: VBA のレジストリ「設備名」と SF3(初期設定)の代わり。**未設定のときは誰でも決められる**
  (VBA も未設定なら SF3 を自動で出していた)。決まったあとに変えるのは管理者だけ(D12)
- 使用パス: VBA の Setting_Path と直書きのパス(ご依頼)。オンライン/オフラインの切替も。変えるのは管理者だけ
- 権限: アクセス権限の表(python-web-tools と共通。D12)を読んだ結果
- 不具合の扱い: `config/quirks.json` の中身(D3)

`GET /api/fs/list` はサーバ側のフォルダを一覧する(「参照…」。python-web-tools と同じ)。
"""
from __future__ import annotations

from flask import Blueprint, jsonify, render_template, request

from ufkensa import access_rights, auth, config, logging_utils, trace, user_settings
from ufkensa.kensa import paths, quirks, terminal
from ufkensa.logging_utils import get_logger
from ufkensa.presenters import fs_browse

from ..shell import shell_context

log = get_logger("app.routes.settings")

bp = Blueprint("settings", __name__)

MSG_ADMIN = auth.MSG_NO_LEVEL


def _err(code: str, message: str, status: int = 400):
    return jsonify({"error": {"code": code, "message": message}}), status


def _state() -> dict:
    s = auth.ensure()
    d = s.rights
    m = paths.mode()
    return {
        "admin": auth.can_edit(),
        "terminal": {
            "facility": terminal.facility(),
            "line_no": terminal.line_no(),
            "label": terminal.line_label(),
            "facilities": list(terminal.FACILITIES),
            "numbers": list(terminal.MARU_NUMBERS),
            "unset": terminal.line() is None,
        },
        "mode": m,
        "modes": [{"key": k, "label": v} for k, v in paths.MODE_LABELS.items()],
        "paths": {k: paths.overview(k) + [_log_path_row()] for k in paths.MODE_LABELS},
        "rights": {
            "user": s.user_id,
            "pc": auth.pc_name(),
            "label": s.label,
            "source": d.source if d else "",
            "reason": d.reason if d else "",
            "codes": d.codes if d else [],
            "table": access_rights.overview(),
            "known": [{"code": k, "label": v} for k, v in access_rights.CODES.items()],
        },
        "quirks": {"fix": quirks.summary()["fix"], "titles": quirks.TITLES},
        "log": _log_state(),
    }


LOG_KEY = "log"


def _log_path_row() -> dict:
    """使用パスの一覧に出す「ログの出力パス」の行(中身は 設定 → ログ と同じ設定。モードによらず1つ)。"""
    requested = str(user_settings.get(config.KEY_LOG_DIR, "") or "")
    return {"key": LOG_KEY, "label": "ログの出力パス", "vba": "(新規。all-tools と同じ決まり)",
            "use": "この下に ufkensa フォルダを作り ufkensa_<PC名>_日付.log / ufkensa_記録_<PC名>_日付.jsonl を書く。"
                   "空ならこの端末のローカル(logs)。オンライン/オフライン共通",
            "need": "write", "kind": "dir", "path": requested, "default": "", "changed": bool(requested)}


def _set_log_dir(text: str):
    """ログの出力パスを変える。書けなければ変えずに理由を返す(今までの置き場所のまま)。"""
    text = (text or "").strip().strip('"')
    target = logging_utils.resolve_log_dir(text)
    why = logging_utils.set_log_dir(target)
    if why:
        return _err("not_writable", f"そのフォルダに書けません: {target}({why})。今までの置き場所のままです。")
    user_settings.save(config.KEY_LOG_DIR, text)
    trace.event("settings.log", f"ログの置き場所を {target} にしました")
    return None


def _log_state() -> dict:
    """ログ(all-tools と同じ決まり): <出力パス>\\ufkensa\\ufkensa_<PC>_日付.log と ufkensa_記録_<PC>_日付.jsonl。"""
    requested = str(user_settings.get(config.KEY_LOG_DIR, "") or "")
    problems = []
    try:
        problems = [{"t": r.get("t", "")[:19].replace("T", " "), "lv": r.get("lv", ""), "ev": r.get("ev", ""),
                     "msg": str(r.get("msg", ""))[:200], "proc": r.get("proc", ""), "ms": r.get("ms")}
                    for r in trace.read(7, only_problems=True, limit=30)]
    except Exception:                                  # noqa: BLE001 - 読めなくても設定画面は出す
        pass
    return {
        "dir": str(config.LOG_DIR), "default": str(config.DEFAULT_LOG_DIR), "requested": requested,
        "problem": logging_utils.probe_writable(config.LOG_DIR),
        "keep_days": int(user_settings.get(config.KEY_LOG_KEEP_DAYS, config.DEFAULT_LOG_KEEP_DAYS) or 0),
        "tool_id": config.TOOL_LOG_NAME, "pc": trace.pc_name(), "problems": problems,
    }


@bp.get("/settings")
def page():
    return render_template("settings.html", **shell_context("settings"))


@bp.get("/api/settings/state")
def state():
    return jsonify(_state())


@bp.post("/api/settings/terminal")
def save_terminal():
    body = request.get_json(silent=True) or {}
    if terminal.line() is not None and not auth.can_edit():
        return _err("not_admin", MSG_ADMIN, 403)
    try:
        terminal.set_facility(str(body.get("facility", "")), str(body.get("line_no", "") or ""))
    except ValueError as e:
        return _err("bad_value", str(e))
    log.info("この端末の設備名を %s(%s)にしました", terminal.facility(), terminal.line_no())
    return jsonify(_state())


@bp.post("/api/settings/mode")
def save_mode():
    if not auth.can_edit():
        return _err("not_admin", MSG_ADMIN, 403)
    body = request.get_json(silent=True) or {}
    try:
        paths.set_mode(str(body.get("mode", "")))
    except ValueError:
        return _err("bad_value", "オンラインかオフラインを選んでください")
    log.info("モードを %s にしました", paths.mode())
    return jsonify(_state())


@bp.post("/api/settings/path")
def save_path():
    if not auth.can_edit():
        return _err("not_admin", MSG_ADMIN, 403)
    body = request.get_json(silent=True) or {}
    key, mode = str(body.get("key", "")), str(body.get("mode", "") or paths.mode())
    if key == LOG_KEY:
        refused = _set_log_dir("" if body.get("reset") else str(body.get("path", "")))
        return refused or jsonify(_state())
    if key not in paths.BY_KEY or mode not in paths.MODE_LABELS:
        return _err("bad_value", "その項目はありません")
    if body.get("reset"):
        paths.reset_path(key, mode)
    else:
        paths.set_path(key, str(body.get("path", "")), mode)
    log.info("使用パス %s(%s)を %s にしました", key, mode, paths.get(key, mode))
    return jsonify(_state())


@bp.post("/api/settings/path/check")
def check_path():
    body = request.get_json(silent=True) or {}
    key, mode = str(body.get("key", "")), str(body.get("mode", "") or paths.mode())
    if key == LOG_KEY:
        target = logging_utils.resolve_log_dir(str(user_settings.get(config.KEY_LOG_DIR, "") or ""))
        why = logging_utils.probe_writable(target)
        return jsonify({"key": key, "path": str(target), "exists": target.is_dir(), "readable": not why,
                        "writable": not why, "ok": not why,
                        "message": f"書けます({target})" if not why else f"書けません: {why}"})
    if key not in paths.BY_KEY or mode not in paths.MODE_LABELS:
        return _err("bad_value", "その項目はありません")
    return jsonify(paths.check(key, mode))


@bp.post("/api/settings/rights/reload")
def reload_rights():
    auth.end()
    auth.ensure()
    return jsonify(_state())


@bp.get("/api/fs/list")
def fs_list():
    """サーバ側のフォルダを一覧する(「参照…」)。"""
    view = fs_browse.browse(request.args.get("path", ""), access=request.args.get("access") == "1")
    return jsonify(fs_browse.to_dict(view))


@bp.post("/api/settings/log")
def save_log():
    """ログの出力パスと残す日数。出力パスの下に ufkensa フォルダを作って書く(直下には置かない)。"""
    if not auth.can_edit():
        return _err("not_admin", MSG_ADMIN, 403)
    body = request.get_json(silent=True) or {}
    if "keep_days" in body:
        try:
            days = int(body["keep_days"])
        except (TypeError, ValueError):
            return _err("bad_value", "残す日数は 0 以上の数で入れてください(0 = 消さない)")
        if days < 0:
            return _err("bad_value", "残す日数は 0 以上の数で入れてください(0 = 消さない)")
        user_settings.save(config.KEY_LOG_KEEP_DAYS, days)
    if "dir" in body:
        refused = _set_log_dir(str(body.get("dir") or ""))
        if refused:
            return refused
    return jsonify(_state())
