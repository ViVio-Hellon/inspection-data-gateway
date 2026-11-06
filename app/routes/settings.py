"""設定画面(この端末・使用パス・権限・不具合の扱い)

- この端末: VBA のレジストリ「設備名」と SF3(初期設定)の代わり。**未設定のときは誰でも決められる**
  (VBA も未設定なら SF3 を自動で出していた)。決まったあとに変えるのは管理者だけ(D12)
- 使用パス: VBA の Setting_Path と直書きのパス(ご依頼)。D26: 仕掛台帳だけ 第1パス / 第2パス(+ 2つ目の名前)、
  ほかはパス 1 つ。**変えるには `auth.master_edit_ok()`(管理者パスワード + アクセス権限の管理者)が要る**。
  アクセス権限マスタから決まらないあいだ(未設定・読めない・管理者の行が無い)は管理者パスワードだけでよい(D30 設定の途中)。
  断るときは 403(code "locked" = パスワード待ち、"not_admin" = 管理者でない。文言は `auth.master_edit_reason()`)。
  「既定に戻す」(項目ごと)・「すべて既定に戻す」も同じ鍵(`POST /api/settings/path/reset`。前後をログに残す)。
  オンライン/オフラインの切替(仕掛台帳の読み方。端末の設定)は管理者だけ
- この端末: ラインはアクセス権限マスタで決まれば自動で当たる(D30)。決まったあとに手で変えるのは管理者パスワードで
- 権限: アクセス権限マスタ(梱包資材マスタの「アクセス権限」。all-tools と共通。D12・D30)を読んだ結果と表ぜんぶの点検
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
        "paths": paths.overview() + [_log_path_row()],
        "master_edit": {"ok": auth.master_edit_ok(), "reason": auth.master_edit_reason(),
                        "setup": auth.setup_mode(), "setup_reason": d.setup_reason if d else "",
                        "possible": auth.can_edit() or auth.setup_mode()},
        "rights": {
            "user": s.user_id,
            "pc": auth.pc_name(),
            "label": "管理者(管理者パスワード・設定の途中)" if auth.can_edit() and not s.can_edit else s.label,
            "admin": auth.can_edit(),
            "source": d.source if d else "",
            "reason": d.reason if d else "",
            "codes": d.codes if d else [],
            "matched": d.matched if d else [],
            "line": {"facility": d.facility if d else "", "line_no": d.line_no if d else "",
                     "text": d.line_text() if d else "", "lines": d.lines if d else [],
                     "problem": d.line_problem if d else "", "notes": d.line_notes if d else []},
            "ignored": d.ignored if d else [],
            "applied": auth.applied_note(),
            "setup": d.setup if d else False,
            "table": access_rights.overview(),
            "known": [{"code": k, "label": v} for k, v in access_rights.CODES.items()],
            "aliases": [{"code": access_rights.ADMIN_ALIAS, "same_as": access_rights.ADMINISTRATOR}],
            "lines": [{"master": k + ("A1〜中板A11" if k == access_rights.MARU else ""), "facility": v}
                      for k, v in access_rights.LINE_FACILITY.items()],
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
            "need": "write", "kind": "dir", "single": True, "path": requested, "default": "",
            "changed": bool(requested)}


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
    """未設定なら誰でも(VBA の SF3)。決まったあとに変えるのは管理者パスワード(all-tools「パスワードがあれば現行の通り変更可」)。"""
    body = request.get_json(silent=True) or {}
    if terminal.line() is not None:
        refused = _refuse_path_edit()
        if refused:
            return refused
    before = terminal.line_label()
    try:
        terminal.set_facility(str(body.get("facility", "")), str(body.get("line_no", "") or ""))
    except ValueError as e:
        return _err("bad_value", str(e))
    log.info("この端末の設備名を %s(%s)にしました", terminal.facility(), terminal.line_no())
    trace.event("settings.terminal", f"この端末のラインを {before or '(未設定)'} → {terminal.line_label()} にしました",
                data={"before": before, "after": terminal.line_label()})
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


def _refuse_path_edit():
    """使用パスを変えてよいか(D26: 管理者パスワード + アクセス権限の管理者。D30: 設定の途中はパスワードだけ)。だめなら 403。"""
    auth.ensure()
    if auth.master_edit_ok():
        return None
    return _err("locked" if (auth.can_edit() or auth.setup_mode()) else "not_admin", auth.master_edit_reason(), 403)


@bp.post("/api/settings/path")
def save_path():
    """使用パスを変える。body: {key, path, reset}(パス 1 つの項目)、
    仕掛台帳は {key: "lot", slot: "first"/"second", path, reset} または {key: "lot", db, second_name}。"""
    body = request.get_json(silent=True) or {}
    key, slot = str(body.get("key", "")), str(body.get("slot", "") or "")
    if key == LOG_KEY:
        # ログの出力パスは 設定 → ログ と同じ設定なので、同じく管理者だけ(パスワードは要らない)
        if not auth.can_edit():
            return _err("not_admin", MSG_ADMIN, 403)
        refused = _set_log_dir("" if body.get("reset") else str(body.get("path", "")))
        return refused or jsonify(_state())
    refused = _refuse_path_edit()
    if refused:
        return refused
    if key not in paths.BY_KEY:
        return _err("bad_value", "その項目はありません")
    if "db" in body:
        db = str(body.get("db", ""))
        if key != paths.LOT or db not in paths.LOT_DBS:
            return _err("bad_value", "2つ目の名前を決められるのは仕掛台帳の SIKALOT / SIKAHIKI / SIKAODR だけです")
        try:
            if "first_name" in body:
                paths.set_name(db, paths.FIRST, "" if body.get("reset") else str(body.get("first_name", "")))
                done = f"第1の名前 {db} を {paths.first_name(db)}"
            else:
                paths.set_second_name(db, "" if body.get("reset") else str(body.get("second_name", "")))
                done = f"第2の名前 {db} を {paths.second_name(db)}"
        except ValueError as e:
            return _err("bad_value", str(e))
        log.info("仕掛台帳の%sにしました", done)
        trace.event("settings.path", f"仕掛台帳の{done}にしました")
        return jsonify(_state())
    if key == paths.LOT and slot not in paths.SLOT_LABELS:
        return _err("bad_value", "仕掛台帳は第1パスか第2パスかを選んでください")
    s = slot if key == paths.LOT else paths.FIRST
    before = paths.slot_value(key, s)
    if body.get("reset"):
        paths.reset_path(key, s)
    else:
        paths.set_path(key, str(body.get("path", "")), s)
    label = paths.BY_KEY[key].label + (f"({paths.SLOT_LABELS[s]})" if key == paths.LOT else "")
    log.info("使用パス %s を %s にしました", label, paths.slot_value(key, s))
    trace.event("settings.path", f"使用パス {label} を {paths.slot_value(key, s)} にしました",
                data={"before": before, "after": paths.slot_value(key, s)})
    if key == "access_master":                      # 権限の表の場所が変わったので読み直す(鍵は開けたまま)
        auth.end()
        auth.ensure()
    return jsonify(_state())


@bp.post("/api/settings/path/reset")
def reset_path():
    """使用パスを既定(paths.py の既定 = VBA のオンラインの値)に戻す。body: {key} または {all: true}。

    仕掛台帳は 第1・第2・2つ目の名前をまとめて戻す。ログの出力パスは「すべて」には入れない(設定 → ログ)。
    前後の値をログ(trace)に残す。
    """
    refused = _refuse_path_edit()
    if refused:
        return refused
    body = request.get_json(silent=True) or {}
    if body.get("all"):
        keys = [i.key for i in paths.ITEMS]
        label = "すべての使用パス"
    else:
        key = str(body.get("key", ""))
        if key not in paths.BY_KEY:
            return _err("bad_value", "その項目はありません")
        keys = [key]
        label = "使用パス " + paths.BY_KEY[key].label
    with trace.timed("settings.path_reset", f"{label}を既定に戻す", keys=keys) as info:
        before = {k: paths.snapshot(k) for k in keys}
        if body.get("all"):
            paths.reset_all()
        else:
            paths.reset_item(keys[0])
        after = {k: paths.snapshot(k) for k in keys}
        diff = {k: {"before": before[k], "after": after[k]} for k in keys if before[k] != after[k]}
        info.update(changed=list(diff))
    trace.event("settings.path", f"{label}を既定に戻しました(変わった項目 {len(diff)})", data=diff)
    log.info("%sを既定に戻しました: %s", label, ", ".join(diff) or "変わった項目なし")
    if "access_master" in diff:                     # 権限の表の場所が変わったので読み直す
        auth.end()
        auth.ensure()
    out = _state()
    out["reset"] = {"keys": list(diff), "message": f"{label}を既定に戻しました"
                    + (f"({len(diff)} 項目)" if body.get("all") else "")}
    return jsonify(out)


@bp.post("/api/settings/path/check")
def check_path():
    """確認ボタン(誰でも)。仕掛台帳は第1パスと第2パスの両方を確かめる。"""
    body = request.get_json(silent=True) or {}
    key = str(body.get("key", ""))
    # 画面で入力中(保存前)の値。送られて来なければ保存した値を確かめる
    typed = {f: str(body[f]) for f in ("path", "first", "second") if body.get(f) is not None}
    names = body.get("second_names") if isinstance(body.get("second_names"), dict) else None
    fnames = body.get("first_names") if isinstance(body.get("first_names"), dict) else None
    if key == LOG_KEY:
        log_text = typed.get("path", str(user_settings.get(config.KEY_LOG_DIR, "") or ""))
        target = logging_utils.resolve_log_dir(log_text.strip().strip('"'))
        why = logging_utils.probe_writable(target)
        one = {"key": key, "slot": "", "label": "", "path": str(target), "exists": target.is_dir(),
               "readable": not why, "writable": not why, "ok": not why,
               "message": f"書けます({target})" if not why else f"書けません: {why}"}
        results = [one]
    elif key in paths.BY_KEY:
        if key == paths.LOT:
            results = [paths.check(key, s, text=typed.get(s), names=names, first_names=fnames) for s in paths.SLOT_LABELS]
        else:
            results = [paths.check(key, text=typed.get("path"))]
        if key == "access_master" and results[0]["path"] and typed.get("path", "").strip() in ("", paths.slot_value(key)):
            # フォルダでもファイルでもよい。実際に読むファイルと「アクセス権限」表を確かめる
            ov = access_rights.overview()
            results[0].update(ok=ov["found"], readable=ov["found"], path=ov["path"],
                              message=(f"読めます({ov['path']} の「{ov['table']}」{ov['rows']} 行)" if ov["found"]
                                       else f"読めません: {ov['problem']}"))
    else:
        return _err("bad_value", "その項目はありません")
    m = paths.mode()
    return jsonify({"key": key, "results": results, "mode": m, "in_use": paths.in_use_text(key, m),
                    "ok": any(r["ok"] for r in results),
                    "message": " / ".join((r["label"] + ": " if r["label"] else "") + r["message"] for r in results)})


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
