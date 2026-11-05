"""保守画面(VBA の Test フォームのうち業務で要る 2 つと、初期設定.bas の端末設定)

2026-11-02 の決め: **誰でも開けて、読むだけの確かめ(歪公差・ソリ・設備仕様)は誰でも使える。**
編集(端末設定のセット・仮ロットの作成/変更/削除)は 管理者パスワード + アクセス権限の「管理者」
(`auth.master_edit_ok()`)。足りなければ 403 {code: "locked"} を返し、画面は鍵の窓(admin_unlock.js)を出して呼び直す。

- 歪公差チェック: 品質基準の表を変えたあとの確認(Test.CommandButton34)。ロット 1 つ / 仕掛の全ロット。
  別の 品質基準.sqlite3 を指定すると、いまの表と並べて違いに印を付ける。ソリ公差(CommandButton24)・
  設備仕様(CommandButton38)の表引きも。**DB は読むだけ**
- 端末設定のセット: 初期設定.AIM / HVC / DelSet(レジストリ → user_settings)
- 仮ロット: 利用者が作るロット(手元DB のシート「仮ロット」。ufkensa/kensa/kari_lot.py)。歪公差チェックで引ける
組み立ては `ufkensa/kensa/maint.py`。判断はすべてサーバ。
"""
from __future__ import annotations

import re

from flask import Blueprint, jsonify, render_template, request

from ufkensa import auth, trace
from ufkensa.kensa import kari_lot, maint, master_admin, store as st, terminal
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


def _need_edit():
    """編集してよくなければ 403 {code: "locked"}(理由はそのまま画面に出す)。"""
    if not auth.master_edit_ok():
        return _err("locked", auth.master_edit_reason() or auth.MSG_NO_LEVEL, 403)
    return None


def _store() -> st.LocalStore:
    return st.LocalStore()


def _lock_state() -> dict:
    return {"unlocked": auth.master_unlocked(), "can_edit": auth.master_edit_ok(), "admin": auth.can_edit(),
            "reason": auth.master_edit_reason()}


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
    return jsonify({"admin": auth.can_edit(), "lock": _lock_state(), "terminal": maint.terminal_state(),
                    "quality_db": str(maint.current_quality_db()), "facilities": list(terminal.FACILITIES),
                    "labels": maint.DIST_LABELS, "keys": list(maint.DIST_KEYS)})


@bp.post("/api/maint/dist")
def dist():
    """歪公差チェック(ロット 1 つ)。Test.CommandButton34_Click。"""
    body = _body()
    lot = str(body.get("lot", "") or "").strip()
    if not _LOT_RE.match(lot):
        return _err("bad_value", "ロットNo は英数字で入れてください", field="lot")
    try:
        alt = _alt(body)
    except ValueError as e:
        return _err("bad_alt", str(e), field="alt")
    with _store() as s:
        kari = kari_lot.get(s, lot)
    with trace.timed("maint.dist", "保守: 歪公差チェック(1ロット)", expected=(Stop,),
                     lot=lot.upper(), kari=bool(kari) or None, alt=str(alt) if alt else None) as info:
        if kari is not None:
            row = maint.dist_check_kari(kari, alt)
            if kari_lot.real_lot_exists(kari.lot):
                row["messages"].append(f"仕掛にも同じロットNo {kari.lot} があります。仮ロットで引きました")
        else:
            row = maint.dist_check(lot, alt)
        info["status"] = row["current"]["status"]
        info["diff"] = row["diff"] or None
    return jsonify({"row": row, "quality_db": str(maint.current_quality_db()), "alt": str(alt) if alt else ""})


@bp.post("/api/maint/dist_all")
def dist_all():
    """歪公差チェック(仕掛の全ロット)。表を変えたあとの確認。"""
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
        kari = None
        if body.get("include_kari", True):
            with _store() as s:
                kari = kari_lot.list_all(s)
        out = maint.dist_check_all(alt, plate_only=bool(body.get("plate_only", True)), prefix=prefix, kari_lots=kari)
        info.update(count=out["count"], diffs=out["diffs"], by_status=out["by_status"], kari=out["kari_used"] or None)
    out["quality_db"] = str(maint.current_quality_db())
    return jsonify(out)


@bp.post("/api/maint/sori")
def sori():
    """ソリ公差(ソリ4級)。Test.CommandButton24_Click。"""
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
    if (d := _need_edit()) is not None:
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


# ------------------------------------------------------------------
# 仮ロット(一覧・写す は誰でも。作る・直す・消す は編集の鍵が要る)
# ------------------------------------------------------------------
@bp.get("/api/maint/kari")
def kari_list():
    with _store() as s:
        rows = [k.to_dict() for k in kari_lot.list_all(s)]
    return jsonify({"rows": rows, "fields": kari_lot.field_specs(), "ken_len": kari_lot.KEN_LEN})


@bp.get("/api/maint/kari/from_real")
def kari_from_real():
    """本物のロットの値を写す(保存はしない。画面の欄に入れるだけ)。"""
    lot = kari_lot.normalize_lot(request.args.get("lot", ""))
    if not _LOT_RE.match(lot):
        return _err("bad_value", "写す元のロットNo は英数字で入れてください", field="source")
    try:
        with trace.timed("maint.kari_from_real", "保守: 仮ロットに本物のロットを写す", expected=(Stop, kari_lot.KariError),
                         lot=lot):
            out = kari_lot.from_real(lot)
    except kari_lot.KariError as e:
        return _err("not_found", str(e), 404, e.field)
    return jsonify(out)


@bp.post("/api/maint/kari/dist")
def kari_dist():
    """画面の欄の値(保存の前でも)で歪公差を引く。読むだけなので鍵は要らない。"""
    b = _body()
    try:
        k = kari_lot.validate(b.get("lot"), b.get("ken_head"), b.get("values"), b.get("hiki_extra"), b.get("source", ""))
        alt = _alt(b)
    except kari_lot.KariError as e:
        return _err("bad_value", str(e), field=e.field)
    except ValueError as e:
        return _err("bad_alt", str(e), field="alt")
    with trace.timed("maint.kari_dist", "保守: 歪公差チェック(仮ロットの欄の値)", expected=(Stop,),
                     lot=k.lot, alt=str(alt) if alt else None) as info:
        row = maint.dist_check_kari(k, alt)
        info["status"] = row["current"]["status"]
    return jsonify({"row": row, "quality_db": str(maint.current_quality_db()), "alt": str(alt) if alt else ""})


@bp.post("/api/maint/kari")
def kari_save():
    """仮ロットを作る(original 空)・直す(original = 直す前のロットNo)。"""
    if (d := _need_edit()) is not None:
        return d
    b = _body()
    try:
        k = kari_lot.validate(b.get("lot"), b.get("ken_head"), b.get("values"), b.get("hiki_extra"), b.get("source", ""))
    except kari_lot.KariError as e:
        return _err("bad_value", str(e), field=e.field)
    original = str(b.get("original", "") or "")
    try:
        with _store() as s, trace.timed("maint.kari_save", "保守: 仮ロットを保存", expected=(kari_lot.KariError,),
                                        lot=k.lot, original=original or None) as info:
            action, k = kari_lot.save(s, k, original)
            trace.note_used(s.path, "save")
            info["action"] = action
            rows = [x.to_dict() for x in kari_lot.list_all(s)]
    except kari_lot.KariError as e:
        return _err("bad_value", str(e), field=e.field)
    log.info("保守: 仮ロットを%s %s(検番 %s)", action, k.lot, k.ken_no)
    return jsonify({"action": action, "row": k.to_dict(), "rows": rows})


@bp.post("/api/maint/kari/delete")
def kari_delete():
    if (d := _need_edit()) is not None:
        return d
    lot = kari_lot.normalize_lot(_body().get("lot", ""))
    with _store() as s, trace.timed("maint.kari_delete", "保守: 仮ロットを消す", lot=lot) as info:
        ok = kari_lot.delete(s, lot)
        trace.note_used(s.path, "save")
        info["deleted"] = ok
        rows = [x.to_dict() for x in kari_lot.list_all(s)]
    if not ok:
        return _err("not_found", f"仮ロット {lot} が見つかりません", 404)
    log.info("保守: 仮ロットを消しました %s", lot)
    return jsonify({"rows": rows})


# ------------------------------------------------------------------
# マスタ管理(見るのは誰でも。直す・足す・消す は編集の鍵が要る)。ufkensa/kensa/master_admin.py
# ------------------------------------------------------------------
def _master_err(e: master_admin.MasterError):
    return _err(e.code, e.message, e.status, e.field)


@bp.get("/api/maint/master/sources")
def master_sources():
    return jsonify({"sources": master_admin.list_sources(), "limit": master_admin.ROW_LIMIT})


@bp.get("/api/maint/master/tables")
def master_tables():
    try:
        return jsonify(master_admin.tables(request.args.get("source", "")))
    except master_admin.MasterError as e:
        return _master_err(e)


@bp.get("/api/maint/master/rows")
def master_rows():
    a = request.args
    try:
        with trace.timed("maint.master_rows", "保守: マスタの行を読む", expected=(master_admin.MasterError,),
                         source=a.get("source", ""), table=a.get("table", ""), q=a.get("q", "") or None):
            return jsonify(master_admin.page(a.get("source", ""), a.get("table", ""), query=a.get("q", ""),
                                             sort=a.get("sort", ""), sort_dir=a.get("dir", "asc")))
    except master_admin.MasterError as e:
        return _master_err(e)


@bp.post("/api/maint/master/row")
def master_row():
    """1 行を 直す(action=save)・足す(add)・消す(delete)。"""
    if (d := _need_edit()) is not None:
        return d
    b = _body()
    src, table, action = str(b.get("source", "")), str(b.get("table", "")), str(b.get("action", ""))
    values = b.get("values") if isinstance(b.get("values"), dict) else {}
    seen = b.get("seen") if isinstance(b.get("seen"), dict) else None
    try:
        if action == "save":
            out = master_admin.save_row(src, table, b.get("rowid"), values, seen)
        elif action == "add":
            out = master_admin.add_row(src, table, values)
        elif action == "delete":
            out = master_admin.delete_row(src, table, b.get("rowid"), seen)
        else:
            return _err("bad_value", f"知らない操作です: {action}")
    except master_admin.MasterError as e:
        return _master_err(e)
    log.info("保守: マスタ %s/%s を%s rowid=%s", src, table, action, out.get("rowid"))
    return jsonify(out)


@bp.get("/api/maint/master/history")
def master_history():
    a = request.args
    return jsonify({"rows": master_admin.history(source_key=a.get("source", ""), table=a.get("table", ""))})


# ------------------------------------------------------------------
# 測定機器ID の伝送テスト(読むだけ。誰でも)
# ------------------------------------------------------------------
@bp.get("/api/maint/device/defaults")
def device_defaults():
    """いま検査中のロットの値(読むだけ)。無ければこの端末のライン。"""
    out = {"line": "", "atu": "", "hab": "", "youto": "", "tak": "", "lot": "", "lines": maint.device_lines()}
    try:
        from ufkensa.kensa import session
        s = session.get().state
        out.update(line=s.line_name or "", atu=s.atu or "", hab=s.hab or "", youto=s.youto or "", tak=s.tak or "",
                   lot=s.lot or "")
    except Exception:                                  # noqa: BLE001 - 検査画面がまだ開いていない
        pass
    if not out["line"]:
        from ufkensa.kensa import lines
        out["line"] = lines.vba_name(terminal.facility() or "")
    return jsonify(out)


@bp.post("/api/maint/device/test")
def device_test():
    b = _body()
    try:
        with trace.timed("maint.device_test", "保守: 測定機器の伝送テスト", expected=(ValueError,),
                         text=str(b.get("text", ""))[:60], line=str(b.get("line", ""))) as info:
            out = maint.device_test(str(b.get("text", "")), str(b.get("line", "")), str(b.get("atu", "")),
                                    str(b.get("hab", "")), str(b.get("youto", "")), str(b.get("st", "S1")),
                                    str(b.get("tak", "")))
            info.update(kind=out["kind"], code=out["code"], ok=out["judge"]["ok"])
    except ValueError as e:
        return _err("bad_value", str(e), field="st")
    return jsonify(out)
