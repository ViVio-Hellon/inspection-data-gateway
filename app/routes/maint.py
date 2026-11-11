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
import sqlite3
from typing import Optional

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
# 判定の画像: 表示時間・表示サイズ(PC ごと)。見るのは誰でも、変えるのは編集の鍵(端末設定と同じ)
@bp.get("/api/maint/judge_display")
def judge_display_get():
    from ufkensa.kensa import judge_images
    return jsonify({"display": judge_images.display(), **judge_images.display_meta(),
                    "slots": [{**judge_images.slot_info(k), "kind": judge_images.SLOTS[k][0]}
                              for k in judge_images.SLOTS]})


@bp.post("/api/maint/judge_display")
def judge_display_set():
    if (d := _need_edit()) is not None:
        return d
    from ufkensa.kensa import judge_images
    before = judge_images.display()
    from ufkensa import user_settings
    try:
        after = judge_images.save_display(_body().get("display"))
    except (OSError, user_settings.NotSaved) as e:
        return _err("not_saved", str(e), 500)
    log.info("保守: 判定の画像の表示 %s → %s", before, after)
    return jsonify({"display": after})


# ------------------------------------------------------------------
# 公差外の確認を出す行(2026-11-11)。伝送用ファイルの共有マスタ。見るのは誰でも、変えるのは編集の鍵
def _confi_rows_payload(rows: dict) -> dict:
    from ufkensa.kensa import confi_rows as cr
    return {"kinds": [{"kind": k, "label": cr.LABELS[k], "max": cr.MAX_ROW[k], "rows": cr.fmt(rows[k]),
                       "default": cr.fmt(cr.DEFAULTS[k])} for k in cr.KINDS]}


@bp.get("/api/maint/confi_rows")
def confi_rows_get():
    from ufkensa.kensa import confi_rows as cr
    cr.ensure_table()                       # 無ければ VBA の既定で作る(読めなくても画面は既定を出す)
    return jsonify(_confi_rows_payload(cr.rows_map()))


@bp.post("/api/maint/confi_rows")
def confi_rows_set():
    if (d := _need_edit()) is not None:
        return d
    from ufkensa.kensa import confi_rows as cr
    vals = _body().get("rows")
    if not isinstance(vals, dict):
        return _err("bad_request", "rows がありません")
    before = cr.rows_map()
    try:
        after = cr.save(vals)
    except cr.Invalid as e:
        return _err("invalid", str(e), 400, field=e.kind)
    except (OSError, sqlite3.Error) as e:
        return _err("not_saved", f"保存できませんでした: {e}", 500)
    log.info("保守: 公差外の確認行 %s → %s", {k: cr.fmt(v) for k, v in before.items()},
             {k: cr.fmt(v) for k, v in after.items()})
    return jsonify(_confi_rows_payload(after))


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


# ------------------------------------------------------------------
# COM受信テスト(測定器の受信機の COM ポートを直接読む。見せるだけ)
# ------------------------------------------------------------------
KEY_COM = "com_receive"         # この端末の受信の設定(次に開いたときの初期値)


def _com_saved() -> dict:
    from ufkensa import user_settings
    from ufkensa.kensa import comport
    d = dict(comport.DEFAULTS, port="", kensa=False)
    d.update({k: v for k, v in (user_settings.get(KEY_COM, {}) or {}).items() if k in d})
    return d


@bp.get("/api/maint/com")
def com_state():
    from ufkensa.kensa import comport
    since = int(request.args.get("since", "0") or 0)
    return jsonify({**comport.status(since), "ports": comport.list_ports(), "saved": _com_saved(),
                    "kensa": com_kensa_list()})


@bp.post("/api/maint/com/start")
def com_start():
    from ufkensa import user_settings
    from ufkensa.kensa import comport
    b = _body()
    try:
        conf = {"port": str(b.get("port", "")).strip(), "baud": int(b.get("baud") or comport.DEFAULT_BAUD),
                "bytesize": int(b.get("bytesize") or 8), "parity": str(b.get("parity") or "N"),
                "stopbits": int(b.get("stopbits") or 1), "flow": str(b.get("flow") or "rtscts")}
        out = comport.start(**conf)
        conf["kensa"] = bool(_com_saved().get("kensa"))
    except (ValueError, comport.ComError) as e:
        trace.event("maint.com", f"COM受信テストを始められません: {e}")
        return _err("com", str(e))
    user_settings.save_or_raise(KEY_COM, conf)
    trace.event("maint.com", f"COM受信テスト開始 {conf}")
    return jsonify({**out, "ports": comport.list_ports(), "saved": conf, "kensa": com_kensa_list()})


@bp.post("/api/maint/com/stop")
def com_stop():
    from ufkensa.kensa import comport
    return jsonify(comport.stop())


KEY_COM_KENSA = "com_kensa"     # 検査画面で受け取る受信機の一覧 [{port, baud, ...}]


def com_kensa_list() -> list:
    from ufkensa import user_settings
    v = user_settings.get(KEY_COM_KENSA, None)
    if v is None:                                   # 前の形(com_receive の kensa=True)を読み替える
        old = dict(user_settings.get(KEY_COM, {}) or {})
        v = [{k: old[k] for k in ("port", "baud", "bytesize", "parity", "stopbits", "flow") if k in old}] \
            if old.get("kensa") and old.get("port") else []
    return [x for x in v if isinstance(x, dict) and x.get("port")]


@bp.post("/api/maint/com/kensa")
def com_kensa():
    """「検査画面でこの受信機から受け取る」を入れる・外す(この端末の設定。受信機は何台でも)。"""
    from ufkensa import user_settings
    b = _body()
    port = str(b.get("port", "")).strip()
    if not port:
        return _err("com", "先に受信機(ポート)を選んでください")
    items = [x for x in com_kensa_list() if x["port"] != port]
    if b.get("on"):
        conf = {"port": port}
        conf.update({k: b[k] for k in ("baud", "bytesize", "parity", "stopbits", "flow") if k in b})
        items.append(conf)
    user_settings.save_or_raise(KEY_COM_KENSA, items)
    trace.event("maint.com", f"検査画面で受け取る受信機: {[x['port'] for x in items]}")
    return jsonify({"kensa": items})


# ------------------------------------------------------------------
# 測定機器ID 画面(2026-11-06 の選定: 受信機カード + 受信記録 + 段の流れ + 当たった行を光らせる)
# 保守で動かす受信は「mnt:<ポート>」(検査には流さない)。検査画面は「kensa:<ポート>」で別に動かす。
# ------------------------------------------------------------------
KEY_COM_CARDS = "com_cards"     # この端末で並べる受信機 {ポート: 受信の設定}


def _cards_saved() -> dict:
    from ufkensa import user_settings
    v = user_settings.get(KEY_COM_CARDS, {}) or {}
    return {k: dict(c) for k, c in v.items() if isinstance(c, dict)}


def _save_cards(cards: dict) -> None:
    from ufkensa import user_settings
    user_settings.save_or_raise(KEY_COM_CARDS, cards)


def _conf_of(b: dict, base: Optional[dict] = None) -> dict:
    from ufkensa.kensa import comport
    c = dict(comport.DEFAULTS)
    c.update(base or {})
    for k, f in (("baud", int), ("bytesize", int), ("stopbits", int), ("parity", str), ("flow", str)):
        if b.get(k) not in (None, ""):
            c[k] = f(b[k])
    return c


def _rx_cards() -> list:
    from ufkensa.kensa import comport
    saved = _cards_saved()
    kensa = {x["port"]: x for x in com_kensa_list()}
    running = {r["port"]: r for r in comport.receivers()}
    ports = {p["port"]: p for p in comport.list_ports()}
    names = list(saved) + [p for p in kensa if p not in saved]
    lines, _last = comport.log_lines("", 0)
    out = []
    for port in names:
        conf = _conf_of({}, saved.get(port) or kensa.get(port) or {})
        r = running.get(port)
        mine = [ln for ln in lines if ln["port"] == port]
        kind = comport.kind_of(port)
        out.append({
            "port": port, "kind": kind, "label": (ports.get(port) or {}).get("label") or port,
            "present": port in ports or not comport.IS_WINDOWS, "conf": conf,
            "running": bool(r and r["running"]), "where": (r["slot"].split(":")[0] if r else ""),
            "error": (r or {}).get("error", ""), "bytes": (r or {}).get("bytes", 0),
            "count": len(mine), "last": mine[-1] if mine else None, "kensa": port in kensa})
    return out


@bp.get("/api/maint/rx")
def rx_state():
    """受信機カード・足せるポート・受信記録(since より新しい行。各行に伝送テストの結果を付ける)。"""
    from ufkensa.kensa import comport
    since = int(request.args.get("since", "0") or 0)
    lines, last = comport.log_lines("mnt:", since)
    a = request.args
    for ln in lines:
        try:
            ln["test"] = maint.device_test(ln["text"], a.get("line", ""), a.get("atu", ""), a.get("hab", ""),
                                           a.get("youto", ""), a.get("st", "S1") or "S1", a.get("tak", ""))
        except ValueError as e:
            ln["test"] = {"error": str(e)}
        except Exception as e:                       # noqa: BLE001
            # 2026-11-09: 判定(測定機器ID の表を読む)で失敗しても、受信記録は返す(前は一覧ごと返らなかった)
            ln["test"] = {"error": f"判定できませんでした: {e}"}
    return jsonify({"cards": _rx_cards(), "ports": comport.list_ports(), "lines": lines, "last": last})


@bp.post("/api/maint/rx/start")
def rx_start():
    from ufkensa.kensa import comport
    b = _body()
    port = str(b.get("port", "")).strip()
    cards = _cards_saved()
    conf = _conf_of(b, cards.get(port))
    # 2026-11-09: 検査画面が同じポートで受信中(kensa:)だと開けない。保守で始めるときは検査画面の受信を止める
    comport.stop("kensa:" + port)
    try:
        comport.start(port, conf["baud"], conf["flow"], conf["bytesize"], conf["parity"], conf["stopbits"],
                      slot="mnt:" + port)
    except (comport.ComError, ValueError) as e:
        trace.event("maint.com", f"受信を始められません {port}: {e}")
        return _err("com", str(e))
    cards[port] = conf
    try:
        _save_cards(cards)
    except Exception as e:                           # noqa: BLE001
        # 2026-11-09: 受信は始まっている。設定を覚えられなくても受信は止めない(知らせるだけ)
        trace.event("maint.com", f"受信機の設定を覚えられませんでした {port}: {e}")
        return jsonify({"cards": _rx_cards(), "warning": f"受信は始めましたが、設定を覚えられませんでした: {e}"})
    trace.event("maint.com", f"受信を始めました {port} {conf}")
    return jsonify({"cards": _rx_cards()})


@bp.post("/api/maint/rx/stop")
def rx_stop():
    from ufkensa.kensa import comport
    port = str(_body().get("port", "")).strip()
    comport.stop("mnt:" + port)
    return jsonify({"cards": _rx_cards()})


@bp.post("/api/maint/rx/add")
def rx_add():
    """受信機カードを足す(まだ受信は始めない)。"""
    from ufkensa.kensa import comport
    b = _body()
    port = str(b.get("port", "")).strip()
    if not port:
        return _err("com", "足す受信機(ポート)を選んでください")
    cards = _cards_saved()
    cards.setdefault(port, _conf_of(b))
    _save_cards(cards)
    return jsonify({"cards": _rx_cards(), "ports": comport.list_ports()})


@bp.post("/api/maint/rx/conf")
def rx_conf():
    b = _body()
    port = str(b.get("port", "")).strip()
    cards = _cards_saved()
    cards[port] = _conf_of(b, cards.get(port))
    _save_cards(cards)
    # 検査画面で受け取る設定にも同じ値を入れる
    from ufkensa import user_settings
    items = [dict(x, **cards[port]) if x["port"] == port else x for x in com_kensa_list()]
    user_settings.save_or_raise(KEY_COM_KENSA, items)
    return jsonify({"cards": _rx_cards()})


@bp.post("/api/maint/rx/remove")
def rx_remove():
    from ufkensa import user_settings
    from ufkensa.kensa import comport
    port = str(_body().get("port", "")).strip()
    comport.stop("mnt:" + port)
    cards = _cards_saved()
    cards.pop(port, None)
    _save_cards(cards)
    user_settings.save_or_raise(KEY_COM_KENSA, [x for x in com_kensa_list() if x["port"] != port])
    return jsonify({"cards": _rx_cards()})


# ----------------------------------------------------------------------
# 履歴出力(2026-11-07): 検査を済ませなくても PDF・xlsx・CSV の保存を試す
# ----------------------------------------------------------------------
def _kensa_line() -> tuple[str, str]:
    """この端末の (Lname, ライン名)。検査画面と同じ起動をしてから読む(ライン設定が無ければ Stop)。"""
    from ufkensa.kensa import session
    s = session.get()
    r = s.start()
    if r.stop is not None:
        raise Stop(r.stop.text, source="履歴出力")
    return s.state.lname, s.state.line_name


def _stop_json(e: Stop):
    return _err("stop", e.text, 400)


@bp.get("/api/maint/history/state")
def history_state():
    from ufkensa.kensa import history_export as hx
    s = _store()
    try:
        return jsonify({"test_loaded": hx.has_test_history(s), "test_lot": hx.json_lot()})
    finally:
        s.close()


@bp.post("/api/maint/history/test")
def history_test():
    """テスト履歴を入れる(add=true)/ 消す(add=false)。"""
    from ufkensa.kensa import history_export as hx
    body = _body()
    s = _store()
    try:
        if body.get("add", True):
            lname, _line = _kensa_line()
            with trace.timed("maint.history", "テスト履歴を入れる", line=lname):
                out = hx.load_test_history(s, lname)
        else:
            with trace.timed("maint.history", "テスト履歴を消す"):
                out = hx.remove_test_history(s)
        return jsonify(out)
    except Stop as e:
        return _stop_json(e)
    finally:
        s.close()


@bp.post("/api/maint/history/pdf")
def history_pdf():
    from ufkensa.kensa import history_export as hx
    body = _body()
    try:
        row = int(body.get("row", 0))
    except (TypeError, ValueError):
        row = 0
    if row < 2:
        return _err("bad_value", "履歴の行を選んでください", field="row")
    s = _store()
    try:
        with trace.timed("maint.history", "履歴 PDF 出力", expected=(Stop,), row=row):
            return jsonify(hx.export_pdf(s, row, mimi=bool(body.get("mimi"))))
    except Stop as e:
        return _stop_json(e)
    finally:
        s.close()


@bp.post("/api/maint/history/csv")
def history_csv():
    from ufkensa.kensa import history_export as hx
    body = _body()
    s = _store()
    try:
        lname, line = _kensa_line()
        with trace.timed("maint.history", "履歴 CSV 出力", expected=(Stop,), line=line):
            return jsonify(hx.export_csv(s, lname, line, mimi=bool(body.get("mimi"))))
    except Stop as e:
        return _stop_json(e)
    finally:
        s.close()


@bp.post("/api/maint/history/nippou")
def history_nippou():
    """いまの直の日報を all-tools の日報モードが読める CSV で書き出す(2026-11-08)。"""
    from ufkensa.kensa import nippou_export, session
    s = session.get()
    r = s.start()
    if r.stop is not None:
        return _stop_json(r.stop)
    with s._lock:                                                     # noqa: SLF001
        try:
            p = nippou_export.export_now(s.store, s.state, s._env())     # noqa: SLF001
        except (OSError, ValueError) as e:
            return _err("write", f"日報の CSV を書き出せませんでした: {e}", 500)
    if p is None:
        return jsonify({"files": [], "messages": [{"text": "いまの直の日報に行がありません", "level": "warn"}]})
    return jsonify({"files": [str(p)], "messages": []})
