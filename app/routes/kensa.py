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
    "lot-source": ("set_lot_source", ("on", "lot", "answers")),     # D28 控え(第2)から読む
    # 外観・不良要因・検入高さ(MPEX の各ページ)
    "visual": ("visual", ("lc", "items")),
    "visual-clear": ("visual_clear", ("lc",)),
    "defect-reflect": ("defect_reflect", ("items",)),
    "defect-clear": ("defect_clear", ()),
    "yamalb": ("yama_lb", ("value",)),
    # オーダーNo・包装仕様書(UF_PackNo)
    "orders": ("ord_no_open", ()),
    "order": ("select_order", ("order",)),
    "pack-select": ("pack_select", ("i",)),
    "pack-close": ("pack_close", ("confirmed",)),
    # 保留・まとめて削除・廃棄フラグ・印刷履歴クリア・サイズ修正・名簿画面の案内
    "hold": ("hold", ("answers",)),
    "ng-delete": ("ng_delete", ("pos", "answers")),
    "izyou": ("izyou", ("answers",)),
    "print-reset": ("print_reset", ("answers",)),
    "resize": ("resize", ("answers",)),
    "namelist-ack": ("name_list_ack", ()),
}

#: 画像の種類 → (使用パスの項目, その下のフォルダ)。VC 画像は SaveLoad.VC_Picture(PackingPath\VC種類画像)。
#: 包装仕様書は /api/kensa/pack_image(いま出している UF_PackNo の画像だけ。全ヵ所・キャスクと同じ packing.py)
PICTURES: dict[str, tuple[str, str]] = {"vc": ("packing", "VC種類画像")}

MSG_NOT_READY = "この操作はまだ結合中です(工程 8c)。いまは VBA の UF検査判定.xlsm を使ってください。"


@bp.get("/kensa")
def page():
    return render_template("kensa.html", **shell_context("kensa"))


@bp.get("/api/kensa/picture/<kind>/<path:name>")
def picture(kind: str, name: str):
    """VC 画像・包装仕様書の画像(使用パスの下のフォルダの <名前>.JPG)。無ければ 404(VBA は出さないだけ)。"""
    from pathlib import Path
    from flask import send_file
    from ufkensa.kensa import paths
    if kind not in PICTURES:
        return jsonify({"error": {"code": "not_found", "message": "その画像の種類はありません"}}), 404
    safe = Path(name).name
    if safe != name or not safe.lower().endswith((".jpg", ".jpeg")):
        return jsonify({"error": {"code": "bad_value", "message": "画像の名前が正しくありません"}}), 400
    key, sub = PICTURES[kind]
    f = paths.to_local(paths.join(paths.get(key), sub, safe))
    if not f.is_file():
        return jsonify({"error": {"code": "not_found", "message": f"画像がありません: {safe}"}}), 404
    return send_file(f, mimetype="image/jpeg", max_age=60)


@bp.get("/api/kensa/pack_image")
def pack_image():
    """いま出している包装仕様書の画像(UF_PackNo.PacIM。使用パス QualityPath\\包装仕様書)。"""
    from flask import send_file
    f = session.get().pack_image()
    if f is None or not f.is_file():
        return jsonify({"error": {"code": "not_found", "message": "包装仕様書の画像がありません"}}), 404
    return send_file(f, mimetype="image/jpeg", max_age=0)


@bp.get("/api/kensa/state")
def state():
    s = session.get()
    return jsonify(s.current().to_dict())


@bp.post("/api/kensa/<op>")
def operate(op: str):
    if op not in OPS:
        return jsonify({"error": {"code": "not_found", "message": "その操作はありません"}}), 404
    method, names = OPS[op]
    body = request.get_json(silent=True) or {}
    kwargs = {n: body[n] for n in names if n in body}
    s = session.get()
    from ufkensa import trace
    try:
        with trace.timed(f"kensa.{op}", f"検査画面: {op}", expected=(Stop, NotImplementedError),
                         lot=s.state.lot or None):
            result = getattr(s, method)(**kwargs)
    except NotImplementedError:
        return jsonify({"error": {"code": "not_ready", "message": MSG_NOT_READY}}), 501
    except Stop as e:                                   # 念のため: セッションが Result に包まずに投げたとき
        return jsonify({"view": view.build(s.state), "messages": [m.to_dict() for m in e.messages],
                        "question": None, "stop": e.to_dict()["stop"]})
    return jsonify(result.to_dict())


# ------------------------------------------------------------------
# 受信機から直接受け取る(保守 → 測定機器ID の COM受信テストで決めた受信機。2026-11-06)
# 測定器の値を、画面のカーソルに関係なくサーバが受け取り、検査の receive に渡す。
# 画面は /api/kensa/recv を見て、届いた結果(描き直し・知らせ)を出す。
# ------------------------------------------------------------------
import threading as _threading
from collections import deque as _deque

_recv_lock = _threading.Lock()
_recv_results: "_deque[dict]" = _deque(maxlen=50)
_recv_seq = 0


def _to_kensa(text: str) -> None:
    global _recv_seq
    s = session.get()
    res = s.receive(text)
    out = res.to_dict()
    out["text"] = text
    with _recv_lock:
        _recv_seq += 1
        out["seq"] = _recv_seq
        _recv_results.append(out)


@bp.post("/api/kensa/recv/start")
def recv_start():
    """この端末の受信の設定で、検査画面の受け取りを始める(設定で「検査画面で受け取る」が入っているときだけ)。"""
    from ufkensa import user_settings
    from ufkensa.kensa import comport
    conf = dict(user_settings.get("com_receive", {}) or {})
    if not conf.get("kensa") or not conf.get("port"):
        return jsonify({"on": False, **comport.status(10 ** 9)})
    st = comport.status(10 ** 9)
    if st["running"] and st["to_kensa"] and st["port"] == conf["port"]:
        return jsonify({"on": True, **st})
    try:
        comport.start(conf["port"], int(conf.get("baud") or 57600), str(conf.get("flow") or "rtscts"),
                      int(conf.get("bytesize") or 8), str(conf.get("parity") or "N"), int(conf.get("stopbits") or 1),
                      sink=_to_kensa)
    except comport.ComError as e:
        return jsonify({"on": False, **comport.status(10 ** 9), "error": str(e)})
    return jsonify({"on": True, **comport.status(10 ** 9)})


@bp.get("/api/kensa/recv")
def recv_poll():
    from ufkensa.kensa import comport
    since = int(request.args.get("since", "0") or 0)
    with _recv_lock:
        items = [r for r in _recv_results if r["seq"] > since]
        last = _recv_seq
    st = comport.status(10 ** 9)
    return jsonify({"results": items, "last": last, "running": st["running"] and st["to_kensa"],
                    "port": st["port"], "error": st["error"]})
