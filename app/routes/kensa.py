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
    "calibrated": ("after_calibration", ()),          # 校正に合格して検査へ戻る(UF校正.INPG_Exit)
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


#: いま開いている画面(受信の結果を取りに来た画面 → 最後に来た時刻)。2026-11-06 の見直し:
#: 前は届いた値をいつも通常検査に入れていたので、校正・Aインプット・キャスク・全ヵ所を開いて測った値が、
#: 見えないまま通常検査の欄に入るおそれがあった。いま開いている(最後に取りに来た)画面にだけ渡す
_recv_seen: dict[str, float] = {}
RECV_ALIVE_SEC = 3.0
_recv_dropped = 0


def _recv_target() -> str:
    import time as _t
    now = _t.monotonic()
    alive = {k: v for k, v in _recv_seen.items() if now - v <= RECV_ALIVE_SEC}
    return max(alive, key=alive.get) if alive else ""


def _to_kensa(text: str) -> None:
    """受信機から届いた 1 件。通常検査が開いていればサーバで受け取る(カーソルに関係なく入る)。
    ほかの画面なら、その画面が取りに来たときに渡す(画面の「受け取る欄」に入れたのと同じ処理になる)。
    どの画面も開いていなければ捨てて記録する(見えないまま検査に入れない)。"""
    global _recv_seq, _recv_dropped
    with _recv_lock:
        target = _recv_target()
    if target == "":
        _recv_dropped += 1
        from ufkensa.logging_utils import get_logger
        get_logger(__name__).warning("受信機の値を受け取る画面が開いていないので捨てました: %r", text)
        return
    if target == "kensa":
        out = session.get().receive(text).to_dict()
    else:
        out = {}
    out["text"], out["screen"] = text, target
    with _recv_lock:
        _recv_seq += 1
        out["seq"] = _recv_seq
        _recv_results.append(out)


@bp.post("/api/kensa/recv/start")
def recv_start():
    """保守で「検査画面で受け取る」にした受信機(何台でも)から受け取りを始める。動いているものはそのまま。"""
    from ufkensa.kensa import comport
    from .maint import com_kensa_list
    want = com_kensa_list()
    running = {r["port"]: r for r in comport.receivers() if r["slot"].startswith("kensa:") and r["running"]}
    errors = []
    for c in want:
        if c["port"] in running:
            continue
        try:
            comport.start(c["port"], int(c.get("baud") or 57600), str(c.get("flow") or "rtscts"),
                          int(c.get("bytesize") or 8), str(c.get("parity") or "N"), int(c.get("stopbits") or 1),
                          sink=_to_kensa, slot="kensa:" + c["port"])
        except (comport.ComError, ValueError) as e:
            errors.append(str(e))
    return jsonify(_recv_state(errors))


def _recv_state(errors=None) -> dict:
    from ufkensa.kensa import comport
    rs = [r for r in comport.receivers() if r["slot"].startswith("kensa:")]
    on = [r["port"] for r in rs if r["running"]]
    errs = list(errors or []) + [r["error"] for r in rs if r["error"]]
    return {"on": bool(on), "ports": on, "port": "・".join(on), "error": " / ".join(errs)}


@bp.get("/api/kensa/recv")
def recv_poll():
    """画面が受信の結果を取りに来る(screen = kensa / calibration / ainput / cask / zenkasho)。来た画面を「開いている」とする。"""
    import time as _t
    since = int(request.args.get("since", "0") or 0)
    screen = str(request.args.get("screen", "kensa") or "kensa")
    with _recv_lock:
        if request.args.get("visible", "1") != "0":
            _recv_seen[screen] = _t.monotonic()
        else:
            _recv_seen.pop(screen, None)
        items = [r for r in _recv_results if r["seq"] > since and r.get("screen", "kensa") == screen]
        last = _recv_seq
    st = _recv_state()
    return jsonify({"results": items, "last": last, "running": st["on"], "port": st["port"], "error": st["error"]})
