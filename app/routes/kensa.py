"""検査画面(A案「計測コックピット」。D6)

画面の操作 1 つ = `session.Session` のメソッド 1 つ。返り値は `Result`(view・messages・question・stop)。
判断はすべてセッション(業務ロジック)が行い、ここは受け渡すだけ。
"""
from __future__ import annotations

import re
from pathlib import Path

from flask import Blueprint, jsonify, render_template, request, send_file

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
    "lot-reload": ("reload_lot", ()),                                  # ロット情報の読み直し(2026-11-07)
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
#: 画像の種類 → 使用パスの項目(2026-11-09: 設定の「VC種類画像」を見る。空なら 参照用ファイル\\VC種類画像)
PICTURES: dict[str, str] = {"vc": "vc_images"}

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
    folder = paths.to_local(paths.get(PICTURES[kind]))
    f = folder / safe
    if not f.is_file():
        # 2026-11-09: 名前の大文字・小文字(VE-20NF.JPG / VE-20NF.jpg)が違っても同じ画像として探す
        try:
            f = next((c for c in folder.iterdir() if c.name.casefold() == safe.casefold() and c.is_file()), f)
        except OSError:
            pass
    if not f.is_file():
        return jsonify({"error": {"code": "not_found", "message": f"画像がありません: {safe}(探した場所: {folder})"}}), 404
    return send_file(f, mimetype="image/jpeg", max_age=60)


@bp.get("/api/kensa/judge_images")
def judge_images_list():
    """判定の画像(OK・NG・対角線)で出せるもの。使用パス「判定の画像」が空・無いときは全部 [](画像を出さない)。"""
    from ufkensa.kensa import judge_images
    return jsonify({"images": judge_images.available(), "show_ms": judge_images.SHOW_MS})


@bp.get("/api/kensa/judge_image/<kind>/<path:name>")
def judge_image(kind: str, name: str):
    """判定の画像 1 枚(VBA の UFOK / UFNG / UF対角 の絵)。決まった名前のものだけ配る。"""
    import mimetypes
    from flask import send_file
    from ufkensa.kensa import judge_images
    f = judge_images.find(kind, name)
    if f is None:
        return jsonify({"error": {"code": "not_found", "message": f"画像がありません: {name}"}}), 404
    return send_file(f, mimetype=mimetypes.guess_type(f.name)[0] or "image/jpeg", max_age=60)


@bp.get("/api/kensa/pack_image")
def pack_image():
    """いま出している包装仕様書の画像(UF_PackNo.PacIM。使用パス QualityPath\\包装仕様書)。"""
    from flask import send_file
    f = session.get().pack_image()
    if f is None or not f.is_file():
        return jsonify({"error": {"code": "not_found", "message": "包装仕様書の画像がありません"}}), 404
    return send_file(f, mimetype="image/jpeg", max_age=0)


@bp.get("/api/kensa/report/<path:name>")
def report(name: str):
    """印刷した帳票(D8: ブラウザで開いて印刷する HTML)。手元のデータフォルダの print にあるものだけ(キャスクと同じ)。"""
    safe = Path(name).name
    if not re.fullmatch(r"[^\\/]+\.html", safe):
        return jsonify({"error": {"code": "bad_value", "message": "帳票の名前が正しくありません"}}), 400
    f = session.get().store.path.parent / "print" / safe
    if not f.is_file():
        return jsonify({"error": {"code": "not_found", "message": f"帳票がありません: {safe}"}}), 404
    return send_file(f, mimetype="text/html", max_age=0)


@bp.get("/api/kensa/lot-watch")
def lot_watch():
    """ロット情報の見張り(画面が定期に呼ぶ)。登録・読み直しの後に変わった種類を返す。"""
    return jsonify(session.get().lot_watch())


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
#: 2026-11-09: 最後に届いた値(画面の受信状態に出す)
_recv_last: dict = {}


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
    why = ""
    if target == "kensa":
        sess = session.get()
        res = sess.receive(text)
        out = res.to_dict()
        # 2026-11-09: 入らなかった理由(検査員未登録・校正未済・ロット未登録・入力先が無い・パケット不一致 など)
        why = res.stop.text if res.stop else getattr(sess, "receive_why", "")
        out["why"] = why
    else:
        out = {}
    out["text"], out["screen"] = text, target
    import time as _t
    _recv_last.update({"text": text, "at": _t.strftime("%H:%M:%S"), "screen": target, "why": why})
    with _recv_lock:
        _recv_seq += 1
        out["seq"] = _recv_seq
        _recv_results.append(out)


def _kensa_receivers() -> list:
    """検査画面で使う受信機。2026-11-09: 現場で「検査・校正で測れない(保守の受信を始めるでは届く)」と言われた。
    「検査画面で受け取る」に印が無いと 1 台も開かなかったため。いまは:
    - 保守で「検査画面で受け取る」に印を付けた受信機があれば、それだけを使う
    - 1 台も印が無ければ、保守 → 測定機器ID で足した受信機(カード)を全部使う"""
    from .maint import _cards_saved, com_kensa_list
    want = com_kensa_list()
    if want:
        return want
    return [dict(c, port=p) for p, c in _cards_saved().items() if p]


@bp.post("/api/kensa/recv/start")
def recv_start():
    """検査画面で使う受信機(_kensa_receivers。何台でも)から受け取りを始める。動いているものはそのまま。"""
    from ufkensa.kensa import comport
    want = _kensa_receivers()
    running = {r["port"]: r for r in comport.receivers() if r["slot"].startswith("kensa:") and r["running"]}
    errors = []
    for c in want:
        if c["port"] in running:
            continue
        # 2026-11-09: 保守で同じポートを受信中(mnt:)だとポートが塞がって開けない。検査画面を開いたら保守の受信を止める
        comport.stop("mnt:" + c["port"])
        try:
            comport.start(c["port"], int(c.get("baud") or 57600), str(c.get("flow") or "rtscts"),
                          int(c.get("bytesize") or 8), str(c.get("parity") or "N"), int(c.get("stopbits") or 1),
                          sink=_to_kensa, slot="kensa:" + c["port"])
        except (comport.ComError, ValueError) as e:
            errors.append(f"{c['port']} を開けません: {e}")
    return jsonify(_recv_state(errors))


def _recv_state(errors=None) -> dict:
    from ufkensa.kensa import comport
    rs = [r for r in comport.receivers() if r["slot"].startswith("kensa:")]
    on = [r["port"] for r in rs if r["running"]]
    errs = list(errors or []) + [f"{r['port']}: {r['error']}" for r in rs if r["error"]]
    # 2026-11-09: 受け取れないときは「なぜ」を人の言葉で返す
    why = ""
    if not on:
        if not _kensa_receivers():
            why = "受信機が登録されていません。保守 → 測定機器ID で受信機を足してください"
        elif errs:
            why = "受信機を開けません(ほかで使用中・抜けている など): " + " / ".join(errs)
        else:
            why = "受信機が止まっています(保守で受信中なら、保守の受信を止めてからこの画面を開き直してください)"
    return {"on": bool(on), "ports": on, "port": "・".join(on), "error": " / ".join(errs), "why": why,
            "receivers": [{"port": r["port"], "running": r["running"], "error": r["error"]} for r in rs],
            "last": dict(_recv_last)}


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
    return jsonify({"results": items, "last": last, "running": st["on"], "port": st["port"], "error": st["error"],
                    "why": st["why"], "received": st["last"]})
