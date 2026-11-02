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
    # 外観・不良要因・検入高さ(MPEX の各ページ)
    "visual": ("visual", ("lc", "items")),
    "visual-clear": ("visual_clear", ("lc",)),
    "defect-reflect": ("defect_reflect", ("items",)),
    "defect-clear": ("defect_clear", ()),
    "yamalb": ("yama_lb", ("value",)),
    # オーダーNo・包装仕様書(UF_PackNo)
    "orders": ("ord_no_open", ()),
    "order": ("select_order", ("order",)),
    "pack-select": ("pack_select", ("row",)),
    "pack-close": ("pack_close", ("checked",)),
    # 保留・まとめて削除・廃棄フラグ・印刷履歴クリア・サイズ修正・名簿画面の案内
    "hold": ("hold", ("answers",)),
    "ng-delete": ("ng_delete", ("pos", "answers")),
    "izyou": ("izyou", ("answers",)),
    "print-reset": ("print_reset", ("answers",)),
    "resize": ("resize", ("answers",)),
    "namelist-ack": ("name_list_ack", ()),
}

#: 画像の種類 → (使用パスの項目, その下のフォルダ)。VC 画像は SaveLoad.VC_Picture、包装仕様書は UFPack_Open
PICTURES: dict[str, tuple[str, str]] = {"vc": ("packing", "VC種類画像"), "pack": ("quality", "包装仕様書")}

MSG_NOT_READY = "この操作はまだ結合中です(工程 8c)。いまは VBA の UF検査判定.xlsm を使ってください。"


@bp.get("/kensa")
def page():
    return render_template("kensa.html", **shell_context("kensa"))


@bp.get("/api/kensa/sound/<path:name>")
def sound(name: str):
    """判定音・案内の音声(VBA の C:\\VBAWAV\\*.wav。使用パス「音声」)。無ければ 404(VBA も鳴らないだけ)。"""
    from pathlib import Path
    from flask import send_file
    from ufkensa.kensa import paths
    safe = Path(name).name
    if not safe.lower().endswith(".wav"):
        return jsonify({"error": {"code": "bad_value", "message": "wav だけ鳴らせます"}}), 400
    f = paths.to_local(paths.join(paths.get("wav"), safe))
    if not f.is_file():
        return jsonify({"error": {"code": "not_found", "message": f"音声がありません: {safe}"}}), 404
    return send_file(f, mimetype="audio/wav", max_age=3600)


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
