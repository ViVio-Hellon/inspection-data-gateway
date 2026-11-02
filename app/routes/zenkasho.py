"""全ヵ所検査(一般 / 5N / 三菱受託)・H494 検査の画面(VBA の UFMitsubishi + 全ヵ所検査 / H494検査 シート)

画面の操作 1 つ = `zenkasho_session.ZenkashoSession` のメソッド 1 つ。返り値は `Result`(view・messages・question・stop)。
判断はすべてセッション(業務ロジック)が行い、ここは受け渡すだけ。
"""
from __future__ import annotations

from pathlib import Path

from flask import Blueprint, jsonify, render_template, request, send_file

from ufkensa import trace
from ufkensa.kensa import zenkasho_session
from ufkensa.kensa.flow import Stop
from ufkensa.logging_utils import get_logger

from ..shell import shell_context

log = get_logger("app.routes.zenkasho")

bp = Blueprint("zenkasho", __name__)

#: 画面の操作 → (セッションのメソッド, 受け取る引数)
OPS: dict[str, tuple[str, tuple[str, ...]]] = {
    "start": ("start", ("choice_input",)),
    "choose": ("choose", ("mode", "choice_input")),
    "select": ("select", ("r", "c")),
    "receive": ("receive", ("text",)),
    "erase": ("erase", ()),
    "form": ("set_form", ("name", "value")),
    "height": ("save_height", ()),
    "weight": ("weight", ()),
    "transfer": ("transfer", ()),
    "print": ("print", ("which", "answers")),
    "new": ("new_sheet", ("answers",)),
    "resume": ("resume", ()),
    "next-empty": ("next_empty", ()),
    "pack-select": ("pack_select", ("i",)),
    "pack-close": ("pack_close", ("confirmed",)),
}

#: 包装仕様書の画像の種類(拡張子 → MIME)。VBA は LoadPicture で .JPG を読む
_IMAGE_TYPES = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".gif": "image/gif",
                ".bmp": "image/bmp"}


@bp.get("/zenkasho")
def page():
    return render_template("zenkasho.html", **shell_context("zenkasho"))


@bp.get("/api/zenkasho/state")
def state():
    return jsonify(zenkasho_session.get().state().to_dict())


@bp.get("/api/zenkasho/report/<path:name>")
def report(name: str):
    """印刷した帳票の HTML(手元のデータフォルダの print)。ブラウザで開いて印刷する(D8)。"""
    s = zenkasho_session.get()
    safe = Path(name).name
    if not safe.lower().endswith(".html"):
        return jsonify({"error": {"code": "bad_value", "message": "帳票は .html だけです"}}), 400
    f = s._print_dir() / safe
    if not f.is_file():
        return jsonify({"error": {"code": "not_found", "message": f"帳票がありません: {safe}"}}), 404
    return send_file(f, mimetype="text/html")


@bp.get("/api/zenkasho/pack_image")
def pack_image():
    """いま出している包装仕様書の画像(UF_PackNo.PacIM)。使用パス(QualityPath\\包装仕様書)のファイルだけを返す。"""
    f = zenkasho_session.get().pack_image()
    if f is None or not f.is_file():
        return jsonify({"error": {"code": "not_found", "message": "包装仕様書の画像がありません"}}), 404
    return send_file(f, mimetype=_IMAGE_TYPES.get(f.suffix.lower(), "application/octet-stream"), max_age=0)


@bp.post("/api/zenkasho/<op>")
def operate(op: str):
    if op not in OPS:
        return jsonify({"error": {"code": "not_found", "message": "その操作はありません"}}), 404
    method, names = OPS[op]
    body = request.get_json(silent=True) or {}
    kwargs = {n: body[n] for n in names if n in body}
    s = zenkasho_session.get()
    with trace.timed(f"zenkasho.{op}", f"全ヵ所・H494 画面: {op}", expected=(Stop,), page=s.page):
        result = getattr(s, method)(**kwargs)
    return jsonify(result.to_dict())
