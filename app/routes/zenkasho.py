"""全ヵ所検査(一般 / 5N / 三菱受託)の画面(VBA の UFMitsubishi + 全ヵ所検査 シート)

画面の操作 1 つ = `zenkasho_session.ZenkashoSession` のメソッド 1 つ。返り値は `Result`(view・messages・question・stop)。
判断はすべてセッション(業務ロジック)が行い、ここは受け渡すだけ。
"""
from __future__ import annotations

from flask import Blueprint, jsonify, render_template

from ufkensa.kensa import zenkasho_session
from ufkensa.logging_utils import get_logger

from ..shell import shell_context
from . import screen_common

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

@bp.get("/zenkasho")
def page():
    return render_template("zenkasho.html", **shell_context("zenkasho"))


@bp.get("/api/zenkasho/state")
def state():
    return jsonify(zenkasho_session.get().state().to_dict())


@bp.get("/api/zenkasho/report/<path:name>")
def report(name: str):
    """印刷した帳票の HTML(手元のデータフォルダの print)。ブラウザで開いて印刷する(D8)。
    整理 6: 検査・キャスクと同じ確かめ(名前は print 直下の *.html)・max_age=0 にそろえた。"""
    return screen_common.report_file(zenkasho_session.get().store.print_dir, name)


@bp.get("/api/zenkasho/pack_image")
def pack_image():
    """いま出している包装仕様書の画像(UF_PackNo.PacIM)。"""
    return screen_common.pack_image(zenkasho_session.get().pack_image())


@bp.post("/api/zenkasho/<op>")
def operate(op: str):
    return screen_common.operate(op, OPS, zenkasho_session.get, area="zenkasho", label="全ヵ所画面",
                                 trace_fields=lambda s: {"page": s.page},
                                 stop_view=lambda s: s.state().to_dict()["view"])
