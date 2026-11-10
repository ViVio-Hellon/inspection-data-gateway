"""キャスク検査の画面(VBA の UFCask + キャスク検査シート)

画面の操作 1 つ = `cask_session.CaskSession` のメソッド 1 つ。返り値は `Result`(view・messages・question・stop)。
判断はすべてセッション(部品 modes/cask.py・modes/sheet.py)が行い、ここは受け渡すだけ。
"""
from __future__ import annotations

from flask import Blueprint, jsonify, render_template

from ufkensa.kensa import cask_session
from ufkensa.logging_utils import get_logger

from ..shell import shell_context
from . import screen_common

log = get_logger("app.routes.cask")

bp = Blueprint("cask", __name__)

#: 画面の操作 → (セッションのメソッド, 受け取る引数)
OPS: dict[str, tuple[str, tuple[str, ...]]] = {
    "open": ("open", ()),
    "lot": ("set_lot", ("lot",)),
    "field": ("set_field", ("name", "value")),
    "zero": ("zero", ("text",)),
    "receive": ("receive", ("text",)),
    "select": ("select", ("row", "col")),
    "transfer": ("transfer", ()),
    "print": ("print_out", ("answers",)),
    "save": ("save", ()),
    "new": ("new_sheet", ("answers",)),
    "cancel": ("cancel", ()),
    "inspection": ("toggle_inspection", ()),
    "resume": ("resume", ()),
    "proofread": ("proofread", ()),
    "close": ("close", ()),
    "pack-select": ("pack_select", ("i",)),
    "pack-close": ("pack_close", ("confirmed",)),
}

@bp.get("/cask")
def page():
    return render_template("cask.html", **shell_context("cask"))


@bp.get("/api/cask/state")
def state():
    return jsonify(cask_session.get().state().to_dict())


@bp.get("/api/cask/report/<path:name>")
def report(name: str):
    """印刷した帳票(D8: ブラウザで開いて印刷する HTML)。手元のデータフォルダの print にあるものだけ。"""
    return screen_common.report_file(cask_session.get().store.print_dir, name)


@bp.get("/api/cask/pack_image")
def pack_image():
    """いま出している包装仕様書の画像(UF_PackNo.PacIM)。"""
    return screen_common.pack_image(cask_session.get().pack_image())


def _lot(s) -> dict:
    return {"lot": (s.book.cask.lot_box if s.book is not None else "") or None}


@bp.post("/api/cask/<op>")
def operate(op: str):
    return screen_common.operate(op, OPS, cask_session.get, area="cask", label="キャスク検査",
                                 required=lambda n: n != "answers", trace_fields=_lot,
                                 stop_view=lambda s: s.state().to_dict()["view"])
