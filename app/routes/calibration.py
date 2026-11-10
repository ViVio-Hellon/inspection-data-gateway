"""校正画面(VBA の UF校正)

直尺(伝送コンベックス)の校正: 1000mm 直尺 → ライン別ストレートエッヂ → 手元DB の「校正記録」へ記録。
画面の操作 1 つ = `calibration_session.CalibrationSession` のメソッド 1 つ。返り値は検査画面と同じ形
(view・messages・question・stop)。判断はすべて `ufkensa/kensa/calibration.py`。
"""
from __future__ import annotations

from flask import Blueprint, jsonify, render_template

from ufkensa.kensa import calibration_session
from ufkensa.logging_utils import get_logger

from ..shell import shell_context
from . import screen_common

log = get_logger("app.routes.calibration")

bp = Blueprint("calibration", __name__)

#: 画面の操作 → (セッションのメソッド, 受け取る引数, 説明)
OPS: dict[str, tuple[str, tuple[str, ...], str]] = {
    "open": ("open", (), "校正画面を開く(UF校正.Initialize)"),
    "receive": ("receive", ("text",), "測定値の受信(INP1000 / INPG)"),
    "clear": ("clear", (), "ｸﾘｱ＆ﾘｾｯﾄ(CBcl)"),
    "skip-1000": ("skip_1000", (), "1000mm を送る(CommandButton5)"),
    "rank": ("rank", ("kind",), "品証順位 / 校正順位"),
}


@bp.get("/calibration")
def page():
    return render_template("calibration.html", **shell_context("calibration"))


@bp.get("/api/calibration/state")
def state():
    return jsonify({"view": _view(calibration_session.get()), "messages": [], "question": None, "stop": None})


def _view(s):
    with s._lock:
        return s.view()


@bp.post("/api/calibration/<op>")
def operate(op: str):
    return screen_common.operate(op, OPS, calibration_session.get, area="calibration", label="校正画面",
                                 required=lambda n: True, as_text=True, stop_view=_view)
