"""品質基準画面(VBA の UF品質基準)

いま検査中のロット(または入れたロットNo)の品質基準を、規格ごとのカードで読み取り専用に見せる。
組み立ては `ufkensa/kensa/quality_session.py`(照合は quality.py / quality_view.py)。マスタの編集はしない
(UF品質基準 にも無い)。
"""
from __future__ import annotations

import re

from flask import Blueprint, jsonify, render_template, request

from ufkensa import trace
from ufkensa.kensa import quality_session, session
from ufkensa.kensa.flow import Stop
from ufkensa.logging_utils import get_logger

from ..shell import shell_context

log = get_logger("app.routes.quality")

bp = Blueprint("quality", __name__)

#: ロットNo(英数字。UF1.Lot と同じく大文字にする)
_LOT_RE = re.compile(r"^[0-9A-Za-z]{1,20}$")


@bp.get("/quality")
def page():
    return render_template("quality.html", **shell_context("quality"))


@bp.get("/api/quality/current")
def current():
    """通常検査でいま検査中のロット(読むだけ)。"""
    return jsonify(quality_session.current_lot())


@bp.get("/api/quality/lot")
def lot():
    """ロットの品質基準。lot を省くといま検査中のロット。"""
    lot_no = (request.args.get("lot") or "").strip()
    if lot_no and not _LOT_RE.match(lot_no):
        return jsonify({"error": {"code": "bad_value", "message": "ロットNo は英数字で入れてください", "field": "lot"}}), 400
    k = session.get()
    with trace.timed("quality.lot", "品質基準画面: ロットの品質基準を引く", expected=(Stop,),
                     lot=lot_no.upper() or k.state.lot or None):
        data = quality_session.build(lot_no, store=k.store, kensa=k)
    return jsonify(data)
