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


# ------------------------------------------------------------------
# 品質写真(VBA 補助.AddPagesForSubFolders / LoadImagesIntoMultiPageFromTemp。2026-11-05)
# 使用パス「品質写真」のサブフォルダごとにタブ、中の .jpg を出す。配るのはそのフォルダの中だけ
# ------------------------------------------------------------------
_PHOTO_EXT = (".jpg", ".jpeg", ".png", ".gif", ".bmp")


def _photo_root():
    from pathlib import Path
    from ufkensa.kensa import paths
    return Path(paths.to_local(paths.get("quality_photo")))


@bp.get("/api/quality/photos")
def photos():
    from ufkensa.kensa import paths
    root = _photo_root()
    folder = paths.get("quality_photo")
    try:
        if not root.is_dir():
            return jsonify({"folder": folder, "tabs": [], "message": f"品質写真のフォルダがありません: {folder}"})
        tabs = []
        for d in sorted((x for x in root.iterdir() if x.is_dir()), key=lambda x: x.name):
            imgs = sorted(f.name for f in d.iterdir() if f.is_file() and f.suffix.lower() in _PHOTO_EXT)
            tabs.append({"name": d.name, "images": imgs})
        return jsonify({"folder": folder, "tabs": tabs, "message": "" if tabs else "品質写真のフォルダにサブフォルダがありません"})
    except OSError as e:
        log.warning("品質写真を読めません: %s", e)
        return jsonify({"folder": folder, "tabs": [], "message": f"品質写真を読めません: {e}"})


@bp.get("/api/quality/photo")
def photo():
    from flask import abort, send_file
    root = _photo_root()
    sub, name = request.args.get("dir", ""), request.args.get("name", "")
    if not sub or not name or any(x in (sub + name) for x in ("/", "\\", "..")):
        abort(400)
    path = root / sub / name
    if path.suffix.lower() not in _PHOTO_EXT or not path.is_file():
        abort(404)
    return send_file(path)
