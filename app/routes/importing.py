"""取り込み画面(Access → SQLite。D4)

取り込みは長時間処理(`jobs`)で走らせ、帯の「いま動いているもの」に出す。
取り込めるのは管理者(ufkensa:admin)だけ。
"""
from __future__ import annotations

from flask import Blueprint, jsonify, render_template, request

from ufkensa import auth, jobs
from ufkensa.kensa import importing
from ufkensa.logging_utils import get_logger

from ..shell import shell_context

log = get_logger("app.routes.importing")

bp = Blueprint("importing", __name__)


def _err(code: str, message: str, status: int = 400):
    return jsonify({"error": {"code": code, "message": message}}), status


@bp.get("/import")
def page():
    return render_template("import.html", **shell_context("import"))


@bp.get("/api/import/plan")
def plan():
    ok, why = importing.can_run_here()
    return jsonify({"rows": [r.to_dict() for r in importing.plan()], "admin": auth.can_edit(),
                    "can_run": ok, "why": why})


@bp.post("/api/import/run")
def run():
    if not auth.can_edit():
        return _err("not_admin", auth.MSG_NO_LEVEL, 403)
    ok, why = importing.can_run_here()
    if not ok:
        return _err("cannot_run", why)
    body = request.get_json(silent=True) or {}
    names = [n for n in (body.get("names") or []) if n in importing.TARGETS]
    if not names:
        return _err("bad_value", "取り込む DB を選んでください")
    sources = body.get("sources") or {}

    def work(progress):
        done = []
        for i, name in enumerate(names):
            def step(pct, msg, i=i, name=name):
                progress(int((i * 100 + pct) / len(names)), f"{name}: {msg}")
            res = importing.run(name, str(sources.get(name, "") or ""), progress=step)
            done.append(f"{name}({res.summary()})")
        return "取り込みました: " + "、".join(done)

    try:
        job = jobs.get_registry().start("import", "Access の取り込み", work)
    except jobs.JobBusy as e:
        return _err("busy", f"ほかの処理が動いています: {e}", 409)
    return jsonify({"job": job.id})
