"""名簿・日報の画面(VBA の UFNameList)

VBA の UFNameList は 3 つの用途で開いていた(OutTolerance.NameList_Open)。
    mode 1 「検査員不一致」 … 印刷前の Name_Match で登録検査員と測定者が合わないとき(と、その印が残っている間の
                              UF1 起動時)。WorkerTime_<L> の A2:L12。「不一致ﾘｽﾄ」で O:Q 列(不一致の一覧)に切り替え
    mode 3 「梱包枚数」     … Number_Count で日報の梱包数が検入枚数を超えたとき。梱包枚数計算_<L>。「編集起動」で mode 2 へ
    mode 2 「日報閲覧」     … Daily_<L> の A2:Y。セルを直して「反映」(ListView.ListPosting)で日報へ書き戻す。
                              「選択行削除」は一覧から行を外す(書き戻すまでシートは変わらない)
読み書きするのは検査画面と同じ手元DB(`session.get().store`)。ラインは検査画面のセッションから読むだけ。
"""
from __future__ import annotations

from typing import Any, Optional

from flask import Blueprint, jsonify, render_template, request

from ufkensa import trace
from ufkensa.kensa import legacy_log, records, session, workers
from ufkensa.kensa.flow import Question, Stop
from ufkensa.kensa.store import cell_text, line_sheet
from ufkensa.logging_utils import get_logger

from ..shell import shell_context

log = get_logger("app.routes.meibo")

bp = Blueprint("meibo", __name__)

CONTENT = {1: "検査員不一致", 2: "日報閲覧", 3: "梱包枚数"}
NO_LINE = "ラインが決まっていません。検査画面を先に開いてください。"


@bp.get("/meibo")
def page():
    return render_template("meibo.html", **shell_context("meibo"))


def _range(sh: Any, r1: int, c1: int, r2: int, c2: int) -> list[list[str]]:
    """`sh.Range(Cells(r1,c1), Cells(r2,c2))` を文字にして(Range は向きを問わない)。ListView の SubItems と同じ文字。"""
    ra, rb = min(r1, r2), max(r1, r2)
    ca, cb = min(c1, c2), max(c1, c2)
    return [[cell_text(sh.get(r, c)) for c in range(ca, cb + 1)] for r in range(ra, rb + 1)]


def build(mode: int, list_kind: str = "open") -> dict[str, Any]:
    """UFNameList に出す表(NameList_Open / CB1_Click)。シートが無ければ rows は空で message を付ける。"""
    s = session.get()
    lname = s.state.lname
    out: dict[str, Any] = {"mode": mode, "content": CONTENT[mode], "line": lname, "head": [], "rows": [],
                           "editable": mode == 2, "list": list_kind, "message": ""}
    if lname == "":
        out["message"] = NO_LINE
        return out
    store = s.store
    if mode == 1:
        name = line_sheet(records.WORKER_TIME, lname)
        if not store.exists(name):
            out["message"] = f"{name} がありません"
            return out
        sh = store.sheet(name)
        if list_kind == "mismatch":                                     # CB1「不一致ﾘｽﾄ」(UFNameList:362-370)
            xyz = max(sh.last_row(15), sh.last_row(16), sh.last_row(17))
            out["head"] = _range(sh, 1, 15, 1, 17)[0]
            out["rows"] = _range(sh, 2, 15, xyz, 17)
        else:
            # NameList_Open(1) は A3:L12(最終行のﾊﾟｹｯﾄﾃﾞｰﾀを除くため固定)、CB1「全ﾘｽﾄ」は A3:L(A 列の最終行)
            out["head"] = _range(sh, 2, 1, 2, 12)[0]
            out["rows"] = _range(sh, 3, 1, sh.last_row(1) if list_kind == "all" else 12, 12)
    elif mode == 2:
        name = line_sheet(records.DAILY, lname)
        if not store.exists(name):
            out["message"] = f"{name} がありません"
            return out
        sh = store.sheet(name)
        out["head"] = _range(sh, 2, 1, 2, 25)[0]
        out["rows"] = _range(sh, 3, 1, sh.last_row(6), 25)
    else:
        name = line_sheet(records.PACK_COUNT, lname)
        if not store.exists(name):
            out["message"] = f"{name} がありません"
            return out
        sh = store.sheet(name)
        out["head"] = _range(sh, 1, 1, 1, max(sh.last_col(1), 1))[0]
        out["rows"] = _range(sh, 2, 1, sh.last_row(1), 4)
    return out


@bp.get("/api/meibo/view")
def view():
    try:
        mode = int(request.args.get("mode", "3"))
    except ValueError:
        mode = 3
    if mode not in CONTENT:
        return jsonify({"error": {"code": "bad_value", "message": "その表示はありません"}}), 400
    list_kind = request.args.get("list", "open")
    with trace.timed("meibo.view", f"名簿・日報: {CONTENT[mode]}", expected=(Stop,)):
        out = build(mode, list_kind)
        if request.args.get("open") == "1" and out["line"] != "":
            # UFNameList.UserForm_Activate のログ(UFNameList:304-320)
            legacy_log.line_entry(out["line"], "*UFKen:UFNameList起動:" + out["content"]).write()
    return jsonify({"view": out, "messages": [], "question": None, "stop": None})


def write_back(rows: list[list[str]], answers: dict[str, bool]) -> Optional[Question]:
    """UFNameList.CommandButton2_Click「反映」(UFNameList:261-276) → ListView.ListPosting (ListView:504-521)。

    日報_<L> の 3 行目から、行数 + 20 行目まで・2 行目の最終列までを消して、一覧の行を書く。値は Split の結果
    (String の配列)を入れるので文字のまま(空は空セル)。
    """
    ans = answers.get("daily")
    if ans is None:
        return Question("daily", "日報ﾃﾞｰﾀを修正しますか？", source="UFNameList:264")
    if not ans:
        return None
    s = session.get()
    lname = s.state.lname
    if lname == "":
        raise Stop(NO_LINE, source="meibo")
    name = line_sheet(records.DAILY, lname)
    if not s.store.exists(name):                                        # WsCheck = False → Exit Sub
        return None
    sh = s.store.sheet(name)
    wb = workers.list_posting([[("" if v is None else str(v)) for v in r] for r in rows], lname, sh.last_col(2))
    trace.note_used(s.store.path, "save")
    with trace.timed("db.save", "日報を直す(ListPosting)", path=str(s.store.path), table=name) as info:
        r1, c1, r2, c2 = wb.clear
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                sh.set(r, c, None)
        for w in wb.writes:
            sh.write(w.row, w.col, w.value, "S")
        info["rows"] = len(rows)
    return None


@bp.post("/api/meibo/daily")
def daily():
    body = request.get_json(silent=True) or {}
    rows = body.get("rows") or []
    answers = body.get("answers") or {}
    if not isinstance(rows, list) or not all(isinstance(r, list) for r in rows):
        return jsonify({"error": {"code": "bad_value", "message": "行の形が正しくありません"}}), 400
    try:
        with trace.timed("meibo.daily", "名簿・日報: 日報の反映", expected=(Stop,)):
            q = write_back(rows, answers)
    except Stop as e:
        return jsonify({"view": build(2), "messages": [], "question": None, "stop": e.to_dict()["stop"]})
    return jsonify({"view": build(2), "messages": [], "question": q.to_dict()["question"] if q else None,
                    "stop": None})
