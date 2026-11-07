"""保守 → 履歴出力。検査済みの履歴(保存用・耳用)から、印刷のときと同じ PDF・xlsx・CSV を出し直す。

検査を最後まで済ませなくても、PDF(Edge)と共有フォルダへの保存を試せるようにするためのもの(2026-11-07 のご依頼)。

- PDF 出力: 選んだ行を、手元DB の写し(一時ファイル)の上で検査画面に読み戻し(UF管理 の読み込み)、帳票を作って
  印刷のときと同じ場所へ PDF(履歴保存)と履歴ブック xlsx を書く。いま検査中の画面・手元DB は触らない。
  保存用・測定DB への登録はしない(出し直しなので)。
- CSV 出力: 印刷のときと同じ 履歴 CSV(保存用 / 耳用)と Lot情報履歴 の CSV を、手元と共有へ。
- テスト履歴: 試験用のロット 1 件分(保存用・Lot情報履歴・作業保存など)を手元DB に入れる(`data/test_history.json`)。
"""
from __future__ import annotations

import json
import shutil
import sqlite3
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional

from .. import trace
from . import checks, records, report_model, wiring
from .flow import Message, Stop
from .store import HOZON, LOT_HISTORY, MIMI, LocalStore

TEST_HISTORY_FILE = Path(__file__).with_name("data") / "test_history.json"
#: テスト履歴の印(保存用の行の「検査員」欄など。本物の履歴と見分ける)
TEST_MARK = "テスト履歴"


def _copy_store(store: LocalStore, folder: Path) -> LocalStore:
    """手元DB の写し(書きかけの WAL も含めて写すため sqlite の backup を使う)。"""
    dst = folder / "kensa_local_copy.sqlite3"
    with store._lock:                                         # noqa: SLF001 - 同じ接続で写す
        out = sqlite3.connect(str(dst))
        try:
            store.conn.backup(out)
        finally:
            out.close()
    return LocalStore(dst)


def export_pdf(store: LocalStore, row: int, *, mimi: bool = False,
               session_factory: Optional[Callable[[LocalStore], Any]] = None,
               now: Optional[datetime] = None) -> dict[str, Any]:
    """選んだ履歴の帳票 PDF と履歴ブック xlsx を出し直す。返り値: {"lot", "pdf": [...], "xlsx": [...], "messages"}。"""
    from . import session as sess_mod
    tmp = Path(tempfile.mkdtemp(prefix="ufkensa_histout_"))
    copy = _copy_store(store, tmp)
    try:
        # 写しでは検査中の作業保存を外す(残すと「別Lotデータがあります」で読み戻せない。本物の手元DB はそのまま)
        for name in copy.sheet_names():
            if name.startswith(records.WORK_SAVE + "_"):
                copy.delete_sheet(name)
        make = session_factory or (lambda st: sess_mod.Session(st, measure_db=tmp / "unused.sqlite3",
                                                                **({"now": (lambda: now)} if now else {})))
        s = make(copy)
        res = s.history_restore(row, mimi)
        if res.stop is not None:
            raise Stop(res.stop.text, source="履歴出力")
        st, env = s.state, s._env()                                   # noqa: SLF001
        msgs: list[Message] = list(res.messages)
        snap = wiring.to_snapshot(st, copy)
        if not any((snap.t1, snap.t2, snap.t3)) or snap.worker_shift not in ("1", "2", "3"):
            # 出し直しでは検査員の登録が残っていないことが多い。保存用の「検査員」欄(77 列目)の人で帳票を作る
            hrow = copy.sheet(MIMI if mimi else HOZON).row(row)
            who = records.cell_text(hrow[76]) if len(hrow) > 76 else ""
            snap.worker_shift, snap.t1 = "1", snap.t1 or who or TEST_MARK
        shifts = env.shift_times()
        penv = report_model.PrintEnv(now=env.now(), lateral_check=checks.lateralbow_check(wiring.to_uf1(st)),
                                     machine_transfer=s._machine_transfer,     # noqa: SLF001
                                     start1=shifts.start1, end3=shifts.end3)
        with trace.timed("history.export", "履歴出力: 帳票を作る", lot=st.lot, row=row):
            pr = report_model.print_transfer(copy, snap, penv)
        msgs.extend(pr.messages)
        if pr.report is None:
            raise Stop("帳票を作れませんでした" + (": " + pr.messages[-1].text if pr.messages else ""),
                       source="履歴出力")
        env.messages.clear()
        pdf = s._history_pdf(pr.report, pr.defect)                    # noqa: SLF001 - 印刷と同じ置き場所
        msgs.extend(env.messages)
        xlsx: list[Path] = []
        try:
            xlsx = records.export_lot_book(copy, st.lname, st.line_name, st.lot, env.now(), pr.report, pr.defect)
        except OSError as e:
            msgs.append(Message(f"履歴ブック(xlsx)を書き出せませんでした: {e}", "warn", source="履歴出力"))
        return {"lot": st.lot, "line": st.line_name,
                "pdf": [f for d in pdf for f in d.values() if isinstance(f, str)],
                "xlsx": [str(p) for p in xlsx],
                # 読み戻しの案内(「Load完了…」など)は出さず、注意と誤りだけ
                "messages": _notices(msgs)}
    finally:
        copy.close()
        shutil.rmtree(tmp, ignore_errors=True)


def _notices(msgs: list[Message]) -> list[dict[str, Any]]:
    """画面に出す知らせ: 注意と誤りだけ、同じ文は 1 回。写しの上の作業中保存の失敗は出さない(本物は触らないので)。"""
    out, seen = [], set()
    for m in msgs:
        if m.level not in ("warn", "error") or m.text.startswith("作業中の保存ができませんでした") or m.text in seen:
            continue
        seen.add(m.text)
        out.append(m.to_dict())
    return out


def export_csv(store: LocalStore, lname: str, line: str, *, mimi: bool = False) -> dict[str, Any]:
    """印刷のときと同じ 履歴 CSV と Lot情報履歴 CSV を、手元と共有へ書く。"""
    msgs: list[Message] = []
    files: list[Path] = []
    if not store.exists(MIMI if mimi else HOZON):
        raise Stop(f"{MIMI if mimi else HOZON} に履歴がありません", source="履歴出力")
    with trace.timed("history.export", "履歴出力: CSV", line=line):
        try:
            files += records.export_history_csv(store, lname, line, mimi, local_dir=store.path.parent)
        except OSError as e:
            msgs.append(Message(f"履歴 CSV を書き出せませんでした: {e}", "warn", source="履歴出力"))
        try:
            files += records.export_lot_history(store, line, local_dir=store.path.parent)
        except OSError as e:
            msgs.append(Message(f"Lot情報履歴を書き出せませんでした: {e}", "warn", source="履歴出力"))
    return {"files": [str(p) for p in files], "messages": [m.to_dict() for m in msgs]}


# ----------------------------------------------------------------------
# テスト履歴
# ----------------------------------------------------------------------
def load_test_history(store: LocalStore, lname: str) -> dict[str, Any]:
    """テスト履歴(ロット 1 件分)を手元DB に入れる。シート名の <L> はこの端末のライン名に置き換える。

    保存用・Lot情報履歴 は行を足す(同じテスト履歴が入っていれば足さない)。ライン別のシート(作業保存_<L> など)は
    検査中のロットを壊さないよう、まだ無いときだけ作る。
    """
    data = json.loads(TEST_HISTORY_FILE.read_text(encoding="utf-8"))
    lot = data["lot"]
    added: list[str] = []
    with store.transaction():
        for name in (HOZON, LOT_HISTORY):
            rows = data["rows"].get(name) or []
            if not rows:
                continue
            sh = store.sheet(name)
            have = {records.cell_text(v) for _, r in sh.rows() for v in r[:3]}
            if lot in have:
                continue
            for r in rows:
                sh.append(r)
            added.append(f"{name}: {len(rows)} 行")
        for base, cells in data.get("line_sheets", {}).items():
            name = f"{base}_{lname}"
            if store.exists(name):
                continue
            sh = store.ensure_sheet(name)
            for key, v in cells.items():
                r, c = (int(x) for x in key.split(","))
                sh.write(r, c, v, "O")
            added.append(name)
    return {"lot": lot, "added": added}


def remove_test_history(store: LocalStore) -> dict[str, Any]:
    """テスト履歴の行(保存用・Lot情報履歴)を消す。ライン別のシートは検査で使うので消さない。"""
    data = json.loads(TEST_HISTORY_FILE.read_text(encoding="utf-8"))
    lot = data["lot"]
    removed: list[str] = []
    with store.transaction():
        for name in (HOZON, LOT_HISTORY):
            if not store.exists(name):
                continue
            sh = store.sheet(name)
            hits = [n for n, r in sh.rows() if lot in {records.cell_text(v) for v in r[:3]}]
            for n in sorted(hits, reverse=True):
                sh.delete_row(n)
            if hits:
                removed.append(f"{name}: {len(hits)} 行")
    return {"lot": lot, "removed": removed}


def has_test_history(store: LocalStore) -> bool:
    lot = json.loads(TEST_HISTORY_FILE.read_text(encoding="utf-8"))["lot"]
    if not store.exists(HOZON):
        return False
    return any(records.cell_text(r[1]) == lot for _, r in store.sheet(HOZON).rows())


def json_lot() -> str:
    return json.loads(TEST_HISTORY_FILE.read_text(encoding="utf-8"))["lot"]
