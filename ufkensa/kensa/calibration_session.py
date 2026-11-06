"""校正画面(VBA の UF校正)のセッション ── 画面(app/routes/calibration.py)はここだけを呼ぶ

判断は `calibration.py`(Proofreading / UF校正 の移植)が行い、ここは
- 手元DB(通常検査と同じ `LocalStore`)の「校正記録」「品証検証」を読んで部品に渡し、
- 部品が返した「書くセル」「赤字にするセル」「ログ」を手元DB・ログへ反映し、
- 画面に出す値(view)を作る。

通常検査の状態(`session.get().state`)は**読むだけ**(ライン名・検査員・DBflag)。D11 により、通常検査は
入力のたびに校正記録を見て判定するので、ここで記録すれば検査画面のブロックは自動で外れる。

各操作と、対応する VBA のイベント:
    open()          UF1.CommandButton6_Click(手で開く。検査員は登録日時が一番新しい直)+ UF校正.Initialize / Activate
                    (品証ランク付け)。校正_Open(未校正で自動で開く)と同じ画面
    receive(text)   INP1000_AfterUpdate(緑 = 1000mm)/ INPG_AfterUpdate(白 = ストレートエッヂ)。
                    測定器の伝送文字列(DT… / #L…)。合格で INPG_Exit(閉じて検査へ戻る)
    clear()         CBcl_Click「ｸﾘｱ＆ﾘｾｯﾄ」
    skip_1000()     CommandButton5_Click「送り」(B14: 直すと押せない)
    rank(kind)      CommandButton4_Click「品証順位」/ CommandButton6_Click「校正順位」

移行のための補い(VBA には無い): 「品証検証」シートは xlsm の取り込み対象に無かったため、手元DB に無ければ
VBA のブックと同じ枠(5 行目の見出し、6〜98 行目の「d日」×「1直/2直/3直」)を作る。VBA はシートが無いと End。
"""
from __future__ import annotations

import re
import threading
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Callable, Optional

from .. import trace
from . import calibration, datasource, legacy_log, lines, quality, quirks, terminal, vbx, workers, wiring
from .calibration import GREEN, WHITE, CalibrationState, RankResult, StepResult
from .flow import ERROR, Message, Stop
from .store import GridSheet, LocalStore

KIROKU = "校正記録"
HINSHO = "品証検証"
#: 赤字 RGB(255,128,128)。手元DB のセルの色は 1 つしか持たないので、この 2 シートでは文字の色として使う
RED_RGB = 255 + 128 * 256 + 128 * 65536
RECORD_ROWS = 20                    # 画面に出す校正記録(新しい順)
SKIP_TEXT = "この端末は校正をとばせる設定です(権限: 校正をとばせる)"

#: 品証検証 の枠(VBA のブックのまま。A1/A3/A4 の注記は判定に使わないので見出しだけ)
HINSHO_HEADERS = {(5, 1): "日付", (5, 3): "デジタル表示(1000㎜)", (5, 6): "スケール目盛読み取り(1000㎜)", (5, 9): "ライン",
                  (5, 10): "測定者", (5, 11): "測定日時", (5, 13): "デジタル表示(1000㎜)", (5, 14): "ライン",
                  (5, 15): "測定者", (5, 16): "測定日時"}


@dataclass
class CalResult:
    """操作の返り値(検査画面の Result と同じ形)。"""
    view: dict[str, Any]
    messages: list[Message] = field(default_factory=list)
    stop: Optional[Stop] = None

    def to_dict(self) -> dict[str, Any]:
        return {"view": self.view, "messages": [m.to_dict() for m in self.messages], "question": None,
                "stop": self.stop.to_dict()["stop"] if self.stop else None}


# ------------------------------------------------------------------
# 小道具
# ------------------------------------------------------------------
def _cdate(text: str) -> float:
    """`Day1 = Replace(UF1.Day1, ".", "/")`(Date へ代入)。読めなければ On Error Resume Next で 0 のまま。"""
    m = re.match(r"^\s*(\d{4})/(\d{1,2})/(\d{1,2})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?\s*$", text)
    if not m:
        return 0.0
    try:
        d = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4) or 0), int(m.group(5) or 0),
                     int(m.group(6) or 0))
    except ValueError:
        return 0.0
    return d.toordinal() + (d.hour * 3600 + d.minute * 60 + d.second) / 86400


def manual_kname(reg: workers.WorkerRegistry) -> str:
    """UF1.CommandButton6_Click (UF1:3030-3045)。手で開くときの検査員 = 登録日時が一番新しい直の T の末尾の人。

    `Select Case Max(Day1, Day2, Day3)` は Day1 → Day2 → Day3 の順に当てる。登録が無い(どれも 0)と Day1 に当たり、
    `Worker1(UBound(Worker1))`(Split("") は空の配列)で添字エラー → On Error Resume Next で Kname は "" のまま。
    """
    days = [_cdate(d.replace(".", "/")) if d != "" else 0.0 for d in (reg.day1, reg.day2, reg.day3)]
    mx = max(days)
    for d, t in zip(days, (reg.t1, reg.t2, reg.t3)):
        if d == mx:
            return t.split("_")[-1] if t != "" else ""
    return ""


def _time_text(v: Any) -> str:
    """校正記録 C 列。xlsm から取り込んだ行は Excel が時刻(日の割合)にした数、新しい行は "h時mm分ss秒"。"""
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 0 <= v < 1:
        s = int(round(v * 86400))
        return f"{s // 3600}時{s % 3600 // 60:02d}分{s % 60:02d}秒"
    return calibration._text(v)


def _num_text(v: Any) -> str:
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return vbx.fmt(v, "0.0")
    return calibration._text(v)


def ensure_hinsho_sheet(store: LocalStore) -> bool:
    """品証検証 が無ければ VBA のブックと同じ枠を作る(移行の補い。作ったら True)。"""
    if store.exists(HINSHO):
        return False
    sh = store.ensure_sheet(HINSHO)
    for (r, c), v in HINSHO_HEADERS.items():
        sh.write(r, c, v, "S")
    for day in range(1, 32):
        r = 6 + (day - 1) * 3
        sh.write(r, 1, f"{day}日", "S")
        for i, shift in enumerate(("1直", "2直", "3直")):
            sh.write(r + i, 2, shift, "S")
    return True


# ------------------------------------------------------------------
# セッション
# ------------------------------------------------------------------
class CalibrationSession:
    """プロセスに 1 つの校正画面。操作は `_lock` で 1 つずつ通す。

    試験では `store`・`now`・`facility`・`shifts`・`skip`・`kensa`(通常検査のセッション)・`write_logs` を差し替える。
    """

    def __init__(self, store: Optional[LocalStore] = None, *, now: Optional[Callable[[], datetime]] = None,
                 facility: Callable[[], str] = terminal.facility, shifts: Optional[workers.ShiftTimes] = None,
                 skip: Optional[Callable[[], bool]] = None, kensa: Any = None, write_logs: bool = True) -> None:
        self._lock = threading.RLock()
        self._store = store
        self._now = now
        self.facility_fn = facility
        self._shifts = shifts
        self._skip = skip
        self._kensa = kensa
        self.write_logs = write_logs
        self.state: Optional[CalibrationState] = None
        self.rank_kind = ""                 # 最後に出した順位("品証" / "校正")
        self.template_created = False

    # --- 土台 -----------------------------------------------------------
    def kensa(self) -> Any:
        if self._kensa is not None:
            return self._kensa
        from . import session
        return session.get()

    @property
    def store(self) -> LocalStore:
        if self._store is None:
            self._store = self.kensa().store
        return self._store

    def now(self) -> datetime:
        return (self._now or self.kensa().now)()

    def facility(self) -> str:
        return self.facility_fn() or workers.DEFAULT_FACILITY

    def shifts(self) -> workers.ShiftTimes:
        if self._shifts is None:
            try:
                self._shifts = workers.load_shift_times()
            except Exception:                           # noqa: BLE001 - 読めなければ Env と同じく空の時刻
                self._shifts = workers.ShiftTimes()
        return self._shifts

    def skip(self) -> bool:
        return (self._skip or self.kensa().calibration_skip)()

    def _kstate(self) -> Any:
        return self.kensa().state

    def _line_name(self) -> str:
        """UF校正.LineName = UF1.LineName & UF1.LineNo。検査画面がまだ起動していなければ端末の設定から。"""
        st = self._kstate()
        if st.line_name != "":
            return st.lname
        fac = self.facility_fn()
        return (lines.vba_name(fac) + terminal.line_no()) if fac else ""

    def _log(self, entries: list[legacy_log.LogEntry]) -> None:
        if not self.write_logs:
            return
        for e in entries:
            e.write(now=self.now())

    def _cells(self, name: str) -> Optional[dict[tuple[int, int], Any]]:
        return wiring.sheet_cells(self.store, name)

    def _red_rows(self, name: str, col: int = 1) -> set[int]:
        if not self.store.exists(name):
            return set()
        sh = self.store.sheet(name)
        assert isinstance(sh, GridSheet)
        rows = self.store.conn.execute('SELECT "行" FROM "セル" WHERE "シート" = ? AND "列" = ? AND "色" = ?',
                                       (name, col, RED_RGB)).fetchall()
        return {r[0] for r in rows}

    # --- 手元DB へ反映 ----------------------------------------------------
    def _paint(self, name: str, cells: list[tuple[int, int]]) -> None:
        sh = self.store.ensure_sheet(name)
        assert isinstance(sh, GridSheet)
        for r, c in cells:
            sh.set_color(r, c, RED_RGB)

    def _apply_rank(self, name: str, rank: RankResult) -> None:
        """校正ランク付け / 品証ランク付け がシートに書くもの(列のクリア → 値の書き直し → 写し)。"""
        sh = self.store.ensure_sheet(name)
        lo, hi = rank.clear_columns
        for r, c, _ in sh.used():
            if lo <= c <= hi:
                sh.set(r, c, None)
        for w in rank.rewrites + rank.writes:
            sh.write(w.row, w.col, w.value)

    def _apply(self, res: StepResult) -> None:
        if not (res.kiroku_writes or res.kiroku_red or res.kiroku_create or res.hinsho_writes or res.hinsho_red
                or res.kiroku_rank):
            return
        with trace.timed("db.save", "校正の記録(校正記録・品証検証)", path=str(self.store.path)) as info:
            if res.kiroku_create:
                self.store.ensure_sheet(KIROKU)
            if res.hinsho_writes:
                wiring.write_cells(self.store, HINSHO, res.hinsho_writes)
            if res.hinsho_red:
                self._paint(HINSHO, res.hinsho_red)
            if res.kiroku_rank is not None:
                self._apply_rank(KIROKU, res.kiroku_rank)
            if res.kiroku_writes:
                wiring.write_cells(self.store, KIROKU, res.kiroku_writes)
            if res.kiroku_red:
                self._paint(KIROKU, res.kiroku_red)
            info["cells"] = len(res.kiroku_writes) + len(res.hinsho_writes)
        trace.note_used(self.store.path, "save")

    # --- 操作 -----------------------------------------------------------
    def _run(self, fn: Callable[..., list[Message]], *args: Any) -> CalResult:
        with self._lock:
            try:
                try:
                    msgs = fn(*args)
                except datasource.DbMissing as e:              # VBA は End(D10: 止めるが状態は保つ)
                    raise Stop(str(e), source="datasource") from e
                st = self.view().get("status") or {}
                if st.get("error"):                              # 直の時刻が読めない(時間用が無い)
                    raise Stop(st["error"], source="Worker.Set_Shift")
            except Stop as e:
                return CalResult(self.view(), list(e.messages), stop=e)
            return CalResult(self.view(), msgs)

    def open(self) -> CalResult:
        return self._run(self._open)

    def _open(self) -> list[Message]:
        """UF1.CommandButton6_Click + UF校正.Initialize(UF校正:342-391)/ Activate(UF校正:314-339)。"""
        line = self._line_name()
        if line == "":
            raise Stop("ライン設定がありません" + "\r\n" + "担当者に連絡して下さい", source="UF1:225")
        st = self._kstate()
        kname = manual_kname(wiring.registry(st))
        self.state = calibration.initialize(line, kname, st.line_name or line)
        self.rank_kind = ""
        self.template_created = ensure_hinsho_sheet(self.store)
        if not self.store.exists(KIROKU):
            # VBA は UF校正 を開く前に 校正()(Proofreading:28-61)が AddWS している。無いまま測ると 測定完了 が
            # シートを作るだけで記録しない(Proofreading:81-84)ので、開くときに作っておく
            self.store.ensure_sheet(KIROKU)
        hinsho = self._cells(HINSHO)
        if hinsho is not None:
            rank = calibration.hinsho_rank(hinsho, now=self.now())          # Initialize の 品証ランク付け
            with trace.timed("db.save", "品証ランク付け(品証検証)", path=str(self.store.path)):
                self._apply_rank(HINSHO, rank)
            self.state.ranks.update(rank.ranks)
            self.rank_kind = "品証"
        self._log([calibration.activate_log(self.state)])
        return []

    def _need_state(self) -> CalibrationState:
        if self.state is None:
            self._open()
        assert self.state is not None
        return self.state

    def receive(self, text: str) -> CalResult:
        return self._run(self._receive, text)

    def _receive(self, text: str) -> list[Message]:
        state = self._need_state()
        now = self.now()
        if state.mode == GREEN:
            res = calibration.input_1000(state, text, self._cells(HINSHO), self._cells(KIROKU), now=now,
                                         facility=self.facility(), shifts=self.shifts(),
                                         red_rows=self._red_rows(KIROKU))
            if res.kiroku_rank is not None:
                self.rank_kind = "校正"
        else:
            res = calibration.input_gauge(state, text, self._cells(KIROKU), self._straight_edge(),
                                          now=now, facility=self.facility(), shifts=self.shifts())
        self._apply(res)
        self._log(res.logs)
        calibration.exit_gauge(state)                    # INPG_Exit: 合格(F1="end")なら閉じて検査へ
        return res.messages

    def _straight_edge(self) -> Optional[tuple[Any, Any]]:
        """QP_DB.StraightEdge(設備名)。"" (行が無い)・None(列が取れない)は VBA の IsEmptyArray → 記録しない。"""
        se = quality.straight_edge(self.facility())
        return se if isinstance(se, tuple) else None

    def straight_edge_info(self) -> Optional[dict[str, str]]:
        """画面に出す 校正ストレートエッヂ の行(長さ・許容範囲・更新日。判定は `_straight_edge` の 長さ下〜長さ上)。"""
        try:
            se = self._straight_edge()
        except datasource.DbMissing:
            return None
        if se is None:
            return None
        out = {"lo": calibration._text(se[0]), "hi": calibration._text(se[1]), "length": "", "tol": "", "updated": ""}
        try:
            db, f = quality._transfer_fields("校正ストレートエッヂ")
            if f is not None:
                for r in datasource.get_records(db, "校正ストレートエッヂ") or []:
                    if lines.matches(quality.cell(r, quality.col_of(f, "ライン名")), self.facility()):
                        out.update(length=quality.vstr(quality.cell(r, quality.col_of(f, "長さ"))),
                                   tol=quality.vstr(quality.cell(r, quality.col_of(f, "許容範囲"))),
                                   updated=quality.vstr(quality.cell(r, quality.col_of(f, "更新日"))))
                        break
        except (datasource.DbMissing, IndexError):
            pass
        return out

    def clear(self) -> CalResult:
        return self._run(lambda: calibration.clear_button(self._need_state()) or [])

    def skip_1000(self) -> CalResult:
        return self._run(self._skip_1000)

    def _skip_1000(self) -> list[Message]:
        state = self._need_state()
        if state.mode != GREEN:
            return []
        if not calibration.skip_1000_button(state):
            return [Message("1000mm はとばせません(B14 を直した設定)", ERROR, source="UF校正:133")]
        self.rank_kind = ""
        return []

    def rank(self, kind: str) -> CalResult:
        return self._run(self._rank, kind)

    def _rank(self, kind: str) -> list[Message]:
        """CommandButton4「品証順位」/ CommandButton6「校正順位」。当たった順位だけ書き換える(前の表示は残る)。"""
        state = self._need_state()
        now = self.now()
        if kind == "hinsho":
            cells = self._cells(HINSHO)
            if cells is None:
                raise Stop("品証検証 シートがありません", source="Proofreading:518")
            r = calibration.hinsho_rank(cells, now=now)
            name, label = HINSHO, "品証"
        else:
            cells = self._cells(KIROKU)
            if cells is None:
                raise Stop("校正記録 シートがありません", source="Proofreading:388")
            r = calibration.calibration_rank(cells, state.line_name, now=now, red_rows=self._red_rows(KIROKU))
            name, label = KIROKU, "校正"
        with trace.timed("db.save", f"{label}ランク付け", path=str(self.store.path)):
            self._apply_rank(name, r)
        state.ranks.update(r.ranks)
        self.rank_kind = label
        return []

    # --- 画面の値 ---------------------------------------------------------
    def status(self) -> dict[str, Any]:
        """いまの(直を考慮した日付, 直)と、校正記録の最終行。D11: 毎回ここで判定する。"""
        now = self.now()
        fac = self.facility()
        shifts = self.shifts()
        cells = self._cells(KIROKU)
        skip = self.skip()
        db_flag = self._kstate().dbflag
        try:
            return self._status(now, fac, shifts, cells, skip, db_flag)
        except workers.ShiftTimesMissing as e:
            # 画面は描けるように空の値で返し、理由は error に入れる(_run が止めて知らせる)
            return {"calibrated": False, "skip": skip, "skip_text": "", "date": "", "shift": "",
                    "last_date": "", "last_shift": "", "now": now.strftime("%Y/%m/%d %H:%M"), "error": e.text}

    def _status(self, now, fac, shifts, cells, skip, db_flag) -> dict[str, Any]:  # noqa: ANN001
        chk = calibration.check_calibration(workers.hhmm(now), cells, now=now, facility=fac, shifts=shifts,
                                            jump=skip, db_flag=db_flag)
        last_d, last_t = calibration.last_record(cells) if cells else (None, None)
        return {"calibrated": not chk.uncalibrated, "skip": skip or db_flag != "",
                "skip_text": SKIP_TEXT if skip else ("DBflag があるので校正を見ません" if db_flag else ""),
                "date": workers.yea_dayr_check(workers.hhmm(now), now=now, facility=fac, shifts=shifts),
                "shift": workers.time_check(workers.hhmm(now), facility=fac, shifts=shifts),
                "last_date": calibration._text(last_d), "last_shift": calibration._text(last_t),
                "now": now.strftime("%Y/%m/%d %H:%M")}

    def records(self) -> list[dict[str, Any]]:
        """校正記録の最近の行(新しい順)。H 列「〆」= 不合格(ストレートエッヂ NG)。"""
        if not self.store.exists(KIROKU):
            return []
        sh = self.store.sheet(KIROKU)
        last = sh.last_row(1)
        red = self._red_rows(KIROKU)
        out = []
        for r in range(last, max(1, last - RECORD_ROWS), -1):
            vals = [sh.get(r, c) for c in range(1, 9)]
            if all(v is None for v in vals):
                continue
            out.append({"row": r, "date": calibration._text(vals[0]), "night": calibration._text(vals[1]) != "",
                        "time": _time_text(vals[2]), "shift": calibration._text(vals[3]), "value": _num_text(vals[4]),
                        "kname": calibration._text(vals[5]), "line": calibration._text(vals[6]),
                        "ng": calibration._text(vals[7]) == "〆", "red": r in red})
        return out

    def masters(self) -> list[dict[str, str]]:
        """画面に出す参照マスタ(ファイルと表)。"""
        from . import paths
        out = []
        try:
            db, _f = quality._transfer_fields("校正ストレートエッヂ")
            out.append({"label": "ストレートエッヂ", "path": str(db), "table": "校正ストレートエッヂ"})
        except Exception:                                   # noqa: BLE001 - 読めなくても場所だけ出す
            out.append({"label": "ストレートエッヂ", "path": paths.db_file("伝送用ファイル"), "table": "校正ストレートエッヂ"})
        out.append({"label": "直の時刻", "path": paths.db_file("伝送用ファイル")
                    + "(無ければ " + paths.db_file("伝送用ファイル", base_key="packing") + ")", "table": "時間用"})
        out.append({"label": "記録先", "path": str(self.store.path), "table": f"{KIROKU} / {HINSHO}"})
        return out

    def view(self) -> dict[str, Any]:
        st = self.state
        line = st.line_name if st else self._line_name()
        se = self.straight_edge_info() if line else None
        center = calibration.RANK_CENTER.get(line)
        test = line == "Test"
        v: dict[str, Any] = {
            "status": self.status() if line else None,
            "line": line, "facility": self.facility_fn(),
            "opened": st is not None,
            "range_1000": ["50", "3000"] if test else ["999.5", "1000.5"],
            "straight_edge": se,
            "masters": self.masters(),
            "center": vbx.fmt(center, "0.0") if center is not None else "",
            "month": f"{self.now().month}月度",
            "records": self.records(),
            "template_created": self.template_created,
            "can_skip_1000": quirks.reproduce("B14"),
        }
        if st is not None:
            step = "done" if st.f1 == "end" else ("1000" if st.mode == GREEN else "gauge")
            v.update({
                "step": step, "kname": st.kname, "gauge": st.gauge, "sokutei": st.sokutei,
                "ng_value": st.ng_value if st.ng_visible else "", "saiken": st.saiken,
                "accepting": st.f1 == "1", "f1": st.f1, "closed": st.close,
                "ranks": [{"pos": k, "text": st.ranks.get(k, "")} for k in range(1, 7)],
                "rank_kind": self.rank_kind,
            })
            v["can_skip_1000"] = v["can_skip_1000"] and st.mode == GREEN and st.f1 != "end"
        return v


_session: Optional[CalibrationSession] = None
_slock = threading.Lock()


def get() -> CalibrationSession:
    global _session
    with _slock:
        if _session is None:
            _session = CalibrationSession()
        return _session


def reset_for_test(s: Optional[CalibrationSession] = None) -> CalibrationSession:
    global _session
    with _slock:
        _session = s or CalibrationSession()
        return _session
