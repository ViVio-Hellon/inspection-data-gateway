"""検査の流れ(VBA の UF1 のイベント)をまとめる入口 ── 画面(app/routes/kensa.py)はここだけを呼ぶ

1 台の PC は 1 ラインを受け持ち、検査中のロットは 1 つ(VBA と同じ)。状態(`grid.KensaState`)は
プロセスに 1 つ持ち、操作のたびに手元 DB(`store.py`)へ書く(VBA の作業中保存 Posting に当たる)。
アプリを開き直したら手元 DB から読み戻す(VBA の Load_Text)。部品との受け渡しは `wiring.py`。

**画面への返し方**(すべての操作で同じ形。`flow.py`):
    Result.view        いまの画面の値(`view.build(state, extra)` の形)
    Result.messages    知らせ(MsgBox・音声)。文言は VBA の原文
    Result.question    聞くこと(vbYesNo)。画面は答えを `answers` に入れて同じ操作を呼び直す
    Result.stop        止めた理由(D10: 状態は保ったまま)

各操作と、対応する VBA のイベント:
    start()                       UF1.UserForm_Initialize(ライン・作業中保存の読み戻し・直)+ Activate
    register_worker(card)         WorkButton → Worker_Change(社員カード13桁 → 名簿 → T1〜T3)→ T?_Change
    register_lot(lot)             CommandButton10 → Lot_Change → CB_Change(重複確認・AccReflec・包装仕様・歪公差)
    select_work(work)             厚板/薄板/少枚数/極厚/スラブ_Click → LSelect
    set_yama(yama)                Yama_Change
    toggle_at() / toggle_inspe()  AtButt_Click / InspeButt_Click
    receive(text)                 StartTB → INP_Change(伝送文字列 DT…/#L…/#〆…)
    click(name) / dblclick(name)  ラベルのクリック(入力先を固定)/ダブルクリック(消去と理由)。name="" は UserForm_Click
    set_value(name, value)        歪・ラテラル・直角度・BH・検入高さ・外観・不良枚数などの手入力欄、
                                  公差外の確認(UFConfi.TB1 / UFConfi.List2 / UFConfi.List1 / UFConfi.その他)
    min_max()                     CommandButton4
    print_check(answers)          PrintB_Click の前半(先行データ票・必須チェック・検査員照合・梱包枚数)
    print_commit(answers)         PrintB_Click の後半(帳票・保存用への追記・測定DBへの登録・履歴・CSV・PDF)
    new_inspection(answers)       NewKen_Click

D11: 校正の要否は入力のたびに `calibration.input_blocked` で判定する。権限 ufkensa:calibration_skip を持つ端末は
VBA の WorkButton.Tag="jump" と同じ扱い。
"""
from __future__ import annotations

import threading
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional

from .. import trace
from . import (calibration, checks, datasource, grid, input_router, judge, layout, legacy_log, lines, lot as lotmod, minmax,
               paths, quality, records, report_html, report_model, terminal, view, vbx, workers, wiring)
from .flow import ERROR, INFO, WARN, Message, Question, Stop
from .grid import KensaState
from .store import HOZON, LOT_HISTORY, MIMI, LocalStore, line_sheet
from .wiring import Env, RealHooks

#: 設備名(レジストリ)→ UF1 のライン OptionButton(UF1:230-238)
FACILITY_OPTION = {"L-1": "L1Work", "LVC": "LVC", "HVC": "HVC", "バランサー": "Bala", "トット": "Tott",
                   "Test": "Test", "丸徳": "Maru"}
NO_LINE_TEXT = "ライン設定がありません" + "\r\n" + "担当者に連絡して下さい"          # UF1:225
CONFI_CHOICES = ("公差外", "再測定", "許容", "保留")
CONFI_REASONS = ("入力位置間違い", "測定位置間違い", "測定ボタン押し間違い", "再測定", "0点確認", "板廃棄", "その他")


@dataclass
class Result:
    view: dict[str, Any]
    messages: list[Message] = field(default_factory=list)
    question: Optional[Question] = None
    stop: Optional[Stop] = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "view": self.view,
            "messages": [m.to_dict() for m in self.messages],
            "question": self.question.to_dict()["question"] if self.question else None,
            "stop": self.stop.to_dict()["stop"] if self.stop else None,
        }


def _calibration_skip_right() -> bool:
    try:
        from .. import access_rights, auth
        return auth.has(access_rights.CALIBRATION_SKIP)
    except Exception:                                  # noqa: BLE001 - 権限が読めなければ持たない扱い
        return False


class Session:
    """プロセスに 1 つの検査セッション。操作は `_lock` で 1 つずつ通す。

    試験では `store`(手元DB)・`now`(時計)・`calibration_skip`(権限)・`write_logs`・`measure_db` を差し替える。
    """

    def __init__(self, store: Optional[LocalStore] = None, *, now: Callable[[], datetime] = datetime.now,
                 calibration_skip: Callable[[], bool] = _calibration_skip_right, write_logs: bool = True,
                 measure_db: Optional[Path] = None) -> None:
        self._lock = threading.RLock()
        self.state = KensaState()
        self._store = store
        self.now = now
        self.calibration_skip = calibration_skip
        self.write_logs = write_logs
        self.measure_db = measure_db
        self.env: Optional[Env] = None
        self.hooks: Optional[RealHooks] = None
        self.started = False
        self.confi: Optional[dict[str, Any]] = None        # 出している UFConfi(公差外・理由選択)
        self.checks: list[dict[str, Any]] = []
        self.print_ready = False
        self.report: dict[str, Any] = {}
        self.calibration_required = False

    # ------------------------------------------------------------------
    # 土台
    # ------------------------------------------------------------------
    @property
    def store(self) -> LocalStore:
        if self._store is None:
            self._store = LocalStore()
        return self._store

    def _env(self) -> Env:
        if self.env is None:
            self.env = Env(store=self.store, now=self.now, facility=terminal.facility,
                           calibration_skip=self.calibration_skip, write_logs=self.write_logs)
            self.hooks = RealHooks(self.env)
        return self.env

    def _log(self, text: str, file: str = "") -> None:
        self._env().log(legacy_log.line_entry(self.state.line_name, text, file))

    def _extra(self) -> dict[str, Any]:
        st = self.state
        reg = wiring.registry(st)
        workers_list = []
        for t, day, label in ((reg.t1, reg.day1, "T1"), (reg.t2, reg.day2, "T2"), (reg.t3, reg.day3, "T3")):
            if t != "" or day != "":
                workers_list.append(f"{label}: {t}" + (f"({day})" if day else ""))
        return {"workers": workers_list, "checks": self.checks, "print_ready": self.print_ready,
                "confi": self.confi, "calibration_required": self.calibration_required,
                "report": self.report or None, "proofread": "ok" if self._calibrated() else "",
                "range": {"ops": st.ops, "ope": st.ope,
                          "labels": [st.v(f"OPS{i}").value if st.v(f"OPS{i}").visible else "" for i in range(1, 7)]}}

    def _view(self) -> dict[str, Any]:
        st = self.state
        st.proofread_flag = "ok" if self._calibrated() else ""
        return view.build(st, self._extra())

    def _calibrated(self) -> bool:
        st = self.state
        if st.line_name == "":
            return False
        env = self._env()
        cells = wiring.sheet_cells(self.store, "校正記録")
        return not calibration.input_blocked(cells, now=env.now(), facility=env.facility() or workers.DEFAULT_FACILITY,
                                             shifts=env.shift_times(), jump=self.calibration_skip(),
                                             db_flag=st.dbflag)

    def _run(self, fn: Callable[..., Any], *args: Any, posting: bool = True) -> Result:
        with self._lock:
            env = self._env()
            env.messages, env.confi, env.reason, env.proofread = [], [], None, None
            try:
                try:
                    q = fn(*args)
                except datasource.DbMissing as e:                       # VBA は End(D10: 止めるが状態は保つ)
                    raise Stop(str(e), messages=list(env.messages), source="datasource") from e
            except Stop as e:
                msgs = list(env.messages)
                for m in e.messages:
                    if not any(m is x for x in msgs):
                        msgs.append(m)
                return Result(self._view(), msgs, stop=e)
            if env.confi:
                self._open_confi(env.confi[-1])
            if env.reason is not None:
                self._open_confi(env.reason)
            if posting and isinstance(q, Question) is False and self.started:
                assert self.hooks is not None
                self.hooks.posting(self.state)
            return Result(self._view(), list(env.messages), question=q if isinstance(q, Question) else None)

    def _open_confi(self, req: checks.ConfiRequest) -> None:
        """UFConfi を出す(vbModeless。最後の 1 件が残る)。page 0 = 公差外、1 = 理由選択。"""
        self.confi = {"posi": req.posi, "n": req.n, "line_name": req.line_name, "moto": req.moto,
                      "return_moto": req.return_moto, "del_count": req.del_count, "page": req.page,
                      "ken_name": req.ken_name or "", "list2": False,
                      "choices": list(CONFI_CHOICES) if req.page == 0 else list(CONFI_REASONS)}
        if req.page == 0:
            self.env.messages.append(Message("検査ﾃﾞｰﾀに公差外有り！", WARN, title="2024.7.26_Ver",
                                             source="UFConfi.Label1"))

    def _require_line(self) -> None:
        if self.state.line_name == "":
            raise Stop(NO_LINE_TEXT, level=WARN, source="UF1:225")

    # ------------------------------------------------------------------
    # 起動
    # ------------------------------------------------------------------
    def start(self) -> Result:
        return self._run(self._start)

    def _start(self) -> None:
        """UF1.UserForm_Initialize (UF1:173-328) と UserForm_Activate の校正済み表示。"""
        fac = terminal.facility()
        if fac == "":
            raise Stop(NO_LINE_TEXT, level=WARN, source="UF1:224")
        env, hooks = self._env(), self.hooks
        assert hooks is not None
        with trace.timed("kensa.start", "検査画面の起動", expected=(Stop,), facility=fac):
            report_model.clean_old_history(self.store, env.now())       # CleanOldHistory
            st = self.state = KensaState()
            self.confi, self.checks, self.print_ready, self.report = None, [], False, {}
            st.auto, st.form_tag, st.load, st.flag_label, st.inp_tag, st.picfm_tag = "on", "STOP", "off", 1, "", "0"
            st.v("Abno").value = "異常値on"
            st.v("DistortionFrame").value = "歪公差"
            opt = FACILITY_OPTION.get(fac, "")
            if opt:
                st.line_option = opt
                maru = 0
                if opt == "Maru" and terminal.line_no() in terminal.MARU_NUMBERS:
                    maru = terminal.MARU_NUMBERS.index(terminal.line_no()) + 1
                layout.line_auto_selection(st, "", hooks, maru)
            if st.sika_set_atu == 0 and st.lot != "":
                try:
                    all_ = lotmod.sika_atu_create(st.lot)
                except Exception:                                       # noqa: BLE001 - SikaAtuCreate の End
                    all_ = None
                if all_:
                    st.sika_set_atu = vbx.single(vbx.val(all_[0]))
            st.form_tag = ""
            st.load = "on"
            self.started = True

    # ------------------------------------------------------------------
    # 検査員
    # ------------------------------------------------------------------
    def register_worker(self, card: str, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._register_worker, card, answers or {})

    def _register_worker(self, card: str, answers: dict[str, bool]) -> Optional[Question]:
        """WorkButton_Click (UF1:4743-4786) → Worker_Change (UF1:4788-4865) → T?_Change (UF1:4099-4160)。"""
        self._require_line()
        st, env, hooks = self.state, self._env(), self.hooks
        assert hooks is not None
        if st.line_option != "Bala":
            ans = answers.get("worker_register")
            if ans is None:
                return Question("worker_register", "検査員登録しますか", title="確認", source="UF1:4755")
            if not ans:
                return None
        if hooks.register_check(st):                                    # Daily_Create
            ctx = wiring.daily_context(st, env)
            res = workers.daily_create(ctx)
            wiring.write_cells(self.store, line_sheet(records.DAILY, st.lname), res.writes, recreate=res.recreate)
        # Worker_Change
        if st.flag_label != 1:
            return None
        if len(card) != 13:
            return None
        ln = lines.find(st.line_name)
        test_line = ln is not None and ln.official == "Test"
        reg0 = wiring.registry(st)
        res2 = workers.register_card(card, reg0, now=env.now(), facility=env.facility() or workers.DEFAULT_FACILITY,
                                     shifts=env.shift_times(), line=st.line_name,
                                     name_lookup=lambda c: workers.name_conv(c, test_line=test_line))
        if res2 is None:
            return None
        env.messages.extend(res2.messages)
        if res2.name == "":
            return None
        st.flag_label = 0
        wiring.set_registry(st, res2.registry)
        if res2.log is not None:
            env.log(res2.log)
        st.picfm_tag = "0"
        if judge.advance_check(st.lot, st.dbflag, arrbox=hooks._arrbox_fn(st)):
            wiring._msg_error(env, st, judge.ADVANCE_MSG, 4)
            env.sound(judge.SOUND_NEED, "UF1:4851")
        wiring.daily_report(st, env, 1)
        st.flag_label = 1
        # T?_Change(T1/T2/T3 のうち変わった欄)
        changed = {"日勤": "1", "1直": "1", "2直": "2", "3直": "3"}.get(res2.shift, "")
        if changed and st.load != "off":
            wiring.write_cells(self.store, line_sheet("検査員", st.lname), workers.worker_save_rows(res2.registry))
            day = {"1": res2.registry.day1, "2": res2.registry.day2, "3": res2.registry.day3}[changed]
            now = env.now()
            fac = env.facility() or workers.DEFAULT_FACILITY
            if vbx.mid(day, 1, 10) == workers.yea_dayr_check(workers.hhmm(now), now=now, facility=fac,
                                                               shifts=env.shift_times()):
                chk = calibration.check_calibration(workers.hhmm(now), wiring.sheet_cells(self.store, "校正記録"),
                                                    now=now, facility=fac, shifts=env.shift_times(),
                                                    jump=self.calibration_skip(), db_flag=st.dbflag)
                if chk.create_sheet:
                    self.store.ensure_sheet("校正記録")
                if chk.uncalibrated:
                    self.calibration_required = True                     # 校正_Open(校正画面へ)
                    env.sound("確認作業.wav", "UF1:4116")
        return None

    # ------------------------------------------------------------------
    # ロット
    # ------------------------------------------------------------------
    def register_lot(self, lot: str, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._register_lot, lot, answers or {})

    def _register_lot(self, lot: str, answers: dict[str, bool]) -> Optional[Question]:
        """CommandButton10_Click (UF1:2937-3002) → Lot_Change (UF1:4660-4706) → CB_Change (UF1:4494-4535)。"""
        self._require_line()
        st, env, hooks = self.state, self._env(), self.hooks
        assert hooks is not None
        if st.line_option != "Bala":
            ans = answers.get("lot_register")
            if ans is None:
                return Question("lot_register", "Lot登録しますか", title="確認", source="UF1:2949")
            if not ans:
                return None
            if st.lot != "":
                a1 = answers.get("lot_split")
                if a1 is None:
                    return Question("lot_split", "ロット分割品ですか？", title="確認", source="UF1:2951")
                if a1:
                    a2 = answers.get("lot_split_sure")
                    if a2 is None:
                        return Question("lot_split_sure", "間違いなくロット分割ですか？？", title="確認",
                                        source="UF1:2952")
                    if a2:
                        st.v("SplitFlag").value = "Lot分割"
        # line:(入力欄を空にする)
        st.lot = ""
        st.yama = ""
        st.ken_no_caption = ""
        for n in ("HizCome", "QF", "LastHot", "dataLoding", "DebugTaikaku", "DebugLate"):
            st.v(n).value = ""
        st.v("HizCome").tag = ""
        st.flatnes_text = ""
        st.machin = ""
        # Lot_KeyPress: 英数字だけ・大文字
        text = "".join(ch.upper() if "a" <= ch <= "z" else ch for ch in lot
                       if ch.isascii() and ch.isalnum())
        if len(text) != 7:
            return None
        st.lot = text
        # Lot_Change
        if st.form_tag != "STOP":
            if self.write_logs:
                legacy_log.write_log("Lot登録|" + text + "|" + st.lname, "")
            self._log("Lot登録:" + text + ":" + st.lname)
            st.machin = ""
            self._log("Lot登録:" + text + ":" + st.lname, legacy_log.MEASURED_LOG)
        # CB_Change
        if st.load == "off":
            return None
        rows = wiring.sheet_rows(self.store, HOZON) or []
        if judge.check_for_lot(rows, text, 2):
            a3 = answers.get("lot_recent")
            if a3 is None:
                return Question("lot_recent", "直近5Lot内に同じLotNoがあります" + "\r\n" + "このまま作業しますか？",
                                title="確認", source="UF1:4505")
            if not a3:
                return None
        self._acc_reflec()
        # DistArray → OpForcedSelection → FormReflectDist
        lot_data = quality.LotData(hooks.arrbox(st, "SIKALOTNOW.accdb", "板", "Fields"),
                                   hooks.arrbox(st, "SIKALOTNOW.accdb", st.lot, "Record"))
        qp = None
        try:
            qp = quality.dist_array(st.lot, st.ken_no_caption, quality.hiki_arr_create(st.lot), lot_data)
            env.messages.extend(qp.messages)
            for k, v in qp.uf1.items():
                wiring.put(st, k, v)
        except (vbx.TypeMismatch, quality.VbaRuntimeError, IndexError) as e:
            env.messages.append(Message(str(e), ERROR, source="QP.DistArray"))
        layout.op_forced_selection(st)
        if qp is not None and qp.kousa:
            r = quality.form_reflect_dist(qp.kousa, quality_visible=False, mitsubishi_visible=False)
            for k, v in r.values.items():
                wiring.put(st, k, v)
        return None

    def _acc_reflec(self) -> None:
        """SaveLoad.AccReflec (SaveLoad:56-562)。"""
        st, env, hooks = self.state, self._env(), self.hooks
        assert hooks is not None
        st.v("TaikText").value = ""
        st.lateral_text = st.lateral_text_tag = st.sumisei = st.additional = st.comment = st.reien = ""
        st.lateral_text_colored = False
        st.v("KenH").value = ""
        hist = wiring.sheet_rows(self.store, LOT_HISTORY)
        henv = lotmod.HistoryEnv(dbflag=st.dbflag, now=env.now(), over12=hooks._over12(st),
                                 uf1_test=st.line_option == "Test")
        try:
            with trace.timed("db.save", "引当の取り込み(AccReflec / HistorySave)", expected=(Stop,),
                             path=str(self.store.path), table=LOT_HISTORY) as info:
                r = lotmod.acc_reflec(wiring.to_lot_screen(st, env.facility()),
                                      check_equipment=quality.check_equipment_db, history_rows=hist,
                                      history_env=henv)
                if r.history is not None:
                    info["rows"] = self._apply_history(r.history)
        except Stop as e:                                                # 設備仕様 NG(MsgError …, 2, "END")
            wiring._msg_error(env, st, e.text, 2, "END")
            raise
        env.messages.extend(r.messages)
        wiring.apply_lot_reflect(st, r)
        if r.status != "ok":
            return
        wiring._tolerance_text(st, r)
        if r.lateral_text is not None:
            try:
                layout.lselect_lateral(st, r.ken_no)
            except layout.VBA_ERRORS:
                pass
        if st.atu != "":
            w = quality.weight(vbx.val(st.atu), vbx.val(st.hab), vbx.val(st.tak))
            layout.atu_opselect(st, vbx.single(vbx.val(st.atu)), hooks, w, env.messages)
        temp_e = r.temp_e
        uf = wiring.to_uf1(st)
        if checks.conductiv_check(uf):
            add = "ハンドの場合、導電率測定が必要です。注意願います。"
            temp_e = add if temp_e == "" else temp_e + "," + add
            cr = checks.conductiv_coolingtime_check(st.lot)
            env.messages.extend(cr.messages)
            qf = cr.value if hasattr(cr, "value") else ""
            if vbx.is_numeric(qf) or (qf not in ("", "LotNoがない")):
                st.v("ConductivL").visible = True
                st.conductiv_text = qf
            elif qf == "":
                st.v("ConductivL").visible = False
                env.messages.append(Message("指定ファイルに写真確認Noに入力がない" + "\r\n" + "品証連絡", WARN,
                                            source="SaveLoad:404"))
            else:
                st.v("ConductivL").visible = False
                env.messages.append(Message("指定ファイルにLotNo入力がない" + "\r\n" + "品証連絡", WARN,
                                            source="SaveLoad:409"))
        result = lotmod.dedupe_message(temp_e)
        if result != "":
            wiring._msg_error(env, st, result, 4)
        try:
            ord_no, msgs, _ = lotmod.materials_order_no(st.lot, r.past_hiki[0], st.dbflag, now=env.now(),
                                                         over12=hooks._over12(st))
            env.messages.extend(msgs)
            wiring.apply_judge(st, judge.materials_msgbox(wiring.to_judge_form(st), ord_no,
                                                          arrbox=hooks._arrbox_fn(st)), env)
        except (lotmod.SubscriptOutOfRange, judge.VbaError, vbx.TypeMismatch) as e:
            env.messages.append(Message(str(e), ERROR, source="Judg.MaterialsMsgBox"))
        h_cours = r.row.get("実績_設備ｺｰｽ") or ""
        gauge = quality.is_gauge_required(vbx.val(st.atu), vbx.val(st.hab), vbx.val(st.tak), st.zai, st.tyu, h_cours,
                                          env.messages)
        for p in grid.POSITIONS:
            st.v("LZ" + p).value = "幅ｹﾞｰｼﾞ" if gauge else "定盤"
            st.v("LG" + p).value = "丈ｹﾞｰｼﾞ" if gauge else "ｹﾞｰｼﾞ"
        rows = hooks.get_lot_info(st, st.lot)
        if rows is None:
            return
        try:
            layout.judge_lot_allocations(st, rows, hooks)
        except layout.VBA_ERRORS:
            pass

    def _apply_history(self, h: lotmod.HistorySaveResult) -> int:
        """HistorySave の結果を Lot情報履歴 に反映する(行を消してから足す)。"""
        self.env.messages.extend(h.messages)
        sh = self.store.sheet(LOT_HISTORY)
        for r in sorted(h.delete_rows, reverse=True):
            sh.delete_row(r)
        n = 0
        for row in h.new_rows:
            sh.append(row[:26])
            n += 1
        if h.cache is not None:
            c = self.state.cache
            c.lot_no, c.cours, c.ken_no = h.cache["lot_no"], h.cache["cours"], h.cache["ken_no"]
        return n

    # ------------------------------------------------------------------
    # 作業選択・山数・切替
    # ------------------------------------------------------------------
    def select_work(self, work: str) -> Result:
        return self._run(self._select_work, work)

    def _select_work(self, work: str) -> None:
        self._require_line()
        if work not in grid.WORKS:
            raise Stop("作業選択してください", level=WARN, source="UF1:4615")
        layout.select_work(self.state, work, self.hooks, self.env.messages)

    def set_yama(self, yama: str) -> Result:
        return self._run(self._set_yama, yama)

    def _set_yama(self, yama: str) -> None:
        self.state.yama = yama
        layout.yama_change(self.state, self.hooks, self.env.messages)

    def toggle_at(self) -> Result:
        return self._run(lambda: layout.at_butt_click(self.state, self.hooks))

    def toggle_inspe(self) -> Result:
        return self._run(lambda: layout.inspe_butt_click(self.state, self.hooks))

    # ------------------------------------------------------------------
    # 測定値の受信・ラベル
    # ------------------------------------------------------------------
    def receive(self, text: str) -> Result:
        return self._run(self._receive, text, posting=False)

    def _receive(self, text: str) -> None:
        """StartTB → INP_Change。D11: 校正の要否をこの入力の時点で判定して ProofreadFlag に入れる。"""
        self._require_line()
        st = self.state
        st.work_button_tag = "jump" if self.calibration_skip() else ""
        st.proofread_flag = "ok" if self._calibrated() else ""
        # 入力先が決まらず抜けたとき FlagLabel=0 が残るのは VBA どおり(click("") = UserForm_Click で戻る)
        self.env.messages.extend(input_router.receive(st, text, self.hooks))
        if self.env.proofread is not None:
            self.calibration_required = True

    def click(self, name: str) -> Result:
        return self._run(self._click, name)

    def _click(self, name: str) -> None:
        st = self.state
        if name in ("", "UserForm"):
            self.hooks.user_form_click(st)                                # UserForm_Click
            return
        if not st.has_ctl(name):
            raise Stop(f"UF1 にコントロール {name!r} がありません", source="LabelCls")
        input_router.text_tag(st, name, self.hooks)

    def dblclick(self, name: str, answers: Optional[dict[str, bool]] = None, reason: str = "") -> Result:
        return self._run(self._dblclick, name, reason)

    def _dblclick(self, name: str, reason: str) -> None:
        """LabelCls の ObjX_Click → ObjX_DblClick。値があれば消去理由(ConfiReason_Open)を出す。"""
        st = self.state
        if not st.has_ctl(name):
            raise Stop(f"UF1 にコントロール {name!r} がありません", source="LabelCls")
        input_router.text_tag(st, name, self.hooks)
        res = input_router.cell_dbl_click(st, name, self.hooks)
        if res.reason and reason != "":
            self.env.reason = None
            self._confi_reason([reason], "")

    # ------------------------------------------------------------------
    # 手入力欄
    # ------------------------------------------------------------------
    def set_value(self, name: str, value: str) -> Result:
        return self._run(self._set_value, name, value)

    def _set_value(self, name: str, value: str) -> None:
        st, env, hooks = self.state, self._env(), self.hooks
        assert hooks is not None
        value = "" if value is None else str(value)
        if name.startswith("UFConfi."):
            self._confi_input(name[len("UFConfi."):], value)
            return
        if name == "Manual":                                             # CommandButton3(手入力)
            env.messages.extend(input_router.manual_entry(st, value, hooks))
            return
        if name == "Abno":                                               # Abno_Click
            st.v("Abno").value = "異常値off" if st.v("Abno").value == "異常値on" else "異常値on"
            return
        if len(name) == 2 and name[0] in "ZG" and name[1] in grid.POSITIONS:
            self._strain_value(name, value)
            return
        if name in st.strain_cb:
            st.strain_cb[name].value = value
            if env.write_logs:
                for e in legacy_log.log_out_combo_entries(name, value, st.line_name, hooks.worker_name(st),
                                                          aorb_caption=st.strain_face):
                    env.log(e)
            return
        if (name.startswith("Lateral") and name[7:] in grid.POSITIONS) or (
                name.startswith("Diag") and name.endswith("Kou") and name[4:5] in grid.POSITIONS):
            self._lateral_or_angle(name, value)
            return
        if len(name) == 3 and name.endswith("BH") and name[0] in st.bh:   # ?BH_Change → AfterUpdate
            v = value if vbx.is_numeric(value) else ""
            st.bh[name[0]].value = vbx.fmt(v, "0") if v != "" else ""
            bh = quality.bh_search(st.zai, st.tyu)
            wiring.apply_judge(st, judge.simply_judge(wiring.to_judge_form(st), name,
                                                      judge.bh_judg(st.bh[name[0]].value, bh),
                                                      hooks.worker_name(st)), env)
            return
        if len(name) == 4 and name.endswith("Bad") and name[0] in st.bad:
            st.bad[name[0]].value = value if vbx.is_numeric(value) else ""
            input_router.bad_mouse_down(st, name[0], hooks)
            return
        if name.startswith("KenHBox") and name[7:].isdigit() and int(name[7:]) in st.ken_h:
            i = int(name[7:])
            if i == 7:
                v = value if vbx.is_numeric(value) else ""
                st.ken_h[7].value = vbx.fmt(v, "0") if v != "" else ""
            else:
                st.ken_h[i].value = vbx.fmt(value, "0.0")
                hooks.calculate_sheets(st, vbx.val(st.atu))
            return
        if name == "KenCount":                                           # KenCount_AfterUpdate
            v = value if vbx.is_numeric(value) else ""
            st.ken_count = vbx.fmt(v, "0") if v != "" else ""
            th = 0
            if st.atu != "":
                if st.ken_h[7].value == "":
                    th = vbx.clng(vbx.fmt(vbx.vround(hooks.calculate_sheets(st, vbx.single(vbx.val(st.atu))), 1), "0"))
                elif st.ken_count != "":
                    th = vbx.clng(st.ken_count)
                if th >= 5:
                    th = 5
            if st.work_caption.get("厚板") == "H" and st.is_work("少枚数"):
                hooks.small_number_count(st, th)
            return
        if name.endswith(("_UpVisual", "_UnderVisual")) and name[0] in grid.POSITIONS:
            st.v(name).value = value                                     # 外観(LC ボタンの選択結果)
            hooks.check_labels_by_category(st, name[0])
            return
        if name[:3] in ("OPS", "OPE") and name[3:].isdigit() and 1 <= int(name[3:]) <= 6:
            minmax.select_range(st, "S" if name[:3] == "OPS" else "E", int(name[3:]))   # OPS?_Change / OPE?_Change
            return
        if name == "AorB":                                               # AorB_Click(A/B 面の切替)
            if st.strain_face == "Now_A":
                st.strain_face = "Now_B"
                env.messages.append(Message("B面歪 入力画面です", INFO, source="UF1:487"))
            elif st.strain_face == "Now_B":
                st.strain_face = "Now_A"
            r = records.reflect_strain(self.store, wiring.to_snapshot(st, self.store), vbx.right(st.strain_face, 1))
            if r is not None:
                wiring.apply_restore(st, r, hooks)
            st.form_tag = "STOP"
            for c in list(st.strain_z.values()) + list(st.strain_g.values()):   # SetSpecificTextBoxesBackColorToWhite
                c.color = grid.COLOR_WHITE
            hooks.hiz_loop_judg(st, 0)
            st.form_tag = ""
            return
        if name == "DBflag":
            st.dbflag = value
            return
        raise Stop(f"手入力できない欄です: {name}", source="session.set_value")

    def _strain_value(self, name: str, value: str) -> None:
        """Z?/G? の Enter → Change → BeforeUpdate → AfterUpdate(UF1:2532 / 4901 / 5084 / 5047)。"""
        st, env, hooks = self.state, self._env(), self.hooks
        p = name[1]
        input_router.tag_set(st, name)                                   # ?_Enter: TagSet
        if name[0] == "Z" and st.strain_cb[f"CB{p}1"].value in ("耳", "中"):
            text = value
        else:
            text = value if vbx.is_numeric(value) else ""
        st.ctl(name).value = text
        if vbx.is_numeric(st.ctl(name).value):
            st.ctl(name).value = vbx.fmt(st.ctl(name).value, "0.00")
        records.reflect_strain(self.store, wiring.to_snapshot(st, self.store), vbx.right(st.strain_face, 1), "保存")
        wiring.apply_judge(st, judge.hizumi_judg(wiring.to_judge_form(st), name), env)
        if env.write_logs:
            for e in legacy_log.log_out_entries(name, st.ctl(name).value, st.line_name, hooks.worker_name(st),
                                                aorb="AorB", aorb_caption=st.strain_face):
                env.log(e)
        hooks.out_tolerance(st)
        hooks.check_labels_by_category(st, p)

    def _lateral_or_angle(self, name: str, value: str) -> None:
        """Lateral?/Diag?Kou の Enter → Change → BeforeUpdate → AfterUpdate(UF1:991 / 771 / 1287 / 1098)。"""
        st, env, hooks = self.state, self._env(), self.hooks
        lateral = name.startswith("Lateral")
        p = name[7] if lateral else name[4]
        input_router.tag_set(st, name)
        text = value if (vbx.is_numeric(value) or value == "-") else ""
        st.ctl(name).value = text
        env.log(legacy_log.LogEntry("*UFKen:" + name + "_BeforeUpdate:" + text + ":" + st.lname))
        if vbx.is_numeric(text) and text != "-":
            conv = judge.unit_record_input(text)
            target = name
            if not lateral and p == "D" and quirks_reproduce_diag_d():
                target = "DiagSKou"                                       # UF1:1195-1196(書く先が S)
                st.ctl(target).value = vbx.fmt(st.ctl(name).value, "0.0")
            else:
                st.ctl(target).value = conv
            form = wiring.to_judge_form(st)
            out = (judge.lateralbow_judg_color(form, name, hooks.worker_name(st)) if lateral
                   else judge.right_angle_judg_color(form, name, hooks.worker_name(st)))
            wiring.apply_judge(st, out, env)
            hooks.posting(st)
        if st.ctl(name).value == "":
            st.ctl(name).color = grid.COLOR_WHITE
        hooks.out_tolerance(st)
        if not lateral:
            hooks.check_labels_by_category(st, p)

    # ------------------------------------------------------------------
    # UFConfi(公差外の確認・消去理由)
    # ------------------------------------------------------------------
    def _confi_input(self, part: str, value: str) -> None:
        if self.confi is None:
            raise Stop("公差外の確認は出ていません", source="UFConfi")
        if part == "その他":
            self.confi["other"] = value
            return
        if part == "TB1":
            self._confi_send(value)
        elif part == "List2":
            self._confi_send(self.confi.get("tb1", "再測定"), reasons=[r.strip() for r in value.split(",") if r.strip()])
        elif part == "List1":
            self._confi_reason([r.strip() for r in value.split(",") if r.strip()], self.confi.get("other", ""))
        else:
            raise Stop(f"UFConfi に {part} はありません", source="UFConfi")

    def _confi_send(self, tb1: str, reasons: Optional[list[str]] = None) -> None:
        """UFConfi.CommandButton1_Click(送信)。"""
        st, env, hooks = self.state, self._env(), self.hooks
        c = self.confi
        assert c is not None
        posi, knum, temp, k = c["posi"], c["n"], c["moto"], c["ken_name"]
        if c.get("list2"):
            if not reasons:
                env.messages.append(Message("選択してください", WARN, source="UFConfi:200"))
                return
            rows = ", ".join(reasons)
            other = c.get("other", "")
            if not vbx.like(rows, "*" + other + "*"):
                rows = rows + ", " + other
            self._confi_action(tb1, posi, knum)
            env.log(legacy_log.line_entry(st.line_name, "公差外:" + posi + str(knum) + "_" + temp + ":" + "理由選択:"
                                          + "再測定," + rows + ":" + k))
            return
        c["tb1"] = tb1
        d = checks.confi_decide(tb1, str(c["del_count"]))
        if d.action == "need_reason":
            c["list2"] = True
            c["choices"] = list(d.reasons)
            env.messages.append(d.message)
            return
        if d.action == "error":
            env.messages.append(d.message)
        else:
            self._confi_action(tb1, posi, knum)
        env.log(legacy_log.line_entry(st.line_name, "*UFConfi.TB1入力:" + tb1))

    def _confi_action(self, tb1: str, posi: str, knum: int) -> None:
        st, hooks = self.state, self.hooks
        if tb1 == "公差外":
            self._ng_posting(posi)
            self._ng_erase(posi)
        elif tb1 in ("許容", "保留"):
            for k, v in checks.ng_pass(posi, tb1).items():
                wiring.put(st, k, v)
        elif tb1 == "再測定":
            hooks.apply_erase(st, checks.ng_back(wiring.to_uf1(st), posi, knum))
        self.confi = None

    def _ng_posting(self, posi: str) -> None:
        """OutTolerance.NG_Posting → 異常ロット用_<L>。IzyouFl を赤に。"""
        st = self.state
        name = line_sheet(records.ABNORMAL, st.lname)
        old = None
        cells = wiring.sheet_cells(self.store, name)
        if cells is not None:
            last = max([r for (r, c) in cells], default=1)
            rows = [checks.NgRow([records.cell_text(cells.get((r, c))) for c in range(1, 20)])
                    for r in range(2, last + 1)]
            old = checks.AbnormalSheet(records.cell_text(cells.get((1, 1))), rows)
        sheet = checks.ng_posting(wiring.to_uf1(st), posi, old)
        sh = self.store.add_sheet(name)
        sh.write(1, 1, sheet.lot)
        for i, row in enumerate(sheet.rows, start=2):
            for c, v in enumerate(row.values, start=1):
                if v != "":
                    sh.write(i, c, v)
            for c in row.ng_cols:
                sh.set_color(i, c, checks.NG_COLOR)
        st.v("IzyouFl").color = "255"

    def _ng_erase(self, posi: str, all_: str = "") -> None:
        """OutTolerance.NG_Erase。作業保存の行を消し、Tag を消し、作業保存を読み直す。"""
        st = self.state
        act = checks.ng_erase(posi, all_)
        name = line_sheet(records.WORK_SAVE, st.lname)
        if self.store.exists(name):
            sh = self.store.sheet(name)
            for r1, c1, r2, c2 in act.ranges:
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        sh.set(r, c, None)
        if act.split_z_tag:
            p = act.split_z_tag
            for n in (f"CB{p}1", f"CB{p}2", f"CB{p}3", f"CB{p}4", "G" + p, "Z" + p):
                st.ctl(n).tag = ""
        st.load = "off"
        self.hooks.load_text(st, "")
        st.load = "on"

    def _confi_reason(self, reasons: list[str], other: str) -> None:
        """UFConfi.CommandButton2_Click(理由選択の決定)。"""
        st, env = self.state, self._env()
        c = self.confi or {}
        if not reasons:
            env.messages.append(Message("選択してください", WARN, source="UFConfi:31"))
            return
        rows = ", ".join(reasons)
        if vbx.instr(rows, "再測定") > 0 or vbx.instr(rows, "その他") > 0:
            if other == "":
                env.messages.append(Message("理由を入力してください", WARN, source="UFConfi:35"))
                return
            rows = rows + ", " + other
        if not vbx.like(rows, "*" + other + "*"):
            rows = rows + ", " + other
        posi, knum, temp = c.get("posi", ""), c.get("n", 0), c.get("moto", "")
        if vbx.instr(rows, "板廃棄") > 0:
            if knum != 0 and posi in grid.POSITIONS:
                st.ctl(f"{posi}{knum}").value = temp
                rp = "SUCDE".find(posi) + 1
                if rp > 0:
                    wiring.apply_judge(st, judge.loop_judg(wiring.to_judge_form(st), 0, rp, rp,
                                                           bh_lower=quality.bh_search(st.zai, st.tyu)), env)
            self._ng_posting(posi)
            self._ng_erase(posi, "板廃棄")
        env.log(legacy_log.line_entry(c.get("line_name", st.line_name),
                                      posi + str(knum) + "_" + temp + ":" + "理由選択:" + rows + ":"
                                      + c.get("ken_name", "")))
        self.confi = None

    # ------------------------------------------------------------------
    # Min/Max
    # ------------------------------------------------------------------
    def min_max(self) -> Result:
        return self._run(lambda: self.env.messages.extend(minmax.run_min_max(self.state, self.hooks)))

    # ------------------------------------------------------------------
    # 印刷
    # ------------------------------------------------------------------
    def print_check(self, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._print_check, answers or {})

    def _print_check(self, answers: dict[str, bool]) -> Optional[Question]:
        """PrintB_Click の前半 (UF1:3619-3730)。結果を view の `checks` と `print_ready` に入れる。"""
        self._require_line()
        st, env, hooks = self.state, self._env(), self.hooks
        assert hooks is not None
        self.print_ready = False
        self.report = {}
        if st.load == "off":
            return None
        lname = st.lname
        if report_model.is_printed(self.store, st.lot):                 # PrintOnly
            res = report_model.print_only(self.store, lname, st.lot, answers, env.now())
            self.checks = [{"label": "印刷済み(再印刷)", "ok": None}]
            if isinstance(res, Question):
                env.messages.extend(res.messages)
                return res
            if res is not None:
                env.messages.extend(res.messages)
                self.report = self._report_files(res.report, res.defect, reprint=True)
            return None
        st.v("EndTime").value = records.end_time_register(self.store, lname, env.now())
        items: list[dict[str, Any]] = []
        if judge.advance_check(st.lot, st.dbflag, arrbox=hooks._arrbox_fn(st)):
            ans = answers.get("advance_sheet")
            if ans is None:
                env.sound("ありますか.wav", "UF1:3647")
                return Question("advance_sheet", "先行データ票はありますか？", source="UF1:3648")
            items.append({"label": "先行データ票", "ok": bool(ans)})
            if not ans:
                self.checks = items
                return None
        else:
            items.append({"label": "先行データ票", "ok": None})
        uf = wiring.to_uf1(st, wiring.worker_day_max(self.store, lname))
        dist = wiring.work_dist(self.store, lname)
        pre = checks.print_precheck(uf, dist)
        for cb in pre.cleared:                                          # Distortion_Check が形状を空にする
            st.strain_cb[cb].value = ""
        z = checks.distortion_check(uf, "Z")
        g = checks.distortion_check(uf, "G")
        rate = checks.lateralbow_check(uf)
        sumi = uf.t("Sumisei") != '"'
        items += [
            {"label": "定盤", "ok": z.text == ""},
            {"label": "ゲージ", "ok": g.text == ""},
            {"label": "歪 A/B 面", "ok": checks.withor_without_gauge(uf, dist) == ""},
            {"label": "ラテラルボー", "ok": not (rate != "" and sumi
                                                and vbx.instr(uf.t("Additional"), "ﾗﾃﾗﾙﾎﾞｰ検査不要") == 0)},
            {"label": "直角度", "ok": not (checks.angle_check(uf) != "" and uf.t("Youto") == "R115")},
            {"label": "外観", "ok": checks.visual_check(uf) == ""},
            {"label": "検入高さ・枚数", "ok": checks.h_check([uf.t(f"KenHBox{i}") for i in range(1, 8)]) == ""},
            {"label": "山数", "ok": not (st.yama_visible and st.yama == "")},
        ]
        if pre.text != "":
            wiring._msg_error(env, st, pre.text, 1)
            self.checks = items + [{"label": "検査員の照合", "ok": None}, {"label": "梱包枚数", "ok": None}]
            return None
        env.sound("おまちください.wav", "UF1:3717")
        # Name_Match
        wt = line_sheet(records.WORKER_TIME, lname)
        nm = workers.name_match(wiring.sheet_cells(self.store, wt) or {}, wiring.registry(st))
        sh = self.store.ensure_sheet(wt)
        for (r, c, _v) in list(sh.used()):
            if nm.clear_columns[0] <= c <= nm.clear_columns[1]:
                sh.set(r, c, None)
        for w in nm.writes:
            sh.write(w.row, w.col, w.value)
        items.append({"label": "検査員の照合", "ok": not nm.out})
        # Number_Count
        nc = records.number_count(self.store, lname, st.lot)
        stop = nc is not None and nc.stop
        if nc is not None:
            env.messages.extend(nc.messages)
        items.append({"label": "梱包枚数", "ok": None if nc is None else not nc.stop})
        self.checks = items
        self.print_ready = not stop
        return None

    def print_commit(self, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._print_commit, answers or {})

    def _print_commit(self, answers: dict[str, bool]) -> Optional[Question]:
        """PrintB_Click の後半 (UF1:3731-3845)。印刷転送 → 印刷する？ → 保存用 → 測定DB → 履歴(PDF) → CSV → 印刷履歴。"""
        self._require_line()
        st, env = self.state, self._env()
        if not self.print_ready:
            q = self._print_check(answers)
            if q is not None or not self.print_ready:
                return q
        lname = st.lname
        snap = wiring.to_snapshot(st, self.store)
        shifts = env.shift_times()
        penv = report_model.PrintEnv(now=env.now(), lateral_check=checks.lateralbow_check(wiring.to_uf1(st)),
                                     machine_transfer=self._machine_transfer, start1=shifts.start1, end3=shifts.end3)
        with trace.timed("db.save", "帳票を作る(印刷転送)", expected=(Stop,), path=str(self.store.path),
                         table=line_sheet(report_model.PRINT_SHEET, lname)) as info:
            res = report_model.print_transfer(self.store, snap, penv)
            info["cells"] = len(res.report.cells) if res.report else 0
        st.v("EndTime").value = res.end_time or st.v("EndTime").value
        env.messages.extend(res.messages)
        if res.stopped:
            env.messages.append(Message("ｷｬﾝｾﾙ処理", INFO, source="UF1:3740"))
            self.print_ready = False
            return None
        ans = answers.get("print")
        if ans is None:
            env.sound("印刷.wav", "UF1:3763")
            return Question("print", "印刷する？", source="UF1:3764")
        if not ans:
            env.messages.append(Message("ｷｬﾝｾﾙ処理", INFO, source="UF1:3840"))
            return None
        reg = wiring.registry(st)
        self._log("*UFKen:" + "印刷" + " T1:" + reg.t1 + " T2:" + reg.t2 + " T3:" + reg.t3)
        self.report = self._report_files(res.report, res.defect)          # 印刷(画面で開く HTML)
        st.v("dataLoding").value = ""
        snap = wiring.to_snapshot(st, self.store)
        db = self.measure_db or paths.to_local(paths.db_file("梱包検査ﾃﾞｰﾀ"))
        trace.note_used(db, "save")
        trace.note_used(self.store.path, "save")
        with trace.timed("db.save", "保存用への追記と測定DBへの登録(MainDataSave)", expected=(Stop,),
                         path=str(db), table="耳付検査ﾃﾞｰﾀ" if snap.kind == "極厚" else "検査ﾃﾞｰﾀ") as info:
            records.ensure_measure_db(db)
            saved = records.main_data_save(self.store, snap, db_path=db)
            info["rows"] = saved.db.rowcount if saved.db else 0
            info["mode"] = saved.db.mode if saved.db else ""
        if saved.db is not None:
            env.messages.extend(saved.db.messages)
        self.report["saved"] = {"sheet": saved.sheet, "row": saved.row, "table": saved.db_table,
                                "mode": saved.db.mode if saved.db else ""}
        self.report["pdf"] = self._history_pdf(res.report, res.defect)    # 履歴保存
        with trace.timed("db.save", "履歴 CSV を書き出す(ExpCSV)", expected=(Stop,),
                         path=paths.get("history"), table=MIMI if snap.kind == "極厚" else HOZON) as info:
            csvs = records.export_history_csv(self.store, lname, st.line_name, snap.kind == "極厚",
                                              local_dir=self.store.path.parent)
            for p in csvs:
                trace.note_used(p, "save")
            info["files"] = len(csvs)
        self.report["csv"] = [str(p) for p in csvs]
        report_model.set_printed(self.store, st.lot, lname, env.now())
        env.sound("完了.wav", "UF1:3832")
        self.print_ready = False
        return None

    def _machine_transfer(self, sh: Any) -> None:
        """MachineTransfer.機器転送(sh, 9, 5, ATU, Hab, Youto)(Print:541-544)。"""
        from . import instruments
        st = self.state
        ln = lines.find(st.line_name)
        t = instruments.TransferInput(9, 5, st.atu, st.hab, st.youto, line=ln.official if ln else "", tak=st.tak,
                                      hab=st.hab, cells=sh.get)
        for r, c, v in instruments.machine_transfer(t):
            sh.put(r, c, v)

    def _report_files(self, report: Optional[report_model.ReportModel], defect: Optional[report_model.ReportModel],
                      reprint: bool = False) -> dict[str, Any]:
        """印刷(D8: ブラウザで開く HTML)。手元のデータフォルダの print に書く。"""
        folder = self.store.path.parent / "print"
        out: dict[str, Any] = {"reprint": reprint}
        edge = None
        for key, model, stem in (("report", report, f"印刷用_{self.state.lname}"),
                                 ("defect", defect, f"不良要因印刷_{self.state.lname}")):
            if model is None:
                continue
            with trace.timed("db.save", "帳票の HTML", path=str(folder), table=stem) as info:
                files = report_html.save_html_and_pdf(model, folder, stem, edge)
                info["files"] = len(files)
            out[key] = files
        return out

    def _history_pdf(self, report: Optional[report_model.ReportModel],
                     defect: Optional[report_model.ReportModel]) -> list[dict[str, str]]:
        """Input.履歴保存 (Input:1014-1062)。手元(C:\\検査履歴)と工場共有へ成績票・不良要因票の PDF(D7)。"""
        st, env = self.state, self._env()
        edge = paths.to_local(paths.get("edge"))
        edge_s = str(edge) if edge.is_file() else None
        out = []
        for folder in wiring.pdf_folders(st.line_name, env.now()):
            for model, stem in ((report, records.cell_text(report.cell("C2")["value"]) if report and report.cell("C2")
                                 else st.lot),
                                (defect, (records.cell_text(defect.cell("B1")["value"]) if defect and defect.cell("B1")
                                          else st.lot) + "_不良要因")):
                if model is None:
                    continue
                local = wiring.legacy_log.local_path(folder)
                trace.note_used(local, "save")
                with trace.timed("db.save", "帳票の PDF(履歴保存)", path=str(local), table=stem) as info:
                    try:
                        files = report_html.save_html_and_pdf(model, local, stem, edge_s)
                    except (OSError, report_html.PdfError) as e:
                        env.messages.append(Message("PDFファイルが作成されませんでした: " + str(local / (stem + ".pdf"))
                                                    + "\r\n" + str(e), WARN, source="補助:739"))
                        files = {}
                    info["files"] = len(files)
                out.append(files)
        return out

    # ------------------------------------------------------------------
    # 新規検査
    # ------------------------------------------------------------------
    def new_inspection(self, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._new_inspection, answers or {})

    def _new_inspection(self, answers: dict[str, bool]) -> Optional[Question]:
        """NewKen_Click (UF1:3359-3533)。Lot が空なら確認なしで消去へ(GoTo line)。"""
        self._require_line()
        st, env, hooks = self.state, self._env(), self.hooks
        assert hooks is not None
        lname = st.lname
        if st.lot != "":
            ans = answers.get("new")
            if ans is None:
                return Question("new", "新規検査作業を開始しますか", title="確認", source="UF1:3373")
            if not ans:
                return None
            found = False
            for sheet in (HOZON, MIMI):
                rows = wiring.sheet_rows(self.store, sheet) or []
                if any(records.cell_text(r[1]) == st.lot for r in rows[1:] if len(r) > 1):
                    found = True
                    break
            if not found:
                wiring._msg_error(env, st, "該当LotNoが保存されていません", 2, "END")
        # line:
        self._log("UFKen" + "*新規検査実行")
        self.store.add_sheet(line_sheet(records.WORK_SAVE, lname))
        self.store.add_sheet(line_sheet(records.WORKER_TIME, lname))
        for name in (line_sheet(report_model.PRINT_SHEET, lname), line_sheet(records.ABNORMAL, lname),
                     line_sheet(report_model.DEFECT_SHEET, lname)):
            if self.store.exists(name):
                self.store.delete_sheet(name)
        now = env.now()
        if workers.register_check(workers.hhmmss(now), wiring.registry(st), wiring.worker_sheet_a(self.store, lname),
                                  now=now, facility=env.facility() or workers.DEFAULT_FACILITY,
                                  shifts=env.shift_times(), n="新規"):
            self.store.add_sheet(line_sheet("検査員", lname))
        layout.clear_lselect_cache(st)
        for p in grid.POSITIONS:
            st.v("LZ" + p).value = "定盤:"
            st.v("LG" + p).value = "ｹﾞｰｼﾞ"
        st.amin = st.amax = st.hmin = st.hmax = st.tmin = st.tmax = ""
        st.flatnes_text = st.conductiv_text = st.comment = st.reien = st.ken_no_caption = st.additional = ""
        for n in ("flatnes", "QF", "LastHot", "dataLoding", "KenH", "HizCome"):
            st.v(n).value = ""
        st.v("HizCome").tag = ""
        st.yama_tag = ""
        st.lateral_text_colored = False
        st.v("IzyouFl").color = grid.COLOR_BUTTON
        for p in grid.POSITIONS[:5]:
            st.v("TS" + p).visible = True
            st.v("TE" + p).visible = True
        for p in grid.POSITIONS:                                        # Split_Z_Tag_0
            for n in (f"CB{p}1", f"CB{p}2", f"CB{p}3", f"CB{p}4", "G" + p, "Z" + p):
                st.ctl(n).tag = ""
        if st.strain_face == "Now_B":
            st.strain_face = "Now_A"
        ws = self.store.sheet(line_sheet(records.WORK_SAVE, lname))
        wiring.apply_restore(st, records.losd_strain("A", ws.read_range(2, 15, 7, 16), ws.read_range(8, 15, 13, 16)),
                             hooks)
        hooks.load_text(st, "ReLoad")
        hooks.worker_load(st)
        layout.at_count_select(st, 4, hooks)
        layout.inspection_select(st, "表示", hooks)
        layout.right_angle_display(st, "新規")
        layout.lselect_lateral(st, "新規")
        st.load = "on"
        self.confi, self.checks, self.print_ready, self.report = None, [], False, {}
        return None


def quirks_reproduce_diag_d() -> bool:
    """UF1:1195-1196 DiagDKou_BeforeUpdate が変換結果を DiagSKou に書く(B 一覧に無い。judge の報告で発見)。

    B 番号が無いので quirks のスイッチは使わず、VBA どおり常に再現する。
    """
    return True


_session: Optional[Session] = None
_slock = threading.Lock()


def get() -> Session:
    global _session
    with _slock:
        if _session is None:
            _session = Session()
        return _session


def reset_for_test(s: Optional[Session] = None) -> Session:
    global _session
    with _slock:
        _session = s or Session()
        return _session
