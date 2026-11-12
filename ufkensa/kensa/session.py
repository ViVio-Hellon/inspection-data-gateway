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
    set_lot_source(on, lot)       (新規 D28)ロットごとの「控え(第2)から読む」。VBA の DBflag の運用の代わり

D28: 「控え(第2)から読む」を選んだロットを読むあいだ(`_reads_second`)、操作はすべて
`paths.lot_source(...)` の中で動かす。仕掛台帳の読み(Arrbox・AccReflec・HistorySave・SikaAtuCreate・
LSelect / Atu_OPselect の引当・引当の読み・オーダーNo など)は paths を通るので、モードによらず第2だけになる。

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
from ..logging_utils import get_logger
from . import (calibration, checks, datasource, grid, input_router, judge, layout, legacy_log, lines, lot as lotmod, minmax, packing,
               paths, quality, records, report_html, report_model, terminal, view, vbx, workers, wiring)
from . import lot_watch, nippou_export, shadow, size_thresholds, worker_archive
from . import reasons as reason_master
from .model import CellKey, Face, FormTag, Switch, bad_name, ken_h_name
from . import sound as snd
from .flow import ERROR, INFO, WARN, Message, Question, Stop, unexpected_stop
from .grid import KensaState
from .store import HOZON, LOT_HISTORY, MIMI, LocalStore, line_sheet
from .wiring import Env, RealHooks

#: 設備名(レジストリ)→ UF1 のライン OptionButton(UF1:230-238)
FACILITY_OPTION = {"L-1": "L1Work", "LVC": "LVC", "HVC": "HVC", "バランサー": "Bala", "トット": "Tott",
                   "Test": "Test", "丸徳": "Maru"}
log = get_logger(__name__)
NO_LINE_TEXT = "ライン設定がありません" + "\r\n" + "担当者に連絡して下さい"          # UF1:225
CONFI_CHOICES = ("公差外", "再測定", "許容", "保留")
#: 理由の選択肢は 伝送用ファイル の表「理由選択」(reasons.py。無ければ VBA と同じ既定で自動で作る)
CONFI_REASONS = reason_master.DEFAULTS[reason_master.KIND_DELETE]


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
        #: 2026-11-09: 最後に受けた測定値が入らなかった理由(入ったら "")。画面の受信状態に出す(view には入れない)
        self.receive_why = ""
        self.visual_lists: checks.VisualLists = checks.VisualLists()   # Visual_Add(外観・山・不良要因の選択肢)
        self.orders: list[str] = []                                   # OrdNoCB の一覧(OrdNoOpen)
        self.order: str = ""                                          # OrdNoCB.Text
        self.packing_info = {"no": "", "form": "", "comment": ""}     # PackingNo / PackingForm / PackingComment
        self.pack: Optional[packing.PackForm] = None                  # 出している UF_PackNo(包装仕様書の確認)
        self.name_list: Optional[int] = None                          # 開くべき UFNameList(1 不一致 / 3 梱包枚数)

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

    def _prune_workers(self) -> None:
        """検査を持ち越していない(ロットが入っていない)なら、いまの直の検査員だけ残す。"""
        st, env = self.state, self._env()
        if st.lot != "":
            return
        try:
            reg = workers.current_shift_only(wiring.registry(st), now=env.now(),
                                             facility=env.facility() or workers.DEFAULT_FACILITY,
                                             shifts=env.shift_times())
        except workers.ShiftTimesMissing:
            return
        wiring.set_registry(st, reg)

    def _extra(self) -> dict[str, Any]:
        st = self.state
        reg = wiring.registry(st)
        workers_list = []
        for t, day, label in ((reg.t1, reg.day1, "T1"), (reg.t2, reg.day2, "T2"), (reg.t3, reg.day3, "T3")):
            if t != "" or day != "":
                workers_list.append(f"{label}: {t}" + (f"({day})" if day else ""))
        return {**self._extra_more(), "lot_source": self._lot_source_view(), "workers": workers_list,
                "worker_time": self._worker_time(), "checks": self.checks, "print_ready": self.print_ready,
                "confi": self._confi_view(), "calibration_required": self.calibration_required and not self._calibrated(),
                "report": self.report or None, "proofread": "ok" if self._calibrated() else "",
                "range": {"ops": st.ops, "ope": st.ope,
                          "labels": [st.v(f"OPS{i}").value if st.v(f"OPS{i}").visible else "" for i in range(1, 7)]}}

    def _lot_source_view(self) -> dict[str, Any]:
        """D28: 「控え(第2)から読む」の表示。active = いま出しているロットがそれ(バッジを出す)。"""
        st = self.state
        on = st.lot_source == wiring.LOT_SECOND
        return {"on": on, "lot": st.lot_source_lot if on else "", "active": wiring.reads_second(st, st.lot),
                "mode": paths.MODE_LABELS.get(paths.mode(), "")}

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
        try:
            return not calibration.input_blocked(cells, now=env.now(),
                                                 facility=env.facility() or workers.DEFAULT_FACILITY,
                                                 shifts=env.shift_times(), jump=self.calibration_skip(),
                                                 db_flag=st.dbflag)
        except (workers.ShiftTimesMissing, datasource.DbMissing, datasource.TableMissing):
            return False      # 直の時刻が読めない(画面を描くために校正済みとはしない。理由は操作のときに出る)

    def _reads_second(self) -> bool:
        """D28: いまのロットを控え(第2)だけから読むか(読むたびに状態を見る)。"""
        return wiring.reads_second(self.state, self.state.lot)

    def _run(self, fn: Callable[..., Any], *args: Any, posting: bool = True, keep_ready: bool = False) -> Result:
        with self._lock, paths.lot_source(self._reads_second):
            env = self._env()
            env.messages, env.confi, env.reason, env.proofread = [], [], None, None
            # 処理の途中だけ立てる旗(VBA の Load = "off"・FormTag = "STOP")。途中で止まったら元に戻す。
            # 戻さないと、次からの入力・描画が全部止まったままになる(2026-11-06 の構造の見直し)
            st0 = self.state
            if self.started and (st0.flag_label != 1 or st0.load != Switch.ON or st0.form_tag == FormTag.STOP):
                # 操作は 1 つずつ順に来る(ロック)ので、操作の始めに「処理中」の旗が立っているのは、前の操作が途中で
                # 止まって戻し忘れたもの。VBA なら End で全部初めに戻る。残すと測定値・カードを黙って受け付けなくなる
                log.warning("前の操作の旗が残っていたので戻します(FlagLabel=%r Load=%r FormTag=%r)",
                            st0.flag_label, st0.load, st0.form_tag)
                st0.flag_label, st0.load = 1, "on"
                if st0.form_tag == FormTag.STOP:
                    st0.form_tag = ""
            flags = (self.state.load, self.state.form_tag)
            try:
                try:
                    q = fn(*args)
                except datasource.DbMissing as e:                       # VBA は End(D10: 止めるが状態は保つ)
                    raise Stop(str(e), messages=list(env.messages), source="datasource") from e
                except (Stop, NotImplementedError):
                    raise
                except Exception as e:                                   # noqa: BLE001 - 想定外でも画面は止めない
                    raise unexpected_stop(e, "kensa." + getattr(fn, "__name__", ""), env.messages) from e
            except Stop as e:
                self.state.load, self.state.form_tag = flags
                shadow.check_verdicts(self.state, "kensa." + getattr(fn, "__name__", ""))
                msgs = list(env.messages)
                for m in e.messages:
                    if not any(m is x for x in msgs):
                        msgs.append(m)
                return Result(self._view(), msgs, stop=e)
            if not keep_ready:
                # 印刷前チェックの後に値・ロットが変わったら、チェックはやり直し(印刷確定がもう一度チェックする)
                self.print_ready = False
                self._print_pending = None           # 確認の後に何か変わったら、帳票は作り直す
            if env.confi:
                self._open_confi(env.confi[-1])
            if env.reason is not None:
                self._open_confi(env.reason)
            if posting and isinstance(q, Question) is False and self.started:
                assert self.hooks is not None
                try:
                    with self.store.transaction():          # 作業中保存は全部書けたときだけ確定(途中で落ちたら前のまま)
                        self.hooks.posting(self.state)
                except Exception as e:                       # noqa: BLE001 - 保存できなくても画面は止めない
                    log.exception("作業中保存(Posting)に失敗")
                    env.messages.append(Message(f"作業中の保存ができませんでした({type(e).__name__}: {e})。"
                                                "このまま続けられますが、再起動すると直前の操作が戻りません", WARN,
                                                source="Posting"))
            shadow.check_verdicts(self.state, "kensa." + getattr(fn, "__name__", ""))
            return Result(self._view(), list(env.messages), question=q if isinstance(q, Question) else None)

    def _confi_view(self) -> Optional[dict[str, Any]]:
        """出している確認。理由の選択肢はマスタ(理由選択)をその都度読む(保守で直したらすぐ出る)。"""
        c = self.confi
        if c is None:
            return None
        if c.get("list2"):
            c["choices"] = reason_master.choices(reason_master.KIND_REMEASURE) or c["choices"]
        elif c.get("page") == 1:
            c["choices"] = reason_master.choices(reason_master.KIND_DELETE)
        return c

    def _open_confi(self, req: checks.ConfiRequest) -> None:
        """UFConfi を出す(vbModeless。最後の 1 件が残る)。page 0 = 公差外、1 = 理由選択。"""
        self.confi = {"posi": req.posi, "n": req.n, "line_name": req.line_name, "moto": req.moto,
                      "return_moto": req.return_moto, "del_count": req.del_count, "page": req.page,
                      "ken_name": req.ken_name or "", "list2": False,
                      "choices": list(CONFI_CHOICES) if req.page == 0 else reason_master.choices(reason_master.KIND_DELETE)}
        if req.page == 0:
            self.env.messages.append(Message("検査ﾃﾞｰﾀに公差外有り！", WARN, title="2024.7.26_Ver",
                                             source="UFConfi.Label1"))

    def _require_line(self) -> None:
        if self.state.line_name == "":
            raise Stop(NO_LINE_TEXT, level=WARN, source="UF1:225")

    def current(self) -> Result:
        """いまの画面の値だけ(操作をしない。画面を開いたときの最初の描画)。"""
        with self._lock, paths.lot_source(self._reads_second):
            try:
                return Result(self._view())
            except (datasource.DbMissing, datasource.TableMissing) as e:
                return Result(view.build(self.state, {}), stop=Stop(str(e), source="datasource"))

    # ------------------------------------------------------------------
    # 起動
    # ------------------------------------------------------------------
    def start(self) -> Result:
        return self._run(self._start)

    def _start(self) -> None:
        """UF1.UserForm_Initialize (UF1:173-328) と UserForm_Activate の校正済み表示。

        VBA の Initialize はフォームを開いたとき 1 回だけ。Web では画面を開くたび(再読込・別の画面から戻る・
        2 つ目のタブ)に呼ばれるので、2 回目からは作り直さない(公差外の確認・分割フラグ・2 度押しの記憶などが
        消えないように。2026-11-06 の見直し)。ラインを変えたときだけ作り直す。
        """
        fac = terminal.facility()
        if fac == "":
            raise Stop(NO_LINE_TEXT, level=WARN, source="UF1:224")
        if self.started and getattr(self, "_started_facility", None) == fac:
            self._prune_workers()
            return None
        self._started_facility = fac
        env, hooks = self._env(), self.hooks
        assert hooks is not None
        with trace.timed("kensa.start", "検査画面の起動", expected=(Stop,), facility=fac):
            report_model.clean_old_history(self.store, env.now())       # CleanOldHistory
            st = self.state = KensaState()
            self.confi, self.checks, self.print_ready, self.report = None, [], False, {}
            st.auto, st.form_tag, st.load, st.flag_label, st.inp_tag, st.picfm_tag = "on", "STOP", "off", 1, "", "0"
            st.v("Abno").value = "異常値on"
            st.v("DistortionFrame").value = "歪公差"
            self.orders, self.order, self.pack, self.name_list = [], "", None, None
            self.packing_info = {"no": "", "form": "", "comment": ""}
            env.vc = None
            self._visual_add()                                          # UF1:226 Visual_Add
            opt = FACILITY_OPTION.get(fac, "")
            if opt:
                st.line_option = opt
                maru = 0
                if opt == "Maru" and terminal.line_no() in terminal.MARU_NUMBERS:
                    maru = terminal.MARU_NUMBERS.index(terminal.line_no()) + 1
                layout.line_auto_selection(st, "", hooks, maru)
            # UF1:110 の SikaSetAtu は UF1 の Private 変数で、SaveLoad の Public SikaSetAtu(state.sika_set_atu)とは別。
            # Initialize で書いた値は LSelect / CBAdd 側には届かない(フォームを開いた直後は 0)
            uf1_sika_set_atu = 0.0
            if st.lot != "":
                try:
                    all_ = lotmod.sika_atu_create(st.lot)
                except Exception:                                       # noqa: BLE001 - SikaAtuCreate の End
                    all_ = None
                if all_:
                    uf1_sika_set_atu = vbx.single(vbx.val(all_[0]))
            if uf1_sika_set_atu != 0:                                   # UF1:253-261 歪コンボ
                if st.blindfold_tag != "Set":
                    hooks.cb_add(st, uf1_sika_set_atu)
                else:
                    st.blindfold_tag = ""
            else:
                hooks.cb_add(st, size_thresholds.thin_atu())   # VBA は 6.0 固定(= 薄板の形状)。しきい値ちょうどを渡す
            st.form_tag = ""
            st.load = Switch.ON
            self.started = True
            if st.lot != "" and st.work != "" and st.inp_tag == "" and st.auto == Switch.ON:
                # 開き直した直後も「次に測る」を出す(VBA は InpTag を保存しないので空だった。値は続きの欄に入っていた)
                try:
                    input_router.next_blank(st, input_router.ken_no_for_input(st))
                except Exception:                                       # noqa: BLE001 - 表示だけなので止めない
                    log.warning("次に測る欄を出せませんでした", exc_info=True)
            self._prune_workers()                 # 持ち越しの検査が無ければ、前の直の検査員は残さない
            self._activate_name_check()

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
        self._prune_workers()
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
        if changed and st.load != Switch.OFF:
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
                    env.sound(snd.KENSA_CALIBRATION, "UF1:4116")
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
        prev_lot = st.lot
        tol_before = lot_watch.tolerance_view(st)          # 画面を空にする前の公差(判定し直すかの比べ元)
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
        # 2026-11-06: 桁の確かめと「直近5Lot」の確認は、画面の値を消す前に行う。前は先に消していたので、
        # 8 桁を入れると黙って作業保存のロットが空になり、「直近5Lot」でいいえと答えると新しいロット番号に
        # 前のロットの公差が残ったまま保存されていた
        text = "".join(ch.upper() if "a" <= ch <= "z" else ch for ch in lot
                       if ch.isascii() and ch.isalnum())
        if len(text) != 7:
            env.messages.append(Message(f"LotNo は英数字 7 桁で入れてください(入れた値: {lot or '(空)'})。"
                                        "いまのロットはそのままです", WARN, source="Lot_Change"))
            return None
        if st.load != Switch.OFF and judge.check_for_lot(wiring.sheet_rows(self.store, HOZON) or [], text, 2):
            a3 = answers.get("lot_recent")
            if a3 is None:
                return Question("lot_recent", "直近5Lot内に同じLotNoがあります" + "\r\n" + "このまま作業しますか？",
                                title="確認", source="UF1:4505")
            if not a3:
                return None                          # いいえ: 何も変えない
        # line:(入力欄を空にする)
        st.lot = ""
        st.yama = ""
        st.ken_no_caption = ""
        st.ken_head = ""                     # 2026-11-10: 作った検番の頭も前のロットを残さない
        for n in ("HizCome", "QF", "LastHot", "dataLoding", "DebugTaikaku", "DebugLate"):
            st.v(n).value = ""
        st.v("HizCome").tag = ""
        st.flatnes_text = ""
        st.machin = ""
        split = st.view.get("SplitFlag")
        if text != prev_lot and (split is None or split.value == ""):
            # 別のロット = 持ち越しではない(st.lot は空にしてある)。ロット分割は前半の検査員を残す(VBA も触らない)
            self._prune_workers()
        st.lot = text
        if st.lot_source != "" and st.lot_source_lot != text:            # D28: 別のロットは いつもの読み方に戻す
            self._set_lot_source("", "", f"別のロット({text})を登録したので")
        # Lot_Change
        if st.form_tag != FormTag.STOP:
            if self.write_logs:
                legacy_log.write_log("Lot登録|" + text + "|" + st.lname, "")
            self._log("Lot登録:" + text + ":" + st.lname)
            st.machin = ""
            self._log("Lot登録:" + text + ":" + st.lname, legacy_log.MEASURED_LOG)
        return self._cb_change(answers, tol_before)

    def _cb_change(self, answers: dict[str, bool], before: Optional[dict[str, str]] = None) -> Optional[Question]:
        """CB_Change の後、公差が変わっていれば測定済みの値を新しい公差で判定し直す(2026-11-11 のご依頼)。

        VBA の CB_Change(AccReflec)は公差の欄を書き換えるだけで、入っている測定値の色は前の公差のまま
        (次に LoopJudg が走る操作 = Min/Max・再測定 などまで古い判定が残る)。別ロット登録・ロット分割・
        読み直し・控えから読む のどれでも、ここで LoopJudg → HizLoopJudg → CheckLabelsByCategory をやり直す。
        許容・保留(NG? のラベル)はそのまま。
        """
        if before is None:
            before = lot_watch.tolerance_view(self.state)
        try:
            return self._cb_change_body(answers)
        finally:
            if lot_watch.tolerance_view(self.state) != before:
                self._rejudge_measured()

    def _rejudge_measured(self) -> None:
        """測定値があれば、いまの公差で全部判定し直す(公差外の確認は出さない。許容・保留は変えない)。"""
        st, env, hooks = self.state, self._env(), self.hooks
        assert hooks is not None
        extra = [n for p in grid.POSITIONS for n in ("Z" + p, "G" + p, "Lateral" + p, "Diag" + p + "Kou")]
        if not (any(c.value != "" for _, c in st.iter_cells())
                or any(n in st.view and st.view[n].value != "" for n in extra)):
            return
        try:
            hooks.loop_judg(st)
            hooks.hiz_loop_judg(st, 0)
            for p in grid.POSITIONS:
                hooks.check_labels_by_category(st, p)
        except Stop as e:
            env.messages.append(Message("新しい公差での判定し直しが途中で止まりました: " + e.text, WARN,
                                        source="ロット読み込み"))
            return
        env.messages.append(Message("ロットの公差が変わったので、入っている測定値を新しい公差で判定し直しました",
                                    INFO, source="ロット読み込み"))

    def _cb_change_body(self, answers: dict[str, bool]) -> Optional[Question]:
        """UF1.CB_Change (UF1:4494-4535)。ロットの読み込み(Lot登録と「控えから読む」の切替から呼ぶ)。"""
        st, env, hooks = self.state, self._env(), self.hooks
        assert hooks is not None
        text = st.lot
        if st.load == Switch.OFF:
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
        self._lot_sig = lot_watch.signature(st.lot, st.dbflag)            # 見張りの基準(読み直し・watch で比べる)
        self._lot_sig_told = None
        # CheckAndUpdateButtonColor は view が印刷履歴から毎回決める(印刷 / 再印刷・印刷履歴クリア)
        self._ord_no_open()                                              # UF1:4518
        if not self._packing_open():                                     # UF1:4519-4520
            return None                                                  # 実行時エラーで VBA はここで止まる
        # DistArray → OpForcedSelection → FormReflectDist
        lot_data = quality.LotData(hooks.arrbox(st, "SIKALOTNOW.accdb", "板", "Fields"),
                                   hooks.arrbox(st, "SIKALOTNOW.accdb", st.lot, "Record"))
        qp = None
        try:
            # 決定 D15(W16): VBA はここで検番がまだ空(LSelect が後で入れる)のまま引いていた。
            # 品質基準画面と同じく、ロットの検番で引く。
            from .quality_session import ken_no_of
            ken = st.ken_no_caption or ken_no_of(st.lot, lot_data, st.dbflag)
            qp = quality.dist_array(st.lot, ken, quality.hiki_arr_create(st.lot), lot_data)
            env.messages.extend(qp.messages)
            wiring.apply_values(st, qp.uf1)
        except (vbx.TypeMismatch, quality.VbaRuntimeError, IndexError) as e:
            env.messages.append(Message(str(e), ERROR, source="QP.DistArray"))
        layout.op_forced_selection(st)
        if qp is not None and qp.kousa:
            r = quality.form_reflect_dist(qp.kousa, quality_visible=False, mitsubishi_visible=False)
            wiring.apply_values(st, r.values)
        return None

    # ------------------------------------------------------------------
    # ロット情報の読み直し・見張り(2026-11-07。VBA には無い)
    # ------------------------------------------------------------------
    def reload_lot(self) -> Result:
        return self._run(self._reload_lot)

    def _reload_lot(self) -> None:
        """いまのロットの情報(仕掛・引当・オーダー・品質基準)を読み直す。測定値・検査員はそのまま。

        読み直した後に判定をやり直し、変わったもの(種類と公差の前後)を知らせる。
        """
        self._require_line()
        st, env, hooks = self.state, self._env(), self.hooks
        assert hooks is not None
        if st.lot == "":
            raise Stop("ロットが登録されていません(読み直すロットがありません)", level=WARN, source="読み直し")
        old_sig = getattr(self, "_lot_sig", None)
        before = lot_watch.tolerance_view(st)
        keep = {n: (c.value, c.tag) for n, c in st.iter_cells()}
        with trace.timed("kensa.reload", "ロット情報の読み直し", lot=st.lot):
            self._cb_change({"lot_recent": True})
        # 読み直しで測定値は変えない(AccReflec は公差・規格の欄だけを書く。念のため戻す)
        for n, (v, t) in keep.items():
            c = st.ctl(n)
            if c.value != v:
                c.value, c.tag = v, t
        hooks.loop_judg(st)                                              # 新しい公差で判定し直す
        hooks.hiz_loop_judg(st)
        for p in grid.POSITIONS:
            hooks.check_labels_by_category(st, p)
        new_sig = getattr(self, "_lot_sig", None)
        kinds = new_sig.changed(old_sig) if new_sig is not None and old_sig is not None else []
        diff = lot_watch.tolerance_diff(before, lot_watch.tolerance_view(st))
        if kinds or diff:
            text = "ロット情報を読み直しました。変わったもの: " + ("・".join(kinds) if kinds else "公差・規格")
            if diff:
                text += "\r\n" + "\r\n".join(diff[:20])
            env.messages.append(Message(text, WARN, source="読み直し"))
        else:
            env.messages.append(Message("ロット情報を読み直しました(変わったところはありません)", INFO, source="読み直し"))
        self._log("読み直し:" + st.lot + ":" + ("・".join(kinds) or "-") + ":" + st.lname)

    def lot_watch(self) -> dict[str, Any]:
        """画面から定期に呼ぶ見張り。いまのロットの情報が登録(読み直し)の後に変わっていれば、その種類を返す。"""
        with self._lock:
            st = self.state
            base = getattr(self, "_lot_sig", None)
            if not self.started or st.lot == "" or base is None or base.lot != st.lot:
                return {"lot": st.lot if self.started else "", "changed": [], "new": False}
            now = lot_watch.signature(st.lot, st.dbflag)
            kinds = now.changed(base)
            new = bool(kinds) and kinds != getattr(self, "_lot_sig_told", None)
            if new:
                self._lot_sig_told = kinds
                trace.event("kensa.lot_watch", f"ロット {st.lot} の情報が更新されています: {'・'.join(kinds)}",
                            data={"lot": st.lot, "changed": kinds})
            return {"lot": st.lot, "changed": kinds, "new": new}

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
        if r.status == "not_found":
            # VBA は何も言わずに欄を空にするだけ(SkipNoItti)。画面では「何も起きない」に見えるので知らせる
            where = getattr(r, "db_path", "") or ""
            env.messages.append(Message(
                f"LotNo {st.lot} は仕掛台帳(SIKALOT)にありません。"
                + (f"読んだファイル: {where}。" if where else "")
                + "設定 → この端末 のモード(オンライン=台帳→控え、オフライン=控えだけ)と、使用パスの「仕掛台帳」を確かめてください",
                WARN, source="SaveLoad.AccReflec"))
        if r.status != "ok":
            return
        wiring._tolerance_text(st, r)
        if r.lateral_text is not None:
            try:
                layout.lselect_lateral(st, r.ken_no)
            except layout.VBA_ERRORS as _vba_e:
                layout.vba_swallow(_vba_e, "session.py _acc_reflec")
                pass
        try:                                                             # SaveLoad:378-381
            if r.vc_action == "clear":
                hooks.vc_clear(st)
            elif r.vc_action == "picture":
                hooks.vc_picture(st)
        except (datasource.DbMissing, datasource.TableMissing, vbx.TypeMismatch, IndexError) as e:
            env.messages.append(Message(str(e), ERROR, source="SaveLoad.VC_Picture"))
        if st.atu != "":
            w = quality.weight(vbx.val(st.atu), vbx.val(st.hab), vbx.val(st.tak))
            layout.atu_opselect(st, vbx.val(st.atu), hooks, w, env.messages)   # SaveLoad:384 val(…) は Double
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
        hooks.cb_add(st, st.sika_set_atu)                                # SaveLoad:501 / 515
        rows = hooks.get_lot_info(st, st.lot)
        if rows is None:
            return
        try:
            layout.judge_lot_allocations(st, rows, hooks)
        except layout.VBA_ERRORS as _vba_e:
            layout.vba_swallow(_vba_e, "session.py _acc_reflec")
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

    CONFI_WAIT = "理由を選んでから測ってください(確認の画面で理由を選ぶまで次の測定はできません)"

    def _confi_block(self, what: str = "") -> None:
        """2026-11-10: 確認(UFConfi)を出している間は測定を受け付けない(現場の声: 理由を選ぶまで次へ進めない)。
        VBA は vbModeless だが、選ばずに測り続けられて理由が抜けるので、選ぶまで止める。"""
        if self.confi is not None:
            self.receive_why = self.CONFI_WAIT
            raise Stop(self.CONFI_WAIT, source="UFConfi", level=WARN)

    def _receive(self, text: str) -> None:
        """StartTB → INP_Change。D11: 校正の要否をこの入力の時点で判定して ProofreadFlag に入れる。"""
        self._require_line()
        self._confi_block()
        st = self.state
        st.work_button_tag = "jump" if self.calibration_skip() else ""
        st.proofread_flag = "ok" if self._calibrated() else ""
        # 入力先が決まらず抜けたとき FlagLabel=0 が残るのは VBA どおり(click("") = UserForm_Click で戻る)
        if st.inp_tag == "" and st.lot != "" and st.work != "" and st.load != Switch.OFF:
            # 2026-11-08: 「〜検査済」の後は、VBA ではフォームの余白を押して次の入力先を決めていた(UserForm_Click)。
            # 画面には余白が無く、測定器の値が黙って捨てられていたので、入力先が無ければここで同じことをしてから受ける
            self.hooks.user_form_click(st)
        # 2026-11-09: 「検査できない理由を表示して」(現場の声)。入らなかったとき、その理由を残す。
        # VBA は黙って捨てる道(ロット未登録・入力先が無い・同じ値の続き)があるので、ここで言葉にする(動きは変えない)
        self.receive_why = ""
        pre = ("ロット未登録(LotNo を入れてください)" if st.lot == "" else
               "同じ値が続けて届いたので受け付けませんでした" if st.inp == text and text != "" else "")
        n0 = len(self.env.messages)
        try:
            self.env.messages.extend(input_router.receive(st, text, self.hooks))
        except Stop as e:
            self.receive_why = e.text
            raise
        landed = st.last_landed            # 値が入った欄(入らなければ "")。input_router.inp_change が残す
        if self.env.proofread is not None:
            self.calibration_required = True
        if landed and st.has_ctl(landed) and st.ctl(landed).value == "":
            # 入力先には入ったが、異常値・2度押しなどで消された
            texts = [" ".join(m.text.split()) for m in self.env.messages[n0:] if m.text.strip()]
            try:
                val = f"(値 {input_router.vbx.cstr(input_router.vbx.val(input_router.conv(text)))})"
            except Exception:                                          # noqa: BLE001 - 理由の文言だけ
                val = ""
            self.receive_why = (f"入れた値{val}を消しました: " + texts[0]) if texts else f"入れた値{val}が消されました(異常値・2度押しなど)"
        if not landed:
            warn = [m.text for m in self.env.messages[n0:] if m.level in (WARN, ERROR) or "検査済" in m.text]
            if st.lot == "":
                why = pre
            elif st.work == "":
                why = "作業選択してください(作業が選ばれていません)"
            elif self.env.proofread is not None or (st.proofread_flag != "ok" and st.work_button_tag != "jump"):
                why = "校正未済(校正画面で校正してから測ってください)"
            elif warn:
                why = ("パケット不一致(DT・#L・#〆 で始まらない値)" if "ﾊﾟｹｯﾄ無し" in warn[0] else warn[0])
            else:
                why = pre or ("入力先が無い(入れる欄が決まりません。欄を押してください)"
                              if input_router.maker(text) in input_router.PACKETS else
                              "パケット不一致(DT・#L・#〆 で始まらない値)")
            self.receive_why = why

    def click(self, name: str) -> Result:
        return self._run(self._click, name)

    def _click(self, name: str) -> None:
        st = self.state
        self._confi_block()
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
        self._confi_block()
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
        if (name == "Manual" or name.startswith("Lateral") or name in st.cells
                or (len(name) == 2 and name[0] in "ZG" and name[1] in grid.POSITIONS)):
            self._confi_block()                                          # 測定値の手入力は確認の後
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
            # BH は値を持つだけ(下限値の表 Q_S_KE_ALL_053 は品質基準 DB から無くなった)。下限 0 = 「不要」
            wiring.apply_judge(st, judge.simply_judge(wiring.to_judge_form(st), name,
                                                      judge.bh_judg(st.bh[name[0]].value, 0),
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
                if input_router.ken_count_missing(st):
                    # 2026-11-10: 枚数を入れたら、検査回数も入れるよう知らせる(入れるまで測定は止まる)
                    env.messages.append(Message(input_router.KEN_COUNT_MISSING, WARN, source="KenHBox7_AfterUpdate"))
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
            if st.strain_face == Face.A:
                st.strain_face = Face.B
                env.messages.append(Message("B面歪 入力画面です", INFO, source="UF1:487"))
            elif st.strain_face == Face.B:
                st.strain_face = Face.A
            r = records.reflect_strain(self.store, wiring.to_snapshot(st, self.store), vbx.right(st.strain_face, 1))
            if r is not None:
                wiring.apply_restore(st, r, hooks)
            with st.stopped():
                for k in [f"Z{p}" for p in st.strain_z] + [f"G{p}" for p in st.strain_g]:  # SetSpecific…ToWhite
                    st.set_color(k, grid.COLOR_WHITE)
                hooks.hiz_loop_judg(st, 0)
            return
        if name == "DBflag":
            st.dbflag = value
            return
        if name in ("ATU", "Hab", "Tak", "Ap", "Am", "Hp", "Hm", "Tp", "Tm", "TAIK", "Lateral", "TaikText",
                    "LateralText"):
            self._resize_value(name, value)
            return
        if name == "Byousya":                                            # Byousya_AfterUpdate(印を付けても外しても)
            st.byousya = value in ("True", "true", "1", "on")
            layout.byousya_after_update(st, hooks)
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
            st.ctl(name).value = conv          # B43 を直した: D も自身の欄に変換して書く(VBA は S 欄へ)
            form = wiring.to_judge_form(st)
            out = (judge.lateralbow_judg_color(form, name, hooks.worker_name(st)) if lateral
                   else judge.right_angle_judg_color(form, name, hooks.worker_name(st)))
            wiring.apply_judge(st, out, env)
            hooks.posting(st)
        if st.ctl(name).value == "":
            st.set_color(name, grid.COLOR_WHITE)
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
        st, env = self.state, self._env()
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
            c["choices"] = reason_master.choices(reason_master.KIND_REMEASURE) or list(d.reasons)   # 理由選択のマスタ
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
            wiring.apply_values(st, checks.ng_pass(posi, tb1))
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
        with st.paused():
            self.hooks.load_text(st, "")

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
                    wiring.apply_judge(st, judge.loop_judg(wiring.to_judge_form(st), 0, rp, rp), env)
            self._ng_posting(posi)
            self._ng_erase(posi, "板廃棄")
        env.log(legacy_log.line_entry(c.get("line_name", st.line_name),
                                      posi + str(knum) + "_" + temp + ":" + "理由選択:" + rows + ":"
                                      + c.get("ken_name", "")))
        self.confi = None

    # ------------------------------------------------------------------
    # 画面に足りなかった操作(外観・不良要因・歪コンボ・検入高さ・オーダー・包装仕様・VC・保留・廃棄フラグ・
    # 印刷履歴クリア・サイズ修正・強制描写・検査員不一致)
    # ------------------------------------------------------------------
    def _worker_time(self) -> dict[str, dict[str, str]]:
        """測定値ごとの検査員・時刻・パケット(WorkerTime_<L>。VBA の Time_Posting と同じシート。2026-11-08 に画面へ)。"""
        st = self.state
        if st.lname == "":
            return {}
        return worker_archive.read_sheet(self.store, st.lname, st.lot)

    def _extra_more(self) -> dict[str, Any]:
        """view に足す値(選択肢・オーダー・包装仕様・VC・印刷ボタン)。判断はここで済ませる。"""
        st, env = self.state, self._env()
        vl = self.visual_lists
        vc = env.vc
        printed = st.lot != "" and report_model.is_printed(self.store, st.lot)
        worker_needed = False
        if self.started and st.line_name != "" and self.hooks is not None:
            try:
                worker_needed = self.hooks.register_check(st)            # UF1:4227 WorkButton を赤に
            except (datasource.DbMissing, IndexError, vbx.TypeMismatch):
                worker_needed = False
        cb = {}
        for p in grid.POSITIONS:
            c1, c3 = st.strain_cb[f"CB{p}1"].value, st.strain_cb[f"CB{p}3"].value
            cb[p] = {"1": list(env.cb_items["shape"]), "2": checks.cb_updown_items(c1),
                     "3": list(env.cb_items["gauge"]), "4": checks.cbg_updown_items(c3)}
        return {
            "lists": {"visual": vl.visual or [], "yama": vl.yama or [],
                      "defect": [{"name": n, "code": c} for n, c in (vl.defect or [])]},
            "cb_choices": cb,
            "yama_lb": st.v("YamaLB").value,
            "ken_hl": {k: {"value": st.v(k).value, "visible": st.v(k).visible}
                       for k in [f"HL{i}" for i in range(1, 7)] + [f"KenHL{i}" for i in range(1, 11)]},
            "ken_formula": st.v("高さ計算式").value, "ken_h_label": st.v("KenH").value,
            "ken_count_visible": st.v("KenCount").visible,
            "visual_caps": {f"{p}_{k}": {"value": st.v(f"{p}_{k}").value, "visible": st.v(f"{p}_{k}").visible}
                            for p in grid.POSITIONS for k in ("UpVisual", "UnderVisual")},
            "order": {"list": list(self.orders), "value": self.order, **self.packing_info},
            "pack": self.pack.to_dict() if self.pack is not None else None,
            "vc": None if vc is None else {
                "a": {"label": vc.l_a_caption, "visible": vc.l_a_visible and vc.a_visible,
                      "image": _picture_name(vc.a_picture)},
                "b": {"label": vc.l_b_caption, "visible": vc.l_b_visible and vc.b_visible,
                      "image": _picture_name(vc.b_picture)},
                "nothing": vc.nothing_text if vc.nothing_visible else ""},
            "buttons": {"abno": st.v("Abno").value, "aorb": st.strain_face,
                        "izyou": st.v("IzyouFl").color == "255",
                        "print": "再印刷" if printed else "印刷", "print_reset": printed,
                        "resize": st.resize_tag == "有", "byousya": bool(st.byousya),
                        "conductiv": st.v("Conductiv").color == "255"},
            "worker_needed": worker_needed,
            "name_list": self.name_list,
            "bh_note": {"qf": st.v("QF").value, "last_hot": st.v("LastHot").value},
            "strain_note": self._strain_note(),
            "inp_tag": st.inp_tag,
            "strain_frame": {p: {"lz": st.v("LZ" + p).value, "lg": st.v("LG" + p).value,
                                 "visible": st.v("FM" + p).visible} for p in grid.POSITIONS},
            "taik_text": st.v("TaikText").value,
            # 2026-11-10: 歪の公差(VBA の UF1 の HaHiz / TaHiz / Naka / Mimi / MG のラベル)を画面の「公差」と歪の見出しに出す
            "strain_tol": {"ha": st.v("HaHiz").value, "ta": st.v("TaHiz").value, "naka": st.v("Naka").value,
                           "mimi": st.v("Mimi").value, "mg": st.v("MG").value, "flatnes": st.v("flatnes").value},
        }

    def _strain_note(self) -> str:
        """UF1.MPEX_Change の歪ページの案内(UF1:1652-1663)。A 案は歪の欄を常に見せるので、文言を欄の見出しに出す。"""
        st = self.state
        if st.atu == "":
            return ""
        w = quality.weight(vbx.val(st.atu), vbx.val(st.hab), vbx.val(st.tak))
        if w == 0:
            return ""
        try:
            ab = checks.withor_without_gauge(wiring.to_uf1(st), wiring.work_dist(self.store, st.lname))
        except (vbx.TypeMismatch, IndexError, checks.VbaRuntimeError):
            return ""
        out = []
        if ab != "":
            out.append("ゲージ不要 * 両面定盤歪測定必要" if 35 >= w else "ゲージによる測定必要")
        if st.strain_face == Face.B and w > 35:
            out.append("B面歪 入力画面です")
        return " / ".join(out)

    def _visual_add(self) -> None:
        """OutTolerance.Visual_Add (OutTolerance:1513-1614)。外観異常名(品質基準「欠陥名」)・山数・不良要因。"""
        from .report_model import XSheet
        t = XSheet.from_template("不良要因印刷原紙", "不良要因印刷原紙")
        cell = lambda r, c: t.cells[(r, c)].value if (r, c) in t.cells else None   # noqa: E731
        self.visual_lists = checks.visual_add([cell(2, c) for c in range(3, 25)], [cell(9, c) for c in range(3, 25)])
        if self.visual_lists.error is not None:
            self._env().messages.append(self.visual_lists.error)

    def _activate_name_check(self) -> None:
        """UF1.UserForm_Activate の検査員不一致フラグ (UF1:4229-4246)。WorkerTime_<L> の R1 = "OUT" なら知らせて名簿画面へ。"""
        st, env = self.state, self._env()
        wt = line_sheet(records.WORKER_TIME, st.lname)
        if st.lname == "" or not self.store.exists(wt):
            return
        if records.cell_text(self.store.sheet(wt).get(1, 18)) == "OUT":
            env.messages.append(Message("Excelを閉じても解除はできません" + "\r\n" + "隠蔽工作Log出力", WARN,
                                        source="UF1:4231"))
            self._log("*UFKen:UFNameList起動隠蔽対策:Excel再起動の形跡")
            self._name_list_open(1)

    def _name_list_open(self, n: int) -> None:
        """OutTolerance.NameList_Open(n) (OutTolerance:545-)。名簿画面(/meibo)を開くよう画面に伝える。"""
        st = self.state
        if n == 1:
            reg = wiring.registry(st)
            self._log("*UFKen:" + "UFNameList起動" + " T1:" + reg.t1 + " T2:" + reg.t2 + " T3:" + reg.t3)
            if not self.store.exists(line_sheet(records.WORKER_TIME, st.lname)):
                return
        self.name_list = n

    def after_calibration(self) -> Result:
        """UF校正.INPG_Exit (UF校正:281-296)。校正に合格して戻ったら、止めていた判定をやり直す(LoopJudg・HizLoopJudg)。"""
        return self._run(self._after_calibration)

    def _after_calibration(self) -> None:
        st, hooks = self.state, self.hooks
        if not self.started or hooks is None:
            return None
        with st.paused():
            hooks.loop_judg(st)
            hooks.hiz_loop_judg(st)
        self.calibration_required = False
        return None

    def name_list_ack(self) -> Result:
        """名簿画面(UFNameList)への案内を出し終えた(画面が開いた)。"""
        return self._run(self._name_list_ack, posting=False)

    def _name_list_ack(self) -> None:
        self.name_list = None

    # --- 外観(LC1〜LC12)----------------------------------------------
    def visual(self, lc: int, items: Optional[list[int]] = None) -> Result:
        return self._run(self._visual, lc, items or [])

    def _visual(self, lc: int, items: list[int]) -> None:
        """UF1.LC1_Click〜LC12_Click (UF1:1458-1505) → LcColor (OutTolerance:1628-1704) → CheckLabelsByCategory。

        items = LBVisual で選んでいる行の番号(0 始まり。先頭は空白行)。
        """
        st, env, hooks = self.state, self._env(), self.hooks
        lc = int(lc)
        if lc not in st.visual:
            raise Stop(f"外観ボタン LC{lc} はありません", source="UF1.LC_Click")
        text = checks.get_selected_items(self.visual_lists.visual or [], [int(i) for i in items])
        res = checks.lc_color(lc, 1 if lc <= 6 else 7, text, st.lname, hooks.worker_name(st),
                              [i for i, c in st.visual.items() if c.color == wiring.COLOR_CYAN])
        st.visual[lc].value = ""                                          # UF1(La.Name).Caption = ""
        for i, c in st.visual.items():
            c.color = wiring.COLOR_CYAN if i in res.cyan else grid.COLOR_WHITE
        for k, v in res.captions.items():
            st.v(k).value = v
        # WriteLog(…, "", Lpath, "Measured.log") と WriteLog(…, "")(LcColor:1701-1702。Lpath は Lname で決まる)
        env.log(legacy_log.LogEntry(res.log_text, paths.join(paths.get("history"), res.log_folder,
                                                             res.log_folder + "_Log"), legacy_log.MEASURED_LOG))
        env.log(legacy_log.LogEntry(res.log_text))
        hooks.check_labels_by_category(st, "SUCDEF"[(lc - 1) % 6])

    def visual_clear(self, lc: int) -> Result:
        return self._run(self._visual_clear, lc)

    def _visual_clear(self, lc: int) -> None:
        """UF1.LC1_DblClick〜LC10_DblClick (UF1:1507-1546)。ボタンを白に戻す(LBVisual の選択は画面で外す)。"""
        lc = int(lc)
        if lc not in range(1, 11):                                        # LC11・LC12 には DblClick が無い
            return
        self.state.visual[lc].color = grid.COLOR_WHITE

    # --- 不良要因(DefectFactorLB・ListReflect・Cle)---------------------
    def defect_reflect(self, items: Optional[list[int]] = None) -> Result:
        return self._run(self._defect_reflect, items or [])

    def _defect_reflect(self, items: list[int]) -> None:
        """UF1.ListReflect_Click (UF1:1548-1605) → TecReturn → SelectionInput (OutTolerance:2356-2449)。"""
        st, env = self.state, self._env()
        defect = self.visual_lists.defect or []
        sel = sorted({int(i) for i in items if 0 <= int(i) < len(defect)})
        temps = checks.get_selected_items(defect, sel, defect=True)
        if not sel:                                                       # ListIndex = -1
            env.messages.append(Message("不良要因が選択されていません", ERROR, source="UF1:1552"))
            return
        if all(st.bad[p].value == "" for p in grid.POSITIONS):
            env.messages.append(Message("枚数入力がありません", ERROR, source="UF1:1558"))
            return
        n = {p: (vbx.clng(vbx.val(st.bad[p].value)) if vbx.is_numeric(st.bad[p].value) else 0) for p in grid.POSITIONS}
        if temps != checks.SPOT and any(v > 50 for v in n.values()):
            env.messages.append(Message("すみません、50枚以上は分けて入力してください", INFO, source="UF1:1572"))
            return
        te = checks.tec_return([bad_name(p) for p in grid.POSITIONS if st.bad[p].color == wiring.COLOR_CYAN])
        if te == "":                                                      # UF1("") → 実行時エラー
            raise Stop("実行時エラー '-2147024809 (80070057)': 指定されたオブジェクトが見つかりません。",
                       source="UF1:1578")
        # SelectionInput: 単一選択で名前が空の行は .Value = "" で抜ける
        if len(sel) == 1 and defect[sel[0]][0] == "":
            pass
        else:
            sheet = report_model.ensure_defect_sheet(self.store, st.lname, st.lot)
            code_row = [sheet.get(9, c) for c in range(3, 25)]
            try:
                w = checks.selection_input(temps, te, st.bad[te[0]].value, code_row, wiring.to_uf1(st))
            except checks.VbaRuntimeError as e:
                raise Stop("実行時エラー '9': インデックスが有効範囲にありません。", source="SelectionInput") from e
            if w is not None:
                with trace.timed("db.save", "不良要因の枚数(SelectionInput)", path=str(self.store.path),
                                 table=line_sheet(report_model.DEFECT_SHEET, st.lname)):
                    for r, c, v in w.cells:
                        sheet.write(r, c, v)
        # 水色(選んでいる)欄だけが反映される。ほかの欄に入れた枚数は VBA と同じく消えるので、黙って消さずに知らせる
        dropped = [f"{st.header(p) or p} {st.bad[p].value}枚" for p in grid.POSITIONS
                   if st.bad[p].value != "" and st.bad[p].color != wiring.COLOR_CYAN]
        if dropped:
            env.messages.append(Message("反映したのは選んでいる(水色の)欄だけです。次の枚数は反映していません: "
                                        + "、".join(dropped) + "。欄を押して選んでから、もう一度入れてください",
                                        WARN, source="ListReflect"))
        for p in grid.POSITIONS:                                          # リセット
            st.bad[p].value = ""
            st.bad[p].color = grid.COLOR_WHITE

    def defect_clear(self) -> Result:
        return self._run(self._defect_clear)

    def _defect_clear(self) -> None:
        """UF1.Cle_Click (UF1:567-577)。不良要因印刷_<L> の C3:X8・C10:X10・C11 を消す。"""
        name = line_sheet(report_model.DEFECT_SHEET, self.state.lname)
        if not self.store.exists(name):
            return
        sh = self.store.sheet(name)
        for r1, c1, r2, c2 in ((3, 3, 8, 24), (10, 3, 10, 24), (11, 3, 11, 3)):
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    sh.set(r, c, None)

    def defect_preview_html(self) -> Optional[str]:
        """不良タブの見本(2026-11-11): 印刷するのと同じ 不良要因印刷_<L> を帳票 HTML にする。読むだけ(シートは作らない)。

        まだシートが無い(一度も反映していない)ときは None。"""
        with self._lock:
            lname = self.state.lname
            if not lname:
                return None
            model = report_model.render_defect_sheet(self.store, lname)
            return None if model is None else report_html.render(model, title="不良要因票")

    # --- 検入高さの山数(YamaLB)-----------------------------------------
    def yama_lb(self, value: str) -> Result:
        return self._run(self._yama_lb, value)

    def _yama_lb(self, value: str) -> None:
        """UF1.YamaLB_Click (UF1:2508-2511) → DrawListBoxContent。厚が空なら選ぶだけ。"""
        if value not in checks.YAMA_ITEMS:
            raise Stop(f"山数 {value!r} は選べません", source="UF1.YamaLB")
        st = self.state
        st.v("YamaLB").value = value
        st.v("YamaLB").tag = str(checks.YAMA_ITEMS.index(value))
        if st.atu == "":
            return
        self.hooks.draw_list_box_content(st)

    # --- オーダーNo・包装仕様 -------------------------------------------
    def _ord_no_open(self) -> None:
        """SaveLoad.OrdNoOpen。オーダーNo の一覧を作り、包装仕様の表示を空にする。"""
        orders = wiring.ord_no_open(wiring.sheet_rows(self.store, LOT_HISTORY), self.state.lot)
        if orders is None:
            return
        self.orders, self.order = orders, ""
        self.packing_info = {"no": "", "form": "", "comment": ""}

    def ord_no_open(self) -> Result:
        """UF1.NoL_Click (UF1:1717-1719)。オーダーNo を読み直す。"""
        return self._run(self._ord_no_open)

    def select_order(self, order: str) -> Result:
        return self._run(self._select_order, order)

    def _select_order(self, order: str) -> None:
        """UF1.OrdNoCB_Change (UF1:1721-1790)。包装仕様No・梱包形態・タイトル → VC_Picture(オーダー) → MaterialsMsgBox。"""
        st, env, hooks = self.state, self._env(), self.hooks
        self.order = order
        hist = wiring.sheet_rows(self.store, LOT_HISTORY)
        try:
            r = wiring.ord_no_change(hist, st.lot, order, self.packing_info["no"], test=st.line_option == "Test")
        except datasource.TableMissing as e:
            env.messages.append(Message(str(e), ERROR, source="UF1.OrdNoCB_Change"))
            return
        if r.packing_no is not None:
            self.packing_info["no"] = r.packing_no
        if r.packing_form is not None:
            self.packing_info["form"] = r.packing_form
        if r.packing_comment is not None:
            self.packing_info["comment"] = r.packing_comment
        if not r.reload:
            return
        try:
            hooks.vc_picture(st, order)
        except (datasource.DbMissing, datasource.TableMissing, vbx.TypeMismatch, IndexError) as e:
            env.messages.append(Message(str(e), ERROR, source="SaveLoad.VC_Picture"))
        try:
            wiring.apply_judge(st, judge.materials_msgbox(wiring.to_judge_form(st), order,
                                                          arrbox=hooks._arrbox_fn(st)), env)
        except (lotmod.SubscriptOutOfRange, judge.VbaError, vbx.TypeMismatch) as e:
            env.messages.append(Message(str(e), ERROR, source="Judg.MaterialsMsgBox"))

    def _packing_open(self) -> bool:
        """FormRelat.PackingNoCheck → OutTolerance.UFPack_Open(UF1:4519-4520)。`packing.py`(全ヵ所・キャスクと共通)。

        画像(使用パス「品証やり取り」\\包装仕様書\\<No>.JPG)があるときだけ UF_PackNo を出す。VBA はモーダルで、
        「確認」に印を付けないと閉じられない。実行時エラー(76 / 9)は VBA と同じくロット登録をここで止める
        (偽を返す。DistArray 以降は通らない。全ヵ所の画面と同じ扱い)。
        """
        st, env = self.state, self._env()
        self.pack = None
        hist = wiring.sheet_rows(self.store, LOT_HISTORY)
        msgs: list[Message] = []
        try:
            n = packing.packing_no_check(hist, st.lot, messages=msgs)
            form = packing.pack_open(n, st.lot, hist, messages=msgs) if n != "" else None
        except packing.PackingError as e:
            env.messages.extend(msgs)
            env.messages.append(Message(f"実行時エラー '{e.number}': {e.description}", ERROR,
                                        source="FormRelat.PackingNoCheck"))
            return False
        except (datasource.DbMissing, datasource.TableMissing, quality.VbaRuntimeError, vbx.TypeMismatch) as e:
            env.messages.extend(msgs)
            env.messages.append(Message(str(e), ERROR, source="FormRelat.PackingNoCheck"))
            return False
        env.messages.extend(msgs)
        self.pack = form
        return True

    def pack_select(self, i: int) -> Result:
        return self._run(self._pack_select, i, posting=False)

    def _pack_select(self, i: int) -> None:
        """UF_PackNo の引当の行を押す(LabelCls → Loop_Label → FormRelat.JPGLoad)。"""
        if self.pack is None:
            raise Stop("包装仕様書の画面は開いていません", level=WARN, source="UF_PackNo")
        try:
            n = int(i)
        except (TypeError, ValueError):
            raise Stop("行の番号が正しくありません", level=WARN) from None
        msgs: list[Message] = []
        try:
            packing.pack_select(self.pack, n, wiring.sheet_rows(self.store, LOT_HISTORY), messages=msgs)
        except (datasource.DbMissing, datasource.TableMissing) as e:
            msgs.append(Message(str(e), ERROR, source="QP.MaterialsUpperPart"))
        self._env().messages.extend(msgs)

    def pack_close(self, confirmed: bool = False) -> Result:
        return self._run(self._pack_close, bool(confirmed), posting=False)

    def _pack_close(self, confirmed: bool) -> None:
        """UF_PackNo.CommandButton2_Click (UF_PackNo:25-48)。「確認」に印が無ければ閉じない。"""
        if self.pack is None:
            return
        m = packing.pack_close(confirmed)
        if m is not None:
            self._env().messages.append(m)
            return
        self.pack = None

    def pack_image(self) -> Optional[Path]:
        """いま出している包装仕様書の画像(使用パス QualityPath\\包装仕様書)。"""
        with self._lock:
            return packing.picture_file(self.pack.picture) if self.pack is not None else None

    # --- 保留・まとめて削除・廃棄フラグ ----------------------------------
    def hold(self, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._hold, answers or {})

    def _hold(self, answers: dict[str, bool]) -> Optional[Question]:
        """UF1.CommandButton18_Click「保留」(UF1:120-170)。いまの入力先(INP.Tag)の位置で NG_Pass(位置, "保留")。"""
        st, env = self.state, self._env()
        tag = st.inp_tag
        left = vbx.left(tag, 1)
        for prefix in ("Lateral", "Diag"):
            if tag.startswith(prefix) and (tag[len(prefix):] in grid.POSITIONS
                                           or (prefix == "Diag" and tag.endswith("Kou") and tag[4:5] in grid.POSITIONS)):
                left = tag[len(prefix)]
        if len(tag) == 2 and tag[0] in "ZG" and tag[1] in grid.POSITIONS:
            left = tag[1]
        if left == "":
            env.messages.append(Message("保留の根拠位置をクリックしてください。", INFO, source="UF1:155"))
            return None
        temp = st.header(left) if left in grid.POSITIONS else ""
        ans = answers.get("hold")
        if ans is None:
            return Question("hold", temp + "で保留処理を行いますか？", source="UF1:169")
        if ans:
            wiring.apply_values(st, checks.ng_pass(left, "保留"))
            self.confi = None                                             # NG_Pass: Unload UFConfi
        return None

    def ng_delete(self, pos: str, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._ng_delete, pos, answers or {})

    def _ng_delete(self, pos: str, answers: dict[str, bool]) -> Optional[Question]:
        """UF1.NGS_DblClick〜NGF_DblClick (UF1:330-365)。まとめて削除 → ConfiReason_Open(位置, 0, …, "S板1枚", DeleteC)。"""
        st = self.state
        if pos not in grid.POSITIONS:
            raise Stop(f"NG{pos} はありません", source="UF1.NG_DblClick")
        ans = answers.get("ng_delete")
        if ans is None:
            return Question("ng_delete", "まとめて削除しますか?", source="UF1:331")
        if ans:
            self.hooks.confi_reason_open(st, pos, 0, st.lname, pos + "板1枚", vbx.val(st.delete_c))
        return None

    def ng_release(self, pos: str, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._ng_release, pos, answers or {})

    def _ng_release(self, pos: str, answers: dict[str, bool]) -> Optional[Question]:
        """許容・保留を外す(2026-11-11 追加。VBA には無い: NG?_DblClick は「まとめて削除」、許容・保留は NG_Erase で
        NG 色が消えたときにしか消えない)。NG の値は消さず、NG? を「NG」(まだ NG 色の値があれば)か "" に戻し、
        CheckLabelsByCategory をやり直して、理由の行と同じ形でライン別 Access.log に残す。"""
        st, env = self.state, self._env()
        if pos not in grid.POSITIONS:
            raise Stop(f"NG{pos} はありません", source="ng_release")
        flag = st.ctl("NG" + pos).value.strip()
        if flag not in ("許容", "保留"):
            env.messages.append(Message(f"{pos} は許容・保留ではありません", INFO, source="ng_release"))
            return None
        ans = answers.get("ng_release")
        if ans is None:
            return Question("ng_release", f"{st.header(pos)}の{flag}を外しますか", source="ng_release")
        if not ans:
            return None
        rows = range(1, 11) if pos == "F" else range(1, 15)
        has_ng = any(st.ctl(f"{pos}{i}").value != "" and st.is_ng(f"{pos}{i}") for i in rows)
        st.ctl("NG" + pos).value = "NG" if has_ng else ""
        self.hooks.check_labels_by_category(st, pos)
        env.log(legacy_log.line_entry(st.line_name, pos + "_" + flag + ":" + "理由選択:" + flag + "解除:"
                                      + self.hooks.worker_name(st)))
        return None

    def izyou(self, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._izyou, answers or {})

    def _izyou(self, answers: dict[str, bool]) -> Optional[Question]:
        """UF1.IzyouFl_Click「廃棄ﾌﾗｸﾞ」(UF1:3213-3228)。異常ロット用_<L> の A1 がこのロットなら、2 回聞いてシートを消す。"""
        st = self.state
        name = line_sheet(records.ABNORMAL, st.lname)
        if not self.store.exists(name):
            return None
        if records.cell_text(self.store.sheet(name).get(1, 1)) != st.lot:
            return None
        st.v("IzyouFl").color = "255"
        a1 = answers.get("izyou")
        if a1 is None:
            return Question("izyou", "異常内容印刷フラグが立っています,取り消しますか？", source="UF1:3218")
        if not a1:
            return None
        a2 = answers.get("izyou_sure")
        if a2 is None:
            return Question("izyou_sure", "本当に取り消しますか？", source="UF1:3219")
        if a2:
            if self.store.exists(name):
                self.store.delete_sheet(name)
            st.v("IzyouFl").color = grid.COLOR_BUTTON
        return None

    # --- 印刷履歴クリア ---------------------------------------------------
    def print_reset(self, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._print_reset, answers or {})

    def _print_reset(self, answers: dict[str, bool]) -> Optional[Question]:
        """UF1.PrintBreset_Click「印刷履歴クリア」(UF1:3599-3617)。印刷済み(オレンジ)のときだけ。"""
        st, env = self.state, self._env()
        if st.lot == "" or not report_model.is_printed(self.store, st.lot):
            return None
        ans = answers.get("print_reset")
        if ans is None:
            return Question("print_reset", "このロットの印刷履歴を消しますか？" + "\r\n"
                            + "消せば印刷実行時にもう一度保存処理されます。", title="履歴削除の確認", source="UF1:3604")
        if ans:
            report_model.clear_printed(self.store, st.lot)
            env.messages.append(Message("履歴を削除しました。", INFO, source="UF1:3612"))
        else:
            env.messages.append(Message("履歴削除処理を中断しました。", WARN, source="UF1:3614"))
        return None

    # --- サイズ修正 -------------------------------------------------------
    def resize(self, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._resize, answers or {})

    def _resize(self, answers: dict[str, bool]) -> Optional[Question]:
        """UF1.ReSize_Click「ｻｲｽﾞ修正」(UF1:4057-4097)。はい なら ReSize.Tag="有"(サイズ・公差を編集できる)。最後に Posting。"""
        st = self.state
        ans = answers.get("resize")
        if ans is None:
            return Question("resize", "Size、公差の編集を行いますか？", source="UF1:4064")
        if ans:
            st.resize_tag = "有"
            st.v("ResizText").value = st.resize_tag
        return None

    def _resize_value(self, name: str, value: str) -> None:
        """Atu/Hab/Tak/Ap/Am/Hp/Hm/Tp/Tm/TAIK/Lateral/TaikText/LateralText の Change(UF1:442-551, 3199-3211, 4161-4200)。"""
        st = self.state
        if st.resize_tag != "有":
            raise Stop("ｻｲｽﾞ修正を押してから編集してください(Locked)", level=WARN, source="UF1.ReSize")
        if name in ("ATU", "Hab", "Tak", "Ap", "Am", "Hp", "Hm", "Tp", "Tm"):
            wiring.put(st, name, value if vbx.is_numeric(value) else "")
            if name not in ("ATU", "Hab", "Tak"):
                layout.form_formula(st)
        elif name == "TAIK":
            st.taik = value
            st.v("TaikText").value = ""                                   # Taik_Change(DebugTaikaku は表示だけ)
        elif name == "Lateral":
            st.lateral_grade = value
            st.lateral_text = ""                                          # Lateral_Change(DebugLate は表示だけ)
        elif name == "TaikText":
            st.v("TaikText").value = value
        elif name == "LateralText":
            st.lateral_text = value

    # ------------------------------------------------------------------
    # Min/Max
    # ------------------------------------------------------------------
    def min_max(self) -> Result:
        return self._run(lambda: self.env.messages.extend(minmax.run_min_max(self.state, self.hooks)))

    # ------------------------------------------------------------------
    # 印刷
    # ------------------------------------------------------------------
    def print_check(self, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._print_check, answers or {}, keep_ready=True)

    def _print_check(self, answers: dict[str, bool]) -> Optional[Question]:
        """PrintB_Click の前半 (UF1:3619-3730)。結果を view の `checks` と `print_ready` に入れる。"""
        self._require_line()
        st, env, hooks = self.state, self._env(), self.hooks
        assert hooks is not None
        self.print_ready = False
        self.report = {}
        if st.load == Switch.OFF:
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
                env.sound(snd.KENSA_ADVANCE_SHEET, "UF1:3647")
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
            {"label": "検入高さ・枚数", "ok": checks.h_check([uf.t(ken_h_name(i)) for i in range(1, 8)]) == ""},
            {"label": "山数", "ok": not (st.yama_visible and st.yama == "")},
        ]
        if pre.text != "":
            wiring._msg_error(env, st, pre.text, 1)
            self.checks = items + [{"label": "検査員の照合", "ok": None}, {"label": "梱包枚数", "ok": None}]
            return None
        env.sound(snd.KENSA_WAIT, "UF1:3717")
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
        if nm.out:                                                       # Name_Match → NameList_Open(1)
            self._name_list_open(1)
        # Number_Count
        nc = records.number_count(self.store, lname, st.lot)
        stop = nc is not None and nc.stop
        if nc is not None:
            env.messages.extend(nc.messages)
            if nc.stop:                                                  # Number_Count → NameList_Open(3)
                self._name_list_open(3)
        items.append({"label": "梱包枚数", "ok": None if nc is None else not nc.stop})
        if nm.out:
            # VBA は UFNameList を vbModal で出し、閉じるまで印刷に進まない。Web では印刷確定を止め、名簿・日報へ案内する
            stop = True
            env.messages.append(Message("登録検査員と測定者が一致しないので、印刷を止めました。"
                                        "「名簿・日報を開く」で直してから、もう一度 印刷・保存 を押してください",
                                        WARN, source="Name_Match"))
        self.checks = items
        self.print_ready = not stop
        return None

    def print_commit(self, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._print_commit, answers or {}, keep_ready=True)

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
        ans = answers.get("print")
        pending = getattr(self, "_print_pending", None)
        if ans is not None and pending is not None and pending[0] == st.lot:
            # 「印刷する？」に答えて来た: 確認の前に作った帳票をそのまま使う(作り直すと終了時刻が後の時刻に変わる)
            res = pending[1]
            self._print_pending = None
        else:
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
        if ans is None:
            self._print_pending = (st.lot, res)
            env.sound(snd.KENSA_PRINT, "UF1:3763")
            return Question("print", "印刷する？", source="UF1:3764")
        if not ans:
            env.messages.append(Message("ｷｬﾝｾﾙ処理", INFO, source="UF1:3840"))
            return None
        db = self.measure_db or paths.to_local(paths.db_file("梱包検査ﾃﾞｰﾀ"))
        if self.measure_db is None and not Path(db).is_file():
            # 前は無いと黙って新しく作り、取り違えたパスの空の DB に登録が溜まった。VBA と同じく止める
            raise Stop(f"測定DB(梱包検査ﾃﾞｰﾀ)が見つからないので保存できません: {db}\r\n"
                       "設定 → 使用パス の「梱包検査データ」を確かめてください(印刷・保存はしていません)",
                       source="SQL_追加")
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
            with self.store.transaction():                  # 保存用・耳用の 1 行は全部書けたときだけ
                saved = records.main_data_save(self.store, snap, db_path=db)
                # 測定値ごとの検査員と時刻を、履歴画面のために控える(2026-11-08。鍵 = 保存用の A 列)
                worker_archive.archive(self.store, lname, st.lot, records.cell_text(saved.values[1]))
            info["rows"] = saved.db.rowcount if saved.db else 0
            info["mode"] = saved.db.mode if saved.db else ""
        if saved.db is not None:
            env.messages.extend(saved.db.messages)
        self.report["saved"] = {"sheet": saved.sheet, "row": saved.row, "table": saved.db_table,
                                "mode": saved.db.mode if saved.db else ""}
        self.report["pdf"] = self._history_pdf(res.report, res.defect)    # 履歴保存
        with trace.timed("db.save", "履歴 CSV を書き出す(ExpCSV)", expected=(Stop,),
                         path=paths.get("history"), table=MIMI if snap.kind == "極厚" else HOZON) as info:
            try:
                csvs = records.export_history_csv(self.store, lname, st.line_name, snap.kind == "極厚",
                                                  local_dir=self.store.path.parent)
            except OSError as e:
                # 書き出せなくても印刷済みの印は付ける(付かないと印刷し直しで保存用・DB の行が二重になる)
                csvs = []
                env.messages.append(Message(f"履歴 CSV を書き出せませんでした: {e}(保存用・測定DB には保存済み)",
                                            WARN, source="ExpCSV"))
            for p in csvs:
                trace.note_used(p, "save")
            info["files"] = len(csvs)
        self.report["csv"] = [str(p) for p in csvs]
        with trace.timed("db.save", "履歴ブック(xlsx)と Lot情報履歴 を書き出す", expected=(Stop,),
                         path=paths.get("local_history"), table=LOT_HISTORY) as info:
            books: list[Path] = []
            try:
                books += records.export_lot_book(self.store, lname, st.line_name, st.lot, env.now(),
                                                 res.report, res.defect)
            except OSError as e:
                env.messages.append(Message(f"履歴ブック(xlsx)を書き出せませんでした: {e}", WARN, source="履歴保存"))
            try:
                books += records.export_lot_history(self.store, st.line_name, local_dir=self.store.path.parent)
            except OSError as e:
                env.messages.append(Message(f"Lot情報履歴を書き出せませんでした: {e}", WARN, source="履歴保存"))
            for p in books:
                trace.note_used(p, "save")
            info["files"] = len(books)
        self.report["xlsx"] = [str(p) for p in books]
        printed_at = env.now()
        report_model.set_printed(self.store, st.lot, lname, printed_at)
        try:                                            # 日報を all-tools の日報モードへ(2026-11-08)
            p = nippou_export.export_now(self.store, st, env, now=printed_at)
            self.report["nippou"] = str(p) if p else ""
        except (OSError, ValueError) as e:
            env.messages.append(Message(f"日報の CSV を書き出せませんでした: {e}", WARN, source="日報"))
        env.sound(snd.KENSA_DONE, "UF1:3832")
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
        folder = self.store.print_dir
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
    def history_restore(self, row: int, mimi: bool = False) -> Result:
        """UF管理 の読み込み(CommandButton1_Click → HistoryLoad。UF管理:620-635)。

        履歴画面で選んだ保存用(耳用)の行を検査画面に戻す。別のロットを検査中なら VBA と同じく止める
        (「別Lotデータがあります、新規登録処理を行ってください」)。
        """
        return self._run(self._history_restore, row, mimi)

    def _history_restore(self, row: int, mimi: bool) -> None:
        if not self.started:
            self._start()
        st, hooks = self.state, self.hooks
        assert hooks is not None
        lots = records.kanri_lots(self.store, mimi=mimi)
        if not 2 <= row < len(lots) + 2:
            raise Stop("その行はありません", source="UF管理")
        sheet = self.store.sheet(MIMI if mimi else HOZON)
        lot, when = lots[row - 2], records.cell_text(sheet.get(row, 1))
        r = records.kanri_load(self.store, lot, when, uf1_visible=True, uf1_lot=st.lot, mimi=mimi,
                               lname=st.lname, aorb=st.strain_face, row_no=row)
        if not r.found:
            # 止めた理由は Stop で 1 回だけ出す(知らせにも入れると同じ文言が 2 回出る)
            hooks.env.messages.extend(r.messages[:-1])
            raise Stop(r.messages[-1].text if r.messages else "該当する履歴がありません", source="UF管理:622")
        hooks.env.messages.extend(r.messages)
        wiring.apply_restore(st, r, hooks)

    # ------------------------------------------------------------------
    # D28: ロットごとの「控え(第2)から読む」
    # ------------------------------------------------------------------
    def set_lot_source(self, on: Any = True, lot: str = "", answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._lot_source, on in (True, "true", "True", "1", 1, "on"), lot, answers or {})

    def _set_lot_source(self, source: str, lot: str, why: str) -> None:
        st = self.state
        before = (st.lot_source, st.lot_source_lot)
        st.lot_source, st.lot_source_lot = (wiring.LOT_SECOND, lot) if source else ("", "")
        if before != (st.lot_source, st.lot_source_lot):
            trace.event("kensa.lot_source", ("控え(第2)から読む: " + lot) if source else
                        f"控え(第2)から読むのをやめました({why})",
                        data={"lot": lot or before[1], "on": bool(source), "before": before[1] if before[0] else "",
                              "why": why, "mode": paths.mode(), "line": st.lname})
            self._log(("控えから読む:" + lot if source else "控えから読む解除:" + before[1]) + ":" + st.lname)

    def _lot_source(self, on: bool, lot: str, answers: dict[str, bool]) -> Optional[Question]:
        """「控え(第2)から読む」の切替(新規 D28。VBA で DBflag に文字を入れて CB_Change を起こしていた運用の代わり)。

        - 入れる: 入力欄の LotNo(空ならいまのロット)に選ぶ。いまのロットならすぐ控えから読み直す(CB_Change)。
          別のロットならそのロットを Lot登録 する(確認は いつもの Lot登録 と同じ)
        - 外す: いまのロットなら いつもの読み方(モードどおり)で読み直す
        DBflag と違い、校正・Min/Max の確かめは飛ばさない。
        """
        self._require_line()
        st = self.state
        text = "".join(ch.upper() if "a" <= ch <= "z" else ch for ch in str(lot or "")
                       if ch.isascii() and ch.isalnum())
        if not on:
            prev = st.lot_source_lot if st.lot_source else ""
            if prev == "":
                return None
            self._set_lot_source("", "", "画面で外した")
            if prev == st.lot and st.lot != "":
                layout_cache_forget(st)
                self.env.messages.append(Message("Lot " + st.lot + " を いつもの読み方(" +
                                                 paths.MODE_LABELS.get(paths.mode(), "") + ")で読み直します",
                                                 INFO, source="D28"))
                return self._cb_change(answers)
            return None
        target = text if len(text) == 7 else st.lot
        if target == "":
            raise Stop("控え(第2)から読むロットの LotNo(7桁)を入れてください", level=WARN, source="D28")
        self._set_lot_source(wiring.LOT_SECOND, target, "画面で選んだ")
        if target == st.lot:
            layout_cache_forget(st)
            self.env.messages.append(Message("Lot " + target + " を控え(第2)から読み直します", INFO, source="D28"))
            return self._cb_change(answers)
        return self._register_lot(target, answers)

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
            # 2026-11-09(意図した差異。docs/06): 測定値が 1 つも無ければ、残すものが無いので履歴の保存を求めずに新規へ。
            # VBA は保存用・耳用にロットが無いと「該当LotNoが保存されていません」で止めていた(Lot登録だけで測っていなくても)
            # 2026-11-09(続き): 数えるのは測定表の欄(S1〜F14)と歪の数値だけ。空白・見出しの文字・"0.00" だけの歪は数えない
            def _num(v: str) -> bool:
                t = str(v or "").strip()
                return any(ch.isdigit() and ch != "0" for ch in t)
            # F7(2026-11-10): 測定表の欄も歪と同じ決まり(数値だけ。"0.00"・空白・文字だけは数えない)
            hits = [k for k, c in st.cells.items() if CellKey.parse(k) is not None and _num(c.value)]
            hits += [f"歪{k}" for k, z in st.strain_z.items() if _num(z.value)]
            measured = bool(hits)
            if measured:
                self._log("UFKen" + "*新規検査: 測定値あり " + ",".join(hits[:10]))
            found = not measured
            for sheet in (HOZON, MIMI) if measured else ():
                rows = wiring.sheet_rows(self.store, sheet) or []
                if any(records.cell_text(r[1]) == st.lot for r in rows[1:] if len(r) > 1):
                    found = True
                    break
            if not found:
                wiring._msg_error(env, st, "該当LotNoが保存されていません"
                                  + (f"\r\n(測定値があります: {'・'.join(hits[:5])})" if hits else ""), 2, "END")
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
        if st.lot_source != "":
            self._set_lot_source("", "", "新規検査なので")                   # D28
        for p in grid.POSITIONS:
            st.v("LZ" + p).value = "定盤:"
            st.v("LG" + p).value = "ｹﾞｰｼﾞ"
        hooks.cb_add(st, st.sika_set_atu)                               # UF1:3431 CBAdd(歪コンボ。2026.1.12 の直し)
        st.v("VC_L_A").tag = st.v("VC_L_B").tag = ""                    # UF1:3445-3446
        st.v("SplitFlag").value = ""                                    # 分割ロットの印を次のロットへ持ち越さない
        self.orders, self.order, self.pack = [], "", None
        self.packing_info = {"no": "", "form": "", "comment": ""}
        self.calibration_required = False
        st.amin = st.amax = st.hmin = st.hmax = st.tmin = st.tmax = ""
        st.flatnes_text = st.conductiv_text = st.comment = st.reien = st.ken_no_caption = st.ken_head = st.additional = ""
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
        if st.strain_face == Face.B:
            st.strain_face = Face.A
        ws = self.store.sheet(line_sheet(records.WORK_SAVE, lname))
        wiring.apply_restore(st, records.losd_strain("A", ws.read_range(2, 15, 7, 16), ws.read_range(8, 15, 13, 16)),
                             hooks)
        hooks.load_text(st, "ReLoad")
        hooks.worker_load(st)
        layout.at_count_select(st, 4, hooks)
        layout.inspection_select(st, "表示", hooks)
        layout.right_angle_display(st, "新規")
        layout.lselect_lateral(st, "新規")
        hooks.vc_clear(st)                                              # UF1:3494 VC_Clear(前ロットの VC 画像を消す)
        st.load = Switch.ON
        self.confi, self.checks, self.print_ready, self.report = None, [], False, {}
        return None


def layout_cache_forget(st: KensaState) -> None:
    """D28: 読み先を変えたので、LSelect の控え(g_CacheLotNo など。いまのロットの検番・コース・厚)を使わせない。

    ClearLSelectCache と違い「該当ロットがありません」を 1 回黙らせる印(g_SkipLotNotFoundMsg)は立てない。
    """
    c = st.cache
    c.lot_no = c.ken_no = c.cours = c.c_course = ""
    c.ord_atu = 0.0


def lot_reads_second(lot: str, kensa: Optional["Session"] = None) -> bool:
    """D28: ほかの画面(品質基準・キャスク・全ヵ所)が、このロットを控え(第2)だけから読むべきか。

    通常検査でそのロットに「控え(第2)から読む」を選んでいるときだけ真(状態は読むだけ)。
    """
    try:
        st = (kensa or get()).state
    except Exception:                                    # noqa: BLE001 - 通常検査が無ければ いつもの読み方
        return False
    return wiring.reads_second(st, lot)


def _picture_name(path: Optional[str]) -> str:
    """画像のパスからファイル名だけ(画面は /api/kensa/picture/<種類>/<名前> で読む)。"""
    if not path:
        return ""
    return path.replace("/", "\\").rsplit("\\", 1)[-1]



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
