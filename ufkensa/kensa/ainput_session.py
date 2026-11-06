"""Aインプット画面(VBA の UFAinput)のセッション ── 画面(app/routes/ainput.py)はここだけを呼ぶ

VBA では検査画面(UF1)の「Aインプット」ボタン(AMode_Click)が `Ainput_Startup("UF1", KenM)` を呼び、
UF1 のロット・Min/Max・検査員を UFAinput に写して開いた。ここでも開くときに検査画面のセッション
(`session.get().state`)から読む。**読むだけで、検査画面の状態は書き換えない。**

部品は `ainput.py`(移植済み)。判断はすべてそちら。ここは
    - UFAinput のコントロールの値(`ainput.AinputForm`)と、フォームが持っていた Hiki_Lot シート・
      MaterialFlag などを 1 つ持つ
    - 検査画面の状態・手元DB(Lot情報履歴・Daily_<L>)・測定機器ID を部品に渡す
    - 画面に出す形(`view()`)を作る
だけを行う。

**画面への返し方**は検査画面と同じ `session.Result`(view・messages・question・stop)。

各操作と、対応する VBA のイベント:
    open(ken_m)            UF1.AMode_Click → Input.Ainput_Startup → UFAinput.UserForm_Activate
    set_value(name, value) 入力欄の Change / BeforeUpdate(HikiNo・KONNO・Remain・MAI1・PileMatch …・Ai/Ga/Vab)
    spin(name, dir)        MAISpin / MaterialSpin / ExMaterialSpin(2 段目・風袋 2〜3 を出す/消す)
    set_mount(value)       Mount_Count_Click → Input.Number_Inp(山数。None = InputBox のキャンセル)
    receive(text)          StartTB_KeyDown(Tab)(伝送文字列 → 次の空いた山)
    clear_height(n)        動的ラベル(clsDynLabel)のクリック = その山の高さを消す
    press(answers)         AinButt_Click「Icas出力」(2 回押し)
    unlock()               CommandButton10「外注出荷」
    reset(answers)         CommandButton2「Reset」(入力を消して hostsend.txt を書き直し、送信アプリを起動)
"""
from __future__ import annotations

import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional

from .. import trace
from . import ainput, datasource, instruments, lines, paths, records, vbx, wiring, workers
from .ainput import MAX_MOUNT, AinputForm
from .flow import ERROR, INFO, WARN, Message, Question, Stop
from .store import LOT_HISTORY, line_sheet

#: UF1 の検査法 OptionButton(Inspection1〜4 の Caption。UF1:448-459)
KEN_METHODS = ("L1", "L2", "M6", "M2")
#: 送信の段階(画面の右に並べる)
STEPS = ("入力チェック", "梱包高さ", "転送の確認", "CSV・hostsend.txt", "送信アプリ起動", "result.txt", "日報")
RESULT_CODES = ("A000", "A001", "A002", "A003")
NUMERIC_FIELDS = ("mai1", "cou1", "mai2", "cou2", "material1", "material2", "material3", "gw", "ex_case_no",
                  "ex_material1", "ex_material2", "ex_material3", "ex_gw", "outside_atu", "outside_haba",
                  "outside_take")


def _kensa():
    from . import session
    return session.get()


def _f2(x: float) -> str:
    """表示用 `Format(<Single>, "0.00")`。"""
    return vbx.fmt(vbx.cstr(vbx.single(x), is_single=True), "0.00")


@dataclass
class _Last:
    """最後の「Icas出力」の結果(画面の右の段階表示)。"""
    stage: str = ""
    at: str = ""
    csv_path: str = ""
    hostsend_path: str = ""
    launched: Optional[bool] = None
    launch_ok: Optional[bool] = None
    result_text: str = ""
    result_missing: bool = False
    result_ok: Optional[bool] = None
    daily_rows: int = 0
    note: str = ""


@dataclass
class _Received:
    """最後に受けた伝送文字列と、入った山(入らなかったときはその訳。VBA は黙って消すだけ)。"""
    text: str = ""
    index: int = 0
    reason: str = ""


class AinputSession:
    """プロセスに 1 つの Aインプット画面。試験では `kensa`(検査画面のセッション)・`launcher`・`sleep` を差し替える。"""

    def __init__(self, kensa: Any = None, *, launcher: Optional[Callable[[Path], bool]] = None,
                 sleep: Callable[[float], None] = time.sleep) -> None:
        self._lock = threading.RLock()
        self._kensa_obj = kensa
        self.launcher = launcher
        self.sleep = sleep
        self.form: Optional[AinputForm] = None
        self.ken_m = ""
        self.source = ""                     # 呼び元(""=通常検査 / Cask / Mitsubishi / H494)
        self.hiki_rows: list[list[Any]] = []        # Hiki_Lot シート(No_Search の結果)
        self.hiki_list: list[str] = []              # HikiNo の選択肢(UserForm_Activate のソート後)
        self.adjust_label = "調整No"
        self.adjust_no = ""
        self.adjust_visible = True
        self.material_flag = "1"
        self.ex_material_flag = "1"
        self.check: Optional[dict[str, Any]] = None  # 最後の梱包高さチェック
        self.last = _Last()
        self.sends: list[dict[str, Any]] = []
        self.received = _Received()
        self.messages: list[Message] = []

    # ------------------------------------------------------------------
    # 土台
    # ------------------------------------------------------------------
    @property
    def kensa(self) -> Any:
        return self._kensa_obj if self._kensa_obj is not None else _kensa()

    def _run(self, fn: Callable[..., Any], *args: Any) -> Any:
        from .session import Result
        with self._lock:
            self.messages = []
            try:
                try:
                    q = fn(*args)
                except datasource.DbMissing as e:                       # VBA は End(D10: 止めるが状態は保つ)
                    raise Stop(str(e), messages=list(self.messages), source="datasource") from e
                except vbx.TypeMismatch as e:                            # VBA の実行時エラー 13
                    raise Stop("実行時エラー 13: 型が一致しません。" + str(e), messages=list(self.messages)) from e
                except IndexError as e:                                  # VBA の実行時エラー 9(B69 など)
                    raise Stop("実行時エラー 9: インデックスが有効範囲にありません。" + str(e),
                               messages=list(self.messages)) from e
            except Stop as e:
                msgs = list(self.messages)
                for m in e.messages:
                    if not any(m is x for x in msgs):
                        msgs.append(m)
                return Result(self.view(), msgs, stop=e)
            return Result(self.view(), list(self.messages), question=q if isinstance(q, Question) else None)

    def _form(self) -> AinputForm:
        if self.form is None:
            raise Stop("Aインプットを開いてください(検査法を選んで「開く」)。", level=WARN)
        return self.form

    def _env(self) -> wiring.Env:
        from . import terminal
        k = self.kensa
        return wiring.Env(store=k.store, now=k.now, facility=terminal.facility, write_logs=k.write_logs)

    def _history(self) -> Optional[list[list[Any]]]:
        """Lot情報履歴(1 行目 = 見出し)。"""
        return wiring.sheet_rows(self.kensa.store, LOT_HISTORY)

    # ------------------------------------------------------------------
    # 開く
    # ------------------------------------------------------------------
    def open(self, ken_m: str = "", source: str = "") -> Any:
        return self._run(self._open, ken_m, source)

    def _source_sheet(self, source: str) -> tuple[str, dict[tuple[int, int], Any], str, str]:
        """キャスク・全ヵ所・H494 から開いたとき(Input:247-252, 361-364): そのシートのロット(D2)・Min/Max のセル・
        検査員の文字・ライン名。前は呼び元にかかわらず通常検査(UF1)のロット・Min/Max で開いていた(2026-11-06)。"""
        from .modes.sheet import CASK, H494, ZEN
        if source == "Cask":
            from . import cask_session
            book = cask_session.get().book
            if book is None or not book.has(CASK):
                raise Stop("キャスク画面でロットを開いてから、Aインプットを開いてください", level=WARN, source="Input:249")
            ws, workers_txt, line = book.ws(CASK), "UFｷｬｽｸ:" + book.cask.kname, book.cask.line_name
        else:
            from . import zenkasho_session
            book = zenkasho_session.get().book
            name = ZEN if source == "Mitsubishi" else H494
            if book is None or not book.has(name):
                raise Stop("全ヵ所・H494 の画面でロットを開いてから、Aインプットを開いてください", level=WARN,
                           source="Input:250")
            ws, line = book.ws(name), book.mit.line_name
            workers_txt = "UF三菱:" + (book.mit.kname if source == "Mitsubishi" else book.mit.h494_kname)
        cells = {(r, c): ws.get(r, c) for r in (2, 3, 4) for c in range(1, 28)}
        lot = "" if cells.get((2, 4)) is None else str(cells[(2, 4)])
        return lot, cells, workers_txt, line

    def _open(self, ken_m: str, source: str = "") -> None:
        """UF1.AMode_Click (UF1:447-479) → Input.Ainput_Startup (Input:225-400) → UFAinput.UserForm_Activate。"""
        from .session import NO_LINE_TEXT
        k = self.kensa
        st = k.state
        if st.line_name == "":
            raise Stop(NO_LINE_TEXT, level=WARN, source="UF1:225")
        ken_m = ken_m or self.ken_m
        if ken_m not in KEN_METHODS:
            raise Stop("検査法を選択してください", level=ERROR, source="UF1:457")
        self.ken_m = ken_m
        if st.load == "off":                                            # VBA は何もせず抜ける(UF1:461)
            raise Stop("検査画面の処理中です。終わってからもう一度開いてください。", level=INFO, source="UF1:461")
        ln = lines.find(st.line_name)
        test = st.line_option == "Test" or (ln is not None and ln.official == "Test")
        store = k.store
        source = source if source in ("Cask", "Mitsubishi", "H494") else ""
        self.source = source
        with trace.timed("db.read", "Lot情報履歴 を読む(Aインプット)", path=str(store.path)):
            history = self._history()
        if source:
            lot, cells, workers_txt, line = self._source_sheet(source)
            res = ainput.startup(source, ken_m, history, working_lot=lot, lname=line or st.lname,
                                 uf1_line=st.line_name, sheet_cells=cells, workers=workers_txt)
        else:
            mm = {"amin": st.amin, "hmin": st.hmin, "tmin": st.tmin, "amax": st.amax, "hmax": st.hmax,
                  "tmax": st.tmax}
            if not (test or st.dbflag != "" or all(v != "" for v in mm.values())):
                raise Stop("MinMax処理を行ってください", level=INFO, source="UF1:474")
            reg = wiring.registry(st)
            day_max = wiring.worker_day_max(store, st.lname)
            res = ainput.startup("UF1", ken_m, history, working_lot=st.lot, lname=st.lname, uf1_line=st.line_name,
                                 uf1_minmax=mm, workers=workers.workers_text(day_max, reg))
        self.messages += res.messages
        if res.form is None:
            last = self.messages.pop() if self.messages else Message("Aインプットを開けません", ERROR)
            raise Stop(last.text, level=last.level, source=last.source, messages=list(self.messages))
        if res.log is not None and k.write_logs:
            res.log.write(now=k.now())
        f = res.form
        f.remain = "0"                      # Remain の設計時の値
        f.mount_count = "1"                 # Mount_Count の設計時の Caption
        self.form = f
        hiki = res.hiki
        self.hiki_rows = list(hiki.rows) if hiki is not None and hiki.rows is not None else []
        self.hiki_list = ainput.sort_hiki([ainput.nz(r[0] if r else None) for r in self.hiki_rows])
        self.adjust_label, self.adjust_no, self.adjust_visible = "調整No", "", True
        self.material_flag = self.ex_material_flag = "1"
        self.check = None
        self.last = _Last()
        self.received = _Received()

    # ------------------------------------------------------------------
    # 入力欄
    # ------------------------------------------------------------------
    def set_value(self, name: str, value: Any) -> Any:
        return self._run(self._set_value, name, value)

    def _set_value(self, name: str, value: Any) -> None:
        f = self._form()
        v = "" if value is None else str(value)
        if name == "hiki_no":
            self._hiki_no(v)
        elif name == "konno":                               # KONNO_Change → KONNO_BeforeUpdate → Change
            f.konno = ainput.numeric_only(ainput.konno_before_update(ainput.numeric_only(v)))
        elif name == "remain":                              # Remain_Change → Remain_BeforeUpdate(→ Change)
            f.remain = v
            ainput.remain_change(f)
            after = ainput.remain_before_update(f.remain)
            if after != f.remain:
                f.remain = after
                ainput.remain_change(f)
        elif name in NUMERIC_FIELDS:
            setattr(f, name, ainput.numeric_only(v))
        elif name == "pile_match":                          # PileMatch_KeyPress(1 文字ずつ)→ PileMatch_Change
            f.pile_match = "".join(ainput.pile_match_key(f.remain, ch) for ch in v)
            self.messages += ainput.pile_match_change(f)
        elif name == "pile_match_no":
            f.pile_match_no = v
            ainput.pile_match_no_change(f)
        elif name == "gaityuu_text":                        # GaityuuText_BeforeUpdate(ログ)
            if f.locked:
                return
            if v != f.gaityuu_text and self.kensa.write_logs:
                ainput.gaityuu_log(self.kensa.state.line_name).write(now=self.kensa.now())
            f.gaityuu_text = v
        elif name in ("ai", "ga", "vab"):                   # Locked の間は変えられない
            if f.locked:
                return
            setattr(f, name, v.lower() in ("1", "true", "on"))
        else:
            raise Stop(f"その入力欄はありません: {name}", level=WARN)

    def _hiki_no(self, v: str) -> None:
        """UFAinput.HikiNo_Change (UFAinput:317-367)。一覧に無い値は ListIndex = -1(オーダーNo だけ消える)。"""
        f = self._form()
        f.hiki_no = v
        f.order_no = ""
        self.adjust_visible = True
        sel = ainput.hiki_no_change(v if v in self.hiki_list else None, self.hiki_rows, db_flag=f.db_flag)
        if v not in self.hiki_list:
            return
        f.order_no = sel.order_no
        if sel.lot is not None:
            f.lot = sel.lot
        if sel.ken_no is not None:
            f.ken_no = sel.ken_no
        if sel.adjust_label:
            self.adjust_label, self.adjust_no = sel.adjust_label, sel.adjust_no
        self.adjust_visible = sel.adjust_visible
        if sel.export_visible is not None:
            f.ex_visible = sel.export_visible

    def spin(self, name: str, dir: str) -> Any:                     # noqa: A002 - 画面の項目名
        return self._run(self._spin, name, dir)

    def _spin(self, name: str, dir: str) -> None:                   # noqa: A002
        """MAISpin / MaterialSpin / ExMaterialSpin の SpinUp(減らす)・SpinDown(増やす)(UFAinput:393-465)。"""
        f = self._form()
        up = dir == "up"
        if name == "mai":
            if up and f.mai_flag == "2":
                f.mai2 = f.cou2 = ""
                f.mai_flag = "1"
            elif not up and f.mai_flag == "1":
                f.mai_flag = "2"
            return
        if name not in ("material", "ex_material"):
            raise Stop(f"その切り替えはありません: {name}", level=WARN)
        prefix = "material" if name == "material" else "ex_material"
        flag = self.material_flag if name == "material" else self.ex_material_flag
        if up:
            if flag == "3":
                setattr(f, prefix + "3", "")
                flag = "2"
            elif flag == "2":
                setattr(f, prefix + "2", "")
                flag = "1"
        else:
            flag = {"1": "2", "2": "3"}.get(flag, flag)
        if name == "material":
            self.material_flag = flag
        else:
            self.ex_material_flag = flag

    def set_mount(self, value: Optional[str]) -> Any:
        return self._run(self._set_mount, value)

    def _set_mount(self, value: Optional[str]) -> None:
        """Mount_Count_Click → Number_Inp (Input:402-428)。処理中(Tag="Load")は受け付けない。"""
        f = self._form()
        if f.tag == "Load":
            return
        self.messages += ainput.input_mount_count(f, None if value is None else str(value))

    # ------------------------------------------------------------------
    # 山高さ
    # ------------------------------------------------------------------
    def receive(self, text: str) -> Any:
        return self._run(self._receive, text)

    def _receive(self, text: str) -> None:
        """StartTB_KeyDown(Tab)(UFAinput:697-746)。≤300 はデプス(機器番号が測定機器ID と同じ)、>300 は #L。"""
        f = self._form()
        text = "" if text is None else str(text)
        self.received = _Received(text)
        if f.start_tb_tag != "" or f.mount_count == "":
            self.received.reason = "山数がありません"
            return
        mc = vbx.clng(vbx.val(f.mount_count))
        idx = ainput.next_empty_height_index(f, mc)
        lh = ainput.l_height(f)
        depth = ""
        if lh <= 300.0:
            depth = instruments.db_return_micro(f.line_name, ainput.depth_set_atu(f), " ", "", "デプス")
        before = list(f.heights)
        self.messages += ainput.accept_height(f, text, depth)
        if f.heights != before:
            self.received.index = idx
            return
        # 入らなかった訳(VBA は StartTB を消すだけ。画面で分かるように添える)
        if idx == -1:
            self.received.reason = "全山入力済みです(測り直す山は消してから)"
        elif lh <= 300.0:
            if ainput.maker(text) == "DT":
                self.received.reason = (f"Depth測定です。機器 {vbx.mid(text, 3, 5)} は、このラインのデプス"
                                        f"({depth or '登録なし'})ではありません")
            else:
                self.received.reason = "Depth測定です(デプスの伝送 DT… を送ってください)"
        else:
            self.received.reason = "ｺﾝﾍﾞｯｸｽ測定です(#L… を送ってください)"

    def clear_height(self, n: Any) -> Any:
        return self._run(self._clear_height, n)

    def _clear_height(self, n: Any) -> None:
        """clsDynLabel.m_lbl_Click: その山の高さを消す。"""
        f = self._form()
        try:
            i = int(n)
        except (TypeError, ValueError):
            raise Stop("山の番号が正しくありません", level=WARN) from None
        if 1 <= i <= MAX_MOUNT:
            f.heights[i - 1] = ""

    # ------------------------------------------------------------------
    # Icas出力(2 回押し)
    # ------------------------------------------------------------------
    def press(self, answers: Optional[dict[str, bool]] = None) -> Any:
        return self._run(self._press, answers or {})

    def _press(self, answers: dict[str, bool]) -> Optional[Question]:
        """UFAinput.AinButt_Click (UFAinput:932-1137)。判断は ainput.ain_button。"""
        f = self._form()
        k = self.kensa
        st = k.state
        store = k.store
        mc = vbx.clng(vbx.val(f.mount_count))
        snapshot = list(f.heights[:mc]) if 1 <= mc <= MAX_MOUNT else []
        history = self._history()
        env = self._env()
        daily = wiring.daily_context(st, env)
        now = k.now()
        r = ainput.ain_button(f, history_rows=history[1:] if history is not None else None, answers=answers,
                              registry=wiring.registry(st), day_max=wiring.worker_day_max(store, st.lname), now=now,
                              daily=daily, launcher=self.launcher, sleep=self.sleep)
        self.messages += r.messages
        if r.height is not None:
            self.check = self._check_view(r.height, snapshot, history is not None, r.stage)
        if r.output is not None:
            trace.note_used(r.output.csv_path, "save")
            trace.note_used(r.output.hostsend_path, "save")
        if r.daily_writes:
            name = line_sheet(records.DAILY, st.lname)
            with trace.timed("db.save", f"{name} に日報を書く(Aインプット)", path=str(store.path)):
                wiring.write_cells(store, name, r.daily_writes)
            trace.note_used(store.path, "save")
        if r.stage not in ("dbflag", "missing", "busy"):
            self._record(r, now)
        return r.question

    def _record(self, r: ainput.AinResult, now: datetime) -> None:
        last = _Last(stage=r.stage, at=now.strftime("%H:%M:%S"))
        if r.output is not None:
            last.csv_path, last.hostsend_path = r.output.csv_path, r.output.hostsend_path
        if r.launch is not None:
            last.launch_ok, last.launched = r.launch.ok, r.launch.launched
        if r.result is not None:
            last.result_ok, last.result_text, last.result_missing = r.result.ok, r.result.text, r.result.missing
        last.daily_rows = len(r.daily_writes)
        if r.stage == "input_error" and r.messages:
            last.note = r.messages[-1].text
        elif r.stage == "height_ng" and r.height is not None:
            last.note = r.height.msg_error
        elif r.stage == "stopped":
            last.note = r.messages[-1].text if r.messages and r.messages[-1].text else \
                "送信停止中(前の失敗の Ngflag が残っています。開き直すと解けます)"
        self.last = last
        if r.stage in ("sent", "stopped"):
            f = self.form
            self.sends.insert(0, {"at": last.at, "stage": r.stage, "hiki_no": f.hiki_no if f else "",
                                  "lot": f.lot if f else "", "note": last.note or self._result_label(last)})
            del self.sends[10:]

    def _check_view(self, hc: ainput.HeightCheckResult, heights: list[str], has_history: bool,
                    stage: str) -> dict[str, Any]:
        """梱包高さチェックの見せ方。判定は合計(Height_Check)だけ。山ごとの外れは目安(合計の理論値 ÷ 山数)。"""
        judged = has_history and not hc.show_uf1
        mc = len(heights)
        lo = vbx.single(hc.gm / mc) if mc else 0.0
        hi = vbx.single(hc.gp / mc) if mc else 0.0
        rows = []
        for i, h in enumerate(heights, start=1):
            v = vbx.single(vbx.val(h))
            diff = 0.0
            if judged and h != "":
                if v > hi:
                    diff = vbx.single(v - hi)
                elif v < lo:
                    diff = vbx.single(v - lo)
            out = (hc.kind == "over" and diff > 0) or (hc.kind == "short" and diff < 0)
            rows.append({"n": i, "value": h, "diff": ("+" if diff > 0 else "") + _f2(diff) if diff else "",
                         "out": bool(out)})
        if hc.ng and rows:                         # 外れの大きい山を 1 つ目立たせる
            key = (lambda r: vbx.val(r["diff"])) if hc.kind == "over" else (lambda r: -vbx.val(r["diff"]))
            worst = max(rows, key=key)
            if worst["diff"]:
                worst["worst"] = True
        label = {"over": "許容値ｵｰﾊﾞｰ", "short": "枚数不足(?)"}.get(hc.kind, "合格")
        if not judged:
            label = "判定なし(Lot情報履歴にｵｰﾀﾞｰNOが無いため合格扱い)" if hc.show_uf1 else "判定なし(Lot情報履歴なし)"
        return {
            "judged": judged, "ok": not hc.ng, "kind": hc.kind, "label": label, "msg": hc.msg_error,
            "stage": stage, "mount": mc,
            "hei": _f2(sum(vbx.val(h) for h in heights)) if heights else "",
            "gm": _f2(hc.gm), "gp": _f2(hc.gp), "lo": _f2(lo), "hi": _f2(hi),
            "diff": _f2(hc.sap if hc.kind == "over" else hc.sam),
            "allow": _f2(hc.error_a if hc.kind == "over" else hc.error_b),
            "allow_over": _f2(hc.error_a), "allow_short": _f2(hc.error_b),
            "ais": _f2(hc.ais), "vcs": _f2(hc.vcs), "gai": _f2(hc.gai),
            "rows": rows,
        }

    # ------------------------------------------------------------------
    # そのほかのボタン
    # ------------------------------------------------------------------
    def unlock(self) -> Any:
        return self._run(lambda: self.messages.extend(ainput.unlock_gaityuu(self._form())))

    def reset(self, answers: Optional[dict[str, bool]] = None) -> Any:
        return self._run(self._reset, answers or {})

    def _reset(self, answers: dict[str, bool]) -> Optional[Question]:
        """CommandButton2「Reset」(UFAinput:127-176)。VBA は確かめずに実行する。画面では押し間違いを避けて 1 度聞く。"""
        f = self._form()
        ans = answers.get("ainput_reset")
        if ans is None:
            return Question("ainput_reset", "入力を消して hostsend.txt を書き直し、送信アプリを起動します。よろしいですか？",
                            "okcancel", source="UFAinput:127")
        if not ans:
            return None
        path, launch = ainput.clear_and_resend(f, launcher=self.launcher, sleep=self.sleep)
        trace.note_used(path, "save")
        f.order_no = ""                                    # HikiNo = "" → HikiNo_Change
        self.adjust_visible = True
        self.messages += launch.messages
        if not launch.ok:
            f.ng_flag = "stop"
        now = self.kensa.now()
        self.last = _Last(stage="reset", at=now.strftime("%H:%M:%S"), hostsend_path=path, launch_ok=launch.ok,
                          launched=launch.launched)
        self.check = None
        return None

    # ------------------------------------------------------------------
    # 画面に出す形
    # ------------------------------------------------------------------
    def _result_label(self, last: _Last) -> str:
        if last.stage == "sent":
            if last.result_missing:
                return "result.txt が無いまま成功扱い(VBA どおり)"
            return last.result_text or ainput.SUCCESS_TEXT
        if last.result_ok is False:
            return ainput.result_con(last.result_text)
        return last.note

    def _steps(self) -> list[dict[str, str]]:
        last = self.last
        states = [""] * len(STEPS)
        s = last.stage

        def upto(n: int, mark: str = "ok") -> None:
            for i in range(n):
                states[i] = mark
        if s == "input_error":
            states[0] = "ng"
        elif s == "height_ng":
            upto(1)
            states[1] = "ng"
        elif s == "question":
            upto(2)
            states[2] = "wait"
        elif s == "cancel":
            upto(2)
            states[2] = "skip"
        elif s == "stopped":
            if last.launch_ok is False or last.result_ok is None:
                upto(4)
                states[4] = "ng"
            else:
                upto(5)
                states[5] = "ng"
        elif s == "sent":
            upto(len(STEPS))
            if last.result_missing:
                states[5] = "warn"
            if last.daily_rows == 0:
                states[6] = "skip"
        if s in ("sent", "stopped") and last.launched is False and last.launch_ok:
            states[4] = "skip"                         # Windows 以外: 送信アプリは起動しない
        return [{"label": lab, "state": st} for lab, st in zip(STEPS, states)]

    def view(self) -> dict[str, Any]:
        k = self.kensa
        st = k.state
        ks = {"lot": st.lot, "line": st.lname, "minmax": all(getattr(st, a) != "" for a in
                                                            ("amin", "amax", "hmin", "hmax", "tmin", "tmax"))}
        out: dict[str, Any] = {"ready": self.form is not None, "ken_m": self.ken_m, "source": self.source, "ken_methods": list(KEN_METHODS),
                               "kensa": ks}
        f = self.form
        if f is None:
            return out
        mc = vbx.clng(vbx.val(f.mount_count))
        mc_ok = 1 <= mc <= MAX_MOUNT
        try:
            lh = ainput.l_height(f) if mc_ok else None
        except ZeroDivisionError:
            lh = None
        filled = sum(1 for i in range(mc) if mc_ok and f.heights[i] != "")
        nxt = ainput.next_empty_height_index(f, mc) if mc_ok else -1
        heights = [{"n": i, "value": f.heights[i - 1], "next": i == nxt} for i in range(1, mc + 1)] if mc_ok else []
        if not mc_ok:
            hint = ""
        elif filled == 0:
            hint = "1 山目を測定してください"
        elif filled == mc:
            hint = "全山入力済み。測り直す山は × で消してから測定"
        else:
            hint = f"{filled + 1} 山目を測定してください"
        second = f.l1_visible and ainput.all_height_filled(f, mc)
        last = self.last
        out.update({
            "busy": f.tag == "Load",
            "lot": {"line": f.line_name, "lot": f.lot, "ken_no": f.ken_no, "zai": f.zai, "tyo": f.tyo, "atu": f.atu,
                    "haba": f.haba, "take": f.take, "youto": f.youto, "ken_law": f.ken_law, "order_no": f.order_no,
                    "adjust_label": self.adjust_label, "adjust_no": self.adjust_no,
                    "adjust_visible": self.adjust_visible, "db_flag": f.db_flag == "true",
                    "amin": f.amin, "amax": f.amax, "hmin": f.hmin, "hmax": f.hmax, "tmin": f.tmin, "tmax": f.tmax,
                    "other_lot": st.lot != "" and f.lot != st.lot and f.db_flag != "true"},
            "fields": {k_: getattr(f, k_) for k_ in ("hiki_no", "konno", "remain", "pile_match_no", "pile_match",
                                                     "gaityuu_text") + NUMERIC_FIELDS},
            "hiki_list": self.hiki_list,
            "show": {"mai2": f.mai_flag == "2", "material": f.ken_law in ("M2", "M6"),
                     "material2": self.material_flag in ("2", "3"), "material3": self.material_flag == "3",
                     "ex": f.ex_visible, "ex_material2": self.ex_material_flag in ("2", "3"),
                     "ex_material3": self.ex_material_flag == "3", "gaityuu": not f.locked},
            "mai_label": "個装高さ" if f.ken_law == "M2" else "梱包枚数",
            "mount": {"count": f.mount_count, "visible": f.l1_visible, "device": f.l1_caption if f.l1_visible else "",
                      "per_mount": _f2(lh) if lh is not None else "", "heights": heights, "filled": filled,
                      "next": nxt, "hint": hint,
                      "total": _f2(ainput.calc_mount_height(f, mc)) if mc_ok and filled else "",
                      "received": {"text": self.received.text, "index": self.received.index,
                                   "reason": self.received.reason}},
            "extras": {"ai": f.ai, "ga": f.ga, "vab": f.vab, "locked": f.locked,
                       "aisi": self._caption(f.ai_text, f.ai, ai=True), "gaisousi": self._caption(f.ga_text, f.ga),
                       "vcab": self._caption(f.vcab, f.vab)},
            "check": self.check,
            "press": {"phase": "second" if second else "first",
                      "text": "2回目: 入力チェック → 梱包高さ → 転送" if second else "1回目: 山ごとの高さ入力欄を出す"},
            "send": {"stage": last.stage, "at": last.at, "steps": self._steps(), "label": self._result_label(last),
                     "note": last.note, "csv_path": last.csv_path, "hostsend_path": last.hostsend_path,
                     "result_text": last.result_text, "result_missing": last.result_missing,
                     "ng_flag": f.ng_flag, "current": ainput.host_data_confi(),
                     "codes": [{"code": c, "text": ainput.host_con(c)} for c in RESULT_CODES],
                     "sender": paths.join(ainput.tools_dir(), ainput.SENDER_DIR, ainput.SENDER_EXE),
                     "history": self.sends},
        })
        return out

    def _caption(self, text: str, checked: bool, *, ai: bool = False) -> str:
        """Aisi / Gaisousi / VCAB の表示。AiText_Change などで写し、外注で外したものは Ai_BeforeUpdate など
        (UFAinput:73-106)で「なし」にする(合紙は表示が空か 0 のときはそのまま)。"""
        f = self.form
        if f is not None and f.gaityuu_text == "外注" and checked and text != "":
            if ai and text == "0":
                return text
            return "なし"
        return text


_session: Optional[AinputSession] = None
_slock = threading.Lock()


def get() -> AinputSession:
    global _session
    with _slock:
        if _session is None:
            _session = AinputSession()
        return _session


def reset_for_test(s: Optional[AinputSession] = None) -> AinputSession:
    global _session
    with _slock:
        _session = s or AinputSession()
        return _session
