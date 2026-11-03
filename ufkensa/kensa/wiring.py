"""検査の流れの結合(工程 8c): grid.Hooks の本物の実装と、各部品の画面値 ⇔ KensaState の変換

VBA の UF1 は 1 つのフォームを全モジュールが `UF1.xxx` で直接読み書きしていた。移植では各担当が
「画面の値を dataclass で受け取り、変更を返す」純粋な関数にしたので、ここでそれを `grid.KensaState` と
つなぐ。**判定・計算はここでしない**(各部品の関数を呼ぶだけ)。

【名前で読む・書く】`get(state, name)` / `put(state, name, value)`
VBA のコントロール名(`ATU`、`S1`、`DiagSKou`、`LateralText.Tag`、`厚板.Caption` など)で KensaState を読み書きする。
各部品が返す「コントロール名 → 値」(FormRestore・ReturnItemResult.uf1・LotReflect.form・ReflectResult など)は
すべてこれで反映する。名前の対応:

    VBA の名前                         KensaState
    ---------------------------------  ------------------------------------------------
    S1〜E14・F1〜F10                    cells[名前]
    DiagS〜F / DiagSKou〜 / LateralS〜  diag / diag_kou / lateral[位置]
    NGS〜 / ZS〜 / GS〜 / CBS1〜CBF4     ng / strain_z / strain_g / strain_cb
    SBH〜EBH / SBad〜FBad               bh / bad[位置]
    KenHBox1〜7 / LC1〜12               ken_h / visual[番号]
    `SCALAR` の名前(Lot, ATU, Ap …)    同名の属性(下の表)
    `TAG_SCALAR` の名前の .Tag          属性(LateralText.Tag → lateral_text_tag など)
    厚板.Caption など                   work_caption[作業]
    それ以外(T1〜T3, Day1〜3, StartTime, EndTime, TaikText, flatnes, MG, HaHiz, TaHiz, Naka, Mimi,
    HizCome, DistortionFrame, KenH, LastHot, QF, 高さ計算式, YamaLB, dataLoding, SplitFlag, Abno,
    列見出し S〜F, LZS, LGS, S_UpVisual …)
                                        view[名前](Ctl。Caption = value、Tag = tag)

【各部品の画面値の型との変換】
    checks.Uf1           to_uf1(state)                     text = 上の名前すべて、visible、ng = 色が NG の名前、
                                                           work、line_name/line_no、worker_day_max(検査員シート)
    judge.Form           to_judge_form(state)              cells(測定・対角・直角度・ラテラル・歪・BH)の
                                                           value/state/visible/tag/border、captions、combos、公差
                         apply_judge(state, Outcome)       state "OK"/"NG"/""/"黄" → color "OK"/"NG"/"WHITE"/"YELLOW"、
                                                           border "赤"/"" → "RED"/"BLACK"、values/tags は put
    lot.LotScreen        to_lot_screen(state, facility)    lot=Lot、cb_text=Lot(UF1.CB = Lot)、dbflag
    lot.LotReflect       apply_lot_reflect(state, r)       form の名前を put、last_hot/qf、キャッシュ、指示公差
    records.KensaSnapshot to_snapshot(state, store)        測定 14 項目・歪(表示面と Tag)・対角・直角度・NG・
                                                           ラテラル・外観・BH・位置名・作業・T1〜T3 など
    records.FormRestore  apply_restore(state, r, hooks)    values/tags/back_colors を put、option → work、
                                                           calls(Atu_OPselect・LSelect…)を順に呼ぶ
    workers.WorkerRegistry registry(state) / set_registry  view の T1〜T3・Day1〜3
    checks.WorkDist      work_dist(store, lname)           作業保存_<L> の O2:P13
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Callable, Optional, Sequence

from . import (calibration, checks, datasource, grid, input_router, instruments, judge, layout, legacy_log, lines,
               lot as lotmod, paths, quality, records, report_model, vbx, workers)
from .. import trace
from .flow import ERROR, INFO, WARN, Message, Stop
from .grid import Ctl, Hooks, KensaState
from .store import LOT_HISTORY, ROW_TABLES, LocalStore, cell_text, line_sheet

#: TeColor・LcColor の水色(RGB(0,255,255))。不良枚数欄・外観ボタンの「選んでいる欄」
COLOR_CYAN = "CYAN"

# ------------------------------------------------------------------
# 名前 → KensaState
# ------------------------------------------------------------------
#: VBA のコントロール名(Caption / Text / Value)→ KensaState の属性
SCALAR: dict[str, str] = {
    "Lot": "lot", "LineName": "line_name", "LineNo": "line_no", "KenNo": "ken_no_caption", "Reien": "reien",
    "Comment": "comment", "Youto": "youto", "Zai": "zai", "Tyu": "tyu", "Sumisei": "sumisei",
    "Additional": "additional", "LateralText": "lateral_text", "flatnesText": "flatnes_text", "HURIS": "huris",
    "ConductivText": "conductiv_text", "ATU": "atu", "Hab": "hab", "Tak": "tak", "Ap": "ap", "Am": "am",
    "Hp": "hp", "Hm": "hm", "Tp": "tp", "Tm": "tm", "ApK": "apk", "AmK": "amk", "HpK": "hpk", "HmK": "hmk",
    "TpK": "tpk", "TmK": "tmk", "TAIK": "taik", "Lateral": "lateral_grade", "AMin": "amin", "AMax": "amax",
    "HMin": "hmin", "HMax": "hmax", "TMin": "tmin", "TMax": "tmax", "DBflag": "dbflag", "Yama": "yama",
    "KenCount": "ken_count", "Machin": "machin", "INP": "inp", "Tagtext": "tagtext", "PakeD": "paked",
    "DeleteP": "delete_p", "DeleteC": "delete_c", "AorB": "strain_face", "AtCount": "at_count", "Auto": "auto",
    "Load": "load", "ProofreadFlag": "proofread_flag", "EachTime": "each_time", "EachValue": "each_value",
    "MoveFlag": "move_flag", "UFTag": "form_tag",
}
#: `名前.Tag` → KensaState の属性
TAG_SCALAR: dict[str, str] = {
    "LateralText": "lateral_text_tag", "ReSize": "resize_tag", "INP": "inp_tag", "Yama": "yama_tag",
    "Abno": "abno_tag", "MPEX": "mpex_tag", "PicFM": "picfm_tag", "blindfold": "blindfold_tag", "SFM": "sfm_tag",
    "WorkButton": "work_button_tag",
}
_PROPS = {"Caption": "value", "Text": "value", "Value": "value", "Tag": "tag", "BackColor": "color",
          "Visible": "visible", "Enabled": "enabled", "BackStyle": "backstyle"}
_COLOR_BY_VBA = {v: k for k, v in grid.VBA_COLOR.items()}
_COLOR_BY_VBA[0x80000005] = grid.COLOR_DEFAULT           # records.RGB_WHITE_DEFAULT(正の値で来る)


def _split(name: str) -> tuple[str, str]:
    if "." in name:
        base, prop = name.rsplit(".", 1)
        if prop in _PROPS:
            return base, _PROPS[prop]
    return name, "value"


def _ctl(state: KensaState, base: str) -> Ctl:
    return state.ctl(base) if state.has_ctl(base) else state.v(base)


def color_name(v: Any) -> str:
    """VBA の BackColor の値 → `Ctl.color` の名前(知らない値は数の文字列)。"""
    if isinstance(v, str):
        return v
    return _COLOR_BY_VBA.get(int(v), str(int(v)))


def put(state: KensaState, name: str, value: Any) -> None:
    """VBA の `UF1(name) = value`(`名前.Tag` などのプロパティつきも可)。"""
    base, prop = _split(name)
    if prop == "value":
        text = value if isinstance(value, str) else cell_text(value)
        if base in SCALAR:
            setattr(state, SCALAR[base], text)
        elif base in grid.WORKS and name.endswith(".Caption"):
            state.work_caption[base] = text
        else:
            _ctl(state, base).value = text
    elif prop == "tag":
        text = value if isinstance(value, str) else cell_text(value)
        if base in TAG_SCALAR:
            setattr(state, TAG_SCALAR[base], text)
        else:
            _ctl(state, base).tag = text
    elif prop == "color":
        if base == "LateralText":
            state.lateral_text_colored = color_name(value) not in (grid.COLOR_DEFAULT, "-2147483642")
        _ctl(state, base).color = color_name(value)
    elif prop == "visible":
        _ctl(state, base).visible = bool(value)
    elif prop == "enabled":
        _ctl(state, base).enabled = bool(value)
    # BackStyle は表示だけ(移さない)


def get(state: KensaState, name: str) -> str:
    """VBA の `UF1(name)`(Caption / Text)。`名前.Tag` も読める。"""
    base, prop = _split(name)
    if prop == "tag":
        if base in TAG_SCALAR:
            return getattr(state, TAG_SCALAR[base])
        return _ctl(state, base).tag
    if base in SCALAR:
        return getattr(state, SCALAR[base])
    if base in grid.WORKS and name.endswith(".Caption"):
        return state.work_caption.get(base, base)
    return _ctl(state, base).value


def all_controls(state: KensaState) -> dict[str, Ctl]:
    """名前つきのコントロールすべて(グリッド・付随欄・表示だけのもの)。"""
    out: dict[str, Ctl] = dict(state.cells)
    for p, c in state.diag.items():
        out["Diag" + p] = c
    for p, c in state.diag_kou.items():
        out[f"Diag{p}Kou"] = c
    for prefix, group in (("Lateral", state.lateral), ("NG", state.ng), ("Z", state.strain_z), ("G", state.strain_g)):
        for p, c in group.items():
            out[prefix + p] = c
    out.update(state.strain_cb)
    for p, c in state.bh.items():
        out[p + "BH"] = c
    for p, c in state.bad.items():
        out[p + "Bad"] = c
    for i, c in state.ken_h.items():
        out[f"KenHBox{i}"] = c
    for i, c in state.visual.items():
        out[f"LC{i}"] = c
    for k, c in state.view.items():
        out.setdefault(k, c)
    return out


# ------------------------------------------------------------------
# 手元DB の小道具
# ------------------------------------------------------------------
def sheet_cells(store: LocalStore, name: str) -> Optional[dict[tuple[int, int], Any]]:
    """シートのセル `{(行, 列): 値}`(シートが無ければ None)。"""
    if not store.exists(name):
        return None
    return {(r, c): v for r, c, v in store.sheet(name).used()}


def sheet_rows(store: LocalStore, name: str) -> Optional[list[list[Any]]]:
    """列の決まった表を「1 行目(見出し)〜A 列の最終行」の形で(VBA の `sh.Range(A1, …)`)。"""
    if not store.exists(name):
        return None
    sh = store.sheet(name)
    headers = list(ROW_TABLES[name]) if name in ROW_TABLES else []
    last = sh.last_row(1)
    width = len(headers) or max((c for _, c, _ in sh.used()), default=1)
    rows = [headers or [sh.get(1, c) for c in range(1, width + 1)]]
    for r in range(2, last + 1):
        rows.append([sh.get(r, c) for c in range(1, width + 1)])
    return rows


def write_cells(store: LocalStore, name: str, writes: Sequence[workers.CellWrite], *, recreate: bool = False) -> None:
    sh = store.add_sheet(name) if recreate else store.ensure_sheet(name)
    for w in writes:
        sh.write(w.row, w.col, w.value)


def work_dist(store: LocalStore, lname: str) -> Optional[checks.WorkDist]:
    """作業保存_<L> の O2:P7(A 面)・O8:P13(B 面)。"""
    name = line_sheet(records.WORK_SAVE, lname)
    if not store.exists(name):
        return None
    sh = store.sheet(name)
    return checks.WorkDist(sh.read_range(2, 15, 7, 16), sh.read_range(8, 15, 13, 16))


def worker_sheet_a(store: LocalStore, lname: str) -> Optional[tuple[Any, Any, Any]]:
    """検査員_<L> の A1:A3(シートが無ければ None)。"""
    cells = sheet_cells(store, line_sheet("検査員", lname))
    if cells is None:
        return None
    return cells.get((1, 1)), cells.get((2, 1)), cells.get((3, 1))


def worker_day_max(store: LocalStore, lname: str) -> str:
    a = worker_sheet_a(store, lname)
    return workers.worke_day_max(*a) if a else ""


def registry(state: KensaState) -> workers.WorkerRegistry:
    """UF1 の T1〜T3・Day1〜Day3。"""
    return workers.WorkerRegistry(t1=state.v("T1").value, t2=state.v("T2").value, t3=state.v("T3").value,
                                  day1=state.v("Day1").value, day2=state.v("Day2").value,
                                  day3=state.v("Day3").value)


def set_registry(state: KensaState, reg: workers.WorkerRegistry) -> None:
    for name, v in (("T1", reg.t1), ("T2", reg.t2), ("T3", reg.t3),
                    ("Day1", reg.day1), ("Day2", reg.day2), ("Day3", reg.day3)):
        state.v(name).value = v


# ------------------------------------------------------------------
# checks.Uf1
# ------------------------------------------------------------------
def to_uf1(state: KensaState, worker_day_max_: str = "") -> checks.Uf1:
    """KensaState → checks.Uf1(読むだけの写し)。"""
    text: dict[str, str] = {}
    visible: dict[str, bool] = {}
    ng: set[str] = set()
    for name, c in all_controls(state).items():
        text[name] = c.value
        visible[name] = c.visible
        if c.color == grid.COLOR_NG:
            ng.add(name)
    for name, attr in SCALAR.items():
        text[name] = getattr(state, attr)
    for name, attr in TAG_SCALAR.items():
        text[name + ".Tag"] = getattr(state, attr)
    for w in grid.WORKS:
        text[w + ".Caption"] = state.work_caption.get(w, w)
    visible["Yama"] = state.yama_visible
    visible["AtButt"] = state.at_butt_visible
    visible["InspeButt"] = state.inspe_butt_visible
    text["AtButt"] = state.at_butt_caption
    text["InspeButt"] = state.inspe_butt_caption
    return checks.Uf1(text=text, visible=visible, ng=ng, work=state.work, line_name=state.line_name,
                      line_no=state.line_no, worker_day_max=worker_day_max_)


# ------------------------------------------------------------------
# judge.Form
# ------------------------------------------------------------------
_TO_JUDGE_STATE = {grid.COLOR_OK: judge.OK, grid.COLOR_NG: judge.NG, grid.COLOR_YELLOW: judge.YELLOW}
_FROM_JUDGE_STATE = {judge.OK: grid.COLOR_OK, judge.NG: grid.COLOR_NG, judge.YELLOW: grid.COLOR_YELLOW,
                     judge.WHITE: grid.COLOR_WHITE}


def to_judge_form(state: KensaState) -> judge.Form:
    """KensaState → judge.Form。"""
    f = judge.Form()
    for name, cell in f.cells.items():
        c = state.ctl(name)
        cell.value, cell.visible, cell.tag = c.value, c.visible, c.tag
        cell.state = _TO_JUDGE_STATE.get(c.color, judge.WHITE)
        cell.border = judge.RED if c.border == grid.BORDER_RED else ""
    for p in grid.POSITIONS:
        f.captions[p] = state.v(p).value
    f.captions["LZS"] = state.v("LZS").value
    f.captions["LGS"] = state.v("LGS").value
    for k in f.combos:
        f.combos[k] = state.strain_cb[k].value
    f.ApK, f.AmK, f.HpK, f.HmK, f.TpK, f.TmK = state.apk, state.amk, state.hpk, state.hmk, state.tpk, state.tmk
    f.Am, f.Hm, f.Tm = state.am, state.hm, state.tm
    f.TaikText = state.v("TaikText").value
    f.LateralText = state.lateral_text
    f.gokuatsu = state.is_work("極厚")
    f.HaHiz, f.TaHiz = state.v("HaHiz").value, state.v("TaHiz").value
    f.Naka, f.Mimi = state.v("Naka").value, state.v("Mimi").value
    f.HURIS, f.flatnesText, f.MoveFlag = state.huris, state.flatnes_text, state.move_flag
    f.Zai, f.Tyu, f.ATU, f.Hab, f.Tak, f.Lot = state.zai, state.tyu, state.atu, state.hab, state.tak, state.lot
    f.DBflag, f.LineName, f.LineNo, f.AbnoTag = state.dbflag, state.line_name, state.line_no, state.abno_tag
    return f


# ------------------------------------------------------------------
# 1 回の操作のあいだ集めるもの
# ------------------------------------------------------------------
@dataclass
class Env:
    """部品を呼ぶのに要る、画面(UF1)以外のもの。"""
    store: LocalStore
    now: Callable[[], datetime] = datetime.now
    facility: Callable[[], str] = lambda: ""
    calibration_skip: Callable[[], bool] = lambda: False
    shifts: Optional[workers.ShiftTimes] = None
    write_logs: bool = True
    messages: list[Message] = field(default_factory=list)
    confi: list[checks.ConfiRequest] = field(default_factory=list)
    reason: Optional[checks.ConfiRequest] = None       # 消去理由の入力(ConfiReason_Open)
    proofread: Optional[calibration.ProofreadLoad] = None
    #: CBAdd が最後に入れた歪コンボの選択肢(CB?1 = 定盤の形状、CB?3 = ゲージの形状)。画面の候補
    cb_items: dict[str, list[str]] = field(default_factory=lambda: {"shape": checks.cb_shape_items(),
                                                                    "gauge": checks.cbg_shape_items({})})
    #: VC_Picture / VC_Clear の結果(None = VC_Clear 後・未表示)
    vc: Optional[lotmod.VcView] = None

    def shift_times(self) -> workers.ShiftTimes:
        if self.shifts is None:
            try:
                self.shifts = workers.load_shift_times()
            except (datasource.DbMissing, IndexError):
                self.shifts = workers.ShiftTimes()
        return self.shifts

    def log(self, entry: legacy_log.LogEntry) -> None:
        if self.write_logs:
            entry.write(now=self.now())

    def sound(self, wav: str, source: str = "") -> None:
        self.messages.append(Message("", INFO, sound=wav, source=source))


def _msg_error(env: Env, state: KensaState, text: str, n: int, temp: str = "") -> None:
    """OutTolerance.MsgError(text, Lname, n, temp)。表示とログ。"END" は Stop(D10)。"""
    reg = registry(state)
    out = checks.msg_error(text, state.lname, n, temp, (reg.t1, reg.t2, reg.t3))
    env.messages.append(out.message)
    env.log(legacy_log.line_entry(state.line_name, out.log_text))
    if out.end:
        raise Stop(out.message.text, messages=list(env.messages), source="OutTolerance.MsgError")


def apply_judge(state: KensaState, out: judge.Outcome, env: Env) -> None:
    """judge.Outcome を KensaState に反映する(色・値・Tag・枠、音、ログ、MsgError、NGErase)。"""
    for k, s in out.states.items():
        state.ctl(k).color = _FROM_JUDGE_STATE.get(s, grid.COLOR_WHITE)
    for k, v in out.values.items():
        put(state, k, v)
    for k, v in out.tags.items():
        state.ctl(k).tag = v
    for k, b in out.borders.items():
        state.ctl(k).border = grid.BORDER_RED if b == judge.RED else grid.BORDER_BLACK
    if out.move_flag is not None:
        state.move_flag = out.move_flag
    if out.mpex_tag is not None:
        state.mpex_tag = out.mpex_tag
    if state.form_tag != "STOP":                      # UF1.Tag="STOP" の間は音を出さない
        for s in out.sounds:
            env.sound(s, "Judg")
    env.messages.extend(out.messages)
    for lg in out.logs:
        if env.write_logs:
            legacy_log.write_log(lg.text, lg.sepl, lg.target_path, lg.target_file, now=env.now())
    for m in out.msg_errors:
        _msg_error(env, state, m.text, m.icon, m.temp)
    if out.ng_erase is not None:
        input_router.ng_erase(state, out.ng_erase)


# ------------------------------------------------------------------
# lot.LotScreen / LotReflect
# ------------------------------------------------------------------
def to_lot_screen(state: KensaState, facility: str) -> lotmod.LotScreen:
    """`UF1.CB = Lot`(Lot_Change)なので cb_text は Lot。"""
    return lotmod.LotScreen(lot=state.lot, cb_text=state.lot, dbflag=state.dbflag,
                            facility=facility or workers.DEFAULT_FACILITY)


def apply_lot_reflect(state: KensaState, r: lotmod.LotReflect) -> None:
    """AccReflec の結果を画面へ(SaveLoad:180-345 / SkipNoItti)。"""
    for k, v in r.form.items():
        put(state, k, v)
    if r.status == "not_found":
        for k in ("MG", "HaHiz", "TaHiz", "Naka", "Mimi"):
            state.v(k).value = ""
        return
    if r.status != "ok":
        return
    if r.last_hot is not None:
        state.v("LastHot").value = r.last_hot
    if r.qf is not None:
        state.v("QF").value = r.qf
    state.sika_set_atu = r.sika_set_atu
    if r.taik_text is not None:
        state.v("TaikText").value = r.taik_text
    if r.lateral_text is not None:
        state.lateral_text = r.lateral_text
        state.lateral_text_tag = "有"
        state.lateral_text_colored = True
    state.additional = r.additional
    if r.cache_ord_atu is not None:
        state.cache.ord_atu = r.cache_ord_atu
    if r.cache_c_course is not None:
        state.cache.c_course = r.cache_c_course


def _tolerance_text(state: KensaState, r: lotmod.LotReflect) -> None:
    """AccReflec の対角・ラテラルボー公差(指示が無いとき。SaveLoad:283-327)。"""
    def f1(x: float) -> str:
        return vbx.fmt(x, "0.0")
    hab, tak, atu = vbx.single(vbx.val(state.hab)), vbx.single(vbx.val(state.tak)), vbx.single(vbx.val(state.atu))
    if r.taik_text is None:
        if state.taik == "4":
            state.v("TaikText").value = f1(judge.diagonal(hab, tak))
        elif state.taik == "3":
            state.v("TaikText").value = f1(judge.re_diagonal(hab, tak))
    if r.lateral_text is None and state.lateral_grade != "":
        txt = None
        if state.lateral_grade == "4":
            txt = f1(judge.lateralbow_j(hab, tak, atu))
        elif state.lateral_grade == "3":
            txt = f1(judge.re_lateralbow(hab, tak, atu))
        if txt is not None:
            state.lateral_text = "" if vbx.u1_compare(txt, 0) == 0 else txt


# ------------------------------------------------------------------
# records.KensaSnapshot / FormRestore
# ------------------------------------------------------------------
def to_snapshot(state: KensaState, store: LocalStore) -> records.KensaSnapshot:
    """KensaState → records.KensaSnapshot(保存・帳票・作業中保存に使う UF1 の値)。"""
    v = lambda n: state.v(n).value          # noqa: E731
    s = records.KensaSnapshot()
    s.line_name = state.lname
    s.lot = state.lot
    for attr in ("atu", "hab", "tak", "am", "ap", "hm", "hp", "tm", "tp", "amk", "apk", "hmk", "hpk", "tmk", "tpk",
                 "zai", "tyu", "youto", "taik", "machin", "resize_tag", "sumisei", "additional", "lateral_text",
                 "huris", "ken_count", "lateral_text_tag", "delete_p", "delete_c", "at_count", "flatnes_text",
                 "at_butt_visible", "inspe_butt_visible", "inspe_butt_caption", "yama", "yama_visible"):
        setattr(s, attr, getattr(state, attr))
    s.lateral = state.lateral_grade
    s.taik_text = v("TaikText")
    s.flatnes = v("flatnes")
    for f in grid.POSITIONS:
        rows = grid.ROW_COUNT[f]
        s.measure[f] = [state.cells[f"{f}{i}"].value if i <= rows else "" for i in range(1, 15)]
        cb = state.strain_cb
        s.strain[f] = records.Strain(cb[f"CB{f}1"].value, cb[f"CB{f}2"].value, state.strain_z[f].value,
                                     cb[f"CB{f}3"].value, cb[f"CB{f}4"].value, state.strain_g[f].value)
        s.strain_tag[f] = records.Strain(cb[f"CB{f}1"].tag, cb[f"CB{f}2"].tag, state.strain_z[f].tag,
                                         cb[f"CB{f}3"].tag, cb[f"CB{f}4"].tag, state.strain_g[f].tag)
        s.diag[f] = state.diag[f].value
        s.diag_kou[f] = state.diag_kou[f].value
        s.ng[f] = state.ng[f].value
        s.lateral_face[f] = state.lateral[f].value
        s.visual_up[f] = v(f + "_UpVisual")
        s.visual_under[f] = v(f + "_UnderVisual")
        s.bh[f] = state.bh[f].value if f in state.bh else ""
        s.positions[f] = state.header(f)
    s.aorb = state.strain_face
    s.kind = state.work
    s.atsuita_caption = state.work_caption.get("厚板", "厚板")
    s.usuita_caption = state.work_caption.get("薄板", "薄板")
    s.start_time, s.end_time = v("StartTime"), v("EndTime")
    s.t1, s.t2, s.t3 = v("T1"), v("T2"), v("T3")
    s.worker_shift = worker_day_max(store, state.lname)
    s.data_loading = v("dataLoding")
    s.lgs_caption, s.lzs_caption = v("LGS"), v("LZS")
    s.mg, s.hahiz, s.tahiz, s.naka, s.mimi = v("MG"), v("HaHiz"), v("TaHiz"), v("Naka"), v("Mimi")
    s.distortion_frame, s.hiz_come = v("DistortionFrame"), v("HizCome")
    s.ken_hbox = [state.ken_h[i].value for i in range(1, 8)]
    s.yama_lb, s.height_formula = v("YamaLB"), v("高さ計算式")
    s.ken_no, s.ken_h = state.ken_no_caption, v("KenH")
    s.last_hot, s.qf = v("LastHot"), v("QF")
    s.lateral_text_backcolor = quality.rgb(51, 204, 204) if state.lateral_text_colored else 0
    s.lateral_text_backstyle = 1 if state.lateral_text_colored else 0
    s.lot_source_lot = state.lot_source_lot if state.lot_source == LOT_SECOND else ""     # D28
    return s


#: D28: 「控え(第2)から読む」(state.lot_source の値)
LOT_SECOND = "second"


def reads_second(state: Any, lot: str) -> bool:
    """D28: このロットを控え(第2)だけから読むか(検査画面でそのロットに「控えから読む」を選んでいるとき)。"""
    if state is None or lot is None:
        return False
    lot = str(lot).strip().upper()
    return (getattr(state, "lot_source", "") == LOT_SECOND and lot != ""
            and getattr(state, "lot_source_lot", "") == lot)


def restore_lot_source(state: KensaState, lot: str) -> None:
    """D28: 作業保存から読んだ「控えから読む」のロットを状態へ戻す("" なら いつもの読み方)。"""
    if lot:
        state.lot_source, state.lot_source_lot = LOT_SECOND, lot
    else:
        state.lot_source, state.lot_source_lot = "", ""


def apply_restore(state: KensaState, r: records.FormRestore, hooks: "RealHooks") -> None:
    """records.FormRestore(Load_Text・LosdStrain の結果)を画面へ戻し、続けて呼ぶ手続きを呼ぶ。"""
    for k, val in r.values.items():
        put(state, k, val)
    for k, val in r.tags.items():
        put(state, k + ".Tag", val)
    for k, col in r.back_colors.items():
        put(state, k + ".BackColor", col)
    if "YamaLB" in r.list_index:
        state.v("YamaLB").tag = str(r.list_index["YamaLB"])
    hooks.env.messages.extend(r.messages)
    # 作業選択は Atu_OPselect(全部 False にする)の後で `UF1.厚板 = True` など(SaveLoad:2839-2867)。
    # OptionButton を True にすると Click イベント(厚板_Click → LSelect …)が走る。シートに記録が無ければ選ばない
    ws = line_sheet(records.WORK_SAVE, state.lname)
    at_sl = cell_text(hooks.store.sheet(ws).get(1, 29)) if hooks.store.exists(ws) else ""
    for call in r.calls:
        hooks.restore_call(state, call)
        if call[0] == "Atu_OPselect" and at_sl in grid.WORKS and r.option == at_sl:
            try:
                layout.select_work(state, at_sl, hooks, hooks.env.messages)
            except layout.VBA_ERRORS:
                state.work = at_sl


# ------------------------------------------------------------------
# LoopText の中身(FormRelat:278-302)
# ------------------------------------------------------------------
def time_posting(state: KensaState, env: Env, position: str, each_t: Optional[datetime], temp: str = "1") -> None:
    """SaveLoad.Time_Posting (SaveLoad:2288-2358)。WorkerTime_<L> に測定者・時刻・種別を書く(temp="" で消す)。"""
    if position == "":
        return
    store = env.store
    name = line_sheet(records.WORKER_TIME, state.lname)
    if not store.exists(name):
        store.add_sheet(name)
    if state.v("SplitFlag").value == "":
        if not records.variant_eq(store.sheet(name).get(1, 1), state.lot):
            store.add_sheet(name)
    sh = store.sheet(name)
    sh.write(1, 1, state.lot)
    sh.write(1, 2, records.vb_format_date(state.v("StartTime").value, "yyyy/mm/dd hh:mm:ss"), "S")
    if state.is_work("極厚"):
        fields = ["厚1", "厚2", "厚3", "厚4", "幅1", "幅2", "幅3", "幅4", "幅5", "幅6", "丈1", "丈2"]
    else:
        fields = ["厚1", "厚2", "厚3", "厚4", "幅1", "幅2", "丈1", "丈2", "対角A", "対角B"]
    sh.write_row(2, 1, fields)
    i = {"S": 1, "U": 2, "C": 3, "D": 4, "E": 5, "F": 6}.get(vbx.left(position, 1), 0)
    x = vbx.clng(vbx.right(position, 1)) if len(position) == 2 else vbx.clng(vbx.mid(position, 2, 2))
    reg = registry(state)
    if temp == "1":
        dm = worker_day_max(env.store, state.lname)
        t = {"1": reg.t1, "2": reg.t2, "3": reg.t3}.get(dm)
        if t is not None:
            if t != "":
                sh.write(i + 2, x, t.split("_")[-1])
            sh.write(i + 8, x, records.vb_cstr_date(each_t) if each_t else "", "S")
            sh.write(i + 14, x, state.paked)
    else:
        sh.set(i + 2, x, None)
        sh.set(i + 8, x, None)
        sh.set(i + 14, x, None)


def daily_context(state: KensaState, env: Env) -> workers.DailyContext:
    name = line_sheet(records.DAILY, state.lname)
    cells = sheet_cells(env.store, name)
    hist = sheet_rows(env.store, LOT_HISTORY)
    return workers.DailyContext(exists=cells is not None, cells=dict(cells or {}),
                                history_rows=hist[1:] if hist is not None else None, working_lot=state.lot,
                                now=env.now(), facility=env.facility() or workers.DEFAULT_FACILITY,
                                shifts=env.shift_times(), registry=registry(state), uf1_line=state.line_name)


def daily_report(state: KensaState, env: Env, n: int) -> None:
    """Worker.Daily_Report(n)(日報 Daily_<L> に書く)。"""
    ctx = daily_context(state, env)
    writes = workers.daily_report(n, ctx)
    if writes:
        write_cells(env.store, line_sheet(records.DAILY, state.lname), writes)


def time_record(state: KensaState, env: Env, st: str) -> None:
    """FormRelat.TimeRecord (FormRelat:303-345)。測定時刻の記録と 2 度押し(4 秒以内に同じ値)の消去。"""
    if st == "" or state.abno_tag == "NG":
        return
    now = env.now().replace(microsecond=0)
    now_t = records.vb_format_date(now, "yyyy/mm/dd hh:mm:ss")
    if st == "S1":
        state.v("StartTime").value = now_t
    if env.store.exists(line_sheet(records.DAILY, state.lname)):
        daily_report(state, env, 2)
    value = vbx.cstr(vbx.val(input_router.conv(state.inp)))
    if state.each_time != "" and state.each_value != "":
        same = state.ctl(st).value == state.each_value
        try:
            secs = (now - records.vb_cdate(state.each_time)).total_seconds()
        except (vbx.TypeMismatch, ValueError):
            secs = 99.0
        if same and secs <= 4:
            state.flag_label = 0
            state.ctl(st).value = ""
            state.flag_label = 1
            state.ctl(st).color = grid.COLOR_DEFAULT
            state.each_time, state.each_value = now_t, value
            env.messages.append(Message("2度押し", WARN, source="FormRelat:319"))
            time_posting(state, env, st, now, "")
        else:
            state.each_time, state.each_value = now_t, value
            time_posting(state, env, st, now)
    elif state.each_time == "" and state.each_value == "":
        state.each_time, state.each_value = now_t, value
        time_posting(state, env, st, now)


# ------------------------------------------------------------------
# grid.Hooks の本物
# ------------------------------------------------------------------
class RealHooks(Hooks):
    """grid.Hooks を移植済みの部品で実装したもの。知らせ・音は `env.messages` に積む。"""

    def __init__(self, env: Env) -> None:
        self.env = env

    # --- 小道具 -------------------------------------------------------
    @property
    def store(self) -> LocalStore:
        return self.env.store

    def _over12(self, state: KensaState) -> bool:
        try:
            return self.is_over_12_hours(state)
        except (vbx.TypeMismatch, ValueError):
            return False

    def worker_name(self, state: KensaState) -> str:
        """Worker.GetWorkerName(判定ログの検査員)。登録が無ければ ""。"""
        try:
            return workers.get_worker_name(worker_day_max(self.store, state.lname), registry(state))
        except IndexError:
            return ""

    def _arrbox_fn(self, state: KensaState) -> Callable[..., Any]:
        def f(target: str, sql_filter: str = "", use: str = "Record", path: str = "") -> Any:
            return self.arrbox(state, target, sql_filter, use)
        return f

    # --- 値を返すもの -------------------------------------------------
    def calculate_sheets(self, state: KensaState, atu: float) -> float:
        r = checks.calculate_sheets([state.ken_h[i].value for i in range(1, 7)], atu)
        if r.formula is not None:
            state.v("高さ計算式").value = r.formula
        return r.quantity

    def register_check(self, state: KensaState) -> bool:
        now = self.env.now()
        return workers.register_check(workers.hhmmss(now), registry(state), worker_sheet_a(self.store, state.lname),
                                      now=now, facility=self.env.facility() or workers.DEFAULT_FACILITY,
                                      shifts=self.env.shift_times())

    def is_over_12_hours(self, state: KensaState) -> bool:
        cells = sheet_cells(self.store, line_sheet(records.WORKER_TIME, state.lname))
        return workers.is_over_12_hours((cells or {}).get((1, 3)), self.env.now(), sheet_exists=cells is not None)

    def get_lot_info(self, state: KensaState, lot: str) -> Optional[list[list[str]]]:
        return lotmod.get_lot_info(sheet_rows(self.store, LOT_HISTORY), lot)

    def db_return_item(self, state: KensaState, delivery: str, customer: str, youto: str, zai: str, tyu: str,
                       atu: float, hab: float, tak: float) -> str:
        lateral_lot = quality.LotData(self.arrbox(state, "SIKALOTNOW.accdb", "板", "Fields"),
                                      self.arrbox(state, "SIKALOTNOW.accdb", state.lot, "Record"))
        r = quality.db_return_item(delivery, customer, youto, zai, tyu, atu, hab, tak,
                                   additional_caption=state.additional, lateral_lot=lateral_lot,
                                   lateral_quiet=True)
        for k, v in r.uf1.items():
            put(state, k, v)
        if r.cb_add_atu is not None:
            self.cb_add(state, r.cb_add_atu)
        self.env.messages.extend(r.messages)
        return r.items

    def is_gauge_required(self, atu: float, hab: float, tak: float, zai: str, tyu: str, cours: str) -> bool:
        return quality.is_gauge_required(atu, hab, tak, zai, tyu, cours, self.env.messages)

    def sika_atu_create(self, state: KensaState, lot: str) -> Optional[list[Any]]:
        return lotmod.sika_atu_create(lot)

    def arrbox(self, state: KensaState, target: str, sql_filter: str = "", use: str = "Record") -> Optional[list[Any]]:
        name = target[:-6] if target.endswith(".accdb") else target
        r = lotmod.arrbox(name, sql_filter, use, state.dbflag, now=self.env.now(),
                          over12=self.is_over_12_hours(state) if self.store.exists(
                              line_sheet(records.WORKER_TIME, state.lname)) else False)
        self.env.messages.extend(r.messages)
        return r.value

    def atu_hiki_records(self, state: KensaState, lot: str) -> tuple[bool, Optional[list[Any]]]:
        """Atu_OPselect の SIKAHIKINOW (FormRelat:6227-6247)。切替の条件は Arrbox と同じ(lot.arrbox を参照)。

        ログ(TimeDesignation / FileDesignation の WriteLog)は arrbox と同じく書かない。
        """
        target = "SIKAHIKINOW"
        main, nxt = paths.get("main"), paths.get("next")
        fpath = main if state.dbflag == "" else nxt
        file_path = lotmod.db_path(fpath, target)
        if fpath == main:
            # D26 改: 時刻窓・PLEASE WAIT は見ない(ReadCheck = 手元へ写せたか だけ)
            if lotmod.read_check(file_path, 1) is False:
                file_path = lotmod.db_path(nxt, target)
                if lotmod.read_check(file_path, 1) is False:
                    return False, None
        flt = "[ﾛｯﾄ番号]='" + lot + "'"
        rows = datasource.get_records(file_path, lotmod.TABLE, flt)
        if rows is None and file_path == lotmod.db_path(main, target):
            # D26(ロットごとの第2): 第1の SIKAHIKI にロットが無ければ第2(控え)を見る
            rows, _second = lotmod.lot_from_second(target, flt, file_path)
        return True, rows

    # --- 値を返さないもの ---------------------------------------------
    def loop_text(self, state: KensaState, st: str) -> None:
        """FormRelat.LoopText (FormRelat:278-302)。異常値除去 → 機器 → 判定 → 時刻 → 桁 → 入力済みの印。"""
        if st == "":
            return
        env = self.env
        hit = checks.abnormal_check(to_uf1(state), st)                 # 異常値除去 (O:1321-1415)
        if hit is not None:
            state.abno_tag = "NG"                                       # AbnoFormStart (O:796-805)
            ws = line_sheet(records.WORK_SAVE, state.lname)
            row1 = self.store.sheet(ws).read_range(1, 2, 1, 10)[0] if self.store.exists(ws) else [""] * 9
            form = checks.abnormal_form_text(row1)
            env.messages.append(Message(hit.value + "\r\n" + "\r\n".join(form.values()), WARN, title="異常",
                                        source="OutTolerance:796"))
            act = checks.ng_back_tb(hit.name)                           # UF異常値発見 の唯一のボタン
            self.apply_erase(state, act)
        state.machin = instruments.machin_text(state.inp, state.machin)
        mj = instruments.micro_juge(instruments.MicroInput(st=st, machin=state.machin, atu=state.atu, hab=state.hab,
                                                           youto=state.youto, lname=state.lname))
        if mj.stop is not None:
            if mj.clear:
                state.ctl(mj.clear).value = ""
            _msg_error(env, state, mj.stop.text, 1)
            raise Stop(mj.stop.text, messages=list(env.messages), source=mj.stop.source, level=WARN)
        apply_judge(state, judge.text_judg(to_judge_form(state), st, self.worker_name(state)), env)
        time_record(state, env, st)
        state.abno_tag = ""
        layout.num_dig(state)
        for p in grid.POSITIONS:
            self.check_labels_by_category(state, p)

    def apply_erase(self, state: KensaState, act: checks.EraseAction) -> None:
        """checks.EraseAction(異常値・再測定で値を消す指示)を反映する。"""
        for name in act.clear:
            state.ctl(name).value = ""
        for name, col in act.back_color:
            state.ctl(name).color = color_name(col)
        if act.clear_inp:
            state.inp = ""
        for name in act.erase_count:
            input_router.erase_count(state, name, self)
        if act.posting:
            self.posting(state)
        if act.rejudge:
            self.loop_judg(state)
            self.hiz_loop_judg(state, 0)
        if act.look_blank:
            input_router.look_blank(state, *act.look_blank)

    def posting(self, state: KensaState) -> None:
        """SaveLoad.Posting。直角度の要否(NeedRight)は Lot情報履歴 の引当で判定する(SaveLoad:2117-2125)。"""
        if state.lname == "":
            return
        snap = to_snapshot(state, self.store)
        need_right: Optional[bool] = None
        rows = self.get_lot_info(state, state.lot) if state.lot != "" else None
        if rows is not None:
            try:
                need_right = layout.judge_lot_allocations(state, rows, self).need_right
            except layout.VBA_ERRORS:
                need_right = None
        sheet = line_sheet(records.WORK_SAVE, state.lname)
        trace.note_used(self.store.path, "save")
        with trace.timed("db.save", "作業中保存(Posting)", expected=(Stop,), path=str(self.store.path),
                         table=sheet) as info:
            sh = records.posting(self.store, snap, need_right)
            info["cells"] = len(sh.used())

    def out_tolerance(self, state: KensaState) -> None:
        """OutTolerance.Out_Tolerance。UFConfi に出すもの(最後の 1 件が画面に残る)を env.confi に積む。"""
        reqs = checks.out_tolerance(to_uf1(state, worker_day_max(self.store, state.lname)))
        if reqs:
            self.env.confi = reqs

    def loop_judg(self, state: KensaState, n: float = 0) -> None:
        bh = quality.bh_search(state.zai, state.tyu) if state.zai != "" else 0
        try:
            out = judge.loop_judg(to_judge_form(state), int(n), bh_lower=bh)
        except judge.LoopJudgAborted as e:                  # VBA はそこまで塗った色が残って止まる
            apply_judge(state, e.outcome, self.env)
            raise Stop("実行時エラー '13': 型が一致しません。", messages=list(self.env.messages), source="Judg.LoopJudg")
        apply_judge(state, out, self.env)

    def hiz_loop_judg(self, state: KensaState, n: float = 0) -> None:
        apply_judge(state, judge.hiz_loop_judg(to_judge_form(state), int(n)), self.env)

    def check_labels_by_category(self, state: KensaState, pos: str) -> None:
        col = checks.check_labels_by_category(to_uf1(state), pos, work_dist(self.store, state.lname))
        if col is not None:
            state.v("CheckT" + pos).color = color_name(col)

    def proofread_load(self, state: KensaState) -> None:
        a = worker_sheet_a(self.store, state.lname) or (None, None, None)
        pl = calibration.proofread_load(registry(state), a, state.line_name, blocked=True)
        if pl is not None:
            self.env.proofread = pl
            self.env.messages.append(pl.message)

    def time_posting(self, state: KensaState, st: str, mode: Optional[str] = None) -> None:
        time_posting(state, self.env, st, self.env.now(), "1" if mode is None else mode)

    def confi_reason_open(self, state: KensaState, pos: str, num: int, lname: str, value: str, count: float) -> None:
        self.env.reason = checks.confi_reason_open(to_uf1(state, worker_day_max(self.store, lname)), pos, num, lname,
                                                   value, vbx.clng(count))

    def te_color(self, state: KensaState, name: str) -> None:
        """OutTolerance.TeColor (OutTolerance:2451-2469)。不良枚数欄を全部白にして、押した欄を水色にする。

        色そのものは表示だけだが、「反映」(ListReflect → TecReturn)は水色の欄を書く先にするので状態として持つ。
        """
        for c in state.bad.values():
            c.color = grid.COLOR_WHITE
        state.ctl(name).color = COLOR_CYAN

    def cb_add(self, state: KensaState, sika_set_atu: float) -> None:
        """OutTolerance.CBAdd (OutTolerance:1427-1485)。歪コンボ CB?1(定盤の形状)・CB?3(ゲージの形状)の選択肢。

        値は変えない(VBA の `.Clear` は一覧だけ)。ゲージ側はその時点の LG? の Caption で決まるので、ここで控える。
        """
        self.env.cb_items = {"shape": checks.cb_shape_items(sika_set_atu),
                             "gauge": checks.cbg_shape_items({"LG" + p: state.v("LG" + p).value
                                                              for p in grid.POSITIONS})}

    def draw_list_box_content(self, state: KensaState) -> None:
        """OutTolerance.DrawListBoxContent (OutTolerance:1747-1840)。検入高さの山数(YamaLB)で入力欄を出し分け、
        CalculateSheets で枚数と式を出し直す。YamaLB が空なら何もしない。"""
        lay = checks.yama_layout(state.v("YamaLB").value)
        if lay is None:
            return
        for k, vis in lay.visible.items():
            if k.startswith("KenHBox"):
                state.ken_h[int(k[7:])].visible = vis
            else:
                state.v(k).visible = vis
        for k, cap in lay.hl_captions.items():
            state.v(k).value = cap
        if lay.formula_reset:
            state.v("高さ計算式").value = ""
        self.calculate_sheets(state, vbx.single(vbx.val(state.atu)))

    def vc_picture(self, state: KensaState, order_no: str = "") -> None:
        """SaveLoad.VC_Picture (SaveLoad:1554-1758)。VC 画像と VC厚 の系統・備考。受注番号を渡すと MaterialsUpperPart も。"""
        hist = sheet_rows(self.store, LOT_HISTORY) or []
        mup = None
        if order_no != "":
            mup = lambda o, b: quality.materials_upper_part(o, hist or None, b)   # noqa: E731
        self.env.vc = lotmod.vc_picture(hist, state.lot, order_no, test=state.line_option == "Test",
                                        materials_upper_part=mup)
        self.env.messages.extend(self.env.vc.messages)

    def vc_clear(self, state: KensaState) -> None:
        """FormRelat.VC_Clear (FormRelat:6465-)。VC の画像とラベルを隠す。"""
        self.env.vc = None

    def write_log(self, state: KensaState, text: str) -> None:
        self.env.log(legacy_log.line_entry(state.line_name, text))

    def user_form_click(self, state: KensaState) -> None:
        """UF1.UserForm_Click (UF1:4300-4416)。再判定・入力済みの印・次の入力欄・少枚数の枚数。"""
        env = self.env
        state.mpex_tag = ""
        state.move_flag = ""
        state.picfm_tag = vbx.cstr(vbx.val(state.picfm_tag) + 1)
        fl = judge.advance_check(state.lot, state.dbflag, arrbox=self._arrbox_fn(state))
        if state.load == "off":
            return
        if fl and vbx.val(state.picfm_tag) == 1:
            _msg_error(env, state, judge.ADVANCE_MSG, 4)
            env.sound(judge.SOUND_NEED, "UF1:4324")
        self.loop_judg(state, vbx.val(state.picfm_tag))
        self.hiz_loop_judg(state, vbx.val(state.picfm_tag))
        for p in grid.POSITIONS:
            self.check_labels_by_category(state, p)
        grid.fm_clear(state, "Click")
        state.auto = "on"
        ken_no = vbx.left(state.ken_no_caption, 1)
        apply_judge(state, judge.looking_yellow(to_judge_form(state)), env)
        if state.is_work("薄板"):
            input_router.look_s_blank(state, 1, 10, ken_no)
        elif state.is_work("極厚"):
            input_router.look_blank(state, 1, 14)
        elif state.is_work("少枚数") and state.ken_count != "":
            input_router.smallnumber_blank(state, 1, 10, vbx.clng(vbx.val(state.ken_count)))
        else:
            input_router.look_blank(state, 1, 10)
        if state.is_work("薄板") and state.yama_visible and state.yama == "":
            env.messages.append(Message("山数の入力がありません" + "\r\n" + "選択してください", WARN, source="UF1:4385"))
            return
        th = grid.small_count_th(state, self.calculate_sheets, by_ken_count_blank=True)
        if state.work_caption.get("厚板") == "H" and state.is_work("少枚数") and vbx.val(state.picfm_tag) == 1:
            self.small_number_count(state, th)
        state.flag_label = 1

    def small_number_count(self, state: KensaState, th: int) -> None:
        def dri(nou: str, tori: str, youtc: str, zai: str, tyou: str, atu: float, haba: float, take: float) -> str:
            return self.db_return_item(state, nou, tori, youtc, zai, tyou, atu, haba, take)
        ch = checks.small_number_count(th, to_uf1(state), self.get_lot_info(state, state.lot), dri)
        for k, v in ch.visible.items():
            put(state, k + ".Visible", v)
        for k, v in ch.enabled.items():
            put(state, k + ".Enabled", v)
        for k, v in ch.values.items():
            put(state, k, v)

    def load_text(self, state: KensaState, n: str = "") -> None:
        r = records.load_text(self.store, state.lname, n, current_option=state.work)
        if r is not None:
            restore_lot_source(state, records.load_lot_source(self.store, state.lname))   # D28: 読み戻しの前に
            apply_restore(state, r, self)

    def worker_load(self, state: KensaState) -> None:
        cells = sheet_cells(self.store, line_sheet("検査員", state.lname)) or {}
        set_registry(state, workers.worker_load(cells))

    # --- Load_Text の続き(FormRestore.calls)---------------------------
    def restore_call(self, state: KensaState, call: tuple) -> None:
        name, args = call[0], call[1:]
        safe = layout.VBA_ERRORS
        try:
            if name == "Atu_OPselect":
                # SaveLoad:2839 `Atu_OPselect(val(…))`: ATU As Double。Single に入れない
                layout.atu_opselect(state, args[0], self, msgs=self.env.messages)
            elif name == "LSelect":
                layout.lselect(state, args[0], self, msgs=self.env.messages)
            elif name == "LSelectLateral":
                layout.lselect_lateral(state, args[0])
            elif name == "Six_Display":
                layout.six_display(state, args[0] if state.yama_visible else "off", self)
            elif name == "AtCountSelect":
                layout.at_count_select(state, args[0], self)
            elif name == "InspectionSelect":
                layout.inspection_select(state, args[0], self)
            elif name == "SlabSet":
                layout.slab_set(state, self)
            elif name == "SmallNumberCount判定":
                # SaveLoad:2914-2927。ATU が空なら th = 0(呼ぶか否かは ATU に関係しない)
                th = grid.small_count_th(state, self.calculate_sheets)
                if state.work_caption.get("厚板") == "H" and state.is_work("少枚数"):
                    self.small_number_count(state, th)
            elif name == "DrawListBoxContent":
                # YamaLB.ListIndex を入れると YamaLB_Click が走る(Load_Text。SaveLoad:2830-2836)
                idx = state.v("YamaLB").tag
                if idx == "last":
                    state.v("YamaLB").value = checks.YAMA_ITEMS[-1]
                elif idx.lstrip("-").isdigit() and 0 <= int(idx) < len(checks.YAMA_ITEMS):
                    state.v("YamaLB").value = checks.YAMA_ITEMS[int(idx)]
                if state.atu != "":
                    self.draw_list_box_content(state)
            elif name == "VC_Clear":
                self.vc_clear(state)
            elif name == "VC_Picture":
                self.vc_picture(state)
            # InitialSetting / Conductiv_Check / UFProgress は表示だけ
        except safe:
            pass
        except (datasource.DbMissing, datasource.TableMissing, vbx.TypeMismatch, IndexError) as e:   # VC厚 が読めない
            self.env.messages.append(Message(str(e), ERROR, source="SaveLoad.VC_Picture"))


# ------------------------------------------------------------------
# 帳票の PDF の置き場所(HTML/PDF は report_html.py)
# ------------------------------------------------------------------
def pdf_folders(line: str, now: datetime) -> list[str]:
    """Input.履歴保存 の PDF の置き場所(Input:1027-1062)。`SaveDateLocate(パス, LN)` = パス\\LN\\yyyy.mm\\dd日。

    手元(C:\\検査履歴 = 使用パス local_history)と工場共有(FactoryyPath & "【梱包】" & LN)。
    ファイル名は成績票 `<C2 の Lot>.pdf`、不良要因票 `<B1 の Lot>_不良要因.pdf`(PDFCREATE 補助:717-741)。
    """
    ln = lines.find(line)
    big = ln.big if ln else "Test"
    month, day = now.strftime("%Y.%m"), now.strftime("%d") + "日"
    return [paths.join(paths.get("local_history"), big, month, day),
            paths.join(paths.get("factory"), "【梱包】" + big, month, day)]


# ------------------------------------------------------------------
# オーダーNo(OrdNoOpen / OrdNoCB_Change)。包装仕様書の確認(PackingNoCheck / UF_PackNo)は packing.py
# ------------------------------------------------------------------
def _hist_cols(hist: Sequence[Sequence[Any]], names: Sequence[str]) -> dict[str, int]:
    """Lot情報履歴 の 1 行目から列番号(1 始まり。無ければ 0 = VBA の未代入の Long)。同名なら後の列。"""
    out = {n: 0 for n in names}
    for i, v in enumerate(hist[0] if hist else [], start=1):
        if cell_text(v) in out:
            out[cell_text(v)] = i
    return out


def _hcell(row: Sequence[Any], col: int) -> Any:
    """`Hist(i, col)`。列 0(見出しが無い)は VBA では添字エラー 9。"""
    if col < 1:
        raise IndexError("インデックスが有効範囲にありません(Lot情報履歴の見出しが無い)")
    return row[col - 1] if col <= len(row) else None


def _lot_eq(v: Any, lot: str) -> bool:
    """`Hist(i, R_Lot) = LotNo`(Variant と String)。"""
    try:
        return vbx.compare_typed(v, lot, vbx.VAR, vbx.STR) == 0
    except vbx.TypeMismatch:
        return False


def _lot_hits(hist: Sequence[Sequence[Any]], lot: str, cols: dict[str, int], take: Sequence[str]) -> list[list[Any]]:
    """`For i = LBound(Hist, 1) To UBound(Hist, 1)`(見出し行も対象)でロットが一致する行の、指定の列。"""
    if not hist:
        return []
    r_lot = cols["ﾛｯﾄNo"]
    return [[_hcell(row, cols[n]) for n in take] for row in hist if _lot_eq(_hcell(row, r_lot), lot)]


def ord_no_open(hist: Optional[Sequence[Sequence[Any]]], lot: str) -> Optional[list[str]]:
    """SaveLoad.OrdNoOpen (SaveLoad:2452-2499)。ロットのオーダーNo を引当No 順に(空は入れない)。

    None = シートが無い・ロットが空・一致なし(VBA の Exit Sub。コンボも PackingNo なども変えない)。
    一覧を返したときは PackingNo / PackingForm / PackingComment を空にする(呼び手)。
    """
    if hist is None or lot == "":
        return None
    cols = _hist_cols(hist, ("ﾛｯﾄNo", "ｵｰﾀﾞｰNo", "引当NO", "包装仕様NO"))
    rows = _lot_hits(hist, lot, cols, ("ｵｰﾀﾞｰNo", "包装仕様NO", "引当NO"))
    if not rows:
        return None
    from .packing import merge_sort2
    rows = merge_sort2(rows, 3)                                        # 補助.merge_sort2(引当NO)
    return [cell_text(r[0]) for r in rows if cell_text(r[0]) != ""]


@dataclass
class OrderPacking:
    """OrdNoCB_Change の結果(PackingNo / PackingForm / PackingComment の Caption)。None = 変えない。"""
    packing_no: Optional[str] = None
    packing_form: Optional[str] = None
    packing_comment: Optional[str] = None
    reload: bool = False                     # VC_Picture(オーダー) と MaterialsMsgBox(オーダー) を呼ぶ


def ord_no_change(hist: Optional[Sequence[Sequence[Any]]], lot: str, ord_no: str, current_packing_no: str,
                  test: bool = False) -> OrderPacking:
    """UF1.OrdNoCB_Change (UF1:1721-1790)。選んだオーダーの包装仕様No と、伝送用ファイル PkgSpecData の梱包形態・タイトル。

    - PackingNo は一致した最後の行の値。一致が無ければ前の値のまま(VBA と同じ)
    - 伝送用ファイルは PackingPath(Test でも同じ)。無ければここで終わり(ReadCheck)
    """
    out = OrderPacking()
    if hist is None or lot == "":
        return out
    cols = _hist_cols(hist, ("ﾛｯﾄNo", "ｵｰﾀﾞｰNo", "包装仕様NO"))
    rows = _lot_hits(hist, lot, cols, ("ｵｰﾀﾞｰNo", "包装仕様NO"))
    if not rows:
        return out
    pno = current_packing_no
    for r in rows:
        try:
            if vbx.compare_typed(ord_no, r[0], vbx.STR, vbx.VAR) == 0:
                pno = cell_text(r[1])
        except vbx.TypeMismatch:
            pass
    out.packing_no = pno
    if pno == "":
        return out
    spec = _pkg_spec(pno)
    if spec is None:
        return out
    fields, rec = spec
    if rec is not None:
        c_form, c_title = datasource.column_index(fields, "梱包形態"), datasource.column_index(fields, "タイトル")
        # 見出しが無いと VBA の C_梱包形態 = 0 → 1 列目(LBound の列)を読む
        out.packing_form = rec[0][max(c_form, 0)]
        out.packing_comment = rec[0][max(c_title, 0)]
    out.reload = True
    return out


def _pkg_spec(n: str) -> Optional[tuple[list[str], Optional[list[list[str]]]]]:
    """伝送用ファイル PkgSpecData の `[包装仕様書№]='n'`(PackingPath)。ファイルが無ければ None(ReadCheck)。"""
    from .lot import read_check
    file_path = paths.join(paths.get("packing"), "伝送用ファイル.sqlite3")
    if not read_check(file_path, 1):
        return None
    flt = "[包装仕様書№]='" + n + "'"
    rec = datasource.get_records(file_path, "PkgSpecData", flt)
    fields = datasource.get_fields(file_path, "PkgSpecData", "")
    if fields is None:
        return None
    return fields, rec


__all__ = ["SCALAR", "TAG_SCALAR", "put", "get", "all_controls", "to_uf1", "to_judge_form", "apply_judge",
           "to_lot_screen", "apply_lot_reflect", "to_snapshot", "apply_restore", "registry", "set_registry",
           "work_dist", "worker_sheet_a", "worker_day_max", "sheet_cells", "sheet_rows", "write_cells", "Env",
           "RealHooks", "time_posting", "time_record", "daily_report", "pdf_folders"]
