"""直尺(伝送コンベックス)の校正

VBA `【UFKen_10】Proofreading` と `UF校正` の移植。

- 校正済みかは「校正記録」シートの**最終行の 1 行**(A=日付、D=直)が、いまの(直を考慮した日付、直)と
  一致するかで決まる(`check_calibration`)。不合格の行は D が空なので、最後が不合格なら未校正。
- D11: VBA は UF1.ProofreadFlag を UF1 の起動・検査員登録・測定完了のときしか更新しないため、
  開いたまま直が変わっても止まらなかった。移植では**入力のたびに判定**する(`input_blocked`)。
- 画面は 2 段階: ① 1000mm 直尺(品証用、999.5〜1000.5。Test は 50〜3000)→ 合格で白い画面に
  ② ライン別ストレートエッヂの全長(DB「校正ストレートエッヂ」の 長さ下〜長さ上。両端を含む)→ 合格で記録・校正済み。
- シートは今のセルを受け取り、書くセルを返す(保存先の表は別担当)。

B14: VBA は「送り」ボタン(CommandButton5)で 1000mm をとばせる。また現行の 測定完了 は不合格の行に赤字を
付けないので、赤字の行を拾う 校正ランク付け に新しい記録が出ない。直すと、とばせない・不合格を赤字にする。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Mapping, Optional

from . import legacy_log, lines, quirks, vbx
from . import sound as snd
from .flow import ERROR, INFO, Message, Stop
from .workers import (CellWrite, ShiftTimes, WorkerRegistry, _seconds, hhmm, last_row, time_check,
                      worke_day_max, yea_dayr_check)

Cells = Mapping[tuple[int, int], Any]
GREEN, WHITE = "1000", "gauge"                  # 背景 &H80FF80(1000mm)/ &HFFFFFF(ストレートエッヂ)
PASS_SOUND, FAIL_SOUND = snd.CALIBRATION_OK, snd.CALIBRATION_NG   # 音楽合格 / 音楽不合格(D29: 出来事の鍵)
DEFAULT_GAUGE = "1000mm直尺の全長を測定してください"   # Gauge の設計時の Caption(G6)
NG_1000_TEXT = "999.5 ≦" + "測定値" + "≦ 1000.5 になっていません"   # Proofreading:365(Test でも同じ文)
RED = "red"                                      # RGB(255,128,128) の赤字


def _text(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return vbx.cstr(v)


def _official(line: object) -> str:
    ln = lines.find(line)
    return ln.official if ln else ""


# ------------------------------------------------------------------
# 未校正の判定
# ------------------------------------------------------------------
@dataclass
class CalibrationCheck:
    uncalibrated: bool             # True = 未校正(VBA の 校正() の戻り値)
    proofread_flag: Optional[str]  # UF1.ProofreadFlag に入れる値(None = 変えない)
    create_sheet: bool = False     # 校正記録 シートを作る(AddWS)
    clear_jump_tag: bool = False   # UF1.WorkButton.Tag = ""


def last_record(cells: Cells) -> tuple[Any, Any]:
    """校正記録の A 列の最終行の (A, D)(`Cells(Rows.Count, 1).End(xlUp).Row`)。"""
    k = last_row(cells, 1)
    return cells.get((k, 1)), cells.get((k, 4))


def check_calibration(st_time: str, cells: Optional[Cells], *, now: datetime, facility: str, shifts: ShiftTimes,
                      jump: bool = False, db_flag: str = "") -> CalibrationCheck:
    """Proofreading.校正 (Proofreading:28-61)。cells は 校正記録 シート(無ければ None)。

    - WorkButton.Tag="jump" か DBflag があれば済み扱い(Tag は消す)
    - シートが無ければ作って未校正
    - 最終行の A が YeaDayrCheck(st_time)、D が TimeCheck(st_time) と両方一致すれば済み(Flag="ok")
    """
    if jump or db_flag != "":
        return CalibrationCheck(False, None, clear_jump_tag=True)
    if cells is None:
        return CalibrationCheck(True, None, create_sheet=True)
    x = time_check(st_time, facility=facility, shifts=shifts)
    y = yea_dayr_check(st_time, now=now, facility=facility, shifts=shifts)
    temp_d, temp_t = last_record(cells)
    if _text(temp_d) != y:
        return CalibrationCheck(True, "")
    if _text(temp_t) != x:
        return CalibrationCheck(True, "")
    return CalibrationCheck(False, "ok")


def input_blocked(cells: Optional[Cells], *, now: datetime, facility: str, shifts: ShiftTimes,
                  jump: bool = False, db_flag: str = "") -> bool:
    """測定値の入力を止めるか(UF1:4615-4626 の `ProofreadFlag <> "ok"`)。D11 により毎回 校正() で判定する。

    jump / DBflag のときは止めない(VBA と同じ)。VBA と違い、ここでは Tag を消さない。
    """
    if jump or db_flag != "":
        return False
    return check_calibration(hhmm(now), cells, now=now, facility=facility, shifts=shifts).uncalibrated


@dataclass
class ProofreadLoad:
    message: Message
    kname: str                      # 校正画面の検査員(最新の直の末尾の人)
    line_name: str


def proofread_load(registry: WorkerRegistry, sheet_a: tuple[Any, Any, Any], line_name: str, *,
                   blocked: bool) -> Optional[ProofreadLoad]:
    """Proofreading.Proofread_Load (Proofreading:638-654)。未校正なら案内を出して UF校正 を開く。

    blocked は `input_blocked(...)` の結果(VBA は jump/DBflag/ProofreadFlag="ok" で抜ける)。
    sheet_a は 検査員_<L> の A1:A3(WorkeDayMax 用)。
    """
    if not blocked:
        return None
    t = {"1": registry.t1, "2": registry.t2, "3": registry.t3}.get(worke_day_max(*sheet_a), "")
    end_work = t.split("_")[-1] if t != "" else ""
    return ProofreadLoad(Message("校正が済んでいません。校正を起動します。", ERROR, source="Proofreading:652"),
                         end_work, line_name)


# ------------------------------------------------------------------
# 案内文
# ------------------------------------------------------------------
def open_guidance(uf1_line: str) -> str:
    """Proofreading.校正_Open (Proofreading:10-17)。UF1 で選んでいるラインで Gauge の文を決める。丸徳は設計時の文。"""
    return {
        "L-1": "[L-1] P2:1000mm直尺の全長を測定してください",
        "LVC": "[LVC] P2:1000mm直尺の全長を測定してください",
        "HVC": "[HVC] P2:1000mm直尺の全長を測定してください",
        "バランサー": "[大板] P2:1000mm直尺の全長を測定してください",
        "トット": "[大板] P2:1000mm直尺の全長を測定してください",
        "Test": "[Test] xx:1000mm直尺の全長を測定してください",
    }.get(_official(uf1_line), DEFAULT_GAUGE)


def gauge_guidance(line_name: str) -> str:
    """1000mm 合格後の案内(Proofreading:325-333)。L-1 は「S-2」、LVC/HVC/大板は「S-1」。"""
    return {
        "L-1": "[L-1]:直尺「S-2」の全長を測定してください",
        "LVC": "[LVC]:直尺「S-1」の全長を測定してください",
        "HVC": "[HVC]:直尺「S-1」の全長を測定してください",
        "バランサー": "[大板]:直尺「S-1」の全長を測定してください",
        "トット": "[大板]:直尺「S-1」の全長を測定してください",
        "Test": "[Test]:直尺「xxx」の全長を測定してください",
    }.get(_official(line_name), "[中板]:直尺「xxx」の全長を測定してください")


def _hinsho_line(line_name: str) -> str:
    """品証測定 のライン(Proofreading:265-273)。"""
    return {"L-1": "L-1", "LVC": "LVC", "HVC": "HVC", "バランサー": "ﾊﾞﾗﾝｻｰ", "トット": "ﾄｯﾄ",
            "Test": "Test"}.get(_official(line_name), "丸徳")


# ------------------------------------------------------------------
# 伝送文字列
# ------------------------------------------------------------------
def _narrow(s: str) -> str:
    """StrConv(…, vbNarrow) の英数字・記号・空白の部分。"""
    out = []
    for ch in s:
        o = ord(ch)
        if 0xFF01 <= o <= 0xFF5E:
            out.append(chr(o - 0xFEE0))
        elif ch == "　":
            out.append(" ")
        else:
            out.append(ch)
    return "".join(out)


def ter_maker(trm: str) -> str:
    """UF校正.TerMaker (UF校正:36-47)。7 文字以上で、先頭 2 文字(大文字・半角にして)が DT / #L。"""
    if trm == "" or len(trm) < 7:
        return ""
    head = _narrow(trm[:2].upper())
    return head if head in ("DT", "#L") else ""


def ter_conv(ter: str) -> tuple[str, list[Message]]:
    """UF校正.TerConv (UF校正:24-35)。"DT…" → 14 文字目から 6 文字、"#L…" → 3 文字目から。ほかは "" と警告。"""
    if ter == "":
        return "", []
    kind = ter_maker(ter)
    if kind == "DT":
        return vbx.mid(ter, 14, 6), []
    if kind == "#L":
        return vbx.mid(ter, 3, len(ter)), []
    return "", [Message("正しいデータが伝送できていません", ERROR, source="UF校正:33")]


def safe_num(s: Any) -> tuple[float, bool]:
    """Proofreading.SafeNum (Proofreading:65-78)。カンマを除き、小数点 2 個以上・数でないものは失敗。"""
    if s is None:
        return 0.0, False
    t = _text(s).strip()
    if t == "":
        return 0.0, False
    t = t.replace(",", "")
    if len(t) - len(t.replace(".", "")) > 1:
        return 0.0, False
    if not vbx.is_numeric(t):
        return 0.0, False
    return vbx.cdbl(t), True


# ------------------------------------------------------------------
# UF校正 の状態と入力
# ------------------------------------------------------------------
@dataclass
class CalibrationState:
    """UF校正 の画面(コントロールの値・色・フラグ)。"""
    line_name: str                   # UF校正.LineName(UF1 の LineName & LineNo)
    kname: str = ""
    gauge: str = DEFAULT_GAUGE
    mode: str = GREEN                # 背景色: 緑=1000mm、白=ストレートエッヂ
    f1: str = "1"                    # 1=入力待ち 0=処理中 2=品証NG "end"=完了
    sokutei: str = ""                # SokuteiH
    inp1000: str = ""
    inpg: str = ""
    ng_value: str = ""
    ng_visible: bool = False
    saiken: bool = False             # 「ｸﾘｱを押してもう1度測定してください」
    ranks: dict[int, str] = field(default_factory=dict)
    proofread_ok: bool = False       # UF1.ProofreadFlag = "ok" にした
    close: bool = False              # INPG_Exit で閉じて UF1 を再判定(LoopJudg・HizLoopJudg)


@dataclass
class StepResult:
    messages: list[Message] = field(default_factory=list)
    logs: list[legacy_log.LogEntry] = field(default_factory=list)
    kiroku_writes: list[CellWrite] = field(default_factory=list)   # 校正記録
    kiroku_red: list[tuple[int, int]] = field(default_factory=list)
    kiroku_create: bool = False
    hinsho_writes: list[CellWrite] = field(default_factory=list)   # 品証検証
    hinsho_red: list[tuple[int, int]] = field(default_factory=list)
    kiroku_rank: Optional["RankResult"] = None                      # 1000mm 合格後の 校正ランク付け(校正記録を書き換える)


def _log(line_name: str, text: str) -> legacy_log.LogEntry:
    """UF校正 のログ `*<フォルダ名>_<文>`(UF校正:229-235 ほか)。"""
    return legacy_log.line_entry(line_name, "*" + legacy_log.log_folder(line_name) + "_" + text)


def initialize(line_name: str, kname: str, uf1_line: str) -> CalibrationState:
    """校正_Open + UF校正.Initialize(UF校正:342-391)。品証ランク付けは呼び手が `hinsho_rank` で。"""
    return CalibrationState(line_name=line_name, kname=kname, gauge=open_guidance(uf1_line))


def activate_log(state: CalibrationState) -> legacy_log.LogEntry:
    """UF校正.Activate のログ(UF校正:329-339)。"""
    return _log(state.line_name, "UF校正起動")


def _hms(now: datetime) -> str:
    return f"{now.hour}時{now.minute:02d}分{now.second:02d}秒"


def _night(now: datetime, shifts: ShiftTimes) -> bool:
    """`Time >= CDate("00:00") And Time <= CDate(End3)`(秒まで比べる)。"""
    t = now.hour * 3600 + now.minute * 60 + now.second
    e3 = _seconds(shifts.end3)
    return t >= 0 and t <= e3


def input_1000(state: CalibrationState, raw: str, hinsho: Optional[Cells], kiroku: Optional[Cells], *,
               now: datetime, facility: str, shifts: ShiftTimes, red_rows: Optional[set[int]] = None) -> StepResult:
    """UF校正.INP1000_AfterUpdate (UF校正:220-240) + Proofreading.品証測定 (Proofreading:258-386)。

    hinsho は 品証検証 シート(無ければ VBA は End → Stop)、kiroku は 校正記録(合格後の 校正ランク付け 用。
    red_rows はその赤字の行)。
    白い画面(ストレートエッヂ)では INP1000 に入れない(INP1000_Enter)ので何もしない。
    """
    res = StepResult()
    if state.mode != GREEN or state.f1 != "1":
        return res
    state.f1 = "0"
    state.inp1000, msgs = ter_conv(raw)
    res.messages += msgs
    state.sokutei = state.inp1000
    if state.sokutei != "":
        res.logs.append(_log(state.line_name, "1000mm測定:" + state.sokutei))
        _hinsho_measure(state, hinsho, kiroku, res, now=now, facility=facility, shifts=shifts, red_rows=red_rows)
    state.f1 = "1"
    return res


def _hinsho_measure(state: CalibrationState, sh: Optional[Cells], kiroku: Optional[Cells], res: StepResult, *,
                    now: datetime, facility: str, shifts: ShiftTimes, red_rows: Optional[set[int]] = None) -> None:
    """Proofreading.品証測定 (Proofreading:258-386)。"""
    if sh is None:
        raise Stop("品証検証 シートがありません(VBA は End)", messages=res.messages, source="Proofreading:260")
    k = last_row(sh, 13) + 1
    line = _hinsho_line(state.line_name)
    x = time_check(hhmm(now), facility=facility, shifts=shifts)
    if x == "日勤":
        x = "1直"
    yea_dayr_check(hhmm(now), now=now, facility=facility, shifts=shifts)     # y は使わない(VBA どおり計算だけ)
    formay = f"{now.day}日"
    if state.line_name == "Test":
        koumin, koumax = 50.0, 3000.0
    else:
        koumin, koumax = 999.5, 1000.5
    if not vbx.is_numeric(state.inp1000):
        state.inp1000 = ""
        state.sokutei = state.inp1000
        return
    v = vbx.val(state.sokutei)
    if v >= vbx.val(koumin) and v <= vbx.val(koumax):
        res.messages.append(Message("", INFO, sound=PASS_SOUND, source="Proofreading:297"))
        state.saiken = False
        stamp = now.strftime("%Y.%m.%d ") + _hms(now)
        for i in range(6, 97):
            if _text(sh.get((i, 1))) == formay:
                for r in (i, i + 1, i + 2):
                    if _text(sh.get((r, 2))) == x:
                        res.hinsho_writes += [CellWrite(r, 3, state.sokutei), CellWrite(r, 9, line),
                                              CellWrite(r, 10, state.kname), CellWrite(r, 11, stamp)]
                        break
        state.saiken = False
        state.ng_value = ""
        state.inp1000 = ""
        state.sokutei = ""
        state.gauge = gauge_guidance(state.line_name)
        state.ranks = {}
        if kiroku is not None:
            res.kiroku_rank = calibration_rank(kiroku, state.line_name, now=now, red_rows=red_rows)
            state.ranks.update(res.kiroku_rank.ranks)
        state.mode = WHITE
        state.f1 = "0"
    else:
        res.messages.append(Message("", INFO, sound=FAIL_SOUND, source="Proofreading:361"))
        state.sokutei = state.inp1000
        state.saiken = True
        state.ng_visible = True
        state.ng_value = NG_1000_TEXT
        state.f1 = "2"
        if state.inp1000 == "":
            state.saiken = False
            state.ng_value = ""
        res.hinsho_writes += [CellWrite(k, 16, now.strftime("%Y.%m.%d ") + _hms(now)),
                              CellWrite(k, 13, state.sokutei), CellWrite(k, 15, state.kname), CellWrite(k, 14, line)]
        res.hinsho_red += [(k, 13), (k, 14), (k, 15), (k, 16)]
        state.inp1000 = ""
        state.f1 = "0"


def input_gauge(state: CalibrationState, raw: str, kiroku: Optional[Cells],
                straight_edge: Optional[tuple[Any, Any]], *, now: datetime, facility: str,
                shifts: ShiftTimes) -> StepResult:
    """UF校正.INPG_AfterUpdate (UF校正:250-271) + Proofreading.測定完了 (Proofreading:80-155)。

    straight_edge は QP_DB.StraightEdge(設備名) の (長さ下, 長さ上)(該当なしは None)。別担当の関数の結果を受け取る。
    kiroku は 校正記録(無ければ作るだけで記録しない)。緑の画面(1000mm)では INPG に入れない(INPG_Enter)。
    """
    res = StepResult()
    if state.mode != WHITE or state.f1 != "1":
        return res
    state.f1 = "0"
    state.inpg, msgs = ter_conv(raw)
    res.messages += msgs
    state.sokutei = state.inpg
    if state.sokutei != "":
        measured = state.sokutei
        _measure_done(state, kiroku, straight_edge, res, now=now, facility=facility, shifts=shifts)
        if quirks.reproduce("B66"):
            # VBA: 不合格だと 測定完了 が SokuteiH を消した後にログを書くので、値が残らない
            res.logs.append(_log(state.line_name, "ｹﾞｰｼﾞ測定:" + state.sokutei))
        else:
            res.logs.append(_log(state.line_name, "ｹﾞｰｼﾞ測定:" + measured))
    else:
        state.f1 = "1"
    return res


def _measure_done(state: CalibrationState, ws: Optional[Cells], straight_edge: Optional[tuple[Any, Any]],
                  res: StepResult, *, now: datetime, facility: str, shifts: ShiftTimes) -> None:
    """Proofreading.測定完了 (Proofreading:80-155)。"""
    if ws is None:
        res.kiroku_create = True
        return
    x = time_check(hhmm(now), facility=facility, shifts=shifts)
    y = yea_dayr_check(hhmm(now), now=now, facility=facility, shifts=shifts)
    k = last_row(ws, 1) + 1
    if straight_edge is None:
        return
    koumin, ok_lo = safe_num(straight_edge[0])
    koumax, ok_hi = safe_num(straight_edge[1])
    if not ok_lo or not ok_hi:
        res.messages.append(Message("校正基準値の設定エラーです。管理者に連絡してください。", ERROR,
                                    source="Proofreading:102"))
        return
    v = vbx.val(state.sokutei)
    if v >= vbx.val(koumin) and v <= vbx.val(koumax):
        night = _night(now, shifts)
        res.kiroku_writes.append(CellWrite(k, 1, y))
        if night:
            res.kiroku_writes.append(CellWrite(k, 2, "夜勤日付処理"))
        res.kiroku_writes += [CellWrite(k, 3, _hms(now)), CellWrite(k, 4, x), CellWrite(k, 5, state.sokutei),
                              CellWrite(k, 6, state.kname), CellWrite(k, 7, facility)]
        state.kname = ""
        state.saiken = False
        res.messages.append(Message("", INFO, sound=PASS_SOUND, source="Proofreading:124"))
        state.f1 = "end"
        state.proofread_ok = True
    else:
        state.ng_visible = True
        state.ng_value = ("測定値" + " " + state.inpg + " " + "mm") if state.inpg != "" else ""
        night = _night(now, shifts)
        res.kiroku_writes.append(CellWrite(k, 1, y))
        if night:
            res.kiroku_writes.append(CellWrite(k, 2, "夜勤日付処理"))
        res.kiroku_writes += [CellWrite(k, 3, _hms(now)), CellWrite(k, 5, state.sokutei),
                              CellWrite(k, 6, state.kname), CellWrite(k, 7, facility), CellWrite(k, 8, "〆")]
        if not quirks.reproduce("B14"):
            res.kiroku_red += [(k, 1), (k, 3), (k, 5), (k, 6), (k, 7)]       # 旧コード P:243-247 と同じ赤字
        res.messages.append(Message("", INFO, sound=FAIL_SOUND, source="Proofreading:148"))
        state.inpg = ""
        state.sokutei = ""
        state.f1 = "0"
        state.saiken = True


def exit_gauge(state: CalibrationState) -> bool:
    """UF校正.INPG_Exit (UF校正:281-296)。F1="end" なら閉じて UF1 を表示し LoopJudg・HizLoopJudg で再判定。"""
    if state.f1 == "end":
        state.close = True
    return state.close


def clear_button(state: CalibrationState) -> None:
    """UF校正.CBcl_Click (UF校正:49-66)「ｸﾘｱ&ﾘｾｯﾄ」。F1 を 1 に戻す(ストレートエッヂ不合格の後はこれが要る)。"""
    state.f1 = "1"
    if state.mode == GREEN:
        state.inp1000 = ""
    else:
        state.inpg = ""
    state.sokutei = ""
    state.ng_value = ""
    if state.inp1000 == "":
        state.saiken = False
        state.ng_value = ""


def skip_1000_button(state: CalibrationState) -> bool:
    """UF校正.CommandButton5_Click (UF校正:133-172)「送り」。1000mm をとばして白い画面にする。

    B14: 直すときは何もしない(False)。丸徳は Gauge の文を変えない。
    """
    if not quirks.reproduce("B14"):
        return False
    state.saiken = False
    state.ng_value = ""
    state.inp1000 = ""
    state.sokutei = ""
    g = {"L-1": "[L-1]:直尺「S-2」の全長を測定してください", "LVC": "[LVC]:直尺「S-1」の全長を測定してください",
         "HVC": "[HVC]:直尺「S-1」の全長を測定してください", "バランサー": "[大板]:直尺「S-1」の全長を測定してください",
         "トット": "[大板]:直尺「S-1」の全長を測定してください",
         "Test": "[TEST]:直尺「S-1」の全長を測定してください"}.get(_official(state.line_name))
    if g is not None:
        state.gauge = g
    state.ranks = {}
    state.mode = WHITE
    return True


def hinsho_button1(state: CalibrationState, sh: Cells, *, now: datetime, facility: str,
                   shifts: ShiftTimes) -> list[CellWrite]:
    """UF校正.CommandButton1_Click (UF校正:68-125)「測定完了」(旧い記録の道)。

    品証検証の A 列(「d日」)を YeaDayrCheck の "yyyy.mm.dd" と比べるので、実際には書く行が見つからない。
    SokuteiH が空なら再検査の表示だけ。
    """
    line = {"L-1": "L-1", "LVC": "LVC", "HVC": "HVC", "バランサー": "ﾊﾞﾗﾝｻｰ", "トット": "ﾄｯﾄ",
            "Test": "Test"}.get(_official(state.line_name), "")
    x = time_check(hhmm(now), facility=facility, shifts=shifts)
    y = yea_dayr_check(hhmm(now), now=now, facility=facility, shifts=shifts)
    writes: list[CellWrite] = []
    if state.sokutei == "":
        state.saiken = True
        state.ng_value = ""
        state.inp1000 = ""
        return writes
    stamp = now.strftime("%m/%d ") + _hms(now)
    for i in range(6, 97):
        if _text(sh.get((i, 1))) == y:
            for r in (i, i + 1, i + 2):
                if _text(sh.get((r, 2))) == x:
                    writes += [CellWrite(r, 3, state.sokutei), CellWrite(r, 9, line),
                               CellWrite(r, 10, state.kname), CellWrite(r, 11, stamp)]
                    break
    state.kname = ""
    state.saiken = False
    return writes


# ------------------------------------------------------------------
# 月ごとの不合格値の順位
# ------------------------------------------------------------------
RANK_CENTER = {"L-1": 2053.2, "LVC": 2000.5, "HVC": 2003.8, "ﾊﾞﾗﾝｻｰ": 2051.8, "ﾄｯﾄ": 2051.8}   # Proofreading:417-465


@dataclass
class RankResult:
    rewrites: list[CellWrite]                 # E 列(または M 列)を Format(Val,"0.0") で書き直す
    writes: list[CellWrite]                   # I:O(または R:X)
    clear_columns: tuple[int, int]
    ranks: dict[int, str]                     # Rank1-6 に表示する文(当たったものだけ)


def _nums(vals: list[Any]) -> list[float]:
    return [float(v) for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)]


def _top3(vals: list[Any], *, largest: bool) -> tuple[float, float, float]:
    """MAX/LARGE(または MIN/SMALL)と COUNTIF で 1-3 位(エラーの所は On Error Resume Next で 0 のまま)。"""
    nums = sorted(_nums(vals), reverse=largest)
    first = (max(nums) if largest else min(nums)) if nums else 0.0

    def nth(k: float, keep: float) -> float:
        k = int(k)
        return nums[k - 1] if 1 <= k <= len(nums) else keep
    c1 = sum(1 for v in nums if v == first)
    second = nth(c1 + 1, 0.0)
    c2 = sum(1 for v in nums if v == second)
    third = nth(c1 + 1 + c2 + 1, 0.0)
    return first, second, third


def _cmpnum(v: Any) -> Any:
    return 0 if v is None or v == "" else v


def _rank_text(pos: int, when: Any, value: float, who: Any) -> str:
    return f"{pos}位：　" + _text(when) + "　値：" + vbx.fmt(value, "0.0") + "　作業者：" + _text(who)


def calibration_rank(cells: Cells, line_name: str, *, now: datetime,
                     red_rows: Optional[set[int]] = None) -> RankResult:
    """Proofreading.校正ランク付け (Proofreading:388-517)。

    当月(A 列の先頭 7 文字 = yyyy.mm)の最初の行から、**A 列が赤字の行**を、ラインの中心値の以上(I:K)・未満(M:O)に
    分けて写し、上位・下位 3 位を作る。red_rows は赤字の行番号(B14 を再現している限り新しい記録は赤字にならない)。
    J・N 列の最終行は**消す前の**値で数える(VBA どおり。B67 を直すと写した後で数える)。
    """
    red_rows = red_rows or set()
    k_old = last_row(cells, 10)
    j = last_row(cells, 1)
    l_old = last_row(cells, 14)
    sheet = {key: v for key, v in cells.items() if not (9 <= key[1] <= 15)}
    rewrites = []
    for i in range(2, j + 1):
        txt = vbx.fmt(vbx.val(_text(sheet.get((i, 5)))), "0.0")
        sheet[(i, 5)] = vbx.val(txt)
        rewrites.append(CellWrite(i, 5, vbx.val(txt)))
    month = now.strftime("%Y.%m")
    tuki = None
    for b in range(j, 1, -1):
        if _text(sheet.get((b, 1)))[:7] == month:
            tuki = b
    writes: list[CellWrite] = []
    n = m = 0
    center = RANK_CENTER.get(line_name)
    if tuki is not None and center is not None:
        for c in range(tuki, j + 1):
            if c not in red_rows:
                continue
            e = vbx.val(_text(sheet.get((c, 5))))
            if e >= center:
                n += 1
                cols = (9, 10, 11)
                r = n + 1
            else:
                m += 1
                cols = (13, 14, 15)
                r = m + 1
            for col, src in zip(cols, (1, 5, 6)):
                sheet[(r, col)] = sheet.get((c, src))
                writes.append(CellWrite(r, col, sheet.get((c, src))))
    if not quirks.reproduce("B67"):
        # 直した側: I:O を消して書き写した後の J・N 列で最終行を数える(新しい記録がその回に出る)
        k_old = last_row(sheet, 10)
        l_old = last_row(sheet, 14)
    ran1 = [sheet.get((r, 10)) for r in range(min(2, k_old), max(2, k_old) + 1)]
    ran2 = [sheet.get((r, 14)) for r in range(min(2, l_old), max(2, l_old) + 1)]
    mx = _top3(ran1, largest=True)
    mn = _top3(ran2, largest=False)
    ranks: dict[int, str] = {}
    for a in range(2, k_old + 1):
        v = _cmpnum(sheet.get((a, 10)))
        for pos, target in ((1, mx[0]), (2, mx[1]), (3, mx[2])):
            if v == target:
                ranks[pos] = _rank_text(pos, sheet.get((a, 9)), target, sheet.get((a, 11)))
                break
    for d in range(2, l_old + 1):
        v = _cmpnum(sheet.get((d, 14)))
        for pos, target in ((1, mn[0]), (2, mn[1]), (3, mn[2])):
            if v == target:
                ranks[pos + 3] = _rank_text(pos, sheet.get((d, 13)), target, sheet.get((d, 15)))
                break
    return RankResult(rewrites, writes, (9, 15), ranks)


#: 品証ランク付けの `.Cells(6, 13) = ""`(数になった M6 と "" の比較)が型エラーになるか(U1)。
#: なる場合は On Error Resume Next で Then 側(Exit Sub)へ進むので、順位は出ない。
_U1_NUMBER_VS_EMPTY_ERRORS = True


def hinsho_rank(cells: Cells, *, now: datetime) -> RankResult:
    """Proofreading.品証ランク付け (Proofreading:518-636)。品証検証の M 列(不合格値)を 1000 以上と未満に分ける。"""
    k_old = last_row(cells, 18)
    j = last_row(cells, 16)
    l_old = last_row(cells, 22)
    sheet = {key: v for key, v in cells.items() if not (18 <= key[1] <= 24)}
    rewrites: list[CellWrite] = []
    for i in range(6, j + 1):
        p = _text(sheet.get((i, 16)))
        if p == "":
            continue
        if p[:2] != "20":
            sheet[(i, 16)] = "2019." + p.replace("/", ".")
            rewrites.append(CellWrite(i, 16, sheet[(i, 16)]))
    for i in range(6, j + 1):
        v = vbx.val(vbx.fmt(vbx.val(_text(sheet.get((i, 13)))), "0.0"))
        sheet[(i, 13)] = v
        rewrites.append(CellWrite(i, 13, v))
    month = now.strftime("%Y.%m")
    tuki = None
    for b in range(j, 5, -1):
        if _text(sheet.get((b, 16)))[:7] == month:
            tuki = b
        else:
            break
    m6 = sheet.get((6, 13))
    if m6 is None or m6 == "" or (isinstance(m6, (int, float)) and _U1_NUMBER_VS_EMPTY_ERRORS):
        return RankResult(rewrites, [], (18, 24), {})
    writes: list[CellWrite] = []
    n = m = 5
    if tuki is not None:
        for c in range(tuki, j + 1):
            e = vbx.val(_text(sheet.get((c, 13))))
            if e >= 1000:
                n += 1
                r, cols = n, (18, 19, 20)
            else:
                m += 1
                r, cols = m, (22, 23, 24)
            for col, src in zip(cols, (13, 15, 16)):
                sheet[(r, col)] = sheet.get((c, src))
                writes.append(CellWrite(r, col, sheet.get((c, src))))
    ran1 = [sheet.get((r, 18)) for r in range(min(6, k_old), max(6, k_old) + 1)]
    ran2 = [sheet.get((r, 22)) for r in range(min(6, l_old), max(6, l_old) + 1)]
    mx = _top3(ran1, largest=True)
    mn = _top3(ran2, largest=False)
    ranks: dict[int, str] = {}
    for a in range(6, k_old + 1):
        v = _cmpnum(sheet.get((a, 18)))
        for pos, target in ((1, mx[0]), (2, mx[1]), (3, mx[2])):
            if v == target:
                ranks[pos] = _rank_text(pos, sheet.get((a, 20)), target, sheet.get((a, 19)))
                break
    for d in range(6, l_old + 1):
        v = _cmpnum(sheet.get((d, 22)))
        for pos, target in ((1, mn[0]), (2, mn[1]), (3, mn[2])):
            if v == target:
                ranks[pos + 3] = _rank_text(pos, sheet.get((d, 24)), target, sheet.get((d, 23)))
                break
    return RankResult(rewrites, writes, (18, 24), ranks)
