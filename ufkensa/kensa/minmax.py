"""Min/Max と下線(FormRelat.MinMax / UnderBar / NGrange_Check ほか)

範囲は OPS1〜6(開始)・OPE1〜6(終了)= 位置 S,U,C,D,E,F の番号 1〜6。
許容でない NG(NG 色で、その位置の NG? が "許容" でない)が範囲にあれば「許容でない公差外が存在します」で止める。
"""
from __future__ import annotations

from typing import Any

from . import vbx
from .flow import WARN, Message, Stop
from .grid import COLOR_DEFAULT, COLOR_NG, POSITIONS, Hooks, KensaState, remove_underline, small_count_th
from .grid import typed_cmp, variant_eq

#: TB_BackColor の色(FormRelat:5911-5912)
COLOR_RANGE_START = "26367"
COLOR_RANGE_END = "13421619"


# ------------------------------------------------------------------
# NG の確認
# ------------------------------------------------------------------
def backcolor_check(state: KensaState, name: str) -> bool:
    """FormRelat.BackColor_Check (FormRelat:5495-5503)。値があり NG 色で、NG? が "許容" でなければ真。"""
    c = state.ctl(name)
    if c.value != "" and c.color == COLOR_NG:
        if state.ctl("NG" + vbx.left(name, 1)).value != "許容":
            return True
    return False


def backcolor_check_tb(state: KensaState, name: str) -> bool:
    """FormRelat.BackColor_CheckTB (FormRelat:5504-5532)。TextBox 版(ZS・GS・LateralS・DiagSKou など)。"""
    temp = ""
    if len(name) == 2:
        temp = vbx.right(name, 1)
        if vbx.is_numeric(temp):
            temp = vbx.left(name, 1)
    elif len(name) == 3:
        temp = vbx.right(name, 2)
        if vbx.is_numeric(temp):
            temp = vbx.left(name, 1)
    if vbx.instr(name, "Lateral") > 0:
        temp = vbx.right(name, 1)
    if vbx.instr(name, "Diag") > 0:
        temp = vbx.mid(name, 5, 1)
    c = state.ctl(name)
    if c.value != "" and c.color == COLOR_NG:
        if state.ctl("NG" + temp).value != "許容":
            return True
    return False


def ngrange_check(state: KensaState, n: int, k: int) -> bool:
    """FormRelat.NGrange_Check (FormRelat:5533-5601)。位置 n〜k に許容でない NG があれば真。

    行は 1〜8(極厚は 1〜12。対角は見ない)、歪 Z?・G?、Sumisei があればラテラルボー、さらに R115 なら直角度。
    """
    result = False
    x = 12 if state.is_work("極厚") else 8
    for i in range(1, x + 1):
        for a in range(n, k + 1):
            if 1 <= a <= 6 and backcolor_check(state, f"{POSITIONS[a - 1]}{i}"):
                result = True
    for prefix in ("Z", "G"):
        for a in range(n, k + 1):
            if 1 <= a <= 6 and backcolor_check_tb(state, prefix + POSITIONS[a - 1]):
                result = True
    if state.sumisei != "":
        for a in range(n, k + 1):
            if 1 <= a <= 6 and backcolor_check_tb(state, "Lateral" + POSITIONS[a - 1]):
                result = True
    if state.sumisei != "" and state.youto == "R115":
        for a in range(n, k + 1):
            if 1 <= a <= 6 and backcolor_check_tb(state, "Diag" + POSITIONS[a - 1] + "Kou"):
                result = True
    return result


# ------------------------------------------------------------------
# 範囲の値
# ------------------------------------------------------------------
def uf_array(state: KensaState, n: int, k: int, x: int, y: int) -> list[str]:
    """SaveLoad.UF_Array(n, k, x, y)(SaveLoad:2135-)の行 1〜14 の部分。行 n〜k × 位置 x〜y の Caption。

    行 11〜14 の F は "" (F11 以降は無い)。並びは VBA の配列と同じく行ごと・位置順。
    """
    out: list[str] = []
    for i in range(n, k + 1):
        for a in range(x, y + 1):
            if not 1 <= a <= 6:
                continue
            if 11 <= i <= 14 and a == 6:
                out.append("")
            elif 1 <= i <= 14:
                out.append(state.cells[f"{POSITIONS[a - 1]}{i}"].value)
    return out


def _numbers(values: list[str]) -> list[Any]:
    """`If Aran(a, i) <> "" Then Aran(a, i) = val(Aran(a, i))`(空はそのまま)。"""
    return [vbx.val(v) if v != "" else v for v in values]


def _ws_min(values: list[Any]) -> float:
    """WorksheetFunction.Min(配列)。配列の中の文字列は無視、数が無ければ 0。"""
    nums = [v for v in values if isinstance(v, float)]
    return min(nums) if nums else 0.0


def _ws_max(values: list[Any]) -> float:
    nums = [v for v in values if isinstance(v, float)]
    return max(nums) if nums else 0.0


def _ranges(state: KensaState, rows: tuple[int, int], s: int, e: int, th: int) -> tuple[int, int, int, int]:
    """幅・丈の範囲(MinMax:5694-5795)。厚板で S=E=U なら S、S=E=D なら C を見る(巾丈を測らない位置)。"""
    a, b = rows
    if state.is_work("極厚") or state.is_work("薄板"):
        return a, b, s, e
    if state.is_work("厚板") or (state.is_work("少枚数") and th == 5):
        if s == 2 and e == 2:
            return a, b, 1, 1
        if s == 4 and e == 4:
            return a, b, 3, 3
        return a, b, s, e
    if state.is_work("少枚数") and th == 4:
        if s == 2 and e == 2:
            return a, b, 1, 1
        if s == 4 and e == 4:
            return a, b, 4, 4
    return a, b, s, e


def min_max(state: KensaState, hooks: Hooks) -> list[Message]:
    """FormRelat.MinMax (FormRelat:5657-5906)。範囲の厚・幅・丈の最小・最大(AMin..TMax)と下線。

    B1: 厚の書式の `If 0 < val(ATU) <= 74# Then` は VBA では常に真(`(0 < x)` が -1/0)→ 厚 74 超でも "0.000"。
    値が 0(数が無い)なら "" にする。
    """
    s, e = state.ops, state.ope
    if s == 0 or e == 0:
        return []
    if s > e:
        return []
    if ngrange_check(state, s, e):
        text = "許容でない公差外が存在します"
        raise Stop(text, messages=[Message(text, WARN, title="警告", source="FormRelat:5683")], level=WARN,
                   source="FormRelat:5683")

    aran = _numbers(uf_array(state, 1, 4, s, e))
    th = 0
    hrng = (5, 10) if state.is_work("極厚") else (5, 6)
    if state.is_work("少枚数"):                              # VBA は幅と丈で 2 回求める
        th = small_count_th(state, hooks.calculate_sheets)
    hran = _numbers(uf_array(state, *_ranges(state, hrng, s, e, th)))
    trng = (11, 14) if state.is_work("極厚") else (7, 8)
    if state.is_work("少枚数"):
        th = small_count_th(state, hooks.calculate_sheets)
    tran = _numbers(uf_array(state, *_ranges(state, trng, s, e, th)))

    atu = vbx.val(state.atu)
    first = 0 < atu <= 74.0
    if first:
        state.amin = vbx.fmt(_ws_min(aran), "0.000")
        state.amax = vbx.fmt(_ws_max(aran), "0.000")
    elif atu > 74:
        state.amin = vbx.fmt(_ws_min(aran), "0.00")
        state.amax = vbx.fmt(_ws_max(aran), "0.00")
    state.hmin = vbx.fmt(_ws_min(hran), "0.0")
    state.hmax = vbx.fmt(_ws_max(hran), "0.0")
    state.tmin = vbx.fmt(_ws_min(tran), "0.0")
    state.tmax = vbx.fmt(_ws_max(tran), "0.0")

    # `If UF1.AMin = 0 Then UF1.AMin = ""`(Label.Caption は String。空なら型エラー)
    for attr in ("amin", "amax", "hmin", "hmax", "tmin", "tmax"):
        if typed_cmp(getattr(state, attr), 0) == 0:
            setattr(state, attr, "")

    remove_underline(state)
    if state.amin != "" and state.amax != "":
        under_bar(state, "A", s, e, vbx.cdbl(state.amin), vbx.cdbl(state.amax))
    if state.hmin != "" and state.hmax != "":
        _, _, hs, he = _ranges(state, (0, 0), s, e, th)
        under_bar(state, "H", hs, he, vbx.cdbl(state.hmin), vbx.cdbl(state.hmax))
    if state.tmin != "" and state.tmax != "":
        _, _, ts, te = _ranges(state, (0, 0), s, e, th)
        under_bar(state, "T", ts, te, vbx.cdbl(state.tmin), vbx.cdbl(state.tmax))
    return []


def under_bar(state: KensaState, buf: str, n: int, k: int, mn: float, mx: float) -> None:
    """FormRelat.UnderBar (FormRelat:5961-6167)。位置 n〜k の行で、Caption が最小・最大と等しい欄に下線。

    行: "A" 1〜4、"H" 5〜6(極厚 5〜10)、"T" 7〜8(極厚 11〜14)。比較は Caption(Variant)と Double の `=`。
    """
    if buf == "A":
        x, y = 1, 4
    elif buf == "H":
        x, y = (5, 10) if state.is_work("極厚") else (5, 6)
    elif buf == "T":
        x, y = (11, 14) if state.is_work("極厚") else (7, 8)
    else:
        return
    if not (1 <= n <= 6 and 1 <= k <= 6) or n > k:
        return
    cols = POSITIONS[n - 1:k]
    for i in range(x, y + 1):
        for target in (mn, mx):
            for p in cols:
                c = state.ctl(f"{p}{i}")
                if variant_eq(c.value, target):
                    c.underline = True


# ------------------------------------------------------------------
# 範囲の選択(OPS / OPE)
# ------------------------------------------------------------------
def tb_backcolor_clear(state: KensaState, which: str) -> None:
    """FormRelat.TB_BackColor_Clear (FormRelat:5476-5494)。which = "S"(開始)/ "E"(終了)。"""
    for p in POSITIONS:
        state.v(("TS" if which == "S" else "TE") + p).color = COLOR_DEFAULT


def op_count(state: KensaState) -> None:
    """FormRelat.OPcount (FormRelat:5942-5960)。開始 > 終了なら SFM.Tag = "Stop"。"""
    if state.ops and state.ope and state.ops > state.ope:
        state.sfm_tag = "Stop"


def tb_backcolor(state: KensaState, which: str, idx: int) -> None:
    """FormRelat.TB_BackColor (FormRelat:5908-5941)。OPS?/OPE? が True になったとき範囲の色を付ける。"""
    if (state.ops if which == "S" else state.ope) != idx:
        return
    tb_backcolor_clear(state, which)
    op_count(state)
    if state.sfm_tag != "":
        state.sfm_tag = ""
        return
    if 1 <= idx <= 6:
        if which == "S":
            state.v("TS" + POSITIONS[idx - 1]).color = COLOR_RANGE_START
        else:
            state.v("TE" + POSITIONS[idx - 1]).color = COLOR_RANGE_END


def select_range(state: KensaState, which: str, idx: int) -> None:
    """UF1.OPS1_Change..OPE6_Change (UF1:3535-3571)。OptionButton を選んで範囲の色を付ける。"""
    if which == "S":
        state.ops = idx
    else:
        state.ope = idx
    tb_backcolor(state, which, idx)


def run_min_max(state: KensaState, hooks: Hooks) -> list[Message]:
    """UF1.CommandButton4_Click (UF1:4459-4468)。再判定してから MinMax。Load を止めて最後に戻す。"""
    with state.paused():                                   # MinMax の Exit Sub・途中のエラーの後も Load="on" に戻る
        hooks.loop_judg(state, vbx.val(state.picfm_tag))
        hooks.hiz_loop_judg(state, vbx.val(state.picfm_tag))
        return min_max(state, hooks)


__all__ = ["backcolor_check", "backcolor_check_tb", "ngrange_check", "uf_array", "min_max", "under_bar",
           "tb_backcolor", "tb_backcolor_clear", "op_count", "select_range", "run_min_max"]
