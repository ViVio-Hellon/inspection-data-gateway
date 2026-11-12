"""伝送文字列の解析と次の入力欄(UF1.INP_Change と FormRelat の PucketMove / AllCheck / Look_*_Blank)

流れ(UF1.INP_Change UF1:4565-4651):
    受信(StartTB → INP) → `maker` で種別 → `pucket_move` で入力先(INP.Tag)→ `conv` の値を入れる
    → 検査員・作業選択・校正のブロック → LoopText(判定。別担当)→ Posting → Out_Tolerance
    → `next_blank` で次の入力欄(黄色) → INP を空にする
判定(LoopText 以降)は `grid.Hooks` 経由で呼ぶだけ。ここは「入力先と値」を決めるところまで。
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional

from . import vbx
from .model import NORMAL, THICK_ROWS, Switch, bad_name, diag_name, is_thick
from .flow import ERROR, INFO, WARN, Message, Stop
from .grid import (BORDER_BLACK, BORDER_RED, COLOR_DEFAULT, COLOR_WHITE, COLOR_YELLOW, Hooks, KensaState,
                   fm_clear, ken_flags, small_count_th)

PACKETS = ("DT", "#L", "#〆")


# ------------------------------------------------------------------
# 伝送文字列
# ------------------------------------------------------------------
def narrow(s: str) -> str:
    """`StrConv(s, vbNarrow)` の英数記号部分(全角 ASCII → 半角、全角空白 → 半角)。〆 などはそのまま。"""
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


def maker(trm: str) -> str:
    """補助.Maker (補助:393-407)。先頭 2 文字を大文字・半角にして "DT" / "#L" / "#〆" を返す。3 文字未満は ""。"""
    if trm == "":
        return ""
    if len(trm) >= 3:
        head = narrow(vbx.left(trm, 2).upper())
        if head in PACKETS:
            return head
    return ""


def conv(trm: str) -> str:
    """補助.Conv (補助:380-392)。DT は `Mid(Trm, 9, 11)`、#L・#〆 は 3 文字目から。ほかはそのまま。

    例: "DT10002+0000021.712M" → "0000021.712"、"#L395.8" → "395.8"
    """
    m = maker(trm)
    if m == "DT":
        return vbx.mid(trm, 9, 11)
    if m in ("#L", "#〆"):
        return vbx.mid(trm, 3, len(trm))
    return trm


# ------------------------------------------------------------------
# 次に入れる欄の表(リファクタリング 第3段。docs/測定値のつなぎ方_DT_L.md 3〜5 章)
# ------------------------------------------------------------------
# VBA は「入れる先」(PucketMove)・「〜検査済」(AllCheck)・黄色の印(Look_Blank / Look_S_Blank / Smallnumber_Blank)を
# 作業ごとに別々に書いていて、1 つだけ直すと食い違った(B60・B90・B91・B93)。ここではどれも次の 1 つの表から決める。
#
#   1. 作業(と薄板の検番・山数、少枚数の検査回数)ごとに、行の種類 → その行で値を入れる位置(下の表)
#   2. 行の範囲: DT = 厚の行、#L = 幅・丈・対角の行、#〆 と黄色の印 = 両方
#   3. 欄の順: 位置ごとに(S の行を上から、次に U の行を上から …)。表に無い位置・行は飛ばす
#   4. 使える欄 = 空で、出ていて(Visible)使える(Enabled)欄
#   入れる先 = 範囲の中で最初に使える欄 / 「〜検査済」= 範囲に使える欄が 1 つも無い / 黄色の印 = 全部の範囲で最初に使える欄

ORDER = "SUCDEF"
#: 行の種類 → その行で値を入れる位置。厚 = 1〜4 行、幅丈 = 5〜8 行(極厚は 5〜14 行)、対角 = 9〜10 行
NEXT_CELL_TABLE: dict[str, dict[str, str]] = {
    # 行 5〜10 の U・D は無い(厚だけの位置 S1・E1)
    "厚板": {"厚": "SUCDE", "幅丈": "SCE", "対角": "SCE"},
    # 厚は 1-上(S)・1-下(C)。U・D は幅丈の 5〜8 行だけ(B90・B91)。E は無い
    "スラブ": {"厚": "SC", "幅丈": "SUCD", "対角": "SC"},
    # 行 5〜14(幅1〜6・丈1〜4)。F は無い(B63)
    "極厚": {"厚": "SUCDE", "幅丈": "SUCDE"},
    # 厚板の検番が H でない少枚数(検査回数で絞らない)
    "少枚数": {"厚": "SUCDE", "幅丈": "SUCDE", "対角": "SUCDE"},
}
#: 少枚数(厚板の検番 H)は検査回数 th(AllCheck の th)で位置を絞る。表に無い th(0)は入れる欄が無い(すぐ「検査済」)
SMALL_COUNT_TABLE: dict[int, dict[str, str]] = {
    5: {"厚": "SUCDE", "幅丈": "SCE", "対角": "SCE"},     # 2026-11-08: 行 5〜10 は U・D を隠す(SmallNumberCount)
    4: {"厚": "SUCD", "幅丈": "SCD", "対角": "SCD"},
    3: {"厚": "SUC", "幅丈": "SUC", "対角": "SUC"},
    2: {"厚": "SU", "幅丈": "SU", "対角": "SU"},
    1: {"厚": "S", "幅丈": "S", "対角": "S"},
}
#: 薄板は検番・山数・検入で位置が決まる(どの行も同じ)。上から順に、最初に当たった旗(grid.ken_flags)の位置
#: (旗, 入れる先・検査済 の位置, 黄色の印 の位置)
THIN_TABLE: tuple[tuple[int, str, str], ...] = (
    (6, "SUCDEF", "SUCDEF"),     # 山数 4(6 回検査)
    (3, "SUCDE", "SUCDE"),       # 山数 3(5 回)
    (5, "SUCDE", "SUCDE"),       # 検番なし(入れる先。黄色の印は検番なしを冷延と同じに扱う: Look_S_Blank)
    (4, "SE", "SE"),             # 冷延で検入≦25(2 回)
    (1, "SCE", "SCE"),           # 冷延(3 回)
    # 検番 H だけ(山数 1・2)。VBA は AllCheck・PucketMove のどれにも当たらず、DT・#L・#〆 はすぐ「検査済」なのに、
    # 黄色の印は出ていた(Look_S_Blank の H の枝。G1 §11-8)。2026-11-10(第3段の続き 1。intended I5): 印の欄で測れるよう、
    # 入れる先・検査済も印と同じ位置にする(山数 1・2 は U・D が隠れているので S → C → E)
    (2, "SUCDE", "SUCDE"),
)


def _row_kind(state: KensaState, row: int) -> str:
    if is_thick(row):
        return "厚"
    if state.is_work("極厚") or row <= 8:
        return "幅丈"
    return "対角"


def _thick_rows(state: KensaState) -> range:
    """厚の行。薄板だけ「厚2ヵ所/厚4ヵ所」のボタンで決まる(ボタンが「厚4ヵ所」= いま 2 ヵ所検査 → 1〜2 行)。

    2026-11-09(B93)ほか: ボタンは薄板のもの。厚板・スラブ・極厚・少枚数は、前の薄板の表示が残っていても 1〜4 行。
    """
    if not state.is_work("薄板"):
        return THICK_ROWS
    if state.at_butt_caption == "厚2ヵ所":       # ボタン表示が2ヵ所 = いま 4 回検査
        return THICK_ROWS
    if state.at_butt_caption == "厚4ヵ所":
        return range(1, 3)
    return range(0)


def _size_rows(state: KensaState) -> range:
    return range(5, 15) if state.is_work("極厚") else NORMAL.size_rows


def _packet_rows(state: KensaState, pucket: str) -> list[int]:
    """伝送の種別ごとの行の範囲(PucketMove の a〜n)。DT = 厚、#L = 幅・丈・対角、#〆 = 両方。"""
    if pucket == "DT":
        return list(_thick_rows(state))
    if pucket == "#L":
        return list(_size_rows(state))
    return list(_thick_rows(state)) + list(_size_rows(state))


def _rows_in(state: KensaState, a: int, n: int) -> list[int]:
    """行 a〜n のうち、作業の厚の行と幅・丈・対角の行(薄板の厚2ヵ所検査中は 3・4 行目を見ない)。"""
    return [r for r in list(_thick_rows(state)) + list(_size_rows(state)) if a <= r <= n]


def _columns(state: KensaState, ken_no: str, hooks: Optional[Hooks], *, marker: bool) -> Optional[dict[str, str]]:
    """いまの作業の、行の種類 → 位置(表を引く)。None は作業が無い。

    hooks は少枚数(厚板の検番 H)の検査回数を出すときだけ使う(AllCheck の th。印は検査回数で絞らない)。
    """
    if state.is_work("薄板"):
        f = ken_flags(state, ken_no, blank_in_flag1=marker)
        for flag, cols, mark in THIN_TABLE:
            if f[flag]:
                p = mark if marker else cols
                return {"厚": p, "幅丈": p, "対角": p}
        return {"厚": "", "幅丈": "", "対角": ""}
    if state.is_work("少枚数") and state.work_caption["厚板"] == "H" and not marker:
        th = small_count_th(state, hooks.calculate_sheets, start_from_ken_count=True) if hooks else 0
        return SMALL_COUNT_TABLE.get(th, {})
    for work, table in NEXT_CELL_TABLE.items():
        if state.is_work(work):
            return table
    return None


def next_cells(state: KensaState, rows: list[int], ken_no: str, hooks: Optional[Hooks] = None, *,
               marker: bool = False) -> list[str]:
    """表から、行 rows の欄を見る順に並べる(位置ごとに行を上から)。使えるかどうかは見ない。"""
    cols = _columns(state, ken_no, hooks, marker=marker)
    if not cols:
        return []
    return [f"{p}{r}" for p in ORDER for r in rows
            if p in cols.get(_row_kind(state, r), "") and (p != "F" or r <= 10)]


def _usable(state: KensaState, name: str) -> bool:
    """値を入れられる欄: 空で、出ていて、使える。"""
    c = state.ctl(name)
    return c.value == "" and c.visible and c.enabled


def _first_usable(state: KensaState, names: list[str]) -> Optional[str]:
    return next((n for n in names if _usable(state, n)), None)


def _set_target(state: KensaState, name: str) -> None:
    state.inp_tag = name
    state.tagtext = name


# ------------------------------------------------------------------
# AllCheck / PucketMove
# ------------------------------------------------------------------
def all_check(state: KensaState, a: int, n: int, ken_no: str, temp: str, hooks: Hooks) -> bool:
    """FormRelat.AllCheck (FormRelat:429-715)。行 a〜n に、表の使える欄が残っていれば真。

    temp(VBA の "厚板" / "少枚数" / "")は表の引き分けと同じ(作業と厚板の検番から決まる)ので見ない。
    AllCheck は ATU があれば作業に関係なく CalculateSheets を呼ぶ(FormRelat:436。VBA どおり)。
    """
    if a == 0 or n == 0:
        return False
    small_count_th(state, hooks.calculate_sheets, start_from_ken_count=True)
    return _first_usable(state, next_cells(state, _rows_in(state, a, n), ken_no, hooks)) is not None


def _no_work(state: KensaState) -> bool:
    return state.work == ""


def pucket_move(state: KensaState, pucket: str, ken_no: str, youto: str, hooks: Hooks,
                msgs: Optional[list[Message]] = None) -> None:
    """FormRelat.PucketMove (FormRelat:716-1016)。伝送の種別の行の範囲で、表の最初に使える欄を INP.Tag に入れる。

    使える欄が無ければ「〜検査済」と知らせ、Auto="off"・INP と INP.Tag を空・FlagLabel=1。
    """
    msgs = msgs if msgs is not None else []
    if _no_work(state):
        return
    if pucket not in PACKETS:
        msgs.append(Message("ﾊﾟｹｯﾄ無し:Exit Sub" + "_PucketMove", WARN, source="FormRelat:774"))
        return
    rows = _packet_rows(state, pucket)
    target = None
    if rows:
        small_count_th(state, hooks.calculate_sheets, start_from_ken_count=True)    # AllCheck(FormRelat:436)
        target = _first_usable(state, next_cells(state, rows, ken_no, hooks))
    if target is None:
        msgs.append(Message(pucket + "検査済", INFO, source="FormRelat:792"))   # 知らせ(警告ではない)
        state.auto = Switch.OFF
        state.inp = ""
        state.inp_tag = ""
        state.flag_label = 1
        return
    _set_target(state, target)


# ------------------------------------------------------------------
# 次の空き欄(黄色の印)
# ------------------------------------------------------------------
def lookblank_color(state: KensaState, a: str) -> None:
    """FormRelat.LookBlank_Color (FormRelat:1272-1282)。MPEX.Tag が空なら赤枠・黄色、そうでなければ黒枠・白。"""
    state.tagtext = a
    c = state.ctl(a)
    if state.mpex_tag == "":
        c.border = BORDER_RED
        state.set_color(a, COLOR_YELLOW)
    else:
        c.border = BORDER_BLACK
        state.set_color(a, COLOR_WHITE)


def _marker(state: KensaState, a: int, n: int, ken_no: str) -> Optional[str]:
    """行 a〜n で、表の最初に使える欄。"""
    return _first_usable(state, next_cells(state, _rows_in(state, a, n), ken_no, marker=True))


def look_s_blank(state: KensaState, a: int, n: int, ken_no: str) -> None:
    """FormRelat.Look_S_Blank (FormRelat:1018-1271)。薄板の次の空き欄。Auto が "on" のときだけ。色は LookBlank_Color。"""
    if state.auto != Switch.ON:
        return
    name = _marker(state, a, n, ken_no)
    if name is not None:
        _set_target(state, name)
        lookblank_color(state, name)


def smallnumber_blank(state: KensaState, a: int, n: int, xa: int) -> None:
    """FormRelat.Smallnumber_Blank (FormRelat:1283-1340)。少枚数(枚数 xa が 1〜5 のとき)の次の空き欄。色は LookBlank_Color。"""
    if state.auto != Switch.ON:
        return
    if xa not in (5, 4, 3, 2, 1):
        return
    name = _marker(state, a, n, "")
    if name is not None:
        _set_target(state, name)
        lookblank_color(state, name)


def look_blank(state: KensaState, a: int, n: int) -> None:
    """FormRelat.Look_Blank (FormRelat:1341-1500)。次の空き欄を赤枠・黄色にする。Auto が "on" のときだけ。

    厚板・極厚・スラブ・少枚数(KenCount なし)の印。再測定(NG_Back)は作業に関係なくこれを呼ぶ(薄板も表どおり)。
    """
    if state.auto != Switch.ON:
        return
    name = _marker(state, a, n, ken_no_for_input(state))
    if name is not None:
        _set_target(state, name)
        state.ctl(name).border = BORDER_RED
        state.set_color(name, COLOR_YELLOW)


def next_blank(state: KensaState, ken_no: str) -> None:
    """次の入力欄を探す(INP_Change UF1:4639-4647 / UserForm_Click UF1:4367-4375 の分岐)。

    検番は INP_Change では冷延欄(B3)、UserForm_Click では KenNo 欄の頭 1 文字を渡す。
    """
    if state.is_work("薄板"):
        look_s_blank(state, 1, 10, ken_no)
    elif state.is_work("極厚"):
        look_blank(state, 1, 14)
    elif state.is_work("少枚数") and state.ken_count != "":
        smallnumber_blank(state, 1, 10, vbx.clng(vbx.val(state.ken_count)))
    else:
        look_blank(state, 1, 10)



def ken_no_for_input(state: KensaState) -> str:
    """次に入れる欄の検番(INP_Change UF1:4585 / UserForm_Click UF1:4362)。

    B3: VBA の INP_Change は冷延欄 `Reien` の頭 1 文字(UserForm_Click は `KenNo`)。ここはどちらも KenNo 欄の頭。
    2026-11-10(第3段の続き 3。intended I7): 薄板で KenNo 欄の頭が "0"/空(DB にまだ検査番号が無い)なら、
    Atu_OPselect が最終圧延機から作った頭(state.ken_head。補助.FindPriority)を使う。VBA は作った頭を作業選択ボタンの
    Caption にしか使わず、薄板は「検番なし」(5 回)か、どの位置にも当たらない("0"。すぐ検査済)になっていた。
    """
    head = vbx.left(state.ken_no_caption, 1)
    if head in ("0", "") and state.is_work("薄板") and state.ken_head not in ("0", ""):
        return state.ken_head
    return head


def ken_count_missing(state: KensaState) -> bool:
    """少枚数(厚板の検番 H)で、枚数(KenHBox7)があるのに検査回数(KenCount)が無い(または 0)。

    2026-11-10(第3段の続き 2。intended I6): このとき AllCheck の th は 0 で、DT・#L・#〆 はすぐ「検査済」になり、
    何も測れないまま進めた(VBA。KenCount は手入力で、KenHBox7 から作る決まりは無い: UF1:745-765 / 1696-1712)。
    検査回数を入れるまで測定を止める(`inp_change`)。画面は検査回数の欄を「必須」にする(kensa.js)。
    """
    return (state.is_work("少枚数") and state.work_caption.get("厚板") == "H"
            and state.ken_h[7].value != "" and vbx.val(state.ken_count) < 1)


KEN_COUNT_MISSING = "検査回数を入れてください"


# ------------------------------------------------------------------
# INP_Change
# ------------------------------------------------------------------
_CELL = re.compile(r"^[SUCDEF](\d{1,2})$")


def manual_target_mismatch(state: KensaState, pucket: str) -> bool:
    """手で選んだ入力先(Auto="off" の INP.Tag)が、届いた種類(DT / #L)の行の範囲に無いか(I8)。

    測定の欄(S1〜F14)だけを見る。#〆(手入力)・測定の欄でない入力先は、VBA どおり選んだ欄へ入れる。
    """
    if pucket not in ("DT", "#L") or state.inp_tag == "":
        return False
    m = _CELL.match(state.inp_tag)
    if not m:
        # 2026-11-12: 歪(ZS 等)など測定表でない欄を手で選んだまま DT / #L が届くと、その欄の名前を位置として扱い
        # 止まっていた。測定器の値は測定表の欄にしか入れない → 次に空いた欄へ回す
        return True
    row = int(m.group(1))
    own = is_thick(row) if pucket == "DT" else not is_thick(row)
    return not own


def _msg_error(text: str, n: int, source: str) -> Message:
    """OutTolerance.MsgError(Trm, Lname, n) の表示(1=!、2=×、4=i「注意」)。"""
    level = {1: WARN, 2: ERROR, 4: INFO}.get(n, INFO)
    return Message(text, level, title="注意" if n == 4 else "警告", source=source)


def inp_change(state: KensaState, hooks: Hooks) -> list[Message]:
    """UF1.INP_Change (UF1:4565-4651)。受信した `state.inp` を入力先へ入れ、判定を呼び、次の入力欄を決める。

    中断(MsgBox のあと Exit Sub)は `Stop`。PucketMove の「〜検査済」は知らせて終わり(戻り値の Message)。
    B55: VBA は入力先が決まらず抜けるとき(INP.Tag="")に FlagLabel を 0 のまま残す。
    """
    msgs: list[Message] = []
    state.last_landed = ""
    if state.flag_label != 1:
        return msgs
    if state.inp == "":
        return msgs
    if state.yama_visible and state.yama == "":
        raise Stop("山数の入力がありません" + "\r\n" + "選択してください", level=WARN, source="UF1:4569")
    if state.yama_tag == "別フォーム":
        raise Stop("別フォームで検査してください", level=ERROR, source="UF1:4574")
    if ken_count_missing(state):
        # 2026-11-10(第3段の続き 2): 枚数があって検査回数が無いと、すぐ「検査済」で何も入らなかった。入れるまで止める。
        # 山数と違い INP を空に戻す(検査回数を入れた後、同じ値が届いても受ける)
        state.inp = ""
        raise Stop(KEN_COUNT_MISSING, level=WARN, source="input_router.ken_count_missing")

    state.flag_label = 0
    fm_clear(state, "All")
    state.paked = maker(state.inp)

    ken_no = ken_no_for_input(state)
    youc = state.youto if state.youto != "" else ""
    if state.auto == Switch.ON:
        pucket_move(state, state.paked, ken_no, youc, hooks, msgs)
        # B61: 「ﾊﾟｹｯﾄ無し」でも VBA は前回の INP.Tag の欄で LoopText・Posting まで進む。直す: そこで抜ける
        if state.paked not in PACKETS:
            state.inp = ""
            state.flag_label = 1
            return msgs
    elif manual_target_mismatch(state, state.paked):
        # 2026-11-11(I8。利用者の決定): 手で選んだ欄が届いた種類の行でなければ(DT を幅・丈・対角へ、#L を厚へ)、
        # その欄には入れず、種類の行の範囲で次に空いた欄へ回す。VBA は Auto="off" なら PucketMove を飛ばして選んだ欄へ入れていた
        pucket_move(state, state.paked, ken_no, youc, hooks, msgs)

    if state.inp_tag == "":
        # B55: VBA は FlagLabel=0 のまま抜け、フォームを押すまで次の伝送を受け付けない。直す: ほかの抜け方と同じく戻す
        state.inp = ""
        state.flag_label = 1
        return msgs
    st = state.inp_tag
    if maker(state.inp) in PACKETS:
        state.ctl(st).value = vbx.cstr(vbx.val(conv(state.inp)))

    def block(text: str, n: int, source: str) -> Stop:
        """MsgError(text, Lname, n) を出して値を消し Exit Sub。"""
        state.ctl(st).value = ""
        state.inp = ""
        state.flag_label = 1
        return Stop(text, messages=msgs, level=_msg_error(text, n, source).level, source=source)

    if hooks.register_check(state):
        raise block("検査員登録をして下さい", 2, "UF1:4609")
    if _no_work(state):
        raise block("作業選択してください", 1, "UF1:4615")
    if state.proofread_flag != "ok":
        if state.work_button_tag == "jump" or state.dbflag != "":
            pass
        else:
            hooks.proofread_load(state)                       # 校正画面を開いて抜ける(MsgBox は無い)
            state.ctl(st).value = ""
            state.inp = ""
            state.flag_label = 1
            return msgs

    state.last_landed = st          # ここから先は入力先に入った(LoopText の異常値で消されても入った後)
    hooks.loop_text(state, st)
    hooks.posting(state)
    hooks.out_tolerance(state)

    state.auto = Switch.ON
    state.load = Switch.ON
    next_blank(state, ken_no)
    fm_clear(state, "Click")
    state.inp = ""
    state.flag_label = 1
    return msgs


def receive(state: KensaState, text: str, hooks: Hooks) -> list[Message]:
    """UF1.StartTB_KeyDown (UF1:3295-3306)。`INP = StartTB`。値が変わったときだけ INP_Change が走る(MSForms の Change)。

    FlagLabel が 1 でないとき・StartTB が空のときは INP に入れずに抜ける(UF1:3297-3298)。
    値が入った欄は `state.last_landed` に残る(入らなければ "")。
    INP は前の値のまま残るので、その後に同じ文字列が来ても Change は起きない(VBA どおり)。
    """
    state.last_landed = ""
    if state.flag_label != 1:
        return []
    if text == "":
        return []
    return _set_inp(state, text, hooks)


def _set_inp(state: KensaState, text: str, hooks: Hooks) -> list[Message]:
    """`UF1.INP = text`。値が変わったときだけ INP_Change が走る(MSForms の Change)。"""
    if state.inp == text:
        return []
    state.inp = text
    return inp_change(state, hooks)


def manual_entry(state: KensaState, manual: str, hooks: Hooks) -> list[Message]:
    """UF1.CommandButton3_Click (UF1:4552-4558)。手入力: `INP = "#〆" & Manual`(INP が 3 文字以下なら "0.000" に整える)。

    StartTB_KeyDown と違い FlagLabel を見ずに INP へ入れる(止めるのは INP_Change の FlagLabel の確認)。
    """
    if len(state.inp) >= 4:
        text = "#〆" + manual
    else:
        text = "#〆" + vbx.fmt(manual, "0.000")
    return _set_inp(state, text, hooks)


# ------------------------------------------------------------------
# 手で入力先を選ぶ・消す(LabelCls)
# ------------------------------------------------------------------
def tag_set(state: KensaState, name: str) -> str:
    """FormRelat.TagSet (FormRelat:1505-1509)。TextBox に入ったとき、そこを入力先にする。"""
    _set_target(state, name)
    return name


def text_tag(state: KensaState, name: str, hooks: Hooks) -> str:
    """FormRelat.TextTag (FormRelat:1510-1529)。測定ラベルのクリック(LabelCls ObjX_Click)。

    自動の入力先探しを止め(Auto="off")、再判定してから、そのラベルを赤枠・黄色にして入力先にする。
    黄色は判定色を上書きする(VBA どおり)。
    """
    state.auto = Switch.OFF
    with state.stopped():
        hooks.loop_judg(state)
        hooks.hiz_loop_judg(state)
        state.ctl(name).border = BORDER_RED
        state.set_color(name, COLOR_YELLOW)
        _set_target(state, name)
    return name


def diagonal_erase(state: KensaState, name: str) -> None:
    """FormRelat.DiagonalErase (FormRelat:1530-1547)。極厚以外で行 9・10 を消したら、その位置の対角差を消す。"""
    if state.is_work("極厚"):
        return
    last = vbx.right(name, 1)
    if last in ("9", "0"):
        p = vbx.left(name, 1)
        state.ctl(diag_name(p)).value = ""
        state.set_color(p + "10", COLOR_DEFAULT)
        state.set_color(p + "9", COLOR_DEFAULT)


def ng_erase(state: KensaState, name: str) -> None:
    """FormRelat.NGErase (FormRelat:1548-1600)。その位置の列に NG 色が残っていなければ NG? を空にする。

    B87: VBA は S..E の行 1〜12 しか見ない(極厚の行 13・14 の NG を見落とす)。
    """
    p = vbx.left(name, 1)
    count = 10 if p == "F" else 12
    if p != "F":
        count = 14            # B87: 極厚の丈3・4(行13・14)も見る
    fl = False
    for i in range(1, count + 1):
        c = state.ctl(f"{p}{i}")
        if c.value != "" and state.is_ng(f"{p}{i}"):
            fl = True
            break
    if name.startswith(("Lateral", "Z", "G")) or (name.startswith("Diag") and name.endswith("Kou")):
        if name in ("LateralS", "LateralU", "LateralC", "LateralD", "LateralE", "LateralF",
                    "DiagSKou", "DiagUKou", "DiagCKou", "DiagDKou", "DiagEKou", "DiagFKou",
                    "ZS", "ZU", "ZC", "ZD", "ZE", "ZF", "GS", "GU", "GC", "GD", "GE", "GF"):
            if state.is_ng(name):
                fl = True
    if not fl:
        state.ctl("NG" + p).value = ""


def erase_count(state: KensaState, name: str, hooks: Hooks) -> None:
    """FormRelat.EraseCount (FormRelat:6911-6935)。消した座標(DeleteP)と、同じ座標を続けて消した回数(DeleteC)。

    厚板・薄板・少枚数では行 9 と 10 を同じ座標 "?10" とみなす。
    """
    temp = name
    letter = vbx.left(name, 1)
    digit = 0
    if len(name) == 3:
        digit = vbx.clng(vbx.right(name, 2))
    if len(name) == 2:
        digit = vbx.clng(vbx.right(name, 1))
    if state.is_work("厚板") or state.is_work("薄板") or state.is_work("少枚数"):
        if digit in (9, 10):
            temp = letter + "10"
    if state.delete_p == temp:
        state.delete_c = vbx.cstr(vbx.val(state.delete_c) + 1)
    else:
        state.delete_c = "1"
        state.delete_p = temp
    hooks.posting(state)


@dataclass
class EraseResult:
    """ダブルクリック消去の結果。`reason` が真なら消去理由の入力(ConfiReason_Open)が要る。"""
    target: str
    old_value: str
    reason: bool
    pos: str = ""
    num: int = 0
    count: float = 0.0


def cell_dbl_click(state: KensaState, name: str, hooks: Hooks) -> EraseResult:
    """LabelCls.ObjX_DblClick (LabelCls:166-196 ほか)。測定ラベルのダブルクリックで消す。

    MSForms はダブルクリックの前に Click(= `text_tag`)が来るので、画面側は `text_tag` → これ の順に呼ぶ。
    消す値は INP.Tag の欄の値(VBA どおり。Click の後なら `name` と同じ)。
    値があったとき、DeleteC が 1 以上(= 毎回)で理由の入力を求める(2024.2.6 品証指示。コメントの「2回連続」ではない)。
    """
    st = state.inp_tag
    temp = state.ctl(st).value
    hooks.time_posting(state, st, "")
    diagonal_erase(state, name)
    c = state.ctl(name)
    c.value = ""
    c.tag = ""
    erase_count(state, name, hooks)
    state.set_color(name, COLOR_DEFAULT)
    ng_erase(state, name)

    pos = vbx.left(st, 1)
    num = 0
    if len(st) == 2:
        num = vbx.clng(vbx.right(st, 1))
    elif len(st) == 3:
        num = vbx.clng(vbx.right(st, 2))
    count = vbx.val(state.delete_c)
    reason = temp != "" and count >= 1
    if reason:
        hooks.confi_reason_open(state, pos, num, state.lname, temp, count)
    return EraseResult(st, temp, reason, pos, num, count)


# ------------------------------------------------------------------
# 外観不良枚数
# ------------------------------------------------------------------
def bad_mouse_down(state: KensaState, pos: str, hooks: Hooks) -> None:
    """UF1.SBad_MouseDown..FBad_MouseDown (UF1:1809-1834)。不良枚数を "0" に整えて色を付ける。

    B4: F の欄は VBA では `UF1.EBad = Format(UF1.FBad, "0")`(E 欄に書く)。
    """
    src = state.bad[pos]
    text = vbx.fmt(src.value, "0")
    dest = pos
    state.bad[dest].value = text if vbx.is_numeric(text) else ""     # ?Bad_Change: 数でなければ空
    hooks.te_color(state, bad_name(pos))
