"""作業選択・山数・検入枚数・厚2/4ヵ所 → 見える位置・行・位置名(FormRelat / UF1 の画面制御)

VBA の LSelect / AtCountSelect / InspectionSelect / Six_Display / SlabSet / Yama_Change / Atu_OPselect などを、
`grid.KensaState` を書き換える関数にした。ピクセル位置・フォームの大きさ(Left/Top/Width/Height)は移さない
(Web 画面は自前で並べる)。見える/見えない・有効/無効・Caption・判定に効く値は VBA と同じにする。

DB(仕掛)は `Hooks.arrbox`、品質基準の照合は `Hooks.db_return_item` / `Hooks.is_gauge_required` で受け取る。
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional

from . import paths, vbx
from .model import FormTag, NORMAL, Switch, THICK_ROWS, bh_name, diag_name, kou_name, lateral_name, thick_format
from .flow import ERROR, INFO, Message
from .grid import (THIN_PLATE_ATU, WORKS, Hooks, KensaState, ken_flags, small_count_th, typed_cmp,
                   variant_eq)

#: VBA の On Error Resume Next で飛ばされる実行時エラーに当たる Python の例外
VBA_ERRORS = (vbx.TypeMismatch, KeyError, IndexError, ValueError, TypeError, ZeroDivisionError, OverflowError)

#: VBA の型・添字の誤り(VBA でも On Error Resume Next で飛ばしたもの)。これ以外を飛ばしたときは
#: Python のプログラムの誤りのおそれがあるので、ログに残して見えるようにする(2026-11-06 の見直し)
_VBA_LIKE = (vbx.TypeMismatch, IndexError, ZeroDivisionError, OverflowError)


def vba_swallow(e: BaseException, where: str = "") -> None:
    """On Error Resume Next で飛ばした誤りを記録する(動きは変えない)。"""
    import logging
    lg = logging.getLogger("ufkensa.kensa.vba_swallow")
    if isinstance(e, _VBA_LIKE):
        lg.debug("飛ばした(VBA と同じ): %s: %s %s", type(e).__name__, e, where)
    else:
        lg.warning("飛ばした誤り(プログラムの誤りのおそれ): %s: %s %s", type(e).__name__, e, where, exc_info=e)

#: ライン OptionButton → Caption(G6 UF1 WorkFrame)
LINE_CAPTIONS = {"L1Work": "L-1", "LVC": "LVC", "HVC": "HVC", "Tott": "ﾄｯﾄ", "Bala": "ﾊﾞﾗﾝｻｰ", "Test": "Test", "Maru": "丸徳"}
#: 丸徳の号機 A1..A11 → Caption(G6 UF1 NoFrame)
MARU_NO_CAPTIONS = {1: "A22", 2: "A21", 3: "A20", 4: "A19", 5: "A18", 6: "場外", 7: "A16", 8: "A15", 9: "A14",
                    10: "A13", 11: "B5"}

FIVE = ("S", "U", "C", "D", "E")


# ------------------------------------------------------------------
# 小さな道具
# ------------------------------------------------------------------
def _show(state: KensaState, name: str, visible: Optional[bool] = None, enabled: Optional[bool] = None) -> None:
    c = state.ctl(name)
    if visible is not None:
        c.visible = visible
    if enabled is not None:
        c.enabled = enabled


def _cells(state: KensaState, rows, cols, visible: Optional[bool] = None, enabled: Optional[bool] = None) -> None:
    for i in rows:
        for p in cols:
            _show(state, f"{p}{i}", visible, enabled)


def _caps(state: KensaState, **caps: str) -> None:
    """列見出し・歪見出し・OPS/OPE の Caption を、書いた分だけ設定する(キーは S, FS, OPS1 など)。"""
    for k, v in caps.items():
        state.v(k).value = v


def _cell1(rows: Optional[list[Any]], col: int) -> Any:
    """VBA の `Arr(1, col)`(1 始まり)。Empty・列 0 はエラー(呼び手の On Error に任せる)。"""
    if rows is None:
        raise vbx.TypeMismatch("Empty の配列を参照")
    if col < 1:
        raise IndexError("インデックスが有効範囲にありません")
    return rows[0][col - 1]


def _field_no(fields: Optional[list[str]], name: str, current: int = 0) -> int:
    """`For i = 1 To UBound(F, 2): Select Case F(1, i): Case name: no = i`(最後に当たった列。無ければ元のまま)。"""
    if fields is None:
        return current
    for i, f in enumerate(fields, start=1):
        if f == name:
            current = i
    return current


def _csng(text: Any) -> float:
    """Single への代入(`Dim x As Single: x = 文字列`)。数にならない文字列は型エラー 13。"""
    return vbx.single(vbx.cdbl(text))


def find_priority(s: str) -> str:
    """補助.FindPriority (補助:1455-1490)。設備コースの最後の圧延機から検査番号の頭を作る。

    後ろから CM1→L、CM2→N、CM3→K、2HI→T を探し、無ければ HOT→H。どれも無ければ ""。
    """
    words = s.split(" ") if s != "" else []
    priority = ""
    for w in reversed(words):
        for key in ("CM1", "CM2", "CM3", "2HI"):
            if key == w:
                priority = key
                break
        if priority != "":
            break
    if priority == "":
        for w in words:
            if w == "HOT":
                priority = "HOT"
                break
    return {"CM1": "L", "CM2": "N", "CM3": "K", "2HI": "T", "HOT": "H"}.get(priority, priority)


def lookup_ken_no(state: KensaState, hooks: Hooks) -> str:
    """検査番号の頭 1 文字(AtCountSelect:1954-1985 / InspectionSelect:2187-2220 / LSelect の非キャッシュ部)。

    SIKAHIKINOW の [ﾛｯﾄ番号]=Lot、無ければ SIKALOTNOW。どちらにも無ければ ""。
    Lot が空なら絞り込み無し(Arrbox の仕様)で先頭行の検査番号になる(VBA どおり)。
    """
    fields = hooks.arrbox(state, "SIKAHIKINOW.accdb", "", "Fields")
    no = _field_no(fields, "検査番号")
    ken_acc = hooks.arrbox(state, "SIKAHIKINOW.accdb", state.lot, "Record")
    if ken_acc is None:
        fields = hooks.arrbox(state, "SIKALOTNOW.accdb", "", "Fields")
        no = _field_no(fields, "検査番号", no)
        ken_acc = hooks.arrbox(state, "SIKALOTNOW.accdb", state.lot, "Record")
    if ken_acc is not None:
        return vbx.left(_cell1(ken_acc, no), 1)
    return ""


def clear_lselect_cache(state: KensaState) -> None:
    """FormRelat.ClearLSelectCache (FormRelat:2568-2577)。新規検査で呼ぶ。"""
    c = state.cache
    c.lot_no = ""
    c.ken_no = ""
    c.cours = ""
    c.ord_atu = 0.0
    c.c_course = ""
    c.skip_lot_not_found_msg = True


# ------------------------------------------------------------------
# 表示桁・公差の上下限
# ------------------------------------------------------------------
def num_dig(state: KensaState) -> None:
    """FormRelat.NumDig (FormRelat:345-428)。測定値の表示桁(LoopText の最後に呼ぶ)。

    厚(行1〜4): ATU > 74 なら "0.00"、ほかは "0.000"。
    幅(行5〜6): "0.0"。
    丈(S..E は行7〜12、F は 7〜10)・対角・直角度: "0.0"。
    """
    if state.atu == "":
        return
    pat = thick_format(state.atu)
    for _key, c in state.measures.thick_ctls():
        c.value = vbx.fmt(c.value, pat)
    pat = "0.0"
    for i in NORMAL.width:
        for p in FIVE + ("F",):
            c = state.cells[f"{p}{i}"]
            c.value = vbx.fmt(c.value, pat)
    for i in range(7, 13):
        for p in FIVE:
            c = state.cells[f"{p}{i}"]
            c.value = vbx.fmt(c.value, "0.0")
    for i in range(7, 11):
        c = state.cells[f"F{i}"]
        c.value = vbx.fmt(c.value, "0.0")
    for group in (state.diag, state.diag_kou):
        for p in FIVE + ("F",):
            group[p].value = vbx.fmt(group[p].value, "0.0")


def form_formula(state: KensaState) -> None:
    """FormRelat.Form_Formula (FormRelat:6169-6184)。公差の手修正中(ReSize.Tag="有")だけ上下限を計算する。"""
    if state.resize_tag != "有":
        return
    v = vbx.val
    if state.atu != "" and state.ap != "":
        state.apk = vbx.fmt(v(state.atu) + v(state.ap), "0.000")
    if state.atu != "" and state.am != "":
        state.amk = vbx.fmt(v(state.atu) - v(state.am), "0.000")
    if state.hab != "" and state.hp != "":
        state.hpk = vbx.fmt(v(state.hab) + v(state.hp), "0.0")
    if state.hab != "" and state.hm != "":
        state.hmk = vbx.fmt(v(state.hab) - v(state.hm), "0.0")
    if state.tak != "" and state.tp != "":
        state.tpk = vbx.fmt(v(state.tak) + v(state.tp), "0.0")
    if state.tak != "" and state.tm != "":
        state.tmk = vbx.fmt(v(state.tak) - v(state.tm), "0.0")


# ------------------------------------------------------------------
# 歪フレーム・4 山・スラブ
# ------------------------------------------------------------------
def dist_select(state: KensaState, n: int) -> None:
    """FormRelat.DistSelect (FormRelat:6600-6617)。歪フレームの数(25=2回、26=3回、ほか=5回)。"""
    if n == 25:
        for p, vis in (("U", False), ("C", False), ("D", False)):
            _show(state, "FM" + p, vis)
    elif n == 26:
        for p, vis in (("U", False), ("C", True), ("D", False)):
            _show(state, "FM" + p, vis)
    else:
        for p in FIVE:
            _show(state, "FM" + p, True)


def six_display(state: KensaState, n: str, hooks: Hooks) -> None:
    """FormRelat.Six_Display (FormRelat:7024-7211)。4 山(6 回検査)の表示。"on" で F 列を出す。"""
    if n == "on":
        _cells(state, range(1, 11), FIVE, True, True)
        for p in FIVE:
            _show(state, diag_name(p), True, True)
            _show(state, "FM" + p, True)
        for i in list(range(1, 6)) + list(range(7, 12)):
            _show(state, f"LC{i}", True)
        for p in FIVE:
            _show(state, f"{p}_UpVisual", True)
            _show(state, f"{p}_UnderVisual", True)
        if state.youto != "R115":
            _show(state, "L10", False)
            for p in FIVE:
                _show(state, kou_name(p), False)
        # SkipA:
        for i in range(1, 11):
            _show(state, f"F{i}", True)
        for name in ("OPS6", "OPE6", "FBad", "LC6", "LC12", "F_UnderVisual", "F_UpVisual", "FMF", "FF", "NGF", "F",
                     "TSF", "TEF", "CheckTF"):
            _show(state, name, True)
        _show(state, "DiagF", True, True)
        state.inspe_butt_visible = False
        _caps(state, S="4-4-S", U="4-4-E", C="4-3-2", D="4-2-2", E="4-1-S", F="4-1-E",
              FS="4-4-S", FU="4-4-E", FC="4-3-2", FD="4-2-2", FE="4-1-S", FF="4-1-E",
              OPS1="4-4-S", OPS2="4-4-E", OPS3="4-3-2", OPS4="4-2-2", OPS5="4-1-S", OPS6="4-1-E",
              OPE1="4-4-S", OPE2="4-4-E", OPE3="4-3-2", OPE4="4-2-2", OPE5="4-1-S", OPE6="4-1-E")
        if state.at_butt_caption == "厚2ヵ所":
            at_count_select(state, 4, hooks)
        elif state.at_butt_caption == "厚4ヵ所":
            at_count_select(state, 2, hooks)
    else:
        for i in range(1, 11):
            _show(state, f"F{i}", False)
        for name in ("OPS6", "OPE6", "FBad", "LC6", "LC12", "F_UnderVisual", "F_UpVisual", "FMF", "FF", "NGF", "F",
                     "TSF", "TEF", "CheckTF", "LateralF", "DiagFKou"):
            _show(state, name, False)
        _show(state, "DiagF", False, False)
        if state.is_work("極厚"):
            for p in FIVE:
                _show(state, diag_name(p), False, False)
            _show(state, "Loading", False)
            return
    _show(state, "Loading", False)


def slab_set(state: KensaState, hooks: Hooks) -> None:
    """FormRelat.SlabSet (FormRelat:5104-5257)。スラブの表示(1-上 / 1-下 / 2-上 / 2-下)。"""
    six_display(state, "off", hooks)
    for i in THICK_ROWS:
        _show(state, f"AL{i}", True)
        for p, on in (("S", True), ("U", False), ("C", True), ("D", False), ("E", False)):
            _show(state, f"{p}{i}", on, on)
    for i in NORMAL.size_rows:
        if i <= 8:
            pat = (("S", True), ("U", True), ("C", True), ("D", True), ("E", False))
        else:
            pat = (("S", True), ("U", False), ("C", True), ("D", False), ("E", False))
        for p, on in pat:
            _show(state, f"{p}{i}", on, on)
    for i in list(range(1, 6)) + list(range(7, 12)):
        _show(state, f"LC{i}", True)
    for i in range(7, 15):
        if i in (7, 8, 12, 13):
            _show(state, f"L{i}", False)
        elif i in (9, 10):
            _show(state, f"U{i}", False, False)
            _show(state, f"D{i}", False, False)
    _cells(state, range(11, 15), FIVE, False, False)
    _caps(state, L1="幅1", L2="幅2", L3="丈1", L4="丈2", L5="対角A", L6="対角B")
    state.at_count = ""
    state.at_butt_visible = False
    state.inspe_butt_visible = False
    state.yama_visible = False
    _show(state, "YamaLabel", False)
    _caps(state, S="1-上", U="1-下", C="2-上", D="2-下", E="", FS="1", FU="", FC="2", FD="", FE="",
          OPS1="1", OPS2="1", OPS3="2", OPS4="2", OPS5="", OPE1="1", OPE2="1", OPE3="2", OPE4="2", OPE5="")
    for name in ("S_UpVisual", "C_UpVisual", "S_UnderVisual", "C_UnderVisual", "FMS", "FMC"):
        _show(state, name, True)
    for name in ("FMU", "FMD", "FME", "U_UpVisual", "U_UnderVisual", "D_UpVisual", "D_UnderVisual", "E_UpVisual",
                 "E_UnderVisual", "LC2", "LC8"):
        _show(state, name, False)
    for i in range(4, 7):
        _show(state, f"LC{i}", False)
        _show(state, f"LC{i + 6}", False)
    for name in ("TSE", "TEE", "E", "NGE"):
        _show(state, name, False)
    for p in ("U", "D", "E"):
        _show(state, diag_name(p), False, False)


# ------------------------------------------------------------------
# 厚2/4ヵ所・検入≦25/≧26
# ------------------------------------------------------------------
def at_count_select(state: KensaState, n: int, hooks: Hooks) -> None:
    """FormRelat.AtCountSelect (FormRelat:1950-2181)。薄板の厚測定箇所 2/4 の切替(行 3・4 の表示)。

    B59: 検番 T・山数 4 のとき F 列は Visible を 2 回設定し Enabled を設定しない(VBA。G1 §0-3-6)。
    直す: 2 回目の Visible を Enabled の意図とみなし、F3/F4 も Enabled にする。
    """
    ken_no = lookup_ken_no(state, hooks)
    f = ken_flags(state, ken_no)
    if n == 2:
        state.at_count = "厚2ヵ所検査中"
        state.at_butt_caption = "厚4ヵ所"
        on = False
    elif n == 4:
        state.at_count = "厚4ヵ所検査中"
        state.at_butt_caption = "厚2ヵ所"
        on = True
    else:
        on = None
    if on is not None:
        for i in (3, 4):
            _show(state, f"AL{i}", on)
            if f[6]:
                cols = FIVE + ("F",)
            elif f[4]:
                cols = ("S", "E")
            elif f[3]:
                cols = FIVE
            elif f[1]:
                cols = ("S", "C", "E")
            elif f[2]:
                cols = FIVE
            else:
                cols = ()
            _cells(state, [i], cols, on, on)
    if ken_no == "T":
        for i in (3, 4):
            _show(state, f"AL{i}", True)
            if f[6]:
                _cells(state, [i], FIVE + ("F",), True)
                _cells(state, [i], FIVE + ("F",), enabled=True)
            elif f[4]:
                _cells(state, [i], ("S", "E"), True, True)
            elif f[3]:
                _cells(state, [i], FIVE, True, True)
            else:
                _cells(state, [i], ("S", "C", "E"), True, True)


def _b5_2yama_ops5() -> str:
    """B5: InspectionSelect の 2 山の OPS5 は VBA では "2-2-E"(他の箇所は "2-1-E")。"""
    return "2-1-E"


def inspection_select(state: KensaState, n: str, hooks: Hooks) -> None:
    """FormRelat.InspectionSelect (FormRelat:2183-2494)。薄板の検入枚数(≦25 = 2 回 / ≧26 = 3 回)の切替。

    "非表示" = 検入≦25(C 列を消す。InspeButt は "検入≧26" を表示)、"表示" = 検入≧26。
    """
    ken_no = lookup_ken_no(state, hooks)
    f = ken_flags(state, ken_no)
    yama = state.yama
    if n == "非表示":
        state.inspe_butt_caption = "検入≧26"
        _cells(state, range(1, 11), ("C",), False, False)
        _show(state, "DiagC", False)
        _caps(state, FC="", C="")
        if variant_eq(yama, 1):
            _caps(state, S="S", C="", E="E", FS="S", FC="", FE="E", OPS1="S", OPS3="", OPS5="E",
                  OPE1="S", OPE3="", OPE5="E", U="", FU="", D="", FD="")
        elif variant_eq(yama, 2):
            _caps(state, S="2-2-S", U="", C="", D="", E="2-1-E", FS="2-2-S", FU="", FC="", FD="", FE="2-1-E",
                  OPS1="2-2-S", OPS2="", OPS3="2-2-E", OPS4="", OPS5=_b5_2yama_ops5(),
                  OPE1="2-2-S", OPE2="", OPE3="2-2-E", OPE4="", OPE5="2-1-E")
            _caps(state, U="", FU="", D="", FD="")
        elif variant_eq(yama, 3):
            _caps(state, S="3-3-S", U="", C="", D="", E="3-1-E", FS="3-3-S", FU="", FC="", FD="", FE="3-1-E",
                  OPS1="3-3-S", OPS2="", OPS3="", OPS4="", OPS5="3-1-E",
                  OPE1="3-3-S", OPE2="", OPE3="", OPE4="", OPE5="3-1-E")
        dist_select(state, 25)
        for i in (2, 3, 4, 8, 9, 10):
            _show(state, f"LC{i}", False)
        for name in ("U_UpVisual", "C_UpVisual", "D_UpVisual", "U_UnderVisual", "C_UnderVisual", "D_UnderVisual",
                     "UBH", "CBH", "DBH"):
            _show(state, name, False)
        if state.sumisei != "" or state.lateral_text_tag == "有":
            for p in ("U", "C", "D"):
                _show(state, lateral_name(p), False)
    elif n == "表示":
        state.inspe_butt_caption = "検入≦25"
        if ken_no == "T" or state.at_butt_caption == "厚2ヵ所":
            _cells(state, range(1, 11), ("C",), True, True)
        elif state.at_butt_caption == "厚4ヵ所":
            for i in range(1, 11):
                on = i not in (3, 4)
                _show(state, f"C{i}", on, on)
        if f[6]:
            pass                                       # 歪表示 6 回は別関数
        elif f[1]:
            dist_select(state, 26)
            for i, on in ((2, False), (4, False), (3, True), (8, False), (10, False), (9, True)):
                _show(state, f"LC{i}", on)
            for name, on in (("U_UpVisual", False), ("D_UpVisual", False), ("C_UpVisual", True),
                             ("U_UnderVisual", False), ("D_UnderVisual", False), ("C_UnderVisual", True),
                             ("UBH", False), ("CBH", True), ("DBH", False)):
                _show(state, name, on)
            if state.sumisei != "" or state.lateral_text_tag == "有":
                for p, on in (("U", False), ("C", True), ("D", False)):
                    _show(state, lateral_name(p), on)
        if ken_no == "T":
            dist_select(state, 26)
        if f[3] or f[2]:
            dist_select(state, 5)
            for i in (2, 3, 4, 8, 9, 10):
                _show(state, f"LC{i}", True)
            # B58: VBA は C_UpVisual を 2 回表示にして C_UnderVisual を触らない(FormRelat:2375-2380)
            c_under = "C_UnderVisual"
            for name in ("U_UpVisual", "C_UpVisual", "D_UpVisual", "U_UnderVisual", c_under, "D_UnderVisual",
                         "UBH", "CBH", "DBH"):
                _show(state, name, True)
        if variant_eq(yama, 1):
            _caps(state, S="S", C="2", E="E", FS="S", FC="2", FE="E", OPS1="S", OPS3="2", OPS5="E",
                  OPE1="S", OPE3="2", OPE5="E", U="", FU="", D="", FD="")
        elif variant_eq(yama, 2):
            _caps(state, S="2-2-S", U="", C="2-2-E", D="", E="2-1-E", FS="2-2-S", FU="", FC="2-2-E", FD="",
                  FE="2-1-E", OPS1="2-2-S", OPS2="", OPS3="2-2-E", OPS4="", OPS5=_b5_2yama_ops5(),
                  OPE1="2-2-S", OPE2="", OPE3="2-2-E", OPE4="", OPE5="2-1-E")
            _caps(state, U="", FU="", D="", FD="")
        elif variant_eq(yama, 3):
            _caps(state, S="3-3-S", U="3-3-E", C="3-2-2", D="3-1-S", E="3-1-E",
                  FS="3-3-S", FU="3-3-E", FC="3-2-2", FD="3-1-S", FE="3-1-E",
                  OPS1="3-3-S", OPS2="3-3-E", OPS3="3-2-2", OPS4="3-1-S", OPS5="3-1-E",
                  OPE1="3-3-S", OPE2="3-3-E", OPE3="3-2-2", OPE4="3-1-S", OPE5="3-1-E")
        _show(state, "DiagC", True)


def at_butt_click(state: KensaState, hooks: Hooks) -> None:
    """UF1.AtButt_Click (UF1:4469-4480)。On Error Resume Next(途中のエラーは握りつぶす)。"""
    try:
        if state.at_butt_caption == "厚2ヵ所":
            at_count_select(state, 2, hooks)
        elif state.at_butt_caption == "厚4ヵ所":
            at_count_select(state, 4, hooks)
    except VBA_ERRORS as _vba_e:
        vba_swallow(_vba_e, "layout.py at_butt_click")
        pass


def inspe_butt_click(state: KensaState, hooks: Hooks) -> None:
    """UF1.InspeButt_Click (UF1:4481-4492)。On Error Resume Next。"""
    try:
        if state.inspe_butt_caption == "検入≧26":
            inspection_select(state, "表示", hooks)
        elif state.inspe_butt_caption == "検入≦25":
            inspection_select(state, "非表示", hooks)
    except VBA_ERRORS as _vba_e:
        vba_swallow(_vba_e, "layout.py inspe_butt_click")
        pass


# ------------------------------------------------------------------
# 直角度・ラテラルボー・引当
# ------------------------------------------------------------------
def right_angle_display(state: KensaState, n: str, count: str = "") -> None:
    """FormRelat.RightAngleDisplay (FormRelat:2495-2566)。直角度欄の表示。n = 検番 / "強制" / "新規"。"""
    if state.sumisei != "":
        _show(state, "L10", True)
        for p in FIVE:
            _show(state, kou_name(p), True)
    if variant_eq(state.yama, 4):
        _show(state, "DiagFKou", True)
    if not variant_eq(state.yama, 4):
        _show(state, "DiagFKou", False)
    if count != "":
        _show(state, "DiagFKou", True)
    if n in ("N", "K", "T", "S", "F"):
        if variant_eq(state.yama, 3):
            if state.youto == "R115":
                _show(state, "DiagUKou", True)
                _show(state, "DiagDKou", True)
            _show(state, "DiagU", True, True)
            _show(state, "DiagD", True, True)
        elif variant_eq(state.yama, 2) or variant_eq(state.yama, 1):
            if state.youto == "R115":
                _show(state, "DiagUKou", False)
                _show(state, "DiagDKou", False)
            _show(state, "DiagU", False, False)
            _show(state, "DiagD", False, False)
    elif n == "強制":
        _show(state, "L10", True)
        for p in FIVE:
            _show(state, kou_name(p), True)
    elif n == "新規":
        _show(state, "L10", False)
        for p in FIVE:
            _show(state, kou_name(p), False)


def lselect_lateral(state: KensaState, n: str, count: str = "") -> None:
    """FormRelat.LSelectLateral (FormRelat:5335-5421)。ラテラルボー欄の表示。

    極厚は Additional を "ﾗﾃﾗﾙﾎﾞｰ検査不要_耳付き" にして終わり。Additional に "ﾗﾃﾗﾙﾎﾞｰ検査不要" があれば何もしない。
    ATU・Hab・Tak を Single に入れるので、数にならないと型エラー(途中で止まる。VBA どおり)。
    """
    if state.is_work("極厚"):
        state.additional = "ﾗﾃﾗﾙﾎﾞｰ検査不要_耳付き"
        return
    if vbx.instr(state.additional, "ﾗﾃﾗﾙﾎﾞｰ検査不要") > 0:
        return
    if state.atu == "":
        return
    _csng(state.atu)
    _csng(state.hab)
    _csng(state.tak)
    _show(state, "L11", True)
    for p in FIVE:
        _show(state, lateral_name(p), True)
    _show(state, "LateralF", state.yama == "4")
    if count != "":
        _show(state, "LateralF", True)
    if n == "H":
        _show(state, "LateralU", False)
        _show(state, "LateralD", False)
    elif n in ("N", "K", "T", "S", "F"):
        if variant_eq(state.yama, 2) or variant_eq(state.yama, 1):
            _show(state, "LateralU", False)
            _show(state, "LateralD", False)
    elif n == "強制":
        _show(state, "L11", True)
        for p in FIVE:
            _show(state, lateral_name(p), True)
    elif n == "新規":
        _show(state, "L11", False)
        for p in FIVE:
            _show(state, lateral_name(p), False)


@dataclass
class LotAllocation:
    """JudgeLotAllocations の ByRef 引数(検番, NeedLateral, NeedRight, NeedFlat)。途中でエラーになっても入った分は残る。"""
    ken_no: str = ""
    need_lateral: bool = False
    need_right: bool = False
    need_flat: bool = False


def judge_lot_allocations(state: KensaState, rows: list[list[str]], hooks: Hooks,
                          result: Optional[LotAllocation] = None) -> LotAllocation:
    """FormRelat.JudgeLotAllocations (FormRelat:5259-5333)。引当(Lot情報履歴)ごとに品質基準を照合する。

    rows の列(1 始まり)= 1 納入先, 2 得意先, 3 用途ｺｰﾄﾞ, 4 材質, 5 調質, 6 厚, 7 巾, 8 丈, 9 ﾛｯﾄNo, 10 検査No。
    納入先が空の行は見ない。8 項目(厚巾丈は Single の CStr)が同じ行は 1 回だけ照合。1 件でも要なら要。
    検番は**最終行**の検査No の頭 1 文字。厚・巾・丈が数にならないと型エラー(result は途中まで入る)。
    """
    r = result if result is not None else LotAllocation()
    org_lateral_text = state.lateral_text
    org_additional = state.additional
    keep_lateral_text = keep_sumisei = keep_huris = ""
    done_keys = "\n"
    r.need_lateral = r.need_right = r.need_flat = False
    r.ken_no = ""
    for row in rows:
        r.ken_no = row[9]
        delivery = row[0]
        if delivery != "":
            customer, youto, zai, tyu = row[1], row[2], row[3], row[4]
            atu = _csng(row[5])
            hab = _csng(row[6])
            tak = _csng(row[7])
            key = "\t".join([delivery, customer, youto, zai, tyu, vbx.cstr(atu, is_single=True),
                             vbx.cstr(hab, is_single=True), vbx.cstr(tak, is_single=True)])
            if vbx.instr(done_keys, "\n" + key + "\n") == 0:
                done_keys += key + "\n"
                ret = hooks.db_return_item(state, delivery, customer, youto, zai, tyu, atu, hab, tak)
                if vbx.instr(ret, "ラテラルボー") > 0:
                    r.need_lateral = True
                    keep_lateral_text = state.lateral_text
                    keep_sumisei = state.sumisei
                if vbx.instr(ret, "直角度") > 0:
                    r.need_right = True
                if vbx.instr(ret, "フリー反り") > 0:
                    r.need_flat = True
                    keep_huris = state.huris
    r.ken_no = vbx.left(r.ken_no, 1)
    if vbx.instr(org_additional, "ﾗﾃﾗﾙﾎﾞ-追記公差あり") > 0:
        state.additional = org_additional
        state.lateral_text = org_lateral_text
        if r.need_lateral and keep_sumisei != "":
            state.sumisei = keep_sumisei
    elif r.need_lateral:
        if keep_lateral_text == "":
            keep_lateral_text = org_lateral_text
        state.lateral_text = keep_lateral_text
        state.sumisei = keep_sumisei
        if vbx.instr(state.additional, "ﾗﾃﾗﾙﾎﾞｰ検査不要") > 0:
            state.additional = ""
    if r.need_flat:
        state.huris = keep_huris
    return r


# ------------------------------------------------------------------
# LSelect
# ------------------------------------------------------------------
def _thin_yama_captions(state: KensaState, ken_no: str, *, with_ok: bool, four: bool = False,
                        at_butt: bool = True) -> None:
    """LSelect の薄板 `Select Case UF1.Yama.Text`(FormRelat:2902-2967 / 3421-3497)。With UF1 の中の `.` の文。

    at_butt=False(B56 を直したとき、2HI・6mm 超)は AtButt を表示し直さない。
    """
    if not with_ok:
        return
    y = state.yama
    ken_no_for_at = ken_no if at_butt else "T"      # 「検番が T 以外なら AtButt 表示」を効かせない(InspeButt は検番どおり)
    if y == "1":
        _caps(state, S="S", C="2", E="E", FS="S", FC="2", FE="E", OPS1="S", OPS3="2", OPS5="E",
              OPE1="S", OPE3="2", OPE5="E", U="", FU="", D="", FD="")
        if four:
            _caps(state, OPS2="", OPS4="", OPE2="", OPE4="")
        if ken_no_for_at != "T":
            state.at_butt_visible = True
        if state.yama_visible:
            state.inspe_butt_visible = True
        if ken_no == "T":
            state.inspe_butt_visible = False
    elif y == "2":
        _caps(state, S="2-2-S", U="", C="2-2-E", D="", E="2-1-E", FS="2-2-S", FU="", FC="2-2-E", FD="", FE="2-1-E",
              OPS1="2-2-S", OPS2="", OPS3="2-2-E", OPS4="", OPS5="2-1-E",
              OPE1="2-2-S", OPE2="", OPE3="2-2-E", OPE4="", OPE5="2-1-E")
        if ken_no_for_at != "T":
            state.at_butt_visible = True
        state.inspe_butt_visible = True
        if ken_no == "T":
            state.inspe_butt_visible = False
    elif y == "3":
        _caps(state, S="3-3-S", U="3-3-E", C="3-2-2", D="3-1-S", E="3-1-E",
              FS="3-3-S", FU="3-3-E", FC="3-2-2", FD="3-1-S", FE="3-1-E",
              OPS1="3-3-S", OPS2="3-3-E", OPS3="3-2-2", OPS4="3-1-S", OPS5="3-1-E",
              OPE1="3-3-S", OPE2="3-3-E", OPE3="3-2-2", OPE4="3-1-S", OPE5="3-1-E")
        if ken_no_for_at != "T":
            state.at_butt_visible = True
        state.inspe_butt_visible = False
    elif y == "4" and four:
        if ken_no_for_at != "T":
            state.at_butt_visible = True
        state.inspe_butt_visible = False


@dataclass
class LSelectInfo:
    """LSelect の途中で決まった値(試験・画面用)。"""
    ken_no: str = ""
    hi_flag: bool = False
    fs4_flag: bool = False
    atu_flag: bool = False
    new_atu: float = 0.0
    course: str = ""
    use_cache: bool = False
    jumped: bool = False          # GoTo LineSkip で入った(With UF1 の `.` の文は効かない)
    allocation: Optional[LotAllocation] = None


def _safe(fn, *a, **k):
    """LSelect は手続き全体が On Error Resume Next。呼んだ先の VBA の誤りは飛ばして None。"""
    try:
        return fn(*a, **k)
    except VBA_ERRORS as _vba_e:
        vba_swallow(_vba_e, "layout.py lselect")
        return None


def lselect(state: KensaState, n: str, hooks: Hooks, skip_db_check: bool = False,
            msgs: Optional[list[Message]] = None) -> LSelectInfo:
    """FormRelat.LSelect (FormRelat:2579-3864)。作業選択ごとの表示・2HI 等のコース判定・ゲージ要否・引当判定。

    n = "厚板" / "薄板" / "少枚数" / "極厚" / "スラブ" / "住精"。手続き全体が On Error Resume Next。
    強制描写(Byousya)と、住精で仕掛 DB が読めないときは `GoTo LineSkip`(薄板の Case の途中)へ飛ぶ。
    With UF1 の途中へ飛び込むので、With の `.` で書いた文はエラー91 で全部飛ばされる(UF1. と書いた文だけ効く)。
    B7: キャッシュ使用時のオーダー厚は IIf の両辺評価のエラーで 0 → 画面の厚(ATU)で 6mm 判定する。

    2026-11-07: 310 行の 1 本を、VBA の上から順の 8 つの段に分けた(中身・順番は同じ):
      1 ロットから決まること(検番・コース・2HI・6mm) 2 見出し 3 どの作業でも出す欄 4 作業ごとの行
      5 住精 6 ゲージ 7 対角 8 引当の判定
    目隠し(「少々お待ちください」)は途中で止まっても必ず消す。
    """
    msgs = msgs if msgs is not None else []
    info = LSelectInfo()
    state.blind_message = "少々お待ちください"            # frameblindfold("目隠し")
    try:
        _lselect_lot_facts(state, n, hooks, msgs, info)
        ken_no, hi, atu_flag, course = info.ken_no, info.hi_flag, info.atu_flag, info.course
        jumped = info.jumped
        with_ok = not jumped
        # B57: GoTo LineSkip が With UF1 の途中へ飛び、`.` の文がエラー91 で飛ばされる。直すと LineSkip 以降の `.` の文は効く
        fix_jump = jumped
        # B56: 2HI 通過・6mm 超で消した AtButt を LineSkip でまた表示する(検番 T 以外)。直すと消したまま(DB を見たときだけ)
        keep_at_hidden = (not jumped) and (hi or not atu_flag)
        _lselect_headers(state, n, hooks, ken_no, hi, atu_flag, with_ok, fix_jump, keep_at_hidden)
        _lselect_common_cells(state)
        _lselect_rows_by_work(state, n, hooks)
        _lselect_sumisei(state, n, hooks, ken_no, with_ok, fix_jump, keep_at_hidden)
        _lselect_gauge(state, n, hooks, course)
        _lselect_diag(state, n, hooks, with_ok, fix_jump)
        if not skip_db_check:
            _lselect_allocations(state, hooks, info)
    finally:
        state.blind_message = ""                         # frameblindfold("通常")。途中で止まっても目隠しを残さない
    return info


def _lselect_lot_facts(state: KensaState, n: str, hooks: Hooks, msgs: list[Message], info: LSelectInfo) -> None:
    """段 1(FormRelat:2585-2760): 仕掛からロットの検番・設計コース・オーダー厚を読み、2HI・L-2/L-3・6mm 以下を決める。"""
    safe = _safe
    if state.sika_set_atu != 0:
        try:
            all_ = hooks.sika_atu_create(state, state.lot)
            state.sika_atu_all = all_
            if all_ is not None:
                state.sika_set_atu = vbx.single(vbx.cdbl(all_[0]))
        except VBA_ERRORS as _vba_e:
            vba_swallow(_vba_e, "layout.py lselect")
            pass

    jumped = bool(state.byousya)
    ken_no = ""
    hi = fs4 = atu_flag = False
    course = ""
    new_atu = 0.0
    if not jumped:
        c = state.cache
        use_cache = c.lot_no == state.lot and c.cours != "" and c.ord_atu != 0
        info.use_cache = use_cache
        lot_acc = None
        cours = ord_atu = c_course = 0
        if use_cache:
            ken_no = c.ken_no
            if ken_no in ("0", ""):
                r = find_priority(c.cours)
                if r != "":
                    ken_no = r
        else:
            r12 = safe(hooks.is_over_12_hours, state)            # エラーなら代入ごと飛ばされ False(2640)
            result_t = bool(r12) if r12 is not None else False
            hi_acc = safe(hooks.arrbox, state, "SIKALOTNOW.accdb", "板", "Fields")
            if hi_acc is not None:
                cours = _field_no(hi_acc, "設計_設備ｺｰｽ")
                ord_atu = _field_no(hi_acc, "ｵｰﾀﾞｰ板厚")
                c_course = _field_no(hi_acc, "実績_設備ｺｰｽ")
            elif n == "住精":
                jumped = True
            if not jumped:
                lot_acc = None if state.lot == "" else safe(hooks.arrbox, state, "SIKALOTNOW.accdb", state.lot, "Record")
                if lot_acc is None:
                    if state.load == Switch.ON and not result_t and not c.skip_lot_not_found_msg:
                        msgs.append(Message("SIKALOTNOWに該当ロットがありません" + "LSelect", ERROR, source="FormRelat:2665"))
                    if n == "住精":
                        jumped = True
            if not jumped:
                c.skip_lot_not_found_msg = False
                k = safe(lookup_ken_no, state, hooks)
                ken_no = k if k is not None else ""
                if ken_no in ("0", ""):
                    try:
                        input_string = str(_cell1(lot_acc, c_course))
                    except VBA_ERRORS as _vba_e:
                        vba_swallow(_vba_e, "layout.py lselect")
                        input_string = ""
                    r = find_priority(input_string)
                    if r != "":
                        ken_no = r
        if not jumped:
            state.reien = ""
            state.comment = ""
            if use_cache:
                course = c.cours
            elif lot_acc is not None:
                try:
                    course = str(_cell1(lot_acc, cours))
                except VBA_ERRORS as _vba_e:
                    vba_swallow(_vba_e, "layout.py lselect")
                    course = ""
            if vbx.instr(course, "2HI") > 0:
                hi = True
                state.reien = ken_no + "_2HI"
                state.comment = "2HI材:厚測定強制4ヵ所"
            if vbx.instr(course, "L-2") > 0:
                fs4 = True
                state.reien = ken_no + "_L-2" + ("_2HI" if hi else "")
            elif vbx.instr(course, "L-3") > 0:
                fs4 = True
                state.reien = ken_no + "_L-3" + ("_2HI" if hi else "")
            if fs4 and vbx.instr(course, "LL7") > 0:
                state.reien = ken_no + "_LL7"
            # NewAtu(Single) = IIf(bUseCache, g_CacheOrdAtu, LotAcc(1, OrdAtu))
            try:
                raw = _cell1(lot_acc, ord_atu)                    # IIf は両辺を評価する(キャッシュ時は Empty → エラー)
            except VBA_ERRORS as _vba_e:
                vba_swallow(_vba_e, "layout.py lselect")
                new_atu = 0.0                                     # 代入ごと飛ばされる
                if use_cache:
                    new_atu = c.ord_atu
            else:
                try:
                    new_atu = c.ord_atu if use_cache else _csng(raw)
                except VBA_ERRORS as _vba_e:
                    vba_swallow(_vba_e, "layout.py lselect")
                    new_atu = 0.0
            if new_atu == 0 and state.atu != "":
                new_atu = vbx.single(vbx.val(state.atu))
            if vbx.val_single(new_atu) <= THIN_PLATE_ATU:
                atu_flag = True
    info.ken_no, info.hi_flag, info.fs4_flag, info.atu_flag = ken_no, hi, fs4, atu_flag
    info.new_atu, info.course, info.jumped = new_atu, course, jumped


def _lselect_headers(state: KensaState, n: str, hooks: Hooks, ken_no: str, hi: bool, atu_flag: bool,
                     with_ok: bool, fix_jump: bool, keep_at_hidden: bool) -> None:
    """段 2: 行の見出し(幅1…)・位置の見出し(S/S1/2…)・山数・AtButt。"""
    if with_ok:
        if n in ("厚板", "薄板", "少枚数"):
            _caps(state, L1="幅1", L2="幅2", L3="丈1", L4="丈2", L5="対角A", L6="対角B")
        elif n == "極厚":
            _caps(state, L1="幅1", L2="幅2", L3="幅3", L4="幅4", L5="幅5", L6="幅6", L7="丈1", L8="丈2",
                  L12="丈3", L13="丈4")
        if n == "厚板":
            state.at_count = ""
            state.at_butt_visible = False
            state.inspe_butt_visible = False
            state.yama_visible = False
            six_display(state, "off", hooks)
            _show(state, "YamaLabel", False)
            _caps(state, S="S", U="S1", C="2", D="E1", E="E", FS="S", FU="S1", FC="2", FD="E1", FE="E",
                  OPS1="S", OPS2="S1", OPS3="2", OPS4="E1", OPS5="E", OPE1="S", OPE2="S1", OPE3="2", OPE4="E1", OPE5="E")
        elif n in ("少枚数", "極厚"):
            state.at_count = ""
            state.at_butt_visible = False
            state.inspe_butt_visible = False
            state.yama_visible = False
            six_display(state, "off", hooks)
            _show(state, "YamaLabel", False)
            _caps(state, S="1", U="2", C="3", D="4", E="5", FS="1", FU="2", FC="3", FD="4", FE="5",
                  OPS1="1", OPS2="2", OPS3="3", OPS4="4", OPS5="5", OPE1="1", OPE2="2", OPE3="3", OPE4="4", OPE5="5")
        elif n == "薄板":
            if not hi and not atu_flag:
                state.at_butt_visible = False
            if hi:
                state.at_butt_visible = False
            state.inspe_butt_visible = True
            state.yama_visible = True
            _show(state, "YamaLabel", True)
            if state.yama == "":
                six_display(state, "off", hooks)
            if ken_no == "T":
                state.inspe_butt_visible = False
            # LineSkip:(With が効いているときだけ)
            _line_skip(state, ken_no, at_butt=not keep_at_hidden)
    elif fix_jump and (n == "薄板" or (n == "住精" and state.is_work("薄板"))):
        # B57 を直す: GoTo LineSkip の先(薄板の見出し・AtButt)を効かせる。薄板以外の作業では薄板の見出しにしない
        _line_skip(state, ken_no, at_butt=True)



def _lselect_common_cells(state: KensaState) -> None:
    """段 3: どの作業でも出す欄(E 列・厚 1〜4・外観・BH)と、隠す欄(ラテラル・R115 以外の直角度)。"""
    for name in ("TSE", "TEE", "E", "NGE", "LC2", "LC8"):
        _show(state, name, True)
    _cells(state, THICK_ROWS, FIVE, True, True)
    for p in FIVE:
        _show(state, "FM" + p, True)
    _show(state, "L11", False)
    for p in FIVE:
        _show(state, lateral_name(p), False)
    for i in list(range(1, 6)) + list(range(7, 12)):
        _show(state, f"LC{i}", True)
    for p in FIVE:
        _show(state, f"{p}_UpVisual", True)
        _show(state, f"{p}_UnderVisual", True)
        _show(state, bh_name(p), True)
    if state.youto != "R115":
        _show(state, "L10", False)
        for p in FIVE:
            _show(state, kou_name(p), False)



def _lselect_rows_by_work(state: KensaState, n: str, hooks: Hooks) -> None:
    """段 4: 作業ごとの行(幅・丈・対角・極厚の 11〜14 行)の表示。"""
    if n in ("厚板", "薄板", "少枚数"):
        for i in (7, 8, 12, 13):
            _show(state, f"L{i}", False)
        _cells(state, range(11, 15), FIVE, False, False)
    elif n == "住精":
        _cells(state, range(13, 15), FIVE, False, False)
    if n == "極厚":
        for i in range(7, 15):
            if i in (7, 8, 12, 13):                    # Select Case は最初の Case だけ(行 12・13 のセルは出さない)
                _show(state, f"L{i}", True)
            elif i in (11, 14):
                _cells(state, [i], FIVE, True, True)
    if n in ("厚板", "少枚数", "極厚"):
        for i in THICK_ROWS:
            _show(state, f"AL{i}", True)
        _cells(state, THICK_ROWS, FIVE, True, True)

    if n != "住精":                                    # 住精は GoTo skipSmisei
        if n == "厚板":
            _cells(state, NORMAL.size_rows, ("U", "D"), False, False)
            _cells(state, NORMAL.size_rows, ("C", "E"), True, True)
            for i in list(range(1, 6)) + list(range(7, 12)):
                _show(state, f"LC{i}", True)
            for p in ("U", "C", "D"):
                _show(state, f"{p}_UpVisual", True)
                _show(state, f"{p}_UnderVisual", True)
        elif n == "薄板":
            _thin_body(state, hooks, set_diag=False)
        elif n == "少枚数":
            _cells(state, NORMAL.size_rows, FIVE, True, True)
            _lc_and_visual(state)
        elif n == "極厚":
            _cells(state, range(5, 15), FIVE, True, True)
            _lc_and_visual(state)



def _lselect_sumisei(state: KensaState, n: str, hooks: Hooks, ken_no: str, with_ok: bool, fix_jump: bool,
                     keep_at_hidden: bool) -> None:
    """段 5: 住精(いまの作業が厚板か薄板かで、その作業の表示に合わせる)。"""
    if n == "住精":
        if state.is_work("厚板"):
            six_display(state, "off", hooks)
            _cells(state, NORMAL.size_rows, ("U", "D"), False, False)
            _cells(state, NORMAL.size_rows, ("C", "E"), True, True)
            for p in ("U", "C", "D"):
                _show(state, f"{p}_UpVisual", True)
                _show(state, f"{p}_UnderVisual", True)
            _show(state, "DiagS", True)
            _show(state, "DiagU", False, False)
            _show(state, "DiagC", True)
            _show(state, "DiagD", False, False)
            _show(state, "DiagE", True)
            for i in list(range(1, 6)) + list(range(7, 12)):
                _show(state, f"LC{i}", True)
            _caps(state, S="S", U="S1", C="2", D="E1", E="E", FS="S", FU="S1", FC="2", FD="E1", FE="E",
                  OPS1="S", OPS2="S1", OPS3="2", OPS4="E1", OPS5="E", OPE1="S", OPE2="S1", OPE3="2", OPE4="E1", OPE5="E")
        elif state.is_work("薄板"):
            _thin_body(state, hooks, set_diag=True)
            _thin_yama_captions(state, ken_no, with_ok=with_ok or fix_jump, four=True, at_butt=not keep_at_hidden)



def _lselect_gauge(state: KensaState, n: str, hooks: Hooks, course: str) -> None:
    """段 6: ゲージが要るか(要れば 幅ｹﾞｰｼﾞ/丈ｹﾞｰｼﾞ と歪コンボ)。"""
    safe = _safe
    if n in ("極厚", "厚板", "薄板", "少枚数", "スラブ"):
        try:
            gauge = hooks.is_gauge_required(vbx.val(state.atu), vbx.val(state.hab), vbx.val(state.tak),
                                            state.zai, state.tyu, course)
        except NotImplementedError:
            raise
        except Exception:                               # If の条件でエラー → Resume Next で Then の中へ進む
            gauge = True
        if gauge:
            for p in FIVE + ("F",):
                state.v("LZ" + p).value = "幅ｹﾞｰｼﾞ"
                state.v("LG" + p).value = "丈ｹﾞｰｼﾞ"
            safe(hooks.cb_add, state, state.sika_set_atu)
            state.blindfold_tag = "Set"
        else:
            for p in FIVE + ("F",):
                state.v("LZ" + p).value = "定盤"
                state.v("LG" + p).value = "ｹﾞｰｼﾞ"



def _lselect_diag(state: KensaState, n: str, hooks: Hooks, with_ok: bool, fix_jump: bool) -> None:
    """段 7: 対角(極厚は出さない・スラブ・住精の山数・厚板/薄板/少枚数)。"""
    safe = _safe
    if n == "極厚":
        if with_ok or fix_jump:
            _show(state, "L9", False)
            _show(state, "L10", False)
        for p in FIVE:
            _show(state, diag_name(p), False)
            _show(state, kou_name(p), False)
    if n == "スラブ":
        safe(slab_set, state, hooks)

    if n == "住精":
        if state.yama_visible and state.yama in ("1", "2"):
            _show(state, "DiagS", True)
            _show(state, "DiagU", False, False)
            _show(state, "DiagC", True)
            _show(state, "DiagD", False, False)
            _show(state, "DiagE", True)
        elif state.yama_visible and state.yama == "3":
            _show(state, "DiagU", True, True)
            _show(state, "DiagD", True, True)
    elif n == "厚板" or (n == "薄板" and state.yama != "3"):
        _show(state, "L9", True)
        _show(state, "DiagS", True)
        _show(state, "DiagU", False, False)
        _show(state, "DiagC", True)
        _show(state, "DiagD", False, False)
        _show(state, "DiagE", True)
    elif n == "薄板":
        _show(state, "L9", True)
        for p in FIVE:
            _show(state, diag_name(p), True)
        _show(state, "DiagU", enabled=True)
        _show(state, "DiagD", enabled=True)
    elif n == "少枚数":
        _show(state, "L9", True)
        for p in FIVE:
            _show(state, diag_name(p), True, True)



def _lselect_allocations(state: KensaState, hooks: Hooks, info: LSelectInfo) -> None:
    """段 8: Lot情報履歴の引当ごとの判定(ラテラル・直角度・フリーソリ。1 件でも要なら要)。"""
    safe = _safe
    rows = safe(hooks.get_lot_info, state, state.lot)
    if rows is not None:
        alloc = LotAllocation()
        info.allocation = alloc
        safe(judge_lot_allocations, state, rows, hooks, alloc)
        if alloc.need_lateral:
            safe(lselect_lateral, state, alloc.ken_no)
        if alloc.need_right:
            safe(right_angle_display, state, alloc.ken_no)
        if alloc.need_flat:
            state.flatnes_text = "ﾌﾘｰｿﾘ要"
        if state.lateral_text_colored:
            safe(lselect_lateral, state, alloc.ken_no)


def _line_skip(state: KensaState, ken_no: str, *, at_butt: bool) -> None:
    """LSelect の `LineSkip:` から薄板の Case の終わりまで(FormRelat:2897-3007)。"""
    state.at_butt_caption = "厚2ヵ所"
    state.at_count = "厚4ヵ所検査中"
    if ken_no != "T" and at_butt:
        state.at_butt_visible = True
    _thin_yama_captions(state, ken_no, with_ok=True, at_butt=at_butt)
    if not state.yama_visible:
        _caps(state, S="S", U="U", C="2", D="D", E="E", FS="S", FU="U", FC="2", FD="D", FE="E",
              OPS1="S", OPS2="U", OPS3="2", OPS4="D", OPS5="E", OPE1="S", OPE2="U", OPE3="2", OPE4="D", OPE5="E")


def _lc_and_visual(state: KensaState) -> None:
    for i in list(range(1, 6)) + list(range(7, 12)):
        _show(state, f"LC{i}", True)
    for p in ("U", "C", "D"):
        _show(state, f"{p}_UpVisual", True)
        _show(state, f"{p}_UnderVisual", True)


def _thin_body(state: KensaState, hooks: Hooks, *, set_diag: bool) -> None:
    """LSelect の薄板の山数別の表示(FormRelat:3193-3247 / 住精 3372-3419)。set_diag は住精側(対角も触る)。"""
    if state.yama == "4":
        try:                                           # AtCountSelect のエラーは LSelect の Resume Next で次の文へ
            six_display(state, "on", hooks)
        except VBA_ERRORS as _vba_e:
            vba_swallow(_vba_e, "layout.py _thin_body")
            pass
    elif state.yama == "3":
        _cells(state, range(1, 11), FIVE, True, True)
        if set_diag:
            for p in FIVE:
                _show(state, diag_name(p), True, True)
        for i in list(range(1, 6)) + list(range(7, 12)):
            _show(state, f"LC{i}", True)
        _show(state, "LC6", False)
        _show(state, "LC12", False)
        dist_select(state, 5)
    else:
        _cells(state, range(1, 11), ("U", "D"), False, False)
        if set_diag:
            _show(state, "DiagU", False, False)
            _show(state, "DiagD", False, False)
        for i in (2, 4, 8, 10, 6, 12):
            _show(state, f"LC{i}", False)
        for name in ("U_UpVisual", "D_UpVisual", "U_UnderVisual", "D_UnderVisual", "UBH", "DBH"):
            _show(state, name, False)
        dist_select(state, 26)


# ------------------------------------------------------------------
# 作業選択ボタン(オーダー厚)
# ------------------------------------------------------------------
def atu_opselect(state: KensaState, atu: float, hooks: Hooks, weight_g: float = 0.0,
                 msgs: Optional[list[Message]] = None) -> None:
    """FormRelat.Atu_OPselect (FormRelat:6185-6435)。オーダー厚で作業選択ボタンを出し分け、検査番号を表示する。

    `Select Case NewAtu` の範囲は VBA のまま(境界の隙間 6〜6.001・29.9〜30・199.9〜200・200〜200.1・299〜300 では消えない)。
    エラー処理は無い(DB の値が数にならない等は例外がそのまま上がる)。
    SIKAHIKINOW の行は Arrbox を通さず GetRecordsArr を直接読む(`hooks.atu_hiki_records`)。MainPath・NextPath とも
    無ければ「日稲:SIKAHIKINOW.accdbがありません_ReadCheckAtu_OPselect」(IsOver12Hours に関係なく出す)、Comment を消して
    リセットの前に抜ける(FormRelat:6233-6242)。
    """
    msgs = msgs if msgs is not None else []
    ken_no = ""
    new_atu = 0.0
    if atu != 0:
        hi_acc = hooks.arrbox(state, "SIKALOTNOW.accdb", "板", "Fields")
        ord_atu = _field_no(hi_acc, "ｵｰﾀﾞｰ板厚")
        no1 = _field_no(hi_acc, "検査番号")
        c_course = _field_no(hi_acc, "実績_設備ｺｰｽ")
        lot_acc = hooks.arrbox(state, "SIKALOTNOW.accdb", state.lot, "Record")
        if lot_acc is not None:
            new_atu = vbx.cdbl(_cell1(lot_acc, ord_atu))
        if vbx.val(atu) != vbx.val(new_atu):
            state.comment = "ｵｰﾀﾞｰ厚参照_" + vbx.fmt(new_atu, "0.000")
        atu = new_atu
        fields = hooks.arrbox(state, "SIKAHIKINOW.accdb", "", "Fields")
        no2 = _field_no(fields, "検査番号")
        # FormRelat:6227-6247: Arrbox を通さず、自前で MainPath/NextPath を選んで GetRecordsArr(Filter は Lot が空でも付く)
        found, ken_acc = hooks.atu_hiki_records(state, state.lot)
        if not found:
            msgs.append(Message(paths.missing_text("SIKAHIKINOW"), ERROR,
                                source="FormRelat:6238"))
            state.comment = ""
            return                                          # Exit Sub(SkipA のリセットは走らない)
        if ken_acc is not None:
            v = _cell1(ken_acc, no2)
            ken_no = vbx.left(v, 1)
            state.ken_no_caption = v
        elif lot_acc is not None:
            v = _cell1(lot_acc, no1)
            ken_no = vbx.left(v, 1)
            state.ken_no_caption = v
        if lot_acc is not None and ken_no in ("0", ""):
            input_string = _cell1(lot_acc, c_course)
            result = find_priority(input_string)
            if result != "":
                msgs.append(Message("最終圧延機から検査番号頭を生成します" + "\r\n" + "オーダーが付いていないので納入先が不明"
                                    + "\r\n" + "ラテラルボーが必要の可能性があります" + "\r\n"
                                    + "Q-S-QA-ALL-088_検査番号の付け方", INFO, source="FormRelat:6283"))
                ken_no = result
            if result == "":
                msgs.append(Message("データ上にまだ検査番号がありません" + "\r\n" + "(一部動作しない機能があります)",
                                    INFO, source="FormRelat:6290"))
                state.comment = ""
    # SkipA: ラベル・フレームの表示を戻す(For Each MyCtrl In UF1.Controls)
    for name, c in state.iter_cells():
        if len(name) == 2:
            c.visible = c.enabled = True
        elif name.endswith("0"):
            c.visible = c.enabled = True
        else:
            c.visible = c.enabled = False
    for p in FIVE + ("F",):
        _show(state, p, True, False)                 # 列見出し(1 文字)
        _show(state, "F" + p, True, False)           # 歪見出し FS..FF
        _show(state, f"{p}_UpVisual", True, True)
        _show(state, f"{p}_UnderVisual", True, True)
        _show(state, diag_name(p), True, True)
        _show(state, "FM" + p, True)
    for i in THICK_ROWS:
        _show(state, f"AL{i}", True)
    _show(state, "L11", state.sumisei != "" or state.lateral_text_tag == "有")
    _show(state, "L12", False)
    _show(state, "L13", False)
    for i in range(1, 10):
        _show(state, f"L{i}", i not in (7, 8))
    _show(state, "ConductivL", False)
    state.work = ""
    for w in WORKS:
        state.work_visible[w] = True
    state.at_butt_visible = False
    state.inspe_butt_visible = False
    state.yama_visible = False
    _show(state, "YamaLabel", False)
    state.conductiv_text = ""

    if ken_no == "":
        state.comment = ""
        return
    if ken_no == "T":
        state.work_caption["薄板"] = ken_no
        for w in WORKS:
            state.work_visible[w] = w == "薄板"
        return
    x = new_atu
    if 1 <= x <= 6.0:
        state.work_visible["厚板"] = False
        state.work_caption["薄板"] = ken_no
    elif 200.1 <= x <= 299.0:
        if weight_g != 0 and weight_g > 1500.0:
            state.work_visible["厚板"] = False
    elif 300.0 <= x <= 800.0:
        state.work_visible["厚板"] = False
    if 6.001 <= x <= 800.0:
        state.work_visible["薄板"] = False
    state.work_caption["厚板"] = ken_no
    if 1 <= x <= 6.0:
        state.work_visible["少枚数"] = False
    if 1 <= x <= 29.9:
        state.work_visible["極厚"] = False
    if 1 <= x <= 199.9:
        state.work_visible["スラブ"] = False


def op_forced_selection(state: KensaState) -> None:
    """FormRelat.OpForcedSelection (FormRelat:6974-7022)。作業未選択なら「作業選択を選んでください」(大きさの変更は移さない)。"""
    if state.load == Switch.OFF or state.lot == "":
        return
    state.blind_message = "作業選択を選んでください" if state.work == "" else ""


# ------------------------------------------------------------------
# 画面のイベント
# ------------------------------------------------------------------
def select_work(state: KensaState, name: str, hooks: Hooks, msgs: Optional[list[Message]] = None) -> None:
    """作業選択ボタン(UF1.厚板_Click 5204 / 少枚数_Click 5244 / 薄板_Click 5278 / 極厚_Click 5302 / スラブ_Click 5312)。"""
    msgs = msgs if msgs is not None else []
    state.work = name
    if name == "厚板":
        lselect(state, "厚板", hooks, True, msgs)
        lselect(state, "住精" if state.sumisei != "" else "厚板", hooks, False, msgs)
        op_forced_selection(state)
    elif name == "少枚数":
        lselect(state, "少枚数", hooks, False, msgs)
        th = small_count_th(state, hooks.calculate_sheets)
        if state.work_caption["厚板"] == "H" and state.is_work("少枚数"):
            hooks.small_number_count(state, th)
            op_forced_selection(state)
            return
        lselect(state, "住精" if state.sumisei != "" else "少枚数", hooks, False, msgs)
        op_forced_selection(state)
    elif name == "薄板":
        lselect(state, "薄板", hooks, True, msgs)
        if state.sumisei != "":
            lselect(state, "住精", hooks, False, msgs)
    elif name == "極厚":
        lselect(state, "極厚", hooks, False, msgs)
        op_forced_selection(state)
    elif name == "スラブ":
        lselect(state, "住精" if state.sumisei != "" else "スラブ", hooks, False, msgs)
        op_forced_selection(state)
    else:
        raise ValueError(name)


def byousya_after_update(state: KensaState, hooks: Hooks) -> None:
    """UF1.Byousya_AfterUpdate (UF1:554-566)。強制描写: 作業選択を外し 6 回目まで表示、ラテラル・直角度を強制表示。

    B62: VBA はスラブを外さない(厚板・薄板・少枚数・極厚だけ False にする)。直すとスラブも外す。On Error Resume Next。
    """
    works = ("厚板", "薄板", "少枚数", "極厚", "スラブ")
    if state.work in works:
        state.work = ""
    for fn, args in ((six_display, ("on", hooks)), (lselect_lateral, ("強制", "6回目も表示")),
                     (right_angle_display, ("強制", "6回目も表示"))):
        try:
            fn(state, *args)
        except VBA_ERRORS as _vba_e:
            vba_swallow(_vba_e, "layout.py byousya_after_update")
            pass
    hooks.initial_setting(state)


def yama_change(state: KensaState, hooks: Hooks, msgs: Optional[list[Message]] = None) -> None:
    """UF1.Yama_Change (UF1:2246-2460)。山数の変更。1〜4 は UF1、5 以上は「別フォームで検査してください」。

    `Select Case UF1.Yama.Text: Case 1, 2, 3, 4`(.Text は String)。数にならない文字は型エラー(U1)。
    数字なら "0" 書式に直す(値が変わると Change がもう一度走る)。
    """
    msgs = msgs if msgs is not None else []
    if state.load == Switch.OFF:
        return
    if vbx.is_numeric(state.yama):
        new = vbx.fmt(state.yama, "0")
        if new != state.yama:
            state.yama = new
            yama_change(state, hooks, msgs)
    y = state.yama
    if any(typed_cmp(y, k) == 0 for k in (1, 2, 3, 4)):
        state.yama_tag = "UF1"
    elif y == "":
        state.yama_tag = "UF1"
        return
    else:
        state.yama_tag = "別フォーム"
        msgs.append(Message("別フォームで検査してください", ERROR, source="UF1:2266"))
        return

    if y == "4":
        six_display(state, "on", hooks)
        if state.sumisei != "":
            lselect(state, "住精", hooks, False, msgs)
        hooks.initial_setting(state)
        return
    six_display(state, "off", hooks)

    if state.inspe_butt_visible and state.inspe_butt_caption == "検入≧26":
        if typed_cmp(y, 1) == 0:
            _caps(state, S="S", C="", E="E", FS="S", FC="", FE="E", OPS1="S", OPS3="", OPS5="E",
                  OPE1="S", OPE3="", OPE5="E", U="", FU="", D="", FD="")
        elif typed_cmp(y, 2) == 0:
            _caps(state, S="2-2-S", U="", C="", D="", E="2-1-E", FS="2-2-S", FU="", FC="", FD="", FE="2-1-E",
                  OPS1="2-2-S", OPS2="", OPS3="", OPS4="", OPS5="2-1-E",
                  OPE1="2-2-S", OPE2="", OPE3="", OPE4="", OPE5="2-1-E")
    elif state.inspe_butt_visible and state.inspe_butt_caption == "検入≦25":
        if typed_cmp(y, 1) == 0:
            _caps(state, S="S", C="2", E="E", FS="S", FC="2", FE="E", OPS1="S", OPS3="2", OPS5="E",
                  OPE1="S", OPE3="2", OPE5="E", U="", FU="", D="", FD="")
        elif typed_cmp(y, 2) == 0:
            _caps(state, S="2-2-S", U="", C="2-2-E", D="", E="2-1-E", FS="2-2-S", FU="", FC="2-2-E", FD="",
                  FE="2-1-E", OPS1="2-2-S", OPS2="", OPS3="2-2-E", OPS4="", OPS5="2-1-E",
                  OPE1="2-2-S", OPE2="", OPE3="2-2-E", OPE4="", OPE5="2-1-E")

    if state.sumisei != "":
        lselect(state, "住精", hooks, False, msgs)
    else:
        if state.load != Switch.OFF:
            lselect(state, "薄板", hooks, False, msgs)
        if state.work_caption["薄板"] == "T" and state.yama == "1":
            _caps(state, OPS1="S", OPS2="", OPS3="2", OPS4="", OPS5="E", OPE1="S", OPE2="", OPE3="2", OPE4="", OPE5="E")
        elif state.work_caption["薄板"] == "T" and state.yama == "2":
            _caps(state, OPS1="2-2-S", OPS2="", OPS3="2-2-E", OPS4="", OPS5="2-1-E",
                  OPE1="2-2-S", OPE2="", OPE3="2-2-E", OPE4="", OPE5="2-1-E")
        for i in (3, 4):
            if state.yama == "3":
                _show(state, f"C{i}", True, True)
                _show(state, "DiagU", True, True)
                _show(state, "DiagE", True, True)
        if state.load != Switch.OFF:
            if state.at_butt_caption == "厚2ヵ所":
                at_count_select(state, 4, hooks)
            elif state.at_butt_caption == "厚4ヵ所":
                at_count_select(state, 2, hooks)
        if state.load != Switch.OFF:
            inspection_select(state, "表示", hooks)
    # skipSmisei:
    hooks.initial_setting(state)
    hooks.user_form_click(state)
    if state.load != Switch.OFF:
        hooks.posting(state)
        hooks.write_log(state, "*UFKen:山数ﾁｪﾝｼﾞ:" + state.yama)


def line_auto_selection(state: KensaState, n: str, hooks: Hooks, maru_no: int = 0) -> None:
    """FormRelat.Line_AutoSelection (FormRelat:6485-6580)。ライン・号機を決め、作業保存を読み、引当を判定する。

    `maru_no` は丸徳の号機 OptionButton A1..A11 のうち True の番号(0 = なし。LineNo はそのまま)。
    B6: Lot情報履歴にロットが無いと WorkerLoad の後 Exit Sub し、Load="off"・Tag="STOP" のまま残る。
        直した場合は SkipLoad へ進む(WorkerLoad → 再判定 → Load="on")。
    """
    if state.line_option in LINE_CAPTIONS:
        state.line_name = LINE_CAPTIONS[state.line_option]
    if state.line_option == "Maru":
        if maru_no in MARU_NO_CAPTIONS:
            state.line_no = MARU_NO_CAPTIONS[maru_no]
    else:
        state.line_no = ""
    state.v("ZEB").value = "検入4山以上" if vbx.mid(state.line_name, 1, 2) == "丸徳" else "全数.三菱.5N"
    state.form_tag = FormTag.STOP
    state.load = Switch.OFF
    try:
        hooks.load_text(state, n)
        if state.lot != "":
            rows = hooks.get_lot_info(state, state.lot)
            if rows is not None:
                alloc = judge_lot_allocations(state, rows, hooks)
                if alloc.need_flat:
                    state.flatnes_text = "ﾌﾘｰｿﾘ要"
        hooks.worker_load(state)
        hooks.loop_judg(state)
        hooks.hiz_loop_judg(state)
        for p in FIVE + ("F",):
            hooks.check_labels_by_category(state, p)
    finally:
        state.load = Switch.ON                           # 途中で止まっても旗は戻す(2026-11-06。B6 を直した)
        state.form_tag = ""
