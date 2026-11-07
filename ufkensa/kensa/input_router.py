"""伝送文字列の解析と次の入力欄(UF1.INP_Change と FormRelat の PucketMove / AllCheck / Look_*_Blank)

流れ(UF1.INP_Change UF1:4565-4651):
    受信(StartTB → INP) → `maker` で種別 → `pucket_move` で入力先(INP.Tag)→ `conv` の値を入れる
    → 検査員・作業選択・校正のブロック → LoopText(判定。別担当)→ Posting → Out_Tolerance
    → `next_blank` で次の入力欄(黄色) → INP を空にする
判定(LoopText 以降)は `grid.Hooks` 経由で呼ぶだけ。ここは「入力先と値」を決めるところまで。
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterator, Optional

from . import vbx
from .model import NORMAL, Switch, set_verdict, bad_name, diag_name, is_thick
from .flow import ERROR, INFO, WARN, Message, Stop
from .grid import (BORDER_BLACK, BORDER_RED, COLOR_DEFAULT, COLOR_WHITE, COLOR_YELLOW, Hooks, KensaState,
                   fm_clear, ken_flags, small_count_th, yama_is)

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


def _blank(state: KensaState, name: str) -> bool:
    """`UF1(name) = ""`(Label の既定プロパティ Caption)。無い名前は KeyError(VBA の実行時エラー)。"""
    return state.ctl(name).value == ""


def _blank_vis(state: KensaState, name: str) -> bool:
    """`UF1(name) = "" And UF1(name).Visible`(And は両方とも評価する)。"""
    c = state.ctl(name)
    return c.value == "" and c.visible


def _set_target(state: KensaState, name: str) -> None:
    state.inp_tag = name
    state.tagtext = name


# ------------------------------------------------------------------
# AllCheck / PucketMove
# ------------------------------------------------------------------
def all_check(state: KensaState, a: int, n: int, ken_no: str, temp: str, hooks: Hooks) -> bool:
    """FormRelat.AllCheck (FormRelat:429-715)。行 a〜n の対象列に空きがあれば真。

    temp: "厚板"(行5〜10 は S/C/E、ほかは S/U/C/D/E)/ "少枚数"(枚数 th で列が変わる)/ ""(薄板・少枚数・極厚・スラブ)
    """
    if a == 0 or n == 0:
        return False
    th = small_count_th(state, hooks.calculate_sheets, start_from_ken_count=True)

    def scan(cols_for_row) -> bool:
        for i in range(a, n + 1):
            for p in cols_for_row(i):
                if _blank(state, f"{p}{i}"):
                    return True
        return False

    five = ("S", "U", "C", "D", "E")
    if temp == "厚板":
        return scan(lambda i: ("S", "C", "E") if 5 <= i <= 10 else five)
    if temp == "少枚数":
        if th == 5:
            return scan(lambda i: five)
        if th == 4:
            return scan(lambda i: ("S", "C", "D") if 5 <= i <= 10 else ("S", "U", "C", "D"))
        if th == 3:
            return scan(lambda i: ("S", "U", "C"))
        if th == 2:
            return scan(lambda i: ("S", "U"))
        if th == 1:
            return scan(lambda i: ("S",))
        return False
    if temp == "":
        if state.is_work("薄板"):
            f = ken_flags(state, ken_no)
            if f[6]:
                return scan(lambda i: five + ("F",))
            if f[3] or f[5]:
                return scan(lambda i: five)
            if f[4]:
                return scan(lambda i: ("S", "E"))
            if f[1]:
                return scan(lambda i: ("S", "C", "E"))
            return False          # 検番 H だけ(flag2)はどれにも当たらない(G1 §11-8)
        if state.is_work("少枚数") or state.is_work("極厚"):
            return scan(lambda i: five)
        if state.is_work("スラブ"):
            return scan(lambda i: ("S", "C") if is_thick(i) else (("S", "U", "C", "D") if 5 <= i <= 10 else ()))
    return False


def _no_work(state: KensaState) -> bool:
    return state.work == ""


def pucket_move(state: KensaState, pucket: str, ken_no: str, youto: str, hooks: Hooks,
                msgs: Optional[list[Message]] = None) -> None:
    """FormRelat.PucketMove (FormRelat:716-1016)。伝送の種別から行範囲を決め、空いた欄を INP.Tag に入れる。

    見つからないときは INP.Tag を変えない(前回の入力先のまま。VBA どおり)。
    """
    msgs = msgs if msgs is not None else []
    if _no_work(state):
        return
    f = ken_flags(state, ken_no)
    a = n = 0
    if pucket == "DT":
        if state.at_butt_caption == "厚2ヵ所":       # ボタン表示が2ヵ所 = いま 4 回検査
            a, n = 1, 4
        elif state.at_butt_caption == "厚4ヵ所":
            a, n = 1, 2
    elif pucket == "#L":
        a, n = (5, 14) if state.is_work("極厚") else (5, 10)
    elif pucket == "#〆":
        a, n = (1, 14) if state.is_work("極厚") else (1, 10)
    else:
        msgs.append(Message("ﾊﾟｹｯﾄ無し:Exit Sub" + "_PucketMove", WARN, source="FormRelat:774"))
        return

    temp = "厚板" if state.is_work("厚板") else ""
    if state.work_caption["厚板"] == "H" and state.is_work("少枚数"):
        temp = "少枚数"

    if not all_check(state, a, n, ken_no, temp, hooks):
        msgs.append(Message(pucket + "検査済", WARN, source="FormRelat:792"))
        state.auto = Switch.OFF
        state.inp = ""
        state.inp_tag = ""
        state.flag_label = 1
        return

    rows = range(a, n + 1)
    for i in rows:                                              # S
        if _blank(state, f"S{i}"):
            return _set_target(state, f"S{i}")

    # U
    if state.is_work("厚板"):
        for i in rows:
            if is_thick(i):
                if _blank(state, f"U{i}"):
                    return _set_target(state, f"U{i}")
            elif 5 <= i <= 10:
                for b in NORMAL.size_rows:
                    if _blank(state, f"C{b}"):
                        return _set_target(state, f"C{b}")
    elif state.is_work("スラブ"):
        for i in rows:
            if is_thick(i):
                for b in range(5, 9):
                    if _blank(state, f"U{b}"):
                        return _set_target(state, f"U{b}")
            elif 5 <= i <= 8:
                if _blank(state, f"U{i}"):
                    return _set_target(state, f"U{i}")
    elif state.is_work("薄板"):
        if f[2] or f[5] or f[3] or f[6]:
            for i in rows:
                if _blank(state, f"U{i}"):
                    return _set_target(state, f"U{i}")
        elif f[4]:
            for i in rows:
                if 1 <= i <= 10 and _blank(state, f"E{i}"):
                    return _set_target(state, f"E{i}")
        elif f[1]:
            for i in rows:
                if 1 <= i <= 10 and _blank(state, f"C{i}"):
                    return _set_target(state, f"C{i}")
    else:
        for i in rows:
            if _blank_vis(state, f"U{i}"):
                return _set_target(state, f"U{i}")

    # C
    if not state.is_work("薄板"):
        for i in rows:
            if _blank_vis(state, f"C{i}"):
                return _set_target(state, f"C{i}")
    elif f[1] or f[2] or f[3] or f[5] or f[6]:
        for i in rows:
            if _blank(state, f"C{i}"):
                return _set_target(state, f"C{i}")
    elif f[4]:
        for i in rows:
            if 1 <= i <= 10 and _blank(state, f"E{i}"):
                return _set_target(state, f"E{i}")

    # D
    if state.is_work("厚板"):
        for i in rows:
            if is_thick(i):
                if _blank(state, f"D{i}"):
                    return _set_target(state, f"D{i}")
            elif 5 <= i <= 10:
                for b in NORMAL.size_rows:
                    if _blank(state, f"E{b}"):
                        return _set_target(state, f"E{b}")
    elif state.is_work("スラブ"):
        for i in rows:
            if is_thick(i):
                for b in range(5, 9):
                    if _blank(state, f"D{b}"):
                        return _set_target(state, f"D{b}")
            elif 5 <= i <= 8:
                if _blank(state, f"D{i}"):
                    return _set_target(state, f"D{i}")
    elif state.is_work("薄板"):
        if f[2] or f[5] or f[3] or f[6]:
            for i in rows:
                if _blank(state, f"D{i}"):
                    return _set_target(state, f"D{i}")
        elif f[4] or f[1]:
            for i in rows:
                if 1 <= i <= 10 and _blank(state, f"E{i}"):
                    return _set_target(state, f"E{i}")
    else:
        for i in rows:
            if _blank_vis(state, f"D{i}"):
                return _set_target(state, f"D{i}")

    for i in rows:                                              # E
        if _blank_vis(state, f"E{i}"):
            return _set_target(state, f"E{i}")

    # F(F11 以降は無い → KeyError = VBA の実行時エラー)
    # B63: 極厚(行 1〜14)で山数 4 が残っていると F11 を見てエラー。直す: 極厚には F 列が無いので見ない
    if f[6] and (not state.is_work("極厚")):
        for i in rows:
            if _blank_vis(state, f"F{i}"):
                return _set_target(state, f"F{i}")
    return None


# ------------------------------------------------------------------
# 次の空き欄
# ------------------------------------------------------------------
def lookblank_color(state: KensaState, a: str) -> None:
    """FormRelat.LookBlank_Color (FormRelat:1272-1282)。MPEX.Tag が空なら赤枠・黄色、そうでなければ黒枠・白。"""
    state.tagtext = a
    c = state.ctl(a)
    if state.mpex_tag == "":
        c.border, c.color = BORDER_RED, COLOR_YELLOW
    else:
        c.border, c.color = BORDER_BLACK, COLOR_WHITE
    set_verdict(state.verdicts, a, c.color)


def _skip34_rows(a: int, n: int) -> Iterator[Optional[int]]:
    """Look_S_Blank の厚4ヵ所の For ループ(FormRelat:1070-1085)。

        For i = a To n
          Case 1, 2, 5 To 10: 行 i を見る
          Case 3: i = i + 2 → 行 5 を見る     Case 4: i = i + 1 → 行 5 を見る
        Next i
    ループ変数を書き換えるので、3 の次は 6(4 は通らない)。見ない回は None。
    """
    i = a
    while i <= n:
        if i in (1, 2, 5, 6, 7, 8, 9, 10):
            yield i
        elif i == 3:
            i += 2
            yield i
        elif i == 4:
            i += 1
            yield i
        else:
            yield None
        i += 1


def look_s_blank(state: KensaState, a: int, n: int, ken_no: str) -> None:
    """FormRelat.Look_S_Blank (FormRelat:1018-1271)。薄板の次の空き欄。Auto が "on" のときだけ。"""
    if state.auto != Switch.ON:
        return
    f = ken_flags(state, ken_no, blank_in_flag1=True)
    thin = state.is_work("薄板")

    def pick(name: str) -> bool:
        _set_target(state, name)
        lookblank_color(state, name)
        return True

    if thin and state.at_butt_caption == "厚4ヵ所":            # いま 2 ヵ所検査(行 3・4 を飛ばす)
        def scan(p: str, vis: bool) -> bool:
            for i in _skip34_rows(a, n):
                if i is None:
                    continue
                if (_blank_vis if vis else _blank)(state, f"{p}{i}"):
                    return pick(f"{p}{i}")
            return False
        if scan("S", False):
            return
        if f[4] and scan("E", False):
            return
        if (f[2] or f[3] or f[5] or f[6]) and scan("U", True):
            return
        if (f[1] or f[2] or f[3] or f[5] or f[6]) and scan("C", True):
            return
        if (f[2] or f[3] or f[5] or f[6]) and scan("D", True):
            return
        if scan("E", True):
            return
        if f[6] and scan("F", True):
            return
        return            # B60: 外の F ループは厚2ヵ所(4回検査)用。厚4ヵ所で隠れた F3/F4 を選ばない
    elif thin and state.at_butt_caption == "厚2ヵ所":
        def scan2(p: str) -> bool:
            for i in range(a, n + 1):
                if _blank(state, f"{p}{i}"):
                    return pick(f"{p}{i}")
            return False
        if scan2("S"):
            return
        if f[4] and scan2("E"):
            return
        if (f[2] or f[3] or f[5] or f[6]) and scan2("U"):
            return
        if (f[1] or f[2] or f[3] or f[5] or f[6]) and scan2("C"):
            return
        if (f[2] or f[3] or f[5] or f[6]) and scan2("D"):
            return
        if scan2("E"):
            return
    if f[6]:                                                    # Select の外。表示を見ない(FormRelat:1260-1270。B60)
        for i in range(a, n + 1):
            if _blank(state, f"F{i}"):
                pick(f"F{i}")
                return


def smallnumber_blank(state: KensaState, a: int, n: int, xa: int) -> None:
    """FormRelat.Smallnumber_Blank (FormRelat:1283-1340)。少枚数(枚数 xa が 1〜5 のとき)の次の空き欄。"""
    if state.auto != Switch.ON:
        return
    if xa not in (5, 4, 3, 2, 1):
        return
    for p, vis in (("S", False), ("U", True), ("C", True), ("D", True), ("E", True), ("F", True)):
        for i in range(a, n + 1):
            if (_blank_vis if vis else _blank)(state, f"{p}{i}"):
                _set_target(state, f"{p}{i}")
                lookblank_color(state, f"{p}{i}")
                return


def look_blank(state: KensaState, a: int, n: int) -> None:
    """FormRelat.Look_Blank (FormRelat:1341-1500)。厚板・極厚・スラブ・少枚数(枚数なし)の次の空き欄。

    表示(Visible)は見ない。厚板の行5〜10 は C1〜C10・E1〜E10 から探す(PucketMove は 5〜10)。
    """
    if state.auto != Switch.ON:
        return

    def pick(name: str) -> bool:
        _set_target(state, name)
        c = state.ctl(name)
        c.border, c.color = BORDER_RED, COLOR_YELLOW
        set_verdict(state.verdicts, name, c.color)
        return True

    rows = range(a, n + 1)

    def side(p: str, mid: str) -> bool:
        if state.is_work("厚板"):
            for i in rows:
                if is_thick(i):
                    if _blank(state, f"{p}{i}"):
                        return pick(f"{p}{i}")
                elif 5 <= i <= 10:
                    for b in range(1, 11):
                        if _blank(state, f"{mid}{b}"):
                            return pick(f"{mid}{b}")
        elif state.is_work("スラブ"):
            for i in rows:
                if is_thick(i):
                    for b in range(5, 9):
                        if _blank(state, f"{p}{b}"):
                            return pick(f"{p}{b}")
                elif 5 <= i <= 8:
                    if _blank(state, f"{p}{i}"):
                        return pick(f"{p}{i}")
        else:
            for i in rows:
                if _blank(state, f"{p}{i}"):
                    return pick(f"{p}{i}")
        return False

    for i in rows:
        if _blank(state, f"S{i}"):
            pick(f"S{i}")
            return
    if side("U", "C"):
        return
    for i in rows:
        if _blank(state, f"C{i}"):
            pick(f"C{i}")
            return
    if side("D", "E"):
        return
    for i in rows:
        if _blank(state, f"E{i}"):
            pick(f"E{i}")
            return
    # B63: 極厚(行 1〜14)で山数 4 が残っていると F11 を見て実行時エラー。直す: 極厚には F 列が無いので見ない
    if yama_is(state, 4) and (not state.is_work("極厚")):
        for i in rows:
            if _blank(state, f"F{i}"):
                pick(f"F{i}")
                return


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
    """INP_Change の検番(UF1:4585)。B3: VBA は冷延欄 `Reien` の頭 1 文字(UserForm_Click は `KenNo`)。"""
    return vbx.left(state.ken_no_caption, 1)


# ------------------------------------------------------------------
# INP_Change
# ------------------------------------------------------------------
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
    if state.flag_label != 1:
        return msgs
    if state.inp == "":
        return msgs
    if state.yama_visible and state.yama == "":
        raise Stop("山数の入力がありません" + "\r\n" + "選択してください", level=WARN, source="UF1:4569")
    if state.yama_tag == "別フォーム":
        raise Stop("別フォームで検査してください", level=ERROR, source="UF1:4574")

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
    INP は前の値のまま残るので、その後に同じ文字列が来ても Change は起きない(VBA どおり)。
    """
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
        c = state.ctl(name)
        c.border, c.color = BORDER_RED, COLOR_YELLOW
        set_verdict(state.verdicts, name, c.color)
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
        state.ctl(p + "10").color = COLOR_DEFAULT
        state.ctl(p + "9").color = COLOR_DEFAULT
        set_verdict(state.verdicts, p + "10", COLOR_DEFAULT)
        set_verdict(state.verdicts, p + "9", COLOR_DEFAULT)


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
        if c.value != "" and c.is_ng:
            fl = True
            break
    if name.startswith(("Lateral", "Z", "G")) or (name.startswith("Diag") and name.endswith("Kou")):
        if name in ("LateralS", "LateralU", "LateralC", "LateralD", "LateralE", "LateralF",
                    "DiagSKou", "DiagUKou", "DiagCKou", "DiagDKou", "DiagEKou", "DiagFKou",
                    "ZS", "ZU", "ZC", "ZD", "ZE", "ZF", "GS", "GU", "GC", "GD", "GE", "GF"):
            if state.ctl(name).is_ng:
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
    c.color = COLOR_DEFAULT
    set_verdict(state.verdicts, name, COLOR_DEFAULT)
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
