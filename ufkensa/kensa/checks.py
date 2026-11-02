"""公差外・入力チェック・印刷前チェック(VBA `【UFKen_05】OutTolerance`)

VBA は UF1 のコントロール(Caption・Text・BackColor・Visible)を直接読み書きしていた。
ここでは画面の値を `Uf1`(読むだけ)で受け取り、結果(出す確認画面・消す欄・メッセージ)を
dataclass で返す。画面へ反映するのは呼び手(Web 側)の仕事。

【この章で決めた読み替え】
- `Uf1.text[名前]` = そのコントロールの Caption(Label)/Text・Value(TextBox・ComboBox)。
  ラベル `S`〜`F` は位置の見出し(UF1.S.Caption)、`INP.Tag` は入力先(TagSet が入れる名前)。
- `Uf1.ng` = 背景が NG 色(8421631)のコントロール名。VBA の `BackColor = Ncol` の判定はこれだけを見る。
- `Uf1.visible[名前]` = Visible。書いていないものは非表示とみなす(呼び手が全部渡す)。
- 作業保存シート `O2:P13`(歪 A/B 面の 定盤・ゲージ)は `WorkDist` で受け取る。シートが無いときは None。

【U1(数にならない文字列と数の比較)】
`Select Case Rig` の `Case 2, 4, …`(Rig が "" や "S")と `.ZS = 0`(ZS が空)は、
`vbx.u1_compare`(U1 のスイッチ `vbx.U1_TEXT_VS_NUMBER_RAISES`。既定はエラーにせず「数 < 文字列」= 一致しない、
真なら型エラー13 = `vbx.TypeMismatch`)で比べる。比べる所は `_num_eq` / `_case_between` に集めた。
"""
from __future__ import annotations

import math
import re
import zipfile
from dataclasses import dataclass, field
from decimal import ROUND_DOWN, Decimal
from pathlib import Path
from typing import Callable, Iterable, Optional, Sequence, Union
from xml.etree import ElementTree as ET

from . import datasource, paths, quirks, vbx
from .flow import ERROR, INFO, WARN, Message

# ------------------------------------------------------------------
# 色(VBA の RGB 値)
# ------------------------------------------------------------------
NG_COLOR = 8421631            # RGB(255,128,128) 公差外
OK_COLOR = 13434828           # RGB(204,255,204) 合格
WHITE = 16777215              # RGB(255,255,255)
CYAN = 16776960               # RGB(0,255,255) 選択中(外観・不良要因)
GREEN = 65280                 # RGB(0,255,0) CheckT? 入力済み
RED = 255                     # RGB(255,0,0) CheckT? 未入力・異常ロット表示
WINDOW_BG = 0x80000005        # &H80000005(システムのウィンドウ背景色)

POSITIONS = ("S", "U", "C", "D", "E", "F")


class VbaRuntimeError(RuntimeError):
    """VBA の実行時エラー(エラー処理なしで止まる所)。`number` は VBA のエラー番号。"""

    def __init__(self, number: int, text: str) -> None:
        super().__init__(f"実行時エラー {number}: {text}")
        self.number = number


# ------------------------------------------------------------------
# 画面の値
# ------------------------------------------------------------------
@dataclass
class Uf1:
    """UF1 のうち、この章の手続きが読む値(読むだけ)。"""
    text: dict[str, str] = field(default_factory=dict)       # 名前 → Caption / Text
    visible: dict[str, bool] = field(default_factory=dict)   # 名前 → Visible
    ng: set[str] = field(default_factory=set)                # 背景が NG 色の名前
    work: str = ""               # 作業選択(OptionButton): 厚板 / 薄板 / 少枚数 / 極厚 / スラブ
    line_name: str = ""          # LineName.Caption(VBA の書き方: L-1 / ﾊﾞﾗﾝｻｰ / ﾄｯﾄ / 丸徳 …)
    line_no: str = ""            # LineNo.Caption
    worker_day_max: str = ""     # Worker.WorkeDayMax の戻り値("1"/"2"/"3")

    def t(self, name: str) -> str:
        v = self.text.get(name, "")
        return "" if v is None else str(v)

    def vis(self, name: str) -> bool:
        return bool(self.visible.get(name, False))

    def is_ng(self, name: str) -> bool:
        return name in self.ng

    @property
    def lname(self) -> str:
        """VBA の `Lname = LineName.Caption & LineNo.Caption`。"""
        return self.line_name + self.line_no


@dataclass
class WorkDist:
    """作業保存シートの歪(O2:P7 = A 面、O8:P13 = B 面。行は S,U,C,D,E,F、列は 定盤, ゲージ)。

    値は "幅_上_0.30" のような文字列。空セルは "" か None、数のセルは数のまま渡してよい。
    """
    a: Sequence[Sequence[object]]
    b: Sequence[Sequence[object]]


def _cell_text(v: object) -> str:
    """シートのセル(Variant)を文字列と比べるときの形(Empty → ""、数 → CStr)。"""
    if v is None:
        return ""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return vbx.cstr(v)
    return str(v)


def _num_eq(text: object, n: float) -> bool:
    """VBA の `文字列 = 数のリテラル`(TextBox の値・String 変数と数)。

    数にならない文字列(空文字を含む)は U1 のスイッチに従う(既定は一致しない。真なら `vbx.TypeMismatch`)。
    """
    return vbx.u1_compare("" if text is None else str(text), n) == 0


def _case_in(text: str, values: Iterable[float]) -> bool:
    """`Select Case text` の `Case 2, 4, …`(左から順に比べ、最初の一致で真)。"""
    for v in values:
        if _num_eq(text, v):
            return True
    return False


def _case_between(text: str, lo: float, hi: float) -> bool:
    """`Case lo To hi`。"""
    s = "" if text is None else str(text)
    return vbx.u1_compare(s, lo) >= 0 and vbx.u1_compare(s, hi) <= 0


_LOG_FOLDERS = {"L-1": "L-1", "LVC": "LVC", "HVC": "HVC", "ﾊﾞﾗﾝｻｰ": "バランサー", "ﾄｯﾄ": "トット", "Test": "Test"}


def log_folder(lname: str) -> str:
    """ログの置き場所(`Select Case Lname`。L-1/LVC/HVC/ﾊﾞﾗﾝｻｰ/ﾄｯﾄ/Test、それ以外は 丸徳)。

    `HistoryPath & <folder> & "\\" & <folder> & "_Log"` のフォルダ名。Lname は LineName & LineNo の完全一致
    (号機付きの名前はすべて「丸徳」になる。VBA のとおり)。(O:784-792 ほか)
    """
    return _LOG_FOLDERS.get(lname, "丸徳")


def last_worker(uf: Uf1) -> Optional[str]:
    """`Split(UF1.T?, "_")` の最後(WorkeDayMax が "1"/"2"/"3" のとき)。それ以外は None(画面を変えない)。

    VBA は Split の結果が空配列(T? が "")だと `Worker(UBound(Worker))` が添字エラー(9)。
    B48 を直すと None(`If IsEmpty(Worker) = False` の意図どおり KenName を変えない)。
    (O:633-640, 662-669, 705-712)
    """
    key = {"1": "T1", "2": "T2", "3": "T3"}.get(uf.worker_day_max)
    if key is None:
        return None
    parts = uf.t(key).split("_") if uf.t(key) != "" else []
    if not parts:
        if quirks.reproduce("B48"):
            raise VbaRuntimeError(9, "インデックスが有効範囲にありません(検査員欄が空)")
        return None
    return parts[-1]


# ------------------------------------------------------------------
# 重量・異常値
# ------------------------------------------------------------------
def _excel_rounddown(x: float, digits: int) -> float:
    """Excel の ROUNDDOWN(有効 15 桁の 10 進にしてから 0 の方向へ切り捨て)。"""
    if x == 0:
        return 0.0
    d = Decimal(vbx.cstr(float(x)))
    return float(d.quantize(Decimal(1).scaleb(-digits), rounding=ROUND_DOWN))


def weight(thickness: object, width: object, length: object) -> float:
    """重量 kg。`OutTolerance.Weight (O:813-817)`

    比重 2.68、`Application.RoundDown(厚 * 幅 * 長 * 2.68 / 10 ^ 6, 3)`。
    どれかが数でない(IsNumeric が偽)・"" なら 0。
    """
    vals = (thickness, width, length)
    if not all(vbx.is_numeric(v) for v in vals):
        return 0.0
    if any(isinstance(v, str) and v == "" for v in vals):
        return 0.0
    t, w, ln = (vbx.cdbl(v) for v in vals)
    return _excel_rounddown(t * w * ln * 2.68 / 10.0 ** 6, 3)


@dataclass(frozen=True)
class AbnormalLimits:
    """異常値のしきい値(下限未満・上限超えが異常)。"""
    atu1: float
    atu2: float
    haba1: float
    haba2: float
    take1: float
    take2: float
    tai1: float
    tai2: float


def abnormal_limits(atu: object, hab: object, tak: object, am: object, ap: object,
                    hm: object, hp: object, tm: object, tp: object) -> AbnormalLimits:
    """公差の 2 倍・対角は √(巾²+丈²)±6。`異常値除去 (O:1324-1347)`、`UF異常値発見.Initialize (:47-81)`

    ATU は Double、巾・丈は Single(`HABA = Val(UF1.Hab)`)。Format の結果を Double に入れ直す。
    """
    am2, ap2, hm2, hp2, tm2, tp2 = (vbx.val(x) * 2 for x in (am, ap, hm, hp, tm, tp))
    a = vbx.val(atu)
    h = vbx.single(vbx.val(hab))
    k = vbx.single(vbx.val(tak))
    diag = math.sqrt(vbx.val_single(h) ** 2 + vbx.val_single(k) ** 2)
    f3 = lambda x: vbx.cdbl(vbx.fmt(x, "0.000"))     # noqa: E731
    f1 = lambda x: vbx.cdbl(vbx.fmt(x, "0.0"))       # noqa: E731
    return AbnormalLimits(f3(a - am2), f3(a + ap2), f1(h - hm2), f1(h + hp2),
                          f1(k - tm2), f1(k + tp2), f1(diag - 6), f1(diag + 6))


def era_flag(caption: object, kp: float, km: float) -> bool:
    """`Era_flag (O:807-811)`: `Val(TB) < KM Or Val(TB) > KP`(境界ちょうどは正常)。"""
    v = vbx.val(caption)
    return v < km or v > kp


@dataclass(frozen=True)
class AbnormalHit:
    """異常値が見つかった(VBA: Abno.Tag="NG" にして UF異常値発見 をモーダルで出す。AbnoFormStart O:796-805)。"""
    name: str           # ラベル名(UF異常値発見.TBname)
    value: str          # 値(UF異常値発見.IJO)
    kind: str           # 厚 / 巾 / 丈 / 対角
    lower: float
    upper: float


def abnormal_check(uf: Uf1, name: str) -> Optional[AbnormalHit]:
    """測定値ラベル 1 つの異常値判定。`異常値除去 (O:1321-1415)`

    `UF1.Abno.Caption = "異常値on"` のときだけ。厚=1〜4、巾=5,6(極厚は 5〜10)、丈=7,8(極厚は 11〜14)、
    対角=9,10(極厚は見ない)。マイナス公差(Am/Hm/Tm)が空なら判定しない。
    """
    if uf.t("Abno") != "異常値on":
        return None
    lim = abnormal_limits(uf.t("ATU"), uf.t("Hab"), uf.t("Tak"), uf.t("Am"), uf.t("Ap"),
                          uf.t("Hm"), uf.t("Hp"), uf.t("Tm"), uf.t("Tp"))
    cap = uf.t(name)
    if cap == "":
        return None
    gokuatsu = uf.work == "極厚"

    def hit(kind: str, kp: float, km: float) -> Optional[AbnormalHit]:
        return AbnormalHit(name, cap, kind, km, kp) if era_flag(cap, kp, km) else None

    # 厚
    if len(name) == 2 and _case_in(vbx.right(name, 1), (1, 2, 3, 4)):
        if uf.t("Am") == "":
            return None
        r = hit("厚", lim.atu2, lim.atu1)
        if r:
            return r
    # 巾
    if gokuatsu:
        if len(name) == 2:
            if _case_between(vbx.right(name, 1), 5, 9):
                if uf.t("Hm") == "":
                    return None
                r = hit("巾", lim.haba2, lim.haba1)
                if r:
                    return r
        elif len(name) == 3:
            if _case_in(vbx.right(name, 2), (10,)):
                if uf.t("Hm") == "":
                    return None
                r = hit("巾", lim.haba2, lim.haba1)
                if r:
                    return r
    else:
        if len(name) == 2 and _case_in(vbx.right(name, 1), (5, 6)):
            if uf.t("Hm") == "":
                return None
            r = hit("巾", lim.haba2, lim.haba1)
            if r:
                return r
    # 丈
    if gokuatsu:
        if len(name) == 3 and _case_in(vbx.right(name, 2), (11, 12, 13, 14)):
            if uf.t("Tm") == "":
                return None
            r = hit("丈", lim.take2, lim.take1)
            if r:
                return r
    else:
        if len(name) == 2 and _case_in(vbx.right(name, 1), (7, 8)):
            if uf.t("Tm") == "":
                return None
            r = hit("丈", lim.take2, lim.take1)
            if r:
                return r
    if gokuatsu:
        return None
    # 対角(マイナス公差は丈の Tm を見る)
    if len(name) == 3:
        if _case_in(vbx.right(name, 2), (10,)):
            if uf.t("Tm") == "":
                return None
            return hit("対角", lim.tai2, lim.tai1)
    elif len(name) == 2:
        if _case_in(vbx.right(name, 1), (9,)):
            if uf.t("Tm") == "":
                return None
            return hit("対角", lim.tai2, lim.tai1)
    return None


def abnormal_form_text(saved_row1: Sequence[object]) -> dict[str, str]:
    """UF異常値発見 の表示文。`UF異常値発見.UserForm_Initialize (:47-81)`

    `saved_row1` = 作業保存シート 1 行目の B〜J(厚, 巾, 丈, Am, Ap, Hm, Hp, Tm, Tp)。
    """
    v = list(saved_row1) + [""] * 9
    lim = abnormal_limits(*v[:9])
    return {
        "TEIGI": "異常値の閾値　公差の2.0倍とする",
        "TEIGITAI": "対角線　a^2＋b^2±6で比較",
        "ATU": "厚 " + vbx.fmt(lim.atu1, "0.000") + "未満 " + vbx.fmt(lim.atu2, "0.000") + "超え",
        "HABA": "巾 " + vbx.fmt(lim.haba1, "0.0") + "未満 " + vbx.fmt(lim.haba2, "0.0") + "超え",
        "TAKE": "丈 " + vbx.fmt(lim.take1, "0.0") + "未満 " + vbx.fmt(lim.take2, "0.0") + "超え",
        "TAIK": "対角線 " + vbx.fmt(lim.tai1, "0.0") + "未満 " + vbx.fmt(lim.tai2, "0.0") + "超え",
    }


@dataclass(frozen=True)
class EraseAction:
    """画面の値を消す指示。"""
    clear: tuple[str, ...] = ()                       # 値を "" にするコントロール(順に)
    erase_count: tuple[str, ...] = ()                 # FormRelat.EraseCount を呼ぶラベル(順に)
    back_color: tuple[tuple[str, int], ...] = ()      # 背景色を変える(名前, 色)
    clear_inp: bool = False                           # UF1.INP = ""
    posting: bool = False                             # SaveLoad.Posting を呼ぶ
    rejudge: bool = False                             # LoopJudg → HizLoopJudg(0)
    look_blank: Optional[tuple[int, int]] = None      # Look_Blank(1, 10) / (1, 14)
    close_confi: bool = False                         # UF1 を出して UFConfi を閉じる


def ng_back_tb(name: str) -> EraseAction:
    """異常値の値を消す。`NG_Back_TB (O:1416-1425)`(UF異常値発見 のボタン)"""
    return EraseAction(clear=(name,), back_color=((name, WINDOW_BG),), clear_inp=True, posting=True)


# ------------------------------------------------------------------
# 公差外の確認(UFConfi)
# ------------------------------------------------------------------
@dataclass(frozen=True)
class ConfiRequest:
    """UFConfi に出す内容。`Confi_Open (O:673-715)` / `ConfiReason_Open (O:646-672)`

    n: 測定欄の番号(1〜14)、96=ラテラルボー、97=直角度、98=定盤歪、99=ゲージ歪。
    page: 0=公差外(MP1 のページ0。P2 を隠す)、1=理由選択(P1 を隠す)。
    """
    posi: str                        # Kpop
    n: int                           # Knum
    line_name: str                   # LineName(= Lname)
    moto: str                        # MotoData.Caption
    return_moto: Optional[str]       # return_MotoData.Caption(理由選択では変えない = None)
    del_count: int                   # DelC.Caption
    page: int
    ken_name: Optional[str]          # KenName(None = 変えない)


def confi_open(uf: Uf1, posi: str, n: int, lname: str, moto: str, del_count: int = 0) -> ConfiRequest:
    """公差外の確認フォームの内容。`Confi_Open (O:673-715)`

    厚板・薄板・少枚数で n が 9 または 10(対角)なら、元データを「線差 」& Diag? の値にする。
    """
    m = moto
    if uf.work in ("厚板", "薄板", "少枚数") and n in (9, 10):
        m = "線差 " + uf.t("Diag" + posi)
    return ConfiRequest(posi, n, lname, m, m, del_count, 0, last_worker(uf))


def confi_reason_open(uf: Uf1, posi: str, n: int, lname: str, moto: str, del_count: int = 0) -> ConfiRequest:
    """削除理由のフォームの内容。`ConfiReason_Open (O:646-672)`"""
    return ConfiRequest(posi, n, lname, moto, None, del_count, 1, last_worker(uf))


def out_tolerance(uf: Uf1, temp: str = "") -> list[ConfiRequest]:
    """入力後に公差外を探して確認フォームを出す。`Out_Tolerance (O:818-977)`

    戻り値は Confi_Open を呼んだ順。UFConfi は 1 つしかなく vbModeless で出すので、
    **画面に残るのは最後の 1 件**(G2 §12 Q3 実機確認待ち)。
    その行の NG? ラベルが「許容」なら確認しない(「保留」は確認する)。
    """
    if uf.t("INP.Tag") == "":
        return []
    st = uf.t("INP.Tag")
    if temp != "":
        st = temp
    lname = uf.lname
    alp = ""
    if len(st) == 2:
        alp = vbx.right(st, 1)
        if vbx.is_numeric(alp):
            alp = vbx.left(st, 1)
    elif len(st) == 3:
        alp = vbx.right(st, 2)
        if vbx.is_numeric(alp):
            alp = vbx.left(st, 1)
    if vbx.instr(st, "Lateral") > 0:
        alp = vbx.right(st, 1)
    if vbx.instr(st, "Diag") > 0:
        alp = vbx.mid(st, 5, 1)
    rig = ""
    if len(st) == 3:
        rig = vbx.right(st, 2)
    elif len(st) == 2:
        rig = vbx.right(st, 1)
    posi = alp
    if posi not in POSITIONS:
        # UF1("NG" & Posi) のコントロールが無い
        raise VbaRuntimeError(-2147024809, f"指定されたオブジェクトが見つかりません: NG{posi}")
    ng_posi = uf.t("NG" + posi)
    dc = vbx.clng(vbx.val(uf.t("DeleteC")))
    out: list[ConfiRequest] = []

    def scan(last: int, rows: tuple[str, ...]) -> None:
        for i in range(1, last + 1):
            if uf.t("NG" + posi) == "許容":
                break
            if posi in rows and uf.is_ng(f"{posi}{i}"):
                out.append(confi_open(uf, posi, i, lname, uf.t(f"{posi}{i}"), dc))

    if uf.work == "薄板":
        if _case_in(rig, (2, 4, 6, 8, 10)):
            scan(10, POSITIONS)
    elif uf.work == "スラブ":
        if _case_in(rig, (2, 4, 6, 8, 10)):
            scan(10, POSITIONS[:5])
    elif uf.work in ("厚板", "少枚数"):
        if _case_in(rig, (4, 10)):
            scan(10, POSITIONS)
    else:
        if _case_in(rig, (4, 7, 10, 12, 14)):
            scan(14, POSITIONS[:5])

    if ng_posi != "許容":
        for p in POSITIONS:
            name = "Lateral" + p
            if st == name and uf.vis(name) and uf.is_ng(name):
                out.append(confi_open(uf, p, 96, lname, uf.t(name), dc))
        for p in POSITIONS:
            name = f"Diag{p}Kou"
            if st == name and uf.vis(name) and uf.is_ng(name):
                out.append(confi_open(uf, p, 97, lname, uf.t(name), dc))
    if ng_posi != "許容":
        for p in POSITIONS:
            if st == "Z" + p and uf.is_ng("Z" + p):
                out.append(confi_open(uf, p, 98, lname, uf.t("Z" + p), dc))
        for p in POSITIONS:
            if st == "G" + p and uf.is_ng("G" + p):
                out.append(confi_open(uf, p, 99, lname, uf.t("G" + p), dc))
    return out


@dataclass(frozen=True)
class ConfiDecision:
    """UFConfi の公差外ページのボタン(CommandButton1)で何をするか。"""
    action: str                 # posting_erase / pass / back / hold / need_reason / error
    message: Optional[Message] = None
    reasons: tuple[str, ...] = ()


CONFI_REASONS_2 = ("入力位置間違い", "測定位置間違い", "測定ボタン押し間違い", "0点確認", "その他")


def confi_decide(tb1: str, del_count_caption: str) -> ConfiDecision:
    """公差外の扱い(TB1 の文字)。`UFConfi.CommandButton1_Click (UFConfi:SkipA 以降)`

    公差外 → NG_Posting → NG_Erase、許容 → NG_Pass("許容")、保留 → NG_Pass("保留")、
    再測定 → 同じ位置を 2 回以上消していれば理由を選ばせる(Q-S-QA-DIM-174 より厳しく 2 回)、それ以外は NG_Back。
    """
    if tb1 == "公差外":
        return ConfiDecision("posting_erase")
    if tb1 == "許容":
        return ConfiDecision("pass")
    if tb1 == "再測定":
        if vbx.val(del_count_caption) >= 2:
            return ConfiDecision("need_reason", Message("理由を選択してください", WARN, source="UFConfi"),
                                 CONFI_REASONS_2)
        return ConfiDecision("back")
    if tb1 == "保留":
        return ConfiDecision("hold")
    return ConfiDecision("error", Message("指定文字列を入力してください", ERROR, source="UFConfi"))


def ng_pass(posi: str, temp: str = "許容") -> dict[str, str]:
    """NG? ラベルに「許容」「保留」を書く。`NG_Pass (O:1157-1170)`(この後 UF1 を出して UFConfi を閉じる)"""
    return {"NG" + posi: temp} if posi in POSITIONS else {}


def ng_back(uf: Uf1, posi: str, n: int) -> EraseAction:
    """再測定: NG 色の欄だけ消し、EraseCount に記録して再判定。`NG_Back (O:1053-1155)`

    n=96/97/98/99 はラテラルボー/直角度/定盤歪/ゲージ歪の欄も消す(EraseCount はしない)。
    """
    clear: list[str] = []
    counted: list[str] = []
    if posi in POSITIONS[:5]:
        for i in range(1, 15):
            name = f"{posi}{i}"
            if uf.is_ng(name):
                clear.append(name)
                counted.append(name)
    if posi == "F":
        for i in range(1, 11):
            name = f"F{i}"
            if uf.is_ng(name):
                clear.append(name)
                counted.append(name)
    if posi in POSITIONS:
        extra = {96: "Lateral" + posi, 97: f"Diag{posi}Kou", 98: "Z" + posi, 99: "G" + posi}.get(n)
        if extra:
            clear.append(extra)
    return EraseAction(clear=tuple(clear), erase_count=tuple(counted), rejudge=True,
                       look_blank=(1, 14) if uf.work == "極厚" else (1, 10), close_confi=True)


@dataclass(frozen=True)
class NgEraseAction:
    """公差外の行を作業保存シートから消して読み直す。`NG_Erase (O:978-1052)`"""
    ranges: tuple[tuple[int, int, int, int], ...]     # (行1, 列1, 行2, 列2) を ClearContents
    split_z_tag: Optional[str]                         # SaveLoad.Split_Z_Tag_0(位置)
    reload: bool = True                                # Load_Text(作業保存の読み込み)
    close_confi: bool = True                           # All="" かつ UFConfi 表示中なら UF1 を出して閉じる


def ng_erase(posi: str, all_: str = "") -> NgEraseAction:
    """作業保存_<Lname> の消す範囲: 行(2〜7 の A:T)、外観BH(10〜15 の A:C)、歪B(8〜13 の O:P)。"""
    if posi not in POSITIONS:
        return NgEraseAction((), None, True, all_ == "")
    k = POSITIONS.index(posi)
    return NgEraseAction(((2 + k, 1, 2 + k, 20), (10 + k, 1, 10 + k, 3), (8 + k, 15, 8 + k, 16)),
                         posi, True, all_ == "")


@dataclass
class NgRow:
    """異常ロット用シートの 1 行(列 1〜18 = 値、19 = 位置)。"""
    values: list[str]
    ng_cols: set[int] = field(default_factory=set)    # NG 色を塗った列(1 始まり)


@dataclass
class AbnormalSheet:
    """シート「異常ロット用_<Lname>」(A1 = ロット、2 行目から公差外の行)。"""
    lot: str
    rows: list[NgRow] = field(default_factory=list)


def ng_posting(uf: Uf1, posi: str, sheet: Optional[AbnormalSheet]) -> AbnormalSheet:
    """公差外の行を異常ロット用シートへ写す。`NG_Posting (O:1171-1320)`

    列 1〜14 = 位置の 1〜14(F は 1〜10)、15 = 直角度(Diag?Kou)、16 = ラテラルボー、
    17 = 定盤歪("A:"/"B:" & 値)、18 = ゲージ歪("ｹﾞ:" & 値)、19 = 位置。NG 色の欄は同じ列を塗る。
    シートが無い・A1 がロットと違うときは作り直す。
    書く行は「A 列の最後の値がある行 + 1」(`UR(sh, 1) + 1`)。そのため **前の行の A 列が空だと
    その行に上書きする**(値は置き換え、色は前の行の色が残る。VBA のとおり。B44)。
    B44 を直すと、何か書いてある最後の行(19 列目の位置は必ず書く)の次に書く。
    この後、画面の IzyouFl の背景を赤(RED)にする。
    """
    lot = uf.t("Lot")
    if sheet is None or sheet.lot != lot:
        sheet = AbnormalSheet(lot, [])
    last = 1
    for idx, row in enumerate(sheet.rows):
        if quirks.reproduce("B44"):
            used = bool(row.values) and row.values[0] != ""
        else:
            used = any(v != "" for v in row.values)
        if used:
            last = idx + 2
    lr = last + 1
    ng = [""] * 18
    if posi in POSITIONS[:5]:
        for i in range(1, 15):
            ng[i - 1] = uf.t(f"{posi}{i}")
    if posi == "F":
        for i in range(1, 11):
            ng[i - 1] = uf.t(f"F{i}")
    if posi in POSITIONS:
        if uf.t("G" + posi) != "":
            ng[17] = "ｹﾞ:" + uf.t("G" + posi)
        if uf.t("AorB") == "Now_A" and uf.t("Z" + posi) != "":
            ng[16] = "A:" + uf.t("Z" + posi)
        if uf.t("AorB") == "Now_B" and uf.t("Z" + posi) != "":
            ng[16] = "B:" + uf.t("Z" + posi)
        ng[14] = uf.t(f"Diag{posi}Kou")
        ng[15] = uf.t("Lateral" + posi)
    colored: set[int] = set()
    if posi in POSITIONS[:5]:
        colored |= {i for i in range(1, 15) if uf.is_ng(f"{posi}{i}")}
    if posi == "F":
        colored |= {i for i in range(1, 11) if uf.is_ng(f"F{i}")}
    if posi in POSITIONS:
        for col, name in ((18, "G" + posi), (17, "Z" + posi), (16, "Lateral" + posi), (15, f"Diag{posi}Kou")):
            if uf.is_ng(name):
                colored.add(col)
    idx = lr - 2
    while len(sheet.rows) <= idx:
        sheet.rows.append(NgRow([""] * 19, set()))
    row = sheet.rows[idx]
    row.values = ng + [posi]
    row.ng_cols |= colored
    return sheet


# ------------------------------------------------------------------
# 選択肢(ComboBox / ListBox)
# ------------------------------------------------------------------
def cb_shape_items(atu: float = 6.0) -> list[str]:
    """定盤歪の形状(CB?1)。厚 6.00mm 以下は 幅/丈/耳/中、超えると 幅/丈。`CBAdd (O:1427-1460)`"""
    return ["幅", "丈", "耳", "中"] if vbx.single(atu) <= 6.0 else ["幅", "丈"]


def cbg_shape_items(lg_captions: dict[str, str]) -> list[str]:
    """ゲージ歪の形状(CB?3)。LG? のどれかが「ﾌﾘｰｿﾘ」なら 上/下、それ以外は 幅/丈。`CBAdd (O:1462-1485)`"""
    free = any(lg_captions.get("LG" + p, "") == "ﾌﾘｰｿﾘ" for p in POSITIONS)
    return ["上", "下"] if free else ["幅", "丈"]


def cb_updown_items(shape: str) -> list[str]:
    """定盤歪の上下・型(CB?2)。`CB_UpDownAdd (O:1487-1502)`"""
    if shape in ("幅", "丈"):
        return ["上", "下"]
    if shape == "中":
        return ["Ⅰ型", "Ⅱ型", "Ⅲ型"]
    if shape == "耳":
        return ["Ⅱ型", "Ⅲ型"]
    return []


def cbg_updown_items(shape: str) -> list[str]:
    """ゲージ歪の上下(CB?4)。`CBG_UpDownAdd (O:1504-1511)`"""
    return ["上", "下"] if shape in ("幅", "丈") else []


YAMA_ITEMS = ("1", "2", "3", "4", "5", "6", "枚数")
DELETE_REASONS = ("入力位置間違い", "測定位置間違い", "測定ボタン押し間違い", "再測定", "0点確認", "板廃棄", "その他")
SPOT = "ｽﾎﾟｯﾄ処理"


@dataclass
class VisualLists:
    """外観タブ・検入高さ・不良タブの選択肢。None = その ListBox は設定しない(VBA の Exit Sub)。"""
    visual: Optional[list[str]] = None                     # LBVisual(先頭は空白行)
    yama: Optional[list[str]] = None                       # YamaLB
    defect: Optional[list[tuple[str, str]]] = None         # DefectFactorLB(名前, コード)
    error: Optional[Message] = None


def defect_factor_items(names_row: Sequence[object], codes_row: Sequence[object]) -> list[tuple[str, str]]:
    """不良要因の選択肢。原紙「不良要因印刷原紙」の C2:X2(名前)と C9:X9(コード)を縦にして、
    最後に「ｽﾎﾟｯﾄ処理」(コードなし)。`Visual_Add (O:1590-1607)`
    """
    names = [_cell_text(v) for v in names_row]
    codes = [_cell_text(v) for v in codes_row]
    n = max(len(names), len(codes))
    names += [""] * (n - len(names))
    codes += [""] * (n - len(codes))
    return list(zip(names, codes)) + [(SPOT, "")]


def visual_add(names_row: Optional[Sequence[object]] = None,
               codes_row: Optional[Sequence[object]] = None) -> VisualLists:
    """外観異常名(品質基準「欠陥名」)・山数・不良要因の選択肢。`Visual_Add (O:1513-1614)`

    names_row/codes_row が None = 原紙シートが無い(不良要因は設定しない)。
    DB を読めないときは VBA の ErrHandler と同じ文言を返す。
    """
    out = VisualLists()
    step = "欠陥名フィールド取得"
    try:
        fields = datasource.get_fields("品質基準", "欠陥名", "")
        if fields is None:
            return out
        step = "欠陥名フィールド検索"
        idx = datasource.column_index(fields, "欠陥名")
        if idx == -1:
            return out
        step = "欠陥名レコード取得"
        recs = datasource.get_records("品質基準", "欠陥名", "")
        if recs is None:
            return out
    except (datasource.DbMissing, datasource.TableMissing) as e:
        out.error = Message("外観異常名登録でエラーが発生しました\n段階: " + step + "\n内容: " + str(e),
                            ERROR, source="OutTolerance:1611")
        return out
    out.visual = [""] + [r[idx] for r in recs if len(r[idx]) > 0]
    out.yama = list(YAMA_ITEMS)
    if names_row is None or codes_row is None:
        return out
    out.defect = defect_factor_items(names_row, codes_row)
    return out


def get_selected_items(items: Sequence[Union[str, Sequence[str]]], selected: Iterable[int],
                       defect: bool = False) -> str:
    """ListBox で選んだ項目をスペースでつなぐ(不良要因は「名前_コード」)。`GetSelectedItems (O:1705-1739)`

    VBA の比較 `.List(i) = SelectedItems` も同じにする(1 件目が空白行なら飛ばす、2 件目以降の空白行は
    " " だけ足す)。最後が "_" なら 1 文字消す。
    """
    sel = set(selected)

    def col(i: int, c: int) -> str:
        it = items[i]
        if isinstance(it, str):
            return it if c == 0 else ""
        return str(it[c]) if c < len(it) and it[c] is not None else ""

    s = ""
    for i in range(len(items)):
        if i not in sel:
            continue
        piece = (col(i, 0) + "_" + col(i, 1)) if defect else col(i, 0)
        if s == "":
            if not (col(i, 0) == s):
                s = piece
        else:
            if not (col(i, 0) == s):
                s = s + " " + piece
    if s == "":
        return ""
    if vbx.right(s, 1) == "_":
        s = vbx.left(s, len(s) - 1)
    return s


_VISUAL_TARGETS = {1: "S_UpVisual", 2: "U_UpVisual", 3: "C_UpVisual", 4: "D_UpVisual", 5: "E_UpVisual",
                   6: "F_UpVisual", 7: "S_UnderVisual", 8: "U_UnderVisual", 9: "C_UnderVisual",
                   10: "D_UnderVisual", 11: "E_UnderVisual", 12: "F_UnderVisual"}


@dataclass(frozen=True)
class LcColorResult:
    cyan: frozenset[int]               # 水色の LC 番号
    captions: dict[str, str]           # 書く外観ラベル → 文字
    target: str                        # ログに出す欄(最後に書いた欄)
    log_text: str                      # WriteLog の文(Measured.log と既定の出力先の 2 か所)
    log_folder: str


def lc_color(label_no: int, n: int, selected_text: str, lname: str, worker: str,
             cyan_before: Iterable[int] = ()) -> LcColorResult:
    """外観の入力欄を選んで、選んだ欠陥名を書く。`LcColor (O:1628-1704)`

    UF1 は LC1〜6 を n=1、LC7〜12 を n=7 で呼ぶ(どちらも先に全部白に戻す)。押した LC の Caption は空にする。
    """
    cyan = set(cyan_before)
    if n in (1, 7):
        cyan = set()
    cyan.add(label_no)
    caps: dict[str, str] = {}
    tb = ""
    for k in range(1, 7):
        if k in cyan:
            caps[_VISUAL_TARGETS[k]] = selected_text
            tb = _VISUAL_TARGETS[k]
            break
    for k in range(7, 13):
        if k in cyan:
            caps[_VISUAL_TARGETS[k]] = selected_text
            tb = _VISUAL_TARGETS[k]
            break
    log = "*UFKenl" + tb + "l" + selected_text + "l" + lname + "|" + worker
    return LcColorResult(frozenset(cyan), caps, tb, log, log_folder(lname))


def tec_return(cyan: Iterable[str]) -> str:
    """水色の不良枚数欄(SBad〜FBad の最後に当たったもの)。`TecReturn (O:2471-2494)`"""
    c = set(cyan)
    out = ""
    for name in ("SBad", "UBad", "CBad", "DBad", "EBad", "FBad"):
        if name in c:
            out = name
    return out


# ------------------------------------------------------------------
# 検入高さ
# ------------------------------------------------------------------
@dataclass(frozen=True)
class YamaLayout:
    visible: dict[str, bool]
    hl_captions: dict[str, str]
    formula_reset: bool = True       # 高さ計算式.Caption = ""


def yama_layout(selected: str, mitsubishi: bool = False) -> Optional[YamaLayout]:
    """山数で検入高さの入力欄を出し分ける。`DrawListBoxContent (O:1747-1840)`(この後 CalculateSheets)

    選択が "" なら None(何もしない)。
    """
    if selected == "":
        return None
    vis: dict[str, bool] = {}
    caps: dict[str, str] = {}
    for i in range(1, 7):
        vis[f"HL{i}"] = False
    for i in range(1, 8):
        vis[f"KenHBox{i}"] = False
    tem_a = 7 if mitsubishi else 10
    for i in range(1, tem_a + 1):
        vis[f"KenHL{i}"] = False
    if not mitsubishi:
        vis["KenCount"] = False
    if selected in ("1", "2", "3", "4", "5", "6"):
        k = int(selected)
        for i in range(1, k + 1):
            vis[f"HL{i}"] = True
            vis[f"KenHBox{i}"] = True
            vis[f"KenHL{i}"] = True
            caps[f"HL{i}"] = "1" if k == 1 else f"{k}-{k + 1 - i}"
    elif selected == "枚数":
        vis["KenHBox7"] = True
        if not mitsubishi:
            vis["KenCount"] = True
        for i in range(7, tem_a + 1):
            vis[f"KenHL{i}"] = True
    return YamaLayout(vis, caps)


@dataclass(frozen=True)
class SheetsResult:
    quantity: float                  # 戻り値(Single)
    formula: Optional[str] = None    # 高さ計算式.Caption(None = 変えない)
    tag7: Optional[str] = None       # UFMitsubishi の KenHBox7.Tag(None = 変えない)


def calculate_sheets(boxes: Sequence[str], thickness: float, mitsubishi: bool = False,
                     box7: str = "") -> SheetsResult:
    """検入高さ → 枚数 Σ(高さ ÷ 厚)。`CalculateSheets (O:2509-2564)`

    `boxes` = KenHBox1〜6 の文字。厚は Single。数の欄が無い・厚 0 なら 0(式は変えない)。
    Quantity は Single に足していき、戻り値は `Format(Quantity, "0.0000")` を Single にしたもの。
    式は `(h÷厚) + (h÷厚)=` & Quantity(Single の CStr)。
    """
    th = vbx.single(thickness)
    if th == 0:
        return SheetsResult(0.0)
    vals = [("" if v is None else str(v)) for v in list(boxes)[:6]] + [""] * max(0, 6 - len(boxes))
    nums = [v for v in vals if vbx.is_numeric(v)]
    if not nums:
        if mitsubishi and box7 != "":
            return SheetsResult(0.0, None, box7)
        return SheetsResult(0.0)
    q = 0.0
    for v in nums:
        q = vbx.single(q + vbx.cdbl(v) / th)
    ret = vbx.single(vbx.cdbl(vbx.fmt(q, "0.0000")))
    th_s = vbx.cstr(th, is_single=True)
    formula = "".join("(" + v + "÷" + th_s + ")" + " + " for v in nums)
    formula = vbx.left(formula, len(formula) - 3)
    return SheetsResult(ret, formula + "=" + vbx.cstr(q, is_single=True))


def h_check(boxes: Sequence[str]) -> str:
    """検入高さ KenHBox1〜7 がすべて空なら "入力なし"。`H_Check (O:2158-2170)`"""
    for v in list(boxes)[:7]:
        if v is not None and str(v) != "":
            return ""
    return "入力なし"


# ------------------------------------------------------------------
# 平坦度(ゲージ要否)
# ------------------------------------------------------------------
def _dist_flags(dist: WorkDist) -> tuple[list[list[str]], list[list[str]]]:
    def clean(rows: Sequence[Sequence[object]]) -> list[list[str]]:
        out = []
        for r in list(rows)[:6]:
            r = list(r) + [None, None]
            row = []
            for v in r[:2]:
                s = _cell_text(v)
                row.append("" if s in ("_", "__") else s)
            out.append(row)
        while len(out) < 6:
            out.append(["", ""])
        return out
    return clean(dist.a), clean(dist.b)


def withor_without_gauge(uf: Uf1, dist: Optional[WorkDist]) -> str:
    """平坦度測定基準(Q-S-QA-DIM-173)の未入力チェック(印刷前)。`WithorWithout_Gauge (O:1842-1983)`

    35kg 以下(パターンA)は両面の定盤、35kg 超え(パターンB)は片面ならゲージが要る。
    後から代入したメッセージが残る。パターンB で 1 か所でも両面の定盤があれば、それまでのメッセージで終わる。
    B10: C 行の文言が「2の…」、B 面のゲージは F 行を見ない(VBA どおり)。
    """
    tw = weight(vbx.val(uf.t("ATU")), vbx.val(uf.t("Hab")), vbx.val(uf.t("Tak")))
    if tw == 0:
        return ""
    if dist is None:
        return ""
    a, b = _dist_flags(dist)
    yama4 = uf.t("Yama") == "4"
    fa = [a[i][0] != "" for i in range(6)]
    fb = [b[i][0] != "" for i in range(5)] + [yama4 and b[5][0] != ""]
    fag = [a[i][1] != "" for i in range(5)] + [yama4 and a[5][1] != ""]
    fbg = [b[i][1] != "" for i in range(5)] + [yama4 and b[5][1] != ""]
    atu = [uf.t(p + "1") != "" for p in POSITIONS[:5]] + [yama4 and uf.t("F1") != ""]
    if quirks.reproduce("B10"):
        names = ["S", "U", "2", "D", "E", "F"]
        b_gauge_rows = 5
    else:
        names = list(POSITIONS)
        b_gauge_rows = 6
    msg = ""
    for i in range(6):
        if atu[i] and not fa[i]:
            msg = names[i] + "のA面 定盤 未入力"
    for i in range(6):
        if atu[i] and tw <= 35.0 and fa[i] and not fb[i]:
            msg = names[i] + "のB面 定盤 未入力_ﾊﾟﾀｰﾝA"
    for i in range(6):
        if atu[i] and tw <= 35.0 and fb[i] and not fa[i]:
            msg = names[i] + "のA面 定盤 未入力_ﾊﾟﾀｰﾝA"
    for i in range(6):
        if atu[i] and tw > 35.0 and fa[i] and fb[i]:
            return msg
    for i in range(6):
        if atu[i] and tw > 35.0 and fa[i] and not fb[i] and not fag[i]:
            msg = names[i] + "のA面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"
    for i in range(b_gauge_rows):
        if atu[i] and tw > 35.0 and fb[i] and not fa[i] and not fbg[i]:
            msg = names[i] + "のB面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"
    return msg


def is_valid_numeric_with_slash(s: str) -> bool:
    """"/" を含み、"/" を除くと数になるか。`IsValidNumericWithSlash (O:2872-2887)`"""
    if vbx.instr(s, "/") == 0:
        return False
    return vbx.is_numeric(s.replace("/", ""))


def is_valid_format(s: str) -> bool:
    """「文字_文字_数値」の形か。`IsValidFormat (O:2840-2870)`"""
    parts = s.split("_") if s != "" else []
    if len(parts) != 3:
        return False
    if (parts[0] != "" or parts[1] != "" or parts[2] != "") and (parts[0] == "" or parts[1] == "" or parts[2] == ""):
        if parts[2] != "0.00":
            return False
    n0, n1, n2 = (vbx.is_numeric(p) for p in parts)
    if n0 and n1 and not n2:
        return False
    if not n0 and not n1 and not n2:
        return is_valid_numeric_with_slash(parts[2])
    return True


def gauge_judgment(uf: Uf1, temp: str, dist: Optional[WorkDist]) -> str:
    """位置 1 か所の平坦度チェック(CheckT? の色付け用)。`GaugeJudgment (O:2746-2838)`

    WithorWithout_Gauge と違い、最初に当たった条件で終わる。文言は全角の「パターン」「ゲージ」。
    """
    tw = weight(vbx.val(uf.t("ATU")), vbx.val(uf.t("Hab")), vbx.val(uf.t("Tak")))
    if tw == 0:
        return ""
    if dist is None:
        return ""
    if temp not in POSITIONS:
        return ""
    i = POSITIONS.index(temp)
    a, b = _dist_flags(dist)
    fa = a[i][0] != "" and is_valid_format(a[i][0])
    fb = b[i][0] != "" and is_valid_format(b[i][0])
    fag = a[i][1] != "" and is_valid_format(a[i][1])
    fbg = b[i][1] != "" and is_valid_format(b[i][1])
    if temp == "F":
        atu = uf.t("F1") != "" if uf.t("Yama") == "4" else False
    else:
        atu = uf.t(temp + "1") != ""
    if atu and not fa:
        return temp + "のA面 定盤 未入力"
    if atu and tw <= 35:
        if fa and not fb:
            return temp + "のB面 定盤 未入力_パターンA"
        if fb and not fa:
            return temp + "のA面 定盤 未入力_パターンA"
    if atu and tw > 35:
        if fa and fb:
            return ""
        if fa and not fb and not fag:
            return temp + "のA面 ゲージ 未入力_パターンB"
        if fb and not fa and not fbg:
            return temp + "のB面 ゲージ 未入力_パターンB"
    return ""


def _control_kind(name: str) -> Optional[str]:
    """UF1 にあるコントロールの種類(GetLabelState が見るものだけ)。無ければ None。"""
    m = re.fullmatch(r"([SUCDEF])(\d+)", name)
    if m:
        k = int(m.group(2))
        limit = 10 if m.group(1) == "F" else 14
        return "Label" if 1 <= k <= limit else None
    if re.fullmatch(r"[SUCDEF]_(Up|Under)Visual", name):
        return "Label"
    if re.fullmatch(r"Lateral[SUCDEF]|Diag[SUCDEF]Kou", name):
        return "TextBox"
    return None


def get_label_state(uf: Uf1, name: str) -> str:
    """入力欄の状態。表示かつ空="否"、表示かつ入力あり="可"、非表示="すべて非表示"、無い="無効"。
    `GetLabelState (O:2640-2674)`
    """
    kind = _control_kind(name)
    if kind is None:
        return "無効"
    if not uf.vis(name):
        return "すべて非表示"
    return "否" if uf.t(name) == "" else "可"


def check_labels_by_category(uf: Uf1, category: str, dist: Optional[WorkDist]) -> Optional[int]:
    """CheckT<位置> の背景色(GREEN/WHITE/RED、None = 変えない)。`CheckLabelsByCategory (O:2676-2744)`

    1〜14 の測定欄 → ラテラルボー・直角度・上下外観 → GaugeJudgment の順。
    「否」「無効」で赤にして終わる(F は 11 番目=F11 が無いので、そこで番号の欄を打ち切る)。
    """
    color: Optional[int] = None
    flag = False
    if category in POSITIONS:
        for i in range(1, 15):
            state = get_label_state(uf, f"{category}{i}")
            if state == "可":
                color, flag = GREEN, True
            elif state == "すべて非表示":
                color = WHITE
            else:
                if category == "F" and i == 11:
                    break
                return RED
    else:
        return RED
    for name in ("Lateral" + category, f"Diag{category}Kou", category + "_UpVisual", category + "_UnderVisual"):
        state = get_label_state(uf, name)
        if state == "可":
            color, flag = GREEN, True
        elif state == "すべて非表示":
            color = WHITE
        else:
            return RED
    g = gauge_judgment(uf, category, dist)
    if g == "":
        if flag:
            color = GREEN
    else:
        color = RED
    return color


# ------------------------------------------------------------------
# 歪 3 項目・未入力チェック
# ------------------------------------------------------------------
@dataclass(frozen=True)
class DistortionResult:
    text: str                        # "" = 問題なし
    cleared: tuple[str, ...] = ()    # 値 0 のため "" にした形状 ComboBox(画面へ反映。CB?1_Change が動く)


def _zero_clear(uf: Uf1) -> tuple[dict[str, str], list[str]]:
    vals = dict(uf.text)
    cleared: list[str] = []
    # 定盤は ZF → CBF1 が無い(VBA のとおり。B46)
    pairs = [("ZS", "CBS1"), ("ZU", "CBU1"), ("ZC", "CBC1"), ("ZD", "CBD1"), ("ZE", "CBE1"),
             ("GS", "CBS3"), ("GU", "CBU3"), ("GC", "CBC3"), ("GD", "CBD3"), ("GE", "CBE3"), ("GF", "CBF3")]
    if not quirks.reproduce("B46"):
        pairs.insert(5, ("ZF", "CBF1"))                      # 直した側: F 位置の定盤も 0 なら形状を空にする
    for val_name, cb in pairs:
        if _num_eq(vals.get(val_name, ""), 0):
            vals[cb] = ""
            cleared.append(cb)
    return vals, cleared


def distortion_check(uf: Uf1, n: str) -> DistortionResult:
    """歪 3 項目(形状・上下・値)の入力漏れ。`Distortion_Check (O:1984-2073)`

    n="Z": 定盤(CB?1, CB?2, Z?)。n="G": flatnesText が "" ならゲージ(CB?3, CB?4, G?)、
    "ﾌﾘｰｿﾘ要" ならフリーソリ(CB?3, G?)。どれかが入っていて、どれかが空で、値が 0 でない位置。
    複数なら最後(F 側)。先に「値が 0 なら形状を空にする」(B46: ZF→CBF1 は無い。直すと足す)。
    B9: 定盤の F 行は ZE の値を見る(VBA どおり)。
    """
    vals, cleared = _zero_clear(uf)
    g = lambda k: "" if vals.get(k) is None else str(vals.get(k, ""))   # noqa: E731
    text = ""
    if n == "Z":
        tr = {p: g(f"CB{p}1") != "" or g(f"CB{p}2") != "" or g("Z" + p) != "" for p in POSITIONS}
        fa = {p: g(f"CB{p}1") == "" or g(f"CB{p}2") == "" or g("Z" + p) == "" for p in POSITIONS}
        for p in POSITIONS:
            src = "Z" + p
            if p == "F" and quirks.reproduce("B9"):
                src = "ZE"
            nz = not _num_eq(g(src), 0)
            if tr[p] and fa[p] and nz:
                text = "定盤:" + p
    elif n == "G":
        flat = uf.t("flatnesText")
        if flat == "":
            tr = {p: g(f"CB{p}3") != "" or g(f"CB{p}4") != "" or g("G" + p) != "" for p in POSITIONS}
            fa = {p: g(f"CB{p}3") == "" or g(f"CB{p}4") == "" or g("G" + p) == "" for p in POSITIONS}
            head = "ｹﾞｰｼﾞ:"
        elif flat == "ﾌﾘｰｿﾘ要":
            tr = {p: g(f"CB{p}3") != "" or g("G" + p) != "" for p in POSITIONS}
            fa = {p: g(f"CB{p}3") == "" or g("G" + p) == "" for p in POSITIONS}
            head = "ﾌﾘｰｿﾘ:"
        else:
            return DistortionResult("", tuple(cleared))
        for p in POSITIONS:
            nz = not _num_eq(g("G" + p), 0)
            if tr[p] and fa[p] and nz:
                text = head + p
    return DistortionResult(text, tuple(cleared))


def single_unit_distortion(uf: Uf1, n: str, temp: str) -> str:
    """Distortion_Check の位置別版(呼び出し元なし)。`SingleUnitDistortion (O:2889-2953)`"""
    if temp not in POSITIONS:
        value, cb1, cb2 = "", "", ""
    elif n == "Z":
        value, cb1, cb2 = uf.t("Z" + temp), uf.t(f"CB{temp}1"), uf.t(f"CB{temp}2")
    else:
        value, cb1, cb2 = uf.t("G" + temp), uf.t(f"CB{temp}3"), uf.t(f"CB{temp}4")
    if n not in ("Z", "G"):
        return ""
    tr = cb1 != "" or cb2 != "" or (value != "" and value != "0")
    fa = cb1 == "" or cb2 == "" or value == ""
    if tr and fa:
        return ("定盤:" if n == "Z" else "ｹﾞｰｼﾞ:") + temp
    return ""


def _first_row(uf: Uf1, cond: Callable[[str], bool]) -> str:
    """S→F の順に、厚 1 か所目の測定があり cond が真の最初の位置の見出し(F は Yama="4" のときだけ)。"""
    for p in POSITIONS:
        if p == "F" and uf.t("Yama") != "4":
            continue
        if uf.t(p + "1") != "" and cond(p):
            return uf.t(p)
    return ""


def angle_check(uf: Uf1) -> str:
    """直角度の未入力(表示中の欄だけ)。`Angle_Check (O:2074-2088)`"""
    return _first_row(uf, lambda p: uf.vis(f"Diag{p}Kou") and uf.t(f"Diag{p}Kou") == "")


def lateralbow_check(uf: Uf1) -> str:
    """ラテラルボーの未入力(表示中の欄だけ)。`Lateralbow_Check (O:2090-2105)`"""
    return _first_row(uf, lambda p: uf.vis("Lateral" + p) and uf.t("Lateral" + p) == "")


def visual_check(uf: Uf1) -> str:
    """外観(上・下)の未入力。`Visual_Check (O:2106-2120)`"""
    return _first_row(uf, lambda p: uf.t(p + "_UpVisual") == "" or uf.t(p + "_UnderVisual") == "")


def bh_check(uf: Uf1) -> str:
    """BH の未入力(QF 通過材。呼び出し元なし)。`BH_Check (O:2121-2157)`"""
    flag = vbx.instr(uf.t("QF"), "QF") > 0
    s, u, c, d, e = (uf.t(p + "1") for p in POSITIONS[:5])
    count = [s != "" and u == "" and c == "" and d == "" and e == "",
             u != "" and c == "" and d == "" and e == "",
             c != "" and d == "" and e == "",
             d != "" and e == "",
             e != ""]
    out = ""
    for k, last in enumerate(("S", "U", "C", "D", "E")):
        if count[k]:
            if flag and uf.t("SBH") == "":
                out = uf.t("S")
            if k > 0 and flag and uf.t(last + "BH") == "":
                out = uf.t(last)
            break
    return out


# ------------------------------------------------------------------
# 印刷前チェック
# ------------------------------------------------------------------
@dataclass(frozen=True)
class PrecheckResult:
    """印刷前の必須チェックの結果。text が "" なら印刷へ進む(VBA は MsgError(text, Lname, 1) で止める)。"""
    text: str
    source: str = ""
    cleared: tuple[str, ...] = ()


def print_precheck(uf: Uf1, dist: Optional[WorkDist]) -> PrecheckResult:
    """印刷前の必須チェック。`UF1.PrintB_Click (UF1:3659-3714)`(先行データ票の確認の後)

    7 つのチェックを全部求めてから(Distortion_Check は形状を空にする)、
    定盤 → ゲージ → 歪AB → ラテラルボー → 直角度 → 外観 → 検入高さ/枚数 → 山数 の順に最初の該当を返す。
    B2: ラテラルボーは Sumisei.Caption を「ダブルクォート 1 文字」と比べる(ほぼ常に真)。
    """
    z = distortion_check(uf, "Z")
    after_z = Uf1(dict(uf.text), dict(uf.visible), set(uf.ng), uf.work, uf.line_name, uf.line_no, uf.worker_day_max)
    for cb in z.cleared:
        after_z.text[cb] = ""
    g = distortion_check(after_z, "G")
    rate = lateralbow_check(uf)
    gaikan = visual_check(uf)
    ken_h = h_check([uf.t(f"KenHBox{i}") for i in range(1, 8)])
    ab = withor_without_gauge(uf, dist)
    chokkaku = angle_check(uf)
    cleared = tuple(dict.fromkeys(z.cleared + g.cleared))

    if quirks.reproduce("B2"):
        sumisei_ok = uf.t("Sumisei") != '"'
    else:
        sumisei_ok = uf.t("Sumisei") != ""
    if z.text != "":
        return PrecheckResult(z.text + "入力なし", "UF1:3674", cleared)
    if g.text != "":
        return PrecheckResult(g.text + "入力なし", "UF1:3679", cleared)
    if ab != "":
        return PrecheckResult(ab, "UF1:3684", cleared)
    if rate != "" and (sumisei_ok and vbx.instr(uf.t("Additional"), "ﾗﾃﾗﾙﾎﾞｰ検査不要") == 0):
        return PrecheckResult(rate + ":ﾗﾃﾗﾙﾎﾞｰ入力なし", "UF1:3689", cleared)
    if chokkaku != "" and uf.t("Youto") == "R115":
        return PrecheckResult(chokkaku + ":直角度入力なし", "UF1:3694", cleared)
    if gaikan != "":
        return PrecheckResult(gaikan + ":外観入力なし", "UF1:3699", cleared)
    if ken_h == "入力なし":
        return PrecheckResult("検入高さor枚数　入力なし", "UF1:3704", cleared)
    if uf.vis("Yama") and uf.t("Yama") == "":
        return PrecheckResult("山数の入力がありません", "UF1:3709", cleared)
    return PrecheckResult("", "", cleared)


def allow_check(uf: Uf1) -> str:
    """許容・保留の位置の一覧(成績票 B10 用)。`allow_Check (O:2598-2638)`

    NGS〜NGF の Caption に「許容」(先)か「保留」を含む位置。2 件目からは ", 位置 許容" の形。
    B34: 1 件目には「許容」「保留」の文字を付けない(そのため 1 か所だけだと B10 が出ない)。
    最後に NG? を位置の見出し(UF1.S.Caption など)に置き換える。
    """
    result = ""
    for name in ("NGS", "NGU", "NGC", "NGD", "NGE", "NGF"):
        cap = uf.t(name)
        if vbx.instr(cap, "許容") > 0:
            word = "許容"
        elif vbx.instr(cap, "保留") > 0:
            word = "保留"
        else:
            continue
        if result == "":
            result = name if quirks.reproduce("B34") else name + " " + word
        else:
            result = result + ", " + name + " " + word
    for p in POSITIONS:
        result = result.replace("NG" + p, uf.t(p))
    return result


# ------------------------------------------------------------------
# 導電率(QA-021-08)
# ------------------------------------------------------------------
def conductiv_check(uf: Uf1) -> bool:
    """導電率の測定が要る材質・調質か(75S: T651/T6、14S: T6/T651)。`Conductiv_Check (O:2172-2183)`"""
    zai, tyu = uf.t("Zai"), uf.t("Tyu")
    if zai == "75S":
        return tyu in ("T651", "T6")
    if zai == "14S":
        return tyu in ("T6", "T651")
    return False


@dataclass(frozen=True)
class CoolingResult:
    value: str                                   # 写真確認No / "" / "LotNo記載がない"
    messages: tuple[Message, ...] = ()


def coolingtime_lookup(rows: Sequence[Sequence[object]], lot: str) -> str:
    """QF空冷時間チェックリスト Sheet1 の 5 行目から(A=写真確認No、B=LotNo)。
    一致した**最後の行**の A を返す(空なら "")。無ければ "LotNo記載がない"。
    B12: 呼び出し側(SaveLoad:409)は "LotNoがない" と比べている。直す場合はこちらを "LotNoがない" にする。
    `Conductiv_Coolingtime_Check (O:2226-2249)`
    """
    found = False
    value = ""
    for r in rows:
        r = list(r) + [None, None]
        if _cell_text(r[1]) == lot:
            found = True
            a = _cell_text(r[0])
            value = a if a != "" else ""
    if not found:
        return "LotNo記載がない" if quirks.reproduce("B12") else "LotNoがない"
    return value


def coolingtime_path() -> str:
    """`QualityPath & "75S・14S焼入れチェック\\QF空冷時間チェックリスト.xlsx"`"""
    return paths.join(paths.get("quality"), "75S・14S焼入れチェック", "QF空冷時間チェックリスト.xlsx")


_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
       "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
       "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}


def _col_no(ref: str) -> int:
    n = 0
    for ch in re.match(r"[A-Z]+", ref).group(0):
        n = n * 26 + ord(ch) - 64
    return n


def _si_text(si) -> str:
    """文字列 1 つの本文。ふりがな(rPh)は読まない(xlsm_import.si_text と同じ)。"""
    from .xlsm_import import si_text
    return si_text(si)


def read_xlsx_sheet(path: Path, sheet_name: str) -> Optional[dict[tuple[int, int], object]]:
    """xlsx の 1 シートを {(行, 列): 値} で読む(標準ライブラリだけ)。シートが無ければ None。

    数は float、文字は str、真偽は bool。日付も数(シリアル値)のまま。
    """
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rid = None
        for s in wb.iter(f"{{{_NS['m']}}}sheet"):
            if s.get("name") == sheet_name:
                rid = s.get(f"{{{_NS['r']}}}id")
                break
        if rid is None:
            return None
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = None
        for rel in rels.iter(f"{{{_NS['rel']}}}Relationship"):
            if rel.get("Id") == rid:
                target = rel.get("Target")
        if target is None:
            return None
        target = target.lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        shared: list[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            ss = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in ss.iter(f"{{{_NS['m']}}}si"):
                shared.append(_si_text(si))
        sh = ET.fromstring(z.read(target))
        cells: dict[tuple[int, int], object] = {}
        for c in sh.iter(f"{{{_NS['m']}}}c"):
            ref = c.get("r") or ""
            m = re.match(r"([A-Z]+)(\d+)", ref)
            if not m:
                continue
            key = (int(m.group(2)), _col_no(m.group(1)))
            t = c.get("t", "n")
            v = c.find(f"{{{_NS['m']}}}v")
            if t == "inlineStr":
                is_ = c.find(f"{{{_NS['m']}}}is")
                cells[key] = _si_text(is_) if is_ is not None else ""
            elif v is None or v.text is None:
                continue
            elif t == "s":
                cells[key] = shared[int(v.text)]
            elif t in ("str", "e"):
                cells[key] = v.text
            elif t == "b":
                cells[key] = v.text == "1"
            else:
                cells[key] = float(v.text)
        return cells


def conductiv_coolingtime_check(lot: str, path: Optional[Path] = None) -> CoolingResult:
    """QF空冷時間チェックリスト.xlsx の写真確認No。`Conductiv_Coolingtime_Check (O:2184-2297)`

    ファイルが無い・Sheet1 が無いときは VBA の MsgBox(vbCritical)の文言を付けて "" を返す。
    範囲は 5 行目〜B 列の最後の行、A 列〜4 行目の最後の列(B 列が無いと添字エラー 9)。
    """
    text_path = coolingtime_path()
    p = path or paths.to_local(text_path)
    if not p.is_file():
        return CoolingResult("", (Message(text_path + "がありません" + "_Conductiv_Coolingtime_Check", ERROR,
                                          source="OutTolerance:2193"),))
    cells = read_xlsx_sheet(p, "Sheet1")
    if cells is None:
        return CoolingResult("", (Message("指定シートがありません" + "_Conductiv_Coolingtime_Check", ERROR,
                                          source="OutTolerance:2251"),))
    nonempty = lambda v: v is not None and _cell_text(v) != ""   # noqa: E731
    end_r = max([r for (r, c), v in cells.items() if c == 2 and nonempty(v)], default=1)
    end_c = max([c for (r, c), v in cells.items() if r == 4 and nonempty(v)], default=1)
    if end_c < 2:
        raise VbaRuntimeError(9, "インデックスが有効範囲にありません(4 行目が B 列まで無い)")
    r1, r2 = sorted((5, end_r))
    rows = [[cells.get((r, 1)), cells.get((r, 2))] for r in range(r1, r2 + 1)]
    return CoolingResult(coolingtime_lookup(rows, lot))


def conductiv_sheet_values(now_text: str, uf: Uf1, worker_cell: str) -> dict[str, str]:
    """導電率測定結果記入表(シート「導電率測定」)の B4〜B9。`ConductivLoad (O:2334-2349)`

    now_text = Format(Now, "yyyy/mm/dd")。worker_cell = 検査員_<Lname> の B1/B2/B3(WorkeDayMax の行)。
    """
    parts = worker_cell.split("_") if worker_cell != "" else []
    if not parts:
        raise VbaRuntimeError(9, "インデックスが有効範囲にありません(検査員が空)")
    return {
        "B4": now_text,
        "B5": uf.t("Lot"),
        "B6": uf.t("Zai") + "-" + uf.t("Tyu"),
        "B7": uf.t("ATU") + "×" + uf.t("Hab") + "×" + uf.t("Tak"),
        "B8": uf.lname,
        "B9": parts[-1],
    }


# ------------------------------------------------------------------
# 不良要因の枚数
# ------------------------------------------------------------------
@dataclass(frozen=True)
class SelectionWrite:
    """不良要因印刷_<Lname> に書くセル(行, 列, 値)。"""
    cells: tuple[tuple[int, int, str], ...]


def selection_input(selected_value: str, te_name: str, te_value: str, code_row: Sequence[object],
                    uf: Optional[Uf1] = None, lot: str = "") -> Optional[SelectionWrite]:
    """不良要因の枚数を不良要因印刷シートへ。`SelectionInput (O:2356-2449)`

    selected_value = GetSelectedItems(DefectFactorLB)("" なら何もしない)。code_row = シートの C9:X9。
    uf を渡すと UF1 から(B1=ロット、B3〜B8=位置の見出し)、None なら三菱(B1 だけ)。
    「ｽﾎﾟｯﾄ処理」は C11 に「枚数 & 枚」。TP は 22 列目(X 列)固定。それ以外はコード(最後の "_" の後)が一致する列。
    コードが見つからない・欄名が S〜F でないときは、VBA は b_Arr の添字エラー(9)で止まる
    (`itti = 0` のやり直し処理には届かない)。
    B45: 複数選択(「汚_12 TP」など。一覧は複数選択できる)だと最後の "_" の後("12 TP")で探すので
    添字エラー(9)、または最後の項目の列だけに書く。直すと、空白で区切った項目ごとに同じ枚数を書く
    (TP は 22 列目、ｽﾎﾟｯﾄ処理 は C11)。
    """
    if selected_value == "":
        return None
    cells: list[tuple[int, int, str]] = [(1, 2, uf.t("Lot") if uf is not None else lot)]
    if uf is not None:
        for k, p in enumerate(POSITIONS):
            cells.append((3 + k, 2, "'" + uf.t(p)))
    if selected_value == SPOT:
        cells.append((11, 3, te_value + "枚"))
        return SelectionWrite(tuple(cells))
    a = {"S": 1, "U": 2, "C": 3, "D": 4, "E": 5, "F": 6}.get(vbx.left(te_name, 1), 0)
    codes = [_cell_text(v) for v in code_row]
    if not quirks.reproduce("B45"):
        # 直した側: GetSelectedItems は項目を " " でつなぐので、項目ごとに書く
        for piece in selected_value.split(" "):
            if piece == "":
                continue
            if piece == SPOT:
                cells.append((11, 3, te_value + "枚"))
                continue
            if piece == "TP":
                col = 22
            else:
                code = piece.split("_")[-1]
                col = codes.index(code) + 1 if code in codes else 0
            if a == 0 or col == 0:
                raise VbaRuntimeError(9, "インデックスが有効範囲にありません(不良要因の列・行)")
            cells.append((2 + a, 2 + col, te_value))
        return SelectionWrite(tuple(cells))
    parts = selected_value.split("_")
    code = parts[-1] if parts else ""
    itti = 0
    for i, v in enumerate(codes, start=1):
        if v == code:
            itti = i
            break
    col = 22 if selected_value == "TP" else itti
    if a == 0 or col == 0:
        raise VbaRuntimeError(9, "インデックスが有効範囲にありません(不良要因の列・行)")
    cells.append((2 + a, 2 + col, te_value))
    return SelectionWrite(tuple(cells))


# ------------------------------------------------------------------
# そのほか
# ------------------------------------------------------------------
def find_right_string(s: str) -> str:
    """コース文字列から最終熱処理(ANI/ANF/CAI/CAF)を右から探す。`FindRightString (O:2495-2507)`"""
    for w in reversed(s.split(" ")):
        if w in ("ANI", "ANF", "CAI", "CAF"):
            return w
    return ""


def uftag_array(tags: dict[str, str], x: int, y: int) -> dict[int, tuple[str, str]]:
    """歪の Tag を「形状_上下_値」に(位置 x〜y、(定盤, ゲージ))。`UFtag_Array (O:2566-2597)`"""
    t = lambda k: "" if tags.get(k) is None else str(tags.get(k, ""))   # noqa: E731
    out: dict[int, tuple[str, str]] = {}
    for n in range(x, y + 1):
        if 1 <= n <= 6:
            p = POSITIONS[n - 1]
            out[n] = (t(f"CB{p}1") + "_" + t(f"CB{p}2") + "_" + t("Z" + p),
                      t(f"CB{p}3") + "_" + t(f"CB{p}4") + "_" + t("G" + p))
        else:
            out[n] = ("", "")
    return out


def pack_picture_path(n: str) -> str:
    """包装仕様書の画像。`UFPack_Open (O:524-525)`: `QualityPath & "包装仕様書\\" & n & ".JPG"`"""
    return paths.join(paths.get("quality"), "包装仕様書", n + ".JPG")


@dataclass(frozen=True)
class MsgErrorOut:
    message: Message
    log_text: str
    log_folder: str
    end: bool                 # Temp="END"(UFMsgBox の OK で VBA は End。D10: 止めるが状態は保つ)


def msg_error(trm: str, lname: str, n: int = 0, temp: str = "", t123: tuple[str, str, str] = ("", "", "")) -> MsgErrorOut:
    """警告・注意のボックス。`MsgError (O:716-794)`

    カンマで改行(空白だけの行は出さない)。カンマが無ければそのまま。
    タイトルは n=4 のとき「注意」、それ以外は「警告」。アイコン 1=!、2=×、3=?、4=i。
    ログ: "*UFKen:MsgError:" & Trm & " T1:… T2:… T3:…"(ライン別フォルダ)。表示中は UF1 を隠す。
    """
    if "," not in trm:
        text = trm
    else:
        text = "\n".join(p for p in trm.split(",") if p.strip(" ") != "")
    level = {1: WARN, 2: ERROR, 3: INFO, 4: INFO}.get(n, WARN)
    msg = Message(text, level, "注意" if n == 4 else "警告", source="OutTolerance.MsgError")
    log = "*UFKen:MsgError:" + trm + " T1:" + t123[0] + " T2:" + t123[1] + " T3:" + t123[2]
    return MsgErrorOut(msg, log, log_folder(lname), temp == "END")


# ------------------------------------------------------------------
# 少枚数の検査位置(SmallNumberCount)
# ------------------------------------------------------------------
@dataclass
class ScreenChanges:
    """画面の変更(後から書いたものが残る)。書いていないコントロールは変えない。"""
    visible: dict[str, bool] = field(default_factory=dict)
    enabled: dict[str, bool] = field(default_factory=dict)
    values: dict[str, str] = field(default_factory=dict)
    lateral: bool = False            # ラテラルボー欄を出した(ラテflag かつ 検査不要でない)
    initial_setting: bool = True     # 最後に FormRelat.InitialSetting(UF1) を呼ぶ

    def vis(self, names: Iterable[str], v: bool, enabled: Optional[bool] = None) -> None:
        for n in names:
            self.visible[n] = v
            if enabled is not None:
                self.enabled[n] = enabled

    def en(self, names: Iterable[str], v: bool) -> None:
        for n in names:
            self.enabled[n] = v

    def set(self, **kw: str) -> None:
        self.values.update(kw)


DbReturnItem = Callable[[str, str, str, str, str, float, float, float], str]


def small_number_count(tb: int, uf: Uf1, lot_rows: Optional[Sequence[Sequence[object]]],
                       db_return_item: DbReturnItem) -> ScreenChanges:
    """少枚数の検査回数(1〜5)で入力欄を出し分ける。`SmallNumberCount (O:3-519)`(Q-S-CH-DIM-163, 2025.7)

    lot_rows = AccHIki.GetLotInfo("Lot情報履歴", Lot)(None = Empty)。各行の列 1=納入, 2=取引, 3=用途コード,
    4=材質, 5=調質, 6=厚, 7=幅, 8=丈, 10=検番。db_return_item = QP_DB.DbReturnItem。
    ラテラルボーは DbReturnItem の結果に「ラテラルボー」があり、Additional の右 4 文字が「検査不要」でないとき。
    納入先が前の行と同じなら DbReturnItem を呼ばない。用途コード〜丈は**最後の行の値**で全行を照合する(VBA のとおり)。
    位置(Left/Top)・幅は移さない。
    B47 を直すと: 照合はその行の用途コード〜丈で(前と同じ照合値なら呼ばない)、ラテラル不要のロットでは
    ラテラル欄(L11・Lateral?)を隠し、5 回では隠した DiagU/D の直角度欄(DiagUKou/DKou)も隠す。
    """
    ng_temp = vbx.right(uf.t("Additional"), 4)
    ch = ScreenChanges()
    rows5 = POSITIONS[:5]
    for i in range(11, 15):
        ch.vis([f"{p}{i}" for p in rows5], False, False)
    ch.set(S="1", U="2", C="3", D="4", E="5", FS="1", FU="2", FC="3", FD="4", FE="5",
           OPS1="1", OPS2="2", OPS3="3", OPS4="4", OPS5="5",
           OPE1="1", OPE2="2", OPE3="3", OPE4="4", OPE5="5")
    ch.vis(["FM" + p for p in rows5], True)
    ch.vis([f"LC{i}" for i in (1, 2, 3, 4, 5, 7, 8, 9, 10, 11)], True)
    ch.vis([p + "_UpVisual" for p in rows5] + [p + "_UnderVisual" for p in rows5], True)
    for i in range(1, 11):
        ch.vis([f"{p}{i}" for p in rows5], True, True)
    ch.vis(["Diag" + p for p in rows5], True, True)
    ch.vis([p + "BH" for p in rows5], True)
    ch.vis(["TS" + p for p in rows5] + ["TE" + p for p in rows5], True)

    late = False
    if lot_rows is not None:
        rows = [list(r) + [""] * 10 for r in lot_rows]
        nounyu = [_cell_text(r[0]) for r in rows]
        torihiki = [_cell_text(r[1]) for r in rows]
        youto = zai = tyo = ""
        atu = haba = take = 0.0
        for r in rows:
            youto, zai, tyo = _cell_text(r[2]), _cell_text(r[3]), _cell_text(r[4])
            atu, haba, take = (vbx.single(vbx.cdbl(r[k])) for k in (5, 6, 7))
        if quirks.reproduce("B47"):
            prev = ""
            for nou, tori in zip(nounyu, torihiki):
                if nou != "":
                    if nou == prev:
                        continue
                    prev = nou
                    temp = db_return_item(nou, tori, youto, zai, tyo, atu, haba, take)
                    if vbx.instr(temp, "ラテラルボー") > 0:
                        late = True
        else:
            # 直した側: 行ごとの値で照合する(同じ照合値が続くときだけ省く)
            prev_key: Optional[tuple] = None
            for r in rows:
                key = (_cell_text(r[0]), _cell_text(r[1]), _cell_text(r[2]), _cell_text(r[3]), _cell_text(r[4]),
                       *(vbx.single(vbx.cdbl(r[k])) for k in (5, 6, 7)))
                if key[0] != "":
                    if key == prev_key:
                        continue
                    prev_key = key
                    temp = db_return_item(*key)
                    if vbx.instr(temp, "ラテラルボー") > 0:
                        late = True
    show_late = late and ng_temp != "検査不要"
    ch.lateral = show_late
    if show_late:
        ch.vis(["L11"] + ["Lateral" + p for p in rows5], True)
    elif not quirks.reproduce("B47"):
        ch.vis(["L11"] + ["Lateral" + p for p in rows5], False)     # 前ロットの表示を残さない
    r115 = uf.t("Youto") == "R115"
    kou = [f"Diag{p}Kou" for p in rows5]
    if r115:
        ch.vis(["L10"] + kou, True)

    if tb == 5:
        if show_late:
            ch.visible.update(LateralS=True, LateralU=False, LateralC=True, LateralD=False, LateralE=True)
        for i in range(5, 11):
            ch.vis([f"U{i}", f"D{i}"], False, False)
        if r115:
            ch.vis(["L10"] + kou, True)
        ch.vis(["DiagU", "DiagD"], False, False)
        if r115 and not quirks.reproduce("B47"):
            ch.vis(["DiagUKou", "DiagDKou"], False, False)           # 隠した対角の直角度欄も隠す
    elif tb == 4:
        ch.set(E="", FE="", OPS5="", OPE5="")
        ch.vis(["FME", "LC5", "LC11", "E_UpVisual", "E_UnderVisual"], False)
        ch.vis([f"E{i}" for i in range(1, 11)], False, False)
        ch.vis([f"U{i}" for i in range(5, 11)], False, False)
        ch.vis(["DiagU", "DiagE"], False, False)
        if r115:
            ch.vis(["DiagEKou"], False, False)
        ch.vis(["EBH", "TSE", "TEE"], False)
        if show_late:
            ch.vis(["LateralU", "LateralE"], False)
    elif tb == 3:
        ch.set(D="", E="", FD="", FE="", OPS4="", OPS5="", OPE4="", OPE5="")
        ch.vis(["FMD", "FME", "LC5", "LC11", "LC4", "LC10",
                "D_UpVisual", "E_UpVisual", "D_UnderVisual", "E_UnderVisual"], False)
        ch.vis([f"{p}{i}" for i in range(1, 11) for p in ("D", "E")], False, False)
        ch.vis(["DiagD"], False, False)
        if r115:
            ch.vis(["DiagDKou", "DiagEKou"], False, False)
        ch.vis(["DiagE"], False, False)
        ch.vis(["DBH", "EBH", "TSD", "TSE", "TED", "TEE"], False)
        if show_late:
            ch.vis(["LateralD", "LateralE"], False)
    elif tb == 2:
        ch.set(C="", D="", E="", FC="", FD="", FE="", OPS3="", OPS4="", OPS5="", OPE3="", OPE4="", OPE5="")
        ch.vis(["FMC", "FMD", "FME", "LC5", "LC11", "LC4", "LC10", "LC3", "LC9"]
               + [p + "_UpVisual" for p in "CDE"] + [p + "_UnderVisual" for p in "CDE"], False)
        ch.vis([f"{p}{i}" for i in range(1, 11) for p in "CDE"], False, False)
        ch.vis(["DiagC"], False, False)
        if r115:
            ch.vis(["DiagCKou", "DiagDKou", "DiagEKou"], False, False)
        ch.vis(["DiagD", "DiagE"], False, False)
        ch.vis(["CBH", "DBH", "EBH", "TSC", "TSD", "TSE", "TEC", "TED", "TEE"], False)
        if show_late:
            ch.vis(["LateralC", "LateralD", "LateralE"], False)
    elif tb == 1:
        ch.set(U="", C="", D="", E="", FU="", FC="", FD="", FE="",
               OPS2="", OPS3="", OPS4="", OPS5="", OPE2="", OPE3="", OPE4="", OPE5="")
        ch.vis(["FMU", "FMC", "FMD", "FME", "LC5", "LC11", "LC4", "LC10", "LC3", "LC9", "LC2", "LC8"]
               + [p + "_UpVisual" for p in "UCDE"] + [p + "_UnderVisual" for p in "UCDE"], False)
        ch.vis([f"{p}{i}" for i in range(1, 11) for p in "UCDE"], False, False)
        if r115:
            ch.vis(["DiagUKou", "DiagCKou", "DiagDKou", "DiagEKou"], False, False)
        ch.vis(["DiagU", "DiagC", "DiagD", "DiagE"], False, False)
        ch.vis(["CBH", "DBH", "EBH", "UBH", "TSU", "TSC", "TSD", "TSE", "TEU", "TEC", "TED", "TEE"], False)
        if show_late:
            ch.vis(["LateralU", "LateralC", "LateralD", "LateralE"], False)
    return ch
