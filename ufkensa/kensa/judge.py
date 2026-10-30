"""寸法判定(VBA `【UFKen_04】Judg`)

厚・巾・丈、対角線差、直角度、ラテラルボー、BH、歪の判定と、公差表(品質基準)の参照、
0.5 刻み変換、先行データ票、包装仕様の注意文を移植したもの。
根拠は docs/analysis/G2_Judg_Tolerance.md(§0、§3.0〜§3.8、§3.15、§3.16、§4)。行番号の `J:` は Judg。

【画面の扱い】
VBA は UF1 のラベル・テキストボックスの Caption/Text/BackColor/Tag を直接読み書きしていた。
ここでは画面の値を `Form`(コントロール名 → `Cell`)で受け取り、変えるもの(判定状態・表示値・Tag・
音・画像・ログ・メッセージ)を `Outcome` で返す。画面は `apply_outcome()` で反映する。

判定状態(`Cell.state`)と VBA の BackColor の対応:
    "OK" = 13434828(薄い緑) / "NG" = 8421631(コーラル) / "" = 白(16777215) / "黄" = RGB(255,255,0)

【他の担当に任せているもの(値を引数で受け取る)】
- BhSearch(品質基準 Q_S_KE_ALL_053 のバーコール下限値。Quality_DB)→ `bh_lower`
- GetWorkerName(Worker)→ `worker_name`
- Arrbox / ReadCheck / TimeDesignation / FileDesignation(AccHIki・SaveLoad)→ 呼び出し可能な引数
- MsgError(OutTolerance)→ `Outcome.msg_errors` に積む。表示・ログ・UF1.Hide は MsgError 側
- WriteLog(WriteLog)→ `Outcome.logs` に積む。書き込みは別担当
- NGErase(OutTolerance)→ `Outcome.ng_erase` にセル名を入れる

【移植しないもの】
Judg_ImageSize(画像の位置・大きさ。`Outcome.images` に "Y"/"N"/"Neko" を積むだけ)、
Test_Tolerance(動作確認用 MsgBox。D9)、NflatRewrite(Posting を呼ぶだけで呼び出し元なし)。
"""
from __future__ import annotations

import copy
import threading
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Optional, Sequence, Union

from . import datasource, flow, paths, quirks, vbx

# ------------------------------------------------------------------
# 定数
# ------------------------------------------------------------------
OK, NG, WHITE, YELLOW = "OK", "NG", "", "黄"
RED = "赤"                                   # 枠線(LookingYellow の赤枠)。"" = 黒

POSITIONS = ("S", "U", "C", "D", "E", "F")   # 検査位置(LoopJudg の a = 1..6)
POS5 = POSITIONS[:5]                         # S〜E は 1〜14 行、F は 1〜10 行

SOUND_OK = "chime2.wav"                      # 音楽合格(補助:305)
SOUND_NG = "tada.wav"                        # 音楽不合格(補助:357)
SOUND_DIAG = "chimes.wav"                    # 音楽対角(補助:282)
SOUND_NEED = "必要です.wav"                   # MaterialsMsgBox(J:2072)
SOUND_RIGHT_ANGLE = "直角度.wav"              # MaterialsMsgBox(J:2105)

IMAGE_OK, IMAGE_NG, IMAGE_NEKO = "Y", "N", "Neko"   # Judg_ImageSize の引数

MEASURED_LOG = "Measured.log"
RIGHT_ANGLE_TOL = 1.0                        # 直角度の公差 1.0 固定(Q-S-QA-DIM-154、J:1902-1904)

TBL_DIAG = "製品規格直角度"
TBL_LAT = "製品規格ラテラルボー"
DIAG_COLS = ("管理番号", "等級", "巾_下限", "巾_上限", "丈_下限", "丈_上限", "公差")
LAT_COLS = ("管理番号", "等級", "厚_下限", "厚_上限", "巾_下限", "巾_上限", "丈_下限", "丈_上限", "公差")

PACKING_MASTER_FILE = "梱包資材マスタ.sqlite3"   # VBA: PackingPath & "梱包資材マスタ.accdb"
VB_OBJECT_ERROR = -2147221504                    # vbObjectError

BH_MSG_FIRST = "作業中止・保留・ロット異常連絡票発行"                       # J:471
BH_MSG_LAST = "梱包完了させてAインプット・保留・ロット異常連絡票発行"         # J:472
LATERAL_NO_TOL_MSG = "ﾗﾃﾗﾙﾎﾞｰ公差がありません。(規定寸法範囲外)"           # J:1808
ADVANCE_MSG = "先行データ票が必要です。"                                    # J:2065


class VbaError(RuntimeError):
    """VBA の実行時エラー(9 = インデックスが有効範囲にありません、5 = プロシージャの呼び出しが不正)。"""

    def __init__(self, number: int, text: str) -> None:
        super().__init__(f"実行時エラー {number}: {text}")
        self.number = number


class LoopJudgAborted(RuntimeError):
    """LoopJudg が途中で実行時エラーになった(VBA はそこまで塗った色が画面に残る)。

    `outcome` はエラーの直前までの変更。`error` は元の例外(vbx.TypeMismatch)。
    """

    def __init__(self, outcome: "Outcome", error: Exception) -> None:
        super().__init__(str(error))
        self.outcome = outcome
        self.error = error


# ------------------------------------------------------------------
# 画面の値(UF1)と、返す変更
# ------------------------------------------------------------------
@dataclass
class Cell:
    """UF1 のコントロール1つ(ラベル・テキストボックス)。"""
    value: str = ""          # Caption / Text
    state: str = WHITE       # 判定状態(BackColor)
    visible: bool = True
    tag: str = ""
    border: str = ""         # "" = 黒 / "赤"


def _blank_cells() -> dict[str, Cell]:
    c: dict[str, Cell] = {}
    for p in POS5:
        for i in range(1, 15):
            c[f"{p}{i}"] = Cell()
    for i in range(1, 11):
        c[f"F{i}"] = Cell()
    for p in POSITIONS:
        for pre in ("Diag", "Lateral", "Z", "G"):
            c[f"{pre}{p}"] = Cell()
        c[f"Diag{p}Kou"] = Cell()
    for p in POS5:
        c[f"{p}BH"] = Cell()
    return c


def _blank_combos() -> dict[str, str]:
    return {f"CB{p}{n}": "" for p in POSITIONS for n in range(1, 5)}


@dataclass
class Form:
    """UF1 の判定に使う値。属性名は VBA のコントロール名(G1 §3.1〜3.3、G2 §3.0)。

    cells の名前: 測定ラベル S1〜S14・U・C・D・E(1〜14)・F1〜F10(行と項目は G2 §3.0)、
    対角線差 DiagS〜DiagF、直角度 DiagSKou〜DiagFKou、ラテラルボー LateralS〜LateralF、
    BH SBH〜EBH、歪 ZS〜ZF(定盤)・GS〜GF(ゲージ)。
    captions: 位置見出し "S"〜"F" の Caption、歪の "LZS"・"LGS"。
    """
    cells: dict[str, Cell] = field(default_factory=_blank_cells)
    captions: dict[str, str] = field(default_factory=lambda: {**{p: p for p in POSITIONS}, "LZS": "", "LGS": ""})
    combos: dict[str, str] = field(default_factory=_blank_combos)   # CBS1〜CBF4(歪の形状)
    # 公差・上下限(SaveLoad:220-266 で設定。Format 済みの文字列)
    ApK: str = ""
    AmK: str = ""
    HpK: str = ""
    HmK: str = ""
    TpK: str = ""
    TmK: str = ""
    Am: str = ""
    Hm: str = ""
    Tm: str = ""
    TaikText: str = ""        # 対角線差の公差
    LateralText: str = ""     # ラテラルボーの公差
    gokuatsu: bool = False    # UF1.極厚
    # 歪の公差(FormRelat で設定)
    HaHiz: str = ""
    TaHiz: str = ""
    Naka: str = ""
    Mimi: str = ""
    HURIS: str = ""
    flatnesText: str = ""
    MoveFlag: str = ""
    # その他
    Zai: str = ""
    Tyu: str = ""
    ATU: str = ""
    Hab: str = ""
    Tak: str = ""
    Lot: str = ""
    DBflag: str = ""
    LineName: str = ""
    LineNo: str = ""
    AbnoTag: str = ""         # UF1.Abno.Tag
    kosei_visible: bool = False    # UF校正.Visible
    msgbox_visible: bool = False   # UFMsgBox.Visible

    def v(self, name: str) -> str:
        return self.cells[name].value

    @property
    def lname(self) -> str:
        """`UF1.LineName.Caption & UF1.LineNo.Caption`"""
        return self.LineName + self.LineNo


@dataclass
class LogLine:
    """`WriteLog(text, sepl, target_path, target_file)` の引数。target_path "" は既定(Access.log)。"""
    text: str
    sepl: str = ""
    target_path: str = ""
    target_file: str = ""


@dataclass
class MsgErrorCall:
    """`MsgError(Trm, Lname, n, Temp)` の引数(表示・ログ・UF1.Hide は MsgError の担当)。"""
    text: str
    lname: str
    icon: int = 0
    temp: str = ""


@dataclass
class Outcome:
    """判定が画面へ返す変更。dict は代入の順に最後の値が残る(VBA で上書きした結果と同じ)。"""
    result: Any = None
    states: dict[str, str] = field(default_factory=dict)
    values: dict[str, str] = field(default_factory=dict)
    tags: dict[str, str] = field(default_factory=dict)
    borders: dict[str, str] = field(default_factory=dict)
    sounds: list[str] = field(default_factory=list)
    images: list[str] = field(default_factory=list)
    logs: list[LogLine] = field(default_factory=list)
    messages: list[flow.Message] = field(default_factory=list)
    msg_errors: list[MsgErrorCall] = field(default_factory=list)
    move_flag: Optional[str] = None     # 新しい UF1.MoveFlag(None は変えない)
    mpex_tag: Optional[str] = None      # 新しい UF1.MPEX.Tag
    ng_erase: Optional[str] = None      # NGErase(TB) を呼ぶセル名

    def merge(self, other: "Outcome") -> None:
        if other.result is not None:
            self.result = other.result
        self.states.update(other.states)
        self.values.update(other.values)
        self.tags.update(other.tags)
        self.borders.update(other.borders)
        self.sounds += other.sounds
        self.images += other.images
        self.logs += other.logs
        self.messages += other.messages
        self.msg_errors += other.msg_errors
        if other.move_flag is not None:
            self.move_flag = other.move_flag
        if other.mpex_tag is not None:
            self.mpex_tag = other.mpex_tag
        if other.ng_erase is not None:
            self.ng_erase = other.ng_erase


def apply_outcome(form: Form, out: Outcome) -> Form:
    """`Outcome` を画面の値に反映する(form をその場で変えて返す)。"""
    for k, s in out.states.items():
        form.cells[k].state = s
    for k, s in out.values.items():
        form.cells[k].value = s
    for k, s in out.tags.items():
        form.cells[k].tag = s
    for k, s in out.borders.items():
        form.cells[k].border = s
    if out.move_flag is not None:
        form.MoveFlag = out.move_flag
    return form


# ------------------------------------------------------------------
# 小さな道具
# ------------------------------------------------------------------
def csng(x: Any) -> float:
    """VBA `CSng`(文字列 → Single)。数にならなければ型エラー 13。"""
    return vbx.single(vbx.cdbl(x))


def _resume(fn: Callable[[], float], keep: float) -> float:
    """`On Error Resume Next` の下での代入: エラーなら元の値のまま。"""
    try:
        return fn()
    except vbx.TypeMismatch:
        return keep


def _last_index(fields: Sequence[str], name: str) -> int:
    """`For i = … : Select Case F(1, i): Case "列名": C = i` (同名が複数あれば最後)。無ければ -1(VBA の 0)。"""
    idx = -1
    for i, f in enumerate(fields):
        if f == name:
            idx = i
    return idx


def _at(row: Sequence[str], col: int) -> str:
    """`arr(r, C)`。列が見つからず C = 0 のときは VBA の実行時エラー 9(配列は 1 始まり)。"""
    if col < 0:
        raise VbaError(9, "インデックスが有効範囲にありません")
    return row[col]


def vba_trim(s: str) -> str:
    """VBA `Trim`(半角空白だけを取る)。"""
    return s.strip(" ")


def log_dir(lname: str) -> str:
    """ライン → ログの置き場所(J:449-457 ほか。`HistoryPath & "L-1\\L-1_Log"`)。

    `Select Case Lname` の完全一致(LineName & LineNo)。それ以外は丸徳。
    """
    sub = {
        "L-1": "L-1\\L-1_Log",
        "LVC": "LVC\\LVC_Log",
        "HVC": "HVC\\HVC_Log",
        "ﾊﾞﾗﾝｻｰ": "バランサー\\バランサー_Log",
        "ﾄｯﾄ": "トット\\トット_Log",
        "Test": "Test\\Test_Log",
    }.get(lname, "丸徳\\丸徳_Log")
    return paths.join(paths.get("history"), sub)


def _judge_logs(text: str, lname: str) -> list[LogLine]:
    """判定ログ: Measured.log と既定の出力先に同じ文を2回(J:539-540 ほか)。"""
    return [LogLine(text, "", log_dir(lname), MEASURED_LOG), LogLine(text, "", "", "")]


def _ufken_text(value: str, name: str, mark: str, lname: str, worker: str) -> str:
    return f"*UFKen|{value}|({name})|{mark}|{lname}|{worker}"


# ------------------------------------------------------------------
# 判定の芯
# ------------------------------------------------------------------
def aht_judg(caption: str, kp: float, km: float) -> bool:
    """厚・巾・丈の判定 `Judg.AHTJudg` (J:362-369)。下限 ≤ 値 ≤ 上限で OK(両端を含む)。

    数にならない値は False。値は Val で読む。
    """
    if not vbx.is_numeric(caption):
        return False
    v = vbx.val(caption)
    if v < km or v > kp:
        return False
    if km <= v <= kp:
        return True
    return False


def diagonal_judg(tai: float, kou: float) -> bool:
    """対角線差の判定 `Judg.DiagonalJudg` (J:380-394)。引数は Single。

    公差 0 は False(差 0 でも NG)。差 0 は True。`Val(kou) >= Val(Tai)` は有効 7 桁の 10 進で比べる(V4)。
    """
    tai = vbx.single(tai)
    kou = vbx.single(kou)
    if kou == 0:
        return False
    if tai == 0:
        return True
    return vbx.val_single(kou) >= vbx.val_single(tai)


def lateral_judg(a: float, kou: str) -> bool:
    """ラテラルボー・直角度の判定 `Judg.LateralJudg` (J:395-409)。|a| ≤ 公差で OK。

    `kou` が "0" か "" なら False("0.0" は一致しないので 0 >= |a| の比較になる)。a は Single。
    """
    a = vbx.single(a)
    if kou == "0" or kou == "":
        return False
    if a == 0:
        return True
    return vbx.val(kou) >= abs(vbx.val_single(a))


def bh_judg(text: str, bh_lower: int) -> str:
    """BH(バーコール硬度)の判定 `Judg.BhJudg` (J:410-418)。

    `bh_lower` は BhSearch(材質, 調質) の結果(Long。担当外)。0 か入力が空なら "不要"。
    下限値 ≤ Val(入力) で "OK"、下限値 > Val(入力) で "NG"。
    """
    bh = int(bh_lower)
    if bh == 0 or text == "":
        return "不要"
    v = vbx.val(text)
    result = ""
    if bh > v:
        result = "NG"
    if bh <= v:
        result = "OK"
    return result


def unit_record(n: float) -> float:
    """0.5 刻みへの切り上げ `Judg.UnitRecord` (J:2421-2443、Q-S-QA-ALL-153)。n は Single。

    B13: n が Single なので 0.51・1.01・1.51・2.01・2.51・3.01・3.51 が区間の下限よりわずかに小さくなり 0 になる。
    999 を超えても 0。直した場合は有効 7 桁の 10 進(VBA の CStr(Single))で区間を比べ、5 を超えたらそのまま。
    0 のとき VBA は -0 になるが、-0 は 0 に揃える(V10)。
    """
    n = vbx.single(n)
    positive = n > 0                       # Sgn(n) = 1
    n = abs(n)
    if quirks.reproduce("B13"):
        x = n                              # Single を Double の区間と比べる
        upper = 999.0
    else:
        x = vbx.val_single(n)
        upper = float("inf")
    r = 0.0
    if x == 0:
        r = 0.0
    elif 0.1 <= x <= 0.5:
        r = 0.5
    elif 0.51 <= x <= 1.0:
        r = 1.0
    elif 1.01 <= x <= 1.5:
        r = 1.5
    elif 1.51 <= x <= 2.0:
        r = 2.0
    elif 2.01 <= x <= 2.5:
        r = 2.5
    elif 2.51 <= x <= 3.0:
        r = 3.0
    elif 3.01 <= x <= 3.5:
        r = 3.5
    elif 3.51 <= x <= 4.0:
        r = 4.0
    elif 4.01 <= x <= 4.5:
        r = 4.5
    elif 4.51 <= x <= 5.0:
        r = 5.0
    elif 5.01 <= x <= upper:
        r = n
    r = vbx.single(r)
    if not positive:
        r = vbx.single(vbx.val_single(r) * -1)
    return r + 0.0 if r != 0 else 0.0


def unit_record_input(text: str) -> str:
    """直角度・ラテラルボーの入力の前処理(UF1:1104-1110、1296-1300)。判定の前に欄へ書き戻す値。

    数でなければそのまま("-" = 測定不要もそのまま)。数なら `UnitRecord(Val(x))` を CStr(Single) で
    書いてから `Format(…, "0.0")`。
    """
    if not vbx.is_numeric(text):
        return text
    if text == "-":
        return text
    r = unit_record(vbx.single(vbx.val(text)))
    return vbx.fmt(vbx.cstr(r, is_single=True), "0.0")


# ------------------------------------------------------------------
# 1 マスの判定(色・音・画像・ログ)
# ------------------------------------------------------------------
def judg_color(form: Form, name: str, n: str, worker_name: str = "") -> Outcome:
    """1 マスの判定 `Judg.JudgColor` (J:482-747)。n: "A"=厚 / "H"=巾 / "T"=丈 / "D"=対角。

    `On Error Resume Next` の下なので、上下限・対角公差が数にならないと 0 のまま判定する
    (対角公差が空なら kou=0 で NG)。マイナス公差(Am/Hm/Tm)が空なら判定しない(J:533-535)。
    対角は 9 と 10 の両方が入っているときだけ判定し、差を Diag? に CStr(Single) で書く。
    """
    out = Outcome()
    lname = form.lname
    kp = km = 0.0
    kou = 0.0
    if n == "A":
        kp = _resume(lambda: vbx.cdbl(form.ApK), kp)
        km = _resume(lambda: vbx.cdbl(form.AmK), km)
    elif n == "H":
        kp = _resume(lambda: vbx.cdbl(form.HpK), kp)
        km = _resume(lambda: vbx.cdbl(form.HmK), km)
    elif n == "T":
        kp = _resume(lambda: vbx.cdbl(form.TpK), kp)
        km = _resume(lambda: vbx.cdbl(form.TmK), km)
    elif n == "D":
        kou = _resume(lambda: csng(form.TaikText), kou)

    caption = form.v(name)

    def ok_ng(ok: bool, cells: Sequence[str]) -> None:
        mark = "○" if ok else "×"
        out.logs.extend(_judge_logs(_ufken_text(caption, name, mark, lname, worker_name), lname))
        for c in cells:
            out.states[c] = OK if ok else NG
        out.sounds.append(SOUND_OK if ok else SOUND_NG)
        out.images.append(IMAGE_OK if ok else IMAGE_NG)

    if n in ("A", "H", "T"):
        if (n == "A" and form.Am == "") or (n == "H" and form.Hm == "") or (n == "T" and form.Tm == ""):
            return out                                            # GoTo SkipD
        ok = aht_judg(caption, kp, km)
        out.result = ok
        ok_ng(ok, [name])
    elif n == "D":
        row = name[:1]
        last = name[-1:]
        if last == "9":
            if form.v(row + "10") == "":
                text = f"*UFKen:{caption}:({name}):_:{lname}"      # 判定前の旧形式(J:587-588)
                out.logs.extend(_judge_logs(text, lname))
                return out                                        # GoTo SkipD
            tai = vbx.single(abs(vbx.val(form.v(row + "10")) - vbx.val(caption)))
            ok = diagonal_judg(tai, kou)
            ok_ng(ok, [row + "10", name])
        elif last == "0":
            if form.v(row + "9") == "":
                return out                                        # Exit Sub(ログなし)
            tai = vbx.single(abs(vbx.val(form.v(row + "9")) - vbx.val(caption)))
            ok = diagonal_judg(tai, kou)
            ok_ng(ok, [row + "9", name])
        else:
            tai = 0.0
            ok = None
        out.result = ok
        out.values["Diag" + row] = vbx.cstr(tai, is_single=True)   # UF1("Diag" & 行) = Tai(J:710)
    return out


def text_judg(form: Form, name: str, worker_name: str = "") -> Outcome:
    """測定ラベルの判定の振り分け `Judg.TextJudg` (J:748-835)。最後に NGErase(TB) を呼ぶ。

    行と項目: 1-4 厚、5-6 巾、7-8 丈、9-10 対角。極厚は 5-10 巾、11-14 丈、対角なし(G2 §3.0)。
    9 を入れたときは対角の音と猫の画像を出してから判定する(J:810-829)。
    """
    out = Outcome()
    if form.msgbox_visible:
        return out
    if form.v(name) == "":
        return out
    if form.AbnoTag == "NG":
        return out
    ln = len(name)
    r1 = name[-1:]
    r2 = name[-2:]
    if ln == 2 and r1 in ("1", "2", "3", "4"):
        out.merge(judg_color(form, name, "A", worker_name))
    if form.gokuatsu:
        if ln == 2 and r1 in ("5", "6", "7", "8", "9"):
            out.merge(judg_color(form, name, "H", worker_name))
        elif ln == 3 and r2 == "10":
            out.merge(judg_color(form, name, "H", worker_name))
    else:
        if ln == 2 and r1 in ("5", "6"):
            out.merge(judg_color(form, name, "H", worker_name))
    if form.gokuatsu:
        if ln == 3 and r2 in ("11", "12", "13", "14"):
            out.merge(judg_color(form, name, "T", worker_name))
    else:
        if ln == 2 and r1 in ("7", "8"):
            out.merge(judg_color(form, name, "T", worker_name))
    if not form.gokuatsu:
        if ln == 3 and r2 == "10":
            out.merge(judg_color(form, name, "D", worker_name))
        elif ln == 2 and r1 == "9":
            out.sounds.append(SOUND_DIAG)
            out.images.append(IMAGE_NEKO)
            out.merge(judg_color(form, name, "D", worker_name))
    out.ng_erase = name
    return out


def lateralbow_judg_color(form: Form, name: str, worker_name: str = "") -> Outcome:
    """ラテラルボーの判定 `Judg.LateralbowJudgColor` (J:1796-1892)。

    空・"-"・数でない入力は何もしない。公差(LateralText)が数でなければ「公差がありません」を出し、
    Tag="範囲外" にして判定しない。最後に MPEX.Tag を "" に戻す。
    """
    out = Outcome()
    text = form.v(name)
    if text == "" or text == "-" or not vbx.is_numeric(text):
        return out
    a = csng(text)
    lname = form.lname
    if not vbx.is_numeric(form.LateralText):
        out.messages.append(flow.Message(LATERAL_NO_TOL_MSG, flow.INFO, source="Judg:1808"))
        out.mpex_tag = ""
        out.tags[name] = "範囲外"
        return out
    kou = form.LateralText
    ok = lateral_judg(a, kou)
    out.result = ok
    out.logs.extend(_judge_logs(_ufken_text(text, name, "○" if ok else "×", lname, worker_name), lname))
    out.states[name] = OK if ok else NG
    out.sounds.append(SOUND_OK if ok else SOUND_NG)
    out.images.append(IMAGE_OK if ok else IMAGE_NG)
    out.mpex_tag = ""
    return out


def right_angle_judg_color(form: Form, name: str, worker_name: str = "") -> Outcome:
    """直角度の判定 `Judg.RightAngleJudgColor` (J:1894-1985)。公差は 1.0 固定(R115 のときだけ表示される欄)。

    `Dim kou As String: kou = 1#` なので公差の文字列は "1"。空・数でない入力は何もしない。
    """
    out = Outcome()
    text = form.v(name)
    if text == "" or not vbx.is_numeric(text):
        return out
    a = csng(text)
    lname = form.lname
    kou = vbx.cstr(RIGHT_ANGLE_TOL)          # "1"
    ok = lateral_judg(a, kou)
    out.result = ok
    out.logs.extend(_judge_logs(_ufken_text(text, name, "○" if ok else "×", lname, worker_name), lname))
    out.states[name] = OK if ok else NG
    out.sounds.append(SOUND_OK if ok else SOUND_NG)
    out.images.append(IMAGE_OK if ok else IMAGE_NG)
    out.mpex_tag = ""
    return out


def simply_judge(form: Form, name: str, n: str, worker_name: str = "") -> Outcome:
    """BH の色・音・ログ `Judg.simplyjudge` (J:420-478)。n は bh_judg の結果。

    NG のとき、最後に検査した位置(1 行目の S1〜E1 の入り方)で工程の指示を出す:
    S だけ検査済みで SBH が NG → BH_MSG_FIRST、E まで検査済みで EBH が NG → BH_MSG_LAST。
    比較は `Iti = UF1.S.Caption` / `Iti = UF1.E.Caption` なので、Iti が "" で見出しが "" の位置があると
    そちらのメッセージも出る(VBA どおり)。
    """
    out = Outcome()
    lname = form.lname
    s1, u1, c1, d1, e1 = (form.v(p + "1") for p in POS5)
    count_l = {
        1: s1 != "" and u1 == "" and c1 == "" and d1 == "" and e1 == "",
        2: u1 != "" and c1 == "" and d1 == "" and e1 == "",
        3: c1 != "" and d1 == "" and e1 == "",
        4: d1 != "" and e1 == "",
        5: e1 != "",
    }
    iti = ""
    if count_l[1]:
        if name == "SBH":
            iti = form.captions["S"]
    elif count_l[5]:
        if name == "EBH":
            iti = form.captions["E"]
    text = form.v(name)
    if n == "OK":
        out.states[name] = OK
        out.sounds.append(SOUND_OK)
        out.logs.extend(_judge_logs(f"*BH|{text}|({name})|○|{lname}|{worker_name}", lname))
    elif n == "NG":
        out.states[name] = NG
        out.sounds.append(SOUND_NG)
        out.logs.extend(_judge_logs(f"*BH|{text}|({name})|×|{lname}|{worker_name}", lname))
        if iti == form.captions["S"]:
            out.msg_errors.append(MsgErrorCall(BH_MSG_FIRST, lname, 1))
        if iti == form.captions["E"]:
            out.msg_errors.append(MsgErrorCall(BH_MSG_LAST, lname, 1))
    elif n == "不要":
        out.states[name] = WHITE
    return out


def hizumi_judg(form: Form, name: str, loop_j: str = "", *, move_flag: Optional[str] = None) -> Outcome:
    """歪 1 マスの判定 `Judg.HizumiJudg` (J:134-291)。name は ZS〜ZF(定盤)か GS〜GF(ゲージ)。

    順序: 空なら白 → フリーソリ(G 欄、"/" より前 > HURIS で NG) → CB?1 が空なら終了 →
    両方ゲージ測定(LZS="幅ｹﾞｰｼﾞ" かつ LGS="丈ｹﾞｰｼﾞ")なら Z は CB?1、G は CB?3 の幅/丈公差と比べる →
    通常は Z 欄だけ CB?1 の公差と比べ(公差ちょうどは OK)、G 欄は白。
    音は loop_j が "" か MoveFlag が "on" のときだけ。判定したら MoveFlag を "" に戻す。
    両方ゲージの分岐は "/" を処理せず CDbl(TB) なので、"0.8/3" や公差 "" は型エラー 13(vbx.TypeMismatch)。
    """
    out = Outcome()
    mf = form.MoveFlag if move_flag is None else move_flag
    tb = form.v(name)
    if tb == "":
        out.states[name] = WHITE
        return out
    lor_r = name[:1]
    temp = name[-1:]
    cb_z = form.combos["CB" + temp + "1"]
    cb_g = form.combos["CB" + temp + "3"]
    sound = loop_j == "" or mf == "on"

    def color(ng: bool) -> None:
        out.states[name] = NG if ng else OK
        if sound:
            out.sounds.append(SOUND_NG if ng else SOUND_OK)

    # フリーソリ(2026.03.05、J:155-176)
    if lor_r == "G" and form.HURIS != "" and vbx.is_numeric(form.HURIS) and form.flatnesText == "ﾌﾘｰｿﾘ要":
        idx2 = vbx.instr(tb, "/")
        number_t2 = vbx.left(tb, idx2 - 1) if idx2 > 0 else tb
        if vbx.is_numeric(number_t2):
            color(vbx.cdbl(number_t2) > vbx.cdbl(form.HURIS))
        out.move_flag = ""
        return out

    if cb_z == "":
        return out

    cv = ""
    if cb_z == "幅":
        cv = form.HaHiz
    elif cb_z == "丈":
        cv = form.TaHiz
    elif cb_z == "中":
        cv = form.Naka
    elif cb_z == "耳":
        cv = form.Mimi

    # 両方ゲージ測定(J:198-231)
    if form.captions.get("LZS", "") == "幅ｹﾞｰｼﾞ" and form.captions.get("LGS", "") == "丈ｹﾞｰｼﾞ":
        if cb_z == "幅":
            cv = form.HaHiz
        elif cb_z == "丈":
            cv = form.TaHiz
        cvg = ""
        if cb_g == "幅":
            cvg = form.HaHiz
        elif cb_g == "丈":
            cvg = form.TaHiz
        if lor_r == "Z" and vbx.is_numeric(cv):
            c = vbx.vcompare(vbx.cdbl(tb), cv)
            if c > 0:
                color(True)
            elif c <= 0:
                color(False)
        elif lor_r == "G":
            if cb_g == "":
                return out
            c = vbx.vcompare(vbx.cdbl(tb), cvg)
            if c > 0:
                color(True)
            elif c <= 0:
                color(False)
        out.move_flag = ""
        return out

    skip_g = False
    if cv == "":
        if cb_g != "":
            skip_g = True
    if not skip_g and not vbx.is_numeric(cv):
        return out
    idx = vbx.instr(tb, "/")
    number_t = vbx.left(tb, idx - 1) if idx > 0 else tb
    if not vbx.is_numeric(number_t):
        return out
    if number_t == "":
        return out
    if lor_r == "Z" and vbx.is_numeric(cv):
        c = vbx.vcompare(vbx.cdbl(number_t), cv)
        if c > 0:
            color(True)
        elif c <= 0:
            color(False)
    else:
        out.states[name] = WHITE
    out.move_flag = ""
    return out


def hiz_loop_judg(form: Form, n: int = 0, tmp: str = "Loop") -> Outcome:
    """歪 12 マスの再判定 `Judg.HizLoopJudg` (J:293-332)。Z 欄 S→F、G 欄 S→F の順。空欄は白。

    MoveFlag は1マス判定するたびに "" に戻るので、"on" で音が鳴るのは最初の1マスだけ。
    """
    out = Outcome()
    if n >= 2:
        return out
    if form.Lot == "":
        return out
    if form.HaHiz == "" and form.TaHiz == "":
        return out
    if form.kosei_visible:
        return out
    mf = form.MoveFlag
    names = [f"Z{p}" for p in POSITIONS] + [f"G{p}" for p in POSITIONS]
    for nm in names:
        if form.v(nm) != "":
            sub = hizumi_judg(form, nm, tmp, move_flag=mf)
            if sub.move_flag is not None:
                mf = sub.move_flag
            out.merge(sub)
    for nm in names:
        if form.v(nm) == "":
            out.states[nm] = WHITE
    if mf != form.MoveFlag:
        out.move_flag = mf
    return out


# ------------------------------------------------------------------
# 画面全体の再判定
# ------------------------------------------------------------------
def re_diag(form: Form) -> Outcome:
    """再読み込み時の対角線差の再計算 `Judg.ReDiag` (J:370-379)。極厚でないときだけ、`Format(…, "0.0")`。"""
    out = Outcome()
    if not form.gokuatsu:
        for p in POSITIONS:
            a, b = form.v(f"{p}9"), form.v(f"{p}10")
            if a != "" and b != "":
                out.values["Diag" + p] = vbx.fmt(abs(vbx.val(a) - vbx.val(b)), "0.0")
    return out


def loop_judg(form: Form, n: int = 0, j: int = 1, k: int = 6, *, bh_lower: int = 0) -> Outcome:
    """画面全体の再判定 `Judg.LoopJudg` (J:836-1116)。j〜k は位置(1=S … 6=F)。

    1. 全マスを白・黒枠に戻し、値のあるマス・表示中のラテラルボー欄・直角度欄を NG にする
       (マイナス公差 Am/Hm/Tm が空の行は白。ただし F 列は戻さない)
    2. 厚・巾・丈・対角を判定し、OK のマスだけ緑にする
       - 巾の上下限が空なら巾をとばす。丈 7-8 の上下限が空なら丈と**対角も**とばす(J:983)
       - 極厚の丈 11-14 は上下限の空欄を見ない(空なら型エラー 13 → LoopJudgAborted)
    3. LateralText が数でなければここで終わる(以降の欄は NG 色のまま、J:1059)
    4. ラテラルボー・BH(OK は白、NG はコーラル、ほかは変えない)・直角度(公差 "1.0")
    5. ラテラルボー・直角度の "" と "-" は白
    B8: 直角度の D と E の事前の色付けが入れ替わっている(J:927-928)。
    `bh_lower` は BhSearch(Zai, Tyu) の結果(担当外)。
    """
    out = Outcome()
    if n >= 2:
        return out
    if form.ApK == "" or form.AmK == "":
        return out
    if form.kosei_visible:
        return out
    try:
        _loop_judg_body(form, j, k, bh_lower, out)
    except vbx.TypeMismatch as e:
        raise LoopJudgAborted(out, e) from e
    return out


def _pos(a: int) -> Optional[str]:
    return POSITIONS[a - 1] if 1 <= a <= 6 else None


def _loop_judg_body(form: Form, j: int, k: int, bh_lower: int, out: Outcome) -> None:
    st = out.states
    v = form.v
    # 枠線リセット(J:850-865)
    for i in range(1, 15):
        for p in POS5:
            out.borders[f"{p}{i}"] = ""
            st[f"{p}{i}"] = WHITE
    for i in range(1, 11):
        out.borders[f"F{i}"] = ""
        st[f"F{i}"] = WHITE
        if v(f"F{i}") != "":
            st[f"F{i}"] = NG
    # 正以外の色付け(J:868-878)
    for i in range(1, 15):
        for p in POS5:
            if v(f"{p}{i}") != "":
                st[f"{p}{i}"] = NG
    for i in range(1, 11):
        if v(f"F{i}") != "":
            st[f"F{i}"] = NG
    # マイナス公差が空の行は白(J:881-910、F 列は対象外)
    for blank, rows in ((form.Am, range(1, 5)), (form.Hm, range(5, 7)), (form.Tm, range(7, 9))):
        if blank == "":
            for i in rows:
                for p in POS5:
                    if v(f"{p}{i}") != "":
                        st[f"{p}{i}"] = WHITE
    # ラテラルボー・直角度は表示中なら NG(J:913-929)
    for p in POSITIONS:
        if form.cells[f"Lateral{p}"].visible:
            st[f"Lateral{p}"] = NG
    kou_targets = [("S", "S"), ("U", "U"), ("C", "C")]
    if quirks.reproduce("B8"):
        kou_targets += [("E", "D"), ("D", "E")]         # 表示を見る欄と塗る欄が逆
    else:
        kou_targets += [("D", "D"), ("E", "E")]
    kou_targets.append(("F", "F"))
    for seen, painted in kou_targets:
        if form.cells[f"Diag{seen}Kou"].visible:
            st[f"Diag{painted}Kou"] = NG

    # 厚(J:931-946)
    def thick() -> None:
        for i in range(1, 5):
            for a in range(j, k + 1):
                if form.Am == "":
                    return                                   # GoTo SkipH
                p = _pos(a)
                if p and aht_judg(v(f"{p}{i}"), vbx.cdbl(form.ApK), vbx.cdbl(form.AmK)):
                    st[f"{p}{i}"] = OK
    thick()

    # 巾(J:948-967)
    def width() -> None:
        x = 10 if form.gokuatsu else 6
        for i in range(5, x + 1):
            for a in range(j, k + 1):
                if form.HpK == "" or form.HmK == "":
                    return                                   # GoTo SkipT
                p = _pos(a)
                if p and aht_judg(v(f"{p}{i}"), vbx.cdbl(form.HpK), vbx.cdbl(form.HmK)):
                    st[f"{p}{i}"] = OK
    width()

    # 丈と対角(J:969-1056)。丈 7-8 の上下限が空なら対角もとばす
    def length_and_diag() -> None:
        x, y = (14, 11) if form.gokuatsu else (8, 7)
        for i in range(y, x + 1):
            if 7 <= i <= 8:
                for a in range(j, k + 1):
                    if form.TpK == "" or form.TmK == "":
                        return                               # GoTo SkipL
                    p = _pos(a)
                    if p and aht_judg(v(f"{p}{i}"), vbx.cdbl(form.TpK), vbx.cdbl(form.TmK)):
                        st[f"{p}{i}"] = OK
            elif 11 <= i <= 14:
                for a in range(j, k + 1):
                    p = _pos(a)
                    if p and p != "F" and aht_judg(v(f"{p}{i}"), vbx.cdbl(form.TpK), vbx.cdbl(form.TmK)):
                        st[f"{p}{i}"] = OK
        if vbx.is_numeric(form.TaikText):
            for a in range(j, k + 1):
                p = _pos(a)
                if p and v(f"Diag{p}") != "":
                    if diagonal_judg(csng(v(f"Diag{p}")), csng(form.TaikText)):
                        st[f"{p}9"] = OK
                        st[f"{p}10"] = OK
    length_and_diag()

    # SkipL: ラテラルボー(J:1059-1069)
    if not vbx.is_numeric(form.LateralText):
        return
    for a in range(j, k + 1):
        p = _pos(a)
        if p and vbx.is_numeric(v(f"Lateral{p}")):
            if lateral_judg(csng(v(f"Lateral{p}")), form.LateralText):
                st[f"Lateral{p}"] = OK
    # BH(J:1071-1088): OK は白、NG はコーラル
    for want, color in (("OK", WHITE), ("NG", NG)):
        for a in range(j, k + 1):
            p = _pos(a)
            if p and p != "F" and vbx.is_numeric(v(f"{p}BH")):
                if bh_judg(v(f"{p}BH"), bh_lower) == want:
                    st[f"{p}BH"] = color
    # 直角度(J:1090-1099)
    for a in range(j, k + 1):
        p = _pos(a)
        if p and vbx.is_numeric(v(f"Diag{p}Kou")):
            if lateral_judg(csng(v(f"Diag{p}Kou")), "1.0"):
                st[f"Diag{p}Kou"] = OK
    for pre, post in (("Lateral", ""), ("Diag", "Kou")):
        for p in POSITIONS:
            nm = f"{pre}{p}{post}"
            if v(nm) == "" or v(nm) == "-":
                st[nm] = WHITE


def looking_yellow(form: Form) -> Outcome:
    """黄色の背景と赤枠を消す `Judg.LookingYellow` (J:2444-2466)。"""
    out = Outcome()
    names = [f"{p}{i}" for i in range(1, 15) for p in POS5] + [f"F{i}" for i in range(1, 11)]
    for nm in names:
        c = form.cells[nm]
        if c.border == RED:
            out.borders[nm] = ""
        if c.state == YELLOW:
            out.states[nm] = WHITE
    return out


# ------------------------------------------------------------------
# 公差表(品質基準.製品規格直角度 / 製品規格ラテラルボー)
# ------------------------------------------------------------------
@dataclass
class _TolTables:
    mtime: Optional[float]
    loaded: bool = False
    failed: bool = False
    diag: Optional[list[list[str]]] = None
    diag_idx: list[int] = field(default_factory=list)
    lat: Optional[list[list[str]]] = None
    lat_idx: list[int] = field(default_factory=list)


_tol_lock = threading.Lock()
_tol_cache: dict[str, _TolTables] = {}


class _LoadError(Exception):
    def __init__(self, number: int, text: str) -> None:
        super().__init__(text)
        self.number = number
        self.text = text


def _load_table(db: Union[str, Path], table: str, cols: Sequence[str]) -> tuple[Optional[list[list[str]]], list[int]]:
    """`Judg.LoadTable` (J:1250-1274)。列名から列番号を引き、全行を取る。"""
    try:
        flds = datasource.get_fields(db, table, "")
    except (datasource.DbMissing, datasource.TableMissing) as e:
        raise _LoadError(0, str(e)) from e      # VBA の結果は U7 実機確認待ち
    if flds is None:
        raise _LoadError(VB_OBJECT_ERROR + 1, f"「{table}」のフィールドを取得できません")
    idx: list[int] = []
    for c in cols:
        i = datasource.column_index(flds, c)
        if i == -1:
            raise _LoadError(VB_OBJECT_ERROR + 2, f"「{table}」に「{c}」フィールドがありません")
        idx.append(i)
    return datasource.get_records(db, table, ""), idx


def _ensure_loaded(db: Union[str, Path, None], messages: Optional[list[flow.Message]]) -> Optional[_TolTables]:
    """`Judg.EnsureLoaded` (J:1214-1248)。

    VBA は初回だけ読み、失敗したら ReloadTolerance まで再試行しない(MsgBox は1回だけ)。
    D11 により、ファイルが更新されたら読み直す(同じファイルなら結果は VBA と同じ)。
    """
    db = db if db is not None else "品質基準"
    p = datasource.resolve(db)
    key = str(p)
    try:
        mtime: Optional[float] = p.stat().st_mtime
    except OSError:
        mtime = None
    with _tol_lock:
        st = _tol_cache.get(key)
        if st is not None and st.mtime == mtime:
            if st.loaded:
                return st
            if st.failed:
                return None
        st = _TolTables(mtime=mtime)
        step = ""
        try:
            step = TBL_DIAG + " 読込"
            st.diag, st.diag_idx = _load_table(p, TBL_DIAG, DIAG_COLS)
            step = TBL_LAT + " 読込"
            st.lat, st.lat_idx = _load_table(p, TBL_LAT, LAT_COLS)
            st.loaded = True
        except _LoadError as e:
            st.failed = True
            if messages is not None:
                text = ("公差テーブルを読み込めませんでした" + "\r\n" + "段階: " + step + "\r\n"
                        + "内容: " + str(e.number) + " " + e.text + "\r\n" + "\r\n"
                        + "公差は表示されません。原因を解消後 ReloadTolerance を実行してください。")
                messages.append(flow.Message(text, flow.ERROR, source="Judg:1243"))
        _tol_cache[key] = st
        return st if st.loaded else None


def reload_tolerance(db: Union[str, Path, None] = None) -> list[flow.Message]:
    """公差表の読み直し `Judg.ReloadTolerance` (J:1138-1146)。読めたら案内を返す。"""
    with _tol_lock:
        _tol_cache.clear()
    msgs: list[flow.Message] = []
    if _ensure_loaded(db, msgs) is not None:
        msgs.append(flow.Message("公差テーブルを読み込み直しました", flow.INFO, source="Judg:1145"))
    return msgs


def _in_range(v: float, lo: str, hi: str) -> bool:
    """`Judg.InRange` (J:1205-1209)。下限 ≤ v ≤ 上限。空欄(Null)は不一致。And は両辺とも評価する。"""
    if len(lo) == 0 or len(hi) == 0:
        return False
    lo_d = vbx.cdbl(lo)
    hi_d = vbx.cdbl(hi)
    return v >= lo_d and v <= hi_d


def _lookup(rows: Optional[list[list[str]]], idx: list[int], grade: str,
            conds: Sequence[tuple[float, int, int]], tol_pos: int) -> float:
    """範囲がすべて合う行のうち管理番号の最小の行の公差(CSng)。該当なしは 0。"""
    if rows is None:
        return 0.0
    result = 0.0
    best_no = 0.0
    hit = False
    for r in rows:
        if r[idx[1]] != grade:                               # CStr(等級 & "") = grade
            continue
        if not all_in_range(r, idx, conds):
            continue
        # `Not hit Or 管理番号 < bestNo` は短絡しないので、比較は毎回する(数にならなければ型エラー)
        less = vbx.vcompare(r[idx[0]], best_no) < 0
        if (not hit) or less:
            hit = True
            best_no = vbx.cdbl(r[idx[0]])
            result = csng(r[idx[tol_pos]])
    return result


def all_in_range(r: Sequence[str], idx: list[int], conds: Sequence[tuple[float, int, int]]) -> bool:
    """入れ子の If InRange(…) Then(先に外れたら後ろは評価しない)。"""
    for v, lo_pos, hi_pos in conds:
        if not _in_range(v, r[idx[lo_pos]], r[idx[hi_pos]]):
            return False
    return True


def _round2(x: float) -> float:
    """`Round(CDbl(Single), 2)`(Single の誤差対策。銀行丸め)。"""
    return vbx.vround(vbx.single(x), 2)


def lookup_diag(grade: str, a: float, b: float, *, db: Union[str, Path, None] = None,
                messages: Optional[list[flow.Message]] = None) -> float:
    """直角度(対角線差)の公差 `Judg.LookupDiag` (J:1153-1176)。a=巾、b=丈(Single)。"""
    st = _ensure_loaded(db, messages)
    if st is None:
        return 0.0
    w, l_ = _round2(a), _round2(b)
    return _lookup(st.diag, st.diag_idx, grade, ((w, 2, 3), (l_, 4, 5)), 6)


def lookup_lat(grade: str, a: float, b: float, c: float, *, db: Union[str, Path, None] = None,
               messages: Optional[list[flow.Message]] = None) -> float:
    """ラテラルボーの公差 `Judg.LookupLat` (J:1178-1203)。a=巾、b=丈、c=厚(Single)。厚→巾→丈の順に見る。"""
    st = _ensure_loaded(db, messages)
    if st is None:
        return 0.0
    w, l_, t = _round2(a), _round2(b), _round2(c)
    return _lookup(st.lat, st.lat_idx, grade, ((t, 2, 3), (w, 4, 5), (l_, 6, 7)), 8)


def diagonal(a: float, b: float, **kw: Any) -> float:
    """4級(厳しい)直角度 `Judg.Diagonal` (J:1118-1121)。(巾, 丈)"""
    return lookup_diag("4", a, b, **kw)


def re_diagonal(a: float, b: float, **kw: Any) -> float:
    """3級(緩い)直角度 `Judg.ReDiagonal` (J:1123-1126)。(巾, 丈)"""
    return lookup_diag("3", a, b, **kw)


def lateralbow_j(a: float, b: float, c: float, **kw: Any) -> float:
    """4級ラテラルボー `Judg.Lateralbow_J` (J:1128-1131)。(巾, 丈, 厚)"""
    return lookup_lat("4", a, b, c, **kw)


def re_lateralbow(a: float, b: float, c: float, **kw: Any) -> float:
    """3級ラテラルボー `Judg.ReLateralbow` (J:1133-1136)。(巾, 丈, 厚)"""
    return lookup_lat("3", a, b, c, **kw)


# ------------------------------------------------------------------
# 先行データ票・包装仕様の注意文
# ------------------------------------------------------------------
Arrbox = Callable[[str, str, str, str], Optional[list]]


def simple_arrbox(target: str, sql_filter: str = "", use: str = "Record", path: str = "") -> Optional[list]:
    """`AccHIki.Arrbox` の最小限(担当外)。path が "" なら MainPath、そうでなければ NextPath の「仕掛」。

    本物の Arrbox は ReadCheck・TimeDesignation・FileDesignation で控えに切り替える。それは AccHIki の担当が作る。
    """
    name = target.rsplit(".", 1)[0]
    db = paths.db_file(name) if path == "" else paths.db_file(name, base_key="next")
    if sql_filter == "":
        flt = ""
    elif sql_filter == "板":
        flt = "[ｵｰﾀﾞｰ板丈]<>0"
    elif sql_filter == "コイル":
        flt = "[ｵｰﾀﾞｰ板丈]=0  AND  [BOX設計_設備名] = 'KEN' "
    else:
        flt = "[ﾛｯﾄ番号]='" + sql_filter + "'"
    if use == "Fields":
        return datasource.get_fields(db, "仕掛", "")
    return datasource.get_records(db, "仕掛", flt)


def advance_check(lot: str, db_flag: str = "", *, arrbox: Optional[Arrbox] = None) -> bool:
    """先行データ票が必要か `Judg.AdvanceCheck` (J:334-361)。

    SIKALOTNOW「仕掛」のロットの 1 行目で、
    flag1: Val(製造板厚) ≤ 3(Single)、flag2: 品質ｸﾞﾚｰﾄﾞ_表面処理 が S でも T でもない、
    flag3: 用途ｺｰﾄﾞ の頭が T か D1・D2。(flag1 And flag2) Or (flag2 And flag3) で True。
    """
    if lot == "":
        return False
    ab = arrbox or simple_arrbox
    fields = ab("SIKALOTNOW.accdb", "", "Fields", db_flag)
    if fields is None:
        return False
    recs = ab("SIKALOTNOW.accdb", lot, "Record", db_flag)
    if recs is None:
        return False
    c_surface = _last_index(fields, "品質ｸﾞﾚｰﾄﾞ_表面処理")
    c_thick = _last_index(fields, "製造板厚")
    c_use = _last_index(fields, "用途ｺｰﾄﾞ")
    row = recs[0]
    thick = vbx.single(vbx.val(_at(row, c_thick)))
    surface = _at(row, c_surface)
    youto = _at(row, c_use)
    f1 = thick <= 3.0
    f2 = surface != "S" and surface != "T"
    f3 = vbx.left(youto, 1) == "T" or vbx.left(youto, 2) == "D1" or vbx.left(youto, 2) == "D2"
    return (f1 and f2) or (f2 and f3)


def _text_instr(target: str, core: str) -> bool:
    """`InStr(1, target, core, vbTextCompare) > 0`(大文字小文字・全角半角を区別しない。vbx.text_eq と同じ揃え方)。"""
    if core == "":
        return True
    na = unicodedata.normalize("NFKC", target).casefold()
    nb = unicodedata.normalize("NFKC", core).casefold()
    return nb in na


def is_wildcard_match(pattern: str, target: str) -> bool:
    """`Judg.IsWildcardMatch` (J:2390-2411)。

    `*xxx*` 部分一致(テキスト比較)、`xxx*` 前方一致(バイナリ)、`*xxx` 後方一致(バイナリ)、`xxx` 完全一致(テキスト)。
    パターンが "*" 1 文字だと `Left("", -1)` で実行時エラー 5(VBA どおり)。
    """
    has_head = vbx.left(pattern, 1) == "*"
    has_tail = vbx.right(pattern, 1) == "*"
    core = pattern
    if has_head:
        core = vbx.mid(core, 2)
    if has_tail:
        if len(core) - 1 < 0:
            raise VbaError(5, "プロシージャの呼び出し、または引数が不正です")
        core = vbx.left(core, len(core) - 1)
    if has_head and has_tail:
        return _text_instr(target, core)
    if has_tail:
        return vbx.left(target, len(core)) == core
    if has_head:
        return vbx.right(target, len(core)) == core
    return vbx.text_eq(target, core)


def is_match_condition_field(field_value: str, target: str) -> bool:
    """`Judg.IsMatchConditionField` (J:2346-2388)。マスタの条件が target に合うか。

    空・"＊"・"*" は無条件で True。カンマ区切りは OR。先頭 "≠" は「合わなければ True」(除外どうしも OR)。
    """
    cond = vba_trim(field_value)
    if cond == "" or cond == "＊" or cond == "*":
        return True
    for tok in cond.split(","):
        token = vba_trim(tok)
        if token == "":
            continue
        exclude = False
        if vbx.left(token, 1) == "≠":
            exclude = True
            token = vbx.mid(token, 2)
        hit = is_wildcard_match(token, target)
        if exclude:
            if not hit:
                return True
        elif hit:
            return True
    return False


def _dims_in(row: Sequence[str], c: dict[str, int], thick: float, width: float, length: float) -> bool:
    """寸法の条件(両端を含む)。厚・幅・丈は Single のまま、Val(マスタ) の Double と比べる。"""
    # VBA の And は短絡しないので、先に全部読む(列が無ければ実行時エラー 9)
    lo_t, hi_t = vbx.val(_at(row, c["厚下"])), vbx.val(_at(row, c["厚上"]))
    lo_w, hi_w = vbx.val(_at(row, c["板幅下"])), vbx.val(_at(row, c["板幅上"]))
    lo_l, hi_l = vbx.val(_at(row, c["板丈下"])), vbx.val(_at(row, c["板丈上"]))
    return lo_t <= thick <= hi_t and lo_w <= width <= hi_w and lo_l <= length <= hi_l


def packing_master_path() -> str:
    """`PackingPath & "梱包資材マスタ.accdb"`(SQLite に切り替え済み)。"""
    return paths.join(paths.get("packing"), PACKING_MASTER_FILE)


_PACK_COLS = ("用途コード", "包装仕様NO", "コメント", "備考", "受注材質", "受注調質", "厚下", "厚上",
              "板幅下", "板幅上", "板丈下", "板丈上", "フラグ", "形状", "納入先", "取引先")


def packaging_specification_no(packaging_no: str, youto: str, thick: float, width: float, length: float,
                               zai: str, tyo: str, nounyu: str = "", torihiki: str = "", *,
                               db: Union[str, Path, None] = None) -> str:
    """包装仕様NO の注意文 `Judg.PackagingSpecificationNo` (J:2188-2344)。結果は "形状_コメント…"。

    梱包資材マスタ「注意_包装仕様」を `[包装仕様NO]='…'` で引く。無ければ ""。
    1 件: 条件(フラグ: ""/用途コード/寸法/納入先/寸法納入/寸法取引)に合えばコメント、合わなければ "形状_"。
    2 件以上: 合ったコメントを前に足していく("C,B,A," のように DB と逆順で末尾に "," が残る)。
    B11: 「寸法納入」に合うと前に足さずに上書きする(それまでのコメントが消える)。
    厚・幅・丈は Single なので、上限ちょうどの値が Single の誤差で外れることがある(VBA どおり)。
    """
    db = db if db is not None else packing_master_path()
    table = "注意_包装仕様"
    datasource.get_fields(db, table, "")                     # Hikifields(使っていない)
    pack_f = datasource.get_fields(db, table, "")
    if pack_f is None:
        # `Dim PackF(): PackF = Empty` は動的配列に Empty を入れるので型エラー 13(IsEmpty の判定まで行かない)
        raise vbx.TypeMismatch("PackF = GetFieldsArr(…) が Empty")
    c = {name: _last_index(pack_f, name) for name in _PACK_COLS}
    rows = datasource.get_records(db, table, "[包装仕様NO]='" + packaging_no + "'")
    if rows is None:
        return ""
    thick, width, length = vbx.single(thick), vbx.single(width), vbx.single(length)

    def matched(r: Sequence[str]) -> Optional[str]:
        """フラグごとの条件。合えばそのコメント、合わなければ None。"""
        flag = _at(r, c["フラグ"])
        if flag == "":
            return _at(r, c["コメント"])
        if flag == "用途コード":
            return _at(r, c["コメント"]) if youto == _at(r, c["用途コード"]) else None
        if flag == "寸法":
            return _at(r, c["コメント"]) if _dims_in(r, c, thick, width, length) else None
        if flag == "納入先":
            if nounyu != "" and is_match_condition_field(_at(r, c["納入先"]), nounyu):
                return _at(r, c["コメント"])
            return None
        if flag == "寸法納入":
            if nounyu != "" and is_match_condition_field(_at(r, c["納入先"]), nounyu):
                if _dims_in(r, c, thick, width, length):
                    return _at(r, c["コメント"])
            return None
        if flag == "寸法取引":
            if torihiki != "":
                dims = _dims_in(r, c, thick, width, length)
                m1 = is_match_condition_field(_at(r, c["納入先"]), nounyu)
                m2 = is_match_condition_field(_at(r, c["取引先"]), torihiki)
                if dims and m1 and m2:
                    return _at(r, c["コメント"])
            return None
        return None

    result = ""
    if len(rows) == 1:
        m = matched(rows[0])
        if m is not None:
            result = m
        return _at(rows[0], c["形状"]) + "_" + result
    for r in rows:
        m = matched(r)
        if m is None:
            continue
        if _at(r, c["フラグ"]) == "寸法納入" and quirks.reproduce("B11"):
            result = m                                       # 上書き(J:2306)
        else:
            result = m + "," + result
    return _at(rows[0], c["形状"]) + "_" + result


def _split(s: str, sep: str) -> list[str]:
    """VBA `Split`。空文字は要素 0 個(UBound = -1)。"""
    return [] if s == "" else s.split(sep)


def materials_msgbox(form: Form, ord_no: str, *,
                     arrbox: Optional[Arrbox] = None,
                     read_check: Optional[Callable[[str], bool]] = None,
                     time_designation: Callable[[], bool] = lambda: False,
                     file_designation: Callable[[], bool] = lambda: True,
                     packing_db: Union[str, Path, None] = None) -> Outcome:
    """包装仕様の注意文と先行データ票の案内 `Judg.MaterialsMsgBox` (J:1987-2122)。

    SIKAODRNOW「仕掛」をオーダーNo で引き、包装仕様NO・用途コード・材質・調質・納入先・取引先を取る。
    厚・巾・丈は画面の製造寸法(ATU/Hab/Tak)。PackagingSpecificationNo の結果を "_" で分け、
    コイルなら何も出さない。注意文は MsgError(…, 4) で出す。
    B11: 結果が "板_"(コメントなし)のとき、先行データ票の案内も出さずに終わる。
    `read_check` / `time_designation` / `file_designation` は担当外(既定: ファイルがあるか / 偽 / 真)。
    """
    out = Outcome()
    if form.Lot == "":
        return out
    lname = form.lname
    main, nxt = paths.get("main"), paths.get("next")
    fpath = main if form.DBflag == "" else nxt
    target = "SIKAODRNOW.accdb"
    rc = read_check or (lambda p: datasource.exists(p))

    def file_of(base: str) -> str:
        return paths.join(base, "SIKAODRNOW.sqlite3")

    file_path = file_of(fpath)
    if (not rc(file_path)) or time_designation() or (not file_designation()):
        missing = "日稲:" + target + "がありません" + "_MaterialsMsgBox"
        if fpath == nxt:
            out.messages.append(flow.Message(missing, flow.ERROR, source="Judg:2004"))
            return out
        if fpath == main:
            fpath = nxt
            file_path = file_of(fpath)
            if not rc(file_path):
                out.messages.append(flow.Message(missing, flow.ERROR, source="Judg:2010"))
                return out

    fields = datasource.get_fields(file_path, "仕掛", "")
    names = ("受注番号", "包装仕様NO", "受注板厚", "受注板幅", "受注板丈", "受注材質", "受注調質",
             "用途ｺｰﾄﾞ", "合紙", "納入先名称", "取引先名称")
    col = {n: (_last_index(fields, n) if fields is not None else -1) for n in names}
    rows = datasource.get_records(file_path, "仕掛", "[受注番号]='" + ord_no + "'")
    if rows is None:
        return out
    r = rows[0]
    packing_no = _at(r, col["包装仕様NO"])
    youto = _at(r, col["用途ｺｰﾄﾞ"])
    thick = vbx.single(vbx.val(form.ATU)) if form.ATU != "" else 0.0
    width = vbx.single(vbx.val(form.Hab)) if form.Hab != "" else 0.0
    length = vbx.single(vbx.val(form.Tak)) if form.Tak != "" else 0.0
    zai = _at(r, col["受注材質"])
    tyo = _at(r, col["受注調質"])
    nounyu = _at(r, col["納入先名称"])
    torihiki = _at(r, col["取引先名称"])

    flag = advance_check(form.Lot, form.DBflag, arrbox=arrbox)
    temp = packaging_specification_no(packing_no, youto, thick, width, length, zai, tyo, nounyu, torihiki,
                                      db=packing_db)
    if temp == "":
        if flag:
            out.msg_errors.append(MsgErrorCall(ADVANCE_MSG, lname, 4))
            out.sounds.append(SOUND_NEED)
        return out

    arr = temp.split("_")
    if arr[0] != "コイル":
        parts = _split(arr[-1], ",")
        n_count = len(parts) - 1
        temp = ""
        if n_count >= 2:
            for s in parts:
                if not vbx.like(temp, "*" + s + "*"):
                    temp = temp + "," + s
        elif n_count <= 1:
            if n_count == -1:
                if not quirks.reproduce("B11") and flag:
                    out.msg_errors.append(MsgErrorCall(ADVANCE_MSG, lname, 4))
                    out.sounds.append(SOUND_NEED)
                return out
            temp = parts[0]
        if flag:
            temp = temp + "," + ADVANCE_MSG
            out.sounds.append(SOUND_NEED)
        if vbx.instr(temp, "直角度の測定") > 0:
            out.sounds.append(SOUND_RIGHT_ANGLE)
        temp = ",".join(dict.fromkeys(temp.split(",")))     # Dictionary で重複を除く(入れた順)
        out.msg_errors.append(MsgErrorCall(temp, lname, 4))
    return out


def check_for_lot(rows: Sequence[Sequence[Any]], lot: str, dcaunt: int) -> bool:
    """直近 5 行に同じロットがあるか `Judg.CheckForLot` (J:2468-2488)。

    rows はシートの 1 行目から最終行(UR(ws, 1))まで。最終行から 5 行の dcaunt 列に lot を含むか(InStr)。
    """
    if lot == "":
        return False
    last = len(rows)
    start = max(1, last - 4)
    for i in range(start, last + 1):
        row = rows[i - 1]
        cell = row[dcaunt - 1] if dcaunt - 1 < len(row) else None
        if vbx.instr(vbx.cstr(cell), lot) > 0:
            return True
    return False


def new_form(**kw: Any) -> Form:
    """試験・画面用: 既定の Form を作り、属性を上書きする。"""
    f = Form()
    for k_, v_ in kw.items():
        setattr(f, k_, v_)
    return f


def clone(form: Form) -> Form:
    return copy.deepcopy(form)
