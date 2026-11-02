"""特殊検査モードの作業シート(キャスク検査・全ヵ所検査・H494検査)の共通部分

VBA では、測定値を**シートに直接入力**し(キーボードウェッジの `DT…`/`#L…` パケット・手入力)、
シートモジュールの `Worksheet_Change` / `Worksheet_SelectionChange` が判定を動かしていた。
判定の本体は `【UFKen_12】Cask` にあり、3 つのシートで共用している(シート名で分岐)。

ここでは次のものを持つ:

- `Sheet` … 作業シート 1 枚(セルの値・Q2:S4 の数式・塗り・下線・表示形式・フォントの大きさ)
- `Book`  … ブック(シートの集まり)と Excel の状態(EnableEvents・アクティブセル・Undo)、
             2 つのフォーム(`CaskForm` = UFCask、`MitForm` = UFMitsubishi)の状態
- Excel のイベントの決まり: マクロがセルに書くと(EnableEvents が真なら)そのシートの Change が
  **入れ子で**走り、`Activate` で SelectionChange が走る。VBA の動き(判定記号を書いた入れ子の
  Change で次の入力セルへ移る、など)はこの決まりの上でそのまま再現する
- `Cask` モジュールのうちシートの判定に使う手続き(AlTerRange・Special_Move・検査判定・対角・
  判定付け・ラテラルボー・歪・公差外検索・作業毎保存・MoveTg)

画面の状態(Caption/Tag/BackColor)はフォームの dataclass のフィールドで持つ。
MsgBox・WriteLog・音・フォーム表示・日報呼び出しは `Effects` に積んで返す(呼ぶべきもの)。

**On Error Resume Next**: シートの Change は `On Error Resume Next` の下で動く。呼び出した手続きの中で
実行時エラーが起きると、その手続きの残りは飛ばされ、Change の次の文から続く。ここでは
`VbaError` を投げ、Change の各文を `_resume()` で包んで同じ動きにする(`Effects.errors` に残す)。
"""
from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Callable, Iterable, Optional

from .. import flow, lines, paths, quirks, vbx

# ------------------------------------------------------------------
# シート名
# ------------------------------------------------------------------
CASK = "キャスク検査"
ZEN = "全ヵ所検査"
H494 = "H494検査"
CASK_TEMPLATE = "キャスク原紙"
ZEN_TEMPLATE = "全ヵ所検査用原紙"
H494_TEMPLATE = "H494_原紙"
MITSU_TEMPLATE = "三菱受託原紙"

RGB = tuple
WHITE = (255, 255, 255)
LIGHT_BLUE = (153, 204, 255)
GRAY = (128, 128, 128)
OK_GREEN = (0, 128, 0)
NG_RED = (200, 60, 50)
#: Excel 既定パレットの ColorIndex(この移植で使うもの)
COLOR_INDEX = {2: (255, 255, 255), 3: (255, 0, 0), 10: (0, 128, 0), 15: (192, 192, 192), 27: (255, 255, 0)}


class VbaError(Exception):
    """VBA の実行時エラー(Err.Number / Err.Description)。"""

    def __init__(self, number: int, description: str, where: str = "") -> None:
        super().__init__(f"{number}: {description}")
        self.number = number
        self.description = description
        self.where = where


def type_mismatch(where: str = "") -> VbaError:
    return VbaError(13, "型が一致しません。", where)


class ExcelError(str):
    """セルのエラー値(#VALUE! など)。"""


# ------------------------------------------------------------------
# セルの値(VBA から代入したときの Excel の変換・VBA から読んだときの型)
# ------------------------------------------------------------------
_NUM_RE = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$")
_DATE_FORMATS = ("%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d")


def to_cell(value: Any) -> Any:
    """VBA で `Range.Value = value` としたときにセルに入る値。

    - "" / None → 空(None)
    - 数として読める文字列("0000150.712" など) → 数(float)。先頭の `'` は文字列の印(値に含めない)
    - "yyyy/mm/dd h:mm" の形の文字列 → 日時
    - 数 → float、日時 → そのまま、それ以外の文字列 → 文字列
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, datetime):
        return value
    s = str(value)
    if s == "":
        return None
    if s.startswith("'"):
        return s[1:]
    t = s.strip()
    if _NUM_RE.match(t):
        return float(t)
    for f in _DATE_FORMATS:
        try:
            return datetime.strptime(t, f)
        except ValueError:
            pass
    return s


def cstr_date(d: datetime) -> str:
    """VBA `CStr(Date)`(日本の地域設定: yyyy/mm/dd h:mm:ss。時刻 0 なら日付だけ)。"""
    if d.hour == 0 and d.minute == 0 and d.second == 0:
        return d.strftime("%Y/%m/%d")
    return f"{d:%Y/%m/%d} {d.hour}:{d:%M:%S}"


def cell_str(v: Any) -> str:
    """セルの値を VBA の String に入れたとき(`Dim s As String: s = Cells(..)`)。"""
    if isinstance(v, ExcelError):
        raise type_mismatch("CStr(エラー値)")
    if isinstance(v, datetime):
        return cstr_date(v)
    return vbx.cstr(v)


def is_numeric(v: Any) -> bool:
    """VBA `IsNumeric(セル)`。空セル(Empty)は真。"""
    if v is None or isinstance(v, bool):
        return True
    if isinstance(v, ExcelError) or isinstance(v, datetime):
        return False
    if isinstance(v, (int, float)):
        return True
    return vbx.is_numeric(v)


def blank(v: Any) -> bool:
    """`Cells(..) = ""`。数の入ったセルは偽(Variant の数と文字列は等しくならない)。"""
    return v is None or v == ""


def is_text(v: Any) -> bool:
    """`VarType(セル) = vbString`。"""
    return isinstance(v, str) and not isinstance(v, ExcelError)


def val_cell(v: Any) -> float:
    """VBA `Val(セル)`(セルの値を文字列にしてから読む)。"""
    if isinstance(v, ExcelError):
        raise type_mismatch("Val(エラー値)")
    return vbx.val(cell_str(v))


def eq_single(v: Any, s: float) -> bool:
    """`Cells(..) = 単精度の変数`(Variant と数値型)。VBA は Single と Double を比べるとき Double を Single の精度に丸める
    (全ヵ所作業中保存の W3 = 10.51200008392334 が B7 = 10.512 と一致して書かれている実例と合う)。

    空セルは 0、数にならない文字列は型エラー 13(`vbx.compare_typed` の表どおり)。
    """
    if isinstance(v, ExcelError):
        raise type_mismatch("エラー値の比較")
    if isinstance(v, datetime):
        return False
    try:
        vbx.compare_typed(v, float(s), vbx.VAR, vbx.NUM)
    except vbx.TypeMismatch as e:
        raise type_mismatch(str(e)) from e
    x = 0.0 if v is None else (vbx.cdbl(v) if isinstance(v, str) else float(v))
    return vbx.single(x) == vbx.single(s)


_FMT_DATE_TOKENS = re.compile(r"yyyy|mm|dd|hh|h|nn|ss")


def fmt_date(v: Any, pattern: str) -> str:
    """VBA `Format$(日時, "yyyy/mm/dd h:mm")` など(このモジュールで使う形だけ)。

    空は ""、日時にならない文字列はそのまま返す。`h` の後ろの `mm` は分。
    """
    if v is None or v == "":
        return ""
    if isinstance(v, str):
        d = to_cell(v)
        if not isinstance(d, datetime):
            return v
        v = d
    if not isinstance(v, datetime):
        return vbx.fmt(v, pattern) if not any(t in pattern for t in ("yyyy", "dd")) else cell_str(v)
    out = []
    pos = 0
    after_h = False
    for m in _FMT_DATE_TOKENS.finditer(pattern):
        out.append(pattern[pos:m.start()])
        tok = m.group(0)
        if tok == "yyyy":
            out.append(f"{v.year:04d}")
        elif tok == "mm":
            out.append(f"{v.minute:02d}" if after_h else f"{v.month:02d}")
        elif tok == "dd":
            out.append(f"{v.day:02d}")
        elif tok == "hh":
            out.append(f"{v.hour:02d}")
            after_h = True
        elif tok == "h":
            out.append(str(v.hour))
            after_h = True
        elif tok == "nn":
            out.append(f"{v.minute:02d}")
        elif tok == "ss":
            out.append(f"{v.second:02d}")
        pos = m.end()
    out.append(pattern[pos:])
    return "".join(out)


# ------------------------------------------------------------------
# 入力パケット(補助 Maker / Conv)
# ------------------------------------------------------------------
def maker(trm: Any) -> str:
    """補助.Maker (補助:393-407)。先頭 2 文字を大文字・半角にして "DT" / "#L" / "#〆" を返す。長さ 3 未満は ""。"""
    s = "" if trm is None else str(trm)
    if s == "":
        return ""
    if len(s) >= 3:
        head = s[:2].upper()
        head = _narrow(head)
        if head in ("DT", "#L", "#〆"):
            return head
    return ""


def _narrow(s: str) -> str:
    """StrConv(…, vbNarrow) の英数記号部分(全角 ASCII → 半角)。"""
    out = []
    for ch in s:
        o = ord(ch)
        if 0xFF01 <= o <= 0xFF5E:
            out.append(chr(o - 0xFEE0))
        elif o == 0x3000:
            out.append(" ")
        else:
            out.append(ch)
    return "".join(out)


def conv(trm: Any) -> str:
    """補助.Conv (補助:380-391)。DT は `Mid(Trm, 9, 11)`(**8 文字目の符号は捨てる**)、#L/#〆 は 3 文字目から。"""
    s = "" if trm is None else str(trm)
    k = maker(s)
    if k == "DT":
        return vbx.mid(s, 9, 11)
    if k in ("#L", "#〆"):
        return vbx.mid(s, 3, len(s))
    return s


# ------------------------------------------------------------------
# 判定の関数(Cask / Judg)
# ------------------------------------------------------------------
def unit_record(n: float) -> float:
    """Judg.UnitRecord (Judg:2421-2444)。ラテラルボー値を 0.5 刻みに切り上げる。引数・戻り値は Single。

    B13: Single の誤差で 0.51・1.01… が範囲から外れて 0 になる、999 超も 0(VBA どおり)。
    直した場合は judge.unit_record_fixed(区間の間の値・0.1 未満・999 超も落とさない)。
    """
    n = vbx.single(n)
    flag = n > 0
    n = vbx.single(abs(n))
    if not quirks.reproduce("B13"):
        from ..judge import unit_record_fixed      # 通常検査と同じ直し方(区間の間の値も落とさない)
        r = vbx.single(unit_record_fixed(n))
        return r if flag else vbx.single(vbx.val_single(r) * -1)
    x = n
    big_ok = lambda v: 5.01 <= v <= 999  # noqa: E731
    table = ((0.1, 0.5, 0.5), (0.51, 1.0, 1.0), (1.01, 1.5, 1.5), (1.51, 2.0, 2.0), (2.01, 2.5, 2.5),
             (2.51, 3.0, 3.0), (3.01, 3.5, 3.5), (3.51, 4.0, 4.0), (4.01, 4.5, 4.5), (4.51, 5.0, 5.0))
    r = 0.0
    if x == 0:
        r = 0.0
    else:
        for lo, hi, out in table:
            if lo <= x <= hi:
                r = out
                break
        else:
            if big_ok(x):
                r = n
    r = vbx.single(r)
    if not flag:
        r = vbx.single(vbx.val_single(r) * -1)
    return r


def ken_judg(v: Any, kp: float, km: float) -> bool:
    """Cask.KenJudg (Cask:923-931)。下限 ≦ 値 ≦ 上限で真。数にならない値は偽。"""
    if not is_numeric(v):
        return False
    x = 0.0 if v is None else float(v)
    if x < km or x > kp:
        return False
    return x >= km and x <= kp


def tai_judg(s: Any, t: Any, kou: float) -> bool:
    """Cask.TaiJudg (Cask:932-977) の比較部分。差の絶対値を **Single** に入れて Single の公差と比べる。

    両方が数で空でないときだけ判定し、それ以外は偽。
    """
    if not (is_numeric(s) and is_numeric(t) and not blank(s) and not blank(t)):
        return False
    tai = vbx.single(abs(val_cell(s) - val_cell(t)))
    kou = vbx.single(kou)
    if kou >= tai:
        return True
    return False


def _abs_judg(v: Any, kou: str) -> bool:
    """Cask.HizumiJ / LateralJ (Cask:902-921)。公差が "0" か "" なら偽(×)。|値| ≦ 公差で真。"""
    if not is_numeric(v):
        return False
    if kou == "0" or kou == "":
        return False
    res = False
    if v is None or (not isinstance(v, str) and float(v) == 0):
        res = True
    if vbx.val(kou) >= abs(val_cell(v)):
        res = True
    elif vbx.val(kou) < abs(val_cell(v)):
        res = False
    return res


hizumi_j = _abs_judg
lateral_j = _abs_judg


def log_dir(line_name: str) -> str:
    """判定付けの Lpath (Cask:1381-1389)。ライン名は VBA の書き方(ﾊﾞﾗﾝｻｰ/ﾄｯﾄ)で比べ、それ以外は丸徳。"""
    sub = {"L-1": "L-1\\L-1_Log", "LVC": "LVC\\LVC_Log", "HVC": "HVC\\HVC_Log",
           "ﾊﾞﾗﾝｻｰ": "バランサー\\バランサー_Log", "ﾄｯﾄ": "トット\\トット_Log", "Test": "Test\\Test_Log"}
    return paths.join(paths.get("history"), sub.get(line_name, "丸徳\\丸徳_Log"))


def weight(t: Any, w: Any, l: Any) -> float:
    """OutTolerance.Weight (OutTolerance:813-817) 厚×幅×丈×2.68/10^6 を小数 3 桁で切り捨て。"""
    if not (vbx.is_numeric(t) and vbx.is_numeric(w) and vbx.is_numeric(l)):
        return 0.0
    if t == "" or w == "" or l == "":
        return 0.0
    x = float(t) * float(w) * float(l) * 2.68 / 10 ** 6
    from decimal import Decimal, ROUND_DOWN
    return float(Decimal(vbx.cstr(x)).quantize(Decimal("0.001"), rounding=ROUND_DOWN))


def weight_message(t: Any, w: Any, l: Any) -> Optional[flow.Message]:
    """UFMitsubishi の Frame2/4/6・シート発行・UserForm_Click・MLP_Click (UFMit:287-290 ほか)。"""
    tw = weight(t, w, l)
    if tw == 0:
        return None
    if 35 >= tw:
        return flow.Message("ゲージ不要 * 両面定盤歪測定必要", source="UFMit:289")
    return flow.Message("ゲージによる測定必要", source="UFMit:290")


# ------------------------------------------------------------------
# 返すもの
# ------------------------------------------------------------------
@dataclass
class LogCall:
    """WriteLog(text, b, path, fname) を呼ぶべきもの。"""
    text: str
    b: str = ""
    path: str = ""
    fname: str = ""


@dataclass
class Call:
    """他モジュールの手続きを呼ぶべきもの(Daily_Report(2)・Ainput_Startup("Cask","L2") など)。"""
    name: str
    args: tuple = ()


@dataclass
class FormShow:
    """表示するフォーム(UFOK / UFNG / UFcount(Maru|Batu, TR, TC) / UF対角(TR, TC))。"""
    name: str
    kind: str = ""
    row: int = 0
    col: int = 0


@dataclass
class Effects:
    messages: list = field(default_factory=list)     # flow.Message
    logs: list = field(default_factory=list)         # LogCall
    sounds: list = field(default_factory=list)       # "音楽合格" / "音楽不合格" / "音楽対角" / "印刷.wav" …
    forms: list = field(default_factory=list)        # FormShow
    calls: list = field(default_factory=list)        # Call
    errors: list = field(default_factory=list)       # On Error Resume Next で握りつぶされた実行時エラー
    prints: list = field(default_factory=list)       # 印刷するもの(PrintJob)
    saves: list = field(default_factory=list)        # 保存するもの(SavePlan)
    history: list = field(default_factory=list)      # `保存用` シートに足す行({列番号: 値})

    def msg(self, text: str, level: str = flow.INFO, source: str = "", title: str = "") -> None:
        self.messages.append(flow.Message(text, level, title=title, source=source))

    def extend(self, other: "Effects") -> None:
        for k in ("messages", "logs", "sounds", "forms", "calls", "errors", "prints", "saves", "history"):
            getattr(self, k).extend(getattr(other, k))


# ------------------------------------------------------------------
# シート
# ------------------------------------------------------------------
def col_letter(c: int) -> str:
    s = ""
    while c > 0:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def col_number(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def a1(addr: str) -> tuple[int, int]:
    m = re.match(r"^\$?([A-Z]+)\$?(\d+)$", addr)
    if not m:
        raise ValueError(addr)
    return int(m.group(2)), col_number(m.group(1))


_FORMULA_RE = re.compile(r"^=([A-Z]+\d+)([+-])([A-Z]+\d+)$")


class Sheet:
    """作業シート 1 枚。セルは (行, 列)、1 始まり。"""

    def __init__(self, name: str) -> None:
        self.name = name
        self.values: dict[tuple[int, int], Any] = {}
        self.formulas: dict[tuple[int, int], str] = {}
        self.interior: dict[tuple[int, int], RGB] = {}
        self.font_color: dict[tuple[int, int], RGB] = {}
        self.font_size: dict[tuple[int, int], float] = {}
        self.number_format: dict[tuple[int, int], str] = {}
        self.underline: set[tuple[int, int]] = set()
        self.visible = False
        self.locked: set[tuple[int, int]] = set()   # 保護中に利用者が書けないセル(ここで扱うものだけ)
        self.col_width: dict[int, float] = {}
        self.tab_color: Optional[int] = None

    # 値 ------------------------------------------------------------
    def get(self, r: int, c: int) -> Any:
        f = self.formulas.get((r, c))
        if f is not None:
            return self._eval(f)
        return self.values.get((r, c))

    def set(self, r: int, c: int, v: Any) -> None:
        self.formulas.pop((r, c), None)
        if v is None:
            self.values.pop((r, c), None)
        else:
            self.values[(r, c)] = v

    def set_formula(self, r: int, c: int, formula: str) -> None:
        self.values.pop((r, c), None)
        self.formulas[(r, c)] = formula

    def _eval(self, f: str) -> Any:
        """原紙・フォームが書く数式 `=B3-M2` などを Excel と同じく計算する(空は 0、文字は #VALUE!)。"""
        m = _FORMULA_RE.match(f)
        if not m:
            raise NotImplementedError(f"数式 {f} は未対応")
        x = self._operand(self.get(*a1(m.group(1))))
        y = self._operand(self.get(*a1(m.group(3))))
        if isinstance(x, ExcelError) or isinstance(y, ExcelError):
            return ExcelError("#VALUE!")
        return x + y if m.group(2) == "+" else x - y

    @staticmethod
    def _operand(v: Any) -> Any:
        if v is None:
            return 0.0
        if isinstance(v, bool):
            return 1.0 if v else 0.0
        if isinstance(v, (int, float)):
            return float(v)
        if isinstance(v, datetime):
            return (v - datetime(1899, 12, 30)).total_seconds() / 86400
        if isinstance(v, str) and vbx.is_numeric(v):
            return float(v.strip().replace(",", ""))
        return ExcelError("#VALUE!")

    def last_row(self, col: int) -> int:
        """`Cells(Rows.Count, col).End(xlUp).Row`。何も無ければ 1。"""
        rows = [r for (r, c) in list(self.values) + list(self.formulas) if c == col]
        return max(rows) if rows else 1

    # 書式 ------------------------------------------------------------
    def fill(self, r1: int, c1: int, r2: int, c2: int, color: Optional[RGB]) -> None:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if color is None:
                    self.interior.pop((r, c), None)
                else:
                    self.interior[(r, c)] = color

    def clear_underline(self, r1: int = 1, c1: int = 1, r2: int = 10 ** 6, c2: int = 10 ** 6) -> None:
        self.underline = {(r, c) for (r, c) in self.underline if not (r1 <= r <= r2 and c1 <= c <= c2)}

    def copy(self, name: Optional[str] = None) -> "Sheet":
        s = copy.deepcopy(self)
        if name:
            s.name = name
        return s

    def snapshot(self) -> dict[str, Any]:
        """A1 形式の値の辞書(保存・帳票用)。"""
        out: dict[str, Any] = {}
        keys = set(self.values) | set(self.formulas)
        for (r, c) in sorted(keys):
            out[f"{col_letter(c)}{r}"] = self.get(r, c)
        return out


def template(name: str) -> Sheet:
    """原紙シート(xlsm の `キャスク原紙`・`全ヵ所検査用原紙`・`H494_原紙`・`三菱受託原紙`)の中身。"""
    ws = Sheet(name)

    def put(cells: Iterable[tuple[str, Any]]) -> None:
        for addr, v in cells:
            ws.set(*a1(addr), v)

    common_head = [("B2", "LotNo"), ("K2", "厚公差"), ("L2", "－"), ("N2", "+"), ("C3", "×"), ("E3", "×"),
                   ("K3", "巾公差"), ("L3", "－"), ("N3", "+"), ("B4", "公差外時の処理確認なし"), ("K4", "丈公差"),
                   ("L4", "－"), ("N4", "+")]
    thick_head = [("B6", "厚1"), ("C6", "判定"), ("D6", "厚2"), ("E6", "判定"), ("F6", "厚3"), ("G6", "判定"),
                  ("H6", "厚4"), ("I6", "判定"), ("K6", "幅1"), ("L6", "判定"), ("M6", "幅2"), ("N6", "判定")]
    if name == CASK_TEMPLATE:
        put(common_head)
        put(thick_head)
        put([("R2", "～"), ("V2", "厚"), ("W2", "巾"), ("X2", "丈"), ("R3", "～"), ("U3", "MAX"), ("R4", "～"),
             ("U4", "MIN"), ("B5", "ﾌｫｰﾑ上のチェックﾎﾞｯｸｽ入で巾丈測定可能"), ("K5", "自動判定するのは、厚巾丈対角だけです"),
             ("V5", "ﾌｫｰﾑ上に対角公差入力、入力がない場合JIS4級で判定"), ("Z5", "タイピング入力できる方はこちらに入力"),
             ("O6", "丈1"), ("P6", "判定"), ("Q6", "丈2"), ("R6", "判定"), ("S6", "対角1"), ("T6", "対角2"),
             ("U6", "判定"), ("V6", "対角線差"), ("X6", "ﾗﾃﾗﾙﾎﾞｰ"), ("Y6", "幅歪"), ("Z6", "丈歪"), ("AA6", "ｹﾞｰｼﾞ歪"),
             ("AB6", "外観")])
        for i in range(1, 701):
            ws.set(i + 6, 1, float(i))
    elif name == ZEN_TEMPLATE:
        put(common_head)
        put(thick_head)
        put([("K1", "ﾗﾃﾗﾙﾎﾞｰ等級"), ("Q1", "直角度等級"), ("R2", "～"), ("W2", "厚"), ("Y2", "巾"), ("AA2", "丈"),
             ("R3", "～"), ("V3", "MAX"), ("R4", "～"), ("V4", "MIN"), ("K5", "歪公差"), ("L5", "幅"), ("N5", "丈"),
             ("W5", "ﾀｲﾋﾟﾝｸﾞ入力の事,5N時は判定あり"), ("A6", "MinMax\n指定用番号"), ("O6", "丈1"), ("P6", "判定"),
             ("Q6", "丈2"), ("R6", "判定"), ("S6", "対角1"), ("T6", "対角2"), ("U6", "対角線差"), ("V6", "判定"),
             ("W6", "ﾗﾃﾗﾙﾎﾞｰ"), ("X6", "判定"), ("Y6", "幅歪"), ("Z6", "判定"), ("AA6", "丈歪"), ("AB6", "判定"),
             ("AC6", "ｹﾞｰｼﾞ"), ("AD6", "外観"), ("AE6", "元板NO"), ("AG6", "切断ﾌﾞﾛｯｸNO"), ("AH6", "三菱"),
             ("AI6", "5N")])
        for i in range(1, 71):
            ws.set(i + 6, 1, float(i))
        for r in (7, 9, 11, 13):
            ws.set(r, col_number("AF"), "-")
            ws.set(r, col_number("AH"), "ﾌﾞﾛｯｸ上")
            ws.set(r + 1, col_number("AH"), "ﾌﾞﾛｯｸ下")
    elif name == H494_TEMPLATE:
        put(common_head)
        put(thick_head)
        put([("K1", "ﾗﾃﾗﾙﾎﾞｰ等級"), ("Q1", "対角公差"), ("R2", "~"), ("V2", "厚"), ("W2", "巾"), ("X2", "丈"),
             ("R3", "~"), ("U3", "MAX"), ("V3", 0.0), ("W3", 0.0), ("X3", 0.0), ("R4", "~"), ("U4", "MIN"),
             ("V4", 0.0), ("W4", 0.0), ("X4", 0.0),
             ("B5", "対角線差4級内＝○、対角線差3級内=△、対角線差3級外れ＝×"), ("K5", "自動判定するのは、厚巾丈対角だけです"),
             ("Y5", "ﾀｲﾋﾟﾝｸﾞ入力の事"), ("A6", "MinMax\n指定用番号"), ("O6", "幅3"), ("P6", "判定"), ("Q6", "丈1"),
             ("R6", "判定"), ("S6", "丈2"), ("T6", "判定"), ("U6", "対角1"), ("V6", "対角2"), ("W6", "判定"),
             ("X6", "対角線差"), ("Y6", "ﾗﾃﾗﾙﾎﾞｰ"), ("Z6", "定盤"), ("AA6", "ｹﾞｰｼﾞ"), ("AB6", "外観")])
        for i in range(1, 11):
            ws.set(i + 6, 1, float(i))
    elif name == MITSU_TEMPLATE:
        put([("B2", "検査番号"), ("C2", "NLM"), ("F2", "検査機器"), ("Q2", "承認"), ("R2", "確認"),
             ("B3", "スラブ番号"), ("C3", "顧客"), ("B4", "合金－質別"), ("C4", "顧客"), ("D4", "MR664-H112"),
             ("B5", "寸法"), ("C5", "NLM"), ("D5", "厚(mm)"), ("F5", "幅(mm)"), ("H5", "長さ(mm)"),
             ("B7", "原板\n積載"), ("C7", "切断\nブロック"), ("D7", "梱包番号"), ("F7", "厚さ"), ("J7", "幅"),
             ("L7", "長さ"), ("N7", "対角線差"), ("O7", "歪"), ("Q7", "外観"), ("E8", "公差"), ("J8", "-3.0\n+3.0"),
             ("L8", "-4.0\n+4.0"), ("N8", "4.0mm以下"), ("O8", "1.0mm以下/製品幅"), ("P8", "1.0mm以下/製品長さ"),
             ("E9", "原板位置"), ("F9", "#1"), ("G9", "#2"), ("H9", "#3"), ("I9", "#4"), ("J9", 1.0), ("K9", 2.0),
             ("L9", 1.0), ("M9", 2.0), ("O9", "幅"), ("P9", "長さ"), ("P19", "開始終了"), ("P20", "検査員")])
        for r in range(10, 18):
            ws.set(r, 5, "上" if r % 2 == 0 else "下")
    else:
        raise KeyError(name)
    return ws


#: Q2:S4 の数式(UFCask:473-496 / UFMit:683-706 の R1C1 を A1 に直したもの)
TOLERANCE_FORMULAS = {(2, 19): "=B3+O2", (2, 17): "=B3-M2", (3, 19): "=D3+O3", (3, 17): "=D3-M3",
                      (4, 19): "=F3+O4", (4, 17): "=F3-M4"}


# ------------------------------------------------------------------
# フォームの状態
# ------------------------------------------------------------------
@dataclass
class CaskForm:
    """UFCask の状態(コントロールの値・Caption を明示的なフィールドに)。"""
    visible: bool = False
    line_name: str = ""          # VBA の LineName(ﾊﾞﾗﾝｻｰ/ﾄｯﾄ/丸徳…)
    lot_box: str = ""
    combo: str = ""
    ken_no: str = ""
    kname: str = ""
    cno: str = ""
    mno: str = ""
    seri_no: str = ""
    min_text: str = ""           # スタートNo
    max_text: str = ""           # エンドNo
    tai: str = ""                # 対角公差
    micro: str = ""              # ゼロ校正値(Caption)
    convex: str = ""
    convex_1000: str = ""
    l1: str = ""                 # ﾏｲｸﾛｾﾞﾛ点 / ﾉｷﾞｽｾﾞﾛ点
    l2_visible: bool = False
    l3_visible: bool = False
    count_st: str = ""           # "3" / "4"(CountSt.Caption)
    inspection: str = "OFF"      # 巾丈測定
    read: str = "ON"
    cnt: str = "0"
    undo_count: str = "0"
    tag: str = ""
    hflag: str = ""
    loadtext: str = "on"
    combo_tag: str = ""
    labels: list = field(default_factory=list)   # lblCask0.. の Caption(表示中のもの)

    def __post_init__(self) -> None:
        self.line_name = lines.vba_name(self.line_name) if self.line_name else ""


@dataclass
class MitForm:
    """UFMitsubishi の状態。"""
    visible: bool = False
    line_name: str = ""
    work_flag: str = ""          # "P1"(全ヵ所) / "P_H494"
    lot_box: str = ""
    combo: str = ""
    ken_no: str = ""
    kname: str = ""
    youto_no: str = ""
    choice: str = ""             # "三菱" / "5N" / "一般"
    five_op: bool = False        # OptionButton(一般 と同じグループ)
    ippan: bool = False
    work: str = "OFF"
    inflag: str = ""
    h494: bool = False           # CheckBox「用途ｺｰﾄﾞH494」(幅 1〜299)
    min_text: str = ""
    max_text: str = ""
    inspection_method: str = ""  # Inspection1〜3 で選ばれた Caption(無ければ "")
    h494_lot: str = ""
    h494_ken_no: str = ""
    h494_kname: str = ""
    h494_youto: str = ""
    h494_min_text: str = ""
    h494_max_text: str = ""
    h494_method: str = ""        # op1〜3 の Caption
    kenh: dict = field(default_factory=lambda: {i: "" for i in range(1, 8)})   # KenHBox1〜7
    takasa_formula: str = ""     # 高さ計算式.Caption
    yama_index: Any = None       # YamaLB.ListIndex(0〜5 = 1〜6 山、6 = 枚数)
    cnt: str = "0"
    undo_count: str = "0"
    tag: str = ""
    pack_no_on_off: str = "on"
    combo_tag: str = ""
    normal_visible: bool = False
    mitsubishi_visible: bool = False
    fini_visible: bool = False

    def __post_init__(self) -> None:
        self.line_name = lines.vba_name(self.line_name) if self.line_name else ""

    def set_option(self, which: str) -> None:
        """OptionButton(FiveOP / 一般)は同じグループ: 片方を真にすると他方は偽。"""
        self.five_op = which == "5N"
        self.ippan = which == "一般"


# ------------------------------------------------------------------
# ブック(Excel の状態とイベント)
# ------------------------------------------------------------------
#: シート名 → (Change(book, sheet, r, c), SelectionChange(book, r, c))。各モードのモジュールが登録する
HANDLERS: dict[str, tuple[Callable, Callable]] = {}


class Book:
    """作業シートの入ったブックと Excel の状態。

    - `user_edit(r, c, text)` … 利用者がアクティブシートのセルに入力した(Change が走る)
    - `user_select(r, c)` … 利用者がセルを選んだ(SelectionChange が走る)
    - マクロからの書き込みは `write()`(EnableEvents が真なら入れ子の Change)、`activate()`(SelectionChange)
    """

    def __init__(self, *, now: Optional[Callable[[], datetime]] = None,
                 other_sheets: Iterable[str] = (), exists: Optional[Callable[[str], bool]] = None,
                 with_work_sheets: bool = True) -> None:
        self.sheets: dict[str, Sheet] = {}
        self.events_enabled = True
        self.active_name: Optional[str] = None
        self.active_cell: tuple[int, int] = (1, 1)
        self.cask = CaskForm()
        self.mit = MitForm()
        self.fx = Effects()
        self.now: Callable[[], datetime] = now or datetime.now
        #: 作業シート以外でブックにあるシート名(Daily_LVC・保存用・品証検証・不良要因印刷_LVC …)
        self.other_sheets: set[str] = set(other_sheets)
        #: ファイル・フォルダがあるか(保存の連番で使う)。既定は使用パスを読み替えて調べる
        self.exists: Callable[[str], bool] = exists or (lambda p: paths.to_local(p).exists())
        #: シートの内容を持つ他のシート(保存用・品証検証・作業中保存 …)
        self.extra: dict[str, Sheet] = {}
        self._undo: Optional[tuple[str, int, int, Any]] = None
        if with_work_sheets:
            for tname, wname in ((CASK_TEMPLATE, CASK), (ZEN_TEMPLATE, ZEN), (H494_TEMPLATE, H494)):
                ws = template(tname)
                ws.name = wname
                self.sheets[wname] = ws

    # シート ----------------------------------------------------------
    def has(self, name: str) -> bool:
        """補助.WsCheck"""
        return name in self.sheets or name in self.extra or name in self.other_sheets

    def ws(self, name: str) -> Sheet:
        if name in self.sheets:
            return self.sheets[name]
        if name in self.extra:
            return self.extra[name]
        raise VbaError(9, "インデックスが有効範囲にありません。", f"Sheets({name})")

    def visible(self, name: str) -> bool:
        """`ThisWorkbook.Sheets(name).Visible`(無ければ実行時エラー 9)。"""
        return self.ws(name).visible

    @property
    def active(self) -> Sheet:
        if self.active_name is None:
            raise VbaError(91, "オブジェクト変数または With ブロック変数が設定されていません。", "ActiveSheet")
        return self.ws(self.active_name)

    def activate_sheet(self, name: str) -> None:
        self.ws(name)
        self.active_name = name

    def add_sheet(self, ws: Sheet) -> None:
        """補助.AddWS + Pesute(同名は消して作り直す)。"""
        self.extra[ws.name] = ws

    def delete_sheet(self, name: str) -> None:
        self.extra.pop(name, None)
        self.sheets.pop(name, None)
        self.other_sheets.discard(name)

    # セル ------------------------------------------------------------
    def write(self, ws: Sheet, r: int, c: int, v: Any) -> None:
        """マクロからの代入。Undo の履歴は消え、EnableEvents が真ならそのシートの Change が走る。"""
        ws.set(r, c, to_cell(v))
        self._undo = None
        if self.events_enabled:
            self._fire_change(ws, r, c)

    def write_formula(self, ws: Sheet, r: int, c: int, f: str) -> None:
        ws.set_formula(r, c, f)
        self._undo = None
        if self.events_enabled:
            self._fire_change(ws, r, c)

    def activate(self, r: int, c: int) -> None:
        """`Cells(r, c).Activate`(アクティブシート)。EnableEvents が真なら SelectionChange。"""
        self.active_cell = (r, c)
        if self.events_enabled:
            h = HANDLERS.get(self.active.name)
            if h:
                h[1](self, r, c)

    def undo(self) -> None:
        """`Application.Undo`。直前が利用者の入力ならその前の値に戻す。マクロが書いた後は何もしない。"""
        if self._undo is None:
            return
        name, r, c, old = self._undo
        self._undo = None
        self.ws(name).set(r, c, old)

    def _fire_change(self, ws: Sheet, r: int, c: int) -> None:
        h = HANDLERS.get(ws.name)
        if h:
            h[0](self, ws, r, c)

    # 利用者の操作 -----------------------------------------------------
    def begin(self) -> Effects:
        self.fx = Effects()
        return self.fx

    def user_edit(self, r: int, c: int, text: Any, sheet: Optional[str] = None) -> Effects:
        """利用者がセルに入力した(`text` は打ち込んだ文字)。返り値はこの操作で起きたこと。"""
        fx = self.begin()
        if sheet:
            self.activate_sheet(sheet)
        ws = self.active
        old = ws.get(r, c)
        ws.set(r, c, to_cell(text))
        self._undo = (ws.name, r, c, old)
        self.active_cell = (r, c)
        if self.events_enabled:
            self._fire_change(ws, r, c)
        return fx

    def user_select(self, r: int, c: int, sheet: Optional[str] = None) -> Effects:
        fx = self.begin()
        if sheet:
            self.activate_sheet(sheet)
        self.activate(r, c)
        return fx

    def user_editable(self, r: int, c: int) -> bool:
        """保護中のシートで利用者が書けるセルか(キャスクの巾丈測定 OFF・3 点時の H:I)。"""
        return (r, c) not in self.active.locked

    def resume(self, fn: Callable[[], Any], where: str = "") -> Any:
        """On Error Resume Next の下で 1 文を実行する。"""
        try:
            return fn()
        except VbaError as e:
            self.fx.errors.append(f"{e.number}: {e.description} ({e.where or where})")
            return None


# ------------------------------------------------------------------
# Cask モジュールのシート処理(3 シート共用)
# ------------------------------------------------------------------
def move_tg(book: Book, n: str, sh: str) -> Optional[tuple[int, int]]:
    """Cask.MoveTg (Cask:2-40)。キャスク 7-706 行 B/D/F、全ヵ所 B/D/F/H の最初の空セル。無ければ None。

    決定 D22(W12): VBA は全ヵ所を 7-36 行(No.1-30)しか見ないが、31 枚以上もあるのでシートの全行 7-76 を見る。
    """
    ws = book.ws(sh)
    book.activate_sheet(sh)
    if n != "A":
        return None
    if ws.name == CASK:
        rows, cols = range(7, 707), range(2, 7, 2)
    elif ws.name == ZEN:
        rows, cols = range(7, 77), range(2, 9, 2)          # D22: VBA は For i = 7 To 36(Cask:25)
    else:
        return None
    for i in rows:
        for j in cols:
            if blank(ws.get(i, j)):
                return (i, j)
    return None


def al_ter_range(book: Book, r: int, c: int, befor: str = "") -> None:
    """Cask.AlTerRange (Cask:43-308)。入力パケットを見て値を変換・移送する(表は G5 §4.1)。"""
    ws = book.active
    tgt = ws.get(r, c)
    temp = maker(cell_str(tgt))
    TR, TC = r, c

    def conv_here() -> None:
        book.write(ws, TR, TC, conv(cell_str(ws.get(TR, TC))))

    def clear_here() -> None:
        book.write(ws, TR, TC, "")

    def jump(rows: range, cols: Iterable[int]) -> None:
        cols = list(cols)
        for i in rows:
            for j in cols:
                if blank(ws.get(i, j)):
                    book.undo()
                    book.activate(i, j)
                    return

    if ws.name == CASK:
        book.cask.cnt = "1"
        inrow = 7 <= TR <= 706
        if inrow and TC in (2, 4, 6, 8):
            if temp == "DT":
                conv_here()
            elif temp == "#L":
                special_move_l(book, TR, TC, befor)
            else:
                clear_here()
        elif inrow and TC in (11, 13, 15, 17, 19, 20):
            if temp == "DT":
                special_move_dt(book, TR, TC, befor)
            elif temp == "#L":
                conv_here()
            else:
                clear_here()
        elif inrow and TC in (24, 25, 26, 27, 28):
            if temp == "DT":
                special_move_dt(book, TR, TC, befor)
            elif temp == "#L":
                special_move_l(book, TR, TC, befor)
            else:
                conv_here()
        else:
            if temp == "DT":
                n = {"4": 8, "3": 6}.get(book.cask.count_st, 0)
                jump(range(7, 707), range(2, n + 1, 2))
            elif temp == "#L":
                jump(range(7, 707), range(11, 18, 2))
            else:
                conv_here()
    elif ws.name == ZEN:
        book.mit.cnt = "1"
        inrow = 7 <= TR <= 76
        if inrow and TC in (2, 4, 6, 8):
            if temp == "DT":
                conv_here()
            elif temp == "#L":
                special_move_l(book, TR, TC, befor)
            else:
                clear_here()
        elif inrow and TC in (11, 13) and book.mit.h494:
            if temp == "DT":
                conv_here()
                ws.number_format[(TR, TC)] = "0.00_);[赤](0.00)"
            else:
                special_move_l(book, TR, TC, befor)
        elif inrow and TC in (11, 13, 15, 17, 19, 20):
            if temp == "DT":
                special_move_dt(book, TR, TC, befor)
            elif temp == "#L":
                conv_here()
                ws.number_format[(TR, TC)] = "0.0_);[赤](0.0)"
        elif inrow and TC in (24, 25, 26):
            if temp == "DT":
                special_move_dt(book, TR, TC, befor)
            elif temp == "#L":
                special_move_l(book, TR, TC, befor)
            else:
                conv_here()
        elif inrow and TC == 23:
            if temp == "DT":
                special_move_dt(book, TR, TC, befor)
            elif temp == "#L":
                special_move_l(book, TR, TC, befor)
        else:
            if temp == "DT":
                jump(range(7, 77), range(2, 9, 2))
            elif temp == "#L":
                jump(range(7, 77), range(11, 18, 2))
            else:
                conv_here()
    elif ws.name == H494:
        book.mit.cnt = "1"
        inrow = 7 <= TR <= 16
        if inrow and TC in (2, 4, 6, 8):
            if temp == "DT":
                conv_here()
            elif temp == "#L":
                special_move_l(book, TR, TC, befor)
            else:
                clear_here()
        elif inrow and TC in (11, 13, 15):
            if temp == "DT":
                conv_here()
                ws.number_format[(TR, TC)] = "0.00_);[赤](0.00)"
            else:
                special_move_l(book, TR, TC, befor)
        elif inrow and TC in (17, 19, 21, 22):
            if temp == "DT":
                special_move_dt(book, TR, TC, befor)
            elif temp == "#L":
                conv_here()
                ws.number_format[(TR, TC)] = "0.0_);[赤](0.0)"
        elif inrow and TC in (26, 27):
            if temp == "DT":
                special_move_dt(book, TR, TC, befor)
            elif temp == "#L":
                special_move_l(book, TR, TC, befor)
            else:
                conv_here()
        elif inrow and TC == 25:
            if temp == "DT":
                special_move_dt(book, TR, TC, befor)
            elif temp == "#L":
                special_move_l(book, TR, TC, befor)
        else:
            if temp == "DT":
                jump(range(7, 17), range(2, 9, 2))
            elif temp == "#L":
                jump(range(7, 17), range(11, 20, 2))
            else:
                conv_here()


def _set_undo_count(book: Book, start: bool, ch_befor: str) -> None:
    """Special_Move の最初(0 にする)と最後(移送元に前の値があれば 1)の UndoCount。"""
    if start:
        if book.visible(CASK):
            book.cask.undo_count = "0"
        elif book.visible(ZEN):
            book.mit.undo_count = "0"
        return
    if book.visible(CASK):
        book.cask.undo_count = "0" if ch_befor == "" else "1"
    elif book.visible(ZEN) or book.visible(H494):
        book.mit.undo_count = "0" if ch_befor == "" else "1"


def _move_into(book: Book, ws: Sheet, i: int, j: int, TR: int, TC: int, ch_after: str, ch_befor: str,
               then: Callable[[Book, int, int], None]) -> None:
    book.write(ws, i, j, conv(ch_after))
    book.write(ws, TR, TC, ch_befor)
    book.events_enabled = True
    then(book, i, j)
    book.events_enabled = False


def special_move_dt(book: Book, TR: int, TC: int, befor: str = "") -> None:
    """Cask.Special_Move_DT (Cask:1591-1714)。DT パケットを**この行から下の**最初の空き厚セルへ移し、元のセルは前の値に戻す。"""
    ws = book.active
    book.events_enabled = False
    ch_after = cell_str(ws.get(TR, TC))
    ch_befor = befor
    _set_undo_count(book, True, ch_befor)
    found = False
    if ws.name == CASK:
        n = {"4": 8, "3": 6}.get(book.cask.count_st, 0)
        for i in range(TR, 707):
            for j in range(2, n + 1, 2):
                if blank(ws.get(i, j)):
                    _move_into(book, ws, i, j, TR, TC, ch_after, ch_befor, kensa_hantei)
                    found = True
                    break
            if found:
                break
    elif ws.name == ZEN:
        for i in range(TR, 77):
            for j in range(2, 9, 2):
                if blank(ws.get(i, j)):
                    _move_into(book, ws, i, j, TR, TC, ch_after, ch_befor, kensa_hantei)
                    found = True
                    break
            if book.mit.h494 and not found:
                for nj in (11, 13):
                    if blank(ws.get(i, nj)):
                        _move_into(book, ws, i, nj, TR, TC, ch_after, ch_befor, kensa_hantei)
                        found = True
                        break
            if found:
                break
    elif ws.name == H494:
        for i in range(TR, 17):
            for j in (2, 4, 6, 8, 11, 13, 15):
                if blank(ws.get(i, j)):
                    _move_into(book, ws, i, j, TR, TC, ch_after, ch_befor, kensa_hantei)
                    found = True
                    break
            if found:
                break
    if ws.name in (CASK, ZEN, H494) and not found:
        book.fx.msg("厚欄はすべて埋まっています", flow.ERROR, "Cask:1634")
        book.write(ws, TR, TC, ch_befor)
    _set_undo_count(book, False, ch_befor)


def special_move_l(book: Book, TR: int, TC: int, befor: str = "") -> None:
    """Cask.Special_Move_L (Cask:1715-1881)。#L パケットを最初の空き巾丈(→検査判定)・対角(→対角)セルへ移す。"""
    ws = book.active
    book.events_enabled = False
    ch_after = cell_str(ws.get(TR, TC))
    ch_befor = befor
    _set_undo_count(book, True, ch_befor)
    found = False

    def scan(last_row: int, plan: list[tuple[int, Callable]]) -> bool:
        for i in range(TR, last_row + 1):
            for j, fn in plan:
                if blank(ws.get(i, j)):
                    _move_into(book, ws, i, j, TR, TC, ch_after, ch_befor, fn)
                    return True
        return False

    if ws.name == CASK:
        if book.cask.inspection == "ON":
            found = scan(706, [(11, kensa_hantei), (13, kensa_hantei), (15, kensa_hantei), (17, kensa_hantei),
                               (19, taikaku), (20, taikaku)])
            if not found:
                book.fx.msg("巾丈対角欄はすべて埋まっています", flow.ERROR, "Cask:1771")
                book.write(ws, TR, TC, ch_befor)
        else:
            book.fx.msg("巾丈測定がONになっていません", flow.ERROR, "Cask:1776")
            book.write(ws, TR, TC, ch_befor)
    elif ws.name == ZEN:
        plan = [] if book.mit.h494 else [(11, kensa_hantei), (13, kensa_hantei)]
        plan += [(15, kensa_hantei), (17, kensa_hantei), (19, taikaku), (20, taikaku)]
        found = scan(76, plan)
        if not found:
            book.fx.msg("巾丈対角欄はすべて埋まっています", flow.ERROR, "Cask:1828")
            book.write(ws, TR, TC, ch_befor)
    elif ws.name == H494:
        found = scan(16, [(17, kensa_hantei), (19, kensa_hantei), (21, taikaku), (22, taikaku)])
        if not found:
            book.fx.msg("巾丈対角欄はすべて埋まっています", flow.ERROR, "Cask:1865")
            book.write(ws, TR, TC, ch_befor)
    _set_undo_count(book, False, ch_befor)


def _lname_judge(book: Book) -> str:
    """判定付け系の Lname (Cask:1373-1376)。全ヵ所 → キャスク の順。H494 は無い(B30)。"""
    if book.visible(ZEN):
        return book.mit.line_name
    if book.visible(CASK):
        return book.cask.line_name
    if not quirks.reproduce("B30") and book.visible(H494):
        return book.mit.line_name
    return ""


def kensa_hantei(book: Book, r: int, c: int) -> None:
    """Cask.検査判定 (Cask:1028-1115)。厚・巾・丈の判定の入口。"""
    if book.visible(ZEN):
        lname = book.mit.line_name
    elif book.visible(CASK):
        lname = book.cask.line_name
    elif book.visible(H494):
        lname = book.mit.line_name
    else:
        lname = ""
    if book.cask.visible:
        if book.cask.micro == "" or book.cask.convex == "":
            book.fx.msg("ゼロ校正を実行してください！！", flow.WARN, "Cask:1041")
            return
    ws = book.active
    ter = ws.get(r, c)
    if (not is_numeric(ws.get(3, 2))) and is_numeric(ws.get(3, 4)) and is_numeric(ws.get(3, 6)):
        return
    if not is_numeric(ter):
        return
    if book.has("Daily_" + lname):
        # WorkLot は読むだけ(Cask:1059-1063)
        book.visible(ZEN)
        book.fx.calls.append(Call("Daily_Report", (2,)))
    if is_text(ter) or blank(ter):
        return
    atu1, atu2 = val_cell(ws.get(2, 17)), val_cell(ws.get(2, 19))
    haba1, haba2 = val_cell(ws.get(3, 17)), val_cell(ws.get(3, 19))
    take1, take2 = val_cell(ws.get(4, 17)), val_cell(ws.get(4, 19))
    if ws.name == CASK:
        book.cask.tag = "off"
        if 7 <= r <= 706 and c in (2, 4, 6, 8):
            hantei_tsuke(book, ken_judg(ter, atu2, atu1), r, c)
        elif 7 <= r <= 706 and c in (11, 13):
            hantei_tsuke(book, ken_judg(ter, haba2, haba1), r, c)
        elif 7 <= r <= 706 and c in (15, 17):
            hantei_tsuke(book, ken_judg(ter, take2, take1), r, c)
        book.cask.tag = ""
    elif ws.name == ZEN:
        book.mit.tag = "off"
        if 7 <= r <= 76 and c in (2, 4, 6, 8):
            hantei_tsuke(book, ken_judg(ter, atu2, atu1), r, c)
        elif 7 <= r <= 76 and c in (11, 13):
            hantei_tsuke(book, ken_judg(ter, haba2, haba1), r, c)
        elif 7 <= r <= 76 and c in (15, 17):
            hantei_tsuke(book, ken_judg(ter, take2, take1), r, c)
        book.mit.tag = ""
    elif ws.name == H494:
        book.mit.tag = "off"
        if 7 <= r <= 16 and c in (2, 4, 6, 8):
            hantei_tsuke(book, ken_judg(ter, atu2, atu1), r, c)
        elif 7 <= r <= 16 and c in (11, 13, 15):
            hantei_tsuke(book, ken_judg(ter, haba2, haba1), r, c)
        elif 7 <= r <= 16 and c in (17, 19):
            hantei_tsuke(book, ken_judg(ter, take2, take1), r, c)
        book.mit.tag = ""


def _judge_logs(book: Book, ws: Sheet, r: int, c: int, mark: str, lname: str, word: str = "*座標:",
                with_sheet: bool = True) -> None:
    v = cell_str(ws.get(r, c))
    head = "*UFKen:" + (ws.name if with_sheet else "") + "*入力:" + v
    pos = (" " + word + "(" + str(r) + "," + str(c) + ")") if word == "*座標:" else (" " + word + str(r) + "," + str(c))
    book.fx.logs.append(LogCall(head + pos + "：" + mark, "", ""))
    book.fx.logs.append(LogCall(head + pos + "：" + mark + ":" + lname, "", log_dir(lname), "Measured.log"))


def _ufcount(book: Book, r: int, c: int, kind: str) -> None:
    """Cask.UFcount起動 (Cask:1882-1898)。"""
    ws = book.active
    if 7 <= r <= 706 and c in (2, 4, 6, 8, 11, 13, 15, 17, 19, 20) and not blank(ws.get(r, c)):
        book.fx.forms.append(FormShow("UFcount", kind, r, c))


def hantei_tsuke(book: Book, judg: bool, r: int, c: int) -> None:
    """Cask.判定付け (Cask:1365-1472)。判定列(TC+1)に ○/× と色(ColorIndex 10/3)、ログ、音、次の入力セルの色。"""
    ws = book.active
    if not is_numeric(ws.get(r, c)):
        return
    lname = _lname_judge(book)
    # UndoCount(移送のあとの元セルの再判定を 1 回止める)
    if book.visible(CASK):
        if book.cask.undo_count == "1":
            book.cask.undo_count = "0"
            return
    elif book.visible(ZEN) or book.visible(H494):
        if book.mit.undo_count == "1":
            book.mit.undo_count = "0"
            return
    mark = "○" if judg else "×"
    book.write(ws, r, c + 1, mark)
    ws.interior[(r, c + 1)] = COLOR_INDEX[10 if judg else 3]
    _judge_logs(book, ws, r, c, mark, lname)
    book.fx.sounds.append("音楽合格" if judg else "音楽不合格")
    if c in (2, 4, 6, 11, 13, 15, 17, 20):
        ws.interior[(r, c + 2)] = LIGHT_BLUE
    elif c == 8:
        ws.interior[(r, c + 3)] = LIGHT_BLUE
    if book.visible(CASK):
        _ufcount(book, r, c, "Maru" if judg else "Batu")
    else:
        book.fx.forms.append(FormShow("UFOK" if judg else "UFNG"))


def taikaku(book: Book, r: int, c: int) -> None:
    """Cask.対角 (Cask:1117-1235)。全ヵ所・H494 の公差は S1(Single)、キャスクはフォームの対角公差。"""
    ws = book.active
    zen = {"P1": ZEN, "P_H494": H494}.get(book.mit.work_flag)
    kou = 0.0
    if not book.cask.visible:
        if zen is None:
            raise VbaError(91, "オブジェクト変数または With ブロック変数が設定されていません。", "Cask:1202")
        kou = vbx.single(val_cell(book.ws(zen).get(1, 19)))
    if ws.name == CASK and c == 20:
        if book.cask.tai == "":
            book.fx.msg("対角公差が入力されていません！！", flow.WARN, "Cask:1210")
            return
        if book.cask.inspection == "ON" and book.cask.tai != "" and vbx.is_numeric(book.cask.tai):
            kou = vbx.single(vbx.val(book.cask.tai))
        taikaku_tsuke(book, tai_judg(ws.get(r, 19), ws.get(r, 20), kou), r, c, kou)
    if ws.name == ZEN and c == 20:
        taikaku_tsuke(book, tai_judg(ws.get(r, 19), ws.get(r, 20), kou), r, c, kou)
    if ws.name == H494 and c == 22:
        taikaku_tsuke(book, tai_judg(ws.get(r, 21), ws.get(r, 22), kou), r, c, kou)


def _diff_text(a: Any, b: Any) -> float:
    """`Abs(Format(val(a) - val(b), "0.0"))`"""
    return abs(float(vbx.fmt(val_cell(a) - val_cell(b), "0.0")))


def taikaku_tsuke(book: Book, judg: bool, r: int, c: int, kou: float) -> None:
    """Cask.対角判定付け (Cask:1473-1590)。"""
    ws = book.active
    lname = _lname_judge(book)
    if ws.name == CASK:
        s, t = ws.get(r, 19), ws.get(r, 20)
        if is_numeric(s) and is_numeric(t) and not blank(s) and not blank(t):
            book.write(ws, r, 22, _diff_text(s, t))
            if judg:
                book.write(ws, r, 21, "○")
                ws.interior[(r, 21)] = LIGHT_BLUE
                _judge_logs(book, ws, r, c, "○", lname, word="*位置:")
                book.fx.sounds.append("音楽合格")
                if book.visible(CASK):
                    book.fx.forms.append(FormShow("UFcount.Maru.Visible"))
                book.write(ws, r, 23, "")
            else:
                book.write(ws, r, 21, "×")
                ws.interior[(r, 21)] = NG_RED
                _judge_logs(book, ws, r, c, "×", lname, word="*位置:", with_sheet=False)
                book.fx.sounds.append("音楽不合格")
                if book.visible(CASK):
                    book.fx.forms.append(FormShow("UFcount.Batu.Visible"))
                book.write(ws, r, 23, "")
                if book.cask.tai != "":
                    book.write(ws, r, 23, "公差:[" + vbx.fmt(book.cask.tai, "0.0") + "]")
                else:
                    book.write(ws, r, 23, "4級公差:[" + vbx.fmt(kou, "0.0") + "]")
                ws.font_size[(r, 23)] = 8
    elif ws.name == ZEN:
        s, t = ws.get(r, 19), ws.get(r, 20)
        book.write(ws, r, 21, _diff_text(s, t))
        s, t = ws.get(r, 19), ws.get(r, 20)
        if is_numeric(s) and is_numeric(t) and not blank(s) and not blank(t):
            if judg:
                book.write(ws, r, 22, "○")
                ws.interior[(r, 22)] = OK_GREEN
                _judge_logs(book, ws, r, c, "○", lname, word="*位置:", with_sheet=False)
                book.fx.sounds.append("音楽合格")
                book.fx.forms.append(FormShow("UFOK"))
            else:
                ws.interior[(r, 22)] = NG_RED
                book.write(ws, r, 22, "×")
                _judge_logs(book, ws, r, c, "×", lname, word="*位置:", with_sheet=False)
                book.fx.sounds.append("音楽不合格")
                book.fx.forms.append(FormShow("UFNG"))
    elif ws.name == H494:
        u, v = ws.get(r, 21), ws.get(r, 22)
        book.write(ws, r, 24, _diff_text(v, u))
        u, v = ws.get(r, 21), ws.get(r, 22)
        if is_numeric(u) and is_numeric(v) and not blank(u) and not blank(v):
            if judg:
                book.write(ws, r, 23, "○")
                ws.interior[(r, 23)] = LIGHT_BLUE
                _judge_logs(book, ws, r, c, "○", lname, word="*位置:", with_sheet=False)
                book.fx.sounds.append("音楽合格")
                book.fx.forms.append(FormShow("UFOK"))
            else:
                ws.interior[(r, 23)] = NG_RED
                book.write(ws, r, 23, "×")
                _judge_logs(book, ws, r, c, "×", lname, word="*位置:", with_sheet=False)
                book.fx.sounds.append("音楽不合格")
                book.fx.forms.append(FormShow("UFNG"))


def lateralbow_z(book: Book, r: int, c: int) -> None:
    """Cask.Lateralbow_Z (Cask:2040-2076)。全ヵ所 W 列(7-76 行)、H494 Y 列(**7-10 行**)。公差は M1(Val を String に)。"""
    ws = book.active
    book.visible(ZEN)
    if (not is_numeric(ws.get(3, 2))) and is_numeric(ws.get(3, 4)) and is_numeric(ws.get(3, 6)):
        return
    ter = ws.get(r, c)
    if not is_numeric(ter):
        return
    if is_text(ter) or blank(ter):
        return
    kou = vbx.cstr(val_cell(ws.get(1, 13)))
    if ws.name == ZEN:
        book.mit.tag = "off"
        if 7 <= r <= 76 and c == 23:
            lateralbow_tsuke(book, lateral_j(ter, kou), r, c)
        book.mit.tag = ""
    elif ws.name == H494:
        book.mit.tag = "off"
        if 7 <= r <= 10 and c == 25:
            lateralbow_tsuke(book, lateral_j(ter, kou), r, c)
        book.mit.tag = ""


def lateralbow_tsuke(book: Book, judg: bool, r: int, c: int) -> None:
    """Cask.Lateralbow_判定付け (Cask:2077-2158)。全ヵ所は X 列に ○/×、キャスク・H494 はセル自体を塗る。"""
    ws = book.active
    if not is_numeric(ws.get(r, c)):
        return
    lname = _lname_judge(book)
    if book.visible(ZEN):
        if book.mit.undo_count == "1":
            book.mit.undo_count = "0"
            return
    mark = "○" if judg else "×"
    if ws.name == ZEN:
        book.write(ws, r, c + 1, mark)
        ws.interior[(r, c + 1)] = COLOR_INDEX[10 if judg else 3]
    elif ws.name in (CASK, H494):
        ws.interior[(r, c)] = COLOR_INDEX[10] if judg else NG_RED
    _judge_logs(book, ws, r, c, mark, lname)
    book.fx.sounds.append("音楽合格" if judg else "音楽不合格")
    if book.visible(CASK):
        _ufcount(book, r, c, "Maru" if judg else "Batu")
    else:
        book.fx.forms.append(FormShow("UFOK" if judg else "UFNG"))


def hizumi_z(book: Book, r: int, c: int) -> None:
    """Cask.Hizumi_Z (Cask:1994-2039)。5N(FiveOP)は幅 2.0 / 丈 3.0 固定、一般は M5 / O5。三菱(どちらも偽)は判定しない。"""
    ws = book.active
    sh: Optional[Sheet] = None
    if book.visible(ZEN):
        sh = book.ws(ZEN)
    elif book.visible(CASK):
        pass
    ter = ws.get(r, c)
    if not is_numeric(ter):
        return
    if is_text(ter) or blank(ter):
        return
    book.mit.tag = "off"
    if book.mit.five_op:
        if ws.name == ZEN:
            if 7 <= r <= 76 and c == 25:
                hizumi_tsuke(book, hizumi_j(ter, vbx.cstr(2.0)), r, c)
            elif 7 <= r <= 76 and c == 27:
                hizumi_tsuke(book, hizumi_j(ter, vbx.cstr(3.0)), r, c)
    elif book.mit.ippan:
        if sh is None:
            raise VbaError(91, "オブジェクト変数または With ブロック変数が設定されていません。", "Cask:2026")
        set_h = cell_str(sh.get(5, 13))
        set_t = cell_str(sh.get(5, 15))
        if ws.name == ZEN:
            if 7 <= r <= 76 and c == 25:
                hizumi_tsuke(book, hizumi_j(ter, set_h), r, c)
            elif 7 <= r <= 76 and c == 27:
                hizumi_tsuke(book, hizumi_j(ter, set_t), r, c)
    book.mit.tag = ""


def hizumi_tsuke(book: Book, judg: bool, r: int, c: int) -> None:
    """Cask.Hizumi_判定付け (Cask:2159-2239)。全ヵ所の Z/AB 列に ○/×(ColorIndex 10/3)。"""
    ws = book.active
    if not is_numeric(ws.get(r, c)):
        return
    lname = _lname_judge(book)
    if book.visible(ZEN):
        if book.mit.undo_count == "1":
            book.mit.undo_count = "0"
            return
    if ws.name != ZEN:
        return
    mark = "○" if judg else "×"
    book.write(ws, r, c + 1, mark)
    ws.interior[(r, c + 1)] = COLOR_INDEX[10 if judg else 3]
    _judge_logs(book, ws, r, c, mark, lname)
    book.fx.sounds.append("音楽合格" if judg else "音楽不合格")
    if book.visible(CASK):
        _ufcount(book, r, c, "Maru" if judg else "Batu")
    else:
        book.fx.forms.append(FormShow("UFOK" if judg else "UFNG"))


WORK_SAVE_SHEET = {ZEN: "全ヵ所作業中保存", CASK: "ｷｬｽｸ作業中保存", H494: "H494作業中保存"}


def koshagai_kensaku(book: Book, r: int, c: int) -> None:
    """Cask.公差外検索 (Cask:1335-1364)。名前と違い、決まった列の入力で作業中保存をするだけ。"""
    ws = book.active
    if ws.name == CASK:
        st = book.cask.count_st
        if (c in (2, 6) and st == "3") or (c in (4, 8) and st == "4") or c in (13, 17, 20, 26):
            sagyo_hozon(book)
    elif ws.name == ZEN:
        if c in (4, 8, 13, 17, 20, 23):
            sagyo_hozon(book)
    elif ws.name == H494:
        if c in (4, 8, 15, 19, 22, 28):
            sagyo_hozon(book)


def sagyo_hozon(book: Book) -> None:
    """Cask.作業毎保存 (Cask:1919-1937)。作業中保存シートへ全セルを写す(AddWS + Pesute。塗りは消える)。"""
    ws = book.active
    name = WORK_SAVE_SHEET.get(ws.name)
    if not name:
        return
    cp = ws.copy(name)
    cp.interior = {}
    cp.font_color = {}
    cp.visible = False
    book.add_sheet(cp)
    book.fx.calls.append(Call("作業毎保存", (name,)))


def uf_taikaku(book: Book, r: int, c: int) -> None:
    """Cask.UF対角起動 (Cask:1899-1918)。全ヵ所は**シート名の誤記("全ヵ所検査検査")で出ない**。"""
    ws = book.active
    if ws.name == CASK and c == 19 and not blank(ws.get(r, 19)):
        book.fx.forms.append(FormShow("UF対角", "", r, c))
    elif ws.name == "全ヵ所検査検査" and c == 19 and not blank(ws.get(r, 19)):
        book.fx.forms.append(FormShow("UF対角", "", r, c))
    elif ws.name == H494 and c == 21 and not blank(ws.get(r, 21)):
        book.fx.forms.append(FormShow("UF対角", "", r, c))


def take_befor(book: Book, form: Any, r: int, c: int) -> str:
    """各シートの Change の冒頭(KySh:19-27 ほか)。Cnt が "0" のときだけ Undo で変更前の値を取る。"""
    if form.cnt != "0":
        return ""
    ws = book.active
    after = cell_str(ws.get(r, c))
    book.events_enabled = False
    book.undo()
    befor = cell_str(ws.get(r, c))
    book.write(ws, r, c, after)
    book.events_enabled = True
    return befor


# ------------------------------------------------------------------
# 印刷・保存(何を・どこに)
# ------------------------------------------------------------------
@dataclass
class PrintJob:
    """印刷するシートと設定(PageSetup)。帳票は D7 で HTML→PDF にする。"""
    sheet: str
    print_area: str = ""
    left_header: str = ""
    center_header: str = ""
    right_header: str = ""
    left_footer: str = ""
    center_footer: str = ""
    right_footer: str = ""
    paper: str = "A4"
    orientation: str = "landscape"
    fit_wide: Optional[int] = 1
    fit_tall: Optional[int] = None
    margins: dict = field(default_factory=dict)      # inch
    center_h: bool = True
    center_v: bool = False
    snapshot: dict = field(default_factory=dict)     # 印刷時のセルの値(A1 → 値)


@dataclass
class SaveStep:
    kind: str            # "save"(新しいブックとして保存) / "copy"(FileCopy) / "mkdir"
    path: str
    src: str = ""
    password: str = ""
    readonly: bool = False


@dataclass
class SavePlan:
    """保存の手順(VBA の xlsx 保存を D7 で PDF+CSV に置き換えるための材料)。"""
    name: str                                         # 手続き名(キャスク専用保存 など)
    sheets: dict = field(default_factory=dict)        # 保存するシート名 → セルの値(A1 → 値)
    steps: list = field(default_factory=list)         # SaveStep
    error: str = ""                                   # 途中で止まったとき(VBA の実行時エラー・End)


def numbered_path(book: Book, set_path: str, day_lot: str) -> tuple[Optional[str], str]:
    """`{SetPath}\\{DayLot}.xlsx`、あれば `.002`〜`.999`(Cask:625-641 ほか)。999 を超えたら (None, 文言)。"""
    first = paths.join(set_path, day_lot + ".xlsx")
    if not book.exists(first):
        return first, ""
    for n in range(2, 1000):
        p = paths.join(set_path, day_lot + "." + vbx.right("00" + str(n), 3) + ".xlsx")
        if not book.exists(p):
            return p, ""
        if n == 999:
            return None, "連番をこれ以上付けることが出来ません。"
    return None, "連番をこれ以上付けることが出来ません。"


def history_row(book: Book, ws: Sheet, kind: str, *, cask_bug: bool = False) -> Optional[dict[int, Any]]:
    """`保存用` シートの最終行の次に書く履歴行 (UFCask:148-162 / UFMit:1576-1590 ほか)。列番号 → 値。

    A=開始～印刷日時、B=ロット、C=厚、D=幅、E=丈、BY(77)="別ファイル参照"、CE(83)=種別。
    `cask_bug`: キャスク(UFCask:156-157)は厚を D 列に書いてから幅で上書きする(B29)。
    """
    if not book.has("保存用"):
        return None
    row: dict[int, Any] = {2: ws.get(2, 4), 83: kind}
    if cask_bug and quirks.reproduce("B29"):
        row[4] = ws.get(3, 2)
        row[4] = ws.get(3, 4)
    else:
        row[3] = ws.get(3, 2)
        row[4] = ws.get(3, 4)
    row[5] = ws.get(3, 6)
    row[1] = fmt_date(ws.get(1, 2), "yyyy/mm/dd  h:mm") + "～" + fmt_date(book.now(), "yyyy/mm/dd  h:mm")
    row[77] = "別ファイル参照"
    return row
