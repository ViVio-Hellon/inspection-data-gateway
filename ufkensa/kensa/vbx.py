"""VBA の数値・文字列の挙動を再現する(UF検査判定の判定・計算はすべてこれを通す)

判定の結果を VBA と一致させるための道具です。**ここにある癖は直さない**こと。
直したくなったら `quirks.py` のスイッチにして、選んでもらいます(判断事項 D3)。

【VBA の性質(docs/analysis/G2 §0)】
V1 DB の値はすべて文字列(GetString)。Null は ""            → datasource が文字列で返す
V3 Single(単精度)の誤差                                       → `single()`
V4 Val(Single) は CStr を通る(有効 7 桁)                      → `val_single()`
V5 CStr(Double) は有効 15 桁                                  → `cstr()`
V6 Format は 0 から遠い方へ四捨五入、Round は銀行丸め         → `fmt()` / `vround()`
V7 文字列と数値の比較(数にならない文字列は型エラー 13)        → `vcompare()` が `TypeMismatch` を投げる
V9 Long への代入は銀行丸め                                    → `clng()`
V10 -0                                                         → `fmt()` は "-0.0" を出さない

`Format` の .5 境界の丸め(10 進で丸めるか 2 進のまま丸めるか)は実機確認待ち(U5)。
ここでは「値を有効 15 桁の 10 進に直してから四捨五入」で実装し、`_FORMAT_FROM_15_DIGITS` で切り替えられる。
"""
from __future__ import annotations

import math
import re
import struct
from decimal import ROUND_HALF_EVEN, ROUND_HALF_UP, Decimal
from typing import Any, Optional, Union

Number = Union[int, float]


#: U1(実機確認待ち)の扱いを 1 か所で決めるスイッチ。対象は「コントロールの値・Variant・String の
#: **数にならない文字列(空文字を含む)** と数との比較」(例 `UF1.Yama = 3` の Yama が ""、`Select Case Rig` の
#: Rig が "S"、`.ZS = 0` の ZS が "")。
#:   False(既定)… エラーにしない。Variant の規則どおり「数 < 文字列」で比べる(= は偽)。
#:                 本番で動いている箇所が多数ある(空欄のまま数と比べても止まっていない)ため、こちらを既定にした
#:   True       … VBA の比較演算子の表どおり型エラー 13(`TypeMismatch`)
#: 見る所: grid.variant_cmp、quality.vcmp、checks._num_eq / _case_between(`u1_compare` を通す)。
U1_TEXT_VS_NUMBER_RAISES = False


class TypeMismatch(ValueError):
    """VBA の実行時エラー 13(型が一致しません)。呼び手が VBA の On Error に合わせて扱う。"""


# ------------------------------------------------------------------
# 単精度
# ------------------------------------------------------------------
def single(x: Number) -> float:
    """VBA の Single に入れた値(float32 に丸めて float に戻す)。"""
    return struct.unpack("<f", struct.pack("<f", float(x)))[0]


def _digits_text(x: float, digits: int) -> str:
    """有効 `digits` 桁の 10 進文字列(VBA の CStr の書き方。末尾の 0 と小数点は付けない)。"""
    if x == 0:
        return "0"
    if math.isnan(x) or math.isinf(x):
        raise OverflowError(x)
    s = f"{x:.{digits}g}"
    if "e" in s:
        mant, exp = s.split("e")
        e = int(exp)
        # VBA は |x| < 1E-15 くらいから指数表記(Double)。それ以外は普通の小数で書く
        if -15 <= e < digits:
            d = Decimal(s)
            s = format(d, "f")
            if "." in s:
                s = s.rstrip("0").rstrip(".")
        else:
            mant = mant.rstrip("0").rstrip(".") if "." in mant else mant
            s = f"{mant}E{'+' if e >= 0 else '-'}{abs(e):02d}"
    if s.startswith("0.") or s.startswith("-0."):
        pass
    return s


def cstr(x: Any, *, is_single: bool = False) -> str:
    """VBA `CStr`。数は有効 15 桁(Single は 7 桁)、文字列はそのまま、None(Empty)は ""。"""
    if x is None:
        return ""
    if isinstance(x, bool):
        return "True" if x else "False"
    if isinstance(x, int):
        return str(x)
    if isinstance(x, float):
        if x == 0:
            return "0"
        return _digits_text(single(x) if is_single else x, 7 if is_single else 15)
    return str(x)


# ------------------------------------------------------------------
# Val / CDbl / CLng
# ------------------------------------------------------------------
_VAL_RE = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eEdD][+-]?\d+)?")


def val(x: Any) -> float:
    """VBA `Val`。先頭の数として読める所だけ読む。空白・タブ・改行はどこにあっても無視。

    数(int/float)を渡したときは VBA と同じく **CStr を通してから**読む(Double は有効 15 桁)。
    Single の値を渡すときは `val_single` を使う。
    """
    if x is None:
        return 0.0
    if isinstance(x, bool):
        return -1.0 if x else 0.0
    if isinstance(x, (int, float)):
        x = cstr(x)
    s = re.sub(r"[ \t\r\n]", "", str(x))
    if s[:2].lower() == "&h":
        try:
            return float(int(re.match(r"[0-9a-fA-F]*", s[2:]).group(0) or "0", 16))
        except ValueError:
            return 0.0
    m = _VAL_RE.match(s)
    if not m:
        return 0.0
    text = m.group(0).replace("d", "e").replace("D", "e")
    try:
        return float(text)
    except ValueError:
        return 0.0


def val_single(x: float) -> float:
    """`Val(単精度の変数)`。CStr(Single) は有効 7 桁なので、7 桁に丸めた 10 進を読む(V4)。"""
    return val(cstr(single(x), is_single=True))


def is_numeric(x: Any) -> bool:
    """VBA `IsNumeric`(よく使う範囲)。空文字は偽、前後の空白は許す、カンマ区切りも数。"""
    if x is None:
        return False
    if isinstance(x, bool):
        return True
    if isinstance(x, (int, float)):
        return True
    s = str(x).strip()
    if not s:
        return False
    s = s.replace(",", "")
    return bool(re.fullmatch(r"[+-]?(\d+\.?\d*|\.\d+)([eEdD][+-]?\d+)?", s))


def cdbl(x: Any) -> float:
    """VBA `CDbl`。数にならない文字列(空文字を含む)は型エラー 13。Empty(None)は 0。"""
    if x is None:
        return 0.0
    if isinstance(x, bool):
        return -1.0 if x else 0.0
    if isinstance(x, (int, float)):
        return float(x)
    if not is_numeric(x):
        raise TypeMismatch(f"CDbl({x!r})")
    return float(str(x).strip().replace(",", "").replace("d", "e").replace("D", "e"))


def _half_even(x: float) -> int:
    d = Decimal(repr(float(x))).quantize(Decimal(1), rounding=ROUND_HALF_EVEN)
    return int(d)


def clng(x: Any) -> int:
    """VBA `CLng` / Long への代入。銀行丸め(2.5→2、3.5→4)。数にならない文字列は型エラー。"""
    return _half_even(cdbl(x))


def cint(x: Any) -> int:
    """VBA `CInt`(範囲チェックは省略。銀行丸めは CLng と同じ)。"""
    return clng(x)


def fix(x: Number) -> int:
    """VBA `Fix`(0 の方向へ切り捨て)。"""
    return int(math.trunc(float(x)))


def vint(x: Number) -> int:
    """VBA `Int`(小さい方へ切り捨て)。"""
    return int(math.floor(float(x)))


def vround(x: Any, digits: int = 0) -> float:
    """VBA `Round`。銀行丸め。

    VBA は値を 10 進で見て .5 ちょうどを偶数へ寄せる(例 Round(0.125, 2) = 0.12、Round(2.5) = 2)。
    2 進の誤差で .5 の手前になっている値(2.675 など)の扱いは U5 の実機確認待ち。
    ここでは有効 15 桁の 10 進に直してから丸める(VBA の観測例 Round(2.675, 2) = 2.68 と合う)。
    """
    v = cdbl(x)
    d = Decimal(cstr(v)) if v != 0 else Decimal(0)
    q = Decimal(1).scaleb(-digits)
    r = d.quantize(q, rounding=ROUND_HALF_EVEN)
    return float(r)


# ------------------------------------------------------------------
# Format
# ------------------------------------------------------------------
_FORMAT_FROM_15_DIGITS = True     # U5: 実機で 2 進のまま丸めると分かったら False にする
_FMT_RE = re.compile(r"^(#*0*)(?:\.(0+|#+))?$")


def _decimal_of(v: float) -> Decimal:
    if _FORMAT_FROM_15_DIGITS:
        return Decimal(cstr(v)) if v != 0 else Decimal(0)
    return Decimal(v)


def fmt(x: Any, pattern: str) -> str:
    """VBA `Format(x, "0.000")` などの数の書式。0 から遠い方へ四捨五入(V6)。

    対応する書式: `0`、`0.0`、`0.00`、`0.000`、`0.0000`、`#`、`#.#` などの単純な形、`#,##0`。
    - 空文字・None は ""(VBA の Format("", "0.000") は "")
    - 数にならない文字列はそのまま返す(VBA と同じ)
    - -0 は "0.0" のように符号なし(V10)
    """
    if x is None:
        return ""
    if isinstance(x, str):
        if x == "":
            return ""
        if not is_numeric(x):
            return x
        v = cdbl(x)
    elif isinstance(x, bool):
        v = -1.0 if x else 0.0
    else:
        v = float(x)
    comma = "," in pattern
    p = pattern.replace(",", "")
    m = _FMT_RE.match(p)
    if not m:
        raise NotImplementedError(f"Format の書式 {pattern!r} は未対応")
    int_part, dec_part = m.group(1), m.group(2) or ""
    places = len(dec_part)
    d = _decimal_of(v)
    q = Decimal(1).scaleb(-places)
    r = d.quantize(q, rounding=ROUND_HALF_UP)
    if r == 0:
        r = abs(r)
    s = format(r, f",.{places}f" if comma else f".{places}f")
    if dec_part.startswith("#"):
        if "." in s:
            s = s.rstrip("0").rstrip(".")
    if int_part and set(int_part) == {"#"} and s.startswith("0") and (len(s) == 1 or s[1] == "."):
        s = s[1:] if s != "0" else ""
    min_int = int_part.count("0")
    if min_int > 1:
        neg = s.startswith("-")
        body = s[1:] if neg else s
        head, _, tail = body.partition(".")
        head = head.zfill(min_int)
        body = head + ("." + tail if tail else "")
        s = ("-" if neg else "") + body
    return s


# ------------------------------------------------------------------
# 比較(Variant)
# ------------------------------------------------------------------
def _is_num(x: Any) -> bool:
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def vcompare(a: Any, b: Any) -> int:
    """VBA の Variant 比較。-1 / 0 / 1 を返す。

    - 数と数 → 数で比べる
    - 文字列と文字列 → バイナリ比較(Option Compare 指定なし)
    - 数と文字列 → 文字列を数に直して比べる。直せなければ型エラー 13(V7。実機確認 U1)
    - Empty(None)は 0 または "" として扱う
    """
    if a is None:
        a = 0 if _is_num(b) else ""
    if b is None:
        b = 0 if _is_num(a) else ""
    if _is_num(a) and _is_num(b):
        return (a > b) - (a < b)
    if isinstance(a, str) and isinstance(b, str):
        return (a > b) - (a < b)
    if _is_num(a) and isinstance(b, str):
        if not is_numeric(b):
            raise TypeMismatch(f"{a!r} と {b!r} の比較")
        bb = cdbl(b)
        return (a > bb) - (a < bb)
    if isinstance(a, str) and _is_num(b):
        return -vcompare(b, a)
    raise TypeMismatch(f"{a!r} と {b!r} の比較")


def veq(a: Any, b: Any) -> bool:
    return vcompare(a, b) == 0


def u1_compare(a: Any, b: Any, raises: Optional[bool] = None) -> int:
    """`vcompare` に U1 の扱いを足したもの。数と「数にならない文字列」の比較は
    `U1_TEXT_VS_NUMBER_RAISES`(`raises` で上書き可)が偽なら「数 < 文字列」、真なら `TypeMismatch`。"""
    try:
        return vcompare(a, b)
    except TypeMismatch:
        if U1_TEXT_VS_NUMBER_RAISES if raises is None else raises:
            raise
        if _is_num(a) and isinstance(b, str):
            return -1
        if isinstance(a, str) and _is_num(b):
            return 1
        raise


def text_eq(a: Any, b: Any) -> bool:
    """`StrComp(a, b, vbTextCompare) = 0`。大文字小文字・全角半角を区別しない比較。"""
    import unicodedata
    na = unicodedata.normalize("NFKC", "" if a is None else str(a)).casefold()
    nb = unicodedata.normalize("NFKC", "" if b is None else str(b)).casefold()
    return na == nb


def like(text: Any, pattern: str) -> bool:
    """VBA `Like`(バイナリ比較)。`*` `?` `#` `[...]` に対応。"""
    s = "" if text is None else str(text)
    out = []
    i = 0
    while i < len(pattern):
        c = pattern[i]
        if c == "*":
            out.append(".*")
        elif c == "?":
            out.append(".")
        elif c == "#":
            out.append(r"\d")
        elif c == "[":
            j = pattern.find("]", i + 1)
            if j < 0:
                out.append(re.escape(c))
            else:
                body = pattern[i + 1:j]
                if body.startswith("!"):
                    body = "^" + body[1:]
                out.append("[" + body.replace("\\", "\\\\") + "]")
                i = j
        else:
            out.append(re.escape(c))
        i += 1
    return re.fullmatch("".join(out), s, flags=re.DOTALL) is not None


# ------------------------------------------------------------------
# 文字列
# ------------------------------------------------------------------
def left(s: Any, n: int) -> str:
    return ("" if s is None else str(s))[:max(n, 0)]


def right(s: Any, n: int) -> str:
    t = "" if s is None else str(s)
    return t[-n:] if n > 0 else ""


def mid(s: Any, start: int, length: Optional[int] = None) -> str:
    """VBA `Mid(s, start[, length])`。start は 1 始まり。"""
    t = "" if s is None else str(s)
    if start < 1:
        raise ValueError("Mid の開始位置は 1 以上")
    return t[start - 1:] if length is None else t[start - 1:start - 1 + length]


def instr(s: Any, sub: str, start: int = 1) -> int:
    """VBA `InStr(start, s, sub)`(バイナリ比較)。見つからなければ 0。"""
    t = "" if s is None else str(s)
    if t == "" or start > len(t):          # 調べる文字列が空・開始位置が後ろ → 0(探す文字列が空でも)
        return 0
    if sub == "":
        return start
    i = t.find(sub, start - 1)
    return i + 1


# ------------------------------------------------------------------
# 型つきの比較(VBA の比較演算子の表どおり。U1)
# ------------------------------------------------------------------
NUM, STR, VAR = "num", "str", "var"


def compare_typed(a: Any, b: Any, a_type: str = VAR, b_type: str = VAR) -> int:
    """VBA の `a < b` などの比較を、**両側の宣言の型**まで含めて再現する。-1 / 0 / 1 を返す。

    型の指定(VBA で何として書かれているか):
        NUM … 数値型の変数・数値の定数(`Dim x As Double`、`0`、`74#`)
        STR … String 型(`Dim s As String`、`.Caption`、`.Text`)
        VAR … Variant(`.Value` の既定・Variant 変数・配列の要素・`Cells().Value`)。中身は Python の型で見る

    VBA の表(比較演算子のヘルプ):
    - 数値型 と 数値型                        → 数で比べる
    - String と String                        → 文字で比べる(バイナリ)
    - 数値型 と 数になる Variant              → 数で比べる
    - 数値型 と 数にならない文字列の Variant  → **型エラー 13**
    - String と 数値型                        → 文字を数に直して比べる(数にならなければ型エラー 13)
    - String と(Null 以外の)Variant          → 文字で比べる(Variant の数は CStr して比べる)
    - Empty と 数値型 / String                → 0 / "" として比べる
    - Variant(数) と Variant(文字列)        → **数のほうが小さい**(エラーにしない)
    """
    def kind(x: Any, t: str) -> str:
        if t != VAR:
            return t
        if x is None:
            return "empty"
        if isinstance(x, (int, float)) and not isinstance(x, bool):
            return "vnum"
        if isinstance(x, bool):
            return "vnum"
        return "vstr"

    ka, kb = kind(a, a_type), kind(b, b_type)

    def num(x: Any) -> float:
        if isinstance(x, bool):
            return -1.0 if x else 0.0
        return float(x) if not isinstance(x, str) else cdbl(x)

    def cmp(x: Any, y: Any) -> int:
        return (x > y) - (x < y)

    # Empty は相手に合わせて 0 / ""
    if ka == "empty":
        a, ka = (0, NUM) if kb in (NUM, "vnum") else ("", STR)
        if kb == "empty":
            return 0
    if kb == "empty":
        b, kb = (0, NUM) if ka in (NUM, "vnum") else ("", STR)

    if ka == STR and kb == STR:
        return cmp(str(a), str(b))
    if (ka == STR and kb == NUM) or (ka == NUM and kb == STR):   # String と数値型 → 文字を数に直して比べる
        s = a if ka == STR else b
        if not is_numeric(s):
            raise TypeMismatch(f"{a!r} と {b!r} の比較(String と数値型。数にならない文字列)")
        return cmp(num(a), num(b))
    if ka == STR or kb == STR:                     # String と Variant → 文字で比べる
        sa = str(a) if ka in (STR, "vstr") else cstr(a)
        sb = str(b) if kb in (STR, "vstr") else cstr(b)
        return cmp(sa, sb)
    if ka == NUM or kb == NUM:                     # 数値型が片方にある
        other, other_k = (b, kb) if ka == NUM else (a, ka)
        if other_k == "vstr" and not is_numeric(other):
            raise TypeMismatch(f"{a!r} と {b!r} の比較(数値型と数にならない文字列)")
        return cmp(num(a), num(b))
    # 両方 Variant
    if ka == "vnum" and kb == "vnum":
        return cmp(num(a), num(b))
    if ka == "vstr" and kb == "vstr":
        return cmp(str(a), str(b))
    return -1 if ka == "vnum" else 1               # 数の Variant < 文字列の Variant
