"""品質基準 DB の照合エンジン(VBA `【UFKen_11】QP_DB` と、照合に使う `Quality_DB`・`【UFKen_17】補助` の関数)

VBA の行番号は `vba_src/【UFKen_11】QP_DB.bas`(QP)、`【UFKen_11】Quality_DB.bas`(QDB)、
`【UFKen_17】補助.bas`(H) のもの。解析は docs/analysis/G4_Quality_Admin.md。

【この移植の約束】
- 値は VBA と同じく文字列のまま受け取り、Val / CSng / CLng / Format は vbx を通す。
- VBA の配列の「列番号」は 1 始まりのまま扱う(列が見つからなければ 0。0 の列を読むと VBA は
  「添字が範囲外」になるので `VbaRuntimeError(9)` を投げる)。行は Python のリスト(0 始まり)。
- 行の順は DB の取り込み順(=Access の返却順)。「最初の一致」「最後の一致」を VBA のまま再現する。
- MsgBox は `flow.Message` で返す。VBA の実行時エラーは例外(vbx.TypeMismatch / VbaRuntimeError)。
- 画面(UF1・UF品質基準・シート)に書いていた値は、戻り値の dict に「コントロール名 → 値」で入れる。

【U1(数と数にならない文字列の比較)】
VBA の `Kousa(1,4) = 0`(中耳列が "" や "6㎜以下,…")、`JoinQuality(i, 番号) = NewArray(a)`(見出し行
"管理番号" と Long)、`UF1.MG = 0`(空のテキストボックス)は、毎日動いている経路で必ず通る。
比較は `vbx.compare_typed`(VBA の比較演算子の表)に両側の宣言の型を渡して行う。表では「数値型 と
数にならない文字列の Variant」は型エラー 13 だが、上の経路ではエラー 13 になるなら本番が動かないので、
**エラーにせず Variant の規則(数 < 文字列。= は偽)**で比べる(`vbx.U1_TEXT_VS_NUMBER_RAISES = False`)。
実機確認 U1 でエラーと分かったら vbx のスイッチを True にする(grid・checks も同じスイッチを見る)。
"""
from __future__ import annotations

import unicodedata
from dataclasses import dataclass, field
from decimal import ROUND_DOWN, Decimal
from pathlib import Path
from typing import Any, Callable, Optional, Sequence, Union

from . import cache_registry, datasource, lines, match_rules, paths, size_thresholds, vbx
from .flow import ERROR, Message

VBCRLF = "\r\n"

#: U1 の扱い(上の説明)。`vbx.U1_TEXT_VS_NUMBER_RAISES` 1 か所で決める。ここは個別に上書きするときだけ
#: (None = vbx に従う、True = 数と数にならない文字列の比較で TypeMismatch、False = 数 < 文字列)
STRING_VS_NUMBER_RAISES: Optional[bool] = None

Row = list
Rows = list


class VbaRuntimeError(RuntimeError):
    """VBA の実行時エラー(番号つき)。9 = 添字が範囲外、5 = 手続きの呼び出しが不正。"""

    def __init__(self, number: int, text: str) -> None:
        super().__init__(f"実行時エラー {number}: {text}")
        self.number = number


# ------------------------------------------------------------------
# VBA 互換の小道具(共通化の候補)
# ------------------------------------------------------------------
def col_of(fields: Optional[Sequence[str]], name: str) -> int:
    """`For i: Select Case F(1, i): Case "名": C = i` の C(1 始まり)。

    ループは途中で抜けないので、同じ名前が2つあれば**後の列**。無ければ 0。
    """
    c = 0
    for i, f in enumerate(fields or (), start=1):
        if f == name:
            c = i
    return c


def cell(row: Sequence[Any], c: int) -> Any:
    """`Arr(i, c)`(c は 1 始まり)。範囲外は VBA と同じく添字エラー。"""
    if c < 1 or c > len(row):
        raise VbaRuntimeError(9, f"列 {c} は範囲外(列数 {len(row)})")
    return row[c - 1]


def vstr(v: Any) -> str:
    """String 変数への代入・`&` での連結。Empty(None) は ""、数は CStr。"""
    return vbx.cstr(v)


def csng(v: Any) -> float:
    """`Dim x As Single: x = 値`。空文字・数にならない文字列は型エラー 13。Empty は 0。"""
    if v is None:
        return 0.0
    if isinstance(v, str):
        return vbx.single(vbx.cdbl(v))
    return vbx.single(float(v))


def clng(v: Any) -> int:
    """`Dim x As Long: x = 値`(銀行丸め)。Empty は 0。"""
    if v is None:
        return 0
    return vbx.clng(v)


def cbool(v: Any) -> bool:
    """VBA `CBool`。"True"/"False"(大小無視)と数の文字列。それ以外は型エラー 13。"""
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    s = str(v).strip()
    if s.lower() == "true":
        return True
    if s.lower() == "false":
        return False
    return vbx.cdbl(s) != 0


NUM, STR, VAR = vbx.NUM, vbx.STR, vbx.VAR


def vcmp(a: Any, b: Any, a_type: str = VAR, b_type: str = VAR) -> int:
    """VBA の比較(両側の宣言の型つき。`vbx.compare_typed`)。

    数値型と「数にならない文字列の Variant」(VBA の表では型エラー 13)だけは U1 の扱いにする:
    `STRING_VS_NUMBER_RAISES` が偽(既定)なら Variant の規則(数 < 文字列)で比べてエラーにしない。
    """
    try:
        return vbx.compare_typed(a, b, a_type, b_type)
    except vbx.TypeMismatch:
        raises = vbx.U1_TEXT_VS_NUMBER_RAISES if STRING_VS_NUMBER_RAISES is None else STRING_VS_NUMBER_RAISES
        if raises:
            raise
        a_num = a_type == NUM or (a_type == VAR and isinstance(a, (int, float)) and not isinstance(a, bool))
        return -1 if a_num else 1


def veq(a: Any, b: Any, a_type: str = VAR, b_type: str = VAR) -> bool:
    """VBA の `=`(型つき。U1 は vcmp と同じ)。"""
    return vcmp(a, b, a_type, b_type) == 0


def isnum(v: Any) -> bool:
    """VBA `IsNumeric(Variant)`。Empty(None) は True(VBA の仕様)。"""
    if v is None:
        return True
    return vbx.is_numeric(v)


def vtrim(s: str) -> str:
    """VBA `Trim`(日本語版は全角空白も取る)。"""
    return s.strip(" 　")


def vleft(s: str, n: int) -> str:
    """VBA `Left`。長さが負なら実行時エラー 5。"""
    if n < 0:
        raise VbaRuntimeError(5, f"Left の長さ {n}")
    return vbx.left(s, n)


def vsplit(s: str, sep: str) -> list[str]:
    """VBA `Split`。空文字は要素 0 個(UBound = -1)。"""
    return [] if s == "" else s.split(sep)


# StrConv(vbNarrow): 全角 → 半角(カタカナは濁点を分けた半角カナ)。ひらがな・漢字はそのまま
_NARROW: dict[str, str] = {}
for _c in range(0xFF01, 0xFF5F):
    _NARROW[chr(_c)] = chr(_c - 0xFEE0)
_NARROW["　"] = " "
_NARROW["￥"] = "\\"
_NARROW.update({"‘": "`", "’": "'", "“": '"', "”": '"',
                "゛": "ﾞ", "゜": "ﾟ"})
for _c in range(0xFF61, 0xFFA0):
    _h = chr(_c)
    _f = unicodedata.normalize("NFKC", _h)
    if len(_f) == 1 and _f not in _NARROW and _f != _h:
        _NARROW[_f] = _h
    for _m in ("ﾞ", "ﾟ"):
        _f2 = unicodedata.normalize("NFKC", _h + _m)
        if len(_f2) == 1 and _f2 not in _NARROW:
            _NARROW[_f2] = _h + _m
_NARROW.pop("゙", None)
_NARROW.pop("゚", None)


def narrow(s: str) -> str:
    """VBA `StrConv(s, vbNarrow)`(日本語版)。"""
    return "".join(_NARROW.get(ch, ch) for ch in s)


# ------------------------------------------------------------------
# 仕掛(SIKALOTNOW)の受け取り
# ------------------------------------------------------------------
@dataclass
class LotData:
    """VBA `Arrbox("SIKALOTNOW.accdb", "板", "Fields")` と `Arrbox(…, ロット, "Record")` の結果。

    Main/Next の切替(TimeDesignation・PLEASE WAIT)は Arrbox の担当。ここでは結果だけ受け取る。
    None は VBA の Empty。
    """
    fields: Optional[list[str]]
    records: Optional[list[list[str]]]

    def value(self, name: str) -> str:
        """`LotAcc(1, C_名)`(先頭行。列が無ければ添字エラー)。"""
        return vstr(cell(self.records[0], col_of(self.fields, name)))


def read_lot(lot_no: str, db: Union[str, Path] = "SIKALOTNOW") -> LotData:
    """切替なしで仕掛を読む(試験・単純な呼び出し用)。"""
    fields = datasource.get_fields(db, "仕掛")
    records = datasource.get_records(db, "仕掛", "[ﾛｯﾄ番号]='" + lot_no + "'")
    return LotData(fields, records)


def _packing_db(file: str) -> Path:
    """PackingPath & ファイル名(代替の置き場)。"""
    return paths.to_local(paths.join(paths.get("packing"), file))


def _quality_db() -> Path:
    return datasource.resolve("品質基準")


def _transfer_db() -> Path:
    return datasource.resolve("伝送用ファイル")


# ------------------------------------------------------------------
# 照合関数(QP L804-1173)
# ------------------------------------------------------------------
def find_in_list(tmp: str, lst: str, blank_flag: str = "") -> bool:
    """QP.FindInList (QP:804-913)。tmp = DB 側の条件、lst = 探す値。

    - tmp が "" なら真(blank_flag を渡したときは調べない)
    - カンマで分けて Trim し、左から評価。肯定条件は一致すると真にするだけ(後勝ちの OR)
    - `≠X` は X と違えば**真にする**、同じなら偽で終わる。`≠*Y*`/`≠*Y`/`≠Y*` は該当すれば偽で終わる
    - `*Y*` 部分一致、`*Y` 後方一致、`Y*` 前方一致、それ以外は完全一致(大小・全角半角を区別)
    """
    result = False
    if blank_flag == "" and tmp == "":
        return True
    for item in [vtrim(x) for x in tmp.split(",")]:
        if vbx.left(item, 1) == "≠":
            cmp = vbx.mid(item, 2)
            if cmp != lst:
                result = True
            else:
                return False
            if vbx.left(cmp, 1) == "*":
                if vbx.right(cmp, 1) == "*":
                    mid_item = vbx.mid(cmp, 2)
                    left_item = vleft(mid_item, len(mid_item) - 1)
                    if vbx.instr(lst, left_item) > 0:
                        return False
                else:
                    mid_item = vbx.mid(cmp, 2)
                    if vbx.right(lst, len(mid_item)) == mid_item:
                        return False
            elif vbx.right(cmp, 1) == "*":
                left_item = vleft(cmp, len(cmp) - 1)
                if vbx.left(lst, len(left_item)) == left_item:
                    return False
            else:
                if cmp == lst:
                    return False
        elif vbx.left(item, 1) == "*":
            if vbx.right(item, 1) == "*":
                mid_item = vbx.mid(item, 2)
                left_item = vleft(mid_item, len(mid_item) - 1)
                if vbx.instr(lst, left_item) > 0:
                    result = True
            else:
                mid_item = vbx.mid(item, 2)
                if vbx.right(lst, len(mid_item)) == mid_item:
                    result = True
        elif vbx.right(item, 1) == "*":
            left_item = vleft(item, len(item) - 1)
            if vbx.left(lst, len(left_item)) == left_item:
                result = True
        else:
            if item == lst:
                result = True
    return result


def check_condition(row: Sequence[Any], column: int, criteria: str, blank_flag: str = "") -> bool:
    """QP.CheckCondition (QP:916-926)。DB 側だけ半角化して「ｰ」→「-」にしてから FindInList。"""
    tmp = vstr(cell(row, column))
    out = narrow(tmp) if tmp != "" else ""
    out = out.replace("ｰ", "-")
    return find_in_list(out, criteria, blank_flag)


def _order_val(order: float, is_single: bool) -> float:
    return vbx.val_single(order) if is_single else vbx.val(order)


def check_thickness(row: Sequence[Any], c1: int, c2: int, c3: int, order: float, *,
                    is_single: bool = True, same_precision: bool = False) -> bool:
    """QP.CheckThickness (QP:928-945)。下・上・値列の3つとも "" なら真。それ以外は 下 <= x <= 上。

    下・上は Single(`Val` を Single に入れる)、x は `Val(Single の引数)`(有効 7 桁)。
    same_precision: 下・上も `Val` のまま(Double)で比べる(B71 を直した側。呼び出し側は真を渡す)。
    """
    # VBA の And は短絡しないので3列とも必ず読む(値列が無い表でも添字エラー 9。QP:930-932)
    v1, v2, v3 = cell(row, c1), cell(row, c2), cell(row, c3)
    if vstr(v1) == "" and vstr(v2) == "" and vstr(v3) == "":
        return True
    if same_precision:
        lower = vbx.val(cell(row, c1))
        upper = vbx.val(cell(row, c2))
    else:
        lower = vbx.single(vbx.val(cell(row, c1)))
        upper = vbx.single(vbx.val(cell(row, c2)))
    x = _order_val(order, is_single)
    return lower <= x and upper >= x


def check_width(row: Sequence[Any], c1: int, c2: int, c3: int, order: float, *,
                is_single: bool = True, same_precision: bool = False) -> bool:
    """QP.CheckWidth (QP:947-964)。CheckThickness と同じ判定(幅・丈)。"""
    return check_thickness(row, c1, c2, c3, order, is_single=is_single, same_precision=same_precision)


def check_within(row: Sequence[Any], c1: int, c2: int, order: float, *, is_single: bool = True,
                 same_precision: bool = False) -> bool:
    """QP.CheckWithin (QP:966-972)。空の確認なしで 下 <= x <= 上(両方空なら 0..0)。

    same_precision: 下・上も `Val` のまま(Double)で比べる(B71 を直した場合)。
    """
    if same_precision:
        lower = vbx.val(cell(row, c1))
        upper = vbx.val(cell(row, c2))
    else:
        lower = vbx.single(vbx.val(cell(row, c1)))
        upper = vbx.single(vbx.val(cell(row, c2)))
    x = _order_val(order, is_single)
    return lower <= x and upper >= x


def check_size(row: Sequence[Any], c1: int, order_size: Any) -> bool:
    """QP.CheckSize (QP:974-981)。空なら真。それ以外は Val(指示) >= CLng(Val(DB))。"""
    if vstr(cell(row, c1)) == "":
        return True
    low = vbx.clng(vbx.val(cell(row, c1)))
    return vbx.val(order_size) >= low


def find_indices(fields: Sequence[str], width_length: str, return_size: float) -> tuple[int, Optional[int]]:
    """QP.FindIndices (QP:985-1049)。("H"/"T" の列番号, 9999 または None)。

    - 列名が「H/T + 数」の列を集めて昇順に並べ、上端を含む区間 (前列, 当列] で探す
    - 最小列は 0 < x <= 最小。範囲外(最大超・0 以下)は最大列と 9999
    - 並べ替えの入れ替え用変数が Long なので小数の列名は丸まる(B27。直すと丸めない)。
      丸まった値で列名を探すので見つからず、Temp の添字がそのまま列番号として返る(VBA どおり)
    - 該当列が無いと VBA は実行時エラー(9)
    """
    temp: list[float] = []
    for name in fields:
        name = vstr(name)
        if len(name) < 1:
            raise VbaRuntimeError(5, "空の列名")
        if vbx.left(name, 1) == width_length and vbx.is_numeric(vbx.mid(name, 2, len(name) - 1)):
            temp.append(vbx.cdbl(vbx.mid(name, 2, len(name) - 1)))
    if not temp:
        raise VbaRuntimeError(9, f"{width_length} の列がありません")
    reproduce = False
    for i in range(len(temp) - 1):
        for j in range(i + 1, len(temp)):
            if temp[i] > temp[j]:
                t = temp[i]
                temp[i] = temp[j]
                temp[j] = float(vbx.clng(t)) if reproduce else t
    rs = return_size
    amatch = 0
    flag: Optional[int] = None
    ub = len(temp) - 1
    a = ub
    while a >= 0:
        n = a - 1
        if n == 0:
            if temp[n] >= rs and 0 < rs:
                amatch = n
                break
            if temp[a] >= rs and temp[n] < rs:
                amatch = a
                break
        else:
            if n == -1:
                amatch = ub
                flag = 9999
                break
            if temp[a] >= rs and temp[n] < rs:
                amatch = a
                break
        a -= 1
    match_size = width_length + vbx.cstr(temp[amatch])
    for i, name in enumerate(fields, start=1):
        if name == match_size:
            amatch = i
    return amatch, flag


def extract_num_from_part(part: str, condition: str) -> str:
    """QP.ExtractNumFromPart (QP:1139-1172)。condition を含む部分の、最初の「数字の並びと小数点1個」。"""
    if vbx.instr(part, condition) == 0:
        return ""
    result = ""
    found = False
    has_point = False
    for ch in part:
        if _char_isnumeric(ch):
            result += ch
            found = True
        elif ch == "." and not has_point:
            result += ch
            has_point = True
            found = True
        elif found:
            break
    return result


def _char_isnumeric(ch: str) -> bool:
    """`IsNumeric(1文字)`。全角数字も数(VBA は全角数字を数として読む)。"""
    return vbx.is_numeric(unicodedata.normalize("NFKC", ch)) if ch.strip() else False


def extract_nums(tmp: str, condition: str) -> str:
    """QP.ExtractNums (QP:1115-1137)。カンマで分け、condition を含む部分の数を ", " でつなぐ。"""
    out = ""
    for p in vsplit(tmp, ","):
        num = extract_num_from_part(p, condition)
        if num != "":
            out += num + ", "
    return out[:-2] if out else ""


def extract_number(s: str) -> Any:
    """QP.ExtractNumber (QP:1051-1091。呼び出し元なし)。以上/以下があれば最初の数を Double で。"""
    if vbx.instr(s, "以上") == 0 and vbx.instr(s, "以下") == 0:
        return ""
    result = ""
    found = False
    has_point = False
    for ch in s:
        if _char_isnumeric(ch):
            result += ch
            found = True
        elif ch == "." and not has_point:
            result += ch
            has_point = True
            found = True
        elif found:
            break
    return "" if result == "" else vbx.cdbl(unicodedata.normalize("NFKC", result))


def compare_strings(tmp: str, lot_atu: float, db_atu: float) -> bool:
    """QP.CompareStrings (QP:1093-1101)。"以上" なら Lot >= DB、なければ "以下" なら Lot <= DB。"""
    if vbx.instr(tmp, "以上") > 0:
        return lot_atu >= db_atu
    if vbx.instr(tmp, "以下") > 0:
        return lot_atu <= db_atu
    return False


def is_half_width(s: str) -> bool:
    """QP.IsHalfWidth (QP:1103-1113。呼び出し元なし)。"""
    return all(ord(ch) <= 0xFF for ch in s)


def _hit_ret(hit: Optional[dict[str, Any]], ri: int, col: int, row: Sequence[Any]) -> float:
    """LoopArray / SoriArray の戻り値。hit を渡すと当たった行(0 始まり)と列(1 始まり)を書く(結果は変えない)。"""
    v = csng(vbx.fmt(cell(row, col), "0.0"))
    if hit is not None:
        hit.update(row=ri, col=col, raw=vstr(cell(row, col)))
    return v


def loop_array(rows: Sequence[Sequence[Any]], fields: Sequence[str], return_size: float, atu: float,
               hit: Optional[dict[str, Any]] = None) -> float:
    """QDB.LoopArray (QDB:1386-1453)。列名が数の列を「列の並び順」でサイズ区間にして公差を引く。

    - 板厚下 <= 厚 <= 板厚上 の**最初の行**
    - a を右から: 区間 (前列, 当列] で一致する列。最小列は 0 < x <= 最小
    - 当てはまらない(最大超・0 以下)ときは**最後の数の列**(L1433-1436)
    - 戻り値は `Format(セル, "0.0")` を Single に入れたもの(空や文字はエラー 13)。該当行が無ければ 0
    """
    c_lo = col_of(fields, "板厚下")
    c_hi = col_of(fields, "板厚上")
    sizes = [vbx.val(f) for f in fields if vbx.is_numeric(f)]
    if not sizes:
        raise VbaRuntimeError(9, "ReDim TempSize(1 To 0)")
    for ri, row in enumerate(rows):
        und = vbx.single(vbx.val(cell(row, c_lo)))
        ove = vbx.single(vbx.val(cell(row, c_hi)))
        if not (und <= atu and ove >= atu):
            continue
        a = len(sizes)
        while a >= 1:
            n = a - 1
            if n == 0:
                if sizes[a - 1] >= return_size and 0 < return_size:
                    for ic, name in enumerate(fields, start=1):
                        if vbx.val(name) == sizes[0]:
                            return _hit_ret(hit, ri, ic, row)
                # GoTo skipErr: 最後の数の列
                a = len(sizes)
                for ic, name in enumerate(fields, start=1):
                    if vbx.is_numeric(name) and vbx.val(name) == vbx.val(sizes[a - 1]):
                        return _hit_ret(hit, ri, ic, row)
                break
            if sizes[a - 1] >= return_size and sizes[n - 1] < return_size:
                for ic, name in enumerate(fields, start=1):
                    if vbx.is_numeric(name) and vbx.val(name) == vbx.val(sizes[a - 1]):
                        return _hit_ret(hit, ri, ic, row)
            a -= 1
        # VBA は一致の有無に関わらず matchFlag で抜ける(列が見つからない場合は次の行へ)
    return 0.0


def sori_array(rows: Sequence[Sequence[Any]], fields: Sequence[str], atu: float, col: int,
               hit: Optional[dict[str, Any]] = None) -> float:
    """QP.SoriArray (QP:1458-1497)。col が 0 なら 0。板厚下 <= 厚 <= 板厚上 の最初の行の col 列。"""
    if col == 0:
        return 0.0
    c_lo = col_of(fields, "板厚下")
    c_hi = col_of(fields, "板厚上")
    for ri, row in enumerate(rows):
        lo = vbx.single(vbx.val(cell(row, c_lo)))
        hi = vbx.single(vbx.val(cell(row, c_hi)))
        if lo <= atu and hi >= atu:
            return _hit_ret(hit, ri, col, row)
    return 0.0


# ------------------------------------------------------------------
# 補助(H)のうち照合で使うもの
# ------------------------------------------------------------------
def find_priority(s: str) -> str:
    """H.FindPriority (H:1455-1491)。コース → 検査番号の頭(CM1→L, CM2→N, CM3→K, 2HI→T, HOT→H)。

    空白で分けて**右から** CM1/CM2/CM3/2HI の完全一致を探し、無ければ HOT。
    """
    words = s.split(" ")
    pri = ""
    for w in reversed(words):
        if w in ("CM1", "CM2", "CM3", "2HI"):
            pri = w
            break
    if pri == "":
        for w in words:
            if w == "HOT":
                pri = "HOT"
                break
    return {"CM1": "L", "CM2": "N", "CM3": "K", "2HI": "T", "HOT": "H"}.get(pri, pri)


def all_elements_same(arr: Sequence[Any]) -> bool:
    """H.AllElementsSame (H:1623-1634)。先頭要素の文字列と全部が等しいか。"""
    check = vstr(arr[0])
    return all(vstr(x) == check for x in arr)


def split_width_length(width_length: Sequence[str]) -> tuple[list[Any], list[Any]]:
    """H.SplitWidthLength (H:1636-1651)。"幅_丈" を分ける。"" の要素は Empty(None)のまま。"""
    w: list[Any] = [None] * len(width_length)
    l: list[Any] = [None] * len(width_length)
    for i, s in enumerate(width_length):
        if s != "":
            parts = s.split("_")
            w[i] = parts[0]
            if len(parts) < 2:
                raise VbaRuntimeError(9, "Split の結果に (1) がありません")
            l[i] = parts[1]
    return w, l


class NotAvailable(ValueError):
    """`CVErr(xlErrNA)`。Double に入れると型エラーになる値。"""


def second_smallest(arr: Sequence[float]) -> float:
    """H.SecondSmallest (H:1653-1680)。並べて、最小と異なる次の値。無ければ CVErr(#N/A)。"""
    t = sorted(arr)
    for i in range(len(t) - 1):
        if t[i] != t[i + 1]:
            return t[i + 1]
    raise NotAvailable("#N/A")


def weight(atu: float, haba: float, take: float) -> float:
    """OutTolerance.Weight (OutTolerance:813-817)。厚×幅×丈×2.68/10^6 を小数3桁で切り捨て(RoundDown)。

    別担当(OutTolerance)の関数。IsGaugeRequired だけのためにここにも置く(共通化の候補)。
    """
    x = atu * haba * take * 2.68 / 10 ** 6
    d = Decimal(vbx.cstr(x)) if x != 0 else Decimal(0)
    return float(d.quantize(Decimal("0.001"), rounding=ROUND_DOWN))


# ------------------------------------------------------------------
# DBQP(QP L69-801): 本番の歪公差
# ------------------------------------------------------------------
#: Kousa の列(VBA の 1〜9 → 0〜8)
K_GAUGE, K_HA, K_TA, K_NAKA, K_MIMI, K_COMMENT, K_SOURCE, K_NO, K_FLAG = range(9)


@dataclass
class QpResult:
    """DBQP / DistArray の結果。

    - kousa: VBA の `Kousa(行, 1..9)`。None は Empty(呼び手は何もしない)
    - form: UF品質基準 に書く値(Execution="" のとき。UF品質基準 が非表示でも VBA は暗黙に読み込んで書く)
    - uf1: UF1 が表示中なら書く値(コントロール名 → 値。".Caption" 付きはラベル)
    """
    kousa: Optional[list[list[Any]]] = None
    messages: list[Message] = field(default_factory=list)
    form: dict[str, Any] = field(default_factory=dict)
    uf1: dict[str, Any] = field(default_factory=dict)
    #: どの表のどの条件で何行残ったか(保守 → 歪公差チェックで見せる)。{"step", "detail", "ok": True/False/None, "rows"?}
    trace: list[dict[str, Any]] = field(default_factory=list)


def _t(res: "QpResult", step: str, detail: str = "", ok: Optional[bool] = None, **extra: Any) -> None:
    res.trace.append({"step": step, "detail": detail, "ok": ok, **extra})


#: 当たった行の見本(表の行をいくつまで画面に送るか)。保守 → 歪公差チェックの表プレビュー
_PREVIEW_AROUND = 3


def _preview(fields: Sequence[str], rows: Sequence[Sequence[Any]], hit_rows: Sequence[int],
             cols: Sequence[int], around: int = _PREVIEW_AROUND) -> dict[str, Any]:
    """表のプレビュー。hit_rows(0 始まり)の前後 around 行。cols は 1 始まりの列番号。行番号は 1 始まり。"""
    keep: set[int] = set()
    for h in hit_rows:
        keep.update(range(max(0, h - around), min(len(rows), h + around + 1)))
    if not hit_rows:
        keep.update(range(min(len(rows), around * 2 + 1)))
    cols = [c for c in dict.fromkeys(cols) if 0 < c <= len(fields)]
    return {"fields": [vstr(fields[c - 1]) for c in cols],
            "rows": [{"no": i + 1, "cells": [vstr(cell(rows[i], c)) for c in cols]} for i in sorted(keep)],
            "hit": [h + 1 for h in hit_rows], "total": len(rows)}


def _grade_cols(fields: Sequence[str], hit_col: int) -> list[int]:
    """等級表で見せる列: 板厚下・板厚上と、当たった列(と数・H・T の列)。"""
    out = [col_of(fields, "板厚下"), col_of(fields, "板厚上")]
    for i, name in enumerate(fields, start=1):
        n = vstr(name)
        if vbx.is_numeric(n) or (n[:1] in ("H", "T") and vbx.is_numeric(n[1:])):
            out.append(i)
    if hit_col:
        out.append(hit_col)
    return out


def _grade_trace(res: "QpResult", out: str, table: str, fields: Sequence[str], rows: Sequence[Sequence[Any]],
                 hit: dict[str, Any], atu: float, size: float, value: str) -> None:
    ri = hit.get("row")
    col = hit.get("col", 0)
    colname = vstr(fields[col - 1]) if 0 < col <= len(fields) else ""
    lo, hi = col_of(fields, "板厚下"), col_of(fields, "板厚上")
    if ri is None:
        detail = f"表「{table}」 厚 {vbx.fmt(atu, '0.000')}・寸法 {vbx.fmt(size, '0.0')} に当たる行がありません → {value}"
        hits: list[dict[str, Any]] = []
    else:
        r = rows[ri]
        detail = (f"表「{table}」 {ri + 1} 行目(板厚 {vstr(cell(r, lo))}〜{vstr(cell(r, hi))})の列「{colname}」"
                  f" = {hit.get('raw', '')} → {value}(厚 {vbx.fmt(atu, '0.000')} / 寸法 {vbx.fmt(size, '0.0')})")
        hits = [{"no": ri + 1, "keys": {"板厚下": vstr(cell(r, lo)), "板厚上": vstr(cell(r, hi))},
                 "col": colname, "value": hit.get("raw", "")}]
    _t(res, f"{out}: 当たった行", detail, ri is not None, table=table, output=out, hits=hits,
       compared={"板厚": vbx.fmt(atu, "0.000"), "寸法": vbx.fmt(size, "0.0")}, chosen=value,
       preview=_preview(fields, rows, [ri] if ri is not None else [], _grade_cols(fields, col)))


def _loop_hit(res: "QpResult", out: str, table: str, fields: Sequence[str], rows: Sequence[Sequence[Any]],
              size: float, atu: float) -> str:
    """`Format(LoopArray(...), "0.0")` と同じ値を返し、当たった行を trace に書く。"""
    hit: dict[str, Any] = {}
    v = vbx.fmt(loop_array(rows, fields, size, atu, hit=hit), "0.0")
    _grade_trace(res, out, table, fields, rows, hit, atu, size, v)
    return v


def _sori_hit(res: "QpResult", out: str, table: str, fields: Sequence[str], rows: Sequence[Sequence[Any]],
              atu: float, col: int, size: float) -> str:
    """`Format(SoriArray(...), "0.0")` と同じ値を返し、当たった行を trace に書く。"""
    hit: dict[str, Any] = {}
    v = vbx.fmt(sori_array(rows, fields, atu, col, hit=hit), "0.0")
    if col and "col" not in hit:
        hit["col"] = col
    _grade_trace(res, out, table, fields, rows, hit, atu, size, v)
    return v


def _table(db: Union[str, Path], table: str) -> tuple[Optional[list[str]], Optional[list[list[str]]]]:
    """`GetFieldsArr(表, "")` と `GetRecordsArr(表, "")`(VBA は2回開くが結果は同じ)。"""
    return datasource.get_fields(db, table), datasource.get_records(db, table)


def dbqp(lot_no: str, ken_no: str, delivery: str, good: str, lot: Optional[LotData], *,
         execution: str = "", uf1_visible: bool = False) -> QpResult:
    """QP.DBQP (QP:69-801)。ロットの歪公差 Kousa(1..n, 1..9) を決める。

    lot: Arrbox の結果。None は VBA の ReadCheck で SIKALOTNOW が無かったとき(QP:91-93)。
    delivery / good: 引当の納入先名称・取引先名称(DistArray から)。
    """
    res = QpResult()
    msgs = res.messages
    if lot is None:
        msgs.append(Message(paths.missing_text("SIKALOTNOW"), ERROR, source="QP:92"))
        _t(res, "仕掛(SIKALOT)を読む", "ファイルが読めません", False)
        return res
    hi = lot.fields
    if hi is None:
        _t(res, "仕掛(SIKALOT)を読む", "列名が取れません", False)
        return res
    if lot.records is None:
        msgs.append(Message("SIKALOTNOWに該当ロットがありません_DBQP", ERROR, source="QP:162"))
        _t(res, "仕掛(SIKALOT)を読む", f"ロット {lot_no} がありません", False)
        return res
    la = lot.records[0]

    def lv(name: str) -> Any:
        return cell(la, col_of(hi, name))

    ken_a = vbx.left(ken_no, 1)
    if ken_a == "0" or ken_a == "":
        result_c = find_priority(vstr(lv("実績_設備ｺｰｽ")))
        if result_c != "":
            ken_a = result_c
        if result_c == "":
            msgs.append(Message("検査Noがないファイルです_DBQP", ERROR, source="QP:174"))
            _t(res, "検番の頭", f"検番 {ken_no or '(空)'}・実績_設備ｺｰｽ から決められません", False)
            return res

    o_atu = csng(lv("ｵｰﾀﾞｰ板厚"))
    o_haba = csng(lv("ｵｰﾀﾞｰ板幅"))
    o_take = csng(lv("ｵｰﾀﾞｰ板丈"))
    m_atu = csng(lv("製造板厚"))
    m_haba = csng(lv("製造板幅"))
    m_take = csng(lv("製造板丈"))
    flat_raw = lv("品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ")
    if not vbx.is_numeric(flat_raw):
        flat = 0                          # B74 を直す: 等級が「.」などの数でない値も「3,4 以外」として扱う
    else:
        flat = clng(flat_raw)
    grade = vstr(lv("品質ｸﾞﾚｰﾄﾞ_表面処理"))
    vstr(lv("納入先名"))
    nou_code = vstr(lv("納入先ｺｰﾄﾞ"))
    okuri = vstr(lv("送り先名"))
    okuri_code = vstr(lv("送り先ｺｰﾄﾞ"))
    vstr(lv("取引先名"))
    tori_code = vstr(lv("取引先ｺｰﾄﾞ"))
    youto = vstr(lv("用途ｺｰﾄﾞ"))
    vstr(lv("用途名"))
    m_zai = vstr(lv("製造材質"))
    m_tyo = vstr(lv("製造調質"))
    o_zai = vstr(lv("ｵｰﾀﾞｰ材質"))
    o_tyo = vstr(lv("ｵｰﾀﾞｰ調質"))
    yokowari = vstr(lv("BOX設計_横割数(板)"))
    naka_sika = vstr(lv("指示_中歪_高さ"))
    mimi_sika = vstr(lv("指示_耳歪_高さ"))
    vstr(lv("指示_そり巾_方向高さ"))
    vstr(lv("指示_そり丈_方向高さ"))

    db = _quality_db()
    table_name = "平坦度一覧QP追加"
    _t(res, "ロットの値", f"検番の頭 {ken_a} / 用途 {youto} / 材質 {m_zai}({o_zai}) / 調質 {m_tyo}({o_tyo}) / "
       f"オーダー {vbx.fmt(o_atu, '0.000')}×{vbx.fmt(o_haba, '0.0')}×{vbx.fmt(o_take, '0.0')} / "
       f"歪等級 {vstr(flat_raw)} / 納入先 {delivery} / 取引先 {good}", None)
    fq = datasource.get_fields(db, table_name)
    if fq is None:
        _t(res, "参照マスタ", f"{db} の表「{table_name}」が読めません(空)", False)
        return res
    C = {n: col_of(fq, n) for n in (
        "平坦度上限値参照元", "用途コード", "板厚下", "板厚上", "1m当たり平坦度上限", "納入先", "納入先コード",
        "送り先", "送り先コード", "取引先", "取引先コード", "巾取数幅分割数のこと", "検査番号の頭", "板幅下", "板幅上",
        "板丈下", "板丈上", "オーダー材質", "オーダー調質", "製造調質", "製造材質", "用途名", "表示コメント", "管理番号",
        "オーダー板厚", "オーダー板幅", "オーダー板丈", "中歪・耳歪", "フラグ", "更新日")}
    flt = "[検査番号の頭] LIKE '*" + ken_a + "*' OR [検査番号の頭] =  ''"
    aq = datasource.get_records(db, table_name, flt)
    if aq is None:
        _t(res, "参照マスタ", f"{db} の表「{table_name}」に 検査番号の頭 が {ken_a} か空の行がありません", False)
        return res
    _t(res, "参照マスタ", f"{db} の表「{table_name}」 検査番号の頭 {ken_a} か空 → {len(aq)} 行", True, master=str(db),
       table=table_name)
    update_day = vstr(cell(aq[0], C["更新日"]))
    hit_haba = find_indices(fq, "H", o_haba)
    hit_take = find_indices(fq, "T", o_take)
    search_s = ""

    matching: list[int] = []
    same = True
    last = 15
    cond_names = ["", "用途コード", "製造材質", "オーダー材質", "製造調質", "オーダー調質", "板厚(下〜上)",
                  "板幅(下〜上)", "板丈(下〜上)", "納入先", "納入先コード", "取引先", "取引先コード", "送り先",
                  "送り先コード", "巾取数"]
    cond_vals = ["", youto, m_zai, o_zai, m_tyo, o_tyo, vbx.fmt(o_atu, "0.000"), vbx.fmt(o_haba, "0.0"),
                 vbx.fmt(o_take, "0.0"), delivery, nou_code, good, tori_code, okuri, okuri_code, yokowari]
    cond_cols = [None, ("用途コード",), ("製造材質",), ("オーダー材質",), ("製造調質",), ("オーダー調質",),
                 ("板厚下", "板厚上"), ("板幅下", "板幅上"), ("板丈下", "板丈上"), ("納入先",), ("納入先コード",),
                 ("取引先",), ("取引先コード",), ("送り先",), ("送り先コード",), ("巾取数幅分割数のこと",)]
    survive = [0] * 16
    prev_rows: list[list[Any]] = [[] for _ in range(16)]
    kept_rows: list[list[int]] = [[] for _ in range(16)]     # 条件 i まで合った行(0 始まり。見本 20 行まで)
    matched_idx: list[int] = []
    for ri, r in enumerate(aq):
        c = [False] * 16
        c[1] = check_condition(r, C["用途コード"], youto)
        c[2] = c[1] & check_condition(r, C["製造材質"], m_zai)
        c[3] = c[2] & check_condition(r, C["オーダー材質"], o_zai)
        c[4] = c[3] & check_condition(r, C["製造調質"], m_tyo)
        c[5] = c[4] & check_condition(r, C["オーダー調質"], o_tyo)
        # B71(QP:313-315): 下・上を Single に入れて Val(オーダー寸法) の Double と比べる。直すと同じ精度
        c[6] = c[5] & check_thickness(r, C["板厚下"], C["板厚上"], C["オーダー板厚"], o_atu, same_precision=same)
        c[7] = c[6] & check_width(r, C["板幅下"], C["板幅上"], C["オーダー板幅"], o_haba, same_precision=same)
        c[8] = c[7] & check_width(r, C["板丈下"], C["板丈上"], C["オーダー板丈"], o_take, same_precision=same)
        c[9] = c[8] & check_condition(r, C["納入先"], delivery)
        c[10] = c[9] & check_condition(r, C["納入先コード"], nou_code)
        c[11] = c[10] & check_condition(r, C["取引先"], good)
        c[12] = c[11] & check_condition(r, C["取引先コード"], tori_code)
        c[13] = c[12] & check_condition(r, C["送り先"], okuri)
        c[14] = c[13] & check_condition(r, C["送り先コード"], okuri_code)
        c[15] = c[14] & check_size(r, C["巾取数幅分割数のこと"], yokowari)
        # B24: 送り先コード・巾取数が最終判定に入っていない(last = 13)
        for i in range(1, 16):
            if c[i]:
                survive[i] += 1
                if len(kept_rows[i]) < 20:
                    kept_rows[i].append(ri)
            elif (i == 1 or c[i - 1]) and len(prev_rows[i]) < 5:
                prev_rows[i].append(r)       # 前の条件までは合って、ここで外れた行(見本 5 行まで)
        if all(c[1:last + 1]):
            matching.append(clng(cell(r, C["管理番号"])))
            matched_idx.append(ri)

    before = len(aq)
    stuck = False
    for i in range(1, last + 1):
        ok = survive[i] > 0
        cols = cond_cols[i]
        miss = []
        if survive[i] < before:
            for r in prev_rows[i]:
                miss.append(f"管理番号{vstr(cell(r, C['管理番号']))}: " + "〜".join(vstr(cell(r, C[k])) or "(空)" for k in cols))
        _t(res, f"条件 {i}: {cond_names[i]}", f"ロット {cond_vals[i] or '(空)'} → {before} 行中 {survive[i]} 行が合う",
           ok if not stuck else None, rows=miss, kept=survive[i], before=before, table=table_name,
           hits=[{"no": x + 1, "keys": {"管理番号": vstr(cell(aq[x], C["管理番号"])),
                                        **{k: vstr(cell(aq[x], C[k])) for k in cols}}} for x in kept_rows[i]])
        if not ok and not stuck:
            stuck = True
        before = survive[i]
    if not matching:
        _t(res, "平坦度一覧QP追加の判定", "合う行がありません → 等級表(ソリ・中歪・耳歪)へ(SkipQP)", False)
        return _dbqp_skip_qp(res, db, flat, m_tyo, o_atu, o_haba, o_take, naka_sika, mimi_sika,
                             execution, uf1_visible, update_day, m_atu, m_haba, m_take, m_zai, youto,
                             grade, delivery, flat_text=vstr(flat_raw))

    new_array = list(matching)
    # 優先: 納入先・取引先・送り先とコード・巾取数のどれかが記入された最初の行(QP:352-376)
    prio: Optional[int] = None
    for r in aq:
        for cm in new_array:
            if cm != 0:
                if veq(cell(r, C["管理番号"]), cm, VAR, NUM):
                    filled = [vstr(cell(r, C[k])) != "" for k in (
                        "納入先", "納入先コード", "取引先", "取引先コード", "送り先", "送り先コード", "巾取数幅分割数のこと")]
                    if any(filled):
                        prio = cm
                        break
        if prio is not None:
            break
    if prio is not None:
        new_array[0] = prio
    key_names = ["管理番号"] + [k for cs in cond_cols[1:] for k in cs]
    chosen_idx = next((x for x in matched_idx if veq(cell(aq[x], C["管理番号"]), new_array[0], VAR, NUM)), None)
    qp_hits = [{"no": x + 1, "keys": {k: vstr(cell(aq[x], C[k])) for k in key_names},
                "chosen": x == chosen_idx} for x in matched_idx]
    show_cols = [C[k] for k in key_names] + [C["1m当たり平坦度上限"], hit_haba[0], hit_take[0], C["中歪・耳歪"],
                                             C["平坦度上限値参照元"], C["フラグ"]]
    _t(res, "平坦度一覧QP追加の判定", f"合った管理番号 {', '.join(str(x) for x in matching)}"
       + (f"。納入先などが書いてある {prio} を優先" if prio is not None else f"。先頭の {new_array[0]} を使う"), True,
       table=table_name, hits=qp_hits, chosen=str(new_array[0]),
       preview=_preview(fq, aq, matched_idx, show_cols, around=1))

    kousa: list[list[Any]] = [[None] * 9 for _ in new_array]
    for a, no in enumerate(new_array):
        for r in aq:
            if veq(cell(r, C["管理番号"]), no, VAR, NUM):
                kousa[a][K_GAUGE] = cell(r, C["1m当たり平坦度上限"])
                kousa[a][K_HA] = cell(r, hit_haba[0])
                kousa[a][K_TA] = cell(r, hit_take[0])
                kousa[a][K_NAKA] = cell(r, C["中歪・耳歪"])
                kousa[a][K_MIMI] = ""
                kousa[a][K_COMMENT] = cell(r, C["表示コメント"])
                kousa[a][K_SOURCE] = cell(r, C["平坦度上限値参照元"])
                kousa[a][K_NO] = cell(r, C["管理番号"])
                kousa[a][K_FLAG] = cell(r, C["フラグ"])

    # 「重複可」のコメントを重ねる(QP:401-426)
    if len(kousa) >= 2:
        for k in kousa:
            if not veq(k[K_FLAG], "", VAR, STR) and vbx.instr(vstr(k[K_FLAG]), "重複可") > 0:
                if vbx.instr(vstr(kousa[0][K_COMMENT]), vstr(k[K_COMMENT])) == 0:
                    kousa[0][K_COMMENT] = vstr(kousa[0][K_COMMENT]) + VBCRLF + vstr(k[K_COMMENT])

    if hit_haba[1] == 9999:
        if vbx.instr(vstr(kousa[0][K_SOURCE]), "別表A") > 0:
            kousa[0][K_COMMENT] = vstr(kousa[0][K_COMMENT]) + VBCRLF + "3054mm越えは3050mmと同様の公差"

    if flat != 3 and flat != 4:
        msgs.append(Message("歪等級が3,4以外です", ERROR, source="QP:450"))
        _t(res, "歪等級", f"品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ = {vstr(flat_raw) or '(空)'}(3・4 以外)なので止まります", False)
        return res

    k0 = kousa[0]
    _t(res, "中歪・耳歪の欄", f"「{vstr(k0[K_NAKA]) or '(空)'}」 幅の列 {fq[hit_haba[0] - 1] if 0 < hit_haba[0] <= len(fq) else '?'}"
       f" / 丈の列 {fq[hit_take[0] - 1] if 0 < hit_take[0] <= len(fq) else '?'}", None,
       table=table_name, output="ソリ幅・ソリ丈・中歪/耳歪",
       hits=[] if chosen_idx is None else [{
           "no": chosen_idx + 1, "keys": {"管理番号": vstr(k0[K_NO])}, "chosen": True,
           "col": " / ".join(vstr(fq[c - 1]) if 0 < c <= len(fq) else "?" for c in (hit_haba[0], hit_take[0])),
           "value": f"ソリ幅 {vstr(k0[K_HA]) or '(空)'} / ソリ丈 {vstr(k0[K_TA]) or '(空)'}"
                    f" / 中歪・耳歪 {vstr(k0[K_NAKA]) or '(空)'}"}])
    goto = ""
    table_n = table_m = ""
    v4 = k0[K_NAKA]
    if veq(v4, "不問", VAR, STR) or veq(v4, "不問とする。", VAR, STR):
        goto = "SkipNoQuestionsAsked"
    elif veq(v4, 0, VAR, NUM) or veq(v4, "無き事", VAR, STR):
        goto = "SkipEndA"
    elif veq(v4, "", VAR, STR):
        goto = "SkipT"
    else:
        tmp = vstr(v4)
        elements = vsplit(tmp, ",")
        what_up = extract_nums(tmp, "以上")
        what_dow = extract_nums(tmp, "以下")
        up_flag = what_up != "" and what_up != " "
        dow_flag = what_dow != "" and what_dow != " "
        result = 0.0
        if up_flag:
            result = vbx.single(vbx.val(extract_nums(tmp, "以上")))
        if dow_flag:
            result = vbx.single(vbx.val(extract_nums(tmp, "以下")))
        has3 = vbx.instr(tmp, "3級")
        has4 = vbx.instr(tmp, "4級")
        tabl_flag = False
        if has3 > 0 and (up_flag or dow_flag):
            if compare_strings(tmp, o_atu, result):
                table_n, table_m, search_s, tabl_flag = "中歪3級", "耳歪3級", "Q-E-001_品質規格:等級3", True
        elif has4 > 0 and (up_flag or dow_flag):
            if compare_strings(tmp, o_atu, result):
                table_n, table_m, search_s, tabl_flag = "中歪4級", "耳歪4級", "Q-E-001_品質規格:等級4", True
        elif has3 > 0 and len(elements) - 1 == 0:
            table_n, table_m, search_s, tabl_flag = "中歪3級", "耳歪3級", "Q-E-001_品質規格:等級3", True
        elif has4 > 0 and len(elements) - 1 == 0:
            table_n, table_m, search_s, tabl_flag = "中歪4級", "耳歪4級", "Q-E-001_品質規格:等級4", True
        if not tabl_flag:
            k0[K_NAKA] = ""
        direct = not up_flag and not dow_flag
        if direct and tabl_flag:
            direct = False          # B26 を直す: 「3級」「4級」だけの行は等級表を引く
        if direct:
            naka_s: Any = extract_nums(tmp, "中")
            mimi_s: Any = extract_nums(tmp, "耳")
            if naka_s != "":
                naka_s = vbx.val(naka_s)
            if mimi_s != "":
                mimi_s = vbx.val(mimi_s)
            k0[K_NAKA] = naka_s
            k0[K_MIMI] = mimi_s
            goto = "skipEnd"
        elif has3 == 0 and has4 == 0:
            if compare_strings(tmp, o_atu, result):
                if flat == 3:
                    table_n, table_m, search_s = "中歪3級", "耳歪3級", "Q-E-001_品質規格:等級3"
                elif flat == 4:
                    table_n, table_m, search_s = "中歪4級", "耳歪4級", "Q-E-001_品質規格:等級4"
                f, rws = _table(db, table_n)
                if f is None or rws is None:
                    return res
                k0[K_NAKA] = _loop_hit(res, "中歪", table_n, f, rws, o_take, o_atu)
                f, rws = _table(db, table_m)
                if f is None or rws is None:
                    return res
                k0[K_MIMI] = _loop_hit(res, "耳歪", table_m, f, rws, o_take, o_atu)
        if goto == "":
            if table_n == "":
                goto = "SkipT"
            else:
                f, rws = _table(db, table_n)
                if f is None or rws is None:
                    return res
                k0[K_NAKA] = _loop_hit(res, "中歪", table_n, f, rws, o_take, o_atu)
                f, rws = _table(db, table_m)
                if f is None or rws is None:
                    return res
                k0[K_MIMI] = _loop_hit(res, "耳歪", table_m, f, rws, o_take, o_atu)
                if vbx.instr(vstr(k0[K_SOURCE]), search_s) == 0:
                    k0[K_SOURCE] = vstr(k0[K_SOURCE]) + " " + search_s
                goto = "SkipT"

    if goto in ("SkipT", "SkipNoQuestionsAsked"):
        if veq(k0[K_HA], "フラットネス品質等級", VAR, STR) or veq(k0[K_TA], "フラットネス品質等級", VAR, STR):
            sori = "ソリ3級" if flat == 3 else "ソリ4級"
            search_s = "Q-E-001_品質規格:等級3" if flat == 3 else "Q-E-001_品質規格:等級4"
            f, rws = _table(db, sori)
            if f is None or rws is None:
                return res
            sh = find_indices(f, "H", o_haba)
            st = find_indices(f, "T", o_take)
            k0[K_HA] = _sori_hit(res, "ソリ幅", sori, f, rws, o_atu, sh[0], o_haba)
            k0[K_TA] = _sori_hit(res, "ソリ丈", sori, f, rws, o_atu, st[0], o_take)
            if vbx.instr(vstr(k0[K_SOURCE]), search_s) == 0:
                k0[K_SOURCE] = vstr(k0[K_SOURCE]) + " " + search_s
            _t(res, "等級表", f"巾歪・丈歪が「フラットネス品質等級」なので 表「{sori}」から引きました", True)
    elif goto == "SkipEndA":
        k0[K_NAKA] = 0
        k0[K_MIMI] = 0

    if table_n:
        _t(res, "等級表", f"中 {table_n} / 耳 {table_m} を引きました", True)
    res.kousa = kousa
    return _dbqp_end(res, naka_sika, mimi_sika, execution, update_day, m_atu, m_haba, m_take,
                     m_zai, m_tyo, youto, grade, delivery)


def _dbqp_skip_qp(res: QpResult, db: Path, flat: int, m_tyo: str, o_atu: float, o_haba: float,
                  o_take: float, naka_sika: str, mimi_sika: str, execution: str, uf1_visible: bool,
                  update_day: str, m_atu: float, m_haba: float, m_take: float, m_zai: str, youto: str,
                  grade: str, delivery: str, flat_text: str = "") -> QpResult:
    """DBQP の SkipQP(QP:638-760)。QP 表に該当なし → O 材・Hx8 材の判定不要、または等級表。"""
    kousa = [[None] * 9]
    k0 = kousa[0]
    res.kousa = kousa
    if m_tyo == "O":
        _t(res, "SkipQP", "製造調質 O → 別表C 判定不要", True)
        k0[K_COMMENT] = "別表Cに記載のないO材_判定不要"
        k0[K_SOURCE] = "Q-S-QA-DIM-122_別表C"
        if execution == "" and uf1_visible:
            res.uf1["HizCome.Caption"] = "別表Cに記載のないO材_判定不要"
            res.uf1["DistortionFrame.Caption"] = "参照元:" + "Q-S-QA-DIM-122_別表C"
        return res
    if len(m_tyo) == 3 and vbx.left(m_tyo, 1) == "H" and vbx.right(m_tyo, 1) == "8":
        _t(res, "SkipQP", f"製造調質 {m_tyo}(Hx8) → 別表D 判定不要", True)
        k0[K_COMMENT] = "別表Dに記載のないHx8材_判定不要"
        k0[K_SOURCE] = "Q-S-QA-DIM-122_別表D"
        if execution == "" and uf1_visible:
            res.uf1["HizCome.Caption"] = "別表Dに記載のないHx8材_判定不要"
        return res

    table_name = "平坦度一覧QP追加"       # 等級が 3/4 以外ならこのまま(QP:229。B74)
    table_n = table_m = search_s = ""
    if flat == 3:
        table_n, table_m, table_name, search_s = "中歪3級", "耳歪3級", "ソリ3級", "Q-E-001_品質規格:等級3"
    elif flat == 4:
        table_n, table_m, table_name, search_s = "中歪4級", "耳歪4級", "ソリ4級", "Q-E-001_品質規格:等級4"

    def sori() -> bool:
        f, rws = _table(db, table_name)
        if f is None or rws is None:
            _t(res, "等級表", f"{db} の表「{table_name}」が読めません(空)", False)
            return False
        _t(res, "等級表", f"ソリ(幅・丈)を 表「{table_name}」から引きました", True)
        sh = find_indices(f, "H", o_haba)
        st = find_indices(f, "T", o_take)
        k0[K_HA] = _sori_hit(res, "ソリ幅", table_name, f, rws, o_atu, sh[0], o_haba)
        k0[K_TA] = _sori_hit(res, "ソリ丈", table_name, f, rws, o_atu, st[0], o_take)
        if vbx.instr(vstr(k0[K_SOURCE]), search_s) == 0:
            k0[K_SOURCE] = vstr(k0[K_SOURCE]) + " " + search_s
        return True

    if flat != 3 and flat != 4:
        # B74 を直す(2026-11-02 の回答で見直し): 等級が 3/4 以外なら、どの表の公差も当てはまらないので
        # ソリ(幅・丈)・中・耳は空にして知らせる。平坦度一覧QP追加の値は入れない
        _t(res, "歪等級", f"品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ = {flat_text or flat}(3・4 以外)なので公差は出しません", False)
        res.messages.append(Message(f"歪等級が3,4以外です(品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ={flat_text or flat})。"
                                    "歪公差は出しません。品証に確認してください", ERROR, source="QP:732"))
        k0[K_HA] = k0[K_TA] = k0[K_NAKA] = k0[K_MIMI] = ""
        return _dbqp_end(res, naka_sika, mimi_sika, execution, update_day, m_atu, m_haba, m_take,
                         m_zai, m_tyo, youto, grade, delivery)
    if not sori():
        res.kousa = None
        return res
    if size_thresholds.thin_atu() >= o_atu:          # 保守 → 寸法のしきい値(既定 6.0)
        if naka_sika == "" or mimi_sika == "":
            f, rws = _table(db, table_n)
            if f is None or rws is None:
                res.kousa = None
                return res
            k0[K_NAKA] = _loop_hit(res, "中歪", table_n, f, rws, o_take, o_atu)
            f, rws = _table(db, table_m)
            if f is None or rws is None:
                res.kousa = None
                return res
            k0[K_MIMI] = _loop_hit(res, "耳歪", table_m, f, rws, o_take, o_atu)
            _t(res, "等級表", f"中 {table_n} / 耳 {table_m} を引きました(厚 6.0 以下)", True)
    else:
        k0[K_NAKA] = ""
        k0[K_MIMI] = ""
        if flat != 3 and flat != 4:
            res.messages.append(Message("歪等級が3,4以外です", ERROR, source="QP:732"))
        else:
            table_name = "ソリ3級" if flat == 3 else "ソリ4級"
            search_s = "Q-E-001_品質規格:等級3" if flat == 3 else "Q-E-001_品質規格:等級4"
            if not sori():
                res.kousa = None
                return res
    return _dbqp_end(res, naka_sika, mimi_sika, execution, update_day, m_atu, m_haba, m_take,
                     m_zai, m_tyo, youto, grade, delivery)


def _dbqp_end(res: QpResult, naka_sika: str, mimi_sika: str, execution: str, update_day: str,
              m_atu: float, m_haba: float, m_take: float, m_zai: str, m_tyo: str, youto: str,
              grade: str, delivery: str) -> QpResult:
    """DBQP の skipEnd(QP:769-800)。"不問" を消し、SIKALOT の指示値で補い、UF品質基準 に書く値を作る。"""
    k0 = res.kousa[0]
    if veq(k0[K_GAUGE], "不問とする。", VAR, STR):
        k0[K_GAUGE] = ""
    if veq(k0[K_NAKA], "不問とする。", VAR, STR):
        k0[K_NAKA] = ""
    if veq(k0[K_NAKA], "不問", VAR, STR):
        k0[K_NAKA] = ""
    if veq(k0[K_MIMI], "不問", VAR, STR):
        k0[K_MIMI] = ""
    if veq(k0[K_NAKA], "", VAR, STR) and naka_sika != "":
        k0[K_NAKA] = naka_sika
    if veq(k0[K_MIMI], "", VAR, STR) and mimi_sika != "":
        k0[K_MIMI] = mimi_sika
    if execution != "":
        return res
    res.form = {
        "Uatu": vbx.fmt(m_atu, "0.000"), "Uhaba": vbx.fmt(m_haba, "0.0"), "Utake": vbx.fmt(m_take, "0.0"),
        "Uzai": m_zai, "Utyou": m_tyo, "UYoucode": youto, "Uhyou": grade, "Noun": delivery,
        "KousinL2.Caption": update_day, "KousinL3.Caption": update_day,
    }
    return res


# ------------------------------------------------------------------
# HikiArrCreate / DistArray / FormReflectDist(QP L1867-2055)
# ------------------------------------------------------------------
def hiki_arr_create(lot_no: str, main_db: Union[str, Path, None] = None,
                    next_db: Union[str, Path, None] = None) -> Optional[list[tuple[str, str]]]:
    """QP.HikiArrCreate (QP:1867-1909)。引当の (受注_納入先名称, 受注_取引先名称) を行順に。

    Main(MainPath)で列・行が取れなければ Next(NextPath)。行が Next から来ても列番号は Main のもの(VBA どおり)。
    `main_db` / `next_db` を渡さなければ、第1/第2の読み分けは検査画面の引当と同じ `lot_ledger`
    (ファイル・ロットが無ければ次に見る側、次に見る側にだけある引当も足す。列は Main の並びに直してある)。
    読めるファイルがどちらにも無ければ `datasource.DbMissing`(前と同じ)。
    """
    if lot_no == "":
        return None
    if main_db is None and next_db is None:
        from . import lot_ledger
        f = lot_ledger.read_fields("SIKAHIKINOW")
        if not f.found:
            raise datasource.DbMissing(f"DB ファイルが見つかりません: {f.file_path}")
        if f.cols is None:
            return None
        rows = lot_ledger.read_rows("SIKAHIKINOW", "[ﾛｯﾄ番号]='" + lot_no + "'").rows
        if rows is None:
            return None
        c_nou, c_tori = col_of(f.cols, "受注_納入先名称"), col_of(f.cols, "受注_取引先名称")
        return [(vstr(cell(r, c_nou)), vstr(cell(r, c_tori))) for r in rows]
    main_db = main_db or paths.db_file("SIKAHIKINOW")
    next_db = next_db or paths.db_file("SIKAHIKINOW", base_key="next")
    db: Union[str, Path] = main_db
    hf = datasource.get_fields(db, "仕掛")
    if hf is None:
        db = next_db
        hf = datasource.get_fields(db, "仕掛")
        if hf is None:
            return None
    c_nou = col_of(hf, "受注_納入先名称")
    c_tori = col_of(hf, "受注_取引先名称")
    flt = "[ﾛｯﾄ番号]='" + lot_no + "'"
    rows = datasource.get_records(db, "仕掛", flt)
    if rows is None:
        rows = datasource.get_records(next_db, "仕掛", flt)
        if rows is None:
            return None
    return [(vstr(cell(r, c_nou)), vstr(cell(r, c_tori))) for r in rows]


def dist_array(lot_no: str, ken_no: str, hiki: Optional[Sequence[tuple[str, str]]], lot: Optional[LotData], *,
               execution: str = "", uf1_visible: bool = False) -> QpResult:
    """QP.DistArray (QP:1911-1984)。引当ごとに DBQP を実行し、納入先で公差が違えば厳しい方で上書き。

    hiki: HikiArrCreate の結果(None = Empty で何もしない)。
    - 納入先名称が全部同じなら1件目の(納入, 取引)で1回だけ(取引先が違っても1件目)
    - 違えば引当ごとに実行。幅・丈の公差文字列が違えば Val の最小で1件目の結果を上書き。
      最小が 0 なら SecondSmallest(全部同じだと #N/A を Double に入れて型エラー = B27)
    - ゲージ・中・耳は1件目のまま。UF品質基準 に書く値は最後に実行した DBQP のもの
    """
    out = QpResult()
    if hiki is None:
        _t(out, "引当(SIKAHIKI)", f"ロット {lot_no} の引当がありません(VBA は何もしない)", False)
        return out
    result = [h[0] for h in hiki]

    def run(nou: str, tori: str) -> QpResult:
        r = dbqp(lot_no, ken_no, nou, tori, lot, execution=execution, uf1_visible=uf1_visible)
        if len(hiki) > 1 and not all_elements_same(result):
            _t(out, f"引当 納入先 {nou}", "この納入先で引きます", None)
        out.trace.extend(r.trace)
        out.messages.extend(r.messages)
        if r.form:
            out.form = r.form
        if r.uf1:
            out.uf1.update(r.uf1)
        return r

    if not all_elements_same(result):
        width_length = [""] * len(hiki)
        top: Optional[list[list[Any]]] = None
        for i, (nou, tori) in enumerate(hiki):
            former = run(nou, tori).kousa
            if i == 0:
                top = former
            if former is not None:
                width_length[i] = vstr(former[0][K_HA]) + "_" + vstr(former[0][K_TA])
        if not all_elements_same(width_length):
            wa, la = split_width_length(width_length)
            wn = [vbx.val(x) for x in wa]
            ln = [vbx.val(x) for x in la]
            min_w = min(wn)
            min_l = min(ln)
            try:
                if min_w == 0:
                    min_w = second_smallest(wn)
                if min_l == 0:
                    min_l = second_smallest(ln)
            except NotAvailable:
                # 直す: 全部同じ(0)なら 0 のまま
                pass
            if top is None:
                raise vbx.TypeMismatch("1件目の DBQP が Empty(TopFormerArr を参照できない)")
            top[0][K_HA] = min_w
            top[0][K_TA] = min_l
            top[0][K_COMMENT] = vstr(top[0][K_COMMENT]) + VBCRLF + "納入先による公差違い有,厳しい公差で上書き"
        out.kousa = top
    else:
        out.kousa = run(hiki[0][0], hiki[0][1]).kousa
    return out


@dataclass
class ReflectResult:
    """FormReflectDist の書き先と値。target: "UF1" / "UF品質基準" / "シート" / ""(書かない)。"""
    target: str = ""
    values: dict[str, Any] = field(default_factory=dict)
    sheet: str = ""
    cells: dict[tuple[int, int], Any] = field(default_factory=dict)


def form_reflect_dist(kousa: list[list[Any]], *, quality_visible: bool, mitsubishi_visible: bool,
                      uf1_visible: bool = True, zenkasho_sheet: bool = False) -> ReflectResult:
    """QP.FormReflectDist (QP:1986-2055)。公差を UF1 / UF品質基準 / 全ヵ所検査シートに反映する値。

    - UF品質基準・UFMitsubishi とも非表示: UF1 の MG/HaHiz/TaHiz/Naka/Mimi("0.0"、0 なら空)。
      HaHiz/TaHiz は数でなければ文字のまま、MG/Naka/Mimi は空。コメントの末尾 "_" を除き "。"→改行
    - UF品質基準 表示中: MG1〜Mimi1(数でなければ文字のまま)、Sansyou1 = 参照元 & "_" & コメント
    - UF1 非表示・全ヵ所検査シートあり・UFMitsubishi 表示中: セル (5,13)=幅歪、(5,15)=丈歪
    """
    k = kousa[0]
    out = ReflectResult()
    if not quality_visible and not mitsubishi_visible:
        v: dict[str, Any] = {}

        def num_or(x: Any, keep_text: bool) -> str:
            if isnum(x):
                s = vbx.fmt(x, "0.0")
            else:
                s = vstr(x) if keep_text else ""
            return "" if veq(s, 0, VAR, NUM) else s      # `UF1.MG = 0`(TextBox の Value は Variant)

        v["MG"] = num_or(k[K_GAUGE], False)
        v["HaHiz"] = num_or(k[K_HA], True)
        v["TaHiz"] = num_or(k[K_TA], True)
        v["Naka"] = num_or(k[K_NAKA], False)
        v["Mimi"] = num_or(k[K_MIMI], False)
        com = vstr(k[K_COMMENT])
        if vbx.right(com, 1) == "_":
            com = com[:-1]
        v["HizCome.Caption"] = com.replace("。", VBCRLF)
        v["DistortionFrame.Caption"] = "参照元:" + vstr(k[K_SOURCE])
        out.target, out.values = "UF1", v
    elif quality_visible:
        v = {}
        for name, i in (("MG1", K_GAUGE), ("HaHiz1", K_HA), ("TaHiz1", K_TA), ("Naka1", K_NAKA), ("Mimi1", K_MIMI)):
            v[name] = vbx.fmt(k[i], "0.0") if isnum(k[i]) else vstr(k[i])
        s = vstr(k[K_SOURCE]) + "_" + vstr(k[K_COMMENT])
        if vbx.right(s, 1) == "_":
            s = s[:-1]
        v["Sansyou1"] = s
        out.target, out.values = "UF品質基準", v
    elif not uf1_visible and zenkasho_sheet and mitsubishi_visible:
        out.target, out.sheet = "シート", "全ヵ所検査"
        out.cells = {(5, 13): k[K_HA], (5, 15): k[K_TA]}
    return out


def update_qp_arry_d(d: list[list[Any]], g: list[list[Any]]) -> list[list[Any]]:
    """QP.UpdateQpArryD (QP:2057-2081。呼び出し側はコメント)。幅・丈の小さい方で上書き。"""
    d12, d13 = vbx.val(d[0][K_HA]), vbx.val(d[0][K_TA])
    g12, g13 = vbx.val(g[0][K_HA]), vbx.val(g[0][K_TA])
    changed = False
    if d12 != g12:
        changed = True
    if d12 > g12:
        d[0][K_HA] = g12
    if d13 != g13:
        changed = True
    if d13 > g13:
        d[0][K_TA] = g13
    if changed:
        d[0][K_COMMENT] = vstr(d[0][K_COMMENT]) + VBCRLF + "納入先と取引先よる公差違い有,厳しい公差で上書き"
    return d


# ------------------------------------------------------------------
# DBdisplayラテラル(QDB L1491-1696)・DbReturnItem(QP L1176-1456)
# ------------------------------------------------------------------
def dbdisplay_lateral(lot: Optional[LotData], target_text: str, *, quiet: bool = False,
                      messages: Optional[list[Message]] = None) -> float:
    """QDB.DBdisplayラテラル (QDB:1491-1696)。ラテラルボー公差(DIM_179 日車/白銅滑川を製造丈で表引き)。

    lot: Arrbox の結果(None は SIKALOTNOW が無い)。quiet: IsOver12Hours が真(MsgBox を出さない)。
    条件: 材質・製造厚(板厚下/上)・製造幅(板幅下/上)。一致行で LoopArray(製造丈, 製造厚)。
    表が 0 件だと VBA は UBound で実行時エラー(B27。直すと 0)。
    """
    msgs = messages if messages is not None else []
    if lot is None:
        if not quiet:
            msgs.append(Message(paths.missing_text("SIKALOTNOW"), ERROR, source="QDB:1510"))
        return 0.0
    if lot.fields is None:
        return 0.0
    if lot.records is None:
        if not quiet:
            msgs.append(Message("SIKALOTNOWに該当ロットがありません_DBdisplayﾗﾃﾗﾙ", ERROR, source="QDB:1554"))
        return 0.0
    la, hf = lot.records[0], lot.fields
    for n in ("ｵｰﾀﾞｰ板厚", "ｵｰﾀﾞｰ板幅", "ｵｰﾀﾞｰ板丈"):
        csng(cell(la, col_of(hf, n)))
    m_atu = csng(cell(la, col_of(hf, "製造板厚")))
    m_haba = csng(cell(la, col_of(hf, "製造板幅")))
    m_take = csng(cell(la, col_of(hf, "製造板丈")))
    for n in ("納入先名", "用途ｺｰﾄﾞ", "用途名", "ｵｰﾀﾞｰ番号"):
        vstr(cell(la, col_of(hf, n)))
    zai = vstr(cell(la, col_of(hf, "製造材質")))
    vstr(cell(la, col_of(hf, "製造調質")))
    if target_text == "日車":
        table = "Q_S_QA_DIM_179_日車"
    elif target_text == "白銅滑川":
        table = "Q_S_QA_DIM_179_白銅滑川"
    else:
        raise VbaRuntimeError(3131, "FROM 句の構文エラー(表名が空)")
    db = _quality_db()
    fq = datasource.get_fields(db, table)
    if fq is None:
        return 0.0
    c_zai, c_lo, c_hi, c_atu = col_of(fq, "材質"), col_of(fq, "板厚下"), col_of(fq, "板厚上"), col_of(fq, "厚")
    c_wlo, c_whi, c_w, c_no = col_of(fq, "板幅下"), col_of(fq, "板幅上"), col_of(fq, "幅"), col_of(fq, "管理番号")
    aq = datasource.get_records(db, table)
    if aq is None:
        return 0.0
    matching: list[int] = []
    same = True          # B71(QDB:1637-1638): 直すと下・上も Double で比べる
    for r in aq:
        c1 = check_condition(r, c_zai, zai)
        c2 = c1 & check_thickness(r, c_lo, c_hi, c_atu, m_atu, same_precision=same)
        c3 = c2 & check_width(r, c_wlo, c_whi, c_w, m_haba, same_precision=same)
        if c1 and c2 and c3:
            matching.append(clng(cell(r, c_no)))
    if not matching:
        return 0.0
    hit = _hit_rows(aq, c_no, matching)
    return csng(vbx.fmt(loop_array(hit, fq, m_take, m_atu), "0.0"))


def _hit_rows(rows: Sequence[Sequence[Any]], c_no: int, numbers: Sequence[int]) -> list[list[Any]]:
    """`HitArry(a, j) = TempArr(i, j)`(管理番号が一致する行。同じ番号が複数あれば**後の行**)。"""
    width = len(rows[0]) if rows else 0
    hit: list[list[Any]] = []
    for no in numbers:
        h: list[Any] = [None] * width
        for r in rows:
            if veq(cell(r, c_no), no, VAR, NUM):
                h = list(r)
        hit.append(h)
    return hit


@dataclass
class ReturnItemResult:
    """DbReturnItem の結果。

    - items: 戻り値(",フリー反り,ラテラルボー,直角度" の部分集合)
    - uf1: UF1 に書く値(キー: "HURIS"・"Sumisei.Caption"・"Additional.Caption"・"LateralText"・
      "LateralText.BackColor"・"HizCome.Tag"・"HizCome.Caption"・"flatnesText"・"flatnesText.BackColor"・
      "LGS.Caption" など)。書いた順に入る
    - cb_add_atu: `CBAdd(ATU)` を呼ぶなら厚(OutTolerance の担当)
    - sheet / cells: UFMitsubishi 表示中に書くシートとセル
    """
    items: str = ""
    uf1: dict[str, Any] = field(default_factory=dict)
    cb_add_atu: Optional[float] = None
    sheet: str = ""
    cells: dict[tuple[int, int], Any] = field(default_factory=dict)
    messages: list[Message] = field(default_factory=list)


def rgb(r: int, g: int, b: int) -> int:
    """VBA `RGB(r, g, b)`。"""
    return r + g * 256 + b * 65536


ORANGE = rgb(255, 153, 0)


def _shadow_on(g: "match_rules.Group") -> bool:
    """旧コードと並べるか(試験、または既定のルールのままのとき。ルールを変えたら違って当然なので並べない)。"""
    if match_rules.SHADOW_STRICT:
        return g.rules == match_rules.BY_NAME[g.name].rules
    return False


def _ratefuri_old(ta: Sequence[Sequence[Any]], C: dict[str, int], args: tuple) -> tuple[list[int], list[bool]]:
    """旧: DbReturnItem の照合(QP:1240-1270)をコードで。照合ルールの shadow で差が 0 件と確かめたら消す。"""
    nou, tori, youtc, zai, tyou, atu, haba, take = args
    matching: list[int] = []
    c = [False] * 9
    # B71: 下限・上限を Single に入れて Double と比べるので、厚 0.6–0.6 のような境界ちょうどで外れる。
    # 直した場合は同じ精度(Double)で比べる(QP:1263-1265)
    same = True
    for r in ta:
        c = [False] * 9
        c[1] = check_condition(r, C["納入先"], nou)
        c[2] = c[1] & check_condition(r, C["取引先"], tori)
        c[3] = c[2] & check_condition(r, C["用途コード"], youtc)
        c[4] = c[3] & check_condition(r, C["製造材質"], zai)
        c[5] = c[4] & check_condition(r, C["製造調質"], tyou)
        c[6] = c[5] & check_thickness(r, C["厚下"], C["厚上"], C["製造板厚"], atu, same_precision=same)
        c[7] = c[6] & check_width(r, C["幅下"], C["幅上"], C["製造板幅"], haba, same_precision=same)
        c[8] = c[7] & check_width(r, C["丈下"], C["丈上"], C["製造板丈"], take, same_precision=same)
        if all(c[1:9]):
            matching.append(clng(cell(r, C["管理番号"])))
    return matching, [c[4], c[6], c[7], c[8]]


#: 該当なしの文言(材質・厚・幅・丈)を決める目印のルール(表の列の最初の列)
_RATEFURI_ANCHORS = ("製造材質", "厚下", "幅下", "丈下")


def ratefuri_ctx(args: tuple) -> dict[str, Any]:
    return dict(zip(match_rules.RATEFURI_FIELDS, args))


def _ratefuri_new(g: "match_rules.Group", ta: Sequence[Sequence[Any]], pf: Sequence[str], C: dict[str, int],
                  args: tuple) -> tuple[list[int], list[bool]]:
    """新: 照合ルール「ラテフリ直角」で絞る。戻りは旧と同じ形(管理番号, 最後の行の Conditions(4,6,7,8))。"""
    res = match_rules.match(g, ta, pf, ratefuri_ctx(args))
    matching = [clng(cell(ta[i], C["管理番号"])) for i in res.rows]
    return matching, ratefuri_flags(g, res)


def ratefuri_flags(g: "match_rules.Group", res: "match_rules.MatchResult") -> list[bool]:
    if res.evaluated == 0:
        return [False] * 4
    if res.last_fail is None:
        return [True] * 4
    rules = [r for r in sorted(g.rules, key=lambda r: r.order) if r.enabled]
    out = []
    for a in _RATEFURI_ANCHORS:
        pos = next((k for k, r in enumerate(rules) if r.columns and r.columns[0] == a), None)
        out.append(True if pos is None else res.last_fail > pos)
    return out


def db_return_item(nou: str, tori: str, youtc: str, zai: str, tyou: str, atu: float, haba: float, take: float, *,
                   additional_caption: str = "", mitsubishi_visible: bool = False, work_flag: str = "",
                   zenkasho_sheet: bool = True, lateral_lot: Optional[LotData] = None,
                   lateral_quiet: bool = False) -> ReturnItemResult:
    """QP.DbReturnItem (QP:1176-1456)。「ラテフリ直角」でフリー反り・ラテラルボー・直角度の要否と表示。

    atu/haba/take は VBA の Single 引数(製造の寸法)。additional_caption: いまの UF1.Additional.Caption。
    lateral_lot: 別表A/B のとき DBdisplayラテラル に渡すロット(UF1.Lot / UFMitsubishi.LotBox の Arrbox 結果)。
    """
    atu, haba, take = vbx.single(atu), vbx.single(haba), vbx.single(take)
    out = ReturnItemResult()
    u = out.uf1
    table = "ラテフリ直角"
    db: Path = _quality_db()
    pf = datasource.get_fields(db, table)
    if pf is None:
        # `Dim PackF(): PackF = GetFieldsArr(…)`(QP:1184)。動的配列に Empty を代入すると型エラー 13 なので、
        # PackingPath へ切り替える処理(QP:1196-1201)には到達しない(UF1.HURIS のリセットより前に止まる)
        raise vbx.TypeMismatch("ラテフリ直角 の列名が Empty (Dim PackF() への代入 QP:1184)")
    u["HURIS"] = ""
    C = {n: col_of(pf, n) for n in ("管理番号", "納入先", "取引先", "用途コード", "製造材質", "製造調質", "厚下",
                                    "厚上", "幅下", "幅上", "丈下", "丈上", "フリー反り", "ラテラルボー", "直角度",
                                    "製造板厚", "製造板幅", "製造板丈", "公差")}
    ta = datasource.get_records(db, table)
    if ta is None:
        return out
    args = (nou, tori, youtc, zai, tyou, atu, haba, take)
    g = match_rules.group(match_rules.RATEFURI)
    if g.new_way:
        matching, c = _ratefuri_new(g, ta, pf, C, args)
        if _shadow_on(g):
            match_rules.compare(match_rules.RATEFURI, _ratefuri_old(ta, C, args), (matching, c), repr(args))
    else:
        matching, c = _ratefuri_old(ta, C, args)

    additional = additional_caption
    if not matching:
        # 該当なし: 最後に評価した行の Conditions で文言を決める(QP:1276-1306)
        if vbx.like(nou, "*ﾆﾎﾝｼﾔﾘﾖｳｾｲｿﾞｳ*"):
            u["Sumisei.Caption"] = nou
            if not vbx.like(additional, "*ﾗﾃﾗﾙﾎﾞ-追記公差あり*"):
                if not c[0]:
                    text = "ﾗﾃﾗﾙﾎﾞｰ検査不要_製造材質"
                elif not c[1]:
                    text = "ﾗﾃﾗﾙﾎﾞｰ検査不要_ｻｲｽﾞ厚"
                elif not c[2]:
                    text = "ﾗﾃﾗﾙﾎﾞｰ検査不要_ｻｲｽﾞ幅"
                elif not c[3]:
                    text = "ﾗﾃﾗﾙﾎﾞｰ検査不要_ｻｲｽﾞ丈"
                else:
                    text = ""
                if text:
                    u["Additional.Caption"] = text
                    u["LateralText"] = ""
                    additional = text
        if vbx.like(tori, "ﾅﾒｶﾜｱﾙﾐ(ｶ*"):
            u["Sumisei.Caption"] = nou
            if vbx.like(nou, "*K.T.N. CO. LTD*"):
                u["Additional.Caption"] = "K.T.N ﾗﾃﾗﾙﾎﾞｰ検査不要"
                u["LateralText"] = ""
            if vbx.like(nou, "ﾆﾎﾝﾊﾂｼﾞﾖｳ*"):
                u["Additional.Caption"] = "ﾆﾎﾝﾊﾂｼﾞﾖｳ ﾗﾃﾗﾙﾎﾞｰ検査不要"
                u["LateralText"] = ""
        return out

    hit = _hit_rows(ta, C["管理番号"], matching)
    items = ""
    huris_set = False
    for h in hit:
        if veq(cell(h, C["フリー反り"]), "要", VAR, STR):
            items += ",フリー反り"
            # B72(QP:1337): 再現は「要」の最後の行の公差が残る。直すと「要」の最初の行(下の kousa_f と同じ)
            if not huris_set:
                u["HURIS"] = cell(h, C["公差"])
                huris_set = True
        if veq(cell(h, C["ラテラルボー"]), "要", VAR, STR):
            items += ",ラテラルボー"
        if veq(cell(h, C["直角度"]), "要", VAR, STR):
            items += ",直角度"
    out.items = items
    kousa1 = cell(hit[0], C["公差"])
    # B72: ラテラルボーの公差も HitArry(1)(最初に一致した行)から取る(UF1.LateralText QP:1355-1362、
    # UFMitsubishi のシート QP:1425-1430)。直した場合はラテラルボーが「要」の最初の行
    kousa_l = kousa1
    for h in hit:
        if veq(cell(h, C["ラテラルボー"]), "要", VAR, STR):
            kousa_l = cell(h, C["公差"])
            break

    # B72: フリー反りの公差(HizCome.Tag・HURIS。QP:1394-1395・1407)も HitArry(1) から取る。直した場合は
    # フリー反りが「要」の最初の行(ループ中の UF1.HURIS(QP:1337)も同じ行にそろえる)
    kousa_f = kousa1
    for h in hit:
        if veq(cell(h, C["フリー反り"]), "要", VAR, STR):
            kousa_f = cell(h, C["公差"])
            break

    def lateral(text: str) -> str:
        return vbx.fmt(dbdisplay_lateral(lateral_lot, text, quiet=lateral_quiet, messages=out.messages), "0.0")

    if not mitsubishi_visible:
        if vbx.instr(items, "ラテラルボー") > 0:
            u["Sumisei.Caption"] = nou
            if not vbx.like(additional, "*ﾗﾃﾗﾙﾎﾞ-追記公差あり*"):
                if isnum(kousa_l):
                    u["LateralText"] = vbx.fmt(kousa_l, "0.0")
                elif veq(kousa_l, "別表B", VAR, STR):
                    u["LateralText"] = lateral("日車")
                    u["LateralText.BackColor"] = ORANGE
                elif veq(kousa_l, "別表A", VAR, STR):
                    u["LateralText"] = lateral("白銅滑川")
                    u["LateralText.BackColor"] = ORANGE
        if vbx.instr(items, "フリー反り") > 0:
            tag = vstr(kousa_f) + "mm以下"
            u["HizCome.Tag"] = tag
            u["flatnesText.BackColor"] = ORANGE
            u["flatnesText"] = "ﾌﾘｰｿﾘ要"
            for lg in ("LGS", "LGU", "LGC", "LGD", "LGE", "LGF"):
                u[lg + ".Caption"] = "ﾌﾘｰｿﾘ"
            out.cb_add_atu = atu
            u["HURIS"] = kousa_f
            u["HizCome.Caption"] = "ﾌﾘｰ反り" + tag
    if mitsubishi_visible:
        if not zenkasho_sheet:
            return out
        if work_flag == "P1":
            out.sheet = "全ヵ所検査"
        else:
            raise VbaRuntimeError(91, "ZenSh が Nothing(WorkFlag が P1 以外)")
        if vbx.instr(items, "ラテラルボー") > 0:
            out.cells[(1, 15)] = nou + zai
            if isnum(kousa_l):
                out.cells[(1, 13)] = vbx.fmt(kousa_l, "0.0")
            elif veq(kousa_l, "別表B", VAR, STR):
                out.cells[(1, 13)] = lateral("日車")
            elif veq(kousa_l, "別表A", VAR, STR):
                out.cells[(1, 13)] = lateral("白銅滑川")
        if vbx.instr(items, "フリー反り") > 0:
            out.cells[(6, 29)] = "ﾌﾘｰｿﾘ"
    return out


# ------------------------------------------------------------------
# DB を読み直すキャッシュ(判断事項 D11: ファイルの更新で読み直す)
# ------------------------------------------------------------------
class _FileCache:
    """(ファイル, 更新時刻) が同じ間は読んだ結果を使う。VBA はブックを開いている間ずっと初回の結果。"""

    def __init__(self) -> None:
        self._key: Optional[tuple[str, int]] = None
        self.value: Any = None
        self.failed = False

    @staticmethod
    def key_of(path: Path) -> tuple[str, int]:
        try:
            return str(path), path.stat().st_mtime_ns
        except OSError:
            return str(path), -1

    def valid(self, path: Path) -> bool:
        return self._key == self.key_of(path)

    def store(self, path: Path, value: Any, failed: bool = False) -> None:
        self._key = self.key_of(path)
        self.value = value
        self.failed = failed

    def clear(self) -> None:
        self._key = None
        self.value = None
        self.failed = False


_gaichu_cache = _FileCache()
_gauge_cache = _FileCache()
_pack_cache = _FileCache()
for _n, _c in (("gaichu", _gaichu_cache), ("gauge", _gauge_cache), ("pack", _pack_cache)):   # F3(2026-11-10)
    cache_registry.register(f"quality._{_n}_cache", _c.clear, lambda c=_c: 0 if c._key is None else 1)  # noqa: SLF001


def clear_caches() -> None:
    """QP.ClearMaterialsCache / ReloadGaichu / ReloadGauge の代わり(次の呼び出しで読み直す)。"""
    for c in (_gaichu_cache, _gauge_cache, _pack_cache):
        c.clear()


class _LoadError(Exception):
    def __init__(self, number: int, text: str) -> None:
        super().__init__(text)
        self.number = number
        self.text = text


VB_OBJECT_ERROR = -2147221504


def _load_table(db: Path, table: str, col_names: Sequence[str]) -> tuple[Optional[list[list[str]]], list[int]]:
    """QP.LoadTable (QP:2432-2471)。必須列の確認(最初の一致)と全行の読込。"""
    try:
        flds = datasource.get_fields(db, table)
    except datasource.TableMissing as e:
        raise _LoadError(-2147217865, str(e)) from e
    if flds is None:
        raise _LoadError(VB_OBJECT_ERROR + 1, "「" + table + "」のフィールドを取得できません")
    idx: list[int] = []
    for name in col_names:
        k = -1
        for i, f in enumerate(flds, start=1):
            if f == name:
                k = i
                break
        if k == -1:
            raise _LoadError(VB_OBJECT_ERROR + 2, "「" + table + "」に「" + name + "」フィールドがありません")
        idx.append(k)
    return datasource.get_records(db, table), idx


# ------------------------------------------------------------------
# 外注条件(QP L2250-2430)
# ------------------------------------------------------------------
_H_COLS = ("管理番号", "条件名", "戻り値", "優先順位", "有効")
_D_COLS = ("条件管理番号", "条件番号", "対象項目", "判定方法", "判定値", "有効")


def _ensure_gaichu(messages: list[Message]) -> Optional[tuple[Any, list[int], Any, list[int]]]:
    """QP.EnsureLoaded (QP:2396-2430)。失敗したらファイルが更新されるまで再試行しない(VBA は手動の ReloadGaichu まで)。"""
    db = _quality_db()
    if _gaichu_cache.valid(db):
        return None if _gaichu_cache.failed else _gaichu_cache.value
    step = ""
    try:
        step = "外注条件 読込"
        head, hidx = _load_table(db, "外注条件", _H_COLS)
        step = "外注条件明細 読込"
        det, didx = _load_table(db, "外注条件明細", _D_COLS)
    except _LoadError as e:
        _gaichu_cache.store(db, None, failed=True)
        messages.append(Message("外注条件テーブルを読み込めませんでした" + VBCRLF + "段階: " + step + VBCRLF
                                + "内容: " + str(e.number) + " " + e.text + VBCRLF + VBCRLF
                                + "外注梱包の判定は行われません。原因を解消後 ReloadGaichu を実行してください。",
                                ERROR, source="QP:2423"))
        return None
    value = (head, hidx, det, didx)
    _gaichu_cache.store(db, value)
    return value


def _judge_one(cell_val: Any, method: str, judge_val: str) -> bool:
    """QP.JudgeOne (QP:2317-2339)。= / Like / 含む / >= / <=。"""
    s = vstr(cell_val)
    if method == "=":
        if vbx.is_numeric(s) and vbx.is_numeric(judge_val):
            return vbx.vround(vbx.cdbl(s), 3) == vbx.vround(vbx.cdbl(judge_val), 3)
        return s == judge_val
    if method == "Like":
        return vbx.like(s, judge_val)
    if method == "含む":
        return len(judge_val) > 0 and vbx.instr(s, judge_val) > 0
    if method == ">=":
        return vbx.is_numeric(s) and vbx.is_numeric(judge_val) and vbx.cdbl(s) >= vbx.cdbl(judge_val)
    if method == "<=":
        return vbx.is_numeric(s) and vbx.is_numeric(judge_val) and vbx.cdbl(s) <= vbx.cdbl(judge_val)
    return False


def gaichu_packing_no(hist: Sequence[Sequence[Any]], lot_no: str, r_lot: int,
                      messages: Optional[list[Message]] = None) -> str:
    """QP.GaichuPackingNo (QP:2250-2269)。外注条件に該当すれば包装仕様書NO、無ければ ""。

    hist: Lot情報履歴シートの値(1行目が見出し。セルの値は Excel のまま: 数は float、空は None)。
    r_lot: ﾛｯﾄNo の列番号(1 始まり)。
    - 有効な条件セットごとに、明細の条件番号どうしは AND、同じ番号の中は OR。明細が無いセットは不成立
    - ﾛｯﾄNo が一致する行のどれか1行で成立すればよい(見出し行も対象)
    - 成立したセットのうち優先順位が最小(同値なら先の行)の戻り値
    """
    msgs = messages if messages is not None else []
    loaded = _ensure_gaichu(msgs)
    if loaded is None:
        return ""
    head, hidx, det, didx = loaded
    if head is None or det is None:
        return ""
    h_no, h_ret, h_pri, h_valid = hidx[0], hidx[2], hidx[3], hidx[4]
    d_hno, d_cno, d_tgt, d_mtd, d_val, d_valid = didx

    def find_hist_col(target: str) -> int:
        for i, v in enumerate(hist[0], start=1):
            if vstr(v) == target:
                return i
        return 0

    def get_cond_nos(head_no: float) -> Optional[list[float]]:
        nos: list[float] = []
        for d in det:
            if vbx.cdbl(cell(d, d_hno)) == head_no and cbool(cell(d, d_valid)):
                v = vbx.cdbl(cell(d, d_cno))
                if v not in nos:
                    nos.append(v)
        return nos or None

    def match_cond_no(head_no: float, cond_no: float) -> bool:
        for d in det:
            if vbx.cdbl(cell(d, d_hno)) == head_no and vbx.cdbl(cell(d, d_cno)) == cond_no \
                    and cbool(cell(d, d_valid)):
                col = find_hist_col(vstr(cell(d, d_tgt)))
                if col > 0:
                    for row in hist:
                        if vstr(cell(row, r_lot)) == lot_no:
                            if _judge_one(cell(row, col), vstr(cell(d, d_mtd)), vstr(cell(d, d_val))):
                                return True
        return False

    def match_cond_set(head_no: float) -> bool:
        nos = get_cond_nos(head_no)
        if nos is None:
            return False
        return all(match_cond_no(head_no, k) for k in nos)

    best_pri = 0.0
    hit = False
    ret = ""
    for h in head:
        if cbool(cell(h, h_valid)):
            head_no = vbx.cdbl(cell(h, h_no))
            if match_cond_set(head_no):
                if not hit or vbx.cdbl(cell(h, h_pri)) < best_pri:
                    hit = True
                    best_pri = vbx.cdbl(cell(h, h_pri))
                    ret = vstr(cell(h, h_ret))
    return ret


# ------------------------------------------------------------------
# ゲージ判定(QP L2473-2552)
# ------------------------------------------------------------------
_G_COLS = ("有効", "重量_超", "丈_以上", "丈_未満", "丈_超", "幅_未満", "材質", "調質", "コース含む")


def _ensure_gauge(messages: list[Message]) -> Optional[tuple[Any, list[int]]]:
    """QP.EnsureLoadedGauge (QP:2528-2552)。"""
    db = _quality_db()
    if _gauge_cache.valid(db):
        return None if _gauge_cache.failed else _gauge_cache.value
    try:
        rule, idx = _load_table(db, "ゲージ判定条件", _G_COLS)
    except _LoadError as e:
        _gauge_cache.store(db, None, failed=True)
        messages.append(Message("ゲージ判定条件を読み込めませんでした" + VBCRLF + "内容: " + str(e.number) + " "
                                + e.text + VBCRLF + VBCRLF
                                + "定盤表示になります。原因を解消後 ReloadGauge を実行してください。",
                                ERROR, source="QP:2548"))
        return None
    _gauge_cache.store(db, (rule, idx))
    return rule, idx


def _has_val(v: Any) -> bool:
    """QP.HasVal (QP:2514-2518)。"""
    if v is None:
        return False
    return len(vtrim(vstr(v)).strip(" ")) > 0


def is_gauge_required(atu: float, haba: float, take: float, zai: str, tyo: str, course: str,
                      messages: Optional[list[Message]] = None) -> bool:
    """QP.IsGaugeRequired (QP:2473-2511)。True = ゲージ検査、False = 定盤検査。

    有効な行ごとに、値の入っている条件をすべて満たせば成立(1行でも成立で True)。
    重量 = Weight(厚,幅,丈)、重量_超(>)・丈_以上(>=)・丈_未満(<)・丈_超(>)・幅_未満(<)・材質(=)・調質(=)・コース含む。
    """
    msgs = messages if messages is not None else []
    loaded = _ensure_gauge(msgs)
    if loaded is None:
        return False
    rule, idx = loaded
    if rule is None:
        return False
    w = weight(atu, haba, take)
    for r in rule:
        if not cbool(cell(r, idx[0])):
            continue
        ok = True
        n = 0
        checks: list[tuple[int, Callable[[Any], bool]]] = [
            (1, lambda v: w > vbx.cdbl(v)),
            (2, lambda v: take >= vbx.cdbl(v)),
            (3, lambda v: take < vbx.cdbl(v)),
            (4, lambda v: take > vbx.cdbl(v)),
            (5, lambda v: haba < vbx.cdbl(v)),
            (6, lambda v: zai == vstr(v)),
            (7, lambda v: tyo == vstr(v)),
            (8, lambda v: vbx.instr(course, vstr(v)) > 0),
        ]
        for k, test in checks:
            v = cell(r, idx[k])
            if _has_val(v):
                n += 1
                if not test(v):
                    ok = False
        if ok and n > 0:
            return True
    return False


# ------------------------------------------------------------------
# 梱包保護材(QP L1618-1783)
# ------------------------------------------------------------------
C_TARGET_HOSONOS = "1P0001,1P0104,1P0110,1P0115,1P0116,1P0118,1P0119,1P0120"


def is_target_hoso_no(hoso_no: str) -> bool:
    """QP.IsTargetHosoNo (QP:1618-1631)。蓋アングルの絡む包装仕様書No か。"""
    return any(vtrim(t) == hoso_no for t in C_TARGET_HOSONOS.split(","))


_PACK_COLS = ("管理番号", "包装仕様書", "用途コード", "材質", "調質", "板厚下", "板厚上", "板厚", "板幅下", "板幅上",
              "板幅", "板丈下", "板丈上", "板丈", "使用保護材", "備考")


def materials_upper_part(ord_no: str, hist: Optional[Sequence[Sequence[Any]]], check_bikou: bool = False,
                         db: Union[str, Path, None] = None) -> str:
    """QP.MaterialsUpperPart (QP:1674-1783)。上部保護材(梱包資材マスタ「梱包保護材」)。

    hist: Lot情報履歴シートの値(見出し行を含む。None はシートが無い)。
    - [ｵｰﾀﾞｰNo] = OrdNo("'" を除く)の最初の行(見出し行から探す)
    - 包装仕様NO が対象外なら ""
    - 備考="当板用" が check_bikou と同じ行だけ、包装仕様書・材質・調質・用途・厚/幅/丈で照合し、最初の一致の
      [使用保護材]。一致が無ければ "一致なし"
    キャッシュ: VBA はオーダーNo が同じ間は読み直さない。ここではファイルの更新で読み直す(D11)。
    """
    if hist is None or ord_no == "":
        return ""
    if not hist:
        return ""
    hdr = hist[0]
    names = {n: 0 for n in ("ﾛｯﾄNo", "ｵｰﾀﾞｰNo", "引当NO", "包装仕様NO", "厚", "巾", "丈", "材質", "調質", "用途ｺｰﾄﾞ")}
    for i, v in enumerate(hdr, start=1):
        if vstr(v) in names:
            names[vstr(v)] = i
    clean = ord_no.replace("'", "")
    a: Optional[Sequence[Any]] = None
    for row in hist:
        if veq(cell(row, names["ｵｰﾀﾞｰNo"]), clean, VAR, STR):
            a = row
            break
    if a is None:
        return ""
    pac = vstr(cell(a, names["包装仕様NO"]))
    if not is_target_hoso_no(pac):
        return ""
    you = vstr(cell(a, names["用途ｺｰﾄﾞ"]))
    zai = vstr(cell(a, names["材質"]))
    tyo = vstr(cell(a, names["調質"]))
    atu = csng(cell(a, names["厚"]))
    hab = csng(cell(a, names["巾"]))
    tak = csng(cell(a, names["丈"]))
    pack = vstr(cell(a, names["包装仕様NO"]))

    path = datasource.resolve(db) if db is not None else _packing_db("梱包資材マスタ.sqlite3")
    if _pack_cache.valid(path) and _pack_cache.value is not None:
        C, rows, pf = _pack_cache.value
    else:
        pf = datasource.get_fields(path, "梱包保護材")
        if pf is None:
            # 不明点: VBA の EnsureCache は `Dim PackF(): PackF = GetFieldsArr(…)`(QP:1641)で、表が空だと
            # 代入で型エラー 13 になるはず。梱包資材マスタが未受領で空の表のまま動かしているため、ここは "" で返す
            return ""
        C = {n: col_of(pf, n) for n in _PACK_COLS}
        rows = datasource.get_records(path, "梱包保護材")
        if rows is None:
            # `m_PackQuality() As Variant`(動的配列)に Empty を入れると型エラー
            raise vbx.TypeMismatch("梱包保護材が 0 件(動的配列に Empty を代入)")
        _pack_cache.store(path, (C, rows, list(pf)))
    vals = (pack, zai, tyo, you, atu, hab, tak, check_bikou)
    g = match_rules.group(match_rules.PACK)
    if not g.new_way:
        return _pack_old(rows, C, vals)
    ctx = pack_ctx(hdr, a, check_bikou)
    res = match_rules.match(g, rows, pf, ctx)
    got = vstr(cell(rows[res.rows[0]], C["使用保護材"])) if res.rows else "一致なし"
    if _shadow_on(g):
        match_rules.compare(match_rules.PACK, _pack_old(rows, C, vals), got, repr(vals))
    return got


def pack_ctx(hdr: Sequence[Any], a: Sequence[Any], check_bikou: bool) -> dict[str, Any]:
    """梱包保護材のロットの項目(Lot情報履歴 の見出し → 値。厚・巾・丈 は CSng。+ 当板用か)。"""
    ctx: dict[str, Any] = {}
    for i, h in enumerate(hdr, start=1):
        name = vstr(h)
        if name and name not in ctx:
            ctx[name] = cell(a, i)
    for n in ("厚", "巾", "丈"):
        if n in ctx:
            ctx[n] = csng(ctx[n])
    ctx["当板用"] = "当板用" if check_bikou else ""
    return ctx


def _pack_old(rows: Sequence[Sequence[Any]], C: dict[str, int], vals: tuple) -> str:
    """旧: MaterialsUpperPart の照合(QP:1750-1770)をコードで。照合ルールの shadow で差が 0 件と確かめたら消す。"""
    pack, zai, tyo, you, atu, hab, tak, check_bikou = vals
    same = True          # B71(QP:1764-1766): 直すと下・上も Double で比べる
    for r in rows:
        is_bikou = vstr(cell(r, C["備考"])) == "当板用"
        if is_bikou != check_bikou:
            continue
        c1 = check_condition(r, C["包装仕様書"], pack)
        c2 = c1 & check_condition(r, C["材質"], zai)
        c3 = c2 & check_condition(r, C["調質"], tyo)
        c4 = c3 & check_condition(r, C["用途コード"], you)
        c5 = c4 & check_thickness(r, C["板厚下"], C["板厚上"], C["板厚"], atu, same_precision=same)
        c6 = c5 & check_width(r, C["板幅下"], C["板幅上"], C["板幅"], hab, same_precision=same)
        c7 = c6 & check_width(r, C["板丈下"], C["板丈上"], C["板丈"], tak, same_precision=same)
        if c7:
            return vstr(cell(r, C["使用保護材"]))
    return "一致なし"


# ------------------------------------------------------------------
# 伝送用ファイル: 時間用・校正ストレートエッヂ・設備仕様
# ------------------------------------------------------------------
def _transfer_fields(table: str) -> tuple[Optional[Path], Optional[list[str]]]:
    """QualityPath の伝送用ファイルの列名。

    VBA は `Dim PackF(): PackF = GetFieldsArr(…)`(QP:1792・1837・2091)で、表が空(Empty)だと動的配列への
    代入で型エラー 13 になる。そのため PackingPath へ切り替える処理には到達しない(MachineTransfer と同じ)。
    """
    db = _transfer_db()
    f = datasource.get_fields(db, table)
    if f is None:
        raise vbx.TypeMismatch(f"{table} の列名が Empty (Dim PackF() への代入)")
    return db, f


def shift_time() -> Any:
    """QP.ShiftTime (QP:1786-1829)。「時間用」の**先頭4行**の (直, 開始, 終了)。

    表が空なら型エラー 13(`_transfer_fields`)、行が無ければ ""。4行未満は VBA と同じく添字エラー。
    """
    db, f = _transfer_fields("時間用")
    if f is None:
        return None
    c_s, c_e, c_c = col_of(f, "開始"), col_of(f, "終了"), col_of(f, "直")
    rows = datasource.get_records(db, "時間用")
    if rows is None:
        return ""
    if len(rows) < 4:
        raise VbaRuntimeError(9, "時間用が4行未満")
    return [(cell(r, c_c), cell(r, c_s), cell(r, c_e)) for r in rows[:4]]


def straight_edge(line_name: str) -> Any:
    """QP.StraightEdge (QP:1831-1864)。ライン名の行の (先頭行の長さ下, 最終行の長さ上)。

    ライン名の照合は D13(文字の幅だけ揃える `lines.matches`)。表が空なら型エラー 13、行が無ければ ""。
    """
    db, f = _transfer_fields("校正ストレートエッヂ")
    if f is None:
        return None
    c_l, c_lo, c_hi = col_of(f, "ライン名"), col_of(f, "長さ下"), col_of(f, "長さ上")
    rows = [r for r in (datasource.get_records(db, "校正ストレートエッヂ") or [])
            if lines.matches(cell(r, c_l), line_name)]
    if not rows:
        return ""
    return cell(rows[0], c_lo), cell(rows[-1], c_hi)


def check_equipment_db(line: str, zai: str, atu: float, haba: float, take: float, wei: float) -> str:
    """QP.CheckEquipmentDb (QP:2084-2190)。設備仕様に合うか。

    - 設備名(D13: 文字の幅だけ揃えて照合)の行が無ければ "可"
    - 製造材質(CheckCondition)と、厚・幅・丈・単重の MIN–MAX(CheckWithin)
    - 一致なし: **最後に評価した行**で "オーダー厚 範囲外" / "C_オーダー幅 範囲外" / "オーダー丈 範囲外" /
      "単重 範囲外"(材質不一致も "オーダー厚 範囲外"。B27。直すと "製造材質 範囲外"・"オーダー幅 範囲外")
    - 一致あり: 空でない [作業可否] の最後のもの(全部空なら "")
    """
    atu, haba, take, wei = (vbx.single(x) for x in (atu, haba, take, wei))
    db, f = _transfer_fields("設備仕様")
    if f is None:
        return ""
    C = {n: col_of(f, n) for n in ("管理番号", "設備名", "製造材質", "オーダー厚MIN", "オーダー厚MAX", "オーダー幅MIN",
                                   "オーダー幅MAX", "オーダー丈MIN", "オーダー丈MAX", "単重MIN", "単重MAX", "作業可否")}
    rows = [r for r in (datasource.get_records(db, "設備仕様") or []) if lines.matches(cell(r, C["設備名"]), line)]
    if not rows:
        return "可"
    matching: list[int] = []
    c = [False] * 6
    same = True          # B71(QP:2140-2143): 直すと MIN・MAX も Double で比べる
    for r in rows:
        c = [False] * 6
        c[1] = check_condition(r, C["製造材質"], zai)
        c[2] = c[1] & check_within(r, C["オーダー厚MIN"], C["オーダー厚MAX"], atu, same_precision=same)
        c[3] = c[2] & check_within(r, C["オーダー幅MIN"], C["オーダー幅MAX"], haba, same_precision=same)
        c[4] = c[3] & check_within(r, C["オーダー丈MIN"], C["オーダー丈MAX"], take, same_precision=same)
        c[5] = c[4] & check_within(r, C["単重MIN"], C["単重MAX"], wei, same_precision=same)
        if all(c[1:6]):
            matching.append(clng(cell(r, C["管理番号"])))
    if not matching:
        if not c[1]:
            return "製造材質 範囲外"
        if not c[2]:
            return "オーダー厚 範囲外"
        if not c[3]:
            return "オーダー幅 範囲外"
        if not c[4]:
            return "オーダー丈 範囲外"
        if not c[5]:
            return "単重 範囲外"
    out = ""
    for h in _hit_rows(rows, C["管理番号"], matching):
        if vstr(cell(h, C["作業可否"])) != "":
            out = vstr(cell(h, C["作業可否"]))
    return out


# ------------------------------------------------------------------
# バーコール(QDB L1454-1490)
# ------------------------------------------------------------------
def bh_search(zai: str, tyo: str) -> int:
    """QDB.BhSearch (QDB:1454-1490)。バーコール下限値。材質は4文字以上なら右3文字で LIKE '*材質*'。

    調質が完全一致する**最後の行**の値(Long)。無ければ 0。
    """
    db = _quality_db()
    f = datasource.get_fields(db, "Q_S_KE_ALL_053")
    if f is None:
        return 0
    c_tyo, c_bh = col_of(f, "調質"), col_of(f, "バーコール下限値")
    if len(zai) >= 4:
        zai = vbx.right(zai, 3)
    rows = datasource.get_records(db, "Q_S_KE_ALL_053", "[材質] LIKE '*" + zai + "*'")
    if rows is None:
        return 0
    out = 0
    for r in rows:
        if vstr(cell(r, c_tyo)) == tyo:
            out = clng(cell(r, c_bh))
    return out
