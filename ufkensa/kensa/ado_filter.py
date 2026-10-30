"""ADO の `Recordset.Filter` と同じ規則で行を絞る

VBA は `SELECT * FROM 表` で全件を取り、`Recordset.Filter = "[列]='値' AND …"` で絞っていた
(SQL_追加 L293-296)。SQLite に移しても**同じ行が残る**よう、絞り込みは Python でこの規則で行う。

対応する書き方(VBA に出てくるもの全部):
    [列]='値'    [列]<>'値'   [列]=数   [列]<>数   [列] LIKE '*文字*'
    条件 AND 条件 / 条件 OR 条件(括弧も可)

【比較の規則】(実機確認 U3 で変わりうる所は関数を分けてある)
- 文字列の比較は**大文字小文字を区別しない**(ADO Filter の仕様)
- 全角/半角は区別する(ADO は区別する見込み。ライン名だけは `lines.same` で揃えて呼ぶ。D13)
- 末尾の空白はそのまま比べる(B20: `'値 '` は一致しない)
- `LIKE` のワイルドカードは `*` と `%`(ADO は前後にだけ書ける)
- 数との比較: 列の値が数として読めれば数で、読めなければ一致しない(Text 列 × 数。U3)
- Null は空文字と同じに扱う(GetString の NullExpr が "")。`= ''` は Null の行にも一致する(U3)
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable, Optional, Union

Row = dict[str, Optional[str]]

_TOKEN = re.compile(r"""
    \s*(?:
      (?P<lp>\()|(?P<rp>\))|
      (?P<and>AND\b)|(?P<or>OR\b)|
      \[(?P<col>[^\]]+)\]|
      (?P<op><>|<=|>=|=|<|>|LIKE\b)|
      '(?P<str>(?:[^']|'')*)'|
      \#(?P<date>[^#]*)\#|
      (?P<num>[+-]?(?:\d+\.?\d*|\.\d+))|
      (?P<bare>[^\s\[\]()'=<>]+)
    )""", re.VERBOSE | re.IGNORECASE)


class FilterSyntaxError(ValueError):
    pass


@dataclass
class _Cond:
    col: str
    op: str
    value: Union[str, float]
    is_num: bool


def _tokens(expr: str) -> list[tuple[str, str]]:
    out: list[tuple[str, str]] = []
    pos = 0
    expr = expr.rstrip()
    while pos < len(expr):
        m = _TOKEN.match(expr, pos)
        if not m or m.end() == pos:
            raise FilterSyntaxError(f"読めない Filter: {expr!r} の {pos} 文字目")
        kind = m.lastgroup
        text = m.group(kind)
        if kind == "str":
            text = text.replace("''", "'")
        out.append((kind, text))
        pos = m.end()
    return out


def _norm_case(s: str) -> str:
    return s.casefold()


def _as_number(text: Optional[str]) -> Optional[float]:
    if text is None:
        return None
    t = str(text).strip()
    if not t:
        return None
    try:
        return float(t.replace(",", ""))
    except ValueError:
        return None


def _match(cond: _Cond, row: Row) -> bool:
    if cond.col not in row:
        raise KeyError(f"Filter の列 [{cond.col}] が表にありません")
    raw = row[cond.col]
    cell = "" if raw is None else str(raw)
    op = cond.op.upper()
    if op == "LIKE":
        pat = str(cond.value)
        core = pat.strip("*%")
        c, p = _norm_case(cell), _norm_case(core)
        head = pat[:1] in ("*", "%")
        tail = pat[-1:] in ("*", "%") and len(pat) > 1
        if head and tail:
            return p in c
        if head:
            return c.endswith(p)
        if tail:
            return c.startswith(p)
        return c == p
    if cond.is_num:
        n = _as_number(cell)
        if n is None:
            return op == "<>"
        v = float(cond.value)
        return {"=": n == v, "<>": n != v, "<": n < v, ">": n > v, "<=": n <= v, ">=": n >= v}[op]
    a, b = _norm_case(cell), _norm_case(str(cond.value))
    return {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[op]


def compile_filter(expr: str) -> Callable[[Row], bool]:
    """Filter の文字列 → 行を受け取って真偽を返す関数。空なら全行が真。"""
    if not expr or not expr.strip():
        return lambda row: True
    toks = _tokens(expr)
    pos = 0

    def peek() -> Optional[tuple[str, str]]:
        return toks[pos] if pos < len(toks) else None

    def take() -> tuple[str, str]:
        nonlocal pos
        t = toks[pos]
        pos += 1
        return t

    def cond() -> Callable[[Row], bool]:
        t = peek()
        if t and t[0] == "lp":
            take()
            f = disj()
            if not peek() or peek()[0] != "rp":
                raise FilterSyntaxError(f"括弧が閉じていない: {expr!r}")
            take()
            return f
        k, col = take()
        if k != "col":
            raise FilterSyntaxError(f"列名 [..] が要る所に {col!r}: {expr!r}")
        k, op = take()
        if k != "op":
            raise FilterSyntaxError(f"比較の記号が要る所に {op!r}: {expr!r}")
        k, v = take()
        if k == "str":
            c = _Cond(col, op, v, False)
        elif k == "num":
            c = _Cond(col, op, float(v), True)
        elif k == "date":
            c = _Cond(col, op, v, False)
        else:
            raise FilterSyntaxError(f"値が要る所に {v!r}: {expr!r}")
        return lambda row, c=c: _match(c, row)

    def conj() -> Callable[[Row], bool]:
        parts = [cond()]
        while peek() and peek()[0] == "and":
            take()
            parts.append(cond())
        return lambda row: all(p(row) for p in parts)

    def disj() -> Callable[[Row], bool]:
        parts = [conj()]
        while peek() and peek()[0] == "or":
            take()
            parts.append(conj())
        return lambda row: any(p(row) for p in parts)

    f = disj()
    if pos != len(toks):
        raise FilterSyntaxError(f"余分な語: {toks[pos:]!r} in {expr!r}")
    return f


def quote(value: object) -> str:
    """VBA で `"[列]='" & 値 & "'"` と組み立てていた所の値(エスケープはしない。VBA と同じ)。"""
    return "" if value is None else str(value)
