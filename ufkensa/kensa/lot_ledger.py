"""仕掛台帳(SIKALOT / SIKAHIKI / SIKAODR)の第1/第2の読み分け(リファクタリング 第2段。**ここ 1 か所**)

前は読み手ごと(arrbox・acc_reflec・履歴保存・SikaAtuCreate・厚の作業選択・grid の既定・judge・lot_watch・
品質基準の引当)に第1/第2の切り替えを手で書いていて、オフラインで第1を見る所と見ない所が混ざっていた。
読み手は全部ここを通す。決まり:

| 読み方                                   | まず読む側(primary) | 次に見る側(other) |
|------------------------------------------|----------------------|--------------------|
| オンライン                               | 第1                  | 第2                 |
| オフライン・「控えから読む」(lot_source) | 第2                  | 第1                 |
| オンラインで DBflag あり(VBA)          | 第2(NextPath)だけ   | なし(VBA どおり)   |

1. ファイル: まず読む側のファイルが読めない(無い・手元へ写せない)→ 次に見る側のファイル。どちらも駄目なら
   「…がありません」(`paths.missing_text`。文言は呼び手がそれぞれの VBA の出し方で出す)
2. ロット: まず読む側のファイルに条件の行(ロット)が無い → 次に見る側の同じ条件の行(列の並びはまず読む側に直す)
3. 列名: まず読む側の表が空 → 次に見る側の列名
4. 引当(SIKAHIKI): まず読む側にロットの行があっても、次に見る側にだけある引当番号の行を後ろに足す
   (知らせ `hiki_added_message`)
5. オーダー(SIKAODR): 引当の受注番号のうち、まず読む側に無いものを次に見る側から足す
6. 列(第2の台帳に無い列。公差・品質ｸﾞﾚｰﾄﾞ): 第1の同じロットの行で補う(`datasource.fill_from_first`。
   表を読む所で済ませて覚えるので、ここでは何もしない)

- 第1と第2が同じフォルダ(試験・1 か所に置いたとき)は次に見る側なし
- 次に見る側へ移るのは 1 回だけ(ファイルで移ったら、ロットではもう移らない)
- 「ファイルがあるか」はどの読み手も `readable`(= `paths.readable`。仕掛 DB は手元へ写せたときだけ真)
- DBflag(VBA)のときはファイルを確かめずに読む(無ければ `datasource.DbMissing`。VBA の End)。
  確かめてから「…がありません」を出す読み手は `check_explicit=True`
- 名前で引く DB の場所(`paths.db_file("SIKALOTNOW")` など)も 1. のファイルの決まり(`lot_file`)

覚え(キャッシュ)はここでは持たない。表は `datasource` が覚えた形(`_lot_ready`・`_eq_index`)をそのまま使う。
"""
from __future__ import annotations

import sqlite3
from dataclasses import dataclass, field
from typing import Any, Callable, Iterable, Optional, Sequence

from . import ado_filter, datasource, paths, vbx
from .flow import INFO, Message

TABLE = "仕掛"
LOT, HIKI, ODR = paths.LOT_DBS              # "SIKALOTNOW", "SIKAHIKINOW", "SIKAODRNOW"
HIKI_NO, ORDER_NO = "引当番号", "受注番号"

_READ_ERRORS = (datasource.DbMissing, datasource.TableMissing, sqlite3.Error)


# ------------------------------------------------------------------
# どちらの側を読むか・ファイル
# ------------------------------------------------------------------
@dataclass(frozen=True)
class Sides:
    primary: str             # まず読む側のフォルダ
    other: str = ""          # 次に見る側のフォルダ("" = なし)
    explicit: bool = False   # DBflag: NextPath だけ(VBA どおり確かめずに読む)
    vba_main: bool = False   # VBA の FPath = MainPath(第1)で、NextPath が別の場所(VBA の文言の出し分けに使う)


def sides(dbflag: str = "", m: Optional[str] = None) -> Sides:
    """いまの読み方の側。`dbflag` は UF1.DBflag(空以外で VBA の NextPath)。"""
    main, nxt = paths.get("main", m), paths.get("next", m)
    if dbflag != "" and nxt != main:
        return Sides(nxt, "", explicit=True)
    if paths.lot_mode(m) == paths.ONLINE:
        other = nxt
    else:
        other = paths.slot_value(paths.LOT, paths.FIRST)
    other = "" if (not other or other == main) else other
    return Sides(main, other, vba_main=(main != nxt))


def other_base(dbflag: str = "") -> str:
    """次に見る側のフォルダ(無ければ "")。"""
    return sides(dbflag).other


def db_path(base: str, name: str, m: Optional[str] = None) -> str:
    """VBA の `FPath & "SIKA….accdb"`(D4 でファイルは .sqlite3)。

    D26: 仕掛の 3 ファイルは、第2(控え)のフォルダなら設定「2つ目の名前」で探す(`paths.lot_file_name`)。
    """
    if name in paths.LOT_DBS:
        return paths.join(base, paths.lot_file_name(name, base, m))
    return paths.join(base, paths.DB_FILES[name][1] if name in paths.DB_FILES else name + ".sqlite3")


def readable(path_text: Any) -> bool:
    """ファイルを読めるか(どの読み手もこれ)。仕掛 DB は手元へ写せたときだけ真。"""
    return paths.readable(path_text)


@dataclass(frozen=True)
class Pick:
    file: str                # 読むファイル(見つからなければ最後に見たファイル)
    found: bool              # 読めるファイルがあった
    on_primary: bool         # まず読む側のファイル


def pick(target: str, dbflag: str = "", *, check_explicit: bool = False,
         check: Optional[Callable[[str], bool]] = None, m: Optional[str] = None) -> Pick:
    """決まり 1(ファイル)。"""
    ok = check or readable
    s = sides(dbflag, m)
    f = db_path(s.primary, target, m)
    if s.explicit:
        return Pick(f, ok(f) if check_explicit else True, True)
    if ok(f):
        return Pick(f, True, True)
    if s.other:
        f2 = db_path(s.other, target, m)
        return Pick(f2, ok(f2), False)
    return Pick(f, False, True)


def lot_file(name: str, m: Optional[str] = None) -> str:
    """名前で引く仕掛 DB の場所(`paths.db_file`)。読めるファイルが無ければ、まず読む側のファイル。"""
    p = pick(name, m=m)
    return p.file if p.found else db_path(sides(m=m).primary, name, m)


# ------------------------------------------------------------------
# 読む
# ------------------------------------------------------------------
@dataclass
class LotRead:
    """読んだ結果。`rows` / `cols` が None は VBA の Empty。"""
    rows: Optional[list] = None
    cols: Optional[list] = None
    file_path: str = ""          # 行を読んだファイル(ファイルが無ければ最後に見たファイル)
    found: bool = False          # 読めるファイルがあった
    on_primary: bool = True      # まず読む側のファイルを読んだ(偽 = ファイルで次に見る側へ移った)
    other: str = "none"          # 次に見る側: "none"(なし・もう移った)/ "unusable"(読めない)/ "tried"(見た)
    from_other: bool = False     # 行を次に見る側から取った(決まり 2・3)
    added: list[str] = field(default_factory=list)   # 足した引当番号 / 受注番号(決まり 4・5)

    def columns(self) -> Optional[list]:
        """行の列名(読んでいなければ読んだファイルの列名)。"""
        if self.cols is None and self.found:
            self.cols = _columns(self.file_path)
        return self.cols


def _columns(file_path: str) -> Optional[list]:
    """表の列名(0 件でも返す)。覚えた表から取るので読み直さない。"""
    try:
        return datasource.columns(file_path, TABLE)
    except _READ_ERRORS:
        return None


def other_file(target: str, dbflag: str = "") -> str:
    """次に見る側のファイル(無ければ "")。まず読む側のファイルを読んでいるときだけ使う。"""
    s = sides(dbflag)
    return db_path(s.other, target) if (s.other and not s.explicit) else ""


def _remap(rows: list, src: Optional[list], dst: Optional[list]) -> list:
    """次に見る側の行を、まず読む側の列の並びに直す(2つ目は列の数・並びが違うことがある。無い列は "")。"""
    if not src or not dst or list(src) == list(dst):
        return rows
    idx = [list(src).index(c) if c in src else -1 for c in dst]
    return [[r[i] if 0 <= i < len(r) else "" for i in idx] for r in rows]


def _other_rows(target: str, flt: str, primary_file: str, ofile: str) -> tuple[Optional[list], Optional[list]]:
    """次に見る側の条件の行を、まず読む側の列の並びで(行, 列名)。無い・読めなければ (None, None)。"""
    try:
        rows = datasource.get_records(ofile, TABLE, flt)
        if rows is None:
            return None, None
        src = datasource.get_fields(ofile, TABLE, "")
        dst = datasource.get_fields(primary_file, TABLE, "") if readable(primary_file) else None
    except _READ_ERRORS:
        return None, None
    return _remap(rows, src, dst), (dst or src)


def read_rows(target: str, flt: str = "", dbflag: str = "", *, merge: bool = True, check_explicit: bool = False,
              check: Optional[Callable[[str], bool]] = None) -> LotRead:
    """決まり 1・2・4。`flt` は ADO の Filter(`arrbox_filter` など)。`merge` が偽なら引当を足さない。"""
    p = pick(target, dbflag, check_explicit=check_explicit, check=check)
    r = LotRead(file_path=p.file, found=p.found, on_primary=p.on_primary)
    if not p.found:
        return r
    r.rows = datasource.get_records(p.file, TABLE, flt)
    other = other_file(target, dbflag) if p.on_primary else ""
    if not other:
        return r
    if r.rows is None:
        if not (check or readable)(other):
            r.other = "unusable"
            return r
        r.other = "tried"
        rows, cols = _other_rows(target, flt, p.file, other)
        if rows is not None:
            r.rows, r.cols, r.file_path, r.from_other = rows, cols, other, True
        return r
    if merge and target == HIKI:
        _merge_hiki(r, flt, other)
    return r


def _merge_hiki(r: LotRead, flt: str, other: str) -> None:
    """決まり 4(2026-11-08 ご依頼): 次に見る側にだけある引当番号の行を後ろに足す(まず読む側の行はそのまま)。"""
    cols = r.columns() or []
    i = datasource.column_index(cols, HIKI_NO)
    if i < 0 or not r.rows or not readable(other):
        return
    extra, _cols = _other_rows(HIKI, flt, r.file_path, other)
    have = {vbx.cstr(x[i]) for x in r.rows if i < len(x)}
    add = [x for x in (extra or []) if i < len(x) and vbx.cstr(x[i]) not in have]
    if not add:
        return
    r.added = sorted({vbx.cstr(x[i]) for x in add})
    r.rows = list(r.rows) + add
    from ..logging_utils import get_logger
    get_logger("kensa.lot").info("%sに無い引当を%sから足しました: %s (%s)", *_here_there(), ", ".join(r.added), flt)


def read_fields(target: str, dbflag: str = "", *, check_explicit: bool = False) -> LotRead:
    """決まり 1・3。VBA の GetFieldsArr(表が空なら None)。"""
    p = pick(target, dbflag, check_explicit=check_explicit)
    r = LotRead(file_path=p.file, found=p.found, on_primary=p.on_primary)
    if not p.found:
        return r
    r.cols = datasource.get_fields(p.file, TABLE, "")
    other = other_file(target, dbflag) if p.on_primary else ""
    if r.cols is None and other:
        if not readable(other):
            r.other = "unusable"
            return r
        r.other = "tried"
        try:
            cols = datasource.get_fields(other, TABLE, "")
        except _READ_ERRORS:
            cols = None
        if cols is not None:
            r.cols, r.file_path, r.from_other = cols, other, True
    return r


def read_orders(orders: Iterable[str], dbflag: str = "", *, check_explicit: bool = False) -> LotRead:
    """SIKAODR の全行(決まり 1・2)に、まず読む側に無い受注番号の行を次に見る側から足す(決まり 5)。"""
    r = read_rows(ODR, "", dbflag, merge=False, check_explicit=check_explicit)
    want = {o for o in orders if o is not None}
    if not r.rows or r.from_other or not r.on_primary or not want:
        return r
    other = other_file(ODR, dbflag)
    cols = r.columns() or []
    j = datasource.column_index(cols, ORDER_NO)
    if not other or j < 0 or not readable(other):
        return r
    missing = want - {x[j] for x in r.rows if j < len(x)}
    if not missing:
        return r
    extra, _cols = _other_rows(ODR, "", r.file_path, other)
    add = [x for x in (extra or []) if j < len(x) and x[j] in missing]
    if add:
        r.rows = list(r.rows) + add
        r.added = sorted({x[j] for x in add})
    return r


# ------------------------------------------------------------------
# Arrbox の絞り込み・知らせ
# ------------------------------------------------------------------
def arrbox_filter(sql_filter: str) -> str:
    """Arrbox の SQLfilter のキーワード → ADO Filter (AccHIki:42-47)。"""
    if sql_filter == "":
        return ""
    if sql_filter == "板":
        return "[ｵｰﾀﾞｰ板丈]<>0"
    if sql_filter == "コイル":
        return "[ｵｰﾀﾞｰ板丈]=0  AND  [BOX設計_設備名] = 'KEN' "
    return "[ﾛｯﾄ番号]='" + ado_filter.quote(sql_filter) + "'"


def arrbox_value(target: str, sql_filter: str = "", use: str = "Record", dbflag: str = "") -> Optional[list]:
    """知らせを出さない Arrbox(grid・judge の既定)。読めるファイルが無ければ None。"""
    name = target[:-6] if target.endswith(".accdb") else target
    if use == "Fields":
        return read_fields(name, dbflag).cols
    return read_rows(name, arrbox_filter(sql_filter), dbflag).rows


def _here_there() -> tuple[str, str]:
    return ("第1パス", "第2パス") if paths.lot_mode() == paths.ONLINE else ("第2パス", "第1パス")


def hiki_added_message(nos: Sequence[str], source: str) -> Message:
    here, there = _here_there()
    return Message(f"{here}に無い引当を{there}から足しました: 引当NO " + ", ".join(nos), INFO, source=source)
