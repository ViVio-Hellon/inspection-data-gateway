"""UF品質基準 の表示内容(VBA `【UFKen_11】Quality_DB`: DBdisplay外観_H / DBdisplay外観_S)

DBdisplay歪(DIM_122 / QA_022_01)は VBA で止め忘れの経路なので移植しない(品質基準画面の歪は平坦度一覧QP追加)。

画面のコントロールは `QualityForm` のフィールドで受け取り、書き換えた写しを返す(元は変えない)。
VBA はモードレスのフォームに書き足していくので、前回の値が残る所(Sansyou1 への追記、範囲外の厚で
前回の公差が残る、タブが増え続ける)も、渡された form の値から続けて再現する。

照合は `quality`(QP_DB の移植)の関数を使う。行番号は `vba_src/【UFKen_11】Quality_DB.bas`(QDB)。
"""
from __future__ import annotations

import copy
from dataclasses import dataclass, field
from typing import Any, Optional, Sequence

from . import datasource, match_rules, paths, vbx
from .flow import ERROR, Message
from .quality import (NUM, STR, VAR, LotData, VbaRuntimeError, cell, check_condition, check_thickness, check_width, clng, col_of,
                      csng, rgb, sika_ctx, veq, vstr, _quality_db)

VBCRLF = "\r\n"
SYSTEM_BUTTON_FACE = -2147483648           # &H80000000(Noun の既定色)
WHITE = rgb(255, 255, 255)


@dataclass
class QualityForm:
    """UF品質基準 のコントロール(名前は VBA のまま。ラベルの Caption も値として持つ)。"""
    ATU: str = ""
    HABA: str = ""
    TAKE: str = ""
    Zai: str = ""
    TYO: str = ""
    YoutC: str = ""
    Noun: str = ""
    Noun_BackColor: int = SYSTEM_BUTTON_FACE
    Uatu: str = ""
    Uhaba: str = ""
    Utake: str = ""
    Uzai: str = ""
    Utyou: str = ""
    UYoucode: str = ""
    Uhyou: str = ""
    Upater: str = ""
    SelectO: str = ""
    SelectS: str = ""
    SelectK: str = ""
    SelectA: str = ""
    MG1: str = ""
    HaHiz1: str = ""
    TaHiz1: str = ""
    Naka1: str = ""
    Mimi1: str = ""
    Sansyou1: str = ""
    KousinL1: str = ""
    KousinL2: str = ""
    KousinL3: str = ""
    KousinL4: str = ""
    AbnoName: str = ""
    AbnoLabel: str = ""
    gradeTB: str = ""
    AbnoTB: str = ""
    tabs: list[str] = field(default_factory=list)        # TabAb のタブ(追加するだけで消さない)
    L2_visible: bool = False                              # 「読込中」表示

    def apply(self, values: dict[str, Any]) -> None:
        """quality.dbqp の form・quality.form_reflect_dist の values("KousinL2.Caption" など)を書く。"""
        for k, v in values.items():
            name = k.replace(".Caption", "")
            if not hasattr(self, name):
                raise KeyError(f"UF品質基準 に {k} はありません")
            setattr(self, name, vstr(v))


@dataclass
class ViewResult:
    form: QualityForm
    messages: list[Message] = field(default_factory=list)
    grade_array: Optional[list[list[Any]]] = None        # DBdisplay外観_H の戻り値(GradeArray)


def _lot_head(lot: LotData) -> tuple[Sequence[Any], list[str]]:
    return lot.records[0], lot.fields


def _lv(row: Sequence[Any], fields: list[str], name: str) -> Any:
    return cell(row, col_of(fields, name))


# ------------------------------------------------------------------
# 厚板の外観(QDB L5-460)
# ------------------------------------------------------------------
_H_COND_COLS = ("管理番号", "材質", "用途コード", "板厚下", "板厚上", "厚", "更新日", "備考", "調質", "納入先",
                "板幅下", "板幅上", "幅", "板丈下", "板丈上", "丈")


def false_search_l(join_arr: list[list[Any]], resul: list[list[Any]]) -> list[list[Any]]:
    """QDB.FalseSearch_L (QDB:52-83)。別表1(見出し行つき)から (欠陥名, 許容内容, グレード) を作る。

    resul: (欠陥名, グレード) の行と最後の ("管理番号", 番号)。結果は欠陥名の数の行(埋まらない行は None)。
    同じ欠陥名・グレードが複数あると VBA と同じく添字エラー。
    """
    size = len(resul) - 1
    if size < 1:
        raise VbaRuntimeError(9, "ReDim TempSize(1 To 0)")
    out: list[list[Any]] = [[None, None, None] for _ in range(size)]
    c_name = col_of(join_arr[0], "欠陥名")
    amatch = 0
    for r in resul:
        for row in join_arr:
            if veq(cell(row, c_name), r[0]):
                for ic in range(1, len(join_arr[0]) + 1):
                    if veq(r[1], cell(join_arr[0], ic)):
                        amatch += 1
                        if amatch > size:
                            raise VbaRuntimeError(9, "FalseSearch_L の行が足りません")
                        out[amatch - 1] = [cell(row, c_name), cell(row, ic), cell(join_arr[0], ic)]
    return out


def _youto_filter(youto: str) -> str:
    return "[用途コード] LIKE '*" + youto + "*'"


def _pre_rows(pre: "match_rules.Group", db: Any, table: str, fq: Sequence[str], ctx: dict
              ) -> Optional[list[list[Any]]]:
    """新: 先に外すルール(用途コード 含む(SQL))で絞った行。0 件は None(GetRecordsArr と同じ)。"""
    rows = datasource.get_records(db, table) or []
    return [rows[i] for i in match_rules.match(pre, rows, fq, ctx).matched] or None


def _h_old(aq: Sequence[Sequence[Any]], R: dict[str, int], zai: str, tyo: str, m_atu: float, o_haba: float,
           o_take: float, nou: str) -> list[int]:
    """旧: DBdisplay外観_H の照合(QDB:296-305)をコードで。照合ルールの shadow で差が 0 件と確かめたら消す。"""
    matching: list[int] = []
    same = True          # B71(QDB:300-302): 直すと下・上も Double で比べる
    for r in aq:
        c1 = check_condition(r, R["材質"], zai)
        # 直す: 調質で照合する(表に調質の列が無ければ条件なし。いまの別表2 には列が無い)
        c2 = c1 & (check_condition(r, R["調質"], tyo) if R["調質"] else True)
        c3 = c2 & check_thickness(r, R["板厚下"], R["板厚上"], R["厚"], m_atu, same_precision=same)
        c4 = c3 & check_width(r, R["板幅下"], R["板幅上"], R["幅"], o_haba, same_precision=same)
        c5 = c4 & check_width(r, R["板丈下"], R["板丈上"], R["丈"], o_take, same_precision=same)
        c6 = c5 & check_condition(r, R["納入先"], nou)
        if c1 and c2 and c3 and c4 and c5 and c6:
            matching.append(clng(cell(r, R["管理番号"])))
    return matching


def _s_old(aq: Sequence[Sequence[Any]], C: dict[str, int], zai: str, grade: str) -> Optional[int]:
    """旧: DBdisplay外観_S の照合(QDB:598-599)をコードで。最初に合う行の位置(無ければ None)。"""
    for i, r in enumerate(aq):
        if check_condition(r, C["材質"], zai):
            if vbx.like(vstr(cell(r, C["表面処理コード"])), "*" + grade + "*"):
                return i
    return None


def add_pages_arr_n(form: QualityForm, n: Optional[list[list[Any]]]) -> str:
    """H.AddPagesArrN (H:1569-1621)。各行の欠陥名でタブを**追加**し、欠陥名 & 改行 を連ねて返す。"""
    if not n:
        return ""
    rep = ""
    for row in n:
        form.tabs.append(vstr(row[0]))
        rep += vstr(row[0]) + VBCRLF
    return rep


def uf_db_display(form: QualityForm, temp_arr: Optional[list[list[Any]]], tab_caption: str) -> None:
    """QDB.UF_DbDisplay (QDB:411-460)。タブ名と一致する行(最後の一致)を AbnoLabel/gradeTB/AbnoTB に。"""
    if not temp_arr:
        return
    for row in temp_arr:
        if veq(tab_caption, row[0], STR, VAR):
            form.AbnoLabel = vstr(row[0])
            form.gradeTB = vstr(row[-1])
            form.AbnoTB = vstr(row[1])


def display_appearance_h(lot: Optional[LotData], form: QualityForm, *, quiet: bool = False) -> ViewResult:
    """QDB.DBdisplay外観_H (QDB:86-409)。厚板(オーダー厚 >= 8)の外観グレード。

    lot: Arrbox の結果(None は ReadCheck で SIKALOTNOW が無い)。quiet: IsOver12Hours が真(MsgBox を出さない)。
    - Q_S_QA_124_別表2 を `[用途コード] LIKE '*用途*'` で絞り、材質・(調質 B25)・製造厚・オーダー幅/丈・納入先で照合
    - 一致した管理番号の欠陥名グレードを**後の番号で上書き**(最後の一致が残る)
    - 別表1 から許容内容を引き、タブを追加して先頭の欠陥を表示。戻り値(grade_array)は GradeArray
    """
    form = copy.deepcopy(form)
    res = ViewResult(form)
    if lot is None:
        if not quiet:
            res.messages.append(Message(paths.missing_text("SIKALOTNOW"), ERROR, source="QDB:103"))
        return res
    if lot.fields is None:
        return res
    if lot.records is None:
        if not quiet:
            res.messages.append(Message("SIKALOTNOW.accdbに該当ロットがありません_DBdisplay外観", ERROR,
                                        source="QDB:148"))
        return res
    la, hf = _lot_head(lot)
    o_atu = csng(_lv(la, hf, "ｵｰﾀﾞｰ板厚"))
    o_haba = csng(_lv(la, hf, "ｵｰﾀﾞｰ板幅"))
    o_take = csng(_lv(la, hf, "ｵｰﾀﾞｰ板丈"))
    m_atu = csng(_lv(la, hf, "製造板厚"))
    m_haba = csng(_lv(la, hf, "製造板幅"))
    m_take = csng(_lv(la, hf, "製造板丈"))
    vstr(_lv(la, hf, "品質ｸﾞﾚｰﾄﾞ_表面処理"))
    nou = vstr(_lv(la, hf, "納入先名"))
    youto = vstr(_lv(la, hf, "用途ｺｰﾄﾞ"))
    vstr(_lv(la, hf, "用途名"))
    vstr(_lv(la, hf, "ｵｰﾀﾞｰ番号"))
    zai = vstr(_lv(la, hf, "製造材質"))
    tyo = vstr(_lv(la, hf, "製造調質"))
    form.ATU = vbx.fmt(m_atu, "0.000")
    form.HABA = vbx.fmt(m_haba, "0.0")
    form.TAKE = vbx.fmt(m_take, "0.0")
    form.Zai, form.TYO, form.YoutC, form.Noun = zai, tyo, youto, nou
    if not o_atu >= 8.0:
        return res

    db = _quality_db()
    t2, t1 = "Q_S_QA_124_別表2", "Q_S_QA_124_別表1"
    fq = datasource.get_fields(db, t2)
    if fq is None:
        return res
    fq2 = datasource.get_fields(db, t1)
    if fq2 is None:
        return res
    R = {n: col_of(fq, n) for n in ("材質", "用途コード", "板厚下", "板厚上", "厚", "更新日", "管理番号", "調質", "納入先",
                                    "板幅下", "板幅上", "幅", "板丈下", "板丈上", "丈")}
    arry_l: list[list[Any]] = [[f, None] for f in fq if f not in _H_COND_COLS]
    arry_l.append([None, None])
    g = match_rules.group(match_rules.APPEAR_H)
    shadow = g.new_way and match_rules.shadow_on(g)
    ctx = sika_ctx(lot)
    if g.new_way:
        pre, rest = match_rules.split_pre(g)
        aq = _pre_rows(pre, db, t2, fq, ctx)
        if shadow:
            match_rules.compare(match_rules.APPEAR_H, datasource.get_records(db, t2, _youto_filter(youto)), aq,
                                "用途 " + youto)
    else:
        aq = datasource.get_records(db, t2, _youto_filter(youto))
    if aq is None:
        return res
    join_quality = [list(fq)] + aq
    aq2 = datasource.get_records(db, t1)
    if aq2 is None:
        return res
    join_arr = [list(fq2)] + aq2

    def old() -> list[int]:
        return _h_old(aq, R, zai, tyo, m_atu, o_haba, o_take, nou)
    if g.new_way:
        matching = [clng(cell(aq[i], R["管理番号"])) for i in match_rules.match(rest, aq, fq, ctx).rows]
        if shadow:
            match_rules.compare(match_rules.APPEAR_H, old(), matching, repr((zai, tyo, m_atu, o_haba, o_take, nou)))
    else:
        matching = old()
    if not matching:
        return res

    for no in matching:
        for row in join_quality:
            if veq(cell(row, R["管理番号"]), no, VAR, NUM):    # 見出し行 "管理番号" と Long も比べる(U1)
                arry_l[-1][0] = "管理番号"
                arry_l[-1][1] = no
                for item in arry_l:
                    for jc in range(1, len(join_quality[0]) + 1):
                        if veq(item[0], cell(join_quality[0], jc)):
                            item[1] = cell(row, jc)

    temp_size_l = false_search_l(join_arr, arry_l)
    temp = add_pages_arr_n(form, temp_size_l)
    form.AbnoName = "「異常グレード確認一覧」" + VBCRLF + temp
    all_rows = datasource.get_records(db, t2)
    if all_rows is not None:
        form.KousinL1 = vstr(cell(all_rows[0], R["更新日"]))
    uf_db_display(form, temp_size_l, vstr(temp_size_l[0][0]))
    res.grade_array = temp_size_l
    return res


# ------------------------------------------------------------------
# 薄板の表面(QDB L5-50, L462-698)
# ------------------------------------------------------------------
def false_search_s(join_arr: list[list[Any]], pattern: int) -> list[list[Any]]:
    """QDB.FalseSearch_S (QDB:5-50)。別表2(見出し行つき)からパターンの行(O/S/K/A)× 異常名の表(5行)。

    見出し行も対象になる(Val("パターン") = 0 なのでパターン 0 のとき見出し行が一致する)。
    同じパターンの行が5行以上あると添字エラー。
    """
    hdr = join_arr[0]
    names = [h for h in hdr if h not in ("管理番号", "パターン", "備考", "更新日")]
    size: list[list[Any]] = [list(names)] + [[None] * len(names) for _ in range(4)]
    c_pat = col_of(hdr, "パターン")
    n = 0
    for row in join_arr:
        if vbx.val(cell(row, c_pat)) == vbx.val(pattern):
            n += 1
            if n + 1 > 5:
                raise VbaRuntimeError(9, "FalseSearch_S の行が5行を超えた")
            for a, name in enumerate(names):
                for ic in range(1, len(hdr) + 1):
                    if veq(cell(hdr, ic), name):
                        size[n][a] = cell(row, ic)
    return size


def kijun_hanei(form: QualityForm, temp: str, arr: list[list[Any]]) -> None:
    """QDB.基準反映 (QDB:676-698)。押したボタンの異常名の列で、1列目が O/S/K/A の行を表示する。"""
    form.SelectO = form.SelectS = form.SelectK = form.SelectA = ""
    for i, name in enumerate(arr[0]):
        if veq(name, temp, VAR, STR):
            for a in range(1, len(arr)):
                for key in ("O", "S", "K", "A"):
                    if veq(arr[a][0], key, VAR, STR):
                        setattr(form, "Select" + key, vstr(arr[a][i]))
            break


def display_appearance_s(lot: Optional[LotData], out_name: str, form: QualityForm, *,
                         quiet: bool = False) -> ViewResult:
    """QDB.DBdisplay外観_S (QDB:462-675)。薄板(オーダー厚 < 8)の表面パターンと、押した異常の基準。

    out_name: 押したボタン(B1〜B11)の Caption。
    - SRF_128_別表1 を `[用途コード] LIKE '*用途*'` で絞り、材質(CheckCondition)と表面処理コード Like *グレード* の
      **最初の行**でパターン(厚フラグがあれば 板厚1 >= オーダー厚 でパターン1、超えればパターン2)
    - 別表2 でパターンの行を集め、ボタンの異常名の O/S/K/A を表示
    """
    form = copy.deepcopy(form)
    res = ViewResult(form)
    if lot is None:
        res.messages.append(Message(paths.missing_text("SIKALOTNOW"), ERROR, source="QDB:476"))
        return res
    form.L2_visible = True
    if lot.fields is None:
        form.L2_visible = False
        return res
    if lot.records is None:
        if not quiet:
            res.messages.append(Message("SIKALOTNOWに該当ロットがありません_DBdisplay外観", ERROR, source="QDB:525"))
        form.L2_visible = False
        return res
    la, hf = _lot_head(lot)
    o_atu = csng(_lv(la, hf, "ｵｰﾀﾞｰ板厚"))
    m_atu = csng(_lv(la, hf, "製造板厚"))
    m_haba = csng(_lv(la, hf, "製造板幅"))
    m_take = csng(_lv(la, hf, "製造板丈"))
    grade = vstr(_lv(la, hf, "品質ｸﾞﾚｰﾄﾞ_表面処理"))
    vstr(_lv(la, hf, "納入先名"))
    youto = vstr(_lv(la, hf, "用途ｺｰﾄﾞ"))
    vstr(_lv(la, hf, "用途名"))
    vstr(_lv(la, hf, "ｵｰﾀﾞｰ番号"))
    zai = vstr(_lv(la, hf, "製造材質"))
    tyo = vstr(_lv(la, hf, "製造調質"))
    form.Uatu = vbx.fmt(m_atu, "0.000")
    form.Uhaba = vbx.fmt(m_haba, "0.0")
    form.Utake = vbx.fmt(m_take, "0.0")
    form.Uzai, form.Utyou, form.UYoucode, form.Uhyou = zai, tyo, youto, grade
    if not o_atu < 8.0:
        form.Upater = "該当板厚なし"
        form.L2_visible = False
        return res

    db = _quality_db()
    t1, t2 = "Q_S_QA_SRF_128_別表1", "Q_S_QA_SRF_128_別表2"
    fq = datasource.get_fields(db, t1)
    if fq is None:
        return res
    C = {n: col_of(fq, n) for n in ("表面処理コード", "材質", "厚フラグ", "板厚1", "パターン1", "パターン2", "更新日")}
    g = match_rules.group(match_rules.APPEAR_S)
    shadow = g.new_way and match_rules.shadow_on(g)
    ctx = sika_ctx(lot)
    if g.new_way:
        pre, rest = match_rules.split_pre(g)
        aq = _pre_rows(pre, db, t1, fq, ctx)
        if shadow:
            match_rules.compare(match_rules.APPEAR_S, datasource.get_records(db, t1, _youto_filter(youto)), aq,
                                "用途 " + youto)
    else:
        aq = datasource.get_records(db, t1, _youto_filter(youto))
    if aq is None:
        form.L2_visible = False
        form.Upater = "該当用途ｺｰﾄﾞなし"
        return res
    fq2 = datasource.get_fields(db, t2)
    if fq2 is None:
        return res                       # VBA は Exit Sub の後に L2 を消す行があり、消えない
    aq2 = datasource.get_records(db, t2)
    if aq2 is None:
        return res
    join_arr = [list(fq2)] + aq2

    if g.new_way:
        used = match_rules.match(rest, aq, fq, ctx).rows
        hit = used[0] if used else None
        if shadow:
            match_rules.compare(match_rules.APPEAR_S, _s_old(aq, C, zai, grade), hit, repr((zai, grade)))
    else:
        hit = _s_old(aq, C, zai, grade)
    arry_s = 0
    if hit is not None:
        r = aq[hit]
        if vstr(cell(r, C["厚フラグ"])) != "":
            if vbx.val(cell(r, C["板厚1"])) >= vbx.val_single(o_atu):
                arry_s = clng(cell(r, C["パターン1"]))
            elif vbx.val(cell(r, C["板厚1"])) < vbx.val_single(o_atu):
                arry_s = clng(cell(r, C["パターン2"]))
        else:
            arry_s = clng(cell(r, C["パターン1"]))
    form.Upater = str(arry_s)
    temp_size_s = false_search_s(join_arr, arry_s)
    if arry_s == 0:
        form.Upater = "該当ﾊﾟﾀｰﾝなし"
        form.L2_visible = False
        return res
    kijun_hanei(form, out_name, temp_size_s)
    form.KousinL4 = vstr(cell(aq[0], C["更新日"]))
    form.L2_visible = False
    if not vbx.like(form.Sansyou1, "* Q-S-QA-SRF-128*"):
        form.Sansyou1 = form.Sansyou1 + " Q-S-QA-SRF-128"
    return res
