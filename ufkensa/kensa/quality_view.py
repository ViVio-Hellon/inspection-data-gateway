"""UF品質基準 の表示内容(VBA `【UFKen_11】Quality_DB`: DBdisplay外観_H / DBdisplay外観_S / DBdisplay歪)

画面のコントロールは `QualityForm` のフィールドで受け取り、書き換えた写しを返す(元は変えない)。
VBA はモードレスのフォームに書き足していくので、前回の値が残る所(Sansyou1 への追記、範囲外の厚で
前回の公差が残る、タブが増え続ける)も、渡された form の値から続けて再現する。

照合は `quality`(QP_DB の移植)の関数を使う。行番号は `vba_src/【UFKen_11】Quality_DB.bas`(QDB)。
"""
from __future__ import annotations

import copy
from dataclasses import dataclass, field
from typing import Any, Optional, Sequence

from . import datasource, quirks, vbx
from .flow import ERROR, Message
from .quality import (NUM, STR, VAR, LotData, VbaRuntimeError, cell, check_condition, check_thickness, check_width, clng, col_of,
                      csng, find_indices, loop_array, rgb, sori_array, vcmp, veq, vsplit, vstr, vtrim, _quality_db)

VBCRLF = "\r\n"
SYSTEM_BUTTON_FACE = -2147483648           # &H80000000(Noun の既定色)
WHITE = rgb(255, 255, 255)
HAKUDO_RED = rgb(175, 15, 5)               # 白銅(DIM_122_A)
O_YELLOW = rgb(255, 255, 153)              # O 材(DIM_122_C)
HX8_BLUE = rgb(0, 102, 204)                # Hx8 材(DIM_122_D)


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
    Naka2: str = ""
    Mimi2: str = ""
    Toukyu: str = ""
    VcFlag: str = ""                       # VcFlag.Caption
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


def _set(form: QualityForm, name: str, value: Any) -> None:
    setattr(form, name, vstr(value))


def _zero_clear(form: QualityForm, name: str) -> None:
    """`If UF品質基準.X = 0 Then UF品質基準.X = ""`(テキストボックスの文字と 0 の比較)。"""
    if veq(getattr(form, name), 0, VAR, NUM):
        setattr(form, name, "")


def _dedupe(text: str) -> str:
    """Sansyou1 の空白区切りの語を Scripting.Dictionary で重複除去(最初の位置を残す)して Join。"""
    return " ".join(dict.fromkeys(vsplit(text, " ")))


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
            res.messages.append(Message("日稲:SIKALOTNOW.accdbがありません_ReadCheck_DBdisplay", ERROR, source="QDB:103"))
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
    aq = datasource.get_records(db, t2, "[用途コード] LIKE '*" + youto + "*'")
    if aq is None:
        return res
    join_quality = [list(fq)] + aq
    aq2 = datasource.get_records(db, t1)
    if aq2 is None:
        return res
    join_arr = [list(fq2)] + aq2

    matching: list[int] = []
    same = not quirks.reproduce("B71")          # B71(QDB:300-302): 直すと下・上も Double で比べる
    for r in aq:
        c1 = check_condition(r, R["材質"], zai)
        if quirks.reproduce("B25"):
            c2 = c1 & check_condition(r, R["材質"], zai)        # B25: 調質を見ずに材質をもう一度
        else:
            # 直す: 調質で照合する(表に調質の列が無ければ条件なし。いまの別表2 には列が無い)
            c2 = c1 & (check_condition(r, R["調質"], tyo) if R["調質"] else True)
        c3 = c2 & check_thickness(r, R["板厚下"], R["板厚上"], R["厚"], m_atu, same_precision=same)
        c4 = c3 & check_width(r, R["板幅下"], R["板幅上"], R["幅"], o_haba, same_precision=same)
        c5 = c4 & check_width(r, R["板丈下"], R["板丈上"], R["丈"], o_take, same_precision=same)
        c6 = c5 & check_condition(r, R["納入先"], nou)
        if c1 and c2 and c3 and c4 and c5 and c6:
            matching.append(clng(cell(r, R["管理番号"])))
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
        res.messages.append(Message("日稲:SIKALOTNOW.accdbがありません_ReadCheck_DBdisplay", ERROR, source="QDB:476"))
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
    aq = datasource.get_records(db, t1, "[用途コード] LIKE '*" + youto + "*'")
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

    arry_s = 0
    for r in aq:
        if check_condition(r, C["材質"], zai):
            if vbx.like(vstr(cell(r, C["表面処理コード"])), "*" + grade + "*"):
                if vstr(cell(r, C["厚フラグ"])) != "":
                    if vbx.val(cell(r, C["板厚1"])) >= vbx.val_single(o_atu):
                        arry_s = clng(cell(r, C["パターン1"]))
                        break
                    if vbx.val(cell(r, C["板厚1"])) < vbx.val_single(o_atu):
                        arry_s = clng(cell(r, C["パターン2"]))
                        break
                else:
                    arry_s = clng(cell(r, C["パターン1"]))
                    break
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


# ------------------------------------------------------------------
# 歪(QDB L700-1384)。UF品質基準 の手動検索だけで使う
# ------------------------------------------------------------------
VC_TARGETS = ("208PRNF", "V325NNF", "V311PNF")


def display_distortion(lot: Optional[LotData], form: QualityForm, *, odr_db: Any = "SIKAODRNOW",
                       quiet: bool = False) -> ViewResult:
    """QDB.DBdisplay歪 (QDB:700-1384)。歪公差を UF品質基準 に表示する(DIM_122 系)。

    odr_db: Arrbox と同じ場所(Main/Next)の SIKAODRNOW(VC の確認に使う)。
    流れ: オーダー厚 <= 6 なら中耳(4級。フラットネス 3 なら 3級も) → VC 指定銘柄なら QA_022_01 →
    DIM_122_B(鉄道ボディ) → DIM_122_A(白銅優先) → DIM_122_C(O 材) → DIM_122_D(Hx8 材)。
    """
    form = copy.deepcopy(form)
    res = ViewResult(form)
    msgs = res.messages
    if lot is None:
        msgs.append(Message("日稲:SIKALOTNOW.accdbがありません_ReadCheck_DBdisplay", ERROR, source="QDB:716"))
        return res
    if lot.fields is None:
        return res
    if lot.records is None:
        if not quiet:
            msgs.append(Message("SIKALOTNOWに該当ロットがありません_DBdisplay歪", ERROR, source="QDB:769"))
        return res
    la, hf = _lot_head(lot)
    o_atu = csng(_lv(la, hf, "ｵｰﾀﾞｰ板厚"))
    o_haba = csng(_lv(la, hf, "ｵｰﾀﾞｰ板幅"))
    o_take = csng(_lv(la, hf, "ｵｰﾀﾞｰ板丈"))
    m_atu = csng(_lv(la, hf, "製造板厚"))
    m_haba = csng(_lv(la, hf, "製造板幅"))
    m_take = csng(_lv(la, hf, "製造板丈"))
    grade = vstr(_lv(la, hf, "品質ｸﾞﾚｰﾄﾞ_表面処理"))
    flat = clng(_lv(la, hf, "品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ"))
    nou = vstr(_lv(la, hf, "納入先名"))
    youto = vstr(_lv(la, hf, "用途ｺｰﾄﾞ"))
    vstr(_lv(la, hf, "用途名"))
    order_no = vstr(_lv(la, hf, "ｵｰﾀﾞｰ番号"))
    zai = vstr(_lv(la, hf, "製造材質"))
    tyo = vstr(_lv(la, hf, "製造調質"))
    form.Uatu = vbx.fmt(m_atu, "0.000")
    form.Uhaba = vbx.fmt(m_haba, "0.0")
    form.Utake = vbx.fmt(m_take, "0.0")
    form.Uzai, form.Utyou, form.UYoucode, form.Uhyou = zai, tyo, youto, grade
    form.Noun = nou
    form.Noun_BackColor = SYSTEM_BUTTON_FACE
    atu_flag = o_atu <= 6.0
    vc_flag = vbx.val(_lv(la, hf, "VC")) == 1

    use_flag = False
    if vc_flag:
        vf = datasource.get_fields(odr_db, "仕掛")
        if vf is not None:
            c_a, c_b = col_of(vf, "VC_表"), col_of(vf, "VC_裏")
            vcr = datasource.get_records(odr_db, "仕掛", "[受注番号]='" + order_no + "'")
            if vcr is not None:
                vca, vcb = vstr(cell(vcr[0], c_a)), vstr(cell(vcr[0], c_b))
                use_flag = vca in VC_TARGETS or vcb in VC_TARGETS

    db = _quality_db()

    def dedupe() -> None:
        form.Sansyou1 = _dedupe(form.Sansyou1)

    # A. 中耳歪(オーダー厚 <= 6)
    if atu_flag:
        form.Toukyu = "4級"
        done = False
        f = datasource.get_fields(db, "耳歪4級")
        rows = datasource.get_records(db, "耳歪4級") if f is not None else None
        if f is not None and rows is not None:
            form.Mimi1 = vbx.fmt(loop_array(rows, f, o_take, o_atu), "0.0")
            _zero_clear(form, "Mimi1")
            f = datasource.get_fields(db, "中歪4級")
            rows = datasource.get_records(db, "中歪4級") if f is not None else None
            if f is not None and rows is not None:
                form.Naka1 = vbx.fmt(loop_array(rows, f, o_take, o_atu), "0.0")
                _zero_clear(form, "Naka1")
                form.Sansyou1 = form.Sansyou1 + " ﾌﾗｯﾄﾈｽ製品規格中耳歪(ｵｰﾀﾞｰ丈判定)"
                done = True
                if flat == 3:
                    # 3級も表を引く(列は中歪4級のものを使う。VBA どおり)
                    rows = datasource.get_records(db, "耳歪3級")
                    if rows is None:
                        done = False
                    else:
                        form.Mimi2 = vbx.fmt(loop_array(rows, f, o_take, o_atu), "0.0")
                        _zero_clear(form, "Mimi2")
                        rows = datasource.get_records(db, "中歪3級")
                        if rows is None:
                            done = False
                        else:
                            form.Naka2 = vbx.fmt(loop_array(rows, f, o_take, o_atu), "0.0")
                            _zero_clear(form, "Naka2")
                            form.Toukyu = "3級"
                            form.Sansyou1 = form.Sansyou1 + " ﾌﾗｯﾄﾈｽ製品規格中耳歪(ｵｰﾀﾞｰ丈判定)"
        if done:
            dedupe()

    # B. VC 指定(QA_022_01)
    goto = ""
    if use_flag:
        f = datasource.get_fields(db, "QA_022_01")
        if f is None:
            return res
        c_lo, c_hi, c_mg = col_of(f, "板厚下"), col_of(f, "板厚上"), col_of(f, "1m当たり平坦度上限")
        rows = datasource.get_records(db, "QA_022_01")
        if rows is None:
            goto = "skipO"
        else:
            for r in rows:
                lo = vbx.single(vbx.val(cell(r, c_lo)))
                hi = vbx.single(vbx.val(cell(r, c_hi)))
                if lo <= m_atu and hi >= m_atu:
                    form.MG1 = vbx.fmt(cell(r, c_mg), "0.0")
                    break
            form.HaHiz1 = vbx.fmt(loop_array(rows, f, o_haba, o_atu), "0.0")
            _zero_clear(form, "HaHiz1")
            form.TaHiz1 = vbx.fmt(loop_array(rows, f, o_take, o_atu), "0.0")
            _zero_clear(form, "TaHiz1")
            form.VcFlag = "あり"
            form.Sansyou1 = form.Sansyou1 + " QA-022-01"
            dedupe()
            return res

    # C. DIM_122_B(鉄道ボディ)
    if goto == "":
        f = datasource.get_fields(db, "Q_S_QA_DIM_122_B")
        if f is None:
            goto = "skipO"
        else:
            C = {n: col_of(f, n) for n in ("用途コード", "板厚下", "板厚上", "板幅下", "板幅上", "板幅", "板丈")}
            rows = datasource.get_records(db, "Q_S_QA_DIM_122_B", "[用途コード] LIKE '*" + youto + "*'")
            if rows is None:
                goto = "skip122A"
            else:
                found = any(vtrim(x) == youto for x in vsplit(vstr(cell(rows[0], C["用途コード"])), ","))
                if found:
                    if len(zai) >= 4:
                        zai = vbx.right(zai, 3)
                    rows = datasource.get_records(db, "Q_S_QA_DIM_122_B",
                                                  "[材質] LIKE '*" + zai + "*' AND  [調質] LIKE  '*" + tyo + "*'")
                    if rows is None:
                        return res
                    for r in rows:
                        lo = vbx.single(vbx.val(cell(r, C["板厚下"])))
                        hi = vbx.single(vbx.val(cell(r, C["板厚上"])))
                        wlo = vbx.single(vbx.val(cell(r, C["板幅下"])))
                        whi = vbx.single(vbx.val(cell(r, C["板幅上"])))
                        if lo <= o_atu and hi >= o_atu and wlo <= o_haba and whi >= o_haba:
                            form.HaHiz1 = vbx.fmt(cell(r, C["板幅"]), "0.0")
                            _zero_clear(form, "HaHiz1")
                            form.TaHiz1 = vbx.fmt(cell(r, C["板丈"]), "0.0")
                            _zero_clear(form, "TaHiz1")
                            form.Sansyou1 = form.Sansyou1 + " Q-S-QA-DIM-122_B"
                            form.Mimi1 = form.Naka1 = form.Mimi2 = form.Naka2 = "不問"
                            break
                    dedupe()
                    return res
                goto = "skip122A"

    # D. DIM_122_A(白銅を優先)
    if goto == "skip122A":
        goto = "skipO"
        f = datasource.get_fields(db, "Q_S_QA_DIM_122_A")
        if f is not None:
            C = {n: col_of(f, n) for n in ("板厚下", "板厚上", "1m当たり平坦度上限", "納入フラグ", "材質")}
            rows = datasource.get_records(db, "Q_S_QA_DIM_122_A",
                                          "[用途コード] LIKE '*" + youto + "*' AND  [調質] LIKE  '*" + tyo + "*'")
            if rows is not None:
                width = len(rows[0])
                arr_h: list[list[Any]] = [[None] * width for _ in rows]
                n = 0
                if len(zai) >= 4:
                    zai = vbx.right(zai, 3)
                haku = False
                for r in rows:
                    lo = vbx.single(vbx.val(cell(r, C["板厚下"])))
                    hi = vbx.single(vbx.val(cell(r, C["板厚上"])))
                    if vbx.like(vstr(cell(r, C["材質"])), "*" + zai + "*"):
                        if lo <= m_atu and hi >= m_atu:
                            if veq(cell(r, C["納入フラグ"]), "白銅", VAR, STR) and vbx.like(nou, "*ﾊｸﾄﾞｳ*"):
                                arr_h[n] = list(r)
                                n += 1
                                form.MG1 = vbx.fmt(cell(r, C["1m当たり平坦度上限"]), "0.0")
                                form.Noun_BackColor = HAKUDO_RED
                                haku = True
                                break
                if not haku:
                    for r in rows:
                        lo = vbx.single(vbx.val(cell(r, C["板厚下"])))
                        hi = vbx.single(vbx.val(cell(r, C["板厚上"])))
                        if vbx.like(vstr(cell(r, C["材質"])), "*" + zai + "*"):
                            if lo <= m_atu and hi >= m_atu:
                                if not veq(cell(r, C["納入フラグ"]), "白銅", VAR, STR):
                                    arr_h[n] = list(r)
                                    n += 1
                                    form.MG1 = vbx.fmt(cell(r, C["1m当たり平坦度上限"]), "0.0")
                                    form.Noun_BackColor = WHITE
                form.HaHiz1 = vbx.fmt(loop_array(arr_h, f, m_haba, m_atu), "0.0")
                form.TaHiz1 = vbx.fmt(loop_array(arr_h, f, m_take, m_atu), "0.0")
                _zero_clear(form, "HaHiz1")
                _zero_clear(form, "TaHiz1")
                if not (form.HaHiz1 == "" and form.TaHiz1 == ""):
                    form.Sansyou1 = form.Sansyou1 + " Q-S-QA-DIM-122"
                    dedupe()
                    return res

    # E. DIM_122_C(O 材)
    f = datasource.get_fields(db, "Q_S_QA_DIM_122_C")
    if f is None:
        return res
    C = {n: col_of(f, n) for n in ("製造材質", "用途コード", "板厚下", "板厚上", "オーダー板幅", "幅公差", "丈公差",
                                   "納入先名", "パターン")}
    rows = datasource.get_records(db, "Q_S_QA_DIM_122_C", "[調質] LIKE '*" + tyo + "*'")
    if rows is not None:
        if len(zai) >= 4:
            zai = vbx.right(zai, 3)
        matched = False
        for r in rows:
            if not vbx.like(vstr(cell(r, C["製造材質"])), "*" + zai + "*"):
                continue
            if not vbx.like(vstr(cell(r, C["用途コード"])), "*" + youto + "*"):
                continue
            if not vbx.like(nou, "*" + vstr(cell(r, C["納入先名"])) + "*"):
                continue
            # B71(QDB:1245-1247): 板厚下・上は Single、Val(オーダー厚) は Double。直すと下・上も Double
            if quirks.reproduce("B71"):
                lo = vbx.single(vbx.val(cell(r, C["板厚下"])))
                hi = vbx.single(vbx.val(cell(r, C["板厚上"])))
            else:
                lo = vbx.val(cell(r, C["板厚下"]))
                hi = vbx.val(cell(r, C["板厚上"]))
            if not (lo <= vbx.val_single(o_atu) and hi >= vbx.val_single(o_atu)):
                continue
            # B71(QDB:1251・1257): DB の幅(文字列 → Double)と Single のオーダー幅を比べる。
            # 直すとオーダー幅も Val(Single)(有効 7 桁の Double)にして同じ精度で比べる
            w = o_haba if quirks.reproduce("B71") else vbx.val_single(o_haba)
            pat = cell(r, C["パターン"])
            hit = False
            if vstr(pat) != "":
                if veq(pat, 1, VAR, NUM):
                    hit = vcmp(cell(r, C["オーダー板幅"]), w, VAR, NUM) >= 0
                elif veq(pat, 2, VAR, NUM):
                    hit = vcmp(cell(r, C["オーダー板幅"]), w, VAR, NUM) <= 0
            else:
                hit = True
            if hit:
                form.HaHiz1 = vbx.fmt(cell(r, C["幅公差"]), "0.0")
                form.TaHiz1 = vbx.fmt(cell(r, C["丈公差"]), "0.0")
                form.Noun_BackColor = O_YELLOW
                matched = True                       # 最後に一致した行が残る
        _zero_clear(form, "HaHiz1")
        _zero_clear(form, "TaHiz1")
        if matched or quirks.reproduce("B27"):
            # B27: 一致が無くても「不問」と参照元を付ける
            form.Mimi1 = form.Naka1 = form.Mimi2 = form.Naka2 = "不問"
            form.Sansyou1 = form.Sansyou1 + " Q-S-QA-DIM-122_O"
        dedupe()
        return res

    # F. DIM_122_D(Hx8 材)
    f = datasource.get_fields(db, "Q_S_QA_DIM_122_D")
    if f is None:
        return res
    C = {n: col_of(f, n) for n in ("製造材質", "製造調質", "用途コード", "オーダー板厚", "幅公差", "丈公差", "納入先名",
                                   "フラットネス品質等級")}
    if len(zai) >= 4:
        zai = vbx.right(zai, 3)
    rows = datasource.get_records(db, "Q_S_QA_DIM_122_D", "[製造材質] LIKE '*" + zai + "*'")
    if rows is None:
        return res
    tyo_a, tyo_k = vbx.left(tyo, 1), vbx.right(tyo, 1)
    # VBA の For は上限を最初に決める。列番号(C)も最初の表のまま。配列だけ差し替わることがある(B27)
    arr, fields = rows, f
    itti = False
    hit_any = False                 # 条件(調質・用途・納入先・厚)に当たった行があった(B73)
    for i in range(len(rows)):
        r = arr[i] if i < len(arr) else None
        if r is None:
            raise VbaRuntimeError(9, "差し替えたソリ表の行が足りません")
        if not vbx.like(vstr(cell(r, C["製造調質"])), tyo_a + "*" + tyo_k):
            continue
        if not vbx.like(vstr(cell(r, C["用途コード"])), "*" + youto + "*"):
            continue
        if not vbx.like(nou, "*" + vstr(cell(r, C["納入先名"])) + "*"):
            continue
        if vbx.val(cell(r, C["オーダー板厚"])) != vbx.val_single(o_atu):
            continue
        hit_any = True
        if vstr(cell(r, C["フラットネス品質等級"])) == "":
            form.HaHiz1 = vbx.fmt(cell(r, C["幅公差"]), "0.0")
            _zero_clear(form, "HaHiz1")
            form.TaHiz1 = vbx.fmt(cell(r, C["丈公差"]), "0.0")
            _zero_clear(form, "TaHiz1")
            form.Noun_BackColor = HX8_BLUE
            form.Sansyou1 = form.Sansyou1 + " Q-S-QA-DIM-122_D"
            itti = True
        else:
            sori = "ソリ4級" if flat == 4 else ("ソリ3級" if flat == 3 else "")
            if sori:
                sf = datasource.get_fields(db, sori)
                if sf is None:
                    return res
                sr = datasource.get_records(db, sori)
                if sr is None:
                    return res
                form.Sansyou1 = form.Sansyou1 + " 製品規格_" + sori
            if quirks.reproduce("B27"):
                # B27: ループ中の配列をソリ表に差し替え、列名が H/T 形式のソリ表を LoopArray で引く
                # (数の列名が無いので LoopArray が添字エラーになる)
                if sori:
                    arr, fields = sr, sf
                form.HaHiz1 = vbx.fmt(loop_array(arr, fields, o_haba, o_atu), "0.0")
                form.TaHiz1 = vbx.fmt(loop_array(arr, fields, o_take, o_atu), "0.0")
                form.Noun_BackColor = HX8_BLUE
            else:
                # 直す: ソリ表は DBQP と同じく FindIndices + SoriArray で引き、ここで止める
                if sori:
                    form.HaHiz1 = vbx.fmt(sori_array(sr, sf, o_atu, find_indices(sf, "H", o_haba)[0]), "0.0")
                    form.TaHiz1 = vbx.fmt(sori_array(sr, sf, o_atu, find_indices(sf, "T", o_take)[0]), "0.0")
                form.Noun_BackColor = HX8_BLUE
                itti = True
        if itti:
            break
    _zero_clear(form, "HaHiz1")
    _zero_clear(form, "TaHiz1")
    if hit_any or quirks.reproduce("B73"):
        # B73: VBA は一致が無くても中耳を「不問」で上書きする。直した場合は一致した行があったときだけ
        form.Mimi1 = form.Naka1 = form.Mimi2 = form.Naka2 = "不問"
    dedupe()
    return res
