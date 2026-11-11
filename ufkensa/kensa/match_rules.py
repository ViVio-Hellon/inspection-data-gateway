"""品質基準の照合ルールをデータで持つ(2026-11-11 設計 quality_rules_design.md の第1段)。

伝送用ファイル.sqlite3 の表
- 「照合ルール」 列: ルール群・順番・表の列・ロットの項目・比べ方・前処理・空のとき・有効・備考
- 「照合ルール群」列: ルール群・使う行(最初 / 最後 / 全部)・新方式(1 = ルールで照合、0 = いままでのコード)・備考

第1段で移したのは ラテフリ直角(QP.DbReturnItem)と 梱包保護材(QP.MaterialsUpperPart)の行の絞り込みだけ。
どの表を使うか・結果の後処理(要の集め方・該当なしの文言・管理番号での引き直し)はコードに残す。

**既定のルール = いままでのコードを 1 行ずつ写したもの。** 表が無いときは既定で照合する(検査中は書かない)。
表は保守 → 照合ルール を開いたとき・保存したときに作る。読めないときは最後に読めたルール(無ければ既定)。

比べ方(新しい比較は書かない。quality.py の部品をそのまま呼ぶ):
    条件式      check_condition(表のセルを半角化し、カンマ区切りの X / X* / *X / *X* / ≠… で見る)
    範囲        check_thickness(表の列 = 「下,上,値」。3 列とも空なら通す。空のとき=通さない なら check_within)
    以上        check_size(ロットの値 ≧ CLng(表)。空なら通す)
    一致        表のセル(文字)とロットの値(文字)が同じ
    当板用か    表のセルが「当板用」かどうかが、ロットの値(「当板用」/ それ以外)と同じ。合わない行は先に外す
    ライン      表のセル(ライン名)とラインが lines.matches で同じ(文字の幅だけ揃える。D13)。合わない行は先に外す
    含む条件式  表のセル(カンマ区切り)のどれかに当たる: ≠X = X でない / *X・X* = X を**含む** / それ以外は完全一致
                (MachineTransfer.CheckThickAndCode の用途。条件式 の前方一致とは別。空のセルは当たらない)
    含む(SQL)   `[表の列] LIKE '*値*'` を ADO Filter と同じ規則で(大文字小文字を区別しない・Null は空・' は構文エラー)。
                合わない行は先に外す(VBA の SQL で絞った行に当たる)
    含む(Like)  表のセル Like "*値*"(VBA の Like。大文字小文字を区別する)
第4段(2026-11-11): 外観_厚板(Q_S_QA_124_別表2)・外観_薄板(SRF_128_別表1)・ラテラル_日車 / 白銅滑川(DIM_179)・
バーコール(Q_S_KE_ALL_053)。どの表を使うか(オーダー厚 8・画面の文字)と後処理(欠陥グレードの上書き・パターン 1/2・
区間列の LoopArray・最後の行の下限値)はコードに残す。外観_薄板 の VBA は入れ子の If(短絡)だが、Like が例外を
出す書き方のグレード(閉じない [ など)でなければ結果は同じ。
前処理: 右3文字 = ロットの値が 4 文字以上なら右 3 文字(BhSearch)/ 列が無ければ通す = 表にその列が無ければ条件なし
(124 別表2 の 調質。B25 を直した側)
第3段(測定機器ID・設備仕様)の 範囲 は 空のとき=通さない(check_within: 空は 0)。B52 を再現するなら 前処理=Single丸め。
前処理: 範囲 で「Single丸め」= 下・上を Single にして比べる(B71 の再現側)。既定は なし(Double。いまのコード)
空のとき: 条件式 で「通さない」= 表のセルが空でも通さない(blank_flag)

VBA の And は短絡しないので、行ごとに全部のルールを評価する(列が無ければ添字エラーになるのも同じ)。
ただし 当板用か は VBA でも先に `If … Then GoTo` で外しているので、合わない行はそこで止める。

旧コードの道(quality.py の _ratefuri_old / _pack_old)は、本番の shadow で差が 0 件と確かめたら消す。
"""
from __future__ import annotations

import functools
import sqlite3
import threading
import time
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Any, Mapping, Optional, Sequence

from .. import safe_io
from ..logging_utils import get_logger
from . import cache_registry, datasource

log = get_logger("kensa.match_rules")

DB = "伝送用ファイル"
TABLE = "照合ルール"
GROUP_TABLE = "照合ルール群"
COLUMNS = ("ルール群", "順番", "表の列", "ロットの項目", "比べ方", "前処理", "空のとき", "有効", "備考")
GROUP_COLUMNS = ("ルール群", "使う行", "新方式", "備考")
#: 照合ルールの版(ルール群ごと。保存のたびに 1 増える。表が無い・一度も保存していない = 0 = 既定)
VERSION_TABLE = "照合ルールの版"
USE_EV = "match_rules.used"
CHANGE_EV = "match_rules.change"
TTL_SEC = 5.0

COMPARES = ("条件式", "範囲", "以上", "一致", "当板用か", "ライン", "含む条件式", "含む(SQL)", "含む(Like)")
#: 合わない行を先に外す比べ方(VBA の GoTo / ADO Filter。最後に評価した行には数えない)
PREFILTER = ("当板用か", "ライン", "含む(SQL)")
PREPS = ("なし", "Single丸め", "右3文字", "列が無ければ通す")
BLANKS = ("通す", "通さない")
USE_ROWS = ("最初", "最後", "全部")

RATEFURI = "ラテフリ直角"
PACK = "梱包保護材"
QP = "平坦度一覧QP追加"
MICRO = "測定機器ID"
EQUIP = "設備仕様"
APPEAR_H = "外観_厚板"
APPEAR_S = "外観_薄板"
LATERAL_N = "ラテラル_日車"
LATERAL_H = "ラテラル_白銅滑川"
BH = "バーコール"

#: 試験用: 新旧を両方走らせ、違えば例外(tests/_isolation で真にする)
SHADOW_STRICT = False


@dataclass(frozen=True)
class Rule:
    order: int
    columns: tuple[str, ...]
    field: str
    compare: str
    prep: str = "なし"
    blank: str = "通す"
    enabled: bool = True
    note: str = ""


@dataclass(frozen=True)
class Group:
    name: str
    table: str                 # 引く表(品質基準 / 梱包資材マスタ)
    source: str                # 表の DB の説明(画面用)
    fields: tuple[str, ...]    # 選べるロットの項目
    use_rows: str
    new_way: bool
    rules: tuple[Rule, ...]
    note: str = ""
    version: int = 0           # 照合ルールの版(保存のたびに 1 増える)


def _r(order: int, cols: str, fld: str, cmp: str, note: str = "", blank: str = "通す") -> Rule:
    return Rule(order, tuple(cols.split(",")), fld, cmp, blank=blank, note=note)


#: DbReturnItem の引数(引当の行。Lot情報履歴 / SIKAHIKINOW + 仕掛の製造寸法)。既定のルールはこれだけを見る
RATEFURI_ARGS = ("納入先", "取引先", "用途ｺｰﾄﾞ", "製造材質", "製造調質", "製造板厚", "製造板幅", "製造板丈")
#: ロットの行(SIKALOT 仕掛)から足す項目。オーダー寸法・BOX実績の寸法に切り替えられるように(2026-11-11)
RATEFURI_LOT_FIELDS = ("ｵｰﾀﾞｰ材質", "ｵｰﾀﾞｰ調質", "ｵｰﾀﾞｰ板厚", "ｵｰﾀﾞｰ板幅", "ｵｰﾀﾞｰ板丈",
                       "BOX実績_板厚", "BOX実績_板幅", "BOX実績_板丈")
#: ラテフリ直角 で選べるロットの項目
RATEFURI_FIELDS = RATEFURI_ARGS + RATEFURI_LOT_FIELDS
#: 梱包保護材 で選べるロットの項目(Lot情報履歴 の見出し。+ 当板用か)
PACK_FIELDS = ("包装仕様NO", "材質", "調質", "用途ｺｰﾄﾞ", "厚", "巾", "丈", "ﾛｯﾄNo", "ｵｰﾀﾞｰNo", "引当NO", "当板用")

#: 平坦度一覧QP追加 で選べるロットの項目(仕掛の行 + 引当の納入先・取引先。QP.DBQP)
QP_FIELDS = ("用途ｺｰﾄﾞ", "製造材質", "ｵｰﾀﾞｰ材質", "製造調質", "ｵｰﾀﾞｰ調質", "ｵｰﾀﾞｰ板厚", "ｵｰﾀﾞｰ板幅", "ｵｰﾀﾞｰ板丈",
             "製造板厚", "製造板幅", "製造板丈", "BOX実績_板厚", "BOX実績_板幅", "BOX実績_板丈",
             "引当_納入先", "納入先ｺｰﾄﾞ", "引当_取引先", "取引先ｺｰﾄﾞ", "送り先名", "送り先ｺｰﾄﾞ", "BOX設計_横割数(板)")
#: 測定機器ID で選べる項目(MachineTransfer.DbReturnMicro の引数: ライン・画面の厚・用途)
MICRO_FIELDS = ("ライン", "厚", "用途")
#: 設備仕様 で選べる項目(QP.CheckEquipmentDb の引数: 端末の設備名・仕掛の製造材質・オーダー寸法・製品単重)
EQUIP_FIELDS = ("ライン", "製造材質", "オーダー厚", "オーダー幅", "オーダー丈", "単重")

#: 外観_厚板・外観_薄板・ラテラル で選べるロットの項目(仕掛の行。QDB.DBdisplay外観_H / 外観_S / DBdisplayラテラル)
SIKA_FIELDS = ("用途ｺｰﾄﾞ", "製造材質", "製造調質", "ｵｰﾀﾞｰ材質", "ｵｰﾀﾞｰ調質", "製造板厚", "製造板幅", "製造板丈",
               "ｵｰﾀﾞｰ板厚", "ｵｰﾀﾞｰ板幅", "ｵｰﾀﾞｰ板丈", "BOX実績_板厚", "BOX実績_板幅", "BOX実績_板丈",
               "納入先名", "納入先ｺｰﾄﾞ", "品質ｸﾞﾚｰﾄﾞ_表面処理")
#: バーコール で選べる項目(QDB.BhSearch の引数: 画面の材質・調質)
BH_FIELDS = ("材質", "調質")

#: 平坦度一覧QP追加 の既定のルールの見出し(歪公差チェックの trace に出す。順番 → 名前)
QP_LABELS = {10: "用途コード", 20: "製造材質", 30: "オーダー材質", 40: "製造調質", 50: "オーダー調質",
             60: "板厚(下〜上)", 70: "板幅(下〜上)", 80: "板丈(下〜上)", 90: "納入先", 100: "納入先コード",
             110: "取引先", 120: "取引先コード", 130: "送り先", 140: "送り先コード", 150: "巾取数"}

DEFAULT_GROUPS: tuple[Group, ...] = (
    Group(RATEFURI, "ラテフリ直角", "品質基準", RATEFURI_FIELDS, "全部", True, (
        _r(10, "納入先", "納入先", "条件式", "QP:1252"),
        _r(20, "取引先", "取引先", "条件式", "QP:1253"),
        _r(30, "用途コード", "用途ｺｰﾄﾞ", "条件式", "QP:1254"),
        _r(40, "製造材質", "製造材質", "条件式", "QP:1255"),
        _r(50, "製造調質", "製造調質", "条件式", "QP:1256"),
        _r(60, "厚下,厚上,製造板厚", "製造板厚", "範囲", "QP:1263"),
        _r(70, "幅下,幅上,製造板幅", "製造板幅", "範囲", "QP:1264"),
        _r(80, "丈下,丈上,製造板丈", "製造板丈", "範囲", "QP:1265"),
    ), "一致した行は全部使う(管理番号で引き直す)。該当なしの文言は最後の行で最初に外れたルールで決める"),
    Group(PACK, "梱包保護材", "梱包資材マスタ", PACK_FIELDS, "最初", True, (
        _r(10, "備考", "当板用", "当板用か", "QP:1753"),
        _r(20, "包装仕様書", "包装仕様NO", "条件式", "QP:1758"),
        _r(30, "材質", "材質", "条件式", "QP:1759"),
        _r(40, "調質", "調質", "条件式", "QP:1760"),
        _r(50, "用途コード", "用途ｺｰﾄﾞ", "条件式", "QP:1761"),
        _r(60, "板厚下,板厚上,板厚", "厚", "範囲", "QP:1764"),
        _r(70, "板幅下,板幅上,板幅", "巾", "範囲", "QP:1765"),
        _r(80, "板丈下,板丈上,板丈", "丈", "範囲", "QP:1766"),
    ), "最初に一致した行の 使用保護材"),
    Group(QP, "平坦度一覧QP追加", "品質基準", QP_FIELDS, "全部", True, (
        _r(10, "用途コード", "用途ｺｰﾄﾞ", "条件式", "DBQP 条件1"),
        _r(20, "製造材質", "製造材質", "条件式", "DBQP 条件2"),
        _r(30, "オーダー材質", "ｵｰﾀﾞｰ材質", "条件式", "DBQP 条件3"),
        _r(40, "製造調質", "製造調質", "条件式", "DBQP 条件4"),
        _r(50, "オーダー調質", "ｵｰﾀﾞｰ調質", "条件式", "DBQP 条件5"),
        _r(60, "板厚下,板厚上,オーダー板厚", "ｵｰﾀﾞｰ板厚", "範囲", "DBQP 条件6"),
        _r(70, "板幅下,板幅上,オーダー板幅", "ｵｰﾀﾞｰ板幅", "範囲", "DBQP 条件7"),
        _r(80, "板丈下,板丈上,オーダー板丈", "ｵｰﾀﾞｰ板丈", "範囲", "DBQP 条件8"),
        _r(90, "納入先", "引当_納入先", "条件式", "DBQP 条件9"),
        _r(100, "納入先コード", "納入先ｺｰﾄﾞ", "条件式", "DBQP 条件10"),
        _r(110, "取引先", "引当_取引先", "条件式", "DBQP 条件11"),
        _r(120, "取引先コード", "取引先ｺｰﾄﾞ", "条件式", "DBQP 条件12"),
        _r(130, "送り先", "送り先名", "条件式", "DBQP 条件13"),
        _r(140, "送り先コード", "送り先ｺｰﾄﾞ", "条件式", "DBQP 条件14"),
        _r(150, "巾取数幅分割数のこと", "BOX設計_横割数(板)", "以上", "DBQP 条件15"),
    ), "検査番号の頭(検番の頭か空)で先に絞った行に当てる。一致した行は全部。納入先などが書いてある行を優先(コード)"),
    Group(MICRO, "測定機器ID", "伝送用ファイル", MICRO_FIELDS, "最初", True, (
        _r(10, "ライン", "ライン", "ライン", "MT:431 [ライン]='<ライン>'(D13)"),
        _r(20, "用途", "用途", "含む条件式", "MT:516-545 CheckThickAndCode"),
        _r(30, "厚下,厚上", "厚", "範囲", "MT:547・489-490(B52)", "通さない"),
    ), "最初に合う行の出力テキスト(機器名・ID)。用途の列がこのラインで全部空・用途で合う行が無いときは、"
       "含む条件式 のルールを外して(厚だけで)引き直す。デプスなどの機器名指定はラインだけで引く(コード)"),
    Group(EQUIP, "設備仕様", "伝送用ファイル", EQUIP_FIELDS, "全部", True, (
        _r(10, "設備名", "ライン", "ライン", "QP:2105(D13)"),
        _r(20, "製造材質", "製造材質", "条件式", "QP:2136"),
        _r(30, "オーダー厚MIN,オーダー厚MAX", "オーダー厚", "範囲", "QP:2140", "通さない"),
        _r(40, "オーダー幅MIN,オーダー幅MAX", "オーダー幅", "範囲", "QP:2141", "通さない"),
        _r(50, "オーダー丈MIN,オーダー丈MAX", "オーダー丈", "範囲", "QP:2142", "通さない"),
        _r(60, "単重MIN,単重MAX", "単重", "範囲", "QP:2143", "通さない"),
    ), "設備名の行が無ければ 可。合う行が無ければ最後の行で最初に外れたルールで「… 範囲外」。"
       "合えば空でない作業可否の最後のもの(コード)"),
)
BY_NAME = {g.name: g for g in DEFAULT_GROUPS}


def _lateral(name: str, table: str) -> Group:
    return Group(name, table, "品質基準", SIKA_FIELDS, "全部", True, (
        _r(10, "材質", "製造材質", "条件式", "QDB:1630"),
        _r(20, "板厚下,板厚上,厚", "製造板厚", "範囲", "QDB:1637(B71 は直した側 = なし)"),
        _r(30, "板幅下,板幅上,幅", "製造板幅", "範囲", "QDB:1638"),
    ), "一致した行は全部(管理番号で引き直す)。値は 製造丈 の区間列 × 製造厚 で LoopArray(コード)。"
       "日車 / 白銅滑川 のどちらの表を使うかは画面の文字で決める(コード)")


#: 第4段(2026-11-11)。SQL の LIKE による先の絞り込みは 含む(SQL)(先に外す)にした
DEFAULT_GROUPS = DEFAULT_GROUPS + (
    Group(APPEAR_H, "Q_S_QA_124_別表2", "品質基準", SIKA_FIELDS, "全部", True, (
        _r(10, "用途コード", "用途ｺｰﾄﾞ", "含む(SQL)", "QDB:213 [用途コード] LIKE '*用途*'"),
        _r(20, "材質", "製造材質", "条件式", "QDB:296"),
        replace(_r(30, "調質", "製造調質", "条件式", "QDB:298 B25(いまの別表2 には列が無い)"), prep="列が無ければ通す"),
        _r(40, "板厚下,板厚上,厚", "製造板厚", "範囲", "QDB:300 製造厚"),
        _r(50, "板幅下,板幅上,幅", "ｵｰﾀﾞｰ板幅", "範囲", "QDB:301 オーダー幅"),
        _r(60, "板丈下,板丈上,丈", "ｵｰﾀﾞｰ板丈", "範囲", "QDB:302 オーダー丈"),
        _r(70, "納入先", "納入先名", "条件式", "QDB:303"),
    ), "オーダー厚 8 以上のとき(コード)。一致した行は全部、後の管理番号の欠陥グレードで上書き(コード)"),
    Group(APPEAR_S, "Q_S_QA_SRF_128_別表1", "品質基準", SIKA_FIELDS, "最初", True, (
        _r(10, "用途コード", "用途ｺｰﾄﾞ", "含む(SQL)", "QDB:560 [用途コード] LIKE '*用途*'"),
        _r(20, "材質", "製造材質", "条件式", "QDB:598"),
        _r(30, "表面処理コード", "品質ｸﾞﾚｰﾄﾞ_表面処理", "含む(Like)", "QDB:599 Like *グレード*"),
    ), "オーダー厚 8 未満のとき(コード)。最初に合う行の パターン(厚フラグがあれば 板厚1 とオーダー厚で 1/2。コード)。"
       "用途で残る行が無ければ「該当用途ｺｰﾄﾞなし」"),
    _lateral(LATERAL_N, "Q_S_QA_DIM_179_日車"),
    _lateral(LATERAL_H, "Q_S_QA_DIM_179_白銅滑川"),
    Group(BH, "Q_S_KE_ALL_053", "品質基準", BH_FIELDS, "全部", True, (
        replace(_r(10, "材質", "材質", "含む(SQL)", "QDB:1479-1481 4 文字以上は右3文字"), prep="右3文字"),
        _r(20, "調質", "調質", "一致", "QDB:1484"),
    ), "一致した行を順に読み、最後の行の バーコール下限値(コード)。無ければ 0(BH 不要)"),
)
BY_NAME = {g.name: g for g in DEFAULT_GROUPS}

#: ロットの項目の出どころ(画面に出す見出し)
FIELD_SOURCE: dict[str, str] = {
    "納入先": "引当", "取引先": "引当", "用途ｺｰﾄﾞ": "引当",
    "製造材質": "製造", "製造調質": "製造", "製造板厚": "製造", "製造板幅": "製造", "製造板丈": "製造",
    "ｵｰﾀﾞｰ材質": "オーダー", "ｵｰﾀﾞｰ調質": "オーダー", "ｵｰﾀﾞｰ板厚": "オーダー", "ｵｰﾀﾞｰ板幅": "オーダー",
    "ｵｰﾀﾞｰ板丈": "オーダー", "BOX実績_板厚": "実績", "BOX実績_板幅": "実績", "BOX実績_板丈": "実績",
    "包装仕様NO": "引当", "材質": "引当", "調質": "引当", "厚": "引当", "巾": "引当", "丈": "引当",
    "ﾛｯﾄNo": "引当", "ｵｰﾀﾞｰNo": "引当", "引当NO": "引当", "当板用": "引当",
    "引当_納入先": "引当", "引当_取引先": "引当", "納入先ｺｰﾄﾞ": "仕掛", "取引先ｺｰﾄﾞ": "仕掛", "送り先名": "仕掛",
    "送り先ｺｰﾄﾞ": "仕掛", "BOX設計_横割数(板)": "仕掛",
    "ライン": "端末", "厚": "画面", "用途": "画面", "オーダー厚": "仕掛", "オーダー幅": "仕掛", "オーダー丈": "仕掛",
    "単重": "仕掛",
    "納入先名": "仕掛", "納入先ｺｰﾄﾞ": "仕掛", "品質ｸﾞﾚｰﾄﾞ_表面処理": "仕掛",
}
#: 群ごとに出どころを変える項目(製造材質 は 設備仕様では仕掛の行)
FIELD_SOURCE_BY_GROUP: dict[str, dict[str, str]] = {EQUIP: {"製造材質": "仕掛"},
                                                   BH: {"材質": "画面", "調質": "画面"}}
#: 「製造寸法 ⇔ オーダー寸法をまとめて切り替え」で入れ替える組(寸法だけ)
DIM_SWAP: dict[str, tuple[tuple[str, str], ...]] = {
    RATEFURI: (("製造板厚", "ｵｰﾀﾞｰ板厚"), ("製造板幅", "ｵｰﾀﾞｰ板幅"), ("製造板丈", "ｵｰﾀﾞｰ板丈")),
    QP: (("製造板厚", "ｵｰﾀﾞｰ板厚"), ("製造板幅", "ｵｰﾀﾞｰ板幅"), ("製造板丈", "ｵｰﾀﾞｰ板丈")),
}
for _n in (APPEAR_H, LATERAL_N, LATERAL_H):
    DIM_SWAP[_n] = DIM_SWAP[QP]
HELP_SWITCH = ("参照する寸法を変えるには:ロットの項目を『製造板厚』→『ｵｰﾀﾞｰ板厚』に選び直し、"
               "試し照合で行が変わるか確かめてから保存")


def field_source(name: str, group: str = "") -> str:
    return FIELD_SOURCE_BY_GROUP.get(group, {}).get(name) or FIELD_SOURCE.get(name, "")


def split_pre(g: Group) -> tuple[Group, Group]:
    """(先に外すルールだけの群, 残りのルールの群)。"""
    return (replace(g, rules=tuple(r for r in g.rules if r.compare in PREFILTER), use_rows="全部"),
            replace(g, rules=tuple(r for r in g.rules if r.compare not in PREFILTER)))


def shadow_on(g: Group) -> bool:
    """旧コードと並べるか(試験で、既定のルールのままのとき。ルールを変えたら違って当然なので並べない)。"""
    if SHADOW_STRICT:
        return g.rules == BY_NAME[g.name].rules
    return False


def swap_dims(g: Group) -> Group:
    """寸法のルールのロットの項目を 製造 ⇔ オーダー でまとめて入れ替える(保存はしない)。"""
    pairs = DIM_SWAP.get(g.name, ())
    m = {a: b for a, b in pairs} | {b: a for a, b in pairs}
    return replace(g, rules=tuple(replace(r, field=m.get(r.field, r.field)) for r in g.rules))


def lot_ctx(fields: Optional[Sequence[str]], record: Optional[Sequence[Any]]) -> dict[str, Any]:
    """ロットの行(仕掛)から照合に足す項目。列が無い・行が無いときは入れない。"""
    out: dict[str, Any] = {}
    if not fields or not record:
        return out
    idx = {str(f): i for i, f in enumerate(fields)}
    for name in RATEFURI_LOT_FIELDS:
        i = idx.get(name)
        if i is not None and i < len(record):
            v = record[i]
            out[name] = "" if v is None else v
    return out


class Invalid(ValueError):
    def __init__(self, message: str, group: str = "", order: int = 0):
        super().__init__(message)
        self.group, self.order = group, order


# ------------------------------------------------------------------
# 照合
# ------------------------------------------------------------------
@dataclass
class Step:
    rule: Rule
    remain: list[int]          # このルールまでで残った行(0 始まりの行位置)


@dataclass
class MatchResult:
    rows: list[int]            # 使う行(使う行の設定を当てたあと。0 始まり)
    matched: list[int]         # 全部のルールを通った行
    steps: list[Step] = field(default_factory=list)
    last_fail: Optional[int] = None      # 最後に評価した行で、最初に外れたルールの位置(0 始まり。通った・行が無い → None)
    evaluated: int = 0


@functools.lru_cache(maxsize=512)
def _sql_like(col: str, value: str) -> Any:
    """`[列] LIKE '*値*'` を ADO Filter と同じ規則で(datasource の絞り込みと同じ関数。' があれば構文エラー)。"""
    from . import ado_filter
    return ado_filter.compile_filter("[" + col + "] LIKE '*" + value + "*'")


def prep_value(rule: Rule, value: Any) -> Any:
    """前処理 右3文字: 4 文字以上なら右 3 文字(QDB.BhSearch)。"""
    if rule.prep == "右3文字":
        from . import quality as q, vbx
        v = q.vstr(value)
        return vbx.right(v, 3) if len(v) >= 4 else v
    return value


def _test(rule: Rule, row: Sequence[Any], fields: Sequence[str], value: Any) -> bool:
    from . import quality as q
    cols = [q.col_of(fields, c) for c in rule.columns]
    cmp = rule.compare
    if rule.prep == "列が無ければ通す" and not all(cols):
        return True
    value = prep_value(rule, value)
    if cmp == "含む(SQL)":
        return _sql_like(rule.columns[0], q.vstr(value))(dict(zip(fields, row)))
    if cmp == "含む(Like)":
        from . import vbx
        return vbx.like(q.vstr(q.cell(row, cols[0])), "*" + q.vstr(value) + "*")
    if cmp == "条件式":
        return q.check_condition(row, cols[0], q.vstr(value), "" if rule.blank == "通す" else "1")
    if cmp == "範囲":
        same = rule.prep != "Single丸め"
        x = q.csng(value)
        if rule.blank == "通さない":
            return q.check_within(row, cols[0], cols[1], x, same_precision=same)
        return q.check_thickness(row, cols[0], cols[1], cols[2] if len(cols) > 2 else cols[1], x, same_precision=same)
    if cmp == "以上":
        return q.check_size(row, cols[0], value)
    if cmp == "一致":
        return q.vstr(q.cell(row, cols[0])) == q.vstr(value)
    if cmp == "当板用か":
        return (q.vstr(q.cell(row, cols[0])) == "当板用") == (q.vstr(value) == "当板用")
    if cmp == "ライン":
        from . import lines
        return lines.matches(q.cell(row, cols[0]), q.vstr(value))
    if cmp == "含む条件式":
        from . import instruments
        return instruments.youto_match(q.cell(row, cols[0]), q.vstr(value))
    raise Invalid(f"比べ方「{cmp}」は使えません")


def match(group: Group, rows: Sequence[Sequence[Any]], fields: Sequence[str], ctx: Mapping[str, Any]) -> MatchResult:
    """rows(表の行)を group のルールで絞る。ctx はロットの項目 → 値。"""
    rules = [r for r in sorted(group.rules, key=lambda r: r.order) if r.enabled]
    steps = [Step(r, []) for r in rules]
    matched: list[int] = []
    last_fail: Optional[int] = None
    evaluated = 0
    for i, row in enumerate(rows):
        ok = True
        first_fail: Optional[int] = None
        skipped = False
        for k, r in enumerate(rules):
            t = _test(r, row, fields, ctx.get(r.field, ""))
            if not t and r.compare in PREFILTER:
                skipped = True            # VBA でも先に外す(GoTo)。Conditions は前の行のまま
                break
            if not t and first_fail is None:
                first_fail = k
            ok = ok and t
            if ok:
                steps[k].remain.append(i)
        if skipped:
            continue
        evaluated += 1
        last_fail = first_fail
        if ok:
            matched.append(i)
    use = matched
    if matched and group.use_rows == "最初":
        use = matched[:1]
    elif matched and group.use_rows == "最後":
        use = matched[-1:]
    return MatchResult(use, matched, steps, last_fail, evaluated)


# ------------------------------------------------------------------
# 読み書き
# ------------------------------------------------------------------
_lock = threading.Lock()
_last_known: dict[str, tuple[float, dict[str, Group]]] = {}
cache_registry.register_dict("match_rules._last_known", _last_known, _lock)


def _db_path() -> Path:
    return datasource.resolve(DB)


def _truthy(v: Any) -> bool:
    return str(v).strip() in ("1", "True", "true", "-1", "✓")


def _has(con: sqlite3.Connection, name: str) -> bool:
    return con.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?", (name,)).fetchone() is not None


def _rule_rows(groups: Sequence[Group]) -> list[tuple]:
    return [(g.name, r.order, ",".join(r.columns), r.field, r.compare, r.prep, r.blank, 1 if r.enabled else 0, r.note)
            for g in groups for r in g.rules]


def _group_rows(groups: Sequence[Group]) -> list[tuple]:
    return [(g.name, g.use_rows, 1 if g.new_way else 0, g.note) for g in groups]


def _create(con: sqlite3.Connection) -> None:
    con.execute(f'CREATE TABLE IF NOT EXISTS "{TABLE}" ("ルール群" TEXT NOT NULL, "順番" INTEGER NOT NULL, '
                f'"表の列" TEXT NOT NULL, "ロットの項目" TEXT NOT NULL, "比べ方" TEXT NOT NULL, '
                f'"前処理" TEXT NOT NULL DEFAULT \'なし\', "空のとき" TEXT NOT NULL DEFAULT \'通す\', '
                f'"有効" INTEGER NOT NULL DEFAULT 1, "備考" TEXT)')
    con.execute(f'CREATE TABLE IF NOT EXISTS "{GROUP_TABLE}" ("ルール群" TEXT NOT NULL PRIMARY KEY, '
                f'"使う行" TEXT NOT NULL, "新方式" INTEGER NOT NULL DEFAULT 1, "備考" TEXT)')
    con.execute(f'CREATE TABLE IF NOT EXISTS "{VERSION_TABLE}" ("ルール群" TEXT NOT NULL PRIMARY KEY, '
                f'"版" INTEGER NOT NULL DEFAULT 0, "日時" TEXT)')


def ensure_table(path: Optional[Path] = None) -> bool:
    """表が無ければ作って既定のルール(いまのコードと同じ)を入れる。足りない群も足す。"""
    p = path or _db_path()
    if not p.is_file():
        return False
    try:
        con = safe_io.connect_existing(p, timeout=5)
        try:
            with con:
                _create(con)
                have = {r[0] for r in con.execute(f'SELECT "ルール群" FROM "{GROUP_TABLE}"')}
                missing = [g for g in DEFAULT_GROUPS if g.name not in have]
                if missing:
                    con.execute(f'DELETE FROM "{TABLE}" WHERE "ルール群" IN ({",".join("?" * len(missing))})',
                                [g.name for g in missing])
                    con.executemany(f'INSERT INTO "{TABLE}" VALUES (?,?,?,?,?,?,?,?,?)', _rule_rows(missing))
                    con.executemany(f'INSERT INTO "{GROUP_TABLE}" VALUES (?,?,?,?)', _group_rows(missing))
                    log.info("照合ルールの既定を入れました: %s %s", p, [g.name for g in missing])
            return True
        finally:
            con.close()
    except sqlite3.Error as e:
        log.warning("照合ルールのマスタを作れませんでした(%s): %s", p, e)
        return False


def _read(p: Path) -> Optional[dict[str, Group]]:
    """読めたら群 → Group(表が無ければ既定)。読めなければ None。"""
    try:
        if not p.is_file():
            return None
        con = safe_io.connect_existing(p, timeout=5)
        try:
            if not _has(con, TABLE) or not _has(con, GROUP_TABLE):
                return dict(BY_NAME)
            rules = con.execute(f'SELECT {",".join(chr(34) + c + chr(34) for c in COLUMNS)} FROM "{TABLE}"').fetchall()
            grows = con.execute(f'SELECT {",".join(chr(34) + c + chr(34) for c in GROUP_COLUMNS)} '
                                f'FROM "{GROUP_TABLE}"').fetchall()
            versions: dict[str, int] = {}
            if _has(con, VERSION_TABLE):
                for vn, vv in con.execute(f'SELECT "ルール群", "版" FROM "{VERSION_TABLE}"'):
                    try:
                        versions[str(vn)] = int(vv or 0)
                    except (TypeError, ValueError):
                        pass
        finally:
            con.close()
    except (OSError, sqlite3.Error) as e:
        log.warning("照合ルールを読めません(%s): %s", p, e)
        return None
    out = dict(BY_NAME)
    for name, use, new, note in grows:
        base = BY_NAME.get(str(name))
        if base is None:
            continue
        mine = []
        for r in rules:
            if str(r[0]) != base.name:
                continue
            try:
                mine.append(Rule(int(r[1]), tuple(c.strip() for c in str(r[2]).split(",")), str(r[3]), str(r[4]),
                                 str(r[5] or "なし"), str(r[6] or "通す"), _truthy(r[7]), str(r[8] or "")))
            except (TypeError, ValueError):
                log.warning("照合ルールの行が悪いので飛ばします: %s", r)
        g = replace(base, use_rows=str(use) if str(use) in USE_ROWS else base.use_rows, new_way=_truthy(new),
                    rules=tuple(mine), note=str(note or ""), version=versions.get(base.name, 0))
        try:
            check(g)
        except Invalid as e:
            log.warning("照合ルール「%s」が悪いので既定を使います: %s", base.name, e)
            g = replace(base, version=versions.get(base.name, 0))
        out[base.name] = g
    return out


def groups(path: Optional[Path] = None, *, fresh: bool = False) -> dict[str, Group]:
    p = path or _db_path()
    key, now = str(p), time.monotonic()
    with _lock:
        hit = _last_known.get(key)
    if hit and not fresh and now - hit[0] < TTL_SEC:
        return dict(hit[1])
    got = _read(p)
    if got is None:
        return dict(hit[1]) if hit else dict(BY_NAME)
    with _lock:
        _last_known[key] = (now, got)
    return dict(got)


def group(name: str) -> Group:
    return groups().get(name) or BY_NAME[name]


def check(g: Group, table_fields: Optional[Sequence[str]] = None) -> None:
    """ルールの形を確かめる(表の列を渡すと、その表に在る列だけにする)。悪ければ Invalid。"""
    if g.use_rows not in USE_ROWS:
        raise Invalid(f"使う行「{g.use_rows}」は 最初・最後・全部 のどれかです", g.name)
    seen = set()
    for r in g.rules:
        if r.order in seen:
            raise Invalid(f"順番 {r.order} が重なっています", g.name, r.order)
        seen.add(r.order)
        if r.compare not in COMPARES:
            raise Invalid(f"比べ方「{r.compare}」は使えません", g.name, r.order)
        if r.prep not in PREPS or r.blank not in BLANKS:
            raise Invalid("前処理・空のときの値が悪いです", g.name, r.order)
        if r.field not in g.fields:
            raise Invalid(f"ロットの項目「{r.field}」はこの群では選べません", g.name, r.order)
        need = (2, 3) if r.compare == "範囲" else (1,)
        if len(r.columns) not in need or not all(r.columns):
            raise Invalid("表の列の数が合いません(範囲 は 下・上・値 の 3 列、ほかは 1 列)", g.name, r.order)
        if table_fields is not None and r.prep != "列が無ければ通す":
            bad = [c for c in r.columns if c not in table_fields]
            if bad:
                raise Invalid(f"表「{g.table}」に列「{'・'.join(bad)}」がありません", g.name, r.order)


def from_payload(name: str, d: Mapping[str, Any]) -> Group:
    base = BY_NAME.get(name)
    if base is None:
        raise Invalid(f"ルール群「{name}」はありません", name)
    rules = []
    for x in d.get("rules") or []:
        try:
            order = int(x.get("order"))
        except (TypeError, ValueError):
            raise Invalid("順番は整数です", name) from None
        cols = x.get("columns") or []
        if isinstance(cols, str):
            cols = cols.split(",")
        rules.append(Rule(order, tuple(str(c).strip() for c in cols), str(x.get("field", "")),
                          str(x.get("compare", "")), str(x.get("prep") or "なし"), str(x.get("blank") or "通す"),
                          bool(x.get("enabled", True)), str(x.get("note") or "")[:200]))
    return replace(base, use_rows=str(d.get("use_rows", base.use_rows)), new_way=bool(d.get("new_way", True)),
                   rules=tuple(sorted(rules, key=lambda r: r.order)), note=str(d.get("note") or "")[:200])


def to_payload(g: Group) -> dict[str, Any]:
    return {"name": g.name, "table": g.table, "source": g.source, "fields": list(g.fields), "use_rows": g.use_rows,
            "new_way": g.new_way, "note": g.note, "version": g.version,
            "rules": [{"order": r.order, "columns": list(r.columns), "field": r.field, "compare": r.compare,
                       "prep": r.prep, "blank": r.blank, "enabled": r.enabled, "note": r.note} for r in g.rules]}


def save(g: Group, path: Optional[Path] = None, *, reason: str = "") -> Group:
    p = path or _db_path()
    check(g)
    before = groups(p, fresh=True).get(g.name)
    g = replace(g, version=(before.version if before else 0) + 1)
    if not ensure_table(p):
        raise OSError(f"伝送用ファイルに書けません: {p}")
    con = safe_io.connect_existing(p, timeout=10)
    try:
        with con:
            con.execute(f'DELETE FROM "{TABLE}" WHERE "ルール群" = ?', (g.name,))
            con.executemany(f'INSERT INTO "{TABLE}" VALUES (?,?,?,?,?,?,?,?,?)', _rule_rows([g]))
            con.execute(f'INSERT OR REPLACE INTO "{GROUP_TABLE}" VALUES (?,?,?,?)', _group_rows([g])[0])
            con.execute(f'INSERT OR REPLACE INTO "{VERSION_TABLE}" VALUES (?,?,datetime(\'now\', \'localtime\'))',
                        (g.name, g.version))
    finally:
        con.close()
    with _lock:
        _last_known.clear()
    try:
        from .. import trace
        trace.event(CHANGE_EV, f"保守: 照合ルール「{g.name}」を変更",
                    data={"group": g.name, "reason": reason, "version": g.version,
                          "summary": _summary(to_payload(before) if before else None, to_payload(g))})
    except Exception as e:  # noqa: BLE001 - 履歴が残せなくても保存は済んでいる
        log.warning("照合ルールの変更履歴を残せませんでした: %s", e)
    return g


def versions() -> dict[str, int]:
    """いまの照合ルールの版(ルール群 → 版)。"""
    return {n: g.version for n, g in groups().items()}


def record_use(lot: str, where: str = "") -> dict[str, int]:
    """検査の保存のとき、使った照合ルールの版を記録に残す(ufkensa_記録 の jsonl。保存用の列は変えない)。"""
    v = {}
    try:
        v = versions()
        from .. import trace
        trace.event(USE_EV, f"照合ルールの版: " + "・".join(f"{k} 版{n}" for k, n in v.items()) + f"(ロット {lot})",
                    data={"lot": lot, "where": where, "versions": v})
    except Exception as e:  # noqa: BLE001 - 記録が残せなくても保存は済んでいる
        log.warning("照合ルールの版を記録できませんでした: %s", e)
    return v


def history(days: int = 90, limit: int = 30) -> list[dict[str, Any]]:
    from .. import trace
    out = []
    for r in trace.read(days, q=CHANGE_EV, limit=2000):
        if r.get("ev") != CHANGE_EV:
            continue
        d = r.get("data") or {}
        if not isinstance(d, dict):
            continue
        out.append({"t": r.get("t", ""), "user": r.get("user", ""), "pc": r.get("pc", ""), "group": d.get("group", ""),
                    "version": d.get("version"),
                    "reason": d.get("reason", ""), "summary": d.get("summary") or []})
        if len(out) >= limit:
            break
    return out


def _summary(before: Any, after: Any) -> list[str]:
    def lines(p: Any) -> dict[int, str]:
        if not isinstance(p, dict):
            return {}
        return {int(r["order"]): f"{','.join(r['columns'])} ← {r['field']} : {r['compare']}"
                + ("" if r.get("enabled", True) else "(無効)") for r in p.get("rules") or []}
    a, b = lines(before), lines(after)
    out = [f"順番{k}: {a.get(k, '(なし)')} → {b.get(k, '(なし)')}" for k in sorted(set(a) | set(b)) if a.get(k) != b.get(k)]
    if isinstance(before, dict) and isinstance(after, dict):
        for key, lab in (("use_rows", "使う行"), ("new_way", "新方式")):
            if before.get(key) != after.get(key):
                out.append(f"{lab}: {before.get(key)} → {after.get(key)}")
    return out


# ------------------------------------------------------------------
# 新旧を並べる(shadow)
# ------------------------------------------------------------------
class ShadowMismatch(AssertionError):
    pass


def compare(name: str, old: Any, new: Any, detail: str = "") -> None:
    """新旧の結果が違えばログ(試験では例外)。"""
    if old == new:
        return
    msg = f"[照合ルール] {name}: 旧 {old!r} / 新 {new!r} {detail}"
    log.warning(msg)
    try:
        from .. import trace
        trace.event("match_rules.diff", msg)
    except Exception:  # noqa: BLE001
        pass
    if SHADOW_STRICT:
        raise ShadowMismatch(msg)
