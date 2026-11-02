"""品質基準 DB の照合(VBA QP_DB / Quality_DB / 補助)の試験

期待値は VBA のコードを手で追ったもの(根拠の行番号をコメントに)。実データは tests/fixtures/kensa
(品質基準・伝送用ファイルは全表、SIKA* は 108 ロット)。行の順序で結果が変わる所は、DB の写しで
行を並べ替えて確かめる。
"""
from __future__ import annotations

import os
import shutil
import sqlite3
import tempfile
import unittest
from pathlib import Path
from typing import Any

from tests import kensa_fixture
from ufkensa.kensa import quality as q
from ufkensa.kensa import quality_view as qv
from ufkensa.kensa import quirks, vbx

CRLF = "\r\n"


def setUpModule() -> None:
    kensa_fixture.use_offline()
    q.clear_caches()
    quirks.set_for_test(None)


def lot_with(lot_no: str, changes: dict[str, str]) -> q.LotData:
    """実データのロットの一部の列を差し替えた LotData。"""
    ld = q.read_lot(lot_no)
    rec = list(ld.records[0])
    for k, v in changes.items():
        rec[ld.fields.index(k)] = v
    return q.LotData(ld.fields, [rec])


class _DbCopy(unittest.TestCase):
    """DB の写しを使う試験(行の並べ替え・値の書き換え)。"""

    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="ufk_quality_"))
        kensa_fixture.copy_to(self.tmp)
        kensa_fixture.use_offline(self.tmp)
        q.clear_caches()

    def tearDown(self) -> None:
        kensa_fixture.use_offline()
        q.clear_caches()
        quirks.set_for_test(None)
        shutil.rmtree(self.tmp, ignore_errors=True)

    def sql(self, db: str, stmt: str, *args: Any) -> None:
        path = self.tmp / kensa_fixture.file_name(db)
        before = path.stat().st_mtime_ns
        with sqlite3.connect(path) as c:
            c.execute(stmt, args)
        c.close()
        if path.stat().st_mtime_ns == before:          # D11 のキャッシュが更新に気づくように
            os.utime(path, ns=(before + 1_000_000, before + 1_000_000))

    def move_to_end(self, db: str, table: str, col: str, value: str) -> None:
        """その行を取り出して末尾に入れ直す(取り込み順 = rowid が最後になる)。"""
        path = self.tmp / kensa_fixture.file_name(db)
        with sqlite3.connect(path) as c:
            rows = c.execute(f'SELECT * FROM "{table}" WHERE "{col}" = ?', (value,)).fetchall()
            c.execute(f'DELETE FROM "{table}" WHERE "{col}" = ?', (value,))
            ph = ",".join("?" * len(rows[0]))
            for r in rows:
                c.execute(f'INSERT INTO "{table}" VALUES ({ph})', r)
        c.close()


# ------------------------------------------------------------------
# 照合関数(QP L804-1173)
# ------------------------------------------------------------------
class FindInListTest(unittest.TestCase):
    def test_blank(self) -> None:
        self.assertTrue(q.find_in_list("", "52S"))                 # QP:813
        self.assertFalse(q.find_in_list("", "52S", "x"))           # BlankFlag ありは空でも調べる → "" ≠ 52S
        self.assertTrue(q.find_in_list("", "", "x"))

    def test_patterns(self) -> None:
        self.assertTrue(q.find_in_list("*52S", "F52S"))            # 後方一致 QP:885
        self.assertFalse(q.find_in_list("*52S", "52SX"))
        self.assertTrue(q.find_in_list("52S*", "52SX"))            # 前方一致 QP:893
        self.assertTrue(q.find_in_list("*2S*", "F52SX"))           # 部分一致 QP:878
        self.assertTrue(q.find_in_list("52S", "52S"))
        self.assertFalse(q.find_in_list("52s", "52S"))             # 大小を区別(Binary)
        self.assertTrue(q.find_in_list(" 52S , 5083 ", "5083"))    # Trim

    def test_not_equal_sets_true(self) -> None:
        # ≠ は違えば True にする(前の肯定条件が外れていても)。G4 §4.4 の例
        self.assertTrue(q.find_in_list("52S,≠F52S", "5083"))
        self.assertTrue(q.find_in_list("A,≠B", "C"))
        self.assertFalse(q.find_in_list("SUS304,≠SUS304", "SUS304"))
        self.assertFalse(q.find_in_list("≠*52S", "F52S"))          # QP:850 後方一致で除外
        self.assertTrue(q.find_in_list("≠*52S", "5083"))
        self.assertFalse(q.find_in_list("≠52S*", "52SX"))          # QP:858
        self.assertFalse(q.find_in_list("≠*2S*", "A52SB"))         # QP:842

    def test_order_later_wins(self) -> None:
        # 肯定条件は途中で抜けない(後勝ちの OR)。≠ の不一致は即 False
        self.assertTrue(q.find_in_list("5083,≠52S", "5083"))
        self.assertTrue(q.find_in_list("≠52S,5083", "5083"))
        self.assertFalse(q.find_in_list("5083,≠5083", "5083"))
        self.assertFalse(q.find_in_list("≠5083,5083", "5083"))
        # 実データ(平坦度一覧QP追加 管理番号84): "52S,*52S,≠(P52S)"
        self.assertTrue(q.find_in_list("52S,*52S,≠(P52S)", "F52S"))
        self.assertFalse(q.find_in_list("52S,*52S,≠(P52S)", "(P52S)"))

    def test_lone_asterisk_is_runtime_error(self) -> None:
        # "*" だけ: Left(MidItem, Len-1) が Left("", -1) → 実行時エラー 5
        with self.assertRaises(q.VbaRuntimeError) as cm:
            q.find_in_list("*", "x")
        self.assertEqual(cm.exception.number, 5)


class CheckConditionTest(unittest.TestCase):
    def test_db_side_narrowed(self) -> None:
        self.assertTrue(q.check_condition(["５２Ｓ"], 1, "52S"))     # QP:920 StrConv(vbNarrow)
        self.assertTrue(q.check_condition(["ニホンシャリョウ"], 1, "ﾆﾎﾝｼｬﾘｮｳ"))
        self.assertTrue(q.check_condition(["ﾆﾂﾎﾟﾝﾌﾙﾊｰﾌ(ｶ"], 1, "ﾆﾂﾎﾟﾝﾌﾙﾊ-ﾌ(ｶ"))   # ｰ → -(QP:922)
        self.assertFalse(q.check_condition(["52S"], 1, "５２Ｓ"))   # 探す側は変換しない

    def test_missing_column(self) -> None:
        with self.assertRaises(q.VbaRuntimeError):
            q.check_condition(["a"], 0, "a")                       # 列番号 0 → 添字エラー


class RangeCheckTest(unittest.TestCase):
    def test_thickness(self) -> None:
        self.assertTrue(q.check_thickness(["", "", ""], 1, 2, 3, 5))     # 3列とも空(QP:930)
        self.assertTrue(q.check_thickness(["4", "50", ""], 1, 2, 3, 4))  # 下限ちょうど
        self.assertTrue(q.check_thickness(["4", "50", ""], 1, 2, 3, 50))  # 上限ちょうど
        self.assertFalse(q.check_thickness(["4", "50", ""], 1, 2, 3, 50.01))
        self.assertFalse(q.check_thickness(["4", "", ""], 1, 2, 3, 5))   # 上限が空 → 0
        self.assertFalse(q.check_thickness(["", "", "x"], 1, 2, 3, 5))   # 値列だけ入っている → 0..0
        self.assertTrue(q.check_thickness(["", "", "x"], 1, 2, 3, 0))

    def test_single_boundary(self) -> None:
        # 下・上は Single(QP:937)、厚は Val(CStr(Single)) の Double。Single(0.6) > 0.6 なので一致しない
        self.assertFalse(q.check_thickness(["0.6", "0.6", "0.6"], 1, 2, 3, 0.6))
        self.assertTrue(q.check_thickness(["0.5", "0.5", "0.5"], 1, 2, 3, 0.5))
        self.assertFalse(q.check_thickness(["0", "12.7", ""], 1, 2, 3, 12.7))   # Single(12.7) < 12.7
        self.assertTrue(q.check_thickness(["6.1", "80", ""], 1, 2, 3, 6.1))     # Single(6.1) < 6.1

    def test_within_and_size(self) -> None:
        self.assertTrue(q.check_within(["", ""], 1, 2, 0))           # QP:966 空でも 0..0
        self.assertFalse(q.check_within(["", ""], 1, 2, 1))
        self.assertTrue(q.check_size([""], 1, "1"))                  # QP:975
        self.assertFalse(q.check_size(["2≦"], 1, "1"))               # Val("2≦")=2、1 >= 2 は偽
        self.assertTrue(q.check_size(["2"], 1, "2"))
        self.assertTrue(q.check_size(["2.5"], 1, "2"))               # Long は銀行丸め 2.5→2
        self.assertFalse(q.check_size(["3.5"], 1, "3.9"))            # 3.5→4


class FindIndicesTest(unittest.TestCase):
    SORI = ["板厚下", "板厚上", "H500", "H1000", "H1500", "H1530", "H2000", "H9999", "T500", "T1000", "T1500",
            "T1530", "T2000", "T9999", "備考", "管理番号", "T1999"]

    def test_intervals(self) -> None:
        f = self.SORI
        self.assertEqual(q.find_indices(f, "H", 500), (3, None))      # 最小列 0 < x <= 500
        self.assertEqual(q.find_indices(f, "H", 500.5), (4, None))    # (500, 1000]
        self.assertEqual(q.find_indices(f, "H", 1530), (6, None))     # 上端を含む
        self.assertEqual(q.find_indices(f, "H", 10000), (8, 9999))    # 最大超 → 最大列 + 9999
        self.assertEqual(q.find_indices(f, "H", 0), (8, 9999))        # 0 以下も最大列
        self.assertEqual(q.find_indices(f, "T", 1800), (17, None))    # 末尾の T1999 も並べ替えて使う
        self.assertEqual(q.find_indices(f, "T", 1999.5), (13, None))

    def test_no_column_is_error(self) -> None:
        with self.assertRaises(q.VbaRuntimeError):
            q.find_indices(["板厚下", "板厚上"], "H", 100)

    def test_long_rounding_B27(self) -> None:
        f = ["X", "H100.5", "H50", "H200"]
        # 再現: 入れ替えで 100.5 → 100(銀行丸め)。"H100" の列は無いので Temp の添字 1 がそのまま列番号に(QP:1003,1042)
        self.assertEqual(q.find_indices(f, "H", 80), (1, None))
        quirks.set_for_test({"B27": "fix"})
        try:
            self.assertEqual(q.find_indices(f, "H", 80), (2, None))
        finally:
            quirks.set_for_test(None)


class LoopArrayTest(unittest.TestCase):
    F = ["板厚下", "板厚上", "500", "1000", "2000"]
    ROWS = [["0", "10", "1", "2", "3"], ["5", "20", "7", "8", "9"]]

    def test_first_row_and_intervals(self) -> None:
        f, r = self.F, self.ROWS
        self.assertEqual(q.loop_array(r, f, 500, 6), 1.0)             # 厚 6 は 2 行とも範囲内 → 最初の行
        self.assertAlmostEqual(q.loop_array(r, f, 500.25, 6), 2.0)
        self.assertEqual(q.loop_array(r, f, 2000, 6), 3.0)
        self.assertEqual(q.loop_array(r, f, 2001, 6), 3.0)            # 範囲外は最後の数の列(QDB:1433)
        self.assertEqual(q.loop_array(r, f, 0, 6), 3.0)               # 0 以下も
        self.assertEqual(q.loop_array(r, f, 500, 15), 7.0)
        self.assertEqual(q.loop_array(r, f, 500, 25), 0.0)            # 厚の行が無い

    def test_column_order_not_sorted(self) -> None:
        # 列は並べ替えない(列順で区間)。1000, 500 の順だと 300 は最初の列 1000 になる
        f = ["板厚下", "板厚上", "1000", "500"]
        self.assertEqual(q.loop_array([["0", "9", "4", "5"]], f, 300, 1), 4.0)
        # 最大超は「最後の数の列」(500)
        self.assertEqual(q.loop_array([["0", "9", "4", "5"]], f, 3000, 1), 5.0)

    def test_min_column_matched_by_val(self) -> None:
        # 最小列の判定は IsNumeric を見ずに Val(列名) で探す(QDB:1412)。"1m" も Val=1
        f = ["1m", "板厚下", "板厚上", "1", "5"]
        self.assertEqual(q.loop_array([["9", "0", "9", "4", "5"]], f, 0.5, 1), 9.0)

    def test_blank_cell_is_type_error(self) -> None:
        with self.assertRaises(vbx.TypeMismatch):
            q.loop_array([["0", "9", "", "5", "6"]], self.F, 100, 1)

    def test_no_numeric_column(self) -> None:
        with self.assertRaises(q.VbaRuntimeError):
            q.loop_array([["0", "9", "1"]], ["板厚下", "板厚上", "H500"], 100, 1)


class SmallFunctionsTest(unittest.TestCase):
    def test_extract(self) -> None:
        self.assertEqual(q.extract_nums("6㎜以下,フラットネス4級", "以下"), "6")
        self.assertEqual(q.extract_nums("中2.5,耳2.5", "中"), "2.5")
        self.assertEqual(q.extract_nums("3以上,5以上", "以上"), "3, 5")
        self.assertEqual(q.extract_nums("1.2.3以上", "以上"), "1.2")
        self.assertEqual(q.extract_nums("不問", "以上"), "")
        self.assertEqual(q.extract_number("6㎜以下"), 6.0)

    def test_compare_strings(self) -> None:
        self.assertTrue(q.compare_strings("6㎜以下", 6, 6))
        self.assertFalse(q.compare_strings("6㎜以下", 6.5, 6))
        self.assertTrue(q.compare_strings("3以上以下", 3, 3))          # "以上" が先
        self.assertFalse(q.compare_strings("4級", 1, 1))

    def test_find_priority(self) -> None:
        self.assertEqual(q.find_priority("HOT PSW ST2 QF 2HI ST2"), "T")   # H:1458
        self.assertEqual(q.find_priority("HOT CM1 CM2"), "N")             # 右から
        self.assertEqual(q.find_priority("HOT FS4"), "H")
        self.assertEqual(q.find_priority("CM1X"), "")
        self.assertEqual(q.find_priority(""), "")

    def test_second_smallest_and_weight(self) -> None:
        self.assertEqual(q.second_smallest([0, 1.2, 1.2, 0.8]), 0.8)
        with self.assertRaises(q.NotAvailable):
            q.second_smallest([0, 0])
        self.assertEqual(q.weight(10, 1000, 3000), 80.4)
        self.assertEqual(q.weight(3, 1001, 1001), 8.056)                 # 8.05608804 を切り捨て
        self.assertEqual(q.weight(1, 1, 1), 0.0)

    def test_narrow(self) -> None:
        self.assertEqual(q.narrow("ガＡ１　ー・ヴ"), "ｶﾞA1 ｰ･ｳﾞ")
        self.assertEqual(q.narrow("ひらがな漢字"), "ひらがな漢字")

    def test_variant_compare_u1(self) -> None:
        V, N, S = q.VAR, q.NUM, q.STR
        self.assertFalse(q.veq("", 0, V, N))          # 表では型エラー 13。U1: エラーにせず「等しくない」
        self.assertFalse(q.veq("6㎜以下,フラットネス4級", 0, V, N))
        self.assertEqual(q.vcmp("abc", 300.0, V, N), 1)             # 数 < 文字列
        self.assertTrue(q.veq("0.0", 0, V, N))
        self.assertTrue(q.veq(None, 0, V, N))                       # Empty = 0
        self.assertTrue(q.veq("13", 13, V, N))                      # 数になる Variant は数で
        self.assertFalse(q.veq("13", 13, V, V))                     # Variant どうし: 数 < 文字列
        self.assertTrue(q.veq(12345.0, "12345", V, S))              # String と Variant は文字で(CStr)
        self.assertFalse(q.veq(0, "", V, S))                        # Kousa(1,4)=0(SkipEndA)と "" → "0" ≠ ""
        q.STRING_VS_NUMBER_RAISES = True
        try:
            with self.assertRaises(vbx.TypeMismatch):
                q.veq("", 0, V, N)
        finally:
            q.STRING_VS_NUMBER_RAISES = False

    def test_u1_strict_table_breaks_main_paths(self) -> None:
        # 表どおり(型エラー 13)にすると、毎日通る経路が止まる: DBQP の `Kousa(1,4) = 0`(QP:462)、
        # 外観_H の見出し行 "管理番号" = Long(QDB:331)。U1 で「エラーにしない」を既定にした根拠
        q.STRING_VS_NUMBER_RAISES = True
        try:
            with self.assertRaises(vbx.TypeMismatch):
                q.dbqp("H6083T1", "H", "ｶ) ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ", "x", q.read_lot("H6083T1"))
            with self.assertRaises(vbx.TypeMismatch):
                qv.display_appearance_h(q.read_lot("H8401Z0"), qv.QualityForm())
        finally:
            q.STRING_VS_NUMBER_RAISES = False

    @unittest.skip("U1 実機確認待ち")
    def test_u1_string_vs_number(self) -> None:
        # VBA で `Kousa(1,4) = 0`(中耳列が "")がエラー 13 にならないこと。エラーなら STRING_VS_NUMBER_RAISES = True
        raise NotImplementedError


# ------------------------------------------------------------------
# DBQP(QP L69-801)
# ------------------------------------------------------------------
class DbqpTest(unittest.TestCase):
    def test_H6083T1_table_A3(self) -> None:
        # H205 F52S R、オーダー 85×1525×2500、等級4、検番 H。一致 [6, 105](行順)。優先(納入先などの記入)なし
        # 幅 1525 → (1254,1750] → H1750=2、丈 2500 → (2004,2504] → T2504=3(QP:384-385)
        # 中耳 "6㎜以下,フラットネス4級": 以下 6、4級 → 85 <= 6 は偽 → 中耳 ""(QP:494,503)
        r = q.dbqp("H6083T1", "H6X1922", "ｶ) ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ", "ﾆﾂﾃﾂﾌﾞﾂｻﾝ(ｶ)ｷﾉｳﾏﾃﾘｱﾙﾌﾞ", q.read_lot("H6083T1"))
        k = r.kousa
        self.assertEqual(len(k), 2)
        self.assertEqual(k[0], ["1.2", "2", "3", "", "",
                                "100mm超えにて規格外れが発生した場合、異常登録して保留し、品証課にて再判定を行う。",
                                "Q-S-QA-DIM-122_別表A③", "6", ""])
        self.assertEqual(k[1][q.K_NO], "105")
        self.assertEqual(r.form["Uatu"], "85.000")
        self.assertEqual(r.form["Uhaba"], "1529.0")
        self.assertEqual(r.form["Noun"], "ｶ) ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ")
        # Execution を渡すとフォームに書かない(QP:784)
        r2 = q.dbqp("H6083T1", "H6X1922", "ｶ) ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ", "x", q.read_lot("H6083T1"), execution="x")
        self.assertEqual(r2.form, {})

    def test_H9370Y0_priority_and_duplicate_comment(self) -> None:
        # K21S NP56 O 45×1220×2440。一致 [25, 98, 104]。25 は納入先(≠ﾕ)ｻﾝｴﾂ)が記入 → 優先(既に先頭)
        # 98 は「重複可」だがコメントが同文なので重ねない(QP:415)。中耳 "不問とする。" → 空
        r = q.dbqp("H9370Y0", "H6X1696", "ｶ)ﾋﾀﾁｾｲｻｸｼﾖ ｺｳﾂｳｼｽﾃﾑｼﾞｷﾞﾖｳﾌﾞ", "ｶﾅｲｷﾝｿﾞｸｺｳｷﾞﾖｳ(ｶ",
                   q.read_lot("H9370Y0"))
        self.assertEqual([x[q.K_NO] for x in r.kousa], ["25", "98", "104"])
        self.assertEqual(r.kousa[0][:6], ["", "2.5", "3.5", "", "", "異常発生時廃棄せずに品証又は生管に連絡する。"])

    def test_N6540J0_flatness_grade_sori(self) -> None:
        # K21B 5005 H14 3×1220×2440 等級4。管理番号102(≠NP56,≠U74S)。幅丈が「フラットネス品質等級」→ ソリ4級
        # H: 1220 → (1000,1500] → H1500、厚 3 は 2.01–7 の行 → 2。T: 2440 → (2000,9999] → T9999 → 3(QP:599-629)
        r = q.dbqp("N6540J0", "K690494", "ｶﾅｲｷﾝｿﾞｸｺｳｷﾞﾖｳ(ｶ", "ｶﾅｲｷﾝｿﾞｸｺｳｷﾞﾖｳ(ｶ", q.read_lot("N6540J0"))
        k = r.kousa[0]
        self.assertEqual(k[:5], ["", "2.0", "3.0", "", ""])
        self.assertEqual(k[q.K_SOURCE], "品質等級 Q-E-001_品質規格:等級4")
        self.assertEqual(k[q.K_NO], "102")

    def test_N9480A0_direct_naka_mimi(self) -> None:
        # K21S NP56 O 3×1220×2440。一致 [24, 98]。中耳 "中2.5,耳2.5" は以上/以下なし → 直接(QP:521-530)
        r = q.dbqp("N9480A0", "H", "ｶﾅｲｷﾝｿﾞｸｺｳｷﾞﾖｳ(ｶ", "ｶﾅｲｷﾝｿﾞｸｺｳｷﾞﾖｳ(ｶ", q.read_lot("N9480A0"))
        k = r.kousa[0]
        self.assertEqual(k[:5], ["", "2.5", "3.5", 2.5, 2.5])
        self.assertEqual(q.form_reflect_dist(r.kousa, quality_visible=False, mitsubishi_visible=False)
                         .values["Naka"], "2.5")

    def test_skip_qp_grade_tables(self) -> None:
        # L8281G0: E133 PF95 H14、0.3×984×0、等級3。QP に該当なし → SkipQP。ソリ3級 H1000=4、丈 0 → 最大列 T9999=5
        # 6 >= 0.3 で指示なし → 中歪3級/耳歪3級 を丈 0(0 以下 → 最大列 4000)で 4(QP:704-718)
        r = q.dbqp("L8281G0", "N6X1410", "ﾄｳﾖｳｱﾙﾐ ｶﾝﾊﾞﾗｾｲｿﾞｳｼﾖ", "x", q.read_lot("L8281G0"))
        self.assertEqual(r.kousa, [[None, "4.0", "5.0", "4.0", "4.0", None, " Q-E-001_品質規格:等級3", None, None]])

    def test_skip_qp_sika_instruction(self) -> None:
        # L7240E1: 指示 中歪 2 / 耳歪 3 があるので中耳の表は引かず、最後に指示値で補う(QP:705, 778-779)
        r = q.dbqp("L7240E1", "N695003", "ｱﾒﾘｶｶﾞﾂｼﾕｳｺｸ", "x", q.read_lot("L7240E1"))
        self.assertEqual(r.kousa[0][:5], [None, "2.0", "3.0", "2", "3"])

    def test_o_and_hx8(self) -> None:
        r = q.dbqp("L9280A0", "N6X1763", "x", "x", q.read_lot("L9280A0"), uf1_visible=True)
        self.assertEqual(r.kousa[0][5:7], ["別表Cに記載のないO材_判定不要", "Q-S-QA-DIM-122_別表C"])
        self.assertEqual(r.uf1, {"HizCome.Caption": "別表Cに記載のないO材_判定不要",
                                 "DistortionFrame.Caption": "参照元:Q-S-QA-DIM-122_別表C"})
        self.assertEqual(r.form, {})                               # skipEnd を通らない
        r = q.dbqp("N7312H0", "K6X1475", "x", "x", q.read_lot("N7312H0"), uf1_visible=True)
        self.assertEqual(r.kousa[0][5], "別表Dに記載のないHx8材_判定不要")
        self.assertEqual(r.uf1, {"HizCome.Caption": "別表Dに記載のないHx8材_判定不要"})   # 参照元は書かない
        r = q.dbqp("N7312H0", "K6X1475", "x", "x", q.read_lot("N7312H0"), uf1_visible=False)
        self.assertEqual(r.uf1, {})

    def test_B24_okuri_code_and_haba_split(self) -> None:
        # R6550H0: K434 75S T6 6.6×215×850 横割数 1。管理番号89 は 巾取数 "2≦" → CheckSize 偽だが最終判定に入らない
        lot = q.read_lot("R6550H0")
        r = q.dbqp("R6550H0", "H", "ｶ)ｼﾏﾉ", "x", lot)
        self.assertEqual([x[q.K_NO] for x in r.kousa], ["89", "104"])
        self.assertEqual(r.kousa[0][1:3], ["1", "1"])
        quirks.set_for_test({"B24": "fix"})
        try:
            r = q.dbqp("R6550H0", "H", "ｶ)ｼﾏﾉ", "x", lot)
            self.assertEqual([x[q.K_NO] for x in r.kousa], ["104"])
            self.assertEqual(r.kousa[0][1:3], ["2", "3"])          # 104: H500=2、T1004=3
        finally:
            quirks.set_for_test(None)

    def test_messages_and_errors(self) -> None:
        r = q.dbqp("X", "H", "", "", None)
        self.assertIsNone(r.kousa)
        self.assertEqual(r.messages[0].text, "日稲:SIKALOTNOW.accdbがありません_ReadCheck_DBQP")
        ld = q.read_lot("無いロット")
        r = q.dbqp("無いロット", "H", "", "", ld)
        self.assertEqual(r.messages[0].text, "SIKALOTNOWに該当ロットがありません_DBQP")
        # 検番が空、コースも空(H6083T1 の実績_設備ｺｰｽ は空)
        r = q.dbqp("H6083T1", "", "x", "x", q.read_lot("H6083T1"))
        self.assertEqual(r.messages[0].text, "検査Noがないファイルです_DBQP")
        # 検番 "0…" はコースから(HOT → H)
        r = q.dbqp("H9370Y0", "0X", "ｶ)ﾋﾀﾁｾｲｻｸｼﾖ ｺｳﾂｳｼｽﾃﾑｼﾞｷﾞﾖｳﾌﾞ", "x", q.read_lot("H9370Y0"))
        self.assertEqual(r.kousa[0][q.K_NO], "25")
        # 等級 "." は CLng で型エラー(L8131C0)
        with self.assertRaises(vbx.TypeMismatch):
            q.dbqp("L8131C0", "N6X1883", "x", "x", q.read_lot("L8131C0"))
        # 等級が 3/4 以外(QP 一致あり)→ MsgBox で Empty(QP:449-452)
        r = q.dbqp("H6083T1", "H", "ｶ) ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ", "x", lot_with("H6083T1", {"品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ": "5"}))
        self.assertIsNone(r.kousa)
        self.assertEqual(r.messages[0].text, "歪等級が3,4以外です")


class DbqpRowOrderTest(_DbCopy):
    def test_priority_row_later_in_table(self) -> None:
        # 管理番号25 を表の末尾へ → 一致順 [98, 104, 25]。優先行 25 が NewArray(1) を上書きし、98 は消え 25 が重複(QP:378)
        self.move_to_end("品質基準", "平坦度一覧QP追加", "管理番号", "25")
        r = q.dbqp("H9370Y0", "H6X1696", "ｶ)ﾋﾀﾁｾｲｻｸｼﾖ ｺｳﾂｳｼｽﾃﾑｼﾞｷﾞﾖｳﾌﾞ", "x", q.read_lot("H9370Y0"))
        self.assertEqual([x[q.K_NO] for x in r.kousa], ["25", "104", "25"])

    def test_duplicate_comment_appended(self) -> None:
        self.sql("品質基準", 'UPDATE "平坦度一覧QP追加" SET "表示コメント"=? WHERE "管理番号"=?', "追加注意", "98")
        r = q.dbqp("H9370Y0", "H6X1696", "ｶ)ﾋﾀﾁｾｲｻｸｼﾖ ｺｳﾂｳｼｽﾃﾑｼﾞｷﾞﾖｳﾌﾞ", "x", q.read_lot("H9370Y0"))
        self.assertEqual(r.kousa[0][5], "異常発生時廃棄せずに品証又は生管に連絡する。" + CRLF + "追加注意")

    def test_B26_grade_only(self) -> None:
        # 中耳が「4級」だけの行: 再現は等級表を引かず中耳 ""(QP:521-530)。直すと 中歪4級/耳歪4級(30.01–120、丈2500 → 4000列 1.5)
        self.sql("品質基準", 'UPDATE "平坦度一覧QP追加" SET "中歪・耳歪"=? WHERE "管理番号"=?', "4級", "6")
        lot = q.read_lot("H6083T1")
        r = q.dbqp("H6083T1", "H", "ｶ) ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ", "x", lot)
        self.assertEqual(r.kousa[0][3:5], ["", ""])
        quirks.set_for_test({"B26": "fix"})
        r = q.dbqp("H6083T1", "H", "ｶ) ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ", "x", lot)
        self.assertEqual(r.kousa[0][3:5], ["1.5", "1.5"])
        self.assertEqual(r.kousa[0][6], "Q-S-QA-DIM-122_別表A③ Q-E-001_品質規格:等級4")

    def test_nakimono_and_ijou_grade(self) -> None:
        self.sql("品質基準", 'UPDATE "平坦度一覧QP追加" SET "中歪・耳歪"=? WHERE "管理番号"=?', "無き事", "6")
        r = q.dbqp("H6083T1", "H", "ｶ) ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ", "x", q.read_lot("H6083T1"))
        self.assertEqual(r.kousa[0][3:5], [0, 0])                  # SkipEndA(QP:765)
        v = q.form_reflect_dist(r.kousa, quality_visible=False, mitsubishi_visible=False).values
        self.assertEqual((v["Naka"], v["Mimi"]), ("", ""))         # "0.0" = 0 → 空
        # 「以上」+3級: 85 >= 3 → 中歪3級/耳歪3級(30.01–120 → 3)
        self.sql("品質基準", 'UPDATE "平坦度一覧QP追加" SET "中歪・耳歪"=? WHERE "管理番号"=?', "3㎜以上,フラットネス3級", "6")
        r = q.dbqp("H6083T1", "H", "ｶ) ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ", "x", q.read_lot("H6083T1"))
        self.assertEqual(r.kousa[0][3:5], ["3.0", "3.0"])


class DistArrayTest(unittest.TestCase):
    def test_different_deliveries_same_tolerance(self) -> None:
        # H6083T1 は納入先名称が2種 → 引当ごとに DBQP。幅丈は同じ "2_3" なので上書きしない。フォームは最後の DBQP の値
        hiki = q.hiki_arr_create("H6083T1", "SIKAHIKINOW", "SIKAHIKINOW")
        self.assertEqual(hiki, [("ｶ) ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ", "ﾆﾂﾃﾂﾌﾞﾂｻﾝ(ｶ)ｷﾉｳﾏﾃﾘｱﾙﾌﾞ"),
                                ("ｶ)ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ ｶﾝｻｲｺｳｼﾞﾖｳ", "ﾆﾂﾃﾂﾌﾞﾂｻﾝ(ｶ)ｷﾉｳﾏﾃﾘｱﾙﾌﾞ")])
        r = q.dist_array("H6083T1", "H6X1922", hiki, q.read_lot("H6083T1"))
        self.assertEqual(r.kousa[0][1:3], ["2", "3"])
        self.assertNotIn("納入先による公差違い", r.kousa[0][5])
        self.assertEqual(r.form["Noun"], "ｶ)ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ ｶﾝｻｲｺｳｼﾞﾖｳ")

    def test_strictest_override(self) -> None:
        # 1件目 ﾊｸﾄﾞｳ → 管理番号14(H1750=0.6、T3054=0.7)。2件目は白銅以外 → 管理番号104(2/3)
        # Val の最小で1件目を上書き(Double)。ゲージ・中・耳は1件目のまま(QP:1953-1974)
        hiki = [("ﾊｸﾄﾞｳ(ｶ)ｶﾅｶﾞﾜｼﾞｷﾞﾖｳｼﾖ", "ﾊｸﾄﾞｳ(ｶ"), ("ﾍﾞﾂﾉ(ｶ", "ﾍﾞﾂﾉ(ｶ")]
        r = q.dist_array("H9437E0", "H6X1526", hiki, q.read_lot("H9437E0"), execution="x")
        k = r.kousa[0]
        self.assertEqual(k[q.K_NO], "14")
        self.assertEqual((k[1], k[2]), (0.6, 0.7))
        self.assertTrue(k[5].endswith(CRLF + "納入先による公差違い有,厳しい公差で上書き"))
        self.assertEqual(q.form_reflect_dist(r.kousa, quality_visible=False, mitsubishi_visible=False)
                         .values["HaHiz"], "0.6")

    def test_empty_hiki(self) -> None:
        self.assertIsNone(q.dist_array("X", "H", None, None).kousa)
        self.assertIsNone(q.hiki_arr_create("", "SIKAHIKINOW", "SIKAHIKINOW"))
        self.assertIsNone(q.hiki_arr_create("無いロット", "SIKAHIKINOW", "SIKAHIKINOW"))


class FormReflectTest(unittest.TestCase):
    K = [["1.2", "2", "フラットネス品質等級", "0", None, "注意。です_", "参照", "6", ""]]

    def test_uf1(self) -> None:
        v = q.form_reflect_dist(self.K, quality_visible=False, mitsubishi_visible=False).values
        self.assertEqual(v["MG"], "1.2")
        self.assertEqual(v["HaHiz"], "2.0")
        self.assertEqual(v["TaHiz"], "フラットネス品質等級")        # 数でなければ文字のまま(QP:1998)
        self.assertEqual(v["Naka"], "")                            # "0.0" = 0 → 空
        self.assertEqual(v["Mimi"], "")                            # Empty は IsNumeric 真 → Format は ""
        self.assertEqual(v["HizCome.Caption"], "注意" + CRLF + "です")  # 末尾 "_" を除き "。" → 改行
        self.assertEqual(v["DistortionFrame.Caption"], "参照元:参照")
        v = q.form_reflect_dist([["0.5以下", "", "", "", "", "", "", "", ""]],
                                quality_visible=False, mitsubishi_visible=False).values
        self.assertEqual(v["MG"], "")                              # 数でない MG は空

    def test_quality_form_and_sheet(self) -> None:
        r = q.form_reflect_dist(self.K, quality_visible=True, mitsubishi_visible=False)
        self.assertEqual(r.target, "UF品質基準")
        self.assertEqual(r.values["TaHiz1"], "フラットネス品質等級")
        self.assertEqual(r.values["Naka1"], "0.0")                 # こちらは 0 を消さない
        self.assertEqual(r.values["Sansyou1"], "参照_注意。です")    # 末尾 "_" を1つ除く
        r = q.form_reflect_dist(self.K, quality_visible=False, mitsubishi_visible=True, uf1_visible=False,
                                zenkasho_sheet=True)
        self.assertEqual((r.target, r.cells), ("シート", {(5, 13): "2", (5, 15): "フラットネス品質等級"}))
        form = qv.QualityForm()
        form.apply(q.form_reflect_dist(self.K, quality_visible=True, mitsubishi_visible=False).values)
        self.assertEqual(form.MG1, "1.2")


# ------------------------------------------------------------------
# DbReturnItem・ラテラル(QP L1176-1456, QDB L1491-1696)
# ------------------------------------------------------------------
class ReturnItemTest(unittest.TestCase):
    def test_lateral_table_B(self) -> None:
        # H9022K1(ﾆﾎﾝｼﾔﾘﾖｳｾｲｿﾞｳ K21S NP56 O 製造 23×672×2703): ラテフリ直角 管理番号8 → ラテラルボー要、公差 "別表B"
        # → DIM_179_日車: 材質 NP56、23 は 6.51–245、幅 672 は 450.1–9999 → 丈 2703 は (2300,3050] → 3
        lot = q.read_lot("H9022K1")
        r = q.db_return_item("ﾆﾎﾝｼﾔﾘﾖｳｾｲｿﾞｳ(ｶ", "ﾐｽﾞﾉｷﾝｿﾞｸｼﾖｳｼﾞ(ｶ", "K21S", "NP56", "O", 23, 672, 2703,
                             lateral_lot=lot)
        self.assertEqual(r.items, ",ラテラルボー")
        self.assertEqual(r.uf1, {"HURIS": "", "Sumisei.Caption": "ﾆﾎﾝｼﾔﾘﾖｳｾｲｿﾞｳ(ｶ", "LateralText": "3.0",
                                 "LateralText.BackColor": q.rgb(255, 153, 0)})
        self.assertEqual(q.dbdisplay_lateral(lot, "日車"), 3.0)
        # 追記公差があれば上書きしない
        r = q.db_return_item("ﾆﾎﾝｼﾔﾘﾖｳｾｲｿﾞｳ(ｶ", "x", "K21S", "NP56", "O", 23, 672, 2703, lateral_lot=lot,
                             additional_caption="ﾗﾃﾗﾙﾎﾞ-追記公差あり")
        self.assertNotIn("LateralText", r.uf1)

    def test_lateral_table_A_and_mitsubishi(self) -> None:
        # ﾊｸﾄﾞｳ: 管理番号5 → 別表A → DIM_179_白銅滑川(一致 [4, 8] の最初の行 4 → 丈 3053 は (2500,9999] → 3)
        lot = q.read_lot("H9437E0")
        args = ("ﾊｸﾄﾞｳ(ｶ)ｶﾅｶﾞﾜｼﾞｷﾞﾖｳｼﾖ", "ﾊｸﾄﾞｳ(ｶ", "H279", "F52S", "R", 36, 1528, 3053)
        r = q.db_return_item(*args, lateral_lot=lot)
        self.assertEqual(r.uf1["LateralText"], "3.0")
        r = q.db_return_item(*args, lateral_lot=lot, mitsubishi_visible=True, work_flag="P_H494")
        self.assertEqual((r.sheet, r.cells), ("H494検査", {(1, 15): "ﾊｸﾄﾞｳ(ｶ)ｶﾅｶﾞﾜｼﾞｷﾞﾖｳｼﾖF52S", (1, 13): "3.0"}))
        self.assertNotIn("LateralText", r.uf1)

    def test_first_hit_tolerance_used_for_lateral(self) -> None:
        # ｽﾐﾄﾓｾｲﾐﾂ R115: 一致 [9(直角度・公差1), 10(ラテラル・等級)]。ラテラルの公差に1行目(直角度の行)の "1" を使う
        r = q.db_return_item("ｽﾐﾄﾓｾｲﾐﾂｺｳｷﾞﾖｳ(ｶ", "x", "R115", "3S", "O", 10, 1000, 2000)
        self.assertEqual(r.items, ",直角度,ラテラルボー")
        self.assertEqual(r.uf1["LateralText"], "1.0")

    def test_no_match_messages(self) -> None:
        # 該当なし: 最後に評価した行(管理番号11 ﾀﾞｲｹﾝ)で判定。納入先で外れ → C4 偽 → 「_製造材質」(B27)
        r = q.db_return_item("ﾆﾎﾝｼﾔﾘﾖｳｾｲｿﾞｳ(ｶ", "x", "K21S", "NP56", "O", 300, 672, 2703)
        self.assertEqual(r.items, "")
        self.assertEqual(r.uf1["Additional.Caption"], "ﾗﾃﾗﾙﾎﾞｰ検査不要_製造材質")
        self.assertEqual(r.uf1["LateralText"], "")
        r = q.db_return_item("K.T.N. CO. LTD", "ﾅﾒｶﾜｱﾙﾐ(ｶ", "H283", "F52S", "R", 10, 1528, 3053)
        self.assertEqual(r.uf1["Additional.Caption"], "K.T.N ﾗﾃﾗﾙﾎﾞｰ検査不要")

    def test_single_boundary_free_sori(self) -> None:
        # 管理番号1(厚 0.6–0.6)は Single(0.6) > 0.6 のため 0.6 のロットに一致しない(新しく見つけた不具合)
        r = q.db_return_item("ﾆﾂﾎﾟﾝﾌﾙﾊ-ﾌ(ｶ", "x", "F975", "52S", "H18", 0.6, 1217, 2438)
        self.assertEqual(r.items, "")

    def test_lateral_none_lot(self) -> None:
        msgs: list = []
        self.assertEqual(q.dbdisplay_lateral(q.read_lot("無いロット"), "日車", messages=msgs), 0.0)
        self.assertEqual(msgs[0].text, "SIKALOTNOWに該当ロットがありません_DBdisplayﾗﾃﾗﾙ")
        msgs = []
        q.dbdisplay_lateral(q.read_lot("無いロット"), "日車", quiet=True, messages=msgs)
        self.assertEqual(msgs, [])


class ReturnItemDbTest(_DbCopy):
    def test_free_sori(self) -> None:
        self.sql("品質基準", 'UPDATE "ラテフリ直角" SET "厚下"=?, "厚上"=? WHERE "管理番号"=?', "0.5", "0.7", "1")
        r = q.db_return_item("ﾆﾂﾎﾟﾝﾌﾙﾊ-ﾌ(ｶ", "x", "F975", "52S", "H18", 0.6, 1217, 2438)
        self.assertEqual(r.items, ",フリー反り")
        self.assertEqual(r.uf1["HURIS"], "40")
        self.assertEqual(r.uf1["HizCome.Tag"], "40mm以下")
        self.assertEqual(r.uf1["HizCome.Caption"], "ﾌﾘｰ反り40mm以下")
        self.assertEqual(r.uf1["flatnesText"], "ﾌﾘｰｿﾘ要")
        self.assertEqual(r.uf1["LGF.Caption"], "ﾌﾘｰｿﾘ")
        self.assertAlmostEqual(r.cb_add_atu, 0.6, places=6)

    def test_B27_message_mapping(self) -> None:
        # 最後の行を「納入・取引・用途・材質は一致、調質で外れる」にする → 再現は「_ｻｲｽﾞ幅」、直すと「_ｻｲｽﾞ厚」
        self.sql("品質基準", 'UPDATE "ラテフリ直角" SET "納入先"=?, "用途コード"=NULL, "製造調質"=? WHERE "管理番号"=?',
                 "ﾆﾎﾝｼﾔﾘﾖｳｾｲｿﾞｳ(ｶ", "T4", "11")
        args = ("ﾆﾎﾝｼﾔﾘﾖｳｾｲｿﾞｳ(ｶ", "x", "K21S", "NP56", "O", 300, 672, 2703)
        self.assertEqual(q.db_return_item(*args).uf1["Additional.Caption"], "ﾗﾃﾗﾙﾎﾞｰ検査不要_ｻｲｽﾞ幅")
        quirks.set_for_test({"B27": "fix"})
        self.assertEqual(q.db_return_item(*args).uf1["Additional.Caption"], "ﾗﾃﾗﾙﾎﾞｰ検査不要_ｻｲｽﾞ厚")

    def test_lateral_empty_table(self) -> None:
        # 表が 0 件なら GetFieldsArr が先に Empty を返して抜ける(QDB:1600)。UBound のエラー(B27 m)には届かない
        self.sql("品質基準", 'DELETE FROM "Q_S_QA_DIM_179_日車"')
        self.assertEqual(q.dbdisplay_lateral(q.read_lot("H9022K1"), "日車"), 0.0)


# ------------------------------------------------------------------
# 設備仕様・時間用・ストレートエッヂ・バーコール
# ------------------------------------------------------------------
class TransferTest(unittest.TestCase):
    def test_equipment(self) -> None:
        self.assertEqual(q.check_equipment_db("LVC", "52S", 2, 1000, 2000, 30), "可")   # 管理番号2 に一致
        # 最後に評価した行(管理番号2)で厚が外れる。材質リストに無い 52S でも同じ文言(QP:2155)
        self.assertEqual(q.check_equipment_db("LVC", "A2ST", 0.5, 1000, 2000, 30), "オーダー厚 範囲外")
        self.assertEqual(q.check_equipment_db("LVC", "52S", 2, 500, 2000, 30), "C_オーダー幅 範囲外")
        self.assertEqual(q.check_equipment_db("SVC", "52S", 3, 1000, 3000, 31), "オーダー丈 範囲外")
        self.assertEqual(q.check_equipment_db("SVC", "52S", 3, 1000, 2000, 31), "単重 範囲外")
        self.assertEqual(q.check_equipment_db("ﾄｯﾄ", "52S", 3, 1000, 3000, 30), "可")    # D13: 幅だけ揃えて トット
        self.assertEqual(q.check_equipment_db("丸徳", "52S", 3, 1, 1, 1), "可")          # 設備名が無い

    def test_equipment_B27_fix(self) -> None:
        quirks.set_for_test({"B27": "fix"})
        try:
            self.assertEqual(q.check_equipment_db("LVC", "52S", 2, 500, 2000, 30), "オーダー幅 範囲外")
        finally:
            quirks.set_for_test(None)

    def test_shift_time_and_straight_edge(self) -> None:
        self.assertEqual(q.shift_time(), [("1", "07:00", "15:00"), ("2", "15:00", "22:50"),
                                          ("3", "22:50", "07:00"), ("昼", "08:15", "17:00")])
        self.assertEqual(q.straight_edge("ﾊﾞﾗﾝｻｰ"), ("2051.3", "2052.3"))   # マスタは全角 バランサー
        self.assertEqual(q.straight_edge("L-1"), ("2052.7", "2053.7"))
        self.assertEqual(q.straight_edge("丸徳"), ("", ""))
        self.assertEqual(q.straight_edge("無いライン"), "")

    def test_bh_search(self) -> None:
        self.assertEqual(q.bh_search("U74", "T4"), 78)
        self.assertEqual(q.bh_search("XU74", "T6"), 80)          # 4文字以上は右3文字(QDB:1479)
        self.assertEqual(q.bh_search("U74", "O"), 0)
        # マスタの 75Ｓ は全角。ADO Filter を全角半角を区別するとみなすと一致しない(U3)
        self.assertEqual(q.bh_search("75S", "T6"), 0)


class TransferDbTest(_DbCopy):
    def test_equipment_last_work_flag(self) -> None:
        # 一致した行の空でない [作業可否] の**最後**(QP:2187)
        self.sql("伝送用ファイル", 'UPDATE "設備仕様" SET "製造材質"=NULL, "作業可否"=? WHERE "管理番号"=?', "条件付", "1")
        self.sql("伝送用ファイル", 'UPDATE "設備仕様" SET "作業可否"=NULL WHERE "管理番号"=?', "2")
        self.assertEqual(q.check_equipment_db("LVC", "52S", 2, 1000, 2000, 30), "条件付")
        self.sql("伝送用ファイル", 'UPDATE "設備仕様" SET "作業可否"=? WHERE "管理番号"=?', "可", "2")
        self.assertEqual(q.check_equipment_db("LVC", "52S", 2, 1000, 2000, 30), "可")

    def test_straight_edge_first_and_last(self) -> None:
        self.sql("伝送用ファイル", 'INSERT INTO "校正ストレートエッヂ" ("番号","ライン名","長さ下","長さ上") VALUES (?,?,?,?)',
                 "8", "LVC", "1", "9")
        self.assertEqual(q.straight_edge("LVC"), ("2052.6", "9"))   # 先頭行の下・最終行の上

    def test_bh_last_match(self) -> None:
        self.sql("品質基準", 'INSERT INTO "Q_S_KE_ALL_053" VALUES (?,?,?,?)', "U74", "T4", "99", "0")
        self.assertEqual(q.bh_search("U74", "T4"), 99)

    def test_shift_time_short(self) -> None:
        self.sql("伝送用ファイル", 'DELETE FROM "時間用" WHERE "直"=?', "昼")
        with self.assertRaises(q.VbaRuntimeError):
            q.shift_time()


# ------------------------------------------------------------------
# 外注条件・ゲージ判定・梱包保護材(D11: ファイルの更新で読み直す)
# ------------------------------------------------------------------
HIST_HEAD = ["登録日", "ｵｰﾀﾞｰNo", "ﾛｯﾄNo", "厚", "巾", "丈", "検査No", "材質", "調質", "用途ｺｰﾄﾞ", "ｺｰｽ", "包装仕様NO"]


def hist(*rows: list) -> list[list]:
    return [HIST_HEAD] + [list(r) for r in rows]


class GaichuTest(unittest.TestCase):
    def setUp(self) -> None:
        q.clear_caches()

    def test_set1(self) -> None:
        h = hist([None, "A", "L1", 6.75, 215.0, 850.0, "H1", "75S", "T6", "K435", "HOT FS4 GFS", ""])
        self.assertEqual(q.gaichu_packing_no(h, "L1", 3), "1P0122")
        h[1][3] = 6.7504                                          # Round(,3) = 6.75
        self.assertEqual(q.gaichu_packing_no(h, "L1", 3), "1P0122")
        h[1][3] = 6.7506
        self.assertEqual(q.gaichu_packing_no(h, "L1", 3), "")
        h[1][3] = 6.75
        h[1][10] = "HOT FS4"                                      # ｺｰｽ条件(番号5)が成立しない
        self.assertEqual(q.gaichu_packing_no(h, "L1", 3), "")

    def test_set2_or_within_number(self) -> None:
        # 条件番号3(H19S/H197)と4(>=322.5 / >=100.0)は同じ番号内で OR → H19S で厚 150 でも成立
        h = hist([None, "A", "L2", 150.0, 1000.0, 2000.0, "H", "F52S", "R", "H19S", "HOT GFS", ""])
        self.assertEqual(q.gaichu_packing_no(h, "L2", 3), "1P0113")

    def test_any_row_of_lot(self) -> None:
        # ﾛｯﾄNo が同じ行のどれかで成立すればよい(行ごとに別の条件を満たしても成立)
        h = hist([None, "A", "L3", 6.75, 1, 1, "H", "75S", "T6", "K434", "", ""],
                 [None, "A", "L3", 1.0, 1, 1, "H", "", "", "", "GCT", ""])
        self.assertEqual(q.gaichu_packing_no(h, "L3", 3), "1P0122")
        self.assertEqual(q.gaichu_packing_no(h, "別", 3), "")


class GaugeTest(unittest.TestCase):
    def setUp(self) -> None:
        q.clear_caches()

    def test_rules(self) -> None:
        self.assertFalse(q.is_gauge_required(10, 1000, 3000, "52S", "H34", ""))     # 80.4kg
        self.assertTrue(q.is_gauge_required(5, 1000, 6000, "52S", "H34", ""))       # 丈_超 5000
        self.assertFalse(q.is_gauge_required(5, 1000, 5000, "52S", "H34", ""))      # 5000 は超えない
        self.assertTrue(q.is_gauge_required(1, 1000, 1000, "x", "F", "HOT SSW"))     # 調質 F かつ SSW
        # 重量 1300 超 かつ 丈 2500 未満(行1): 50×1500×2400×2.68/1e6 = 482.4 → 不成立、200 厚で 1929.6 → 成立
        self.assertFalse(q.is_gauge_required(50, 1500, 2400, "52S", "R", ""))
        self.assertTrue(q.is_gauge_required(200, 1500, 2400, "52S", "R", ""))


class MasterCacheTest(_DbCopy):
    def test_gauge_load_failure_and_reload(self) -> None:
        self.sql("品質基準", 'ALTER TABLE "ゲージ判定条件" RENAME COLUMN "有効" TO "有効X"')
        msgs: list = []
        self.assertFalse(q.is_gauge_required(5, 1000, 6000, "52S", "H34", "", msgs))
        self.assertEqual(msgs[0].text, "ゲージ判定条件を読み込めませんでした" + CRLF + "内容: -2147221502 "
                         "「ゲージ判定条件」に「有効」フィールドがありません" + CRLF + CRLF
                         + "定盤表示になります。原因を解消後 ReloadGauge を実行してください。")
        msgs = []
        self.assertFalse(q.is_gauge_required(5, 1000, 6000, "52S", "H34", "", msgs))
        self.assertEqual(msgs, [])                                # 同じファイルのうちは再試行しない
        self.sql("品質基準", 'ALTER TABLE "ゲージ判定条件" RENAME COLUMN "有効X" TO "有効"')
        self.assertTrue(q.is_gauge_required(5, 1000, 6000, "52S", "H34", "", msgs))   # 更新で読み直す(D11)

    def test_gaichu_failure_and_update(self) -> None:
        self.sql("品質基準", 'DROP TABLE "外注条件明細"')
        msgs: list = []
        h = hist([None, "A", "L1", 6.75, 1, 1, "H", "75S", "T6", "K435", "GFS", ""])
        self.assertEqual(q.gaichu_packing_no(h, "L1", 3, msgs), "")
        self.assertIn("段階: 外注条件明細 読込", msgs[0].text)

    def test_gaichu_reads_update(self) -> None:
        h = hist([None, "A", "L1", 6.75, 1, 1, "H", "75S", "T6", "K435", "GFS", ""])
        self.assertEqual(q.gaichu_packing_no(h, "L1", 3), "1P0122")
        self.sql("品質基準", 'UPDATE "外注条件" SET "戻り値"=? WHERE "管理番号"=?', "1P9999", "1")
        self.assertEqual(q.gaichu_packing_no(h, "L1", 3), "1P9999")


class MaterialsUpperPartTest(unittest.TestCase):
    COLS = ["管理番号", "包装仕様書", "用途コード", "材質", "調質", "板厚下", "板厚上", "板厚", "板幅下", "板幅上", "板幅",
            "板丈下", "板丈上", "板丈", "使用保護材", "備考"]

    def setUp(self) -> None:
        q.clear_caches()
        self.tmp = Path(tempfile.mkdtemp(prefix="ufk_pack_"))
        self.db = self.tmp / "梱包資材マスタ.sqlite3"
        rows = [
            ["1", "1P0104", None, "*52S", None, "0", "10", None, None, None, None, None, None, None, "ザラ板", None],
            ["2", "1P0104", None, "*52S", None, "0", "100", None, None, None, None, None, None, None, "アングル", None],
            ["3", "1P0104", None, None, None, None, None, None, None, None, None, None, None, None, "当板", "当板用"],
        ]
        with sqlite3.connect(self.db) as c:
            c.execute("CREATE TABLE 梱包保護材 (" + ",".join(f'"{x}" TEXT' for x in self.COLS) + ")")
            c.executemany(f"INSERT INTO 梱包保護材 VALUES ({','.join('?' * len(self.COLS))})", rows)
        c.close()

    def tearDown(self) -> None:
        q.clear_caches()
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_match(self) -> None:
        h = hist([None, "12345", "L1", 5.0, 1000.0, 2000.0, "H", "F52S", "H34", "H197", "", "1P0104"],
                 [None, "999", "L2", 50.0, 1000.0, 2000.0, "H", "F52S", "H34", "H197", "", "1P0104"],
                 [None, "888", "L3", 5.0, 1000.0, 2000.0, "H", "2ST", "H34", "H197", "", "1P0104"],
                 [None, "777", "L4", 5.0, 1000.0, 2000.0, "H", "F52S", "H34", "H197", "", "1P0200"])
        self.assertEqual(q.materials_upper_part("'12345", h, db=self.db), "ザラ板")    # 最初の一致
        self.assertEqual(q.materials_upper_part("12345", h, True, db=self.db), "当板")
        self.assertEqual(q.materials_upper_part("999", h, db=self.db), "アングル")
        self.assertEqual(q.materials_upper_part("888", h, db=self.db), "一致なし")
        self.assertEqual(q.materials_upper_part("777", h, db=self.db), "")            # 対象外の包装仕様書
        self.assertEqual(q.materials_upper_part("000", h, db=self.db), "")
        self.assertEqual(q.materials_upper_part("12345", None, db=self.db), "")       # シートが無い
        # D11: ファイルを直したら次の呼び出しで読み直す(VBA はオーダーNo が同じ間は古いまま)
        before = self.db.stat().st_mtime_ns
        with sqlite3.connect(self.db) as c:
            c.execute('UPDATE 梱包保護材 SET "使用保護材"=? WHERE "管理番号"=?', ("ザラ板2", "1"))
        c.close()
        os.utime(self.db, ns=(before + 1_000_000, before + 1_000_000))
        self.assertEqual(q.materials_upper_part("12345", h, db=self.db), "ザラ板2")

    def test_target_list(self) -> None:
        self.assertTrue(q.is_target_hoso_no("1P0120"))
        self.assertFalse(q.is_target_hoso_no("1P0121"))


# ------------------------------------------------------------------
# UF品質基準 の表示(QDB)
# ------------------------------------------------------------------
class AppearanceHTest(unittest.TestCase):
    def test_H8401Z0(self) -> None:
        # H283 F52S R 製造 15×1528×3053。別表2 を '*H283*' で絞ると 14-19 と 20-25 の組。15 は 15 と 21 に一致
        # → 後の 21 が残る(QDB:328-342)。21 のグレード 4,4,4,3,3,3
        r = qv.display_appearance_h(q.read_lot("H8401Z0"), qv.QualityForm())
        f = r.form
        self.assertEqual((f.ATU, f.HABA, f.TAKE, f.Zai, f.TYO, f.YoutC), ("15.000", "1528.0", "3053.0", "F52S", "R",
                                                                          "H283"))
        names = ["飛込み", "キズ/スリキズ/スリ跡", "鋳造剥離AL圧込", "AL付着跡", "汚れボタ落ち", "カミコミ跡ショック跡"]
        self.assertEqual(f.tabs, names)
        self.assertEqual([g[2] for g in r.grade_array], ["4", "4", "4", "3", "3", "3"])
        self.assertEqual(f.AbnoName, "「異常グレード確認一覧」" + CRLF + "".join(n + CRLF for n in names))
        self.assertEqual((f.AbnoLabel, f.gradeTB), ("飛込み", "4"))
        self.assertTrue(f.AbnoTB.startswith("触感:あってもOK\n色:―"))
        self.assertIn("1mmφ未満", f.AbnoTB)                      # 別表1 飛込み のグレード4
        self.assertEqual(f.KousinL1, "G279,H179追加_2024.7.15 ")  # 別表2 全件の先頭行の更新日
        # タブ表示(TabAb_Change)
        qv.uf_db_display(f, r.grade_array, "AL付着跡")
        self.assertEqual(f.gradeTB, "3")

    def test_thin_and_missing(self) -> None:
        r = qv.display_appearance_h(q.read_lot("L726M50"), qv.QualityForm())
        self.assertIsNone(r.grade_array)                         # オーダー厚 2 < 8
        self.assertEqual(r.form.ATU, "1.980")
        r = qv.display_appearance_h(q.read_lot("無いロット"), qv.QualityForm())
        self.assertEqual(r.messages[0].text, "SIKALOTNOW.accdbに該当ロットがありません_DBdisplay外観")
        r = qv.display_appearance_h(q.read_lot("無いロット"), qv.QualityForm(), quiet=True)
        self.assertEqual(r.messages, [])

    def test_tabs_accumulate(self) -> None:
        f = qv.QualityForm(tabs=["前回"])
        r = qv.display_appearance_h(q.read_lot("H8401Z0"), f)
        self.assertEqual(r.form.tabs[0], "前回")                  # 消さずに追加(H:1569)
        self.assertEqual(len(r.form.tabs), 7)
        self.assertEqual(f.tabs, ["前回"])                        # 渡した form は変えない


class AppearanceHDbTest(_DbCopy):
    def test_last_match_wins(self) -> None:
        self.sql("品質基準", 'UPDATE "Q_S_QA_124_別表2" SET "飛込み"=? WHERE "管理番号"=?', "1", "15")
        r = qv.display_appearance_h(q.read_lot("H8401Z0"), qv.QualityForm())
        self.assertEqual(r.grade_array[0][2], "4")                # 21 が後なので 15 の変更は効かない
        self.move_to_end("品質基準", "Q_S_QA_124_別表2", "管理番号", "15")
        r = qv.display_appearance_h(q.read_lot("H8401Z0"), qv.QualityForm())
        self.assertEqual(r.grade_array[0][2], "1")                # 15 が最後 → 15 の値

    def test_B25(self) -> None:
        # いまの別表2 には調質の列が無い(直しても条件なし)。列を足して確かめる
        self.sql("品質基準", 'ALTER TABLE "Q_S_QA_124_別表2" ADD COLUMN "調質" TEXT')
        self.sql("品質基準", 'UPDATE "Q_S_QA_124_別表2" SET "調質"=? WHERE "管理番号"=?', "H34", "21")
        self.sql("品質基準", 'UPDATE "Q_S_QA_124_別表2" SET "飛込み"=? WHERE "管理番号"=?', "1", "15")
        r = qv.display_appearance_h(q.read_lot("H8401Z0"), qv.QualityForm())
        self.assertEqual(r.grade_array[0][2], "4")                # 調質を見ない → 21 も一致
        quirks.set_for_test({"B25": "fix"})
        r = qv.display_appearance_h(q.read_lot("H8401Z0"), qv.QualityForm())
        self.assertEqual(r.grade_array[0][2], "1")                # R ≠ H34 で 21 が外れ 15 だけ


class AppearanceSTest(unittest.TestCase):
    def test_pattern_and_buttons(self) -> None:
        # L726M50: T245 NK01 H14 グレード B、オーダー厚 2。別表1 管理番号13(T238,T245、厚フラグなし)→ パターン1 = 4
        r = qv.display_appearance_s(q.read_lot("L726M50"), "トビコミ", qv.QualityForm())
        f = r.form
        self.assertEqual(f.Upater, "4")
        self.assertEqual((f.SelectO, f.SelectS, f.SelectK, f.SelectA), ("2㎜以下", "爪に掛らない", "ﾋﾟｯﾁ300㎜以上", "/"))
        self.assertEqual(f.Sansyou1, " Q-S-QA-SRF-128")
        self.assertEqual(f.KousinL4, "")
        self.assertFalse(f.L2_visible)
        r2 = qv.display_appearance_s(q.read_lot("L726M50"), "無いボタン", f)
        self.assertEqual((r2.form.SelectO, r2.form.Sansyou1), ("", " Q-S-QA-SRF-128"))   # 追記しない

    def test_thickness_flag(self) -> None:
        # L9150S0: U171 52S グレード B。管理番号3(厚フラグあり、板厚1 2.5)。1.0 → パターン1、2.5 ちょうども 1、2.51 → 2
        self.assertEqual(qv.display_appearance_s(q.read_lot("L9150S0"), "キズ,スリキズ", qv.QualityForm()).form.Upater,
                         "1")
        self.assertEqual(qv.display_appearance_s(lot_with("L9150S0", {"ｵｰﾀﾞｰ板厚": "2.5"}), "x",
                                                 qv.QualityForm()).form.Upater, "1")
        self.assertEqual(qv.display_appearance_s(lot_with("L9150S0", {"ｵｰﾀﾞｰ板厚": "2.51"}), "x",
                                                 qv.QualityForm()).form.Upater, "2")

    def test_no_pattern_and_thick(self) -> None:
        f = qv.display_appearance_s(q.read_lot("L9280A0"), "x", qv.QualityForm()).form
        self.assertEqual(f.Upater, "該当ﾊﾟﾀｰﾝなし")                 # R292 の行は材質 995 だけ
        self.assertEqual(qv.display_appearance_s(q.read_lot("H8401Z0"), "x", qv.QualityForm()).form.Upater,
                         "該当板厚なし")
        # R6550H0: オーダー厚 6.6(< 8)だが用途 K434 が別表1 に無い
        self.assertEqual(qv.display_appearance_s(q.read_lot("R6550H0"), "x", qv.QualityForm()).form.Upater,
                         "該当用途ｺｰﾄﾞなし")


class DistortionViewTest(unittest.TestCase):
    def test_122_B(self) -> None:
        # H9022K1: K21S NP56 O 23×670×2700。DIM_122_B 先頭行の用途に K21S → 材質は右3文字 "P56" と調質 O で絞り
        # 1–75 かつ 幅 0.1–9999 の最初の行 → 板幅 2.5 / 板丈 3.5、中耳4欄「不問」
        f = qv.display_distortion(q.read_lot("H9022K1"), qv.QualityForm()).form
        self.assertEqual((f.HaHiz1, f.TaHiz1, f.Naka1, f.Mimi2), ("2.5", "3.5", "不問", "不問"))
        self.assertEqual(f.Sansyou1, " Q-S-QA-DIM-122_B")

    def test_122_A_and_boundary(self) -> None:
        # H6083T1: H205 R 製造 85×1529×2502。白銅ではない → 50.01–100 の行(MG 1.2)。幅 1529 は (1254,1529] → 2
        f = qv.display_distortion(q.read_lot("H6083T1"), qv.QualityForm()).form
        self.assertEqual((f.MG1, f.HaHiz1, f.TaHiz1, f.Noun_BackColor), ("1.2", "2.0", "3.0", qv.WHITE))
        self.assertEqual(f.Sansyou1, " Q-S-QA-DIM-122")
        # H8401Z0: H283 15 → 12–50 の行(MG 0.4)、1528 → 1529 列 0.6、3053 → 3054 列 0.7
        f = qv.display_distortion(q.read_lot("H8401Z0"), qv.QualityForm()).form
        self.assertEqual((f.MG1, f.HaHiz1, f.TaHiz1), ("0.4", "0.6", "0.7"))

    def test_naka_mimi_and_122_D(self) -> None:
        # N7120E1: S139 1070 H18 オーダー 6.0×1150×800 等級4。6 以下 → 耳歪4級/中歪4級(5.01–12、800 → 4000 列 2)
        # その後 DIM_122_D(1070 → "070"、H*8、S139、ｱｷﾀｾｲﾚﾝ、6.0)で 1.5/1(QDB:1322)。中耳は最後に「不問」で上書き
        f = qv.display_distortion(q.read_lot("N7120E1"), qv.QualityForm()).form
        self.assertEqual((f.HaHiz1, f.TaHiz1, f.Naka1, f.Toukyu), ("1.5", "1.0", "不問", "4級"))
        self.assertEqual(f.Sansyou1, " ﾌﾗｯﾄﾈｽ製品規格中耳歪(ｵｰﾀﾞｰ丈判定) Q-S-QA-DIM-122_D")
        self.assertEqual(f.Noun_BackColor, qv.HX8_BLUE)

    def test_grade3_and_dedupe(self) -> None:
        # L8281G0: 0.3 等級3。4級 0.01–0.5、丈 0 → 最大列 4000 = 2。3級も同じく 4。参照元の重複は除く
        f = qv.display_distortion(q.read_lot("L8281G0"), qv.QualityForm(Sansyou1="前回 前回")).form
        self.assertEqual((f.Naka1, f.Mimi1, f.Naka2, f.Mimi2, f.Toukyu), ("2.0", "2.0", "4.0", "4.0", "3級"))
        self.assertEqual(f.Sansyou1, "前回 ﾌﾗｯﾄﾈｽ製品規格中耳歪(ｵｰﾀﾞｰ丈判定)")

    def test_122_C_no_match_B27(self) -> None:
        # L9280A0: R292 BE5Y O 1.0。DIM_122_C を調質 O で絞るが一致なし → 再現は「不問」と参照元を付ける(QDB:1276-1282)
        f = qv.display_distortion(q.read_lot("L9280A0"), qv.QualityForm()).form
        self.assertEqual(f.Naka1, "不問")
        self.assertTrue(f.Sansyou1.endswith(" Q-S-QA-DIM-122_O"))
        quirks.set_for_test({"B27": "fix"})
        try:
            f = qv.display_distortion(q.read_lot("L9280A0"), qv.QualityForm()).form
            self.assertEqual(f.Naka1, "2.0")                       # 中歪4級 0.51–2、丈 0 → 4000 列 = 2 が残る
            self.assertNotIn("_O", f.Sansyou1)
        finally:
            quirks.set_for_test(None)

    def test_122_C_match(self) -> None:
        # H6540C0: R112 NP56 O 16 ｽﾐﾄﾓｾｲﾐﾂ → 管理番号1行目(0.1–999、すべて)と NP56 の行に一致、最後の一致 → 10/10、黄
        f = qv.display_distortion(q.read_lot("H6540C0"), qv.QualityForm()).form
        self.assertEqual((f.HaHiz1, f.TaHiz1, f.Noun_BackColor), ("10.0", "10.0", qv.O_YELLOW))

    def test_122_D_grade_B27(self) -> None:
        # 等級ありの行(K184 NP56 H38 ﾐｸﾆﾊﾟ-ﾃﾂｸ 2.0): 再現はソリ表(列名 H/T)を LoopArray で引いて添字エラー
        lot = lot_with("N7120E1", {"用途ｺｰﾄﾞ": "K184", "製造材質": "NP56", "製造調質": "H38",
                                   "納入先名": "ﾐｸﾆﾊﾟ-ﾃﾂｸ(ｶ)ｻｶﾞﾗｺｳｼﾞﾖｳ", "ｵｰﾀﾞｰ板厚": "2", "ｵｰﾀﾞｰ板幅": "1000",
                                   "ｵｰﾀﾞｰ板丈": "2000", "品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ": "4"})
        with self.assertRaises(q.VbaRuntimeError):
            qv.display_distortion(lot, qv.QualityForm())
        quirks.set_for_test({"B27": "fix"})
        try:
            f = qv.display_distortion(lot, qv.QualityForm()).form
            # ソリ4級 厚 2 → 0.51–2 の行。H1000 = 2、T2000 = 3
            self.assertEqual((f.HaHiz1, f.TaHiz1), ("2.0", "3.0"))
            self.assertIn(" 製品規格_ソリ4級", f.Sansyou1)
        finally:
            quirks.set_for_test(None)


class DistortionVcTest(_DbCopy):
    def test_qa_022_01(self) -> None:
        # H8401Z0(VC 01、受注 06023667)の VC_表 を 208PRNF に → QA_022_01。製造厚 15 → 12–50 の行
        # MG1 は Format("0.4以下") = 文字のまま。幅 1525 → (1255,1530] → 0.6、丈 3050 → (2505,3055] → 0.7
        self.sql("SIKAODRNOW", 'UPDATE "仕掛" SET "VC_表"=? WHERE "受注番号"=?', "208PRNF", "06023667")
        f = qv.display_distortion(q.read_lot("H8401Z0"), qv.QualityForm(), odr_db="SIKAODRNOW").form
        self.assertEqual((f.MG1, f.HaHiz1, f.TaHiz1, f.VcFlag), ("0.4以下", "0.6", "0.7", "あり"))
        self.assertEqual(f.Sansyou1, " QA-022-01")


# ------------------------------------------------------------------
# 突き合わせ検証で見つけた差(VBA の手追い)
# ------------------------------------------------------------------
class VerifierRangeCheckTest(unittest.TestCase):
    def test_value_column_always_read(self) -> None:
        # CheckThickness/CheckWidth の `A = "" And B = "" And C = ""` は短絡しない(QP:930-932・949-951)。
        # 値列(Col3)が 0(表に無い)なら、下・上に値があっても ArrQuality(row, 0) で添字エラー 9
        with self.assertRaises(q.VbaRuntimeError) as cm:
            q.check_thickness(["4", "50"], 1, 2, 0, 5)
        self.assertEqual(cm.exception.number, 9)
        with self.assertRaises(q.VbaRuntimeError):
            q.check_width(["4", "50"], 1, 2, 3, 5)


class VerifierEmptyFieldsTest(_DbCopy):
    """`Dim PackF(): PackF = GetFieldsArr(…)` は表が空(Empty)だと代入で型エラー 13。
    PackingPath へ切り替える `If IsEmpty(PackF) Then …` には到達しない。"""

    def test_db_return_item_empty_table(self) -> None:
        self.sql("品質基準", 'DELETE FROM "ラテフリ直角"')
        with self.assertRaises(vbx.TypeMismatch):          # QP:1184
            q.db_return_item("ｽﾐﾄﾓｾｲﾐﾂｺｳｷﾞﾖｳ(ｶ", "x", "R115", "3S", "O", 10, 1000, 2000)

    def test_transfer_tables_empty(self) -> None:
        self.sql("伝送用ファイル", 'DELETE FROM "設備仕様"')
        with self.assertRaises(vbx.TypeMismatch):          # QP:2091
            q.check_equipment_db("LVC", "52S", 2, 1000, 2000, 30)
        self.sql("伝送用ファイル", 'DELETE FROM "時間用"')
        with self.assertRaises(vbx.TypeMismatch):          # QP:1792
            q.shift_time()
        self.sql("伝送用ファイル", 'DELETE FROM "校正ストレートエッヂ"')
        with self.assertRaises(vbx.TypeMismatch):          # QP:1837
            q.straight_edge("LVC")


if __name__ == "__main__":
    unittest.main()


# ------------------------------------------------------------------
# 直した場合の分岐(B71・B72・B73・B74)
# ------------------------------------------------------------------
class FixBranchTest(unittest.TestCase):
    def setUp(self) -> None:
        kensa_fixture.use_offline()
        q.clear_caches()

    def tearDown(self) -> None:
        quirks.set_for_test(None)

    def test_b71_boundary_same_precision(self) -> None:
        args = ("ﾆﾂﾎﾟﾝﾌﾙﾊ-ﾌ(ｶ", "x", "F975", "52S", "H18", 0.6, 1217, 2438)
        self.assertEqual(q.db_return_item(*args).items, "")          # 再現: Single(0.6) > 0.6 で外れる
        quirks.set_for_test({"B71": "fix"})
        r = q.db_return_item(*args)                                  # 直す: 管理番号1(0.6–0.6)に一致
        self.assertEqual(r.items, ",フリー反り")
        self.assertEqual((r.uf1["HURIS"], r.uf1["HizCome.Tag"]), ("40", "40mm以下"))
        # 0.65–0.65(管理番号2、丈 2730)
        self.assertEqual(q.db_return_item("ﾆﾂﾎﾟﾝﾌﾙﾊ-ﾌ(ｶ", "x", "F975", "52S", "H18", 0.65, 1217, 2730).items,
                         ",フリー反り")
        # 範囲の外はそのまま外れる
        self.assertEqual(q.db_return_item("ﾆﾂﾎﾟﾝﾌﾙﾊ-ﾌ(ｶ", "x", "F975", "52S", "H18", 0.61, 1217, 2438).items, "")
        # check_thickness の既定(他の表)は Single のまま
        self.assertFalse(q.check_thickness(["0.6", "0.6", "0.6"], 1, 2, 3, 0.6))

    def test_b72_lateral_tolerance_from_lateral_row(self) -> None:
        args = ("ｽﾐﾄﾓｾｲﾐﾂｺｳｷﾞﾖｳ(ｶ", "x", "R115", "3S", "O", 10, 1000, 2000)
        self.assertEqual(q.db_return_item(*args).uf1["LateralText"], "1.0")   # 再現: 直角度の行の "1"
        quirks.set_for_test({"B72": "fix"})
        r = q.db_return_item(*args)
        self.assertEqual(r.items, ",直角度,ラテラルボー")
        # 直す: ラテラルボーが「要」の行(管理番号10)の公差 "等級" → 数でも別表でもないので書かない(QP:1366)
        self.assertNotIn("LateralText", r.uf1)
        self.assertEqual(r.uf1["Sumisei.Caption"], "ｽﾐﾄﾓｾｲﾐﾂｺｳｷﾞﾖｳ(ｶ")

    def test_b73_122_D_no_match(self) -> None:
        # LX120J0: DIM_122_D に一致なし。再現は中耳4欄を「不問」で上書き(QDB:1366-1369)
        f = qv.display_distortion(q.read_lot("LX120J0"), qv.QualityForm()).form
        self.assertEqual((f.Mimi1, f.Naka1, f.Mimi2, f.Naka2), ("不問",) * 4)
        quirks.set_for_test({"B73": "fix"})
        f = qv.display_distortion(q.read_lot("LX120J0"), qv.QualityForm()).form
        self.assertEqual((f.Mimi1, f.Naka1, f.Mimi2, f.Naka2), ("2.0", "2.0", "", ""))   # 等級表の値が残る
        # 一致があれば今までどおり「不問」(N7120E1)
        f = qv.display_distortion(q.read_lot("N7120E1"), qv.QualityForm()).form
        self.assertEqual((f.HaHiz1, f.TaHiz1, f.Naka1, f.Mimi1), ("1.5", "1.0", "不問", "不問"))

    def test_b74_grade_not_3_4_in_skip_qp(self) -> None:
        # L1099R0(0.5mm)は QP 表に一致なし → SkipQP。等級 5 だと TableN が空のまま表を開いて止まる
        lot = lot_with("L1099R0", {"品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ": "5"})
        from ufkensa.kensa import datasource
        with self.assertRaises(datasource.TableMissing):
            q.dbqp("L1099R0", "H", "x", "x", lot)
        quirks.set_for_test({"B74": "fix"})
        r = q.dbqp("L1099R0", "H", "x", "x", lot)
        self.assertEqual([m.text for m in r.messages], ["歪等級が3,4以外です"])
        self.assertEqual(r.kousa[0][q.K_NAKA:q.K_MIMI + 1], [None, None])         # 中耳は空のまま skipEnd
        # 指示値があれば skipEnd で補う(QP:779-780)
        lot = lot_with("L1099R0", {"品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ": "5", "指示_中歪_高さ": "2", "指示_耳歪_高さ": ""})
        r = q.dbqp("L1099R0", "H", "x", "x", lot)
        self.assertEqual(r.kousa[0][q.K_NAKA], "2")
        # 等級 3/4 は変わらない
        r = q.dbqp("L1099R0", "H", "x", "x", q.read_lot("L1099R0"))
        self.assertEqual(r.messages, [])


# ------------------------------------------------------------------
# B71・B72 を同じ種類の箇所すべてに広げた分(直した側と再現側)
# ------------------------------------------------------------------
class FixB71WideTest(_DbCopy):
    """B71: 下限・上限を Single に入れて Double(Val(Single) など)と比べ、境界ちょうどが外れる。"""

    def test_check_within_param(self) -> None:
        row = ["2.1", "2.1"]
        self.assertFalse(q.check_within(row, 1, 2, 2.1))                       # Single(2.1) > 2.1(QP:967-970)
        self.assertTrue(q.check_within(row, 1, 2, 2.1, same_precision=True))

    def test_dbqp_thickness(self) -> None:
        # QP:313: 管理番号6 の板厚を 85.3–85.3、オーダー厚 85.3 → 再現は 6 が外れて [105] だけ
        self.sql("品質基準", 'UPDATE "平坦度一覧QP追加" SET "板厚下"=?, "板厚上"=? WHERE "管理番号"=?', "85.3", "85.3", "6")
        lot = lot_with("H6083T1", {"ｵｰﾀﾞｰ板厚": "85.3"})
        r = q.dbqp("H6083T1", "H6X1922", "ｶ) ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ", "x", lot)
        self.assertEqual([x[q.K_NO] for x in r.kousa], ["105"])
        quirks.set_for_test({"B71": "fix"})
        r = q.dbqp("H6083T1", "H6X1922", "ｶ) ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ", "x", lot)
        self.assertEqual([x[q.K_NO] for x in r.kousa], ["6", "105"])

    def test_dbdisplay_lateral(self) -> None:
        # QDB:1637-1638: DIM_179_日車 の NP56 行は 厚 6.51–245・幅 450.1–9999。境界ちょうどで再現は一致なし → 0
        for change in ({"製造板厚": "6.51"}, {"製造板幅": "450.1"}):
            lot = lot_with("H9022K1", change)
            quirks.set_for_test(None)
            self.assertEqual(q.dbdisplay_lateral(lot, "日車"), 0.0)
            quirks.set_for_test({"B71": "fix"})
            self.assertEqual(q.dbdisplay_lateral(lot, "日車"), 3.0)              # 丈 2703 → (2300,3050] → 3

    def test_check_equipment_db(self) -> None:
        # QP:2140-2143: 管理番号2(LVC)の単重MIN を 30.1 → 単重 30.1 は再現で「単重 範囲外」
        self.sql("伝送用ファイル", 'UPDATE "設備仕様" SET "単重MIN"=? WHERE "管理番号"=?', "30.1", "2")
        self.assertEqual(q.check_equipment_db("LVC", "52S", 2, 1000, 2000, 30.1), "単重 範囲外")
        quirks.set_for_test({"B71": "fix"})
        self.assertEqual(q.check_equipment_db("LVC", "52S", 2, 1000, 2000, 30.1), "可")
        self.assertEqual(q.check_equipment_db("LVC", "52S", 2, 1000, 2000, 30.09), "単重 範囲外")   # 外は外のまま

    def test_appearance_h_124(self) -> None:
        # QDB:300: 別表2 の 15・21 は 板厚 10.01–30。製造厚 10.01 は再現で両方外れ(グレードなし)
        lot = lot_with("H8401Z0", {"製造板厚": "10.01"})
        self.assertIsNone(qv.display_appearance_h(lot, qv.QualityForm()).grade_array)
        quirks.set_for_test({"B71": "fix"})
        r = qv.display_appearance_h(lot, qv.QualityForm())
        self.assertEqual([g[2] for g in r.grade_array], ["4", "4", "4", "3", "3", "3"])   # 後の 21 が残る

    def test_122_C_thickness(self) -> None:
        # QDB:1245-1247: DIM_122_C の C211 ﾆﾂﾎﾟﾝﾌﾙﾊ-ﾌ 行(3.2–3.2、5/5)。オーダー厚 3.2 は再現で外れる
        lot = lot_with("H6540C0", {"用途ｺｰﾄﾞ": "C211", "納入先名": "ﾆﾂﾎﾟﾝﾌﾙﾊ-ﾌ(ｶ", "ｵｰﾀﾞｰ板厚": "3.2"})
        f = qv.display_distortion(lot, qv.QualityForm()).form
        self.assertEqual((f.HaHiz1, f.TaHiz1), ("", ""))
        quirks.set_for_test({"B71": "fix"})
        f = qv.display_distortion(lot, qv.QualityForm()).form
        self.assertEqual((f.HaHiz1, f.TaHiz1, f.Noun_BackColor), ("5.0", "5.0", qv.O_YELLOW))

    def test_122_C_width_pattern(self) -> None:
        # QDB:1257: K126 ｽ-ﾊﾟ-ｽﾀ- のパターン2 行(オーダー板幅 740.1 <= 幅 で 4/4)。幅 740.1 は Single(740.1) < 740.1 で外れる
        lot = lot_with("H6540C0", {"製造材質": "52S", "用途ｺｰﾄﾞ": "K126", "納入先名": "ｶ)ｽ-ﾊﾟ-ｽﾀ-", "ｵｰﾀﾞｰ板厚": "6",
                                   "ｵｰﾀﾞｰ板幅": "740.1"})
        self.assertEqual(qv.display_distortion(lot, qv.QualityForm()).form.HaHiz1, "")
        quirks.set_for_test({"B71": "fix"})
        self.assertEqual(qv.display_distortion(lot, qv.QualityForm()).form.HaHiz1, "4.0")
        # パターン1(740 >= 幅)はちょうど 740 で今までどおり 2/2
        lot = lot_with("H6540C0", {"製造材質": "52S", "用途ｺｰﾄﾞ": "K126", "納入先名": "ｶ)ｽ-ﾊﾟ-ｽﾀ-", "ｵｰﾀﾞｰ板厚": "6",
                                   "ｵｰﾀﾞｰ板幅": "740"})
        self.assertEqual(qv.display_distortion(lot, qv.QualityForm()).form.HaHiz1, "2.0")


class FixB71MaterialsTest(MaterialsUpperPartTest):
    """QP:1764-1766: 梱包保護材の 板厚下 0.6 と 厚 0.6。"""

    def tearDown(self) -> None:
        quirks.set_for_test(None)
        super().tearDown()

    def test_match(self) -> None:          # 親の試験は繰り返さない
        with sqlite3.connect(self.db) as c:
            c.execute('UPDATE 梱包保護材 SET "板厚下"=? WHERE "管理番号"=?', ("0.6", "1"))
        c.close()
        h = hist([None, "12345", "L1", 0.6, 1000.0, 2000.0, "H", "F52S", "H34", "H197", "", "1P0104"])
        self.assertEqual(q.materials_upper_part("12345", h, db=self.db), "アングル")    # 再現: 1 が外れて 2
        quirks.set_for_test({"B71": "fix"})
        self.assertEqual(q.materials_upper_part("12345", h, db=self.db), "ザラ板")

    def test_target_list(self) -> None:
        pass


class FixB72WideTest(_DbCopy):
    def test_free_sori_tolerance_from_free_row(self) -> None:
        # ｽﾐﾄﾓ R115: 一致 [9(直角度・公差 1), 10]。10 をフリー反り「要」・公差 30 にする
        # 再現: HizCome.Tag・HURIS は HitArry(1)(行 9)の 1(QP:1395・1407)。直す: フリー反りが「要」の行 10 の 30
        self.sql("品質基準", 'UPDATE "ラテフリ直角" SET "フリー反り"=?, "公差"=? WHERE "管理番号"=?', "要", "30", "10")
        args = ("ｽﾐﾄﾓｾｲﾐﾂｺｳｷﾞﾖｳ(ｶ", "x", "R115", "3S", "O", 10, 1000, 2000)
        r = q.db_return_item(*args)
        self.assertEqual(r.items, ",直角度,フリー反り,ラテラルボー")
        self.assertEqual((r.uf1["HizCome.Tag"], r.uf1["HURIS"], r.uf1["HizCome.Caption"]),
                         ("1mm以下", "1", "ﾌﾘｰ反り1mm以下"))
        quirks.set_for_test({"B72": "fix"})
        r = q.db_return_item(*args)
        self.assertEqual((r.uf1["HizCome.Tag"], r.uf1["HURIS"], r.uf1["HizCome.Caption"]),
                         ("30mm以下", "30", "ﾌﾘｰ反り30mm以下"))
        # UFMitsubishi 表示中は HizCome を書かず、ループ中の HURIS(要の行。QP:1337)のまま
        r = q.db_return_item(*args, mitsubishi_visible=True, work_flag="P1")
        self.assertEqual(r.uf1["HURIS"], "30")
        self.assertEqual(r.cells[(1, 13)], "30.0")                             # ラテラル公差も行 10(既存の B72)
