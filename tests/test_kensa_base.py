"""UF検査判定の土台(VBA 互換・不具合スイッチ・ADO Filter・DB の読み方・ライン名・使用パス)"""
from __future__ import annotations

import unittest

from tests import kensa_fixture
from ufkensa.kensa import ado_filter, datasource, lines, paths, quirks, vbx


class VbxTest(unittest.TestCase):
    def test_single(self) -> None:
        # 0.5 刻み変換(UnitRecord)で問題になる値: Single では 1.01 が 1.00999999…
        self.assertLess(vbx.single(1.01), 1.01)
        self.assertEqual(vbx.cstr(1.01, is_single=True), "1.01")
        self.assertEqual(vbx.val_single(1.01), 1.01)

    def test_cstr_15_digits(self) -> None:
        self.assertEqual(vbx.cstr(1.005 * 1000), "1005")      # Python の repr は 1004.9999999999999
        self.assertEqual(vbx.cstr(0.1 + 0.2), "0.3")
        self.assertEqual(vbx.cstr(2053.2), "2053.2")
        self.assertEqual(vbx.cstr(-0.5), "-0.5")
        self.assertEqual(vbx.cstr(None), "")

    def test_val(self) -> None:
        self.assertEqual(vbx.val(" 12 3abc"), 123.0)            # 空白はどこでも無視
        self.assertEqual(vbx.val("1,000"), 1.0)                  # カンマで止まる
        self.assertEqual(vbx.val(""), 0.0)
        self.assertEqual(vbx.val("-.5"), -0.5)
        self.assertEqual(vbx.val("&H10"), 16.0)

    def test_format_half_up(self) -> None:
        self.assertEqual(vbx.fmt(2.675, "0.00"), "2.68")
        self.assertEqual(vbx.fmt(0.125, "0.00"), "0.13")         # Round(0.125, 2) は 0.12
        self.assertEqual(vbx.fmt(-0.04, "0.0"), "0.0")           # -0 を出さない
        self.assertEqual(vbx.fmt("", "0.000"), "")
        self.assertEqual(vbx.fmt("3", "0.000"), "3.000")
        self.assertEqual(vbx.fmt("S", "0.0"), "S")
        self.assertEqual(vbx.fmt(1234.5, "#,##0"), "1,235")
        self.assertEqual(vbx.fmt(7, "00"), "07")

    def test_round_bankers(self) -> None:
        self.assertEqual(vbx.vround(2.5), 2.0)
        self.assertEqual(vbx.vround(3.5), 4.0)
        self.assertEqual(vbx.vround(0.125, 2), 0.12)
        self.assertEqual(vbx.clng(2.5), 2)
        self.assertEqual(vbx.clng("3.5"), 4)

    def test_compare(self) -> None:
        self.assertEqual(vbx.vcompare(2, "2.0"), 0)
        self.assertEqual(vbx.vcompare("b", "a"), 1)
        with self.assertRaises(vbx.TypeMismatch):
            vbx.vcompare(2, "S")
        with self.assertRaises(vbx.TypeMismatch):
            vbx.cdbl("")

    def test_compare_typed_follows_vba_table(self) -> None:
        N, S, V = vbx.NUM, vbx.STR, vbx.VAR
        with self.assertRaises(vbx.TypeMismatch):               # 数値型 と 数にならない文字列の Variant
            vbx.compare_typed("", 0, V, N)
        self.assertEqual(vbx.compare_typed("", 0, V, V), 1)     # Variant どうし: 文字列 > 数(エラーにしない)
        self.assertEqual(vbx.compare_typed("10", 9.0, S, V), -1)  # String と Variant: 文字で比べる("10" < "9")
        self.assertEqual(vbx.compare_typed("3", 2, V, N), 1)    # 数になる Variant と数値型: 数で比べる
        self.assertEqual(vbx.compare_typed(None, 0, V, N), 0)   # Empty は 0
        self.assertEqual(vbx.compare_typed("10", 9, S, N), 1)   # String と数値型: 数に直す(.DiagS = 0 など)
        with self.assertRaises(vbx.TypeMismatch):
            vbx.compare_typed("", 0, S, N)

    def test_like_mid_instr(self) -> None:
        self.assertTrue(vbx.like("ABC52S1", "*52S*"))
        self.assertFalse(vbx.like("abc", "A*"))                 # バイナリ比較
        self.assertEqual(vbx.mid("DT0001", 3, 2), "00")
        self.assertEqual(vbx.instr("ﾗﾃﾗﾙﾎﾞｰ検査不要", "検査不要"), 8)
        self.assertEqual(vbx.instr("", ""), 0)                  # 調べる文字列が空なら 0
        self.assertEqual(vbx.instr("ABC", ""), 1)               # 探す文字列が空なら開始位置   # ﾎﾞ は半角2文字


class QuirksTest(unittest.TestCase):
    def test_summary_matches_decision(self) -> None:
        sm = quirks.summary()
        self.assertEqual(len(sm["fix"]), 90)   # 2026-11-09: B93 を足した
        self.assertIn("B19", sm["keep"])
        self.assertIn("B37", sm["keep"])

    def test_93_items(self) -> None:
        self.assertEqual(len(quirks.TITLES), 93)
        self.assertEqual(list(quirks.TITLES), [f"B{i}" for i in range(1, 94)])


class AdoFilterTest(unittest.TestCase):
    rows = [
        {"ﾛｯﾄ番号": "H7372B0", "ｵｰﾀﾞｰ板丈": "2000", "BOX設計_設備名": "KEN", "残仕掛設備ｺｰｽ": "L-3 KEN"},
        {"ﾛｯﾄ番号": "h7372b0", "ｵｰﾀﾞｰ板丈": "0", "BOX設計_設備名": "KEN", "残仕掛設備ｺｰｽ": "HOT"},
        {"ﾛｯﾄ番号": "X", "ｵｰﾀﾞｰ板丈": "", "BOX設計_設備名": None, "残仕掛設備ｺｰｽ": ""},
    ]

    def pick(self, expr: str) -> list[int]:
        f = ado_filter.compile_filter(expr)
        return [i for i, r in enumerate(self.rows) if f(r)]

    def test_eq_is_case_insensitive(self) -> None:
        self.assertEqual(self.pick("[ﾛｯﾄ番号]='H7372B0'"), [0, 1])

    def test_number_against_text_column(self) -> None:
        self.assertEqual(self.pick("[ｵｰﾀﾞｰ板丈]<>0"), [0, 2])
        self.assertEqual(self.pick("[ｵｰﾀﾞｰ板丈]=0  AND  [BOX設計_設備名] = 'KEN' "), [1])

    def test_like_and(self) -> None:
        self.assertEqual(self.pick("[ｵｰﾀﾞｰ板丈]<>0 AND [残仕掛設備ｺｰｽ] LIKE '*KEN*' AND  [BOX設計_設備名] = 'KEN' "), [0])

    def test_null_equals_empty(self) -> None:
        self.assertEqual(self.pick("[BOX設計_設備名]=''"), [2])

    def test_trailing_space_does_not_match(self) -> None:
        self.assertEqual(self.pick("[ﾛｯﾄ番号]='X '"), [])        # B20 の前提

    def test_or_and_paren(self) -> None:
        self.assertEqual(self.pick("([ﾛｯﾄ番号]='X') OR [ｵｰﾀﾞｰ板丈]='0'"), [1, 2])

    def test_column_name_case(self) -> None:
        self.assertTrue(ado_filter.compile_filter("[No]='1'")({"NO": "1"}))   # ADO は列名の大小を区別しない

    def test_unknown_column(self) -> None:
        with self.assertRaises(KeyError):
            self.pick("[無い列]='1'")


class DatasourceTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        kensa_fixture.use_offline()

    def test_records_are_text_in_import_order(self) -> None:
        recs = datasource.get_records("伝送用ファイル", "時間用")
        self.assertEqual([r[1:4] for r in recs],
                         [["1", "07:00", "15:00"], ["2", "15:00", "22:50"], ["3", "22:50", "07:00"], ["昼", "08:15", "17:00"]])
        self.assertTrue(all(isinstance(v, str) for r in recs for v in r))

    def test_filter_and_empty(self) -> None:
        f = datasource.get_fields("伝送用ファイル", "測定機器ID", "[ライン]='L-1'")
        self.assertEqual(f[:4], ["ライン", "機器名", "ID", "出力テキスト"])
        self.assertIsNone(datasource.get_records("伝送用ファイル", "測定機器ID", "[ライン]='無い'"))
        self.assertIsNone(datasource.get_fields("伝送用ファイル", "測定機器ID", "[ライン]='無い'"))

    def test_numbers_become_vba_text(self) -> None:
        recs = datasource.get_records("伝送用ファイル", "キャスク梱包板番号")
        self.assertEqual(recs[0][2], "23")                        # Long は整数の文字
        self.assertEqual(recs[0][4], "2026/02/24 08:59:00")       # 日付は bridge と同じ書き方

    def test_missing_file_does_not_create_one(self) -> None:
        with self.assertRaises(datasource.DbMissing):
            datasource.get_records(str(kensa_fixture.build() / "無い.sqlite3"), "仕掛")
        self.assertFalse((kensa_fixture.build() / "無い.sqlite3").exists())

    def test_column_index(self) -> None:
        f = datasource.get_fields("伝送用ファイル", "測定機器ID")
        self.assertEqual(datasource.column_index(f, "備考"), 7)
        self.assertEqual(datasource.column_index(f, "備"), -1)    # VBA の C_備考 = 0 に当たる(未使用)


class LinesTest(unittest.TestCase):
    def test_width_only(self) -> None:
        self.assertTrue(lines.matches("トット", "ﾄｯﾄ"))
        self.assertTrue(lines.matches("ﾊﾞﾗﾝｻｰ", "バランサー"))
        self.assertFalse(lines.matches("l-1", "L-1"))            # 大文字小文字は揃えない
        self.assertEqual(lines.official("丸徳"), "中板")
        self.assertEqual(lines.vba_name("トット"), "ﾄｯﾄ")
        self.assertEqual(lines.find("トット").folder, "トット")


class PathsTest(unittest.TestCase):
    def test_vba_defaults(self) -> None:
        self.assertEqual(paths.default("main", paths.ONLINE), "\\\\nlmsrvngy03\\Read\\【New】仕掛\\台帳\\")
        self.assertEqual(paths.default("nis", paths.OFFLINE), r"C:\Tools\測定データ電子化")   # D26 改: パス1つ(モードによらない)
        self.assertTrue(paths.default("history", paths.OFFLINE).endswith(("/", "\\")))
        self.assertTrue(paths.db_file("梱包検査ﾃﾞｰﾀ", paths.ONLINE).endswith("【測定DB】梱包\\梱包検査ﾃﾞｰﾀ.sqlite3"))

    def test_override_and_reset(self) -> None:
        paths.set_path("wav", r"D:\音", paths.ONLINE)
        self.assertEqual(paths.get("wav", paths.ONLINE), r"D:\音")
        paths.reset_path("wav", paths.ONLINE)
        self.assertEqual(paths.get("wav", paths.ONLINE), r"C:\VBAWAV" "\\")

    def test_every_item_has_a_label_and_use(self) -> None:
        for row in paths.overview(paths.ONLINE):
            self.assertTrue(row["label"] and row["use"], row)


if __name__ == "__main__":
    unittest.main()
