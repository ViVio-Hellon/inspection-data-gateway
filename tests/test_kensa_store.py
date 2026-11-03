"""手元DB(store.py): シートの形・Excel の代入の型・行の操作"""
from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

from ufkensa.kensa import store as st


class ExcelValueTest(unittest.TestCase):
    """セルへの代入で値がどうなるか(サンプル xlsm のセルの型で確かめた規則)。"""

    def test_variant_string_becomes_number(self) -> None:
        # 作業保存_Test!B1: Array(.ATU.Text …)(Variant)で "25.000" → 数 25(SaveLoad:2017)
        self.assertEqual(st.excel_value("25.000", "V"), 25.0)
        self.assertEqual(st.excel_value(" 12 ", "V"), 12.0)
        self.assertEqual(st.excel_value("1,000", "V"), 1000.0)
        self.assertEqual(st.excel_value("1E5", "V"), 100000.0)
        self.assertEqual(st.excel_value("-0.5", "V"), -0.5)

    def test_variant_text_and_blank(self) -> None:
        self.assertEqual(st.excel_value("H207", "V"), "H207")
        self.assertEqual(st.excel_value("1,00", "V"), "1,00")          # 3 桁区切りでない
        self.assertEqual(st.excel_value("E5", "V"), "E5")
        self.assertIsNone(st.excel_value("", "V"))                      # 空セル
        self.assertEqual(st.excel_value("'3-2", "V"), "3-2")            # 先頭の ' は文字列の印
        self.assertEqual(st.excel_value("'-0.5", "V"), "-0.5")

    def test_string_array_and_object(self) -> None:
        # 作業保存_Test!Q2: String 型の配列 → '0.0' のまま。耳用!C2: Array(.ATU, …)(コントロール)→ '352.500'
        self.assertEqual(st.excel_value("0.0", "S"), "0.0")
        self.assertIsNone(st.excel_value("", "S"))
        self.assertEqual(st.excel_value("352.500", "O"), "352.500")
        self.assertEqual(st.excel_value("", "O"), "")                   # 耳用!CZ2 は長さ 0 の文字列

    def test_cell_text(self) -> None:
        self.assertEqual(st.cell_text(None), "")
        self.assertEqual(st.cell_text(50.0), "50")
        self.assertEqual(st.cell_text(49.776), "49.776")
        self.assertEqual(st.cell_text("x"), "x")

    def test_columns(self) -> None:
        self.assertEqual(st.col_letter(114), "DJ")
        self.assertEqual(st.col_letter(108), "DD")
        self.assertEqual(st.col_index("DJ"), 114)
        self.assertEqual(len(st.HOZON_HEADERS), 114)
        self.assertEqual(st.HOZON_HEADERS[113], "F_直角度")
        self.assertEqual(st.MIMI_HEADERS[107], "E_BH")
        self.assertEqual(st.MIMI_HEADERS[25], "S定盤歪")


class StoreTest(unittest.TestCase):
    def setUp(self) -> None:
        self.dir = Path(tempfile.mkdtemp())
        self.store = st.LocalStore(self.dir / "local.sqlite3")

    def tearDown(self) -> None:
        self.store.close()
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_row_sheet(self) -> None:
        sh = self.store.sheet(st.HOZON)
        self.assertTrue(sh.exists())
        self.assertEqual(sh.get(1, 2), "LOTNO")
        self.assertEqual(sh.last_row(1), 1)
        self.assertEqual(sh.last_col(1), 114)                          # UC(sh, 1) = 見出しの数
        sh.write(2, 2, "A1", "O")
        sh.write(3, 2, "B1", "O")
        sh.write(3, 1, "x")
        self.assertEqual(sh.last_row(2), 3)
        self.assertEqual(sh.last_row(1), 3)
        sh.write(2, 200, "捨てる")                                    # 見出しより右は持たない
        sh.delete_row(2)
        self.assertEqual(sh.get(2, 2), "B1")
        self.assertEqual(sh.last_row(2), 2)
        with self.assertRaises(ValueError):
            sh.write(1, 1, "x")

    def test_grid_sheet(self) -> None:
        self.assertFalse(self.store.exists("作業保存_Test"))
        sh = self.store.add_sheet("作業保存_Test")
        self.assertTrue(self.store.exists("作業保存_Test"))
        sh.write(2, 3, "1.5")
        sh.write(5, 1, "x")
        self.assertEqual(sh.get(2, 3), 1.5)
        self.assertEqual(sh.last_row(1), 5)
        self.assertEqual(sh.last_row(9), 1)
        self.assertEqual(sh.last_col(2), 3)
        self.assertEqual(sh.last_cell(), (5, 3))
        self.assertEqual(sh.read_range(5, 1, 2, 1), [[None], [None], [None], ["x"]])   # 逆向きは並べ替え
        sh.set_color(4, 5, 123)
        self.assertEqual(sh.color(4, 5), 123)
        sh.delete_row(2)
        self.assertEqual(sh.get(4, 1), "x")
        self.store.add_sheet("作業保存_Test")                           # AddWS は作り直し
        self.assertEqual(self.store.sheet("作業保存_Test").used(), [])

    def test_import_cells(self) -> None:
        n = st.import_cells(self.store, st.HOZON, [(1, 1, "見出し"), (2, 2, "L1"), (2, 3, 50.0)])
        self.assertEqual(n, 2)
        self.assertEqual(self.store.sheet(st.HOZON).get(2, 3), 50.0)
        st.import_cells(self.store, "校正記録", [(2, 1, "2026.10.13"), (2, 5, 1999.9)])
        self.assertEqual(self.store.sheet("校正記録").get(2, 5), 1999.9)


if __name__ == "__main__":
    unittest.main()
