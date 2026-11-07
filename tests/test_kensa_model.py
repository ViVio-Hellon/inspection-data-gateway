"""model.py(第3段 案C の型): 欄の名前・厚の行・書式・判定。"""
from __future__ import annotations

import unittest

from ufkensa.kensa import grid, model
from ufkensa.kensa.model import CellKey, Verdict


class ModelTest(unittest.TestCase):
    def test_cell_key(self) -> None:
        self.assertEqual(CellKey.parse("S1"), CellKey("S", 1))
        self.assertEqual(CellKey.parse("E14").name, "E14")
        self.assertIsNone(CellKey.parse("F11"))            # F は 10 行まで
        self.assertIsNone(CellKey.parse("SBH"))
        self.assertIsNone(CellKey.parse("S0"))
        self.assertTrue(CellKey("C", 4).thick)
        self.assertFalse(CellKey("C", 5).thick)

    def test_thick_order_and_formats(self) -> None:
        keys = [k.name for k in model.thick_keys(("S", "U"))]
        self.assertEqual(keys, ["S1", "U1", "S2", "U2", "S3", "U3", "S4", "U4"])
        self.assertEqual(model.row_format(4), "0.000")
        self.assertEqual(model.row_format(5), "0.0")
        self.assertEqual(model.thick_format("74"), "0.000")
        self.assertEqual(model.thick_format("74.5"), "0.00")

    def test_row_layouts(self) -> None:
        self.assertEqual(model.NORMAL.labels(), ["厚1", "厚2", "厚3", "厚4", "幅1", "幅2", "丈1", "丈2", "対角A", "対角B"])
        self.assertEqual(list(model.NORMAL.size_rows), [5, 6, 7, 8, 9, 10])
        self.assertEqual(model.EXTREME.labels()[4:], ["幅1", "幅2", "幅3", "幅4", "幅5", "幅6", "丈1", "丈2", "丈3", "丈4"])
        self.assertEqual(list(model.EXTREME.length), [11, 12, 13, 14])
        self.assertIs(model.row_layout(True), model.EXTREME)

    def test_verdict_is_the_color(self) -> None:
        self.assertEqual(Verdict.NG, grid.COLOR_NG)
        self.assertEqual(Verdict.of(grid.COLOR_OK), Verdict.OK)
        self.assertEqual(Verdict.of(grid.COLOR_BUTTON), Verdict.NONE)

    def test_measures_view(self) -> None:
        st = grid.KensaState()
        st.cells["S2"].value, st.cells["S2"].color = "15.012", grid.COLOR_NG
        m = st.measures
        self.assertEqual(m.thick("S"), ["", "15.012", "", ""])
        self.assertEqual(m.verdict("S", 2), Verdict.NG)
        self.assertIs(m.ctl(CellKey("S", 2)), st.cells["S2"])      # 置き場所は同じ(写しではない)


if __name__ == "__main__":
    unittest.main()
