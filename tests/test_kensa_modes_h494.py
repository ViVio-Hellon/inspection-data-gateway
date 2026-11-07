"""H494 検査(h494.py): 幅 3 点(DT 受付・0.00 表示)・ラテラルボーは 7〜10 行・対角・B30"""
from __future__ import annotations

import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from tests import kensa_fixture
from ufkensa.kensa import flow, quirks
from ufkensa.kensa.modes import h494
from ufkensa.kensa.modes import sheet as sh

NOW = datetime(2026, 10, 30, 10, 0)


def h_book() -> sh.Book:
    b = sh.Book(now=lambda: NOW)
    ws = b.ws(sh.H494)
    ws.visible = True
    b.activate_sheet(sh.H494)
    f = b.mit
    f.visible, f.work_flag, f.line_name = True, "P_H494", "LVC"
    for (r, c), v in {(2, 4): "R8405C0", (3, 2): 10.0, (3, 4): 150.0, (3, 6): 2500.0, (2, 13): 0.2, (2, 15): 0.2,
                      (3, 13): 0.5, (3, 15): 0.5, (4, 13): 3.0, (4, 15): 3.0, (1, 19): 1.0, (1, 13): 1.0}.items():
        ws.set(r, c, v)
    for k, fm in sh.TOLERANCE_FORMULAS.items():
        ws.set_formula(*k, fm)
    return b


class H494SheetTest(unittest.TestCase):
    def tearDown(self) -> None:
        pass
    def test_width_dt_and_length(self) -> None:          # Cask:233-250, 1101-1110
        b = h_book()
        ws = b.ws(sh.H494)
        b.user_edit(7, 15, "DT10002+0000150.500M")          # 幅3(O 列)も幅の公差
        self.assertEqual((ws.get(7, 15), ws.get(7, 16)), (150.5, "○"))
        self.assertEqual(ws.number_format[(7, 15)], "0.00_);[赤](0.00)")
        b.user_edit(7, 17, "#L2497.0")
        self.assertEqual(ws.get(7, 18), "○")
        b.user_edit(7, 19, "#L2496.9")
        self.assertEqual(ws.get(7, 20), "×")

    def test_lateral_color_b81_fix(self) -> None:
        b = h_book()
        ws = b.ws(sh.H494)
        b.user_edit(7, 25, "0.3")
        self.assertEqual(ws.interior[(7, 25)], sh.COLOR_INDEX[10])           # 緑が残る
        self.assertEqual(ws.interior[(7, 26)], sh.LIGHT_BLUE)                # 次のセルは薄青
        b.user_edit(8, 25, "3.0")
        self.assertEqual(ws.interior[(8, 25)], sh.NG_RED)
        b.activate(7, 25)                                                    # 判定済みのセルを選んでも色は変えない
        b.activate(9, 2)
        self.assertEqual(ws.interior[(7, 25)], sh.COLOR_INDEX[10])
        self.assertEqual(ws.interior[(9, 2)], sh.LIGHT_BLUE)
        self.assertEqual(ws.interior[(7, 26)], (255, 255, 255))

    def test_diagonal(self) -> None:                     # Cask:1555-1573
        b = h_book()
        ws = b.ws(sh.H494)
        b.user_edit(7, 21, "#L2502.0")
        b.user_edit(7, 22, "#L2501.0")
        self.assertEqual((ws.get(7, 23), ws.get(7, 24)), ("○", 1.0))
        self.assertEqual(b.active_cell, (8, 2))

class H494PrintTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        kensa_fixture.use_offline(kensa_fixture.copy_to(Path(self.tmp.name)))

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def test_print_yes(self) -> None:
        b = h_book()
        b.mit.h494_kname = "検査員02"
        b.other_sheets.add("保存用")
        res = h494.print_and_save(b, {"h494_print": True})
        self.assertEqual(res.history[0][83], "H494")
        self.assertEqual(res.history[0][3], 10.0)
        self.assertEqual(res.prints[0].sheet, sh.H494)
        self.assertTrue(any(s.path.endswith("20261030_R8405C0.xlsx") for s in res.saves[0].steps))
        self.assertEqual(b.mit.tag, "")

    def test_transfer_needs_method(self) -> None:        # UFMit:99-124
        b = h_book()
        fx = h494.transfer(b)
        self.assertEqual(fx.messages[-1].text, "検査法を選択してください")
        b.mit.h494_method, b.mit.h494_min_text, b.mit.h494_max_text = "L2", "1", "2"
        fx = h494.transfer(b)
        self.assertEqual(fx.calls[-1], sh.Call("Ainput_Startup", ("H494", "L2")))


if __name__ == "__main__":
    unittest.main()
