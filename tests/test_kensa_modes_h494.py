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
        quirks.set_for_test(None)

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

    def test_lateral_rows_and_blank(self) -> None:       # H494_SH:62-73, Cask:2070
        b = h_book()
        ws = b.ws(sh.H494)
        fx = b.user_edit(7, 25, "0.3")
        self.assertEqual(ws.get(7, 25), 0.5)
        self.assertTrue(fx.logs[0].text.endswith("：○"))
        # セル自体を緑に塗るが、直後の Cells(TR, 26).Activate の SelectionChange で白に戻る(VBA どおり)
        self.assertEqual(ws.interior[(7, 25)], (255, 255, 255))
        fx = b.user_edit(8, 25, "")                             # 空でも "0.0" を書いて判定
        self.assertEqual(ws.get(8, 25), 0.0)
        self.assertTrue(fx.logs[0].text.endswith("：○"))
        fx = b.user_edit(11, 25, "3.0")                         # 11 行目以降は判定しない
        self.assertEqual(ws.get(11, 25), 3.0)
        self.assertEqual(fx.logs, [])

    def test_diagonal(self) -> None:                     # Cask:1555-1573
        b = h_book()
        ws = b.ws(sh.H494)
        b.user_edit(7, 21, "#L2502.0")
        b.user_edit(7, 22, "#L2501.0")
        self.assertEqual((ws.get(7, 23), ws.get(7, 24)), ("○", 1.0))
        self.assertEqual(b.active_cell, (8, 2))

    def test_b30_log_dir(self) -> None:                  # Cask:1373-1376
        b = h_book()
        fx = b.user_edit(7, 2, "DT10002+0000010.000M")
        self.assertTrue(fx.logs[1].path.endswith("丸徳\\丸徳_Log"))
        quirks.set_for_test({"B30": "fix"})
        fx = b.user_edit(8, 2, "DT10002+0000010.000M")
        self.assertTrue(fx.logs[1].path.endswith("LVC\\LVC_Log"))


class H494PrintTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        kensa_fixture.use_offline(kensa_fixture.copy_to(Path(self.tmp.name)))

    def tearDown(self) -> None:
        self.tmp.cleanup()
        quirks.set_for_test(None)

    def test_print_footer_b30(self) -> None:             # UFMit:138-230 / Mit:922
        b = h_book()
        b.mit.kname = "検査員01"                              # OpenMitubisi は Kname に入れる
        res = h494.print_and_save(b)
        self.assertIsInstance(res, flow.Question)
        ws = b.ws(sh.H494)
        self.assertTrue(sh.cell_str(ws.get(2, 25)).startswith("ﾏｲｸﾛ:LVCM1"))
        from ufkensa.kensa.modes import zenkasho as zk
        job = zk.zen_print_setup(b, sh.H494)
        self.assertEqual(job.left_footer, '&"Meiryo UI"&8*検査員： *測定機器：')
        quirks.set_for_test({"B30": "fix"})
        job = zk.zen_print_setup(b, sh.H494)
        self.assertTrue(job.left_footer.startswith('&"Meiryo UI"&8*検査員：検査員01 *測定機器：ﾏｲｸﾛ:LVCM1'))

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
