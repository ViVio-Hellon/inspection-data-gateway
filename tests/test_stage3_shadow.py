"""第3段7・9 の影の確かめ(shadow.py)が、食い違いを [第3段7] / [第3段9] の印つきで見つけること。"""
from __future__ import annotations

import unittest
from unittest import mock

from tests.test_kensa_session import SessionBase
from ufkensa.kensa import grid, model, records, shadow, wiring


class ShadowTest(SessionBase):
    def test_verdict_in_step_through_a_lot(self) -> None:
        self.register()
        self.fill("DT")
        st = self.s.state
        self.assertEqual(st.verdicts.get("S1"), model.Verdict.OK)
        self.assertEqual(shadow.verdict_diffs(st), [])

    def test_color_changed_without_verdict_is_reported(self) -> None:
        self.register()
        self.fill("DT")
        st = self.s.state
        st.cells["S2"].color = grid.COLOR_NG               # 判定を通さずに色だけ変える(間違いの例)
        with self.assertLogs("ufkensa.stage3", "WARNING") as cm, mock.patch.object(shadow, "STRICT", False):
            self.s.click("")
        self.assertTrue(any("[第3段7]" in x and "S2" in x for x in cm.output), cm.output)
        with self.assertRaises(shadow.ShadowMismatch):
            shadow.check_verdicts(st, "試験")

    def test_restore_difference_is_reported(self) -> None:
        r = records.FormRestore()
        r.values.update({"S1": "15.012", "Lot": "H7345B0"})
        st = grid.KensaState()
        wiring.put(st, "S1", "15.000")                      # 再生の結果が保存の値と違う(間違いの例)
        wiring.put(st, "Lot", "H7345B0")
        with self.assertLogs("ufkensa.stage3", "WARNING") as cm, mock.patch.object(shadow, "STRICT", False):
            shadow.check_restore(st, wiring.restore_typed(r), wiring._restore_names(r), "試験")
        self.assertTrue(any("[第3段9]" in x and "S1" in x for x in cm.output), cm.output)

    def test_undeclared_view_name_is_reported(self) -> None:
        st = grid.KensaState()
        st.v("KenH").value = "1"                            # 宣言済み: 何も出ない
        with self.assertRaises(shadow.ShadowMismatch):
            st.v("KenHH")                                  # 打ち間違い(STRICT の試験では失敗)
        with self.assertLogs("ufkensa.stage3", "WARNING") as cm, mock.patch.object(shadow, "STRICT", False):
            st.v("Lott")
        self.assertTrue(any("[第3段10]" in x and "Lott" in x for x in cm.output), cm.output)

    def test_off_switch(self) -> None:
        with mock.patch.object(shadow, "enabled", return_value=False):
            st = grid.KensaState()
            st.cells["S1"].color = grid.COLOR_NG
            shadow.check_verdicts(st, "試験")              # 止めているので何も起きない


if __name__ == "__main__":
    unittest.main()
