"""2026-11-18: ロットの厚が読めない(空・0)ときも Min/Max が止まらない(前は '' と 0 の比較で TypeMismatch)。"""
from __future__ import annotations

import unittest
from unittest import mock

from ufkensa.kensa import minmax
from ufkensa.kensa.grid import KensaState


class EmptyAtuTest(unittest.TestCase):
    def test_minmax_with_empty_thickness(self) -> None:
        for atu in ("", "0"):
            st = KensaState(lot="LOT0001", work="厚板")
            st.atu = atu
            st.ops, st.ope = 1, 2
            st.cells["S1"].value, st.cells["U1"].value = "174.85", "174.37"
            st.cells["S5"].value, st.cells["S6"].value = "653.3", "653.0"
            hooks = mock.Mock()
            with self.subTest(atu=atu):
                self.assertEqual(minmax.min_max(st, hooks), [])
                self.assertEqual((st.amin, st.amax), ("174.370", "174.850"))
                self.assertEqual((st.hmin, st.hmax), ("653.0", "653.3"))
                self.assertEqual((st.tmin, st.tmax), ("", ""))         # 丈は無いので空(止まらない)


if __name__ == "__main__":
    unittest.main()
