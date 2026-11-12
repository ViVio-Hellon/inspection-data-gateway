"""2026-11-12: ラテラルボー条件に当たったことを画面へ渡す(view.lateral_hit)。"""
from __future__ import annotations

import unittest

from ufkensa.kensa import view, wiring
from ufkensa.kensa.grid import KensaState


class LateralHitTest(unittest.TestCase):
    def test_no_hit(self) -> None:
        h = view.build(KensaState())["lateral_hit"]
        self.assertEqual(h, {"on": False, "who": "", "text": "", "source": []})

    def test_sumisei_and_dim179_table(self) -> None:
        st = KensaState()
        wiring.apply_values(st, {"Sumisei.Caption": "ナメカワ", "LateralText": "3.0"})
        wiring.put(st, "LateralText.BackColor", 255 + 153 * 256)       # quality.ORANGE = DIM_179 の表から
        h = view.lateral_hit(st)
        self.assertTrue(h["on"])
        self.assertEqual(h["who"], "ナメカワ")
        self.assertEqual(h["text"], "3.0")
        self.assertIn("DIM_179 日車/白銅滑川の表", h["source"])
        self.assertIn("製品規格のラテラルボー要", h["source"])

    def test_gijutsu_shiji(self) -> None:
        st = KensaState()
        st.lateral_text, st.lateral_text_tag = "2.5", "有"
        h = view.lateral_hit(st)
        self.assertTrue(h["on"])
        self.assertEqual(h["source"], ["技術指示"])


if __name__ == "__main__":
    unittest.main()

