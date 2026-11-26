"""帳票の見本(保守 → 帳票の見本。2026-11-26): 全部の組み合わせで 測定画面 と 帳票 が作れること。"""
from __future__ import annotations

import unittest

from tests.test_screen_maint import MaintRouteTest
from ufkensa.kensa import print_samples as ps


class PrintSamplesTest(unittest.TestCase):
    def test_every_pattern_makes_screen_and_print(self) -> None:
        pats = ps.all_patterns()
        self.assertEqual(len(pats), ps.options()["total"])
        bad, layouts = [], set()
        for p in pats:
            s = ps.make(p)
            if s.problems or not s.html or s.filled == 0:
                bad.append((p.label(), s.problems))
            layouts.add((p.work, s.filled))
        self.assertEqual(bad, [])
        self.assertGreater(len(layouts), 20)                      # 組み合わせで欄の並びが変わる

    def test_options_change_screen(self) -> None:
        base = ps.make(ps.Pattern(work="N(薄板)", yama="1"))
        lat = ps.make(ps.Pattern(work="N(薄板)", yama="1", lateral=True))
        self.assertFalse(any(c.get("visible") for c in base.screen["grid"]["lateral"]))
        self.assertTrue(any(c.get("visible") for c in lat.screen["grid"]["lateral"]))
        free = ps.make(ps.Pattern(work="H(厚板)", free=True))
        self.assertEqual(free.screen["tolerance"]["flatnes_text"], "ﾌﾘｰｿﾘ要")
        self.assertIn("SAMPLE0", base.html)
        self.assertIn("500.0", ps.make(ps.Pattern(work="N(薄板)", yama="2")).html)   # 検入高さ 2 山ぶん

    def test_bad_work(self) -> None:
        with self.assertRaises(ValueError):
            ps.Pattern.from_dict({"work": "無い作業"})


class PrintSamplesRouteTest(MaintRouteTest):
    def test_routes(self) -> None:
        d = self.client.get("/api/maint/samples/options", headers=self.h).get_json()
        self.assertIn("N(薄板)", d["works"])
        r = self.post("/api/maint/samples/make", {"work": "スラブ", "right": True})
        self.assertEqual(r.status_code, 200)
        j = r.get_json()
        self.assertEqual(j["problems"], [])
        self.assertIn("<table>", j["html"])
        self.assertTrue(j["screen"]["grid"]["rows"])
        self.assertEqual(self.post("/api/maint/samples/make", {"work": "x"}).status_code, 400)
