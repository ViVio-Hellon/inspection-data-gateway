"""照合ルール 第4段: 外観_H(Q_S_QA_124)・外観_S(SRF_128)・ラテラル(DIM_179)・BH(Q_S_KE_ALL_053)の特性テスト(2026-11-11)。

記録(tests/fixtures/match_rules/characterization_p4.json)は、照合ルールへ移す**前の**コードの結果。
作り直しは `python -m tests.match_rules_p4_cases --write`(結果が変わるのを承知のときだけ)。
"""
from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from tests import kensa_fixture
from tests import match_rules_p4_cases as mc


def _p4() -> tuple[str, ...]:
    from ufkensa.kensa import match_rules as mr
    return (mr.APPEAR_H, mr.APPEAR_S, mr.LATERAL_N, mr.LATERAL_H, mr.BH)


P4 = _p4()


def setUpModule() -> None:
    kensa_fixture.use_offline()


class P4CharacterizationTest(unittest.TestCase):
    """新(照合ルール。旧と shadow で並べる)・旧(コード)・表から読んだ既定 のどれでも、記録と同じ。"""

    def setUp(self) -> None:
        from ufkensa.kensa import match_rules as mr
        self._strict = mr.SHADOW_STRICT
        mr.SHADOW_STRICT = True
        mr._last_known.clear()

    def tearDown(self) -> None:
        from ufkensa.kensa import match_rules as mr
        mr.SHADOW_STRICT = self._strict
        mr._last_known.clear()

    def _check(self) -> None:
        want = json.loads(mc.RECORD.read_text(encoding="utf-8"))
        with tempfile.TemporaryDirectory() as t:
            got = mc.record(Path(t))
        for kind in mc.KINDS:
            self.assertEqual(sorted(got[kind]), sorted(want[kind]), kind)
            diff = [k for k in want[kind] if got[kind][k] != want[kind][k]]
            self.assertEqual(diff[:5], [], f"{kind}: {len(diff)} 件ちがう")

    def test_new_default_with_shadow(self) -> None:
        from ufkensa.kensa import match_rules as mr
        self.assertTrue(all(mr.BY_NAME[n].new_way for n in P4))
        self._check()

    def test_old_path(self) -> None:
        from dataclasses import replace
        from unittest import mock
        from ufkensa.kensa import match_rules as mr
        old = {n: replace(g, new_way=False) for n, g in mr.BY_NAME.items()}
        with mock.patch.object(mr, "group", lambda n: old[n]):
            self._check()

    def test_default_rules_round_trip_table(self) -> None:
        from ufkensa.kensa import match_rules as mr
        with tempfile.TemporaryDirectory() as t:
            kensa_fixture.use_offline(kensa_fixture.copy_to(Path(t)))
            try:
                self.assertTrue(mr.ensure_table())
                got = mr.groups(fresh=True)
                for n in P4:
                    self.assertEqual(got[n].rules, mr.BY_NAME[n].rules)
            finally:
                kensa_fixture.use_offline()


class P4RulesTest(unittest.TestCase):
    """ルールを変えたときの動き(製造 ⇔ オーダーの切替・右3文字・列が無ければ通す・含む(Like)・保守の試し照合)。"""

    def setUp(self) -> None:
        from ufkensa.kensa import match_rules as mr
        mr._last_known.clear()

    def _with(self, name: str, **kw):
        from dataclasses import replace
        from unittest import mock
        from ufkensa.kensa import match_rules as mr
        g = replace(mr.BY_NAME[name], **kw)
        return mock.patch.object(mr, "group", lambda n: g if n == name else mr.BY_NAME[n])

    def test_bh_right3_prep(self) -> None:
        from dataclasses import replace
        from ufkensa.kensa import match_rules as mr, quality as q
        self.assertEqual(q.bh_search("XU74", "T6"), 80)
        rules = tuple(replace(r, prep="なし") for r in mr.BY_NAME[mr.BH].rules)
        with self._with(mr.BH, rules=rules):
            self.assertEqual(q.bh_search("XU74", "T6"), 0)            # 右3文字にしないと XU74 を含む行は無い
            self.assertEqual(q.bh_search("U74", "T6"), 80)

    def test_appear_h_switch_to_order_thickness(self) -> None:
        from ufkensa.kensa import match_rules as mr, quality_view as qv
        from tests.match_rules_p4_cases import lot_with
        lot = lot_with("H8401Z0", {"製造板厚": "8", "ｵｰﾀﾞｰ板厚": "15"})
        a = qv.display_appearance_h(lot, qv.QualityForm()).grade_array
        g = mr.swap_dims(mr.BY_NAME[mr.APPEAR_H])
        self.assertEqual([r.field for r in g.rules if r.compare == "範囲"], ["ｵｰﾀﾞｰ板厚", "製造板幅", "製造板丈"])
        with self._with(mr.APPEAR_H, rules=g.rules):
            b = qv.display_appearance_h(lot, qv.QualityForm()).grade_array
        self.assertNotEqual(a, b)

    def test_missing_column_passes(self) -> None:
        from ufkensa.kensa import match_rules as mr
        r = mr.BY_NAME[mr.APPEAR_H].rules[2]
        self.assertEqual((r.columns, r.prep), (("調質",), "列が無ければ通す"))
        self.assertTrue(mr._test(r, ["1"], ["管理番号"], "H112"))
        self.assertFalse(mr._test(r, ["1", "O"], ["管理番号", "調質"], "H112"))

    def test_like_contains_and_sql(self) -> None:
        from ufkensa.kensa import match_rules as mr
        like = mr._r(1, "表面処理コード", "品質ｸﾞﾚｰﾄﾞ_表面処理", "含む(Like)")
        self.assertTrue(mr._test(like, ["B,H,K"], ["表面処理コード"], "H"))
        self.assertTrue(mr._test(like, [None], ["表面処理コード"], ""))
        sql = mr._r(1, "用途コード", "用途ｺｰﾄﾞ", "含む(SQL)")
        self.assertTrue(mr._test(sql, ["h191,H241"], ["用途コード"], "H191"))     # 大文字小文字を区別しない(ADO)
        self.assertFalse(mr._test(sql, ["H241"], ["用途コード"], "H191"))

    def test_maint_trial(self) -> None:
        from ufkensa.kensa import match_rules as mr, match_rules_maint as mm, store
        with tempfile.TemporaryDirectory() as t:
            kensa_fixture.use_offline(kensa_fixture.copy_to(Path(t)))
            try:
                d = mm.trial(mr.BH, "", None, store.LocalStore(), {"材質": "U74", "調質": "T4"})
                self.assertEqual([c["saved"]["result"] for c in d["cases"]], ["下限値 78"])
                d = mm.trial(mr.APPEAR_S, "L726M50", None, store.LocalStore())
                self.assertEqual(d["cases"][0]["saved"]["result"], "管理番号 13 → パターン1 = 4")
                d = mm.trial(mr.LATERAL_N, "H8401Z0", None, store.LocalStore())
                self.assertIn("ラテラルボー 4.0", d["cases"][0]["saved"]["result"])
                for n in (mr.APPEAR_H, mr.BH):
                    self.assertEqual(mm.regression(n, None, store.LocalStore())["changed_count"], 0)
                st = {g["name"]: g for g in mm.state()["groups"]}
                self.assertEqual(st[mr.BH]["field_sources"]["材質"], "画面")
                self.assertEqual(st[mr.APPEAR_H]["field_sources"]["品質ｸﾞﾚｰﾄﾞ_表面処理"], "仕掛")
                self.assertEqual(st[mr.APPEAR_H]["dim_swap"][0], ["製造板厚", "ｵｰﾀﾞｰ板厚"])
            finally:
                kensa_fixture.use_offline()


if __name__ == "__main__":
    unittest.main()
