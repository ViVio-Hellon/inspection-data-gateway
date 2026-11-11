"""照合ルール 第3段: 測定機器ID(DbReturnMicro)・設備仕様(CheckEquipmentDb)の特性テスト(2026-11-11)。

記録(tests/fixtures/match_rules/characterization_equip.json)は、照合ルールへ移す**前の**コードの結果。
作り直しは `python -m tests.match_rules_equip_cases --write`(結果が変わるのを承知のときだけ)。
"""
from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from tests import kensa_fixture
from tests import match_rules_equip_cases as mc


def setUpModule() -> None:
    kensa_fixture.use_offline()


class EquipCharacterizationTest(unittest.TestCase):
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
        for kind in ("micro", "equip"):
            self.assertEqual(sorted(got[kind]), sorted(want[kind]), kind)
            diff = [k for k in want[kind] if got[kind][k] != want[kind][k]]
            self.assertEqual(diff[:5], [], f"{kind}: {len(diff)} 件ちがう")

    def test_new_default_with_shadow(self) -> None:
        from ufkensa.kensa import match_rules as mr
        self.assertTrue(mr.BY_NAME[mr.MICRO].new_way and mr.BY_NAME[mr.EQUIP].new_way)
        self._check()

    def test_old_path(self) -> None:
        from dataclasses import replace
        from unittest import mock
        from ufkensa.kensa import match_rules as mr
        old = {n: replace(g, new_way=False) for n, g in mr.BY_NAME.items()}
        with mock.patch.object(mr, "group", lambda n: old[n]):
            self._check()

    def test_default_rules_round_trip_table(self) -> None:
        """表に入れた既定を読み戻すと既定と同じ(表から読んだルールで照合しても同じ)。"""
        from ufkensa.kensa import match_rules as mr
        with tempfile.TemporaryDirectory() as t:
            kensa_fixture.use_offline(kensa_fixture.copy_to(Path(t)))
            try:
                self.assertTrue(mr.ensure_table())
                got = mr.groups(fresh=True)
                for n in (mr.MICRO, mr.EQUIP):
                    self.assertEqual(got[n].rules, mr.BY_NAME[n].rules)
            finally:
                kensa_fixture.use_offline()



class EquipRulesTest(unittest.TestCase):
    """ルールを変えたときの動き(含む条件式・用途の引き直し・B52 の前処理・設備仕様の文言・保守の試し照合)。"""

    def setUp(self) -> None:
        from ufkensa.kensa import match_rules as mr
        mr._last_known.clear()

    def _with(self, name: str, **kw):
        from dataclasses import replace
        from unittest import mock
        from ufkensa.kensa import match_rules as mr
        g = replace(mr.BY_NAME[name], **kw)
        return mock.patch.object(mr, "group", lambda n: g if n == name else mr.BY_NAME[n])

    def test_contains_means_contains(self) -> None:
        from ufkensa.kensa import instruments
        self.assertTrue(instruments.youto_match("A5*", "BA5"))      # X* も「含む」(前方一致ではない)
        self.assertTrue(instruments.youto_match("*K2", "AK2B"))
        self.assertTrue(instruments.youto_match("≠H191", "H207"))
        self.assertFalse(instruments.youto_match("H191", "H1911"))
        self.assertFalse(instruments.youto_match("", "H191"))

    def test_youto_rule_and_fallback(self) -> None:
        from ufkensa.kensa import instruments
        from ufkensa.kensa import match_rules as mr
        self.assertEqual(instruments.db_return_micro("HVC", 10.0, "H191"), "10002")      # 用途で M4
        self.assertEqual(instruments.db_return_micro("HVC", 10.0, "Z"), "10001")         # 合わないので厚だけ → M1
        g = mr.BY_NAME[mr.MICRO]
        off = tuple(r if r.compare != "含む条件式" else __import__("dataclasses").replace(r, enabled=False)
                    for r in g.rules)
        with self._with(mr.MICRO, rules=off):
            self.assertEqual(instruments.db_return_micro("HVC", 10.0, "H191"), "10001")

    def test_b52_single_prep(self) -> None:
        from dataclasses import replace
        from ufkensa.kensa import instruments
        from ufkensa.kensa import match_rules as mr
        self.assertEqual(instruments.db_return_micro("HVC", 28.01, " ", "機器名"), "M2")
        rules = tuple(replace(r, prep="Single丸め") if r.compare == "範囲" else r for r in mr.BY_NAME[mr.MICRO].rules)
        with self._with(mr.MICRO, rules=rules):
            self.assertEqual(instruments.db_return_micro("HVC", 28.01, " ", "機器名"), "")   # B52 の再現

    def test_equipment_messages(self) -> None:
        from dataclasses import replace
        from ufkensa.kensa import match_rules as mr
        from ufkensa.kensa import quality as q
        self.assertEqual(q.check_equipment_db("LVC", "5052", 25.0, 1000.0, 2000.0, 10.0), "オーダー厚 範囲外")
        self.assertEqual(q.check_equipment_db("無いライン", "5052", 25.0, 1000.0, 2000.0, 10.0), "可")
        rules = tuple(replace(r, enabled=r.order != 30) for r in mr.BY_NAME[mr.EQUIP].rules)
        with self._with(mr.EQUIP, rules=rules):
            self.assertEqual(q.check_equipment_db("LVC", "5052", 25.0, 1000.0, 2000.0, 10.0), "可")

    def test_maint_trial_direct_inputs(self) -> None:
        from ufkensa.kensa import match_rules as mr, match_rules_maint as mm, store
        with tempfile.TemporaryDirectory() as t:
            kensa_fixture.use_offline(kensa_fixture.copy_to(Path(t)))
            try:
                d = mm.trial(mr.MICRO, "", None, store.LocalStore(), {"ライン": "HVC", "厚": "10", "用途": "H191"})
                self.assertEqual([c["saved"]["result"] for c in d["cases"]], ["M4(出力テキスト 10002)"])
                st = {g["name"]: g for g in mm.state()["groups"]}
                self.assertEqual(st[mr.EQUIP]["field_sources"]["製造材質"], "仕掛")
                self.assertEqual(st[mr.MICRO]["inputs"], ["ライン", "厚", "用途"])
            finally:
                kensa_fixture.use_offline()


if __name__ == "__main__":
    unittest.main()
