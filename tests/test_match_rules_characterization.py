"""照合ルール 第1段の特性テスト(2026-11-11)。ラテフリ直角・梱包保護材の照合が、記録した結果と同じか。

記録(tests/fixtures/match_rules/characterization.json)は、照合ルールへ移す**前の**コードの結果。
作り直しは `python -m tests.match_rules_cases --write`(結果が変わるのを承知のときだけ)。
"""
from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from tests import kensa_fixture
from tests import match_rules_cases as mc


def setUpModule() -> None:
    kensa_fixture.use_offline()


class CharacterizationTest(unittest.TestCase):
    """新(照合ルール)・旧(コード)・表に入れた既定のルール のどれでも、記録と同じ。新旧は shadow でも並べる。"""

    def setUp(self) -> None:
        from ufkensa.kensa import match_rules as mr
        self._strict = mr.SHADOW_STRICT
        mr.SHADOW_STRICT = True                   # 新旧を並べ、違えば ShadowMismatch
        mr._last_known.clear()

    def tearDown(self) -> None:
        from ufkensa.kensa import match_rules as mr
        mr.SHADOW_STRICT = self._strict
        mr._last_known.clear()

    def _check(self) -> None:
        want = json.loads(mc.RECORD.read_text(encoding="utf-8"))
        with tempfile.TemporaryDirectory() as t:
            got = mc.record(Path(t))
        for kind in ("ratefuri", "pack"):
            self.assertEqual(sorted(got[kind]), sorted(want[kind]), kind)
            diff = [k for k in want[kind] if got[kind][k] != want[kind][k]]
            self.assertEqual(diff[:5], [], f"{kind}: {len(diff)} 件ちがう")

    def test_new_default(self) -> None:
        self._check()

    def test_old_path(self) -> None:
        from unittest import mock
        from dataclasses import replace
        from ufkensa.kensa import match_rules as mr
        old = {n: replace(g, new_way=False) for n, g in mr.BY_NAME.items()}
        with mock.patch.object(mr, "group", lambda n: old[n]):
            self._check()

    def test_rules_from_table(self) -> None:
        """表「照合ルール」を作って(既定を入れて)、そこから読んだルールでも同じ。"""
        from ufkensa.kensa import match_rules as mr
        with tempfile.TemporaryDirectory() as t:
            db = kensa_fixture.copy_to(Path(t))
            kensa_fixture.use_offline(db)
            try:
                self.assertTrue(mr.ensure_table())
                got = mr.groups(fresh=True)
                self.assertEqual({n: g.rules for n, g in got.items()}, {n: g.rules for n, g in mr.BY_NAME.items()})
                self._check()
            finally:
                kensa_fixture.use_offline()


if __name__ == "__main__":
    unittest.main()
