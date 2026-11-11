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


def setUpModule() -> None:
    kensa_fixture.use_offline()


class P4CharacterizationTest(unittest.TestCase):
    def _check(self) -> None:
        want = json.loads(mc.RECORD.read_text(encoding="utf-8"))
        with tempfile.TemporaryDirectory() as t:
            got = mc.record(Path(t))
        for kind in mc.KINDS:
            self.assertEqual(sorted(got[kind]), sorted(want[kind]), kind)
            diff = [k for k in want[kind] if got[kind][k] != want[kind][k]]
            self.assertEqual(diff[:5], [], f"{kind}: {len(diff)} 件ちがう")

    def test_current_code(self) -> None:
        self._check()


if __name__ == "__main__":
    unittest.main()
