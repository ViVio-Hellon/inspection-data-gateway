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

    def _check(self) -> None:
        want = json.loads(mc.RECORD.read_text(encoding="utf-8"))
        with tempfile.TemporaryDirectory() as t:
            got = mc.record(Path(t))
        for kind in ("micro", "equip"):
            self.assertEqual(sorted(got[kind]), sorted(want[kind]), kind)
            diff = [k for k in want[kind] if got[kind][k] != want[kind][k]]
            self.assertEqual(diff[:5], [], f"{kind}: {len(diff)} 件ちがう")

    def test_current(self) -> None:
        self._check()


if __name__ == "__main__":
    unittest.main()
