"""照合ルール 第2段: 平坦度一覧QP追加(QP.DBQP の照合 15 条件)の特性テストと、ルールを変えたときの動き(2026-11-11)。

記録(tests/fixtures/match_rules/characterization_qp.json)は、照合ルールへ移す**前の**コードの結果。
作り直しは `python -m tests.match_rules_qp_cases --write`(結果が変わるのを承知のときだけ)。
"""
from __future__ import annotations

import json
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path
from unittest import mock

from tests import kensa_fixture
from tests import match_rules_qp_cases as mc
from ufkensa.kensa import match_rules as mr
from ufkensa.kensa import quality as q


def setUpModule() -> None:
    kensa_fixture.use_offline()


class QpCharacterizationTest(unittest.TestCase):
    """新(照合ルール)・旧(コード)・表から読んだ既定 のどれでも、記録と同じ(trace も)。新旧は shadow でも並べる。"""

    def setUp(self) -> None:
        self._strict = mr.SHADOW_STRICT
        mr.SHADOW_STRICT = True
        mr._last_known.clear()

    def tearDown(self) -> None:
        mr.SHADOW_STRICT = self._strict
        mr._last_known.clear()

    def _check(self) -> None:
        want = json.loads(mc.RECORD.read_text(encoding="utf-8"))
        got = mc.record()
        self.assertEqual(sorted(got), sorted(want))
        diff = [k for k in want if got[k] != want[k]]
        self.assertEqual(diff[:5], [], f"{len(diff)} 件ちがう")

    def test_new_default(self) -> None:
        self.assertTrue(mr.BY_NAME[mr.QP].new_way)
        self._check()

    def test_old_path(self) -> None:
        old = {n: replace(g, new_way=False) for n, g in mr.BY_NAME.items()}
        with mock.patch.object(mr, "group", lambda n: old[n]):
            self._check()

    def test_rules_from_table(self) -> None:
        with tempfile.TemporaryDirectory() as t:
            db = kensa_fixture.copy_to(Path(t))
            kensa_fixture.use_offline(db)
            try:
                self.assertTrue(mr.ensure_table())
                self.assertEqual(mr.groups(fresh=True)[mr.QP].rules, mr.BY_NAME[mr.QP].rules)
                self._check()
            finally:
                kensa_fixture.use_offline()


class QpRulesTest(unittest.TestCase):
    LOT = "H6083T1"

    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        kensa_fixture.use_offline(kensa_fixture.copy_to(Path(self.tmp.name)))
        mr._last_known.clear()

    def tearDown(self) -> None:
        kensa_fixture.use_offline()
        mr._last_known.clear()
        self.tmp.cleanup()

    def _run(self):
        return q.dbqp(self.LOT, "H6X1922", "ｶ) ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ", "ﾆﾂﾃﾂﾌﾞﾂｻﾝ(ｶ)ｷﾉｳﾏﾃﾘｱﾙﾌﾞ", q.read_lot(self.LOT),
                      execution="Test")

    def test_trace_carries_rule_and_version(self) -> None:
        r = self._run()
        steps = [t for t in r.trace if t["step"].startswith("条件 ")]
        self.assertEqual(len(steps), 15)
        self.assertEqual(steps[0]["step"], "条件 1: 用途コード")
        self.assertEqual(steps[0]["rule"], {"group": mr.QP, "order": 10, "version": 0, "new_way": True})
        g = mr.BY_NAME[mr.QP]
        mr.save(replace(g, rules=tuple(r_ for r_ in g.rules if r_.order != 10)), reason="試し")
        r2 = self._run()
        steps = [t for t in r2.trace if t["step"].startswith("条件 ")]
        self.assertEqual(len(steps), 14)
        self.assertEqual(steps[0]["step"], "条件 1: 製造材質")
        self.assertEqual((steps[0]["rule"]["order"], steps[0]["rule"]["version"]), (20, 1))
        self.assertEqual(r2.rule_version, 1)

    def test_swap_dims_uses_manufacturing(self) -> None:
        """製造 ⇔ オーダー の切り替え: 範囲のルールがロットの製造寸法を見るようになる。"""
        g = mr.swap_dims(mr.BY_NAME[mr.QP])
        self.assertEqual([r.field for r in g.rules if r.order in (60, 70, 80)], ["製造板厚", "製造板幅", "製造板丈"])
        lot = q.read_lot(self.LOT)
        ctx = q.qp_ctx(lot, "a", "b")
        db = q._quality_db()
        rows, fields = q.datasource.get_records(db, mr.QP), q.datasource.get_fields(db, mr.QP)
        a, b = mr.match(mr.BY_NAME[mr.QP], rows, fields, ctx), mr.match(g, rows, fields, dict(ctx, 製造板厚=99999.0))
        self.assertTrue(a.steps[5].remain)
        self.assertEqual(b.steps[5].remain, [])             # 製造板厚 99999 はどの行の範囲にも入らない


from tests.test_screen_maint import MaintRouteTest  # noqa: E402


class QpMaintRouteTest(MaintRouteTest):
    def setUp(self) -> None:
        super().setUp()
        from app.routes import match_rules as routes
        app = self.client.application
        if "match_rules" not in app.blueprints:
            app.register_blueprint(routes.bp)
        mr._last_known.clear()
        self.addCleanup(mr._last_known.clear)

    def test_qp_group_in_maint(self) -> None:
        d = self.client.get("/api/maint/match_rules", headers=self.h).get_json()
        g = {x["name"]: x for x in d["groups"]}[mr.QP]
        self.assertEqual(len(g["rules"]), 15)
        self.assertIn("検査番号の頭", g["table_fields"])
        self.assertIn(["製造板厚", "ｵｰﾀﾞｰ板厚"], g["dim_swap"])
        r = self.post("/api/maint/match_rules/trial", {"group": mr.QP, "lot": "H6083T1"}).get_json()
        self.assertTrue(r["cases"], r)
        self.assertTrue(r["cases"][0]["saved"]["result"].startswith("管理番号"), r["cases"][0]["saved"])
        self.assertFalse(any(c["changed"] for c in r["cases"]))
        edited = dict(g, rules=[dict(x, field={"ｵｰﾀﾞｰ板厚": "製造板厚"}.get(x["field"], x["field"])) for x in g["rules"]])
        reg = self.post("/api/maint/match_rules/regression", {"group": mr.QP, "edited": edited}).get_json()
        self.assertGreater(reg["checked"], 0)
        same = self.post("/api/maint/match_rules/regression", {"group": mr.QP, "edited": g}).get_json()
        self.assertEqual(same["changed_count"], 0)


for _name in dir(MaintRouteTest):
    if _name.startswith("test_"):
        setattr(QpMaintRouteTest, _name, None)
del MaintRouteTest


if __name__ == "__main__":
    unittest.main()
