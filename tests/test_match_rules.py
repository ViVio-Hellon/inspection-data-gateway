"""照合ルール(ufkensa/kensa/match_rules.py)の読み書き・確かめ・照合の途中経過。"""
from __future__ import annotations

import sqlite3
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path

from tests import kensa_fixture
from ufkensa.kensa import match_rules as mr
from ufkensa.kensa import quality as q


class MatchRulesTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.db = kensa_fixture.copy_to(Path(self.tmp.name))
        kensa_fixture.use_offline(self.db)
        mr._last_known.clear()

    def tearDown(self) -> None:
        kensa_fixture.use_offline()
        mr._last_known.clear()
        self.tmp.cleanup()

    def test_no_table_uses_defaults_without_writing(self) -> None:
        self.assertEqual(mr.group(mr.RATEFURI), mr.BY_NAME[mr.RATEFURI])
        con = sqlite3.connect(mr._db_path())
        self.assertIsNone(con.execute("SELECT 1 FROM sqlite_master WHERE name = ?", (mr.TABLE,)).fetchone())
        con.close()

    def test_save_roundtrip_and_history(self) -> None:
        g = mr.BY_NAME[mr.RATEFURI]
        rules = tuple(replace(r, field="製造板厚") if r.order == 60 else r for r in g.rules)
        mr.save(replace(g, rules=rules), reason="試し")
        got = mr.group(mr.RATEFURI)
        self.assertEqual(got.rules, rules)
        h = mr.history()
        self.assertTrue(h and h[0]["group"] == mr.RATEFURI and h[0]["reason"] == "試し")

    def test_check_rejects(self) -> None:
        g = mr.BY_NAME[mr.RATEFURI]
        bad = [replace(g, rules=(replace(g.rules[0], field="無い項目"),)),
               replace(g, rules=(replace(g.rules[0], compare="なにか"),)),
               replace(g, rules=(g.rules[0], g.rules[0])),
               replace(g, rules=(replace(g.rules[5], columns=("厚下",)),)),
               replace(g, use_rows="真ん中")]
        for b in bad:
            with self.assertRaises(mr.Invalid):
                mr.check(b)
        with self.assertRaises(mr.Invalid):
            mr.check(replace(g, rules=(replace(g.rules[0], columns=("無い列",)),)), ["納入先"])

    def test_disable_rule_changes_result_and_trace(self) -> None:
        args = ("ﾆﾂﾎﾟﾝﾌﾙﾊ-ﾌ(ｶ", "", "F975", "52S", "H18", q.vbx.single(0.6), 1217.0, 2438.0)
        self.assertEqual(q.db_return_item(*args).items, ",フリー反り")
        wrong = args[:3] + ("ZZZ",) + args[4:]
        self.assertEqual(q.db_return_item(*wrong).items, "")
        g = mr.BY_NAME[mr.RATEFURI]
        mr.save(replace(g, rules=tuple(replace(r, enabled=r.order != 40) for r in g.rules)))
        self.assertEqual(q.db_return_item(*wrong).items, ",フリー反り")
        db = q._quality_db()
        res = mr.match(mr.group(mr.RATEFURI), q.datasource.get_records(db, "ラテフリ直角"),
                       q.datasource.get_fields(db, "ラテフリ直角"), q.ratefuri_ctx(wrong))
        self.assertEqual(len(res.steps), 7)
        self.assertTrue(all(len(a.remain) >= len(b.remain) for a, b in zip(res.steps, res.steps[1:])))


if __name__ == "__main__":
    unittest.main()


from tests.test_screen_maint import MaintRouteTest  # noqa: E402


class MatchRulesRouteTest(MaintRouteTest):
    def setUp(self) -> None:
        super().setUp()
        from app.routes import match_rules as routes
        app = self.client.application
        if "match_rules" not in app.blueprints:
            app.register_blueprint(routes.bp)
        mr._last_known.clear()
        self.addCleanup(mr._last_known.clear)

    def test_api(self) -> None:
        d = self.client.get("/api/maint/match_rules", headers=self.h).get_json()
        g = {x["name"]: x for x in d["groups"]}[mr.RATEFURI]
        self.assertIn("厚下", g["table_fields"])
        r = self.post("/api/maint/match_rules/trial", {"group": mr.RATEFURI, "lot": "N8061T0"}).get_json()
        self.assertIn("cases", r)
        edited = dict(g)
        edited["rules"] = [dict(x, enabled=x["order"] != 60) for x in g["rules"]]
        reg = self.post("/api/maint/match_rules/regression", {"group": mr.RATEFURI, "edited": edited}).get_json()
        self.assertGreater(reg["checked"], 0)
        bad = dict(g, rules=[dict(g["rules"][0], columns=["無い列"])])
        self.assertEqual(self.post("/api/maint/match_rules/trial",
                                   {"group": mr.RATEFURI, "lot": "N8061T0", "edited": bad}).status_code, 400)
        r = self.post("/api/maint/match_rules", {"group": mr.RATEFURI, "edited": edited})
        if reg["changed_count"]:
            self.assertEqual(r.get_json()["error"]["code"], "need_reason")
        r = self.post("/api/maint/match_rules", {"group": mr.RATEFURI, "edited": edited, "reason": "試し"})
        self.assertEqual(r.status_code, 200)
        from ufkensa import auth
        auth.lock_master()
        r = self.post("/api/maint/match_rules", {"group": mr.RATEFURI, "edited": g, "reason": "x"})
        self.assertEqual(r.status_code, 403)
        self.assertIn('data-tab="rules"', self.client.get("/maint", headers=self.h).get_data(as_text=True))


for _name in dir(MaintRouteTest):
    if _name.startswith("test_"):
        setattr(MatchRulesRouteTest, _name, None)
del MaintRouteTest
