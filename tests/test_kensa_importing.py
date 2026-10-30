"""取り込み(Access → SQLite。D4)の一覧と権限"""
from __future__ import annotations

import os
import unittest
from unittest import mock

from tests._web import make_client
from ufkensa import auth
from ufkensa.kensa import importing, paths


class PlanTest(unittest.TestCase):
    def test_targets_and_default_source(self):
        rows = {r.name: r for r in importing.plan()}
        self.assertEqual(set(rows), set(importing.TARGETS))
        r = rows["品質基準"]
        self.assertTrue(r.target.endswith("品質基準.sqlite3"))
        self.assertTrue(r.source.endswith("品質基準.accdb"))
        self.assertEqual(r.source[: -len(".accdb")], r.target[: -len(".sqlite3")])   # 同じ場所・同じ名前

    def test_unknown_name(self):
        with self.assertRaises(KeyError):
            importing.run("無いDB")


class ImportApiTest(unittest.TestCase):
    def setUp(self):
        env = mock.patch.dict(os.environ, {"USERNAME": "someone", "COMPUTERNAME": "PC-1"})
        env.start()
        self.addCleanup(env.stop)
        auth.end()
        self.addCleanup(auth.end)
        self.c = make_client()
        self.h = {"X-Tool-Token": self.c.application.config["TOKEN"]}

    def test_page_and_plan(self):
        self.assertEqual(self.c.get("/import").status_code, 200)
        p = self.c.get("/api/import/plan", headers=self.h).get_json()
        self.assertEqual(len(p["rows"]), len(importing.TARGETS))
        if os.name != "nt":
            self.assertFalse(p["can_run"])
            r = self.c.post("/api/import/run", json={"names": ["品質基準"]}, headers=self.h)
            self.assertIn(r.status_code, (400, 403))


if __name__ == "__main__":
    unittest.main()
