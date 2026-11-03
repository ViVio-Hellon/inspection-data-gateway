"""履歴画面(VBA の UF管理): 手元DB の保存用・耳用の一覧と中身"""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest import mock

from tests._web import make_client
from ufkensa.kensa import store as st


class HistoryScreenTest(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        path = Path(tmp.name) / "local.sqlite3"
        s = st.LocalStore(path)
        sh = s.ensure_sheet(st.HOZON)
        sh.set(2, 1, "2026/06/03 09:47～2026/06/03 10:05")
        sh.set(2, 2, "H7372B0")
        sh.set(2, 3, "3.000")
        sh.set(2, 12, "3.012")                                   # S厚1
        s.close()
        p = mock.patch.object(st, "default_path", return_value=path)
        p.start()
        self.addCleanup(p.stop)
        self.c = make_client()
        self.h = {"X-Tool-Token": self.c.application.config["TOKEN"]}

    def test_page_list_and_view(self):
        self.assertEqual(self.c.get("/history").status_code, 200)
        lots = self.c.get("/api/history/lots", headers=self.h).get_json()
        self.assertEqual([r["lot"] for r in lots["rows"]], ["H7372B0"])
        v = self.c.get("/api/history/view?row=2", headers=self.h).get_json()
        self.assertEqual(v["lot"], "H7372B0")
        self.assertTrue(any("3.012" in r["cells"] for r in v["grid"]))
        self.assertEqual(self.c.get("/api/history/view?row=9", headers=self.h).status_code, 404)
        self.assertEqual(self.c.get("/api/history/lots?mimi=1", headers=self.h).get_json()["rows"], [])


if __name__ == "__main__":
    unittest.main()
