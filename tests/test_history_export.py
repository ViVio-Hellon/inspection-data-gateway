"""保守 → 履歴出力(テスト履歴を入れる・PDF 出力・CSV 出力)。"""
from __future__ import annotations

import unittest
import zipfile
from pathlib import Path

from tests.test_kensa_session import SessionBase
from ufkensa.kensa import history_export as hx, records
from ufkensa.kensa.store import HOZON, LOT_HISTORY


class HistoryExportTest(SessionBase):
    def test_load_test_history_then_export(self) -> None:
        r = hx.load_test_history(self.store, "HVC")
        self.assertEqual(r["lot"], "T999T99")
        self.assertIn("保存用: 1 行", r["added"])
        self.assertEqual(hx.load_test_history(self.store, "HVC")["added"], [])        # 2 回目は足さない
        hz = self.store.sheet(HOZON)
        row = next(n for n, v in hz.rows() if records.cell_text(v[1]) == "T999T99")

        out = hx.export_pdf(self.store, row, session_factory=lambda st: self.new_session_on(st))
        self.assertEqual(out["lot"], "T999T99")
        self.assertTrue(any(p.endswith("T999T99.html") for p in out["pdf"]), out)      # Edge が無いので HTML
        book = Path(out["xlsx"][0])
        self.assertEqual(book.name, "20261102_T999T99.xlsx")
        with zipfile.ZipFile(book) as z:
            self.assertIn("印刷用_", z.read("xl/workbook.xml").decode("utf-8"))
        # 本物の手元DB は触らない(印刷用・測定DB の登録なし)
        self.assertFalse(self.store.exists("印刷用_HVC"))
        self.assertEqual(len([1 for _, v in hz.rows() if records.cell_text(v[1]) == "T999T99"]), 1)

        c = hx.export_csv(self.store, "HVC", "HVC")
        names = [Path(p).name for p in c["files"]]
        self.assertIn("HVC_履歴CSV.csv", names)
        self.assertIn("Lot情報履歴_HVC.csv", names)
        self.assertTrue(self.store.exists(LOT_HISTORY))

        self.assertTrue(hx.has_test_history(self.store))
        self.assertEqual(hx.remove_test_history(self.store)["removed"], ["保存用: 1 行", "Lot情報履歴: 1 行"])
        self.assertFalse(hx.has_test_history(self.store))

    def test_routes(self) -> None:
        from unittest import mock
        from tests import _web
        from app.routes import maint as mroute, history as hroute
        from ufkensa.kensa import session, store as stmod
        session.reset_for_test(self.s)
        self.addCleanup(session.reset_for_test, None)
        opener = lambda: stmod.LocalStore(self.store.path)                       # noqa: E731
        for mod in (mroute, hroute):
            p = mock.patch.object(mod, "_store", opener)
            p.start()
            self.addCleanup(p.stop)
        c, h = _web.make_client(), _web.auth()
        self.assertIn('data-panel="hist"', c.get("/maint", headers=h).get_data(as_text=True))
        self.assertFalse(c.get("/api/maint/history/state", headers=h).get_json()["test_loaded"])
        r = c.post("/api/maint/history/test", json={"add": True}, headers=h)
        self.assertEqual(r.status_code, 200, r.get_json())
        self.assertTrue(c.get("/api/maint/history/state", headers=h).get_json()["test_loaded"])
        rows = c.get("/api/history/lots", headers=h).get_json()["rows"]
        row = next(x["row"] for x in rows if x["lot"] == "T999T99")
        r = c.post("/api/maint/history/pdf", json={"row": row}, headers=h)
        self.assertEqual(r.status_code, 200, r.get_json())
        self.assertTrue(r.get_json()["xlsx"])
        r = c.post("/api/maint/history/csv", json={}, headers=h)
        self.assertEqual(r.status_code, 200, r.get_json())
        self.assertTrue(any(f.endswith("Lot情報履歴_HVC.csv") for f in r.get_json()["files"]))
        self.assertEqual(c.post("/api/maint/history/pdf", json={"row": 0}, headers=h).status_code, 400)
        r = c.post("/api/maint/history/test", json={"add": False}, headers=h)
        self.assertEqual(r.get_json()["removed"], ["保存用: 1 行", "Lot情報履歴: 1 行"])

    def new_session_on(self, st):
        from ufkensa.kensa import session
        return session.Session(st, now=self.clock, calibration_skip=lambda: True,
                               measure_db=self.tmp / "unused.sqlite3")


if __name__ == "__main__":
    unittest.main()
