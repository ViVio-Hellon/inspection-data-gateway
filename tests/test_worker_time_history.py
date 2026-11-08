"""履歴画面で、測定値ごとの検査員と時刻を見る(印刷・保存のときに控える。2026-11-08)。"""
from __future__ import annotations

from unittest import mock

from tests.test_kensa_session import LOT, SessionBase
from ufkensa.kensa import session, worker_archive


class WorkerTimeHistoryTest(SessionBase):
    def print_lot(self) -> None:
        s = self.s
        self.register()
        self.fill("DT")
        self.fill("#L")
        self.strain_visual_height()
        self.ok(s.min_max())
        self.ok(s.print_check())
        s.print_commit()
        self.ok(s.print_commit({"print": True}))

    def test_archive_and_history_view(self) -> None:
        self.print_lot()
        from tests import _web
        from app.routes import history as hroute
        from ufkensa.kensa import store as stmod
        session.reset_for_test(self.s)
        self.addCleanup(session.reset_for_test, None)
        p = mock.patch.object(hroute, "_store", lambda: stmod.LocalStore(self.store.path))
        p.start()
        self.addCleanup(p.stop)
        c, h = _web.make_client(), _web.auth()
        rows = c.get("/api/history/lots", headers=h).get_json()["rows"]
        row = next(r["row"] for r in rows if r["lot"] == LOT)
        v = c.get(f"/api/history/view?row={row}", headers=h).get_json()
        self.assertEqual(v["worker_time"]["S1"]["who"], "検査員01")
        self.assertTrue(v["worker_time"]["S5"]["when"].startswith("2026/11/02"))   # 幅も
        self.assertIn('id="histWho"', c.get("/history", headers=h).get_data(as_text=True))

    def test_new_lot_keeps_archive(self) -> None:
        self.print_lot()
        period = self.store.sheet("保存用").get(2, 1)
        before = worker_archive.lookup(self.store, str(period), LOT)
        self.assertTrue(before)
        # 新規検査で WorkerTime は作り直されても、控えは残る
        self.ok(self.s.new_inspection({"new": True}) if hasattr(self.s, "new_inspection") else self.s.start())
        self.assertEqual(worker_archive.lookup(self.store, str(period), LOT), before)

    def test_reprint_replaces(self) -> None:
        self.print_lot()
        period = str(self.store.sheet("保存用").get(2, 1))
        n = len(worker_archive.lookup(self.store, period, LOT))
        worker_archive.archive(self.store, "HVC", LOT, period)                     # もう一度写しても増えない
        self.assertEqual(len(worker_archive.lookup(self.store, period, LOT)), n)
