"""測定値ごとの検査員と時刻(WorkerTime_<L>。VBA の Time_Posting と同じシート)を画面で見る(2026-11-08)。"""
from __future__ import annotations

from tests.test_kensa_session import LOT, SessionBase


class WorkerTimeViewTest(SessionBase):
    def test_each_value_has_worker_and_time(self) -> None:
        self.register()
        self.assertEqual(self.s._view()["worker_time"], {})               # まだ測っていない
        self.fill("DT")
        wt = self.s._view()["worker_time"]
        self.assertEqual(wt["S1"]["who"], "検査員01")
        self.assertTrue(wt["S1"]["when"].startswith("2026/11/02 10:"))
        self.assertIn("S4", wt)
        self.assertNotIn("S5", wt)                                        # 幅はまだ
        # VBA と同じシート・同じ位置(行 3 = S の検査員、行 9 = S の時刻、列 = 欄の行番号)
        sh = self.store.sheet("WorkerTime_HVC")
        self.assertEqual((sh.get(1, 1), sh.get(3, 1)), (LOT, "検査員01"))
        # 消すと消える(Time_Posting temp="")
        r = self.s.dblclick("S1")
        self.assertNotIn("S1", r.view["worker_time"])

    def test_route_and_page(self) -> None:
        from tests import _web
        from ufkensa.kensa import session
        session.reset_for_test(self.s)
        self.addCleanup(session.reset_for_test, None)
        self.register()
        self.fill("DT")
        c, h = _web.make_client(), _web.auth()
        v = c.get("/api/kensa/state", headers=h).get_json()["view"]
        self.assertEqual(v["worker_time"]["S2"]["who"], "検査員01")
        self.assertIn('id="whoToggle"', c.get("/kensa", headers=h).get_data(as_text=True))
