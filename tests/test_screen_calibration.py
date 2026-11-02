"""校正画面 /calibration と /api/calibration/*(app/routes/calibration.py → calibration_session.py)

業務の結果(判定・書くセル)は test_kensa_calibration が見ている。ここは画面の組み立てと手元DB への反映、
通常検査の校正判定(D11)とのつながりを見る。基準データ(tests/kensa_fixture.py)の 校正ストレートエッヂ:
HVC は 長さ下 2052.7・長さ上 2053.7。時計は 2026-11-02 10:05:30 から(HVC 3 交替の 1直)。
"""
from __future__ import annotations

import shutil
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path
from unittest import mock

from flask import Flask

from tests import kensa_fixture
from tests._web import make_client
from ufkensa import user_settings
from ufkensa.kensa import calibration_session as cs
from ufkensa.kensa import quirks, session, store, terminal


class Clock:
    def __init__(self, start: datetime) -> None:
        self.t = start

    def __call__(self) -> datetime:
        self.t = self.t + timedelta(seconds=7)
        return self.t


class CalBase(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="ufkensa_cal_"))
        kensa_fixture.use_offline(kensa_fixture.copy_to(self.tmp / "db"))
        terminal.set_facility("HVC")
        self.clock = Clock(datetime(2026, 11, 2, 10, 5, 30))
        self.store = store.LocalStore(self.tmp / "local.sqlite3")
        self.kensa = session.Session(self.store, now=self.clock, calibration_skip=lambda: False, write_logs=False)
        self.kensa.state.line_name = "HVC"                    # 検査画面は起動済み(UF1.LineName)
        self.c = cs.CalibrationSession(self.store, now=self.clock, kensa=self.kensa, write_logs=False)

    def tearDown(self) -> None:
        self.store.close()
        user_settings.save(terminal.KEY_FACILITY, "")
        shutil.rmtree(self.tmp, ignore_errors=True)


class CalibrationSessionTest(CalBase):
    def test_open_shows_uncalibrated_and_template(self) -> None:
        self.assertFalse(self.kensa._calibrated())
        v = self.c.open().view
        self.assertFalse(v["status"]["calibrated"])
        self.assertEqual((v["status"]["date"], v["status"]["shift"]), ("2026.11.02", "1直"))
        self.assertEqual(v["step"], "1000")
        self.assertEqual(v["gauge"], "[HVC] P2:1000mm直尺の全長を測定してください")      # 校正_Open(Proofreading:12)
        self.assertEqual(v["range_1000"], ["999.5", "1000.5"])
        self.assertEqual((v["straight_edge"]["lo"], v["straight_edge"]["hi"]), ("2052.7", "2053.7"))
        self.assertEqual(v["center"], "2003.8")                                          # RANK_CENTER の HVC
        self.assertTrue(v["template_created"])
        self.assertTrue(self.store.exists(cs.HINSHO) and self.store.exists(cs.KIROKU))
        self.assertEqual(self.store.sheet(cs.HINSHO).text(9, 1), "2日")                   # 6 + (2-1)*3
        self.assertEqual(self.store.sheet(cs.HINSHO).text(98, 2), "3直")

    def test_full_pass_unblocks_kensa(self) -> None:
        self.c.open()
        r = self.c.receive("#L1000.2")
        self.assertEqual([m.sound for m in r.messages], ["chime2.wav"])
        self.assertEqual(r.view["step"], "gauge")
        self.assertEqual(r.view["gauge"], "[HVC]:直尺「S-1」の全長を測定してください")
        hs = self.store.sheet(cs.HINSHO)
        self.assertEqual((hs.get(9, 3), hs.text(9, 9)), (1000.2, "HVC"))                 # 2日 1直 の行(Proofreading:316)
        r = self.c.receive("DT10001+0000002053.0M")                                      # DT は 14 文字目から 6 文字
        self.assertEqual(r.view["step"], "done")
        self.assertTrue(r.view["closed"])
        self.assertTrue(r.view["status"]["calibrated"])
        k = self.store.sheet(cs.KIROKU)
        self.assertEqual((k.text(2, 1), k.text(2, 4), k.get(2, 5), k.text(2, 7)), ("2026.11.02", "1直", 2053.0, "HVC"))
        rec = r.view["records"][0]
        self.assertEqual((rec["value"], rec["shift"], rec["ng"]), ("2053.0", "1直", False))
        self.assertTrue(self.kensa._calibrated())                                       # D11: 検査画面の止めが外れる
        self.assertEqual(self.c.receive("#L2053.0").messages, [])                       # 終わった後は受け付けない

    def test_gauge_ng_then_clear(self) -> None:
        self.c.open()
        self.c.skip_1000()
        r = self.c.receive("#L2060.0")
        self.assertEqual([m.sound for m in r.messages], ["tada.wav"])
        self.assertEqual(r.view["ng_value"], "測定値 2060.0 mm")                         # Proofreading:131
        self.assertTrue(r.view["saiken"])
        self.assertFalse(r.view["accepting"])
        self.assertTrue(r.view["records"][0]["ng"])                                     # H 列「〆」
        self.assertFalse(r.view["status"]["calibrated"])
        self.assertEqual(self.c.receive("#L2053.0").view["step"], "gauge")              # F1=0 のまま: 受け付けない
        v = self.c.clear().view
        self.assertTrue(v["accepting"])
        self.assertEqual(self.c.receive("#L2053.0").view["step"], "done")

    def test_1000_ng_and_bad_text(self) -> None:
        self.c.open()
        r = self.c.receive("#L1001.0")
        self.assertEqual(r.view["ng_value"], "999.5 ≦測定値≦ 1000.5 になっていません")
        self.assertEqual(r.view["step"], "1000")
        self.assertEqual(self.store.sheet(cs.HINSHO).get(6, 13), 1001.0)               # 不合格は M 列に(Proofreading:340)
        r = self.c.receive("1000.0")                                                    # DT / #L でない
        self.assertEqual([m.text for m in r.messages], ["正しいデータが伝送できていません"])

    def test_skip_1000_follows_b14(self) -> None:
        self.c.open()
        self.assertTrue(self.c.view()["can_skip_1000"])
        self.assertEqual(self.c.skip_1000().view["step"], "gauge")
        self.c.open()
        old = quirks.reproduce
        quirks.reproduce = lambda b: False if b == "B14" else old(b)
        self.addCleanup(setattr, quirks, "reproduce", old)
        self.assertFalse(self.c.view()["can_skip_1000"])
        self.assertEqual(self.c.skip_1000().view["step"], "1000")

    def test_kname_from_latest_registration(self) -> None:
        st = self.kensa.state
        st.v("T1").value, st.v("Day1").value = "検査員01", "2026.11.01  10:00"
        st.v("T2").value, st.v("Day2").value = "検査員02_検査員03", "2026.11.02  08:00"
        self.assertEqual(self.c.open().view["kname"], "検査員03")                      # UF1.CommandButton6 の Max(Day)
        self.assertEqual(cs.manual_kname(cs.workers.WorkerRegistry()), "")

    def test_no_line(self) -> None:
        self.kensa.state.line_name = ""
        user_settings.save(terminal.KEY_FACILITY, "")
        r = self.c.open()
        self.assertIn("ライン設定がありません", r.stop.text)
        self.assertIsNone(r.view["status"])

    def test_rank_buttons(self) -> None:
        self.c.open()
        v = self.c.rank("kousei").view
        self.assertEqual(v["rank_kind"], "校正")
        self.assertEqual(len(v["ranks"]), 6)
        self.assertEqual(self.c.rank("hinsho").view["rank_kind"], "品証")


class CalibrationRouteTest(CalBase):
    def setUp(self) -> None:
        super().setUp()
        cs.reset_for_test(self.c)
        self.addCleanup(cs.reset_for_test, None)
        from app import shell                               # レールの「準備中」を外したときと同じにする
        p = mock.patch.object(shell, "READY_SCREENS", shell.READY_SCREENS | {"calibration"})
        p.start()
        self.addCleanup(p.stop)
        client = make_client()
        app: Flask = client.application
        if "calibration" not in app.blueprints:
            from app.routes import calibration
            app.register_blueprint(calibration.bp)
        self.client = client
        self.h = {"X-Tool-Token": app.config["TOKEN"]}

    def post(self, op: str, **body):
        r = self.client.post(f"/api/calibration/{op}", json=body, headers=self.h)
        self.assertEqual(r.status_code, 200, r.get_data(as_text=True))
        return r.get_json()

    def test_page_and_round_trip(self) -> None:
        r = self.client.get("/calibration", headers=self.h)
        self.assertEqual(r.status_code, 200)
        self.assertIn("calIn", r.get_data(as_text=True))
        d = self.post("open")
        self.assertEqual(d["view"]["step"], "1000")
        self.assertIsNone(d["stop"])
        d = self.post("receive", text="#L1000.0")
        self.assertEqual(d["messages"][0]["sound"], "chime2.wav")
        d = self.post("receive", text="#L2053.7")                                       # 上限ちょうどは合格
        self.assertEqual(d["view"]["step"], "done")
        st = self.client.get("/api/calibration/state", headers=self.h).get_json()
        self.assertTrue(st["view"]["status"]["calibrated"])
        self.assertEqual(self.post("clear")["view"]["f1"], "1")

    def test_bad_requests(self) -> None:
        self.assertEqual(self.client.post("/api/calibration/nothing", json={}, headers=self.h).status_code, 404)
        self.assertEqual(self.client.post("/api/calibration/receive", json={}, headers=self.h).status_code, 400)


if __name__ == "__main__":
    unittest.main()
