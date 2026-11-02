"""検査画面 /kensa と /api/kensa/* の往復(app/routes/kensa.py → session.py)"""
from __future__ import annotations

import sqlite3
from datetime import datetime

from tests._web import make_client
from tests.test_kensa_session import CARD, LOT, Clock, SessionBase
from ufkensa import user_settings
from ufkensa.kensa import session, terminal


class KensaScreenTest(SessionBase):
    def setUp(self) -> None:
        super().setUp()
        session.reset_for_test(self.s)
        self.addCleanup(session.reset_for_test, None)
        self.c = make_client()
        self.h = {"X-Tool-Token": self.c.application.config["TOKEN"]}

    def post(self, op: str, **body):
        r = self.c.post(f"/api/kensa/{op}", json=body, headers=self.h)
        self.assertEqual(r.status_code, 200, r.get_data(as_text=True))
        return r.get_json()

    def test_page_and_state(self) -> None:
        r = self.c.get("/kensa", headers=self.h)
        self.assertEqual(r.status_code, 200)
        self.assertIn("mgrid", r.get_data(as_text=True))                 # A案の測定グリッド
        st = self.c.get("/api/kensa/state", headers=self.h).get_json()
        self.assertIn("grid", st["view"])
        self.assertIsNone(st["stop"])

    def test_round_trip(self) -> None:
        d = self.post("start")
        self.assertIsNone(d["stop"])
        self.assertEqual(d["view"]["header"]["line"], "HVC")
        d = self.post("worker", card=CARD)
        self.assertEqual(d["question"]["key"], "worker_register")       # 確認を返す → 答えを持って呼び直す
        d = self.post("worker", card=CARD, answers={"worker_register": True})
        self.assertTrue(d["view"]["workers"][0].startswith("T1: 検査員01"))
        d = self.post("lot", lot=LOT.lower(), answers={"lot_register": True})   # Lot_KeyPress で大文字に
        self.assertEqual(d["view"]["header"]["lot"], LOT)
        d = self.post("work", work="厚板")
        self.assertEqual(d["view"]["header"]["work"], "厚板")
        d = self.post("receive", text="DT10001+0000015.012M")
        s1 = d["view"]["grid"]["rows"][0]["cells"][0]
        self.assertEqual((s1["name"], s1["value"], s1["color"]), ("S1", "15.012", "OK"))
        d = self.post("click", name="S1")
        self.assertEqual(d["view"]["next"]["name"], "S1")
        d = self.post("dblclick", name="S1")
        self.assertEqual(d["view"]["grid"]["rows"][0]["cells"][0]["value"], "")
        self.assertEqual(d["view"]["confi"]["page"], 1)                   # 消去理由(ConfiReason_Open)
        d = self.post("value", name="UFConfi.List1", value="測定位置間違い")
        self.assertIsNone(d["view"]["confi"])
        d = self.post("print-check")
        self.assertFalse(d["view"]["print_ready"])
        self.assertTrue(any(c["ok"] is False for c in d["view"]["checks"]))

    def test_no_line(self) -> None:
        user_settings.save(terminal.KEY_FACILITY, "")
        d = self.post("start")
        self.assertIn("ライン設定がありません", d["stop"]["text"])

    def test_unknown_op(self) -> None:
        r = self.c.post("/api/kensa/nothing", json={}, headers=self.h)
        self.assertEqual(r.status_code, 404)
