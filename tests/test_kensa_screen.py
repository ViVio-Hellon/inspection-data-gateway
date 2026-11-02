"""検査画面 /kensa と /api/kensa/* の往復(app/routes/kensa.py → session.py)"""
from __future__ import annotations

import unittest


from tests._web import make_client
from tests.test_kensa_session import CARD, LOT, SessionBase
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

    def test_new_ops_round_trip(self) -> None:
        """足した操作(外観・不良要因・山数・保留・オーダー・包装仕様書・印刷履歴クリア・サイズ修正)の往復。"""
        self.post("start")
        self.post("worker", card=CARD, answers={"worker_register": True})
        d = self.post("lot", lot=LOT, answers={"lot_register": True})
        v = d["view"]
        for k in ("lists", "cb_choices", "order", "pack", "vc", "buttons", "visual_caps", "ken_hl"):
            self.assertIn(k, v)
        self.post("work", work="厚板")
        d = self.post("visual", lc=2, items=[1])
        self.assertEqual(d["view"]["visual_caps"]["U_UpVisual"]["value"], v["lists"]["visual"][1])
        self.post("value", name="SBad", value="2")
        d = self.post("defect-reflect", items=[2])
        self.assertEqual(d["view"]["bad"][0]["value"], "")
        d = self.post("yamalb", value="3")
        self.assertEqual(d["view"]["ken_hl"]["HL1"]["value"], "3-3")
        d = self.post("hold")
        self.assertIsNone(d["question"])                                  # 入力先が無ければ知らせるだけ
        d = self.post("order", order=v["order"]["list"][0])
        self.assertEqual(d["view"]["order"]["no"], "1P0116")
        d = self.post("resize")
        self.assertEqual(d["question"]["key"], "resize")
        d = self.post("resize", answers={"resize": True})
        self.assertTrue(d["view"]["buttons"]["resize"])
        d = self.post("print-reset")
        self.assertIsNone(d["question"])
        d = self.post("pack-close", confirmed=True)
        self.assertIsNone(d["view"]["pack"])
        d = self.post("ng-delete", pos="S")
        self.assertEqual(d["question"]["key"], "ng_delete")
        d = self.post("namelist-ack")
        self.assertIsNone(d["view"]["name_list"])

    def test_page_has_new_parts(self) -> None:
        html = self.c.get("/kensa", headers=self.h).get_data(as_text=True)
        for part in ("visualList", "defectList", "strainGrid", "yamaLb", "orderSel", "pack_dialog", "holdBtn",
                     "izyouBtn", "resizeBtn", "printResetBtn", "manualIn", "byousyaChk"):
            self.assertIn(part, html)

    def test_picture_route(self) -> None:
        (self.db_dir / "VC種類画像").mkdir()
        (self.db_dir / "VC種類画像" / "208WRYG.JPG").write_bytes(b"\xff\xd8\xff")
        self.assertEqual(self.c.get("/api/kensa/picture/vc/208WRYG.JPG", headers=self.h).status_code, 200)
        self.assertEqual(self.c.get("/api/kensa/picture/vc/無い.JPG", headers=self.h).status_code, 404)
        self.assertEqual(self.c.get("/api/kensa/picture/pack/1P0116.JPG", headers=self.h).status_code, 404)
        self.assertEqual(self.c.get("/api/kensa/pack_image", headers=self.h).status_code, 404)   # 出していない
        self.assertEqual(self.c.get("/api/kensa/picture/vc/a.txt", headers=self.h).status_code, 400)

    def test_no_line(self) -> None:
        user_settings.save(terminal.KEY_FACILITY, "")
        d = self.post("start")
        self.assertIn("ライン設定がありません", d["stop"]["text"])

    def test_unknown_op(self) -> None:
        r = self.c.post("/api/kensa/nothing", json={}, headers=self.h)
        self.assertEqual(r.status_code, 404)


class SoundTest(unittest.TestCase):
    def test_sound_route(self):
        import tempfile
        from pathlib import Path
        from tests._web import make_client
        from ufkensa.kensa import paths
        d = Path(tempfile.mkdtemp())
        (d / "完了.wav").write_bytes(b"RIFF0000WAVE")
        paths.set_path("wav", str(d) + "/")
        self.addCleanup(paths.reset_path, "wav")
        c = make_client()
        h = {"X-Tool-Token": c.application.config["TOKEN"]}
        self.assertEqual(c.get("/api/kensa/sound/完了.wav", headers=h).status_code, 200)
        self.assertEqual(c.get("/api/kensa/sound/無い.wav", headers=h).status_code, 404)
        self.assertEqual(c.get("/api/kensa/sound/..%2Fx.txt", headers=h).status_code, 400)
