"""名簿・日報の画面 /meibo と /api/meibo/*(VBA の UFNameList)"""
from __future__ import annotations

from tests._web import make_client
from tests.test_kensa_session import SessionBase
from ufkensa.kensa import session


class MeiboScreenTest(SessionBase):
    def setUp(self) -> None:
        super().setUp()
        session.reset_for_test(self.s)
        self.addCleanup(session.reset_for_test, None)
        self.c = make_client()
        app = self.c.application
        if "meibo" not in app.blueprints:
            from app.routes import meibo
            app.register_blueprint(meibo.bp)
        self.h = {"X-Tool-Token": app.config["TOKEN"]}
        self.ok(self.s.start())                                           # ライン HVC

    def get(self, q: str):
        r = self.c.get("/api/meibo/view?" + q, headers=self.h)
        self.assertEqual(r.status_code, 200, r.get_data(as_text=True))
        return r.get_json()["view"]

    def test_page(self) -> None:
        r = self.c.get("/meibo", headers=self.h)
        self.assertEqual(r.status_code, 200)
        self.assertIn("meiboTable", r.get_data(as_text=True))

    def test_mismatch(self) -> None:
        self.assertIn("WorkerTime_HVC がありません", self.get("mode=1")["message"])
        sh = self.store.add_sheet("WorkerTime_HVC")
        sh.write(2, 1, "LotNo")
        sh.write(3, 1, "A1")
        sh.write(14, 1, "A14")                                            # 13 行目より下は最初の表示に出さない
        sh.write(1, 15, "不一致")
        sh.write(2, 15, "検査員02")
        v = self.get("mode=1")
        self.assertEqual((v["content"], v["head"][0], len(v["rows"])), ("検査員不一致", "LotNo", 10))
        self.assertEqual(len(self.get("mode=1&list=all")["rows"]), 12)      # 全ﾘｽﾄ: A 列の最終行まで
        v = self.get("mode=1&list=mismatch")
        self.assertEqual((v["head"], v["rows"]), (["不一致", "", ""], [["検査員02", "", ""]]))

    def test_pack_count(self) -> None:
        sh = self.store.add_sheet("梱包枚数計算_HVC")
        for c, h in enumerate(("LotNo", "梱包枚数", "LotNo", "検入枚数"), start=1):
            sh.write(1, c, h)
        sh.write(2, 1, "H7345B0")
        sh.write(2, 2, 20)
        v = self.get("mode=3")
        self.assertEqual(v["head"], ["LotNo", "梱包枚数", "LotNo", "検入枚数"])
        self.assertEqual(v["rows"], [["H7345B0", "20", "", ""]])
        self.assertFalse(v["editable"])

    def test_daily_write_back(self) -> None:
        sh = self.store.add_sheet("Daily_HVC")
        for c in range(1, 14):
            sh.write(2, c, f"h{c}")
        sh.write(3, 1, "H7345B0")
        sh.write(3, 6, "x")
        sh.write(4, 6, "y")
        v = self.get("mode=2")
        self.assertTrue(v["editable"])
        self.assertEqual(len(v["rows"]), 2)
        rows = [r[:] for r in v["rows"]][:1]                              # 2 行目を外して
        rows[0][4] = "12"
        r = self.c.post("/api/meibo/daily", json={"rows": rows}, headers=self.h).get_json()
        self.assertEqual(r["question"]["text"], "日報ﾃﾞｰﾀを修正しますか？")
        r = self.c.post("/api/meibo/daily", json={"rows": rows, "answers": {"daily": False}}, headers=self.h).get_json()
        self.assertEqual(sh.get(4, 6), "y")                               # いいえ: 変えない
        r = self.c.post("/api/meibo/daily", json={"rows": rows, "answers": {"daily": True}}, headers=self.h).get_json()
        self.assertEqual(sh.get(3, 5), "12")                              # Split の結果(文字)のまま
        self.assertIsNone(sh.get(4, 6))                                   # 外した行は消える
        self.assertEqual(len(r["view"]["rows"]), 1)

    def test_bad_mode(self) -> None:
        self.assertEqual(self.c.get("/api/meibo/view?mode=9", headers=self.h).status_code, 400)
