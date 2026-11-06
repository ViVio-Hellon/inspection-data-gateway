"""Aインプット画面 /ainput と /api/ainput/*(app/routes/ainput.py → ainput_session.py → ainput.py)

検査画面のセッションに、検査中のロット H7345B0(HVC・厚 15)と Min/Max を置き、Lot情報履歴 に 1 行入れて
開く → 引当No → 枚数・包数 → 山数 → 1 回目 → デプスの伝送 → 2 回目(「ｱｲｶｽ転送を行いますか？」)→ 送信 → 日報、
と画面の往復を確かめる。業務の結果(高さチェック・電文・CSV)は tests/test_kensa_ainput.py が確かめている。

HVC のデプスは 測定機器ID の 10006(基準データ)。厚 15 × 10 枚 × 1 包 ÷ 1 山 = 150 ≤ 300 なのでデプス。
理論値: Max 15.05 × 10 = 150.5、Min 14.95 × 10 = 149.5。許容: 上 15.05×3/4 = 11.2875、下 14.95/3 = 4.983。
"""
from __future__ import annotations

import unittest
from unittest import mock

from tests._web import make_client
from tests.test_kensa_session import SessionBase
from ufkensa import user_settings
from ufkensa.kensa import ainput_session, paths, quirks, session
from ufkensa.kensa.store import LOT_HISTORY
from tests._isolation import vba_reproduce

LOT = "H7345B0"
HIKI = "60602048"
ROW = ["2026-11-02", "04077194", LOT, 15, 1252, 2502, "H6X1355", "F52S", "R", "H279", "", "", "", "", 10, 0, 0, 0,
       None, None, 0, HIKI, 0, 0, "", "1P1202"]
OK_TEXT = "正常に送信が行われました。"


@vba_reproduce
class AinputScreenTest(SessionBase):
    def setUp(self) -> None:
        super().setUp()
        st = self.s.state
        st.line_name, st.line_option, st.lot = "HVC", "HVC", LOT
        st.amin, st.amax, st.hmin, st.hmax, st.tmin, st.tmax = "14.95", "15.05", "1252.1", "1252.6", "2502.1", "2502.6"
        st.v("T1").value = "検査員01"
        self.store.ensure_sheet(LOT_HISTORY).append(ROW)
        self.nis = self.tmp / "nis"
        (self.nis / "SokuteiDataTrans").mkdir(parents=True)
        (self.nis / "SokuteiDataTrans" / "SokuteiDataTrans.exe").write_bytes(b"")
        paths.set_path("nis", str(self.nis))
        session.reset_for_test(self.s)
        self.addCleanup(session.reset_for_test, None)
        self.launched = []
        self.a = ainput_session.reset_for_test(ainput_session.AinputSession(
            self.s, launcher=lambda p: self.launched.append(p) or True, sleep=lambda s: None))
        self.addCleanup(ainput_session.reset_for_test, None)
        self.addCleanup(quirks.set_for_test, None)
        from app import shell
        p = mock.patch.object(shell, "READY_SCREENS", shell.READY_SCREENS | {"ainput"})
        p.start()
        self.addCleanup(p.stop)
        self.c = make_client()
        from app.routes import ainput as ainput_routes
        if "ainput" not in self.c.application.blueprints:
            self.c.application.register_blueprint(ainput_routes.bp)
        self.h = {"X-Tool-Token": self.c.application.config["TOKEN"]}

    def post(self, op: str, **body):
        r = self.c.post(f"/api/ainput/{op}", json=body, headers=self.h)
        self.assertEqual(r.status_code, 200, r.get_data(as_text=True))
        return r.get_json()

    def ok(self, d):
        self.assertIsNone(d["stop"], d["stop"])
        return d

    def ready(self, mounts: str = "1") -> dict:
        """開いて、引当No・梱包No・枚数・包数・山数を入れ、1 回目を押したところまで。"""
        self.ok(self.post("open", ken_m="L1"))
        self.ok(self.post("value", name="hiki_no", value=HIKI))
        self.ok(self.post("value", name="konno", value="1"))
        self.ok(self.post("value", name="mai1", value="10"))
        self.ok(self.post("value", name="cou1", value="1"))
        self.ok(self.post("mount", value=mounts))
        return self.ok(self.post("press"))

    def result(self, text: str) -> None:
        (self.nis / "result.txt").write_bytes((text + "\r\n").encode("cp932"))

    # ------------------------------------------------------------------
    def test_page_and_state(self) -> None:
        r = self.c.get("/ainput", headers=self.h)
        self.assertEqual(r.status_code, 200)
        html = r.get_data(as_text=True)
        self.assertIn("pressBtn", html)
        self.assertIn("ainput.js", html)
        v = self.c.get("/api/ainput/state", headers=self.h).get_json()["view"]
        self.assertFalse(v["ready"])
        self.assertEqual((v["kensa"]["lot"], v["kensa"]["minmax"]), (LOT, True))
        self.assertEqual(self.c.post("/api/ainput/nothing", json={}, headers=self.h).status_code, 404)

    def test_open_guards(self) -> None:
        d = self.post("open")
        self.assertEqual(d["stop"]["text"], "検査法を選択してください")                 # UF1:457
        self.s.state.amax = ""
        d = self.post("open", ken_m="L1")
        self.assertEqual(d["stop"]["text"], "MinMax処理を行ってください")               # UF1:474
        self.assertFalse(d["view"]["ready"])
        d = self.post("value", name="konno", value="1")
        self.assertIn("開いてください", d["stop"]["text"])

    def test_open_copies_kensa_lot(self) -> None:
        d = self.ok(self.post("open", ken_m="M6"))
        v = d["view"]
        self.assertTrue(v["ready"])
        self.assertEqual((v["lot"]["lot"], v["lot"]["ken_no"], v["lot"]["atu"], v["lot"]["amin"], v["lot"]["ken_law"]),
                         (LOT, "H6X1355", "15.000", "14.95", "M6"))
        self.assertEqual(v["hiki_list"], [HIKI])
        self.assertEqual((v["fields"]["remain"], v["mount"]["count"]), ("0", "1"))     # 設計時の値
        self.assertTrue(v["show"]["material"])                                         # M6 は風袋・総重量を出す
        self.assertEqual(v["mai_label"], "梱包枚数")
        self.assertEqual([c["code"] for c in v["send"]["codes"]], ["A000", "A001", "A002", "A003"])
        self.assertEqual(v["send"]["codes"][2]["text"], "ホストへの送信内容が正しくありません。NISへ連絡。")
        d = self.ok(self.post("value", name="hiki_no", value=HIKI))
        self.assertEqual((d["view"]["lot"]["order_no"], d["view"]["lot"]["adjust_visible"]), ("04077194", False))
        d = self.ok(self.post("value", name="konno", value="1"))
        self.assertEqual(d["view"]["fields"]["konno"], "01")                            # KONNO_BeforeUpdate
        d = self.ok(self.post("value", name="mai1", value="1x"))
        self.assertEqual(d["view"]["fields"]["mai1"], "")                               # 数でなければ消す
        d = self.ok(self.post("spin", name="mai", dir="down"))
        self.assertTrue(d["view"]["show"]["mai2"])
        d = self.ok(self.post("value", name="remain", value="5"))
        self.assertEqual(d["view"]["fields"]["remain"], "")                             # 0/1 以外は空

    def test_two_presses_send_and_daily(self) -> None:
        self.store.ensure_sheet("Daily_HVC").write(1, 1, "2026/11/02_1直")
        d = self.ready()
        v = d["view"]
        self.assertEqual((v["mount"]["visible"], v["mount"]["device"], v["mount"]["per_mount"]),
                         (True, "Depth測定", "150.00"))
        self.assertEqual(v["press"]["phase"], "first")
        d = self.ok(self.post("receive", text="DT10001+0000150.000M"))                  # 機器番号がデプスと違う
        self.assertEqual(d["view"]["mount"]["heights"][0]["value"], "")
        self.assertIn("10006", d["view"]["mount"]["received"]["reason"])
        d = self.ok(self.post("receive", text="DT10006+0000150.000M"))
        self.assertEqual(d["view"]["mount"]["heights"][0]["value"], "150")
        self.assertEqual(d["messages"][0]["sound"], "ainput_height")
        self.assertEqual(d["view"]["press"]["phase"], "second")
        d = self.post("press")
        self.assertEqual(d["question"]["key"], "ainput_send")
        self.assertEqual(d["question"]["text"], "ｱｲｶｽ転送を行いますか？")
        self.assertTrue(d["view"]["check"]["ok"])
        self.assertEqual((d["view"]["check"]["gm"], d["view"]["check"]["gp"]), ("149.50", "150.50"))
        d = self.ok(self.post("press", answers={"ainput_send": False}))
        self.assertEqual(d["view"]["send"]["stage"], "cancel")
        self.result(OK_TEXT)
        d = self.ok(self.post("press", answers={"ainput_send": True}))
        s = d["view"]["send"]
        self.assertEqual((s["stage"], s["label"]), ("sent", OK_TEXT))
        self.assertTrue(all(x["state"] == "ok" for x in s["steps"]), s["steps"])
        self.assertEqual(len(self.launched), 1)
        self.assertTrue((self.nis / "hostsend.txt").read_bytes().startswith(b"KEN:AIN:0185" + HIKI.encode()))
        self.assertIn("引当No_" + HIKI, s["current"])
        self.assertEqual(d["view"]["fields"]["konno"], "")
        self.assertEqual(d["view"]["mount"]["heights"][0]["value"], "150")              # B64: 送信後も山高さが残る
        daily = self.store.sheet("Daily_HVC")
        self.assertEqual(daily.get(2, 4), 1.0)              # 梱包No "01"(Daily_Report(4)。Excel の代入で数になる)
        self.assertEqual(s["history"][0]["hiki_no"], HIKI)

    def test_height_ng_shows_which_mount(self) -> None:
        self.ready("2")
        self.ok(self.post("receive", text="DT10006+0000075.000M"))
        d = self.ok(self.post("receive", text="DT10006+0000090.000M"))
        self.assertEqual(d["view"]["mount"]["total"], "165.00")
        d = self.ok(self.post("press"))
        c = d["view"]["check"]
        self.assertFalse(c["ok"])
        self.assertTrue(c["msg"].startswith("[許容値ｵｰﾊﾞｰ] *測定高さ:165.00"))
        self.assertEqual((c["diff"], c["allow"], c["lo"], c["hi"]), ("14.50", "11.29", "74.75", "75.25"))
        self.assertEqual([(r["value"], r["out"], r["diff"]) for r in c["rows"]],
                         [("75", False, ""), ("90", True, "+14.75")])
        self.assertTrue(c["rows"][1]["worst"])
        self.assertEqual([h["value"] for h in d["view"]["mount"]["heights"]], ["", ""])   # NG で山高さを消す
        self.assertEqual(d["view"]["send"]["steps"][1]["state"], "ng")
        self.assertEqual(self.launched, [])

    def test_result_ng_then_ngflag_sticks(self) -> None:
        self.ready()
        self.post("receive", text="DT10006+0000150.000M")
        self.result("ホストへの送信内容が正しくありません。NISへ連絡してください。")
        d = self.ok(self.post("press", answers={"ainput_send": True}))
        self.assertEqual(d["view"]["send"]["stage"], "stopped")
        self.assertEqual(d["messages"][-1]["text"], "ﾎｽﾄへの送信内容が正しくありません。担当者へ連絡して下さい。\r\nA002")
        self.assertEqual(d["view"]["send"]["steps"][5]["state"], "ng")
        self.result(OK_TEXT)
        d = self.ok(self.post("press", answers={"ainput_send": True}))                  # Ngflag が残る(VBA どおり)
        self.assertEqual((d["view"]["send"]["stage"], d["view"]["send"]["ng_flag"]), ("stopped", "stop"))
        self.assertEqual(d["view"]["send"]["steps"][4]["state"], "ng")
        self.ok(self.post("open", ken_m="L1"))                                          # 開き直すと解ける
        self.assertEqual(self.a.form.ng_flag, "")

    def test_b65_result_missing(self) -> None:
        self.ready()
        self.post("receive", text="DT10006+0000150.000M")
        d = self.ok(self.post("press", answers={"ainput_send": True}))
        self.assertEqual(d["view"]["send"]["stage"], "sent")                            # B65: 無くても成功扱い
        self.assertEqual(d["view"]["send"]["steps"][5]["state"], "warn")
        quirks.set_for_test({"B65": "fix", "B64": "fix"})
        self.ok(self.post("open", ken_m="L1"))
        self.ready()
        self.post("receive", text="DT10006+0000150.000M")
        d = self.ok(self.post("press", answers={"ainput_send": True}))
        self.assertEqual(d["view"]["send"]["stage"], "stopped")
        self.assertIn("result.txt がありません", d["messages"][-1]["text"])
        self.result(OK_TEXT)
        self.ok(self.post("open", ken_m="L1"))
        self.ready()
        self.post("receive", text="DT10006+0000150.000M")
        d = self.ok(self.post("press", answers={"ainput_send": True}))
        self.assertEqual(d["view"]["send"]["stage"], "sent")
        self.assertEqual(d["view"]["mount"]["heights"][0]["value"], "")                 # B64 を直すと消える

    def test_reset_asks_then_resends(self) -> None:
        self.ready()
        d = self.post("reset")
        self.assertEqual(d["question"]["key"], "ainput_reset")
        d = self.ok(self.post("reset", answers={"ainput_reset": True}))
        self.assertEqual((d["view"]["fields"]["hiki_no"], d["view"]["mount"]["count"]), ("", ""))
        self.assertEqual(d["view"]["send"]["stage"], "reset")
        self.assertEqual(len(self.launched), 1)

    def test_gaityuu_unlock(self) -> None:
        self.ok(self.post("open", ken_m="L1"))
        d = self.ok(self.post("value", name="ai", value="1"))
        self.assertFalse(d["view"]["extras"]["ai"])                                     # Locked の間は変えられない
        d = self.ok(self.post("unlock"))
        self.assertEqual(d["messages"][0]["text"], "合紙,VCの有無を変更したい場合\r\n赤BOXに 外注 と入力してください。")
        self.assertTrue(d["view"]["show"]["gaityuu"])
        d = self.ok(self.post("value", name="ai", value="1"))
        self.assertTrue(d["view"]["extras"]["ai"])


if __name__ == "__main__":
    unittest.main()
