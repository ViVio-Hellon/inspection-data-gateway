"""全ヵ所・H494 画面 /zenkasho と /api/zenkasho/*(app/routes/zenkasho.py → zenkasho_session.py)

基準データ(tests/kensa_fixture.py)のロット RX021T0(厚 25 × 巾 1385 × 丈 2770.5、LVC)で、
開く(作業名の InputBox)→ ロット → 伝送文字列で測定 → 作業の切替(5N・三菱・H494)→ 検入高さ → 印刷・保存、
作業中保存からの読み戻しを確かめる。業務の結果そのものは test_kensa_modes_* で確かめている。
"""
from __future__ import annotations

import shutil
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path

from tests import kensa_fixture
from tests._web import make_client
from ufkensa import user_settings
from ufkensa.kensa import paths, store, terminal, zenkasho_session
from ufkensa.kensa.modes import sheet as sh

LOT = "RX021T0"


class Clock:
    def __init__(self) -> None:
        self.t = datetime(2026, 11, 2, 10, 5, 30)

    def __call__(self) -> datetime:
        self.t += timedelta(seconds=7)
        return self.t


class ZenBase(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="ufkensa_zen_"))
        kensa_fixture.use_offline(kensa_fixture.copy_to(self.tmp / "db"))
        paths.set_path("local_history", str(self.tmp / "検査履歴") + "/")
        terminal.set_facility("LVC")
        self.store = store.LocalStore(self.tmp / "data" / "kensa_local.sqlite3")
        self.clock = Clock()
        self.s = self.new_session()

    def new_session(self) -> zenkasho_session.ZenkashoSession:
        return zenkasho_session.ZenkashoSession(self.store, now=self.clock, kensa_state=lambda: None,
                                                write_logs=False)

    def tearDown(self) -> None:
        self.store.close()
        user_settings.save(terminal.KEY_FACILITY, "")
        paths.reset_path("local_history")
        shutil.rmtree(self.tmp, ignore_errors=True)

    def open_ippan(self) -> dict:
        d = self.s.start().to_dict()
        self.assertEqual(d["question"]["key"], "mit_choice")           # V1 が空 → InputBox(UFMit:2210)
        self.assertEqual(d["question"]["buttons"], "input")
        d = self.s.start("一般").to_dict()
        self.assertIsNone(d["question"])
        d = self.s.set_form("LotBox", LOT.lower()).to_dict()          # KeyPress で大文字
        self.assertEqual(d["view"]["form"]["lot"], LOT)
        return d


class SessionTest(ZenBase):
    def test_measure_flow(self) -> None:
        d = self.open_ippan()
        v = d["view"]
        self.assertEqual((v["mode"], v["page"]), ("一般", "P1"))
        self.assertEqual([m["key"] for m in v["modes"] if m["active"]], ["一般"])
        self.assertEqual(v["active"]["label"], "No.1 厚1")
        self.assertEqual(len(v["rows"]), 70)
        self.assertEqual(v["groups"][0], {"label": "厚", "span": 8, "tol": "24.550〜25.450"})
        d = self.s.receive("DT10002+0000025.012M").to_dict()
        row = d["view"]["rows"][0]["cells"]
        self.assertEqual(row[0:2], [["25.012", ""], ["○", "ok"]])
        self.assertEqual(d["view"]["active"]["label"], "No.1 厚2")      # 判定記号の入れ子の Change で次へ
        self.assertEqual(d["view"]["flash"], "OK")
        self.assertIn("chime2.wav", [m["sound"] for m in d["messages"]])
        d = self.s.receive("DT10002+0000026.000M").to_dict()            # 公差外
        self.assertEqual(d["view"]["rows"][0]["cells"][3], ["×", "ng"])
        self.assertEqual(d["view"]["flash"], "NG")
        self.s.select(7, 23)                                             # ラテラル: Change は動かさない → Enter で下へ
        d = self.s.receive("0.3").to_dict()
        self.assertEqual(d["view"]["active"]["label"], "No.2 ﾗﾃﾗﾙﾎﾞｰ")
        self.s.select(7, 30)
        d = self.s.receive("良").to_dict()                               # 文字で Change を抜ける → EnableEvents=False のまま
        self.assertFalse(d["view"]["events_enabled"])
        d = self.s.resume().to_dict()                                    # 処理再開
        self.assertTrue(d["view"]["events_enabled"])
        d = self.s.erase().to_dict()
        self.assertEqual(d["view"]["rows"][1]["cells"][27][0], "")

    def test_modes(self) -> None:
        self.open_ippan()
        d = self.s.choose("5N").to_dict()
        v = d["view"]
        self.assertEqual(v["mode"], "5N")
        hiz = next(g for g in v["groups"] if g["label"] == "歪")
        self.assertEqual(hiz["tol"], "幅 2.0 / 丈 3.0(5N 固定)")
        self.assertEqual([p["which"] for p in v["prints"]], ["fini"])
        self.assertEqual(v["notice"], "")
        d = self.s.choose("三菱").to_dict()
        v = d["view"]
        self.assertEqual(v["notice"], "三菱受託は確認待ちの項目があります(docs/04 W2)")
        self.assertEqual([p["which"] for p in v["prints"]], ["mitsubishi"])
        self.assertIn("元板NO", [c["label"] for c in v["columns"]])
        d = self.s.choose("H494").to_dict()
        v = d["view"]
        self.assertEqual((v["mode"], v["page"]), ("H494", "P_H494"))
        self.assertEqual(len(v["rows"]), 10)
        self.assertIn("幅3", [c["label"] for c in v["columns"]])
        self.assertFalse(self.s.book.sheets[sh.ZEN].visible)            # QueryClose で隠す
        d = self.s.choose("一般").to_dict()
        self.assertEqual(d["view"]["page"], "P1")
        self.assertEqual(d["view"]["mode"], "一般")

    def test_height_and_print_save(self) -> None:
        self.open_ippan()
        self.s.resume()
        self.s.receive("DT10002+0000025.012M")
        d = self.s.set_form("YamaLB", "1").to_dict()
        self.assertEqual([k["n"] for k in d["view"]["form"]["kenh"] if k["visible"]], [1])
        self.s.set_form("KenHBox1", "250")
        d = self.s.save_height().to_dict()
        self.assertEqual(d["view"]["form"]["takasa_formula"], "(250.0÷25)=10")
        self.assertIn("ゲージによる測定必要", [m["text"] for m in d["messages"]])
        self.assertEqual(self.s.book.ws(sh.ZEN).get(2, 30), "1山")
        self.s.set_form("Kname", "検査員01")
        d = self.s.print("normal").to_dict()
        self.assertEqual(d["question"]["key"], "zen_print")
        d = self.s.print("normal", {"zen_print": True}).to_dict()
        self.assertIsNone(d["stop"])
        rep = d["view"]["reports"][0]
        html = Path(rep["path"]).read_text(encoding="utf-8")
        self.assertIn("厚公差：24.550～25.450", html)                  # 右ヘッダ(Mit:901)
        self.assertIn("25.012", html)
        saved = d["view"]["saved"]
        self.assertTrue(any(p.endswith(f"20261102_{LOT}.html") for p in saved))
        self.assertTrue(any(p.endswith(f"20261102_{LOT}_全ヵ所検査保存.csv") for p in saved))
        hz = self.store.sheet(store.HOZON)
        self.assertEqual(hz.get(hz.last_row(2), 2), LOT)                # 保存用へ履歴行
        self.assertEqual(hz.get(hz.last_row(2), 77), "別ファイル参照")
        # 2 回目は連番(.html を見て .002)
        d = self.s.print("normal", {"zen_print": True}).to_dict()
        self.assertTrue(any(p.endswith(f"20261102_{LOT}.002.html") for p in d["view"]["saved"]))

    def test_restore(self) -> None:
        self.open_ippan()
        self.s.receive("DT10002+0000025.012M")
        s2 = self.new_session()                                          # アプリを開き直した
        d = s2.start().to_dict()
        v = d["view"]
        self.assertEqual(v["mode"], "一般")                              # V1 から戻る(InputBox は出ない)
        self.assertEqual(v["rows"][0]["cells"][0], ["25.012", ""])
        self.assertEqual(v["form"]["lot"], LOT)
        self.assertTrue(self.store.exists("全ヵ所作業中保存"))

    def test_new_sheet_question(self) -> None:
        self.open_ippan()
        self.s.receive("DT10002+0000025.012M")
        d = self.s.new_sheet().to_dict()
        self.assertEqual(d["question"]["key"], "zen_new")
        d = self.s.new_sheet({"zen_new": True}).to_dict()
        self.assertEqual(d["view"]["rows"][0]["cells"][0][0], "")
        self.assertEqual(d["view"]["form"]["lot"], "")
        self.assertFalse(d["view"]["events_enabled"])                    # LotBox="" で EnableEvents=False(VBA どおり)


class ScreenTest(ZenBase):
    def setUp(self) -> None:
        super().setUp()
        zenkasho_session.reset_for_test(self.s)
        self.addCleanup(zenkasho_session.reset_for_test, None)
        self.c = make_client()
        app = self.c.application
        from app.routes import zenkasho
        if "zenkasho" not in app.blueprints:
            app.register_blueprint(zenkasho.bp)
        if "pending_zenkasho" in app.view_functions:                    # 「準備中」の同じ URL より先に出す
            app.view_functions["pending_zenkasho"] = zenkasho.page
        self.h = {"X-Tool-Token": app.config["TOKEN"]}

    def post(self, op: str, **body):
        r = self.c.post(f"/api/zenkasho/{op}", json=body, headers=self.h)
        self.assertEqual(r.status_code, 200, r.get_data(as_text=True))
        return r.get_json()

    def test_page_and_round_trip(self) -> None:
        r = self.c.get("/zenkasho", headers=self.h)
        self.assertEqual(r.status_code, 200)
        self.assertIn('id="sheet"', r.get_data(as_text=True))
        st = self.c.get("/api/zenkasho/state", headers=self.h).get_json()
        self.assertIn("rows", st["view"])
        d = self.post("start")
        self.assertEqual(d["question"]["key"], "mit_choice")
        d = self.post("start", choice_input="５N")                        # 全角は 5N に(UFMit:2212)
        self.assertEqual(d["view"]["mode"], "5N")
        d = self.post("form", name="LotBox", value=LOT)
        d = self.post("resume")
        d = self.post("receive", text="DT10002+0000025.012M")
        self.assertEqual(d["view"]["rows"][0]["cells"][1], ["○", "ok"])
        d = self.post("select", r=9, c=11)
        self.assertEqual(d["view"]["active"]["label"], "No.3 幅1")
        d = self.post("select", r=999, c=1)
        self.assertIn("セルの位置", d["stop"]["text"])
        d = self.post("choose", mode="H494")
        self.assertEqual(d["view"]["page"], "P_H494")
        d = self.post("print", which="normal")                           # H494 では一般の印刷は無い
        self.assertIsNotNone(d["stop"])

    def test_unknown_op_and_report(self) -> None:
        r = self.c.post("/api/zenkasho/nothing", json={}, headers=self.h)
        self.assertEqual(r.status_code, 404)
        r = self.c.get("/api/zenkasho/report/x.txt", headers=self.h)
        self.assertEqual(r.status_code, 400)
        r = self.c.get("/api/zenkasho/report/none.html", headers=self.h)
        self.assertEqual(r.status_code, 404)


if __name__ == "__main__":
    unittest.main()
