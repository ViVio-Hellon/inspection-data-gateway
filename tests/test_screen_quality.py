"""品質基準画面 /quality と /api/quality/*(app/routes/quality.py → quality_session.py)

照合の結果そのものは test_kensa_quality が見ている。ここは「規格ごとのカード」への組み立て・参照元の添え方・
いま検査中のロットとの受け渡しを見る。基準データ(tests/kensa_fixture.py)のロット:
H7345B0(厚 15 の厚板)、L726M50(厚 2 の薄板、パターン 4)、H6083T1(DIM_122 別表A)。
"""
from __future__ import annotations

import shutil
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock

from tests import kensa_fixture
from tests._web import make_client
from ufkensa import user_settings
from ufkensa.kensa import quality_session as qs
from ufkensa.kensa import session, store, terminal


class QualityBase(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="ufkensa_q_"))
        kensa_fixture.use_offline(kensa_fixture.copy_to(self.tmp / "db"))
        terminal.set_facility("HVC")
        self.store = store.LocalStore(self.tmp / "local.sqlite3")
        self.kensa = session.Session(self.store, now=lambda: datetime(2026, 11, 2, 10, 5, 30),
                                     calibration_skip=lambda: False, write_logs=False)

    def tearDown(self) -> None:
        self.store.close()
        user_settings.save(terminal.KEY_FACILITY, "")
        shutil.rmtree(self.tmp, ignore_errors=True)

    def build(self, lot: str) -> dict:
        return qs.build(lot, store=self.store, kensa=self.kensa)

    @staticmethod
    def card(d: dict, key: str) -> dict:
        return next(c for c in d["cards"] if c["key"] == key)


class QualityBuildTest(QualityBase):
    def test_thick_lot_cards(self) -> None:
        d = self.build("h7345b0")                                       # 大文字にして引く
        self.assertTrue(d["found"])
        self.assertFalse(d["current"])
        h = d["header"]
        self.assertEqual((h["lot"], h["ken_no"], h["zai"], h["tyu"], h["youto"]), ("H7345B0", "H6X1355", "995", "R", "H165"))
        self.assertEqual(h["equipment"], "可")
        self.assertEqual([c["key"] for c in d["cards"]],
                         ["distortion", "appearance", "bh", "diagonal", "lateral", "ratefuri", "distortion_ref"])
        for c in d["cards"]:
            self.assertTrue(c["sources"], c["key"])                     # 参照元は必ず添える
        dist = self.card(d, "distortion")
        self.assertEqual(dist["sources"][0], "Q-S-QA-DIM-122_別表E")    # 平坦度上限値参照元
        self.assertIn("品質基準: 平坦度一覧QP追加", dist["sources"])
        self.assertEqual(dict(dist["rows"])["巾歪"], "2.0")
        ap = self.card(d, "appearance")
        self.assertEqual(ap["title"], "外観(厚板)")
        self.assertEqual(ap["sources"], ["Q-S-QA-124 別表1", "Q-S-QA-124 別表2"])
        self.assertEqual(ap["table"]["head"], ["欠陥名", "許容内容", "グレード"])
        self.assertEqual(ap["table"]["rows"][0][0], "飛込み")
        bh = self.card(d, "bh")
        self.assertEqual((bh["status"], dict(bh["rows"])["下限値"]), ("none", "不要"))
        self.assertEqual(dict(self.card(d, "lateral")["rows"])["等級からの公差"], "3.0")   # 等級 3 → ReLateralbow
        self.assertEqual(dict(self.card(d, "ratefuri")["rows"])["引当の出どころ"], "引当(SIKAHIKINOW)")

    def test_thin_lot_surface_table(self) -> None:
        ap = self.card(self.build("L726M50"), "appearance")
        self.assertEqual(ap["title"], "外観(薄板)")
        self.assertEqual(dict(ap["rows"])["パターン"], "4")
        self.assertEqual(ap["table"]["rows"][0], ["トビコミ", "2㎜以下", "爪に掛らない", "ﾋﾟｯﾁ300㎜以上", "/"])
        self.assertIn("Q-S-QA-SRF-128 別表1", ap["sources"])

    def test_distortion_ref_and_not_found(self) -> None:
        ref = self.card(self.build("H6083T1"), "distortion_ref")
        self.assertEqual(ref["sources"], ["Q-S-QA-DIM-122"])
        self.assertEqual(dict(ref["rows"])["1M当りｹﾞｰｼﾞ"], "1.2")
        d = self.build("ZZZZZZZ")
        self.assertFalse(d["found"])
        self.assertEqual(d["status"], "not_found")
        self.assertEqual(d["cards"], [])

    def test_current_lot_is_default(self) -> None:
        st = self.kensa.state
        st.lot, st.atu, st.ken_no_caption, st.lateral_text = "H7345B0", "15.000", "H6X1355", "3.0"
        d = self.build("")
        self.assertEqual(d["lot"], "H7345B0")
        self.assertTrue(d["current"])
        self.assertEqual(dict(d["kensa"])["ラテラルボー"], "3.0")      # 検査画面の値は読むだけ
        self.assertEqual(st.lot, "H7345B0")
        self.assertEqual(qs.current_lot(self.kensa)["lot"], "H7345B0")
        st.lot = ""
        self.assertEqual(self.build("")["cards"], [])

    def test_lot_history_allocations(self) -> None:
        sh = self.store.sheet(store.LOT_HISTORY)
        row = [""] * 26
        for i, name in enumerate(store.LOT_HISTORY_HEADERS):
            row[i] = {"ﾛｯﾄNo": "H7345B0", "納入先": "ｸﾉｻﾝｷﾞﾖｳ(ｶ ﾎﾝｼﾔ", "得意先": "x", "用途ｺｰﾄﾞ": "H165", "材質": "995",
                      "調質": "R", "厚": "15", "巾": "1252", "丈": "2502", "検査No": "H6X1355"}.get(name, "")
        sh.append(row)
        rf = self.card(self.build("H7345B0"), "ratefuri")
        self.assertEqual(dict(rf["rows"])["引当の出どころ"], "Lot情報履歴")
        self.assertEqual(rf["table"]["rows"][0][:2], ["ｸﾉｻﾝｷﾞﾖｳ(ｶ ﾎﾝｼﾔ", "x"])


class QualityRouteTest(QualityBase):
    def setUp(self) -> None:
        super().setUp()
        session.reset_for_test(self.kensa)
        self.addCleanup(session.reset_for_test, None)
        from app import shell
        p = mock.patch.object(shell, "READY_SCREENS", shell.READY_SCREENS | {"quality"})
        p.start()
        self.addCleanup(p.stop)
        self.client = make_client()
        app = self.client.application
        if "quality" not in app.blueprints:
            from app.routes import quality
            app.register_blueprint(quality.bp)
        self.h = {"X-Tool-Token": app.config["TOKEN"]}

    def test_page_and_api(self) -> None:
        r = self.client.get("/quality", headers=self.h)
        self.assertEqual(r.status_code, 200)
        self.assertIn("qCards", r.get_data(as_text=True))
        d = self.client.get("/api/quality/lot?lot=H6083T1", headers=self.h).get_json()
        self.assertEqual(d["header"]["lot"], "H6083T1")
        self.assertEqual(len(d["cards"]), 7)
        self.assertEqual(self.client.get("/api/quality/lot", headers=self.h).get_json()["lot"], "")   # 検査中なし
        self.assertEqual(self.client.get("/api/quality/current", headers=self.h).get_json()["lot"], "")
        self.assertEqual(self.client.get("/api/quality/lot?lot=a'b", headers=self.h).status_code, 400)


if __name__ == "__main__":
    unittest.main()
