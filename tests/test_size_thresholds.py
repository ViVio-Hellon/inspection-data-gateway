"""寸法のしきい値のマスタ(伝送用ファイル の表「寸法しきい値」。size_thresholds.py。2026-11-11)。

- 既定(表が無い・既定のまま)で、いままでの直書きと同じ動き(境界の すぐ下・ちょうど・すぐ上)
- 値を変えると、使う所すべてに効く
"""
from __future__ import annotations

import sqlite3
import unittest
from unittest import mock

from tests.test_kensa_layout import fake_db
from tests.test_kensa_session import SessionBase
from tests.test_screen_maint import MaintRouteTest
from ufkensa.kensa import ainput, checks, grid, instruments as ins, judge, layout, size_thresholds as th
from ufkensa.kensa.grid import KensaState


def visible(atu: float, weight: float = 0.0) -> list[str]:
    s = KensaState(lot="LOT0001", work="厚板")
    layout.atu_opselect(s, 3.0, fake_db(atu=str(atu)), weight, [])
    return [w for w in grid.WORKS if s.work_visible[w]]


class _Tmp(SessionBase):
    def db(self):
        return self.db_dir / "伝送用ファイル.sqlite3"


class ParseAndStoreTest(_Tmp):
    def test_defaults_are_vba(self) -> None:
        self.assertEqual(th.values(), th.DEFAULTS)
        self.assertEqual((th.thin_atu(), th.sakko_thick(), th.kennyu_n(), th.get("conb_take"), th.get("depth_conb")),
                         (6.0, 3.0, 25, 5500.0, 300.0))
        self.assertEqual((th.le_caption(), th.ge_caption()), ("検入≦25", "検入≧26"))
        self.assertEqual(KensaState().inspe_butt_caption, "検入≦25")
        c = sqlite3.connect(self.db())                                  # 読むだけでは表を作らない
        self.assertIsNone(c.execute("SELECT 1 FROM sqlite_master WHERE name = ?", (th.TABLE,)).fetchone())
        c.close()

    def test_parse(self) -> None:
        self.assertEqual(th.parse("thin_atu", " 5.5 "), 5.5)
        self.assertEqual(th.parse("kennyu_n", "30"), 30.0)
        for k, v in (("thin_atu", "a"), ("thin_atu", ""), ("thin_atu", "-1"), ("kennyu_n", "2.5"), ("kennyu_n", "0"),
                     ("thin_atu", "nan"), ("nokey", "1")):
            with self.assertRaises(th.Invalid, msg=(k, v)):
                th.parse(k, v)

    def test_ensure_save_read_and_pairs(self) -> None:
        self.assertTrue(th.ensure_table())
        c = sqlite3.connect(self.db())
        self.assertEqual(c.execute(f'SELECT COUNT(*) FROM "{th.TABLE}"').fetchone()[0], len(th.ITEMS))
        c.close()
        self.assertEqual(th.values(fresh=True), th.DEFAULTS)
        th.save({"thin_atu": "5"})
        self.assertEqual(th.values(fresh=True)["thin_atu"], 5.0)
        self.assertEqual(th.values()["op_thin_hi"], 6.0)                # 足りないキーはいまの値
        with self.assertRaises(th.Invalid):
            th.save({"op_thin_lo": "7"})                                 # 下限 > 上限
        with self.assertRaises(th.Invalid):
            th.save({"thin_atu": "x"})
        self.assertEqual(th.values(fresh=True)["thin_atu"], 5.0)        # 悪い値は何も書かない

    def test_bad_row_uses_default_and_unreadable_uses_last_known(self) -> None:
        th.save({"sakko_thick": "2.5"})
        c = sqlite3.connect(self.db())
        with c:
            c.execute(f'UPDATE "{th.TABLE}" SET "値" = ? WHERE "キー" = ?', ("x", "thin_atu"))
        c.close()
        v = th.values(fresh=True)
        self.assertEqual((v["thin_atu"], v["sakko_thick"]), (6.0, 2.5))
        with mock.patch.object(th.datasource, "get_table", side_effect=sqlite3.OperationalError("locked")):
            self.assertEqual(th.values(fresh=True)["sakko_thick"], 2.5)  # 最後に読めた値

    def test_history(self) -> None:
        from ufkensa import trace
        with mock.patch.object(trace, "event", wraps=trace.event) as ev:
            th.save({"conb_take": "6000"})
            th.save({"conb_take": "6000"})                              # 変わらなければ残さない
        mine = [c for c in ev.call_args_list if c.args and c.args[0] == th.CHANGE_EV]
        self.assertEqual(len(mine), 1)
        self.assertEqual(mine[0].kwargs["data"], {"changed": {"conb_take": ["5500", "6000"]}})
        recs = [{"t": "2026-11-11T10:00", "ev": th.CHANGE_EV, "user": "u", "pc": "p",
                 "data": {"changed": {"conb_take": ["5500", "6000"]}}}, {"ev": "other"}]
        with mock.patch.object(trace, "read", return_value=recs):
            h = th.history()
        self.assertEqual(h[0]["changes"], [{"key": "conb_take", "label": th.BY_KEY["conb_take"].label,
                                            "before": "5500", "after": "6000"}])
        self.assertEqual(len(h), 1)


class DefaultsReproduceVbaTest(_Tmp):
    """既定のまま: 境界のすぐ下・ちょうど・すぐ上で、いままでの直書きと同じ。"""

    ALL = ["厚板", "薄板", "少枚数", "極厚", "スラブ"]

    def test_op_select_boundaries(self) -> None:
        cases = {
            0.999: ["厚板", "薄板", "少枚数", "極厚", "スラブ"],
            1.0: ["薄板"],
            6.0: ["薄板"],
            6.0005: ["厚板", "薄板", "少枚数"],                       # 6.0〜6.001 の隙間(VBA のまま)
            6.001: ["厚板", "少枚数"],
            29.9: ["厚板", "少枚数"],
            29.95: ["厚板", "少枚数", "極厚"],
            199.9: ["厚板", "少枚数", "極厚"],
            199.95: ["厚板", "少枚数", "極厚", "スラブ"],
            200.1: ["厚板", "少枚数", "極厚", "スラブ"],
            299.0: ["厚板", "少枚数", "極厚", "スラブ"],
            299.5: ["厚板", "少枚数", "極厚", "スラブ"],
            300.0: ["少枚数", "極厚", "スラブ"],
            800.0: ["少枚数", "極厚", "スラブ"],
            800.5: ["厚板", "薄板", "少枚数", "極厚", "スラブ"],
        }
        for atu, want in cases.items():
            self.assertEqual(visible(atu), want, atu)
        self.assertEqual(visible(250.0, 1500.0), ["厚板", "少枚数", "極厚", "スラブ"])   # 重量ちょうどは隠さない
        self.assertEqual(visible(250.0, 1500.1), ["少枚数", "極厚", "スラブ"])
        self.assertEqual(visible(200.05, 2000.0), ["厚板", "少枚数", "極厚", "スラブ"])  # 200〜200.1 の隙間

    def test_thin_atu_places(self) -> None:
        self.assertEqual(checks.cb_shape_items(5.99), ["幅", "丈", "耳", "中"])
        self.assertEqual(checks.cb_shape_items(6.0), ["幅", "丈", "耳", "中"])
        self.assertEqual(checks.cb_shape_items(6.01), ["幅", "丈"])
        self.assertEqual(checks.cb_shape_items(), ["幅", "丈", "耳", "中"])

    def test_sakko(self) -> None:
        for thick, want in (("2.99", True), ("3", True), ("3.01", False)):
            self.assertEqual(self.sakko(thick), want, thick)

    def sakko(self, thick: str) -> bool:
        flds = ["製造板厚", "品質ｸﾞﾚｰﾄﾞ_表面処理", "用途ｺｰﾄﾞ"]
        def ab(db, lot, kind, flag):
            return [flds] if kind == "Fields" else [[thick, "B", "U19B"]]
        with mock.patch.object(judge, "_last_index", lambda f, n: f[0].index(n)), \
             mock.patch.object(judge, "_at", lambda r, c: r[c]):
            return judge.advance_check("LOT", arrbox=ab)

    def conb(self, tak: str) -> str:
        t = ins.TransferInput(x=9, y=5, atu="3", haba="1000", yout="K126", line="L-1", tak=tak, hab="1000.0")
        return [v for (_r, _c, v) in ins.machine_transfer(t, db=lambda *a, **k: "") if "C1" in v][0]

    def test_conb_and_depth(self) -> None:
        self.assertEqual([self.conb(x) for x in ("5499.9", "5500", "5500.1")], ["L-1 C12", "L-1 C12", "L-1 C11"])
        self.assertEqual([ainput.l1_caption(x) for x in (299.9, 300.0, 300.1)],
                         ["Depth測定", "Depth測定", "ｺﾝﾍﾞｯｸｽ測定"])

    def test_kennyu_caption_state(self) -> None:
        s = KensaState()
        self.assertTrue(th.is_le(s.inspe_butt_caption))
        self.assertTrue(th.is_ge("検入≧26") and not th.is_ge("検入≦25"))


class ChangedValuesTakeEffectTest(_Tmp):
    def test_op_select(self) -> None:
        th.save({"op_thin_hi": "8", "op_notthin_lo": "8.001", "op_notgoku_hi": "40", "op_heavy_weight": "1000"})
        self.assertEqual(visible(7.0), ["薄板"])
        self.assertEqual(visible(35.0), ["厚板", "少枚数"])
        self.assertEqual(visible(250.0, 1200.0), ["少枚数", "極厚", "スラブ"])

    def test_thin_atu(self) -> None:
        th.save({"thin_atu": "5"})
        self.assertEqual(checks.cb_shape_items(5.5), ["幅", "丈"])
        self.assertEqual(checks.cb_shape_items(), ["幅", "丈", "耳", "中"])   # 厚が無いときは薄板の形状のまま

    def test_sakko_conb_depth(self) -> None:
        th.save({"sakko_thick": "4", "conb_take": "6000", "depth_conb": "250"})
        self.assertTrue(DefaultsReproduceVbaTest.sakko(self, "3.5"))
        self.assertEqual(DefaultsReproduceVbaTest.conb(self, "5800"), "L-1 C12")
        self.assertEqual(ainput.l1_caption(260.0), "ｺﾝﾍﾞｯｸｽ測定")

    def test_kennyu_captions(self) -> None:
        th.save({"kennyu_n": "30"})
        self.assertEqual((th.le_caption(), th.ge_caption()), ("検入≦30", "検入≧31"))
        s = KensaState()
        self.assertEqual(s.inspe_butt_caption, "検入≦30")
        self.assertFalse(grid.ken_flags(s, "L")[4])
        s.inspe_butt_caption = "検入≧26"                                # 前の値の文字のまま残っていても同じ意味
        self.assertTrue(grid.ken_flags(s, "L")[4])


class SessionKennyuTest(_Tmp):
    """検査画面: 枚数を変えるとボタンの文字が変わり、切替は今まで通り。"""

    def test_toggle(self) -> None:
        from tests.test_feature_kensa_count import THIN_LOT
        from tests.test_kensa_session import CARD
        from ufkensa.kensa import terminal
        th.save({"kennyu_n": "30"})
        terminal.set_facility("LVC")
        self.write_calibration()
        self.s = self.new_session()
        self.ok(self.s.start())
        self.ok(self.s.register_worker(CARD, {"worker_register": True}))
        self.ok(self.s.register_lot(THIN_LOT, {"lot_register": True}))
        self.ok(self.s.select_work("薄板"))
        self.ok(self.s.set_yama("1"))
        st = self.s.state
        self.assertEqual(st.inspe_butt_caption, "検入≦30")
        self.ok(self.s.toggle_inspe())
        self.assertEqual(st.inspe_butt_caption, "検入≧31")
        self.assertFalse(st.ctl("DiagC").visible)
        self.ok(self.s.toggle_inspe())
        self.assertEqual(st.inspe_butt_caption, "検入≦30")
        self.assertTrue(st.ctl("DiagC").visible)


class ThresholdsRouteTest(MaintRouteTest):
    def setUp(self) -> None:
        super().setUp()
        self.db = self.tmp / "伝送用ファイル.sqlite3"
        sqlite3.connect(self.db).close()
        p = mock.patch.object(th, "_db_path", lambda: self.db)
        p.start()
        self.addCleanup(p.stop)

    def test_api(self) -> None:
        d = self.client.get("/api/maint/thresholds", headers=self.h).get_json()
        items = {i["key"]: i for g in d["groups"] for i in g["items"]}
        self.assertEqual(items["op_notthin_lo"]["value"], "6.001")
        self.assertEqual(items["kennyu_n"]["value"], "25")
        c = sqlite3.connect(self.db)
        self.assertEqual(c.execute(f'SELECT COUNT(*) FROM "{th.TABLE}"').fetchone()[0], len(th.ITEMS))
        c.close()
        r = self.post("/api/maint/thresholds", {"values": {"thin_atu": "5.5"}})
        self.assertEqual(r.status_code, 200)
        self.assertEqual({i["key"]: i["value"] for g in r.get_json()["groups"] for i in g["items"]}["thin_atu"], "5.5")
        r = self.post("/api/maint/thresholds", {"values": {"thin_atu": "x"}})
        self.assertEqual((r.status_code, r.get_json()["error"]["field"]), (400, "thin_atu"))
        self.assertEqual(self.post("/api/maint/thresholds", {}).status_code, 400)
        from ufkensa import auth
        auth.lock_master()
        r = self.post("/api/maint/thresholds", {"values": {"thin_atu": "4"}})
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "locked"))
        self.assertEqual(th.values(fresh=True)["thin_atu"], 5.5)
        self.assertIn('data-tab="thresh"', self.client.get("/maint", headers=self.h).get_data(as_text=True))


for _name in dir(MaintRouteTest):
    if _name.startswith("test_"):
        setattr(ThresholdsRouteTest, _name, None)
del MaintRouteTest

if __name__ == "__main__":
    unittest.main()
