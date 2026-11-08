"""定盤検査 / ゲージ検査(幅ｹﾞｰｼﾞ・丈ｹﾞｰｼﾞ)の切り替えを画面の操作どおりに通す。

VBA: QP_DB.IsGaugeRequired(QP_DB:2473-2511、品質基準「ゲージ判定条件」1行=AND・行間=OR)が True なら
LZS〜LZF="幅ｹﾞｰｼﾞ"・LGS〜LGF="丈ｹﾞｰｼﾞ"(FormRelat:3572-3587 / SaveLoad:487-501)、False なら "定盤"/"ｹﾞｰｼﾞ"。
Judg.HizumiJudg(J:198-231)は両方ゲージなら Z 欄を CB?1、G 欄を CB?3 の幅/丈公差で判定、定盤なら G 欄は白。
Print:1127-1129 は LZS="幅ｹﾞｰｼﾞ" のとき (11,5)="幅ゲージ"・(11,6)="丈ゲージ"。
試験用の「ゲージ判定条件」に 材質=995(H7345B0 の材質)だけの 1 行を足してゲージのロットにする。
"""
from __future__ import annotations

import sqlite3

from tests.test_kensa_session import SessionBase
from ufkensa.kensa.store import HOZON

FACES = "SUCDEF"


class GaugeBase(SessionBase):
    rule: tuple = ()

    def setUp(self) -> None:
        super().setUp()
        if self.rule:
            con = sqlite3.connect(self.db_dir / "品質基準.sqlite3")
            con.execute("INSERT INTO ゲージ判定条件 (管理番号, 材質, 調質, 有効) VALUES (?, ?, ?, ?)", self.rule)
            con.commit()
            con.close()

    @staticmethod
    def caps(st) -> tuple[list[str], list[str]]:
        return [st.v("LZ" + f).value for f in FACES], [st.v("LG" + f).value for f in FACES]

    def finish(self) -> str:
        s = self.s
        self.fill("DT")
        self.fill("#L")
        self.strain_visual_height()
        self.ok(s.min_max())
        r = self.ok(s.print_check())
        self.assertTrue(r.view["print_ready"])
        self.assertEqual([c["label"] for c in r.view["checks"]][1:3], ["定盤", "ゲージ"])
        rep = self.ok(s.print_commit({"print": True})).view["report"]
        self.assertEqual(self.store.sheet(HOZON).row(2)[1], "H7345B0")
        return open(rep["report"]["html"], encoding="utf-8").read()


class SurfacePlateLotTest(GaugeBase):
    """既定の条件(重量/丈)に当たらない 15x1252x2502(約127kg)は定盤検査。"""

    def test_plate(self) -> None:
        s = self.s
        self.register()
        st = s.state
        self.assertEqual(self.caps(st), (["定盤"] * 6, ["ｹﾞｰｼﾞ"] * 6))
        # 定盤: Z 欄は CB?1 の公差(幅=HaHiz 2.0)、公差ちょうどは OK
        self.ok(s.set_value("CBS1", "幅"))
        self.ok(s.set_value("ZS", "2.0"))
        self.assertEqual(st.strain_z["S"].color, "OK")
        self.ok(s.set_value("ZS", "2.1"))
        self.assertEqual(st.strain_z["S"].color, "NG")
        # G 欄(ｹﾞｰｼﾞ)は CB?3 があっても判定しない(白)
        self.ok(s.set_value("CBS3", "丈"))
        self.ok(s.set_value("GS", "9"))
        self.assertNotIn(st.strain_g["S"].color, ("OK", "NG"))
        self.ok(s.set_value("ZS", ""))
        self.ok(s.set_value("GS", ""))
        html = self.finish()
        self.assertNotIn("幅ゲージ", html)
        self.assertNotIn("丈ゲージ", html)
        s2 = self.new_session()
        self.ok(s2.start())
        self.assertEqual(self.caps(s2.state), (["定盤"] * 6, ["ｹﾞｰｼﾞ"] * 6))


class GaugeLotTest(GaugeBase):
    rule = (9001, "995", None, 1)

    def test_gauge(self) -> None:
        s = self.s
        self.register()
        st = s.state
        self.assertEqual(self.caps(st), (["幅ｹﾞｰｼﾞ"] * 6, ["丈ｹﾞｰｼﾞ"] * 6))
        self.assertEqual((st.v("HaHiz").value, st.v("TaHiz").value), ("2.0", "3.0"))
        # 両方ゲージ: Z=CB?1、G=CB?3 の公差で判定(> で NG)
        self.ok(s.set_value("CBS1", "幅"))
        self.ok(s.set_value("ZS", "2.0"))
        self.assertEqual(st.strain_z["S"].color, "OK")
        self.ok(s.set_value("ZS", "2.1"))
        self.assertEqual(st.strain_z["S"].color, "NG")
        self.ok(s.set_value("CBS3", "丈"))
        self.ok(s.set_value("GS", "3.0"))
        self.assertEqual(st.strain_g["S"].color, "OK")
        r = self.ok(s.set_value("GS", "3.1"))
        self.assertEqual(st.strain_g["S"].color, "NG")
        self.assertIn("kensa_ng", [m.sound for m in r.messages])
        self.ok(s.set_value("ZS", ""))
        self.ok(s.set_value("GS", ""))
        html = self.finish()
        self.assertIn("幅ゲージ", html)
        self.assertIn("丈ゲージ", html)
        # 作業中保存からの読み戻しで IsGaugeRequired をやり直してゲージ表示に戻る(SaveLoad:487)
        s2 = self.new_session()
        self.ok(s2.start())
        self.assertEqual(self.caps(s2.state), (["幅ｹﾞｰｼﾞ"] * 6, ["丈ｹﾞｰｼﾞ"] * 6))
        self.ok(s2.set_value("CBS3", "丈"))
        self.ok(s2.set_value("GS", "3.5"))
        self.assertEqual(s2.state.strain_g["S"].color, "NG")


class InvalidRuleTest(GaugeBase):
    """有効=False の行・条件の一部が外れる行(AND)ではゲージにならない。"""
    rule = (9002, "995", "ZZ", 1)

    def test_and_condition(self) -> None:
        self.register()
        self.assertEqual(self.caps(self.s.state), (["定盤"] * 6, ["ｹﾞｰｼﾞ"] * 6))


class DisabledRuleTest(GaugeBase):
    rule = (9003, "995", None, 0)

    def test_disabled(self) -> None:
        self.register()
        self.assertEqual(self.caps(self.s.state), (["定盤"] * 6, ["ｹﾞｰｼﾞ"] * 6))
