"""フリー反り(ﾌﾘｰｿﾘ)の切り替えを画面の操作どおりに通す(session.py + wiring.py)

VBA: QP.DbReturnItem(QP_DB:1392-1409)が「ラテフリ直角」でフリー反り=要の行に当たると
flatnesText="ﾌﾘｰｿﾘ要"・LGS〜LGF="ﾌﾘｰｿﾘ"・HURIS=公差・HizCome="ﾌﾘｰ反り<公差>mm以下" に切り替え、
Judg(J:155-176)は G 欄(ゲージ)を CB?3 に関係なく HURIS で判定する。保存用 84 列 "〆"(SaveLoad)・
帳票(11,6)="ﾌﾘｰｿﾘ"(Print:1130-1131)・履歴/作業中保存の読み戻しで LG? が戻る。
試験用の「ラテフリ直角」に合成の 1 行(納入先だけ一致)を足して、要のロット/要でないロットを比べる。
"""
from __future__ import annotations

import sqlite3

from tests.test_kensa_session import SessionBase
from ufkensa.kensa.store import HOZON

DELIVERY = "ｸﾉｻﾝｷﾞﾖｳ(ｶ ﾎﾝｼﾔ"          # H7345B0 の引当の納入先(基準データ)
FACES = "SUCDEF"


class FreeSoriBase(SessionBase):
    rows: tuple = ()

    def setUp(self) -> None:
        super().setUp()
        con = sqlite3.connect(self.db_dir / "品質基準.sqlite3")
        for no, flat, kousa in self.rows:
            con.execute("INSERT INTO ラテフリ直角 (管理番号, 納入先, フリー反り, 公差) VALUES (?, ?, ?, ?)",
                        (no, DELIVERY, flat, kousa))
        con.commit()
        con.close()

    def captions(self, st) -> list[str]:
        return [st.v("LG" + f).value for f in FACES]


class NoFreeSoriTest(FreeSoriBase):
    def test_normal_lot(self) -> None:
        self.register()
        st = self.s.state
        self.assertEqual((st.flatnes_text, st.huris), ("", ""))
        self.assertEqual(self.captions(st), ["ｹﾞｰｼﾞ"] * 6)
        self.assertEqual(st.v("LZS").value, "定盤")
        # CB?3 未選択なら G 欄は判定しない(J:178 以降の通常判定)
        self.s.set_value("GS", "0.6")
        self.assertNotEqual(st.strain_g["S"].color, "NG")


class FreeSoriTest(FreeSoriBase):
    rows = ((9001, "要", "0.5"),)

    def test_screen_judge_save_reload(self) -> None:
        s = self.s
        self.register()
        st = s.state
        # 画面の切り替え
        self.assertEqual((st.flatnes_text, st.huris), ("ﾌﾘｰｿﾘ要", "0.5"))
        self.assertEqual(self.captions(st), ["ﾌﾘｰｿﾘ"] * 6)
        self.assertEqual(st.v("LZS").value, "定盤")                      # 定盤側は変わらない
        self.assertEqual(st.v("HizCome").value, "ﾌﾘｰ反り0.5mm以下")
        # 判定は HURIS(CB?3 未選択でも。"/" の前で比べる)
        r = self.ok(s.set_value("GS", "0.6"))
        self.assertEqual(st.strain_g["S"].color, "NG")
        self.assertIn("kensa_ng", [m.sound for m in r.messages])
        self.ok(s.set_value("UFConfi.TB1", "許容"))                     # 確認(公差外)に答えてから次へ
        self.ok(s.set_value("GU", "0.5"))                                # 境界は合格(> で NG)
        self.assertEqual(st.strain_g["U"].color, "OK")
        self.ok(s.set_value("GS", "0.2"))
        self.assertEqual(st.strain_g["S"].color, "OK")

        self.fill("DT")
        self.fill("#L")
        self.strain_visual_height()
        self.ok(s.min_max())
        self.assertTrue(self.ok(s.print_check()).view["print_ready"])
        rep = self.ok(s.print_commit({"print": True})).view["report"]
        self.assertEqual(self.store.sheet(HOZON).row(2)[83], "〆")      # 84 列目
        html = open(rep["report"]["html"], encoding="utf-8").read()
        self.assertIn("ﾌﾘｰｿﾘ", html)

        # 作業中保存からの読み戻し
        s2 = self.new_session()
        self.ok(s2.start())
        st2 = s2.state
        self.assertEqual((st2.flatnes_text, st2.huris), ("ﾌﾘｰｿﾘ要", "0.5"))
        self.assertEqual(self.captions(st2), ["ﾌﾘｰｿﾘ"] * 6)
        self.ok(s2.set_value("GS", "0.7"))
        self.assertEqual(st2.strain_g["S"].color, "NG")

    def test_switch_back_on_new(self) -> None:
        """新規検査で flatnesText・LG? が戻り(UF1:3430-3442)、要でなくなったロットは通常の判定に戻る。"""
        s = self.s
        self.register()
        self.fill("DT")
        self.fill("#L")
        self.strain_visual_height()
        self.ok(s.min_max())
        self.ok(s.print_check())
        self.ok(s.print_commit({"print": True}))
        con = sqlite3.connect(self.db_dir / "品質基準.sqlite3")
        con.execute("DELETE FROM ラテフリ直角 WHERE 管理番号 = 9001")
        con.commit()
        con.close()
        self.ok(s.new_inspection({"new": True}))
        st = s.state
        self.assertEqual(st.flatnes_text, "")
        self.assertEqual(self.captions(st), ["ｹﾞｰｼﾞ"] * 6)
        # 要でなくなった同じロットを登録し直す → 通常の画面のまま
        r = s.register_lot("H7345B0", {"lot_register": True, "lot_split": False, "lot_recent": True})
        self.assertIsNone(r.stop, r.stop and r.stop.text)
        self.ok(s.select_work("厚板"))
        self.assertEqual((st.flatnes_text, st.huris), ("", ""))
        self.assertEqual(self.captions(st), ["ｹﾞｰｼﾞ"] * 6)


class FreeSoriToleranceRowTest(FreeSoriBase):
    """B72: 最初に一致した行が要でない行でも、公差はフリー反り=要の行から取る(VBA は HitArry(1) の公差)。"""
    rows = ((9001, "", "9.9"), (9002, "要", "0.5"))

    def test_tolerance_from_free_row(self) -> None:
        self.register()
        st = self.s.state
        self.assertEqual((st.flatnes_text, st.huris), ("ﾌﾘｰｿﾘ要", "0.5"))
        self.assertEqual(st.v("HizCome").value, "ﾌﾘｰ反り0.5mm以下")
        self.ok(self.s.set_value("GS", "0.6"))
        self.assertEqual(st.strain_g["S"].color, "NG")
