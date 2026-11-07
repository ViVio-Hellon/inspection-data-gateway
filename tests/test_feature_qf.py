"""QFフラグ(QF.Caption)の通し試験。

VBA: ロット登録 AccReflec で 実績_設備ｺｰｽ に "QF" を含むと QF.Caption = "QF有り"(SaveLoad:213)、
作業中保存の 17 行 2 列へ書き(SaveLoad:2063)・読み戻す(SaveLoad:2603)、新規検査で消す(UF1:2989/3446)。
BH_Check(O:2121-2157)は QF 通過材の BH 未入力を見る(呼び出し元はコメントアウト UF1:3878)。
"""
from __future__ import annotations

import sqlite3

from tests import kensa_fixture
from tests.test_kensa_session import LOT, SessionBase
from ufkensa.kensa import checks, wiring

QF_COURSE = "HOT ST2 QF ANF ST2 PSW TUR"


class QfFlagTest(SessionBase):
    def setUp(self) -> None:
        super().setUp()
        # 一時コピーの仕掛だけ書き換える(基準データは触らない)
        with sqlite3.connect(self.db_dir / kensa_fixture.file_name("SIKALOTNOW")) as c:
            n = c.execute('UPDATE "仕掛" SET "実績_設備ｺｰｽ"=? WHERE "ﾛｯﾄ番号"=?', (QF_COURSE, LOT)).rowcount
        self.assertGreater(n, 0)

    def test_qf_lot_flow(self) -> None:
        s = self.s
        self.register()
        st = s.state
        self.assertEqual(st.v("QF").value, "QF有り")                     # SaveLoad:213
        self.assertEqual(s._view()["bh_note"]["qf"], "QF有り")
        self.assertEqual(st.v("LastHot").value, "最終熱処理工程：ANF")    # 最右の熱処理工程

        # BH_Check: QF 通過で BH 未入力なら位置名(厚1 を入れた後)
        self.fill("DT")
        self.assertEqual(checks.bh_check(wiring.to_uf1(st)), st.header("E"))
        self.fill("#L")
        self.strain_visual_height()
        self.ok(s.min_max())
        r = self.ok(s.print_check())
        self.assertTrue(r.view["print_ready"])                          # QF は印刷を止めない(BH_Check は呼ばれない)
        self.ok(s.print_commit({"print": True}))

        # 作業中保存 17 行 2 列 → 開き直しで復元(SaveLoad:2063/2603)
        s2 = self.new_session()
        self.ok(s2.start())
        self.assertEqual(s2.state.lot, LOT)
        self.assertEqual(s2.state.v("QF").value, "QF有り")

        # 新規検査で消える(UF1:2989)
        self.ok(s2.new_inspection({"new": True}))
        self.assertEqual(s2.state.v("QF").value, "")

    def test_non_qf_lot_clears(self) -> None:
        """QF 無しの工程では空(B54 の修正: 前ロットの表示を残さない)。"""
        with sqlite3.connect(self.db_dir / kensa_fixture.file_name("SIKALOTNOW")) as c:
            c.execute('UPDATE "仕掛" SET "実績_設備ｺｰｽ"=? WHERE "ﾛｯﾄ番号"=?', ("HOT ST2 TUR", LOT))
        self.s.state.v("QF").value = "QF有り"                             # 前ロットの残り
        self.register()
        self.assertEqual(self.s.state.v("QF").value, "")

    def test_bh_check_unit(self) -> None:
        self.register()
        st = self.s.state
        self.assertEqual(checks.bh_check(wiring.to_uf1(st)), "")       # 測定前は対象外
        st.v("QF").value = ""
        self.fill("DT")
        self.assertEqual(checks.bh_check(wiring.to_uf1(st)), "")       # QF 無しは常に ""
