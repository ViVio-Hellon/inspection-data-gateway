"""技術指示のラテラルボー(指示_ﾗﾃﾗﾙﾎﾞｰ → LateralText.Tag = "有")の通し試験。(「試験指示票」= 先行データ票 は test_feature_shiken_shiji.py)

VBA:
- SaveLoad:70        ロット登録の初めに LateralText.Tag をリセット
- SaveLoad:277-300   SIKALOTNOW「指示_ﾗﾃﾗﾙﾎﾞｰ」が空でなければ LateralText に半角で入れ、色付け、Tag="有"、LSelectLateral
- FormRelat:6317     L11(ラテラルボー見出し)を Sumisei<>"" Or Tag="有" で表示
- FormRelat:2328/2390 検入≧26/≦25 の切替でラテラル欄 U/C/D を出し分け
- Print:1043         印刷用罫線でラテラルボー列を残す条件に Tag="有"
- SaveLoad:2107/2592 作業保存 Cells(1,42) に Tag を書き、読み戻しで復元 → SaveLoad:2878 LSelectLateral
"""
from __future__ import annotations

import sqlite3
import unittest

from tests import kensa_fixture
from tests.test_kensa_session import LOT, SessionBase


class ShikenShijiFlagTest(SessionBase):
    def setUp(self) -> None:
        super().setUp()
        # 一時コピーの SIKALOTNOW に技術指示(全角)を入れる(元の fixture は空)
        f = self.db_dir / kensa_fixture.file_name("SIKALOTNOW")
        with sqlite3.connect(f) as c:
            c.execute('UPDATE "仕掛" SET "指示_ﾗﾃﾗﾙﾎﾞｰ" = ? WHERE "ﾛｯﾄ番号" = ?', ("３.０/2M", LOT))
        c.close()

    def test_flag_display_print_and_reload(self) -> None:
        s = self.s
        self.register()
        st = s.state
        # 登録: フラグ・表示(SaveLoad:295-300)
        self.assertEqual(st.lateral_text_tag, "有")
        self.assertEqual(st.lateral_text, "3.0")
        self.assertTrue(st.lateral_text_colored)
        self.assertEqual(st.additional, "対角ﾗﾃﾗﾙﾎﾞ-追記公差あり")
        self.assertTrue(st.ctl("L11").visible)                          # FormRelat:6317

        # 測定 → 印刷・保存
        self.fill("DT")
        self.fill("#L")
        self.strain_visual_height()
        self.ok(s.min_max())
        r = self.ok(s.print_check())
        self.assertTrue(r.view["print_ready"], r.view.get("checks"))
        r = self.ok(s.print_commit({"print": True}))
        self.assertEqual(r.view["report"]["saved"]["mode"], "ADDNEW")

        # 作業保存 Cells(1,42) = "有"(SaveLoad:2107)
        self.assertEqual(self.store.sheet("作業保存_HVC").get(1, 42), "有")

        # 開き直し: Tag が戻り、L11 も表示(SaveLoad:2592, 2878)
        s2 = self.new_session()
        self.ok(s2.start())
        self.assertEqual(s2.state.lot, LOT)
        self.assertEqual(s2.state.lateral_text_tag, "有")
        self.assertTrue(s2.state.ctl("L11").visible)

    def test_flag_reset_on_next_lot_without_instruction(self) -> None:
        """フラグはロット登録ごとにリセット(SaveLoad:70)。指示を消して登録し直すと Tag は空。"""
        self.register()
        self.assertEqual(self.s.state.lateral_text_tag, "有")
        f = self.db_dir / kensa_fixture.file_name("SIKALOTNOW")
        with sqlite3.connect(f) as c:
            c.execute('UPDATE "仕掛" SET "指示_ﾗﾃﾗﾙﾎﾞｰ" = \'\' WHERE "ﾛｯﾄ番号" = ?', (LOT,))
        c.close()
        s = self.s
        self.ok(s.register_lot(LOT, {"lot_register": True, "lot_split": False, "lot_recent": True}))
        self.assertEqual(s.state.lateral_text_tag, "")
        self.assertFalse(s.state.ctl("L11").visible)


class ConvertToHalfTest(unittest.TestCase):
    def test_fullwidth_decimal_point(self) -> None:
        from ufkensa.kensa import lot
        from ufkensa.kensa.modes import zenkasho
        for f in (lot.convert_to_half, zenkasho.convert_to_half):
            self.assertEqual(f("３．０/2M"), "3.0")        # 前は "30"(全角の「．」が落ちていた)
            self.assertEqual(f("３.０"), "3.0")


if __name__ == "__main__":
    import unittest
    unittest.main()
