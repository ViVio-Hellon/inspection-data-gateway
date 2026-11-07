"""試験指示票(先行データ票)フラグの通し試験。VBA の Judg.AdvanceCheck(J:334-361)。

必要になる条件(SIKALOTNOW「仕掛」のロットの 1 行目):
  (製造板厚 ≤ 3.0 かつ 品質ｸﾞﾚｰﾄﾞ_表面処理 が S・T 以外)または(表面処理 が S・T 以外 かつ 用途ｺｰﾄﾞ の頭が T・D1・D2)

使われる所(VBA):
- ロット登録: 包装仕様の注意文と一緒に「先行データ票が必要です。」(Judg.MaterialsMsgBox J:2067)
- 検査員登録: 「先行データ票が必要です。」(UF1:4850)
- 検査員登録の後の最初のフォームのクリック: 「先行データ票が必要です。」(UserForm_Click、UF1:4312-4325 PicFM.Tag = 1)
- 印刷: 「先行データ票はありますか？」いいえなら印刷しない(UF1:3643-3655)
"""
from __future__ import annotations

import sqlite3
import unittest

from tests import kensa_fixture
from tests.test_kensa_session import CARD, LOT, SessionBase
from ufkensa.kensa import judge

NEED = "先行データ票が必要です。"


class AdvanceCheckRuleTest(unittest.TestCase):
    """条件の組み合わせ(J:351-359)。"""

    def check(self, thick: str, surface: str, youto: str) -> bool:
        fields = ["製造板厚", "品質ｸﾞﾚｰﾄﾞ_表面処理", "用途ｺｰﾄﾞ"]

        def arrbox(file, lot, kind, db_flag):
            return fields if kind == "Fields" else [[thick, surface, youto]]
        return judge.advance_check("L1", arrbox=arrbox)

    def test_rules(self) -> None:
        self.assertTrue(self.check("3.0", "X", "H165"))        # 3mm 以下 かつ S・T 以外
        self.assertTrue(self.check("2.0", "", "H165"))
        self.assertFalse(self.check("3.1", "X", "H165"))       # 3mm を超え、用途も当たらない
        self.assertFalse(self.check("2.0", "S", "H165"))       # 表面処理 S
        self.assertFalse(self.check("2.0", "T", "T123"))       # 表面処理 T
        self.assertTrue(self.check("15", "X", "T123"))         # 用途の頭 T
        self.assertTrue(self.check("15", "X", "D1AB"))         # D1
        self.assertTrue(self.check("15", "X", "D2AB"))         # D2
        self.assertFalse(self.check("15", "X", "D3AB"))        # D3 は当たらない
        self.assertFalse(self.check("15", "S", "D1AB"))


class AdvanceSheetFlowTest(SessionBase):
    need = True

    def setUp(self) -> None:
        super().setUp()
        f = self.db_dir / kensa_fixture.file_name("SIKALOTNOW")
        with sqlite3.connect(f) as c:
            # 厚は 15 のまま(製造板厚は画面の厚にもなる)。表面処理 X と 用途の頭 T の組み合わせで「必要」にする
            c.execute('UPDATE "仕掛" SET "品質ｸﾞﾚｰﾄﾞ_表面処理" = ?, "用途ｺｰﾄﾞ" = ? WHERE "ﾛｯﾄ番号" = ?',
                      ("X" if self.need else "S", "T165" if self.need else "H165", LOT))
        c.close()

    def texts(self, r) -> list[str]:
        return [m.text for m in r.messages]

    def test_flow(self) -> None:
        s = self.s
        self.ok(s.start())
        r = s.register_worker(CARD)
        r = self.ok(s.register_worker(CARD, {"worker_register": True}))
        r = self.ok(s.register_lot(LOT, {"lot_register": True}))
        self.assertEqual(NEED in self.texts(r), self.need, self.texts(r))      # ロット登録(MaterialsMsgBox)
        self.ok(s.select_work("厚板"))
        r = s.click("")                                                        # 最初のクリック(UserForm_Click)
        self.assertEqual(NEED in self.texts(r), self.need, self.texts(r))
        r = s.click("")                                                        # 2 回目からは出ない
        self.assertNotIn(NEED, self.texts(r))
        self.fill("DT")
        self.fill("#L")
        self.strain_visual_height()
        self.ok(s.min_max())
        r = s.print_check()
        if not self.need:
            self.assertIsNone(r.question)
            self.assertTrue(r.view["print_ready"])
            return
        self.assertEqual((r.question.key, r.question.text), ("advance_sheet", "先行データ票はありますか？"))
        r = s.print_check({"advance_sheet": False})                             # いいえ: 印刷しない
        self.assertFalse(r.view["print_ready"])
        r = s.print_check({"advance_sheet": True})                              # はい: 先へ
        self.assertTrue(r.view["print_ready"], r.view.get("checks"))
        self.assertIn({"label": "先行データ票", "ok": True}, r.view["checks"])
        r = self.ok(s.print_commit({"print": True}))
        self.assertEqual(r.view["report"]["saved"]["mode"], "ADDNEW")

    def test_worker_register_after_lot(self) -> None:
        """ロットを先に登録してから検査員を登録しても案内が出る(UF1:4845-4858)。"""
        s = self.s
        self.ok(s.start())
        self.ok(s.register_lot(LOT, {"lot_register": True}))
        s.register_worker(CARD)
        r = self.ok(s.register_worker(CARD, {"worker_register": True}))
        self.assertEqual(NEED in self.texts(r), self.need, self.texts(r))


class NoAdvanceSheetFlowTest(AdvanceSheetFlowTest):
    need = False


if __name__ == "__main__":
    unittest.main()
