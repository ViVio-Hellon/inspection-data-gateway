"""理由選択のマスタ(伝送用ファイル の表「理由選択」。reasons.py)。"""
from __future__ import annotations

import sqlite3

from tests.test_kensa_session import SessionBase
from ufkensa.kensa import master_admin, reasons


class ReasonMasterTest(SessionBase):
    def db(self):
        return self.db_dir / "伝送用ファイル.sqlite3"

    def test_auto_create_with_vba_defaults(self) -> None:
        c = sqlite3.connect(self.db())
        self.assertIsNone(c.execute("SELECT 1 FROM sqlite_master WHERE name = '理由選択'").fetchone())
        c.close()
        self.assertEqual(reasons.choices(reasons.KIND_DELETE), list(reasons.DEFAULTS["消去"]))
        self.assertEqual(reasons.choices(reasons.KIND_REMEASURE), list(reasons.DEFAULTS["再測定"]))
        c = sqlite3.connect(self.db())
        n = c.execute('SELECT COUNT(*) FROM "理由選択"').fetchone()[0]
        c.close()
        self.assertEqual(n, 12)
        self.assertTrue(reasons.ensure_table())                     # 2 回目は作らない
        names = [t["name"] for t in master_admin.tables("denso")["tables"]]
        self.assertIn("理由選択", names)                            # 保守 → マスタ管理 で直せる

    def test_edited_master_is_used_on_screen(self) -> None:
        reasons.ensure_table()
        c = sqlite3.connect(self.db())
        with c:
            c.execute('UPDATE "理由選択" SET "有効" = 0 WHERE "区分" = ? AND "理由" = ?', ("消去", "板廃棄"))
            c.execute('INSERT INTO "理由選択" VALUES (?, ?, ?, 1)', ("消去", 0, "台車移動"))
        c.close()
        got = reasons.choices(reasons.KIND_DELETE)
        self.assertEqual(got[0], "台車移動")                        # 順番どおり
        self.assertNotIn("板廃棄", got)                             # 有効 = 0 は出さない
        # 検査画面: ダブルクリックで消すと、マスタの選択肢が出る
        self.register()
        self.fill("DT")
        r = self.s.dblclick("S1")
        self.assertEqual(r.view["confi"]["page"], 1)
        self.assertEqual(r.view["confi"]["choices"], got)

    def test_master_screen_creates_it(self) -> None:
        names = [t["name"] for t in master_admin.tables("denso")["tables"]]   # 保守 → マスタ管理 を開いただけで作る
        self.assertIn("理由選択", names)

    def test_missing_db_falls_back(self) -> None:
        self.assertEqual(reasons.choices(reasons.KIND_DELETE, path=self.tmp / "無い.sqlite3"),
                         list(reasons.DEFAULTS["消去"]))

    def test_master_admin_edits_show_in_open_dialog(self) -> None:
        """保守 → マスタ管理 で 足す・直す・消す と、出ている確認の選択肢にすぐ出る(古いまま残らない)。"""
        self.register()
        self.fill("DT")
        r = self.s.dblclick("S1")
        self.assertEqual(r.view["confi"]["page"], 1)
        master_admin.add_row("denso", "理由選択", {"区分": "消去", "順番": 99, "理由": "台車移動", "有効": 1})
        rows = master_admin.page("denso", "理由選択", limit=100)["rows"]
        rid = next(x["__rowid"] for x in rows if x["理由"] == "板廃棄")
        master_admin.save_row("denso", "理由選択", rid, {"理由": "板廃棄(割れ)"})
        rid2 = next(x["__rowid"] for x in rows if x["理由"] == "その他")
        master_admin.delete_row("denso", "理由選択", rid2)
        got = self.s._view()["confi"]["choices"]
        self.assertIn("台車移動", got)
        self.assertIn("板廃棄(割れ)", got)
        self.assertNotIn("その他", got)
