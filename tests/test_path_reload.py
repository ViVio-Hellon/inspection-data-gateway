"""使用パスを変えた直後に読み返すか、と 設定 → 使用パス の「読み込む」(2026-11-08)。"""
from __future__ import annotations

import shutil
import sqlite3
from unittest import mock

from tests import _web, kensa_fixture
from tests.test_kensa_session import LOT, SessionBase
from ufkensa import auth
from ufkensa.kensa import datasource, judge, paths, quality, reasons, session


class PathReloadTest(SessionBase):
    def setUp(self) -> None:
        super().setUp()
        session.reset_for_test(self.s)
        self.addCleanup(session.reset_for_test, None)
        p = mock.patch.object(auth, "master_edit_ok", return_value=True)
        p.start()
        self.addCleanup(p.stop)
        self.c, self.h = _web.make_client(), _web.auth()
        # 2 つ目の置き場所(中身を少し変えた写し)
        self.db2 = self.tmp / "db2"
        shutil.copytree(self.db_dir, self.db2)

    def save_path(self, key: str, path: str, slot: str = "") -> None:
        body = {"key": key, "path": path + "/"}
        if slot:
            body["slot"] = slot
        r = self.c.post("/api/settings/path", json=body, headers=self.h)
        self.assertEqual(r.status_code, 200, r.get_json())

    def test_quality_and_denso_reread_right_after_change(self) -> None:
        # いまの品質基準・伝送用ファイルを読んで覚えさせる
        n1 = len(datasource.get_table("品質基準", "欠陥名")[1])
        self.assertEqual(reasons.choices(reasons.KIND_DELETE)[0], "入力位置間違い")
        self.assertIsNotNone(judge._ensure_loaded(None, []))                 # 公差表(覚える)
        # 2 つ目の置き場所の中身を変える
        c = sqlite3.connect(self.db2 / "品質基準.sqlite3")
        with c:
            c.execute('DELETE FROM "欠陥名" WHERE rowid IN (SELECT rowid FROM "欠陥名" LIMIT 5)')
        c.close()
        reasons.ensure_table(self.db2 / "伝送用ファイル.sqlite3")
        c = sqlite3.connect(self.db2 / "伝送用ファイル.sqlite3")
        with c:
            c.execute('UPDATE "理由選択" SET "順番" = -1 WHERE "区分" = ? AND "理由" = ?', ("消去", "その他"))
        c.close()
        # 使用パスを変えた直後に、新しい場所から読む
        self.save_path("quality", str(self.db2))
        self.assertEqual(len(datasource.get_table("品質基準", "欠陥名")[1]), n1 - 5)
        self.assertEqual(reasons.choices(reasons.KIND_DELETE)[0], "その他")
        self.assertIn(str(self.db2), str(datasource.resolve("品質基準")))
        self.assertIsNotNone(quality)                                       # (quality の覚えも捨てている)

    def test_lot_path_change_and_reload_button(self) -> None:
        self.register()
        self.fill("DT")
        st = self.s.state
        self.assertEqual(st.v("TaikText").value, "2.0")
        # 2 つ目の台帳では直角度の指示が 3.0
        c = sqlite3.connect(self.db2 / kensa_fixture.file_name("SIKALOTNOW"))
        with c:
            c.execute('UPDATE "仕掛" SET "指示_直角度" = ? WHERE "ﾛｯﾄ番号" = ?', ("3.0", LOT))
        c.close()
        self.save_path("lot", str(self.db2), slot="first")
        self.save_path("lot", str(self.db2), slot="second")
        self.assertEqual(st.v("TaikText").value, "2.0")                     # 取り込んだ値は読み直すまで前のまま
        r = self.c.post("/api/settings/paths/reload", json={}, headers=self.h).get_json()
        self.assertEqual(r["lot"], LOT)
        self.assertEqual(st.v("TaikText").value, "3.0")                     # 「読み込む」で新しい台帳から
        self.assertEqual(st.cells["S1"].value, "15.012")                    # 測定値はそのまま
        self.assertTrue(any("対角 公差: 2.0 → 3.0" in m["text"] for m in r["messages"]), r["messages"])

    def test_reload_without_lot(self) -> None:
        r = self.c.post("/api/settings/paths/reload", json={}, headers=self.h).get_json()
        self.assertEqual(r["lot"], "")
        self.assertIn("検査中のロットはありません", r["messages"][0]["text"])
        page = self.c.get("/settings", headers=self.h).get_data(as_text=True)
        self.assertIn('id="reloadPaths"', page)
