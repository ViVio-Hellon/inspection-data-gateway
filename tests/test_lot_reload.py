"""ロット情報の読み直しと見張り(lot_watch.py・Session.reload_lot / lot_watch。2026-11-07)。

引当だけ・オーダーだけ・仕掛(指示)だけ・品質基準だけ が変わる場合を、一時コピーの DB を書き換えて確かめる。
"""
from __future__ import annotations

import os
import sqlite3
import time

from tests import kensa_fixture
from tests.test_kensa_session import LOT, SessionBase


class LotReloadTest(SessionBase):
    def db(self, name: str):
        return sqlite3.connect(self.db_dir / kensa_fixture.file_name(name))

    def exec(self, name: str, sql: str, args=()) -> None:
        c = self.db(name)
        with c:
            c.execute(sql, args)
        c.close()
        time.sleep(0.01)

    def start_lot(self) -> None:
        self.register()
        self.fill("DT")

    def watch(self) -> list[str]:
        return self.s.lot_watch()["changed"]

    def test_nothing_changed(self) -> None:
        self.start_lot()
        st = self.s.state
        self.assertEqual(self.watch(), [])
        before = st.cells["S1"].value
        r = self.ok(self.s.reload_lot())
        self.assertIn("変わったところはありません", [m.text for m in r.messages][-1])
        self.assertEqual(st.cells["S1"].value, before)               # 測定値はそのまま
        self.assertEqual(st.verdict("S1").value, "OK")               # 判定し直しても OK のまま

    def test_allocation_only(self) -> None:
        self.start_lot()
        cols = [r[1] for r in self.db("SIKAHIKINOW").execute('PRAGMA table_info("仕掛")')]
        i = cols.index("引当番号")
        c = self.db("SIKAHIKINOW")
        with c:
            row = list(c.execute('SELECT * FROM "仕掛" WHERE "ﾛｯﾄ番号" = ?', (LOT,)).fetchone())
            row[i] = "99999999"
            c.execute(f'INSERT INTO "仕掛" VALUES ({",".join("?" * len(row))})', row)
        c.close()
        self.assertEqual(self.watch(), ["引当"])
        w = self.s.lot_watch()
        self.assertFalse(w["new"])                                    # 同じ変化は 2 回目から「新しい」にしない
        r = self.ok(self.s.reload_lot())
        self.assertIn("変わったもの: 引当", "\n".join(m.text for m in r.messages))
        self.assertEqual(self.watch(), [])                             # 読み直したので基準が新しくなる

    def test_order_only(self) -> None:
        self.start_lot()
        cols = [r[1] for r in self.db("SIKAODRNOW").execute('PRAGMA table_info("仕掛")')]
        self.assertIn("合紙", cols)
        hiki = self.db("SIKAHIKINOW").execute('SELECT "受注番号" FROM "仕掛" WHERE "ﾛｯﾄ番号" = ?', (LOT,)).fetchone()[0]
        self.exec("SIKAODRNOW", 'UPDATE "仕掛" SET "合紙" = ? WHERE "受注番号" = ?', ("ﾃｽﾄ合紙", hiki))
        self.assertEqual(self.watch(), ["オーダー"])

    def test_lot_instruction_changes_tolerance(self) -> None:
        self.start_lot()
        st = self.s.state
        self.assertEqual(st.v("TaikText").value, "2.0")
        self.exec("SIKALOTNOW", 'UPDATE "仕掛" SET "指示_直角度" = ? WHERE "ﾛｯﾄ番号" = ?', ("3.0", LOT))
        self.assertEqual(self.watch(), ["仕掛"])
        r = self.ok(self.s.reload_lot())
        text = "\n".join(m.text for m in r.messages)
        self.assertIn("変わったもの: 仕掛", text)
        self.assertIn("対角 公差: 2.0 → 3.0", text)
        self.assertEqual(st.v("TaikText").value, "3.0")
        self.assertEqual(st.cells["S1"].value, "15.012")              # 測定値はそのまま

    def test_quality_only(self) -> None:
        self.start_lot()
        q = self.db_dir / "品質基準.sqlite3"
        t = time.time() + 5
        os.utime(q, (t, t))
        self.assertEqual(self.watch(), ["品質基準"])

    def test_no_lot(self) -> None:
        self.ok(self.s.start())
        self.assertEqual(self.watch(), [])
        r = self.s.reload_lot()
        self.assertIsNotNone(r.stop)

    def test_route(self) -> None:
        from tests import _web
        from ufkensa.kensa import session
        session.reset_for_test(self.s)
        self.addCleanup(session.reset_for_test, None)
        self.start_lot()
        c, h = _web.make_client(), _web.auth()
        self.assertEqual(c.get("/api/kensa/lot-watch", headers=h).get_json()["changed"], [])
        r = c.post("/api/kensa/lot-reload", json={}, headers=h)
        self.assertEqual(r.status_code, 200)
        self.assertIn('id="lotReloadBtn"', c.get("/kensa", headers=h).get_data(as_text=True))
