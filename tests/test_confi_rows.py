"""公差外の確認を出す行のマスタ(伝送用ファイル の表「公差外の確認行」。confi_rows.py。2026-11-11)。"""
from __future__ import annotations

import sqlite3
from unittest import mock

from tests.test_kensa_checks import make_uf
from tests.test_kensa_session import SessionBase
from tests.test_screen_maint import MaintRouteTest
from ufkensa.kensa import checks, confi_rows as cr


class ParseTest(SessionBase):
    def test_parse(self) -> None:
        self.assertEqual(cr.parse("薄板", " 4, 2，10"), (2, 4, 10))
        self.assertEqual(cr.parse("それ以外", "14"), (14,))
        self.assertEqual(cr.parse("厚板", ""), ())
        for kind, text in (("薄板", "11"), ("スラブ", "0"), ("厚板", "a"), ("それ以外", "15"), ("薄板", "2,2"),
                           ("薄板", "2,,4"), ("薄板", "-1")):
            with self.assertRaises(cr.Invalid, msg=(kind, text)):
                cr.parse(kind, text)
        with self.assertRaises(cr.Invalid):
            cr.parse("極厚", "1")

    def test_kind_of(self) -> None:
        self.assertEqual([cr.kind_of(w) for w in ("薄板", "スラブ", "厚板", "少枚数", "極厚", "")],
                         ["薄板", "スラブ", "厚板", "厚板", "それ以外", "それ以外"])

    def db(self):
        return self.db_dir / "伝送用ファイル.sqlite3"

    def test_read_does_not_create_and_defaults(self) -> None:
        self.assertEqual(cr.rows_map(), cr.DEFAULTS)
        c = sqlite3.connect(self.db())
        self.assertIsNone(c.execute("SELECT 1 FROM sqlite_master WHERE name = ?", (cr.TABLE,)).fetchone())
        c.close()
        self.assertEqual(cr.rows_map(self.tmp / "無い.sqlite3"), cr.DEFAULTS)

    def test_ensure_save_and_read(self) -> None:
        self.assertTrue(cr.ensure_table())
        self.assertEqual(cr.rows_map(), cr.DEFAULTS)
        got = cr.save({"薄板": "1,3", "スラブ": "", "厚板": "4,10", "それ以外": "14"})
        self.assertEqual(got["薄板"], (1, 3))
        self.assertEqual(cr.rows_map(), {"薄板": (1, 3), "スラブ": (), "厚板": (4, 10), "それ以外": (14,)})
        with self.assertRaises(cr.Invalid):
            cr.save({"薄板": "99"})
        self.assertEqual(cr.rows_map()["薄板"], (1, 3))                # 悪い値は何も書かない

    def test_bad_value_in_table_uses_default(self) -> None:
        cr.ensure_table()
        c = sqlite3.connect(self.db())
        with c:
            c.execute(f'UPDATE "{cr.TABLE}" SET "行" = ? WHERE "区分" = ?', ("x", "厚板"))
        c.close()
        self.assertEqual(cr.rows_map()["厚板"], cr.DEFAULTS["厚板"])


class OutToleranceRowsTest(SessionBase):
    def test_defaults_same_as_vba(self) -> None:
        for tag in ("S4", "S3", "S10", "S7"):
            uf = make_uf(text={"INP.Tag": tag}, ng={"S2"})
            self.assertEqual(checks.out_tolerance(uf), checks.out_tolerance(uf, rows=dict(cr.DEFAULTS)), tag)

    def test_custom_rows(self) -> None:
        uf = make_uf(text={"INP.Tag": "S3"}, ng={"S2"})                # 厚板: 既定では 3 行目で出さない
        self.assertEqual(checks.out_tolerance(uf), [])
        self.assertEqual([(r.posi, r.n) for r in checks.out_tolerance(uf, rows={"厚板": (3,)})], [("S", 2)])
        uf = make_uf(text={"INP.Tag": "S4"}, ng={"S2"})
        self.assertEqual(checks.out_tolerance(uf, rows={"厚板": ()}), [])   # 空 = 出さない

    def test_session_uses_master(self) -> None:
        """検査画面: マスタで 1 行目を足すと、1 行目の入力で公差外の確認が出る。"""
        cr.save({"薄板": "2,4,6,8,10", "スラブ": "2,4,6,8,10", "厚板": "1,4,10", "それ以外": "4,7,10,12,14"})
        self.register()
        r = self.s.receive("DT10001+0000015.500M")                      # 厚板 S1 = 15.500(上限 15.350 を超える)
        self.assertEqual((r.view["confi"]["posi"], r.view["confi"]["n"], r.view["confi"]["page"]), ("S", 1, 0))

    def test_session_default_no_confi_on_row1(self) -> None:
        self.register()
        r = self.s.receive("DT10001+0000015.500M")
        self.assertIsNone(r.view["confi"])


class ConfiRowsRouteTest(MaintRouteTest):
    """保守の API(見るのは誰でも、保存は編集の鍵)。"""

    def setUp(self) -> None:
        super().setUp()
        self.db = self.tmp / "伝送用ファイル.sqlite3"
        sqlite3.connect(self.db).close()
        p = mock.patch.object(cr, "_db_path", lambda: self.db)
        p.start()
        self.addCleanup(p.stop)

    def test_confi_rows_api(self) -> None:
        d = self.client.get("/api/maint/confi_rows", headers=self.h).get_json()
        self.assertEqual([(k["kind"], k["rows"], k["max"]) for k in d["kinds"]],
                         [("薄板", "2,4,6,8,10", 10), ("スラブ", "2,4,6,8,10", 10), ("厚板", "4,10", 10),
                          ("それ以外", "4,7,10,12,14", 14)])
        c = sqlite3.connect(self.db)                                   # 開いたら既定で作る
        self.assertEqual(c.execute(f'SELECT COUNT(*) FROM "{cr.TABLE}"').fetchone()[0], 4)
        c.close()
        r = self.post("/api/maint/confi_rows", {"rows": {"薄板": "1,2", "スラブ": "2", "厚板": "4", "それ以外": "14"}})
        self.assertEqual(r.status_code, 200)
        self.assertEqual(r.get_json()["kinds"][0]["rows"], "1,2")
        r = self.post("/api/maint/confi_rows", {"rows": {"薄板": "1", "スラブ": "11", "厚板": "4", "それ以外": "14"}})
        self.assertEqual((r.status_code, r.get_json()["error"]["field"]), (400, "スラブ"))
        self.assertEqual(self.post("/api/maint/confi_rows", {}).status_code, 400)
        from ufkensa import auth
        auth.lock_master()
        r = self.post("/api/maint/confi_rows", {"rows": {"薄板": "2"}})
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "locked"))
        self.assertEqual(cr.rows_map()["薄板"], (1, 2))
        self.assertIn('data-tab="confi"', self.client.get("/maint", headers=self.h).get_data(as_text=True))


for _name in dir(MaintRouteTest):                                        # 親のテストをここで二重に走らせない
    if _name.startswith("test_"):
        setattr(ConfiRowsRouteTest, _name, None)
del MaintRouteTest
