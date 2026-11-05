"""保守画面の マスタ管理(ufkensa/kensa/master_admin.py)と 測定機器ID の伝送テスト(maint.device_test)

- 見る: DB の一覧・表の一覧・1 ページ(絞り込み・並び替え・上限)
- 直す・足す・消す: 列の型で確かめる、画面が見ていた行が変わっていたら断る、仕掛は書かない、前後を trace に残す
- 伝送テスト: 検査画面と同じ部品(input_router.maker/conv、instruments.machin_text/micro_juge)の結果
- API: 読むのは誰でも、書くのは編集の鍵(403 locked)
"""
from __future__ import annotations

import shutil
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from tests import kensa_fixture
from tests._web import make_client
from tests.test_screen_maint import MaintBase
from ufkensa import admin_password, auth, trace
from ufkensa.kensa import datasource, maint, master_admin as ma, paths

TABLE = "測定機器ID"


def _rows(db: str) -> list[list[str]]:
    return datasource.get_table(db, TABLE)[1]


class MasterAdminTest(MaintBase):
    def setUp(self) -> None:
        super().setUp()
        # 梱包資材マスタ(アクセス権限)を置いて使用パスを向ける
        self.access = self.tmp / "db" / "梱包資材マスタ.sqlite3"
        c = sqlite3.connect(self.access)
        c.execute('CREATE TABLE アクセス権限 ("管理番号" INTEGER PRIMARY KEY, "ログインID" TEXT, "PC名" TEXT, '
                  '"権限" TEXT NOT NULL, "有効" INTEGER)')
        c.execute("INSERT INTO アクセス権限 VALUES (1, 'u1', 'PC1', 'Administrator', 1)")
        c.execute('CREATE TABLE ほかの表 (a TEXT)')
        c.commit()
        c.close()
        paths.set_path("access_master", str(self.access))
        self.addCleanup(paths.reset_path, "access_master")

    def test_sources_and_tables(self) -> None:
        s = {x["key"]: x for x in ma.list_sources()}
        self.assertTrue(s["quality"]["editable"] and s["denso"]["editable"] and s["access"]["editable"])
        for k in ("SIKALOTNOW", "SIKAHIKINOW", "SIKAODRNOW"):
            self.assertFalse(s[k]["editable"])
            self.assertTrue(s[k]["exists"])
        t = ma.tables("denso")
        self.assertIn(TABLE, [x["name"] for x in t["tables"]])
        self.assertEqual([x["name"] for x in ma.tables("access")["tables"]], ["アクセス権限"])  # ほかの表は出さない
        with self.assertRaises(ma.MasterError):
            ma.tables("nai")

    def test_page_filter_sort_limit(self) -> None:
        p = ma.page("denso", TABLE)
        self.assertEqual(p["total"], 22)
        self.assertEqual(p["rows"][0]["ライン"], "Test_")                  # 既定は rowid の順
        self.assertIn(ma.ROW_KEY, p["rows"][0])
        cols = {c["name"]: c for c in p["columns"]}
        self.assertTrue(cols["厚上"].get("number_text"))
        self.assertEqual(cols["番号"]["kind"], "int")
        p = ma.page("denso", TABLE, query="HVC")
        self.assertEqual(p["total"], 6)
        p = ma.page("denso", TABLE, sort="厚上", sort_dir="desc")
        self.assertEqual(p["rows"][0]["厚上"], "599.00")                   # 数の大きさで(文字の順なら 74.00)
        self.assertEqual(p["rows"][-1]["厚上"], "")                       # 空は最後
        p = ma.page("denso", TABLE, sort="無い列")
        self.assertEqual(p["sort"], "")
        p = ma.page("denso", TABLE, limit=5)
        self.assertEqual(len(p["rows"]), 5)
        self.assertIn("ほか 17件", p["note"])
        p = ma.page("SIKALOTNOW", "仕掛", limit=3)
        self.assertFalse(p["editable"])
        with self.assertRaises(ma.MasterError):
            ma.page("denso", "無い表")

    def test_save_row_writes_and_logs(self) -> None:
        p = ma.page("denso", TABLE, query="0000158327")
        row = p["rows"][0]
        seen = {k: v for k, v in row.items() if k != ma.ROW_KEY}
        with mock.patch.object(trace, "event", wraps=trace.event) as ev:
            out = ma.save_row("denso", TABLE, row[ma.ROW_KEY], {"厚上": "30.00", "機器名": "M1"}, seen)
        self.assertEqual(out["before"], {"厚上": "28.00"})                 # 変わらない列は書かない
        self.assertEqual(out["after"], {"厚上": "30.00"})
        hit = [c for c in ev.call_args_list if c.args[0] == ma.CHANGE_EV]
        self.assertEqual(hit[0].kwargs["data"]["before"], {"厚上": "28.00"})
        self.assertTrue(any(c.args[0] == "proc.end" and c.kwargs.get("proc") == "db.save" for c in ev.call_args_list))
        # 検査画面が読む形(datasource)でも入っている
        self.assertIn("30.00", [r[5] for r in _rows("伝送用ファイル") if r[2] == "0000158327"])
        # 画面が見ていた値が古ければ断る(rowid が別の行に付き直った場合も)
        with self.assertRaises(ma.MasterError) as cm:
            ma.save_row("denso", TABLE, row[ma.ROW_KEY], {"厚上": "31"}, seen)
        self.assertEqual(cm.exception.code, "stale")
        with self.assertRaises(ma.MasterError) as cm:
            ma.save_row("denso", TABLE, row[ma.ROW_KEY], {"厚上": "30.00"})
        self.assertEqual(cm.exception.code, "no_change")

    def test_validation(self) -> None:
        rid = ma.page("denso", TABLE)["rows"][1][ma.ROW_KEY]
        for values, field in (({"厚上": "abc"}, "厚上"), ({"番号": "1.5"}, "番号"), ({"番号": "x"}, "番号"),
                              ({"無い列": "1"}, "無い列")):
            with self.assertRaises(ma.MasterError) as cm:
                ma.save_row("denso", TABLE, rid, values)
            self.assertEqual(cm.exception.field, field)
        ma.save_row("denso", TABLE, rid, {"厚上": "", "番号": "7"})          # 空は「無し」、整数は整数
        p = ma.page("denso", TABLE, query="0000158327")["rows"][0]
        self.assertEqual((p["厚上"], p["番号"]), (None, 7))
        # NOT NULL の列は空にできない
        with self.assertRaises(ma.MasterError) as cm:
            ma.add_row("access", "アクセス権限", {"ログインID": "u2"})
        self.assertEqual(cm.exception.field, "権限")

    def test_add_delete(self) -> None:
        out = ma.add_row("access", "アクセス権限", {"ログインID": "u2", "PC名": "PC2", "権限": "HVC", "有効": "1"})
        self.assertEqual(out["after"]["有効"], 1)
        p = ma.page("access", "アクセス権限")
        self.assertEqual(p["total"], 2)
        new = [r for r in p["rows"] if r["ログインID"] == "u2"][0]
        self.assertEqual(new["管理番号"], 2)                               # INTEGER PRIMARY KEY は自動
        seen = {k: v for k, v in new.items() if k != ma.ROW_KEY}
        ma.delete_row("access", "アクセス権限", new[ma.ROW_KEY], seen)
        self.assertEqual(ma.page("access", "アクセス権限")["total"], 1)
        with self.assertRaises(ma.MasterError) as cm:
            ma.delete_row("access", "アクセス権限", new[ma.ROW_KEY], seen)
        self.assertEqual(cm.exception.code, "stale")
        with self.assertRaises(ma.MasterError) as cm:
            ma.add_row("access", "ほかの表", {"a": "x"})                    # 梱包資材マスタはアクセス権限だけ
        self.assertEqual(cm.exception.code, "not_editable")

    def test_never_writes_sika(self) -> None:
        f = datasource.resolve("SIKALOTNOW")
        before = f.read_bytes()
        rid = ma.page("SIKALOTNOW", "仕掛", limit=1)["rows"][0][ma.ROW_KEY]
        for call in (lambda: ma.save_row("SIKALOTNOW", "仕掛", rid, {"ﾛｯﾄ番号": "X"}),
                     lambda: ma.add_row("SIKAHIKINOW", "仕掛", {"a": "1"}),
                     lambda: ma.delete_row("SIKAODRNOW", "仕掛", rid)):
            with self.assertRaises(ma.MasterError) as cm:
                call()
            self.assertEqual(cm.exception.code, "not_editable")
        self.assertEqual(f.read_bytes(), before)

    def test_busy_db_waits_then_fails_cleanly(self) -> None:
        """ほかが書いている間は busy_timeout まで待ち、だめなら断る(中途半端に書かない)。"""
        path = datasource.resolve("伝送用ファイル")
        other = sqlite3.connect(path, isolation_level=None)
        other.execute("BEGIN IMMEDIATE")
        try:
            with mock.patch.object(ma, "BUSY_TIMEOUT_MS", 200):
                with self.assertRaises(ma.MasterError) as cm:
                    ma.save_row("denso", TABLE, 1, {"厚上": "1"})
            self.assertEqual(cm.exception.code, "write_failed")
        finally:
            other.execute("ROLLBACK")
            other.close()
        self.assertEqual(ma.page("denso", TABLE)["rows"][0]["厚上"], "7")

    def test_history_reads_trace(self) -> None:
        recs = [{"ev": ma.CHANGE_EV, "t": "2026-11-05T10:00:00", "user": "u",
                 "data": {"source": "denso", "label": "伝送用ファイル", "table": TABLE, "action": "直す", "rowid": 1,
                          "before": {"厚上": "7"}, "after": {"厚上": "8"}}},
                {"ev": "proc.end", "data": {}}]
        with mock.patch.object(trace, "read", return_value=recs):
            h = ma.history(source_key="denso", table=TABLE)
            self.assertEqual(len(h), 1)
            self.assertEqual(h[0]["after"], {"厚上": "8"})
            self.assertEqual(ma.history(source_key="quality"), [])


class DeviceTestTest(MaintBase):
    def test_dt_ok_and_ng(self) -> None:
        r = maint.device_test("DT10006+0000021.712M", "L-1", "3.000", "1000", "K126", "S1")
        self.assertEqual((r["kind"], r["raw"], r["value"], r["code"]), ("DT", "0000021.712", "21.712", "10006"))
        self.assertEqual([h["機器名"] for h in r["hits"]], ["M1"])
        self.assertEqual(r["expect"], {"text": "10006", "name": "M1"})
        self.assertTrue(r["judge"]["ok"])
        self.assertEqual(r["report"][0], "L-1 M1")
        r = maint.device_test("DT10001+0000021.712M", "L-1", "3.000", "1000", "K126", "S1")
        self.assertFalse(r["judge"]["ok"])
        self.assertEqual(r["judge"]["text"], "測定するﾏｲｸﾛを正しく選定して下さい M1を使用の事")   # MT:211 と同じ文言
        self.assertTrue(r["others"])                                       # ほかのラインにも 10001 がある

    def test_line_width_and_packets(self) -> None:
        r = maint.device_test("DT10000+0000001.000M", "バランサー", "3", "1000", "K126", "S1")
        self.assertEqual(r["line"], "ﾊﾞﾗﾝｻｰ")
        self.assertEqual([h["機器名"] for h in r["hits"]], ["M2"])          # D13: 幅だけ揃えて照合
        r = maint.device_test("#L395.8", "L-1", "3", "1000", "K126", "S5")
        self.assertEqual((r["kind"], r["value"]), ("#L", "395.8"))
        self.assertIsNone(r["judge"]["ok"])
        self.assertIn("幅", r["st_label"])
        r = maint.device_test("ab", "L-1", "3", "1000", "K126", "S1")
        self.assertFalse(r["accepted"])
        with self.assertRaises(ValueError):
            maint.device_test("DT1", "L-1", st="Z9")

    def test_thickness_not_number(self) -> None:
        r = maint.device_test("DT10006+0000021.712M", "L-1", "", "1000", "K126", "S1")
        self.assertIsNone(r["expect"])
        self.assertTrue(any("厚" in n for n in r["notes"]))


class MasterRouteTest(MaintBase):
    def setUp(self) -> None:
        super().setUp()
        self.client = make_client()
        app = self.client.application
        from app.routes import maint as maint_routes
        if "maint" not in app.blueprints:
            app.register_blueprint(maint_routes.bp)
        self.h = {"X-Tool-Token": app.config["TOKEN"]}
        p = mock.patch.object(auth, "can_edit", return_value=True)
        p.start()
        self.addCleanup(p.stop)
        self.addCleanup(auth.lock_master)

    def post(self, url: str, body: dict):
        return self.client.post(url, json=body, headers=self.h)

    def test_read_open_write_locked(self) -> None:
        auth.lock_master()
        d = self.client.get("/api/maint/master/sources", headers=self.h).get_json()
        self.assertIn("denso", [s["key"] for s in d["sources"]])
        d = self.client.get(f"/api/maint/master/rows?source=denso&table={TABLE}&q=HVC", headers=self.h).get_json()
        self.assertEqual(d["total"], 6)
        r = self.post("/api/maint/master/row", {"source": "denso", "table": TABLE, "action": "save", "rowid": 1,
                                                "values": {"厚上": "8"}})
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "locked"))
        auth.unlock_master(admin_password.DEFAULT_PASSWORD)
        r = self.post("/api/maint/master/row", {"source": "denso", "table": TABLE, "action": "save", "rowid": 1,
                                                "values": {"厚上": "8"}, "seen": {"厚上": "7"}})
        self.assertEqual(r.status_code, 200, r.get_json())
        r = self.post("/api/maint/master/row", {"source": "denso", "table": TABLE, "action": "save", "rowid": 1,
                                                "values": {"厚上": "x"}})
        self.assertEqual((r.status_code, r.get_json()["error"]["field"]), (400, "厚上"))
        r = self.post("/api/maint/master/row", {"source": "SIKALOTNOW", "table": "仕掛", "action": "delete", "rowid": 1})
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "not_editable"))
        r = self.client.get("/api/maint/master/rows?source=denso&table=nai", headers=self.h)
        self.assertEqual(r.status_code, 404)
        self.assertEqual(self.client.get("/api/maint/master/history", headers=self.h).status_code, 200)

    def test_device_api(self) -> None:
        d = self.client.get("/api/maint/device/defaults", headers=self.h).get_json()
        self.assertEqual(d["line"], "HVC")                                 # 検査中でなければこの端末のライン
        d = self.post("/api/maint/device/test", {"text": "DT10001+0000001.000M", "line": "HVC", "atu": "3",
                                                 "hab": "1000", "youto": "K126", "st": "S1"}).get_json()
        self.assertTrue(d["judge"]["ok"])
        r = self.post("/api/maint/device/test", {"text": "x", "line": "HVC", "st": "Q1"})
        self.assertEqual((r.status_code, r.get_json()["error"]["field"]), (400, "st"))

    def test_page_has_tabs(self) -> None:
        from app import shell
        with mock.patch.object(shell, "READY_SCREENS", shell.READY_SCREENS | {"maint"}):
            html = self.client.get("/maint", headers=self.h).get_data(as_text=True)
        for tab in ("dist", "kari", "master", "device", "lookup", "term"):
            self.assertIn(f'data-panel="{tab}"', html)


if __name__ == "__main__":
    unittest.main()
