"""保守: アイカスのラインコード(共有マスタ)と 出力の確認(2026-11-14)。"""
from __future__ import annotations

import sqlite3
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock

from tests.test_kensa_session import SessionBase
from tests.test_screen_maint import MaintRouteTest
from ufkensa.kensa import icas_lines as il
from ufkensa.kensa import output_check as oc
from ufkensa.kensa import records

VBA_LINCONV = {"L-1": "05", "LVC": "03", "HVC": "04", "ﾊﾞﾗﾝｻｰ": "01", "ﾄｯﾄ": "01", "LS4": "06", "NS1": "07",
               "Test": "AA", "丸徳": "Maru", "どこか": "Maru", "": ""}


class IcasLinesTest(SessionBase):
    def setUp(self) -> None:
        super().setUp()
        il.reset_cache()
        self.db = self.db_dir / "伝送用ファイル.sqlite3"

    def test_defaults_reproduce_vba_and_read_does_not_create(self) -> None:
        from ufkensa.kensa import ainput
        for line, code in VBA_LINCONV.items():
            self.assertEqual(ainput.lin_conv(line), code, line)
        c = sqlite3.connect(self.db)
        self.assertIsNone(c.execute("SELECT 1 FROM sqlite_master WHERE name = ?", (il.TABLE,)).fetchone())
        c.close()

    def test_save_changes_code_and_checks(self) -> None:
        from ufkensa.kensa import ainput
        rows = [{"line": a, "code": b, "note": c} for a, b, c in il.DEFAULTS]
        rows[2]["code"] = "14"                                            # HVC
        rows.append({"line": "丸徳", "code": "09", "note": "号機"})
        il.save(rows)
        il.reset_cache()
        self.assertEqual((ainput.lin_conv("HVC"), ainput.lin_conv("丸徳"), ainput.lin_conv("XYZ")), ("14", "09", "Maru"))
        for bad in ([{"line": "HVC", "code": ""}], [{"line": "HVC", "code": "12345"}], [{"line": "HVC", "code": "あ"}],
                    [{"line": "HVC", "code": "1"}, {"line": "HVC", "code": "2"}], [{"line": "", "code": "1"}]):
            with self.assertRaises(il.Invalid, msg=bad):
                il.save(bad)
        self.assertEqual(ainput.lin_conv("HVC"), "14")                    # 悪い値は何も書かない
        il.save([{"line": "HVC", "code": "04"}])                          # (その他) は無くても足す
        self.assertEqual(il.code_for("丸徳"), "Maru")
        self.assertTrue(il.history())

    def test_unreadable_uses_last_known(self) -> None:
        il.save([{"line": "HVC", "code": "44"}])
        with mock.patch.object(il.safe_io, "connect_existing", side_effect=sqlite3.OperationalError("locked")):
            self.assertEqual(il.rows(fresh=True)[0], ("HVC", "44", ""))


class OutputCheckTest(SessionBase):
    def test_measure_db_reads_last_rows_of_line(self) -> None:
        db = self.tmp / "梱包検査ﾃﾞｰﾀ.sqlite3"
        records.ensure_measure_db(db)
        c = sqlite3.connect(db)
        cols = records.MEASURE_TABLES["検査ﾃﾞｰﾀ"]
        with c:
            for i, line in enumerate(("HVC", "LVC", "HVC")):
                c.execute(f'INSERT INTO "検査ﾃﾞｰﾀ" ("{cols[0]}", "{cols[1]}") VALUES (?, ?)', (line, f"v{i}"))
        c.close()
        before = db.stat().st_mtime
        d = oc.measure_db("HVC", db)
        t = d["tables"][0]
        self.assertEqual((t["table"], t["count"], len(t["rows"])), ("検査ﾃﾞｰﾀ", 3, 2))
        self.assertEqual(t["rows"][0][2], "v2")                            # 新しい順
        self.assertEqual(db.stat().st_mtime, before)                       # 書かない
        self.assertIn("見つかりません", oc.measure_db("HVC", self.tmp / "nai.sqlite3")["problem"])

    def test_listing_and_reading_csv(self) -> None:
        folder = self.tmp / "out"
        folder.mkdir()
        (folder / "HVC_履歴CSV.csv").write_bytes("ロット,厚\r\nLOT1,15.0\r\n".encode("cp932"))
        pl = oc._one("x", folder, ("HVC_履歴CSV.csv", "HVC_耳用履歴CSV.csv"))
        self.assertEqual([f["name"] for f in pl["files"]], ["HVC_履歴CSV.csv"])
        self.assertTrue(oc.allowed(pl["files"][0]["path"], [pl]))
        self.assertFalse(oc.allowed(str(self.tmp / "other.csv"), [pl]))
        t = oc.read_text(pl["files"][0]["path"])
        self.assertEqual((t["encoding"], t["rows"][1]), ("cp932", ["LOT1", "15.0"]))
        miss = oc._scan("y", self.tmp / "nai", ("*.pdf",))
        self.assertIn("フォルダがありません", miss["problem"])
        self.assertFalse((self.tmp / "nai").exists())                      # 作らない


class IcasOutputsRouteTest(MaintRouteTest):
    def setUp(self) -> None:
        super().setUp()
        il.reset_cache()
        self.db = self.tmp / "伝送用ファイル.sqlite3"
        sqlite3.connect(self.db).close()
        p = mock.patch.object(il, "_db_path", lambda: self.db)
        p.start()
        self.addCleanup(p.stop)

    def test_icas_api(self) -> None:
        d = self.client.get("/api/maint/icas_lines", headers=self.h).get_json()
        self.assertEqual({r["line"]: r["code"] for r in d["rows"]}["HVC"], "04")
        rows = d["rows"]
        rows[2]["code"] = "24"
        r = self.post("/api/maint/icas_lines", {"rows": rows})
        self.assertEqual(r.status_code, 200)
        self.assertEqual(il.code_for("HVC"), "24")
        r = self.post("/api/maint/icas_lines", {"rows": [{"line": "HVC", "code": "!!"}]})
        self.assertEqual((r.status_code, r.get_json()["error"]["field"]), (400, "HVC"))
        from ufkensa import auth
        auth.lock_master()
        r = self.post("/api/maint/icas_lines", {"rows": rows})
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "locked"))
        page = self.client.get("/maint", headers=self.h).get_data(as_text=True)
        self.assertIn('data-tab="icas"', page)
        self.assertIn('data-tab="outputs"', page)

    def test_outputs_api_only_listed_files(self) -> None:
        from app.routes import maint as m
        folder = self.tmp / "csvlocal"
        folder.mkdir()
        f = folder / "HVC_履歴CSV.csv"
        f.write_bytes(b"a,b\r\n")
        places = [oc._one("x", folder, ("HVC_履歴CSV.csv",))]
        with mock.patch.object(m, "_output_places", lambda kind: (places, {"line": "HVC", "lname": "HVC"})):
            d = self.client.get("/api/maint/outputs?kind=csv", headers=self.h).get_json()
            self.assertEqual(d["places"][0]["files"][0]["name"], "HVC_履歴CSV.csv")
            r = self.client.get(f"/api/maint/outputs/text?kind=csv&path={f}", headers=self.h)
            self.assertEqual(r.get_json()["rows"], [["a", "b"]])
            r = self.client.get(f"/api/maint/outputs/text?kind=csv&path={self.tmp / 'x.csv'}", headers=self.h)
            self.assertEqual(r.status_code, 404)
            self.assertEqual(self.client.get("/api/maint/outputs?kind=zzz", headers=self.h).status_code, 400)


for _name in dir(MaintRouteTest):
    if _name.startswith("test_"):
        setattr(IcasOutputsRouteTest, _name, None)
del MaintRouteTest

if __name__ == "__main__":
    unittest.main()
