"""日報を all-tools の日報モードへ渡す CSV(nippou_export.py。2026-11-08)。

all-tools がこの PC にあれば(試験の環境)、日報モードの取り込み(nippou.logic.csv_import.parse_file)で
そのまま読めることも確かめる。無ければその部分は飛ばす。
"""
from __future__ import annotations

import csv
import io
import sys
import unittest
from pathlib import Path

from tests.test_kensa_session import LOT, SessionBase
from ufkensa.kensa import nippou_export as nx
from ufkensa.kensa import paths

ALL_TOOLS_NIPPOU = Path("/home/user/vivio-hellon/all-tools/tools/nippou")


class NippouExportTest(SessionBase):
    def print_lot(self):
        s = self.s
        self.register()
        self.fill("DT")
        self.fill("#L")
        self.strain_visual_height()
        self.ok(s.min_max())
        self.ok(s.print_check())
        s.print_commit()
        return self.ok(s.print_commit({"print": True}))

    def test_written_on_print(self) -> None:
        out = self.tmp / "nippou"
        paths.set_path("nippou_out", str(out) + "/")
        self.addCleanup(paths.reset_path, "nippou_out")
        r = self.print_lot()
        p = Path(r.view["report"]["nippou"])
        self.assertEqual(p.parent, out)
        self.assertEqual(p.name, "集計明細_2026年11月2日_HVC_1直.csv")
        raw = p.read_bytes()
        self.assertTrue(raw.startswith(b"\xef\xbb\xbf"))                 # BOM 付き UTF-8
        rows = list(csv.DictReader(io.StringIO(raw.decode("utf-8-sig"))))
        self.assertEqual(list(rows[0])[:5], ["日付", "ライン", "直", "ページ", "行"])
        row = rows[0]
        self.assertEqual((row["日付"], row["ライン"], row["直"], row["行"]), ("2026年11月2日", "HVC", "1直", "1"))
        self.assertEqual(row["ロット№"], LOT)
        self.assertEqual(row["製品寸法 厚×幅×丈"], "15.000×1252.0×2502.0")
        self.assertEqual(row["検入枚数"], "17")
        self.assertEqual(row["作業者"], "検査員01")
        self.assertEqual(row["包装仕様書No"], "")                         # 梱包No(何個目の梱包か)は入れない

    @unittest.skipUnless((ALL_TOOLS_NIPPOU / "nippou").is_dir(), "all-tools の日報モードがこの環境に無い")
    def test_all_tools_nippou_reads_it(self) -> None:
        out = self.tmp / "nippou"
        paths.set_path("nippou_out", str(out) + "/")
        self.addCleanup(paths.reset_path, "nippou_out")
        p = Path(self.print_lot().view["report"]["nippou"])
        sys.path.insert(0, str(ALL_TOOLS_NIPPOU))
        try:
            from nippou.logic import csv_import
            parsed = csv_import.parse_file(p)
        finally:
            sys.path.remove(str(ALL_TOOLS_NIPPOU))
        self.assertEqual(parsed.missing_columns, [])
        self.assertEqual(parsed.problems, [])
        self.assertEqual(len(parsed.pages), 1)
        page = parsed.pages[0]
        self.assertEqual((page.header.line, page.header.shift), ("HVC", "1直"))
        self.assertEqual(page.details[0].lot, LOT)
        self.assertEqual(page.header.worker, "検査員01")

    def test_helpers(self) -> None:
        self.assertEqual(nx.nippou_date("2026.01.05"), "2026年1月5日")
        self.assertEqual(nx.nippou_line("丸徳"), "中板")
        self.assertIsNone(nx.export(self.store, "HVC", "HVC", "2026.11.02", "1直", folder=str(self.tmp)))  # 行なし


class NippouRouteTest(NippouExportTest):
    def test_written_on_print(self) -> None:          # 親の試験はここでは走らせない
        pass

    test_all_tools_nippou_reads_it = None
    test_helpers = None

    def test_route(self) -> None:
        from tests import _web
        from ufkensa.kensa import session
        out = self.tmp / "nippou2"
        paths.set_path("nippou_out", str(out) + "/")
        self.addCleanup(paths.reset_path, "nippou_out")
        session.reset_for_test(self.s)
        self.addCleanup(session.reset_for_test, None)
        c, h = _web.make_client(), _web.auth()
        r = c.post("/api/maint/history/nippou", json={}, headers=h).get_json()
        self.assertEqual(r["files"], [])                                 # まだ行が無い
        self.print_lot()
        r = c.post("/api/maint/history/nippou", json={}, headers=h).get_json()
        self.assertTrue(r["files"][0].endswith("_HVC_1直.csv"), r)
