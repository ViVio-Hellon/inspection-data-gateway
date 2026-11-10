"""2026-11-10: 理由の履歴(ライン別 Access.log の「理由選択」を読む)。"""
from __future__ import annotations

import shutil
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from ufkensa.kensa import legacy_log, paths, reason_log


class ReasonLogTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="ufkensa_reason_"))
        paths.set_path("history", str(self.tmp / "検査履歴") + "/")
        paths.set_path("app_log", str(self.tmp / "app") + "/")
        (self.tmp / "app").mkdir()
        self.addCleanup(lambda: shutil.rmtree(self.tmp, ignore_errors=True))
        self.addCleanup(paths.reset_path, "history")
        self.addCleanup(paths.reset_path, "app_log")

    def write(self) -> None:
        def app(msg: str, at: datetime, pc: str = "PC-A") -> None:
            legacy_log.write_log(msg, "", now=at, computer=pc)

        def line(msg: str, at: datetime, pc: str = "PC-A") -> None:
            legacy_log.write_line_log("HVC", msg, now=at, computer=pc)

        app("Lot登録|LOTA01|HVC", datetime(2026, 11, 9, 8, 0, 0))
        line("公差外:S1_15.2:理由選択:再測定,キズ, もう一度:検査員01", datetime(2026, 11, 9, 8, 5, 0))
        line("*UFConfi.TB1入力:15.0", datetime(2026, 11, 9, 8, 5, 10))
        app("Lot登録|LOTB02|HVC", datetime(2026, 11, 10, 9, 0, 0))
        line("U2_15.012:理由選択:その他, 測り直し:検査員02", datetime(2026, 11, 10, 9, 1, 0))
        app("Lot登録|LOTC03|HVC", datetime(2026, 11, 10, 9, 0, 30), pc="PC-B")

    def test_file_is_cp932_crlf_and_entries_newest_first(self) -> None:
        self.write()
        f = reason_log.files("HVC")
        raw = legacy_log.local_path(f["access"]).read_bytes()
        self.assertIn("理由選択".encode("cp932"), raw)
        self.assertIn(b"\r\n", raw)
        r = reason_log.read_reasons("HVC")
        self.assertEqual(r["error"], "")
        self.assertTrue(r["access"].endswith("HVC_Log\\Access.log"), r["access"])
        self.assertTrue(r["measured"].endswith("HVC_Log\\Measured.log"))
        self.assertEqual(r["total"], 2)
        new, old = r["entries"]
        self.assertEqual((new["at"], new["position"], new["value"], new["reasons"], new["worker"], new["kind"]),
                         ("2026/11/10 09:01:00", "U2", "15.012", "その他, 測り直し", "検査員02", "消去・理由"))
        self.assertEqual(new["lot"], "LOTB02")          # 同じ PC の直前の Lot登録(PC-B の LOTC03 ではない)
        self.assertEqual((old["kind"], old["position"], old["reasons"], old["lot"], old["pc"]),
                         ("公差外の確認", "S1", "再測定,キズ, もう一度", "LOTA01", "PC-A"))

    def test_filters(self) -> None:
        self.write()
        self.assertEqual(reason_log.read_reasons("HVC", date_from="2026-11-10")["total"], 1)
        self.assertEqual(reason_log.read_reasons("HVC", date_to="2026-11-09")["total"], 1)
        r = reason_log.read_reasons("HVC", lot="lota")
        self.assertEqual([e["lot"] for e in r["entries"]], ["LOTA01"])

    def test_missing_file_says_where(self) -> None:
        r = reason_log.read_reasons("LVC")
        self.assertIn("ありません", r["error"])
        self.assertIn("LVC_Log", r["access"])
        self.assertEqual(r["entries"], [])

    def test_unreadable_bytes_are_tolerated(self) -> None:
        f = reason_log.files("HVC")
        p = legacy_log.local_path(f["access"])
        p.parent.mkdir(parents=True)
        p.write_bytes(b"\x81\x00garbage\r\n" + "2026/11/10 10:00:00 PC  S3_1:理由選択:その他, x:誰\r\n".encode("cp932"))
        r = reason_log.read_reasons("HVC")
        self.assertEqual(r["total"], 1)

    def test_settings_row_shows_targets(self) -> None:
        row = next(r for r in paths.overview() if r["key"] == "app_log")
        self.assertTrue(row["targets"]["app_access"].endswith("Access.log"))
        self.assertIn("_Log\\Access.log", row["targets"]["line_access"])

    def test_api(self) -> None:
        self.write()
        from app import create_app
        app = create_app("main", token="t", port=8899)
        app.config["READY"] = True
        c = app.test_client()
        res = c.get("/api/history/reasons?line=HVC", headers={"X-Tool-Token": "t"})
        self.assertEqual(res.status_code, 200, res.data[:300])
        body = res.get_json()
        self.assertEqual(body["total"], 2)
        self.assertTrue(any(l["key"] == "HVC" for l in body["lines"]))


if __name__ == "__main__":
    unittest.main()
