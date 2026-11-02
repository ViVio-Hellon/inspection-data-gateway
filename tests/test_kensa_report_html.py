"""帳票モデル → 印刷用 HTML(D8)と PDF の作り方(D7)"""
from __future__ import annotations

import os
import unittest
from pathlib import Path
import tempfile

import tests.test_kensa_report_model as T
from ufkensa.kensa import report_html


class RenderTest(T.Base):
    def test_inspection_sheet(self):
        res = self.run_print(T.make_snap(ken_hbox=["200", "226", "", "", "", "", ""]))
        html = report_html.render(res.report, title="検査成績票")
        self.assertIn("@page { size: A4", html)
        self.assertIn(res.report.text("C2"), html)                  # LotNo
        self.assertIn("border-", html)                               # 罫線
        self.assertIn(res.report.page.get("footer_right", ""), html)  # フッター(終了時刻)
        d = report_html.render(res.defect, title="不良要因票")
        self.assertIn("<table>", d)

    def test_save_without_edge_writes_html_only(self):
        res = self.run_print()
        out = report_html.save_html_and_pdf(res.report, Path(tempfile.mkdtemp()), "H7372A0_検査成績票", None)
        self.assertEqual(set(out), {"html"})
        self.assertTrue(Path(out["html"]).is_file())

    def test_missing_edge(self):
        with self.assertRaises(report_html.PdfError):
            report_html.make_pdf(Path("x.html"), Path("x.pdf"), "/無い/msedge.exe")


if __name__ == "__main__":
    unittest.main()
