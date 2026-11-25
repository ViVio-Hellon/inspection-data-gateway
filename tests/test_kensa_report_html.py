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


class EdgeProfileTest(unittest.TestCase):
    """2026-11-26(現場の声「PDF を作れませんでした(終了コード 0)」): 開いている Edge に仕事を渡さないよう、
    PDF ごとに別のプロファイル(--user-data-dir)で動かす。使ったプロファイルは消す。"""

    @unittest.skipIf(os.name == "nt", "試しの Edge はシェルの台本")
    def test_each_pdf_gets_own_profile(self):
        import sys
        d = Path(tempfile.mkdtemp())
        log = d / "args.txt"
        fake = d / "msedge"
        fake.write_text(f"""#!{sys.executable}
import sys, pathlib
a = sys.argv[1:]
open({str(log)!r}, "a").write(" ".join(a) + "\\n")
ud = [x.split("=", 1)[1] for x in a if x.startswith("--user-data-dir=")]
out = [x.split("=", 1)[1] for x in a if x.startswith("--print-to-pdf=")][0]
if not ud:                     # 開いている Edge に渡したのと同じ: 何も作らず 0 で終わる
    sys.exit(0)
if "ng" in out:
    sys.stderr.write("ERROR: 印刷できません\\n"); sys.exit(0)
pathlib.Path(out).write_bytes(b"%PDF-1.4")
""", encoding="utf-8")
        fake.chmod(0o755)
        h = d / "a.html"
        h.write_text("<p>x</p>", encoding="utf-8")
        p1 = report_html.make_pdf(h, d / "検査_1.pdf", str(fake))
        p2 = report_html.make_pdf(h, d / "検査_2.pdf", str(fake))
        self.assertTrue(p1.is_file() and p2.is_file())
        dirs = [x.split("--user-data-dir=")[1].split(" ")[0] for x in log.read_text().splitlines()]
        self.assertEqual(len(set(dirs)), 2)                              # 毎回 別
        self.assertFalse(any(Path(x).exists() for x in dirs))            # 使い終わったら消す
        with self.assertRaises(report_html.PdfError) as e:
            report_html.make_pdf(h, d / "ng.pdf", str(fake))
        self.assertIn("印刷できません", str(e.exception))                  # Edge の言い分も出す


if __name__ == "__main__":
    unittest.main()
