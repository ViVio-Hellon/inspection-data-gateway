"""保守 → 日報(all-tools の日報の紙と同じ見せ方。2026-11-14)。"""
from __future__ import annotations

import unittest
from unittest import mock

from tests.test_kensa_session import SessionBase
from tests.test_screen_maint import MaintRouteTest
from ufkensa.kensa import nippou_export as ne
from ufkensa.kensa import nippou_print as npr


def _row(lot: str, cnt: str = "10", wei: str = "1200.5", **kw: str) -> dict:
    return {"ロット№": lot, "材・調質": "5052-H34", "実績合計 枚数": cnt, "実績合計 重量": wei, **kw}


class PaperTest(unittest.TestCase):
    def test_paper_shape_like_all_tools(self) -> None:
        pages = npr.paginate([_row("A1"), _row("A2", "5", "300")], report_date="2026年11月14日", line="HVC",
                             shift="1直", worker="山田")
        self.assertEqual((len(pages), pages[0].page, [r["行"] for r in pages[0].rows]), (1, "1", ["1", "2"]))
        h = npr.build_html(pages, generated_at="x")
        self.assertIn("梱　包　実　績　日　報　表", h)
        self.assertIn("<th>作業日・直</th><td>2026年11月14日1直</td>", h)
        self.assertIn('<th colspan="4">梱包作業時間</th>', h)
        self.assertIn('<th rowspan="2">ロット№</th>', h)
        self.assertEqual(h.count("<tr><td>"), 12 + 1)               # 12 行いつも出す + 合計の行
        self.assertIn('onclick="window.print()">印刷する</button>', h)
        self.assertIn("@page { size: A4 landscape; margin: 0; }", h)
        self.assertIn("1500.5 Kg", h)
        self.assertIn("1.50 T", h)
        self.assertNotIn("作業停止④", h)                             # 使っていなければ紙のまま

    def test_more_than_12_rows_makes_pages_and_added_stops(self) -> None:
        rows = [_row(f"L{i}") for i in range(14)]
        rows[13]["作業停止④ 記号"] = "ア"
        pages = npr.paginate(rows, report_date="d", line="HVC", shift="2直", worker="")
        self.assertEqual([p.page for p in pages], ["1", "2"])
        h = npr.build_html(pages)
        self.assertEqual(h.count('<section class="sheet">'), 2)
        self.assertIn("全2ページ", h)
        self.assertIn("作業停止④", h)                                 # 使ったページだけ足す
        self.assertIn("該当するデータがありません", npr.build_html([]))

    def test_csv_roundtrip_uses_page_and_row(self) -> None:
        import tempfile
        from pathlib import Path
        d = Path(tempfile.mkdtemp())
        text = ",".join(ne.NIPPOU_FIELDS) + "\r\n"
        vals = {f: "" for f in ne.NIPPOU_FIELDS}
        vals.update({"日付": "2026年11月14日", "ライン": "LVC", "直": "3直", "ページ": "1", "行": "5", "ロット№": "X9",
                     "作業者": "佐藤"})
        text += ",".join(vals[f] for f in ne.NIPPOU_FIELDS) + "\r\n"
        (d / "a.csv").write_bytes(text.encode("utf-8-sig"))
        pages = npr.paginate(npr.read_csv(d / "a.csv"), report_date="", line="", shift="", worker="")
        h = npr.build_html(pages)
        self.assertIn("<td>2026年11月14日3直</td>", h)
        self.assertIn("<tr><td>5</td><td>X9</td>", h)                # 行番号の位置に置く


class NippouRouteTest(MaintRouteTest):
    def test_api_and_paper(self) -> None:
        from app.routes import maint as m
        pages = npr.paginate([_row("Q1")], report_date="2026年11月14日", line="HVC", shift="1直", worker="")
        info = {"date": "2026年11月14日", "shift": "1直", "line": "HVC", "rows": 1}
        out = self.tmp / "nip"
        out.mkdir()
        (out / "集計明細_2026年11月14日_HVC_1直.csv").write_bytes(
            (",".join(ne.NIPPOU_FIELDS) + "\r\n").encode("utf-8-sig"))
        with mock.patch.object(m, "_nippou_now", lambda: (pages, info)), mock.patch.object(ne, "out_dir", lambda: out):
            d = self.client.get("/api/maint/nippou", headers=self.h).get_json()
            self.assertEqual((d["now"]["rows"], d["files"][0]["name"]), (1, "集計明細_2026年11月14日_HVC_1直.csv"))
            h = self.client.get("/maint/nippou/paper?src=now", headers=self.h).get_data(as_text=True)
            self.assertIn("<td>Q1</td>", h)
            h = self.client.get("/maint/nippou/paper?src=file&name=../x.csv", headers=self.h).get_data(as_text=True)
            self.assertIn("一覧にありません", h)
            h = self.client.get("/maint/nippou/paper?src=file&name=" + d["files"][0]["name"],
                                headers=self.h).get_data(as_text=True)
            self.assertIn("該当するデータがありません", h)
        self.assertIn('data-tab="nippou"', self.client.get("/maint", headers=self.h).get_data(as_text=True))


for _name in dir(MaintRouteTest):
    if _name.startswith("test_"):
        setattr(NippouRouteTest, _name, None)
del MaintRouteTest

if __name__ == "__main__":
    unittest.main()
