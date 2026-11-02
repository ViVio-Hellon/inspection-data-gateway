"""各ラインの xlsm からの取り込み(D5)。実データ(氏名が入る)は使わず、小さな xlsx を組み立てて確かめる。"""
from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

from ufkensa.kensa import xlsm_import
from ufkensa.kensa.store import LocalStore

M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def make_book(path: Path) -> Path:
    sheets = {"校正記録": "sheet1.xml", "保存用": "sheet2.xml", "Sheet9": "sheet3.xml", "作業保存_LVC": "sheet4.xml"}
    wb = "".join(f'<sheet name="{n}" sheetId="{i}" r:id="rId{i}"/>' for i, n in enumerate(sheets, 1))
    rels = "".join(f'<Relationship Id="rId{i}" Type="x" Target="worksheets/{f}"/>' for i, f in enumerate(sheets.values(), 1))
    shared = ('<sst xmlns="%s"><si><t>1直</t><rPh sb="1" eb="2"><t>チョク</t></rPh></si>'
              '<si><r><t>検査</t></r><r><t>員A</t></r></si><si><t>H7372B0</t></si></sst>' % M)
    styles = ('<styleSheet xmlns="%s"><numFmts count="1"><numFmt numFmtId="176" formatCode="yyyy/m/d h:mm"/></numFmts>'
              '<cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="21"/><xf numFmtId="176"/></cellXfs></styleSheet>' % M)

    def sheet(cells: str) -> str:
        return f'<worksheet xmlns="{M}"><sheetData>{cells}</sheetData></worksheet>'

    with zipfile.ZipFile(path, "w") as z:
        z.writestr("xl/workbook.xml", f'<workbook xmlns="{M}" xmlns:r="{R}"><sheets>{wb}</sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels",
                   f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>')
        z.writestr("xl/sharedStrings.xml", shared)
        z.writestr("xl/styles.xml", styles)
        z.writestr("xl/worksheets/sheet1.xml", sheet(
            '<row r="2"><c r="A2" t="str"><v>2026.10.13</v></c><c r="C2" s="2"><v>0.41950231481481487</v></c>'
            '<c r="D2" t="s"><v>0</v></c><c r="E2"><v>1999.9</v></c><c r="F2" t="s"><v>1</v></c></row>'))
        z.writestr("xl/worksheets/sheet2.xml", sheet(
            '<row r="1"><c r="A1" t="inlineStr"><is><t>検査開始終了</t></is></c></row>'
            '<row r="2"><c r="B2" t="s"><v>2</v></c><c r="C2"><v>3</v></c><c r="D2" s="3"><v>46176.40790509259</v></c></row>'))
        z.writestr("xl/worksheets/sheet3.xml", sheet('<row r="1"><c r="A1"><v>1</v></c></row>'))
        z.writestr("xl/worksheets/sheet4.xml", sheet('<row r="2"><c r="A2" s="1"><v>46176</v></c></row>'))
    return path


class XlsmImportTest(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.tmp = Path(tmp.name)
        self.book = make_book(self.tmp / "UF検査判定.xlsm")
        self.store = LocalStore(self.tmp / "local.sqlite3")
        self.addCleanup(self.store.close)

    def test_targets(self):
        b = xlsm_import.open_book(self.book)
        self.assertEqual(xlsm_import.targets(b), ["校正記録", "保存用", "作業保存_LVC"])   # Sheet9 は取り込まない

    def test_values_dates_and_phonetic(self):
        rep = xlsm_import.import_book(self.book, self.store)
        self.assertIn("Sheet9", rep.skipped)
        k = self.store.sheet("校正記録")
        self.assertEqual(k.get(2, 1), "2026.10.13")              # 文字のまま
        self.assertEqual(k.get(2, 3), "10:04:05")                # 時刻だけの書式 → CStr(Date) の時刻
        self.assertEqual(k.get(2, 4), "1直")                     # ふりがなは読まない
        self.assertEqual(k.get(2, 5), 1999.9)
        self.assertEqual(k.get(2, 6), "検査員A")                 # 書式つきの文字(r)をつなぐ
        h = self.store.sheet("保存用")
        self.assertEqual((h.get(2, 2), h.get(2, 3)), ("H7372B0", 3))
        self.assertEqual(h.get(2, 4), "2026/06/03 9:47:23")      # 独自の日付書式
        self.assertEqual(self.store.sheet("作業保存_LVC").get(2, 1), "2026/06/03")   # 日付だけ

    def test_twice_does_not_duplicate(self):
        xlsm_import.import_book(self.book, self.store)
        rep = xlsm_import.import_book(self.book, self.store)
        self.assertEqual(dict(rep.sheets)["保存用"], 3)              # 見出しの行は数えない


if __name__ == "__main__":
    unittest.main()
