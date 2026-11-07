"""複数シートの Excel(.xlsx)を、標準ライブラリ(zipfile + XML)だけで書く(2026-11-07 のご判断「A」)。

VBA は保存のたびに必要なシートを 1 つのブックに写して xlsx で保存していた(Input:1070-1110・Cask:600-635・
Mitsubishi:1280-1325)。ここでは同じく 1 ファイルにシートを並べる。

- 値だけ(罫線・色・列幅は付けない。見せる用は PDF)
- パスワードは付けない(標準ライブラリでは暗号化できない。2026-11-07 に「問題ない」とのご判断)
- 数に見える文字は数として、ほかは文字として書く(Excel で開いたときに計算に使えるように)。
  先頭が 0 の番号(例 "07322019")や、"2-3" のような日付に化ける値は文字のまま
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any, Iterable, Sequence
from xml.sax.saxutils import escape

_BAD = re.compile(r"[\[\]:*?/\\]")
_NUM = re.compile(r"^-?(0|[1-9]\d*)(\.\d+)?$")
_CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def sheet_title(name: str, used: set[str]) -> str:
    """Excel のシート名の決まり(31 文字まで・[]:*?/\\ は使えない・重複しない)。"""
    t = _BAD.sub("_", str(name or "Sheet"))[:31] or "Sheet"
    base, i = t, 2
    while t.lower() in used:
        suffix = f"({i})"
        t = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(t.lower())
    return t


def col_letter(c: int) -> str:
    s = ""
    while c > 0:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def _cell(ref: str, v: Any) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, bool):
        return f'<c r="{ref}" t="b"><v>{int(v)}</v></c>'
    if isinstance(v, (int, float)):
        return f'<c r="{ref}"><v>{repr(float(v)) if isinstance(v, float) else v}</v></c>'
    text = _CTRL.sub("", str(v))
    if _NUM.match(text) and not (len(text) > 1 and text.startswith("0") and not text.startswith("0.")):
        return f'<c r="{ref}"><v>{text}</v></c>'
    return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{escape(text)}</t></is></c>'


def _sheet_xml(rows: Iterable[Sequence[Any]]) -> str:
    out = []
    for r, row in enumerate(rows, start=1):
        cells = "".join(_cell(f"{col_letter(c)}{r}", v) for c, v in enumerate(row, start=1))
        if cells:
            out.append(f'<row r="{r}">{cells}</row>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<sheetData>{"".join(out)}</sheetData></worksheet>')


def write_xlsx(path: Path, sheets: Sequence[tuple[str, Sequence[Sequence[Any]]]]) -> Path:
    """sheets = [(シート名, 行の並び)]。行の並びは 1 行目から(空の行も位置を保つ)。"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    used: set[str] = set()
    titles = [sheet_title(n, used) for n, _ in sheets] or ["Sheet1"]
    data = list(sheets) or [("Sheet1", [])]
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" '
          'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
    for i in range(1, len(data) + 1):
        ct.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                  'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    ct.append("</Types>")
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
            'officeDocument" Target="xl/workbook.xml"/></Relationships>')
    wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
          + "".join(f'<sheet name="{escape(t)}" sheetId="{i}" r:id="rId{i}"/>' for i, t in enumerate(titles, 1))
          + "</sheets></workbook>")
    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                         f'relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(data) + 1))
               + "</Relationships>")
    tmp = path.with_name(path.name + ".part")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "".join(ct))
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        for i, (_n, rows) in enumerate(data, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(rows))
    tmp.replace(path)                      # 書きかけのファイルを残さない
    return path


def grid_rows(cells: Iterable[tuple[int, int, Any]]) -> list[list[Any]]:
    """(行, 列, 値) の並び → 1 行目からの行の並び(シートのセルをそのまま並べる)。"""
    vals = {(r, c): v for r, c, v in cells}
    if not vals:
        return []
    last_r = max(r for r, _ in vals)
    last_c = max(c for _, c in vals)
    return [[vals.get((r, c)) for c in range(1, last_c + 1)] for r in range(1, last_r + 1)]
