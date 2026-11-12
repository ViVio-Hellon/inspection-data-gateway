"""xlsx の 1 シートを、見た目のまま HTML の表にする(標準ライブラリだけ。読むだけ)

2026-11-12: 保守 → QF空冷時間 の見本に使う。

【all-tools から何を持ってきたか】
all-tools の 点検表(tools/inspection)の見本は、Excel を COM(cscript + VBScript)で開いて
`CopyPicture` した絵を PNG にする仕組みで、Excel の入った Windows でしか動かない
(標準ライブラリだけでは移せない)。そこで、all-tools の 日報(tools/nippou/nippou/logic/xlsx_sheet.py)が
xlsx を ZIP + XML として読む作法(workbook.xml → rels → worksheets、sharedStrings、
`A1` の列名を数へ、数の文字の短くし方)を下敷きに、見た目に要るもの(styles.xml の
書式・太字・塗り・罫線・寄せ、列の幅・行の高さ、セルの結合)を足した。

できること: セルの結合(rowspan / colspan)・列の幅・行の高さ・隠した行列・太字/斜体/文字の大きさ/色・
塗り(RGB と古い色番号)・罫線(線の有無と太さ)・横/縦の寄せ・折り返し・
数の書式(標準・0・0.00・#,##0・%・日付/時刻・文字 @)。
できないこと: テーマの色・条件付き書式・図・数式の再計算(保存されている値を出す)。
"""
from __future__ import annotations

import datetime as _dt
import html
import re
import zipfile
from pathlib import Path
from typing import Any, Optional
from xml.etree import ElementTree as ET

_M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_REF = re.compile(r"^([A-Z]+)(\d+)$")

#: 組み込みの数の書式(よく使うものだけ)
_BUILTIN_FMT = {0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%",
                11: "0.00E+00", 14: "yyyy/m/d", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM",
                19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "yyyy/m/d h:mm", 49: "@",
                56: "yyyy年m月d日", 57: "yyyy年m月", 58: "m月d日"}
#: 古い色番号(indexed)のうち、よく使うもの
_INDEXED = {8: "000000", 9: "FFFFFF", 10: "FF0000", 11: "00FF00", 12: "0000FF", 13: "FFFF00", 14: "FF00FF",
            15: "00FFFF", 22: "C0C0C0", 23: "808080", 43: "FFFF99", 44: "99CCFF", 45: "FF99CC", 47: "FFCC99",
            50: "CCFFCC", 26: "FFFFCC", 27: "CCFFFF", 41: "CCFFFF", 42: "CCFFCC"}
_BORDER_W = {"thin": "1px solid", "hair": "1px dotted", "dotted": "1px dotted", "dashed": "1px dashed",
             "medium": "2px solid", "mediumDashed": "2px dashed", "thick": "3px solid", "double": "3px double"}


class NotXlsx(Exception):
    """xlsx として開けない・シートが無い。理由の文を持つ。"""


def col_no(name: str) -> int:
    n = 0
    for ch in name:
        n = n * 26 + ord(ch) - 64
    return n


def col_name(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _text(si: Optional[ET.Element]) -> str:
    """文字の置き場 1 つ分(ふりがな rPh は除く)。"""
    if si is None:
        return ""
    out = []
    for ch in si:
        if ch.tag == _M + "t":
            out.append(ch.text or "")
        elif ch.tag == _M + "r":
            out += [t.text or "" for t in ch.iter(_M + "t")]
    return "".join(out)


def _color(el: Optional[ET.Element]) -> str:
    if el is None:
        return ""
    rgb = el.get("rgb")
    if rgb and len(rgb) >= 6:
        return "#" + rgb[-6:]
    idx = el.get("indexed")
    if idx is not None and int(idx) in _INDEXED:
        return "#" + _INDEXED[int(idx)]
    return ""


def _styles(z: zipfile.ZipFile) -> list[dict[str, Any]]:
    """cellXfs の並びで [{"fmt", "css"}]。"""
    if "xl/styles.xml" not in z.namelist():
        return []
    st = ET.fromstring(z.read("xl/styles.xml"))
    fmts = dict(_BUILTIN_FMT)
    for nf in st.iter(_M + "numFmt"):
        fmts[int(nf.get("numFmtId", "0"))] = nf.get("formatCode", "General")
    fonts, fills, borders = [], [], []
    fe = st.find(_M + "fonts")
    for f in (fe if fe is not None else []):
        css = []
        if f.find(_M + "b") is not None and f.find(_M + "b").get("val", "1") not in ("0", "false"):
            css.append("font-weight:700")
        if f.find(_M + "i") is not None and f.find(_M + "i").get("val", "1") not in ("0", "false"):
            css.append("font-style:italic")
        sz = f.find(_M + "sz")
        if sz is not None and sz.get("val") and float(sz.get("val")) != 11:   # 11pt は表の既定
            css.append(f"font-size:{float(sz.get('val')):g}pt")
        c = _color(f.find(_M + "color"))
        if c and c.upper() != "#000000":
            css.append("color:" + c)
        fonts.append(css)
    fl = st.find(_M + "fills")
    for f in (fl if fl is not None else []):
        p = f.find(_M + "patternFill")
        c = ""
        if p is not None and p.get("patternType") not in (None, "none", "gray125"):
            c = _color(p.find(_M + "fgColor")) or _color(p.find(_M + "bgColor"))
        fills.append(["background:" + c] if c else [])
    bl = st.find(_M + "borders")
    for b in (bl if bl is not None else []):
        css = []
        for side in ("left", "right", "top", "bottom"):
            e = b.find(_M + side)
            if e is not None and e.get("style") in _BORDER_W:
                css.append(f"border-{side}:{_BORDER_W[e.get('style')]} {_color(e.find(_M + 'color')) or '#000'}")
        borders.append(css)
    out = []
    xfs = st.find(_M + "cellXfs")
    for x in (xfs if xfs is not None else []):
        css: list[str] = []
        for idx, table in ((x.get("fontId"), fonts), (x.get("fillId"), fills), (x.get("borderId"), borders)):
            if idx is not None and int(idx) < len(table):
                css += table[int(idx)]
        al = x.find(_M + "alignment")
        if al is not None:
            h = al.get("horizontal")
            if h in ("left", "center", "right", "justify"):
                css.append("text-align:" + h)
            elif h == "centerContinuous":
                css.append("text-align:center")
            v = al.get("vertical")
            if v in ("top", "center", "bottom"):
                css.append("vertical-align:" + ("middle" if v == "center" else v))
            if al.get("wrapText") in ("1", "true"):
                css.append("white-space:pre-wrap")
        out.append({"fmt": fmts.get(int(x.get("numFmtId", "0")), "General"), "css": css})
    return out


def _sheet_target(z: zipfile.ZipFile, sheet: Optional[str]) -> tuple[str, str]:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    sheets = list(wb.iter(_M + "sheet"))
    if not sheets:
        raise NotXlsx("シートがありません")
    chosen = sheets[0] if sheet is None else next((s for s in sheets if s.get("name") == sheet), None)
    if chosen is None:
        raise NotXlsx(f"指定シートがありません({sheet})")
    rid = chosen.get(_R + "id")
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    for rel in rels.iter(_PR + "Relationship"):
        if rel.get("Id") == rid:
            t = rel.get("Target", "").lstrip("/")
            return chosen.get("name", ""), t if t.startswith("xl/") else "xl/" + t
    raise NotXlsx("シートの場所が分かりません")


# ------------------------------------------------------------------
# 数の書式
# ------------------------------------------------------------------
_DATE_TOKEN = re.compile(r"[ymdhs]", re.I)


def _strip_fmt(code: str) -> str:
    code = re.sub(r'"[^"]*"', "", code)
    code = re.sub(r"\[[^\]]*\]", "", code)
    return re.sub(r"\\.", "", code)


def _short(x: float) -> str:
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    return f"{x:.10g}"


def format_number(x: float, code: str) -> str:
    """数を Excel の書式で文字にする(よく使う形だけ。分からなければ標準)。"""
    code = (code or "General").split(";")[0]
    bare = _strip_fmt(code)
    if code in ("General", "@") or not bare.strip():
        return _short(x)
    if _DATE_TOKEN.search(bare) and "E" not in bare.upper().replace("E+", ""):
        return _format_date(x, bare)
    pct = "%" in bare
    v = x * 100 if pct else x
    dec = 0
    m = re.search(r"\.([0#]+)", bare)
    if m:
        dec = len(m.group(1))
    s = f"{v:,.{dec}f}" if "," in bare else f"{v:.{dec}f}"
    return s + ("%" if pct else "")


def _format_date(x: float, bare: str) -> str:
    try:
        d = _dt.datetime(1899, 12, 30) + _dt.timedelta(days=float(x))
    except (OverflowError, ValueError):
        return _short(x)
    has_date = re.search(r"[yd]", bare, re.I) or re.search(r"m{3,}", bare, re.I)
    has_time = re.search(r"[hs]", bare, re.I)
    if has_date and has_time:
        return f"{d.year}/{d.month}/{d.day} {d.hour}:{d.minute:02d}"
    if has_time:
        return f"{d.hour}:{d.minute:02d}" + (f":{d.second:02d}" if "s" in bare.lower() else "")
    if "年" in bare or "月" in bare:
        return f"{d.year}年{d.month}月{d.day}日" if "日" in bare else f"{d.year}年{d.month}月"
    return f"{d.year}/{d.month}/{d.day}"


# ------------------------------------------------------------------
# 読む・HTML にする
# ------------------------------------------------------------------
def read_sheet(path: Path, sheet: Optional[str] = None, max_rows: int = 2000) -> dict[str, Any]:
    """{"sheet", "cells": {(r, c): {"text", "css"}}, "merges": [(r1, c1, r2, c2)], "widths": {c: px},
    "heights": {r: px}, "hidden_rows", "hidden_cols", "max_row", "max_col", "cut"}。"""
    try:
        z = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError) as e:
        raise NotXlsx(f"xlsx として開けません: {e}") from e
    with z:
        name, target = _sheet_target(z, sheet)
        styles = _styles(z)
        shared: list[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            shared = [_text(si) for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(_M + "si")]
        sh = ET.fromstring(z.read(target))
    fpr = sh.find(_M + "sheetFormatPr")
    def_w = float(fpr.get("defaultColWidth") or (float(fpr.get("baseColWidth", 8)) + 0.71)) if fpr is not None else 8.43
    def_h = float(fpr.get("defaultRowHeight", 15)) if fpr is not None else 15.0
    widths: dict[int, int] = {}
    hidden_cols: set[int] = set()
    for c in sh.iter(_M + "col"):
        lo, hi = int(c.get("min", "1")), int(c.get("max", "1"))
        for n in range(lo, min(hi, 200) + 1):
            if c.get("width"):
                widths[n] = int(float(c.get("width")) * 7 + 5)
            if c.get("hidden") in ("1", "true"):
                hidden_cols.add(n)
    heights: dict[int, int] = {}
    hidden_rows: set[int] = set()
    cells: dict[tuple[int, int], dict[str, Any]] = {}
    max_r = max_c = 0
    for row in sh.iter(_M + "row"):
        r = int(row.get("r", "0") or 0)
        if r > max_rows:
            continue
        if row.get("ht"):
            heights[r] = round(float(row.get("ht")) * 4 / 3)
        if row.get("hidden") in ("1", "true"):
            hidden_rows.add(r)
        for c in row.iter(_M + "c"):
            m = _REF.match(c.get("r") or "")
            if not m:
                continue
            cr, cc = int(m.group(2)), col_no(m.group(1))
            s = int(c.get("s", "0") or 0)
            st = styles[s] if s < len(styles) else {"fmt": "General", "css": []}
            t = c.get("t", "n")
            v = c.find(_M + "v")
            text = ""
            num = False
            if t == "inlineStr":
                text = _text(c.find(_M + "is"))
            elif v is not None and v.text is not None:
                if t == "s":
                    text = shared[int(v.text)] if int(v.text) < len(shared) else ""
                elif t == "b":
                    text = "TRUE" if v.text == "1" else "FALSE"
                elif t in ("str", "e"):
                    text = v.text
                else:
                    try:
                        text, num = format_number(float(v.text), st["fmt"]), True
                    except ValueError:
                        text = v.text
            if not text and not st["css"]:
                continue
            css = list(st["css"])
            if num and not any(x.startswith("text-align") for x in css):
                css.append("text-align:right")
            cells[(cr, cc)] = {"text": text, "css": css}
            if text or any(x.startswith(("background", "border")) for x in css):
                max_r, max_c = max(max_r, cr), max(max_c, cc)
    merges = []
    for mc in sh.iter(_M + "mergeCell"):
        a, _, b = (mc.get("ref") or "").partition(":")
        ma, mb = _REF.match(a), _REF.match(b or a)
        if ma and mb:
            r1, c1, r2, c2 = int(ma.group(2)), col_no(ma.group(1)), int(mb.group(2)), col_no(mb.group(1))
            if r1 <= max_rows:
                merges.append((r1, c1, min(r2, max_rows), c2))
                if (r1, c1) in cells and cells[(r1, c1)]["text"]:
                    max_r, max_c = max(max_r, r2 if r2 <= max_rows else max_r), max(max_c, c2)
    cut = any(int(row.get("r", "0") or 0) > max_rows for row in sh.iter(_M + "row"))
    return {"sheet": name, "cells": cells, "merges": merges, "widths": widths, "def_width": int(def_w * 7 + 5),
            "heights": heights, "def_height": round(def_h * 4 / 3), "hidden_rows": hidden_rows,
            "hidden_cols": hidden_cols, "max_row": max_r, "max_col": max_c, "cut": cut}


def to_html(data: dict[str, Any], *, row_attr: str = "data-r") -> str:
    """read_sheet の結果を <table class="xl"> に。行・列の見出し(Excel と同じ A B C… / 1 2 3…)付き。
    文字はすべて html.escape する。各行の <tr> に `data-r="行番号"`(画面で当たった行を光らせる)。"""
    mr, mc = data["max_row"], data["max_col"]
    cols = [c for c in range(1, mc + 1) if c not in data["hidden_cols"]]
    covered: set[tuple[int, int]] = set()
    span: dict[tuple[int, int], tuple[int, int]] = {}
    for r1, c1, r2, c2 in data["merges"]:
        rs = sum(1 for r in range(r1, r2 + 1) if r not in data["hidden_rows"])
        cs = sum(1 for c in range(c1, c2 + 1) if c not in data["hidden_cols"])
        span[(r1, c1)] = (rs, cs)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if (r, c) != (r1, c1):
                    covered.add((r, c))
    out = ['<table class="xl"><colgroup><col class="xl__rh">']
    total = 40
    for c in cols:
        w = data["widths"].get(c, data["def_width"])
        total += w
        out.append(f'<col style="width:{w}px">')
    out.append("</colgroup><thead><tr><th class=\"xl__corner\"></th>")
    out += [f"<th>{col_name(c)}</th>" for c in cols]
    out.append("</tr></thead><tbody>")
    for r in range(1, mr + 1):
        if r in data["hidden_rows"]:
            continue
        h = data["heights"].get(r, data["def_height"])
        out.append(f'<tr {row_attr}="{r}" style="height:{h}px"><th>{r}</th>')
        for c in cols:
            if (r, c) in covered:
                continue
            cell = data["cells"].get((r, c))
            attrs = ""
            if (r, c) in span:
                rs, cs = span[(r, c)]
                attrs += (f' rowspan="{rs}"' if rs > 1 else "") + (f' colspan="{cs}"' if cs > 1 else "")
            if cell and cell["css"]:
                attrs += ' style="' + html.escape(";".join(cell["css"])) + '"'
            out.append(f"<td{attrs}>{html.escape(cell['text']) if cell else ''}</td>")
        out.append("</tr>")
    out.append("</tbody></table>")
    return "".join(out).replace('<table class="xl">', f'<table class="xl" style="width:{total}px">', 1)


def render(path: Path, sheet: Optional[str] = None, max_rows: int = 2000) -> dict[str, Any]:
    """{"sheet", "html", "rows", "cols", "cut"}。開けなければ NotXlsx。"""
    d = read_sheet(path, sheet, max_rows)
    return {"sheet": d["sheet"], "html": to_html(d), "rows": d["max_row"], "cols": d["max_col"], "cut": d["cut"]}
