"""各ラインの UF検査判定.xlsm に貯まったデータを、手元DB に1回だけ取り込む(判断事項 D5)

ラインPCには追加ライブラリを入れないので、xlsm(zip + XML)を標準ライブラリだけで読む。

- 取り込むシート: 保存用・耳用・Lot情報履歴・印刷履歴管理・校正記録・品証検証・Accesslog と、
  ライン名の付くシート(作業保存_<L>・WorkerTime_<L>・検査員_<L>・Daily_<L>・異常ロット用_<L>・不良要因印刷_<L>・
  印刷用_<L>・梱包枚数計算_<L>)、特殊検査の作業中保存
- 値はセルの値をそのまま(文字は文字、数は数)。**日付・時刻の書式のセルは VBA の CStr(Date) と同じ文字**
  (`yyyy/mm/dd H:mm:ss`)にする(VBA がシートから読むと Date になり、文字にすると この形になるため)
- 同じ名前のシートが手元DB にあれば、**消してから入れ直す**(1回きりの取り込み。2回押しても重ならない)
"""
from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterator, Optional

from .store import LocalStore, import_cells

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}

FIXED_SHEETS = ("保存用", "耳用", "Lot情報履歴", "印刷履歴管理", "校正記録", "品証検証", "Accesslog",
                "全ヵ所作業中保存", "ｷｬｽｸ作業中保存")
LINE_PREFIXES = ("作業保存_", "WorkerTime_", "検査員_", "Daily_", "異常ロット用_", "不良要因印刷_", "印刷用_",
                 "梱包枚数計算_")

#: Excel の組み込み書式のうち日付・時刻のもの(14〜22、45〜47、日本語版の 27〜36・50〜58)
_DATE_BUILTIN = set(range(14, 23)) | {45, 46, 47} | set(range(27, 37)) | set(range(50, 59))
_EXCEL_EPOCH = datetime(1899, 12, 30)


def _col_no(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _is_date_format(code: str) -> bool:
    s = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", code.lower())
    return bool(re.search(r"[ymdhs]", s)) and "0" not in s.replace("00", "") or bool(re.search(r"(yy|dd|hh|mm:|:ss)", s))


def cstr_date(serial: float) -> str:
    """Excel のシリアル値 → VBA の CStr(Date)(日本の地域設定)。"""
    d = _EXCEL_EPOCH + timedelta(days=serial)
    d = d.replace(microsecond=0) + (timedelta(seconds=1) if d.microsecond >= 500_000 else timedelta(0))
    has_time = (d.hour, d.minute, d.second) != (0, 0, 0)
    if int(serial) == 0:
        return f"{d.hour}:{d.minute:02d}:{d.second:02d}"
    s = f"{d.year:04d}/{d.month:02d}/{d.day:02d}"
    return s + (f" {d.hour}:{d.minute:02d}:{d.second:02d}" if has_time else "")


@dataclass
class Book:
    path: Path
    sheets: dict[str, str] = field(default_factory=dict)      # シート名 → zip の中のパス
    shared: list[str] = field(default_factory=list)
    date_styles: set[int] = field(default_factory=set)        # 日付書式の cellXfs の番号


def si_text(si: ET.Element) -> str:
    """共有文字列 1 つの本文。**ふりがな(rPh)は読まない**(Excel は「1直」に「チョク」を添えて保存している)。"""
    m = NS["m"]
    parts = []
    for child in si:
        if child.tag == f"{{{m}}}t":
            parts.append(child.text or "")
        elif child.tag == f"{{{m}}}r":
            parts.extend(x.text or "" for x in child.iter(f"{{{m}}}t"))
    return "".join(parts)


def open_book(path: Path) -> Book:
    book = Book(Path(path))
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        targets = {r.get("Id"): r.get("Target") for r in rels.iter(f"{{{NS['rel']}}}Relationship")}
        for s in wb.iter(f"{{{NS['m']}}}sheet"):
            t = (targets.get(s.get(f"{{{NS['r']}}}id")) or "").lstrip("/")
            book.sheets[s.get("name")] = t if t.startswith("xl/") else "xl/" + t
        if "xl/sharedStrings.xml" in z.namelist():
            ss = ET.fromstring(z.read("xl/sharedStrings.xml"))
            book.shared = [si_text(si) for si in ss.iter(f"{{{NS['m']}}}si")]
        if "xl/styles.xml" in z.namelist():
            st = ET.fromstring(z.read("xl/styles.xml"))
            custom = {int(n.get("numFmtId")): n.get("formatCode") or ""
                      for n in st.iter(f"{{{NS['m']}}}numFmt")}
            xfs = st.find(f"{{{NS['m']}}}cellXfs")
            for i, xf in enumerate(xfs if xfs is not None else []):
                fid = int(xf.get("numFmtId") or 0)
                if fid in _DATE_BUILTIN or (fid in custom and _is_date_format(custom[fid])):
                    book.date_styles.add(i)
    return book


def read_cells(book: Book, name: str) -> Iterator[tuple[int, int, object]]:
    """シートのセル(行, 列, 値)。空のセルは返さない。"""
    with zipfile.ZipFile(book.path) as z:
        root = ET.fromstring(z.read(book.sheets[name]))
    for c in root.iter(f"{{{NS['m']}}}c"):
        m = re.match(r"([A-Z]+)(\d+)", c.get("r") or "")
        if not m:
            continue
        r, col = int(m.group(2)), _col_no(m.group(1))
        t = c.get("t", "n")
        v = c.find(f"{{{NS['m']}}}v")
        if t == "inlineStr":
            is_ = c.find(f"{{{NS['m']}}}is")
            val: object = si_text(is_) if is_ is not None else ""
        elif v is None or v.text is None:
            continue
        elif t == "s":
            val = book.shared[int(v.text)]
        elif t in ("str", "e"):
            val = v.text
        elif t == "b":
            val = v.text == "1"
        else:
            num = float(v.text)
            if int(c.get("s") or 0) in book.date_styles:
                val = cstr_date(num)
            else:
                val = int(num) if num.is_integer() and abs(num) < 2**53 else num
        if val == "":
            continue
        yield r, col, val


def targets(book: Book) -> list[str]:
    return [n for n in book.sheets if n in FIXED_SHEETS or n.startswith(LINE_PREFIXES)]


@dataclass
class ImportReport:
    sheets: list[tuple[str, int]] = field(default_factory=list)   # (シート名, セルの数)
    skipped: list[str] = field(default_factory=list)

    def summary(self) -> str:
        return "、".join(f"{n}({c}セル)" for n, c in self.sheets) or "取り込むシートがありませんでした"


def import_book(path: Path, store: Optional[LocalStore] = None) -> ImportReport:
    from .. import trace
    store = store or LocalStore()
    trace.note_used(path, "xlsm")
    with trace.timed("import.xlsm", "xlsm を取り込む", source=str(path), target=str(store.path)) as info:
        rep = _import(path, store)
        info["sheets"] = len(rep.sheets)
        info["cells"] = sum(n for _, n in rep.sheets)
    return rep


def _import(path: Path, store: LocalStore) -> ImportReport:
    book = open_book(path)
    rep = ImportReport()
    for name in book.sheets:
        if name not in targets(book):
            rep.skipped.append(name)
            continue
        store.add_sheet(name)                    # 消してから入れ直す
        n = import_cells(store, name, read_cells(book, name))
        rep.sheets.append((name, n))
    return rep
