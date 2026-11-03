"""帳票モデル(report_model.ReportModel)→ 印刷用の HTML、と HTML → PDF(Edge のヘッドレス印刷。D7)

VBA は原紙シートに書いて Excel で PrintOut し、同じシートを PDF にしていた。ここでは
帳票モデル(セルの値・書式済みの文字・罫線・列幅・行の高さ・結合・用紙設定)をそのまま HTML の表にする。
ブラウザで開けば印刷(D8)、Edge のヘッドレス印刷にかければ PDF(D7)。

- 列幅: Excel の幅(文字数)→ px は `幅 × 7 + 5`(既定の 11pt フォントのときの Excel の換算)
- 行の高さ: pt のまま
- 罫線: continuous / dash / dot / double × hairline / thin / medium / thick
- 用紙: A4・向き・余白(inch)。倍率(zoom)は CSS の zoom、「1 ページに収める」は印刷時に縮める
"""
from __future__ import annotations

import html
import os
import re
import subprocess
from pathlib import Path
from typing import Optional

from .report_model import ReportModel

_WEIGHT_PX = {"hairline": 1, "thin": 1, "medium": 2, "thick": 3, "": 1}
_STYLE = {"continuous": "solid", "dash": "dashed", "dot": "dotted", "double": "double", "none": "none"}
DEFAULT_COL_WIDTH = 8.43
DEFAULT_ROW_HEIGHT = 18.75


def _col_no(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _split_ref(ref: str) -> tuple[int, int]:
    m = re.match(r"\$?([A-Z]+)\$?(\d+)", ref)
    return int(m.group(2)), _col_no(m.group(1))


def _border(side: list) -> str:
    style, weight = (side + ["", ""])[:2]
    if not style or style == "none":
        return "none"
    px = 3 if style == "double" else _WEIGHT_PX.get(weight, 1)
    return f"{px}px {_STYLE.get(style, 'solid')} #000"


def _width_px(chars: float) -> int:
    return int(round(chars * 7 + 5))


def render(model: ReportModel, *, title: str = "") -> str:
    """帳票 1 枚の HTML(完全な文書)。"""
    if model.print_area:
        r1, c1, r2, c2 = model.print_area
    elif model.cells:
        r1 = min(c["row"] for c in model.cells)
        c1 = min(c["col"] for c in model.cells)
        r2 = max(c["row"] for c in model.cells)
        c2 = max(c["col"] for c in model.cells)
    else:
        r1 = c1 = r2 = c2 = 1
    cells = {(c["row"], c["col"]): c for c in model.cells}
    spans: dict[tuple[int, int], tuple[int, int]] = {}
    covered: set[tuple[int, int]] = set()
    for m in model.merges:
        a, b = m.split(":")
        (ra, ca), (rb, cb) = _split_ref(a), _split_ref(b)
        spans[(ra, ca)] = (rb - ra + 1, cb - ca + 1)
        for r in range(ra, rb + 1):
            for c in range(ca, cb + 1):
                if (r, c) != (ra, ca):
                    covered.add((r, c))
    widths = {_col_no(k): v for k, v in model.col_widths.items()}
    page = model.page or {}
    orient = page.get("orientation", "portrait")
    mg = page.get("margins_inch") or {}
    zoom = page.get("zoom") or 100

    cols = "".join(f'<col style="width:{_width_px(widths.get(c, DEFAULT_COL_WIDTH))}px">' for c in range(c1, c2 + 1))
    rows = []
    for r in range(r1, r2 + 1):
        h = model.row_heights.get(r, DEFAULT_ROW_HEIGHT)
        tds = []
        for c in range(c1, c2 + 1):
            if (r, c) in covered:
                continue
            x = cells.get((r, c))
            attrs = ""
            if (r, c) in spans:
                rs, cs = spans[(r, c)]
                attrs = (f' rowspan="{rs}"' if rs > 1 else "") + (f' colspan="{cs}"' if cs > 1 else "")
            if x is None:
                tds.append(f"<td{attrs}></td>")
                continue
            b = x.get("borders") or {}
            st = [f"border-{s}:{_border(b.get(s, ['none', '']))}" for s in ("top", "bottom", "left", "right")]
            st.append(f"font-size:{x.get('font_size') or 11}pt")
            if x.get("bold"):
                st.append("font-weight:700")
            if x.get("red"):
                st.append("color:#c00")
            ha = {"center": "center", "right": "right", "left": "left"}.get(x.get("h_align") or "")
            if not ha:
                ha = "right" if isinstance(x.get("value"), (int, float)) and not isinstance(x.get("value"), bool) else "left"
            st.append(f"text-align:{ha}")
            va = {"top": "top", "center": "middle", "bottom": "bottom"}.get(x.get("v_align") or "", "middle")
            st.append(f"vertical-align:{va}")
            if x.get("wrap"):
                st.append("white-space:pre-wrap")
            text = html.escape(x.get("text") or "")
            tds.append(f'<td{attrs} style="{";".join(st)}">{text}</td>')
        rows.append(f'<tr style="height:{h}pt">{"".join(tds)}</tr>')

    def hf(key: str) -> str:
        return html.escape(str(page.get(key, "") or ""))

    return f"""<!doctype html>
<html lang="ja"><head><meta charset="utf-8"><title>{html.escape(title or model.kind)}</title>
<style>
@page {{ size: A4 {orient}; margin: {mg.get('top', 0.4)}in {mg.get('right', 0.3)}in {mg.get('bottom', 0.4)}in {mg.get('left', 0.3)}in; }}
body {{ margin: 0; font-family: "Meiryo UI", "MS PGothic", sans-serif; color: #000; background: #fff; }}
.sheet {{ zoom: {zoom}%; }}
.hf {{ display: flex; justify-content: space-between; font-size: {page.get('header_size', 8)}pt; }}
table {{ border-collapse: collapse; table-layout: fixed; }}
td {{ padding: 0 2px; overflow: hidden; white-space: nowrap; }}
</style></head>
<body><div class="sheet">
<div class="hf"><span>{hf('header_left')}</span><span>{hf('header_center')}</span><span>{hf('header_right')}</span></div>
<table><colgroup>{cols}</colgroup>{''.join(rows)}</table>
<div class="hf"><span>{hf('footer_left')}</span><span>{hf('footer_center')}</span><span>{hf('footer_right')}</span></div>
</div></body></html>"""


class PdfError(RuntimeError):
    pass


def make_pdf(html_path: Path, pdf_path: Path, edge: str, *, timeout_sec: float = 60.0) -> Path:
    """Edge のヘッドレス印刷で PDF にする(D7)。Edge の場所は使用パス「Edge(PDF 作成)」。"""
    exe = Path(edge)
    if not exe.is_file():
        raise PdfError(f"Edge が見つかりません: {edge}(設定 → 使用パス で場所を直してください)")
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    cmd = [str(exe), "--headless", "--disable-gpu", "--no-pdf-header-footer",
           f"--print-to-pdf={pdf_path}", html_path.resolve().as_uri()]
    try:
        proc = subprocess.run(cmd, capture_output=True, timeout=timeout_sec,
                              creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except subprocess.TimeoutExpired:
        raise PdfError(f"PDF の作成が {timeout_sec:.0f} 秒で終わりませんでした") from None
    if proc.returncode != 0 or not pdf_path.is_file():
        raise PdfError(f"PDF を作れませんでした(終了コード {proc.returncode})")
    return pdf_path


def save_html_and_pdf(model: ReportModel, folder: Path, stem: str, edge: Optional[str]) -> dict[str, str]:
    """HTML を書き、Edge があれば PDF も作る。作ったファイルのパスを返す(Edge が無ければ html だけ)。"""
    folder.mkdir(parents=True, exist_ok=True)
    h = folder / f"{stem}.html"
    h.write_text(render(model, title=stem), encoding="utf-8")
    out = {"html": str(h)}
    if edge and os.name == "nt":
        out["pdf"] = str(make_pdf(h, folder / f"{stem}.pdf", edge))
    return out
