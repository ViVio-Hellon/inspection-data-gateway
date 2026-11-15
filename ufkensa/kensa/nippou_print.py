"""日報の紙(保守 → 日報。2026-11-14 のご依頼「見せ方は all-tools の印刷画面」)。

all-tools の日報モード(tools/nippou)の紙 `nippou/reporting/print_format.py`・`layout.py`・`reporting/paper.py`
を、標準ライブラリだけで写したもの。紙(梱包実績日報表)と同じ姿にする:

- 頭: 左に 作業日・直 / 梱包ライン / 作業者名、中央に題「梱　包　実　績　日　報　表」、右に ページ・昼稼働・その他
- 明細: 見出しは 2 段(梱包作業時間 → 開始 時・分 / 終了 時・分 など)。**12 行いつも出す**(打っていない行も空で)
- 下: 負荷計算後 係数Lot数 / 直実績合計 ﾛｯﾄ数・枚数・重量(Kg と T)
- 上の帯「印刷する」(紙には出ない)。A4 横、紙の端から 7mm 内側(@page の余白に任せない)

明細の出どころは日報モードへ渡す CSV と同じ(`nippou_export`)── いまの日報シートか、書き出した
`集計明細_*.csv`。紙の行は見出しの名前(日報モードの DETAIL_FIELDNAMES)で引く。
all-tools の紙にある「その直の集計(紙には無い、このツールの追加)」は、日報モードの集計の仕組みを
持ってこないので出さない(紙の姿そのものは同じ)。
"""
from __future__ import annotations

import csv
import html
import io
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Sequence

ROW_COUNT = 12
SHEET_TITLE = "梱　包　実　績　日　報　表"
STOP_TIME_NOTE = "※作業停止時間は5分以上で時間を記入"
REASON_LABEL = "その他"
PRINT_LABEL = "印刷する"
INSET_MM = 7


@dataclass(frozen=True)
class Column:
    key: str          # 日報モードの CSV の見出し(nippou_export.NIPPOU_FIELDS)
    label: str        # 下段の見出し(紙の言葉)
    group: str        # 上段の見出し(空なら 2 段ぶん高い)
    added: bool = False


#: all-tools layout.COLUMNS の紙に載る列(作業時間・単重は印刷範囲の外なので無い)
COLUMNS: tuple[Column, ...] = (
    Column("ロット№", "ロット№", ""), Column("材・調質", "材・調質", ""),
    Column("製品寸法 厚×幅×丈", "厚×幅×丈", "製品寸法"), Column("検入枚数", "検入枚数", ""),
    Column("梱包作業時間 開始 時", "開始 時", "梱包作業時間"), Column("梱包作業時間 開始 分", "開始 分", "梱包作業時間"),
    Column("梱包作業時間 終了 時", "終了 時", "梱包作業時間"), Column("梱包作業時間 終了 分", "終了 分", "梱包作業時間"),
    Column("作業人数", "作業人数", ""), Column("合紙", "合紙", ""),
    Column("梱包数量 個装単位 枚数", "個装単位 枚数", "梱包数量"), Column("梱包数量 梱包単位 包数", "梱包単位 包数", "梱包数量"),
    Column("ＶＣ種別 両面・片面", "両面・片面", "ＶＣ種別"), Column("etc 反転・EX etc", "反転・EX etc", "etc"),
    Column("作業停止① 記号", "記号", "作業停止①"), Column("作業停止① 時間(分)", "時間(分)", "作業停止①"),
    Column("作業停止② 記号", "記号", "作業停止②"), Column("作業停止② 時間(分)", "時間(分)", "作業停止②"),
    Column("作業停止③ 記号", "記号", "作業停止③"), Column("作業停止③ 時間(分)", "時間(分)", "作業停止③"),
    Column("作業停止④ 記号", "記号", "作業停止④", True), Column("作業停止④ 時間(分)", "時間(分)", "作業停止④", True),
    Column("作業停止⑤ 記号", "記号", "作業停止⑤", True), Column("作業停止⑤ 時間(分)", "時間(分)", "作業停止⑤", True),
    Column("実績合計 枚数", "枚数", "実績合計"), Column("実績合計 重量", "重量", "実績合計"),
)
PRINT_COLUMNS = tuple(c for c in COLUMNS if not c.added)

_STYLE = """
body { font-family: "Meiryo UI", "Yu Gothic", sans-serif; font-size: 12px; margin: 16px; color: #000; background: #fff; }
h1 { font-size: 16px; margin: 0 0 8px; }
table { border-collapse: collapse; width: 100%; margin-top: 8px; }
th, td { border: 1px solid #444; padding: 2px 4px; text-align: center; white-space: nowrap; }
th { background: #eee; }
.head { display: flex; align-items: flex-start; gap: 12px; margin-bottom: 4px; }
.head__keys { border-collapse: collapse; width: auto; margin: 0; flex: none; }
.head__keys th, .head__keys td { border: 1px solid #444; padding: 1px 6px; text-align: left;
  white-space: nowrap; font-size: 11px; background: #fff; font-weight: normal; }
.head__keys th { background: #eee; }
.head__title { flex: 1 1 auto; display: flex; align-items: center; justify-content: center; align-self: stretch; min-width: 0; }
.head__title h1 { font-size: 16px; margin: 0; letter-spacing: 2px; white-space: nowrap; }
.head__right { display: flex; gap: 8px; align-items: stretch; flex: none; }
.head__box { border: 1px solid #444; padding: 1px 6px; font-size: 11px; min-width: 5em; }
.head__box b { display: block; font-weight: normal; background: #eee; margin: -1px -6px 2px; padding: 0 6px; }
.head__box--reason { min-width: 16em; }
.note { font-size: 11px; margin: 4px 0 0; }
.generated-at { margin-top: 12px; font-size: 10px; color: #666; }
.totals { display: flex; justify-content: flex-end; margin-top: 4px; }
.totals table { width: auto; margin: 0; }
.totals th, .totals td { padding: 1px 10px; }
.totals .kg { display: block; }
.totals .ton { display: block; font-size: 11px; }
.sheet + .sheet { page-break-before: always; break-before: page; }
@media screen { .sheet + .sheet { border-top: 2px dashed #999; margin-top: 24px; padding-top: 16px; } }
@media print {
  .generated-at { display: none; }
  body { font-size: 11px; }
  th, td { padding: 1px 2px; }
  .sheet { break-inside: avoid; }
  .sheet > table th { white-space: normal; }
}
.toolbar { display: flex; gap: 8px; align-items: center; margin-bottom: 10px; padding: 6px 8px; background: #f4f4f4;
  border: 1px solid #ccc; position: sticky; top: 0; z-index: 1; }
.toolbar button { font: inherit; font-size: 14px; padding: 4px 16px; cursor: pointer; }
.toolbar .hint { font-size: 11px; color: #555; }
@media print { .toolbar { display: none !important; } }
@page { size: A4 landscape; margin: 0; }
@media print {
  html, body { margin: 0 !important; padding: 0 !important; }
  .sheet { box-sizing: border-box; padding: {INSET}mm; }
}
""".replace("{INSET}", str(INSET_MM))


@dataclass
class Page:
    report_date: str
    line: str
    shift: str
    page: str
    worker: str
    rows: list[dict[str, str]]          # 行番号 "行" を持つ
    day_shift: str = ""
    reason: str = ""


def _e(v: object) -> str:
    return html.escape(str(v if v is not None else ""))


def _num(s: str) -> Optional[float]:
    try:
        return float(str(s).replace(",", "").strip())
    except ValueError:
        return None


def _fmt(x: float, nd: int) -> str:
    return f"{x:.{nd}f}"


def column_groups(columns: Sequence[Column]) -> list[tuple[str, int]]:
    out: list[tuple[str, int]] = []
    for c in columns:
        if out and out[-1][0] == c.group and c.group:
            out[-1] = (c.group, out[-1][1] + 1)
        else:
            out.append((c.group, 1))
    return out


def paginate(rows: list[dict[str, str]], *, report_date: str, line: str, shift: str, worker: str) -> list[Page]:
    """行を紙のページ(12 行ずつ)に分ける。CSV が ページ・行 を持っていればそれを使う。"""
    by: dict[str, list[dict[str, str]]] = {}
    for i, r in enumerate(rows):
        pg = (r.get("ページ") or "").strip() or str(i // ROW_COUNT + 1)
        rr = dict(r)
        if not (rr.get("行") or "").strip():
            rr["行"] = str(i % ROW_COUNT + 1)
        by.setdefault(pg, []).append(rr)
    pages = []
    for pg in sorted(by, key=lambda p: (_num(p) or 0, p)):
        first = by[pg][0]
        pages.append(Page(first.get("日付") or report_date, first.get("ライン") or line, first.get("直") or shift, pg,
                          first.get("作業者") or worker, by[pg],
                          day_shift=first.get("昼稼働", ""), reason=first.get("作業コメント", "")))
    return pages


def _head(p: Page) -> str:
    rows = "".join(f"<tr><th>{_e(k)}</th><td>{_e(v)}</td></tr>"
                   for k, v in (("作業日・直", f"{p.report_date}{p.shift}"), ("梱包ライン", p.line), ("作業者名", p.worker)))
    return (f'<div class="head"><table class="head__keys">{rows}</table>'
            f'<div class="head__title"><h1>{_e(SHEET_TITLE)}</h1></div><div class="head__right">'
            f'<div class="head__box"><b>ページ</b>{_e(p.page)}</div>'
            f'<div class="head__box"><b>昼稼働</b>{_e(p.day_shift or "無")}</div>'
            f'<div class="head__box head__box--reason"><b>{_e(REASON_LABEL)}</b>{_e(p.reason)}</div></div></div>')


def shift_totals(pages: Sequence[Page]) -> dict[str, str]:
    """直実績合計(直ぶん。ﾛｯﾄ数は ロット№ の数)。数でない欄は数えない。"""
    rows = [r for p in pages for r in p.rows if (r.get("ロット№") or "").strip()]
    cnt = sum(_num(r.get("実績合計 枚数", "")) or 0 for r in rows)
    wei = sum(_num(r.get("実績合計 重量", "")) or 0 for r in rows)
    coef = [_num(r.get("係数処理ﾛｯﾄ数", "")) for r in rows]
    return {"lots": str(len({(r.get("ロット№") or "").strip() for r in rows})) if rows else "",
            "count": _fmt(cnt, 0) if rows and cnt else "", "kg": _fmt(wei, 1) if rows and wei else "",
            "ton": _fmt(wei / 1000, 2) if rows and wei else "",
            "coef": _fmt(sum(c for c in coef if c is not None), 2) if any(c is not None for c in coef) else ""}


def _totals(t: dict[str, str]) -> str:
    weight = f'<span class="kg">{_e(t["kg"])} Kg</span>' + (f'<span class="ton">{_e(t["ton"])} T</span>' if t["ton"] else "")
    return ('<div class="totals"><table><thead><tr><th>負荷計算後</th><th colspan="3">直実績合計</th></tr>'
            '<tr><th>係数Lot数</th><th>ﾛｯﾄ数</th><th>枚数</th><th>重量</th></tr></thead>'
            f'<tbody><tr><td>{_e(t["coef"])}</td><td>{_e(t["lots"])}</td><td>{_e(t["count"])}</td><td>{weight}</td></tr>'
            '</tbody></table></div>')


def _sheet(p: Page, totals: dict[str, str], generated_at: str) -> str:
    added = any((r.get(c.key) or "").strip() for r in p.rows for c in COLUMNS if c.added)
    cols = COLUMNS if added else PRINT_COLUMNS
    top, i = [], 0
    for group, span in column_groups(cols):
        top.append(f'<th colspan="{span}">{_e(group)}</th>' if group else f'<th rowspan="2">{_e(cols[i].label)}</th>')
        i += span
    bottom = "".join(f"<th>{_e(c.label)}</th>" for c in cols if c.group)
    by_row: dict[int, dict[str, str]] = {}
    extra: list[tuple[int, dict[str, str]]] = []
    for r in p.rows:
        n = int(_num(r.get("行", "")) or 0)
        if 1 <= n <= ROW_COUNT and n not in by_row:
            by_row[n] = r
        else:
            extra.append((n, r))
    body = []
    for n in range(1, ROW_COUNT + 1):
        d = by_row.get(n)
        body.append(f"<tr><td>{n}</td>" + "".join(f"<td>{_e(d.get(c.key, '') if d else '')}</td>" for c in cols) + "</tr>")
    for n, d in extra:                        # 12 行に入らない行も捨てずに下へ
        body.append(f"<tr><td>{n or ''}</td>" + "".join(f"<td>{_e(d.get(c.key, ''))}</td>" for c in cols) + "</tr>")
    title = f"日報 {p.report_date} {p.line} {p.shift} 第{p.page}ページ"
    return (f'<section class="sheet">{_head(p)}<div class="note">{_e(STOP_TIME_NOTE)}</div>'
            f'<table><thead><tr><th rowspan="2">行</th>{"".join(top)}</tr><tr>{bottom}</tr></thead>'
            f'<tbody>{"".join(body)}</tbody></table>{_totals(totals)}'
            f'<div class="generated-at">生成日時: {_e(generated_at)} / {_e(title)}</div></section>')


def build_html(pages: Sequence[Page], *, generated_at: str = "", note: str = "") -> str:
    """直ぶんを 1 つの窓に(ページのあいだで改ページ)。all-tools の build_shift_html と同じ。"""
    if not pages:
        return _doc("日報", '<p class="note">該当するデータがありません。</p>', 0, note)
    t = shift_totals(pages)
    body = "".join(_sheet(p, t, generated_at) for p in pages)
    h = pages[0]
    return _doc(f"日報 {h.report_date} {h.line} {h.shift} 全{len(pages)}ページ", body, len(pages), note)


def _doc(title: str, body: str, pages: int, note: str) -> str:
    close = "刷り終わったら、この窓は閉じてかまいません"
    hint = (f"全{pages}ページ ─ A4 横で{pages}枚に分かれて刷れます。{close}" if pages > 1
            else f"A4 横1枚で刷れます。{close}")
    bar = ""
    if pages:
        bar = (f'<div class="toolbar"><button type="button" onclick="window.print()">{PRINT_LABEL}</button>'
               '<button type="button" onclick="var d=window.__TAURI__&&window.__TAURI__.core;'
               'if(d){d.invoke(\'close_window\')}else{window.close()}">閉じる</button>'
               f'<span class="hint">{_e(hint)}</span></div>\n')
    if note:
        bar += f'<div class="toolbar"><span class="hint">{_e(note)}</span></div>\n'
    return (f'<!doctype html>\n<html lang="ja"><head><meta charset="utf-8"><title>{_e(title)}</title>'
            f'<style>{_STYLE}</style></head><body>\n{bar}{body}</body></html>\n')


def read_csv(path: Path) -> list[dict[str, str]]:
    """日報モードの CSV(BOM 付き UTF-8。だめなら cp932)を 見出し → 値 で。"""
    data = path.read_bytes()
    for enc in ("utf-8-sig", "cp932"):
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = data.decode("utf-8", "replace")
    return [{k: (v or "") for k, v in r.items() if k} for r in csv.DictReader(io.StringIO(text))]
