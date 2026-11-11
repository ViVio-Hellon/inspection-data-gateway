"""帳票の中身(検査成績票・不良要因票)をデータで作る。VBA `【UFKen_08】Print` の移植

VBA は原紙のシートを写して(Pesute)、セルに値を書き、罫線を引き直し、ページ設定をして印刷していた。
ここでは同じ手順を**小さな表計算モデル**(`XSheet`)の上で行い、結果を `ReportModel`
(セルごとの値・書式済みの文字列・文字の大きさ・罫線の種類・印刷範囲・ヘッダー/フッター)で返す。
HTML にするのは別の工程(この結果をそのまま描けばよい)。

- 原紙の様式(値・表示形式・文字の大きさ・配置・罫線・セル結合・列幅・行高)は xlsm の `印刷用原紙` と
  `不良要因印刷原紙` から写したもの(このファイルの末尾 `_TEMPLATES`)
- セルへの代入は Excel と同じ型の変わり方をする(`store.excel_value`。数字の文字列は数になり、表示形式が効く)
- 罫線は Excel と同じく、隣り合うセルの辺を同じ線として扱う(API で引いた線は両側に入る。コピーはセル自身の辺だけ)

対応する VBA:
    印刷転送 (Print:184-734)        → print_transfer()
    印刷用罫線 (Print:736-1136)      → _print_borders()
    印刷範囲 / 印刷範囲2 (Print:1137-1200) → _page_area_abnormal() / _page_area_normal()
    Pesute (Print:1202-1276)         → XSheet.from_template() / paste_template_values()
    不良要因印刷範囲 (Print:1278-1335) → defect_print_range()
    CalculateColumnSums / ConvertNumbersToString (Print:1337-1396) → calculate_column_sums() / convert_numbers_to_string()
    FitTextToCell (Print:1412-1460)  → fit_text_font_size()(文字幅の見積もりによる近似)
    InitializePrintFlags〜CleanOldHistory (Print:8-182) → load_print_history() ほか
    UF1.PrintOnly (UF1:3574-3597)    → print_only()
    OutTolerance.allow_Check (OutTolerance:2598-2633) → allow_check()(B34 の分岐のため、ここにも置く)
    補助.CircledNumber (補助:1205-1216) → circled_number()
"""
from __future__ import annotations

import copy
import itertools
import json
import re
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta, time as dtime
from typing import Any, Callable, Iterable, Mapping, Optional, Sequence

from . import flow, vbx
from .model import EXTREME, Face, NORMAL, THICK_ROWS
from . import sound as snd
from .flow import Message, Question, Stop
from .records import (FACES, KensaSnapshot, cmp_str_num, end_time_register, uf_array, uf_tag_array, vb_cdate,
                      vb_cstr_date, vb_format_date, vb_isdate, variant_eq)
from .store import (GridSheet, LocalStore, PRINT_HISTORY, Value, cell_text, col_index, col_letter, excel_value,
                    line_sheet)

PRINT_SHEET = "印刷用"
DEFECT_SHEET = "不良要因印刷"
ABNORMAL_SHEET = "異常ロット用"
WORKER_TIME = "WorkerTime"

#: Excel の色(RGB 関数の値)
RGB_ORANGE = 255 + 165 * 256            # RGB(255,165,0)
RGB_GRAY = 192 + 192 * 256 + 192 * 65536
RGB_FLAG = 51 + 204 * 256 + 204 * 65536  # RGB(51,204,204) 指示公差の色

# ------------------------------------------------------------------
# 罫線
# ------------------------------------------------------------------
NONE = ("none", "")
_XLSX_BORDER = {
    "thin": ("continuous", "thin"), "medium": ("continuous", "medium"), "thick": ("continuous", "thick"),
    "double": ("double", "thick"), "dotted": ("dot", "thin"), "hair": ("continuous", "hairline"),
    "dashed": ("dash", "thin"), "mediumDashed": ("dash", "medium"), "dashDot": ("dashdot", "thin"),
    "mediumDashDot": ("dashdot", "medium"), "dashDotDot": ("dashdotdot", "thin"),
    "mediumDashDotDot": ("dashdotdot", "medium"), "slantDashDot": ("slantdashdot", "medium"),
}
SIDES = ("top", "bottom", "left", "right")
_OPP = {"top": "bottom", "bottom": "top", "left": "right", "right": "left"}
_STEP = {"top": (-1, 0), "bottom": (1, 0), "left": (0, -1), "right": (0, 1)}
_stamp = itertools.count(1)


def _apply_line(old: tuple[str, str], linestyle: Optional[str], weight: Optional[str]) -> tuple[str, str]:
    """`Border.LineStyle = …` / `Border.Weight = …` を当てた後の (線種, 太さ)。"""
    style, w = old
    if linestyle is not None:
        if linestyle == "none":
            return NONE
        if linestyle == "double":
            return ("double", "thick")
        if linestyle == "dot":
            return ("dot", "thin")
        if linestyle == "continuous":
            return ("continuous", w if style == "continuous" and w else "thin")
        return (linestyle, w or "thin")
    if weight is not None:
        if style in ("none", "double", "dot"):
            return ("continuous", weight)
        return (style, weight)
    return old


@dataclass
class XCell:
    value: Value = None
    fmt: str = "General"
    size: float = 11.0
    bold: bool = False
    font: str = "P"
    ha: str = ""
    va: str = ""
    wrap: bool = False
    shrink: bool = False
    b: dict = field(default_factory=lambda: {s: (NONE, 0) for s in SIDES})   # 辺 → ((線種, 太さ), 書いた順)


class XSheet:
    """VBA が帳票に使う Worksheet の操作だけを持つ小さな表(行・列は 1 始まり)。"""

    def __init__(self, name: str) -> None:
        self.name = name
        self.cells: dict[tuple[int, int], XCell] = {}
        self.merges: list[tuple[int, int, int, int]] = []
        self.widths: dict[int, float] = {}
        self.heights: dict[int, float] = {}
        self.page: dict[str, Any] = {}
        self.print_area: Optional[tuple[int, int, int, int]] = None
        self.figure_at: Optional[str] = None

    # -- 原紙 --------------------------------------------------------
    @classmethod
    def from_template(cls, template: str, name: str) -> "XSheet":
        """VBA `Pesute(原紙, name)`: 原紙の全セル(値・書式・罫線)を写す。塗りなし・文字は黒に戻す。"""
        t = _TEMPLATES[template]
        sh = cls(name)
        for line in t["cells"].strip("\n").split("\n"):
            ref, v, fi, sz, bold, ha, va, wrap, shrink, borders, font = line.split("|")
            r, c = _ref(ref)
            cell = sh.cell(r, c)
            if v.startswith("#"):
                num = float(v[1:])
                cell.value = num
            else:
                cell.value = v if v != "" else None
            cell.fmt = _FORMATS[int(fi)]
            cell.size = float(sz) if sz else 11.0
            cell.bold = bold == "1"
            cell.ha, cell.va = ha, va
            cell.wrap, cell.shrink = wrap == "1", shrink == "1"
            cell.font = font
            for side, st in zip(SIDES, borders.split(",")):
                if st:
                    cell.b[side] = (_XLSX_BORDER.get(st, ("continuous", "thin")), 0)
        sh.merges = [_range_ref(m) for m in t["merges"]]
        sh.widths = {col_index(k): v for k, v in t["widths"].items()}
        sh.heights = dict(t["heights"])
        return sh

    # -- セル ------------------------------------------------------
    def cell(self, r: int, c: int) -> XCell:
        key = (r, c)
        if key not in self.cells:
            self.cells[key] = XCell()
        return self.cells[key]

    def get(self, r: int, c: int) -> Value:
        x = self.cells.get((r, c))
        return x.value if x else None

    def put(self, r: int, c: int, v: Any, kind: str = "V") -> None:
        """`Cells(r, c) = v`。"""
        self.cell(r, c).value = excel_value(v, kind)

    def put_range(self, r: int, c: int, rows: Sequence[Sequence[Any]], kind: str = "V") -> None:
        """`Range(...) = 2 次元配列`(左上 r, c)。"""
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                self.put(r + i, c + j, v, kind)

    def text(self, r: int, c: int) -> str:
        """セルの表示(表示形式つき)。"""
        x = self.cells.get((r, c))
        if not x:
            return ""
        return format_cell(x.value, x.fmt)[0]

    def last_row(self, col: int) -> int:
        rows = [r for (r, c), x in self.cells.items() if c == col and x.value is not None]
        return max(rows) if rows else 1

    # -- 範囲の操作 ----------------------------------------------------
    @staticmethod
    def _norm(r1: int, c1: int, r2: int, c2: int) -> tuple[int, int, int, int]:
        if c1 < 1 or c2 < 1 or r1 < 1 or r2 < 1:
            raise Stop("実行時エラー '1004': アプリケーション定義またはオブジェクト定義のエラーです。",
                       source="Print")
        return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)

    def _iter(self, r1: int, c1: int, r2: int, c2: int) -> Iterable[tuple[int, int]]:
        a, b, c, d = self._norm(r1, c1, r2, c2)
        for r in range(a, c + 1):
            for cc in range(b, d + 1):
                yield r, cc

    def clear_contents(self, r1: int, c1: int, r2: int, c2: int) -> None:
        for r, c in self._iter(r1, c1, r2, c2):
            if (r, c) in self.cells:
                self.cells[(r, c)].value = None

    def copy(self, r1: int, c1: int, r2: int, c2: int, dr: int, dc: int, *, formats_only: bool = False) -> None:
        """`Range.Copy Destination` / `PasteSpecial xlPasteFormats`(左上 dr, dc に同じ大きさで貼る)。"""
        a, b, c, d = self._norm(r1, c1, r2, c2)
        snap = {(r, cc): copy.deepcopy(self.cells.get((r, cc), XCell())) for r in range(a, c + 1)
                for cc in range(b, d + 1)}
        for (r, cc), src in snap.items():
            tr, tc = dr + (r - a), dc + (cc - b)
            dst = self.cell(tr, tc)
            keep = dst.value
            n = next(_stamp)
            new = copy.deepcopy(src)
            new.b = {s: (src.b[s][0], n if src.b[s][0] != NONE else 0) for s in SIDES}
            if formats_only:
                new.value = keep
            self.cells[(tr, tc)] = new

    def _set_side(self, r: int, c: int, side: str, linestyle: Optional[str], weight: Optional[str]) -> None:
        cur = self.effective(r, c, side)
        new = _apply_line(cur, linestyle, weight)
        n = next(_stamp)
        self.cell(r, c).b[side] = (new, n)
        dr, dc = _STEP[side]
        if r + dr >= 1 and c + dc >= 1:
            self.cell(r + dr, c + dc).b[_OPP[side]] = (new, n)

    def border(self, r1: int, c1: int, r2: int, c2: int, index: str, *, linestyle: Optional[str] = None,
               weight: Optional[str] = None) -> None:
        """`Range.Borders(index).LineStyle / .Weight`。index: top/bottom/left/right/inside_h/inside_v。"""
        a, b, c, d = self._norm(r1, c1, r2, c2)
        if index == "top":
            targets = [(a, cc, "top") for cc in range(b, d + 1)]
        elif index == "bottom":
            targets = [(c, cc, "bottom") for cc in range(b, d + 1)]
        elif index == "left":
            targets = [(r, b, "left") for r in range(a, c + 1)]
        elif index == "right":
            targets = [(r, d, "right") for r in range(a, c + 1)]
        elif index == "inside_h":
            targets = [(r, cc, "bottom") for r in range(a, c) for cc in range(b, d + 1)]
        elif index == "inside_v":
            targets = [(r, cc, "right") for r in range(a, c + 1) for cc in range(b, d)]
        else:
            raise ValueError(index)
        for r, cc, side in targets:
            self._set_side(r, cc, side, linestyle, weight)

    def borders_all(self, r1: int, c1: int, r2: int, c2: int, *, linestyle: Optional[str] = None,
                    weight: Optional[str] = None) -> None:
        """`Range.Borders.LineStyle = …`(添字なしの Borders。範囲の各セル自身の 4 辺)。

        Excel の添字なし Borders はセル自身の辺(xlLeft/xlRight/xlTop/xlBottom)に効き、範囲の外の
        隣のセルが持つ向かいの辺は消さない。実機の印刷物(厚板・ラテラルボー無)で、J:N を消しても
        I 列(差)の右の太線が残ることを確かめた(Print:1097-1098 ほか)。
        """
        for r, c in self._iter(r1, c1, r2, c2):
            x = self.cell(r, c)
            for side in SIDES:
                new = _apply_line(x.b[side][0], linestyle, weight)
                x.b[side] = (new, next(_stamp))

    def effective(self, r: int, c: int, side: str) -> tuple[str, str]:
        """見えている線(自分の辺と隣のセルの向かいの辺のうち、後から引いた方)。"""
        own = self.cells.get((r, c))
        o = own.b[side] if own else (NONE, 0)
        dr, dc = _STEP[side]
        nb = self.cells.get((r + dr, c + dc))
        n = nb.b[_OPP[side]] if nb else (NONE, 0)
        cands = [x for x in (o, n) if x[0] != NONE]
        if not cands:
            return NONE
        return max(cands, key=lambda x: x[1])[0]

    def number_format(self, r1: int, c1: int, r2: int, c2: int, fmt: str) -> None:
        for r, c in self._iter(r1, c1, r2, c2):
            self.cell(r, c).fmt = fmt

    def font_size(self, r1: int, c1: int, r2: int, c2: int, size: float) -> None:
        for r, c in self._iter(r1, c1, r2, c2):
            self.cell(r, c).size = size

    def align(self, r1: int, c1: int, r2: int, c2: int, *, ha: Optional[str] = None, va: Optional[str] = None) -> None:
        for r, c in self._iter(r1, c1, r2, c2):
            x = self.cell(r, c)
            if ha is not None:
                x.ha = ha
            if va is not None:
                x.va = va

    def replace(self, r1: int, c1: int, r2: int, c2: int, find: str, repl: str) -> None:
        """`Range.Replace find, repl, xlPart`(置き換えた結果は入力し直したのと同じ扱い)。"""
        for r, c in self._iter(r1, c1, r2, c2):
            x = self.cells.get((r, c))
            if x and isinstance(x.value, str) and find in x.value:
                x.value = excel_value(x.value.replace(find, repl), "V")

    # -- 出力 --------------------------------------------------------
    def to_model(self, kind: str) -> "ReportModel":
        if self.print_area:
            r1, c1, r2, c2 = self.print_area
        else:
            keys = [k for k, x in self.cells.items() if x.value is not None] or [(1, 1)]
            r1, c1 = 1, 1
            r2, c2 = max(k[0] for k in keys), max(k[1] for k in keys)
        out = []
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                x = self.cells.get((r, c))
                borders = {s: list(self.effective(r, c, s)) for s in SIDES}
                if x is None and all(v[0] == "none" for v in borders.values()):
                    continue
                x = x or XCell()
                text, red = format_cell(x.value, x.fmt)
                out.append({"ref": f"{col_letter(c)}{r}", "row": r, "col": c, "value": x.value, "text": text,
                            "red": red, "number_format": x.fmt, "font_size": x.size, "bold": x.bold,
                            "font": {"M": "Meiryo UI", "P": "ＭＳ Ｐゴシック"}.get(x.font, x.font),
                            "h_align": x.ha, "v_align": x.va, "wrap": x.wrap, "shrink": x.shrink,
                            "borders": borders})
        return ReportModel(sheet=self.name, kind=kind, print_area=self.print_area,
                           page=dict(self.page), cells=out,
                           merges=[f"{col_letter(a)}{b}:{col_letter(d)}{c}" for b, a, c, d in
                                   [(m[0], m[1], m[2], m[3]) for m in self.merges]],
                           col_widths={col_letter(k): v for k, v in sorted(self.widths.items())},
                           row_heights=dict(sorted(self.heights.items())), figure_at=self.figure_at)


def _ref(ref: str) -> tuple[int, int]:
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    assert m, ref
    return int(m.group(2)), col_index(m.group(1))


def _range_ref(text: str) -> tuple[int, int, int, int]:
    a, b = text.split(":")
    r1, c1 = _ref(a)
    r2, c2 = _ref(b)
    return r1, c1, r2, c2


# ------------------------------------------------------------------
# 表示形式
# ------------------------------------------------------------------
def _general(v: float) -> str:
    """表示形式「標準」での数の見え方(セル幅いっぱい・有効 10 桁までの近似)。"""
    if v == int(v) and abs(v) < 1e11:
        return str(int(v))
    s = f"{v:.10g}"
    if "e" in s:
        mant, exp = s.split("e")
        e = int(exp)
        return f"{mant}E{'+' if e >= 0 else '-'}{abs(e):02d}"
    return s


def format_cell(value: Value, fmt: str) -> tuple[str, bool]:
    """セルの値を表示形式で文字にする。(文字, 赤字か)。

    使う形: 標準 / `0.0_ ` / `0.000_ ` / `0.00_ ` / `0_ ` / `0.0` / `@` /
    `0.000_);[Red]\\(0.000\\)` / `0.00_);[赤](0.00)` など(正の部;負の部)。`_)` の空き幅は文字にしない。
    """
    if value is None:
        return "", False
    if isinstance(value, str):
        return value, False
    v = float(value)
    if fmt in ("General", "@", ""):
        return _general(v), False
    sections = fmt.split(";")
    sec = sections[0]
    neg = v < 0 and len(sections) > 1
    if neg:
        sec = sections[1]
    red = "[Red]" in sec or "[赤]" in sec
    body = re.sub(r"\[[^\]]*\]", "", sec)
    body = re.sub(r"_.", "", body)
    body = body.replace("\\", "")
    m = re.search(r"0(?:\.(0+))?", body)
    if not m:
        return _general(v), red
    places = len(m.group(1) or "")
    num = vbx.fmt(abs(v) if neg else v, "0" + ("." + "0" * places if places else ""))
    return body[:m.start()] + num + body[m.end():], red


# ------------------------------------------------------------------
# 帳票のデータ
# ------------------------------------------------------------------
@dataclass
class ReportModel:
    sheet: str                       # VBA のシート名(印刷用_<L> / 不良要因印刷_<L>)
    kind: str                        # "検査成績票" / "不良要因票"
    print_area: Optional[tuple[int, int, int, int]]
    page: dict                       # 用紙・向き・余白(inch)・倍率・ヘッダー・フッター
    cells: list[dict]                # 印刷範囲のセル(値・文字・書式・罫線)
    merges: list[str]
    col_widths: dict[str, float]
    row_heights: dict[int, float]
    figure_at: Optional[str] = None  # `図` シートの図形を貼った位置(公差外ありのとき "L11")

    def cell(self, ref: str) -> Optional[dict]:
        for c in self.cells:
            if c["ref"] == ref:
                return c
        return None

    def text(self, ref: str) -> str:
        c = self.cell(ref)
        return c["text"] if c else ""

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False)

    @classmethod
    def from_json(cls, text: str) -> "ReportModel":
        d = json.loads(text)
        d["print_area"] = tuple(d["print_area"]) if d.get("print_area") else None
        d["row_heights"] = {int(k): v for k, v in d.get("row_heights", {}).items()}
        return cls(**d)

    def sections(self) -> dict[str, Any]:
        """成績票の意味ごとのまとまり(見出し・厚の表・幅丈の表・公差外表)。"""
        t = self.text
        out: dict[str, Any] = {
            "lot": t("C2"), "reprint": t("D2"), "size_fix": t("B1"), "history": t("B3"), "line": t("B4"),
            "ken_label": t("F2"), "ken_value": t("G2"), "yama": t("H2"), "formula": t("F1"),
            "size": {"厚": t("D4"), "巾": t("F4"), "丈": t("H4")},
            "tolerance": [[t(f"{c}{r}") for c in "CDEFGH"] for r in (5, 6)],
            "distortion": {"参照元": t("I3"), "ゲージ": t("I6"), "幅": t("J6"), "丈": t("K6"),
                           "中歪": t("Q9"), "耳歪": t("R9"), "対角": t("L6"), "フリー反り": t("L8")},
            "worker": t("B9"), "micro": t("E9"), "convex": t("G9"), "allow": t("B10"),
            "upper": [[t(f"{c}{r}") for c in "BCEFGHIJ"] for r in range(11, 18)],
            "lower": [[t(f"{col_letter(c)}{r}") for c in range(2, 19)] for r in range(18, 25)],
        }
        if t("A26") == "公差外ﾃﾞｰﾀ":
            rows = []
            r = 27
            while self.cell(f"A{r}") or self.cell(f"B{r}"):
                rows.append([t(f"{col_letter(c)}{r}") for c in range(1, 16)])
                r += 1
            out["abnormal"] = {"header": [t(f"{col_letter(c)}26") for c in range(1, 16)], "rows": rows}
        return out


@dataclass
class PrintResult:
    """印刷転送の結果。stopped なら `PrintB.Tag="stop"`(印刷しない)。"""
    stopped: bool
    report: Optional[ReportModel] = None
    defect: Optional[ReportModel] = None
    end_time: str = ""
    messages: list[Message] = field(default_factory=list)


# ------------------------------------------------------------------
# 部品
# ------------------------------------------------------------------
def circled_number(num: int) -> str:
    """補助.CircledNumber (補助:1205-1216)。1〜50 を ①〜㊿ に。範囲外は数字のまま(文字列)。"""
    if num < 1 or num > 50:
        return str(num)
    if num <= 20:
        base = 9311
    elif num <= 35:
        base = 12860
    else:
        base = 12941
    return chr(base + num)


def allow_check(ng_captions: Mapping[str, str], positions: Mapping[str, str]) -> str:
    """OutTolerance.allow_Check (OutTolerance:2598-2633)。許容/保留の位置名を返す。

    B34: VBA は 1 つ目の位置に「許容/保留」を付けないので、1 か所だけだと B10 が出ない。
    直す場合は 1 つ目にも付ける。
    """
    result = ""
    for face in FACES:
        cap = ng_captions.get(face, "") or ""
        name = "NG" + face
        if "許容" in cap:
            label = " 許容"
        elif "保留" in cap:
            label = " 保留"
        else:
            continue
        if result == "":
            result = name + label
        else:
            result = result + ", " + name + label
    for face in FACES:
        result = result.replace("NG" + face, positions.get(face, "") or "")
    return result


def _visual_mark(v: str) -> str:
    """外観に記号: 欠陥あり → △、ちょうど「欠陥ﾅｼ」→ 〇 (Print:205-237)。"""
    if v != "" and not vbx.like(v, "*欠陥ﾅｼ*"):
        v = "△" + v
    if v == "欠陥ﾅｼ":
        v = "〇" + v
    return v


def _visual_ab(a: str, b: str) -> str:
    """外観 A/B を 1 セルに (Print:240-287)。"""
    if a != "" and b != "":
        return "A:" + a + "\r\n" + "B:" + b
    if a != "" and b == "":
        return "A:" + a + "\r\n" + ""
    if a == "" and b != "":
        return "" + "\r\n" + "B:" + b
    return ""


def fit_text_font_size(text: str, width_chars: float, height_pt: float) -> float:
    """FitTextToCell (Print:1412-1460) の近似。

    VBA は一時テキストボックスで文字の外寸を測り、8pt から 0.5pt ずつ下げて収まる大きさを探す。
    結果が 4pt 以上なら 8pt に戻し、3.5pt なら 6pt、それ未満はその値(下限 1pt)。
    ここでは文字幅(全角 1・半角 0.5 字)と行の高さ(文字の 1.3 倍)で見積もる(実機の計測と違いうる)。
    """
    import unicodedata
    width_pt = (width_chars * 7 + 5) * 0.75 - 14.4     # テキストボックスの左右の余白(7.2pt×2)を引く

    def em(s: str) -> float:
        return sum(0.5 if unicodedata.east_asian_width(ch) in ("Na", "H") else 1.0 for ch in s)

    size = 8.0
    paras = text.replace("\r\n", "\n").split("\n") if text else [""]
    while True:
        lines = 0
        for p in paras:
            w = em(p) * size
            lines += max(1, int(-(-w // max(width_pt, 1))))
        if lines * size * 1.3 <= height_pt or size <= 0.5:
            break
        size -= 0.5
    if size >= 4.0:
        size = 8.0
    elif size == 3.5:
        size = 6.0
    if size < 1:
        size = 1.0
    return size


# ------------------------------------------------------------------
# 印刷履歴(印刷履歴管理)
# ------------------------------------------------------------------
#: U12: `IsDate("2026.06.22")` が真か(実機確認待ち)。偽なら「.」区切りの日付は消さない
DOT_DATE_IS_DATE = False


def _print_history_sheet(store: LocalStore):
    return store.sheet(PRINT_HISTORY)


def load_print_history(store: LocalStore) -> dict[str, Value]:
    """LoadPrintHistory (Print:31-53)。LOT → 印刷日時(同じ LOT は後の行が勝つ)。"""
    sh = _print_history_sheet(store)
    out: dict[str, Value] = {}
    last = sh.last_row(1)
    if last > 1:
        for i in range(2, last + 1):
            lot = cell_text(sh.get(i, 1))
            if lot != "":
                out[lot] = sh.get(i, 2)
    return out


def is_printed(store: LocalStore, lot: str) -> bool:
    """IsPrinted (Print:96-99)。毎回読み直す。"""
    return lot in load_print_history(store)


def save_print_history(store: LocalStore, lot: str, lname: str, now: datetime) -> None:
    """SavePrintHistory (Print:56-93)。あれば B 列を今に、無ければ最後の行の次に追記。"""
    sh = _print_history_sheet(store)
    last = sh.last_row(1)
    existing = 0
    if last > 1:
        for i in range(2, last + 1):
            if cell_text(sh.get(i, 1)) == lot:
                existing = i
                break
    stamp = vb_cstr_date(now)
    if existing == 0:
        new = last + 1
        sh.write(new, 1, lot)
        sh.write(new, 2, stamp, "S")
        sh.write(new, 3, lname)
    else:
        sh.write(existing, 2, stamp, "S")


def set_printed(store: LocalStore, lot: str, lname: str, now: datetime) -> None:
    """SetPrinted (Print:126-130)。"""
    save_print_history(store, lot, lname, now)


def clear_printed(store: LocalStore, lot: str) -> None:
    """ClearPrinted (Print:133-157)。下から探して最初の 1 行だけ消す。"""
    sh = _print_history_sheet(store)
    last = sh.last_row(1)
    if last > 1:
        for i in range(last, 1, -1):
            if cell_text(sh.get(i, 1)) == lot:
                sh.delete_row(i)
                break


def clean_old_history(store: LocalStore, today: datetime) -> list[str]:
    """CleanOldHistory (Print:160-182)。消した LOT を返す。

    B37: VBA は `DateDiff("d") > 3`(4 日以上前)で消す。直すとコメント(UF1:184)どおり 30 日以上前。
    """
    sh = _print_history_sheet(store)
    last = sh.last_row(1)
    removed = []
    if last > 1:
        for i in range(last, 1, -1):
            v = sh.get(i, 2)
            if vb_isdate(v, dot_is_date=DOT_DATE_IS_DATE):
                d = vb_cdate(v, dot_is_date=DOT_DATE_IS_DATE)
                diff = (today.date() - d.date()).days
                old = diff > 3
                if old:
                    removed.append(cell_text(sh.get(i, 1)))
                    sh.delete_row(i)
    return removed


def print_button_state(store: LocalStore, lot: str) -> dict[str, Any]:
    """UpdatePrintButtonColor (Print:102-115)。"""
    if is_printed(store, lot):
        return {"caption": "再印刷", "back_color": RGB_ORANGE, "reset_visible": True}
    return {"caption": "印刷", "back_color": RGB_GRAY, "reset_visible": False}


def print_only(store: LocalStore, lname: str, lot: str, answers: Mapping[str, bool], now: datetime):
    """UF1.PrintOnly (UF1:3574-3597)。印刷済みの LOT を、残っている帳票のまま再印刷する。

    「再印刷しますか？」が はい なら成績票の D2 に「再印刷」と書いて、成績票と不良要因票をそのまま返す
    (データの作り直し・保存・CSV・PDF はしない。W8: 別 LOT の内容が残っていればそれが出る)。
    """
    msgs = [Message("", sound=snd.KENSA_REPRINT, source="UF1:3578")]
    ans = answers.get("reprint")
    if ans is None:
        return Question("reprint", "このLOTは印刷済みです。再印刷しますか？", buttons="yesno", source="UF1:3579",
                        messages=msgs)
    if not ans:
        return None
    rep_json = store.load_report(line_sheet(PRINT_SHEET, lname))
    if rep_json is None or not store.exists(line_sheet(DEFECT_SHEET, lname)):
        raise Stop("実行時エラー '9': インデックスが有効範囲にありません。", source="UF1:3575")
    report = ReportModel.from_json(rep_json)
    sh = _sheet_from_model(report)
    sh.put(2, 4, "再印刷")
    report = sh.to_model("検査成績票")
    report.page = dict(sh.page)
    store.save_report(line_sheet(PRINT_SHEET, lname), report.to_json())
    defect = render_defect_sheet(store, lname)
    set_printed(store, lot, lname, now)
    msgs += [Message("", sound=snd.KENSA_REPRINT_DONE, source="UF1:3592"),
             Message("再印刷が完了しました", level=flow.INFO, source="UF1:3593")]
    return PrintResult(stopped=False, report=report, defect=defect, messages=msgs)


def _sheet_from_model(model: ReportModel) -> XSheet:
    """保存しておいた成績票をセル表に戻す(再印刷で D2 を書き足すため)。"""
    sh = XSheet(model.sheet)
    sh.print_area = model.print_area
    sh.page = dict(model.page)
    sh.figure_at = model.figure_at
    sh.widths = {col_index(k): v for k, v in model.col_widths.items()}
    sh.heights = dict(model.row_heights)
    sh.merges = [_range_ref(m) for m in model.merges]
    for c in model.cells:
        x = sh.cell(c["row"], c["col"])
        x.value = c["value"]
        x.fmt = c["number_format"]
        x.size = c["font_size"]
        x.bold = c["bold"]
        x.font = {"Meiryo UI": "M", "ＭＳ Ｐゴシック": "P"}.get(c["font"], c["font"])
        x.ha, x.va, x.wrap, x.shrink = c["h_align"], c["v_align"], c["wrap"], c["shrink"]
        x.b = {s: (tuple(c["borders"][s]), 1 if c["borders"][s][0] != "none" else 0) for s in SIDES}
    return sh


# ------------------------------------------------------------------
# 不良要因票
# ------------------------------------------------------------------
def ensure_defect_sheet(store: LocalStore, lname: str, lot: Optional[str] = None) -> GridSheet:
    """不良要因シートの用意(UF1:3739-3750 と 印刷転送 Print:719-722)。

    `lot` を渡すと UF1 の判定: シートが無い、または B1 ≠ lot なら作り直す(原紙から)。
    `lot` が None なら 印刷転送 の判定: 無いときだけ作る。
    """
    name = line_sheet(DEFECT_SHEET, lname)
    if not store.exists(name) or (lot is not None and not variant_eq(store.sheet(name).get(1, 2), lot)):
        sh = store.add_sheet(name)
        paste_template_values(sh, "不良要因印刷原紙")
    return store.sheet(name)  # type: ignore[return-value]


def paste_template_values(sheet, template: str) -> None:
    """Pesute の値の部分を手元DB のシートへ(書式は帳票を作るときに原紙から当てる)。"""
    t = XSheet.from_template(template, sheet.name)
    for (r, c), x in t.cells.items():
        if x.value is not None:
            sheet.set(r, c, x.value)


def calculate_column_sums(sh: GridSheet) -> None:
    """CalculateColumnSums (Print:1337-1359)。C〜X 列の 3〜8 行を足して 10 行目に(Single)。

    B35: VBA は囲み数字を数と見ないので、印刷をやり直すと合計が消える。直す場合は ①〜㊿ も数として足す。
    """
    for col in range(3, 25):
        total = 0.0
        for row in range(3, 9):
            v = sh.get(row, col)
            if v is None:
                continue                                  # IsNumeric(Empty) は真。足しても 0
            if isinstance(v, (int, float)):
                total = vbx.single(total + float(v))
            elif vbx.is_numeric(v):
                total = vbx.single(total + vbx.cdbl(v))
            else:
                n = _circled_value(v)
                if n is not None:
                    total = vbx.single(total + n)
        sh.set(10, col, total)


def _circled_value(v: Any) -> Optional[int]:
    if not isinstance(v, str) or len(v) != 1:
        return None
    o = ord(v)
    if 0x2460 <= o <= 0x2473:
        return o - 0x2460 + 1
    if 0x3251 <= o <= 0x325F:
        return o - 0x3251 + 21
    if 0x32B1 <= o <= 0x32BF:
        return o - 0x32B1 + 36
    return None


def convert_numbers_to_string(sh: GridSheet) -> None:
    """ConvertNumbersToString (Print:1361-1396)。C3:AA8 と C10:AA10 の数を囲み数字に。

    VBA は数でないセルを書き戻しで**空にする**(B35 の原因)。直す場合は数でないセルをそのまま残す。
    10 行目の 0 は空。
    """
    def is_target(v: Value) -> bool:
        return v is not None and cell_text(v) != "" and vbx.is_numeric(v)

    def conv(v: Value) -> str:
        return circled_number(vbx.clng(v))

    block = [[sh.get(r, c) for c in range(3, 28)] for r in range(3, 9)]
    if any(is_target(v) for row in block for v in row):
        for i, row in enumerate(block):
            for j, v in enumerate(row):
                if is_target(v):
                    sh.set(3 + i, 3 + j, excel_value(conv(v), "V"))
    total = [sh.get(10, c) for c in range(3, 28)]
    flag2 = False
    out: list[Any] = [None] * len(total)          # None = VBA の ConvArr が Empty のまま
    for j, v in enumerate(total):
        if is_target(v):
            out[j] = conv(v)
            flag2 = True
        if vbx.compare_typed(v, 0, vbx.VAR, vbx.NUM) == 0:   # 数にならない文字列なら型エラー 13
            out[j] = ""
    if flag2:
        for j, v in enumerate(out):
            if v is None:
                continue                          # 直す場合: 数でないセルはそのまま
            sh.set(10, 3 + j, excel_value(v, "V"))


def defect_print_range(store: LocalStore, lname: str, worker: str, lot_text: str, *, now: datetime,
                       start_time: str = "", uf1_visible: bool = True, start1: str = "", end3: str = "",
                       end_time_caption: Optional[str] = None, mitsubishi: bool = False,
                       start_override: Optional[Value] = None) -> Optional[dict[str, Any]]:
    """不良要因印刷範囲 (Print:1278-1335)。ページ設定と、合計・囲み数字(手元DB のシートを書き換える)。

    - `mitsubishi`: VBA の `LotN<>""`(全ヵ所・キャスク等から呼ぶ形)。フッターを付けない
    - `start_override`: UFMitsubishi / UFCask が表示中のときの開始(全ヵ所検査!B1 / キャスク検査!B1)
    - `start1` / `end3`: 直の時間(伝送用ファイル `時間用`)。B38 を直す場合は `end_time_caption`(成績票の終了時刻)を使う
    返り値はページ設定(シートが無ければ None)。
    """
    name = line_sheet(DEFECT_SHEET, lname)
    if not store.exists(name):
        return None
    ws = store.sheet(name)
    start_a: Optional[datetime] = None
    if uf1_visible and start_time != "":
        start_a = _date_minute(start_time)
    if start_override is not None:
        start_a = _date_minute(start_override)
    now_s = now.replace(microsecond=0)
    t = now_s.time()
    if end_time_caption is None:
        s1 = vb_cdate(start1).time() if start1 != "" else None
        e3 = vb_cdate(end3).time() if end3 != "" else None
        if s1 is None or e3 is None:
            raise vbx.TypeMismatch("CDate(Start1/End3)")
        if t > s1 and t <= dtime(23, 59):
            end_t = now_s
        elif dtime(0, 0) <= t <= e3:
            end_t = now_s - timedelta(days=1)
        else:
            end_t = now_s
    else:
        end_t = vb_cdate(end_time_caption)
    n = ws.last_row(2)
    page = {"print_area": f"$A$1:$Y${n + 1}", "header_left": lname, "header_center": "",
            "header_right": f"{lot_text} {worker}", "footer_left": "", "footer_center": "",
            "footer_right": "", "header_font": "Meiryo UI", "header_size": 8,
            "margins_inch": {"left": 0, "right": 0, "top": 0, "bottom": 0, "header": 0, "footer": 0},
            "paper": "A4", "orientation": "landscape", "fit_wide": 1, "fit_tall": 1,
            "center_h": True, "center_v": True, "zoom": 100}
    if not mitsubishi:
        sa = vb_format_date(start_a if start_a is not None else datetime(1899, 12, 30), "yyyy/mm/dd  h:mm")
        page["footer_right"] = " " + sa + "～" + vb_format_date(end_t, "yyyy/mm/dd  h:mm") + " " + "印刷"
    calculate_column_sums(ws)
    convert_numbers_to_string(ws)
    store.save_report(name + "#page", json.dumps({"page": page, "area": [1, 1, n + 1, 25]}, ensure_ascii=False))
    return page


def _date_minute(v: Any) -> datetime:
    """`StartA = Format$(x, "yyyy/mm/dd h:mm")`(Date に入れるので秒が落ちる)。"""
    s = vb_format_date(v, "yyyy/mm/dd h:mm")
    return vb_cdate(s)


def render_defect_sheet(store: LocalStore, lname: str) -> Optional[ReportModel]:
    """手元DB の不良要因シートに原紙の書式を当てて帳票にする(最後に決めたページ設定つき)。"""
    name = line_sheet(DEFECT_SHEET, lname)
    if not store.exists(name):
        return None
    sh = XSheet.from_template("不良要因印刷原紙", name)
    for (r, c), x in list(sh.cells.items()):
        x.value = None
    grid = store.sheet(name)
    for r, c, v in grid.used():
        sh.cell(r, c).value = v
    meta = store.load_report(name + "#page")
    if meta:
        d = json.loads(meta)
        sh.page = d["page"]
        sh.print_area = tuple(d["area"])  # type: ignore[assignment]
    return sh.to_model("不良要因票")


# ------------------------------------------------------------------
# 検査成績票(印刷転送)
# ------------------------------------------------------------------
@dataclass
class PrintEnv:
    """印刷転送が UF1 以外から受け取るもの(他の担当の関数の結果)。"""
    now: datetime
    lateral_check: str = ""                 # OutTolerance.Lateralbow_Check の結果(未入力の位置名 or "")
    machine_transfer: Optional[Callable[[XSheet], None]] = None   # MachineTransfer.機器転送(sh,9,5,…)
    figure_sheet: bool = True               # シート「図」がある(xlsm にはある)
    worker_name_visible: bool = False       # UFWorkerName が表示中か
    worker_name_l1: str = ""                # UFWorkerName.L1.Caption(現検査員)
    start1: str = ""                        # 直の時間(不良要因票のフッター。B38)
    end3: str = ""
    uf1_visible: bool = True


def print_transfer(store: LocalStore, snap: KensaSnapshot, env: PrintEnv) -> PrintResult:
    """印刷転送 (Print:184-734)。検査成績票と不良要因票の中身を作る。

    途中で止まる(VBA の `PrintB.Tag = "stop"`)ときは `stopped=True` と MsgError の文言を返す。
    EndTime_Register をここでも呼ぶので、終了時刻(`end_time`)が更新される。
    """
    lname = snap.line_name
    msgs: list[Message] = []
    end_time = end_time_register(store, lname, env.now)          # Print:190
    snap.end_time = end_time
    sh = XSheet.from_template("印刷用原紙", line_sheet(PRINT_SHEET, lname))   # AddWS + Pesute (Print:191-192)
    sh.page = {"zoom": 100}

    a = {f: _visual_mark(snap.visual_up.get(f, "")) for f in FACES}
    b = {f: _visual_mark(snap.visual_under.get(f, "")) for f in FACES}
    vi = [[_visual_ab(a[f], b[f])] for f in FACES]

    ken_h = 0.0
    for i in range(6):                                          # Print:303-306(Single に足していく)
        ken_h = vbx.single(vbx.val(snap.ken_hbox[i]) + ken_h)
    yamasu = ""
    if vbx.is_numeric(snap.ken_hbox[6]):
        yamasu = snap.ken_hbox[6] + "枚"
    else:
        for k in range(6, 0, -1):
            if vbx.is_numeric(snap.ken_hbox[k - 1]):
                yamasu = f"{k}山"
                break
    sh.put(1, 6, snap.height_formula)

    at = [[uf_array(snap, i, f) for i in THICK_ROWS] for f in FACES]
    ht = [[uf_array(snap, i, f) for i in range(5, 15)] for f in FACES]
    tai = [[uf_array(snap, 17, f)] for f in FACES]
    angle = [[uf_array(snap, 18, f) if uf_array(snap, 18, f) != "" else None] for f in FACES]
    late = [[uf_array(snap, 20, f) + "/L" if uf_array(snap, 20, f) != "" else None] for f in FACES]

    shown = [[uf_array(snap, 15, f), uf_array(snap, 16, f)] for f in FACES]
    tags = uf_tag_array(snap)
    if snap.aorb == Face.A:
        hz_a, hz_b = shown, tags
    elif snap.aorb == Face.B:
        hz_a, hz_b = tags, shown
    else:
        raise vbx.TypeMismatch("AorB.Caption が Now_A / Now_B でない(Print:407 の LBound で型エラー)")
    hz = [[(("A:" + hz_a[i][k]) if hz_a[i][k] != "__" else hz_a[i][k]) + "\r\n"
           + (("B:" + hz_b[i][k]) if hz_b[i][k] != "__" else hz_b[i][k]) for k in range(2)] for i in range(6)]

    if snap.resize_tag == "有":
        sh.put(1, 2, "ｻｲｽﾞ修正")
    sh.put(2, 3, snap.lot)
    sh.put(4, 4, snap.atu)
    sh.put(4, 6, snap.hab)
    sh.put(4, 8, snap.tak)
    sh.put(5, 3, "'-" + snap.am)
    sh.put(5, 5, "'-" + snap.hm)
    sh.put(5, 7, "'-" + snap.tm)
    sh.put(5, 4, "'+" + snap.ap)
    sh.put(5, 6, "'+" + snap.hp)
    sh.put(5, 8, "'+" + snap.tp)
    sh.put(6, 3, snap.amk)
    sh.put(6, 5, snap.hmk)
    sh.put(6, 7, snap.tmk)
    sh.put(6, 4, snap.apk)
    sh.put(6, 6, snap.hpk)
    sh.put(6, 8, snap.tpk)
    sh.put(6, 9, snap.mg)
    sh.put(6, 10, snap.hahiz)
    sh.put(6, 11, snap.tahiz)
    sh.put(9, 17, snap.naka)
    sh.put(9, 18, snap.mimi)
    if snap.huris != "":                                         # Print:454-465
        sh.put(8, 12, vbx.fmt(vbx.val(snap.huris), "0.0"))
        sh.put(7, 12, "フリー反り")
        sh.borders_all(8, 12, 8, 12, linestyle="continuous")
        sh.border(7, 12, 8, 12, "top", linestyle="double")
        sh.border(7, 12, 8, 12, "right", weight="thick")
        sh.border(7, 12, 8, 12, "bottom", weight="thick")
        sh.border(7, 12, 8, 12, "left", weight="thick")
        sh.number_format(7, 12, 8, 12, "0.0")
    sh.put(6, 12, vbx.fmt(vbx.val(snap.taik_text), "0.0"))      # Print:467-482
    sh.put(4, 12, "対角")
    sh.border(5, 12, 5, 12, "top", linestyle="double")
    sh.border(4, 12, 6, 12, "top", weight="thick")
    sh.border(4, 12, 6, 12, "right", weight="thick")
    sh.border(4, 12, 6, 12, "left", weight="thick")
    sh.number_format(4, 12, 6, 12, "0.0")
    # B36: VBA は下罫線を耳歪(HizKousa(4))で決める。直すとフリー反り(HizKousa(6))で決める
    under_key = snap.huris
    if under_key != "":
        sh.border(4, 12, 6, 12, "bottom", linestyle="double")
    else:
        sh.border(4, 12, 6, 12, "bottom", weight="thick")
    sh.put(3, 9, snap.distortion_frame)
    sh.font_size(3, 9, 3, 9, 6)

    sh.put_range(12, 5, hz, kind="S")                           # 定盤・ゲージ (Print:490)
    sh.put_range(12, 7, at)                                     # 厚
    sh.put_range(19, 3, ht)                                     # 巾丈
    sh.put_range(19, 14, tai)                                   # 対角
    sh.put_range(19, 16, late)                                  # ラテラルボー
    sh.replace(12, 5, 17, 5, "_", "")
    sh.replace(12, 6, 17, 6, "_", "")
    if snap.youto == "R115":
        sh.put_range(19, 18, angle)
    sh.put_range(12, 3, vi)                                     # 外観
    for r in range(12, 18):
        x = sh.cell(r, 3)
        x.size = fit_text_font_size(cell_text(x.value), sh.widths.get(3, 11.625), sh.heights.get(r, 23.25))
    sh.put(2, 7, ken_h)
    if vbx.is_numeric(snap.ken_hbox[6]):
        sh.put(2, 6, "検入枚数")
        sh.put(2, 7, yamasu)
    else:
        sh.put(2, 8, yamasu)

    if snap.kind in ("厚板", "薄板", "少枚数", "極厚", "スラブ"):  # Print:527-538
        _print_borders(sh, snap.kind, snap, env)

    if env.machine_transfer is not None:                        # Print:541-544
        env.machine_transfer(sh)

    # 検査員 (Print:546-604)
    t_by = {"1": snap.t1, "2": snap.t2, "3": snap.t3}
    # Split("", "_") は空の配列、WorkeDayMax が "" なら Worker は Empty → どちらも IsEmptyArray が真
    worker_list = (t_by[snap.worker_shift].split("_")
                   if snap.worker_shift in t_by and t_by[snap.worker_shift] != "" else [])
    if not worker_list:
        msgs.append(Message("実行時エラー！検査員登録をやり直してください。(変数なし)", level=flow.WARN,
                            source="Print:554"))
        return PrintResult(stopped=True, end_time=end_time, messages=msgs)
    wt_name = line_sheet(WORKER_TIME, lname)
    if not store.exists(wt_name):
        sh.put(9, 2, worker_list[-1])
        raise Stop("実行時エラー '9': インデックスが有効範囲にありません。", source="Print:561", messages=msgs)
    wt = store.sheet(wt_name)
    names: dict[Any, int] = {}
    for row in wt.read_range(3, 1, 7, 12):
        for v in row:
            if v is not None and cell_text(v) != "":
                key = (type(v) is str, v)
                if key not in names:
                    names[key] = len(names) + 1
    if names:
        sh.put(9, 2, "_".join(cell_text(k[1]) for k in names))
    loaded = snap.data_loading == "Load" or variant_eq(wt.get(1, 1), "履歴呼び出し", vbx.STR)
    if not loaded:
        if cell_text(sh.get(9, 2)) == "":
            msgs.append(Message("検査員が空白です。WorkerTime上に履歴がない可能性有。 担当者に連絡してください。",
                                level=flow.WARN, source="Print:593"))
            return PrintResult(stopped=True, end_time=end_time, messages=msgs)
    else:
        sh.put(9, 2, env.worker_name_l1 if env.worker_name_visible else None)
        sh.put(3, 2, "履歴呼び出しから印刷")
    sh.cell(9, 2).shrink = True
    worker_text = cell_text(sh.get(9, 2))

    sh.put(4, 2, lname)                                          # Print:607
    ng = allow_check(snap.ng, snap.positions)                    # Print:609-620
    if ng != "":
        if "許容" in ng:
            sh.put(10, 2, "許容公差外在り" + ":" + ng.replace("許容", ""))
        elif "保留" in ng:
            sh.put(10, 2, "公差外 保留処理" + ":" + ng.replace("保留", ""))

    thick = vbx.val(cell_text(sh.get(4, 4)))                     # Print:625-629
    if thick > 74.0:
        sh.number_format(12, 7, 16, 10, "0.00_);[赤](0.00)")
    else:
        sh.number_format(12, 7, 16, 10, "0.000_);[赤](0.00)")

    # 公差外データ (Print:641-717)
    abn_name = line_sheet(ABNORMAL_SHEET, lname)
    if store.exists(abn_name):
        iz = store.sheet(abn_name)
        n = iz.last_row(1)
        if variant_eq(sh.get(2, 3), iz.get(1, 1)):
            if env.figure_sheet:
                sh.figure_at = "L11"
            lr = 26
            sh.put(lr, 1, "公差外ﾃﾞｰﾀ")
            if snap.kind == "極厚":
                field_date = EXTREME.labels()
                _copy_from(sh, iz, 2, 1, n, 14, lr + 1, 2)
            else:
                field_date = NORMAL.labels() + ["直角度", "ﾗﾃﾗﾙﾎﾞｰ", "定盤", "ｹﾞ-ｼﾞ"]
                _copy_from(sh, iz, 2, 1, n, 10, lr + 1, 2)
                _copy_from(sh, iz, 2, 15, n, 16, lr + 1, 12)
                _copy_from(sh, iz, 2, 17, n, 18, lr + 1, 14)
            sh.put_range(lr, 2, [field_date])
            _copy_from(sh, iz, 2, 19, n, 19, lr + 1, 1)
            last = sh.last_row(1)
            width = len(field_date) + 1
            sh.font_size(lr, 1, last, width, 8)
            sh.align(lr, 1, last, width, ha="left", va="center")
            if thick > 74.0:
                sh.number_format(lr + 1, 1, last, 5, "0.00_);[赤](0.00)")
                sh.number_format(lr + 1, 6, last, 11, "0.0_);[赤](0.0)")
            else:
                sh.number_format(lr + 1, 1, last, 5, "0.000_);[赤](0.00)")
                sh.number_format(lr + 1, 6, last, 11, "0.0_);[赤](0.0)")
            _page_area_abnormal(sh, snap, worker_text)
        else:
            # B36: VBA は C2(LOT)ではなく D4(厚)と比べる。一致するとページ設定をしない
            other = sh.get(2, 3)
            if not variant_eq(other, iz.get(1, 1)):
                _page_area_normal(sh, snap, worker_text)
    else:
        _page_area_normal(sh, snap, worker_text)

    # 不良要因票 (Print:719-726)
    defect_sheet = ensure_defect_sheet(store, lname)
    defect_sheet.write(1, 2, snap.lot, "O")
    defect_print_range(store, lname, worker_text, snap.lot, now=env.now, start_time=snap.start_time,
                       uf1_visible=env.uf1_visible, start1=env.start1, end3=env.end3, end_time_caption=end_time)

    sh.replace(12, 3, 16, 6, "\r\n", "\n")                       # Print:729-730
    report = sh.to_model("検査成績票")
    store.save_report(sh.name, report.to_json())
    return PrintResult(stopped=False, report=report, defect=render_defect_sheet(store, lname),
                       end_time=end_time, messages=msgs)


def _copy_from(dst: XSheet, src, r1: int, c1: int, r2: int, c2: int, dr: int, dc: int) -> None:
    """別シート(手元DB)の範囲を値だけ貼る(書式は異常ロット用シートに無いので値のみ)。範囲は Excel と同じく並べ替える。"""
    rows = src.read_range(r1, c1, r2, c2)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            dst.cell(dr + i, dc + j).value = v


def _footer(snap: KensaSnapshot) -> str:
    return (" " + vb_format_date(snap.start_time, "yyyy/mm/dd  h:mm") + "～"
            + vb_format_date(snap.end_time, "yyyy/mm/dd  h:mm") + " " + "印刷")


def _page_common(sh: XSheet, snap: KensaSnapshot, worker: str) -> dict[str, Any]:
    return {"header_left": snap.line_name, "header_center": "",
            "header_right": cell_text(sh.get(2, 3)) + " " + worker, "footer_left": "", "footer_center": "",
            "footer_right": _footer(snap), "header_font": "Meiryo UI", "header_size": 8,
            "margins_inch": {"left": 0, "right": 0, "top": 0.4, "bottom": 0, "header": 0, "footer": 0},
            "paper": "A4", "orientation": "landscape", "fit_wide": 1, "fit_tall": 1}


def _page_area_abnormal(sh: XSheet, snap: KensaSnapshot, worker: str) -> None:
    """印刷範囲 (Print:1137-1167)。公差外あり。倍率は指定なし(原紙のまま。U11)。"""
    n = sh.last_row(1)
    sh.print_area = (1, 1, n + 1, 16 if snap.kind == "極厚" else 15)
    sh.page.update(_page_common(sh, snap, worker))
    sh.page["print_area"] = f"$A$1:${'P' if snap.kind == '極厚' else 'O'}${n + 1}"


def _page_area_normal(sh: XSheet, snap: KensaSnapshot, worker: str) -> None:
    """印刷範囲2 (Print:1168-1200)。公差外なし。B1:L、倍率 115、水平中央。"""
    n = sh.last_row(2)
    sh.print_area = (1, 2, n + 1, 12)
    sh.page.update(_page_common(sh, snap, worker))
    sh.page.update({"print_area": f"$B$1:$L${n + 1}", "center_h": True, "zoom": 115})


def _print_borders(sh: XSheet, n: str, snap: KensaSnapshot, env: PrintEnv) -> None:
    """印刷用罫線 (Print:736-1136)。検査形態ごとの位置名・列見出し・罫線・ラテラルボー/直角度の列。"""
    ng_temp = "検査不要" if "ﾗﾃﾗﾙﾎﾞｰ検査不要" in snap.additional else ""
    flag_color = False

    def clea(r: int, c1: int, c2: int) -> None:
        sh.borders_all(r, c1, r, c2, linestyle="none")

    if snap.naka != "":
        sh.copy(8, 17, 9, 18, 7, 10)                             # Print:763

    def thick_plate_lines() -> None:
        sh.border(20, 2, 20, 9, "top", linestyle="continuous")
        sh.border(20, 2, 20, 9, "bottom", linestyle="continuous")
        sh.border(22, 2, 22, 9, "top", linestyle="continuous")
        sh.border(22, 2, 22, 9, "bottom", linestyle="continuous")
        sh.border(13, 2, 13, 6, "top", linestyle="continuous")
        sh.border(15, 2, 15, 6, "top", linestyle="continuous")
        sh.border(13, 2, 13, 6, "bottom", linestyle="continuous")
        sh.border(15, 2, 15, 6, "bottom", linestyle="continuous")

    def headers_e_to_i() -> None:
        for c, v in ((5, "丈1"), (6, "丈2"), (7, "対角A"), (8, "対角B"), (9, "差")):
            sh.put(18, c, v)

    def labels(top: Sequence[str], bottom: Sequence[str], top_row: int = 12, bottom_row: int = 19) -> None:
        for i, v in enumerate(top):
            sh.put(top_row + i, 2, v)
        for i, v in enumerate(bottom):
            sh.put(bottom_row + i, 2, v)

    def yama_is(k: int) -> bool:
        return cmp_str_num(snap.yama, k) == 0

    if n == "厚板":
        clea(20, 2, 10)
        clea(22, 2, 10)
        labels(["S", "S1", "2", "E1", "E"], ["S", "", "2", "", "E"])
        headers_e_to_i()
        sh.copy(18, 14, 24, 14, 18, 9)
        for r in (20, 22):
            sh.border(r, 9, r, 9, "left", linestyle="none")
            sh.border(r, 9, r, 9, "right", linestyle="none")
        thick_plate_lines()
    elif n == "少枚数":
        headers_e_to_i()
        sh.copy(18, 14, 24, 14, 18, 9)
    elif n == "薄板":
        labels(["S", "U", "2", "D", "E"], ["S", "U", "2", "D", "E"])
        headers_e_to_i()
        sh.copy(18, 14, 24, 14, 18, 9)
        if yama_is(4):
            labels(["4-S", "4-E", "'3-2", "'2-2", "1-S", "1-E"], ["4-S", "4-E", "'3-2", "'2-2", "1-S", "1-E"])
        elif yama_is(3):
            labels(["3-S", "3-E", "'2-2", "1-S", "1-E"], ["3-S", "3-E", "'2-2", "1-S", "1-E"])
        elif yama_is(2):
            labels(["2-S", "", "2-E", "", "1-E"], ["2-S", "", "2-E", "", "1-E"])
        elif yama_is(1):
            labels(["S", "", "2", "", "E"], ["S", "", "2", "", "E"])
        if snap.yama == "4":                                     # Print:873-899
            flag_color = snap.lateral_text_backcolor == RGB_FLAG and snap.lateral_text_backstyle == 1
            lascol = 0
            if snap.sumisei == "":
                lascol = 9
            if snap.sumisei != "" and ng_temp != "検査不要":
                lascol = 10
            if flag_color:
                lascol = 10
            if lascol == 0:
                # 直した側: 住精あり・ラテラル検査不要・指示色なしは、ラテラルボーの列なし(住精なしと同じ 9 列目まで)
                lascol = 9
            sh.copy(16, 2, 16, lascol, 17, 2, formats_only=True)                # B77: Lascol=0 だと 1004
            sh.border(17, 2, 17, lascol, "top", linestyle="dot")
            sh.copy(23, 2, 23, lascol, 24, 2, formats_only=True)
            sh.border(24, 2, 24, lascol, "top", linestyle="dot")
            sh.heights[24] = sh.heights.get(23, 23.25)
        if snap.inspe_butt_visible and str(snap.inspe_butt_caption).startswith("検入≧"):   # Print:902-917
            sh.put(14, 2, "")
            sh.put(21, 2, "")
            sh.copy(18, 14, 23, 14, 18, 10)
            sh.borders_all(13, 2, 15, 11, linestyle="none")
            sh.borders_all(20, 2, 22, 9, linestyle="none")
            sh.border(12, 2, 12, 11, "bottom", linestyle="continuous")
            sh.border(15, 2, 15, 11, "bottom", linestyle="continuous")
            sh.border(19, 2, 19, 10, "bottom", linestyle="continuous")
            sh.border(23, 2, 23, 10, "top", linestyle="continuous")
            sh.border(13, 11, 15, 11, "top", linestyle="none")
            sh.border(13, 11, 15, 11, "bottom", linestyle="none")
            _lateral_columns(sh, n, snap, env, ng_temp, flag_color)
            return
        if snap.usuita_caption == "T":                           # Print:919-960
            if snap.yama == "3":
                sh.borders_all(18, 12, 23, 12, linestyle="none")
                sh.clear_contents(18, 12, 23, 12)
                sh.copy(18, 14, 23, 14, 18, 9)
            elif snap.yama in ("2", "1") or snap.inspe_butt_visible:
                clea(20, 2, 10)
                clea(22, 2, 10)
                clea(13, 2, 10)
                clea(15, 2, 10)
                for r in (13, 15):
                    sh.border(r, 2, r, 10, "top", linestyle="continuous")
                    sh.border(r, 2, r, 10, "bottom", linestyle="continuous")
                sh.copy(18, 14, 23, 14, 18, 9)
                for r in (20, 22):
                    sh.border(r, 9, r, 9, "left", linestyle="none")
                    sh.border(r, 9, r, 9, "right", linestyle="none")
                thick_plate_lines()
            elif not snap.inspe_butt_visible:
                sh.copy(18, 14, 23, 14, 18, 9)
    elif n == "スラブ":                                          # Print:961-1031
        labels(["1", "", "2", "", ""], ["1上", "1下", "2上", "2下", ""])
        headers_e_to_i()
        sh.copy(18, 14, 23, 14, 18, 9)
        sh.border(20, 8, 20, 8, "left", linestyle="none")
        sh.border(22, 8, 22, 8, "left", linestyle="none")
        for r in (20, 22):
            sh.border(r, 9, r, 9, "left", linestyle="none")
            sh.border(r, 9, r, 9, "right", linestyle="none")
        sh.border(15, 2, 16, 10, "inside_h", linestyle="none")
        sh.border(15, 2, 16, 10, "inside_v", linestyle="none")
        sh.border(15, 2, 16, 10, "left", linestyle="none")
        sh.border(15, 2, 16, 10, "right", linestyle="none")
        sh.border(15, 2, 16, 10, "bottom", linestyle="none")
        sh.border(14, 2, 14, 10, "bottom", linestyle="continuous")
        sh.border(14, 2, 14, 10, "bottom", weight="medium")
        sh.border(23, 2, 23, 9, "inside_h", linestyle="none")
        sh.border(23, 2, 23, 9, "inside_v", linestyle="none")
        sh.border(23, 2, 23, 9, "bottom", linestyle="none")
        for r in (20, 22):
            sh.border(r, 7, r, 9, "top", linestyle="continuous")
        for r in (20, 22):
            sh.border(r, 7, r, 9, "bottom", linestyle="continuous")
        sh.border(22, 7, 22, 9, "bottom", linestyle="none")
        sh.border(23, 2, 23, 9, "left", linestyle="none")
        sh.border(23, 2, 23, 9, "right", linestyle="none")
        sh.border(22, 2, 22, 6, "bottom", linestyle="continuous")
        sh.border(22, 2, 22, 6, "bottom", weight="medium")
        sh.clear_contents(18, 10, 23, 12)
        for idx in ("left", "top", "bottom", "right", "inside_v", "inside_h"):
            sh.border(18, 10, 23, 12, idx, linestyle="none")
        sh.border(18, 9, 19, 9, "right", linestyle="continuous")
        sh.border(18, 9, 19, 9, "right", weight="thick")
        sh.border(21, 9, 21, 9, "right", linestyle="continuous")
        sh.border(21, 9, 21, 9, "right", weight="thick")
    else:                                                        # 耳付き (Print:1033-1035)
        sh.border(18, 9, 23, 9, "left", linestyle="continuous")
        sh.border(18, 9, 23, 9, "left", weight="medium")
    _lateral_columns(sh, n, snap, env, ng_temp, flag_color)


def _lateral_columns(sh: XSheet, n: str, snap: KensaSnapshot, env: PrintEnv, ng_temp: str,
                     flag_color: bool) -> None:
    """印刷用罫線の skipEnd 以降 (Print:1037-1135)。ラテラルボー/直角度の列と、歪の見出し。

    R115 の列(U10): VBA の順に処理すると、R115 用に J/K へ写した列を直後の「ラテ=""」の処理が消す。
    実機の印刷物で確かめるまで VBA のまま再現する(docs/04 U10、G6 §10-21)。
    """
    late = env.lateral_check
    extreme = snap.kind == "極厚"
    slab = snap.kind == "スラブ"
    if (snap.sumisei != "" and ng_temp != "検査不要") or flag_color or snap.lateral_text_tag == "有":
        if not extreme and not slab:
            if snap.youto == "R115":
                sh.copy(18, 16, 24, 16, 18, 11)
                sh.copy(18, 18, 24, 18, 18, 10)
                sh.borders_all(18, 12, 24, 14, linestyle="none")
                sh.clear_contents(18, 12, 24, 14)
            if late == "":
                sh.clear_contents(18, 10, 24, 14)
                sh.borders_all(18, 10, 23, 14, linestyle="none")
                sh.clear_contents(18, 10, 23, 14)
            if snap.youto != "R115":
                sh.copy(18, 16, 24, 16, 18, 10)
                sh.borders_all(18, 11, 24, 14, linestyle="none")
                sh.clear_contents(18, 11, 24, 14)
            if snap.yama == "4":
                sh.copy(19, 9, 19, 11, 24, 9, formats_only=True)
                sh.border(24, 9, 24, 11, "top", linestyle="dot")
                sh.border(24, 9, 24, 11, "bottom", weight="medium")
                sh.heights[24] = sh.heights.get(23, 23.25)
                if snap.youto != "R115":
                    sh.border(24, 11, 24, 11, "top", linestyle="none")
                    sh.border(24, 11, 24, 11, "bottom", linestyle="none")
            if n == "厚板":
                for r in (20, 22):
                    sh.border(r, 2, r, 10, "top", linestyle="continuous")
                for r in (20, 22):
                    sh.border(r, 2, r, 10, "bottom", linestyle="continuous")
    else:
        if n == "極厚":
            sh.borders_all(18, 14, 23, 16, linestyle="none")
            sh.clear_contents(18, 14, 23, 16)
        else:
            if snap.yama == "4":
                sh.borders_all(18, 10, 24, 14, linestyle="none")
                sh.clear_contents(18, 10, 24, 14)
            sh.borders_all(18, 10, 23, 14, linestyle="none")
            sh.clear_contents(18, 10, 23, 14)
        if snap.lateral_text_backcolor == RGB_FLAG and snap.lateral_text_backstyle == 1:
            if late == "" and snap.youto != "R115":
                if not extreme and not slab:
                    sh.clear_contents(18, 10, 24, 14)
                    sh.borders_all(18, 10, 23, 14, linestyle="none")
                    sh.clear_contents(18, 10, 23, 14)
                    sh.copy(18, 16, 24, 16, 18, 10)
                    sh.borders_all(18, 11, 24, 14, linestyle="none")
                    sh.clear_contents(18, 11, 24, 14)
        if ng_temp == "検査不要" and snap.kind == "薄板":
            sh.clear_contents(18, 10, 24, 14)
            sh.borders_all(18, 10, 23, 14, linestyle="none")
            sh.clear_contents(18, 10, 23, 14)
    if snap.lzs_caption == "幅ｹﾞｰｼﾞ":                            # Print:1127-1134
        sh.put(11, 5, "幅ゲージ")
        sh.put(11, 6, "丈ゲージ")
    elif snap.flatnes_text == "ﾌﾘｰｿﾘ要":
        sh.put(11, 6, "ﾌﾘｰｿﾘ")
        sh.replace(12, 6, 16, 6, "A:", "")
        sh.replace(12, 6, 16, 6, "B:", "")


# ------------------------------------------------------------------
# 原紙(xlsm の 印刷用原紙・不良要因印刷原紙 から写したもの)
#   1 行 1 セル: 番地|値(#は数)|表示形式の番号|文字の大きさ|太字|横位置|縦位置|折返し|縮小|上,下,左,右の線|フォント(M=Meiryo UI, P=ＭＳ Ｐゴシック)
# ------------------------------------------------------------------
_FORMATS = ('General', '0.0_ ', '0.000_ ', '0.00_ ', '0.000_);[Red]\\(0.000\\)', '0_ ', '@')
_TEMPLATES = {
    '印刷用原紙': {
        "merges": ['C12:D12', 'C16:D16', 'C15:D15', 'B9:D9', 'C14:D14', 'C13:D13'],
        "widths": {'A': 9.0, 'B': 7.25, 'C': 11.625, 'D': 11.625, 'E': 11.625, 'F': 11.625, 'G': 11.625, 'H': 11.625, 'I': 11.625, 'J': 11.625, 'K': 11.625, 'L': 11.625, 'M': 9.0, 'N': 9.375, 'O': 9.375, 'P': 9.375, 'Q': 9.0, 'R': 9.0, 'S': 9.0, 'T': 9.0, 'U': 9.0},
        "heights": {1: 16.5, 2: 17.25, 3: 16.5, 4: 17.25, 5: 19.5, 6: 19.5, 7: 20.25, 8: 19.5, 9: 19.5, 10: 19.5, 11: 20.25, 12: 23.25, 13: 23.25, 14: 23.25, 15: 23.25, 16: 23.25, 17: 23.25, 18: 23.25, 19: 23.25, 20: 23.25, 21: 23.25, 22: 23.25, 23: 23.25, 24: 18.75, 25: 18.75, 26: 18.75},
        "cells": """
F1||0|9||left|center|||,,,|M
B2|LotNo|0|11||center|center|||medium,medium,medium,medium|M
C2||0|12||left|center|||medium,medium,medium,|M
D2||0|12|||center|||medium,medium,,medium|M
F2|検入高さ(mm)|0|9||center|center|||medium,medium,medium,medium|M
G2||1|10||center|center|||medium,medium,medium,medium|M
H2||1|10||center|center|||medium,medium,medium,medium|M
B3||0|10||left|center|||,,,|M
B4||0|11||center|center|||medium,medium,medium,medium|M
C4|厚|0|12||center|center|||medium,double,medium,|M
D4||2|11|||center|||medium,double,,medium|M
E4|巾|0|12||center|center|||medium,double,medium,|M
F4||1|11|||center|||medium,double,,medium|M
G4|丈|0|12||center|center|||medium,double,,|M
H4||1|11|||center|||medium,double,,medium|M
I4|歪|0|12||center|center|||medium,double,medium,|M
J4||0|6||left|top|||medium,double,,|M
K4||0|8||left|top|||medium,double,,medium|M
C5|ﾏｲﾅｽ|2|11||left|center|||,thin,medium,|M
D5|ﾌﾟﾗｽ|2|11||left|center|||,thin,hair,|M
E5|ﾏｲﾅｽ|1|11||left|center|||,thin,medium,|M
F5|ﾌﾟﾗｽ|1|11||left|center|||,thin,hair,medium|M
G5|ﾏｲﾅｽ|1|11||left|center|||,thin,,|M
H5|ﾌﾟﾗｽ|1|11||left|center|||,thin,hair,medium|M
I5|ｹﾞｰｼﾞ|1|11||left|center|||double,thin,medium,|M
J5|幅|1|11||left|center|||double,thin,hair,|M
K5|丈|0|11|||center|||double,thin,hair,medium|M
C6|ﾏｲﾅｽ実|2|11||left|center|||thin,medium,medium,|M
D6|ﾌﾟﾗｽ実|2|11||left|center|||thin,medium,hair,|M
E6|ﾏｲﾅｽ実|1|11||left|center|||thin,medium,medium,|M
F6|ﾌﾟﾗｽ実|1|11||left|center|||thin,medium,hair,medium|M
G6|ﾏｲﾅｽ実|1|11||left|center|||thin,medium,,|M
H6|ﾌﾟﾗｽ実|1|11||left|center|||thin,medium,hair,medium|M
I6||1|11||left|center|||,medium,medium,hair|M
J6||1|11||left|center|||,medium,hair,|M
K6||1|11||left|center|||,medium,hair,medium|M
B8|検査員|0|8|||top|||medium,dashed,medium,|M
C8||0|8|||top|||medium,dashed,,|M
D8||0|11|||center|||medium,dashed,,medium|M
E8|マイクロ|0|8||left|top|||medium,dashed,medium,|M
F8||0|11|||center|||medium,dashed,,medium|M
G8|コンベックス|0|8||left|top|||medium,dashed,,|M
H8||0|11|||center|||medium,dashed,,medium|M
I8||3|12||center|center|||,,,|M
Q8|中歪|1|11||left|center|||double,thin,medium,hair|M
R8|耳歪|1|11||left|center|||double,thin,hair,medium|M
B9||0|10||left|center|||dashed,medium,medium,medium|M
C9||0|11||||||dashed,medium,,|P
D9||0|11||||||dashed,medium,,medium|P
E9||0|10||left|center|||,medium,medium,|M
F9||0|10||left|center|||,medium,,medium|M
G9||0|10||left|center|||,medium,,|M
H9||0|10||left|center|||,medium,,medium|M
I9||1|12||center|center|||,,,|M
Q9||1|11||left|center|||,medium,medium,hair|M
R9||1|11||left|center|||,medium,hair,medium|M
B11||0|12||center|center|||,,,|M
C11|外観|0|11||center|center|||medium,double,medium,|M
D11||0|11|||center|||medium,double,,medium|M
E11|定盤歪|0|11||center|center|||medium,double,,|M
F11|ゲージ歪|0|11||center|center|||medium,double,hair,medium|M
G11|厚1|0|11||center|center|||medium,double,medium,hair|M
H11|厚2|0|11||center|center|||medium,double,hair,hair|M
I11|厚3|0|11||center|center|||medium,double,hair,hair|M
J11|厚4|0|11||center|center|||medium,double,hair,medium|M
K11||0|11||center|center|||,,,|M
L11||0|7||center|center|||,,,|M
B12|#1|0|12||center|center|||medium,dashed,medium,medium|M
C12||4|8||left|center|||,dashed,medium,medium|M
D12||0|11||||||,dashed,,medium|P
E12||4|7.5||left|top|||,dashed,,|M
F12||4|7.5||left|center|||,dashed,hair,medium|M
G12||4|11||left|center|||,dashed,medium,hair|M
H12||4|11||left|center|||,dashed,hair,hair|M
I12||4|11||left|center|||,dashed,hair,hair|M
J12||4|11||left|center|||,dashed,hair,medium|M
K12||5|11||center|center|||,,,|M
L12||0|8||center|center|||,,,|M
B13|#2|0|12||center|center|||dashed,dashed,medium,medium|M
C13||4|8||left|center|||dashed,dashed,medium,medium|M
D13||0|11||||||dashed,dashed,,medium|P
E13||4|7.5||left|top|||dashed,dashed,,|M
F13||4|7.5||left|center|||dashed,dashed,hair,medium|M
G13||4|11||left|center|||dashed,dashed,medium,hair|M
H13||4|11||left|center|||dashed,dashed,hair,hair|M
I13||4|11||left|center|||dashed,dashed,hair,hair|M
J13||4|11||left|center|||dashed,dashed,hair,medium|M
K13||5|11||center|center|||,,,|M
B14|#3|0|12||center|center|||dashed,dashed,medium,medium|M
C14||4|8||left|center|||dashed,dashed,medium,medium|M
D14||0|11||||||dashed,dashed,,medium|P
E14||4|7.5||left|top|||dashed,dashed,,|M
F14||4|7.5||left|center|||dashed,dashed,hair,medium|M
G14||4|11||left|center|||dashed,dashed,medium,hair|M
H14||4|11||left|center|||dashed,dashed,hair,hair|M
I14||4|11||left|center|||dashed,dashed,hair,hair|M
J14||4|11||left|center|||dashed,dashed,hair,medium|M
K14||5|11||center|center|||,,,|M
B15|#4|0|12||center|center|||dashed,dashed,medium,medium|M
C15||4|8||left|center|||dashed,dashed,medium,medium|M
D15||0|11||||||dashed,dashed,,medium|P
E15||4|7.5||left|top|||dashed,dashed,,|M
F15||4|7.5||left|center|||dashed,dashed,hair,medium|M
G15||4|11||left|center|||dashed,dashed,medium,hair|M
H15||4|11||left|center|||dashed,dashed,hair,hair|M
I15||4|11||left|center|||dashed,dashed,hair,hair|M
J15||4|11||left|center|||dashed,dashed,hair,medium|M
K15||5|11||center|center|||,,,|M
B16|#5|0|12||center|center|||dashed,medium,medium,medium|M
C16||4|8||left|center|||dashed,medium,medium,medium|M
D16||0|11||||||dashed,medium,,medium|P
E16||4|7.5||left|top|||dashed,medium,,|M
F16||4|7.5||left|center|||dashed,medium,hair,medium|M
G16||4|11||left|center|||dashed,medium,medium,hair|M
H16||4|11||left|center|||dashed,medium,hair,hair|M
I16||4|11||left|center|||dashed,medium,hair,hair|M
J16||4|11||left|center|||dashed,medium,hair,medium|M
K16||5|11||center|center|||,,,|M
B17||0|11||center|center|||,,,|M
C17||0|8||left|center|||,,,|M
E17||0|7.5||left|top|||,,,|M
F17||0|7.5||left|top|||,,,|M
G17||2|11||left|center|||,,,|M
H17||2|11||left|center|||,,,|M
I17||2|11||left|center|||,,,|M
J17||2|11||left|center|||,,,|M
C18|幅1|0|11||center|center|||medium,double,medium,hair|M
D18|幅2|0|11||center|center|||medium,double,hair,hair|M
E18|幅3|0|11||center|center|||medium,double,hair,hair|M
F18|幅4|0|11||center|center|||medium,double,hair,hair|M
G18|幅5|0|11||center|center|||medium,double,hair,hair|M
H18|幅6|0|11||center|center|||medium,double,hair,hair|M
I18|丈1|0|11||center|center|||medium,double,hair,hair|M
J18|丈2|0|11||center|center|||medium,double,hair,hair|M
K18|丈3|0|11||center|center|||medium,double,hair,hair|M
L18|丈4|0|11||center|center|||medium,double,hair,medium|M
N18|差|0|11||center|center|||medium,double,double,medium|M
P18|ﾗﾃﾗﾙﾎﾞｰ|0|11||center|center|||medium,double,medium,medium|M
R18|直角度|0|11||center|center|||medium,double,medium,medium|M
B19|#1|0|12||center|center|||medium,dashed,medium,medium|M
C19||1|11||left|center|||,dashed,medium,hair|M
D19||1|11||left|center|||,dashed,hair,hair|M
E19||1|11||left|center|||,dashed,hair,hair|M
F19||1|11||left|center|||,dashed,hair,hair|M
G19||1|11||left|center|||,dashed,hair,hair|M
H19||1|11||left|center|||,dashed,hair,hair|M
I19||1|11||left|center|||,dashed,hair,hair|M
J19||1|11||left|center|||,dashed,hair,hair|M
K19||1|11||left|center|||,dashed,hair,hair|M
L19||1|11||left|center|||,dashed,hair,medium|M
N19||1|11||center|center|||,dashed,double,medium|M
P19||1|11||center|center|||,dashed,medium,medium|M
R19||1|11||center|center|||,dashed,medium,medium|M
B20|#2|0|12||center|center|||dashed,dashed,medium,medium|M
C20||1|11||left|center|||dashed,dashed,medium,hair|M
D20||1|11||left|center|||dashed,dashed,hair,hair|M
E20||1|11||left|center|||dashed,dashed,hair,hair|M
F20||1|11||left|center|||dashed,dashed,hair,hair|M
G20||1|11||left|center|||dashed,dashed,hair,hair|M
H20||1|11||left|center|||dashed,dashed,hair,hair|M
I20||1|11||left|center|||dashed,dashed,hair,hair|M
J20||1|11||left|center|||dashed,dashed,hair,hair|M
K20||1|11||left|center|||dashed,dashed,hair,hair|M
L20||1|11||left|center|||dashed,dashed,hair,medium|M
N20||1|11||center|center|||dashed,dashed,double,medium|M
P20||1|11||center|center|||dashed,dashed,medium,medium|M
R20||1|11||center|center|||dashed,dashed,medium,medium|M
B21|#3|0|12||center|center|||dashed,dashed,medium,medium|M
C21||1|11||left|center|||dashed,dashed,medium,hair|M
D21||1|11||left|center|||dashed,dashed,hair,hair|M
E21||1|11||left|center|||dashed,dashed,hair,hair|M
F21||1|11||left|center|||dashed,dashed,hair,hair|M
G21||1|11||left|center|||dashed,dashed,hair,hair|M
H21||1|11||left|center|||dashed,dashed,hair,hair|M
I21||1|11||left|center|||dashed,dashed,hair,hair|M
J21||1|11||left|center|||dashed,dashed,hair,hair|M
K21||1|11||left|center|||dashed,dashed,hair,hair|M
L21||1|11||left|center|||dashed,dashed,hair,medium|M
N21||1|11||center|center|||dashed,dashed,double,medium|M
P21||1|11||center|center|||dashed,dashed,medium,medium|M
R21||1|11||center|center|||dashed,dashed,medium,medium|M
B22|#4|0|12||center|center|||dashed,dashed,medium,medium|M
C22||1|11||left|center|||dashed,dashed,medium,hair|M
D22||1|11||left|center|||dashed,dashed,hair,hair|M
E22||1|11||left|center|||dashed,dashed,hair,hair|M
F22||1|11||left|center|||dashed,dashed,hair,hair|M
G22||1|11||left|center|||dashed,dashed,hair,hair|M
H22||1|11||left|center|||dashed,dashed,hair,hair|M
I22||1|11||left|center|||dashed,dashed,hair,hair|M
J22||1|11||left|center|||dashed,dashed,hair,hair|M
K22||1|11||left|center|||dashed,dashed,hair,hair|M
L22||1|11||left|center|||dashed,dashed,hair,medium|M
N22||1|11||center|center|||dashed,dashed,double,medium|M
P22||1|11||center|center|||dashed,dashed,medium,medium|M
R22||1|11||center|center|||dashed,dashed,medium,medium|M
B23|#5|0|12||center|center|||dashed,medium,medium,medium|M
C23||1|11||left|center|||dashed,medium,medium,hair|M
D23||1|11||left|center|||dashed,medium,hair,hair|M
E23||1|11||left|center|||dashed,medium,hair,hair|M
F23||1|11||left|center|||dashed,medium,hair,hair|M
G23||1|11||left|center|||dashed,medium,hair,hair|M
H23||1|11||left|center|||dashed,medium,hair,hair|M
I23||1|11||left|center|||dashed,medium,hair,hair|M
J23||1|11||left|center|||dashed,medium,hair,hair|M
K23||1|11||left|center|||dashed,medium,hair,hair|M
L23||1|11||left|center|||dashed,medium,hair,medium|M
N23||1|11||center|center|||dashed,medium,double,medium|M
P23||1|11||center|center|||dashed,medium,medium,medium|M
R23||1|11||center|center|||dashed,medium,medium,medium|M
B24||0|12||center|center|||,,,|M
C24||1|11||left|center|||,,,|M
D24||1|11||left|center|||,,,|M
E24||1|11||left|center|||,,,|M
F24||1|11||left|center|||,,,|M
G24||1|11||left|center|||,,,|M
H24||1|11||left|center|||,,,|M
I24||1|11||left|center|||,,,|M
J24||1|11||left|center|||,,,|M
K24||1|11||left|center|||,,,|M
L24||1|11||left|center|||,,,|M
P24||1|11||center|center|||,,,|M
""",
    },
    '不良要因印刷原紙': {
        "merges": ['C1:D1', 'I1:M1', 'N1:V1', 'E1:H1'],
        "widths": {'A': 3.0, 'B': 10.25, 'C': 5.875, 'D': 5.875, 'E': 5.875, 'F': 5.875, 'G': 5.875, 'H': 5.875, 'I': 5.875, 'J': 5.875, 'K': 5.875, 'L': 5.875, 'M': 5.875, 'N': 5.875, 'O': 5.875, 'P': 5.875, 'Q': 5.875, 'R': 5.875, 'S': 5.875, 'T': 5.875, 'U': 5.875, 'V': 5.875, 'W': 5.875, 'X': 6.0, 'Y': 3.375, 'Z': 9.0},
        "heights": {1: 23.25, 2: 91.5, 3: 44.25, 4: 44.25, 5: 44.25, 6: 44.25, 7: 44.25, 8: 44.25, 9: 22.5, 10: 23.25, 11: 33.0},
        "cells": """
B1||0|10||center|center|||,,,|M
C1|鋳造|0|14||center||||thin,thin,thin,thin|M
D1||0|11||||||thin,thin,,thin|P
E1|熱延|0|14||center||||thin,thin,thin,thin|M
F1||0|11||||||thin,thin,,|P
G1||0|11||||||thin,thin,,|P
H1||0|11||||||thin,thin,,thin|P
I1|冷延|0|14||center||||thin,thin,thin,thin|M
J1||0|11||||||thin,thin,,|P
K1||0|11||||||thin,thin,,|P
L1||0|11||||||thin,thin,,|P
M1||0|11||||||thin,thin,,thin|P
N1|仕上|0|14||center||||thin,thin,thin,thin|M
O1||0|11||||||thin,thin,,|P
P1||0|11||||||thin,thin,,|P
Q1||0|11||||||thin,thin,,|P
R1||0|11||||||thin,thin,,|P
S1||0|11||||||thin,thin,,|P
T1||0|11||||||thin,thin,,|P
U1||0|11||||||thin,thin,,|P
V1||0|11||||||thin,thin,,thin|P
W1||0|10|||center|||thin,,thin,thin|M
X1||0|10|||center|||thin,,thin,thin|M
C2|フクレ|0|9||center|center|||thin,thin,thin,hair|M
D2|　|0|9||center|center|||thin,thin,hair,|M
E2|ムシレ|0|9||center|center|||thin,thin,thin,hair|M
F2|汚|0|9||center|center|||thin,thin,hair,hair|M
G2|リックアップ|0|9||center|center|||thin,thin,hair,hair|M
H2| |0|9||center|center|||thin,thin,hair,thin|M
I2|圧込|0|9||center|center|||thin,thin,,hair|M
J2|汚|0|9||center|center|||thin,thin,hair,hair|M
K2|フラットネス|0|9||center|center|||thin,thin,hair,hair|M
L2|厚|0|9||center|center|||thin,thin,hair,hair|M
M2| |0|9||center|center|||thin,thin,hair,|M
N2|汚|0|9||center|center|||thin,thin,thin,hair|M
O2|飛込|0|9||center|center|||thin,thin,hair,hair|M
P2|当り|0|9||center|center|||thin,thin,hair,hair|M
Q2|擢|0|9||center|center|||thin,thin,hair,hair|M
R2|付着油|0|9||center|center|||thin,thin,hair,hair|M
S2|切断面|0|9||center|center|||thin,thin,hair,hair|M
T2|巾丈|0|9||center|center|||thin,thin,hair,hair|M
U2|特寸|0|9||center|center|||thin,thin,hair,hair|M
V2| |0|9||center|center|||thin,thin,hair,thin|M
W2|他|0|10||center|center|||,thin,thin,thin|M
X2|TP|0|10||center|center|||,thin,thin,thin|M
B3||0|11||center|center|||thin,thin,thin,thin|M
C3||0|11||center|center|||thin,dotted,thin,hair|M
D3||0|11||center|center|||thin,dotted,hair,|M
E3||0|11||center|center|||thin,dotted,thin,hair|M
F3||0|11||center|center|||thin,dotted,hair,hair|M
G3||0|11||center|center|||thin,dotted,hair,hair|M
H3||0|11||center|center|||thin,dotted,hair,thin|M
I3||0|11||center|center|||thin,dotted,,hair|M
J3||0|11||center|center|||thin,dotted,hair,hair|M
K3||0|11||center|center|||thin,dotted,hair,hair|M
L3||0|11||center|center|||thin,dotted,hair,hair|M
M3||0|11||center|center|||thin,dotted,hair,|M
N3||0|11||center|center|||thin,dotted,thin,hair|M
O3||0|11||center|center|||thin,dotted,hair,hair|M
P3||0|11||center|center|||thin,dotted,hair,hair|M
Q3||0|11||center|center|||thin,dotted,hair,hair|M
R3||0|11||center|center|||thin,dotted,hair,hair|M
S3||0|11||center|center|||thin,dotted,hair,hair|M
T3||0|11||center|center|||thin,dotted,hair,hair|M
U3||0|11||center|center|||thin,dotted,hair,hair|M
V3||0|11||center|center|||thin,dotted,hair,thin|M
W3||0|11||center|center|||thin,dotted,thin,thin|M
X3||0|11||center|center|||thin,dotted,thin,thin|M
B4||0|11||center|center|||thin,thin,thin,thin|M
C4||0|11||center|center|||dotted,dotted,thin,hair|M
D4||0|11||center|center|||dotted,dotted,hair,|M
E4||0|11||center|center|||dotted,dotted,thin,hair|M
F4||0|11||center|center|||dotted,dotted,hair,hair|M
G4||0|11||center|center|||dotted,dotted,hair,hair|M
H4||0|11||center|center|||dotted,dotted,hair,thin|M
I4||0|11||center|center|||dotted,dotted,,hair|M
J4||0|11||center|center|||dotted,dotted,hair,hair|M
K4||0|11||center|center|||dotted,dotted,hair,hair|M
L4||0|11||center|center|||dotted,dotted,hair,hair|M
M4||0|11||center|center|||dotted,dotted,hair,|M
N4||0|11||center|center|||dotted,dotted,thin,hair|M
O4||0|11||center|center|||dotted,dotted,hair,hair|M
P4||0|11||center|center|||dotted,dotted,hair,hair|M
Q4||0|11||center|center|||dotted,dotted,hair,hair|M
R4||0|11||center|center|||dotted,dotted,hair,hair|M
S4||0|11||center|center|||dotted,dotted,hair,hair|M
T4||0|11||center|center|||dotted,dotted,hair,hair|M
U4||0|11||center|center|||dotted,dotted,hair,hair|M
V4||0|11||center|center|||dotted,dotted,hair,thin|M
W4||0|11||center|center|||dotted,dotted,thin,thin|M
X4||0|11||center|center|||dotted,dotted,thin,thin|M
B5||0|11||center|center|||thin,thin,thin,thin|M
C5||0|11||center|center|||dotted,dotted,thin,hair|M
D5||0|11||center|center|||dotted,dotted,hair,|M
E5||0|11||center|center|||dotted,dotted,thin,hair|M
F5||0|11||center|center|||dotted,dotted,hair,hair|M
G5||0|11||center|center|||dotted,dotted,hair,hair|M
H5||0|11||center|center|||dotted,dotted,hair,thin|M
I5||0|11||center|center|||dotted,dotted,,hair|M
J5||0|11||center|center|||dotted,dotted,hair,hair|M
K5||0|11||center|center|||dotted,dotted,hair,hair|M
L5||0|11||center|center|||dotted,dotted,hair,hair|M
M5||0|11||center|center|||dotted,dotted,hair,|M
N5||0|11||center|center|||dotted,dotted,thin,hair|M
O5||0|11||center|center|||dotted,dotted,hair,hair|M
P5||0|11||center|center|||dotted,dotted,hair,hair|M
Q5||0|11||center|center|||dotted,dotted,hair,hair|M
R5||0|11||center|center|||dotted,dotted,hair,hair|M
S5||0|11||center|center|||dotted,dotted,hair,hair|M
T5||0|11||center|center|||dotted,dotted,hair,hair|M
U5||0|11||center|center|||dotted,dotted,hair,hair|M
V5||0|11||center|center|||dotted,dotted,hair,thin|M
W5||0|11||center|center|||dotted,dotted,thin,thin|M
X5||0|11||center|center|||dotted,dotted,thin,thin|M
B6||0|11||center|center|||thin,thin,thin,thin|M
C6||0|11||center|center|||dotted,dotted,thin,hair|M
D6||0|11||center|center|||dotted,dotted,hair,|M
E6||0|11||center|center|||dotted,dotted,thin,hair|M
F6||0|11||center|center|||dotted,dotted,hair,hair|M
G6||0|11||center|center|||dotted,dotted,hair,hair|M
H6||0|11||center|center|||dotted,dotted,hair,thin|M
I6||0|11||center|center|||dotted,dotted,,hair|M
J6||0|11||center|center|||dotted,dotted,hair,hair|M
K6||0|11||center|center|||dotted,dotted,hair,hair|M
L6||0|11||center|center|||dotted,dotted,hair,hair|M
M6||0|11||center|center|||dotted,dotted,hair,|M
N6||0|11||center|center|||dotted,dotted,thin,hair|M
O6||0|11||center|center|||dotted,dotted,hair,hair|M
P6||0|11||center|center|||dotted,dotted,hair,hair|M
Q6||0|11||center|center|||dotted,dotted,hair,hair|M
R6||0|11||center|center|||dotted,dotted,hair,hair|M
S6||0|11||center|center|||dotted,dotted,hair,hair|M
T6||0|11||center|center|||dotted,dotted,hair,hair|M
U6||0|11||center|center|||dotted,dotted,hair,hair|M
V6||0|11||center|center|||dotted,dotted,hair,thin|M
W6||0|11||center|center|||dotted,dotted,thin,thin|M
X6||0|11||center|center|||dotted,dotted,thin,thin|M
B7||0|11||center|center|||thin,thin,thin,thin|M
C7||0|11||center|center|||dotted,dashed,thin,hair|M
D7||0|11||center|center|||dotted,dashed,hair,|M
E7||0|11||center|center|||dotted,dashed,thin,hair|M
F7||0|11||center|center|||dotted,dashed,hair,hair|M
G7||0|11||center|center|||dotted,dashed,hair,hair|M
H7||0|11||center|center|||dotted,dashed,hair,thin|M
I7||0|11||center|center|||dotted,dashed,,hair|M
J7||0|11||center|center|||dotted,dashed,hair,hair|M
K7||0|11||center|center|||dotted,dashed,hair,hair|M
L7||0|11||center|center|||dotted,dashed,hair,hair|M
M7||0|11||center|center|||dotted,dashed,hair,|M
N7||0|11||center|center|||dotted,dashed,thin,hair|M
O7||0|11||center|center|||dotted,dashed,hair,hair|M
P7||0|11||center|center|||dotted,dashed,hair,hair|M
Q7||0|11||center|center|||dotted,dashed,hair,hair|M
R7||0|11||center|center|||dotted,dashed,hair,hair|M
S7||0|11||center|center|||dotted,dashed,hair,hair|M
T7||0|11||center|center|||dotted,dashed,hair,hair|M
U7||0|11||center|center|||dotted,dashed,hair,hair|M
V7||0|11||center|center|||dotted,dashed,hair,thin|M
W7||0|11||center|center|||dotted,dashed,thin,thin|M
X7||0|11||center|center|||dotted,dashed,thin,thin|M
B8||0|11||center|center|||thin,thin,thin,thin|M
C8||0|11||center|center|||,double,thin,hair|M
D8||0|11||center|center|||,double,hair,|M
E8||0|11||center|center|||,double,thin,hair|M
F8||0|11||center|center|||,double,hair,hair|M
G8||0|11||center|center|||,double,hair,hair|M
H8||0|11||center|center|||,double,hair,thin|M
I8||0|11||center|center|||,double,,hair|M
J8||0|11||center|center|||,double,hair,hair|M
K8||0|11||center|center|||,double,hair,hair|M
L8||0|11||center|center|||,double,hair,hair|M
M8||0|11||center|center|||,double,hair,|M
N8||0|11||center|center|||,double,thin,hair|M
O8||0|11||center|center|||,double,hair,hair|M
P8||0|11||center|center|||,double,hair,hair|M
Q8||0|11||center|center|||,double,hair,hair|M
R8||0|11||center|center|||,double,hair,hair|M
S8||0|11||center|center|||,double,hair,hair|M
T8||0|11||center|center|||,double,hair,hair|M
U8||0|11||center|center|||,double,hair,hair|M
V8||0|11||center|center|||,double,hair,thin|M
W8||0|11||center|center|||,double,,thin|M
X8||0|11||center|center|||,double,,thin|M
C9|01|6|11||center|center|||double,thin,thin,hair|M
D9|02|6|11||center|center|||double,thin,hair,|M
E9|11|6|11||center|center|||double,thin,thin,hair|M
F9|12|6|11||center|center|||double,thin,hair,hair|M
G9|13|6|11||center|center|||double,thin,hair,hair|M
H9|14|6|11||center|center|||double,thin,hair,thin|M
I9|21|6|11||center|center|||double,thin,,hair|M
J9|22|6|11||center|center|||double,thin,hair,hair|M
K9|23|6|11||center|center|||double,thin,hair,hair|M
L9|24|6|11||center|center|||double,thin,hair,hair|M
M9|25|6|11||center|center|||double,thin,hair,|M
N9|31|6|11||center|center|||double,thin,thin,hair|M
O9|32|6|11||center|center|||double,thin,hair,hair|M
P9|33|6|11||center|center|||double,thin,hair,hair|M
Q9|34|6|11||center|center|||double,thin,hair,hair|M
R9|35|6|11||center|center|||double,thin,hair,hair|M
S9|36|6|11||center|center|||double,thin,hair,hair|M
T9|37|6|11||center|center|||double,thin,hair,hair|M
U9|38|6|11||center|center|||double,thin,hair,hair|M
V9|39|6|11||center|center|||double,thin,hair,thin|M
W9|41|6|11||center|center|||double,thin,,thin|M
X9||6|11||center|center|||double,thin,,thin|M
B10|不良合計|0|10||center|center|||,,,|M
C10||0|11||center|center|||thin,thin,thin,hair|M
D10||0|11||center|center|||thin,thin,hair,|M
E10||0|11||center|center|||thin,thin,thin,hair|M
F10||0|11||center|center|||thin,thin,hair,hair|M
G10||0|11||center|center|||thin,thin,hair,hair|M
H10||0|11||center|center|||thin,thin,hair,thin|M
I10||0|11||center|center|||thin,thin,,hair|M
J10||0|11||center|center|||thin,thin,hair,hair|M
K10||0|11||center|center|||thin,thin,hair,hair|M
L10||0|11||center|center|||thin,thin,hair,hair|M
M10||0|11||center|center|||thin,thin,hair,|M
N10||0|11||center|center|||thin,thin,thin,hair|M
O10||0|11||center|center|||thin,thin,hair,hair|M
P10||0|11||center|center|||thin,thin,hair,hair|M
Q10||0|11||center|center|||thin,thin,hair,hair|M
R10||0|11||center|center|||thin,thin,hair,hair|M
S10||0|11||center|center|||thin,thin,hair,hair|M
T10||0|11||center|center|||thin,thin,hair,hair|M
U10||0|11||center|center|||thin,thin,hair,hair|M
V10||0|11||center|center|||thin,thin,hair,thin|M
W10||0|11||center|center|||thin,thin,,thin|M
X10||0|11||center|center|||thin,thin,,thin|M
B11|スポット|0|11||center|center|||,,,|M
C11||0|11||left|center|||,,,|M
D11||0|11||left|center|||,,,|M
""",
    },
}
