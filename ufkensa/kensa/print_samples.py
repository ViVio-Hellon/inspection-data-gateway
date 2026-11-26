"""帳票の見本(保守 → 帳票の見本。2026-11-26 現場の声)。

「検査のすべてのパターンで、測定画面と印刷フォーマットがセットで正しく出るか」を、現物のロット無しで確かめる。

- 測定画面は**本物の作業選択**(layout.select_work・yama_change・at_butt_click・inspe_butt_click、
  ラテラル・直角度・フリー反りは lselect_lateral・right_angle_display・ﾌﾘｰｿﾘ要)で作る。ロットは見本の値
  (仕掛台帳は読まない。hooks は見本の DB を返す)
- 見えている測定欄に見本の値(公差の中の値)を入れ、**本物の印刷転送**(report_model.print_transfer)で帳票を作る
- 書くのは使い捨ての手元DB(一時フォルダ)だけ。本物の検査・保存・測定DB には何も書かない
"""
from __future__ import annotations

import itertools
import math
import shutil
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

from . import grid, layout, report_html, report_model, view, vbx, wiring
from .grid import KensaState
from .store import LocalStore

LINE = "Test"
LOT = "SAMPLE0"

#: 作業(画面のボタン)→ 内部の作業名・検番の頭・見本の寸法(厚・幅・丈)
WORKS: dict[str, tuple[str, str, float, float, float]] = {
    "H(厚板)": ("厚板", "H", 50.0, 1528.0, 3053.0),
    "少枚数": ("少枚数", "H", 10.0, 1250.0, 2500.0),
    "スラブ": ("スラブ", "H", 250.0, 1250.0, 2500.0),
    "耳付(極厚)": ("極厚", "H", 120.0, 1250.0, 2500.0),
    "N(薄板)": ("薄板", "N", 2.0, 1000.0, 2000.0),
    "L(薄板)": ("薄板", "L", 2.0, 1000.0, 2000.0),
    "T(薄板)": ("薄板", "T", 2.0, 1000.0, 2000.0),
}
YAMA = ("1", "2", "3", "4")
AT = ("4", "2")              # 厚4ヵ所 / 厚2ヵ所
INSPE = ("ge26", "le25")     # 検入≧26 / 検入≦25
LOT_FIELDS = ["検査番号", "ﾛｯﾄ番号", "ｵｰﾀﾞｰ板厚", "設計_設備ｺｰｽ", "実績_設備ｺｰｽ"]
HIKI_FIELDS = ["ﾛｯﾄ番号", "x", "検査番号"]


@dataclass
class Pattern:
    work: str = "H(厚板)"
    yama: str = "1"          # 薄板だけ
    at: str = "4"            # 薄板・厚板で「厚2ヵ所」ボタンが出るとき
    inspe: str = "ge26"      # 薄板(山数 1・2)で「検入≦25」ボタンが出るとき
    lateral: bool = False
    right: bool = False      # 直角度
    free: bool = False       # フリー反り
    sheets: int = 3          # 少枚数の枚数

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "Pattern":
        p = cls()
        for k in ("work", "yama", "at", "inspe"):
            if d.get(k) not in (None, ""):
                setattr(p, k, str(d[k]))
        for k in ("lateral", "right", "free"):
            setattr(p, k, str(d.get(k, "")).lower() in ("1", "true", "on", "有"))
        if str(d.get("sheets", "")).isdigit():
            p.sheets = max(1, min(5, int(d["sheets"])))
        if p.work not in WORKS:
            raise ValueError(f"作業が違います: {p.work}")
        return p

    def label(self) -> str:
        name, ken = WORKS[self.work][0], WORKS[self.work][1]
        parts = [self.work]
        if name == "薄板":
            parts += [f"{self.yama}山", "厚2ヵ所" if self.at == "2" else "厚4ヵ所",
                      "検入≦25" if self.inspe == "le25" else "検入≧26"]
        elif name == "厚板":
            parts.append("厚2ヵ所" if self.at == "2" else "厚4ヵ所")
        elif name == "少枚数":
            parts.append(f"{self.sheets}枚")
        parts += [f"ラテラル{'有' if self.lateral else '無'}", f"直角度{'有' if self.right else '無'}",
                  f"フリー反り{'有' if self.free else '無'}"]
        return "・".join(parts)


class SampleHooks(grid.Hooks):
    """見本の DB を返す呼び出し口(本物の仕掛台帳・品質基準は読まない)。"""

    def __init__(self, ken: str, atu: float, sheets: float) -> None:
        self.ken_no = ken + "6X0001"
        self.atu = atu
        self.sheets = sheets

    def calculate_sheets(self, state, atu):
        return vbx.single(self.sheets)

    def register_check(self, state):
        return False

    def is_over_12_hours(self, state):
        return False

    def get_lot_info(self, state, lot):
        return None

    def db_return_item(self, state, *args):
        return ""

    def is_gauge_required(self, *args):
        return False

    def sika_atu_create(self, state, lot):
        return None

    def arrbox(self, state, target, sql_filter="", use="Record"):
        if use == "Fields":
            return {"SIKALOTNOW.accdb": LOT_FIELDS, "SIKAHIKINOW.accdb": HIKI_FIELDS}.get(target)
        if target == "SIKALOTNOW.accdb":
            return [[self.ken_no, LOT, vbx.fmt(self.atu, "0.000"), "HOT L-1 KEN", ""]]
        return None

    def atu_hiki_records(self, state, lot):
        return True, None

    def small_number_count(self, state, th):
        wiring.RealHooks.small_number_count(self, state, th)        # type: ignore[arg-type]


def _tol(state: KensaState, atu: float, hab: float, tak: float) -> None:
    f3, f1 = (lambda x: vbx.fmt(x, "0.000")), (lambda x: vbx.fmt(x, "0.0"))
    state.atu, state.ap, state.am = f3(atu), "0.300", "0.300"
    state.apk, state.amk = f3(atu + 0.3), f3(atu - 0.3)
    state.hab, state.hp, state.hm = f1(hab), "3.0", "3.0"
    state.hpk, state.hmk = f1(hab + 3), f1(hab - 3)
    state.tak, state.tp, state.tm = f1(tak), "3.0", "3.0"
    state.tpk, state.tmk = f1(tak + 3), f1(tak - 3)


def build_state(p: Pattern) -> tuple[KensaState, SampleHooks]:
    """測定画面の状態を、本物の作業選択で作る。"""
    name, ken, atu, hab, tak = WORKS[p.work]
    s = KensaState(line_name=LINE, lot=LOT, proofread_flag="ok", ken_no_caption=ken + "6X0001")
    s.work_caption["厚板"] = ken if name != "薄板" else "H"
    s.work_caption["薄板"] = ken
    s.zai, s.tyu, s.youto = "5052", "H34", "H207"
    _tol(s, atu, hab, tak)
    h = SampleHooks(ken, atu, p.sheets)
    if name == "薄板":
        layout.select_work(s, "薄板", h)
        s.yama = p.yama
        layout.yama_change(s, h)
        if p.at == "2" and s.at_butt_visible and s.at_butt_caption == "厚2ヵ所":
            layout.at_butt_click(s, h)
        if p.inspe == "le25" and s.inspe_butt_visible and s.inspe_butt_caption == "検入≦25":
            layout.inspe_butt_click(s, h)
    else:
        layout.select_work(s, name, h)
        if name == "厚板" and p.at == "2" and s.at_butt_visible and s.at_butt_caption == "厚2ヵ所":
            layout.at_butt_click(s, h)
    # ロットの引当で「要」になったときと同じ表示(layout._lselect_allocations)
    if p.lateral:
        layout.lselect_lateral(s, s.ken_no_caption[:1])
    if p.right:
        layout.right_angle_display(s, s.ken_no_caption[:1])
    if p.free:
        s.flatnes_text = "ﾌﾘｰｿﾘ要"
    return s, h


def _sample_value(label: str, s: KensaState) -> str:
    """欄の名前から、公差の中の見本の値。"""
    if "幅" in label or "巾" in label:
        return vbx.fmt(vbx.val(s.hab) + 0.5, "0.0")
    if "丈" in label or "長" in label:
        return vbx.fmt(vbx.val(s.tak) + 0.5, "0.0")
    if "対角" in label:
        return vbx.fmt(math.hypot(vbx.val(s.hab), vbx.val(s.tak)), "0.0")
    return vbx.fmt(vbx.val(s.atu) + 0.012, "0.000")


def fill_values(s: KensaState) -> int:
    """見えていて入れられる測定欄に見本の値を入れる。入れた数を返す。"""
    v = view.build(s)
    n = 0
    for row in v["grid"]["rows"]:
        if not row["visible"]:
            continue
        for cell in row["cells"]:
            if cell.get("name") and cell["visible"] and cell.get("enabled", True):
                s.cells[cell["name"]].value = _sample_value(row["label"], s)
                n += 1
    for d in (s.diag, s.lateral):
        for pos, c in d.items():
            if c.visible and c.enabled:
                c.value = "1.0" if d is s.lateral else _sample_value("対角", s)
                n += 1
    # 検入高さ: 山数ぶん(山数が無い作業は 1 山)。帳票の「検入高さ(mm)」に出る
    yama = int(vbx.val(s.yama)) if vbx.is_numeric(s.yama) else 1
    for i in range(1, max(1, min(yama, 6)) + 1):
        s.ken_h[i].value = "250"
    return n


@dataclass
class Sample:
    pattern: Pattern
    screen: dict[str, Any] = field(default_factory=dict)
    html: str = ""
    filled: int = 0
    problems: list[str] = field(default_factory=list)


def make(p: Pattern, *, now: Optional[datetime] = None) -> Sample:
    """1 つの組み合わせの 測定画面(view)と 帳票(HTML)。問題は problems に入れる(止めない)。"""
    out = Sample(p)
    try:
        s, _h = build_state(p)
    except Exception as e:                                   # noqa: BLE001 - 見本は止めずに理由を出す
        out.problems.append(f"測定画面を作れませんでした: {type(e).__name__}: {e}")
        return out
    out.filled = fill_values(s)
    out.screen = view.build(s)
    tmp = Path(tempfile.mkdtemp(prefix="ufkensa_sample_"))
    store = LocalStore(tmp / "sample.sqlite3")
    try:
        store.add_sheet(f"WorkerTime_{LINE}").write(3, 1, "見本_検査員")
        s.v("StartTime").value = "2026/01/01 08:00:00"
        snap = wiring.to_snapshot(s, store)
        snap.t1, snap.worker_shift = "見本_検査員", "1"               # 検査員登録の代わり
        snap.start_time = snap.start_time or "2026/01/01 08:00:00"
        res = report_model.print_transfer(store, snap, report_model.PrintEnv(now=now or datetime.now()))
        if res.stopped or res.report is None:
            out.problems.append("印刷転送が止まりました: " + " / ".join(m.text for m in res.messages))
        else:
            out.html = report_html.render(res.report, title=p.label())
            if out.filled == 0:
                out.problems.append("測定欄が 1 つも見えていません")
    except Exception as e:                                   # noqa: BLE001
        out.problems.append(f"帳票を作れませんでした: {type(e).__name__}: {e}")
    finally:
        store.close()
        shutil.rmtree(tmp, ignore_errors=True)
    return out


def all_patterns() -> list[Pattern]:
    """全部の組み合わせ(画面で選べるもの。意味のない組み合わせは作らない)。"""
    out: list[Pattern] = []
    flags = list(itertools.product((False, True), repeat=3))
    for work, (name, *_rest) in WORKS.items():
        if name == "薄板":
            base = [Pattern(work=work, yama=y, at=a, inspe=i) for y in YAMA for a in AT for i in INSPE]
        elif name == "厚板":
            base = [Pattern(work=work, at=a) for a in AT]
        elif name == "少枚数":
            base = [Pattern(work=work, sheets=n) for n in (1, 3, 5)]
        else:
            base = [Pattern(work=work)]
        for b in base:
            for lat, right, free in flags:
                out.append(Pattern(b.work, b.yama, b.at, b.inspe, lat, right, free, b.sheets))
    return out


def options() -> dict[str, Any]:
    return {"works": list(WORKS), "yama": list(YAMA), "at": [{"v": "4", "label": "厚4ヵ所"}, {"v": "2", "label": "厚2ヵ所"}],
            "inspe": [{"v": "ge26", "label": "検入≧26"}, {"v": "le25", "label": "検入≦25"}],
            "thin": [w for w, t in WORKS.items() if t[0] == "薄板"], "small": [w for w, t in WORKS.items() if t[0] == "少枚数"],
            "thick": [w for w, t in WORKS.items() if t[0] == "厚板"], "total": len(all_patterns())}
