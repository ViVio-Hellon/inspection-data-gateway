"""検査画面(A案「計測コックピット」)に渡す値を、状態モデルから作る

画面(JS)は判断しない。ここで「何を・どの色で・見せるか隠すか」まで決めた値を渡す。
色は `grid.Ctl.color` の名前("OK" / "NG" / "YELLOW" / "" …)をそのまま渡し、CSS が塗る。

グリッドの行の見出し(VBA の UF1 のラベル。G1 §3.1):
    行 1〜4   AL1〜AL4("厚1"〜"厚4")
    行 5〜10  L1〜L6(通常: 幅1・幅2・丈1・丈2・対角A・対角B / 極厚: 幅1〜幅6)
    行 11,12  L7, L8(極厚: 丈1・丈2)
    行 13,14  L12, L13(極厚: 丈3・丈4)
"""
from __future__ import annotations

from typing import Any, Optional

from . import grid
from .grid import KensaState

_ROW_LABEL = {**{r: f"AL{r}" for r in range(1, 5)}, **{r: f"L{r - 4}" for r in range(5, 11)},
              11: "L7", 12: "L8", 13: "L12", 14: "L13"}


def _ctl(c: grid.Ctl) -> dict[str, Any]:
    return {"value": c.value, "color": c.color, "border": c.border, "visible": c.visible,
            "enabled": c.enabled, "underline": c.underline}


def _row_visible(state: KensaState, r: int) -> bool:
    return any(state.cells[f"{p}{r}"].visible for p in grid.POSITIONS if r <= grid.ROW_COUNT[p])


def build(state: KensaState, extra: Optional[dict[str, Any]] = None) -> dict[str, Any]:
    v = state.view
    positions = [{"key": p, "label": v[p].value or p, "visible": v[p].visible} for p in grid.POSITIONS]
    rows = []
    for r in range(1, 15):
        label = v.get(_ROW_LABEL[r])
        rows.append({
            "no": r,
            "label": label.value if label else "",
            "visible": _row_visible(state, r) and (label.visible if label else True),
            "cells": [({"name": f"{p}{r}", **_ctl(state.cells[f"{p}{r}"])} if r <= grid.ROW_COUNT[p]
                       else {"name": "", "value": "", "color": "", "visible": False}) for p in grid.POSITIONS],
        })

    def by_pos(d: dict[str, grid.Ctl]) -> list[dict[str, Any]]:
        return [{"pos": p, **_ctl(d[p])} if p in d else {"pos": p, "visible": False} for p in grid.POSITIONS]

    out: dict[str, Any] = {
        "header": {
            "line": state.line_name, "line_no": state.line_no, "lot": state.lot,
            "ken_no": state.ken_no_caption, "reien": state.reien, "comment": state.comment,
            "youto": state.youto, "zai": state.zai, "tyu": state.tyu,
            "work": state.work, "yama": state.yama, "yama_visible": state.yama_visible,
            "at_count": state.at_count, "at_butt": {"caption": state.at_butt_caption, "visible": state.at_butt_visible},
            "inspe_butt": {"caption": state.inspe_butt_caption, "visible": state.inspe_butt_visible},
            "proofread": state.proofread_flag == "ok",
            "works": [{"key": w, "caption": state.work_caption.get(w, w), "visible": state.work_visible.get(w, True)}
                      for w in grid.WORKS],
        },
        "tolerance": {
            "atu": state.atu, "ap": state.ap, "am": state.am, "apk": state.apk, "amk": state.amk,
            "hab": state.hab, "hp": state.hp, "hm": state.hm, "hpk": state.hpk, "hmk": state.hmk,
            "tak": state.tak, "tp": state.tp, "tm": state.tm, "tpk": state.tpk, "tmk": state.tmk,
            "taik": state.taik, "lateral_grade": state.lateral_grade, "lateral_text": state.lateral_text,
            "flatnes_text": state.flatnes_text, "additional": state.additional, "sumisei": state.sumisei,
            "conductiv_text": state.conductiv_text,
        },
        "grid": {"positions": positions, "rows": rows,
                 "diag_label": v["L9"].value, "diag_kou_label": v["L10"].value, "lateral_label": v["L11"].value,
                 "diag": by_pos(state.diag), "diag_kou": by_pos(state.diag_kou), "lateral": by_pos(state.lateral),
                 "ng": by_pos(state.ng)},
        "strain": {"face": state.strain_face, "z": by_pos(state.strain_z), "g": by_pos(state.strain_g),
                   "cb": {k: _ctl(c) for k, c in state.strain_cb.items()}},
        "bh": by_pos(state.bh),
        "bad": by_pos(state.bad),
        "ken_h": [{"no": i, **_ctl(c)} for i, c in sorted(state.ken_h.items())],
        "ken_count": state.ken_count,
        "visual": [{"no": i, **_ctl(c)} for i, c in sorted(state.visual.items())],
        "next": {"name": state.inp_tag, "label": state.tagtext, "kind": state.paked, "machin": state.machin,
                 "auto": state.auto == "on"},
        "minmax": {"amin": state.amin, "amax": state.amax, "hmin": state.hmin, "hmax": state.hmax,
                   "tmin": state.tmin, "tmax": state.tmax},
        "blind": state.blind_message,
    }
    if extra:
        out.update(extra)
    return out
