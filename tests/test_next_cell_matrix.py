"""「次に入れる欄」の書き留め(リファクタリング 第3段・改善案 8)

入れる先(`pucket_move`)・「〜検査済」(`all_check`)・黄色の印(`look_blank`・`look_s_blank`・`smallnumber_blank`)・
振り分け(`next_blank`)を 1 つの表にまとめる前に、今の動きを書き留めた(`fixtures/next_cell_matrix.json`。作り直さない)。
まとめた後の動きは、書き留めに `fixtures/next_cell_matrix_intended.json`(決めた決まりにそろえた所。直し。理由つき)を
重ねたものと同じであること。

画面の状態は本物の作業選択(`layout.select_work`・`yama_change`・`at_butt_click`・`inspe_butt_click`・
少枚数の `SmallNumberCount`)で作る。ロットは試験用の DB(検番・オーダー厚)を差し込む。

- 作業: 厚板(厚2ヵ所/厚4ヵ所/空 が残ったもの、薄板の後)・スラブ・極厚(耳付。薄板 4 山の後)・
  薄板(検番 L/H/空/T × 山数 1〜4 × 厚2ヵ所/厚4ヵ所 × 検入≦25/≧26、山数なし)・
  少枚数(厚板.Caption=H で 枚数 1〜5 × KenCount 空/1〜5 × KenHBox7 空/あり、H でないもの)
- 並び(seq): 空から DT・#L・#〆 を「〜検査済」まで(4 通りの順)、乱数(種を固定)で途中まで入れた所から
  DT/#L/#〆・ダブルクリックで消す・欄を押す(Auto off)・フォームの余白を押す(UserForm_Click)を混ぜたもの
  1 手ごとに: 入った欄 / 黄色の欄 / 知らせ / INP.Tag / Auto / FlagLabel
- 調べ(probe): 乱数で埋めた状態ごとに、DT/#L/#〆 の入れる先(または 検査済)・next_blank の印・look_blank の印
  (look_blank は再測定 NG_Back が作業に関係なく呼ぶ)

書き留め直すとき(まとめる前のコードでだけ): UFKENSA_RECORD_NEXT_CELL=1 python -m pytest tests/test_next_cell_matrix.py
"""
from __future__ import annotations

import json
import os
import random
import unittest
from pathlib import Path
from typing import Any, Callable, Optional

from tests.test_kensa_grid import FakeHooks
from ufkensa.kensa import grid, layout, vbx, wiring
from ufkensa.kensa import input_router as ir
from ufkensa.kensa.flow import Stop
from ufkensa.kensa.grid import KensaState
from ufkensa.kensa.model import POSITIONS, ROW_COUNT, Switch

BASELINE = Path(__file__).parent / "fixtures" / "next_cell_matrix.json"
INTENDED = Path(__file__).parent / "fixtures" / "next_cell_matrix_intended.json"
HIKI_FIELDS = ["ﾛｯﾄ番号", "x", "検査番号"]
LOT_FIELDS = ["検査番号", "ﾛｯﾄ番号", "ｵｰﾀﾞｰ板厚", "設計_設備ｺｰｽ", "実績_設備ｺｰｽ"]
TEXT = {"DT": "DT10001+0000015.012M", "#L": "#L1000.5", "#〆": "#〆2.5"}
ALL_CELLS = [f"{p}{r}" for p in POSITIONS for r in range(1, ROW_COUNT[p] + 1)]


class MatrixHooks(FakeHooks):
    """試験用の DB(検番・オーダー厚 2mm)と、入った欄の記録。SmallNumberCount は本物(wiring)。"""

    def __init__(self, ken: str, sheets: float) -> None:
        kenno = (ken + "6X0001") if ken else ""
        super().__init__(sheets=sheets,
                         records={"SIKAHIKINOW.accdb": None,
                                  "SIKALOTNOW.accdb": [[kenno, "LOT0001", "2", "HOT L-1 KEN", ""]]},
                         fields={"SIKAHIKINOW.accdb": HIKI_FIELDS, "SIKALOTNOW.accdb": LOT_FIELDS})
        self.landed: list[str] = []

    def calculate_sheets(self, state, atu):
        return vbx.single(self.sheets)

    def loop_text(self, state, st):
        self.landed.append(st)

    def small_number_count(self, state, th):
        wiring.RealHooks.small_number_count(self, state, th)        # type: ignore[arg-type]


# ------------------------------------------------------------------
# 画面の状態(本物の作業選択で作る)
# ------------------------------------------------------------------
def _base(ken: str, sheets: float = 0.0) -> tuple[KensaState, MatrixHooks]:
    s = KensaState(lot="LOT0001", atu="2", hab="1000", tak="2000", proofread_flag="ok",
                   ken_no_caption=(ken + "6X0001") if ken else "")
    s.work_caption["厚板"] = ken
    s.work_caption["薄板"] = ken
    return s, MatrixHooks(ken, sheets)


def _thin(s: KensaState, h: MatrixHooks, yama: str, at: int, inspe: str) -> None:
    layout.select_work(s, "薄板", h)
    s.yama = yama
    if yama != "":
        layout.yama_change(s, h)
    # ボタンは出ているときだけ押せる(検番 T・2HI・6mm 超は 厚2/4ヵ所 が無い、山数 3・4 は 検入 が無い)
    if at == 2 and s.at_butt_visible and s.at_butt_caption == "厚2ヵ所":
        layout.at_butt_click(s, h)                     # → 厚2ヵ所検査中(ボタンは「厚4ヵ所」)
    if inspe == "le25" and s.inspe_butt_visible and s.inspe_butt_caption == "検入≦25":
        layout.inspe_butt_click(s, h)                  # → 検入≦25(ボタンは「検入≧26」)


def scenarios() -> dict[str, Callable[[], tuple[KensaState, MatrixHooks]]]:
    out: dict[str, Callable[[], tuple[KensaState, MatrixHooks]]] = {}

    def add(key: str, fn: Callable[[], tuple[KensaState, MatrixHooks]]) -> None:
        out[key] = fn

    for cap in ("厚2ヵ所", "厚4ヵ所", ""):
        def atu(cap=cap):
            s, h = _base("H")
            s.at_butt_caption = cap
            layout.select_work(s, "厚板", h)
            return s, h
        add(f"厚板|at={cap or '空'}", atu)
    def goku():
        s, h = _base("H")
        layout.select_work(s, "極厚", h)
        return s, h
    add("極厚", goku)

    # 薄板の後に作業を変える(山数・厚2ヵ所/厚4ヵ所 のボタンが残る。極厚の 4 山は B63)
    for work in ("厚板", "スラブ", "極厚", "少枚数"):
        for yama, at in (("1", 2), ("4", 4)):
            def after(work=work, yama=yama, at=at):
                s, h = _base("H", 3)
                _thin(s, h, yama, at, "ge26")
                layout.select_work(s, work, h)
                return s, h
            add(f"{work}|after薄板{yama}山at{at}", after)

    for ken in ("L", "H", "", "T"):
        for yama in ("1", "2", "3", "4"):
            for at in (4, 2):
                for inspe in ("ge26", "le25"):
                    def thin(ken=ken, yama=yama, at=at, inspe=inspe):
                        s, h = _base(ken)
                        _thin(s, h, yama, at, inspe)
                        return s, h
                    add(f"薄板|ken={ken or '空'}|yama={yama}|at={at}|{inspe}", thin)

    def thin_noyama():
        s, h = _base("L")
        _thin(s, h, "", 4, "ge26")
        return s, h
    add("薄板|ken=L|yama=空", thin_noyama)

    for sheets in (1, 2, 3, 4, 5):
        for kc in ("", "1", "2", "3", "4", "5"):
            for h7 in ("", "10"):
                def small(sheets=sheets, kc=kc, h7=h7):
                    s, h = _base("H", sheets)
                    s.ken_h[7].value = h7
                    s.ken_count = kc
                    layout.select_work(s, "少枚数", h)
                    return s, h
                add(f"少枚数H|sheets={sheets}|kc={kc or '空'}|h7={h7 or '空'}", small)
    for kc in ("", "3"):
        def small_l(kc=kc):
            s, h = _base("L", 3)
            s.ken_count = kc
            layout.select_work(s, "少枚数", h)
            return s, h
        add(f"少枚数L|kc={kc or '空'}", small_l)
    return out


# ------------------------------------------------------------------
# 操作と記録
# ------------------------------------------------------------------
def _ken(s: KensaState) -> str:
    return vbx.left(s.ken_no_caption, 1)


def _yellow(s: KensaState) -> str:
    return ",".join(n for n in ALL_CELLS if s.cells[n].color == grid.COLOR_YELLOW)


def _reset_colors(s: KensaState) -> None:
    for n in ALL_CELLS:
        if s.cells[n].color != grid.COLOR_DEFAULT:
            s.set_color(n, grid.COLOR_DEFAULT)


def _ufc(s: KensaState) -> None:
    """UserForm_Click の「次の入力欄」の所(wiring.RealHooks.user_form_click: Auto=on → 作業ごとの印)。"""
    s.auto = Switch.ON
    ir.next_blank(s, _ken(s))


def do_op(s: KensaState, h: MatrixHooks, op: str) -> str:
    """1 手。session._receive と同じく、入力先が空なら先に UserForm_Click をしてから受ける。"""
    _reset_colors(s)
    h.landed.clear()
    msgs: list[str] = []
    err = ""
    try:
        if op in TEXT:
            if s.inp_tag == "" and s.work != "" and s.load != Switch.OFF:
                _ufc(s)
                _reset_colors(s)
            msgs = [m.text for m in ir.receive(s, TEXT[op], h)]
        elif op == "ufc":
            _ufc(s)
        elif op.startswith("clk:"):
            ir.text_tag(s, op[4:], h)
        elif op.startswith("era:"):
            ir.text_tag(s, op[4:], h)
            _reset_colors(s)
            ir.cell_dbl_click(s, op[4:], h)
            s.flag_label = 1
        else:
            raise ValueError(op)
    except Stop as e:
        err = "STOP:" + e.text.replace("\r\n", " ")
        s.flag_label = 1
    except KeyError as e:
        err = "ERR:" + str(e)
        s.flag_label = 1
        s.inp = ""
    landed = h.landed[0] if h.landed else "-"
    return f"{op}>{landed}|{_yellow(s)}|{';'.join(msgs)}{err}|{s.inp_tag}|{s.auto}{s.flag_label}"


def _done(rec: str, kind: str) -> bool:
    return (kind + "検査済") in rec or "STOP:" in rec or "ERR:" in rec


def run_until(s: KensaState, h: MatrixHooks, kinds: list[str], cap: int = 70) -> list[str]:
    """kinds の順に、それぞれ「〜検査済」まで送る(cap 手で打ち切り)。kinds が複数で 'alt' なら交互。"""
    out: list[str] = []
    for kind in kinds:
        for _ in range(cap):
            rec = do_op(s, h, kind)
            out.append(rec)
            if _done(rec, kind):
                break
    return out


def run_alt(s: KensaState, h: MatrixHooks, cap: int = 140) -> list[str]:
    out: list[str] = []
    left = ["DT", "#L"]
    i = 0
    while left and len(out) < cap:
        kind = left[i % len(left)]
        rec = do_op(s, h, kind)
        out.append(rec)
        if _done(rec, kind):
            left.remove(kind)
        else:
            i += 1
    out += run_until(s, h, ["#〆"])
    return out


def _r(rng: random.Random, n: int) -> int:
    return int(rng.random() * n)


def prefill(s: KensaState, rng: random.Random) -> str:
    """乱数で欄を埋める(表示の欄は p、隠れた欄は 0.1)。埋めた欄を返す。"""
    p = (0.2, 0.5, 0.8, 0.95)[_r(rng, 4)]
    filled = []
    for n in ALL_CELLS:
        c = s.cells[n]
        if rng.random() < (p if c.visible else 0.1):
            c.value = "1.5"
            filled.append(n)
    return ",".join(filled)


def random_ops(s: KensaState, h: MatrixHooks, rng: random.Random, n: int = 40) -> list[str]:
    out: list[str] = []
    vis = [c for c in ALL_CELLS if s.cells[c].visible]
    for _ in range(n):
        x = rng.random()
        if x < 0.3:
            op = "DT"
        elif x < 0.6:
            op = "#L"
        elif x < 0.7:
            op = "#〆"
        elif x < 0.82:
            filled = [c for c in ALL_CELLS if s.cells[c].value != ""]
            op = ("era:" + filled[_r(rng, len(filled))]) if filled else "ufc"
        elif x < 0.92:
            op = "clk:" + vis[_r(rng, len(vis))]
        else:
            op = "ufc"
        out.append(do_op(s, h, op))
    return out


_SCALARS = ("inp_tag", "tagtext", "auto", "inp", "flag_label")


def probe(s: KensaState, h: MatrixHooks) -> dict[str, str]:
    """この状態で、DT/#L/#〆 の入れる先・next_blank の印・look_blank の印(状態は元に戻す)。"""
    saved = {k: getattr(s, k) for k in _SCALARS}
    colors = {n: (s.cells[n].color, s.cells[n].border) for n in ALL_CELLS}
    verdicts = dict(s.verdicts)

    def restore() -> None:
        for k, v in saved.items():
            setattr(s, k, v)
        for n, (col, bor) in colors.items():
            s.cells[n].color, s.cells[n].border = col, bor
        s.verdicts.clear()
        s.verdicts.update(verdicts)

    out: dict[str, str] = {}
    for kind in ("DT", "#L", "#〆"):
        s.auto, s.inp_tag = Switch.ON, "?"
        msgs = []
        try:
            ir.pucket_move(s, kind, _ken(s), s.youto, h, msgs)
            out[kind] = "済" if any("検査済" in m.text for m in msgs) else s.inp_tag
        except KeyError as e:
            out[kind] = "ERR:" + str(e)
        restore()
    for key, fn in (("mark", lambda: ir.next_blank(s, _ken(s))),
                    ("look", lambda: ir.look_blank(s, 1, 14 if s.is_work("極厚") else 10))):
        _reset_colors(s)
        s.auto, s.inp_tag = Switch.ON, "?"
        try:
            fn()
            out[key] = s.inp_tag + "/" + _yellow(s)
        except KeyError as e:
            out[key] = "ERR:" + str(e)
        restore()
    return out


def record() -> dict[str, Any]:
    data: dict[str, Any] = {}
    for idx, (key, make) in enumerate(scenarios().items()):
        entry: dict[str, Any] = {}
        s, h = make()
        entry["visible"] = ",".join(n for n in ALL_CELLS if s.cells[n].visible and s.cells[n].enabled)
        entry["state"] = {"at_butt": s.at_butt_caption, "inspe": s.inspe_butt_caption, "yama": s.yama,
                          "ken": _ken(s), "ken_count": s.ken_count}
        seqs: dict[str, list[str]] = {}
        for name, kinds in (("DT>#L>#〆", ["DT", "#L", "#〆"]), ("#L>DT>#〆", ["#L", "DT", "#〆"]),
                            ("#〆", ["#〆"])):
            s, h = make()
            seqs[name] = run_until(s, h, kinds)
        s, h = make()
        seqs["DT・#L交互>#〆"] = run_alt(s, h)
        for r in range(3):
            rng = random.Random(f"{idx}-{r}")
            s, h = make()
            pre = prefill(s, rng)
            seqs[f"rand{r}"] = ["pre:" + pre] + random_ops(s, h, rng)
        entry["seq"] = seqs
        probes: list[list[Any]] = []
        rng = random.Random(f"probe-{idx}")
        for _ in range(30):
            s, h = make()
            pre = prefill(s, rng)
            probes.append([pre, probe(s, h)])
        entry["probe"] = probes
        data[key] = entry
    return data


def expected() -> dict[str, Any]:
    base = json.loads(BASELINE.read_text(encoding="utf-8"))
    if INTENDED.exists():
        for path, item in json.loads(INTENDED.read_text(encoding="utf-8")).items():
            if path.startswith("_"):
                continue
            node = base
            keys = path.split("/")
            for k in keys[:-1]:
                node = node[int(k)] if isinstance(node, list) else node[k]
            last = keys[-1]
            if isinstance(node, list):
                node[int(last)] = item["value"]
            else:
                node[last] = item["value"]
    return base


class NextCellMatrixTest(unittest.TestCase):
    maxDiff = None

    def test_matrix(self) -> None:
        got = record()
        if os.environ.get("UFKENSA_RECORD_NEXT_CELL") == "1":
            BASELINE.write_text(json.dumps(got, ensure_ascii=False, indent=0, sort_keys=True) + "\n", encoding="utf-8")
            self.skipTest("書き留めた")
        want = expected()
        self.assertEqual(sorted(got), sorted(want))
        diffs = []
        for key in want:
            if got[key] != want[key]:
                g, w = got[key], want[key]
                for part in ("visible", "state"):
                    if g[part] != w[part]:
                        diffs.append(f"{key} {part}: {g[part]} != {w[part]}")
                for name in w["seq"]:
                    for i, (a, b) in enumerate(zip(g["seq"][name], w["seq"][name])):
                        if a != b:
                            diffs.append(f"{key} seq {name}[{i}]: {a} != {b}")
                            break
                    if len(g["seq"][name]) != len(w["seq"][name]):
                        diffs.append(f"{key} seq {name}: len {len(g['seq'][name])} != {len(w['seq'][name])}")
                for i, (a, b) in enumerate(zip(g["probe"], w["probe"])):
                    if a != b:
                        diffs.append(f"{key} probe[{i}]: {a[1]} != {b[1]}")
        self.assertEqual(diffs[:40], [], f"{len(diffs)} 件の違い")


if __name__ == "__main__":
    unittest.main()
