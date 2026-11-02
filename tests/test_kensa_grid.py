"""検査画面の状態モデル(grid.py)"""
from __future__ import annotations

import unittest
from typing import Any, Optional

from ufkensa.kensa import grid, quirks, vbx
from ufkensa.kensa.grid import KensaState


class FakeHooks(grid.Hooks):
    """試験用の呼び出し口。DB は `records` / `fields`(target → 値)で差し替える。None なら既定(試験用 DB)。"""

    def __init__(self, *, sheets: float = 0.0, registered: bool = True, over12: bool = False,
                 lot_info: Optional[list[list[str]]] = None, delivery: Any = "", gauge: bool = False,
                 records: Optional[dict[str, Any]] = None, fields: Optional[dict[str, Any]] = None) -> None:
        self.sheets = sheets
        self.registered = registered
        self.over12 = over12
        self.lot_info = lot_info
        self.delivery = delivery
        self.gauge = gauge
        self.records = records
        self.fields = fields
        self.calls: list[tuple] = []

    def calculate_sheets(self, state, atu):
        self.calls.append(("calculate_sheets", atu))
        return vbx.single(self.sheets)

    def register_check(self, state):
        return not self.registered

    def is_over_12_hours(self, state):
        return self.over12

    def get_lot_info(self, state, lot):
        return self.lot_info

    def db_return_item(self, state, *args):
        self.calls.append(("db_return_item",) + args)
        if callable(self.delivery):
            return self.delivery(*args)
        return self.delivery

    def is_gauge_required(self, *args):
        if isinstance(self.gauge, Exception):
            raise self.gauge
        return self.gauge

    def arrbox(self, state, target, sql_filter="", use="Record"):
        if self.records is None:
            return super().arrbox(state, target, sql_filter, use)
        if use == "Fields":
            return (self.fields or {}).get(target)
        return self.records.get(target)

    #: Atu_OPselect の SIKAHIKINOW が MainPath・NextPath とも無い(False)を試すとき
    hiki_path_ok = True

    def atu_hiki_records(self, state, lot):
        if self.records is None:
            return super().atu_hiki_records(state, lot)
        return self.hiki_path_ok, self.records.get("SIKAHIKINOW.accdb")

    def loop_text(self, state, st):
        self.calls.append(("loop_text", st))

    def posting(self, state):
        self.calls.append(("posting",))

    def out_tolerance(self, state):
        self.calls.append(("out_tolerance",))

    def proofread_load(self, state):
        self.calls.append(("proofread_load",))

    def confi_reason_open(self, state, pos, num, lname, value, count):
        self.calls.append(("confi_reason_open", pos, num, value, count))

    def te_color(self, state, name):
        self.calls.append(("te_color", name))

    def worker_load(self, state):
        self.calls.append(("worker_load",))

    def write_log(self, state, text):
        self.calls.append(("write_log", text))

    def small_number_count(self, state, th):
        self.calls.append(("small_number_count", th))

    def cb_add(self, state, atu):
        self.calls.append(("cb_add", atu))


def fill(state: KensaState, names, value: str = "1.000") -> None:
    for n in names:
        state.ctl(n).value = value


def cells(cols, rows):
    return [f"{p}{i}" for p in cols for i in rows]


class QuirkCase(unittest.TestCase):
    def tearDown(self) -> None:
        quirks.set_for_test(None)


class StateTest(unittest.TestCase):
    def test_grid_shape(self) -> None:
        s = KensaState()
        self.assertEqual(len(s.cells), 5 * 14 + 10)
        self.assertIn("F10", s.cells)
        self.assertNotIn("F11", s.cells)
        with self.assertRaises(KeyError):          # VBA: UF1("F11") は実行時エラー
            s.ctl("F11")

    def test_design_time_visibility(self) -> None:
        # G6 UF1: S..E の 1〜10 は表示、11〜14 と F は非表示。直角度・ラテラルボーは非表示
        s = KensaState()
        self.assertTrue(s.ctl("E10").visible)
        self.assertFalse(s.ctl("S11").visible)
        self.assertFalse(s.ctl("F1").visible)
        self.assertFalse(s.ctl("DiagSKou").visible)
        self.assertFalse(s.ctl("LateralS").visible)
        self.assertFalse(s.ctl("DiagF").visible)
        self.assertTrue(s.ctl("DiagS").visible)
        self.assertEqual(s.at_butt_caption, "厚2ヵ所")
        self.assertEqual(s.inspe_butt_caption, "検入≦25")
        self.assertEqual(s.work_caption["極厚"], "耳付")

    def test_resolver_names(self) -> None:
        s = KensaState()
        for name, group, key in (("DiagSKou", s.diag_kou, "S"), ("DiagE", s.diag, "E"), ("LateralF", s.lateral, "F"),
                                 ("NGC", s.ng, "C"), ("ZD", s.strain_z, "D"), ("GU", s.strain_g, "U"),
                                 ("EBH", s.bh, "E"), ("FBad", s.bad, "F")):
            self.assertIs(s.ctl(name), group[key], name)
        self.assertIs(s.ctl("CBS1"), s.strain_cb["CBS1"])
        self.assertIs(s.ctl("KenHBox7"), s.ken_h[7])
        self.assertIs(s.ctl("LC12"), s.visual[12])
        with self.assertRaises(KeyError):
            s.ctl("FBH")                             # FBH は存在しない

    def test_judge_is_color(self) -> None:
        c = grid.Ctl(value="1.0", color=grid.COLOR_NG)
        self.assertEqual(c.judge, "NG")
        c.color = grid.COLOR_YELLOW                  # 黄色で塗ると判定は消える(VBA と同じ 1 つの色)
        self.assertEqual(c.judge, "")
        self.assertEqual(grid.VBA_COLOR[grid.COLOR_OK], 13434828)
        self.assertEqual(grid.VBA_COLOR[grid.COLOR_NG], 8421631)

    def test_lname(self) -> None:
        self.assertEqual(KensaState(line_name="丸徳", line_no="A22").lname, "丸徳A22")


class CompareTest(unittest.TestCase):
    def test_variant_string_vs_number(self) -> None:
        self.assertTrue(grid.variant_eq("3", 3))
        self.assertTrue(grid.variant_eq("21.712", 21.712))
        self.assertFalse(grid.variant_eq("", 3))          # 空は数より大きい(U1)
        self.assertEqual(grid.variant_cmp("", 300.0), 1)
        self.assertEqual(grid.variant_cmp("abc", 3), 1)

    def test_typed_string_raises(self) -> None:
        with self.assertRaises(vbx.TypeMismatch):
            grid.typed_cmp("", 0)

    @unittest.skip("U1 実機確認待ち: Variant の空文字と数の比較が実行時エラー13 になる場合")
    def test_variant_raises_if_u1_says_so(self) -> None:
        grid.VARIANT_STRING_GT_NUMBER = False
        try:
            with self.assertRaises(vbx.TypeMismatch):
                grid.variant_cmp("", 300.0)
        finally:
            grid.VARIANT_STRING_GT_NUMBER = True

    def test_hab_lt_300(self) -> None:
        self.assertTrue(grid.hab_lt_300(KensaState(hab="299.9")))
        self.assertFalse(grid.hab_lt_300(KensaState(hab="300")))
        self.assertFalse(grid.hab_lt_300(KensaState(hab="")))

    def test_yama_is(self) -> None:
        self.assertTrue(grid.yama_is(KensaState(yama="3"), 3))
        self.assertFalse(grid.yama_is(KensaState(yama=""), 3))


class FlagTest(unittest.TestCase):
    def test_ken_flags(self) -> None:
        s = KensaState(yama="3", inspe_butt_caption="検入≧26")
        f = grid.ken_flags(s, "L")
        self.assertEqual([f[i] for i in range(1, 7)], [True, False, True, True, False, False])
        f = grid.ken_flags(s, "H")
        self.assertEqual([f[i] for i in range(1, 7)], [False, True, True, False, False, False])
        f = grid.ken_flags(KensaState(yama="4"), "")
        self.assertEqual([f[i] for i in range(1, 7)], [False, False, False, False, True, True])

    def test_look_s_blank_definition(self) -> None:
        # Look_S_Blank(FormRelat:1063): 検番 "" は flag1、flag5 は立てない
        f = grid.ken_flags(KensaState(inspe_butt_caption="検入≧26"), "", blank_in_flag1=True)
        self.assertTrue(f[1] and f[4])
        self.assertFalse(f[5])


class HelperTest(unittest.TestCase):
    def test_fm_clear_formats_z_and_removes_underline(self) -> None:
        s = KensaState()
        s.strain_z["S"].value = "1.5"
        s.strain_z["U"].value = "耳"
        s.strain_z["F"].value = "0.125"
        s.cells["S1"].underline = True
        s.cells["F10"].underline = True
        grid.fm_clear(s, "All")
        self.assertEqual(s.strain_z["S"].value, "1.50")
        self.assertEqual(s.strain_z["U"].value, "耳")
        self.assertEqual(s.strain_z["F"].value, "0.13")      # Format は四捨五入
        self.assertFalse(s.cells["S1"].underline)
        self.assertFalse(s.cells["F10"].underline)

    def test_n_count(self) -> None:
        s = KensaState()
        self.assertFalse(grid.n_count(s))
        s.cells["E13"].value = "1"                  # 行 13・14 は見ない
        self.assertFalse(grid.n_count(s))
        s.cells["F10"].value = "1"
        self.assertTrue(grid.n_count(s))

    def test_small_count_th(self) -> None:
        h = FakeHooks(sheets=2.25)
        s = KensaState(atu="3")
        # Round(2.25, 1) = 2.2(銀行丸め。2.25 は Single でちょうど表せる)→ Format "0" → "2"
        self.assertEqual(grid.small_count_th(s, h.calculate_sheets), 2)
        h.sheets = 2.5                                # Round = 2.5 → Format(2.5, "0") = "3"(四捨五入)
        self.assertEqual(grid.small_count_th(s, h.calculate_sheets), 3)
        h.sheets = 3.45                               # Single の 3.45 は 3.4500000477 → Round = 3.5 → "4"(U5)
        self.assertEqual(grid.small_count_th(s, h.calculate_sheets), 4)
        h.sheets = 12
        self.assertEqual(grid.small_count_th(s, h.calculate_sheets), 5)   # 5 まで
        s.ken_h[7].value = "x"
        s.ken_count = "2.5"
        self.assertEqual(grid.small_count_th(s, h.calculate_sheets), 2)   # Long への代入は銀行丸め
        self.assertEqual(grid.small_count_th(KensaState(ken_count="3"), h.calculate_sheets), 0)
        self.assertEqual(grid.small_count_th(KensaState(ken_count="3"), h.calculate_sheets,
                                             start_from_ken_count=True), 3)
        # UserForm_Click は KenCount が空かどうかで決める
        s2 = KensaState(atu="3", ken_count="2")
        s2.ken_h[7].value = ""
        self.assertEqual(grid.small_count_th(s2, h.calculate_sheets, by_ken_count_blank=True), 2)


if __name__ == "__main__":
    unittest.main()
