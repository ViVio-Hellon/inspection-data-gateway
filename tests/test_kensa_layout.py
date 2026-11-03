"""作業選択・山数・検入枚数 → 表示(layout.py)"""
from __future__ import annotations

import unittest

from tests import kensa_fixture
from tests.test_kensa_grid import FakeHooks, QuirkCase, cells
from ufkensa.kensa import grid, layout, quirks, vbx
from ufkensa.kensa.grid import KensaState

FIVE = ("S", "U", "C", "D", "E")
HIKI_FIELDS = ["ﾛｯﾄ番号", "x", "検査番号"]
LOT_FIELDS = ["検査番号", "ﾛｯﾄ番号", "ｵｰﾀﾞｰ板厚", "設計_設備ｺｰｽ", "実績_設備ｺｰｽ"]


def fake_db(ken: str = "L6X0001", atu: str = "2", course: str = "HOT L-1 KEN", c_course: str = "HOT CM1",
            in_hiki: bool = False) -> FakeHooks:
    rec_lot = [[ken, "LOT0001", atu, course, c_course]]
    return FakeHooks(records={"SIKAHIKINOW.accdb": [["LOT0001", "", ken]] if in_hiki else None,
                              "SIKALOTNOW.accdb": rec_lot},
                     fields={"SIKAHIKINOW.accdb": HIKI_FIELDS, "SIKALOTNOW.accdb": LOT_FIELDS})


def caps(s: KensaState, *names: str) -> list[str]:
    return [s.v(n).value for n in names]


class FindPriorityTest(unittest.TestCase):
    def test_find_priority(self) -> None:
        self.assertEqual(layout.find_priority("HOT FS4"), "H")
        self.assertEqual(layout.find_priority("HOT PSW ST2 QF 2HI ST2"), "T")
        self.assertEqual(layout.find_priority("HOT CM1 CM3 ANF"), "K")          # 後ろから探す
        self.assertEqual(layout.find_priority("HOT L-1 L-1 CAL L-2"), "H")
        self.assertEqual(layout.find_priority("PSW ANF"), "")
        self.assertEqual(layout.find_priority(""), "")
        self.assertEqual(layout.find_priority("hot"), "")                        # 大文字小文字を区別


class NumDigTest(unittest.TestCase):
    def test_thickness(self) -> None:
        s = KensaState(atu="74")
        s.cells["S1"].value = "21.7"
        s.cells["F4"].value = "3"
        layout.num_dig(s)
        self.assertEqual((s.cells["S1"].value, s.cells["F4"].value), ("21.700", "3.000"))
        s = KensaState(atu="74.01")                        # > 74 で 2 桁
        s.cells["S1"].value = "80.125"
        layout.num_dig(s)
        self.assertEqual(s.cells["S1"].value, "80.13")
        s = KensaState(atu="")
        s.cells["S1"].value = "1"
        layout.num_dig(s)
        self.assertEqual(s.cells["S1"].value, "1")

    def test_width_h494(self) -> None:
        for youto, hab, expect in (("H494", "250", "250.13"), ("H494", "299", "299.00"), ("H494", "0.5", "250.1"),
                                   ("H494", "300", "250.1"), ("R115", "250", "250.1")):
            s = KensaState(atu="3", youto=youto, hab=hab)
            s.cells["S5"].value = "250.125" if expect != "299.00" else "299"
            layout.num_dig(s)
            self.assertEqual(s.cells["S5"].value, expect, (youto, hab))

    def test_length_diag(self) -> None:
        s = KensaState(atu="3")
        s.cells["E12"].value = "2000.05"
        s.cells["E13"].value = "2000.05"                  # 行13・14 は整えない
        s.cells["F10"].value = "5"
        s.diag["S"].value = "1.25"
        s.diag_kou["F"].value = "0.05"
        layout.num_dig(s)
        self.assertEqual([s.cells["E12"].value, s.cells["E13"].value, s.cells["F10"].value,
                          s.diag["S"].value, s.diag_kou["F"].value], ["2000.1", "2000.05", "5.0", "1.3", "0.1"])


class FormFormulaTest(unittest.TestCase):
    def test_form_formula(self) -> None:
        s = KensaState(atu="3", ap="0.05", am="0.1", hab="1000", hp="5", hm="", tak="2000", tp="10", tm="0")
        layout.form_formula(s)
        self.assertEqual(s.apk, "")                       # ReSize.Tag が "有" でなければ何もしない
        s.resize_tag = "有"
        layout.form_formula(s)
        self.assertEqual([s.apk, s.amk, s.hpk, s.hmk, s.tpk, s.tmk], ["3.050", "2.900", "1005.0", "", "2010.0", "2000.0"])


class KenNoLookupTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        kensa_fixture.use_offline()

    def test_hiki_first(self) -> None:
        self.assertEqual(layout.lookup_ken_no(KensaState(lot="H7485T0"), grid.Hooks()), "H")

    def test_lotnow_fallback(self) -> None:
        self.assertEqual(layout.lookup_ken_no(KensaState(lot="R6550H0"), grid.Hooks()), "0")

    def test_missing(self) -> None:
        self.assertEqual(layout.lookup_ken_no(KensaState(lot="ZZZZZZZ"), grid.Hooks()), "")


class LSelectFixtureTest(unittest.TestCase):
    """受領 DB の仕掛(108 ロット)で LSelect のコース判定を確かめる。"""

    @classmethod
    def setUpClass(cls) -> None:
        kensa_fixture.use_offline()

    def test_2hi_from_design_course(self) -> None:
        # R6550H0: 検査番号 "0000000000" → "0" → FindPriority(実績 "HOT FS4") = "H"。設計に 2HI → H_2HI
        s = KensaState(lot="R6550H0", work="厚板")
        info = layout.lselect(s, "厚板", FakeHooks(), skip_db_check=True)
        self.assertEqual((info.ken_no, info.hi_flag, s.reien, s.comment), ("H", True, "H_2HI", "2HI材:厚測定強制4ヵ所"))
        self.assertAlmostEqual(info.new_atu, vbx.single(6.6))
        self.assertFalse(info.atu_flag)

    def test_l2(self) -> None:
        s = KensaState(lot="L726M50", work="薄板")
        info = layout.lselect(s, "薄板", FakeHooks(), skip_db_check=True)
        self.assertEqual((info.ken_no, info.fs4_flag, s.reien), ("H", True, "H_L-2"))
        self.assertTrue(info.atu_flag)

    def test_l3_ken_from_hiki(self) -> None:
        s = KensaState(lot="N7180C0", work="薄板")
        info = layout.lselect(s, "薄板", FakeHooks(), skip_db_check=True)
        self.assertEqual((info.ken_no, s.reien), ("K", "K_L-3"))
        self.assertTrue(info.atu_flag)                    # オーダー厚 6 は 6 以下

    def test_lot_not_found(self) -> None:
        s = KensaState(lot="ZZZZZZZ", work="厚板", atu="5")
        msgs = []
        info = layout.lselect(s, "厚板", FakeHooks(), skip_db_check=True, msgs=msgs)
        self.assertEqual(msgs[0].text, "SIKALOTNOWに該当ロットがありませんLSelect")
        self.assertEqual(info.new_atu, 5.0)               # オーダー厚が取れなければ画面の厚
        msgs = []
        layout.lselect(s, "厚板", FakeHooks(over12=True), skip_db_check=True, msgs=msgs)
        self.assertEqual(msgs, [])                        # IsOver12Hours のときは出さない


class LSelectTest(QuirkCase):
    def test_atuita(self) -> None:
        s = KensaState(lot="LOT0001", work="厚板", yama_visible=True, at_butt_visible=True)
        for i in range(11, 15):
            s.cells[f"S{i}"].visible = True
        layout.lselect(s, "厚板", fake_db(), skip_db_check=True)
        self.assertEqual(caps(s, "S", "U", "C", "D", "E", "OPS2", "OPE4"), ["S", "S1", "2", "E1", "E", "S1", "E1"])
        self.assertEqual((s.yama_visible, s.at_butt_visible), (False, False))
        self.assertFalse(s.cells["U5"].visible or s.cells["D10"].visible)
        self.assertTrue(s.cells["C5"].visible and s.cells["U4"].visible)
        self.assertFalse(s.cells["S11"].visible)
        self.assertFalse(s.diag["U"].visible)
        self.assertEqual(caps(s, "LZS", "LGS"), ["定盤", "ｹﾞｰｼﾞ"])
        self.assertEqual(caps(s, "L1", "L6"), ["幅1", "対角B"])

    def test_gauge_and_error(self) -> None:
        s = KensaState(lot="LOT0001", work="厚板")
        h = fake_db()
        h.gauge = True
        layout.lselect(s, "厚板", h, skip_db_check=True)
        self.assertEqual(caps(s, "LZS", "LGF"), ["幅ｹﾞｰｼﾞ", "丈ｹﾞｰｼﾞ"])
        self.assertEqual(s.blindfold_tag, "Set")
        # IsGaugeRequired がエラー → On Error Resume Next で Then の中へ進む
        s = KensaState(lot="LOT0001", work="厚板")
        h = fake_db()
        h.gauge = ValueError("DB なし")
        layout.lselect(s, "厚板", h, skip_db_check=True)
        self.assertEqual(caps(s, "LZS"), ["幅ｹﾞｰｼﾞ"])

    def test_is_over_12_hours_error(self) -> None:
        # FormRelat:2640 `ResultT = IsOver12Hours`: IsOver12Hours の CDate がエラー(Worker:342)でも
        # On Error Resume Next で ResultT = False のまま続く → ロットが無ければ MsgBox(2665)
        class H(type(fake_db())):
            def is_over_12_hours(self, state):
                raise vbx.TypeMismatch("CDate")
        h = H(records={"SIKAHIKINOW.accdb": None, "SIKALOTNOW.accdb": None},
              fields={"SIKAHIKINOW.accdb": HIKI_FIELDS, "SIKALOTNOW.accdb": LOT_FIELDS})
        s = KensaState(lot="LOT0001", work="厚板", load="on")
        msgs = []
        layout.lselect(s, "厚板", h, skip_db_check=True, msgs=msgs)
        self.assertEqual([m.text for m in msgs], ["SIKALOTNOWに該当ロットがありませんLSelect"])
        self.assertEqual(caps(s, "S", "U"), ["S", "S1"])

    def test_four_yama_at_count_error_resumes(self) -> None:
        # 4 山: Six_Display("on") → AtCountSelect(4) の KenAcc(1, 0) が添字エラー(FormRelat:1985、検査番号の列が無い)。
        # エラーは LSelect の On Error Resume Next で受け、続き(対角・ゲージ)は実行される(3195-3196 の次から)
        h = FakeHooks(records={"SIKAHIKINOW.accdb": [["LOT0001", ""]], "SIKALOTNOW.accdb": [["L6X0001", "LOT0001", "2",
                                                                                          "HOT", "HOT CM1"]]},
                      fields={"SIKAHIKINOW.accdb": ["ﾛｯﾄ番号", "x"], "SIKALOTNOW.accdb": LOT_FIELDS})
        s = KensaState(lot="LOT0001", work="薄板", yama="4", at_butt_caption="厚2ヵ所")
        layout.lselect(s, "薄板", h, skip_db_check=True)
        self.assertTrue(s.cells["F1"].visible)
        self.assertEqual(caps(s, "LZS"), ["定盤"])
        self.assertTrue(s.ctl("L9").visible)

    def test_thin_at_butt_reshown(self) -> None:
        # 6mm 超(AtuFlag=False)で AtButt を消しても、直後の LineSkip: で検番が T 以外なら表示に戻る
        s = KensaState(lot="LOT0001", work="薄板", yama="1")
        info = layout.lselect(s, "薄板", fake_db(atu="8"), skip_db_check=True)
        self.assertFalse(info.atu_flag)
        self.assertTrue(s.at_butt_visible)
        self.assertEqual((s.at_butt_caption, s.at_count), ("厚2ヵ所", "厚4ヵ所検査中"))
        self.assertTrue(s.inspe_butt_visible)
        self.assertEqual(caps(s, "S", "U", "C", "E"), ["S", "", "2", "E"])
        self.assertFalse(s.cells["U5"].visible)            # 3 回検査: U・D を消す
        self.assertEqual((s.v("FMU").visible, s.v("FMC").visible), (False, True))     # DistSelect(26)
        s = KensaState(lot="LOT0001", work="薄板", yama="2")
        layout.lselect(s, "薄板", fake_db(ken="T6X0001"), skip_db_check=True)
        self.assertEqual((s.at_butt_visible, s.inspe_butt_visible), (False, False))
        self.assertEqual(caps(s, "OPS5"), ["2-1-E"])

    def test_thin_yama3(self) -> None:
        s = KensaState(lot="LOT0001", work="薄板", yama="3")
        layout.lselect(s, "薄板", fake_db(), skip_db_check=True)
        self.assertEqual(caps(s, "S", "U", "C", "D", "E"), ["3-3-S", "3-3-E", "3-2-2", "3-1-S", "3-1-E"])
        self.assertTrue(s.cells["U5"].visible and s.diag["U"].visible and s.diag["U"].enabled)
        self.assertFalse(s.inspe_butt_visible)

    def test_gokuatsu(self) -> None:
        s = KensaState(lot="LOT0001", work="極厚")
        layout.lselect(s, "極厚", fake_db(), skip_db_check=True)
        self.assertEqual(caps(s, "S", "E", "L7", "L13"), ["1", "5", "丈1", "丈4"])
        self.assertTrue(all(s.cells[n].visible for n in cells(FIVE, range(1, 15))))
        self.assertFalse(s.diag["S"].visible or s.v("L9").visible)

    def test_slab(self) -> None:
        s = KensaState(lot="LOT0001", work="スラブ")
        layout.lselect(s, "スラブ", fake_db(), skip_db_check=True)
        self.assertEqual(caps(s, "S", "U", "C", "D", "E", "OPS2"), ["1-上", "1-下", "2-上", "2-下", "", "1"])
        self.assertFalse(s.cells["U1"].visible)
        self.assertTrue(s.cells["U5"].visible)
        self.assertFalse(s.cells["U9"].visible or s.cells["E5"].visible)

    def test_byousya_skips_with_statements(self) -> None:
        # 強制描写: GoTo LineSkip で With UF1 の中へ飛ぶので、`.S = ...` などは効かない(UF1. の文だけ効く)
        s = KensaState(lot="LOT0001", work="", byousya=True)
        s.v("S").value = "X"
        s.at_butt_caption = "厚4ヵ所"
        info = layout.lselect(s, "厚板", fake_db(), skip_db_check=True)
        self.assertTrue(info.jumped)
        self.assertEqual((s.v("S").value, s.at_butt_caption), ("X", "厚4ヵ所"))
        self.assertFalse(s.cells["U5"].visible)             # UF1("U" & i).Visible = False は効く

    def test_b7_cache(self) -> None:
        s = KensaState(lot="LOT0001", work="薄板", atu="5")
        s.cache.lot_no, s.cache.cours, s.cache.ord_atu, s.cache.ken_no = "LOT0001", "HOT L-1", 8.0, "L"
        info = layout.lselect(s, "薄板", fake_db(), skip_db_check=True)
        self.assertTrue(info.use_cache)
        self.assertEqual((info.new_atu, info.atu_flag), (5.0, True))      # 画面の厚
        quirks.set_for_test({"B7": "fix"})
        info = layout.lselect(s, "薄板", fake_db(), skip_db_check=True)
        self.assertEqual((info.new_atu, info.atu_flag), (8.0, False))

    def test_ll7(self) -> None:
        s = KensaState(lot="LOT0001", work="厚板")
        layout.lselect(s, "厚板", fake_db(course="HOT L-3 LL7 KEN"), skip_db_check=True)
        self.assertEqual(s.reien, "L_LL7")

    def test_allocation(self) -> None:
        rows = [["納A", "得", "R115", "5052", "H34", "3", "1000", "2000", "LOT0001", "N6X0001"],
                ["納A", "得", "R115", "5052", "H34", "3.0", "1000", "2000", "LOT0001", "L6X0002"]]
        h = fake_db()
        h.lot_info = rows
        h.delivery = "ラテラルボー/直角度/フリー反り"
        s = KensaState(lot="LOT0001", work="厚板", atu="3", hab="1000", tak="2000", sumisei="住精",
                       additional="ﾗﾃﾗﾙﾎﾞｰ検査不要")
        info = layout.lselect(s, "厚板", h)
        a = info.allocation
        self.assertEqual((a.ken_no, a.need_lateral, a.need_right, a.need_flat), ("L", True, True, True))
        self.assertEqual(len([c for c in h.calls if c[0] == "db_return_item"]), 1)   # "3" と "3.0" は同じキー
        self.assertEqual(s.additional, "")
        self.assertEqual(s.flatnes_text, "ﾌﾘｰｿﾘ要")
        self.assertTrue(s.lateral["S"].visible)
        self.assertTrue(s.diag_kou["S"].visible)


class JudgeLotAllocationsTest(unittest.TestCase):
    def test_partial_on_type_error(self) -> None:
        rows = [["納A", "", "", "", "", "3", "1000", "2000", "", "H6X1"],
                ["納B", "", "", "", "", "", "1000", "2000", "", "N6X1"]]
        h = FakeHooks(delivery="直角度")
        alloc = layout.LotAllocation()
        with self.assertRaises(vbx.TypeMismatch):
            layout.judge_lot_allocations(KensaState(), rows, h, alloc)
        self.assertEqual((alloc.ken_no, alloc.need_right), ("N6X1", True))     # 頭 1 文字にする前に止まる

    def test_blank_delivery_and_additional(self) -> None:
        rows = [["", "", "", "", "", "x", "", "", "", "S6X1"]]
        s = KensaState(additional="ﾗﾃﾗﾙﾎﾞ-追記公差あり", lateral_text="0.5")
        a = layout.judge_lot_allocations(s, rows, FakeHooks(delivery="ラテラルボー"))
        self.assertEqual((a.ken_no, a.need_lateral), ("S", False))
        self.assertEqual((s.additional, s.lateral_text), ("ﾗﾃﾗﾙﾎﾞ-追記公差あり", "0.5"))


class CountSelectTest(QuirkCase):
    def test_at_count_select_thin_l(self) -> None:
        s = KensaState(lot="LOT0001", work="薄板")
        layout.at_count_select(s, 2, fake_db())
        self.assertEqual((s.at_butt_caption, s.at_count), ("厚4ヵ所", "厚2ヵ所検査中"))
        self.assertFalse(s.v("AL3").visible)
        self.assertFalse(s.cells["S3"].visible or s.cells["C4"].enabled or s.cells["E3"].visible)
        self.assertTrue(s.cells["U3"].visible)              # 3 回検査は S・C・E だけ
        layout.at_count_select(s, 4, fake_db())
        self.assertEqual(s.at_butt_caption, "厚2ヵ所")
        self.assertTrue(s.cells["S3"].visible and s.cells["E4"].enabled)

    def test_at_count_select_t_yama4(self) -> None:
        s = KensaState(lot="LOT0001", work="薄板", yama="4")
        layout.at_count_select(s, 2, fake_db(ken="T6X0001"))
        # 検番 T の 6 回: 表示に戻すが F の Enabled は設定しない(VBA どおり)
        self.assertTrue(s.cells["F3"].visible)
        self.assertFalse(s.cells["F3"].enabled)
        self.assertTrue(s.cells["S3"].enabled)

    def test_inspection_select_b5(self) -> None:
        s = KensaState(lot="LOT0001", work="薄板", yama="2", at_butt_caption="厚2ヵ所")
        layout.inspection_select(s, "表示", fake_db())
        self.assertEqual(caps(s, "OPS5", "OPE5", "C"), ["2-2-E", "2-1-E", "2-2-E"])
        self.assertEqual(s.inspe_butt_caption, "検入≦25")
        self.assertTrue(s.cells["C1"].visible)
        quirks.set_for_test({"B5": "fix"})
        layout.inspection_select(s, "非表示", fake_db())
        self.assertEqual(caps(s, "OPS5", "C", "E"), ["2-1-E", "", "2-1-E"])
        self.assertEqual(s.inspe_butt_caption, "検入≧26")
        self.assertFalse(s.cells["C1"].visible or s.diag["C"].visible)

    def test_inspection_select_c_under_visual_not_set(self) -> None:
        s = KensaState(lot="LOT0001", work="薄板", yama="1")
        s.v("C_UnderVisual").visible = False
        layout.inspection_select(s, "表示", fake_db(ken="H6X0001"))      # 検番 H(flag2)の 5 回表示
        self.assertTrue(s.v("C_UpVisual").visible)
        self.assertFalse(s.v("C_UnderVisual").visible)     # VBA は C_UpVisual を 2 回書いている

    def test_buttons_swallow_errors(self) -> None:
        s = KensaState(lot="LOT0001", work="薄板")
        h = FakeHooks(records={"SIKAHIKINOW.accdb": [["x"]]}, fields={"SIKAHIKINOW.accdb": ["a"]})
        layout.at_butt_click(s, h)                          # 列 0 の参照エラー → On Error Resume Next
        self.assertEqual(s.at_butt_caption, "厚2ヵ所")
        with self.assertRaises(IndexError):
            layout.at_count_select(s, 2, h)


class SixDisplayTest(unittest.TestCase):
    def test_on_off(self) -> None:
        s = KensaState(lot="LOT0001", work="薄板", yama="4")
        layout.six_display(s, "on", fake_db())
        self.assertEqual(caps(s, "F", "OPS6", "FF"), ["4-1-E", "4-1-E", "4-1-E"])
        self.assertTrue(s.cells["F1"].visible and s.v("OPS6").visible and s.diag["F"].enabled)
        self.assertFalse(s.inspe_butt_visible)
        self.assertEqual(s.at_count, "厚4ヵ所検査中")         # AtButt "厚2ヵ所" → AtCountSelect(4)
        layout.six_display(s, "off", fake_db())
        self.assertFalse(s.cells["F1"].visible or s.lateral["F"].visible or s.diag["F"].visible)


class AtuOpSelectTest(unittest.TestCase):
    def run_with(self, atu: str, *, ken: str = "L6X0001", weight: float = 0.0, screen_atu: float = 3.0):
        s = KensaState(lot="LOT0001", work="厚板")
        msgs = []
        layout.atu_opselect(s, screen_atu, fake_db(ken=ken, atu=atu), weight, msgs)
        return s, msgs

    def visible(self, s: KensaState) -> list[str]:
        return [w for w in grid.WORKS if s.work_visible[w]]

    def test_ranges_and_gaps(self) -> None:
        cases = {
            "0.5": ["厚板", "薄板", "少枚数", "極厚", "スラブ"],
            "1": ["薄板"],
            "6": ["薄板"],
            "6.0005": ["厚板", "薄板", "少枚数"],          # 6〜6.001 の隙間
            "6.001": ["厚板", "少枚数"],                   # 少枚数は 1〜6 のときだけ消える
            "29.9": ["厚板", "少枚数"],
            "29.95": ["厚板", "少枚数", "極厚"],           # 29.9〜30 の隙間
            "199.95": ["厚板", "少枚数", "極厚", "スラブ"],
            "200.05": ["厚板", "少枚数", "極厚", "スラブ"],  # 200〜200.1 の隙間
            "299.5": ["厚板", "少枚数", "極厚", "スラブ"],
            "300": ["少枚数", "極厚", "スラブ"],
            "850": ["厚板", "薄板", "少枚数", "極厚", "スラブ"],
        }
        for atu, expect in cases.items():
            s, _ = self.run_with(atu)
            self.assertEqual(self.visible(s), expect, atu)
            self.assertEqual(s.work, "")
            self.assertEqual(s.work_caption["厚板"], "L")
        s, _ = self.run_with("250", weight=1500.0)
        self.assertIn("厚板", self.visible(s))
        s, _ = self.run_with("250", weight=1500.1)
        self.assertNotIn("厚板", self.visible(s))

    def test_comment_and_caption(self) -> None:
        s, _ = self.run_with("6", screen_atu=6.5)
        self.assertEqual(s.comment, "ｵｰﾀﾞｰ厚参照_6.000")
        self.assertEqual((s.ken_no_caption, s.work_caption["薄板"]), ("L6X0001", "L"))
        s, _ = self.run_with("6", screen_atu=6.0)
        self.assertEqual(s.comment, "")

    def test_ken_t(self) -> None:
        s, _ = self.run_with("10", ken="T6X0001")
        self.assertEqual(self.visible(s), ["薄板"])
        self.assertEqual(s.work_caption["薄板"], "T")

    def test_ken_zero(self) -> None:
        s, msgs = self.run_with("10", ken="0000000000")
        self.assertIn("最終圧延機から検査番号頭を生成します", msgs[0].text)
        self.assertEqual(s.work_caption["厚板"], "L")       # 実績 "HOT CM1" → L
        s = KensaState(lot="LOT0001")
        msgs = []
        h = fake_db(ken="0000000000", atu="10", c_course="HOT PSW")
        h.records["SIKALOTNOW.accdb"][0][4] = "PSW"
        layout.atu_opselect(s, 3.0, h, 0.0, msgs)
        self.assertIn("データ上にまだ検査番号がありません", msgs[0].text)

    def test_hiki_missing_exits_before_reset(self) -> None:
        # FormRelat:6233-6242: SIKAHIKINOW が MainPath・NextPath とも無い → MsgBox、Comment="" で Exit Sub。
        # SkipA のリセット(作業選択の解除・ボタン非表示・行 11〜14 の非表示)は走らない
        s = KensaState(lot="LOT0001", work="厚板", at_butt_visible=True, ken_no_caption="前")
        s.cells["S11"].visible = True
        h = fake_db(ken="L6X0001", atu="6")
        h.hiki_path_ok = False
        msgs = []
        layout.atu_opselect(s, 6.5, h, 0.0, msgs)
        self.assertEqual([m.text for m in msgs], ["日稲:SIKAHIKINOW.accdbがありません_ReadCheckAtu_OPselect"])
        self.assertEqual(msgs[0].level, "error")
        self.assertEqual(s.comment, "")                      # "ｵｰﾀﾞｰ厚参照_6.000" を消す
        self.assertEqual((s.work, s.at_butt_visible, s.ken_no_caption), ("厚板", True, "前"))
        self.assertTrue(s.cells["S11"].visible)
        self.assertTrue(all(s.work_visible[w] for w in grid.WORKS))

    def test_zero_skips_db_and_resets(self) -> None:
        s = KensaState(work="薄板", comment="x", at_butt_visible=True)
        for n in ("S11", "S5"):
            s.cells[n].visible = n == "S11"
        layout.atu_opselect(s, 0.0, FakeHooks(records={}))
        self.assertEqual((s.work, s.comment, s.at_butt_visible), ("", "", False))
        self.assertTrue(s.cells["S5"].visible)
        self.assertFalse(s.cells["S11"].visible)
        self.assertTrue(s.cells["F10"].visible)


class AtuOpSelectFixtureTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        kensa_fixture.use_offline()

    def test_empty_lot_uses_lot_filter_for_hiki(self) -> None:
        # Lot が空: SIKALOTNOW は Arrbox(…, "") で絞り込み無し → 先頭行 RX021T0(ｵｰﾀﾞｰ板厚 25, 検査番号 T6X1913)。
        # SIKAHIKINOW は GetRecordsArr(…, "[ﾛｯﾄ番号]=''")(FormRelat:6246-6247)で 0 件 → LotAcc の検査番号(6255-6258)。
        # Arrbox(…, "") で読むと SIKAHIKINOW の先頭行 N6X1699 になってしまう
        s = KensaState(lot="")
        msgs = []
        layout.atu_opselect(s, 3.0, grid.Hooks(), 0.0, msgs)
        self.assertEqual(s.ken_no_caption, "T6X1913")
        self.assertEqual(s.work_caption["薄板"], "T")
        self.assertEqual([w for w in grid.WORKS if s.work_visible[w]], ["薄板"])
        self.assertEqual(s.comment, "ｵｰﾀﾞｰ厚参照_25.000")

    def test_hiki_found(self) -> None:
        s = KensaState(lot="H7485T0")
        layout.atu_opselect(s, 3.0, grid.Hooks(), 0.0, [])
        self.assertEqual(s.ken_no_caption, "H6X1879")


class EventsTest(QuirkCase):
    def test_select_work_small_h(self) -> None:
        s = KensaState(lot="LOT0001", atu="3")
        s.work_caption["厚板"] = "H"
        h = fake_db()
        h.sheets = 3
        layout.select_work(s, "少枚数", h)
        self.assertIn(("small_number_count", 3), h.calls)
        self.assertEqual(caps(s, "S", "E"), ["1", "5"])

    def test_yama_change(self) -> None:
        s = KensaState(lot="LOT0001", work="薄板", yama="02", inspe_butt_visible=True)
        h = fake_db()
        layout.yama_change(s, h)
        self.assertEqual((s.yama, s.yama_tag), ("2", "UF1"))
        self.assertIn(("write_log", "*UFKen:山数ﾁｪﾝｼﾞ:2"), h.calls)
        msgs = []
        s.yama = "5"
        layout.yama_change(s, h, msgs)
        self.assertEqual((s.yama_tag, msgs[0].text), ("別フォーム", "別フォームで検査してください"))
        s.load = "off"
        s.yama = "7"
        layout.yama_change(s, h)
        self.assertEqual(s.yama_tag, "別フォーム")          # Load="off" のときは何もしない

    def test_yama_change_blank_u1(self) -> None:
        # `Select Case UF1.Yama.Text: Case 1, 2, 3, 4`(.Text は String)。空は数と比べて型エラー(U1 待ち)
        s = KensaState(work="薄板", yama="")
        with self.assertRaises(vbx.TypeMismatch):
            layout.yama_change(s, fake_db())

    def test_byousya(self) -> None:
        s = KensaState(lot="LOT0001", work="スラブ", atu="3", hab="1000", tak="2000")
        layout.byousya_after_update(s, fake_db())
        self.assertEqual(s.work, "スラブ")                  # スラブは外さない(VBA どおり)
        self.assertTrue(s.lateral["S"].visible and s.lateral["F"].visible and s.diag_kou["F"].visible)
        s.work = "薄板"
        layout.byousya_after_update(s, fake_db())
        self.assertEqual(s.work, "")

    def test_line_auto_selection_b6(self) -> None:
        s = KensaState(lot="LOT0001", line_option="Tott")
        h = FakeHooks(lot_info=None)
        layout.line_auto_selection(s, "", h)
        self.assertEqual((s.line_name, s.line_no, s.load, s.form_tag), ("ﾄｯﾄ", "", "off", "STOP"))
        self.assertEqual(s.v("ZEB").value, "全数.三菱.5N")
        quirks.set_for_test({"B6": "fix"})
        s = KensaState(lot="LOT0001", line_option="Maru")
        h = FakeHooks(lot_info=None)
        layout.line_auto_selection(s, "", h, maru_no=6)
        self.assertEqual((s.line_name, s.line_no, s.load, s.form_tag), ("丸徳", "場外", "on", ""))
        self.assertEqual(s.v("ZEB").value, "検入4山以上")
        self.assertEqual([c for c in h.calls if c[0] == "worker_load"], [("worker_load",)])

    def test_line_auto_selection_flat(self) -> None:
        s = KensaState(lot="LOT0001", line_option="LVC")
        h = FakeHooks(lot_info=[["納", "", "", "", "", "1", "1", "1", "", "L1"]], delivery="フリー反り")
        layout.line_auto_selection(s, "", h)
        self.assertEqual((s.flatnes_text, s.load), ("ﾌﾘｰｿﾘ要", "on"))

    def test_right_angle_and_lateral(self) -> None:
        s = KensaState(yama="3", youto="R115", sumisei="")
        layout.right_angle_display(s, "N")
        self.assertTrue(s.diag_kou["U"].visible and s.diag["U"].enabled)
        self.assertFalse(s.diag_kou["F"].visible)
        layout.right_angle_display(s, "新規", "x")
        self.assertTrue(s.diag_kou["F"].visible)
        self.assertFalse(s.diag_kou["S"].visible)
        s = KensaState(yama="1", atu="3", hab="1000", tak="2000", work="薄板")
        layout.lselect_lateral(s, "K")
        self.assertTrue(s.lateral["S"].visible)
        self.assertFalse(s.lateral["U"].visible or s.lateral["F"].visible)
        s = KensaState(atu="3", hab="", tak="2000")
        with self.assertRaises(vbx.TypeMismatch):          # Single に "" を入れて型エラー
            layout.lselect_lateral(s, "H")
        self.assertFalse(s.lateral["S"].visible)
        s = KensaState(atu="3", work="極厚")
        layout.lselect_lateral(s, "H")
        self.assertEqual(s.additional, "ﾗﾃﾗﾙﾎﾞｰ検査不要_耳付き")

    def test_clear_cache(self) -> None:
        s = KensaState()
        s.cache.lot_no, s.cache.ord_atu = "X", 3.0
        layout.clear_lselect_cache(s)
        self.assertEqual((s.cache.lot_no, s.cache.ord_atu, s.cache.skip_lot_not_found_msg), ("", 0.0, True))

    def test_op_forced(self) -> None:
        s = KensaState(lot="LOT0001")
        layout.op_forced_selection(s)
        self.assertEqual(s.blind_message, "作業選択を選んでください")


class LayoutFixedQuirksTest(QuirkCase):
    """B56〜B59・B62 の「再現」(既定)と「直す」。"""

    def test_b56_at_butt_stays_hidden(self) -> None:
        for atu, course in (("8", "HOT L-1 KEN"), ("2", "HOT 2HI KEN")):    # 6mm 超 / 2HI 通過
            for yama in ("1", "2", "3"):
                quirks.set_for_test(None)
                s = KensaState(lot="LOT0001", work="薄板", yama=yama)
                layout.lselect(s, "薄板", fake_db(atu=atu, course=course), skip_db_check=True)
                self.assertTrue(s.at_butt_visible, (atu, yama))              # 再現: LineSkip でまた表示
                quirks.set_for_test({"B56": "fix"})
                s = KensaState(lot="LOT0001", work="薄板", yama=yama)
                layout.lselect(s, "薄板", fake_db(atu=atu, course=course), skip_db_check=True)
                self.assertFalse(s.at_butt_visible, (atu, yama))
                self.assertEqual((s.at_butt_caption, s.at_count), ("厚2ヵ所", "厚4ヵ所検査中"))
        # 6mm 以下・2HI なしは表示(変わらない)
        s = KensaState(lot="LOT0001", work="薄板", yama="1")
        layout.lselect(s, "薄板", fake_db(atu="2"), skip_db_check=True)
        self.assertTrue(s.at_butt_visible)
        # 住精の 2 回目(薄板)も表示し直さない
        s = KensaState(lot="LOT0001", work="薄板", yama="1", sumisei="有")
        layout.lselect(s, "薄板", fake_db(atu="8"), skip_db_check=True)
        layout.lselect(s, "住精", fake_db(atu="8"), skip_db_check=True)
        self.assertFalse(s.at_butt_visible)

    def test_b57_byousya_line_skip(self) -> None:
        def run() -> KensaState:
            s = KensaState(lot="LOT0001", work="", byousya=True, yama="2", yama_visible=True)
            s.v("S").value = "X"
            s.at_butt_caption = "厚4ヵ所"
            layout.lselect(s, "薄板", fake_db(), skip_db_check=True)
            return s
        s = run()
        self.assertEqual((s.v("S").value, s.at_butt_caption, s.at_butt_visible), ("X", "厚4ヵ所", False))   # 再現
        quirks.set_for_test({"B57": "fix"})
        s = run()
        self.assertEqual((s.at_butt_caption, s.at_count, s.at_butt_visible), ("厚2ヵ所", "厚4ヵ所検査中", True))
        self.assertEqual(caps(s, "S", "C", "E", "OPS5"), ["2-2-S", "2-2-E", "2-1-E", "2-1-E"])
        # 薄板以外の作業では薄板の見出しにしない
        s = KensaState(lot="LOT0001", work="", byousya=True)
        s.v("S").value = "X"
        s.at_butt_caption = "厚4ヵ所"
        layout.lselect(s, "厚板", fake_db(), skip_db_check=True)
        self.assertEqual((s.v("S").value, s.at_butt_caption), ("X", "厚4ヵ所"))
        # 極厚の L9・L10 を消す `.` の文は効く
        s = KensaState(lot="LOT0001", work="", byousya=True)
        s.v("L9").visible = True
        layout.lselect(s, "極厚", fake_db(), skip_db_check=True)
        self.assertFalse(s.v("L9").visible)

    def test_b57_sumisei_db_missing(self) -> None:
        h = FakeHooks(records={}, fields={})                               # 仕掛 DB が読めない → GoTo LineSkip
        s = KensaState(lot="LOT0001", work="薄板", yama="1", yama_visible=True)
        s.at_butt_caption = "厚4ヵ所"
        info = layout.lselect(s, "住精", h, skip_db_check=True)
        self.assertTrue(info.jumped)
        self.assertEqual(s.at_butt_caption, "厚4ヵ所")                     # 再現
        quirks.set_for_test({"B57": "fix"})
        s = KensaState(lot="LOT0001", work="薄板", yama="1", yama_visible=True)
        s.at_butt_caption = "厚4ヵ所"
        layout.lselect(s, "住精", h, skip_db_check=True)
        self.assertEqual((s.at_butt_caption, s.at_butt_visible), ("厚2ヵ所", True))
        self.assertEqual(caps(s, "S", "C", "E"), ["S", "2", "E"])

    def test_b58_c_under_visual(self) -> None:
        def run() -> KensaState:
            s = KensaState(lot="LOT0001", work="薄板", yama="1")
            s.v("C_UnderVisual").visible = False
            layout.inspection_select(s, "表示", fake_db(ken="H6X0001"))
            return s
        self.assertFalse(run().v("C_UnderVisual").visible)               # 再現
        quirks.set_for_test({"B58": "fix"})
        s = run()
        self.assertTrue(s.v("C_UnderVisual").visible and s.v("C_UpVisual").visible)

    def test_b59_f_enabled(self) -> None:
        s = KensaState(lot="LOT0001", work="薄板", yama="4")
        layout.at_count_select(s, 2, fake_db(ken="T6X0001"))
        self.assertFalse(s.cells["F3"].enabled)                            # 再現
        quirks.set_for_test({"B59": "fix"})
        s = KensaState(lot="LOT0001", work="薄板", yama="4")
        layout.at_count_select(s, 2, fake_db(ken="T6X0001"))
        self.assertTrue(s.cells["F3"].visible and s.cells["F3"].enabled and s.cells["F4"].enabled)
        self.assertTrue(s.cells["S3"].enabled)

    def test_b62_slab_cleared(self) -> None:
        s = KensaState(lot="LOT0001", work="スラブ", atu="3", hab="1000", tak="2000")
        layout.byousya_after_update(s, fake_db())
        self.assertEqual(s.work, "スラブ")                                  # 再現
        quirks.set_for_test({"B62": "fix"})
        layout.byousya_after_update(s, fake_db())
        self.assertEqual(s.work, "")


if __name__ == "__main__":
    unittest.main()


class AtuOpSelectRealHooksTest(unittest.TestCase):
    """wiring.RealHooks で SIKAHIKINOW が無いとき(FormRelat:6222-6242)。"""

    def setUp(self) -> None:
        import tempfile
        from pathlib import Path
        self.tmp = Path(tempfile.mkdtemp(prefix="ufkensa_atuop_"))
        folder = kensa_fixture.copy_to(self.tmp / "db")
        kensa_fixture.use_offline(folder)
        self.folder = folder

    def tearDown(self) -> None:
        import shutil
        kensa_fixture.use_offline()
        shutil.rmtree(self.tmp, ignore_errors=True)

    def run_atu(self, now):
        from ufkensa.kensa import wiring
        from ufkensa.kensa.store import LocalStore
        env = wiring.Env(store=LocalStore(self.tmp / "local.sqlite3"), now=lambda: now, write_logs=False)
        s = KensaState(lot="H7485T0", work="厚板", at_butt_visible=True)
        layout.atu_opselect(s, 3.0, wiring.RealHooks(env), 0.0, env.messages)
        return s, [m.text for m in env.messages]

    def test_missing_both(self) -> None:
        from datetime import datetime
        (self.folder / "SIKAHIKI.sqlite3").unlink()
        s, texts = self.run_atu(datetime(2026, 11, 2, 10, 30, 0))
        # 列名の Arrbox(6221)が「…_ReadCheck_Arrbox」、行は自前の判定で「…_ReadCheckAtu_OPselect」→ Exit Sub
        self.assertEqual(texts, ["日稲:SIKAHIKINOW.accdbがありません_ReadCheck_Arrbox",
                                 "日稲:SIKAHIKINOW.accdbがありません_ReadCheckAtu_OPselect"])
        self.assertEqual((s.comment, s.work, s.at_butt_visible), ("", "厚板", True))

    def test_found_in_time_window(self) -> None:
        from datetime import datetime
        # 時刻窓(xx:59:50)でも NextPath(同じフォルダ)にあれば読む
        s, texts = self.run_atu(datetime(2026, 11, 2, 10, 59, 50))
        self.assertEqual(texts, [])
        self.assertEqual((s.ken_no_caption, s.work), ("H6X1879", ""))


class SmallNumberCountRestoreTest(unittest.TestCase):
    """Load_Text の「少枚数の枚数」(SaveLoad:2914-2927)。"""

    def run_call(self, atu: str):
        import tempfile
        from pathlib import Path
        from ufkensa.kensa import wiring
        from ufkensa.kensa.store import LocalStore

        class Hooks(wiring.RealHooks):
            def __init__(self, env):
                super().__init__(env)
                self.got = []

            def calculate_sheets(self, state, atu):
                return vbx.single(3.0)

            def small_number_count(self, state, th):
                self.got.append(th)

        env = wiring.Env(store=LocalStore(Path(tempfile.mkdtemp(prefix="ufkensa_snc_")) / "l.sqlite3"),
                         write_logs=False)
        h = Hooks(env)
        s = KensaState(atu=atu, work="少枚数")
        s.work_caption["厚板"] = "H"
        h.restore_call(s, ("SmallNumberCount判定", ""))
        return h.got

    def test_atu_blank_still_calls(self) -> None:
        # SaveLoad:2925-2927: `If Kno = "H" And flag Then Call SmallNumberCount(th)` は If UF1.ATU <> "" の外。
        # ATU が空なら th = 0 のまま呼ぶ
        self.assertEqual(self.run_call(""), [0])

    def test_atu_set(self) -> None:
        self.assertEqual(self.run_call("3"), [3])


class AtuOpSelectRestoreTest(unittest.TestCase):
    """Load_Text の `Call Atu_OPselect(val(sh.Cells(1, 2)))`(SaveLoad:2839)。引数は Double のまま渡る。"""

    def test_double_argument(self) -> None:
        import tempfile
        from pathlib import Path
        from ufkensa.kensa import wiring
        from ufkensa.kensa.store import LocalStore
        fake = fake_db(ken="L6X0001", atu="6.6")

        class Hooks(wiring.RealHooks):
            def arrbox(self, state, target, sql_filter="", use="Record"):
                return fake.arrbox(state, target, sql_filter, use)

            def atu_hiki_records(self, state, lot):
                return fake.atu_hiki_records(state, lot)

        env = wiring.Env(store=LocalStore(Path(tempfile.mkdtemp(prefix="ufkensa_atur_")) / "l.sqlite3"),
                         write_logs=False)
        s = KensaState(lot="LOT0001")
        Hooks(env).restore_call(s, ("Atu_OPselect", vbx.val("6.6")))
        # FormRelat:6214: val(ATU)=6.6 と val(NewAtu)=CDbl("6.6") は等しい → Comment は書かない。
        # Single に入れてから渡すと 6.59999990463257 になり「ｵｰﾀﾞｰ厚参照_6.600」が出てしまう
        self.assertEqual(s.comment, "")
        self.assertEqual(s.ken_no_caption, "L6X0001")
