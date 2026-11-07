"""Min/Max と下線(minmax.py)"""
from __future__ import annotations

import unittest

from tests.test_kensa_grid import FakeHooks, QuirkCase, cells, fill
from ufkensa.kensa import grid, minmax, quirks, vbx
from ufkensa.kensa.flow import Stop
from ufkensa.kensa.grid import KensaState

FIVE = ("S", "U", "C", "D", "E")


def atuita(**kw) -> KensaState:
    """厚板・S〜E。行5〜8 は S・C・E だけ(U・D は空のまま)。"""
    base = dict(work="厚板", atu="3", hab="1000", tak="2000", ops=1, ope=5)
    base.update(kw)
    s = KensaState(**base)
    fill(s, cells(FIVE, range(1, 5)), "3.000")
    s.cells["S1"].value = "2.995"
    s.cells["U2"].value = "3.010"
    for name, v in (("S5", "1000.5"), ("C5", "999.5"), ("E5", "1000.0"), ("S6", "1001.0"), ("C6", "1000.0"),
                    ("E6", "1000.0"), ("S7", "2000.0"), ("C7", "2001.5"), ("E7", "1999.5"), ("S8", "2000.0"),
                    ("C8", "2000.0"), ("E8", "2000.0")):
        s.cells[name].value = v
    return s


def underlined(s: KensaState) -> list[str]:
    return sorted(n for n, c in s.iter_cells() if c.underline)


class MinMaxTest(QuirkCase):
    def test_atuita_full_range(self) -> None:
        s = atuita()
        s.cells["E8"].underline = True                  # 前回の下線は外す
        self.assertEqual(minmax.min_max(s, FakeHooks()), [])
        self.assertEqual([s.amin, s.amax, s.hmin, s.hmax, s.tmin, s.tmax],
                         ["2.995", "3.010", "999.5", "1001.0", "1999.5", "2001.5"])
        # U5・D5 は空のまま数と比べる(Variant の比較。U1)
        self.assertEqual(underlined(s), ["C5", "C7", "E7", "S1", "S6", "U2"])

    def test_ng_blocks(self) -> None:
        s = atuita()
        s.cells["S2"].color = grid.COLOR_NG
        with self.assertRaises(Stop) as cm:
            minmax.min_max(s, FakeHooks())
        self.assertEqual(cm.exception.text, "許容でない公差外が存在します")
        self.assertEqual(s.amin, "")
        s.ng["S"].value = "許容"
        minmax.min_max(s, FakeHooks())
        self.assertEqual(s.amin, "2.995")
        s.ng["S"].value = "保留"                         # 保留は止める
        with self.assertRaises(Stop):
            minmax.min_max(s, FakeHooks())

    def test_ng_outside_range_or_rows(self) -> None:
        s = atuita(ops=2, ope=5)
        s.cells["S2"].color = grid.COLOR_NG             # 範囲外の位置
        s.cells["E9"].value = "1"
        s.cells["E9"].color = grid.COLOR_NG             # 対角(行9)は見ない
        minmax.min_max(s, FakeHooks())
        self.assertEqual(s.amin, "3.000")

    def test_strain_lateral_rightangle_ng(self) -> None:
        s = atuita()
        s.strain_g["C"].value = "2"
        s.strain_g["C"].color = grid.COLOR_NG
        with self.assertRaises(Stop):
            minmax.min_max(s, FakeHooks())
        s = atuita()
        s.lateral["S"].value = "1.0"
        s.lateral["S"].color = grid.COLOR_NG
        minmax.min_max(s, FakeHooks())                  # Sumisei が空ならラテラルボーは見ない
        s.sumisei = "住精"
        with self.assertRaises(Stop):
            minmax.min_max(s, FakeHooks())
        s = atuita(sumisei="住精")
        s.diag_kou["E"].value = "1.0"
        s.diag_kou["E"].color = grid.COLOR_NG
        minmax.min_max(s, FakeHooks())                  # R115 でなければ直角度は見ない
        s.youto = "R115"
        with self.assertRaises(Stop):
            minmax.min_max(s, FakeHooks())

    def test_atuita_u_only_uses_s(self) -> None:
        s = atuita(ops=2, ope=2)
        minmax.min_max(s, FakeHooks())
        self.assertEqual([s.amin, s.amax, s.hmin, s.hmax], ["3.000", "3.010", "1000.5", "1001.0"])
        self.assertIn("S5", underlined(s))
        s = atuita(ops=4, ope=4)
        minmax.min_max(s, FakeHooks())
        self.assertEqual([s.hmin, s.hmax], ["999.5", "1000.0"])     # D → C

    def test_small_count_th4(self) -> None:
        s = KensaState(work="少枚数", atu="3", ops=4, ope=4)
        fill(s, cells("D", range(1, 9)), "5.0")
        s.cells["C5"].value = "1.0"
        minmax.min_max(s, FakeHooks(sheets=4))
        self.assertEqual((s.hmin, s.tmin), ("5.0", "5.0"))           # th=4 は D のまま
        minmax.min_max(s, FakeHooks(sheets=5))
        self.assertEqual(s.hmin, "1.0")                              # th=5 は D → C

    def test_h494(self) -> None:
        s = atuita(youto="H494")
        s.hab = "250"
        minmax.min_max(s, FakeHooks())
        self.assertEqual((s.hmin, s.hmax), ("999.50", "1001.00"))

    def test_empty_and_guards(self) -> None:
        s = KensaState(work="厚板", atu="3", ops=1, ope=5)
        minmax.min_max(s, FakeHooks())
        self.assertEqual([s.amin, s.amax, s.hmin, s.hmax, s.tmin, s.tmax], [""] * 6)
        s = atuita(ops=0)
        minmax.min_max(s, FakeHooks())
        self.assertEqual(s.amin, "")
        s = atuita(ops=4, ope=2)
        minmax.min_max(s, FakeHooks())
        self.assertEqual(s.amin, "")

    def test_gokuatsu_f_rows(self) -> None:
        s = KensaState(work="極厚", atu="3", ops=1, ope=6)
        with self.assertRaises(KeyError):               # 極厚は行12 まで見るので F11 で実行時エラー(VBA どおり)
            minmax.min_max(s, FakeHooks())

    def test_run_min_max_restores_load(self) -> None:
        s = atuita()
        s.cells["S1"].color = grid.COLOR_NG
        with self.assertRaises(Stop):
            minmax.run_min_max(s, FakeHooks())
        self.assertEqual(s.load, "on")

    def test_under_bar_columns(self) -> None:
        s = KensaState(work="薄板")
        fill(s, ["S1", "C1", "E1"], "2.000")
        minmax.under_bar(s, "A", 3, 5, 2.0, 9.0)
        self.assertEqual(underlined(s), ["C1", "E1"])
        self.assertTrue(grid.variant_eq("2.000", vbx.cdbl("2")))


class RangeTest(unittest.TestCase):
    def test_select_range(self) -> None:
        s = KensaState()
        minmax.select_range(s, "S", 3)
        self.assertEqual(s.v("TSC").color, minmax.COLOR_RANGE_START)
        minmax.select_range(s, "S", 1)
        self.assertEqual((s.v("TSC").color, s.v("TSS").color), (grid.COLOR_DEFAULT, minmax.COLOR_RANGE_START))
        minmax.select_range(s, "E", 5)
        self.assertEqual(s.v("TEE").color, minmax.COLOR_RANGE_END)
        minmax.select_range(s, "S", 6)                  # 開始 > 終了 → 色を付けずに SFM.Tag を戻す
        self.assertEqual((s.v("TSF").color, s.sfm_tag), (grid.COLOR_DEFAULT, ""))
        self.assertEqual(s.v("TSS").color, grid.COLOR_DEFAULT)

    def test_backcolor_check_tb_names(self) -> None:
        s = KensaState()
        for name, pos in (("ZS", "S"), ("GF", "F"), ("LateralC", "C"), ("DiagDKou", "D")):
            c = s.ctl(name)
            c.value, c.color = "1", grid.COLOR_NG
            self.assertTrue(minmax.backcolor_check_tb(s, name), name)
            s.ng[pos].value = "許容"
            self.assertFalse(minmax.backcolor_check_tb(s, name), name)

    def test_uf_array(self) -> None:
        s = KensaState()
        s.cells["S11"].value = "1"
        self.assertEqual(minmax.uf_array(s, 11, 11, 1, 6), ["1", "", "", "", "", ""])


if __name__ == "__main__":
    unittest.main()
