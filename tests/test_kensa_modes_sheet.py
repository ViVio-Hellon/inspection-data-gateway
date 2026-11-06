"""特殊検査モードの共通部分(sheet.py): パケット変換・判定の関数・セルの値・Excel のイベントの決まり

期待値は VBA を手で追ったもの(根拠の行番号はコメント)。
"""
from __future__ import annotations

import unittest
from datetime import datetime

from ufkensa.kensa import quirks, vbx
from ufkensa.kensa.modes import sheet as sh
from tests._isolation import vba_reproduce


class PacketTest(unittest.TestCase):
    def test_maker_conv(self) -> None:          # 補助:380-407, G5 §11.1
        self.assertEqual(sh.maker("DT10002+0000150.712M"), "DT")
        self.assertEqual(sh.conv("DT10002+0000150.712M"), "0000150.712")
        self.assertEqual(sh.conv("DT10002-0000000.012M"), "0000000.012")    # 符号は捨てる
        self.assertEqual(sh.conv("dt10002+0000010.512M"), "0000010.512")
        self.assertEqual(sh.maker("ＤＴ10002+0000010.512M"), "DT")         # 全角も DT(Mid の位置は元の文字列)
        self.assertEqual(sh.conv("#L395.8"), "395.8")
        self.assertEqual(sh.conv("#〆12.3"), "12.3")
        self.assertEqual(sh.maker("DT"), "")
        self.assertEqual(sh.conv("12"), "12")
        self.assertEqual(sh.maker("1528.2"), "")

    def test_to_cell(self) -> None:
        self.assertEqual(sh.to_cell("0000150.712"), 150.712)
        self.assertIsNone(sh.to_cell(""))
        self.assertEqual(sh.to_cell("-"), "-")
        self.assertEqual(sh.to_cell("' -0.4mm"), " -0.4mm")
        self.assertEqual(sh.to_cell("2026/06/02 9:26"), datetime(2026, 6, 2, 9, 26))
        self.assertEqual(sh.cstr_date(datetime(2026, 2, 24, 8, 59)), "2026/02/24 8:59:00")


@vba_reproduce
class JudgeFunctionTest(unittest.TestCase):
    def tearDown(self) -> None:
        quirks.set_for_test(None)

    def test_ken_judg_boundaries(self) -> None:  # Cask:923-931 境界は ○
        self.assertTrue(sh.ken_judg(35.6, 36.4, 35.6))
        self.assertFalse(sh.ken_judg(35.599, 36.4, 35.6))
        self.assertTrue(sh.ken_judg(36.4, 36.4, 35.6))
        self.assertFalse(sh.ken_judg(36.401, 36.4, 35.6))
        self.assertFalse(sh.ken_judg("abc", 36.4, 35.6))   # 数でなければ ×

    def test_tai_judg_single(self) -> None:      # Cask:932-977 差を Single に入れて Single の公差と比べる
        self.assertTrue(sh.tai_judg(1530.4, 1528.3, vbx.single(2.1)))   # 倍精度なら ×
        self.assertTrue(sh.tai_judg(1000.3, 1002.4, vbx.single(2.1)))
        self.assertFalse(sh.tai_judg(2000.0, 2002.1, vbx.single(2.0)))
        self.assertTrue(sh.tai_judg(1528.3, 1526.3, vbx.single(2.0)))
        self.assertFalse(sh.tai_judg(1530.4, None, 2.1))                # 空は ×
        self.assertTrue(sh.tai_judg(1000.0, 1000.0, 0.0))               # 公差 0(S1 が空)は差 0 だけ ○

    def test_lateral_hizumi(self) -> None:       # Cask:902-921
        self.assertFalse(sh.lateral_j(1.0, "0"))   # 公差 0 / 空は常に ×
        self.assertFalse(sh.lateral_j(0.0, ""))
        self.assertTrue(sh.lateral_j(1.0, "1"))
        self.assertTrue(sh.lateral_j(-1.0, "1"))
        self.assertFalse(sh.lateral_j(1.5, "1"))
        self.assertTrue(sh.hizumi_j(2.0, "2"))
        self.assertFalse(sh.hizumi_j(2.1, "2"))

    def test_unit_record_b13(self) -> None:      # Judg:2421-2444
        f = lambda x: vbx.fmt(sh.unit_record(x), "0.0")  # noqa: E731
        self.assertEqual(f(0.3), "0.5")
        self.assertEqual(f(1.0), "1.0")
        self.assertEqual(f(1.2), "1.5")
        self.assertEqual(f(-0.7), "-1.0")
        self.assertEqual(f(0.05), "0.0")
        self.assertEqual(f(5.3), "5.3")
        self.assertEqual(f(0.51), "0.0")          # B13: Single の 0.51 は 0.51 未満
        self.assertEqual(f(1000), "0.0")          # B13: 999 超も 0
        quirks.set_for_test({"B13": "fix"})
        self.assertEqual(f(0.51), "1.0")
        self.assertEqual(f(1000), "1000.0")

    def test_eq_single(self) -> None:
        # Double を Single の精度に丸めて比べる(全ヵ所作業中保存 W3=10.51200008392334 の実例)
        self.assertTrue(sh.eq_single(10.512, vbx.single(10.512)))
        self.assertTrue(sh.eq_single(None, 0.0))
        with self.assertRaises(sh.VbaError):      # 数にならない文字列の Variant と数値型 → 型エラー 13
            sh.eq_single("厚1", 0.0)

    def test_weight(self) -> None:               # OutTolerance:813-817
        self.assertEqual(sh.weight(36, 1528, 3053), 450.077)
        self.assertEqual(sh.weight_message(10, 150, 2502.5).text, "ゲージ不要 * 両面定盤歪測定必要")
        self.assertEqual(sh.weight_message(36, 1528, 3053).text, "ゲージによる測定必要")
        self.assertIsNone(sh.weight_message(0, 1528, 3053))

    def test_fmt_date(self) -> None:
        d = datetime(2026, 10, 30, 9, 5)
        self.assertEqual(sh.fmt_date(d, "yyyy/mm/dd  h:mm"), "2026/10/30  9:05")
        self.assertEqual(sh.fmt_date(None, "yyyy/mm/dd h:mm"), "")


class TemplateTest(unittest.TestCase):
    def test_templates(self) -> None:
        c = sh.template(sh.CASK_TEMPLATE)
        self.assertEqual(c.get(706, 1), 700.0)
        self.assertEqual(c.get(6, 22), "対角線差")
        z = sh.template(sh.ZEN_TEMPLATE)
        self.assertEqual(z.get(76, 1), 70.0)
        self.assertEqual(z.get(8, 34), "ﾌﾞﾛｯｸ下")
        self.assertEqual(z.get(6, 31), "元板NO")
        h = sh.template(sh.H494_TEMPLATE)
        self.assertEqual(h.get(6, 15), "幅3")
        self.assertEqual(h.get(3, 22), 0.0)
        m = sh.template(sh.MITSU_TEMPLATE)
        self.assertEqual(m.get(4, 4), "MR664-H112")

    def test_formula(self) -> None:              # UFCask:473-496 =B3-M2 など。空は 0、文字は #VALUE!
        ws = sh.Sheet("x")
        ws.set(3, 2, 36.0)
        ws.set(2, 13, 0.4)
        ws.set_formula(2, 17, "=B3-M2")
        self.assertEqual(sh.val_cell(ws.get(2, 17)), 35.6)
        ws.set_formula(3, 17, "=D3-M3")
        self.assertEqual(ws.get(3, 17), 0.0)
        ws.set(3, 4, "abc")
        self.assertIsInstance(ws.get(3, 17), sh.ExcelError)
        with self.assertRaises(sh.VbaError):
            sh.val_cell(ws.get(3, 17))


class EventRuleTest(unittest.TestCase):
    def test_undo_only_after_user_edit(self) -> None:
        b = sh.Book()
        ws = b.ws(sh.ZEN)
        b.activate_sheet("全ヵ所検査")
        b.events_enabled = False
        b.user_edit(30, 40, "x")
        b.undo()
        self.assertIsNone(ws.get(30, 40))
        b.write(ws, 30, 40, "y")
        b.undo()                                   # マクロが書いた後は戻らない
        self.assertEqual(ws.get(30, 40), "y")

    def test_numbered_path(self) -> None:        # Cask:625-641
        seen = {"C:\\a\\x.xlsx", "C:\\a\\x.002.xlsx"}
        b = sh.Book(exists=lambda p: p in seen)
        self.assertEqual(sh.numbered_path(b, "C:\\a", "x"), ("C:\\a\\x.003.xlsx", ""))
        b2 = sh.Book(exists=lambda p: True)
        self.assertEqual(sh.numbered_path(b2, "C:\\a", "x"), (None, "連番をこれ以上付けることが出来ません。"))


if __name__ == "__main__":
    unittest.main()
