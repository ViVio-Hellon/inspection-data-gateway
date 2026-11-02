"""全ヵ所検査(zenkasho.py): 一般 / 5N / 三菱受託。歪・ラテラル・対角1(B32)・MAX/MIN(B28)・検入高さ(B31)・保存(B33)・三菱(W2)"""
from __future__ import annotations

import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from tests import kensa_fixture
from ufkensa.kensa import flow, quirks, vbx
from ufkensa.kensa.modes import sheet as sh
from ufkensa.kensa.modes import zenkasho as zk

NOW = datetime(2026, 10, 30, 10, 0)


def zen_book(choice: str = "一般", line: str = "LVC") -> sh.Book:
    b = sh.Book(now=lambda: NOW)
    ws = b.ws(sh.ZEN)
    ws.visible = True
    b.activate_sheet(sh.ZEN)
    f = b.mit
    f.visible, f.work_flag, f.line_name, f.choice = True, "P1", line, choice
    f.set_option({"一般": "一般", "5N": "5N"}.get(choice, ""))
    for (r, c), v in {(2, 4): "H3084S0", (3, 2): 36.0, (3, 4): 1528.0, (3, 6): 3053.0, (2, 13): 0.4, (2, 15): 0.4,
                      (3, 13): 3.0, (3, 15): 3.0, (4, 13): 3.0, (4, 15): 3.0, (1, 19): 2.1, (1, 13): 1.0,
                      (5, 13): 0.6, (5, 15): 0.7}.items():
        ws.set(r, c, v)
    for k, fm in sh.TOLERANCE_FORMULAS.items():
        ws.set_formula(*k, fm)
    return b


class SheetInputTest(unittest.TestCase):
    def tearDown(self) -> None:
        quirks.set_for_test(None)

    def test_width_manual_kept_and_judged(self) -> None:   # Cask:153-162 巾の手入力は残す
        b = zen_book()
        ws = b.ws(sh.ZEN)
        fx = b.user_edit(7, 11, "1525.0")
        self.assertEqual(ws.get(7, 11), 1525.0)
        self.assertEqual(ws.get(7, 12), "○")                 # 下限 1525.0 ちょうど
        self.assertIn(sh.FormShow("UFOK"), fx.forms)
        b.user_edit(7, 13, "1524.9")
        self.assertEqual(ws.get(7, 14), "×")
        self.assertEqual(ws.interior[(7, 14)], (255, 0, 0))

    def test_hizumi_modes(self) -> None:                  # Cask:1994-2039
        b = zen_book("5N")
        ws = b.ws(sh.ZEN)
        b.user_edit(7, 25, "2.0")
        self.assertEqual(ws.get(7, 26), "○")
        b.user_edit(8, 25, "-2.1")
        self.assertEqual(ws.get(8, 26), "×")
        b.user_edit(7, 27, "3.01")
        self.assertEqual(ws.get(7, 28), "×")
        b = zen_book("一般")
        ws = b.ws(sh.ZEN)
        b.user_edit(7, 25, "0.6")
        self.assertEqual(ws.get(7, 26), "○")
        b.user_edit(8, 25, "0.7")
        self.assertEqual(ws.get(8, 26), "×")
        ws.set(5, 13, None)
        b.user_edit(9, 25, "0.1")                            # M5 が空 → ×
        self.assertEqual(ws.get(9, 26), "×")
        b = zen_book("三菱")
        b.user_edit(7, 25, "9.9")
        self.assertIsNone(b.ws(sh.ZEN).get(7, 26))           # 三菱は歪判定なし

    def test_lateral(self) -> None:                       # ZenSh:63-75
        b = zen_book()
        ws = b.ws(sh.ZEN)
        b.user_edit(7, 23, "0.3")
        self.assertEqual((ws.get(7, 23), ws.get(7, 24)), (0.5, "○"))
        b.user_edit(8, 23, "-0.7")
        self.assertEqual((ws.get(8, 23), ws.get(8, 24)), (-1.0, "○"))
        b.user_edit(9, 23, "-")
        self.assertEqual((ws.get(9, 23), ws.get(9, 24)), ("-", None))
        ws.set(1, 13, None)
        b.user_edit(10, 23, "0.3")                           # M1 が空 → ×
        self.assertEqual(ws.get(10, 24), "×")

    def test_diagonal_and_b32(self) -> None:              # ZenSh:44-60, Cask:1536-1554
        b = zen_book()
        ws = b.ws(sh.ZEN)
        b.user_edit(7, 23, "1.0")
        b.user_edit(7, 19, "#L1530.4")
        self.assertIsNone(ws.get(7, 23))                     # B32: ラテラルボー値まで消える
        b.user_edit(7, 20, "#L1528.3")
        self.assertEqual((ws.get(7, 21), ws.get(7, 22)), (2.1, "○"))
        self.assertEqual(b.active_cell, (7, 23))
        quirks.set_for_test({"B32": "fix"})
        b.user_edit(8, 23, "1.0")
        b.user_edit(8, 19, "#L1000.0")
        self.assertEqual(ws.get(8, 23), 1.0)

    def test_red_x_count(self) -> None:                   # Mit:1387-1397
        b = zen_book()
        b.user_edit(7, 2, "DT10002+0000036.500M")
        b.user_edit(7, 25, "0.9")
        ws = b.ws(sh.ZEN)
        self.assertEqual(zk.check_red_x(ws, 7, 2, 76, 9), 1)
        self.assertEqual(zk.check_red_x(ws, 7, 25, 76, 28), 1)


class StringToolsTest(unittest.TestCase):
    def test_tools(self) -> None:                         # Mit:77-94, 1398-1434 / 補助:1105-1130
        self.assertEqual(zk.remove_duplicates("田中 鈴木 田中"), "田中 鈴木")
        self.assertEqual(zk.remove_duplicates("田中 中"), "田中")             # 部分一致で消える(VBA どおり)
        self.assertEqual(zk.extract_nth_text("(10.0÷2)(20.5÷3)", 2), "20.5")
        self.assertEqual(zk.extract_nth_text("(10.0)", 1), "")
        self.assertEqual(zk.extract_nth_text("(10.0÷2)", 0), "")
        self.assertEqual(zk.convert_to_half("２．０mm/3"), "20")    # 全角の "．" は半角にせず捨てる
        self.assertEqual(zk.convert_to_half("1.0/2M"), "1.0")


class DeviceTest(unittest.TestCase):
    def test_kiki(self) -> None:                          # Mit:986-1056
        b = zen_book(line="LVC")
        ws = b.ws(sh.ZEN)
        b.events_enabled = False
        sp = "　　"
        cases = {"28": "ﾏｲｸﾛ:LVCM1" + sp + "ｺﾝﾍﾞｯｸｽ:LVC C10", "28.001": "ﾏｲｸﾛ:LVCM2" + sp + "ｺﾝﾍﾞｯｸｽ:LVC C10",
                 "74": "ﾏｲｸﾛ:LVCM3" + sp + "ｺﾝﾍﾞｯｸｽ:LVC C10", "74.5": "ﾉｷﾞｽ:LVCN5" + sp + "ｺﾝﾍﾞｯｸｽ:LVC C10"}
        for atu, text in cases.items():
            zk.kiki_tensou(b, ws, 3, 29, atu, "1528", "")
            self.assertEqual(ws.get(3, 29), text)
        ws.set(3, 29, None)
        zk.kiki_tensou(b, ws, 3, 29, "0", "1528", "")
        self.assertIsNone(ws.get(3, 29))
        ws.set(14, 2, 299.5)
        ws.set(15, 4, 300.2)
        zk.kiki_tensou(b, ws, 3, 29, "300", "1528", "")
        self.assertEqual(ws.get(3, 29), "ﾉｷﾞｽ：LVCN5,N6" + sp + "ｺﾝﾍﾞｯｸｽ:LVCC11")
        zk.kiki_tensou(b, ws, 3, 29, "10", "150", "H494")
        self.assertEqual(ws.get(3, 29), "ﾏｲｸﾛ:LVCM1" + sp + "幅:ﾉｷﾞｽ:N5_丈:ｺﾝﾍﾞｯｸｽ:LVC C10")
        b.mit.line_name = "Test"
        zk.kiki_tensou(b, ws, 3, 29, "10", "1528", "")
        self.assertEqual(ws.get(3, 29), "ﾏｲｸﾛ:M1" + sp + "ｺﾝﾍﾞｯｸｽ:")


class MaxMinTest(unittest.TestCase):
    def tearDown(self) -> None:
        quirks.set_for_test(None)

    def test_b28_and_single(self) -> None:                # Mit:134-322
        b = zen_book()
        b.mit.min_text, b.mit.max_text = "1", "4"
        ws = b.ws(sh.ZEN)
        for r, v in zip(range(7, 11), (10.5, 10.6, 10.4, 10.5)):
            ws.set(r, 2, v)
        for r, v in zip(range(7, 11), (10.512, 10.7, 10.3, 10.5)):
            ws.set(r, 6, v)                                  # F 列(3 番目)にだけ最小 10.3
        b.events_enabled = False
        zk.max_min(b)
        self.assertEqual(ws.get(3, 23), vbx.single(10.7))     # Single の値が書かれる(10.69999980926514)
        # B28: 列最小 = [10.4, 0(D 空), 入らない, 入らない] → 2・3 番目も 0 → MIN 0 → 空セル(Empty=0)と一致して 0
        self.assertEqual(ws.get(4, 23), 0.0)
        self.assertIn((7, 4), ws.underline)
        quirks.set_for_test({"B28": "fix"})
        zk.max_min(b)
        self.assertEqual(ws.get(4, 23), vbx.single(10.3))


class FormTest(unittest.TestCase):
    def tearDown(self) -> None:
        quirks.set_for_test(None)

    def test_forn_select_5n(self) -> None:                # Mit:1465-1476
        b = zen_book("一般")
        b.mit.choice = "5N"
        zk.forn_select(b)
        ws = b.ws(sh.ZEN)
        self.assertEqual((ws.get(1, 22), ws.get(5, 13), ws.get(5, 15)), ("5N", 2.0, 3.0))
        self.assertTrue(b.mit.five_op and not b.mit.ippan)
        self.assertFalse(b.events_enabled)                    # V1 の書き込みで Change が走り、文字なので False のまま
        b.mit.choice = "三菱"
        zk.forn_select(b)
        self.assertTrue(b.mit.five_op)                        # 三菱では OptionButton を触らない(前の 5N が残る)
        self.assertEqual((b.mit.work, b.mit.inflag), ("ON", "flagon"))

    def test_kenhbox7_b31(self) -> None:                  # UFMit:379-388
        b = zen_book()
        f = b.mit
        f.kenh[6], f.kenh[7] = "12.34", "15"
        fx = zk.kenhbox_after_update(b, 7)
        self.assertEqual(f.kenh[6], "12")                     # B31: KenHBox6 を整形してしまう
        self.assertEqual(fx.calls[0].args[0], "KenHBox6")
        self.assertEqual(b.ws(sh.ZEN).get(3, 30), "15枚")
        quirks.set_for_test({"B31": "fix"})
        f.kenh[6], f.kenh[7] = "12.34", "15.4"
        fx = zk.kenhbox_after_update(b, 7)
        self.assertEqual((f.kenh[6], f.kenh[7]), ("12.34", "15"))

    def test_choice_change(self) -> None:
        b = zen_book()
        b.mit.choice = "５N"
        zk.choice_change(b)
        self.assertEqual(b.mit.choice, "5N")
        b.mit.choice = "その他"
        fx = zk.choice_change(b)
        self.assertEqual(fx.messages[-1].text, "正しい作業選択名を入力してください")

    def test_print_requires_height(self) -> None:          # UFMit:1670
        b = zen_book()
        res = zk.ippan_print(b)
        self.assertEqual(res.messages[-1].text, "検入高さを入力してください。")
        b.ws(sh.ZEN).set(3, 30, "10枚")
        b.ws(sh.ZEN).set(7, 3, "×")
        b.ws(sh.ZEN).interior[(7, 3)] = (255, 0, 0)
        res = zk.ippan_print(b)
        self.assertIsInstance(res, flow.Question)
        self.assertEqual(res.text, "公差外判定の板があります 厚:1箇所　" + "\r\n" + "処理を続行しますか？")


class SaveTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        kensa_fixture.use_offline(kensa_fixture.copy_to(Path(self.tmp.name)))

    def tearDown(self) -> None:
        self.tmp.cleanup()
        quirks.set_for_test(None)

    def test_fiven_save_b33(self) -> None:                # Mit:1355-1373
        b = zen_book("5N")
        b.mit.work = "ON"                                     # 新規発行(CB1)の後
        plan = zk.fiven_save(b)
        self.assertFalse(any("【梱包】５N" in s.path for s in plan.steps))
        self.assertTrue(any(s.path.endswith("全ヵ所検査_履歴\\2026.10\\20261030_H3084S0.xlsx") for s in plan.steps))
        quirks.set_for_test({"B33": "fix"})
        plan = zk.fiven_save(b)
        self.assertTrue(any("【梱包】５N\\2026.10" in s.path for s in plan.steps))

    def test_zen_save_marutoku_and_mitsubishi(self) -> None:   # Mit:597-614
        b = zen_book(line="丸徳")
        plan = zk.zen_save(b)
        self.assertEqual(plan.steps[-1].path, "\\20261030_H3084S0.xlsx")
        b = zen_book("三菱")
        b.mit.inflag = "flagon"
        with self.assertRaises(flow.Stop):
            zk.zen_save(b)
        plan = b.fx.saves[-1]
        self.assertTrue(any("三菱受託検査_履歴\\2026.10\\20261030_H3084S0.xlsx" in s.path for s in plan.steps))
        self.assertIn("パスが見つかりません", plan.error)

    def test_mitsubishi_transfer(self) -> None:           # Mit:740-855(W2: 列がずれたまま再現)
        b = zen_book("三菱")
        f = b.mit
        f.kname, f.ken_no, f.inflag, f.work = "検査員01", "H651019", "flagon", "ON"
        ws = b.ws(sh.ZEN)
        b.events_enabled = False
        for c, v in {2: 36.1, 4: 36.2, 11: 1528.1, 22: "○", 23: 0.5, 25: 0.3, 27: 0.4, 29: 0.2, 30: "良"}.items():
            ws.set(7, c, v)
        ws.set(7, 24, "○")
        ws.set(7, 26, "○")
        zk.zen_transfer(b)
        cp = b.extra["ｺﾋﾟ-用"]
        self.assertEqual((cp.get(2, 5), cp.get(2, 9), cp.get(2, 13), cp.get(2, 14)), (36.1, 1528.1, "○", 0.5))
        t = b.extra["転送済"]
        self.assertEqual(t.get(10, 6), 36.1)
        self.assertEqual(t.get(10, 14), "○")                   # N(対角線差)に判定記号
        self.assertEqual(t.get(10, 15), 0.5)                   # O(歪 幅)にラテラルボー
        self.assertEqual(t.get(10, 2), 0.2)                    # B(原板積載)にゲージ
        self.assertEqual(t.get(10, 4), 1.0)
        self.assertEqual(t.get(8, 6), " -0.4mm \n+0.4mm")
        ws.set(7, 24, None)
        with self.assertRaises(flow.Stop) as cm:
            zk.zen_transfer(b)
        self.assertEqual(cm.exception.text, "歪入力がありません,はじめからやり直してください  *[座標](7,24)")
        self.assertNotIn("転送済", b.extra)

    def test_block_no_check(self) -> None:                # UFMit:1532-1537
        b = zen_book("三菱")
        b.ws(sh.ZEN).set(7, 2, 36.0)
        res = zk.mitsubishi_print(b)
        self.assertEqual(res.messages[-1].text, "ﾌﾞﾛｯｸNOの入力をしてください")


class LotReadTest(unittest.TestCase):
    def setUp(self) -> None:
        kensa_fixture.use_offline()

    def test_combo_change(self) -> None:                  # UFMit:596-875
        b = zen_book()
        b.mit.combo = "RX021T0"
        called = []
        fx = zk.combo_change(b, diagonal=lambda a, c: 9.9, re_diagonal=lambda a, c: 1.5,
                             lateralbow_j=lambda *a: 9.9, re_lateralbow=lambda *a: (called.append(a), 2.5)[1])
        ws = b.ws(sh.ZEN)
        self.assertEqual((ws.get(3, 2), ws.get(3, 4), ws.get(3, 6)), (25.0, 1385.0, 2770.5))
        self.assertEqual((ws.get(1, 18), ws.get(1, 19)), ("追記", 2.0))     # 指示_直角度 2.0 が優先
        self.assertEqual((ws.get(1, 12), ws.get(1, 13)), (3.0, 2.5))       # 等級 3 → ReLateralbow
        self.assertEqual(called[0], (vbx.single(1385.0), vbx.single(2770.5), 25.0))
        self.assertTrue(b.events_enabled)
        self.assertIn("HistorySave", [c.name for c in fx.calls])
        b.mit.combo = "L726M50"                              # 幅 638.9 / 等級 4・指示 2.0
        zk.combo_change(b, diagonal=lambda a, c: 9.9, re_diagonal=lambda a, c: 1.5,
                        lateralbow_j=lambda *a: 0.0, re_lateralbow=lambda *a: 0.0)
        self.assertEqual(ws.get(1, 13), 2.0)                 # 指示_ﾗﾃﾗﾙﾎﾞｰ 2.0 → L1="追記"
        self.assertEqual(ws.get(1, 12), "追記")


if __name__ == "__main__":
    unittest.main()
