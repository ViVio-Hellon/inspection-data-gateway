"""検査回数の切り替え(端から端まで)。

A) 薄板「25枚以上・以下で検査回数の切り替え」(InspeButt)
   VBA: FormRelat.InspectionSelect (FormRelat:2183-2494)。UF1.InspeButt_Click (UF1:4481-4492) で手で切り替える
   (枚数からの自動切替は無い。ボタンの Caption は「次に切り替わる先」)。
     - Caption "検入≦25" = いま 検入≧26(3 回検査: S/C/E、山 1 の C 見出し "2")。LSelect 薄板・山数変更の後は必ずこちら
       (UF1:2430 `InspectionSelect("表示")`)。
     - Caption "検入≧26" = いま 検入≦25(2 回検査: S/E。C1〜C10・DiagC を隠す。FormRelat:2228-2262)。
     - 検番フラグ 4(冷延系かつ "検入≧26")で AllCheck / PucketMove は S・E だけを見る(FormRelat:587-592)。
     - 印刷: Print:902-917 "検査回数2回" の罫線。作業中保存の復元 SaveLoad:2900-2903。
   Python: layout.inspection_select / inspe_butt_click(layout.py:359-461)、grid.ken_flags(grid.py:476)、
   input_router.all_check / pucket_move、session.toggle_inspe(session.py:710)。
   ロット N7180C0(検番 K、厚 6.06 × 巾 1530 × 丈 2505)を LVC で使う(HVC は設備仕様外)。

B) 少枚数「検査回数(枚数)での切り替え」(SmallNumberCount)
   VBA: OutTolerance.SmallNumberCount (O:3-519)。th = KenHBox7 が空なら Round(CalculateSheets,1)、
   でなければ KenCount、5 以上は 5(UF1 少枚数_Click 5244-5266 / KenCount_AfterUpdate 745-765)。検番頭 H のときだけ。
     th=5: 全 5 列(行 5〜10 は U/D を隠す) / 4: E 列なし(行 5〜10 は U も隠す) / 3: D・E なし / 2: C・D・E なし /
     1: S だけ / 0 (KenCount 空) : 出し分けない。
   Python: checks.small_number_count(checks.py:1542)、grid.small_count_th(grid.py:542)、
   session._set_value KenCount(session.py:810-823)、layout.select_work 少枚数(layout.py:1224)。
"""
from __future__ import annotations

import unittest

from tests.test_kensa_session import CARD, LOT, SessionBase
from ufkensa.kensa import terminal

THIN_LOT = "N7180C0"
THIN_WL = {"5": "1530.2", "6": "1530.4", "7": "2505.1", "8": "2505.3", "9": "2929.9", "0": "2930.1"}
PS = "SUCDE"


class CountBase(SessionBase):
    def cells(self, rows=range(1, 11)) -> dict[str, list[bool]]:
        st = self.s.state
        return {p: [st.ctl(f"{p}{i}").visible for i in rows] for p in PS}

    def filled(self) -> set[str]:
        st = self.s.state
        return {f"{p}{i}" for p in PS for i in range(1, 11) if st.ctl(f"{p}{i}").value != ""}

    def send_until_done(self, kind: str, value: str, wl: dict[str, str], dev: str = "10001") -> None:
        for _ in range(60):
            nxt = self.s.state.inp_tag
            text = ("DT" + dev + "+" + value.rjust(11, "0") + "M") if kind == "DT" else "#L" + wl.get(nxt[-1:], "0")
            r = self.s.receive(text)
            self.assertIsNone(r.stop, r.stop and r.stop.text)
            if any(m.text == kind + "検査済" for m in r.messages):
                self.ok(self.s.click(""))
                return
        self.fail(kind + " が検査済にならない")


# ----------------------------------------------------------------------
# A) 薄板 検入≦25 / ≧26
# ----------------------------------------------------------------------
class ThinInspeCountTest(CountBase):
    def setUp(self) -> None:
        super().setUp()
        terminal.set_facility("LVC")
        self.write_calibration()
        self.s = self.new_session()
        self.ok(self.s.start())
        self.ok(self.s.register_worker(CARD, {"worker_register": True}))
        self.ok(self.s.register_lot(THIN_LOT, {"lot_register": True}))
        self.ok(self.s.select_work("薄板"))
        self.ok(self.s.set_yama("1"))

    def test_default_is_26_or_more(self) -> None:
        """山数選択後は必ず 検入≧26(3 回: S/C/E)。InspeButt は表示され、Caption は次の切替先 "検入≦25"(UF1:2430)。"""
        st = self.s.state
        self.assertTrue(st.inspe_butt_visible)
        self.assertEqual(st.inspe_butt_caption, "検入≦25")
        c = self.cells()
        self.assertEqual((c["S"], c["C"], c["E"]), ([True] * 10,) * 3)
        self.assertEqual((c["U"], c["D"]), ([False] * 10,) * 2)
        self.assertEqual((st.ctl("S").value, st.ctl("C").value, st.ctl("E").value), ("S", "2", "E"))
        self.assertTrue(st.ctl("DiagC").visible)

    def test_toggle_to_25_or_less_and_back(self) -> None:
        """InspeButt で 検入≦25(2 回: S/E)→ もう一度で 検入≧26 に戻る(FormRelat:2228-2262 / 2333-2393)。"""
        st = self.s.state
        self.ok(self.s.toggle_inspe())
        self.assertEqual(st.inspe_butt_caption, "検入≧26")
        c = self.cells()
        self.assertEqual(c["C"], [False] * 10)
        self.assertEqual((c["S"], c["E"]), ([True] * 10,) * 2)
        self.assertEqual((st.ctl("C").value, st.ctl("OPS3").value), ("", ""))
        self.assertFalse(st.ctl("DiagC").visible)
        self.ok(self.s.toggle_inspe())
        self.assertEqual(st.inspe_butt_caption, "検入≦25")
        self.assertEqual(self.cells()["C"], [True] * 10)
        self.assertEqual(st.ctl("C").value, "2")
        self.assertTrue(st.ctl("DiagC").visible)

    def test_yama2_captions_both_sides(self) -> None:
        """山 2: ≧26 は C 見出し "2-2-E"、≦25 は空(FormRelat:2240-2246 / 2384-2388)。"""
        st = self.s.state
        self.ok(self.s.set_yama("2"))
        self.assertEqual(st.inspe_butt_caption, "検入≦25")
        self.assertEqual((st.ctl("S").value, st.ctl("C").value, st.ctl("E").value), ("2-2-S", "2-2-E", "2-1-E"))
        self.ok(self.s.toggle_inspe())
        self.assertEqual((st.ctl("S").value, st.ctl("C").value, st.ctl("E").value), ("2-2-S", "", "2-1-E"))

    def test_yama_change_resets_to_26(self) -> None:
        """≦25 にしてから山数を変えると ≧26 に戻る(UF1:2430 チェンジ後は必ず表示)。"""
        self.ok(self.s.toggle_inspe())
        self.ok(self.s.set_yama("2"))
        self.assertEqual(self.s.state.inspe_butt_caption, "検入≦25")
        self.assertEqual(self.cells()["C"], [True] * 10)

    def test_yama3_hides_button(self) -> None:
        """山 3 は 5 回検査で InspeButt なし(FormRelat:2934 付近 / layout._thin_yama_captions)。"""
        self.ok(self.s.set_yama("3"))
        self.assertFalse(self.s.state.inspe_butt_visible)

    def measure(self) -> set[str]:
        self.send_until_done("DT", "6.060", THIN_WL, dev="10000")      # LVC・厚 6.06 は M1 = "10000"
        self.send_until_done("#L", "", THIN_WL)
        return self.filled()

    def test_routing_26_or_more_measures_sce(self) -> None:
        """≧26: 入力の振り先は S/C/E(厚 2 ヵ所なら行 3・4 を除く)。U/D には入らない。"""
        got = self.measure()
        cols = {n[0] for n in got}
        self.assertEqual(cols, {"S", "C", "E"})
        self.assertIn("C1", got)
        self.assertIn("C5", got)

    def test_routing_25_or_less_skips_c(self) -> None:
        """≦25: 入力の振り先は S/E だけ(検番フラグ 4。FormRelat:587-592 / 739-744)。"""
        self.ok(self.s.toggle_inspe())
        got = self.measure()
        self.assertEqual({n[0] for n in got}, {"S", "E"})
        self.assertNotIn("C1", got)

    def minmax_print_check(self) -> list[str]:
        self.measure()
        self.ok(self.s.min_max())
        return [m.text for m in self.s.print_check().messages]

    def test_print_check_26_or_more(self) -> None:
        """≧26: Min/Max・印刷前チェックまで止まらない(未入力の指摘に C の測定欄は出ない)。"""
        texts = self.minmax_print_check()
        self.assertFalse(any("未入力" in t and t[:1] in "SCE" and "定盤" not in t for t in texts), texts)

    def test_print_check_25_or_less(self) -> None:
        """≦25: C 列を測らなくても印刷前チェックで C を未入力と言わない。"""
        self.ok(self.s.toggle_inspe())
        texts = self.minmax_print_check()
        self.assertFalse(any("未入力" in t and t[:1] in "SCE" and "定盤" not in t for t in texts), texts)


# ----------------------------------------------------------------------
# B) 少枚数(SmallNumberCount)
# ----------------------------------------------------------------------
class SmallCountTest(CountBase):
    EXPECT = {   # th → (行 1〜4 の列, 行 5〜10 の列)  O:3-519
        5: ("SUCDE", "SCE"),
        4: ("SUCD", "SCD"),
        3: ("SUC", "SUC"),
        2: ("SU", "SU"),
        1: ("S", "S"),
    }

    def setUp(self) -> None:
        super().setUp()
        self.ok(self.s.start())
        self.ok(self.s.register_worker(CARD, {"worker_register": True}))
        self.ok(self.s.register_lot(LOT, {"lot_register": True}))

    def small(self, count: str) -> None:
        self.ok(self.s.select_work("少枚数"))
        self.ok(self.s.yama_lb("枚数"))
        self.ok(self.s.set_value("KenHBox7", count))
        self.ok(self.s.set_value("KenCount", count))

    def visible_cols(self, row: int) -> str:
        return "".join(p for p in PS if self.s.state.ctl(f"{p}{row}").visible)

    def test_each_count(self) -> None:
        for n, (thick, size) in self.EXPECT.items():
            with self.subTest(count=n):
                self.small(str(n))
                self.assertEqual(self.visible_cols(1), thick)
                self.assertEqual(self.visible_cols(5), size)
                self.assertEqual(self.s.state.ctl("S").value, "1")

    def test_six_or_more_is_five(self) -> None:
        """5 以上は 5 回(UF1:762 `If th >= 5 Then th = 5`)。境界 5 / 6 / 30。"""
        for n in ("5", "6", "30"):
            with self.subTest(count=n):
                self.small(n)
                self.assertEqual((self.visible_cols(1), self.visible_cols(5)), self.EXPECT[5])

    def test_switch_back_up(self) -> None:
        """1 → 5 と増やすと隠した列が戻る(SmallNumberCount は最初に全部表示し直す。O:20-40)。"""
        self.small("1")
        self.assertEqual(self.visible_cols(1), "S")
        self.ok(self.s.set_value("KenCount", "5"))
        self.assertEqual(self.visible_cols(1), "SUCDE")

    def test_blank_count_shows_all(self) -> None:
        """KenHBox7 に値があり KenCount が空なら th=0 → 出し分けない(全 5 列。VBA どおり)。"""
        self.small("2")
        self.ok(self.s.set_value("KenCount", ""))
        self.assertEqual(self.visible_cols(1), "SUCDE")

    def test_count_from_height(self) -> None:
        """KenHBox7 が空なら 検入高さ ÷ 厚 の枚数で決まる(厚 15: 30mm = 2 枚、45mm = 3 枚)。少枚数_Click (UF1:5248)。"""
        for h, n in (("30", 2), ("45", 3), ("90", 5)):
            with self.subTest(height=h):
                self.ok(self.s.yama_lb("1"))
                self.ok(self.s.set_value("KenHBox1", h))
                self.ok(self.s.select_work("少枚数"))
                self.assertEqual(self.visible_cols(1), self.EXPECT[n][0])

    def test_not_small_work_ignores_count(self) -> None:
        """厚板のときは KenCount を入れても出し分けない(少枚数 かつ 検番 H のときだけ。UF1:764)。"""
        self.ok(self.s.select_work("厚板"))
        self.ok(self.s.set_value("KenCount", "2"))
        self.assertEqual(self.visible_cols(1), "SUCDE")

    def measured_cols(self, count: str) -> set[str]:
        self.small(count)
        self.send_until_done("DT", "15.012", {})
        from tests.test_kensa_session import WL
        self.send_until_done("#L", "", WL)
        return self.filled()

    def test_routing_and_print_check_2(self) -> None:
        """2 回: 入力は S/U だけに入り、Min/Max・印刷前チェックで C/D/E を未入力と言わない。"""
        got = self.measured_cols("2")
        self.assertEqual({n[0] for n in got}, {"S", "U"})
        self.ok(self.s.min_max())
        texts = [m.text for m in self.s.print_check().messages]
        self.assertFalse(any(t.startswith(("C", "D", "E")) and "未入力" in t for t in texts), texts)

    def test_routing_4(self) -> None:
        """4 回: S/U/C/D に入る。行 5〜10 は U・E を隠すので U5・E5 は空(O:430-470 / AllCheck:455-460)。"""
        got = self.measured_cols("4")
        self.assertEqual({x[0] for x in got}, {"S", "U", "C", "D"})
        self.assertFalse({"U5", "E5"} & got)

    # 不具合(VBA と同じ): 少枚数 5 回のとき、SmallNumberCount は行 5〜10 の U/D を隠す(O:259-264)のに、
    # AllCheck の Case "少枚数" th=5 は行 5〜10 も U/D の空きを見る(FormRelat:476-494)。隠れた U5〜D10 が空のまま
    # なので「#L検査済」にならず、PucketMove は振り先が無く E10 に留まる。Python も同じ
    # (ufkensa/kensa/input_router.py:102-103 `if th == 5: return scan(lambda i: five)`)。
    # 直し方: `return scan(lambda i: ("S", "C", "E") if 5 <= i <= 10 else five)`(th=4 と同じ形)。
    def test_routing_5(self) -> None:
        """5 回(境界): 全列。行 5〜10 は U・D を隠す(O:400-420)。"""
        got = self.measured_cols("5")
        self.assertEqual({x[0] for x in got}, set(PS))
        self.assertFalse({"U5", "D5"} & got)

if __name__ == "__main__":
    unittest.main()
