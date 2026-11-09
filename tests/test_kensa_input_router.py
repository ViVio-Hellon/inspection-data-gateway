"""伝送文字列の解析と次の入力欄(input_router.py)"""
from __future__ import annotations

import unittest

from tests.test_kensa_grid import FakeHooks, QuirkCase, cells, fill
from ufkensa.kensa import grid, input_router as ir, quirks
from ufkensa.kensa.flow import Stop
from ufkensa.kensa.grid import KensaState

FIVE = ("S", "U", "C", "D", "E")


def ready(work: str = "厚板", **kw) -> KensaState:
    """検査員登録済み・校正済み・作業選択済みの状態。"""
    s = KensaState(work=work, proofread_flag="ok", atu="3", hab="1000", tak="2000", **kw)
    return s


class MakerConvTest(unittest.TestCase):
    def test_maker(self) -> None:
        self.assertEqual(ir.maker("DT10002+0000021.712M"), "DT")
        self.assertEqual(ir.maker("dt10002+0000021.712M"), "DT")       # UCase
        self.assertEqual(ir.maker("ＤＴ10002"), "DT")                    # vbNarrow
        self.assertEqual(ir.maker("#L395.8"), "#L")
        self.assertEqual(ir.maker("#l3"), "#L")
        self.assertEqual(ir.maker("＃〆5"), "#〆")
        self.assertEqual(ir.maker("#〆12.3"), "#〆")
        self.assertEqual(ir.maker("DT"), "")                             # 3 文字未満
        self.assertEqual(ir.maker(""), "")
        self.assertEqual(ir.maker("AB123"), "")

    def test_conv(self) -> None:
        # 補助:383 DT は Mid(Trm, 9, 11)
        self.assertEqual(ir.conv("DT10002+0000021.712M"), "0000021.712")
        self.assertEqual(ir.conv("#L395.8"), "395.8")
        self.assertEqual(ir.conv("#〆5"), "5")
        self.assertEqual(ir.conv("XYZ"), "XYZ")


class PucketMoveTest(unittest.TestCase):
    def test_atuita_order(self) -> None:
        s = ready()
        h = FakeHooks()
        ir.pucket_move(s, "DT", "", "", h)
        self.assertEqual((s.inp_tag, s.tagtext), ("S1", "S1"))
        fill(s, cells("S", range(1, 5)))
        ir.pucket_move(s, "DT", "", "", h)
        self.assertEqual(s.inp_tag, "U1")               # 厚板の行1〜4 は S→U→C→D→E
        fill(s, cells("UCDE", range(1, 5)))
        msgs = []
        ir.pucket_move(s, "DT", "", "", h, msgs)
        self.assertEqual(msgs[0].text, "DT検査済")        # FormRelat:792
        self.assertEqual(msgs[0].level, "info")          # 知らせ(警告の色にしない)
        self.assertEqual((s.auto, s.inp, s.inp_tag, s.flag_label), ("off", "", "", 1))

    def test_atuita_width_goes_to_c(self) -> None:
        # 厚板 #L: S5〜S10 の後は U ではなく C5〜(FormRelat:869-879)
        s = ready()
        fill(s, cells("S", range(5, 11)))
        ir.pucket_move(s, "#L", "", "", FakeHooks())
        self.assertEqual(s.inp_tag, "C5")
        fill(s, cells("C", range(5, 11)))
        ir.pucket_move(s, "#L", "", "", FakeHooks())
        self.assertEqual(s.inp_tag, "E5")               # D の代わりに E

    def test_dt_rows_by_at_butt(self) -> None:
        s = ready("少枚数", at_butt_caption="厚4ヵ所")   # いま 2 ヵ所検査 → 行1〜2(厚板は B93 で 1〜4 行)
        fill(s, ["S1", "S2"])
        ir.pucket_move(s, "DT", "", "", FakeHooks())
        self.assertEqual(s.inp_tag, "U1")
        s = ready("少枚数", at_butt_caption="")                   # どちらでもない → a=n=0 → 「DT検査済」(G1 §11-7)
        msgs = []
        ir.pucket_move(s, "DT", "", "", FakeHooks(), msgs)
        self.assertEqual(msgs[0].text, "DT検査済")


    def test_hab_blank_is_not_error(self) -> None:
        s = ready()
        s.hab = ""
        ir.pucket_move(s, "DT", "", "", FakeHooks())     # U1(Variant の "" < 300 は偽)
        self.assertEqual(s.inp_tag, "S1")

    def test_thin_flags(self) -> None:
        # 薄板 冷延(L)= 3 回: S → C → E
        s = ready("薄板")
        fill(s, cells("S", range(1, 5)))
        ir.pucket_move(s, "DT", "L", "", FakeHooks())
        self.assertEqual(s.inp_tag, "C1")
        # 25 枚以下(flag4)= 2 回: S → E
        s = ready("薄板", inspe_butt_caption="検入≧26")
        fill(s, cells("S", range(1, 5)))
        ir.pucket_move(s, "DT", "L", "", FakeHooks())
        self.assertEqual(s.inp_tag, "E1")
        # 山数 3 = 5 回: S → U(検番 H だけでは AllCheck の分岐に当たらない → test_thin_h_only_counts_as_done)
        s = ready("薄板", yama="3")
        fill(s, cells("S", range(1, 5)))
        ir.pucket_move(s, "DT", "H", "", FakeHooks())
        self.assertEqual(s.inp_tag, "U1")
        # 山数 4 = 6 回: F まで
        s = ready("薄板", yama="4")
        fill(s, cells(FIVE, range(1, 5)))
        s.cells["F1"].visible = True
        ir.pucket_move(s, "DT", "L", "", FakeHooks())
        self.assertEqual(s.inp_tag, "F1")

    def test_thin_h_only_counts_as_done(self) -> None:
        # 薄板・検番 H・山数 1 は AllCheck の分岐に当たらず「検査済」(G1 §11-8)
        s = ready("薄板", yama="1")
        msgs = []
        ir.pucket_move(s, "DT", "H", "", FakeHooks(), msgs)
        self.assertEqual(msgs[0].text, "DT検査済")

    def test_small_count_th_columns(self) -> None:
        # 厚板.Caption="H" で少枚数 → AllCheck は th で列を決める(th=2 → S,U)
        s = ready("少枚数")
        s.work_caption["厚板"] = "H"
        fill(s, cells("SU", range(1, 5)))
        msgs = []
        ir.pucket_move(s, "DT", "", "", FakeHooks(sheets=2), msgs)
        self.assertEqual(msgs[0].text, "DT検査済")

    def test_unknown_packet_keeps_target(self) -> None:
        s = ready(inp_tag="S3")
        msgs = []
        ir.pucket_move(s, "", "", "", FakeHooks(), msgs)
        self.assertEqual(msgs[0].text, "ﾊﾟｹｯﾄ無し:Exit Sub_PucketMove")
        self.assertEqual(s.inp_tag, "S3")

    def test_no_work(self) -> None:
        s = KensaState(inp_tag="S2")
        ir.pucket_move(s, "DT", "", "", FakeHooks())
        self.assertEqual(s.inp_tag, "S2")

    def test_slab(self) -> None:
        # 2026-11-08: スラブの厚は S1〜S4 の次に C1(B90。VBA FormRelat:831-847 は U5 へ入れていた)
        s = ready("スラブ")
        fill(s, cells("S", range(1, 5)))
        ir.pucket_move(s, "DT", "", "", FakeHooks())
        self.assertEqual(s.inp_tag, "C1")
        # #L は S5〜S10 → U5〜U8 → C5〜C10 → D5〜D8(隠れた U9・U10 は飛ばす)
        fill(s, cells("S", range(5, 11)) + cells("U", range(5, 9)))
        ir.pucket_move(s, "#L", "", "", FakeHooks())
        self.assertEqual(s.inp_tag, "C5")


class LookBlankTest(unittest.TestCase):
    def test_look_blank_atuita_c_from_row1(self) -> None:
        # Look_Blank の厚板は行5〜10 のとき C1〜C10 から探す(FormRelat:1369)
        s = ready()
        fill(s, cells("S", range(1, 11)) + cells("U", range(1, 5)))
        ir.look_blank(s, 1, 10)
        self.assertEqual(s.inp_tag, "C1")
        self.assertEqual((s.cells["C1"].color, s.cells["C1"].border), (grid.COLOR_YELLOW, grid.BORDER_RED))

    def test_look_blank_auto_off(self) -> None:
        s = ready(auto="off", inp_tag="E3")
        ir.look_blank(s, 1, 10)
        self.assertEqual(s.inp_tag, "E3")

    def test_look_s_blank_skips_rows_3_4(self) -> None:
        s = ready("薄板", at_butt_caption="厚4ヵ所")
        fill(s, ["S1", "S2"])
        ir.look_s_blank(s, 1, 10, "L")
        self.assertEqual(s.inp_tag, "S5")               # 3 → 5 に飛ぶ
        s = ready("薄板", at_butt_caption="厚4ヵ所")
        fill(s, [f"S{i}" for i in (1, 2, 3, 5, 6, 7, 8, 9, 10)])
        ir.look_s_blank(s, 1, 10, "L")
        self.assertEqual(s.inp_tag, "C1")               # S4 は見ない
        self.assertEqual(s.tagtext, "C1")

    def test_look_s_blank_hidden_cells(self) -> None:
        s = ready("薄板", at_butt_caption="厚4ヵ所")
        fill(s, cells("S", range(1, 11)))
        s.cells["C1"].visible = False
        ir.look_s_blank(s, 1, 10, "L")
        self.assertEqual(s.inp_tag, "C2")               # 厚4ヵ所の分岐は Visible を見る

    def test_look_s_blank_blank_ken_is_three_times(self) -> None:
        # 検番 "" は flag1(3 回扱い。AllCheck の 5 回扱いと違う)
        s = ready("薄板")
        fill(s, cells("S", range(1, 11)))
        ir.look_s_blank(s, 1, 10, "")
        self.assertEqual(s.inp_tag, "C1")

    def test_look_s_blank_mpex_tag_white(self) -> None:
        s = ready("薄板", mpex_tag="Click")
        ir.look_s_blank(s, 1, 10, "L")
        self.assertEqual((s.cells["S1"].color, s.cells["S1"].border), (grid.COLOR_WHITE, grid.BORDER_BLACK))

    def test_smallnumber_blank(self) -> None:
        s = ready("少枚数")
        fill(s, cells("S", range(1, 11)))
        s.cells["U1"].visible = False
        ir.smallnumber_blank(s, 1, 10, 3)
        self.assertEqual(s.inp_tag, "U2")
        s2 = ready("少枚数", inp_tag="X")
        ir.smallnumber_blank(s2, 1, 10, 6)              # 1〜5 以外は何もしない
        self.assertEqual(s2.inp_tag, "X")

    def test_next_blank_small_count(self) -> None:
        s = ready("少枚数", ken_count="2.5")             # Long へ: 2.5 → 2
        ir.next_blank(s, "")
        self.assertEqual(s.inp_tag, "S1")


class InpChangeTest(QuirkCase):
    def test_normal_flow(self) -> None:
        s = ready()
        h = FakeHooks()
        msgs = ir.receive(s, "DT10002+0000021.712M", h)
        self.assertEqual(msgs, [])
        self.assertEqual(s.cells["S1"].value, "21.712")     # CStr(Val("0000021.712"))
        self.assertEqual(s.paked, "DT")
        self.assertEqual(s.inp_tag, "S2")                   # 次の欄
        self.assertEqual(s.cells["S2"].color, grid.COLOR_YELLOW)
        self.assertEqual((s.inp, s.flag_label, s.auto, s.load), ("", 1, "on", "on"))
        # AllCheck は ATU があれば作業に関係なく CalculateSheets を呼ぶ(FormRelat:436)
        self.assertEqual([c[0] for c in h.calls], ["calculate_sheets", "loop_text", "posting", "out_tolerance"])
        self.assertEqual(h.calls[1], ("loop_text", "S1"))

    def test_register_check_blocks(self) -> None:
        s = ready()
        with self.assertRaises(Stop) as cm:
            ir.receive(s, "#L395.8", FakeHooks(registered=False))
        self.assertEqual(cm.exception.text, "検査員登録をして下さい")
        self.assertEqual((s.cells["S5"].value, s.inp, s.flag_label), ("", "", 1))

    def test_receive_ignored_while_flag_not_1(self) -> None:
        # UF1:3297 `If Me.FlagLabel <> 1 Then Exit Sub` → INP は変わらない。
        # その後 FlagLabel が戻っても、同じ文字列なら INP の Change が起きず何もしない
        s = ready()
        s.inp = "DT10002+0000021.712M"
        s.flag_label = 0
        h = FakeHooks()
        self.assertEqual(ir.receive(s, "#〆1.5", h), [])
        self.assertEqual(s.inp, "DT10002+0000021.712M")
        s.flag_label = 1
        self.assertEqual(ir.receive(s, "DT10002+0000021.712M", h), [])
        self.assertEqual((s.cells["S1"].value, h.calls), ("", []))

    def test_receive_empty_text_keeps_inp(self) -> None:
        # UF1:3298 `If UF1.StartTB = "" Then Exit Sub` → INP を "" にしない
        s = ready()
        s.inp = "#〆1.5"
        self.assertEqual(ir.receive(s, "", FakeHooks()), [])
        self.assertEqual(s.inp, "#〆1.5")

    def test_manual_entry_ignores_flag(self) -> None:
        # CommandButton3_Click (UF1:4552-4557) は FlagLabel を見ずに INP へ入れる。INP_Change が FlagLabel で抜ける
        s = ready()
        s.flag_label = 0
        self.assertEqual(ir.manual_entry(s, "1.5", FakeHooks()), [])
        self.assertEqual((s.inp, s.cells["S1"].value), ("#〆1.500", ""))

    def test_no_work_with_previous_target(self) -> None:
        s = KensaState(proofread_flag="ok", inp_tag="S1")
        with self.assertRaises(Stop) as cm:
            ir.receive(s, "#〆1.5", FakeHooks())
        self.assertEqual(cm.exception.text, "作業選択してください")
        self.assertEqual(s.cells["S1"].value, "")

    def test_not_calibrated(self) -> None:
        s = ready()
        s.proofread_flag = ""
        h = FakeHooks()
        self.assertEqual(ir.receive(s, "#〆1.5", h), [])
        self.assertIn(("proofread_load",), h.calls)
        self.assertEqual((s.cells["S1"].value, s.flag_label), ("", 1))
        s = ready(work_button_tag="jump")
        s.proofread_flag = ""
        ir.receive(s, "#〆1.5", FakeHooks())
        self.assertEqual(s.cells["S1"].value, "1.5")

    def test_yama_missing(self) -> None:
        s = ready("薄板", yama_visible=True)
        with self.assertRaises(Stop) as cm:
            ir.receive(s, "#〆1", FakeHooks())
        self.assertIn("山数の入力がありません", cm.exception.text)
        self.assertEqual(s.flag_label, 1)
        s = ready("薄板", yama_tag="別フォーム")
        with self.assertRaises(Stop):
            ir.receive(s, "#〆1", FakeHooks())

    def test_same_text_no_change_event(self) -> None:
        s = ready(inp="#〆1", flag_label=1)
        self.assertEqual(ir.receive(s, "#〆1", FakeHooks()), [])
        self.assertEqual(s.cells["S1"].value, "")

    def test_manual_target(self) -> None:
        s = ready(auto="off", inp_tag="DiagSKou")
        ir.receive(s, "#〆1.5", FakeHooks())
        self.assertEqual(s.diag_kou["S"].value, "1.5")

    def test_manual_entry(self) -> None:
        s = ready()
        ir.manual_entry(s, "12.3456", FakeHooks())
        self.assertEqual(s.cells["S1"].value, "12.346")      # "#〆" & Format(Manual, "0.000")

class EraseTest(unittest.TestCase):
    def test_click_then_dblclick(self) -> None:
        s = ready()
        h = FakeHooks()
        s.cells["S1"].value = "21.712"
        s.set_color("S1", grid.COLOR_NG)
        s.cells["S1"].tag = "pass"
        s.ng["S"].value = "NG"
        ir.text_tag(s, "S1", h)
        self.assertEqual((s.auto, s.inp_tag, s.cells["S1"].color), ("off", "S1", grid.COLOR_YELLOW))
        r = ir.cell_dbl_click(s, "S1", h)
        self.assertEqual((r.old_value, r.reason, r.pos, r.num, r.count), ("21.712", True, "S", 1, 1.0))
        self.assertEqual((s.cells["S1"].value, s.cells["S1"].tag, s.cells["S1"].color), ("", "", grid.COLOR_DEFAULT))
        self.assertEqual((s.delete_p, s.delete_c), ("S1", "1"))
        self.assertEqual(s.ng["S"].value, "")
        self.assertIn(("confi_reason_open", "S", 1, "21.712", 1.0), h.calls)
        # 同じ座標をもう一度
        s.cells["S1"].value = "21.700"
        ir.text_tag(s, "S1", h)
        r = ir.cell_dbl_click(s, "S1", h)
        self.assertEqual((s.delete_c, r.count), ("2", 2.0))
        # 空欄を消しても理由は聞かない
        ir.text_tag(s, "S1", h)
        r = ir.cell_dbl_click(s, "S1", h)
        self.assertFalse(r.reason)
        self.assertEqual(s.delete_c, "3")

    def test_rows_9_10_same_coordinate(self) -> None:
        s = ready()
        h = FakeHooks()
        s.cells["C9"].value = "1"
        s.diag["C"].value = "0.5"
        ir.text_tag(s, "C9", h)
        ir.cell_dbl_click(s, "C9", h)
        self.assertEqual((s.delete_p, s.delete_c, s.diag["C"].value), ("C10", "1", ""))
        s.cells["C10"].value = "1"
        ir.text_tag(s, "C10", h)
        ir.cell_dbl_click(s, "C10", h)
        self.assertEqual((s.delete_p, s.delete_c), ("C10", "2"))
        s2 = ready("極厚")
        s2.diag["C"].value = "0.5"
        ir.text_tag(s2, "C9", h)
        ir.cell_dbl_click(s2, "C9", h)
        self.assertEqual((s2.delete_p, s2.diag["C"].value), ("C9", "0.5"))

    def test_ng_erase_keeps_other_ng(self) -> None:
        s = ready()
        s.ng["E"].value = "NG"
        s.cells["E2"].value = "3.1"
        s.set_color("E2", grid.COLOR_NG)
        ir.ng_erase(s, "E1")
        self.assertEqual(s.ng["E"].value, "NG")
        s.set_color("E2", grid.COLOR_OK)
        ir.ng_erase(s, "E1")
        self.assertEqual(s.ng["E"].value, "")


class FixedQuirksTest(QuirkCase):
    """B55・B60・B61・B63・B87 の「再現」(既定)と「直す」。"""

    def _b60_state(self) -> KensaState:
        s = ready("薄板", at_butt_caption="厚4ヵ所", yama="4")
        fill(s, cells(FIVE, range(1, 11)))
        fill(s, [f"F{i}" for i in (1, 2, 5, 6, 7, 8, 9, 10)])
        for i in range(1, 11):
            s.cells[f"F{i}"].visible = True
        s.cells["F3"].visible = False
        s.cells["F4"].visible = False
        return s

if __name__ == "__main__":
    unittest.main()


class AtuitaThickRowsTest(unittest.TestCase):
    """2026-11-09(B93): 厚板に薄板の「厚4ヵ所」(2 ヵ所検査)が残っていても、DT は厚 1〜4 行を全部埋める。"""

    def _dt(self, s: KensaState) -> str:
        ir.pucket_move(s, "DT", "", "", FakeHooks())
        tag = s.inp_tag
        if tag:
            s.cells[tag].value = "15.0"
        return tag

    def test_user_sequence_reaches_s3(self) -> None:
        s = ready(at_butt_caption="厚4ヵ所")
        got = [self._dt(s) for _ in range(5)]
        # VBA は S1, S2, U1, U2, C1(S3 へ行かない)
        self.assertEqual(got, ["S1", "S2", "S3", "S4", "U1"])
        ir.look_blank(s, 1, 10)
        self.assertEqual(s.inp_tag, "S5")               # 厚が埋まったので次は幅(黄色の印も DT の入る先と合う)

    def test_marker_matches_next_dt(self) -> None:
        s = ready(at_butt_caption="厚4ヵ所")
        fill(s, ["S1", "S2"])
        ir.look_blank(s, 1, 10)
        marker = s.inp_tag
        ir.pucket_move(s, "DT", "", "", FakeHooks())
        self.assertEqual((marker, s.inp_tag), ("S3", "S3"))

    def test_disabled_thick_cells_skipped(self) -> None:
        s = ready()
        for p in FIVE:
            for i in (3, 4):
                if p != "S":
                    s.cells[f"{p}{i}"].enabled = False
        got = [self._dt(s) for _ in range(12)]
        self.assertEqual(got, ["S1", "S2", "S3", "S4", "U1", "U2", "C1", "C2", "D1", "D2", "E1", "E2"])
        msgs = []
        ir.pucket_move(s, "DT", "", "", FakeHooks(), msgs)
        self.assertEqual(msgs[0].text, "DT検査済")        # 使えない U3 などは待たない
        s.auto = "on"
        ir.look_blank(s, 1, 10)
        self.assertEqual(s.inp_tag, "S5")
