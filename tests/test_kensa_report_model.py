"""帳票の中身(report_model.py)。期待値は Print.bas を手で追った結果(行番号はコメント)"""
from __future__ import annotations

import shutil
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from ufkensa.kensa import quirks, records as rc, report_model as rm, store as st
from ufkensa.kensa.flow import Question, Stop
from tests.test_kensa_records import make_snap

NOW = datetime(2024, 11, 2, 8, 55, 3)


class PartsTest(unittest.TestCase):
    def tearDown(self) -> None:
        pass
    def test_circled(self) -> None:
        self.assertEqual([rm.circled_number(n) for n in (1, 20, 21, 35, 36, 50)], ["①", "⑳", "㉑", "㉟", "㊱", "㊿"])
        self.assertEqual(rm.circled_number(51), "51")
        self.assertEqual(rm.circled_number(0), "0")

    def test_format_cell(self) -> None:
        f = "0.000_);[赤](0.00)"
        self.assertEqual(rm.format_cell(49.776, f), ("49.776", False))
        self.assertEqual(rm.format_cell(-0.125, f), ("(0.13)", True))           # 負は 2 桁・赤
        self.assertEqual(rm.format_cell(120.11, "0.00_);[赤](0.00)"), ("120.11", False))
        self.assertEqual(rm.format_cell(17.040000915527344, "0.0_ "), ("17.0", False))
        self.assertEqual(rm.format_cell("A:x", f), ("A:x", False))
        self.assertEqual(rm.format_cell(3.0, "General"), ("3", False))
        self.assertEqual(rm.format_cell(None, f), ("", False))

    def test_fit_text(self) -> None:
        self.assertEqual(rm.fit_text_font_size("A:△ｶﾐｺﾐ跡\r\nB:△ｽﾘｷｽﾞ", 11.625, 23.25), 8.0)
        self.assertLess(rm.fit_text_font_size("あ" * 400, 11.625, 23.25), 4.0)


class Base(unittest.TestCase):
    def setUp(self) -> None:
        self.dir = Path(tempfile.mkdtemp())
        self.store = st.LocalStore(self.dir / "local.sqlite3")
        self.store.add_sheet("WorkerTime_Test").write(3, 1, "検査員01")
        self.env = rm.PrintEnv(now=NOW, start1="7:00", end3="6:50")

    def tearDown(self) -> None:
        self.store.close()
        shutil.rmtree(self.dir, ignore_errors=True)

    def run_print(self, snap=None, env=None) -> rm.PrintResult:
        return rm.print_transfer(self.store, snap or make_snap(), env or self.env)


class PrintTransferTest(Base):
    def test_header_cells(self) -> None:
        s = make_snap(ken_hbox=["200", "226", "", "", "", "", ""], resize_tag="有")
        res = self.run_print(s)
        self.assertFalse(res.stopped)
        rep = res.report
        self.assertEqual(rep.text("C2"), "H7372A0")
        self.assertEqual(rep.text("B1"), "ｻｲｽﾞ修正")
        self.assertEqual(rep.text("C5"), "-0.500")                 # "'-" で文字列のまま(Print:431)
        self.assertEqual(rep.cell("C5")["value"], "-0.500")
        self.assertEqual(rep.text("C6"), "49.500")                 # 数になり 0.000_ で表示
        self.assertEqual(rep.text("G2"), "426.0")                  # 検入高さの合計(Single)
        self.assertEqual(rep.text("H2"), "2山")
        self.assertEqual(rep.text("B4"), "Test")
        self.assertEqual(rep.text("B9"), "検査員01")
        self.assertEqual(rep.text("L6"), "0.0")                    # 対角公差が空でも 0.0
        self.assertEqual(rep.text("C12"), "A:△ｶﾐｺﾐ跡\nB:△ｽﾘｷｽﾞ")
        self.assertEqual(rep.text("E12"), "A:丈上0.20\n")          # "_" を消し、B は "__" → ""
        self.assertEqual(rep.text("B13"), "S1")                    # 厚板の位置名
        self.assertEqual(rep.text("I18"), "差")
        self.assertEqual(rep.page["print_area"], "$B$1:$L$24")
        self.assertEqual(rep.page["zoom"], 115)
        self.assertEqual(rep.page["header_right"], "H7372A0 検査員01")
        self.assertEqual(rep.page["footer_right"], " 2024/11/02  5:29～2024/11/02  8:55 印刷")
        self.assertEqual(res.end_time, "2024/11/02 8:55:03")
        self.assertEqual(self.store.sheet("WorkerTime_Test").get(1, 3), "2024/11/02 8:55:03")

    def test_kenhbox7_count(self) -> None:
        rep = self.run_print(make_snap(ken_hbox=["200", "", "", "", "", "", "5"])).report
        self.assertEqual(rep.text("F2"), "検入枚数")
        self.assertEqual(rep.text("G2"), "5枚")
        self.assertEqual(rep.text("H2"), "")

    def test_thickness_digits(self) -> None:
        s = make_snap(atu="120.000")
        for f in "SUCDE":
            s.measure[f][0] = "120.110"
        rep = self.run_print(s).report
        self.assertEqual(rep.text("G12"), "120.11")                # 厚 > 74 → 2 桁(Print:625)
        s2 = make_snap(atu="74.000")
        self.assertEqual(self.run_print(s2).report.text("G12"), "49.776")


    def test_thin_yama_labels(self) -> None:
        s = make_snap(kind="薄板", usuita_caption="L", yama="3")
        rep = self.run_print(s).report
        self.assertEqual([rep.text(f"B{r}") for r in range(12, 17)], ["3-S", "3-E", "2-2", "1-S", "1-E"])
        s4 = make_snap(kind="薄板", usuita_caption="L", yama="4")
        rep4 = self.run_print(s4).report
        self.assertEqual(rep4.text("B17"), "1-E")
        self.assertEqual(rep4.cell("B17")["borders"]["top"], ["dot", "thin"])

    def test_slab_labels(self) -> None:
        rep = self.run_print(make_snap(kind="スラブ")).report
        self.assertEqual([rep.text(f"B{r}") for r in range(19, 24)], ["1上", "1下", "2上", "2下", ""])

    def test_lateral_column(self) -> None:
        rep = self.run_print(make_snap(sumisei="白銅")).report
        self.assertEqual(rep.text("J18"), "ﾗﾃﾗﾙﾎﾞｰ")
        self.assertEqual(rep.text("J19"), "1.5/L")
        rep2 = self.run_print(make_snap()).report
        self.assertEqual(rep2.text("J18"), "")                      # 要らなければ消す

    def test_r115_columns_vanish(self) -> None:
        # U10: R115 では J/K に写したラテラル・直角度を直後の「ラテ=""」の処理が消す(VBA のまま)
        s = make_snap(sumisei="白銅", youto="R115")
        s.diag_kou.update({f: "0.3" for f in "SUCDE"})
        rep = self.run_print(s).report
        self.assertEqual(rep.text("J18"), "")
        self.assertEqual(rep.text("K18"), "")

    @unittest.skip("U10 実機確認待ち(R115 の印刷物にラテラルボー/直角度の列が出ているか)")
    def test_r115_columns_on_paper(self) -> None:
        pass

    def test_worker_stops(self) -> None:
        res = self.run_print(make_snap(worker_shift=""))
        self.assertTrue(res.stopped)
        self.assertEqual(res.messages[-1].text, "実行時エラー！検査員登録をやり直してください。(変数なし)")
        self.store.add_sheet("WorkerTime_Test")
        res2 = self.run_print()
        self.assertTrue(res2.stopped)
        self.assertEqual(res2.messages[-1].text,
                         "検査員が空白です。WorkerTime上に履歴がない可能性有。 担当者に連絡してください。")
        self.store.delete_sheet("WorkerTime_Test")
        with self.assertRaises(Stop):
            self.run_print()

    def test_history_loaded_worker(self) -> None:
        env = rm.PrintEnv(now=NOW, start1="7:00", end3="6:50", worker_name_visible=True, worker_name_l1="阿部")
        rep = self.run_print(make_snap(data_loading="Load"), env).report
        self.assertEqual(rep.text("B9"), "阿部")
        self.assertEqual(rep.text("B3"), "履歴呼び出しから印刷")

    def test_abnormal_table(self) -> None:
        iz = self.store.add_sheet("異常ロット用_Test")
        iz.write(1, 1, "H7372A0")
        iz.write(2, 1, "50.6")
        iz.write(2, 5, "1530.1")
        iz.write(2, 19, "S")
        rep = self.run_print().report
        self.assertEqual(rep.text("A26"), "公差外ﾃﾞｰﾀ")
        self.assertEqual(rep.text("B26"), "厚1")
        self.assertEqual(rep.text("O26"), "ｹﾞ-ｼﾞ")
        self.assertEqual(rep.text("A27"), "S")
        self.assertEqual(rep.text("B27"), "50.600")
        self.assertEqual(rep.text("F27"), "1530.1")
        self.assertEqual(rep.figure_at, "L11")
        self.assertEqual(rep.page["print_area"], "$A$1:$O$28")
        sec = rep.sections()
        self.assertEqual(sec["abnormal"]["rows"][0][:2], ["S", "50.600"])

    def test_saved_and_reprint(self) -> None:
        self.run_print()
        rm.set_printed(self.store, "H7372A0", "Test", NOW)
        self.assertTrue(rm.is_printed(self.store, "H7372A0"))
        q = rm.print_only(self.store, "Test", "H7372A0", {}, NOW)
        self.assertIsInstance(q, Question)
        self.assertEqual(q.text, "このLOTは印刷済みです。再印刷しますか？")
        self.assertIsNone(rm.print_only(self.store, "Test", "H7372A0", {"reprint": False}, NOW))
        res = rm.print_only(self.store, "Test", "H7372A0", {"reprint": True}, NOW)
        self.assertEqual(res.report.text("D2"), "再印刷")
        self.assertEqual(res.report.text("C2"), "H7372A0")
        self.assertEqual(res.messages[-1].text, "再印刷が完了しました")


class DefectTest(Base):
    def defect(self) -> st.GridSheet:
        sh = rm.ensure_defect_sheet(self.store, "Test", "H7372A0")
        sh.write(3, 3, "2")
        sh.write(5, 3, "3")
        sh.write(4, 24, "1")
        return sh

    def test_template_and_sums(self) -> None:
        sh = self.defect()
        self.assertEqual(sh.get(2, 3), "フクレ")
        page = rm.defect_print_range(self.store, "Test", "検査員01", "H7372A0", now=NOW, start_time="2024/11/02 05:29:10",
                                     start1="7:00", end3="6:50", end_time_caption="2024/11/02 8:55:03")
        self.assertEqual(page["print_area"], "$A$1:$Y$12")
        self.assertEqual(page["footer_right"], " 2024/11/02  5:29～2024/11/02  8:55 印刷")
        self.assertEqual(sh.get(3, 3), "②")
        self.assertEqual(sh.get(10, 3), "⑤")
        self.assertEqual(sh.get(10, 24), "①")
        self.assertIsNone(sh.get(10, 4))                           # 合計 0 は空

    def test_mitsubishi_no_footer(self) -> None:
        self.defect()
        p = rm.defect_print_range(self.store, "Test", "検査員01", "L", now=NOW, start1="7:00", end3="6:50",
                                  mitsubishi=True)
        self.assertEqual(p["footer_right"], "")

    def test_render(self) -> None:
        self.defect()
        m = rm.render_defect_sheet(self.store, "Test")
        self.assertEqual(m.text("C1"), "鋳造")
        self.assertEqual(m.text("C3"), "2")


class PrintHistoryTest(Base):
    def test_history(self) -> None:
        rm.set_printed(self.store, "L1", "Test", datetime(2026, 10, 25, 9, 0))
        rm.set_printed(self.store, "L2", "Test", datetime(2026, 10, 29, 9, 0))
        rm.set_printed(self.store, "L2", "Test", datetime(2026, 10, 30, 9, 0))    # 同じ LOT は日時だけ更新
        sh = self.store.sheet(st.PRINT_HISTORY)
        self.assertEqual(sh.last_row(1), 3)
        self.assertEqual(sh.get(3, 2), "2026/10/30 9:00:00")
        self.assertEqual(rm.print_button_state(self.store, "L1")["caption"], "再印刷")
        self.assertEqual(rm.print_button_state(self.store, "X")["back_color"], rm.RGB_GRAY)
        st.import_cells(self.store, st.PRINT_HISTORY, [(4, 1, "L346A20"), (4, 2, "2026.06.22"), (4, 3, "TEST")])
        removed = rm.clean_old_history(self.store, datetime(2026, 10, 30))
        self.assertEqual(removed, ["L1"])                          # 5 日前 > 3。「.」区切りは日付と見ない(U12)
        rm.clear_printed(self.store, "L2")
        self.assertFalse(rm.is_printed(self.store, "L2"))


if __name__ == "__main__":
    unittest.main()
