"""直尺の校正(Proofreading / UF校正)"""
from __future__ import annotations

import unittest
from datetime import datetime

from ufkensa.kensa import calibration as cal
from ufkensa.kensa import flow, quirks
from ufkensa.kensa.workers import CellWrite, ShiftTimes, WorkerRegistry

SHIFTS = ShiftTimes("07:00", "15:00", "15:00", "22:50", "22:50", "07:00", "08:15", "17:00")


def at(h: int, m: int = 0, s: int = 0, day: int = 30) -> datetime:
    return datetime(2026, 10, day, h, m, s)


def kiroku(*rows: tuple) -> dict:
    """校正記録(1 行目は見出し)。rows = (A, D) or (A, D, E)。"""
    c = {(1, 1): "校正日時", (1, 4): "直", (1, 5): "測定値", (1, 6): "検査員", (1, 7): "ライン"}
    for i, r in enumerate(rows, start=2):
        c[(i, 1)] = r[0]
        if r[1] is not None:
            c[(i, 4)] = r[1]
        if len(r) > 2:
            c[(i, 5)] = r[2]
    return c


def hinsho() -> dict:
    """品証検証(6 行目から 1日 1直/2直/3直 …)。"""
    c = {}
    for d in range(1, 32):
        r = 6 + (d - 1) * 3
        c[(r, 1)] = f"{d}日"
        for k in range(3):
            c[(r + k, 2)] = f"{k + 1}直"
    c[(6, 13)] = 5
    c[(6, 16)] = "2026.07.27 9時11分45秒"
    return c


class CheckTest(unittest.TestCase):
    kw = dict(facility="L-1", shifts=SHIFTS)

    def test_last_row_matches(self) -> None:
        r = cal.check_calibration("10:00", kiroku(("2026.10.30", "1直")), now=at(10), **self.kw)
        self.assertEqual((r.uncalibrated, r.proofread_flag), (False, "ok"))

    def test_other_shift_or_day(self) -> None:
        c = kiroku(("2026.10.30", "1直"))
        self.assertTrue(cal.check_calibration("16:00", c, now=at(16), **self.kw).uncalibrated)
        self.assertTrue(cal.check_calibration("10:00", c, now=at(10, day=31), **self.kw).uncalibrated)
        self.assertEqual(cal.check_calibration("16:00", c, now=at(16), **self.kw).proofread_flag, "")

    def test_only_last_row_counts(self) -> None:
        # 不合格の行(D 空)が最後にあれば未校正
        c = kiroku(("2026.10.30", "1直"), ("2026.10.30", None))
        self.assertTrue(cal.check_calibration("10:00", c, now=at(10), **self.kw).uncalibrated)

    def test_night(self) -> None:
        c = kiroku(("2026.10.29", "3直"))
        self.assertFalse(cal.check_calibration("00:30", c, now=at(0, 30), **self.kw).uncalibrated)

    def test_day_shift_line(self) -> None:
        c = kiroku(("2026.10.30", "日勤"))
        r = cal.check_calibration("23:00", c, now=at(23), facility="バランサー", shifts=SHIFTS)
        self.assertFalse(r.uncalibrated)

    def test_sheet_missing_and_jump(self) -> None:
        r = cal.check_calibration("10:00", None, now=at(10), **self.kw)
        self.assertEqual((r.uncalibrated, r.create_sheet), (True, True))
        r = cal.check_calibration("10:00", None, now=at(10), jump=True, **self.kw)
        self.assertEqual((r.uncalibrated, r.clear_jump_tag, r.proofread_flag), (False, True, None))
        self.assertFalse(cal.check_calibration("10:00", None, now=at(10), db_flag="x", **self.kw).uncalibrated)

    def test_input_blocked_each_time(self) -> None:
        # D11: 直が変わればその時点で止まる
        c = kiroku(("2026.10.30", "1直"))
        self.assertFalse(cal.input_blocked(c, now=at(14, 59), **self.kw))
        self.assertTrue(cal.input_blocked(c, now=at(15, 0), **self.kw))
        self.assertFalse(cal.input_blocked(c, now=at(15, 0), jump=True, **self.kw))

    def test_proofread_load(self) -> None:
        reg = WorkerRegistry(t1="阿部_佐藤")
        p = cal.proofread_load(reg, ("2026.10.30  07:10", None, None), "L-1", blocked=True)
        self.assertEqual(p.message.text, "校正が済んでいません。校正を起動します。")
        self.assertEqual(p.kname, "佐藤")
        self.assertIsNone(cal.proofread_load(reg, ("", "", ""), "L-1", blocked=False))


class TerminalTest(unittest.TestCase):
    def test_ter_conv(self) -> None:
        self.assertEqual(cal.ter_conv("DT10002+0000021.712M"), ("21.712", []))
        self.assertEqual(cal.ter_conv("#L395.8"), ("395.8", []))
        self.assertEqual(cal.ter_conv("ｄｔ10002+0000021.712M")[0], "21.712")   # 全角小文字も DT
        self.assertEqual(cal.ter_conv("#L2053.2")[0], "2053.2")
        v, m = cal.ter_conv("ABC")
        self.assertEqual((v, m[0].text), ("", "正しいデータが伝送できていません"))
        self.assertEqual(cal.ter_conv("#L1.5")[0], "")                          # 7 文字未満
        self.assertEqual(cal.ter_conv(""), ("", []))

    def test_safe_num(self) -> None:
        self.assertEqual(cal.safe_num("2,052.7"), (2052.7, True))
        self.assertEqual(cal.safe_num("2053..5"), (0.0, False))
        self.assertEqual(cal.safe_num(""), (0.0, False))
        self.assertEqual(cal.safe_num(None), (0.0, False))
        self.assertEqual(cal.safe_num("abc"), (0.0, False))

    def test_guidance(self) -> None:
        self.assertEqual(cal.open_guidance("L-1"), "[L-1] P2:1000mm直尺の全長を測定してください")
        self.assertEqual(cal.open_guidance("ﾄｯﾄ"), "[大板] P2:1000mm直尺の全長を測定してください")
        self.assertEqual(cal.open_guidance("丸徳"), "1000mm直尺の全長を測定してください")
        self.assertEqual(cal.gauge_guidance("L-1"), "[L-1]:直尺「S-2」の全長を測定してください")
        self.assertEqual(cal.gauge_guidance("HVC"), "[HVC]:直尺「S-1」の全長を測定してください")
        self.assertEqual(cal.gauge_guidance("丸徳A1"), "[中板]:直尺「xxx」の全長を測定してください")


class GaugeTest(unittest.TestCase):
    se = ("2052.7", "2053.7")

    def setUp(self) -> None:
        self.st = cal.CalibrationState("L-1", kname="阿部", mode=cal.WHITE)

    def tearDown(self) -> None:
        quirks.set_for_test(None)

    def run_gauge(self, raw: str, now=None, ws=None, se=se):
        return cal.input_gauge(self.st, raw, kiroku() if ws is None else ws, se, now=now or at(10, 5, 9),
                               facility="L-1", shifts=SHIFTS)

    def test_pass_both_ends(self) -> None:
        for raw in ("#L2052.7", "#L2053.7"):
            self.st = cal.CalibrationState("L-1", kname="阿部", mode=cal.WHITE)
            r = self.run_gauge(raw)
            self.assertTrue(self.st.proofread_ok)
            self.assertEqual(self.st.f1, "end")
        self.assertEqual(r.kiroku_writes, [CellWrite(2, 1, "2026.10.30"), CellWrite(2, 3, "10時05分09秒"),
                                           CellWrite(2, 4, "1直"), CellWrite(2, 5, "2053.7"),
                                           CellWrite(2, 6, "阿部"), CellWrite(2, 7, "L-1")])
        self.assertEqual(r.logs[0].msg, "*L-1_ｹﾞｰｼﾞ測定:2053.7")
        self.assertEqual(self.st.kname, "")
        self.assertTrue(cal.exit_gauge(self.st))

    def test_fail(self) -> None:
        quirks.set_for_test({})
        r = self.run_gauge("#L2052.6")
        self.assertFalse(self.st.proofread_ok)
        self.assertEqual(self.st.ng_value, "測定値 2052.6 mm")
        self.assertIn(CellWrite(2, 8, "〆"), r.kiroku_writes)
        self.assertNotIn(4, [w.col for w in r.kiroku_writes])          # D(直)は書かない
        self.assertEqual(r.kiroku_red, [])                             # B14: 赤字にしない
        self.assertEqual(r.logs[0].msg, "*L-1_ｹﾞｰｼﾞ測定:")               # 消した後に書くので値が残らない
        self.assertEqual(self.st.f1, "0")
        # F1=0 のままなので、クリアを押すまで次の入力は受け付けない
        self.assertEqual(self.run_gauge("#L2053.0").kiroku_writes, [])
        cal.clear_button(self.st)
        self.assertEqual(self.st.f1, "1")
        self.run_gauge("#L2053.0")
        self.assertTrue(self.st.proofread_ok)

    def test_fail_b66_fix_logs_value(self) -> None:
        quirks.set_for_test({"B66": "fix"})
        r = self.run_gauge("#L2052.6")
        self.assertEqual(r.logs[0].msg, "*L-1_ｹﾞｰｼﾞ測定:2052.6")            # 消す前の値をログに書く
        self.assertEqual(self.st.sokutei, "")                            # 画面は VBA どおり消える
        self.st = cal.CalibrationState("L-1", kname="阿部", mode=cal.WHITE)
        self.assertEqual(self.run_gauge("#L2053.0").logs[0].msg, "*L-1_ｹﾞｰｼﾞ測定:2053.0")

    def test_fail_b14_fix_red(self) -> None:
        quirks.set_for_test({"B14": "fix"})
        r = self.run_gauge("#L2053.8")
        self.assertEqual(r.kiroku_red, [(2, 1), (2, 3), (2, 5), (2, 6), (2, 7)])

    def test_night_mark(self) -> None:
        r = self.run_gauge("#L2053.0", now=at(6, 59, 0))
        self.assertEqual(r.kiroku_writes[:2], [CellWrite(2, 1, "2026.10.29"), CellWrite(2, 2, "夜勤日付処理")])
        self.assertIn(CellWrite(2, 4, "3直"), r.kiroku_writes)

    def test_bad_master(self) -> None:
        r = self.run_gauge("#L2053.0", se=("2053..5", "2053.7"))
        self.assertEqual(r.messages[0].text, "校正基準値の設定エラーです。管理者に連絡してください。")
        self.assertFalse(self.st.proofread_ok)
        self.st.f1 = "1"
        r = self.run_gauge("#L2053.0", se=None)                         # 設備名が DB に無い → 何もしない
        self.assertEqual(r.kiroku_writes, [])

    def test_green_mode_ignores_gauge(self) -> None:
        self.st.mode = cal.GREEN
        self.assertEqual(self.run_gauge("#L2053.0").logs, [])

    def test_missing_sheet(self) -> None:
        r = cal.input_gauge(self.st, "#L2053.0", None, self.se, now=at(10), facility="L-1", shifts=SHIFTS)
        self.assertTrue(r.kiroku_create)
        self.assertEqual(self.st.f1, "0")


class Hinsho1000Test(unittest.TestCase):
    def tearDown(self) -> None:
        quirks.set_for_test(None)

    def test_pass(self) -> None:
        st = cal.initialize("L-1", "阿部", "L-1")
        r = cal.input_1000(st, "#L1000.5", hinsho(), kiroku(), now=at(10, 5, 9), facility="L-1", shifts=SHIFTS)
        row = 6 + 29 * 3
        self.assertEqual(r.hinsho_writes, [CellWrite(row, 3, "1000.5"), CellWrite(row, 9, "L-1"),
                                           CellWrite(row, 10, "阿部"), CellWrite(row, 11, "2026.10.30 10時05分09秒")])
        self.assertEqual(r.logs[0].msg, "*L-1_1000mm測定:1000.5")
        self.assertEqual((st.mode, st.f1, st.gauge), (cal.WHITE, "1", "[L-1]:直尺「S-2」の全長を測定してください"))
        self.assertEqual(r.kiroku_rank.clear_columns, (9, 15))           # 合格後に 校正ランク付け(校正記録を書き換える)

    def test_pass_low_and_day_shift_written_as_1(self) -> None:
        st = cal.initialize("ﾊﾞﾗﾝｻｰ", "阿部", "ﾊﾞﾗﾝｻｰ")
        r = cal.input_1000(st, "#L999.50", hinsho(), None, now=at(16), facility="バランサー", shifts=SHIFTS)
        row = 6 + 29 * 3
        self.assertIn(CellWrite(row, 9, "ﾊﾞﾗﾝｻｰ"), r.hinsho_writes)      # 日勤は 1直 の行
        self.assertEqual(r.logs[0].target_path[-len("バランサー\\バランサー_Log"):], "バランサー\\バランサー_Log")

    def test_fail(self) -> None:
        st = cal.initialize("L-1", "阿部", "L-1")
        r = cal.input_1000(st, "#L999.40", hinsho(), None, now=at(10, 5, 9), facility="L-1", shifts=SHIFTS)
        self.assertEqual(st.ng_value, "999.5 ≦測定値≦ 1000.5 になっていません")
        self.assertEqual(r.hinsho_writes, [CellWrite(7, 16, "2026.10.30 10時05分09秒"), CellWrite(7, 13, "999.40"),
                                           CellWrite(7, 15, "阿部"), CellWrite(7, 14, "L-1")])
        self.assertEqual((st.mode, st.f1), (cal.GREEN, "1"))

    def test_test_line_range(self) -> None:
        st = cal.initialize("Test", "阿部", "Test")
        cal.input_1000(st, "#L0060.0", hinsho(), None, now=at(10), facility="Test", shifts=SHIFTS)
        self.assertEqual(st.mode, cal.WHITE)

    def test_missing_sheet_is_stop(self) -> None:
        st = cal.initialize("L-1", "阿部", "L-1")
        with self.assertRaises(flow.Stop):
            cal.input_1000(st, "#L1000.0", None, None, now=at(10), facility="L-1", shifts=SHIFTS)

    def test_skip_button_b14(self) -> None:
        quirks.set_for_test({})
        st = cal.initialize("HVC", "阿部", "HVC")
        self.assertTrue(cal.skip_1000_button(st))
        self.assertEqual((st.mode, st.gauge), (cal.WHITE, "[HVC]:直尺「S-1」の全長を測定してください"))
        quirks.set_for_test({"B14": "fix"})
        st = cal.initialize("HVC", "阿部", "HVC")
        self.assertFalse(cal.skip_1000_button(st))
        self.assertEqual(st.mode, cal.GREEN)


class RankTest(unittest.TestCase):
    def test_rank_only_red_rows(self) -> None:
        c = kiroku(("2026.09.30", None, 2053.9), ("2026.10.01", None, 2053.9), ("2026.10.02", None, "2052.5"),
                   ("2026.10.03", "1直", 2053.0), ("2026.10.04", None, 2054.1))
        c[(2, 6)] = c[(3, 6)] = c[(4, 6)] = c[(6, 6)] = "阿部"
        r = cal.calibration_rank(c, "L-1", now=at(10))
        self.assertEqual(r.ranks, {})                     # 赤字が無ければ何も出ない(B14 の状態)
        self.assertIn(CellWrite(4, 5, 2052.5), r.rewrites)
        r = cal.calibration_rank(c, "L-1", now=at(10), red_rows={2, 3, 4, 6})
        # 当月の最初の行(3 行目)から写す。J・N 列の最終行は消す前の値(=1)で数えるので、1 回目は表示されない
        self.assertIn(CellWrite(2, 10, 2053.9), r.writes)
        self.assertIn(CellWrite(3, 10, 2054.1), r.writes)
        self.assertIn(CellWrite(2, 14, 2052.5), r.writes)
        self.assertEqual(r.ranks, {})
        # 2 回目(写した I:O が残っている)で出る。3 位は LARGE の範囲外で 0 のまま
        for w in r.writes:
            c[(w.row, w.col)] = w.value
        r = cal.calibration_rank(c, "L-1", now=at(10), red_rows={2, 3, 4, 6})
        self.assertEqual(r.ranks, {1: "1位：　2026.10.04　値：2054.1　作業者：阿部",
                                   2: "2位：　2026.10.01　値：2053.9　作業者：阿部",
                                   4: "1位：　2026.10.02　値：2052.5　作業者：阿部"})

    def test_rank_b67_fix_shows_new_records_at_once(self) -> None:
        quirks.set_for_test({"B67": "fix"})
        try:
            c = kiroku(("2026.09.30", None, 2053.9), ("2026.10.01", None, 2053.9), ("2026.10.02", None, "2052.5"),
                       ("2026.10.03", "1直", 2053.0), ("2026.10.04", None, 2054.1))
            c[(2, 6)] = c[(3, 6)] = c[(4, 6)] = c[(6, 6)] = "阿部"
            r = cal.calibration_rank(c, "L-1", now=at(10), red_rows={2, 3, 4, 6})
            self.assertEqual(r.ranks, {1: "1位：　2026.10.04　値：2054.1　作業者：阿部",
                                       2: "2位：　2026.10.01　値：2053.9　作業者：阿部",
                                       4: "1位：　2026.10.02　値：2052.5　作業者：阿部"})
            # 前回の写し(I:O)が今回より長く残っていても、消した後で数えるので古い行は数えない
            c[(9, 10)] = 2060.0
            c[(9, 9)] = "古い"
            r = cal.calibration_rank(c, "L-1", now=at(10), red_rows={2, 3, 4, 6})
            self.assertEqual(r.ranks[1], "1位：　2026.10.04　値：2054.1　作業者：阿部")
        finally:
            quirks.set_for_test(None)

    def test_hinsho_rank_exits_on_m6(self) -> None:
        r = cal.hinsho_rank(hinsho(), now=at(10))
        self.assertEqual(r.ranks, {})
        self.assertIn(CellWrite(6, 13, 5.0), r.rewrites)


if __name__ == "__main__":
    unittest.main()
