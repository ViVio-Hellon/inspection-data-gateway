"""作業者: 直と日付・名簿・検査員シート・作業時間・照合・日報(Worker / UFNameList)"""
from __future__ import annotations

import unittest
from datetime import datetime

from tests import kensa_fixture
from ufkensa.kensa import quirks, vbx, workers
from ufkensa.kensa.workers import CellWrite, ShiftTimes, WorkerRegistry

# DB「時間用」の値(fixtures)
SHIFTS = ShiftTimes("07:00", "15:00", "15:00", "22:50", "22:50", "07:00", "08:15", "17:00")


def at(h: int, m: int = 0, s: int = 0, day: int = 30) -> datetime:
    return datetime(2026, 10, day, h, m, s)


def tc(t: str, fac: str = "L-1") -> str:
    return workers.time_check(t, facility=fac, shifts=SHIFTS)


class ShiftTest(unittest.TestCase):
    def test_load_from_db(self) -> None:
        kensa_fixture.use_offline()
        self.assertEqual(workers.load_shift_times(), SHIFTS)

    def test_less_than_4_rows_is_error_9(self) -> None:
        with self.assertRaises(IndexError):
            workers.shift_from_records(["番号", "直", "開始", "終了"], [["1", "1", "07:00", "15:00"]])

    def test_empty_table_keeps_previous(self) -> None:
        self.assertEqual(workers.shift_from_records(None, None, SHIFTS), SHIFTS)

    def test_time_check_boundaries(self) -> None:
        # Worker:261-271
        self.assertEqual(tc("07:00"), "3直")          # Start1 < t(含まない)→ 0:00〜End3 の 3直
        self.assertEqual(tc("07:00:01"), "1直")       # 秒まで渡すと 1直
        self.assertEqual(tc("07:01"), "1直")
        self.assertEqual(tc("14:59:59"), "1直")
        self.assertEqual(tc("15:00"), "2直")          # Start2 <= t
        self.assertEqual(tc("22:49"), "2直")
        self.assertEqual(tc("22:50"), "3直")
        self.assertEqual(tc("23:59"), "3直")
        self.assertEqual(tc("23:59:30"), "3直")       # どれにも当たらない → 3直
        self.assertEqual(tc("00:00"), "3直")
        self.assertEqual(tc("10:00", "バランサー"), "日勤")
        self.assertEqual(tc("10:00", "存在なし"), "日勤")
        self.assertEqual(tc("03:00", "Test"), "3直")

    def test_empty_shift_times_is_type_mismatch(self) -> None:
        with self.assertRaises(vbx.TypeMismatch):
            workers.time_check("10:00", facility="L-1", shifts=ShiftTimes())

    def test_yea_dayr_check(self) -> None:
        f = workers.yea_dayr_check
        self.assertEqual(f("06:30", now=at(6, 30), facility="L-1", shifts=SHIFTS), "2026.10.29")
        self.assertEqual(f("07:00", now=at(7, 0), facility="L-1", shifts=SHIFTS), "2026.10.29")
        self.assertEqual(f("07:01", now=at(7, 1), facility="L-1", shifts=SHIFTS), "2026.10.30")
        self.assertEqual(f("23:00", now=at(23), facility="HVC", shifts=SHIFTS), "2026.10.30")
        self.assertEqual(f("03:00", now=at(3), facility="トット", shifts=SHIFTS), "2026.10.30")

    def test_day_check(self) -> None:
        # Worker:209-213: 前日扱いは「前日  (時+24):分」
        self.assertEqual(workers.day_check("03:15", now=at(3, 15), facility="L-1", shifts=SHIFTS),
                         "2026.10.29  27:15")
        self.assertEqual(workers.day_check("00:05", now=at(0, 5), facility="L-1", shifts=SHIFTS),
                         "2026.10.29  24:05")
        self.assertEqual(workers.day_check("10:05", now=at(10, 5), facility="L-1", shifts=SHIFTS),
                         "2026.10.30  10:05")


class NameTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        kensa_fixture.use_offline()

    def test_name_conv(self) -> None:
        # 名簿の列は NO(VBA の Filter は [No]。ADO は列名の大文字小文字を区別しない)
        self.assertEqual(workers.name_conv("9000000000001"), ("検査員01", []))
        self.assertEqual(workers.name_conv("9000000000002", test_line=True)[0], "検査員02")

    def test_unregistered(self) -> None:
        name, msgs = workers.name_conv("1234567890123")
        self.assertEqual(name, "")
        self.assertEqual(msgs[0].sound, "ナンバーNO.wav")
        self.assertEqual(workers.name_conv(""), ("", []))

    def test_register_card(self) -> None:
        look = workers.name_conv
        self.assertIsNone(workers.register_card("900000000000", WorkerRegistry(), now=at(10), facility="L-1",
                                                shifts=SHIFTS, line="L-1", name_lookup=look))
        r = workers.register_card("9000000000001", WorkerRegistry(t1="阿部"), now=at(10, 5), facility="L-1",
                                  shifts=SHIFTS, line="L-1", name_lookup=look)
        self.assertEqual(r.registry.t1, "阿部_検査員01")
        self.assertEqual(r.registry.day1, "2026.10.30  10:05")
        self.assertEqual(r.log.msg, "*UFKen:検査員登録_2026.10.30  10:05:阿部_検査員01")
        self.assertTrue(r.log.target_path.endswith("L-1\\L-1_Log"))
        # 夜勤の 0 時以降は 3直・前日の「26:00」表記。Test のログはアプリのフォルダ(UF1:4842)
        r = workers.register_card("9000000000001", WorkerRegistry(), now=at(2), facility="Test", shifts=SHIFTS,
                                  line="Test", name_lookup=look)
        self.assertEqual((r.shift, r.registry.t3, r.registry.day3), ("3直", "検査員01", "2026.10.29  26:00"))
        self.assertEqual(r.log.target_path, "")
        # 未登録は何も変えない
        r = workers.register_card("1234567890123", WorkerRegistry(t1="x"), now=at(10), facility="L-1",
                                  shifts=SHIFTS, line="L-1", name_lookup=look)
        self.assertEqual((r.name, r.registry.t1, r.log), ("", "x", None))


class RegistrySheetTest(unittest.TestCase):
    def test_worker_save_merges_same_names(self) -> None:
        w = workers.worker_save_rows(WorkerRegistry(t1="阿部_阿部", t2="阿部_佐藤", day1="d1"))
        self.assertIn(CellWrite(1, 2, "阿部"), w)
        self.assertIn(CellWrite(2, 2, "阿部_佐藤"), w)
        self.assertIn(CellWrite(1, 1, "d1"), w)

    def test_worke_day_max(self) -> None:
        f = workers.worke_day_max
        self.assertEqual(f("", "", ""), "")
        self.assertEqual(f("2026.10.30  07:10", None, None), "1")
        self.assertEqual(f("2026.10.30  07:10", "2026.10.30  15:10", None), "2")
        self.assertEqual(f("2026.10.30  07:10", "2026.10.29  23:10", None), "1")
        # 夜勤の「27:15」表記は 27 時として比べる
        self.assertEqual(f("2026.10.29  07:10", "2026.10.29  15:10", "2026.10.29  27:15"), "3")
        # 分は比べない: 同じ時なら先の直
        self.assertEqual(f("2026.10.30  15:59", "2026.10.30  15:01", None), "1")
        # 1,2 が同日で 3 が後の日
        self.assertEqual(f("2026.10.29  07:10", "2026.10.29  15:10", "2026.10.30  23:00"), "3")
        self.assertEqual(f("2026.10.29  07:10", "2026.10.30  15:10", "2026.10.29  23:00"), "2")

    def test_get_worker_name(self) -> None:
        reg = WorkerRegistry(t1="阿部_佐藤", t2="")
        self.assertEqual(workers.get_worker_name("1", reg), "佐藤")
        self.assertEqual(workers.get_worker_name("", reg), "")
        with self.assertRaises(IndexError):            # Split("") の UBound で実行時エラー 9
            workers.get_worker_name("2", reg)
        quirks.set_for_test({"B69": "fix"})
        try:
            self.assertEqual(workers.get_worker_name("2", reg), "")   # 直した側: 空なら名前なし
        finally:
            quirks.set_for_test(None)
        self.assertEqual(workers.workers_text("1", reg), " T1:阿部_佐藤")
        self.assertEqual(workers.workers_text("", reg), "")

    def test_register_check(self) -> None:
        reg = WorkerRegistry(t1="阿部", day1="2026.10.30  07:10")
        kw = dict(now=at(10), facility="L-1", shifts=SHIFTS)
        self.assertFalse(workers.register_check("10:00:00", reg, ("2026.10.30  07:10", None, None), **kw))
        self.assertTrue(workers.register_check("16:00:00", reg, ("2026.10.30  07:10", None, None), **kw))
        self.assertTrue(workers.register_check("10:00:00", reg, None, **kw))
        self.assertTrue(workers.register_check("10:00:00", reg, ("2026.10.29  07:10", None, None), n="新規", **kw))
        self.assertFalse(workers.register_check("10:00:00", reg, ("2026.10.30  07:10", "", None), n="新規", **kw))


class WorkerTimeTest(unittest.TestCase):
    def tearDown(self) -> None:
        quirks.set_for_test(None)

    def test_end_time_register_7_00(self) -> None:
        quirks.set_for_test({})
        f = workers.end_time_register
        self.assertEqual(f(at(6, 59, 59))[0], "2026/10/29 06:59:59")    # 日付だけ 1 日戻す
        self.assertEqual(f(at(7, 0, 0))[0], "2026/10/29 07:00:00")      # 7:00 ちょうどは前日
        self.assertEqual(f(at(7, 0, 1))[0], "2026/10/30 7:00:01")       # CStr(Date) の書き方
        self.assertEqual(f(at(23, 59, 30))[0], "2026/10/30 23:59:30")
        self.assertEqual(f(at(13, 5, 9))[1], CellWrite(1, 3, "2026/10/30 13:05:09"))

    def test_end_time_register_b38_fix(self) -> None:
        quirks.set_for_test({"B38": "fix"})
        sh = ShiftTimes(end3="06:30")
        self.assertEqual(workers.end_time_register(at(6, 45), sh)[0], "2026/10/30 6:45:00")
        quirks.set_for_test({})
        self.assertEqual(workers.end_time_register(at(6, 45), sh)[0], "2026/10/29 06:45:00")

    def test_is_over_12_hours_b19(self) -> None:
        quirks.set_for_test({})
        f = workers.is_over_12_hours
        self.assertTrue(f("2026/10/30 9:59:00", at(10, 0)))      # 時の境界をまたげば 1 分でも True
        self.assertFalse(f("2026/10/30 9:01:00", at(9, 59)))
        self.assertFalse(f("", at(10)))
        self.assertFalse(f("2026/10/30 9:00:00", at(10), sheet_exists=False))
        quirks.set_for_test({"B19": "fix"})
        self.assertFalse(f("2026/10/30 9:59:00", at(10, 0)))
        self.assertTrue(f("2026/10/29 22:00:00", at(10, 0)))

    def test_name_match_b21(self) -> None:
        cells = {(3, 1): "阿部", (4, 1): "阿部", (3, 2): "佐藤", (7, 12): "鈴木",
                 (8, 1): "田中", (3, 13): "山田", (1, 18): "OUT"}
        reg = WorkerRegistry(t1="阿部_佐藤", t3="鈴木")
        quirks.set_for_test({})
        r = workers.name_match(cells, reg)
        self.assertFalse(r.out)                           # F 面(8 行)と 13 列は見ない
        self.assertIn(CellWrite(2, 15, "阿部"), r.writes)
        self.assertIn(CellWrite(3, 15, "佐藤"), r.writes)  # 列→行の順に番号
        self.assertIn(CellWrite(4, 15, "鈴木"), r.writes)
        self.assertIn(CellWrite(1, 16, "登録検査員"), r.writes)
        self.assertNotIn(CellWrite(1, 18, ""), r.writes)   # R1 の "OUT" は残る
        self.assertTrue(r.save_workbook)
        quirks.set_for_test({"B21": "fix"})
        r = workers.name_match(cells, reg)
        self.assertTrue(r.out)
        self.assertEqual(r.open_name_list, 1)
        q = [w.value for w in r.writes if w.col == 17]
        self.assertEqual(q, ["一致無し検査員", "田中", "山田"])
        self.assertEqual(r.writes[0], CellWrite(1, 18, ""))

    def test_name_match_space_splits_names(self) -> None:
        quirks.set_for_test({})
        r = workers.name_match({(3, 1): "阿部 久知"}, WorkerRegistry(t1="阿部 久知"))
        # 登録検査員は Split(Join(...)) で半角空白でも分かれる → 照合で一致しない(VBA どおり)
        self.assertTrue(r.out)

    def test_name_match_b86_clears_out(self) -> None:
        quirks.set_for_test({})
        reg = WorkerRegistry(t1="阿部")
        r = workers.name_match({(1, 18): "OUT", (3, 1): "阿部"}, reg)
        self.assertNotIn(18, [w.col for w in r.writes])                      # VBA: R1 の OUT は残る
        quirks.set_for_test({"B86": "fix"})
        r = workers.name_match({(1, 18): "OUT", (3, 1): "阿部"}, reg)
        self.assertEqual([w for w in r.writes if w.col == 18], [CellWrite(1, 18, "")])
        r = workers.name_match({(1, 18): "OUT", (3, 1): "田中"}, reg)
        self.assertEqual([w.value for w in r.writes if w.col == 18], ["", "OUT"])  # 不一致なら書き直す

    def test_name_match_b68_fix(self) -> None:
        quirks.set_for_test({"B68": "fix"})
        reg = WorkerRegistry(t1="阿部 久知_佐藤", t3="鈴木")
        r = workers.name_match({(3, 1): "阿部 久知", (4, 1): "佐藤", (3, 2): "阿部"}, reg)
        self.assertEqual([w.value for w in r.writes if w.col == 16], ["登録検査員", "阿部 久知", "佐藤", "鈴木"])
        self.assertEqual([w.value for w in r.writes if w.col == 17], ["一致無し検査員", "阿部"])
        r = workers.name_match({(3, 1): "阿部 久知", (4, 1): "佐藤"}, reg)
        self.assertFalse(r.out)


HIST = [
    # 2 行目からの Lot情報履歴(列 1..26)
    ["2026-06-03", "57076309", "N456J52", 3, 1000, 2000, "K650969", "995", "H24", "H174", "", "", "", "",
     150, 2439, 1, 0, None, None, 16.26, "60524074", 0, 2000, "", "1P0116"],
    ["2026-06-03", "04077194", "H3416A0", 15, 1528, 3053, "H651336", "F52S", "R", "H279", "", "", "", "",
     15, 2812.9, 0, 1, "208PRWH", "208PRWH", 187.53235, "60602048", 0, 0, "", "1P1202"],
]


def ctx(now: datetime, cells=None, lot: str = "N456J52", **kw) -> workers.DailyContext:
    c = {(2, 4): "梱包No"} if cells is None else cells
    return workers.DailyContext(True, c, HIST, lot, now, "Test", SHIFTS,
                                registry=WorkerRegistry(t1="阿部", t2="佐藤", t3="鈴木"), uf1_line="Test", **kw)


class DailyTest(unittest.TestCase):
    def tearDown(self) -> None:
        quirks.set_for_test(None)

    def test_report_1_and_2(self) -> None:
        c = ctx(at(10, 5))
        w = workers.daily_report(1, c)
        self.assertEqual(w, [CellWrite(1, 1, "2026.10.30_1直"), CellWrite(1, 2, "Test"), CellWrite(1, 3, "阿部")])
        w = workers.daily_report(2, c)
        self.assertIn(CellWrite(3, 1, "N456J52"), w)
        self.assertIn(CellWrite(3, 2, "995-H24"), w)
        self.assertIn(CellWrite(3, 3, "3.000×1000.0×2000.0"), w)
        self.assertIn(CellWrite(3, 11, "有"), w)
        self.assertIn(CellWrite(3, 25, "16.26"), w)
        c = ctx(at(10, 5), lot="H3416A0")
        w = workers.daily_report(2, c)
        self.assertIn(CellWrite(3, 11, "無"), w)
        self.assertIn(CellWrite(3, 14, "A:208PRWH_B:208PRWH"), w)

    def test_line_b1(self) -> None:
        c = ctx(at(10))
        c.uf1_line = "丸徳"                       # UF1.Maru は Case Else → "Test"
        self.assertIn(CellWrite(1, 2, "Test"), workers.daily_report(1, c))
        c.uf1_line = "ﾊﾞﾗﾝｻｰ"
        self.assertIn(CellWrite(1, 2, "ﾊﾞﾗﾝｻｰ"), workers.daily_report(1, c))

    def test_report_3_b18(self) -> None:
        quirks.set_for_test({})
        w = workers.daily_report(3, ctx(at(9, 5)))
        self.assertEqual(w[-2:], [CellWrite(3, 6, "09"), CellWrite(3, 6, "05")])   # 分で時を上書き
        w = workers.daily_report(3, ctx(at(9, 15)))
        self.assertEqual(w[-2:], [CellWrite(3, 6, "09"), CellWrite(3, 7, "15")])
        quirks.set_for_test({"B18": "fix"})
        w = workers.daily_report(3, ctx(at(9, 5)))
        self.assertEqual(w[-2:], [CellWrite(3, 6, "09"), CellWrite(3, 7, "05")])

    def test_report_4(self) -> None:
        cells = {(2, 4): "梱包No", (3, 1): "N456J52", (3, 6): "09", (3, 7): "11"}
        c = ctx(at(10, 7), cells)
        c.ainput = workers.AinputValues("N456J52", "01", "50", "2", "", "")
        w = workers.daily_report(4, c)
        self.assertEqual(w, [CellWrite(4, 1, "N456J52"), CellWrite(3, 4, "01"), CellWrite(4, 1, "N456J52"),
                             CellWrite(3, 4, "01"), CellWrite(3, 8, "10"), CellWrite(3, 9, "07"),
                             CellWrite(3, 12, 50), CellWrite(3, 13, 2), CellWrite(4, 6, "10"),
                             CellWrite(4, 7, "07")])
        # 2 段(包数2 あり)は 2 行に分け、終了時刻は次の行に書く
        cells = {(2, 4): "梱包No", (3, 1): "N456J52"}
        c = ctx(at(10, 7), cells)
        c.ainput = workers.AinputValues("N456J52", "02", "50", "2", "30", "1")
        w = workers.daily_report(4, c)
        self.assertIn(CellWrite(4, 8, "10"), w)
        self.assertIn(CellWrite(4, 12, 30), w)
        self.assertIn(CellWrite(5, 1, "N456J52"), w)
        self.assertIn(CellWrite(5, 6, "10"), w)

    def test_report_no_lot_or_no_sheet(self) -> None:
        self.assertEqual(workers.daily_report(1, ctx(at(10), lot="XXX")), [])
        c = ctx(at(10))
        c.exists = False
        self.assertEqual(workers.daily_report(1, c), [])
        c = ctx(at(10))
        c.history_rows = None
        with self.assertRaises(IndexError):
            workers.daily_report(1, c)
        quirks.set_for_test({"B69": "fix"})
        self.assertEqual(workers.daily_report(1, c), [])            # 直した側: 一致なしと同じく何もしない
        self.assertEqual(workers.daily_create(c).writes[-1], CellWrite(2, 25, " 単重"))

    def test_daily_a1_7_00_b70(self) -> None:
        quirks.set_for_test({})
        c = ctx(at(7, 0, 30))
        c.facility = "L-1"
        self.assertIn(CellWrite(1, 1, "2026.10.29_1直"), workers.daily_report(1, c))   # VBA: 前日_1直
        c = ctx(at(7, 0, 30), cells={(1, 1): "2026.10.29_1直"})
        c.facility = "L-1"
        self.assertFalse(workers.daily_create(c).recreate)
        quirks.set_for_test({"B70": "fix"})
        c = ctx(at(7, 0, 30))
        c.facility = "L-1"
        w = workers.daily_report(1, c)
        self.assertIn(CellWrite(1, 1, "2026.10.29_3直"), w)                          # 日付・直とも hh:mm
        self.assertIn(CellWrite(1, 3, "鈴木"), w)                                     # C1(T3)と揃う
        c = ctx(at(7, 0, 30), cells={(1, 1): "2026.10.29_3直"})
        c.facility = "L-1"
        self.assertFalse(workers.daily_create(c).recreate)
        c = ctx(at(7, 1, 0))
        c.facility = "L-1"
        self.assertIn(CellWrite(1, 1, "2026.10.30_1直"), workers.daily_report(1, c))

    def test_daily_create(self) -> None:
        quirks.set_for_test({})
        c = ctx(at(10, 15), cells={(1, 1): "2026.10.30_1直", (2, 4): "梱包No", (3, 1): "N456J52", (3, 4): "01"})
        r = workers.daily_create(c)
        self.assertFalse(r.recreate)
        self.assertIn(CellWrite(4, 6, "10"), r.writes)      # 開始時刻は次の行
        c = ctx(at(16, 0), cells={(1, 1): "2026.10.30_1直"})
        r = workers.daily_create(c)
        self.assertTrue(r.recreate)
        self.assertEqual(r.writes[0], CellWrite(2, 1, "ﾛｯﾄNo"))
        self.assertEqual(r.writes[15], CellWrite(2, 16, " 停止①記号"))
        self.assertIn(CellWrite(1, 1, "2026.10.30_2直"), r.writes)

    def test_daily_sheet_name(self) -> None:
        self.assertEqual(workers.daily_sheet_name("UF1", "Test"), "Daily_Test")
        self.assertEqual(workers.daily_sheet_name("キャスク", "Test"), "Daily_")


class NameListTest(unittest.TestCase):
    def tearDown(self) -> None:
        quirks.set_for_test(None)

    def test_write_back_b22(self) -> None:
        quirks.set_for_test({})
        wb = workers.name_list_write_back([["L1", "x"], ["L2"]], "L-1", 3)
        self.assertEqual(wb.sheet, "Daily_Test")
        self.assertEqual(wb.clear, (3, 1, 22, 3))
        self.assertEqual(wb.writes[:3], [CellWrite(3, 1, "L1"), CellWrite(3, 2, "x"), CellWrite(3, 3, "")])
        quirks.set_for_test({"B22": "fix"})
        self.assertEqual(workers.name_list_write_back([], "L-1", 3).sheet, "Daily_L-1")
        self.assertEqual(workers.list_posting([["a"]], "L-1", 2).sheet, "Daily_L-1")    # 日報修正は常に自ラインに


if __name__ == "__main__":
    unittest.main()
