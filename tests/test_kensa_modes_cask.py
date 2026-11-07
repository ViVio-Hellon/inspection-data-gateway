"""キャスク検査(cask.py): シート入力・ゼロ校正・対角・ロット読込(U4)・梱包板番号 DB・最大最小(B28)・保存(B29)"""
from __future__ import annotations

import sqlite3
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from tests import kensa_fixture
from ufkensa.kensa import datasource, flow, quirks, vbx
from ufkensa.kensa.modes import cask
from ufkensa.kensa.modes import sheet as sh

NOW = datetime(2026, 10, 30, 10, 0)


def cask_book(*, micro: str = "0.000", convex: str = "0.0", count_st: str = "3") -> sh.Book:
    b = sh.Book(now=lambda: NOW)
    ws = b.ws(sh.CASK)
    ws.visible = True
    b.activate_sheet(sh.CASK)
    f = b.cask
    f.visible, f.line_name, f.micro, f.convex, f.count_st = True, "LVC", micro, convex, count_st
    for (r, c), v in {(2, 4): "H3084S0", (3, 2): 36.0, (3, 4): 1528.0, (3, 6): 3053.0, (2, 13): 0.4, (2, 15): 0.4,
                      (3, 13): 3.0, (3, 15): 3.0, (4, 13): 3.0, (4, 15): 3.0}.items():
        ws.set(r, c, v)
    for k, fm in sh.TOLERANCE_FORMULAS.items():
        ws.set_formula(*k, fm)
    return b


class CaskSheetTest(unittest.TestCase):
    def tearDown(self) -> None:
        pass
    def test_zero_calibration_required(self) -> None:   # Cask:1039-1044
        b = cask_book(micro="")
        fx = b.user_edit(7, 2, "DT10002+0000036.000M")
        self.assertEqual(b.ws(sh.CASK).get(7, 2), 36.0)
        self.assertIsNone(b.ws(sh.CASK).get(7, 3))
        self.assertIn("ゼロ校正を実行してください！！", [m.text for m in fx.messages])

    def test_thickness_judge_and_cursor(self) -> None:
        b = cask_book()
        ws = b.ws(sh.CASK)
        fx = b.user_edit(7, 2, "DT10002+0000035.600M")
        self.assertEqual(ws.get(7, 3), "○")                 # 下限ちょうど
        self.assertEqual(ws.interior[(7, 3)], (0, 128, 0))
        self.assertEqual(b.active_cell, (7, 4))             # 判定を書いた入れ子の Change で次の入力セルへ
        self.assertEqual(ws.get(1, 2), datetime(2026, 10, 30, 10, 0))   # D7 を選んだので開始日時
        self.assertIn(sh.FormShow("UFcount", "Maru", 7, 2), fx.forms)
        self.assertEqual(fx.logs[1].path.split("\\")[-1], "LVC_Log")
        b.user_edit(7, 4, "DT10002+0000035.599M")
        self.assertEqual(ws.get(7, 5), "×")
        # 決定 D14(W18): 3 点のとき厚3 の後は、巾丈測定 OFF なら次の行の厚1(VBA は保護した H 列へ移っていた)
        b.cask.inspection = "OFF"
        b.user_edit(7, 6, "DT10002+0000036.000M")
        self.assertEqual(b.active_cell, (8, 2))
        # 巾丈測定 ON なら幅1
        b.cask.inspection = "ON"
        b.user_edit(8, 6, "DT10002+0000036.000M")
        self.assertEqual(b.active_cell, (8, 11))

    def test_four_points_after_thickness4(self) -> None:
        b = cask_book(count_st="4")
        b.cask.inspection = "OFF"
        b.user_edit(7, 8, "DT10002+0000036.000M")
        self.assertEqual(b.active_cell, (8, 2))

    def test_manual_thickness_is_erased(self) -> None:   # Cask:63-66 厚欄の手入力は消す
        b = cask_book()
        b.user_edit(7, 2, "36.0")
        self.assertIsNone(b.ws(sh.CASK).get(7, 2))
        self.assertIsNone(b.ws(sh.CASK).get(7, 3))
        self.assertTrue(b.events_enabled)                 # 空は IsNumeric が真なので最後まで進み EnableEvents は戻る

    def test_l_packet_needs_inspection_on(self) -> None:  # Cask:1774-1778
        b = cask_book()
        fx = b.user_edit(7, 2, "#L1528.0")
        self.assertIn("巾丈測定がONになっていません", [m.text for m in fx.messages])

    def test_dt_in_width_moves_to_thickness(self) -> None:   # Special_Move_DT Cask:1611-1636
        b = cask_book()
        b.cask.inspection = "ON"
        ws = b.ws(sh.CASK)
        b.user_select(7, 11)
        b.user_edit(7, 11, "DT10002+0000036.100M")
        self.assertEqual(ws.get(7, 2), 36.1)
        self.assertEqual(ws.get(7, 3), "○")
        self.assertIsNone(ws.get(7, 11))

    def test_diagonal(self) -> None:                      # Cask:1206-1221, 1499-1535
        b = cask_book()
        b.cask.inspection = "ON"
        ws = b.ws(sh.CASK)
        b.user_edit(7, 19, "#L2000.0")
        fx = b.user_edit(7, 20, "#L2002.1")
        self.assertIn("対角公差が入力されていません！！", [m.text for m in fx.messages])
        b.cask.tai = "2"
        b.user_edit(7, 20, "#L2002.1")
        self.assertEqual(ws.get(7, 21), "×")
        self.assertEqual(ws.get(7, 22), 2.1)
        self.assertEqual(ws.get(7, 23), "公差:[2.0]")
        b.cask.tai = "2.1"
        b.user_edit(7, 19, "#L1530.4")
        b.user_edit(7, 20, "#L1528.3")
        self.assertEqual(ws.get(7, 21), "○")              # Single で比べるので ○

    def test_loop_judge(self) -> None:                    # Cask:979-1027
        b = cask_book()
        ws = b.ws(sh.CASK)
        ws.set(8, 2, 36.5)
        ws.set(8, 11, 1528.0)
        cask.loop_judge(b)
        self.assertEqual(ws.get(8, 3), "×")
        self.assertEqual(ws.interior[(8, 3)], (200, 60, 50))
        self.assertEqual(ws.get(8, 12), "○")
        with self.assertRaises(sh.VbaError):
            cask.loop_judge(b, u6_raises=True)


class CalibrationTest(unittest.TestCase):
    def test_start_tb(self) -> None:                      # UFCask:736-776 / Cask:1956-1962
        b = cask_book(micro="", convex="")
        f = b.cask
        cask.start_tb_after_update(b, "DT10002+0000000.000M")
        self.assertEqual((f.micro, f.l1), ("0.000", "ﾏｲｸﾛｾﾞﾛ点"))
        fx = cask.start_tb_after_update(b, "DT10002+0000000.001M")
        self.assertEqual(f.micro, "")
        self.assertIn("マイクロ校正を正しく実行してください！！", [m.text for m in fx.messages])
        cask.start_tb_after_update(b, "DT10002+0000000.0004M")   # "0.000" に丸めてから見る → OK
        self.assertEqual(f.micro, "0.000")
        cask.start_tb_after_update(b, "#L0.5")
        self.assertEqual(f.convex, "0.5")
        cask.start_tb_after_update(b, "#L-0.5")
        self.assertEqual(f.convex, "-0.5")
        cask.start_tb_after_update(b, "#L0.54")              # "0.5" に丸めてから見る → OK
        self.assertEqual(f.convex, "0.5")
        cask.start_tb_after_update(b, "#L0.6")
        self.assertEqual(f.convex, "")
        b.ws(sh.CASK).set(3, 2, 120.0)
        cask.start_tb_after_update(b, "DT10002+0000000.000M")
        self.assertEqual((f.micro, f.l1), ("0.00", "ﾉｷﾞｽｾﾞﾛ点"))
        cask.start_tb_after_update(b, "#L")
        self.assertTrue(f.l2_visible)

    def test_tolerance_pair(self) -> None:                # UFCask:470
        self.assertEqual(cask.tolerance_pair(0.3, 0.3, 0.2, 0.2), (0.3, 0.3))
        self.assertEqual(cask.tolerance_pair(0, 0.3, 0.2, 0.2), (0.2, 0.2))
        self.assertEqual(cask.tolerance_pair(0.3, 0, 0.2, 0.2), (0.2, 0.2))


class LotReadTest(unittest.TestCase):
    def setUp(self) -> None:
        kensa_fixture.use_offline()

    def test_u4_fullwidth_field_error(self) -> None:     # UFCask:436 全角の欄名で読めない場合
        b = cask_book()
        b.cask.combo = "RX021T0"
        fx = cask.combo_change(b, fullwidth_fields_readable=False)
        self.assertEqual(fx.messages[-1].text, "DB読み込みエラー:" + "\r\n" + cask.FIELD_MISSING)
        self.assertEqual(b.ws(sh.CASK).get(3, 2), 36.0)      # シートは変わらない
        self.assertTrue(b.events_enabled)

    def test_u4_readable(self) -> None:
        b = cask_book()
        b.cask.combo = "RX021T0"
        fx = cask.combo_change(b, fullwidth_fields_readable=True)
        ws = b.ws(sh.CASK)
        self.assertEqual(ws.get(2, 4), "RX021T0")
        self.assertEqual(ws.get(3, 2), 25.0)
        self.assertEqual(b.cask.count_st, "3")
        self.assertEqual((ws.get(2, 13), ws.get(2, 15)), (0.45, 0.45))   # オーダー公差が空 → 製造公差
        self.assertEqual(sh.val_cell(ws.get(2, 17)), 24.55)
        self.assertEqual(b.cask.ken_no, "T6X1913")
        self.assertEqual(fx.calls[0].name, "HistorySave")

    def test_lot_not_found(self) -> None:
        b = cask_book()
        b.cask.combo = "ZZZZZZZ"
        fx = cask.combo_change(b, fullwidth_fields_readable=True)
        self.assertIn("Text.データなし　手入力よろしく", [m.text for m in fx.messages])
        self.assertIsNone(b.ws(sh.CASK).get(3, 2))
        self.assertFalse(b.events_enabled)                # VBA どおり False のまま


class PackDbTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = kensa_fixture.copy_to(Path(self.tmp.name))
        kensa_fixture.use_offline(self.dir)

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def rows(self, lot: str) -> list:
        with sqlite3.connect(self.dir / "伝送用ファイル.sqlite3") as c:
            return c.execute('SELECT "スタートNo","エンドNo","管理番号" FROM "キャスク梱包板番号" WHERE "LotNo"=? '
                             'ORDER BY "スタートNo"', (lot,)).fetchall()

    def form(self, start: str, end: str) -> sh.Book:
        b = cask_book()
        b.cask.lot_box, b.cask.ken_no, b.cask.min_text, b.cask.max_text = "H473C51", "H451076", start, end
        return b

    def test_sequence(self) -> None:                      # UFCask:926-1053(既存 1-22,23-29,30-33,34-38,39-41)
        self.assertTrue(cask.check_cask_sequence(self.form("42", "50")))
        b = self.form("43", "50")
        self.assertFalse(cask.check_cask_sequence(b))
        self.assertEqual(b.fx.messages[-1].title, "連番チェック")
        self.assertTrue(cask.check_cask_sequence(self.form("23", "29")))   # 同じスタートNo は置き換え
        b = self.form("30", "20")
        self.assertFalse(cask.check_cask_sequence(b))
        self.assertTrue(b.fx.messages[-1].text.startswith("スタートNo がエンドNo より大きい値です。"))
        self.assertTrue(cask.check_cask_sequence(self.form("1", "5")))
        b = self.form("A", "5")
        self.assertFalse(cask.check_cask_sequence(b))
        self.assertTrue(b.fx.messages[-1].text.startswith("連番チェックエラー："))

    def test_save_insert_update_delete(self) -> None:     # UFCask:1056-1190
        self.assertTrue(cask.save_cask_to_db(self.form("42", "50"), NOW))
        r = self.rows("H473C51")
        self.assertEqual(r[-1], (42, 50, 14))              # 管理番号 = MAX(13)+1
        self.assertTrue(cask.save_cask_to_db(self.form("42", "55"), NOW))
        self.assertEqual(self.rows("H473C51")[-1], (42, 55, 14))
        self.assertTrue(cask.save_cask_to_db(self.form("1", "10"), NOW))
        self.assertEqual(self.rows("H473C51"), [(1, 10, 1)])  # スタートNo=1 なら他は削除
        self.assertEqual(len(self.rows("RX092D0")), 3)        # 他のロットはそのまま

    def test_column_order_d18(self) -> None:
        # 決定 D18: 列名で読む → 梱包範囲は「No1:スタートNo～エンドNo」(VBA は位置で読み、1列ずれていた)
        b = self.form("1", "1")
        cask.add_cask_labels(b, "H473C51", "H451076")
        self.assertEqual(b.cask.labels, ["No1:1～22", "No2:23～29", "No3:30～33", "No4:34～38", "No5:39～41"])
        ws = b.ws(sh.CASK)
        ws.set(47, 2, 36.0)                                  # A47 = 41 = DB の最後のエンドNo
        self.assertTrue(cask.check_cask_end_no(b, "H473C51", "H451076"))
        ws.set(48, 2, 36.0)                                  # A48 = 42 → 合わない
        self.assertFalse(cask.check_cask_end_no(b, "H473C51", "H451076"))
        b2 = self.form("1", "1")
        self.assertFalse(cask.check_cask_end_no(b2, "NOLOT", "X"))
        self.assertTrue(b2.fx.messages[-1].text.startswith("DBにデータが見つかりません。"))


class EventsTest(unittest.TestCase):
    def tearDown(self) -> None:
        pass
class MaxMinTest(unittest.TestCase):
    def tearDown(self) -> None:
        pass
    def book(self, count_st: str) -> sh.Book:
        b = cask_book(count_st=count_st)
        b.cask.min_text, b.cask.max_text = "1", "3"
        ws = b.ws(sh.CASK)
        for r in (7, 8, 9):
            for c in (2, 4, 6, 8):
                ws.set(r, c, 36.5)
        ws.set(7, 8, 36.0)                                   # 最小は H7 だけ
        ws.set(8, 11, 1528.5)
        ws.set(9, 11, 1527.5)
        return b

    def test_b80_empty_column(self) -> None:
        b = self.book("4")
        cask.max_min(b)
        ws = b.ws(sh.CASK)
        self.assertEqual((ws.get(3, 23), ws.get(4, 23)), (1528.5, 1527.5))   # 空の M 列は除く
        self.assertIn((9, 11), ws.underline)
        self.assertNotIn((7, 13), ws.underline)
        self.assertIn(ws.get(3, 24), (None, ""))               # 丈が全部空なら書かない
        self.assertEqual(ws.get(4, 22), 36.0)

class SaveTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        kensa_fixture.use_offline(kensa_fixture.copy_to(Path(self.tmp.name)))

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def test_cask_save_plan(self) -> None:               # Cask:445-682
        b = cask_book()
        f = b.cask
        f.lot_box, f.ken_no, f.kname, f.cno, f.mno, f.seri_no = "H3084S0", "H651019", "検査員01", "C1", "M1", "S9"
        b.ws(sh.CASK).set(7, 2, 36.1)
        plan = cask.cask_save(b)
        save = [s for s in plan.steps if s.kind == "save"][0]
        self.assertTrue(save.path.endswith("2026.10\\20261030_H3084S0\\20261030_H3084S0.xlsx"))
        self.assertEqual(save.password, "H3084S0.2026.10.30")
        copies = [s.path for s in plan.steps if s.kind == "copy"]
        self.assertTrue(copies[0].endswith("キャスク_履歴\\2026.10\\20261030_H3084S0.xlsx"))
        self.assertTrue(copies[1].endswith("【梱包】キャスク\\2026.10\\20261030_H3084S0.xlsx"))
        hin = plan.sheets["品証ｺﾋﾟ-用"]
        self.assertEqual((hin["D1"], hin["E1"], hin["E2"], hin["M1"]), ("板番号", "厚1", 36.1, "対角線差"))
        self.assertEqual(hin["B12"], "梱包No1:1～3")          # 決定 D18
        self.assertNotIn("キャスク保存", b.extra)

    def test_required_fields(self) -> None:
        b = cask_book()
        self.assertIsNone(cask.cask_save(b))
        self.assertEqual(b.fx.messages[-1].text, "フォーム上で必要欄を記入してください")

    def test_print_question_and_setup(self) -> None:     # UFCask:110-174
        b = cask_book()
        f = b.cask
        f.lot_box, f.ken_no, f.kname, f.cno, f.mno, f.seri_no = "H3084S0", "H651019", "検査員01", "C1", "M1", "S9"
        b.ws(sh.CASK).set(22, 2, 36.0)                      # A22 = 16、DB の最後のエンドNo は 15 → 一致しない
        res = cask.print_and_save(b)
        self.assertNotIsInstance(res, flow.Question)
        self.assertTrue(res.messages[-1].text.startswith("エンドNoが一致しません。"))
        self.assertEqual(f.tag, "off")
        b.ws(sh.CASK).set(22, 2, None)                      # A21 = 15 で一致(決定 D18)→ 印刷するか聞く
        b.ws(sh.CASK).set(21, 2, 36.0)
        self.assertIsInstance(cask.print_and_save(b), flow.Question)
        job = cask.print_setup(b)
        self.assertIn("*厚公差：35.600～36.400", job.right_header)
        self.assertTrue(job.left_footer.startswith('&"Meiryo UI"&10梱包No1_1～3'))
        self.assertEqual(job.print_area, "$A$6:$AB$21")


if __name__ == "__main__":
    unittest.main()
