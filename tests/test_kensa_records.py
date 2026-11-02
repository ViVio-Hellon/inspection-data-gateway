"""保存・測定DB 登録・作業中保存・履歴(records.py)。期待値は VBA を手で追った結果(行番号はコメント)"""
from __future__ import annotations

import shutil
import sqlite3
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from ufkensa.kensa import quirks, records as rc, store as st, vbx
from ufkensa.kensa.flow import Stop


def make_snap(**kw) -> rc.KensaSnapshot:
    s = rc.KensaSnapshot(line_name="Test", lot="H7372A0", atu="50.000", hab="1528.0", tak="3053.0", am="0.500",
                         ap="0.500", hm="3.0", hp="3.0", tm="3.0", tp="3.0", amk="49.500", apk="50.500",
                         hmk="1525.0", hpk="1531.0", tmk="3050.0", tpk="3056.0", zai="F52S", tyu="5",
                         youto="H207", kind="厚板", atsuita_caption="H", t1="佐藤_検査員01", worker_shift="1",
                         start_time="2024/11/02 05:29:10", end_time="2024/11/02 08:55:00", resize_tag="無")
    for f in "SUCDE":
        s.measure[f] = ["49.776", "49.807", "49.896", "49.793", "1528.6", "1528.5", "3053.9", "3054.0", "3414.8",
                        "3414.3", "", "", "", ""]
        s.strain[f] = rc.Strain("丈", "上", "0.20", "", "", "0.00")
        s.visual_up[f] = "ｶﾐｺﾐ跡"
        s.visual_under[f] = "ｽﾘｷｽﾞ"
        s.lateral_face[f] = "1.5"
        s.positions[f] = f
    for k, v in kw.items():
        setattr(s, k, v)
    return s


class Base(unittest.TestCase):
    def setUp(self) -> None:
        self.dir = Path(tempfile.mkdtemp())
        self.store = st.LocalStore(self.dir / "local.sqlite3")
        self.db = rc.ensure_measure_db(self.dir / "梱包検査ﾃﾞｰﾀ.sqlite3")
        quirks.set_for_test(None)

    def tearDown(self) -> None:
        quirks.set_for_test(None)
        self.store.close()
        shutil.rmtree(self.dir, ignore_errors=True)

    def rows(self, table: str) -> list[tuple]:
        with sqlite3.connect(self.db) as c:
            return c.execute(f'SELECT * FROM "{table}" ORDER BY rowid').fetchall()


class VisualJoinTest(unittest.TestCase):
    def tearDown(self) -> None:
        quirks.set_for_test(None)

    def test_join(self) -> None:
        self.assertEqual(rc.visual_save_join("a", "b", face="S"), "A_a,B_b")
        self.assertEqual(rc.visual_save_join("a", "", face="S"), "A_a,")
        self.assertEqual(rc.visual_save_join("", "b", face="S"), " ,B_b")
        self.assertEqual(rc.visual_save_join("", "", face="S"), "")

    def test_b17_c_face(self) -> None:
        self.assertEqual(rc.visual_save_join("", "b", face="C"), ",B_b")          # SaveLoad:4268
        quirks.set_for_test({"B17": "fix"})
        self.assertEqual(rc.visual_save_join("", "b", face="C"), " ,B_b")

    def test_b17_f_face(self) -> None:
        # SaveLoad:4290: F の「A のみ」を E 面で判定する
        self.assertEqual(rc.visual_save_join("a", "", face="F", a_e="", b_e=""), "")
        self.assertEqual(rc.visual_save_join("a", "b", face="F", a_e="", b_e=""), "A_a,B_b")
        self.assertEqual(rc.visual_save_join("", "", face="F", a_e="e", b_e=""), "A_,")
        quirks.set_for_test({"B17": "fix"})
        self.assertEqual(rc.visual_save_join("a", "", face="F", a_e="", b_e=""), "A_a,")

    def test_strain_ab(self) -> None:
        s = make_snap()
        s.strain_tag["S"] = rc.Strain("幅", "下", "0.30", "", "", "")
        ab = rc.strain_ab(s)
        self.assertEqual(ab[0], ["A:丈_上_0.20,B:幅_下_0.30", "A:__0.00"])   # Tag のゲージは "__" → B は空
        self.assertEqual(ab[1], ["A:丈_上_0.20", "A:__0.00"])                     # B が "__" → ""
        s.aorb = "Now_B"
        self.assertEqual(rc.strain_ab(s)[0][0], "A:幅_下_0.30,B:丈_上_0.20")
        s.aorb = ""
        with self.assertRaises(rc.VbaRuntimeError):
            rc.strain_ab(s)


class MainDataSaveTest(Base):
    def test_hozon_row(self) -> None:
        s = make_snap(lgs_caption="ﾌﾘｰｿﾘ")
        s.diag_kou["S"] = "0.3"
        s.measure["F"] = ["1"] * 10 + [""] * 4
        s.measure["U"][4] = ""
        r = rc.main_data_save(self.store, s, db_path=self.db)
        sh = self.store.sheet(st.HOZON)
        self.assertEqual((r.sheet, r.row), ("保存用", 2))
        self.assertEqual(sh.get(2, 1), "2024/11/02 05:29～2024/11/02 08:55")
        self.assertEqual(sh.get(2, 3), "50.000")              # Array(.ATU, …) は文字列のまま
        self.assertEqual(sh.get(2, 12), "49.776")             # Split の String 配列も文字列のまま
        self.assertIsNone(sh.get(2, 28))                      # U巾1 が空 → 空セル
        self.assertEqual(sh.get(2, 22), "A:丈_上_0.20")       # 歪は A:/B: で上書き
        self.assertEqual(sh.get(2, 23), "A:__0.00")
        self.assertEqual(sh.get(2, 72), 1.5)                  # ラテラルボー(Variant)は数
        self.assertEqual(sh.get(2, 77), "検査員01")           # T1 の末尾
        self.assertEqual(sh.get(2, 78), "無")
        self.assertEqual(sh.get(2, 83), "H")                  # 厚板 → 厚板.Caption
        self.assertEqual(sh.get(2, 84), "〆")                 # LGS = ﾌﾘｰｿﾘ
        self.assertEqual(sh.get(2, 85), "A_ｶﾐｺﾐ跡,B_ｽﾘｷｽﾞ")
        self.assertIsNone(sh.get(2, 90))                      # B17: BH は書かない
        self.assertEqual(sh.get(2, 95), 0.3)
        self.assertEqual(sh.get(2, 100), 1.0)                 # F 面(Variant 配列)
        self.assertIsNone(sh.get(2, 110))                     # F の歪 "__" → 空
        self.assertEqual(r.values[0], "Test")
        self.assertEqual(len(r.values), 115)
        self.assertEqual(r.db.mode, "ADDNEW")
        row = self.rows("検査ﾃﾞｰﾀ")[0]
        self.assertEqual(row[0], "Test")
        self.assertEqual(row[1], "2024/11/02 05:29～2024/11/02 08:55")
        self.assertEqual(row[73], "1.5")                      # CStr(1.5)
        self.assertEqual(row[28], "")                         # U巾1(空)は ""

    def test_kind_frequency(self) -> None:
        for kind, expect in (("薄板", "L:3"), ("少枚数", "H:少枚数"), ("スラブ", "ｽﾗﾌﾞ"), ("", "ｽﾗﾌﾞ")):
            s = make_snap(kind=kind, usuita_caption="L", yama="3")
            r = rc.main_data_save(self.store, s, register=False)
            self.assertEqual(self.store.sheet(st.HOZON).get(r.row, 83), expect)

    def test_b17_fix_writes_bh(self) -> None:
        quirks.set_for_test({"B17": "fix"})
        s = make_snap()
        s.bh["S"] = "65"
        r = rc.main_data_save(self.store, s, register=False)
        self.assertEqual(self.store.sheet(st.HOZON).get(r.row, 90), "65")

    def test_b20_key_with_space(self) -> None:
        s = make_snap()
        rc.main_data_save(self.store, s, db_path=self.db)
        r2 = rc.main_data_save(self.store, s, db_path=self.db)
        self.assertEqual(r2.db.mode, "ADDNEW")                 # 'キー ' は一致しない → 毎回追加(B20)
        self.assertEqual(len(self.rows("検査ﾃﾞｰﾀ")), 2)

    def test_b20_fix_updates(self) -> None:
        quirks.set_for_test({"B20": "fix"})
        s = make_snap()
        rc.main_data_save(self.store, s, db_path=self.db)
        s2 = make_snap(line_name="LVC", zai="52S")
        r2 = rc.main_data_save(self.store, s2, db_path=self.db)
        self.assertEqual(r2.db.mode, "UPDATE")
        rows = self.rows("検査ﾃﾞｰﾀ")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][0], "Test")                   # 1 列目(ライン名)は UPDATE しない
        self.assertEqual(rows[0][79], "52S")                   # 材質(CA = 79 列目 → 位置 79)

    def test_mimi(self) -> None:
        s = make_snap(kind="極厚", data_loading="Load")
        for f in "SUCDE":
            s.measure[f] = ["351.89"] * 4 + ["1283.1"] * 6 + ["2009.1"] * 4
        r = rc.main_data_save(self.store, s, db_path=self.db)
        sh = self.store.sheet(st.MIMI)
        self.assertEqual(r.sheet, "耳用")
        self.assertEqual(sh.get(2, 26), "丈_上_0.20")         # 耳用の歪は A:/B: を付けない
        self.assertEqual(sh.get(2, 97), "〆")
        self.assertEqual(sh.get(2, 98), "耳付き")
        self.assertEqual(sh.get(2, 104), "")                  # BH(コントロール)は長さ 0 の文字列
        self.assertEqual(len(r.values), 109)
        self.assertEqual(len(self.rows("耳付検査ﾃﾞｰﾀ")), 1)

    def test_worker_errors(self) -> None:
        with self.assertRaises(rc.VbaRuntimeError) as cm:
            rc.main_data_save(self.store, make_snap(worker_shift=""), register=False)
        self.assertIn("'13'", cm.exception.text)
        with self.assertRaises(rc.VbaRuntimeError) as cm:
            rc.main_data_save(self.store, make_snap(t1=""), register=False)
        self.assertIn("'9'", cm.exception.text)


class RegisterTest(Base):
    def test_more_columns_than_values(self) -> None:
        r = rc.register_measurement(self.db, "検査ﾃﾞｰﾀ", ["Test", "k"], "[検査開始終了] ='k'")
        self.assertEqual(r.mode, "ERROR")
        self.assertEqual(r.messages[0].text, "エラー番号  :9\r\nエラーの内容:インデックスが有効範囲にありません。")
        self.assertEqual(self.rows("検査ﾃﾞｰﾀ"), [])

    def test_case_insensitive_key_and_quote(self) -> None:
        vals = ["L", "AbC"] + [""] * 113
        self.assertEqual(rc.register_measurement(self.db, "検査ﾃﾞｰﾀ", vals, "[検査開始終了] ='AbC'").mode, "ADDNEW")
        vals2 = ["L", "abc", "x'y"] + [""] * 112
        r = rc.register_measurement(self.db, "検査ﾃﾞｰﾀ", vals2, "[検査開始終了] ='abc'")
        self.assertEqual(r.mode, "ERROR")                      # UPDATE の SQL に ' を埋め込んで構文エラー
        vals3 = ["L", "abc", "z"] + [""] * 112
        self.assertEqual(rc.register_measurement(self.db, "検査ﾃﾞｰﾀ", vals3, "[検査開始終了] ='abc'").mode, "UPDATE")
        self.assertEqual(self.rows("検査ﾃﾞｰﾀ")[0][1:3], ("abc", "z"))

    def test_missing_db(self) -> None:
        with self.assertRaises(Stop):
            rc.register_measurement(self.dir / "無い.sqlite3", "検査ﾃﾞｰﾀ", ["x"], "[検査開始終了] ='x'")

    def test_lock_retry(self) -> None:
        other = sqlite3.connect(self.db, isolation_level=None)
        other.execute("BEGIN IMMEDIATE")
        old = rc.MEASURE_BUSY_TIMEOUT_MS
        rc.MEASURE_BUSY_TIMEOUT_MS = 50
        try:
            r = rc.register_measurement(self.db, "検査ﾃﾞｰﾀ", ["L", "k"] + [""] * 113, "[検査開始終了] ='k'",
                                        max_retry=1, retry_wait_sec=0.01)
        finally:
            rc.MEASURE_BUSY_TIMEOUT_MS = old
            other.execute("ROLLBACK")
            other.close()
        self.assertEqual(r.mode, "ERROR")
        self.assertIn("locked", r.messages[0].text)


class PostingLoadTest(Base):
    def snap(self) -> rc.KensaSnapshot:
        s = make_snap(atu="25.000", kind="薄板", usuita_caption="L", yama="3", at_butt_visible=True,
                      at_count="厚2ヵ所検査", ken_no="T6X1913", ken_hbox=["200", "226", "", "", "", "", ""],
                      yama_lb="2", delete_p="C2", delete_c="1", sumisei="", additional="対角追記公差あり",
                      taik_text="3", huris="1.2", lateral_text_tag="有")
        s.diag["S"] = "0.0"
        s.strain_tag["S"] = rc.Strain("幅", "下", "0.30", "", "", "")
        s.measure["U"] = [""] * 14
        return s

    def test_posting_cells(self) -> None:
        sh = rc.posting(self.store, self.snap(), need_right=False)
        self.assertEqual(sh.get(1, 2), 25.0)                    # Variant 配列 → 数
        self.assertEqual(sh.get(1, 29), "薄板")
        self.assertEqual(sh.get(1, 30), "厚2ヵ所")              # Mid(AtCount, 1, 4)
        self.assertEqual(sh.get(1, 32), 3.0)
        self.assertEqual(sh.get(2, 17), "0.0")                  # String 配列 → '0.0' のまま(作業保存_Test!Q2)
        self.assertEqual(sh.get(2, 15), "丈_上_0.20")           # A 面(表示中)
        self.assertEqual(sh.get(8, 15), "幅_下_0.30")           # B 面(Tag)
        self.assertEqual(sh.get(9, 3), 200.0)
        self.assertEqual(sh.get(9, 2), 2.0)
        self.assertIsNone(sh.get(2, 18))                        # 直角度が不要 → R 列を消す

    def test_load_text_round_trip(self) -> None:
        rc.posting(self.store, self.snap(), need_right=True)
        out = rc.load_text(self.store, "Test")
        v = out.values
        self.assertEqual(v["ATU"], "25.000")
        self.assertEqual(v["Hab"], "1528.0")
        self.assertEqual(v["S1"], "49.776")
        self.assertEqual(v["U1"], "")                            # 0 は空に
        self.assertEqual(out.back_colors["U1"], rc.RGB_WHITE_DEFAULT)
        self.assertEqual(v["S7"], "3053.9")
        self.assertEqual(v["DiagS"], "0.0")                      # S10 があるので 0 でも残す(SaveLoad:2831)
        self.assertEqual(v["KenHBox1"], "200.0")
        self.assertEqual(v["KenHBox3"], "")
        self.assertEqual(out.list_index["YamaLB"], 1)
        self.assertEqual(v["HURIS"], "1.2")
        self.assertEqual(out.tags["LateralText"], "有")
        self.assertEqual(out.tags["ReSize"], "無")
        self.assertEqual(out.back_colors["TaikText"], 51 + 204 * 256 + 204 * 65536)
        self.assertEqual(v["CBS1"], "丈")                        # A 面を表示
        self.assertEqual(out.tags["CBS1"], "幅")                 # B 面は Tag
        self.assertEqual(out.option, "薄板")
        self.assertEqual(v["Yama"], "3")
        self.assertIn(("LSelectLateral", "T"), out.calls)       # Tag="有"
        self.assertEqual(v["KenNo"], "T6X1913")
        self.assertIsNone(rc.load_text(self.store, "LVC"))

    def test_reload(self) -> None:
        rc.posting(self.store, self.snap(), need_right=None)
        out = rc.load_text(self.store, "Test", "ReLoad")
        self.assertEqual(out.tags["ReSize"], "無")


class HistoryTest(Base):
    def test_history_load(self) -> None:
        s = make_snap()
        s.strain_tag["S"] = rc.Strain("幅", "下", "0.30", "", "", "")
        rc.main_data_save(self.store, s, register=False)
        out = rc.history_load(self.store, "H7372A0", lname="Test", current={"F1": ""})
        self.assertTrue(out.found)
        v = out.values
        self.assertEqual(v["ATU"], "50.000")
        self.assertEqual(v["ApK"], "50.500")
        self.assertEqual(v["S1"], "49.776")
        self.assertEqual(v["S10"], "3414.3")
        self.assertEqual(v["DiagS"], "0.5")                       # |3414.8 - 3414.3|
        self.assertEqual(v["LateralS"], "1.5")
        self.assertEqual(v["S_UpVisual"], "ｶﾐｺﾐ跡")
        self.assertEqual(v["S_UnderVisual"], "ｽﾘｷｽﾞ")
        self.assertEqual(v["CBS1"], "丈")
        self.assertEqual(out.tags["CBS1"], "幅")
        self.assertEqual(v["F1"], "")
        self.assertEqual(self.store.sheet("WorkerTime_Test").get(1, 1), "履歴呼び出し")
        self.assertEqual(out.messages[0].text, "Load完了,印刷する場合はフォームを閉じないでください。")
        self.assertFalse(rc.history_load(self.store, "無い").found)

    def test_kanri_view(self) -> None:
        rc.main_data_save(self.store, make_snap(), register=False)
        rc.main_data_save(self.store, make_snap(start_time="2024/11/03 05:00:00"), register=False)
        lots = rc.kanri_lots(self.store)
        self.assertEqual(lots, ["H7372A0", "H7372A0"])
        view = rc.kanri_view(self.store, 0, lots, "H7372A0")
        self.assertEqual(view.messages[0].text, "同LotNoあり、作業時間を選択してください")
        self.assertEqual(view.new_cb, ["2024/11/02 05:29～2024/11/02 08:55", "2024/11/03 05:00～2024/11/02 08:55"])
        self.assertEqual(view.values["Worker"], "検査員01")
        self.assertEqual(view.values["SZH"], "A:丈_上_0.20")    # "," が無ければ A: のまま(UF管理:236)
        self.assertEqual(view.values["Count"], "作業数2件")
        self.assertEqual(rc.kanri_select_time(self.store, "2024/11/03 05:00～2024/11/02 08:55", "H7372A0"), 1)

    def test_kanri_b27_fgh(self) -> None:
        sh = self.store.sheet(st.HOZON)
        sh.write(2, 2, "L1")
        sh.write(2, 110, "__")
        sh.write(2, 111, "A:x_y_z")
        self.assertEqual(rc.kanri_view(self.store, 0, ["L1"], "L1").values["FGH"], "")   # 110 列を見る
        quirks.set_for_test({"B27": "fix"})
        self.assertEqual(rc.kanri_view(self.store, 0, ["L1"], "L1").values["FGH"], "A:x_y_z")

    def test_kanri_load_other_lot(self) -> None:
        out = rc.kanri_load(self.store, "L1", "", uf1_visible=True, uf1_lot="X")
        self.assertEqual(out.messages[0].text, "別Lotデータがあります、新規登録処理を行ってください")


class NumberCountTest(Base):
    def daily(self, rows) -> None:
        sh = self.store.add_sheet("Daily_Test")
        for c, h in enumerate(["ﾛｯﾄNo", "材・調質", "製品寸法", "梱包No", "検入枚数"], start=1):
            sh.write(2, c, h)
        for r, row in enumerate(rows, start=3):
            for c, v in row.items():
                sh.write(r, c, v)

    def test_over(self) -> None:
        self.daily([{1: "N1", 5: "100", 12: "30", 13: "2"}, {12: "25", 13: "2"}, {1: "N2", 5: "10", 12: "1", 13: "1"}])
        res = rc.number_count(self.store, "Test", "N1")
        self.assertTrue(res.stop)                                # 30×2 + 25×2 = 110 > 100
        self.assertEqual(res.red_row, 3)
        self.assertEqual(res.packed["N1"], 110)
        self.assertEqual(self.store.sheet("Daily_Test").color(3, 5), rc.RGB_RED_NG)
        self.assertEqual(res.messages[0].text, "梱包数と検入枚数が一致しません")
        ws = self.store.sheet("梱包枚数計算_Test")
        self.assertEqual([ws.get(1, c) for c in range(1, 5)], ["LotNo", "梱包枚数", "LotNo", "検入枚数"])

    def test_equal_is_ok(self) -> None:
        self.daily([{1: "N1", 5: "100", 12: "50", 13: "2"}])
        self.assertFalse(rc.number_count(self.store, "Test", "N1").stop)    # > でなければ通す

    def test_no_rows_reads_header(self) -> None:
        # Daily の 3 行目以降が空: Range(A3:M2) は A2:M3 になり、見出しを数える(梱包枚数計算_Test の実物どおり)
        self.daily([])
        res = rc.number_count(self.store, "Test", "N1")
        ws = self.store.sheet("梱包枚数計算_Test")
        self.assertEqual((ws.get(2, 1), ws.get(2, 2), ws.get(2, 3), ws.get(2, 4)), ("ﾛｯﾄNo", 0, "ﾛｯﾄNo", "検入枚数"))
        self.assertFalse(res.stop)

    def test_no_daily(self) -> None:
        self.assertIsNone(rc.number_count(self.store, "Test", "N1"))


class EndTimeTest(Base):
    def test_boundaries(self) -> None:
        self.store.add_sheet("WorkerTime_Test")
        self.assertEqual(rc.end_time_register(self.store, "Test", datetime(2026, 10, 30, 7, 0, 0)),
                         "2026/10/29 07:00:00")                  # 7:00 ちょうどは前日(Worker:319)
        self.assertEqual(rc.end_time_register(self.store, "Test", datetime(2026, 10, 30, 7, 0, 1)),
                         "2026/10/30 7:00:01")
        self.assertEqual(rc.end_time_register(self.store, "Test", datetime(2026, 10, 30, 3, 5, 9)),
                         "2026/10/29 03:05:09")
        self.assertEqual(rc.end_time_register(self.store, "Test", datetime(2026, 10, 30, 23, 59, 30)),
                         "2026/10/30 23:59:30")                  # 23:59 より後は Else(今)
        self.assertEqual(self.store.sheet("WorkerTime_Test").get(1, 3), "2026/10/30 23:59:30")


class CsvTest(Base):
    def test_exp_csv(self) -> None:
        sh = self.store.add_sheet("A_Icas_Test")
        sh.write(1, 1, "a,b", "S")
        sh.write(1, 2, "1.50")
        sh.write(2, 1, "①㉑")
        p = rc.exp_csv(sh, "x", self.dir)
        self.assertEqual(p.read_bytes(), "a,b,1.5,\r\n①?,,\r\n".encode("cp932"))
        p2 = rc.exp_csv(sh, "y", self.dir, sep="\t")
        self.assertEqual(p2.suffix, ".txt")

    def test_hostsend(self) -> None:
        sh = self.store.add_sheet("A_Icas_Test")
        sh.write(1, 1, "ABC", "S")
        self.assertEqual(rc.exp_csv(sh, "hostsend", self.dir, sep="\t").read_bytes(), b"ABC\r\n")
        sh.write(1, 2, "D", "S")
        with self.assertRaises(rc.VbaRuntimeError):
            rc.exp_csv(sh, "hostsend", self.dir, sep="\t")


class DateTest(unittest.TestCase):
    def test_format(self) -> None:
        self.assertEqual(rc.vb_format_date("2024/11/02 05:29:10", "yyyy/mm/dd  h:mm"), "2024/11/02  5:29")
        self.assertEqual(rc.vb_format_date("2024/11/2 5:09", "yyyy/mm/dd hh:mm"), "2024/11/02 05:09")
        self.assertEqual(rc.vb_format_date("", "yyyy"), "")
        self.assertEqual(rc.vb_format_date("abc", "yyyy"), "abc")
        self.assertEqual(rc.vb_cstr_date(datetime(2026, 10, 30, 9, 5, 3)), "2026/10/30 9:05:03")
        self.assertEqual(rc.vb_cstr_date(datetime(2026, 10, 30)), "2026/10/30")
        self.assertFalse(rc.vb_isdate("2026.06.22"))              # U12 の既定
        self.assertTrue(rc.vb_isdate("2026.06.22", dot_is_date=True))

    def test_variant_eq(self) -> None:
        self.assertTrue(rc.variant_eq(None, ""))
        self.assertTrue(rc.variant_eq(None, 0.0))
        self.assertFalse(rc.variant_eq(50.0, "50"))               # Variant どうし: 数と文字列は等しくない
        self.assertTrue(rc.variant_eq(50.0, "50", vbx.STR))       # String と Variant: CStr して比べる
        self.assertEqual(rc.cmp_str_num("0.000", 0), 0)           # Caption と 0: 数で比べる
        with self.assertRaises(vbx.TypeMismatch):
            rc.cmp_str_num("", 4)
        self.assertTrue(rc.variant_eq("a", "a"))


if __name__ == "__main__":
    unittest.main()
