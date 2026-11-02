"""Aインプット(電文・CSV・送信・結果)と梱包高さチェック(Input / UFAinput)"""
from __future__ import annotations

import shutil
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from tests import kensa_fixture
from ufkensa import user_settings
from ufkensa.kensa import ainput, flow, paths, vbx
from ufkensa.kensa.ainput import AinputForm
from ufkensa.kensa.workers import CellWrite, ShiftTimes, WorkerRegistry

NOW = datetime(2026, 10, 30, 10, 7, 0)
SHIFTS = ShiftTimes("07:00", "15:00", "15:00", "22:50", "22:50", "07:00", "08:15", "17:00")
S = vbx.single

# Lot情報履歴(2 行目から)。列 2=ｵｰﾀﾞｰNo 3=ﾛｯﾄNo 17=合紙 18=VC 19=VC_A 20=VC_B
HIST = [
    ["2026-06-03", "57076309", "N456J52", 3, 1000, 2000, "K650969", "995", "H24", "H174", "", "", "", "",
     150, 2439, 1, 0, None, None, 16.26, "60524074", 0, 2000, "", "1P0116"],
    ["2026-06-03", "04077194", "H3416A0", 15, 1528, 3053, "H651336", "F52S", "R", "H279", "", "", "", "",
     15, 2812.9, 0, 1, "208PRWH", "208PRWH", 187.53235, "60602048", 0, 0, "", "1P1202"],
    ["2026-06-03", "00000001", "X", 1, 1, 1, "K", "", "", "", "", "", "", "", 1, 1, 0, 0, None, None, 1, "1"],
]
HEAD = ["登録日", "ｵｰﾀﾞｰNo", "ﾛｯﾄNo", "厚", "巾", "丈", "検査No", "材質", "調質", "用途ｺｰﾄﾞ", "ｺｰｽ", "得意先", "納入先",
        "送り先", "前工程実績枚数", "前工程実績重量", "合紙", "VC", "VC_A", "VC_B", "単重", "引当NO", "引当調整NO",
        "引当数量", "引当ﾌｧｲﾙ登録", "包装仕様NO"]


def vc_stub(code: str):
    return S(0.09), []


def hc(atu1, atu2, hei, mai, order="00000001", form=None, hist=HIST):
    return ainput.height_check(atu1, atu2, hei, mai, order, form or AinputForm(remain="0"), hist, vc=vc_stub)


class TelegramTest(unittest.TestCase):
    def test_zer_con(self) -> None:
        # G2 §10.10
        self.assertEqual(ainput.zer_con(31, "AB"), "AB        ")
        self.assertEqual(ainput.zer_con(25, ""), "     ")
        self.assertEqual(ainput.zer_con(12, "123456"), "23456")            # あふれは左を切る
        self.assertEqual(ainput.zer_con(3, "123"), "123")                  # 引当No は桁を揃えない
        self.assertEqual(ainput.zer_con(3, ""), "00000000")
        self.assertEqual(ainput.zer_con(4, "5"), "05")
        self.assertEqual(ainput.zer_con(13, "2"), "002")
        self.assertEqual(ainput.zer_con(26, "12"), "0012")
        self.assertEqual(ainput.zer_con(40, ""), "000000")
        self.assertEqual(ainput.zer_con(30, ""), " ")
        self.assertEqual(ainput.zer_con(31, "12345678901"), "1234567890")

    def test_st_con(self) -> None:
        self.assertEqual(ainput.st_con(1.005, "厚"), "0001005")              # CStr 15 桁(Python なら 1004.99…)
        self.assertEqual(ainput.st_con("1102.3", "巾"), "0110230")
        self.assertEqual(ainput.st_con("3053.3", "丈"), "0030533")
        self.assertEqual(ainput.st_con("0", "厚"), "")
        self.assertEqual(ainput.st_con("0.2855", "厚"), "00285.5")
        with self.assertRaises(vbx.TypeMismatch):
            ainput.st_con("S", "厚")

    def test_lin_conv(self) -> None:
        self.assertEqual(ainput.lin_conv("丸徳"), "Maru")
        self.assertEqual(ainput.lin_conv("ﾄｯﾄ"), "01")
        self.assertEqual(ainput.lin_conv("Test"), "AA")
        self.assertEqual(ainput.lin_conv(""), "")

    def form(self) -> AinputForm:
        return AinputForm(line_name="Test", uf1_line="Test", lot="N456J52", ken_no="K650969", order_no="57076309",
                          atu="3.000", amin="2.995", amax="3.005", hiki_no="60524074", konno="01", mai1="50",
                          cou1="1", mount_count="1", heights=["152.5"] + [""] * 14)

    def test_telegram(self) -> None:
        t = ainput.telegram(self.form())
        expect = ("KEN:AIN:0185" + "60524074" + "01" + "0" + "0002995" + "0003005" + "00050" + "001" + "00000" + "000"
                  + "00000" * 6 + "0000" * 2 + "00000" + "     " + "0000" * 4 + " " + " " * 10 + "00000" + "000" * 3
                  + "0000" + "00000" * 3 + "000000" + "AA")
        self.assertEqual(t, expect)
        self.assertEqual(vbx.mid(t, 13, 8), "60524074")

    def test_csv_lines(self) -> None:
        rows = ainput.csv_lines(self.form())
        self.assertEqual(len(rows), 56)
        self.assertEqual(rows[0], "1_固定値,KEN:AIN:,")
        self.assertEqual(rows[2], "3_引当No,60524074,")
        self.assertEqual(rows[17], "18_GW2_x,0,")
        self.assertEqual(rows[29], "30_積み合せｺｰﾄﾞ, ,")              # 変換なしでも 30/31 は ZerCon 済み
        self.assertEqual(rows[40], "41_ﾗｲﾝ名,AA,")
        self.assertEqual(rows[41], "42_梱包高さ1,152.5,")
        self.assertEqual(rows[55], "56_梱包高さ15,,")


class HeightCheckTest(unittest.TestCase):
    def test_example_a_boundaries(self) -> None:
        # G2 §10.9 例A: GP=20、GM=15、ErrorA=1.5、ErrorB=0.5
        self.assertFalse(hc(1.5, 2.0, 21.5, 10).ng)
        r = hc(1.5, 2.0, 21.51, 10)
        self.assertEqual((r.ng, r.kind), (True, "over"))
        self.assertFalse(hc(1.5, 2.0, 21.504, 10).ng)                   # "1.50" で OK
        self.assertTrue(hc(1.5, 2.0, 21.506, 10).ng)                    # "1.51" で NG
        self.assertFalse(hc(1.5, 2.0, 14.5, 10).ng)
        r = hc(1.5, 2.0, 14.49, 10)
        self.assertEqual((r.ng, r.kind), (True, "short"))
        self.assertEqual(r.msg_error, "[枚数不足(?)] *測定高さ:14.49*推定枚数:9.66 *入力枚数:10枚")

    def test_ng_clears_heights_and_tag(self) -> None:
        f = AinputForm(remain="0", tag="Load", heights=["1"] * 15)
        r = hc(1.5, 2.0, 21.51, 10, form=f)
        self.assertEqual(r.msg_error, "[許容値ｵｰﾊﾞｰ] *測定高さ:21.51*推定枚数:10.76 *入力枚数:10枚")
        self.assertEqual(r.messages[-1].source, "Input:570")
        self.assertEqual((f.tag, f.heights), ("", [""] * 15))

    def test_example_b_paper(self) -> None:
        # 合紙あり・個装なし: 現行は 枚数×0.029(旧版なら (枚数+1)×0.029)
        f = AinputForm()
        r = hc(2.95, 3.05, 310.1, 100, order="57076309", form=f)
        self.assertAlmostEqual(r.ais, 2.9, places=5)
        self.assertEqual(f.ai_text, "通常2.9")
        self.assertFalse(r.ng)
        self.assertTrue(hc(2.95, 3.05, 310.2, 100, order="57076309", form=AinputForm()).ng)
        self.assertFalse(hc(2.95, 3.05, 297.0, 100, order="57076309", form=AinputForm()).ng)
        self.assertTrue(hc(2.95, 3.05, 296.9, 100, order="57076309", form=AinputForm()).ng)

    def test_example_c_packages(self) -> None:
        f = AinputForm(cou1="3")
        r = hc(2.95, 3.05, 300, 100, order="57076309", form=f)
        self.assertEqual(r.ais, S(S(103 * S(0.029))))
        self.assertEqual(r.gai, S(3 * (2 * 0.119 + 0.119)))
        self.assertEqual(f.ga_text, "1.071")

    def test_outsourcing_four_cases(self) -> None:
        # 外注 × Ai × 合紙あり/なし(現行コードの結果)
        f = AinputForm(gaityuu_text="外注", ai=True)
        self.assertAlmostEqual(hc(3, 3, 300, 100, order="57076309", form=f).ais, 2.9, places=5)
        f = AinputForm(gaityuu_text="外注", ai=False)
        r = hc(3, 3, 300, 100, order="57076309", form=f)
        self.assertEqual((r.ais, f.ai_text), (0.0, "0"))
        f = AinputForm(gaityuu_text="外注", ai=True, remain="0")
        self.assertAlmostEqual(hc(3, 3, 300, 100, form=f).ais, 2.9, places=5)
        f = AinputForm(gaityuu_text="外注", ai=False, remain="0")
        r = hc(3, 3, 300, 100, form=f)
        self.assertEqual((r.ais, f.ai_text), (0.0, ""))

    def test_store_paper(self) -> None:
        f = AinputForm(remain="1")
        r = hc(3, 3, 300, 100, form=f)
        self.assertEqual(r.ais, S(100 * S(0.042)))
        self.assertTrue(f.ai_text.startswith("ストア"))

    def test_vc(self) -> None:
        f = AinputForm(remain="0")
        r = hc(15, 15, 300, 20, order="04077194", form=f)
        self.assertEqual(r.vcs, S(S(S(0.09) + S(0.09)) * 20))
        # ストアでも高精度用途なら加算、それ以外は加算しない(VCAB も変えない)
        f = AinputForm(remain="1", youto="H191")
        self.assertGreater(hc(15, 15, 300, 20, order="04077194", form=f).vcs, 0)
        f = AinputForm(remain="1", youto="H174")
        self.assertEqual((hc(15, 15, 300, 20, order="04077194", form=f).vcs, f.vcab), (0.0, ""))
        f = AinputForm(remain="0", gaityuu_text="外注", vab=True)
        self.assertEqual(hc(15, 15, 300, 20, order="04077194", form=f).vcs, 0.0)
        f = AinputForm(remain="0")
        hc(1, 1, 1, 1, form=f)
        self.assertEqual(f.vcab, "0")                                      # VC=0 の行

    def test_no_order_is_pass(self) -> None:
        r = hc(1.5, 2.0, 999, 10, order="nope")
        self.assertFalse(r.ng)
        self.assertEqual(r.messages[0].text, "Lot情報履歴内にｵｰﾀﾞｰNO無し")
        self.assertTrue(r.show_uf1)
        self.assertFalse(hc(1.5, 2.0, 999, 10, hist=None).ng)

    @unittest.skip("U1 実機確認待ち: 合紙なしで出荷ストア欄が空のとき `Remain = 1` が型エラー13 になるか")
    def test_blank_remain(self) -> None:
        hc(1.5, 2.0, 20, 10, form=AinputForm(remain=""))

    def test_blank_remain_raises_by_policy(self) -> None:
        with self.assertRaises(vbx.TypeMismatch):
            hc(1.5, 2.0, 20, 10, form=AinputForm(remain=""))

    def test_vc_conv_db(self) -> None:
        kensa_fixture.use_offline()
        self.assertEqual(ainput.vc_conv("208PRWH"), (S(0.09), []))
        v, m = ainput.vc_conv("NOPE")
        self.assertEqual((v, m[0].text), (S(0.08406977), "VC厚登録なし,担当者に連絡のこと"))


class HeightInputTest(unittest.TestCase):
    def test_calc_mount_height(self) -> None:
        f = AinputForm(heights=["100.1", "100.2", "x"] + [""] * 12)
        self.assertEqual(ainput.calc_mount_height(f, 2), S(S(100.1) + 100.2))
        self.assertEqual(ainput.calc_mount_height(f, 20), S(S(100.1) + 100.2))
        self.assertTrue(ainput.all_height_filled(f, 0))
        self.assertEqual(ainput.next_empty_height_index(f, 4), 4)

    def test_accept_depth_and_convex(self) -> None:
        f = AinputForm(line_name="L-1", atu="3.000", mai1="50", cou1="1", mount_count="2")
        self.assertEqual(ainput.l1_caption(ainput.l_height(f)), "Depth測定")    # 150/2=75 ≤ 300
        ainput.accept_height(f, "DT10002+0000075.512M", "10003")                # 機器番号が違う → 入れない
        self.assertEqual(f.heights[0], "")
        m = ainput.accept_height(f, "DT10002+0000075.512M", "10002")
        self.assertEqual((f.heights[0], m[0].sound), ("75.512", "chimes.wav"))
        f = AinputForm(line_name="L-1", atu="10", mai1="70", cou1="1", mount_count="2")
        self.assertEqual(ainput.l1_caption(ainput.l_height(f)), "ｺﾝﾍﾞｯｸｽ測定")   # 350 > 300
        ainput.accept_height(f, "DT10002+0000075.512M", "10002")
        self.assertEqual(f.heights[0], "")
        ainput.accept_height(f, "#L351.0", "")
        self.assertEqual(f.heights[0], "351")
        # 丸徳は機器番号が無い(DT でない)ことが条件
        f = AinputForm(line_name="丸徳A1", atu="1", mai1="10", cou1="1", mount_count="1")
        ainput.accept_height(f, "#L10.5", "10002")
        self.assertEqual(f.heights[0], "10.5")

    def test_mount_count(self) -> None:
        f = AinputForm(atu="1", mai1="1", cou1="1")
        m = ainput.input_mount_count(f, "16")
        self.assertEqual(m[0].text, "山数は1～15で入力してください。")
        ainput.input_mount_count(f, "3")
        self.assertEqual((f.mount_count, f.l1_visible, f.l1_caption), ("3", True, "Depth測定"))


class InputCheckTest(unittest.TestCase):
    def test_vcab_cleared_with_vctext(self) -> None:
        # UFAinput:254(CommandButton9)/947(AinButt)で VcText = "" → VcText_Change(UFAinput:69-71)で VCAB も ""。
        # ストア(Remain=1)・用途が高精度でないと Height_Check の VC 分岐(Input:513)を通らず VCAB は "" のまま
        f = AinputForm(remain="1", youto="H279", order_no="04077194", amin="15", amax="15", mount_count="1",
                       heights=["15"] + [""] * 14, mai1="1", cou1="1", vc_text="1.2", vcab="1.2")
        ainput.recheck_height(f, HIST, vc=vc_stub)
        self.assertEqual((f.vc_text, f.vcab), ("", ""))
        # VcText がもともと "" なら Change は起きず、VCAB(Else 分岐の "0" など)はそのまま
        f = AinputForm(remain="1", youto="H279", order_no="04077194", amin="15", amax="15", mount_count="1",
                       heights=["15"] + [""] * 14, mai1="1", cou1="1", vc_text="", vcab="0")
        ainput.recheck_height(f, HIST, vc=vc_stub)
        self.assertEqual(f.vcab, "0")

    def test_konno(self) -> None:
        self.assertEqual(ainput.konno_before_update("5"), "05")
        self.assertEqual(ainput.konno_before_update("0"), "")
        self.assertEqual(ainput.konno_before_update("12"), "12")

    def test_pile_match(self) -> None:
        f = AinputForm(remain="1", lot="S123", pile_match="L234567")
        self.assertEqual(ainput.pile_match_change(f), [])                  # S/F は L でも可
        self.assertEqual(f.pile_match, "L234567")
        f = AinputForm(remain="1", lot="H123", pile_match="L234567")
        m = ainput.pile_match_change(f)
        self.assertEqual((m[0].text, f.pile_match), ("ストアカードの頭文字が違います", ""))
        f = AinputForm(remain="1", lot="H123", pile_match="H23")             # 7 桁になるまで見ない
        self.assertEqual((ainput.pile_match_change(f), f.pile_match), ([], "H23"))
        f = AinputForm(remain="0", pile_match="x")
        ainput.pile_match_change(f)
        self.assertEqual(f.pile_match, "")
        f = AinputForm(remain="1", lot="H1", pile_match="H234567", pile_match_no="1")
        ainput.pile_match_change(f)
        self.assertEqual(f.pile_match, "")                                  # 積合せなら数でないものは消す
        self.assertEqual(ainput.pile_match_key("1", "a"), "A")
        self.assertEqual(ainput.pile_match_key("1", "-"), "")
        self.assertEqual(ainput.pile_match_key("0", "-"), "-")

    def test_pile_match_no(self) -> None:
        f = AinputForm(pile_match_no="1", hiki_no="60524074", konno="01")
        ainput.pile_match_no_change(f)
        self.assertEqual(f.pile_match, "6052407401")
        f = AinputForm(pile_match_no="9", pile_match="x")
        ainput.pile_match_no_change(f)
        self.assertEqual((f.pile_match_no, f.pile_match), ("9", ""))
        # 1/9 以外は "" にした後 `PileMatchNo = 1` と比べる → U1 の方針では型エラー 13
        with self.assertRaises(vbx.TypeMismatch):
            ainput.pile_match_no_change(AinputForm(pile_match_no="5", pile_match="x"))

    def test_remain(self) -> None:
        self.assertEqual(ainput.remain_before_update("2"), "")
        self.assertEqual(ainput.remain_before_update("1"), "1")
        f = AinputForm(remain="1", ex_visible=True)
        ainput.remain_change(f)
        self.assertEqual((f.outside_atu, f.outside_haba, f.outside_take), ("0", "0", "0"))


class HikiTest(unittest.TestCase):
    def test_sort(self) -> None:
        self.assertEqual(ainput.sort_hiki(["10", "9", "A", "", "2"]), ["2", "9", "10", "A"])

    def test_no_search_from_history(self) -> None:
        r = ainput.no_search("N456J52", [HEAD] + HIST)
        self.assertEqual(r.rows, [["60524074", "57076309", "N456J52", "K650969", 0, 2000, None, "1P0116"]])
        self.assertFalse(r.db_flag)

    def test_no_search_db_missing(self) -> None:
        kensa_fixture.use_offline()
        r = ainput.no_search("ZZZ", [HEAD] + HIST)
        self.assertTrue(r.db_flag)
        self.assertEqual(r.messages[0].text, "日稲:SIKACOMP.accdbがありません_ReadCheck_No_Search")

    def test_hiki_no_change(self) -> None:
        rows = [["61029053", "04093803", "RX021T0", "T6X1913", "0", "0"],
                ["61029052", "57193316", "RX021T0", "T6X1913", "0", "4"]]
        s = ainput.hiki_no_change("61029053", rows)
        self.assertEqual((s.order_no, s.adjust_visible, s.export_visible, s.form_height),
                         ("04093803", False, False, 490))
        s = ainput.hiki_no_change("61029052", rows)
        self.assertEqual((s.adjust_label, s.adjust_no, s.export_visible), ("引当数量", "4", True))
        self.assertEqual(ainput.hiki_no_change(None, rows).order_no, "")

    def test_startup(self) -> None:
        r = ainput.startup("UF1", "L1", [HEAD] + HIST, working_lot="N456J52", lname="Test", uf1_line="Test",
                           uf1_minmax={"amin": "2.995"}, workers=" T1:阿部")
        self.assertEqual((r.form.ken_no, r.form.atu, r.form.haba, r.form.amin), ("K650969", "3.000", "1000.0", "2.995"))
        self.assertEqual(r.log.msg, "*UFKen:UFAinput起動 T1:阿部")
        self.assertEqual(len(r.hiki.rows), 1)
        r = ainput.startup("Mitsubishi", "", [HEAD] + HIST, working_lot="N456J52", lname="L-1",
                           sheet_cells={(4, 23): 2.9951, (3, 25): 0})
        self.assertEqual((r.form.amin, r.form.hmax), ("2.995", ""))
        self.assertIsNone(ainput.startup("Cask", "", [HEAD] + HIST, working_lot="?", lname="L-1").form)


class FilesTest(unittest.TestCase):
    def setUp(self) -> None:
        kensa_fixture.use_offline()
        self.tmp = Path(tempfile.mkdtemp(prefix="ufk_ain_"))
        d = user_settings.get(paths.KEY_PATHS)
        d[paths.OFFLINE]["history"] = str(self.tmp / "hist") + "/"
        d[paths.OFFLINE]["nis"] = str(self.tmp / "nis")
        user_settings.save(paths.KEY_PATHS, d)
        (self.tmp / "nis").mkdir()
        self.form = AinputForm(line_name="Test", uf1_line="Test", lot="N456J52", ken_no="K650969",
                               order_no="57076309", atu="3.000", amin="2.995", amax="3.005", hiki_no="60524074",
                               konno="01", mai1="50", cou1="1", mount_count="1")

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_csv_numbering_and_stop(self) -> None:
        p1 = ainput.save_ainput_csv(self.form, "N456J52", "K650969", now=NOW)
        self.assertTrue(p1.endswith("Test\\Test_Ainput\\2026.10\\N456J52\\20261030_K650969.csv"))
        folder = self.tmp / "hist" / "Test" / "Test_Ainput" / "2026.10" / "N456J52"
        data = (folder / "20261030_K650969.csv").read_bytes()
        self.assertTrue(data.startswith("1_固定値,KEN:AIN:,\r\n2_固定値,0185,\r\n".encode("cp932")))
        p2 = ainput.save_ainput_csv(self.form, "N456J52", "K650969", now=NOW)
        self.assertTrue(p2.endswith("20261030_K650969.02.csv"))
        for n in range(3, 100):
            (folder / f"20261030_K650969.{n:02d}.csv").write_text("")
        with self.assertRaises(flow.Stop) as cm:
            ainput.save_ainput_csv(self.form, "N456J52", "K650969", now=NOW)
        self.assertEqual(cm.exception.text, "連番をこれ以上付けることが出来ません。")

    def test_folder_by_uf1_line(self) -> None:
        self.assertEqual(ainput.ainput_folder("ﾊﾞﾗﾝｻｰ"), "バランサー")
        self.assertEqual(ainput.ainput_folder("丸徳"), "丸徳")
        self.assertEqual(ainput.ainput_folder(""), "Test")

    def test_hostsend(self) -> None:
        p = ainput.write_hostsend(self.form)
        data = (self.tmp / "nis" / "hostsend.txt").read_bytes()
        self.assertEqual(data, (ainput.telegram(self.form) + "\r\n").encode("cp932"))
        self.assertTrue(p.endswith("hostsend.txt"))
        self.assertEqual(ainput.host_data_confi(), "現添付Data【引当No_60524074 梱包No_01】")

    def test_launch(self) -> None:
        r = ainput.app_activation(lambda p: True)
        self.assertFalse(r.ok)
        self.assertTrue(r.messages[0].text.endswith("SokuteiDataTrans.exeを開くことができません。\r\n担当者に連絡してください"))
        exe = self.tmp / "nis" / "SokuteiDataTrans" / "SokuteiDataTrans.exe"
        exe.parent.mkdir()
        exe.write_bytes(b"")
        self.assertTrue(ainput.app_activation(lambda p: True).ok)
        r = ainput.app_activation(lambda p: False)
        self.assertEqual((r.ok, r.messages[0].text), (False, "失敗"))

    def test_result(self) -> None:
        sleeps = []
        r = ainput.inter_rupt_text(5000, sleep=sleeps.append)
        self.assertEqual((r.ok, r.missing, sleeps), (True, True, [1.0] * 5))    # 無ければ先へ進む
        res = self.tmp / "nis" / "result.txt"
        res.write_bytes("前の結果\r\n正常に送信が行われました。\r\n".encode("cp932"))
        self.assertTrue(ainput.inter_rupt_text(0, sleep=lambda s: None).ok)
        res.write_bytes("ホストへの送信内容が正しくありません。NISへ連絡してください。".encode("cp932"))
        r = ainput.inter_rupt_text(0, sleep=lambda s: None)
        self.assertFalse(r.ok)
        self.assertEqual(r.messages[0].text, "ﾎｽﾄへの送信内容が正しくありません。担当者へ連絡して下さい。\r\nA002")
        res.write_bytes(b"")
        r = ainput.inter_rupt_text(0, sleep=lambda s: None)
        self.assertEqual((r.ok, r.messages[0].text), (False, ""))

    def test_host_con(self) -> None:
        self.assertEqual(ainput.host_con("KEN:RES:0004A001"), "作成されたテキストの内容が正しくありません。品質保証課へ連絡。")
        self.assertEqual(ainput.host_con("A009"), "")


class AinButtonTest(FilesTest):
    def press(self, answers=None, daily=None):
        return ainput.ain_button(self.form, history_rows=HIST, answers=answers or {},
                                 registry=WorkerRegistry(t1="阿部"), day_max="1", now=NOW, daily=daily, vc=vc_stub,
                                 launcher=lambda p: True, sleep=lambda s: None, computer="PC")

    def test_guards(self) -> None:
        self.form.db_flag = "true"
        self.assertEqual(self.press().stage, "dbflag")
        self.form.db_flag = ""
        self.form.hiki_no = ""
        self.assertEqual(self.press().messages[0].text, "必要項目の入力がありません(引当No)")
        self.form.hiki_no = "1"
        self.form.tag = "Load"
        self.assertEqual(self.press().stage, "busy")

    def test_two_presses_and_send(self) -> None:
        r = self.press()
        self.assertEqual((r.stage, self.form.l1_visible, self.form.l1_caption, self.form.tag),
                         ("first", True, "Depth測定", ""))
        r = self.press()                                   # 山高さが未入力 → もう一度 1 回目の扱い
        self.assertEqual(r.stage, "first")
        self.form.heights[0] = "152.5"
        r = self.press()
        self.assertEqual((r.stage, r.question.text, self.form.tag), ("question", "ｱｲｶｽ転送を行いますか？", ""))
        self.assertEqual(self.press({"ainput_send": False}).stage, "cancel")
        exe = self.tmp / "nis" / "SokuteiDataTrans" / "SokuteiDataTrans.exe"
        exe.parent.mkdir()
        exe.write_bytes(b"")
        (self.tmp / "nis" / "result.txt").write_bytes("正常に送信が行われました。\r\n".encode("cp932"))
        from ufkensa.kensa.workers import DailyContext
        from tests.test_kensa_workers import HIST as WH
        daily = DailyContext(True, {(2, 4): "梱包No", (3, 1): "N456J52"}, WH, "N456J52", NOW, "Test", SHIFTS)
        r = self.press({"ainput_send": True}, daily)
        self.assertEqual(r.stage, "sent")
        self.assertEqual(r.logs_written[0].msg, "*UFKen:50枚_152.5mm * T1:阿部")
        log = (self.tmp / "hist" / "Test" / "Test_Log" / "Access.log").read_bytes().decode("cp932")
        self.assertIn("PC  *UFKen:50枚_152.5mm * T1:阿部 \r\n", log)
        self.assertTrue((self.tmp / "nis" / "hostsend.txt").is_file())
        self.assertIn(CellWrite(3, 4, "01"), r.daily_writes)
        self.assertEqual(self.form.konno, "")
        self.assertEqual(self.form.heights[0], "152.5")    # 送信後も山高さは残る(VBA どおり)
        self.assertTrue(self.form.locked)

    def test_input_errors(self) -> None:
        self.form.l1_visible = True
        self.form.heights[0] = "152.5"
        self.form.cou1 = ""
        self.assertEqual(self.press().messages[0].text, "必要項目の入力がありません(枚数1,包数1)")
        self.form.cou1 = "1"
        self.form.ex_visible, self.form.remain = True, "0"
        self.assertEqual(self.press().messages[0].text, "ｹｰｽNoの入力がありません。")
        self.form.ex_case_no = "1"
        self.form.outside_atu, self.form.outside_haba, self.form.outside_take = "5", "100", "100"
        self.assertEqual(self.press().messages[0].text, "EX外寸に在り得ない数値の入力があります。")
        self.form.ex_visible, self.form.remain = False, "1"
        self.assertEqual(self.press().messages[0].text, "ｽﾄｱｶｰﾄﾞNo入力がありません。")
        self.assertEqual(self.form.tag, "")

    def test_height_ng_logs_msg_error(self) -> None:
        self.form.l1_visible = True
        self.form.heights[0] = "160"
        r = self.press()
        self.assertEqual(r.stage, "height_ng")
        self.assertTrue(r.logs_written[0].msg.startswith("*UFKen:MsgError:[許容値ｵｰﾊﾞｰ] *測定高さ:160.00"))
        self.assertEqual(self.form.heights[0], "")

    def test_marutoku_log_separator(self) -> None:
        self.form.line_name = "丸徳A1"
        self.form.l1_visible = True
        self.form.heights[0] = "152.5"
        exe = self.tmp / "nis" / "SokuteiDataTrans" / "SokuteiDataTrans.exe"
        exe.parent.mkdir()
        exe.write_bytes(b"")
        r = self.press({"ainput_send": True})
        self.assertEqual(r.logs_written[0].msg, "*UFKen:50枚\\152.5mm * T1:阿部")
        self.assertTrue(r.logs_written[0].target_path.endswith("丸徳\\丸徳_Log"))


if __name__ == "__main__":
    unittest.main()
