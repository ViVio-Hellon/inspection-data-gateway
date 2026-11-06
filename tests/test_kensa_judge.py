"""寸法判定(VBA【UFKen_04】Judg)の試験。期待値は VBA を手で追った結果(行番号は J: = Judg)。"""
from __future__ import annotations

import os
import shutil
import sqlite3
import tempfile
import time
import unittest
from ufkensa.kensa import paths
from pathlib import Path

from tests import kensa_fixture
from ufkensa.kensa import judge, quirks, vbx
from ufkensa.kensa.judge import NG, OK, WHITE
from tests._isolation import vba_reproduce


def spec_form(**kw) -> judge.Form:
    """作業保存_Test の値(G2 §10.1/10.8): 厚 24.550〜25.450、巾 1383.0〜1387.0、丈 2768.0〜2773.0。"""
    base = dict(ApK="25.450", AmK="24.550", Am="0.45", HpK="1387.0", HmK="1383.0", Hm="2",
                TpK="2773.0", TmK="2768.0", Tm="2.5", TaikText="1.5", LateralText="2.0",
                LineName="L-1", LineNo="", Lot="L726M50")
    base.update(kw)
    return judge.new_form(**base)


def put(form: judge.Form, **values: str) -> judge.Form:
    for k, v in values.items():
        form.cells[k].value = v
    return form


class QuirkCase(unittest.TestCase):
    def tearDown(self) -> None:
        quirks.set_for_test(None)


# ------------------------------------------------------------------
class AhtJudgTest(unittest.TestCase):
    """AHTJudg(J:362-369)。下限 ≤ 値 ≤ 上限(両端を含む)。G2 §10.1"""

    def test_boundaries(self) -> None:
        kp, km = 25.45, 24.55
        self.assertTrue(judge.aht_judg("24.55", kp, km))      # 下限ちょうど
        self.assertFalse(judge.aht_judg("24.549", kp, km))
        self.assertTrue(judge.aht_judg("25.45", kp, km))      # 上限ちょうど
        self.assertFalse(judge.aht_judg("25.451", kp, km))

    def test_not_numeric(self) -> None:
        self.assertFalse(judge.aht_judg("abc", 25.45, 24.55))
        self.assertFalse(judge.aht_judg("", 25.45, 24.55))
        self.assertFalse(judge.aht_judg("12.3abc", 25.45, 0))  # IsNumeric で先に弾く

    def test_judg_color_thickness(self) -> None:
        f = put(spec_form(), S1="25.45")
        out = judge.judg_color(f, "S1", "A", "検査太郎")
        self.assertTrue(out.result)
        self.assertEqual(out.states, {"S1": OK})
        self.assertEqual(out.sounds, [judge.SOUND_OK])
        self.assertEqual(out.images, ["Y"])
        # ログは Measured.log と既定の2回(J:539-540)
        self.assertEqual(len(out.logs), 2)
        self.assertEqual(out.logs[0].text, "*UFKen|25.45|(S1)|○|L-1|検査太郎")
        self.assertEqual(out.logs[0].target_file, "Measured.log")
        self.assertTrue(out.logs[0].target_path.endswith("L-1\\L-1_Log"))
        self.assertEqual((out.logs[1].target_path, out.logs[1].target_file), ("", ""))

    def test_judg_color_ng(self) -> None:
        f = put(spec_form(), S5="1382.9")
        out = judge.judg_color(f, "S5", "H")
        self.assertFalse(out.result)
        self.assertEqual(out.states, {"S5": NG})
        self.assertEqual(out.logs[0].text, "*UFKen|1382.9|(S5)|×|L-1|")

    def test_minus_tolerance_blank_skips(self) -> None:
        # Am="" なら判定しない(J:533)。プラス側 Ap は見ない
        f = put(spec_form(Am=""), S1="99")
        out = judge.judg_color(f, "S1", "A")
        self.assertEqual((out.states, out.logs), ({}, []))

    def test_limit_text_not_numeric_is_zero(self) -> None:
        # On Error Resume Next: TpK="" の代入はエラーで KP=0 のまま → 値 > 0 は NG(J:483, 499)
        f = put(spec_form(TpK=""), S7="2770")
        self.assertEqual(judge.judg_color(f, "S7", "T").states, {"S7": NG})

    def test_line_folder_is_exact_match(self) -> None:
        f = put(spec_form(LineName="ﾄｯﾄ"), S1="25")
        self.assertTrue(judge.judg_color(f, "S1", "A").logs[0].target_path.endswith("トット\\トット_Log"))
        f = put(spec_form(LineName="丸徳", LineNo="A1"), S1="25")
        self.assertTrue(judge.judg_color(f, "S1", "A").logs[0].target_path.endswith("丸徳\\丸徳_Log"))
        f = put(spec_form(LineName="L-1", LineNo="2"), S1="25")      # "L-12" は Case Else
        self.assertTrue(judge.judg_color(f, "S1", "A").logs[0].target_path.endswith("丸徳\\丸徳_Log"))


class DiagonalTest(unittest.TestCase):
    """対角線差(J:380-394, 584-715)。G2 §10.2(TaikText="1.5")"""

    def judge_pair(self, s9: str, s10: str, **kw) -> judge.Outcome:
        f = put(spec_form(**kw), S9=s9, S10=s10)
        return judge.judg_color(f, "S10", "D")

    def test_cases(self) -> None:
        out = self.judge_pair("3145.9", "3145.9")
        self.assertEqual((out.result, out.values["DiagS"]), (True, "0"))
        out = self.judge_pair("3145.9", "3147.4")                 # 差 1.5 = 公差 → OK
        self.assertEqual((out.result, out.values["DiagS"]), (True, "1.5"))
        self.assertEqual(out.states, {"S9": OK, "S10": OK})
        out = self.judge_pair("3145.9", "3147.5")                 # 差 1.6 → NG
        self.assertEqual((out.result, out.values["DiagS"]), (False, "1.6"))
        self.assertEqual(out.states, {"S9": NG, "S10": NG})
        out = self.judge_pair("2947.3", "2946.4")                 # Double では 0.900000000000091
        self.assertEqual((out.result, out.values["DiagS"]), (True, "0.9"))

    def test_tolerance_blank_or_zero_is_ng(self) -> None:
        self.assertFalse(self.judge_pair("3145.9", "3145.9", TaikText="").result)     # kou=0 のまま
        self.assertFalse(self.judge_pair("3145.9", "3145.9", TaikText="0.0").result)  # 公差 0 は差 0 でも NG

    def test_first_of_pair(self) -> None:
        # 9 だけ入れたときは旧形式のログだけ書いて判定しない(J:587-592)
        f = put(spec_form(), S9="3145.9")
        out = judge.judg_color(f, "S9", "D")
        self.assertEqual(out.states, {})
        self.assertNotIn("DiagS", out.values)
        self.assertEqual([x.text for x in out.logs], ["*UFKen:3145.9:(S9):_:L-1"] * 2)
        # 10 だけ入れたときはログもなし
        f = put(spec_form(), S10="3145.9")
        self.assertEqual(judge.judg_color(f, "S10", "D").logs, [])

    def test_diagonal_judg_single(self) -> None:
        self.assertTrue(judge.diagonal_judg(1.5, 1.5))
        self.assertFalse(judge.diagonal_judg(0, 0))
        self.assertTrue(judge.diagonal_judg(0, 1))
        self.assertTrue(judge.diagonal_judg(1.0000001, 1.0000001))

    def test_re_diag(self) -> None:
        f = put(spec_form(), S9="2947.3", S10="2946.4", U9="10", E10="1")
        self.assertEqual(judge.re_diag(f).values, {"DiagS": "0.9"})
        f.gokuatsu = True
        self.assertEqual(judge.re_diag(f).values, {})


@vba_reproduce
class LateralTest(QuirkCase):
    """LateralJudg / UnitRecord(J:395-409, 2421-2443)。G2 §10.3(LateralText="2.0")"""

    def test_lateral_judg(self) -> None:
        self.assertTrue(judge.lateral_judg(0, "2.0"))
        self.assertTrue(judge.lateral_judg(2.0, "2.0"))
        self.assertFalse(judge.lateral_judg(2.5, "2.0"))
        self.assertFalse(judge.lateral_judg(-2.5, "2.0"))
        self.assertTrue(judge.lateral_judg(-2.0, "2.0"))
        self.assertFalse(judge.lateral_judg(0, "0"))           # kou="0" は先に False
        self.assertFalse(judge.lateral_judg(1, ""))
        self.assertTrue(judge.lateral_judg(0, "0.0"))           # "0.0" は "0" と一致しない → a=0 で True
        self.assertFalse(judge.lateral_judg(0.5, "0.0"))

    def test_unit_record(self) -> None:
        cases = {0: 0.0, 0.05: 0.0, 0.3: 0.5, 0.5: 0.5, 1.6: 2.0, 2.1: 2.5, -2.1: -2.5, -1.9: -2.0,
                 2.6: 3.0, 4.01: 4.5, 4.51: 5.0, 5.0: 5.0, 5.01: vbx.single(5.01), 5.3: vbx.single(5.3)}
        for n, want in cases.items():
            self.assertEqual(judge.unit_record(n), want, n)
        self.assertEqual(str(judge.unit_record(-0.0)), "0.0")    # -0 は出さない(V10)

    def test_b13_single_gap(self) -> None:
        for n in (0.51, 1.01, 1.51, 2.01, 2.51, 3.01, 3.51, 1000):
            self.assertEqual(judge.unit_record(n), 0.0, n)       # B13 再現
        quirks.set_for_test({"B13": "fix"})
        fixed = {0.51: 1.0, 1.01: 1.5, 1.51: 2.0, 2.01: 2.5, 2.51: 3.0, 3.01: 3.5, 3.51: 4.0,
                 -1.01: -1.5, 1000: 1000.0, 0.05: 0.5}   # 0.1 未満も区間に入れる(2026-11-02 の回答)
        for n, want in fixed.items():
            self.assertEqual(judge.unit_record(n), want, n)

    def test_unit_record_input(self) -> None:
        self.assertEqual(judge.unit_record_input("1.6"), "2.0")
        self.assertEqual(judge.unit_record_input("-.3"), "-0.5")
        self.assertEqual(judge.unit_record_input("0"), "0.0")
        self.assertEqual(judge.unit_record_input("-"), "-")
        self.assertEqual(judge.unit_record_input("abc"), "abc")
        self.assertEqual(judge.unit_record_input(".51"), "0.0")    # B13

    def test_lateralbow_judg_color(self) -> None:
        f = put(spec_form(), LateralS="2.0")
        out = judge.lateralbow_judg_color(f, "LateralS", "W")
        self.assertEqual((out.result, out.states), (True, {"LateralS": OK}))
        self.assertEqual(out.logs[0].text, "*UFKen|2.0|(LateralS)|○|L-1|W")
        self.assertEqual(out.mpex_tag, "")
        f = put(spec_form(), LateralS="-2.5")
        self.assertEqual(judge.lateralbow_judg_color(f, "LateralS").states, {"LateralS": NG})
        for v in ("", "-", "abc"):
            self.assertEqual(judge.lateralbow_judg_color(put(spec_form(), LateralS=v), "LateralS").states, {})

    def test_lateralbow_no_tolerance(self) -> None:
        f = put(spec_form(LateralText=""), LateralC="1.0")
        out = judge.lateralbow_judg_color(f, "LateralC")
        self.assertEqual(out.messages[0].text, "ﾗﾃﾗﾙﾎﾞｰ公差がありません。(規定寸法範囲外)")
        self.assertEqual(out.tags, {"LateralC": "範囲外"})
        self.assertEqual((out.states, out.logs), ({}, []))

    def test_right_angle_fixed_1(self) -> None:
        for v, want in (("1.0", OK), ("-1.0", OK), ("1.5", NG), ("0.0", OK)):
            f = put(spec_form(LateralText="0"), DiagSKou=v)       # 公差は LateralText と無関係
            self.assertEqual(judge.right_angle_judg_color(f, "DiagSKou").states, {"DiagSKou": want}, v)
        self.assertEqual(judge.right_angle_judg_color(put(spec_form(), DiagSKou="-"), "DiagSKou").states, {})
        # 1.1 は入力時に UnitRecord で 1.5 になってから判定される(UF1:1109)
        self.assertEqual(judge.unit_record_input("1.1"), "1.5")


@vba_reproduce
class BhTest(unittest.TestCase):
    """BhJudg / simplyjudge(J:410-478)。G2 §10.4(下限値 40)"""

    def test_bh_judg(self) -> None:
        self.assertEqual(judge.bh_judg("39", 40), "NG")
        self.assertEqual(judge.bh_judg("40", 40), "OK")
        self.assertEqual(judge.bh_judg("", 40), "不要")
        self.assertEqual(judge.bh_judg("50", 0), "不要")
        self.assertEqual(judge.bh_judg(vbx.fmt("39.5", "0"), 40), "OK")   # 入力時に Format "0"
        self.assertEqual(judge.bh_judg("abc", 40), "NG")                   # Val=0

    def test_simply_judge_messages(self) -> None:
        # S だけ検査済みで SBH が NG → 作業中止…
        f = put(spec_form(), S1="25", SBH="30")
        out = judge.simply_judge(f, "SBH", "NG", "W")
        self.assertEqual(out.states, {"SBH": NG})
        self.assertEqual(out.logs[0].text, "*BH|30|(SBH)|×|L-1|W")
        self.assertEqual([m.text for m in out.msg_errors], ["作業中止・保留・ロット異常連絡票発行"])
        self.assertEqual(out.msg_errors[0].icon, 1)
        # E まで検査済みで EBH が NG
        f = put(spec_form(), S1="25", U1="25", C1="25", D1="25", E1="25", EBH="30")
        out = judge.simply_judge(f, "EBH", "NG")
        self.assertEqual([m.text for m in out.msg_errors], ["梱包完了させてAインプット・保留・ロット異常連絡票発行"])
        # 途中の位置は出さない
        f = put(spec_form(), S1="25", U1="25", UBH="30")
        self.assertEqual(judge.simply_judge(f, "UBH", "NG").msg_errors, [])

    def test_simply_judge_blank_caption(self) -> None:
        # Iti="" と見出し "" が一致してメッセージが出る(スラブの E 見出しは "")。VBA どおり
        f = put(spec_form(), S1="25", U1="25", UBH="30")
        f.captions["E"] = ""
        self.assertEqual([m.text for m in judge.simply_judge(f, "UBH", "NG").msg_errors],
                         ["梱包完了させてAインプット・保留・ロット異常連絡票発行"])

    def test_simply_judge_ok_and_unneeded(self) -> None:
        f = put(spec_form(), SBH="45")
        out = judge.simply_judge(f, "SBH", "OK")
        self.assertEqual((out.states, out.sounds), ({"SBH": OK}, [judge.SOUND_OK]))
        out = judge.simply_judge(f, "SBH", "不要")
        self.assertEqual((out.states, out.logs, out.sounds), ({"SBH": WHITE}, [], []))


class HizumiTest(unittest.TestCase):
    """HizumiJudg(J:134-291)。G2 §10.6(HaHiz="1.0"、TaHiz="2.0"、CBS1="幅")"""

    def form(self, **values: str) -> judge.Form:
        f = spec_form(HaHiz="1.0", TaHiz="2.0")
        f.combos["CBS1"] = "幅"
        return put(f, **values)

    def test_surface_plate(self) -> None:
        self.assertEqual(judge.hizumi_judg(self.form(ZS="1.0"), "ZS").states, {"ZS": OK})
        out = judge.hizumi_judg(self.form(ZS="1.01"), "ZS")
        self.assertEqual((out.states, out.sounds, out.move_flag), ({"ZS": NG}, [judge.SOUND_NG], ""))
        self.assertEqual(judge.hizumi_judg(self.form(ZS="0.8/3"), "ZS").states, {"ZS": OK})
        self.assertEqual(judge.hizumi_judg(self.form(ZS="abc"), "ZS").states, {})
        self.assertEqual(judge.hizumi_judg(self.form(ZS=""), "ZS").states, {"ZS": WHITE})
        f = self.form(ZS="1.0")
        f.combos["CBS1"] = ""
        self.assertEqual(judge.hizumi_judg(f, "ZS").states, {})

    def test_gauge_is_white(self) -> None:
        self.assertEqual(judge.hizumi_judg(self.form(GS="0.5"), "GS").states, {"GS": WHITE})

    def test_sound_only_single_or_mousedown(self) -> None:
        self.assertEqual(judge.hizumi_judg(self.form(ZS="1.0"), "ZS", "Loop").sounds, [])
        self.assertEqual(judge.hizumi_judg(self.form(ZS="1.0"), "ZS", "Loop", move_flag="on").sounds,
                         [judge.SOUND_OK])

    def test_free_sori(self) -> None:
        f = self.form(GS="3")
        f.HURIS, f.flatnesText = "3", "ﾌﾘｰｿﾘ要"
        f.combos["CBS1"] = ""                                      # 形状が空でも判定する
        self.assertEqual(judge.hizumi_judg(f, "GS").states, {"GS": OK})
        f.cells["GS"].value = "3.1/2"
        self.assertEqual(judge.hizumi_judg(f, "GS").states, {"GS": NG})

    def test_both_gauges(self) -> None:
        f = self.form(ZS="1.5", GS="1.5")
        f.captions.update(LZS="幅ｹﾞｰｼﾞ", LGS="丈ｹﾞｰｼﾞ")
        f.combos["CBS3"] = "丈"
        self.assertEqual(judge.hizumi_judg(f, "ZS").states, {"ZS": NG})     # 幅 1.0 と比べる
        self.assertEqual(judge.hizumi_judg(f, "GS").states, {"GS": OK})     # 丈 2.0 と比べる
        f.cells["ZS"].value = "0.8/3"                                      # "/" を処理しない → 型エラー
        with self.assertRaises(vbx.TypeMismatch):
            judge.hizumi_judg(f, "ZS")

    @unittest.skip("U1 実機確認待ち: 数値と '' の比較が型エラーになるか")
    def test_both_gauges_blank_tolerance_u1(self) -> None:
        f = self.form(GS="1.5")
        f.captions.update(LZS="幅ｹﾞｰｼﾞ", LGS="丈ｹﾞｰｼﾞ")
        f.combos["CBS3"] = "中"                                      # ComValueG = ""
        with self.assertRaises(vbx.TypeMismatch):
            judge.hizumi_judg(f, "GS")

    def test_hiz_loop_move_flag(self) -> None:
        f = self.form(ZS="1.0", ZU="2", GS="0.5")
        f.combos["CBU1"] = "丈"
        f.MoveFlag = "on"
        out = judge.hiz_loop_judg(f)
        self.assertEqual(out.states["ZS"], OK)
        self.assertEqual(out.states["ZU"], OK)
        self.assertEqual(out.states["GS"], WHITE)
        self.assertEqual(out.states["ZF"], WHITE)
        self.assertEqual(out.sounds, [judge.SOUND_OK])             # 1マス目で MoveFlag が "" に戻る
        self.assertEqual(out.move_flag, "")
        self.assertEqual(judge.hiz_loop_judg(self.form(ZS="1"), n=2).states, {})


@vba_reproduce
class LoopJudgTest(QuirkCase):
    """LoopJudg(J:836-1116)"""

    def test_basic(self) -> None:
        f = put(spec_form(), S1="25", S2="26", S5="1385", S7="2770", S9="3145.9", S10="3147.4",
                DiagS="1.5", LateralS="2.0", F1="25", SBH="45", DiagSKou="1.0")
        out = judge.loop_judg(f, bh_lower=40)
        st = out.states
        self.assertEqual((st["S1"], st["S2"], st["S5"], st["S7"], st["S9"], st["S10"]), (OK, NG, OK, OK, OK, OK))
        self.assertEqual((st["F1"], st["S3"], st["S11"]), (OK, WHITE, WHITE))
        self.assertEqual((st["LateralS"], st["DiagSKou"], st["SBH"]), (OK, OK, WHITE))   # BH の OK は白
        self.assertEqual(st["LateralU"], WHITE)                       # "" は白
        self.assertEqual(out.borders["S1"], "")

    def test_bh_ng_and_unneeded(self) -> None:
        f = put(spec_form(), SBH="30", UBH="45")
        st = judge.loop_judg(f, bh_lower=40).states
        self.assertEqual((st["SBH"], st["UBH"]), (NG, WHITE))
        self.assertNotIn("SBH", judge.loop_judg(f, bh_lower=0).states)   # 不要は色を変えない

    def test_exit_conditions(self) -> None:
        f = put(spec_form(), S1="25")
        self.assertEqual(judge.loop_judg(f, n=2).states, {})
        self.assertEqual(judge.loop_judg(put(spec_form(ApK=""), S1="25")).states, {})
        f.kosei_visible = True
        self.assertEqual(judge.loop_judg(f).states, {})

    def test_minus_blank_rows_white_except_f(self) -> None:
        f = put(spec_form(Am=""), S1="99", F1="99")
        st = judge.loop_judg(f).states
        self.assertEqual((st["S1"], st["F1"]), (WHITE, NG))

    def test_length_blank_skips_diagonal(self) -> None:
        # 丈の上下限が空なら対角もとばす(J:983 GoTo SkipL)
        f = put(spec_form(TpK=""), S9="3145.9", S10="3145.9", DiagS="0")
        st = judge.loop_judg(f).states
        self.assertEqual((st["S9"], st["S10"]), (NG, NG))
        f.TpK = "2773.0"
        st = judge.loop_judg(f).states
        self.assertEqual((st["S9"], st["S10"]), (OK, OK))

    def test_lateral_text_not_numeric_stops(self) -> None:
        # LateralText が数でないと BH・直角度・空欄の白戻しをとばす(J:1059)
        f = put(spec_form(LateralText=""), LateralS="1.0", DiagSKou="0.5", SBH="45")
        st = judge.loop_judg(f, bh_lower=40).states
        self.assertEqual((st["LateralS"], st["DiagSKou"], st["LateralU"]), (NG, NG, NG))
        self.assertNotIn("SBH", st)

    def test_b8_right_angle_d_e(self) -> None:
        f = put(spec_form(LateralText=""))
        f.cells["DiagDKou"].visible = False                         # 4山などで D を出さない
        st = judge.loop_judg(f).states
        self.assertEqual(st.get("DiagDKou"), NG)                    # B8 再現: E が表示 → D を塗る
        self.assertNotIn("DiagEKou", st)
        quirks.set_for_test({"B8": "fix"})
        st = judge.loop_judg(f).states
        self.assertNotIn("DiagDKou", st)
        self.assertEqual(st.get("DiagEKou"), NG)

    def test_gokuatsu(self) -> None:
        f = put(spec_form(gokuatsu=True), S9="1385", S11="2770", S10="1385")
        st = judge.loop_judg(f).states
        self.assertEqual((st["S9"], st["S10"], st["S11"]), (OK, OK, OK))

    def test_gokuatsu_blank_length_is_runtime_error(self) -> None:
        f = put(spec_form(gokuatsu=True, TpK=""), S1="25", S11="2770")
        with self.assertRaises(judge.LoopJudgAborted) as cm:
            judge.loop_judg(f)
        self.assertEqual(cm.exception.outcome.states["S1"], OK)    # エラーまでに塗った色は残る

    def test_range(self) -> None:
        f = put(spec_form(), S1="25", U1="25")
        st = judge.loop_judg(f, j=2, k=2).states
        self.assertEqual((st["S1"], st["U1"]), (NG, OK))

    def test_looking_yellow(self) -> None:
        f = spec_form()
        f.cells["S3"].state, f.cells["S3"].border = judge.YELLOW, judge.RED
        f.cells["F2"].border = judge.RED
        out = judge.looking_yellow(f)
        self.assertEqual(out.states, {"S3": WHITE})
        self.assertEqual(out.borders, {"S3": "", "F2": ""})
        judge.apply_outcome(f, out)
        self.assertEqual((f.cells["S3"].state, f.cells["S3"].border), (WHITE, ""))


class TextJudgTest(unittest.TestCase):
    """TextJudg(J:748-835)"""

    def test_routing(self) -> None:
        f = put(spec_form(), S9="3145.9", S10="3145.9")
        out = judge.text_judg(f, "S9")
        self.assertEqual(out.sounds[:1], [judge.SOUND_DIAG])
        self.assertEqual(out.images[:1], ["Neko"])
        self.assertEqual(out.states, {"S10": OK, "S9": OK})
        self.assertEqual(out.ng_erase, "S9")
        f.gokuatsu = True
        out = judge.text_judg(f, "S9")                              # 極厚の 9 は巾
        self.assertEqual(out.states, {"S9": NG})
        self.assertNotIn(judge.SOUND_DIAG, out.sounds)

    def test_skip(self) -> None:
        f = put(spec_form(), S1="25", S11="5")
        self.assertEqual(judge.text_judg(f, "S11").states, {})      # 極厚でない 11 は判定なし
        self.assertEqual(judge.text_judg(f, "S11").ng_erase, "S11")
        f.AbnoTag = "NG"
        self.assertIsNone(judge.text_judg(f, "S1").ng_erase)
        f.AbnoTag, f.msgbox_visible = "", True
        self.assertIsNone(judge.text_judg(f, "S1").ng_erase)


class ToleranceTableTest(unittest.TestCase):
    """LookupDiag / LookupLat(J:1153-1209)。受領した品質基準の表で。G2 §10.5"""

    @classmethod
    def setUpClass(cls) -> None:
        kensa_fixture.use_offline()
        judge.reload_tolerance()

    def test_test_tolerance_examples(self) -> None:
        # Test_Tolerance の期待値(J:1783-1786)
        self.assertEqual(judge.diagonal(500, 2000), 1.0)
        self.assertEqual(judge.re_diagonal(1000, 5000), 5.0)
        self.assertEqual(judge.lateralbow_j(500, 2000, 3.2), 1.5)
        self.assertEqual(judge.re_lateralbow(300, 4000, 8), 15.0)

    def test_boundaries(self) -> None:
        self.assertEqual(judge.diagonal(300, 3700), vbx.single(0.8))  # 両端を含む。戻り値は CSng
        self.assertEqual(judge.diagonal(300.1, 3701), 1.5)
        self.assertEqual(judge.diagonal(1500.1, 1000), 2.5)        # Single 1500.0999… → Round で 1500.1
        self.assertEqual(judge.diagonal(300.05, 1000), 0.0)        # 区分のすき間 → 0
        self.assertEqual(judge.diagonal(500, 3700.5), 0.0)
        self.assertEqual(judge.diagonal(500, 7000.1), 0.0)
        self.assertEqual(judge.diagonal(2400, 7000), 3.0)

    def test_overlap_smallest_number(self) -> None:
        # 巾 900 は No.89(250.1-900, 公差 2)と No.97(900-2000, 公差 1)の両方 → 管理番号の小さい 89
        self.assertEqual(judge.re_lateralbow(900, 1000, 1), 2.0)
        self.assertEqual(judge.re_lateralbow(900.1, 1000, 1), 1.0)


class ToleranceLoadTest(unittest.TestCase):
    """読み込みの失敗と D11(ファイル更新で読み直す)"""

    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="judge_tol_"))
        self.db = self.tmp / "品質基準.sqlite3"

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp, ignore_errors=True)

    def make(self, tol: str, with_tol_col: bool = True) -> None:
        if self.db.exists():
            self.db.unlink()
        con = sqlite3.connect(self.db)
        dcols = ["管理番号", "等級", "巾_下限", "巾_上限", "丈_下限", "丈_上限"] + (["公差"] if with_tol_col else [])
        con.execute("CREATE TABLE 製品規格直角度 (" + ",".join(f'"{c}"' for c in dcols) + ")")
        row = [5, "4", 1, 1000, 1, 1000] + ([tol] if with_tol_col else [])
        con.execute(f"INSERT INTO 製品規格直角度 VALUES ({','.join('?' * len(dcols))})", row)
        con.execute("INSERT INTO 製品規格直角度 VALUES (" + ",".join("?" * len(dcols)) + ")",
                    [3, "4", 1, 1000, 1, 1000] + (["9"] if with_tol_col else []))
        lcols = judge.LAT_COLS
        con.execute("CREATE TABLE 製品規格ラテラルボー (" + ",".join(f'"{c}"' for c in lcols) + ")")
        con.execute("INSERT INTO 製品規格ラテラルボー VALUES (" + ",".join("?" * 9) + ")",
                    [1, "4", 0.1, 10, 1, 1000, 1, 1000, tol])
        con.commit()
        con.close()

    def test_missing_column_message_once(self) -> None:
        self.make("1.5", with_tol_col=False)
        msgs: list = []
        self.assertEqual(judge.diagonal(500, 500, db=self.db, messages=msgs), 0.0)
        self.assertEqual(judge.diagonal(500, 500, db=self.db, messages=msgs), 0.0)
        self.assertEqual(len(msgs), 1)
        self.assertEqual(msgs[0].text,
                         "公差テーブルを読み込めませんでした\r\n段階: 製品規格直角度 読込\r\n"
                         "内容: -2147221502 「製品規格直角度」に「公差」フィールドがありません\r\n\r\n"
                         "公差は表示されません。原因を解消後 ReloadTolerance を実行してください。")

    def test_smallest_number_and_reload_on_update(self) -> None:
        self.make("1.5")
        self.assertEqual(judge.diagonal(500, 500, db=self.db), 9.0)   # 管理番号 3 < 5
        self.assertEqual(judge.lateralbow_j(500, 500, 1, db=self.db), 1.5)
        time.sleep(0.01)
        self.make("2.5")
        st = os.stat(self.db)
        os.utime(self.db, (st.st_atime, st.st_mtime + 5))
        self.assertEqual(judge.lateralbow_j(500, 500, 1, db=self.db), 2.5)   # D11


class AdvanceCheckTest(unittest.TestCase):
    """AdvanceCheck(J:334-361)。受領 SIKALOTNOW の 108 ロットから"""

    @classmethod
    def setUpClass(cls) -> None:
        kensa_fixture.use_offline()

    def test_lots(self) -> None:
        self.assertTrue(judge.advance_check("L726M50"))    # 1.98mm・表面 B・用途 T245
        self.assertTrue(judge.advance_check("LX120J0"))    # 1mm・表面 B・用途 U19B(flag1 and flag2)
        self.assertFalse(judge.advance_check("L8281G0"))   # 0.3mm だが表面 S
        self.assertFalse(judge.advance_check("R6550H0"))   # 6.75mm・用途 K434
        self.assertFalse(judge.advance_check(""))
        self.assertFalse(judge.advance_check("NOLOT00"))

    def test_boundary_single(self) -> None:
        fields = ["ﾛｯﾄ番号", "製造板厚", "品質ｸﾞﾚｰﾄﾞ_表面処理", "用途ｺｰﾄﾞ"]

        def box(thick: str):
            return lambda t, f, use, p: fields if use == "Fields" else [["X", thick, "B", "K1"]]
        self.assertTrue(judge.advance_check("X", arrbox=box("3.0")))
        self.assertTrue(judge.advance_check("X", arrbox=box("3.00000001")))   # Single では 3.0
        self.assertFalse(judge.advance_check("X", arrbox=box("3.01")))

    def test_missing_column_is_error_9(self) -> None:
        box = lambda t, f, use, p: ["ﾛｯﾄ番号"] if use == "Fields" else [["X"]]
        with self.assertRaises(judge.VbaError):
            judge.advance_check("X", arrbox=box)


@vba_reproduce
class ConditionFieldTest(unittest.TestCase):
    """IsMatchConditionField / IsWildcardMatch(J:2346-2411)。G2 §10.12"""

    def test_cases(self) -> None:
        m = judge.is_match_condition_field
        self.assertFalse(m("≠ﾆﾎﾝﾊﾂｼﾞﾖｳ*", "ﾆﾎﾝﾊﾂｼﾞﾖｳ(ｶ"))
        self.assertTrue(m("≠ﾆﾎﾝﾊﾂｼﾞﾖｳ*", "ﾄﾖﾀ"))
        self.assertTrue(m("＊", "何でも"))
        self.assertTrue(m(" ", "x"))
        self.assertTrue(m("abc", "ABC"))                  # 完全一致は大文字小文字を区別しない
        self.assertFalse(m("abc*", "ABCD"))               # 前方一致は区別する
        self.assertTrue(m("ABC*", "ABCD"))
        self.assertTrue(m("*BCD", "ABCD"))
        self.assertTrue(m("*bc*", "ABCD"))                # 部分一致は区別しない
        self.assertTrue(m("X, ABC*", "ABCD"))             # カンマ区切りの OR
        self.assertTrue(m("≠A,≠B", "A"))                 # 除外どうしも OR → ほぼ常に True
        self.assertFalse(m("X,Y", "Z"))

    def test_lone_star_token_is_error_5(self) -> None:
        with self.assertRaises(judge.VbaError):
            judge.is_match_condition_field("A,*", "B")


@vba_reproduce
class PackagingTest(QuirkCase):
    """PackagingSpecificationNo / MaterialsMsgBox(J:1987-2344)"""

    PCOLS = ("包装仕様NO", "用途コード", "コメント", "備考", "受注材質", "受注調質", "厚下", "厚上",
             "板幅下", "板幅上", "板丈下", "板丈上", "フラグ", "形状", "納入先", "取引先")

    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="judge_pack_"))
        kensa_fixture.use_offline(self.tmp)
        self.pack = self.tmp / judge.PACKING_MASTER_FILE
        con = sqlite3.connect(self.pack)
        con.execute("CREATE TABLE 注意_包装仕様 (" + ",".join(f'"{c}"' for c in self.PCOLS) + ")")
        con.execute("INSERT INTO 注意_包装仕様 (包装仕様NO, 形状, コメント) VALUES ('ZZ', '板', '別の仕様')")
        con.commit()
        con.close()
        con = sqlite3.connect(self.tmp / "SIKAODR.sqlite3")
        con.execute('CREATE TABLE 仕掛 ("受注番号","包装仕様NO","用途ｺｰﾄﾞ","受注材質","受注調質","納入先名称","取引先名称")')
        con.execute("INSERT INTO 仕掛 VALUES ('ORD1','P1','R115','52S','H34','ﾆｯｼｬ','ﾄﾘ')")
        con.commit()
        con.close()

    def tearDown(self) -> None:
        super().tearDown()
        shutil.rmtree(self.tmp, ignore_errors=True)

    def add(self, **kw) -> None:
        row = {c: "" for c in self.PCOLS}
        row.update({"包装仕様NO": "P1", "形状": "板"})
        row.update(kw)
        con = sqlite3.connect(self.pack)
        con.execute("INSERT INTO 注意_包装仕様 VALUES (" + ",".join("?" * len(self.PCOLS)) + ")",
                    [row[c] for c in self.PCOLS])
        con.commit()
        con.close()

    def psn(self, thick=1.0, width=1000.0, length=2000.0, nounyu="ﾆｯｼｬ", torihiki="ﾄﾘ") -> str:
        return judge.packaging_specification_no("P1", "R115", thick, width, length, "52S", "H34", nounyu, torihiki)

    def msgbox(self, flag: bool, **form_kw) -> judge.Outcome:
        f = spec_form(Lot="L1", ATU="1.0", Hab="1000", Tak="2000", **form_kw)
        fields = ["ﾛｯﾄ番号", "製造板厚", "品質ｸﾞﾚｰﾄﾞ_表面処理", "用途ｺｰﾄﾞ"]
        rec = [["L1", "1" if flag else "9", "B", "K1"]]
        return judge.materials_msgbox(f, "ORD1", arrbox=lambda t, fl, use, p: fields if use == "Fields" else rec)

    def test_single_row(self) -> None:
        self.assertEqual(self.psn(), "")                                    # レコードなし
        self.add(コメント="注意A")
        self.assertEqual(self.psn(), "板_注意A")

    def test_empty_table_is_type_mismatch(self) -> None:
        # `Dim PackF(): PackF = GetFieldsArr(…)` に Empty が入ると型エラー 13(J:2195-2196)
        con = sqlite3.connect(self.pack)
        con.execute("DELETE FROM 注意_包装仕様")
        con.commit()
        con.close()
        with self.assertRaises(vbx.TypeMismatch):
            self.psn()

    def test_single_row_dims(self) -> None:
        self.add(フラグ="寸法", コメント="寸法注意", 厚下="0.5", 厚上="1.2", 板幅下="900", 板幅上="1100",
                 板丈下="1500", 板丈上="2500")
        self.assertEqual(self.psn(thick=1.0), "板_寸法注意")
        self.assertEqual(self.psn(thick=0.5), "板_寸法注意")             # Single(0.5) は正確
        self.assertEqual(self.psn(thick=1.2), "板_")                      # Single(1.2) > 1.2 で上限から外れる
        self.assertEqual(self.psn(thick=1.3), "板_")

    def test_multi_rows_prepend_and_b11(self) -> None:
        self.add(コメント="A")
        self.add(フラグ="用途コード", 用途コード="R115", コメント="B")
        self.add(フラグ="用途コード", 用途コード="X", コメント="X")
        self.assertEqual(self.psn(), "板_B,A,")
        self.add(フラグ="寸法納入", 納入先="ﾆｯ*", コメント="C", 厚下="0", 厚上="9", 板幅下="0", 板幅上="9999",
                 板丈下="0", 板丈上="9999")
        self.add(フラグ="寸法取引", 取引先="ﾄﾘ", コメント="D", 厚下="0", 厚上="9", 板幅下="0", 板幅上="9999",
                 板丈下="0", 板丈上="9999")
        self.assertEqual(self.psn(), "板_D,C")                             # B11 再現: C で上書き
        self.assertEqual(self.psn(torihiki=""), "板_C")
        self.assertEqual(self.psn(nounyu=""), "板_D,B,A,")                 # 納入先 "" は寸法納入を見ない
        quirks.set_for_test({"B11": "fix"})
        self.assertEqual(self.psn(), "板_D,C,B,A,")

    def test_materials_message(self) -> None:
        self.add(コメント="直角度の測定をすること")
        self.add(フラグ="用途コード", 用途コード="R115", コメント="注意_B")
        out = self.msgbox(flag=True)
        # "板_注意_B,直角度の測定をすること," → 最後の要素 "B,直角度の測定をすること," → 2個以上の連結
        self.assertEqual([m.text for m in out.msg_errors], [",B,直角度の測定をすること,先行データ票が必要です。"])
        self.assertEqual(out.msg_errors[0].icon, 4)
        self.assertEqual(out.sounds, [judge.SOUND_NEED, judge.SOUND_RIGHT_ANGLE])

    def test_materials_single_comment(self) -> None:
        self.add(コメント="注意A")
        out = self.msgbox(flag=False)
        self.assertEqual([m.text for m in out.msg_errors], ["注意A"])
        self.assertEqual(out.sounds, [])

    def test_materials_no_record(self) -> None:
        out = self.msgbox(flag=True)
        self.assertEqual([m.text for m in out.msg_errors], ["先行データ票が必要です。"])
        self.assertEqual(out.sounds, [judge.SOUND_NEED])
        self.assertEqual(self.msgbox(flag=False).msg_errors, [])

    def test_materials_board_without_comment_b11(self) -> None:
        self.add(フラグ="用途コード", 用途コード="X", コメント="X")          # 合わない → "板_"
        self.assertEqual(self.msgbox(flag=True).msg_errors, [])            # B11 再現: 先行も出ない
        quirks.set_for_test({"B11": "fix"})
        self.assertEqual([m.text for m in self.msgbox(flag=True).msg_errors], ["先行データ票が必要です。"])

    def test_materials_coil(self) -> None:
        self.add(形状="コイル", コメント="注意A")
        out = self.msgbox(flag=True)
        self.assertEqual((out.msg_errors, out.sounds), ([], []))

    def test_materials_db_missing(self) -> None:
        os.remove(self.tmp / "SIKAODR.sqlite3")
        out = self.msgbox(flag=True)
        # オフラインでは MainPath = NextPath なので Case FPath = NextPath で止まる
        self.assertEqual([m.text for m in out.messages], [paths.missing_text("SIKAODRNOW")])

    def test_materials_no_lot(self) -> None:
        self.add(コメント="注意A")
        out = judge.materials_msgbox(spec_form(Lot=""), "ORD1")
        self.assertEqual(out.msg_errors, [])


class CheckForLotTest(unittest.TestCase):
    """CheckForLot(J:2468-2488)"""

    def test_last_five(self) -> None:
        rows = [["", f"LOT{i}"] for i in range(1, 11)]
        self.assertTrue(judge.check_for_lot(rows, "LOT6", 2))
        self.assertFalse(judge.check_for_lot(rows, "LOT5", 2))
        self.assertTrue(judge.check_for_lot(rows, "OT1", 2))          # 部分一致(LOT10)
        self.assertFalse(judge.check_for_lot(rows, "", 2))
        self.assertTrue(judge.check_for_lot([["", "A"]], "A", 2))
        self.assertFalse(judge.check_for_lot([["", None]], "A", 2))


if __name__ == "__main__":
    unittest.main()


# ------------------------------------------------------------------
# 「直す」と決まった不具合(B39〜B42)の直した側
# ------------------------------------------------------------------
class FixedBhTest(QuirkCase):
    def test_b39(self) -> None:
        quirks.set_for_test({"B39": "fix"})
        # 見出し E が "" でも途中の位置では出さない
        f = put(spec_form(), S1="25", U1="25", UBH="30")
        f.captions["E"] = ""
        self.assertEqual(judge.simply_judge(f, "UBH", "NG").msg_errors, [])
        # S だけ検査済みの SBH / E まで検査済みの EBH は今までどおり
        f = put(spec_form(), S1="25", SBH="30")
        f.captions["E"] = ""
        self.assertEqual([m.text for m in judge.simply_judge(f, "SBH", "NG").msg_errors],
                         ["作業中止・保留・ロット異常連絡票発行"])
        f = put(spec_form(), S1="25", U1="25", C1="25", D1="25", E1="25", EBH="30")
        self.assertEqual([m.text for m in judge.simply_judge(f, "EBH", "NG").msg_errors],
                         ["梱包完了させてAインプット・保留・ロット異常連絡票発行"])
        f = put(spec_form(), S1="25", U1="25", C1="25", D1="25", E1="25", DBH="30")
        self.assertEqual(judge.simply_judge(f, "DBH", "NG").msg_errors, [])


class FixedConditionTest(QuirkCase):
    def test_b42(self) -> None:
        quirks.set_for_test({"B42": "fix"})
        self.assertTrue(judge.is_match_condition_field("A,*", "B"))      # "*" は何にでも合う
        self.assertTrue(judge.is_wildcard_match("*", ""))
        self.assertFalse(judge.is_match_condition_field("≠*", "B"))     # 除外の "*" は誰にも当たらない
        self.assertFalse(judge.is_match_condition_field("abc*", "ABCD"))  # ほかは今までどおり


@vba_reproduce
class FixedPackagingTest(PackagingTest):
    def test_b40_empty_table(self) -> None:
        con = sqlite3.connect(self.pack)
        con.execute("DELETE FROM 注意_包装仕様")
        con.commit()
        con.close()
        quirks.set_for_test({"B40": "fix"})
        self.assertEqual(self.psn(), "")
        self.assertEqual(self.msgbox(flag=True).msg_errors[0].text, "先行データ票が必要です。")

    def test_b41_boundary(self) -> None:
        self.add(フラグ="寸法", コメント="寸法注意", 厚下="0.5", 厚上="1.2", 板幅下="900", 板幅上="1100.3",
                 板丈下="1500", 板丈上="2500")
        quirks.set_for_test({"B41": "fix"})
        self.assertEqual(self.psn(thick=1.2), "板_寸法注意")              # 上限ちょうどは範囲内
        self.assertEqual(self.psn(thick=1.2, width=1100.3), "板_寸法注意")
        self.assertEqual(self.psn(thick=1.3), "板_")
        self.assertEqual(self.psn(thick=1.2000001), "板_寸法注意")         # Single では 1.2 と同じ値
        self.assertEqual(self.psn(thick=1.21), "板_")
        quirks.set_for_test(None)
        self.assertEqual(self.psn(thick=1.2, width=1000), "板_")          # 再現


@vba_reproduce
class UnitRecordB13FixTest(unittest.TestCase):
    """B13 を直した側: 区間の間の値・0.1 未満・999 超も必ずどこかの区間に入れる(2026-11-02 の回答)。"""

    def tearDown(self) -> None:
        quirks.set_for_test(None)

    def test_fixed(self) -> None:
        from ufkensa.kensa.modes import sheet
        quirks.set_for_test({"B13": "fix"})
        cases = {0.05: 0.5, 0.1: 0.5, 0.5: 0.5, 0.505: 1.0, 0.51: 1.0, 1.0: 1.0, 1.01: 1.5, 1.009: 1.5,
                 3.51: 4.0, 4.99: 5.0, 5.0: 5.0, -1.01: -1.5}
        for v, want in cases.items():
            self.assertEqual(judge.unit_record(v), want, v)
            self.assertEqual(sheet.unit_record(v), want, v)
        self.assertEqual(judge.unit_record(1500), 1500.0)                 # 999 超はそのまま
        self.assertAlmostEqual(judge.unit_record(12.3), 12.3, places=5)  # 5 超はそのまま

    def test_vba(self) -> None:
        self.assertEqual(judge.unit_record(1.01), 0.0)                    # Single の誤差で区間に入らない
        self.assertEqual(judge.unit_record(1500), 0.0)
