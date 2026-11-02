"""公差外・入力チェック・印刷前チェック(VBA 【UFKen_05】OutTolerance)

期待値は VBA のコードを手で追った結果。根拠の行番号を O:(OutTolerance)で書く。
"""
from __future__ import annotations

import contextlib
import tempfile
import unittest
import zipfile
from pathlib import Path

from tests import kensa_fixture
from ufkensa.kensa import checks, quirks, vbx
from ufkensa.kensa.checks import Uf1, WorkDist

ZERO_DIST = {f"{k}{p}": "0.00" for k in ("Z", "G") for p in "SUCDEF"}


@contextlib.contextmanager
def u1(raises: bool):
    """U1 のスイッチ(vbx.U1_TEXT_VS_NUMBER_RAISES)を一時的に切り替える。"""
    old = vbx.U1_TEXT_VS_NUMBER_RAISES
    vbx.U1_TEXT_VS_NUMBER_RAISES = raises
    try:
        yield
    finally:
        vbx.U1_TEXT_VS_NUMBER_RAISES = old


def make_uf(**kw) -> Uf1:
    """S 行だけ測定済み・入力がそろった画面(印刷前チェックが "" になる)。"""
    text = {"S": "S", "U": "U", "C": "C", "D": "D", "E": "E", "F": "F",
            "S1": "25.01", "S_UpVisual": "欠陥ﾅｼ", "S_UnderVisual": "欠陥ﾅｼ",
            "KenHBox1": "200", "ATU": "25", "Hab": "1385", "Tak": "2770.5",
            "Lot": "RX021T0", "Sumisei": "", "Additional": "", "Youto": "H197"}
    text.update(ZERO_DIST)
    text.update(kw.pop("text", {}))
    return Uf1(text=text, visible=kw.pop("visible", {}), ng=set(kw.pop("ng", set())),
               work=kw.pop("work", "厚板"), line_name=kw.pop("line_name", "Test"),
               line_no=kw.pop("line_no", ""), worker_day_max=kw.pop("worker_day_max", ""))


def dist(a=None, b=None) -> WorkDist:
    """既定: S 行の A・B 両面に定盤がある(他は "__")。"""
    a = a or {}
    b = b or {}
    rows_a = [[a.get(p, ("幅_上_0.30" if p == "S" else "__", "__"))[0], a.get(p, ("", "__"))[1]] for p in "SUCDEF"]
    rows_b = [[b.get(p, ("幅_下_0.20" if p == "S" else "__", "__"))[0], b.get(p, ("", "__"))[1]] for p in "SUCDEF"]
    return WorkDist(rows_a, rows_b)


class QuirkCase(unittest.TestCase):
    def tearDown(self) -> None:
        quirks.set_for_test(None)

    def fix(self, *bugs: str) -> None:
        quirks.set_for_test({b: "fix" for b in bugs})


# ------------------------------------------------------------------
class WeightTest(unittest.TestCase):
    def test_weight(self) -> None:
        # O:816 25×1385×2770.5×2.68/1e6 = 257.0885475 → 下3桁切り捨て
        self.assertEqual(checks.weight(25, 1385, 2770.5), 257.088)
        self.assertEqual(checks.weight("25", "1385", "2770.5"), 257.088)

    def test_not_numeric(self) -> None:
        # O:814-815
        self.assertEqual(checks.weight("", 1385, 2770.5), 0.0)
        self.assertEqual(checks.weight("S", 1, 1), 0.0)
        self.assertEqual(checks.weight(0, 1385, 2770.5), 0.0)

    def test_35kg_boundary(self) -> None:
        # 1000×13060×2.68/1e6 = 35.0008 → 35.000(≤35 = パターンA)、13060.5 → 35.002
        self.assertEqual(checks.weight(1, 1000, 13060), 35.0)
        self.assertEqual(checks.weight(1, 1000, 13060.5), 35.002)


class AbnormalTest(unittest.TestCase):
    """作業保存_Test の値(G2 §10.8): ATU=25、Am=Ap=0.45、HABA=1385、Hm=Hp=2、TAKE=2770.5、Tm=Tp=2.5"""

    def uf(self, name: str, value: str, work: str = "厚板", **extra: str) -> Uf1:
        t = {"Abno": "異常値on", "ATU": "25", "Hab": "1385", "Tak": "2770.5", "Am": "0.45", "Ap": "0.45",
             "Hm": "2", "Hp": "2", "Tm": "2.5", "Tp": "2.5", name: value}
        t.update(extra)
        return Uf1(text=t, work=work)

    def test_limits(self) -> None:
        lim = checks.abnormal_limits(25, 1385, 2770.5, 0.45, 0.45, 2, 2, 2.5, 2.5)
        self.assertEqual((lim.atu1, lim.atu2), (24.1, 25.9))
        self.assertEqual((lim.haba1, lim.haba2), (1381.0, 1389.0))
        self.assertEqual((lim.take1, lim.take2), (2765.5, 2775.5))
        # √(1385²+2770.5²) = 3097.401… ± 6 を Format "0.0"
        self.assertEqual((lim.tai1, lim.tai2), (3091.4, 3103.4))

    def test_thickness_boundary(self) -> None:
        # O:1352-1354 境界ちょうどは正常(Era_flag O:808)
        self.assertIsNone(checks.abnormal_check(self.uf("S1", "24.100"), "S1"))
        hit = checks.abnormal_check(self.uf("S1", "24.099"), "S1")
        self.assertEqual((hit.kind, hit.value, hit.lower, hit.upper), ("厚", "24.099", 24.1, 25.9))
        self.assertIsNone(checks.abnormal_check(self.uf("U4", "25.9"), "U4"))
        self.assertIsNotNone(checks.abnormal_check(self.uf("U4", "25.901"), "U4"))

    def test_width_length_diag(self) -> None:
        self.assertIsNone(checks.abnormal_check(self.uf("S5", "1381.0"), "S5"))
        self.assertEqual(checks.abnormal_check(self.uf("S6", "1380.9"), "S6").kind, "巾")
        self.assertIsNone(checks.abnormal_check(self.uf("S7", "2775.5"), "S7"))
        self.assertEqual(checks.abnormal_check(self.uf("S8", "2775.6"), "S8").kind, "丈")
        self.assertIsNone(checks.abnormal_check(self.uf("S9", "3091.4"), "S9"))
        self.assertEqual(checks.abnormal_check(self.uf("S10", "3103.5"), "S10").kind, "対角")

    def test_gokuatsu(self) -> None:
        # 極厚: 9 と 10 は巾(O:1361, 1367)、11〜14 は丈、対角は見ない(O:1399)
        self.assertEqual(checks.abnormal_check(self.uf("S9", "1380.9", "極厚"), "S9").kind, "巾")
        self.assertEqual(checks.abnormal_check(self.uf("S10", "1389.1", "極厚"), "S10").kind, "巾")
        self.assertEqual(checks.abnormal_check(self.uf("S13", "2760", "極厚"), "S13").kind, "丈")
        self.assertEqual(checks.abnormal_check(self.uf("S7", "9999", "極厚"), "S7").kind, "巾")   # 5 To 9

    def test_skip(self) -> None:
        self.assertIsNone(checks.abnormal_check(self.uf("S1", "1", Abno="異常値off"), "S1"))   # O:1322
        self.assertIsNone(checks.abnormal_check(self.uf("S1", ""), "S1"))                      # O:1348
        self.assertIsNone(checks.abnormal_check(self.uf("S1", "1", Am=""), "S1"))              # O:1353
        # 対角はマイナス公差に丈の Tm を見る(O:1410)
        self.assertIsNone(checks.abnormal_check(self.uf("S9", "1", Tm=""), "S9"))

    def test_form_text(self) -> None:
        t = checks.abnormal_form_text([25, 1385, 2770.5, 0.45, 0.45, 2, 2, 2.5, 2.5])
        self.assertEqual(t["ATU"], "厚 24.100未満 25.900超え")
        self.assertEqual(t["HABA"], "巾 1381.0未満 1389.0超え")
        self.assertEqual(t["TAKE"], "丈 2765.5未満 2775.5超え")
        self.assertEqual(t["TAIK"], "対角線 3091.4未満 3103.4超え")

    def test_ng_back_tb(self) -> None:
        a = checks.ng_back_tb("S1")
        self.assertEqual(a.clear, ("S1",))
        self.assertEqual(a.back_color, (("S1", checks.WINDOW_BG),))
        self.assertTrue(a.clear_inp and a.posting)


class OutToleranceTest(unittest.TestCase):
    def test_atsuita_row(self) -> None:
        # 厚板: Rig 4 で 1〜10 の NG を全部(O:899-913)。DeleteC は Long
        uf = make_uf(text={"INP.Tag": "S4", "S2": "24.0", "S3": "26.1", "DeleteC": "1"}, ng={"S2", "S3"})
        req = checks.out_tolerance(uf)
        self.assertEqual([(r.posi, r.n, r.moto, r.del_count, r.page) for r in req],
                         [("S", 2, "24.0", 1, 0), ("S", 3, "26.1", 1, 0)])

    def test_rig_not_timing(self) -> None:
        uf = make_uf(text={"INP.Tag": "S3"}, ng={"S2"})
        self.assertEqual(checks.out_tolerance(uf), [])
        self.assertEqual(checks.out_tolerance(make_uf(text={"INP.Tag": ""}, ng={"S2"}), "S4"), [])   # O:819

    def test_allow_and_hold(self) -> None:
        # 「許容」なら確認しない、「保留」は確認する(O:903)
        self.assertEqual(checks.out_tolerance(make_uf(text={"INP.Tag": "S4", "NGS": "許容"}, ng={"S2"})), [])
        self.assertEqual(len(checks.out_tolerance(make_uf(text={"INP.Tag": "S4", "NGS": "保留"}, ng={"S2"}))), 1)

    def test_temp_overrides(self) -> None:
        uf = make_uf(text={"INP.Tag": "S1"}, ng={"U5"})
        self.assertEqual([(r.posi, r.n) for r in checks.out_tolerance(uf, "U10")], [("U", 5)])

    def test_diag_moto(self) -> None:
        # 厚板・薄板・少枚数で 9,10 は「線差 」& Diag?(O:689-696)
        uf = make_uf(text={"INP.Tag": "C10", "C9": "3097.0", "DiagC": "0.6"}, ng={"C9"})
        r = checks.out_tolerance(uf)[0]
        self.assertEqual((r.moto, r.return_moto), ("線差 0.6", "線差 0.6"))
        uf.work = "極厚"
        r = checks.out_tolerance(uf)[0]
        self.assertEqual(r.moto, "3097.0")

    def test_work_patterns(self) -> None:
        # 薄板は F 行あり、Rig 2
        uf = make_uf(text={"INP.Tag": "F2"}, ng={"F1"}, work="薄板")
        self.assertEqual([(r.posi, r.n) for r in checks.out_tolerance(uf)], [("F", 1)])
        # スラブは F 行なし
        uf.work = "スラブ"
        self.assertEqual(checks.out_tolerance(uf), [])
        # 極厚(作業未選択も同じ)は 4,7,10,12,14 で 1〜14
        uf = make_uf(text={"INP.Tag": "S12"}, ng={"S11", "S14"}, work="極厚")
        self.assertEqual([r.n for r in checks.out_tolerance(uf)], [11, 14])
        uf.work = ""
        self.assertEqual([r.n for r in checks.out_tolerance(uf)], [11, 14])

    def test_ken_name(self) -> None:
        uf = make_uf(text={"INP.Tag": "S4", "T2": "山田_鈴木"}, ng={"S1"}, worker_day_max="2")
        self.assertEqual(checks.out_tolerance(uf)[0].ken_name, "鈴木")
        uf = make_uf(text={"INP.Tag": "S4", "T1": ""}, ng={"S1"}, worker_day_max="1")
        with self.assertRaises(checks.VbaRuntimeError):            # Split("") → Worker(-1)
            checks.out_tolerance(uf)

    def test_lateral_and_distortion(self) -> None:
        # U1: Rig が "" や "S" のとき `Select Case Rig` の `Case 2, 4, …`。既定(エラーにしない)では一致しないだけ
        with u1(False):
            uf = make_uf(text={"INP.Tag": "LateralS", "LateralS": "2.5"}, visible={"LateralS": True}, ng={"LateralS"})
            self.assertEqual([(r.posi, r.n) for r in checks.out_tolerance(uf)], [("S", 96)])
            uf = make_uf(text={"INP.Tag": "ZS", "ZS": "3.00"}, ng={"ZS"})
            self.assertEqual([(r.posi, r.n) for r in checks.out_tolerance(uf)], [("S", 98)])

    def test_lateral_u1_switch(self) -> None:
        # U1 のスイッチ: 真なら "" と 2 の比較は型エラー13、偽ならエラーにせず進む
        uf = make_uf(text={"INP.Tag": "LateralS"}, visible={"LateralS": True}, ng={"LateralS"})
        with u1(True):
            with self.assertRaises(vbx.TypeMismatch):
                checks.out_tolerance(uf)
        with u1(False):
            self.assertEqual([(r.posi, r.n) for r in checks.out_tolerance(uf)], [("S", 96)])


class ConfiTest(unittest.TestCase):
    def test_decide(self) -> None:
        self.assertEqual(checks.confi_decide("公差外", "0").action, "posting_erase")
        self.assertEqual(checks.confi_decide("許容", "0").action, "pass")
        self.assertEqual(checks.confi_decide("保留", "0").action, "hold")
        self.assertEqual(checks.confi_decide("再測定", "1").action, "back")
        d = checks.confi_decide("再測定", "2")
        self.assertEqual((d.action, d.reasons[0]), ("need_reason", "入力位置間違い"))
        self.assertEqual(checks.confi_decide("x", "0").message.text, "指定文字列を入力してください")

    def test_ng_pass(self) -> None:
        self.assertEqual(checks.ng_pass("S"), {"NGS": "許容"})
        self.assertEqual(checks.ng_pass("F", "保留"), {"NGF": "保留"})

    def test_ng_back(self) -> None:
        uf = make_uf(ng={"S2", "S12", "LateralS"})
        a = checks.ng_back(uf, "S", 96)
        self.assertEqual(a.clear, ("S2", "S12", "LateralS"))      # O:1058-1108
        self.assertEqual(a.erase_count, ("S2", "S12"))
        self.assertEqual(a.look_blank, (1, 10))
        uf = make_uf(ng={"F2"}, work="極厚")
        a = checks.ng_back(uf, "F", 99)
        self.assertEqual((a.clear, a.look_blank), (("F2", "GF"), (1, 14)))

    def test_ng_erase(self) -> None:
        a = checks.ng_erase("U")
        self.assertEqual(a.ranges, ((3, 1, 3, 20), (11, 1, 11, 3), (9, 15, 9, 16)))   # O:983, 990, 997
        self.assertEqual(a.split_z_tag, "U")
        self.assertFalse(checks.ng_erase("U", "板廃棄").close_confi)

    def test_reason_open(self) -> None:
        r = checks.confi_reason_open(make_uf(), "S", 2, "Test", "24.0", 1)
        self.assertEqual((r.page, r.return_moto, r.ken_name), (1, None, None))

    def test_reasons(self) -> None:
        self.assertEqual(checks.DELETE_REASONS[-1], "その他")
        self.assertEqual(len(checks.DELETE_REASONS), 7)


class NgPostingTest(unittest.TestCase):
    def test_posting(self) -> None:
        uf = make_uf(text={"S1": "24.0", "S2": "25.0", "DiagSKou": "0.5", "LateralS": "1.0",
                           "ZS": "1.00", "GS": "2.00", "AorB": "Now_B"}, ng={"S1", "ZS"})
        sh = checks.ng_posting(uf, "S", None)
        row = sh.rows[0]
        self.assertEqual(sh.lot, "RX021T0")
        self.assertEqual(row.values[:2], ["24.0", "25.0"])
        self.assertEqual(row.values[14:19], ["0.5", "1.0", "B:1.00", "ｹﾞ:2.00", "S"])
        self.assertEqual(row.ng_cols, {1, 17})
        sh = checks.ng_posting(uf, "S", sh)
        self.assertEqual(len(sh.rows), 2)                                 # 同じロットは下へ足す

    def test_overwrite_when_col1_blank(self) -> None:
        # UR(sh, 1)+1(O:1177): 前の行の A 列が空だと、その行に上書きする(色は残る)
        uf = make_uf(text={"S1": "", "S5": "1380"}, ng={"S5"})
        sh = checks.ng_posting(uf, "S", None)
        uf2 = make_uf(text={"S1": "24.0"}, ng={"S1"})
        sh = checks.ng_posting(uf2, "S", sh)
        self.assertEqual(len(sh.rows), 1)
        self.assertEqual(sh.rows[0].values[0], "24.0")
        self.assertEqual(sh.rows[0].ng_cols, {1, 5})

    def test_other_lot_resets(self) -> None:
        uf = make_uf(text={"S1": "24.0"})
        sh = checks.ng_posting(uf, "S", checks.AbnormalSheet("ZZZ", [checks.NgRow(["x"] * 19)]))
        self.assertEqual((sh.lot, len(sh.rows)), ("RX021T0", 1))

    def test_f_row(self) -> None:
        uf = make_uf(text={"F1": "24", "F10": "9"})
        row = checks.ng_posting(uf, "F", None).rows[0]
        self.assertEqual((row.values[0], row.values[9], row.values[10], row.values[18]), ("24", "9", "", "F"))


class ListsTest(unittest.TestCase):
    def test_cb(self) -> None:
        self.assertEqual(checks.cb_shape_items(6.0), ["幅", "丈", "耳", "中"])   # O:1447 6.00 以下
        self.assertEqual(checks.cb_shape_items(6.01), ["幅", "丈"])
        self.assertEqual(checks.cbg_shape_items({"LGF": "ﾌﾘｰｿﾘ"}), ["上", "下"])
        self.assertEqual(checks.cbg_shape_items({"LGS": "ｹﾞｰｼﾞ:"}), ["幅", "丈"])
        self.assertEqual(checks.cb_updown_items("耳"), ["Ⅱ型", "Ⅲ型"])
        self.assertEqual(checks.cb_updown_items("中"), ["Ⅰ型", "Ⅱ型", "Ⅲ型"])
        self.assertEqual(checks.cbg_updown_items("中"), [])

    def test_selected_items(self) -> None:
        items = ["", "当り跡ﾞ", "扱いｷｽﾞ"]
        self.assertEqual(checks.get_selected_items(items, [1, 2]), "当り跡ﾞ 扱いｷｽﾞ")
        self.assertEqual(checks.get_selected_items(items, [0, 2]), "扱いｷｽﾞ")     # 1 件目の空白行は飛ばす
        self.assertEqual(checks.get_selected_items(items, [1, 0][::-1]), "当り跡ﾞ")
        self.assertEqual(checks.get_selected_items(["a", ""], [0, 1]), "a ")       # 2 件目以降の空白は " "
        d = checks.defect_factor_items(["汚", "TP"], ["12", None])
        self.assertEqual(d[-1], (checks.SPOT, ""))
        self.assertEqual(checks.get_selected_items(d, [1], defect=True), "TP")     # "TP_" → 最後の "_" を消す
        self.assertEqual(checks.get_selected_items(d, [2], defect=True), checks.SPOT)
        self.assertEqual(checks.get_selected_items(d, [0, 1], defect=True), "汚_12 TP")

    def test_lc_color(self) -> None:
        r = checks.lc_color(8, 7, "ｷｽﾞ", "ﾄｯﾄ", "鈴木", cyan_before={1})
        self.assertEqual(r.cyan, frozenset({8}))
        self.assertEqual(r.captions, {"U_UnderVisual": "ｷｽﾞ"})
        self.assertEqual(r.log_text, "*UFKenlU_UnderVisuallｷｽﾞlﾄｯﾄ|鈴木")
        self.assertEqual(r.log_folder, "トット")

    def test_tec_return(self) -> None:
        self.assertEqual(checks.tec_return({"SBad", "DBad"}), "DBad")
        self.assertEqual(checks.tec_return(set()), "")


class VisualAddTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        kensa_fixture.use_offline()

    def test_from_db(self) -> None:
        v = checks.visual_add(["フクレ", "　"], ["01", "02"])
        self.assertEqual(v.visual[:2], ["", "当り跡ﾞ"])               # 先頭は空白行(O:1569)
        self.assertEqual(len(v.visual), 76)
        self.assertEqual(v.visual[-1], "ｶｯﾌﾟﾙ跡")
        self.assertEqual(v.yama, ["1", "2", "3", "4", "5", "6", "枚数"])
        self.assertEqual(v.defect, [("フクレ", "01"), ("　", "02"), ("ｽﾎﾟｯﾄ処理", "")])
        self.assertIsNone(checks.visual_add().defect)                 # 原紙が無い(O:1591)


class SheetsTest(unittest.TestCase):
    def test_calculate(self) -> None:
        # 作業保存_Test の J9「(200.0÷25) + (226.0÷25)=17.04」
        r = checks.calculate_sheets(["200.0", "226.0", "", "", "", ""], 25)
        self.assertEqual(r.formula, "(200.0÷25) + (226.0÷25)=17.04")
        self.assertEqual(r.quantity, vbx.single(17.04))

    def test_single_thickness(self) -> None:
        r = checks.calculate_sheets(["100"], 0.3)
        self.assertEqual(r.formula, "(100÷0.3)=333.3333")
        self.assertEqual(r.quantity, vbx.single(333.3333))

    def test_nothing(self) -> None:
        self.assertEqual(checks.calculate_sheets(["200"], 0), checks.SheetsResult(0.0))
        self.assertEqual(checks.calculate_sheets(["", "x"], 25), checks.SheetsResult(0.0))
        self.assertEqual(checks.calculate_sheets([""], 25, True, "17"), checks.SheetsResult(0.0, None, "17"))

    def test_h_check(self) -> None:
        self.assertEqual(checks.h_check([""] * 7), "入力なし")
        self.assertEqual(checks.h_check([""] * 6 + ["12"]), "")

    def test_yama_layout(self) -> None:
        y = checks.yama_layout("2")
        self.assertEqual(y.hl_captions, {"HL1": "2-2", "HL2": "2-1"})
        self.assertTrue(y.visible["KenHBox2"] and not y.visible["KenHBox3"])
        self.assertEqual(checks.yama_layout("1").hl_captions, {"HL1": "1"})
        y = checks.yama_layout("枚数", mitsubishi=True)
        self.assertTrue(y.visible["KenHBox7"] and y.visible["KenHL7"])
        self.assertNotIn("KenCount", y.visible)
        self.assertIsNone(checks.yama_layout(""))


class GaugeTest(QuirkCase):
    def heavy(self, **text) -> Uf1:      # 257kg(パターンB)
        return make_uf(text=text)

    def light(self, **text) -> Uf1:      # 35.000kg(パターンA)
        t = {"ATU": "1", "Hab": "1000", "Tak": "13060"}
        t.update(text)
        return make_uf(text=t)

    def test_ok(self) -> None:
        self.assertEqual(checks.withor_without_gauge(self.heavy(), dist()), "")
        self.assertEqual(checks.withor_without_gauge(self.heavy(ATU=""), dist(a={"S": ("", "")})), "")  # 重量 0
        self.assertEqual(checks.withor_without_gauge(self.heavy(), None), "")                           # シートなし

    def test_pattern_a(self) -> None:
        d = dist(b={"S": ("__", "__")})
        self.assertEqual(checks.withor_without_gauge(self.light(), d), "SのB面 定盤 未入力_ﾊﾟﾀｰﾝA")
        d = dist(a={"S": ("_", "")})
        self.assertEqual(checks.withor_without_gauge(self.light(), d), "SのA面 定盤 未入力_ﾊﾟﾀｰﾝA")
        # 35.002kg はパターンB(B 面だけでゲージなし → 後のループの文言が残る)
        self.assertEqual(checks.withor_without_gauge(self.light(Tak="13060.5"), d), "SのB面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB")

    def test_pattern_b(self) -> None:
        d = dist(b={"S": ("__", "__")})
        self.assertEqual(checks.withor_without_gauge(self.heavy(), d), "SのA面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB")
        d = dist(b={"S": ("__", "__")}, a={"S": ("幅_上_0.30", "幅_上_1.00")})
        self.assertEqual(checks.withor_without_gauge(self.heavy(), d), "")

    def test_pattern_b_exit(self) -> None:
        # S が両面、U が A のみゲージなし → S で Exit Function(O:1952)
        d = dist(a={"U": ("幅_上_0.30", "__")})
        self.assertEqual(checks.withor_without_gauge(self.heavy(U1="25.0"), d), "")

    def test_b10_c_row(self) -> None:
        d = dist(a={"S": ("__", "__")}, b={"S": ("__", "__")})
        uf = self.heavy(S1="", C1="25.0")
        self.assertEqual(checks.withor_without_gauge(uf, d), "2のA面 定盤 未入力")       # O:1914
        self.fix("B10")
        self.assertEqual(checks.withor_without_gauge(uf, d), "CのA面 定盤 未入力")

    def test_b10_f_row(self) -> None:
        # F は B 面だけ・ゲージなし(Yama=4)。VBA の B 面ループは 1 To 5(O:1973)
        d = dist(a={"S": ("__", "__")}, b={"S": ("__", "__"), "F": ("幅_下_0.20", "__")})
        uf = self.heavy(S1="", F1="25.0", Yama="4")
        self.assertEqual(checks.withor_without_gauge(uf, d), "FのA面 定盤 未入力")
        self.fix("B10")
        self.assertEqual(checks.withor_without_gauge(uf, d), "FのB面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB")

    def test_gauge_judgment(self) -> None:
        d = dist(b={"S": ("__", "__")})
        self.assertEqual(checks.gauge_judgment(self.light(), "S", d), "SのB面 定盤 未入力_パターンA")
        self.assertEqual(checks.gauge_judgment(self.heavy(), "S", d), "SのA面 ゲージ 未入力_パターンB")
        self.assertEqual(checks.gauge_judgment(self.heavy(C1="1"), "C", d), "CのA面 定盤 未入力")
        self.assertEqual(checks.gauge_judgment(self.heavy(), "X", d), "")
        self.assertEqual(checks.gauge_judgment(self.heavy(F1="1"), "F", d), "")         # Yama≠4
        # 形が「文字_文字_数値」でない値は入力なし扱い
        d = dist(a={"S": ("幅_上", "")})
        self.assertEqual(checks.gauge_judgment(self.heavy(), "S", d), "SのA面 定盤 未入力")

    def test_is_valid_format(self) -> None:
        self.assertTrue(checks.is_valid_format("幅_上_0.30"))
        self.assertTrue(checks.is_valid_format("__0.00"))         # 空があっても 0.00 なら次へ
        self.assertFalse(checks.is_valid_format("幅__0.30"))
        self.assertFalse(checks.is_valid_format("1_2_上"))
        self.assertTrue(checks.is_valid_format("幅_上_1/2"))
        self.assertFalse(checks.is_valid_format("幅_上_ab"))
        self.assertFalse(checks.is_valid_format(""))
        self.assertFalse(checks.is_valid_format("a_b"))

    def test_check_labels(self) -> None:
        vis = {f"S{i}": True for i in range(1, 3)}
        uf = make_uf(text={"S2": "1"}, visible=vis)
        # 非表示の欄は白にするが、最後に「可」があれば緑(O:2740)
        self.assertEqual(checks.check_labels_by_category(uf, "S", dist()), checks.GREEN)
        self.assertEqual(checks.check_labels_by_category(make_uf(), "S", dist()), checks.WHITE)
        uf = make_uf(text={"S2": ""}, visible=vis)
        self.assertEqual(checks.check_labels_by_category(uf, "S", dist()), checks.RED)
        vis.update({"S_UpVisual": True, "S_UnderVisual": True})
        uf = make_uf(text={"S2": "1"}, visible=vis)
        self.assertEqual(checks.check_labels_by_category(uf, "S", dist()), checks.GREEN)
        self.assertEqual(checks.check_labels_by_category(uf, "S", dist(b={"S": ("__", "__")})), checks.RED)
        # F は F11 が無いので番号の欄はそこで打ち切り
        uf = make_uf(text={"F1": "1"}, visible={"F1": True})
        self.assertEqual(checks.check_labels_by_category(uf, "F", dist()), checks.GREEN)
        self.assertEqual(checks.check_labels_by_category(uf, "X", dist()), checks.RED)
        self.assertEqual(checks.get_label_state(uf, "F11"), "無効")
        self.assertEqual(checks.get_label_state(uf, "S14"), "すべて非表示")


class DistortionTest(QuirkCase):
    def test_zero_clears_shape(self) -> None:
        uf = make_uf(text={"CBS1": "幅", "CBS2": "上", "CBF1": "幅", "CBF2": "上", "CBS3": "幅"})
        r = checks.distortion_check(uf, "Z")
        self.assertEqual(r.text, "")
        self.assertIn("CBS1", r.cleared)
        self.assertIn("CBS3", r.cleared)
        self.assertNotIn("CBF1", r.cleared)                     # ZF → CBF1 は無い(O:1989-1993)

    def test_z_missing(self) -> None:
        uf = make_uf(text={"ZU": "1.00", "CBU1": "幅"})
        self.assertEqual(checks.distortion_check(uf, "Z").text, "定盤:U")
        uf = make_uf(text={"ZU": "1.00", "CBU1": "幅", "ZD": "0.50", "CBD2": "上"})
        self.assertEqual(checks.distortion_check(uf, "Z").text, "定盤:D")      # 後の位置が残る

    def test_b9(self) -> None:
        # F 行は ZE を見る(O:2022)。ZE=0 なので F の未入力に気付かない
        uf = make_uf(text={"ZF": "1.00", "CBF1": "幅"})
        self.assertEqual(checks.distortion_check(uf, "Z").text, "")
        self.fix("B9")
        self.assertEqual(checks.distortion_check(uf, "Z").text, "定盤:F")

    def test_gauge_and_free(self) -> None:
        uf = make_uf(text={"GC": "1.00", "CBC3": "幅"})
        self.assertEqual(checks.distortion_check(uf, "G").text, "ｹﾞｰｼﾞ:C")
        uf = make_uf(text={"GC": "1.00", "CBC3": "上", "flatnesText": "ﾌﾘｰｿﾘ要"})
        self.assertEqual(checks.distortion_check(uf, "G").text, "")             # 2 項目そろっている
        uf = make_uf(text={"GC": "1.00", "flatnesText": "ﾌﾘｰｿﾘ要"})
        self.assertEqual(checks.distortion_check(uf, "G").text, "ﾌﾘｰｿﾘ:C")
        uf = make_uf(text={"GC": "1.00", "flatnesText": "その他"})
        self.assertEqual(checks.distortion_check(uf, "G").text, "")

    def test_empty_value(self) -> None:
        # U1: ZS が空のとき `.ZS = 0`(O:2003)。既定(エラーにしない)は「数 < 文字列」で 0 と一致しない
        # → CBS1 は消されず、形状あり・上下と値が空・値が 0 でない で「定盤:S」(O:2030-2036)
        uf = make_uf(text={"ZS": "", "CBS1": "幅"})
        with u1(False):
            self.assertEqual(checks.distortion_check(uf, "Z").text, "定盤:S")
        with u1(True):
            with self.assertRaises(vbx.TypeMismatch):
                checks.distortion_check(uf, "Z")

    def test_single_unit(self) -> None:
        uf = make_uf(text={"ZS": "0", "CBS1": "幅"})
        self.assertEqual(checks.single_unit_distortion(uf, "Z", "S"), "定盤:S")
        uf = make_uf(text={"ZS": "0"})
        self.assertEqual(checks.single_unit_distortion(uf, "Z", "S"), "")
        uf = make_uf(text={"GS": "", "CBS3": "幅"})
        self.assertEqual(checks.single_unit_distortion(uf, "G", "S"), "ｹﾞｰｼﾞ:S")


class PrecheckTest(QuirkCase):
    def test_ok(self) -> None:
        self.assertEqual(checks.print_precheck(make_uf(), dist()).text, "")

    def test_order(self) -> None:
        # 定盤とゲージ両方 → 定盤が先
        uf = make_uf(text={"ZS": "1.00", "GS": "1.00", "CBS1": "幅", "CBS3": "幅", "S_UpVisual": ""})
        self.assertEqual(checks.print_precheck(uf, dist()).text, "定盤:S入力なし")
        uf = make_uf(text={"GS": "1.00", "CBS3": "幅", "S_UpVisual": ""})
        self.assertEqual(checks.print_precheck(uf, dist()).text, "ｹﾞｰｼﾞ:S入力なし")
        uf = make_uf(text={"S_UpVisual": ""})
        self.assertEqual(checks.print_precheck(uf, dist(b={"S": ("__", "__")})).text,
                         "SのA面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB")
        self.assertEqual(checks.print_precheck(uf, dist()).text, "S:外観入力なし")
        uf = make_uf(text={"KenHBox1": ""}, visible={"Yama": True})
        self.assertEqual(checks.print_precheck(uf, dist()).text, "検入高さor枚数　入力なし")
        uf = make_uf(visible={"Yama": True})
        self.assertEqual(checks.print_precheck(uf, dist()).text, "山数の入力がありません")

    def test_angle_r115(self) -> None:
        uf = make_uf(visible={"DiagSKou": True})
        self.assertEqual(checks.print_precheck(uf, dist()).text, "")
        uf.text["Youto"] = "R115"
        self.assertEqual(checks.print_precheck(uf, dist()).text, "S:直角度入力なし")

    def test_lateral_b2(self) -> None:
        uf = make_uf(visible={"LateralS": True})
        # B2: Sumisei が "" でも止める(`<> """"`)
        self.assertEqual(checks.print_precheck(uf, dist()).text, "S:ﾗﾃﾗﾙﾎﾞｰ入力なし")
        uf.text["Additional"] = "ﾗﾃﾗﾙﾎﾞｰ検査不要"
        self.assertEqual(checks.print_precheck(uf, dist()).text, "")
        uf.text["Additional"] = ""
        uf.text["Sumisei"] = '"'
        self.assertEqual(checks.print_precheck(uf, dist()).text, "")
        self.fix("B2")
        uf.text["Sumisei"] = ""
        self.assertEqual(checks.print_precheck(uf, dist()).text, "")
        uf.text["Sumisei"] = "住友精密"
        self.assertEqual(checks.print_precheck(uf, dist()).text, "S:ﾗﾃﾗﾙﾎﾞｰ入力なし")

    def test_row_checks(self) -> None:
        uf = make_uf(text={"S1": "", "U1": "1", "LateralU": ""}, visible={"LateralU": True, "LateralS": True})
        self.assertEqual(checks.lateralbow_check(uf), "U")
        uf = make_uf(text={"F1": "1"}, visible={"DiagFKou": True})
        self.assertEqual(checks.angle_check(uf), "")              # Yama≠4 は F を見ない
        uf.text["Yama"] = "4"
        self.assertEqual(checks.angle_check(uf), "F")
        self.assertEqual(checks.visual_check(make_uf(text={"C1": "1"})), "C")

    def test_bh_check(self) -> None:
        uf = make_uf(text={"QF": "QF有り", "U1": "1", "SBH": "", "UBH": ""})
        self.assertEqual(checks.bh_check(uf), "U")                 # countL(2): S のあと U で上書き
        uf = make_uf(text={"QF": "QF有り", "SBH": ""})
        self.assertEqual(checks.bh_check(uf), "S")
        uf = make_uf(text={"QF": "", "SBH": ""})
        self.assertEqual(checks.bh_check(uf), "")


class AllowCheckTest(QuirkCase):
    def test_b34(self) -> None:
        uf = make_uf(text={"NGS": "許容", "S": "1"})
        self.assertEqual(checks.allow_check(uf), "1")              # 「許容」が付かない(O:2615-2616)
        uf = make_uf(text={"NGS": "許容", "NGU": "保留", "S": "1", "U": "2"})
        self.assertEqual(checks.allow_check(uf), "1, 2 保留")
        self.fix("B34")
        self.assertEqual(checks.allow_check(make_uf(text={"NGS": "許容", "S": "1"})), "1 許容")
        self.assertEqual(checks.allow_check(uf), "1 許容, 2 保留")

    def test_none(self) -> None:
        self.assertEqual(checks.allow_check(make_uf()), "")


def _write_xlsx(path: Path, cells: dict[str, object]) -> None:
    rows: dict[int, list[str]] = {}
    for ref, v in cells.items():
        r = int("".join(ch for ch in ref if ch.isdigit()))
        if isinstance(v, str):
            x = f'<c r="{ref}" t="inlineStr"><is><t>{v}</t></is></c>'
        else:
            x = f'<c r="{ref}"><v>{v}</v></c>'
        rows.setdefault(r, []).append(x)
    data = "".join(f'<row r="{r}">{"".join(cs)}</row>' for r, cs in sorted(rows.items()))
    m = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rr = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("xl/workbook.xml", f'<workbook xmlns="{m}" xmlns:r="{rr}"><sheets>'
                   f'<sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels",
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr("xl/worksheets/sheet1.xml", f'<worksheet xmlns="{m}"><sheetData>{data}</sheetData></worksheet>')


class ConductivTest(QuirkCase):
    def test_check(self) -> None:
        self.assertTrue(checks.conductiv_check(make_uf(text={"Zai": "75S", "Tyu": "T651"})))
        self.assertTrue(checks.conductiv_check(make_uf(text={"Zai": "14S", "Tyu": "T6"})))
        self.assertFalse(checks.conductiv_check(make_uf(text={"Zai": "17S", "Tyu": "T3"})))   # コメントアウト済み
        self.assertFalse(checks.conductiv_check(make_uf(text={"Zai": "75S", "Tyu": "T7351"})))

    def test_lookup(self) -> None:
        rows = [[12.0, "R1"], ["", "R2"], ["A-3", "R1"], [None, "R3"]]
        self.assertEqual(checks.coolingtime_lookup(rows, "R1"), "A-3")       # 最後に一致した行
        self.assertEqual(checks.coolingtime_lookup(rows[:1], "R1"), "12")
        self.assertEqual(checks.coolingtime_lookup(rows, "R2"), "")
        self.assertEqual(checks.coolingtime_lookup(rows, "R9"), "LotNo記載がない")
        self.fix("B12")
        self.assertEqual(checks.coolingtime_lookup(rows, "R9"), "LotNoがない")

    def test_xlsx(self) -> None:
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "list.xlsx"
            _write_xlsx(p, {"A4": "写真確認No", "B4": "LotNo", "A5": 3, "B5": "RX021T0", "B6": "R6550H0"})
            self.assertEqual(checks.conductiv_coolingtime_check("RX021T0", p).value, "3")
            self.assertEqual(checks.conductiv_coolingtime_check("R6550H0", p).value, "")
            self.assertEqual(checks.conductiv_coolingtime_check("X", p).value, "LotNo記載がない")
            r = checks.conductiv_coolingtime_check("X", Path(d) / "none.xlsx")
            self.assertEqual(r.value, "")
            self.assertTrue(r.messages[0].text.endswith("QF空冷時間チェックリスト.xlsxがありません_Conductiv_Coolingtime_Check"))

    def test_sheet_values(self) -> None:
        uf = make_uf(text={"Zai": "75S", "Tyu": "T6", "ATU": "6.75"}, line_name="LVC")
        v = checks.conductiv_sheet_values("2026/10/30", uf, "山田_鈴木")
        self.assertEqual((v["B6"], v["B7"], v["B8"], v["B9"]), ("75S-T6", "6.75×1385×2770.5", "LVC", "鈴木"))


class SelectionInputTest(unittest.TestCase):
    CODES = ["01", "02", "11", "12", "13", "14", "21", "22", "23", "24", "25", "31", "32", "33", "34", "35",
             "36", "37", "38", "39", "41", None]

    def test_write(self) -> None:
        w = checks.selection_input("汚_12", "UBad", "3", self.CODES, make_uf())
        self.assertEqual(w.cells[0], (1, 2, "RX021T0"))
        self.assertIn((3, 2, "'S"), w.cells)
        self.assertEqual(w.cells[-1], (4, 6, "3"))                 # b_Arr(2, 4) → 行 4・列 F
        w = checks.selection_input("TP", "FBad", "1", self.CODES, None, "LOT")
        self.assertEqual(w.cells, ((1, 2, "LOT"), (8, 24, "1")))     # TP は 22 列目(X)
        w = checks.selection_input(checks.SPOT, "SBad", "5", self.CODES, None, "LOT")
        self.assertEqual(w.cells[-1], (11, 3, "5枚"))
        self.assertIsNone(checks.selection_input("", "SBad", "5", self.CODES))

    def test_not_found(self) -> None:
        # b_Arr(a, 0) で添字エラー(O:2435)。itti=0 のやり直しには届かない
        with self.assertRaises(checks.VbaRuntimeError):
            checks.selection_input("汚_12 TP", "SBad", "1", self.CODES)


class SmallNumberTest(unittest.TestCase):
    ROWS = [["滑川", "白銅", "R115", "52S", "H34", "6", "1250", "2500", "L1", "K1"],
            ["滑川", "白銅", "R115", "52S", "H34", "6", "1250", "2500", "L1", "K1"],
            ["日車", "", "R115", "52S", "H34", "6.5", "1250", "2500", "L1", "K1"]]

    def test_count4_lateral(self) -> None:
        calls = []

        def dbr(*a):
            calls.append(a)
            return "ラテラルボー,直角度" if a[0] == "日車" else ""

        uf = make_uf(text={"Youto": "R115"})
        ch = checks.small_number_count(4, uf, self.ROWS, dbr)
        # 同じ納入先は呼ばない。用途〜丈は最後の行の値(厚 6.5)
        self.assertEqual([c[0] for c in calls], ["滑川", "日車"])
        self.assertEqual(calls[0][5], 6.5)
        self.assertTrue(ch.lateral)
        v = ch.visible
        self.assertEqual((v["LateralS"], v["LateralU"], v["LateralC"], v["LateralD"], v["LateralE"]),
                         (True, False, True, True, False))
        self.assertFalse(v["DiagEKou"] or v["DiagU"] or v["E1"] or v["U5"] or v["S11"])
        self.assertTrue(v["U4"] and v["DiagDKou"] and v["S10"])
        self.assertEqual((ch.values["E"], ch.values["OPS5"], ch.values["D"]), ("", "", "4"))

    def test_count5_and_no_lateral(self) -> None:
        uf = make_uf(text={"Additional": "ﾗﾃﾗﾙﾎﾞｰ検査不要"})
        ch = checks.small_number_count(5, uf, self.ROWS, lambda *a: "ラテラルボー")
        self.assertFalse(ch.lateral)
        self.assertNotIn("LateralS", ch.visible)                    # 触らない
        self.assertFalse(ch.visible["DiagU"] or ch.visible["U5"])
        self.assertTrue(ch.visible["U4"] and ch.visible["DiagC"])
        ch = checks.small_number_count(5, make_uf(), None, lambda *a: "ラテラルボー")
        self.assertFalse(ch.lateral)

    def test_count1(self) -> None:
        ch = checks.small_number_count(1, make_uf(), self.ROWS, lambda *a: "ラテラルボー")
        self.assertEqual({k: ch.visible[k] for k in ("LateralS", "LateralU", "LateralE")},
                         {"LateralS": True, "LateralU": False, "LateralE": False})
        self.assertEqual((ch.values["S"], ch.values["U"], ch.values["OPE2"]), ("1", "", ""))
        self.assertNotIn("DiagUKou", ch.visible)                    # R115 でなければ直角度は触らない

    def test_bad_thickness(self) -> None:
        rows = [["A", "", "x", "", "", "", "1", "1"]]
        with self.assertRaises(vbx.TypeMismatch):                   # "" を Single に入れる(型エラー13)
            checks.small_number_count(3, make_uf(), rows, lambda *a: "")


class MiscTest(unittest.TestCase):
    def test_find_right_string(self) -> None:
        self.assertEqual(checks.find_right_string("HOT ANI CR CAF CR"), "CAF")
        self.assertEqual(checks.find_right_string("HOT CR"), "")

    def test_uftag(self) -> None:
        t = checks.uftag_array({"CBS1": "幅", "CBS2": "上", "ZS": "0.30", "GS": "1.0"}, 1, 2)
        self.assertEqual(t[1], ("幅_上_0.30", "__1.0"))
        self.assertEqual(t[2], ("__", "__"))

    def test_msg_error(self) -> None:
        m = checks.msg_error("ABC,  ,DEF", "ﾊﾞﾗﾝｻｰ", 4, t123=("a", "", ""))
        self.assertEqual(m.message.text, "ABC\nDEF")
        self.assertEqual(m.message.title, "注意")
        self.assertEqual(m.log_folder, "バランサー")
        self.assertEqual(m.log_text, "*UFKen:MsgError:ABC,  ,DEF T1:a T2: T3:")
        self.assertFalse(m.end)
        m = checks.msg_error("止めます", "丸徳A1", 1, "END")
        self.assertEqual((m.message.text, m.message.title, m.log_folder, m.end), ("止めます", "警告", "丸徳", True))

    def test_log_folder(self) -> None:
        self.assertEqual(checks.log_folder("ﾄｯﾄ"), "トット")
        self.assertEqual(checks.log_folder("トット"), "丸徳")          # VBA の Select Case は完全一致

    def test_pack_picture(self) -> None:
        self.assertTrue(checks.pack_picture_path("1P0113").endswith("包装仕様書\\1P0113.JPG"))


if __name__ == "__main__":
    unittest.main()
