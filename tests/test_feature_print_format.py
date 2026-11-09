"""ラテラルボー有無 × フリー反り有無 で検査成績票の書式が切り替わることを Session の操作どおりに確かめる。

VBA(【UFKen_08】Print.bas 印刷用罫線 skipEnd 以降):
- Print:1043  (Sumisei<>"" And NgTemp<>"検査不要") Or FlagColor Or LateralText.Tag="有" → ラテラルボーの列を残す
    Print:1052 ラテ="" なら J18:N24 を一旦消し、Print:1058 P 列(ﾗﾃﾗﾙﾎﾞｰ)を J 列へ写して K:N を消す
    (R115 なら Print:1046 で P→K・R(直角度)→J)
- Print:1083 Else: J18:N23(4 山は N24 まで)の罫線と値を消す = ラテラルボーの列なし
- Print:1127 LZS="幅ｹﾞｰｼﾞ" → (11,5)="幅ゲージ" / (11,6)="丈ゲージ"
- Print:1130 flatnesText="ﾌﾘｰｿﾘ要" → (11,6)="ﾌﾘｰｿﾘ"、F12:F16 の "A:" "B:" を消す
Python: ufkensa/kensa/report_model.py _lateral_columns (1275-1342)。

合成データ(一時コピーの 品質基準.ラテフリ直角 にだけ入れる): H7345B0 の納入先 × 厚15/巾1252/丈2502 に当たる 1 行。
"""
from __future__ import annotations

import sqlite3

from tests.test_kensa_session import SessionBase
from ufkensa.kensa import report_model

DELIVERY = "ｸﾉｻﾝｷﾞﾖｳ(ｶ ﾎﾝｼﾔ"
LATERAL_BLOCK = [c + str(r) for r in range(18, 25) for c in "JKLMN"]
COMMON = [c + str(r) for r in range(11, 25) for c in "BCDEGHI"]   # 切り替わらない欄(F 列とラテラル列以外)


class PrintFormatBase(SessionBase):
    lateral = False
    free = False

    def setUp(self) -> None:
        super().setUp()
        if self.lateral or self.free:
            con = sqlite3.connect(self.db_dir / "品質基準.sqlite3")
            con.execute("INSERT INTO ラテフリ直角 (管理番号, 納入先, 厚下, 厚上, 幅下, 幅上, 丈下, 丈上,"
                        " ラテラルボー, フリー反り, 公差) VALUES (9001, ?, 10, 20, 1000, 1500, 2000, 3000, ?, ?, ?)",
                        (DELIVERY, "要" if self.lateral else "", "要" if self.free else "",
                         "5" if self.lateral else "0.5"))
            con.commit()
            con.close()

    def printed(self) -> report_model.ReportModel:
        s = self.s
        self.register()
        self.fill("DT")
        self.fill("#L")
        self.strain_visual_height()
        for p in "SUCDE":
            if s.state.ctl("Lateral" + p).visible:
                self.ok(s.set_value("Lateral" + p, "1.0"))
        self.ok(s.min_max())
        self.assertTrue(self.ok(s.print_check()).view["print_ready"])
        cap, orig = [], report_model.print_transfer

        def spy(*a, **k):
            res = orig(*a, **k)
            cap.append(res)
            return res
        report_model.print_transfer = spy
        try:
            self.ok(s.print_commit({"print": True}))
        finally:
            report_model.print_transfer = orig
        self.assertEqual(len(cap), 1)
        return cap[0].report

    @staticmethod
    def val(m, ref):
        c = m.cell(ref)
        return c["value"] if c else None

    @staticmethod
    def bd(m, ref):
        c = m.cell(ref)
        return {k: tuple(v) for k, v in c["borders"].items()} if c else None

    # -- 期待値 ------------------------------------------------------
    def check(self, m) -> None:
        st = self.s.state
        self.assertEqual(st.flatnes_text, "ﾌﾘｰｿﾘ要" if self.free else "")
        self.assertEqual(self.s.state.ctl("L11").visible, self.lateral)
        # (11,5) は定盤のまま、(11,6) は ﾌﾘｰｿﾘ/ゲージ歪 (Print:1127-1131)
        self.assertEqual(self.val(m, "E11"), "定盤歪")
        self.assertEqual(self.val(m, "F11"), "ﾌﾘｰｿﾘ" if self.free else "ゲージ歪")
        for r in range(12, 17):                                  # A:/B: の表記 (Print:1132-1133)
            v = self.val(m, f"F{r}")
            if self.free:
                self.assertFalse(v.startswith(("A:", "B:")), v)
                self.assertEqual(v, "幅上0.20\n")
            else:
                self.assertEqual(v, "A:幅上0.20\n")
        # ラテラルボーの列 (Print:1043-1080 / 1083-1095)
        if self.lateral:
            self.assertEqual(self.val(m, "J18"), "ﾗﾃﾗﾙﾎﾞｰ")
            self.assertEqual(self.bd(m, "J18")["left"], ("continuous", "medium"))
            self.assertEqual(self.bd(m, "J18")["bottom"], ("double", "thick"))
            for r in (19, 21, 23):                               # 厚板 S/2/E に値
                self.assertEqual(self.val(m, f"J{r}"), "1.0/L")
                self.assertEqual(self.bd(m, f"J{r}")["left"], ("continuous", "medium"))
            for ref in LATERAL_BLOCK:
                if ref[0] != "J":
                    self.assertIsNone(self.val(m, ref), ref)     # 直角度(R115 以外)・K:N は無い
            self.assertNotIn("直角度", [c["value"] for c in m.cells])
        else:
            for ref in LATERAL_BLOCK[:-5]:                       # J18:N23 は値も罫線もなし
                self.assertIsNone(self.val(m, ref), ref)
                b = self.bd(m, ref)
                if b:
                    # J の左は I(差)の右と共有の辺(I 自身の線は消えない)。それ以外の辺は無い
                    self.assertTrue(all(v[0] == "none" for k, v in b.items()
                                        if not (ref[0] == "J" and k == "left")), (ref, b))
            self.assertNotIn("ﾗﾃﾗﾙﾎﾞｰ", [c["value"] for c in m.cells])
        # I 列(差)の右の線は J 列と共有の辺。ラテラルありは J の左(P 列から写した medium)で閉じ、
        # なしも I 自身の右の線(N 列から写した medium)が残って閉じる。添字なしの Borders=None(J18:N23)は
        # 隣の I の辺を消さない(実機の印刷物 L-1 H8460F0 で確認)。20/22 行は Print:787-790 で I 自身の右を消す。
        for r in range(18, 24):
            closed = self.lateral or r not in (20, 22)
            self.assertEqual(self.bd(m, f"I{r}")["right"],
                             ("continuous", "medium") if closed else ("none", ""), r)
        self.common = {ref: (self.val(m, ref), {k: v for k, v in self.bd(m, ref).items()
                                                if not (ref[0] == "I" and 18 <= int(ref[1:]) <= 23 and k == "right")}
                             if self.bd(m, ref) else None) for ref in COMMON}


class _Common:
    """4 通りで切り替わらない欄が同じこと(B〜I 列から F 列を除く 11〜24 行)。"""
    baseline: dict = {}

    def run_case(self) -> None:
        m = self.printed()
        self.check(m)
        key = "ref"
        if key not in _Common.baseline:
            _Common.baseline[key] = self.common
        else:
            self.assertEqual(self.common, _Common.baseline[key])


class LateralNoFreeNo(_Common, PrintFormatBase):
    def test_format(self) -> None:
        self.run_case()


class LateralYesFreeNo(_Common, PrintFormatBase):
    lateral = True

    def test_format(self) -> None:
        self.run_case()


class LateralNoFreeYes(_Common, PrintFormatBase):
    free = True

    def test_format(self) -> None:
        self.run_case()


class LateralYesFreeYes(_Common, PrintFormatBase):
    lateral = True
    free = True

    def test_format(self) -> None:
        self.run_case()
