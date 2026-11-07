"""測定機器のチェックと帳票の機器名(ufkensa.kensa.instruments)

`測定機器ID` は受領 DB(伝送用ファイル)の実データ。期待値は VBA(MachineTransfer)を手で追ったもの。
"""
from __future__ import annotations

import unittest

from tests import kensa_fixture
from ufkensa.kensa import flow, instruments as ins, quirks, vbx


class _Base(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        kensa_fixture.use_offline()


class CheckThickAndCodeTest(unittest.TestCase):
    def test_youto_conditions(self) -> None:
        # MT:522-545
        f = ins.check_thick_and_code
        self.assertTrue(f("0", "100", "H191,H207", 3.0, "H207"))
        self.assertFalse(f("0", "100", "H191,H207", 3.0, "H208"))
        self.assertFalse(f("0", "100", "≠H207", 3.0, "H207"))
        self.assertTrue(f("0", "100", "≠H207", 3.0, "H191"))
        self.assertTrue(f("0", "100", "≠H207,≠H191", 3.0, "H207"))     # 2 つ目の ≠ で当たる
        self.assertTrue(f("0", "100", "*20", 3.0, "H207"))
        self.assertTrue(f("0", "100", "H2*", 3.0, "H207"))
        self.assertFalse(f("0", "100", "H3*", 3.0, "H207"))
        self.assertFalse(f("0", "100", "", 3.0, "H207"))                # Split("") は要素なし
        self.assertFalse(f("0", "100", "*20", 3.0, ""))                 # InStr("", …) = 0

class DbReturnMicroTest(_Base):
    def test_l1_by_thickness(self) -> None:
        d = ins.db_return_micro
        # B52 を直した: 境界ちょうどの厚も次の区間に入る(VBA は Single の丸めで外れて "")
        cases = {3.0: ("10006", "M1"), 28.0: ("10006", "M1"), 28.01: ("10001", "M2"), 28.02: ("10001", "M2"),
                 54.0: ("10001", "M2"), 60.0: ("10003", "M3"), 74.01: ("10005", "N5"), 100.0: ("10005", "N5"),
                 300.0: ("10002", "N6"), 0.0: ("10006", "M1"), 600.0: ("", "")}
        for atu, (text, name) in cases.items():
            with self.subTest(atu=atu):
                self.assertEqual(d("L-1", atu, "H207"), text)
                self.assertEqual(d("L-1", atu, "H207", "機器名"), name)
        self.assertEqual(d("L-1", 3.0, "", "ID"), "0000158327")

    def test_line_names_d13(self) -> None:
        d = ins.db_return_micro
        self.assertEqual(d("ﾄｯﾄ", 3.0, ""), "11001")                     # マスタは全角 トット(幅だけ揃える)
        self.assertEqual(d("ﾊﾞﾗﾝｻｰ", 3.0, "", "機器名"), "M2")
        self.assertEqual(d("大板", 3.0, ""), "")                          # 機器転送の "大板" はマスタに無い
        self.assertEqual(d("丸徳A1", 3.0, ""), "")
        self.assertEqual(d("test", 3.0, ""), "")                          # 大文字小文字は揃えない(D13)
        # Test: デプス(厚下・厚上が空 = 0〜0)が M1 より前の行
        self.assertEqual(d("Test", 0.0, "", "機器名"), "デプス")
        self.assertEqual(d("Test", 3.0, "", "機器名"), "M1")

    def test_dep(self) -> None:
        d = ins.db_return_micro
        self.assertEqual(d("L-1", 3.0, "H494", "", "N5"), "10005")
        self.assertEqual(d("L-1", 3.0, "H494", "機器名", "n5"), "10005")  # Filter は大文字小文字を区別しない、n は見ない
        self.assertEqual(d("L-1", 3.0, "", "", "デプス"), "10004")
        self.assertEqual(d("LVC", 3.0, "", "", "N5"), "")


class MakerTest(unittest.TestCase):
    def test_maker_and_machin_text(self) -> None:
        self.assertEqual(ins.maker("DT10002+0000021.712M"), "DT")
        self.assertEqual(ins.maker("dt10002"), "DT")
        self.assertEqual(ins.maker("ＤＴ10002"), "DT")
        self.assertEqual(ins.maker("#L395.8"), "#L")
        self.assertEqual(ins.maker("#〆1"), "#〆")
        self.assertEqual(ins.maker("DT"), "")                             # 3 文字未満
        self.assertEqual(ins.maker(""), "")
        self.assertEqual(ins.machin_text("DT10002+0000021.712M", "前"), "10002")
        self.assertEqual(ins.machin_text("#L395.8", "前"), "前")


class MicroJugeTest(_Base):
    def mi(self, **kw) -> ins.MicroInput:
        base = dict(st="S1", machin="10006", atu="3.000", hab="1000.0", youto="K126", lname="L-1")
        base.update(kw)
        return ins.MicroInput(**base)

    def test_ok_and_ng(self) -> None:
        self.assertTrue(ins.micro_juge(self.mi()).ok)
        r = ins.micro_juge(self.mi(machin="10001"))
        self.assertFalse(r.ok)
        self.assertEqual(r.clear, "S1")
        self.assertIsInstance(r.stop, flow.Stop)
        self.assertEqual(r.stop.text, "測定するﾏｲｸﾛを正しく選定して下さい M1を使用の事")
        self.assertTrue(ins.micro_juge(self.mi(machin="")).ok)           # Machin 空は何もしない (MT:167)

    def test_h494_narrow(self) -> None:
        h = dict(youto="H494", hab="250.0")
        self.assertTrue(ins.micro_juge(self.mi(st="S5", machin="10005", **h)).ok)
        r = ins.micro_juge(self.mi(st="S5", machin="10006", **h))
        self.assertEqual(r.stop.text, "用途コード_H494の幅測定はノギスで行ってください *Not:L-1N5")
        self.assertEqual(r.stop.level, "warn")
        self.assertTrue(ins.micro_juge(self.mi(st="S1", machin="99999", **h)).ok)       # 幅以外は確かめない
        self.assertTrue(ins.micro_juge(self.mi(st="S6", machin="99999", youto="H494", hab="299.5")).ok)
        self.assertFalse(ins.micro_juge(self.mi(st="S6", machin="99999", youto="H494", hab="300.0")).ok)
        # L-1 以外は H494 でも通常の判定
        self.assertFalse(ins.micro_juge(self.mi(st="S5", machin="10005", lname="HVC", **h)).ok)
        # S10 は右 2 文字 "10"(座標 5〜6 でない)
        self.assertTrue(ins.micro_juge(self.mi(st="S10", machin="99999", **h)).ok)

    def test_type_errors(self) -> None:
        with self.assertRaises(vbx.TypeMismatch):
            ins.micro_juge(self.mi(atu=""))                               # Single への代入 (MT:170)
        with self.assertRaises(vbx.TypeMismatch):
            ins.micro_juge(self.mi(hab="幅"))
        with self.assertRaises(vbx.TypeMismatch):
            ins.micro_juge(self.mi(st="SX", youto="H494", hab="250"))     # Case 5 To 6 と "X"


class MachineTransferTest(_Base):
    def t(self, **kw) -> ins.TransferInput:
        base = dict(x=9, y=5, atu="3", haba="1000", yout="K126", line="L-1", tak="2000.0", hab="1000.0")
        base.update(kw)
        return ins.TransferInput(**base)

    def test_atu_300(self) -> None:
        cells = {(14, 2): "299.5", (15, 4): 300.0, (16, 6): "", (17, 8): "abc"}
        got = ins.machine_transfer(self.t(atu="300", cells=lambda r, c: cells.get((r, c))))
        self.assertEqual(got[-1], (9, 5, "L-1 N5,N6"))
        self.assertEqual(ins.apply_cells(got)[(9, 5)], "L-1 N5,N6")
        got = ins.machine_transfer(self.t(atu="300", cells=lambda r, c: 300 if (r, c) == (18, 8) else None))
        self.assertEqual(got[-1], (9, 5, "L-1 N6"))
        got = ins.machine_transfer(self.t(atu="300.000", cells=lambda r, c: "12" if (r, c) == (14, 2) else None))
        self.assertEqual(got[-1], (9, 5, "L-1 N5"))
        got = ins.machine_transfer(self.t(atu="300", cells=lambda r, c: None))
        self.assertEqual(got, [(9, 5, "L-1 N6"), (9, 7, "L-1 C12")])
        got = ins.machine_transfer(self.t(atu="300", cells=lambda r, c: 5 if (r, c) == (14, 3) else None))
        self.assertEqual(len(got), 2)                                     # 奇数列は見ない

    def test_h494(self) -> None:
        got = ins.machine_transfer(self.t(yout="H494", hab="250.0", haba="250"))
        self.assertEqual(got[-2:], [(8, 7, "幅:ﾉｷﾞｽ , 丈:ｺﾝﾍﾞｯｸｽ"), (9, 7, "L-1N5, C12")])
        got = ins.machine_transfer(self.t(yout="H494", hab="250.0", haba="0.5"))
        self.assertEqual(len(got), 2)
        got = ins.machine_transfer(self.t(yout="H494", hab="300.0", haba="250"))
        self.assertEqual(len(got), 2)
        with self.assertRaises(vbx.TypeMismatch):                         # UF1.Hab < 300# は And でも評価される
            ins.machine_transfer(self.t(hab=""))


if __name__ == "__main__":
    unittest.main()


# ------------------------------------------------------------------
# 直した場合の分岐(B51・B52・B53)
# ------------------------------------------------------------------
class FixBranchTest(_Base):
    def tearDown(self) -> None:
        pass
    def test_b52_bounds_same_precision(self) -> None:
        f = ins.check_thick_and_code
        # 厚下・厚上も Double(Val のまま)で比べる → 境界ちょうどが範囲に入る
        self.assertTrue(f("28.01", "54", "X", vbx.single(28.01), "X"))
        self.assertTrue(f("0", "24.99", "X", vbx.single(24.99), "X"))
        self.assertFalse(f("28.01", "54", "X", vbx.single(28.0), "X"))
        self.assertFalse(f("0", "24.99", "X", vbx.single(25.0), "X"))
        d = ins.db_return_micro
        # 再現では 28.01・74.01 は "" (DbReturnMicroTest)。直すと M2・N5
        self.assertEqual(d("L-1", 28.01, "H207", "機器名"), "M2")
        self.assertEqual(d("L-1", 74.01, "H207", "機器名"), "N5")
        self.assertEqual(d("L-1", 54.01, "H207", "機器名"), "M3")
        self.assertEqual(d("L-1", 28.0, "H207", "機器名"), "M1")
        self.assertEqual(d("HVC", 24.99, "H207", "機器名"), "M4")       # 用途の行(0–24.99)

    def test_b53_balancer_tott(self) -> None:
        f = ins.machine_transfer
        base = dict(x=9, y=5, atu="3", haba="1000", yout="K126", tak="2000.0", hab="1000.0")
        # 帳票のライン名は「大板」のまま、機器は実際のラインの行から(ﾊﾞﾗﾝｻｰ M2・トット M1)
        self.assertEqual(f(ins.TransferInput(line="ﾊﾞﾗﾝｻｰ", **base)), [(9, 5, "大板 M2"), (9, 7, "大板 C12")])
        self.assertEqual(f(ins.TransferInput(line="トット", **base)), [(9, 5, "大板 M1"), (9, 7, "大板 C11")])
        # 範囲外の厚は空のまま
        base["atu"] = "30"
        self.assertEqual(f(ins.TransferInput(line="ﾄｯﾄ", **base))[0], (9, 5, "大板 "))
        # マスタに「大板」の行があればそちらが先
        calls: list = []

        def db(line: str, *a) -> str:
            calls.append(line)
            return "X9" if line == "大板" else "M1"
        self.assertEqual(f(ins.TransferInput(line="トット", **base), db=db)[0], (9, 5, "大板 X9"))
        self.assertEqual(calls, ["大板"])
