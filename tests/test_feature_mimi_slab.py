"""耳付き(極厚)とスラブを Session の操作どおりに通す: 登録 → 測定 → 検査済 → Min/Max → 印刷 → 保存 → 履歴から戻す。

VBA の参照先:
- 検査箇所: FormRelat.AllCheck (FormRelat:570-713) / PucketMove (716-1016) / Look_Blank (1341-1500)
    極厚: 厚 1〜4・巾 5〜10・丈 11〜14 をすべて S・U・C・D・E(#L は a=5, n=14。FormRelat:763-769, 664-682)
    スラブ: 厚 1〜4 は S・C、巾丈対角 5〜10 は S・U・C・D(FormRelat:683-711)。表示は U・D の 9・10 が無い
- 保存: SaveLoad.MainDataSave (SaveLoad:4416-4452 耳用 / 4454-4525 保存用。スラブの検査頻度は "ｽﾗﾌﾞ")
- 測定DB: 極厚は 耳付検査ﾃﾞｰﾀ、ほかは 検査ﾃﾞｰﾀ(SaveLoad:4535-4561)
"""
from __future__ import annotations

import sqlite3

from tests.test_kensa_session import SessionBase, LOT
from ufkensa.kensa import report_model
from ufkensa.kensa.store import HOZON, MIMI

DT = "DT10001+0000015.012M"
WIDTH, LENGTH, DIAG = "1252.3", "2502.4", "2800.1"


class MimiSlabBase(SessionBase):
    work = ""

    def start_work(self) -> None:
        s = self.s
        self.register()                       # 厚板を選んだあと
        self.ok(s.select_work(self.work))
        self.assertEqual(s.state.work, self.work)

    def send(self, text: str, expect_target: str = "") -> list:
        r = self.s.receive(text)
        self.assertIsNone(r.stop, r.stop and r.stop.text)
        return [m.text for m in r.messages]

    def send_seq(self, seq: list[tuple[str, str]]) -> None:
        """(欄, 伝送) を順に送り、その欄に値が入ることを確かめる(PucketMove の入力先 = VBA の順番)。"""
        st = self.s.state
        for name, text in seq:
            before = {k: c.value for k, c in st.cells.items()}
            msgs = self.send(text)
            changed = [k for k, c in st.cells.items() if c.value != before[k]]
            self.assertEqual(changed, [name], (name, msgs))

    def visual(self, faces: str) -> None:
        for p in faces:
            for n, v in ((f"CB{p}1", "幅"), (f"CB{p}2", "上"), (f"Z{p}", "0.3"), (f"CB{p}3", "幅"), (f"CB{p}4", "上"),
                         (f"G{p}", "0.2"), (f"{p}_UpVisual", "欠陥ﾅｼ"), (f"{p}_UnderVisual", "欠陥ﾅｼ")):
                if self.s.state.ctl(n).visible:
                    self.ok(self.s.set_value(n, v))
        self.ok(self.s.set_value("KenHBox1", "200"))
        self.ok(self.s.set_value("OPS1", "True"))
        self.ok(self.s.set_value("OPE5", "True"))

    def print_and_save(self) -> report_model.ReportModel:
        s = self.s
        r = self.ok(s.print_check())
        self.assertTrue(r.view["print_ready"], r.view.get("checks"))
        cap, orig = [], report_model.print_transfer

        def spy(*a, **k):
            res = orig(*a, **k)
            cap.append(res)
            return res
        report_model.print_transfer = spy
        try:
            self.saved = self.ok(s.print_commit({"print": True})).view["report"]
        finally:
            report_model.print_transfer = orig
        self.assertEqual(len(cap), 1)
        return cap[0].report

    @staticmethod
    def val(m, ref):
        c = m.cell(ref)
        return c["value"] if c else None

    def db_rows(self, table: str) -> list:
        with sqlite3.connect(self.db_dir / "梱包検査ﾃﾞｰﾀ.sqlite3") as c:
            try:
                return c.execute(f'SELECT * FROM "{table}"').fetchall()
            except sqlite3.OperationalError:
                return []

    def restore(self, mimi: bool):
        s2 = self.new_session()
        self.ok(s2.start())
        # 作業保存(Load_Text)から開き直すと、作業選択も値も戻る
        st, st2 = self.s.state, s2.state
        self.assertEqual((st2.lot, st2.work), (LOT, self.work))
        for name in st.cells:
            self.assertEqual(st2.cells[name].value, st.cells[name].value, name)
        self.ok(s2.new_inspection({"new": True}))
        self.assertEqual(s2.state.lot, "")
        self.ok(s2.history_restore(2, mimi=mimi))
        return s2


class MimiTest(MimiSlabBase):
    work = "極厚"

    def test_flow(self) -> None:
        s = self.s
        self.start_work()
        st = s.state
        self.assertEqual([st.header(p) for p in "SUCDEF"], ["1", "2", "3", "4", "5", ""])   # LSelect 極厚
        self.assertEqual(st.additional, "対角追記公差あり")
        five = "SUCDE"
        self.send_seq([(f"{p}{i}", DT) for p in five for i in range(1, 5)])
        self.assertIn("DT検査済", self.send(DT))
        self.ok(s.click(""))
        seq = [(f"{p}{i}", "#L" + (WIDTH if i <= 10 else LENGTH)) for p in five for i in range(5, 15)]
        self.send_seq(seq)
        self.assertIn("#L検査済", self.send("#L" + WIDTH))
        self.ok(s.click(""))
        for p in five:                                            # 対角差は出さない(Judg.DiagonalJudg は極厚以外)
            self.assertEqual(st.diag[p].value, "")
            for i in range(1, 15):
                self.assertEqual(st.cells[f"{p}{i}"].color, "OK", f"{p}{i}")
        self.visual(five)
        r = self.ok(s.min_max())
        self.assertEqual(r.view["minmax"], {"amin": "15.012", "amax": "15.012", "hmin": WIDTH, "hmax": WIDTH,
                                            "tmin": LENGTH, "tmax": LENGTH})
        rep = self.print_and_save()
        # 公差外なしの印刷範囲は 極厚も $B$1:$L$(Print:1175。公差外ありは P 列まで Print:1145)
        self.assertEqual((rep.print_area[1], rep.print_area[3]), (2, 12))
        # 厚 G12:J16、巾丈 C19:L23(巾1〜6 = C〜H、丈1〜4 = I〜L。Print:359-360)、対角 N19 は空
        for r, p in zip(range(12, 17), five):
            self.assertEqual([self.val(rep, c + str(r)) for c in "GHIJ"], [15.012] * 4, p)
        for r, p in zip(range(19, 24), five):
            self.assertEqual([self.val(rep, c + str(r)) for c in "CDEFGH"], [float(WIDTH)] * 6, p)
            self.assertEqual([self.val(rep, c + str(r)) for c in "IJKL"], [float(LENGTH)] * 4, p)
            self.assertIn(self.val(rep, "N" + str(r)), (None, ""), p)

        # 耳用 の 1 行(保存用には書かない)
        self.assertEqual(self.store.sheet(HOZON).last_row(1), 1)
        row = self.store.sheet(MIMI).row(2)
        self.assertEqual(row[1], LOT)
        self.assertEqual(row[2:5], ["15.000", "1252.0", "2502.0"])
        ken = row[11:91]                                          # UF_Array(1, 16, 1, 5): 位置ごとに 16 個
        self.assertEqual(len(ken), 80)
        for k, p in enumerate(five):
            blk = ken[k * 16:(k + 1) * 16]
            self.assertEqual(blk[0:4], ["15.012"] * 4, p)
            self.assertEqual(blk[4:10], [WIDTH] * 6, p)
            self.assertEqual(blk[10:14], [LENGTH] * 4, p)
        self.assertEqual(row[97], "耳付き")
        self.assertEqual(row[91], "検査員01")
        self.assertEqual(row[98:103], ["A_欠陥ﾅｼ,B_欠陥ﾅｼ"] * 5)
        # 履歴 CSV は耳用(UF1:3780-3815)
        self.assertTrue(any(p.endswith("HVC_耳用履歴CSV.csv") for p in self.saved["csv"]), self.saved["csv"])
        # 測定DB は 耳付検査ﾃﾞｰﾀ
        self.assertEqual(self.db_rows("検査ﾃﾞｰﾀ"), [])
        rows = self.db_rows("耳付検査ﾃﾞｰﾀ")
        self.assertEqual(len(rows), 1)
        self.assertEqual((rows[0][0], rows[0][2], rows[0][12]), ("HVC", LOT, "15.012"))
        self.assertEqual(rows[0][98], "耳付き")

        # 履歴(耳用)から戻す
        s2 = self.restore(mimi=True)
        st2 = s2.state
        # HistoryLoad は作業選択を戻さない(SaveLoad:4034-4037 がコメントアウト。VBA どおり)
        self.assertEqual((st2.lot, st2.work), (LOT, ""))
        for p in five:
            for i in range(1, 15):
                self.assertEqual(st2.cells[f"{p}{i}"].value, st.cells[f"{p}{i}"].value, f"{p}{i}")


class SlabTest(MimiSlabBase):
    work = "スラブ"

    def test_flow(self) -> None:
        s = self.s
        self.start_work()
        st = s.state
        self.assertEqual([st.header(p) for p in "SUCD"], ["1-上", "1-下", "2-上", "2-下"])   # SlabSet
        for p, rows in (("S", range(1, 11)), ("C", range(1, 11)), ("U", range(5, 9)), ("D", range(5, 9))):
            for i in range(1, 15):
                self.assertEqual(st.ctl(f"{p}{i}").visible, i in rows, f"{p}{i}")
        self.assertFalse(any(st.ctl(f"E{i}").visible for i in range(1, 11)))
        val = {5: WIDTH, 6: WIDTH, 7: LENGTH, 8: LENGTH, 9: DIAG, 10: DIAG}
        # 厚は S → 巾丈(#L)→ 厚 C の順で測る(厚 1〜4 の S のあと PucketMove は U5〜U8 を見る。FormRelat:831-847)
        self.send_seq([(f"S{i}", DT) for i in range(1, 5)])
        self.send_seq([(f"S{i}", "#L" + val[i]) for i in range(5, 11)]
                      + [(f"U{i}", "#L" + val[i]) for i in range(5, 9)]
                      + [(f"C{i}", "#L" + val[i]) for i in range(5, 11)]
                      + [(f"D{i}", "#L" + val[i]) for i in range(5, 9)])
        self.send_seq([(f"C{i}", DT) for i in range(1, 5)])
        self.assertIn("DT検査済", self.send(DT))
        self.ok(s.click(""))
        for p in "SC":
            self.assertEqual(st.diag[p].value, "0.0")
        self.visual("SUCD")
        r = self.ok(s.min_max())
        self.assertEqual(r.view["minmax"], {"amin": "15.012", "amax": "15.012", "hmin": WIDTH, "hmax": WIDTH,
                                            "tmin": LENGTH, "tmax": LENGTH})
        rep = self.print_and_save()
        self.assertEqual((rep.print_area[1], rep.print_area[3]), (2, 12))
        for r, p in zip(range(12, 16), "SUCD"):                  # 厚は S(1上)・C(2上)だけ
            want = [15.012] * 4 if p in "SC" else [None] * 4
            self.assertEqual([self.val(rep, c + str(r)) or None for c in "GHIJ"], want, p)
        for r, p in zip(range(19, 23), "SUCD"):
            got = [self.val(rep, c + str(r)) or None for c in "CDEFGH"]
            want = [float(WIDTH)] * 2 + [float(LENGTH)] * 2 + ([float(DIAG)] * 2 if p in "SC" else [None] * 2)
            self.assertEqual(got, want, p)
        self.assertEqual([self.val(rep, c + "23") or None for c in "CDEFGH"], [None] * 6)   # E 面なし

        self.assertEqual(self.store.sheet(MIMI).last_row(1), 1)
        row = self.store.sheet(HOZON).row(2)
        self.assertEqual(row[1], LOT)
        self.assertEqual(row[82], "ｽﾗﾌﾞ")                          # 検査頻度
        self.assertEqual(row[11:15], ["15.012"] * 4)              # S 厚1〜4
        self.assertEqual(row[11 + 24:11 + 28], ["15.012"] * 4)    # C 厚1〜4(S・U の 12 個ずつのあと)
        self.assertTrue(any(p.endswith("HVC_履歴CSV.csv") for p in self.saved["csv"]), self.saved["csv"])
        rows = self.db_rows("検査ﾃﾞｰﾀ")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][83], "ｽﾗﾌﾞ")
        self.assertEqual(self.db_rows("耳付検査ﾃﾞｰﾀ"), [])

        s2 = self.restore(mimi=False)
        st2 = s2.state
        self.assertEqual((st2.lot, st2.work), (LOT, ""))
        for p in "SUCD":
            for i in range(1, 11):
                self.assertEqual(st2.cells[f"{p}{i}"].value, st.cells[f"{p}{i}"].value, f"{p}{i}")
