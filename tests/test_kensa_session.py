"""検査の流れを端から端まで(session.py + wiring.py)。基準データ(tests/kensa_fixture.py)で 1 ロットを通す

ロット H7345B0(厚 15 × 巾 1252 × 丈 2502、検番 H6X1355、HVC で設備仕様「可」)を、
検査員登録 → ロット登録 → 作業選択(厚板)→ 伝送文字列で測定値を全部 → 歪・外観・検入高さ → Min/Max
→ 印刷前チェック → 印刷・保存 まで通し、保存用の 1 行・測定DB の 1 行・作業中保存からの読み戻しを確かめる。

伝送文字列は G2 §3.20 / input_router.maker・conv の形: "DT10001+0000015.012M"(DT + 機器 5 桁、値は 9 文字目から
11 文字)、"#L1252.3"(#L の後ろが値)。HVC・厚 15 の正しい機器は測定機器ID の M1 = "10001"(MT:406-501)。
時計は 1 操作ごとに 7 秒進める(TimeRecord の「4 秒以内に同じ値 = 2 度押し」を避ける。FormRelat:316)。
開始は 10:05:30(毎時 00:00〜00:45 は TimeDesignation で控えの DB を見るので避ける。補助:1027-1040)。
"""
from __future__ import annotations

import shutil
import sqlite3
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path
from unittest import mock

from tests import kensa_fixture
from ufkensa import user_settings
from ufkensa.kensa import paths, records, session, store, terminal, workers
from ufkensa.kensa.store import HOZON, LOT_HISTORY

LOT = "H7345B0"
CARD = "9000000000001"            # 名簿「検査員01」
PACK_COLS = ("用途コード", "包装仕様NO", "コメント", "備考", "受注材質", "受注調質", "厚下", "厚上", "板幅下", "板幅上",
             "板丈下", "板丈上", "フラグ", "形状", "納入先", "取引先")
WL = {"5": "1252.3", "6": "1252.5", "7": "2502.4", "8": "2502.6", "9": "2800.1", "0": "2800.6"}


class Clock:
    def __init__(self, start: datetime) -> None:
        self.t = start

    def __call__(self) -> datetime:
        self.t = self.t + timedelta(seconds=7)
        return self.t


class SessionBase(unittest.TestCase):
    calibrated = True

    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="ufkensa_session_"))
        self.db_dir = kensa_fixture.copy_to(self.tmp / "db")
        kensa_fixture.use_offline(self.db_dir)
        # 梱包資材マスタは未受領なので、注意_包装仕様 に別の仕様の 1 行だけの合成 DB を置く(test_kensa_judge と同じ)
        con = sqlite3.connect(self.db_dir / "梱包資材マスタ.sqlite3")
        con.execute("CREATE TABLE 注意_包装仕様 (" + ",".join(f'"{c}"' for c in PACK_COLS) + ")")
        con.execute("INSERT INTO 注意_包装仕様 (包装仕様NO, 形状, コメント) VALUES ('ZZ', '板', '別の仕様')")
        con.commit()
        con.close()
        paths.set_path("local_history", str(self.tmp / "検査履歴"))
        terminal.set_facility("HVC")
        p = mock.patch.object(paths, "app_dir", return_value=str(self.tmp) + "/")
        p.start()
        self.addCleanup(p.stop)
        self.clock = Clock(datetime(2026, 11, 2, 10, 5, 30))
        self.store = store.LocalStore(self.tmp / "data" / "kensa_local.sqlite3")
        if self.calibrated:
            self.write_calibration()
        self.s = self.new_session()

    def tearDown(self) -> None:
        self.store.close()
        user_settings.save(terminal.KEY_FACILITY, "")
        paths.reset_path("local_history")
        shutil.rmtree(self.tmp, ignore_errors=True)

    def new_session(self, skip: bool = False) -> session.Session:
        return session.Session(self.store, now=self.clock, calibration_skip=lambda: skip,
                               measure_db=self.db_dir / "梱包検査ﾃﾞｰﾀ.sqlite3")

    def write_calibration(self) -> None:
        """校正記録の最終行 = (直を考慮した日付, 直)(Proofreading:28-61)。HVC は 3 交替、10 時台は 1直。"""
        now = self.clock.t
        shifts = workers.load_shift_times()
        sh = self.store.ensure_sheet("校正記録")
        sh.write(2, 1, workers.yea_dayr_check(workers.hhmm(now), now=now, facility="HVC", shifts=shifts), "S")
        sh.write(2, 4, workers.time_check(workers.hhmm(now), facility="HVC", shifts=shifts), "S")

    # --- 操作 ------------------------------------------------------------
    def ok(self, r: session.Result) -> session.Result:
        self.assertIsNone(r.stop, r.stop and r.stop.text)
        self.assertIsNone(r.question, r.question and r.question.text)
        return r

    def register(self) -> None:
        self.ok(self.s.start())
        r = self.s.register_worker(CARD)
        self.assertEqual(r.question.key, "worker_register")             # WorkButton(UF1:4755)
        self.assertEqual(r.question.text, "検査員登録しますか")
        self.ok(self.s.register_worker(CARD, {"worker_register": True}))
        self.ok(self.s.register_lot(LOT, {"lot_register": True}))
        self.ok(self.s.select_work("厚板"))

    def fill(self, kind: str, value: str = "15.012") -> list[str]:
        """同じ種別の伝送を「〜検査済」が出るまで送る。入った欄の名前を返す。最後にフォームをクリック(UserForm_Click)。"""
        names = []
        for _ in range(40):
            nxt = self.s.state.inp_tag
            if kind == "DT":
                text = "DT10001+" + value.rjust(11, "0") + "M"
            else:
                text = "#L" + WL.get(nxt[-1:], "1252.1")
            r = self.s.receive(text)
            self.assertIsNone(r.stop, r.stop and r.stop.text)
            if any(m.text == kind + "検査済" for m in r.messages):
                self.ok(self.s.click(""))
                return names
            names.append(nxt)
        self.fail("検査済にならない")

    def strain_visual_height(self) -> None:
        for p in "SUCDE":
            for n, v in ((f"CB{p}1", "幅"), (f"CB{p}2", "上"), (f"Z{p}", "0.3"), (f"CB{p}3", "幅"), (f"CB{p}4", "上"),
                         (f"G{p}", "0.2"), (f"{p}_UpVisual", "欠陥ﾅｼ"), (f"{p}_UnderVisual", "欠陥ﾅｼ")):
                self.ok(self.s.set_value(n, v))
        self.ok(self.s.set_value("KenHBox1", "200"))
        self.ok(self.s.set_value("OPS1", "True"))
        self.ok(self.s.set_value("OPE5", "True"))


class EndToEndTest(SessionBase):
    def test_one_lot(self) -> None:
        s = self.s
        self.register()
        st = s.state
        # ロット登録(AccReflec。SaveLoad:180-251)
        self.assertEqual((st.lot, st.zai, st.tyu, st.youto), (LOT, "995", "R", "H165"))
        self.assertEqual((st.atu, st.apk, st.amk), ("15.000", "15.350", "14.650"))
        self.assertEqual((st.hab, st.hpk, st.hmk, st.tak, st.tpk, st.tmk),
                         ("1252.0", "1254.0", "1250.0", "2502.0", "2504.0", "2500.0"))
        self.assertEqual(st.ken_no_caption, "H6X1355")
        self.assertEqual(st.work_caption["厚板"], "H")
        self.assertEqual(st.additional, "対角追記公差あり")           # 指示_直角度 2.0(SaveLoad:279-287)
        self.assertEqual(st.v("TaikText").value, "2.0")
        v = s._view()
        self.assertTrue(v["workers"][0].startswith("T1: 検査員01("))
        self.assertTrue(v["header"]["proofread"])
        # Lot情報履歴(HistorySave)に引当の行
        hist = self.store.sheet(LOT_HISTORY)
        self.assertTrue(any(records.cell_text(r[2]) == LOT for _, r in hist.rows()))
        # 厚板 H の位置名(LSelect FormRelat:2800)
        self.assertEqual([st.header(p) for p in "SUCDE"], ["S", "S1", "2", "E1", "E"])

        self.fill("DT")
        self.fill("#L")
        for p in "SUCDE":
            for i in range(1, 5):
                self.assertEqual((st.cells[f"{p}{i}"].value, st.cells[f"{p}{i}"].color), ("15.012", "OK"))
        for p in "SCE":                                                   # 厚板: 幅・丈・対角は S・2・E だけ
            self.assertEqual([st.cells[f"{p}{i}"].value for i in range(5, 11)],
                             ["1252.3", "1252.5", "2502.4", "2502.6", "2800.1", "2800.6"])
            self.assertEqual(st.diag[p].value, "0.5")                    # Abs(2800.6 - 2800.1) を CStr(Single)
        self.assertEqual(st.cells["U5"].value, "")
        # WorkerTime に測定者と時刻(Time_Posting SaveLoad:2333-2337)
        wt = self.store.sheet("WorkerTime_HVC")
        self.assertEqual(wt.get(1, 1), LOT)
        self.assertEqual(wt.get(3, 1), "検査員01")
        self.assertEqual(wt.get(15, 1), "DT")

        self.strain_visual_height()
        r = self.ok(s.min_max())
        self.assertEqual(r.view["minmax"], {"amin": "15.012", "amax": "15.012", "hmin": "1252.3", "hmax": "1252.5",
                                            "tmin": "2502.4", "tmax": "2502.6"})
        r = self.ok(s.print_check())
        self.assertTrue(r.view["print_ready"])
        self.assertEqual([c["label"] for c in r.view["checks"]][:3], ["先行データ票", "定盤", "ゲージ"])
        self.assertTrue(all(c["ok"] in (True, None) for c in r.view["checks"]))
        r = s.print_commit()
        self.assertEqual((r.question.key, r.question.text), ("print", "印刷する？"))   # UF1:3764
        self.assertIn("印刷.wav", [m.sound for m in r.messages])
        r = self.ok(s.print_commit({"print": True}))
        self.assertIn("完了.wav", [m.sound for m in r.messages])
        rep = r.view["report"]
        self.assertEqual(rep["saved"]["mode"], "ADDNEW")

        # 保存用の 1 行(MainDataSave SaveLoad:4454-4525)
        row = self.store.sheet(HOZON).row(2)
        self.assertEqual(row[1], LOT)
        self.assertEqual(row[2:5], ["15.000", "1252.0", "2502.0"])     # Array(.ATU…) はコントロールの文字列のまま("O")
        self.assertEqual(row[11:15], ["15.012"] * 4)                    # S 厚1〜4(String の配列のまま)
        self.assertEqual(row[76], "検査員01")                            # 検査員(WorkeDayMax の末尾の人)
        self.assertEqual(row[82], "H")                                   # 検査頻度 = 厚板.Caption
        self.assertTrue(str(row[0]).startswith("2026/11/02 10:"))        # 検査開始～終了
        # 測定DB の 1 行(adoSQL ADDNEW。1 列目 = ライン名)
        with sqlite3.connect(self.db_dir / "梱包検査ﾃﾞｰﾀ.sqlite3") as c:
            rows = c.execute('SELECT * FROM "検査ﾃﾞｰﾀ"').fetchall()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][0], "HVC")
        self.assertEqual(rows[0][2], LOT)
        self.assertEqual(rows[0][12], "15.012")
        # 帳票(HTML)・履歴の PDF の置き場所(C:\検査履歴\HVC\yyyy.mm\dd日。Edge が無いので HTML だけ)・CSV
        self.assertTrue(Path(rep["report"]["html"]).is_file())
        self.assertTrue(any(f.get("html", "").endswith(str(Path("HVC", "2026.11", "02日", LOT + ".html")))
                            for f in rep["pdf"]))
        self.assertTrue(any(p.endswith("HVC_履歴CSV.csv") for p in rep["csv"]))
        # 印刷履歴 → 次の印刷は再印刷の確認(PrintOnly UF1:3579)
        r = s.print_check()
        self.assertEqual(r.question.key, "reprint")

        # 作業中保存からの読み戻し(Load_Text)。開き直した別のセッションで同じ画面になる
        s2 = self.new_session()
        self.ok(s2.start())
        st2 = s2.state
        self.assertEqual((st2.lot, st2.work, st2.atu, st2.tpk), (LOT, "厚板", "15.000", "2504.0"))
        for name in ("S1", "E4", "C5", "E10"):
            self.assertEqual(st2.cells[name].value, st.cells[name].value)
            self.assertEqual(st2.cells[name].color, "OK")                 # 読み戻し後の LoopJudg
        self.assertEqual(st2.strain_z["S"].value, "0.30")
        self.assertEqual(st2.strain_cb["CBS1"].value, "幅")
        self.assertEqual(st2.v("T1").value, "検査員01")                   # WorkerLoad(検査員_HVC)

        # 同じロットをもう一度登録 → 直近 5 Lot の確認(UF1:4505)
        r = s2.register_lot(LOT, {"lot_register": True, "lot_split": False})
        self.assertEqual(r.question.key, "lot_recent")
        self.assertIn("直近5Lot内に同じLotNoがあります", r.question.text)

        # 新規検査(NewKen_Click)。保存用にあるので確認のあと画面を空に
        r = s2.new_inspection()
        self.assertEqual((r.question.key, r.question.text), ("new", "新規検査作業を開始しますか"))
        self.ok(s2.new_inspection({"new": True}))
        self.assertEqual((st2.lot, st2.cells["S1"].value, st2.amin), ("", "", ""))


class NgTest(SessionBase):
    def test_ng_value(self) -> None:
        s = self.s
        self.register()
        st = s.state
        r = self.ok(s.receive("DT10001+0000015.500M"))                  # 上限 15.350 を超える(Judg.AHTJudg)
        self.assertEqual((st.cells["S1"].value, st.cells["S1"].color), ("15.500", "NG"))           # NumDig で "0.000"
        self.assertIn("tada.wav", [m.sound for m in r.messages])         # 音楽不合格
        for _ in range(3):
            r = self.ok(s.receive("DT10001+0000015.012M"))
        # S4 を入れた時点で Out_Tolerance → UFConfi(厚板は Rig 4 と 10。O:888-892)
        self.assertIsNotNone(r.view["confi"])
        self.assertEqual((r.view["confi"]["posi"], r.view["confi"]["n"], r.view["confi"]["moto"]), ("S", 1, "15.500"))
        self.assertIn("検査ﾃﾞｰﾀに公差外有り！", [m.text for m in r.messages])
        # 許容でない公差外があると Min/Max は止まる(FormRelat:5683)
        s.set_value("OPS1", "True")
        s.set_value("OPE1", "True")
        r = s.min_max()
        self.assertEqual(r.stop.text, "許容でない公差外が存在します")
        self.assertEqual(st.cells["S1"].value, "15.500")                 # D10: 状態は保つ
        # 「許容」と送信 → NGS = 許容(NG_Pass)→ Min/Max が通る
        self.ok(s.set_value("UFConfi.TB1", "許容"))
        self.assertEqual(st.ng["S"].value, "許容")
        self.assertIsNone(s._view()["confi"])
        r = self.ok(s.min_max())
        self.assertEqual((r.view["minmax"]["amin"], r.view["minmax"]["amax"]), ("15.012", "15.500"))

    def test_remeasure(self) -> None:
        s = self.s
        self.register()
        st = s.state
        s.receive("DT10001+0000015.500M")
        for _ in range(3):
            s.receive("DT10001+0000015.012M")
        self.ok(s.set_value("UFConfi.TB1", "再測定"))                    # NG_Back: NG の欄だけ消す(O:1053-1155)
        self.assertEqual(st.cells["S1"].value, "")
        self.assertEqual(st.cells["S2"].value, "15.012")
        self.assertEqual(st.delete_p, "S1")


class StrainFaceTest(SessionBase):
    def test_a_b_face(self) -> None:
        # AorB_Click: 表示中(A 面)の値を作業保存へ書き、B 面を表示へ。戻すと A 面が戻る(SaveLoad:1855-1979)
        s = self.s
        self.register()
        st = s.state
        for n, v in (("CBS1", "幅"), ("CBS2", "上"), ("ZS", "0.3")):
            self.ok(s.set_value(n, v))
        r = self.ok(s.set_value("AorB", ""))
        self.assertEqual(st.strain_face, "Now_B")
        self.assertIn("B面歪 入力画面です", [m.text for m in r.messages])   # UF1:487
        self.assertEqual((st.strain_z["S"].value, st.strain_z["S"].tag), ("", "0.30"))
        self.ok(s.set_value("ZS", "0.4"))
        self.ok(s.set_value("AorB", ""))
        self.assertEqual(st.strain_face, "Now_A")
        self.assertEqual((st.strain_z["S"].value, st.strain_z["S"].tag), ("0.30", "0.40"))


class CalibrationTest(SessionBase):
    calibrated = False

    def test_uncalibrated_blocks_input(self) -> None:
        s = self.s
        self.register()
        r = s.receive("DT10001+0000015.012M")
        self.assertIn("校正が済んでいません。校正を起動します。", [m.text for m in r.messages])   # Proofreading:652
        self.assertEqual(s.state.cells["S1"].value, "")                  # 値を消して抜ける(UF1:4621-4625)
        self.assertTrue(r.view["calibration_required"])
        self.assertFalse(r.view["header"]["proofread"])
        # 校正を済ませると、次の入力から通る(D11: 入力のたびに判定)
        self.write_calibration()
        self.ok(s.receive("DT10001+0000015.013M"))
        self.assertEqual(s.state.cells["S1"].value, "15.013")

    def test_calibration_skip_right(self) -> None:
        # 権限 ufkensa:calibration_skip = WorkButton.Tag="jump"(UF1:4617)
        self.s = self.new_session(skip=True)
        self.register()
        self.ok(self.s.receive("DT10001+0000015.012M"))
        self.assertEqual(self.s.state.cells["S1"].value, "15.012")


class NoLineTest(SessionBase):
    def test_no_line(self) -> None:
        user_settings.save(terminal.KEY_FACILITY, "")
        r = self.s.start()
        self.assertEqual(r.stop.text, "ライン設定がありません" + "\r\n" + "担当者に連絡して下さい")   # UF1:225
        r = self.s.register_worker(CARD, {"worker_register": True})
        self.assertIsNotNone(r.stop)

    def test_worker_needed(self) -> None:
        self.ok(self.s.start())
        self.ok(self.s.register_lot(LOT, {"lot_register": True}))
        self.ok(self.s.select_work("厚板"))
        r = self.s.receive("DT10001+0000015.012M")
        self.assertEqual(r.stop.text, "検査員登録をして下さい")            # UF1:4609
        self.assertEqual(self.s.state.cells["S1"].value, "")

    def test_unknown_card(self) -> None:
        self.ok(self.s.start())
        r = self.ok(self.s.register_worker("9999999999999", {"worker_register": True}))
        self.assertIn(workers.UNREGISTERED_SOUND, [m.sound for m in r.messages])
        self.assertEqual(r.view["workers"], [])


if __name__ == "__main__":
    unittest.main()


class HistoryRestoreTest(SessionBase):
    """履歴画面からの読み込み(UF管理 → HistoryLoad)。印刷・保存したロットを新規検査のあと戻す。"""

    def test_restore_after_new(self) -> None:
        s = self.s
        self.register()
        self.fill("DT")
        self.fill("#L")
        self.strain_visual_height()
        self.ok(s.min_max())
        self.ok(s.print_check())
        self.ok(s.print_commit({"print": True}))
        self.ok(s.new_inspection({"new": True}))
        self.assertEqual((s.state.lot, s.state.cells["S1"].value), ("", ""))
        self.assertEqual(records.kanri_lots(self.store), [LOT])
        r = self.ok(s.history_restore(2))                                  # 保存用の 2 行目を戻す
        st = s.state
        self.assertEqual(st.cells["S1"].value, "15.012")
        self.assertEqual(st.cells["C5"].value, "1252.3")

    def test_other_lot_is_refused(self) -> None:
        sh = self.store.ensure_sheet(store.HOZON)
        sh.set(2, 1, "2026/11/01 09:00～2026/11/01 09:30")
        sh.set(2, 2, "ZZZZZZ1")
        self.register()                                                   # 検査中のロットがある
        r = self.s.history_restore(2)
        self.assertIsNotNone(r.stop)
        self.assertIn("別Lotデータがあります", r.stop.text)

    def test_bad_row(self) -> None:
        self.ok(self.s.start())
        r = self.s.history_restore(99)
        self.assertIsNotNone(r.stop)


class SameToleranceAsQualityScreenTest(SessionBase):
    """決定 D15(W16): ロット登録の歪公差(DistArray)も、品質基準画面と同じくロットの検番で引く。"""

    def test_register_uses_lot_ken_no(self) -> None:
        from ufkensa.kensa import quality
        from ufkensa.kensa.quality_session import ken_no_of
        seen = []
        real = quality.dist_array

        def spy(lot_no, ken_no, *a, **k):
            seen.append((lot_no, ken_no))
            return real(lot_no, ken_no, *a, **k)

        with mock.patch.object(quality, "dist_array", spy):
            self.register()
        self.assertTrue(seen)
        lot_data = quality.LotData(*[None] * 2)
        self.assertEqual(seen[0], (LOT, ken_no_of(LOT, lot_data)))
        self.assertNotEqual(seen[0][1], "")

    def test_ken_no_changes_the_table(self) -> None:
        # 検番 K のロット: 検番なしだと設備ｺｰｽ(厚板QP)で引いていた → 検番で薄板QP
        from ufkensa.kensa import lot as lotmod, quality
        from ufkensa.kensa.quality_session import ken_no_of
        lot = "N7180C0"
        data = quality.LotData(lotmod.arrbox("SIKALOTNOW", "板", "Fields", "").value,
                               lotmod.arrbox("SIKALOTNOW", lot, "Record", "").value)
        ken = ken_no_of(lot, data)
        self.assertTrue(ken.startswith("K"))
        old = quality.dist_array(lot, "", quality.hiki_arr_create(lot), data).kousa
        new = quality.dist_array(lot, ken, quality.hiki_arr_create(lot), data).kousa
        self.assertEqual(old[0][6], "厚板QP")
        self.assertEqual(new[0][6], "薄板QP")
