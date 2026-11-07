"""ロット分割(SplitFlag)の端から端まで。

VBA: UF1.CommandButton10_Click (UF1:2937-2960) — Lot が入っている状態で Lot 登録 → 「ロット分割品ですか？」
→「間違いなくロット分割ですか？？」で SplitFlag.Caption = "Lot分割"。
SaveLoad.Time_Posting (SaveLoad:2293-2300) — SplitFlag が空でなければ WorkerTime_<L> を作り直さず前ロットから続行。
空なら A1 のロットが違うときだけ作り直す。NewKen(UF1:3445 付近の Python 版 session.py:1825)で印を消す。
"""
from __future__ import annotations

import sqlite3
import unittest
from datetime import timedelta

from tests.test_kensa_session import LOT, SessionBase
from ufkensa.kensa import records, workers
from ufkensa.kensa.store import HOZON, LOT_HISTORY

LOT2 = "H7345B1"     # 分割後の子ロット番号(合成。setUp で仕掛台帳の一時コピーに足す)


class LotSplitTest(SessionBase):
    def setUp(self) -> None:
        super().setUp()
        # 合成データ: 一時コピーの仕掛台帳に LOT の行を LOT2 として写す(分割した子ロットの想定。tests/fixtures は触らない)
        for f in ("SIKALOT.sqlite3", "SIKAHIKI.sqlite3"):
            con = sqlite3.connect(self.db_dir / f)
            cols = [r[1] for r in con.execute('PRAGMA table_info("仕掛")')]
            sel = ",".join("?" if c == "ﾛｯﾄ番号" else f'"{c}"' for c in cols)
            con.execute(f'INSERT INTO "仕掛" SELECT {sel} FROM "仕掛" WHERE "ﾛｯﾄ番号"=?', (LOT2, LOT))
            con.commit()
            con.close()

    def part1(self) -> None:
        self.register()
        self.fill("DT")
        self.fill("#L")
        self.strain_visual_height()
        self.ok(self.s.min_max())
        self.assertTrue(self.ok(self.s.print_check()).view["print_ready"])
        r = self.ok(self.s.print_commit({"print": True}))
        self.assertEqual(r.view["report"]["saved"]["mode"], "ADDNEW")

    def hozon_rows(self, lot: str) -> list:
        return [r for _, r in self.store.sheet(HOZON).rows() if records.cell_text(r[1]) == lot]

    def wt(self):
        return self.store.sheet("WorkerTime_HVC")

    def test_split_question_flow(self) -> None:
        """Lot が入っているときだけ分割の 2 段確認。いいえなら印は付かない(UF1:2950-2959)。"""
        self.register()
        st = self.s.state
        self.assertEqual(st.v("SplitFlag").value, "")
        r = self.s.register_lot(LOT, {"lot_register": True})
        self.assertEqual((r.question.key, r.question.source), ("lot_split", "UF1:2951"))
        r = self.s.register_lot(LOT, {"lot_register": True, "lot_split": True})
        self.assertEqual(r.question.key, "lot_split_sure")
        self.s.register_lot(LOT, {"lot_register": True, "lot_split": True, "lot_split_sure": False,
                                  "lot_recent": True})
        self.assertEqual(st.v("SplitFlag").value, "")
        self.s.register_lot(LOT, {"lot_register": True, "lot_split": True, "lot_split_sure": True,
                                  "lot_recent": True})
        self.assertEqual(st.v("SplitFlag").value, "Lot分割")

    def split_to(self, lot: str) -> None:
        self.ok(self.s.register_lot(lot, {"lot_register": True, "lot_split": True, "lot_split_sure": True,
                                          "lot_recent": True}))
        self.assertEqual((self.s.state.lot, self.s.state.v("SplitFlag").value), (lot, "Lot分割"))

    def test_split_mid_measure_carries_data_and_worker_time(self) -> None:
        """厚(DT)を測ったところで分割 → 新しい LotNo で残り(#L)を測って印刷。
        画面の測定値は消えず(CommandButton10 は Lot 欄などだけ空にする UF1:2971-2995)、
        WorkerTime_<L> も作り直さず続行(SaveLoad:2294-2296)、A1 だけ新ロットに(SaveLoad:2302)。"""
        self.register()
        self.fill("DT")
        wt = self.wt()
        dt_rows = [wt.get(9, c) for c in range(1, 5)]
        self.assertEqual(wt.get(1, 1), LOT)
        self.split_to(LOT2)
        st = self.s.state
        self.assertEqual(st.cells["S1"].value, "15.012")           # 前半の測定値は残る
        self.ok(self.s.select_work("厚板"))
        self.fill("#L")
        wt = self.wt()
        self.assertEqual(wt.get(1, 1), LOT2)
        self.assertEqual([wt.get(9, c) for c in range(1, 5)], dt_rows)   # 前半の時刻が残る
        self.assertNotIn(wt.get(9, 5), (None, ""))                        # 後半(幅)が足される
        self.strain_visual_height()
        self.ok(self.s.min_max())
        self.assertTrue(self.ok(self.s.print_check()).view["print_ready"])
        r = self.ok(self.s.print_commit({"print": True}))
        self.assertEqual(r.view["report"]["saved"]["mode"], "ADDNEW")
        rows = self.hozon_rows(LOT2)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][11:15], ["15.012"] * 4)
        self.assertEqual(rows[0][76], "検査員01")
        self.assertEqual(self.hozon_rows(LOT), [])
        # 開き直し: 作業保存から LOT2 と前後半の値。印はフォームにだけ(SaveLoad:2295 のコメント)
        s2 = self.new_session()
        self.ok(s2.start())
        self.assertEqual((s2.state.lot, s2.state.cells["S1"].value, s2.state.cells["S5"].value),
                         (LOT2, "15.012", "1252.3"))
        self.assertEqual(s2.state.v("SplitFlag").value, "")

    def test_split_across_shift_keeps_previous_worker(self) -> None:
        """不具合: 直をまたいで分割登録すると、前の直の検査員(T1)が登録から消える。
        VBA の CommandButton10_Click は検査員を触らない(UF1:2937-3002)。Python だけの _prune_workers
        (session.py:140-151)は「ロットが入ったまま = 持ち越し」なら残す作りだが、_register_lot が先に
        st.lot = "" にしてから呼ぶ(session.py:452-461)ので、持ち越しの分割でも必ず刈り込まれる。
        WorkerTime は分割で続行するのに、検査員の履歴(保存用の 検査員 列)からは前半の人が落ちる。"""
        self.register()
        self.fill("DT")
        shifts = workers.load_shift_times()
        first = workers.time_check(workers.hhmmss(self.clock.t), facility="HVC", shifts=shifts)
        for _ in range(48):                                         # 30 分ずつ進めて次の直へ
            self.clock.t += timedelta(minutes=30)
            if workers.time_check(workers.hhmmss(self.clock.t), facility="HVC", shifts=shifts) != first:
                break
        self.clock.t += timedelta(minutes=30)
        self.write_calibration()
        self.split_to(LOT2)
        self.assertEqual(self.s.state.v("T1").value, "検査員01")

    def test_no_split_recreates_worker_time(self) -> None:
        """対照: 分割でなければ別ロットの最初の Time_Posting で WorkerTime が作り直され、前半の記録が消える。"""
        self.register()
        self.fill("DT")
        self.ok(self.s.register_lot(LOT2, {"lot_register": True, "lot_split": False, "lot_recent": True}))
        self.assertEqual(self.s.state.v("SplitFlag").value, "")
        self.ok(self.s.select_work("厚板"))
        self.fill("#L")
        wt = self.wt()
        self.assertEqual(wt.get(1, 1), LOT2)
        self.assertEqual([wt.get(9, c) for c in range(1, 5)], [None] * 4)

    def test_split_after_print_then_new_inspection(self) -> None:
        """1 回目を印刷 → 分割で LOT2 → 新規検査は LOT2 が保存用に無いので断られ、印も残る。
        LOT2 を印刷すれば保存用は 1 ロット 1 行ずつ。"""
        self.part1()
        self.split_to(LOT2)
        r = self.s.new_inspection({"new": True})
        self.assertEqual(self.s.state.v("SplitFlag").value, "Lot分割")
        self.ok(self.s.select_work("厚板"))
        self.ok(self.s.min_max())
        self.ok(self.s.print_check())
        self.ok(self.s.print_commit({"print": True}))
        self.assertEqual((len(self.hozon_rows(LOT)), len(self.hozon_rows(LOT2))), (1, 1))
        self.ok(self.s.new_inspection({"new": True}))
        st = self.s.state
        self.assertEqual((st.lot, st.v("SplitFlag").value), ("", ""))   # Python は新規検査で印を消す(VBA は消さない)
        self.assertEqual(self.wt().get(3, 1), None)                      # NewKen が WorkerTime を作り直す(UF1:3405)


if __name__ == "__main__":
    unittest.main()
