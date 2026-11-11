"""ロットを読み直したとき(読み直し・別ロット登録・ロット分割)、測定済みの値を新しい公差で判定し直す(2026-11-11)。"""
from __future__ import annotations

from tests.test_lot_reload import LOT, LotReloadTest

TIGHT = ("0.01", "0.01")      # 厚 +公差 0.01 → 上限 15.010。15.012 は NG
LOOSE = ("0.35", "0.35")


class RejudgeOnReread(LotReloadTest):
    def set_plus(self, v: tuple[str, str]) -> None:
        self.exec("SIKALOTNOW", 'UPDATE "仕掛" SET "板厚公差_製造_ﾌﾟﾗｽ" = ?, "板厚公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ" = ? WHERE "ﾛｯﾄ番号" = ?',
                  (*v, LOT))

    def verdicts(self) -> set[str]:
        st = self.s.state
        return {st.verdict(n).name for n in ("S1", "U1", "C1")}

    def reregister(self, split: bool) -> None:
        r = self.s.register_lot(LOT, {"lot_register": True, "lot_split": split, "lot_split_sure": True,
                                      "lot_recent": True})
        self.assertIsNone(r.stop, r.stop and r.stop.text)
        self.assertIsNone(r.question, r.question and r.question.text)
        self.assertIn("新しい公差で判定し直しました", "\n".join(m.text for m in r.messages))

    def test_reload_ok_to_ng(self) -> None:
        self.start_lot()
        self.assertEqual(self.verdicts(), {"OK"})
        self.set_plus(TIGHT)
        self.ok(self.s.reload_lot())
        self.assertEqual(self.verdicts(), {"NG"})
        self.assertEqual(self.s.state.cells["S1"].value, "15.012")

    def test_split_ok_to_ng(self) -> None:
        self.start_lot()
        self.set_plus(TIGHT)
        self.reregister(split=True)
        self.assertEqual(self.s.state.v("SplitFlag").value, "Lot分割")
        self.assertEqual(self.verdicts(), {"NG"})

    def start_ng(self) -> None:
        self.start_lot()
        self.set_plus(TIGHT)
        self.ok(self.s.reload_lot())
        self.assertEqual(self.verdicts(), {"NG"})

    def test_other_lot_ng_to_ok(self) -> None:
        self.start_ng()
        self.set_plus(LOOSE)
        self.reregister(split=False)
        self.assertEqual(self.verdicts(), {"OK"})

    def test_split_ng_to_ok_keeps_allowed(self) -> None:
        self.start_ng()
        st = self.s.state
        st.ng["S"].value = "許容"
        st.ng["U"].value = "保留"
        self.set_plus(LOOSE)
        self.reregister(split=True)
        self.assertEqual(self.verdicts(), {"OK"})
        self.assertEqual((st.ng["S"].value, st.ng["U"].value), ("許容", "保留"))

    def test_no_change_no_rejudge_message(self) -> None:
        self.start_lot()
        r = self.s.register_lot(LOT, {"lot_register": True, "lot_split": False, "lot_recent": True})
        self.assertNotIn("判定し直しました", "\n".join(m.text for m in r.messages))
        self.assertEqual(self.verdicts(), {"OK"})
