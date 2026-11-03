"""突き合わせ(保存用の公差 vs Python で登録し直した公差)。"""
from __future__ import annotations

import shutil
import sqlite3
import tempfile
import unittest
from pathlib import Path

from tests import kensa_fixture
from ufkensa import user_settings
from ufkensa.kensa import store, terminal, verify_saved


class VerifySavedTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="ufkensa_verify_test_"))
        db = kensa_fixture.copy_to(self.tmp / "db")
        kensa_fixture.use_offline(db)
        con = sqlite3.connect(db / "梱包資材マスタ.sqlite3")
        con.execute('CREATE TABLE 注意_包装仕様 ("包装仕様NO", "形状", "コメント")')
        con.commit()
        con.close()
        terminal.set_facility("HVC")
        self.local = store.LocalStore(self.tmp / "local.sqlite3")
        sh = self.local.ensure_sheet(store.HOZON)
        good = ["15.000", "1252.0", "2502.0", "0.350", "0.350", "2.0", "2.0", "2.0", "2.0"]
        sh.write(2, 1, "2026/10/01 08:00～2026/10/01 09:00"); sh.write(2, 2, "H7345B0", "O")
        sh.write_row(2, 3, ["15.000", "1252.0", "2502.0", "0.300", "0.300", "2.0", "2.0", "2.0", "2.0"], "O")
        sh.write(3, 1, "2026/10/02 08:00～2026/10/02 09:00"); sh.write(3, 2, "H7345B0", "O")
        sh.write_row(3, 3, good, "O")
        sh.write(4, 2, "ZZZZZZZ", "O")
        sh.write_row(4, 3, good, "O")
        sh.write(5, 2, "H6083T1", "O")
        sh.write_row(5, 3, ["85.000", "1529.0", "2502.0", "0.850", "0.850", "4", "4", "2", "2"], "O")

    def tearDown(self) -> None:
        self.local.close()
        user_settings.save(terminal.KEY_FACILITY, "")
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_compare(self) -> None:
        rows = verify_saved.saved_rows(self.local)
        self.assertEqual([r.lot for r in rows], ["H7345B0", "ZZZZZZZ", "H6083T1"])   # 同じロットは新しい行
        outs = {o.row.lot: o for o in verify_saved.compare(rows, work_dir=self.tmp / "w")}
        self.assertEqual(outs["H7345B0"].status, "一致")
        self.assertEqual(outs["ZZZZZZZ"].status, "仕掛に無い")
        self.assertEqual(outs["H6083T1"].status, "一致")                                 # "4" と "4.0" は同じ
        old = verify_saved.saved_rows(self.local, latest_only=False)[0]
        o = verify_saved.compare([old], work_dir=self.tmp / "w2")[0]
        self.assertEqual(o.status, "違う")
        self.assertEqual(o.diffs, ["厚-: VBA 0.300 / Python 0.350", "厚+: VBA 0.300 / Python 0.350"])
        p = verify_saved.write_report([o], self.tmp / "r.csv")
        self.assertIn("厚-: VBA 0.300", p.read_text(encoding="cp932"))

    def test_does_not_touch_local_store(self) -> None:
        def counts():
            return {n: self.local.sheet(n).last_row(1) for n in (store.HOZON, store.LOT_HISTORY) if self.local.exists(n)}
        before = counts()
        verify_saved.compare(verify_saved.saved_rows(self.local), work_dir=self.tmp / "w")
        self.assertEqual(counts(), before)
