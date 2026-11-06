"""重さの見直し(2026-11-06): ログのまとめ書き・表の読み直しの控え・共有フォルダの確かめの間引き"""
import os
import shutil
import sqlite3
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

from tests import _isolation  # noqa: F401
from ufkensa import trace
from ufkensa.kensa import datasource, lot_copy


class TraceBatchTest(unittest.TestCase):
    def test_lines_are_batched_and_flushed(self) -> None:
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, tmp, True)
        f = tmp / "log.jsonl"
        opened = []
        real_open = open
        def counting_open(p, *a, **k):
            if str(p) == str(f):
                opened.append(p)
            return real_open(p, *a, **k)
        with mock.patch.object(trace, "SYNC", False), mock.patch.object(trace, "file_for", return_value=f), \
                mock.patch("builtins.open", counting_open):
            for i in range(30):
                trace.event("x.test", f"行 {i}")
            trace.flush()
        self.assertEqual(len(f.read_text(encoding="utf-8").splitlines()), 30)
        self.assertEqual(len(opened), 1)                       # 30 行を 1 回開いて書く

    def test_error_is_written_at_once(self) -> None:
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, tmp, True)
        f = tmp / "log.jsonl"
        with mock.patch.object(trace, "SYNC", False), mock.patch.object(trace, "file_for", return_value=f):
            trace.event("x.err", "エラー", level="ERROR")
            self.assertTrue(f.is_file())                        # 貯めずにすぐ書く


class TableCacheTest(unittest.TestCase):
    def test_same_table_read_once_until_changed(self) -> None:
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, tmp, True)
        db = tmp / "x.sqlite3"
        with sqlite3.connect(db) as c:
            c.execute("CREATE TABLE t (a TEXT)")
            c.execute("INSERT INTO t VALUES ('1')")
        c.close()
        calls = []
        real = datasource._read_uncached
        with mock.patch.object(datasource, "_read_uncached", lambda *a: calls.append(a) or real(*a)):
            for _ in range(10):
                self.assertEqual(datasource.get_records(str(db), "t"), [["1"]])
            self.assertEqual(len(calls), 1)
            time.sleep(0.01)
            with sqlite3.connect(db) as c:
                c.execute("INSERT INTO t VALUES ('2')")
            c.close()
            os.utime(db, None)
            self.assertEqual(datasource.get_records(str(db), "t"), [["1"], ["2"]])   # 変わったら読み直す
            self.assertEqual(len(calls), 2)

    def test_rows_are_copies(self) -> None:
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, tmp, True)
        db = tmp / "y.sqlite3"
        with sqlite3.connect(db) as c:
            c.execute("CREATE TABLE t (a TEXT)")
            c.execute("INSERT INTO t VALUES ('1')")
        c.close()
        r = datasource.get_records(str(db), "t")
        r[0][0] = "変えた"
        self.assertEqual(datasource.get_records(str(db), "t"), [["1"]])   # 呼び手が変えても控えは壊れない


class CopyRecheckTest(unittest.TestCase):
    def test_source_checked_at_most_every_few_seconds(self) -> None:
        self.assertEqual(lot_copy.RECHECK_SEC, 0.0)                         # 試験の既定(間引かない)
        with mock.patch.object(lot_copy, "RECHECK_SEC", 5.0):
            tmp = Path(tempfile.mkdtemp())
            self.addCleanup(shutil.rmtree, tmp, True)
            src = tmp / "SIKALOT.sqlite3"
            with sqlite3.connect(src) as c:
                c.execute("CREATE TABLE 仕掛 (a TEXT)")
            c.close()
            lot_copy.forget()
            with mock.patch.object(lot_copy, "copy_dir", return_value=tmp / "copy"):
                (tmp / "copy").mkdir()
                first = lot_copy.local_copy(str(src))
                self.assertIsNotNone(first)
                stats = []
                real = lot_copy._stamp
                with mock.patch.object(lot_copy, "_stamp", lambda p: stats.append(p) or real(p)):
                    for _ in range(20):
                        self.assertEqual(lot_copy.local_copy(str(src)), first)
                self.assertEqual(stats, [])                                  # 共有フォルダへは聞きに行かない
            lot_copy.forget()


if __name__ == "__main__":
    unittest.main()


class ProgressTest(unittest.TestCase):
    def test_label_and_estimate(self) -> None:
        import threading
        trace._progress_total.clear()
        seen = []
        def op(n: int, sample: bool) -> None:
            with trace.timed("x.op", "操作"):
                for i in range(n):
                    with trace.timed("x.step", f"段 {i}"):
                        if sample:
                            seen.append(trace.progress()[0])
        op(4, False)                                   # 1 回目で段の数を覚える
        op(4, True)
        self.assertEqual([s["label"] for s in seen], ["段 0", "段 1", "段 2", "段 3"])
        self.assertEqual([s["pct"] for s in seen], [0, 24, 48, 71])
        self.assertEqual(trace.progress(), [])         # 終わったら消える
        seen.clear()
        op(6, True)                                    # 前回より多い → 見積もりを超えたら % を出さない
        self.assertIsNone(seen[-1]["pct"])


class HistoryRowTest(unittest.TestCase):
    def test_selected_row_is_restored(self) -> None:
        """同じ LOT を 2 回保存したとき、履歴で選んだ行を戻す(前は最初に一致した行だった)。"""
        from ufkensa.kensa import records, store
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, tmp, True)
        st = store.LocalStore(tmp / "l.sqlite3")
        self.addCleanup(st.close)
        sh = st.ensure_sheet(store.HOZON)
        for r, (when, atu) in enumerate((("2026/11/06 09:00～09:30", "1.000"), ("2026/11/06 10:00～10:30", "2.000")), 2):
            sh.set(r, 1, when)
            sh.set(r, 2, "A111A11")
            sh.set(r, 3, atu)
        first = records.history_load(st, "A111A11", "", mimi=False)
        picked = records.history_load(st, "A111A11", "", mimi=False, row_no=3)
        self.assertTrue(first.found and picked.found)
        self.assertNotEqual(first.values, picked.values)
        wrong = records.history_load(st, "B222B22", "", mimi=False, row_no=3)   # LOT が合わない行番号は使わない
        self.assertFalse(wrong.found)
