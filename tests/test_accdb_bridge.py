"""Access 取り込み(TSV → sqlite3)の試験。基準データは tests/fixtures/kensa(受領した 品質基準.accdb ほか)。"""
from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

from ufkensa.accdb_bridge import importer

FIXTURES = Path(__file__).resolve().parent / "fixtures" / "kensa" / "伝送用ファイル"


class UnescapeTest(unittest.TestCase):
    def test_escapes(self):
        self.assertEqual(importer.unescape("a\\tb\\nc\\rd\\\\e"), "a\tb\nc\rd\\e")
        self.assertIsNone(importer.unescape("\\N"))
        self.assertEqual(importer.unescape("plain"), "plain")

    def test_type_map(self):
        self.assertEqual(importer.sqlite_type(3), "INTEGER")
        self.assertEqual(importer.sqlite_type(11), "INTEGER")
        self.assertEqual(importer.sqlite_type(6), "REAL")
        self.assertEqual(importer.sqlite_type(7), "TEXT")
        self.assertEqual(importer.sqlite_type(202), "TEXT")


class ImportFixturesTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.target = Path(self.tmp.name) / "伝送用ファイル.sqlite3"

    def test_builds_all_tables(self):
        steps = []
        result = importer.import_tsv_dir(FIXTURES, self.target,
                                         progress=lambda p, m: steps.append((p, m)))
        names = dict(result.tables)
        self.assertEqual(names["測定機器ID"], 22)
        self.assertEqual(names["時間用"], 4)
        self.assertEqual(names["PkgSpecData"], 871)          # 値の中の改行を行の区切りにしない
        self.assertEqual(names["キャスク梱包板番号"], 13)
        self.assertTrue(steps and steps[-1][0] >= 90)
        conn = sqlite3.connect(str(self.target))
        try:
            cols = {r[1]: r[2] for r in conn.execute("PRAGMA table_info(キャスク梱包板番号)")}
            self.assertEqual(cols["管理番号"], "INTEGER")
            self.assertEqual(cols["LotNo"], "TEXT")
            self.assertEqual(cols["作業時間"], "TEXT")       # 日付は文字で持つ
            idx = conn.execute("SELECT sql FROM sqlite_master WHERE name='ix_キャスク梱包板番号_管理番号'").fetchone()[0]
            self.assertIn("UNIQUE", idx)
            # 取り込み順(Access の返却順)のまま
            first = conn.execute("SELECT 直, 開始 FROM 時間用 ORDER BY rowid").fetchall()
            self.assertEqual(first, [("1", "07:00"), ("2", "15:00"), ("3", "22:50"), ("昼", "08:15")])
            self.assertEqual(conn.execute("PRAGMA journal_mode").fetchone()[0], "delete")
        finally:
            conn.close()

    def test_swaps_in_with_backup(self):
        importer.import_tsv_dir(FIXTURES, self.target)
        first = self.target.read_bytes()[:100]
        result = importer.import_tsv_dir(FIXTURES, self.target)
        self.assertIsNotNone(result.backup)
        self.assertTrue(result.backup.exists())
        self.assertEqual(result.backup.read_bytes()[:100], first)
        self.assertFalse(self.target.with_name(self.target.name + ".tmp").exists())

    def test_duplicate_key_gets_plain_index_and_warning(self):
        tables = [importer.Table("T", [importer.Column("管理番号", 3), importer.Column("x")],
                                 [["1", "a"], ["1", "b"], ["2", "c"]])]
        result = importer.write_sqlite(tables, self.target)
        self.assertTrue(any("重複" in w for w in result.warnings))
        conn = sqlite3.connect(str(self.target))
        idx = conn.execute("SELECT sql FROM sqlite_master WHERE name='ix_T_管理番号'").fetchone()[0]
        conn.close()
        self.assertNotIn("UNIQUE", idx)

    def test_cscript_missing_is_a_clear_error(self):
        if importer.shutil.which("cscript"):
            self.skipTest("Windows")
        with self.assertRaises(importer.BridgeError):
            importer.list_tables(FIXTURES / "nothing.accdb")


if __name__ == "__main__":
    unittest.main()
