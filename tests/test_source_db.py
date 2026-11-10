"""取り込み元の sqlite3(ufkensa/source_db.py)の読み書き

* 読む口(read_table / read_query / list_tables / columns / table_counts)は読み取り専用で、元のファイルを書き換えない
* 読むときは手元へ写す。写しは中身が変わったときだけ作り直す。WAL で書き手が開いたままでも確定した行を読む
* 文字の入れ方(UTF-8 / CP932)はファイル単位で決めて読む
* 書き戻しの口(SourceConnection)は insert / update / delete / execute / query / transaction。
  where の無い update・delete は断る。失敗した文の鍵を持ったままにしない
* 開けない・無い・別物のファイルは SourceError(診断の probe は例外を投げない)
"""
from __future__ import annotations

import os
import sqlite3
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

import tests  # noqa: F401  ログなどを一時フォルダへ(ufkensa より先に)
from ufkensa import source_db as sdb


def _make(path: Path, *, wal: bool = False) -> Path:
    con = sqlite3.connect(path)
    if wal:
        con.execute("PRAGMA journal_mode=WAL")
    con.execute('CREATE TABLE "品目" ("コード" TEXT PRIMARY KEY, "名前" TEXT, "数" INTEGER)')
    con.executemany('INSERT INTO "品目" VALUES (?, ?, ?)', [("A1", "上蓋", 3), ("B2", "下敷", 5)])
    con.execute('CREATE TABLE "空" (x)')
    con.commit()
    con.close()
    return path


def _rows(path: Path, sql: str) -> list[tuple]:
    con = sqlite3.connect(path)
    try:
        return con.execute(sql).fetchall()
    finally:
        con.close()


class _Tmp(unittest.TestCase):
    def setUp(self) -> None:
        self._td = tempfile.TemporaryDirectory()
        self.addCleanup(self._td.cleanup)
        self.dir = Path(self._td.name)
        self.db = _make(self.dir / "マスタ.sqlite3")


class Helpers(unittest.TestCase):
    def test_error_kinds(self) -> None:
        self.assertTrue(sdb.is_lock_error("database is locked"))
        self.assertTrue(sdb.is_lock_error("SQLITE_BUSY"))
        self.assertFalse(sdb.is_lock_error(""))
        self.assertFalse(sdb.is_lock_error(None))           # type: ignore[arg-type]
        self.assertTrue(sdb.is_duplicate_error("UNIQUE constraint failed: t.a"))
        self.assertFalse(sdb.is_duplicate_error("no such table"))
        self.assertTrue(sdb.is_already_exists_error("duplicate column name: x"))
        self.assertTrue(sdb.is_already_exists_error("table t already exists"))

    def test_quote_identifier(self) -> None:
        self.assertEqual(sdb.quote_identifier("ﾛｯﾄ番号"), '"ﾛｯﾄ番号"')
        self.assertEqual(sdb.quote_identifier('a"b'), '"a""b"')

    def test_to_uri_keeps_unc_as_path(self) -> None:
        self.assertEqual(sdb.to_uri(Path(r"\\srv\共有\x y.sqlite3")),
                         "file:////srv/%E5%85%B1%E6%9C%89/x%20y.sqlite3")
        self.assertEqual(sdb.to_uri(Path("//srv/share/a.db")), "file:////srv/share/a.db")
        local = Path(os.path.abspath("a.sqlite3"))
        self.assertEqual(sdb.to_uri(local), local.as_uri())

    def test_decode_text(self) -> None:
        self.assertEqual(sdb.decode_text("上蓋".encode("utf-8")), "上蓋")
        self.assertEqual(sdb.decode_text("上蓋".encode("cp932")), "上蓋")    # 読めなければもう一方
        before = len(sdb.decode_failures)
        self.assertIn("\ufffd", sdb.decode_text(b"\x81\x20"))                # どちらでも読めない
        self.assertEqual(len(sdb.decode_failures), before + 1)              # 黙って潰さない
        del sdb.decode_failures[before:]


class Reading(_Tmp):
    def test_read_table_and_helpers(self) -> None:
        rows = sdb.read_table(self.db, "品目")
        self.assertEqual(rows, [{"コード": "A1", "名前": "上蓋", "数": 3}, {"コード": "B2", "名前": "下敷", "数": 5}])
        self.assertEqual(sdb.list_tables(self.db), ["品目", "空"])
        self.assertEqual(sdb.columns(self.db, "品目"), ["コード", "名前", "数"])
        self.assertEqual(sdb.table_counts(self.db), {"品目": 2, "空": 0})
        self.assertEqual(sdb.read_query(self.db, 'SELECT "数" FROM "品目" WHERE "コード" = ?', ["B2"]), [{"数": 5}])
        self.assertTrue(sdb.is_readable(self.db))

    def test_read_paths_never_write(self) -> None:
        before = self.db.read_bytes()
        with self.assertRaises(sdb.SourceError):
            sdb.read_query(self.db, 'UPDATE "品目" SET "数" = 0')
        with self.assertRaises(sdb.SourceError):
            sdb.read_query(self.db, 'DELETE FROM "品目"')
        self.assertEqual(self.db.read_bytes(), before)
        self.assertEqual(_rows(self.db, 'SELECT count(*) FROM "品目"'), [(2,)])

    def test_missing_or_bad_files(self) -> None:
        missing = self.dir / "無い.sqlite3"
        with self.assertRaises(sdb.SourceError):
            sdb.read_table(missing, "品目")
        with self.assertRaises(sdb.SourceError):
            sdb.read_table(self.db, "無い表")
        bogus = self.dir / "偽物.sqlite3"
        bogus.write_bytes(b"Standard Jet DB\x00" + b"\x00" * 4096)
        self.assertFalse(sdb.is_readable(bogus))
        self.assertFalse(sdb.is_readable(missing))
        self.assertEqual(sdb.list_tables(bogus), [])           # 一覧・列・行数は空で返す(落とさない)
        self.assertEqual(sdb.columns(missing, "品目"), [])
        self.assertEqual(sdb.table_counts(missing), {})
        with self.assertRaises(sdb.SourceError) as cm:
            sdb.read_query(bogus, "SELECT 1")
        self.assertIn("偽物.sqlite3", str(cm.exception))

    def test_copy_is_reused_until_the_source_changes(self) -> None:
        with mock.patch.object(sdb, "_snapshot", wraps=sdb._snapshot) as snap:
            sdb.read_table(self.db, "品目")
            sdb.read_table(self.db, "品目")
            sdb.list_tables(self.db)
            self.assertEqual(snap.call_count, 1, "中身が同じなら写し直さない")
            con = sqlite3.connect(self.db)
            con.execute('INSERT INTO "品目" VALUES (?, ?, ?)', ("C3", "当て木", 7))
            con.commit()
            con.close()
            st = self.db.stat()
            os.utime(self.db, ns=(st.st_atime_ns, st.st_mtime_ns + 1_000_000))   # 秒の粒が粗い FS でも変わったと分かるように
            rows = sdb.read_table(self.db, "品目")
            self.assertEqual(snap.call_count, 2, "変わったら写し直す")
        self.assertEqual([r["コード"] for r in rows], ["A1", "B2", "C3"])

    def test_wal_reads_committed_rows_while_writer_is_open(self) -> None:
        db = _make(self.dir / "wal.sqlite3", wal=True)
        writer = sqlite3.connect(db)
        self.addCleanup(writer.close)
        writer.execute('INSERT INTO "品目" VALUES (?, ?, ?)', ("W1", "書き手が開いたまま", 1))
        writer.commit()
        self.assertTrue(Path(str(db) + "-wal").exists())
        rows = sdb.read_table(db, "品目")
        self.assertIn("W1", [r["コード"] for r in rows], "-wal にしか無い確定済みの行も読む")

    def test_cp932_file_is_read_as_cp932(self) -> None:
        db = self.dir / "cp932.sqlite3"
        con = sqlite3.connect(db)
        con.execute('CREATE TABLE "t" ("名前")')
        for word in ("上蓋", "燿　", "梱包保護材", "下用", "ﾛｯﾄ番号"):
            con.execute('INSERT INTO "t" VALUES (CAST(? AS TEXT))', (word.encode("cp932"),))
        con.commit()
        con.close()
        self.assertEqual([r["名前"] for r in sdb.read_table(db, "t")], ["上蓋", "燿　", "梱包保護材", "下用", "ﾛｯﾄ番号"])
        self.assertEqual(sdb.probe(db).encoding, sdb.ENCODING_CP932)
        self.assertEqual(sdb.probe(self.db).encoding, sdb.ENCODING_UTF8)

    def test_sniff_ignores_access_internal_tables(self) -> None:
        con = sqlite3.connect(":memory:")
        con.execute('CREATE TABLE "MSysObjects" (x)')
        con.execute('INSERT INTO "MSysObjects" VALUES (CAST(? AS TEXT))', (b"\x83\x41\x83\x43\x83\x45",))
        con.execute('CREATE TABLE "t" (x)')
        con.execute('INSERT INTO "t" VALUES (?)', ("上蓋",))
        self.assertEqual(sdb.sniff_encoding(con), sdb.ENCODING_UTF8)
        con.close()


class Probing(_Tmp):
    def test_probe_ok(self) -> None:
        p = sdb.probe(self.db)
        self.assertTrue(p.ok and p.normal and p.exists and p.is_sqlite)
        self.assertEqual(p.opened_by, sdb.WAY_COPY)
        self.assertEqual(p.tables, ["品目", "空"])
        self.assertEqual(p.failures, [])
        self.assertGreater(p.size, 0)

    def test_probe_never_raises(self) -> None:
        p = sdb.probe(self.dir / "無い.sqlite3")
        self.assertFalse(p.ok)
        self.assertFalse(p.exists)
        self.assertTrue(p.error)
        bogus = self.dir / "偽物.db"
        bogus.write_bytes(b"not a database at all" * 100)
        p = sdb.probe(bogus)
        self.assertTrue(p.exists)
        self.assertFalse(p.is_sqlite)
        self.assertFalse(p.ok)
        self.assertTrue(p.failures, "折れた手を全部残す")
        self.assertIn("偽物.db", p.error)


class Writing(_Tmp):
    def test_insert_update_delete_execute_query(self) -> None:
        with sdb.connect(self.db) as con:
            self.assertEqual(con.path, self.db)
            self.assertEqual(con.table_names(), ["品目", "空"])
            self.assertEqual(con.insert("品目", {"コード": "C3", "名前": "当て木", "数": 7}), 1)
            self.assertEqual(con.update("品目", {"数": 9}, {"コード": "A1"}), 1)
            self.assertEqual(con.update("品目", {"数": 9}, {"コード": "無い"}), 0)
            self.assertEqual(con.delete("品目", {"コード": "B2"}), 1)
            self.assertEqual(con.execute('UPDATE "品目" SET "名前" = "名前" || ?', ["*"]), 2)
            self.assertEqual(con.query('SELECT "コード", "数" FROM "品目" ORDER BY 1'),
                             [{"コード": "A1", "数": 9}, {"コード": "C3", "数": 7}])
        # 1 文ごとに確定している(別の接続から見える)
        self.assertEqual(_rows(self.db, 'SELECT "コード", "名前", "数" FROM "品目" ORDER BY 1'),
                         [("A1", "上蓋*", 9), ("C3", "当て木*", 7)])

    def test_values_are_bound_not_embedded(self) -> None:
        nasty = "x'); DROP TABLE \"品目\"; --"
        with sdb.connect(self.db) as con:
            con.insert("品目", {"コード": nasty, "名前": nasty, "数": 0})
        self.assertEqual(_rows(self.db, 'SELECT "名前" FROM "品目" WHERE "コード" = \'' + nasty.replace("'", "''") + "'"),
                         [(nasty,)])

    def test_refuses_empty_values_and_where(self) -> None:
        with sdb.connect(self.db) as con:
            with self.assertRaises(sdb.SourceError):
                con.insert("品目", {})
            with self.assertRaises(sdb.SourceError):
                con.update("品目", {}, {"コード": "A1"})
            with self.assertRaises(sdb.SourceError):
                con.update("品目", {"数": 0}, {})                # 全行を書き換えない
            with self.assertRaises(sdb.SourceError):
                con.delete("品目", {})                            # 全行を消さない
        self.assertEqual(_rows(self.db, 'SELECT sum("数") FROM "品目"'), [(8,)])

    def test_errors_become_source_error(self) -> None:
        with sdb.connect(self.db) as con:
            with self.assertRaises(sdb.SourceError) as cm:
                con.insert("品目", {"コード": "A1", "名前": "重複", "数": 0})
            self.assertTrue(sdb.is_duplicate_error(str(cm.exception)))
            with self.assertRaises(sdb.SourceError):
                con.query("SELECT * FROM 無い表")
            with self.assertRaises(sdb.SourceError):
                con.execute("UPDATE 無い表 SET x = 1")
        with self.assertRaises(sdb.SourceError):
            sdb.connect(self.dir / "無い.sqlite3")

    def test_failed_statement_releases_the_write_lock(self) -> None:
        """一意制約で落ちても(二重送信の防止が効いた、ふつうの形)鍵を持ったままにしない。"""
        con = sdb.connect(self.db)
        self.addCleanup(con.close)
        with self.assertRaises(sdb.SourceError):
            con.insert("品目", {"コード": "A1", "名前": "重複", "数": 0})
        other = sqlite3.connect(self.db, timeout=0.2)
        try:
            other.execute('INSERT INTO "品目" VALUES (?, ?, ?)', ("Z9", "別の端末", 1))
            other.commit()
        finally:
            other.close()
        with con.transaction() as tx:                          # 次の取引も始められる
            tx.insert("品目", {"コード": "Z8", "名前": "続き", "数": 1})
        self.assertEqual(_rows(self.db, 'SELECT count(*) FROM "品目"'), [(4,)])

    def test_transaction_commits_together(self) -> None:
        with sdb.connect(self.db) as con:
            with con.transaction() as tx:
                rowid = tx.insert("空", {"x": "ヘッダ"})
                self.assertGreater(rowid, 0)
                self.assertEqual(tx.execute('UPDATE "品目" SET "数" = "数" + 1'), 2)
                self.assertEqual(tx.query('SELECT count(*) AS n FROM "空"'), [{"n": 1}])
                # 確定前は別の接続から見えない
                self.assertEqual(_rows(self.db, 'SELECT count(*) FROM "空"'), [(0,)])
        self.assertEqual(_rows(self.db, 'SELECT count(*) FROM "空"'), [(1,)])
        self.assertEqual(_rows(self.db, 'SELECT sum("数") FROM "品目"'), [(10,)])

    def test_transaction_rolls_back_everything(self) -> None:
        with sdb.connect(self.db) as con:
            with self.assertRaises(sdb.SourceError):
                with con.transaction() as tx:
                    tx.insert("空", {"x": "ヘッダ"})
                    tx.insert("品目", {"コード": "A1", "名前": "重複", "数": 0})     # 明細で落ちる
            with self.assertRaises(KeyError):
                with con.transaction() as tx:
                    tx.insert("空", {"x": "ヘッダ"})
                    raise KeyError("呼び手の失敗")
            with self.assertRaises(sdb.SourceError):
                with con.transaction() as tx:
                    tx.insert("空", {})
            self.assertEqual(con.query('SELECT count(*) AS n FROM "空"'), [{"n": 0}])
        self.assertEqual(_rows(self.db, 'SELECT count(*) FROM "空"'), [(0,)])

    def test_transaction_takes_the_write_lock_first(self) -> None:
        holder = sqlite3.connect(self.db, isolation_level=None)
        holder.execute("BEGIN IMMEDIATE")
        try:
            con = sdb.connect(self.db)
            self.addCleanup(con.close)
            con._conn.execute("PRAGMA busy_timeout = 50")
            with self.assertRaises(sdb.SourceError) as cm:
                with con.transaction():
                    pass
            self.assertTrue(sdb.is_lock_error(str(cm.exception)))
        finally:
            holder.execute("ROLLBACK")
            holder.close()

    def test_write_connection_is_the_real_file_not_a_copy(self) -> None:
        with mock.patch.object(sdb, "_try_copy", side_effect=AssertionError("書き戻しを写しへ向けない")):
            with sdb.connect(self.db) as con:
                con.insert("空", {"x": 1})
        self.assertEqual(_rows(self.db, 'SELECT x FROM "空"'), [(1,)])


class Finding(_Tmp):
    def test_find_and_list(self) -> None:
        other = _make(self.dir / "SIKALOT.db")
        (self.dir / "メモ.txt").write_text("x", encoding="utf-8")
        self.assertEqual(sdb.find(self.dir, "SIKALOT.accdb"), other)       # 拡張子違いも見る
        self.assertEqual(sdb.find(self.dir, "無い", "マスタ"), self.db)
        self.assertIsNone(sdb.find(self.dir, "無い"))
        self.assertEqual(sdb.list_source_files(self.dir), sorted([self.db, other]))
        self.assertEqual(sdb.list_source_files(self.dir / "無いフォルダ"), [])


class Sweeping(unittest.TestCase):
    def test_sweep_removes_only_old_foreign_copies(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            old = root / (sdb._COPY_PREFIX + "old")
            new = root / (sdb._COPY_PREFIX + "new")
            mine = root / (sdb._COPY_PREFIX + "mine")
            other = root / "someone_else"
            for d in (old, new, mine, other):
                d.mkdir()
                (d / "x.sqlite3").write_bytes(b"x")
            ago = time.time() - 7 * 3600
            for d in (old, mine, other):
                os.utime(d, (ago, ago))
            with mock.patch.object(sdb.tempfile, "gettempdir", return_value=td), \
                    mock.patch.object(sdb, "_COPY_DIR", mine):
                self.assertEqual(sdb.sweep_old_copies(keep_hours=6), 1)
            self.assertFalse(old.exists())
            self.assertTrue(new.exists(), "まだ新しいもの(隣で動いている起動の分)は残す")
            self.assertTrue(mine.exists(), "自分の写しは残す")
            self.assertTrue(other.exists(), "名前の違うものには触らない")


if __name__ == "__main__":
    unittest.main()
