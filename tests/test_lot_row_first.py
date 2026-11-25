"""2026-11-24(現場の声): 第1パスで取れた値は保持し、無いものだけ第2パス(圧縮版 SIKALOT)で補う。

ライン PC の設定と同じ形: 第1・第2 が同じフォルダ、オフライン、第2 は 2つ目の名前(SIKALOT_圧縮)で公差の列が無い。
前の版はこの形で第1も「2つ目の名前」になり、圧縮版だけを読んで公差が 0、キャスク・全ヵ所は列が無くて止まった。
"""
from __future__ import annotations

import shutil
import sqlite3
import tempfile
import unittest
from pathlib import Path

from tests import kensa_fixture
from ufkensa import user_settings
from ufkensa.kensa import cache_registry, lot, paths
from ufkensa.kensa.modes import cask

LOT = "N9170B0"


class LotRowFirstTest(unittest.TestCase):
    def setUp(self) -> None:
        self._saved = {k: user_settings.get(k, None) for k in (paths.KEY_PATHS, paths.KEY_MODE)}
        self.tmp = Path(tempfile.mkdtemp(prefix="ufk_lotrow_"))
        self.d = kensa_fixture.copy_to(self.tmp / "db")
        kensa_fixture.use_offline(self.d)
        paths.set_name("SIKALOTNOW", paths.SECOND, "SIKALOT_圧縮")
        comp = self.d / "SIKALOT_圧縮.sqlite3"
        shutil.copy(self.d / "SIKALOT.sqlite3", comp)
        with sqlite3.connect(comp) as c:
            for col in [r[1] for r in c.execute('PRAGMA table_info("仕掛")')]:
                if "公差" in col or col.startswith("品質ｸﾞﾚｰﾄﾞ"):
                    c.execute(f'ALTER TABLE "仕掛" DROP COLUMN "{col}"')
            c.execute('UPDATE "仕掛" SET "BOX実績_枚本数" = ? WHERE "ﾛｯﾄ番号" = ?', ("999", LOT))
        with sqlite3.connect(self.d / "SIKALOT.sqlite3") as c:
            c.execute('UPDATE "仕掛" SET "BOX実績_枚本数" = ? WHERE "ﾛｯﾄ番号" = ?', ("17", LOT))
        self.comp = comp
        cache_registry.clear_all()

    def tearDown(self) -> None:
        for k, v in self._saved.items():
            if v is None:
                user_settings.delete(k)
            else:
                user_settings.save(k, v)
        cache_registry.clear_all()
        shutil.rmtree(self.tmp, ignore_errors=True)

    def reflect(self) -> lot.LotReflect:
        return lot.acc_reflec(lot.LotScreen(lot=LOT, cb_text=LOT), check_equipment=lambda *a: "可")

    def test_first_file_uses_first_name_even_in_same_folder(self) -> None:
        self.assertTrue(lot.first_lot_file().endswith("SIKALOT.sqlite3"))

    def test_lot_registration_keeps_first(self) -> None:
        r = self.reflect()
        self.assertTrue(r.db_path.endswith("/SIKALOT.sqlite3"), r.db_path)
        self.assertNotIn("公差が 0", " ".join(m.text for m in r.messages))
        self.assertNotEqual(r.form["Ap"], "0.000")
        self.assertEqual(r.form["KenH"], "前工程実績17枚")              # 第1の値を第2で上書きしない

    def test_blank_in_first_filled_from_second(self) -> None:
        with sqlite3.connect(self.d / "SIKALOT.sqlite3") as c:
            c.execute('UPDATE "仕掛" SET "BOX実績_枚本数" = \'\' WHERE "ﾛｯﾄ番号" = ?', (LOT,))
        cache_registry.clear_all()
        self.assertEqual(self.reflect().form["KenH"], "前工程実績999枚")

    def test_only_in_second_still_reads_and_notes_why(self) -> None:
        with sqlite3.connect(self.d / "SIKALOT.sqlite3") as c:
            c.execute('DELETE FROM "仕掛" WHERE "ﾛｯﾄ番号" = ?', (LOT,))
        cache_registry.clear_all()
        r = self.reflect()
        self.assertTrue(r.db_path.endswith("SIKALOT_圧縮.sqlite3"))
        self.assertIn("第1パスの台帳", " ".join(m.text for m in r.messages))   # 公差 0 の知らせに理由を足す

    def test_cask_and_zenkasho_row_has_tolerance(self) -> None:
        row = cask.read_lot_row(LOT)
        self.assertIsNotNone(row)
        self.assertIn("板厚公差_製造_ﾌﾟﾗｽ", row)
        self.assertEqual(str(row["BOX実績_枚本数"]), "17")
        # 第1から消えても列は空で足される(「コレクションに見つかりません」で止めない)
        with sqlite3.connect(self.d / "SIKALOT.sqlite3") as c:
            c.execute('DELETE FROM "仕掛" WHERE "ﾛｯﾄ番号" = ?', (LOT,))
        cache_registry.clear_all()
        row = cask.read_lot_row(LOT)
        self.assertIn("板厚公差_製造_ﾌﾟﾗｽ", row)


if __name__ == "__main__":
    unittest.main()


class LinePcShapeTest(unittest.TestCase):
    """2026-11-25: ライン PC の形。第1 = Read\\【New】仕掛\\台帳(SIKALOT.sqlite3・全列)、
    第2 = 梱包課共有\\仕掛(SIKALOT_圧縮.sqlite3・46 列で公差の列が無い)、フォルダは別。
    「控え(第2)から読む」が入っていても、第1の公差を使う(前は第1を見ずに公差 0)。"""

    def setUp(self) -> None:
        self._saved = {k: user_settings.get(k, None) for k in (paths.KEY_PATHS, paths.KEY_MODE)}
        self.tmp = Path(tempfile.mkdtemp(prefix="ufk_linepc_"))
        base = kensa_fixture.copy_to(self.tmp / "db")
        first, second = self.tmp / "Read_台帳", self.tmp / "梱包課共有_仕掛"
        first.mkdir()
        second.mkdir()
        for n in ("SIKALOT", "SIKAHIKI", "SIKAODR"):
            shutil.copy(base / f"{n}.sqlite3", first / f"{n}.sqlite3")
            shutil.copy(base / f"{n}.sqlite3", second / f"{n}.sqlite3")
        comp = second / "SIKALOT_圧縮.sqlite3"
        (second / "SIKALOT.sqlite3").rename(comp)
        with sqlite3.connect(comp) as c:
            for col in [r[1] for r in c.execute('PRAGMA table_info("仕掛")')]:
                if "公差" in col:
                    c.execute(f'ALTER TABLE "仕掛" DROP COLUMN "{col}"')
        f = str(base) + "/"
        user_settings.save(paths.KEY_PATHS, {
            paths.KEY_ITEMS: {k: f for k in ("quality", "packing", "history", "factory")},
            paths.KEY_LOT: {paths.FIRST: str(first) + "/", paths.SECOND: str(second) + "/"},
            paths.KEY_SECOND_NAMES: {"SIKALOTNOW": "SIKALOT_圧縮.sqlite3"}})
        cache_registry.clear_all()

    def tearDown(self) -> None:
        for k, v in self._saved.items():
            if v is None:
                user_settings.delete(k)
            else:
                user_settings.save(k, v)
        cache_registry.clear_all()
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_every_mode_and_lot_source_uses_first_tolerance(self) -> None:
        online = None
        for mode in (paths.ONLINE, paths.OFFLINE):
            for forced in (False, True):
                paths.set_mode(mode)
                cache_registry.clear_all()
                with paths.lot_source(forced):
                    r = lot.acc_reflec(lot.LotScreen(lot=LOT, cb_text=LOT), check_equipment=lambda *a: "可")
                    row = cask.read_lot_row(LOT)
                why = f"{mode} 控え={forced}"
                self.assertIn("Read_台帳", r.db_path, why)
                self.assertNotIn("公差が 0", " ".join(m.text for m in r.messages), why)
                online = online or (r.form["Ap"], r.form["Hp"], r.form["Tp"])
                self.assertEqual((r.form["Ap"], r.form["Hp"], r.form["Tp"]), online, why)
                self.assertNotIn(str(row["板厚公差_製造_ﾌﾟﾗｽ"] or ""), ("", "0"), why)
        paths.set_mode(paths.ONLINE)
