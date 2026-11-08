"""「無い」は読めて確かめたときだけ。読めないときは 書かない・消さない・覚えない(2026-11-08)。"""
from __future__ import annotations

import json
import shutil
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from ufkensa import config, distribution, safe_io, settings_file, user_settings
from ufkensa.kensa import lot_copy, reasons, records


class _Tmp(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="ufk_unread_"))
        self.addCleanup(shutil.rmtree, self.tmp, True)


class SafeIoTest(_Tmp):
    def test_states(self) -> None:
        p = self.tmp / "a.json"
        self.assertEqual(safe_io.read_json_or_absent(p), (safe_io.ABSENT, None))
        p.write_text("{壊れ", encoding="utf-8")
        self.assertEqual(safe_io.read_json_or_absent(p)[0], safe_io.UNREADABLE)
        p.write_text('{"a": 1}', encoding="utf-8")
        self.assertEqual(safe_io.read_json_or_absent(p), (safe_io.OK, {"a": 1}))
        # 親フォルダに届かない → 無いとは言わない
        far = self.tmp / "unreachable" / "share" / "a.json"
        self.assertEqual(safe_io.read_json_or_absent(far)[0], safe_io.UNREADABLE)
        self.assertFalse(safe_io.confirm_absent(far))

    def test_connect_existing_does_not_create(self) -> None:
        p = self.tmp / "master.sqlite3"
        with self.assertRaises(sqlite3.OperationalError):
            safe_io.connect_existing(p)
        self.assertFalse(p.exists())


class UserSettingsTest(_Tmp):
    def setUp(self) -> None:
        super().setUp()
        self.path = self.tmp / "data" / "user_config.json"
        p = mock.patch.object(config, "USER_CONFIG_PATH", self.path)
        p.start()
        self.addCleanup(p.stop)

    def test_corrupt_json_save_does_not_wipe(self) -> None:
        self.assertTrue(user_settings.save("a", 1))           # 無い → 作る(初めての起動)
        self.assertTrue(user_settings.save("b", 2))
        good = self.path.read_text(encoding="utf-8")
        broken = good[:-3]                                     # 書きかけ・壊れ
        self.path.write_text(broken, encoding="utf-8")
        self.assertFalse(user_settings.save("c", 3))
        self.assertFalse(user_settings.delete("a"))
        self.assertEqual(self.path.read_text(encoding="utf-8"), broken)     # 上書きしていない
        self.assertIn("設定ファイルを読めないので保存しませんでした", user_settings.last_error)
        with self.assertRaises(user_settings.NotSaved):
            user_settings.save_or_raise("c", 3)
        self.path.write_text(good, encoding="utf-8")
        self.assertTrue(user_settings.save("c", 3))
        self.assertEqual(json.loads(self.path.read_text(encoding="utf-8")), {"a": 1, "b": 2, "c": 3})

    def test_distribution_does_not_overwrite_when_unreadable(self) -> None:
        self.path.parent.mkdir(parents=True)
        self.path.write_text("{壊れ", encoding="utf-8")
        key = next(iter(distribution.ITEM_KEYS))
        bundle = distribution.Bundle(settings={key: "配布の値"}, contents=None, created_at="", created_on="")
        r = distribution._apply(bundle, overwrite=False)       # noqa: SLF001
        self.assertFalse(r.ok)
        self.assertEqual(self.path.read_text(encoding="utf-8"), "{壊れ")


class SettingsFileTest(_Tmp):
    def test_unreachable_parent_no_mkdir_no_create(self) -> None:
        target = self.tmp / "share_root" / "UF" / settings_file.DEFAULT_NAME
        with self.assertRaises(OSError):
            settings_file.ensure(target)
        self.assertFalse((self.tmp / "share_root").exists())

    def test_existing_vanished_momentarily_no_empty_file(self) -> None:
        target = self.tmp / settings_file.DEFAULT_NAME
        settings_file.ensure(target)
        target.unlink()
        # このプロセスで一度そろえたファイル。exists() のあとに一瞬見えなくなった場合
        with mock.patch.object(Path, "exists", return_value=True):
            with self.assertRaises(sqlite3.Error):
                settings_file.ensure(target)
        self.assertFalse(target.exists())


class MeasureDbTest(_Tmp):
    def test_unreachable_parent_no_create(self) -> None:
        p = self.tmp / "unreachable_share" / "梱包検査ﾃﾞｰﾀ.sqlite3"
        with mock.patch.object(safe_io, "confirm_absent", return_value=False):
            with self.assertRaises(records.Stop):
                records.ensure_measure_db(p)
        self.assertFalse(p.parent.exists())

    def test_vanished_master_not_recreated_empty(self) -> None:
        p = self.tmp / "梱包検査ﾃﾞｰﾀ.sqlite3"
        with mock.patch.object(Path, "is_file", return_value=True):
            with self.assertRaises(records.Stop):
                records.register_measurement(p, "検査ﾃﾞｰﾀ", ["L1"], "")
        self.assertFalse(p.exists())


class ReasonsTest(_Tmp):
    def test_unreadable_master_shows_last_known_and_not_recreated(self) -> None:
        db = self.tmp / "伝送用ファイル.sqlite3"
        with sqlite3.connect(db) as con:
            con.execute("CREATE TABLE x (a)")
        self.assertTrue(reasons.ensure_table(db))
        with sqlite3.connect(db) as con:
            con.execute(f'INSERT INTO "{reasons.TABLE}" VALUES (?, 0, ?, 1)', (reasons.KIND_DELETE, "利用者が足した理由"))
        got = reasons.choices(reasons.KIND_DELETE, db)
        self.assertEqual(got[0], "利用者が足した理由")
        before = db.read_bytes()
        # 読めない(壊れ・ロック・共有切れ)
        with mock.patch.object(safe_io, "connect_existing", side_effect=sqlite3.OperationalError("locked")):
            again = reasons.choices(reasons.KIND_DELETE, db)
        self.assertEqual(again, got)                           # 消えない(最後に読めたもの)
        self.assertEqual(db.read_bytes(), before)              # 作り直していない

    def test_unreadable_without_history_shows_defaults_without_writing(self) -> None:
        db = self.tmp / "伝送用ファイル2.sqlite3"
        db.write_bytes(b"not a sqlite file" * 100)
        before = db.read_bytes()
        self.assertEqual(reasons.choices(reasons.KIND_REMEASURE, db), list(reasons.DEFAULTS[reasons.KIND_REMEASURE]))
        self.assertEqual(db.read_bytes(), before)
        missing = self.tmp / "gone.sqlite3"
        self.assertFalse(reasons.ensure_table(missing))
        self.assertFalse(missing.exists())


class LotCopyTest(_Tmp):
    def test_first_unreadable_keeps_local_copy(self) -> None:
        src = self.tmp / "share" / "SIKALOTNOW.sqlite3"
        src.parent.mkdir()
        with sqlite3.connect(src) as con:
            con.execute("CREATE TABLE t (a)")
            con.execute("INSERT INTO t VALUES (1)")
        cache = self.tmp / "cache"
        cache.mkdir()
        lot_copy.forget()
        self.addCleanup(lot_copy.forget)
        with mock.patch.object(lot_copy, "copy_dir", return_value=cache), \
                mock.patch("ufkensa.kensa.paths.to_local", side_effect=lambda t: Path(t)):
            dst = lot_copy.local_copy(str(src))
            self.assertIsNotNone(dst)
            sidecar = Path(str(dst) + "-journal")
            sidecar.write_bytes(b"keep")
            before = dst.read_bytes()
            src.touch()                                        # 変わったように見せる
            lot_copy._checked.clear()                          # noqa: SLF001
            with mock.patch.object(lot_copy, "_snapshot", return_value=False), \
                    mock.patch.object(lot_copy, "_raw_copy", side_effect=PermissionError("届かない")):
                self.assertIsNone(lot_copy.local_copy(str(src)))
            self.assertEqual(dst.read_bytes(), before)         # 前の写しはそのまま
            self.assertTrue(sidecar.exists())                  # 付属ファイルも消していない


if __name__ == "__main__":
    unittest.main()
