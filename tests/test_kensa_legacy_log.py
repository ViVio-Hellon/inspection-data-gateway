"""VBA と同じログ(WriteLog / LogOut / LogOutCombo)"""
from __future__ import annotations

import shutil
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock

from ufkensa import user_settings
from ufkensa.kensa import legacy_log, paths

NOW = datetime(2026, 10, 30, 9, 5, 3)


def use_history(tmp: Path) -> None:
    paths.set_mode(paths.OFFLINE)
    user_settings.save(paths.KEY_PATHS, {paths.OFFLINE: {"history": str(tmp) + "/"}})


class FormatTest(unittest.TestCase):
    def test_line_format_with_empty_sepl(self) -> None:
        # WriteLog:34。呼び出しのほとんどは Sepl="" → 時刻の後に空白 1、PC 名の後に空白 2、末尾に空白
        s = legacy_log.format_log_line("*UFKen|S1|24.868|L-1|阿部", "", now=NOW, computer="PC01")
        self.assertEqual(s, "2026/10/30 09:05:03 PC01  *UFKen|S1|24.868|L-1|阿部 ")

    def test_line_format_default_sepl(self) -> None:
        s = legacy_log.format_log_line("x", now=NOW, computer="PC01")
        self.assertEqual(s, "2026/10/30 09:05:03 ,PC01 , x ")

    def test_folder(self) -> None:
        # J:449-457 ほか: ﾊﾞﾗﾝｻｰ→バランサー、ﾄｯﾄ→トット、それ以外(丸徳+号機など)→丸徳
        self.assertEqual(legacy_log.log_folder("ﾊﾞﾗﾝｻｰ"), "バランサー")
        self.assertEqual(legacy_log.log_folder("ﾄｯﾄ"), "トット")
        self.assertEqual(legacy_log.log_folder("Test"), "Test")
        self.assertEqual(legacy_log.log_folder("丸徳A1"), "丸徳")
        self.assertEqual(legacy_log.log_folder(""), "丸徳")


class WriteTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="ufk_log_"))
        use_history(self.tmp)

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_new_file_has_title_cp932_crlf(self) -> None:
        self.assertTrue(legacy_log.write_line_log("L-1", "*UFKen:UFAinput起動 T1:阿部", now=NOW, computer="PC01"))
        p = self.tmp / "L-1" / "L-1_Log" / "Access.log"
        data = p.read_bytes()
        self.assertEqual(data, ("管理ﾌｧｲﾙ使用状況\r\n"
                                "2026/10/30 09:05:03 PC01  *UFKen:UFAinput起動 T1:阿部 \r\n").encode("cp932"))
        # 2 回目は追記(見出しは 1 回だけ)
        legacy_log.write_line_log("L-1", "2", now=NOW, computer="PC01")
        self.assertEqual(p.read_bytes().count("管理ﾌｧｲﾙ使用状況".encode("cp932")), 1)
        self.assertTrue(p.read_bytes().endswith(" 2 \r\n".encode("cp932")))

    def test_balancer_folder_full_width(self) -> None:
        legacy_log.write_line_log("ﾊﾞﾗﾝｻｰ", "x", now=NOW, computer="P")
        self.assertTrue((self.tmp / "バランサー" / "バランサー_Log" / "Access.log").is_file())

    def test_log_out_writes_measured_and_app_access(self) -> None:
        # 補助:1173-1178: [A] 付きは Measured.log と「Access.log などの置き場所」(VBA はアプリのフォルダ)の Access.log の 2 か所
        app = Path(tempfile.mkdtemp(prefix="ufk_app_"))
        try:
            with mock.patch.object(paths, "write_dir", return_value=str(app) + "/"):
                legacy_log.log_out("S1", "24.868", "Test", "阿部", aorb="A", aorb_caption="歪 A",
                                   now=NOW, computer="PC")
            m = (self.tmp / "Test" / "Test_Log" / "Measured.log").read_bytes().decode("cp932")
            self.assertIn(" *UFKen|[A] S1|24.868|Test|阿部 \r\n", m)
            a = (app / "Access.log").read_bytes().decode("cp932")
            self.assertIn(" *UFKen|[A] S1|24.868|Test|阿部 \r\n", a)
        finally:
            shutil.rmtree(app, ignore_errors=True)

    def test_log_out_entries(self) -> None:
        e = legacy_log.log_out_entries("S1", 24.5, "L-1", "阿部")
        self.assertEqual(e[0].msg, "*UFKen|S1|24.5|L-1|阿部")
        self.assertEqual(e[0].target_file, "Measured.log")
        self.assertEqual(e[1].target_path, "")
        e = legacy_log.log_out_entries("S1", "1", "L-1", "阿部", mitsu="三菱", aorb="A", aorb_caption="A")
        self.assertEqual(e[0].msg, "*UF全数|S1|1|L-1|阿部")
        e = legacy_log.log_out_combo_entries("CBS1", "幅", "ﾄｯﾄ", "阿部", "面B")
        self.assertEqual(e[0].msg, "*UFKen|[B] CBS1|幅|ﾄｯﾄ|阿部")
        self.assertTrue(e[0].target_path.endswith("トット\\トット_Log"))

    def test_unwritable_returns_false(self) -> None:
        # On Error Resume Next: 書けなくても例外にしない
        (self.tmp / "L-1").write_text("file, not folder")
        self.assertFalse(legacy_log.write_line_log("L-1", "x", now=NOW, computer="P"))

    def test_msg_error_entry(self) -> None:
        e = legacy_log.msg_error_entry("[許容値ｵｰﾊﾞｰ] *測定高さ:21.51", "LVC", "阿部", "", "佐藤")
        self.assertEqual(e.msg, "*UFKen:MsgError:[許容値ｵｰﾊﾞｰ] *測定高さ:21.51 T1:阿部 T2: T3:佐藤")


if __name__ == "__main__":
    unittest.main()


class WriteDirTest(unittest.TestCase):
    """共有に置いた本体のフォルダには書かない(2026-11-07)。"""

    def test_default_is_local_and_setting_wins(self) -> None:
        from ufkensa import app_config
        from ufkensa.kensa import paths as p
        self.assertEqual(Path(p.write_dir()), app_config.local_dir("logs"))
        self.assertNotEqual(Path(p.write_dir()), Path(p.app_dir()))
        p.set_path("app_log", r"\\server\share\logs" + "\\")
        try:
            self.assertEqual(p.write_dir(), r"\\server\share\logs" + "\\")
        finally:
            p.reset_path("app_log")
