"""使用パスの分け方(2026-11-16)と Aインプットの CSV の置き場所。"""
from __future__ import annotations

import unittest
from datetime import datetime

from tests.test_kensa_session import SessionBase
from ufkensa.kensa import ainput, paths


class SectionsTest(unittest.TestCase):
    def test_every_item_has_a_section_and_overview_is_grouped(self) -> None:
        self.assertEqual(set(paths.SECTION_OF), {i.key for i in paths.ITEMS})
        rows = paths.overview()
        secs = [r["section"] for r in rows]
        # 同じ分けは続けて並ぶ(枠が割れない)
        seen: list[str] = []
        for s in secs:
            if not seen or seen[-1] != s:
                self.assertNotIn(s, seen, secs)
                seen.append(s)
        self.assertEqual(seen, ["入力(読むもの)", "出力(検査で使う)", "出力(その他)", "道具"])
        by = {r["key"]: r for r in rows}
        self.assertIn("Aインプット", by["ainput_csv"]["frame"])
        self.assertIn("測定データ", by["history"]["frame"])


class AinputCsvRootTest(SessionBase):
    def test_default_is_vba_folder_and_setting_wins(self) -> None:
        root = ainput.ainput_csv_root("HVC")
        self.assertTrue(root.replace("/", "\\").endswith("HVC\\HVC_Ainput"), root)
        own = str(self.tmp / "ain")
        paths.set_path("ainput_csv", own)
        self.assertEqual(ainput.ainput_csv_root("HVC"), own)
        form = ainput.AinputForm(lot="LOT1234", uf1_line="HVC")
        p = ainput.save_ainput_csv(form, "LOT1234", "K1", now=datetime(2026, 11, 16, 9, 0))
        self.assertTrue(str(p).replace("\\", "/").startswith(own.replace("\\", "/")), p)
        self.assertIn("2026.11", str(p))


if __name__ == "__main__":
    unittest.main()


class MeasureDbFileTest(SessionBase):
    def test_folder_and_name(self) -> None:
        paths.set_path("measure_db", str(self.tmp / "off"))
        self.assertTrue(paths.db_file("梱包検査ﾃﾞｰﾀ").endswith("梱包検査ﾃﾞｰﾀ.sqlite3"))
        paths.set_path("measure_db_name", "測定")
        self.assertTrue(paths.db_file("梱包検査ﾃﾞｰﾀ").replace("\\", "/").endswith("off/測定.sqlite3"))
        paths.set_path("measure_db", str(self.tmp / "x" / "a.sqlite3"))      # 0.3.12 の書き方(ファイルまで)
        self.assertTrue(paths.db_file("梱包検査ﾃﾞｰﾀ").endswith("a.sqlite3"))
        c = paths.check("measure_db_name")
        self.assertTrue(c["ok"])
        self.assertIn("まだありません", c["message"])
