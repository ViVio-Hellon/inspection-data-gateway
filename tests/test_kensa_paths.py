"""使用パスの 第1パス / 第2パス(判断事項 D26。2026-11-02 の決め)

- オンライン: 第1を読み、第1に無いものだけ第2。オフライン: 第2だけ。書く項目は オンライン第1・オフライン第2
- 仕掛台帳: 第1 = 台帳(MainPath)、第2 = 控え(NextPath)。ファイルごと・ロットごとに第2。2つ目の名前
- 前の保存(オンライン/オフライン)の読み替え
- 設定画面の API: 変えるには管理者パスワード(auth.master_edit_ok)。断りは 403
"""
from __future__ import annotations

import datetime as dt
import shutil
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from tests import kensa_fixture
from ufkensa import auth, user_settings
from ufkensa.kensa import datasource, lot, paths

T_NORMAL = dt.datetime(2026, 10, 30, 10, 30, 0)     # 時刻窓の外


def _ok(*_a) -> str:
    return "可"


class _Saved(unittest.TestCase):
    """使用パスとモードを試験の前後で戻す。"""

    def setUp(self) -> None:
        self._paths = user_settings.get(paths.KEY_PATHS, None)
        self._mode = user_settings.get(paths.KEY_MODE, None)
        self.root = Path(tempfile.mkdtemp(prefix="ufk_paths_"))
        self.addCleanup(shutil.rmtree, self.root, True)
        self.d1, self.d2 = self.root / "first", self.root / "second"
        self.d1.mkdir()
        self.d2.mkdir()

    def tearDown(self) -> None:
        for k, v in ((paths.KEY_PATHS, self._paths), (paths.KEY_MODE, self._mode)):
            if v is None:
                user_settings.delete(k)
            else:
                user_settings.save(k, v)

    def use(self, mode: str, **extra) -> None:
        paths.set_mode(mode)
        f = {k: str(self.d1) + "/" for k in (paths.LOT, "quality", "packing", "history", "factory")}
        s = {k: str(self.d2) + "/" for k in (paths.LOT, "quality", "packing", "history", "factory")}
        user_settings.save(paths.KEY_PATHS, {paths.FIRST: f, paths.SECOND: s, **extra})

    def put(self, where: Path, name: str, as_name: str = "") -> None:
        shutil.copy2(kensa_fixture.build() / kensa_fixture.file_name(name), where / (as_name or kensa_fixture.file_name(name)))


# ------------------------------------------------------------------
class ReadWriteMatrixTest(_Saved):
    """読む項目(品質基準)と書く項目(検査履歴)の モード × ある/無い。"""

    def test_read_online(self) -> None:
        self.use(paths.ONLINE)
        p1, p2 = str(self.d1) + "/" + "品質基準.sqlite3", str(self.d2) + "/" + "品質基準.sqlite3"
        self.assertEqual(paths.db_file("品質基準"), p1)              # どちらにも無い → 第1(呼び手が「ありません」)
        self.put(self.d2, "品質基準")
        self.assertEqual(paths.db_file("品質基準"), p2)              # 第1に無い → 第2
        self.put(self.d1, "品質基準")
        self.assertEqual(paths.db_file("品質基準"), p1)              # 両方ある → 第1
        self.assertIsNotNone(datasource.get_fields("品質基準", "QA_022_01", ""))

    def test_read_offline(self) -> None:
        self.use(paths.OFFLINE)
        self.put(self.d1, "品質基準")
        p2 = str(self.d2) + "/" + "品質基準.sqlite3"
        self.assertEqual(paths.db_file("品質基準"), p2)              # 第1にあっても見ない
        with self.assertRaises(datasource.DbMissing):                # 第2に無ければ「見つかりません」
            datasource.get_fields("品質基準", "QA_022_01", "")
        self.put(self.d2, "品質基準")
        self.assertEqual(paths.db_file("品質基準"), p2)

    def test_folder_level_read(self) -> None:
        self.use(paths.ONLINE)
        paths.set_path("wav", str(self.root / "無い") + "/", paths.FIRST)
        paths.set_path("wav", str(self.d2) + "/", paths.SECOND)
        self.assertEqual(paths.get("wav"), str(self.d2) + "/")        # 第1のフォルダが無い → 第2
        self.assertEqual(paths.which("wav"), paths.SECOND)
        paths.set_path("wav", str(self.d1) + "/", paths.FIRST)
        self.assertEqual(paths.get("wav"), str(self.d1) + "/")
        paths.set_mode(paths.OFFLINE)
        self.assertEqual(paths.get("wav"), str(self.d2) + "/")

    def test_write_items_never_fall_back(self) -> None:
        self.use(paths.ONLINE)
        paths.set_path("history", str(self.root / "届かない") + "/", paths.FIRST)
        self.assertEqual(paths.get("history"), str(self.root / "届かない") + "/")   # 第1が無くても第1(書けないと知らせる)
        self.assertEqual(paths.db_file("梱包検査ﾃﾞｰﾀ"), str(self.d1) + "/" + "【測定DB】梱包\\梱包検査ﾃﾞｰﾀ.sqlite3")
        paths.set_mode(paths.OFFLINE)
        self.assertEqual(paths.get("history"), str(self.d2) + "/")
        self.assertEqual(paths.get("nis"), "")                        # 第2の既定は空(VBA はオフライン時に空)
        self.put(self.d2, "品質基準")
        paths.set_mode(paths.ONLINE)
        self.assertEqual(paths.db_file("品質基準", write=True), str(self.d1) + "/" + "品質基準.sqlite3")


class LotFileTest(_Saved):
    """仕掛台帳: ファイルごとの第2・作り直し中・2つ目の名前・SIKACOMP は第2。"""

    def test_aliases(self) -> None:
        self.use(paths.ONLINE)
        self.assertEqual((paths.get("main"), paths.get("next")), (str(self.d1) + "/", str(self.d2) + "/"))
        paths.set_mode(paths.OFFLINE)
        self.assertEqual((paths.get("main"), paths.get("next")), (str(self.d2) + "/", str(self.d2) + "/"))

    def test_per_file(self) -> None:
        self.use(paths.ONLINE, second_names={"SIKALOTNOW": "SIKALOT_圧縮"})
        self.put(self.d2, "SIKALOTNOW", "SIKALOT_圧縮.sqlite3")
        self.put(self.d2, "SIKAHIKINOW")
        self.put(self.d1, "SIKAHIKINOW")
        self.assertEqual(paths.db_file("SIKALOTNOW"), str(self.d2) + "/SIKALOT_圧縮.sqlite3")    # 第1に無い → 第2の名前
        self.assertEqual(paths.db_file("SIKAHIKINOW"), str(self.d1) + "/SIKAHIKI.sqlite3")       # 第1にある
        (self.d1 / paths.PLEASE_WAIT).write_text("x")
        self.assertEqual(paths.db_file("SIKAHIKINOW"), str(self.d2) + "/SIKAHIKI.sqlite3")       # 作り直し中 → 第2
        self.assertEqual(paths.db_file("SIKACOMP"), str(self.d2) + "/SIKACOMP.sqlite3")          # VBA は NextPath
        paths.set_mode(paths.OFFLINE)
        self.assertEqual(paths.db_file("SIKALOTNOW"), str(self.d2) + "/SIKALOT_圧縮.sqlite3")
        self.assertEqual(lot.db_path(paths.get("main"), "SIKALOTNOW"), str(self.d2) + "/SIKALOT_圧縮.sqlite3")

    def test_second_name_rules(self) -> None:
        self.use(paths.ONLINE)
        self.assertEqual(paths.second_name("SIKAHIKINOW"), "SIKAHIKI.sqlite3")
        paths.set_second_name("SIKAHIKINOW", "HIKI2.db")
        self.assertEqual(paths.second_name("SIKAHIKINOW"), "HIKI2.db")                          # 拡張子まで入れればそのまま
        self.assertEqual(lot.db_path(paths.get("next"), "SIKAHIKINOW"), str(self.d2) + "/HIKI2.db")
        self.assertEqual(lot.db_path(paths.get("main"), "SIKAHIKINOW"), str(self.d1) + "/SIKAHIKI.sqlite3")
        with self.assertRaises(ValueError):
            paths.set_second_name("SIKAHIKINOW", r"C:\x\HIKI.sqlite3")                          # フォルダは入れない
        paths.set_second_name("SIKAHIKINOW", "")
        self.assertEqual(paths.second_name("SIKAHIKINOW"), "SIKAHIKI.sqlite3")


class PerLotTest(_Saved):
    """ロットごとの第2: 第1の SIKAHIKI / SIKALOT / SIKAODR にロットが無ければ第2を見る。"""

    def setUp(self) -> None:
        super().setUp()
        for name in ("SIKALOTNOW", "SIKAHIKINOW", "SIKAODRNOW"):
            self.put(self.d1, name)
            self.put(self.d2, name)

    def delete(self, where: Path, name: str, col: str, values: list[str]) -> None:
        with sqlite3.connect(where / kensa_fixture.file_name(name)) as c:
            c.executemany(f'DELETE FROM "仕掛" WHERE "{col}" = ?', [(v,) for v in values])
        c.close()

    def reverse_columns(self, path: Path) -> None:
        with sqlite3.connect(path) as c:
            cols = [r[1] for r in c.execute('PRAGMA table_info("仕掛")')]
            sel = ", ".join(f'"{x}"' for x in reversed(cols))
            c.execute(f'CREATE TABLE t2 AS SELECT {sel} FROM "仕掛" ORDER BY rowid')
            c.execute('DROP TABLE "仕掛"')
            c.execute('ALTER TABLE t2 RENAME TO "仕掛"')
        c.close()

    def test_arrbox_lot_from_second(self) -> None:
        self.use(paths.ONLINE)
        base = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL)
        self.assertEqual(base.file_path, str(self.d1) + "/SIKAHIKI.sqlite3")
        self.delete(self.d1, "SIKAHIKINOW", "ﾛｯﾄ番号", ["H2500K0"])
        r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL)
        self.assertEqual(r.file_path, str(self.d2) + "/SIKAHIKI.sqlite3")
        self.assertEqual(r.value, base.value)
        self.assertEqual(r.messages, [])
        # 第2にも無い → VBA と同じく Empty(メッセージなし)
        self.delete(self.d2, "SIKAHIKINOW", "ﾛｯﾄ番号", ["H2500K0"])
        self.assertIsNone(lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL).value)

    def test_second_with_other_name_and_column_order(self) -> None:
        self.use(paths.ONLINE, second_names={"SIKAHIKINOW": "SIKAHIKI_圧縮.sqlite3"})
        base = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL).value
        (self.d2 / "SIKAHIKI.sqlite3").rename(self.d2 / "SIKAHIKI_圧縮.sqlite3")
        self.reverse_columns(self.d2 / "SIKAHIKI_圧縮.sqlite3")
        self.delete(self.d1, "SIKAHIKINOW", "ﾛｯﾄ番号", ["H2500K0"])
        r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL)
        self.assertEqual(r.file_path, str(self.d2) + "/SIKAHIKI_圧縮.sqlite3")
        self.assertEqual(r.value, base)                               # 第1の列の並びに直してある

    def test_offline_reads_only_second(self) -> None:
        self.use(paths.OFFLINE)
        self.delete(self.d2, "SIKAHIKINOW", "ﾛｯﾄ番号", ["H2500K0"])
        r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL)
        self.assertEqual(r.file_path, str(self.d2) + "/SIKAHIKI.sqlite3")
        self.assertIsNone(r.value)                                    # 第1にあっても見ない
        (self.d2 / "SIKAHIKI.sqlite3").unlink()
        r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL)
        self.assertEqual([m.text for m in r.messages], ["日稲:SIKAHIKINOW.accdbがありません_ReadCheck_Arrbox"])

    def test_acc_reflec_lot_from_second(self) -> None:
        self.use(paths.ONLINE)
        screen = lot.LotScreen(lot="H2500K0", cb_text="H2500K0", work_time="T")
        base = lot.acc_reflec(screen, check_equipment=_ok)
        self.delete(self.d1, "SIKALOTNOW", "ﾛｯﾄ番号", ["H2500K0"])
        r = lot.acc_reflec(screen, check_equipment=_ok)
        self.assertEqual((r.status, r.db_path), ("ok", str(self.d2) + "/SIKALOT.sqlite3"))
        self.assertEqual(r.past_hiki, base.past_hiki)
        self.delete(self.d2, "SIKALOTNOW", "ﾛｯﾄ番号", ["H2500K0"])
        self.assertEqual(lot.acc_reflec(screen, check_equipment=_ok).status, "not_found")

    def test_history_save_odr_from_second(self) -> None:
        self.use(paths.ONLINE)
        ph = lot.acc_reflec(lot.LotScreen(lot="H2500K0", cb_text="H2500K0", work_time="T"), check_equipment=_ok).past_hiki
        sheet = [list(lot.HISTORY_HEADER)]
        base = lot.history_save(ph, sheet, lot.HistoryEnv(now=T_NORMAL))
        self.delete(self.d1, "SIKAODRNOW", "受注番号", ["06017610", "06017503", "06019756"])
        r = lot.history_save(ph, sheet, lot.HistoryEnv(now=T_NORMAL))
        self.assertTrue(r.completed)
        self.assertEqual(r.new_rows, base.new_rows)
        # SIKAHIKI にロットが無い → VBA どおり控えへ(SaveLoad:1343-1352)
        self.delete(self.d1, "SIKAHIKINOW", "ﾛｯﾄ番号", ["H2500K0"])
        r = lot.history_save(ph, sheet, lot.HistoryEnv(now=T_NORMAL))
        self.assertEqual(r.file_path, str(self.d2) + "/SIKAHIKI.sqlite3")
        self.assertEqual(r.new_rows, base.new_rows)

    def test_sika_atu_create_from_second(self) -> None:
        self.use(paths.ONLINE)
        base = lot.sika_atu_create("H2500K0")
        self.delete(self.d1, "SIKALOTNOW", "ﾛｯﾄ番号", ["H2500K0"])
        self.reverse_columns(self.d2 / "SIKALOT.sqlite3")
        self.assertEqual(lot.sika_atu_create("H2500K0"), base)


class MigrationTest(_Saved):
    """前の保存(オンライン/オフライン)→ 第1 / 第2。"""

    def test_old_online_offline(self) -> None:
        user_settings.save(paths.KEY_PATHS, {
            paths.ONLINE: {"main": r"\\a\台帳" "\\", "next": r"\\a\控え" "\\", "wav": r"D:\音1"},
            paths.OFFLINE: {"main": r"C:\app" "\\", "next": r"C:\app" "\\", "wav": r"D:\音2", "nis": ""}})
        self.assertEqual(paths.slot_value(paths.LOT, paths.FIRST), r"\\a\台帳" "\\")
        self.assertEqual(paths.slot_value(paths.LOT, paths.SECOND), r"\\a\控え" "\\")    # VBA オンラインの NextPath
        self.assertEqual((paths.slot_value("wav", paths.FIRST), paths.slot_value("wav", paths.SECOND)),
                         (r"D:\音1", r"D:\音2"))
        self.assertEqual(paths.slot_value("nis", paths.SECOND), "")
        stored = user_settings.get(paths.KEY_PATHS)
        self.assertNotIn(paths.ONLINE, stored)                                          # 新しい形で保存し直す
        self.assertEqual(stored[paths.FIRST]["wav"], r"D:\音1")

    def test_old_offline_only(self) -> None:
        user_settings.save(paths.KEY_PATHS, {paths.OFFLINE: {"main": "/m/", "next": "/n/", "history": "/h/"}})
        paths.set_mode(paths.OFFLINE)
        self.assertEqual((paths.get("main"), paths.get("next"), paths.get("history")), ("/n/", "/n/", "/h/"))
        self.assertEqual(paths.slot_value(paths.LOT, paths.FIRST), paths.slot_default(paths.LOT, paths.FIRST))

    def test_old_call_style(self) -> None:
        paths.set_path("wav", r"D:\x", paths.ONLINE)                 # 前の呼び方(モード)は 第1 / 第2 に読み替え
        self.assertEqual(paths.slot_value("wav", paths.FIRST), r"D:\x")
        paths.set_path("next", "/n/")
        self.assertEqual(paths.slot_value(paths.LOT, paths.SECOND), "/n/")
        paths.reset_path("next")
        self.assertEqual(paths.slot_value(paths.LOT, paths.SECOND), paths.slot_default(paths.LOT, paths.SECOND))


class SettingsApiTest(_Saved):
    """設定画面の使用パス: 変えるには管理者パスワード(403 locked / not_admin)。確認は両方。"""

    def setUp(self) -> None:
        super().setUp()
        from tests._web import make_client
        self.c = make_client()
        self.h = {"X-Tool-Token": self.c.application.config["TOKEN"]}
        self.use(paths.ONLINE)

    def post(self, url, **body):
        return self.c.post(url, json=body, headers=self.h)

    def test_locked_and_not_admin(self) -> None:
        with mock.patch.object(auth, "can_edit", return_value=True), \
                mock.patch.object(auth, "master_edit_ok", return_value=False):
            r = self.post("/api/settings/path", key="wav", slot="first", path="x")
            self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "locked"))
            r = self.post("/api/settings/path", key="lot", db="SIKALOTNOW", second_name="x")
            self.assertEqual(r.status_code, 403)
        with mock.patch.object(auth, "can_edit", return_value=False):
            r = self.post("/api/settings/path", key="wav", slot="second", path="x")
            self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "not_admin"))
            self.assertEqual(r.get_json()["error"]["message"], "編集には、アクセス権限の「管理者」が要ります")
        self.assertNotEqual(paths.slot_value("wav", paths.FIRST), "x")

    def test_edit_when_unlocked(self) -> None:
        with mock.patch.object(auth, "master_edit_ok", return_value=True):
            st = self.post("/api/settings/path", key="wav", slot="second", path=str(self.d2) + "/").get_json()
            row = [p for p in st["paths"] if p["key"] == "wav"][0]
            self.assertEqual((row["second"], row["second_changed"]), (str(self.d2) + "/", True))
            st = self.post("/api/settings/path", key="lot", db="SIKAHIKINOW", second_name="HIKI2").get_json()
            lot_row = [p for p in st["paths"] if p["key"] == "lot"][0]
            self.assertEqual([n["effective"] for n in lot_row["second_names"]],
                             ["SIKALOT.sqlite3", "HIKI2.sqlite3", "SIKAODR.sqlite3"])
            self.assertEqual(lot_row["in_use"][paths.OFFLINE], "第2")
            self.assertEqual(self.post("/api/settings/path", key="wav", path="x").status_code, 400)   # slot が要る

    def test_check_both(self) -> None:
        self.put(self.d2, "SIKALOTNOW")
        d = self.post("/api/settings/path/check", key="lot").get_json()
        self.assertTrue(d["first"]["ok"] and d["second"]["ok"])
        self.assertEqual([f["exists"] for f in d["second"]["files"]][:3], [True, False, False])
        self.assertIn("第1パス", d["message"])
