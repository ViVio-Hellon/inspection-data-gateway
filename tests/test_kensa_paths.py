"""使用パス(判断事項 D26。2026-11-02 決め、11-03 改)

- 仕掛台帳だけ 第1 / 第2。オンラインは第1 → 無いファイル・ロットだけ第2、オフラインは第2だけ
- 仕掛 DB は手元へ写してから読む。写せないファイルは読まない(第2へ / 「ありません」)
- 上流の作り直しの決まり(PLEASE WAIT・時刻窓)は使わない
- ほかの項目はパス 1 つ(モードによらない)。前の保存の読み替え
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
from ufkensa.kensa import datasource, lot, lot_copy, paths

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
        items = {k: str(self.d1) + "/" for k in ("quality", "packing", "history", "factory")}
        user_settings.save(paths.KEY_PATHS, {paths.KEY_ITEMS: items,
                                             paths.KEY_LOT: {paths.FIRST: str(self.d1) + "/",
                                                             paths.SECOND: str(self.d2) + "/"}, **extra})

    def put(self, where: Path, name: str, as_name: str = "") -> None:
        shutil.copy2(kensa_fixture.build() / kensa_fixture.file_name(name), where / (as_name or kensa_fixture.file_name(name)))


# ------------------------------------------------------------------
class SinglePathTest(_Saved):
    """仕掛台帳でない項目はパス 1 つ。モードで変わらない・第2へも逃げない。"""

    def test_mode_does_not_matter(self) -> None:
        for mode in (paths.ONLINE, paths.OFFLINE):
            self.use(mode)
            self.assertEqual(paths.get("quality"), str(self.d1) + "/")
            self.assertEqual(paths.db_file("品質基準"), str(self.d1) + "/品質基準.sqlite3")
            self.assertEqual(paths.get("nis"), r"C:\Tools\測定データ電子化")         # 既定は VBA のオンラインの値
        paths.set_path("wav", str(self.root / "無い") + "/")
        self.assertEqual(paths.get("wav"), str(self.root / "無い") + "/")           # 無くても同じパス(落ち先なし)
        with self.assertRaises(datasource.DbMissing):
            datasource.get_fields("品質基準", "QA_022_01", "")
        self.put(self.d1, "品質基準")
        self.assertIsNotNone(datasource.get_fields("品質基準", "QA_022_01", ""))


class LotFileTest(_Saved):
    """仕掛台帳: ファイルごとの第2・写せないとき・2つ目の名前・SIKACOMP は第2。"""

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
        (self.d1 / "PLEASE WAIT.txt").write_text("x")
        self.assertEqual(paths.db_file("SIKAHIKINOW"), str(self.d1) + "/SIKAHIKI.sqlite3")       # D26 改: 見ない
        self.assertEqual(paths.db_file("SIKACOMP"), str(self.d2) + "/SIKACOMP.sqlite3")          # VBA は NextPath
        self.assertEqual(paths.db_file("SIKAODRNOW"), str(self.d1) + "/SIKAODR.sqlite3")         # どちらにも無い → 第1
        paths.set_mode(paths.OFFLINE)
        self.assertEqual(paths.db_file("SIKALOTNOW"), str(self.d2) + "/SIKALOT_圧縮.sqlite3")
        self.assertEqual(lot.db_path(paths.get("main"), "SIKALOTNOW"), str(self.d2) + "/SIKALOT_圧縮.sqlite3")

    def test_copy_failure_falls_back(self) -> None:
        self.use(paths.ONLINE)
        self.put(self.d1, "SIKAHIKINOW")
        self.put(self.d2, "SIKAHIKINOW")
        with mock.patch.object(lot_copy.shutil, "copyfile", side_effect=OSError("共有に届かない")):
            lot_copy.forget()
            self.assertFalse(lot.read_check(str(self.d1) + "/SIKAHIKI.sqlite3", 1))
            self.assertEqual(paths.db_file("SIKAHIKINOW"), str(self.d1) + "/SIKAHIKI.sqlite3")   # 第2も写せない → 第1
            with self.assertRaises(datasource.DbMissing):                                       # 共有を直接は開かない
                datasource.get_fields(str(self.d1) + "/SIKAHIKI.sqlite3", "仕掛", "")
            r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL)
            self.assertEqual([m.text for m in r.messages], [paths.missing_text("SIKAHIKINOW")])
        lot_copy.forget()
        real = shutil.copyfile

        def only_second(src, dst, *a, **k):
            if str(self.d1) in str(src):
                raise OSError("第1に届かない")
            return real(src, dst, *a, **k)
        with mock.patch.object(lot_copy.shutil, "copyfile", side_effect=only_second):
            self.assertEqual(paths.db_file("SIKAHIKINOW"), str(self.d2) + "/SIKAHIKI.sqlite3")   # 写せない第1 → 第2
            r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL)
            self.assertEqual(r.file_path, str(self.d2) + "/SIKAHIKI.sqlite3")
            self.assertIsNotNone(r.value)
        lot_copy.forget()

    def test_reads_copy_and_recopies_on_change(self) -> None:
        self.use(paths.ONLINE)
        self.put(self.d1, "SIKAHIKINOW")
        src = str(self.d1) + "/SIKAHIKI.sqlite3"
        c1 = lot_copy.local_copy(src)
        self.assertTrue(str(c1).startswith(str(lot_copy.copy_dir())))
        with mock.patch.object(lot_copy.shutil, "copyfile", side_effect=AssertionError("写し直さない")):
            self.assertEqual(lot_copy.local_copy(src), c1)                    # 変わっていなければ写し直さない
        n = len(datasource.get_records(src, "仕掛", ""))
        with sqlite3.connect(src) as c:
            c.execute('DELETE FROM "仕掛" WHERE rowid = (SELECT MIN(rowid) FROM "仕掛")')
        c.close()
        self.assertEqual(len(datasource.get_records(src, "仕掛", "")), n - 1)   # 変わったら写し直して読む

    def test_second_name_rules(self) -> None:
        self.use(paths.ONLINE)
        self.assertEqual(paths.second_name("SIKAHIKINOW"), "SIKAHIKI.sqlite3")
        paths.set_second_name("SIKAHIKINOW", "HIKI2.db")
        self.assertEqual(paths.second_name("SIKAHIKINOW"), "HIKI2.db")                          # 拡張子まで入れればそのまま
        self.assertEqual(lot.db_path(paths.get("next"), "SIKAHIKINOW"), str(self.d2) + "/HIKI2.db")
        self.assertEqual(lot.db_path(paths.get("main"), "SIKAHIKINOW"), str(self.d1) + "/SIKAHIKI.sqlite3")
        self.assertTrue(paths.is_copied(str(self.d2) + "/HIKI2.db"))
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

    def rename_columns(self, path: Path, mapping: dict[str, str], drop: tuple[str, ...] = ()) -> None:
        with sqlite3.connect(path) as c:
            for a, b in mapping.items():
                c.execute(f'ALTER TABLE "仕掛" RENAME COLUMN "{a}" TO "{b}"')
            for d in drop:
                c.execute(f'ALTER TABLE "仕掛" DROP COLUMN "{d}"')
        c.close()

    def test_second_with_other_column_names_offline(self) -> None:
        """2つ目のファイルの列名が全角カナなど書き方違い(python-web-tools と同じ見比べ方)でも、第1の名前で読める。"""
        self.use(paths.ONLINE)
        base = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL).value
        self.use(paths.OFFLINE, second_names={"SIKAHIKINOW": "SIKAHIKI_2"})
        (self.d2 / "SIKAHIKI.sqlite3").rename(self.d2 / "SIKAHIKI_2.sqlite3")
        self.rename_columns(self.d2 / "SIKAHIKI_2.sqlite3", {"ﾛｯﾄ番号": "ロット番号", "引当番号": "引当番号 "})
        r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL)
        self.assertEqual(r.file_path, str(self.d2) + "/SIKAHIKI_2.sqlite3")
        self.assertEqual(r.value, base)
        f = lot.arrbox("SIKAHIKINOW", "", "Fields", "", now=T_NORMAL).value
        self.assertIn("ﾛｯﾄ番号", f)

    def test_acc_reflec_second_other_column_names(self) -> None:
        """ロット登録(AccReflec の SELECT)も 2つ目の列名の書き方違いで読める(オフライン = 第2だけ)。"""
        screen = lot.LotScreen(lot="H2500K0", cb_text="H2500K0", work_time="T")
        self.use(paths.OFFLINE)
        base = lot.acc_reflec(screen, check_equipment=_ok)
        self.use(paths.OFFLINE, second_names={"SIKALOTNOW": "LOT2"})
        (self.d2 / "SIKALOT.sqlite3").rename(self.d2 / "LOT2.sqlite3")
        self.rename_columns(self.d2 / "LOT2.sqlite3", {"ﾛｯﾄ番号": "ロット番号", "ｵｰﾀﾞｰ板丈": "オーダー板丈"})
        r = lot.acc_reflec(screen, check_equipment=_ok)
        self.assertEqual((r.status, r.db_path), ("ok", str(self.d2) + "/LOT2.sqlite3"))
        self.assertEqual(r.past_hiki, base.past_hiki)

    def test_check_typed_values_and_missing_columns(self) -> None:
        """確認ボタン: 保存前に入力中のパス・2つ目の名前を確かめ、2つ目の足りない列を出す。"""
        self.use(paths.ONLINE)
        (self.d2 / "SIKALOT.sqlite3").rename(self.d2 / "LOT2.sqlite3")
        self.rename_columns(self.d2 / "LOT2.sqlite3", {"ﾛｯﾄ番号": "ロット番号"}, drop=("製造材質",))
        saved = paths.check(paths.LOT, paths.SECOND)
        self.assertIn("SIKALOT.sqlite3 なし", saved["message"])
        r = paths.check(paths.LOT, paths.SECOND, text=str(self.d2), names={"SIKALOTNOW": "LOT2"})
        f = r["files"][0]
        self.assertEqual((f["name"], f["copied"], f["lacking"]), ("LOT2.sqlite3", True, ["製造材質"]))
        self.assertIn("足りない列", r["message"])
        r = paths.check(paths.LOT, paths.FIRST, text=str(self.d2 / "無い"))
        self.assertEqual(r["message"], "見つかりません")

    def test_second_missing_columns_read_as_blank(self) -> None:
        """2つ目のファイルに列が足りなくても(実機: SIKALOT_圧縮 は 46 列)止まらず、無い列は空として読む。"""
        from ufkensa.kensa import quality, datasource as ds
        screen = lot.LotScreen(lot="H2500K0", cb_text="H2500K0", work_time="T")
        self.use(paths.OFFLINE, second_names={"SIKALOTNOW": "LOT2"})
        (self.d2 / "SIKALOT.sqlite3").rename(self.d2 / "LOT2.sqlite3")
        self.rename_columns(self.d2 / "LOT2.sqlite3", {}, drop=("指示_中歪_高さ", "品質ｸﾞﾚｰﾄﾞ_直角度", "BOX設計_横割数(板)"))
        r = lot.acc_reflec(screen, check_equipment=_ok)
        self.assertEqual(r.status, "ok")
        f = lot.arrbox("SIKALOTNOW", "板", "Fields").value
        rec = lot.arrbox("SIKALOTNOW", "H2500K0", "Record").value
        self.assertEqual(quality.cell(rec[0], quality.col_of(f, "指示_中歪_高さ")), "")
        self.assertIn("ﾛｯﾄ番号", ds.get_fields(str(self.d2 / "LOT2.sqlite3"), "仕掛"))

    def test_first_with_other_name(self) -> None:
        """第1パスにも名前を決められる(オンライン: 第1の別名のファイルでロット登録できる)。"""
        screen = lot.LotScreen(lot="H2500K0", cb_text="H2500K0", work_time="T")
        self.use(paths.ONLINE)
        base = lot.acc_reflec(screen, check_equipment=_ok)
        for n in ("SIKALOT", "SIKAHIKI", "SIKAODR"):
            (self.d2 / f"{n}.sqlite3").unlink()
            (self.d1 / f"{n}.sqlite3").rename(self.d1 / f"{n}_1.sqlite3")
        self.use(paths.ONLINE, first_names={"SIKALOTNOW": "SIKALOT_1", "SIKAHIKINOW": "SIKAHIKI_1",
                                            "SIKAODRNOW": "SIKAODR_1"})
        self.assertEqual(paths.first_name("SIKALOTNOW"), "SIKALOT_1.sqlite3")
        r = lot.acc_reflec(screen, check_equipment=_ok)
        self.assertEqual((r.status, r.db_path), ("ok", str(self.d1) + "/SIKALOT_1.sqlite3"))
        self.assertEqual(r.past_hiki, base.past_hiki)
        c = paths.check(paths.LOT, paths.FIRST)
        self.assertIn("SIKALOT_1.sqlite3 あり", c["message"])
        c = paths.check(paths.LOT, paths.FIRST, first_names={"SIKALOTNOW": ""})
        self.assertIn("SIKALOT.sqlite3 なし", c["message"])
        paths.set_name("SIKALOTNOW", paths.FIRST, "")
        self.assertEqual(paths.first_name("SIKALOTNOW"), "SIKALOT.sqlite3")

    def test_offline_reads_second_then_first(self) -> None:
        # 2026-11-09(ご依頼): オフラインは第2を読み、第2に無いもの(ロット・ファイル)は第1を見に行く
        self.use(paths.OFFLINE)
        r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL)
        self.assertEqual(r.file_path, str(self.d2) + "/SIKAHIKI.sqlite3")
        self.delete(self.d2, "SIKAHIKINOW", "ﾛｯﾄ番号", ["H2500K0"])
        r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL)
        self.assertIsNotNone(r.value)
        self.assertIn("first", r.file_path)
        (self.d2 / "SIKAHIKI.sqlite3").unlink()
        r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL)
        self.assertIsNotNone(r.value)
        (self.d1 / "SIKAHIKI.sqlite3").unlink()
        r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL)
        self.assertEqual([m.text for m in r.messages], [paths.missing_text("SIKAHIKINOW")])

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
    """前の保存 → 仕掛台帳は第1/第2、ほかはいまのモードで使っていた値。"""

    OLD = {paths.ONLINE: {"main": r"\\a\台帳" "\\", "next": r"\\a\控え" "\\", "wav": r"D:\音1"},
           paths.OFFLINE: {"main": r"C:\app" "\\", "next": r"C:\app" "\\", "wav": r"D:\音2", "nis": ""}}

    def test_old_online_offline_in_online(self) -> None:
        paths.set_mode(paths.ONLINE)
        user_settings.save(paths.KEY_PATHS, self.OLD)
        self.assertEqual(paths.slot_value(paths.LOT, paths.FIRST), r"\\a\台帳" "\\")
        self.assertEqual(paths.slot_value(paths.LOT, paths.SECOND), r"\\a\控え" "\\")   # VBA オンラインの NextPath
        self.assertEqual(paths.get("wav"), r"D:\音1")
        self.assertEqual(paths.get("nis"), paths.slot_default("nis"))
        stored = user_settings.get(paths.KEY_PATHS)
        self.assertNotIn(paths.ONLINE, stored)                                          # 新しい形で保存し直す
        self.assertEqual(stored[paths.KEY_ITEMS]["wav"], r"D:\音1")

    def test_old_online_offline_in_offline(self) -> None:
        paths.set_mode(paths.OFFLINE)
        user_settings.save(paths.KEY_PATHS, self.OLD)
        self.assertEqual((paths.get("wav"), paths.get("nis")), (r"D:\音2", ""))         # オフラインで使っていた値
        self.assertEqual(paths.get("main"), r"\\a\控え" "\\")

    def test_interim_first_second(self) -> None:
        paths.set_mode(paths.OFFLINE)
        user_settings.save(paths.KEY_PATHS, {paths.FIRST: {"lot": "/1/", "quality": "/q1/"},
                                             paths.SECOND: {"lot": "/2/", "quality": "/q2/"},
                                             "second_names": {"SIKALOTNOW": "L2"}})
        self.assertEqual((paths.slot_value(paths.LOT, paths.FIRST), paths.slot_value(paths.LOT, paths.SECOND)),
                         ("/1/", "/2/"))
        self.assertEqual(paths.get("quality"), "/q2/")
        self.assertEqual(paths.second_name("SIKALOTNOW"), "L2.sqlite3")

    def test_old_call_style(self) -> None:
        paths.set_path("wav", r"D:\x", paths.ONLINE)                 # slot はパス1つの項目では見ない
        self.assertEqual(paths.get("wav"), r"D:\x")
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
        setup = mock.patch.object(auth, "setup_mode", return_value=False)      # マスタで決まっている(設定の途中でない)
        setup.start()
        self.addCleanup(setup.stop)
        with mock.patch.object(auth, "can_edit", return_value=True), \
                mock.patch.object(auth, "master_edit_ok", return_value=False):
            r = self.post("/api/settings/path", key="wav", slot="first", path="x")
            self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "locked"))
            r = self.post("/api/settings/path", key="lot", db="SIKALOTNOW", second_name="x")
            self.assertEqual(r.status_code, 403)
        with mock.patch.object(auth, "can_edit", return_value=False):
            r = self.post("/api/settings/path", key="wav", slot="second", path="x")
            self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "not_admin"))
            self.assertEqual(r.get_json()["error"]["message"], "編集には、アクセス権限の「管理者」(Administrator)が要ります")
        self.assertNotEqual(paths.slot_value("wav", paths.FIRST), "x")

    def test_edit_when_unlocked(self) -> None:
        with mock.patch.object(auth, "master_edit_ok", return_value=True):
            st = self.post("/api/settings/path", key="wav", path=str(self.d2) + "/").get_json()
            row = [p for p in st["paths"] if p["key"] == "wav"][0]
            self.assertEqual((row["path"], row["changed"], row["single"]), (str(self.d2) + "/", True, True))
            st = self.post("/api/settings/path", key="lot", slot="second", path=str(self.root) + "/").get_json()
            lot_row = [p for p in st["paths"] if p["key"] == "lot"][0]
            self.assertEqual((lot_row["second"], lot_row["second_changed"]), (str(self.root) + "/", True))
            st = self.post("/api/settings/path", key="lot", db="SIKAHIKINOW", second_name="HIKI2").get_json()
            lot_row = [p for p in st["paths"] if p["key"] == "lot"][0]
            self.assertEqual([n["effective"] for n in lot_row["second_names"]],
                             ["SIKALOT.sqlite3", "HIKI2.sqlite3", "SIKAODR.sqlite3"])
            self.assertEqual(lot_row["in_use"][paths.OFFLINE], "第2だけ")
            self.assertEqual(self.post("/api/settings/path", key="lot", path="x").status_code, 400)   # 仕掛は slot が要る

    def test_check_both(self) -> None:
        self.put(self.d2, "SIKALOTNOW")
        d = self.post("/api/settings/path/check", key="lot").get_json()
        first, second = d["results"]
        self.assertTrue(first["ok"] and second["ok"])
        self.assertEqual([f["copied"] for f in second["files"]][:3], [True, False, False])
        self.assertIn("第1パス", d["message"])
        d = self.post("/api/settings/path/check", key="quality").get_json()
        self.assertEqual(len(d["results"]), 1)
        # 保存前に入力中の値を確かめる(画面の「確認」は入力欄の値を送る)
        d = self.post("/api/settings/path/check", key="quality", path=str(self.root / "無い")).get_json()
        self.assertEqual((d["results"][0]["path"], d["results"][0]["message"]), (str(self.root / "無い"), "見つかりません"))
        d = self.post("/api/settings/path/check", key="lot", first=str(self.root / "無い"), second=str(self.d2),
                      second_names={"SIKALOTNOW": ""}).get_json()
        self.assertEqual([r["ok"] for r in d["results"]], [False, True])


class JudgeImagesTest(_Saved):
    """2026-11-09: 判定の画像(VBA の UFOK / UFNG / UF対角 の絵)は枠ごとのファイル。空・無いときは同梱の絵。"""

    DIAG = [f"judge_diag{i}" for i in range(1, 6)]

    def client(self):
        from app import create_app
        app = create_app("main", token="t", port=0)
        app.config["READY"] = True
        return app.test_client(), {"X-Tool-Token": "t"}

    def test_unset_gives_bundled_vba_pictures(self) -> None:
        from ufkensa.kensa import judge_images
        user_settings.delete(paths.KEY_PATHS)
        self.assertEqual(judge_images.available(), {"ok": ["judge_ok"], "ng": ["judge_ng"], "diag": self.DIAG})
        self.assertEqual([f.name for f in judge_images.effective("diag")], [f"対角{i}.jpg" for i in range(1, 6)])
        self.assertEqual(judge_images.resolve("judge_ok"), (judge_images.DEFAULT_DIR / "OK.jpg", "default"))
        self.assertTrue(paths.check("judge_ok")["ok"])

    def test_any_file_name_and_diag_pool(self) -> None:
        from ufkensa.kensa import judge_images
        a = self.d1 / "よくできた.PNG"; a.write_bytes(b"\x89PNG")
        b = self.d1 / "x.gif"; b.write_bytes(b"GIF8")
        paths.set_path("judge_ok", str(a))
        paths.set_path("judge_diag3", str(b))
        self.assertEqual(judge_images.resolve("judge_ok"), (a, "set"))
        self.assertEqual(judge_images.available()["diag"], [f"judge_diag{i}" for i in range(1, 6)])   # 決めた枠と同梱の絵を混ぜて乱数
        eff = judge_images.effective("diag")                                     # 対角3 は決めた画像、ほかは同梱の絵
        self.assertEqual((len(eff), eff[2]), (5, b))
        self.assertEqual(eff[0].name, "対角1.jpg")
        self.assertEqual(judge_images.resolve("judge_ng")[1], "default")
        self.assertEqual(judge_images.find("diag", "judge_diag1").name, "対角1.jpg")   # 空の枠は同梱の絵
        row = [r for r in paths.overview() if r["key"] == "judge_ok"][0]
        self.assertEqual((row["kind"], row["image"]["source"]), ("file", "set"))
        self.assertNotIn("judge_images", paths.BY_KEY)                             # フォルダの項目は一覧に出さない

    def test_missing_file_gives_bundled_and_note(self) -> None:
        from ufkensa.kensa import judge_images
        paths.set_path("judge_ng", str(self.d1 / "無い.jpg"))
        f, src = judge_images.resolve("judge_ng")
        self.assertEqual((f.name, src), ("NG.jpg", "missing"))
        self.assertIn("ファイルがありません", judge_images.slot_info("judge_ng")["note"])
        self.assertIn("ファイルがありません", paths.check("judge_ng")["message"])
        c, h = self.client()
        r = c.get("/api/kensa/judge_image/judge_ng", headers=h)
        self.assertEqual((r.status_code, r.headers["X-Judge-Source"]), (200, "missing"))
        self.assertEqual(r.data, (judge_images.DEFAULT_DIR / "NG.jpg").read_bytes())

    def test_route_serves_only_configured_slot(self) -> None:
        from ufkensa.kensa import judge_images
        a = self.d1 / "mine.png"; a.write_bytes(b"\x89PNG-mine")
        (self.d1 / "secret.png").write_bytes(b"\x89PNG-secret")
        paths.set_path("judge_ok", str(a))
        c, h = self.client()
        r = c.get("/api/kensa/judge_image/judge_ok", headers=h)
        self.assertEqual((r.status_code, r.data, r.headers["X-Judge-Source"]), (200, b"\x89PNG-mine", "set"))
        self.assertEqual(c.get("/api/kensa/judge_image/judge_ng", headers=h).data,
                         (judge_images.DEFAULT_DIR / "NG.jpg").read_bytes())          # 空 → 同梱
        for bad in ("secret.png", "..", "judge_images", "OK.jpg"):
            self.assertEqual(c.get(f"/api/kensa/judge_image/{bad}", headers=h).status_code, 404, bad)
        self.assertEqual(c.get("/api/kensa/judge_image/ok/secret.png", headers=h).status_code, 404)
        self.assertEqual(c.get("/api/kensa/judge_image/ng/judge_ok", headers=h).status_code, 404)
        self.assertEqual(c.get("/api/kensa/judge_image/ok/judge_ok", headers=h).status_code, 200)
        d = c.get("/api/kensa/judge_images", headers=h).get_json()
        self.assertEqual(d["images"]["ok"], ["judge_ok"])
        self.assertEqual(d["show_ms"], {"ok": 3000, "ng": 3000, "diag": 2000})

    def test_migrates_old_folder_once(self) -> None:
        from ufkensa.kensa import judge_images
        (self.d1 / "ok.JPG").write_bytes(b"\xff\xd8\xff")
        (self.d1 / "対角2.png").write_bytes(b"\x89PNG")
        (self.d1 / "不合格のとき.bmp").write_bytes(b"BM")        # 決まった名前でなくても言葉で
        user_settings.save(paths.KEY_PATHS, {paths.KEY_ITEMS: {"judge_images": str(self.d1), "judge_ok": str(self.d2 / "既.jpg")}})
        self.assertEqual(paths.get("judge_ok"), str(self.d2 / "既.jpg"))         # 決めてある枠は変えない
        self.assertEqual(paths.get("judge_ng"), paths.join(str(self.d1), "不合格のとき.bmp"))
        self.assertEqual(paths.get("judge_diag1"), paths.join(str(self.d1), "対角2.png"))
        saved = user_settings.get(paths.KEY_PATHS)[paths.KEY_ITEMS]
        self.assertNotIn("judge_images", saved)                                  # 1 回だけ
        self.assertEqual(judge_images.available()["diag"], [f"judge_diag{i}" for i in range(1, 6)])
