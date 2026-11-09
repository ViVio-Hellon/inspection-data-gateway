"""ロットごとの「控え(第2)から読む」(判断事項 D28。VBA の DBflag の運用の代わり)

第1(台帳)と第2(控え)で同じロット H7345B0 の内容を変えて(第2 は BOX実績_枚本数 = 999 → KenH「前工程実績999枚」)、
入/切 × オンライン/オフライン × ロットが第1だけ/第2だけ/両方 で、どちらから読んだかを確かめる。
作業保存での持ち越し、新規検査・別のロットの登録で外れることも確かめる。
"""
from __future__ import annotations

import shutil
import sqlite3
import unittest

from tests import kensa_fixture
from tests.test_kensa_session import LOT, SessionBase
from ufkensa import user_settings
from ufkensa.kensa import lot as lotmod
from ufkensa.kensa import paths, quality_session, records, wiring
from ufkensa.kensa.store import HOZON

KEN_FIRST = "前工程実績17枚"           # 基準データの BOX実績_枚本数(第1)
KEN_SECOND = "前工程実績999枚"         # 第2だけ書き換える
OTHER = "RX021T0"                     # 別のロット
REG = {"lot_register": True, "lot_split": False, "lot_recent": True}


class _TwoSlots(SessionBase):
    """仕掛台帳の第1 = db_dir/first、第2 = db_dir/second(ほかの項目は db_dir)。"""

    def setUp(self) -> None:
        super().setUp()
        self.first = self.db_dir / "first"
        self.second = self.db_dir / "second"
        for d in (self.first, self.second):
            d.mkdir()
            for name in paths.LOT_DBS:
                shutil.copy2(self.db_dir / kensa_fixture.file_name(name), d)
        with sqlite3.connect(self.second / "SIKALOT.sqlite3") as c:
            c.execute('UPDATE "仕掛" SET "BOX実績_枚本数" = ? WHERE "ﾛｯﾄ番号" = ?', ("999", LOT))
        d = user_settings.get(paths.KEY_PATHS, {})
        d[paths.KEY_LOT] = {paths.FIRST: str(self.first) + "/", paths.SECOND: str(self.second) + "/"}
        user_settings.save(paths.KEY_PATHS, d)
        paths.set_mode(paths.ONLINE)

    def delete_lot(self, where, lot: str = LOT) -> None:
        for name in paths.LOT_DBS:
            with sqlite3.connect(where / kensa_fixture.file_name(name)) as c:
                c.execute('DELETE FROM "仕掛" WHERE "ﾛｯﾄ番号" = ?', (lot,))

    def ken_h(self) -> str:
        return self.s.state.v("KenH").value

    def lot_source_on(self, lot: str = "") -> None:
        r = self.s.set_lot_source(True, lot, REG)
        self.assertIsNone(r.stop, r.stop and r.stop.text)

    def reregister(self) -> None:
        r = self.s.register_lot(LOT, REG)
        self.assertIsNone(r.stop, r.stop and r.stop.text)


class MatrixTest(_TwoSlots):
    """入/切 × モード × ロットの在りか。"""

    def test_off_online_both_reads_first(self) -> None:
        self.register()
        self.assertEqual(self.ken_h(), KEN_FIRST)
        self.assertFalse(self.s.current().view["lot_source"]["active"])

    def test_on_online_both_reads_second(self) -> None:
        self.register()
        self.lot_source_on()                                  # いまのロット → すぐ読み直す
        self.assertEqual(self.ken_h(), KEN_SECOND)
        v = self.s.current().view["lot_source"]
        self.assertEqual((v["on"], v["lot"], v["active"]), (True, LOT, True))
        self.reregister()                                     # 同じロットの登録し直しでも控えのまま
        self.assertEqual(self.ken_h(), KEN_SECOND)
        self.assertEqual(self.s.state.lot_source, "second")

    def test_off_offline_reads_second(self) -> None:
        paths.set_mode(paths.OFFLINE)
        self.register()
        self.assertEqual(self.ken_h(), KEN_SECOND)

    def test_on_offline_reads_second(self) -> None:
        paths.set_mode(paths.OFFLINE)
        self.register()
        self.lot_source_on()
        self.assertEqual(self.ken_h(), KEN_SECOND)

    def test_mode_switch_keeps_second(self) -> None:
        """オフラインで控えを選んだロットは、オンラインに切り替えても第1で置き換わらない。"""
        paths.set_mode(paths.OFFLINE)
        self.register()
        self.lot_source_on()
        paths.set_mode(paths.ONLINE)
        self.reregister()
        self.assertEqual(self.ken_h(), KEN_SECOND)
        self.assertEqual(self.s.state.lot_source_lot, LOT)

    def test_mode_switch_without_choice_replaces(self) -> None:
        """選んでいなければ D26 どおり(オンラインは第1)。"""
        paths.set_mode(paths.OFFLINE)
        self.register()
        self.assertEqual(self.ken_h(), KEN_SECOND)
        paths.set_mode(paths.ONLINE)
        self.reregister()
        self.assertEqual(self.ken_h(), KEN_FIRST)

    def test_lot_only_in_second(self) -> None:
        self.delete_lot(self.first)
        self.register()                                       # 切: D26 のロットごとの第2
        self.assertEqual(self.ken_h(), KEN_SECOND)
        self.lot_source_on()
        self.assertEqual(self.ken_h(), KEN_SECOND)

    def test_lot_only_in_first_on_reads_first(self) -> None:
        self.delete_lot(self.second)
        self.register()
        self.assertEqual(self.ken_h(), KEN_FIRST)
        self.s.set_lot_source(True, "", REG)
        # 2026-11-09(ご依頼): 第2に無いものは第1を見に行く(オフライン・控えから読む でも)
        self.assertEqual(self.ken_h(), KEN_FIRST)
        r = self.s.set_lot_source(False, "", REG)             # 外すと いつもの読み方(第1)で読み直す
        self.assertEqual(self.ken_h(), KEN_FIRST)
        self.assertFalse(r.view["lot_source"]["on"])

    def test_on_with_typed_lot_registers_it(self) -> None:
        """登録前に LotNo を入れて「控えから読む」→ そのロットを控えから登録する(台帳にロットが無いとき)。"""
        self.delete_lot(self.first)
        self.ok(self.s.start())
        self.ok(self.s.register_worker("9000000000001", {"worker_register": True}))
        self.lot_source_on(LOT.lower())
        self.assertEqual(self.s.state.lot, LOT)
        self.assertEqual(self.ken_h(), KEN_SECOND)
        self.assertTrue(self.s.current().view["lot_source"]["active"])

    def test_on_without_lot_stops(self) -> None:
        self.ok(self.s.start())
        r = self.s.set_lot_source(True, "", {})
        self.assertIsNotNone(r.stop)
        self.assertIn("LotNo", r.stop.text)

    def test_off_rereads_first(self) -> None:
        self.register()
        self.lot_source_on()
        self.assertEqual(self.ken_h(), KEN_SECOND)
        self.s.set_lot_source(False, "", REG)
        self.assertEqual(self.ken_h(), KEN_FIRST)
        self.assertEqual((self.s.state.lot_source, self.s.state.lot_source_lot), ("", ""))


class LifecycleTest(_TwoSlots):
    def test_saved_and_restored(self) -> None:
        self.register()
        self.lot_source_on()
        sheet = self.store.sheet("作業保存_" + self.s.state.lname)
        self.assertEqual(sheet.get(*records.LOT_SOURCE_CELL), LOT)
        s2 = self.new_session()                               # 開き直し
        self.ok(s2.start())
        self.assertEqual((s2.state.lot, s2.state.lot_source, s2.state.lot_source_lot), (LOT, "second", LOT))
        self.s = s2
        self.reregister()
        self.assertEqual(self.ken_h(), KEN_SECOND)

    def test_not_saved_when_off(self) -> None:
        self.register()
        sheet = self.store.sheet("作業保存_" + self.s.state.lname)
        self.assertIn(sheet.get(*records.LOT_SOURCE_CELL), (None, ""))
        self.assertEqual(records.load_lot_source(self.store, self.s.state.lname), "")

    def test_cleared_by_new_inspection(self) -> None:
        self.register()
        self.lot_source_on()
        sh = self.store.sheet(HOZON)                          # 保存済み(でないと新規検査は止まる。UF1:3380)
        sh.set(2, 1, "HVC")
        sh.set(2, 2, LOT)
        r = self.s.new_inspection({"new": True})
        self.assertEqual((self.s.state.lot_source, self.s.state.lot_source_lot), ("", ""))
        self.assertFalse(r.view["lot_source"]["on"])
        s2 = self.new_session()
        self.ok(s2.start())
        self.assertEqual(s2.state.lot_source, "")

    def test_cleared_by_other_lot(self) -> None:
        self.register()
        self.lot_source_on()
        self.s.register_lot(OTHER, REG)
        self.assertEqual(self.s.state.lot, OTHER)
        self.assertEqual((self.s.state.lot_source, self.s.state.lot_source_lot), ("", ""))
        self.s.register_lot(LOT, REG)                         # 戻しても既定は切(第1)
        self.assertEqual(self.ken_h(), KEN_FIRST)


class OtherScreensTest(_TwoSlots):
    def test_quality_screen_follows(self) -> None:
        self.register()
        out = quality_session.build("", store=self.store, kensa=self.s, facility="HVC")
        self.assertFalse(out["lot_source_second"])
        self.lot_source_on()
        out = quality_session.build("", store=self.store, kensa=self.s, facility="HVC")
        self.assertTrue(out["lot_source_second"])
        self.assertTrue(out["found"])
        # 別のロットは いつもの読み方
        self.assertFalse(quality_session.build(OTHER, store=self.store, kensa=self.s, facility="HVC")
                         ["lot_source_second"])

    def test_quality_screen_second_only_in_second(self) -> None:
        self.register()
        self.lot_source_on()
        self.delete_lot(self.second)
        out = quality_session.build("", store=self.store, kensa=self.s, facility="HVC")
        self.assertTrue(out["found"])                         # 2026-11-09: 第2に無ければ第1を見に行く


    def test_cask_and_zenkasho_follow_same_lot(self) -> None:
        """キャスク・全ヵ所の操作は、そのロットが「控えから読む」のロットなら第2だけ(HistorySave・ロットの読みとも)。"""
        from ufkensa.kensa import cask_session, zenkasho_session
        self.register()
        self.lot_source_on()
        st = self.s.state
        seen: list[bool] = []
        cs = cask_session.CaskSession(self.store, now=self.clock, uf1=lambda: st, write_logs=False)
        for lot, want in ((LOT, True), (OTHER, False)):
            def fn(lot: str = lot) -> None:
                cs.book.cask.lot_box = lot
                seen.append(paths.lot_second_forced())
            cs._run(fn)
            self.assertEqual(seen[-1], want)
        zs = zenkasho_session.ZenkashoSession(self.store, now=self.clock, kensa_state=lambda: st, write_logs=False)
        for lot, want in ((LOT, True), (OTHER, False)):
            def fz(lot: str = lot) -> None:
                zs.book.mit.lot_box = lot
                seen.append(paths.lot_second_forced())
            zs._run(fz)
            self.assertEqual(seen[-1], want)


class PathsTest(unittest.TestCase):
    """paths.lot_source の読み方(モードによらず第2だけ)。"""

    def setUp(self) -> None:
        self.saved = (user_settings.get(paths.KEY_MODE), user_settings.get(paths.KEY_PATHS))
        user_settings.save(paths.KEY_PATHS, {paths.KEY_LOT: {paths.FIRST: "/a/", paths.SECOND: "/b/"}})

    def tearDown(self) -> None:
        user_settings.save(paths.KEY_MODE, self.saved[0] or paths.ONLINE)
        user_settings.save(paths.KEY_PATHS, self.saved[1] or {})

    def test_main_is_second_when_forced(self) -> None:
        for m in (paths.ONLINE, paths.OFFLINE):
            paths.set_mode(m)
            with paths.lot_source(True):
                self.assertEqual(paths.get("main"), "/b/")
                self.assertEqual(paths.get("main"), paths.get("next"))
                self.assertTrue(paths.db_file("SIKALOTNOW").startswith("/b/"))
            self.assertEqual(paths.mode(), m)                 # 設定のモードは変えない
        paths.set_mode(paths.ONLINE)
        self.assertEqual(paths.get("main"), "/a/")
        with paths.lot_source(False):
            self.assertEqual(paths.get("main"), "/a/")

    def test_callable_is_evaluated_each_time(self) -> None:
        paths.set_mode(paths.ONLINE)
        flag = {"on": False}
        with paths.lot_source(lambda: flag["on"]):
            self.assertEqual(paths.get("main"), "/a/")
            flag["on"] = True
            self.assertEqual(paths.get("main"), "/b/")
        self.assertFalse(paths.lot_second_forced())

    def test_explicit_mode_wins(self) -> None:
        with paths.lot_source(True):
            self.assertEqual(paths.get("main", paths.ONLINE), "/a/")

    def test_reads_second_helper(self) -> None:
        class S:
            lot_source, lot_source_lot = "second", LOT
        self.assertTrue(wiring.reads_second(S, LOT.lower()))
        self.assertFalse(wiring.reads_second(S, OTHER))
        self.assertFalse(wiring.reads_second(S, ""))
        self.assertFalse(wiring.reads_second(None, LOT))


class ArrboxTest(_TwoSlots):
    def test_arrbox_lot_only_in_first(self) -> None:
        self.delete_lot(self.second)
        r = lotmod.arrbox("SIKALOTNOW", LOT, "Record", now=self.clock.t)
        self.assertIsNotNone(r.value)
        with paths.lot_source(True):
            r = lotmod.arrbox("SIKALOTNOW", LOT, "Record", now=self.clock.t)
        self.assertIsNotNone(r.value)                          # 2026-11-09: 第2に無ければ第1を見に行く
        self.assertIn("first", r.file_path)


class RouteTest(_TwoSlots):
    def test_route(self) -> None:
        from tests import _web
        from ufkensa.kensa import session
        session.reset_for_test(self.s)
        self.register()
        c = _web.make_client()
        res = c.post("/api/kensa/lot-source", json={"on": True, "lot": "", "answers": REG}, headers=_web.auth())
        self.assertEqual(res.status_code, 200)
        body = res.get_json()
        self.assertTrue(body["view"]["lot_source"]["active"])
        self.assertEqual(self.ken_h(), KEN_SECOND)
        page = c.get("/kensa", headers=_web.auth()).get_data(as_text=True)
        self.assertIn('id="lotSrcBtn"', page)
        self.assertIn('id="lotSrcLv"', page)


if __name__ == "__main__":
    unittest.main()
