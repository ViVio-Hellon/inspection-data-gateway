"""ロット引当・仕掛 DB の切替(ufkensa.kensa.lot)

期待値は VBA を手で追ったもの(根拠の行番号はコメント)。DB は受領 DB から作った fixtures(SIKA* は 108 ロット)。
"""
from __future__ import annotations

import datetime as dt
import shutil
import sqlite3
import tempfile
import unittest
from pathlib import Path

from tests import kensa_fixture
from ufkensa import user_settings
from ufkensa.kensa import datasource, flow, lot, paths, quirks, vbx

T_NORMAL = dt.datetime(2026, 10, 30, 10, 30, 0)     # 時刻窓の外
T_WINDOW = dt.datetime(2026, 10, 30, 10, 59, 50)    # 毎時 59:46〜59:59 の中


def _ok(*_a) -> str:
    return "可"


class _Dirs:
    """main / next / packing を別フォルダにした使用パス(オフライン)。"""

    def __init__(self, *, main: bool = True, nxt: bool = True) -> None:
        self.root = Path(tempfile.mkdtemp(prefix="ufkensa_lot_"))
        self.main = self.root / "main"
        self.next = self.root / "next"
        self.pack = self.root / "pack"
        for d in (self.main, self.next, self.pack):
            d.mkdir()
        src = kensa_fixture.build()
        for name in ("SIKALOTNOW", "SIKAHIKINOW", "SIKAODRNOW"):
            if main:
                shutil.copy2(src / kensa_fixture.file_name(name), self.main)
            if nxt:
                shutil.copy2(src / kensa_fixture.file_name(name), self.next)
        shutil.copy2(src / "伝送用ファイル.sqlite3", self.main)
        # D26: 仕掛台帳の第1 = MainPath(台帳)、第2 = NextPath(控え)。VBA の切替はオンラインで確かめる
        paths.set_mode(paths.ONLINE)
        user_settings.save(paths.KEY_PATHS, {
            paths.KEY_ITEMS: {"quality": str(self.main) + "/", "packing": str(self.pack) + "/",
                              "history": str(self.root) + "/", "factory": str(self.root) + "/"},
            paths.KEY_LOT: {paths.FIRST: str(self.main) + "/", paths.SECOND: str(self.next) + "/"}})

    def please_wait(self) -> None:
        (self.main / "PLEASE WAIT.txt").write_text("x")             # D26 改: 置いても見ない

    def remove(self, where: Path, name: str) -> None:
        (where / kensa_fixture.file_name(name)).unlink()

    def delete_lot(self, where: Path, name: str, lot_no: str) -> None:
        with sqlite3.connect(where / kensa_fixture.file_name(name)) as c:
            c.execute('DELETE FROM "仕掛" WHERE "ﾛｯﾄ番号" = ?', (lot_no,))

    def cleanup(self) -> None:
        shutil.rmtree(self.root, ignore_errors=True)


class _Base(unittest.TestCase):
    def setUp(self) -> None:
        self.d = _Dirs()

    def tearDown(self) -> None:
        self.d.cleanup()


# ------------------------------------------------------------------
class ReadCheckTest(_Base):
    """D26 改: VBA の TimeDesignation / FileDesignation(上流の作り直しの決まり)は外した。"""

    def test_read_check(self) -> None:
        self.assertTrue(lot.read_check(str(self.d.main / "SIKALOT.sqlite3"), 1))
        self.assertFalse(lot.read_check(str(self.d.main / "無い.sqlite3"), 1))
        self.assertFalse(lot.read_check(str(self.d.main), 1))        # フォルダはファイルではない
        self.assertTrue(lot.read_check(str(self.d.main), 2))
        self.assertFalse(lot.read_check(str(self.d.main), 3))


class IsOver12HoursTest(unittest.TestCase):
    def tearDown(self) -> None:
        pass
    def test_reproduce_b19_one_hour_boundary(self) -> None:
        # Worker:342 `DateDiff("h", C1, Now) >= 1` は時の境界の数
        self.assertFalse(lot.is_over_12_hours("2026/10/30 10:00:00", dt.datetime(2026, 10, 30, 10, 59, 59)))
        self.assertTrue(lot.is_over_12_hours("2026/10/30 10:59:59", dt.datetime(2026, 10, 30, 11, 0, 0)))
        self.assertFalse(lot.is_over_12_hours("2026/10/30 12:00:00", dt.datetime(2026, 10, 30, 11, 0, 0)))


    def test_empty_and_errors(self) -> None:
        now = dt.datetime(2026, 10, 30, 12, 0, 0)
        self.assertFalse(lot.is_over_12_hours("2026/10/29 1:00:00", now, sheet_exists=False))   # Worker:333
        self.assertFalse(lot.is_over_12_hours("", now))                                          # Worker:337
        self.assertFalse(lot.is_over_12_hours(None, now))
        self.assertTrue(lot.is_over_12_hours(dt.datetime(2026, 10, 29, 23, 0), now))
        with self.assertRaises(vbx.TypeMismatch):                                                # CDate エラー 13
            lot.is_over_12_hours("終了", now)


# ------------------------------------------------------------------
class ArrboxTest(_Base):
    def test_main_normal(self) -> None:
        r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL)
        self.assertEqual(r.file_path, lot.db_path(paths.get("main"), "SIKAHIKINOW"))
        self.assertEqual([x[3] for x in r.value], ["06017610", "06019756", "06017503"])   # 受注番号、取り込み順
        self.assertEqual(r.logs, [])
        self.assertEqual(r.messages, [])

    def test_rebuild_rule_removed(self) -> None:
        # D26 改: 時刻窓(毎時 59:46〜00:45)も PLEASE WAIT.txt も見ない → 第1(MainPath)のまま。ログも出ない
        self.d.please_wait()
        r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_WINDOW)
        self.assertEqual(r.file_path, lot.db_path(paths.get("main"), "SIKAHIKINOW"))
        self.assertEqual(r.logs, [])

    def test_main_missing_then_next_missing(self) -> None:
        self.d.remove(self.d.main, "SIKAHIKINOW")
        r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL)
        self.assertEqual(r.file_path, lot.db_path(paths.get("next"), "SIKAHIKINOW"))
        self.assertIsNotNone(r.value)
        self.d.remove(self.d.next, "SIKAHIKINOW")
        r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL)
        self.assertIsNone(r.value)
        self.assertEqual([m.text for m in r.messages], [paths.missing_text("SIKAHIKINOW")])
        r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "", now=T_NORMAL, over12=True)
        self.assertEqual(r.messages, [])                                  # IsOver12Hours の間は出さない

    def test_dbflag_uses_next_without_checks(self) -> None:
        r = lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "1", now=T_WINDOW)
        self.assertEqual(r.file_path, lot.db_path(paths.get("next"), "SIKAHIKINOW"))
        self.assertEqual(r.logs, [])                                      # FPath<>MainPath なので確認しない
        self.d.remove(self.d.next, "SIKAHIKINOW")
        with self.assertRaises(datasource.DbMissing):                     # VBA は End
            lot.arrbox("SIKAHIKINOW", "H2500K0", "Record", "1", now=T_NORMAL)

    def test_fields_and_keywords(self) -> None:
        f = lot.arrbox("SIKAHIKINOW", "板", "Fields", "", now=T_NORMAL)
        self.assertEqual(f.value[:4], ["ﾛｯﾄ番号", "引当番号", "検査番号", "受注番号"])
        self.assertIsNone(lot.arrbox("SIKAHIKINOW", "無いロット", "Record", "", now=T_NORMAL).value)
        # "板" は [ｵｰﾀﾞｰ板丈]<>0(Text 列と数の比較)
        rows = datasource.get_records("SIKALOTNOW", "仕掛")
        cols = datasource.get_fields("SIKALOTNOW", "仕掛")
        k = cols.index("ｵｰﾀﾞｰ板丈")
        r = lot.arrbox("SIKALOTNOW", "板", "Record", "", now=T_NORMAL)
        self.assertEqual(len(r.value), sum(1 for x in rows if x[k] != "0"))
        self.assertEqual(lot.arrbox_filter("コイル"), "[ｵｰﾀﾞｰ板丈]=0  AND  [BOX設計_設備名] = 'KEN' ")
        self.assertEqual(lot.arrbox_filter(""), "")
        self.assertEqual(lot.arrbox_filter("H2500K0"), "[ﾛｯﾄ番号]='H2500K0'")
        with self.assertRaises(ValueError):
            lot.arrbox("SIKAHIKINOW", "", "All")


# ------------------------------------------------------------------
class AccReflecTest(_Base):
    def reflect(self, lot_no: str, **kw) -> lot.LotReflect:
        return lot.acc_reflec(lot.LotScreen(lot=lot_no, cb_text=lot_no, **kw), check_equipment=_ok)

    def test_shiji_full_width_text(self) -> None:
        r = self.reflect("N9170B0")
        self.assertEqual(r.shiji_taik, "ﾓｸﾋﾖｳ3MM")
        self.assertEqual(r.taik_text, "3.0")

    def test_zero_tolerance_names_missing_columns(self) -> None:
        # 2026-11-09: 読んだ台帳に公差の列が無いと公差 0(下限 = 上限)になる → 無い列を知らせる
        f = self.d.main / kensa_fixture.file_name("SIKALOTNOW")
        with sqlite3.connect(f) as c:
            for k in ("ｵｰﾀﾞｰ_ﾌﾟﾗｽ", "ｵｰﾀﾞｰ_ﾏｲﾅｽ", "製造_ﾌﾟﾗｽ", "製造_ﾏｲﾅｽ"):
                c.execute(f'ALTER TABLE "仕掛" DROP COLUMN "板厚公差_{k}"')
        r = self.reflect("N9170B0")
        texts = " ".join(m.text for m in r.messages)
        self.assertIn("板厚の公差が 0 になります", texts)
        self.assertIn("列 板厚公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ", texts)

    def test_offline_fills_missing_columns_from_first(self) -> None:
        # 2026-11-09: 第2(控え)に公差の列が無くても、第1の同じロットで補う(オフラインでも第1に届けば)
        f = self.d.next / kensa_fixture.file_name("SIKALOTNOW")
        with sqlite3.connect(f) as c:
            for k in ("ｵｰﾀﾞｰ_ﾌﾟﾗｽ", "ｵｰﾀﾞｰ_ﾏｲﾅｽ", "製造_ﾌﾟﾗｽ", "製造_ﾏｲﾅｽ"):
                c.execute(f'ALTER TABLE "仕掛" DROP COLUMN "板厚公差_{k}"')
            c.execute('ALTER TABLE "仕掛" DROP COLUMN "品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ"')
        online = self.reflect("N9170B0")
        paths.set_mode(paths.OFFLINE)
        try:
            datasource._table_cache.clear()
            r = self.reflect("N9170B0")
            self.assertEqual((r.form["Ap"], r.form["Am"]), (online.form["Ap"], online.form["Am"]))
            self.assertNotIn("公差が 0", " ".join(m.text for m in r.messages))
            fields = lot.arrbox("SIKALOTNOW", "", "Fields", "").value
            rec = lot.arrbox("SIKALOTNOW", "N9170B0", "Record", "").value[0]
            self.assertNotEqual(rec[fields.index("品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ")], "")
        finally:
            paths.set_mode(paths.ONLINE)

    def test_full_width_tolerance(self) -> None:
        f = self.d.main / kensa_fixture.file_name("SIKALOTNOW")
        with sqlite3.connect(f) as c:
            c.execute('UPDATE "仕掛" SET "板厚公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ" = ?, "板厚公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ" = ? WHERE "ﾛｯﾄ番号" = ?',
                      ("０．２", "０．１", "N9170B0"))
        r = self.reflect("N9170B0")
        self.assertEqual((r.form["Ap"], r.form["Am"]), ("0.200", "0.100"))

    def test_lot_case_insensitive(self) -> None:
        # Access の Text の = は大文字小文字を区別しない
        r = lot.acc_reflec(lot.LotScreen(lot="x", cb_text="h2500k0"), check_equipment=_ok)
        self.assertEqual(r.row["ﾛｯﾄ番号"], "H2500K0")

    def test_not_found_and_order_length_zero(self) -> None:
        # L8410N0 は ｵｰﾀﾞｰ板丈 が '0' の行だけ (SQL の <> '0')
        r = self.reflect("L8410N0")
        self.assertEqual(r.status, "not_found")
        self.assertEqual(set(r.form), set(lot.CLEAR_ON_NOT_FOUND))
        self.assertTrue(all(v == "" for v in r.form.values()))
        self.assertEqual(self.reflect("ZZZZZZZ").status, "not_found")
        self.assertEqual(lot.acc_reflec(lot.LotScreen(lot=""), check_equipment=_ok).status, "no_lot")

    def test_all_108_lots(self) -> None:
        # 実データ: 見つかるのは ﾛｯﾄ番号 が一致し ｵｰﾀﾞｰ板丈 が Null でも '0' でもない行があるロット。採用は先頭行
        with sqlite3.connect(self.d.main / "SIKALOT.sqlite3") as c:
            rows = c.execute('SELECT rowid, "ﾛｯﾄ番号", "ｵｰﾀﾞｰ板丈", "製造板厚", "BOX実績_板厚", "設計_設備ｺｰｽ" '
                             'FROM "仕掛" ORDER BY rowid').fetchall()
        lots = sorted({r[1] for r in rows})
        self.assertEqual(len(lots), 108)
        found = 0
        for L in lots:
            want = [r for r in rows if r[1] == L and r[2] is not None and r[2] != "0"]
            r = self.reflect(L)
            with self.subTest(lot=L):
                if want:
                    found += 1
                    self.assertEqual(r.status, "ok")
                    g = any(k in want[0][5] for k in ("GFS", "GCT", "GSS"))      # 外注は BOX実績
                    self.assertEqual(r.g_flag, g)
                    self.assertEqual(r.form["ATU"], vbx.fmt(vbx.val(want[0][4 if g else 3] or ""), "0.000"))
                    self.assertEqual(r.past_hiki[1], L)
                else:
                    self.assertEqual(r.status, "not_found")
        self.assertEqual(found, 63)

    def test_path_rules(self) -> None:
        # SaveLoad:101-112 DBflag > MainPath(D26 改: PLEASE WAIT は見ない)
        self.assertEqual(lot.acc_reflec_path(""), paths.get("main"))
        self.assertEqual(lot.acc_reflec_path("1"), paths.get("next"))
        self.d.please_wait()
        self.assertEqual(lot.acc_reflec_path(""), paths.get("main"))

    def test_db_missing(self) -> None:
        self.d.remove(self.d.main, "SIKALOTNOW")
        self.assertEqual(self.reflect("H2500K0").status, "ok")      # 2026-11-09: 第1に無ければ第2の台帳を読む
        self.d.remove(self.d.next, "SIKALOTNOW")
        r = self.reflect("H2500K0")
        self.assertEqual(r.status, "db_missing")
        self.assertEqual(r.messages[0].text, "DBファイルが見つかりません:" + "\r\n" + r.db_path)

    def test_equipment_ng_stops(self) -> None:
        seen = []

        def ng(*a):
            seen.append(a)
            return "厚さ範囲外"
        with self.assertRaises(flow.Stop) as cm:
            lot.acc_reflec(lot.LotScreen(lot="H2500K0", cb_text="H2500K0", facility="L-1"), check_equipment=ng)
        self.assertEqual(cm.exception.text, "設備仕様から外れています。,厚さ範囲外,終了します。")
        self.assertEqual(seen[0], ("L-1", "N17S", 50.0, 1250.0, 2500.0, vbx.single(435.505)))

    def test_history_inside_and_error_handler(self) -> None:
        hdr = [list(lot.HISTORY_HEADER)]
        r = lot.acc_reflec(lot.LotScreen(lot="H2500K0", cb_text="H2500K0"), check_equipment=_ok,
                           history_rows=hdr, history_env=lot.HistoryEnv(now=T_NORMAL))
        self.assertEqual(len(r.history.new_rows), 3)
        r = lot.acc_reflec(lot.LotScreen(lot="H2500K0", cb_text="H2500K0", work_time="T"), check_equipment=_ok,
                           history_rows=hdr)
        self.assertIsNone(r.history)
        # 空の Lot情報履歴(A1 の 1 セル)は型エラー 13 → ErrHandler (SaveLoad:572)
        r = lot.acc_reflec(lot.LotScreen(lot="H2500K0", cb_text="H2500K0"), check_equipment=_ok,
                           history_rows=[[None]], history_env=lot.HistoryEnv(now=T_NORMAL))
        self.assertEqual(r.status, "error")
        self.assertEqual(r.messages[-1].text, "DB読み込みエラー:\r\n型が一致しません。")

    def test_quote_in_lot_and_missing_odr(self) -> None:
        # SQL に値をそのまま埋めるので ' があると構文エラー → ErrHandler (SaveLoad:120, 564-572)
        r = self.reflect("A'B")
        self.assertEqual((r.status, r.messages[0].text), ("error", "DB読み込みエラー:\r\nクエリ式の構文エラー"))
        # HistorySave の SIKAODRNOW が無い → VBA は End(DbMissing のまま上へ)
        self.d.remove(self.d.main, "SIKAODRNOW")
        with self.assertRaises(datasource.DbMissing):
            lot.acc_reflec(lot.LotScreen(lot="H2500K0", cb_text="H2500K0"), check_equipment=_ok,
                           history_rows=_sheet(), history_env=lot.HistoryEnv(now=T_NORMAL))

    def test_dedupe_and_materials_order(self) -> None:
        self.assertEqual(lot.dedupe_message("a,b,a,c,b"), "a,b,c")
        self.assertEqual(lot.dedupe_message(""), "")
        no, msgs, _ = lot.materials_order_no("H2500K0", "'06017503", now=T_NORMAL)
        self.assertEqual(no, "06017610")                           # 引当の先頭行の受注番号 (SaveLoad:445)
        no, msgs, _ = lot.materials_order_no("L726M50", "'35000001", now=T_NORMAL)
        self.assertEqual(no, "35000001")                           # 引当なし → PastHiki(0) から ' を除く

    def test_small_helpers(self) -> None:
        self.assertEqual(lot.convert_to_half("２．５ＭＭ"), "2.5")    # 全角の . も半角に(2026-11-07 に直した。VBA は "25")
        self.assertEqual(lot.convert_to_half("1.0/1M"), "1.0")
        self.assertEqual(lot.find_right_string("HOT ANF CAI PSW"), "CAI")
        self.assertEqual(lot.find_right_string(""), "")


# ------------------------------------------------------------------
def _sheet(*rows: list) -> list[list]:
    return [list(lot.HISTORY_HEADER)] + [list(r) for r in rows]


def _hrow(lot_no: str, order: str = "1") -> list:
    r = [None] * 26
    r[1], r[2] = order, lot_no
    return r


class HistorySaveTest(_Base):
    def ph(self, lot_no: str = "H2500K0") -> list:
        r = lot.acc_reflec(lot.LotScreen(lot=lot_no, cb_text=lot_no, work_time="T"), check_equipment=_ok)
        return r.past_hiki

    def test_b15_lateral_grade(self) -> None:
        ph = self.ph()
        self.assertEqual(ph[17], "3")
        res = lot.history_save(ph, _sheet(), lot.HistoryEnv(now=T_NORMAL))
        self.assertEqual(res.new_rows[0][18], "V325NYE")              # S 列はラテラルボー等級を上書き
        res = lot.history_save(ph, _sheet(), lot.HistoryEnv(now=T_NORMAL))
        self.assertEqual(res.new_rows[0][18], "V325NYE")
        self.assertEqual(res.new_rows[0][26], "3")                    # 直した場合は AA 列に残す

    def test_window_is_ignored(self) -> None:
        # D26 改: 時刻窓は見ない。DBflag で控えを選んでいれば控えを読み、MainPath なら MainPath のまま
        res = lot.history_save(self.ph(), _sheet(), lot.HistoryEnv(dbflag="1", now=T_WINDOW))
        self.assertEqual((res.messages, res.logs), ([], []))
        self.assertEqual(res.file_path, lot.db_path(paths.get("next"), "SIKAHIKINOW"))
        self.assertEqual(len(res.new_rows), 3)
        res = lot.history_save(self.ph(), _sheet(), lot.HistoryEnv(now=T_WINDOW))
        self.assertEqual(res.file_path, lot.db_path(paths.get("main"), "SIKAHIKINOW"))
        self.assertEqual(len(res.new_rows), 3)

    def test_special_sheets_use_main(self) -> None:
        # D26 改: PLEASE WAIT は見ない。全ヵ所/キャスク表示中は DBflag より MainPath が先 (SaveLoad:1303-1308)
        self.d.please_wait()
        res = lot.history_save(self.ph(), _sheet(), lot.HistoryEnv(now=T_NORMAL, cask_visible=True))
        self.assertEqual(res.messages, [])
        res = lot.history_save(self.ph(), _sheet(), lot.HistoryEnv(dbflag="1", now=T_NORMAL, zenkasho_visible=True))
        self.assertEqual(res.file_path, lot.db_path(paths.get("main"), "SIKAHIKINOW"))

    def test_no_hiki(self) -> None:
        ph = self.ph("N8096B0")
        self.d.delete_lot(self.d.main, "SIKAHIKINOW", "N8096B0")
        # MainPath に無い → NextPath で再検索 (SaveLoad:1343-1352)
        res = lot.history_save(ph, _sheet(), lot.HistoryEnv(now=T_NORMAL))
        self.assertEqual(res.file_path, lot.db_path(paths.get("next"), "SIKAHIKINOW"))
        self.assertTrue(res.completed)
        self.d.delete_lot(self.d.next, "SIKAHIKINOW", "N8096B0")
        res = lot.history_save(ph, _sheet(), lot.HistoryEnv(now=T_NORMAL))
        self.assertEqual([m.text for m in res.messages], ["LOT一覧：LOTNO有_HIKIDB：LOTNO無_HistorySave"])
        res = lot.history_save(ph, _sheet(), lot.HistoryEnv(dbflag="1", now=T_NORMAL))
        self.assertEqual([m.text for m in res.messages], ["SIKAHIKINOW.accdbにN8096B0ﾃﾞｰﾀがありません_HistorySave"])
        # Test のときは黙って抜ける (SaveLoad:1337-1339)
        for env in (lot.HistoryEnv(now=T_NORMAL, uf1_test=True),
                    lot.HistoryEnv(now=T_NORMAL, cask_visible=True, cask_line="Test"),
                    lot.HistoryEnv(now=T_NORMAL, zenkasho_visible=True, mitsubishi_line="Test")):
            res = lot.history_save(ph, _sheet(), env)
            self.assertEqual(res.messages, [])
            self.assertFalse(res.completed)

    def test_offline_same_paths(self) -> None:
        # オフラインで MainPath = NextPath のときは `Case FPath = NextPath` が先に当たる
        kensa_fixture.use_offline(kensa_fixture.copy_to(self.d.root / "same"))
        res = lot.history_save(self.ph(), _sheet(), lot.HistoryEnv(now=T_WINDOW))   # D26 改: 時刻窓は見ない
        self.assertEqual(len(res.new_rows), 3)
        res = lot.history_save(self.ph(), _sheet(), lot.HistoryEnv(now=T_NORMAL))
        self.assertEqual(len(res.new_rows), 3)

# ------------------------------------------------------------------
class LotInfoTest(_Base):
    def test_get_lot_info(self) -> None:
        r1 = ["2026/10/30", "1", "L1", "3", "1000", "2000", "K1", "5052", "H34", "U1", "C", "得意", "納入", "送り"]
        r2 = list(r1)
        r2[2] = "L2"
        sheet = [list(lot.HISTORY_HEADER), r1, r2, list(r1)]
        got = lot.get_lot_info(sheet, "L1")
        self.assertEqual(got, [["納入", "得意", "U1", "5052", "H34", "3", "1000", "2000", "L1", "K1"]] * 2)
        self.assertIsNone(lot.get_lot_info(sheet, "Z"))
        self.assertIsNone(lot.get_lot_info(None, "L1"))
        self.assertIsNone(lot.get_lot_info([["ﾛｯﾄNo", "厚"], ["L1", "3"]], "L1"))   # 見出しが足りない

    def test_sika_atu_create(self) -> None:
        self.assertEqual(lot.sika_atu_create("H2500K0"), ["50", "50"])
        self.assertIsNone(lot.sika_atu_create(""))
        self.d.delete_lot(self.d.main, "SIKALOTNOW", "H2500K0")
        self.assertEqual(lot.sika_atu_create("H2500K0"), ["50", "50"])      # NextPath (SaveLoad:40-45)
        self.d.delete_lot(self.d.next, "SIKALOTNOW", "H2500K0")
        self.assertIsNone(lot.sika_atu_create("H2500K0"))
        self.d.remove(self.d.main, "SIKALOTNOW")
        with self.assertRaises(datasource.DbMissing):                        # VBA は End
            lot.sika_atu_create("H2500K0")


class VcTest(_Base):
    def sheet(self, vc="1", a="AM-500", b="BE-702H", order="111") -> list:
        r = _hrow("L1", order)
        r[17], r[18], r[19] = vc, a, b
        return _sheet(_hrow("L0", "000"), r)

    def pic(self, name: str) -> None:
        d = self.d.pack / "VC種類画像"
        d.mkdir(exist_ok=True)
        (d / name).write_bytes(b"jpg")

    def test_both_pictures(self) -> None:
        self.pic("AM-500.JPG")
        self.pic("BE-702H.JPG")
        v = lot.vc_picture(self.sheet(), "L1")
        self.assertTrue(v.matched and v.tall_form and v.l_a_tag_on and v.l_b_tag_on)
        self.assertEqual((v.l_a_caption, v.l_b_caption), ("A面:ﾄｸｼﾕﾌｲﾙﾑ_ﾄｳﾒｲ", "B面:ﾎﾟﾘｴﾁﾚﾝ_ﾌﾞﾙ-"))
        self.assertTrue(v.a_picture.endswith("AM-500.JPG"))
        self.assertTrue(v.a_visible and v.b_visible and not v.nothing_visible)
        self.assertEqual(v.b_left, "A幅+30")

    def test_missing_picture(self) -> None:
        self.pic("AM-500.JPG")
        v = lot.vc_picture(self.sheet(), "L1")
        self.assertTrue(v.nothing_visible)
        self.assertEqual(v.nothing_text, "BE-702H.JPG" + lot.VC_NOTHING_SUFFIX)
        self.assertFalse(v.a_visible or v.b_visible or v.l_a_visible or v.l_b_visible)

    def test_only_b_and_no_vc(self) -> None:
        self.pic("BE-702H.JPG")
        v = lot.vc_picture(self.sheet(a=None), "L1")
        self.assertTrue(v.b_visible and v.l_b_visible and not v.a_visible)
        self.assertEqual(v.b_left, "B幅+30")
        self.assertEqual(v.name_a, "")
        v = lot.vc_picture(self.sheet(vc="0"), "L1")
        self.assertEqual(v.nothing_text, "111:受注番号一致なし(進度ﾌｧｲﾙ異常)")
        self.assertFalse(v.tall_form)
        self.assertFalse(lot.vc_picture(self.sheet(), "無い").matched)

    def test_order_no_and_materials(self) -> None:
        self.pic("AM-500.JPG")
        self.pic("BE-702H.JPG")
        calls = []

        def mup(order: str, flag: bool) -> str:
            calls.append((order, flag))
            return "当板あり" if flag else "アングル"
        v = lot.vc_picture(self.sheet(vc="0"), "L1", "111", materials_upper_part=mup)
        self.assertTrue(v.a_visible)                               # 受注番号指定は VC 列を見ない
        self.assertEqual(calls, [("111", False), ("111", True)])
        self.assertEqual([m.text for m in v.messages], ["アングル", "当板あり"])

    def test_vc_conv(self) -> None:
        r = lot.vc_conv("AM-500")
        self.assertEqual(r.value, vbx.single(0.039))
        r = lot.vc_conv("無い")
        self.assertEqual(r.value, vbx.single(0.08406977))
        self.assertEqual(r.messages[0].text, "VC厚登録なし,担当者に連絡のこと")
        r = lot.vc_conv("AM-500", test=True)                         # Test は PackingPath(ファイルなし)
        self.assertEqual((r.value, r.messages[0].text), (0.0, paths.missing_text("伝送用ファイル", [paths.db_file("伝送用ファイル", base_key="packing")])))


class VerifierAccReflecNullTest(_Base):
    def test_ken_no_null_with_lateral_instruction(self) -> None:
        # 指示_ﾗﾃﾗﾙﾎﾞｰ があるときだけ `Dim kenNoTemp As String: kenNoTemp = PastHiki(5)`(SaveLoad:301)を通る。
        # 検査番号が Null ならそこでエラー 94 → null_uses に残す(L8430D0 は指示 "2.0MM/1M" あり)
        from unittest import mock
        orig = lot._select_lot

        def null_ken(file_path: str, cb_text: str):
            cols, values = orig(file_path, cb_text)
            v = list(values)
            v[cols.index("検査番号")] = None
            return cols, tuple(v)

        with mock.patch.object(lot, "_select_lot", null_ken):
            r = lot.acc_reflec(lot.LotScreen(lot="L8430D0", cb_text="L8430D0"), check_equipment=_ok)
        self.assertEqual(r.null_uses, ["SaveLoad:301 検査番号"])
        self.assertEqual(r.ken_no, "")


if __name__ == "__main__":
    unittest.main()


# ------------------------------------------------------------------
# 直した場合の分岐(B49・B50・B51・B54)。既定(再現)の側は上の試験で確かめている
# ------------------------------------------------------------------
class FixBranchTest(_Base):
    def ph(self, lot_no: str = "H2500K0") -> list:
        r = lot.acc_reflec(lot.LotScreen(lot=lot_no, cb_text=lot_no, work_time="T"), check_equipment=_ok)
        return r.past_hiki

    def test_b49_next_selected_missing(self) -> None:
        # D26 改で時刻窓・PLEASE WAIT を外したので、控えを選んで ReadCheck に落ちるのは控えのファイルが無いときだけ。
        # そのときは再現も直すも「ありません」で抜ける(B49 の差は出なくなった)
        ph = self.ph()
        self.d.remove(self.d.next, "SIKAHIKINOW")
        for mode in ("reproduce", "fix"):
            res = lot.history_save(ph, _sheet(), lot.HistoryEnv(dbflag="1", now=T_WINDOW))
            self.assertEqual([m.text for m in res.messages], [paths.missing_text("SIKAHIKINOW")])
            self.assertFalse(res.completed)

    def test_b49_offline_same_paths(self) -> None:
        # オフライン(MainPath = NextPath)で時刻窓: D26 改で時刻窓を見ないので、どちらも同じファイルを読む
        kensa_fixture.use_offline(kensa_fixture.copy_to(self.d.root / "same"))
        ph = self.ph()
        self.assertEqual(len(lot.history_save(ph, _sheet(), lot.HistoryEnv(now=T_WINDOW)).new_rows), 3)
        self.assertEqual(len(lot.history_save(ph, _sheet(), lot.HistoryEnv(now=T_WINDOW)).new_rows), 3)

    def test_b50_delete_only_after_db_checks(self) -> None:
        ph = self.ph()
        sheet = _sheet(_hrow("H2500K0"))
        # 最後まで進めば消す
        res = lot.history_save(ph, sheet, lot.HistoryEnv(now=T_NORMAL))
        self.assertTrue(res.completed)
        self.assertEqual(res.delete_rows, [2])
        # 途中で抜けたら消さない(再現は [2]。test_delete_happens_before_db_checks)
        self.d.remove(self.d.main, "SIKAHIKINOW")
        self.d.remove(self.d.next, "SIKAHIKINOW")
        res = lot.history_save(ph, sheet, lot.HistoryEnv(now=T_NORMAL))
        self.assertFalse(res.completed)
        self.assertEqual(res.delete_rows, [])
        # Test で黙って抜けるときも消さない
        res = lot.history_save(ph, sheet, lot.HistoryEnv(now=T_NORMAL, uf1_test=True))
        self.assertEqual(res.delete_rows, [])

    def test_b51_odr_fallback(self) -> None:
        ph = self.ph()
        with sqlite3.connect(self.d.main / "SIKAODR.sqlite3") as c:
            c.execute('DELETE FROM "仕掛"')
        # 直す: MainPath の SIKAODRNOW が空 → NextPath へ切り替えて続ける (SaveLoad:1391-1401)
        res = lot.history_save(ph, _sheet(), lot.HistoryEnv(now=T_NORMAL))
        self.assertTrue(res.completed)
        self.assertEqual(len(res.new_rows), 3)
        # 控えを選んでいたら「…がありません」で抜ける (SaveLoad:1393)
        with sqlite3.connect(self.d.next / "SIKAODR.sqlite3") as c:
            c.execute('DELETE FROM "仕掛"')
        res = lot.history_save(ph, _sheet(), lot.HistoryEnv(dbflag="1", now=T_NORMAL))
        self.assertEqual([m.text for m in res.messages], ["SIKAODRNOW.accdbにVCFieldﾃﾞｰﾀがありません_HistorySave"])
        self.assertFalse(res.completed)
        # MainPath → NextPath も空 → 「VCFieldData無」
        res = lot.history_save(ph, _sheet(), lot.HistoryEnv(now=T_NORMAL))
        self.assertEqual([m.text for m in res.messages], ["VCFieldData無_HistorySave"])
        res = lot.history_save(ph, _sheet(), lot.HistoryEnv(now=T_NORMAL, over12=True))
        self.assertEqual(res.messages, [])
        # NextPath のファイルが無い → 「日稲:…がありません」
        self.d.remove(self.d.next, "SIKAODRNOW")
        res = lot.history_save(ph, _sheet(), lot.HistoryEnv(now=T_NORMAL))
        self.assertEqual([m.text for m in res.messages], [paths.missing_text("SIKAODRNOW")])


class HikiMergeTest(_Base):
    """2026-11-08: 第1の SIKAHIKI に引当が足りないとき、第2にだけある引当番号を足す。"""

    def _extra_in_second(self) -> tuple[str, str]:
        f = self.d.next / kensa_fixture.file_name("SIKAHIKINOW")
        with sqlite3.connect(f) as c:
            cols = [r[1] for r in c.execute('PRAGMA table_info("仕掛")')]
            row = list(c.execute('SELECT * FROM "仕掛" LIMIT 1').fetchone())
            lot_no = row[cols.index("ﾛｯﾄ番号")]
            row[cols.index("引当番号")] = "99999999"
            c.execute(f'INSERT INTO "仕掛" VALUES ({",".join("?" * len(row))})', row)
        return lot_no, "99999999"

    def _nos(self, rows: list) -> set:
        cols = datasource.get_fields(lot.db_path(paths.get("main"), "SIKAHIKINOW"), lot.TABLE, "")
        i = datasource.column_index(cols, "引当番号")
        return {vbx.cstr(r[i]) for r in rows}

    def test_arrbox_adds_missing_hiki(self) -> None:
        lot_no, extra = self._extra_in_second()
        r = lot.arrbox("SIKAHIKINOW", lot_no, "Record", "", now=T_NORMAL)
        self.assertIn(extra, self._nos(r.value))
        self.assertTrue(any("第2パスから足しました" in m.text for m in r.messages))

    def test_nothing_added_when_same(self) -> None:
        f = self.d.next / kensa_fixture.file_name("SIKAHIKINOW")
        with sqlite3.connect(f) as c:
            lot_no = c.execute('SELECT "ﾛｯﾄ番号" FROM "仕掛" LIMIT 1').fetchone()[0]
        r = lot.arrbox("SIKAHIKINOW", lot_no, "Record", "", now=T_NORMAL)
        self.assertEqual(r.messages, [])

    def test_offline_reads_second_only(self) -> None:
        lot_no, extra = self._extra_in_second()
        paths.set_mode(paths.OFFLINE)
        try:
            r = lot.arrbox("SIKAHIKINOW", lot_no, "Record", "", now=T_NORMAL)
            self.assertIn(extra, self._nos(r.value))       # 第2だけを読む(足す処理はしない)
            self.assertEqual(r.messages, [])
        finally:
            paths.set_mode(paths.ONLINE)

    def test_watch_sees_added_hiki(self) -> None:
        from ufkensa.kensa import lot_watch
        f = self.d.next / kensa_fixture.file_name("SIKAHIKINOW")
        with sqlite3.connect(f) as c:
            lot_no = c.execute('SELECT "ﾛｯﾄ番号" FROM "仕掛" LIMIT 1').fetchone()[0]
        before = lot_watch.signature(lot_no)
        self._extra_in_second()
        self.assertIn("引当", before.changed(lot_watch.signature(lot_no)))
