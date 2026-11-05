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
        quirks.set_for_test(None)
        self.d = _Dirs()

    def tearDown(self) -> None:
        quirks.set_for_test(None)
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
        quirks.set_for_test(None)

    def test_reproduce_b19_one_hour_boundary(self) -> None:
        # Worker:342 `DateDiff("h", C1, Now) >= 1` は時の境界の数
        self.assertFalse(lot.is_over_12_hours("2026/10/30 10:00:00", dt.datetime(2026, 10, 30, 10, 59, 59)))
        self.assertTrue(lot.is_over_12_hours("2026/10/30 10:59:59", dt.datetime(2026, 10, 30, 11, 0, 0)))
        self.assertFalse(lot.is_over_12_hours("2026/10/30 12:00:00", dt.datetime(2026, 10, 30, 11, 0, 0)))

    def test_fix_b19_twelve_hours(self) -> None:
        quirks.set_for_test({"B19": "fix"})
        # 10 時台 → 21 時台は境界 11 個、22 時台で 12 個
        self.assertFalse(lot.is_over_12_hours("2026/10/30 10:59:59", dt.datetime(2026, 10, 30, 21, 59, 59)))
        self.assertTrue(lot.is_over_12_hours("2026/10/30 10:59:59", dt.datetime(2026, 10, 30, 22, 0, 0)))

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

    def test_gaichu_and_tolerances_L8430D0(self) -> None:
        r = self.reflect("L8430D0")
        self.assertEqual(r.status, "ok")
        self.assertEqual(r.row["ﾛｯﾄ番号"], "L8430D0")
        # 設計_設備ｺｰｽ に GFS (SaveLoad:139-149)
        self.assertTrue(r.g_flag)
        self.assertEqual(r.temp_e, "GFSｺｰｽ有、進度にあった製品ｻｲｽﾞ入力をお願いします。")
        # 外注は BOX実績 (SaveLoad:158-162, 186-190, 221-225)
        self.assertEqual(r.equipment_args[1:], ("3004", vbx.single(0.3), 1185.0, 0.0, vbx.single(0.09763)))
        self.assertEqual(r.past_hiki[0], "'35008168")
        self.assertEqual(r.past_hiki[2:5], ["0.3", "1185", "0"])
        f = r.form
        self.assertEqual((f["ATU"], f["Hab"], f["Tak"]), ("0.300", "1185.0", "0.0"))
        # 厚はオーダー公差(どちらも 0 でない)、幅・丈はオーダーのマイナスが 0 → 製造公差 (SaveLoad:231/240/249)
        self.assertEqual((f["Ap"], f["Am"], f["Hp"], f["Hm"], f["Tp"], f["Tm"]),
                         ("0.040", "0.040", "0.5", "0.5", "1.0", "0.5"))
        self.assertEqual((f["ApK"], f["AmK"]), ("0.340", "0.260"))
        self.assertEqual((f["HpK"], f["HmK"], f["TpK"], f["TmK"]), ("", "", "", ""))   # SaveLoad:264-266
        self.assertEqual((f["TAIK"], f["Lateral"], f["flatnes"]), ("3", "", "3"))      # "." は数でない → ""
        self.assertEqual(f["KenH"], "前工程実績1枚")
        self.assertEqual(r.last_hot, "最終熱処理工程：ANF")
        self.assertIsNone(r.qf)
        # 指示 "4.5MM" → 4.5、"2.0MM/1M" → "/" の前 → 2.0 (補助 ConvertToHalf/numExtract)
        self.assertEqual((r.taik_text, r.lateral_text), ("4.5", "2.0"))
        self.assertEqual(r.additional, "対角ﾗﾃﾗﾙﾎﾞ-追記公差あり")
        self.assertEqual(r.ken_no, "N6X1262")
        self.assertEqual(r.vc_action, "clear")                     # VC "00" = 0
        self.assertEqual(r.sika_set_atu, vbx.single(0.3))
        self.assertEqual(r.cache_c_course, "HOT L-1 L-1 L-1 L-1 DL2 L-2 L-2 L-2 ANF")
        self.assertEqual(r.null_uses, [])
        self.assertTrue(r.run_history_save)

    def test_order_minus_zero_and_taik_dot_N8096B0(self) -> None:
        r = self.reflect("N8096B0")
        f = r.form
        self.assertFalse(r.g_flag)
        self.assertEqual(f["TAIK"], "4")                           # "." → "4" (SaveLoad:214)
        self.assertEqual((f["ATU"], f["Ap"], f["Am"], f["ApK"], f["AmK"]), ("6.620", "0.080", "0.120", "6.700", "6.500"))
        self.assertEqual((f["Hab"], f["Hp"], f["Hm"], f["HpK"], f["HmK"]), ("810.0", "5.0", "5.0", "815.0", "805.0"))
        self.assertEqual((f["Tak"], f["Tp"], f["Tm"], f["TpK"], f["TmK"]), ("777.5", "2.5", "2.5", "780.0", "775.0"))
        self.assertIsNone(r.last_hot)
        self.assertEqual((r.taik_text, r.lateral_text, r.additional), (None, None, ""))
        # 仕掛の列はすべて必須(NOT NULL)なので空欄は "" で入っている。Null を読む所は無い
        self.assertEqual(r.null_uses, [])

    def test_vc_picture_and_null_tolerance_H2500K0(self) -> None:
        r = self.reflect("H2500K0")
        self.assertEqual(r.vc_action, "picture")                   # "01" = 1
        self.assertEqual(r.taik_text, "2.0")
        self.assertEqual(r.additional, "対角追記公差あり")
        self.assertIsNone(r.qf)                                    # QF は設計にだけある(実績で見る)
        self.assertEqual(r.form["Ap"], "0.650")
        self.assertEqual(r.null_uses, [])                          # 必須の列なので Null は無い("")

    def test_shiji_full_width_text(self) -> None:
        r = self.reflect("N9170B0")
        self.assertEqual(r.shiji_taik, "ﾓｸﾋﾖｳ3MM")
        self.assertEqual(r.taik_text, "3.0")

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
        self.assertEqual(lot.convert_to_half("２．５ＭＭ"), "25")     # 全角の . は残らない
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

    def test_rows_from_real_data(self) -> None:
        ph = self.ph()
        res = lot.history_save(ph, _sheet(), lot.HistoryEnv(now=T_NORMAL))
        self.assertTrue(res.completed)
        self.assertEqual(res.messages, [])
        # 外側 = SIKAODRNOW の順(06017610, 06017503, 06019756)、内側 = 引当 (SaveLoad:1466-1467)
        self.assertEqual([r[21] for r in res.new_rows], ["61026069", "61026068", "61026070"])
        first = res.new_rows[0]
        self.assertEqual(len(first), 26)
        self.assertEqual(first[0], T_NORMAL)
        self.assertEqual(first[1:11], ["06017610", "H2500K0", "50", "1252", "2502.5", "H6X1690", "N17S", "T351",
                                       "H197", "HOT PSW ST2 QF ST2 PSW VC KEN"])
        self.assertEqual(first[11:14], ["ﾆﾂﾃﾂﾌﾞﾂｻﾝ(ｶ)ｷﾉｳﾏﾃﾘｱﾙﾌﾞ", "ｶ)ｼﾝｸｽｺ-ﾎﾟﾚ-ｼﾖﾝ ｶﾝｻｲｺｳｼﾞﾖｳ",
                                        "ｶ)ﾋﾖｳｺｳ ﾛﾂｺｳﾌﾞﾂﾘﾕｳｾﾝﾀ-"])
        self.assertEqual(first[14:], ["8", "4288.8", "0", "01", "V325NYE", "V325NYE", "435.50507", "61026069",
                                      "0", "4", "2026-10-26 08:53:25", "1P1125"])
        self.assertEqual(res.cache, {"lot_no": "H2500K0", "cours": "HOT PSW ST2 QF ST2 PSW VC KEN", "ken_no": "H"})

    def test_b15_lateral_grade(self) -> None:
        ph = self.ph()
        self.assertEqual(ph[17], "3")
        res = lot.history_save(ph, _sheet(), lot.HistoryEnv(now=T_NORMAL))
        self.assertEqual(res.new_rows[0][18], "V325NYE")              # S 列はラテラルボー等級を上書き
        quirks.set_for_test({"B15": "fix"})
        res = lot.history_save(ph, _sheet(), lot.HistoryEnv(now=T_NORMAL))
        self.assertEqual(res.new_rows[0][18], "V325NYE")
        self.assertEqual(res.new_rows[0][26], "3")                    # 直した場合は AA 列に残す

    def test_b16_delete_range(self) -> None:
        sheet = _sheet(_hrow("A"), _hrow("H2500K0"), _hrow("B"), _hrow("H2500K0"), _hrow("C"))
        # シートの行番号: 2=A 3=H 4=B 5=H 6=C
        self.assertEqual(lot.history_delete_rows(sheet, "H2500K0"), [3, 4, 5])   # 間の B も消える
        quirks.set_for_test({"B16": "fix"})
        self.assertEqual(lot.history_delete_rows(sheet, "H2500K0"), [3, 5])
        self.assertEqual(lot.history_delete_rows(sheet, "Z"), [])
        self.assertEqual(lot.history_delete_rows(_sheet(), "Z"), [])
        with self.assertRaises(vbx.TypeMismatch):                     # 空のシート(A1 だけ)
            lot.history_delete_rows([[None]], "Z")
        # 数のセルとロットの文字列は一致しない(Variant の比較)
        self.assertEqual(lot.history_delete_rows(_sheet(_hrow(1234567)), "1234567"), [])

    def test_delete_happens_before_db_checks(self) -> None:
        self.d.remove(self.d.main, "SIKAHIKINOW")
        self.d.remove(self.d.next, "SIKAHIKINOW")
        res = lot.history_save(self.ph(), _sheet(_hrow("H2500K0")), lot.HistoryEnv(now=T_NORMAL))
        self.assertEqual(res.delete_rows, [2])
        self.assertEqual([m.text for m in res.messages], [paths.missing_text("SIKAHIKINOW")])
        self.assertFalse(res.completed)
        res = lot.history_save(self.ph(), _sheet(), lot.HistoryEnv(now=T_NORMAL, over12=True))
        self.assertEqual(res.messages, [])

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

    def test_odr_empty_is_type_mismatch(self) -> None:
        with sqlite3.connect(self.d.main / "SIKAODR.sqlite3") as c:
            c.execute('DELETE FROM "仕掛"')
        with self.assertRaises(vbx.TypeMismatch) as cm:
            lot.history_save(self.ph(), _sheet(_hrow("H2500K0")), lot.HistoryEnv(now=T_NORMAL))
        self.assertEqual(cm.exception.partial.delete_rows, [2])


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
            quirks.set_for_test({"B49": mode})
            res = lot.history_save(ph, _sheet(), lot.HistoryEnv(dbflag="1", now=T_WINDOW))
            self.assertEqual([m.text for m in res.messages], [paths.missing_text("SIKAHIKINOW")])
            self.assertFalse(res.completed)

    def test_b49_offline_same_paths(self) -> None:
        # オフライン(MainPath = NextPath)で時刻窓: D26 改で時刻窓を見ないので、どちらも同じファイルを読む
        kensa_fixture.use_offline(kensa_fixture.copy_to(self.d.root / "same"))
        ph = self.ph()
        self.assertEqual(len(lot.history_save(ph, _sheet(), lot.HistoryEnv(now=T_WINDOW)).new_rows), 3)
        quirks.set_for_test({"B49": "fix"})
        self.assertEqual(len(lot.history_save(ph, _sheet(), lot.HistoryEnv(now=T_WINDOW)).new_rows), 3)

    def test_b50_delete_only_after_db_checks(self) -> None:
        ph = self.ph()
        sheet = _sheet(_hrow("H2500K0"))
        quirks.set_for_test({"B50": "fix"})
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

    def test_b50_fix_error_keeps_rows(self) -> None:
        quirks.set_for_test({"B50": "fix"})
        for d in (self.d.main, self.d.next):
            with sqlite3.connect(d / "SIKAODR.sqlite3") as c:
                c.execute('DELETE FROM "仕掛"')
        with self.assertRaises(vbx.TypeMismatch) as cm:          # B51 は再現のまま → 型エラー 13
            lot.history_save(self.ph(), _sheet(_hrow("H2500K0")), lot.HistoryEnv(now=T_NORMAL))
        self.assertEqual(cm.exception.partial.delete_rows, [])

    def test_b51_odr_fallback(self) -> None:
        ph = self.ph()
        with sqlite3.connect(self.d.main / "SIKAODR.sqlite3") as c:
            c.execute('DELETE FROM "仕掛"')
        quirks.set_for_test({"B51": "fix"})
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

    def test_b54_clear_previous_captions(self) -> None:
        # 再現: 熱処理工程・QF が無いロットは None(前ロットの表示を変えない)(SaveLoad:211-213)
        r = lot.acc_reflec(lot.LotScreen(lot="R9022F3", cb_text="R9022F3", work_time="T"), check_equipment=_ok)
        self.assertIsNone(r.last_hot)
        self.assertIsNone(r.qf)
        quirks.set_for_test({"B54": "fix"})
        r = lot.acc_reflec(lot.LotScreen(lot="R9022F3", cb_text="R9022F3", work_time="T"), check_equipment=_ok)
        self.assertEqual((r.last_hot, r.qf), ("", ""))
        r = lot.acc_reflec(lot.LotScreen(lot="RX021T0", cb_text="RX021T0", work_time="T"), check_equipment=_ok)
        self.assertEqual((r.last_hot, r.qf), ("", "QF有り"))
        r = lot.acc_reflec(lot.LotScreen(lot="L8430D0", cb_text="L8430D0", work_time="T"), check_equipment=_ok)
        self.assertEqual((r.last_hot, r.qf), ("最終熱処理工程：ANF", ""))
