"""保守画面 /maint と /api/maint/*(app/routes/maint.py → ufkensa/kensa/maint.py)

歪公差の照合そのものは test_kensa_quality が見ている。ここは Test フォームと同じ 8 項目・書式への組み立て、
別の品質基準との比べ方(このスレッドだけ差し替える)、ソリ公差・設備仕様の表引き、端末設定のセット、
編集の鍵(2026-11-02: 開く・読むは誰でも、編集は 管理者パスワード + 管理者の権限)、仮ロット、管理者パスワードの変更を見る。
基準データ(tests/kensa_fixture.py)のロット: H9022K1(別表B 管理番号 25。巾 2.5 / 丈 3.5)。
"""
from __future__ import annotations

import re
import shutil
import sqlite3
import tempfile
import threading
import unittest
from pathlib import Path
from unittest import mock

from tests import kensa_fixture
from tests._web import make_client
from ufkensa import admin_password, auth, config, user_settings
from ufkensa.kensa import kari_lot, maint, paths, quality, store as st, terminal


def _make_alt(src: Path, dst: Path, no: int = 25, value: str = "9") -> Path:
    """品質基準の写しを作り、平坦度一覧QP追加 の管理番号 no の幅・丈の列を value にする(変更後の表のつもり)。"""
    shutil.copy(src, dst)
    c = sqlite3.connect(dst)
    cols = [x[1] for x in c.execute('pragma table_info("平坦度一覧QP追加")')]
    ht = [x for x in cols if re.fullmatch(r"[HT]\d+(\.\d+)?", x)]
    c.execute('update "平坦度一覧QP追加" set ' + ",".join(f'"{x}"=?' for x in ht) + ' where "管理番号"=?',
              [value] * len(ht) + [no])
    c.commit()
    c.close()
    return dst


class MaintBase(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="ufkensa_maint_"))
        kensa_fixture.use_offline(kensa_fixture.copy_to(self.tmp / "db"))
        terminal.set_facility("HVC")

    def tearDown(self) -> None:
        user_settings.save(terminal.KEY_FACILITY, "")
        user_settings.save(terminal.KEY_LINE_NO, "")
        paths.set_mode(paths.OFFLINE)
        shutil.rmtree(self.tmp, ignore_errors=True)

    def alt(self) -> Path:
        (self.tmp / "変更後").mkdir(exist_ok=True)
        return _make_alt(maint.current_quality_db(), self.tmp / "変更後" / "品質基準.sqlite3")


class DistCheckTest(MaintBase):
    def test_one_lot_like_test_form(self) -> None:
        r = maint.dist_check("h9022k1")                                   # 大文字にして引く
        self.assertEqual((r["lot"], r["ken_no"]), ("H9022K1", "H6X1176"))  # D15: ロットの検番
        cur = r["current"]
        self.assertEqual(cur["status"], "ok")
        v = cur["values"]
        # Test:96-105: 1〜5 列は Format "0.0"(空は空)、6〜8 列はそのまま
        self.assertEqual((v["MG"], v["HaHiz"], v["TaHiz"], v["Naka"], v["Mimi"]), ("", "2.5", "3.5", "", ""))
        self.assertEqual(v["Source"], "Q-S-QA-DIM-122_別表B")
        self.assertEqual(v["NoTex"], "25")
        self.assertEqual(v["HizCome"], "異常発生時廃棄せずに品証又は生管に連絡する。")
        self.assertIsNone(r["alt"])
        self.assertEqual(r["diff"], [])

    def test_same_as_register_lot_parts(self) -> None:
        """通常検査のロット登録と同じ部品(hiki_arr_create → dist_array)の結果と一致する。"""
        from ufkensa.kensa import lot as lotmod, quality_session
        lot = "L8281G0"
        data = quality.LotData(lotmod.arrbox("SIKALOTNOW", "板", "Fields").value,
                               lotmod.arrbox("SIKALOTNOW", lot, "Record").value)
        qp = quality.dist_array(lot, quality_session.ken_no_of(lot, data), quality.hiki_arr_create(lot), data)
        v = maint.dist_check(lot)["current"]["values"]
        self.assertEqual(v["HaHiz"], "4.0")
        self.assertEqual(v["HaHiz"], quality.vbx.fmt(qp.kousa[0][quality.K_HA], "0.0"))
        self.assertEqual(v["Source"], qp.kousa[0][quality.K_SOURCE])

    def test_not_found_and_none(self) -> None:
        r = maint.dist_check("ZZZZZZZ")
        self.assertEqual(r["current"]["status"], "not_found")
        self.assertEqual(maint.dist_check("R6550H0")["current"]["status"], "none")   # 検番 0000… で公差なし

    def test_alt_side_by_side(self) -> None:
        alt = self.alt()
        before = maint.current_quality_db()
        r = maint.dist_check("H9022K1", alt)
        self.assertEqual(r["current"]["values"]["HaHiz"], "2.5")
        self.assertEqual((r["alt"]["values"]["HaHiz"], r["alt"]["values"]["TaHiz"]), ("9.0", "9.0"))
        self.assertEqual(r["diff"], ["HaHiz", "TaHiz"])
        self.assertEqual(quality._quality_db.__name__, "_quality_db")        # 差し替えは元に戻る
        self.assertEqual(quality._quality_db(), before)

    def test_override_is_thread_local(self) -> None:
        alt = self.alt()
        seen: dict[str, Path] = {}
        go, done = threading.Event(), threading.Event()

        def other() -> None:
            go.wait(5)
            seen["other"] = quality._quality_db()
            done.set()
        t = threading.Thread(target=other)
        t.start()
        with maint.quality_db(alt):
            seen["mine"] = quality._quality_db()
            go.set()
            done.wait(5)
        t.join(5)
        self.assertEqual(seen["mine"], alt)
        self.assertEqual(seen["other"], maint.current_quality_db())      # 通常検査の要求はいまの表のまま

    def test_all_lots(self) -> None:
        out = maint.dist_check_all(self.alt())
        self.assertEqual(out["count"], 63)                                # 板(オーダー板丈<>0)のロット
        self.assertEqual(out["by_status"]["error"], 0)
        diff_lots = [r["lot"] for r in out["rows"] if r["diff"]]
        self.assertIn("H9022K1", diff_lots)
        self.assertEqual(out["diffs"], len(diff_lots))
        self.assertGreater(len(maint.lot_numbers(plate_only=False)), out["count"])
        self.assertTrue(all(x.startswith("H9") for x in maint.lot_numbers(prefix="h9")))

    def test_resolve_alt(self) -> None:
        self.assertIsNone(maint.resolve_alt("  "))
        with self.assertRaises(ValueError):
            maint.resolve_alt(str(self.tmp / "nai.sqlite3"))
        with self.assertRaises(ValueError):
            maint.resolve_alt(str(self.tmp))                              # .sqlite3 でない


class SoriEquipmentTest(MaintBase):
    def test_sori(self) -> None:
        # ソリ4級: 板厚 2.01〜7 の行。幅 1200 → H1500(2)、丈は T1530 まで 2、超えると T1999(3)
        r = maint.sori_check("3", "1200", "1530")
        self.assertEqual((r["text"], r["haba_col"], r["take_col"]), ("幅2.0 丈2.0", "H1500", "T1530"))
        self.assertEqual(maint.sori_check("3", "1200", "1531")["text"], "幅2.0 丈3.0")
        self.assertEqual(maint.sori_check("3", "1200", "99999")["take_col"], "T9999")   # 範囲外は最大列
        self.assertEqual(maint.sori_check("1000", "1200", "1500")["text"], "幅0.0 丈0.0")  # 該当行なし(999 超) → 0
        with self.assertRaises(ValueError):
            maint.sori_check("", "1200", "1500")                          # VBA は実行時エラー 13

    def test_equipment(self) -> None:
        # Test:188 の直書きの値。LVC の厚 MIN は 1 なので 0.49 は範囲外
        self.assertEqual(maint.equipment_check("LVC", "5NB", "0.49", "730", "1500", "225")["text"], "オーダー厚 範囲外")
        self.assertEqual(maint.equipment_check("LVC", "5NB", "2", "730", "1500", "225")["text"], "可")
        self.assertEqual(maint.equipment_check("ZZZ", "5NB", "2", "730", "1500", "225")["text"], "可")  # 設備の行なし


class TerminalTest(MaintBase):
    def test_presets_and_clear(self) -> None:
        paths.set_mode(paths.OFFLINE)
        t = maint.apply_preset("HVC")                                     # 初期設定.HVC: モードはそのまま
        self.assertEqual((t["facility"], t["mode"]), ("HVC", paths.OFFLINE))
        t = maint.apply_preset("AIM")                                     # 初期設定.AIM: Test + オンライン
        self.assertEqual((t["facility_text"], t["mode"], t["mode_saved"]), ("Test", paths.ONLINE, True))
        terminal.set_facility("丸徳", "A3")
        t = maint.clear_terminal()                                        # 初期設定.DelSet
        self.assertEqual((t["facility"], t["facility_text"], t["line_no"]), ("", "存在なし", ""))
        self.assertTrue(t["unset"])
        self.assertFalse(t["mode_saved"])
        self.assertIsNone(terminal.line())                                # 次の起動で「この端末」を求める
        with self.assertRaises(ValueError):
            maint.apply_preset("XXX")


class MaintRouteTest(MaintBase):
    def setUp(self) -> None:
        super().setUp()
        from app import shell
        p = mock.patch.object(shell, "READY_SCREENS", shell.READY_SCREENS | {"maint"})
        p.start()
        self.addCleanup(p.stop)
        self.admin = mock.patch.object(auth, "can_edit", return_value=True)
        self.admin.start()
        self.addCleanup(self.admin.stop)
        self.client = make_client()
        app = self.client.application
        from app.routes import admin_auth, maint as maint_routes
        if "maint" not in app.blueprints:
            app.register_blueprint(maint_routes.bp)
        if "admin_auth" not in app.blueprints:
            app.register_blueprint(admin_auth.bp)
        admin_auth.reset_misses()
        p = mock.patch.object(maint_routes, "_store", lambda: st.LocalStore(self.tmp / "local.sqlite3"))
        p.start()
        self.addCleanup(p.stop)
        auth.unlock_master(admin_password.DEFAULT_PASSWORD)              # 既定は鍵を開けた状態で試す
        self.addCleanup(auth.lock_master)
        self.h = {"X-Tool-Token": app.config["TOKEN"]}

    def post(self, url: str, body: dict):
        return self.client.post(url, json=body, headers=self.h)

    def test_page_and_state(self) -> None:
        r = self.client.get("/maint", headers=self.h)
        self.assertEqual(r.status_code, 200)
        self.assertIn("mResult", r.get_data(as_text=True))
        d = self.client.get("/api/maint/state", headers=self.h).get_json()
        self.assertTrue(d["admin"])
        self.assertEqual(d["terminal"]["facility"], "HVC")
        self.assertEqual(d["keys"][:5], ["MG", "HaHiz", "TaHiz", "Naka", "Mimi"])

    def test_dist_api(self) -> None:
        d = self.post("/api/maint/dist", {"lot": "H9022K1"}).get_json()
        self.assertEqual(d["row"]["current"]["values"]["HaHiz"], "2.5")
        alt = self.alt()
        d = self.post("/api/maint/dist", {"lot": "H9022K1", "alt": str(alt)}).get_json()
        self.assertEqual(d["row"]["diff"], ["HaHiz", "TaHiz"])
        self.assertEqual(self.post("/api/maint/dist", {"lot": "a'b"}).status_code, 400)
        r = self.post("/api/maint/dist", {"lot": "H9022K1", "alt": str(self.tmp / "nai.sqlite3")})
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (400, "bad_alt"))
        d = self.post("/api/maint/dist_all", {"prefix": "H9", "alt": str(alt)}).get_json()
        self.assertTrue(all(x["lot"].startswith("H9") for x in d["rows"]))
        self.assertGreaterEqual(d["diffs"], 1)

    def test_sori_equipment_api(self) -> None:
        d = self.post("/api/maint/sori", {"atu": "3", "haba": "1200", "take": "1531"}).get_json()
        self.assertEqual(d["current"]["text"], "幅2.0 丈3.0")
        self.assertIsNone(d["alt"])
        self.assertEqual(self.post("/api/maint/sori", {"atu": "x", "haba": "1", "take": "1"}).status_code, 400)
        d = self.post("/api/maint/equipment", {"line": "LVC", "zai": "5NB", "atu": "2", "haba": "730", "take": "1500",
                                               "wei": "225"}).get_json()
        self.assertEqual(d["text"], "可")

    def test_terminal_api_logs(self) -> None:
        from ufkensa import trace
        with mock.patch.object(trace, "event", wraps=trace.event) as ev:
            d = self.post("/api/maint/terminal", {"action": "AIM"}).get_json()
        self.assertEqual(d["terminal"]["facility"], "Test")
        self.assertTrue(any(c.args[0] == "proc.end" and c.kwargs.get("proc") == "maint.terminal" for c in ev.call_args_list))
        d = self.post("/api/maint/terminal", {"action": "clear"}).get_json()
        self.assertTrue(d["terminal"]["unset"])
        self.assertEqual(self.post("/api/maint/terminal", {"action": "x"}).status_code, 400)

    def test_reads_open_to_everyone(self) -> None:
        """2026-11-02: 管理者でなくても、鍵を開けていなくても、読むだけの確かめは使える。"""
        auth.lock_master()
        self.admin.stop()
        with mock.patch.object(auth, "can_edit", return_value=False):
            self.assertEqual(self.client.get("/maint", headers=self.h).status_code, 200)
            for url, body in (("/api/maint/dist", {"lot": "H9022K1"}), ("/api/maint/dist_all", {"prefix": "H9"}),
                              ("/api/maint/sori", {"atu": "3", "haba": "1200", "take": "1500"}),
                              ("/api/maint/equipment", {"line": "LVC", "zai": "5NB", "atu": "2", "haba": "730",
                                                        "take": "1500", "wei": "225"})):
                self.assertEqual(self.post(url, body).status_code, 200, url)
            self.assertEqual(self.client.get("/api/maint/kari", headers=self.h).status_code, 200)
            d = self.client.get("/api/maint/state", headers=self.h).get_json()
            self.assertFalse(d["admin"])
            self.assertFalse(d["lock"]["can_edit"])
        self.admin.start()

    def test_edit_needs_unlock(self) -> None:
        """編集は 403 {code: locked}(理由は master_edit_reason)。鍵を開けると通る。"""
        auth.lock_master()
        r = self.post("/api/maint/terminal", {"action": "clear"})
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "locked"))
        self.assertEqual(r.get_json()["error"]["message"], auth.master_edit_reason())
        self.assertIn("管理者パスワード", r.get_json()["error"]["message"])
        self.assertEqual(self.post("/api/maint/kari", _kari_body()).status_code, 403)
        self.assertEqual(self.post("/api/maint/kari/delete", {"lot": "A111A11"}).status_code, 403)
        self.assertEqual(terminal.facility(), "HVC")                      # 変わっていない
        # 間違ったパスワード → 開かない。正しいパスワード → 開いて呼び直せる(JS の withUnlock と同じ流れ)
        d = self.post("/api/admin/unlock", {"password": "chigau"}).get_json()
        self.assertEqual((d["ok"], d["reason"]), (False, "管理者パスワードが違います"))
        self.assertFalse(d["state"]["can_edit"])
        d = self.post("/api/admin/unlock", {"password": admin_password.DEFAULT_PASSWORD}).get_json()
        self.assertTrue(d["ok"])
        self.assertTrue(d["state"]["can_edit"])
        self.assertEqual(self.post("/api/maint/terminal", {"action": "clear"}).status_code, 200)

    def test_admin_right_also_needed(self) -> None:
        """パスワードを通していても、管理者の権限が無ければ編集できない(両方要る。マスタで決まっているとき)。"""
        self.admin.stop()
        with mock.patch.object(auth, "can_edit", return_value=False), \
                mock.patch.object(auth, "setup_mode", return_value=False):
            r = self.post("/api/maint/terminal", {"action": "clear"})
            self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "locked"))
            self.assertIn("管理者", r.get_json()["error"]["message"])
            auth.lock_master()
            d = self.post("/api/admin/unlock", {"password": admin_password.DEFAULT_PASSWORD}).get_json()
            self.assertFalse(d["ok"])
            self.assertIn("管理者", d["reason"])
        self.admin.start()
        self.assertFalse(auth.master_unlocked())


def _kari_body(lot: str = "A111A11", **values: str) -> dict:
    v = {f.key: f.default for f in kari_lot.FIELDS}
    v.update({"用途コード": "C1", "製造材質": "5NB", "オーダー材質": "5NB", "製造調質": "H32", "オーダー調質": "H32",
              "オーダー板厚": "3", "オーダー板幅": "1250", "オーダー板丈": "2500",
              "製造板厚": "3", "製造板幅": "1250", "製造板丈": "2500"})
    v.update(values)
    return {"lot": lot, "ken_head": "H", "values": v}


class KariLotTest(MaintBase):
    def setUp(self) -> None:
        super().setUp()
        self.store = st.LocalStore(self.tmp / "local.sqlite3")
        self.addCleanup(self.store.close)

    def make(self, lot: str = "A111A11", **values: str) -> kari_lot.KariLot:
        b = _kari_body(lot, **values)
        return kari_lot.validate(b["lot"], b["ken_head"], b["values"])

    def test_validate_lot_and_ken(self) -> None:
        k = kari_lot.validate("ａ１２３ｂ４５", "h", _kari_body()["values"])     # 全角・小文字は直して受ける
        self.assertEqual((k.lot, k.ken_head, k.ken_no), ("A123B45", "H", "H111111"))
        # 検番の長さは本物(SIKAHIKINOW・SIKALOTNOW の 検査番号 H6X1176 など)と同じ 7
        self.assertEqual(len(k.ken_no), len(maint.dist_check("H9022K1")["ken_no"]))
        for bad in ("A11A11", "A111A1", "A111A111", "", "A111A1-", "N8281C"):
            with self.assertRaises(kari_lot.KariError, msg=bad) as cm:
                kari_lot.validate(bad, "H", _kari_body()["values"])
            self.assertEqual(cm.exception.field, "lot")
        for bad in ("", "1", "HH", "-"):
            with self.assertRaises(kari_lot.KariError, msg=bad) as cm:
                kari_lot.validate("A111A11", bad, _kari_body()["values"])
            self.assertEqual(cm.exception.field, "ken_head")

    def test_validate_numbers_and_defaults(self) -> None:
        for v in ("", "abc", "1.2.3"):
            with self.assertRaises(kari_lot.KariError) as cm:
                self.make(**{"オーダー板幅": v})
            self.assertEqual(cm.exception.field, "オーダー板幅")
        k = self.make(**{"製造板丈": "0", "オーダー板厚": " 0.5 "})
        self.assertEqual((k.values["製造板丈"], k.values["オーダー板厚"]), ("0", "0.5"))
        k = kari_lot.validate("A111A11", "H", {"オーダー板厚": "1", "オーダー板幅": "1", "オーダー板丈": "1",
                                               "製造板厚": "1", "製造板幅": "1", "製造板丈": "1"})
        self.assertEqual(k.values["歪等級"], "4")                          # 品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ の既定
        self.assertEqual(k.values["納入先"], "")

    def test_lot_data_and_hiki_mapping(self) -> None:
        k = self.make(**{"納入先": "ﾆﾎﾝ", "取引先": "ﾄﾘ", "送り先コード": "S9", "歪等級": "3"})
        d = k.lot_data()
        self.assertEqual((d.value("ﾛｯﾄ番号"), d.value("検査番号")), ("A111A11", "H111111"))
        self.assertEqual((d.value("ｵｰﾀﾞｰ板幅"), d.value("納入先名"), d.value("送り先ｺｰﾄﾞ")), ("1250", "ﾆﾎﾝ", "S9"))
        self.assertEqual(d.value("品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ"), "3")
        self.assertEqual(d.value("実績_設備ｺｰｽ"), "")                       # 指定の無い列は空
        self.assertEqual(k.hiki(), [("ﾆﾎﾝ", "ﾄﾘ")])                         # HikiArrCreate の (受注_納入先名称, 受注_取引先名称)

    def test_crud(self) -> None:
        self.assertEqual(kari_lot.list_all(self.store), [])
        self.assertEqual(kari_lot.save(self.store, self.make())[0], "作成")
        self.assertEqual(kari_lot.save(self.store, self.make("B222B22", **{"用途名": "テスト"}))[0], "作成")
        self.assertEqual([k.lot for k in kari_lot.list_all(self.store)], ["A111A11", "B222B22"])
        with self.assertRaises(kari_lot.KariError):                         # 同じロットNo は作れない
            kari_lot.save(self.store, self.make())
        # 直す(ロットNo も変えられる)
        action, k = kari_lot.save(self.store, self.make("C333C33", **{"オーダー板幅": "1000"}), "A111A11")
        self.assertEqual(action, "変更")
        self.assertIsNone(kari_lot.get(self.store, "A111A11"))
        got = kari_lot.get(self.store, "c333c33")
        self.assertEqual((got.values["オーダー板幅"], got.ken_no), ("1000", "H111111"))
        self.assertTrue(got.updated)
        with self.assertRaises(kari_lot.KariError):
            kari_lot.save(self.store, self.make("D444D44"), "ZZZZZZZ")      # 直す元が無い
        self.assertTrue(kari_lot.delete(self.store, "B222B22"))
        self.assertFalse(kari_lot.delete(self.store, "B222B22"))
        self.assertEqual([k.lot for k in kari_lot.list_all(self.store)], ["C333C33"])
        st2 = st.LocalStore(self.tmp / "local.sqlite3")                     # 保存されている(開き直しても同じ)
        self.addCleanup(st2.close)
        self.assertEqual(kari_lot.get(st2, "C333C33").values["オーダー板幅"], "1000")

    def test_real_lot_number_refused(self) -> None:
        with self.assertRaises(kari_lot.KariError) as cm:
            kari_lot.save(self.store, self.make("L726M50"))                   # 仕掛にあるロット(A111A11 の形)
        self.assertIn("仕掛に同じロットNo", str(cm.exception))

    def test_from_real(self) -> None:
        d = kari_lot.from_real("h9022k1")
        self.assertEqual((d["source"], d["ken_no"], d["ken_head"]), ("H9022K1", "H6X1176", "H"))
        hiki = quality.hiki_arr_create("H9022K1")
        self.assertEqual((d["values"]["納入先"], d["values"]["取引先"]), hiki[0])
        data = quality.read_lot("H9022K1")
        self.assertEqual(d["values"]["オーダー板幅"], data.value("ｵｰﾀﾞｰ板幅"))
        self.assertEqual(d["values"]["歪等級"], data.value("品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ"))
        with self.assertRaises(kari_lot.KariError):
            kari_lot.from_real("ZZZZZZZ")

    def test_same_result_as_real_lot(self) -> None:
        """本物のロットと同じ値の仮ロットは、歪公差も同じになる(引当のある仕掛の全ロット)。"""
        checked = 0
        for lot in maint.lot_numbers():
            if not quality.hiki_arr_create(lot):
                continue                                                     # 本物は引当が無いと引かない
            d = kari_lot.from_real(lot)
            k = kari_lot.validate("A111A11", d["ken_head"], d["values"], d["hiki_extra"], lot)
            real, kari = maint.dist_check(lot), maint.dist_check_kari(k)
            self.assertEqual(kari["current"]["status"], real["current"]["status"], lot)
            self.assertEqual(kari["current"]["values"], real["current"]["values"], lot)
            self.assertEqual(kari["current"]["messages"], real["current"]["messages"], lot)
            checked += 1
        self.assertGreaterEqual(checked, 20)
        r = maint.dist_check_kari(kari_lot.validate("A111A11", "H", kari_lot.from_real("H9022K1")["values"]))
        self.assertEqual((r["kari"], r["ken_no"], r["current"]["values"]["HaHiz"]), (True, "H111111", "2.5"))

    def test_kari_with_alt_and_all(self) -> None:
        d = kari_lot.from_real("H9022K1")
        k = kari_lot.validate("H902K21", d["ken_head"], d["values"], d["hiki_extra"], "H9022K1")
        r = maint.dist_check_kari(k, self.alt())
        self.assertEqual(r["diff"], ["HaHiz", "TaHiz"])                     # 変更後の表で 9.0
        out = maint.dist_check_all(prefix="H9", kari_lots=[k, self.make()])
        self.assertEqual(out["kari_used"], ["H902K21"])                     # 頭 H9 に合う仮ロットだけ
        kari_rows = [x for x in out["rows"] if x["kari"]]
        self.assertEqual([x["lot"] for x in kari_rows], ["H902K21"])
        self.assertEqual(out["count"], len(out["rows"]))
        self.assertFalse(any(x["kari"] for x in maint.dist_check_all(prefix="H9")["rows"]))


class KariRouteTest(MaintRouteTest):
    """仮ロットの API(MaintRouteTest の試験も、この setUp でもう一度通る)。"""

    def test_kari_api_flow(self) -> None:
        d = self.client.get("/api/maint/kari", headers=self.h).get_json()
        self.assertEqual((d["rows"], d["ken_len"]), ([], 7))
        self.assertIn({"key": "納入先", "column": "納入先名", "hiki": "受注_納入先名称"},
                      [{k: f[k] for k in ("key", "column", "hiki")} for f in d["fields"]])
        r = self.post("/api/maint/kari", _kari_body("A11A11"))
        self.assertEqual((r.status_code, r.get_json()["error"]["field"]), (400, "lot"))
        src = self.client.get("/api/maint/kari/from_real?lot=H9022K1", headers=self.h).get_json()
        self.assertEqual(src["ken_head"], "H")
        self.assertEqual(self.client.get("/api/maint/kari/from_real?lot=ZZZZZZZ", headers=self.h).status_code, 404)
        body = {"lot": "H902K21", "ken_head": src["ken_head"], "values": src["values"],
                "hiki_extra": src["hiki_extra"], "source": "H9022K1"}
        # 保存の前でも欄の値で引ける(読むだけなので鍵は要らない)
        auth.lock_master()
        d = self.post("/api/maint/kari/dist", body).get_json()
        self.assertEqual((d["row"]["kari"], d["row"]["current"]["values"]["HaHiz"]), (True, "2.5"))
        auth.unlock_master(admin_password.DEFAULT_PASSWORD)
        d = self.post("/api/maint/kari", body).get_json()
        self.assertEqual((d["action"], d["row"]["ken_no"], len(d["rows"])), ("作成", "H111111", 1))
        self.assertEqual(self.post("/api/maint/kari", body).status_code, 400)           # 同じロットNo
        # 歪公差チェックのロット欄に仮ロットを入れられる
        d = self.post("/api/maint/dist", {"lot": "h902k21"}).get_json()
        self.assertTrue(d["row"]["kari"])
        self.assertEqual(d["row"]["current"]["values"],
                         self.post("/api/maint/dist", {"lot": "H9022K1"}).get_json()["row"]["current"]["values"])
        # まとめて引く: 仮ロットも(既定)/ 外す
        d = self.post("/api/maint/dist_all", {"prefix": "H9"}).get_json()
        self.assertEqual(d["kari_used"], ["H902K21"])
        d = self.post("/api/maint/dist_all", {"prefix": "H9", "include_kari": False}).get_json()
        self.assertEqual(d["kari_used"], [])
        # 直す・消す
        body2 = dict(body, original="H902K21", values=dict(body["values"], **{"オーダー板幅": "999"}))
        self.assertEqual(self.post("/api/maint/kari", body2).get_json()["action"], "変更")
        self.assertEqual(self.post("/api/maint/kari/delete", {"lot": "H902K21"}).get_json()["rows"], [])
        self.assertEqual(self.post("/api/maint/kari/delete", {"lot": "H902K21"}).status_code, 404)


class AdminAuthTest(MaintRouteTest):
    """/api/admin/*(app/routes/admin_auth.py)。"""

    def tearDown(self) -> None:
        user_settings.delete(config.KEY_ADMIN_PASSWORD)
        super().tearDown()

    def test_state_lock_unlock(self) -> None:
        d = self.client.get("/api/admin/state", headers=self.h).get_json()
        self.assertEqual((d["unlocked"], d["can_edit"], d["admin"], d["reason"]), (True, True, True, ""))
        self.assertFalse(d["custom"])
        self.assertEqual(d["min_length"], admin_password.MIN_LENGTH)
        d = self.post("/api/admin/lock", {}).get_json()["state"]
        self.assertEqual((d["unlocked"], d["can_edit"]), (False, False))
        self.assertEqual(d["reason"], auth.master_edit_reason())

    def test_password_not_logged(self) -> None:
        from ufkensa import trace
        auth.lock_master()
        secret = "himitsu-XYZ"
        with mock.patch.object(trace, "event", wraps=trace.event) as ev, self.assertLogs(level="DEBUG") as logs:
            self.post("/api/admin/unlock", {"password": secret})
            self.post("/api/admin/password", {"current": secret, "new": secret + "2", "confirm": secret + "2"})
        self.assertNotIn(secret, repr(ev.call_args_list))
        self.assertNotIn(secret, "\n".join(logs.output))
        self.assertTrue(any(c.kwargs.get("proc") == "admin.unlock" for c in ev.call_args_list))

    def test_cool_down(self) -> None:
        auth.lock_master()
        for _ in range(admin_password.MAX_MISSES):
            self.assertFalse(self.post("/api/admin/unlock", {"password": "x"}).get_json()["ok"])
        d = self.post("/api/admin/unlock", {"password": admin_password.DEFAULT_PASSWORD}).get_json()
        self.assertFalse(d["ok"])
        self.assertIn("待って", d["reason"])
        from app.routes import admin_auth
        admin_auth.reset_misses()
        self.assertTrue(self.post("/api/admin/unlock", {"password": admin_password.DEFAULT_PASSWORD}).get_json()["ok"])

    def test_change_password(self) -> None:
        dflt = admin_password.DEFAULT_PASSWORD
        cases = ((("chigau", "abcd", "abcd"), "wrong_password", "current"),
                 ((dflt, "abc", "abc"), "too_short", "new"),
                 ((dflt, "abcd", "abce"), "mismatch", "confirm"),
                 ((dflt, dflt, dflt), "same", "new"))
        for (cur, new, conf), code, field in cases:
            r = self.post("/api/admin/password", {"current": cur, "new": new, "confirm": conf})
            self.assertEqual((r.status_code, r.get_json()["error"]["code"], r.get_json()["error"]["field"]),
                             (400, code, field), code)
        self.assertFalse(admin_password.is_custom())
        d = self.post("/api/admin/password", {"current": dflt, "new": "atarashii", "confirm": "atarashii"}).get_json()
        self.assertTrue(d["ok"])
        self.assertTrue(d["state"]["custom"])
        self.assertNotIn("atarashii", str(user_settings.get(config.KEY_ADMIN_PASSWORD)))   # 平文で持たない
        auth.lock_master()
        self.assertFalse(self.post("/api/admin/unlock", {"password": dflt}).get_json()["ok"])
        self.assertTrue(self.post("/api/admin/unlock", {"password": "atarashii"}).get_json()["ok"])

    def test_change_password_needs_admin_right(self) -> None:
        self.admin.stop()
        with mock.patch.object(auth, "can_edit", return_value=False):
            r = self.post("/api/admin/password", {"current": admin_password.DEFAULT_PASSWORD, "new": "abcd", "confirm": "abcd"})
            self.assertEqual(r.status_code, 403)
        self.admin.start()
        self.assertFalse(admin_password.is_custom())


if __name__ == "__main__":
    unittest.main()



class JudgeDisplayTest(MaintRouteTest):
    """保守 → 判定の画像: 表示時間・大きさ(PC ごと。2026-11-09)。"""

    def tearDown(self) -> None:
        user_settings.save("judge_display", {})
        super().tearDown()

    def test_normalize_clamps(self) -> None:
        from ufkensa.kensa import judge_images as ji
        d = ji.normalize_display({})
        self.assertEqual(d, {"ok": {"ms": 3000, "size": 80}, "ng": {"ms": 3000, "size": 80},
                             "diag": {"ms": 2000, "size": 80}})
        d = ji.normalize_display({"ok": {"ms": 10, "size": 5}, "ng": {"ms": 99999, "size": 500},
                                  "diag": {"ms": "x", "size": "60.4"}})
        self.assertEqual(d["ok"], {"ms": 500, "size": 20})
        self.assertEqual(d["ng"], {"ms": 30000, "size": 100})
        self.assertEqual(d["diag"], {"ms": 2000, "size": 60})
        self.assertEqual(ji.normalize_display("bad")["ok"]["ms"], 3000)

    def test_route_get_save_and_kensa_uses_it(self) -> None:
        d = self.client.get("/api/maint/judge_display", headers=self.h).get_json()
        self.assertEqual(d["display"]["diag"], {"ms": 2000, "size": 80})
        self.assertEqual(len([x for x in d["slots"] if x["kind"] == "diag"]), 5)
        d = self.post("/api/maint/judge_display", {"display": {"ok": {"ms": 1500, "size": 100},
                                                               "ng": {"ms": 60000, "size": 40}}}).get_json()
        self.assertEqual(d["display"]["ok"], {"ms": 1500, "size": 100})
        self.assertEqual(d["display"]["ng"], {"ms": 30000, "size": 40})
        self.assertEqual(user_settings.get("judge_display")["ok"]["size"], 100)
        from app.routes import kensa as kensa_routes
        app = self.client.application
        if "kensa" not in app.blueprints:
            app.register_blueprint(kensa_routes.bp)
        k = self.client.get("/api/kensa/judge_images", headers=self.h).get_json()
        self.assertEqual(k["show_ms"], {"ok": 1500, "ng": 30000, "diag": 2000})
        self.assertEqual(k["size"], {"ok": 100, "ng": 40, "diag": 80})

    def test_save_needs_unlock(self) -> None:
        auth.lock_master()
        r = self.post("/api/maint/judge_display", {"display": {"ok": {"ms": 1000}}})
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "locked"))
        self.assertEqual(self.client.get("/api/maint/judge_display", headers=self.h).status_code, 200)
