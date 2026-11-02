"""保守画面 /maint と /api/maint/*(app/routes/maint.py → ufkensa/kensa/maint.py)

歪公差の照合そのものは test_kensa_quality が見ている。ここは Test フォームと同じ 8 項目・書式への組み立て、
別の品質基準との比べ方(このスレッドだけ差し替える)、ソリ公差・設備仕様の表引き、端末設定のセット、管理者の制限を見る。
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
from ufkensa import auth, user_settings
from ufkensa.kensa import maint, paths, quality, terminal


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
        if "maint" not in app.blueprints:
            from app.routes import maint as maint_routes
            app.register_blueprint(maint_routes.bp)
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

    def test_not_admin(self) -> None:
        self.admin.stop()
        with mock.patch.object(auth, "can_edit", return_value=False):
            for url, body in (("/api/maint/dist", {"lot": "H9022K1"}), ("/api/maint/dist_all", {}),
                              ("/api/maint/sori", {}), ("/api/maint/equipment", {}),
                              ("/api/maint/terminal", {"action": "clear"})):
                self.assertEqual(self.post(url, body).status_code, 403, url)
            self.assertFalse(self.client.get("/api/maint/state", headers=self.h).get_json()["admin"])
        self.assertEqual(terminal.facility(), "HVC")                      # 変わっていない
        self.admin.start()


if __name__ == "__main__":
    unittest.main()
