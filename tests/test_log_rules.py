"""ログの決まり(all-tools に合わせる): 置き場所・名前・処理の開始と終了・処理時間・エラーの詳細"""
from __future__ import annotations

import json
import time
import unittest
from unittest import mock

from tests.test_trace import _LogDirCase
from ufkensa import config, trace


class LogRulesTest(_LogDirCase):
    def lines(self):
        p = trace.file_for()
        return [json.loads(l) for l in p.read_text(encoding="utf-8").splitlines()] if p.exists() else []

    def test_place_and_name(self):
        from ufkensa import logging_utils
        place = logging_utils.resolve_log_dir(str(self.dir.parent / "共有ログ"))
        self.assertEqual(place, self.dir.parent / "共有ログ" / "ufkensa")   # <出力パス>\<ツールID>\(直下に置かない)
        self.assertEqual(config.tool_log_dir(self.dir).name, "ufkensa")
        self.assertTrue(trace.file_for().name.startswith(f"ufkensa_記録_{trace.pc_name()}_"))   # ツールIDで始まり PC 名
        lp = logging_utils.log_path_for(__import__("datetime").date.today())
        self.assertTrue(lp.name.startswith(f"ufkensa_{trace.pc_name()}_"))
        self.assertEqual(config.TOOL_LOG_NAME, "ufkensa")                  # 半角英小文字

    def test_start_and_end_are_paired_with_time(self):
        with trace.timed("db.read", "品質基準 を読む", path="p", table="t") as info:
            info["rows"] = 5
        evs = [(r["ev"], r["lv"]) for r in self.lines() if r.get("proc") == "db.read"]
        self.assertEqual(evs, [("proc.start", "INFO"), ("proc.end", "INFO")])
        end = [r for r in self.lines() if r["ev"] == "proc.end"][-1]
        self.assertIn("ms", end)
        self.assertEqual(end["data"]["rows"], 5)

    def test_slow_steps(self):
        self.assertEqual(trace.slow_label(2.9), "")
        self.assertEqual(trace.slow_label(3.0), "3秒以上")
        self.assertEqual(trace.slow_label(10.0), "10秒以上")
        self.assertEqual(trace.slow_label(45.0), "30秒以上")
        with mock.patch.object(trace, "SLOW_STEPS", ((0.02, "3秒以上"),)):
            with trace.timed("slow"):
                time.sleep(0.03)
        end = [r for r in self.lines() if r.get("proc") == "slow" and r["ev"] == "proc.end"][-1]
        self.assertEqual((end["lv"], end["slow"]), ("WARN", "3秒以上"))

    def test_still_running_is_noticed(self):
        with mock.patch.object(trace, "STILL_RUNNING_AT", (0.03,)):
            with trace.timed("hang"):
                time.sleep(0.08)
        self.assertTrue([r for r in self.lines() if r["ev"] == "proc.slow" and r["lv"] == "WARN"])

    def test_error_details(self):
        trace.note_used("/data/品質基準.sqlite3", "read")
        trace.event("screen", "検査画面を開いた")
        with self.assertRaises(ValueError):
            with trace.timed("kensa.lot", "ロット登録"):
                raise ValueError("こわれた")
        e = [r for r in self.lines() if r["ev"] == "proc.error"][-1]
        self.assertEqual(e["lv"], "ERROR")
        for k in ("t", "app", "ver", "pc", "procs", "used", "recent", "exc"):
            self.assertIn(k, e)
        self.assertEqual(e["procs"], ["kensa.lot"])                        # 実行していた処理
        self.assertIn("read:/data/品質基準.sqlite3", e["used"])           # 使っていたファイル・DB
        self.assertTrue(any("検査画面を開いた" in x for x in e["recent"]))  # 直前の処理状況
        self.assertEqual(e["exc"]["type"], "ValueError")
        self.assertTrue(e["exc"]["tb"])                                    # スタックトレース

    def test_expected_stop_is_a_warning(self):
        from ufkensa.kensa.flow import Stop
        with self.assertRaises(Stop):
            with trace.timed("kensa.print-check", expected=(Stop,)):
                raise Stop("外観未入力")
        e = [r for r in self.lines() if r.get("proc") == "kensa.print-check"][-1]
        self.assertEqual((e["ev"], e["lv"], e["stopped"]), ("proc.end", "WARN", "Stop"))


if __name__ == "__main__":
    unittest.main()


class LogPathSettingTest(_LogDirCase):
    """ログの出力パスは 設定 → 使用パス(と ログ)で決められる。下に ufkensa フォルダを作る。"""

    def setUp(self):
        super().setUp()
        import os
        from unittest import mock as _m
        from tests._web import make_client
        from ufkensa import auth
        env = _m.patch.dict(os.environ, {"USERNAME": "admin1", "COMPUTERNAME": "PC-1"})
        env.start()
        self.addCleanup(env.stop)
        auth.end()
        self.addCleanup(auth.end)
        admin = _m.patch.object(auth, "can_edit", return_value=True)
        admin.start()
        self.addCleanup(admin.stop)
        self.c = make_client()
        self.h = {"X-Tool-Token": self.c.application.config["TOKEN"]}

    def test_set_from_paths_tab(self):
        out = self.dir.parent / "共有" / "ログ"
        st = self.c.post("/api/settings/path", json={"key": "log", "path": str(out)}, headers=self.h).get_json()
        row = [p for p in st["paths"][st["mode"]] if p["key"] == "log"][0]
        self.assertEqual(row["path"], str(out))
        self.assertEqual(config.LOG_DIR, out / "ufkensa")                 # 直下ではなく ufkensa フォルダ
        self.assertEqual(st["log"]["dir"], str(out / "ufkensa"))          # ログの面と同じ設定
        chk = self.c.post("/api/settings/path/check", json={"key": "log"}, headers=self.h).get_json()
        self.assertTrue(chk["ok"])
        trace.event("x", "y")
        self.assertTrue((out / "ufkensa").glob("ufkensa_記録_*.jsonl"))
        st = self.c.post("/api/settings/path", json={"key": "log", "reset": True}, headers=self.h).get_json()
        self.assertEqual([p for p in st["paths"][st["mode"]] if p["key"] == "log"][0]["path"], "")

    def test_unwritable_keeps_current(self):
        before = config.LOG_DIR
        blocker = self.dir.parent / "ファイル"
        blocker.write_text("x", encoding="utf-8")
        r = self.c.post("/api/settings/path", json={"key": "log", "path": str(blocker / "下")}, headers=self.h)
        self.assertEqual(r.status_code, 400)
        self.assertIn("書けません", r.get_json()["error"]["message"])
        self.assertEqual(config.LOG_DIR, before)
