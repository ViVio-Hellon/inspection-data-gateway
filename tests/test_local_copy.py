"""local_copy.py: 共有の本体を各 PC の手元へ写して、手元から起動する。"""
from __future__ import annotations

import os
import shutil
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

import local_copy as lc


class LocalCopyTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="ufk_lc_"))
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.src = self.tmp / "share" / "UF検査判定"
        for rel, text in (("start_app.py", "print(1)"), ("ufkensa/a.py", "x = 1"), ("app/static/x.js", "1"),
                          ("tests/t.py", "no"), ("docs/d.md", "no"), ("ufkensa/__pycache__/a.pyc", "no"),
                          ("Access.log", "no"), ("配布設定/設定.json", "{}")):
            p = self.src / rel
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(text, encoding="utf-8")
        self.base = self.tmp / "local" / "app"

    def test_network_path(self) -> None:
        self.assertTrue(lc.is_network_path(Path(r"\\server\share\UF")))
        self.assertTrue(lc.is_network_path(Path("//server/share/UF")))
        self.assertFalse(lc.is_network_path(self.src))

    def test_copy_once_and_new_version_on_change(self) -> None:
        d1 = lc.ensure_copy(self.src, self.base)
        self.assertTrue((d1 / "start_app.py").is_file())
        self.assertTrue((d1 / "app/static/x.js").is_file())
        for skipped in ("tests", "docs", "ufkensa/__pycache__", "Access.log", "配布設定"):
            self.assertFalse((d1 / skipped).exists(), skipped)
        self.assertTrue((d1 / ".complete").is_file())
        self.assertEqual(lc.ensure_copy(self.src, self.base), d1)              # 同じ版は写さない
        time.sleep(1.1)
        (self.src / "ufkensa/a.py").write_text("x = 2", encoding="utf-8")       # 本体を更新
        d2 = lc.ensure_copy(self.src, self.base)
        self.assertNotEqual(d2, d1)
        self.assertEqual((d2 / "ufkensa/a.py").read_text(encoding="utf-8"), "x = 2")
        self.assertEqual((d1 / "ufkensa/a.py").read_text(encoding="utf-8"), "x = 1")   # 動いている写しはそのまま
        time.sleep(1.1)
        (self.src / "ufkensa/a.py").write_text("x = 3", encoding="utf-8")
        d3 = lc.ensure_copy(self.src, self.base)
        self.assertEqual(sorted(p.name for p in self.base.glob("v-*")), sorted([d2.name, d3.name]))  # 2 つまで残す

    def test_relaunch(self) -> None:
        env = {k: v for k, v in os.environ.items() if k not in (lc.ENV_FROM, lc.ENV_OFF, "UFKENSA_DISTRIBUTION_DIR")}
        with mock.patch.dict(os.environ, env, clear=True):
            self.assertIsNone(lc.relaunch_from_local(self.src, []))            # 手元の本体ならそのまま
            with mock.patch.object(lc, "is_network_path", return_value=True), \
                    mock.patch.object(lc, "local_base", return_value=self.base), \
                    mock.patch.object(lc.subprocess, "call", return_value=0) as call:
                self.assertEqual(lc.relaunch_from_local(self.src, ["--no-browser"]), 0)
            args, kw = call.call_args
            self.assertTrue(args[0][1].endswith("start_app.py"))
            self.assertIn(str(self.base), args[0][1])
            self.assertEqual(args[0][2:], ["--no-browser"])
            self.assertEqual(kw["env"][lc.ENV_FROM], str(self.src))
            self.assertEqual(kw["env"]["UFKENSA_DISTRIBUTION_DIR"], str(self.src / "配布設定"))
            # 写した先から起動された 2 回目は、写し直さない
            with mock.patch.dict(os.environ, {lc.ENV_FROM: str(self.src)}):
                self.assertIsNone(lc.relaunch_from_local(self.src, []))

    def test_copy_failure_falls_back(self) -> None:
        with mock.patch.object(lc, "is_network_path", return_value=True), \
                mock.patch.object(lc, "ensure_copy", side_effect=OSError("disk full")), \
                mock.patch.dict(os.environ, {}, clear=False):
            os.environ.pop(lc.ENV_FROM, None)
            os.environ.pop(lc.ENV_OFF, None)
            self.assertIsNone(lc.relaunch_from_local(self.src, []))            # 共有からそのまま起動


if __name__ == "__main__":
    unittest.main()
