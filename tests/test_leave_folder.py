"""終わったあとにアプリのフォルダを消せる(0.25.0)

Windows はプロセスがいまいるフォルダを消せない。プレビューの cscript・Excel、「現物を開く」の Excel・Word、
「フォルダを開く」のエクスプローラは、サーバのいまいるフォルダを引き継ぐので、起動したらアプリのフォルダの外へ移る。
終わるときは、プレビューで起こして残った隠れた Excel・Word を片付ける(前は次の起動と stop.bat だけ)。
"""
from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path
from unittest import mock

import start_app
from ufkensa import app_config


class LeaveAppFolderTest(unittest.TestCase):
    def setUp(self):
        self.cwd = os.getcwd()
        self.path = list(sys.path)
        self.addCleanup(os.chdir, self.cwd)
        self.addCleanup(setattr, sys, "path", self.path)

    def test_moves_out_and_keeps_imports_working(self):
        os.chdir(start_app.APP_ROOT)
        sys.path.insert(0, "")                                         # python start_app.py のときの形
        start_app.leave_app_folder()
        here = Path(os.getcwd()).resolve()
        self.assertEqual(here, app_config.local_dir("work").resolve())
        self.assertNotIn(start_app.APP_ROOT.resolve(), [here, *here.parents])   # アプリのフォルダの外
        self.assertNotIn("", sys.path)
        self.assertIn(str(start_app.APP_ROOT), sys.path)
        import importlib
        importlib.import_module("ufkensa.kensa.vbx")                 # 後から読み込む部品も見つかる

    def test_children_tidied_on_exit(self):
        start_app.tidy_children()                                        # 止めるものが無くても終わる

if __name__ == "__main__":
    unittest.main()
