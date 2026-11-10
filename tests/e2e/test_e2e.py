"""ブラウザの通しの試験(2026-11-10 リファクタリング 第1段 5)

画面(JS)の動きを本物のブラウザ(Playwright の Chromium)で確かめる。中身は tests/e2e/kensa.spec.js:
受信機の状態・検査の通し(検査員・ロット・作業・受信機/受信 API の測定・Min/Max・印刷)・公差外の確認の窓・
サイズ修正の窓・1366×768 のはみ出し・作業範囲の選び欄(F4)・全ヵ所の検入高さ(F5)。
続けて tests/e2e/screens.spec.js(整理 6): キャスク・全ヵ所・Aインプット・校正を開き、問いのある操作を画面の
確認窓で答え、受信機の値がその画面に届くこと。

**ふだんは飛ばす。** `UFKENSA_E2E=1` のときだけ動く(node と Playwright・Chromium が要り、1 分ほどかかるため。
ラインPC・CI に入っていないことがある)。動かし方:

    UFKENSA_E2E=1 python -m pytest -q tests/e2e

Python の playwright は使わない(入っていない)。node の playwright を `require` できること
(NODE_PATH など)。画面の写しが欲しいときは `UFKENSA_E2E_SHOTS=<フォルダ>`。
"""
from __future__ import annotations

import os
import shutil
import socket
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
ENABLED = os.environ.get("UFKENSA_E2E") == "1"
#: 動かす順(screens.spec.js は kensa.spec.js の後の状態を使う)
SPECS = ("kensa.spec.js", "screens.spec.js")


def _node() -> str:
    return shutil.which("node") or ""


def _playwright_ok(node: str) -> str:
    """node から playwright と Chromium が使えるか。使えなければ理由を返す("" なら使える)。"""
    probe = ("const { chromium } = require('playwright');"
             "chromium.launch().then(b => b.close()).then(() => process.exit(0), e => { console.error(e.message); process.exit(3); });")
    try:
        r = subprocess.run([node, "-e", probe], capture_output=True, text=True, timeout=60, cwd=HERE)
    except (OSError, subprocess.TimeoutExpired) as e:
        return str(e)
    return "" if r.returncode == 0 else (r.stderr.strip().splitlines() or ["playwright を読めません"])[-1]


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


@unittest.skipUnless(ENABLED, "UFKENSA_E2E=1 のときだけ動かす(ブラウザの通しの試験)")
class BrowserE2E(unittest.TestCase):
    def setUp(self) -> None:
        self.node = _node()
        if not self.node:
            self.skipTest("node がありません")
        why = _playwright_ok(self.node)
        if why:
            self.skipTest("Playwright(node)が使えません: " + why)

    def test_kensa_screens(self) -> None:
        port = _free_port()
        env = dict(os.environ)
        env.setdefault("PYTHONPATH", str(ROOT))
        err = tempfile.TemporaryFile(mode="w+", encoding="utf-8")      # サーバの stderr(詰まらないようファイルへ)
        self.addCleanup(err.close)
        srv = subprocess.Popen([sys.executable, "-m", "tests.e2e.server", str(port)], cwd=ROOT, env=env,
                               stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=err, text=True)
        try:
            line = srv.stdout.readline()
            if not line.startswith("READY"):
                err.seek(0)
                self.fail("試験のサーバが起動しません: " + line + err.read()[-2000:])
            pty_name = line.split()[2]
            nenv = dict(os.environ, E2E_BASE=f"http://127.0.0.1:{port}", E2E_PTY=pty_name)
            if os.environ.get("UFKENSA_E2E_SHOTS"):
                nenv["E2E_SHOTS"] = os.environ["UFKENSA_E2E_SHOTS"]
            # 順に動かす(screens は kensa で済ませた検査員・ロット・Min/Max・受信機を使う)
            for spec in SPECS:
                with self.subTest(spec=spec):
                    r = subprocess.run([self.node, str(HERE / spec)], cwd=HERE, env=nenv,
                                       capture_output=True, text=True, timeout=240)
                    out = r.stdout + r.stderr
                    self.assertIn("E2E_RESULT", out, out[-4000:])
                    self.assertEqual(r.returncode, 0, "\n".join(l for l in out.splitlines() if "FAIL" in l or "PAGEERROR" in l
                                                              or "E2E_RESULT" in l) or out[-4000:])
        finally:
            try:
                srv.stdin.close()
                srv.wait(timeout=15)
            except (OSError, subprocess.TimeoutExpired):
                srv.kill()
            srv.stdout.close()


if __name__ == "__main__":
    unittest.main()
