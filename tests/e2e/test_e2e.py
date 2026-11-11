"""ブラウザの通しの試験(2026-11-10 リファクタリング 第1段 5)

画面(JS)の動きを本物のブラウザ(Playwright の Chromium)で確かめる。中身は tests/e2e/kensa.spec.js:
受信機の状態・検査の通し(検査員・ロット・作業・受信機/受信 API の測定・Min/Max・印刷)・公差外の確認の窓・
サイズ修正の窓・1366×768 のはみ出し・作業範囲の選び欄(F4)・全ヵ所の検入高さ(F5)。
続けて tests/e2e/screens.spec.js(整理 6): キャスク・全ヵ所・Aインプット・校正を開き、問いのある操作を画面の
確認窓で答え、受信機の値がその画面に届くこと。

**ふだんは飛ばす。** `UFKENSA_E2E=1` のときだけ動く(node と Playwright・Chromium が要り、1 分半ほどかかるため。
ラインPC・CI に入っていないことがある)。動かし方(どちらも同じ):

    python -m tests.e2e                      # 前提を確かめてから流す(画面の写しは --shots <フォルダ>)
    UFKENSA_E2E=1 python -m pytest -q tests/e2e

Python の playwright は使わない(入っていない)。node の playwright を `require` できること
(NODE_PATH など)。画面の写しが欲しいときは `UFKENSA_E2E_SHOTS=<フォルダ>`。

後始末: 試験のサーバ・node(とその子の Chromium)はそれぞれ別のプロセスグループで起こし、終わったら・時間切れ・
中断(Ctrl+C)のときもグループごと止める。サーバは一時フォルダ(台帳の写し)を消し、疑似端末を閉じて終わる。
"""
from __future__ import annotations

import os
import shutil
import signal
import socket
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
ENABLED = os.environ.get("UFKENSA_E2E") == "1"
#: 動かす順(screens.spec.js は kensa.spec.js の後の状態を使う。single_screen.spec.js は画面の取り合い)
SPECS = ("kensa.spec.js", "screens.spec.js", "single_screen.spec.js")


def _node() -> str:
    return shutil.which("node") or ""


def _playwright_ok(node: str) -> str:
    """node から playwright と Chromium が使えるか。使えなければ理由を返す("" なら使える)。"""
    probe = ("const { chromium } = require('playwright');"
             "chromium.launch().then(b => b.close()).then(() => process.exit(0), e => { console.error(e.message); process.exit(3); });")
    try:
        code, out = run_node(node, ["-e", probe], dict(os.environ), timeout=60)
    except OSError as e:
        return str(e)
    return "" if code == 0 else (out.strip().splitlines() or ["playwright を読めません"])[-1]


def _killpg(proc: subprocess.Popen) -> None:
    """``proc`` のプロセスグループ(node なら子の Chromium も)を止める。もう居なければ何もしない。"""
    if proc.poll() is not None and os.name != "posix":
        return
    try:
        if os.name == "posix":
            os.killpg(proc.pid, signal.SIGKILL)
        else:
            proc.kill()
    except (ProcessLookupError, PermissionError, OSError):
        pass
    try:
        proc.wait(timeout=10)
    except subprocess.TimeoutExpired:
        pass


def _group() -> dict:
    """子プロセスを別のプロセスグループで起こす(まとめて止められるように)。"""
    if os.name == "posix":
        return {"start_new_session": True}
    return {"creationflags": getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)}


def run_node(node: str, args: list[str], env: dict, timeout: float = 240) -> tuple[int, str]:
    """node を args で動かし (終了コード, 出力) を返す。時間切れ・中断ならグループごと止める。"""
    proc = subprocess.Popen([node, *args], cwd=HERE, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                            text=True, encoding="utf-8", errors="replace", **_group())
    try:
        out, _ = proc.communicate(timeout=timeout)
        return proc.returncode, out
    except subprocess.TimeoutExpired:
        _killpg(proc)
        out, _ = proc.communicate()
        return -1, (out or "") + f"\n時間切れ({timeout:.0f} 秒)"
    finally:
        _killpg(proc)                      # 中断のとき・node が先に終わって Chromium が残ったときも


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
                               stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=err, text=True, **_group())
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
                    code, out = run_node(self.node, [str(HERE / spec)], nenv)
                    self.assertIn("E2E_RESULT", out, out[-4000:])
                    self.assertEqual(code, 0, "\n".join(l for l in out.splitlines() if "FAIL" in l or "PAGEERROR" in l
                                                              or "E2E_RESULT" in l) or out[-4000:])
        finally:
            try:
                srv.stdin.close()                       # サーバは標準入力が閉じると後始末をして止まる
                srv.wait(timeout=15)
            except (OSError, subprocess.TimeoutExpired):
                pass
            _killpg(srv)                                # 止まらなければ・残った子があればグループごと
            srv.stdout.close()


if __name__ == "__main__":
    unittest.main()
