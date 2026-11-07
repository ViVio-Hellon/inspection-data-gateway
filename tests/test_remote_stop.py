"""管理者以外停止モード(`remote_stop.py` は all-tools の写し。受け方は `ufkensa/remote_stop_link.py`)

    何も残っていなければ        帯を見せてから、まもなく止まる
    残っているもの              利用者を待つ。期限(5分。試験は短く)で強制で止める。止める前に控えを書き、
                                次の起動で一度だけ知らせる
    お願いが出ているあいだ      起動しない(待機画面に理由を出して、しばらくして終わる)
    stop.bat                    止めずに理由を返す
    お願いを出す                保守 →「管理者以外停止」。管理者パスワード(鍵)かつ Administrator

本物の子プロセス(`start_app.main(["--no-browser"])`)で確かめる。時間は `ALLTOOLS_REMOTE_STOP_TIMES` で短くし、
「残っているもの」は試験用の `ALLTOOLS_REMOTE_STOP_TEST_LEFT` で作る。
"""
from __future__ import annotations

import io
import json
import os
import socket
import subprocess
import sys
import tempfile
import threading
import time
import unittest
import urllib.request
from contextlib import redirect_stdout
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import remote_stop  # noqa: E402
from ufkensa import remote_stop_link as link  # noqa: E402

_LOCAL = urllib.request.build_opener(urllib.request.ProxyHandler({}))
TIMES = "grace=4,notice=0.5,poll=0.2,flush=0.5,block=4,beat=1"
CHILD = """
import sys
sys.path.insert(0, {root!r})
from ufkensa import app_config
app_config.port_candidates = lambda mode="main": [{port}]
import start_app
sys.exit(start_app.main(["--no-browser"]))
"""


class _Saved(unittest.TestCase):
    def setUp(self) -> None:
        saved = remote_stop.Setup(**vars(remote_stop.setup()))
        self.addCleanup(lambda: remote_stop.configure(**vars(saved)))
        remote_stop.reset()
        self.addCleanup(remote_stop.reset)
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.tmp = Path(tmp.name)
        self.share = self.tmp / "共有"
        self.share.mkdir()


class SameAsAllToolsTests(unittest.TestCase):
    def test_写しであることを書き_中身は変えていない(self) -> None:
        text = (ROOT / "remote_stop.py").read_text(encoding="utf-8")
        self.assertIn("all-tools", text.split('"""')[1])
        self.assertIn("中身を変えずに", text)
        for name in ("def configure", "def install", "def gate", "def watch", "def save_keep", "class Watcher"):
            self.assertIn(name, text)


class LinkTests(_Saved):
    def test_置き場所は使用パス_空なら検査履歴(self) -> None:
        from ufkensa.kensa import paths
        with mock.patch.dict(os.environ, {remote_stop.ROOT_ENV: ""}):
            with mock.patch.object(paths, "get", side_effect=lambda k, m=None: {"remote_stop": "", "history": str(self.share)}[k]):
                self.assertEqual(link.shared_root(), self.share)
            other = self.tmp / "別"
            with mock.patch.object(paths, "get", side_effect=lambda k, m=None: {"remote_stop": str(other), "history": "x"}[k]):
                self.assertEqual(link.shared_root(), other)
            with mock.patch.object(paths, "get", return_value=""):
                with self.assertRaises(remote_stop.Unreadable):
                    link.shared_root()
        self.assertIn("remote_stop", paths.BY_KEY)

    def test_このツールの名前と手元の領域(self) -> None:
        rs = link.configured()
        self.assertEqual((rs.setup().app, rs.setup().app_id, rs.setup().hosted), ("UF検査判定", "ufkensa", False))
        from ufkensa import app_config
        self.assertEqual(rs.setup().local_root, app_config.local_root())

    def test_残っているもの_操作の途中_印刷の確認_窓(self) -> None:
        from ufkensa.kensa import session as ks
        s = SimpleNamespace(_lock=threading.RLock(), state=SimpleNamespace(lot="H9022K1"), print_ready=True,
                            _print_pending=None, confi={"x": 1}, pack=None)
        with mock.patch.object(ks, "_session", s):
            self.assertEqual(link.left(), ["印刷・保存がまだ済んでいません(ロット H9022K1)", "公差外の確認の窓(ロット H9022K1)"])
            held = threading.Event()
            done = threading.Event()

            def hold():
                with s._lock:
                    held.set()
                    done.wait(5)
            t = threading.Thread(target=hold)
            t.start()
            held.wait(5)
            try:
                self.assertIn("検査の操作の途中", link.left())
            finally:
                done.set()
                t.join()
            kept = link.keep()
            self.assertTrue(any("H9022K1" in k and "手元DB" in k for k in kept), kept)
            self.assertTrue(any("印刷・保存" in k for k in kept), kept)
            s.print_ready, s.confi = False, None
            self.assertEqual(link.left(), [])
            self.assertEqual(link.keep(), [])
        with mock.patch.object(ks, "_session", None):
            self.assertEqual(link.left(), [])

    def test_打ちかけはお願いを受けているときだけ控えに書く(self) -> None:
        remote_stop.configure(local_root=self.tmp / "手元", root=lambda: self.share)
        self.assertEqual(link.save_drafts("/kensa", [{"label": "ロットNo", "value": "H1"}]), 0)
        w = mock.Mock(request={"id": "r1", "by_pc": "ADMIN", "created_at": 0})
        with mock.patch.object(remote_stop, "_watcher", w), \
                mock.patch.object(link, "configured", return_value=remote_stop):
            self.assertEqual(link.save_drafts("/kensa", [{"label": "ロットNo", "value": "H1"}, {"label": "空", "value": " "}]), 1)
        keeps = list((self.tmp / "手元" / "data" / "停止控え").glob("keep-*.json"))
        self.assertEqual(len(keeps), 1)
        data = json.loads(keeps[0].read_text(encoding="utf-8"))
        self.assertEqual(data["kept"], ["入力しかけ「ロットNo」: H1"])
        notes = remote_stop.restore_keeps()
        self.assertIn("入力しかけ「ロットNo」: H1", notes[0])


class AdminRouteTests(_Saved):
    """保守 →「管理者以外停止」の API。鍵(パスワード)かつ Administrator。"""

    def setUp(self) -> None:
        super().setUp()
        from app import create_app
        self.app = create_app("main", token="t")
        self.client = self.app.test_client()
        env = mock.patch.dict(os.environ, {remote_stop.ROOT_ENV: str(self.share), "ALLTOOLS_PC_NAME": "PC-ADMIN"})
        env.start()
        self.addCleanup(env.stop)

    def call(self, method: str, url: str, body=None):
        fn = getattr(self.client, method)
        return fn(url, json=body, headers={"X-Tool-Token": "t"})

    def who(self, password: bool, admin: bool):
        return mock.patch.object(link, "who", return_value={"password_ok": password, "administrator": admin})

    def test_鍵が閉じていれば_locked(self) -> None:
        with self.who(False, True):
            r = self.call("post", "/api/remote_stop/issue", {"reason": "更新"})
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "locked"))
        self.assertFalse((self.share / "管理者以外停止" / "お願い.json").exists())

    def test_Administratorでなければ断る(self) -> None:
        with self.who(True, False):
            r = self.call("post", "/api/remote_stop/issue", {"reason": "更新"})
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "not_administrator"))

    def test_出して_様子を見て_解除する(self) -> None:
        with self.who(True, True):
            r = self.call("post", "/api/remote_stop/issue", {"reason": "版上げ"})
            self.assertEqual(r.status_code, 200, r.get_json())
            self.assertTrue(r.get_json()["state"]["request"]["active"])
            again = self.call("post", "/api/remote_stop/issue", {})
            self.assertEqual(again.get_json()["error"]["code"], "already_active")
            st = self.call("get", "/api/remote_stop/state").get_json()
            self.assertIn("版上げ", st["request_text"])
            r = self.call("post", "/api/remote_stop/clear", {})
            self.assertEqual(r.status_code, 200)
            self.assertIsNone(r.get_json()["state"]["request"])

    def test_本当の判定_設定の途中はパスワードだけでは出せない(self) -> None:
        from ufkensa import auth
        with mock.patch.object(auth, "master_unlocked", return_value=True), \
                mock.patch.object(auth, "level", return_value=1):
            self.assertEqual(link.who(), {"password_ok": True, "administrator": False})
        with mock.patch.object(auth, "master_unlocked", return_value=True), \
                mock.patch.object(auth, "level", return_value=auth.EDIT_LEVEL):
            self.assertEqual(link.who(), {"password_ok": True, "administrator": True})


class StopEntryTests(_Saved):
    def run_main(self, running: bool):
        import process_manager
        out = io.StringIO()
        with mock.patch.dict(os.environ, {remote_stop.ROOT_ENV: str(self.share), "ALLTOOLS_PC_NAME": "PC-ADMIN"}), \
                redirect_stdout(out), \
                mock.patch("launch_guard.read_lock", return_value=object() if running else None), \
                mock.patch("launch_guard.desktop_running", return_value=None), \
                mock.patch.object(process_manager, "stop") as stop:
            stop.return_value = mock.Mock(stopped=True)
            code = process_manager.main([])
        return code, out.getvalue(), stop

    def test_お願いが出ていて動いていれば止めない(self) -> None:
        with mock.patch.dict(os.environ, {remote_stop.ROOT_ENV: str(self.share), "ALLTOOLS_PC_NAME": "PC-ADMIN"}):
            self.assertTrue(link.configured().issue(True, True, "更新").ok)
        code, out, stop = self.run_main(True)
        self.assertEqual(code, 1)
        self.assertTrue(out.startswith("止めませんでした: 管理者以外停止モード中"), out)
        stop.assert_not_called()
        _code, _out, stop = self.run_main(False)
        stop.assert_called()


class RealProcessTests(_Saved):
    """本物の子プロセス。"""

    def setUp(self) -> None:
        super().setUp()
        config = self.tmp / "user_config.json"
        config.write_text(json.dumps({"online": False, "local_dir": str(self.tmp), "db_file": "無い.sqlite3",
                                      "access_log": False}), encoding="utf-8")
        self.local = self.tmp / "local"
        self.left_file = self.tmp / "残っているもの.txt"
        self.env = dict(os.environ, UFKENSA_CONFIG_PATH=str(config), UFKENSA_LOG_DIR=str(self.tmp / "logs"),
                        UFKENSA_LOCAL_DIR=str(self.local), ALLTOOLS_REMOTE_STOP_DIR=str(self.share),
                        ALLTOOLS_PC_NAME="PC-LINE1", ALLTOOLS_LOGIN_ID="genba", ALLTOOLS_REMOTE_STOP_TIMES=TIMES,
                        ALLTOOLS_REMOTE_STOP_TEST_LEFT=str(self.left_file), PYTHONIOENCODING="utf-8")
        self.lock = self.local / "runtime" / "main.lock"
        remote_stop.configure(root=lambda: self.share, local_root=self.tmp / "管理者の手元", app="UF検査判定")

    def start(self) -> subprocess.Popen:
        with socket.socket() as sock:
            sock.bind(("127.0.0.1", 0))
            free = sock.getsockname()[1]
        try:
            self.lock.unlink()
        except OSError:
            pass
        proc = subprocess.Popen([sys.executable, "-c", CHILD.format(root=str(ROOT), port=free)], env=self.env,
                                cwd=str(ROOT), stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        self.addCleanup(lambda: (proc.poll() is None and proc.kill(), proc.wait(timeout=10), proc.stdout.close()))
        deadline = time.monotonic() + 60
        while time.monotonic() < deadline and proc.poll() is None:
            try:
                info = json.loads(self.lock.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                info = {}
            if info.get("port") and info.get("token") and info.get("pid") == proc.pid:
                self.port = int(info["port"])
                return proc
            time.sleep(0.05)
        self.fail("待ち受けが始まらない: " + (proc.stdout.read().decode("utf-8", "replace")[-2000:]
                                              if proc.poll() is not None else ""))

    def health(self, query: str = "") -> dict:
        with _LOCAL.open(f"http://127.0.0.1:{self.port}/api/health{query}", timeout=5) as r:
            return json.loads(r.read().decode("utf-8"))

    def wait_ready(self, proc) -> None:
        deadline = time.monotonic() + 90
        while time.monotonic() < deadline:
            try:
                if self.health().get("remote_stop") is not None:
                    return
            except (OSError, ValueError):
                pass
            if proc.poll() is not None:
                self.fail("起動の途中で終わりました: " + proc.stdout.read().decode("utf-8", "replace")[-2000:])
            time.sleep(0.2)
        self.fail("見張りが立ちません")

    def issue(self) -> None:
        with mock.patch.dict(os.environ, {"ALLTOOLS_PC_NAME": "PC-ADMIN", "ALLTOOLS_LOGIN_ID": "kanri"}), \
                mock.patch.object(remote_stop, "GRACE_SEC", 4.0):
            self.assertTrue(remote_stop.issue(True, True, "更新のため").ok)

    def clear(self) -> None:
        with mock.patch.dict(os.environ, {"ALLTOOLS_PC_NAME": "PC-ADMIN"}):
            self.assertTrue(remote_stop.clear(True, True).ok)

    def answer(self) -> dict:
        found = [json.loads(p.read_text(encoding="utf-8"))
                 for p in (self.share / "管理者以外停止" / "稼働中").glob("*.json")]
        mine = [a for a in found if a.get("pc") == "PC-LINE1"]
        self.assertEqual(len(mine), 1, found)
        return mine[0]

    def test_お願いで止まる_何も残っていなければまもなく(self) -> None:
        proc = self.start()
        self.wait_ready(proc)
        self.issue()
        self.assertEqual(proc.wait(40), 0, "お願いで止まらない")
        answer = self.answer()
        self.assertEqual((answer["state"], answer["app"]), (remote_stop.STATE_STOPPED, "UF検査判定"))

    def test_残っていれば待ち_期限で強制で止め_控えを次の起動で知らせる(self) -> None:
        self.left_file.write_text("試験の打ちかけ\n", encoding="utf-8")
        proc = self.start()
        self.wait_ready(proc)
        self.issue()
        time.sleep(1.5)
        state = self.health("?remote_stop=1&page=t")["remote_stop"]["state"]
        self.assertEqual(state["left"], ["試験の打ちかけ"], "帯に残っているものを出す")
        self.assertIsNone(proc.poll(), "残っているのに止まった")
        self.assertEqual(proc.wait(40), 0)
        answer = self.answer()
        self.assertEqual(answer["state"], remote_stop.STATE_FORCED)
        self.assertIn("試験の打ちかけ", answer["detail"])
        keeps = list((self.local / "data" / "停止控え").glob("keep-*.json"))
        self.assertEqual(len(keeps), 1, "この PC の手元の控え")
        self.assertFalse(any(self.share.rglob("keep-*")), "共有には書かない")

        self.clear()
        self.left_file.unlink()
        again = self.start()
        self.wait_ready(again)
        notes = self.health("?remote_stop=1&page=t")["remote_stop"]["notes"]
        self.assertEqual(len(notes), 1, notes)
        self.assertIn("強制で止めました", notes[0])
        self.assertIn("試験の打ちかけ", notes[0])
        self.assertEqual(self.health("?remote_stop=1&page=t")["remote_stop"]["notes"], [], "一度だけ")

    def test_お願い中は起動せず理由を出して終わる_解除すれば起動できる(self) -> None:
        self.issue()
        proc = self.start()
        deadline = time.monotonic() + 30
        why = ""
        while time.monotonic() < deadline and "管理者" not in why:
            try:
                why = str(self.health().get("startup_error") or "")
            except OSError:
                pass
            time.sleep(0.2)
        self.assertIn("更新が終わるまで起動できません", why)
        self.assertEqual(proc.wait(40), 0, "しばらくして終わる")
        self.clear()
        again = self.start()
        self.wait_ready(again)
        self.assertIsNone(again.poll())


class BandScriptTests(unittest.TestCase):
    def test_帯と打ちかけの部品をどの画面にも置く(self) -> None:
        base = (ROOT / "app" / "templates" / "base.html").read_text(encoding="utf-8")
        self.assertIn("js/remote_stop_band.js", base)
        self.assertIn("js/remote_stop_page.js", base)
        page = (ROOT / "app" / "static" / "js" / "remote_stop_page.js").read_text(encoding="utf-8")
        self.assertIn("__remoteStopUnsaved", page)
        self.assertIn("/api/remote_stop/drafts", page)
        maint = (ROOT / "app" / "templates" / "maint.html").read_text(encoding="utf-8")
        self.assertIn('data-tab="rstop"', maint)


if __name__ == "__main__":
    unittest.main()
