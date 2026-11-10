"""2026-11-10: 2回目の起動は、動いている1つ目を開くだけ(2つ目のサーバを立てない)。

現場の声「多重起動できます」。2回目の起動が1つ目を「応答なし」「死んだロック」と取り違えると、
ロックを消して2つ目を立てていた。
"""
from __future__ import annotations

import json
import os
import socket
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

import launch_guard
import start_app


def _info(pid: int = 4242, port: int = 9, started_at: float | None = None) -> launch_guard.LockInfo:
    return launch_guard.LockInfo(app_id="x", mode="main", pid=pid, port=port,
                                 url=f"http://127.0.0.1:{port}/",
                                 started_at=time.time() if started_at is None else started_at)


def _ours(pid: int = 4242, port: int = 9, version: str | None = None) -> dict:
    from ufkensa import app_config
    return {"app_id": app_config.app_id(), "mode": "main", "pid": pid, "port": port,
            "version": version or app_config.version()}


class _LocalDir(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        patcher = mock.patch.dict(os.environ, {"UFKENSA_LOCAL_DIR": self.tmp.name})
        patcher.start()
        self.addCleanup(patcher.stop)
        self.addCleanup(self.tmp.cleanup)


class BusyServerIsJoinedTest(_LocalDir):
    """待ち受けてはいるが `/api/health` が返らない(検査の途中で忙しい)。"""

    def setUp(self) -> None:
        super().setUp()
        # accept しない待ち受け = TCP はつながるが HTTP は返らない
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.bind(("127.0.0.1", 0))
        self.sock.listen(16)
        self.addCleanup(self.sock.close)
        self.port = self.sock.getsockname()[1]

    def test_busy_server_that_cannot_be_verified_is_joined_not_duplicated(self) -> None:
        info = _info(port=self.port)
        with mock.patch.object(launch_guard, "read_lock", return_value=info), \
                mock.patch.object(launch_guard, "is_lock_owner_alive", return_value=True), \
                mock.patch.object(launch_guard, "HUNG_WAIT_SEC", 0.2), \
                mock.patch.object(launch_guard, "BUSY_WAIT_SEC", 0.5), \
                mock.patch.object(launch_guard, "HEALTH_TIMEOUT_SEC", 0.1), \
                mock.patch.object(launch_guard, "looks_like_our_process", return_value=False), \
                mock.patch.object(launch_guard, "remove_lock") as rm, \
                mock.patch("os.kill") as k:
            r = launch_guard.check_existing("main")
        self.assertFalse(r.should_start, r.reason)
        self.assertEqual(r.url, info.url)
        rm.assert_not_called()
        k.assert_not_called()

    def test_busy_server_that_answers_later_is_joined(self) -> None:
        info = _info(port=self.port)
        answers = iter([None] * 3)

        def probe(port, timeout=1.5):
            return next(answers, _ours(port=self.port))

        with mock.patch.object(launch_guard, "read_lock", return_value=info), \
                mock.patch.object(launch_guard, "is_lock_owner_alive", return_value=True), \
                mock.patch.object(launch_guard, "HUNG_WAIT_SEC", 0.2), \
                mock.patch.object(launch_guard, "BUSY_WAIT_SEC", 10), \
                mock.patch.object(launch_guard, "probe_health", side_effect=probe), \
                mock.patch.object(launch_guard, "_kill_if_ours") as kill, \
                mock.patch.object(launch_guard, "remove_lock") as rm:
            r = launch_guard.check_existing("main")
        self.assertFalse(r.should_start)
        kill.assert_not_called()
        rm.assert_not_called()

    def test_server_that_stopped_listening_is_still_replaced(self) -> None:
        """待ち受けていない(本当に応答なし)なら、今までどおり片付けて起動する。"""
        self.sock.close()
        info = _info(port=self.port)
        with mock.patch.object(launch_guard, "read_lock", return_value=info), \
                mock.patch.object(launch_guard, "is_lock_owner_alive", return_value=True), \
                mock.patch.object(launch_guard, "HUNG_WAIT_SEC", 0.2), \
                mock.patch.object(launch_guard, "_kill_if_ours", return_value=True), \
                mock.patch.object(launch_guard, "remove_lock") as rm:
            r = launch_guard.check_existing("main")
        self.assertTrue(r.should_start)
        rm.assert_called_once()


class StartTimeMisjudgedTest(_LocalDir):
    def test_owner_that_answers_with_same_pid_is_alive(self) -> None:
        """起動時刻の比べが外れても、本人が同じ PID で答えれば合流する(Windows で起きうる)。"""
        info = _info(pid=777, port=8800)
        with mock.patch.object(launch_guard, "read_lock", return_value=info), \
                mock.patch.object(launch_guard, "is_lock_owner_alive", return_value=False), \
                mock.patch.object(launch_guard, "probe_health", return_value=_ours(pid=777, port=8800)), \
                mock.patch.object(launch_guard, "remove_lock") as rm:
            r = launch_guard.check_existing("main")
        self.assertFalse(r.should_start)
        self.assertEqual(r.url, info.url)
        rm.assert_not_called()

    def test_other_pid_on_port_is_not_trusted(self) -> None:
        info = _info(pid=777, port=8800)
        with mock.patch.object(launch_guard, "read_lock", return_value=info), \
                mock.patch.object(launch_guard, "is_lock_owner_alive", return_value=False), \
                mock.patch.object(launch_guard, "probe_health", return_value=_ours(pid=778, port=8800)), \
                mock.patch.object(launch_guard, "remove_lock") as rm:
            r = launch_guard.check_existing("main")
        self.assertTrue(r.should_start)
        rm.assert_called_once()


class RemoveOnlyOwnLockTest(_LocalDir):
    def test_exiting_instance_does_not_remove_newer_lock(self) -> None:
        launch_guard.write_lock(_info(pid=os.getpid() + 1, port=8800))
        launch_guard.remove_lock("main", only_pid=os.getpid())
        self.assertIsNotNone(launch_guard.read_lock("main"))
        launch_guard.remove_lock("main", only_pid=os.getpid() + 1)
        self.assertIsNone(launch_guard.read_lock("main"))


class AppLockFallbackTest(_LocalDir):
    """ロックが消えていても、同時起動の錠をブラウザ版が握っていれば、動いているほうを開く。"""

    def tearDown(self) -> None:
        launch_guard.release_app_lock()

    def test_second_launch_joins_via_app_lock(self) -> None:
        runtime = Path(self.tmp.name) / "runtime"
        runtime.mkdir(parents=True)
        (runtime / launch_guard.APP_OWNER).write_text(json.dumps(
            {"kind": launch_guard.KIND_BROWSER, "pid": 99999, "started_at": time.time()}), encoding="utf-8")
        opened = []
        with mock.patch.object(launch_guard, "hold_app_lock", return_value=None), \
                mock.patch.object(launch_guard, "find_running", return_value=_ours(pid=99999, port=8811)), \
                mock.patch.object(start_app, "log_environment"), \
                mock.patch("webbrowser.open", side_effect=opened.append):
            code = start_app.start("main", open_browser=True)
        self.assertEqual(code, 0)
        self.assertEqual(opened, ["http://127.0.0.1:8811/"])
        lock = launch_guard.read_lock("main")
        self.assertIsNotNone(lock)
        self.assertEqual((lock.pid, lock.port), (99999, 8811))

    def test_desktop_owner_still_refuses(self) -> None:
        runtime = Path(self.tmp.name) / "runtime"
        runtime.mkdir(parents=True)
        (runtime / launch_guard.APP_OWNER).write_text(json.dumps(
            {"kind": launch_guard.KIND_DESKTOP, "pid": 99999}), encoding="utf-8")
        with mock.patch.object(launch_guard, "hold_app_lock", return_value=None), \
                mock.patch.object(start_app, "log_environment"), \
                mock.patch("webbrowser.open"):
            with self.assertRaises(start_app.StartupError):
                start_app.start("main", open_browser=True)
        self.assertIsNone(launch_guard.read_lock("main"))


if __name__ == "__main__":
    unittest.main()
