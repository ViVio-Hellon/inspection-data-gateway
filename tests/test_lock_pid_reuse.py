"""2026-11-08: 再起動で PID が使い回されても「起動中」と誤らない・stop.bat が片付ける。"""
from __future__ import annotations

import os
import subprocess
import sys
import time
import unittest
from unittest import mock

import launch_guard
import process_manager


def _info(pid: int, started_at: float, port: int = 1) -> launch_guard.LockInfo:
    return launch_guard.LockInfo(app_id="x", mode="main", pid=pid, port=port, url="", started_at=started_at)


class StartTimeTest(unittest.TestCase):
    def test_own_process_is_owner(self) -> None:
        st = launch_guard.process_start_time(os.getpid())
        if st is None:
            self.skipTest("起動時刻が取れない環境")
        self.assertLess(abs(st - time.time()), 24 * 3600)
        self.assertTrue(launch_guard.is_lock_owner_alive(_info(os.getpid(), time.time())))

    def test_process_started_after_lock_is_not_owner(self) -> None:
        if launch_guard.process_start_time(os.getpid()) is None:
            self.skipTest("起動時刻が取れない環境")
        # ロックは 1 時間前に書かれた = いまのプロセスは後から同じ番号を持った別物
        self.assertFalse(launch_guard.is_lock_owner_alive(_info(os.getpid(), time.time() - 3600)))

    def test_dead_pid(self) -> None:
        p = subprocess.Popen([sys.executable, "-c", "pass"])
        p.wait()
        self.assertFalse(launch_guard.is_lock_owner_alive(_info(p.pid, time.time())))


class CheckExistingTest(unittest.TestCase):
    def test_reused_pid_lock_is_cleaned(self) -> None:
        info = _info(os.getpid(), time.time() - 3600, port=9)
        with mock.patch.object(launch_guard, "read_lock", return_value=info), \
                mock.patch.object(launch_guard, "is_lock_owner_alive", return_value=False), \
                mock.patch.object(launch_guard, "remove_lock") as rm:
            r = launch_guard.check_existing("main")
        self.assertTrue(r.should_start)
        rm.assert_called_once()

    def test_hung_own_app_is_killed_after_wait(self) -> None:
        info = _info(4242, time.time(), port=9)
        with mock.patch.object(launch_guard, "read_lock", return_value=info), \
                mock.patch.object(launch_guard, "is_lock_owner_alive", return_value=True), \
                mock.patch.object(launch_guard, "probe_health", return_value=None), \
                mock.patch.object(launch_guard, "HUNG_WAIT_SEC", 0.6), \
                mock.patch.object(launch_guard, "_kill_if_ours", return_value=True) as kill, \
                mock.patch.object(launch_guard, "remove_lock") as rm:
            r = launch_guard.check_existing("main")
        self.assertTrue(r.should_start)
        kill.assert_called_once()
        rm.assert_called_once()

    def test_not_our_process_is_never_killed(self) -> None:
        info = _info(4242, time.time())
        with mock.patch.object(launch_guard, "is_lock_owner_alive", return_value=True), \
                mock.patch.object(launch_guard, "looks_like_our_process", return_value=False), \
                mock.patch("os.kill") as k:
            self.assertFalse(launch_guard._kill_if_ours(info))
        k.assert_not_called()


class StopTest(unittest.TestCase):
    def test_stop_cleans_lock_of_reused_pid(self) -> None:
        info = _info(os.getpid(), time.time() - 3600, port=9)
        with mock.patch.object(launch_guard, "read_lock", return_value=info), \
                mock.patch.object(launch_guard, "probe_health", return_value=None), \
                mock.patch.object(launch_guard, "is_lock_owner_alive", return_value=False), \
                mock.patch.object(launch_guard, "remove_lock") as rm, mock.patch("os.kill") as k:
            r = process_manager.stop("main")
        self.assertTrue(r.stopped)
        self.assertIn("片付けました。Start.vbs で起動できます", r.message)
        rm.assert_called_once()
        k.assert_not_called()


if __name__ == "__main__":
    unittest.main()
