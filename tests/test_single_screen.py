"""操作できる画面は1つだけ・2回目の起動は開いている画面を前に出す(2026-11-11)。

現場の声「まだ多重起動できる」── 起動をもう一度押すとブラウザの画面がもう1つ開き、どちらでも操作できた。
"""
from __future__ import annotations

import unittest
from unittest import mock

import launch_guard
import start_app
from tests._web import TOKEN, make_client
from ufkensa import screens


class _Clock:
    def __init__(self) -> None:
        self.t = 1000.0

    def __call__(self) -> float:
        return self.t


class RegistryTest(unittest.TestCase):
    def setUp(self) -> None:
        self.clock = _Clock()
        self.reg = screens.ScreenRegistry(self.clock)

    def test_first_screen_is_active_second_is_not(self) -> None:
        self.assertTrue(self.reg.ping("A")["active"])
        self.assertFalse(self.reg.ping("B")["active"])
        self.assertTrue(self.reg.may_write("A"))
        self.assertFalse(self.reg.may_write("B"))
        self.assertTrue(self.reg.may_write(""))           # 番号の無い要求は通す

    def test_claim_takes_over(self) -> None:
        self.reg.ping("A"), self.reg.ping("B")
        self.reg.claim("B")
        self.assertFalse(self.reg.ping("A")["active"])
        self.assertTrue(self.reg.may_write("B"))
        self.assertFalse(self.reg.may_write("A"))

    def test_closed_active_screen_frees_after_heartbeat_stops(self) -> None:
        self.reg.ping("A"), self.reg.ping("B")
        self.clock.t += screens.OPEN_SEC + 1
        self.assertTrue(self.reg.ping("B")["active"])

    def test_reload_hands_over_but_duplicate_does_not(self) -> None:
        self.reg.ping("A")
        self.reg.leaving("A")                              # 読み込み直し: pagehide
        self.reg.ping("X")                                 # 別のタブが割り込んでも
        self.assertFalse(self.reg.is_active("X"))          # 少しのあいだは取られない
        self.reg.touch("A2", prev="A")                     # 同じタブの次のページ
        self.assertTrue(self.reg.may_write("A2"))
        self.assertFalse(self.reg.may_write("X"))

    def test_leaving_frees_after_grace(self) -> None:
        self.reg.ping("A")
        self.reg.leaving("A")
        self.clock.t += screens.HANDOFF_SEC + 1
        self.assertTrue(self.reg.ping("B")["active"])

    def test_raise_only_when_a_visible_screen_is_open(self) -> None:
        self.assertFalse(self.reg.request_raise())         # 開いていない → ブラウザを開く
        before = self.reg.ping("A")["raise_seq"]
        self.assertTrue(self.reg.request_raise())
        self.assertEqual(self.reg.ping("A")["raise_seq"], before + 1)
        self.clock.t += screens.OPEN_SEC + 1               # 心拍が途絶えた(閉じた)
        self.assertFalse(self.reg.request_raise())

    def test_closed_screen_is_not_raised(self) -> None:
        self.reg.ping("A")
        self.reg.leaving("A")
        self.assertFalse(self.reg.request_raise())

    def test_hidden_screen_keeps_active_but_is_not_raised(self) -> None:
        self.reg.ping("A")
        self.reg.hidden("A")
        self.clock.t += screens.OPEN_SEC + 5
        self.assertFalse(self.reg.ping("B")["active"])     # 裏に回っても使っている画面のまま
        self.assertTrue(self.reg.is_active("A"))


class ServerRefusalTest(unittest.TestCase):
    def setUp(self) -> None:
        self.c = make_client()

    def h(self, sid: str) -> dict:
        return {"X-Tool-Token": TOKEN, "X-Screen-Id": sid}

    def test_inactive_screen_write_is_refused(self) -> None:
        self.assertTrue(self.c.get("/api/screen", headers=self.h("A")).get_json()["active"])
        self.assertFalse(self.c.get("/api/screen", headers=self.h("B")).get_json()["active"])
        r = self.c.post("/api/no-such-thing", headers=self.h("B"), json={})
        self.assertEqual(r.status_code, 409)
        self.assertEqual(r.get_json()["error"]["code"], "inactive_screen")
        self.assertIn("ほかのウィンドウ", r.get_json()["error"]["message"])
        # 使っている画面は通る(経路が無いので 404)
        self.assertEqual(self.c.post("/api/no-such-thing", headers=self.h("A"), json={}).status_code, 404)
        # 読むだけ・番号の無い要求は通す
        self.assertNotEqual(self.c.get("/api/health", headers=self.h("B")).status_code, 409)
        self.assertEqual(self.c.post("/api/no-such-thing", headers={"X-Tool-Token": TOKEN},
                                     json={}).status_code, 404)

    def test_claim_switches_and_old_screen_is_refused(self) -> None:
        self.c.get("/api/screen", headers=self.h("A"))
        self.c.get("/api/screen", headers=self.h("B"))
        self.assertEqual(self.c.post("/api/screen/claim", headers=self.h("B"), json={}).status_code, 200)
        self.assertFalse(self.c.get("/api/screen", headers=self.h("A")).get_json()["active"])
        self.assertEqual(self.c.post("/api/no-such-thing", headers=self.h("A"), json={}).status_code, 409)
        self.assertEqual(self.c.post("/api/no-such-thing", headers=self.h("B"), json={}).status_code, 404)

    def test_raise_needs_no_token_and_reports_open_screen(self) -> None:
        self.assertFalse(self.c.post("/api/screen/raise", json={}).get_json()["raised"])
        self.c.get("/api/screen", headers=self.h("A"))
        self.assertTrue(self.c.post("/api/screen/raise", json={}).get_json()["raised"])
        self.assertEqual(self.c.get("/api/screen", headers=self.h("A")).get_json()["raise_seq"], 1)

    def test_leaving_beacon_releases(self) -> None:
        self.c.get("/api/screen", headers=self.h("A"))
        self.c.post("/api/alive", json={"screen_id": "A", "leaving": True})
        self.assertFalse(self.c.post("/api/screen/raise", json={}).get_json()["raised"])

    def test_receiver_results_go_only_to_active_screen(self) -> None:
        try:
            from app.routes import kensa  # noqa: F401
        except Exception:                 # noqa: BLE001
            self.skipTest("検査の画面が無い")
        self.c.get("/api/screen", headers=self.h("A"))
        self.c.get("/api/screen", headers=self.h("B"))
        r = self.c.get("/api/kensa/recv?since=0&screen=calibration", headers=self.h("B"))
        if r.status_code == 404:
            self.skipTest("検査の画面が無い")
        self.assertTrue(r.get_json().get("inactive"))
        self.assertEqual(r.get_json()["results"], [])
        r = self.c.get("/api/kensa/recv?since=0&screen=calibration", headers=self.h("A"))
        self.assertFalse(r.get_json().get("inactive"))


class JoinDecisionTest(unittest.TestCase):
    def test_open_screen_is_raised_and_browser_not_opened(self) -> None:
        with mock.patch.object(launch_guard, "raise_open_screen", return_value=True) as raised, \
                mock.patch.object(start_app.webbrowser, "open") as opened:
            self.assertFalse(start_app.join_running("http://127.0.0.1:9/", 9, "tok"))
        raised.assert_called_once_with(9, "tok")
        opened.assert_not_called()

    def test_no_screen_opens_browser(self) -> None:
        with mock.patch.object(launch_guard, "raise_open_screen", return_value=False), \
                mock.patch.object(start_app.webbrowser, "open") as opened:
            self.assertTrue(start_app.join_running("http://127.0.0.1:9/", 9, "tok"))
        opened.assert_called_once_with("http://127.0.0.1:9/")

    def test_raise_against_real_server(self) -> None:
        import threading
        from werkzeug.serving import make_server
        from app import create_app
        app = create_app("main", token=TOKEN, port=0)
        app.config["READY"] = True
        srv = make_server("127.0.0.1", 0, app)
        threading.Thread(target=srv.serve_forever, daemon=True).start()
        self.addCleanup(srv.shutdown)
        port = srv.server_port
        # 画面が無い → 開く。画面がある → 前に出す
        self.assertFalse(launch_guard.raise_open_screen(port, TOKEN))
        screens.of(app).ping("A")
        self.assertTrue(launch_guard.raise_open_screen(port, TOKEN))
        self.assertFalse(launch_guard.raise_open_screen(1, TOKEN, timeout=0.5))   # 応答なし → 開く


if __name__ == "__main__":
    unittest.main()
