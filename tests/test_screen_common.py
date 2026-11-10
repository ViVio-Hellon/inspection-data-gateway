"""画面の受け口の共通部分(app/routes/screen_common.py。整理 6)

帳票を開く・包装仕様書の画像・操作の受け口を、検査・キャスク・全ヵ所・Aインプット・校正で 1 つにした。
各画面の通しは test_screen_* が見ている。ここは共通部分の決まり(404・400・501・止め・型・max_age)を見る。
"""
from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

from flask import Flask

from app.routes import screen_common
from ufkensa.kensa.flow import Message, Stop


class _Result:
    def __init__(self, d: dict) -> None:
        self.d = d

    def to_dict(self) -> dict:
        return self.d


class _Session:
    def __init__(self) -> None:
        self.calls: list = []

    def go(self, **kw):
        self.calls.append(kw)
        return _Result({"view": {"ok": 1}, "messages": [], "question": None, "stop": None, "kw": kw})

    def todo(self, **kw):
        raise NotImplementedError

    def stop(self, **kw):
        raise Stop("止めました", messages=[Message("知らせ")], source="test")


OPS = {"go": ("go", ("text", "answers")), "todo": ("todo", ()), "stop": ("stop", (), "止める")}


class ScreenCommonTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="ufkensa_common_"))
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.s = _Session()
        app = Flask(__name__)
        opts: dict = {}

        @app.post("/op/<op>")
        def op(op: str):
            return screen_common.operate(op, OPS, lambda: self.s, area="t", label="試験", **opts)

        @app.get("/report/<path:name>")
        def report(name: str):
            return screen_common.report_file(self.tmp, name)

        @app.get("/img/<name>")
        def img(name: str):
            return screen_common.pack_image(None if name == "none" else self.tmp / name)

        self.opts = opts
        self.c = app.test_client()

    def test_unknown_op_404(self) -> None:
        r = self.c.post("/op/nothing", json={})
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (404, "not_found"))

    def test_passes_only_named_args(self) -> None:
        d = self.c.post("/op/go", json={"text": "DT1", "other": 1}).get_json()
        self.assertEqual(d["kw"], {"text": "DT1"})

    def test_missing_args_400_only_when_required(self) -> None:
        self.assertEqual(self.c.post("/op/go", json={}).status_code, 200)        # 既定は確かめない(検査・全ヵ所・A)
        self.opts["required"] = lambda n: n != "answers"                           # キャスク
        r = self.c.post("/op/go", json={})
        self.assertEqual(r.status_code, 400)
        self.assertEqual(r.get_json()["error"], {"code": "bad_value", "message": "足りない値があります: text"})
        self.assertEqual(self.c.post("/op/go", json={"text": "x"}).status_code, 200)

    def test_as_text(self) -> None:
        self.opts["as_text"] = True                                                # 校正
        d = self.c.post("/op/go", json={"text": 12}).get_json()
        self.assertEqual(d["kw"], {"text": "12"})

    def test_not_implemented_501(self) -> None:
        r = self.c.post("/op/todo", json={})
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (501, "not_ready"))

    def test_raw_stop_becomes_result(self) -> None:
        self.opts["stop_view"] = lambda s: {"v": 1}
        d = self.c.post("/op/stop", json={}).get_json()
        self.assertEqual(d["view"], {"v": 1})
        self.assertEqual(d["stop"]["text"], "止めました")
        self.assertEqual([m["text"] for m in d["messages"]], ["知らせ"])
        self.assertIsNone(d["question"])

    def test_report(self) -> None:
        (self.tmp / "印刷_A.html").write_text("<p>帳票</p>", encoding="utf-8")
        r = self.c.get("/report/印刷_A.html")
        self.assertEqual((r.status_code, r.mimetype), (200, "text/html"))
        self.assertIn("max-age=0", r.headers.get("Cache-Control", ""))
        r.close()
        self.assertEqual(self.c.get("/report/none.html").status_code, 404)
        for bad in ("x.txt", "x.HTML", "a.html.txt"):
            self.assertEqual(self.c.get(f"/report/{bad}").status_code, 400, bad)

    def test_pack_image_types(self) -> None:
        self.assertEqual(self.c.get("/img/none").status_code, 404)
        self.assertEqual(self.c.get("/img/nofile.jpg").status_code, 404)
        for name, mime in (("a.JPG", "image/jpeg"), ("b.png", "image/png"), ("c.bmp", "image/bmp"),
                           ("d.xyz", "application/octet-stream")):
            (self.tmp / name).write_bytes(b"\x00")
            r = self.c.get(f"/img/{name}")
            self.assertEqual((r.status_code, r.mimetype), (200, mime), name)
            r.close()


if __name__ == "__main__":
    unittest.main()
