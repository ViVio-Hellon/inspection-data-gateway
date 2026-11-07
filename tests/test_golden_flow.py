"""第3段(案C)の安全網: 検査の流れの途中ごとに、画面・手元DB の全シート・帳票を記録して、作り直しの前後で突き合わせる。

記録は tests/golden/*.json。作り直しの前のコードで作り、以後は一致しなければ失敗する。
業務の動きを意図して変えたときだけ、`UFKENSA_GOLDEN_UPDATE=1 python -m pytest tests/test_golden_flow.py` で作り直す。

記録するもの(各チェックポイントで)
- 手元DB の全シートのセル(保存用・作業保存・印刷用・不良要因印刷・WorkerTime・Daily・Lot情報履歴 …)
- 画面の値・表示・NG(checks.Uf1 の写し)と、画面に返す view
- 保存・帳票に渡す KensaSnapshot
"""
from __future__ import annotations

import dataclasses
import json
import os
import re
import unittest
from pathlib import Path
from typing import Any

from tests.test_kensa_session import SessionBase
from ufkensa.kensa import wiring

GOLDEN = Path(__file__).with_name("golden")
UPDATE = os.environ.get("UFKENSA_GOLDEN_UPDATE") == "1"
_TMP = re.compile(r"/tmp/[^\s\"']*?/|[A-Za-z]:\\\\[^\"']*?\\\\")


def _plain(v: Any) -> Any:
    if dataclasses.is_dataclass(v) and not isinstance(v, type):
        return {f.name: _plain(getattr(v, f.name)) for f in dataclasses.fields(v)}
    if isinstance(v, dict):
        return {str(k): _plain(x) for k, x in sorted(v.items(), key=lambda kv: str(kv[0]))}
    if isinstance(v, (list, tuple)):
        return [_plain(x) for x in v]
    if isinstance(v, set):
        return sorted(_plain(x) for x in v)
    if isinstance(v, float):
        return round(v, 9)
    if isinstance(v, Path):
        return v.name
    if isinstance(v, str):
        return _TMP.sub("<tmp>/", v)
    return v


class GoldenFlowTest(SessionBase):
    maxDiff = None

    def snap(self, label: str) -> None:
        s, st = self.s, self.s.state
        sheets = {}
        for name in sorted(self.store.sheet_names()):
            sheets[name] = {f"{r},{c}": v for r, c, v in self.store.sheet(name).used()}
        # 帳票(成績票・不良要因票)は手元DB に残らないので、書いた HTML の中身を記録する
        for key in ("report", "defect"):
            f = (getattr(s, "report", {}) or {}).get(key) or {}
            if f.get("html") and Path(f["html"]).is_file():
                sheets["帳票:" + key] = {"html": Path(f["html"]).read_text(encoding="utf-8")}
        uf1 = wiring.to_uf1(st)
        view = s._view()
        view.pop("server_time", None)
        self.points[label] = _plain({
            "sheets": sheets,
            "uf1": {"text": uf1.text, "visible": uf1.visible, "ng": uf1.ng, "work": uf1.work},
            "snapshot": wiring.to_snapshot(st, self.store),
            "view": view,
        })

    def check(self, name: str) -> None:
        path = GOLDEN / f"{name}.json"
        text = json.dumps(self.points, ensure_ascii=False, indent=0, sort_keys=True, default=str)
        if UPDATE or not path.exists():
            path.write_text(text, encoding="utf-8")
            if not UPDATE:
                self.skipTest(f"記録を作りました: {path.name}(次から突き合わせます)")
            return
        want = json.loads(path.read_text(encoding="utf-8"))
        got = json.loads(text)
        for label in want:
            for part in want[label]:
                with self.subTest(point=label, part=part):
                    self.assertEqual(got.get(label, {}).get(part), want[label][part])
        self.assertEqual(list(got), list(want))

    def setUp(self) -> None:
        super().setUp()
        self.points: dict[str, Any] = {}

    # --- 1 ロットを最後まで(厚板・A 面) ----------------------------------
    def test_one_lot(self) -> None:
        s = self.s
        self.register()
        self.snap("1_登録")
        self.fill("DT")
        self.snap("2_厚")
        self.fill("#L")
        self.snap("3_幅丈対角")
        self.strain_visual_height()
        self.snap("4_歪外観高さ")
        self.ok(s.min_max())
        self.ok(s.print_check())
        s.print_commit()
        self.ok(s.print_commit({"print": True}))
        self.snap("5_印刷保存")
        s2 = self.new_session()
        self.ok(s2.start())
        self.s = s2
        self.snap("6_読み戻し")
        self.check("one_lot")

    # --- 公差外・許容・再測定 ---------------------------------------------
    def test_ng_flow(self) -> None:
        s = self.s
        self.register()
        s.receive("DT10001+0000015.500M")
        for _ in range(3):
            s.receive("DT10001+0000015.012M")
        self.snap("1_公差外")
        self.ok(s.set_value("UFConfi.TB1", "許容"))
        self.snap("2_許容")
        s.receive("DT10001+0000015.600M")
        for _ in range(3):
            s.receive("DT10001+0000015.012M")
        self.ok(s.set_value("UFConfi.TB1", "再測定"))
        self.snap("3_再測定")
        self.check("ng_flow")

    # --- 歪の A 面 / B 面 ----------------------------------------------
    def test_strain_faces(self) -> None:
        s = self.s
        self.register()
        for n, v in (("CBS1", "幅"), ("CBS2", "上"), ("ZS", "0.3"), ("GS", "0.2")):
            self.ok(s.set_value(n, v))
        self.ok(s.set_value("AorB", ""))
        self.snap("1_B面へ")
        self.ok(s.set_value("ZS", "0.4"))
        self.ok(s.set_value("AorB", ""))
        self.snap("2_A面へ戻す")
        s2 = self.new_session()
        self.ok(s2.start())
        self.s = s2
        self.snap("3_読み戻し")
        self.check("strain_faces")


if __name__ == "__main__":
    unittest.main()
