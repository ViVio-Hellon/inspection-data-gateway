"""テスト履歴(ufkensa/kensa/data/test_history.json)を作る。

試験と同じ合成データ(tests/fixtures/kensa。名簿は伏せ字)で 1 ロットを最後まで検査・印刷し、
保存用・Lot情報履歴 の行と、ライン別のシート(WorkerTime・Daily)を書き出す。

    python tools/make_test_history.py
"""
from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import tests  # noqa: E402,F401 -- 試験と同じ隔離  試験と同じ隔離(本番の設定・共有に触れない)
from tests.test_kensa_session import LOT, SessionBase  # noqa: E402
from ufkensa.kensa import records  # noqa: E402
from ufkensa.kensa.store import HOZON, LOT_HISTORY  # noqa: E402

TEST_LOT = "T999T99"
OUT = ROOT / "ufkensa" / "kensa" / "data" / "test_history.json"


class Make(SessionBase):
    def runTest(self) -> None:  # noqa: N802
        s = self.s
        self.register()
        self.fill("DT")
        self.fill("#L")
        self.strain_visual_height()
        self.ok(s.min_max())
        self.ok(s.print_check())
        s.print_commit()
        self.ok(s.print_commit({"print": True}))
        lname = s.state.lname
        hz = self.store.sheet(HOZON)
        hist = self.store.sheet(LOT_HISTORY)
        data = {
            "_説明": "保守 → 履歴出力 の「テスト履歴を入れる」で使う 1 ロット分(試験用の合成データ。名簿は伏せ字)。"
                     "tools/make_test_history.py で作り直す",
            "lot": LOT,
            "rows": {
                HOZON: [r for _, r in hz.rows() if records.cell_text(r[1]) == LOT],
                LOT_HISTORY: [r for _, r in hist.rows() if LOT in {records.cell_text(v) for v in r[:3]}],
            },
            "line_sheets": {},
        }
        for base in ("WorkerTime", "Daily"):
            name = f"{base}_{lname}"
            if self.store.exists(name):
                data["line_sheets"][base] = {f"{r},{c}": v for r, c, v in self.store.sheet(name).used()}
        # 本物のロット・納入先と見分けがつくように置き換える
        text = json.dumps(data, ensure_ascii=False, indent=1).replace(LOT, TEST_LOT)
        for row in data["rows"][LOT_HISTORY]:
            for i in (11, 12, 13):
                if i < len(row) and row[i]:
                    text = text.replace(json.dumps(row[i], ensure_ascii=False), json.dumps("ﾃｽﾄﾉｳﾆﾕｳｻｷ", ensure_ascii=False))
        data = json.loads(text)
        OUT.parent.mkdir(parents=True, exist_ok=True)
        OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
        print("書きました:", OUT, {k: len(v) for k, v in data["rows"].items()}, list(data["line_sheets"]))


if __name__ == "__main__":
    r = unittest.TextTestRunner(verbosity=0).run(Make())
    sys.exit(0 if r.wasSuccessful() else 1)
