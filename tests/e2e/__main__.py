"""ブラウザの通しの試験を流す入口

    python -m tests.e2e                    # 前提(node・Playwright・Chromium)を確かめてから流す
    python -m tests.e2e --shots <フォルダ>  # 画面の写しも残す(kensa_ops_1366x768_dark.png など)
    python -m tests.e2e --check            # 前提が揃っているかだけ見る(揃っていれば 0)

`UFKENSA_E2E=1 python -m pytest -q tests/e2e` と同じ試験を、pytest なしで(unittest で)流す。
前提が揃っていなければ、飛ばさずに理由を出して 2 で終わる(「流したつもりで飛んでいた」を防ぐ)。
"""
from __future__ import annotations

import argparse
import os
import sys
import unittest
from pathlib import Path


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="python -m tests.e2e", description="ブラウザの通しの試験(Playwright)を流す")
    ap.add_argument("--shots", metavar="DIR", help="画面の写しを置くフォルダ(UFKENSA_E2E_SHOTS と同じ)")
    ap.add_argument("--check", action="store_true", help="前提が揃っているかだけ見る")
    args = ap.parse_args(argv)

    os.environ["UFKENSA_E2E"] = "1"                 # 試験の読み込みより先に
    if args.shots:
        os.environ["UFKENSA_E2E_SHOTS"] = str(Path(args.shots).resolve())
    from tests.e2e import test_e2e

    node = test_e2e._node()
    why = "node がありません(PATH)" if not node else test_e2e._playwright_ok(node)
    if why:
        print("ブラウザの通しの試験を流せません: " + why, file=sys.stderr)
        print("node の playwright を require できるように(NODE_PATH など)し、`npx playwright install chromium` を済ませてください",
              file=sys.stderr)
        return 2
    if args.check:
        print("前提は揃っています: node " + node)
        return 0
    suite = unittest.defaultTestLoader.loadTestsFromModule(test_e2e)
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    if result.skipped:
        return 2
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    sys.exit(main())
