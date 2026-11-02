"""VBA との突き合わせ: 保存用・耳用に残っている公差と、Python で登録し直した公差を比べる(工程 10)。

ラインPCで、xlsm の取り込み(取り込み画面)が済んだあとに実行する:

    python scripts\\compare_saved_tolerance.py            (ロットごとに新しい行だけ)
    python scripts\\compare_saved_tolerance.py --all      (全部の行)

本番の手元DB・測定DB・共有フォルダには書かない。結果は ログの出力パス の
`ufkensa\\ufkensa_突き合わせ_<PC名>_<日付>.csv` に書き、件数を画面に出す。
"""
from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ufkensa import config, logging_utils, trace  # noqa: E402
from ufkensa.kensa import store as st, verify_saved  # noqa: E402


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--all", action="store_true", help="同じロットの古い行も比べる")
    ap.add_argument("--out", help="結果の CSV の置き場所(既定はログの出力パス)")
    a = ap.parse_args(argv)
    logging_utils.apply_settings()
    local = st.LocalStore()
    try:
        rows = verify_saved.saved_rows(local, latest_only=not a.all)
    finally:
        local.close()
    if not rows:
        print("保存用・耳用に行がありません(先に取り込み画面で xlsm を取り込んでください)")
        return 1
    with trace.timed("verify.saved", "保存用の公差と Python の公差を比べる", rows=len(rows)):
        outs = verify_saved.compare(rows)
    out = Path(a.out) if a.out else config.LOG_DIR / f"ufkensa_突き合わせ_{trace.pc_name()}_{date.today():%Y%m%d}.csv"
    verify_saved.write_report(outs, out)
    for k, v in verify_saved.summary(outs).items():
        print(f"{k}: {v} 件")
    print(f"結果: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
