"""日報を all-tools の日報モード(日報管理ツール nippou)へ渡す CSV(2026-11-08 のご依頼)。

検査の日報(手元DB の `Daily_<L>`。VBA の Worker.Daily_Create / Daily_Report と同じシート)を、
日報モードの「CSV から日報を戻す」(nippou/logic/csv_import.py)がそのまま読める形で書き出す。

- 見出しは日報モードが書き出す「集計明細」の見出し(DETAIL_FIELDNAMES)と同じ並び。日報モードは**見出しの名前**で
  列を探すので、ここに無い列は空のまま読まれる
- 必須の 日付・ライン・直 は、直の考え方(VBA の YeaDayrCheck / TimeCheck)で決める。
  日付は日報モードの書き方(`2026年11月2日`。月日は 0 で埋めない)、ラインは日報モードの正規の呼び名
  (L-1 / LVC / HVC / トット / バランサー / 中板)、直は 1直 / 2直 / 3直 / 日勤
- 文字コードは BOM 付き UTF-8(日報モードが自分で書く形)。改行は CRLF
- ファイル名は `集計明細_<日付>_<ライン>_<直>.csv`(日報モードの書き出しと同じ始まり)。同じ直は上書き
"""
from __future__ import annotations

import csv
import io
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

from . import lines, paths
from .store import LocalStore, cell_text, line_sheet

#: 日報モード(all-tools tools/nippou)の DETAIL_FIELDNAMES と同じ(2026-11-08 時点)
NIPPOU_FIELDS: tuple[str, ...] = (
    "日付", "ライン", "直", "ページ", "行", "係数処理ﾛｯﾄ数", "ロット№", "材・調質", "製品寸法 厚×幅×丈", "検入枚数",
    "梱包作業時間 開始 時", "梱包作業時間 開始 分", "梱包作業時間 終了 時", "梱包作業時間 終了 分", "作業人数", "合紙",
    "梱包数量 個装単位 枚数", "梱包数量 梱包単位 包数", "ＶＣ種別 両面・片面", "etc 反転・EX etc",
    "作業停止① 記号", "作業停止① 時間(分)", "作業停止② 記号", "作業停止② 時間(分)", "作業停止③ 記号",
    "作業停止③ 時間(分)", "実績合計 枚数", "実績合計 重量", "作業時間(分)", "単重", "用途コード", "用途名", "納入先",
    "包装仕様書No", "コイル縦割", "コイル横縦割", "作業者", "昼稼働", "作業コメント", "引当番号",
    "作業停止④ 記号", "作業停止④ 時間(分)", "作業停止⑤ 記号", "作業停止⑤ 時間(分)",
)

#: 検査の日報シート(2 行目の見出し。前後の空白は無視)→ 日報モードの見出し
COLUMN_MAP: dict[str, str] = {
    # 梱包No(そのロットの何個目の梱包か。NO1・NO2…)は日報モードに当たる欄が無いので渡さない(2026-11-08)。
    # 1 ロットに梱包が 3 つあれば、日報モードでも同じロットの行が 3 行並ぶ
    "ﾛｯﾄNo": "ロット№", "材・調質": "材・調質", "製品寸法": "製品寸法 厚×幅×丈",
    "検入枚数": "検入枚数", "開始時": "梱包作業時間 開始 時", "開始分": "梱包作業時間 開始 分",
    "終了時": "梱包作業時間 終了 時", "終了分": "梱包作業時間 終了 分", "作業人数": "作業人数", "合紙": "合紙",
    "枚数": "梱包数量 個装単位 枚数", "包数": "梱包数量 梱包単位 包数", "VC種別": "ＶＣ種別 両面・片面",
    "etc": "etc 反転・EX etc", "停止①記号": "作業停止① 記号", "停止①時間": "作業停止① 時間(分)",
    "停止②記号": "作業停止② 記号", "停止②時間": "作業停止② 時間(分)", "停止③記号": "作業停止③ 記号",
    "停止③時間": "作業停止③ 時間(分)", "実績枚数": "実績合計 枚数", "実績重量": "実績合計 重量",
    "作業時間": "作業時間(分)", "単重": "単重",
}

#: このツールのライン名(lines.official)→ 日報モードの正規の呼び名
NIPPOU_LINE = {"L-1": "L-1", "LVC": "LVC", "HVC": "HVC", "トット": "トット", "バランサー": "バランサー",
               "中板": "中板", "丸徳": "中板"}


def nippou_date(business_date: str) -> str:
    """"2026.11.02" / "2026/11/02" → "2026年11月2日"(日報モードの書き方)。"""
    parts = business_date.replace("/", ".").replace("-", ".").split(".")
    y, m, d = (int(p) for p in parts[:3])
    return f"{y}年{m}月{d}日"


def nippou_line(line: str) -> str:
    official = lines.official(line) or line
    return NIPPOU_LINE.get(official, official)


def rows_of(store: LocalStore, lname: str) -> list[dict[str, str]]:
    """日報シートの明細行(3 行目から、ロット№ のある行)を 日報モードの見出し → 値 で。"""
    name = line_sheet("Daily", lname)
    if not store.exists(name):
        return []
    cells: dict[tuple[int, int], Any] = {(r, c): v for r, c, v in store.sheet(name).used()}
    heads = {c: cell_text(v).strip() for (r, c), v in cells.items() if r == 2}
    last = max((r for r, _ in cells), default=2)
    out = []
    for r in range(3, last + 1):
        row = {COLUMN_MAP[h]: cell_text(cells.get((r, c))) for c, h in heads.items() if h in COLUMN_MAP}
        if row.get("ロット№", "").strip():
            out.append(row)
    return out


def build_csv(store: LocalStore, lname: str, line: str, business_date: str, shift: str, *,
              worker: str = "") -> tuple[str, int]:
    """CSV の中身と行数。"""
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=list(NIPPOU_FIELDS), lineterminator="\r\n", extrasaction="ignore")
    w.writeheader()
    rows = rows_of(store, lname)
    for i, row in enumerate(rows, start=1):
        w.writerow({"日付": nippou_date(business_date), "ライン": nippou_line(line), "直": shift,
                    "ページ": "1", "行": str(i), "作業者": worker, **row})
    return buf.getvalue(), len(rows)


def file_name(line: str, business_date: str, shift: str) -> str:
    return f"集計明細_{nippou_date(business_date)}_{nippou_line(line)}_{shift}.csv"


def export(store: LocalStore, lname: str, line: str, business_date: str, shift: str, *, worker: str = "",
           folder: Optional[str] = None) -> Optional[Path]:
    """書き出す。行が無ければ書かない(None)。置き場所は使用パス「日報の受け渡し先」、空なら手元の logs。"""
    text, n = build_csv(store, lname, line, business_date, shift, worker=worker)
    if n == 0:
        return None
    base = folder if folder is not None else (paths.get("nippou_out") or paths.write_dir())
    target = paths.to_local(paths.join(base, file_name(line, business_date, shift)))
    if os.sep == "/" and "\\" in str(target):
        target = Path(str(target).replace("\\", "/"))
    # 2026-11-08: 受け渡し先(共有)に届かないときに、フォルダの形を作らない。決めたフォルダ 1 段だけは作ってよい
    target.parent.mkdir(exist_ok=True)
    tmp = target.with_name(target.name + ".part")
    tmp.write_bytes(text.encode("utf-8-sig"))
    tmp.replace(target)
    return target


def export_now(store: LocalStore, state: Any, env: Any, now: Optional[datetime] = None) -> Optional[Path]:
    """いまの直の日報を書き出す(検査画面の状態から ライン・直・検査員 を決める)。"""
    from . import wiring, workers
    now = now or env.now()
    fac = env.facility() or workers.DEFAULT_FACILITY
    shifts = env.shift_times()
    day = workers.yea_dayr_check(workers.hhmm(now), now=now, facility=fac, shifts=shifts)
    shift = workers.time_check(workers.hhmmss(now), facility=fac, shifts=shifts)
    reg = wiring.registry(state)
    who = {"1直": reg.t1, "2直": reg.t2, "3直": reg.t3}.get(shift, "") or reg.t1
    return export(store, state.lname, state.line_name, day, shift, worker=who.replace("_", "・"))
