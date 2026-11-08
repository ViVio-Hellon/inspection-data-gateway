"""理由選択のマスタ(2026-11-08 のご依頼)。伝送用ファイル.sqlite3 の表「理由選択」。

VBA は理由の選択肢をフォーム(UFConfi)に直書きしていた。ここでは表で持ち、保守 → マスタ管理 で直せるようにする。
**表が無ければ、VBA と同じ選択肢で自動で作る。**

    区分        いつ出すか
    消去        測定値をダブルクリックで消したとき・公差外の確認で理由を聞くとき(VBA の List1)
    再測定      同じ位置を 2 回以上「再測定」したとき(VBA の List2)

列: 区分・順番・理由・有効(1 = 出す、0 = 出さない)。順番の小さい順に並べる。
特別な意味を持つ理由: 「板廃棄」(消した値を戻して異常ロットへ写す)・「再測定」「その他」(理由の文字が要る)。
表を読めない・書けないときは、VBA と同じ既定の選択肢を使う(画面は止めない)。
"""
from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Optional

from .. import safe_io
from ..logging_utils import get_logger
from . import datasource

log = get_logger("kensa.reasons")

TABLE = "理由選択"
DB = "伝送用ファイル"
KIND_DELETE = "消去"
KIND_REMEASURE = "再測定"
DEFAULTS: dict[str, tuple[str, ...]] = {
    KIND_DELETE: ("入力位置間違い", "測定位置間違い", "測定ボタン押し間違い", "再測定", "0点確認", "板廃棄", "その他"),
    KIND_REMEASURE: ("入力位置間違い", "測定位置間違い", "測定ボタン押し間違い", "0点確認", "その他"),
}
COLUMNS = ("区分", "順番", "理由", "有効")

#: 2026-11-08: 最後に読めた選択肢((ファイル, 区分) → 選択肢)。読めないときは空や既定にせず、これを出す
#: (保守で足した理由が、共有フォルダが一瞬切れただけで画面から消えないように)。読めたときだけ覚える
_last_known: dict[tuple[str, str], list[str]] = {}


def _db_path() -> Path:
    return datasource.resolve(DB)


def ensure_table(path: Optional[Path] = None) -> bool:
    """表が無ければ作って既定の選択肢を入れる。作った・もうあれば True、作れなければ False。"""
    p = path or _db_path()
    if not p.is_file():
        return False
    try:
        # 2026-11-08: マスタが一瞬見えないとき sqlite3.connect は空の DB を作る。あるファイルだけを開く
        con = safe_io.connect_existing(p, timeout=5)
        try:
            have = con.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?", (TABLE,)).fetchone()
            if have:
                return True
            # ここに来るのは sqlite_master を読めて、表が本当に無いと確かめられたときだけ(読めなければ sqlite3.Error)
            with con:
                con.execute(f'CREATE TABLE "{TABLE}" ("区分" TEXT NOT NULL, "順番" INTEGER NOT NULL, '
                            f'"理由" TEXT NOT NULL, "有効" INTEGER NOT NULL DEFAULT 1)')
                for kind, items in DEFAULTS.items():
                    con.executemany(f'INSERT INTO "{TABLE}" VALUES (?, ?, ?, 1)',
                                    [(kind, i, r) for i, r in enumerate(items, start=1)])
            log.info("理由選択のマスタを作りました: %s", p)
            return True
        finally:
            con.close()
    except sqlite3.Error as e:
        log.warning("理由選択のマスタを作れませんでした(%s): %s", p, e)
        return False


def choices(kind: str, path: Optional[Path] = None) -> list[str]:
    """出す選択肢(有効なものを順番どおり)。読めなければ VBA と同じ既定。"""
    p = path or _db_path()
    # 2026-11-08: 読めないときは、最後に読めた選択肢(無ければ VBA の既定)。表は作り直さない・書かない
    default = list(_last_known.get((str(p), kind)) or DEFAULTS.get(kind, ()))
    if not ensure_table(p):
        log.warning("理由選択のマスタを読めないので、最後に読めた選択肢(または既定)を出します: %s", p)
        return default
    try:
        cols, rows = datasource.get_table(str(p), TABLE)
    except (datasource.DbMissing, datasource.TableMissing, OSError, sqlite3.Error):
        log.warning("理由選択のマスタを読めないので、最後に読めた選択肢(または既定)を出します: %s", p)
        return default
    idx = {c: i for i, c in enumerate(cols)}
    if not all(c in idx for c in COLUMNS):
        return default
    out = []
    for r in rows:
        if str(r[idx["区分"]]).strip() != kind or str(r[idx["有効"]]).strip() not in ("1", "True", "true", "-1"):
            continue
        text = str(r[idx["理由"]]).strip()
        if text:
            try:
                order = float(r[idx["順番"]])
            except (TypeError, ValueError):
                order = 9999.0
            out.append((order, text))
    got = [t for _, t in sorted(out, key=lambda x: x[0])] or list(DEFAULTS.get(kind, ()))
    _last_known[(str(p), kind)] = got
    return got
