"""Access(.accdb)→ SQLite(.sqlite3)の取り込み(判断事項 D4: すべて SQLite に切り替え)

アプリは SQLite だけを読み書きする。Access のファイルしか無いあいだは、この取り込みで
同じ場所・同じ名前の .sqlite3 を作る(cscript → bridge.vbs → ADODB → ACE。追加ライブラリは要らない)。

- 取り込むと、表の行の順は Access の返却順のまま入る(VBA の「最初に一致した行」を同じにするため)
- 書き切ってから差し替える(`<名前>.sqlite3.tmp` → 前のものは `.bak` に退ける)。読みかけの PC に書きかけを見せない
- 仕掛(SIKALOT / SIKAHIKI / SIKAODR)は上流が SQLite で毎時作り直すので、ここでは取り込まない
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Optional

from ..accdb_bridge import importer
from ..logging_utils import get_logger
from . import paths

log = get_logger("kensa.importing")

#: 取り込む DB(VBA が読み書きしていた Access)。梱包検査ﾃﾞｰﾀ は後日受領(04 の資料)
#: 仕掛の SIKALOT / SIKAHIKI / SIKAODR は上流が SQLite で作るので取り込まない
TARGETS: tuple[str, ...] = ("SIKACOMP", "SIKATWOYEAR", "品質基準", "伝送用ファイル",
                            "梱包検査ﾃﾞｰﾀ")


@dataclass
class PlanRow:
    name: str
    source: str
    target: str
    source_exists: bool
    target_exists: bool
    target_mtime: Optional[float]

    def to_dict(self) -> dict[str, Any]:
        return self.__dict__.copy()


def default_source(name: str) -> str:
    """取り込み元の既定: 取り込み先と同じ場所・同じ名前の .accdb(VBA が読んでいたファイル)。"""
    target = paths.db_file(name)
    return target[: -len(".sqlite3")] + ".accdb"


def plan() -> list[PlanRow]:
    rows = []
    for name in TARGETS:
        src, dst = default_source(name), paths.db_file(name)
        sp, dp = paths.to_local(src), paths.to_local(dst)
        try:
            mtime = dp.stat().st_mtime if dp.is_file() else None
        except OSError:
            mtime = None
        rows.append(PlanRow(name, src, dst, _is_file(sp), mtime is not None, mtime))
    return rows


def _is_file(p: Path) -> bool:
    try:
        return p.is_file()
    except OSError:
        return False


def run(name: str, source: str = "", progress: Optional[Callable[[int, str], None]] = None) -> importer.ImportResult:
    """1 つ取り込む。`source` を省くと既定の .accdb。"""
    if name not in TARGETS:
        raise KeyError(f"{name} は取り込む DB の一覧にありません")
    src = paths.to_local(source or default_source(name))
    dst = paths.to_local(paths.db_file(name))
    from .. import trace
    trace.note_used(src, "accdb")
    trace.note_used(dst, "save")
    with trace.timed("import.accdb", f"{name} を取り込む", source=str(src), target=str(dst)) as info:
        res = importer.import_accdb(src, dst, progress=progress)
        info["tables"] = len(res.tables)
        info["rows"] = sum(n for _, n in res.tables)
    return res


def can_run_here() -> tuple[bool, str]:
    """この PC で取り込めるか(cscript があるか)。"""
    if os.name != "nt":
        return False, "Access の取り込みは Windows(cscript.exe と Access Database Engine)でだけ動きます。"
    try:
        importer.cscript_path()
    except importer.BridgeError as e:
        return False, str(e)
    return True, ""
