"""保守 → 出力の確認(2026-11-14 のご依頼)。書き出したものを、見るだけで確かめる。

- Aインプットの CSV … <検査履歴>\\<L>\\<L>_Ainput\\yyyy.mm\\<Lot>\\yyyymmdd_<検査No>(.nn).csv(save_ainput_csv)
                        と hostsend.txt(送信アプリへ渡す電文 1 行)
- 検査結果の PDF    … 手元(C:\\検査履歴)と工場共有の <LN>\\yyyy.mm\\dd日\\<Lot>.pdf(wiring.pdf_folders)
- 検査結果の CSV    … 手元の <L>_履歴CSV.csv・検査履歴\\<ライン>_履歴\\<L>_File.csv・工場共有(export_history_csv)
- 測定DB(梱包検査ﾃﾞｰﾀ)… 検査ﾃﾞｰﾀ / 耳付検査ﾃﾞｰﾀ の最後に足された行

**何も書かない・作らない。** 届かない場所は「見られません」と理由を出すだけ。開けるファイルは、ここで並べた
場所の中だけ(画面から来たパスはここで並べ直したものと突き合わせる)。
"""
from __future__ import annotations

import csv
import io
import os
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Optional

from .. import safe_io
from . import lines, paths, records, wiring

MAX_FILES = 40
MAX_DB_ROWS = 30
MAX_TEXT_BYTES = 2_000_000


def _local(text: str) -> Path:
    p = paths.to_local(str(text))
    if os.sep == "/" and "\\" in str(p):
        p = Path(str(p).replace("\\", "/"))
    return p


def _stat(p: Path) -> Optional[dict[str, Any]]:
    try:
        st = p.stat()
    except OSError:
        return None
    return {"path": str(p), "name": p.name, "size": st.st_size,
            "mtime": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S"), "_t": st.st_mtime}


def _place(label: str, folder: Path, files: list[dict[str, Any]], problem: str = "") -> dict[str, Any]:
    files.sort(key=lambda f: f["_t"], reverse=True)
    for f in files:
        f.pop("_t", None)
    return {"label": label, "folder": str(folder), "files": files[:MAX_FILES], "more": max(0, len(files) - MAX_FILES),
            "problem": problem}


def _scan(label: str, folder: Path, patterns: tuple[str, ...], *, recursive: bool = False) -> dict[str, Any]:
    try:
        if not folder.is_dir():
            return _place(label, folder, [], "フォルダがありません(まだ書いていないか、届きません)")
        found: list[dict[str, Any]] = []
        for pat in patterns:
            for p in (folder.rglob(pat) if recursive else folder.glob(pat)):
                s = _stat(p) if p.is_file() else None
                if s:
                    found.append(s)
        return _place(label, folder, found)
    except OSError as e:
        return _place(label, folder, [], f"見られません: {e}")


def _months(now: datetime, n: int = 2) -> list[str]:
    out, d = [], now.replace(day=1)
    for _ in range(n):
        out.append(d.strftime("%Y.%m"))
        d = (d - timedelta(days=1)).replace(day=1)
    return out


# ------------------------------------------------------------------
def ainput(line: str, now: datetime) -> list[dict[str, Any]]:
    from . import ainput as ain
    folder = ain.ainput_folder(line)
    places = []
    for m in _months(now):
        base = _local(paths.join(paths.get("history"), folder, folder + "_Ainput", m))
        places.append(_scan(f"Aインプットの CSV({m})", base, ("*.csv",), recursive=True))
    host = _local(ain.hostsend_path())
    s = _stat(host)
    places.append(_place("hostsend.txt(送信アプリへ渡す電文)", host.parent, [s] if s else [],
                         "" if s else "まだ書いていません(または届きません)"))
    return places


def pdf(line: str, now: datetime) -> list[dict[str, Any]]:
    places = []
    for day in (now, now - timedelta(days=1)):
        for i, folder in enumerate(wiring.pdf_folders(line, day)):
            where = "手元" if i == 0 else "工場共有"
            places.append(_scan(f"検査結果の PDF {where}({day:%m/%d})", _local(folder), ("*.pdf", "*.html")))
    return places


def result_csv(line: str, lname: str, store_dir: Path) -> list[dict[str, Any]]:
    official = lines.official(line)
    names = (lname + "_履歴CSV.csv", lname + "_耳用履歴CSV.csv")
    local_dir = _local(paths.get("csv_local")) if paths.get("csv_local") else store_dir
    out = [_one("検査結果の CSV 手元", local_dir, names)]
    new = (lname + "_File.csv", lname + "_耳用File.csv")
    hist = records._HISTORY_DIRS.get(official, "Test\\Test_履歴")
    out.append(_one("検査結果の CSV 検査履歴", _local(paths.join(paths.get("history"), hist)), new))
    if official in records._FACTORY_DIRS:
        out.append(_one("検査結果の CSV 工場共有",
                        _local(paths.join(paths.get("factory"), records._FACTORY_DIRS[official])), new))
    return out


def _one(label: str, folder: Path, names: tuple[str, ...]) -> dict[str, Any]:
    files = [s for s in (_stat(folder / n) for n in names) if s]
    return _place(label, folder, files, "" if files else "まだ書いていません(または届きません)")


def measure_db(line: str, db: Optional[Path] = None) -> dict[str, Any]:
    """測定DB の最後に足された行(ラインで絞る。新しい順)。開くのは読むだけ。"""
    p = db or paths.to_local(paths.db_file("梱包検査ﾃﾞｰﾀ"))
    out: dict[str, Any] = {"path": str(p), "tables": [], "problem": ""}
    s = _stat(Path(p))
    if not s:
        out["problem"] = "測定DB(梱包検査ﾃﾞｰﾀ)が見つかりません。設定 → 使用パス を確かめてください"
        return out
    out["mtime"] = s["mtime"]
    vba = (lines.find(line).vba if lines.find(line) else line)
    try:
        con = safe_io.connect_existing(Path(p), timeout=5)
        try:
            for table in records.MEASURE_TABLES:
                if not con.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone():
                    out["tables"].append({"table": table, "columns": [], "rows": [], "count": 0, "problem": "表がありません"})
                    continue
                cols = [r[1] for r in con.execute(f'PRAGMA table_info("{table}")')]
                count = con.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0]
                first = cols[0] if cols else ""
                recs = con.execute(f'SELECT rowid, * FROM "{table}" WHERE "{first}" IN (?, ?) ORDER BY rowid DESC LIMIT ?',
                                   (line, vba, MAX_DB_ROWS)).fetchall() if first else []
                out["tables"].append({"table": table, "columns": ["rowid"] + cols, "count": count,
                                      "rows": [["" if v is None else str(v) for v in r] for r in recs], "problem": ""})
        finally:
            con.close()
    except (OSError, sqlite3.Error) as e:
        out["problem"] = f"測定DB を読めませんでした: {e}"
    return out


# ------------------------------------------------------------------
def allowed(path: str, places: list[dict[str, Any]]) -> bool:
    return any(f["path"] == path for pl in places for f in pl["files"])


def read_text(path: str) -> dict[str, Any]:
    """CSV・txt を表として返す(cp932 → だめなら UTF-8)。"""
    p = Path(path)
    data = p.read_bytes()[:MAX_TEXT_BYTES]
    for enc in ("cp932", "utf-8-sig"):
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text, enc = data.decode("cp932", "replace"), "cp932?"
    rows = list(csv.reader(io.StringIO(text))) if p.suffix.lower() == ".csv" else [[ln] for ln in text.splitlines()]
    return {"path": path, "encoding": enc, "rows": rows[-300:] if len(rows) > 300 else rows, "total": len(rows),
            "cut": len(rows) > 300}
