"""保守 → 出力の確認(2026-11-14 のご依頼)。書き出したものを、見るだけで確かめる。

- Aインプットの CSV … <検査履歴>\\<L>\\<L>_Ainput\\yyyy.mm\\<Lot>\\yyyymmdd_<検査No>(.nn).csv(save_ainput_csv)
                        と hostsend.txt(送信アプリへ渡す電文 1 行)
- 検査結果の PDF    … 手元(C:\\検査履歴)と工場共有の <LN>\\yyyy.mm\\dd日\\<Lot>.pdf(wiring.pdf_folders)
- 検査結果の CSV    … 手元の <L>_履歴CSV.csv・検査履歴\\<ライン>_履歴\\<L>_File.csv・工場共有(export_history_csv)
- 測定DB(梱包検査ﾃﾞｰﾀ)… 検査ﾃﾞｰﾀ / 耳付検査ﾃﾞｰﾀ の最後に足された行

**書き換えない・消さない。**(Aインプットの今月の yyyy.mm フォルダだけは無ければ作る。2026-11-18) 届かない場所は「見られません」と理由を出すだけ。開けるファイルは、ここで並べた
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


MAX_DIRS = 60


def _clean_rel(rel: Any) -> list[str]:
    """画面から来た「根の下の道」(/ 区切り)。.. や絶対パスは捨てる。"""
    parts = [p for p in str(rel or "").replace("\\", "/").split("/") if p and p not in (".", "..")]
    return [p for p in parts if ":" not in p]


def _browse(label: str, root: str, start: list[str], patterns: tuple[str, ...],
            rel: Optional[list[str]] = None) -> dict[str, Any]:
    """2026-11-23(現場の声): 使用パス(根)から下のフォルダを選んで見る。
    初めは今月・今日のフォルダ(start)、無ければ在る所までさかのぼる。配下のフォルダとファイルを並べる。"""
    base = _local(root)
    want = list(start if rel is None else rel)
    note = ""
    while want and not base.joinpath(*want).is_dir():
        if rel is None and not note:
            note = f"「{'\\'.join(start)}」はまだありません(書いたときに作ります)。在るところを出しています"
        want.pop()
    folder = base.joinpath(*want) if want else base
    dirs: list[dict[str, Any]] = []
    files: list[dict[str, Any]] = []
    problem = ""
    try:
        if not folder.is_dir():
            problem = "使用パスのフォルダがありません(届かないか、まだありません)"
        else:
            for p in folder.iterdir():
                if p.is_dir():
                    try:
                        t = p.stat().st_mtime
                    except OSError:
                        continue
                    dirs.append({"name": p.name, "mtime": datetime.fromtimestamp(t).strftime("%Y-%m-%d %H:%M"), "_t": t})
                elif p.is_file() and any(p.match(pat) for pat in patterns):
                    s = _stat(p)
                    if s:
                        files.append(s)
    except OSError as e:
        problem = f"見られません: {e}"
    dirs.sort(key=lambda d: (d["_t"], d["name"]), reverse=True)
    for d in dirs:
        d.pop("_t", None)
    out = _place(label, folder, files, problem)
    out.update({"browse": True, "root": str(base), "rel": want, "dirs": dirs[:MAX_DIRS],
                "dirs_more": max(0, len(dirs) - MAX_DIRS), "note": note, "patterns": list(patterns)})
    return out


def _months(now: datetime, n: int = 2) -> list[str]:
    out, d = [], now.replace(day=1)
    for _ in range(n):
        out.append(d.strftime("%Y.%m"))
        d = (d - timedelta(days=1)).replace(day=1)
    return out


# ------------------------------------------------------------------
def ainput(line: str, now: datetime, rels: Optional[dict[int, list[str]]] = None) -> list[dict[str, Any]]:
    from . import ainput as ain
    places = []
    # 2026-11-18(現場の声): 今月の yyyy.mm フォルダは無ければ作る(送信のときにも作る。届かない共有には作らない)
    try:
        safe_io.mkdir_reachable(_local(paths.join(ain.ainput_csv_root(line), now.strftime("%Y.%m"))))
    except OSError:
        pass
    rels = rels or {}
    places.append(_browse("Aインプットの CSV", ain.ainput_csv_root(line), [now.strftime("%Y.%m")], ("*.csv",), rels.get(0)))
    host = _local(ain.hostsend_path())
    s = _stat(host)
    places.append(_place("hostsend.txt(送信アプリへ渡す電文)", host.parent, [s] if s else [],
                         "" if s else "まだ書いていません(または届きません)"))
    return places


def pdf(line: str, now: datetime, rels: Optional[dict[int, list[str]]] = None) -> list[dict[str, Any]]:
    rels = rels or {}
    start = [now.strftime("%Y.%m"), now.strftime("%d") + "日"]
    return [_browse(f"検査結果の PDF {'手元' if i == 0 else '工場共有'}", root, start, ("*.pdf", "*.html"), rels.get(i))
            for i, root in enumerate(wiring.pdf_roots(line))]


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
    """一覧に出したファイル、または選べるフォルダ(根)の下で決まった種類のファイルだけ。"""
    if any(f["path"] == path for pl in places for f in pl["files"]):
        return True
    p = Path(path)
    if ".." in p.parts:
        return False
    for pl in places:
        if not pl.get("browse"):
            continue
        root = Path(pl["root"])
        try:
            inside = p.resolve().is_relative_to(root.resolve())
        except OSError:
            continue
        if inside and p.is_file() and any(p.match(pat) for pat in pl["patterns"]):
            return True
    return False


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
