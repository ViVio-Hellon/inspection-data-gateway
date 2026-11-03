"""(開発用・Linux)mdbtools で .accdb を bridge.vbs と同じ形の TSV にする

本番の取り込みは Windows の cscript + bridge.vbs(ufkensa/accdb_bridge)。
これは**試験用の基準データ**(tests/fixtures/kensa/)を作るためだけの道具で、ラインPCでは使わない。

    python scripts/dev_accdb_to_tsv.py 品質基準.accdb tests/fixtures/kensa/品質基準
    python scripts/dev_accdb_to_tsv.py SIKALOTNOW.accdb out --where 仕掛:ﾛｯﾄ番号=N456J52,H7372B0
    python scripts/dev_accdb_to_tsv.py 伝送用ファイル.accdb out --mask 名簿

出力は bridge.vbs と同じ: `_schema.tsv`(表 列 ADO型 大きさ)と `<表>.tsv`(1行目が列名、\\t \\n \\r \\\\ をエスケープ、
NULL は \\N、日付は yyyy/mm/dd hh:nn:ss)。文字コードは UTF-8(importer はどちらも読む)。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
from datetime import datetime
from pathlib import Path

ADO = {"Text": 202, "Memo": 203, "Long Integer": 3, "Integer": 2, "Byte": 17, "Double": 5, "Single": 4,
       "Currency": 6, "Boolean": 11, "DateTime": 7, "Replication ID": 72, "Numeric": 131, "OLE": 205}
ENV = dict(os.environ, LANG="C.UTF-8", LC_ALL="C.UTF-8")


def run(*args: str) -> str:
    return subprocess.run(args, check=True, capture_output=True, env=ENV).stdout.decode("utf-8")


NOT_NULL: dict[str, set[str]] = {}


def schema(accdb: str) -> dict[str, list[tuple[str, int, int]]]:
    """表 → (列, ADO 型, 大きさ)。必須(NOT NULL)の列は NOT_NULL に控える。

    mdbtools は長さ 0 の文字列も NULL と同じに書き出す。必須の列に NULL は入らない(Access が断る)ので、
    必須の列で値が無いものは "" に戻す(VBA が読む値と同じにするため)。"""
    out: dict[str, list[tuple[str, int, int]]] = {}
    table = None
    for line in run("mdb-schema", accdb).splitlines():
        m = re.match(r"CREATE TABLE \[(.+)\]", line)
        if m:
            table = m.group(1)
            out[table] = []
            continue
        m = re.match(r"\s*\[(.+?)\]\s+(.+?)(?:\s*\((\d+)\))?(?:\s+NOT NULL)?,?\s*$", line)
        if m and table is not None:
            typ = m.group(2).strip()
            out[table].append((m.group(1), ADO.get(typ, 202), int(m.group(3) or 0)))
            if "NOT NULL" in line:
                NOT_NULL.setdefault(table, set()).add(m.group(1))
        if line.startswith(");"):
            table = None
    return out


def esc(v) -> str:
    if v is None:
        return "\\N"
    s = str(v)
    return s.replace("\\", "\\\\").replace("\t", "\\t").replace("\n", "\\n").replace("\r", "\\r")


def fix_date(v):
    if isinstance(v, str):
        for f in ("%m/%d/%y %H:%M:%S", "%m/%d/%Y %H:%M:%S"):
            try:
                d = datetime.strptime(v, f)
                return d.strftime("%Y/%m/%d") if d.time() == datetime.min.time() else d.strftime("%Y/%m/%d %H:%M:%S")
            except ValueError:
                pass
    return v


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("accdb")
    ap.add_argument("out")
    ap.add_argument("--tables", default="", help="カンマ区切り。省くと ~TMP 以外の全部")
    ap.add_argument("--where", action="append", default=[], help="表:列=値1,値2 (その値の行だけ)")
    ap.add_argument("--mask", action="append", default=[], help="個人情報を伏せる表(名簿)")
    a = ap.parse_args()
    sch = schema(a.accdb)
    want = [t for t in (a.tables.split(",") if a.tables else sch) if t and not t.startswith("~TMP")]
    wh: dict[str, tuple[str, set[str]]] = {}
    for w in a.where:
        t, cond = w.split(":", 1)
        c, vals = cond.split("=", 1)
        wh[t] = (c, set(vals.split(",")))
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    with open(out / "_schema.tsv", "w", encoding="utf-8", newline="") as f:
        for t in want:
            for c, ado, size in sch[t]:
                f.write(f"{t}\t{c}\t{ado}\t{size}\r\n")
    for t in want:
        cols = [c for c, _, _ in sch[t]]
        dates = {c for c, ado, _ in sch[t] if ado == 7}
        texts_required = {c for c, ado, _ in sch[t] if ado in (202, 203) and c in NOT_NULL.get(t, set())}
        rows = [json.loads(l) for l in run("mdb-json", a.accdb, t).splitlines() if l.strip()]
        if t in wh:
            c, vals = wh[t]
            rows = [r for r in rows if str(r.get(c, "")) in vals]
        if t in a.mask:
            for i, r in enumerate(rows, 1):
                for k in list(r):
                    if k in ("名前", "よみ"):
                        r[k] = f"検査員{i:02d}" if k == "名前" else f"けんさいん{i:02d}"
                    elif k in ("NO", "バーコード", "NO加工無"):
                        r[k] = f"{9000000000000 + i}" if k != "バーコード" else f"*{9000000000000 + i}*"
        for r in rows:
            for c in texts_required:
                if r.get(c) is None:
                    r[c] = ""
        with open(out / f"{t}.tsv", "w", encoding="utf-8", newline="") as f:
            f.write("\t".join(esc(c) for c in cols) + "\r\n")
            for r in rows:
                f.write("\t".join(esc(fix_date(r.get(c)) if c in dates else r.get(c)) for c in cols) + "\r\n")
        print(t, len(rows))


if __name__ == "__main__":
    main()
