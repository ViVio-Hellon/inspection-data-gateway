"""「無い」と「読めない」を取り違えない読み書きの道具(2026-11-08)

決まり: **「無い」と言ってよいのは、読みに行って本当に無いと確かめられたときだけ。**
読めなかった(権限・ネットワーク切れ・JSON や sqlite の壊れ・ロック・親フォルダに届かない)ときは、
書かない・消さない・覚えない。

- `read_json_or_absent(path)` → (状態, 中身)。状態は OK / ABSENT / UNREADABLE
- `confirm_absent(path)` → 本当に無いと確かめられたら True(親フォルダに届かなければ False)
- `connect_existing(path)` → あるファイルだけを sqlite で開く(無ければ作らずに例外)
"""
from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from typing import Any, Union
from urllib.parse import quote as _urlquote

OK = "ok"
ABSENT = "absent"
UNREADABLE = "unreadable"


class Unreadable(OSError):
    """読めなかった(無いとは言えない)。文言はそのまま画面・ログに出せる。"""


def confirm_absent(path: Union[str, Path], *, local: bool = False) -> bool:
    """`path` が本当に無いか。親フォルダが読めて、その中に無いときだけ True。

    `Path.exists()` は共有フォルダに届かないときも False を返すので、それだけでは「無い」と言えない。
    `local=True`(この PC の手元のファイル)は、親が無くても、さかのぼって最初にあるものがフォルダなら「無い」とする
    (初めて起動したときなど。共有フォルダには使わない)。
    """
    p = Path(path)
    try:
        if p.exists():
            return False
        if p.parent.is_dir():
            return True
        if not local:
            return False
        for up in p.parents:
            if up.exists():
                return up.is_dir()
        return False
    except OSError:
        return False


def mkdir_under(root: Union[str, Path], target: Union[str, Path]) -> Path:
    """`root`(設定で決めた共有の置き場所)に届くときだけ、その下の `target` を作る。

    届かないのに mkdir(parents=True) すると、共有の形を勝手に作ってしまう(ドライブの割り当てが外れたときなど)。
    """
    r, t = Path(root), Path(target)
    try:
        ok = r.is_dir()
    except OSError:
        ok = False
    if not ok:
        raise Unreadable(f"置き場所に届かないのでフォルダを作りません: {r}")
    t.mkdir(parents=True, exist_ok=True)
    return t


def mkdir_reachable(target: Union[str, Path]) -> Path:
    """出力先のフォルダを作る。ただし、設定のパスの根元(共有 \\\\server\\share やドライブ)より下に、
    **すでにあるフォルダが 1 つでも見えるとき**だけ(2026-11-08)。

    共有に届かない・ドライブの割り当てが外れたときは、根元しか見えない(または何も見えない)ので作らずに
    `Unreadable`。手元のフォルダ(すでにある所の下)はこれまでどおり作る。
    """
    t = Path(target)
    try:
        if t.is_dir():
            return t
        anchor = Path(t.anchor) if t.anchor else None
        base = None
        for up in t.parents:
            if anchor is not None and up == anchor:
                break
            if up.exists():
                base = up
                break
    except OSError:
        base = None
    if base is None or not base.is_dir():
        raise Unreadable(f"出力先に届かないのでフォルダを作りません: {t}(共有フォルダならつながっているか確かめてください)")
    t.mkdir(parents=True, exist_ok=True)
    return t


def read_json_or_absent(path: Union[str, Path], encoding: str = "utf-8", *, local: bool = False) -> tuple[str, Any]:
    """JSON を読む → (状態, 中身)。無い → (ABSENT, None)、読めない・壊れている → (UNREADABLE, 理由の文字)。"""
    p = Path(path)
    try:
        text = p.read_text(encoding=encoding)
    except FileNotFoundError:
        return (ABSENT, None) if confirm_absent(p, local=local) else (UNREADABLE, f"フォルダに届きません: {p.parent}")
    except (OSError, UnicodeError) as exc:
        return UNREADABLE, str(exc)
    try:
        return OK, json.loads(text)
    except ValueError as exc:
        return UNREADABLE, f"JSON として読めません: {exc}"


def _uri(p: Path, mode: str) -> str:
    text = str(p).replace("\\", "/")
    if text.startswith("//"):
        text = "//" + text
    return "file:" + _urlquote(text, safe="/:") + "?mode=" + mode


def connect_existing(path: Union[str, Path], *, timeout: float = 10.0, **kwargs: Any) -> sqlite3.Connection:
    """あるファイルだけを読み書きで開く(`mode=rw`)。無い・届かなければ作らずに sqlite3.OperationalError。

    `sqlite3.connect(path)` は、共有のマスタが一瞬見えないだけで空の DB を作ってしまう。
    """
    p = Path(path)
    return sqlite3.connect(_uri(p.absolute(), "rw"), uri=True, timeout=timeout, **kwargs)
