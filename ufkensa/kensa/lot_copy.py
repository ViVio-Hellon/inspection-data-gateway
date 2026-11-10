"""仕掛 DB は手元へ写してから読む(判断事項 D26 改。python-web-tools の source_db と同じ考え方)

- 仕掛(SIKALOT / SIKAHIKI / SIKAODR と 2つ目の名前のもの、SIKACOMP / SIKATWOYEAR)は、
  **共有のファイルを直接開かない**。手元(`app_config.local_dir("cache")/kensa_lot`)へ写し、写しを読む
  (共有のファイルを開いたままにすると上流が差し替えられないため)
- 写せなければそのファイルは読まない(`local_copy` が None)。呼び手は「無い」として扱う:
  オンラインは第2へ、オフラインは「…がありません」
- 写し直すのは、元の 大きさ・更新時刻・inode が変わったときだけ
- 書いている途中を写さないよう、読む鍵(BEGIN + SELECT)を取ってから写す。WAL なら -wal も一緒に写す
"""
from __future__ import annotations

import hashlib
import shutil
import sqlite3
import threading
import time
from pathlib import Path
from typing import Optional

from ..logging_utils import get_logger

log = get_logger("kensa.lot_copy")

BUSY_TIMEOUT_MS = 10_000
_SIDECARS = ("-wal", "-shm", "-journal")
_lock = threading.Lock()
_copies: dict[str, tuple[tuple[int, int, int], Path]] = {}
_checked: dict[str, float] = {}      # 最後に共有フォルダの元を確かめた時刻
RECHECK_SEC = 5.0

from . import cache_registry as _cr                       # noqa: E402 - F3(2026-11-10)
_cr.register_dict("lot_copy._copies", _copies, _lock)
_cr.register_dict("lot_copy._checked", _checked, _lock)


def copy_dir() -> Path:
    from .. import app_config
    d = app_config.local_dir("cache") / "kensa_lot"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _stamp(p: Path) -> tuple[int, int, int]:
    st = p.stat()
    return (st.st_size, st.st_mtime_ns, st.st_ino)


def _uri(p: Path) -> str:
    from .datasource import sqlite_uri
    return sqlite_uri(p.resolve(), "ro")


def _snapshot(src: Path, dst: Path) -> bool:
    """読む鍵を取ってから写す(書き終わった姿だけが写る)。写せたら True。WAL は False(ファイルのまま写す)。"""
    if Path(str(src) + "-wal").exists():
        return False
    try:
        conn = sqlite3.connect(_uri(src), uri=True, timeout=BUSY_TIMEOUT_MS / 1000, isolation_level=None)
    except sqlite3.Error:
        return False
    try:
        conn.execute(f"PRAGMA busy_timeout = {BUSY_TIMEOUT_MS}")
        conn.execute("BEGIN")
        conn.execute("SELECT COUNT(*) FROM sqlite_master").fetchone()
        shutil.copyfile(src, dst)
        conn.execute("COMMIT")
        return True
    except (sqlite3.Error, OSError):
        dst.unlink(missing_ok=True)
        return False
    finally:
        conn.close()


def _raw_copy(src: Path, dst: Path) -> None:
    wal = Path(str(src) + "-wal")
    if wal.exists():
        shutil.copyfile(wal, Path(str(dst) + "-wal"))
    shutil.copyfile(src, dst)


def local_copy(path_text: str) -> Optional[Path]:
    """仕掛 DB(Windows のパス文字列)を手元へ写した実体。写せなければ None(理由は trace に残す)。"""
    from .. import trace
    from . import paths
    src = paths.to_local(path_text)
    key = str(src)
    # 1 回の操作(ロット登録など)で同じ DB を何十回も読む。そのたびに共有フォルダへ「変わったか」を聞くと遅いので、
    # 確かめてから RECHECK_SEC 秒の間は、手元の写しをそのまま使う(2026-11-06 重さの見直し)
    with _lock:
        known = _copies.get(key)
        seen = _checked.get(key, 0.0)
        if known is not None and time.monotonic() - seen < RECHECK_SEC and known[1].is_file():
            return known[1]
    try:
        if not src.is_file():
            return None
        stamp = _stamp(src)
    except OSError:
        return None
    with _lock:
        known = _copies.get(key)
        if known is not None and known[0] == stamp and known[1].is_file():
            _checked[key] = time.monotonic()
            return known[1]
        try:
            dst = copy_dir() / (hashlib.sha1(key.encode("utf-8")).hexdigest()[:16] + "_" + src.name)
            tmp = dst.with_name(dst.name + ".part")
            if not _snapshot(src, tmp):
                _raw_copy(src, tmp)
            # 2026-11-08: 元を写し終えてから、前の写しの付属ファイルを消す(元を読めずに失敗したとき、
            # 前の写しを壊さない)
            for extra in _SIDECARS:
                Path(str(dst) + extra).unlink(missing_ok=True)
            if Path(str(tmp) + "-wal").exists():
                Path(str(tmp) + "-wal").replace(Path(str(dst) + "-wal"))
            tmp.replace(dst)
            after = _stamp(src)
        except OSError as e:
            trace.event("db.copy", f"仕掛 DB を手元へ写せなかったので読みません: {src}({e})", level="WARNING",
                        data={"path": key})
            log.warning("仕掛 DB を手元へ写せませんでした: %s (%s)", src, e)
            _copies.pop(key, None)
            return None
        _copies[key] = (after if after == stamp else stamp, dst)
        _checked[key] = time.monotonic()
        trace.note_used(src, "copy")
        log.info("仕掛 DB を手元へ写しました: %s → %s", src, dst)
        return dst


def forget() -> None:
    """覚えている写しを忘れる(試験用)。"""
    with _lock:
        _copies.clear()
        _checked.clear()
