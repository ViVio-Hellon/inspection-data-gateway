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


#: 最後に写せなかった理由(画面の「第1パスの台帳を読めません」に添える)
_last_error: dict[str, str] = {}
_cr.register_dict("lot_copy._last_error", _last_error, _lock)


def last_error(path_text: str) -> str:
    from . import paths
    with _lock:
        return _last_error.get(str(paths.to_local(path_text)), "")


def _base_name(key: str, src: Path) -> Path:
    return copy_dir() / (hashlib.sha1(key.encode("utf-8")).hexdigest()[:16] + "_" + src.name)


def _newest_on_disk(base: Path) -> Optional[Path]:
    """前に写した姿(この PC に残っているもの。前の起動のものも)。新しい順で最初の 1 つ。"""
    try:
        found = [q for q in base.parent.glob(base.stem + "*" + base.suffix)
                 if q.is_file() and not q.name.endswith(".part")]
    except OSError:
        return None
    found.sort(key=lambda q: q.stat().st_mtime, reverse=True)
    return found[0] if found else None


def _fallback(key: str, base: Path, why: str) -> Optional[Path]:
    """2026-11-26(現場の声「第1パスの台帳を読めません」): 元を今は写せない(共有に届かない・上流が作り直し中・
    前の写しを開いている)とき、前に写せた姿があればそれを使う(少し古いだけで、読めないよりよい)。"""
    _last_error[key] = why
    prev = _copies.get(key)
    old = prev[1] if prev is not None and prev[1].is_file() else _newest_on_disk(base)
    if old is not None:
        log.warning("仕掛 DB を今は写せないので、前に写した姿を使います: %s(%s。写し %s)", key, why, old.name)
        _checked[key] = time.monotonic()
        if prev is None:
            _copies[key] = ((-1, -1, -1), old)
        return old
    log.warning("仕掛 DB を手元へ写せず、前の写しもありません: %s(%s)", key, why)
    return None


def local_copy(path_text: str) -> Optional[Path]:
    """仕掛 DB(Windows のパス文字列)を手元へ写した実体。写せなければ前に写した姿、それも無ければ None。"""
    from .. import trace
    from . import paths
    src = paths.to_local(path_text)
    key = str(src)
    base = _base_name(key, src)
    # 1 回の操作(ロット登録など)で同じ DB を何十回も読む。そのたびに共有フォルダへ「変わったか」を聞くと遅いので、
    # 確かめてから RECHECK_SEC 秒の間は、手元の写しをそのまま使う(2026-11-06 重さの見直し)
    with _lock:
        known = _copies.get(key)
        seen = _checked.get(key, 0.0)
        if known is not None and time.monotonic() - seen < RECHECK_SEC and known[1].is_file():
            return known[1]
    try:
        if not src.is_file():
            if src.parent.is_dir():                 # フォルダには届く = 本当に無い(前の写しは使わない)
                with _lock:
                    _last_error[key] = "共有のフォルダにファイルがありません"
                return None
            with _lock:
                return _fallback(key, base, "共有のフォルダに届きません")
        stamp = _stamp(src)
    except OSError as e:
        with _lock:
            return _fallback(key, base, f"共有のファイルを確かめられません: {e}")
    with _lock:
        known = _copies.get(key)
        if known is not None and known[0] == stamp and known[1].is_file():
            _checked[key] = time.monotonic()
            _last_error.pop(key, None)
            return known[1]
        tmp = base.with_name(base.name + ".part")
        try:
            if not _snapshot(src, tmp):
                _raw_copy(src, tmp)
            after = _stamp(src)
        except OSError as e:
            tmp.unlink(missing_ok=True)
            trace.event("db.copy", f"仕掛 DB を手元へ写せませんでした: {src}({e})", level="WARNING", data={"path": key})
            return _fallback(key, base, f"写せませんでした: {e}")
        dst = _place(tmp, base, key)
        if dst is None:
            return _fallback(key, base, "手元の写しを置き換えられませんでした(前の写しを開いている)")
        _copies[key] = (after if after == stamp else stamp, dst)
        _checked[key] = time.monotonic()
        _last_error.pop(key, None)
        trace.note_used(src, "copy")
        log.info("仕掛 DB を手元へ写しました: %s → %s", src, dst)
        return dst


def _place(tmp: Path, base: Path, key: str) -> Optional[Path]:
    """写したものを置く。Windows は開いているファイルを置き換えられないので、置き換えられなければ
    別の名前(base_<番号>)に置く(2026-11-26。前は ここで失敗して「読めません」になっていた)。"""
    for i in range(0, 20):
        dst = base if i == 0 else base.with_name(f"{base.stem}_{int(time.time() * 1000)}_{i}{base.suffix}")
        try:
            for extra in _SIDECARS:
                Path(str(dst) + extra).unlink(missing_ok=True)
            if Path(str(tmp) + "-wal").exists():
                Path(str(tmp) + "-wal").replace(Path(str(dst) + "-wal"))
            tmp.replace(dst)
        except OSError as e:
            log.info("手元の写し %s を置き換えられませんでした(%s)。別の名前で置きます", dst.name, e)
            continue
        _sweep(base, keep=dst)
        return dst
    tmp.unlink(missing_ok=True)
    return None


def _sweep(base: Path, keep: Path) -> None:
    """古い別名の写しを消す(開いていて消せないものは次の機会に)。"""
    try:
        for q in base.parent.glob(base.stem + "_*" + base.suffix):
            if q != keep:
                try:
                    q.unlink()
                except OSError:
                    pass
        if keep != base:
            try:
                base.unlink()
            except OSError:
                pass
    except OSError:
        pass


def forget() -> None:
    """覚えている写しを忘れる(試験用)。"""
    with _lock:
        _copies.clear()
        _checked.clear()
        _last_error.clear()
