"""共有に置いた本体を、起動のときに各 PC の手元へ写して、手元の写しから動かす(2026-11-07)。

【なぜ】
共有の本体を直接動かすと、
- 動いている途中で本体を差し替えると、途中で読む部品だけ新しくなり、古い部品と混ざる
- 共有が一瞬切れると、画面の部品(HTML・JS)や途中で読む部品が読めず、検査の途中で止まる

【しくみ】(標準ライブラリだけ。ufkensa は読み込まない)
1. start_app.py は起動の最初にここを呼ぶ。本体が**ネットワーク上**(\\\\サーバ\\… か ネットワークドライブ)なら、
2. 本体の中身の一覧(ファイル名・大きさ・更新時刻)から版の印(ハッシュ)を作り、
3. 手元 `%LOCALAPPDATA%\\UFKensa\\app\\v-<印>\\` が無ければ、まず仮のフォルダへ丸ごと写してから名前を変える
   (写しかけを使わない)。印が同じなら写さない(2 回目からは写さないので速い)
4. 手元の start_app.py を同じ引数で起動し、終わるのを待つ。共有の本体を書き換えても、動いている PC は
   手元の写しのまま(次に起動したときに新しい版を写す)
5. 写せなかったら(共有が遅い・空きが無い…)、今まで通り共有から動かす(起動はあきらめない)

配布設定(`配布設定\\`)は共有のものを見る(環境変数 UFKENSA_DISTRIBUTION_DIR)。
`UFKENSA_NO_LOCAL_COPY=1` で写さない(試すとき)。
"""
from __future__ import annotations

import hashlib
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Optional

APP_NAME_DIR = "UFKensa"
#: 写さないもの(動かすのに要らない・各 PC で別に持つもの)
SKIP_DIRS = {".git", "tests", "docs", "vba_src", "__pycache__", "配布設定", ".pytest_cache", "tools"}
SKIP_FILES = {"Access.log"}
SKIP_SUFFIXES = (".pyc", ".log", ".tmp")
#: 手元に残す版の数(いま使っているもの + 1 つ前)
KEEP_VERSIONS = 2
ENV_FROM = "UFKENSA_FROM_SHARE"          # 写しから起動したとき、元の共有の本体の場所
ENV_OFF = "UFKENSA_NO_LOCAL_COPY"


def is_network_path(p: Path) -> bool:
    """\\\\サーバ\\… か、ネットワークドライブ(Windows の DRIVE_REMOTE)か。"""
    s = str(p)
    if s.startswith("\\\\") or s.startswith("//"):
        return True
    if os.name != "nt" or len(s) < 2 or s[1] != ":":
        return False
    try:
        import ctypes
        return ctypes.windll.kernel32.GetDriveTypeW(s[:2] + "\\") == 4      # DRIVE_REMOTE
    except Exception:                                   # noqa: BLE001
        return False


def local_base() -> Path:
    override = os.environ.get("UFKENSA_LOCAL_DIR")
    if override and override.strip():
        return Path(override.strip()) / "app"
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / ".local" / "share")
    return Path(base) / APP_NAME_DIR / "app"


def _files(root: Path) -> list[Path]:
    out = []

    def _fail(exc: OSError) -> None:
        raise exc

    # 2026-11-08: os.walk は読めないフォルダを黙って飛ばす。共有が一瞬切れると「中身の欠けた版」を
    # 完成として写し、古い版を消していた。読めなければ OSError(写さない・消さない)
    for dirpath, dirnames, filenames in os.walk(root, onerror=_fail):
        dirnames[:] = sorted(d for d in dirnames if d not in SKIP_DIRS and not d.startswith("."))
        for f in sorted(filenames):
            if f in SKIP_FILES or f.endswith(SKIP_SUFFIXES):
                continue
            out.append(Path(dirpath) / f)
    return out


def version_mark(root: Path, files: Optional[list[Path]] = None) -> str:
    h = hashlib.sha1()
    for f in files if files is not None else _files(root):
        st = f.stat()
        h.update(f"{f.relative_to(root).as_posix()}|{st.st_size}|{int(st.st_mtime)}\n".encode("utf-8"))
    return h.hexdigest()[:12]


def ensure_copy(src: Path, dst_base: Optional[Path] = None) -> Path:
    """手元の写し(v-<印>)を用意して、その場所を返す。写せなければ OSError。"""
    dst_base = dst_base or local_base()
    files = _files(src)
    if not any(f.name == "start_app.py" for f in files):
        # 2026-11-08: 本体が見えない = 読めていない。空の版を作らない
        raise OSError(f"共有の本体を読めません(start_app.py が見えない): {src}")
    mark = version_mark(src, files)
    dst = dst_base / f"v-{mark}"
    if (dst / ".complete").is_file():
        return dst
    tmp = dst_base / f"tmp-{mark}-{os.getpid()}"
    shutil.rmtree(tmp, ignore_errors=True)
    for f in files:
        to = tmp / f.relative_to(src)
        to.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(f, to)
    (tmp / ".complete").write_text(f"{src}\n{time.strftime('%Y-%m-%d %H:%M:%S')}\n", encoding="utf-8")
    try:
        tmp.rename(dst)
    except OSError:
        if not (dst / ".complete").is_file():          # ほかの起動が先に作ったのでなければ失敗
            raise
        shutil.rmtree(tmp, ignore_errors=True)
    _cleanup(dst_base, keep=dst)
    return dst


def _cleanup(base: Path, keep: Path) -> None:
    """古い版を消す(使っている最中のものは消せないので、消せなければそのまま)。"""
    vers = sorted((p for p in base.glob("v-*") if p.is_dir()), key=lambda p: p.stat().st_mtime, reverse=True)
    for p in [v for v in vers if v != keep][KEEP_VERSIONS - 1:]:
        shutil.rmtree(p, ignore_errors=True)
    for p in base.glob("tmp-*"):
        if time.time() - p.stat().st_mtime > 3600:
            shutil.rmtree(p, ignore_errors=True)


def relaunch_from_local(app_root: Path, argv: list[str]) -> Optional[int]:
    """共有から起動されたら手元の写しから起動し、その終了コードを返す。写さないときは None(そのまま続ける)。"""
    if os.environ.get(ENV_OFF) == "1" or os.environ.get(ENV_FROM):
        return None
    if not is_network_path(app_root):
        return None
    try:
        local = ensure_copy(app_root)
    except OSError as e:
        sys.stderr.write(f"手元への写しに失敗したので、共有から起動します: {e}\n")
        return None
    env = dict(os.environ)
    env[ENV_FROM] = str(app_root)
    env.setdefault("UFKENSA_DISTRIBUTION_DIR", str(app_root / "配布設定"))
    return subprocess.call([sys.executable, str(local / "start_app.py"), *argv], env=env, cwd=str(local))
