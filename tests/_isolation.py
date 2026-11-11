"""テストが本物の設定・記録を書き換えないようにする

【なぜ `tests/__init__.py` だけでは足りないのか】
`tests/__init__.py` は環境変数を先に立てることで `ufkensa.config` に
一時フォルダを掴ませている。これは**その `__init__.py` が読まれる場合に
限って**効く。ところが unittest の探索は、開始フォルダをそのまま
最上位フォルダとして扱うと、テストを**パッケージ配下ではなく単体の
モジュールとして**取り込む。つまり:

    python3 -m unittest discover -s tests        # tests/__init__.py は読まれない
    python3 -m unittest discover -s tests -t .   # 読まれる(READMEはこちら)

前者で流すと安全網が丸ごと外れ、`data/user_config.json`(追跡対象)に
テストの値が書き込まれ、`%LOCALAPPDATA%` の記録も本番と混ざる。
**黙って起きる**のが厄介で、テストは全部OKのまま作業ツリーだけが汚れる。

そこでこのモジュールが、環境変数に頼らず**いま掴んでいるパスを見て**
リポジトリ内や本番領域を指していたら一時フォルダへ向け直す。

【なぜ取り込み時に直せば間に合うのか】
unittest はテストを**全部取り込んでから**走らせる。書き込みが起きるのは
走っている最中なので、どれか1つのテストモジュールが取り込み時にここを
呼んでいれば、その時点で全テストぶんの向き先が直っている
(`tests/test_isolation.py` が呼んでいる)。
"""
from __future__ import annotations

import atexit
import os
import shutil
import tempfile
from pathlib import Path

# 何度呼ばれても1回だけ効かせる(取り込み順に依存させない)
_done = False


def _under(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
    except (ValueError, OSError):
        return False
    return True


def _escape(name: str) -> Path:
    tmp = Path(tempfile.mkdtemp(prefix=f"ufkensa_test_{name}_"))
    atexit.register(shutil.rmtree, tmp, True)
    return tmp


def ensure_isolated() -> list[str]:
    """危ない向き先を一時フォルダへ替える。替えた項目の名前を返す。

    既に `tests/__init__.py` が逃がしてあれば何もしない(戻り値は空)。
    """
    global _done
    if _done:
        return []
    _done = True

    from ufkensa import config, logging_utils

    repo = Path(config.BASE_DIR)
    moved: list[str] = []

    # 利用者設定: リポジトリの `data/user_config.json` は**追跡対象**。
    # ここに書かれると、テストを流しただけで差分が出る
    if _under(config.USER_CONFIG_PATH, repo) or not _under(config.USER_CONFIG_PATH, Path(tempfile.gettempdir())):
        config.USER_CONFIG_PATH = _escape("config") / "user_config.json"
        moved.append("USER_CONFIG_PATH")

    # ログと、ロック・長時間処理の記録が入るローカル領域。
    # 本番の端末で流すと、動いているアプリの記録に混ざる
    for name in ("LOG_DIR",):
        current = getattr(config, name)
        if _under(current, repo) or not _under(current, Path(tempfile.gettempdir())):
            setattr(config, name, config.tool_log_dir(_escape(name.lower())))     # 0.33.0: ツール名のフォルダ
            moved.append(name)
    # 記録の置き場所を設定で変えても、既定(戻り先)は一時フォルダ
    config.DEFAULT_LOG_DIR = config.LOG_DIR

    # 配布設定(0.30.0): 既定はリポジトリ直下。試験で書き出すと作業ツリーにできてしまう
    from ufkensa import distribution
    if _under(distribution.DIR, repo):
        distribution.DIR = _escape("dist") / "配布設定"
        moved.append("distribution.DIR")

    # 不具合の扱い(config/quirks.json): 試験も本番と同じ扱いで動かす(下)。
    # 試験は書いた直後にファイル・DB を読むので、まとめ書き(trace)と共有フォルダの確かめの間引き(lot_copy)は切る。
    # 本番の動き(まとめ書き・間引き)は tests/test_speed.py で確かめる
    from ufkensa import trace as _trace
    from ufkensa.kensa import lot_copy as _lot_copy
    _trace.SYNC = True
    # 第3段7・9 の影の確かめ: 試験では食い違いがあれば失敗させる
    from ufkensa.kensa import shadow as _shadow
    _shadow.STRICT = True
    from ufkensa.kensa import match_rules as _mr
    _mr.SHADOW_STRICT = True
    _lot_copy.RECHECK_SEC = 0.0
    from ufkensa.kensa import quirks
    if _under(quirks.QUIRKS_FILE, repo):
        quirks.QUIRKS_FILE = _escape("quirks") / "quirks.json"
        moved.append("quirks.QUIRKS_FILE")
        import shutil as _sh
        quirks.QUIRKS_FILE.parent.mkdir(parents=True, exist_ok=True)
        _sh.copy2(repo / "config" / "quirks.json", quirks.QUIRKS_FILE)

    # ローカル領域(%LOCALAPPDATA%\…): アクセス権限の「前回の結果」(runtime\access_rights.json)・仕掛の写し・控え。
    # `tests/__init__.py` が読まれない流しかたでも、本番の端末の記録と混ぜない
    import os
    local = os.environ.get("UFKENSA_LOCAL_DIR", "")
    if not local or not _under(Path(local), Path(tempfile.gettempdir())):
        os.environ["UFKENSA_LOCAL_DIR"] = str(_escape("local"))
        moved.append("UFKENSA_LOCAL_DIR")

    # 使用パス: 試験が保存していない項目は既定(本番の共有 \\nlmsrvngy03\… や C:\検査履歴)のまま。
    # 共有と、一時フォルダの外のドライブのパスは、ここの下へ読み替える(`paths.to_local` の安全網)。
    # 試験が kensa_fixture.use_offline などで一時フォルダへ向けた項目はそのまま
    from ufkensa.kensa import paths
    if not os.environ.get(paths.GUARD_ENV):
        os.environ[paths.GUARD_ENV] = str(_escape("paths"))
        moved.append(paths.GUARD_ENV)

    # Access.log: 本体と同じフォルダ(試験ではリポジトリ直下)に書いてしまうので止める
    from ufkensa import user_settings
    user_settings.save(config.KEY_ACCESS_LOG, False)

    logging_utils.silence_console()
    return moved

