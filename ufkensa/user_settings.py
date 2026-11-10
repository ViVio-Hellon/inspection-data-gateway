"""利用者ごとの設定の永続化 (VBA `GetSetting`/`SaveSetting` の代替)

VBA版は Windows レジストリに保存していた。

    GetSetting(SysName, "起動設定", "モード", "存在なし")
    SaveSetting SysName, "起動設定", "モード", "オフライン"

Python版はレジストリに依存しないよう、JSONファイル(`config.USER_CONFIG_PATH`)
に保存する。読み書きのたびにファイルを開くので、外部から書き換えても
次回の読み取りで反映される。鍵の一覧は `config.KEY_*`。
"""
from __future__ import annotations

import json
from typing import Any

from . import config, safe_io
from .logging_utils import get_logger

log = get_logger("user_settings")

# 表示用。VBA `GetSetting(..., "存在なし")` に相当する「未設定」の言葉
UNSET_LABEL = "未設定"


def _load() -> tuple[str, dict[str, Any]]:
    """(状態, 中身)。2026-11-08: 「無い」と「読めない」を分ける(読めないのに空として保存すると、ほかの設定が消える)。"""
    path = config.USER_CONFIG_PATH
    state, data = safe_io.read_json_or_absent(path, local=True)
    if state == safe_io.ABSENT:
        return state, {}
    if state == safe_io.UNREADABLE:
        log.warning("設定ファイルを読めませんでした(既定値で続行・保存はしません): %s (%s)", path, data)
        return state, {}
    if not isinstance(data, dict):
        log.warning("設定ファイルの中身が想定と違います(保存はしません): %s", path)
        return safe_io.UNREADABLE, {}
    return state, data


def load_with_state() -> tuple[str, dict[str, Any]]:
    """(状態, 中身)。状態は safe_io.OK / ABSENT / UNREADABLE。覚えてよいかを呼び手が決める用(paths。2026-11-10)。"""
    return _load()


#: 保存・削除のあとに呼ぶ(覚えた設定を捨てる。paths が登録する。2026-11-10)
_listeners: list = []


def on_change(fn) -> None:
    if fn not in _listeners:
        _listeners.append(fn)


def _changed() -> None:
    for fn in list(_listeners):
        try:
            fn()
        except Exception as exc:  # noqa: BLE001 - 捨てられなくても保存は済んでいる
            log.warning("設定の覚えを捨てられませんでした: %s", exc)


def load_all() -> dict[str, Any]:
    """設定ファイル全体を読む。読めなければ例外にせず空(既定値)で返す。**保存には使わない**(`save` は読めないとき書かない)。"""
    return _load()[1]


#: 最後に「設定ファイルを読めないので保存しなかった」ときの文言(画面に出す用。2026-11-08)
last_error = ""


class NotSaved(RuntimeError):
    """設定ファイルを読めないので保存しなかった(2026-11-08)。文言はそのまま画面に出す。"""


def save_or_raise(key: str, value: Any) -> None:
    """`save` と同じ。ただし設定ファイルを読めずに保存しなかったときは `NotSaved`(画面の操作から呼ぶ。2026-11-08)。"""
    if not save(key, value) and last_error:
        raise NotSaved(last_error)


def get(key: str, default: Any = None) -> Any:
    """VBA `GetSetting` 相当。"""
    return load_all().get(key, default)


# ログに値を出さない鍵。伏せるだけで、保存はふつうに行う
HIDDEN_IN_LOG = ("admin_password", "password")


def save(key: str, value: Any) -> bool:
    """VBA `SaveSetting` 相当。書き込めなければ False を返す。"""
    global last_error
    state, data = _load()
    last_error = ""
    if state == safe_io.UNREADABLE:
        # 2026-11-08: 読めないのに空として書くと、ほかの設定がすべて消える。書かない
        last_error = f"設定ファイルを読めないので保存しませんでした: {config.USER_CONFIG_PATH}"
        log.warning("%s(%s)", last_error, key)
        return False
    data[key] = value
    try:
        config.USER_CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        config.USER_CONFIG_PATH.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError as exc:
        log.warning("設定ファイルに書けませんでした: %s", exc)
        _changed()
        return False
    _changed()
    # **値を出さない鍵がある。** 管理者パスワードは撹拌済みだが、
    # ログに残す理由が無い(残せば持ち出せる)
    log.info("設定を保存しました: %s=%s", key,
             "(伏せます)" if key in HIDDEN_IN_LOG else value)
    return True


def delete(key: str) -> bool:
    """VBA `DeleteSetting` 相当。無ければ何もしない(True)。"""
    global last_error
    state, data = _load()
    if state == safe_io.UNREADABLE:
        # 2026-11-08: 読めないときは消さない(書き戻すとほかの設定が消える)
        last_error = f"設定ファイルを読めないので保存しませんでした: {config.USER_CONFIG_PATH}"
        log.warning("%s(%s を消す)", last_error, key)
        return False
    if key not in data:
        return True
    del data[key]
    try:
        config.USER_CONFIG_PATH.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError as exc:
        log.warning("設定ファイルに書けませんでした: %s", exc)
        _changed()
        return False
    _changed()
    log.info("設定を消しました: %s", key)
    return True
