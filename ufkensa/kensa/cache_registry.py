"""覚え(キャッシュ)の一覧(2026-11-10 リファクタリング 第1段 1・F3)

読んだ表・仕掛の写し・公差表などを覚えている所は、モジュールの読み込み時にここへ「捨て方」と「大きさ」を
登録する。設定 → 使用パス の保存・「読み込む」(`reload_data.clear_caches`)は、登録された覚えを全部捨てる。
覚えを足したら、そのモジュールで `register` を呼ぶこと(捨て忘れ F3 の再発を防ぐ)。

覚えるモジュールは `MODULES` に並べる(`load_all` で読み込んで、登録をそろえる)。
"""
from __future__ import annotations

import importlib
import threading
from typing import Any, Callable, Optional

#: 覚えを登録するモジュール(ufkensa.kensa の下)。足したらここにも足す
MODULES = ("datasource", "judge", "lot_copy", "paths", "quality", "reasons")

_lock = threading.Lock()
_entries: dict[str, tuple[Callable[[], None], Callable[[], int]]] = {}


def register(name: str, clear: Callable[[], None], size: Callable[[], int]) -> None:
    """覚えを 1 つ登録する。`clear` は捨てる、`size` は覚えている数(空なら 0)。同じ名前は置き換える。"""
    with _lock:
        _entries[name] = (clear, size)


def register_dict(name: str, d: Any, lock: Optional[Any] = None) -> None:
    """dict / set / list の覚えを登録する(`lock` があればそれを取って捨てる)。"""
    def clear() -> None:
        if lock is None:
            d.clear()
        else:
            with lock:
                d.clear()
    register(name, clear, lambda: len(d))


def load_all() -> None:
    for m in MODULES:
        importlib.import_module(f"{__package__}.{m}")


def names() -> list[str]:
    load_all()
    with _lock:
        return sorted(_entries)


def sizes() -> dict[str, int]:
    load_all()
    with _lock:
        items = list(_entries.items())
    return {n: size() for n, (_clear, size) in items}


def clear_all() -> list[str]:
    """登録された覚えを全部捨てる。捨てた名前を返す。"""
    load_all()
    with _lock:
        items = list(_entries.items())
    for _name, (clear, _size) in items:
        clear()
    return [n for n, _ in items]
