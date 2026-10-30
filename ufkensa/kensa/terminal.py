"""この端末の設定(VBA のレジストリ `UF検査判定システム\\起動設定\\設備名` と SF3 フォーム)

VBA は 1 台の PC が 1 ラインを受け持ち、レジストリの「設備名」でラインを決めていた。
未設定なら SF3(初期設定)を出し、UF1 は「ライン設定がありません」で止まる(UF1:220-228)。

- 設備名の選択肢は SF3 と同じ(L-1 / LVC / HVC / バランサー / トット / Test / 丸徳。SF3:74)。
  SF3 は "AIM" も受け付けていた(SF3:28-33)ので残す
- 丸徳(中板)は号機を持つ(UF1 の `LineNo`。VBA は A1〜A11 のボタン。all-tools の中板は 1〜7 → W14)
- 設定の保存は `user_settings`(レジストリの代わり)。変更は管理者(ufkensa:admin)だけ(D12)
"""
from __future__ import annotations

from typing import Optional

from .. import user_settings
from . import lines

KEY_FACILITY = "kensa_facility"     # VBA の「設備名」(レジストリの値そのまま)
KEY_LINE_NO = "kensa_line_no"       # 丸徳の号機(UF1.LineNo)

FACILITIES = ("L-1", "LVC", "HVC", "バランサー", "トット", "Test", "丸徳", "AIM")
MARU_NUMBERS = tuple(f"A{i}" for i in range(1, 12))


def facility() -> str:
    return str(user_settings.get(KEY_FACILITY, "") or "")


def line() -> Optional[lines.Line]:
    """この端末のライン。未設定・知らない値なら None(VBA は UF1 を開かない)。"""
    f = facility()
    return lines.find(f) if f else None


def line_no() -> str:
    return str(user_settings.get(KEY_LINE_NO, "") or "")


def set_facility(value: str, number: str = "") -> None:
    if value not in FACILITIES:
        raise ValueError(f"設備名 {value!r} は選べません(選べるもの: {', '.join(FACILITIES)})")
    if number and number not in MARU_NUMBERS:
        raise ValueError(f"号機 {number!r} は選べません")
    user_settings.save(KEY_FACILITY, value)
    user_settings.save(KEY_LINE_NO, number if value == "丸徳" else "")


def line_label() -> str:
    """帯に出す名前(正規の呼び名。丸徳は号機つき)。"""
    ln = line()
    if ln is None:
        return facility()
    if ln.official == "中板" and line_no():
        return f"{ln.official}{line_no()}"
    return ln.official
