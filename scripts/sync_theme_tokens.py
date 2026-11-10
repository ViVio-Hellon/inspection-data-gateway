"""暗いテーマの色を 1 か所から写す(app/static/css/tokens.css)

暗いテーマの値は tokens.css に 2 回書く必要がある:

* ``@media (prefers-color-scheme: dark){ :root:not([data-theme="light"]){ … } }`` … OS が暗い・明示の指定なし
* ``:root[data-theme="dark"]{ … }``                                               … 利用者が「暗い」を選んだ

CSS だけでは(入れ子の CSS を使わずに)この 2 つを 1 つの規則にまとめられない。そこで **@media の中を元** にして、
下の ``:root[data-theme="dark"]`` の中身をこの道具で書き写す。手で 2 か所を直して食い違うのを防ぐ。

    python scripts/sync_theme_tokens.py          # 写す(変わったときだけ書く)
    python scripts/sync_theme_tokens.py --check  # 揃っているかだけ見る(揃っていなければ 1 で終わる)

tests/test_theme_tokens.py が ``--check`` と同じことを確かめる。
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

TOKENS = Path(__file__).resolve().parent.parent / "app" / "static" / "css" / "tokens.css"
MEDIA_HEAD = "@media (prefers-color-scheme: dark){"
MEDIA_RULE = ':root:not([data-theme="light"]){'
DARK_RULE = ':root[data-theme="dark"]{'


def _block(text: str, start: int) -> tuple[int, int]:
    """``start`` にある ``{`` の中身の範囲 [開き括弧の次, 閉じ括弧) を返す(注釈の中の括弧は数えない)。"""
    i = text.index("{", start) + 1
    depth, j = 1, i
    while depth:
        if text.startswith("/*", j):
            j = text.index("*/", j) + 2
            continue
        c = text[j]
        depth += (c == "{") - (c == "}")
        j += 1
    return i, j - 1


def split(text: str) -> tuple[str, tuple[int, int]]:
    """(@media の中の暗い値の本文, 写し先 ``:root[data-theme="dark"]`` の中身の範囲) を返す。"""
    m = text.index(MEDIA_HEAD)
    r = text.index(MEDIA_RULE, m)
    src = text[slice(*_block(text, r))]
    d = text.index(DARK_RULE, m)
    return src, _block(text, d)


def render(src: str) -> str:
    """@media の中(字下げ 4)を、トップの規則(字下げ 2)の中身にする。"""
    lines = src.strip("\n").split("\n")
    out = []
    for ln in lines:
        if ln.startswith("    "):
            out.append(ln[2:])
        elif ln.strip():
            out.append("  " + ln.lstrip(" "))
        else:
            out.append("")
    return "\n" + "\n".join(out).rstrip() + "\n"


def declarations(body: str) -> list[tuple[str, str]]:
    """本文の ``--名前: 値;`` を順に返す(注釈は除く)。"""
    body = re.sub(r"/\*.*?\*/", "", body, flags=re.S)
    return [(k, " ".join(v.split())) for k, v in re.findall(r"(--[\w-]+)\s*:\s*([^;]+);", body)]


def synced(text: str) -> str:
    src, (a, b) = split(text)
    return text[:a] + render(src) + text[b:]


def main(argv: list[str]) -> int:
    text = TOKENS.read_text(encoding="utf-8")
    new = synced(text)
    if "--check" in argv:
        if new != text:
            print("tokens.css: :root[data-theme=\"dark\"] が @media の中と食い違っています。"
                  "python scripts/sync_theme_tokens.py で写してください", file=sys.stderr)
            return 1
        return 0
    if new != text:
        TOKENS.write_text(new, encoding="utf-8")
        print("tokens.css を書き直しました")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
