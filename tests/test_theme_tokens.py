"""色のトークン(app/static/css/tokens.css)の約束を確かめる

* 暗いテーマの値は 2 か所(OS が暗い ``@media`` の中と、利用者が選んだ ``[data-theme="dark"]``)に要る。
  元は ``@media`` の中で、``scripts/sync_theme_tokens.py`` で写す。2 か所が食い違っていないこと
* 暗いテーマが上書きする名前は、明るいテーマ(``:root``)にもあること(片方だけの名前は打ち間違い)
* 部品の CSS に色の直値(``#rrggbb``)を書かないこと。色は tokens.css から引く
* 部品の CSS が使う ``var(--名前)`` は、どこかで決めてあること(決めていない名前は後ろの既定値で黙って塗られる)
"""
from __future__ import annotations

import re
import subprocess
import sys
import unittest
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CSS = ROOT / "app" / "static" / "css"
sys.path.insert(0, str(ROOT / "scripts"))
import sync_theme_tokens as sync  # noqa: E402

#: CSS では決めない名前(画面の JS が要素に書く・既定値つきで使う)
SET_BY_JS = {"--judge-size", "--judge-size-h", "--w"}


def _strip_comments(text: str) -> str:
    return re.sub(r"/\*.*?\*/", "", text, flags=re.S)


class ThemeTokens(unittest.TestCase):
    def setUp(self) -> None:
        self.text = (CSS / "tokens.css").read_text(encoding="utf-8")

    def _light(self) -> list[tuple[str, str]]:
        a, b = sync._block(self.text, self.text.index(":root{"))
        return sync.declarations(self.text[a:b])

    def test_dark_blocks_are_identical(self) -> None:
        src, (a, b) = sync.split(self.text)
        media, chosen = sync.declarations(src), sync.declarations(self.text[a:b])
        self.assertTrue(media)
        self.assertEqual(media, chosen, "@media の中と [data-theme=dark] が食い違う。scripts/sync_theme_tokens.py で写す")
        self.assertEqual(sync.synced(self.text), self.text, "[data-theme=dark] が @media の中の写しになっていない")

    def test_sync_script_check(self) -> None:
        r = subprocess.run([sys.executable, str(ROOT / "scripts" / "sync_theme_tokens.py"), "--check"],
                           capture_output=True, text=True, timeout=60)
        self.assertEqual(r.returncode, 0, r.stderr)

    def test_sync_rewrites_a_drifted_copy(self) -> None:
        drifted = self.text.replace(':root[data-theme="dark"]{\n  --ground:#101820;', ':root[data-theme="dark"]{\n  --ground:#000000;')
        self.assertNotEqual(drifted, self.text)
        self.assertEqual(sync.synced(drifted), self.text)

    def test_no_duplicate_names(self) -> None:
        src, (a, b) = sync.split(self.text)
        for name, decls in (("light", self._light()), ("dark", sync.declarations(src))):
            dup = [k for k, n in Counter(k for k, _ in decls).items() if n > 1]
            self.assertEqual(dup, [], f"{name}: 同じ名前を 2 回決めている")

    def test_dark_names_exist_in_light(self) -> None:
        src, _ = sync.split(self.text)
        light = {k for k, _ in self._light()}
        extra = sorted({k for k, _ in sync.declarations(src)} - light)
        self.assertEqual(extra, [], "暗いテーマにだけある名前")

    def test_component_css_has_no_hex_colours(self) -> None:
        bad = []
        for f in sorted(CSS.glob("*.css")):
            if f.name == "tokens.css":
                continue
            for n, line in enumerate(_strip_comments(f.read_text(encoding="utf-8")).splitlines(), 1):
                if re.search(r"#[0-9a-fA-F]{3,8}\b", line):
                    bad.append(f"{f.name}: {line.strip()}")
        self.assertEqual(bad, [], "色は tokens.css のトークンから引く")

    def test_used_variables_are_defined(self) -> None:
        defined = set(SET_BY_JS)
        used = set()
        for f in CSS.glob("*.css"):
            t = _strip_comments(f.read_text(encoding="utf-8"))
            defined |= set(re.findall(r"(--[\w-]+)\s*:", t))
            used |= set(re.findall(r"var\(\s*(--[\w-]+)", t))
        self.assertEqual(sorted(used - defined), [], "決めていない名前を使っている")


if __name__ == "__main__":
    unittest.main()
