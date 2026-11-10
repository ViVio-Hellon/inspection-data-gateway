"""検査画面の受け渡しの名前(view の一番上の名前)の約束(2026-11-10 リファクタリング 第1段 3)

- view.build が作る名前 = view.BASE_KEYS、Session が渡す view 全体 = view.VIEW_KEYS
- 画面 kensa.js が読む `v.名前` / `V.名前` / `v["名前"]` は、すべて VIEW_KEYS にある(F6 のような名前のずれを見つける)
"""
from __future__ import annotations

import re
import unittest
from pathlib import Path

from tests.test_kensa_session import SessionBase
from ufkensa.kensa import view
from ufkensa.kensa.grid import KensaState

JS = Path(__file__).resolve().parent.parent / "app" / "static" / "js" / "views" / "kensa.js"
#: kensa.js で view を受け取る変数は v(描く関数の引数)と V(最後に受け取った view)
_USE = re.compile(r"""(?<![\w.$])[vV](?:\.([A-Za-z_]\w*)|\[["']([A-Za-z_]\w*)["']\])""")


def js_view_keys(text: str) -> set[str]:
    return {a or b for a, b in _USE.findall(text)}


class ViewContractTest(SessionBase):
    def test_build_matches_base_keys(self) -> None:
        self.assertEqual(set(view.build(KensaState())), set(view.BASE_KEYS))

    def test_session_view_matches_declaration(self) -> None:
        self.assertEqual(set(self.s.current().view), set(view.VIEW_KEYS))       # 開く前
        self.register()
        self.assertEqual(set(self.s.current().view), set(view.VIEW_KEYS))       # ロット登録の後
        r = self.s.receive("DT10001+0000015.012M")
        self.assertEqual(set(r.view), set(view.VIEW_KEYS))                      # 操作の結果

    def test_js_uses_only_declared_keys(self) -> None:
        used = js_view_keys(JS.read_text(encoding="utf-8"))
        self.assertGreater(len(used), 20)                                       # 読み取りの形が変わったら気づく
        self.assertEqual(used - set(view.VIEW_KEYS), set())

    def test_scanner(self) -> None:
        self.assertEqual(js_view_keys('v.header.lot; V.pack; v["range"]; const vv = x.v; v2.no; a.v.b'), {"header", "pack", "range"})


if __name__ == "__main__":
    unittest.main()
