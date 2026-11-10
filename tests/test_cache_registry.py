"""覚え(キャッシュ)の一覧と「読み込む」で全部捨てること(2026-11-10 第1段 1・F3)

F3: reload_data.clear_caches が、あとから足した覚え(仕掛台帳の補った形・索引・第1の索引・理由の選択肢・
届かないパスの覚え・仕掛の写し)を捨てていなかった。いまは各モジュールが cache_registry に登録し、
clear_caches が登録された覚えを全部捨てる。
"""
from __future__ import annotations

import unittest

from tests.test_kensa_lot_source import _TwoSlots
from ufkensa.kensa import cache_registry, datasource, judge, lot_copy, paths, quality, reasons, reload_data

#: これだけは必ず登録されていること(F3 で捨て忘れていたものを含む)
EXPECTED = {
    "datasource._table_cache", "datasource._lot_ready", "datasource._first_index", "datasource._eq_index",
    "lot_copy._copies", "lot_copy._checked", "paths._slow", "reasons._last_known",
    "quality._gaichu_cache", "quality._gauge_cache", "quality._pack_cache", "judge._tol_cache",
    "paths._settings",
}


class RegistryTest(unittest.TestCase):
    def test_expected_caches_registered(self) -> None:
        self.assertLessEqual(EXPECTED, set(cache_registry.names()))

    def test_clear_empties_every_registered_cache(self) -> None:
        # 覚えに何か入れる(中身は捨てられるかだけを見る)
        datasource._table_cache[("x", "t")] = ((1, 1), [], [])              # noqa: SLF001
        datasource._lot_ready["x"] = None                                   # noqa: SLF001
        datasource._first_index["SIKALOTNOW"] = None                        # noqa: SLF001
        datasource._eq_index[(1, 2, 3)] = None                              # noqa: SLF001
        lot_copy._copies["x"] = ((1, 1, 1), None)                           # noqa: SLF001
        lot_copy._checked["x"] = 1.0                                        # noqa: SLF001
        paths._slow[("x", "any")] = (True, 0.0)                             # noqa: SLF001
        reasons._last_known[("x", "消去")] = ["a"]                          # noqa: SLF001
        judge._tol_cache["x"] = None                                        # noqa: SLF001
        from pathlib import Path
        quality._gaichu_cache.store(Path("無い"), None)                     # noqa: SLF001
        paths.get("main")                                                   # 設定の覚え(第1段 2)
        sizes = cache_registry.sizes()
        for name in EXPECTED - {"quality._gauge_cache", "quality._pack_cache"}:
            self.assertGreater(sizes[name], 0, name)
        cleared = reload_data.clear_caches()
        self.assertLessEqual(EXPECTED, set(cleared))
        self.assertEqual({n: s for n, s in cache_registry.sizes().items() if s}, {})


class ReloadAfterRealReadTest(_TwoSlots):
    """実際にロットを登録して覚えがたまった後、clear_caches で全部空になり、次の登録も同じ結果になる。"""

    def test_clear_after_register(self) -> None:
        self.register()
        before = self.ken_h()
        sizes = cache_registry.sizes()
        self.assertGreater(sizes["datasource._table_cache"], 0)
        self.assertGreater(sizes["lot_copy._copies"], 0)
        reload_data.clear_caches()
        self.assertEqual({n: s for n, s in cache_registry.sizes().items() if s}, {})
        self.reregister()
        self.assertEqual(self.ken_h(), before)


if __name__ == "__main__":
    unittest.main()
