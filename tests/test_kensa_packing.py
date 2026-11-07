"""包装仕様書の確認(ufkensa/kensa/packing.py: PackingNoCheck・UFPack_Open・UF_PackNo)

期待値は VBA を手で追ったもの(FormRelat:6619-6754 / O:520-543 / UF_PackNo.UserForm_Activate / FormRelat:6826-6915)。
"""
from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

from tests import kensa_fixture
from ufkensa.kensa import packing, paths, quirks
from ufkensa.kensa.store import LOT_HISTORY_HEADERS

LOT = "RX021T0"


def row(lot: str, hiki: object, ord_no: str, pac: object) -> list:
    r: list = [None] * len(LOT_HISTORY_HEADERS)
    r[LOT_HISTORY_HEADERS.index("ﾛｯﾄNo")] = lot
    r[LOT_HISTORY_HEADERS.index("引当NO")] = hiki
    r[LOT_HISTORY_HEADERS.index("ｵｰﾀﾞｰNo")] = ord_no
    r[LOT_HISTORY_HEADERS.index("包装仕様NO")] = pac
    return r


def no_gaichu(*_a, **_k) -> str:
    return ""


class PackingTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="ufkensa_pack_"))
        kensa_fixture.use_offline(kensa_fixture.copy_to(self.tmp / "db"))
        self.folder = paths.to_local(packing.folder_path())
        self.hist = [list(LOT_HISTORY_HEADERS), row(LOT, "61029053", "04093803", "1P1216"),
                     row("OTHER01", "1", "1", "1P0113"), row(LOT, "61029052", "04093316", "1P0113")]

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp, ignore_errors=True)

    def files(self, *names: str) -> None:
        self.folder.mkdir(parents=True, exist_ok=True)
        for n in names:
            (self.folder / n).write_bytes(b"x")

    def check(self, lot: str = LOT, hist=None, messages=None) -> str:
        return packing.packing_no_check(self.hist if hist is None else hist, lot, gaichu=no_gaichu, messages=messages)

    def test_no_sheet_or_lot(self) -> None:
        self.assertEqual(packing.packing_no_check(None, LOT), "")          # WsCheck("Lot情報履歴") = False
        self.assertEqual(self.check(""), "")                               # LotNo = ""
        self.assertEqual(self.check("NOLOT00"), "")                        # a = 0 → Exit Function(フォルダは見ない)

    def test_last_match_in_folder_order_wins(self) -> None:
        self.files("1P0113.JPG", "1P1216.jpg", "9Z9999.JPG")
        self.assertEqual(self.check(), "1P1216")                           # 内側の Exit For だけ → 後の一致が残る
        self.files("1P0113.JPG")
        hist = [self.hist[0], row(LOT, "1", "1", "1P0113")]
        self.assertEqual(self.check(hist=hist), "1P0113")
        self.files("1P0113.jpeg")                                          # "." が 7 文字目 → 左 6 文字
        hist = [self.hist[0], row(LOT, "1", "1", "XXXXXX")]
        self.assertEqual(self.check(hist=hist), "")

    def test_numeric_lot_cell(self) -> None:
        self.files("1P1216.JPG")
        hist = [self.hist[0], row(1234567.0, "1", "1", "1P1216")]          # セルが数 → CStr して文字で比べる
        self.assertEqual(self.check("1234567", hist), "1P1216")

    def test_gaichu_first(self) -> None:
        got = packing.packing_no_check(self.hist, LOT, gaichu=lambda h, lot, r_lot, m: "1P0122")
        self.assertEqual(got, "1P0122")                                    # 外注条件はフォルダを見る前に返す

    def test_merge_sort2(self) -> None:
        rows = [["b", "10"], ["a", "2"], ["c", 3.0], ["d", "10"]]
        # 文字は文字順("10" < "2")、数 < 文字、同じ値は元の順(補助:942-959 の組 → merge2 を手で追った結果)
        self.assertEqual([r[0] for r in packing.merge_sort2(rows, 2)], ["c", "b", "d", "a"])
        rows = [[str(i), v] for i, v in enumerate(["5", "3", "9", "1", "7", "2"])]    # 6 件(2 回目の組で全部まとまる)
        self.assertEqual([r[1] for r in packing.merge_sort2(rows, 2)], ["1", "2", "3", "5", "7", "9"])
        self.assertEqual(packing.merge_sort2([["x", 1]], 2), [["x", 1]])

    def test_pack_open_and_select(self) -> None:
        self.assertIsNone(packing.pack_open("1P1216", LOT, self.hist))      # 画像が無い → フォームは出ない
        self.files("1P1216.jpg", "1P0113.JPG")
        f = packing.pack_open("1P1216", LOT, self.hist)
        self.assertEqual(f.caption, LOT)
        self.assertEqual([(r["hiki"], r["pac"]) for r in f.rows], [("61029052", "1P0113"), ("61029053", "1P1216")])
        self.assertEqual(f.picture, "1P1216")                              # 画像は PackingNoCheck の No
        self.assertEqual(f.spec[0], "1P0113")                              # L1〜L5 は 1 行目(PackSelect(Pac1))
        self.assertEqual(f.selected, 0)
        packing.pack_select(f, 2, self.hist)
        self.assertEqual((f.selected, f.picture, f.spec[0]), (2, "1P1216", "1P1216"))
        packing.pack_select(f, 4, self.hist)                               # 空の行: 画像が無い → 赤字だけ
        self.assertEqual((f.selected, f.picture), (4, "1P1216"))

    def test_pack_open_max_five_rows(self) -> None:
        self.files("1P1216.JPG")
        hist = [self.hist[0]] + [row(LOT, str(9 - i), str(i), "1P1216") for i in range(7)]
        f = packing.pack_open("1P1216", LOT, hist)
        # 先頭 5 行(6 行目以降は TempHiki(6, …) の添字エラーを Resume Next で飛ばす)を引当NO で並べる
        self.assertEqual([r["hiki"] for r in f.rows], ["5", "6", "7", "8", "9"])

    def test_pack_spec(self) -> None:
        self.assertEqual(packing.pack_spec("1P1216")[1:3], ["梱包形態:パレット", "梱包指定:指定（ユーザー指定）"])
        self.assertEqual(packing.pack_spec("ZZZZZZ"), ["ZZZZZZ", "", "", "", ""])   # 行が無い → L1 だけ

    def test_pack_close(self) -> None:
        self.assertEqual(packing.pack_close(False).text, "確認後に押してください。")
        self.assertIsNone(packing.pack_close(True))


if __name__ == "__main__":
    unittest.main()


class PackingB84FixTest(PackingTest):
    """B84 を直した場合: PacN(デバッグ用)を作らないので、空のフォルダ・"." が 7 文字目でない名前で止まらない。"""

    def setUp(self) -> None:
        super().setUp()

    def tearDown(self) -> None:
        super().tearDown()

    def test_folder_errors(self) -> None:
        msgs = []
        self.assertEqual(self.check(messages=msgs), "")                     # フォルダが無くても止めない(画像を出さないだけ)
        self.assertIn("包装仕様書のフォルダが見つからない", msgs[0].text)
        self.files()
        self.assertEqual(self.check(), "")                                 # 空のフォルダ → 一致なし
        self.files("0.txt", "1P1216.JPG")
        self.assertEqual(self.check(), "1P1216")
        hist = [self.hist[0], row(LOT, "1", "1", "0.txt")]                 # 7 文字目でない名前はそのまま照合
        self.assertEqual(self.check(hist=hist), "0.txt")
