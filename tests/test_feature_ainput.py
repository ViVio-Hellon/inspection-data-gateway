"""機能テスト: AINPUTフォーム起動 / 高さ測定での挙動・判定 / EX(輸出)を認識し表示が切り替わるか

VBA: vba_src/UFAinput.frm HikiNo_Change(317-367) — ｵｰﾀﾞｰNo の 3 文字目が "1" なら輸出(EX):
ExportLabel・ExFrame を出しフォーム高さ 620、それ以外は隠して 490。Remain_Change(545-573) は EX 表示中に
ストア(1)なら外寸 0,0,0、出荷(0)なら空。AinButt_Click(990-1020) は EX のとき ｹｰｽNo・外寸を検査する。
同じロットに引当No が 2 つ(通常 04077194 / 輸出 01177194)ある Lot情報履歴 で、画面(API)を通して確かめる。
"""
from __future__ import annotations

import unittest

from tests.test_screen_ainput import AinputScreenTest, HIKI, LOT, OK_TEXT, ROW
from ufkensa.kensa.store import LOT_HISTORY

HIKI_EX = "60602049"
ROW_EX = list(ROW)
ROW_EX[1], ROW_EX[21] = "01177194", HIKI_EX


class FeatureAinputTest(AinputScreenTest):
    def setUp(self) -> None:
        super().setUp()
        self.store.ensure_sheet(LOT_HISTORY).append(ROW_EX)   # 一時フィクスチャにだけ足す

    def measure(self, hiki: str, height: str = "150.000") -> dict:
        self.ok(self.post("open", ken_m="L1"))
        d = self.ok(self.post("value", name="hiki_no", value=hiki))
        for n, v in (("konno", "1"), ("mai1", "10"), ("cou1", "1")):
            self.ok(self.post("value", name=n, value=v))
        self.ok(self.post("mount", value="1"))
        self.ok(self.post("press"))
        return self.ok(self.post("receive", text=f"DT10006+0000{height}M"))

    # 1) 起動 ------------------------------------------------------------
    def test_startup_from_kensa_loads_lot_and_hiki(self) -> None:
        v = self.ok(self.post("open", ken_m="L1"))["view"]
        self.assertTrue(v["ready"])
        self.assertEqual((v["lot"]["lot"], v["lot"]["ken_no"], v["lot"]["amin"], v["lot"]["amax"], v["ken_m"]),
                         (LOT, "H6X1355", "14.95", "15.05", "L1"))
        self.assertEqual(sorted(v["hiki_list"]), sorted([HIKI, HIKI_EX]))
        self.assertFalse(v["show"]["ex"])                       # 引当No を選ぶまでは EX は出ない

    def test_startup_errors(self) -> None:
        self.assertEqual(self.post("open")["stop"]["text"], "検査法を選択してください")
        d = self.post("open", ken_m="L1", source="Cask")
        self.assertIn("キャスク画面でロットを開いて", d["stop"]["text"])
        d = self.post("open", ken_m="L1", source="Mitsubishi")
        self.assertIn("全ヵ所の画面でロットを開いて", d["stop"]["text"])
        self.s.state.tmin = ""
        self.assertEqual(self.post("open", ken_m="L1")["stop"]["text"], "MinMax処理を行ってください")

    # 2) 高さ -----------------------------------------------------------
    def test_height_ok_sends_and_writes(self) -> None:
        d = self.measure(HIKI)
        self.assertEqual(d["view"]["mount"]["total"], "150.00")
        self.result(OK_TEXT)
        d = self.post("press")
        self.assertEqual(d["question"]["key"], "ainput_send")
        d = self.ok(self.post("press", answers={"ainput_send": True}))
        self.assertTrue(d["view"]["check"]["ok"])
        s = d["view"]["send"]
        self.assertEqual(s["stage"], "sent", s)
        self.assertTrue(s["csv_path"] and s["hostsend_path"])
        self.assertEqual(len(self.launched), 1)

    def test_height_short_is_ng(self) -> None:
        self.measure(HIKI, "140.000")
        d = self.ok(self.post("press"))
        c = d["view"]["check"]
        self.assertFalse(c["ok"])
        self.assertTrue(c["msg"].startswith("[枚数不足(?)]"))
        self.assertEqual(d["view"]["send"]["steps"][1]["state"], "ng")
        self.assertEqual(self.launched, [])

    # 3) EX -------------------------------------------------------------
    def test_ex_toggle_by_order_no(self) -> None:
        self.ok(self.post("open", ken_m="L1"))
        v = self.ok(self.post("value", name="hiki_no", value=HIKI_EX))["view"]
        self.assertEqual(v["lot"]["order_no"], "01177194")
        self.assertTrue(v["show"]["ex"])
        v = self.ok(self.post("value", name="hiki_no", value=HIKI))["view"]
        self.assertFalse(v["show"]["ex"])
        v = self.ok(self.post("value", name="hiki_no", value="9999"))["view"]   # 一覧外: EX はそのまま
        self.assertFalse(v["show"]["ex"])

    def test_ex_remain_sets_outside(self) -> None:
        self.ok(self.post("open", ken_m="L1"))
        self.ok(self.post("value", name="hiki_no", value=HIKI_EX))
        f = self.ok(self.post("value", name="remain", value="1"))["view"]["fields"]
        self.assertEqual((f["outside_atu"], f["outside_haba"], f["outside_take"]), ("0", "0", "0"))
        f = self.ok(self.post("value", name="remain", value="0"))["view"]["fields"]
        self.assertEqual((f["outside_atu"], f["outside_haba"], f["outside_take"]), ("", "", ""))

    def test_non_ex_remain_keeps_outside(self) -> None:
        self.ok(self.post("open", ken_m="L1"))
        self.ok(self.post("value", name="hiki_no", value=HIKI))
        f = self.ok(self.post("value", name="remain", value="1"))["view"]["fields"]
        self.assertEqual(f["outside_atu"], "")

    def test_ex_input_checks_then_send(self) -> None:
        self.measure(HIKI_EX)
        d = self.ok(self.post("press"))
        self.assertEqual(d["view"]["send"]["note"], "ｹｰｽNoの入力がありません。")
        self.ok(self.post("value", name="ex_case_no", value="3"))
        for n, v in (("outside_atu", "5"), ("outside_haba", "1300"), ("outside_take", "2600")):
            self.ok(self.post("value", name=n, value=v))
        d = self.ok(self.post("press"))
        self.assertEqual(d["view"]["send"]["note"], "EX外寸に在り得ない数値の入力があります。")
        self.ok(self.post("value", name="outside_atu", value=""))
        # 空欄は「EX外寸の入力がありません」(2026-11-08 に直した。VBA は "" と 1〜9 の比較で実行時エラー 13 になり、
        # この案内に届かなかった)
        d = self.ok(self.post("press"))
        self.assertEqual(d["view"]["send"]["note"], "EX外寸の入力がありません｡")
        for n, v in (("ex_case_no", "3"), ("outside_atu", "200"), ("outside_haba", "1300"), ("outside_take", "2600")):
            self.ok(self.post("value", name=n, value=v))
        self.result(OK_TEXT)
        d = self.post("press")
        self.assertEqual(d["question"]["key"], "ainput_send")
        d = self.ok(self.post("press", answers={"ainput_send": True}))
        self.assertEqual(d["view"]["send"]["stage"], "sent")
        name = d["view"]["send"]["csv_path"].replace("\\", "/").split("/")[-1]
        hits = [p for p in self.tmp.rglob("*.csv") if p.name.endswith(name.split("_")[-1])] or \
            list(self.tmp.rglob("*H6X1355*.csv"))
        self.assertTrue(hits, d["view"]["send"]["csv_path"])
        csv = hits[0].read_bytes().decode("cp932")
        self.assertIn("25_EXｹｰｽNo,", csv)
        self.assertIn("2600", csv)

    def test_non_ex_skips_ex_checks(self) -> None:
        self.measure(HIKI)
        self.result(OK_TEXT)
        self.assertEqual(self.post("press")["question"]["key"], "ainput_send")

    def test_press_after_runtime_error_not_stuck(self) -> None:
        # 2026-11-08 に直した: 途中で実行時エラーになっても、次のボタンは効く(前は Tag="Load" が残り黙って効かなかった)
        from unittest import mock
        from ufkensa.kensa import ainput, vbx
        self.measure(HIKI_EX)
        self.ok(self.post("value", name="ex_case_no", value="3"))
        with mock.patch.object(ainput, "_in_1_to_9", side_effect=vbx.TypeMismatch("試験")):
            self.ok(self.post("value", name="outside_atu", value="200"))
            self.assertIsNotNone(self.post("press")["stop"])        # 実行時エラーで止まる
        for n, v in (("outside_atu", "200"), ("outside_haba", "1300"), ("outside_take", "2600")):
            self.ok(self.post("value", name=n, value=v))
        self.result(OK_TEXT)
        self.assertIsNotNone(self.post("press")["question"])

    def test_ex_store_hides_case_no(self) -> None:
        # VBA Remain_Change(UFAinput:551): EX 表示中のストア(1)は EXｹｰｽNo 欄を隠す(2026-11-08 に足した)
        self.ok(self.post("open", ken_m="L1"))
        self.ok(self.post("value", name="hiki_no", value=HIKI_EX))
        v = self.ok(self.post("value", name="remain", value="1"))["view"]
        self.assertFalse(v["show"]["ex_case_no"])


# 親(test_screen_ainput)のテストはここでは走らせない(setUp とヘルパーだけ使う)
for _n in [n for n in vars(AinputScreenTest) if n.startswith("test_")]:
    if _n not in vars(FeatureAinputTest):
        setattr(FeatureAinputTest, _n, None)


if __name__ == "__main__":
    unittest.main()
