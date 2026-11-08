"""確認フォーム(UFConfi)の挙動を Session で端から端まで

1) 測定値のダブルクリック消去 → 理由選択ページ(ConfiReason_Open O:646-672、LabelCls:166-196、UFConfi:16-88)
2) 公差外 → 公差外ページ(Out_Tolerance O:818-977 → Confi_Open O:673-715。厚板は Rig 4・10)
3) 公差外/再測定/許容/保留 の後(UFConfi.CommandButton1_Click UFConfi:171-273、NG_Back/NG_Pass/NG_Erase/NG_Posting)
"""
from __future__ import annotations

import unittest
from unittest import mock

from ufkensa.kensa import legacy_log

from tests.test_kensa_session import SessionBase

NG = "DT10001+0000015.500M"     # 上限 15.350 超え
OK = "DT10001+0000015.012M"


class ConfiBase(SessionBase):
    def setUp(self) -> None:
        super().setUp()
        self.logs: list[str] = []
        orig = legacy_log.LogEntry.write

        def cap(entry, *a, **k):
            self.logs.append(entry.msg if hasattr(entry, "msg") else str(entry))
            return orig(entry, *a, **k)
        p = mock.patch.object(legacy_log.LogEntry, "write", autospec=True, side_effect=cap)
        p.start()
        self.addCleanup(p.stop)
        self.register()
        self.st = self.s.state

    def ng_row_s(self):
        """S1 を公差外、S2〜S4 を合格で入れる(S4 で Out_Tolerance が走る)。"""
        self.s.receive(NG)
        r = None
        for _ in range(3):
            r = self.s.receive(OK)
        return r

    def log_has(self, text: str) -> bool:
        return any(text in m for m in self.logs)


class EraseReasonTest(ConfiBase):
    """1) ダブルクリックで消した場合"""

    def test_dblclick_opens_reason_page(self) -> None:
        s, st = self.s, self.st
        s.receive(OK)                                     # S1 = 15.012
        r = self.ok(s.dblclick("S1"))
        self.assertEqual(st.cells["S1"].value, "")       # 先に消える(LabelCls:173)
        self.assertNotEqual(st.cells["S1"].color, "NG")
        self.assertEqual((st.delete_p, st.delete_c), ("S1", "1"))   # EraseCount
        c = r.view["confi"]
        self.assertIsNotNone(c)                           # 1 回目から理由(2024.2.6 品証指示)
        self.assertEqual((c["page"], c["posi"], c["n"], c["moto"]), (1, "S", 1, "15.012"))
        self.assertEqual(c["choices"], ["入力位置間違い", "測定位置間違い", "測定ボタン押し間違い", "再測定",
                                        "0点確認", "板廃棄", "その他"])
        self.assertNotIn("検査ﾃﾞｰﾀに公差外有り！", [m.text for m in r.messages])
        self.assertEqual(st.inp_tag, "S1")                # 次はまた S1 に入る

    def test_blank_dblclick_no_form(self) -> None:
        r = self.ok(self.s.dblclick("S3"))
        self.assertIsNone(r.view["confi"])

    def test_reason_required_and_other_text(self) -> None:
        s = self.s
        s.receive(OK)
        s.dblclick("S1")
        r = s.set_value("UFConfi.List1", "")
        self.assertIn("選択してください", [m.text for m in r.messages])
        self.assertIsNotNone(r.view["confi"])
        r = s.set_value("UFConfi.List1", "その他")       # その他・再測定 は文字が要る(UFConfi:36-39)
        self.assertIn("理由を入力してください", [m.text for m in r.messages])
        self.assertIsNotNone(r.view["confi"])
        r = s.set_value("UFConfi.List1", "再測定")
        self.assertIn("理由を入力してください", [m.text for m in r.messages])
        self.ok(s.set_value("UFConfi.その他", "治具ずれ"))
        r = self.ok(s.set_value("UFConfi.List1", "その他"))
        self.assertIsNone(r.view["confi"])
        self.assertTrue(self.log_has("S1_15.012:理由選択:その他, 治具ずれ:"))

    def test_simple_reason_logs_and_closes(self) -> None:
        s = self.s
        s.receive(OK)
        s.dblclick("S1")
        r = self.ok(s.set_value("UFConfi.List1", "入力位置間違い,測定位置間違い"))
        self.assertIsNone(r.view["confi"])
        self.assertTrue(self.log_has("S1_15.012:理由選択:入力位置間違い, 測定位置間違い:"))
        self.assertEqual(self.st.cells["S1"].value, "")
        r = self.ok(s.receive(OK))                        # 次の伝送は消した S1 へ
        self.assertEqual(self.st.cells["S1"].value, "15.012")

    def test_dblclick_ng_value_clears_ng_label(self) -> None:
        s, st = self.s, self.st
        self.ng_row_s()
        s.set_value("UFConfi.TB1", "許容")
        self.assertEqual(st.ng["S"].value, "許容")
        r = s.dblclick("S1")                              # NG 色が無くなると NGS も空(NGErase)
        self.assertEqual(st.ng["S"].value, "")
        self.assertEqual(r.view["confi"]["page"], 1)

    def test_count_same_coordinate(self) -> None:
        s, st = self.s, self.st
        s.receive(OK)
        s.dblclick("S1")
        s.set_value("UFConfi.List1", "入力位置間違い")
        s.receive(OK)
        s.dblclick("S1")
        self.assertEqual(st.delete_c, "2")
        self.assertEqual(s._view()["confi"]["del_count"], 2)

    def test_ita_haiki_restores_and_posts(self) -> None:
        s, st = self.s, self.st
        s.receive(OK)
        s.dblclick("S1")
        r = self.ok(s.set_value("UFConfi.List1", "板廃棄"))
        self.assertIsNone(r.view["confi"])
        self.assertTrue(self.store.exists("異常ロット用_HVC"))   # NG_Posting
        self.assertEqual(st.v("IzyouFl").color, "255")
        self.assertTrue(self.log_has("S1_15.012:理由選択:板廃棄"))

    def test_reload_keeps_reason_form(self) -> None:
        s = self.s
        s.receive(OK)
        s.dblclick("S1")
        r = self.ok(s.start())                            # 画面の開き直し(2 回目の Initialize は作り直さない)
        self.assertEqual(r.view["confi"]["page"], 1)


class OutToleranceTest(ConfiBase):
    """2) 公差外の確認が出るとき"""

    def test_not_before_rig4(self) -> None:
        s = self.s
        r = s.receive(NG)
        self.assertIsNone(r.view["confi"])
        r = s.receive(OK)
        r = s.receive(OK)
        self.assertIsNone(r.view["confi"])                # Rig 2,3 は見ない
        r = s.receive(OK)
        c = r.view["confi"]
        self.assertEqual((c["page"], c["posi"], c["n"], c["moto"]), (0, "S", 1, "15.500"))
        self.assertEqual(c["choices"], ["公差外", "再測定", "許容", "保留"])
        self.assertIn("検査ﾃﾞｰﾀに公差外有り！", [m.text for m in r.messages])

    def test_no_confi_without_ng(self) -> None:
        for _ in range(4):
            r = self.ok(self.s.receive(OK))
        self.assertIsNone(r.view["confi"])

    def test_bad_text(self) -> None:
        self.ng_row_s()
        r = self.s.set_value("UFConfi.TB1", "あいう")
        self.assertIn("指定文字列を入力してください", [m.text for m in r.messages])
        self.assertIsNotNone(r.view["confi"])
        self.assertTrue(self.log_has("*UFConfi.TB1入力:あいう"))


class AfterChoiceTest(ConfiBase):
    """3) 処理入力後"""

    def test_kyoyou(self) -> None:
        s, st = self.s, self.st
        self.ng_row_s()
        r = self.ok(s.set_value("UFConfi.TB1", "許容"))
        self.assertEqual(st.ng["S"].value, "許容")
        self.assertEqual((st.cells["S1"].value, st.cells["S1"].color), ("15.500", "NG"))   # 値・色は残す
        self.assertIsNone(r.view["confi"])
        self.assertTrue(self.log_has("*UFConfi.TB1入力:許容"))
        r = self.ok(s.receive(OK))                        # 許容の行はもう聞かない
        self.assertIsNone(r.view["confi"])

    def test_hold_blocks_minmax(self) -> None:
        s, st = self.s, self.st
        self.ng_row_s()
        self.ok(s.set_value("UFConfi.TB1", "保留"))
        self.assertEqual(st.ng["S"].value, "保留")
        s.set_value("OPS1", "True")
        s.set_value("OPE1", "True")
        r = s.min_max()
        self.assertEqual(r.stop.text if r.stop else None, "許容でない公差外が存在します")

    def test_remeasure_first(self) -> None:
        s, st = self.s, self.st
        self.ng_row_s()
        r = self.ok(s.set_value("UFConfi.TB1", "再測定"))
        self.assertEqual(st.cells["S1"].value, "")
        self.assertEqual(st.cells["S4"].value, "15.012")
        self.assertEqual((st.delete_p, st.delete_c), ("S1", "1"))
        self.assertIsNone(r.view["confi"])
        self.assertEqual(st.ng["S"].value, "")
        self.assertEqual(st.inp_tag, "S1")                # Look_Blank: 空いた S1 へ
        self.assertTrue(self.log_has("*UFConfi.TB1入力:再測定"))

    def test_remeasure_twice_needs_reason(self) -> None:
        s, st = self.s, self.st
        self.ng_row_s()
        s.set_value("UFConfi.TB1", "再測定")              # DeleteC = 1
        r = s.receive(NG)                                 # S1 へ再び NG(S4 は入っているので Rig は 1)
        # Out_Tolerance は Rig 4/10 の入力でだけ動くので S1 では出ない(VBA どおり)
        self.assertIsNone(r.view["confi"])
        st.inp_tag = "S4"
        s.confi = None
        # S4 を消して入れ直し → 確認(DelC は S4 の座標で数えるので 1)
        s.dblclick("S4")
        s.set_value("UFConfi.List1", "入力位置間違い")
        r = s.receive(OK)
        c = r.view["confi"]
        self.assertEqual(c["page"], 0)
        # 同じ座標を 2 回以上消している状態で再測定 → List2(理由)
        s.confi["del_count"] = 2
        r = s.set_value("UFConfi.TB1", "再測定")
        self.assertIn("理由を選択してください", [m.text for m in r.messages])
        self.assertTrue(r.view["confi"]["list2"])
        self.assertEqual(r.view["confi"]["choices"], ["入力位置間違い", "測定位置間違い", "測定ボタン押し間違い",
                                                      "0点確認", "その他"])
        r = s.set_value("UFConfi.List2", "")
        self.assertIn("選択してください", [m.text for m in r.messages])
        r = self.ok(s.set_value("UFConfi.List2", "0点確認"))
        self.assertIsNone(r.view["confi"])
        self.assertEqual(st.cells["S1"].value, "")
        self.assertTrue(self.log_has("公差外:S1_15.500:理由選択:再測定,0点確認"))

    def test_koshagai_posts_and_erases(self) -> None:
        s, st = self.s, self.st
        self.ng_row_s()
        r = self.ok(s.set_value("UFConfi.TB1", "公差外"))
        self.assertIsNone(r.view["confi"])
        self.assertTrue(self.store.exists("異常ロット用_HVC"))
        self.assertEqual(st.v("IzyouFl").color, "255")
        for i in range(1, 5):                             # 作業保存の S 行を消して読み直す
            self.assertEqual(st.cells[f"S{i}"].value, "", f"S{i}")
        self.assertTrue(self.log_has("*UFConfi.TB1入力:公差外"))

    def test_last_request_wins_and_reload(self) -> None:
        s = self.s
        self.ng_row_s()
        r = self.ok(s.start())
        self.assertEqual(r.view["confi"]["page"], 0)
        s2 = self.new_session()                           # 再起動(新しい Session)では確認は戻らない
        r = s2.start()
        self.assertEqual(s2.state.cells["S1"].value, "15.500")
        self.assertIsNone(r.view["confi"])

    def test_no_confi_input_when_closed(self) -> None:
        r = self.s.set_value("UFConfi.TB1", "許容")
        self.assertEqual(r.stop.text, "公差外の確認は出ていません")


if __name__ == "__main__":
    unittest.main()
