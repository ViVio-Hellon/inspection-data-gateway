"""業務の流れの試験(STEP 4)。画面と同じ HTTP(/api/kensa/*)を通し、**本番の不具合の扱い(config/quirks.json)** で動かす

ほかの試験は「VBA どおり再現」で書いてあるので、本番で通る「直す」側の道が端から端まで試されていなかった。
ここでは現場の操作の順(カード → ロット → 作業選択 → 山数 → 測定 → 歪・外観 → Min/Max → 印刷前チェック → 印刷・保存
→ 開き直し → 新規検査)と、よくある異常(空欄・桁違い・無いロット・範囲外の値・未校正・前の直の検査員・2 つ目のタブ)を通す。
画面がやるのと同じく、確認(question)には決めた答えを付けて送り直す。
"""
from __future__ import annotations

import json
import unittest
from datetime import datetime, timedelta
from typing import Any, Optional

from tests import test_kensa_session as base
from tests._web import make_client
from ufkensa import config
from ufkensa.kensa import quirks, session, store

LOT, CARD, WL = base.LOT, base.CARD, base.WL


def production_quirks() -> dict[str, str]:
    data = json.loads((config.BASE_DIR / "config" / "quirks.json").read_text(encoding="utf-8"))
    return {k: v for k, v in data.items() if k.startswith("B")}


class FlowBase(base.SessionBase):
    """SessionBase の DB・時計・校正に、本番の不具合の扱いと HTTP の入口を足す。"""

    def setUp(self) -> None:
        quirks.set_for_test(production_quirks())
        self.addCleanup(quirks.set_for_test, None)
        super().setUp()
        session.reset_for_test(self.s)
        self.addCleanup(session.reset_for_test, None)
        self.c = make_client()
        self.h = {"X-Tool-Token": "test-token-abc123"}
        self.answers: dict[str, bool] = {}           # 確認への答え(無いものは「はい」)

    # --- 画面の op() と同じ ------------------------------------------------
    def op(self, op_name: str, **body: Any) -> dict:
        answers: dict[str, bool] = {}
        for _ in range(8):
            r = self.c.post(f"/api/kensa/{op_name}", json={**body, "answers": answers}, headers=self.h)
            self.assertLess(r.status_code, 500, r.get_data(as_text=True)[:300])
            d = r.get_json()
            q = d.get("question")
            if not q:
                return d
            answers[q["key"]] = self.answers.get(q["key"], True)
        self.fail(f"{op_name}: 確認が終わらない {answers}")

    def texts(self, d: dict) -> list[str]:
        return [m.get("text", "") for m in d.get("messages", [])] + ([d["stop"]["text"]] if d.get("stop") else [])

    def ready(self) -> dict:
        self.op("start")
        self.op("worker", card=CARD)
        d = self.op("lot", lot=LOT.lower())
        self.assertEqual(d["view"]["header"]["lot"], LOT, self.texts(d))
        return d

    def measure_all(self) -> int:
        """「次に測る」に従って伝送文字列を送り、厚・幅丈対角が検査済になるまで。送った数を返す。"""
        sent = 0
        for kind in ("DT", "#L"):
            for _ in range(60):
                nxt = self.s.state.inp_tag
                text = "DT10001+0000015.012M" if kind == "DT" else "#L" + WL.get(nxt[-1:], "1252.1")
                d = self.op("receive", text=text)
                sent += 1
                self.assertIsNone(d.get("stop"), self.texts(d))
                if any(t == kind + "検査済" for t in self.texts(d)):
                    self.op("click", name="")
                    break
            else:
                self.fail(kind + " が検査済にならない")
        return sent

    def strain_visual_height(self) -> None:
        for p in "SUCDE":
            for n, v in ((f"CB{p}1", "幅"), (f"CB{p}2", "上"), (f"Z{p}", "0.3"), (f"CB{p}3", "幅"), (f"CB{p}4", "上"),
                         (f"G{p}", "0.2"), (f"{p}_UpVisual", "欠陥ﾅｼ"), (f"{p}_UnderVisual", "欠陥ﾅｼ")):
                d = self.op("value", name=n, value=v)
                self.assertIsNone(d.get("stop"), (n, self.texts(d)))
        for n, v in (("KenHBox1", "200"), ("OPS1", "True"), ("OPE5", "True")):
            self.op("value", name=n, value=v)


class NormalFlowTest(FlowBase):
    def test_one_lot_through_http(self) -> None:
        """厚板 1 ロット: 登録 → 測定 → 印刷・保存 → 開き直して同じ画面 → 新規検査で空。"""
        d = self.ready()
        v = d["view"]
        self.assertEqual((v["header"]["zai"], v["header"]["youto"]), ("995", "H165"))
        self.assertEqual(v["tolerance"]["apk"], "15.350")
        d = self.op("work", work="厚板")
        self.assertEqual(d["view"]["header"]["work"], "厚板")
        self.assertGreater(self.measure_all(), 10)
        cells = self.s.state.cells
        self.assertEqual([cells[f"S{i}"].value for i in range(1, 5)], ["15.012"] * 4)
        self.assertEqual(cells["S1"].color, "OK")
        self.strain_visual_height()
        d = self.op("minmax")
        self.assertEqual(d["view"]["minmax"]["amin"], "15.012")
        d = self.op("print-check")
        self.assertTrue(d["view"]["print_ready"], [c for c in d["view"]["checks"] if c["ok"] is False])
        d = self.op("print-commit")
        self.assertEqual(d["view"]["report"]["saved"]["mode"], "ADDNEW", self.texts(d))
        row = self.store.sheet(store.HOZON).row(2)
        self.assertEqual((row[1], row[11]), (LOT, "15.012"))

        # 再起動(セッションを作り直す)→ 画面を開く(start)→ 同じロット・同じ値
        session.reset_for_test(self.new_session())
        self.s = session.get()
        d = self.op("start")
        v = d["view"]
        self.assertEqual((v["header"]["lot"], v["header"]["work"]), (LOT, "厚板"))
        self.assertEqual(self.s.state.cells["S1"].value, "15.012")
        # 新規検査 → 空
        d = self.op("new")
        self.assertEqual((d["view"]["header"]["lot"], self.s.state.cells["S1"].value), ("", ""))

    def test_second_tab_does_not_reset(self) -> None:
        """2 つ目のタブで検査画面を開いても(start)、測りかけの値は消えない。"""
        self.ready()
        self.op("work", work="厚板")
        self.op("receive", text="DT10001+0000015.012M")
        first = self.s.state.cells["S1"].value
        self.assertEqual(first, "15.012")
        other = make_client()
        r = other.post("/api/kensa/start", json={"answers": {}}, headers=self.h)
        self.assertEqual(r.status_code, 200)
        self.assertEqual((self.s.state.lot, self.s.state.cells["S1"].value), (LOT, "15.012"))

    def test_interrupted_and_restarted(self) -> None:
        """測りかけで再起動 → 測った値も、次に測る欄も戻る。"""
        self.ready()
        self.op("work", work="厚板")
        for _ in range(3):
            self.op("receive", text="DT10001+0000015.012M")
        nxt = self.s.state.inp_tag
        vals = {k: c.value for k, c in self.s.state.cells.items() if c.value}
        session.reset_for_test(self.new_session())
        self.s = session.get()
        self.op("start")
        self.assertEqual({k: c.value for k, c in self.s.state.cells.items() if c.value}, vals)
        # 次に測る欄は開き直した直後は空(VBA も InpTag は保存しない)。次の値は自動で続きの欄に入る
        d = self.op("receive", text="DT10001+0000015.012M")
        self.assertIsNone(d.get("stop"), self.texts(d))
        self.assertEqual(self.s.state.cells[nxt].value, "15.012")


class RecoveryTest(FlowBase):
    def test_unexpected_error_does_not_freeze(self) -> None:
        """操作の途中で想定外のエラー → 知らせて止まるが、旗(Load・FormTag)は戻り、次の測定は入る。"""
        from unittest import mock
        self.ready()
        self.op("work", work="厚板")
        st = self.s.state
        def boom(*a, **k):
            st.load, st.form_tag = "off", "STOP"
            raise ValueError("試験のエラー")
        from ufkensa.kensa import minmax
        with mock.patch.object(minmax, "run_min_max", boom):
            d = self.op("minmax")
        self.assertIn("試験のエラー", d["stop"]["text"])
        self.assertEqual((st.load, st.form_tag), ("on", ""))
        d = self.op("receive", text="DT10001+0000015.012M")
        self.assertIsNone(d.get("stop"), self.texts(d))
        self.assertEqual(st.cells["S1"].value, "15.012")


class StuckFlagTest(FlowBase):
    def test_left_flag_does_not_block_measurement(self) -> None:
        """前の操作が途中で止まって FlagLabel=0・Load=off が残っても、次の測定は入る(黙って捨てない)。"""
        self.ready()
        self.op("work", work="厚板")
        st = self.s.state
        st.flag_label, st.load = 0, "off"
        d = self.op("receive", text="DT10001+0000015.012M")
        self.assertIsNone(d.get("stop"), self.texts(d))
        self.assertEqual(st.cells["S1"].value, "15.012")

    def test_reopen_keeps_out_of_tolerance_question(self) -> None:
        """画面を開き直しても(start)、出ている公差外の確認は消えない。"""
        self.ready()
        self.op("work", work="厚板")
        self.op("receive", text="DT10001+0000015.400M")          # 公差外
        confi = self.s.confi
        self.op("start")
        self.assertEqual(self.s.confi, confi)


class ReceiverRoutingTest(FlowBase):
    def test_value_goes_to_open_screen_only(self) -> None:
        """受信機の値は、いま開いている画面にだけ渡す。開いていなければ捨てる(見えないまま検査に入れない)。"""
        from app.routes import kensa as rk
        self.ready()
        self.op("work", work="厚板")
        rk._recv_seen.clear()
        rk._to_kensa("DT10001+0000015.012M")                     # どの画面も開いていない → 捨てる
        self.assertEqual(self.s.state.cells["S1"].value, "")
        self.c.get("/api/kensa/recv?since=0&screen=calibration", headers=self.h)   # 校正画面を開いている
        rk._to_kensa("DT10001+0000015.012M")
        self.assertEqual(self.s.state.cells["S1"].value, "")      # 検査には入らない
        d = self.c.get("/api/kensa/recv?since=0&screen=calibration", headers=self.h).get_json()
        self.assertEqual([r["text"] for r in d["results"]], ["DT10001+0000015.012M"])   # 校正画面が受け取る
        self.c.get("/api/kensa/recv?since=0&screen=calibration&visible=0", headers=self.h)  # 校正を閉じた
        self.c.get("/api/kensa/recv?since=0&screen=kensa", headers=self.h)                 # 検査を開いた
        rk._to_kensa("DT10001+0000015.012M")
        self.assertEqual(self.s.state.cells["S1"].value, "15.012")


class SaveIntegrityTest(FlowBase):
    def test_posting_is_all_or_nothing(self) -> None:
        """作業中保存の途中で落ちても、前に保存した値は残る(1 セルずつ確定していたのを直した)。"""
        from unittest import mock
        from ufkensa.kensa import store as stmod
        self.ready()
        self.op("work", work="厚板")
        for _ in range(3):
            self.op("receive", text="DT10001+0000015.012M")
        saved = self.store.sheet("作業保存_HVC").used()
        real = stmod.GridSheet.set
        n = {"i": 0}
        def flaky(sh, r, c, v):
            if sh.name == "作業保存_HVC":
                n["i"] += 1
            if n["i"] == 50:
                raise OSError("試験: 書き込み失敗")
            return real(sh, r, c, v)
        with mock.patch.object(stmod.GridSheet, "set", flaky):
            d = self.op("receive", text="DT10001+0000015.012M")
        self.assertTrue(any("書き込み失敗" in t for t in self.texts(d)), self.texts(d))
        self.assertEqual(self.store.sheet("作業保存_HVC").used(), saved)       # 途中までの書きかけは残らない
        self.assertGreater(n["i"], 49)

    def test_change_after_check_requires_recheck(self) -> None:
        """印刷前チェックの後に値を変えたら、確定の前にもう一度チェックする。"""
        self.ready()
        self.op("work", work="厚板")
        self.measure_all()
        self.strain_visual_height()
        self.op("minmax")
        self.assertTrue(self.op("print-check")["view"]["print_ready"])
        self.op("value", name="ZS", value="")                    # 歪の値を消す
        self.assertFalse(self.s.print_ready)

    def test_worker_mismatch_stops_print(self) -> None:
        """検査員と測定者が一致しないと、印刷・保存に進まず名簿・日報へ案内する(VBA は UFNameList で止まる)。"""
        from unittest import mock
        from ufkensa.kensa import workers
        self.ready()
        self.op("work", work="厚板")
        self.measure_all()
        self.strain_visual_height()
        self.op("minmax")
        real = workers.name_match
        def mismatch(*a, **k):
            r = real(*a, **k)
            r.out = True
            return r
        with mock.patch.object(workers, "name_match", mismatch):
            d = self.op("print-commit")
        self.assertIn("一致しない", " ".join(self.texts(d)))
        self.assertEqual(d["view"]["name_list"], 1)
        self.assertEqual(self.store.sheet(store.HOZON).last_row(1), 1)     # 保存していない

    def test_missing_measure_db_stops_before_save(self) -> None:
        self.ready()
        self.op("work", work="厚板")
        self.measure_all()
        self.strain_visual_height()
        self.op("minmax")
        self.s.measure_db = None                                 # 本番と同じく使用パスの測定DB を見る
        from ufkensa.kensa import paths
        from unittest import mock
        with mock.patch.object(paths, "db_file", return_value=str(self.tmp / "無い" / "梱包検査ﾃﾞｰﾀ.sqlite3")):
            d = self.op("print-commit")
        self.assertIn("測定DB", d["stop"]["text"])
        self.assertFalse((self.tmp / "無い").exists())           # 黙って作らない
        self.assertEqual(self.store.sheet(store.HOZON).last_row(1), 1)


class AbnormalInputTest(FlowBase):
    def test_card_and_lot_wrong_length(self) -> None:
        self.op("start")
        d = self.op("worker", card="900000000000")             # 12 桁
        self.assertEqual(d["view"]["workers"], [])
        d = self.op("worker", card="1234567890123")            # 名簿に無い
        self.assertEqual(d["view"]["workers"], [])
        self.assertTrue(any("名簿" in t for t in self.texts(d)), self.texts(d))
        self.op("worker", card=CARD)
        d = self.op("lot", lot="H7345B")                        # 6 桁
        self.assertEqual(d["view"]["header"]["lot"], "")
        d = self.op("lot", lot="")
        self.assertEqual(d["view"]["header"]["lot"], "")

    def test_wrong_length_keeps_current_lot(self) -> None:
        """8 桁などを入れても、いまのロット・公差はそのまま(前は黙って作業保存のロットが空になった)。"""
        self.ready()
        d = self.op("lot", lot="H7345B00")
        self.assertEqual((d["view"]["header"]["lot"], d["view"]["tolerance"]["apk"]), (LOT, "15.350"))
        self.assertTrue(any("7 桁" in t for t in self.texts(d)), self.texts(d))
        self.assertEqual(self.store.sheet("作業保存_HVC").get(1, 1), LOT)

    def test_recent_lot_no_keeps_everything(self) -> None:
        """「直近5Lot内に同じLotNo」でいいえ → 何も変えない(前は新しい番号に前のロットの公差が残った)。"""
        self.ready()
        sh = self.store.ensure_sheet(store.HOZON)
        sh.set(2, 1, "2026/11/01 10:00～10:30")
        sh.set(2, 2, "RX021T0")
        self.answers["lot_recent"] = False
        d = self.op("lot", lot="RX021T0")
        self.assertEqual((d["view"]["header"]["lot"], d["view"]["tolerance"]["apk"]), (LOT, "15.350"))

    def test_lot_not_in_shikake(self) -> None:
        self.op("start")
        self.op("worker", card=CARD)
        d = self.op("lot", lot="Z9999Z9")
        self.assertTrue(any("Z9999Z9" in t and "仕掛" in t for t in self.texts(d)), self.texts(d))

    def test_garbage_and_out_of_tolerance(self) -> None:
        self.ready()
        self.op("work", work="厚板")
        before = {k: c.value for k, c in self.s.state.cells.items()}
        d = self.op("receive", text="XYZ")                      # 頭が DT / #L ではない
        self.assertEqual({k: c.value for k, c in self.s.state.cells.items()}, before)
        d = self.op("receive", text="DT10001+0000016.500M")      # 公差の 2 倍の外 = 異常値(入れずに知らせる)
        self.assertTrue(any("異常値" in t for t in self.texts(d)), self.texts(d))
        self.assertEqual({k: c.value for k, c in self.s.state.cells.items()}, before)
        d = self.op("receive", text="DT10001+0000015.400M")      # 公差 14.650〜15.350 の外(異常値ではない)
        self.assertIsNone(d.get("stop"), self.texts(d))
        hit = [k for k, c in self.s.state.cells.items() if c.value == "15.400"]
        self.assertEqual(len(hit), 1)
        self.assertEqual(self.s.state.cells[hit[0]].color, "NG")
        d = self.op("print-check")
        self.assertFalse(d["view"]["print_ready"])

    def test_blank_values(self) -> None:
        self.ready()
        self.op("work", work="厚板")
        for n in ("ZS", "GS", "KenHBox1", "S_UpVisual"):
            d = self.op("value", name=n, value="")
            self.assertIsNone(d.get("stop"), (n, self.texts(d)))


class CalibrationAndShiftTest(FlowBase):
    calibrated = False

    def test_uncalibrated_lot_ok_measure_refused(self) -> None:
        d = self.ready()                                         # ロットは登録できる(2026-11-06)
        self.op("work", work="厚板")
        d = self.op("receive", text="DT10001+0000015.012M")
        self.assertTrue(any("校正" in t for t in self.texts(d)), self.texts(d))
        self.assertEqual([k for k, c in self.s.state.cells.items() if c.value], [])
        self.assertTrue(d["view"]["calibration_required"])

    def test_after_calibration_rejudges(self) -> None:
        """校正に合格して戻ったら(UF校正.INPG_Exit)、判定をやり直し、校正の案内を消す。"""
        self.ready()
        self.op("work", work="厚板")
        self.op("receive", text="DT10001+0000015.012M")          # 未校正なので受け付けない
        self.write_calibration()
        d = self.op("calibrated")
        self.assertIsNone(d.get("stop"), self.texts(d))
        self.assertFalse(d["view"]["calibration_required"])
        self.assertEqual(self.s.state.load, "on")
        d = self.op("receive", text="DT10001+0000015.012M")
        self.assertEqual(self.s.state.cells["S1"].value, "15.012")

    def test_previous_day_worker_not_kept(self) -> None:
        """前の日の検査員は、持ち越しの検査が無ければ残らない(2026-11-06)。"""
        self.write_calibration()
        self.op("start")
        self.op("worker", card=CARD)
        self.assertEqual(len(self.op("start")["view"]["workers"]), 1)
        self.clock.t += timedelta(days=1)
        self.write_calibration()
        session.reset_for_test(self.new_session())
        self.s = session.get()
        d = self.op("start")
        self.assertEqual(d["view"]["workers"], [])
        self.assertTrue(d["view"]["worker_needed"])


if __name__ == "__main__":
    unittest.main()
