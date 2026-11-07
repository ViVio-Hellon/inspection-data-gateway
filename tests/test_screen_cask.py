"""キャスク検査の画面 /cask と /api/cask/*(app/routes/cask.py → cask_session.py → modes/cask.py・sheet.py)

部品の判定そのものは test_kensa_modes_cask.py で確かめてある。ここは画面の組み立てと受け渡し:
開く(検査員の登録が要る)→ ロット → ゼロ校正 → シート上部の手入力 → 伝送文字列 → ○/× と次の入力セル →
作業中保存からの読み戻し → 新規発行の確認 → 印刷(帳票 HTML)・保存(D7: CSV+HTML)・保存用の行。
"""
from __future__ import annotations

import sqlite3
import unittest
from pathlib import Path
from unittest import mock

from app import shell
from tests._web import make_client
from tests.test_kensa_session import CARD, SessionBase
import functools
import shutil

from ufkensa.kensa import cask_session, packing, paths, store, vbx, wiring
from ufkensa.kensa.grid import KensaState
from ufkensa.kensa.modes import cask as caskmod
from ufkensa.kensa.modes import sheet as sh
from tests._isolation import vba_reproduce


@vba_reproduce
class CaskScreenTest(SessionBase):
    def setUp(self) -> None:
        super().setUp()
        self.s.start()
        r = self.s.register_worker(CARD, {"worker_register": True})
        self.assertIsNone(r.stop)
        self.cs = self.cask_session()
        cask_session.reset_for_test(self.cs)
        self.addCleanup(cask_session.reset_for_test, None)
        p = mock.patch.object(shell, "READY_SCREENS", shell.READY_SCREENS | {"cask"})
        p.start()
        self.addCleanup(p.stop)
        self.c = make_client()
        from app.routes import cask as cask_routes
        if "cask" not in self.c.application.blueprints:
            self.c.application.register_blueprint(cask_routes.bp)
        self.h = {"X-Tool-Token": self.c.application.config["TOKEN"]}

    def cask_session(self, uf1=None) -> cask_session.CaskSession:
        return cask_session.CaskSession(self.store, now=self.clock, uf1=uf1 or (lambda: self.s.state),
                                        write_logs=False)

    def post(self, op: str, **body):
        r = self.c.post(f"/api/cask/{op}", json=body, headers=self.h)
        self.assertEqual(r.status_code, 200, r.get_data(as_text=True))
        return r.get_json()

    @staticmethod
    def texts(d) -> list[str]:
        return [m["text"] for m in d["messages"] if m["text"]]

    @staticmethod
    def cell(d, row: int, col: int) -> list:
        """view.rows の (行, 列) のセル [文字, 判定, 状態]。"""
        cols = [c["col"] for c in d["view"]["columns"]]
        for r in d["view"]["rows"]:
            if r[0] == row:
                return r[2][cols.index(col)]
        raise AssertionError(f"行 {row} がありません")

    def ready(self) -> dict:
        """開く → ロット → ゼロ校正 → 上部(寸法・公差)を手入力。"""
        d = self.post("open")
        self.assertTrue(d["view"]["visible"], d)
        d = self.post("lot", lot="zz-zzzzz")
        self.assertEqual(d["view"]["form"]["lot_box"], "ZZZZZZZ")                     # KeyPress で英数字・大文字
        # 仕掛に無い → 上部を消して Q2:S4 の数式だけ置き直す(UFCask:520-531)。次の入力セルは B7
        self.assertIn("Text.データなし　手入力よろしく", self.texts(d))
        self.assertEqual((d["view"]["next"]["row"], d["view"]["next"]["col"]), (7, 2))
        self.assertEqual(d["view"]["focus"], "zero")                                   # ゼロ校正が未
        d = self.post("zero", text="DT10002+0000000.000M")
        self.assertEqual(d["view"]["form"]["micro"], "0.000")
        d = self.post("zero", text="#L0.0")
        self.assertEqual(d["view"]["form"]["convex"], "0.0")
        self.assertEqual(d["view"]["focus"], "meas")
        for (r, c), v in {(3, 2): "36", (3, 4): "1528", (3, 6): "3053", (2, 13): "0.4", (2, 15): "0.4"}.items():
            self.post("select", row=r, col=c)
            d = self.post("receive", text=v)
        tol = d["view"]["header"]["tolerance"][0]
        self.assertEqual((tol["size"], tol["lo"], tol["hi"]), ("36.000", "35.600", "36.400"))   # Q2:S2 の数式
        return d

    def test_page_and_state(self) -> None:
        r = self.c.get("/cask", headers=self.h)
        self.assertEqual(r.status_code, 200)
        self.assertIn("caskGrid", r.get_data(as_text=True))
        st = self.c.get("/api/cask/state", headers=self.h).get_json()
        self.assertFalse(st["view"]["visible"])
        self.assertEqual(st["view"]["total_rows"], 700)
        self.assertEqual([c["label"] for c in st["view"]["columns"]][:4], ["厚1", "厚2", "厚3", "厚4"])

    def test_lot_u4_default(self) -> None:                   # U4: 既定は読める(2026-11-06。全角の欄名でも半角の欄を読む)
        self.post("open")
        d = self.post("lot", lot="rx021t0")
        self.assertEqual(d["view"]["form"]["lot_box"], "RX021T0")
        self.assertNotIn("DB読み込みエラー:\r\n" + caskmod.FIELD_MISSING, self.texts(d))
        self.assertEqual(d["view"]["form"]["ken_no"], "T6X1913")                 # 仕掛から検査番号を読んだ
        self.assertEqual(d["view"]["form"]["count_st"], "3")                     # 製造板厚から検査回数を決めた

    def test_ainput_from_cask_uses_cask_lot(self) -> None:
        """キャスクの「Aインプットへ」: Aインプットは通常検査ではなく、キャスクのシートのロットで開く(Input:247-252)。"""
        from ufkensa.kensa import ainput_session
        self.post("open")
        self.post("lot", lot="rx021t0")
        a = ainput_session.AinputSession(self.s)
        res = a.open("L2", "Cask")
        self.assertIsNone(res.stop, res.stop and res.stop.text)
        v = a.view()
        self.assertEqual((v["source"], a.form.lot), ("Cask", "RX021T0"))
        self.assertNotEqual(a.form.lot, self.s.state.lot)                    # 通常検査のロットではない
        empty = ainput_session.AinputSession(self.s)
        cask_session.reset_for_test(self.cask_session())                     # キャスクでロットを開いていない
        self.assertIn("キャスク画面でロットを開いてから", empty.open("L2", "Cask").stop.text)

    def test_stuck_events_recover(self) -> None:
        """前の操作でイベントが止まったまま(EnableEvents = False の戻し忘れ)でも、次の操作でロットが読める。"""
        self.post("open")
        self.cs.book.events_enabled = False
        d = self.post("lot", lot="rx021t0")
        self.assertEqual(d["view"]["form"]["ken_no"], "T6X1913")

    def test_open_needs_worker(self) -> None:                # UF1.KYB_Click (UF1:3232-3234)
        cs = self.cask_session(uf1=lambda: KensaState())
        res = cs.open()
        self.assertFalse(res.view["visible"])
        self.assertIn("検査員登録をして下さい", [m.text for m in res.messages])
        self.assertIsNotNone(cs.receive("DT10002+0000036.000M").stop)              # 開くまでは入力できない

    def test_measure_judge_and_reload(self) -> None:
        self.ready()
        self.post("select", row=7, col=2)
        d = self.post("receive", text="DT10002+0000036.000M")
        self.assertEqual(self.cell(d, 7, 2)[:2], ["36", "○"])
        self.assertIn("ok", self.cell(d, 7, 2)[2])
        self.assertIn("cask_ok", [m["sound"] for m in d["messages"]])           # 音楽合格
        self.assertEqual(d["view"]["flash"]["mark"], "○")                          # UFcount
        self.assertEqual((d["view"]["next"]["row"], d["view"]["next"]["col"]), (7, 4))
        self.assertIn("sel", self.cell(d, 7, 4)[2])
        d = self.post("receive", text="DT10002+0000035.500M")
        self.assertEqual(self.cell(d, 7, 4)[:2], ["35.5", "×"])
        self.assertIn("cask_ng", [m["sound"] for m in d["messages"]])
        self.assertTrue(d["view"]["header"]["start"].startswith("2026/11/02"))     # D7 を選んだので開始日時
        # 作業中保存: 手元DB の ｷｬｽｸ作業中保存 に値、開き直すと読み戻す(フォームは Initialize でシートから)
        self.assertEqual(self.store.sheet(cask_session.WORK_SAVE).get(7, 2), 36.0)
        cs2 = self.cask_session()
        res = cs2.state()
        self.assertFalse(res.view["visible"])
        self.assertEqual(self.cell({"view": res.view}, 7, 4)[:2], ["35.5", "×"])
        res = cs2.open()
        self.assertTrue(res.view["visible"])
        self.assertEqual(res.view["header"]["lot"], "")                            # 仕掛に無いので D2 は空
        # 厚欄の手入力(DT でない)は消す(Cask:63-66)
        cs2.select(7, 6)
        res = cs2.receive("36.0")
        self.assertEqual(self.cell({"view": res.view}, 7, 6)[0], "")

    def test_locked_cell_and_resume(self) -> None:
        self.ready()
        ws = self.cs.book.ws(sh.CASK)
        ws.locked.add((7, 8))                                 # 3 点のとき H:I はロック(CountSelect)
        self.post("select", row=7, col=8)
        d = self.post("receive", text="DT10002+0000036.000M")
        self.assertIn(cask_session.PROTECTED_TEXT, self.texts(d))
        self.assertEqual(self.cell(d, 7, 8)[0], "")
        # 数でない文字が残ると EnableEvents=False のまま(VBA どおり)→ 処理再開で戻す
        self.post("inspection")
        self.post("select", row=7, col=28)
        d = self.post("receive", text="良")
        self.assertFalse(d["view"]["events"])
        d = self.post("resume")
        self.assertTrue(d["view"]["events"])

    def test_new_sheet_question(self) -> None:
        self.ready()
        self.post("select", row=7, col=2)
        self.post("receive", text="DT10002+0000036.000M")
        d = self.post("new")
        self.assertEqual(d["question"]["key"], "cask_new")
        d = self.post("new", answers={"cask_new": False})
        self.assertIn("ｷｬﾝｾﾙ処理", self.texts(d))
        d = self.post("new", answers={"cask_new": True})
        self.assertEqual(self.cell(d, 7, 2)[0], "")
        self.assertEqual(d["view"]["form"]["micro"], "")
        self.assertEqual(d["view"]["form"]["inspection"], "OFF")
        # JustinCase_KySave(D7: CSV と HTML)
        files = [Path(p) for p in d["view"]["saved"]["files"]]
        self.assertIn("キャスク検査", _xlsx_sheets(files))                      # 1 つの xlsx にシートを並べる

    def test_print_and_save(self) -> None:
        d = self.ready()
        for name, v in {"kname": "検査員01", "cno": "C1", "mno": "M1", "seri_no": "s9", "ken_no": "K1"}.items():
            d = self.post("field", name=name, value=v)
        self.assertEqual(d["view"]["form"]["seri_no"], "S9")
        r = self.c.post("/api/cask/field", json={"name": "nothing", "value": "x"}, headers=self.h).get_json()
        self.assertIsNotNone(r["stop"])
        # 梱包板番号: キャスク梱包板番号の [4](受領 DB の列順では 作業時間)とシート最終行の A 列を照合する(W15 は部品どおり)
        with sqlite3.connect(self.db_dir / "伝送用ファイル.sqlite3") as con:
            con.execute('INSERT INTO "キャスク梱包板番号" ("LotNo", "検査No", "スタートNo", "エンドNo", "作業時間", '
                        '"管理番号") VALUES (?, ?, ?, ?, ?, ?)', ("ZZZZZZZ", "K1", 1, 1, "1", 9001))
        self.post("select", row=7, col=2)
        for v in ("36.0", "36.1", "36.2"):
            d = self.post("receive", text="DT10002+00000" + v + "00M")
        before = self.store.sheet(store.HOZON).last_row(1)
        d = self.post("print")
        self.assertEqual(d["question"]["key"], "cask_print", d["messages"])
        d = self.post("print", answers={"cask_print": True})
        self.assertIsNone(d["stop"])
        rep = d["view"]["report"]
        self.assertTrue(rep["name"].endswith(".html"))
        html = self.c.get(f"/api/cask/report/{rep['name']}", headers=self.h).get_data(as_text=True)
        self.assertIn("厚公差：35.600～36.400", html)                                 # 右ヘッダー
        self.assertIn("梱包No1_1～1", html)                                          # 左フッター
        self.assertEqual(self.c.get("/api/cask/report/none.html", headers=self.h).status_code, 404)
        # キャスク専用保存(D7: シートごとの CSV + HTML、サーバ・工場共有へ写す)
        files = [Path(p) for p in d["view"]["saved"]["files"]]
        self.assertIn("品証ｺﾋﾟ-用", _xlsx_sheets(files))
        self.assertTrue(any("キャスク_履歴" in str(p) for p in files))
        self.assertTrue(all(p.is_file() for p in files))
        self.assertEqual(self.store.sheet(store.HOZON).last_row(1), before + 1)      # 保存用に履歴行
        self.assertEqual(self.store.sheet(store.HOZON).get(before + 1, 83), "ｷｬｽｸ")
        # 同じ日にもう一度保存すると連番(.002)
        d = self.post("save")
        self.assertTrue(any(".002_" in p for p in d["view"]["saved"]["files"]), d["view"]["saved"])

    # --- 包装仕様書(PackingNoCheck / UFPack_Open。UFCask:538) ------------------------------
    def lot_readable(self) -> None:
        """U4 を読める側にして(全角の欄名を幅を揃えて読む)、ロットを読み込めるようにする。"""
        p = mock.patch.object(caskmod, "combo_change",
                              functools.partial(caskmod.combo_change, fullwidth_fields_readable=True))
        p.start()
        self.addCleanup(p.stop)

    def pack_folder(self, *names: str) -> None:
        f = paths.to_local(packing.folder_path())
        f.mkdir(parents=True, exist_ok=True)
        for n in names:
            (f / n).write_bytes(b"\xff\xd8\xff\xe0")

    def test_packing_dialog(self) -> None:
        self.lot_readable()
        self.pack_folder("1P1216.JPG")
        self.post("open")
        d = self.post("lot", lot="rx021t0")
        rows = [r for r in (wiring.sheet_rows(self.store, store.LOT_HISTORY) or [])[1:] if r[2] == "RX021T0"]
        self.assertEqual(len(rows), 2)                                                 # HistorySave(UFCask:535)
        p = d["view"]["pack"]
        self.assertEqual((p["picture"], p["caption"]), ("1P1216", "RX021T0"))
        self.assertEqual([r["hiki"] for r in p["rows"]], ["61029052", "61029053"])
        r = self.c.get("/api/cask/pack_image", headers=self.h)
        self.assertEqual((r.status_code, r.mimetype), (200, "image/jpeg"))
        r.close()
        d = self.post("pack-select", i=2)
        self.assertEqual(d["view"]["pack"]["selected"], 2)
        d = self.post("pack-close", confirmed=False)
        self.assertIn(packing.NEED_CONFIRM, self.texts(d))
        self.assertIsNotNone(d["view"]["pack"])
        d = self.post("pack-close", confirmed=True)
        self.assertIsNone(d["view"]["pack"])
        self.assertEqual(self.c.get("/api/cask/pack_image", headers=self.h).status_code, 404)

    def test_packing_folder_missing(self) -> None:
        shutil.rmtree(paths.to_local(packing.folder_path()), ignore_errors=True)
        self.lot_readable()
        self.post("open")
        d = self.post("lot", lot="RX021T0")
        # fso.GetFolder の実行時エラー 76 → ErrHandler_ComboChange(UFCask:554)
        self.assertIn("DB読み込みエラー:\r\nパスが見つかりません。", self.texts(d))
        self.assertIsNone(d["view"]["pack"])

    def test_packing_skipped_after_history_error(self) -> None:
        self.lot_readable()
        self.post("open")
        with mock.patch("ufkensa.kensa.lot.history_save", side_effect=vbx.TypeMismatch("x")):
            d = self.post("lot", lot="RX021T0")
        errs = [t for t in self.texts(d) if t.startswith("DB読み込みエラー:")]
        self.assertEqual(len(errs), 1)                                                 # PackingNoCheck まで進まない
        self.assertNotIn("パスが見つかりません。", errs[0])

    def test_bad_requests(self) -> None:
        self.assertEqual(self.c.post("/api/cask/nothing", json={}, headers=self.h).status_code, 404)
        self.assertEqual(self.c.post("/api/cask/select", json={"row": 7}, headers=self.h).status_code, 400)



def _xlsx_sheets(files) -> list[str]:
    """保存した xlsx(1 つ)のシート名。"""
    import re
    import zipfile
    xs = [p for p in files if str(p).endswith(".xlsx") and Path(p).is_file()]
    assert xs, files
    with zipfile.ZipFile(xs[0]) as z:
        return re.findall(r'<sheet name="([^"]+)"', z.read("xl/workbook.xml").decode("utf-8"))


class SheetJsonTest(unittest.TestCase):
    def test_round_trip(self) -> None:
        from datetime import datetime
        ws = sh.template(sh.CASK_TEMPLATE)
        ws.set(1, 2, datetime(2026, 11, 2, 10, 0))
        ws.set_formula(2, 17, "=B3-M2")
        ws.fill(6, 11, 706, 28, sh.GRAY)
        ws.underline.add((7, 2))
        ws.locked.update({(r, 8) for r in range(7, 707)})
        ws.col_width[8] = 0.05
        back = cask_session.sheet_from_json(sh.CASK, cask_session.sheet_to_json(ws))
        self.assertEqual(back.values, ws.values)
        self.assertEqual(back.formulas, ws.formulas)
        self.assertEqual(back.interior, ws.interior)
        self.assertEqual((back.underline, back.locked, back.col_width), (ws.underline, ws.locked, ws.col_width))


if __name__ == "__main__":
    unittest.main()
