"""全ヵ所・H494 画面 /zenkasho と /api/zenkasho/*(app/routes/zenkasho.py → zenkasho_session.py)

基準データ(tests/kensa_fixture.py)のロット RX021T0(厚 25 × 巾 1385 × 丈 2770.5、LVC)で、
開く(作業名の InputBox)→ ロット → 伝送文字列で測定 → 作業の切替(5N・三菱・H494)→ 検入高さ → 印刷・保存、
作業中保存からの読み戻しを確かめる。業務の結果そのものは test_kensa_modes_* で確かめている。
"""
from __future__ import annotations

import shutil
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path

from tests import kensa_fixture
from tests._web import make_client
from ufkensa import user_settings
from unittest import mock

from ufkensa.kensa import packing, paths, quality, quirks, store, terminal, wiring, zenkasho_session
from ufkensa.kensa.grid import KensaState
from ufkensa.kensa.modes import sheet as sh
from tests._isolation import vba_reproduce

LOT = "RX021T0"


class Clock:
    def __init__(self) -> None:
        self.t = datetime(2026, 11, 2, 10, 5, 30)

    def __call__(self) -> datetime:
        self.t += timedelta(seconds=7)
        return self.t


class ZenBase(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="ufkensa_zen_"))
        kensa_fixture.use_offline(kensa_fixture.copy_to(self.tmp / "db"))
        paths.set_path("local_history", str(self.tmp / "検査履歴") + "/")
        terminal.set_facility("LVC")
        self.store = store.LocalStore(self.tmp / "data" / "kensa_local.sqlite3")
        self.clock = Clock()
        # UF1 の検査員は登録済み(Register_Check = 偽)にしておく。登録の要否そのものは RegisterCheckTest で見る
        p = mock.patch.object(zenkasho_session.ZenkashoSession, "_registered", return_value=False)
        self.registered = p.start()
        self._reg_patch = p
        self.addCleanup(mock.patch.stopall)
        self.ks = KensaState(line_name="LVC")
        self.s = self.new_session()

    def new_session(self) -> zenkasho_session.ZenkashoSession:
        return zenkasho_session.ZenkashoSession(self.store, now=self.clock, kensa_state=lambda: self.ks,
                                                write_logs=False)

    def tearDown(self) -> None:
        self.store.close()
        user_settings.save(terminal.KEY_FACILITY, "")
        paths.reset_path("local_history")
        shutil.rmtree(self.tmp, ignore_errors=True)

    def open_ippan(self) -> dict:
        d = self.s.start().to_dict()
        self.assertEqual(d["question"]["key"], "mit_choice")           # V1 が空 → InputBox(UFMit:2210)
        self.assertEqual(d["question"]["buttons"], "input")
        d = self.s.start("一般").to_dict()
        self.assertIsNone(d["question"])
        d = self.s.set_form("LotBox", LOT.lower()).to_dict()          # KeyPress で大文字
        self.assertEqual(d["view"]["form"]["lot"], LOT)
        return d


@vba_reproduce
class SessionTest(ZenBase):
    def test_measure_flow(self) -> None:
        d = self.open_ippan()
        v = d["view"]
        self.assertEqual((v["mode"], v["page"]), ("一般", "P1"))
        self.assertEqual([m["key"] for m in v["modes"] if m["active"]], ["一般"])
        self.assertEqual(v["active"]["label"], "No.1 厚1")
        self.assertEqual(len(v["rows"]), 70)
        self.assertEqual(v["groups"][0], {"label": "厚", "span": 8, "tol": "24.550〜25.450"})
        d = self.s.receive("DT10002+0000025.012M").to_dict()
        row = d["view"]["rows"][0]["cells"]
        self.assertEqual(row[0:2], [["25.012", ""], ["○", "ok"]])
        self.assertEqual(d["view"]["active"]["label"], "No.1 厚2")      # 判定記号の入れ子の Change で次へ
        self.assertEqual(d["view"]["flash"], "OK")
        self.assertIn("zenkasho_ok", [m["sound"] for m in d["messages"]])
        d = self.s.receive("DT10002+0000026.000M").to_dict()            # 公差外
        self.assertEqual(d["view"]["rows"][0]["cells"][3], ["×", "ng"])
        self.assertEqual(d["view"]["flash"], "NG")
        self.s.select(7, 23)                                             # ラテラル: Change は動かさない → Enter で下へ
        d = self.s.receive("0.3").to_dict()
        self.assertEqual(d["view"]["active"]["label"], "No.2 ﾗﾃﾗﾙﾎﾞｰ")
        self.s.select(7, 30)
        d = self.s.receive("良").to_dict()                               # 文字で Change を抜ける → EnableEvents=False のまま
        self.assertFalse(d["view"]["events_enabled"])
        d = self.s.resume().to_dict()                                    # 処理再開
        self.assertTrue(d["view"]["events_enabled"])
        d = self.s.erase().to_dict()
        self.assertEqual(d["view"]["rows"][1]["cells"][27][0], "")

    def test_modes(self) -> None:
        self.open_ippan()
        d = self.s.choose("5N").to_dict()
        v = d["view"]
        self.assertEqual(v["mode"], "5N")
        hiz = next(g for g in v["groups"] if g["label"] == "歪")
        self.assertEqual(hiz["tol"], "幅 2.0 / 丈 3.0(5N 固定)")
        self.assertEqual([p["which"] for p in v["prints"]], ["fini"])
        self.assertEqual(v["notice"], "")
        d = self.s.choose("三菱").to_dict()
        v = d["view"]
        self.assertEqual(v["notice"], "三菱受託は確認待ちの項目があります(docs/04 W2)")
        self.assertEqual([p["which"] for p in v["prints"]], ["mitsubishi"])
        self.assertIn("元板NO", [c["label"] for c in v["columns"]])
        d = self.s.choose("H494").to_dict()
        v = d["view"]
        self.assertEqual((v["mode"], v["page"]), ("H494", "P_H494"))
        self.assertEqual(len(v["rows"]), 10)
        self.assertIn("幅3", [c["label"] for c in v["columns"]])
        self.assertFalse(self.s.book.sheets[sh.ZEN].visible)            # QueryClose で隠す
        d = self.s.choose("一般").to_dict()
        self.assertEqual(d["view"]["page"], "P1")
        self.assertEqual(d["view"]["mode"], "一般")

    def test_height_and_print_save(self) -> None:
        self.open_ippan()
        self.s.resume()
        self.s.receive("DT10002+0000025.012M")
        d = self.s.set_form("YamaLB", "1").to_dict()
        self.assertEqual([k["n"] for k in d["view"]["form"]["kenh"] if k["visible"]], [1])
        self.s.set_form("KenHBox1", "250")
        d = self.s.save_height().to_dict()
        self.assertEqual(d["view"]["form"]["takasa_formula"], "(250.0÷25)=10")
        self.assertIn("ゲージによる測定必要", [m["text"] for m in d["messages"]])
        self.assertEqual(self.s.book.ws(sh.ZEN).get(2, 30), "1山")
        self.s.set_form("Kname", "検査員01")
        d = self.s.print("normal").to_dict()
        self.assertEqual(d["question"]["key"], "zen_print")
        d = self.s.print("normal", {"zen_print": True}).to_dict()
        self.assertIsNone(d["stop"])
        rep = d["view"]["reports"][0]
        html = Path(rep["path"]).read_text(encoding="utf-8")
        self.assertIn("厚公差：24.550～25.450", html)                  # 右ヘッダ(Mit:901)
        self.assertIn("25.012", html)
        saved = d["view"]["saved"]
        self.assertTrue(any(p.endswith(f"20261102_{LOT}.html") for p in saved))
        self.assertTrue(any(p.endswith(f"20261102_{LOT}_全ヵ所検査保存.csv") for p in saved))
        hz = self.store.sheet(store.HOZON)
        self.assertEqual(hz.get(hz.last_row(2), 2), LOT)                # 保存用へ履歴行
        self.assertEqual(hz.get(hz.last_row(2), 77), "別ファイル参照")
        # 2 回目は連番(.html を見て .002)
        d = self.s.print("normal", {"zen_print": True}).to_dict()
        self.assertTrue(any(p.endswith(f"20261102_{LOT}.002.html") for p in d["view"]["saved"]))

    def test_restore(self) -> None:
        self.open_ippan()
        self.s.receive("DT10002+0000025.012M")
        s2 = self.new_session()                                          # アプリを開き直した
        d = s2.start().to_dict()
        v = d["view"]
        self.assertEqual(v["mode"], "一般")                              # V1 から戻る(InputBox は出ない)
        self.assertEqual(v["rows"][0]["cells"][0], ["25.012", ""])
        self.assertEqual(v["form"]["lot"], LOT)
        self.assertTrue(self.store.exists("全ヵ所作業中保存"))

    def test_new_sheet_question(self) -> None:
        self.open_ippan()
        self.s.receive("DT10002+0000025.012M")
        d = self.s.new_sheet().to_dict()
        self.assertEqual(d["question"]["key"], "zen_new")
        d = self.s.new_sheet({"zen_new": True}).to_dict()
        self.assertEqual(d["view"]["rows"][0]["cells"][0][0], "")
        self.assertEqual(d["view"]["form"]["lot"], "")
        self.assertFalse(d["view"]["events_enabled"])                    # LotBox="" で EnableEvents=False(VBA どおり)


class ScreenTest(ZenBase):
    def setUp(self) -> None:
        super().setUp()
        zenkasho_session.reset_for_test(self.s)
        self.addCleanup(zenkasho_session.reset_for_test, None)
        self.c = make_client()
        app = self.c.application
        from app.routes import zenkasho
        if "zenkasho" not in app.blueprints:
            app.register_blueprint(zenkasho.bp)
        if "pending_zenkasho" in app.view_functions:                    # 「準備中」の同じ URL より先に出す
            app.view_functions["pending_zenkasho"] = zenkasho.page
        self.h = {"X-Tool-Token": app.config["TOKEN"]}

    def post(self, op: str, **body):
        r = self.c.post(f"/api/zenkasho/{op}", json=body, headers=self.h)
        self.assertEqual(r.status_code, 200, r.get_data(as_text=True))
        return r.get_json()

    def test_page_and_round_trip(self) -> None:
        r = self.c.get("/zenkasho", headers=self.h)
        self.assertEqual(r.status_code, 200)
        self.assertIn('id="sheet"', r.get_data(as_text=True))
        st = self.c.get("/api/zenkasho/state", headers=self.h).get_json()
        self.assertIn("rows", st["view"])
        d = self.post("start")
        self.assertEqual(d["question"]["key"], "mit_choice")
        d = self.post("start", choice_input="５N")                        # 全角は 5N に(UFMit:2212)
        self.assertEqual(d["view"]["mode"], "5N")
        d = self.post("form", name="LotBox", value=LOT)
        d = self.post("resume")
        d = self.post("receive", text="DT10002+0000025.012M")
        self.assertEqual(d["view"]["rows"][0]["cells"][1], ["○", "ok"])
        d = self.post("select", r=9, c=11)
        self.assertEqual(d["view"]["active"]["label"], "No.3 幅1")
        d = self.post("select", r=999, c=1)
        self.assertIn("セルの位置", d["stop"]["text"])
        d = self.post("choose", mode="H494")
        self.assertEqual(d["view"]["page"], "P_H494")
        d = self.post("print", which="normal")                           # H494 では一般の印刷は無い
        self.assertIsNotNone(d["stop"])

    def test_unknown_op_and_report(self) -> None:
        r = self.c.post("/api/zenkasho/nothing", json={}, headers=self.h)
        self.assertEqual(r.status_code, 404)
        r = self.c.get("/api/zenkasho/report/x.txt", headers=self.h)
        self.assertEqual(r.status_code, 400)
        r = self.c.get("/api/zenkasho/report/none.html", headers=self.h)
        self.assertEqual(r.status_code, 404)

    def test_pack_image_and_ops(self) -> None:
        r = self.c.get("/api/zenkasho/pack_image", headers=self.h)
        self.assertEqual(r.status_code, 404)
        f = paths.to_local(packing.folder_path())
        f.mkdir(parents=True, exist_ok=True)
        (f / "1P1216.JPG").write_bytes(b"\xff\xd8\xff\xe0")
        self.post("start", choice_input="一般")
        d = self.post("form", name="LotBox", value=LOT)
        self.assertEqual(d["view"]["pack"]["picture"], "1P1216")
        r = self.c.get("/api/zenkasho/pack_image?n=1P1216", headers=self.h)
        self.assertEqual((r.status_code, r.mimetype), (200, "image/jpeg"))
        r.close()
        d = self.post("pack-select", i=1)
        self.assertEqual(d["view"]["pack"]["selected"], 1)
        d = self.post("pack-close", confirmed=True)
        self.assertIsNone(d["view"]["pack"])

@vba_reproduce
class ConnectTest(ZenBase):
    """部品の Effects.calls をつないだもの(HistorySave・GetLotInfo/DbReturnItem・包装仕様書・不良要因票・日報・UF対角)。

    通常検査(UF1)の状態は KensaState(ライン LVC・ロット AAA0001)を読むだけで渡す。
    """

    def new_session(self) -> zenkasho_session.ZenkashoSession:
        self.ks.lot = "AAA0001"
        return super().new_session()

    def pack_folder(self, *names: str) -> Path:
        f = paths.to_local(packing.folder_path())
        f.mkdir(parents=True, exist_ok=True)
        for n in names:
            (f / n).write_bytes(b"\xff\xd8\xff\xe0")
        return f

    def history_rows(self) -> list:
        rows = wiring.sheet_rows(self.store, store.LOT_HISTORY) or []
        return [r for r in rows[1:] if r[2] == LOT]

    def test_history_save_and_packing_folder_missing(self) -> None:
        d = self.open_ippan()
        rows = self.history_rows()                                       # HistorySave(UFMit:804): 引当ごとに 1 行
        self.assertEqual(sorted(r[21] for r in rows), ["61029052", "61029053"])
        self.assertEqual({r[25] for r in rows}, {"1P1216"})
        # 包装仕様書のフォルダが無い → fso.GetFolder の実行時エラー 76 → ErrHandler_MitCombo(UFMit:874)
        self.assertIn("DB読み込みエラー:\r\nパスが見つかりません。", [m["text"] for m in d["messages"]])
        self.assertIsNone(d["view"]["pack"])
        self.assertTrue(d["view"]["events_enabled"])                     # LotBox_Change が最後に True に戻す
        self.s.set_form("LotBox", LOT)                                   # 同じロットを読み直しても行は増えない
        self.assertEqual(len(self.history_rows()), 2)

    def test_packing_dialog(self) -> None:
        self.pack_folder("1P1216.jpg", "1P0113.JPG")
        d = self.open_ippan()
        p = d["view"]["pack"]
        self.assertEqual((p["tb_load"], p["picture"], p["caption"], p["selected"]), ("1P1216", "1P1216", LOT, 0))
        self.assertEqual([r["hiki"] for r in p["rows"]], ["61029052", "61029053"])    # 引当NO で並べる(merge_sort2)
        self.assertEqual(p["spec"][0:2], ["1P1216", "梱包形態:パレット"])               # PackSelect(PkgSpecData)
        d = self.s.pack_select(2).to_dict()
        self.assertEqual(d["view"]["pack"]["selected"], 2)
        d = self.s.pack_close(False).to_dict()                                         # 確認の印が無い
        self.assertIn(packing.NEED_CONFIRM, [m["text"] for m in d["messages"]])
        self.assertIsNotNone(d["view"]["pack"])
        self.assertEqual(self.s.pack_image().name, "1P1216.jpg")
        d = self.s.pack_close(True).to_dict()
        self.assertIsNone(d["view"]["pack"])
        self.assertIsNone(self.s.pack_image())

    def test_packing_not_while_opening(self) -> None:
        self.pack_folder("1P1216.JPG")
        self.open_ippan()
        self.s.pack_close(True)
        s2 = self.new_session()                                          # 開き直す: PackNo_on_off = "off" の間の読込
        d = s2.start().to_dict()
        self.assertIsNone(d["view"]["pack"])
        # 開く途中の FormReflectDist は UF1 の欄にしか書かない → 書かずに理由を残す
        self.assertTrue(any(x.startswith("FormReflectDist(開く途中") for x in d["view"]["pending"]))

    def test_get_lot_info_empty_exits(self) -> None:
        self.pack_folder("1P1216.JPG")
        with mock.patch("ufkensa.kensa.lot.get_lot_info", return_value=None), \
                mock.patch.object(quality, "dist_array") as dist:
            d = self.open_ippan()
        dist.assert_not_called()                                         # If IsEmpty(TempArry) Then Exit Sub
        self.assertIsNone(d["view"]["pack"])

    def test_db_return_item(self) -> None:
        calls = []

        def fake(nou, tori, youtc, zai, tyou, atu, haba, take, **kw):
            calls.append((nou, tori, youtc, zai, tyou, atu, haba, take, kw["mitsubishi_visible"], kw["work_flag"]))
            r = quality.ReturnItemResult(items=",ラテラルボー,フリー反り")
            r.uf1["HURIS"] = ""
            if kw["mitsubishi_visible"]:
                r.sheet = "全ヵ所検査"
                r.cells[(1, 15)] = nou + zai
                r.cells[(6, 29)] = "ﾌﾘｰｿﾘ"
            return r
        with mock.patch.object(quality, "db_return_item", side_effect=fake):
            self.open_ippan()
        # 納入先が同じ 2 引当 → 1 回だけ(UFMit:838-839)。厚・幅・丈は Single
        self.assertEqual(len(calls), 1)
        self.assertEqual(calls[0][:8], ("ｺ-ﾐｷﾝｿﾞｸ(ｶ ﾌｸｼﾏｼﾃﾝ", "ｺ-ﾐｷﾝｿﾞｸ(ｶ", "H197", "R61S", "T651", 25.0, 1385.0, 2770.5))
        self.assertEqual(calls[0][8:], (True, "P1"))
        ws = self.s.book.ws(sh.ZEN)
        self.assertEqual(ws.get(1, 15), "ｺ-ﾐｷﾝｿﾞｸ(ｶ ﾌｸｼﾏｼﾃﾝR61S")
        self.assertEqual(ws.get(6, 29), "ﾌﾘｰｿﾘ")
        self.assertEqual(self.s.pending, [])

    def print_normal(self) -> dict:
        self.open_ippan()
        self.s.resume()
        self.s.receive("DT10002+0000025.012M")
        self.s.set_form("YamaLB", "1")
        self.s.set_form("KenHBox1", "250")
        self.s.save_height()
        self.s.set_form("Kname", "検査員01")
        return self.s.print("normal", {"zen_print": True}).to_dict()

    def test_defect_report(self) -> None:
        d = self.print_normal()
        self.assertIsNone(d["stop"])
        labels = [r["label"] for r in d["view"]["reports"]]
        self.assertEqual(labels, ["全ヵ所検査", "不良要因印刷_LVC"])
        self.assertNotIn("印刷:不良要因印刷_LVC", d["view"]["pending"])
        name = "不良要因印刷_LVC"
        self.assertEqual(self.store.sheet(name).get(1, 2), LOT)          # 全ヵ所用転送の B1(Mit:735)
        page = __import__("json").loads(self.store.load_report(name + "#page"))["page"]
        # 不良要因印刷範囲(Kname) は LotN を渡さない形 → UF1 の Lname・Lot(Print:1279-1285)
        self.assertEqual((page["header_left"], page["header_right"]), ("LVC", "AAA0001 検査員01"))
        self.assertTrue(page["footer_right"].startswith(" 2026/11/02  10:05～"))     # 開始 = 全ヵ所検査!B1
        html = Path(d["view"]["reports"][1]["path"]).read_text(encoding="utf-8")
        self.assertIn(LOT, html)
        # 新規発行: Register_Check(真 = 直が変わって登録が要る)→ Daily_Create、不良要因シートは消す(UFMit:1273)
        self.registered.return_value = True
        d = self.s.new_sheet({"zen_new": True}).to_dict()
        self.assertIsNone(d["stop"])
        self.assertFalse(self.store.exists(name))
        self.assertTrue(self.store.exists("Daily_LVC"))
        self.assertEqual(self.store.sheet("Daily_LVC").get(2, 1), "ﾛｯﾄNo")

    def test_defect_range_other_lname(self) -> None:
        """丸徳の号機: シート名は UFMitsubishi.LineName、不良要因印刷範囲 は UF1 の Lname(号機つき)を見る。"""
        self.ks.line_no = "A22"                                          # Lname = LVCA22 → シートが無い → Exit Sub
        d = self.print_normal()
        self.assertEqual([r["label"] for r in d["view"]["reports"]], ["全ヵ所検査", "不良要因印刷_LVC"])
        self.assertIsNone(self.store.load_report("不良要因印刷_LVC#page"))

    def test_defect_range_b85_fix(self) -> None:
        """B85 を直す: 全ヵ所のライン名・ロットで動く(号機つきでもページ設定が付き、右ヘッダは全ヵ所のロット)。"""
        quirks.set_for_test({"B85": "fix"})
        self.addCleanup(quirks.set_for_test, None)
        self.ks.line_no = "A22"
        self.print_normal()
        page = __import__("json").loads(self.store.load_report("不良要因印刷_LVC#page"))["page"]
        self.assertEqual((page["header_left"], page["header_right"]), ("LVC", LOT + " 検査員01"))
        self.assertTrue(page["footer_right"].startswith(" 2026/11/02  10:05～"))     # フッターは今までどおり

    def test_uf_taikaku_h494(self) -> None:
        self.open_ippan()
        self.s.choose("H494")
        self.s.set_form("H494_Lot", LOT)
        self.s.resume()
        self.s.select(7, 21)
        d = self.s.receive("1960.5").to_dict()                           # 対角1 を入れた → UF対角(Cask:1899-1918)
        self.assertTrue(d["view"]["diag"])
        self.assertEqual(d["view"]["flash"], "")
        self.assertNotIn("UF対角", d["view"]["pending"])



class RegisterCheckTest(ZenBase):
    """UF1.ZEB_Click / H494B_Click: Register_Check が真なら「検査員登録をして下さい」で開かない(UF1:4870-4875 / 604-609)。"""

    def test_start_needs_worker(self) -> None:
        self.registered.return_value = True
        d = self.s.start("一般").to_dict()
        self.assertIn(zenkasho_session.NEED_WORKER, [m["text"] for m in d["messages"]])
        self.assertFalse(d["view"]["open"])
        d = self.s.choose("H494").to_dict()                              # 閉じているときの H494 も同じ
        self.assertIn(zenkasho_session.NEED_WORKER, [m["text"] for m in d["messages"]])
        self.assertFalse(d["view"]["open"])
        self.registered.return_value = False
        d = self.s.start("一般").to_dict()
        self.assertTrue(d["view"]["open"])
        self.registered.return_value = True                              # 開いた後の切替では見ない(フォーム内のページ)
        d = self.s.choose("H494").to_dict()
        self.assertEqual((d["view"]["open"], d["view"]["page"]), (True, "P_H494"))

    def test_real_register_check(self) -> None:
        self._reg_patch.stop()                                           # 本物: 検査員_LVC が無い → 真(Worker:456-)
        s = self.new_session()
        d = s.start("一般").to_dict()
        self.assertIn(zenkasho_session.NEED_WORKER, [m["text"] for m in d["messages"]])
        self.assertFalse(d["view"]["open"])

    def test_no_uf1_state(self) -> None:
        s = zenkasho_session.ZenkashoSession(self.store, now=self.clock, kensa_state=lambda: None, write_logs=False)
        d = s.start("一般").to_dict()
        self.assertIn(zenkasho_session.NEED_WORKER, [m["text"] for m in d["messages"]])
        self.assertFalse(d["view"]["open"])


class LateralQuietTest(ZenBase):
    def test_quiet_follows_is_over_12_hours(self) -> None:
        """DbReturnItem → DBdisplayラテラル は IsOver12Hours が真なら MsgBox を出さない(VBA どおり)。"""
        from ufkensa.kensa import wiring as w
        for over in (False, True):
            got = []

            def fake(*a, **kw):
                got.append(kw["lateral_quiet"])
                return quality.ReturnItemResult()
            with mock.patch.object(quality, "db_return_item", side_effect=fake), \
                    mock.patch.object(w.RealHooks, "_over12", return_value=over):
                self.s = self.new_session()
                self.s.start("一般")
                self.s.set_form("LotBox", LOT)
            self.assertEqual(set(got), {over})
            self.store.delete_sheet(zenkasho_session.META_SHEET)


if __name__ == "__main__":
    unittest.main()


class ShownModesTest(unittest.TestCase):
    def test_d21_unused_modes_hidden(self) -> None:
        """決定 D21: 三菱受託・5N は使っていないので切り替えに出さない(その作業中なら出す)。"""
        from ufkensa.kensa import zenkasho_session as zs
        self.assertEqual(zs.SHOWN_MODES, ("一般", "H494"))
        keys = [k for k, _ in zs.MODES if k in zs.SHOWN_MODES or k == "三菱"]
        self.assertEqual(keys, ["一般", "三菱", "H494"])
