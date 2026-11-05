"""権限(判断事項 D12・D30: 梱包資材マスタの「アクセス権限」。all-tools 日報ツールと同じ読み方)

表は 管理番号・ログインID・PC名・権限・有効・備考(ほかのツールと共有)。
- この PC の行 = ログインID と PC名 の両方が当たる行(空欄は問わない・両方空は誰の行でもない・大文字小文字/全角半角は
  問わない・DOMAIN\\user は user でも当たる・有効 0/いいえ/無効 は読まない)
- 権限: Administrator(= ufkensa:admin)・ufkensa:calibration_skip・正規のライン名(→ この端末の設備名)。ほかは読まない
- 違うラインが 2 つ → 決めない。表のラインは前に当てたものから変わったときだけ当てる
- マスタから決まらない(未設定・読めない・管理者の行が無い)あいだは、管理者パスワードだけで設定できる(設定の途中)
- 使用パス「既定に戻す」(項目ごと・すべて)。鍵は使用パスの変更と同じ。前後をログに残す
"""
from __future__ import annotations

import os
import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest import mock

from tests._web import make_client
from ufkensa import access_rights, admin_password, auth, config, trace, user_settings
from ufkensa.kensa import paths, terminal

A = access_rights


def master_db(path: Path, rows) -> Path:
    """梱包資材マスタと同じ形の「アクセス権限」表。rows = (ログインID, PC名, 権限, 有効[, 備考])。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    with closing(sqlite3.connect(path)) as c:
        c.execute('CREATE TABLE IF NOT EXISTS "アクセス権限" ("管理番号" INTEGER PRIMARY KEY, "ログインID" TEXT,'
                  ' "PC名" TEXT, "権限" TEXT, "有効" INTEGER, "備考" TEXT)')
        c.execute('DELETE FROM "アクセス権限"')
        c.executemany('INSERT INTO "アクセス権限" (ログインID, PC名, 権限, 有効, 備考) VALUES (?, ?, ?, ?, ?)',
                      [tuple(r) + ("",) * (5 - len(r)) for r in rows])
        c.commit()
    return path


def rows(*items):
    return [{"管理番号": i, "ログインID": r[0], "PC名": r[1], "権限": r[2], "有効": r[3] if len(r) > 3 else 1}
            for i, r in enumerate(items, 1)]


class _Env(unittest.TestCase):
    """一時フォルダ・この PC の名前・設定を毎回きれいにする。"""
    login, pc = "genba", "NLM-L1-01"

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.tmp = Path(tmp.name)
        env = mock.patch.dict(os.environ, {"UFKENSA_LOCAL_DIR": str(self.tmp / "local"),
                                           "USERNAME": self.login, "COMPUTERNAME": self.pc})
        env.start()
        self.addCleanup(env.stop)
        for k in ("kensa_facility", "kensa_line_no", A.KEY_LINE_APPLIED):
            user_settings.save(k, "")
            self.addCleanup(user_settings.save, k, "")
        paths.reset_item("access_master")
        self.addCleanup(paths.reset_item, "access_master")
        auth.end()
        auth.lock_master()
        self.addCleanup(auth.end)
        self.addCleanup(auth.lock_master)

    def use_master(self, rows_, where: str = "file") -> Path:
        path = master_db(self.tmp / "参照用" / A.DEFAULT_FILE, rows_)
        paths.set_path("access_master", str(path if where == "file" else path.parent))
        auth.end()
        return path


# ----------------------------------------------------------------------
# 判断(純粋)
# ----------------------------------------------------------------------
class MatchTest(unittest.TestCase):
    def test_both_columns(self):
        r = lambda login, pc: {"ログインID": login, "PC名": pc}
        self.assertTrue(A.applies(r("genba", "PC-1"), "genba", "PC-1"))
        self.assertFalse(A.applies(r("genba", "PC-2"), "genba", "PC-1"))     # 共通アカウントでも PC が違えば当たらない
        self.assertTrue(A.applies(r("", "PC-1"), "anyone", "PC-1"))          # 空欄は問わない
        self.assertTrue(A.applies(r("genba", ""), "genba", "どこでも"))
        self.assertFalse(A.applies(r("", ""), "genba", "PC-1"))               # 両方空は誰の行でもない
        self.assertTrue(A.applies(r("ＧＥＮＢＡ", "pc-1"), "genba", "PC-1"))  # 全角・大文字小文字
        self.assertTrue(A.applies(r("NIKKEI\\genba", "PC-1"), "genba", "PC-1"))   # DOMAIN\user
        self.assertTrue(A.applies(r("genba", "PC-1"), "NIKKEI\\genba", "PC-1"))

    def test_enabled(self):
        for off in (0, "0", "いいえ", "無効", "false", "×"):
            self.assertFalse(A.enabled(off), off)
        for on in (1, "1", None, "", "はい"):
            self.assertTrue(A.enabled(on), on)

    def test_admin_and_codes(self):
        d = lambda rs, who="taro", pc="PC-1": A.decide(rows(*rs), who, pc)
        self.assertTrue(d([("taro", "PC-1", "Administrator")]).admin)
        self.assertTrue(d([("taro", "PC-1", "administrator")]).admin)
        self.assertTrue(d([("taro", "", "ufkensa:admin")]).admin)            # 前の書き方(互換)
        self.assertTrue(d([("taro", "PC-1", "mode:field、Administrator")]).admin)   # 並べて書ける
        self.assertFalse(d([("taro", "PC-1", "Administrator", 0)]).admin)    # 有効でない
        self.assertFalse(d([("taro", "PC-2", "Administrator")]).admin)
        self.assertFalse(d([("", "", "Administrator")]).admin)
        self.assertFalse(d([("taro", "PC-1", "編集可")]).admin)              # db-viewer の書き方は読まない
        r = d([("taro", "PC-1", "ufkensa:calibration_skip / mode:material")])
        self.assertEqual((r.admin, r.calibration_skip, r.ignored), (False, True, ("mode:material",)))


class LineTest(unittest.TestCase):
    def test_all_tools_example_common_account(self):
        """all-tools v4.12.1: 共通アカウントが 12 台ぶんの行に入っていても、この PC の行だけでラインが決まる。"""
        table = rows(("genba", "NLM-L1-01", "L-1"), ("genba", "NLM-HVC-01", "HVC"),
                     ("genba", "NLM-LVC-01", "LVC"), ("genba", "NLM-L1-01", "mode:field"),
                     ("genba", "NLM-BAL-01", "ﾊﾞﾗﾝｻｰ"), ("yamada", "", "Administrator"))
        r = A.decide(table, "genba", "NLM-L1-01")
        self.assertEqual((r.facility, r.line_no, r.admin, r.problem), ("L-1", "", False, ""))
        self.assertEqual(r.ignored, ("mode:field",))
        self.assertEqual(len(r.matched), 2)
        self.assertEqual(A.decide(table, "genba", "NLM-BAL-01").facility, "バランサー")   # 幅だけ揃える(D13)
        self.assertEqual(A.decide(table, "yamada", "NLM-L1-01").admin, True)

    def test_official_names_only(self):
        for token, facility in (("L-1", "L-1"), ("LVC", "LVC"), ("ＨＶＣ", "HVC"), ("ﾄｯﾄ", "トット"),
                                ("バランサー", "バランサー"), ("中板A3", "丸徳")):
            self.assertEqual(A.read_line(token).facility, facility, token)
        for token in ("L1", "l-1", "L 1", "丸徳", "TOT", "機側", "AIM", "NS1", "コイル", "Test", "作業長"):
            self.assertEqual(A.read_line(token).facility, "", token)
        self.assertIn("正規は「L-1」", A.hint("L1"))
        self.assertIn("正規は「L-1」", A.hint("l-1"))
        self.assertIn("正規は「中板」", A.hint("丸徳"))
        self.assertIn("号機を続けて", A.hint("中板 3"))
        for silent in ("mode:field", "作業長", "コイル", "機側", "AIM"):
            self.assertEqual(A.hint(silent), "", silent)

    def test_maru_number(self):
        self.assertEqual(A.read_line("中板A3"), A.LineReading("丸徳", "A3"))
        self.assertEqual(A.read_line("中板Ａ１１").number, "A11")
        r = A.read_line("中板3")                                             # 日報ツールのラインNO の書き方
        self.assertEqual((r.facility, r.number), ("丸徳", ""))
        self.assertIn("号機(A1〜A11)は決めません", r.note)
        self.assertEqual((A.read_line("中板").facility, A.read_line("中板").number), ("丸徳", ""))
        self.assertIn("A1〜A11 ではありません", A.read_line("中板A12").problem)
        r = A.decide(rows(("genba", "PC-1", "中板A12")), "genba", "PC-1")
        self.assertEqual(r.facility, "")
        self.assertIn("中板A12", r.ignored[0])

    def test_conflict_and_merge(self):
        r = A.decide(rows(("genba", "PC-1", "L-1"), ("", "PC-1", "HVC")), "genba", "PC-1")
        self.assertEqual(r.facility, "")
        self.assertIn("ラインが2つ", r.problem)
        self.assertEqual(r.lines, ("L-1", "HVC"))
        r = A.decide(rows(("genba", "PC-1", "中板A3, 中板A5")), "genba", "PC-1")   # 号機違いも決めない
        self.assertEqual(r.facility, "")
        self.assertIn("丸徳A3・丸徳A5", r.problem)
        r = A.decide(rows(("genba", "PC-1", "中板"), ("genba", "PC-1", "中板A4")), "genba", "PC-1")
        self.assertEqual((r.facility, r.line_no, r.problem), ("丸徳", "A4", ""))
        r = A.decide(rows(("genba", "PC-1", "L-1"), ("genba", "PC-1", "L-1")), "genba", "PC-1")   # 同じラインは1つ
        self.assertEqual((r.facility, r.problem), ("L-1", ""))

    def test_inspect(self):
        table = rows(("genba", "PC-1", "L1"), ("", "", "HVC"), ("genba", "PC-2", "LVC"), ("genba", "PC-2", "HVC"),
                     ("genba", "PC-3", "mode:field・作業長"), ("genba", "PC-4", "丸徳"), ("x", "PC-5", "L-1", 0))
        found = [f.describe() for f in A.inspect(table)]
        self.assertTrue(any("管理番号 1" in f and "正規は「L-1」" in f for f in found), found)
        self.assertTrue(any("管理番号 2" in f and "どのPCの行にもなりません" in f for f in found), found)
        self.assertTrue(any("ラインが2つ" in f and "HVC・LVC" in f for f in found), found)
        self.assertTrue(any("管理番号 6" in f and "正規は「中板」" in f for f in found), found)
        self.assertFalse(any("管理番号 5" in f for f in found), found)        # ラインでない値は言わない
        self.assertEqual(len(found), 4)

    def test_should_apply(self):
        self.assertTrue(A.should_apply("L-1", ""))
        self.assertFalse(A.should_apply("L-1", "L-1"))
        self.assertTrue(A.should_apply("丸徳:A3", "L-1"))
        self.assertFalse(A.should_apply("", "L-1"))


# ----------------------------------------------------------------------
# 表を読む・前回の結果・設定の途中
# ----------------------------------------------------------------------
class EvaluateTest(_Env):
    def test_unset_is_setup(self):
        d = A.evaluate(self.login, self.pc)
        self.assertEqual((d.source, d.edit, d.setup), ("unset", True, True))
        self.assertEqual(d.setup_reason, A.MSG_SETUP)
        self.assertIsNone(A.master_path())

    def test_folder_or_file(self):
        path = self.use_master([(self.login, self.pc, "Administrator")], where="folder")
        self.assertEqual(A.master_path(), path)                              # フォルダなら中の 梱包資材マスタ.sqlite3
        d = A.evaluate(self.login, self.pc)
        self.assertEqual((d.source, d.edit, d.setup, d.codes), ("master", True, False, ["Administrator"]))
        paths.set_path("access_master", str(path))
        self.assertEqual(A.master_path(), path)

    def test_not_admin_and_line(self):
        self.use_master([(self.login, self.pc, "トット"), ("yamada", "", "Administrator")])
        d = A.evaluate(self.login, self.pc)
        self.assertEqual((d.edit, d.setup, d.facility, d.line_key), (False, False, "トット", "トット"))
        self.assertFalse(d.has(A.CALIBRATION_SKIP))

    def test_no_admin_row_is_setup(self):
        self.use_master([(self.login, self.pc, "L-1"), ("", "", "Administrator"), ("x", "", "Administrator", 0)])
        d = A.evaluate(self.login, self.pc)
        self.assertEqual((d.source, d.edit, d.setup), ("master", False, True))
        self.assertEqual(d.setup_reason, A.MSG_SETUP_NO_ADMIN)

    def test_unreachable_uses_cache_then_setup(self):
        path = self.use_master([(self.login, self.pc, "Administrator, 中板A2")])
        self.assertTrue(A.evaluate(self.login, self.pc).edit)                # 読めた結果を覚える
        path.unlink()
        d = A.evaluate(self.login, self.pc)
        self.assertEqual((d.source, d.edit, d.setup, d.facility, d.line_no), ("cache", True, False, "丸徳", "A2"))
        self.assertIn("前回", d.reason)
        d = A.evaluate("ken", self.pc)                                       # 前回の結果が無い人
        self.assertEqual((d.source, d.edit, d.setup), ("unreachable", False, True))
        self.assertEqual(d.setup_reason, A.MSG_SETUP_UNREAD)

    def test_missing_table(self):
        p = self.tmp / "空.sqlite3"
        with closing(sqlite3.connect(p)) as c:
            c.execute("CREATE TABLE x (a)")
        paths.set_path("access_master", str(p))
        ov = A.overview()
        self.assertFalse(ov["found"])
        self.assertIn("アクセス権限", ov["problem"])

    def test_overview(self):
        self.use_master([(self.login, self.pc, "L1"), ("yamada", "", "Administrator")])
        ov = A.overview()
        self.assertEqual((ov["found"], ov["rows"], len(ov["admin_rows"])), (True, 2, 1))
        self.assertIn("正規は「L-1」", ov["findings"][0])


class ApplyLineTest(_Env):
    def test_applies_only_when_master_changes(self):
        path = self.use_master([(self.login, self.pc, "HVC")])
        auth.ensure()
        self.assertEqual(terminal.facility(), "HVC")
        self.assertIn("HVC", auth.applied_note())
        terminal.set_facility("LVC")                                         # 管理者パスワードで手で変えた
        auth.end()
        auth.ensure()
        self.assertEqual(terminal.facility(), "LVC")                         # 表が同じなら戻さない
        master_db(path, [(self.login, self.pc, "中板A7")])
        auth.end()
        auth.ensure()
        self.assertEqual((terminal.facility(), terminal.line_no(), terminal.line_label()), ("丸徳", "A7", "中板A7"))

    def test_maru_without_number_keeps_current(self):
        terminal.set_facility("丸徳", "A5")
        self.use_master([(self.login, self.pc, "中板3")])
        auth.ensure()
        self.assertEqual((terminal.facility(), terminal.line_no()), ("丸徳", "A5"))

    def test_conflict_or_cache_does_not_apply(self):
        path = self.use_master([(self.login, self.pc, "L-1"), (self.login, "", "HVC")])
        auth.ensure()
        self.assertEqual(terminal.facility(), "")
        master_db(path, [(self.login, self.pc, "L-1")])
        d = A.evaluate(self.login, self.pc)
        path.unlink()
        cached = A.evaluate(self.login, self.pc)
        self.assertEqual((cached.source, cached.facility), ("cache", "L-1"))
        self.assertEqual(A.apply_line(cached), "")                           # 前回の結果では当てない
        self.assertTrue(A.apply_line(d))


# ----------------------------------------------------------------------
# 鍵(設定の途中)
# ----------------------------------------------------------------------
class BootstrapTest(_Env):
    def test_unreachable_master_password_alone(self):
        """試験用 PC: 使用パスの表に届かない → 一般だが、管理者パスワードだけで開けられる(詰まない)。"""
        paths.set_path("access_master", str(self.tmp / "無いフォルダ"))
        auth.end()
        self.assertFalse(auth.can_edit())
        self.assertTrue(auth.setup_mode())
        self.assertIn(A.MSG_SETUP_UNREAD, auth.master_edit_reason())
        self.assertEqual(auth.unlock_master("違う")[0], False)
        self.assertTrue(auth.unlock_master(admin_password.DEFAULT_PASSWORD)[0])
        self.assertTrue(auth.master_edit_ok())
        self.assertTrue(auth.can_edit())
        self.assertEqual(auth.master_edit_reason(), "")
        auth.lock_master()
        self.assertFalse(auth.can_edit())

    def test_unset_master(self):
        self.assertTrue(auth.setup_mode())
        self.assertIn(A.MSG_SETUP, auth.master_edit_reason())
        self.assertTrue(auth.unlock_master(admin_password.DEFAULT_PASSWORD)[0])
        self.assertTrue(auth.master_edit_ok())

    def test_readable_master_needs_administrator(self):
        self.use_master([("yamada", "", "Administrator")])
        self.assertFalse(auth.setup_mode())
        ok, why = auth.unlock_master(admin_password.DEFAULT_PASSWORD)
        self.assertFalse(ok)
        self.assertIn("管理者", why)
        self.assertFalse(auth.master_edit_ok())
        self.use_master([(self.login, "", "Administrator")])
        self.assertTrue(auth.can_edit())
        self.assertFalse(auth.master_edit_ok())                              # パスワードも要る
        self.assertEqual(auth.master_edit_reason(), "編集するには、管理者パスワードを入れてください")
        self.assertTrue(auth.unlock_master(admin_password.DEFAULT_PASSWORD)[0])
        self.assertTrue(auth.master_edit_ok())


# ----------------------------------------------------------------------
# 設定画面の API
# ----------------------------------------------------------------------
class SettingsApiTest(_Env):
    def setUp(self):
        super().setUp()
        self.c = make_client()
        self.h = {"X-Tool-Token": self.c.application.config["TOKEN"]}
        user_settings.delete(paths.KEY_PATHS)
        self.addCleanup(user_settings.delete, paths.KEY_PATHS)

    def post(self, url, **body):
        return self.c.post(url, json=body, headers=self.h)

    def state(self):
        return self.c.get("/api/settings/state", headers=self.h).get_json()

    def unlock(self):
        return self.post("/api/admin/unlock", password=admin_password.DEFAULT_PASSWORD).get_json()

    def test_bootstrap_on_a_test_pc(self):
        """試験用 PC の流れ: 表に届かない → パスワードだけで「アクセス権限マスタ」を直す → 以後はマスタの決まり。"""
        paths.set_path("access_master", str(self.tmp / "まだ無い"))
        auth.end()
        st = self.state()
        self.assertTrue(st["master_edit"]["setup"])
        self.assertTrue(st["master_edit"]["possible"])
        self.assertEqual(st["rights"]["source"], "unreachable")
        r = self.post("/api/settings/path", key="access_master", path="x")
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "locked"))
        self.assertIn("管理者パスワードだけで設定できます", r.get_json()["error"]["message"])
        self.assertTrue(self.unlock()["ok"])
        master = master_db(self.tmp / "試験" / A.DEFAULT_FILE, [(self.login, self.pc, "Administrator, LVC"),
                                                                ("genba", "PC-X", "L1")])
        st = self.post("/api/settings/path", key="access_master", path=str(master.parent)).get_json()
        rt = st["rights"]
        self.assertEqual((rt["source"], rt["setup"], rt["admin"], rt["codes"]), ("master", False, True, ["Administrator"]))
        self.assertEqual(rt["table"]["path"], str(master))
        self.assertEqual(rt["line"]["facility"], "LVC")
        self.assertEqual(st["terminal"]["facility"], "LVC")                  # ラインも自動で
        self.assertTrue(any("L1" in f for f in rt["table"]["findings"]))
        self.assertEqual(len(rt["matched"]), 1)
        self.assertTrue(st["master_edit"]["ok"])                              # 鍵は開けたまま
        ck = self.post("/api/settings/path/check", key="access_master").get_json()
        self.assertTrue(ck["ok"])
        self.assertIn("2 行", ck["message"])

    def test_not_admin_is_403(self):
        self.use_master([("yamada", "", "Administrator")])
        self.unlock()
        for url, body in (("/api/settings/path", {"key": "wav", "path": "x"}),
                          ("/api/settings/path/reset", {"key": "wav"}),
                          ("/api/settings/path/reset", {"all": True})):
            r = self.post(url, **body)
            self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "not_admin"), url)
        st = self.state()
        self.assertFalse(st["master_edit"]["possible"])
        self.assertFalse(st["master_edit"]["setup"])

    def test_reset_item_and_all(self):
        self.use_master([(self.login, "", "Administrator")])
        r = self.post("/api/settings/path/reset", key="wav")
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "locked"))
        self.assertTrue(self.unlock()["ok"])
        t = str(self.tmp) + "/"
        for key in ("quality", "wav", "history"):
            paths.set_path(key, t)
        paths.set_path("lot", t, paths.FIRST)
        paths.set_path("lot", t, paths.SECOND)
        paths.set_second_name("SIKALOTNOW", "SIKALOT_2")
        with mock.patch.object(trace, "event", wraps=trace.event) as ev:
            st = self.post("/api/settings/path/reset", key="wav").get_json()
        self.assertEqual(paths.get("wav"), paths.BY_KEY["wav"].path)
        self.assertEqual(paths.get("quality"), t)                            # ほかの項目はそのまま
        self.assertEqual(st["reset"]["keys"], ["wav"])
        data = [c.kwargs.get("data") for c in ev.call_args_list if c.args and c.args[0] == "settings.path"][-1]
        self.assertEqual(data["wav"], {"before": {"音声": t}, "after": {"音声": paths.BY_KEY["wav"].path}})
        st = self.post("/api/settings/path/reset", key="lot").get_json()     # 第1・第2・2つ目の名前をまとめて
        lot = [p for p in st["paths"] if p["key"] == "lot"][0]
        self.assertEqual((lot["first"], lot["second"]), (paths.BY_KEY["lot"].path, paths.BY_KEY["lot"].second))
        self.assertEqual(paths.second_name("SIKALOTNOW"), "SIKALOT.sqlite3")
        paths.set_path("lot", t, paths.SECOND)
        paths.set_mode(paths.OFFLINE)
        self.addCleanup(paths.set_mode, paths.ONLINE)
        with mock.patch.object(trace, "event", wraps=trace.event) as ev:
            st = self.post("/api/settings/path/reset", all=True).get_json()
        for i in paths.ITEMS:
            self.assertFalse(paths.item_changed(i.key), i.key)
            if i.key != "lot":
                self.assertEqual(paths.get(i.key), paths.slot_default(i.key), i.key)   # 画像のフォルダは元のフォルダの下
        self.assertEqual(paths.mode(), paths.OFFLINE)                        # モードは変えない
        # 包装仕様書・品質写真は 品証やり取り の下(決めていなければ)なので、品証やり取りを戻すと一緒に変わる
        self.assertEqual(set(st["reset"]["keys"]), {"quality", "history", "lot", "access_master", "pack_spec", "quality_photo"})
        data = [c.kwargs.get("data") for c in ev.call_args_list if c.args and c.args[0] == "settings.path"][-1]
        self.assertEqual(data["quality"]["before"], {"品証やり取り": t})
        # アクセス権限マスタも既定(空)に戻ったので、設定の途中に戻る
        self.assertTrue(st["master_edit"]["setup"])
        self.assertEqual(st["rights"]["source"], "unset")

    def test_terminal_auto_and_manual_needs_password(self):
        self.use_master([(self.login, self.pc, "バランサー"), ("yamada", "", "Administrator")])
        st = self.post("/api/settings/rights/reload").get_json()
        self.assertEqual(st["terminal"]["facility"], "バランサー")
        self.assertIn("バランサー", st["rights"]["applied"])
        r = self.post("/api/settings/terminal", facility="HVC")              # 一般・鍵なし
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "not_admin"))
        self.use_master([(self.login, self.pc, "バランサー"), (self.login, "", "Administrator")])
        r = self.post("/api/settings/terminal", facility="HVC")
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "locked"))
        self.unlock()
        self.assertEqual(self.post("/api/settings/terminal", facility="HVC").status_code, 200)
        st = self.post("/api/settings/rights/reload").get_json()
        self.assertEqual(st["terminal"]["facility"], "HVC")                  # 表が同じなら手で変えたまま

    def test_unset_terminal_anyone(self):
        self.use_master([("yamada", "", "Administrator")])
        self.assertEqual(self.post("/api/settings/terminal", facility="LVC").status_code, 200)   # 未設定なら誰でも
        self.assertEqual(self.post("/api/settings/terminal", facility="HVC").status_code, 403)
        self.assertEqual(self.post("/api/settings/mode", mode="offline").status_code, 403)
        self.assertEqual(self.post("/api/settings/path/check", key="wav").status_code, 200)       # 確認は誰でも


class IsolationTest(unittest.TestCase):
    """試験が本番のパス・本番のローカル領域を読み書きしない(tests/_isolation.py)。"""

    def test_production_defaults_are_guarded(self):
        guard = os.environ.get(paths.GUARD_ENV)
        self.assertTrue(guard)
        tmp = Path(tempfile.gettempdir()).resolve()
        for i in paths.ITEMS:
            for text in filter(None, (i.path, i.second)):
                local = paths.to_local(text).resolve()
                self.assertTrue(str(local).startswith(str(Path(guard).resolve())), (i.key, local))
        self.assertTrue(str(A._cache_file().resolve()).startswith(str(tmp)))
        self.assertTrue(str(Path(config.USER_CONFIG_PATH).resolve()).startswith(str(tmp)))

    def test_temp_paths_are_not_rewritten(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(paths.to_local(d + "/"), Path(d + "/"))


class AdminPasswordTest(unittest.TestCase):
    """管理者パスワード(基盤。all-tools と同じ作り): 既定 nisk・変えたものは平文で持たない。"""

    def setUp(self):
        self.ap = admin_password
        self.ap.session.lock()
        self.addCleanup(self.ap.session.lock)
        self.addCleanup(user_settings.save, config.KEY_ADMIN_PASSWORD, None)
        user_settings.save(config.KEY_ADMIN_PASSWORD, None)

    def test_default_and_change(self):
        self.assertTrue(self.ap.verify("nisk"))
        self.assertFalse(self.ap.verify("NISK"))
        self.assertTrue(self.ap.change("nisk", "新しい鍵", "新しい鍵").ok)
        stored = user_settings.get(config.KEY_ADMIN_PASSWORD)
        self.assertNotIn("新しい鍵", stored)
        self.assertTrue(self.ap.verify("新しい鍵"))


if __name__ == "__main__":
    unittest.main()
