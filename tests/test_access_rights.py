"""権限(判断事項 D12: python-web-tools と同じ「アクセス権限」表)

表は 管理番号・ログインID・PC名・権限・有効・備考(ほかのツールと共有)。
UF検査判定の権限は `ufkensa:admin`(設定・取り込み。VBA の weekend99 の代わり)と
`ufkensa:calibration_skip`(校正をとばす。VBA の jump の代わり)。
ほかのツールの行は読み飛ばす・空欄は問わない・両方空は効かない・大文字小文字と全角半角は区別しない・
有効 0 は効かない・ufkensa の行が 1 行も無ければ誰でも管理者・届かなければ前回の結果。
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
from ufkensa import access_rights, auth, config, user_settings
from ufkensa.kensa import paths


def master_db(path: Path, rows) -> Path:
    """python-web-tools の梱包資材マスタと同じ形の「アクセス権限」表。rows = (ログインID, PC名, 権限, 有効)。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    with closing(sqlite3.connect(path)) as c:
        c.execute('CREATE TABLE IF NOT EXISTS "アクセス権限" ("管理番号" INTEGER PRIMARY KEY, "ログインID" TEXT,'
                  ' "PC名" TEXT, "権限" TEXT, "有効" INTEGER, "備考" TEXT)')
        c.execute('DELETE FROM "アクセス権限"')
        c.executemany('INSERT INTO "アクセス権限" (ログインID, PC名, 権限, 有効) VALUES (?, ?, ?, ?)', rows)
        c.commit()
    return path


def use_master(case: unittest.TestCase, path: Path) -> Path:
    paths.set_path("access_master", str(path))
    case.addCleanup(paths.reset_path, "access_master")
    return path


class DecideTest(unittest.TestCase):
    def rows(self, *rows):
        return [{"ログインID": a, "PC名": b, "権限": c, "有効": d} for a, b, c, d in rows]

    def test_admin_rules(self):
        d = lambda rows, who="taro", pc="PC-1": access_rights.decide(self.rows(*rows), who, pc)[0]
        self.assertTrue(d([("taro", "PC-1", "ufkensa:admin", 1)]))
        self.assertTrue(d([("TARO", "pc-1", "UFKENSA:ADMIN", 1)]))                # 大文字小文字
        self.assertTrue(d([("ｔａｒｏ", "", "ufkensa:admin", 1)]))                 # 全角も・PC は問わない
        self.assertTrue(d([("", "PC-1", "ufkensa:admin", None)]))                 # 有効の空欄は有効
        self.assertFalse(d([("", "", "ufkensa:admin", 1)]))                       # 両方空は誰にも効かない
        self.assertFalse(d([("taro", "PC-2", "ufkensa:admin", 1)]))
        self.assertFalse(d([("taro", "PC-1", "ufkensa:calibration_skip", 1)]))    # 校正スキップは管理者ではない
        self.assertFalse(d([("taro", "PC-1", "編集可", 1)]))                      # db-viewer の書き方は読まない
        self.assertFalse(d([("taro", "PC-1", "mode:material", 1)]))               # ほかのツールの権限
        for off in (0, "0", "false", "×", "無効"):
            self.assertFalse(d([("taro", "PC-1", "ufkensa:admin", off)]), off)

    def test_codes(self):
        rows = self.rows(("taro", "", "ufkensa:calibration_skip", 1), ("taro", "", "mode:field", 1),
                         ("", "PC-1", "ufkensa:admin", 0))
        self.assertEqual(access_rights.codes_of(rows, "taro", "PC-1"), ["ufkensa:calibration_skip"])


class EvaluateTest(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.tmp = Path(tmp.name)
        env = mock.patch.dict(os.environ, {"UFKENSA_LOCAL_DIR": str(self.tmp / "local")})
        env.start()
        self.addCleanup(env.stop)
        self.master = use_master(self, master_db(self.tmp / "梱包資材マスタ.sqlite3", [
            ("taro", "PC-1", "ufkensa:admin", 1), ("hanako", "", "ufkensa:calibration_skip", 1),
            ("jiro", "", "mode:material", 1), ("", "PC-9", "ufkensa:admin", 0)]))

    def test_shared_master_decides(self):
        d = access_rights.evaluate("taro", "PC-1")
        self.assertEqual((d.edit, d.source, d.label), (True, "master", "管理者"))
        self.assertTrue(d.has(access_rights.ADMIN))
        d = access_rights.evaluate("hanako", "PC-1")
        self.assertEqual((d.edit, d.label), (False, "一般"))
        self.assertTrue(d.has(access_rights.CALIBRATION_SKIP))
        d = access_rights.evaluate("jiro", "PC-1")                                # ほかのツールの権限だけ
        self.assertEqual((d.edit, d.codes), (False, []))

    def test_no_ufkensa_rows_means_everyone_is_admin(self):
        master_db(self.master, [("jiro", "", "mode:material", 1)])                # ほかのツールの行だけ
        d = access_rights.evaluate("anyone", "PC-X")
        self.assertEqual((d.edit, d.source), (True, "unset"))
        self.assertTrue(d.has(access_rights.CALIBRATION_SKIP))
        self.assertIn("まだ無い", d.reason)

    def test_unreachable_uses_last_result_then_general(self):
        self.assertTrue(access_rights.evaluate("taro", "PC-1").edit)              # 読めた結果を覚える
        with mock.patch.object(access_rights, "read_rows", side_effect=OSError("共有フォルダに届きません")):
            d = access_rights.evaluate("taro", "PC-1")
            self.assertEqual((d.edit, d.source), (True, "cache"))
            self.assertIn("前回", d.reason)
            d = access_rights.evaluate("ken", "PC-2")                             # 前回の結果が無い人は一般
            self.assertEqual((d.edit, d.source), (False, "unreachable"))

    def test_without_master_path_uses_local_settings_file(self):
        paths.reset_path("access_master")
        self.assertNotEqual(access_rights.master_path(), self.master)


class SettingsApiRightsTest(unittest.TestCase):
    """設定画面: ラインは未設定なら誰でも決められる。決まったあと・使用パス・モードは管理者だけ。"""

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.tmp = Path(tmp.name)
        env = mock.patch.dict(os.environ, {"UFKENSA_LOCAL_DIR": str(self.tmp / "local"),
                                           "USERNAME": "hanako", "COMPUTERNAME": "PC-1"})
        env.start()
        self.addCleanup(env.stop)
        use_master(self, master_db(self.tmp / "梱包資材マスタ.sqlite3", [("taro", "", "ufkensa:admin", 1)]))
        user_settings.save("kensa_facility", "")
        self.addCleanup(user_settings.save, "kensa_facility", "")
        auth.end()
        self.addCleanup(auth.end)
        self.c = make_client()
        self.h = {"X-Tool-Token": self.c.application.config["TOKEN"]}

    def post(self, url, **body):
        return self.c.post(url, json=body, headers=self.h)

    def test_general_user(self):
        st = self.c.get("/api/settings/state", headers=self.h).get_json()
        self.assertFalse(st["admin"])
        self.assertTrue(st["terminal"]["unset"])
        self.assertEqual(self.post("/api/settings/terminal", facility="LVC").status_code, 200)   # 未設定なら誰でも
        self.assertEqual(self.post("/api/settings/terminal", facility="HVC").status_code, 403)   # 変えるのは管理者
        self.assertEqual(self.post("/api/settings/mode", mode="offline").status_code, 403)
        r = self.post("/api/settings/path", key="wav", slot="first", path=r"D:\\x")
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "not_admin"))
        auth.unlock_master("nisk")                     # 管理者でなければパスワードを入れても開かない
        self.assertEqual(self.post("/api/settings/path", key="wav", slot="first", path="x").status_code, 403)
        self.assertEqual(self.post("/api/settings/path/check", key="wav").status_code, 200)       # 確認は誰でも

    def test_admin(self):
        with mock.patch.dict(os.environ, {"USERNAME": "taro"}):
            auth.end()
            self.assertTrue(self.c.get("/api/settings/state", headers=self.h).get_json()["admin"])
            self.assertEqual(self.post("/api/settings/terminal", facility="丸徳", line_no="A3").status_code, 200)
            # 使用パスは管理者パスワードも要る(D26)。開けるまでは 403 locked
            r = self.post("/api/settings/path", key="wav", slot="first", path=r"D:\音")
            self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (403, "locked"))
            self.assertEqual(r.get_json()["error"]["message"], auth.master_edit_reason())
            self.assertTrue(auth.unlock_master("nisk")[0])
            self.addCleanup(auth.lock_master)
            st = self.post("/api/settings/path", key="wav", slot="first", path=r"D:\音").get_json()
            self.addCleanup(paths.reset_path, "wav")
            self.assertEqual([p for p in st["paths"] if p["key"] == "wav"][0]["path"], r"D:\音")
            self.assertEqual(st["terminal"]["label"], "中板A3")
            self.assertEqual(self.post("/api/settings/terminal", facility="無い").status_code, 400)


class AdminPasswordTest(unittest.TestCase):
    """管理者パスワード(基盤。all-tools と同じ作り): 既定 nisk・変えたものは平文で持たない・続けてまちがえると待つ。"""

    def setUp(self):
        from ufkensa import admin_password
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
        self.assertNotIn("新しい鍵", stored)                                      # 平文で持たない
        self.assertTrue(self.ap.verify("新しい鍵"))


if __name__ == "__main__":
    unittest.main()
