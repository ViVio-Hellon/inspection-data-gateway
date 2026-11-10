"""音の設定(D29: 出来事ごとに 鳴らさない / 音声フォルダのファイル を選ぶ。all-tools と同じ決まり)"""
from __future__ import annotations

import re
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from tests._web import make_client
from ufkensa import auth, user_settings
from ufkensa.kensa import ainput, calibration, cask_session, judge, paths, report_model, sound, workers, zenkasho_session
from ufkensa.kensa.flow import Message

ROOT = Path(__file__).resolve().parent.parent

#: 移す前に呼び出し側へ直書きされていたファイル(VBA の C:\VBAWAV の .wav)。既定はこれと同じでなければならない
OLD_FILES = {
    "kensa_ok": "chime2.wav", "kensa_ng": "tada.wav", "kensa_diag": "chimes.wav",          # judge.SOUND_*(補助:282-358)
    "kensa_need": "必要です.wav", "kensa_right_angle": "直角度.wav",                         # Judg:2072 / 2105
    "kensa_abnormal": "異常値.wav",                                                         # UF異常値発見:29(移植漏れを足した)
    "kensa_calibration": "確認作業.wav", "kensa_worker_unregistered": "ナンバーNO.wav",
    "kensa_advance_sheet": "ありますか.wav", "kensa_wait": "おまちください.wav",
    "kensa_print": "印刷.wav", "kensa_done": "完了.wav",
    "kensa_reprint": "印刷.wav", "kensa_reprint_done": "完了.wav",                          # report_model.print_only
    "cask_ok": "chime2.wav", "cask_ng": "tada.wav", "cask_diag": "chimes.wav",
    "cask_print": "印刷.wav", "cask_done": "完了.wav",
    "zenkasho_ok": "chime2.wav", "zenkasho_ng": "tada.wav", "zenkasho_diag": "chimes.wav",
    "zenkasho_wait": "おまちください.wav", "zenkasho_print": "印刷.wav", "zenkasho_done": "完了.wav",
    "ainput_height": "chimes.wav", "ainput_wait": "お待ちください.wav", "ainput_done": "完了.wav",
    "calibration_ok": "chime2.wav", "calibration_ng": "tada.wav",
}


class _Folder(unittest.TestCase):
    """音声フォルダ(使用パス「音声」)を一時フォルダに向け、音の設定を空にする。"""

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = Path(tmp.name) / "VBAWAV"
        self.dir.mkdir()
        paths.set_path("wav", str(self.dir) + "/")
        self.addCleanup(paths.reset_path, "wav")
        for k in sound.SOUND_KEYS:
            user_settings.delete(sound.setting_key(k))
            self.addCleanup(user_settings.delete, sound.setting_key(k))

    def put(self, name: str, data: bytes = b"RIFF0000WAVEfmt ") -> Path:
        p = self.dir / name
        p.write_bytes(data)
        return p


class CatalogTest(unittest.TestCase):
    def test_defaults_are_the_old_vba_files(self):
        self.assertEqual({s.key: s.default_file for s in sound.SOUNDS}, OLD_FILES)

    def test_every_event_has_label_place_when(self):
        for s in sound.SOUNDS:
            with self.subTest(s.key):
                self.assertTrue(s.label and s.place and s.when and s.message)
                self.assertIn(Path(s.default_file).suffix.lower(), sound.AUDIO_SUFFIXES)
                self.assertNotIn(s.key, ("state", "cues", "preview"))          # /api/sound/<key> と被らない

    def test_call_sites_use_event_keys(self):
        self.assertEqual((judge.SOUND_OK, judge.SOUND_NG, judge.SOUND_DIAG), ("kensa_ok", "kensa_ng", "kensa_diag"))
        self.assertEqual((judge.SOUND_NEED, judge.SOUND_RIGHT_ANGLE), ("kensa_need", "kensa_right_angle"))
        self.assertEqual((calibration.PASS_SOUND, calibration.FAIL_SOUND), ("calibration_ok", "calibration_ng"))
        self.assertEqual(workers.UNREGISTERED_SOUND, "kensa_worker_unregistered")
        self.assertEqual((ainput.CHIMES, ainput.WAIT_SOUND, ainput.DONE_SOUND),
                         ("ainput_height", "ainput_wait", "ainput_done"))
        self.assertEqual(cask_session.SOUNDS, {"音楽合格": "cask_ok", "音楽不合格": "cask_ng", "音楽対角": "cask_diag"})
        self.assertEqual(zenkasho_session.SOUNDS,
                         {"音楽合格": "zenkasho_ok", "音楽不合格": "zenkasho_ng", "音楽対角": "zenkasho_diag"})

    def test_no_file_names_at_call_sites(self):
        """呼び出し側に音声ファイル名を直書きしない(出来事の一覧 sound.py だけが既定のファイルを持つ)。"""
        pat = re.compile(r'"[^"\n]+\.(wav|mp3|ogg|m4a|aac|flac)"', re.I)
        hits = []
        for p in list((ROOT / "ufkensa").rglob("*.py")) + list((ROOT / "app").rglob("*.py")) \
                + list((ROOT / "app" / "static" / "js").rglob("*.js")):
            if p.name == "sound.py" and p.parent.name == "kensa":
                continue
            for i, line in enumerate(p.read_text(encoding="utf-8").splitlines(), 1):
                if pat.search(line) and not line.lstrip().startswith(("#", "//", "*")):
                    hits.append(f"{p.relative_to(ROOT)}:{i}")
        self.assertEqual(hits, [])

    def test_every_screen_uses_the_shared_player(self):
        # 整理 6: 知らせの音は共通の op()(screen_op.js)が鳴らす。検査は自分の showMessage で鳴らす
        shared = (ROOT / "app/static/js/screen_op.js").read_text(encoding="utf-8")
        self.assertIn('from "./sound.js"', shared)
        self.assertIn("playMessage(m)", shared)
        for v in ("kensa", "cask", "zenkasho", "ainput", "calibration"):
            js = (ROOT / "app/static/js/views" / f"{v}.js").read_text(encoding="utf-8")
            with self.subTest(v):
                uses_shared = 'from "../screen_op.js"' in js and "showMessage" not in js
                if not uses_shared:
                    self.assertIn('from "../sound.js"', js)
                    self.assertIn("playMessage(m)", js)
                self.assertNotIn("/api/kensa/sound", js)


class ChoiceTest(_Folder):
    def test_default_then_choice_then_off(self):
        self.assertEqual(sound.chosen("kensa_ok"), "chime2.wav")
        sound.set_choice("kensa_ok", "ピンポン.mp3")
        self.assertEqual(sound.chosen("kensa_ok"), "ピンポン.mp3")
        self.assertEqual(sound.chosen("cask_ok"), "chime2.wav")                 # 画面ごとに別
        self.assertEqual(sound.chosen("calibration_ok"), "chime2.wav")
        sound.set_choice("kensa_ok", sound.SOUND_OFF)
        self.assertEqual((sound.chosen("kensa_ok"), sound.enabled("kensa_ok")), ("", False))
        self.assertEqual(user_settings.get("sound_file_kensa_ok"), "鳴らさない")
        sound.reset("kensa_ok")
        self.assertEqual(sound.chosen("kensa_ok"), "chime2.wav")

    def test_refuses_paths_and_other_types(self):
        for bad in ("..\\x.wav", "a/b.wav", "C:\\VBAWAV\\x.wav", "x.txt", "", ".."):
            with self.subTest(bad):
                with self.assertRaises(ValueError):
                    sound.set_choice("kensa_ok", bad)
        with self.assertRaises(KeyError):
            sound.set_choice("無い", "a.wav")
        for ok in sound.AUDIO_SUFFIXES:
            sound.set_choice("kensa_ok", "a" + ok.upper())

    def test_message_carries_key_only_when_enabled(self):
        m = Message("", sound="kensa_ng")
        self.assertEqual((m.to_dict()["sound"], m.to_dict()["sound_text"]), ("kensa_ng", "不合格"))
        sound.set_choice("kensa_ng", sound.SOUND_OFF)
        self.assertEqual(m.to_dict()["sound"], "")                              # 鳴らさない
        self.assertNotIn("sound_text", m.to_dict())
        self.assertEqual(Message("", sound="tada.wav").to_dict()["sound"], "")  # 一覧に無いもの(ファイル名)は渡さない

    def test_missing_file_keeps_choice(self):
        self.put("chime2.wav")
        sound.set_choice("cask_ok", "無い.wav")
        rows = {r["key"]: r for r in sound.overview()["sounds"]}
        self.assertEqual((rows["cask_ok"]["file"], rows["cask_ok"]["exists"]), ("無い.wav", False))
        self.assertIn("フォルダにありません", rows["cask_ok"]["note"])
        self.assertTrue(rows["kensa_ok"]["exists"])
        self.assertEqual(sound.playable_keys(), ["kensa_ok", "zenkasho_ok", "calibration_ok"])

    def test_reprint_and_kensa_session_keys(self):
        self.assertEqual(sound.chosen(report_model.snd.KENSA_REPRINT), "印刷.wav")


class RouteTest(_Folder):
    def setUp(self):
        super().setUp()
        self.c = make_client()
        self.h = {"X-Tool-Token": self.c.application.config["TOKEN"]}

    def get(self, url, **kw):
        return self.c.get(url, headers=self.h, **kw)

    def test_served_only_by_event_key(self):
        self.put("chime2.wav", b"RIFF-chime2")
        self.put("別の音.wav", b"RIFF-other")
        r = self.get("/api/sound/kensa_ok")
        self.assertEqual((r.status_code, r.data, r.mimetype), (200, b"RIFF-chime2", "audio/wav"))
        sound.set_choice("kensa_ok", "別の音.wav")                                 # 設定で鳴る音が変わる
        self.assertEqual(self.get("/api/sound/kensa_ok").data, b"RIFF-other")
        self.assertEqual(self.get("/api/sound/cask_ok").data, b"RIFF-chime2")     # キャスクは前のまま
        sound.set_choice("kensa_ok", sound.SOUND_OFF)
        r = self.get("/api/sound/kensa_ok")
        self.assertEqual((r.status_code, r.get_json()["error"]["code"]), (404, "off"))
        self.assertEqual(self.get("/api/sound/kensa_done").status_code, 404)      # ファイルが無い
        for bad in ("/api/sound/chime2.wav", "/api/sound/..%2Fchime2.wav", "/api/sound/無い"):
            with self.subTest(bad):
                self.assertEqual(self.get(bad).status_code, 404)
        self.assertEqual(self.c.get("/api/sound/cask_ok").status_code, 401)       # トークンが要る
        self.assertEqual(self.c.get("/api/sound/cask_ok?t=" + self.h["X-Tool-Token"]).status_code, 200)

    def test_cues(self):
        self.put("tada.wav")
        sound.set_choice("cask_ng", sound.SOUND_OFF)
        self.assertEqual(self.get("/api/sound/cues").get_json()["cues"], ["kensa_ng", "zenkasho_ng", "calibration_ng"])

    def test_preview_inside_folder_only(self):
        self.put("試し.mp3", b"ID3")
        r = self.get("/api/sound/preview?name=試し.mp3")
        self.assertEqual((r.status_code, r.data, r.mimetype), (200, b"ID3", "audio/mpeg"))
        self.assertEqual(self.get("/api/sound/preview?name=無い.wav").status_code, 404)
        outside = self.dir.parent / "外.wav"
        outside.write_bytes(b"RIFF")
        for bad in ("../外.wav", "..\\外.wav", str(outside), "sub/x.wav", "x.txt", ""):
            with self.subTest(bad):
                self.assertEqual(self.get("/api/sound/preview", query_string={"name": bad}).status_code, 400)
        self.assertEqual(self.c.get("/api/sound/preview?name=試し.mp3").status_code, 401)

    def test_settings_api_without_admin_password(self):
        self.put("chime2.wav")
        self.put("ポン.ogg")
        with mock.patch.object(auth, "can_edit", return_value=False), \
                mock.patch.object(auth, "master_edit_ok", return_value=False):
            st = self.get("/api/sound/state").get_json()
            self.assertEqual(st["dir"], str(self.dir) + "/")
            self.assertEqual((st["files"], st["dir_error"], st["off"]), (["chime2.wav", "ポン.ogg"], "", "鳴らさない"))
            self.assertEqual([r["key"] for r in st["sounds"]], list(OLD_FILES))
            r = self.c.post("/api/settings/sound", json={"key": "kensa_ok", "file": "ポン.ogg"}, headers=self.h)
            self.assertEqual(r.status_code, 200)                                  # 一般の人でも選べる
            row = [x for x in r.get_json()["sounds"] if x["key"] == "kensa_ok"][0]
            self.assertEqual((row["file"], row["custom"], row["exists"], row["default_file"]),
                             ("ポン.ogg", True, True, "chime2.wav"))
            r = self.c.post("/api/settings/sound", json={"key": "kensa_ok", "file": "鳴らさない"}, headers=self.h)
            row = [x for x in r.get_json()["sounds"] if x["key"] == "kensa_ok"][0]
            self.assertEqual((row["off"], row["note"]), (True, "鳴らしません"))
            self.assertNotIn("kensa_ok", r.get_json()["cues"])
            r = self.c.post("/api/settings/sound", json={"key": "kensa_ok", "reset": True}, headers=self.h)
            self.assertEqual([x for x in r.get_json()["sounds"] if x["key"] == "kensa_ok"][0]["file"], "chime2.wav")
            self.assertIsNone(user_settings.get("sound_file_kensa_ok"))
            for body in ({"key": "kensa_ok", "file": "..\\x.wav"}, {"key": "kensa_ok", "file": "x.exe"},
                         {"key": "無い", "file": "a.wav"}):
                with self.subTest(body):
                    self.assertEqual(self.c.post("/api/settings/sound", json=body, headers=self.h).status_code, 400)

    def test_folder_missing(self):
        paths.set_path("wav", str(self.dir / "無い") + "/")
        st = self.get("/api/sound/state").get_json()
        self.assertIn("音声フォルダがありません", st["dir_error"])
        self.assertEqual(st["cues"], [])
        self.assertTrue(all(not r["exists"] for r in st["sounds"]))

    def test_settings_page_has_sound_tab(self):
        html = self.c.get("/settings").get_data(as_text=True)
        self.assertIn('data-key="sound"', html)
        self.assertIn('id="soundTable"', html)
        self.assertIn('href="?tab=paths"', html)


class SessionSoundTest(_Folder):
    """呼び出し側は出来事の鍵を出し、鳴らさないと決めた出来事は画面へ渡らない。"""

    def test_calibration_pass_is_calibration_event(self):
        m = Message("", sound=calibration.PASS_SOUND)
        self.assertEqual(m.to_dict()["sound"], "calibration_ok")
        sound.set_choice("calibration_ok", sound.SOUND_OFF)
        self.assertEqual(m.to_dict()["sound"], "")
        self.assertEqual(Message("", sound=judge.SOUND_OK).to_dict()["sound"], "kensa_ok")   # 検査は別に選ぶ


if __name__ == "__main__":
    unittest.main()


from tests.test_kensa_session import SessionBase  # noqa: E402


class KensaAbnormalSoundTest(SessionBase):
    """移植漏れだった UF異常値発見 の音(音楽異常値 → 異常値.wav)を出来事 kensa_abnormal で鳴らす。"""

    def test_abnormal_value_rings_abnormal_event(self) -> None:
        self.register()
        r = self.s.receive("DT10001+0000030.000M")                       # 厚の公差の2倍を超える(異常値除去 O:1321)
        self.assertIn("異常", [m.title for m in r.messages])
        self.assertIn(sound.KENSA_ABNORMAL, [m.sound for m in r.messages])
