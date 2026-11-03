"""音(どの出来事で、どの音声ファイルを鳴らすか)── all-tools 日報ツールと同じ決まり(D29)

【鳴らすのはブラウザ、決めるのはサーバ】
VBA は `mciSendString("Play C:\\VBAWAV\\…")` でその場で鳴らしていた(補助:251-358)。Web 版で
サーバ側から鳴らすと、画面を見ている人に届く保証がない。そこで分ける:

    ここ(サーバ)  … 「どの出来事か」(業務の部品は出来事の鍵だけを返す)と、
                      「その出来事で何を鳴らすか」(設定)を決める
    ブラウザ      … `app/static/js/sound.js` が `/api/sound/<出来事>` を鳴らす

【出来事 = 鳴らす場所 × 意味】
画面ごとに別々に選べるよう、同じ音(chime2.wav)でも「検査の合格」「キャスクの合格」「校正の合格」は
別の出来事にした。**既定のファイルは VBA が鳴らしていたファイルそのもの**(何も設定しなければ今までどおり)。
どの出来事で何を鳴らすかは 設定 →「音」で選ぶ ── 「鳴らさない」か、音声フォルダ(使用パス「音声」)のファイル。

- 設定は `user_settings` の `sound_file_<出来事>`。入っていなければ既定、`鳴らさない` という字なら鳴らさない
- **ファイル名は呼び出し側に持たない**(呼び出し側は出来事の鍵を `Message.sound` に入れる)
- ファイルが無くても選択は残す(画面に「フォルダにありません」と出す)。音は補助で、文言は画面に出る
- 配るのは出来事の鍵で決まるファイルだけ(`/api/sound/<出来事>`)。試聴はフォルダの中のファイル名だけ
  (`/api/sound/preview?name=`)。道(フォルダの区切り)は受け取らない
- 変えるのに管理者パスワードは要らない(間違えても音が鳴らないだけ。all-tools と同じ考え)

出来事を足すときは、`SOUNDS` に1行と、鳴らすきっかけ(`Message(sound=<鍵>)`)を1か所足す。
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

from ufkensa import user_settings

from . import paths

#: 音声ファイルとして受け付ける拡張子(配る口・設定の選択肢・フォルダの一覧はこれ1つを見る)
AUDIO_SUFFIXES: tuple[str, ...] = (".wav", ".mp3", ".ogg", ".m4a", ".aac", ".flac")

#: 「鳴らさない」を選んだときに設定へ入れる字(拡張子が無いのでファイル名とは混ざらない)
SOUND_OFF = "鳴らさない"

#: 設定の鍵の頭(`sound_file_<出来事>`)
SETTING_PREFIX = "sound_file_"

#: 音声フォルダ = 使用パス「音声」(VBA の C:\VBAWAV\)
FOLDER_KEY = "wav"

MIME = {".wav": "audio/wav", ".mp3": "audio/mpeg", ".ogg": "audio/ogg", ".m4a": "audio/mp4",
        ".aac": "audio/aac", ".flac": "audio/flac"}


@dataclass(frozen=True)
class SoundSpec:
    """1つの出来事。`message` は鳴らせなかったとき(自動再生の制限)に画面へ出す文言。"""

    key: str
    label: str           # 設定画面の名前
    place: str           # 鳴らす場所(画面)
    when: str            # いつ鳴るか
    default_file: str    # 既定のファイル(VBA が鳴らしていたもの)
    message: str         # 鳴らせなかったときの文言
    source: str = ""     # VBA の場所


KENSA, CASK, ZENKASHO, AINPUT, CALIBRATION = "検査", "キャスク", "全ヵ所", "A入力", "校正"

# 鍵(呼び出し側はこれを使う。文字列を直書きしない)
KENSA_OK, KENSA_NG, KENSA_DIAG = "kensa_ok", "kensa_ng", "kensa_diag"
KENSA_NEED, KENSA_RIGHT_ANGLE, KENSA_ABNORMAL = "kensa_need", "kensa_right_angle", "kensa_abnormal"
KENSA_CALIBRATION, KENSA_WORKER_UNREGISTERED = "kensa_calibration", "kensa_worker_unregistered"
KENSA_ADVANCE_SHEET, KENSA_WAIT, KENSA_PRINT = "kensa_advance_sheet", "kensa_wait", "kensa_print"
KENSA_DONE, KENSA_REPRINT, KENSA_REPRINT_DONE = "kensa_done", "kensa_reprint", "kensa_reprint_done"
CASK_OK, CASK_NG, CASK_DIAG, CASK_PRINT, CASK_DONE = "cask_ok", "cask_ng", "cask_diag", "cask_print", "cask_done"
ZENKASHO_OK, ZENKASHO_NG, ZENKASHO_DIAG = "zenkasho_ok", "zenkasho_ng", "zenkasho_diag"
ZENKASHO_WAIT, ZENKASHO_PRINT, ZENKASHO_DONE = "zenkasho_wait", "zenkasho_print", "zenkasho_done"
AINPUT_HEIGHT, AINPUT_WAIT, AINPUT_DONE = "ainput_height", "ainput_wait", "ainput_done"
CALIBRATION_OK, CALIBRATION_NG = "calibration_ok", "calibration_ng"

SOUNDS: tuple[SoundSpec, ...] = (
    # ---- 検査(通常検査 UF1)
    SoundSpec(KENSA_OK, "合格", KENSA, "測定値を判定して公差内だったとき", "chime2.wav", "合格", "補助:305 音楽合格"),
    SoundSpec(KENSA_NG, "不合格", KENSA, "測定値を判定して公差外だったとき", "tada.wav", "不合格", "補助:357 音楽不合格"),
    SoundSpec(KENSA_DIAG, "対角線", KENSA, "対角線の欄を判定するとき", "chimes.wav", "対角線を判定しました",
              "Judg 音楽対角"),
    SoundSpec(KENSA_NEED, "必要です", KENSA, "先行データ票が必要なロット(ロット読込・検査員登録・資材の案内)",
              "必要です.wav", "先行データ票が必要です", "Judg:2072 / UF1:4326 / UF1:4857"),
    SoundSpec(KENSA_RIGHT_ANGLE, "直角度", KENSA, "資材の案内に「直角度の測定」があるとき", "直角度.wav",
              "直角度の測定があります", "Judg:2105"),
    SoundSpec(KENSA_ABNORMAL, "異常値", KENSA, "入力が異常値(公差の2倍など)で消したとき", "異常値.wav",
              "異常値です", "UF異常値発見:29 音楽異常値"),
    SoundSpec(KENSA_CALIBRATION, "校正の確認作業", KENSA, "検査員を登録して、その直の校正がまだのとき",
              "確認作業.wav", "校正(確認作業)をしてください", "UF1:4117 音楽検査員登録"),
    SoundSpec(KENSA_WORKER_UNREGISTERED, "名簿にない作業者", KENSA, "作業者カードの番号が名簿(伝送用ファイル)にないとき",
              "ナンバーNO.wav", "名簿にない番号です", "Worker:74 音楽未登録"),
    SoundSpec(KENSA_ADVANCE_SHEET, "先行データ票はありますか", KENSA, "ロット登録で先行データ票の有無を聞くとき",
              "ありますか.wav", "先行データ票はありますか？", "UF1:3647"),
    SoundSpec(KENSA_WAIT, "お待ちください(保存)", KENSA, "保存を始めたとき", "おまちください.wav",
              "お待ちください", "UF1:3715"),
    SoundSpec(KENSA_PRINT, "印刷する？", KENSA, "保存のあと「印刷する？」と聞くとき", "印刷.wav", "印刷する？",
              "UF1:3752"),
    SoundSpec(KENSA_DONE, "完了(保存)", KENSA, "保存・印刷が済んだとき", "完了.wav", "完了しました", "UF1:3837"),
    SoundSpec(KENSA_REPRINT, "再印刷しますか", KENSA, "印刷済みのロットで印刷を押したとき", "印刷.wav",
              "再印刷しますか？", "UF1:3578"),
    SoundSpec(KENSA_REPRINT_DONE, "完了(再印刷)", KENSA, "再印刷が済んだとき", "完了.wav", "再印刷が完了しました",
              "UF1:3592"),
    # ---- キャスク(UFCask)
    SoundSpec(CASK_OK, "合格", CASK, "測定値を判定して公差内だったとき", "chime2.wav", "合格", "補助:305 音楽合格"),
    SoundSpec(CASK_NG, "不合格", CASK, "測定値を判定して公差外だったとき", "tada.wav", "不合格", "補助:357 音楽不合格"),
    SoundSpec(CASK_DIAG, "対角線", CASK, "対角線の欄を判定するとき", "chimes.wav", "対角線を判定しました",
              "KySh 音楽対角"),
    SoundSpec(CASK_PRINT, "印刷", CASK, "印刷(保存)を始めたとき", "印刷.wav", "印刷します", "UFCask:138"),
    SoundSpec(CASK_DONE, "完了", CASK, "印刷・保存が済んだとき", "完了.wav", "完了しました", "UFCask:164"),
    # ---- 全ヵ所(UFMitsubishi。全ヵ所・H494 のシート)
    SoundSpec(ZENKASHO_OK, "合格", ZENKASHO, "測定値を判定して公差内だったとき", "chime2.wav", "合格",
              "補助:305 音楽合格"),
    SoundSpec(ZENKASHO_NG, "不合格", ZENKASHO, "測定値を判定して公差外だったとき", "tada.wav", "不合格",
              "補助:357 音楽不合格"),
    SoundSpec(ZENKASHO_DIAG, "対角線", ZENKASHO, "対角線の欄を判定するとき", "chimes.wav", "対角線を判定しました",
              "ZenSh / H494_SH 音楽対角"),
    SoundSpec(ZENKASHO_WAIT, "お待ちください", ZENKASHO, "保存・印刷を始めたとき", "おまちください.wav",
              "お待ちください", "UFMitsubishi:153"),
    SoundSpec(ZENKASHO_PRINT, "印刷", ZENKASHO, "印刷するとき", "印刷.wav", "印刷します", "UFMitsubishi:170"),
    SoundSpec(ZENKASHO_DONE, "完了", ZENKASHO, "保存・印刷が済んだとき", "完了.wav", "完了しました",
              "UFMitsubishi:217"),
    # ---- A入力(UFAinput)
    SoundSpec(AINPUT_HEIGHT, "山高さを読んだ", AINPUT, "測定器から山高さを読み取ったとき", "chimes.wav",
              "山高さを読みました", "UFAinput:729 音楽対角"),
    SoundSpec(AINPUT_WAIT, "お待ちください", AINPUT, "A入力の出力を始めたとき", "お待ちください.wav",
              "お待ちください", "UFAinput:1050"),
    SoundSpec(AINPUT_DONE, "完了", AINPUT, "A入力の出力が済んだとき", "完了.wav", "完了しました", "UFAinput:1086"),
    # ---- 校正(Proofreading)
    SoundSpec(CALIBRATION_OK, "合格", CALIBRATION, "品証スケール・ゲージの校正が合格したとき", "chime2.wav", "合格",
              "Proofreading:124 / :297"),
    SoundSpec(CALIBRATION_NG, "不合格", CALIBRATION, "品証スケール・ゲージの校正が不合格のとき", "tada.wav", "不合格",
              "Proofreading:148 / :361"),
)

BY_KEY: dict[str, SoundSpec] = {s.key: s for s in SOUNDS}
SOUND_KEYS: tuple[str, ...] = tuple(BY_KEY)

#: VBA の音の手続き(シートのモジュールが積む名前)→ 画面ごとの出来事
VBA_PROCS = ("音楽合格", "音楽不合格", "音楽対角")


def by_procedure(ok: str, ng: str, diag: str) -> dict[str, str]:
    """`音楽合格 / 音楽不合格 / 音楽対角`(補助:282-358)を、その画面の出来事へ読み替える表。"""
    return dict(zip(VBA_PROCS, (ok, ng, diag)))


def spec(key: str) -> Optional[SoundSpec]:
    return BY_KEY.get(key)


def setting_key(key: str) -> str:
    return SETTING_PREFIX + key


def is_audio_name(name: str) -> bool:
    """フォルダの中のファイル名1つとして通せるか(フォルダの区切り・.. は通さない)。"""
    text = (name or "").strip()
    if not text or "/" in text or "\\" in text or ":" in text or text in (".", ".."):
        return False
    return Path(text).suffix.lower() in AUDIO_SUFFIXES


def stored(key: str) -> Optional[str]:
    """設定に入っている値(無ければ None = 既定)。"""
    v = user_settings.get(setting_key(key))
    return None if v is None else str(v)


def chosen(key: str) -> str:
    """その出来事で鳴らすファイル名。鳴らさない・知らない出来事は ""。"""
    s = spec(key)
    if s is None:
        return ""
    v = stored(key)
    if v is None:
        return s.default_file
    v = v.strip()
    if v == SOUND_OFF or not is_audio_name(v):
        return ""
    return v


def enabled(key: str) -> bool:
    """鳴らすと決めてあるか(ファイルがあるかは見ない)。"""
    return chosen(key) != ""


def set_choice(key: str, value: str) -> None:
    """選ぶ。`鳴らさない` か、音声ファイルの名前(フォルダに無くてもよい)。"""
    if key not in BY_KEY:
        raise KeyError(key)
    value = (value or "").strip()
    if value != SOUND_OFF and not is_audio_name(value):
        raise ValueError("「鳴らさない」か、音声ファイルの名前(" + " / ".join(AUDIO_SUFFIXES)
                         + "。フォルダは含めない)を選んでください")
    user_settings.save(setting_key(key), value)


def reset(key: str) -> None:
    """既定(VBA のファイル)に戻す。"""
    if key not in BY_KEY:
        raise KeyError(key)
    user_settings.delete(setting_key(key))


def folder_text() -> str:
    return paths.get(FOLDER_KEY)


def folder() -> Path:
    return paths.to_local(folder_text())


def file_path(name: str) -> Optional[Path]:
    """音声フォルダの中のファイル(試聴用)。フォルダの外は指せない。無ければ None。"""
    if not is_audio_name(name):
        return None
    p = folder() / name.strip()
    try:
        return p if p.is_file() else None
    except OSError:
        return None


def sound_path(key: str) -> Optional[Path]:
    """出来事で鳴らすファイル(あれば)。"""
    name = chosen(key)
    return file_path(name) if name else None


def playable(key: str) -> bool:
    return sound_path(key) is not None


def playable_keys() -> list[str]:
    """いま鳴らせる出来事(鳴らすと決めてあり、ファイルがある)。画面はここにあるものだけ鳴らしに行く。"""
    try:
        names = set(folder_files()[0])
    except OSError:
        return []
    return [s.key for s in SOUNDS if chosen(s.key) in names]


def folder_files() -> tuple[list[str], str]:
    """音声フォルダの音声ファイル(名前順)と、読めなかったときの理由。"""
    d = folder()
    try:
        names = sorted((p.name for p in d.iterdir() if p.is_file() and is_audio_name(p.name)), key=str.casefold)
    except FileNotFoundError:
        return [], f"音声フォルダがありません: {folder_text()}"
    except OSError as e:
        return [], f"音声フォルダを読めません: {folder_text()} ({e})"
    return names, ""


def mimetype(path: Path) -> str:
    return MIME.get(path.suffix.lower(), "application/octet-stream")


def describe(key: str, files: Optional[set[str]] = None) -> dict[str, Any]:
    """設定画面の1行。"""
    s = BY_KEY[key]
    if files is None:
        files = set(folder_files()[0])
    v = stored(key)
    name = chosen(key)
    off = name == ""
    present = (not off) and name in files
    if off:
        note = "鳴らしません"
    elif present:
        note = ""
    else:
        note = "フォルダにありません(音は鳴らず、画面の文言だけになります)"
    return {"key": s.key, "label": s.label, "place": s.place, "when": s.when, "message": s.message,
            "source": s.source, "default_file": s.default_file, "file": name, "off": off,
            "custom": v is not None, "exists": present, "playable": present, "note": note}


def overview() -> dict[str, Any]:
    """設定画面の「音」の面(いまどこを見ているか・フォルダの中身・出来事ごとの選択)。"""
    files, problem = folder_files()
    fs = set(files)
    return {"dir": folder_text(), "dir_local": str(folder()), "dir_error": problem, "files": files,
            "off": SOUND_OFF, "suffixes": list(AUDIO_SUFFIXES), "folder_key": FOLDER_KEY,
            "places": list(dict.fromkeys(s.place for s in SOUNDS)),
            "sounds": [describe(s.key, fs) for s in SOUNDS]}


def message_text(key: str) -> str:
    s = spec(key)
    return s.message if s else ""
