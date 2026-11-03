"""音(`/api/sound/*`)と 設定 →「音」の保存(`/api/settings/sound`)── D29

    サーバ  … どの出来事か(業務の部品が `Message.sound` に出来事の鍵を入れる)と、何を鳴らすか(設定)を決める
    ブラウザ … `app/static/js/sound.js` が鳴らす(自動再生の制限で鳴らせなければ文言を出す)

**配るのは出来事の鍵で決まるファイルだけ**(`/api/sound/<出来事>`)。道を受け取る口にすると、そこが
読み取り口になる。試聴は音声フォルダの中のファイル名1つだけ(`/api/sound/preview?name=`)。
どちらも `/api/` の下なので起動トークンが要る(`<audio>` は `?t=` で付ける)。

音の選択を変えるのに管理者パスワードは要らない(間違えても音が鳴らないだけ。all-tools と同じ)。
"""
from __future__ import annotations

from flask import Blueprint, jsonify, request, send_file

from ufkensa import trace
from ufkensa.kensa import sound
from ufkensa.logging_utils import get_logger

log = get_logger("app.routes.sound")

bp = Blueprint("sound", __name__)


def _err(code: str, message: str, status: int = 400):
    return jsonify({"error": {"code": code, "message": message}}), status


def _state() -> dict:
    out = sound.overview()
    out["cues"] = [s["key"] for s in out["sounds"] if s["playable"]]
    return out


@bp.get("/api/sound/state")
def state():
    """設定画面の「音」の面: 音声フォルダ(いまどこを見ているか)・中のファイル・出来事ごとの選択。"""
    return jsonify(_state())


@bp.get("/api/sound/cues")
def cues():
    """いま鳴らせる出来事(鳴らすと決めてあり、ファイルがある)。画面はここにあるものだけ鳴らしに行く。"""
    return jsonify({"cues": sound.playable_keys()})


@bp.get("/api/sound/preview")
def preview():
    """試聴。音声フォルダの中の音声ファイルだけを名前で返す(フォルダの区切りを含む名前は断る)。"""
    name = request.args.get("name", "")
    if not sound.is_audio_name(name):
        return _err("bad_value", "音声フォルダの中のファイル名だけ試聴できます", 400)
    p = sound.file_path(name)
    if p is None:
        return _err("not_found", f"音声フォルダにありません: {name}", 404)
    return send_file(p, mimetype=sound.mimetype(p), conditional=True, max_age=0)


@bp.get("/api/sound/<key>")
def play(key: str):
    """出来事の音声ファイル。鳴らさないと決めてある・ファイルが無いときは 404(画面は文言だけになる)。"""
    if sound.spec(key) is None:
        return _err("not_found", "その音はありません", 404)
    if not sound.enabled(key):
        return _err("off", "この出来事は鳴らさない設定です", 404)
    p = sound.sound_path(key)
    if p is None:
        return _err("not_found", f"音声フォルダにありません: {sound.chosen(key)}", 404)
    return send_file(p, mimetype=sound.mimetype(p), conditional=True, max_age=0)


@bp.post("/api/settings/sound")
def save():
    """出来事ごとの選択。body: {key, file}(file は「鳴らさない」か音声ファイルの名前)または {key, reset: true}。

    管理者パスワードは要らない(間違えても音が鳴らないだけ)。
    """
    body = request.get_json(silent=True) or {}
    key = str(body.get("key", ""))
    s = sound.spec(key)
    if s is None:
        return _err("bad_value", "その音はありません")
    try:
        if body.get("reset"):
            sound.reset(key)
        else:
            sound.set_choice(key, str(body.get("file", "")))
    except ValueError as e:
        return _err("bad_value", str(e))
    now = sound.chosen(key) or sound.SOUND_OFF
    log.info("音 %s(%s)を %s にしました", s.label, s.place, now)
    trace.event("settings.sound", f"音 {s.place}・{s.label} を {now} にしました")
    return jsonify(_state())
