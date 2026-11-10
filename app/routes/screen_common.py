"""画面の受け口で共通のもの(検査・キャスク・全ヵ所・Aインプット・校正)

整理 6(2026-11-10): 前は帳票を開く・包装仕様書の画像・操作の受け口を画面ごとに写していて、少しずつ違っていた
(全ヵ所の帳票だけ名前の確かめが緩く max_age が無い、検査の画像だけ型が JPEG 決め打ち、400 の返し方が 2 通り)。
ここに 1 つだけ置き、各画面の routes はこれを呼ぶ。URL と返す形は前と同じ。
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Callable, Mapping, Optional, Sequence

from flask import jsonify, request, send_file

from ufkensa import trace
from ufkensa.kensa.flow import Stop

#: 包装仕様書の画像の種類(拡張子 → MIME)。VBA は LoadPicture で .JPG を読む
IMAGE_TYPES = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".gif": "image/gif",
               ".bmp": "image/bmp"}

#: 帳票の名前(print フォルダの直下の .html だけ)
REPORT_NAME = re.compile(r"[^\\/]+\.html")

MSG_NOT_READY = "この操作はまだ結合中です(工程 8c)。いまは VBA の UF検査判定.xlsm を使ってください。"


def error(code: str, message: str, status: int):
    return jsonify({"error": {"code": code, "message": message}}), status


def report_file(folder: Path, name: str):
    """印刷した帳票(D8: ブラウザで開いて印刷する HTML)。手元のデータフォルダの print にあるものだけ。"""
    safe = Path(name).name
    if not REPORT_NAME.fullmatch(safe):
        return error("bad_value", "帳票の名前が正しくありません", 400)
    f = folder / safe
    if not f.is_file():
        return error("not_found", f"帳票がありません: {safe}", 404)
    return send_file(f, mimetype="text/html", max_age=0)


def pack_image(f: Optional[Path]):
    """いま出している包装仕様書の画像(UF_PackNo.PacIM)。使用パス(QualityPath\\包装仕様書)のファイルだけを返す。"""
    if f is None or not f.is_file():
        return error("not_found", "包装仕様書の画像がありません", 404)
    return send_file(f, mimetype=IMAGE_TYPES.get(f.suffix.lower(), "application/octet-stream"), max_age=0)


def operate(op: str, ops: Mapping[str, Sequence[Any]], get_session: Callable[[], Any], *, area: str, label: str,
            required: Optional[Callable[[str], bool]] = None, as_text: bool = False,
            stop_view: Optional[Callable[[Any], Any]] = None,
            trace_fields: Optional[Callable[[Any], Mapping[str, Any]]] = None):
    """画面の操作 1 つ = セッションのメソッド 1 つ。`ops[op]` = (メソッド, 受け取る引数[, ログの説明])。

    - 無い操作は 404
    - `required(name)` が真の引数が本文に無ければ 400(bad_value。足りない名前を並べる)
    - `as_text`: 引数を文字にして渡す(校正)
    - セッションが NotImplementedError を出したら 501(not_ready)
    - セッションが Result に包まずに Stop を投げたら、止めとして返す(view は `stop_view(session)`)
    """
    if op not in ops:
        return error("not_found", "その操作はありません", 404)
    entry = tuple(ops[op])
    method, names = entry[0], entry[1]
    desc = entry[2] if len(entry) > 2 else op
    body = request.get_json(silent=True) or {}
    kwargs = {n: (str(body[n]) if as_text else body[n]) for n in names if n in body}
    s = get_session()
    if required is not None:
        missing = [n for n in names if required(n) and n not in kwargs]
        if missing:
            return error("bad_value", "足りない値があります: " + ", ".join(missing), 400)
    fields = dict(trace_fields(s)) if trace_fields else {}
    try:
        with trace.timed(f"{area}.{op}", f"{label}: {desc}", expected=(Stop, NotImplementedError), **fields):
            result = getattr(s, method)(**kwargs)
    except NotImplementedError:
        return error("not_ready", MSG_NOT_READY, 501)
    except Stop as e:                                   # 念のため: セッションが Result に包まずに投げたとき
        return jsonify({"view": stop_view(s) if stop_view else None,
                        "messages": [m.to_dict() for m in e.messages],
                        "question": None, "stop": e.to_dict()["stop"]})
    return jsonify(result.to_dict())
