"""ブラウザの通しの試験用のサーバ(2026-11-10 リファクタリング 第1段 5)

    python -m tests.e2e.server <port>

試験用の台帳(tests/kensa_fixture)を一時フォルダに写し、オフラインの使用パスにして、校正済みにしてから
アプリを 127.0.0.1:<port> で動かす。トークンは `e2e-token`。準備ができたら標準出力に `READY <port> <pty>` を出す。
標準入力が閉じたら(呼び手の試験が終わったら)止まり、一時フォルダを消して疑似端末を閉じる。

受信機の代わりに疑似端末(pty。Linux・macOS)を 1 つ作る。`POST /__e2e/feed {"text": "..."}` で
その端末へ 1 行書く(受信機から測定値が届いたのと同じ)。pty が作れない OS では `<pty>` は `-`。
"""
from __future__ import annotations

import os
import shutil
import signal
import sqlite3
import sys
import tempfile
import threading
from datetime import datetime
from pathlib import Path
from unittest import mock

import tests  # noqa: F401  ログ・設定・ローカル領域を一時フォルダへ(ufkensa より先に)
from tests import _isolation, kensa_fixture

TOKEN = "e2e-token"


def _prepare() -> Path:
    from tests.test_kensa_session import PACK_COLS
    from ufkensa.kensa import paths, records, store, terminal, workers
    tmp = Path(tempfile.mkdtemp(prefix="ufkensa_e2e_"))
    db = kensa_fixture.copy_to(tmp / "db")
    kensa_fixture.use_offline(db)
    con = sqlite3.connect(db / "梱包資材マスタ.sqlite3")
    con.execute("CREATE TABLE 注意_包装仕様 (" + ",".join(f'"{c}"' for c in PACK_COLS) + ")")
    con.execute("INSERT INTO 注意_包装仕様 (包装仕様NO, 形状, コメント) VALUES ('ZZ', '板', '別の仕様')")
    con.commit()
    con.close()
    (db / "包装仕様書").mkdir(exist_ok=True)
    (db / "包装仕様書" / "0P0000.JPG").write_bytes(b"\xff\xd8\xff")
    (db / "【測定DB】梱包").mkdir(exist_ok=True)
    records.ensure_measure_db(db / "【測定DB】梱包" / "梱包検査ﾃﾞｰﾀ.sqlite3")
    paths.set_path("local_history", str(tmp / "検査履歴"))
    terminal.set_facility("HVC")
    mock.patch.object(paths, "app_dir", return_value=str(tmp) + "/").start()
    now = datetime.now()
    shifts = workers.load_shift_times()
    st = store.LocalStore()
    sh = st.ensure_sheet("校正記録")
    sh.write(2, 1, workers.yea_dayr_check(workers.hhmm(now), now=now, facility="HVC", shifts=shifts), "S")
    sh.write(2, 4, workers.time_check(workers.hhmm(now), facility="HVC", shifts=shifts), "S")
    st.close()
    return tmp


def _open_pty() -> tuple[int, int, str]:
    """(親の fd, 子の fd, 子の名前)。作れなければ (-1, -1, "-")。"""
    try:
        import pty
        master, slave = pty.openpty()
        return master, slave, os.ttyname(slave)
    except (ImportError, OSError):
        return -1, -1, "-"


def main(port: int) -> None:
    import logging
    logging.getLogger("werkzeug").setLevel(logging.ERROR)
    _isolation.ensure_isolated()
    if hasattr(signal, "SIGHUP"):
        # 呼び手は別のセッションで起こす(まとめて止めるため)。そのとき疑似端末がこのセッションの端末になり、
        # 閉じると SIGHUP で落ちて後始末(一時フォルダを消す)が飛ぶ。無視する
        signal.signal(signal.SIGHUP, signal.SIG_IGN)
    tmp = _prepare()
    master, slave, pty_name = _open_pty()
    from flask import jsonify, request
    from werkzeug.serving import make_server
    from app import create_app
    app = create_app("main", token=TOKEN, port=port)
    app.config["READY"] = True

    def feed():
        text = str((request.get_json(silent=True) or {}).get("text", ""))
        if master < 0:
            return jsonify({"ok": False, "error": "pty なし"}), 400
        os.write(master, text.encode("ascii") + b"\r")
        return jsonify({"ok": True})
    app.add_url_rule("/__e2e/feed", "e2e_feed", feed, methods=["POST"])
    srv = make_server("127.0.0.1", port, app, threaded=True)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    print(f"READY {port} {pty_name}", flush=True)
    try:
        sys.stdin.read()                   # 呼び手が閉じるまで
    finally:
        srv.shutdown()
        srv.server_close()
        shutil.rmtree(tmp, ignore_errors=True)
        for fd in (master, slave):
            if fd >= 0:
                try:
                    os.close(fd)
                except OSError:
                    pass


if __name__ == "__main__":
    main(int(sys.argv[1]))
