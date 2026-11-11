"""開いている画面(タブ)の名簿と「いま使っている画面」(2026-11-11)。

現場の声「まだ多重起動できる」── サーバは1つになったが、起動をもう一度押すと
**ブラウザの画面がもう1つ開き、どちらでも操作できた。** そこで:

- 画面は数秒ごとに `GET /api/screen` を送る(心拍)。ここで最後に見た時刻を覚える
- 2回目の起動は、見えている画面があれば新しく開かずに `POST /api/screen/raise` で
  「前に出て」と頼む(画面は次の心拍でそれを受け取り、前に出す・題を点滅・知らせを出す)
- 書く操作を受け付けるのは**いま使っている画面**(最後に「こちらで使う」を押した画面)だけ。
  ほかの画面からの書く操作は断る(`inactive_screen`)

画面番号の無い要求(試験・curl・古い画面)は今までどおり通す。
"""
from __future__ import annotations

import threading
import time
from dataclasses import dataclass
from typing import Optional

#: 見えている画面は `POLL_MS` ごとに心拍を送る。これより長く来なければ「閉じた」とみなす
POLL_MS = 3000
OPEN_SEC = 10.0
#: 裏に回った画面はブラウザが心拍を間引く(Chrome は1分に1回まで)。そのぶん長く待つ
HIDDEN_OPEN_SEC = 150.0
#: 閉じた画面が「使っている画面」を持ち続ける長さ。同じタブの読み込み直しが引き継ぐまでの間、
#: ほかのタブに取られないように
HANDOFF_SEC = 5.0

INACTIVE_CODE = "inactive_screen"
INACTIVE_MESSAGE = ("この画面はほかのウィンドウで使われています。"
                    "この画面で続けるには「こちらで使う」を押してください")


@dataclass
class _Seen:
    at: float
    visible: bool = True
    gone: bool = False          # 閉じた(読み込み直しなら、すぐ次の番号が引き継ぐ)


class ScreenRegistry:
    def __init__(self, clock=time.monotonic) -> None:
        self._clock = clock
        self._lock = threading.Lock()
        self._seen: dict[str, _Seen] = {}
        self._active = ""
        self._raise_seq = 0

    # ---- 内部 ----------------------------------------------------------
    def _alive(self, sid: str, now: float) -> bool:
        s = self._seen.get(sid)
        if s is None:
            return False
        if s.gone:
            return now - s.at <= HANDOFF_SEC
        return now - s.at <= (OPEN_SEC if s.visible else HIDDEN_OPEN_SEC)

    def _settle(self, now: float) -> None:
        """閉じた画面を名簿から外す。いま使っている画面が閉じていれば空ける。"""
        for k in [k for k in self._seen if not self._alive(k, now)]:
            del self._seen[k]
        if self._active and self._active not in self._seen:
            self._active = ""

    # ---- 画面から ------------------------------------------------------
    def ping(self, sid: str, *, visible: bool = True, prev: str = "") -> dict:
        """心拍。`prev` は同じタブで読み込み直す前の番号(引き継ぐ)。"""
        if not sid:
            return {"active": True, "raise_seq": self._raise_seq}
        with self._lock:
            now = self._clock()
            if prev and prev != sid and self._active == prev:
                self._active = sid
                self._seen.pop(prev, None)
            self._seen[sid] = _Seen(now, visible)
            self._settle(now)
            if not self._active:
                self._active = sid
            return {"active": self._active == sid, "raise_seq": self._raise_seq}

    def touch(self, sid: str, prev: str = "") -> None:
        """心拍以外の要求も「開いている」の合図にする(見える・見えないは変えない)。
        `prev` は同じタブで読み込み直す前の番号。使っている画面だったなら引き継ぐ。"""
        if not sid:
            return
        with self._lock:
            if prev and prev != sid and self._active == prev:
                self._active = sid
                self._seen.pop(prev, None)
            s = self._seen.get(sid)
            if s is None:
                self._seen[sid] = _Seen(self._clock())
            else:
                s.at, s.gone = self._clock(), False

    def hidden(self, sid: str) -> None:
        if not sid:
            return
        with self._lock:
            # 閉じるときは「裏に回った」と「閉じた」が続けて届き、順番は決まらない。
            # 閉じた画面・知らない画面を「裏に回った」で生き返らせない
            s = self._seen.get(sid)
            if s is not None and not s.gone:
                s.at, s.visible = self._clock(), False

    def claim(self, sid: str) -> None:
        """「こちらで使う」。この画面をいま使っている画面にする。"""
        if not sid:
            return
        with self._lock:
            self._seen[sid] = _Seen(self._clock())
            self._active = sid

    def leaving(self, sid: str) -> None:
        """閉じた(`pagehide`)。使っている画面なら、少し待ってから空ける(読み込み直しが引き継ぐ)。"""
        if not sid:
            return
        with self._lock:
            if self._active == sid:
                self._seen[sid] = _Seen(self._clock(), True, gone=True)
            else:
                self._seen.pop(sid, None)

    def may_write(self, sid: str) -> bool:
        """この画面からの書く操作を受け付けるか。番号が無ければ通す。"""
        if not sid:
            return True
        with self._lock:
            now = self._clock()
            self._settle(now)
            if not self._active:
                self._active = sid
                self._seen.setdefault(sid, _Seen(now))
            return self._active == sid

    def is_active(self, sid: str) -> bool:
        if not sid:
            return True
        with self._lock:
            self._settle(self._clock())
            return not self._active or self._active == sid

    # ---- 起動から ------------------------------------------------------
    def open_screen(self) -> Optional[str]:
        """見えている(最近心拍の来た)画面。いま使っている画面を優先。無ければ None。"""
        with self._lock:
            now = self._clock()
            self._settle(now)
            visible = [k for k, s in self._seen.items()
                       if s.visible and not s.gone and now - s.at <= OPEN_SEC]
            if not visible:
                return None
            return self._active if self._active in visible else visible[0]

    def request_raise(self) -> bool:
        """見えている画面があれば「前に出て」を出して True。無ければ False(ブラウザを開く)。"""
        if self.open_screen() is None:
            return False
        with self._lock:
            self._raise_seq += 1
        return True

    @property
    def active(self) -> str:
        return self._active


def of(app) -> ScreenRegistry:
    reg = app.extensions.get("ufkensa.screens")
    if reg is None:
        reg = app.extensions["ufkensa.screens"] = ScreenRegistry()
    return reg
