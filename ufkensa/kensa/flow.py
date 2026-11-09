"""業務ロジックから画面へ伝えるものの型(MsgBox・MsgError・End の置き換え)

VBA は処理の途中で MsgBox を出し、Yes/No で分岐し、`MsgError … "END"` や `End` で止めていた。
Web では処理の途中で画面に聞けないので、業務ロジックは次の形で返す:

- **知らせる**(vbOKOnly の MsgBox・音声・ログ) → `Message` を `messages` に積む。文言は VBA の原文のまま
- **聞く**(vbYesNo / vbOKCancel)               → `Question` を返して止まる。画面が答えを持って呼び直す
                                                  (関数は `answers: dict[str, bool]` を受け取り、
                                                  `Question.key` の答えがあればそれで進む)
- **止める**(MsgError "END"・End・Exit Sub で中断) → `Stop` 例外。D10 で「止めるが状態は保つ」に決まったので、
                                                  画面は状態をそのままにしてメッセージを出す

    def print_check(state, answers):
        msgs = []
        if need_advance_sheet(state):
            ans = answers.get("advance_sheet")
            if ans is None:
                return Question("advance_sheet", "先行データ票はありますか？", buttons="yesno")
            if not ans:
                raise Stop("先行データ票はありますか？ → いいえ", messages=msgs)
        ...
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional

INFO, WARN, ERROR = "info", "warn", "error"


@dataclass
class Message:
    text: str                      # VBA の原文
    level: str = INFO              # info / warn / error(VBA の vbInformation / vbExclamation / vbCritical)
    title: str = ""                # MsgBox のタイトル
    sound: str = ""                # 鳴らす音の出来事の鍵(`sound.SOUNDS`。ファイル名は設定で決まる D29)
    source: str = ""               # VBA の場所(例 "Judg:471")

    def to_dict(self) -> dict[str, Any]:
        d = {"text": self.text, "level": self.level, "title": self.title, "sound": "", "source": self.source}
        if self.sound:
            d["cue"] = self.sound          # 2026-11-09: 出来事の鍵(音を鳴らさない設定でも、判定の画像はこれで出す)
        if self.sound:
            from . import sound as _sound                 # 鳴らさないと決めた出来事は画面へ渡さない(D29)
            if _sound.enabled(self.sound):
                d["sound"] = self.sound
                d["sound_text"] = _sound.message_text(self.sound)
        return d


@dataclass
class Question:
    key: str                       # 答えを返すときの鍵
    text: str                      # VBA の原文
    buttons: str = "yesno"         # yesno / okcancel
    title: str = ""
    source: str = ""
    messages: list[Message] = field(default_factory=list)   # 聞く前に出ていた知らせ

    def to_dict(self) -> dict[str, Any]:
        return {"question": {"key": self.key, "text": self.text, "buttons": self.buttons,
                             "title": self.title, "source": self.source},
                "messages": [m.to_dict() for m in self.messages]}


class Stop(Exception):
    """処理を止める(VBA の End / MsgError "END" / 途中の Exit Sub で中断)。状態は呼び手が保つ。"""

    def __init__(self, text: str, *, messages: Optional[list[Message]] = None, source: str = "",
                 level: str = ERROR) -> None:
        super().__init__(text)
        self.text = text
        self.source = source
        self.level = level
        self.messages = list(messages or [])

    def to_dict(self) -> dict[str, Any]:
        return {"stop": {"text": self.text, "source": self.source, "level": self.level},
                "messages": [m.to_dict() for m in self.messages]}


def unexpected_stop(e: BaseException, where: str = "", messages: Optional[list] = None) -> "Stop":
    """想定外のエラー(Python の誤りなど)を、画面を止めない「止まりました」の知らせにする(2026-11-06 の見直し)。

    前は 500 で返り、画面に「思わぬエラー」が出て状態が半端なままだった。ログには詳細(トレースバック)を残す。
    検査・キャスク・全ヵ所・Aインプット・校正の操作の入口で共通に使う。
    """
    import logging
    logging.getLogger("ufkensa.kensa.flow").error("操作 %s で想定外のエラー", where or "?", exc_info=e)
    return Stop(f"処理の途中でエラーになりました({type(e).__name__}: {e})。"
                "この操作はやり直してください。続くときはログを管理者に渡してください",
                messages=list(messages or []), source=where)
