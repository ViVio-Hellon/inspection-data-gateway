"""検査の流れ(VBA の UF1 のイベント)をまとめる入口 ── 画面(app/routes/kensa.py)はここだけを呼ぶ

1 台の PC は 1 ラインを受け持ち、検査中のロットは 1 つ(VBA と同じ)。状態(`grid.KensaState`)は
プロセスに 1 つ持ち、操作のたびに手元 DB(`store.py`)へ書く(VBA の作業中保存 Posting に当たる)。
アプリを開き直したら手元 DB から読み戻す(VBA の Load_Text)。

**画面への返し方**(すべての操作で同じ形。`flow.py`):
    Result.view        いまの画面の値(`view.build(state)` の形)
    Result.messages    知らせ(MsgBox・音声)。文言は VBA の原文
    Result.question    聞くこと(vbYesNo)。画面は答えを `answers` に入れて同じ操作を呼び直す
    Result.stop        止めた理由(D10: 状態は保ったまま)

各操作と、対応する VBA のイベント:
    start()                       UF1.UserForm_Initialize(ライン・作業中保存の読み戻し・直)
    register_worker(card)         WorkButton → Worker_Change(社員カード13桁 → 名簿 → T1〜T3)
    register_lot(lot)             CommandButton10 → Lot_Change → CB_Change(重複確認・AccReflec・包装仕様・歪公差)
    select_work(work)             厚板/薄板/少枚数/極厚/スラブ_Click → LSelect
    set_yama(yama)                Yama_Change
    toggle_at() / toggle_inspe()  AtButt_Click / InspeButt_Click
    receive(text)                 StartTB → INP_Change(伝送文字列 DT…/#L…/#〆…)
    click(name) / dblclick(name)  ラベルのクリック(入力先を固定)/ダブルクリック(消去と理由)
    set_value(name, value)        歪・ラテラル・直角度・BH・検入高さ・外観・不良枚数などの手入力欄
    min_max()                     CommandButton4
    print_check(answers)          PrintB_Click の前半(先行データ票・必須チェック・検査員照合・梱包枚数)
    print_commit(answers)         PrintB_Click の後半(保存用への追記・測定DBへの登録・履歴・CSV・PDF)
    new_inspection(answers)       NewKen_Click
"""
from __future__ import annotations

import threading
from dataclasses import dataclass, field
from typing import Any, Optional

from .flow import Message, Question, Stop
from .grid import KensaState


@dataclass
class Result:
    view: dict[str, Any]
    messages: list[Message] = field(default_factory=list)
    question: Optional[Question] = None
    stop: Optional[Stop] = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "view": self.view,
            "messages": [m.to_dict() for m in self.messages],
            "question": self.question.to_dict()["question"] if self.question else None,
            "stop": self.stop.to_dict()["stop"] if self.stop else None,
        }


class Session:
    """プロセスに 1 つの検査セッション。操作は `_lock` で 1 つずつ通す。"""

    def __init__(self) -> None:
        self._lock = threading.RLock()
        self.state = KensaState()

    # 以下の操作は工程 8c(結合)で実装する。画面はこの名前と引数で呼ぶ
    def start(self) -> Result:
        raise NotImplementedError

    def register_worker(self, card: str, answers: Optional[dict[str, bool]] = None) -> Result:
        raise NotImplementedError

    def register_lot(self, lot: str, answers: Optional[dict[str, bool]] = None) -> Result:
        raise NotImplementedError

    def select_work(self, work: str) -> Result:
        raise NotImplementedError

    def set_yama(self, yama: str) -> Result:
        raise NotImplementedError

    def toggle_at(self) -> Result:
        raise NotImplementedError

    def toggle_inspe(self) -> Result:
        raise NotImplementedError

    def receive(self, text: str) -> Result:
        raise NotImplementedError

    def click(self, name: str) -> Result:
        raise NotImplementedError

    def dblclick(self, name: str, answers: Optional[dict[str, bool]] = None, reason: str = "") -> Result:
        raise NotImplementedError

    def set_value(self, name: str, value: str) -> Result:
        raise NotImplementedError

    def min_max(self) -> Result:
        raise NotImplementedError

    def print_check(self, answers: Optional[dict[str, bool]] = None) -> Result:
        raise NotImplementedError

    def print_commit(self, answers: Optional[dict[str, bool]] = None) -> Result:
        raise NotImplementedError

    def new_inspection(self, answers: Optional[dict[str, bool]] = None) -> Result:
        raise NotImplementedError


_session: Optional[Session] = None
_slock = threading.Lock()


def get() -> Session:
    global _session
    with _slock:
        if _session is None:
            _session = Session()
        return _session


def reset_for_test(s: Optional[Session] = None) -> Session:
    global _session
    with _slock:
        _session = s or Session()
        return _session
