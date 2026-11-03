"""誰が使っているか と、編集可 / 閲覧のみ(0.16.0)

【0.16.0 から】VBA 版から続いていた「ID の表 + パスワード + Level 1〜5」をやめた。
Windows の**ログインID**と**PC名**で、設定ファイルの `アクセス権限` 表(0.22.0)を引き、
**編集可**(権限 = 編集可 の行がある)か **閲覧のみ** かを決める(`access_rights`)。パスワードは要らない。
画面を開けば自動で入る(ログインの画面は無い)。

中の決まり(どの操作に何が要るか)は Level の数のまま持つ: 編集可 = 5、閲覧のみ = 1。
閲覧のみは 新規・修正・削除・列の追加・設定の保存 ができない(ダブルクリックは参照で開く)。

【セッション】プロセスに 1 つ(1 台の PC を 1 人が使う前提)。権限は `access_rights.RECHECK_SEC` ごとに
読み直す(取り消しが効くように。読み直しは裏で、画面は待たせない)。設定「無操作で権限を読み直す(分)」が
0 より大きければ、その時間触っていないと次の操作で読み直す。
"""
from __future__ import annotations

import os
import socket
import threading
import time
from dataclasses import dataclass
from typing import Optional

from . import access_log, access_rights, config, user_settings
from .logging_utils import get_logger

log = get_logger("auth")

# db-viewer の VBA にあった開発者の特例。UF検査判定の VBA には無いので空(権限はアクセス権限の表だけで決める)
DEVELOPER_IDS: tuple[str, ...] = ()
DEVELOPER_PCS: tuple[str, ...] = ()

EDIT_LEVEL = 5
VIEW_LEVEL = 1
MSG_NO_LEVEL = ("管理者の権限がありません。アクセス権限マスタ(梱包資材マスタの「アクセス権限」)に、"
                "この ログインID・PC名 で 権限 = Administrator の行を足してもらってください。")


def network_id() -> str:
    """VBA `NetWorkID`(WScript.Network.UserName)。"""
    return os.environ.get("USERNAME") or os.environ.get("USER") or ""


def pc_name() -> str:
    """VBA `UsePCName`(WScript.Network.ComputerName)。"""
    return os.environ.get("COMPUTERNAME") or socket.gethostname()


def is_developer(use_id: str, use_pc: str) -> bool:
    return use_id in DEVELOPER_IDS or use_pc in DEVELOPER_PCS


@dataclass
class LoginResult:
    ok: bool
    message: str = ""
    note: str = ""
    user_id: str = ""
    level: int = 0
    developer: bool = False
    rights: Optional[access_rights.Decision] = None


def rights_result() -> LoginResult:
    """この PC のログイン ID・PC 名で、設定ファイルのアクセス権限を引く(パスワードは無い)。"""
    use_id, use_pc = network_id(), pc_name()
    d = access_rights.evaluate(use_id, use_pc)
    dev = is_developer(use_id, use_pc)
    edit = d.edit or dev
    access_log.write("認証成功")
    user_settings.save(config.KEY_LAST_USER, use_id)
    return LoginResult(True, note=d.reason + ("(開発者)" if dev and not d.edit else ""), user_id=use_id,
                       level=EDIT_LEVEL if edit else VIEW_LEVEL, developer=dev, rights=d)


# ------------------------------------------------------------------
# セッション(プロセスに 1 つ)
# ------------------------------------------------------------------
@dataclass
class Session:
    user_id: str
    level: int
    developer: bool
    logged_in_at: float
    last_seen: float
    rights: Optional[access_rights.Decision] = None
    checked_at: float = 0.0

    @property
    def can_edit(self) -> bool:
        return self.level >= EDIT_LEVEL

    @property
    def label(self) -> str:
        return access_rights.LABEL_EDIT if self.can_edit else access_rights.LABEL_VIEW

    def touch(self) -> None:
        self.last_seen = time.time()

    def expired(self) -> bool:
        minutes = int(user_settings.get(config.KEY_RELOGIN_MIN, 0) or 0)
        if minutes <= 0:
            return False
        return time.time() - self.last_seen > minutes * 60


_lock = threading.Lock()
_session: Optional[Session] = None


def current() -> Optional[Session]:
    """いまのログイン。無操作で切れていれば消して `None`。"""
    global _session
    with _lock:
        if _session is not None and _session.expired():
            log.info("無操作だったので権限を読み直します: %s", _session.user_id)
            _session = None
        return _session


def start(result: LoginResult) -> Session:
    global _session
    now = time.time()
    with _lock:
        _session = Session(result.user_id, result.level, result.developer, now, now, result.rights, now)
        return _session


_applied_note = ""


def sign_in() -> Session:
    """アクセス権限マスタを引いて、セッションを始める(画面を開けば自動で・「権限を読み直す」)。

    マスタでこの PC のラインが決まれば、この端末の設備名に当てる(前に当てたものから変わったときだけ。
    `access_rights.apply_line`)。裏の読み直し(`ensure`)では当てない ── 検査の途中でラインを変えないため。
    """
    global _applied_note
    result = rights_result()
    s = start(result)
    if result.rights is not None:
        try:
            note = access_rights.apply_line(result.rights)
        except Exception:                                # noqa: BLE001 - 当てられなくても使い始める
            log.warning("アクセス権限のラインを当てられませんでした", exc_info=True)
            note = ""
        if note:
            _applied_note = note
    log.info("権限: %s / %s → %s(%s)", s.user_id, pc_name(), s.label, result.rights.source if result.rights else "")
    return s


def applied_note() -> str:
    """最後にアクセス権限マスタからラインを当てたときの知らせ(設定の画面に出す)。"""
    return _applied_note


_refreshing = threading.Event()


def ensure() -> Session:
    """いまのセッション。無ければ始める。古ければ裏で読み直す(取り消しが効くように)。"""
    s = current()
    if s is None:
        return sign_in()
    if time.time() - s.checked_at > access_rights.RECHECK_SEC and not _refreshing.is_set():
        _refreshing.set()
        s.checked_at = time.time()

        def again() -> None:
            try:
                result = rights_result()
                with _lock:
                    if _session is s:
                        if s.level != result.level:
                            log.info("権限が変わりました: %s → %s", s.label,
                                     access_rights.LABEL_EDIT if result.level >= EDIT_LEVEL else access_rights.LABEL_VIEW)
                        s.level, s.rights, s.checked_at = result.level, result.rights, time.time()
            except Exception:                            # noqa: BLE001 - 読み直せなければ今のまま
                log.warning("権限を読み直せませんでした", exc_info=True)
            finally:
                _refreshing.clear()

        threading.Thread(target=again, daemon=True, name="rights-recheck").start()
    return s


def end() -> None:
    global _session
    with _lock:
        _session = None


def level() -> int:
    s = current()
    return s.level if s else 0


def setup_mode() -> bool:
    """設定の途中か = アクセス権限マスタから決まらない(未設定・読めない・管理者の行が無い)。

    このあいだは**管理者パスワードだけで**マスタ編集(使用パスなど)を開けられる。権限の表の場所が使用パスに
    あるので、そうしないと新しい PC・試験用の PC で「パスを直すには権限、権限を読むにはパス」と詰む。
    """
    s = current()
    return bool(s is not None and s.rights is not None and s.rights.setup)


def can_edit() -> bool:
    """管理者か(マスタの Administrator。設定の途中は 管理者パスワードを通したら 管理者として扱う)。"""
    return level() >= EDIT_LEVEL or (_master_unlocked and setup_mode())


def has(code: str) -> bool:
    """この人・この端末がその権限(`access_rights.CODES`)を持つか。"""
    if code in (access_rights.ADMIN, access_rights.ADMIN_ALIAS):
        return can_edit()
    s = current()
    if s is None or s.rights is None:
        return False
    return s.rights.has(code)


# ------------------------------------------------------------------
# レベル判定(VBA `ボタンクリック制御`。編集可 = 5 / 閲覧のみ = 1)
# ------------------------------------------------------------------
def required_level(mode: str) -> int:
    if mode in ("新規", "削除"):
        return 3
    if mode == "修正":
        return 2
    return 0


def check_mode(mode: str, lv: int, click_type: str = "ButtonClick") -> tuple[str, str]:
    """(通したモード, 断りの文言)。ダブルクリックで修正に足りなければ参照に落とす。"""
    need = required_level(mode)
    if lv >= need:
        return mode, ""
    if mode == "修正" and click_type == "DblClick":
        return "参照", ""
    return "", MSG_NO_LEVEL


# ------------------------------------------------------------------
# マスタ編集の認証(python-web-tools と同じ考え方)。2026-11-02 の決め:
# マスタ管理・保守の画面は誰でも開けるが、編集は「管理者パスワード」と「アクセス権限の管理者」の両方が要る。
# パスワードを入れるとこのアプリを閉じるまで(または lock_master まで)編集できる。
# 2026-11-03: アクセス権限マスタから決まらないあいだ(setup_mode)は、管理者パスワードだけで開ける(D30)。
# ------------------------------------------------------------------
_master_unlocked = False


def unlock_master(password: str) -> tuple[bool, str]:
    """管理者パスワードで「マスタ編集」を開ける。(できたか, 断った理由)。"""
    global _master_unlocked
    from . import admin_password
    ensure()
    if not can_edit() and not setup_mode():
        return False, "この端末・この人にアクセス権限の「管理者」がありません(設定 → 権限 で確かめてください)"
    if not admin_password.verify(str(password or "")):
        return False, "管理者パスワードが違います"
    _master_unlocked = True
    log.info("マスタ編集の認証を通りました: %s%s", network_id(), "(設定の途中: パスワードだけ)" if setup_mode() else "")
    return True, ""


def lock_master() -> None:
    global _master_unlocked
    _master_unlocked = False


def master_unlocked() -> bool:
    return _master_unlocked


def master_edit_ok() -> bool:
    """マスタ・保守の編集をしてよいか(パスワードを通した かつ 管理者の権限がある。設定の途中はパスワードだけ)。"""
    return _master_unlocked and (can_edit() or setup_mode())


def master_edit_reason() -> str:
    """編集できない理由(できるなら "")。画面にそのまま出す。"""
    if setup_mode():
        if not _master_unlocked:
            s = current()
            why = s.rights.setup_reason if s and s.rights else access_rights.MSG_SETUP
            return f"{why}。編集するには、管理者パスワードを入れてください"
        return ""
    if not can_edit():
        return "編集には、アクセス権限の「管理者」(Administrator)が要ります"
    if not _master_unlocked:
        return "編集するには、管理者パスワードを入れてください"
    return ""
