"""管理者以外停止モード ── UF検査判定の受け方(`remote_stop.py` に渡す、このツールの事情)

`remote_stop.py`(アプリのフォルダ直下)は all-tools から**中身を変えずに**写したもの。違うところはここで渡す:

    名前                 UF検査判定(app_id = ufkensa)
    手元の領域           `app_config.local_root()`(%LOCALAPPDATA%\\UFKensa)。控え = data\\停止控え\\
    置き場所             設定 → 使用パス「管理者以外停止の置き場所」。空なら 検査履歴。その下の 管理者以外停止\\
    お願いを出せる人     管理者パスワードを通した(保守の鍵) **かつ** アクセス権限が Administrator
                         (設定の途中 = パスワードだけで鍵が開くあいだは出せない)

【残っているもの】(帯に出し、利用者を待つ。5分で強制)
検査の作業の状態は操作のたびに手元DB(SQLite)へ書いてあるので、入れたロット・測定値は止めても消えない。
消えるのは「操作の途中」と「画面にしかないもの」だけ:

    - 実行中の処理(取り込み・書き戻しなど `jobs` の busy)
    - 検査の操作の途中(検査セッションの錠を握っている = 1 つの操作が走っている)
    - 「印刷する？」の答え待ち・印刷・保存の確認が済んだまま(印刷・保存がまだ済んでいない)
    - 開いたままの確認の窓(公差外の理由選択 UFConfi・包装仕様書の確認)
    - 画面の打ちかけ(帯が `remote_stop_page.js` から言い添える。入力して Enter・確定していない欄)

強制で止める前に、何が残っていたか・どうすれば戻るかを手元の控え(`remote_stop.save_keep`)に書き、
次の起動で一度だけ知らせる。画面の打ちかけは、止めると決まったとき画面が値ごと送ってきて控えに書く(`save_drafts`)。

**業務(ufkensa/kensa)は変えない。** ここは覗くだけ(錠は取れるか試してすぐ返す)。
"""
from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any, Optional

from .logging_utils import get_logger

log = get_logger("remote_stop")

APP_NAME = "UF検査判定"
APP_ID = "ufkensa"


def _module():
    root = str(Path(__file__).resolve().parent.parent)
    if root not in sys.path:
        sys.path.insert(0, root)
    import remote_stop
    return remote_stop


def shared_root() -> Path:
    """お願いを置くフォルダの親(この下に 管理者以外停止\\)。決まらなければ `Unreadable`。"""
    rs = _module()
    env = os.environ.get(rs.ROOT_ENV, "").strip()
    if env:
        return Path(env)
    from .kensa import paths
    text = (paths.get("remote_stop") or "").strip() or (paths.get("history") or "").strip()
    if not text:
        raise rs.Unreadable("設定 → 使用パス の「管理者以外停止の置き場所」も「検査履歴」も空です")
    return paths.to_local(text.rstrip("\\/") or text)


def configured():
    """このツールの事情を渡した `remote_stop`。"""
    from . import app_config
    rs = _module()
    rs.configure(app=APP_NAME, app_id=APP_ID, version=app_config.version(), local_root=app_config.local_root(),
                 root=shared_root, bundle_root=None, hosted=False, logger=log)
    return rs


# ------------------------------------------------------------------
# 残っているもの・控え
# ------------------------------------------------------------------
def _kensa_session():
    """検査セッション(**作らない**。まだ無ければ None)。"""
    mod = sys.modules.get("ufkensa.kensa.session")
    return getattr(mod, "_session", None) if mod is not None else None


def _lot(s: Any) -> str:
    try:
        return str(s.state.lot or "")
    except Exception:                             # noqa: BLE001
        return ""


def left() -> list:
    """いま止めると失われるもの(呼び名)。"""
    items: list = []
    try:
        from . import jobs
        items += [f"実行中の処理({x})" for x in jobs.get_registry().busy_labels()]
    except Exception:                             # noqa: BLE001
        log.debug("実行中の処理を数えられませんでした", exc_info=True)
    s = _kensa_session()
    if s is None:
        return items
    lock = getattr(s, "_lock", None)
    if lock is not None:
        if lock.acquire(blocking=False):
            lock.release()
        else:
            items.append("検査の操作の途中")
    lot = _lot(s)
    on = f"(ロット {lot})" if lot else ""
    if getattr(s, "_print_pending", None) is not None:
        items.append(f"「印刷する？」の答え待ち{on}")
    elif getattr(s, "print_ready", False):
        items.append(f"印刷・保存がまだ済んでいません{on}")
    if getattr(s, "confi", None):
        items.append(f"公差外の確認の窓{on}")
    if getattr(s, "pack", None) is not None:
        items.append(f"包装仕様書の確認の窓{on}")
    return items


def keep() -> list:
    """強制で止める前に、手元に置いたもの・戻し方(控えに書く一言)。"""
    now = left()
    if not now:
        return []
    out = []
    s = _kensa_session()
    lot = _lot(s) if s is not None else ""
    if lot:
        out.append(f"検査中のロット {lot} の入力は手元DBに残っています。検査画面で同じロットを開くと続きから使えます。")
    if any("印刷" in x for x in now):
        out.append(f"印刷・保存は済んでいません。{('ロット ' + lot + ' を') if lot else ''}もう一度「印刷・保存」してください。")
    if any(x.startswith("実行中の処理") for x in now):
        out.append("途中で止めた処理(取り込みなど)は、もう一度行ってください。")
    return out


def save_drafts(page: str, drafts: Any) -> int:
    """画面が送ってきた打ちかけ(止めると決まったとき)を手元の控えに書く。書いた欄の数。"""
    rs = _module()
    w = rs.get()
    if w is None or w.request is None or not isinstance(drafts, list):
        return 0
    kept = []
    for d in drafts[:30]:
        if not isinstance(d, dict):
            continue
        label = str(d.get("label") or "").strip()[:80] or "(名前の無い欄)"
        value = str(d.get("value") or "").strip()[:500]
        if value:
            kept.append(f"入力しかけ「{label}」: {value}")
    if not kept:
        return 0
    where = f"(画面 {str(page or '')[:80]})" if page else ""
    left_now = [f"画面の打ちかけ{where}"]
    return len(kept) if rs.save_keep(left_now, kept, w.request) else 0


# ------------------------------------------------------------------
# 起動の入口から(start_app.py は1行ずつ呼ぶだけ)
# ------------------------------------------------------------------
def gate(srv) -> bool:
    """お願い中(更新中)なら理由を待機画面に出して True(しばらくして終わる)。判定で起動を止めない。"""
    try:
        return configured().gate(srv.mark_error, srv.stop)
    except Exception as exc:                      # noqa: BLE001
        log.warning("管理者以外停止のお願いを確かめられませんでした: %s", exc)
        return False


def watch(srv, *, bare: bool = False, kind: str = "browser") -> None:
    """お願いを見張る。止め方は `/api/shutdown` と同じ道(`srv.stop`)。`bare` = 本体が無い(理由の画面のまま)。"""
    try:
        rs = configured()
        rs.watch(srv.stop, (lambda: []) if bare else left, kind=kind,
                 before_force=None if bare else keep, restore=not bare)
    except Exception as exc:                      # noqa: BLE001 - 見張りが立たなくても起動は続ける
        log.warning("管理者以外停止の見張りを立てられませんでした: %s", exc)


def leave() -> None:
    mod = sys.modules.get("remote_stop")
    if mod is None:
        return
    try:
        mod.leave()
    except Exception:                             # noqa: BLE001 - 終わるときの後始末で止まらない
        pass


def stop_bat_refusal(running: bool) -> str:
    """stop.bat: お願いが出ていてこのツールが動いていれば、止めない理由。読めなければ断らない。"""
    if not running:
        return ""
    try:
        return configured().stop_bat_refusal() or ""
    except Exception:                             # noqa: BLE001
        return ""


# ------------------------------------------------------------------
# 管理者の画面(保守 → 管理者以外停止)
# ------------------------------------------------------------------
def who() -> dict:
    """お願いを出せるか(パスワード = 保守の鍵を開けた / Administrator = アクセス権限。設定の途中は数えない)。"""
    from . import auth
    password_ok = auth.master_unlocked()
    try:
        administrator = auth.level() >= auth.EDIT_LEVEL
    except Exception:                             # noqa: BLE001
        administrator = False
    return {"password_ok": password_ok, "administrator": administrator}


def overview() -> dict:
    rs = configured()
    out = rs.overview().to_dict()
    out.update(who())
    return out


def issue(reason: str = ""):
    w = who()
    return configured().issue(w["password_ok"], w["administrator"], reason)


def clear():
    w = who()
    return configured().clear(w["password_ok"], w["administrator"])


def note_for(result: Optional[Any]) -> dict:
    return {"ok": bool(result.ok), "message": result.message, "reason": result.reason}
