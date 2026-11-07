"""画面の外枠に渡す値

左のレールの並びと、テンプレートが必要とする共通の値をここで作る。
UF検査判定の画面の並び(docs/03_移植計画.md の機能対応表)。
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional

from flask import current_app, request

from ufkensa import app_config, idle_exit, modes, user_settings


@dataclass
class NavItem:
    key: str
    label: str
    url: str
    note: str = ""
    badge: str = ""
    badge_kind: str = "todo"
    ready: bool = True


# 画面の並び。上から、ふだん使う順(A案「計測コックピット」が中心。D6)
# 名簿・日報(/meibo)はメニューに出さない。検査で名簿の不一致などが出たときだけ、その知らせから開く
NAV: tuple[tuple[str, str, str, str], ...] = (
    ("kensa", "検査", "/kensa", "測る・判定・Min/Max・印刷"),
    ("cask", "キャスク", "/cask", "キャスク検査(最大700枚)"),
    ("zenkasho", "全ヵ所", "/zenkasho", "全ヵ所(一般/5N/三菱受託)"),
    ("ainput", "Aインプット", "/ainput", "梱包高さ・アイカス送信"),
    ("calibration", "校正", "/calibration", "直尺の校正"),
    ("quality", "品質基準", "/quality", "このロットの品質基準を見る"),
    ("history", "履歴", "/history", "保存した検査を見る・再印刷"),
    ("import", "取り込み", "/import", "Access・xlsm からの取り込み"),
    ("maint", "保守", "/maint", "歪公差チェック・端末設定のセット(管理者)"),
    ("settings", "設定", "/settings", "この端末・使用パス・権限"),
)

# もう作った画面。ここに挙がっていないものは「準備中」ページを出す
READY_SCREENS = frozenset({"settings", "import", "kensa", "history", "ainput", "calibration", "quality", "zenkasho", "cask", "meibo", "maint"})

# 準備中の画面の説明。key: (作る予定, 何ができるようになるか, いま使う画面)
PENDING_SCREENS: dict[str, tuple[str, str, str]] = {
    "kensa": ("工程 8c", "ロット登録から測定・判定・Min/Max・印刷まで、1画面で行えるようになります(A案)。", "VBA の UF検査判定.xlsm"),
    "cask": ("工程 8e", "キャスク検査の入力・判定・印刷・保存。", "VBA の UF検査判定.xlsm"),
    "zenkasho": ("工程 8e", "全ヵ所検査(一般・5N・三菱受託)。", "VBA の UF検査判定.xlsm"),
    "ainput": ("工程 8c", "梱包高さチェックとアイカスへの送信。", "VBA の UF検査判定.xlsm"),
    "calibration": ("工程 8c", "直尺の校正の記録と判定。", "VBA の UF検査判定.xlsm"),
    "quality": ("工程 8c", "ロットの品質基準(外観・歪・BH・ラテラル)の表示。", "VBA の UF検査判定.xlsm"),
    "history": ("工程 8c", "保存した検査の一覧・復元・再印刷。", "VBA の UF検査判定.xlsm"),
    "import": ("工程 8a", "Access(.accdb)を SQLite に取り込み、各ラインの xlsm から過去のデータを1回取り込みます。", "設定"),
}


def nav_items(mode: str = modes.DEFAULT,
              badges: Optional[dict[str, tuple[str, str]]] = None) -> list[NavItem]:
    badges = badges or {}
    items = []
    for key, label, url, note in NAV:
        ready = key in READY_SCREENS
        badge, kind = badges.get(key, ("", "todo"))
        if not ready and not badge:
            badge, kind = "準備中", "todo"
        items.append(NavItem(key=key, label=label, url=url, note=note,
                             badge=badge, badge_kind=kind, ready=ready))
    return items


def pending_items(mode: str = modes.DEFAULT) -> list[NavItem]:
    return [item for item in nav_items(mode) if not item.ready]


_FAVICON = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">'
            '<rect width="32" height="32" rx="6" fill="{color}"/>'
            '<text x="16" y="21" font-size="13" font-family="sans-serif" font-weight="bold"'
            ' text-anchor="middle" fill="#ffffff">{mark}</text></svg>')


def favicon(mode: str = modes.DEFAULT) -> str:
    item = modes.get(modes.normalize(mode))
    return _FAVICON.format(color=item.color, mark=item.mark)


FALLBACK_URL = "/settings"


def home_url(mode: str = modes.DEFAULT) -> str:
    """`/` を開いたときに入る画面。レールの先頭で、**実際に登録されている**もの。"""
    registered = {rule.rule for rule in current_app.url_map.iter_rules()}
    for item in nav_items(mode):
        if item.ready and item.url in registered:
            return item.url
    return FALLBACK_URL


def ribbon_values() -> dict[str, str]:
    """帯に常時出す値。ライン(この端末の設備名)・モード(オンライン/オフライン)・利用者・権限。"""
    from ufkensa import auth
    from ufkensa.kensa import paths, terminal

    s = auth.ensure()
    return {
        "line": terminal.line_label(),
        "mode": paths.MODE_LABELS.get(paths.mode(), ""),
        "user": s.user_id if s else "",
        "level": s.label if s else "",
        "developer": False,
    }


def shell_context(active: str, *,
                  badges: Optional[dict[str, tuple[str, str]]] = None) -> dict[str, Any]:
    """`base.html` が必要とする値一式。"""
    config = current_app.config
    mode = config["MODE"]
    badges = dict(badges or {})
    if "settings" not in badges:
        from ufkensa.kensa import terminal
        if not terminal.line():
            badges["settings"] = ("ライン未設定", "todo")
    return {
        "active": active,
        "nav": nav_items(mode, badges),
        "mode": mode,
        "display_name": config["DISPLAY_NAME"],
        "version_label": app_config.version_label(),
        "favicon": favicon(mode),
        "token": config["TOKEN"],
        # このタブの番号。自動終了の見張りが画面ごとに数えるためのもの
        "screen_id": request.headers.get("X-Screen-Id", "") or _new_screen_id(),
        "ribbon": ribbon_values(),
        "health_poll_ms": app_config.health_poll_seconds() * 1000,
        "log_poll_ms": max(500, app_config.job_poll_ms() * 3),
        "alive_poll_ms": idle_exit.HEARTBEAT_MS,
    }


def _new_screen_id() -> str:
    import uuid
    return uuid.uuid4().hex
