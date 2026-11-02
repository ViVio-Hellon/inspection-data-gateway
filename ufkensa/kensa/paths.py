"""使用パス(ご依頼: 設定画面で見て直せるようにする)

VBA の `初期設定.Setting_Path` の 7 本と、コードに直書きされていたパスを、ここ 1 か所に集める。
**業務ロジックはパスを文字列で持たず、必ず `paths.get("main")` のように引く。**

- オンライン / オフラインの 2 組を持つ(VBA のレジストリ「モード」と同じ考え方)。
  オフラインの既定は VBA と同じく「アプリのフォルダ」(ThisWorkbook.Path に当たる)。
- 端末ごとに `user_settings` に保存する。保存されていない項目は下の既定値。
- DB は SQLite に切り替えた(D4)。ファイル名は VBA の .accdb と同じ名前で拡張子が .sqlite3。
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path, PureWindowsPath
from typing import Any, Optional

from .. import config, user_settings

KEY_MODE = "kensa_mode"              # "online" / "offline"
KEY_PATHS = "kensa_paths"            # {"online": {key: path}, "offline": {key: path}}
ONLINE, OFFLINE = "online", "offline"
MODE_LABELS = {ONLINE: "オンライン", OFFLINE: "オフライン"}


@dataclass(frozen=True)
class PathItem:
    key: str
    label: str
    vba: str                 # VBA での名前(変数名・直書きの場所)
    online: str              # オンラインの既定値(VBA の値)
    offline: str             # オフラインの既定値("" = アプリのフォルダ)
    use: str                 # 何に使うか
    need: str = "read"       # read / write / exe(確認ボタンで見ること)
    kind: str = "dir"        # dir / file


_APP = "<APP>"   # オフラインの既定 = アプリのフォルダ(VBA の ThisWorkbook.Path)

ITEMS: tuple[PathItem, ...] = (
    PathItem("main", "仕掛台帳", "MainPath", r"\\nlmsrvngy03\Read\【New】仕掛\台帳" "\\", _APP,
             "SIKALOT / SIKAHIKI / SIKAODR.sqlite3(上流が毎時作り直す)"),
    PathItem("next", "仕掛の控え", "NextPath",
             r"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\仕掛" "\\", _APP,
             "上流の作り直し中(PLEASE WAIT.txt・毎時 59:46〜00:45)に読む控え"),
    PathItem("quality", "品証やり取り", "QualityPath",
             r"\\nlmsrvngy03\工場内共有\検査データ\品証-梱包やり取りフォルダ\Accdb" "\\", _APP,
             "品質基準・伝送用ファイル・包装仕様書(画像)"),
    PathItem("packing", "参照用ファイル", "PackingPath",
             r"\\nlmfangyshrd\各課共有\0130_日軽稲沢\梱包課\AIM\【■】_参照用ファイル" "\\", _APP,
             "伝送用ファイルの代替・VC種類画像"),
    PathItem("history", "検査履歴", "HistoryPath",
             r"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\検査履歴" "\\", _APP,
             "ライン別ログ・Aインプット CSV・履歴のコピー", need="write"),
    PathItem("factory", "工場共有 検査データ", "FactoryyPath", r"\\nlmsrvngy03\工場内共有\検査データ" "\\", _APP,
             "梱包検査ﾃﾞｰﾀ(測定DB)・PDF・キャスク/5N", need="write"),
    PathItem("nis", "測定データ電子化", "NisPath", r"C:\Tools\測定データ電子化", "",
             "hostsend.txt・SokuteiDataTrans.exe・result.txt(VBA はオフライン時は空)", need="write"),
    PathItem("local_history", "手元の検査履歴", "C:\\検査履歴(直書き)", r"C:\検査履歴", r"C:\検査履歴",
             "帳票 PDF の手元保存", need="write"),
    PathItem("wav", "音声", "C:\\VBAWAV(直書き)", r"C:\VBAWAV" "\\", r"C:\VBAWAV" "\\",
             "判定音・案内の音声(.wav)"),
    PathItem("nps_mitsubishi", "三菱受託検査 履歴(NPS)", "直書き(Mitsubishi)",
             r"\\nlmsrvngy03\NPS\全社研修会トレーナー教育\第36回NPSトレーナー教育（名古屋）2017年\梱包実績の見える化\三菱受託検査_履歴" "\\",
             r"\\nlmsrvngy03\NPS\全社研修会トレーナー教育\第36回NPSトレーナー教育（名古屋）2017年\梱包実績の見える化\三菱受託検査_履歴" "\\",
             "三菱受託の xlsx(→PDF+CSV)の写し", need="write"),
    PathItem("access_master", "アクセス権限マスタ", "(python-web-tools と共通)", "", "",
             "「アクセス権限」表がある sqlite3(梱包資材マスタ)。空なら権限の表を読まない", kind="file"),
    PathItem("edge", "Edge(PDF 作成)", "(新規 D7)",
             r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
             r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
             "帳票の PDF をヘッドレス印刷で作る", need="exe", kind="file"),
)
BY_KEY = {i.key: i for i in ITEMS}

#: DB ファイル(D4: すべて SQLite)。場所 → ファイル名
DB_FILES = {
    # 仕掛の3つは上流が SQLite で作る(VBA の SIKALOTNOW.accdb → SIKALOT.sqlite3 など)。列名は同じ
    "SIKALOTNOW": ("main", "SIKALOT.sqlite3"),
    "SIKAHIKINOW": ("main", "SIKAHIKI.sqlite3"),
    "SIKAODRNOW": ("main", "SIKAODR.sqlite3"),
    "SIKACOMP": ("next", "SIKACOMP.sqlite3"),
    "SIKATWOYEAR": ("next", "SIKATWOYEAR.sqlite3"),     # 仕掛_完了2年(Input:293。VBA は NextPath から読む)
    "品質基準": ("quality", "品質基準.sqlite3"),
    "伝送用ファイル": ("quality", "伝送用ファイル.sqlite3"),
    "梱包検査ﾃﾞｰﾀ": ("factory", "【測定DB】梱包\\梱包検査ﾃﾞｰﾀ.sqlite3"),
}


def mode() -> str:
    m = str(user_settings.get(KEY_MODE, ONLINE) or ONLINE)
    return m if m in (ONLINE, OFFLINE) else ONLINE


def set_mode(m: str) -> None:
    if m not in (ONLINE, OFFLINE):
        raise ValueError(m)
    user_settings.save(KEY_MODE, m)


def _saved() -> dict[str, dict[str, str]]:
    d = user_settings.get(KEY_PATHS, {}) or {}
    return d if isinstance(d, dict) else {}


def app_dir() -> str:
    return str(config.BASE_DIR) + os.sep


def default(key: str, m: Optional[str] = None) -> str:
    item = BY_KEY[key]
    v = item.online if (m or mode()) == ONLINE else item.offline
    return app_dir() if v == _APP else v


def get(key: str, m: Optional[str] = None) -> str:
    """いまのモードでのパス(文字列のまま。Windows のパス)。"""
    m = m or mode()
    saved = _saved().get(m, {})
    v = saved.get(key)
    return default(key, m) if v is None else str(v)


def set_path(key: str, value: str, m: Optional[str] = None) -> None:
    if key not in BY_KEY:
        raise KeyError(key)
    m = m or mode()
    d = _saved()
    d.setdefault(m, {})[key] = str(value).strip()
    user_settings.save(KEY_PATHS, d)


def reset_path(key: str, m: Optional[str] = None) -> None:
    m = m or mode()
    d = _saved()
    if key in d.get(m, {}):
        del d[m][key]
        user_settings.save(KEY_PATHS, d)


def join(base: str, *parts: str) -> str:
    """VBA の `path & "名前"` と同じく文字列でつなぐ(区切りが無ければ足す)。"""
    out = base
    for p in parts:
        if out and not out.endswith(("\\", "/")):
            out += "\\"
        out += p
    return out


def db_file(name: str, m: Optional[str] = None, *, base_key: Optional[str] = None) -> str:
    """DB の場所。`base_key` で別の場所(例 仕掛の控え "next")から引ける。"""
    where, file = DB_FILES[name]
    return join(get(base_key or where, m), file)


def to_local(path_text: str) -> Path:
    """Windows のパス文字列 → この OS の Path(試験では UFKENSA_PATH_ROOT の下に読み替える)。"""
    root = os.environ.get("UFKENSA_PATH_ROOT")
    if not root and os.name != "nt" and "\\" in path_text and not path_text.startswith("\\\\"):
        return Path(path_text.replace("\\", "/"))             # 試験(Linux)で "…/fixtures\\品質基準.sqlite3" を読む
    if root:
        pw = PureWindowsPath(path_text)
        parts = [p.replace(":", "").strip("\\") for p in pw.parts if p.strip("\\")]
        return Path(root, *parts)
    return Path(path_text)


def check(key: str, m: Optional[str] = None) -> dict[str, Any]:
    """確認ボタン: ある・読める・書けるか。"""
    item = BY_KEY[key]
    text = get(key, m)
    out: dict[str, Any] = {"key": key, "path": text, "exists": False, "readable": False,
                           "writable": False, "ok": False, "message": ""}
    if not text:
        out["message"] = "設定されていません"
        return out
    p = to_local(text)
    try:
        out["exists"] = p.exists()
        if item.kind == "dir":
            out["readable"] = out["exists"] and os.access(p, os.R_OK | os.X_OK)
            out["writable"] = out["exists"] and os.access(p, os.W_OK)
        else:
            out["readable"] = out["exists"] and os.access(p, os.R_OK)
            out["writable"] = out["exists"] and os.access(p, os.W_OK)
    except OSError as e:
        out["message"] = f"確かめられません: {e}"
        return out
    if not out["exists"]:
        out["message"] = "見つかりません"
    elif item.need == "write" and not out["writable"]:
        out["message"] = "書き込めません"
    elif not out["readable"]:
        out["message"] = "読めません"
    else:
        out["ok"] = True
        out["message"] = "書けます" if item.need == "write" else ("あります" if item.need == "exe" else "読めます")
    return out


def overview(m: Optional[str] = None) -> list[dict[str, Any]]:
    m = m or mode()
    saved = _saved().get(m, {})
    return [{"key": i.key, "label": i.label, "vba": i.vba, "use": i.use, "need": i.need, "kind": i.kind,
             "path": get(i.key, m), "default": default(i.key, m), "changed": i.key in saved}
            for i in ITEMS]
