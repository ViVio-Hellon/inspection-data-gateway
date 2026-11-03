"""使用パス(ご依頼: 設定画面で見て直せるようにする)

VBA の `初期設定.Setting_Path` の 7 本と、コードに直書きされていたパスを、ここ 1 か所に集める。
**業務ロジックはパスを文字列で持たず、必ず `paths.get("quality")` のように引く。**

【第1パス / 第2パス(判断事項 D26。2026-11-02 の決め)】
どの項目も **第1パス** と **第2パス** を持つ(前の「オンライン / オフライン」の 2 組を言い換えたもの。
保存済みの設定は 旧オンライン → 第1、旧オフライン → 第2 に読み替える `_saved`)。
モード(VBA のレジストリ「モード」)で読み方が変わる:

- オンライン: 第1パスを読む。**第1に無いものだけ** 第2パスを読む(読む項目)
- オフライン: 第2パスだけを読む。無ければ「…がありません」(呼び手がそのまま VBA の文言で出す)
- 書く項目(検査履歴・工場共有・測定データ電子化・手元の検査履歴・NPS): オンライン → 第1、オフライン → 第2。
  書けないからといって黙ってもう一方へ書くことはしない(VBA と同じく書けなかったと知らせる)

【仕掛台帳(引当マスタ)】python-web-tools と同じ考え方。第1 = 台帳(VBA の MainPath)、第2 = 控え(NextPath)。
`get("main")` / `get("next")` は残す(呼び手が多く、VBA の `FPath = MainPath` の比べ方をそのまま使うため):
    main … オンラインは第1、オフラインは第2      next … いつも第2
2つ目(第2)は SIKALOT / SIKAHIKI / SIKAODR の**ファイル名を変えて**置ける(設定「2つ目の名前」。任意)。
ファイルごと・ロットごとに第2を見るのは `db_file`(ファイル)と `lot.py`(ロット)。
上流の作り直し(PLEASE WAIT.txt・毎時 59:46〜00:45)は「第1が使えない」に含める(lot.py の VBA どおりの切替)。
SIKACOMP / SIKATWOYEAR は VBA が NextPath から読むので、仕掛台帳の第2。

- 端末ごとに `user_settings` に保存する。保存されていない項目は下の既定値。
- DB は SQLite に切り替えた(D4)。ファイル名は VBA の .accdb と同じ名前で拡張子が .sqlite3。
"""
from __future__ import annotations

import os
import time
from dataclasses import dataclass
from pathlib import Path, PureWindowsPath
from typing import Any, Optional

from .. import config, user_settings
from ..logging_utils import get_logger

log = get_logger("kensa.paths")

KEY_MODE = "kensa_mode"              # "online" / "offline"
KEY_PATHS = "kensa_paths"            # {"first": {key: path}, "second": {key: path}, "second_names": {DB名: ファイル名}}
ONLINE, OFFLINE = "online", "offline"
MODE_LABELS = {ONLINE: "オンライン", OFFLINE: "オフライン"}
FIRST, SECOND = "first", "second"
SLOT_LABELS = {FIRST: "第1パス", SECOND: "第2パス"}
KEY_SECOND_NAMES = "second_names"

LOT = "lot"                          # 仕掛台帳(VBA の MainPath = 第1 / NextPath = 第2)
ALIASES = ("main", "next")           # 前の項目名。main = いまのモードで読む側、next = 第2
PLEASE_WAIT = "PLEASE WAIT.txt"      # 上流の作り直し中の印(lot.PLEASE_WAIT と同じ)


@dataclass(frozen=True)
class PathItem:
    key: str
    label: str
    vba: str                 # VBA での名前(変数名・直書きの場所)
    first: str               # 第1パスの既定値(VBA のオンラインの値)
    second: str              # 第2パスの既定値(VBA のオフラインの値。"<APP>" = アプリのフォルダ)
    use: str                 # 何に使うか
    need: str = "read"       # read / write / exe(確認ボタンで見ること。write は落ち先を持たない)
    kind: str = "dir"        # dir / file


_APP = "<APP>"   # 既定 = アプリのフォルダ(VBA の ThisWorkbook.Path)

ITEMS: tuple[PathItem, ...] = (
    PathItem(LOT, "仕掛台帳", "MainPath / NextPath", r"\\nlmsrvngy03\Read\【New】仕掛\台帳" "\\",
             r"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\仕掛" "\\",
             "SIKALOT / SIKAHIKI / SIKAODR.sqlite3(上流が毎時作り直す)。第1 = 台帳、第2 = 控え"
             "(作り直し中・第1に無いファイルやロットは第2を読む)。SIKACOMP / SIKATWOYEAR は第2"),
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

#: 仕掛台帳の 3 ファイル(第2はファイル名を変えられる)
LOT_DBS = ("SIKALOTNOW", "SIKAHIKINOW", "SIKAODRNOW")

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


# ------------------------------------------------------------------
# モード
# ------------------------------------------------------------------
def mode() -> str:
    m = str(user_settings.get(KEY_MODE, ONLINE) or ONLINE)
    return m if m in (ONLINE, OFFLINE) else ONLINE


def set_mode(m: str) -> None:
    if m not in (ONLINE, OFFLINE):
        raise ValueError(m)
    user_settings.save(KEY_MODE, m)


def mode_slot(m: Optional[str] = None) -> str:
    """そのモードで書く側(オンライン = 第1、オフライン = 第2)。"""
    return FIRST if (m or mode()) == ONLINE else SECOND


# ------------------------------------------------------------------
# 保存の形と、前の形(オンライン/オフライン)からの読み替え
# ------------------------------------------------------------------
def _norm_slots(raw: dict) -> dict[str, Any]:
    """保存値 → {"first": {}, "second": {}, "second_names": {}}。

    前の形 {"online": {...}, "offline": {...}} は 旧オンライン → 第1、旧オフライン → 第2。
    仕掛の main / next は 1 つの「仕掛台帳」にまとめる: 第1 ← 旧オンラインの main、
    第2 ← 旧オンラインの next(VBA のオンラインの NextPath)、無ければ旧オフラインの next、main の順。
    新しい形の値があればそちらを優先する。
    """
    first: dict[str, str] = {}
    second: dict[str, str] = {}
    lot2: list[str] = []
    on = raw.get(ONLINE) if isinstance(raw.get(ONLINE), dict) else {}
    off = raw.get(OFFLINE) if isinstance(raw.get(OFFLINE), dict) else {}
    for k, v in on.items():
        if k == "main":
            first[LOT] = str(v)
        elif k == "next":
            lot2.append(str(v))
        else:
            first[k] = str(v)
    for k in ("next", "main"):
        if k in off:
            lot2.append(str(off[k]))
    for k, v in off.items():
        if k not in ALIASES:
            second[k] = str(v)
    if lot2:
        second[LOT] = lot2[0]
    for slot, dest in ((FIRST, first), (SECOND, second)):
        d = raw.get(slot) if isinstance(raw.get(slot), dict) else {}
        for k, v in d.items():
            if k == "main":
                dest[LOT] = str(v)
            elif k == "next":
                second[LOT] = str(v)
            else:
                dest[k] = str(v)
    names = raw.get(KEY_SECOND_NAMES) if isinstance(raw.get(KEY_SECOND_NAMES), dict) else {}
    return {FIRST: first, SECOND: second, KEY_SECOND_NAMES: {k: str(v) for k, v in names.items() if k in LOT_DBS}}


def _saved() -> dict[str, Any]:
    raw = user_settings.get(KEY_PATHS, {}) or {}
    if not isinstance(raw, dict):
        raw = {}
    d = _norm_slots(raw)
    if ONLINE in raw or OFFLINE in raw or any(a in raw.get(s, {}) for s in (FIRST, SECOND)
                                              if isinstance(raw.get(s), dict) for a in ALIASES):
        # 前の形を見つけたら、読み替えた形で保存し直す(1 回だけ)
        user_settings.save(KEY_PATHS, d)
        log.info("使用パスの保存の形を 第1パス / 第2パス に読み替えました")
    return d


def app_dir() -> str:
    return str(config.BASE_DIR) + os.sep


def _key(key: str) -> str:
    return LOT if key in ALIASES else key


def _slot_arg(slot: Optional[str], key: str) -> str:
    """`set_path` などの slot 引数。前の呼び方(モード)も受ける: online → 第1、offline → 第2。"""
    if key == "main":
        return FIRST
    if key == "next":
        return SECOND
    if slot in (FIRST, SECOND):
        return slot
    if slot == ONLINE:
        return FIRST
    if slot == OFFLINE:
        return SECOND
    if slot is None:
        return mode_slot()
    raise ValueError(slot)


# ------------------------------------------------------------------
# 値を引く
# ------------------------------------------------------------------
def slot_default(key: str, slot: str) -> str:
    item = BY_KEY[_key(key)]
    v = item.first if slot == FIRST else item.second
    return app_dir() if v == _APP else v


def slot_value(key: str, slot: str) -> str:
    """その側(第1 / 第2)に決めてあるパス(決めていなければ既定)。"""
    k = _key(key)
    v = _saved()[slot].get(k)
    return slot_default(k, slot) if v is None else str(v)


_SLOW_SEC = 0.5          # これより時間がかかった確認(届かない共有など)は
_SLOW_KEEP_SEC = 30.0    # この間は覚えておき、毎回待たせない
_slow: dict[tuple[str, str], tuple[bool, float]] = {}


def exists(path_text: str, kind: str = "any") -> bool:
    """パス(Windows の文字列)があるか。届かないときも偽。

    第1 → 第2 の判定は読むたびに呼ばれるので、時間のかかった確認だけ少しの間覚えておく。
    """
    if not path_text:
        return False
    key = (path_text, kind)
    hit = _slow.get(key)
    now = time.monotonic()
    if hit is not None and now - hit[1] < _SLOW_KEEP_SEC:
        return hit[0]
    try:
        p = to_local(path_text)
        if kind == "dir":
            ok = p.is_dir()
        elif kind == "file":
            ok = p.is_file()
        else:
            ok = p.exists()
    except (OSError, ValueError):
        ok = False
    if time.monotonic() - now > _SLOW_SEC:
        _slow[key] = (ok, time.monotonic())
    else:
        _slow.pop(key, None)
    return ok


def which(key: str, m: Optional[str] = None) -> str:
    """いまのモードでどちらの側を使うか(FIRST / SECOND)。

    - next(仕掛の控え)… いつも第2。main / 仕掛台帳 … オンライン第1・オフライン第2
      (仕掛の第1 → 第2 の切替は lot.py が VBA どおりに、ファイルごと・ロットごとに行う)
    - 書く項目 … オンライン第1・オフライン第2(落ち先を持たない)
    - 読む項目 … オフラインは第2。オンラインは第1、第1に無く第2にあれば第2
    """
    m = m or mode()
    if key == "next":
        return SECOND
    k = _key(key)
    item = BY_KEY[k]
    if m == OFFLINE:
        return SECOND
    if k == LOT or item.need == "write":
        return FIRST
    p1, p2 = slot_value(k, FIRST), slot_value(k, SECOND)
    if not p2 or p1 == p2 or exists(p1, item.kind):
        return FIRST
    return SECOND if exists(p2, item.kind) else FIRST


def default(key: str, m: Optional[str] = None) -> str:
    """そのモードで主に使う側の既定値(オンライン = 第1、オフライン = 第2。next は第2)。"""
    slot = SECOND if key == "next" else mode_slot(m)
    return slot_default(key, slot)


def get(key: str, m: Optional[str] = None) -> str:
    """いまのモードでのパス(文字列のまま。Windows のパス)。"""
    return slot_value(key, which(key, m))


def set_path(key: str, value: str, slot: Optional[str] = None) -> None:
    """`slot` は FIRST / SECOND(前の呼び方の ONLINE / OFFLINE も可。省くといまのモードで書く側)。"""
    k = _key(key)
    if k not in BY_KEY:
        raise KeyError(key)
    s = _slot_arg(slot, key)
    d = _saved()
    d[s][k] = str(value).strip()
    user_settings.save(KEY_PATHS, d)


def reset_path(key: str, slot: Optional[str] = None) -> None:
    k = _key(key)
    s = _slot_arg(slot, key)
    d = _saved()
    if k in d[s]:
        del d[s][k]
        user_settings.save(KEY_PATHS, d)


# ------------------------------------------------------------------
# 仕掛台帳の 2つ目の名前
# ------------------------------------------------------------------
def first_name(db: str) -> str:
    return DB_FILES[db][1]


def second_name_setting(db: str) -> str:
    """設定「2つ目の名前」の入力そのまま(空 = 1つ目と同じ名前)。"""
    return str(_saved()[KEY_SECOND_NAMES].get(db, "") or "")


def second_name(db: str) -> str:
    """第2で探すファイル名。拡張子まで入れればその名前、省けば .sqlite3 を付ける(python-web-tools と同じ)。"""
    if db not in LOT_DBS:
        return DB_FILES[db][1]
    v = second_name_setting(db).strip()
    if not v:
        return first_name(db)
    return v if "." in v else v + ".sqlite3"


def set_second_name(db: str, name: str) -> None:
    """2つ目の名前を決める(空で 1つ目と同じ名前に戻す)。フォルダは入れられない。"""
    if db not in LOT_DBS:
        raise KeyError(db)
    name = str(name or "").strip().strip('"')
    if any(c in name for c in "\\/:"):
        raise ValueError("2つ目の名前にはファイル名だけを入れてください(フォルダは入れません)")
    d = _saved()
    if name:
        d[KEY_SECOND_NAMES][db] = name
    else:
        d[KEY_SECOND_NAMES].pop(db, None)
    user_settings.save(KEY_PATHS, d)


def lot_file_name(db: str, base: str, m: Optional[str] = None) -> str:
    """仕掛の `FPath & "SIKA….accdb"` のファイル名。第2(控え)のフォルダなら 2つ目の名前。

    VBA の切替(lot.py)は FPath を文字列で比べるので、ここもフォルダの文字列で見分ける。
    オフラインは main も next も第2なので 2つ目の名前。オンラインで第1と第2が同じフォルダなら 1つ目の名前。
    """
    m = m or mode()
    second = slot_value(LOT, SECOND)
    if base == second and (m == OFFLINE or base != slot_value(LOT, FIRST)):
        return second_name(db)
    return first_name(db)


# ------------------------------------------------------------------
# つなぐ・DB の場所
# ------------------------------------------------------------------
def join(base: str, *parts: str) -> str:
    """VBA の `path & "名前"` と同じく文字列でつなぐ(区切りが無ければ足す)。"""
    out = base
    for p in parts:
        if out and not out.endswith(("\\", "/")):
            out += "\\"
        out += p
    return out


def rebuilding(folder: str) -> bool:
    """上流が作り直し中か(第1の台帳に PLEASE WAIT.txt がある。VBA の FileDesignation)。"""
    return exists(join(folder, PLEASE_WAIT), "file")


def db_file(name: str, m: Optional[str] = None, *, base_key: Optional[str] = None, write: bool = False) -> str:
    """DB の場所。`base_key` で別の場所(例 仕掛の控え "next"・"packing")から引ける。

    ファイルごとに第1 → 第2 を見る(オンラインで読むとき)。第1に無く第2にも無ければ第1の場所を返す
    (呼び手が VBA どおり「…がありません」と出す)。オフラインは第2だけ。`write=True` は落ち先を持たない。
    仕掛の 3 ファイルは、第1に PLEASE WAIT.txt があるときも第1を使えないとみなす(上流の作り直し中)。
    """
    m = m or mode()
    where, file = DB_FILES[name]
    key = base_key or where
    if key == "next":
        return join(slot_value(LOT, SECOND), second_name(name))
    k = _key(key)
    item = BY_KEY[k]
    if write or item.need == "write":
        slot = mode_slot(m)
        return join(slot_value(k, slot), second_name(name) if (k == LOT and slot == SECOND) else file)
    if m == OFFLINE:
        return join(slot_value(k, SECOND), second_name(name) if k == LOT else file)
    d1, d2 = slot_value(k, FIRST), slot_value(k, SECOND)
    p1 = join(d1, file)
    p2 = join(d2, second_name(name) if k == LOT else file)
    if p1 == p2 or not d2:
        return p1
    if exists(p1, "file") and not (k == LOT and rebuilding(d1)):
        return p1
    return p2 if exists(p2, "file") else p1


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


# ------------------------------------------------------------------
# 確認ボタン・一覧
# ------------------------------------------------------------------
def check(key: str, slot: Optional[str] = None) -> dict[str, Any]:
    """確認ボタン: その側(第1 / 第2)が ある・読める・書けるか。仕掛台帳は 3 ファイルがあるかも見る。"""
    k = _key(key)
    item = BY_KEY[k]
    s = _slot_arg(slot, key)
    text = slot_value(k, s)
    out: dict[str, Any] = {"key": k, "slot": s, "path": text, "exists": False, "readable": False,
                           "writable": False, "ok": False, "message": "", "files": []}
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
    if k == LOT and out["exists"]:
        dbs = list(LOT_DBS) + (["SIKACOMP", "SIKATWOYEAR"] if s == SECOND else [])
        have = []
        for db in dbs:
            fname = second_name(db) if s == SECOND else first_name(db)
            ok = exists(join(text, fname), "file")
            out["files"].append({"db": db, "name": fname, "exists": ok})
            have.append(f"{fname} {'あり' if ok else 'なし'}")
        if s == FIRST and rebuilding(text):
            have.append(PLEASE_WAIT + " あり(作り直し中。第2を読みます)")
        out["message"] += "(" + "・".join(have) + ")"
    return out


def in_use_text(key: str, m: str) -> str:
    """一覧に出す「このモードで使う側」。"""
    k = _key(key)
    item = BY_KEY[k]
    if m == OFFLINE:
        return "第2"
    if k == LOT:
        return "第1(無い分は第2)"
    if item.need == "write":
        return "第1"
    return "第1" if which(k, m) == FIRST else "第2(第1に無い)"


def overview(m: Optional[str] = None) -> list[dict[str, Any]]:
    """設定画面の一覧。第1 / 第2 の値と、モードごとに使う側。"""
    d = _saved()
    rows = []
    for i in ITEMS:
        row: dict[str, Any] = {
            "key": i.key, "label": i.label, "vba": i.vba, "use": i.use, "need": i.need, "kind": i.kind,
            "first": slot_value(i.key, FIRST), "second": slot_value(i.key, SECOND),
            "first_default": slot_default(i.key, FIRST), "second_default": slot_default(i.key, SECOND),
            "first_changed": i.key in d[FIRST], "second_changed": i.key in d[SECOND],
            "in_use": {mm: in_use_text(i.key, mm) for mm in MODE_LABELS},
        }
        if m is not None:
            row["path"] = get(i.key, m)
        if i.key == LOT:
            row["second_names"] = [{"db": db, "first": first_name(db), "value": second_name_setting(db),
                                    "effective": second_name(db)} for db in LOT_DBS]
        rows.append(row)
    return rows
