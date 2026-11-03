"""使用パス(ご依頼: 設定画面で見て直せるようにする)

VBA の `初期設定.Setting_Path` の 7 本と、コードに直書きされていたパスを、ここ 1 か所に集める。
**業務ロジックはパスを文字列で持たず、必ず `paths.get("quality")` のように引く。**

【判断事項 D26(2026-11-02 決め、11-03 改)】
- **仕掛台帳だけ** 第1パス / 第2パス を持つ(python-web-tools と同じ)。第1 = 台帳(VBA の MainPath)、
  第2 = 控え(NextPath)。2つ目(第2)は SIKALOT / SIKAHIKI / SIKAODR の**ファイル名を変えて**置ける
  (設定「2つ目の名前」。任意)。SIKACOMP / SIKATWOYEAR は VBA が NextPath から読むので第2。
- モード(オンライン / オフライン。VBA のレジストリ「モード」)は**仕掛台帳にだけ効く**:
    オンライン … 第1を読み、第1に無いもの(ファイルごと・ロットごと)だけ第2
    オフライン … 第2だけ。無ければ「…がありません」(呼び手が VBA の文言で出す)
- 仕掛 DB は手元へ写してから読む(`lot_copy`)。写せないファイルは「無い」とみなす
- 上流の作り直しの決まり(PLEASE WAIT.txt・毎時 59:46〜00:45)は使わない(11-03 のご指示で外した)
- ほかの項目(品証やり取り・参照用ファイル・検査履歴・工場共有・測定データ電子化・手元の検査履歴・音声・
  NPS・アクセス権限マスタ・Edge)は**パス 1 つ**。モードによらない

【D28】ロットごとの「控え(第2)から読む」(`lot_source`)の中は、モードによらずオフラインと同じ読み方(第2だけ)。

`get("main")` / `get("next")` は残す(呼び手が多く、VBA の `FPath = MainPath` の比べ方をそのまま使うため):
    main … オンラインは第1、オフラインは第2      next … いつも第2

- 端末ごとに `user_settings` に保存する。保存されていない項目は下の既定値(VBA のオンラインの値)。
  前の保存(オンライン/オフラインの 2 組、または 第1/第2 の 2 組)は `_saved` で読み替える。
- DB は SQLite に切り替えた(D4)。ファイル名は VBA の .accdb と同じ名前で拡張子が .sqlite3。
"""
from __future__ import annotations

import contextvars
import os
import time
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path, PureWindowsPath
from typing import Any, Callable, Iterator, Optional, Union

from .. import config, user_settings
from ..logging_utils import get_logger

log = get_logger("kensa.paths")

KEY_MODE = "kensa_mode"              # "online" / "offline"
KEY_PATHS = "kensa_paths"            # {"items": {key: path}, "lot": {"first": p, "second": p}, "second_names": {DB名: 名前}}
ONLINE, OFFLINE = "online", "offline"
MODE_LABELS = {ONLINE: "オンライン", OFFLINE: "オフライン"}
FIRST, SECOND = "first", "second"
SLOT_LABELS = {FIRST: "第1パス", SECOND: "第2パス"}
KEY_ITEMS, KEY_LOT, KEY_SECOND_NAMES = "items", "lot", "second_names"

LOT = "lot"                          # 仕掛台帳(VBA の MainPath = 第1 / NextPath = 第2)
ALIASES = ("main", "next")           # 前の項目名。main = いまのモードで読む側、next = 第2


@dataclass(frozen=True)
class PathItem:
    key: str
    label: str
    vba: str                 # VBA での名前(変数名・直書きの場所)
    path: str                # 既定値(VBA のオンラインの値)。仕掛台帳は第1の既定
    use: str                 # 何に使うか
    need: str = "read"       # read / write / exe(確認ボタンで見ること)
    kind: str = "dir"        # dir / file
    second: str = ""         # 仕掛台帳だけ: 第2の既定


ITEMS: tuple[PathItem, ...] = (
    PathItem(LOT, "仕掛台帳", "MainPath / NextPath", r"\\nlmsrvngy03\Read\【New】仕掛\台帳" "\\",
             "SIKALOT / SIKAHIKI / SIKAODR.sqlite3(上流が毎時作り直す)。第1 = 台帳、第2 = 控え。"
             "オンラインは第1に無いファイル・ロットだけ第2、オフラインは第2だけ。SIKACOMP / SIKATWOYEAR は第2。"
             "手元へ写してから読む",
             second=r"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\仕掛" "\\"),
    PathItem("quality", "品証やり取り", "QualityPath",
             r"\\nlmsrvngy03\工場内共有\検査データ\品証-梱包やり取りフォルダ\Accdb" "\\",
             "品質基準・伝送用ファイル・包装仕様書(画像)"),
    PathItem("packing", "参照用ファイル", "PackingPath",
             r"\\nlmfangyshrd\各課共有\0130_日軽稲沢\梱包課\AIM\【■】_参照用ファイル" "\\",
             "伝送用ファイルの代替・VC種類画像"),
    PathItem("history", "検査履歴", "HistoryPath",
             r"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\検査履歴" "\\",
             "ライン別ログ・Aインプット CSV・履歴のコピー", need="write"),
    PathItem("factory", "工場共有 検査データ", "FactoryyPath", r"\\nlmsrvngy03\工場内共有\検査データ" "\\",
             "梱包検査ﾃﾞｰﾀ(測定DB)・PDF・キャスク/5N", need="write"),
    PathItem("nis", "測定データ電子化", "NisPath", r"C:\Tools\測定データ電子化",
             "hostsend.txt・SokuteiDataTrans.exe・result.txt(空なら使わない。VBA はオフライン時に空)", need="write"),
    PathItem("local_history", "手元の検査履歴", "C:\\検査履歴(直書き)", r"C:\検査履歴",
             "帳票 PDF の手元保存", need="write"),
    PathItem("wav", "音声", "C:\\VBAWAV(直書き)", r"C:\VBAWAV" "\\",
             "判定音・案内の音声のフォルダ(何を鳴らすかは 設定 → 音。D29)"),
    PathItem("nps_mitsubishi", "三菱受託検査 履歴(NPS)", "直書き(Mitsubishi)",
             r"\\nlmsrvngy03\NPS\全社研修会トレーナー教育\第36回NPSトレーナー教育（名古屋）2017年\梱包実績の見える化\三菱受託検査_履歴" "\\",
             "三菱受託の xlsx(→PDF+CSV)の写し", need="write"),
    PathItem("access_master", "アクセス権限マスタ", "(python-web-tools と共通)", "",
             "「アクセス権限」表がある sqlite3(梱包資材マスタ)。空なら権限の表を読まない", kind="file"),
    PathItem("edge", "Edge(PDF 作成)", "(新規 D7)",
             r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
             "帳票の PDF をヘッドレス印刷で作る", need="exe", kind="file"),
)
BY_KEY = {i.key: i for i in ITEMS}

#: 仕掛台帳の 3 ファイル(第2はファイル名を変えられる)
LOT_DBS = ("SIKALOTNOW", "SIKAHIKINOW", "SIKAODRNOW")
#: 仕掛として手元へ写して読む DB
COPIED_DBS = LOT_DBS + ("SIKACOMP", "SIKATWOYEAR")

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
# モード(仕掛台帳にだけ効く)
# ------------------------------------------------------------------
def mode() -> str:
    m = str(user_settings.get(KEY_MODE, ONLINE) or ONLINE)
    return m if m in (ONLINE, OFFLINE) else ONLINE


def set_mode(m: str) -> None:
    if m not in (ONLINE, OFFLINE):
        raise ValueError(m)
    user_settings.save(KEY_MODE, m)


# ------------------------------------------------------------------
# ロットごとの「控え(第2)から読む」(判断事項 D28)
# ------------------------------------------------------------------
#: いまの操作で仕掛台帳を第2だけから読むか。真偽、または読むたびに決める関数(検査画面の状態を見る)
_lot_second: contextvars.ContextVar[Union[bool, Callable[[], bool]]] = contextvars.ContextVar(
    "kensa_lot_second", default=False)


@contextmanager
def lot_source(second: Union[bool, Callable[[], bool]]) -> Iterator[None]:
    """この中の仕掛台帳の読みを、真なら**モードによらず第2(控え)だけ**にする(オフラインと同じ読み方)。

    `second` に関数を渡すと読むたびに呼ぶ(操作の途中でロットが変わっても、そのロットの選択に従う)。
    VBA の DBflag の代わり(D27 で DBflag は使わない)。校正・Min/Max の確かめには効かない。
    """
    token = _lot_second.set(second)
    try:
        yield
    finally:
        _lot_second.reset(token)


def lot_second_forced() -> bool:
    """いま仕掛台帳を第2だけから読む指定があるか。"""
    v = _lot_second.get()
    try:
        return bool(v()) if callable(v) else bool(v)
    except Exception:                                   # noqa: BLE001 - 決められなければ いつもの読み方
        return False


def lot_mode(m: Optional[str] = None) -> str:
    """仕掛台帳の読み方に使うモード。指定があればそれ、「控えから読む」中はオフライン、ほかは設定のモード。"""
    if m:
        return m
    return OFFLINE if lot_second_forced() else mode()


# ------------------------------------------------------------------
# 保存の形と、前の形からの読み替え
# ------------------------------------------------------------------
def _d(v: Any) -> dict:
    return v if isinstance(v, dict) else {}


def _norm(raw: dict, m: str) -> dict[str, Any]:
    """保存値 → {"items": {}, "lot": {"first"?, "second"?}, "second_names": {}}。

    前の形は 2 つ:
    1. {"online": {...}, "offline": {...}}(最初の形)
    2. {"first": {...}, "second": {...}}(11-02 の 第1/第2 の形)
    パス 1 つの項目は **いまのモードで使っていた値** を残す(オンラインなら旧オンライン / 第1、
    オフラインなら旧オフライン / 第2)。仕掛台帳は 第1 ← 旧オンラインの main(または第1)、
    第2 ← 旧オンラインの next → 旧オフラインの next → main の順(または第2)。
    """
    items: dict[str, str] = {}
    lot: dict[str, str] = {}
    on, off = _d(raw.get(ONLINE)), _d(raw.get(OFFLINE))
    if on or off:
        cur = on if m == ONLINE else off
        items.update({k: str(v) for k, v in cur.items() if k not in ALIASES and k in BY_KEY})
        if "main" in on:
            lot[FIRST] = str(on["main"])
        for v in (on.get("next"), off.get("next"), off.get("main")):
            if v is not None:
                lot[SECOND] = str(v)
                break
    f, s = _d(raw.get(FIRST)), _d(raw.get(SECOND))
    if f or s:
        cur = f if m == ONLINE else s
        items.update({k: str(v) for k, v in cur.items() if k not in ALIASES and k != LOT and k in BY_KEY})
        for src in (f, s):
            if "main" in src:
                lot[FIRST if src is f else SECOND] = str(src["main"])
            if "next" in src:
                lot[SECOND] = str(src["next"])
        if LOT in f:
            lot[FIRST] = str(f[LOT])
        if LOT in s:
            lot[SECOND] = str(s[LOT])
    items.update({k: str(v) for k, v in _d(raw.get(KEY_ITEMS)).items() if k in BY_KEY and k != LOT})
    lot.update({k: str(v) for k, v in _d(raw.get(KEY_LOT)).items() if k in (FIRST, SECOND)})
    names = {k: str(v) for k, v in _d(raw.get(KEY_SECOND_NAMES)).items() if k in LOT_DBS}
    return {KEY_ITEMS: items, KEY_LOT: lot, KEY_SECOND_NAMES: names}


def _saved() -> dict[str, Any]:
    raw = _d(user_settings.get(KEY_PATHS, {}) or {})
    d = _norm(raw, mode())
    if any(k in raw for k in (ONLINE, OFFLINE, FIRST, SECOND)):
        user_settings.save(KEY_PATHS, d)            # 前の形を見つけたら、読み替えた形で保存し直す(1 回だけ)
        log.info("使用パスの保存の形を読み替えました(仕掛台帳だけ第1/第2、ほかはパス1つ)")
    return d


def app_dir() -> str:
    return str(config.BASE_DIR) + os.sep


def _key(key: str) -> str:
    return LOT if key in ALIASES else key


def _lot_slot(key: str, slot: Optional[str]) -> str:
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
        return FIRST
    raise ValueError(slot)


# ------------------------------------------------------------------
# 値を引く
# ------------------------------------------------------------------
def slot_default(key: str, slot: str = FIRST) -> str:
    """既定値。仕掛台帳は第1 / 第2、ほかは slot によらず 1 つ。"""
    item = BY_KEY[_key(key)]
    return item.second if (item.key == LOT and slot == SECOND) else item.path


def slot_value(key: str, slot: str = FIRST) -> str:
    """仕掛台帳はその側(第1 / 第2)のパス。ほかの項目は slot によらずパス 1 つ。"""
    k = _key(key)
    d = _saved()
    if k == LOT:
        v = d[KEY_LOT].get(slot)
    else:
        v = d[KEY_ITEMS].get(k)
    return slot_default(k, slot) if v is None else str(v)


def which(key: str, m: Optional[str] = None) -> str:
    """仕掛台帳でどちらの側を使うか。main / 仕掛台帳 = オンライン第1・オフライン第2、next = 第2。"""
    if key == "next":
        return SECOND
    return FIRST if lot_mode(m) == ONLINE else SECOND


def default(key: str, m: Optional[str] = None) -> str:
    k = _key(key)
    return slot_default(k, which(key, m)) if k == LOT else slot_default(k)


def get(key: str, m: Optional[str] = None) -> str:
    """パス(文字列のまま。Windows のパス)。仕掛台帳(main / next)だけモードで変わる。"""
    k = _key(key)
    if k == LOT:
        return slot_value(LOT, which(key, m))
    return slot_value(k)


def set_path(key: str, value: str, slot: Optional[str] = None) -> None:
    """パスを決める。仕掛台帳は `slot`(FIRST / SECOND。main = 第1、next = 第2)。ほかは slot を見ない。"""
    k = _key(key)
    if k not in BY_KEY:
        raise KeyError(key)
    d = _saved()
    if k == LOT:
        d[KEY_LOT][_lot_slot(key, slot)] = str(value).strip()
    else:
        d[KEY_ITEMS][k] = str(value).strip()
    user_settings.save(KEY_PATHS, d)


def reset_path(key: str, slot: Optional[str] = None) -> None:
    k = _key(key)
    d = _saved()
    bucket, name = (d[KEY_LOT], _lot_slot(key, slot)) if k == LOT else (d[KEY_ITEMS], k)
    if name in bucket:
        del bucket[name]
        user_settings.save(KEY_PATHS, d)


def snapshot(key: str) -> dict[str, str]:
    """いまの値(ログに前後を残す形)。仕掛台帳は 第1・第2・2つ目の名前。"""
    k = _key(key)
    if k == LOT:
        out = {SLOT_LABELS[FIRST]: slot_value(LOT, FIRST), SLOT_LABELS[SECOND]: slot_value(LOT, SECOND)}
        out.update({f"2つ目の名前 {db}": second_name(db) for db in LOT_DBS})
        return out
    return {BY_KEY[k].label: slot_value(k)}


def item_changed(key: str) -> bool:
    """その項目のどこかが既定から変わっているか(仕掛台帳は 第1・第2・2つ目の名前のどれか)。"""
    k = _key(key)
    d = _saved()
    if k == LOT:
        return bool(d[KEY_LOT]) or bool(d[KEY_SECOND_NAMES])
    return k in d[KEY_ITEMS]


def reset_item(key: str) -> None:
    """項目を既定に戻す(設定「既定に戻す」)。仕掛台帳は 第1・第2・2つ目の名前をまとめて戻す。"""
    k = _key(key)
    if k not in BY_KEY:
        raise KeyError(key)
    d = _saved()
    if k == LOT:
        d[KEY_LOT].clear()
        d[KEY_SECOND_NAMES].clear()
    else:
        d[KEY_ITEMS].pop(k, None)
    user_settings.save(KEY_PATHS, d)


def reset_all() -> list[str]:
    """すべての項目を既定(VBA のオンラインの値)に戻す。戻した項目の key を返す。モードは変えない。"""
    done = [i.key for i in ITEMS if item_changed(i.key)]
    user_settings.save(KEY_PATHS, {KEY_ITEMS: {}, KEY_LOT: {}, KEY_SECOND_NAMES: {}})
    return done


def changed(key: str, slot: str = FIRST) -> bool:
    k = _key(key)
    d = _saved()
    return slot in d[KEY_LOT] if k == LOT else k in d[KEY_ITEMS]


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
    m = lot_mode(m)
    second = slot_value(LOT, SECOND)
    if base == second and (m == OFFLINE or base != slot_value(LOT, FIRST)):
        return second_name(db)
    return first_name(db)


def copied_names() -> set[str]:
    """手元へ写して読む仕掛 DB のファイル名(1つ目・2つ目の名前・SIKACOMP・SIKATWOYEAR)。"""
    out = {DB_FILES[db][1] for db in COPIED_DBS}
    out |= {second_name(db) for db in LOT_DBS}
    return {n.lower() for n in out}


def is_copied(path: Any) -> bool:
    """仕掛 DB(手元へ写して読むもの)か。ファイル名で見分ける。"""
    name = Path(str(path).replace("\\", "/")).name if not isinstance(path, Path) else path.name
    return name.lower() in copied_names()


def lot_usable(path_text: str) -> bool:
    """仕掛 DB を読めるか = 手元へ写せたか(写せなければ読まない。D26)。"""
    from . import lot_copy
    return lot_copy.local_copy(path_text) is not None


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


def db_file(name: str, m: Optional[str] = None, *, base_key: Optional[str] = None, write: bool = False) -> str:
    """DB の場所。`base_key` で別の場所(例 仕掛の控え "next"・"packing")から引ける。

    仕掛の 3 ファイル(main): オンラインは第1(写せれば)→ 第2(写せれば)、どちらも駄目なら第1の場所
    (呼び手が VBA どおり「…がありません」と出す)。オフラインは第2だけ。
    SIKACOMP / SIKATWOYEAR(next)はいつも第2。ほかの DB はその項目のパス 1 つ。`write` は前の呼び方の名残。
    """
    m = lot_mode(m)
    where, file = DB_FILES[name]
    key = base_key or where
    if key == "next":
        return join(slot_value(LOT, SECOND), second_name(name))
    k = _key(key)
    if k != LOT:
        return join(slot_value(k), file)
    if m == OFFLINE:
        return join(slot_value(LOT, SECOND), second_name(name))
    p1 = join(slot_value(LOT, FIRST), file)
    p2 = join(slot_value(LOT, SECOND), second_name(name))
    if p1 == p2 or lot_usable(p1):
        return p1
    return p2 if lot_usable(p2) else p1


#: 試験の安全網(tests/_isolation.py が立てる)。立っていれば、共有(\\server\…)と、一時フォルダの外のドライブのパス
#: (C:\検査履歴 など = 既定の本番のパス)を、この下へ読み替える。試験が本番のフォルダを読み書きしないように
GUARD_ENV = "UFKENSA_TEST_PATH_GUARD"


def _guarded(path_text: str) -> Optional[Path]:
    guard = os.environ.get(GUARD_ENV)
    if not guard or not path_text or os.environ.get("UFKENSA_PATH_ROOT"):
        return None
    t = path_text.strip()
    unc = t.startswith(("\\\\", "//"))
    drive = len(t) >= 3 and t[0].isalpha() and t[1] == ":" and t[2] in "\\/"
    if not unc and not drive:
        return None
    if drive and os.name == "nt":
        import tempfile
        tmp = os.path.normcase(os.path.abspath(tempfile.gettempdir()))
        if os.path.normcase(os.path.abspath(t)).startswith(tmp):
            return None                                   # 試験が作った一時フォルダ(Windows)はそのまま
    pw = PureWindowsPath(t.replace("/", "\\"))
    parts = [p.replace(":", "").strip("\\") for p in pw.parts if p.strip("\\")]
    return Path(guard, "本番のパス", *parts)


def to_local(path_text: str) -> Path:
    """Windows のパス文字列 → この OS の Path(試験では UFKENSA_PATH_ROOT の下に読み替える)。"""
    guarded = _guarded(path_text)
    if guarded is not None:
        return guarded
    root = os.environ.get("UFKENSA_PATH_ROOT")
    if not root and os.name != "nt" and "\\" in path_text and not path_text.startswith("\\\\"):
        return Path(path_text.replace("\\", "/"))             # 試験(Linux)で "…/fixtures\\品質基準.sqlite3" を読む
    if root:
        pw = PureWindowsPath(path_text)
        parts = [p.replace(":", "").strip("\\") for p in pw.parts if p.strip("\\")]
        return Path(root, *parts)
    return Path(path_text)


_SLOW_SEC = 0.5          # これより時間がかかった確認(届かない共有など)は
_SLOW_KEEP_SEC = 30.0    # この間は覚えておき、毎回待たせない
_slow: dict[tuple[str, str], tuple[bool, float]] = {}


def exists(path_text: str, kind: str = "any") -> bool:
    """パス(Windows の文字列)があるか。届かないときも偽。時間のかかった確認だけ少しの間覚えておく。"""
    if not path_text:
        return False
    key = (path_text, kind)
    hit = _slow.get(key)
    now = time.monotonic()
    if hit is not None and now - hit[1] < _SLOW_KEEP_SEC:
        return hit[0]
    try:
        p = to_local(path_text)
        ok = p.is_dir() if kind == "dir" else (p.is_file() if kind == "file" else p.exists())
    except (OSError, ValueError):
        ok = False
    if time.monotonic() - now > _SLOW_SEC:
        _slow[key] = (ok, time.monotonic())
    else:
        _slow.pop(key, None)
    return ok


# ------------------------------------------------------------------
# 確認ボタン・一覧
# ------------------------------------------------------------------
def check(key: str, slot: Optional[str] = None) -> dict[str, Any]:
    """確認ボタン: ある・読める・書けるか。仕掛台帳はその側(第1 / 第2)と、各ファイルを手元へ写せるか。"""
    k = _key(key)
    item = BY_KEY[k]
    s = _lot_slot(key, slot) if k == LOT else ""
    text = slot_value(k, s or FIRST)
    out: dict[str, Any] = {"key": k, "slot": s, "label": SLOT_LABELS.get(s, ""), "path": text, "exists": False,
                           "readable": False, "writable": False, "ok": False, "message": "", "files": []}
    if not text:
        out["message"] = "設定されていません"
        return out
    p = to_local(text)
    try:
        out["exists"] = p.exists()
        if item.kind == "dir":
            out["readable"] = out["exists"] and os.access(p, os.R_OK | os.X_OK)
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
            there = exists(join(text, fname), "file")
            ok = there and lot_usable(join(text, fname))
            out["files"].append({"db": db, "name": fname, "exists": there, "copied": ok})
            have.append(f"{fname} " + ("あり" if ok else ("写せません" if there else "なし")))
        out["message"] += "(" + "・".join(have) + ")"
    return out


def in_use_text(key: str, m: str) -> str:
    """一覧に出す「このモードで使う側」(仕掛台帳だけ)。"""
    if _key(key) != LOT:
        return ""
    return "第1(無い分は第2)" if m == ONLINE else "第2だけ"


def overview(m: Optional[str] = None) -> list[dict[str, Any]]:
    """設定画面の一覧。仕掛台帳は第1 / 第2 / 2つ目の名前、ほかはパス 1 つ。"""
    rows = []
    for i in ITEMS:
        row: dict[str, Any] = {"key": i.key, "label": i.label, "vba": i.vba, "use": i.use, "need": i.need,
                               "kind": i.kind, "single": i.key != LOT}
        if i.key == LOT:
            row.update({
                "first": slot_value(LOT, FIRST), "second": slot_value(LOT, SECOND),
                "first_default": slot_default(LOT, FIRST), "second_default": slot_default(LOT, SECOND),
                "first_changed": changed(LOT, FIRST), "second_changed": changed(LOT, SECOND),
                "in_use": {mm: in_use_text(LOT, mm) for mm in MODE_LABELS},
                "second_names": [{"db": db, "first": first_name(db), "value": second_name_setting(db),
                                  "effective": second_name(db)} for db in LOT_DBS],
            })
        else:
            row.update({"path": slot_value(i.key), "default": slot_default(i.key), "changed": changed(i.key)})
        rows.append(row)
    return rows
