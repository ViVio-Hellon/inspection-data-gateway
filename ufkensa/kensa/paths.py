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
KEY_FIRST_NAMES = "first_names"      # 第1パスで探すファイル名 {DB名: 名前}

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
    PathItem("pack_spec", "包装仕様書(画像)", "QualityPath & \"包装仕様書\"", "",
             "包装仕様書の図(<包装仕様No>.JPG)。空なら 品証やり取り\\包装仕様書"),
    PathItem("vc_images", "VC種類画像", "PackingPath & \"VC種類画像\"", "",
             "VC の種類の画像(<品名ｺｰﾄﾞ>.JPG)。空なら 参照用ファイル\\VC種類画像"),
    PathItem("quality_photo", "品質写真", "QualityPath & \"品質写真\"", "",
             "品質基準画面の写真(フォルダごとにタブ、中の .jpg)。空なら 品証やり取り\\品質写真"),
    # 2026-11-09: 判定の画像はフォルダをやめ、絵ごとにファイルを決める(空なら同梱の VBA の絵)。
    # 前の「判定の画像」フォルダ(judge_images)は _saved で 1 回だけ枠へ読み替える
    PathItem("judge_ok", "合格の画像", "UFOK のフォームの絵", "",
             "合格のとき 3 秒大きく出す画像のファイル(jpg・jpeg・png・gif・bmp。名前は自由)。"
             "空・ファイルが無いときは同梱の VBA の絵(OK.jpg)", kind="file"),
    PathItem("judge_ng", "不合格の画像", "UFNG のフォームの絵", "",
             "不合格のとき 3 秒大きく出す画像のファイル。空・ファイルが無いときは同梱の VBA の絵(NG.jpg)", kind="file"),
    *(PathItem(f"judge_diag{n}", f"対角線の画像{n}", f"UF対角 の DB{n}", "",
               "対角線の欄に値が入ったとき 2 秒出す画像のファイル(5 枠から乱数で 1 枚。空の枠は同梱の絵。"
               f"1 つも決めなければ同梱の 対角1〜5)。空・ファイルが無いときは同梱の VBA の絵(対角{n}.jpg)",
               kind="file") for n in range(1, 6)),
    PathItem("history", "検査履歴", "HistoryPath",
             r"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\検査履歴" "\\",
             "ライン別ログ・Aインプット CSV・履歴のコピー", need="write"),
    PathItem("remote_stop", "管理者以外停止の置き場所", "(新規。all-tools と同じ管理者以外停止モード)", "",
             "管理者が「ほかの PC を止める」お願い(お願い.json)と、各 PC の稼働中の知らせを置くフォルダ。"
             "この下に 管理者以外停止\\ を作る。空なら 検査履歴(その下の 管理者以外停止\\)", need="write"),
    PathItem("factory", "工場共有 検査データ", "FactoryyPath", r"\\nlmsrvngy03\工場内共有\検査データ" "\\",
             "梱包検査ﾃﾞｰﾀ(測定DB)・PDF・キャスク/5N", need="write"),
    PathItem("nis", "測定データ電子化", "NisPath", r"C:\Tools\測定データ電子化",
             "hostsend.txt・SokuteiDataTrans.exe・result.txt(空なら使わない。VBA はオフライン時に空)", need="write"),
    PathItem("local_history", "手元の検査履歴", "C:\\検査履歴(直書き)", r"C:\検査履歴",
             "帳票 PDF・履歴ブック(xlsx)の手元保存", need="write"),
    PathItem("csv_local", "履歴 CSV の手元の出し先", "ThisWorkbook.Path(VBA は xlsm と同じフォルダ)", "",
             "印刷・保守の CSV 出力で、手元に書く 履歴 CSV(<L>_履歴CSV.csv)と Lot情報履歴 CSV のフォルダ。"
             "空ならこのツールの手元DB と同じフォルダ", need="write"),
    PathItem("lot_history_share", "Lot情報履歴の共有先", "(新規)", "",
             "印刷のたびに Lot情報履歴 を CSV で写す共有フォルダ。空なら 検査履歴\\<ライン>\\<ライン>_履歴",
             need="write"),
    PathItem("app_log", "Access.log などの置き場所", "ThisWorkbook.Path(VBA は xlsm と同じフォルダ)", "",
             "Access.log・ライン別ログの既定の置き場所(書き先を指定しないもの)と、hostsend.txt の予備の場所。"
             "空ならこの PC の手元(ローカル領域の logs)。共有フォルダにすると全 PC のログが 1 か所に集まる",
             need="write"),
    PathItem("nippou_out", "日報の受け渡し先", "(新規。all-tools の日報モードへ)", "",
             "印刷のたびに、その直の日報を all-tools の日報モードが読める CSV(集計明細_<日付>_<ライン>_<直>.csv)で"
             "書くフォルダ。日報モードの「CSV から日報を戻す」で取り込む。空なら Access.log などの置き場所",
             need="write"),
    PathItem("wav", "音声", "C:\\VBAWAV(直書き)", r"C:\VBAWAV" "\\",
             "判定音・案内の音声のフォルダ(何を鳴らすかは 設定 → 音。D29)"),
    PathItem("nps_mitsubishi", "三菱受託検査 履歴(NPS)", "直書き(Mitsubishi)",
             r"\\nlmsrvngy03\NPS\全社研修会トレーナー教育\第36回NPSトレーナー教育（名古屋）2017年\梱包実績の見える化\三菱受託検査_履歴" "\\",
             "三菱受託の xlsx(→PDF+CSV)の写し", need="write"),
    PathItem("access_master", "アクセス権限マスタ", "(python-web-tools と共通)", "",
             "「アクセス権限」表がある 梱包資材マスタ.sqlite3 のファイル、またはそのフォルダ(フォルダなら中の"
             " 梱包資材マスタ.sqlite3)。管理者(Administrator)とこの PC のラインを決める(D30)。"
             "空のあいだは設定の途中(管理者パスワードだけで設定できる)", kind="file"),
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
    user_settings.save_or_raise(KEY_MODE, m)   # 2026-11-08


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
    fnames = {k: str(v) for k, v in _d(raw.get(KEY_FIRST_NAMES)).items() if k in LOT_DBS}
    return {KEY_ITEMS: items, KEY_LOT: lot, KEY_SECOND_NAMES: names, KEY_FIRST_NAMES: fnames}


def _saved() -> dict[str, Any]:
    raw = _d(user_settings.get(KEY_PATHS, {}) or {})
    d = _norm(raw, mode())
    old = str(_d(raw.get(KEY_ITEMS)).get("judge_images") or "").strip()
    if "judge_images" in _d(raw.get(KEY_ITEMS)):
        # 2026-11-09: 前の「判定の画像」フォルダ → 枠ごとのファイル(1 回だけ。決めてある枠は変えない)
        from . import judge_images
        if old:
            for k, v in judge_images.migrate_from_folder(old).items():
                d[KEY_ITEMS].setdefault(k, v)
        try:
            user_settings.save_or_raise(KEY_PATHS, d)
            log.info("判定の画像のフォルダを、画像ごとのファイルに読み替えました: %s", old)
        except Exception as e:  # noqa: BLE001  保存できなくても読み替えた値で動く(次に読むときもう一度)
            log.warning("判定の画像の読み替えを保存できません: %s", e)
    elif any(k in raw for k in (ONLINE, OFFLINE, FIRST, SECOND)):
        user_settings.save(KEY_PATHS, d)            # 前の形を見つけたら、読み替えた形で保存し直す(1 回だけ)
        log.info("使用パスの保存の形を読み替えました(仕掛台帳だけ第1/第2、ほかはパス1つ)")
    return d


def app_dir() -> str:
    """本体のフォルダ(共有に置くと全 PC が同じ所を見る。**書き込みには使わない**。2026-11-07)。"""
    return str(config.BASE_DIR) + os.sep


def write_dir() -> str:
    """VBA で ThisWorkbook.Path(本体のフォルダ)に書いていたものの置き場所。

    使用パス「Access.log などの置き場所」、空ならこの PC の手元(ローカル領域の logs)。
    共有に置いた本体のフォルダへは書かない(全 PC が同じファイルを取り合う・読み取り専用だと書けないため)。
    """
    v = get("app_log")
    if v:
        return v
    from .. import app_config
    d = app_config.local_dir("logs")
    d.mkdir(parents=True, exist_ok=True)
    return str(d) + os.sep


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
#: 画像のフォルダ(2026-11-05 依頼で使用パスに出した)。決めていなければ VBA と同じく元のフォルダの下
DERIVED = {"pack_spec": ("quality", "包装仕様書"), "vc_images": ("packing", "VC種類画像"),
           "quality_photo": ("quality", "品質写真")}


def slot_default(key: str, slot: str = FIRST) -> str:
    """既定値。仕掛台帳は第1 / 第2、ほかは slot によらず 1 つ。画像のフォルダは元のフォルダの下。"""
    item = BY_KEY[_key(key)]
    if item.key in DERIVED:
        base, sub = DERIVED[item.key]
        return join(slot_value(base), sub) + "\\"
    return item.second if (item.key == LOT and slot == SECOND) else item.path


def _saved_value(key: str) -> Optional[str]:
    """その項目を利用者が決めているか(決めていなければ None)。"""
    v = _saved()[KEY_ITEMS].get(_key(key))
    return None if v is None else str(v)


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
    user_settings.save_or_raise(KEY_PATHS, d)   # 2026-11-08: 読めなければ保存しない(画面に出す)


def reset_path(key: str, slot: Optional[str] = None) -> None:
    k = _key(key)
    d = _saved()
    bucket, name = (d[KEY_LOT], _lot_slot(key, slot)) if k == LOT else (d[KEY_ITEMS], k)
    if name in bucket:
        del bucket[name]
        user_settings.save_or_raise(KEY_PATHS, d)   # 2026-11-08: 読めなければ保存しない(画面に出す)


def snapshot(key: str) -> dict[str, str]:
    """いまの値(ログに前後を残す形)。仕掛台帳は 第1・第2・2つ目の名前。"""
    k = _key(key)
    if k == LOT:
        out = {SLOT_LABELS[FIRST]: slot_value(LOT, FIRST), SLOT_LABELS[SECOND]: slot_value(LOT, SECOND)}
        out.update({f"第1の名前 {db}": first_name(db) for db in LOT_DBS})
        out.update({f"第2の名前 {db}": second_name(db) for db in LOT_DBS})
        return out
    return {BY_KEY[k].label: slot_value(k)}


def item_changed(key: str) -> bool:
    """その項目のどこかが既定から変わっているか(仕掛台帳は 第1・第2・2つ目の名前のどれか)。"""
    k = _key(key)
    d = _saved()
    if k == LOT:
        return bool(d[KEY_LOT]) or bool(d[KEY_SECOND_NAMES]) or bool(d[KEY_FIRST_NAMES])
    return k in d[KEY_ITEMS]


def reset_item(key: str) -> None:
    """項目を既定に戻す(設定「既定に戻す」)。仕掛台帳は 第1・第2・2つ目の名前をまとめて戻す。"""
    k = _key(key)
    if k not in BY_KEY:
        raise KeyError(key)
    d = _saved()
    if k == LOT:
        d[KEY_LOT].clear()
        d[KEY_SECOND_NAMES].clear(); d[KEY_FIRST_NAMES].clear()
    else:
        d[KEY_ITEMS].pop(k, None)
    user_settings.save_or_raise(KEY_PATHS, d)   # 2026-11-08: 読めなければ保存しない(画面に出す)


def reset_all() -> list[str]:
    """すべての項目を既定(VBA のオンラインの値)に戻す。戻した項目の key を返す。モードは変えない。"""
    done = [i.key for i in ITEMS if item_changed(i.key)]
    user_settings.save_or_raise(KEY_PATHS, {KEY_ITEMS: {}, KEY_LOT: {}, KEY_SECOND_NAMES: {}, KEY_FIRST_NAMES: {}})
    return done


def changed(key: str, slot: str = FIRST) -> bool:
    k = _key(key)
    d = _saved()
    return slot in d[KEY_LOT] if k == LOT else k in d[KEY_ITEMS]


# ------------------------------------------------------------------
# 仕掛台帳の 2つ目の名前
# ------------------------------------------------------------------
def default_name(db: str) -> str:
    return DB_FILES[db][1]


def first_name_setting(db: str) -> str:
    """設定「第1の名前」の入力そのまま(空 = 既定の名前)。"""
    return str(_saved().get(KEY_FIRST_NAMES, {}).get(db, "") or "")


def first_name(db: str) -> str:
    """第1で探すファイル名。設定が空なら既定(SIKALOT.sqlite3 など)。"""
    if db not in LOT_DBS:
        return default_name(db)
    v = first_name_setting(db).strip()
    if not v:
        return default_name(db)
    return v if "." in v else v + ".sqlite3"


def set_name(db: str, slot: str, name: str) -> None:
    """第1 / 第2 のファイル名を決める(空で既定に戻す)。"""
    if slot == SECOND:
        return set_second_name(db, name)
    if db not in LOT_DBS:
        raise KeyError(db)
    name = str(name or "").strip().strip('"')
    if any(c in name for c in "\\/:"):
        raise ValueError("ファイル名だけを入れてください(フォルダは入れません)")
    d = _saved()
    d.setdefault(KEY_FIRST_NAMES, {})
    if name:
        d[KEY_FIRST_NAMES][db] = name
    else:
        d[KEY_FIRST_NAMES].pop(db, None)
    user_settings.save_or_raise(KEY_PATHS, d)   # 2026-11-08: 読めなければ保存しない(画面に出す)


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
    user_settings.save_or_raise(KEY_PATHS, d)   # 2026-11-08: 読めなければ保存しない(画面に出す)


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
    out |= {second_name(db) for db in LOT_DBS} | {first_name(db) for db in LOT_DBS}
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
    p1 = join(slot_value(LOT, FIRST), first_name(name))
    p2 = join(slot_value(LOT, SECOND), second_name(name))
    if p1 == p2 or lot_usable(p1):
        return p1
    return p2 if lot_usable(p2) else p1


#: 「…がありません」で言う DB の呼び名(現場の言い方)
DB_LABELS = {"SIKALOTNOW": "仕掛台帳(ロット)", "SIKAHIKINOW": "仕掛台帳(引当)", "SIKAODRNOW": "仕掛台帳(受注)",
             "SIKACOMP": "仕掛台帳(完了)", "SIKATWOYEAR": "仕掛台帳(完了2年)",
             "品質基準": "品質基準", "伝送用ファイル": "伝送用ファイル"}


def looked_files(name: str, m: Optional[str] = None) -> list[str]:
    """`name`(SIKALOTNOW など)を探した実際のファイル(探す順)。仕掛の 3 ファイルはオンラインが第1 → 第2、
    オフラインは第2だけ(第2は「2つ目の名前」)。"""
    name = name[:-6] if name.endswith(".accdb") else name
    if name not in DB_FILES:
        return []
    where, file = DB_FILES[name]
    if name in LOT_DBS:
        p2 = join(slot_value(LOT, SECOND), second_name(name))
        if lot_mode(m) == OFFLINE:
            return [p2]
        p1 = join(slot_value(LOT, FIRST), first_name(name))
        return [p1] if p1 == p2 else [p1, p2]
    return [db_file(name, m)]


def missing_text(name: str, files: Optional[list[str]] = None, m: Optional[str] = None) -> str:
    """「ファイルが見つからない(読めない)」の知らせ。VBA の「日稲:SIKALOTNOW.accdbがありません_ReadCheck_…」の代わりに、
    実際に探したファイル名とフォルダを言う(このアプリは .accdb を読まない)。"""
    name = name[:-6] if name.endswith(".accdb") else name
    files = [f for f in (files if files is not None else looked_files(name, m)) if f]
    label = DB_LABELS.get(name, name)
    if not files:
        return f"{label} のファイルの場所が設定されていません。設定 → 使用パス で決めてください。"
    parts = []
    for f in files:
        cut = max(f.rfind("\\"), f.rfind("/"))
        folder, fname = f[:cut].rstrip("\\/"), f[cut + 1:]
        parts.append(f"{fname}(フォルダ {folder or '(空)'})")
    mode_note = ""
    if name in LOT_DBS:
        mode_note = "オフラインなので第2パス(控え)だけを見ました。第1パス(台帳)を読むには 設定 → この端末 のモードを「オンライン」にするか、第2パスの名前を合わせてください。" if lot_mode(m) == OFFLINE else ""
    what = "見つからないか、手元へ写せません" if name in COPIED_DBS else "見つかりません"
    return (f"{label} のファイルが{what}: " + " / ".join(parts) + "。"
            + mode_note + "設定 → 使用パス の「確認」で場所と名前を確かめてください。")


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
def check(key: str, slot: Optional[str] = None, *, text: Optional[str] = None,
          names: Optional[dict[str, str]] = None,
          first_names: Optional[dict[str, str]] = None) -> dict[str, Any]:
    """確認ボタン: ある・読める・書けるか。仕掛台帳はその側(第1 / 第2)と、各ファイルを手元へ写せるか。

    `text` / `names` は画面で入力中(保存前)の値。渡せば保存した値の代わりにそれを確かめる
    (names は 2つ目の名前 {DB名: 名前}。空は1つ目と同じ名前)。第2のファイルは足りない列も見る。
    """
    k = _key(key)
    item = BY_KEY[k]
    s = _lot_slot(key, slot) if k == LOT else ""
    text = slot_value(k, s or FIRST) if text is None else str(text).strip().strip('"')
    out: dict[str, Any] = {"key": k, "slot": s, "label": SLOT_LABELS.get(s, ""), "path": text, "exists": False,
                           "readable": False, "writable": False, "ok": False, "message": "", "files": []}
    if not text:
        out["message"] = "設定されていません"
        if k.startswith("judge_"):
            out["ok"] = True
            out["message"] = "空なので同梱の VBA の絵を出します"   # 2026-11-09
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
            if s == SECOND and names is not None and db in LOT_DBS:
                fname = _name_of(names.get(db, ""), db)
            elif s == FIRST and first_names is not None and db in LOT_DBS:
                v = str(first_names.get(db, "") or "").strip().strip('"')
                fname = (v if "." in v else v + ".sqlite3") if v else default_name(db)
            else:
                fname = second_name(db) if s == SECOND else first_name(db)
            there = exists(join(text, fname), "file")
            ok = there and lot_usable(join(text, fname))
            lack = _lacking(join(text, fname), db) if ok and db in LOT_DBS else []
            out["files"].append({"db": db, "name": fname, "exists": there, "copied": ok, "lacking": lack})
            have.append(f"{fname} " + ("あり" if ok else ("写せません" if there else "なし"))
                        + (f"(足りない列 {len(lack)}: {_short(lack)})" if lack else ""))
        out["message"] += "(" + "・".join(have) + ")"
    if k.startswith("judge_"):
        # 2026-11-09: 判定の画像。空・無いときは同梱の絵を出す(ここで知らせる)
        from . import judge_images
        if k in judge_images.SLOTS:
            if not out["exists"]:
                out["message"] = judge_images.MISSING_NOTE
            elif Path(str(p)).suffix.lower() not in judge_images.EXT:
                out["ok"] = False
                out["message"] = "画像ではありません(jpg・jpeg・png・gif・bmp)。同梱の VBA の絵を出します"
    return out


def _name_of(value: str, db: str) -> str:
    """入力中の「2つ目の名前」→ 探すファイル名(second_name と同じ決まり)。"""
    v = str(value or "").strip().strip('"')
    if not v:
        return first_name(db)
    return v if "." in v else v + ".sqlite3"


def _lacking(path_text: str, db: str) -> list[str]:
    """そのファイルの表「仕掛」に足りない列(列名の書き方の違いは直してから見る。python-web-tools「足りない列」)。"""
    from . import datasource, lot_columns
    try:
        cols, _rows = datasource.get_table(path_text, "仕掛")
    except Exception:                                   # noqa: BLE001 - 読めないことは別に出る
        return []
    return lot_columns.missing(db, cols) if cols else []


def _short(cols: list[str], n: int = 8) -> str:
    return "、".join(cols[:n]) + (f" ほか {len(cols) - n} 列" if len(cols) > n else "")


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
                "first_names": [{"db": db, "default": default_name(db), "value": first_name_setting(db),
                                 "effective": first_name(db)} for db in LOT_DBS],
            })
        else:
            row.update({"path": slot_value(i.key), "default": slot_default(i.key), "changed": changed(i.key)})
            if i.key.startswith("judge_"):
                from . import judge_images          # 2026-11-09: 見本(小さな画像)と「ファイルがありません」
                if i.key in judge_images.SLOTS:
                    row["image"] = judge_images.slot_info(i.key)
        rows.append(row)
    return rows
