"""アクセス権限 ── 梱包資材マスタの「アクセス権限」表から、この PC の「管理者か」と「ライン」を決める

all-tools 日報ツール(`tools/nippou/nippou/logic/access_rights.py` v4.12.0〜v4.12.3)と同じ表・同じ読み方
(判断事項 D12・D30)。2026-11-03 のご依頼:

    梱包資材マスタ.sqlite3 のアクセス権限を読む。アクセス権限でログインID・PC名がヒットし、権限: Administrator。
    権限にはライン名も入る。参考: all-tools

【表(梱包資材マスタ.sqlite3 の「アクセス権限」。ほかのツールと共有)】

    管理番号 / ログインID / PC名 / 権限 / 有効 / 備考

1 行が「このログインIDで、この PC に入ったときの権限」。現場のライン PC は共通のログインIDで使い、PC ごとに行がある。

【どの行がこの PC か ── ログインIDとPC名の両方】
* ログインID と PC名 が**両方とも**この PC と同じ行(欄が空ならその欄は問わない。**両方空の行は誰の行でもない**)
* 大文字小文字・全角半角は問わない。`DOMAIN\\user` と書いてあれば `user` でも当たる
* 「有効」が 0 / いいえ / 無効 / false / × の行は読まない(空は有効)

【権限の読み方】1 つの欄に「,」「、」「/」「・」「;」で並べてあれば 1 つずつ読む(空白では区切らない)

    Administrator              … 管理者(このツールの ADMIN。VBA の隠しパスワード weekend99 の代わり)
    ufkensa:admin              … Administrator と同じ(前の書き方。互換のため読む)
    ufkensa:calibration_skip   … 校正をとばせる(VBA の隠しパスワード jump の代わり。このツールだけの値)
    ライン名(正規の呼び名だけ) … この PC のライン = 設定「この端末」の設備名を自動で決める
                                 L-1・LVC・HVC・トット・バランサー → 同じ設備名、中板 → 丸徳。
                                 中板は号機を続けて書く(`中板A3` → 丸徳・号機 A3。W14: 号機は A1〜A11)。
                                 `中板` だけ・`中板3`(日報ツールのラインNO の書き方)は 丸徳 にするが号機は決めない
    それ以外                   … 読まない(mode:field・mode:fullaccess・作業長・機側・NS1・AIM・コイル など、
                                 ほかのツールの権限やこのツールに無いライン)

ライン名は**文字の幅だけ**揃える(D13。`ﾄｯﾄ` = `トット`)。大文字小文字・記号・空白は揃えない ── `L1`・`l-1`・`丸徳`
は読まず、設定の画面に「正規は…」を添えて出す(`inspect`)。

【ラインを決めないとき】この PC の行に**違うラインが 2 つ以上**(号機違いも)あれば決めず、理由を出す。
ラインを書いていなければ触らない。表のラインは**前に当てたものから変わったときだけ**当てる(`should_apply`):
管理者パスワードで変えたライン(「パスワードがあれば現行の通り変更可」)が、起動のたびに戻らないように。

【表がどこにあるか】設定 → 使用パス の「アクセス権限マスタ」。**フォルダでもファイルでもよい**
(フォルダなら その中の `梱包資材マスタ.sqlite3`。all-tools の「梱包資材マスタのフォルダ + ファイル名」と同じ)。

【設定の途中(setup)】次のときは「マスタからは決まらない」ので、**管理者パスワードだけで**使用パスなどを直せる
(`auth.setup_mode`)。権限が使用パスにある表で決まるので、そうしないと新しい PC・試験用の PC で詰む:
    * アクセス権限マスタが未設定(空)       … 前と同じく、誰でも管理者として動く(入れたばかりの PC)
    * 設定してあるが読めない(前回の結果も無い)… 一般。管理者パスワードで設定できる
    * 読めたが、表に管理者の行が 1 つも無い  … 一般。管理者パスワードで設定できる(誰も直せなくならないように)

【届かないとき】前回読めた結果(同じ表・同じ人・同じ PC)をこの PC に覚えておき、それで決める。
"""
from __future__ import annotations

import json
import re
import sqlite3
import time
import unicodedata
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Iterable, Mapping, Optional

from . import app_config
from .logging_utils import get_logger

log = get_logger("access_rights")

TABLE = "アクセス権限"
DEFAULT_FILE = "梱包資材マスタ.sqlite3"
MASTER_SUFFIXES = (".sqlite3", ".sqlite", ".db")

ADMINISTRATOR = "Administrator"
ADMIN_ALIAS = "ufkensa:admin"                 # 前の書き方(互換)
CALIBRATION_SKIP = "ufkensa:calibration_skip"
ADMIN = ADMINISTRATOR
CODES = {ADMINISTRATOR: "管理者(設定・使用パス・取り込み。ufkensa:admin も同じ)",
         CALIBRATION_SKIP: "校正をとばせる(このツールだけの値)"}

COL_NO, COL_LOGIN, COL_PC, COL_PERMISSION, COL_ENABLED, COL_NOTE = (
    "管理番号", "ログインID", "PC名", "権限", "有効", "備考")
COLUMNS = (COL_NO, COL_LOGIN, COL_PC, COL_PERMISSION, COL_ENABLED, COL_NOTE)
_OFF = {"0", "false", "no", "off", "×", "✕", "x", "無効", "いいえ", "n"}
_SEPARATORS = re.compile(r"[,、/／・;；]+")
RECHECK_SEC = 300.0

SOURCE_UNSET, SOURCE_MASTER, SOURCE_CACHE, SOURCE_UNREACHABLE = "unset", "master", "cache", "unreachable"
LABEL_EDIT, LABEL_VIEW = "管理者", "一般"

#: マスタのライン名(正規の呼び名)→ この端末の設備名(terminal.FACILITIES)。中板は号機を続けて書く
MARU = "中板"
LINE_FACILITY = {"L-1": "L-1", "LVC": "LVC", "HVC": "HVC", "トット": "トット", "バランサー": "バランサー", MARU: "丸徳"}
MARU_NUMBERS = tuple(f"A{i}" for i in range(1, 12))          # W14: VBA どおり A1〜A11
#: 正規でない書き方 → 正規(読まずに「正規は…」と添えるだけ)
_HINTS = {"L1": "L-1", "TOT": "トット", "BALA": "バランサー", "MARU": MARU, "丸徳": MARU}
_MARU_FORM = re.compile(r"^中板(?P<no>[A-Za-z]?\d+)$")

MSG_SETUP = "アクセス権限マスタが未設定のため、管理者パスワードだけで設定できます"
MSG_SETUP_UNREAD = "アクセス権限マスタを読めないため、管理者パスワードだけで設定できます"
MSG_SETUP_NO_ADMIN = "アクセス権限マスタに管理者(Administrator)の行が1つも無いため、管理者パスワードだけで設定できます"


# ------------------------------------------------------------------
# 比べる形
# ------------------------------------------------------------------
def fold(text: Any) -> str:
    """ログインID・PC名・権限の比べる形(全角半角・大文字小文字・前後の空白を揃える)。"""
    return unicodedata.normalize("NFKC", "" if text is None else str(text)).strip().casefold()


def width(text: Any) -> str:
    """ライン名の比べる形(D13: 文字の幅と前後の空白だけ)。"""
    return unicodedata.normalize("NFKC", "" if text is None else str(text)).strip()


def enabled(value: Any) -> bool:
    return fold(value) not in _OFF


def tokens(value: Any) -> list[str]:
    """1 つの欄に並んだ権限を 1 つずつ。"""
    return [t.strip() for t in _SEPARATORS.split(str(value or "")) if t.strip()]


def is_admin_token(token: Any) -> bool:
    return fold(token) in (fold(ADMINISTRATOR), fold(ADMIN_ALIAS))


def _col(row: Mapping[str, Any], name: str) -> Any:
    if name in row:
        return row.get(name)
    want = fold(name)
    for k, v in row.items():
        if fold(k) == want:
            return v
    return None


def _same(cell: Any, mine: str) -> bool:
    text, me = fold(cell), fold(mine)
    if not text or not me:
        return False
    me_short = me.rsplit("\\", 1)[-1]
    return text in (me, me_short) or text.rsplit("\\", 1)[-1] in (me, me_short)


def applies(row: Mapping[str, Any], login_id: str, pc: str) -> bool:
    """その行がこの人・この PC に当てはまるか。空欄は問わない。両方空の行は誰にも効かない。"""
    rid, rpc = str(_col(row, COL_LOGIN) or "").strip(), str(_col(row, COL_PC) or "").strip()
    if not rid and not rpc:
        return False
    return (not rid or _same(rid, login_id)) and (not rpc or _same(rpc, pc))


@dataclass(frozen=True)
class LineReading:
    """権限の値 1 つをラインとして読んだ結果。"""
    facility: str = ""        # 設備名(読めなければ空)
    number: str = ""          # 丸徳の号機(A1〜A11。決まらなければ空)
    note: str = ""            # 読めたが号機は決めない、など(画面に出す)
    problem: str = ""         # 書き間違い(点検で言う)


def read_line(token: Any) -> LineReading:
    """権限の値 → この端末のライン。正規の呼び名(幅だけ揃える)だけを読む。"""
    s = width(token)
    if s in LINE_FACILITY and s != MARU:
        return LineReading(LINE_FACILITY[s])
    if s == MARU:
        return LineReading("丸徳", note="「中板」に号機が無いので、号機は決めません(中板A3 のように続けて書きます)")
    m = _MARU_FORM.match(s)
    if not m:
        return LineReading()
    no = m.group("no")
    if no.isdigit():
        return LineReading("丸徳", note=f"「{s}」は日報ツールのラインNO の書き方なので、号機(A1〜A11)は決めません"
                                        f"(中板A{no} と書けば号機も決めます)")
    if no in MARU_NUMBERS:
        return LineReading("丸徳", no)
    return LineReading("丸徳", problem=f"「{str(token).strip()}」の号機は A1〜A11 ではありません")


def hint(token: Any) -> str:
    """ラインのつもりで書いたのに読めない値なら、その一言(ラインでない値は空)。"""
    s = width(token)
    if not s:
        return ""
    reading = read_line(s)
    if reading.facility:
        return reading.problem
    loose = re.sub(r"[\W_]+", "", s).upper()
    for official in LINE_FACILITY:
        if loose == re.sub(r"[\W_]+", "", official).upper():
            return f"「{str(token).strip()}」は読みません(正規は「{official}」)"
    if s in _HINTS or s.upper() in _HINTS:
        return f"「{str(token).strip()}」は読みません(正規は「{_HINTS.get(s) or _HINTS[s.upper()]}」)"
    if s.startswith(MARU):
        return f"「{str(token).strip()}」は読みません(中板は「中板A3」のように号機を続けて書きます)"
    return ""


# ------------------------------------------------------------------
# 判断(純粋)
# ------------------------------------------------------------------
@dataclass
class Decision:
    edit: bool
    source: str                         # unset | master | cache | unreachable
    reason: str
    login_id: str = ""
    pc: str = ""
    path: str = ""
    matched: list[dict[str, Any]] = field(default_factory=list)   # この PC に当てはまった行
    checked_at: float = 0.0
    codes: list[str] = field(default_factory=list)                # 読めた権限(Administrator / ufkensa:calibration_skip)
    facility: str = ""                  # 決まった設備名(決まらなければ空)
    line_no: str = ""                   # 丸徳の号機
    lines: list[str] = field(default_factory=list)               # この PC の行に書いてあったライン
    ignored: list[str] = field(default_factory=list)             # 読まなかった権限の値
    line_problem: str = ""              # ラインを決めなかった理由
    line_notes: list[str] = field(default_factory=list)
    has_admin_row: bool = True          # 表に有効な管理者の行があるか(無ければ設定の途中)

    @property
    def setup(self) -> bool:
        """マスタから決まらない(未設定・読めない・管理者の行が無い)= 管理者パスワードだけで設定できる。"""
        return self.source in (SOURCE_UNSET, SOURCE_UNREACHABLE) or (
            self.source == SOURCE_MASTER and not self.has_admin_row)

    @property
    def setup_reason(self) -> str:
        if self.source == SOURCE_UNSET:
            return MSG_SETUP
        if self.source == SOURCE_UNREACHABLE:
            return MSG_SETUP_UNREAD
        if self.setup:
            return MSG_SETUP_NO_ADMIN
        return ""

    def has(self, code: str) -> bool:
        if code in (ADMIN, ADMIN_ALIAS):
            return self.edit
        return code in self.codes or self.source == SOURCE_UNSET

    @property
    def label(self) -> str:
        return LABEL_EDIT if self.edit else LABEL_VIEW

    @property
    def line_key(self) -> str:
        """前に当てたものと比べる形(`丸徳:A3` / `L-1`)。決まっていなければ空。"""
        if not self.facility:
            return ""
        return f"{self.facility}:{self.line_no}" if self.line_no else self.facility

    def line_text(self) -> str:
        if not self.facility:
            return ""
        return f"{self.facility}{self.line_no}" if self.line_no else self.facility

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d.update(label=self.label, setup=self.setup, setup_reason=self.setup_reason, line_text=self.line_text())
        return d


@dataclass(frozen=True)
class Reading:
    """表の行とこの PC から読めたこと(`decide`)。"""
    admin: bool
    calibration_skip: bool
    matched: list[dict[str, Any]]
    facility: str = ""
    line_no: str = ""
    lines: tuple[str, ...] = ()
    ignored: tuple[str, ...] = ()
    problem: str = ""
    notes: tuple[str, ...] = ()
    has_admin_row: bool = True


def _row_view(number: int, row: Mapping[str, Any]) -> dict[str, Any]:
    d = {k: _col(row, k) for k in COLUMNS}
    d["行"] = number
    return d


def decide(rows: Iterable[Mapping[str, Any]], login_id: str, pc: str) -> Reading:
    """表の行とこの PC の名前から決める(all-tools `logic.access_rights.decide` と同じ考え方)。"""
    rows = list(rows)
    matched: list[dict[str, Any]] = []
    admin = skip = False
    found: list[tuple[str, str]] = []
    ignored: list[str] = []
    notes: list[str] = []
    has_admin_row = any(enabled(_col(r, COL_ENABLED)) and (str(_col(r, COL_LOGIN) or "").strip()
                                                           or str(_col(r, COL_PC) or "").strip())
                        and any(is_admin_token(t) for t in tokens(_col(r, COL_PERMISSION))) for r in rows)
    for number, row in enumerate(rows, 1):
        if not enabled(_col(row, COL_ENABLED)) or not applies(row, login_id, pc):
            continue
        matched.append(_row_view(number, row))
        for token in tokens(_col(row, COL_PERMISSION)):
            if is_admin_token(token):
                admin = True
                continue
            if fold(token) == fold(CALIBRATION_SKIP):
                skip = True
                continue
            reading = read_line(token)
            if reading.facility and not reading.problem:
                if reading.note and reading.note not in notes:
                    notes.append(reading.note)
                if (reading.facility, reading.number) not in found:
                    found.append((reading.facility, reading.number))
                continue
            why = reading.problem or hint(token)
            label = f"{token}({why})" if why else token
            if label not in ignored:
                ignored.append(label)
    # 号機の無い 丸徳 は、号機つきの 丸徳 があればそちらにまとめる
    numbered = {f for f, n in found if n}
    found = [(f, n) for f, n in found if n or f not in numbered]
    shown = tuple(f"{f}{n}" for f, n in found)
    facility, number, problem = "", "", ""
    if len(found) == 1:
        facility, number = found[0]
    elif len(found) > 1:
        problem = (f"この PC の行にラインが{len(found)}つあります({'・'.join(shown)})。"
                   "どちらか分からないので、ラインは自動では決めません")
    return Reading(admin, skip, matched, facility, number, shown, tuple(ignored), problem, tuple(notes),
                   has_admin_row)


@dataclass(frozen=True)
class Finding:
    """点検で気になった 1 行。"""
    number: int
    key: str
    message: str

    def describe(self) -> str:
        where = f"管理番号 {self.key}" if self.key else f"{self.number}行目"
        return f"{where}: {self.message}"


def inspect(rows: Iterable[Mapping[str, Any]]) -> list[Finding]:
    """表の**全部の行**を確かめる(この PC の行に限らない)。ラインでない値(mode:field など)は言わない。"""
    rows = list(rows)
    found: list[Finding] = []
    by_pair: dict[tuple[str, str], list[tuple[int, str, str]]] = {}
    for number, row in enumerate(rows, 1):
        key = str(_col(row, COL_NO) or "").strip()
        perm = _col(row, COL_PERMISSION)
        for token in tokens(perm):
            if is_admin_token(token) or fold(token) == fold(CALIBRATION_SKIP):
                continue
            why = hint(token)
            if why:
                found.append(Finding(number, key, f"権限: {why}"))
        login, pc = str(_col(row, COL_LOGIN) or "").strip(), str(_col(row, COL_PC) or "").strip()
        if not login and not pc:
            found.append(Finding(number, key, "ログインIDもPC名も空なので、どのPCの行にもなりません"))
            continue
        if not enabled(_col(row, COL_ENABLED)):
            continue
        for token in tokens(perm):
            r = read_line(token)
            if r.facility and not r.problem:
                by_pair.setdefault((fold(login), fold(pc)), []).append((number, key, f"{r.facility}{r.number}"))
    for hits in by_pair.values():
        distinct = sorted({h[2] for h in hits})
        if len(distinct) > 1:
            number, key, _ = hits[0]
            where = "・".join((f"管理番号 {k}" if k else f"{n}行目") for n, k, _ in hits)
            found.append(Finding(number, key, (f"同じログインID・PC名にラインが{len(distinct)}つあります"
                                               f"({'・'.join(distinct)}、{where})── どちらか分からないので決めません")))
    found.sort(key=lambda f: f.number)
    return found


def should_apply(decided_key: str, last_applied: str) -> bool:
    """表のラインを、いまこの端末に当てるか。**前に当てたものから変わったときだけ**(all-tools と同じ)。"""
    return bool(decided_key) and decided_key != (last_applied or "")


# ------------------------------------------------------------------
# どのファイルか
# ------------------------------------------------------------------
def configured_text() -> str:
    from .kensa import paths
    return str(paths.get("access_master") or "").strip().strip('"')


def resolve(text: str) -> Path:
    """使用パスの文字 → 表のファイル。フォルダなら その中の 梱包資材マスタ.sqlite3。"""
    from .kensa import paths
    p = paths.to_local(text)
    if p.suffix.lower() in MASTER_SUFFIXES and not p.is_dir():
        return p
    return p / DEFAULT_FILE


def master_path() -> Optional[Path]:
    """読む表のファイル。使用パス「アクセス権限マスタ」が空なら None(未設定)。"""
    text = configured_text()
    return resolve(text) if text else None


def read_rows(path: Path) -> list[dict[str, Any]]:
    """`アクセス権限` の行(読み取り専用で開く)。表が無い・読めないときは例外。"""
    if not path.is_file():
        raise FileNotFoundError(f"見つかりません: {path}")
    from .kensa.datasource import sqlite_uri
    uri = sqlite_uri(path, "ro")
    conn = sqlite3.connect(uri, uri=True, timeout=3)
    try:
        conn.row_factory = sqlite3.Row
        names = [r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")]
        table = next((n for n in names if fold(n) == fold(TABLE)), None)
        if table is None:
            raise LookupError(f"{path.name} に「{TABLE}」の表がありません")
        return [dict(r) for r in conn.execute(f'SELECT * FROM "{table}"')]
    finally:
        conn.close()


# ------------------------------------------------------------------
# 前回読めた結果(届かないとき用)
# ------------------------------------------------------------------
def _cache_file() -> Path:
    return app_config.local_dir("runtime") / "access_rights.json"


def _save_cache(d: Decision) -> None:
    try:
        path = _cache_file()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(asdict(d), ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def _load_cache(path: str, login_id: str, pc: str) -> Optional[dict[str, Any]]:
    try:
        d = json.loads(_cache_file().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    if not isinstance(d, dict) or d.get("path") != path or fold(d.get("login_id")) != fold(login_id) \
            or fold(d.get("pc")) != fold(pc):
        return None
    return d


# ------------------------------------------------------------------
def evaluate(login_id: str, pc: str) -> Decision:
    """この人・この PC の権限とライン。"""
    now = time.time()
    path = master_path()
    if path is None:
        return Decision(True, SOURCE_UNSET,
                        "使用パス「アクセス権限マスタ」が未設定なので、誰でも管理者として動きます。"
                        f"{MSG_SETUP}(梱包資材マスタ.sqlite3 のファイルか、そのフォルダを入れてください)。",
                        login_id, pc, "", [], now, sorted(CODES))
    try:
        rows = read_rows(path)
    except Exception as exc:                             # noqa: BLE001 - 届かない・壊れている
        cached = _load_cache(str(path), login_id, pc)
        if cached is not None:
            when = time.strftime("%m/%d %H:%M", time.localtime(float(cached.get("checked_at") or 0)))
            log.warning("アクセス権限マスタを読めないので前回の結果(%s)を使います: %s", when, exc)
            known = {f.name for f in Decision.__dataclass_fields__.values()}
            d = Decision(**{k: v for k, v in cached.items() if k in known})
            d.source, d.checked_at = SOURCE_CACHE, now
            d.reason = f"アクセス権限マスタに届かないので、前回({when})読めた結果を使っています: {exc}"
            return d
        log.warning("アクセス権限マスタを読めません(前回の結果も無いので一般): %s", exc)
        return Decision(False, SOURCE_UNREACHABLE,
                        f"アクセス権限マスタを読めないので一般にしています: {exc}。{MSG_SETUP_UNREAD}。",
                        login_id, pc, str(path), [], now)
    r = decide(rows, login_id, pc)
    codes = ([ADMINISTRATOR] if r.admin else []) + ([CALIBRATION_SKIP] if r.calibration_skip else [])
    if not r.matched:
        reason = f"アクセス権限マスタに、この PC(ログインID {login_id} / PC名 {pc})の行がありません。"
    else:
        reason = f"アクセス権限マスタの {len(r.matched)} 行が当てはまりました(" + (
            "・".join(codes + list(r.lines)) or "このツールの権限なし") + ")。"
    if not r.has_admin_row:
        reason += MSG_SETUP_NO_ADMIN + "。"
    d = Decision(r.admin, SOURCE_MASTER, reason, login_id, pc, str(path), r.matched, now, codes,
                 r.facility, r.line_no, list(r.lines), list(r.ignored), r.problem, list(r.notes), r.has_admin_row)
    _save_cache(d)
    return d


def overview() -> dict[str, Any]:
    """設定の画面: どのファイルを読んでいるか・読めるか・管理者の行・表ぜんぶの点検。"""
    text = configured_text()
    path = master_path()
    out: dict[str, Any] = {"configured": bool(text), "text": text, "path": str(path) if path else "",
                           "default_file": DEFAULT_FILE, "table": TABLE, "found": False, "rows": 0,
                           "admin_rows": [], "findings": [], "problem": ""}
    if path is None:
        out["problem"] = "使用パス「アクセス権限マスタ」が未設定です"
        return out
    try:
        rows = read_rows(path)
    except Exception as exc:                             # noqa: BLE001
        out["problem"] = str(exc)
        return out
    out["found"] = True
    out["rows"] = len(rows)
    out["admin_rows"] = [_row_view(n, r) for n, r in enumerate(rows, 1)
                         if any(is_admin_token(t) for t in tokens(_col(r, COL_PERMISSION)))]
    out["findings"] = [f.describe() for f in inspect(rows)]
    return out


# ------------------------------------------------------------------
# この端末に当てる
# ------------------------------------------------------------------
KEY_LINE_APPLIED = "kensa_access_line_applied"


def apply_line(d: Decision) -> str:
    """表で決まったラインを この端末の設備名 にする(前に当てたものから変わったときだけ)。知らせを返す。

    マスタを読めたとき(source = master)だけ当てる。号機が決まらない 丸徳 は、いまの号機を残す。
    """
    from . import user_settings
    from .kensa import terminal
    if d.source != SOURCE_MASTER or not d.facility:
        return ""
    last = str(user_settings.get(KEY_LINE_APPLIED, "") or "")
    if not should_apply(d.line_key, last):
        return ""
    number = d.line_no
    if not number and d.facility == "丸徳" and terminal.facility() == "丸徳":
        number = terminal.line_no()
    before = terminal.line_label() or "(未設定)"
    try:
        terminal.set_facility(d.facility, number)
    except ValueError as exc:
        log.warning("アクセス権限のラインを当てられません: %s", exc)
        return ""
    user_settings.save(KEY_LINE_APPLIED, d.line_key)
    msg = f"アクセス権限マスタから、この端末のラインを {before} → {terminal.line_label()} にしました"
    log.info("%s", msg)
    return msg
