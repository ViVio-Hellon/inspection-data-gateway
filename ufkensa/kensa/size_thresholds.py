"""寸法のしきい値のマスタ(2026-11-11 のご依頼)。伝送用ファイル.sqlite3 の表「寸法しきい値」。

VBA はよく変わる製品仕様のしきい値をコードに直書きしていた。ここでは共有の表で持ち、
保守 → 寸法のしきい値 で直せるようにする(全 PC で同じになるよう 伝送用ファイル に置く。配布の対象ではない)。

    列: キー・値・単位・説明

**既定は VBA と同じ値**(境界の隙間 6.0〜6.001 などもそのまま)。それ以外のしきい値は直書きのまま。
- 検査中に読むときは **書かない**(表が無ければ既定)。表は保守で開いたとき・保存したときに作る
- 読めないときは最後に読めた値(無ければ既定)。画面は止めない。悪い値のキーは既定
- 何度も読むので、読めた値を数秒だけ覚える(保存・「読み込む」で捨てる)
"""
from __future__ import annotations

import sqlite3
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

from .. import safe_io
from ..logging_utils import get_logger
from . import cache_registry, datasource

log = get_logger("kensa.size_thresholds")

TABLE = "寸法しきい値"
DB = "伝送用ファイル"
COLUMNS = ("キー", "値", "単位", "説明")
CHANGE_EV = "thresholds.change"
TTL_SEC = 5.0


@dataclass(frozen=True)
class Item:
    key: str
    group: str
    label: str
    default: float
    unit: str
    lo: float
    hi: float
    integer: bool = False
    note: str = ""


#: グループ名 → 説明(画面の見出し)
GROUPS: tuple[tuple[str, str], ...] = (
    ("op", "作業区分の厚の境界(オーダー厚で作業選択ボタンを隠す。FormRelat Atu_OPselect)"),
    ("thin", "薄板の厚(中耳歪・定盤歪の形状・薄板判定)"),
    ("sakko", "先行データ票"),
    ("kennyu", "薄板の検入枚数"),
    ("instr", "測定機器"),
)

ITEMS: tuple[Item, ...] = (
    Item("op_thin_lo", "op", "薄板・少枚数の範囲 下限(厚板・少枚数を隠す)", 1.0, "mm", 0, 10000),
    Item("op_thin_hi", "op", "薄板・少枚数の範囲 上限", 6.0, "mm", 0, 10000),
    Item("op_notthin_lo", "op", "薄板不可 下限", 6.001, "mm", 0, 10000,
         note="VBA は 6.0 と 6.001 の間に隙間がある(そのまま)"),
    Item("op_notthin_hi", "op", "薄板不可 上限", 800.0, "mm", 0, 10000),
    Item("op_heavy_lo", "op", "厚板不可(重い) 下限", 200.1, "mm", 0, 10000),
    Item("op_heavy_hi", "op", "厚板不可(重い) 上限", 299.0, "mm", 0, 10000),
    Item("op_heavy_weight", "op", "厚板不可(重い) 重量がこれを超える", 1500.0, "g", 0, 1e9),
    Item("op_notthick_lo", "op", "厚板不可 下限", 300.0, "mm", 0, 10000),
    Item("op_notthick_hi", "op", "厚板不可 上限", 800.0, "mm", 0, 10000),
    Item("op_notgoku_lo", "op", "極厚不可 下限", 1.0, "mm", 0, 10000),
    Item("op_notgoku_hi", "op", "極厚不可 上限", 29.9, "mm", 0, 10000),
    Item("op_notslab_lo", "op", "スラブ不可 下限", 1.0, "mm", 0, 10000),
    Item("op_notslab_hi", "op", "スラブ不可 上限", 199.9, "mm", 0, 10000),
    Item("thin_atu", "thin", "薄板とみなす厚(これ以下)", 6.0, "mm", 0, 10000,
         note="中歪/耳歪 4 級表・UF品質基準の中耳表示・定盤歪の形状(耳/中)・薄板判定(厚2/4ヵ所)"),
    Item("sakko_thick", "sakko", "製造板厚がこれ以下なら先行データ票", 3.0, "mm", 0, 10000),
    Item("kennyu_n", "kennyu", "検入枚数の境(これ以下 = 2 回検査、超え = 3 回検査)", 25, "枚", 1, 9999, integer=True,
         note="ボタンの文字「検入≦N」「検入≧N+1」もこの値から作る"),
    Item("conb_take", "instr", "L-1 のコンベックス C11/C12 を分ける丈(超え = C11)", 5500.0, "mm", 0, 100000),
    Item("depth_conb", "instr", "デプス/コンベックスを分ける高さ(以下 = デプス)", 300.0, "mm", 0, 100000),
)
BY_KEY = {i.key: i for i in ITEMS}
DEFAULTS: dict[str, float] = {i.key: i.default for i in ITEMS}
#: 下限 ≦ 上限 でなければならない組
PAIRS = (("op_thin_lo", "op_thin_hi"), ("op_notthin_lo", "op_notthin_hi"), ("op_heavy_lo", "op_heavy_hi"),
         ("op_notthick_lo", "op_notthick_hi"), ("op_notgoku_lo", "op_notgoku_hi"), ("op_notslab_lo", "op_notslab_hi"))

_lock = threading.Lock()
#: パス → (読んだ時刻, 値)。読めないときはこれを出す(最後に読めた値)
_last_known: dict[str, tuple[float, dict[str, float]]] = {}
cache_registry.register_dict("size_thresholds._last_known", _last_known, _lock)


class Invalid(ValueError):
    def __init__(self, key: str, message: str):
        super().__init__(message)
        self.key = key


def parse(key: str, value: Any) -> float:
    it = BY_KEY.get(key)
    if it is None:
        raise Invalid(key, f"キーが違います: {key}")
    s = str(value if value is not None else "").strip()
    try:
        x = float(s)
    except ValueError:
        raise Invalid(key, f"{it.label}: 「{s or '(空)'}」は数ではありません") from None
    if x != x or x in (float("inf"), float("-inf")):
        raise Invalid(key, f"{it.label}: 数ではありません")
    if it.integer and x != int(x):
        raise Invalid(key, f"{it.label}: 整数で書いてください")
    if not it.lo <= x <= it.hi:
        raise Invalid(key, f"{it.label}: {s} は範囲外です({fmt(it.lo)}〜{fmt(it.hi)})")
    return float(int(x)) if it.integer else x


def check_all(vals: dict[str, float]) -> None:
    for a, b in PAIRS:
        if vals[a] > vals[b]:
            raise Invalid(a, f"{BY_KEY[a].label} が {BY_KEY[b].label} より大きくなっています")


def fmt(x: float) -> str:
    return str(int(x)) if float(x).is_integer() else repr(float(x))


def _db_path() -> Path:
    return datasource.resolve(DB)


def _has_table(con: sqlite3.Connection) -> bool:
    return con.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?", (TABLE,)).fetchone() is not None


def _rows(vals: dict[str, float]) -> list[tuple[str, str, str, str]]:
    return [(i.key, fmt(vals[i.key]), i.unit, i.label) for i in ITEMS]


def ensure_table(path: Optional[Path] = None) -> bool:
    """表が無ければ作って VBA の既定を入れる(保守からだけ呼ぶ)。足りないキーは既定で足す。"""
    p = path or _db_path()
    if not p.is_file():
        return False
    try:
        con = safe_io.connect_existing(p, timeout=5)
        try:
            with con:
                if not _has_table(con):
                    con.execute(f'CREATE TABLE "{TABLE}" ("キー" TEXT PRIMARY KEY, "値" TEXT NOT NULL, '
                                f'"単位" TEXT NOT NULL DEFAULT \'\', "説明" TEXT NOT NULL DEFAULT \'\')')
                    log.info("寸法しきい値のマスタを作りました: %s", p)
                con.executemany(f'INSERT OR IGNORE INTO "{TABLE}" VALUES (?, ?, ?, ?)', _rows(DEFAULTS))
            return True
        finally:
            con.close()
    except sqlite3.Error as e:
        log.warning("寸法しきい値のマスタを作れませんでした(%s): %s", p, e)
        return False


def _read(p: Path, base: dict[str, float]) -> Optional[dict[str, float]]:
    """読めたら値(表が無ければ既定)、読めなければ None。"""
    try:
        if not p.is_file():
            return None
        con = safe_io.connect_existing(p, timeout=5)
        try:
            if not _has_table(con):
                return dict(DEFAULTS)
        finally:
            con.close()
        cols, recs = datasource.get_table(str(p), TABLE)
    except datasource.TableMissing:
        return dict(DEFAULTS)
    except (datasource.DbMissing, OSError, sqlite3.Error):
        return None
    idx = {c: i for i, c in enumerate(cols)}
    if "キー" not in idx or "値" not in idx:
        return dict(DEFAULTS)
    got = dict(DEFAULTS)
    for r in recs:
        k = str(r[idx["キー"]]).strip()
        if k not in BY_KEY:
            continue
        try:
            got[k] = parse(k, r[idx["値"]])
        except Invalid as e:
            log.warning("寸法しきい値の値が悪いので既定を使います: %s", e)
    try:
        check_all(got)
    except Invalid as e:
        log.warning("寸法しきい値の組が悪いので範囲の値は既定を使います: %s", e)
        for a, b in PAIRS:
            got[a], got[b] = DEFAULTS[a], DEFAULTS[b]
    return got


def values(path: Optional[Path] = None, *, fresh: bool = False) -> dict[str, float]:
    """キー → 値。表が無ければ既定(作らない)。読めなければ最後に読めた値(無ければ既定)。"""
    p = path or _db_path()
    key = str(p)
    now = time.monotonic()
    with _lock:
        hit = _last_known.get(key)
    if hit and not fresh and now - hit[0] < TTL_SEC:
        return dict(hit[1])
    got = _read(p, DEFAULTS)
    if got is None:
        if p.is_file() or hit:
            log.warning("寸法しきい値のマスタを読めないので、最後に読めた値(または既定)を使います: %s", p)
        return dict(hit[1]) if hit else dict(DEFAULTS)
    with _lock:
        _last_known[key] = (now, got)
    return dict(got)


def get(key: str) -> float:
    return values()[key]


def save(new: dict, path: Optional[Path] = None, *, user: str = "") -> dict[str, float]:
    """キー → 文字を確かめて保存(足りないキーはいまの値)。悪ければ Invalid(何も書かない)。前 → 後 を trace に残す。"""
    p = path or _db_path()
    before = values(p, fresh=True)
    vals = {k: parse(k, (new or {}).get(k, fmt(before[k]))) for k in BY_KEY}
    check_all(vals)
    if not ensure_table(p):
        raise OSError(f"伝送用ファイルに書けません: {p}")
    con = safe_io.connect_existing(p, timeout=10)
    try:
        with con:
            con.executemany(f'INSERT OR REPLACE INTO "{TABLE}" VALUES (?, ?, ?, ?)', _rows(vals))
    finally:
        con.close()
    with _lock:
        _last_known[str(p)] = (time.monotonic(), dict(vals))
    changed = {k: [fmt(before[k]), fmt(vals[k])] for k in BY_KEY if before[k] != vals[k]}
    log.info("寸法しきい値を保存しました: %s", changed)
    if changed:
        try:
            from .. import trace
            trace.event(CHANGE_EV, "保守: 寸法のしきい値を変更", data={"changed": changed})
        except Exception as e:  # noqa: BLE001 - 履歴が残せなくても保存は済んでいる
            log.warning("寸法しきい値の変更履歴を残せませんでした: %s", e)
    return vals


def history(days: int = 90, limit: int = 50) -> list[dict[str, Any]]:
    from .. import trace
    try:
        trace.flush()                                   # 保存したばかりの記録も出す
    except Exception:  # noqa: BLE001
        pass
    out = []
    for r in trace.read(days, q=CHANGE_EV, limit=2000):
        if r.get("ev") != CHANGE_EV:
            continue
        d = r.get("data") or {}
        ch = d.get("changed") if isinstance(d, dict) else None
        if not isinstance(ch, dict):
            continue
        out.append({"t": r.get("t", ""), "user": r.get("user", ""), "pc": r.get("pc", ""),
                    "changes": [{"key": k, "label": BY_KEY[k].label if k in BY_KEY else k,
                                 "before": v[0], "after": v[1]} for k, v in ch.items() if isinstance(v, list) and len(v) == 2]})
        if len(out) >= limit:
            break
    return out


# ------------------------------------------------------------------
# 使う側の小さな関数
# ------------------------------------------------------------------
def thin_atu() -> float:
    return get("thin_atu")


def sakko_thick() -> float:
    return get("sakko_thick")


def kennyu_n() -> int:
    return int(get("kennyu_n"))


#: 検入ボタンの文字の頭。状態はこの頭で見分ける(値を変えても前の文字のまま残った状態を取り違えない)
LE_PREFIX = "検入≦"
GE_PREFIX = "検入≧"


def le_caption() -> str:
    """「検入≦25」(いまが 検入≧N+1 = 3 回検査 のときのボタンの文字。押すと ≦N になる)。"""
    return f"{LE_PREFIX}{kennyu_n()}"


def ge_caption() -> str:
    return f"{GE_PREFIX}{kennyu_n() + 1}"


def is_le(caption: str) -> bool:
    return str(caption or "").startswith(LE_PREFIX)


def is_ge(caption: str) -> bool:
    return str(caption or "").startswith(GE_PREFIX)
