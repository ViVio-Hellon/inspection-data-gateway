"""手元DB(端末ごとの SQLite)。VBA がブック内のシートに貯めていたものの置き場

VBA は検査結果・作業中のデータ・履歴をブック(xlsm)のシートに書いていた(G3 §3.3)。
D4/D5 で、これを端末ごとの SQLite に移す。**シートの形(行・列・値の型)はそのまま**にして、
VBA のセル操作(`Cells(r, c)`・`UR`・`UC`・`AddWS`・`Rows(i).Delete`)を同じ意味で行えるようにする。

【2 種類の表】
- **列の決まった表**(1 行目が見出しのシート): `保存用`(114 列)・`耳用`(108 列)・`Lot情報履歴`(26 列)・
  `印刷履歴管理`(3 列)。列名 = シートの 1 行目の見出し。`行` 列がシートの行番号(2 から)。
- **セルの表**(形が決まっていないシート): `作業保存_<L>`・`WorkerTime_<L>`・`検査員_<L>`・`Daily_<L>`・
  `梱包枚数計算_<L>`・`校正記録`・`Accesslog`・`異常ロット用_<L>`・`不良要因印刷_<L>` など。
  `セル(シート, 行, 列, 値)` に 1 セル 1 行で持つ。

【値の型】セルの値は Excel と同じく **空(None)・文字列・数** のどれか。列に型を付けない(SQLite の型なし列)
ので、入れた型のまま戻る。VBA の代入で Excel が文字列を数にする規則は `excel_value()`(下)。

【Excel の代入の規則】(サンプルの xlsm で確かめた形。G3 §9.6 の注を参照)
- Variant の文字列(1 セルへの代入・Variant 配列)… 数として読める文字列は**数**になる。"" は**空セル**。
  先頭の `'` は付けた文字列のまま(`'` は消える)
- String 型の配列(`Split` の結果・`Dim x() As String`)… 文字列のまま。"" は空セル
- コントロールそのもの(`Array(.ATU, …)`)… 既定のプロパティの文字列のまま。"" も**長さ 0 の文字列**として残る
  → `write(r, c, v, kind)` の `kind` = "V" / "S" / "O"
"""
from __future__ import annotations

import re
import sqlite3
import threading
from pathlib import Path
from typing import Any, Iterable, Optional, Sequence, Union

from .. import app_config
from ..logging_utils import get_logger

log = get_logger("kensa.store")

Value = Union[None, str, float, int]

BUSY_TIMEOUT_MS = 10_000

# ------------------------------------------------------------------
# 見出し(シートの 1 行目そのまま)
# ------------------------------------------------------------------
HOZON = "保存用"
MIMI = "耳用"
LOT_HISTORY = "Lot情報履歴"
PRINT_HISTORY = "印刷履歴管理"

HOZON_HEADERS: tuple[str, ...] = (
    '検査開始終了', 'LOTNO', '厚', '幅', '丈', '厚-', '厚+', '幅-', '幅+', '丈-', '丈+',
    'S厚1', 'S厚2', 'S厚3', 'S厚4', 'S巾1', 'S巾2', 'S丈1', 'S丈2', 'S対角A', 'S対角B', 'S定盤歪', 'Sゲージ歪',
    'U厚1', 'U厚2', 'U厚3', 'U厚4', 'U巾1', 'U巾2', 'U丈1', 'U丈2', 'U対角A', 'U対角B', 'U定盤歪', 'Uゲージ歪',
    'C厚1', 'C厚2', 'C厚3', 'C厚4', 'C巾1', 'C巾2', 'C丈1', 'C丈2', 'C対角A', 'C対角B', 'C定盤歪', 'Cゲージ歪',
    'D厚1', 'D厚2', 'D厚3', 'D厚4', 'D巾1', 'D巾2', 'D丈1', 'D丈2', 'D対角A', 'D対角B', 'D定盤歪', 'Dゲージ歪',
    'E厚1', 'E厚2', 'E厚3', 'E厚4', 'E巾1', 'E巾2', 'E丈1', 'E丈2', 'E対角A', 'E対角B', 'E定盤歪', 'Eゲージ歪',
    'Sﾗﾃﾗﾙﾎﾞｰ', 'Uﾗﾃﾗﾙﾎﾞｰ', 'Cﾗﾃﾗﾙﾎﾞｰ', 'Dﾗﾃﾗﾙﾎﾞｰ', 'Eﾗﾃﾗﾙﾎﾞｰ', '検査員', '公差修正', '材質', '調質',
    '用途コード', 'LoadFile', '検査頻度', 'フリー反り', 'S外観', 'U外観', 'C外観', 'D外観', 'E外観',
    'S_BH', 'U_BH', 'C_BH', 'D_BH', 'E_BH', 'S_直角度', 'U_直角度', 'C_直角度', 'D_直角度', 'E_直角度',
    'F厚1', 'F厚2', 'F厚3', 'F厚4', 'F巾1', 'F巾2', 'F丈1', 'F丈2', 'F対角A', 'F対角B', 'F定盤歪', 'Fゲージ歪',
    'Fﾗﾃﾗﾙﾎﾞｰ', 'F外観', 'F_直角度')

MIMI_HEADERS: tuple[str, ...] = tuple(
    ['検査開始終了', 'LOTNO', '厚', '幅', '丈', '厚-', '厚+', '幅-', '幅+', '丈-', '丈+']
    + [f"{f}{k}" for f in "SUCDE" for k in
       ("厚1", "厚2", "厚3", "厚4", "巾1", "巾2", "巾3", "巾4", "巾5", "巾6", "丈1", "丈2", "丈3", "丈4",
        "定盤歪", "ゲージ歪")]
    + ['検査員', '公差修正', '材質', '調質', '用途コード', 'LoadFile', '検査頻度',
       'S外観', 'U外観', 'C外観', 'D外観', 'E外観', 'S_BH', 'U_BH', 'C_BH', 'D_BH', 'E_BH'])

LOT_HISTORY_HEADERS: tuple[str, ...] = (
    '登録日', 'ｵｰﾀﾞｰNo', 'ﾛｯﾄNo', '厚', '巾', '丈', '検査No', '材質', '調質', '用途ｺｰﾄﾞ', 'ｺｰｽ', '得意先',
    '納入先', '送り先', '前工程実績枚数', '前工程実績重量', '合紙', 'VC', 'VC_A', 'VC_B', '単重', '引当NO',
    '引当調整NO', '引当数量', '引当ﾌｧｲﾙ登録', '包装仕様NO')

PRINT_HISTORY_HEADERS: tuple[str, ...] = ("LOT番号", "印刷日時", "ライン名")

ROW_TABLES: dict[str, tuple[str, ...]] = {
    HOZON: HOZON_HEADERS,
    MIMI: MIMI_HEADERS,
    LOT_HISTORY: LOT_HISTORY_HEADERS,
    PRINT_HISTORY: PRINT_HISTORY_HEADERS,
}

assert len(HOZON_HEADERS) == 114 and len(MIMI_HEADERS) == 108 and len(LOT_HISTORY_HEADERS) == 26


def line_sheet(prefix: str, lname: str) -> str:
    """ライン別シートの名前(VBA の `"作業保存" & "_" & Lname`)。`lname` は LineName.Caption & LineNo.Caption。"""
    return f"{prefix}_{lname}"


# ------------------------------------------------------------------
# Excel の代入で値がどうなるか
# ------------------------------------------------------------------
_NUM_RE = re.compile(r"^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d*)?(?:[eE][+-]?\d+)?$")


def excel_number(text: str) -> Optional[float]:
    """Excel が「入力された文字列」を数と読むときの値。読めなければ None。

    対応: 前後の半角空白、符号、桁区切りのカンマ(3 桁ごと)、小数、指数、末尾の %。
    日付・時刻(`2-3`・`9:30` など)は数にしない(VBA が `'` を付けて避けている。`'3-2` など)。
    """
    s = text.strip(" ")
    if not s:
        return None
    pct = s.endswith("%")
    if pct:
        s = s[:-1].strip(" ")
    if not s or not _NUM_RE.match(s) or not re.search(r"\d", s):
        return None
    if re.match(r"^[+-]?[eE.]", s.replace(",", "")) and not re.match(r"^[+-]?\.\d", s):
        return None
    try:
        v = float(s.replace(",", ""))
    except ValueError:
        return None
    return v / 100 if pct else v


def excel_value(v: Any, kind: str = "V") -> Value:
    """VBA の代入 `Cells(r, c) = v` でセルに入る値。kind は上の注を参照("V" / "S" / "O")。"""
    if v is None:
        return None
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"     # VBA の True を代入すると Excel の論理値(表示は TRUE)
    if isinstance(v, (int, float)):
        return v
    s = str(v)
    if kind == "O":
        return s
    if s == "":
        return None
    if kind == "S":
        return s
    if s.startswith("'"):
        return s[1:]
    n = excel_number(s)
    return n if n is not None else s


def is_blank(v: Value) -> bool:
    """End(xlUp) などで「何も入っていない」セルか(長さ 0 の文字列は値ありと見る)。"""
    return v is None


def cell_text(v: Value) -> str:
    """セルの値を VBA の文字列にしたもの(CStr。空は "")。"""
    from . import vbx
    if v is None:
        return ""
    if isinstance(v, float):
        return vbx.cstr(v)
    return str(v)


def col_letter(c: int) -> str:
    s = ""
    while c > 0:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _q(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


# ------------------------------------------------------------------
# 手元DB
# ------------------------------------------------------------------
def default_path() -> Path:
    return app_config.local_dir("data") / "kensa_local.sqlite3"


class LocalStore:
    """端末ごとの手元DB。`LocalStore()` で既定の場所、試験では `LocalStore(tmp_path)`。"""

    def __init__(self, path: Optional[Union[str, Path]] = None) -> None:
        self.path = Path(path) if path else default_path()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.RLock()
        self.conn = sqlite3.connect(str(self.path), timeout=BUSY_TIMEOUT_MS / 1000, check_same_thread=False)
        self.conn.execute(f"PRAGMA busy_timeout = {BUSY_TIMEOUT_MS}")
        self._create()

    def close(self) -> None:
        self.conn.close()

    def __enter__(self) -> "LocalStore":
        return self

    def __exit__(self, *exc: Any) -> None:
        self.close()

    def _create(self) -> None:
        with self._lock, self.conn:
            c = self.conn
            c.execute('CREATE TABLE IF NOT EXISTS "シート" ("名前" TEXT PRIMARY KEY, "順" INTEGER)')
            c.execute('CREATE TABLE IF NOT EXISTS "セル" ("シート" TEXT NOT NULL, "行" INTEGER NOT NULL, '
                      '"列" INTEGER NOT NULL, "値", "色" INTEGER, PRIMARY KEY ("シート", "行", "列"))')
            c.execute('CREATE TABLE IF NOT EXISTS "帳票" ("シート" TEXT PRIMARY KEY, "内容" TEXT NOT NULL)')
            for name, headers in ROW_TABLES.items():
                cols = ", ".join(_q(h) for h in headers)
                c.execute(f'CREATE TABLE IF NOT EXISTS {_q(name)} ("行" INTEGER PRIMARY KEY, {cols})')
                c.execute('INSERT OR IGNORE INTO "シート" ("名前", "順") VALUES (?, 0)', (name,))

    # -- シート ------------------------------------------------------
    def sheet(self, name: str) -> "Sheet":
        if name in ROW_TABLES:
            return RowSheet(self, name, ROW_TABLES[name])
        return GridSheet(self, name)

    def exists(self, name: str) -> bool:
        """VBA `WsCheck(name)`。"""
        r = self.conn.execute('SELECT 1 FROM "シート" WHERE "名前" = ?', (name,)).fetchone()
        return r is not None

    def add_sheet(self, name: str) -> "Sheet":
        """VBA `AddWS(name)`。同じ名前があれば**消してから**作り直す(中身は空)。"""
        with self._lock, self.conn:
            if name in ROW_TABLES:
                self.conn.execute(f"DELETE FROM {_q(name)}")
            else:
                self.conn.execute('DELETE FROM "セル" WHERE "シート" = ?', (name,))
            self.conn.execute('INSERT OR REPLACE INTO "シート" ("名前", "順") VALUES (?, 0)', (name,))
        return self.sheet(name)

    def ensure_sheet(self, name: str) -> "Sheet":
        """VBA `ShJuge(name)`(無ければ作る)。"""
        if not self.exists(name):
            return self.add_sheet(name)
        return self.sheet(name)

    def delete_sheet(self, name: str) -> None:
        with self._lock, self.conn:
            if name in ROW_TABLES:
                self.conn.execute(f"DELETE FROM {_q(name)}")
            else:
                self.conn.execute('DELETE FROM "セル" WHERE "シート" = ?', (name,))
            self.conn.execute('DELETE FROM "シート" WHERE "名前" = ?', (name,))

    def sheet_names(self) -> list[str]:
        return [r[0] for r in self.conn.execute('SELECT "名前" FROM "シート" ORDER BY "名前"')]

    # -- 帳票(印刷用_<L> の出来上がり。再印刷 PrintOnly で使う) --------
    def save_report(self, name: str, content_json: str) -> None:
        with self._lock, self.conn:
            self.conn.execute('INSERT OR REPLACE INTO "帳票" ("シート", "内容") VALUES (?, ?)', (name, content_json))

    def load_report(self, name: str) -> Optional[str]:
        r = self.conn.execute('SELECT "内容" FROM "帳票" WHERE "シート" = ?', (name,)).fetchone()
        return r[0] if r else None


class Sheet:
    """シート 1 枚(VBA の Worksheet に当たる操作だけ)。行・列は 1 始まり。"""

    def __init__(self, store: LocalStore, name: str) -> None:
        self.store = store
        self.name = name

    @property
    def conn(self) -> sqlite3.Connection:
        return self.store.conn

    def exists(self) -> bool:
        return self.store.exists(self.name)

    # 下位クラスで実装
    def get(self, r: int, c: int) -> Value:
        raise NotImplementedError

    def set(self, r: int, c: int, v: Value) -> None:
        raise NotImplementedError

    def used(self) -> list[tuple[int, int, Value]]:
        raise NotImplementedError

    def delete_row(self, r: int) -> None:
        raise NotImplementedError

    # 共通
    def write(self, r: int, c: int, v: Any, kind: str = "V") -> None:
        """VBA `Cells(r, c) = v`(Excel の型の変わり方つき)。"""
        self.set(r, c, excel_value(v, kind))

    def write_row(self, r: int, c: int, values: Sequence[Any], kind: str = "V") -> None:
        """`Range(Cells(r, c), Cells(r, c + n - 1)) = 配列`。"""
        for i, v in enumerate(values):
            self.write(r, c + i, v, kind)

    def text(self, r: int, c: int) -> str:
        return cell_text(self.get(r, c))

    def read_range(self, r1: int, c1: int, r2: int, c2: int) -> list[list[Value]]:
        """`Range(Cells(r1, c1), Cells(r2, c2)).Value`。Excel と同じく逆向きの指定は並べ替える。"""
        ra, rb = sorted((r1, r2))
        ca, cb = sorted((c1, c2))
        return [[self.get(r, c) for c in range(ca, cb + 1)] for r in range(ra, rb + 1)]

    def last_row(self, col: int) -> int:
        """`Cells(Rows.Count, col).End(xlUp).Row`(列に何も無ければ 1)。"""
        rows = [r for r, c, v in self.used() if c == col and not is_blank(v)]
        return max(rows) if rows else 1

    def last_col(self, row: int) -> int:
        """`Cells(row, Columns.Count).End(xlToLeft).Column`(行に何も無ければ 1)。"""
        cols = [c for r, c, v in self.used() if r == row and not is_blank(v)]
        return max(cols) if cols else 1

    def last_cell(self) -> tuple[int, int]:
        """`SpecialCells(xlLastCell)`(値のある範囲の右下。書式だけのセルは見ない近似)。"""
        cells = [(r, c) for r, c, v in self.used() if not is_blank(v)]
        if not cells:
            return 1, 1
        return max(r for r, _ in cells), max(c for _, c in cells)


class GridSheet(Sheet):
    """形の決まっていないシート(セル単位)。"""

    def get(self, r: int, c: int) -> Value:
        row = self.conn.execute('SELECT "値" FROM "セル" WHERE "シート" = ? AND "行" = ? AND "列" = ?',
                                (self.name, r, c)).fetchone()
        return row[0] if row else None

    def set(self, r: int, c: int, v: Value) -> None:
        with self.store._lock, self.conn:
            if not self.store.exists(self.name):
                self.conn.execute('INSERT OR IGNORE INTO "シート" ("名前", "順") VALUES (?, 0)', (self.name,))
            if v is None:
                self.conn.execute('UPDATE "セル" SET "値" = NULL WHERE "シート" = ? AND "行" = ? AND "列" = ?',
                                  (self.name, r, c))
                self.conn.execute('DELETE FROM "セル" WHERE "シート" = ? AND "行" = ? AND "列" = ? '
                                  'AND "色" IS NULL', (self.name, r, c))
            else:
                self.conn.execute('INSERT INTO "セル" ("シート", "行", "列", "値") VALUES (?, ?, ?, ?) '
                                  'ON CONFLICT ("シート", "行", "列") DO UPDATE SET "値" = excluded."値"',
                                  (self.name, r, c, v))

    def color(self, r: int, c: int) -> Optional[int]:
        row = self.conn.execute('SELECT "色" FROM "セル" WHERE "シート" = ? AND "行" = ? AND "列" = ?',
                                (self.name, r, c)).fetchone()
        return row[0] if row else None

    def set_color(self, r: int, c: int, rgb: Optional[int]) -> None:
        """背景色(VBA `Interior.Color`。None = 塗りなし)。"""
        with self.store._lock, self.conn:
            self.conn.execute('INSERT INTO "セル" ("シート", "行", "列", "値", "色") VALUES (?, ?, ?, NULL, ?) '
                              'ON CONFLICT ("シート", "行", "列") DO UPDATE SET "色" = excluded."色"',
                              (self.name, r, c, rgb))

    def used(self) -> list[tuple[int, int, Value]]:
        return [(r, c, v) for r, c, v in self.conn.execute(
            'SELECT "行", "列", "値" FROM "セル" WHERE "シート" = ?', (self.name,))]

    def last_row(self, col: int) -> int:
        r = self.conn.execute('SELECT MAX("行") FROM "セル" WHERE "シート" = ? AND "列" = ? AND "値" IS NOT NULL',
                              (self.name, col)).fetchone()
        return r[0] if r and r[0] else 1

    def last_col(self, row: int) -> int:
        r = self.conn.execute('SELECT MAX("列") FROM "セル" WHERE "シート" = ? AND "行" = ? AND "値" IS NOT NULL',
                              (self.name, row)).fetchone()
        return r[0] if r and r[0] else 1

    def delete_row(self, r: int) -> None:
        """`Rows(r).Delete`(下の行を 1 つ上げる)。"""
        with self.store._lock, self.conn:
            self.conn.execute('DELETE FROM "セル" WHERE "シート" = ? AND "行" = ?', (self.name, r))
            self.conn.execute('UPDATE "セル" SET "行" = -("行" - 1) WHERE "シート" = ? AND "行" > ?', (self.name, r))
            self.conn.execute('UPDATE "セル" SET "行" = -"行" WHERE "シート" = ? AND "行" < 0', (self.name,))

    def clear(self) -> None:
        with self.store._lock, self.conn:
            self.conn.execute('DELETE FROM "セル" WHERE "シート" = ?', (self.name,))


class RowSheet(Sheet):
    """1 行目が見出しの表(保存用・耳用・Lot情報履歴・印刷履歴管理)。1 行目は見出しで固定。"""

    def __init__(self, store: LocalStore, name: str, headers: tuple[str, ...]) -> None:
        super().__init__(store, name)
        self.headers = headers

    def _col(self, c: int) -> str:
        if not 1 <= c <= len(self.headers):
            raise IndexError(f"{self.name} に {col_letter(c)} 列はありません(列は {len(self.headers)} まで)")
        return self.headers[c - 1]

    def get(self, r: int, c: int) -> Value:
        if r == 1:
            return self.headers[c - 1] if 1 <= c <= len(self.headers) else None
        if not 1 <= c <= len(self.headers):
            return None
        row = self.conn.execute(f'SELECT {_q(self._col(c))} FROM {_q(self.name)} WHERE "行" = ?', (r,)).fetchone()
        return row[0] if row else None

    def set(self, r: int, c: int, v: Value) -> None:
        if r == 1:
            raise ValueError(f"{self.name} の見出し(1 行目)は変えられません")
        if c > len(self.headers):
            # 見出しより右のセル。VBA のシートには書けるが、DB 登録(UC(sh,1) まで)にも使われないので持たない
            log.warning("%s の %s%d(見出しより右)への書き込みを捨てました: %r", self.name, col_letter(c), r, v)
            return
        col = self._col(c)
        with self.store._lock, self.conn:
            self.conn.execute(f'INSERT OR IGNORE INTO {_q(self.name)} ("行") VALUES (?)', (r,))
            self.conn.execute(f'UPDATE {_q(self.name)} SET {_q(col)} = ? WHERE "行" = ?', (v, r))

    def row(self, r: int) -> list[Value]:
        """1 行(A 列から最後の見出しの列まで)。"""
        if r == 1:
            return list(self.headers)
        cols = ", ".join(_q(h) for h in self.headers)
        got = self.conn.execute(f'SELECT {cols} FROM {_q(self.name)} WHERE "行" = ?', (r,)).fetchone()
        return list(got) if got else [None] * len(self.headers)

    def rows(self) -> list[tuple[int, list[Value]]]:
        cols = ", ".join(_q(h) for h in self.headers)
        return [(r[0], list(r[1:])) for r in self.conn.execute(
            f'SELECT "行", {cols} FROM {_q(self.name)} ORDER BY "行"')]

    def used(self) -> list[tuple[int, int, Value]]:
        out: list[tuple[int, int, Value]] = [(1, i + 1, h) for i, h in enumerate(self.headers)]
        for r, vals in self.rows():
            out.extend((r, i + 1, v) for i, v in enumerate(vals) if v is not None)
        return out

    def last_row(self, col: int) -> int:
        if not 1 <= col <= len(self.headers):
            return 1
        r = self.conn.execute(f'SELECT MAX("行") FROM {_q(self.name)} WHERE {_q(self._col(col))} IS NOT NULL'
                              ).fetchone()
        return r[0] if r and r[0] else 1

    def last_col(self, row: int) -> int:
        if row == 1:
            return len(self.headers)
        vals = self.row(row)
        idx = [i + 1 for i, v in enumerate(vals) if v is not None]
        return max(idx) if idx else 1

    def delete_row(self, r: int) -> None:
        if r < 2:
            raise ValueError("見出しの行は消せません")
        with self.store._lock, self.conn:
            self.conn.execute(f'DELETE FROM {_q(self.name)} WHERE "行" = ?', (r,))
            self.conn.execute(f'UPDATE {_q(self.name)} SET "行" = -("行" - 1) WHERE "行" > ?', (r,))
            self.conn.execute(f'UPDATE {_q(self.name)} SET "行" = -"行" WHERE "行" < 0')

    def clear(self) -> None:
        with self.store._lock, self.conn:
            self.conn.execute(f"DELETE FROM {_q(self.name)}")

    def append(self, values: Sequence[Any], kind: str = "O") -> int:
        """取り込み用: 最後の行の次に 1 行入れる(`kind` の既定は文字列のまま)。入れた行番号を返す。"""
        r = self.conn.execute(f'SELECT MAX("行") FROM {_q(self.name)}').fetchone()[0]
        n = max(int(r or 1), 1) + 1
        self.write_row(n, 1, list(values)[:len(self.headers)], kind)
        return n


def import_cells(store: LocalStore, name: str, cells: Iterable[tuple[int, int, Value]]) -> int:
    """D5 の取り込み: xlsm のシートのセル(行, 列, 値)をそのまま(型も)入れる。見出しの表は 2 行目から。

    値は Excel のセルの値(None・文字列・数。日付は呼び手で CStr の文字にしておく)。入れたセルの数を返す。
    """
    sh = store.ensure_sheet(name)
    n = 0
    for r, c, v in cells:
        if isinstance(sh, RowSheet) and r == 1:
            continue
        sh.set(r, c, v)
        n += 1
    return n
