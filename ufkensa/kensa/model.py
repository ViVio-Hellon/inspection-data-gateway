"""検査の値の型(第3段 案C)。測定グリッドの「位置 × 行」が何の項目かと、判定の結果をここで決める。

VBA の画面は S1〜E14・F1〜F10 のラベルに値を入れ、行の意味(1〜4 行目 = 厚、5 行目から先は作業で変わる)を
あちこちの `For i = 1 To 4` で表していた。ここに 1 か所でまとめ、ほかのモジュールは名前でなくこの型で扱う。

段ごとに置き換えていく(段1: 厚)。業務の動きと保存・帳票のセル位置は変えない(tests/test_golden_flow.py)。
"""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import TYPE_CHECKING, Iterator

from . import vbx

if TYPE_CHECKING:                                  # pragma: no cover
    from .grid import Ctl, KensaState

#: 測定の位置(列)。F は極厚などで使う 6 番目
POSITIONS: tuple[str, ...] = ("S", "U", "C", "D", "E", "F")
#: F 以外の 5 位置
FIVE: tuple[str, ...] = POSITIONS[:5]
#: 位置ごとの行数(S..E は 1〜14、F は 1〜10)
ROW_COUNT: dict[str, int] = {"S": 14, "U": 14, "C": 14, "D": 14, "E": 14, "F": 10}

#: 厚の行(どの作業でも 1〜4 行目)
THICK_ROWS = range(1, 5)
#: 厚の書式(厚 74mm を超えると小数 2 桁。FormRelat の NumDig)
THICK_FORMAT = "0.000"
THICK_FORMAT_EXTREME = "0.00"
#: 厚以外(幅・丈・対角)の書式
LENGTH_FORMAT = "0.0"


def is_thick(row: int) -> bool:
    """行が厚か(1〜4 行目)。"""
    return row in THICK_ROWS


def row_format(row: int) -> str:
    """保存・読み戻しで値を揃える書式(厚 "0.000"、ほか "0.0")。"""
    return THICK_FORMAT if is_thick(row) else LENGTH_FORMAT


def thick_format(atu: str) -> str:
    """画面に出す厚の書式。製品の厚が 74mm を超えると "0.00"(NumDig)。"""
    return THICK_FORMAT_EXTREME if vbx.val(atu) > 74.0 else THICK_FORMAT


@dataclass(frozen=True)
class RowLayout:
    """作業ごとの行の意味。厚はいつも 1〜4 行目。

    - ふつう(厚板・薄板・少枚数・スラブ): 幅 5〜6、丈 7〜8、対角 9〜10
    - 極厚(耳付): 幅 5〜10(幅1〜6)、丈 11〜14(丈1〜4)、対角なし
    """
    width: range
    length: range
    diag: range

    @property
    def thick(self) -> range:
        return THICK_ROWS

    @property
    def size_rows(self) -> range:
        """厚の次から(幅・丈・対角。ふつうは 5〜10)。"""
        last = max(self.diag.stop, self.length.stop)
        return range(self.width.start, last)

    def labels(self) -> list[str]:
        """行の見出し(WorkerTime・公差外表の 厚1… 幅1… 丈1… 対角A/B)。"""
        out = [f"厚{i}" for i in range(1, len(self.thick) + 1)]
        out += [f"幅{i}" for i in range(1, len(self.width) + 1)]
        out += [f"丈{i}" for i in range(1, len(self.length) + 1)]
        out += [f"対角{c}" for c in "AB"[:len(self.diag)]]
        return out


NORMAL = RowLayout(width=range(5, 7), length=range(7, 9), diag=range(9, 11))
EXTREME = RowLayout(width=range(5, 11), length=range(11, 15), diag=range(11, 11))


def row_layout(extreme: bool) -> RowLayout:
    """極厚(耳付)かどうかで行の意味を返す。"""
    return EXTREME if extreme else NORMAL


@dataclass(frozen=True)
class CellKey:
    """測定グリッドの 1 欄(位置 × 行)。VBA のラベル名は `name`("S1" など)。"""
    pos: str
    row: int

    @property
    def name(self) -> str:
        return f"{self.pos}{self.row}"

    @property
    def thick(self) -> bool:
        return is_thick(self.row)

    @classmethod
    def parse(cls, name: str) -> "CellKey | None":
        """"S1"・"E14" → CellKey。グリッドの欄でなければ None。"""
        if len(name) < 2 or name[0] not in POSITIONS or not name[1:].isdigit():
            return None
        row = int(name[1:])
        return cls(name[0], row) if 1 <= row <= ROW_COUNT[name[0]] else None


def thick_keys(positions: tuple[str, ...] = POSITIONS) -> Iterator[CellKey]:
    """厚の欄を、行 → 位置 の順に(VBA の `For i = 1 To 4: For Each p` と同じ順)。"""
    for row in THICK_ROWS:
        for p in positions:
            yield CellKey(p, row)


class Verdict(str, Enum):
    """欄の判定(画面の色)。値は grid の COLOR_* と同じ文字なので、そのまま比べられる。"""
    NONE = ""
    OK = "OK"
    NG = "NG"
    YELLOW = "YELLOW"
    WHITE = "WHITE"

    @classmethod
    def of(cls, color: str) -> "Verdict":
        try:
            return cls(color)
        except ValueError:
            return cls.NONE


class Measures:
    """KensaState の測定グリッドを、位置 × 行 の型で読み書きする窓口。値の置き場所は KensaState のまま。"""

    def __init__(self, state: "KensaState") -> None:
        self._state = state

    def ctl(self, key: CellKey) -> "Ctl":
        return self._state.cells[key.name]

    def value(self, pos: str, row: int) -> str:
        return self._state.cells[f"{pos}{row}"].value

    def verdict(self, pos: str, row: int) -> Verdict:
        return Verdict.of(self._state.cells[f"{pos}{row}"].color)

    def thick(self, pos: str) -> list[str]:
        """その位置の厚 1〜4 の値。"""
        return [self.value(pos, r) for r in THICK_ROWS]

    def thick_ctls(self, positions: tuple[str, ...] = POSITIONS) -> Iterator[tuple[CellKey, "Ctl"]]:
        for k in thick_keys(positions):
            yield k, self.ctl(k)


# ----------------------------------------------------------------------
# 位置ごとの欄の名前(段3・段4: 対角・直角度・ラテラルボー・BH)。VBA のコントロール名をここだけで組み立てる
# ----------------------------------------------------------------------
def diag_name(pos: str) -> str:
    """対角線差のラベル(DiagS…)。"""
    return f"Diag{pos}"


def kou_name(pos: str) -> str:
    """直角度の欄(DiagSKou…)。"""
    return f"Diag{pos}Kou"


def lateral_name(pos: str) -> str:
    """ラテラルボーの欄(LateralS…)。"""
    return f"Lateral{pos}"


def bh_name(pos: str) -> str:
    """BH の欄(SBH…。F は無い)。"""
    return f"{pos}BH"


# ----------------------------------------------------------------------
# 画面の切り替え・旗(段5・段8)。VBA では Caption / Tag の文字だったもの。値はそのまま(保存・ログと同じ文字)
# ----------------------------------------------------------------------
class Face:
    """歪の入力面(AorB.Caption)。"""
    A = "Now_A"
    B = "Now_B"


class Switch:
    """Load.Caption(フォームのイベントを止める)・Auto(次の入力欄を自動で決める)。"""
    ON = "on"
    OFF = "off"


class FormTag:
    """UF1.Tag。STOP の間は音声・MsgBox を止める。"""
    NONE = ""
    STOP = "STOP"
