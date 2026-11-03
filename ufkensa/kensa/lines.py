"""ライン名(判断事項 D13: all-tools と同じ扱い)

all-tools(`tools/nippou/nippou/logic/line_names.py`)の決まりに合わせる:
- **正規の呼び名**: L-1 / LVC / HVC / トット / バランサー / 中板(+ 試験用の Test)
- マスタ(測定機器ID・校正ストレートエッヂ など)と照合するときは**文字の幅だけ**揃える
  (`ﾄｯﾄ` = `トット`)。大文字小文字・記号・空白は揃えない
- VBA の画面のライン名(UF1 の OptionButton の Caption)は `ﾊﾞﾗﾝｻｰ`・`ﾄｯﾄ`・`丸徳`。
  ログのフォルダ名・帳票・電文など**VBA の書き方がそのまま外に出る所**は VBA の名前を使う
  (`vba_name()`)。どこで何を使うかは呼び手の docstring に書く。
"""
from __future__ import annotations

import unicodedata
from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class Line:
    official: str        # 正規の呼び名(all-tools)
    vba: str             # VBA の LineName(UF1 の Caption)
    folder: str          # VBA のログ・履歴フォルダ名(HistoryPath\<folder>\<folder>_Log)
    big: str             # 履歴の LN(L-1 / LVC / HVC / 大板 / 中板 / Test)
    registry: str        # VBA のレジストリ「設備名」


LINES: tuple[Line, ...] = (
    Line("L-1", "L-1", "L-1", "L-1", "L-1"),
    Line("LVC", "LVC", "LVC", "LVC", "LVC"),
    Line("HVC", "HVC", "HVC", "HVC", "HVC"),
    Line("バランサー", "ﾊﾞﾗﾝｻｰ", "バランサー", "大板", "バランサー"),
    Line("トット", "ﾄｯﾄ", "トット", "大板", "トット"),
    Line("中板", "丸徳", "丸徳", "中板", "丸徳"),
    Line("Test", "Test", "Test", "Test", "Test"),
)


def same(text: object) -> str:
    """比べる形。文字の幅と前後の空白だけ揃える。"""
    return unicodedata.normalize("NFKC", "" if text is None else str(text)).strip()


def find(text: object) -> Optional[Line]:
    """正規・VBA・レジストリのどの書き方からでも引く(幅だけ揃える)。"""
    s = same(text)
    for ln in LINES:
        if s in (same(ln.official), same(ln.vba), same(ln.registry)):
            return ln
    return None


def official(text: object) -> str:
    ln = find(text)
    return ln.official if ln else ("" if text is None else str(text))


def vba_name(text: object) -> str:
    ln = find(text)
    return ln.vba if ln else ("" if text is None else str(text))


def matches(master_value: object, line: object) -> bool:
    """マスタの「ライン」列の値がこのラインか(D13)。"""
    return same(master_value) == same(official(line)) or same(master_value) == same(vba_name(line))
