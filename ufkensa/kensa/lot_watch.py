"""ロット情報の読み直しと見張り(2026-11-07 のご依頼。VBA には無い)。

検査中のロットについて、元の情報を種類ごとに「指紋」(ハッシュ)にして覚えておく。
- 仕掛: 仕掛台帳(SIKALOT)のそのロットの行(製造の厚・表面処理・用途・指示 など)
- 引当: 引当(SIKAHIKI)のそのロットの行(引当の数・受注)
- オーダー: 引当の受注番号の、オーダー(SIKAODR)の行(納入先・包装仕様・VC など)
- 品質基準: 品質基準 DB(公差の表)が変わったか(ファイルの更新)

ロット登録のときに覚え、作業者の「読み直し」と、画面からの定期の見張り(watch)で今の指紋と比べる。
どれかが変わっていれば、変わった種類を返す(引当だけ・オーダーだけ・公差だけ… の場合がある)。
"""
from __future__ import annotations

import hashlib
import os
from dataclasses import dataclass, field
from typing import Any, Optional

from . import datasource, lot as lotmod, paths

KINDS = ("仕掛", "引当", "オーダー", "品質基準")

#: 画面に出す公差・規格の項目(読み直しの前後で比べる。KensaState の属性 / 表示の欄)
TOLERANCE_ATTRS = {
    "atu": "厚", "hab": "巾", "tak": "丈", "apk": "厚 上限", "amk": "厚 下限", "hpk": "巾 上限", "hmk": "巾 下限",
    "tpk": "丈 上限", "tmk": "丈 下限", "ap": "厚 +公差", "am": "厚 -公差", "hp": "巾 +公差", "hm": "巾 -公差",
    "tp": "丈 +公差", "tm": "丈 -公差", "lateral_text": "ラテラルボー公差", "huris": "フリー反り公差",
    "flatnes_text": "フリー反り", "youto": "用途", "zai": "材質", "tyu": "調質", "sumisei": "ラテラル要の納入先",
    "additional": "技術指示", "ken_no_caption": "検査番号",
}
TOLERANCE_VIEWS = {"TaikText": "対角 公差", "HaHiz": "歪 巾", "TaHiz": "歪 丈", "Naka": "歪 中", "Mimi": "歪 耳",
                   "MG": "歪 MG"}


@dataclass
class Signature:
    lot: str
    parts: dict[str, str] = field(default_factory=dict)

    def changed(self, other: "Signature") -> list[str]:
        """変わった種類(同じロットのときだけ。読めなかった種類は比べない)。"""
        if other is None or other.lot != self.lot:
            return []
        return [k for k in KINDS if k in self.parts and k in other.parts and self.parts[k] != other.parts[k]]


def _digest(rows: Any) -> str:
    return hashlib.sha1(repr(rows).encode("utf-8")).hexdigest()


def _base(dbflag: str) -> str:
    return paths.get("main") if dbflag == "" else paths.get("next")


def _rows(path: str, lot: str) -> Optional[list[list[str]]]:
    try:
        return datasource.get_records(path, lotmod.TABLE, lotmod.arrbox_filter(lot))
    except (datasource.DbMissing, datasource.TableMissing, OSError):
        return None


def signature(lot: str, dbflag: str = "") -> Signature:
    """いまのロット情報の指紋。読めない種類は入れない(比べない)。"""
    sig = Signature(lot)
    if lot == "":
        return sig
    base = _base(dbflag)
    lot_rows = _rows(lotmod.db_path(base, "SIKALOTNOW"), lot)
    if lot_rows is not None:
        sig.parts["仕掛"] = _digest(lot_rows)
    hiki_path = lotmod.db_path(base, "SIKAHIKINOW")
    hiki = _rows(hiki_path, lot)
    if hiki is not None:
        sig.parts["引当"] = _digest(hiki)
        try:
            cols = datasource.get_fields(hiki_path, lotmod.TABLE, "") or []
            i = datasource.column_index(cols, "受注番号")
            orders = sorted({r[i] for r in hiki}) if i >= 0 else []
            odr_path = lotmod.db_path(base, "SIKAODRNOW")
            ocols = datasource.get_fields(odr_path, lotmod.TABLE, "") or []
            j = datasource.column_index(ocols, "受注番号")
            orows = (datasource.get_records(odr_path, lotmod.TABLE, "") or []) if j >= 0 else []
            sig.parts["オーダー"] = _digest(sorted(r for r in orows if r[j] in orders) if j >= 0 else [])
        except (datasource.DbMissing, datasource.TableMissing, OSError):
            pass
    try:
        q = paths.to_local(paths.db_file("品質基準"))
        st = os.stat(q)
        sig.parts["品質基準"] = f"{st.st_mtime_ns}:{st.st_size}"
    except (OSError, KeyError, ValueError):
        pass
    return sig


def tolerance_view(state: Any) -> dict[str, str]:
    """画面の公差・規格(読み直しの前後で比べる)。"""
    out = {label: str(getattr(state, a, "")) for a, label in TOLERANCE_ATTRS.items()}
    for name, label in TOLERANCE_VIEWS.items():
        out[label] = state.view[name].value if name in state.view else ""
    return out


def tolerance_diff(before: dict[str, str], after: dict[str, str]) -> list[str]:
    return [f"{k}: {before.get(k, '') or '(空)'} → {after.get(k, '') or '(空)'}"
            for k in after if before.get(k, "") != after.get(k, "")]
