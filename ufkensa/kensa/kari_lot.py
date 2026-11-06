"""仮ロット(保守画面の歪公差チェック用。2026-11-02 の依頼)

品質基準の表を変えたあと、仕掛に無い組み合わせでも公差を確かめられるように、利用者が自分で
ロットを作って手元DB(`store.LocalStore`)のシート「仮ロット」に保存する。

- ロットNo は英数字 7 文字(例 `A111A11`。本物の形 `N8281C0` も可)
- 検番は「頭の英字 1 文字 + 残りは 1」(例 "H" → "H111111")。長さは本物の検査番号(7 桁)と同じ。
  DBQP が検番で見るのは頭の 1 文字だけ(QP: `KenA = Left(KenNo, 1)`)
- 歪公差は本物のロットと同じ部品(`quality.dist_array` → `dbqp`)で引く。仕掛の行(LotData)と引当
  (HikiArrCreate の (受注_納入先名称, 受注_取引先名称))を**メモリ上で**組み立てるだけで、本物の DB には書かない

シートの形: 1 行目は見出し、2 行目から 1 ロット 1 行(1 列目 = ロットNo、2 列目 = 中身の JSON)。
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Optional

from . import lot as lotmod, quality, vbx
from .store import LocalStore

SHEET = "仮ロット"
HEADER = ("ロットNo", "内容(JSON)")

LOT_RE = re.compile(r"^[A-Z0-9]{7}$")   # 英数字 7 文字(本物の形 N8281C0 も A111A11 も可)
HEAD_RE = re.compile(r"^[A-Z]$")
KEN_LEN = 7                     # SIKAHIKINOW / SIKALOTNOW の 検査番号(例 H6X1176)


@dataclass(frozen=True)
class Field:
    key: str            # 画面の項目名(保存の鍵)
    column: str         # SIKALOTNOW 仕掛 の列名
    num: bool = False   # 数(Single に入れる。空・数でない文字列は VBA では実行時エラー 13)
    group: str = "main"  # main = 利用者が決める欄 / detail = DBQP が読む詳しい欄
    default: str = ""
    hiki: str = ""      # SIKAHIKINOW 仕掛 の列名(引当。DBQP の納入先・取引先はこちらを見る)
    note: str = ""


#: 項目 → 本物の列(QP.DBQP が LotAcc から読む列 QP:176-205、HikiArrCreate QP:1867-1909)
FIELDS: tuple[Field, ...] = (
    Field("用途名", "用途名", note="DBQP は読むが照合には使わない"),
    Field("用途コード", "用途ｺｰﾄﾞ"),
    Field("製造材質", "製造材質"),
    Field("オーダー材質", "ｵｰﾀﾞｰ材質"),
    Field("製造調質", "製造調質"),
    Field("オーダー調質", "ｵｰﾀﾞｰ調質"),
    Field("オーダー板厚", "ｵｰﾀﾞｰ板厚", num=True),
    Field("オーダー板幅", "ｵｰﾀﾞｰ板幅", num=True),
    Field("オーダー板丈", "ｵｰﾀﾞｰ板丈", num=True),
    Field("製造板厚", "製造板厚", num=True),
    Field("製造板幅", "製造板幅", num=True),
    Field("製造板丈", "製造板丈", num=True),
    Field("納入先", "納入先名", hiki="受注_納入先名称",
          note="照合は引当の 受注_納入先名称(SIKALOT の 納入先名 は読むだけ)"),
    Field("納入先コード", "納入先ｺｰﾄﾞ"),
    Field("送り先", "送り先名"),
    Field("送り先コード", "送り先ｺｰﾄﾞ"),
    Field("取引先", "取引先名", hiki="受注_取引先名称",
          note="照合は引当の 受注_取引先名称(SIKALOT の 取引先名 は読むだけ)"),
    Field("取引先コード", "取引先ｺｰﾄﾞ"),
    # DBQP が読む欄(利用者の項目には無いが、結果を変えるので詳しい欄として持つ)
    Field("歪等級", "品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ", group="detail", default="4",
          note="QP 表に該当なしのとき 3/4 で等級表を引く。空は VBA では実行時エラー 13"),
    Field("表面処理等級", "品質ｸﾞﾚｰﾄﾞ_表面処理", group="detail", note="品質基準の画面に出すだけ"),
    Field("横割数", "BOX設計_横割数(板)", group="detail", note="QP 表の 巾取数幅分割数のこと と照合"),
    Field("指示_中歪", "指示_中歪_高さ", group="detail", note="中歪が空のとき補う"),
    Field("指示_耳歪", "指示_耳歪_高さ", group="detail", note="耳歪が空のとき補う"),
)
FIELD_BY_KEY = {f.key: f for f in FIELDS}

#: 仮ロットの LotData の列(DBQP が読む列すべて。本物の SIKALOTNOW を見ないので、ここに無い列は使えない)
LOT_COLUMNS: tuple[str, ...] = ("ﾛｯﾄ番号", "検査番号", "実績_設備ｺｰｽ") + tuple(f.column for f in FIELDS) + (
    "指示_そり巾_方向高さ", "指示_そり丈_方向高さ")


class KariError(ValueError):
    """入力の誤り(文言は画面にそのまま出す)。"""

    def __init__(self, message: str, field: str = "") -> None:
        super().__init__(message)
        self.field = field


@dataclass
class KariLot:
    lot: str
    ken_head: str
    values: dict[str, str]
    hiki_extra: list[tuple[str, str]] = field(default_factory=list)   # 2 件目からの引当(写したときだけ)
    source: str = ""          # 写した元のロットNo
    updated: str = ""

    @property
    def ken_no(self) -> str:
        return ken_no_of(self.ken_head)

    def lot_data(self) -> quality.LotData:
        """仕掛(SIKALOTNOW)の 1 行の代わり。指定の無い列は空。"""
        rec = {c: "" for c in LOT_COLUMNS}
        rec["ﾛｯﾄ番号"] = self.lot
        rec["検査番号"] = self.ken_no
        for f in FIELDS:
            rec[f.column] = self.values.get(f.key, "")
        return quality.LotData(list(LOT_COLUMNS), [[rec[c] for c in LOT_COLUMNS]])

    def hiki(self) -> list[tuple[str, str]]:
        """HikiArrCreate の代わり: 1 件目は 納入先・取引先 の欄、2 件目からは写した元の引当。"""
        return [(self.values.get("納入先", ""), self.values.get("取引先", ""))] + list(self.hiki_extra)

    def to_dict(self) -> dict[str, Any]:
        return {"lot": self.lot, "ken_head": self.ken_head, "ken_no": self.ken_no, "values": dict(self.values),
                "hiki_extra": [list(h) for h in self.hiki_extra], "source": self.source, "updated": self.updated}


def ken_no_of(head: str) -> str:
    """検番 = 頭の英字 + 残りは "1"(例 "H" → "H111111")。"""
    return head + "1" * (KEN_LEN - len(head))


def field_specs() -> list[dict[str, Any]]:
    return [{"key": f.key, "column": f.column, "hiki": f.hiki, "num": f.num, "group": f.group,
             "default": f.default, "note": f.note} for f in FIELDS]


# ------------------------------------------------------------------
# 入力の確かめ
# ------------------------------------------------------------------
def normalize_lot(text: Any) -> str:
    return quality.narrow(str(text or "")).strip().upper()


def validate(lot: Any, ken_head: Any, values: Any, hiki_extra: Any = None, source: Any = "") -> KariLot:
    """画面の入力 → KariLot。誤りは KariError(field = 項目名)。"""
    lot_no = normalize_lot(lot)
    if not LOT_RE.match(lot_no):
        raise KariError("ロットNo は英数字 7 文字で入れてください(例 A111A11・N8281C0)", "lot")
    head = normalize_lot(ken_head)
    if not HEAD_RE.match(head):
        raise KariError("検番の頭は英字 1 文字で選んでください(例 H → H111111)", "ken_head")
    if not isinstance(values, dict):
        raise KariError("項目の値がありません", "values")
    out: dict[str, str] = {}
    for f in FIELDS:
        v = values.get(f.key, f.default)
        v = "" if v is None else str(v).strip()
        if len(v) > 100:
            raise KariError(f"{f.key} が長すぎます(100 文字まで)", f.key)
        if f.num:
            if v == "" or not vbx.is_numeric(v):
                raise KariError(f"{f.key} に数を入れてください(DBQP は Single に入れるので空・文字は実行時エラー 13)", f.key)
        out[f.key] = v
    extra: list[tuple[str, str]] = []
    for h in hiki_extra or []:
        if not isinstance(h, (list, tuple)) or len(h) != 2:
            raise KariError("引当(2件目から)の形が違います", "hiki_extra")
        extra.append((str(h[0] or ""), str(h[1] or "")))
    return KariLot(lot_no, head, out, extra, normalize_lot(source))


# ------------------------------------------------------------------
# 保存(手元DB のシート「仮ロット」)
# ------------------------------------------------------------------
def _rows(store: LocalStore) -> list[tuple[int, KariLot]]:
    if not store.exists(SHEET):
        return []
    sh = store.sheet(SHEET)
    out = []
    for r in range(2, sh.last_row(1) + 1):
        lot_no = sh.text(r, 1)
        if not lot_no:
            continue
        try:
            d = json.loads(sh.text(r, 2) or "{}")
        except ValueError:
            d = {}
        out.append((r, KariLot(lot_no, str(d.get("ken_head", "")), {k: str(v) for k, v in (d.get("values") or {}).items()},
                               [tuple(h) for h in d.get("hiki_extra") or []], str(d.get("source", "")),
                               str(d.get("updated", "")))))
    return out


def list_all(store: LocalStore) -> list[KariLot]:
    return [k for _, k in _rows(store)]


def get(store: LocalStore, lot_no: str) -> Optional[KariLot]:
    lot_no = normalize_lot(lot_no)
    return next((k for _, k in _rows(store) if k.lot == lot_no), None)


def _row_of(store: LocalStore, lot_no: str) -> int:
    return next((r for r, k in _rows(store) if k.lot == lot_no), 0)


def real_lot_exists(lot_no: str) -> bool:
    """仕掛(SIKALOTNOW / SIKAHIKINOW)に同じロットNo があるか(読むだけ。DB が無ければ False)。"""
    try:
        return bool(lotmod.arrbox("SIKALOTNOW", lot_no, "Record").value) or bool(quality.hiki_arr_create(lot_no))
    except Exception:                                   # noqa: BLE001 - 読めなければ重なりは調べない
        return False


def save(store: LocalStore, k: KariLot, original: str = "") -> tuple[str, KariLot]:
    """作る(original 空)・直す(original = 直す前のロットNo)。("作成" / "変更", 保存したもの)。"""
    original = normalize_lot(original)
    if k.lot != original and get(store, k.lot) is not None:
        raise KariError(f"仮ロット {k.lot} はもうあります", "lot")
    if k.lot != original and real_lot_exists(k.lot):
        raise KariError(f"仕掛に同じロットNo {k.lot} があります。別のロットNo にしてください", "lot")
    k.updated = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    body = json.dumps({"ken_head": k.ken_head, "values": k.values, "hiki_extra": [list(h) for h in k.hiki_extra],
                       "source": k.source, "updated": k.updated}, ensure_ascii=False)
    sh = store.ensure_sheet(SHEET)
    sh.write_row(1, 1, HEADER, "S")
    if original:
        r = _row_of(store, original)
        if r == 0:
            raise KariError(f"仮ロット {original} が見つかりません(消されたかもしれません)", "lot")
        action = "変更"
    else:
        r = sh.last_row(1) + 1
        action = "作成"
    sh.write(r, 1, k.lot, "S")
    sh.write(r, 2, body, "S")
    return action, k


def delete(store: LocalStore, lot_no: str) -> bool:
    r = _row_of(store, normalize_lot(lot_no))
    if r == 0:
        return False
    store.sheet(SHEET).delete_row(r)
    return True


# ------------------------------------------------------------------
# 本物のロットから写す
# ------------------------------------------------------------------
def from_real(lot_no: str) -> dict[str, Any]:
    """仕掛(SIKALOTNOW 1 行目・SIKAHIKINOW の引当・検番 D15)から項目の値を作る(保存はしない)。

    検番の頭は本物の検番の頭(数字・空なら DBQP と同じく 実績_設備ｺｰｽ から FindPriority)。
    """
    from . import quality_session
    lot_no = normalize_lot(lot_no)
    f = lotmod.arrbox("SIKALOTNOW", "板", "Fields")
    r = lotmod.arrbox("SIKALOTNOW", lot_no, "Record")
    data = quality.LotData(f.value, r.value)
    hiki = quality.hiki_arr_create(lot_no) or []
    if not r.value or not f.value:
        why = " ".join(m.text for m in (f.messages + r.messages)) or f"読んだファイル: {r.file_path or f.file_path}"
        raise KariError(f"仕掛(SIKALOTNOW)にロット {lot_no} がありません。{why}", "source")

    def col(name: str) -> str:
        c = quality.col_of(f.value, name)
        return quality.vstr(r.value[0][c - 1]) if c else ""
    values = {fd.key: col(fd.column) for fd in FIELDS}
    notes: list[str] = []
    if hiki:
        values["納入先"], values["取引先"] = hiki[0]
    else:
        notes.append("引当(SIKAHIKINOW)が無いので、納入先・取引先は SIKALOT の 納入先名・取引先名 を入れました"
                     "(本物のロットは引当が無いと歪公差を引きません)")
    ken = quality_session.ken_no_of(lot_no, data)
    head = vbx.left(ken, 1).upper()
    if not HEAD_RE.match(head):
        head = quality.find_priority(col("実績_設備ｺｰｽ"))
        notes.append(f"検番 {ken or '(空)'} の頭が英字でないので、実績_設備ｺｰｽ から {head or '(なし)'} にしました")
    if len(hiki) > 1:
        notes.append(f"引当が {len(hiki)} 件あります。2 件目からは元のロットのまま持ちます")
    return {"source": lot_no, "ken_no": ken, "ken_head": head, "values": values,
            "hiki_extra": [list(h) for h in hiki[1:]], "notes": notes}
