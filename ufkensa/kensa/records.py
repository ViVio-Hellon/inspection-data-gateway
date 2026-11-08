"""検査結果の保存・測定DB への登録・作業中保存・履歴(VBA `【UFKen_03】SaveLoad` ほか)

対応する VBA:
    SaveLoad.MainDataSave (SaveLoad:4216-4569)  → main_data_save()
    SQL_追加.adoSQL (SQL_追加:525-704)          → register_measurement()(ADDNEW / UPDATE)
    SaveLoad.Posting (SaveLoad:1981-2133)       → posting()
    SaveLoad.ReflectStrain / LosdStrain (SaveLoad:1855-1979) → reflect_strain() / losd_strain()
    SaveLoad.Load_Text (SaveLoad:2501-2992)     → load_text()
    SaveLoad.HistoryLoad (SaveLoad:3409-4049)   → history_load()
    SaveLoad.UF_Array / UF_SaveArray (SaveLoad:2135-2286), OutTolerance.UFtag_Array → uf_array() ほか
    Worker.Number_Count (Worker:948-1023)       → number_count()
    Worker.EndTime_Register (Worker:312-325)    → end_time_register()
    Input.ExpCSV (Input:838-874)                → exp_csv() / export_history_csv()(UF1:3780-3815 の出し先)
    UF管理 (UF管理.frm)                         → kanri_lots() / kanri_view() / kanri_select_time() / kanri_load()

画面(UF1)のコントロールの値は `KensaSnapshot` で受け取る。復元(Load_Text・HistoryLoad)は
「どのコントロールに何を入れるか」(`FormRestore`)と、続けて画面側で呼ぶ VBA の手続き(`calls`)を返す。
"""
from __future__ import annotations

import os
import re
import sqlite3
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta, time as dtime
from pathlib import Path
from typing import Any, Mapping, Optional, Sequence, Union

from . import ado_filter, datasource, flow, lines, paths, vbx
from .. import safe_io
from .model import Face, NORMAL, UiValues, bh_name, diag_name, is_thick, kou_name, lateral_name, row_format, visual_under_name, visual_up_name
from .flow import Message, Stop
from .store import (HOZON, HOZON_HEADERS, LOT_HISTORY, MIMI, MIMI_HEADERS, GridSheet, LocalStore, Value,
                    cell_text, excel_value, line_sheet)
from ..sqlite_toolkit import is_lock_error

FACES = ("S", "U", "C", "D", "E", "F")
WORK_SAVE = "作業保存"
WORKER_TIME = "WorkerTime"
DAILY = "Daily"
PACK_COUNT = "梱包枚数計算"
ABNORMAL = "異常ロット用"

RGB_WHITE_DEFAULT = 0x80000005           # &H80000005(既定のウィンドウ色)
RGB_RED_NG = 255 + 128 * 256 + 128 * 65536


def _rgb(r: int, g: int, b: int) -> int:
    return r + g * 256 + b * 65536


class VbaRuntimeError(Stop):
    """VBA の実行時エラー(ダイアログが出て止まる所)。D10 で状態は保つ。"""


def _err9(source: str) -> VbaRuntimeError:
    return VbaRuntimeError("実行時エラー '9': インデックスが有効範囲にありません。", source=source)


def _err13(source: str) -> VbaRuntimeError:
    return VbaRuntimeError("実行時エラー '13': 型が一致しません。", source=source)


# ------------------------------------------------------------------
# 画面の値
# ------------------------------------------------------------------
@dataclass
class Strain:
    """歪 1 面ぶん(定盤 CBx1_CBx2_Zx・ゲージ CBx3_CBx4_Gx)。表示中の値か、裏の面(Tag)の値。"""
    cb1: str = ""
    cb2: str = ""
    z: str = ""
    cb3: str = ""
    cb4: str = ""
    g: str = ""

    def plate(self) -> str:
        return f"{self.cb1}_{self.cb2}_{self.z}"

    def gauge(self) -> str:
        return f"{self.cb3}_{self.cb4}_{self.g}"


def _faces(default: Any = "") -> dict[str, Any]:
    return {f: default for f in FACES}


@dataclass
class KensaSnapshot:
    """帳票・保存に要る UF1 の値(コントロールの Caption / Text / Tag と状態)。

    grid.py(測定グリッドの状態モデル)と対応付けて作る。キーの面は "S","U","C","D","E","F"。
    """
    line_name: str = ""                     # LineName.Caption & LineNo.Caption(シート名の接尾辞・帳票の左上)
    lot: str = ""
    atu: str = ""; hab: str = ""; tak: str = ""                        # noqa: E702
    am: str = ""; ap: str = ""; hm: str = ""; hp: str = ""; tm: str = ""; tp: str = ""   # noqa: E702
    amk: str = ""; apk: str = ""; hmk: str = ""; hpk: str = ""; tmk: str = ""; tpk: str = ""  # noqa: E702
    zai: str = ""; tyu: str = ""; youto: str = ""; taik: str = ""; machin: str = ""     # noqa: E702
    resize_tag: str = ""                    # ReSize.Tag("有"/"無")
    lateral: str = ""; taik_text: str = ""; lateral_text: str = ""; flatnes: str = ""    # noqa: E702
    sumisei: str = ""; additional: str = ""                                              # noqa: E702
    measure: dict = field(default_factory=lambda: {f: [""] * 14 for f in FACES})   # 面 → 項目 1〜14 の Caption
    strain: dict = field(default_factory=lambda: {f: Strain() for f in FACES})     # 表示中の歪
    strain_tag: dict = field(default_factory=lambda: {f: Strain() for f in FACES}) # 裏の面(Tag)
    diag: dict = field(default_factory=_faces)          # DiagX.Caption(対角線差)
    diag_kou: dict = field(default_factory=_faces)      # DiagXKou.Text(直角度)
    ng: dict = field(default_factory=_faces)            # NGX.Caption(公差外判定ﾌﾗｸﾞ: "許容"/"保留"など)
    lateral_face: dict = field(default_factory=_faces)  # LateralX.Text
    visual_up: dict = field(default_factory=_faces)     # X_UpVisual.Caption
    visual_under: dict = field(default_factory=_faces)  # X_UnderVisual.Caption
    bh: dict = field(default_factory=_faces)            # XBH(S〜E)
    positions: dict = field(default_factory=_faces)     # 位置名 UF1.S.Caption など
    aorb: str = Face.A                     # AorB.Caption
    kind: str = ""                          # 厚板 / 薄板 / 少枚数 / 極厚 / スラブ(どれも選ばれていなければ "")
    atsuita_caption: str = ""               # 厚板.Caption(H / T)
    usuita_caption: str = ""                # 薄板.Caption
    yama: str = ""; yama_visible: bool = False                                           # noqa: E702
    start_time: str = ""; end_time: str = ""                                             # noqa: E702
    t1: str = ""; t2: str = ""; t3: str = ""                                             # noqa: E702
    worker_shift: str = ""                  # Worker.WorkeDayMax の結果("1"/"2"/"3"/"")
    data_loading: str = ""                  # dataLoding.Caption("Load" = 履歴呼び出し後)
    lgs_caption: str = ""; lzs_caption: str = ""; flatnes_text: str = ""                 # noqa: E702
    mg: str = ""; hahiz: str = ""; tahiz: str = ""; naka: str = ""; mimi: str = ""      # noqa: E702
    distortion_frame: str = ""; hiz_come: str = ""; huris: str = ""                     # noqa: E702
    ken_hbox: list = field(default_factory=lambda: [""] * 7)   # KenHBox1〜7
    yama_lb: str = ""; ken_count: str = ""; lateral_text_tag: str = ""; height_formula: str = ""  # noqa: E702
    delete_p: str = ""; delete_c: str = ""; ken_no: str = ""; ken_h: str = ""            # noqa: E702
    last_hot: str = ""; qf: str = ""; at_count: str = ""                                 # noqa: E702
    at_butt_visible: bool = False; inspe_butt_visible: bool = False; inspe_butt_caption: str = ""  # noqa: E702
    lateral_text_backcolor: int = 0; lateral_text_backstyle: int = 0                     # noqa: E702
    lot_source_lot: str = ""                # D28: 「控え(第2)から読む」を選んだロット(選んでいなければ "")

    def m(self, face: str, i: int) -> str:
        items = self.measure.get(face) or []
        return items[i - 1] if 0 < i <= len(items) else ""


def uf_array(snap: KensaSnapshot, i: int, face: str) -> str:
    """SaveLoad.UF_Array (SaveLoad:2135-2233) の 1 要素(項目 i、面 face)。"""
    if 1 <= i <= 10:
        return snap.m(face, i)
    if 11 <= i <= 14:
        return "" if face == "F" else snap.m(face, i)
    s = snap.strain.get(face, Strain())
    if i == 15:
        return s.plate()
    if i == 16:
        return s.gauge()
    if i == 17:
        return snap.diag.get(face, "")
    if i == 18:
        return snap.diag_kou.get(face, "")
    if i == 19:
        return snap.ng.get(face, "")
    if i == 20:
        return snap.lateral_face.get(face, "")
    raise IndexError(i)


def uf_save_array(snap: KensaSnapshot, i: int, face: str) -> str:
    """SaveLoad.UF_SaveArray (SaveLoad:2236-2286)。1〜10 測定、11 定盤歪、12 ゲージ歪。"""
    if 1 <= i <= 10:
        return snap.m(face, i)
    if i == 11:
        return snap.strain.get(face, Strain()).plate()
    if i == 12:
        return snap.strain.get(face, Strain()).gauge()
    raise IndexError(i)


def uf_tag_array(snap: KensaSnapshot) -> list[list[str]]:
    """OutTolerance.UFtag_Array(1, 6) (OutTolerance:2566-2596)。6 面 × [定盤, ゲージ](裏の面)。"""
    return [[snap.strain_tag.get(f, Strain()).plate(), snap.strain_tag.get(f, Strain()).gauge()] for f in FACES]


# ------------------------------------------------------------------
# 日付・比較(VBA の Date と Variant)
# ------------------------------------------------------------------
_BASE = datetime(1899, 12, 30)
_DATE_RE = re.compile(r"^\s*(\d{4})([/\-.])(\d{1,2})\2(\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?\s*$")
_TIME_RE = re.compile(r"^\s*(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s*$")


def _parse_date(v: Any, dot_is_date: bool) -> Optional[datetime]:
    if isinstance(v, datetime):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return _BASE + timedelta(days=float(v))
    if not isinstance(v, str):
        return None
    m = _DATE_RE.match(v)
    if m:
        if m.group(2) == "." and not dot_is_date:
            return None
        try:
            return datetime(int(m.group(1)), int(m.group(3)), int(m.group(4)), int(m.group(5) or 0),
                            int(m.group(6) or 0), int(m.group(7) or 0))
        except ValueError:
            return None
    m = _TIME_RE.match(v)
    if m:
        h, mi, s = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
        if h < 24 and mi < 60 and s < 60:
            return _BASE.replace(hour=h, minute=mi, second=s)
    return None


def vb_isdate(v: Any, *, dot_is_date: bool = False) -> bool:
    """VBA `IsDate`(この移植で出てくる形: yyyy/m/d [h:mm[:ss]]・h:mm・Date)。「.」区切りは U12。"""
    return _parse_date(v, dot_is_date) is not None


def vb_cdate(v: Any, *, dot_is_date: bool = False) -> datetime:
    """VBA `CDate`。読めなければ型エラー 13。"""
    d = _parse_date(v, dot_is_date)
    if d is None:
        raise vbx.TypeMismatch(f"CDate({v!r})")
    return d


def vb_cstr_date(d: datetime) -> str:
    """`CStr(Date)`(日本の地域設定: yyyy/mm/dd H:mm:ss。時刻 0:00:00 は日付だけ、日付 0 は時刻だけ)。"""
    has_time = (d.hour, d.minute, d.second) != (0, 0, 0)
    if d.date() == _BASE.date():
        return f"{d.hour}:{d.minute:02d}:{d.second:02d}"
    s = f"{d.year:04d}/{d.month:02d}/{d.day:02d}"
    return s + (f" {d.hour}:{d.minute:02d}:{d.second:02d}" if has_time else "")


_TOKENS = re.compile(r"yyyy|mm|dd|hh|h|nn|ss")


def vb_format_date(v: Any, pattern: str) -> str:
    """`Format$(v, "yyyy/mm/dd hh:mm")` など。v が空なら ""、日付にできない文字列はそのまま返す。

    `mm` は直前に h/hh があれば分、それ以外は月(VBA と同じ)。
    """
    if v is None or (isinstance(v, str) and v == ""):
        return ""
    d = _parse_date(v, False)
    if d is None:
        return str(v)
    out = []
    pos = 0
    after_hour = False
    for m in _TOKENS.finditer(pattern):
        out.append(pattern[pos:m.start()])
        t = m.group(0)
        if t == "yyyy":
            out.append(f"{d.year:04d}")
        elif t == "mm":
            out.append(f"{d.minute:02d}" if after_hour else f"{d.month:02d}")
        elif t == "dd":
            out.append(f"{d.day:02d}")
        elif t == "hh":
            out.append(f"{d.hour:02d}")
        elif t == "h":
            out.append(f"{d.hour}")
        elif t == "nn":
            out.append(f"{d.minute:02d}")
        elif t == "ss":
            out.append(f"{d.second:02d}")
        after_hour = t in ("h", "hh")
        pos = m.end()
    out.append(pattern[pos:])
    return "".join(out)


def variant_eq(a: Value, b: Value, b_type: str = vbx.VAR) -> bool:
    """セルの値(Variant)と b の `=`(vbx.compare_typed)。b が String 型の変数・文字列定数なら b_type=STR
    (そのときは文字で比べる。数のセルは CStr する)。Variant どうしなら数と文字列は等しくない。"""
    return vbx.compare_typed(a, b, vbx.VAR, b_type) == 0


def cmp_str_num(a: Any, b: Any) -> int:
    """String 型(Caption・Text)と数値の定数の比較(`.DiagS = 0`・`Select Case Yama.Text … Case 4`)。

    VBA は String を数に直して比べる(直せなければ型エラー 13)。compare_typed(STR, NUM) は文字で比べてしまうので、
    中身が同じになる VAR(文字列) と NUM の比較で行う(共通部品への依頼を最終報告に記載)。
    """
    return vbx.compare_typed(a, b, vbx.VAR, vbx.NUM)


def _vsplit(text: Any, sep: str) -> list[str]:
    """VBA `Split`(空文字は要素 0 個の配列)。"""
    s = cell_text(text) if not isinstance(text, str) else text
    return [] if s == "" else s.split(sep)


# ------------------------------------------------------------------
# MainDataSave: 保存用 / 耳用 の 1 行
# ------------------------------------------------------------------
def visual_save_join(a: str, b: str, *, face: str = "", a_e: str = "", b_e: str = "") -> str:
    """MainDataSave の外観 A/B の結合 (SaveLoad:4246-4294)。"A_xxx,B_yyy"。

    B17: C 面だけ B のみのとき先頭の空白が無い(",B_yyy")。F 面は 2 つ目の条件で E 面(a_e, b_e)を見る。
    """
    lead = " "
    if a != "" and b != "":
        return "A_" + a + ",B_" + b
    if a != "" and b == "":
        return "A_" + a + ","
    if a == "" and b != "":
        return lead + ",B_" + b
    return ""


def strain_ab(snap: KensaSnapshot) -> list[list[str]]:
    """MainDataSave の歪 "A:<z>,B:<z>" (SaveLoad:4300-4345)。6 面 × [定盤, ゲージ]。"__" は空。"""
    shown = [[uf_array(snap, 15, f), uf_array(snap, 16, f)] for f in FACES]
    tags = uf_tag_array(snap)
    if snap.aorb == Face.A:
        a, b = shown, tags
    elif snap.aorb == Face.B:
        a, b = tags, shown
    else:
        raise _err13("SaveLoad:4321")            # hizumA が Empty のまま LBound → 型エラー
    out = []
    for i in range(6):
        row = []
        for k in range(2):
            av = "A:" + a[i][k] if a[i][k] != "__" else ""
            bv = ",B:" + b[i][k] if b[i][k] != "__" else ""
            row.append(av + bv)
        out.append(row)
    return out


@dataclass
class RegisterResult:
    """測定DB への登録の結果(adoSQL)。mode: ADDNEW / UPDATE / ERROR。"""
    mode: str
    rowcount: int = 0
    messages: list[Message] = field(default_factory=list)


@dataclass
class SaveResult:
    sheet: str
    row: int
    values: list[Value]                      # AdoTemp(0 = ライン名、1〜 = シートの A 列〜)
    db_table: str
    db: Optional[RegisterResult] = None


def _worker_last(snap: KensaSnapshot, source: str) -> str:
    """`Worker(UBound(Worker))`。WorkeDayMax が "" なら Empty で型エラー、T が空なら添字エラー。"""
    t_by = {"1": snap.t1, "2": snap.t2, "3": snap.t3}
    if snap.worker_shift not in t_by:
        raise _err13(source)
    parts = _vsplit(t_by[snap.worker_shift], "_")
    if not parts:
        raise _err9(source)
    return parts[-1]


def main_data_save(store: LocalStore, snap: KensaSnapshot, *, db_path: Optional[Union[str, Path]] = None,
                   register: bool = True) -> SaveResult:
    """SaveLoad.MainDataSave (SaveLoad:4216-4569)。印刷の後に 1 行を保存用(極厚は耳用)へ書き、測定DB に登録する。

    シートへ書いてから DB に登録する(VBA と同じく、登録に失敗してもシートの行は残る)。
    """
    extreme = snap.kind == "極厚"
    hozon = store.ensure_sheet(HOZON)
    mimi = store.ensure_sheet(MIMI)
    sh = mimi if extreme else hozon

    vis = {f: visual_save_join(snap.visual_up.get(f, ""), snap.visual_under.get(f, ""), face=f,
                               a_e=snap.visual_up.get("E", ""), b_e=snap.visual_under.get("E", ""))
           for f in FACES}
    vis_item = [vis[f] for f in "SUCDE"]
    bh_data = [snap.bh.get(f, "") for f in "SUCDE"]
    hz = strain_ab(snap)

    if extreme:
        ken = [[uf_array(snap, i, f) for i in range(1, 17)] for f in "SUCDE"]
    else:
        ken = [[uf_save_array(snap, i, f) for i in range(1, 13)] for f in "SUCDE"]
    one_ken = " ".join(" ".join(row) for row in ken).split(" ")      # Join → Split(" ")
    lateral = [uf_array(snap, 20, f) for f in "SUCDE"]
    right_angle = [uf_array(snap, 18, f) for f in "SUCDE"]
    six = ([uf_save_array(snap, i, "F") for i in range(1, 11)]
           + [hz[5][0], hz[5][1], uf_array(snap, 20, "F"), vis["F"], uf_array(snap, 18, "F")])
    hiki = [snap.atu, snap.hab, snap.tak, snap.am, snap.ap, snap.hm, snap.hp, snap.tm, snap.tp]
    period = (vb_format_date(snap.start_time, "yyyy/mm/dd hh:mm") + "～"
              + vb_format_date(snap.end_time, "yyyy/mm/dd hh:mm"))

    n = sh.last_row(1) + 1
    if extreme:                                                     # SaveLoad:4416-4452
        sh.write(n, 2, snap.lot, "O")
        sh.write_row(n, 3, hiki, "O")
        sh.write_row(n, 12, one_ken, "S")
        sh.write(n, 98, "耳付き")
        sh.write(n, 1, period)
        sh.write(n, 92, _worker_last(snap, "SaveLoad:4431"))
        sh.write(n, 93, snap.resize_tag)
        sh.write(n, 94, snap.zai, "O")
        sh.write(n, 95, snap.tyu, "O")
        sh.write(n, 96, snap.youto, "O")
        sh.write_row(n, 99, vis_item)
        sh.write_row(n, 104, bh_data, "O")
        if snap.data_loading == "Load":
            sh.write(n, 97, "〆")
    else:                                                           # SaveLoad:4454-4525
        sh.write(n, 2, snap.lot, "O")
        sh.write_row(n, 3, hiki, "O")
        sh.write_row(n, 12, one_ken, "S")
        sh.write_row(n, 72, lateral)
        sh.write_row(n, 95, right_angle)
        sh.write_row(n, 100, six)
        if snap.kind == "厚板":
            freq = snap.atsuita_caption
        elif snap.kind == "薄板":
            freq = snap.usuita_caption + ":" + snap.yama
        elif snap.kind == "少枚数":
            freq = snap.atsuita_caption + ":" + "少枚数"
        else:
            freq = "ｽﾗﾌﾞ"
        sh.write(n, 83, freq)
        sh.write(n, 1, period)
        sh.write(n, 77, _worker_last(snap, "SaveLoad:4487"))
        sh.write(n, 78, snap.resize_tag)
        sh.write(n, 79, snap.zai, "O")
        sh.write(n, 80, snap.tyu, "O")
        sh.write(n, 81, snap.youto, "O")
        if snap.lgs_caption == "ﾌﾘｰｿﾘ":
            sh.write(n, 84, "〆")
        sh.write_row(n, 85, vis_item)
        sh.write_row(n, 90, bh_data, "O")                       # B17: VBA はコメントアウトで書かない
        for i, col in enumerate((22, 34, 46, 58, 70)):
            sh.write_row(n, col, hz[i])
        if snap.data_loading == "Load":
            sh.write(n, 82, "〆")

    save_allowed(store, sh.name, period, snap.lot, snap.ng)
    end_r = sh.last_row(1)                                           # SaveLoad:4535-4561
    end_c = sh.last_col(1)
    values: list[Value] = [snap.line_name] + [sh.get(end_r, c) for c in range(1, end_c + 1)]
    key = cell_text(values[1])
    fil = "[検査開始終了] ='" + key + ("'")
    table = "耳付検査ﾃﾞｰﾀ" if extreme else "検査ﾃﾞｰﾀ"
    result = SaveResult(sheet=sh.name, row=n, values=values, db_table=table)
    if register:
        path = Path(db_path) if db_path else paths.to_local(paths.db_file("梱包検査ﾃﾞｰﾀ"))
        result.db = register_measurement(path, table, values, fil)
    return result


# ------------------------------------------------------------------
# 測定DB(梱包検査ﾃﾞｰﾀ。D4 で SQLite)
# ------------------------------------------------------------------
#: 表の列(VBA は列名でなく**位置**で書く: 1 列目 = ライン名、2 列目以降 = シートの A 列〜)。
#: 実 DB(M1)が届いたら名前を差し替える。登録は既にある表の列の並び(PRAGMA table_info)で位置どおりに書く
MEASURE_TABLES: dict[str, tuple[str, ...]] = {
    "検査ﾃﾞｰﾀ": ("ライン名",) + HOZON_HEADERS,
    "耳付検査ﾃﾞｰﾀ": ("ライン名",) + MIMI_HEADERS,
}
MEASURE_BUSY_TIMEOUT_MS = 10_000
MEASURE_MAX_RETRY = 3
MEASURE_RETRY_WAIT_SEC = 1.0


def _qi(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def ensure_measure_db(path: Union[str, Path]) -> Path:
    """測定DB の表を作る(無いときだけ)。列はすべて TEXT(VBA は全列 CStr の文字列で書く)。"""
    p = Path(path)
    # 2026-11-08: 「無い」と確かめられたときだけ作る。共有フォルダに届かない・一瞬見えないだけなら、
    # フォルダも空の DB も作らずに止める(あるファイルは mode=rw で開くので作らない)
    if p.exists():
        conn = safe_io.connect_existing(p, timeout=MEASURE_BUSY_TIMEOUT_MS / 1000)
    elif safe_io.confirm_absent(p, local=not str(p).startswith(("\\\\", "//"))):
        p.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(p), timeout=MEASURE_BUSY_TIMEOUT_MS / 1000)
    else:
        raise Stop(f"測定DB の場所に届かないので保存できません: {p}", source="SQL_追加")
    try:
        with conn:
            for table, cols in MEASURE_TABLES.items():
                body = ", ".join(f"{_qi(c)} TEXT" for c in cols)
                conn.execute(f"CREATE TABLE IF NOT EXISTS {_qi(table)} ({body})")
    finally:
        conn.close()
    return p


class _SqlError(Exception):
    def __init__(self, number: int, description: str) -> None:
        super().__init__(description)
        self.number = number
        self.description = description


def register_measurement(path: Union[str, Path], table: str, values: Sequence[Value], key_filter: str, *,
                         match_line: bool = True,
                         max_retry: int = MEASURE_MAX_RETRY,
                         retry_wait_sec: float = MEASURE_RETRY_WAIT_SEC) -> RegisterResult:
    """SQL_追加.adoSQL(MODE="ADDNEW", skipmsgflag=True, keySQLfilter) (SQL_追加:525-704)。

    - 表の列を位置で対応させる(列 j ← CStr(values[j]))。表の列が values より多いと添字エラーで取り消し
    - `key_filter`(ADO Filter の書き方)に合い、1 列目(ライン名)も同じ行があれば UPDATE(ライン名は書かない)、
      無ければ追加(`match_line=False` で VBA と同じく key_filter だけで照合)
    - 1 回の登録 = 1 トランザクション(BEGIN IMMEDIATE)。ロック中は待って(busy_timeout)、それでもだめなら再試行
    - 失敗したら取り消して、VBA の msg.Box と同じ「エラー番号 / エラーの内容」を返す(例外にはしない)
    """
    p = Path(path)
    if not p.is_file():
        raise Stop(f"DB ファイルが見つかりません: {p}", source="SQL_追加:297")   # VBA は End
    texts = [cell_text(v) for v in values]
    retry = 0
    while True:
        # 2026-11-08: is_file() のあとで一瞬見えなくなっても空の DB を作らない(あるファイルだけを開く)
        try:
            conn = safe_io.connect_existing(p, timeout=MEASURE_BUSY_TIMEOUT_MS / 1000, isolation_level=None)
        except sqlite3.OperationalError as e:
            raise Stop(f"DB ファイルを開けません: {p}({e})", source="SQL_追加:297") from e
        try:
            conn.execute(f"PRAGMA busy_timeout = {MEASURE_BUSY_TIMEOUT_MS}")
            cols = [r[1] for r in conn.execute(f"PRAGMA table_info({_qi(table)})")]
            if not cols:
                raise Stop(f"表 {table} がありません: {p}", source="SQL_追加:529")
            jcnt = len(cols)
            conn.execute("BEGIN IMMEDIATE")
            try:
                matched = _matching_rowids(conn, table, cols, key_filter)
                if match_line and matched:
                    # 2026-11-08 のご判断: 照合キーにラインも足す(「開始～終了」が分まで同じだと、別ラインの行を
                    # 上書きしていた。VBA は検査開始終了だけで照合)。1 列目(ライン名)が同じ行だけを同じ検査とみなす
                    marks = ", ".join("?" for _ in matched)
                    same = {r[0] for r in conn.execute(
                        f"SELECT rowid FROM {_qi(table)} WHERE rowid IN ({marks}) AND {_qi(cols[0])} = ?",
                        list(matched) + [texts[0]])}
                    matched = [m for m in matched if m in same]
                if not matched:
                    if len(texts) < jcnt:
                        raise _SqlError(9, "インデックスが有効範囲にありません。")
                    conn.execute(f"INSERT INTO {_qi(table)} ({', '.join(_qi(c) for c in cols)}) "
                                 f"VALUES ({', '.join('?' for _ in cols)})", texts[:jcnt])
                    mode, count = "ADDNEW", 1
                else:
                    if len(texts) < jcnt:
                        raise _SqlError(9, "インデックスが有効範囲にありません。")
                    upd = texts[1:jcnt]
                    if "'" in key_filter.split("=", 1)[-1].strip()[1:-1]:
                        # VBA は値を '…' で囲んで SQL に埋め込む(エスケープなし)→ 構文エラー
                        # B75 を直すと、値はパラメータで渡す(下の UPDATE)ので値の ' は通る。
                        # 照合キー(keySQLfilter)の ' は ADO の Filter でも通らないので直した側でも同じ
                        raise _SqlError(-2147217900, "クエリ式の構文エラー(値に ' がある)")
                    sets = ", ".join(f"{_qi(c)} = ?" for c in cols[1:jcnt])
                    marks = ", ".join("?" for _ in matched)
                    cur = conn.execute(f"UPDATE {_qi(table)} SET {sets} WHERE rowid IN ({marks})",
                                       upd + list(matched))
                    mode, count = "UPDATE", cur.rowcount
                conn.execute("COMMIT")
                return RegisterResult(mode=mode, rowcount=count)
            except _SqlError as e:
                conn.execute("ROLLBACK")
                return RegisterResult(mode="ERROR", messages=[_db_error_message(e.number, e.description)])
            except sqlite3.Error:
                if conn.in_transaction:
                    conn.execute("ROLLBACK")
                raise
        except sqlite3.Error as e:
            if is_lock_error(e) and retry < max_retry:
                retry += 1
                time.sleep(retry_wait_sec * retry)
                continue
            return RegisterResult(mode="ERROR", messages=[_db_error_message(-2147467259, str(e))])
        finally:
            conn.close()


def _db_error_message(number: int, description: str) -> Message:
    return Message("エラー番号  :" + str(number) + "\r\n" + "エラーの内容:" + description, level=flow.WARN,
                   source="SQL_追加:697")


def _matching_rowids(conn: sqlite3.Connection, table: str, cols: list[str], key_filter: str) -> list[int]:
    """ADO の Filter に合う行(adoSQL L546。比較の規則は ado_filter)。"""
    refs = [c for c in re.findall(r"\[([^\]]+)\]", key_filter) if c in cols]
    if not refs:
        return []
    test = ado_filter.compile_filter(key_filter)
    sel = ", ".join(_qi(c) for c in refs)
    out = []
    for row in conn.execute(f"SELECT rowid, {sel} FROM {_qi(table)} ORDER BY rowid"):
        if test({c: datasource.as_text(v) for c, v in zip(refs, row[1:])}):
            out.append(row[0])
    return out


# ------------------------------------------------------------------
# 作業中保存(Posting)と歪の A/B
# ------------------------------------------------------------------
def posting(store: LocalStore, snap: KensaSnapshot, need_right: Optional[bool]) -> GridSheet:
    """SaveLoad.Posting (SaveLoad:1981-2133)。作業保存_<L> を作り直して全部書く。

    `need_right`: Lot情報履歴 の引当から決めた「直角度が要るか」(JudgeLotAllocations の NeedRight)。
    None はロットが Lot情報履歴 に無い(VBA はそこで抜けるので R 列を消さない)。
    """
    sh: GridSheet = store.add_sheet(line_sheet(WORK_SAVE, snap.line_name))  # type: ignore[assignment]
    sh.write_row(1, 33, [snap.mg, snap.hahiz, snap.tahiz, snap.naka, snap.mimi, snap.distortion_frame,
                         snap.hiz_come])
    sh.write_row(1, 1, [snap.lot, snap.atu, snap.hab, snap.tak, snap.am, snap.ap, snap.hm, snap.hp, snap.tm,
                        snap.tp, snap.amk, snap.apk, snap.hmk, snap.hpk, snap.tmk, snap.tpk, snap.zai,
                        snap.tyu, snap.youto, snap.taik, snap.machin, snap.resize_tag, snap.lateral,
                        snap.taik_text, snap.lateral_text, snap.flatnes, snap.sumisei, snap.additional])
    last_col = 29
    if snap.kind in ("厚板", "薄板", "少枚数", "極厚", "スラブ"):
        sh.write(1, last_col, snap.kind)
    thin = variant_eq(sh.get(1, last_col), "薄板", vbx.STR)
    if thin and snap.at_butt_visible:
        sh.write(1, last_col + 1, vbx.mid(snap.at_count, 1, 4))
    if thin and snap.inspe_butt_visible:
        sh.write(1, last_col + 2, snap.inspe_butt_caption)
    if thin and snap.yama != "":
        sh.write(1, last_col + 3, snap.yama)
    for r, f in enumerate(FACES):
        sh.write_row(2 + r, 1, [uf_array(snap, i, f) for i in range(1, 15)])
        sh.write_row(2 + r, 17, [uf_array(snap, i, f) for i in range(17, 21)], "S")
    reflect_strain(store, snap, snap.aorb[-1:] if snap.aorb else "", "保存")
    sh.write(8, 1, snap.delete_p)
    sh.write(8, 2, snap.delete_c)
    sh.write(9, 1, snap.ken_no)
    sh.write(16, 1, snap.ken_h)
    sh.write(17, 1, snap.last_hot)
    sh.write(17, 2, snap.qf)
    for r, f in enumerate(FACES):
        sh.write(10 + r, 1, snap.visual_up.get(f, ""))
        sh.write(10 + r, 2, snap.visual_under.get(f, ""))
    for r, f in enumerate("SUCDE"):
        sh.write(10 + r, 3, snap.bh.get(f, ""))
    sh.write(9, 2, snap.yama_lb)
    sh.write_row(9, 3, list(snap.ken_hbox[:7]))
    sh.write(1, 40, snap.ken_count)
    sh.write(1, 41, snap.huris)
    sh.write(1, 42, snap.lateral_text_tag)
    sh.write(9, 10, snap.height_formula)
    if snap.lot_source_lot != "":                      # D28(VBA には無い。VBA の列と重ならない 1 行目 60 列目)
        sh.write(LOT_SOURCE_CELL[0], LOT_SOURCE_CELL[1], snap.lot_source_lot)
    if need_right is None:
        return sh
    if not need_right:
        for r in range(2, 8):
            sh.set(r, 18, None)
    return sh


#: D28: 作業保存_<L> に「控え(第2)から読む」を選んだロットを置くセル(VBA の Posting は 1 行目 42 列目まで)
LOT_SOURCE_CELL = (1, 60)


def load_lot_source(store: LocalStore, lname: str) -> str:
    """作業保存_<L> から「控え(第2)から読む」を選んだロット(D28)。保存のロット(A1)と同じときだけ返す。無ければ ""。"""
    name = line_sheet(WORK_SAVE, lname)
    if lname == "" or not store.exists(name):
        return ""
    sh = store.sheet(name)
    lot = cell_text(sh.get(*LOT_SOURCE_CELL))
    return lot if lot != "" and lot == cell_text(sh.get(1, 1)) else ""


@dataclass
class FormRestore:
    """画面へ戻す値。values: コントロール名 → Caption/Text、tags: → Tag、back_colors: → BackColor など。"""
    found: bool = True
    values: dict[str, str] = field(default_factory=UiValues)
    tags: dict[str, str] = field(default_factory=dict)
    back_colors: dict[str, int] = field(default_factory=dict)
    back_styles: dict[str, int] = field(default_factory=dict)
    list_index: dict[str, Any] = field(default_factory=dict)
    option: str = ""                       # 選ぶ OptionButton(厚板 など)
    calls: list[tuple] = field(default_factory=list)   # 画面側で続けて呼ぶ VBA の手続き(名前, 引数…)
    messages: list[Message] = field(default_factory=list)

    def set(self, name: str, value: Any) -> None:
        self.values[name] = cell_text(value) if not isinstance(value, str) else value


def reflect_strain(store: LocalStore, snap: KensaSnapshot, n: str, default_b: str = "") -> Optional[FormRestore]:
    """SaveLoad.ReflectStrain (SaveLoad:1855-1901)。作業保存_<L> の O:P(A 面 2〜7 行、B 面 8〜13 行)。

    default_b="保存": いま表示中の面へ表示中の値、反対の面へ Tag の値を書く。
    default_b="": A/B の切替ボタン。表示していた面を書いてから反対の面を読み出す(LosdStrain の結果を返す)。
    """
    name = line_sheet(WORK_SAVE, snap.line_name)
    if not store.exists(name):
        return None
    sh = store.sheet(name)
    shown = [[uf_array(snap, 15, f), uf_array(snap, 16, f)] for f in FACES]
    if default_b == "":
        if n == "B":
            _write_block(sh, 2, shown, "S")
        elif n == "A":
            _write_block(sh, 8, shown, "S")
        else:
            return None
        return losd_strain(n, sh.read_range(2, 15, 7, 16), sh.read_range(8, 15, 13, 16))
    if default_b == "保存":
        tags = uf_tag_array(snap)
        if n == "B":
            _write_block(sh, 8, shown, "S")
            _write_block(sh, 2, tags, "V")
        elif n == "A":
            _write_block(sh, 2, shown, "S")
            _write_block(sh, 8, tags, "V")
    return None


def _write_block(sh, top: int, rows: list[list[str]], kind: str) -> None:
    for i, row in enumerate(rows):
        sh.write_row(top + i, 15, row, kind)


def losd_strain(temp: str, a_rows: Sequence[Sequence[Value]], b_rows: Sequence[Sequence[Value]],
                out: Optional[FormRestore] = None) -> FormRestore:
    """SaveLoad.LosdStrain (SaveLoad:1903-1979)。呼び出す面を表示へ、反対の面を Tag へ(Split_Z / Split_G)。"""
    out = out or FormRestore()
    data, opp = (b_rows, a_rows) if temp == "B" else (a_rows, b_rows)
    if temp not in ("A", "B"):
        raise _err13("SaveLoad:1923")                          # データ が Empty のまま
    for i, kind in ((0, "Z"), (1, "G")):
        for a, f in enumerate(FACES):
            za = _vsplit(data[a][i], "_")
            zb = _vsplit(opp[a][i], "_")
            if not za:
                _split_0(out, f, kind)
            elif len(za) - 1 == 2:
                _split(out, f, kind, za, zb, "SaveLoad:1943")
    return out


def _names(face: str, kind: str) -> tuple[str, str, str]:
    if kind == "Z":
        return f"CB{face}1", f"CB{face}2", f"Z{face}"
    return f"CB{face}3", f"CB{face}4", f"G{face}"


def _split_0(out: FormRestore, face: str, kind: str, b: Optional[list[str]] = None, source: str = "") -> None:
    """Split_Z_0 / Split_G_0 (SaveLoad:2394-2407)。表示を空にし、b があれば Tag へ。"""
    for nm in _names(face, kind):
        out.values[nm] = ""
    if b:
        _split_tag(out, face, kind, b, source)


def _split(out: FormRestore, face: str, kind: str, a: list[str], b: Optional[list[str]], source: str) -> None:
    """Split_Z / Split_G (SaveLoad:2408-2425)。"""
    if a:
        n1, n2, n3 = _names(face, kind)
        out.values[n1] = a[0]
        if len(a) < 2:
            raise _err9(source)
        out.values[n2] = a[1]
        out.values[n3] = a[-1]
    if b:
        _split_tag(out, face, kind, b, source)


def _split_tag(out: FormRestore, face: str, kind: str, b: list[str], source: str) -> None:
    n1, n2, n3 = _names(face, kind)
    out.tags[n1] = b[0]
    if len(b) < 2:
        raise _err9(source)
    out.tags[n2] = b[1]
    out.tags[n3] = b[-1]


# ------------------------------------------------------------------
# 作業再開(Load_Text)
# ------------------------------------------------------------------
def _fmt(v: Value, pattern: str) -> str:
    """`Format(セルの値, pattern)`(Empty は ""、数にならない文字列はそのまま)。"""
    return vbx.fmt(v, pattern)


def _fmt_val(v: Value, pattern: str) -> str:
    """`Format(Val(セル), pattern)`。"""
    return vbx.fmt(vbx.val(cell_text(v)), pattern)


def _is_zero_caption(text: str) -> bool:
    """`UF1("S1") = 0`(Caption の文字列を数にして比べる。数にならなければ型エラー)。"""
    return cmp_str_num(text, 0) == 0


def load_text(store: LocalStore, lname: str, n: str = "", *, current_option: str = "",
              lot_now: Optional[str] = None) -> Optional[FormRestore]:
    """SaveLoad.Load_Text (SaveLoad:2501-2992)。作業保存_<L> から画面を戻す。シートが無ければ None。

    `n`: ""(通常)/ "表示"(進捗バー)/ "ReLoad"(ReSize.Tag を "無" にする)。
    `current_option`: いま選ばれている検査形態(シートに記録が無いときはそのまま)。
    画面の部品の描き直し(Atu_OPselect・LSelect など)は `calls` に順に入れる。
    """
    name = line_sheet(WORK_SAVE, lname)
    if not store.exists(name):
        return None
    sh = store.sheet(name)
    out = FormRestore()
    if n == "表示":
        out.calls.append(("UFProgress", "SaveDate_Loading...."))
    kensa = sh.read_range(1, 1, 1, 28)[0]
    data = sh.read_range(2, 1, 9, 20)
    tb = sh.read_range(1, 33, 1, 39)[0]
    vi = sh.read_range(10, 1, 15, 2)
    bh = sh.read_range(10, 3, 14, 3)
    for i, f in enumerate(FACES):
        out.set(visual_up_name(f), vi[i][0])
        out.set(visual_under_name(f), vi[i][1])
        if f != "F":
            out.set(bh_name(f), bh[i][0])
    ken_h = sh.read_range(9, 3, 9, 9)[0]
    for a in range(6):
        out.set(f"KenHBox{a + 1}", _fmt(ken_h[a], "0.0"))
    out.set("KenHBox7", ken_h[6])
    y = sh.get(9, 2)
    if y is not None and cell_text(y) != "":
        if vbx.is_numeric(y):
            out.list_index["YamaLB"] = vbx.clng(vbx.val(cell_text(y)) - 1)
        else:
            out.list_index["YamaLB"] = "last"                   # ListCount - 1
    out.set("KenCount", sh.get(1, 40))
    out.set("HURIS", sh.get(1, 41))
    out.tags["LateralText"] = cell_text(sh.get(1, 42))
    out.set("高さ計算式", sh.get(9, 10))
    out.set("KenH", sh.get(16, 1))
    out.set("LastHot", sh.get(17, 1))
    out.set("QF", sh.get(17, 2))

    k = kensa
    out.set("Lot", k[0])
    for nm, idx, pat in (("ATU", 1, "0.000"), ("Hab", 2, "0.0"), ("Tak", 3, "0.0"), ("Am", 4, "0.000"),
                         ("Ap", 5, "0.000"), ("Hm", 6, "0.0"), ("Hp", 7, "0.0"), ("Tm", 8, "0.0"),
                         ("Tp", 9, "0.0"), ("AmK", 10, "0.000"), ("ApK", 11, "0.000"), ("HmK", 12, "0.0"),
                         ("HpK", 13, "0.0"), ("TmK", 14, "0.0"), ("TpK", 15, "0.0")):
        out.set(nm, _fmt(k[idx], pat))
    for nm, idx in (("Zai", 16), ("Tyu", 17), ("Youto", 18), ("TAIK", 19), ("Machin", 20), ("Lateral", 22)):
        out.set(nm, k[idx])
    out.set("TaikText", _fmt(k[23], "0.0"))
    out.set("LateralText", _fmt(k[24], "0.0"))
    out.set("flatnes", k[25])
    out.set("Sumisei", k[26])
    out.set("MG", _fmt(tb[0], "0.0"))
    out.set("HaHiz", _fmt(tb[1], "0.0"))
    out.set("TaHiz", _fmt(tb[2], "0.0"))
    naka, mimi = cell_text(tb[3]), cell_text(tb[4])
    out.set("Naka", vbx.fmt(naka, "0.0") if vbx.is_numeric(naka) else naka)
    out.set("Mimi", vbx.fmt(mimi, "0.0") if vbx.is_numeric(mimi) else mimi)
    out.set("DistortionFrame", tb[5])
    out.set("HizCome", tb[6])
    if cell_text(k[26]) != "":
        out.back_styles["LateralText"] = 1
        out.back_colors["LateralText"] = _rgb(255, 153, 0)
    if cell_text(k[27]) != "":
        out.set("Additional", k[27])
        head = vbx.mid(cell_text(k[27]), 1, 3)
        flag = _rgb(51, 204, 204)
        if head == "対角ﾗ":
            out.back_styles.update({"TaikText": 1, "LateralText": 1})
            out.back_colors.update({"TaikText": flag, "LateralText": flag})
        elif head == "対角追":
            out.back_styles["TaikText"] = 1
            out.back_colors["TaikText"] = flag
        elif head == "ﾗﾃﾗ":
            out.back_styles["LateralText"] = 1
            out.back_colors["LateralText"] = flag
    out.tags["ReSize"] = "無" if n == "ReLoad" else cell_text(k[21])
    out.set("ResizText", out.tags["ReSize"])

    for i in range(1, 21):
        for fi, f in enumerate(FACES):
            v = data[fi][i - 1]
            if is_thick(i):
                out.set(f"{f}{i}", _fmt_val(v, "0.000"))
            elif i in NORMAL.width:
                out.set(f"{f}{i}", _fmt_val(v, "0.0"))
            elif i in NORMAL.length or i in NORMAL.diag:
                out.set(f"{f}{i}", _fmt_val(v, "0.0"))
            elif 11 <= i <= 14:
                if f != "F":
                    out.set(f"{f}{i}", _fmt_val(v, "0.0"))
            elif i == 17:
                out.set(diag_name(f), _fmt_val(v, "0.0"))
            elif i == 18:
                if v is not None and cell_text(v) != "" and vbx.is_numeric(v):
                    out.set(kou_name(f), _fmt_val(v, "0.0"))
                elif v is None or cell_text(v) == "":
                    out.set(kou_name(f), "")
            elif i == 19:
                out.set(f"NG{f}", v)
            elif i == 20:
                out.set(lateral_name(f), _fmt(v, "0.0"))
    losd_strain("A", sh.read_range(2, 15, 7, 16), sh.read_range(8, 15, 13, 16), out)
    for i in range(1, 15):
        for f in "SUCDE":
            if _is_zero_caption(out.values[f"{f}{i}"]):
                out.values[f"{f}{i}"] = ""
                out.back_colors[f"{f}{i}"] = RGB_WHITE_DEFAULT
    for i in range(1, 11):
        if _is_zero_caption(out.values[f"F{i}"]):
            out.values[f"F{i}"] = ""
            out.back_colors[f"F{i}"] = RGB_WHITE_DEFAULT
    for f in FACES:
        if out.values[f"{f}10"] == "" and _is_zero_caption(out.values[diag_name(f)]):
            out.values[diag_name(f)] = ""
    out.calls.append(("Atu_OPselect", vbx.val(cell_text(sh.get(1, 2)))))

    ken_item = [cell_text(v) for v in sh.read_range(1, 29, 1, 32)[0]]
    at_sl, two_four, above_below, temp_yama = ken_item
    ken_no = vbx.left(cell_text(data[7][0]), 1)
    out.set("Yama", temp_yama)
    option = current_option
    if at_sl == "厚板":
        option = "厚板"
        if ken_no in ("H", "T"):
            out.values["厚板.Caption"] = ken_no
    elif at_sl == "薄板":
        option = "薄板"
        if ken_no in ("L", "N", "K", "T", "S", "F"):
            out.values["薄板.Caption"] = ken_no
    elif at_sl in ("少枚数", "極厚", "スラブ"):
        option = at_sl
    out.option = option
    if out.values["Sumisei"] != "" and option not in ("極厚", "スラブ"):
        out.calls.append(("LSelect", "住精"))
    if out.tags["LateralText"] == "有" or out.values["Sumisei"] != "":
        out.calls.append(("LSelectLateral", ken_no))
    out.calls.append(("Six_Display", "on" if temp_yama == "4" else "off", "Yama.Visible が偽なら off"))
    if two_four == "厚2ヵ所":
        out.calls.append(("AtCountSelect", 2))
    elif two_four == "厚4ヵ所":
        out.calls.append(("AtCountSelect", 4))
    if above_below == "検入≧26":
        out.calls.append(("InspectionSelect", "非表示"))
    elif above_below == "検入≦25":
        out.calls.append(("InspectionSelect", "表示"))
    if option == "スラブ":
        out.calls.append(("SlabSet",))
    out.set("KenNo", data[7][0])
    out.set("DeleteP", data[6][0])
    out.set("DeleteC", data[6][1])
    out.calls.append(("SmallNumberCount判定", "ATU<>\"\" のとき: KenHBox7=\"\" なら Round(CalculateSheets,1)、"
                      "KenCount<>\"\" ならその値。5 まで。厚板.Caption=\"H\" かつ 少枚数 なら SmallNumberCount"))
    if out.values["ATU"] != "":
        out.calls.append(("DrawListBoxContent", out.values["ATU"]))

    lot = out.values["Lot"] if lot_now is None else lot_now
    if not store.exists(LOT_HISTORY):
        return out
    hist = store.sheet(LOT_HISTORY)
    lr = hist.last_row(1)
    rows = hist.read_range(2, 1, lr, 26)
    hit = next((r for r in rows if variant_eq(r[2], lot)), None)
    if hit is None:
        return out
    vc = hit[17]
    if vbx.compare_typed(vc, 0, vbx.VAR, vbx.NUM) == 0:
        out.calls.append(("VC_Clear",))
    if vbx.compare_typed(vc, 1, vbx.VAR, vbx.NUM) == 0:
        out.calls.append(("VC_Picture",))
    out.calls.append(("InitialSetting", "サイズ保存", "UFsizeWidth<>0 のとき"))
    wt = line_sheet(WORKER_TIME, lname)
    if store.exists(wt):
        out.set("StartTime", vb_format_date(store.sheet(wt).get(1, 2), "yyyy/mm/dd hh:mm:ss"))
    ab = line_sheet(ABNORMAL, lname)
    if store.exists(ab) and variant_eq(store.sheet(ab).get(1, 1), out.values["Lot"]):
        out.back_colors["IzyouFl"] = _rgb(255, 0, 0)
    out.calls.append(("Conductiv_Check",))
    return out


# ------------------------------------------------------------------
# 履歴の呼び出し(HistoryLoad)と UF管理
# ------------------------------------------------------------------
def _ab_strain(cell: Value, state: dict[str, str]) -> tuple[str, str]:
    """HistoryLoad の "A:x,B:y" の分け方。VBA は前の面の値を変数に残したまま次の面へ進む(state で再現)。"""
    s = cell_text(cell)
    parts = s.split(",")
    if "A:" in parts[0]:
        state["a"] = parts[0].replace("A:", "")
    elif "B:" in parts[-1]:
        state["a"] = ""
    if "B:" in parts[-1]:
        state["b"] = parts[-1].replace("B:", "")
    return state["a"], state["b"]


def history_load(store: LocalStore, hist_lot: str, temp_date: str = "", *, mimi: bool = False,
                 current: Optional[Mapping[str, str]] = None, lname: str = "",
                 kanri_positions: Optional[Mapping[str, str]] = None, aorb: str = Face.A,
                 row_no: Optional[int] = None) -> FormRestore:
    """SaveLoad.HistoryLoad (SaveLoad:3409-4049)。保存用(耳用)の行を UF1 に戻す。

    `row_no`(2026-11-06): 履歴画面で選んだ行の番号。渡せばその行を戻す(VBA は A 列・LOT で最初に一致した行を
    戻すので、同じ LOT を 2 回保存した耳用や、同じ分に開始・終了した行では別の行が戻った)。

    - `temp_date` があれば A 列(検査開始終了)で、無ければ B 列(LOT)で最初に一致した行(耳用は常に LOT)
    - `current`: いまの UF1 の値(VBA は空のセルのとき前の値を残すので、それを引き継ぐ)
    - `kanri_positions`: UF管理 が表示中なら、その位置名(S〜F)
    見つからなければ found=False(VBA は何もせずに抜ける)。WorkerTime_<L>!A1 に「履歴呼び出し」を書く。
    """
    out = FormRestore()
    cur = dict(current or {})
    out.values.update({k: v for k, v in cur.items()})
    if not store.exists(HOZON):
        out.found = False
        return out
    sh = store.sheet(MIMI if mimi else HOZON)
    lr, lc = sh.last_row(1), sh.last_col(1)
    rows = sh.read_range(2, 1, lr, lc)
    key_col = 1 if (temp_date != "" and not mimi) else 2
    key = temp_date if (temp_date != "" and not mimi) else hist_lot
    if row_no is not None and 2 <= row_no <= lr and variant_eq(rows[row_no - 2][1], hist_lot, vbx.STR):
        hit = rows[row_no - 2]                       # 選んだ行(LOT が合うときだけ)
    else:
        hit = next((r for r in rows if variant_eq(r[key_col - 1], key, vbx.STR)), None)
    if hit is None:
        out.found = False
        out.values = UiValues()
        return out
    row = hit + [None] * (115 - len(hit))

    def c(i: int) -> Value:
        return row[i - 1]

    v = out.values
    for nm, i in (("Lot", 2), ("ATU", 3), ("Hab", 4), ("Tak", 5), ("Am", 6), ("Ap", 7), ("Hm", 8), ("Hp", 9),
                  ("Tm", 10), ("Tp", 11)):
        out.set(nm, c(i))
    for nm in ("Zai", "Tyu", "Youto", "TAIK", "Machin", "Lateral", "Sumisei", "flatnesText", "flatnes",
               "ConductivText", "Comment", "Reien", "KenNo", "KenH", "Additional"):
        v[nm] = ""
    out.tags["ReSize"] = ""
    out.back_colors.update({"TaikText": 0x80000006, "LateralText": 0x80000006, "Conductiv": 0x8000000F,
                            "IzyouFl": 0x8000000F})
    atu, ap, am = vbx.val(v["ATU"]), vbx.val(v["Ap"]), vbx.val(v["Am"])
    v["ApK"] = vbx.fmt(atu + ap, "0.000")
    v["AmK"] = vbx.fmt(atu - am, "0.000")
    hab, tak = vbx.val(v["Hab"]), vbx.val(v["Tak"])
    v["HpK"] = vbx.fmt(hab + vbx.val(v["Hp"]), "0.0")
    v["HmK"] = vbx.fmt(hab - vbx.val(v["Hm"]), "0.0")
    v["TpK"] = vbx.fmt(tak + vbx.val(v["Tp"]), "0.0")
    v["TmK"] = vbx.fmt(tak - vbx.val(v["Tm"]), "0.0")

    vis_cols = range(99, 104) if mimi else range(85, 90)
    if any(c(i) is not None for i in vis_cols):              # IsEmpty でないものがある
        av: list[Any] = [None] * 6
        bv: list[Any] = [None] * 6
        for k, i in enumerate(vis_cols):
            spl = _vsplit(c(i), ",")
            if spl:
                av[k], bv[k] = spl[0], spl[-1]
            if not mimi and c(113) is not None:
                six = _vsplit(c(113), ",")
                if six:
                    av[5], bv[5] = six[0], six[-1]
        faces = FACES if not mimi else "SUCDE"
        for k, f in enumerate(faces):
            v[visual_up_name(f)] = cell_text(av[k]).replace("A_", "")
            v[visual_under_name(f)] = cell_text(bv[k]).replace("B_", "")

    if mimi:
        _history_load_mimi(out, c)
    else:
        _history_load_normal(out, c, cur, aorb)
        if kanri_positions:
            for f in FACES:
                v[f] = kanri_positions.get(f, "")
                v["F" + f] = kanri_positions.get(f, "")
        v["VC_L_A.Tag"] = ""
        v["VC_L_B.Tag"] = ""
        out.calls.append(("InitialSetting",))
    out.messages.append(Message("Load完了,印刷する場合はフォームを閉じないでください。", level=flow.INFO,
                                source="SaveLoad:3839"))
    for name, cap in load_allowed(store, sh.name, c(1), c(2)).items():
        v[name] = cap                                                # 2026-11-08: 許容・保留も戻す
    out.calls.append(("ReDiag",))
    v["dataLoding"] = "Load"
    wt = line_sheet(WORKER_TIME, lname)
    if not store.exists(wt):
        store.add_sheet(wt)
    store.sheet(wt).write(1, 1, "履歴呼び出し")
    return out


#: 位置ごとの「許容」「保留」(NGS〜NGF)の控え(2026-11-08)。保存用・耳用の行には許容の欄が無く、VBA も
#: 履歴から戻すと許容が消えて Min/Max が先へ進めなかった。保存用の列は変えず、別のシートに控えて戻す
ALLOWED = "許容控え"
ALLOWED_HEADERS = ("保存シート", "検査開始終了", "LOTNO") + tuple(f"NG{f}" for f in FACES)


def save_allowed(store: LocalStore, sheet: str, period: str, lot: str, ng: Mapping[str, str]) -> None:
    """許容・保留のある位置だけ控える(NG だけのものは控えない。NG は値から判定し直す)。"""
    caps = [cell_text(ng.get(f, "")) for f in FACES]
    if not any(c in ("許容", "保留") for c in caps):
        return
    sh = store.ensure_sheet(ALLOWED)
    if sh.last_row(1) < 1:
        sh.write_row(1, 1, list(ALLOWED_HEADERS))
    sh.write_row(sh.last_row(1) + 1, 1, [sheet, period, lot] + [c if c in ("許容", "保留") else "" for c in caps])


def load_allowed(store: LocalStore, sheet: str, period: Value, lot: Value) -> dict[str, str]:
    """控えた許容・保留(同じ保存シート・検査開始終了・LOT の最後の控え)。無ければ空。"""
    if not store.exists(ALLOWED):
        return {}
    sh = store.sheet(ALLOWED)
    found: dict[str, str] = {}
    for r in range(2, sh.last_row(1) + 1):
        if (cell_text(sh.get(r, 1)) == sheet and cell_text(sh.get(r, 2)) == cell_text(period)
                and cell_text(sh.get(r, 3)) == cell_text(lot)):
            found = {f"NG{f}": cell_text(sh.get(r, 4 + k)) for k, f in enumerate(FACES)
                     if cell_text(sh.get(r, 4 + k)) != ""}
    return found


def _history_load_normal(out: FormRestore, c, cur: Mapping[str, str], aorb: str) -> None:
    v = out.values
    if variant_eq(c(84), "〆", vbx.STR):
        for f in FACES:
            v[f"LG{f}"] = "ﾌﾘｰｿﾘ"
    tai = []
    for a_col in (20, 32, 44, 56, 68, 108):
        tai.append(vbx.single(abs(vbx.val(cell_text(c(a_col))) - vbx.val(cell_text(c(a_col + 1))))))
    code = cell_text(c(83))
    ken_l, ken_r = vbx.left(code, 1), vbx.right(code, 1)
    if ken_r == "4":
        for i in range(1, 11):
            out.set(f"F{i}", c(i + 99))
        if cell_text(c(111)) != "":
            v["LateralF"] = vbx.fmt(vbx.val(cell_text(c(112))), "0.0")
        if cell_text(c(113)) != "":
            v["DiagFKou"] = vbx.fmt(vbx.val(cell_text(c(114))), "0.0")
    for i in range(1, 11):
        pat = row_format(i)
        for k, f in enumerate("SUCDE"):
            v[f"{f}{i}"] = vbx.fmt(vbx.val(cell_text(c(i + 11 + 12 * k))), pat)
        v[f"F{i}"] = vbx.fmt(vbx.val(v.get(f"F{i}", "")), pat)
    for k, f in enumerate(FACES):
        if v.get(f"{f}10", "") != "":
            # Tai は Single。Format は Single の 7 桁の値で丸める(U5 と同じく実機確認の対象)
            v[diag_name(f)] = vbx.fmt(vbx.cstr(tai[k], is_single=True), "0.0")
    for k, f in enumerate("SUCDE"):
        if cell_text(c(95 + k)) != "":
            v[kou_name(f)] = vbx.fmt(vbx.val(cell_text(c(95 + k))), "0.0")
        if cell_text(c(72 + k)) != "":
            v[lateral_name(f)] = vbx.fmt(vbx.val(cell_text(c(72 + k))), "0.0")
        if cell_text(c(90 + k)) != "":
            v[bh_name(f)] = vbx.fmt(vbx.val(cell_text(c(90 + k))), "0")

    z = {"a": "", "b": ""}
    for k, f in enumerate("SUCDE"):
        cell = c(22 + 12 * k)
        if cell_text(cell) != "":
            _history_strain(out, f, "Z", cell, z, guard=True)
    _history_strain(out, "F", "Z", c(110), z, guard=False)
    g = {"a": "", "b": ""}
    gb_prev: list[str] = []
    for k, f in enumerate("SUCDE"):
        gb_prev = _history_strain(out, f, "G", c(23 + 12 * k), g, guard=False, gb_prev=gb_prev)
    _history_strain(out, "F", "G", c(111), g, guard=False, gb_prev=gb_prev)
    out.calls.append(("ReflectStrain", aorb[-1:], "保存"))
    for i in range(1, 11):
        for f in FACES:
            if _is_zero_caption(v[f"{f}{i}"]):
                v[f"{f}{i}"] = ""
                out.back_colors[f"{f}{i}"] = RGB_WHITE_DEFAULT
    if ken_l == "H" and ken_r == "数":
        out.calls.append(("LSelect", "少枚数"))
    if ken_l == "H" and ken_r == "H":
        out.calls.append(("LSelect", "厚板"))
    if ken_l != "H" and ken_r in ("1", "2", "3"):
        out.calls.append(("LSelect", "薄板"))
    if ken_l != "H" and ken_r == "4":
        out.calls.append(("Six_Display", "on"))
    if any(v.get(lateral_name(f), "") != "" for f in "SUCDE"):
        out.calls.append(("LSelectLateral", "強制"))
    if any(v.get(kou_name(f), "") != "" for f in "SUCDE"):
        out.calls.append(("RightAngleDisplay", "強制"))
    if v.get("LateralF", "") != "":
        out.calls.append(("LSelectLateral", "強制", "LateralF"))
    if v.get("DiagFKou", "") != "":
        out.calls.append(("RightAngleDisplay", "強制", "DiagFKou"))


def _history_strain(out: FormRestore, face: str, kind: str, cell: Value, state: dict[str, str], *,
                    guard: bool, gb_prev: Optional[list[str]] = None) -> list[str]:
    """HistoryLoad の歪 1 面 (SaveLoad:3616-3773)。定盤とゲージで微妙に違う分岐をそのまま再現する。

    B76: VBA は TempA/tempB(ゲージは tempC/TempD)を面ごとに空にしないので前の面の値が残り、ゲージの
    「A のみ」が空の面では前の面の GdataB が Tag に入る。直すと面ごとに空から始め、前の面の B は使わない。
    """
    src = "SaveLoad:3631"
    text = cell_text(cell)
    fix = True
    if fix:
        state["a"] = state["b"] = ""
        gb_prev = None
    gb: list[str] = list(gb_prev or [])
    if "," in text:
        a, b = _ab_strain(text, state)
        za, zb = _vsplit(a, "_"), _vsplit(b, "_")
        gb = zb
        if len(za) - 1 == 2:
            _split(out, face, kind, za, zb, src)
        elif not za:
            if kind == "Z":
                _split_0(out, face, kind, zb, src)
            else:
                _split_0(out, face, kind, None, src)            # ゲージの AB 分岐は Split_G_0(n) だけ
    else:
        za = _vsplit(text.replace("A:", ""), "_")
        if len(za) - 1 == 2:
            _split(out, face, kind, za, None, src)
        elif not za:
            if kind == "Z":
                _split_0(out, face, kind, None, src)
            else:
                _split_0(out, face, kind, gb, src)               # 前の面の GdataB が残っている
    return gb


def _history_load_mimi(out: FormRestore, c) -> None:
    v = out.values
    for i in range(1, 15):
        pat = row_format(i)
        for k, f in enumerate("SUCDE"):
            v[f"{f}{i}"] = vbx.fmt(vbx.val(cell_text(c(i + 11 + 16 * k))), pat)
    for k, f in enumerate("SUCDE"):
        za = _vsplit(c(26 + 16 * k), "_")
        if len(za) - 1 == 2:
            _split(out, f, "Z", za, None, "SaveLoad:3947")
        ga = _vsplit(c(27 + 16 * k), "_")
        if len(ga) - 1 == 2:
            _split(out, f, "G", ga, None, "SaveLoad:3965")
        if cell_text(c(104 + k)) != "":
            v[bh_name(f)] = vbx.fmt(vbx.val(cell_text(c(104 + k))), "0")
    for i in range(1, 15):
        for f in "SUCDE":
            if _is_zero_caption(v[f"{f}{i}"]):
                v[f"{f}{i}"] = ""
                out.back_colors[f"{f}{i}"] = RGB_WHITE_DEFAULT


def kanri_lots(store: LocalStore, mimi: bool = False) -> list[str]:
    """UF管理 の LOT 一覧(UserForm_Initialize / MIMI_AfterUpdate: B 列の 2 行目〜最後の行)。"""
    name = MIMI if mimi else HOZON
    if not store.exists(name):
        return []
    sh = store.sheet(name)
    k = sh.last_row(2)
    return [cell_text(sh.get(a, 2)) for a in range(2, k + 1)]


def kanri_select_time(store: LocalStore, target_text: str, cb1_text: str, mimi: bool = False) -> Optional[int]:
    """UF管理.NewCB_Change (UF管理:11-35)。同じ LOT が複数あるとき、時間で選んだ行の ListIndex(i-2)。"""
    sh = store.sheet(MIMI if mimi else HOZON)
    end_r = sh.last_row(1)
    for i, row in enumerate(sh.read_range(1, 1, end_r, 2), start=1):
        if cell_text(row[0]) == target_text and cell_text(row[1]) == cb1_text:
            return i - 2
    return None


def _kanri_split(v: dict[str, str], z: str, zb: str, g: str, gb: str, source: str) -> None:
    """UF管理 の歪 A/B の分け方 (UF管理:236-253)。"""
    if "," not in v[z]:
        return
    arr = v[z].split(",")
    if "A:" in arr[0]:
        v[z] = arr[0].replace("A:", "")
    elif "B:" in arr[-1]:
        v[z] = ""
    if "B:" in arr[-1]:
        v[zb] = arr[-1].replace("B:", "")
    arr_g = _vsplit(v[g], ",")
    if not arr_g:
        raise _err9(source)
    if "A:" in arr_g[0]:
        v[g] = arr_g[0].replace("A:", "")
    elif "B:" in arr_g[-1]:
        v[g] = ""
    if "B:" in arr_g[-1]:
        v[gb] = arr_g[-1].replace("B:", "")


@dataclass
class KanriView:
    """UF管理 の表示(CB1_Change)。values: ラベル名 → 文字、new_cb: 同じ LOT の時間の候補。"""
    values: dict[str, str] = field(default_factory=dict)
    new_cb: Optional[list[str]] = None
    new_cb_alert: bool = False
    messages: list[Message] = field(default_factory=list)


def kanri_view(store: LocalStore, list_index: int, cb1_list: Sequence[str], cb1_text: str, *, mimi: bool = False,
               cb1_tag: str = "") -> KanriView:
    """UF管理.CB1_Change (UF管理:37-618)。選んだ行の内容を表示用に並べる。

    `list_index` は CB1.ListIndex(行 = list_index + 2)。B27: F 面ゲージ歪の表示条件が F 定盤歪の列(110)を見る。
    """
    out = KanriView()
    v = out.values
    count = sum(1 for x in cb1_list if x == cb1_text)
    sh = store.sheet(MIMI if mimi else HOZON)
    k = sh.last_row(2)
    if count > 1 and cb1_tag != "ループ":
        out.new_cb = [cell_text(sh.get(a, 1)) for a in range(2, k + 1) if cell_text(sh.get(a, 2)) == cb1_text]
    elif count == 1:
        out.new_cb = []
    for i in range(1, 15):
        for f in FACES:
            v[f"{f}{i}"] = ""
    for f, lf in zip(FACES, ("Slate", "Ulate", "Clate", "Dlate", "Elate", "Flate")):
        v[lf] = ""
        v[f"A{f.lower()}_Vis"] = ""
        v[f"B{f.lower()}_Vis"] = ""
        v[f"{f}_BH"] = ""
        v[f"{f}ZH"] = v[f"{f}GH"] = v[f"{f}ZH_B"] = v[f"{f}GH_B"] = ""
        v[f"{f}_Angle"] = ""
    if count > 1 and cb1_tag != "ループ":
        out.messages.append(Message("同LotNoあり、作業時間を選択してください", level=flow.WARN, source="UF管理:156"))
        out.new_cb_alert = True
    if not store.exists(MIMI if mimi else HOZON):
        return out
    r = list_index + 2
    t = lambda col: cell_text(sh.get(r, col))   # noqa: E731
    v["Count"] = "作業数" + str(len(cb1_list)) + "件"
    for nm, col in (("ATU", 3), ("HABA", 4), ("TAKE", 5), ("Am", 6), ("Ap", 7), ("Hm", 8), ("Hp", 9), ("Tm", 10),
                    ("Tp", 11)):
        v[nm] = t(col)
    v["SE"] = t(1)
    av: list[Any] = [None] * 6
    bv: list[Any] = [None] * 6
    if not mimi:
        v.update({"H1": "A_定盤歪", "H2": "A_ｹﾞｰｼﾞ歪", "LhizB": "B_定盤歪", "LhigB": "B_ｹﾞｰｼﾞ歪", "A1": "厚1",
                  "A2": "厚2", "A3": "厚3", "A4": "厚4", "L1": "巾1", "L2": "巾2", "L3": "丈1", "L4": "丈2",
                  "L5": "対角A", "L6": "対角B", "L7": "", "L8": "", "Llate": "ﾗﾃﾗﾙﾎﾞｰ"})
        if t(84) == "〆":
            v["H2"] = "フリー反り"
            v["LhigB"] = ""
        v["Worker"] = t(77)
        v["WorkJudgment"] = t(83)
        for f, col in zip(FACES, (22, 34, 46, 58, 70, 110)):
            if t(col) != "__":
                v[f"{f}ZH"] = t(col)
        for f, col in zip("SUCDE", (23, 35, 47, 59, 71)):
            if t(col) != "__":
                v[f"{f}GH"] = t(col)
        cond = 111
        if t(cond) != "__":
            v["FGH"] = t(111)
        for f in "SUCDE":
            _kanri_split(v, f"{f}ZH", f"{f}ZH_B", f"{f}GH", f"{f}GH_B", "UF管理:246")
        for i in range(1, 11):
            for k2, f in enumerate("SUCDE"):
                v[f"{f}{i}"] = t(i + 11 + 12 * k2)
        for k2, f in enumerate("SUCDE"):
            v[f"{f}late"] = t(72 + k2)
        vis = [sh.get(r, col) for col in range(85, 90)]
        if any(x is not None for x in vis):
            for k2, x in enumerate(vis):
                spl = _vsplit(x, ",")
                if spl:
                    av[k2], bv[k2] = spl[0], spl[-1]
            six = sh.get(r, 113)
            if six is not None:
                spl = _vsplit(six, ",")
                if spl:
                    av[5], bv[5] = spl[0], spl[-1]
            for k2, f in enumerate("SUCDE"):
                v[f"A{f.lower()}_Vis"] = cell_text(av[k2]).replace("A_", "")
                v[f"B{f.lower()}_Vis"] = cell_text(bv[k2]).replace("B_", "")
                v[f"{f}_BH"] = t(90 + k2)
                v[f"{f}_Angle"] = t(95 + k2)
        code = t(83)
        ken_l, ken_r = vbx.left(code, 1), vbx.right(code, 1)
        pos = None
        if ken_l == "H" and ken_r == "数":
            pos = (["1", "2", "3", "4", "5", ""], ["1", "2", "3", "4", "5", ""])
        elif ken_l == "H" and ken_r == "H":
            pos = (["S", "S1", "2", "E1", "E", ""], ["S", "S1", "2", "E1", "E", ""])
        elif ken_l != "H" and ken_r == "1":
            pos = (["S", "", "2", "", "E", ""], ["S", "", "2", "", "E", ""])
        elif ken_l != "H" and ken_r == "2":
            pos = (["2-S", "", "2-E", "", "1-E", ""], ["S", "", "2-E", "", "1-E", ""])
        elif ken_l != "H" and ken_r == "3":
            pos = (["3-S", "3-E", "2-2", "1-S", "1-E", ""], ["3-S", "3-E", "2-2", "1-S", "1-E", ""])
        elif ken_l != "H" and ken_r == "4":
            pos = (["4-S", "4-E", "3-2", "2-2", "1-S", "1-E"], ["4-S", "4-E", "3-2", "2-2", "1-S", "1-E"])
        if pos:
            for f, p1, p2 in zip(FACES, pos[0], pos[1]):
                v[f] = p1
                v[f + "L"] = p2
        if ken_r == "4":
            for i in range(1, 11):
                v[f"F{i}"] = t(i + 99)
            v["Flate"] = t(112)
            v["F_Angle"] = t(114)
            v["Af_Vis"] = cell_text(av[5]).replace("A_", "")
            v["Bf_Vis"] = cell_text(bv[5]).replace("B_", "")
            _kanri_split(v, "FZH", "FZH_B", "FGH", "FGH_B", "UF管理:499")
    else:
        v.update({"H1": "定盤歪", "H2": "ｹﾞｰｼﾞ歪", "A1": "厚1", "A2": "厚2", "A3": "厚3", "A4": "厚4",
                  "L1": "巾1", "L2": "巾2", "L3": "巾3", "L4": "巾4", "L5": "巾5", "L6": "巾6", "L7": "丈1",
                  "L8": "丈2", "L9": "丈3", "L10": "丈4"})
        v["Worker"] = t(92)
        v["WorkJudgment"] = t(98)
        for k2, f in enumerate("SUCDE"):
            v[f"{f}ZH"] = t(26 + 16 * k2)
            v[f"{f}GH"] = t(27 + 16 * k2)
        for i in range(1, 15):
            for k2, f in enumerate("SUCDE"):
                v[f"{f}{i}"] = t(i + 11 + 16 * k2)
        vis = [sh.get(r, col) for col in range(99, 104)]
        if any(x is not None for x in vis):
            for k2, x in enumerate(vis):
                spl = _vsplit(x, ",")
                if spl:
                    av[k2], bv[k2] = spl[0], spl[-1]
            for k2, f in enumerate("SUCDE"):
                v[f"A{f.lower()}_Vis"] = cell_text(av[k2]).replace("A_", "")
                v[f"B{f.lower()}_Vis"] = cell_text(bv[k2]).replace("B_", "")
                v[f"{f}_BH"] = t(104 + k2)
        # 2026-11-08 に直した(B92): 保存値は「耳付き」なので、それでも位置名を 1〜5 に切り替える(「極厚」も残す)。
        # VBA(UF管理.frm:600)は「極厚」か空だけを見るため、耳付きの履歴で位置名が変わらなかった
        if t(98) in ("極厚", "耳付き", ""):
            for f, p1, p2 in zip("SUCDE", ("1", "2", "3", "4", "5"), ("S", "U", "2", "D", "E")):
                v[f] = p1
                v[f + "L"] = p2
    return out


def kanri_load(store: LocalStore, cb1_text: str, new_cb_text: str, *, uf1_visible: bool, uf1_lot: str,
               mimi: bool = False, current: Optional[Mapping[str, str]] = None, lname: str = "",
               kanri_positions: Optional[Mapping[str, str]] = None, aorb: str = Face.A,
               row_no: Optional[int] = None) -> FormRestore:
    """UF管理.CommandButton1_Click (UF管理:620-635)。別の LOT を表示中なら止める。"""
    if uf1_visible and uf1_lot != "":
        out = FormRestore(found=False)
        out.messages.append(Message("別Lotデータがあります、新規登録処理を行ってください", level=flow.ERROR,
                                    source="UF管理:622"))
        return out
    res = history_load(store, cb1_text, new_cb_text, mimi=mimi, current=current, lname=lname,
                       kanri_positions=kanri_positions, aorb=aorb, row_no=row_no)
    res.calls.append(("InitialSetting", "サイズ保存"))
    return res


# ------------------------------------------------------------------
# 梱包枚数チェック・終了時刻
# ------------------------------------------------------------------
@dataclass
class NumberCountResult:
    stop: bool = False                         # UF1.PrintB.Tag = "stop"
    red_row: Optional[int] = None              # Daily_<L> の E 列を赤くした行
    open_namelist: Optional[int] = None        # NameList_Open(3)
    packed: dict = field(default_factory=dict)       # LOT → 梱包枚数
    inspected: dict = field(default_factory=dict)    # LOT → 検入枚数
    messages: list[Message] = field(default_factory=list)


def number_count(store: LocalStore, lname: str, lotname: str) -> Optional[NumberCountResult]:
    """Worker.Number_Count (Worker:948-1023)。日報の梱包枚数(枚数×包数)が検入枚数を超えたら止める。

    Daily_<L> が無ければ None。梱包枚数計算_<L> を作り直して結果を書く。
    """
    daily_name = line_sheet(DAILY, lname)
    if not store.exists(daily_name):
        return None
    sh: GridSheet = store.sheet(daily_name)  # type: ignore[assignment]
    ws = store.add_sheet(line_sheet(PACK_COUNT, lname))
    ur5 = sh.last_row(5)
    for r in range(min(3, ur5), max(3, ur5) + 1):
        if sh.color(r, 5) is not None:
            sh.set_color(r, 5, None)
    ur1 = sh.last_row(1)
    if ur1 < 3:
        total = []                       # 直した側: 3 行目以降が空なら数える行なし(見出しの 2 行目を数えない)
    else:
        total = [list(row) for row in sh.read_range(3, 1, ur1, 13)]   # B78: A3:M2 は A2:M3 になり見出しを数える
    count_dic: dict[Any, Value] = {}
    kara: str = ""
    for row in total:
        if cell_text(row[0]) == "":
            row[0] = kara
        else:
            kara = cell_text(row[0])
            if row[0] not in count_dic:
                count_dic[row[0]] = row[4]
    lot_dic: dict[Any, float] = {}
    for row in total:
        x = vbx.clng(vbx.val(cell_text(row[11])) * vbx.val(cell_text(row[12])))
        if row[0] not in lot_dic:
            lot_dic[row[0]] = x
        else:
            lot_dic[row[0]] = vbx.val(cell_text(lot_dic[row[0]])) + x
    ws.write(1, 1, "LotNo")
    ws.write(1, 2, "梱包枚数")
    for j, (k, val) in enumerate(lot_dic.items(), start=2):
        ws.set(j, 1, excel_value(k))
        ws.set(j, 2, val)
    ws.write(1, 3, "LotNo")
    ws.write(1, 4, "検入枚数")
    for j, (k, val) in enumerate(count_dic.items(), start=2):
        ws.set(j, 3, excel_value(k))
        ws.set(j, 4, val)
    out = NumberCountResult(packed={cell_text(k): v for k, v in lot_dic.items()},
                            inspected={cell_text(k): v for k, v in count_dic.items()})
    if lotname in count_dic:
        if vbx.val(cell_text(lot_dic.get(lotname))) > vbx.val(cell_text(count_dic[lotname])):
            posi = 0
            for i, row in enumerate(total, start=1):
                if cell_text(row[0]) == lotname and cell_text(row[4]) != "":
                    posi = i
                    break
            out.red_row = posi + 2          # 見つからなければ 0 + 2(VBA どおり)
            sh.set_color(out.red_row, 5, RGB_RED_NG)
            out.stop = True
            out.open_namelist = 3
            out.messages.append(Message("梱包数と検入枚数が一致しません", level=flow.ERROR, source="Worker:1018"))
    return out


def end_time_register(store: LocalStore, lname: str, now: datetime) -> str:
    """Worker.EndTime_Register (Worker:312-325)。終了時刻(UF1.EndTime の Caption)を決めて WorkerTime!C1 にも書く。

    7:00 より後〜23:59 は今。0:00〜7:00(7:00 ちょうどを含む)は日付だけ 1 日戻す。それ以外(23:59:01〜)は今。
    文字の形は VBA どおり: 今のときは CStr(Date)("2026/10/30 9:05:03")、前日のときは
    Format("yyyy/mm/dd hh:mm:ss")("2026/10/29 03:15:22")。B38(不良要因票との違い)は印刷側で扱う。
    """
    now_s = now.replace(microsecond=0)
    t = now_s.time()
    if dtime(7, 0) < t <= dtime(23, 59):
        end = vb_cstr_date(now_s)
    elif dtime(0, 0) <= t <= dtime(7, 0):
        end = vb_format_date(now_s - timedelta(days=1), "yyyy/mm/dd hh:mm:ss")
    else:
        end = vb_cstr_date(now_s)
    wt = line_sheet(WORKER_TIME, lname)
    if store.exists(wt):
        store.sheet(wt).write(1, 3, end, "O")
    return end


# ------------------------------------------------------------------
# CSV(ExpCSV)
# ------------------------------------------------------------------
def _csv_quote(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def exp_csv(sheet, file_name: str = "", file_path: Union[str, Path] = "", sep: str = ",", maxi: int = 0,
            maxj: int = 0, *, default_dir: Optional[Union[str, Path]] = None, quote: bool = False) -> Path:
    """Input.ExpCSV (Input:838-874)。シートを区切り文字の文字列に(各行の末尾にも区切り。値は囲まない)。

    文字コードは ANSI(cp932。表せない文字は "?")、改行は CRLF。拡張子は "," なら .csv、それ以外は .txt。
    `file_path` が空なら `default_dir`(VBA の ThisWorkbook.Path に当たる)。"hostsend" は 1 セルを 1 行で書く。
    """
    lr, lc = sheet.last_cell()
    maxi = maxi or lr
    maxj = maxj or lc
    data = sheet.read_range(1, 1, maxi, maxj)
    folder = str(file_path) if str(file_path) else str(default_dir or paths.write_dir())
    name = file_name or "ExportFile"
    ext = ".csv" if sep == "," else ".txt"
    if "\\" in folder:
        target = paths.to_local(paths.join(folder, name + ext))
        if os.sep == "/" and "\\" in str(target):
            target = Path(str(target).replace("\\", "/"))
    else:
        target = Path(folder) / (name + ext)
    target.parent.mkdir(parents=True, exist_ok=True)
    if name == "hostsend":
        if len(data) != 1 or len(data[0]) != 1:
            raise _err13("Input:860")                 # 配列を String に入れる → 型エラー
        lines_out = [cell_text(data[0][0])]
    else:
        if quote and sep == ",":
            # 2026-11-08 のご判断: 履歴 CSV は値を "…" で囲む(中の " は "" にする)。VBA は囲まなかったので、外観などに
            # カンマがあると列がずれていた。各行の末尾の区切りは VBA のまま残す
            lines_out = ["".join(_csv_quote(cell_text(v)) + sep for v in row) for row in data]
        else:
            lines_out = ["".join(cell_text(v) + sep for v in row) for row in data]
    with open(target, "w", encoding="cp932", errors="replace", newline="") as fh:
        for line in lines_out:
            fh.write(line + "\r\n")
    return target


_HISTORY_DIRS = {"L-1": "L-1\\L-1_履歴", "LVC": "LVC\\LVC_履歴", "HVC": "HVC\\HVC_履歴",
                 "バランサー": "バランサー\\バランサー_履歴", "トット": "トット\\トット_履歴", "中板": "丸徳\\丸徳_履歴"}
_FACTORY_DIRS = {"L-1": "【梱包】L-1", "LVC": "【梱包】LVC", "HVC": "【梱包】HVC", "バランサー": "【梱包】大板",
                 "トット": "【梱包】大板", "中板": "【梱包】中板"}


def export_history_csv(store: LocalStore, lname: str, line: str, extreme: bool, *,
                       local_dir: Optional[Union[str, Path]] = None) -> list[Path]:
    """印刷後の CSV(UF1:3780-3815)。保存用(極厚は耳用)を 3 か所へ。

    1. 手元(ThisWorkbook.Path に当たる `local_dir`)に `<L>_履歴CSV.csv`(耳付きは `<L>_耳用履歴CSV.csv`)
    2. HistoryPath\\<ライン>\\<ライン>_履歴 に `<L>_File.csv`(耳付きは `<L>_耳用File.csv`)
    3. FactoryyPath\\【梱包】<L-1/LVC/HVC/大板/中板> に同じ名前(Test などは出さない)
    VBA は Excel で開かれていたら書かない(IsBookOpened)。ここでは書けなければ飛ばす。
    """
    sh = store.sheet(MIMI if extreme else HOZON)
    official = lines.official(line)
    out = [exp_csv(sh, lname + ("_耳用履歴CSV" if extreme else "_履歴CSV"), paths.get("csv_local"),
                   default_dir=local_dir or store.path.parent, quote=True)]
    new_name = lname + ("_耳用File" if extreme else "_File")
    targets = [paths.join(paths.get("history"), _HISTORY_DIRS.get(official, "Test\\Test_履歴"))]
    if official in _FACTORY_DIRS:
        targets.append(paths.join(paths.get("factory"), _FACTORY_DIRS[official]))
    for folder in targets:
        try:
            out.append(exp_csv(sh, new_name, folder, quote=True))
        except OSError:
            continue
    return out


def _local_dir(folder_text: str) -> Path:
    """設定のパス文字列(\\ 区切り)→ この PC で書ける Path(exp_csv と同じ変換)。"""
    p = paths.to_local(str(folder_text))
    if os.sep == "/" and "\\" in str(p):
        p = Path(str(p).replace("\\", "/"))
    return p


def _numbered(path: Path) -> Path:
    """VBA と同じく、あれば .002〜.999 を付けた名前にする(Input:1095-1105)。"""
    if not path.exists():
        return path
    for i in range(2, 1000):
        p = path.with_name(f"{path.stem}.{i:03d}{path.suffix}")
        if not p.exists():
            return p
    raise OSError(f"同じ名前のファイルが 999 個を超えました: {path}")


#: UF1 の履歴ブックに入れるシート(Input:1070-1087)。Lot情報履歴 は別に出す(2026-11-07 のご判断)
_BOOK_SHEETS = ("WorkerTime", "作業保存", "Daily")


def export_lot_book(store: LocalStore, lname: str, line: str, lot: str, now: datetime,
                    report: Any = None, defect: Any = None) -> list[Path]:
    """Input.履歴保存 の xlsx(Input:1070-1140)。1 つのブックに複数シート(値だけ)。

    手元: <手元の検査履歴>\\yyyy.mm\\yyyymmdd_LOT\\yyyymmdd_LOT.xlsx(あれば .002〜)
    共有: <検査履歴>\\<ライン>\\<ライン>_履歴\\yyyy.mm\\ へ同じファイルを写す
    """
    import shutil
    from .xlsx_writer import grid_rows, write_xlsx
    book: list[tuple[str, list]] = []
    if report is not None:
        book.append((f"印刷用_{lname}", grid_rows((c["row"], c["col"], c["value"]) for c in report.cells)))
    for base in _BOOK_SHEETS:
        name = f"{base}_{lname}"
        if store.exists(name):
            book.append((name, grid_rows(store.sheet(name).used())))
    if defect is not None:
        book.append((f"不良要因印刷_{lname}", grid_rows((c["row"], c["col"], c["value"]) for c in defect.cells)))
    month, day_lot = now.strftime("%Y.%m"), f"{now:%Y%m%d}_{lot}"
    local = _numbered(_local_dir(paths.join(paths.get("local_history"), month, day_lot)) / f"{day_lot}.xlsx")
    out = [write_xlsx(local, book)]
    share = _local_dir(paths.join(paths.get("history"),
                                  _HISTORY_DIRS.get(lines.official(line), "Test\\Test_履歴"), month))
    try:
        # 2026-11-08: 設定の「検査履歴」に届くときだけ、その下(ライン・月)のフォルダを作る
        safe_io.mkdir_under(_local_dir(paths.get("history")), share)
        out.append(Path(shutil.copy2(local, share / local.name)))
    except OSError:
        pass                                   # 共有に届かなくても手元は残す(VBA は FileCopy の失敗で止まる)
    return out


def export_lot_history(store: LocalStore, line: str, *, local_dir: Optional[Union[str, Path]] = None) -> list[Path]:
    """Lot情報履歴 を手元と共有(設定「Lot情報履歴の共有先」)に CSV で出す(2026-11-07 のご判断)。

    共有先が空なら <検査履歴>\\<ライン>\\<ライン>_履歴。ファイル名は Lot情報履歴_<ライン>.csv(上書き)。
    """
    if not store.exists(LOT_HISTORY):
        return []
    sh = store.sheet(LOT_HISTORY)
    official = lines.official(line) or line or "Test"
    name = f"Lot情報履歴_{official}"
    out = [exp_csv(sh, name, paths.get("csv_local"), default_dir=local_dir or store.path.parent, quote=True)]
    share = paths.get("lot_history_share") or paths.join(paths.get("history"),
                                                          _HISTORY_DIRS.get(official, "Test\\Test_履歴"))
    try:
        out.append(exp_csv(sh, name, share, quote=True))
    except OSError:
        pass
    return out
