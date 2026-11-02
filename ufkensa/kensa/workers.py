"""作業者: 社員カード → 名簿、直と日付の判定、作業時間、検査員の照合、日報

VBA `【UFKen_07】Worker`、`UFNameList`(書き戻し)、`UFWorkerName`、UF1.Worker_Change(カード登録の中身)の移植。

【シートの扱い】VBA はシート(検査員_<L>・WorkerTime_<L>・Daily_<L>・梱包枚数計算_<L>)に直接書いていた。
ここでは**今のセルの値を受け取り、書くセル(行・列・値)をデータで返す**。保存先の表は別担当。
セルは `{(行, 列): 値}`(1 始まり)。Excel の空セルは None(VBA の Empty)、"" も空として扱う。

【直の判定】設備名(レジストリ「設備名」)が L-1 / LVC / HVC / Test のときだけ 3 交替。ほかは日勤で当日。
直の時刻は DB「時間用」(伝送用ファイル)。`load_shift_times()` で読む(VBA の ShiftTime + Set_Shift)。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field, replace
from datetime import datetime, timedelta
from typing import Any, Callable, Mapping, Optional

from . import datasource, legacy_log, lines, paths, quirks, vbx
from .flow import ERROR, INFO, Message

Cells = Mapping[tuple[int, int], Any]
SHIFT_FACILITIES = ("L-1", "LVC", "HVC", "Test")       # Worker:150, 217, 258
DEFAULT_FACILITY = "存在なし"                           # GetSetting の既定値
UNREGISTERED_SOUND = "ナンバーNO.wav"                    # 音楽未登録(補助:343)
DAILY_FIELDS = ("ﾛｯﾄNo", "材・調質", "製品寸法", "梱包No", "検入枚数", "開始時", "開始分", "終了時", "終了分",
                "作業人数", "合紙", "枚数", "包数", "VC種別", "etc", " 停止①記号", " 停止①時間", " 停止②記号",
                " 停止②時間", " 停止③記号", " 停止③時間", " 実績枚数", " 実績重量", " 作業時間", " 単重")   # Worker:736-737


# ------------------------------------------------------------------
# セルの道具
# ------------------------------------------------------------------
@dataclass(frozen=True)
class CellWrite:
    row: int
    col: int
    value: Any


def _empty(v: Any) -> bool:
    return v is None or v == ""


def _text(v: Any) -> str:
    """セルの値の CStr(Empty は "")。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return vbx.cstr(v)


def _cell(v: Any) -> Any:
    """Excel のセル値: "" は Empty(None)として扱う。"""
    return None if v == "" else v


def _veq(a: Any, b: Any) -> bool:
    return vbx.vcompare(a, b) == 0


def last_row(cells: Cells, col: int) -> int:
    """`sh.Cells(Rows.Count, col).End(xlUp).Row`(空の列は 1)。"""
    rows = [r for (r, c), v in cells.items() if c == col and not _empty(v)]
    return max(rows) if rows else 1


def _num(v: float) -> Any:
    """Val の結果をセルに書く値(整数なら int)。"""
    return int(v) if float(v).is_integer() else v


# ------------------------------------------------------------------
# 直の時刻(DB 時間用)
# ------------------------------------------------------------------
@dataclass(frozen=True)
class ShiftTimes:
    """Worker の Start1-3 / End1-3 / Start昼 / End昼(文字列のまま。VBA の Public 変数)。"""
    start1: str = ""
    end1: str = ""
    start2: str = ""
    end2: str = ""
    start3: str = ""
    end3: str = ""
    start_day: str = ""
    end_day: str = ""


def shift_from_records(fields: Optional[list[str]], records: Optional[list[list[str]]],
                       base: ShiftTimes = ShiftTimes()) -> ShiftTimes:
    """QP_DB.ShiftTime (QP_DB:1786-1828) + Worker.Set_Shift (Worker:12-35)。

    - 先頭 4 行だけを使う(4 行未満は VBA の実行時エラー 9 → IndexError)
    - 表が空なら何もしない(Set_Shift の IsEmptyArray で抜ける → 前の値のまま)
    """
    if fields is None or records is None:
        return base
    c_start = datasource.column_index(fields, "開始")
    c_end = datasource.column_index(fields, "終了")
    c_shift = datasource.column_index(fields, "直")
    if min(c_start, c_end, c_shift) < 0:
        raise IndexError("時間用 の列(直・開始・終了)が見つかりません(VBA の実行時エラー 9)")
    if len(records) < 4:
        raise IndexError("時間用 が 4 行未満(VBA の実行時エラー 9)")
    out = base
    for r in records[:4]:
        key, s, e = r[c_shift], r[c_start], r[c_end]
        if key == "1":
            out = replace(out, start1=s, end1=e)
        elif key == "2":
            out = replace(out, start2=s, end2=e)
        elif key == "3":
            out = replace(out, start3=s, end3=e)
        elif key == "昼":
            out = replace(out, start_day=s, end_day=e)
    return out


def load_shift_times(base: ShiftTimes = ShiftTimes()) -> ShiftTimes:
    """DB「時間用」から直の時刻を読む(ShiftTime: 品証やり取り → 無ければ参照用ファイル)。

    品証やり取りの DB ファイル自体が無いときは VBA は End(D10)→ `datasource.DbMissing` が上がる。
    """
    db = paths.db_file("伝送用ファイル")
    fields = datasource.get_fields(db, "時間用")
    if fields is None:
        db = paths.db_file("伝送用ファイル", base_key="packing")
        fields = datasource.get_fields(db, "時間用")
        if fields is None:
            return base
    return shift_from_records(fields, datasource.get_records(db, "時間用"), base)


_TIME_RE = re.compile(r"^\s*(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s*$")


def _seconds(text: Any) -> int:
    """`CDate("7:00")` の時刻を 0 時からの秒に。読めなければ型エラー 13(CDate("") も)。"""
    if isinstance(text, datetime):
        return text.hour * 3600 + text.minute * 60 + text.second
    m = _TIME_RE.match("" if text is None else str(text))
    if not m:
        raise vbx.TypeMismatch(f"CDate({text!r})")
    h, mi, s = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
    if h > 23 or mi > 59 or s > 59:
        raise vbx.TypeMismatch(f"CDate({text!r})")
    return h * 3600 + mi * 60 + s


_2359 = 23 * 3600 + 59 * 60


def _shift_branch(dtime: str, shifts: ShiftTimes) -> int:
    """YeaDayrCheck / DayCheck / TimeCheck 共通の If…ElseIf(Worker:153-163)。

    1=1直 2=2直 3=3直(当日) 4=3直(0時〜End3。前日扱い) 5=どれでもない(3直・当日)
    VBA の And は両辺を評価するので、各分岐の 2 つの CDate は両方読む。
    """
    t = _seconds(dtime)
    a, b = _seconds(shifts.start1), _seconds(shifts.end1)
    if t > a and t < b:
        return 1
    a, b = _seconds(shifts.start2), _seconds(shifts.end2)
    if t >= a and t < b:
        return 2
    a = _seconds(shifts.start3)
    if t >= a and t <= _2359:
        return 3
    b = _seconds(shifts.end3)
    if t >= 0 and t <= b:
        return 4
    return 5


def yea_dayr_check(dtime: str, *, now: datetime, facility: str, shifts: ShiftTimes) -> str:
    """Worker.YeaDayrCheck (Worker:143-164)。直を考慮した日付 `yyyy.mm.dd`(夜勤の 0 時〜End3 は前日)。"""
    today = now.strftime("%Y.%m.%d")
    if facility not in SHIFT_FACILITIES:
        return today
    if _shift_branch(dtime, shifts) == 4:
        return (now - timedelta(days=1)).strftime("%Y.%m.%d")
    return today


def day_check(dtime: str, *, now: datetime, facility: str, shifts: ShiftTimes) -> str:
    """Worker.DayCheck (Worker:204-231)。`yyyy.mm.dd  hh:nn`。前日扱いは `前日  (時+24):nn`(例 27:15)。"""
    today = now.strftime("%Y.%m.%d  %H:%M")
    if facility not in SHIFT_FACILITIES:
        return today
    if _shift_branch(dtime, shifts) == 4:
        dt = now - timedelta(days=1)
        return dt.strftime("%Y.%m.%d") + "  " + str(now.hour + 24) + dt.strftime(":%M")
    return today


def time_check(dtime: str, *, facility: str, shifts: ShiftTimes) -> str:
    """Worker.TimeCheck (Worker:251-272)。"日勤" / "1直" / "2直" / "3直"。

    1直は Start1 と**同じ時刻を含まない**(7:00 ちょうどは 3直・前日扱い)。
    """
    if facility not in SHIFT_FACILITIES:
        return "日勤"
    br = _shift_branch(dtime, shifts)
    return {1: "1直", 2: "2直"}.get(br, "3直")


def hhmm(now: datetime) -> str:
    """`Format$(Now, "hh:mm")`。"""
    return now.strftime("%H:%M")


def hhmmss(now: datetime) -> str:
    """`Format$(Now, "hh:mm:ss")`。"""
    return now.strftime("%H:%M:%S")


# ------------------------------------------------------------------
# 社員カード → 名簿
# ------------------------------------------------------------------
def _resolve_col(db: str, table: str, name: str) -> tuple[Optional[str], list[str], int]:
    """ADO の Filter の列名は大文字小文字を区別しない(VBA は [No]、表は NO)。実際の列名を探す。"""
    cols, rows = datasource.get_table(db, table)
    for c in cols:
        if c.casefold() == name.casefold():
            return c, cols, len(rows)
    return None, cols, len(rows)


def name_conv(card: str, *, test_line: bool = False) -> tuple[str, list[Message]]:
    """Worker.NameConv (Worker:37-80)。13 桁のカード番号 → 名簿の名前。未登録は "" と音声。

    - Test ラインは参照用ファイル(PackingPath)、ほかは品証やり取り(QualityPath)の 伝送用ファイル
    - ファイルが無ければ(品証やり取りのときだけ)参照用ファイルを見る。両方無ければメッセージで ""
    """
    msgs: list[Message] = []
    if card == "":
        return "", msgs
    key = "packing" if test_line else "quality"
    fpath = paths.get(key)
    db = paths.db_file("伝送用ファイル", base_key=key)
    if not datasource.exists(db):
        if fpath == paths.get("quality"):
            db = paths.db_file("伝送用ファイル", base_key="packing")
        if not datasource.exists(db):
            msgs.append(Message("日稲:" + "伝送用ファイル.accdb" + "がありません" + "_ReadCheck_NameConv",
                                ERROR, source="Worker:52"))
            return "", msgs
    col, cols, n = _resolve_col(db, "名簿", "No")
    records = datasource.get_records(db, "名簿", f"[{col or 'No'}]='{card}'") if col else None
    if n == 0:                                    # GetFieldsArr が Empty
        msgs.append(Message("fieldsなし" + "_NameConv", INFO, source="Worker:63"))
        return "", msgs
    c_name = datasource.column_index(cols, "名前")
    if records is None:
        msgs.append(Message("", INFO, sound=UNREGISTERED_SOUND, source="Worker:74"))
        return "", msgs
    if c_name < 0:
        raise IndexError("名簿 に 名前 列がありません(VBA の実行時エラー 9)")
    return records[0][c_name], msgs


@dataclass
class WorkerRegistry:
    """UF1 の検査員欄(T1-T3 = 名前を _ でつなぐ、Day1-3 = "yyyy.mm.dd  hh:mm")。"""
    t1: str = ""
    t2: str = ""
    t3: str = ""
    day1: str = ""
    day2: str = ""
    day3: str = ""


@dataclass
class RegisterResult:
    registry: WorkerRegistry
    name: str                         # 名簿の名前(未登録は "")
    shift: str = ""                   # 登録した直
    log: Optional[legacy_log.LogEntry] = None
    messages: list[Message] = field(default_factory=list)
    after: tuple[str, ...] = ()       # 呼び手が続けて行うこと(AdvanceCheck・Daily_Report(1))


def register_card(card: str, registry: WorkerRegistry, *, now: datetime, facility: str, shifts: ShiftTimes,
                  line: str, name_lookup: Callable[[str], tuple[str, list[Message]]]) -> Optional[RegisterResult]:
    """UF1.Worker_Change (UF1:4788-4859) の中身。カードが 13 桁のときだけ名簿を引いて直の欄に足す。

    13 桁でなければ None(何もしない)。line は UF1 の LineName(号機なし)。
    ログは Test だけアプリのフォルダの Access.log(UF1:4842 `WriteLog(Temp, "")`)。
    """
    if len(card) != 13:
        return None
    name, msgs = name_lookup(card)
    reg = replace(registry)
    if name == "":
        return RegisterResult(reg, "", messages=msgs)
    shift = time_check(hhmmss(now), facility=facility, shifts=shifts)
    day = day_check(hhmm(now), now=now, facility=facility, shifts=shifts)
    if shift in ("日勤", "1直"):
        reg.day1 = day
        reg.t1 = name if reg.t1 == "" else reg.t1 + "_" + name
        temp = "*UFKen:検査員登録_" + reg.day1 + ":" + reg.t1
    elif shift == "2直":
        reg.day2 = day
        reg.t2 = name if reg.t2 == "" else reg.t2 + "_" + name
        temp = "*UFKen:検査員登録_" + reg.day2 + ":" + reg.t2
    else:
        reg.day3 = day
        reg.t3 = name if reg.t3 == "" else reg.t3 + "_" + name
        temp = "*UFKen:検査員登録_" + reg.day3 + ":" + reg.t3
    ln = lines.find(line)
    entry = legacy_log.LogEntry(temp) if (ln and ln.official == "Test") else legacy_log.line_entry(line, temp)
    return RegisterResult(reg, name, shift, entry, msgs, ("AdvanceCheck", "Daily_Report(1)"))


# ------------------------------------------------------------------
# 検査員シート(検査員_<L> の A1:B3)
# ------------------------------------------------------------------
def epual(wor: list[str]) -> bool:
    """Worker.Epual (Worker:498-508)。全部同じなら True。"""
    for i in range(1, len(wor)):
        if wor[i] != wor[i - 1]:
            return False
    return True


def worker_save_rows(reg: WorkerRegistry) -> list[CellWrite]:
    """Worker.WorkerSave (Worker:513-565)。A1:B3 = [Day1,T1],[Day2,T2],[Day3,T3]。全員同じ名前なら 1 人に。"""
    out: list[CellWrite] = []
    for row, (day, t) in enumerate(((reg.day1, reg.t1), (reg.day2, reg.t2), (reg.day3, reg.t3)), start=1):
        if t != "":
            parts = t.split("_")
            if epual(parts):
                t = parts[-1]
        out += [CellWrite(row, 1, day), CellWrite(row, 2, t)]
    return out


def worker_load(cells: Cells) -> WorkerRegistry:
    """Worker.WorkerLoad (Worker:566-583)。Day は `Format(…, "yyyy.mm.dd  hh:mm")`(日付型のときだけ書式が効く)。"""
    def day(v: Any) -> str:
        if isinstance(v, datetime):
            return v.strftime("%Y.%m.%d  %H:%M")
        return _text(v)
    return WorkerRegistry(t1=_text(cells.get((1, 2))), t2=_text(cells.get((2, 2))), t3=_text(cells.get((3, 2))),
                          day1=day(cells.get((1, 1))), day2=day(cells.get((2, 1))), day3=day(cells.get((3, 1))))


def _date_ordinal(text: str) -> float:
    """`Day1 = Replace(Day1str, ".", "/")`(Date へ代入)。読めなければ On Error Resume Next で 0 のまま。"""
    m = re.match(r"^\s*(\d{4})/(\d{1,2})/(\d{1,2})\s*$", text)
    if not m:
        return 0.0
    try:
        return float(datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).toordinal())
    except ValueError:
        return 0.0


def worke_day_max(a1: Any, a2: Any, a3: Any) -> str:
    """Worker.WorkeDayMax (Worker:585-704)。検査員シート A1:A3 から、いちばん遅い登録の直 "1"/"2"/"3"。

    日付(先頭 10 文字)が同じなら時(13-14 文字目)で比べる。**分は比べない**。同じ値なら T1→T2→T3 の順。
    どれも無ければ ""。
    """
    s1, s2, s3 = _text(a1), _text(a2), _text(a3)
    d1s, d2s, d3s = s1[:10], s2[:10], s3[:10]

    def hour(s: str) -> int:
        if s == "":
            return 0
        h = vbx.mid(s, 13, 2)
        try:
            return vbx.clng(h)
        except vbx.TypeMismatch:
            return 0                                  # On Error Resume Next
    t1, t2, t3 = hour(s1), hour(s2), hour(s3)
    d1 = _date_ordinal(d1s.replace(".", "/")) if d1s else 0.0
    d2 = _date_ordinal(d2s.replace(".", "/")) if d2s else 0.0
    d3 = _date_ordinal(d3s.replace(".", "/")) if d3s else 0.0
    e1, e2, e3 = d1s != "", d2s != "", d3s != ""

    def pick(pairs: list[tuple[float, str]]) -> str:
        mx = max(v for v, _ in pairs)
        for v, k in pairs:
            if v == mx:
                return k
        return ""
    if e1 and not e2 and not e3:
        return "1"
    if not e1 and e2 and not e3:
        return "2"
    if not e1 and not e2 and e3:
        return "3"
    if e1 and e2 and not e3:
        return pick([(t1, "1"), (t2, "2")]) if d1s == d2s else pick([(d1, "1"), (d2, "2")])
    if e1 and not e2 and e3:
        return pick([(t1, "1"), (t3, "3")]) if d1s == d3s else pick([(d1, "1"), (d3, "3")])
    if not e1 and e2 and e3:
        return pick([(t2, "2"), (t3, "3")]) if d2s == d3s else pick([(d2, "2"), (d3, "3")])
    if e1 and e2 and e3:
        if d1s == d2s and d1s == d3s:
            return pick([(t1, "1"), (t2, "2"), (t3, "3")])
        mx = max(d1, d2, d3)
        if d1s == d2s:
            return "3" if mx == d3 else pick([(t1, "1"), (t2, "2")])
        if d1s == d3s:
            return "2" if mx == d2 else pick([(t1, "1"), (t3, "3")])
        if d2s == d3s:
            return "1" if mx == d1 else pick([(t2, "2"), (t3, "3")])
        return pick([(d1, "1"), (d2, "2"), (d3, "3")])
    return ""


def get_worker_name(day_max: str, reg: WorkerRegistry) -> str:
    """Worker.GetWorkerName (Worker:707-717)。最新の直の T の `_` 区切りの末尾。

    その直の T が空だと VBA は `Worker(UBound(Worker))` で実行時エラー 9 → IndexError。
    """
    t = {"1": reg.t1, "2": reg.t2, "3": reg.t3}.get(day_max)
    if t is None:
        return ""
    if t == "":
        raise IndexError("GetWorkerName: 直の検査員が空(VBA の実行時エラー 9)")
    return t.split("_")[-1]


def workers_text(day_max: str, reg: WorkerRegistry) -> str:
    """Aインプットのログに付ける ` T1:<T1>`(UFAinput:1035-1039、Input:339-343)。該当なしは ""。"""
    return {"1": " T1:" + reg.t1, "2": " T2:" + reg.t2, "3": " T3:" + reg.t3}.get(day_max, "")


def register_check(st_time: str, reg: WorkerRegistry, sheet_a: Optional[tuple[Any, Any, Any]], *, now: datetime,
                   facility: str, shifts: ShiftTimes, n: str = "") -> bool:
    """Worker.Register_Check (Worker:456-497)。True なら検査員の登録が必要。

    sheet_a は 検査員_<L> の A1:A3(シートが無ければ None)。
    """
    day_return = yea_dayr_check(hhmm(now), now=now, facility=facility, shifts=shifts)
    if sheet_a is None:
        return True
    if n == "新規":
        for v in sheet_a:
            s = _text(v)
            if s != "" and s[:10] != day_return:
                return True
        return False
    set_time = time_check(st_time, facility=facility, shifts=shifts)
    t, d = {"日勤": (reg.t1, reg.day1), "1直": (reg.t1, reg.day1),
            "2直": (reg.t2, reg.day2), "3直": (reg.t3, reg.day3)}[set_time]
    result = t == "" or d == ""
    if d[:10] != day_return:
        result = True
    return result


# ------------------------------------------------------------------
# 作業時間(WorkerTime_<L>)
# ------------------------------------------------------------------
def _vba_date_text(dt: datetime) -> str:
    """Date を文字列にしたとき(CStr。日本語の既定 "yyyy/MM/dd H:mm:ss")。"""
    return f"{dt.year:04d}/{dt.month:02d}/{dt.day:02d} {dt.hour}:{dt.minute:02d}:{dt.second:02d}"


def end_time_register(now: datetime, shifts: Optional[ShiftTimes] = None) -> tuple[str, CellWrite]:
    """Worker.EndTime_Register (Worker:312-326)。終了時刻と、WorkerTime_<L>!C1 に書く値。

    0:00〜7:00(7:00 を含む)は**前日の日付で時刻はそのまま**。7:00 は DB を使わない固定値。
    B38: 直すときは境界を DB の End3 にする(不良要因票と同じ)。
    """
    if quirks.reproduce("B38"):
        bound = 7 * 3600
    else:
        bound = _seconds((shifts or ShiftTimes()).end3)
    t = now.hour * 3600 + now.minute * 60 + now.second
    if t > bound and t <= _2359:
        text = _vba_date_text(now)
    elif t >= 0 and t <= bound:
        text = (now - timedelta(days=1)).strftime("%Y/%m/%d %H:%M:%S")
    else:
        text = _vba_date_text(now)
    return text, CellWrite(1, 3, text)


_DT_RE = re.compile(r"^\s*(\d{4})[/-](\d{1,2})[/-](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?\s*$")


def _cdate(v: Any) -> datetime:
    if isinstance(v, datetime):
        return v
    m = _DT_RE.match("" if v is None else str(v))
    if not m:
        raise vbx.TypeMismatch(f"CDate({v!r})")
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    h, mi, s = int(m.group(4) or 0), int(m.group(5) or 0), int(m.group(6) or 0)
    return datetime(y, mo, d, h, mi, s)


def is_over_12_hours(c1: Any, now: datetime, *, sheet_exists: bool = True) -> bool:
    """Worker.IsOver12Hours (Worker:328-350)。WorkerTime_<L>!C1(終了時刻)からの経過。

    B19: VBA は `DateDiff("h") >= 1`(時の境界を 1 つでもまたげば True。名前と違い 1 時間)。
    直すときは実際の経過が 12 時間以上。
    """
    if not sheet_exists or _empty(c1):
        return False
    target = _cdate(c1)
    if quirks.reproduce("B19"):
        hours = (now.replace(minute=0, second=0, microsecond=0)
                 - target.replace(minute=0, second=0, microsecond=0)).total_seconds() // 3600
        return hours >= 1
    return (now - target) >= timedelta(hours=12)


@dataclass
class NameMatchResult:
    writes: list[CellWrite]
    clear_columns: tuple[int, int] = (14, 17)     # N:Q を ClearContents
    out: bool = False                             # 不一致あり → UFNameList(1)
    open_name_list: Optional[int] = None
    save_workbook: bool = True                    # ThisWorkbook.Save(不一致の有無に関係なく)


def name_match(cells: Cells, reg: WorkerRegistry) -> NameMatchResult:
    """Worker.Name_Match (Worker:352-455)。WorkerTime の測定者が登録検査員(T1+T2+T3)にいるか。

    書く場所: N(番号)・O(作業済検査員履歴)・P(登録検査員)・Q(一致無し検査員)・R1="OUT"。
    B21: VBA は 3-7 行(S〜E)× 1-12 列だけ読み、F 面(8 行)と 13-14 列を見ない。R1 の "OUT" を消さない。
    """
    fix = not quirks.reproduce("B21")
    last_r, last_c = (8, 14) if fix else (7, 12)
    writes: list[CellWrite] = []
    if fix:
        writes.append(CellWrite(1, 18, ""))
    done: dict[str, int] = {}
    for c in range(1, last_c + 1):                 # 列 → 行の順(VBA の For i(列) / For a(行))
        for r in range(3, last_r + 1):
            v = _text(cells.get((r, c)))
            if v != "" and v not in done:
                done[v] = len(done) + 1
    writes.append(CellWrite(1, 14, None))
    writes.append(CellWrite(1, 15, "作業済検査員履歴"))
    for j, (name, no) in enumerate(done.items(), start=2):
        writes += [CellWrite(j, 14, no), CellWrite(j, 15, name)]
    # 登録検査員: Split(Join(Split(T1,"_")) & " " & …)(名前に半角空白があると分かれる)
    joined = " ".join(t.replace("_", " ") for t in (reg.t1, reg.t2, reg.t3))
    registered: dict[str, int] = {}
    for w in joined.split(" "):
        if w != "" and w not in registered:
            registered[w] = len(registered) + 1
    writes.append(CellWrite(1, 16, "登録検査員"))
    for k, name in enumerate(registered, start=2):
        writes.append(CellWrite(k, 16, name))
    writes.append(CellWrite(1, 17, "一致無し検査員"))
    k = 1
    out = False
    for name in done:
        if name not in registered:
            k += 1
            writes.append(CellWrite(k, 17, name))
            writes.append(CellWrite(1, 18, "OUT"))
            out = True
    return NameMatchResult(writes, out=out, open_name_list=1 if out else None)


# ------------------------------------------------------------------
# 日報(Daily_<L>)
# ------------------------------------------------------------------
@dataclass
class AinputValues:
    """Daily_Report(4) が読む UFAinput の値。"""
    lot: str = ""
    konno: str = ""
    mai1: str = ""
    cou1: str = ""
    mai2: str = ""
    cou2: str = ""


@dataclass
class DailyContext:
    """Daily_Report が読むもの。

    mode: "UF1" / "全ヵ所" / "キャスク" / "H494"(表示中のシート)。UF1 以外は Lname="" で
    `Daily_` シートを探す(実際は無いので何もしない)。
    uf1_line: UF1 で選んでいるライン(L1Work/LVC/HVC/Bala/Tott のどれか。号機なし)。
    sub_line / sub_kname: UFMitsubishi・UFCask の LineName と Kname。
    history_rows: Lot情報履歴 の 2 行目から(列は 1 始まりの位置で row[列-1])。None はシートなし。
    """
    exists: bool
    cells: dict[tuple[int, int], Any]
    history_rows: Optional[list[list[Any]]]
    working_lot: str
    now: datetime
    facility: str
    shifts: ShiftTimes
    registry: WorkerRegistry = field(default_factory=WorkerRegistry)
    mode: str = "UF1"
    uf1_line: str = ""
    sub_line: str = ""
    sub_kname: str = ""
    ainput: AinputValues = field(default_factory=AinputValues)


def daily_sheet_name(mode: str, lname: str) -> str:
    """Daily_Report が使うシート名(Worker:747-750)。UF1 以外は "Daily_"。"""
    return "Daily_" + (lname if mode == "UF1" else "")


def _hist(row: list[Any], col: int) -> Any:
    return _cell(row[col - 1]) if col - 1 < len(row) else None


def _two(n: int) -> str:
    """`Select Case Len(Hour(Now)): Case 1: "'0" & h  Case 2: "'" & h`(先頭の ' は Excel の文字列の印)。"""
    s = str(n)
    return "0" + s if len(s) == 1 else s


def _daily_line_b1(ctx: DailyContext) -> Optional[str]:
    if ctx.mode in ("全ヵ所", "H494"):
        return {"L-1": "L-1", "LVC": "LVC", "HVC": "HVC", "ﾊﾞﾗﾝｻｰ": "ﾊﾞﾗﾝｻｰ", "ﾄｯﾄ": "ﾄｯﾄ"}.get(ctx.sub_line, "Test")
    if ctx.mode == "キャスク":
        return {"L-1": "L-1", "ﾊﾞﾗﾝｻｰ": "ﾊﾞﾗﾝｻｰ", "ﾄｯﾄ": "ﾄｯﾄ"}.get(ctx.sub_line)
    ln = lines.find(ctx.uf1_line)
    names = {"L-1": "L-1", "LVC": "LVC", "HVC": "HVC", "バランサー": "ﾊﾞﾗﾝｻｰ", "トット": "ﾄｯﾄ"}
    return names.get(ln.official if ln else "", "Test")


def daily_report(n: int, ctx: DailyContext) -> list[CellWrite]:
    """Worker.Daily_Report (Worker:745-918)。日報に書くセル(書いた順)。ctx.cells も書き換える。

    1: A1=日付_直、B1=ライン、C1=当直の検査員 / 2: ロット情報 / 3: 開始時刻 / 4: 終了時刻・梱包No・枚数・包数
    B18: (3) の開始の分が 1 桁だと F 列(時)に書く。直すと G 列。
    Lot情報履歴 が無いと VBA は空のシートを作って配列の範囲外(実行時エラー 9)になる → IndexError。
    """
    writes: list[CellWrite] = []
    if not ctx.exists:
        return writes
    if ctx.history_rows is None:
        raise IndexError("Lot情報履歴 がありません(VBA は空のシートを作り実行時エラー 9)")
    cells = ctx.cells

    def put(r: int, c: int, v: Any) -> None:
        writes.append(CellWrite(r, c, v))
        cells[(r, c)] = v

    lr = last_row(cells, 4) + 1
    x = None
    for row in ctx.history_rows:
        if _veq(_hist(row, 3), ctx.working_lot):
            x = row
            break
    if x is None:
        return writes
    now = ctx.now
    if n == 1:
        put(1, 1, yea_dayr_check(hhmm(now), now=now, facility=ctx.facility, shifts=ctx.shifts) + "_"
            + time_check(hhmmss(now), facility=ctx.facility, shifts=ctx.shifts))
        b1 = _daily_line_b1(ctx)
        if b1 is not None:
            put(1, 2, b1)
        if ctx.mode in ("全ヵ所", "H494", "キャスク"):
            put(1, 3, ctx.sub_kname)
        else:
            sh = time_check(hhmm(now), facility=ctx.facility, shifts=ctx.shifts)
            reg = ctx.registry
            put(1, 3, {"日勤": reg.t1, "1直": reg.t1, "2直": reg.t2, "3直": reg.t3}[sh])
    elif n == 2:
        put(lr, 1, _hist(x, 3))
        put(lr, 2, _text(_hist(x, 8)) + "-" + _text(_hist(x, 9)))
        put(lr, 3, vbx.fmt(_hist(x, 4), "0.000") + "×" + vbx.fmt(_hist(x, 5), "0.0") + "×"
            + vbx.fmt(_hist(x, 6), "0.0"))
        put(lr, 5, _hist(x, 15))
        ai = _hist(x, 17)
        if _veq(ai, 0):
            put(lr, 11, "無")
        elif _veq(ai, 1):
            put(lr, 11, "有")
        put(lr, 25, vbx.fmt(_hist(x, 21), "0.00"))
        if _veq(_hist(x, 18), 1):
            a, b = _text(_hist(x, 19)), _text(_hist(x, 20))
            if a != "" and b != "":
                put(lr, 14, "A" + ":" + a + "_" + "B" + ":" + b)
            elif a != "" and b == "":
                put(lr, 14, "A" + ":" + a)
            elif a == "" and b != "":
                put(lr, 14, "B" + ":" + b)
            else:
                put(lr, 14, "")
        if ctx.mode == "キャスク":
            put(lr, 15, "キャスク")
    elif n == 3:
        writes += daily_report(1, ctx)
        if _empty(cells.get((lr, 1))):
            writes += daily_report(2, ctx)
        put(lr, 6, _two(now.hour))
        mi = str(now.minute)
        if len(mi) == 1:
            put(lr, 6 if quirks.reproduce("B18") else 7, "0" + mi)
        else:
            put(lr, 7, mi)
    elif n == 4:
        if _empty(cells.get((lr, 1))):
            writes += daily_report(2, ctx)
        a = ctx.ainput
        put(lr + 1, 1, a.lot)
        put(lr, 4, a.konno)
        if a.cou2 != "":
            put(lr + 1, 8, _two(now.hour))
            put(lr + 1, 9, _two(now.minute))
            put(lr, 12, _num(vbx.val(a.mai1)))
            put(lr, 13, _num(vbx.val(a.cou1)))
            put(lr + 2, 1, a.lot)
            put(lr + 1, 4, a.konno)
            put(lr + 1, 12, _num(vbx.val(a.mai2)))
            put(lr + 1, 13, _num(vbx.val(a.cou2)))
            put(lr + 2, 6, _two(now.hour))
            put(lr + 2, 7, _two(now.minute))
        else:
            put(lr + 1, 1, a.lot)
            put(lr, 4, a.konno)
            put(lr, 8, _two(now.hour))
            put(lr, 9, _two(now.minute))
            put(lr, 12, _num(vbx.val(a.mai1)))
            put(lr, 13, _num(vbx.val(a.cou1)))
            put(lr + 1, 6, _two(now.hour))
            put(lr + 1, 7, _two(now.minute))
    return writes


@dataclass
class DailyCreateResult:
    recreate: bool                 # AddWS(シートを作り直す。前の直の日報は消える)
    writes: list[CellWrite]


def daily_create(ctx: DailyContext) -> DailyCreateResult:
    """Worker.Daily_Create (Worker:719-743)。直が変わっていたら日報シートを作り直し、見出しと (1)(3) を書く。

    A1 との比較は YeaDayrCheck("hh:mm") と TimeCheck("hh:mm:ss") の組み合わせ(7:00 台の 1 分間は
    「前日_1直」になる)。シートがあっても毎回 (1)(3) を書く(VBA どおり)。
    """
    day_return = yea_dayr_check(hhmm(ctx.now), now=ctx.now, facility=ctx.facility, shifts=ctx.shifts)
    set_time = time_check(hhmmss(ctx.now), facility=ctx.facility, shifts=ctx.shifts)
    recreate = (not ctx.exists) or _text(ctx.cells.get((1, 1))) != day_return + "_" + set_time
    if recreate:
        ctx.cells.clear()
    ctx.exists = True
    writes = [CellWrite(2, c, v) for c, v in enumerate(DAILY_FIELDS, start=1)]
    for w in writes:
        ctx.cells[(w.row, w.col)] = w.value
    writes += daily_report(1, ctx)
    writes += daily_report(3, ctx)
    return DailyCreateResult(recreate, writes)


@dataclass
class WriteBack:
    sheet: str
    clear: tuple[int, int, int, int]      # (行1, 列1, 行2, 列2) を ClearContents
    writes: list[CellWrite]


def name_list_write_back(items: list[list[str]], lname: str, col_count: int) -> WriteBack:
    """UFNameList.CommandButton3_Click (UFNameList:278-293) / ListView.ListPosting (ListView:504-521)。

    items は ListView の行(各行の SubItems(1..) = 日報の 1 列目から)。col_count は `UC(sh, 2)`。
    B22: VBA の CommandButton3 は `Daily_Test` に固定で書く。直すと `Daily_<ライン>`(ListPosting と同じ)。
    """
    sheet = "Daily_Test" if quirks.reproduce("B22") else "Daily_" + lname
    cont = len(items)
    writes = []
    for i, row in enumerate(items, start=1):
        for a in range(1, col_count + 1):
            writes.append(CellWrite(i + 2, a, row[a - 1] if a - 1 < len(row) else ""))
    return WriteBack(sheet, (3, 1, 20 + cont, col_count), writes)


def list_posting(items: list[list[str]], lname: str, col_count: int) -> WriteBack:
    """ListView.ListPosting (ListView:504-521)。UFNameList の「日報修正」(CommandButton2)は `Daily_<ライン>` に書く。"""
    cont = len(items)
    writes = [CellWrite(i + 2, a, row[a - 1] if a - 1 < len(row) else "")
              for i, row in enumerate(items, start=1) for a in range(1, col_count + 1)]
    return WriteBack("Daily_" + lname, (3, 1, 20 + cont, col_count), writes)
