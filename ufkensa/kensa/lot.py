"""ロット引当と仕掛 DB の切替(VBA 【UFKen_01】AccHIki・【UFKen_03】SaveLoad の一部・補助/Worker の時刻判定)

VBA の対応(行番号は vba_src の各ファイル):

| VBA                                   | ここ                       |
|---------------------------------------|----------------------------|
| 補助.ReadCheck (補助:454-464)          | `read_check`               |
| 補助.TimeDesignation / FileDesignation | 移さない(D26 改: 上流の作り直しの決まりは使わない) |
| Worker.IsOver12Hours (Worker:328-349)  | `is_over_12_hours`(B19)  |
| AccHIki.Arrbox (AccHIki:14-92)         | `arrbox`                   |
| AccHIki.GetLotInfo (AccHIki:644-704)   | `get_lot_info`             |
| SaveLoad.SikaAtuCreate (SaveLoad:16-54)| `sika_atu_create`          |
| SaveLoad.AccReflec (SaveLoad:56-573)   | `acc_reflec`(DB を読んで画面に出す値まで)、`materials_order_no`、`dedupe_message` |
| SaveLoad.HistorySave (SaveLoad:1258-1552) | `history_save`(B15・B16・B49・B50・B51) |
| SaveLoad.VC_Picture (SaveLoad:1554-1758)  | `vc_picture`             |
| Input.VcConv (Input:753-800)            | `vc_conv`(VC厚 品名ｺｰﾄﾞ → 厚-係数) |

【仕掛 DB の切替は 3 系統あり、条件が違う(G3 §3.1.1)。まとめずにそれぞれ書く】
1. Arrbox: DBflag が空なら MainPath。MainPath のときは ReadCheck 失敗で NextPath へ
2. AccReflec: DBflag → NextPath、それ以外 MainPath
3. HistorySave: 全ヵ所/キャスク検査の表示中 → MainPath、DBflag → NextPath、それ以外 MainPath。
   そのあと ReadCheck で切替(NextPath を選んでいたら「ありません」で抜ける。B49)
【D26 改(2026-11-03 のご指示)】VBA の上流の作り直しの決まり(TimeDesignation = 毎時 59:46〜00:45、
FileDesignation / AccReflec / HistorySave の PLEASE WAIT.txt)は**外した**。ReadCheck は仕掛 DB を
手元へ写せたか(`paths.lot_usable`。写せなければ読まない)。第1にロットが無いときの第2は `lot_from_second`。
4. SikaAtuCreate: MainPath で列名も行も取れなければ NextPath(ReadCheck・時刻窓は見ない)

【パスの比べ方】VBA は `FPath = MainPath` を文字列で比べる。オフラインでは MainPath と NextPath が
同じ文字列になるので、その場合の分岐の順もそのまま再現する。

【時刻】`now` を引数で受け取る(None なら現在時刻)。

【ログ】VBA の `WriteLog(msg, Sepl, TargetPath)` は `LogLine` として返す(書くのは WriteLog の担当)。

【DB の Null】受領 DB(全列 Text)は値の無い所が Null。VBA の `Val(Null)` や String 変数への代入は
実行時エラー 94 になるはずだが、実データでは SIKALOTNOW の `板厚公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ` が 83 行中 80 行 Null で、
そのとおりなら本番でほぼ全ロットがエラーになる。実機確認まで **Null は "" として読み**、VBA なら
エラー 94 になる所を `LotReflect.null_uses` に順に記録する(先頭が VBA の止まる所)。
"""
from __future__ import annotations

import datetime as _dt
import re
import sqlite3
from dataclasses import dataclass, field
from typing import Any, Callable, Optional, Sequence

from . import ado_filter, datasource, flow, paths, vbx
from .model import UiValues
from .flow import ERROR, INFO, WARN, Message

TABLE = "仕掛"
VC_TABLE = "VC厚"
DENSO = "伝送用ファイル"

#: VBA の Err.Description(AccReflec の ErrHandler が「DB読み込みエラー:」の後に出す)
ERR_TEXT = {
    9: "インデックスが有効範囲にありません。",
    13: "型が一致しません。",
    94: "Null の使い方が不正です。",
}


class SubscriptOutOfRange(IndexError):
    """VBA の実行時エラー 9(列が見つからず添字 0 で配列を引いた、など)。"""


class DbReadError(RuntimeError):
    """ADO のエラー(列が無い 3265・SQL の構文エラーなど)。AccReflec の ErrHandler に入る。"""


@dataclass(frozen=True)
class LogLine:
    """VBA `WriteLog(msg1, Sepl, TargetPath)` の依頼(WriteLog:6-38 が書く 1 行)。"""
    msg: str
    sepl: str = ","
    target_path: str = ""


def vba_error_number(exc: BaseException) -> int:
    """例外 → VBA のエラー番号(13 / 9 / それ以外 0)。"""
    if isinstance(exc, vbx.TypeMismatch):
        return 13
    if isinstance(exc, SubscriptOutOfRange):
        return 9
    return 0


def db_read_error(exc: BaseException) -> Message:
    """AccReflec の ErrHandler (SaveLoad:564-572) の MsgBox「DB読み込みエラー:」+ Err.Description。"""
    desc = ERR_TEXT.get(vba_error_number(exc), str(exc))
    return Message("DB読み込みエラー:" + "\r\n" + desc, ERROR, source="SaveLoad:572")


def vba_target(name: str) -> str:
    """メッセージに出す VBA のファイル名(`SIKAHIKINOW.accdb` など)。"""
    return name + ".accdb"


def db_path(base: str, name: str) -> str:
    """VBA の `FPath & "SIKA….accdb"`(D4 でファイルは .sqlite3)。

    D26: 仕掛の 3 ファイルは、第2(控え)のフォルダなら設定「2つ目の名前」で探す(`paths.lot_file_name`)。
    """
    if name in paths.LOT_DBS:
        return paths.join(base, paths.lot_file_name(name, base))
    return paths.join(base, paths.DB_FILES[name][1] if name in paths.DB_FILES else name + ".sqlite3")


def _remap(rows: list, src: Optional[list], dst: Optional[list]) -> list:
    """第2の行を第1の列の並びに直す(2つ目は列の数・並びが違うことがある。無い列は "")。"""
    if not src or not dst or list(src) == list(dst):
        return rows
    idx = [list(src).index(c) if c in src else -1 for c in dst]
    return [[r[i] if 0 <= i < len(r) else "" for i in idx] for r in rows]


def lot_from_second(target: str, flt: str, first_file: str) -> tuple[Optional[list], str]:
    """D26(ロットごとの第2): 第1の仕掛にロット(条件の行)が無いとき、第2(控え)から同じ条件で読む。

    VBA には無い(VBA は第1にファイルがあればそこで「該当なし」)。第1と第2が同じフォルダ
    (オフライン・1 か所に置いたとき)は何もしない。行は第1の列の並びに直して返す。
    戻り値は (行 / None, 読んだ第2のファイル)。
    """
    main, nxt = paths.get("main"), paths.get("next")
    if main == nxt:
        return None, ""
    second_file = db_path(nxt, target)
    if not read_check(second_file, 1):
        return None, ""
    try:
        rows = datasource.get_records(second_file, TABLE, flt)
        if rows is None:
            return None, second_file
        src = datasource.get_fields(second_file, TABLE, "")
        dst = datasource.get_fields(first_file, TABLE, "") if read_check(first_file, 1) else None
    except (datasource.DbMissing, datasource.TableMissing, sqlite3.Error):
        return None, ""
    return _remap(rows, src, dst), second_file


def merge_hiki_from_second(rows: Optional[list], flt: str, first_file: str) -> tuple[Optional[list], list[str]]:
    """2026-11-08(ご依頼): 第1の SIKAHIKI にロットの行があっても、引当が足りないことがある。

    第2(控え)にだけある引当番号の行を、第1の行の後ろに足す(第1の行はそのまま使う。第2の方が古いこともあるので
    第2を丸ごとは使わない)。オンラインで第1と第2が別の場所のときだけ。戻り値は (行, 足した引当番号)。
    """
    main, nxt = paths.get("main"), paths.get("next")
    if not rows or main == nxt or first_file != db_path(main, "SIKAHIKINOW"):
        return rows, []
    try:
        cols = datasource.get_fields(first_file, TABLE, "") or []
    except (datasource.DbMissing, datasource.TableMissing, sqlite3.Error):
        return rows, []
    i = datasource.column_index(cols, "引当番号")
    if i < 0:
        return rows, []
    extra, _second = lot_from_second("SIKAHIKINOW", flt, first_file)
    have = {vbx.cstr(r[i]) for r in rows if i < len(r)}
    add = [r for r in (extra or []) if i < len(r) and vbx.cstr(r[i]) not in have]
    if not add:
        return rows, []
    nos = sorted({vbx.cstr(r[i]) for r in add})
    from ..logging_utils import get_logger
    get_logger("kensa.lot").info("第1パスに無い引当を第2パスから足しました: %s (%s)", ", ".join(nos), flt)
    return list(rows) + add, nos


def hiki_added_message(nos: Sequence[str], source: str) -> Message:
    return Message("第1パスに無い引当を第2パスから足しました: 引当NO " + ", ".join(nos), INFO, source=source)


def fields_from_second(target: str) -> tuple[Optional[list], str]:
    """D26: 第1の仕掛の表が空(列名が取れない)とき、第2の列名。"""
    main, nxt = paths.get("main"), paths.get("next")
    if main == nxt:
        return None, ""
    second_file = db_path(nxt, target)
    if not read_check(second_file, 1):
        return None, ""
    try:
        return datasource.get_fields(second_file, TABLE, ""), second_file
    except (datasource.DbMissing, datasource.TableMissing, sqlite3.Error):
        return None, ""


def subpath(base: str, *parts: str) -> str:
    """`base & "VC種類画像\" & 名前` のようなつなぎ方。区切りは base に合わせる
    (paths.join は常に "\\" を足すので、試験の Linux のフォルダの下の階層に使えない。共通化の候補)。"""
    sep = "/" if ("/" in base and "\\" not in base) else "\\"
    out = base
    for part in parts:
        if out and not out.endswith(("\\", "/")):
            out += sep
        out += part
    return out


def _now(now: Optional[_dt.datetime]) -> _dt.datetime:
    return now if now is not None else _dt.datetime.now()


# ------------------------------------------------------------------
# Variant の比べ方(シートのセル・配列要素どうし)
# ------------------------------------------------------------------
def _is_num(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def variant_eq(a: Any, b: Any) -> bool:
    """Variant どうしの `=`(セル値と DB 値など)。数と文字列は等しくならない(エラーにもならない)。"""
    if a is None:
        a = "" if not _is_num(b) else 0
    if b is None:
        b = "" if not _is_num(a) else 0
    if _is_num(a) and _is_num(b):
        return float(a) == float(b)
    if isinstance(a, str) and isinstance(b, str):
        return a == b
    if isinstance(a, _dt.datetime) or isinstance(b, _dt.datetime):
        return a == b
    return False


def _cell_ne_empty(v: Any) -> bool:
    """`セル <> ""`。Empty は ""、数は "" と等しくない。"""
    if v is None:
        return False
    if isinstance(v, str):
        return v != ""
    return True


# ------------------------------------------------------------------
# 補助・Worker
# ------------------------------------------------------------------
def read_check(target: str, n: int = 1) -> bool:
    """補助.ReadCheck (補助:454-464)。n=1 ファイル、n=2 フォルダがあれば真。エラーは偽(On Error Resume Next)。"""
    try:
        p = paths.to_local(target)
        if n == 1:
            if paths.is_copied(target):
                # D26: 仕掛 DB は手元へ写せたときだけ「ある」(写せなければ読まない)
                return p.is_file() and paths.lot_usable(target)
            return p.is_file()
        if n == 2:
            return p.is_dir()
    except (OSError, ValueError):
        return False
    return False


_DATE_FORMATS = ("%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
                 "%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d %H:%M:%S", "%Y.%m.%d")


def cdate(value: Any) -> _dt.datetime:
    """VBA `CDate`(よく使う書き方だけ)。日付にできなければ型エラー 13。"""
    if isinstance(value, _dt.datetime):
        return value
    if isinstance(value, _dt.date):
        return _dt.datetime(value.year, value.month, value.day)
    text = "" if value is None else str(value).strip()
    text = re.sub(r"\s+", " ", text)
    for f in _DATE_FORMATS:
        try:
            return _dt.datetime.strptime(text, f)
        except ValueError:
            continue
    raise vbx.TypeMismatch(f"CDate({value!r})")


def _hour_diff(t1: _dt.datetime, t2: _dt.datetime) -> int:
    """VBA `DateDiff("h", t1, t2)`(時の境界をまたいだ数)。"""
    a = t1.replace(minute=0, second=0, microsecond=0)
    b = t2.replace(minute=0, second=0, microsecond=0)
    return int((b - a).total_seconds() // 3600)


def is_over_12_hours(end_time: Any, now: Optional[_dt.datetime] = None, *, sheet_exists: bool = True) -> bool:
    """Worker.IsOver12Hours (Worker:328-349)。`WorkerTime_<L>!C1`(終了時刻)から時の境界を
    1 つ以上またいだら真(真の間は「DB がありません」のメッセージを出さない)。

    - シートが無い・C1 が空 → 偽
    - C1 が日付にできない → 型エラー 13(VBA は On Error なし)

    B19: 名前は 12 時間だが VBA は `DateDiff("h") >= 1`。直した場合は `>= 12`。
    """
    if not sheet_exists:
        return False
    if end_time is None or (isinstance(end_time, str) and end_time == ""):
        return False
    target = cdate(end_time)
    diff = _hour_diff(target, _now(now))
    return diff >= 1


# ------------------------------------------------------------------
# Arrbox(AccHIki:14-92)
# ------------------------------------------------------------------
@dataclass
class ArrboxResult:
    """`Arrbox` の戻り値。`value` は Fields なら列名のリスト、Record なら行のリスト、None は VBA の Empty。"""
    value: Optional[list] = None
    file_path: str = ""
    messages: list[Message] = field(default_factory=list)
    logs: list[LogLine] = field(default_factory=list)


#: Arrbox の SQLfilter のキーワード → ADO Filter (AccHIki:42-47)
def arrbox_filter(sql_filter: str) -> str:
    if sql_filter == "":
        return ""
    if sql_filter == "板":
        return "[ｵｰﾀﾞｰ板丈]<>0"
    if sql_filter == "コイル":
        return "[ｵｰﾀﾞｰ板丈]=0  AND  [BOX設計_設備名] = 'KEN' "
    return "[ﾛｯﾄ番号]='" + ado_filter.quote(sql_filter) + "'"


def arrbox(target: str, sql_filter: str = "", use: str = "Record", path: str = "", *,
           now: Optional[_dt.datetime] = None, over12: bool = False) -> ArrboxResult:
    """AccHIki.Arrbox (AccHIki:14-92)。仕掛 DB(`target` = "SIKAHIKINOW" など)の表「仕掛」を読む。

    - `path` は VBA と同じく UF1.DBflag.Text を渡す(空なら MainPath、それ以外は NextPath)
    - `over12` は IsOver12Hours の結果(真なら「ありません」を出さない)
    - use="Fields" は Filter を掛けずに列名(AccHIki:54)。0 件なら None
    - L60 以降(空のとき NextPath を見直す処理)は到達しないので移さない
    """
    if use not in ("Record", "Fields"):
        raise ValueError(f"Arrbox の Use={use!r} は VBA でも使われていない")
    res = ArrboxResult()
    main, nxt = paths.get("main"), paths.get("next")
    fpath = main if path == "" else nxt
    file_path = db_path(fpath, target)
    if fpath == main:
        # D26 改: 時刻窓・PLEASE WAIT は見ない(ReadCheck = 手元へ写せたか だけ)
        if read_check(file_path, 1) is False:
            fpath = nxt
            file_path = db_path(fpath, target)
            if read_check(db_path(nxt, target), 1) is False:
                if not over12:
                    res.messages.append(Message(paths.missing_text(target),
                                                ERROR, source="AccHIki:33"))
                res.file_path = file_path
                return res
    res.file_path = file_path
    if use == "Record":
        res.value = datasource.get_records(file_path, TABLE, arrbox_filter(sql_filter))
    else:
        res.value = datasource.get_fields(file_path, TABLE, "")
    if use == "Record" and target == "SIKAHIKINOW" and fpath == main and res.value is not None:
        res.value, added = merge_hiki_from_second(res.value, arrbox_filter(sql_filter), file_path)
        if added:
            res.messages.append(hiki_added_message(added, "Arrbox"))
    if res.value is None and fpath == main:
        # D26: 第1にロット(行・列名)が無ければ第2を見る(VBA の L60 以降は到達しないので VBA には無い動き)
        if use == "Record":
            rows, second = lot_from_second(target, arrbox_filter(sql_filter), file_path)
        else:
            rows, second = fields_from_second(target)
        if rows is not None:
            res.value, res.file_path = rows, second
    return res


# ------------------------------------------------------------------
# GetLotInfo(AccHIki:644-704)
# ------------------------------------------------------------------
LOT_INFO_COLUMNS = ("納入先", "得意先", "用途ｺｰﾄﾞ", "材質", "調質", "厚", "巾", "丈", "ﾛｯﾄNo", "検査No")


def get_lot_info(sheet_rows: Optional[Sequence[Sequence[Any]]], lot: str) -> Optional[list[list[Any]]]:
    """AccHIki.GetLotInfo。`Lot情報履歴` の 1 行目(見出し)から列を探し、ロットの行を
    [納入先, 得意先, 用途ｺｰﾄﾞ, 材質, 調質, 厚, 巾, 丈, ﾛｯﾄNo, 検査No] で返す。

    `sheet_rows` はシートの 1 行目〜A 列の最終行(None はシートが無い)。見出しが 1 つでも
    足りない・該当なしは None(VBA の Empty)。見出し行も比較の対象に入る(VBA どおり)。
    """
    if sheet_rows is None:
        return None
    if len(sheet_rows) <= 1 and (len(sheet_rows) == 0 or len(sheet_rows[0]) <= 1):
        raise vbx.TypeMismatch("Lot情報履歴が空(HisArry() = A1 の 1 セル)AccHIki:650")
    header = list(sheet_rows[0])
    idx = {}
    for name in LOT_INFO_COLUMNS:
        # VBA は Select Case で後勝ち(同じ見出しが 2 つあれば右の列)
        found = -1
        for i, h in enumerate(header):
            if variant_eq(h, name) and isinstance(h, str):
                found = i
        idx[name] = found
    if any(v < 0 for v in idx.values()):
        return None

    def cell(row: Sequence[Any], i: int) -> Any:
        return row[i] if i < len(row) else None

    out = []
    for row in sheet_rows:
        if variant_eq(cell(row, idx["ﾛｯﾄNo"]), lot):
            out.append([cell(row, idx[n]) for n in LOT_INFO_COLUMNS])
    return out or None


# ------------------------------------------------------------------
# SikaAtuCreate(SaveLoad:16-54)
# ------------------------------------------------------------------
def sika_atu_create(lot: str) -> Optional[list[str]]:
    """SaveLoad.SikaAtuCreate。SIKALOTNOW のロットの全行の `ｵｰﾀﾞｰ板厚`(文字列)。無ければ None。

    MainPath で列名(表が空なら None)も行も取れなければ NextPath を見る。ReadCheck は見ない。
    ファイルが無いときは VBA の End に当たる `datasource.DbMissing`。
    """
    if lot == "":
        return None
    file_path = db_path(paths.get("main"), "SIKALOTNOW")
    fields = datasource.get_fields(file_path, TABLE, "")
    if fields is None:
        file_path = db_path(paths.get("next"), "SIKALOTNOW")
        fields = datasource.get_fields(file_path, TABLE, "")
        if fields is None:
            return None
    col = datasource.column_index(fields, "ｵｰﾀﾞｰ板厚")
    flt = "[ﾛｯﾄ番号]='" + ado_filter.quote(lot) + "'"
    rows = datasource.get_records(file_path, TABLE, flt)
    if rows is None:
        file_path = db_path(paths.get("next"), "SIKALOTNOW")
        rows = datasource.get_records(file_path, TABLE, flt)
        if rows is None:
            return None
        # D26: 第2は列の並びが違うことがある(2つ目の名前の圧縮版など)。第1の列の並びに直す(同じなら VBA と同じ)
        rows = _remap(rows, datasource.get_fields(file_path, TABLE, ""), fields)
    if col < 0:
        raise SubscriptOutOfRange("SikaAtuCreate: ｵｰﾀﾞｰ板厚 の列が無い (SaveLoad:51)")
    return [r[col] for r in rows]


# ------------------------------------------------------------------
# 小さな文字列関数(他モジュールの VBA。共通化の候補)
# ------------------------------------------------------------------
def num_extract(value: Any) -> str:
    """補助.numExtract (補助:1116-1130)。"/" の前だけを見て、半角の 0-9 と "." だけを残す。"""
    s = "" if value is None else str(value)
    if "/" in s:
        s = s.split("/")[0]
    return "".join(ch for ch in s if ch in "0123456789.")


def convert_to_half(value: Any) -> str:
    """補助.ConvertToHalf (補助:1105-1114)。全角の ０-９ だけ半角にしてから numExtract。"""
    s = "" if value is None else str(value)
    # 全角の数字と小数点「．」を半角に(VBA は「．」を落として "３．０" → "30" になっていた。2026-11-07 に直した)
    s = "".join(chr(ord(ch) - 0xFEE0) if "０" <= ch <= "９" else ("." if ch == "．" else ch) for ch in s)
    return num_extract(s)


def find_right_string(text: str) -> str:
    """OutTolerance.FindRightString (OutTolerance:2495-2506)。空白区切りで右から ANI/ANF/CAI/CAF を探す。"""
    if text == "":
        return ""
    for part in reversed(text.split(" ")):
        if part in ("ANI", "ANF", "CAI", "CAF"):
            return part
    return ""


def dedupe_message(temp_e: str) -> str:
    """AccReflec:419-427。カンマで分けて重複を除き(最初の順)、カンマでつなぎ直す。"""
    if temp_e == "":
        return ""
    seen: dict[str, int] = {}
    for part in temp_e.split(","):
        seen[part] = 0
    return ",".join(seen.keys())


# ------------------------------------------------------------------
# AccReflec(SaveLoad:56-573)
# ------------------------------------------------------------------
#: 該当ロットが無いとき(SkipNoItti, SaveLoad:347-374)に "" にする画面の項目
CLEAR_ON_NOT_FOUND = ("Zai", "Tyu", "Youto", "ATU", "Hab", "Tak", "Ap", "Am", "Hp", "Hm", "Tp", "Tm",
                      "ApK", "AmK", "HpK", "HmK", "TpK", "TmK", "TAIK", "MG", "HaHiz", "TaHiz", "Naka",
                      "Mimi", "HURIS")

#: AccReflec の最初に空にする画面の項目(SaveLoad:64-77。BackStyle の戻しも含む)
CLEAR_ON_START = ("TaikText", "LateralText", "LateralText.Tag", "Sumisei", "DebugTaikaku", "DebugLate",
                  "Additional", "Comment", "Reien", "KenH")

GAICHU = (("GFS", "GFSｺｰｽ有、進度にあった製品ｻｲｽﾞ入力をお願いします。"),
          ("GCT", "GCTｺｰｽ有、進度にあった製品ｻｲｽﾞ入力をお願いします。"),
          ("GSS", "GSSｺｰｽ有、進度にあった製品ｻｲｽﾞ入力をお願いします。"))

#: PastHiki(0..17) の列(SaveLoad:176-199)。外注のときは 2-4 が BOX実績
PAST_HIKI_COLUMNS = ("ｵｰﾀﾞｰ番号", "ﾛｯﾄ番号", "製造板厚", "製造板幅", "製造板丈", "検査番号", "製造材質", "製造調質",
                     "用途ｺｰﾄﾞ", "設計_設備ｺｰｽ", "取引先名", "納入先名", "送り先名", "BOX実績_枚本数",
                     "BOX実績_良品重量", "梱包_合紙", "VC", "品質ｸﾞﾚｰﾄﾞ_ﾗﾃﾗﾙﾎﾞｰ")


@dataclass
class LotScreen:
    """AccReflec が見る画面・設定の値。"""
    lot: str                      # UF1.Lot.Text(空なら何もしない。Arrbox の引当検索にも使う)
    cb_text: str = ""             # UF1.CB.Text(SQL の ﾛｯﾄ番号 に使う。VBA は Lot でなく CB)
    dbflag: str = ""              # UF1.DBflag.Text(空以外で控え NextPath)
    facility: str = "存在なし"    # レジストリ「設備名」(GetSetting の既定値 "存在なし")
    work_time: str = "S"          # "S" のときだけ HistorySave を呼ぶ


@dataclass
class LotReflect:
    """AccReflec が SIKALOTNOW から読んで画面に出す値。

    status: "ok" / "not_found"(SkipNoItti: `CLEAR_ON_NOT_FOUND` を空に) /
            "db_missing"(MsgBox を出して抜ける) / "error"(ErrHandler の MsgBox) / "no_lot"(Lot が空)
    設備仕様 NG は `flow.Stop` を投げる(呼び手は UF1.CB.Tag="end" と同じく以降を止める)。
    """
    status: str = "ok"
    db_path: str = ""
    row: dict[str, Optional[str]] = field(default_factory=dict)   # 採用した 1 行(None = Null)
    messages: list[Message] = field(default_factory=list)
    g_flag: bool = False                    # 外注(GFS/GCT/GSS)
    temp_e: str = ""                        # 注意メッセージ(この後 導電率の文言を足し、dedupe_message → MsgError(…,4))
    equipment_args: tuple = ()              # CheckEquipmentDb に渡した (設備名, 材質, 厚, 幅, 丈, 単重)
    past_hiki: list[Any] = field(default_factory=lambda: [None] * 19)
    form: dict[str, str] = field(default_factory=UiValues)   # 画面の値(Zai, Tyu, Youto, TAIK, …, TmK, KenH)
    last_hot: Optional[str] = None          # LastHot.Caption(None = 変えない。VBA は消さない)
    qf: Optional[str] = None                # QF.Caption(None = 変えない。VBA は消さない)
    sika_set_atu: float = 0.0               # SikaSetAtu(Single)。CBAdd に渡す
    shiji_taik: str = ""                    # 指示_直角度(原文)
    shiji_late: str = ""                    # 指示_ﾗﾃﾗﾙﾎﾞｰ(原文)
    taik_text: Optional[str] = None         # 指示があるときの TaikText(色 RGB(51,204,204))。None は等級から計算
    lateral_text: Optional[str] = None      # 指示があるときの LateralText(色付き、Tag="有")
    taik_flag: bool = False
    late_flag: bool = False
    additional: str = ""                    # Additional.Caption
    ken_no: str = ""                        # PastHiki(5)。指示ラテラルのとき LSelectLateral に渡す
    cache_ord_atu: Optional[float] = None   # g_CacheOrdAtu(Single)
    cache_c_course: Optional[str] = None    # g_CacheC_Course
    vc_action: str = ""                     # "clear"(VC_Clear) / "picture"(VC_Picture) / ""
    run_history_save: bool = False          # WorkTime="S" で HistorySave を呼ぶ
    history: Optional["HistorySaveResult"] = None
    null_uses: list[str] = field(default_factory=list)  # VBA ならエラー 94 になる所(先頭で VBA は止まる)


class _Rs:
    """ADO Recordset の 1 行。Null を "" で返し、VBA がエラー 94 になる使い方を記録する。"""

    def __init__(self, cols: Sequence[str], values: Sequence[Any], null_uses: list[str]) -> None:
        self.data: dict[str, Optional[str]] = {}
        for c, v in zip(cols, values):
            self.data[c] = None if v is None else datasource.as_text(v)
        self.null_uses = null_uses

    def value(self, col: str) -> Optional[str]:
        """`rs.Fields(col).Value`(Variant。None = Null)。列が無ければ ADO のエラー 3265。"""
        if col not in self.data:
            raise DbReadError("要求された名前、または序数に対応する項目がコレクションに見つかりません。")
        return self.data[col]

    def text(self, col: str, where: str) -> str:
        """String 変数への代入・Val・CStr に使う所。Null は "" にして記録する。"""
        v = self.value(col)
        if v is None:
            self.null_uses.append(f"{where} {col}")
            return ""
        return v

    def control(self, col: str) -> str:
        """画面のコントロール・Variant 配列・文字列連結に入れる所(Null でもエラーにならない)。"""
        v = self.value(col)
        return "" if v is None else v


def _half(s: str) -> str:
    """全角の英数記号を半角に(０．５ → 0.5、＋ → +)。"""
    return "".join(chr(ord(c) - 0xFEE0) if 0xFF01 <= ord(c) <= 0xFF5E else c for c in (s or ""))


def _select_lot(file_path: str, cb_text: str) -> tuple[list[str], Optional[tuple]]:
    """`SELECT * FROM 仕掛 WHERE ﾛｯﾄ番号 = '<CB>' AND ｵｰﾀﾞｰ板丈 <> '0'` の先頭行(SaveLoad:120-125)。

    Access の Text の `=` は大文字小文字を区別しない → COLLATE NOCASE。
    `ｵｰﾀﾞｰ板丈 <> '0'` は文字列 '0' との比較(全列 Text)。Null の行は残らない(SQLite も同じ)。
    行の順は取り込み順(rowid)。
    """
    from . import lot_columns
    with datasource.open_ro(file_path) as conn:
        # 2つ目のファイルは列名の書き方が違うことがある(全角カナなど)。実際の列名で引き、第1の名前で返す
        actual = [r[1] for r in conn.execute('PRAGMA table_info("仕掛")')]
        canon = lot_columns.canonical(actual)
        real = dict(zip(canon, actual))
        lot_col = real.get("ﾛｯﾄ番号", "ﾛｯﾄ番号").replace('"', '""')
        take_col = real.get("ｵｰﾀﾞｰ板丈", "ｵｰﾀﾞｰ板丈").replace('"', '""')
        sql = (f'SELECT * FROM "仕掛" WHERE "{lot_col}" = ? COLLATE NOCASE AND "{take_col}" <> ? '
               'ORDER BY rowid LIMIT 1')
        cur = conn.execute(sql, (cb_text, "0"))
        cols = lot_columns.canonical([d[0] for d in cur.description])
        row = cur.fetchone()
        if row is None:
            return cols, None
        cols, rows = lot_columns.pad(cols, [row])
        return cols, tuple(rows[0])


def acc_reflec_path(dbflag: str) -> str:
    """AccReflec の参照先 (SaveLoad:101-112)。DBflag > MainPath(D26 改: PLEASE WAIT.txt は見ない)。"""
    if dbflag != "":
        return paths.get("next")
    return paths.get("main")


def acc_reflec(screen: LotScreen, *,
               check_equipment: Callable[[str, str, float, float, float, float], str],
               history_rows: Optional[Sequence[Sequence[Any]]] = None,
               history_env: Optional["HistoryEnv"] = None) -> LotReflect:
    """SaveLoad.AccReflec (SaveLoad:56-345) の DB を読んで画面に出す値まで。

    - `check_equipment(設備名, 製造材質, 厚, 幅, 丈, 単重)` は QP_DB.CheckEquipmentDb(担当外)。"可" 以外なら
      `flow.Stop("設備仕様から外れています。,<結果>,終了します。")`(MsgError …, 2, "END"。SaveLoad:166-172)
    - `history_rows`(Lot情報履歴のシート)を渡すと、WorkTime="S" のとき HistorySave もここで呼ぶ
      (VBA の順: 公差の表示の後、指示の表示の前。SaveLoad:272-274)。渡さなければ `run_history_save` を見て呼び手が呼ぶ
    - この後の処理(VC 表示、作業選択、導電率、MsgError、MaterialsMsgBox、ゲージ判定、GetLotInfo)は呼び手。
      VC は `vc_action`、MaterialsMsgBox の受注番号は `materials_order_no`、注意文の重複除きは `dedupe_message`
    - 指示が無いときの対角・ラテラルボーの公差(Diagonal/ReDiagonal/Lateralbow_J/ReLateralbow)は担当外
    """
    out = LotReflect()
    if screen.lot == "":
        out.status = "no_lot"
        return out
    base = acc_reflec_path(screen.dbflag)
    out.db_path = file_path = db_path(base, "SIKALOTNOW")
    if not read_check(file_path, 1):
        out.status = "db_missing"
        out.messages.append(Message("DBファイルが見つかりません:" + "\r\n" + file_path, ERROR, source="SaveLoad:116"))
        return out
    # D26: 第1(MainPath)で読むときは、ロットが無ければ第2(控え)の SIKALOT も見る
    second = ""
    if base == paths.get("main") and base != paths.get("next"):
        second = db_path(paths.get("next"), "SIKALOTNOW")
    try:
        _acc_reflec_body(screen, out, file_path, check_equipment, history_rows, history_env, second)
    except flow.Stop:
        raise
    except (vbx.TypeMismatch, SubscriptOutOfRange, DbReadError, sqlite3.Error) as e:
        # ファイルが無い(DbMissing)は VBA の End に当たるのでここでは捕まえない
        out.status = "error"
        out.messages.append(db_read_error(e))
    return out


def _acc_reflec_body(screen: LotScreen, out: LotReflect, file_path: str,
                     check_equipment: Callable[[str, str, float, float, float, float], str],
                     history_rows: Optional[Sequence[Sequence[Any]]],
                     history_env: Optional["HistoryEnv"], second: str = "") -> None:
    if "'" in screen.cb_text:
        # VBA は SQL に値をそのまま埋める(SaveLoad:120)→ 構文エラーで ErrHandler
        raise DbReadError("クエリ式の構文エラー")
    cols, values = _select_lot(file_path, screen.cb_text)
    if values is None and second and read_check(second, 1):
        # D26: ロットごとの第2(列は名前で読むので並びが違っても同じ)
        cols, values = _select_lot(second, screen.cb_text)
        if values is not None:
            out.db_path = second
    if values is None:
        out.status = "not_found"
        for k in CLEAR_ON_NOT_FOUND:
            out.form[k] = ""
        return
    rs = _Rs(cols, values, out.null_uses)
    out.row = dict(rs.data)

    g_cours = rs.text("設計_設備ｺｰｽ", "SaveLoad:135")
    h_cours = rs.text("実績_設備ｺｰｽ", "SaveLoad:136")
    for key, msg in GAICHU:
        if vbx.instr(g_cours, key) > 0:
            out.temp_e = msg
            out.g_flag = True
            break

    zai = rs.text("製造材質", "SaveLoad:152")
    o_atu = vbx.single(vbx.val(rs.text("ｵｰﾀﾞｰ板厚", "SaveLoad:153")))
    o_hab = vbx.single(vbx.val(rs.text("ｵｰﾀﾞｰ板幅", "SaveLoad:154")))
    o_tak = vbx.single(vbx.val(rs.text("ｵｰﾀﾞｰ板丈", "SaveLoad:155")))
    weight = vbx.single(vbx.val(rs.text("製品単重", "SaveLoad:156")))
    if out.g_flag:
        o_atu = vbx.single(vbx.val(rs.text("BOX実績_板厚", "SaveLoad:159")))
        o_hab = vbx.single(vbx.val(rs.text("BOX実績_板幅", "SaveLoad:160")))
        o_tak = vbx.single(vbx.val(rs.text("BOX実績_板丈", "SaveLoad:161")))
    out.equipment_args = (screen.facility, zai, o_atu, o_hab, o_tak, weight)
    line_check = check_equipment(screen.facility, zai, o_atu, o_hab, o_tak, weight)
    if line_check != "可":
        raise flow.Stop("設備仕様から外れています。" + "," + line_check + "," + "終了します。",
                        source="SaveLoad:167")

    # PastHiki(Variant の配列。Null のまま入る)
    ph: list[Any] = [None] * 19
    ph[0] = "'" + rs.control("ｵｰﾀﾞｰ番号")
    for i, col in enumerate(PAST_HIKI_COLUMNS[1:], start=1):
        if out.g_flag and i in (2, 3, 4):
            col = ("BOX実績_板厚", "BOX実績_板幅", "BOX実績_板丈")[i - 2]
        ph[i] = rs.value(col)
    out.past_hiki = ph
    out.ken_no = "" if ph[5] is None else ph[5]

    f = out.form
    f["Zai"] = rs.control("製造材質")
    f["Tyu"] = rs.control("製造調質")
    f["Youto"] = rs.control("用途ｺｰﾄﾞ")
    f["TAIK"] = rs.control("品質ｸﾞﾚｰﾄﾞ_直角度")
    f["Lateral"] = rs.control("品質ｸﾞﾚｰﾄﾞ_ﾗﾃﾗﾙﾎﾞｰ")
    f["flatnes"] = rs.control("品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ")
    out.sika_set_atu = vbx.single(vbx.val(rs.text("ｵｰﾀﾞｰ板厚", "SaveLoad:208")))
    f["KenH"] = "前工程実績" + rs.control("BOX実績_枚本数") + "枚"
    if find_right_string(h_cours) != "":
        out.last_hot = "最終熱処理工程：" + find_right_string(h_cours)
    else:
        out.last_hot = ""                   # B54 を直す: 前ロットの表示を消す
    if vbx.instr(h_cours, "QF") > 0:
        out.qf = "QF有り"
    else:
        out.qf = ""                         # B54 を直す: 前ロットの表示を消す
    if not vbx.is_numeric(f["TAIK"]):
        f["TAIK"] = "4"
    if not vbx.is_numeric(f["Lateral"]):
        f["Lateral"] = ""
    if not vbx.is_numeric(f["flatnes"]):
        f["flatnes"] = ""

    if not out.g_flag:
        f["ATU"] = vbx.fmt(vbx.val(rs.text("製造板厚", "SaveLoad:218")), "0.000")
        f["Hab"] = vbx.fmt(vbx.val(rs.text("製造板幅", "SaveLoad:219")), "0.0")
        f["Tak"] = vbx.fmt(vbx.val(rs.text("製造板丈", "SaveLoad:220")), "0.0")
    else:
        f["ATU"] = vbx.fmt(vbx.val(rs.text("BOX実績_板厚", "SaveLoad:222")), "0.000")
        f["Hab"] = vbx.fmt(vbx.val(rs.text("BOX実績_板幅", "SaveLoad:223")), "0.0")
        f["Tak"] = vbx.fmt(vbx.val(rs.text("BOX実績_板丈", "SaveLoad:224")), "0.0")

    # 公差: オーダーの + か - のどちらかが 0 なら製造公差(SaveLoad:227-251)
    for key, p_key, m_key, line, pattern in (("板厚", "Ap", "Am", 227, "0.000"),
                                             ("板幅", "Hp", "Hm", 236, "0.0"),
                                             ("板丈", "Tp", "Tm", 245, "0.0")):
        raw = [rs.text(f"{key}公差_{k}", f"SaveLoad:{line + i}")
               for i, k in enumerate(("ｵｰﾀﾞｰ_ﾌﾟﾗｽ", "ｵｰﾀﾞｰ_ﾏｲﾅｽ", "製造_ﾌﾟﾗｽ", "製造_ﾏｲﾅｽ"))]
        # 2026-11-09: 全角の数字・記号(０．５・＋)も読む(VBA の Val は全角を 0 にする)
        op, om, mp, mm = (vbx.val(_half(x)) for x in raw)
        if op == 0 or om == 0:
            op, om = mp, mm
        if op == 0 and om == 0:
            # 2026-11-09: 公差が 0(下限 = 上限)になるときは、台帳の値をそのまま知らせる(原因を見つけるため)
            from . import lot_columns
            want = [f"{key}公差_{k}" for k in ("ｵｰﾀﾞｰ_ﾌﾟﾗｽ", "ｵｰﾀﾞｰ_ﾏｲﾅｽ", "製造_ﾌﾟﾗｽ", "製造_ﾏｲﾅｽ")]
            gone = [c for c in want if c in lot_columns.last_padded.get("SIKALOTNOW", [])]
            why = (f"読んだ台帳({out.db_path or file_path})に列 {'・'.join(gone)} がありません"
                   + (f"(台帳にあって読めない列: {'・'.join(lot_columns.last_unknown.get('SIKALOTNOW', [])[:8])})"
                      if lot_columns.last_unknown.get("SIKALOTNOW") else "")
                   if gone else f"読んだ台帳: {out.db_path or file_path}")
            out.messages.append(Message(
                f"{key}の公差が 0 になります(下限と上限が同じ)。仕掛台帳の値: "
                f"ｵｰﾀﾞｰ +{raw[0]!r} -{raw[1]!r} / 製造 +{raw[2]!r} -{raw[3]!r}。{why}", WARN, source="SaveLoad:227"))
        f[p_key] = vbx.fmt(op, pattern)
        f[m_key] = vbx.fmt(om, pattern)

    f["ApK"] = vbx.fmt(vbx.val(f["ATU"]) + vbx.val(f["Ap"]), "0.000")
    f["AmK"] = vbx.fmt(vbx.val(f["ATU"]) - vbx.val(f["Am"]), "0.000")
    f["HpK"] = vbx.fmt(vbx.val(f["Hab"]) + vbx.val(f["Hp"]), "0.0")
    f["HmK"] = vbx.fmt(vbx.val(f["Hab"]) - vbx.val(f["Hm"]), "0.0")
    f["TpK"] = vbx.fmt(vbx.val(f["Tak"]) + vbx.val(f["Tp"]), "0.0")
    f["TmK"] = vbx.fmt(vbx.val(f["Tak"]) - vbx.val(f["Tm"]), "0.0")
    if out.g_flag:
        f["HpK"] = f["HmK"] = f["TpK"] = f["TmK"] = ""

    out.run_history_save = screen.work_time == "S"
    if out.run_history_save and history_rows is not None:
        env = history_env or HistoryEnv(dbflag=screen.dbflag)
        try:
            out.history = history_save(ph, history_rows, env)
        except (vbx.TypeMismatch, SubscriptOutOfRange) as e:
            out.history = getattr(e, "partial", None)
            raise

    out.shiji_taik = rs.text("指示_直角度", "SaveLoad:276")
    out.shiji_late = rs.text("指示_ﾗﾃﾗﾙﾎﾞｰ", "SaveLoad:277")
    if out.shiji_taik != "":
        out.taik_text = vbx.fmt(convert_to_half(out.shiji_taik), "0.0")
        out.taik_flag = True
    if out.shiji_late != "":
        out.lateral_text = convert_to_half(out.shiji_late)
        out.late_flag = True
        if ph[5] is None:
            # `Dim kenNoTemp As String: kenNoTemp = PastHiki(5)`(SaveLoad:301)。Null の代入はエラー 94
            out.null_uses.append("SaveLoad:301 検査番号")
    if out.taik_flag and out.late_flag:
        out.additional = "対角ﾗﾃﾗﾙﾎﾞ-追記公差あり"
    elif out.taik_flag:
        out.additional = "対角追記公差あり"
    elif out.late_flag:
        out.additional = "ﾗﾃﾗﾙﾎﾞ-追記公差あり"

    out.cache_ord_atu = vbx.single(vbx.val(rs.text("ｵｰﾀﾞｰ板厚", "SaveLoad:344")))
    out.cache_c_course = rs.text("実績_設備ｺｰｽ", "SaveLoad:345")

    # VC (SaveLoad:378-381)。PastHiki(16) は Variant の文字列("00"/"01")。Null は比較が偽
    vc = ph[16]
    if vc is not None:
        if vbx.vcompare(vc, 0) == 0:
            out.vc_action = "clear"
        if vbx.vcompare(vc, 1) == 0:
            out.vc_action = "picture"


def materials_order_no(lot: str, past_hiki0: Any, dbflag: str = "", *,
                       now: Optional[_dt.datetime] = None, over12: bool = False
                       ) -> tuple[str, list[Message], list[LogLine]]:
    """AccReflec:431-449。MaterialsMsgBox に渡す受注番号。

    SIKAHIKINOW を Arrbox で読み(列名 → 引当)、引当の先頭行の `受注番号`。引当が無ければ
    PastHiki(0) から ' を除いたもの。列名が取れないのに引当があるときは VBA と同じく添字エラー 9。
    """
    msgs: list[Message] = []
    logs: list[LogLine] = []
    f = arrbox("SIKAHIKINOW", "板", "Fields", dbflag, now=now, over12=over12)
    a = arrbox("SIKAHIKINOW", lot, "Record", dbflag, now=now, over12=over12)
    for r in (f, a):
        msgs += r.messages
        logs += r.logs
    col = -1
    if f.value is not None:
        col = datasource.column_index(f.value, "受注番号")
    if a.value:
        if col < 0:
            raise SubscriptOutOfRange("受注番号 の列が無い (SaveLoad:446)")
        return a.value[0][col], msgs, logs
    return ("" if past_hiki0 is None else str(past_hiki0)).replace("'", ""), msgs, logs


# ------------------------------------------------------------------
# HistorySave(SaveLoad:1258-1552)
# ------------------------------------------------------------------
HISTORY_HEADER = ("登録日", "ｵｰﾀﾞｰNo", "ﾛｯﾄNo", "厚", "巾", "丈", "検査No", "材質", "調質", "用途ｺｰﾄﾞ", "ｺｰｽ",
                  "得意先", "納入先", "送り先", "前工程実績枚数", "前工程実績重量", "合紙", "VC", "VC_A", "VC_B",
                  "単重", "引当NO", "引当調整NO", "引当数量", "引当ﾌｧｲﾙ登録", "包装仕様NO")
#: 文字列の書式にする列(SaveLoad:1460 `Range("B:B,V:V").NumberFormatLocal = "@"`)
HISTORY_TEXT_COLUMNS = ("B", "V")


@dataclass
class HistoryEnv:
    """HistorySave が見る画面・シートの状態。"""
    dbflag: str = ""                 # UF1.DBflag
    now: Optional[_dt.datetime] = None
    over12: bool = False             # IsOver12Hours
    zenkasho_visible: bool = False   # シート「全ヵ所検査」が Visible
    cask_visible: bool = False       # シート「キャスク検査」が Visible
    uf1_test: bool = False           # UF1.Test
    mitsubishi_line: str = ""        # UFMitsubishi.LineName.Caption
    cask_line: str = ""              # UFCask.LineName.Caption


@dataclass
class HistorySaveResult:
    """HistorySave の結果。保存先(シートの代わり)への反映は呼び手。

    1. `delete_rows`(シートの行番号。1 始まり)を消す ── 再現では DB の確認より先に必ず行う
       (B50 を直した場合は最後まで進んだときだけ入る)
    2. 残った行の後ろに `new_rows`(A..Z の 26 列。B15 を直した場合は 27 列目にﾗﾃﾗﾙﾎﾞｰ等級)を足す
    3. `cache` があれば g_CacheLotNo / g_CacheCours / g_CacheKenNo を置き換える(None は前のまま)
    """
    delete_rows: list[int] = field(default_factory=list)
    new_rows: list[list[Any]] = field(default_factory=list)
    cache: Optional[dict[str, str]] = None
    messages: list[Message] = field(default_factory=list)
    logs: list[LogLine] = field(default_factory=list)
    file_path: str = ""              # 引当(SIKAHIKINOW)を読んだファイル
    completed: bool = False          # 最後まで進んだ(途中の MsgBox で抜けたら偽)


def history_delete_rows(sheet_rows: Sequence[Sequence[Any]], lot: Any) -> list[int]:
    """HistorySave:1268-1288。同じロットの行(シートの行番号)。

    B16: VBA は最初と最後に現れた行の間を**まとめて**消す(間の別ロットの行も消える)。
    直した場合はロットが一致する行だけ。シートが空(A1 だけ)なら VBA は型エラー 13。
    """
    if len(sheet_rows) <= 1 and (len(sheet_rows) == 0 or len(sheet_rows[0]) <= 1):
        raise vbx.TypeMismatch("Lot情報履歴が空(buf = A1 の 1 セル)SaveLoad:1270")

    def c(row: Sequence[Any]) -> Any:
        if len(row) < 3:
            raise SubscriptOutOfRange("Lot情報履歴の列が 3 未満 (SaveLoad:1272)")
        return row[2]

    hits = [i + 1 for i in range(1, len(sheet_rows)) if variant_eq(c(sheet_rows[i]), lot)]
    if not hits:
        return []
    return hits


def _history_path(env: HistoryEnv) -> str:
    """HistorySave の参照先 (SaveLoad:1299-1314)。"""
    if env.zenkasho_visible:
        return paths.get("main")
    if env.cask_visible:
        return paths.get("main")
    if env.dbflag != "":
        return paths.get("next")
    return paths.get("main")


def _need(idx: int, what: str) -> int:
    if idx < 0:
        raise SubscriptOutOfRange(f"{what} の列が無い (HistorySave)")
    return idx


def _odr_fallback(fpath: str, odr_path: str, what: str, res: HistorySaveResult,
                  over12: bool) -> tuple[str, str, bool]:
    """B51 を直した場合の SIKAODRNOW の控えへの切替(SaveLoad:1391-1401・1409-1423)。

    控え側を選んでいたら「…がありません」で抜ける。MainPath なら NextPath へ切り替え、ファイルが無ければ抜ける。
    戻り値は (FPath, ファイル, 続けるか)。
    """
    main, nxt = paths.get("main"), paths.get("next")
    target = vba_target("SIKAODRNOW")
    if fpath == nxt:
        if not over12:
            res.messages.append(Message(target + "に" + what + "_HistorySave", ERROR, source="SaveLoad:1393"))
        return fpath, odr_path, False
    if fpath == main:
        fpath = nxt
        odr_path = db_path(fpath, "SIKAODRNOW")
        if not read_check(odr_path, 1):
            if not over12:
                res.messages.append(Message(paths.missing_text("SIKAODRNOW"), ERROR,
                                            source="SaveLoad:1399"))
            return fpath, odr_path, False
    return fpath, odr_path, True


def history_save(past_hiki: Sequence[Any], sheet_rows: Sequence[Sequence[Any]], env: HistoryEnv) -> HistorySaveResult:
    """SaveLoad.HistorySave。`Lot情報履歴` に引当ごとの 1 行を作る(行の内容と消す行の判断まで)。

    `past_hiki` は PastHiki(0..18)(AccReflec・UFCask・UFMitsubishi が作る)。
    `sheet_rows` は `Lot情報履歴` の 1 行目(見出し)〜A 列の最終行。
    """
    res = HistorySaveResult()
    lot = past_hiki[1]
    delete_rows = history_delete_rows(sheet_rows, lot)
    try:
        _history_save_db(past_hiki, env, res)
    except (vbx.TypeMismatch, SubscriptOutOfRange) as e:
        # 再現: 行の削除は済んでいる(VBA は削除の後で止まる)。直した場合は消さない。呼び手が `e.partial` を反映する
        e.partial = res  # type: ignore[attr-defined]
        raise
    if res.completed:
        # B50 を直す: DB の確認を通って新しい行を作れたときだけ古い行を消す
        res.delete_rows = delete_rows
    return res


def _history_save_db(past_hiki: Sequence[Any], env: HistoryEnv, res: HistorySaveResult) -> HistorySaveResult:
    over12 = env.over12
    lot = past_hiki[1]
    lot_no = "" if lot is None else str(lot)

    main, nxt = paths.get("main"), paths.get("next")
    target = "SIKAHIKINOW"
    fpath = _history_path(env)
    file_path = db_path(fpath, target)
    c1 = read_check(file_path, 1) is False          # D26 改: 時刻窓・PLEASE WAIT は見ない
    missing = Message(paths.missing_text(target), ERROR,
                      source="SaveLoad:1322")
    if c1:
        if fpath == nxt:
            # 再現: 控え側を選んでいたら、控えのファイルがあっても抜ける(VBA どおり)
            # B49 を直す: 控えのファイルがあればそのまま控えを読む(無いときだけ抜ける)
            if not read_check(file_path, 1):
                if not over12:
                    res.messages.append(missing)
                res.file_path = file_path
                return res
        elif fpath == main:
            fpath = nxt
            file_path = db_path(fpath, target)
            if not read_check(file_path, 1):
                if not over12:
                    res.messages.append(missing)
                res.file_path = file_path
                return res

    flt = "[ﾛｯﾄ番号]='" + ado_filter.quote(lot_no) + "'"
    temp_hiki = datasource.get_records(file_path, TABLE, flt)
    if temp_hiki is None:
        if env.zenkasho_visible and env.mitsubishi_line == "Test":
            res.file_path = file_path
            return res
        if env.cask_visible and env.cask_line == "Test":
            res.file_path = file_path
            return res
        if not env.cask_visible and env.uf1_test:
            res.file_path = file_path
            return res
        if fpath == nxt:
            if not over12:
                res.messages.append(Message(vba_target(target) + "に" + lot_no + "ﾃﾞｰﾀがありません" + "_HistorySave",
                                            ERROR, source="SaveLoad:1341"))
            res.file_path = file_path
            return res
        if fpath == main:
            fpath = nxt
            file_path = db_path(fpath, target)
            if not read_check(file_path, 1):
                if not over12:
                    res.messages.append(missing)
                res.file_path = file_path
                return res
        temp_hiki = datasource.get_records(file_path, TABLE, flt)
        if temp_hiki is None:
            res.messages.append(Message("LOT一覧：LOTNO有_" + "HIKIDB：LOTNO無" + "_HistorySave", ERROR,
                                        source="SaveLoad:1355"))
            res.file_path = file_path
            return res
    res.file_path = file_path
    if fpath == main:
        temp_hiki, added = merge_hiki_from_second(temp_hiki, flt, file_path)     # 2026-11-08
        if added:
            res.messages.append(hiki_added_message(added, "HistorySave"))

    hiki_fields = datasource.get_fields(file_path, TABLE, "")
    if hiki_fields is None:
        if not over12:
            res.messages.append(Message("Hikifieldなし" + "_HistorySave", ERROR, source="SaveLoad:1364"))
        return res
    h = {n: datasource.column_index(hiki_fields, n)
         for n in ("受注番号", "引当番号", "ﾛｯﾄ番号", "検査番号", "引当数量", "引当調整NO", "引当ﾌｧｲﾙ登録年月日時分秒")}

    # SIKAODRNOW。`Dim VC(): VC = GetFieldsArr(…)` は Empty を代入すると型エラー 13
    # (SaveLoad:1390。そのため L1391-1406 の控えへの切替は到達しない)。B51 を直すと切替まで進む
    odr_path = db_path(fpath, "SIKAODRNOW")
    vc = datasource.get_fields(odr_path, TABLE, "")
    if vc is None:
        fpath, odr_path, ok = _odr_fallback(fpath, odr_path, "VCFieldﾃﾞｰﾀがありません", res, over12)
        if not ok:
            return res
        # VBA の L1402 は GetRecordsArr だが、列名を引く所なので GetFieldsArr(作者の意図)
        vc = datasource.get_fields(odr_path, TABLE, "")
        if vc is None:
            if not over12:
                res.messages.append(Message("VCFieldData無" + "_HistorySave", ERROR, source="SaveLoad:1404"))
            return res
    vcr = datasource.get_records(odr_path, TABLE, "")
    if vcr is None:
        fpath, odr_path, ok = _odr_fallback(fpath, odr_path, "VCRecordﾃﾞｰﾀがありません", res, over12)
        if not ok:
            return res
        vcr = datasource.get_records(odr_path, TABLE, "")
        if vcr is None:
            if not over12:
                res.messages.append(Message("VCRecordData無" + "_HistorySave", ERROR, source="SaveLoad:1426"))
            return res
    r = {n: datasource.column_index(vc, n)
         for n in ("受注番号", "VC_表", "VC_裏", "製品単重", "合紙", "取引先名称", "納入先名称", "送り先名称", "包装仕様NO")}
    if fpath == main and main != nxt and r["受注番号"] >= 0 and h["受注番号"] >= 0:
        # D26(ロットごとの第2): 引当の受注が第1の SIKAODR に無ければ、第2の SIKAODR の行を後ろに足す
        have = {v[r["受注番号"]] for v in vcr}
        want = {x[h["受注番号"]] for x in temp_hiki} - have
        if want:
            extra, _second = lot_from_second("SIKAODRNOW", "", odr_path)
            vcr = list(vcr) + [v for v in (extra or []) if v[r["受注番号"]] in want]

    now = _now(env.now)
    base = list(past_hiki) + [None] * max(0, 19 - len(past_hiki))
    for vrow in vcr:
        for hrow in temp_hiki:
            if hrow[_need(h["受注番号"], "受注番号")] != vrow[_need(r["受注番号"], "受注番号")]:
                continue
            row: list[Any] = [None] * 26
            row[0] = now
            for k in range(19):                     # B..T = PastHiki(0..18)
                row[1 + k] = base[k]
            lateral_grade = base[17]
            row[18] = vrow[_need(r["VC_表"], "VC_表")]           # S(B15: ﾗﾃﾗﾙﾎﾞｰ等級を上書き)
            row[19] = vrow[_need(r["VC_裏"], "VC_裏")]           # T
            row[20] = vrow[_need(r["製品単重"], "製品単重")]      # U
            row[21] = hrow[_need(h["引当番号"], "引当番号")]      # V
            row[1] = hrow[_need(h["受注番号"], "受注番号")]       # B
            row[16] = vrow[_need(r["合紙"], "合紙")]              # Q
            row[11] = vrow[_need(r["取引先名称"], "取引先名称")]  # L
            row[12] = vrow[_need(r["納入先名称"], "納入先名称")]  # M
            row[13] = vrow[_need(r["送り先名称"], "送り先名称")]  # N
            row[22] = hrow[_need(h["引当調整NO"], "引当調整NO")]  # W
            row[23] = hrow[_need(h["引当数量"], "引当数量")]      # X
            row[24] = hrow[_need(h["引当ﾌｧｲﾙ登録年月日時分秒"], "引当ﾌｧｲﾙ登録年月日時分秒")]  # Y
            row[25] = vrow[_need(r["包装仕様NO"], "包装仕様NO")]  # Z
            row.append(lateral_grade)           # AA: ﾗﾃﾗﾙﾎﾞｰ等級(直した場合の置き場所)
            res.new_rows.append(row)

            ken = ""
            for hr in temp_hiki:
                if variant_eq(hr[_need(h["ﾛｯﾄ番号"], "ﾛｯﾄ番号")], lot):
                    ken = vbx.left(hr[_need(h["検査番号"], "検査番号")], 1)
                    break
            res.cache = {"lot_no": lot_no, "cours": "" if base[9] is None else str(base[9]), "ken_no": ken}
    res.completed = True
    return res


# ------------------------------------------------------------------
# VC_Picture(SaveLoad:1554-1758)・VcConv(Input:753-800)
# ------------------------------------------------------------------
VC_NOTHING_SUFFIX = "写真登録がないかﾈｯﾄﾜｰｸに繋がっていません：\\\\nlmsrvngy03\\各課共有\\0130_日軽稲沢\\梱包課\\AIM\\VC種類画像\\"


@dataclass
class VcView:
    """VC_Picture が決める画面(VC_L_A/VC_L_B ラベル、VC_A/VC_B 画像、VcNothing)。"""
    matched: bool = False                 # Lot情報履歴にロットがあった(偽なら全部隠したまま)
    l_a_visible: bool = False
    l_b_visible: bool = False
    a_visible: bool = False
    b_visible: bool = False
    l_a_caption: str = "A面"
    l_b_caption: str = "B面"
    a_picture: Optional[str] = None       # 読み込む画像のパス
    b_picture: Optional[str] = None
    l_a_tag_on: bool = False              # VC_L_A.Tag = "on"
    l_b_tag_on: bool = False
    nothing_visible: bool = False
    nothing_text: str = ""
    tall_form: bool = False               # UF1.Height = 703.5
    b_left: str = ""                      # "" = 既定(18) / "A幅+30" / "B幅+30"
    name_a: str = ""
    name_b: str = ""
    messages: list[Message] = field(default_factory=list)


def _denso_base(test: bool) -> str:
    return paths.get("packing") if test else paths.get("quality")


def vc_picture(sheet_rows: Sequence[Sequence[Any]], lot: Any, order_no: str = "", *, test: bool = False,
               materials_upper_part: Optional[Callable[[str, bool], str]] = None) -> VcView:
    """SaveLoad.VC_Picture。`Lot情報履歴` の VC_A/VC_B(S/T 列)から VC 画像と VC厚 の系統・備考を出す。

    - `sheet_rows` は 1 行目(見出し)〜A 列の最終行。VBA は 2 行目からを配列にする
    - `order_no` を渡すとその受注番号の行の VC を見て、`materials_upper_part(受注番号, 当板?)`
      (MaterialsUpperPart。担当外)の結果を MsgBox で出す
    - 画像は `PackingPath & "VC種類画像\\" & 品名ｺｰﾄﾞ & ".JPG"`
    """
    v = VcView()
    # 使用パス「VC種類画像」(2026-11-05)。決めていなければ VBA と同じ PackingPath & "VC種類画像"
    vc_path = (subpath(_denso_base_for_picture(), "VC種類画像") if paths._saved_value("vc_images") is None
               else paths.get("vc_images").rstrip("\\/"))
    if len(sheet_rows) >= 2:
        body = list(sheet_rows[1:])
    else:
        body = list(sheet_rows[:2])          # Range(A2:Z1) は A1:Z2 になる
    x = -1
    for i, row in enumerate(body):
        if variant_eq(_cell(row, 3), lot):
            x = i
            break
    if x < 0:
        return v
    v.matched = True

    a_fl = b_fl = False
    vca = vcb = ""
    if order_no == "":
        if vbx.vcompare(_cell(body[x], 18), 1) == 0:
            if _cell_ne_empty(_cell(body[x], 19)):
                vca = _cstr(_cell(body[x], 19)) + ".JPG"
                a_fl = True
            if _cell_ne_empty(_cell(body[x], 20)):
                vcb = _cstr(_cell(body[x], 20)) + ".JPG"
                b_fl = True
    else:
        for row in body:
            if variant_eq(_cell(row, 2), order_no):
                if _cell_ne_empty(_cell(row, 19)):
                    vca = _cstr(_cell(row, 19)) + ".JPG"
                    a_fl = True
                if _cell_ne_empty(_cell(row, 20)):
                    vcb = _cstr(_cell(row, 20)) + ".JPG"
                    b_fl = True
        if materials_upper_part is not None:
            r1 = materials_upper_part(order_no, False)
            r2 = materials_upper_part(order_no, True)
            if r1 != "" and r1 != "一致なし":
                v.messages.append(Message(r1, INFO, source="SaveLoad:1621"))
            if r1 == "アングル":
                if r2 != "" and r2 != "一致なし":
                    v.messages.append(Message(r2, INFO, source="SaveLoad:1624"))

    # VC厚 の系統・備考(ファイルが無いと GetRecordsArr の中で End → DbMissing)
    file_path = paths.join(_denso_base(test), DENSO + ".sqlite3")
    flt_a = "[品名ｺｰﾄﾞ]='" + ado_filter.quote(_cstr(_cell(body[x], 19))) + "'"
    flt_b = "[品名ｺｰﾄﾞ]='" + ado_filter.quote(_cstr(_cell(body[x], 20))) + "'"
    temp_a = datasource.get_records(file_path, VC_TABLE, flt_a)
    temp_b = datasource.get_records(file_path, VC_TABLE, flt_b)
    fields = datasource.get_fields(file_path, VC_TABLE, "")
    if fields is None:
        raise vbx.TypeMismatch("VC厚 の列名が Empty (LBound(Empty) SaveLoad:1659)")
    c_kei = datasource.column_index(fields, "系統")
    c_biko = datasource.column_index(fields, "備考")
    if temp_a is not None:
        v.name_a = temp_a[0][_need(c_kei, "系統")] + "_" + temp_a[0][_need(c_biko, "備考")]
    if temp_b is not None:
        v.name_b = temp_b[0][_need(c_kei, "系統")] + "_" + temp_b[0][_need(c_biko, "備考")]

    if a_fl or b_fl:
        v.tall_form = True
    pic_a = subpath(vc_path, vca)
    pic_b = subpath(vc_path, vcb)

    def hide_all(name: str) -> None:
        v.nothing_visible = True
        v.nothing_text = name + VC_NOTHING_SUFFIX
        v.l_a_visible = v.l_b_visible = v.a_visible = v.b_visible = False

    if a_fl and b_fl:
        v.l_a_tag_on = v.l_b_tag_on = True
        v.l_a_visible = v.l_b_visible = v.a_visible = v.b_visible = True
        v.l_a_caption = "A面" + ":" + v.name_a
        v.l_b_caption = "B面" + ":" + v.name_b
        if read_check(pic_a, 1):
            v.a_picture = pic_a
        else:
            hide_all(vca)
        v.b_left = "A幅+30"
        if read_check(pic_b, 1):
            v.b_picture = pic_b
        else:
            hide_all(vcb)
    elif a_fl and not b_fl:
        v.l_a_tag_on = True
        v.l_a_visible = True
        v.l_a_caption = "A面" + ":" + v.name_a
        v.a_visible = True
        if read_check(pic_a, 1):
            v.a_picture = pic_a
            v.l_b_visible = v.b_visible = False
        else:
            hide_all(vca)
    elif not a_fl and b_fl:
        v.l_b_tag_on = True
        v.l_b_visible = True
        v.l_b_caption = "B面" + ":" + v.name_b
        v.b_visible = True
        if read_check(pic_b, 1):
            v.b_picture = pic_b
            v.l_a_visible = v.a_visible = False
        else:
            hide_all(vcb)
        v.b_left = "B幅+30"
    else:
        v.nothing_visible = True
        v.nothing_text = _cstr(_cell(body[x], 2)) + ":受注番号一致なし(進度ﾌｧｲﾙ異常)"
    return v


def _denso_base_for_picture() -> str:
    # VC_Picture:1578-1581。Test でもそれ以外でも PackingPath
    return paths.get("packing")


def _cell(row: Sequence[Any], col: int) -> Any:
    """1 始まりの列番号でセルを引く(無ければ Empty)。"""
    return row[col - 1] if 0 < col <= len(row) else None


def _cstr(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return vbx.cstr(v)
    return str(v)


@dataclass
class VcConvResult:
    value: float = 0.0                    # Single
    messages: list[Message] = field(default_factory=list)


VC_AVERAGE = 0.08406977   # 登録 VC 全平均厚 (Input:796)


def vc_conv(no: str, *, test: bool = False) -> VcConvResult:
    """Input.VcConv (Input:753-800)。VC厚 の `品名ｺｰﾄﾞ` → `厚-係数`(Single)。

    - 伝送用ファイルが無い → MsgBox「日稲:伝送用ファイル.accdbがありません_ReadCheck_VcConv」、0
    - 列名が取れない → MsgBox「fieldsなし_VcConv」、0
    - 登録なし → MsgBox「VC厚登録なし,担当者に連絡のこと」、0.08406977(全平均)
    - 係数が数にならない → 型エラー 13(Single への代入)
    """
    res = VcConvResult()
    file_path = paths.join(_denso_base(test), DENSO + ".sqlite3")
    if not read_check(file_path, 1):
        res.messages.append(Message(paths.missing_text(DENSO, [file_path]), ERROR,
                                    source="Input:772"))
        return res
    temp = datasource.get_records(file_path, VC_TABLE, "[品名ｺｰﾄﾞ]='" + ado_filter.quote(no) + "'")
    fields = datasource.get_fields(file_path, VC_TABLE, "")
    if fields is None:
        res.messages.append(Message("fieldsなし" + "_VcConv", INFO, source="Input:785"))
        return res
    col = datasource.column_index(fields, "厚-係数")
    if temp is None:
        res.messages.append(Message("VC厚登録なし,担当者に連絡のこと", INFO, source="Input:795"))
        res.value = vbx.single(VC_AVERAGE)
        return res
    res.value = vbx.single(vbx.cdbl(temp[0][_need(col, "厚-係数")]))
    return res

