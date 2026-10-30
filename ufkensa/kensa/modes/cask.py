"""キャスク検査(UFCask + キャスク検査シート + Cask モジュールのキャスク専用部分)

- シートのイベント: `KySh1`(= `KySh`。原紙側も同じ内容)の Worksheet_Change / SelectionChange
- フォーム(UFCask)の操作: ロット読込・ゼロ校正・転送(CB3)・印刷(CB2)・保存(CB4)・取消(CB5)・
  巾丈測定の切替(CB9)・新規発行(CB1)
- 梱包板番号(伝送用ファイル `キャスク梱包板番号`): 連番チェック・追加/更新・スタートNo=1 の削除。
  D4 で SQLite。書き込みは `BEGIN IMMEDIATE` と `busy_timeout`(VBA と同じく STEP1 と STEP2 は別々に確定する)

ボタンの操作は `Book.fx`(Effects)に結果を積んで返す。MsgBox で聞く所は `flow.Question` を返し、
`answers` に答えを入れて呼び直す(聞くまでの処理はもう一度走るが、どれも同じ結果になる処理)。
"""
from __future__ import annotations

import sqlite3
import unicodedata
from datetime import datetime
from typing import Any, Optional, Union

from .. import datasource, flow, paths, quirks, vbx
from . import sheet as sh
from .sheet import CASK, Book, Call, Effects, LogCall, PrintJob, SavePlan, SaveStep, VbaError

#: SIKALOTNOW `仕掛` の欄名(UFCask:436-452 の書き方のまま。4 つが全角カナ = U4)
CASK_FIELDS = ("オーダー番号", "ロット番号", "製造板厚", "製造板幅", "製造板丈", "検査番号", "製造材質", "製造調質",
               "用途コード", "設計_設備コース", "取引先名", "納入先名", "送り先名", "BOX実績_枚本数",
               "BOX実績_良品重量", "梱包_合紙", "VC")
FIELD_MISSING = "要求された名前、または序数に対応する項目がコレクションに見つかりません。"
INVALID_NULL = "Null の使い方が不正です。"
TABLE = "キャスク梱包板番号"
_DATE_COLUMNS = {"作業時間"}


def _cask_log_path(line_name: str) -> str:
    """新規発行のログ先(UFCask:53-61)。判定付けと同じ振り分け。"""
    return sh.log_dir(line_name)


# ------------------------------------------------------------------
# シートのイベント(KySh1)
# ------------------------------------------------------------------
_RAN = ((2, 2, 3, 6), (2, 11, 4, 15), (2, 17, 4, 19), (2, 21, 4, 26), (7, 2, 706, 2), (7, 4, 706, 4),
        (7, 6, 706, 6), (7, 8, 706, 8), (7, 11, 706, 11), (7, 13, 706, 13), (7, 15, 706, 15), (7, 17, 706, 17),
        (7, 19, 706, 20), (7, 22, 706, 22), (7, 24, 706, 28))
_HANTEI_COLS = (3, 5, 7, 9, 12, 14, 16, 18, 21)


def _in(rects: tuple, r: int, c: int) -> bool:
    return any(r1 <= r <= r2 and c1 <= c <= c2 for r1, c1, r2, c2 in rects)


def worksheet_change(book: Book, target: sh.Sheet, r: int, c: int) -> None:
    """KySh1.Worksheet_Change (KySh:9-91)。On Error Resume Next の下で動く。"""
    ws = book.active
    befor = book.resume(lambda: sh.take_befor(book, book.cask, r, c), "KySh:19") or ""
    TR, TC = r, c
    if ws.name == CASK and 7 <= TR <= 706:
        cask = book.cask
        if 2 <= TC <= 8:                      # Case 7 は Case 2 To 8 に先に当たる(到達しない)
            book.resume(lambda: book.activate(TR, TC + 1))
        elif TC == 9:
            if cask.count_st == "4":
                if cask.inspection == "OFF":
                    book.resume(lambda: book.activate(TR + 1, 2))
                elif cask.inspection == "ON":
                    book.resume(lambda: book.activate(TR, 11))
        elif 11 <= TC <= 18:
            book.resume(lambda: book.activate(TR, TC + 1))
        elif TC == 19:
            if not sh.blank(ws.get(TR, 19)):
                book.fx.sounds.append("音楽対角")
                book.resume(lambda: sh.uf_taikaku(book, TR, TC))
                book.events_enabled = False
                for cc in range(20, 24):
                    book.write(ws, TR, cc, "")
                book.events_enabled = True
                book.resume(lambda: book.activate(TR, TC + 1))
        elif TC == 20:
            book.resume(lambda: book.activate(TR, TC + 1))
        elif TC == 21:
            book.resume(lambda: book.activate(TR, 24))
        elif 24 <= TC <= 27:
            book.resume(lambda: book.activate(TR, TC + 1))
    if ws.name == CASK:
        book.events_enabled = False
        book.resume(lambda: sh.al_ter_range(book, TR, TC, befor), "AlTerRange")
        if not sh.is_numeric(target.get(TR, TC)):
            return                                # EnableEvents は False のまま(VBA どおり)
        book.events_enabled = True
        book.resume(lambda: sh.kensa_hantei(book, TR, TC), "検査判定")
        book.events_enabled = False
        book.resume(lambda: sh.taikaku(book, TR, TC), "対角")
        book.events_enabled = True
        book.resume(lambda: sh.koshagai_kensaku(book, TR, TC), "公差外検索")


def write_header(book: Book) -> None:
    """KySh:105-126 列 24(X)を選んだときにフォームの値をシートの Y2:AG3 に書く。"""
    ws = book.active
    f = book.cask
    book.events_enabled = False
    book.write(ws, 2, 25, f.ken_no)
    book.write(ws, 3, 25, f.kname)
    book.write(ws, 2, 26, f.cno)
    book.write(ws, 3, 26, f.mno)
    book.write(ws, 4, 26, f.seri_no)
    soku = ""
    if sh.val_cell(ws.get(3, 2)) >= 100.0:
        soku = "ﾉｷﾞｽ"
    if sh.val_cell(ws.get(3, 2)) < 100.0:
        soku = "ﾏｲｸﾛ"
    book.write(ws, 2, 28, "ｾﾞﾛ校正:" + soku)
    book.write(ws, 3, 28, "ｾﾞﾛ校正:" + "ｺﾝﾍﾞｯｸｽ")
    book.write(ws, 3, 30, "1000mm校正:" + "ｺﾝﾍﾞｯｸｽ")
    book.write(ws, 2, 29, f.micro)
    book.write(ws, 3, 29, f.convex)
    book.write(ws, 3, 31, f.convex_1000)
    book.write(ws, 3, 33, f.count_st)
    for r in range(2, 5):
        for c in range(25, 32):
            ws.font_size[(r, c)] = 8
    book.events_enabled = True


def worksheet_selection_change(book: Book, r: int, c: int) -> None:
    """KySh1.Worksheet_SelectionChange (KySh:93-201)。"""
    book.cask.cnt = "0"
    ws = book.active
    if ws.name == CASK:
        if c == 21:
            book.activate(r, 24)
        elif c == 24:
            write_header(book)
    if ws.name == CASK and (r, c) == (7, 4):
        book.events_enabled = False
        if not sh.blank(ws.get(7, 2)):
            book.write(ws, 1, 2, sh.fmt_date(book.now(), "yyyy/mm/dd h:mm"))
            ws.font_size[(1, 2)] = 5
            book.fx.calls.append(Call("Daily_Report", (1,)))
            book.fx.calls.append(Call("Daily_Report", (3,)))
        book.events_enabled = True
    if ws.name != CASK:
        return
    off = book.cask.inspection == "OFF"
    if 7 <= r <= 706 and c in _HANTEI_COLS:
        for rect in _RAN:
            ws.fill(*rect, sh.WHITE)
        if off:
            ws.fill(6, 11, 706, 28, sh.GRAY)
        return
    if _in(_RAN, r, c):
        for rect in _RAN:
            ws.fill(*rect, sh.WHITE)
        ws.interior[(r, c)] = sh.LIGHT_BLUE
        ws.font_color[(r, c)] = (0, 0, 0)
        if off:
            ws.fill(6, 11, 706, 28, sh.GRAY)


sh.HANDLERS[CASK] = (worksheet_change, worksheet_selection_change)


# ------------------------------------------------------------------
# Cask モジュール(キャスク専用)
# ------------------------------------------------------------------
def cask_copy(book: Book) -> None:
    """Cask.キャスクコピー (Cask:310-334)。キャスク検査を消して原紙から作り直す。"""
    ws = sh.template(sh.CASK_TEMPLATE)
    ws.name = CASK
    ws.visible = True
    ws.tab_color = 0x660066
    book.sheets[CASK] = ws
    book.activate_sheet(CASK)


def loop_judge(book: Book, *, u6_raises: bool = False) -> None:
    """Cask.ループ判定 (Cask:979-1027)。キャスク検査の全行の厚・巾・丈を判定し直す(色は RGB(0,128,0)/(200,60,50))。

    `KenJudg((Cells(i, a)), …)` は括弧で値を渡している(U6)。実機でエラーになるなら `u6_raises=True`
    (実行時エラー 424。呼び出し元に処理は無いので止まる)。値は**アクティブシート**から読む(VBA どおり)。
    """
    book.events_enabled = False
    ws = book.ws(CASK)
    act = book.active
    atu1, haba1, take1 = sh.val_cell(ws.get(2, 17)), sh.val_cell(ws.get(3, 17)), sh.val_cell(ws.get(4, 17))
    atu2, haba2, take2 = sh.val_cell(ws.get(2, 19)), sh.val_cell(ws.get(3, 19)), sh.val_cell(ws.get(4, 19))
    for i in range(7, 707):
        for a in range(2, 18):
            if sh.blank(ws.get(i, a)):
                continue
            if a in (2, 4, 6, 8):
                lim = (atu2, atu1)
            elif a in (11, 13):
                lim = (haba2, haba1)
            elif a in (15, 17):
                lim = (take2, take1)
            else:
                continue
            if u6_raises:
                raise VbaError(424, "オブジェクトが必要です。", "Cask:996")
            ok = sh.ken_judg(act.get(i, a), *lim)
            book.write(ws, i, a + 1, "○" if ok else "×")
            ws.interior[(i, a + 1)] = sh.OK_GREEN if ok else sh.NG_RED
    book.events_enabled = True


def count_select(book: Book, a: int) -> None:
    """Cask.CountSelect (Cask:1963-1993)。厚の測定点数(3 = H:I 列を狭めて保護、4 = 広げる)。"""
    if not book.has(CASK):
        return
    ws = book.ws(CASK)
    book.activate_sheet(CASK)
    _lock_hi(ws, True)
    book.cask.count_st = "3"
    if a == 3:
        _lock_hi(ws, True)
        book.cask.count_st = "3"
    elif a == 4:
        _lock_hi(ws, False)
        book.cask.count_st = "4"


def _lock_hi(ws: sh.Sheet, lock: bool) -> None:
    for r in range(7, 707):
        for c in (8, 9):
            (ws.locked.add if lock else ws.locked.discard)((r, c))
    if lock:
        ws.col_width[8] = ws.col_width[9] = 0.05
    else:
        ws.col_width[8], ws.col_width[9] = 8.88, 2.63


def _lock_inspection(ws: sh.Sheet, lock: bool) -> None:
    for r in range(7, 707):
        for c in range(11, 29):
            (ws.locked.add if lock else ws.locked.discard)((r, c))


def numeric_confi(book: Book, n: str) -> bool:
    """Cask.NumericConfi (Cask:1956-1962)。ゼロ点は 0 ちょうどだけ可、コンベックスは ±0.5 以内。"""
    if n == "M":
        return not (vbx.val(book.cask.micro) != 0)
    if n == "C":
        v = vbx.val(book.cask.convex)
        return not (v > 0.5 or v < -0.5)
    return True


def start_tb_after_update(book: Book, text: str) -> Effects:
    """UFCask.StartTB_AfterUpdate (UFCask:736-776)。ゼロ校正の入力(DT = マイクロ/ノギス、#L = コンベックス)。"""
    fx = book.begin()
    f = book.cask
    if text == "#L":
        f.l2_visible = True
        return fx
    temp = sh.maker(text)
    atu = sh.cell_str(book.ws(CASK).get(3, 2))
    if temp == "DT":
        pr = sh.conv(text)
        if atu == "":
            atu = "0"
        try:
            big = vbx.vcompare(atu, 100) >= 0
        except vbx.TypeMismatch:
            raise flow.Stop("実行時エラー '13':" + "\r\n\r\n" + "型が一致しません。", source="UFCask:750")
        if big:
            f.l1 = "ﾉｷﾞｽｾﾞﾛ点"
            f.micro = vbx.fmt(vbx.val(pr), "0.00")
            if not numeric_confi(book, "M"):
                fx.msg("ノギス校正を正しく実行してください！！", flow.WARN, "UFCask:754")
                f.micro = ""
        else:
            f.l1 = "ﾏｲｸﾛｾﾞﾛ点"
            f.micro = vbx.fmt(vbx.val(pr), "0.000")
            if not numeric_confi(book, "M"):
                fx.msg("マイクロ校正を正しく実行してください！！", flow.WARN, "UFCask:761")
                f.micro = ""
    elif temp == "#L":
        pr = sh.conv(text)
        f.convex = vbx.fmt(vbx.val(pr), "0.0")
        if not numeric_confi(book, "C"):
            fx.msg("コンベックス校正を正しく実行してください！！", flow.WARN, "UFCask:769")
            f.convex = ""
    return fx


def proofreading_search(book: Book, to_time: str, to_tyo: str) -> None:
    """Cask.ProofreadingSearch (Cask:1939-1955)。`品証検証` A5:K98 の K 列先頭 10 文字=日付・B 列=直 の C 列を 1000mm 校正値に。

    `to_time` / `to_tyo` は Worker.YeaDayrCheck / TimeCheck の結果(別担当)。
    """
    if not book.has("品証検証") or "品証検証" not in book.extra:
        return
    q = book.extra["品証検証"]
    for i in range(5, 99):
        if vbx.left(sh.cell_str(q.get(i, 11)), 10) == to_time:
            if vbx.veq(q.get(i, 2) if q.get(i, 2) is not None else "", to_tyo):
                book.cask.convex_1000 = sh.cell_str(q.get(i, 3))
                if book.cask.kname == "":
                    book.cask.kname = sh.cell_str(q.get(i, 10))
                book.cask.l3_visible = True


def max_min(book: Book) -> None:
    """Cask.キャスク用最大最小 (Cask:336-443)。MIN〜MAX 行の厚・巾・丈の最大最小を V3:X4 に書き、一致セルに下線。

    比較は Single(`Cells = MAXatu`)。B28: 厚の MIN は行と列の変数が入れ替わっている(行 x〜k・列 2〜n)。
    """
    f = book.cask
    ws = book.ws(CASK)
    n = vbx.clng(vbx.val(f.min_text) + 6)
    k = vbx.clng(vbx.val(f.max_text) + 6)

    def col_vals(col: int) -> list[float]:
        lo, hi = min(n, k), max(n, k)
        out = []
        for r in range(lo, hi + 1):
            v = ws.get(r, col)
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                out.append(float(v))
        return out

    def wmax(col: int) -> float:
        v = col_vals(col)
        return max(v) if v else 0.0

    def wmin(col: int) -> float:
        v = col_vals(col)
        return min(v) if v else 0.0

    max_atu = min_atu = 0.0
    if f.count_st == "4":
        max_atu = vbx.single(max(wmax(2), wmax(4), wmax(6), wmax(8)))
        min_atu = vbx.single(min(wmin(2), wmin(4), wmin(6), wmin(8)))
    elif f.count_st == "3":
        max_atu = vbx.single(max(wmax(2), wmax(4), wmax(6)))
        min_atu = vbx.single(min(wmin(2), wmin(4), wmin(6)))
    max_haba = vbx.single(max(wmax(11), wmax(13)))
    min_haba = vbx.single(min(wmin(11), wmin(13)))
    max_take = vbx.single(max(wmax(15), wmax(17)))
    min_take = vbx.single(min(wmin(15), wmin(17)))
    for r, c in ((3, 22), (4, 22), (3, 23), (4, 23), (3, 24), (4, 24)):
        book.write(ws, r, c, "")
    ws.clear_underline(7, 2, 706, 28)

    def cell(i: int, a: int) -> Any:
        if i < 1 or a < 1:
            raise VbaError(1004, "アプリケーション定義またはオブジェクト定義のエラーです。", "Cask:397")
        return ws.get(i, a)

    def mark(rows: range, cols: range, target: float, out: tuple[int, int]) -> None:
        for i in rows:
            for a in cols:
                if sh.eq_single(cell(i, a), target):
                    ws.underline.add((i, a))
                    book.write(ws, out[0], out[1], target)
                    break

    x = 0
    for i in range(n, k + 1):
        x = {"4": 8, "3": 6}.get(f.count_st, x)
        for a in range(2, x + 1, 2):
            if sh.eq_single(cell(i, a), max_atu):
                ws.underline.add((i, a))
                book.write(ws, 3, 22, max_atu)
                break
    if quirks.reproduce("B28"):
        mark(range(x, k + 1), range(2, n + 1, 2), min_atu, (4, 22))
    else:
        mark(range(n, k + 1), range(2, x + 1, 2), min_atu, (4, 22))
    mark(range(n, k + 1), range(11, 14, 2), max_haba, (3, 23))
    mark(range(n, k + 1), range(11, 14, 2), min_haba, (4, 23))
    mark(range(n, k + 1), range(15, 18, 2), max_take, (3, 24))
    mark(range(n, k + 1), range(15, 18, 2), min_take, (4, 24))


# ------------------------------------------------------------------
# 梱包板番号(伝送用ファイル)
# ------------------------------------------------------------------
def _db_path(name: str):
    return datasource.resolve(name)


def _open_rw(name: str) -> sqlite3.Connection:
    p = _db_path(name)
    if not p.is_file():
        raise VbaError(-2147467259, f"ファイル '{p}' が見つかりません。", "db.Open")
    conn = sqlite3.connect(str(p), timeout=datasource.BUSY_TIMEOUT_MS / 1000, isolation_level=None)
    conn.execute(f"PRAGMA busy_timeout = {datasource.BUSY_TIMEOUT_MS}")
    return conn


def _variant(col: str, v: Any) -> Any:
    """SQLite の値 → ADO の Variant(日付の列は日時に)。"""
    if v is not None and col in _DATE_COLUMNS and isinstance(v, str):
        d = sh.to_cell(v)
        return d if isinstance(d, datetime) else v
    return v


def _vstr(v: Any) -> str:
    """Variant を文字列に連結するとき(`"…" & v`)。"""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return sh.cstr_date(v)
    return vbx.cstr(v)


def get_cask_data(book: Book, lot: str, ken: str) -> Optional[list[list[Any]]]:
    """Cask.GetCaskData + GetRecordsSQL (Cask:2251-2349)。`SELECT * … ORDER BY スタートNo` を 2 次元配列で。0 件・エラーは None。

    配列の列は **表の列の並び**(VBA は 3 = スタートNo、4 = エンドNo と決めて読む。受領 DB では
    LotNo, 検査No, スタートNo, エンドNo, 作業時間, 管理番号 の順なので 3 = エンドNo、4 = 作業時間 になる)。
    """
    try:
        with datasource.open_ro(_db_path("伝送用ファイル")) as conn:
            cur = conn.execute(f'SELECT * FROM "{TABLE}" WHERE "LotNo" = ? COLLATE NOCASE '
                               f'AND "検査No" = ? COLLATE NOCASE ORDER BY "スタートNo"', (lot, ken))
            cols = [d[0] for d in cur.description]
            rows = [[_variant(cols[i], v) for i, v in enumerate(r)] for r in cur.fetchall()]
    except (datasource.DbMissing, sqlite3.Error) as e:
        book.fx.msg("GetRecordsSQL エラー：" + "\r\n" + str(e), flow.ERROR, "Cask:2346")
        return None
    return rows or None


def add_cask_labels(book: Book, lot: str, ken: str) -> None:
    """Cask.AddCaskLabels (Cask:2367-2457)。フォームの梱包範囲ラベル `No{n}:{[3]}～{[4]}`。"""
    book.cask.labels = []
    data = get_cask_data(book, lot, ken)
    if data is None:
        return
    book.cask.labels = ["No" + str(i + 1) + ":" + _vstr(_at(row, 3)) + "～" + _vstr(_at(row, 4))
                        for i, row in enumerate(data)]


def _at(row: list, i: int) -> Any:
    if i >= len(row):
        raise VbaError(9, "インデックスが有効範囲にありません。")
    return row[i]


def _clng_variant(v: Any) -> int:
    """`CLng(Variant)`。日付はシリアル値を銀行丸め。"""
    if isinstance(v, datetime):
        serial = (v - datetime(1899, 12, 30)).total_seconds() / 86400
        return vbx.clng(serial)
    if v is None:
        raise VbaError(94, INVALID_NULL)
    try:
        return vbx.clng(v)
    except vbx.TypeMismatch as e:
        raise sh.type_mismatch(str(e)) from e


def check_cask_end_no(book: Book, lot: str, ken: str) -> bool:
    """Cask.CheckCaskEndNo (Cask:2468-2530)。B 列最終行の A 列の値と DB の最後(スタートNo 順)の [4] を比べる。"""
    fx = book.fx
    try:
        ws = book.ws(CASK)
        last = ws.last_row(2)
        if last < 1:
            fx.msg("「キャスク検査」シートにデータがありません。", flow.WARN, "Cask:2490")
            return False
        a = ws.get(last, 1)
        sh_end = 0 if a is None else _clng_variant(a)
        data = get_cask_data(book, lot, ken)
        if data is None:
            fx.msg("DBにデータが見つかりません。" + "\r\n" + "LotNo：" + lot + "　検査No：" + ken, flow.WARN, "Cask:2503")
            return False
        db_end = _clng_variant(_at(data[-1], 4))
        if sh_end == db_end:
            return True
        fx.msg("エンドNoが一致しません。" + "\r\n" + "\r\n" + "シート最終行（A列）：" + str(sh_end) + "\r\n"
               + "DB エンドNo（最終）：" + str(db_end), flow.WARN, "Cask:2519")
        return False
    except VbaError as e:
        fx.msg("チェックエラー：" + "\r\n" + e.description, flow.ERROR, "Cask:2528")
        return False


def validate_cask_input(book: Book) -> bool:
    """UFCask.ValidateCaskInput (UFCask:898-919)。"""
    f = book.cask
    for v, m in ((f.lot_box, "LotNo が未入力です。"), (f.ken_no, "検査No が未入力です。"),
                 (f.min_text, "スタートNo が未入力です。"), (f.max_text, "エンドNo が未入力です。")):
        if v.strip(" ") == "":
            book.fx.msg(m, flow.WARN, "UFCask:911")
            return False
    return True


def check_cask_sequence(book: Book) -> bool:
    """UFCask.CheckCaskSequence (UFCask:926-1053)。既存の範囲(同じスタートNo は上書きなので除く)と合わせて連番か。"""
    f = book.cask
    fx = book.fx
    try:
        start = _clng_variant(f.min_text)
        end = _clng_variant(f.max_text)
        if start > end:
            fx.msg("スタートNo がエンドNo より大きい値です。" + "\r\n" + "スタートNo：" + str(start) + "　エンドNo："
                   + str(end), flow.WARN, "UFCask:958")
            return False
        if start == 1:
            return True
        try:
            with datasource.open_ro(_db_path("伝送用ファイル")) as conn:
                rows = conn.execute(f'SELECT "スタートNo", "エンドNo" FROM "{TABLE}" WHERE "LotNo" = ? COLLATE NOCASE '
                                    f'AND "検査No" = ? COLLATE NOCASE AND "スタートNo" <> ? ORDER BY "スタートNo"',
                                    (f.lot_box, f.ken_no, start)).fetchall()
        except (datasource.DbMissing, sqlite3.Error) as e:
            raise VbaError(-2147467259, str(e))
        arr_s = [_clng_variant(r[0]) for r in rows]
        arr_e = [_clng_variant(r[1]) for r in rows]
        arr_s.append(start)
        arr_e.append(end)
        cnt = len(arr_s)
        if cnt <= 1:
            return True
        for i in range(cnt - 1):
            for j in range(cnt - 1 - i):
                if arr_s[j] > arr_s[j + 1]:
                    arr_s[j], arr_s[j + 1] = arr_s[j + 1], arr_s[j]
                    arr_e[j], arr_e[j + 1] = arr_e[j + 1], arr_e[j]
        bad = [i for i in range(cnt - 1) if arr_e[i] + 1 != arr_s[i + 1]]
        if bad:
            fx.msg("番号が連番になっていない箇所があります。", flow.WARN, "UFCask:1034", title="連番チェック")
            return False
        return True
    except VbaError as e:
        fx.msg("連番チェックエラー：" + "\r\n" + e.description, flow.ERROR, "UFCask:1043")
        return False


def save_cask_to_db(book: Book, dt_work: datetime) -> bool:
    """UFCask.SaveCaskToDB (UFCask:1056-1190)。

    STEP1: 同じ LotNo・検査No・スタートNo があれば エンドNo・作業時間 を更新、無ければ 管理番号 = MAX+1 で追加。
    STEP2: スタートNo = 1 のとき、同じ LotNo・検査No でスタートNo <> 1 の行を削除(VBA の DAO の物理削除を DELETE で)。
    STEP1 と STEP2 は別々に確定する(VBA もトランザクションなし。STEP2 が失敗しても STEP1 は残る)。
    """
    f = book.cask
    try:
        try:
            start = int(f.min_text)
            end = int(f.max_text)
        except ValueError:
            raise VbaError(-2147217900, "クエリ式の構文エラー")
        conn = _open_rw("伝送用ファイル")
        try:
            when = dt_work.strftime("%Y/%m/%d %H:%M:%S")
            conn.execute("BEGIN IMMEDIATE")
            try:
                hit = conn.execute(f'SELECT rowid FROM "{TABLE}" WHERE "LotNo" = ? COLLATE NOCASE AND "検査No" = ? '
                                   f'COLLATE NOCASE AND "スタートNo" = ? ORDER BY rowid LIMIT 1',
                                   (f.lot_box, f.ken_no, start)).fetchone()
                if hit:
                    conn.execute(f'UPDATE "{TABLE}" SET "エンドNo" = ?, "作業時間" = ? WHERE rowid = ?',
                                 (end, when, hit[0]))
                else:
                    mx = conn.execute(f'SELECT MAX("管理番号") FROM "{TABLE}"').fetchone()[0]
                    nxt = 1 if mx is None else _clng_variant(mx) + 1
                    conn.execute(f'INSERT INTO "{TABLE}" ("管理番号", "LotNo", "検査No", "スタートNo", "エンドNo", "作業時間") '
                                 f'VALUES (?, ?, ?, ?, ?, ?)', (nxt, f.lot_box, f.ken_no, start, end, when))
                conn.execute("COMMIT")
            except BaseException:
                conn.execute("ROLLBACK")
                raise
            if start == 1:
                conn.execute("BEGIN IMMEDIATE")
                try:
                    conn.execute(f'DELETE FROM "{TABLE}" WHERE "LotNo" = ? COLLATE NOCASE AND "検査No" = ? COLLATE NOCASE '
                                 f'AND "スタートNo" <> 1', (f.lot_box, f.ken_no))
                    conn.execute("COMMIT")
                except BaseException:
                    conn.execute("ROLLBACK")
                    raise
        finally:
            conn.close()
        return True
    except (VbaError, sqlite3.Error) as e:
        desc = e.description if isinstance(e, VbaError) else str(e)
        book.fx.msg("DB保存エラー：" + "\r\n" + desc, flow.ERROR, "UFCask:1174")
        return False


# ------------------------------------------------------------------
# ロット読込(UFCask.ComboBox1_Change)
# ------------------------------------------------------------------
def _field(row: dict[str, Any], name: str, *, width_insensitive: bool) -> Any:
    """ADO `rs.Fields(name)`。英字の大小は区別しない。全角/半角は `width_insensitive` のとき同一視(U4)。"""
    if name in row:
        return row[name]
    for k, v in row.items():
        if k.casefold() == name.casefold():
            return v
    if width_insensitive:
        n = unicodedata.normalize("NFKC", name).casefold()
        for k, v in row.items():
            if unicodedata.normalize("NFKC", k).casefold() == n:
                return v
    raise VbaError(3265, FIELD_MISSING, f"Fields({name})")


def _nn(v: Any, null_raises: bool) -> str:
    """Null を String / TextBox / Val に渡したとき。`null_raises` なら実行時エラー 94、そうでなければ ""。"""
    if v is None:
        if null_raises:
            raise VbaError(94, INVALID_NULL)
        return ""
    return str(v)


def read_lot_row(lot: str) -> Optional[dict[str, Any]]:
    """`SELECT * FROM 仕掛 WHERE ﾛｯﾄ番号 = '{lot}' AND ｵｰﾀﾞｰ板丈 <> '0'` の先頭 1 件(UFCask:426 / UFMit:629)。

    値は SQLite のまま(Null は None)。Access の文字の比較は大文字小文字を区別しないので NOCASE で比べる。
    """
    with datasource.open_ro(_db_path("SIKALOTNOW")) as conn:
        cur = conn.execute('SELECT * FROM "仕掛" WHERE "ﾛｯﾄ番号" = ? COLLATE NOCASE AND "ｵｰﾀﾞｰ板丈" <> \'0\' '
                           'ORDER BY rowid LIMIT 1', (lot,))
        cols = [d[0] for d in cur.description]
        r = cur.fetchone()
    return dict(zip(cols, r)) if r else None


def tolerance_pair(order_p: float, order_m: float, made_p: float, made_m: float) -> tuple[float, float]:
    """公差の取り方 (UFCask:470 / UFMit:680)。オーダー公差の ± の**どちらか一方でも 0** なら両方とも製造公差。"""
    if order_p == 0 or order_m == 0:
        return made_p, made_m
    return order_p, order_m


def combo_change(book: Book, *, fullwidth_fields_readable: bool = False, null_raises: bool = False) -> Effects:
    """UFCask.ComboBox1_Change (UFCask:401-555)。SIKALOTNOW `仕掛` からロット情報を読み、シート上部に寸法・公差を書く。

    `fullwidth_fields_readable`(U4): UFCask は欄名の 4 つを全角カナで読む(実際の欄は半角)。
      False = 欄が見つからず最初の `rs.Fields("オーダー番号")` でエラー処理に落ちる(「DB読み込みエラー:」、シートは変わらない)
      True  = ADO が全角/半角を同一視して読める
    `null_raises`: Null の欄を Val / String / TextBox に入れたとき実行時エラー 94 になるか(受領 DB ではオーダー公差の
      9 割が Null。運用で読めているなら Null ではなく長さ 0 の文字列の可能性が高い。既定は "" として読む)。
    """
    fx = book.fx
    f = book.cask
    book.events_enabled = False
    if not book.has(CASK):
        return fx
    if len(f.combo) != 7:
        return fx
    if f.loadtext == "off":
        return fx
    ws = book.ws(CASK)
    no_fl = False
    try:
        try:
            row = read_lot_row(f.combo)
        except datasource.DbMissing as e:
            raise VbaError(-2147467259, str(e))
        if row is not None:
            no_fl = True
            fv = lambda name: _field(row, name, width_insensitive=fullwidth_fields_readable)  # noqa: E731
            past: list[Any] = ["'" + ("" if fv(CASK_FIELDS[0]) is None else str(fv(CASK_FIELDS[0])))]
            for name in CASK_FIELDS[1:]:
                past.append(fv(name))
            book.write(ws, 2, 4, f.combo)
            f.ken_no = _nn(fv("検査番号"), null_raises)
            val = lambda name: vbx.val(_nn(fv(name), null_raises))  # noqa: E731
            book.write(ws, 3, 2, vbx.fmt(val("製造板厚"), "0.000"))
            b3 = ws.get(3, 2)
            if vbx.vcompare(100.0, b3 if b3 is not None else 0.0) > 0:
                count_select(book, 3)
            elif vbx.vcompare(100.0, b3 if b3 is not None else 0.0) <= 0:
                count_select(book, 4)
            book.write(ws, 3, 4, vbx.fmt(val("製造板幅"), "0.0"))
            book.write(ws, 3, 6, vbx.fmt(val("製造板丈"), "0.0"))
            for row_no, key, pat in ((2, "板厚", "0.000"), (3, "板幅", "0.0"), (4, "板丈", "0.0")):
                op, om = val(key + "公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ"), val(key + "公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ")
                mp, mm = val(key + "公差_製造_ﾌﾟﾗｽ"), val(key + "公差_製造_ﾏｲﾅｽ")
                op, om = tolerance_pair(op, om, mp, mm)
                book.write(ws, row_no, 15, vbx.fmt(op, pat))
                book.write(ws, row_no, 13, vbx.fmt(om, pat))
                book.write_formula(ws, row_no, 19, sh.TOLERANCE_FORMULAS[(row_no, 19)])
                book.write_formula(ws, row_no, 17, sh.TOLERANCE_FORMULAS[(row_no, 17)])
        if not no_fl:
            for r, c in ((3, 2), (3, 4), (3, 6), (2, 13), (2, 15), (2, 17), (2, 19), (3, 13), (3, 15), (3, 17),
                         (3, 19), (4, 13), (4, 15), (4, 17), (4, 19)):
                book.write(ws, r, c, "")
            f.ken_no = ""
            if f.combo_tag != "Load":
                fx.msg("Text.データなし　手入力よろしく", source="UFCask:524")
            act = book.active
            for (r, c), formula in sh.TOLERANCE_FORMULAS.items():
                act.set_formula(r, c, formula)
            return fx                               # EnableEvents は False のまま(VBA どおり)
        fx.calls.append(Call("HistorySave", (past,)))
        book.activate_sheet(CASK)
        fx.calls.append(Call("PackingNoCheck/UFPack_Open", (f.lot_box,)))
        book.events_enabled = True
        return fx
    except VbaError as e:
        book.events_enabled = True
        fx.msg("DB読み込みエラー:" + "\r\n" + e.description, flow.ERROR, "UFCask:554")
        return fx


def lot_box_change(book: Book, **kw: Any) -> Effects:
    """UFCask.LotBox_Change (UFCask:374-387)。7 桁になったら ComboBox1 に写す(→ ロット読込)。"""
    fx = book.fx
    if not book.visible(CASK):
        return fx
    book.events_enabled = False
    if len(book.cask.lot_box) != 7:
        return fx
    book.cask.combo = book.cask.lot_box
    combo_change(book, **kw)
    _move_tg_activate(book)
    book.events_enabled = True
    return fx


def _move_tg_activate(book: Book) -> None:
    """`MoveTg("A", "キャスク検査").Activate`。空きが無いと Nothing.Activate で実行時エラー 91。"""
    pos = sh.move_tg(book, "A", CASK)
    if pos is None:
        raise flow.Stop("実行時エラー '91':" + "\r\n\r\n" + "オブジェクト変数または With ブロック変数が設定されていません。",
                        source="Cask:MoveTg", messages=list(book.fx.messages))
    book.activate(*pos)


# ------------------------------------------------------------------
# 印刷・保存
# ------------------------------------------------------------------
def print_setup(book: Book) -> Optional[PrintJob]:
    """Cask.キャスク印刷範囲 (Cask:807-900)。A6:AB{B 列最終行}、A4 横、ヘッダ/フッタ。"""
    if not book.has(CASK):
        return None
    ws = book.ws(CASK)
    n = ws.last_row(2)
    ws.fill(7, 2, 706, 28, sh.WHITE)
    ws.clear_underline(7, 2, 706, 28)
    data = get_cask_data(book, book.cask.lot_box, book.cask.ken_no)
    footer = ""
    if data is not None:
        for i, row in enumerate(data):
            if footer != "":
                footer += "  "
            footer += "梱包No" + str(i + 1) + "_" + _vstr(_at(row, 3)) + "～" + _vstr(_at(row, 4))
    x, y, z = sh.val_cell(ws.get(3, 2)), sh.val_cell(ws.get(3, 4)), sh.val_cell(ws.get(3, 6))
    am, ap = sh.val_cell(ws.get(2, 13)), sh.val_cell(ws.get(2, 15))
    hm, hp = sh.val_cell(ws.get(3, 13)), sh.val_cell(ws.get(3, 15))
    tm, tp = sh.val_cell(ws.get(4, 13)), sh.val_cell(ws.get(4, 15))
    soku = "ﾉｷﾞｽ" if x >= 100 else "ﾏｲｸﾛ"
    f = book.cask
    job = PrintJob(
        sheet=CASK, print_area="$A$6:$AB$" + str(n),
        left_header='&"Meiryo UI"&8' + "*" + sh.cell_str(ws.get(2, 4)) + "*" + " 検査員：" + f.kname + " " + "\r"
        + vbx.fmt(x, "0.000") + "×" + vbx.fmt(y, "0.0") + "×" + vbx.fmt(z, "0.0") + "\r"
        + "*ｾﾞﾛ校正_" + soku + ":" + vbx.fmt(ws.get(2, 29), "0.000")
        + " *ｾﾞﾛ校正_ｺﾝﾍﾞｯｸｽ:" + vbx.fmt(ws.get(3, 29), "0.0")
        + " *1000mm校正_ｺﾝﾍﾞｯｸｽ:" + vbx.fmt(ws.get(3, 31), "0.0"),
        right_header='&"Meiryo UI"&8' + "*厚公差：" + vbx.fmt(x - am, "0.000") + "～" + vbx.fmt(x + ap, "0.000") + "\r"
        + "*巾公差：" + vbx.fmt(y - hm, "0.0") + "～" + vbx.fmt(y + hp, "0.0") + "\r"
        + "*丈公差：" + vbx.fmt(z - tm, "0.0") + "～" + vbx.fmt(z + tp, "0.0"),
        left_footer='&"Meiryo UI"&10' + footer,
        right_footer='&"Meiryo UI"&8' + "*" + sh.fmt_date(ws.get(1, 2), "yyyy/mm/dd h:mm") + "～"
        + sh.fmt_date(book.now(), "yyyy/mm/dd h:mm") + " 印刷",
        paper="A4", orientation="landscape", fit_wide=1, fit_tall=None,
        margins={"header": 0.25, "footer": 0.2, "left": 0, "right": 0, "top": 0.8, "bottom": 0.35},
        center_h=True, snapshot=ws.snapshot())
    return job


def _hinsho_copy(book: Book, save_ws: sh.Sheet, data: Optional[list]) -> sh.Sheet:
    """`品証ｺﾋﾟ-用` (Cask:534-597)。キャスク保存 A6:AB{n} を D1 へ貼り、F,H,J,L:M,O,Q,S,U,V:X,Z を削除したもの。"""
    f = book.cask
    q = sh.Sheet("品証ｺﾋﾟ-用")
    n = save_ws.last_row(2)
    keep = [1, 2, 4, 6, 8, 11, 13, 15, 17, 22, 24, 25, 26, 27, 28]     # 削除後に残る元の列(A,B,D,F,H,K,M,O,Q,V,X〜AB)
    for r in range(6, max(n, 6) + 1):
        for k, src_c in enumerate(keep):
            q.set(r - 5, 4 + k, save_ws.get(r, src_c))
    soku = ""
    if sh.val_cell(save_ws.get(3, 2)) >= 100.0:
        soku = "ﾉｷﾞｽ"
    if sh.val_cell(save_ws.get(3, 2)) < 100.0:
        soku = "ﾏｲｸﾛ"
    put = lambda r, c, v: q.set(r, c, sh.to_cell(v))  # noqa: E731
    put(1, 1, "LotNo")
    q.set(1, 2, save_ws.get(2, 4))
    put(2, 1, "検査No")
    q.set(2, 2, save_ws.get(1, 26))
    put(3, 1, "Size")
    put(3, 2, sh.cell_str(save_ws.get(3, 2)) + "×" + sh.cell_str(save_ws.get(3, 4)) + "×" + sh.cell_str(save_ws.get(3, 6)))
    put(4, 1, "ｺﾝﾍﾞｯｸｽNo")
    put(4, 2, f.cno)
    put(5, 1, "ﾏｲｸﾛNo")
    put(5, 2, f.mno + " S/N" + f.seri_no)
    put(6, 1, "検査員")
    put(6, 2, f.kname)
    put(1, 4, "板番号")
    put(7, 1, "作業日")
    put(7, 2, sh.fmt_date(save_ws.get(1, 2), "yyyy/mm/dd  h:mm") + "～" + sh.fmt_date(save_ws.get(1, 4), "yyyy/mm/dd  h:mm"))
    put(8, 1, "ｾﾞﾛ校正:" + soku)
    put(8, 2, f.micro)
    put(9, 1, "ｾﾞﾛ校正:" + "ｺﾝﾍﾞｯｸｽ")
    put(9, 2, f.convex)
    put(10, 1, "1000mm校正:" + "ｺﾝﾍﾞｯｸｽ")
    put(10, 2, f.convex_1000)
    if data is not None:
        for i, row in enumerate(data):
            put(12 + i, 2, "梱包No" + str(i + 1) + ":" + _vstr(_at(row, 3)) + "～" + _vstr(_at(row, 4)))
    return q


def _save_sheet(book: Book, *, with_hinsho: bool) -> sh.Sheet:
    """キャスク保存シート(Cask:479-522 / 702-734)。"""
    f = book.cask
    src = book.ws(CASK)
    s = src.copy("キャスク保存")
    s.underline = set()
    s.interior = {}
    s.visible = False
    put = lambda r, c, v: s.set(r, c, sh.to_cell(v))  # noqa: E731
    put(1, 25, "検査番号")
    put(1, 26, f.ken_no)
    put(2, 25, "検査員")
    put(3, 25, "ｺﾝﾍﾞｯｸｽ")
    put(4, 25, "ﾏｲｸﾛ")
    put(2, 26, f.kname)
    put(3, 26, f.cno)
    put(4, 26, f.mno + "  S/N" + f.seri_no)
    if with_hinsho:
        soku = ""
        if sh.val_cell(s.get(3, 2)) >= 100.0:
            soku = "ﾉｷﾞｽ"
        if sh.val_cell(s.get(3, 2)) < 100.0:
            soku = "ﾏｲｸﾛ"
        put(2, 28, "ｾﾞﾛ校正:" + soku)
        put(3, 28, "ｾﾞﾛ校正:" + "ｺﾝﾍﾞｯｸｽ")
        put(3, 30, "1000mm校正:" + "ｺﾝﾍﾞｯｸｽ")
        put(2, 29, f.micro)
        put(3, 29, f.convex)
        put(3, 31, f.convex_1000)
    put(1, 3, "～")
    put(1, 4, sh.fmt_date(book.now(), "yyyy/mm/dd h:mm"))
    s.col_width.update({2: 8.5, 11: 8.5, 26: 8.5, 22: 18})
    return s


def _password(lot: str, today: datetime) -> str:
    return lot + "." + today.strftime("%Y.%m.%d")


def _save_numbered(book: Book, plan: SavePlan, set_path: str, day_lot: str, lot: str) -> Optional[str]:
    plan.steps.append(SaveStep("mkdir", set_path))
    path, err = sh.numbered_path(book, set_path, day_lot)
    if path is None:
        plan.error = err
        book.fx.msg(err, flow.ERROR, "Cask:638")
        raise flow.Stop(err, source="Cask:639", messages=list(book.fx.messages))
    plan.steps.append(SaveStep("save", path, password=_password(lot, book.now()), readonly=True))
    return path


def cask_save(book: Book) -> Optional[SavePlan]:
    """Cask.キャスク専用保存 (Cask:445-682)。キャスク保存 + 品証ｺﾋﾟ-用 を `C:\\検査履歴\\yyyy.mm\\yyyymmdd_{Lot}\\` に保存し、
    サーバ(キャスク_履歴)と工場共有(【梱包】キャスク)へ写す。必要欄が空なら「フォーム上で必要欄を記入してください」で中止。
    """
    f = book.cask
    book.delete_sheet("キャスク保存")
    book.delete_sheet("品証ｺﾋﾟ-用")
    save_ws = _save_sheet(book, with_hinsho=True)
    book.add_sheet(save_ws)
    if f.cno == "" or f.mno == "" or f.kname == "" or f.seri_no == "":
        book.fx.msg("フォーム上で必要欄を記入してください", source="Cask:526")
        return None
    data = get_cask_data(book, f.lot_box, f.ken_no)
    hin = _hinsho_copy(book, save_ws, data)
    book.add_sheet(hin)
    plan = SavePlan("キャスク専用保存", sheets={hin.name: hin.snapshot(), save_ws.name: save_ws.snapshot()})
    book.fx.saves.append(plan)
    today = book.now()
    month = today.strftime("%Y.%m")
    lot = sh.cell_str(save_ws.get(2, 4))
    day_lot = today.strftime("%Y%m%d") + "_" + lot
    set_path = paths.join(paths.get("local_history"), month, day_lot)
    saved = _save_numbered(book, plan, set_path, day_lot, lot)
    for name in ("品証ｺﾋﾟ-用", "キャスク保存", "ｷｬｽｸ作業中保存"):
        book.delete_sheet(name)
    hist = paths.get("history")
    srv = paths.join(hist, "キャスク_履歴", month)
    plan.steps.append(SaveStep("mkdir", srv))
    fname = saved.split("\\")[-1]
    plan.steps.append(SaveStep("copy", paths.join(srv, fname), src=saved))
    if f.line_name == "Test":
        kou = paths.join(hist, "Test\\Test(工場共有)_履歴\\キャスク_履歴")
    else:
        kou = paths.join(paths.get("factory"), "【梱包】キャスク", month)
    plan.steps.append(SaveStep("mkdir", kou))
    plan.steps.append(SaveStep("copy", paths.join(kou, fname), src=saved))
    return plan


def justincase_save(book: Book) -> Optional[SavePlan]:
    """Cask.JustinCase_KySave (Cask:683-788)。新規発行の前の念のための保存(キャスク保存 + キャスク検査 → 検査履歴\\JustinCaseSave)。"""
    if not book.has(CASK):
        return None
    f = book.cask
    save_ws = _save_sheet(book, with_hinsho=False)
    save_ws.visible = True
    book.add_sheet(save_ws)
    plan = SavePlan("JustinCase_KySave", sheets={save_ws.name: save_ws.snapshot(), CASK: book.ws(CASK).snapshot()})
    book.fx.saves.append(plan)
    today = book.now()
    hist = paths.get("history")
    base = paths.join(hist, "Test\\Test_履歴\\JustinCaseSave") if f.line_name == "Test" else paths.join(hist, "JustinCaseSave")
    lot = sh.cell_str(save_ws.get(2, 4))
    day_lot = today.strftime("%Y%m%d") + "_" + lot
    _save_numbered(book, plan, paths.join(base, day_lot), day_lot, lot)
    return plan


# ------------------------------------------------------------------
# フォームの操作
# ------------------------------------------------------------------
def _ask(answers: Optional[dict], key: str, text: str, source: str, fx: Effects) -> Optional[bool]:
    a = (answers or {}).get(key)
    if a is None:
        raise _Ask(flow.Question(key, text, source=source, messages=list(fx.messages)))
    return bool(a)


class _Ask(Exception):
    def __init__(self, q: flow.Question) -> None:
        self.q = q


def new_sheet(book: Book, answers: Optional[dict] = None) -> Union[Effects, flow.Question]:
    """UFCask.CommandButton1_Click (UFCask:41-79) 新規発行。"""
    fx = book.begin()
    f = book.cask
    if f.tag == "off":
        return fx
    try:
        yes = _ask(answers, "cask_new", "新規発行しますか？現在のｼｰﾄは削除されます。", "UFCask:44", fx)
    except _Ask as a:
        return a.q
    fx.calls.append(Call("Register_Check/Daily_Create"))
    if yes:
        justincase_save(book)
        cask_copy(book)
        fx.logs.append(LogCall("*UFKen:" + "ｷｬｽｸ検査ｼｰﾄ発行", "", _cask_log_path(f.line_name)))
        f.inspection = "OFF"
        f.lot_box = ""
        f.ken_no = ""
        f.kname = ""
        f.combo = ""
        f.convex = ""
        f.micro = ""
    else:
        fx.msg("ｷｬﾝｾﾙ処理", source="UFCask:76")
    book.events_enabled = True
    return fx


def transfer(book: Book, *, now: Optional[datetime] = None) -> Effects:
    """UFCask.CommandButton3_Click (UFCask:182-232) 転送。入力・連番チェック → 点数切替 → DB 保存 → ラベル → 最大最小 → Ainput。"""
    fx = book.begin()
    f = book.cask
    if f.tag == "off":
        return fx
    book.events_enabled = False
    dt_work = now or book.now()
    if not validate_cask_input(book):
        return fx
    if not check_cask_sequence(book):
        return fx
    ws = book.ws(CASK)
    if not sh.blank(ws.get(3, 2)):
        if 100.0 > sh.val_cell(ws.get(3, 2)):
            count_select(book, 3)
        elif 100.0 <= sh.val_cell(ws.get(3, 2)):
            count_select(book, 4)
    if save_cask_to_db(book, dt_work):
        add_cask_labels(book, f.lot_box, f.ken_no)
        try:
            max_min(book)
        except VbaError as e:
            raise flow.Stop(f"実行時エラー '{e.number}':" + "\r\n\r\n" + e.description, source="Cask:キャスク用最大最小",
                            messages=list(fx.messages))
        fx.calls.append(Call("Ainput_Startup", ("Cask", "L2")))
    book.events_enabled = True
    return fx


def _required_missing(f: sh.CaskForm) -> bool:
    return f.cno == "" or f.mno == "" or f.kname == "" or f.seri_no == ""


def print_and_save(book: Book, answers: Optional[dict] = None, *, u6_raises: bool = False) -> Union[Effects, flow.Question]:
    """UFCask.CommandButton2_Click (UFCask:110-174) 印刷。ループ判定 → 必須欄 → エンドNo 照合 → 印刷 → 保存 → 保存用に履歴。"""
    fx = book.begin()
    f = book.cask
    f.tag = "off"
    book.activate_sheet(CASK)
    try:
        loop_judge(book, u6_raises=u6_raises)
    except VbaError as e:
        raise flow.Stop(f"実行時エラー '{e.number}':" + "\r\n\r\n" + e.description, source="Cask:ループ判定")
    if _required_missing(f):
        fx.msg("フォーム上で必要欄を記入してください", flow.WARN, "UFCask:119")
        return fx                                   # Tag は "off" のまま(VBA どおり)
    if check_cask_end_no(book, f.lot_box, f.ken_no):
        job = print_setup(book)
    else:
        return fx                                   # Tag は "off" のまま
    fx.sounds.append("印刷.wav")
    try:
        yes = _ask(answers, "cask_print", "キャスク検査を印刷しますか？", "UFCask:139", fx)
    except _Ask as a:
        return a.q
    if yes:
        book.ws(CASK).visible = True
        if job is not None:
            fx.prints.append(job)
        cask_save(book)
        row = sh.history_row(book, book.ws(CASK), "ｷｬｽｸ", cask_bug=True)
        if row is not None:
            fx.history.append(row)
        fx.sounds.append("完了.wav")
    else:
        fx.msg("ｷｬﾝｾﾙ処理", source="UFCask:169")
    f.tag = ""
    return fx


def save_only(book: Book) -> Effects:
    """UFCask.CommandButton4_Click (UFCask:234-243) 保存のみ。"""
    fx = book.begin()
    f = book.cask
    f.tag = "off"
    if _required_missing(f):
        fx.msg("フォーム上で必要欄を記入してください", source="UFCask:237")
        return fx
    cask_save(book)
    f.tag = ""
    return fx


def test_print_save(book: Book) -> Effects:
    """UFCask.CommandButton6_Click (UFCask:308-323)。点数を 3 にし、照合が合えば印刷範囲を設定、照合に関係なく保存(試験用の疑い)。"""
    fx = book.begin()
    f = book.cask
    count_select(book, 3)
    if check_cask_end_no(book, f.lot_box, f.ken_no):
        job = print_setup(book)
        if job is not None:
            fx.prints.append(job)
    cask_save(book)
    return fx


def cancel_input(book: Book) -> Effects:
    """UFCask.CommandButton5_Click (UFCask:245-306) 取消。アクティブ列に応じて直前の入力を消す。"""
    fx = book.begin()
    ar, ac = book.active_cell
    book.events_enabled = False
    f = book.cask
    f.hflag = "off"
    if book.has(CASK):
        book.activate_sheet(CASK)
        ws = book.ws(CASK)

        def clear(c1: int, c2: int) -> None:
            for c in range(c1, c2 + 1):
                ws.set(ar, c, None)

        if ac in (4, 6, 13, 15, 17, 19, 22):
            clear(ac - 2, ac - 1)
            ws.fill(ar, ac - 2, ar, ac, sh.WHITE)
            book.activate(ar, ac - 2)
            ws.interior[(ar, ac - 2)] = sh.LIGHT_BLUE
        elif ac == 11:
            clear(6, 7)
            ws.fill(ar, 7, ar, ac, sh.WHITE)
            book.activate(ar, 6)
            ws.interior[(ar, 6)] = sh.LIGHT_BLUE
        elif ac == 20:
            clear(19, 19)
            ws.fill(ar, 19, ar, ac, sh.WHITE)
            book.activate(ar, 19)
            ws.interior[(ar, 19)] = sh.LIGHT_BLUE
        elif ac == 24:
            clear(19, 23)
            ws.fill(ar, 19, ar, ac, sh.WHITE)
            book.activate(ar, 19)
            ws.interior[(ar, 19)] = sh.LIGHT_BLUE
        elif 25 <= ac <= 28:
            ws.interior[(ar, ac)] = sh.WHITE
            clear(ac - 1, ac - 1)
            book.activate(ar, ac - 1)
            ws.interior[(ar, ac - 1)] = sh.LIGHT_BLUE
    f.hflag = "on"
    book.events_enabled = True
    return fx


def toggle_inspection(book: Book, *, u6_raises: bool = False) -> Effects:
    """UFCask.CommandButton9_Click (UFCask:329-359) 巾丈測定 ON/OFF。"""
    fx = book.begin()
    if CASK not in book.sheets:
        return fx
    ws = book.ws(CASK)
    f = book.cask
    if f.inspection == "OFF":
        f.inspection = "ON"
        _lock_inspection(ws, False)
        ws.fill(6, 11, 706, 28, sh.WHITE)
        try:
            loop_judge(book, u6_raises=u6_raises)
        except VbaError as e:
            raise flow.Stop(f"実行時エラー '{e.number}':" + "\r\n\r\n" + e.description, source="Cask:ループ判定")
    elif f.inspection == "ON":
        f.inspection = "OFF"
        ws.fill(6, 11, 706, 28, sh.GRAY)
        _lock_inspection(ws, True)
    _move_tg_activate(book)
    book.events_enabled = True
    return fx


def enable_events(book: Book) -> Effects:
    """UFCask.CommandButton8_Click / UFMitsubishi.CommandButton16・21_Click: EnableEvents を戻す。"""
    fx = book.begin()
    book.events_enabled = True
    return fx


def initialize(book: Book, *, u6_raises: bool = False, **lot_kw: Any) -> Effects:
    """UFCask.UserForm_Initialize (UFCask:817-873)。キャスク検査シートからフォームの値を戻す。"""
    fx = book.begin()
    f = book.cask
    if f.tag == "off":
        return fx
    f.combo_tag = "Load"
    book.events_enabled = True
    if book.has(CASK) and CASK in book.sheets:
        f.loadtext = "off"
        ws = book.ws(CASK)
        book.activate_sheet(CASK)
        f.lot_box = sh.cell_str(ws.get(2, 4))
        lot_box_change(book, **lot_kw)
        f.ken_no = sh.cell_str(ws.get(2, 25))
        f.kname = sh.cell_str(ws.get(3, 25))
        f.cno = sh.cell_str(ws.get(2, 26))
        f.mno = sh.cell_str(ws.get(3, 26))
        f.seri_no = sh.cell_str(ws.get(4, 26))
        f.micro = sh.cell_str(ws.get(2, 29))
        f.convex = sh.cell_str(ws.get(3, 29))
        f.convex_1000 = sh.cell_str(ws.get(3, 31))
        f.count_st = sh.cell_str(ws.get(3, 33))
        try:
            loop_judge(book, u6_raises=u6_raises)
        except VbaError as e:
            raise flow.Stop(f"実行時エラー '{e.number}':" + "\r\n\r\n" + e.description, source="Cask:ループ判定")
        if f.inspection == "OFF":
            ws.fill(6, 11, 706, 28, sh.GRAY)
            _lock_inspection(ws, True)
    fx.calls.append(Call("VisibleSheets"))
    fx.calls.append(Call("HiddenSheets"))
    add_cask_labels(book, f.lot_box, f.ken_no)
    f.combo_tag = ""
    f.loadtext = "on"
    return fx


def activate(book: Book, *, to_time: str = "", to_tyo: str = "") -> Effects:
    """UFCask.UserForm_Activate (UFCask:777-809)。キャスク系シートを表示し、1000mm 校正値を読む。"""
    fx = book.fx
    for name in (CASK, "ｷｬｽｸ作業中保存", "キャスク保存"):
        if name in book.sheets or name in book.extra:
            book.ws(name).visible = True
    if CASK not in book.sheets:
        fx.msg("「キャスク検査」シートが見つかりません。" + "\r\n" + "処理を中断します。", flow.ERROR, "UFCask:800", title="エラー")
        book.cask.visible = False
        return fx
    proofreading_search(book, to_time, to_tyo)
    book.activate_sheet(CASK)
    return fx


def open_form(book: Book, line_name: str, *, to_time: str = "", to_tyo: str = "", **kw: Any) -> Effects:
    """UF1.KYB_Click → UFCask.Show: LineName を入れ、Initialize → 表示 → Activate。"""
    book.cask.line_name = sh.CaskForm(line_name=line_name).line_name
    fx = initialize(book, **kw)
    book.cask.visible = True
    activate(book, to_time=to_time, to_tyo=to_tyo)
    return fx


def query_close(book: Book) -> Effects:
    """UFCask.UserForm_QueryClose (UFCask:875-895)。キャスク系シートを隠す。"""
    fx = book.begin()
    for name in (CASK, "ｷｬｽｸ作業中保存", "キャスク保存"):
        if name in book.sheets or name in book.extra:
            book.ws(name).visible = False
    book.cask.visible = False
    return fx
