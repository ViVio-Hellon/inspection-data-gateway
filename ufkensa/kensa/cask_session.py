"""キャスク検査の画面(VBA の UFCask + キャスク検査シート)のセッション ── 画面(app/routes/cask.py)はここだけを呼ぶ

部品は `modes/sheet.py`(Sheet/Book・判定)と `modes/cask.py`(KySh のイベント・UFCask のボタン)。ここは
- 1 プロセスに 1 冊の `Book`(キャスク検査シートと UFCask の状態)を持ち、操作を 1 つずつ通す
- 部品が返す `Effects`(知らせ・音・ログ・フォーム表示・他モジュール呼び出し・印刷・保存・保存用の行)を実行する
- 操作のたびにシートを手元DB へ書く(VBA の作業中保存 `ｷｬｽｸ作業中保存`)。開き直したら読み戻す

**開き直したとき**: VBA はブックにシートが残り、フォームは Initialize でシートの Y2:AG3 から値を戻す。ここも同じく
シートだけを戻し、フォーム(UFCask)は `open()` の Initialize(`cask.open_form`)で戻す。対角公差・スタート/エンドNo・
巾丈測定の ON は VBA と同じく開き直すと消える(OFF に戻る)。

各操作と VBA:
    open()                 UF1.KYB_Click(Register_Check → 検査員登録をして下さい)→ UFCask.Show(Initialize・Activate)
    set_lot(lot)           LotBox(KeyPress で英数字・大文字)→ LotBox_Change → ComboBox1_Change
    set_field(name, value) KenNo / Kname / Cno / Mno / SeriNo / MINText / MAXText / Tai の入力
    zero(text)             StartTB_AfterUpdate(ゼロ校正)
    receive(text)          測定器(キーボードウェッジ)・手入力 → アクティブセルへ入力 → KySh1.Worksheet_Change
    select(row, col)       セルを選ぶ → KySh1.Worksheet_SelectionChange
    transfer()             CommandButton3「決定」(転送)
    print_out(answers)     CommandButton2「印刷」
    save()                 CommandButton4「ｻｰﾊﾞｰ保存」
    new_sheet(answers)     CommandButton1「ｷｬｽｸ用ｼｰﾄ発行」
    cancel()               CommandButton5「Back space」
    toggle_inspection()    CommandButton9「巾丈測定」
    resume()               CommandButton8「処理再開」(EnableEvents を戻す)
    proofread()            CommandButton15「校正読み込み」
    close()                UserForm_QueryClose(×)
    pack_select(i)         包装仕様書の窓(UF_PackNo)の引当の行をクリック(LabelCls → JPGLoad)
    pack_close(confirmed)  UF_PackNo の「閉じる」(CommandButton2。確認の印が要る)

ComboBox1_Change の続き: HistorySave(Lot情報履歴)→ PackingNoCheck / UFPack_Open(`packing`。包装仕様書の画像と
引当の一覧をモーダルで出す)。どちらかで実行時エラーになれば「DB読み込みエラー:」(ErrHandler_ComboChange)。

印刷(PrintJob)は D8 でブラウザで開く HTML(report_html.render)。保存(SavePlan)は D7 で xlsx の代わりに
シートごとの PDF(Edge。Windows 以外は HTML)+ CSV。
"""
from __future__ import annotations

import json
import re
import logging
import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional, Union

from .. import trace
from ..logging_utils import get_logger
from . import datasource, legacy_log, lot as lotmod, packing, paths, report_html, terminal, vbx, wiring, workers
from . import sound as snd
from .flow import ERROR, INFO, WARN, Message, Question, Stop, unexpected_stop
from .modes import cask
from .modes import sheet as sh
from .report_model import ReportModel, format_cell
from .session import Result
from .special_common import SubInspectionCommon
from .store import HOZON, HOZON_HEADERS, LOT_HISTORY, LocalStore

log = get_logger("kensa.cask_session")

SCREEN = "cask"
STATE_TABLE = "画面状態"                       # この画面が持つ状態(シートの全部)。1 画面 1 行
WORK_SAVE = "ｷｬｽｸ作業中保存"                   # VBA の作業中保存シート(値だけの写し)
#: VBA の音の手続き(補助:282-358)→ キャスクの出来事(D29。ファイルは 設定 →「音」)
SOUNDS = snd.by_procedure(snd.CASK_OK, snd.CASK_NG, snd.CASK_DIAG)
#: 保護中のシートのロックされたセルに打ち込んだときの Excel の警告
PROTECTED_TEXT = ("変更しようとしているセルやグラフは保護されているシート上にあります。"
                  "変更するには、[校閲] タブの [シート保護の解除] をクリックします (パスワードの入力が必要な場合もあります)。")
NEED_WORKER = "検査員登録をして下さい"            # UF1.KYB_Click (UF1:3233)
#: フォームの入力欄 → KeyPress の制限("alnum" = 英数字だけ・小文字は大文字に)
FIELDS = {"ken_no": "", "kname": "", "cno": "", "mno": "", "seri_no": "alnum", "min_text": "alnum",
          "max_text": "alnum", "tai": ""}
#: 画面に並べる列(シートの列番号, 判定の列)。見出しはシートの 6 行目
COLUMNS: tuple[tuple[int, int], ...] = (
    (2, 3), (4, 5), (6, 7), (8, 9), (11, 12), (13, 14), (15, 16), (17, 18), (19, 0), (20, 0), (21, 0), (22, 0),
    (23, 0), (24, 0), (25, 0), (26, 0), (27, 0), (28, 0))
FIRST_ROW, LAST_ROW = 7, 706
#: シート上部の手入力できる欄(ロットが読めないときは「Text.データなし　手入力よろしく」でここに打つ)
HEADER_CELLS = {(3, 2): "寸法 厚", (3, 4): "寸法 巾", (3, 6): "寸法 丈", (2, 13): "厚公差 −", (2, 15): "厚公差 +",
                (3, 13): "巾公差 −", (3, 15): "巾公差 +", (4, 13): "丈公差 −", (4, 15): "丈公差 +"}
#: 表示の色の名前(Interior の色 → 画面の状態)
_COLOR_STATE = {sh.OK_GREEN: "ok", sh.COLOR_INDEX[10]: "ok", sh.NG_RED: "ng", sh.COLOR_INDEX[3]: "ng",
                sh.LIGHT_BLUE: "next", sh.GRAY: "off"}


def _alnum_upper(text: str) -> str:
    """LotBox/ComboBox1/MINText/MAXText/SeriNo の KeyPress(UFCask:17-39 ほか): 0-9 A-Z はそのまま、a-z は大文字、他は捨てる。"""
    out = []
    for ch in str(text or ""):
        if "0" <= ch <= "9" or "A" <= ch <= "Z":
            out.append(ch)
        elif "a" <= ch <= "z":
            out.append(ch.upper())
    return "".join(out)


def cell_text(v: Any, fmt: str = "General") -> str:
    """セルの見え方(表示形式つき)。日時は CStr、エラー値はそのまま。"""
    if v is None:
        return ""
    if isinstance(v, sh.ExcelError):
        return str(v)
    if isinstance(v, datetime):
        return sh.cstr_date(v)
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    return format_cell(v, fmt or "General")[0]


# ------------------------------------------------------------------
# シートの保存形(JSON)
# ------------------------------------------------------------------
def _enc(v: Any) -> Any:
    if isinstance(v, sh.ExcelError):
        return {"e": str(v)}
    if isinstance(v, datetime):
        return {"d": v.isoformat()}
    return v


def _dec(v: Any) -> Any:
    if isinstance(v, dict):
        if "d" in v:
            return datetime.fromisoformat(v["d"])
        if "e" in v:
            return sh.ExcelError(v["e"])
    return v


def _runs(cells: dict) -> list:
    """{(r, c): 値} → 行ごとに同じ値が続く範囲 [r, c1, c2, 値](塗り・ロックは矩形で付くので小さくなる)。"""
    out: list = []
    for (r, c) in sorted(cells):
        v = cells[(r, c)]
        if out and out[-1][0] == r and out[-1][2] == c - 1 and out[-1][3] == v:
            out[-1][2] = c
        else:
            out.append([r, c, c, v])
    return out


def _unruns(runs: list, conv: Callable[[Any], Any] = lambda x: x) -> dict:
    return {(r, c): conv(v) for r, c1, c2, v in runs for c in range(c1, c2 + 1)}


def sheet_to_json(ws: sh.Sheet) -> dict[str, Any]:
    return {
        "values": [[r, c, _enc(v)] for (r, c), v in sorted(ws.values.items())],
        "formulas": [[r, c, f] for (r, c), f in sorted(ws.formulas.items())],
        "interior": _runs({k: list(v) for k, v in ws.interior.items()}),
        "font_color": _runs({k: list(v) for k, v in ws.font_color.items()}),
        "font_size": _runs(dict(ws.font_size)),
        "number_format": _runs(dict(ws.number_format)),
        "underline": sorted([r, c] for r, c in ws.underline),
        "locked": _runs({k: 1 for k in ws.locked}),
        "col_width": {str(k): v for k, v in ws.col_width.items()},
        "tab_color": ws.tab_color,
    }


def sheet_from_json(name: str, d: dict[str, Any]) -> sh.Sheet:
    ws = sh.Sheet(name)
    ws.values = {(r, c): _dec(v) for r, c, v in d.get("values", [])}
    ws.formulas = {(r, c): f for r, c, f in d.get("formulas", [])}
    ws.interior = _unruns(d.get("interior", []), tuple)
    ws.font_color = _unruns(d.get("font_color", []), tuple)
    ws.font_size = _unruns(d.get("font_size", []))
    ws.number_format = _unruns(d.get("number_format", []))
    ws.underline = {(r, c) for r, c in d.get("underline", [])}
    ws.locked = set(_unruns(d.get("locked", [])))
    ws.col_width = {int(k): v for k, v in d.get("col_width", {}).items()}
    ws.tab_color = d.get("tab_color")
    return ws


# ------------------------------------------------------------------
# 帳票(PrintJob / SavePlan のシート → ReportModel)
# ------------------------------------------------------------------
_EXCEL_CODE = re.compile(r'&"[^"]*"|&\d+')
#: 帳票の列幅(文字数)。原紙の列幅は受領資料に無いので、値の列 7.5・判定の列 2.6 にした
_PRINT_WIDTHS = {1: 4.0, 22: 6.5, 23: 10.0, 28: 8.0}
_JUDGE_COLS = {3, 5, 7, 9, 12, 14, 16, 18, 21}
_A4_LANDSCAPE_PX = 1122          # 297mm を 96dpi で


def _header_text(code: str) -> str:
    """PageSetup のヘッダー/フッター(`&"Meiryo UI"&8…`、改行は vbCr)→ 文字。"""
    return _EXCEL_CODE.sub("", code or "").replace("\r", "\n")


def _parse_area(area: str) -> Optional[tuple[int, int, int, int]]:
    m = re.match(r"^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$", area or "")
    if not m:
        return None
    return int(m.group(2)), sh.col_number(m.group(1)), int(m.group(4)), sh.col_number(m.group(3))


def snapshot_model(snapshot: dict[str, Any], *, kind: str, sheet_name: str,
                   area: Optional[tuple[int, int, int, int]] = None, widths: Optional[dict[int, float]] = None,
                   page: Optional[dict[str, Any]] = None, formats: Optional[dict[tuple[int, int], str]] = None,
                   grid: bool = True) -> ReportModel:
    """セルの値(A1 → 値)を帳票モデルにする。罫線は範囲の全部に細線(原紙の罫線は受領資料に無い)。"""
    vals = {sh.a1(k): v for k, v in snapshot.items()}
    if area is None:
        keys = [k for k, v in vals.items() if v is not None] or [(1, 1)]
        area = (1, 1, max(r for r, _ in keys), max(c for _, c in keys))
    r1, c1, r2, c2 = area
    widths = dict(widths or {})
    thin = ["continuous", "thin"] if grid else ["none", ""]
    cells = []
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            v = vals.get((r, c))
            fmt = (formats or {}).get((r, c), "General")
            text = cell_text(v, fmt)
            cells.append({"ref": f"{sh.col_letter(c)}{r}", "row": r, "col": c,
                          "value": v.isoformat() if isinstance(v, datetime) else v, "text": text,
                          "red": False, "number_format": fmt, "font_size": 9, "bold": r == r1 and r1 == 6,
                          "font": "Meiryo UI", "h_align": "center" if c in _JUDGE_COLS or r == 6 else "",
                          "v_align": "center", "wrap": False, "shrink": False,
                          "borders": {s: list(thin) for s in ("top", "bottom", "left", "right")}})
    col_widths = {}
    for c in range(c1, c2 + 1):
        w = widths.get(c)
        if w is None:
            w = _PRINT_WIDTHS.get(c, 2.6 if c in _JUDGE_COLS else 7.5)
        col_widths[sh.col_letter(c)] = w
    pg = dict(page or {})
    total = sum(report_html._width_px(w) for w in col_widths.values())
    if pg.pop("fit_wide", None) == 1 and total > 0:
        mg = pg.get("margins_inch") or {}
        usable = _A4_LANDSCAPE_PX - int((mg.get("left", 0.3) + mg.get("right", 0.3)) * 96)
        pg["zoom"] = min(100, int(usable * 100 / total))
    return ReportModel(sheet=sheet_name, kind=kind, print_area=area, page=pg, cells=cells, merges=[],
                       col_widths=col_widths, row_heights={}, figure_at=None)


def print_model(job: sh.PrintJob, ws: sh.Sheet) -> ReportModel:
    """Cask.キャスク印刷範囲 の PageSetup と印刷範囲 → 帳票モデル。"""
    m = job.margins or {}
    page = {"orientation": job.orientation, "fit_wide": job.fit_wide,
            "margins_inch": {"top": m.get("top", 0.8), "bottom": m.get("bottom", 0.35),
                             "left": m.get("left", 0), "right": m.get("right", 0)},
            "header_left": _header_text(job.left_header), "header_center": _header_text(job.center_header),
            "header_right": _header_text(job.right_header), "footer_left": _header_text(job.left_footer),
            "footer_center": _header_text(job.center_footer), "footer_right": _header_text(job.right_footer),
            "header_size": 8}
    return snapshot_model(job.snapshot, kind="キャスク検査", sheet_name=job.sheet, area=_parse_area(job.print_area),
                          widths=dict(ws.col_width), page=page, formats=dict(ws.number_format))


def render_html(model: ReportModel, title: str) -> str:
    """report_html.render に、ヘッダー/フッターの改行(vbCr)を活かす指定だけ足す。"""
    return report_html.render(model, title=title).replace("</style>", ".hf span { white-space: pre-line; }\n</style>", 1)


# ------------------------------------------------------------------
# セッション
# ------------------------------------------------------------------
class CaskSession(SubInspectionCommon):
    """プロセスに 1 つのキャスク検査。操作は `_lock` で 1 つずつ通す。

    試験では `store`(手元DB)・`now`(時計)・`uf1`(通常検査の状態を返す関数)・`write_logs` を差し替える。
    """

    def __init__(self, store: Optional[LocalStore] = None, *, now: Callable[[], datetime] = datetime.now,
                 uf1: Optional[Callable[[], Any]] = None, write_logs: bool = True,
                 facility: Callable[[], str] = terminal.facility) -> None:
        self._lock = threading.RLock()
        self._store = store
        self.now = now
        self._uf1 = uf1
        self.write_logs = write_logs
        self.facility = facility
        self.book: Optional[sh.Book] = None
        self.flash: Optional[dict[str, Any]] = None      # UFcount(板No・位置と ○/×)
        self.report: Optional[dict[str, Any]] = None     # 印刷した帳票
        self.saved: Optional[dict[str, Any]] = None      # 保存したファイル
        self.ainput = False                              # 転送で Ainput_Startup("Cask", "L2") が呼ばれた
        self.pack: Optional[packing.PackForm] = None     # 包装仕様書の窓(閉じるまで残る。モーダル)
        self._lot_error = False                          # この操作の ComboBox1_Change がエラー処理へ落ちた
        self._env: Optional[wiring.Env] = None

    # ------------------------------------------------------------------
    # 土台
    # ------------------------------------------------------------------
    @property
    def store(self) -> LocalStore:
        if self._store is None:
            self._store = LocalStore()
        return self._store

    def env(self) -> wiring.Env:
        if self._env is None:
            self._env = wiring.Env(store=self.store, now=self.now, facility=self.facility, write_logs=self.write_logs)
        return self._env

    def uf1_state(self) -> Any:
        """通常検査(UF1)の状態。読むだけ(ライン名・検査員の登録・DBflag)。"""
        if self._uf1 is not None:
            return self._uf1()
        from . import session
        return session.get().state

    def _sink(self) -> list[Message]:
        return self.book.fx.messages

    def _facility(self) -> str:
        return self.facility() or workers.DEFAULT_FACILITY

    def _lot_read_failed(self, e: Exception, msgs: list[Message]) -> None:
        """ComboBox1_Change のエラー処理「DB読み込みエラー:」(UFCask:554)。以後の包装仕様書は飛ばす。"""
        self._lot_error = True
        desc = e.description if isinstance(e, packing.PackingError) else str(e)
        msgs.append(Message("DB読み込みエラー:" + "\r\n" + desc, ERROR, source="UFCask:554"))

    def _history_env(self) -> lotmod.HistoryEnv:
        book, uf1 = self.book, self.uf1_state()
        hooks = wiring.RealHooks(self.env())
        return lotmod.HistoryEnv(dbflag=getattr(uf1, "dbflag", "") or "", now=self.now(),
                                 over12=hooks._over12(uf1) if uf1 is not None else False,
                                 cask_visible=book.visible(sh.CASK), uf1_test=getattr(uf1, "line_option", "") == "Test",
                                 cask_line=book.cask.line_name)

    def _ensure_book(self) -> sh.Book:
        if self.book is None:
            self.book = sh.Book(now=self.now, exists=self._exists)
            self._restore()
        self.book.other_sheets = set(self.store.sheet_names())
        return self.book

    # -- 作業中保存 ---------------------------------------------------
    def _table(self) -> None:
        with self.store._lock, self.store.conn:
            self.store.conn.execute(f'CREATE TABLE IF NOT EXISTS "{STATE_TABLE}" ("画面" TEXT PRIMARY KEY, '
                                    '"内容" TEXT NOT NULL, "更新" TEXT)')

    def _restore(self) -> None:
        """手元DB からキャスク検査シートを戻す。全部の形(JSON)が無ければ `ｷｬｽｸ作業中保存`(xlsm から取り込んだもの)の値。"""
        book = self.book
        assert book is not None
        self._table()
        row = self.store.conn.execute(f'SELECT "内容" FROM "{STATE_TABLE}" WHERE "画面" = ?', (SCREEN,)).fetchone()
        if row:
            try:
                d = json.loads(row[0])
                ws = sheet_from_json(sh.CASK, d["sheet"])
                book.sheets[sh.CASK] = ws
                ac = d.get("active_cell") or [FIRST_ROW, 2]
                book.active_cell = (int(ac[0]), int(ac[1]))
                book.active_name = sh.CASK
                return
            except (ValueError, KeyError, TypeError) as e:
                log.warning("キャスク検査の作業中保存を読めませんでした: %s", e)
        if self.store.exists(WORK_SAVE):
            ws = sh.Sheet(sh.CASK)
            for r, c, v in self.store.sheet(WORK_SAVE).used():
                if isinstance(v, str) and sh._FORMULA_RE.match(v):
                    ws.set_formula(r, c, v)
                else:
                    ws.set(r, c, sh.to_cell(v) if isinstance(v, str) else v)
            book.sheets[sh.CASK] = ws
            book.active_name = sh.CASK

    def _persist(self) -> None:
        """シートの全部(JSON)と、値の写し(`ｷｬｽｸ作業中保存`)を手元DB へ書く。"""
        book = self.book
        if book is None or sh.CASK not in book.sheets:
            return
        ws = book.sheets[sh.CASK]
        content = json.dumps({"sheet": sheet_to_json(ws), "active_cell": list(book.active_cell)}, ensure_ascii=False)
        rows = []
        for (r, c) in set(ws.values) | set(ws.formulas):
            v = ws.get(r, c)
            if isinstance(v, datetime):
                v = sh.cstr_date(v)
            elif isinstance(v, (sh.ExcelError, bool)):
                v = str(v)
            if v is not None:
                rows.append((WORK_SAVE, r, c, v))
        s = self.store
        with s._lock, s.conn:
            s.conn.execute(f'INSERT OR REPLACE INTO "{STATE_TABLE}" ("画面", "内容", "更新") VALUES (?, ?, ?)',
                           (SCREEN, content, self.now().strftime("%Y/%m/%d %H:%M:%S")))
            s.conn.execute('DELETE FROM "セル" WHERE "シート" = ?', (WORK_SAVE,))
            s.conn.execute('INSERT OR REPLACE INTO "シート" ("名前", "順") VALUES (?, 0)', (WORK_SAVE,))
            s.conn.executemany('INSERT INTO "セル" ("シート", "行", "列", "値") VALUES (?, ?, ?, ?)', rows)

    # -- 操作の通し方 -------------------------------------------------
    def _reads_second(self) -> bool:
        """D28: キャスクのロットが、通常検査で「控え(第2)から読む」を選んだロットなら第2だけから読む。"""
        book = self.book
        if book is None:
            return False
        try:
            return wiring.reads_second(self.uf1_state(), book.cask.lot_box)
        except Exception:                                  # noqa: BLE001 - 通常検査が無ければ いつもの読み方
            return False

    def _run(self, fn: Callable[..., Any], *args: Any) -> Result:
        with self._lock, paths.lot_source(self._reads_second):
            book = self._ensure_book()
            if not book.events_enabled:
                # 前の操作が途中で止まり、イベントが止まったまま(VBA の EnableEvents = False の戻し忘れ)。戻さないと
                # セルの入力に何も反応しなくなる。操作は 1 つずつ順に来るので、ここで止まっているのは残り物(2026-11-06)
                logging.getLogger(__name__).warning("イベントが止まったままだったので戻します")
                book.events_enabled = True
            book.begin()
            self._lot_error = False
            extra: list[Message] = []
            q: Optional[Question] = None
            stop: Optional[Stop] = None
            try:
                try:
                    out = fn(*args)
                    if isinstance(out, Question):
                        q = out
                except sh.VbaError as e:                      # イベントの外で起きた実行時エラー(VBA は止まる)
                    raise Stop(f"実行時エラー '{e.number}':" + "\r\n\r\n" + e.description, source=e.where or "VBA") from e
                except datasource.DbMissing as e:
                    raise Stop(str(e), source="datasource") from e
                except (Stop, NotImplementedError):
                    raise
                except Exception as e:                       # noqa: BLE001 - 想定外でも画面は止めない
                    raise unexpected_stop(e, "cask." + getattr(fn, "__name__", "")) from e
            except Stop as e:
                stop = e
            fx = book.fx
            msgs = list(fx.messages)
            try:
                extra = self._apply(fx)
            except Stop as e:
                stop = stop or e
            for m in (stop.messages if stop else []) + (q.messages if q else []):
                if not any(m is x for x in msgs):
                    msgs.append(m)
            msgs.extend(extra)
            try:
                with trace.timed("db.save", "キャスク検査の作業中保存", path=str(self.store.path), table=WORK_SAVE):
                    self._persist()
            except Exception as e:                                # noqa: BLE001 - 書けなくても検査は止めない
                msgs.append(Message("作業中保存を書けませんでした: " + str(e), WARN, source="cask_session"))
            return Result(self.view(), msgs, question=q, stop=stop)

    def _apply(self, fx: sh.Effects) -> list[Message]:
        """部品が積んだもの(Effects)を実行する。返り値は足す知らせ。"""
        out: list[Message] = []
        for s in fx.sounds:
            out.append(Message("", INFO, sound=SOUNDS.get(s, s), source="Cask"))
        for lc in fx.logs:
            if self.write_logs:
                legacy_log.write_log(lc.text, lc.b, lc.path, lc.fname, now=self.now())
        for f in fx.forms:
            self._form(f)
        for e in fx.errors:
            log.info("On Error Resume Next で続けた実行時エラー: %s", e)
        for c in fx.calls:
            out.extend(self._call(c))
        for job in fx.prints:
            out.extend(self._print(job))
        for plan in fx.saves:
            out.extend(self._save(plan))
        for row in fx.history:
            out.extend(self._history(row))
        return out

    def _form(self, f: sh.FormShow) -> None:
        """UFcount(板No・位置の表示)/ UF対角(演出)。UFOK/UFNG はキャスクでは出ない。"""
        book = self.book
        assert book is not None
        ws = book.ws(sh.CASK)
        if f.name == "UFcount":
            self.flash = {"mark": "○" if f.kind == "Maru" else "×", "no": cell_text(ws.get(f.row, 1)),
                          "label": cell_text(ws.get(6, f.col)), "row": f.row, "col": f.col}
        elif f.name in ("UFcount.Maru.Visible", "UFcount.Batu.Visible"):
            r = self.flash.get("row") if self.flash else 0
            self.flash = {"mark": "○" if "Maru" in f.name else "×", "no": cell_text(ws.get(r, 1)) if r else "",
                          "label": "対角", "row": r, "col": 21}

    # -- 他モジュールの呼び出し ---------------------------------------
    def _call(self, c: sh.Call) -> list[Message]:
        if c.name == "Daily_Report":
            return self._daily_report(int(c.args[0]))
        if c.name == "Register_Check/Daily_Create":
            return self._daily_create()
        if c.name == "HistorySave":
            return self._history_save(list(c.args[0]))
        if c.name == "Ainput_Startup":
            self.ainput = True
            return []
        if c.name == "PackingNoCheck/UFPack_Open":
            return [] if self._lot_error else self._packing(str(c.args[0]))   # HistorySave で落ちたら来ない
        # VisibleSheets → HiddenSheets(UFCask.Initialize)はシートの表示だけ。キャスク系のシートは直後の
        # UserForm_Activate(UFCask:784-790)で表示に戻るので、業務の結果は変わらない(Effects をまとめて後で
        # 実行するここで行うと Activate の後に隠してしまうので行わない)
        log.info("キャスク: 呼び出しを省いた %s %s", c.name, c.args)
        return []

    def _daily_report(self, n: int) -> list[Message]:
        """Worker.Daily_Report(n)。キャスク検査が表示中なので Lname="" → `Daily_` シート(普通は無い → 何もしない)。"""
        name = workers.daily_sheet_name("キャスク", "")
        cells = wiring.sheet_cells(self.store, name)
        if cells is None:
            return []
        book = self.book
        assert book is not None
        env = self.env()
        hist = wiring.sheet_rows(self.store, LOT_HISTORY)
        ctx = workers.DailyContext(exists=True, cells=dict(cells), history_rows=hist[1:] if hist is not None else None,
                                   working_lot=sh.cell_str(book.ws(sh.CASK).get(2, 4)), now=self.now(),
                                   facility=self.facility() or workers.DEFAULT_FACILITY, shifts=env.shift_times(),
                                   mode="キャスク", sub_line=book.cask.line_name, sub_kname=book.cask.kname)
        try:
            writes = workers.daily_report(n, ctx)
        except IndexError as e:
            raise Stop("実行時エラー '9':" + "\r\n\r\n" + "インデックスが有効範囲にありません。", source="Worker.Daily_Report") from e
        if writes:
            wiring.write_cells(self.store, name, writes)
        return []

    def _daily_create(self) -> list[Message]:
        """新規発行の `If Register_Check(…) Then Call Daily_Create`(UFCask:43)。

        Register_Check は UF1 の検査員で見る。Daily_Create は UF1 が隠れているので Lname = UFCask.LineName
        (Worker:721-725)。見出しを書いた後の Daily_Report(1)(3) はキャスク表示中のため `Daily_` を探して何もしない。
        """
        uf1 = self.uf1_state()
        if uf1 is None or not self._registered(uf1):
            return []
        book = self.book
        assert book is not None
        self._daily_header(book.cask.line_name)
        self._daily_report(1)
        self._daily_report(3)
        return []

    # -- 印刷・保存 ---------------------------------------------------
    def _print_folder(self) -> Path:
        return self.store.path.parent / "print"

    def _print(self, job: sh.PrintJob) -> list[Message]:
        """PrintOut の代わり(D8): 帳票の HTML を書き、画面で開く。"""
        book = self.book
        assert book is not None
        ws = book.ws(sh.CASK)
        lot = sh.cell_str(ws.get(2, 4)) or "キャスク"
        folder = self._print_folder()
        stem = "キャスク検査_" + re.sub(r'[\\/:*?"<>|]', "_", lot)
        with trace.timed("db.save", "キャスク検査の帳票(HTML)", path=str(folder), table=stem):
            folder.mkdir(parents=True, exist_ok=True)
            p = folder / (stem + ".html")
            p.write_text(render_html(print_model(job, ws), stem), encoding="utf-8")
            trace.note_used(p, "save")
        self.report = {"html": str(p), "name": p.name}
        return []

    def _save(self, plan: sh.SavePlan) -> list[Message]:
        """SavePlan の実行: 1 つの xlsx(シートを並べる)と、シートごとの見せる用 HTML(Windows で Edge があれば PDF)。"""
        edge = paths.to_local(paths.get("edge"))
        edge_s = str(edge) if edge.is_file() else None
        files, msgs = self._run_save_plan(plan, lambda folder, stem: self._write_sheets(plan, folder, stem, edge_s))
        self.saved = {"name": plan.name, "files": files}
        return msgs

    def _write_sheets(self, plan: sh.SavePlan, folder: Path, stem: str, edge: Optional[str]) -> list[Path]:
        out: list[Path] = []
        folder.mkdir(parents=True, exist_ok=True)
        # VBA と同じく 1 つの xlsx にシートを並べる(値だけ・パスワードなし。2026-11-07 のご判断)
        from .xlsx_writer import grid_rows, write_xlsx
        book = [(name, grid_rows((*sh.a1(k), v) for k, v in snap.items())) for name, snap in plan.sheets.items()]
        x = write_xlsx(folder / (stem + ".xlsx"), book)
        trace.note_used(x, "save")
        out.append(x)
        for name, snap in plan.sheets.items():
            base = f"{stem}_{name}"
            model = snapshot_model(snap, kind=name, sheet_name=name,
                                   page={"orientation": "landscape", "fit_wide": 1})
            h = folder / (base + ".html")
            h.write_text(render_html(model, base), encoding="utf-8")
            out.append(h)
            if edge:
                import os
                if os.name == "nt":
                    out.append(report_html.make_pdf(h, folder / (base + ".pdf"), edge))
        return out

    def _history(self, row: dict[int, Any]) -> list[Message]:
        """`保存用` の最終行の次に履歴行(UFCask:148-162)。"""
        values = [row.get(i + 1) for i in range(len(HOZON_HEADERS))]
        with trace.timed("db.save", "保存用に履歴行(キャスク)", path=str(self.store.path), table=HOZON):
            self.store.sheet(HOZON).append(values, "O")
            trace.note_used(self.store.path, "save")
        return []

    # ------------------------------------------------------------------
    # 操作
    # ------------------------------------------------------------------
    def state(self) -> Result:
        with self._lock:
            self._ensure_book()
            return Result(self.view())

    def open(self) -> Result:
        return self._run(self._open)

    def _open(self) -> Optional[sh.Effects]:
        """UF1.KYB_Click (UF1:3230-3245) → UFCask.Show。表示中なら何もしない(vbModeless のフォームはそのまま)。"""
        book = self.book
        assert book is not None
        if book.cask.visible:
            return None
        uf1 = self.uf1_state()
        env = self.env()
        env.messages = []
        if uf1 is None:
            book.fx.messages.append(Message(NEED_WORKER, ERROR, "警告", source="UF1:3233"))
            return None
        if self._registered(uf1):
            wiring._msg_error(env, uf1, NEED_WORKER, 2)
            book.fx.messages.extend(env.messages)
            return None
        for name in (sh.CASK, WORK_SAVE, "キャスク保存"):      # HiddenSheets(UF1:3239)
            if name in book.sheets or name in book.extra:
                book.ws(name).visible = False
        self._load_hinsho()
        now = self.now()
        fac = self.facility() or workers.DEFAULT_FACILITY
        to_time = workers.yea_dayr_check(workers.hhmm(now), now=now, facility=fac, shifts=env.shift_times())
        to_tyo = workers.time_check(workers.hhmmss(now), facility=fac, shifts=env.shift_times())
        book.cask = sh.CaskForm()
        return cask.open_form(book, uf1.line_name, to_time=to_time, to_tyo=to_tyo)

    def _load_hinsho(self) -> None:
        """`品証検証` シート(1000mm 校正値。ProofreadingSearch が読む)。"""
        book = self.book
        assert book is not None
        cells = wiring.sheet_cells(self.store, "品証検証")
        if cells is None:
            book.extra.pop("品証検証", None)
            return
        q = sh.Sheet("品証検証")
        for (r, c), v in cells.items():
            q.set(r, c, v)
        book.extra["品証検証"] = q

    def _need_open(self) -> None:
        if self.book is None or not self.book.cask.visible:
            raise Stop("キャスク検査の画面を開いてください(検査員の登録が要ります)", level=WARN, source="UF1.KYB_Click")

    def set_lot(self, lot: str) -> Result:
        return self._run(self._set_lot, lot)

    def _set_lot(self, lot: str) -> sh.Effects:
        self._need_open()
        book = self.book
        book.cask.lot_box = _alnum_upper(lot)
        return cask.lot_box_change(book)

    def set_field(self, name: str, value: str) -> Result:
        return self._run(self._set_field, name, value)

    def _set_field(self, name: str, value: str) -> None:
        self._need_open()
        if name not in FIELDS:
            raise Stop("その欄はありません: " + str(name), source="cask_session")
        v = str(value or "")
        if FIELDS[name] == "alnum":
            v = _alnum_upper(v)
        setattr(self.book.cask, name, v)

    def zero(self, text: str) -> Result:
        return self._run(self._zero, text)

    def _zero(self, text: str) -> sh.Effects:
        self._need_open()
        return cask.start_tb_after_update(self.book, str(text or ""))

    def receive(self, text: str) -> Result:
        return self._run(self._receive, text)

    def _receive(self, text: str) -> Optional[sh.Effects]:
        """アクティブセルへの入力(測定器のキーボードウェッジ・手入力)。保護中のロックされたセルは Excel が断る。"""
        self._need_open()
        book = self.book
        book.activate_sheet(sh.CASK)
        r, c = book.active_cell
        if not book.user_editable(r, c):
            book.fx.messages.append(Message(PROTECTED_TEXT, WARN, title="Microsoft Excel", source="Excel(シート保護)"))
            return None
        return book.user_edit(r, c, str(text or ""))

    def select(self, row: int, col: int) -> Result:
        return self._run(self._select, row, col)

    def _select(self, row: int, col: int) -> sh.Effects:
        self._need_open()
        try:
            r, c = int(row), int(col)
        except (TypeError, ValueError):
            r = c = 0
        if not (1 <= r <= LAST_ROW and 1 <= c <= 28):
            raise Stop("そのセルはありません", source="cask_session")
        return self.book.user_select(r, c, sh.CASK)

    def transfer(self) -> Result:
        return self._run(self._transfer)

    def _transfer(self) -> sh.Effects:
        self._need_open()
        self.ainput = False
        return cask.transfer(self.book, now=self.now())

    def print_out(self, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._print_out, answers or {})

    def _print_out(self, answers: dict[str, bool]) -> Union[sh.Effects, Question]:
        self._need_open()
        return cask.print_and_save(self.book, answers)

    def save(self) -> Result:
        return self._run(self._save_only)

    def _save_only(self) -> sh.Effects:
        self._need_open()
        return cask.save_only(self.book)

    def new_sheet(self, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._new_sheet, answers or {})

    def _new_sheet(self, answers: dict[str, bool]) -> Union[sh.Effects, Question]:
        self._need_open()
        out = cask.new_sheet(self.book, answers)
        if not isinstance(out, Question) and answers.get("cask_new"):
            self.flash = None
            self.report = None
            self.saved = None
            self.ainput = False
        return out

    def cancel(self) -> Result:
        return self._run(self._cancel)

    def _cancel(self) -> sh.Effects:
        self._need_open()
        return cask.cancel_input(self.book)

    def toggle_inspection(self) -> Result:
        return self._run(self._toggle_inspection)

    def _toggle_inspection(self) -> sh.Effects:
        self._need_open()
        return cask.toggle_inspection(self.book)

    def resume(self) -> Result:
        return self._run(self._resume)

    def _resume(self) -> sh.Effects:
        self._need_open()
        return cask.enable_events(self.book)

    def proofread(self) -> Result:
        return self._run(self._proofread)

    def _proofread(self) -> None:
        """CommandButton15_Click (UFCask:99-101) 校正読み込み。"""
        self._need_open()
        self._load_hinsho()
        env = self.env()
        now = self.now()
        fac = self.facility() or workers.DEFAULT_FACILITY
        cask.proofreading_search(self.book, workers.yea_dayr_check(workers.hhmm(now), now=now, facility=fac,
                                                                   shifts=env.shift_times()),
                                 workers.time_check(workers.hhmmss(now), facility=fac, shifts=env.shift_times()))

    def close(self) -> Result:
        return self._run(self._close)

    def _close(self) -> sh.Effects:
        """UserForm_QueryClose (UFCask:875-895)。フォームを閉じる(Unload なので入力欄の値は消える)。"""
        book = self.book
        fx = cask.query_close(book)
        book.cask = sh.CaskForm()
        self.flash = None
        return fx

    # ------------------------------------------------------------------
    # 画面の値
    # ------------------------------------------------------------------
    def view(self) -> dict[str, Any]:
        book = self.book
        assert book is not None
        ws = book.ws(sh.CASK)
        f = book.cask
        hidden_hi = ws.col_width.get(8, 8.88) < 1
        columns = []
        for c, j in COLUMNS:
            columns.append({"col": c, "judge": j, "label": cell_text(ws.get(6, c)) or ("備考" if c == 23 else ""),
                            "hidden": hidden_hi and c == 8, "group": _group(c)})
        ar, ac = book.active_cell if book.active_name == sh.CASK else (0, 0)
        last = FIRST_ROW
        for (r, c) in list(ws.values) + list(ws.formulas):
            if FIRST_ROW <= r <= LAST_ROW and 2 <= c <= 28:
                last = max(last, r)
        until = min(LAST_ROW, max(last, ar if FIRST_ROW <= ar <= LAST_ROW else 0, FIRST_ROW + 19) + 15)
        rows = []
        for r in range(FIRST_ROW, until + 1):
            cells = []
            for c, j in COLUMNS:
                v = ws.get(r, c)
                mark = cell_text(ws.get(r, j)) if j else ""
                flags = []
                st = ""
                if j:
                    st = "ok" if mark == "○" else ("ng" if mark == "×" else "")
                elif c == 21:
                    st = "ok" if cell_text(v) == "○" else ("ng" if cell_text(v) == "×" else "")
                if not st:
                    st = _COLOR_STATE.get(ws.interior.get((r, c)), "")
                    if st == "next":
                        st = ""
                if st:
                    flags.append(st)
                if (r, c) in ws.locked:
                    flags.append("lock")
                if (r, c) in ws.underline:
                    flags.append("u")
                if (r, c) == (ar, ac) or (j and (r, j) == (ar, ac)):
                    flags.append("sel")
                cells.append([cell_text(v, ws.number_format.get((r, c), "General")), mark, " ".join(flags)])
            rows.append([r, cell_text(ws.get(r, 1)), cells])
        nxt = None
        if 1 <= ar and 1 <= ac:
            nxt = {"row": ar, "col": ac, "no": cell_text(ws.get(ar, 1)) if ar >= FIRST_ROW else "",
                   "label": cell_text(ws.get(6, ac)) if ar >= FIRST_ROW
                   else HEADER_CELLS.get((ar, ac), f"{sh.col_letter(ac)}{ar}"),
                   "locked": (ar, ac) in ws.locked}

        def num(r: int, c: int, pat: str) -> str:
            v = ws.get(r, c)
            if v is None or v == "":
                return ""
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                return vbx.fmt(v, pat)
            return cell_text(v)

        tol = []
        for row_no, name, pat, size_c in ((2, "厚", "0.000", 2), (3, "巾", "0.0", 4), (4, "丈", "0.0", 6)):
            tol.append({"name": name, "size": num(3, size_c, pat), "size_cell": [3, size_c],
                        "minus": num(row_no, 13, pat), "minus_cell": [row_no, 13],
                        "plus": num(row_no, 15, pat), "plus_cell": [row_no, 15],
                        "lo": num(row_no, 17, pat), "hi": num(row_no, 19, pat)})
        return {
            "visible": f.visible,
            "events": book.events_enabled,
            "form": {k: getattr(f, k) for k in ("line_name", "lot_box", "ken_no", "kname", "cno", "mno", "seri_no",
                                                "min_text", "max_text", "tai", "micro", "convex", "convex_1000",
                                                "l1", "l2_visible", "l3_visible", "count_st", "inspection", "tag",
                                                "hflag")} | {"labels": list(f.labels)},
            "header": {"lot": cell_text(ws.get(2, 4)), "atu": num(3, 2, "0.000"), "hab": num(3, 4, "0.0"),
                       "tak": num(3, 6, "0.0"), "start": sh.fmt_date(ws.get(1, 2), "yyyy/mm/dd h:mm")
                       if isinstance(ws.get(1, 2), datetime) else cell_text(ws.get(1, 2)), "tolerance": tol,
                       "maxmin": [{"name": n, "max": num(3, c, p), "min": num(4, c, p)}
                                  for n, c, p in (("厚", 22, "0.000"), ("巾", 23, "0.0"), ("丈", 24, "0.0"))]},
            "columns": columns,
            "rows": rows,
            "total_rows": LAST_ROW - FIRST_ROW + 1,
            "next": nxt,
            "flash": self.flash,
            "report": self.report,
            "saved": self.saved,
            "ainput": self.ainput,
            "focus": "zero" if f.visible and (f.micro == "" or f.convex == "") else "meas",
            "pack": self.pack.to_dict() if self.pack is not None else None,
        }


def _group(c: int) -> str:
    if c in (2, 4, 6, 8):
        return "厚"
    if c in (11, 13):
        return "巾"
    if c in (15, 17):
        return "丈"
    if c in (19, 20, 21, 22, 23):
        return "対角"
    return "その他"


# ------------------------------------------------------------------
_session: Optional[CaskSession] = None
_slock = threading.Lock()


def get() -> CaskSession:
    global _session
    with _slock:
        if _session is None:
            _session = CaskSession()
        return _session


def reset_for_test(s: Optional[CaskSession] = None) -> CaskSession:
    global _session
    with _slock:
        _session = s or CaskSession()
        return _session


__all__ = ["CaskSession", "get", "reset_for_test", "Result"]
