"""全ヵ所検査(一般 / 5N / 三菱受託)と H494 検査の画面のセッション(VBA の UFMitsubishi + 全ヵ所検査 / H494検査 シート)

画面(app/routes/zenkasho.py)はここだけを呼ぶ。判断はすべて部品(`modes/zenkasho.py`・`modes/h494.py`・
`modes/sheet.py`)が行い、ここは「部品を呼ぶ順番」と「部品が返した Effects の後始末」だけを受け持つ。

**状態**: ブック(`sheet.Book`: 作業シート 2 枚と UFMitsubishi の状態)をプロセスに 1 つ持つ。
操作のたびに手元DB(通常検査と同じ `store.LocalStore()`)の `全ヵ所作業中保存` / `H494作業中保存` に
シートの中身(値・数式・塗り)を書き、アプリを開き直したらそこから読み戻して `OpenMitubisi` を呼ぶ
(VBA でブックを開き直してから「全数」「H494」を押したのと同じ。フォームだけの値は消え、シートから戻る)。

**操作と VBA のイベント**:
    start()                    UF1 の「全数」/「H494」→ OpenMitubisi + UserForm_Activate(開いていなければ)
    choose(mode)               作業の切替。一般 = 一般_Click、5N = FiveOP_Click、三菱 = Choice へ「三菱」(Choice_Change)、
                               H494 = 閉じて OpenMitubisi("H494")(VBA も P_H494 ページへの切替)
    select(r, c)               セルを選ぶ(Worksheet_SelectionChange)
    receive(text)              選んでいるセルへの入力(測定器の伝送文字列 DT…/#L…、手入力)→ Worksheet_Change。
                               Change が選択を動かさなかったときは Excel の Enter と同じく 1 つ下へ(ENTER_MOVES_DOWN)
    erase()                    選んでいるセルを消す(Delete キー。Change は走る・選択は動かない)
    set_form(name, value)      フォームの入力欄(LotBox・KenNo・Kname・MINText・検査法・KenHBox1〜7・YamaLB、H494_*)
    save_height()              MLP_Click(検入高さを AD1:AD3 へ・重量のゲージ要否)
    weight()                   Frame4 などのクリック(重量のゲージ要否の知らせ)
    transfer()                 CommandButton3 / CommandButton19「決定」(最大最少 → Aインプット)
    print(which, answers)      CommandButton7(一般)/ 15(5N)/ 2(三菱)/ 23(H494)
    new_sheet(answers)         CommandButton1 / ws_out「ｼｰﾄ発行」
    resume()                   CommandButton16 / 21「処理再開」(EnableEvents = True)
    next_empty()               SB_SpinDown(MoveTg: 厚の最初の空き欄へ)
    pack_select(i)             包装仕様書の窓(UF_PackNo)の引当の行をクリック(LabelCls → JPGLoad)
    pack_close(confirmed)      UF_PackNo の「閉じる」(CommandButton2。確認の印が要る)

**Effects の後始末**: 音(音楽合格 → chime2.wav など)・ログ(WriteLog)・UFOK/UFNG(画面の大きな ○/×)・
印刷(PrintJob → 帳票の HTML。D8 でブラウザから印刷)・保存(SavePlan → D7 で xlsx の代わりに HTML/PDF + CSV)・
`保存用` への履歴行。他モジュールの呼び出し(`Effects.calls`)は、部品が積んだその場で処理する
(DrawListBoxContent・DistArray/FormReflectDist・LogOut など。VBA と同じ順番で効くように)。
つないでいない呼び出しは view の `pending` に名前を出す。

**つないだ呼び出し**(部品の `Effects.calls` → ここ):
    HistorySave                 SaveLoad.HistorySave → `lot.history_save`(Lot情報履歴へ。通常検査・キャスクと同じ)
    GetLotInfo/DbReturnItem     AccHIki.GetLotInfo(`lot.get_lot_info`。無ければ ComboBox1_Change を抜ける)→
                                QP.DbReturnItem(`quality.db_return_item`。表示中はシートの O1/M1/AC6 へ)
    DistArray/FormReflectDist   表示中は 全ヵ所検査 の M5/O5。開く途中は UF1 の欄にしか書かないので書かない(pending)
    PackingNoCheck/UFPack_Open  `packing`(包装仕様書の画像と引当の一覧。モーダル。`pack_select` / `pack_close`)
    AddWS/Pesute・不良要因印刷_B1・不良要因印刷範囲・不良要因の PrintJob
                                手元DB の `不良要因印刷_<Line>`(通常検査と同じシート)→ `report_model` → 帳票の HTML
    Register_Check/Daily_Create Worker.Register_Check(UF1 の検査員)→ Daily_Create(`Daily_<UFMitsubishi.LineName>`)
    Daily_Report                Worker.Daily_Report(全ヵ所・H494 表示中は `Daily_`)
    UF対角(フォーム)           2 秒の演出。UFOK/UFNG を閉じる → 画面の大きな表示を「対角」に
VBA の実行時エラーは、ComboBox1_Change の中なら部品のエラー処理(「DB読み込みエラー:」)へ渡すため `sh.VbaError` にする。
"""
from __future__ import annotations

import csv
import hashlib
import json
import os
import re
import shutil
import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional

from .. import trace
from . import checks, datasource, judge, legacy_log, packing, paths, quality, report_html, report_model, terminal, vbx, \
    workers
from . import lot as lotmod
from .flow import ERROR, INFO, WARN, Message, Question, Stop
from .modes import h494 as hm
from .modes import sheet as sh
from .modes import zenkasho as zk
from .modes.sheet import H494, ZEN
from .report_model import ReportModel
from .session import Result
from .store import HOZON, LOT_HISTORY, LocalStore

#: 作業シート → 手元DB のシート名(VBA の作業中保存のシート名)
SAVE_SHEETS = {ZEN: "全ヵ所作業中保存", H494: "H494作業中保存"}
#: 画面の状態(ページ・選んでいるセル・D2 の赤字)を JSON で 1 セルに持つシート
META_SHEET = "全ヵ所作業中保存_画面"
#: 三菱受託の未回答の確認(docs/04 W2)
W2_NOTICE = "三菱受託は確認待ちの項目があります(docs/04 W2)"
#: UF1.ZEB_Click / H494B_Click の Register_Check が真のとき(UF1:4873 / 606)
NEED_WORKER = "検査員登録をして下さい"
#: Excel の既定(Enter で下へ)。Change が選択を動かさなかったとき(ラテラル・歪・外観など)に効く
ENTER_MOVES_DOWN = True

SOUNDS = {"音楽合格": judge.SOUND_OK, "音楽不合格": judge.SOUND_NG, "音楽対角": judge.SOUND_DIAG}
YAMA_LIST = ("1", "2", "3", "4", "5", "6", "枚数")          # YamaLB.List (UFMit:2305)
METHODS = ("L1", "L2", "M6")                               # Inspection1〜3 / op1〜3 の Caption
MODES = (("一般", "一般"), ("5N", "5N"), ("三菱", "三菱受託"), ("H494", "H494"))

#: 表の列(列番号, 見出し, 種類, まとまり)。種類 v=測定値 j=判定 c=計算 t=文字 l=原紙の固定文字
ZEN_COLS: tuple[tuple[int, str, str, str], ...] = (
    (2, "厚1", "v", "厚"), (3, "判", "j", "厚"), (4, "厚2", "v", "厚"), (5, "判", "j", "厚"),
    (6, "厚3", "v", "厚"), (7, "判", "j", "厚"), (8, "厚4", "v", "厚"), (9, "判", "j", "厚"),
    (11, "幅1", "v", "幅"), (12, "判", "j", "幅"), (13, "幅2", "v", "幅"), (14, "判", "j", "幅"),
    (15, "丈1", "v", "丈"), (16, "判", "j", "丈"), (17, "丈2", "v", "丈"), (18, "判", "j", "丈"),
    (19, "対角1", "v", "対角"), (20, "対角2", "v", "対角"), (21, "差", "c", "対角"), (22, "判", "j", "対角"),
    (23, "ﾗﾃﾗﾙ", "v", "ラテラル"), (24, "判", "j", "ラテラル"),
    (25, "幅歪", "v", "歪"), (26, "判", "j", "歪"), (27, "丈歪", "v", "歪"), (28, "判", "j", "歪"),
    (29, "ｹﾞｰｼﾞ", "v", "ゲージ"), (30, "外観", "t", "外観"),
)
MITSU_COLS: tuple[tuple[int, str, str, str], ...] = (
    (31, "元板NO", "t", "三菱"), (33, "切断ﾌﾞﾛｯｸNO", "t", "三菱"), (34, "ﾌﾞﾛｯｸ", "l", "三菱"),
)
H494_COLS: tuple[tuple[int, str, str, str], ...] = (
    (2, "厚1", "v", "厚"), (3, "判", "j", "厚"), (4, "厚2", "v", "厚"), (5, "判", "j", "厚"),
    (6, "厚3", "v", "厚"), (7, "判", "j", "厚"), (8, "厚4", "v", "厚"), (9, "判", "j", "厚"),
    (11, "幅1", "v", "幅"), (12, "判", "j", "幅"), (13, "幅2", "v", "幅"), (14, "判", "j", "幅"),
    (15, "幅3", "v", "幅"), (16, "判", "j", "幅"),
    (17, "丈1", "v", "丈"), (18, "判", "j", "丈"), (19, "丈2", "v", "丈"), (20, "判", "j", "丈"),
    (21, "対角1", "v", "対角"), (22, "対角2", "v", "対角"), (23, "判", "j", "対角"), (24, "差", "c", "対角"),
    (25, "ﾗﾃﾗﾙ", "v", "ラテラル"), (26, "定盤", "v", "歪"), (27, "ｹﾞｰｼﾞ", "v", "歪"), (28, "外観", "t", "外観"),
)
SHEET_ROWS = {ZEN: (7, 76), H494: (7, 16)}
DEFECT_PREFIX = "不良要因印刷_"                            # 不良要因印刷 & "_" & Lname
TYPE_MISMATCH = "型が一致しません。"
SHEET_COLS = {ZEN: 35, H494: 28}


def _rgb_int(rgb: Any) -> Optional[int]:
    """VBA の RGB(r, g, b) = r + g*256 + b*65536。"""
    if rgb is None:
        return None
    r, g, b = rgb
    return int(r) + int(g) * 256 + int(b) * 65536


def _int_rgb(n: Any) -> Optional[tuple[int, int, int]]:
    if n is None:
        return None
    n = int(n)
    return (n % 256, (n // 256) % 256, (n // 65536) % 256)


def _cls(rgb: Any) -> str:
    """塗り → 画面の色の種類。緑(ColorIndex 10)= ok、赤(ColorIndex 3 / RGB(200,60,50))= ng、薄青 = sel(入力先)。"""
    if rgb is None or rgb == sh.WHITE:
        return ""
    if rgb == sh.OK_GREEN:
        return "ok"
    if rgb in ((255, 0, 0), sh.NG_RED):
        return "ng"
    if rgb == sh.LIGHT_BLUE:
        return "sel"
    return "warn"


def _db_value(v: Any) -> Any:
    """手元DB に入れる形(日時は Excel の文字、エラー値は文字)。"""
    if isinstance(v, datetime):
        return sh.cstr_date(v)
    if isinstance(v, sh.ExcelError):
        return str(v)
    return v


def _from_db(v: Any) -> Any:
    if isinstance(v, str):
        for f in sh._DATE_FORMATS:
            try:
                return datetime.strptime(v.strip(), f)
            except ValueError:
                pass
    return v


def _text(ws: sh.Sheet, r: int, c: int) -> str:
    """セルの見た目の文字(表示形式 "0.00_);[赤](0.00)" などは最初の区切りの形で)。"""
    v = ws.get(r, c)
    if v is None:
        return ""
    if isinstance(v, sh.ExcelError):
        return str(v)
    if isinstance(v, datetime):
        return sh.fmt_date(v, "yyyy/mm/dd h:mm")
    nf = ws.number_format.get((r, c))
    if nf and isinstance(v, (int, float)) and not isinstance(v, bool):
        pat = re.split(r"[_;]", nf)[0]
        if pat:
            return vbx.fmt(v, pat)
    try:
        return sh.cell_str(v)
    except sh.VbaError:
        return str(v)


def _fmt_num(v: Any, pat: str) -> str:
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return vbx.fmt(v, pat)
    return "" if v is None else str(v)


def _safe_name(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|\s]+', "_", s).strip("_") or "sheet"


class _Calls(list):
    """`Effects.calls`。部品が積んだその場でセッションに知らせる(VBA の呼び出しと同じ順番で効かせる)。"""

    def __init__(self, hook: Callable[[sh.Call], None]) -> None:
        super().__init__()
        self._hook = hook

    def append(self, item: Any) -> None:            # type: ignore[override]
        super().append(item)
        self._hook(item)


class _Book(sh.Book):
    """`begin()` のたびに前の Effects をセッションに渡すブック(1 操作で部品を何度呼んでも知らせを落とさない)。"""

    def __init__(self, owner: "ZenkashoSession", **kw: Any) -> None:
        self._owner = owner
        super().__init__(**kw)
        self.fx = self._new_fx()

    def _new_fx(self) -> sh.Effects:
        fx = sh.Effects()
        fx.calls = _Calls(self._owner._on_call)
        return fx

    def begin(self) -> sh.Effects:
        self._owner._harvest(self.fx)
        self.fx = self._new_fx()
        return self.fx


class ZenkashoSession:
    """プロセスに 1 つの全ヵ所・H494 のセッション。操作は `_lock` で 1 つずつ通す。

    試験では `store`(手元DB)・`now`(時計)・`kensa_state`(通常検査の状態。読むだけ)・`write_logs` を差し替える。
    """

    def __init__(self, store: Optional[LocalStore] = None, *, now: Callable[[], datetime] = datetime.now,
                 kensa_state: Optional[Callable[[], Any]] = None, write_logs: bool = True) -> None:
        self._lock = threading.RLock()
        self._store = store
        self.now = now
        self.write_logs = write_logs
        self._kensa_state = kensa_state
        self.book: Optional[_Book] = None
        self.page = "P1"
        self._acc: list[sh.Effects] = []
        self._msgs: list[Message] = []
        self._digest = ""
        self._sel: dict[str, tuple[int, int]] = {}
        # 直近の操作の結果(画面に出すもの)
        self.flash = ""
        self.reports: list[dict[str, str]] = []
        self.saved: list[str] = []
        self.ainput: Optional[dict[str, str]] = None
        self.pending: list[str] = []
        self.errors: list[str] = []
        self.pack: Optional[packing.PackForm] = None           # UF_PackNo(閉じるまで残る。モーダル)
        self.diag = False                                      # UF対角(直近の操作で出た)
        self._lot_rows: Optional[list[list[Any]]] = None       # GetLotInfo の結果(DbReturnItem が使う)
        self._defect_had: Optional[str] = None                 # 操作の前にあった不良要因シート
        self._env: Optional[Any] = None

    # ------------------------------------------------------------------
    # 土台
    # ------------------------------------------------------------------
    @property
    def store(self) -> LocalStore:
        if self._store is None:
            self._store = LocalStore()
        return self._store

    def _kstate(self) -> Any:
        """通常検査の状態(読むだけ。UF1 のライン・検査員)。"""
        try:
            if self._kensa_state is not None:
                return self._kensa_state()
            from . import session as ks
            return ks.get().state
        except Exception:                                  # noqa: BLE001 - 読めなければ通常検査の値は使わない
            return None

    def _sheet(self) -> str:
        return ZEN if self.page == "P1" else H494

    def _exists(self, p: str) -> bool:
        """保存の連番の判定(VBA は .xlsx を探す。D7 で xlsx の代わりに .html/.pdf/.csv を書くので、それも見る)。"""
        lp = legacy_log.local_path(p)
        if lp.exists():
            return True
        if p.lower().endswith(".xlsx"):
            stem = lp.name[:-5]
            return any((lp.parent / (stem + ext)).exists() for ext in (".html", ".pdf"))
        return False

    def _new_book(self) -> _Book:
        others = [HOZON] if self.store.exists(HOZON) else []
        b = _Book(self, now=self.now, other_sheets=others, exists=self._exists)
        return b

    def _ensure(self) -> _Book:
        if self.book is None:
            self.book = self._new_book()
            self._restore()
        return self.book

    def _line_name(self) -> str:
        st = self._kstate()
        name = getattr(st, "line_name", "") if st is not None else ""
        return name or terminal.facility()

    def _worker_text(self) -> Optional[str]:
        """OpenMitubisi の `Split(UF1.T?, "_")`(WorkeDayMax の直)。直が無い・欄が空なら None(検査員欄を変えない)。"""
        st = self._kstate()
        if st is None:
            return None
        try:
            from . import wiring
            day = wiring.worker_day_max(self.store, st.lname)
            reg = wiring.registry(st)
        except Exception:                                  # noqa: BLE001
            return None
        t = {"1": reg.t1, "2": reg.t2, "3": reg.t3}.get(day)
        return t if t else None

    def _lot_kw(self) -> dict[str, Any]:
        """ロット読込(ComboBox1_Change)に渡す Judg の公差表の関数。"""
        m = self._msgs
        return {"diagonal": lambda a, b: judge.diagonal(a, b, messages=m),
                "re_diagonal": lambda a, b: judge.re_diagonal(a, b, messages=m),
                "lateralbow_j": lambda a, b, c: judge.lateralbow_j(a, b, c, messages=m),
                "re_lateralbow": lambda a, b, c: judge.re_lateralbow(a, b, c, messages=m),
                "lot_info_found": self._get_lot_info}

    # ------------------------------------------------------------------
    # Effects
    # ------------------------------------------------------------------
    def _harvest(self, fx: sh.Effects) -> None:
        if not any(x is fx for x in self._acc):
            self._acc.append(fx)

    def _on_call(self, call: sh.Call) -> None:
        """部品が `Effects.calls` に積んだ呼び出しを、その場で処理する。"""
        try:
            if call.name == "DrawListBoxContent":
                self._draw_list_box(float(call.args[0]))
            elif call.name == "DistArray/FormReflectDist":
                self._dist_array(str(call.args[0]), str(call.args[1]))
            elif call.name == "LogOut":
                self._log_out(str(call.args[0]), str(call.args[2]) if len(call.args) > 2 else "")
            elif call.name == "Ainput_Startup":
                self.ainput = {"kind": str(call.args[0]), "method": str(call.args[1]), "url": "/ainput"}
            elif call.name == "Daily_Report":
                self._daily_report(int(call.args[0]))
            elif call.name == "HistorySave":
                self._history_save(list(call.args[0]))
            elif call.name == "GetLotInfo/DbReturnItem":
                self._db_return_items()
            elif call.name == "PackingNoCheck/UFPack_Open":
                self._packing(str(call.args[0]))
            elif call.name == "AddWS/Pesute":
                self._add_ws_pesute(str(call.args[0]), str(call.args[1]))
            elif call.name == "不良要因印刷_B1":
                self._defect_b1(str(call.args[0]), call.args[1])
            elif call.name == "不良要因印刷範囲":
                self._defect_range(str(call.args[0]))
            elif call.name == "Register_Check/Daily_Create":
                self._daily_create()
            elif call.name == "作業毎保存":
                pass                                       # 毎操作のあと手元DB へ書いている(_persist)
            else:
                self._pending(call.name)
        except (datasource.DbMissing, vbx.TypeMismatch, quality.VbaRuntimeError, IndexError, OSError) as e:
            self._msgs.append(Message(f"{call.name}: {e}", ERROR, source=call.name))

    def _pending(self, name: str) -> None:
        if name not in self.pending:
            self.pending.append(name)

    def _draw_list_box(self, thick: float) -> None:
        """OutTolerance.DrawListBoxContent (O:1747-1840) の UFMitsubishi 分: 高さ計算式を出し直す(CalculateSheets)。"""
        f = self.book.mit
        idx = f.yama_index
        if not isinstance(idx, int) or not 0 <= idx < len(YAMA_LIST):
            return                                         # YamaLB.Value = "" → Exit Sub
        f.takasa_formula = ""
        res = checks.calculate_sheets([f.kenh[i] for i in range(1, 7)], thick, mitsubishi=True, box7=f.kenh[7])
        if res.formula is not None:
            f.takasa_formula = res.formula

    def _dist_array(self, lot_no: str, ken_no: str) -> None:
        """ComboBox1_Change の DistArray → FormReflectDist (UFMit:846-850 / QP:1986-2055)。

        UFMitsubishi が表示中なら 全ヵ所検査 の M5/O5(幅歪・丈歪の公差)へ。
        開く途中(OpenMitubisi の LotBox 代入。UFMitsubishi.Visible = False)は、FormReflectDist の最初の分岐
        (`UF品質基準.Visible = False And UFMitsubishi.Visible = False`)で **UF1 の MG・HaHiz・TaHiz・Naka・Mimi・
        HizCome・DistortionFrame にだけ**書く(シートには書かない。UF1 はこの後 Hide される)。通常検査の状態は
        書き換えない約束なので書かず、DistArray の知らせだけ出して pending に理由を残す。
        """
        b = self.book
        qp = quality.dist_array(lot_no, ken_no, quality.hiki_arr_create(lot_no), quality.read_lot(lot_no))
        self._msgs.extend(qp.messages)
        if not qp.kousa:
            return
        if not b.mit.visible:
            self._pending("FormReflectDist(開く途中は UF1 の歪公差の欄だけに書く。通常検査の状態は書き換えない)")
            return
        r = quality.form_reflect_dist(qp.kousa, quality_visible=False, mitsubishi_visible=True, uf1_visible=False,
                                      zenkasho_sheet=ZEN in b.sheets)
        if r.target == "シート":
            ws = b.ws(r.sheet)
            for (rr, cc), v in r.cells.items():
                ws.set(rr, cc, sh.to_cell(v))              # EnableEvents は False の間(UFMit:850)

    # ------------------------------------------------------------------
    # ロット読込の続き(HistorySave・GetLotInfo・DbReturnItem・包装仕様書)
    # ------------------------------------------------------------------
    def env(self) -> Any:
        from . import wiring
        if self._env is None:
            self._env = wiring.Env(store=self.store, now=self.now, facility=terminal.facility,
                                   write_logs=self.write_logs)
        return self._env

    @staticmethod
    def _vba_error(e: Exception, number: int = 13, desc: str = TYPE_MISMATCH) -> sh.VbaError:
        """部品のエラー処理(ComboBox1_Change の ErrHandler_MitCombo など)へ渡す形にする。"""
        if isinstance(e, packing.PackingError):
            return sh.VbaError(e.number, e.description, e.where)
        if isinstance(e, quality.VbaRuntimeError):
            return sh.VbaError(getattr(e, "number", number), str(e), "")
        if isinstance(e, datasource.DbMissing):
            return sh.VbaError(0, str(e), "datasource")
        return sh.VbaError(number, desc, str(e))

    def _history_save(self, past: list[Any]) -> None:
        """ComboBox1_Change の `Call HistorySave(PastHiki)`(UFMit:804)。Lot情報履歴 に引当ごとの行。

        途中の実行時エラーは ComboBox1_Change のエラー処理(「DB読み込みエラー:」)に落ちる(行の削除は済んだまま)。
        g_CacheLotNo などの控え(UF1 のロットと同じときだけ効く)は通常検査の状態なので書き換えない。
        """
        b = self.book
        st = self._kstate()
        over12 = False
        if st is not None:
            from . import wiring
            over12 = wiring.RealHooks(self.env())._over12(st)
        henv = lotmod.HistoryEnv(dbflag=getattr(st, "dbflag", "") or "", now=self.now(), over12=over12,
                                 zenkasho_visible=ZEN in b.sheets and b.sheets[ZEN].visible,
                                 uf1_test=getattr(st, "line_option", "") == "Test", mitsubishi_line=b.mit.line_name)
        from . import wiring
        hist = wiring.sheet_rows(self.store, LOT_HISTORY) or []
        with trace.timed("db.save", "Lot情報履歴(HistorySave)", path=str(self.store.path), table=LOT_HISTORY,
                         expected=(sh.VbaError,)) as info:
            try:
                res = lotmod.history_save(past, hist, henv)
            except (vbx.TypeMismatch, lotmod.SubscriptOutOfRange, IndexError, datasource.DbMissing) as e:
                partial = getattr(e, "partial", None)
                if partial is not None:
                    self._apply_history(partial)
                raise self._vba_error(e, 9 if isinstance(e, (lotmod.SubscriptOutOfRange, IndexError)) else 13,
                                      "インデックスが有効範囲にありません。"
                                      if isinstance(e, (lotmod.SubscriptOutOfRange, IndexError)) else TYPE_MISMATCH)
            info["rows"] = self._apply_history(res)
        self._msgs.extend(res.messages)
        if self.write_logs:
            for lg in res.logs:
                legacy_log.write_log(lg.msg, lg.sepl, lg.target_path, now=self.now())

    def _apply_history(self, h: lotmod.HistorySaveResult) -> int:
        """HistorySave の結果を Lot情報履歴 に反映する(行を消してから足す。通常検査・キャスクと同じ)。"""
        lh = self.store.sheet(LOT_HISTORY)
        trace.note_used(self.store.path, "save")
        for r in sorted(h.delete_rows, reverse=True):
            lh.delete_row(r)
        for row in h.new_rows:
            lh.append(row[:26])
        return len(h.new_rows)

    def _get_lot_info(self) -> bool:
        """`TempArry = GetLotInfo("Lot情報履歴", UFMitsubishi.LotBox.Text)`(UFMit:807-809)。Empty なら抜ける(偽)。

        H494 でも LotBox(全ヵ所のロット欄)で引く(VBA どおり)。
        """
        from . import wiring
        try:
            self._lot_rows = lotmod.get_lot_info(wiring.sheet_rows(self.store, LOT_HISTORY), self.book.mit.lot_box)
        except vbx.TypeMismatch as e:
            raise self._vba_error(e) from e
        return self._lot_rows is not None

    def _db_return_items(self) -> None:
        """UFMit:810-845。引当ごとの納入先で DbReturnItem(前回と同じ納入先は飛ばす)。

        用途コード・材質・調質・厚・幅・丈は**最後の行**の値(VBA のループで上書き)。DbReturnItem は
        UFMitsubishi 表示中ならシート(全ヵ所検査 / H494検査)の O1・M1・AC6 に書く。UF1 に書く分(HURIS の空け、
        該当なしのときの Sumisei・Additional・LateralText、表示前の LateralText など)は通常検査の状態なので書かない。
        """
        rows = self._lot_rows or []
        if not rows:
            return
        b, f = self.book, self.book.mit
        st = self._kstate()
        try:
            last = rows[-1]
            youto, zai, tyou = (vbx.cstr(last[i]) if last[i] is not None else "" for i in (2, 3, 4))
            atu, haba, take = (self._single(last[i]) for i in (5, 6, 7))
        except vbx.TypeMismatch as e:
            raise self._vba_error(e) from e
        prev = ""
        lateral_lot = None
        quiet = False                                  # DBdisplayラテラル: IsOver12Hours が真なら MsgBox を出さない
        if st is not None:
            from . import wiring
            quiet = wiring.RealHooks(self.env())._over12(st)
        for r in rows:
            nou = r[0]
            if vbx.compare_typed(nou, "", vbx.VAR, vbx.STR) == 0:
                continue
            nou_s = vbx.cstr(nou)
            tori_s = vbx.cstr(r[1]) if r[1] is not None else ""
            if nou_s == prev:
                continue
            prev = nou_s
            if lateral_lot is None:
                lateral_lot = quality.read_lot(f.lot_box)
            try:
                res = quality.db_return_item(nou_s, tori_s, youto, zai, tyou, atu, haba, take,
                                             additional_caption=getattr(st, "additional", "") or "",
                                             mitsubishi_visible=f.visible, work_flag=f.work_flag,
                                             zenkasho_sheet=ZEN in b.sheets, lateral_lot=lateral_lot,
                                             lateral_quiet=quiet)
            except (quality.VbaRuntimeError, vbx.TypeMismatch, datasource.DbMissing) as e:
                raise self._vba_error(e) from e
            self._msgs.extend(res.messages)
            if res.sheet:
                ws = b.ws(res.sheet)
                for (rr, cc), v in res.cells.items():
                    ws.set(rr, cc, sh.to_cell(v))          # EnableEvents は False の間
            if any(k != "HURIS" or v != "" for k, v in res.uf1.items()) or res.cb_add_atu is not None:
                self._pending("DbReturnItem(UF1 の欄: " + "・".join(k for k in res.uf1 if k != "HURIS")
                              + ")" if res.uf1 else "DbReturnItem(UF1 の CBAdd)")

    @staticmethod
    def _single(v: Any) -> float:
        """`Dim x As Single: x = TempArry(i, n)`(Empty は 0、数にならない文字列は型エラー 13)。"""
        if v is None:
            return 0.0
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return vbx.single(float(v))
        return vbx.single(vbx.cdbl(v))

    def _packing(self, lot_no: str) -> None:
        """`If Not PackingNoCheck(LotNo) = "" Then Call UFPack_Open(PackingNoCheck(LotNo), LotNo)`(UFMit:856-858)。"""
        from . import wiring
        hist = wiring.sheet_rows(self.store, LOT_HISTORY)
        with trace.timed("zenkasho.packing", "包装仕様書の確認(PackingNoCheck)", expected=(sh.VbaError,), lot=lot_no):
            try:
                n = packing.packing_no_check(hist, lot_no, messages=self._msgs)
                if n == "":
                    return
                form = packing.pack_open(n, lot_no, hist, messages=self._msgs)
            except (packing.PackingError, quality.VbaRuntimeError, vbx.TypeMismatch) as e:
                raise self._vba_error(e) from e
        if form is not None:
            self.pack = form

    def pack_select(self, i: Any) -> Result:
        return self._run(self._pack_select, i)

    def _pack_select(self, i: Any) -> None:
        """UF_PackNo の引当の行をクリック(LabelCls → JPGLoad)。"""
        if self.pack is None:
            raise Stop("包装仕様書の画面は開いていません", level=WARN)
        from . import wiring
        try:
            n = int(i)
        except (TypeError, ValueError):
            raise Stop("行の番号が正しくありません", level=WARN) from None
        packing.pack_select(self.pack, n, wiring.sheet_rows(self.store, LOT_HISTORY), messages=self._msgs)

    def pack_close(self, confirmed: Any = False) -> Result:
        return self._run(self._pack_close, bool(confirmed))

    def _pack_close(self, confirmed: bool) -> None:
        """UF_PackNo.CommandButton2_Click(「確認」が無いと閉じない)。"""
        m = packing.pack_close(confirmed)
        if m is not None:
            self._msgs.append(m)
            return
        self.pack = None

    def pack_image(self) -> Optional[Path]:
        """いま出している包装仕様書の画像(使用パスの QualityPath\包装仕様書)。"""
        with self._lock:
            if self.pack is None:
                return None
            return packing.picture_file(self.pack.picture)

    # ------------------------------------------------------------------
    # 不良要因票(手元DB の 不良要因印刷_<Line>。通常検査と同じシート)
    # ------------------------------------------------------------------
    def _defect_name(self) -> str:
        return DEFECT_PREFIX + self.book.mit.line_name

    def _sync_defect(self) -> None:
        """操作の前に、手元DB の不良要因シートの有無と B1 をブックへ映す(部品は `book.has` / B1 で判断する)。"""
        b = self.book
        for k in [k for k in b.extra if k.startswith(DEFECT_PREFIX)]:
            b.extra.pop(k)
        name = self._defect_name()
        self._defect_had = None
        if self.store.exists(name):
            ws = sh.Sheet(name)
            ws.visible = False
            ws.set(1, 2, self.store.sheet(name).get(1, 2))
            b.extra[name] = ws
            self._defect_had = name

    def _defect_after(self) -> None:
        """部品が消した不良要因シート(新規発行の `Sheets(…).Delete`)を手元DB からも消す。"""
        name = self._defect_had
        if name and not self.book.has(name) and self.store.exists(name):
            trace.note_used(self.store.path, "save")
            with trace.timed("db.save", "不良要因シートを消す", path=str(self.store.path), table=name):
                self.store.delete_sheet(name)
        self._defect_had = None

    def _add_ws_pesute(self, template: str, name: str) -> None:
        """補助.AddWS(name) + Print.Pesute(原紙, name)。手元DB のシートを作り直して原紙の値を写す。"""
        trace.note_used(self.store.path, "save")
        with trace.timed("db.save", "不良要因シートを原紙から作る(AddWS/Pesute)", path=str(self.store.path), table=name):
            ws = self.store.add_sheet(name)
            report_model.paste_template_values(ws, template)
        b1 = ws.get(1, 2)
        if name in self.book.extra:
            self.book.extra[name].set(1, 2, b1)
        self._defect_had = None                            # 作り直した(この操作で消したのではない)

    def _defect_b1(self, name: str, lot: Any) -> None:
        """`Sheets("不良要因印刷_" & Lname).Cells(1, 2) = UFMitsubishi.LotBox`(Mit:735)。"""
        if not self.store.exists(name):
            raise sh.VbaError(9, "インデックスが有効範囲にありません。", f"Sheets({name})")
        self.store.sheet(name).write(1, 2, lot, "O")
        trace.note_used(self.store.path, "save")

    def _defect_range(self, kname: str) -> None:
        """Print.不良要因印刷範囲(Kname) (Print:1278-1335)。**LotN を渡さない形**なので UF1 の Lname・Lot を使う。

        - シート名は `不良要因印刷_` & UF1.LineName & UF1.LineNo。全ヵ所が作るのは `不良要因印刷_` & UFMitsubishi.LineName
          なので、号機のある丸徳では別のシート(通常検査のもの)を見る/無ければ何もしない(VBA どおり)
        - 右ヘッダは UF1.Lot と検査員名、右フッターの開始は 全ヵ所検査!B1(UFMitsubishi 表示中。H494 でも)
        """
        b = self.book
        st = self._kstate()
        lname = getattr(st, "lname", "") if st is not None else ""
        lot_c = getattr(st, "lot", "") if st is not None else ""
        start = None
        if b.mit.visible:
            v = b.ws(ZEN).get(1, 2)                        # 全ヵ所検査 が無ければ実行時エラー 9
            start = "" if v is None else v
        shifts = self.env().shift_times()
        try:
            with trace.timed("db.save", "不良要因票のページ設定・合計(不良要因印刷範囲)", path=str(self.store.path),
                             table=DEFECT_PREFIX + lname, expected=(vbx.TypeMismatch,)):
                report_model.defect_print_range(self.store, lname, kname, lot_c, now=self.now(), uf1_visible=False,
                                                start1=shifts.start1, end3=shifts.end3, start_override=start)
        except vbx.TypeMismatch as e:
            raise self._vba_error(e) from e

    def _defect_print(self, job: sh.PrintJob) -> None:
        """不良要因の PrintJob(`sh.PrintOut`)→ 手元DB のシートに原紙の書式を当てた帳票の HTML(D8)。"""
        model = report_model.render_defect_sheet(self.store, job.sheet[len(DEFECT_PREFIX):])
        if model is None:
            raise sh.VbaError(9, "インデックスが有効範囲にありません。", f"Sheets({job.sheet})")
        folder = self._print_dir()
        stem = _safe_name(job.sheet)
        with trace.timed("db.save", "不良要因票の HTML", path=str(folder), table=stem) as info:
            files = report_html.save_html_and_pdf(model, folder, stem, None)
            trace.note_used(files["html"], "save")
            info["files"] = len(files)
        p = Path(files["html"])
        self.reports.append({"label": job.sheet, "name": p.name, "path": str(p)})

    # ------------------------------------------------------------------
    # 日報(Register_Check / Daily_Create / Daily_Report)
    # ------------------------------------------------------------------
    def _registered(self, st: Any) -> bool:
        """Worker.Register_Check(Format$(Now, "hh:mm:ss"))。UF1 の検査員で見る。True = 登録が要る。"""
        from . import wiring
        now = self.now()
        env = self.env()
        return workers.register_check(workers.hhmmss(now), wiring.registry(st), wiring.worker_sheet_a(self.store, st.lname),
                                      now=now, facility=terminal.facility() or workers.DEFAULT_FACILITY,
                                      shifts=env.shift_times())

    def _daily_create(self) -> None:
        """新規発行の `If Register_Check(…) Then Call Daily_Create`(UFMit:1253 / 2117)。

        Daily_Create は UF1 が隠れているので Lname = UFMitsubishi.LineName(Worker:721-725)。
        見出しの後の Daily_Report(1)(3) は全ヵ所・H494 表示中のため `Daily_` を探す。
        """
        from . import wiring
        st = self._kstate()
        if st is None or not self._registered(st):
            return
        b = self.book
        env = self.env()
        now = self.now()
        fac = terminal.facility() or workers.DEFAULT_FACILITY
        lname = b.mit.line_name if b.mit.visible else ""
        name = "Daily_" + lname
        day_return = workers.yea_dayr_check(workers.hhmm(now), now=now, facility=fac, shifts=env.shift_times())
        set_time = workers.time_check(workers.hhmmss(now), facility=fac, shifts=env.shift_times())
        cells = wiring.sheet_cells(self.store, name)
        recreate = cells is None or sh.cell_str(cells.get((1, 1))) != day_return + "_" + set_time
        trace.note_used(self.store.path, "save")
        with trace.timed("db.save", "日報の作成(Daily_Create)", path=str(self.store.path), table=name):
            wiring.write_cells(self.store, name, [workers.CellWrite(2, c, v) for c, v in
                                                  enumerate(workers.DAILY_FIELDS, start=1)], recreate=recreate)
        self._daily_report(1)
        self._daily_report(3)

    def _daily_report(self, n: int) -> None:
        """Worker.Daily_Report(n)。全ヵ所・H494 が表示中なので Lname = "" → `Daily_` シート(普通は無い → 何もしない)。"""
        from . import wiring
        mode = "全ヵ所" if self.page == "P1" else "H494"
        name = workers.daily_sheet_name(mode, "")
        cells = wiring.sheet_cells(self.store, name)
        if cells is None:
            return
        b = self.book
        env = self.env()
        hist = wiring.sheet_rows(self.store, LOT_HISTORY)
        lot_sheet = ZEN if ZEN in b.sheets and b.sheets[ZEN].visible else H494
        working = sh.cell_str(b.ws(lot_sheet).get(2, 4)) if lot_sheet in b.sheets else ""
        ctx = workers.DailyContext(exists=True, cells=dict(cells), history_rows=hist[1:] if hist is not None else None,
                                   working_lot=working, now=self.now(),
                                   facility=terminal.facility() or workers.DEFAULT_FACILITY, shifts=env.shift_times(),
                                   mode=mode, sub_line=b.mit.line_name, sub_kname=b.mit.kname)
        try:
            writes = workers.daily_report(n, ctx)
        except IndexError as e:
            raise sh.VbaError(9, "インデックスが有効範囲にありません。", "Worker.Daily_Report") from e
        if writes:
            trace.note_used(self.store.path, "save")
            with trace.timed("db.save", f"日報(Daily_Report {n})", path=str(self.store.path), table=name):
                wiring.write_cells(self.store, name, writes)

    def _log_out(self, name: str, mitsu: str) -> None:
        """補助.LogOut(KenHBox?, "", "三菱")。ラインと検査員は UF1(通常検査)のもの。"""
        if not self.write_logs:
            return
        f = self.book.mit
        m = re.match(r"KenHBox(\d)$", name)
        value = f.kenh.get(int(m.group(1)), "") if m else ""
        st = self._kstate()
        line = st.lname if st is not None else f.line_name
        worker = ""
        if st is not None:
            try:
                from . import wiring
                worker = workers.get_worker_name(wiring.worker_day_max(self.store, st.lname), wiring.registry(st))
            except Exception:                              # noqa: BLE001 - GetWorkerName が引けなければ空
                worker = ""
        for e in legacy_log.log_out_entries(name, value, line, worker, mitsu=mitsu):
            e.write(now=self.now())

    def _settle(self) -> list[Message]:
        """集めた Effects の後始末。画面に返す知らせ(音を含む)を返す。"""
        msgs: list[Message] = list(self._msgs)
        b = self.book
        if b is not None:
            self._defect_after()
        for fx in self._acc:
            msgs.extend(fx.messages)
            for s in fx.sounds:
                msgs.append(Message("", INFO, sound=SOUNDS.get(s, s)))
            if self.write_logs:
                for lg in fx.logs:
                    legacy_log.write_log(lg.text, lg.b, lg.path, lg.fname, now=self.now())
            for form in fx.forms:
                if form.name in ("UFOK", "UFNG"):
                    self.flash = "OK" if form.name == "UFOK" else "NG"
                elif form.name == "UF対角":
                    # UF対角.UserForm_Activate: 絵を 2 秒出して閉じる。UFOK/UFNG/UFcount は閉じる(演出だけ)
                    self.flash = ""
                    self.diag = True
                else:
                    self._pending(form.name)
            self.errors.extend(fx.errors)
            for job in fx.prints:
                self._print_job(job, msgs)
            for plan in fx.saves:
                self._save_plan(plan, msgs)
            for row in fx.history:
                self._history(row)
        self._acc = []
        self._msgs = []
        if b is not None:
            self._persist()
        return msgs

    # ------------------------------------------------------------------
    # 実行の枠
    # ------------------------------------------------------------------
    def _run(self, fn: Callable[..., Any], *args: Any) -> Result:
        with self._lock:
            b = self._ensure()
            self._acc, self._msgs = [], []
            self.flash, self.reports, self.saved, self.ainput, self.pending, self.errors = "", [], [], None, [], []
            self.diag = False
            self._sync_defect()
            q: Optional[Question] = None
            stop: Optional[Stop] = None
            b.fx = b._new_fx()                         # 前の操作の Effects は片付け済み(集め直さない)
            try:
                res = fn(*args)
                if isinstance(res, Question):
                    q = res
            except Stop as e:
                stop = e
            except sh.VbaError as e:
                stop = Stop(f"実行時エラー '{e.number}':" + "\r\n\r\n" + e.description, source=e.where)
            except vbx.TypeMismatch as e:
                stop = Stop("実行時エラー '13':" + "\r\n\r\n" + "型が一致しません。", source=str(e))
            except datasource.DbMissing as e:
                stop = Stop(str(e), source="datasource")
            finally:
                self._harvest(b.fx)
            msgs = self._settle()
            for extra in (q.messages if q else []) + (stop.messages if stop else []):
                if not any(extra is m for m in msgs):
                    msgs.append(extra)
            return Result(self._view(), msgs, question=q, stop=stop)

    def _call(self, fn: Callable[..., Any], *args: Any, **kw: Any) -> Any:
        """部品を 1 つ呼ぶ(前の Effects は begin で集める)。"""
        self.book.begin()
        return fn(*args, **kw)

    # ------------------------------------------------------------------
    # 手元DB(作業中保存)
    # ------------------------------------------------------------------
    def _persist(self) -> None:
        b = self.book
        rows: list[tuple[str, int, int, Any, Optional[int]]] = []
        for wname, sname in SAVE_SHEETS.items():
            ws = b.sheets.get(wname)
            if ws is None:
                continue
            cells: dict[tuple[int, int], list[Any]] = {}
            for (r, c), v in ws.values.items():
                cells[(r, c)] = [_db_value(v), None]
            for (r, c), fm in ws.formulas.items():
                cells[(r, c)] = [fm, None]
            for (r, c), rgb in ws.interior.items():
                cells.setdefault((r, c), [None, None])[1] = _rgb_int(rgb)
            rows.extend((sname, r, c, v, col) for (r, c), (v, col) in sorted(cells.items()))
        meta = {"page": self.page, "active_name": b.active_name, "active_cell": list(b.active_cell),
                "events_enabled": b.events_enabled,
                "d2_red": {n: b.sheets[n].font_color.get((2, 4)) == (255, 0, 0) for n in SAVE_SHEETS if n in b.sheets},
                "underline": {n: sorted(list(x) for x in b.sheets[n].underline) for n in SAVE_SHEETS if n in b.sheets}}
        rows.append((META_SHEET, 1, 1, json.dumps(meta, ensure_ascii=False), None))
        digest = hashlib.sha1(repr(rows).encode("utf-8")).hexdigest()
        if digest == self._digest:
            return
        st = self.store
        trace.note_used(st.path, "save")
        with trace.timed("db.save", "全ヵ所・H494 の作業中保存", path=str(st.path), table="全ヵ所作業中保存") as info:
            with st._lock, st.conn:
                for name in list(SAVE_SHEETS.values()) + [META_SHEET]:
                    st.conn.execute('DELETE FROM "セル" WHERE "シート" = ?', (name,))
                    st.conn.execute('INSERT OR REPLACE INTO "シート" ("名前", "順") VALUES (?, 0)', (name,))
                st.conn.executemany('INSERT INTO "セル" ("シート", "行", "列", "値", "色") VALUES (?, ?, ?, ?, ?)', rows)
            info["cells"] = len(rows)
        self._digest = digest

    def _restore(self) -> None:
        """手元DB の作業中保存からシートを戻す(無ければ原紙のまま)。"""
        st, b = self.store, self.book
        if not st.exists(META_SHEET):
            return
        with trace.timed("db.read", "全ヵ所・H494 の作業中保存を読む", path=str(st.path)) as info:
            meta_v = st.sheet(META_SHEET).get(1, 1)
            meta = json.loads(meta_v) if isinstance(meta_v, str) else {}
            n = 0
            for wname, sname in SAVE_SHEETS.items():
                if not st.exists(sname):
                    continue
                ws = sh.Sheet(wname)
                for r, c, v, col in st.conn.execute('SELECT "行", "列", "値", "色" FROM "セル" WHERE "シート" = ?',
                                                     (sname,)):
                    if isinstance(v, str) and sh._FORMULA_RE.match(v):
                        ws.set_formula(r, c, v)
                    elif v is not None:
                        ws.set(r, c, _from_db(v))
                    if col is not None:
                        ws.interior[(r, c)] = _int_rgb(col)
                    n += 1
                if (meta.get("d2_red") or {}).get(wname):
                    ws.font_color[(2, 4)] = (255, 0, 0)
                ws.underline = {tuple(x) for x in (meta.get("underline") or {}).get(wname, [])}
                b.sheets[wname] = ws
            info["cells"] = n
        self.page = meta.get("page", "P1") if meta.get("page") in ("P1", "P_H494") else "P1"
        b.events_enabled = bool(meta.get("events_enabled", True))
        if meta.get("active_name") in b.sheets:
            b.active_name = meta["active_name"]
        ac = meta.get("active_cell")
        if isinstance(ac, list) and len(ac) == 2:
            b.active_cell = (int(ac[0]), int(ac[1]))

    # ------------------------------------------------------------------
    # 印刷・保存・履歴
    # ------------------------------------------------------------------
    def _print_dir(self) -> Path:
        return self.store.path.parent / "print"

    def _print_job(self, job: sh.PrintJob, msgs: list[Message]) -> None:
        """PrintJob → 帳票の HTML(D8: ブラウザで開いて印刷)。不良要因印刷_<Line> は手元DB のシートから作る。"""
        if not job.snapshot:
            if job.sheet.startswith(DEFECT_PREFIX):
                try:
                    self._defect_print(job)
                except sh.VbaError as e:
                    msgs.append(Message(f"実行時エラー '{e.number}':" + "\r\n\r\n" + e.description, ERROR,
                                        source="PrintOut"))
                return
            self._pending(f"印刷:{job.sheet}")
            return
        lot = sh.cell_str(job.snapshot.get("D2") or "") if job.sheet != "転送済" else ""
        stem = _safe_name(f"{job.sheet}_{lot or self.book.mit.lot_box or self.book.mit.h494_lot}")
        folder = self._print_dir()
        with trace.timed("db.save", "帳票の HTML", path=str(folder), table=stem) as info:
            folder.mkdir(parents=True, exist_ok=True)
            p = folder / f"{stem}.html"
            p.write_text(self._render(job.snapshot, job), encoding="utf-8")
            trace.note_used(p, "save")
            info["files"] = 1
        self.reports.append({"label": job.sheet, "name": p.name, "path": str(p)})

    def _render(self, snapshot: dict[str, Any], job: Optional[sh.PrintJob] = None, title: str = "") -> str:
        model = snapshot_model(snapshot, job)
        out = report_html.render(model, title=title or (job.sheet if job else ""))
        return out.replace("</style>", ".hf span{white-space:pre-line}\n</style>", 1)

    def _save_plan(self, plan: sh.SavePlan, msgs: list[Message]) -> None:
        """SavePlan の手順を実行する。D7: xlsx の代わりに、保存先のファイル名で .html(+ Windows で Edge があれば .pdf)と
        シートごとの .csv を書く。FileCopy はそれらを写す。パスワード・読み取り専用(SetAttr)は付けない。"""
        produced: dict[str, list[Path]] = {}
        edge = paths.to_local(paths.get("edge"))
        edge_s = str(edge) if edge.is_file() and os.name == "nt" else None
        for step in plan.steps:
            try:
                if step.kind == "mkdir":
                    if step.path:
                        legacy_log.local_path(step.path).mkdir(parents=True, exist_ok=True)
                elif step.kind == "save":
                    with trace.timed("db.save", f"保存({plan.name})", path=step.path) as info:
                        files = self._write_book(plan, step.path, edge_s, msgs)
                        info["files"] = len(files)
                    produced[step.path] = files
                    self.saved.extend(str(f) for f in files)
                elif step.kind == "copy":
                    dest = legacy_log.local_path(step.path)
                    with trace.timed("db.save", f"写し({plan.name})", path=step.path) as info:
                        dest.parent.mkdir(parents=True, exist_ok=True)
                        for f in produced.get(step.src, []):
                            shutil.copyfile(f, dest.parent / f.name)
                            trace.note_used(dest.parent / f.name, "save")
                            self.saved.append(str(dest.parent / f.name))
                        info["files"] = len(produced.get(step.src, []))
            except OSError as e:
                msgs.append(Message("保存できませんでした: " + step.path + "\r\n" + str(e), WARN, source=plan.name))

    def _write_book(self, plan: sh.SavePlan, xlsx_path: str, edge: Optional[str], msgs: list[Message]) -> list[Path]:
        lp = legacy_log.local_path(xlsx_path)
        stem = lp.name[:-5] if lp.name.lower().endswith(".xlsx") else lp.name
        folder = lp.parent
        folder.mkdir(parents=True, exist_ok=True)
        files: list[Path] = []
        main: Optional[tuple[str, dict[str, Any]]] = None
        for name, snap in plan.sheets.items():
            if snap is None:
                self._pending(f"保存:{name}")
                continue
            if main is None:
                main = (name, snap)
            p = folder / f"{stem}_{_safe_name(name)}.csv"
            write_csv(p, snap)
            trace.note_used(p, "save")
            files.append(p)
        if main is not None:
            h = folder / f"{stem}.html"
            h.write_text(self._render(main[1], None, title=f"{stem} {main[0]}"), encoding="utf-8")
            trace.note_used(h, "save")
            files.insert(0, h)
            if edge:
                try:
                    files.insert(1, report_html.make_pdf(h, folder / f"{stem}.pdf", edge))
                except report_html.PdfError as e:
                    msgs.append(Message("PDFファイルが作成されませんでした: " + str(folder / f"{stem}.pdf") + "\r\n" + str(e),
                                        WARN, source="補助:739"))
        return files

    def _history(self, row: dict[int, Any]) -> None:
        """`保存用` の最終行の次に履歴行を書く(UFMit:1576-1590 ほか)。"""
        hz = self.store.sheet(HOZON)
        values = [_db_value(row.get(i)) for i in range(1, len(hz.headers) + 1)]
        trace.note_used(self.store.path, "save")
        with trace.timed("db.save", "保存用へ履歴行", path=str(self.store.path), table=HOZON):
            hz.append(values, kind="V")

    # ------------------------------------------------------------------
    # 操作
    # ------------------------------------------------------------------
    def state(self) -> Result:
        with self._lock:
            self._ensure()
            return Result(self._view())

    def start(self, choice_input: Optional[str] = None) -> Result:
        return self._run(self._start, choice_input)

    def _start(self, choice_input: Optional[str]) -> Optional[Question]:
        if self.book.mit.visible:
            return None
        if self._need_worker():
            return None
        return self._open(self.page, choice_input)

    def _need_worker(self) -> bool:
        """UF1.ZEB_Click (UF1:4870-4875) / H494B_Click (UF1:604-609)。開く前の検査員チェック。

        `If Register_Check(Format$(Now, "hh:mm:ss"), "") Then MsgError "検査員登録をして下さい", Lname, 2: Exit Sub`。
        UF1(通常検査の状態)が無ければ登録が無いのと同じに扱う(キャスクの画面と同じ)。真なら開かない。
        """
        st = self._kstate()
        if st is None:
            self._msgs.append(Message(NEED_WORKER, ERROR, "警告", source="UF1:4873"))
            return True
        if not self._registered(st):
            return False
        from . import wiring
        env = self.env()
        env.messages = []
        wiring._msg_error(env, st, NEED_WORKER, 2)
        self._msgs.extend(env.messages)
        env.messages = []
        return True

    def _open(self, page: str, choice_input: Optional[str]) -> Optional[Question]:
        """Mitsubishi.OpenMitubisi (Mit:3-76) + UserForm_Activate。"""
        b = self.book
        if b.active_name in SHEET_ROWS:
            self._sel[b.active_name] = b.active_cell        # シートごとの選択(Excel はシートごとに覚えている)
        prev = b.active_name
        self.page = page
        temp = "全数" if page == "P1" else "H494"
        res = self._call(zk.open_mitubisi, b, temp, line_name=self._line_name(), worker_text=self._worker_text(),
                         choice_input=choice_input, **self._lot_kw())
        if isinstance(res, Question):
            return res
        name = self._sheet()
        lo, hi = SHEET_ROWS[name]
        if prev != name:
            b.active_cell = self._sel.get(name) or self._first_empty(name) or (lo, 2)
        if not (lo <= b.active_cell[0] <= hi):
            b.active_cell = self._first_empty(name) or (lo, 2)
        b.active_name = name                           # シートを開いたときの選択(Excel はイベントなしで戻す)
        return None

    def _first_empty(self, name: str) -> Optional[tuple[int, int]]:
        ws = self.book.ws(name)
        lo, hi = SHEET_ROWS[name]
        for r in range(lo, hi + 1):
            for c in (2, 4, 6, 8):
                if sh.blank(ws.get(r, c)):
                    return (r, c)
        return None

    def choose(self, mode: str, choice_input: Optional[str] = None) -> Result:
        return self._run(self._choose, mode, choice_input)

    def _choose(self, mode: str, choice_input: Optional[str]) -> Optional[Question]:
        b = self.book
        f = b.mit
        if mode == "H494":
            if f.visible and f.work_flag == "P_H494":
                return None
            if f.visible:
                self._call(zk.query_close, b)
            elif self._need_worker():                  # 閉じている → UF1 の H494B_Click と同じ
                return None
            return self._open("P_H494", None)
        if mode not in ("一般", "5N", "三菱"):
            raise Stop("正しい作業選択名を入力してください", level=WARN, source="UFMit:513")
        if not (f.visible and f.work_flag == "P1"):
            if f.visible:
                self._call(zk.query_close, b)
            elif self._need_worker():                  # 閉じている → UF1 の ZEB_Click と同じ
                return None
            q = self._open("P1", choice_input if choice_input is not None else mode)
            if q is not None:
                return q
        f = b.mit
        if f.choice == mode:
            return None
        if mode == "一般":
            self._call(zk.ippan_click, b)
        elif mode == "5N":
            self._call(zk.five_op_click, b)
        else:
            b.begin()
            f.choice = "三菱"                           # Choice(TextBox)へ入力 → Choice_Change(UFMit:509-515)
            zk.choice_change(b)
        return None

    def select(self, r: int, c: int) -> Result:
        return self._run(self._select, r, c)

    def _check_cell(self, r: Any, c: Any) -> tuple[int, int]:
        try:
            r, c = int(r), int(c)
        except (TypeError, ValueError):
            raise Stop("セルの位置が正しくありません", level=WARN) from None
        name = self._sheet()
        if not (1 <= r <= SHEET_ROWS[name][1] and 1 <= c <= SHEET_COLS[name]):
            raise Stop("セルの位置が正しくありません", level=WARN)
        return r, c

    def _select(self, r: Any, c: Any) -> None:
        r, c = self._check_cell(r, c)
        self._call(self.book.user_select, r, c, self._sheet())

    def receive(self, text: str) -> Result:
        return self._run(self._receive, text, True)

    def erase(self) -> Result:
        return self._run(self._receive, "", False)

    def _receive(self, text: Any, enter: bool) -> None:
        b = self.book
        name = self._sheet()
        if b.active_name != name:
            b.activate_sheet(name)
        r, c = self._check_cell(*b.active_cell)
        text = "" if text is None else str(text).rstrip("\r\n")
        self._call(b.user_edit, r, c, text, name)
        if enter and ENTER_MOVES_DOWN and b.active_name == name and b.active_cell == (r, c) \
                and r < SHEET_ROWS[name][1]:
            self._call(b.activate, r + 1, c)

    def next_empty(self) -> Result:
        return self._run(self._next_empty)

    def _next_empty(self) -> None:
        """UFMitsubishi.SB_SpinDown (UFMit:2183-2187)。`MoveTg("A", "全ヵ所検査").Activate`(7〜36 行)。"""
        b = self.book
        if b.mit.tag == "off" or self.page != "P1":
            return
        b.begin()
        cell = sh.move_tg(b, "A", ZEN)
        if cell is None:
            raise Stop("実行時エラー '91':" + "\r\n\r\n" + "オブジェクト変数または With ブロック変数が設定されていません。",
                       source="UFMit:2185")
        b.activate(*cell)

    def set_form(self, name: str, value: Any) -> Result:
        return self._run(self._set_form, name, value)

    def _set_form(self, name: str, value: Any) -> None:
        b = self.book
        f = b.mit
        v = "" if value is None else str(value)
        upper = "".join(ch.upper() if "a" <= ch <= "z" else ch for ch in v
                        if "0" <= ch <= "9" or "A" <= ch <= "Z" or "a" <= ch <= "z")   # KeyPress の入力制限
        if name == "LotBox":
            f.lot_box = upper
            self._call(zk.lot_box_change, b, **self._lot_kw())
        elif name == "H494_Lot":
            f.h494_lot = upper
            self._call(hm.h494_lot_change, b, **self._lot_kw())
        elif name in ("KenNo", "H494_KenNo"):
            setattr(f, "ken_no" if name == "KenNo" else "h494_ken_no", upper)
        elif name in ("Kname", "H494_Kname", "YoutoNo", "H494_Youto", "MINText", "MAXText", "H494_MINText",
                      "H494_MAXText"):
            attr = {"Kname": "kname", "H494_Kname": "h494_kname", "YoutoNo": "youto_no", "H494_Youto": "h494_youto",
                    "MINText": "min_text", "MAXText": "max_text", "H494_MINText": "h494_min_text",
                    "H494_MAXText": "h494_max_text"}[name]
            setattr(f, attr, v)
        elif name in ("Inspection", "op"):
            if v not in METHODS and v != "":
                raise Stop("検査法を選択してください", level=WARN, source="UFMit:1617")
            setattr(f, "inspection_method" if name == "Inspection" else "h494_method", v)
        elif re.match(r"KenHBox[1-7]$", name):
            n = int(name[-1])
            f.kenh[n] = v
            self._call(zk.kenhbox_after_update, b, n)
        elif name == "YamaLB":
            f.yama_index = YAMA_LIST.index(v) if v in YAMA_LIST else None
            self._yama_click()
        else:
            raise Stop("その入力欄はありません", level=WARN)

    def _yama_click(self) -> None:
        """UFMitsubishi.YamaLB_Click (UFMit:2153-2158)。"""
        b = self.book
        b.begin()
        if ZEN not in b.sheets:
            return
        ws = b.ws(ZEN)
        if sh.blank(ws.get(3, 2)):
            return
        self._draw_list_box(sh.val_cell(ws.get(3, 2)))

    def save_height(self) -> Result:
        return self._run(lambda: self._call(zk.mlp_click, self.book))

    def weight(self) -> Result:
        return self._run(lambda: self._call(zk.weight_click, self.book))

    def transfer(self) -> Result:
        return self._run(self._transfer)

    def _transfer(self) -> None:
        if self.page == "P1":
            self._call(zk.transfer, self.book)
        else:
            self._call(hm.transfer, self.book)

    def print(self, which: str, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._print, which, answers or {})

    def _print(self, which: str, answers: dict[str, bool]) -> Any:
        b = self.book
        fn = {"normal": zk.ippan_print, "fini": zk.fiven_print, "mitsubishi": zk.mitsubishi_print,
              "h494": hm.print_and_save}.get(which)
        if fn is None or (which == "h494") != (self.page == "P_H494"):
            raise Stop("この作業ではその印刷は使えません", level=WARN)
        return self._call(fn, b, answers)

    def new_sheet(self, answers: Optional[dict[str, bool]] = None) -> Result:
        return self._run(self._new_sheet, answers or {})

    def _new_sheet(self, answers: dict[str, bool]) -> Any:
        b = self.book
        fn = zk.new_sheet if self.page == "P1" else hm.new_sheet
        res = self._call(fn, b, answers)
        name = self._sheet()
        if not isinstance(res, Question) and name in b.sheets and b.active_name == name:
            b.active_cell = self._first_empty(name) or (7, 2)
        return res

    def resume(self) -> Result:
        return self._run(self._resume)

    def _resume(self) -> None:
        """CommandButton16 / CommandButton21「処理再開」(UFMit:1499-1502 / 126-129)。"""
        self.book.events_enabled = True

    # ------------------------------------------------------------------
    # 画面の値
    # ------------------------------------------------------------------
    def _view(self) -> dict[str, Any]:
        b = self.book
        f = b.mit
        name = self._sheet()
        ws = b.sheets.get(name) or sh.template(sh.ZEN_TEMPLATE if name == ZEN else sh.H494_TEMPLATE)
        mode = "H494" if self.page == "P_H494" else (f.choice if f.choice in ("一般", "5N", "三菱") else "")
        cols = list(H494_COLS if name == H494 else ZEN_COLS + (MITSU_COLS if mode == "三菱" else ()))
        lo, hi = SHEET_ROWS[name]
        rows = []
        for r in range(lo, hi + 1):
            cells = []
            for c, _label, _kind, _g in cols:
                cls = _cls(ws.interior.get((r, c)))
                if (r, c) in ws.underline:
                    cls = (cls + " u").strip()
                cells.append([_text(ws, r, c), cls])
            rows.append({"r": r, "no": _text(ws, r, 1), "cells": cells})
        active = None
        if b.active_name == name:
            ar, ac = b.active_cell
            active = {"r": ar, "c": ac, "label": self._cell_label(name, ar, ac, cols)}
        layout = checks.yama_layout(YAMA_LIST[f.yama_index] if isinstance(f.yama_index, int)
                                    and 0 <= f.yama_index < len(YAMA_LIST) else "", mitsubishi=True)
        kenh_boxes = [{"n": i, "value": f.kenh.get(i, ""),
                       "visible": bool(layout and layout.visible.get(f"KenHBox{i}")),
                       "caption": (layout.hl_captions.get(f"HL{i}", "") if layout else "") if i < 7 else "枚"}
                      for i in range(1, 8)]
        prints = []
        if self.page == "P_H494":
            prints.append({"which": "h494", "label": "検査印刷"})
        else:
            if f.normal_visible:
                prints.append({"which": "normal", "label": "検査印刷"})
            if f.fini_visible:
                prints.append({"which": "fini", "label": "5N印刷"})
            if f.mitsubishi_visible:
                prints.append({"which": "mitsubishi", "label": "三菱印刷"})
        h = self.page == "P_H494"
        return {
            "page": self.page, "mode": mode, "open": f.visible, "line": f.line_name,
            "modes": [{"key": k, "label": lb, "active": k == mode} for k, lb in MODES],
            "notice": W2_NOTICE if mode == "三菱" else "", "suspect": list(zk.MITSUBISHI_SUSPECT) if mode == "三菱" else [],
            "events_enabled": b.events_enabled,
            "form": {"lot": f.h494_lot if h else f.lot_box, "ken_no": f.h494_ken_no if h else f.ken_no,
                     "kname": f.h494_kname if h else f.kname, "youto": f.h494_youto if h else f.youto_no,
                     "min": f.h494_min_text if h else f.min_text, "max": f.h494_max_text if h else f.max_text,
                     "method": f.h494_method if h else f.inspection_method, "methods": list(METHODS),
                     "h494": f.h494, "yama": YAMA_LIST[f.yama_index] if isinstance(f.yama_index, int)
                     and 0 <= f.yama_index < len(YAMA_LIST) else "", "yama_list": list(YAMA_LIST),
                     "kenh": kenh_boxes, "takasa_formula": f.takasa_formula,
                     "names": {"lot": "H494_Lot" if h else "LotBox", "ken_no": "H494_KenNo" if h else "KenNo",
                               "kname": "H494_Kname" if h else "Kname", "youto": "H494_Youto" if h else "YoutoNo",
                               "min": "H494_MINText" if h else "MINText", "max": "H494_MAXText" if h else "MAXText",
                               "method": "op" if h else "Inspection"}},
            "prints": prints,
            "header": self._header(ws, name, mode),
            "groups": self._groups(ws, name, mode, cols),
            "columns": [{"c": c, "label": lb, "kind": k, "group": g, "letter": sh.col_letter(c)} for c, lb, k, g in cols],
            "rows": rows, "active": active, "flash": self.flash, "reports": self.reports, "saved": self.saved,
            "ainput": self.ainput, "pending": self.pending, "errors": self.errors,
            "can_next_empty": self.page == "P1",
            "diag": self.diag,
            "pack": self.pack.to_dict() if self.pack is not None else None,
        }

    def _cell_label(self, name: str, r: int, c: int, cols: list) -> str:
        lab = next((lb for cc, lb, _k, g in cols if cc == c), "")
        group = next((g for cc, _lb, _k, g in cols if cc == c), "")
        if SHEET_ROWS[name][0] <= r <= SHEET_ROWS[name][1]:
            head = lab if lab != "判" else f"{group}の判定"
            if group == "ラテラル" and lab == "ﾗﾃﾗﾙ":
                head = "ﾗﾃﾗﾙﾎﾞｰ"
            return f"No.{r - 6} {head or sh.col_letter(c)}"
        hdr = next((x["label"] for x in self._header(self.book.ws(name), name, "") if x.get("cell") == [r, c]), "")
        return f"{sh.col_letter(c)}{r} {hdr}".strip()

    def _header(self, ws: sh.Sheet, name: str, mode: str) -> list[dict[str, Any]]:
        t = lambda r, c: _text(ws, r, c)  # noqa: E731
        t1 = lambda r, c: _fmt_num(ws.get(r, c), "0.0")  # noqa: E731
        size = " × ".join(x for x in (_fmt_num(ws.get(3, 2), "0.000"), _fmt_num(ws.get(3, 4), "0.0"),
                                      _fmt_num(ws.get(3, 6), "0.0")) if x)
        items: list[dict[str, Any]] = [
            {"label": "開始", "text": t(1, 2), "cell": [1, 2]},
            {"label": "LotNo", "text": t(2, 4), "cell": [2, 4], "red": ws.font_color.get((2, 4)) == (255, 0, 0)},
            {"label": "検査No", "text": t(2, 8), "cell": [2, 8]},
            {"label": "検査員", "text": t(3, 8), "cell": [3, 8]},
            {"label": "用途", "text": t(4, 8), "cell": [4, 8]},
            {"label": "寸法", "text": size},
            {"label": "ﾗﾃﾗﾙ等級", "text": t(1, 12), "cell": [1, 12]},
            {"label": "ﾗﾃﾗﾙ公差", "text": t1(1, 13), "cell": [1, 13], "edit": True},
            {"label": "直角度等級" if name == ZEN else "対角等級", "text": t(1, 18), "cell": [1, 18]},
            {"label": "対角公差", "text": t1(1, 19), "cell": [1, 19], "edit": True},
        ]
        if name == ZEN:
            items += [
                {"label": "歪公差 幅", "text": t1(5, 13), "cell": [5, 13], "edit": mode == "一般"},
                {"label": "歪公差 丈", "text": t1(5, 15), "cell": [5, 15], "edit": mode == "一般"},
                {"label": "作業", "text": t(1, 22), "cell": [1, 22]},
                {"label": "測定機器", "text": t(3, 29), "cell": [3, 29]},
                {"label": "MAX 厚/巾/丈", "text": " / ".join((t(3, 23), t(3, 25), t(3, 27)))},
                {"label": "MIN 厚/巾/丈", "text": " / ".join((t(4, 23), t(4, 25), t(4, 27)))},
            ]
        else:
            items += [
                {"label": "測定機器", "text": t(2, 25), "cell": [2, 25]},
                {"label": "MAX 厚/巾/丈", "text": " / ".join((t(3, 22), t(3, 23), t(3, 24)))},
                {"label": "MIN 厚/巾/丈", "text": " / ".join((t(4, 22), t(4, 23), t(4, 24)))},
            ]
        items.append({"label": "検入高さ", "text": "  ".join(x for x in (t(1, 30), t(2, 30), t(3, 30)) if x)})
        return items

    def _groups(self, ws: sh.Sheet, name: str, mode: str, cols: list) -> list[dict[str, Any]]:
        """表の見出しのまとまりと、その公差(モードで違う所はここで分かるように)。"""
        def rng(r: int, pat: str) -> str:
            a, z = ws.get(r, 17), ws.get(r, 19)
            if isinstance(a, sh.ExcelError) or isinstance(z, sh.ExcelError):
                return ""
            if sh.blank(ws.get(r, 13)) and sh.blank(ws.get(r, 15)):
                return ""
            return f"{_fmt_num(a, pat)}〜{_fmt_num(z, pat)}"

        m1, s1 = _fmt_num(ws.get(1, 13), "0.0"), _fmt_num(ws.get(1, 19), "0.0")
        if name == H494:
            hiz = "判定なし"
        elif mode == "5N":
            hiz = "幅 2.0 / 丈 3.0(5N 固定)"
        elif mode == "一般":
            hiz = f"幅 {_fmt_num(ws.get(5, 13), '0.0') or '(空)'} / 丈 {_fmt_num(ws.get(5, 15), '0.0') or '(空)'}"
        elif mode == "三菱":
            hiz = "判定なし(三菱)"
        else:
            hiz = ""
        tol = {"厚": rng(2, "0.000"), "幅": rng(3, "0.0"), "丈": rng(4, "0.0"),
               "対角": f"公差 {s1}" if s1 else "公差なし",
               "ラテラル": (f"公差 {m1}" if m1 else "公差なし") + ("(7〜10行のみ判定)" if name == H494 else ""),
               "歪": hiz, "ゲージ": "", "外観": "", "三菱": "ﾌﾞﾛｯｸNO"}
        out: list[dict[str, Any]] = []
        for _c, _lb, _k, g in cols:
            if out and out[-1]["label"] == g:
                out[-1]["span"] += 1
            else:
                out.append({"label": g, "span": 1, "tol": tol.get(g, "")})
        return out


# ------------------------------------------------------------------
# 帳票(シートの値 → report_model.ReportModel → report_html.render)
# ------------------------------------------------------------------
_HF_CODE = re.compile(r'&"[^"]*"|&\d+')
_AREA = re.compile(r"^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$")
_THIN = {s: ["continuous", "thin"] for s in ("top", "bottom", "left", "right")}


def _hf(text: str) -> str:
    """Excel のヘッダ・フッタの書式コード(&"Meiryo UI"&8 など)を除き、改行(\\r)を改行に。"""
    return _HF_CODE.sub("", text or "").replace("\r", "\n")


def snapshot_model(snapshot: dict[str, Any], job: Optional[sh.PrintJob] = None) -> ReportModel:
    """シートの値(A1 → 値)と PrintJob の設定から帳票モデルを作る(罫線は範囲の全部に細線)。"""
    cells_rc: dict[tuple[int, int], Any] = {}
    for ref, v in snapshot.items():
        cells_rc[sh.a1(ref)] = v
    area = None
    if job and job.print_area:
        m = _AREA.match(job.print_area)
        if m:
            area = (int(m.group(2)), sh.col_number(m.group(1)), int(m.group(4)), sh.col_number(m.group(3)))
    if area is None and cells_rc:
        area = (min(r for r, _ in cells_rc), min(c for _, c in cells_rc),
                max(r for r, _ in cells_rc), max(c for _, c in cells_rc))
    r1, c1, r2, c2 = area or (1, 1, 1, 1)
    cells = []
    for (r, c), v in sorted(cells_rc.items()):
        if not (r1 <= r <= r2 and c1 <= c <= c2) or v is None:
            continue
        if isinstance(v, datetime):
            text = sh.fmt_date(v, "yyyy/mm/dd h:mm")
        elif isinstance(v, sh.ExcelError):
            text = str(v)
        else:
            try:
                text = sh.cell_str(v)
            except sh.VbaError:
                text = str(v)
        cells.append({"ref": f"{sh.col_letter(c)}{r}", "row": r, "col": c,
                      "value": v if isinstance(v, (int, float, str)) else text, "text": text, "borders": _THIN,
                      "font_size": 9, "bold": False, "red": text == "×", "h_align": "center" if text in ("○", "×") else "",
                      "v_align": "center", "wrap": "\n" in text})
    for r in range(r1, r2 + 1):                  # 空のセルにも罫線(表の形を見せる)
        for c in range(c1, c2 + 1):
            if (r, c) not in cells_rc or cells_rc[(r, c)] is None:
                cells.append({"ref": f"{sh.col_letter(c)}{r}", "row": r, "col": c, "value": None, "text": "",
                              "borders": _THIN, "font_size": 9})
    widths: dict[str, float] = {}
    for c in range(c1, c2 + 1):
        texts = [x["text"] for x in cells if x["col"] == c and x["text"]]
        w = max([len(t) for t in texts] + [2])
        widths[sh.col_letter(c)] = float(min(max(w + 1, 3), 16))
    total_px = sum(int(round(w * 7 + 5)) for w in widths.values())
    zoom = 100
    if job and job.fit_wide == 1 and total_px > 0:
        zoom = max(30, min(100, int(1000 * 100 / total_px)))
    page = {"orientation": job.orientation if job else "landscape",
            "margins_inch": dict(job.margins) if job and job.margins else {}, "zoom": zoom, "header_size": 8,
            "header_left": _hf(job.left_header) if job else "", "header_center": _hf(job.center_header) if job else "",
            "header_right": _hf(job.right_header) if job else "", "footer_left": _hf(job.left_footer) if job else "",
            "footer_center": _hf(job.center_footer) if job else "",
            "footer_right": _hf(job.right_footer) if job else ""}
    return ReportModel(sheet=job.sheet if job else "", kind=job.sheet if job else "シート", print_area=(r1, c1, r2, c2),
                       page=page, cells=cells, merges=[], col_widths=widths, row_heights={})


def write_csv(path: Path, snapshot: dict[str, Any]) -> None:
    """シートの値を CSV に(UTF-8 BOM 付き。Excel で開ける)。"""
    cells = {sh.a1(ref): v for ref, v in snapshot.items()}
    if not cells:
        path.write_text("", encoding="utf-8-sig")
        return
    rmax = max(r for r, _ in cells)
    cmax = max(c for _, c in cells)
    with open(path, "w", newline="", encoding="utf-8-sig") as fp:
        w = csv.writer(fp)
        for r in range(1, rmax + 1):
            row = []
            for c in range(1, cmax + 1):
                v = cells.get((r, c))
                if v is None:
                    row.append("")
                elif isinstance(v, datetime):
                    row.append(sh.cstr_date(v))
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    row.append(vbx.cstr(v))
                else:
                    row.append(str(v))
            w.writerow(row)


# ------------------------------------------------------------------
# プロセスに 1 つ
# ------------------------------------------------------------------
_session: Optional[ZenkashoSession] = None
_slock = threading.Lock()


def get() -> ZenkashoSession:
    global _session
    with _slock:
        if _session is None:
            _session = ZenkashoSession()
        return _session


def reset_for_test(s: Optional[ZenkashoSession] = None) -> ZenkashoSession:
    global _session
    with _slock:
        _session = s or ZenkashoSession()
        return _session


__all__ = ["ZenkashoSession", "get", "reset_for_test", "snapshot_model", "write_csv", "W2_NOTICE", "SAVE_SHEETS"]
