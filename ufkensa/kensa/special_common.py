"""キャスク・全ヵ所に共通の部品(再構築 第6段)。

VBA では UFCask と UFMitsubishi が同じ処理(HistorySave の反映・包装仕様書・Register_Check・Daily_Create の見出し)を
それぞれ持っていた。ここに 1 つだけ置き、両方の画面のセッションが使う(業務の中身は VBA のまま)。

使う側が持つもの: store / now / _lock / pack / _run / env() / _sink()(知らせの置き場)/ _facility()
"""
from __future__ import annotations

import shutil
from glob import escape as glob_escape
from pathlib import Path
from typing import Any, Callable, Optional

from .. import trace
from . import lot as lotmod
from . import datasource, legacy_log, packing, report_html, quality, terminal, vbx, wiring, workers
from .flow import WARN, Message, Stop
from .modes import sheet as sh
from .store import LOT_HISTORY


class SubInspectionCommon:
    """特殊検査の画面に共通の操作。"""

    pack: Any = None

    # -- 使う側で決めるもの -----------------------------------------------
    def _sink(self) -> list[Message]:
        raise NotImplementedError

    def _facility(self) -> str:
        return terminal.facility() or workers.DEFAULT_FACILITY

    # -- Lot情報履歴 -------------------------------------------------------
    def _apply_history(self, h: lotmod.HistorySaveResult) -> int:
        """HistorySave の結果を Lot情報履歴 に反映する(行を消してから足す。通常検査と同じ)。"""
        lh = self.store.sheet(LOT_HISTORY)
        trace.note_used(self.store.path, "save")
        for r in sorted(h.delete_rows, reverse=True):
            lh.delete_row(r)
        for row in h.new_rows:
            lh.append(row[:26])
        return len(h.new_rows)

    # -- 検査員・日報 ------------------------------------------------------
    def _registered(self, uf1: Any) -> bool:
        """Worker.Register_Check(UF1 の検査員で見る)。True = 登録が要る。"""
        now = self.now()
        return workers.register_check(workers.hhmmss(now), wiring.registry(uf1),
                                      wiring.worker_sheet_a(self.store, uf1.lname), now=now,
                                      facility=self._facility(), shifts=self.env().shift_times())

    def _daily_header(self, lname: str) -> None:
        """Daily_Create の見出し(Worker:721-740)。直が変わっていれば作り直す。Daily_Report(1)(3) は呼び元で。"""
        now = self.now()
        fac, shifts = self._facility(), self.env().shift_times()
        name = "Daily_" + lname
        day_return = workers.yea_dayr_check(workers.hhmm(now), now=now, facility=fac, shifts=shifts)
        set_time = workers.time_check(workers.hhmmss(now), facility=fac, shifts=shifts)
        cells = wiring.sheet_cells(self.store, name)
        recreate = cells is None or sh.cell_str(cells.get((1, 1))) != day_return + "_" + set_time
        trace.note_used(self.store.path, "save")
        with trace.timed("db.save", "日報の作成(Daily_Create)", path=str(self.store.path), table=name):
            wiring.write_cells(self.store, name, [workers.CellWrite(2, c, v) for c, v in
                                                  enumerate(workers.DAILY_FIELDS, start=1)], recreate=recreate)

    # -- 包装仕様書(UF_PackNo) -------------------------------------------
    def pack_select(self, i: Any) -> Any:
        return self._run(self._pack_select, i)

    def _pack_select(self, i: Any) -> None:
        """UF_PackNo の引当の行をクリック(LabelCls → Loop_Label → JPGLoad)。"""
        if self.pack is None:
            raise Stop("包装仕様書の画面は開いていません", level=WARN)
        try:
            n = int(i)
        except (TypeError, ValueError):
            raise Stop("行の番号が正しくありません", level=WARN) from None
        packing.pack_select(self.pack, n, wiring.sheet_rows(self.store, LOT_HISTORY), messages=self._sink())

    def pack_close(self, confirmed: Any = False) -> Any:
        return self._run(self._pack_close, bool(confirmed))

    def _pack_close(self, confirmed: bool) -> None:
        """UF_PackNo.CommandButton2_Click(「確認」が無いと「確認後に押してください。」で閉じない)。"""
        m = packing.pack_close(confirmed)
        if m is not None:
            self._sink().append(m)
            return
        self.pack = None

    def pack_image(self) -> Optional[Path]:
        """いま出している包装仕様書の画像(使用パスの QualityPath\\包装仕様書)。"""
        with self._lock:
            return packing.picture_file(self.pack.picture) if self.pack is not None else None

    # -- ロット選択の続き(HistorySave・包装仕様書) -------------------------
    #: HistorySave の途中で落ちる実行時エラー(ComboBox1_Change のエラー処理へ)
    _HISTORY_ERRORS = (vbx.TypeMismatch, lotmod.SubscriptOutOfRange, IndexError, datasource.DbMissing)
    _PACKING_ERRORS = (packing.PackingError, quality.VbaRuntimeError, vbx.TypeMismatch)

    def _lot_read_failed(self, e: Exception, msgs: list[Message]) -> None:
        """ロット選択のエラー処理。キャスクは「DB読み込みエラー:」を出す、全ヵ所は VBA のエラーで止める。"""
        raise NotImplementedError

    def _history_env(self) -> lotmod.HistoryEnv:
        raise NotImplementedError

    def _after_history(self, res: lotmod.HistorySaveResult) -> None:
        """HistorySave の後に画面ごとにすること(既定は無し)。"""

    def _history_save(self, past: list[Any]) -> list[Message]:
        """`Call HistorySave(PastHiki)`(UFCask:535・UFMit:804)。Lot情報履歴 に引当ごとの行。

        途中の実行時エラーは、そこまでの行の削除を反映してから画面のエラー処理へ(VBA どおり)。
        g_CacheLotNo などの UF1 の控えは通常検査の状態なので書き換えない。
        """
        henv = self._history_env()
        hist = wiring.sheet_rows(self.store, LOT_HISTORY) or []
        msgs: list[Message] = []
        with trace.timed("db.save", "Lot情報履歴(HistorySave)", path=str(self.store.path), table=LOT_HISTORY,
                         expected=(Stop, sh.VbaError)) as info:
            try:
                res = lotmod.history_save(past, hist, henv)
            except self._HISTORY_ERRORS as e:
                partial = getattr(e, "partial", None)
                if partial is not None:
                    self._apply_history(partial)
                self._lot_read_failed(e, msgs)
                return msgs
            info["rows"] = self._apply_history(res)
        msgs.extend(res.messages)
        self._after_history(res)
        return msgs

    def _packing(self, lot_no: str) -> list[Message]:
        """`If Not PackingNoCheck(LotNo) = "" Then Call UFPack_Open(…)`(UFCask:538・UFMit:856-858)。"""
        msgs: list[Message] = []
        hist = wiring.sheet_rows(self.store, LOT_HISTORY)
        with trace.timed("packing", "包装仕様書の確認(PackingNoCheck)", lot=lot_no,
                         expected=(packing.PackingError, sh.VbaError)):
            try:
                n = packing.packing_no_check(hist, lot_no, messages=msgs)
                form = packing.pack_open(n, lot_no, hist, messages=msgs) if n != "" else None
            except self._PACKING_ERRORS as e:
                self._lot_read_failed(e, msgs)
                return msgs
        if form is not None:
            self.pack = form
        return msgs

    # -- 保存(SavePlan) ----------------------------------------------------
    def _exists(self, path_text: str) -> bool:
        """保存の連番(numbered_path)の「もうあるか」。xlsx と、一緒に書く見せる用の .html/.pdf を見る。"""
        p = legacy_log.local_path(path_text)
        if p.exists():
            return True
        stem = _stem(p.name)
        return any((p.parent / (stem + ext)).exists() for ext in (".xlsx", ".html", ".pdf")) or \
            any(p.parent.glob(glob_escape(stem) + "_*.html"))

    def _run_save_plan(self, plan: sh.SavePlan,
                       write: Callable[[Path, str], list[Path]]) -> tuple[list[str], list[Message]]:
        """SavePlan の手順(mkdir / save / copy)を実行する。キャスク・全ヵ所で同じ。

        save: `write(フォルダ, 名前)` で xlsx(+ 見せる用の HTML/PDF)を書く。
        copy: save で作ったファイルを、写し先の名前に付け替えて写す(VBA の FileCopy)。
        パスワード・読み取り専用(xlsx の保護)は付けない。書けなかった手順は知らせて次へ進む。
        """
        msgs: list[Message] = []
        made: dict[str, list[Path]] = {}
        files: list[str] = []
        for step in plan.steps:
            if not step.path:
                continue
            target = legacy_log.local_path(step.path)
            try:
                if step.kind == "mkdir":
                    target.mkdir(parents=True, exist_ok=True)
                elif step.kind == "save":
                    stem = _stem(target.name)
                    with trace.timed("db.save", f"保存({plan.name})", path=str(target.parent), table=stem) as info:
                        target.parent.mkdir(parents=True, exist_ok=True)
                        made[step.path] = write(target.parent, stem)
                        info["files"] = len(made[step.path])
                    files.extend(str(p) for p in made[step.path])
                elif step.kind == "copy":
                    src_stem = _stem(Path(step.src.replace("\\", "/")).name)
                    dst_stem = _stem(target.name)
                    with trace.timed("db.save", f"写し({plan.name})", path=str(target.parent)) as info:
                        target.parent.mkdir(parents=True, exist_ok=True)
                        for p in made.get(step.src, []):
                            dst = target.parent / (dst_stem + p.name[len(src_stem):])
                            shutil.copyfile(p, dst)
                            trace.note_used(dst, "save")
                            files.append(str(dst))
                        info["files"] = len(made.get(step.src, []))
            except (OSError, report_html.PdfError) as e:
                msgs.append(Message("保存できませんでした: " + str(target) + "\r\n" + str(e), WARN, source=plan.name))
        return files, msgs


def _stem(name: str) -> str:
    return name[:-5] if name.lower().endswith(".xlsx") else name
