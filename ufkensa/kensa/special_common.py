"""キャスク・全ヵ所・H494 に共通の部品(再構築 第6段)。

VBA では UFCask と UFMitsubishi が同じ処理(HistorySave の反映・包装仕様書・Register_Check・Daily_Create の見出し)を
それぞれ持っていた。ここに 1 つだけ置き、両方の画面のセッションが使う(業務の中身は VBA のまま)。

使う側が持つもの: store / now / _lock / pack / _run / env() / _sink()(知らせの置き場)/ _facility()
"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Optional

from .. import trace
from . import lot as lotmod
from . import packing, terminal, wiring, workers
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
