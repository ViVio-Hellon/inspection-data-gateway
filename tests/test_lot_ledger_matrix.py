"""仕掛台帳の第1/第2の読み分けの書き留め(リファクタリング 第2段・改善案 7)

読み手ごとに手で書いていた第1/第2の読み分けを 1 か所(`lot_ledger`)にまとめる前に、今の動きを
モード × 台帳の状態 × 読み手 の組み合わせで書き留めた(`fixtures/lot_ledger_matrix.json`。作り直さない)。
まとめた後の動きは、書き留めに `fixtures/lot_ledger_matrix_intended.json`(決めた決まりにそろえた所。直し。
理由つき)を重ねたものと同じであること。

- モード: online / offline / lot_source(オンラインの設定で「控えから読む」)/ dbflag(オンラインで DBflag あり)
- 台帳の状態(ロット H2500K0。第2は値を変えて見分ける: 枚本数 999・ｵｰﾀﾞｰ板厚 9.5・引当数量 999・納入先 ﾀﾞｲ2):
  both / first_file_missing / second_file_missing / lot_only_first / lot_only_second / second_missing_columns /
  hiki_extra_second / hiki_extra_first / odr_missing_first / odr_missing_second /
  hiki_lot_only_first / hiki_lot_only_second(引当だけ片側)/ lot_absent_both
- 読み手: arrbox(Record・Fields)・acc_reflec・history_save・sika_atu_create・atu_hiki_records・
  grid の既定・judge(simple_arrbox・materials_msgbox)・lot_watch.signature・quality.hiki_arr_create・
  paths.db_file

書き留め直すとき(まとめる前のコードでだけ): UFKENSA_RECORD_LEDGER=1 python -m pytest tests/test_lot_ledger_matrix.py
"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import sqlite3
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Callable
from unittest import mock

from tests import kensa_fixture
from ufkensa import user_settings
from ufkensa.kensa import cache_registry, datasource, grid, judge, lot, lot_watch, paths, wiring

BASELINE = Path(__file__).parent / "fixtures" / "lot_ledger_matrix.json"
LOT = "H2500K0"
ORDERS = ("06017610", "06017503", "06019756")
EXTRA_HIKI = "69999999"
MODES = ("online", "offline", "lot_source", "dbflag")
SCENARIOS = ("both", "first_file_missing", "second_file_missing", "lot_only_first", "lot_only_second",
             "second_missing_columns", "hiki_extra_second", "hiki_extra_first", "odr_missing_first",
             "odr_missing_second", "hiki_lot_only_first", "hiki_lot_only_second", "lot_absent_both")
DROP_COLS = ("品質ｸﾞﾚｰﾄﾞ_直角度", "BOX実績_枚本数")


def _ok(*_a: Any) -> str:
    return "可"


def _sql(path: Path, *stmts: tuple) -> None:
    con = sqlite3.connect(path)
    with con:
        for s in stmts:
            con.execute(*s) if isinstance(s, tuple) else con.execute(s)
    con.close()


def _digest(v: Any) -> str:
    return hashlib.sha1(repr(v).encode("utf-8")).hexdigest()[:10]


class _Case:
    """1 つの (状態, モード)。第1 = root/first、第2 = root/second。"""

    def __init__(self, scenario: str, mode: str) -> None:
        self.scenario, self.mode = scenario, mode
        self.root = Path(tempfile.mkdtemp(prefix="ufk_ledger_"))
        self.d1, self.d2 = self.root / "first", self.root / "second"
        src = kensa_fixture.build()
        for d in (self.d1, self.d2):
            d.mkdir()
            for name in paths.LOT_DBS:
                shutil.copy2(src / kensa_fixture.file_name(name), d)
        lotf, hiki, odr = (kensa_fixture.file_name(n) for n in paths.LOT_DBS)
        # 第2を見分ける印
        _sql(self.d2 / lotf, ('UPDATE "仕掛" SET "BOX実績_枚本数" = ?, "ｵｰﾀﾞｰ板厚" = ? WHERE "ﾛｯﾄ番号" = ?',
                              ("999", "9.5", LOT)))
        _sql(self.d2 / hiki, ('UPDATE "仕掛" SET "引当数量" = ? WHERE "ﾛｯﾄ番号" = ?', ("999", LOT)))
        _sql(self.d2 / odr, ('UPDATE "仕掛" SET "納入先名称" = ? WHERE "受注番号" IN (?, ?, ?)', ("ﾀﾞｲ2",) + ORDERS))
        s = scenario
        if s == "first_file_missing":
            for f in (lotf, hiki, odr):
                (self.d1 / f).unlink()
        elif s == "second_file_missing":
            for f in (lotf, hiki, odr):
                (self.d2 / f).unlink()
        elif s in ("lot_only_first", "lot_only_second"):
            d = self.d2 if s == "lot_only_first" else self.d1
            _sql(d / lotf, ('DELETE FROM "仕掛" WHERE "ﾛｯﾄ番号" = ?', (LOT,)))
            _sql(d / hiki, ('DELETE FROM "仕掛" WHERE "ﾛｯﾄ番号" = ?', (LOT,)))
            _sql(d / odr, ('DELETE FROM "仕掛" WHERE "受注番号" IN (?, ?, ?)', ORDERS))
        elif s in ("hiki_lot_only_first", "hiki_lot_only_second"):
            d = self.d2 if s == "hiki_lot_only_first" else self.d1
            _sql(d / hiki, ('DELETE FROM "仕掛" WHERE "ﾛｯﾄ番号" = ?', (LOT,)))
        elif s == "lot_absent_both":
            for d in (self.d1, self.d2):
                _sql(d / lotf, ('DELETE FROM "仕掛" WHERE "ﾛｯﾄ番号" = ?', (LOT,)))
                _sql(d / hiki, ('DELETE FROM "仕掛" WHERE "ﾛｯﾄ番号" = ?', (LOT,)))
        elif s == "second_missing_columns":
            _sql(self.d2 / lotf, *[f'ALTER TABLE "仕掛" DROP COLUMN "{c}"' for c in DROP_COLS])
        elif s in ("hiki_extra_second", "hiki_extra_first"):
            d = self.d2 if s == "hiki_extra_second" else self.d1
            _sql(d / hiki, ('CREATE TEMP TABLE t AS SELECT * FROM "仕掛" WHERE "ﾛｯﾄ番号" = ? AND "引当番号" = ?',
                            (LOT, "61026069")),
                 ('UPDATE t SET "引当番号" = ?', (EXTRA_HIKI,)),
                 'INSERT INTO "仕掛" SELECT * FROM t')
        elif s in ("odr_missing_first", "odr_missing_second"):
            d = self.d1 if s == "odr_missing_first" else self.d2
            _sql(d / odr, ('DELETE FROM "仕掛" WHERE "受注番号" = ?', (ORDERS[0],)))
        paths.set_mode(paths.OFFLINE if mode == "offline" else paths.ONLINE)
        items = {k: str(self.root) + "/" for k in ("quality", "packing", "history", "factory")}
        user_settings.save(paths.KEY_PATHS, {paths.KEY_ITEMS: items,
                                             paths.KEY_LOT: {paths.FIRST: str(self.d1) + "/",
                                                             paths.SECOND: str(self.d2) + "/"}})
        cache_registry.clear_all()
        self.dbflag = "1" if mode == "dbflag" else ""

    def norm(self, v: Any) -> Any:
        if isinstance(v, str):
            return v.replace(str(self.d1), "<1>").replace(str(self.d2), "<2>")
        if isinstance(v, (list, tuple)):
            return [self.norm(x) for x in v]
        if isinstance(v, dict):
            return {k: self.norm(x) for k, x in v.items()}
        return v

    def close(self) -> None:
        shutil.rmtree(self.root, ignore_errors=True)


def _rows(rows: Any, cols: Any, pick: tuple[str, ...]) -> Any:
    if rows is None:
        return None
    idx = [datasource.column_index(cols or [], c) for c in pick]
    return {"n": len(rows), "pick": sorted([[r[i] if 0 <= i < len(r) else "?" for i in idx] for r in rows]),
            "digest": _digest(sorted(map(tuple, rows)))}


def _guard(fn: Callable[[], Any]) -> Any:
    try:
        return fn()
    except Exception as e:                                       # noqa: BLE001 - 例外も書き留める
        return {"raises": type(e).__name__}


LOT_PICK = ("ﾛｯﾄ番号", "BOX実績_枚本数", "ｵｰﾀﾞｰ板厚", "品質ｸﾞﾚｰﾄﾞ_直角度")
HIKI_PICK = ("引当番号", "引当数量", "受注番号")


def observe(c: _Case) -> dict[str, Any]:
    out: dict[str, Any] = {}
    flag = c.dbflag

    def msgs(ms: Any) -> list[str]:
        return [m.text for m in ms]

    # arrbox
    for target, pick in (("SIKALOTNOW", LOT_PICK), ("SIKAHIKINOW", HIKI_PICK)):
        def a(target: str = target, pick: tuple = pick) -> Any:
            f = lot.arrbox(target, "", "Fields", flag, over12=False)
            r = lot.arrbox(target, LOT, "Record", flag, over12=False)
            return {"file": r.file_path, "fields_file": f.file_path, "fields": _digest(f.value),
                    "rows": _rows(r.value, f.value, pick), "msgs": msgs(r.messages) + msgs(f.messages)}
        out[f"arrbox.{target}"] = _guard(a)

    # acc_reflec
    def acc() -> Any:
        r = lot.acc_reflec(lot.LotScreen(lot=LOT, cb_text=LOT, dbflag=flag, work_time="T"), check_equipment=_ok)
        row = r.row or {}
        return {"status": r.status, "file": r.db_path, "msgs": msgs(r.messages),
                "pick": [row.get(k) for k in LOT_PICK], "past_hiki": _digest(r.past_hiki)}
    out["acc_reflec"] = _guard(acc)

    # history_save(Lot情報履歴)
    def hist() -> Any:
        ph = [None] * 19
        ph[0], ph[1], ph[9] = "'06017610", LOT, "KEN"
        r = lot.history_save(ph, [list(lot.HISTORY_HEADER)], lot.HistoryEnv(dbflag=flag))
        return {"completed": r.completed, "file": r.file_path, "msgs": msgs(r.messages),
                "rows": sorted([[x[21], x[23], x[12], x[1]] for x in r.new_rows]),
                "digest": _digest(sorted([repr(x[1:]) for x in r.new_rows]))}
    out["history_save"] = _guard(hist)

    # sika_atu_create(lot_source は呼び手の paths.lot_source の中で読む)
    out["sika_atu_create"] = _guard(lambda: lot.sika_atu_create(LOT))

    # atu_hiki_records(wiring)
    def atu() -> Any:
        found, rows = wiring.RealHooks.atu_hiki_records(None, SimpleNamespace(dbflag=flag), LOT)
        return {"found": found, "rows": _rows(rows, _hiki_cols(c), HIKI_PICK)}
    out["atu_hiki_records"] = _guard(atu)

    # grid の既定
    st = SimpleNamespace(dbflag=flag)
    h = grid.Hooks()
    out["grid.arrbox.SIKALOTNOW"] = _guard(lambda: _rows(h.arrbox(st, "SIKALOTNOW.accdb", LOT, "Record"),
                                                         h.arrbox(st, "SIKALOTNOW.accdb", "", "Fields"), LOT_PICK))
    out["grid.atu_hiki_records"] = _guard(lambda: (lambda fr: {"found": fr[0], "rows": _rows(fr[1], _hiki_cols(c),
                                                                                              HIKI_PICK)})(
        h.atu_hiki_records(st, LOT)))

    # judge
    out["judge.simple_arrbox"] = _guard(lambda: _rows(judge.simple_arrbox("SIKALOTNOW.accdb", LOT, "Record", flag),
                                                      judge.simple_arrbox("SIKALOTNOW.accdb", "", "Fields", flag),
                                                      LOT_PICK))

    def mm() -> Any:
        seen: list = []
        with mock.patch.object(judge, "packaging_specification_no",
                               lambda *a, **k: seen.append(list(a)) or ""):
            o = judge.materials_msgbox(judge.new_form(Lot=LOT, DBflag=flag, ATU="1", Hab="1", Tak="1"), ORDERS[0],
                                       arrbox=lambda t, f, u, p: None)
        return {"args": seen, "msgs": msgs(o.messages)}
    out["judge.materials_msgbox"] = _guard(mm)

    # 引当の納入先(品質基準の DistArray に渡す。quality.hiki_arr_create)
    from ufkensa.kensa import quality
    out["quality.hiki_arr_create"] = _guard(lambda: quality.hiki_arr_create(LOT))
    # 名前で引く DB の場所(全ヵ所・キャスク・保存の確かめなどが datasource.resolve("SIKALOTNOW") で読む)
    out["paths.db_file"] = _guard(lambda: [paths.db_file(n) for n in paths.LOT_DBS])

    # lot_watch
    def sig() -> Any:
        s = lot_watch.signature(LOT, flag)
        return {k: v for k, v in s.parts.items() if k != "品質基準"}
    out["lot_watch.signature"] = _guard(sig)
    return c.norm(out)


def _hiki_cols(c: _Case) -> list[str]:
    for d in (c.d1, c.d2):
        p = d / kensa_fixture.file_name("SIKAHIKINOW")
        if p.exists():
            con = sqlite3.connect(p)
            try:
                cols = [r[1] for r in con.execute('PRAGMA table_info("仕掛")')]
            finally:
                con.close()
            return cols
    return []


def run_matrix() -> dict[str, Any]:
    out: dict[str, Any] = {}
    for s in SCENARIOS:
        for m in MODES:
            c = _Case(s, m)
            try:
                if m == "lot_source":
                    with paths.lot_source(True):
                        obs = observe(c)
                else:
                    obs = observe(c)
            finally:
                c.close()
            for reader, v in obs.items():
                out[f"{s}|{m}|{reader}"] = v
    return out


#: 決めた決まりにそろえた所(書き留めとの差。直し)。キー → {"value": 新しい値, "why": 理由}
INTENDED_FILE = Path(__file__).parent / "fixtures" / "lot_ledger_matrix_intended.json"


def intended() -> dict[str, Any]:
    if not INTENDED_FILE.exists():
        return {}
    return {k: v["value"] for k, v in json.loads(INTENDED_FILE.read_text(encoding="utf-8")).items()}


class _Restore(unittest.TestCase):
    """使用パスとモードを試験の前後で戻す。"""

    def setUp(self) -> None:
        self._paths = user_settings.get(paths.KEY_PATHS, None)
        self._mode = user_settings.get(paths.KEY_MODE, None)

    def tearDown(self) -> None:
        for k, v in ((paths.KEY_PATHS, self._paths), (paths.KEY_MODE, self._mode)):
            if v is None:
                user_settings.delete(k)
            else:
                user_settings.save(k, v)
        cache_registry.clear_all()


class LotLedgerMatrixTest(_Restore):
    def test_matrix(self) -> None:
        got = run_matrix()
        if os.environ.get("UFKENSA_RECORD_LEDGER") == "1":
            BASELINE.write_text(json.dumps(got, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")
            self.skipTest("書き留めました")
        base = json.loads(BASELINE.read_text(encoding="utf-8"))
        want = dict(base)
        want.update(intended())
        self.assertEqual(sorted(got), sorted(want))
        diffs = [k for k in sorted(want) if json.loads(json.dumps(got[k], ensure_ascii=False)) != want[k]]
        self.assertEqual(diffs, [], "\n".join(f"{k}\n  want {want[k]}\n  got  {got[k]}" for k in diffs[:20]))

    def test_intended_are_real_changes(self) -> None:
        """直しの一覧は、書き留めにあるキーで、書き留めと違う値であること(古い直しを残さない)。"""
        base = json.loads(BASELINE.read_text(encoding="utf-8"))
        raw = json.loads(INTENDED_FILE.read_text(encoding="utf-8"))
        for k, v in raw.items():
            self.assertIn(k, base)
            self.assertNotEqual(v["value"], base[k], k)
            self.assertTrue(v["why"], k)



class LedgerRulesTest(_Restore):
    """表に出ない所(VBA の文言の出し分け・Test のロット・覚えた表の使い回し)。"""

    def case(self, scenario: str, mode: str) -> _Case:
        c = _Case(scenario, mode)
        self.addCleanup(c.close)
        return c

    def hist(self, **env: Any) -> lot.HistorySaveResult:
        ph = [None] * 19
        ph[0], ph[1], ph[9] = "'06017610", LOT, "KEN"
        return lot.history_save(ph, [list(lot.HISTORY_HEADER)], lot.HistoryEnv(**env))

    def test_sides(self) -> None:
        from ufkensa.kensa import lot_ledger
        c = self.case("both", "online")
        d1, d2 = str(c.d1) + "/", str(c.d2) + "/"
        self.assertEqual(lot_ledger.sides(), lot_ledger.Sides(d1, d2, vba_main=True))
        self.assertEqual(lot_ledger.sides("1"), lot_ledger.Sides(d2, "", explicit=True))
        with paths.lot_source(True):
            self.assertEqual(lot_ledger.sides(), lot_ledger.Sides(d2, d1))
        paths.set_mode(paths.OFFLINE)
        self.assertEqual(lot_ledger.sides(), lot_ledger.Sides(d2, d1))
        self.assertEqual(lot_ledger.sides("1"), lot_ledger.Sides(d2, d1))      # オフラインの DBflag は同じ読み方

    def test_history_test_lot_reads_other_side(self) -> None:
        """直し: Test のロットでも、第1に無ければ第2を見る(前は第2を見る前に黙って抜けていた)。"""
        self.case("lot_only_second", "online")
        r = self.hist(uf1_test=True)
        self.assertTrue(r.completed)
        self.assertEqual(len(r.new_rows), 3)

    def test_history_messages(self) -> None:
        hiki = kensa_fixture.file_name("SIKAHIKINOW")
        nodata = "SIKAHIKINOW.accdbにH2500K0ﾃﾞｰﾀがありません_HistorySave"
        # オンライン: 第1にロットが無く、第2が読めない → 「…がありません」(SaveLoad:1322。前と同じ)
        c = self.case("lot_only_second", "online")
        (c.d2 / hiki).unlink()
        self.assertEqual([m.text for m in self.hist().messages], [paths.missing_text("SIKAHIKINOW")])
        # オフライン: 第2のファイルが無く第1へ移って、ロットが無い → NextPath の文言(直し。前は LOT一覧…)
        c = self.case("lot_absent_both", "offline")
        (c.d2 / hiki).unlink()
        self.assertEqual([m.text for m in self.hist().messages], [nodata])
        # オフライン: 両方にロットが無い → NextPath の文言(前と同じ)
        self.case("lot_absent_both", "offline")
        self.assertEqual([m.text for m in self.hist().messages], [nodata])
        # IsOver12Hours が真なら出さない(LOT一覧… は VBA どおり出す)
        self.case("lot_absent_both", "online")
        self.assertEqual([m.text for m in self.hist(over12=True).messages],
                         ["LOT一覧：LOTNO有_HIKIDB：LOTNO無_HistorySave"])

    def test_tables_read_once(self) -> None:
        """読み分けで表を読み直さない(2 回目からは覚えた表と索引だけ)。"""
        c = self.case("lot_only_second", "online")
        calls: list = []
        real = datasource._read_uncached
        with mock.patch.object(datasource, "_read_uncached", lambda *a: calls.append(a) or real(*a)):
            for _ in range(5):
                lot.arrbox("SIKAHIKINOW", LOT, "Record")
                lot.arrbox("SIKALOTNOW", LOT, "Record")
                lot.history_save([None, LOT] + [None] * 17, [list(lot.HISTORY_HEADER)], lot.HistoryEnv())
        files = sorted({Path(str(a[0])).name for a in calls})
        self.assertLessEqual(len(calls), 6, files)                    # 3 つの台帳 × 第1・第2 を 1 回ずつまで
        self.assertTrue(c.d1.exists())


if __name__ == "__main__":
    unittest.main()
