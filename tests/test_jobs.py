"""長時間処理の受け付けと見張り(ufkensa/jobs.py)

* 1 つずつしか走らせない(種類が違っても JobBusy)。走っている間は is_busy・busy_labels に出る
* 進捗: pct は 0〜100 に丸める。文言が変わるたびに段が 1 つ増える。空の文言はバーだけ進める。
  ok=False は開いている段にだけ印を付ける
* 終わり方: 結果の summary()/ok を読む・None・文字列・例外(失敗。最後の段だけ失敗)
* 最近の分は新しい順に max_recent 件まで。find で走っているもの・終わったものを引ける
* ファイルに残し、読み直すと「実行中」だったものは「中断」になる。壊れた記録では落ちない
"""
from __future__ import annotations

import json
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest import mock

import tests  # noqa: F401  ログなどを一時フォルダへ(ufkensa より先に)
from ufkensa import jobs


def _wait(cond, timeout: float = 5.0) -> None:
    end = time.time() + timeout
    while not cond():
        if time.time() > end:
            raise AssertionError("待ちきれませんでした")
        time.sleep(0.01)


class _Result:
    def __init__(self, ok: bool, text: str) -> None:
        self.ok, self._text = ok, text

    def summary(self) -> str:
        return self._text


class _Gate:
    """作業スレッドを、試験が許すまで止めておく。"""

    def __init__(self) -> None:
        self.go = threading.Event()
        self.entered = threading.Event()
        self.progress = None

    def work(self, result=None, steps=()):
        def run(progress):
            self.progress = progress
            for pct, msg, *ok in steps:
                progress(pct, msg, *ok)
            self.entered.set()
            assert self.go.wait(5), "試験が先へ進めませんでした"
            return result
        return run


class _Base(unittest.TestCase):
    def setUp(self) -> None:
        self._td = tempfile.TemporaryDirectory()
        self.addCleanup(self._td.cleanup)
        self.store = Path(self._td.name) / "cache" / "jobs.json"
        self.reg = jobs.JobRegistry(store_path=self.store)

    def finish(self, job: jobs.Job) -> jobs.Job:
        _wait(lambda: not self.reg.is_busy())
        found = self.reg.find(job.id)
        self.assertIsNotNone(found)
        return found


class Running(_Base):
    def test_one_at_a_time_even_for_other_kinds(self) -> None:
        gate = _Gate()
        job = self.reg.start("import", "まとめて取り込み", gate.work())
        self.assertTrue(job.is_running)
        self.assertEqual(job.state, jobs.STATE_RUNNING)
        self.assertEqual(job.message, "まとめて取り込みを準備しています")
        self.assertEqual(len(job.id), 12)
        self.assertTrue(self.reg.is_busy())
        self.assertEqual(self.reg.busy_labels(), ["まとめて取り込み"])
        self.assertIs(self.reg.running(), job)
        self.assertIs(self.reg.find(job.id), job)
        with self.assertRaises(jobs.JobBusy) as cm:
            self.reg.start("write_back", "書き戻し", lambda p: None)
        self.assertIs(cm.exception.running, job)
        self.assertIn("まとめて取り込み", str(cm.exception))
        # 走っている間もファイルに残っている(落ちても最後が読める)
        saved = json.loads(self.store.read_text(encoding="utf-8"))
        self.assertEqual(saved["running"]["id"], job.id)
        gate.go.set()
        done = self.finish(job)
        self.assertEqual(done.state, jobs.STATE_DONE)
        self.assertEqual(self.reg.busy_labels(), [])
        self.assertIsNone(self.reg.running())
        # 終われば次を受け付ける
        nxt = self.reg.start("write_back", "書き戻し", lambda p: None)
        self.finish(nxt)
        self.assertEqual([j.id for j in self.reg.recent()], [nxt.id, job.id])

    def test_progress_steps(self) -> None:
        gate = _Gate()
        job = self.reg.start("import", "取り込み", gate.work(steps=[
            (5, "マスタを読み込んでいます"),
            (12, ""),                                   # バーだけ進む(文言・段はそのまま)
            (20, "マスタを読み込んでいます"),           # 同じ文言は同じ段
            (40, "仕掛ロットを読み込んでいます"),
            (45, "仕掛ロットを読み込んでいます", False),  # いまの段で問題
            (70, "丈補填を読み込んでいます"),
            (150, ""),                                  # 100 を超えない
        ]))
        _wait(gate.entered.is_set)
        self.assertEqual(job.pct, 100)
        self.assertEqual(job.message, "丈補填を読み込んでいます")
        self.assertEqual([s.label for s in job.steps],
                         ["マスタを読み込んでいます", "仕掛ロットを読み込んでいます", "丈補填を読み込んでいます"])
        self.assertEqual([s.ok for s in job.steps], [True, False, True])
        self.assertEqual([s.running for s in job.steps], [False, False, True])
        gate.progress(-5, "")
        self.assertEqual(job.pct, 0)                    # 0 を割らない
        gate.go.set()
        done = self.finish(job)
        self.assertEqual(done.state, jobs.STATE_DONE)
        self.assertTrue(done.ok)
        self.assertEqual(done.pct, 100)
        self.assertEqual(done.message, "")
        self.assertEqual([s.ok for s in done.steps], [True, False, True], "全体の結果を最後の段に被せない")
        self.assertTrue(all(not s.running for s in done.steps))
        self.assertGreaterEqual(done.finished_at, done.started_at)
        self.assertGreaterEqual(done.elapsed_sec, 0.0)
        for s in done.steps:
            self.assertGreaterEqual(s.elapsed_sec, 0.0)

    def test_failed_step_mark_is_ignored_without_an_open_step(self) -> None:
        def work(progress):
            progress(10, "", False)                     # 段がまだ無い
        done = self.finish(self.reg.start("import", "取り込み", work))
        self.assertEqual(done.steps, [])
        self.assertEqual(done.state, jobs.STATE_DONE)


class Endings(_Base):
    def test_result_with_summary_and_ok(self) -> None:
        done = self.finish(self.reg.start("import", "取り込み", lambda p: _Result(True, "19表 1,234行")))
        self.assertEqual((done.state, done.ok, done.summary, done.error), (jobs.STATE_DONE, True, "19表 1,234行", ""))
        bad = self.finish(self.reg.start("import", "取り込み", lambda p: (p(30, "マスタ"), _Result(False, "2表失敗"))[1]))
        self.assertEqual((bad.state, bad.ok, bad.summary), (jobs.STATE_FAILED, False, "2表失敗"))
        self.assertEqual(bad.pct, 30, "失敗ならバーは止まったところのまま")
        self.assertTrue(bad.steps[-1].ok, "結果が失敗でも、段の失敗は段自身が報告したものだけ")

    def test_plain_results(self) -> None:
        self.assertEqual(self.finish(self.reg.start("k", "a", lambda p: None)).summary, "")
        self.assertEqual(self.finish(self.reg.start("k", "b", lambda p: 42)).summary, "42")
        self.assertTrue(self.finish(self.reg.start("k", "c", lambda p: "済み")).ok)

    def test_exception_marks_failed_and_last_step(self) -> None:
        def work(progress):
            progress(10, "マスタ")
            progress(50, "仕掛")
            raise ValueError("仕掛台帳が読めません")
        with mock.patch("ufkensa.trace.event") as ev:
            done = self.finish(self.reg.start("import", "取り込み", work))
        self.assertEqual(done.state, jobs.STATE_FAILED)
        self.assertFalse(done.ok)
        self.assertEqual(done.error, "仕掛台帳が読めません")
        self.assertEqual(done.summary, "")
        self.assertEqual(done.pct, 50)
        self.assertEqual([s.ok for s in done.steps], [True, False])
        self.assertEqual(ev.call_args[0][0], "job.failed")
        self.assertFalse(self.reg.is_busy(), "落ちても次を受け付ける")


class Recent(_Base):
    def test_newest_first_and_capped(self) -> None:
        reg = jobs.JobRegistry(store_path=self.store, max_recent=3)
        ids = []
        for i in range(5):
            job = reg.start("k", f"処理{i}", lambda p: None)
            _wait(lambda: not reg.is_busy())
            ids.append(job.id)
        self.assertEqual([j.id for j in reg.recent()], ids[::-1][:3])
        self.assertIsNone(reg.find(ids[0]), "押し出されたものは引けない")
        self.assertIsNone(reg.find("無い"))
        reg.recent().clear()                            # 写しを返す(中身を壊されない)
        self.assertEqual(len(reg.recent()), 3)


class Persistence(_Base):
    def test_restart_turns_running_into_interrupted(self) -> None:
        done = self.finish(self.reg.start("import", "前の取り込み", lambda p: "済み"))
        gate = _Gate()
        job = self.reg.start("write_back", "書き戻し", gate.work(steps=[(10, "送っています")]))
        _wait(gate.entered.is_set)
        # このプロセスが落ちた体で、同じ記録から作り直す
        again = jobs.JobRegistry(store_path=self.store)
        self.assertIsNone(again.running())
        self.assertFalse(again.is_busy())
        got = again.recent()
        self.assertEqual([j.id for j in got], [job.id, done.id])
        stale = got[0]
        self.assertEqual(stale.state, jobs.STATE_INTERRUPTED)
        self.assertFalse(stale.is_running)
        self.assertEqual(stale.message, "")
        self.assertIn("途中で止まっています", stale.error)
        self.assertIsInstance(stale.steps[0], jobs.Step, "段は辞書でなく Step に組み直す")
        self.assertEqual(stale.steps[0].label, "送っています")
        self.assertEqual(got[1].summary, "済み")
        gate.go.set()
        _wait(lambda: not self.reg.is_busy())

    def test_broken_or_odd_store_does_not_break_startup(self) -> None:
        self.store.parent.mkdir(parents=True, exist_ok=True)
        self.store.write_text("{壊れている", encoding="utf-8")
        reg = jobs.JobRegistry(store_path=self.store)
        self.assertEqual(reg.recent(), [])
        self.store.write_text(json.dumps({
            "running": "文字列",
            "recent": [
                {"id": "a1", "kind": "import", "label": "古い形", "state": "done", "未知の欄": 1,
                 "steps": [{"label": "段", "started_at": 1.0, "finished_at": 2.0, "ok": True, "余計": 0}, "壊れた段"]},
                {"kind": "import"},                     # id も label も無い → 捨てる
                "行でない",
            ]}), encoding="utf-8")
        reg = jobs.JobRegistry(store_path=self.store)
        got = reg.recent()
        self.assertEqual([j.id for j in got], ["a1"])
        self.assertEqual([s.label for s in got[0].steps], ["段"])
        self.assertAlmostEqual(got[0].steps[0].elapsed_sec, 1.0)
        self.assertIsNone(reg.running())

    def test_save_failure_does_not_stop_the_work(self) -> None:
        blocker = Path(self._td.name) / "ファイル"
        blocker.write_text("x", encoding="utf-8")
        reg = jobs.JobRegistry(store_path=blocker / "jobs.json")     # 親がファイル → 書けない
        job = reg.start("import", "取り込み", lambda p: (p(10, "読む"), "済み")[1])
        _wait(lambda: not reg.is_busy())
        done = reg.find(job.id)
        self.assertEqual(done.state, jobs.STATE_DONE)

    def test_default_store_is_the_local_cache(self) -> None:
        with mock.patch.object(jobs.app_config, "local_dir", return_value=Path(self._td.name) / "local") as ld:
            reg = jobs.JobRegistry()
            self.assertEqual(reg._store_path(), Path(self._td.name) / "local" / "jobs.json")
        ld.assert_called_with("cache")


class ProcessWide(unittest.TestCase):
    def test_get_and_reset_registry(self) -> None:
        saved = jobs._registry
        self.addCleanup(setattr, jobs, "_registry", saved)
        with tempfile.TemporaryDirectory() as td:
            mine = jobs.JobRegistry(store_path=Path(td) / "jobs.json")
            self.assertIs(jobs.reset_registry(mine), mine)
            self.assertIs(jobs.get_registry(), mine)
            self.assertIs(jobs.get_registry(), mine)


if __name__ == "__main__":
    unittest.main()
