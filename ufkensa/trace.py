"""操作の記録(後から「なぜ?」をたどるための記録)

【なぜ要るか】
これまでの記録(`ufkensa_*.log`)は開発者向けの文で、「誰が・どの画面で・何を押して・
何が返り・その前に何があったか」を後から組み立てられなかった。断り(422/409)は
1 行も残らず、思わぬエラーは Flask の中に消えていた。

【何を残すか】1 行 = 1 つの出来事(JSON)。ファイルは `ufkensa\\ufkensa_記録_<PC名>_YYYYMMDD.jsonl`(0.33.0 から。ツール名のフォルダ・ツール名で始まる名前)。
共有フォルダに置いても PC ごとにファイルが分かれる(同じファイルを取り合わない)。

    t     時刻(ミリ秒まで)        lv    INFO / WARN / ERROR
    ev    出来事の種類(下の表)    msg   画面に出した文言・一言
    rid   記録番号(要求ごと。画面のエラーにも出る)
    sid   画面(タブ)の番号。同じ画面の前後をたどる鍵
    user  ログイン ID          lvl   Level      pc   PC 名     ver   版
    path  画面・操作の URL     method / st(状態) / ms(かかった時間)
    data  中身(直したマスの 前→後、断りの理由コード、送った値の要約 …)
    exc   例外(種類・文言・トレースバック)

    screen         画面を開いた
    request        書く操作・断られた操作(st が 4xx/5xx なら断り・エラー)
    error.server   思わぬエラー(トレースバックつき)
    error.client   ブラウザの中のエラー(JavaScript)
    login.ok / login.ng, cell.save, record.save, master.save / master.delete,
    settings.save, db.open, accdb.convert, export, print, report, job.done / job.failed,
    app.start / app.stop

パスワードなど、名前に pass / token を含む値は「***」にして残す。
**書けなくても業務は止めない。**

【0.11.0】(ViVio-Hellon/vba-daily-report-python-migration の「ログ」にそろえた)
- 設定の出力先に書けないときは、この PC の既定の場所に書き、**5 分ごとに設定の出力先を試し直す**
  (起動し直さなくても戻る)。読むときは両方を読む(書けなかったあいだのぶんも一覧に並ぶ)
- コードの中で `log.warning` / `log.exception` に書いたもの(握りつぶしていたエラー)も拾う(`install_capture`)。
  同じものが続くときは 1 分に 1 度だけ、間引いた数を添える
"""
from __future__ import annotations

import json
import os
import re
import socket
import threading
import traceback
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable, Optional

from . import config

_lock = threading.Lock()
PREFIX = f"{config.TOOL_LOG_NAME}_記録_"        # 0.33.0 からツール名で始める
LEGACY_PREFIX = "記録_"                          # 0.32.0 まで(前の場所の自分の行だけ読む)
SUFFIX = ".jsonl"
SECRET = re.compile(r"pass|token|secret", re.IGNORECASE)
# 文の中の URL に付いた起動トークン(JavaScript のスタックに画面の URL が出る)
TOKEN_IN_URL = re.compile(r"([?&]t=)[^&\s#:'\"]+")
MAX_TEXT = 400
MAX_ITEMS = 20
# 問題の出来事(「問題だけ」で出すもの)
PROBLEM_LEVELS = ("WARN", "ERROR")


def pc_name() -> str:
    name = os.environ.get("COMPUTERNAME") or socket.gethostname() or "PC"
    return re.sub(r'[\\/:*?"<>|\s]+', "_", name)[:40]


def file_for(day: Optional[date] = None, pc: Optional[str] = None) -> Path:
    return config.LOG_DIR / f"{PREFIX}{pc or pc_name()}_{(day or date.today()):%Y%m%d}{SUFFIX}"


def read_dirs() -> list[Path]:
    """読む場所: いまの出力先(ツール名のフォルダ)・この PC の既定の場所(書けなかったあいだのぶん)と、
    **前の場所**(0.32.0 まで書いていた、それぞれの 1 つ上 = 置き場所の直下)。

    前の場所はほかのツールと同じ所なので、**自分の名前のファイルの、自分の行だけ**を読む(`_ours`)。
    ほかのツールのフォルダ(nippou\\ など)の中は見ない。"""
    dirs = [config.LOG_DIR]
    for d in (config.DEFAULT_LOG_DIR, *legacy_dirs()):
        if d not in dirs and d.is_dir():
            dirs.append(d)
    return dirs


def legacy_dirs() -> list[Path]:
    """前の場所(0.32.0 まで): いまの出力先・既定の場所・設定の置き場所の、ツール名のフォルダの 1 つ上。"""
    out: list[Path] = []
    cands = [config.LOG_DIR, config.DEFAULT_LOG_DIR]
    try:
        from . import logging_utils, user_settings
        text = str(user_settings.get(config.KEY_LOG_DIR, "") or "").strip()
        if text:
            cands.append(logging_utils.resolve_log_dir(text))
    except Exception:                                       # noqa: BLE001 - 読めなければいまの場所だけ
        pass
    for d in cands:
        if d.name == config.TOOL_LOG_NAME and d.parent not in out:
            out.append(d.parent)
    return out


def _ours(rec: Any) -> bool:
    """自分(UF検査判定)の記録の行か。前の場所のファイルには、ほかのツールの行が混ざることがある。"""
    return isinstance(rec, dict) and "ev" in rec and rec.get("app") in (None, config.TOOL_LOG_NAME)


RETRY_SECONDS = 300
_retry = {"at": 0.0}


def _maybe_retry() -> None:
    """設定の出力先に書けず既定の場所に書いているとき、5 分ごとに設定の出力先を試し直す。"""
    import time
    from . import logging_utils, user_settings
    now = time.monotonic()
    if now - _retry["at"] < RETRY_SECONDS:
        return
    _retry["at"] = now
    text = str(user_settings.get(config.KEY_LOG_DIR, "") or "").strip()
    if not text:
        return
    wanted = logging_utils.resolve_log_dir(text)
    if wanted != config.LOG_DIR and not logging_utils.set_log_dir(wanted):
        logging_utils.get_logger("ufkensa.trace").info("ログの出力先に書けるようになりました: %s", wanted)


def summarize(value: Any, depth: int = 0) -> Any:
    """送った値などを、記録に残せる大きさに縮める。秘密は伏せる。"""
    if isinstance(value, dict):
        out = {}
        for i, (k, v) in enumerate(value.items()):
            if i >= MAX_ITEMS * 2:
                out["…"] = f"ほか {len(value) - i} 項目"
                break
            out[str(k)] = "***" if SECRET.search(str(k)) else summarize(v, depth + 1)
        return out
    if isinstance(value, (list, tuple)):
        if depth >= 2 or len(value) > MAX_ITEMS:
            head = [summarize(v, depth + 1) for v in list(value)[:3]] if depth < 2 else []
            return {"件数": len(value), "先頭": head} if head else {"件数": len(value)}
        return [summarize(v, depth + 1) for v in value]
    if isinstance(value, str):
        value = TOKEN_IN_URL.sub(r"\1***", value)
        return value if len(value) <= MAX_TEXT else value[:MAX_TEXT] + f"…(全 {len(value)} 字)"
    if value is None or isinstance(value, (bool, int, float)):
        return value
    return summarize(str(value), depth)


def clean_path(text: str) -> str:
    """URL から起動トークン(`t=`)を外す。**記録に鍵を残さない**(共有フォルダに置くこともある)。"""
    text = str(text or "")
    if "?" not in text:
        return text
    head, _, query = text.partition("?")
    kept = [part for part in query.split("&") if part and not part.startswith("t=")]
    return head + ("?" + "&".join(kept) if kept else "")


def mask_by_columns(columns: Optional[list[str]], values: Any) -> Any:
    """列と同じ並びで来た値のうち、秘密の列(名前に pass など。ID 表の PASS)を「***」にする。

    マスタ管理は値を**名前の無い並び**で送るので、`summarize` の名前による伏せ字が効かない。
    列が分からないときは全部伏せる(残さないほうへ倒す)。
    """
    if not isinstance(values, (list, tuple)):
        return values
    if columns is None:
        return ["***"] * len(values)
    return ["***" if i < len(columns) and SECRET.search(str(columns[i])) else v for i, v in enumerate(values)]


def sent_body() -> Any:
    """記録に残す「送った値」。経路が伏せ字にしたもの(`g.trace_sent`)があればそちら。"""
    try:
        from flask import g, request
        if "trace_sent" in g:
            return g.trace_sent
        return request.get_json(silent=True) if request.is_json else None
    except Exception:                                  # noqa: BLE001
        return None


# ------------------------------------------------------------------
# 処理の開始・終了と処理時間(ログの決まり 3)・エラー時の詳細(ログの決まり 2)
# ------------------------------------------------------------------
import contextvars
import threading as _threading
import time as _time
from collections import deque
from contextlib import contextmanager

#: 処理時間 → 残す重さ。通常 INFO、3 秒以上 / 10 秒以上 / 30 秒以上 は WARN(区切りを label に)
SLOW_STEPS: tuple[tuple[float, str], ...] = ((30.0, "30秒以上"), (10.0, "10秒以上"), (3.0, "3秒以上"))
#: 終わらない処理の見張り。この秒数を過ぎるたびに「まだ終わっていません」を残す(10分以上の不具合を見つける)
STILL_RUNNING_AT: tuple[float, ...] = (30.0, 60.0, 120.0, 300.0, 600.0, 1200.0, 1800.0)

_procs: contextvars.ContextVar[tuple[str, ...]] = contextvars.ContextVar("ufkensa_procs", default=())
_used: deque = deque(maxlen=12)                 # 最近使ったファイル・DB(エラーのときに残す)
_recent: deque = deque(maxlen=15)               # 直前の出来事(エラーのときに残す)


def note_used(path: Any, what: str = "") -> None:
    """使ったファイル・DB を覚える(エラーが起きたら「使っていたもの」として残す)。"""
    try:
        item = f"{what}:{clean_path(str(path))}" if what else clean_path(str(path))
        if not _used or _used[-1] != item:
            _used.append(item)
    except Exception:                                  # noqa: BLE001
        pass


def slow_label(sec: float) -> str:
    for limit, label in SLOW_STEPS:
        if sec >= limit:
            return label
    return ""


def _text_log(level: str, fmt: str, *args: Any) -> None:
    """テキストのログ(ufkensa_<PC>_日付.log)にも 1 行。記録(jsonl)へは戻さない(no_trace)。"""
    try:
        import logging
        getattr(logging.getLogger("ufkensa.proc"), level)(fmt, *args, extra={"no_trace": True})
    except Exception:                                  # noqa: BLE001
        pass


@contextmanager
def timed(name: str, msg: str = "", *, expected: tuple[type, ...] = (), **data: Any):
    """処理の開始と終了をセットで残す。処理時間で INFO / WARN、例外なら ERROR(詳細つき)を残して投げ直す。

        with trace.timed("db.read", "品質基準 を読む", path=p, table="ソリ3級") as info:
            ...
            info["rows"] = len(rows)      # 終わりの行に足したい値
    """
    stack = _procs.get() + (name,)
    token = _procs.set(stack)
    info: dict[str, Any] = dict(data)
    t0 = _time.perf_counter()
    event("proc.start", msg or name, proc=name, data=info or None)
    done = _threading.Event()

    def watch(i: int = 0) -> None:
        if i >= len(STILL_RUNNING_AT) or done.is_set():
            return
        wait = STILL_RUNNING_AT[i] - (STILL_RUNNING_AT[i - 1] if i else 0)
        if done.wait(wait):
            return
        event("proc.slow", f"{msg or name}: まだ終わっていません({STILL_RUNNING_AT[i]:.0f}秒経過)",
              level="WARN", proc=name, ms=round((_time.perf_counter() - t0) * 1000))
        watch(i + 1)

    watcher = _threading.Thread(target=watch, daemon=True, name=f"timed-{name}")
    watcher.start()
    try:
        yield info
    except expected as exc:                          # 業務上の中断(VBA の MsgError・End に当たる)。エラーではない
        ms = round((_time.perf_counter() - t0) * 1000)
        event("proc.end", f"{msg or name}: 中断しました({exc})", level="WARN", proc=name, ms=ms,
              stopped=type(exc).__name__, data=info or None)
        _text_log("warning", "処理 %s を中断しました(%d ms): %s", name, ms, exc)
        raise
    except Exception as exc:
        ms = round((_time.perf_counter() - t0) * 1000)
        event("proc.error", f"{msg or name}: {exc}", level="ERROR", proc=name, ms=ms, exc=exc, data=info or None)
        _text_log("error", "処理 %s が失敗しました(%d ms): %s", name, ms, exc)
        raise
    else:
        ms = round((_time.perf_counter() - t0) * 1000)
        label = slow_label(ms / 1000)
        event("proc.end", f"{msg or name}" + (f"({label})" if label else ""), level="WARN" if label else "INFO",
              proc=name, ms=ms, slow=label or None, data=info or None)
        _text_log("warning" if label else "info", "処理 %s: %d ms%s %s", name, ms, f"({label})" if label else "",
                  summarize(info) if info else "")
    finally:
        done.set()
        _procs.reset(token)


def _context() -> dict[str, Any]:
    """いまの要求(あれば)とログインから分かること。"""
    out: dict[str, Any] = {}
    try:
        from flask import g, has_request_context, request
        if has_request_context():
            out["rid"] = g.get("rid", "")
            out["sid"] = request.headers.get("X-Screen-Id", "") or ""
            out["path"] = clean_path(request.full_path.rstrip("?")) if request.query_string else request.path
            out["method"] = request.method
    except Exception:                                  # noqa: BLE001
        pass
    try:
        from . import auth
        s = auth.current()
        if s is not None:
            out["user"] = s.user_id
            out["lvl"] = s.level
    except Exception:                                  # noqa: BLE001
        pass
    return out


def _version() -> str:
    try:
        from . import app_config
        return app_config.version()
    except Exception:                                  # noqa: BLE001
        return ""


def event(ev: str, msg: str = "", *, level: str = "INFO", data: Any = None,
          exc: Optional[BaseException] = None, **fields: Any) -> dict[str, Any]:
    """出来事を 1 行残す。戻り値は残した中身(試験用)。**例外は出さない。**"""
    rec: dict[str, Any] = {"t": datetime.now().isoformat(timespec="milliseconds"), "lv": level, "ev": ev}
    try:
        rec.update(_context())
        rec.update({k: v for k, v in fields.items() if v is not None})
        rec["pc"] = pc_name()
        rec["app"] = config.TOOL_LOG_NAME                 # 中身からもどのツールの記録か分かるように(0.33.0)
        rec["ver"] = _version()
        if msg:
            rec["msg"] = summarize(str(msg))
        if data is not None:
            rec["data"] = summarize(data)
        procs = _procs.get()
        if procs and "proc" not in rec:
            rec["proc"] = procs[-1]
        if exc is not None or level == "ERROR":
            # エラーのときに必ず残すもの: 実行していた処理・使っていたファイル/DB・直前の処理状況
            rec["procs"] = list(procs)
            rec["used"] = list(_used)
            rec["recent"] = list(_recent)
        if exc is not None:
            rec["exc"] = {"type": type(exc).__name__, "msg": summarize(str(exc)),
                          "tb": traceback.format_exception(type(exc), exc, exc.__traceback__)[-30:]}
        line = json.dumps(rec, ensure_ascii=False, default=str)
        _recent.append(f'{rec["t"][11:19]} {ev} {str(rec.get("msg", ""))[:80]}')
        _maybe_retry()
        path = file_for()
        with _lock:
            path.parent.mkdir(parents=True, exist_ok=True)
            with open(path, "a", encoding="utf-8", newline="\n") as f:
                f.write(line + "\n")
    except Exception:                                  # noqa: BLE001 - 記録のせいで業務を止めない
        pass
    return rec


# ------------------------------------------------------------------
# 読む(設定「ログ」の画面)
# ------------------------------------------------------------------
_NAME = re.compile(rf"^(?:{re.escape(PREFIX)}|{re.escape(LEGACY_PREFIX)})(?P<pc>.+)_(?P<day>\d{{8}}){re.escape(SUFFIX)}$")


def pc_of(path: Path) -> str:
    """記録ファイルの名前から PC 名。"""
    m = _NAME.match(path.name)
    return m.group("pc") if m else ""


def files(days: int = 7, folder: Optional[Path] = None) -> list[Path]:
    """直近 `days` 日の記録ファイル(この PC 以外のものも。共有フォルダに置いたとき)。"""
    since = (date.today() - timedelta(days=max(1, days) - 1)).strftime("%Y%m%d")
    out: dict[str, Path] = {}
    for d in ([folder] if folder else read_dirs()):
        try:
            for p in d.iterdir():
                m = _NAME.match(p.name)
                if m and m.group("day") >= since:
                    out.setdefault(str(p.resolve()), p)
        except OSError:
            continue
    return sorted(out.values(), key=lambda p: p.name)


def _lines(paths: Iterable[Path]) -> list[dict[str, Any]]:
    out = []
    for p in paths:
        try:
            with open(p, encoding="utf-8", errors="replace") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        rec = json.loads(line)
                    except ValueError:
                        continue
                    if _ours(rec):
                        out.append(rec)
        except OSError:
            continue
    out.sort(key=lambda r: str(r.get("t", "")))
    return out


def read(days: int = 7, *, only_problems: bool = False, q: str = "", limit: int = 300) -> list[dict[str, Any]]:
    """新しい順。`q` は記録番号・利用者・PC・種類・文言・中身のどれかに含まれる言葉。"""
    recs = _lines(files(days))
    if only_problems:
        recs = [r for r in recs if r.get("lv") in PROBLEM_LEVELS]
    if q:
        words = [w.lower() for w in q.split() if w]
        recs = [r for r in recs
                if all(w in json.dumps(r, ensure_ascii=False).lower() for w in words)]
    return list(reversed(recs))[:max(1, limit)]


def chain(rid: str, *, minutes: int = 15, days: int = 30) -> dict[str, Any]:
    """記録番号の出来事と、その前後(同じ画面・同じ PC の ±minutes 分)。

    「なぜ?」をたどるための並び: いちばん上が古い。どの画面で、何を押し、何が返って、
    その前に何が起きていたか(ほかの端末の書き込みで競合した、など)が一続きで読める。
    """
    recs = _lines(files(days))
    hits = [r for r in recs if r.get("rid") == rid]
    if not hits:
        return {"rid": rid, "found": False, "events": []}
    first = hits[0]
    t0 = datetime.fromisoformat(str(first["t"]))
    lo, hi = t0 - timedelta(minutes=minutes), t0 + timedelta(minutes=minutes)
    sid, pc = first.get("sid") or "", first.get("pc") or ""

    def near(r: dict[str, Any]) -> bool:
        try:
            t = datetime.fromisoformat(str(r.get("t")))
        except ValueError:
            return False
        if not (lo <= t <= hi):
            return False
        if r.get("rid") == rid:
            return True
        if r.get("pc") != pc:
            return False
        # 同じ画面の出来事、画面の無い出来事(長時間処理・起動終了)
        return (sid and r.get("sid") == sid) or not r.get("sid")

    return {"rid": rid, "found": True, "center": first, "events": [r for r in recs if near(r)]}


def cleanup(keep_days: int) -> int:
    """この PC の古い記録を消す(記録・ufkensa_*.log・取り込み診断)。消した数。"""
    if keep_days <= 0:
        return 0
    limit = (date.today() - timedelta(days=keep_days)).strftime("%Y%m%d")
    me = pc_name()
    tool = re.escape(config.TOOL_LOG_NAME)
    pats = [re.compile(rf"^(?:{re.escape(PREFIX)}|{re.escape(LEGACY_PREFIX)}){re.escape(me)}_(\d{{8}}){re.escape(SUFFIX)}$"),
            re.compile(rf"^{tool}_{re.escape(me)}_(\d{{8}})\.log$"),
            re.compile(rf"^{tool}_(\d{{8}})\.log$"),
            re.compile(rf"^(?:{tool}_)?取り込み診断_{re.escape(me)}_(\d{{8}})\.log$")]
    gone = 0
    # いまの置き場所・既定の場所・前の場所(直下)の、**この PC の・自分の名前の**ものだけ。
    # 前の場所の 記録_… は、中の行が全部自分のものと分かったときだけ消す(ほかのツールの行が混ざっていれば残す)
    legacy = set(legacy_dirs())
    for folder in read_dirs():
        try:
            entries = list(folder.iterdir())
        except OSError:
            continue
        for p in entries:
            if not p.is_file():
                continue                                       # ほかのツールのフォルダ(nippou\ など)は見ない
            for pat in pats:
                m = pat.match(p.name)
                if m and m.group(1) < limit:
                    if folder in legacy and p.name.startswith(LEGACY_PREFIX) and not _all_ours(p):
                        break
                    try:
                        p.unlink()
                        gone += 1
                    except OSError:
                        pass
                    break
    return gone


def _all_ours(path: Path) -> bool:
    """ファイルの行が全部自分のものか(読めなければ「分からない」= 消さない)。"""
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    if not _ours(json.loads(line)):
                        return False
                except ValueError:
                    return False
    except OSError:
        return False
    return True


# ------------------------------------------------------------------
# log.warning / log.exception もログに拾う(0.11.0)
# ------------------------------------------------------------------
import logging  # noqa: E402

# 拾わない記録(通信の部品・ログ自身)
_SKIP_LOGGERS = ("waitress", "werkzeug", "urllib3", "ufkensa.trace")
REPEAT_SECONDS = 60


class CaptureHandler(logging.Handler):
    """WARNING 以上を 1 行の出来事にする。同じもの(記録の名前 + 文の型)は 1 分に 1 度、間引いた数を添える。"""

    def __init__(self) -> None:
        super().__init__(level=logging.WARNING)
        self._last: dict[tuple[str, str], list[float]] = {}

    def emit(self, record: logging.LogRecord) -> None:
        try:
            if record.name.startswith(_SKIP_LOGGERS) or getattr(record, "no_trace", False):
                return
            import time
            key = (record.name, str(record.msg))
            now = time.monotonic()
            last = self._last.get(key)
            if last and now - last[0] < REPEAT_SECONDS:
                last[1] += 1
                return
            skipped = int(last[1]) if last else 0
            self._last[key] = [now, 0]
            error = record.levelno >= logging.ERROR
            data: dict[str, Any] = {"logger": record.name, "where": f"{Path(record.pathname).name}:{record.lineno}"}
            if skipped:
                data["間引いた数"] = skipped
            exc = record.exc_info[1] if record.exc_info and record.exc_info[1] else None
            event("log.error" if error else "log.warning", record.getMessage(),
                  level="ERROR" if error else "WARN", data=data, exc=exc)
        except Exception:                               # noqa: BLE001 - ログのせいで業務を止めない
            pass


def install_capture() -> None:
    """`ufkensa` の記録に拾う口を付ける(何度呼んでも 1 つ)。"""
    root = logging.getLogger("ufkensa")
    if not any(isinstance(h, CaptureHandler) for h in root.handlers):
        root.addHandler(CaptureHandler())
