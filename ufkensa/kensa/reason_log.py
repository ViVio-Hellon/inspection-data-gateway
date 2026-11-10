"""理由の履歴(2026-11-10。現場の声「再測定・その他の理由がどこに残っているか分からない」)

再測定・その他・消去の理由は、VBA と同じく **ライン別の Access.log** に 1 行ずつ追記している
(`legacy_log.line_entry` → `<検査履歴>\\<フォルダ>\\<フォルダ>_Log\\Access.log`)。

    2026/11/10 10:05:31 LINE-PC-HVC  公差外:S1_15.2:理由選択:再測定,キズ, もう一度:検査員01
    2026/11/10 10:06:02 LINE-PC-HVC  U2_15.012:理由選択:その他, 測り直し:検査員01

ここでは**読むだけ**(書式・書き先は変えない。品証などが今の Access.log を読んでいる)。

- 文字コードは cp932・改行 CRLF(`legacy_log`)。読めない文字は置き換えて読む
- 理由の行にはロットが入っていない。アプリの Access.log(使用パス「Access.log などの置き場所」)の
  `Lot登録|<ロット>|<ライン>` のうち、同じ PC でその時刻より前の最後のものを「その時のロット」として添える(推定)
- 読めない・無いときは、理由とフルパスを返す(画面がそのまま出す)
"""
from __future__ import annotations

import re
from dataclasses import asdict, dataclass
from datetime import date, datetime
from typing import Any, Optional

from . import legacy_log, lines, paths, terminal

#: 1 行の頭(`yyyy/mm/dd hh:mm:ss `)
_HEAD = re.compile(r"^(\d{4})/(\d{2})/(\d{2}) (\d{2}):(\d{2}):(\d{2}) ")
REASON_MARK = ":理由選択:"
TOL_PREFIX = "公差外:"
LOT_MARK = "Lot登録|"
#: 画面に返す上限(新しいものから)
MAX_ROWS = 2000


@dataclass
class ReasonEntry:
    at: str            # yyyy/mm/dd hh:mm:ss
    pc: str
    kind: str          # 公差外の確認 / 消去・理由
    position: str      # S1 など
    value: str         # そのときの値(VBA の moto)
    reasons: str       # 再測定, キズ, 自由記入 …
    worker: str
    lot: str = ""      # アプリの Access.log の Lot登録 から推した、その時のロット
    raw: str = ""


def _stamp(line: str) -> Optional[datetime]:
    m = _HEAD.match(line)
    if not m:
        return None
    try:
        return datetime(*map(int, m.groups()))
    except ValueError:
        return None


def _split_head(line: str) -> tuple[Optional[datetime], str, str]:
    """(時刻, PC名, 本文)。書式は `legacy_log.format_log_line`(区切りは "")。"""
    at = _stamp(line)
    if at is None:
        return None, "", ""
    rest = line[20:]
    pc, _, msg = rest.partition(" ")
    return at, pc.strip(), msg.strip()


def parse_reason(line: str) -> Optional[ReasonEntry]:
    """理由選択の 1 行 → ReasonEntry。理由の行でなければ None。"""
    at, pc, msg = _split_head(line.rstrip("\r\n"))
    if at is None or REASON_MARK not in msg:
        return None
    kind = "消去・理由"
    if msg.startswith(TOL_PREFIX):
        kind = "公差外の確認"
        msg = msg[len(TOL_PREFIX):]
    head, _, rest = msg.partition(REASON_MARK)
    reasons, sep, worker = rest.rpartition(":")
    if not sep:
        reasons, worker = rest, ""
    pos, _, value = head.partition("_")
    return ReasonEntry(at=at.strftime("%Y/%m/%d %H:%M:%S"), pc=pc, kind=kind, position=pos.strip(),
                       value=value.strip(), reasons=reasons.strip().strip(","), worker=worker.strip(),
                       raw=line.strip())


def folder_for(line: object = None) -> str:
    """ライン → 検査履歴の下のフォルダ名。未指定はこの端末のライン。"""
    if line in (None, ""):
        ln = terminal.line()
        return ln.folder if ln else legacy_log.DEFAULT_FOLDER
    return legacy_log.log_folder(line)


def files(line: object = None) -> dict[str, str]:
    """ライン別の Access.log / Measured.log と、アプリの Access.log の場所(Windows のパス文字列)。"""
    folder = folder_for(line)
    log_dir = paths.join(paths.get("history"), folder, folder + "_Log")
    return {
        "folder": folder,
        "log_dir": log_dir,
        "access": paths.join(log_dir, legacy_log.ACCESS_LOG),
        "measured": paths.join(log_dir, legacy_log.MEASURED_LOG),
        "app_access": paths.join(paths.write_dir(), legacy_log.ACCESS_LOG),
    }


def _read(path_text: str) -> tuple[Optional[list[str]], str]:
    """(行, 読めないときの理由)。"""
    p = legacy_log.local_path(path_text)
    try:
        data = p.read_bytes()
    except FileNotFoundError:
        return None, "ファイルがありません(まだ理由が 1 件も書かれていないか、フォルダが違います)"
    except OSError as e:
        return None, f"読めません: {e.strerror or e}"
    return data.decode(legacy_log.ENCODING, errors="replace").splitlines(), ""


def _lot_marks(path_text: str) -> list[tuple[datetime, str, str]]:
    """アプリの Access.log の Lot登録(時刻, PC, ロット)。読めなければ空。"""
    rows, _ = _read(path_text)
    out = []
    for line in rows or []:
        at, pc, msg = _split_head(line)
        if at is None or not msg.startswith(LOT_MARK):
            continue
        lot = msg[len(LOT_MARK):].split("|", 1)[0].strip()
        if lot:
            out.append((at, pc, lot))
    out.sort(key=lambda t: t[0])
    return out


def _guess_lot(marks: list[tuple[datetime, str, str]], at: datetime, pc: str) -> str:
    best = ""
    for t, p, lot in marks:
        if t > at:
            break
        if p == pc or not p:
            best = lot
    return best


def _day(text: str) -> Optional[date]:
    text = (text or "").strip().replace("/", "-")
    if not text:
        return None
    try:
        return date.fromisoformat(text)
    except ValueError:
        return None


def read_reasons(line: object = None, *, date_from: str = "", date_to: str = "", lot: str = "") -> dict[str, Any]:
    """理由の履歴(新しい順)。画面にそのまま返す形。"""
    f = files(line)
    out: dict[str, Any] = {**f, "access_local": str(legacy_log.local_path(f["access"])),
                           "entries": [], "total": 0, "error": "", "truncated": False}
    rows, err = _read(f["access"])
    if rows is None:
        out["error"] = err
        return out
    marks = _lot_marks(f["app_access"])
    d0, d1 = _day(date_from), _day(date_to)
    want = (lot or "").strip().upper()
    entries = []
    for line in rows:
        e = parse_reason(line)
        if e is None:
            continue
        at = datetime.strptime(e.at, "%Y/%m/%d %H:%M:%S")
        if (d0 and at.date() < d0) or (d1 and at.date() > d1):
            continue
        e.lot = _guess_lot(marks, at, e.pc)
        if want and want not in e.lot.upper() and want not in e.raw.upper():
            continue
        entries.append(e)
    entries.reverse()                      # 追記なので末尾が新しい
    out["total"] = len(entries)
    out["truncated"] = len(entries) > MAX_ROWS
    out["entries"] = [asdict(e) for e in entries[:MAX_ROWS]]
    out["lot_source"] = bool(marks)
    return out


def line_choices() -> list[dict[str, str]]:
    """選べるライン(フォルダ名が同じものは 1 つ)。"""
    seen, out = set(), []
    for ln in lines.LINES:
        if ln.folder in seen:
            continue
        seen.add(ln.folder)
        out.append({"key": ln.folder, "label": ln.official if ln.folder == ln.official else f"{ln.official}({ln.folder})"})
    return out


def describe_targets() -> dict[str, str]:
    """設定 → 使用パス に出す、いまの書き先(フォルダとファイル名)。"""
    f = files()
    return {"app_access": f["app_access"], "line_access": f["access"], "line_measured": f["measured"],
            "folder": f["folder"]}


__all__ = ["ReasonEntry", "parse_reason", "files", "read_reasons", "line_choices", "describe_targets", "MAX_ROWS"]
