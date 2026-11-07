"""VBA と同じ場所・同じ書式のログ(Access.log / Measured.log)

VBA `【UFKen_16】WriteLog.WriteLog` と、`【UFKen_17】補助` の `LogOut` / `LogOutCombo` の移植。
検査履歴のフォルダ(HistoryPath)にラインごとのログを追記する。品証などが今の
Access.log を読んでいるので、**場所・ファイル名・書式・文字コードを VBA と同じ**にする。

- 場所: `<HistoryPath>\\<フォルダ>\\<フォルダ>_Log\\Access.log`(Measured.log)。
  フォルダは `lines.find(ライン).folder`。見つからないライン(丸徳+号機 など)は VBA の Case Else と同じく「丸徳」
- 書式: `yyyy/mm/dd hh:mm:ss` + " " + Sepl + PC名 + " " + Sepl + " " + 本文 + " "(WriteLog:34)
- 文字コード: FSO の OpenTextFile 既定(ANSI = cp932)、改行 CRLF
- 新しく作ったファイルの 1 行目は `管理ﾌｧｲﾙ使用状況`(WriteLog:29)
- VBA は `On Error Resume Next` なので、書けなくても処理は止めない(戻り値 False)
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

from . import lines, paths

ACCESS_LOG_TITLE = "管理ﾌｧｲﾙ使用状況"      # WriteLog.AccessLogTitle
ACCESS_LOG = "Access.log"
MEASURED_LOG = "Measured.log"
ENCODING = "cp932"                          # FSO の TristateFalse(ANSI)
NEWLINE = "\r\n"                            # TextStream.WriteLine
DEFAULT_FOLDER = "丸徳"                     # Select Case の Case Else
#: cp932 で書けない文字の扱い。FSO の実挙動は未確認(不明点)。"replace" は "?" にする
ENCODE_ERRORS = "replace"


def local_path(path_text: str) -> Path:
    """Windows のパス文字列 → この OS の Path。

    `paths.to_local` に加え、Windows 以外では `\\` を区切りとして読む
    (オフラインの使用パスが "/tmp/x/" のとき、VBA の `path & "L-1\\L-1_Log"` を正しく辿るため)。
    共通化の候補(paths.to_local に入れてよい)。
    """
    p = paths.to_local(path_text)
    if os.sep == "/":
        p = Path(str(p).replace("\\", "/"))
    return p


def log_folder(line: object) -> str:
    """ライン名 → 検査履歴の下のフォルダ名(L-1 / LVC / HVC / バランサー / トット / Test / 丸徳)。

    VBA の `Select Case Lname: Case "L-1" … Case Else: 丸徳`(補助:1153-1161 ほか多数)。
    """
    ln = lines.find(line)
    return ln.folder if ln else DEFAULT_FOLDER


def line_log_dir(line: object) -> str:
    """`HistoryPath & "<フォルダ>\\<フォルダ>_Log"`(Windows のパス文字列)。"""
    folder = log_folder(line)
    return paths.join(paths.get("history"), folder, folder + "_Log")


def format_log_line(msg1: str, sepl: str = ",", *, now: Optional[datetime] = None,
                    computer: Optional[str] = None) -> str:
    """WriteLog:34 の 1 行(改行なし)。

    `Format(Now, "YYYY/MM/DD hh:mm:ss") & " " & Sepl & compName & " " & Sepl & " " & msg1 & " "`
    """
    now = now or datetime.now()
    comp = os.environ.get("COMPUTERNAME", "") if computer is None else computer
    stamp = now.strftime("%Y/%m/%d %H:%M:%S")
    return stamp + " " + sepl + comp + " " + sepl + " " + str(msg1) + " "


def _create_folder(path_text: str) -> None:
    """`CreateFolder`(Input:1163)。エラーは無視する。"""
    try:
        local_path(path_text).mkdir(parents=True, exist_ok=True)
    except OSError:
        pass


def write_log(msg1: str, sepl: str = ",", target_path: str = "", target_file: str = "", *,
              now: Optional[datetime] = None, computer: Optional[str] = None) -> bool:
    """WriteLog.WriteLog (WriteLog:6-39)。書けたら True(VBA は On Error Resume Next で黙って続ける)。

    - target_path が空なら 使用パス「Access.log などの置き場所」(空ならこの PC の手元)。VBA は ThisWorkbook.path
    - target_file が空なら Access.log
    - ファイルが無ければ作り、1 行目に `管理ﾌｧｲﾙ使用状況`
    """
    file_name = target_file or ACCESS_LOG
    if target_path == "":
        save_text = paths.join(paths.write_dir(), file_name)
    else:
        _create_folder(target_path)
        save_text = paths.join(target_path, file_name)
    line = format_log_line(msg1, sepl, now=now, computer=computer)
    try:
        p = local_path(save_text)
        new = not p.exists()
        with open(p, "ab") as f:
            if new:
                f.write((ACCESS_LOG_TITLE + NEWLINE).encode(ENCODING, ENCODE_ERRORS))
            f.write((line + NEWLINE).encode(ENCODING, ENCODE_ERRORS))
        return True
    except OSError:
        return False


def write_line_log(line: object, msg1: str, *, file: str = "", now: Optional[datetime] = None,
                   computer: Optional[str] = None) -> bool:
    """`Call WriteLog(msg, "", HistoryPath & "<フォルダ>\\<フォルダ>_Log"[, file])` の形(ライン別ログ)。"""
    return write_log(msg1, "", line_log_dir(line), file, now=now, computer=computer)


@dataclass(frozen=True)
class LogEntry:
    """書くべきログ 1 件(純粋な関数が返し、呼び手が `write()` で書く)。

    target_path が "" のときは VBA の `WriteLog(msg, "")`(アプリのフォルダの Access.log)。
    """
    msg: str
    target_path: str = ""
    target_file: str = ""
    sepl: str = ""

    def write(self, *, now: Optional[datetime] = None, computer: Optional[str] = None) -> bool:
        return write_log(self.msg, self.sepl, self.target_path, self.target_file, now=now, computer=computer)


def line_entry(line: object, msg: str, file: str = "") -> LogEntry:
    """ライン別ログに書く LogEntry。"""
    return LogEntry(msg, line_log_dir(line), file)


def log_out_entries(name: str, value: object, line: str, worker_name: str, *, aorb: str = "",
                    aorb_caption: str = "", mitsu: str = "") -> list[LogEntry]:
    """補助.LogOut (補助:1147-1179) が書く 2 件(ライン別 Measured.log と、アプリのフォルダの Access.log)。

    line は `UF1.LineName.Caption & UF1.LineNo.Caption`、aorb_caption は `UF1.AorB.Caption`(右 1 文字を使う)。
    """
    lpath = line_log_dir(line)
    temp = aorb_caption[-1:] if aorb_caption else ""
    v = "" if value is None else str(value)
    if mitsu == "三菱":
        msg = "*UF全数|" + name + "|" + v + "|" + line + "|" + worker_name
    elif aorb == "":
        msg = "*UFKen|" + name + "|" + v + "|" + line + "|" + worker_name
    else:
        msg = "*UFKen|" + "[" + temp + "] " + name + "|" + v + "|" + line + "|" + worker_name
    return [LogEntry(msg, lpath, MEASURED_LOG), LogEntry(msg)]


def log_out(name: str, value: object, line: str, worker_name: str, *, aorb: str = "", aorb_caption: str = "",
            mitsu: str = "", now: Optional[datetime] = None, computer: Optional[str] = None) -> None:
    """補助.LogOut (補助:1147-1179)。測定値 1 件を Measured.log と Access.log に書く。"""
    for e in log_out_entries(name, value, line, worker_name, aorb=aorb, aorb_caption=aorb_caption, mitsu=mitsu):
        e.write(now=now, computer=computer)


def log_out_combo_entries(name: str, value: object, line: str, worker_name: str,
                          aorb_caption: str = "") -> list[LogEntry]:
    """補助.LogOutCombo (補助:1182-1201)。常に `[A]`/`[B]` 付き。"""
    lpath = line_log_dir(line)
    temp = aorb_caption[-1:] if aorb_caption else ""
    v = "" if value is None else str(value)
    msg = "*UFKen|[" + temp + "] " + name + "|" + v + "|" + line + "|" + worker_name
    return [LogEntry(msg, lpath, MEASURED_LOG), LogEntry(msg)]


def log_out_combo(name: str, value: object, line: str, worker_name: str, aorb_caption: str = "", *,
                  now: Optional[datetime] = None, computer: Optional[str] = None) -> None:
    """補助.LogOutCombo (補助:1182-1201)。"""
    for e in log_out_combo_entries(name, value, line, worker_name, aorb_caption):
        e.write(now=now, computer=computer)


def msg_error_entry(trm: str, line: str, t1: str, t2: str, t3: str) -> LogEntry:
    """OutTolerance.MsgError が書くログ(O:787-795)。`*UFKen:MsgError:<文> T1:… T2:… T3:…`。

    MsgError 本体は別担当。Aインプットの高さ NG(Input:570, 579)で必要なので、ログの形だけここに置く。
    """
    return line_entry(line, "*UFKen:" + "MsgError:" + trm + " T1:" + t1 + " T2:" + t2 + " T3:" + t3)
