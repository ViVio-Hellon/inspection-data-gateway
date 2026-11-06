"""測定器の受信機(ミツトヨ U-WAVE-R など)の COM ポートを読む(保守 → 測定機器ID の「COM受信テスト」)。

U-WAVE-R は USB の仮想COMポート(FTDI)。U-WAVEPAK の設定(FTD2XXRS.232)は 57600bps・8bit・パリティなし・
ストップ1・RTS/CTS。U-WAVEPAK は COM を読んでキー入力に変えているだけなので、COM を直接読めば
画面のカーソルに関係なく受け取れる。

- 追加のライブラリは使わない。Windows は ctypes(CreateFile / SetCommState / ReadFile)、ポートの一覧は winreg
- 1 つの COM ポートを開けるのは 1 つのプログラムだけ。U-WAVEPAK が受信中だと開けない(その旨を出す)
- ここでは受け取った文字列を見せるだけ。検査には流さない・何も書かない
- Windows 以外(試験)では、ファイルのパス(擬似端末など)を渡すとそれを読む
"""
from __future__ import annotations

import os
import sys
import threading
import time
from collections import deque
from dataclasses import dataclass, field
from typing import Any, Optional

from ..logging_utils import get_logger

log = get_logger(__name__)

DEFAULT_BAUD = 57600
#: U-WAVEPAK の FTD2XXRS.232 と同じ既定(57600・8bit・パリティなし・ストップ1・RTS/CTS)
DEFAULTS = {"baud": 57600, "bytesize": 8, "parity": "N", "stopbits": 1, "flow": "rtscts"}
PARITY = {"N": 0, "O": 1, "E": 2}            # なし / 奇数 / 偶数
FLOWS = ("none", "rtscts", "xonxoff")
MAX_LINES = 200
IS_WINDOWS = sys.platform == "win32"


D2XX_PREFIX = "FTDI:"     # 受信機を FTD2XX.dll で直接読むときのポート名(FTDI:<シリアル番号>)


def _d2xx():
    """FTD2XX.dll(FTDI のドライバに付いてくる)。無ければ None。"""
    if not IS_WINDOWS:
        return None
    import ctypes
    for name in ("ftd2xx.dll", "ftd2xx64.dll"):
        try:
            return ctypes.WinDLL(name)
        except OSError:
            continue
    return None


def d2xx_devices() -> list[dict[str, Any]]:
    """FTD2XX.dll で見える FTDI の機器(U-WAVE-R など)。仮想COM が切ってあっても見える。"""
    dll = _d2xx()
    if dll is None:
        return []
    import ctypes
    n = ctypes.c_uint32(0)
    if dll.FT_CreateDeviceInfoList(ctypes.byref(n)) != 0:
        return []
    out = []
    for i in range(n.value):
        flags, typ, did, loc = (ctypes.c_uint32() for _ in range(4))
        serial, desc = ctypes.create_string_buffer(16), ctypes.create_string_buffer(64)
        h = ctypes.c_void_p()
        dll.FT_GetDeviceInfoDetail.argtypes = [ctypes.c_uint32] + [ctypes.POINTER(ctypes.c_uint32)] * 4 + \
            [ctypes.c_char_p, ctypes.c_char_p, ctypes.POINTER(ctypes.c_void_p)]
        if dll.FT_GetDeviceInfoDetail(i, ctypes.byref(flags), ctypes.byref(typ), ctypes.byref(did), ctypes.byref(loc),
                                      serial, desc, ctypes.byref(h)) != 0:
            continue
        sn, ds = serial.value.decode("ascii", "replace"), desc.value.decode("ascii", "replace")
        busy = bool(flags.value & 1)                        # FT_FLAGS_OPENED = ほかで使用中
        out.append({"port": D2XX_PREFIX + sn, "device": ds, "usb": True, "d2xx": True,
                    "label": f"{ds or 'FTDI'}(シリアル {sn})直接" + ("・使用中" if busy else "")})
    return out


def list_ports() -> list[dict[str, str]]:
    """この PC の COM ポート(レジストリ HARDWARE\\DEVICEMAP\\SERIALCOMM)。名前の例: \\Device\\VCP0 = USB の仮想COM。"""
    if not IS_WINDOWS:
        return []
    import winreg
    out = []
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"HARDWARE\DEVICEMAP\SERIALCOMM") as k:
            i = 0
            while True:
                try:
                    dev, port, _t = winreg.EnumValue(k, i)
                except OSError:
                    break
                usb = "VCP" in dev.upper() or "USB" in dev.upper()
                out.append({"port": str(port), "device": dev, "usb": usb})
                i += 1
    except OSError as e:
        log.warning("COM ポートの一覧を読めません: %s", e)
    out = sorted(out, key=lambda x: (not x["usb"], int("".join(c for c in x["port"] if c.isdigit()) or 0)))
    try:
        direct = d2xx_devices()                  # 受信機を直接(仮想COM が切ってある U-WAVEPAK の設定でも見える)
    except Exception as e:                       # noqa: BLE001 - 一覧が取れなくても COM は出す
        log.warning("FTD2XX の機器一覧を読めません: %s", e)
        direct = []
    return direct + out


@dataclass
class Line:
    at: str
    text: str
    hex: str


@dataclass
class _State:
    port: str = ""
    baud: int = DEFAULT_BAUD
    flow: str = "rtscts"
    bytesize: int = 8
    parity: str = "N"
    stopbits: int = 1
    running: bool = False
    error: str = ""
    started: str = ""
    bytes: int = 0
    lines: deque = field(default_factory=lambda: deque(maxlen=MAX_LINES))
    seq: int = 0


class ComError(Exception):
    pass


# ------------------------------------------------------------------
# Windows: ctypes
# ------------------------------------------------------------------
class _WinPort:
    def __init__(self, port: str, baud: int, flow: str, bytesize: int = 8, parity: str = "N",
                 stopbits: int = 1) -> None:
        import ctypes
        from ctypes import wintypes as w
        self.ct, self.w = ctypes, w
        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        self.k32 = k32
        k32.CreateFileW.restype = w.HANDLE
        k32.CreateFileW.argtypes = [w.LPCWSTR, w.DWORD, w.DWORD, w.LPVOID, w.DWORD, w.DWORD, w.HANDLE]
        # 64bit の Windows でハンドルが int に切り詰められないよう、使う関数の引数の型を決めておく
        for fn in ("GetCommState", "SetCommState", "SetCommTimeouts"):
            getattr(k32, fn).argtypes = [w.HANDLE, w.LPVOID]
        k32.PurgeComm.argtypes = [w.HANDLE, w.DWORD]
        k32.ReadFile.argtypes = [w.HANDLE, w.LPVOID, w.DWORD, ctypes.POINTER(w.DWORD), w.LPVOID]
        k32.CloseHandle.argtypes = [w.HANDLE]

        D, W8, W16 = ctypes.c_uint32, ctypes.c_uint8, ctypes.c_uint16      # Windows の DWORD / BYTE / WORD

        class DCB(ctypes.Structure):
            _fields_ = [("DCBlength", D), ("BaudRate", D),
                        ("fBinary", D, 1), ("fParity", D, 1), ("fOutxCtsFlow", D, 1),
                        ("fOutxDsrFlow", D, 1), ("fDtrControl", D, 2), ("fDsrSensitivity", D, 1),
                        ("fTXContinueOnXoff", D, 1), ("fOutX", D, 1), ("fInX", D, 1),
                        ("fErrorChar", D, 1), ("fNull", D, 1), ("fRtsControl", D, 2),
                        ("fAbortOnError", D, 1), ("fDummy2", D, 17),
                        ("wReserved", W16), ("XonLim", W16), ("XoffLim", W16),
                        ("ByteSize", W8), ("Parity", W8), ("StopBits", W8),
                        ("XonChar", ctypes.c_char), ("XoffChar", ctypes.c_char), ("ErrorChar", ctypes.c_char),
                        ("EofChar", ctypes.c_char), ("EvtChar", ctypes.c_char), ("wReserved1", W16)]

        class TIMEOUTS(ctypes.Structure):
            _fields_ = [(n, D) for n in ("ReadIntervalTimeout", "ReadTotalTimeoutMultiplier",
                                               "ReadTotalTimeoutConstant", "WriteTotalTimeoutMultiplier",
                                               "WriteTotalTimeoutConstant")]
        name = port if port.startswith("\\\\.\\") else "\\\\.\\" + port
        h = k32.CreateFileW(name, 0xC0000000, 0, None, 3, 0, None)   # GENERIC_READ|WRITE, OPEN_EXISTING
        if h is None or h == w.HANDLE(-1).value:
            err = ctypes.get_last_error()
            why = {2: "そのポートがありません", 5: "ほかのプログラム(U-WAVEPAK など)が使っています"}.get(err, f"Windows のエラー {err}")
            raise ComError(f"{port} を開けません: {why}")
        self.h = h
        dcb = DCB()
        dcb.DCBlength = ctypes.sizeof(DCB)
        if not k32.GetCommState(h, ctypes.byref(dcb)):
            self.close()
            raise ComError(f"{port} の設定を読めません(エラー {ctypes.get_last_error()})")
        dcb.BaudRate, dcb.ByteSize = baud, bytesize
        dcb.Parity = PARITY.get(parity, 0)
        dcb.StopBits = 2 if stopbits == 2 else 0                                 # 0 = ストップ1 / 2 = ストップ2
        dcb.fBinary, dcb.fParity, dcb.fDtrControl = 1, 1 if dcb.Parity else 0, 1
        dcb.fOutxCtsFlow, dcb.fRtsControl = (1, 2) if flow == "rtscts" else (0, 1)   # RTS/CTS(ハンドシェイク)
        dcb.fOutX = dcb.fInX = 1 if flow == "xonxoff" else 0
        dcb.fOutxDsrFlow = dcb.fDsrSensitivity = dcb.fAbortOnError = 0
        if not k32.SetCommState(h, ctypes.byref(dcb)):
            self.close()
            raise ComError(f"{port} を {baud}bps に設定できません(エラー {ctypes.get_last_error()})")
        to = TIMEOUTS(50, 0, 200, 0, 200)          # 200ms ごとに戻って止める合図を見る
        k32.SetCommTimeouts(h, ctypes.byref(to))
        k32.PurgeComm(h, 0x0008 | 0x0002)          # 受信の残りを捨てる

    def read(self) -> bytes:
        buf = self.ct.create_string_buffer(256)
        n = self.w.DWORD(0)
        if not self.k32.ReadFile(self.h, buf, 256, self.ct.byref(n), None):
            raise ComError(f"読めません(エラー {self.ct.get_last_error()})。受信機が抜けたかもしれません")
        return buf.raw[: n.value]

    def close(self) -> None:
        if getattr(self, "h", None):
            self.k32.CloseHandle(self.h)
            self.h = None


class _D2xxPort:
    """受信機を FTD2XX.dll で直接読む(U-WAVEPAK と同じ読み方。仮想COM が切ってあっても読める)。"""

    FLOW = {"none": 0x0000, "rtscts": 0x0100, "xonxoff": 0x0400}

    def __init__(self, serial: str, baud: int, flow: str, bytesize: int = 8, parity: str = "N",
                 stopbits: int = 1) -> None:
        import ctypes
        self.ct = ctypes
        dll = _d2xx()
        if dll is None:
            raise ComError("FTD2XX.dll が見つかりません(FTDI のドライバが入っていない)")
        self.dll = dll
        dll.FT_OpenEx.argtypes = [ctypes.c_char_p, ctypes.c_uint32, ctypes.POINTER(ctypes.c_void_p)]
        for fn in ("FT_SetBaudRate", "FT_Purge"):
            getattr(dll, fn).argtypes = [ctypes.c_void_p, ctypes.c_uint32]
        dll.FT_SetDataCharacteristics.argtypes = [ctypes.c_void_p, ctypes.c_ubyte, ctypes.c_ubyte, ctypes.c_ubyte]
        dll.FT_SetFlowControl.argtypes = [ctypes.c_void_p, ctypes.c_uint16, ctypes.c_ubyte, ctypes.c_ubyte]
        dll.FT_SetTimeouts.argtypes = [ctypes.c_void_p, ctypes.c_uint32, ctypes.c_uint32]
        dll.FT_Read.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint32, ctypes.POINTER(ctypes.c_uint32)]
        for fn in ("FT_Close", "FT_SetDtr", "FT_SetRts"):
            getattr(dll, fn).argtypes = [ctypes.c_void_p]
        h = ctypes.c_void_p()
        st = dll.FT_OpenEx(serial.encode("ascii"), 1, ctypes.byref(h))      # FT_OPEN_BY_SERIAL_NUMBER
        if st != 0:
            why = {2: "その受信機がありません(抜けている)", 3: "開けません。U-WAVEPAK など、ほかのプログラムが使っています"}.get(
                st, f"FTD2XX のエラー {st}")
            raise ComError(f"受信機(シリアル {serial})を開けません: {why}")
        self.h = h
        steps = [dll.FT_SetBaudRate(h, baud),
                 dll.FT_SetDataCharacteristics(h, bytesize, 2 if stopbits == 2 else 0, PARITY.get(parity, 0)),
                 dll.FT_SetFlowControl(h, self.FLOW.get(flow, 0), 0x11, 0x13),
                 dll.FT_SetTimeouts(h, 200, 200), dll.FT_SetDtr(h), dll.FT_SetRts(h), dll.FT_Purge(h, 1 | 2)]
        if any(steps):
            self.close()
            raise ComError(f"受信機の設定ができません(FTD2XX のエラー {[x for x in steps if x][0]})")

    def read(self) -> bytes:
        buf = self.ct.create_string_buffer(256)
        got = self.ct.c_uint32(0)
        st = self.dll.FT_Read(self.h, buf, 256, self.ct.byref(got))       # 200ms で戻る(FT_SetTimeouts)
        if st != 0:
            raise ComError(f"読めません(FTD2XX のエラー {st})。受信機が抜けたかもしれません")
        return buf.raw[: got.value]

    def close(self) -> None:
        if getattr(self, "h", None):
            self.dll.FT_Close(self.h)
            self.h = None


class _FilePort:
    """Windows 以外(試験): ファイル・擬似端末を読む。"""

    def __init__(self, path: str) -> None:
        try:
            self.fd = os.open(path, os.O_RDONLY | getattr(os, "O_NONBLOCK", 0))
        except OSError as e:
            raise ComError(f"{path} を開けません: {e}") from e

    def read(self) -> bytes:
        import select
        r, _w, _x = select.select([self.fd], [], [], 0.2)
        if not r:
            return b""
        try:
            return os.read(self.fd, 256)
        except BlockingIOError:
            return b""

    def close(self) -> None:
        if self.fd is not None:
            os.close(self.fd)
            self.fd = None


# ------------------------------------------------------------------
# 受信(1 つだけ。開始で前のは止める)
# ------------------------------------------------------------------
_lock = threading.Lock()
_state = _State()
_thread: Optional[threading.Thread] = None
_stop = threading.Event()


def _decode(b: bytes) -> str:
    for enc in ("ascii", "cp932"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            pass
    return b.decode("latin-1")


def _run(dev: Any) -> None:
    buf = b""
    try:
        while not _stop.is_set():
            chunk = dev.read()
            if not chunk:
                continue
            with _lock:
                _state.bytes += len(chunk)
            buf += chunk
            # CR・LF で区切る(U-WAVE-R は CR で終わる)。区切りが来ないまま 1 秒たっても出す
            while True:
                cut = min((i for i in (buf.find(b"\r"), buf.find(b"\n")) if i >= 0), default=-1)
                if cut < 0:
                    break
                one, buf = buf[:cut], buf[cut + 1:]
                if one:
                    _push(one)
    except ComError as e:
        with _lock:
            _state.error = str(e)
    except Exception as e:                       # noqa: BLE001 - 受信の止まった理由を画面に出す
        with _lock:
            _state.error = f"{type(e).__name__}: {e}"
    finally:
        if buf:
            _push(buf)
        dev.close()
        with _lock:
            _state.running = False


def _push(raw: bytes) -> None:
    with _lock:
        _state.seq += 1
        _state.lines.append((_state.seq, Line(time.strftime("%H:%M:%S"), _decode(raw).strip(),
                                              " ".join(f"{c:02X}" for c in raw))))
    log.info("COM受信 %s: %r", _state.port, raw)


def start(port: str, baud: int = DEFAULT_BAUD, flow: str = "rtscts", bytesize: int = 8, parity: str = "N",
          stopbits: int = 1) -> dict[str, Any]:
    global _thread
    stop()
    port = str(port or "").strip()
    if not port:
        raise ComError("COM ポートを選んでください")
    if flow not in FLOWS or parity not in PARITY or bytesize not in (7, 8) or stopbits not in (1, 2) or baud <= 0:
        raise ComError("受信の設定が正しくありません")
    if port.startswith(D2XX_PREFIX):
        dev = _D2xxPort(port[len(D2XX_PREFIX):], baud, flow, bytesize, parity, stopbits)
    else:
        dev = _WinPort(port, baud, flow, bytesize, parity, stopbits) if IS_WINDOWS else _FilePort(port)
    with _lock:
        _state.port, _state.baud, _state.flow = port, baud, flow
        _state.bytesize, _state.parity, _state.stopbits = bytesize, parity, stopbits
        _state.running, _state.error, _state.bytes = True, "", 0
        _state.started = time.strftime("%H:%M:%S")
        _state.lines.clear()
    _stop.clear()
    _thread = threading.Thread(target=_run, args=(dev,), name="com-reader", daemon=True)
    _thread.start()
    log.info("COM受信テストを始めました: %s %dbps %d%s%d フロー=%s", port, baud, bytesize, parity, stopbits, flow)
    return status()


def stop() -> dict[str, Any]:
    global _thread
    if _thread is not None:
        _stop.set()
        _thread.join(timeout=2)
        _thread = None
        log.info("COM受信テストを止めました")
    return status()


def status(since: int = 0) -> dict[str, Any]:
    with _lock:
        lines = [{"seq": s, "at": ln.at, "text": ln.text, "hex": ln.hex} for s, ln in _state.lines if s > since]
        return {"port": _state.port, "baud": _state.baud, "flow": _state.flow, "bytesize": _state.bytesize,
                "parity": _state.parity, "stopbits": _state.stopbits, "defaults": DEFAULTS, "running": _state.running,
                "error": _state.error, "started": _state.started, "bytes": _state.bytes, "lines": lines,
                "last": _state.seq, "windows": IS_WINDOWS}
