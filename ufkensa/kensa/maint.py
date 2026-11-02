"""保守画面(管理者だけ)の組み立て ── VBA の Test フォームと 初期設定.bas のうち、業務で要る 2 つ

Test フォームそのものは移植しない(D9)。ただし持ち主の依頼で、次の 2 つは Web 版にも置く:

1. 歪公差チェック(Test.CommandButton34_Click, Test:58-110)
   品質基準の表を変えたときに結果を確かめる。ロット 1 つ、または仕掛の全ロットについて、通常検査の
   ロット登録と同じ部品(`quality.hiki_arr_create` → `quality.dist_array`、検番は D15 のとおりロットの検番
   `quality_session.ken_no_of`)で歪公差を引き、Test と同じ 8 項目を同じ書式で出す。
   別の 品質基準.sqlite3(変更後の表など)を指定すると、いまの表と並べて違いに印を付ける。
   同じ目的の表引きとして、ソリ公差(CommandButton24, Test:433-463)と設備仕様(CommandButton38, Test:186-191)も置く。
2. 端末設定のセット(初期設定.AIM / HVC / DelSet / test取得。初期設定:51-67)
   レジストリ「起動設定」の代わりは user_settings(terminal.py・paths.set_mode。docs/06 レジストリ→user_settings)。

**読み取り専用**: 1 は DB を読むだけ(datasource は mode=ro で開く)。書くのは 2 の端末設定だけ。
"""
from __future__ import annotations

import contextlib
import threading
import time
from pathlib import Path
from typing import Any, Iterator, Optional

from .. import user_settings
from . import datasource, lot as lotmod, paths, quality, quality_session, terminal, vbx
from .flow import Stop

#: Test フォームの表示欄(QpArryD の 1〜8 列。Test:96-105)と、参考に出す 9 列目
DIST_KEYS = ("MG", "HaHiz", "TaHiz", "Naka", "Mimi", "HizCome", "Source", "NoTex", "Flag")
DIST_LABELS = {"MG": "1M当りｹﾞｰｼﾞ", "HaHiz": "巾歪", "TaHiz": "丈歪", "Naka": "中歪", "Mimi": "耳歪",
               "HizCome": "コメント", "Source": "参照元", "NoTex": "管理番号", "Flag": "フラグ"}
#: 照合で起きうる VBA の実行時エラー(1 ロットの失敗にして続ける)
ERRORS = quality_session.ERRORS + (datasource.DbMissing, datasource.TableMissing, Stop)

SORI_TABLE = "ソリ4級"


# ------------------------------------------------------------------
# 品質基準 DB の差し替え(この処理のスレッドだけ)
# ------------------------------------------------------------------
_alt = threading.local()
_alt_lock = threading.Lock()


@contextlib.contextmanager
def quality_db(path: Optional[Path]) -> Iterator[None]:
    """この中で、**このスレッドの** quality の照合だけ品質基準 DB を `path` に向ける(None なら何もしない)。

    quality.py の照合はすべて `quality._quality_db()` から DB を引くので、そこを差し替える。
    ほかのスレッド(通常検査の要求)はこの間もいまの DB のまま(スレッドごとの値で分ける)。
    """
    if path is None:
        yield
        return
    with _alt_lock:
        original = quality._quality_db

        def patched() -> Path:
            p = getattr(_alt, "path", None)
            return p if p is not None else original()

        quality._quality_db = patched
        _alt.path = path
        try:
            yield
        finally:
            _alt.path = None
            quality._quality_db = original


def current_quality_db() -> Path:
    return datasource.resolve("品質基準")


def resolve_alt(path_text: str) -> Optional[Path]:
    """別の品質基準ファイルの指定(空なら None)。無い・.sqlite3 でなければ ValueError(文言は画面に出す)。"""
    text = (path_text or "").strip().strip('"')
    if text == "":
        return None
    p = paths.to_local(text)
    if p.suffix.lower() != ".sqlite3":
        raise ValueError(f"品質基準の .sqlite3 ファイルを指定してください: {text}")
    if not p.is_file():
        raise ValueError(f"ファイルが見つかりません: {text}")
    return p


# ------------------------------------------------------------------
# 歪公差チェック(Test.CommandButton34_Click)
# ------------------------------------------------------------------
def _dist_values(kousa: list[list[Any]]) -> dict[str, str]:
    """Test:96-105。1〜5 列は `Format(x, "0.0")`、6〜8 列はそのまま。9 列目(フラグ)は参考。"""
    k = kousa[0]
    v = {key: vbx.fmt(k[i], "0.0") for key, i in (("MG", quality.K_GAUGE), ("HaHiz", quality.K_HA),
                                                 ("TaHiz", quality.K_TA), ("Naka", quality.K_NAKA),
                                                 ("Mimi", quality.K_MIMI))}
    v["HizCome"] = quality_session._s(k[quality.K_COMMENT])
    v["Source"] = quality_session._s(k[quality.K_SOURCE])
    v["NoTex"] = quality_session._s(k[quality.K_NO])
    v["Flag"] = quality_session._s(k[quality.K_FLAG])
    return v


def _run_dist(lot_no: str, ken_no: str, hiki: Any, data: quality.LotData) -> dict[str, Any]:
    """DistArray(LotNo, KenNo, "Test", "Test")(Test:91)。Execution="Test" なので画面へは書かない。"""
    out: dict[str, Any] = {"status": "ok", "values": {}, "note": "", "messages": []}
    try:
        qp = quality.dist_array(lot_no, ken_no, hiki, data, execution="Test")
    except ERRORS as e:                     # VBA なら実行時エラーで止まる
        out["status"], out["note"] = "error", f"{type(e).__name__}: {getattr(e, 'text', '') or e}"
        return out
    out["messages"] = [m.text for m in qp.messages]
    if not qp.kousa:
        # VBA は QpArryD が Empty のまま LBound を取って実行時エラー(13)
        out["status"], out["note"] = "none", "該当する歪公差がありません(VBA の Test は実行時エラー)"
        return out
    out["values"] = _dist_values(qp.kousa)
    return out


def _lot_context(lot_no: str) -> dict[str, Any]:
    """ロットの値(SIKALOTNOW)・検番(D15)・引当(HikiArrCreate)。品質基準の差し替えに関係しない部分。"""
    f = lotmod.arrbox("SIKALOTNOW", "板", "Fields")
    r = lotmod.arrbox("SIKALOTNOW", lot_no, "Record")
    data = quality.LotData(f.value, r.value)
    ken = quality_session.ken_no_of(lot_no, data)
    hiki = quality.hiki_arr_create(lot_no)
    return {"data": data, "ken_no": ken, "hiki": hiki, "found": bool(r.value),
            "messages": [m.text for m in f.messages + r.messages]}


def diff_keys(a: dict[str, Any], b: Optional[dict[str, Any]]) -> list[str]:
    """いまの表と別の表で違う所(状態・表示欄)。"""
    if b is None:
        return []
    out = ["status"] if a["status"] != b["status"] else []
    return out + [k for k in DIST_KEYS if a["values"].get(k, "") != b["values"].get(k, "")]


def dist_check(lot_no: str, alt: Optional[Path] = None) -> dict[str, Any]:
    """ロット 1 つの歪公差。alt を渡すと、その品質基準で引いた結果も並べる。"""
    lot_no = lot_no.strip().upper()
    ctx = _lot_context(lot_no)
    row: dict[str, Any] = {"lot": lot_no, "ken_no": ctx["ken_no"], "found": ctx["found"],
                           "hiki": [list(h) for h in (ctx["hiki"] or [])], "messages": ctx["messages"]}
    if not ctx["found"] and ctx["hiki"] is None:
        row["current"] = {"status": "not_found", "values": {}, "note": "仕掛(SIKALOTNOW・SIKAHIKINOW)にロットがありません",
                          "messages": []}
        row["alt"], row["diff"] = None, []
        return row
    row["current"] = _run_dist(lot_no, ctx["ken_no"], ctx["hiki"], ctx["data"])
    row["alt"] = None
    if alt is not None:
        with quality_db(alt):
            row["alt"] = _run_dist(lot_no, ctx["ken_no"], ctx["hiki"], ctx["data"])
    row["diff"] = diff_keys(row["current"], row["alt"])
    return row


def lot_numbers(*, plate_only: bool = True, prefix: str = "") -> list[str]:
    """仕掛(SIKALOTNOW)のロットNo(取り込み順・重複なし)。plate_only は Arrbox の "板"([ｵｰﾀﾞｰ板丈]<>0)。"""
    f = lotmod.arrbox("SIKALOTNOW", "板", "Fields")
    if f.value is None:
        return []
    rows = datasource.get_records(f.file_path, lotmod.TABLE, lotmod.arrbox_filter("板") if plate_only else "") or []
    c = quality.col_of(f.value, "ﾛｯﾄ番号")
    prefix = prefix.strip().upper()
    lots = (quality_session._s(quality.cell(r, c)) for r in rows)
    return list(dict.fromkeys(x for x in lots if x and x.upper().startswith(prefix)))


def dist_check_all(alt: Optional[Path] = None, *, plate_only: bool = True, prefix: str = "") -> dict[str, Any]:
    """仕掛の全ロット(またはロットNo の頭で絞ったもの)の歪公差。表を変えたあとの確認用。"""
    t0 = time.perf_counter()
    rows = [dist_check(lot, alt) for lot in lot_numbers(plate_only=plate_only, prefix=prefix)]
    count = {s: sum(1 for r in rows if r["current"]["status"] == s) for s in ("ok", "none", "error", "not_found")}
    return {"rows": rows, "count": len(rows), "by_status": count,
            "diffs": sum(1 for r in rows if r["diff"]), "alt": str(alt) if alt else "",
            "sec": round(time.perf_counter() - t0, 1)}


# ------------------------------------------------------------------
# ソリ公差(Test.CommandButton24_Click)・設備仕様(Test.CommandButton38_Click)
# ------------------------------------------------------------------
def _single_of(text: str, label: str) -> float:
    """`Dim x As Single: x = Me.TextBox`。空・数にならない文字列は VBA では実行時エラー 13。"""
    try:
        return quality.csng(str(text).strip())
    except vbx.TypeMismatch:
        raise ValueError(f"{label} に数を入れてください(VBA では実行時エラー 13)") from None


def sori_check(atu: str, haba: str, take: str, db: Optional[Path] = None) -> dict[str, Any]:
    """Test.CommandButton24_Click (Test:433-463)。ソリ4級 から 幅反り公差・丈反り公差。

    - 表が空(Fields / Records が Empty)なら VBA は何も出さずに抜ける → status "none"
    - 列は FindIndices("H"/"T", 幅/丈)、行は 板厚下 <= 厚 <= 板厚上 の最初の行(SoriArray)
    - 表示は `"幅" & Format(幅, "0.0") & " " & "丈" & Format(丈, "0.0")`
    VBA は担当者のデスクトップの 品質基準.accdb を直書きで読んでいた。ここでは使用パスの品質基準(または db)。
    """
    a, h, t = _single_of(atu, "厚"), _single_of(haba, "幅"), _single_of(take, "丈")
    src = db or current_quality_db()
    fields = datasource.get_fields(src, SORI_TABLE)
    rows = datasource.get_records(src, SORI_TABLE, "") if fields is not None else None
    if fields is None or rows is None:
        return {"status": "none", "text": "", "note": f"{SORI_TABLE} が空です(VBA は何も出さずに終わる)"}
    try:
        hit_h = quality.find_indices(fields, "H", h)
        hit_t = quality.find_indices(fields, "T", t)
        w = quality.sori_array(rows, fields, a, hit_h[0])
        l_ = quality.sori_array(rows, fields, a, hit_t[0])
    except ERRORS as e:
        return {"status": "error", "text": "", "note": f"{type(e).__name__}: {e}"}

    def col_name(col: int) -> str:
        return fields[col - 1] if 1 <= col <= len(fields) else ""
    ws, ls = vbx.fmt(w, "0.0"), vbx.fmt(l_, "0.0")
    return {"status": "ok", "haba": ws, "take": ls, "text": "幅" + ws + " " + "丈" + ls,
            "haba_col": col_name(hit_h[0]), "take_col": col_name(hit_t[0]), "note": ""}


def equipment_check(line: str, zai: str, atu: str, haba: str, take: str, wei: str) -> dict[str, Any]:
    """Test.CommandButton38_Click (Test:186-191)。CheckEquipmentDb の結果("可" / "… 範囲外" / 作業可否)。

    VBA は引数を直書き(LVC, 5NB, 0.49, 730, 1500, 225)していた。ここでは入れた値で引く(伝送用ファイル の 設備仕様)。
    """
    nums = [_single_of(v, n) for v, n in ((atu, "厚"), (haba, "幅"), (take, "丈"), (wei, "単重"))]
    try:
        text = quality.check_equipment_db(line, zai, *nums)
    except ERRORS as e:
        return {"status": "error", "text": "", "note": f"{type(e).__name__}: {e}"}
    return {"status": "ok", "text": text, "note": ""}


# ------------------------------------------------------------------
# 端末設定のセット(初期設定.AIM / HVC / DelSet / test取得)
# ------------------------------------------------------------------
#: 初期設定.AIM(初期設定:51-54)・HVC(初期設定:55-57)。モード None は「書かない」
PRESETS = {
    "AIM": {"facility": "Test", "mode": paths.ONLINE,
            "label": "AIM(設備名 Test・オンライン)", "vba": "初期設定.AIM"},
    "HVC": {"facility": "HVC", "mode": None, "label": "HVC(設備名 HVC。モードはそのまま)", "vba": "初期設定.HVC"},
}
MISSING = "存在なし"     # GetSetting の既定値(初期設定:59, 70-71)


def terminal_state() -> dict[str, Any]:
    """初期設定.test取得 / Test.CommandButton22: いまの 設備名・モード(と号機)。"""
    fac = terminal.facility()
    saved_mode = user_settings.get(paths.KEY_MODE, None)
    return {
        "facility": fac, "facility_text": fac or MISSING, "line_no": terminal.line_no(),
        "label": terminal.line_label(), "unset": terminal.line() is None,
        "mode": paths.mode(), "mode_saved": saved_mode is not None,
        "mode_text": paths.MODE_LABELS[paths.mode()] + ("" if saved_mode is not None else f"({MISSING}。既定でオンライン)"),
        "presets": [{"key": k, "label": v["label"], "vba": v["vba"]} for k, v in PRESETS.items()],
    }


def apply_preset(key: str) -> dict[str, Any]:
    """初期設定.AIM / HVC。号機は terminal.set_facility のとおり(丸徳以外は空にする)。"""
    p = PRESETS.get(key)
    if p is None:
        raise ValueError(f"知らないセットです: {key}")
    terminal.set_facility(p["facility"])
    if p["mode"] is not None:
        paths.set_mode(p["mode"])
    return terminal_state()


def clear_terminal() -> dict[str, Any]:
    """初期設定.DelSet (初期設定:61-67)。設備名・モードを消す(号機も)。

    次に検査画面を開くと、VBA の 初期設定値確認 → SF3 と同じく「この端末」の設定を求める
    (設備名が無いと UF1 は「ライン設定がありません」)。モードは消すと既定(オンライン)で読む。
    """
    user_settings.delete(terminal.KEY_FACILITY)
    user_settings.delete(terminal.KEY_LINE_NO)
    user_settings.delete(paths.KEY_MODE)
    return terminal_state()
