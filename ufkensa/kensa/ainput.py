"""Aインプット(ホスト送信)と梱包高さチェック

VBA `【UFKen_09】Input` と `UFAinput` の移植。

- 梱包高さチェック `height_check`(Input:430-584 の**現行コード**)。コメントアウトされた旧版(Input:592-751)とは
  合紙・外注・外装紙・VC の計算が違う(不明点 W1)。旧版は移植しない。
- 電文(hostsend.txt): `"KEN:AIN:" & "0185" & A3 … A40 & ライン` の固定長 1 行(FormArray n=2)。
- 梱包課保存用 CSV: 項目名と値の 2 列、各行の末尾にカンマ、cp932、`<検査履歴>\\<ライン>\\<ライン>_Ainput\\yyyy.mm\\<Lot>\\
  yyyymmdd_<検査No>.csv`。同名があれば .02〜.99、.99 もあれば止める(VBA は End)。
- 送信アプリ `SokuteiDataTrans.exe` は Windows のときだけ起動する(ほかの OS では起動しない)。
- 画面(UFAinput)のコントロールの値は `AinputForm` で受け取る。押すたびに呼ぶ `ain_button`(2 回押しの流れ)。
"""
from __future__ import annotations

import os
import subprocess
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional

from . import datasource, legacy_log, lines, paths, vbx
from .flow import ERROR, INFO, WARN, Message, Question, Stop
from .workers import AinputValues, DailyContext, WorkerRegistry, daily_report, workers_text

MAX_MOUNT = 15                       # Input:7
NIS1, NIS2 = "KEN:AIN:", "0185"      # Input:1233-1234
HOSTSEND = "hostsend"                # ExpCSV(…, "hostsend", NisPath, vbTab) → hostsend.txt
RESULT_FILE = "result.txt"
SENDER_DIR, SENDER_EXE = "SokuteiDataTrans", "SokuteiDataTrans.exe"
SUCCESS_TEXT = "正常に送信が行われました。"
VC_DEFAULT = 0.08406977              # 登録 VC 全平均厚(2022.12.19)Input:796
HIGH_PRECISION = ("H191", "H242", "H241")   # Input:512
CHIMES, WAIT_SOUND, DONE_SOUND = "chimes.wav", "お待ちください.wav", "完了.wav"
CSV_ENCODING = "cp932"


def _single(x: float) -> float:
    return vbx.single(x)


def _veq(a: Any, b: Any) -> bool:
    return vbx.vcompare(a, b) == 0


def _cell(v: Any) -> Any:
    """Excel のセル値("" は Empty = None)。"""
    return None if v == "" else v


def _hist(row: list[Any], col: int) -> Any:
    return _cell(row[col - 1]) if col - 1 < len(row) else None


def _nz(v: Any) -> str:
    """Input.Nz (Input:587-589)。Null → ""、ほかは CStr。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return vbx.cstr(v)


# ------------------------------------------------------------------
# 画面の値
# ------------------------------------------------------------------
@dataclass
class AinputForm:
    """UFAinput のコントロール(Caption / Text / Visible / Tag)。"""
    line_name: str = ""          # UFAinput.LineName(UF1.LineName & LineNo)
    uf1_line: str = ""           # UF1 で選んでいるライン(L1Work/LVC/HVC/Bala/Tott/Maru。号機なし)。CSV の置き場所
    lot: str = ""
    ken_no: str = ""
    order_no: str = ""
    zai: str = ""
    tyo: str = ""
    atu: str = ""
    haba: str = ""
    take: str = ""
    ken_law: str = ""
    youto: str = ""
    amin: str = ""
    hmin: str = ""
    tmin: str = ""
    amax: str = ""
    hmax: str = ""
    tmax: str = ""
    hiki_no: str = ""
    konno: str = ""
    remain: str = ""
    mai_flag: str = "1"          # 1 = 枚数1×包数1、2 = 2 段
    mai1: str = ""
    cou1: str = ""
    mai2: str = ""
    cou2: str = ""
    material1: str = ""
    material2: str = ""
    material3: str = ""
    gw: str = ""
    pile_match_no: str = ""
    pile_match: str = ""
    ex_visible: bool = False     # ExFrame(輸出)
    ex_case_no: str = ""
    ex_material1: str = ""
    ex_material2: str = ""
    ex_material3: str = ""
    ex_gw: str = ""
    outside_atu: str = ""
    outside_haba: str = ""
    outside_take: str = ""
    mount_count: str = ""        # Mount_Count.Caption(山数)
    heights: list[str] = field(default_factory=lambda: [""] * MAX_MOUNT)   # DynHL_1〜15 の Caption
    l1_visible: bool = False
    l1_caption: str = ""
    l2_visible: bool = False
    gaityuu_text: str = ""       # 「外注」と入れると合紙・VC・外装紙を外せる
    ai: bool = False
    ga: bool = False
    vab: bool = False
    locked: bool = True          # Ai / Ga / Vab の Locked
    ai_text: str = ""
    ga_text: str = ""
    vc_text: str = ""
    vcab: str = ""               # VCAB.Caption
    db_flag: str = ""            # "true" = Lot情報履歴 に無く SIKACOMP から引いた
    tag: str = ""                # "Load" = 処理中
    ng_flag: str = ""            # "stop" = 送信失敗
    start_tb_tag: str = ""


# ------------------------------------------------------------------
# 山高さ
# ------------------------------------------------------------------
def get_k_height(form: AinputForm, n: int) -> str:
    """Input.GetKHeight (Input:1604-1609)。無い番号は ""。"""
    return form.heights[n - 1] if 1 <= n <= len(form.heights) else ""


def calc_mount_height(form: AinputForm, mount_count: int) -> float:
    """Input.CalcMountHeight (Input:1633-1640)。Σ Val(山高さ)。合計は Single。"""
    total = 0.0
    for i in range(1, mount_count + 1):
        total = _single(total + vbx.val(get_k_height(form, i)))
    return total


def all_height_filled(form: AinputForm, mount_count: int) -> bool:
    """Input.AllHeightFilled (Input:1644-1653)。山数 0 なら True。"""
    return all(get_k_height(form, i) != "" for i in range(1, mount_count + 1))


def next_empty_height_index(form: AinputForm, mount_count: int) -> int:
    """Input.NextEmptyHeightIndex (Input:1656-1665)。全部埋まっていれば -1。"""
    for i in range(1, mount_count + 1):
        if get_k_height(form, i) == "":
            return i
    return -1


def clear_all_k_height(form: AinputForm) -> None:
    """Input.ClearAllKHeight (Input:1668-1673)。"""
    form.heights = [""] * MAX_MOUNT


def _mount_count(form: AinputForm) -> int:
    return vbx.clng(vbx.val(form.mount_count))


def l_height(form: AinputForm) -> float:
    """1 山あたりの高さ `(厚 × 枚数1 × 包数1) / 山数`(Single)。山数 0 は VBA の 0 除算エラー → ZeroDivisionError。"""
    num = vbx.val(form.atu) * vbx.val(form.mai1) * vbx.val(form.cou1)
    return _single(num / vbx.val(form.mount_count))


def l1_caption(lh: float) -> str:
    """≤300 はデプス(Height ゲージ)、>300 はコンベックス(UFAinput:1484-1488)。"""
    if lh <= 300.0:
        return "Depth測定"
    if lh > 300.0:
        return "ｺﾝﾍﾞｯｸｽ測定"
    return "伝送高さ測定"


def show_height_controls(form: AinputForm) -> None:
    """UFAinput.ShowHeightControls (UFAinput:1473-1497)。L1/L2 を出す(山数より後のラベルは隠すだけで消さない)。"""
    form.l1_caption = l1_caption(l_height(form))
    form.l1_visible = True
    form.l2_visible = True


def input_mount_count(form: AinputForm, score: Optional[str]) -> list[Message]:
    """Input.Number_Inp の山数(Input:402-428)。score=None はキャンセル。1〜15 以外は案内して変えない。"""
    msgs: list[Message] = []
    if score is None:
        return msgs
    mc = vbx.clng(vbx.val(score))
    if mc < 1 or mc > MAX_MOUNT:
        msgs.append(Message("山数は1～" + str(MAX_MOUNT) + "で入力してください。", WARN, source="Input:415"))
        return msgs
    form.mount_count = score.lower()
    mc2 = vbx.clng(vbx.val(form.mount_count))
    if 1 <= mc <= MAX_MOUNT:
        form.mount_count = form.mount_count
        _ = mc2
        show_height_controls(form)
    return msgs


# ------------------------------------------------------------------
# 伝送文字列(補助.Maker / Conv の写し。共通化の候補)
# ------------------------------------------------------------------
def _narrow(s: str) -> str:
    out = []
    for ch in s:
        o = ord(ch)
        out.append(chr(o - 0xFEE0) if 0xFF01 <= o <= 0xFF5E else (" " if ch == "\u3000" else ch))
    return "".join(out)


def maker(trm: str) -> str:
    """補助.Maker (補助:393-406)。3 文字以上で先頭 2 文字が DT / #L / #〆。"""
    if trm == "" or len(trm) < 3:
        return ""
    head = _narrow(trm[:2].upper())
    return head if head in ("DT", "#L", "#〆") else ""


def conv(trm: str) -> str:
    """補助.Conv (補助:380-392)。DT → 9 文字目から 11 文字、#L / #〆 → 3 文字目から、ほかはそのまま。"""
    k = maker(trm)
    if k == "DT":
        return vbx.mid(trm, 9, 11)
    if k in ("#L", "#〆"):
        return vbx.mid(trm, 3, len(trm))
    return trm


def depth_set_atu(form: AinputForm) -> float:
    """Height_Machin が DbReturnMicro に渡す厚(厚が空なら 1)。UFAinput:618-620。"""
    return _single(vbx.val(form.atu)) if form.atu != "" else 1.0


def height_machin(trm: str, form: AinputForm, depth_device: str) -> bool:
    """UFAinput.Height_Machin (UFAinput:610-634)。デプスの機器番号が DB の機器と同じか。

    depth_device は `DbReturnMicro(LineName, depth_set_atu(form), " ", "", "デプス")`(別担当)の結果。
    丸徳(上の 6 ライン以外)は機器番号が無い(DT でない)ことが条件。
    """
    machin_no = vbx.mid(trm, 3, 5) if maker(trm) == "DT" else ""
    ln = lines.find(form.line_name)
    if ln is not None and ln.official != "中板":
        return machin_no == depth_device
    return machin_no == ""


def accept_height(form: AinputForm, text: str, depth_device: str) -> list[Message]:
    """UFAinput.StartTB_KeyDown (Tab)(UFAinput:697-746)。伝送値を次の空いた山に入れる。

    ≤300 はデプス(機器番号が DB と一致)、>300 はコンベックス(#L)。合わなければ何も入れない。
    """
    msgs: list[Message] = []
    if form.start_tb_tag != "" or form.mount_count == "":
        return msgs
    lh = l_height(form)
    mc = _mount_count(form)
    idx = next_empty_height_index(form, mc)
    if idx == -1:
        return msgs
    val_str = ""
    if lh <= 300.0:
        if not height_machin(text, form, depth_device):
            return msgs
        val_str = vbx.cstr(vbx.val(conv(text)))
        msgs.append(Message("", INFO, sound=CHIMES, source="UFAinput:729"))
    elif lh > 300.0:
        if maker(text) == "#L":
            val_str = vbx.cstr(vbx.val(conv(text)))
            msgs.append(Message("", INFO, sound=CHIMES, source="UFAinput:733"))
        else:
            return msgs
    if 1 <= idx <= MAX_MOUNT:
        form.heights[idx - 1] = val_str
    return msgs


# ------------------------------------------------------------------
# VC 膜厚
# ------------------------------------------------------------------
def vc_conv(code: str, *, test_line: bool = False) -> tuple[float, list[Message]]:
    """Input.VcConv (Input:753-803)。伝送用ファイル「VC厚」の 厚-係数(Single)。

    登録が無ければ「VC厚登録なし,担当者に連絡のこと」で 0.08406977。ファイルが無ければ案内して 0。
    """
    msgs: list[Message] = []
    db = paths.db_file("伝送用ファイル", base_key="packing" if test_line else "quality")
    if not datasource.exists(db):
        msgs.append(Message("日稲:" + "伝送用ファイル.accdb" + "がありません" + "_ReadCheck_VcConv", ERROR,
                            source="Input:771"))
        return 0.0, msgs
    temp = datasource.get_records(db, "VC厚", f"[品名ｺｰﾄﾞ]='{code}'")
    fields = datasource.get_fields(db, "VC厚")
    if fields is None:
        msgs.append(Message("fieldsなし" + "_VcConv", INFO, source="Input:784"))
        return 0.0, msgs
    c = datasource.column_index(fields, "厚-係数")
    if temp is None:
        msgs.append(Message("VC厚登録なし,担当者に連絡のこと", INFO, source="Input:795"))
        return _single(VC_DEFAULT), msgs
    if c < 0:
        raise IndexError("VC厚 に 厚-係数 列がありません(VBA の実行時エラー 9)")
    return _single(vbx.cdbl(temp[0][c])), msgs


# ------------------------------------------------------------------
# 梱包高さチェック
# ------------------------------------------------------------------
@dataclass
class HeightCheckResult:
    ng: bool                          # True = NG(VBA の戻り値)
    kind: str = ""                    # "over" 許容値ｵｰﾊﾞｰ / "short" 枚数不足 / "" 合格・判定なし
    messages: list[Message] = field(default_factory=list)
    msg_error: str = ""               # MsgError に出す文(NG のとき)
    show_uf1: bool = False            # オーダーNo が無いとき UF1.Show
    ais: float = 0.0
    vcs: float = 0.0
    gai: float = 0.0
    gp: float = 0.0
    gm: float = 0.0
    error_a: float = 0.0
    error_b: float = 0.0
    sap: float = 0.0
    sam: float = 0.0


def _fmt_single(x: float, pattern: str) -> str:
    """`Format(<Single>, …)`。Single は有効 7 桁の 10 進として丸める(V4 と同じ扱い。U5 の実機確認待ち)。"""
    return vbx.fmt(vbx.cstr(_single(x), is_single=True), pattern)


def height_check(atu1: float, atu2: float, hei: float, mai: int, order_no: str, form: AinputForm,
                 history_rows: Optional[list[list[Any]]], *,
                 vc: Callable[[str], tuple[float, list[Message]]] = vc_conv) -> HeightCheckResult:
    """Input.Height_Check (Input:430-584)。梱包の山高さの実測と「板厚×枚数＋合紙＋VC＋外装紙」を比べる。

    - atu1=最小厚、atu2=最大厚、hei=山高さの合計、mai=枚数(引数は VBA と同じく Single / Long に直す)
    - history_rows は Lot情報履歴 の 2 行目から(None はシートなし → 判定せず False)
    - オーダーNo が無ければ「Lot情報履歴内にｵｰﾀﾞｰNO無し」で **False(合格扱い)**
    - NG: 実測 > 最大の理論値 かつ 差(小数 2 桁に四捨五入)> 3/4×最大厚 → [許容値ｵｰﾊﾞｰ]
          実測 < 最小の理論値 かつ 差 > 1/3×最小厚 → [枚数不足(?)]。差がちょうどなら合格。
      NG のときは山高さを消し、Tag を "" にする(form を書き換える)。
    - 画面の AiText / VcText / GaText / VCAB も form に書く。
    """
    res = HeightCheckResult(False)
    if history_rows is None:
        return res
    atu1, atu2, hei = _single(atu1), _single(atu2), _single(hei)
    mai = vbx.clng(mai)
    x = None
    for row in history_rows:
        if _veq(_hist(row, 2), order_no):
            x = row
            break
    if x is None:
        res.messages.append(Message("Lot情報履歴内にｵｰﾀﾞｰNO無し", INFO, source="Input:454"))
        res.show_uf1 = True
        return res

    ais = vcs = gai = a1 = v1 = 0.0
    atu_max, atu_min = atu2, atu1
    gaityuu = form.gaityuu_text == "外注"
    # 合紙
    unit_ais = 0.0
    if vbx.vcompare(_hist(x, 17), 0) != 0:
        unit_ais = _single(0.029)
        form.ai_text = "通常"
    elif _veq(form.remain, 1):
        unit_ais = _single(0.042)
        form.ai_text = "ストア"
    if unit_ais > 0 or (gaityuu and form.ai):
        if gaityuu and form.ai:
            unit_ais = _single(0.029)
            form.ai_text = "通常"
        a1 = unit_ais
        extra = 1
        if form.cou1 != "":
            extra = vbx.clng(extra + vbx.val(form.cou1))
        if form.cou2 != "":
            extra = vbx.clng(extra + vbx.val(form.cou2))
        ais = _single(vbx.val(mai + extra - 1) * a1)
        form.ai_text = form.ai_text + vbx.cstr(ais, is_single=True)
        if gaityuu and form.ai:
            pass
        elif gaityuu:
            a1 = 0.0
            ais = 0.0
            form.ai_text = "0"
    # VC(VBA の Or は両辺を評価する)
    vca = vcb = 0.0
    if _veq(_hist(x, 18), 1):
        not_store = vbx.vcompare(form.remain, 1) != 0
        high = form.youto in HIGH_PRECISION
        if not_store or high:
            if _nz(_hist(x, 19)) != "":
                vca, m = vc(_nz(_hist(x, 19)))
                res.messages += m
            if _nz(_hist(x, 20)) != "":
                vcb, m = vc(_nz(_hist(x, 20)))
                res.messages += m
            v1 = _single(vca + vcb)
            vcs = _single(v1 * vbx.val(mai))
            form.vc_text = vbx.cstr(vcs, is_single=True)
            form.vcab = form.vc_text
            if gaityuu and form.vab:
                v1 = 0.0
                vcs = 0.0
                form.vc_text = "0"
                form.vcab = "0"
    else:
        form.vcab = "0"
    # 外装紙(個装 2 包以上のとき 1 包あたり 0.119×3)
    if vbx.val(form.cou1) > 1 or vbx.val(form.cou2) > 1:
        cnt = 0
        if vbx.val(form.cou1) > 1:
            cnt = vbx.clng(cnt + vbx.val(form.cou1))
        if vbx.val(form.cou2) > 1:
            cnt = vbx.clng(cnt + vbx.val(form.cou2))
        gai = _single(cnt * (2 * 0.119 + 0.119))
    form.ga_text = vbx.cstr(gai, is_single=True)
    if gaityuu and form.ga:
        gai = 0.0
        form.ga_text = "0"
    # 判定
    gosa = _single(_single(ais + vcs) + gai)
    gp = _single(atu_max * mai + gosa)
    gm = _single(atu_min * mai + gosa)
    error_a = _single(vbx.val_single(atu_max) * 3 / 4)
    error_b = _single(vbx.val_single(atu_min) * 1 / 3)
    sap = _single(abs(vbx.val_single(hei) - vbx.val_single(gp)))
    sam = _single(abs(vbx.val_single(hei) - vbx.val_single(gm)))
    res.ais, res.vcs, res.gai, res.gp, res.gm = ais, vcs, gai, gp, gm
    res.error_a, res.error_b, res.sap, res.sam = error_a, error_b, sap, sam
    hv = vbx.val_single(hei)
    if hv > gp and vbx.vcompare(_fmt_single(sap, "0.00"), error_a) > 0:
        count_1 = _single(_single(atu_max + a1) + v1)
        res.ng, res.kind = True, "over"
        res.msg_error = ("[許容値ｵｰﾊﾞｰ]" + " *測定高さ:" + _fmt_single(hei, "0.00") + "*推定枚数:"
                         + _fmt_single(_single(_single(hei - gai) / count_1), "0.00")
                         + " *入力枚数:" + str(mai) + "枚")
    elif hv < gm and vbx.vcompare(_fmt_single(sam, "0.00"), error_b) > 0:
        count_1 = _single(_single(atu_min + a1) + v1)
        res.ng, res.kind = True, "short"
        res.msg_error = ("[枚数不足(?)]" + " *測定高さ:" + _fmt_single(hei, "0.00") + "*推定枚数:"
                         + _fmt_single(_single(_single(hei - gai) / count_1), "0.00")
                         + " *入力枚数:" + str(mai) + "枚")
    if res.ng:
        res.messages.append(Message(res.msg_error, ERROR, title="警告", source="Input:570" if res.kind == "over"
                                    else "Input:579"))
        clear_all_k_height(form)
        form.tag = ""
    return res


# ------------------------------------------------------------------
# 電文と CSV
# ------------------------------------------------------------------
def zer_con(date_no: int, buf: str) -> str:
    """Input.ZerCon (Input:1335-1371)。項目番号ごとの固定長。空欄の既定値と、値の右寄せ 0 埋め(あふれは左を切る)。"""
    if buf == "":
        if date_no == 3:
            return "00000000"
        if date_no in (4, 41):
            return "00"
        if date_no == 5:
            return "0"
        if 6 <= date_no <= 11:
            return "0000000"
        if date_no in (12, 14) or 16 <= date_no <= 21 or date_no == 24 or date_no in (32, 37, 38, 39):
            return "00000"
        if date_no == 25:
            return "     "
        if date_no in (13, 15, 33, 34, 35):
            return "000"
        if date_no in (22, 23, 26, 27, 28, 29, 36):
            return "0000"
        if date_no == 30:
            return " "
        if date_no == 31:
            return "          "
        if date_no == 40:
            return "000000"
        return ""
    if date_no == 3:
        return buf
    if date_no in (4, 41):
        return vbx.right("0" + buf, 2)
    if date_no == 5:
        return buf
    if date_no in (12, 14, 16, 17, 18, 19, 20, 21, 24, 25, 32) or 37 <= date_no <= 39:
        return vbx.right("0000" + buf, 5)
    if date_no in (13, 15, 33, 34, 35):
        return vbx.right("00" + buf, 3)
    if date_no in (22, 23, 26, 27, 28, 29, 36):
        return vbx.right("000" + buf, 4)
    if date_no == 30:
        return buf
    if date_no == 31:
        return vbx.left(buf + "         ", 10)
    if date_no == 40:
        return vbx.right("00000" + buf, 6)
    return ""


def st_con(temp: Any, pits: str) -> str:
    """Input.StCon (Input:1302-1318)。厚×1000・巾×100・丈×10 を CStr(15 桁)して右 7 桁(0 埋め)。0 は ""。

    引数は Double(画面の Caption を CDbl。数にならなければ型エラー 13)。
    """
    t = vbx.cdbl(temp)
    if t == 0:
        return ""
    if pits == "厚":
        cal, ket = t * 1000, "0000"
    elif pits == "巾":
        cal, ket = t * 100, "00000"
    elif pits == "丈":
        cal, ket = t * 10, "000000"
    else:
        cal, ket = 0.0, ""
    return vbx.right(ket + vbx.cstr(cal), 7)


def lin_conv(line: str) -> str:
    """Input.LinConv (Input:1320-1333)。ライン名 → NIS コード。それ以外は "Maru"(4 文字)。"""
    if line == "":
        return ""
    return {"L-1": "05", "LVC": "03", "HVC": "04", "ﾊﾞﾗﾝｻｰ": "01", "ﾄｯﾄ": "01", "LS4": "06", "NS1": "07",
            "Test": "AA"}.get(line, "Maru")


def _converted(form: AinputForm) -> dict[str, str]:
    f = form
    a: dict[str, str] = {
        "A3": zer_con(3, f.hiki_no), "A4": zer_con(4, f.konno), "A5": zer_con(5, f.remain),
        "A12": zer_con(12, f.mai1), "A14": zer_con(14, f.mai2), "A13": zer_con(13, f.cou1),
        "A15": zer_con(15, f.cou2), "A17": zer_con(17, f.material1), "A19": zer_con(19, f.material2),
        "A21": zer_con(21, f.material3), "A16": zer_con(16, f.gw), "A18": zer_con(18, ""),
        "A20": zer_con(20, ""), "A22": zer_con(22, ""), "A23": zer_con(23, ""), "A24": zer_con(24, ""),
        "A32": zer_con(32, ""), "A33": zer_con(33, ""), "A34": zer_con(34, ""), "A35": zer_con(35, ""),
        "A36": zer_con(36, ""), "A40": zer_con(40, ""), "A30": zer_con(30, f.pile_match_no),
        "A31": zer_con(31, f.pile_match), "A25": zer_con(25, f.ex_case_no), "A26": zer_con(26, f.ex_material1),
        "A27": zer_con(27, f.ex_material2), "A28": zer_con(28, f.ex_material3), "A29": zer_con(29, f.ex_gw),
        "A37": zer_con(37, f.outside_atu), "A38": zer_con(38, f.outside_haba), "A39": zer_con(39, f.outside_take),
    }
    a["A6"] = st_con(f.amin, "厚") if f.amin != "" else ""
    a["A9"] = st_con(f.amax, "厚") if f.amax != "" else ""
    a["A7"] = st_con(f.hmin, "巾") if f.hmin != "" else ""
    a["A10"] = st_con(f.hmax, "巾") if f.hmax != "" else ""
    a["A8"] = st_con(f.tmin, "丈") if f.tmin != "" else ""
    a["A11"] = st_con(f.tmax, "丈") if f.tmax != "" else ""
    a["line"] = lin_conv(f.line_name)
    return a


CSV_LABELS = ("1_固定値", "2_固定値", "3_引当No", "4_梱包No", "5_出荷ｽﾄｱ", "6_厚Min", "7_巾Min", "8_丈Min", "9_厚Max",
              "10_巾Max", "11_丈Max", "12_枚数1", "13_包数1", "14_枚数2", "15_包数2", "16_GW1", "17_風袋1", "18_GW2_x",
              "19_風袋2", "20_GW3_x", "21_風袋3", "22_正味_x", "23_風袋_x", "24_梱包ｺｰﾄﾞ_x", "25_EXｹｰｽNo", "26_Ex風袋1",
              "27_Ex風袋2", "28_Ex風袋3", "29_EXGW", "30_積み合せｺｰﾄﾞ", "31_積み合せNo", "32_ｺｲﾙ長さ_x", "33_ｺｲﾙ内径_x",
              "34_ｺｲﾙ肉厚_x", "35_ｺｲﾙ本数_x", "36_ｺｲﾙ巻数_x", "37_EX外寸高さ", "38_EX外寸奥", "39_EX外寸長さ", "40_x",
              "41_ﾗｲﾝ名") + tuple(f"{41 + i}_梱包高さ{i}" for i in range(1, MAX_MOUNT + 1))   # Input:1271-1278


def form_array(form: AinputForm, n: int) -> Any:
    """Input.FormArray (Input:1178-1300)。

    n=1: 梱包課用の 41 要素(変換済み) / n=2: ホスト電文 1 本の文字列 / n=3: 項目名と値(変換なし)の 56 行 × 2 列
    """
    a = _converted(form)
    order = ["A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12", "A13", "A14", "A15", "A16", "A17", "A18",
             "A19", "A20", "A21", "A22", "A23", "A24", "A25", "A26", "A27", "A28", "A29", "A30", "A31", "A32", "A33",
             "A34", "A35", "A36", "A37", "A38", "A39", "A40", "line"]
    if n == 1:
        return [NIS1, NIS2] + [a[k] for k in order]
    if n == 2:
        return NIS1 + NIS2 + "".join(a[k] for k in order)
    if n == 3:
        f = form
        values = [NIS1, NIS2, f.hiki_no, f.konno, f.remain, f.amin, f.hmin, f.tmin, f.amax, f.hmax, f.tmax,
                  f.mai1, f.cou1, f.mai2, f.cou2, f.gw, f.material1, "0", f.material2, "0", f.material3,
                  "0", "0", "0", f.ex_case_no, f.ex_material1, f.ex_material2, f.ex_material3, f.ex_gw, a["A30"],
                  a["A31"], "0", "0", "0", "0", "0", f.outside_atu, f.outside_haba, f.outside_take, "0", a["line"]]
        values += [get_k_height(form, i) for i in range(1, MAX_MOUNT + 1)]
        return [[lab, v] for lab, v in zip(CSV_LABELS, values)]
    return None


def telegram(form: AinputForm) -> str:
    """ホスト電文(hostsend.txt の 1 行)。"""
    return form_array(form, 2)


def csv_lines(form: AinputForm, sep: str = ",") -> list[str]:
    """Input.ExpCSV の行(Input:864-872)。各セルの後ろに区切りを付ける(末尾にもカンマ)。"""
    return ["".join(str(c) + sep for c in row) for row in form_array(form, 3)]


def _write_lines(path_text: str, rows: list[str]) -> Path:
    """`Open … For Output` + `Print #`(ANSI = cp932、CRLF。書けない文字は "?")。"""
    p = legacy_log.local_path(path_text)
    with open(p, "wb") as f:
        for r in rows:
            f.write((r + "\r\n").encode(CSV_ENCODING, "replace"))
    return p


def ainput_folder(uf1_line: str) -> str:
    """A_CsvSave のフォルダ名(Input:884-892)。UF1 のライン、どれでもなければ Test(ログと違い丸徳でない)。"""
    ln = lines.find(uf1_line)
    return ln.folder if ln else "Test"


def save_ainput_csv(form: AinputForm, lot: str, ken: str, *, now: datetime) -> str:
    """Input.A_CsvSave (Input:876-936)。`<検査履歴>\\<L>\\<L>_Ainput\\yyyy.mm\\<Lot>\\yyyymmdd_<Ken>(.nn).csv`。

    日付は暦日(夜勤の前日扱いをしない)。.99 まで埋まっていたら「連番をこれ以上付けることが出来ません。」で止める。
    書いたファイルのパス(Windows の書き方)を返す。
    """
    folder = ainput_folder(form.uf1_line)
    month_path = paths.join(paths.get("history"), folder, folder + "_Ainput", now.strftime("%Y.%m"))
    legacy_log.local_path(month_path).mkdir(parents=True, exist_ok=True)
    set_path = paths.join(month_path, lot)
    legacy_log.local_path(set_path).mkdir(parents=True, exist_ok=True)
    day_ken = now.strftime("%Y%m%d") + "_" + ken
    rows = csv_lines(form)
    first = paths.join(set_path, day_ken + ".csv")
    if not legacy_log.local_path(first).exists():
        _write_lines(first, rows)
        return first
    for n in range(2, 100):
        cand = paths.join(set_path, day_ken + "." + vbx.right("0" + str(n), 2) + ".csv")
        if not legacy_log.local_path(cand).exists():
            _write_lines(cand, rows)
            return cand
    raise Stop("連番をこれ以上付けることが出来ません。", source="Input:929")


def _tools_dir() -> str:
    """測定データ電子化のフォルダ。VBA は送信アプリと result.txt を `C:\\Tools\\測定データ電子化\\` の直書きで見る。

    使用パス nis が空(オフラインの既定)のときは、VBA と同じくオンラインの既定値を使う。
    """
    return paths.get("nis") or paths.default("nis", paths.ONLINE)


def hostsend_path() -> str:
    """hostsend.txt の場所(ExpCSV: NisPath、空ならアプリのフォルダ)。"""
    base = paths.get("nis") or paths.app_dir()
    return paths.join(base, HOSTSEND + ".txt")


def write_hostsend(form: AinputForm) -> str:
    """ExpCSV(…, "hostsend", NisPath, vbTab)(Input:858-862)。電文 1 行 + CRLF を cp932 で上書き。"""
    path = hostsend_path()
    _write_lines(path, [telegram(form)])
    return path


@dataclass
class OutputResult:
    csv_path: str
    hostsend_path: str


def a_output(form: AinputForm, lot: str, ken: str, *, now: datetime) -> OutputResult:
    """Input.A_Output (Input:805-835)。CSV(梱包課保存用)を先に、次に hostsend.txt を書く。"""
    csv = save_ainput_csv(form, lot, ken, now=now)
    host = write_hostsend(form)
    return OutputResult(csv, host)


# ------------------------------------------------------------------
# 送信アプリと結果
# ------------------------------------------------------------------
def host_con(tmp: str) -> str:
    """Input.HostCon (Input:1373-1382)。結果コード(右 4 文字)の説明。VBA では呼ばれていない。"""
    return {"A000": "正常に送信が行われました。",
            "A001": "作成されたテキストの内容が正しくありません。品質保証課へ連絡。",
            "A002": "ホストへの送信内容が正しくありません。NISへ連絡。",
            "A003": "ホストにてAインプットデータの登録に失敗しました。NISへ連絡。"}.get(vbx.right(tmp, 4), "")


def result_con(tmp: Optional[str]) -> str:
    """Input.ResultCon (Input:1384-1392)。result.txt の文 → 画面に出す文。"""
    t = "" if tmp is None else tmp
    return {
        "正常に送信が行われました。": t,
        "作成されたテキストの内容が正しくありません。品質保証課へ連絡してください。":
            "作成されたﾃｷｽﾄの内容が正しくありません。担当者へ連絡して下さい。" + "\r\n" + "A001",
        "ホストへの送信内容が正しくありません。NISへ連絡してください。":
            "ﾎｽﾄへの送信内容が正しくありません。担当者へ連絡して下さい。" + "\r\n" + "A002",
        "ホストへのAインプットデータの送信に失敗しました。NISへ連絡してください。":
            "ﾎｽﾄにてAｲﾝﾌﾟｯﾄﾃﾞｰﾀの送信に失敗しました。担当者へ連絡して下さい。" + "\r\n" + "A003",
    }.get(t, t)


def host_data_confi() -> str:
    """Input.HostDataConfi (Input:1394-1400)。いまの hostsend.txt の引当No と梱包No。

    VBA は C:\\Tools\\測定データ電子化\\hostsend.txt を UTF-8 として読む(書くときは ANSI)。無ければ ""。
    """
    p = legacy_log.local_path(paths.join(_tools_dir(), HOSTSEND + ".txt"))
    if not p.is_file():
        return ""
    try:
        buf = p.read_bytes().decode("utf-8", "replace")
    except OSError:
        buf = ""
    if buf.startswith("\ufeff"):
        buf = buf[1:]
    return "現添付Data【" + "引当No_" + vbx.mid(buf, 13, 8) + " " + "梱包No_" + vbx.mid(buf, 21, 2) + "】"


def _default_launcher(path: Path) -> bool:
    """Shell(exe, vbNormalFocus)。Windows 以外では起動しない(True を返して先へ進む)。"""
    if os.name != "nt":
        return True
    try:
        subprocess.Popen([str(path)], cwd=str(path.parent))
        return True
    except OSError:
        return False


@dataclass
class LaunchResult:
    ok: bool
    launched: bool
    messages: list[Message] = field(default_factory=list)


def app_activation(launcher: Optional[Callable[[Path], bool]] = None) -> LaunchResult:
    """Input.App_Activation (Input:1475-1487)。送信アプリを起動する。無い・失敗なら Ngflag="stop"(ok=False)。"""
    text = paths.join(_tools_dir(), SENDER_DIR, SENDER_EXE)
    p = legacy_log.local_path(text)
    if not p.is_file():
        return LaunchResult(False, False, [Message(text + "を開くことができません。" + "\r\n" + "担当者に連絡してください",
                                                   ERROR, source="Input:1478")])
    launched = os.name == "nt" or launcher is not None
    if not (launcher or _default_launcher)(p):
        return LaunchResult(False, False, [Message("失敗", INFO, source="Input:1484")])
    return LaunchResult(True, launched)


@dataclass
class ResultCheck:
    ok: bool                  # False = Ngflag "stop"
    text: str = ""            # result.txt の最終行
    missing: bool = False     # result.txt が無い(VBA は何もせず先へ進む)
    messages: list[Message] = field(default_factory=list)


def _last_line(data: bytes) -> Optional[str]:
    """`Line Input` を EOF まで回した最後の行(区切りは CR / CRLF)。空のファイルは None(Empty)。"""
    if not data:
        return None
    text = data.decode(CSV_ENCODING, "replace")
    parts = text.replace("\r\n", "\r").split("\r")
    if parts and parts[-1] == "":
        parts = parts[:-1]
    return parts[-1] if parts else ""


def inter_rupt_text(wait_ms: int = 0, *, sleep: Callable[[float], None] = time.sleep) -> ResultCheck:
    """Input.InterRuptText (Input:1402-1459)。待ってから result.txt の最終行を読む。

    5000 のときは 1 秒 × 5 回待つ(進捗表示)。「正常に送信が行われました。」以外は ResultCon の文を出して止める。
    """
    if wait_ms == 5000:
        for _ in range(5):
            sleep(1.0)
    else:
        sleep(wait_ms / 1000.0)
    p = legacy_log.local_path(paths.join(_tools_dir(), RESULT_FILE))
    if not p.is_file():
        return ResultCheck(True, missing=True)
    line = _last_line(p.read_bytes())
    if line == SUCCESS_TEXT:
        sleep(1.0)
        return ResultCheck(True, line)
    return ResultCheck(False, "" if line is None else line,
                       messages=[Message(result_con(line), WARN, source="Input:1445")])


# ------------------------------------------------------------------
# 入力欄のチェック(UFAinput の Change / BeforeUpdate)
# ------------------------------------------------------------------
def numeric_only(text: str) -> str:
    """`If IsNumeric(x) = False Then x = ""`(MAI1・COU1・GW・Material・Ex… など)。"""
    return text if vbx.is_numeric(text) else ""


def konno_before_update(konno: str) -> str:
    """UFAinput.KONNO_BeforeUpdate (UFAinput:375-378)。1 桁なら頭に 0、0 は空。"""
    if len(konno) == 1 and vbx.vcompare(konno, 0) != 0:
        konno = "0" + konno
    if _veq(konno, 0):
        konno = ""
    return konno


def remain_before_update(remain: str) -> str:
    """UFAinput.Remain_BeforeUpdate (UFAinput:541-543)。0 か 1 以外は空。"""
    if not (_veq(remain, 0) or _veq(remain, 1)):
        return ""
    return remain


def remain_change(form: AinputForm) -> None:
    """UFAinput.Remain_Change (UFAinput:545-573)。輸出欄が出ているとき、ストア(1)なら外寸を 0,0,0 に。"""
    form.remain = numeric_only(form.remain)
    if form.ex_visible:
        if _veq(form.remain, 1):
            form.outside_atu = form.outside_haba = form.outside_take = "0"
        elif _veq(form.remain, 0):
            form.outside_atu = form.outside_haba = form.outside_take = ""


def pile_match_change(form: AinputForm) -> list[Message]:
    """UFAinput.PileMatch_Change (UFAinput:476-517)。ストア(1)のときストアカード 7 桁の頭文字をロットと照合する。

    ロットの頭が S / F なら "L" でもよい。積合せ(PileMatchNo あり)のときは数でなければ消す。
    """
    msgs: list[Message] = []
    if form.remain == "1":
        if len(form.pile_match) != 7:
            return msgs
        store, lot_h = form.pile_match[:1], form.lot[:1]
        flag = store == lot_h
        if lot_h in ("S", "F") and store == "L":
            flag = True
        if not flag:
            msgs.append(Message("ストアカードの頭文字が違います", ERROR, source="UFAinput:498"))
            form.pile_match = ""
            return msgs
    elif form.remain == "0" and form.pile_match_no == "":
        form.pile_match = ""
        return msgs
    if form.pile_match_no != "":
        if not vbx.is_numeric(form.pile_match):
            form.pile_match = ""
    return msgs


def pile_match_key(remain: str, ch: str) -> str:
    """UFAinput.PileMatch_KeyPress (UFAinput:519-531)。ストアのときは英数字だけ(小文字は大文字に)。"""
    if not _veq(remain, 1):
        return ch
    o = ord(ch) if len(ch) == 1 else 0
    if 48 <= o <= 57 or 65 <= o <= 90:
        return ch
    if 97 <= o <= 122:
        return chr(o - 32)
    return ""


def pile_match_no_change(form: AinputForm) -> None:
    """UFAinput.PileMatchNo_Change (UFAinput:533-539)。1 か 9 だけ。1(積合せ)なら 引当No & 梱包No を入れる。"""
    form.pile_match_no = numeric_only(form.pile_match_no)
    if not (_veq(form.pile_match_no, 9) or _veq(form.pile_match_no, 1)):
        form.pile_match_no = ""
    if not _veq(form.pile_match_no, 1):
        form.pile_match = ""
    if _veq(form.pile_match_no, 1):
        form.pile_match = form.hiki_no + form.konno


# ------------------------------------------------------------------
# 引当の一覧(Hiki_Lot)
# ------------------------------------------------------------------
@dataclass
class HikiResult:
    rows: Optional[list[list[Any]]]       # Hiki_Lot に貼る行(None = 貼らない)
    recreate: bool = False                # AddWS("Hiki_Lot")(空にする)
    db_flag: bool = False                 # UFAinput.DBflag = "true"
    messages: list[Message] = field(default_factory=list)


def no_search(lot: str, history: Optional[list[list[Any]]]) -> HikiResult:
    """Input.No_Search (Input:77-223)。Lot情報履歴(1 行目=見出し)から引当を集める。無ければ SIKACOMP から。"""
    if history is None or lot == "":
        return HikiResult(None)
    head = history[0]
    want = {"ﾛｯﾄNo": 0, "引当NO": 0, "ｵｰﾀﾞｰNo": 0, "検査No": 0, "引当調整NO": 0, "引当数量": 0,
            "引当ﾌｧｲﾙ登録": 0, "包装仕様NO": 0}
    for i, name in enumerate(head, start=1):
        if name in want:
            want[name] = i
    if want["ﾛｯﾄNo"] == 0:
        raise IndexError("Lot情報履歴 に ﾛｯﾄNo 列がありません(VBA の実行時エラー 9)")
    out = []
    for row in history:
        if _veq(_hist(row, want["ﾛｯﾄNo"]), lot):
            out.append([_hist(row, want[k]) if want[k] else None
                        for k in ("引当NO", "ｵｰﾀﾞｰNo", "ﾛｯﾄNo", "検査No", "引当調整NO", "引当数量", "引当ﾌｧｲﾙ登録",
                                  "包装仕様NO")])
    if out:
        return HikiResult(out, recreate=True)
    res = HikiResult(None, db_flag=True)
    db = paths.db_file("SIKACOMP")
    if not datasource.exists(db):
        res.messages.append(Message("日稲:" + "SIKACOMP.accdb" + "がありません" + "_ReadCheck_No_Search", ERROR,
                                    source="Input:154"))
        return res
    recs = datasource.get_records(db, "仕掛", f"[ﾛｯﾄ番号]='{lot}'")
    if recs is None:
        res.messages.append(Message("SIKACOMP：該当ロットなし" + "_No_Search:empty", INFO, source="Input:159"))
        res.recreate = True
        return res
    fields = datasource.get_fields(db, "仕掛") or []
    idx = {k: datasource.column_index(fields, k) for k in ("ﾛｯﾄ番号", "引当番号", "受注番号", "検査番号", "引当調整NO",
                                                          "引当数量")}
    rows = [[r[idx[k]] if idx[k] >= 0 else None for k in ("引当番号", "受注番号", "ﾛｯﾄ番号", "検査番号", "引当調整NO",
                                                          "引当数量")]
            for r in recs if r[idx["ﾛｯﾄ番号"]] == lot]
    if not rows:
        res.messages.append(Message("SIKACOMP：該当Lotなし" + "_No_Search:LotNo", INFO, source="Input:187"))
        res.recreate = True
        return res
    res.rows = rows
    res.recreate = True
    return res


def sort_hiki(values: list[str]) -> list[str]:
    """UFAinput.UserForm_Activate のバブルソート(UFAinput:828-847)。数は Val の昇順、数でないものは後ろ(文字列順)。"""
    arr = [str(v) for v in values if str(v) != ""]
    n = len(arr)
    for i in range(0, n - 1):
        for j in range(0, n - 1 - i):
            a, b = arr[j], arr[j + 1]
            na, nb = vbx.is_numeric(a), vbx.is_numeric(b)
            if na and nb:
                swap = vbx.val(a) > vbx.val(b)
            elif na and not nb:
                swap = False
            elif not na and nb:
                swap = True
            else:
                swap = a > b
            if swap:
                arr[j], arr[j + 1] = b, a
    return arr


@dataclass
class HikiSelect:
    order_no: str = ""
    lot: Optional[str] = None
    ken_no: Optional[str] = None
    adjust_label: str = ""
    adjust_no: str = ""
    adjust_visible: bool = True
    export_visible: Optional[bool] = None
    form_height: Optional[int] = None


def hiki_no_change(selected: Optional[str], hiki_rows: list[list[Any]], *, db_flag: str = "") -> HikiSelect:
    """UFAinput.HikiNo_Change (UFAinput:317-367)。選んだ引当No の行から オーダーNo・調整No を出す。

    オーダーNo の 3 文字目が 1 なら輸出(ExFrame を出す)。selected=None はリストから選ばれていない(ListIndex=-1)。
    """
    out = HikiSelect()
    if selected is None:
        return out
    row = None
    for r in hiki_rows:
        if _nz(r[0] if r else None) == selected:
            row = r
            break
    if row is None:
        return out

    def c(i: int) -> str:
        return _nz(row[i - 1] if i - 1 < len(row) else None)
    out.order_no = c(2)
    if db_flag == "true":
        out.lot, out.ken_no = c(3), c(4)
    out.adjust_label = "調整No"
    out.adjust_no = c(5)
    if out.adjust_no == "0":
        out.adjust_label = "引当数量"
        out.adjust_no = c(6)
    if out.adjust_no == "0":
        out.adjust_visible = False
    ex_flag = vbx.clng(vbx.mid(out.order_no, 3, 1))
    out.export_visible = ex_flag == 1
    out.form_height = 620 if ex_flag == 1 else 490
    return out


# ------------------------------------------------------------------
# 起動
# ------------------------------------------------------------------
@dataclass
class StartupResult:
    form: Optional[AinputForm]
    messages: list[Message] = field(default_factory=list)
    log: Optional[legacy_log.LogEntry] = None
    hiki: Optional[HikiResult] = None


def startup(temp: str, ken_m: str, history: Optional[list[list[Any]]], *, working_lot: str, lname: str,
            uf1_line: str = "", uf1_minmax: Optional[dict[str, str]] = None,
            sheet_cells: Optional[dict[tuple[int, int], Any]] = None, workers: str = "") -> StartupResult:
    """Input.Ainput_Startup (Input:225-400)。UFAinput を開くときの値。

    temp: "UF1" / "Cask" / "Mitsubishi" / "H494"。history は Lot情報履歴(1 行目=見出し)。
    uf1_minmax は UF1 の AMin..TMax、sheet_cells は キャスク検査/全ヵ所検査/H494検査 のセル(行, 列)。
    workers は UF1 なら ` T1:<T1>`(workers.workers_text)、キャスク・三菱は "UFｷｬｽｸ:<Kname>" などを呼び手が作る。
    """
    res = StartupResult(None)
    if history is None:
        res.messages.append(Message("Lot情報履歴ｼｰﾄがありません" + "_Ainput_Startup", ERROR, source="Input:227"))
        return res
    rows = history[1:]
    match = None
    for r in rows:
        if _veq(_hist(r, 3), working_lot):
            match = r
            break
    form = AinputForm(line_name=lname, uf1_line=uf1_line)
    if match is not None:
        form.lot = working_lot
        form.ken_no = _nz(_hist(match, 7))
        form.zai = _nz(_hist(match, 8))
        form.tyo = _nz(_hist(match, 9))
        form.atu = vbx.fmt(_hist(match, 4), "0.000")
        form.haba = vbx.fmt(_hist(match, 5), "0.0")
        form.take = vbx.fmt(_hist(match, 6), "0.0")
        form.ken_law = ken_m
        form.youto = _nz(_hist(match, 10))
    else:
        if temp != "UF1":
            return res
        form.db_flag = "true"
        res.messages.append(Message("Lot情報履歴：履歴該当なし。" + "\r\n" + "*DB検索start(15秒程かかります)" + "\r\n"
                                    + "_Ainput_Startup", INFO, source="Input:283"))
        db = paths.join(paths.get("next"), "SIKATWOYEAR.sqlite3")
        if not datasource.exists(db):
            res.messages.append(Message("日稲" + "SIKATWOYEAR.accdb" + "がありません" + "_ReadCheck_Ainput_Startup",
                                        ERROR, source="Input:297"))
            return res
        recs = datasource.get_records(db, "仕掛", f"[ﾛｯﾄ番号]='{working_lot}'")
        if recs is None:
            res.messages.append(Message("SIKACOMP：該当ロットなし" + "_Ainput_Startup", INFO, source="Input:302"))
            return res
        fields = datasource.get_fields(db, "仕掛") or []

        def col(name: str) -> str:
            i = datasource.column_index(fields, name)
            if i < 0:
                raise IndexError(f"{name} 列がありません(VBA の実行時エラー 9)")
            return recs[0][i]
        form.lot = working_lot
        form.ken_no = col("検査番号")
        form.zai = col("製造材質")
        form.tyo = col("製造調質")
        form.atu = vbx.fmt(col("製造板厚"), "0.000")
        form.haba = vbx.fmt(col("製造板幅"), "0.0")
        form.take = vbx.fmt(col("製造板丈"), "0.0")
        form.ken_law = ken_m
    cells = sheet_cells or {}

    def put(attr: str, rc: tuple[int, int], pattern: str) -> None:
        v = _cell(cells.get(rc))
        if vbx.vcompare(v, 0) != 0:
            setattr(form, attr, vbx.fmt(v, pattern))
    if temp == "UF1":
        mm = uf1_minmax or {}
        for k in ("amin", "hmin", "tmin", "amax", "hmax", "tmax"):
            setattr(form, k, mm.get(k, ""))
    elif temp == "Mitsubishi":
        put("amin", (4, 23), "0.000"), put("hmin", (4, 25), "0.0"), put("tmin", (4, 27), "0.0")
        put("amax", (3, 23), "0.000"), put("hmax", (3, 25), "0.0"), put("tmax", (3, 27), "0.0")
    elif temp in ("Cask", "H494"):
        put("amin", (4, 22), "0.000"), put("hmin", (4, 23), "0.0"), put("tmin", (4, 24), "0.0")
        put("amax", (3, 22), "0.000"), put("hmax", (3, 23), "0.0"), put("tmax", (3, 24), "0.0")
    res.form = form
    res.log = legacy_log.line_entry(lname, "*UFKen:" + "UFAinput起動" + workers)
    res.hiki = no_search(working_lot, history)
    if res.hiki.db_flag:
        form.db_flag = "true"
    form.ex_visible = False
    return res


# ------------------------------------------------------------------
# 「Aインプット」ボタン(2 回押し)
# ------------------------------------------------------------------
@dataclass
class AinResult:
    stage: str           # dbflag / missing / busy / first / input_error / height_ng / question / cancel / stopped / sent
    messages: list[Message] = field(default_factory=list)
    question: Optional[Question] = None
    height: Optional[HeightCheckResult] = None
    output: Optional[OutputResult] = None
    launch: Optional[LaunchResult] = None
    result: Optional[ResultCheck] = None
    daily_writes: list[Any] = field(default_factory=list)
    logs_written: list[legacy_log.LogEntry] = field(default_factory=list)


def _m_count(form: AinputForm) -> int:
    if form.mai_flag == "1":
        return vbx.clng(vbx.val(form.mai1) * vbx.val(form.cou1))
    if form.mai_flag == "2":
        return vbx.clng(vbx.val(form.mai1) * vbx.val(form.cou1) + vbx.val(form.mai2) * vbx.val(form.cou2))
    return 0


def _in_1_to_9(v: str) -> bool:
    """`Select Case v: Case 1 To 9`(文字列と数の比較。空欄は U1 の実機確認待ち → TypeMismatch)。"""
    return vbx.vcompare(v, 1) >= 0 and vbx.vcompare(v, 9) <= 0


def _lock(form: AinputForm) -> None:
    form.tag = ""
    form.locked = True


def ain_button(form: AinputForm, *, history_rows: Optional[list[list[Any]]], answers: dict[str, bool],
               registry: WorkerRegistry, day_max: str, now: datetime, daily: Optional[DailyContext] = None,
               vc: Callable[[str], tuple[float, list[Message]]] = vc_conv,
               launcher: Optional[Callable[[Path], bool]] = None,
               sleep: Callable[[float], None] = time.sleep, computer: Optional[str] = None) -> AinResult:
    """UFAinput.AinButt_Click (UFAinput:932-1137)。

    1 回目(高さラベルが出ていない、または未入力の山がある): 山数から高さの入力欄を出す。
    2 回目(全山入力済み): 入力チェック → 高さチェック → 「ｱｲｶｽ転送を行いますか？」→ ログ → CSV・hostsend.txt →
    送信アプリ → 1 秒 → result.txt(5 秒待つ)→ 日報(Daily_Report(4))。
    問いは `Question("ainput_send")` を返して止まる。答えを answers に入れて呼び直す(そのため Tag は戻しておく)。
    history_rows は Lot情報履歴 の 2 行目から。daily を渡すと Daily_Report(4) の書き込みを返す。
    """
    res = AinResult("")
    if form.db_flag == "true":
        res.stage = "dbflag"
        res.messages.append(Message("Lot情報履歴に履歴が無いため" + "\r\n" + "A_input転送は使用できません。" + "\r\n"
                                    + "_AinButt:DBflag", WARN, source="UFAinput:935"))
        return res
    if form.hiki_no == "":
        res.stage = "missing"
        res.messages.append(Message("必要項目の入力がありません" + "(引当No)", ERROR, source="UFAinput:939"))
        return res
    if form.tag == "Load":
        res.stage = "busy"
        return res
    form.ai_text = form.ga_text = form.vc_text = ""
    form.tag = "Load"
    m_count = _m_count(form)
    mount_height = calc_mount_height(form, _mount_count(form))
    ain_mc = _mount_count(form)

    def stop_input(text: str, level: str, src: str) -> AinResult:
        form.tag = ""
        res.stage = "input_error"
        res.messages.append(Message(text, level, source=src))
        return res

    if form.l1_visible and all_height_filled(form, ain_mc):
        if get_k_height(form, 1) == "" or form.mount_count == "" or form.hiki_no == "":
            return stop_input("必要項目の入力がありません" + "(高さ測定値,山数,引当No)", ERROR, "UFAinput:973")
        if form.mai_flag == "1" and (form.mai1 == "" or form.cou1 == ""):
            return stop_input("必要項目の入力がありません" + "(枚数1,包数1)", ERROR, "UFAinput:980")
        if form.mai_flag == "2" and (form.mai1 == "" or form.cou1 == "" or form.mai2 == "" or form.cou2 == ""):
            return stop_input("必要項目の入力がありません" + "(枚数1,包数1,枚数2,包数2)", ERROR, "UFAinput:986")
        if form.ex_visible:
            if vbx.vcompare(form.remain, 1) != 0 and form.ex_case_no == "":
                return stop_input("ｹｰｽNoの入力がありません。", WARN, "UFAinput:994")
            fl1 = False
            for v in (form.outside_atu, form.outside_haba, form.outside_take):
                if _in_1_to_9(v):
                    fl1 = True
            fl2 = form.outside_atu == "" or form.outside_haba == "" or form.outside_take == ""
            if fl1:
                return stop_input("EX外寸に在り得ない数値の入力があります。", WARN, "UFAinput:1014")
            if fl2:
                return stop_input("EX外寸の入力がありません｡", WARN, "UFAinput:1019")
        if form.remain == "1" and form.pile_match == "":
            return stop_input("ｽﾄｱｶｰﾄﾞNo入力がありません。", WARN, "UFAinput:1027")
        hc = height_check(vbx.val(form.amin), vbx.val(form.amax), vbx.val_single(mount_height),
                          vbx.val(m_count), form.order_no, form, history_rows, vc=vc)
        res.height = hc
        res.messages += hc.messages
        if hc.ng:
            res.stage = "height_ng"
            entry = legacy_log.msg_error_entry(hc.msg_error, form.line_name, registry.t1, registry.t2, registry.t3)
            entry.write(now=now, computer=computer)
            res.logs_written.append(entry)
            return res
        ans = answers.get("ainput_send")
        if ans is None:
            form.tag = ""
            res.stage = "question"
            res.question = Question("ainput_send", "ｱｲｶｽ転送を行いますか？", "yesno", source="UFAinput:1034",
                                    messages=list(res.messages))
            return res
        if not ans:
            form.tag = ""
            res.stage = "cancel"
            return res
        wk = workers_text(day_max, registry)
        ln = lines.find(form.line_name)
        sep = "枚_" if (ln is not None and ln.official != "中板") else "枚\\"
        entry = legacy_log.line_entry(form.line_name, "*UFKen:" + str(m_count) + sep
                                      + vbx.cstr(mount_height, is_single=True) + "mm *" + wk)
        entry.write(now=now, computer=computer)
        res.logs_written.append(entry)
        res.messages.append(Message("", INFO, sound=WAIT_SOUND, source="UFAinput:1050"))
        res.output = a_output(form, form.lot, form.ken_no, now=now)
        res.launch = app_activation(launcher)
        res.messages += res.launch.messages
        if not res.launch.ok:
            form.ng_flag = "stop"
            form.tag = ""
            res.stage = "stopped"
            return res
        sleep(1.0)
        res.result = inter_rupt_text(5000, sleep=sleep)
        res.messages += res.result.messages
        if not res.result.ok:
            form.ng_flag = "stop"
            form.tag = ""
            res.stage = "stopped"
            return res
        if daily is not None:
            daily.ainput = AinputValues(form.lot, form.konno, form.mai1, form.cou1, form.mai2, form.cou2)
            res.daily_writes = daily_report(4, daily)
        res.messages.append(Message("", INFO, sound=DONE_SOUND, source="UFAinput:1086"))
        form.konno = ""                 # 旧 K_Height 欄は消すが、山高さ(動的ラベル)は消さない(VBA どおり)
        res.stage = "sent"
    else:
        if form.mount_count != "":
            form.l1_visible = True
            form.l2_visible = True
            form.l1_caption = l1_caption(l_height(form))
            mc = _mount_count(form)
            if 1 <= mc <= MAX_MOUNT:
                show_height_controls(form)
            else:
                form.l1_visible = False
                form.l2_visible = False
            res.stage = "first"
        else:
            form.tag = ""
            res.stage = "input_error"
            res.messages.append(Message("山数入力無し", ERROR, source="UFAinput:1126"))
            return res
    _lock(form)
    return res


def recheck_height(form: AinputForm, history_rows: Optional[list[list[Any]]], *,
                   vc: Callable[[str], tuple[float, list[Message]]] = vc_conv) -> HeightCheckResult:
    """UFAinput.CommandButton9_Click (UFAinput:250-277)。外注で合紙・VC・外装紙を外したあとの再計算。"""
    form.ai_text = form.ga_text = form.vc_text = ""
    m_count = _m_count(form)
    mount_height = calc_mount_height(form, _mount_count(form))
    hc = height_check(vbx.val(form.amin), vbx.val(form.amax), vbx.val_single(mount_height), vbx.val(m_count),
                      form.order_no, form, history_rows, vc=vc)
    form.gaityuu_text = ""
    form.locked = True
    return hc


def unlock_gaityuu(form: AinputForm) -> list[Message]:
    """UFAinput.CommandButton10_Click (UFAinput:107-115)。合紙・VC・外装紙のチェックを変えられるようにする。"""
    if form.locked:
        form.locked = False
        return [Message("合紙,VCの有無を変更したい場合" + "\r\n" + "赤BOXに 外注 と入力してください。", INFO,
                        source="UFAinput:113")]
    return []


def gaityuu_log(uf1_line: str) -> legacy_log.LogEntry:
    """UFAinput.GaityuuText_BeforeUpdate (UFAinput:49-61)のログ。"""
    return legacy_log.line_entry(uf1_line, "*UFKen:計算高さ操作")


def clear_and_resend(form: AinputForm, *, launcher: Optional[Callable[[Path], bool]] = None,
                     sleep: Callable[[float], None] = time.sleep) -> tuple[str, LaunchResult]:
    """UFAinput.CommandButton2_Click (UFAinput:127-176)。入力を消して hostsend.txt を書き直し、送信アプリを起動。"""
    for k in ("hiki_no", "remain", "konno", "mai1", "cou1", "mai2", "cou2", "material1", "material2", "material3",
              "gw", "pile_match_no", "pile_match", "mount_count", "ex_case_no", "ex_material1", "ex_material2",
              "ex_material3", "ex_gw", "outside_atu", "outside_haba", "outside_take"):
        setattr(form, k, "")
    clear_all_k_height(form)
    form.l1_visible = False
    form.l2_visible = False
    path = write_hostsend(form)
    sleep(1.0)
    return path, app_activation(launcher)
