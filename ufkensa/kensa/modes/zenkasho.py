"""全ヵ所検査(一般 / 5N / 三菱受託)— UFMitsubishi(P1) + 全ヵ所検査シート + Mitsubishi モジュール

- シートのイベント: `ZenSh1`(= `ZenSh`)の Worksheet_Change / SelectionChange。`ZenSh2` は旧版の残骸(移植しない)
- 作業選択(Choice = 三菱 / 5N / 一般)と `ForN_Select`、ロット読込(H494 モードと共用)、検入高さ
- MAX/MIN、`ｺﾋﾟ-用`・`転送済`(三菱受託原紙)への転記、測定機器の文字列、印刷の設定、保存の手順

三菱受託モードは 2024.1.11 の列追加に追従していない疑いがある(W2)。**VBA どおりの列で再現**し、
ずれていると思われる所は関数の docstring と `MITSUBISHI_SUSPECT` に書いた。
"""
from __future__ import annotations

import unicodedata
from typing import Any, Callable, Optional, Union

from .. import datasource, flow, paths, quirks, vbx
from . import cask as ck
from . import sheet as sh
from .sheet import H494, ZEN, Book, Call, Effects, LogCall, PrintJob, SavePlan, SaveStep, VbaError

#: 三菱受託モードで列がずれていると思われる所(W2。VBA どおりに再現している)
MITSUBISHI_SUSPECT = (
    "ﾌﾞﾛｯｸNO の入力チェック(UFMit:1532-1537)が列 27(AA=丈歪)・29(AC=ｹﾞｰｼﾞ)を見ている。元板NO は AE(31)・切断ﾌﾞﾛｯｸNO は AG(33)",
    "歪・外観の入力チェック(Mit:826-855)が列 24(X=ﾗﾃﾗﾙﾎﾞｰ判定)・25(Y=幅歪)・26(Z=幅歪判定)を見ている",
    "転送済の N(対角線差)に対角の判定記号(V 列)、O(歪 幅)にﾗﾃﾗﾙﾎﾞｰ、P(歪 長さ)に幅歪、Q(外観)に丈歪が入る",
    "転送済の B(原板積載)にｹﾞｰｼﾞ(AC)、C(切断ブロック)には削除後の空の S 列が入る。元板NO・切断ﾌﾞﾛｯｸNO はコピー範囲 A:AD の外",
)

#: SIKALOTNOW `仕掛` の欄名(UFMit:639-656。こちらは半角で正しい)
MIT_FIELDS = ("ｵｰﾀﾞｰ番号", "ﾛｯﾄ番号", "製造板厚", "製造板幅", "製造板丈", "検査番号", "製造材質", "製造調質", "用途ｺｰﾄﾞ",
              "設計_設備ｺｰｽ", "取引先名", "納入先名", "送り先名", "BOX実績_枚本数", "BOX実績_良品重量", "梱包_合紙", "VC",
              "品質ｸﾞﾚｰﾄﾞ_ﾗﾃﾗﾙﾎﾞｰ")

LINE_DEVICE = {"L-1": ("L-1", " C11"), "HVC": ("HVC", " C11"), "LVC": ("LVC", " C10"),
               "ﾊﾞﾗﾝｻｰ": ("ﾊﾞﾗﾝｻｰ", " C11"), "ﾄｯﾄ": ("ﾄｯﾄ", " C12")}
FIVE_LINES = ("L-1", "LVC", "HVC", "ﾊﾞﾗﾝｻｰ", "ﾄｯﾄ")


# ------------------------------------------------------------------
# 文字列の小道具(Mitsubishi / 補助)
# ------------------------------------------------------------------
def remove_duplicates(tmp: str, n: str = "") -> str:
    """Mitsubishi.RemoveDuplicates (Mit:77-94)。**部分一致**で重複とみなす(「田中 中」→「田中」)。"""
    if n != "":
        words = tmp.split(n)
    else:
        words = tmp.split(" ")
        n = " "
    unique = ""
    for w in words:
        if _instr_text(unique, w + n) == 0:
            unique = unique + w + n
    return unique.strip(" ")


def _instr_text(s: str, sub: str) -> int:
    """`InStr(1, s, sub, vbTextCompare)`"""
    a = unicodedata.normalize("NFKC", s).casefold()
    b = unicodedata.normalize("NFKC", sub).casefold()
    if b == "":
        return 1
    return a.find(b) + 1


def extract_before(temp: str, n: str = "÷") -> str:
    """Mitsubishi.ExtractBefore (Mit:1426-1434)。n が無ければ ""。"""
    pos = vbx.instr(temp, n)
    return vbx.left(temp, pos - 1) if pos > 0 else ""


def extract_nth_text(s: str, a: int) -> str:
    """Mitsubishi.ExtractNthText (Mit:1398-1425)。a 番目の ( ) の中の "÷" より前。"""
    if a < 1:
        return ""
    start = end = 0
    for _ in range(a):
        start = vbx.instr(s, "(", start + 1)
        if start == 0:
            return ""
        end = vbx.instr(s, ")", start)
        if end == 0:
            return ""
    if start > 0 and end > start:
        return extract_before(vbx.mid(s, start + 1, end - start - 1))
    return ""


def convert_to_half(temp: Any) -> str:
    """補助.ConvertToHalf + numExtract (補助:1105-1130)。全角数字を半角にし、"/" の前の数字と "." だけ残す。"""
    s = "" if temp is None else str(temp)
    s = "".join(chr(ord(ch) - 0xFEE0) if "０" <= ch <= "９" else ch for ch in s)
    if "/" in s:
        s = s.split("/")[0]
    return "".join(ch for ch in s if ch in "0123456789.")


def check_red_x(ws: sh.Sheet, r1: int, c1: int, r2: int, c2: int) -> int:
    """Mitsubishi.CheckRedXCells (Mit:1387-1397)。背景 RGB(255,0,0) かつ "×" の数(判定付けの ColorIndex 3)。"""
    n = 0
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            if ws.interior.get((r, c)) == (255, 0, 0) and ws.get(r, c) == "×":
                n += 1
    return n


# ------------------------------------------------------------------
# シートのイベント(ZenSh1)
# ------------------------------------------------------------------
_RAN = ((2, 2, 3, 6), (1, 11, 1, 15), (1, 17, 1, 19), (2, 11, 4, 15), (2, 17, 4, 19), (2, 22, 4, 28),
        (7, 2, 76, 2), (7, 4, 76, 4), (7, 6, 76, 6), (7, 8, 76, 8), (7, 11, 76, 11), (7, 13, 76, 13),
        (7, 15, 76, 15), (7, 17, 76, 17), (7, 19, 76, 21), (7, 23, 76, 23), (7, 25, 76, 25), (7, 27, 76, 27),
        (7, 29, 76, 30))
_HANTEI_COLS = (3, 5, 7, 9, 12, 14, 16, 18, 22, 24, 26, 28)


def worksheet_change(book: Book, target: sh.Sheet, r: int, c: int) -> None:
    """ZenSh1.Worksheet_Change (ZenSh:9-125)。On Error Resume Next の下で動く。"""
    ws = book.active
    befor = book.resume(lambda: sh.take_befor(book, book.mit, r, c), "ZenSh:19") or ""
    TR, TC = r, c
    if ws.name == ZEN and 7 <= TR <= 76:
        if 2 <= TC <= 8:
            book.resume(lambda: book.activate(TR, TC + 1))
        elif TC == 9:
            if sh.blank(ws.get(TR, 11)):
                book.resume(lambda: book.activate(TR, 11))
            else:
                book.resume(lambda: book.activate(TR + 1, 2))
        elif 11 <= TC <= 18:
            book.resume(lambda: book.activate(TR, TC + 1))
        elif TC == 19:
            if not sh.blank(ws.get(TR, 19)):
                book.fx.sounds.append("音楽対角")
                book.resume(lambda: sh.uf_taikaku(book, TR, TC))
                book.events_enabled = False
                last = 23 if quirks.reproduce("B32") else 22     # B32: ラテラルボー(W)まで消える
                for cc in range(20, last + 1):
                    book.write(ws, TR, cc, "")
                book.events_enabled = True
                book.resume(lambda: book.activate(TR, TC + 1))
        elif TC == 20:
            book.events_enabled = False
            book.resume(lambda: sh.al_ter_range(book, TR, TC, befor), "AlTerRange")
            book.resume(lambda: sh.taikaku(book, TR, TC), "対角")
            book.events_enabled = True
            book.resume(lambda: book.activate(TR, 23))
            return
        elif TC in (21, 22):
            book.resume(lambda: book.activate(TR, 23))
        elif TC == 23:
            book.events_enabled = False
            v = ws.get(TR, 23)
            if not sh.is_numeric(v):
                pass
            elif v == "-":
                pass
            elif not sh.blank(v):
                book.resume(lambda: book.write(ws, TR, 23, vbx.fmt(sh.unit_record(sh.val_cell(v)), "0.0")))
            book.resume(lambda: sh.lateralbow_z(book, TR, TC), "Lateralbow_Z")
            book.events_enabled = True
            return
        elif TC in (24, 26, 28):
            book.resume(lambda: book.activate(TR, TC + 1))
        elif TC in (25, 27):
            book.events_enabled = False
            v = ws.get(TR, TC)
            if not (not sh.is_numeric(v) or sh.blank(v)):
                book.resume(lambda: sh.hizumi_z(book, TR, TC), "Hizumi_Z")
                book.events_enabled = True
                return
    if ws.name == ZEN:
        book.events_enabled = False
        book.resume(lambda: sh.al_ter_range(book, TR, TC, befor), "AlTerRange")
        if not sh.is_numeric(target.get(TR, TC)):
            return                                     # EnableEvents は False のまま(VBA どおり)
        book.events_enabled = True
        if TC in (2, 4, 6, 8, 11, 13, 15, 17, 19, 20):
            book.resume(lambda: sh.kensa_hantei(book, TR, TC), "検査判定")
        book.events_enabled = False
        book.events_enabled = True
        book.resume(lambda: sh.koshagai_kensaku(book, TR, TC), "公差外検索")


def worksheet_selection_change(book: Book, r: int, c: int) -> None:
    """ZenSh1.Worksheet_SelectionChange (ZenSh:127-212)。"""
    book.mit.cnt = "0"
    ws = book.active
    if ws.name == ZEN:
        if c == 2:
            book.events_enabled = False
            book.write(ws, 2, 8, book.mit.ken_no)
            book.write(ws, 3, 8, book.mit.kname)
            book.write(ws, 4, 8, book.mit.youto_no)
            for rr in (2, 3, 4):
                ws.font_size[(rr, 8)] = 5
            book.events_enabled = True
        elif c in (22, 24, 26, 28):
            book.activate(r, c + 1)
    if ws.name == ZEN and (r, c) == (7, 4):
        book.events_enabled = False
        if not sh.blank(ws.get(7, 2)):
            book.write(ws, 1, 2, sh.fmt_date(book.now(), "yyyy/mm/dd h:mm"))
            ws.font_size[(1, 2)] = 5
            book.fx.calls.append(Call("Daily_Report", (1,)))
            book.fx.calls.append(Call("Daily_Report", (3,)))
        book.events_enabled = True
    if ws.name != ZEN:
        return
    if 7 <= r <= 76 and c in _HANTEI_COLS:
        for rect in _RAN:
            ws.fill(*rect, sh.WHITE)
        return
    if ck._in(_RAN, r, c):
        for rect in _RAN:
            ws.fill(*rect, sh.WHITE)
        ws.interior[(r, c)] = sh.LIGHT_BLUE
        ws.font_color[(r, c)] = (0, 0, 0)


sh.HANDLERS[ZEN] = (worksheet_change, worksheet_selection_change)


# ------------------------------------------------------------------
# 作業選択
# ------------------------------------------------------------------
def _set_choice(book: Book, text: str) -> None:
    """`UFMitsubishi.Choice = text`。値が変わったときだけ Choice_Change が走る。"""
    if book.mit.choice != text:
        book.mit.choice = text
        choice_change(book)


def choice_change(book: Book) -> Effects:
    """UFMitsubishi.Choice_Change (UFMit:509-515)。"""
    fx = book.fx
    if book.mit.choice == "５N":
        book.mit.choice = "5N"
    if book.mit.choice in ("三菱", "一般", "5N"):
        forn_select(book)
    else:
        fx.msg("正しい作業選択名を入力してください", flow.WARN, "UFMit:513")
    return fx


def forn_select(book: Book) -> None:
    """Mitsubishi.ForN_Select (Mit:1448-1513)。作業選択に応じたフラグ・V1・5N の歪公差(M5=2.0 / O5=3.0)・検入高さの復元。

    シートへの書き込みは EnableEvents を止めずに行う(VBA どおり。V1 への書き込みで Change が走り、
    文字なので EnableEvents が False のまま終わる)。
    """
    f = book.mit
    if f.tag == "off":
        return
    if ZEN not in book.sheets:
        return
    ws = book.ws(ZEN)
    book.activate_sheet(ZEN)
    if f.choice == "三菱":
        f.work, f.inflag = "ON", "flagon"
        f.normal_visible, f.mitsubishi_visible, f.fini_visible = False, True, False
        book.write(ws, 1, 22, "三菱")
    elif f.choice == "5N":
        f.work, f.inflag = "OFF", ""
        f.normal_visible, f.mitsubishi_visible, f.fini_visible = False, False, True
        book.write(ws, 1, 22, "5N")
        f.set_option("5N")
        book.write(ws, 5, 13, "2.0")
        book.write(ws, 5, 15, "3.0")
    elif f.choice == "一般":
        f.work, f.inflag = "OFF", ""
        f.normal_visible, f.mitsubishi_visible, f.fini_visible = True, False, False
        book.write(ws, 1, 22, "一般")
        f.set_option("一般")
    ad3 = ws.get(3, 30)
    if not sh.blank(ad3):
        if vbx.right(sh.cell_str(ad3), 1) == "枚":
            f.yama_index = 6
            if not sh.blank(ws.get(3, 2)):
                book.fx.calls.append(Call("DrawListBoxContent", (sh.val_cell(ws.get(3, 2)),)))
            f.kenh[7] = extract_before(sh.cell_str(ad3), "枚")
    ad2 = ws.get(2, 30)
    if not sh.blank(ad2):
        f.yama_index = vbx.clng(vbx.val(vbx.left(sh.cell_str(ad2), 1)) - 1)
    if not sh.blank(ws.get(3, 2)):
        book.fx.calls.append(Call("DrawListBoxContent", (sh.val_cell(ws.get(3, 2)),)))
    if not sh.blank(ad2):
        n = vbx.clng(vbx.val(vbx.left(sh.cell_str(ad2), 1)))
        while n > 0:
            f.kenh[n] = extract_nth_text(sh.cell_str(ws.get(1, 30)), n)
            n -= 1


def five_op_click(book: Book) -> Effects:
    """UFMitsubishi.FiveOP_Click (UFMit:271-274)。"""
    fx = book.begin()
    book.mit.set_option("5N")
    _set_choice(book, "5N")
    forn_select(book)
    return fx


def ippan_click(book: Book) -> Effects:
    """UFMitsubishi.一般_Click (UFMit:2178-2181)。"""
    fx = book.begin()
    book.mit.set_option("一般")
    _set_choice(book, "一般")
    forn_select(book)
    return fx


def cmb_click(book: Book) -> Effects:
    """UFMitsubishi.CMB_Click (UFMit:517-581)。ForN_Select の簡易版(FiveOP/一般・M5/O5・検入高さは触らない)。"""
    fx = book.begin()
    f = book.mit
    if f.tag == "off" or ZEN not in book.sheets:
        return fx
    ws = book.ws(ZEN)
    book.activate_sheet(ZEN)
    if f.choice == "三菱":
        f.work, f.inflag = "ON", "flagon"
        f.normal_visible, f.mitsubishi_visible, f.fini_visible = False, True, False
        book.write(ws, 1, 22, "三菱")
    elif f.choice == "5N":
        f.work, f.inflag = "OFF", ""
        f.normal_visible, f.mitsubishi_visible, f.fini_visible = False, False, True
        book.write(ws, 1, 22, "5N")
    elif f.choice == "一般":
        f.work, f.inflag = "OFF", ""
        f.normal_visible, f.mitsubishi_visible, f.fini_visible = True, False, False
        book.write(ws, 1, 22, "一般")
    return fx


def cycle_choice(book: Book) -> Effects:
    """UFMitsubishi.CommandButton14_DblClick (UFMit:1298-1361)。一般 → 三菱 → 5N → 一般。"""
    fx = book.begin()
    f = book.mit
    if ZEN not in book.sheets:
        return fx
    book.activate_sheet(ZEN)
    nxt = {"一般": "三菱", "三菱": "5N", "5N": "一般"}.get(f.choice)
    if nxt is None:
        return fx
    _set_choice(book, nxt)
    f.work, f.inflag = ("ON", "flagon") if nxt == "三菱" else ("OFF", "")
    f.normal_visible, f.mitsubishi_visible, f.fini_visible = nxt == "一般", nxt == "三菱", nxt == "5N"
    return fx


# ------------------------------------------------------------------
# 検入高さ
# ------------------------------------------------------------------
def kenhbox_after_update(book: Book, n: int) -> Effects:
    """UFMitsubishi.KenHBox1〜7_AfterUpdate (UFMit:331-388)。1〜6 は "0.0" に整形、7 は AD3 に「{n}枚」。

    B31: KenHBox7 の処理で整形・ログの対象が KenHBox6("0" で整形)になっている。
    """
    fx = book.fx
    f = book.mit
    if n == 7:
        tgt = 6 if quirks.reproduce("B31") else 7
        f.kenh[tgt] = vbx.fmt(f.kenh[tgt], "0")
        fx.calls.append(Call("LogOut", (f"KenHBox{tgt}", "", "三菱")))
    else:
        f.kenh[n] = vbx.fmt(f.kenh[n], "0.0")
        fx.calls.append(Call("LogOut", (f"KenHBox{n}", "", "三菱")))
    if ZEN not in book.sheets:
        return fx
    ws = book.ws(ZEN)
    if sh.blank(ws.get(3, 2)):
        return fx
    if n == 7:
        book.write(ws, 3, 30, f.kenh[7] + "枚")
    else:
        fx.calls.append(Call("DrawListBoxContent", (sh.val_cell(ws.get(3, 2)),)))
    return fx


def mlp_click(book: Book) -> Effects:
    """UFMitsubishi.MLP_Click (UFMit:1833-1892)。重量のゲージ要否と、検入高さ(AD1=計算式・AD2=山数・AD3)の保存。"""
    fx = book.begin()
    f = book.mit
    name = {"P1": ZEN, "P_H494": H494}.get(f.work_flag)
    if name is None or name not in book.sheets:
        return fx
    book.activate_sheet(name)
    ws = book.ws(name)
    m = sh.weight_message(sh.val_cell(ws.get(3, 2)), sh.val_cell(ws.get(3, 4)), sh.val_cell(ws.get(3, 6)))
    if m is None:
        return fx
    fx.messages.append(m)
    if sh.blank(ws.get(3, 2)):
        return fx
    fx.calls.append(Call("DrawListBoxContent", (sh.val_cell(ws.get(3, 2)),)))
    yamasu = ""
    for i in (6, 5, 4, 3, 2, 1):
        if vbx.is_numeric(f.kenh[i]):
            yamasu = f"{i}山"
            break
    if f.kenh[7] == "" and all(f.kenh[i] == "" for i in range(1, 7)):
        for rr in (1, 2, 3):
            book.resume(lambda rr=rr: book.write(ws, rr, 30, ""))
    elif f.kenh[7] == "":
        book.resume(lambda: book.write(ws, 3, 30, ""))
    book.resume(lambda: book.write(ws, 1, 30, f.takasa_formula))
    book.resume(lambda: book.write(ws, 2, 30, yamasu))
    return fx


def weight_click(book: Book) -> Effects:
    """Frame2/4/6・シート発行・UserForm_Click の重量メッセージ(UFMit:276-308 ほか)。UserForm_Click は EnableEvents も戻す。"""
    fx = book.begin()
    name = {"P1": ZEN, "P_H494": H494}.get(book.mit.work_flag)
    if name is None:
        return fx
    if name in book.sheets:
        book.activate_sheet(name)
    ws = book.ws(name)
    m = sh.weight_message(sh.val_cell(ws.get(3, 2)), sh.val_cell(ws.get(3, 4)), sh.val_cell(ws.get(3, 6)))
    if m:
        fx.messages.append(m)
    return fx


def userform_click(book: Book) -> Effects:
    """UFMitsubishi.UserForm_Click (UFMit:2270-2292)。EnableEvents = True と重量メッセージ。"""
    if book.mit.tag == "off":
        return book.begin()
    book.events_enabled = True
    return weight_click(book)


# ------------------------------------------------------------------
# ロット読込(全ヵ所 / H494 共用)
# ------------------------------------------------------------------
def combo_change(book: Book, *, diagonal: Optional[Callable[[float, float], float]] = None,
                 re_diagonal: Optional[Callable[[float, float], float]] = None,
                 lateralbow_j: Optional[Callable[[float, float, float], float]] = None,
                 re_lateralbow: Optional[Callable[[float, float, float], float]] = None,
                 null_raises: bool = False, lot_info_found: bool = True) -> Effects:
    """UFMitsubishi.ComboBox1_Change (UFMit:596-875)。SIKALOTNOW `仕掛` から寸法・公差・対角公差 S1・ラテラル公差 M1。

    - 公差: オーダー公差の±どちらかが 0 なら製造公差(`cask.tolerance_pair`)
    - 対角 S1: 品質ｸﾞﾚｰﾄﾞ_直角度 "4" → Diagonal、"3" → ReDiagonal(Judg の公差表。別担当なので関数を渡す)。
      指示_直角度があれば最優先(R1 = "追記")。**それ以外の等級では S1 を消さない**(前ロットの値が残る)
    - ラテラル M1: 指示_ﾗﾃﾗﾙﾎﾞｰ 優先、"3" → ReLateralbow、それ以外 → Lateralbow_J、0 なら ""
    - その後の Lot情報履歴・品証DB(DistArray / FormReflectDist)・包装仕様書は別担当: `Effects.calls` に積む。
      `lot_info_found=False`(Lot情報履歴に無い)なら VBA と同じく EnableEvents が False のまま終わる
    """
    fx = book.fx
    f = book.mit
    if f.combo_tag == "Load":
        return fx
    book.events_enabled = False
    if ZEN not in book.sheets:
        return fx
    if len(f.combo) != 7:
        return fx
    name = {"P1": ZEN, "P_H494": H494}.get(f.work_flag)
    if name is None:
        raise flow.Stop("実行時エラー '91':" + "\r\n\r\n" + "オブジェクト変数または With ブロック変数が設定されていません。",
                        source="UFMit:617")
    zs = book.ws(name)
    zs.set(1, 15, None)
    zs.font_color[(2, 4)] = (0, 0, 0)
    no_fl = False
    try:
        db = datasource.resolve("SIKALOTNOW")
        if not db.is_file():
            fx.msg("DBファイルが見つかりません:" + "\r\n" + paths.db_file("SIKALOTNOW"), flow.ERROR, "UFMit:625")
            raise VbaError(0, "")
        row = ck.read_lot_row(f.combo)
        if row is not None:
            no_fl = True
            fv = lambda n: ck._field(row, n, width_insensitive=False)  # noqa: E731
            past: list[Any] = ["'" + ("" if fv(MIT_FIELDS[0]) is None else str(fv(MIT_FIELDS[0])))]
            for n in MIT_FIELDS[1:]:
                past.append(fv(n))
            nn = lambda n: ck._nn(fv(n), null_raises)  # noqa: E731
            val = lambda n: vbx.val(nn(n))  # noqa: E731
            book.write(zs, 2, 4, f.combo)
            f.ken_no = nn("検査番号")
            f.h494_ken_no = nn("検査番号")
            f.youto_no = nn("用途ｺｰﾄﾞ")
            f.h494_youto = nn("用途ｺｰﾄﾞ")
            if 1.0 <= val("製造板幅") <= 299.0:
                f.h494 = True
                zs.font_color[(2, 4)] = (255, 0, 0)
            book.write(zs, 3, 2, vbx.fmt(val("製造板厚"), "0.000"))
            book.write(zs, 3, 4, vbx.fmt(val("製造板幅"), "0.0"))
            book.write(zs, 3, 6, vbx.fmt(val("製造板丈"), "0.0"))
            for row_no, key, pat in ((2, "板厚", "0.000"), (3, "板幅", "0.0"), (4, "板丈", "0.0")):
                op, om = val(key + "公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ"), val(key + "公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ")
                mp, mm = val(key + "公差_製造_ﾌﾟﾗｽ"), val(key + "公差_製造_ﾏｲﾅｽ")
                op, om = ck.tolerance_pair(op, om, mp, mm)
                book.write(zs, row_no, 15, vbx.fmt(op, pat))
                book.write(zs, row_no, 13, vbx.fmt(om, pat))
                book.write_formula(zs, row_no, 19, sh.TOLERANCE_FORMULAS[(row_no, 19)])
                book.write_formula(zs, row_no, 17, sh.TOLERANCE_FORMULAS[(row_no, 17)])
            book.write(zs, 1, 12, fv("品質ｸﾞﾚｰﾄﾞ_ﾗﾃﾗﾙﾎﾞｰ"))
            book.write(zs, 1, 13, "")
            book.write(zs, 1, 18, fv("品質ｸﾞﾚｰﾄﾞ_直角度"))
            haba_t = _csng(past[3], null_raises)
            take_t = _csng(past[4], null_raises)
            grade = zs.get(1, 18)
            if _case_eq(grade, "4"):
                book.write(zs, 1, 19, vbx.fmt(_need(diagonal, "diagonal")(haba_t, take_t), "0.0"))
            elif _case_eq(grade, "3"):
                book.write(zs, 1, 19, vbx.fmt(_need(re_diagonal, "re_diagonal")(haba_t, take_t), "0.0"))
            shiji = nn("指示_直角度")
            if shiji != "":
                kou = _csng(vbx.fmt(convert_to_half(shiji), "0.0"), False)
                book.write(zs, 1, 19, kou)
                book.write(zs, 1, 18, "追記")
            late_flag = False
            shiji_late = nn("指示_ﾗﾃﾗﾙﾎﾞｰ")
            if shiji_late != "":
                book.write(zs, 1, 13, convert_to_half(shiji_late))
                late_flag = True
            elif sh.blank(zs.get(1, 13)):
                g = zs.get(1, 12)
                args = (_csng(zs.get(3, 4), False), _csng(zs.get(3, 6), False), _csng(zs.get(3, 2), False))
                if _case_eq(g, "3"):
                    v = _need(re_lateralbow, "re_lateralbow")(*args)
                else:
                    v = _need(lateralbow_j, "lateralbow_j")(*args)
                book.write(zs, 1, 13, vbx.fmt(v, "0.0"))
                if vbx.compare_typed(zs.get(1, 13), 0, vbx.VAR, vbx.NUM) == 0:
                    book.write(zs, 1, 13, "")
            if late_flag:
                book.write(zs, 1, 12, "追記")
        if not no_fl:
            for r, c in ((1, 12), (1, 13), (1, 14), (1, 15), (1, 18), (1, 19), (1, 15), (3, 2), (3, 4), (3, 6),
                         (2, 13), (2, 15), (2, 17), (2, 19), (3, 13), (3, 15), (3, 17), (3, 19), (4, 13), (4, 15),
                         (4, 17), (4, 19), (5, 13), (5, 15)):
                book.write(zs, r, c, "")
            f.ken_no = ""
            if f.combo_tag != "Load":
                fx.msg("Text.データなし　手入力よろしく", source="UFMit:794")
            act = book.active
            for (r, c), formula in sh.TOLERANCE_FORMULAS.items():
                act.set_formula(r, c, formula)
            return fx                                   # EnableEvents は False のまま(VBA どおり)
        fx.calls.append(Call("HistorySave", (past,)))
        if not lot_info_found:
            return fx
        fx.calls.append(Call("GetLotInfo/DbReturnItem", ("Lot情報履歴", f.lot_box)))
        if f.lot_box != "":
            fx.calls.append(Call("DistArray/FormReflectDist", (f.lot_box, f.ken_no)))
        book.activate_sheet(name)
        book.events_enabled = True
        lot_no = book.cask.lot_box if book.cask.visible else ""
        lot_no = {"P1": f.lot_box, "P_H494": f.h494_lot}.get(f.work_flag, lot_no)
        if f.pack_no_on_off != "off":
            fx.calls.append(Call("PackingNoCheck/UFPack_Open", (lot_no,)))
        return fx
    except VbaError as e:
        book.events_enabled = True
        fx.msg("DB読み込みエラー:" + "\r\n" + e.description, flow.ERROR, "UFMit:874")
        return fx


def _need(fn: Optional[Callable], name: str) -> Callable:
    if fn is None:
        raise ValueError(f"{name}(Judg の公差表の関数)を渡してください")
    return fn


def _case_eq(v: Any, lit: str) -> bool:
    """`Select Case セル: Case "4"`(Variant と文字列の定数 → 文字で比べる。数のセルは CStr して比べる)。"""
    return vbx.compare_typed(v, lit, vbx.VAR, vbx.STR) == 0


def _csng(v: Any, null_raises: bool) -> float:
    """`Dim x As Single: x = v`。Null は 94(null_raises)または ""、数にならない文字列は 13。"""
    if v is None:
        if null_raises:
            raise VbaError(94, ck.INVALID_NULL)
        v = ""
    if isinstance(v, (int, float)):
        return vbx.single(float(v))
    if not vbx.is_numeric(v):
        raise sh.type_mismatch("CSng")
    return vbx.single(vbx.cdbl(v))


def lot_box_change(book: Book, **kw: Any) -> Effects:
    """UFMitsubishi.LotBox_Change (UFMit:1822-1831)。7 桁で ComboBox1 に写す。"""
    fx = book.fx
    book.events_enabled = False
    if len(book.mit.lot_box) != 7:
        return fx
    book.mit.combo = book.mit.lot_box
    combo_change(book, **kw)
    book.events_enabled = True
    return fx


# ------------------------------------------------------------------
# MAX / MIN
# ------------------------------------------------------------------
def _wvals(ws: sh.Sheet, n: int, k: int, col: int) -> list[float]:
    lo, hi = min(n, k), max(n, k)
    return [float(v) for v in (ws.get(r, col) for r in range(lo, hi + 1))
            if isinstance(v, (int, float)) and not isinstance(v, bool)]


def _wmax(ws: sh.Sheet, n: int, k: int, col: int) -> float:
    v = _wvals(ws, n, k, col)
    return max(v) if v else 0.0


def _wmin(ws: sh.Sheet, n: int, k: int, col: int) -> float:
    v = _wvals(ws, n, k, col)
    return min(v) if v else 0.0


def max_min(book: Book, temp: str = ZEN) -> None:
    """Mitsubishi.全ヵ所用最大最少 (Mit:134-322)。MIN〜MAX 行の最大最小(Single)を書き、一致セルに下線。

    厚の MIN が 0(空の列がある)とき、列ごとの最小を並べて 2 番目(だめなら 3 番目)に小さいものにする。
    B28: 列ごとの最小を入れる Select Case が `Case 1, 2, 6, 8` で 3・4 番目(F・H 列)に入らない。
    """
    f = book.mit
    if f.work_flag == "P1":
        if f.min_text == "" or f.max_text == "":
            return
        n, k = vbx.clng(vbx.val(f.min_text) + 6), vbx.clng(vbx.val(f.max_text) + 6)
        erow, ecol = 36, 13
    elif f.work_flag == "P_H494":
        if f.h494_min_text == "" or f.h494_max_text == "":
            return
        n, k = vbx.clng(vbx.val(f.h494_min_text) + 6), vbx.clng(vbx.val(f.h494_max_text) + 6)
        erow, ecol = 16, 15
    else:
        n = k = erow = ecol = 0
    ws = book.ws(temp)
    mx = lambda c: _wmax(ws, n, k, c)  # noqa: E731
    mn = lambda c: _wmin(ws, n, k, c)  # noqa: E731
    max_atu = vbx.single(max(mx(2), mx(4), mx(6), mx(8)))
    min_atu = vbx.single(min(mn(2), mn(4), mn(6), mn(8)))
    if min_atu == 0:
        atu = [0.0, 0.0, 0.0, 0.0]
        idx = (1, 2, 6, 8) if quirks.reproduce("B28") else (1, 2, 3, 4)
        cols = {1: 2, 2: 4, 3: 6, 4: 8, 6: 6, 8: 8}
        slot = {1: 0, 2: 1, 3: 2, 4: 3, 6: 2, 8: 3}
        for i in range(1, 5):
            if i in idx:
                m = mn(cols[i])
                atu[slot[i]] = 0.0 if m == 0 else vbx.single(m)
        srt = sorted(atu)
        min_atu = vbx.single(srt[1])
        if min_atu == 0:
            min_atu = vbx.single(srt[2])
    if f.work_flag == "P1":
        max_haba = vbx.single(max(mx(11), mx(13)))
        min_haba = vbx.single(min(mn(11), mn(13)))
        max_take = vbx.single(max(mx(15), mx(17)))
        min_take = vbx.single(min(mn(15), mn(17)))
        ta, th, tt = 23, 25, 27
    elif f.work_flag == "P_H494":
        max_haba = vbx.single(max(mx(11), mx(13), mx(15)))
        min_haba = vbx.single(min(mn(11), mn(13), mn(15)))
        max_take = vbx.single(max(mx(17), mx(19)))
        min_take = vbx.single(min(mn(17), mn(19)))
        ta, th, tt = 22, 23, 24
    else:
        max_haba = min_haba = max_take = min_take = 0.0
        ta = th = tt = 0
    for c in (ta, th, tt):
        if c:
            book.write(ws, 3, c, "")
            book.write(ws, 4, c, "")
    ws.clear_underline(7, 2, erow, 22)

    def mark(cols: range, target: float, out: tuple[int, int]) -> None:
        for i in range(n, k + 1):
            for a in cols:
                if sh.eq_single(ws.get(i, a), target):
                    ws.underline.add((i, a))
                    book.write(ws, out[0], out[1], target)
                    break

    mark(range(2, 9, 2), max_atu, (3, ta))
    mark(range(2, 9, 2), min_atu, (4, ta))
    mark(range(11, ecol + 1, 2), max_haba, (3, th))
    mark(range(11, ecol + 1, 2), min_haba, (4, th))
    mark(range(ecol + 2, ecol + 5, 2), max_take, (3, tt))
    mark(range(ecol + 2, ecol + 5, 2), min_take, (4, tt))


# ------------------------------------------------------------------
# 測定機器
# ------------------------------------------------------------------
def kiki_tensou(book: Book, ws: sh.Sheet, x: int, y: int, atu: str, haba: str, yout: str = "") -> None:
    """Mitsubishi.全ヵ所用機器転送 (Mit:986-1056)。板厚・ライン別の測定機器名(全角空白 2 つで区切る)。"""
    line, conb = LINE_DEVICE.get(book.mit.line_name, ("", ""))
    v = vbx.val(atu)
    sp = "\u3000\u3000"
    text: Optional[str] = None
    if 0 < v <= 28.0:
        text = "ﾏｲｸﾛ:" + line + "M1" + sp + "ｺﾝﾍﾞｯｸｽ:" + line + conb
    elif 28 < v <= 54.0:
        text = "ﾏｲｸﾛ:" + line + "M2" + sp + "ｺﾝﾍﾞｯｸｽ:" + line + conb
    elif 54.0 < v <= 74.0:
        text = "ﾏｲｸﾛ:" + line + "M3" + sp + "ｺﾝﾍﾞｯｸｽ:" + line + conb
    elif 74.0 < v < 300:
        text = "ﾉｷﾞｽ:" + line + "N5" + sp + "ｺﾝﾍﾞｯｸｽ:" + line + conb
    elif 300.0 <= v:
        text = "ﾉｷﾞｽ:" + line + "N6" + sp + "ｺﾝﾍﾞｯｸｽ:" + line + conb
    if text is not None:
        book.write(ws, x, y, text)
    if 300.0 <= v and v == 300.0:
        lo = hi = False
        for i in range(14, 19):
            for a in range(2, 9, 2):
                cv = ws.get(i, a)
                if sh.is_numeric(cv) and not sh.blank(cv):
                    if sh.val_cell(cv) < 300.0:
                        lo = True
                    elif sh.val_cell(cv) >= 300.0:
                        hi = True
        if lo and hi:
            book.write(ws, x, y, "ﾉｷﾞｽ：" + line + "N5,N6" + sp + "ｺﾝﾍﾞｯｸｽ:" + line + "C11")
        elif lo and not hi:
            book.write(ws, x, y, "ﾉｷﾞｽ" + line + "N5" + sp + "ｺﾝﾍﾞｯｸｽ:" + line + "C11")
        elif hi and not lo:
            book.write(ws, x, y, "ﾉｷﾞｽ" + line + "N6" + sp + "ｺﾝﾍﾞｯｸｽ:" + line + "C11")
    if yout == "H494":
        if 1.0 <= vbx.val(haba) <= 299.0:
            book.write(ws, x, y, "ﾏｲｸﾛ:" + line + "M1" + sp + "幅:ﾉｷﾞｽ:" + "N5" + "_丈:ｺﾝﾍﾞｯｸｽ:" + line + conb)


# ------------------------------------------------------------------
# 転送(ｺﾋﾟ-用 / 転送済)
# ------------------------------------------------------------------
COPY_KEEP = (1, 2, 4, 6, 8, 11, 13, 15, 17, 22, 23, 25, 27, 29, 30)   # 削除後に残る元の列(A,B,D,F,H,K,M,O,Q,V,W,Y,AA,AC,AD)


def _copy_sheet(src: sh.Sheet) -> sh.Sheet:
    """`ｺﾋﾟ-用`: 全ヵ所検査 A6:AD{n} を D1 へ貼り、F,H,J,L,M,O,Q,S,U:X,AA,AC,AE を削除 (Mit:655-659)。

    結果 D=No, E〜H=厚1〜4, I,J=幅, K,L=丈, **M=対角の判定(V)**, N=ﾗﾃﾗﾙﾎﾞｰ, O=幅歪, P=丈歪, Q=ｹﾞｰｼﾞ, R=外観。
    """
    q = sh.Sheet("ｺﾋﾟ-用")
    n = src.last_row(2)
    for r in range(6, max(n, 6) + 1):
        for k, c in enumerate(COPY_KEEP):
            q.set(r - 5, 4 + k, src.get(r, c))
    return q


def _defect_sheet_name(book: Book) -> str:
    return "不良要因印刷" + "_" + book.mit.line_name


def zen_transfer(book: Book) -> None:
    """Mitsubishi.全ヵ所用転送 (Mit:619-859)。ｺﾋﾟ-用・測定機器(AC3)・不良要因シート、三菱なら 転送済 と歪/外観チェック。

    三菱の歪/外観の入力が無いと、ｺﾋﾟ-用/転送済/全ヵ所検査保存 を消して `End`(→ `flow.Stop`)。
    """
    fx = book.fx
    f = book.mit
    zen = book.ws(ZEN)
    atu, haba = sh.cell_str(zen.get(3, 2)), sh.cell_str(zen.get(3, 4))
    flag2 = book.has("転送済")
    book.delete_sheet("ｺﾋﾟ-用")
    if f.kname == "" and f.ken_no == "":
        fx.msg("検査員名と検査番号を入力してください", source="Mit:644")
        return
    cp = _copy_sheet(zen)
    n = zen.last_row(2)
    book.add_sheet(cp)
    for i in range(2, 9):                      # 数字変換(Mit:661-672)。列削除後の S 列は空なので実際は何もしない
        for num, mark in ((1, "①"), (2, "②"), (3, "③"), (4, "④")):
            try:
                hit = vbx.compare_typed(cp.get(i, 19), num, vbx.VAR, vbx.NUM) == 0
            except vbx.TypeMismatch as e:
                raise sh.type_mismatch(str(e)) from e
            if hit:
                cp.set(i, 19, mark)
                break
    kiki_tensou(book, zen, 3, 29, atu, haba, sh.cell_str(zen.get(4, 8)))
    zen.underline = set()
    put = lambda r, c, v: cp.set(r, c, sh.to_cell(v))  # noqa: E731
    put(1, 1, "LotNo")
    cp.set(1, 2, zen.get(2, 4))
    put(2, 1, "Size")
    put(2, 2, sh.cell_str(zen.get(3, 2)) + "×" + sh.cell_str(zen.get(3, 4)) + "×" + sh.cell_str(zen.get(3, 6)))
    put(3, 1, "測定機器")
    cp.set(3, 2, zen.get(3, 29))
    put(5, 1, "検査員")
    put(5, 2, f.kname)
    put(6, 1, "作業日")
    put(6, 2, sh.fmt_date(zen.get(1, 2), "yyyy/mm/dd  h:mm") + "～" + sh.fmt_date(book.now(), "yyyy/mm/dd h:mm"))
    dname = _defect_sheet_name(book)
    if not book.has(dname):
        book.add_sheet(sh.Sheet(dname))
        fx.calls.append(Call("AddWS/Pesute", ("不良要因印刷原紙", dname)))
    if dname in book.extra:
        book.extra[dname].set(1, 2, sh.to_cell(f.lot_box))
    fx.calls.append(Call("不良要因印刷_B1", (dname, f.lot_box)))
    if book.has(dname):
        fx.calls.append(Call("不良要因印刷範囲", (f.kname,)))
    if f.inflag != "flagon":
        return
    if flag2:
        book.delete_sheet("転送済")
    if not sh.blank(zen.get(2, 4)):
        tz = sh.template(sh.MITSU_TEMPLATE)
        tz.name = "転送済"
    else:
        tz = sh.Sheet("転送済")
    book.add_sheet(tz)
    tput = lambda r, c, v: tz.set(r, c, sh.to_cell(v))  # noqa: E731
    tput(2, 4, f.ken_no)
    tz.set(5, 5, zen.get(3, 2))
    tz.set(5, 7, zen.get(3, 4))
    tz.set(5, 9, zen.get(3, 6))
    for rr, col in ((2, 6), (3, 10), (4, 12)):
        tput(8, col, "' " + "-" + sh.cell_str(zen.get(rr, 13)) + "mm " + "\n" + "+" + sh.cell_str(zen.get(rr, 15)) + "mm")
    tz.set(19, 17, cp.get(6, 2))
    tput(20, 17, f.kname)
    kiki_tensou(book, tz, 2, 7, atu, haba, sh.cell_str(zen.get(4, 8)))
    for i in range(10, 17, 2):
        tz.set(i, 2, cp.get(i - 8, 17))
    for i in range(10, 17, 2):
        tz.set(i, 3, cp.get(i - 8, 19))
    a = 0
    for i in range(2, 9, 2):
        a += 1
        if not sh.blank(cp.get(i, 5)):
            tz.set(i + 8, 4, float(a))
    for i in range(8):
        for j in range(8):
            tz.set(10 + i, 10 + j, cp.get(2 + i, 9 + j))
        for j in range(4):
            tz.set(10 + i, 6 + j, cp.get(2 + i, 5 + j))
    if f.work != "ON":
        return
    for i in range(7, n + 1):
        a = 24
        for a in (24, 25):
            if sh.blank(zen.get(i, a)):
                _mitsubishi_abort(book, "歪入力がありません,はじめからやり直してください" + " " + " *[座標](" + str(i) + "," + str(a) + ")")
        a = 26
        if sh.blank(zen.get(i, a)):
            break
    for i in range(7, n + 1):
        if sh.blank(zen.get(i, 26)):
            _mitsubishi_abort(book, "外観入力がありません,はじめからやり直してください" + " " + " *[座標](" + str(i) + "," + "26" + ")")


def _mitsubishi_abort(book: Book, text: str) -> None:
    book.fx.msg(text, flow.WARN, "Mit:829")
    for name in ("ｺﾋﾟ-用", "転送済", "全ヵ所検査保存"):
        book.delete_sheet(name)
    raise flow.Stop(text, source="Mit:837", messages=list(book.fx.messages))


# ------------------------------------------------------------------
# 印刷の設定
# ------------------------------------------------------------------
def _kname_for(book: Book) -> str:
    f = book.mit
    if f.work_flag == "P_H494":
        if not quirks.reproduce("B30") and f.h494_kname == "":
            return f.kname
        return f.h494_kname
    return f.kname if f.work_flag == "P1" else ""


def zen_print_setup(book: Book, temp: str = ZEN) -> Optional[PrintJob]:
    """Mitsubishi.全ヵ所用印刷範囲 (Mit:860-938)。全ヵ所 A6:AD{n} / H494 A6:AB{n}、A4 横 1×1、ヘッダ/フッタ。

    B30: フッタの測定機器は AC3(全ヵ所の書き込み先)を見る。H494 は Y2 に書くので空欄になる。
    """
    if not book.has(temp) or temp not in book.sheets:
        return None
    ws = book.ws(temp)
    n = ws.last_row(2)
    f = book.mit
    area = ""
    if f.work_flag == "P1":
        area = "$A$6:$AD$" + str(n)
    elif f.work_flag == "P_H494":
        area = "$A$6:$AB$" + str(n)
    kname = _kname_for(book)
    book.activate_sheet(temp)
    ws.fill(7, 2, 76, 28, sh.WHITE)
    v = lambda r, c: sh.val_cell(ws.get(r, c))  # noqa: E731
    x, y, z = v(3, 2), v(3, 4), v(3, 6)
    am, ap, hm, hp, tm, tp = v(2, 13), v(2, 15), v(3, 13), v(3, 15), v(4, 13), v(4, 15)
    device = ws.get(3, 29)
    if temp == H494 and not quirks.reproduce("B30"):
        device = ws.get(2, 25)
    return PrintJob(
        sheet=temp, print_area=area,
        right_footer='&"Meiryo UI"&8' + "*" + sh.fmt_date(ws.get(1, 2), "yyyy/mm/dd  h:mm") + "～"
        + sh.fmt_date(book.now(), "yyyy/mm/dd  h:mm") + " " + "印刷",
        right_header='&"Meiryo UI"&8' + "*" + "厚公差：" + vbx.fmt(x - am, "0.000") + "～" + vbx.fmt(x + ap, "0.000") + "\r" + "*"
        + "巾公差：" + vbx.fmt(y - hm, "0.0") + "～" + vbx.fmt(y + hp, "0.0") + "\r" + "*"
        + "丈公差：" + vbx.fmt(z - tm, "0.0") + "～" + vbx.fmt(z + tp, "0.0") + "\r" + "*"
        + "歪公差：" + "幅:" + vbx.fmt(v(5, 13), "0.0") + "_" + "丈:" + vbx.fmt(v(5, 15), "0.0") + "\r" + "*"
        + "対角公差：" + vbx.fmt(v(1, 19), "0.0") + "\r" + "*"
        + "ﾗﾃﾗﾙﾎﾞｰ公差：" + vbx.fmt(v(1, 13), "0.0"),
        left_header='&"Meiryo UI"&8' + "*" + sh.cell_str(ws.get(2, 4)) + "\r" + "*"
        + vbx.fmt(x, "0.000") + "×" + vbx.fmt(y, "0.0") + "×" + vbx.fmt(z, "0.0") + "\r" + "*"
        + sh.cell_str(ws.get(1, 30)) + "_" + sh.cell_str(ws.get(2, 30)),
        left_footer='&"Meiryo UI"&8' + "*検査員：" + kname + " " + "*測定機器：" + sh.cell_str(device),
        paper="A4", orientation="landscape", fit_wide=1, fit_tall=1,
        margins={"header": 0.2, "footer": 0.2, "left": 0.5, "right": 0.5, "top": 0.8, "bottom": 0.5},
        center_h=True, center_v=True, snapshot=ws.snapshot())


def genshi_print_setup(book: Book) -> Optional[PrintJob]:
    """Mitsubishi.原紙印刷範囲 (Mit:939-985)。5N シート(A1:L{n})または 転送済(B2:R20)。"""
    f = book.mit
    if f.choice == "5N":
        if not book.has("5N"):
            return None
        ws = book.ws("5N")
        area = "$A$1:$L$" + str(ws.last_row(3) + 3)
    elif f.choice in ("三菱", "一般"):
        if not book.has("転送済"):
            return None
        ws = book.ws("転送済")
        area = "$B$2:$R$20"
    else:
        raise VbaError(91, "オブジェクト変数または With ブロック変数が設定されていません。", "Mit:955")
    zen = book.ws(ZEN)
    return PrintJob(
        sheet=ws.name, print_area=area,
        left_header='&"Meiryo UI"&7' + f.line_name + "\r" + "*" + '&"Meiryo UI"&7' + "*" + sh.cell_str(zen.get(1, 30))
        + "_" + sh.cell_str(zen.get(2, 30)),
        paper="A4", orientation="landscape", fit_wide=1, fit_tall=1,
        margins={"header": 0, "footer": 0, "left": 0, "right": 0, "top": 0.5, "bottom": 0},
        center_h=True, snapshot=ws.snapshot())


# ------------------------------------------------------------------
# 保存
# ------------------------------------------------------------------
def _save_copy(book: Book, temp: str, kname: str, *, widths: str) -> sh.Sheet:
    src = book.ws(temp)
    s = src.copy(temp + "保存")
    s.interior = {}
    put = lambda r, c, v: s.set(r, c, sh.to_cell(v))  # noqa: E731
    put(1, 3, "～")
    put(1, 4, sh.fmt_date(book.now(), "yyyy/mm/dd h:mm"))
    if widths == "P1":
        s.col_width.update({2: 7.7, 11: 7.7, 24: 7.7})
    elif widths == "P_H494":
        s.col_width.update({2: 7.7, 11: 7.7, 25: 7.7})
    elif widths == "5N":
        s.col_width.update({c: 8.75 for c in range(1, 31)})
        s.col_width.update({c: 2.5 for c in (3, 5, 7, 9, 10, 12, 14, 16, 18, 21, 24, 26, 28)})
    put(1, 6, "検査員：" + kname)
    s.font_size[(1, 6)] = 7
    return s


def _snap(book: Book, names: list[str]) -> dict[str, Any]:
    out = {}
    for name in names:
        if name in book.extra or name in book.sheets:
            out[name] = book.ws(name).snapshot()
        elif name in book.other_sheets:
            out[name] = None            # 別担当のシート(不良要因印刷_ など)。中身は持っていない
    return out


def justincase_zn_save(book: Book, temp: str = ZEN) -> Optional[SavePlan]:
    """Mitsubishi.JustinCase_ZnSave (Mit:324-426)。新規発行の前の念のための保存(検査履歴\\JustinCaseSave)。

    含めるのは表示中の `{temp}保存` と `全ヵ所検査`(H494 でも 全ヵ所検査 の名前で探す)。
    """
    if not book.has(temp):
        return None
    f = book.mit
    s = _save_copy(book, temp, f.kname, widths="P1" if f.work_flag == "P1" else "")
    s.visible = True
    book.add_sheet(s)
    names = [n for n in (s.name, ZEN) if (n in book.extra or n in book.sheets) and book.ws(n).visible]
    plan = SavePlan("JustinCase_ZnSave", sheets=_snap(book, names))
    if not names:
        return None
    book.fx.saves.append(plan)
    today = book.now()
    hist = paths.get("history")
    base = paths.join(hist, "Test\\Test_履歴\\JustinCaseSave") if f.line_name == "Test" else paths.join(hist, "JustinCaseSave")
    lot = sh.cell_str(s.get(2, 4))
    day_lot = today.strftime("%Y%m%d") + "_" + lot
    ck._save_numbered(book, plan, paths.join(base, day_lot), day_lot, lot)
    return plan


def _new_file_copy(plan: SavePlan, saved: str, dest_dir: str) -> None:
    plan.steps.append(SaveStep("mkdir", dest_dir))
    fname = saved.split("\\")[-1]
    dest = (dest_dir + "\\" + fname) if dest_dir != "" else ("\\" + fname)
    plan.steps.append(SaveStep("copy", dest, src=saved))


def zen_save(book: Book, temp: str = ZEN) -> Optional[SavePlan]:
    """Mitsubishi.全ヵ所専用保存 (Mit:427-618)。`{temp}保存`・ｺﾋﾟ-用・転送済・5N・不良要因印刷_{Line} を `C:\\検査履歴` に保存し、
    三菱(Inflag="flagon")なら NPS の 三菱受託検査_履歴 にも保存、サーバの `{temp}_履歴` へ写す。

    三菱のときは SetPath が NPS の `月\\日付_ロット`(フォルダではない)に置き換わったまま `NewFileName(SetPath\\)` を呼ぶため、
    写しの所で実行時エラーになる(G5 §13-3)。手順はそこまでを返し、`flow.Stop` で止める。
    丸徳など 5 ライン以外は写し先が "" になり `\\{ファイル名}` へ写す(G5 §13-4)。
    """
    fx = book.fx
    if not book.has(temp):
        fx.msg(temp + "がありません", flow.ERROR, "Mit:429")
        return None
    f = book.mit
    kname = f.kname if f.work_flag == "P1" else (_kname_for(book) if f.work_flag == "P_H494" else "")
    s = _save_copy(book, temp, kname, widths=f.work_flag)
    book.add_sheet(s)
    dname = _defect_sheet_name(book)
    order = [n for n in (s.name, "ｺﾋﾟ-用", "転送済", "5N", dname) if book.has(n)]
    plan = SavePlan("全ヵ所専用保存", sheets=_snap(book, order))
    fx.saves.append(plan)
    today = book.now()
    month = today.strftime("%Y.%m")
    lot = sh.cell_str(s.get(2, 4))
    day_lot = today.strftime("%Y%m%d") + "_" + lot
    set_path = paths.join(paths.get("local_history"), month, day_lot)
    saved = ck._save_numbered(book, plan, set_path, day_lot, lot)
    if f.inflag == "flagon":
        month_path = paths.join(paths.get("nps_mitsubishi"), month)
        if not book.exists(month_path):
            plan.steps.append(SaveStep("mkdir", month_path))
        set_path = paths.join(month_path, day_lot)
        if not book.exists(set_path + ".xlsx"):
            p = set_path + ".xlsx"
        else:
            p = None
            for cnt in range(2, 1000):
                cand = set_path + "." + vbx.right("00" + str(cnt), 3) + ".xlsx"
                if not book.exists(cand):
                    p = cand
                    break
            if p is None:
                fx.msg("連番をこれ以上付けることが出来ません。", flow.ERROR, "Mit:585")
                plan.error = "連番をこれ以上付けることが出来ません。"
                raise flow.Stop(plan.error, source="Mit:586", messages=list(fx.messages))
        plan.steps.append(SaveStep("save", p, password=ck._password(lot, today), readonly=True))
    file_n = temp + "_履歴\\"
    hist = paths.get("history")
    if f.line_name in FIVE_LINES:
        srv = paths.join(hist, file_n + month)
    elif f.line_name == "Test":
        srv = paths.join(hist, "Test\\ " + file_n + month)
    else:
        srv = ""
    if f.inflag == "flagon":
        plan.steps.append(SaveStep("mkdir", srv))
        plan.error = "実行時エラー '76': パスが見つかりません。(NewFileName(" + set_path + "\\))"
        raise flow.Stop(plan.error, source="Mit:613", messages=list(fx.messages))
    _new_file_copy(plan, saved, srv)
    return plan


def fiven_save(book: Book) -> Optional[SavePlan]:
    """Mitsubishi.FiveN専用保存 (Mit:1232-1378)。全ヵ所検査保存 他を保存し、`全ヵ所検査_履歴` へ写す。
    5N(Inflag="" かつ Work="OFF" かつ Choice="5N")なら工場共有 `【梱包】５N\\yyyy.mm` にも写す。

    B33: 新規発行(CB1)で Work="ON" になるので、5N を選び直さないと工場共有へ写らない。直した場合は Work を見ない。
    """
    fx = book.fx
    if ZEN not in book.sheets:
        fx.msg("保存ｼｰﾄがありません", flow.ERROR, "Mit:1234")
        return None
    f = book.mit
    s = _save_copy(book, ZEN, f.kname, widths="5N")
    book.add_sheet(s)
    dname = _defect_sheet_name(book)
    order = [n for n in (s.name, "ｺﾋﾟ-用", "転送済", "5N", dname) if book.has(n)]
    plan = SavePlan("FiveN専用保存", sheets=_snap(book, order))
    fx.saves.append(plan)
    today = book.now()
    month = today.strftime("%Y.%m")
    lot = sh.cell_str(s.get(2, 4))
    day_lot = today.strftime("%Y%m%d") + "_" + lot
    set_path = paths.join(paths.get("local_history"), month, day_lot)
    saved = ck._save_numbered(book, plan, set_path, day_lot, lot)
    hist = paths.get("history")
    if f.line_name in FIVE_LINES:
        srv = paths.join(hist, "全ヵ所検査_履歴\\" + month)
    elif f.line_name == "Test":
        srv = paths.join(hist, "Test\\Test_履歴\\全ヵ所検査_履歴\\" + month)
    else:
        srv = ""
    _new_file_copy(plan, saved, srv)
    if quirks.reproduce("B33"):
        cond = f.inflag == "" and f.work == "OFF" and f.choice == "5N"
    else:
        cond = f.inflag == "" and f.choice == "5N"
    if cond:
        if f.line_name in FIVE_LINES:
            kou = paths.join(paths.get("factory"), "【梱包】５N", month)
        elif f.line_name == "Test":
            kou = paths.join(hist, "Test\\5N_履歴\\" + month)
        else:
            kou = ""
        _new_file_copy(plan, saved, kou)
    return plan


# ------------------------------------------------------------------
# フォームの操作
# ------------------------------------------------------------------
def _red_x_question(book: Book, answers: Optional[dict], key: str) -> Optional[Union[flow.Question, bool]]:
    """印刷前の警告(UFMit:1661-1683)。厚 B7:I76・歪 Y7:AB76 の赤背景の × を数える。いいえなら False。"""
    ws = book.ws(ZEN)
    r1 = check_red_x(ws, 7, 2, 76, 9)
    r2 = check_red_x(ws, 7, 25, 76, 28)
    text = None
    if r1 > 0 and r2 > 0:
        text = "公差外判定の板があります " + "厚:" + str(r1) + "箇所　" + "歪:" + str(r2) + "箇所" + "\r\n" + "処理を続行しますか？"
    elif r1 > 0 and r2 == 0:
        text = "公差外判定の板があります " + "厚:" + str(r1) + "箇所　" + "\r\n" + "処理を続行しますか？"
    elif r1 == 0 and r2 > 0:
        text = "公差外判定の板があります " + "歪:" + str(r2) + "箇所" + "\r\n" + "処理を続行しますか？"
    if text is None:
        return True
    a = (answers or {}).get(key)
    if a is None:
        return flow.Question(key, text, source="UFMit:1678", messages=list(book.fx.messages))
    return bool(a)


def _defect_print(book: Book, kname: str, lot_box: str) -> None:
    """不良要因印刷_{Line} を用意して印刷(UFMit:1720-1730 ほか)。中身は別担当(Print.不良要因印刷範囲)。"""
    fx = book.fx
    dname = _defect_sheet_name(book)
    if not book.has(dname):
        book.add_sheet(sh.Sheet(dname))
        fx.calls.append(Call("AddWS/Pesute", ("不良要因印刷原紙", dname)))
    else:
        b1 = book.extra[dname].get(1, 2) if dname in book.extra else None
        if vbx.compare_typed(b1, lot_box, vbx.VAR, vbx.STR) != 0:
            book.add_sheet(sh.Sheet(dname))
            fx.calls.append(Call("AddWS/Pesute", ("不良要因印刷原紙", dname)))
    fx.calls.append(Call("不良要因印刷範囲", (kname,)))
    fx.prints.append(PrintJob(sheet=dname))


def new_sheet(book: Book, answers: Optional[dict] = None) -> Union[Effects, flow.Question]:
    """UFMitsubishi.CommandButton1_Click (UFMit:1250-1287) 全ヵ所の新規発行。最後に Work = "ON"(B33 の元)。

    不良要因印刷_{Line} の削除は DisplayAlerts を止めていないので Excel の確認が出る(ここでは消したことにする)。
    """
    fx = book.begin()
    f = book.mit
    if f.tag == "off":
        return fx
    a = (answers or {}).get("zen_new")
    if a is None:
        return flow.Question("zen_new", "新規発行しますか？現在のｼｰﾄは削除されます。", source="UFMit:1254")
    fx.calls.append(Call("Register_Check/Daily_Create"))
    if a:
        justincase_zn_save(book)
        zen_copy(book)
        fx.logs.append(LogCall("*UFKen:" + "三菱受託検査ｼｰﾄ発行", "", sh.log_dir(f.line_name)))
        dname = _defect_sheet_name(book)
        if book.has(dname):
            book.delete_sheet(dname)
        f.work = "ON"
        if f.lot_box != "":
            f.lot_box = ""
            lot_box_change(book)
        f.ken_no = ""
        if f.combo != "":
            f.combo = ""
            combo_change(book)
        f.youto_no = ""
    else:
        fx.msg("ｷｬﾝｾﾙ処理", source="UFMit:1285")
    return fx


def zen_copy(book: Book, temp: str = ZEN) -> None:
    """Mitsubishi.全ヵ所用コピー (Mit:95-133)。原紙から作業シート(全ヵ所検査 / H494検査)を作り直す。"""
    tname = sh.ZEN_TEMPLATE if temp == ZEN else sh.H494_TEMPLATE
    ws = sh.template(tname)
    ws.name = temp
    ws.visible = True
    ws.tab_color = 16764108
    book.sheets[temp] = ws
    book.activate_sheet(temp)


def transfer(book: Book) -> Effects:
    """UFMitsubishi.CommandButton3_Click (UFMit:1605-1630) 転送。検査法の選択 → 最大最少 → Ainput_Startup("Mitsubishi", 検査法)。"""
    fx = book.begin()
    f = book.mit
    if f.tag == "off":
        return fx
    if f.inspection_method == "":
        fx.msg("検査法を選択してください", flow.ERROR, "UFMit:1617")
        return fx
    book.events_enabled = False
    try:
        max_min(book)
    except VbaError as e:
        raise flow.Stop(f"実行時エラー '{e.number}':" + "\r\n\r\n" + e.description, source="Mit:全ヵ所用最大最少")
    fx.calls.append(Call("Ainput_Startup", ("Mitsubishi", f.inspection_method)))
    book.events_enabled = True
    return fx


def transfer_and_save(book: Book) -> Effects:
    """UFMitsubishi.CommandButton4_Click (UFMit:1632-1641)。"""
    fx = book.begin()
    zen_transfer(book)
    if book.mit.choice != "5N":
        zen_save(book)
    else:
        fiven_save(book)
    return fx


def save_only(book: Book) -> Effects:
    """UFMitsubishi.CommandButton5_Click (UFMit:1643-1645)。"""
    fx = book.begin()
    zen_save(book)
    return fx


def justincase(book: Book) -> Effects:
    """UFMitsubishi.CommandButton6_Click (UFMit:1647-1652)。"""
    fx = book.begin()
    justincase_zn_save(book)
    return fx


def _height_missing(book: Book) -> bool:
    ws = book.ws(ZEN)
    return sh.blank(ws.get(2, 30)) and sh.blank(ws.get(3, 30))


def ippan_print(book: Book, answers: Optional[dict] = None) -> Union[Effects, flow.Question]:
    """UFMitsubishi.CommandButton7_Click (UFMit:1654-1771) 一般の印刷・保存。"""
    fx = book.begin()
    f = book.mit
    if ZEN not in book.sheets:
        return fx
    book.activate_sheet(ZEN)
    if _height_missing(book):
        fx.msg("検入高さを入力してください。", flow.ERROR, "UFMit:1671")
        return fx
    q = _red_x_question(book, answers, "zen_red_x")
    if isinstance(q, flow.Question):
        return q
    if q is False:
        return fx
    f.tag = "off"
    f.work, f.inflag = "OFF", ""
    if f.kname == "" and f.ken_no == "":
        fx.msg("検査員名と検査番号を入力してください", source="UFMit:1690")
        return fx
    fx.sounds.append("おまちください.wav")
    zen_transfer(book)
    job = zen_print_setup(book)
    fx.sounds.append("印刷.wav")
    a = (answers or {}).get("zen_print")
    if a is None:
        return flow.Question("zen_print", "印刷しますか？", source="UFMit:1712", messages=list(fx.messages))
    if a:
        book.ws(ZEN).visible = True
        if job:
            fx.prints.append(job)
        zen_save(book)
        _defect_print(book, f.kname, f.lot_box)
        for name in ("ｺﾋﾟ-用", "全ヵ所検査保存", "全ヵ所作業中保存"):
            book.delete_sheet(name)
        row = sh.history_row(book, book.ws(ZEN), "全ヵ所")
        if row is not None:
            fx.history.append(row)
        fx.sounds.append("完了.wav")
    else:
        fx.msg("ｷｬﾝｾﾙ処理", source="UFMit:1764")
    f.tag = ""
    return fx


def fiven_print(book: Book, answers: Optional[dict] = None) -> Union[Effects, flow.Question]:
    """UFMitsubishi.CommandButton15_Click (UFMit:1363-1497) 5N の印刷・保存(Choice が 5N でなければ何もしない)。"""
    fx = book.begin()
    f = book.mit
    if ZEN not in book.sheets:
        return fx
    if _height_missing(book):
        fx.msg("検入高さを入力してください。", flow.ERROR, "UFMit:1377")
        return fx
    q = _red_x_question(book, answers, "zen_red_x")
    if isinstance(q, flow.Question):
        return q
    if q is False:
        return fx
    if not f.choice == "5N":
        return fx
    book.activate_sheet(ZEN)
    if f.kname == "" and f.ken_no == "":
        fx.msg("検査員名と検査番号を入力してください", source="UFMit:1394")
        return fx
    f.tag = "off"
    fx.sounds.append("おまちください.wav")
    zen_transfer(book)
    job = zen_print_setup(book)
    fx.sounds.append("印刷.wav")
    a = (answers or {}).get("zen_print")
    if a is None:
        return flow.Question("zen_print", "印刷しますか？", source="UFMit:1427", messages=list(fx.messages))
    if a:
        book.ws(ZEN).visible = True
        if job:
            fx.prints.append(job)
        _defect_print(book, f.kname, f.lot_box)
        fiven_save(book)
        for name in ("ｺﾋﾟ-用", "5N", "全ヵ所検査保存", "全ヵ所作業中保存"):
            book.delete_sheet(name)
        if not book.has("保存用"):
            return fx                                   # Tag は "off" のまま(VBA どおり)
        row = sh.history_row(book, book.ws(ZEN), "全ヵ所")
        if row is not None:
            fx.history.append(row)
        fx.sounds.append("完了.wav")
    else:
        fx.msg("ｷｬﾝｾﾙ処理", source="UFMit:1492")
    f.tag = ""
    return fx


def mitsubishi_print(book: Book, answers: Optional[dict] = None) -> Union[Effects, flow.Question]:
    """UFMitsubishi.CommandButton2_Click (UFMit:1515-1603) 三菱受託の印刷・保存。

    ﾌﾞﾛｯｸNO のチェックは行 7,9,11,13 の B 列に値があり、列 27(AA)または 29(AC)が空のとき(W2: 列がずれている疑い)。
    """
    fx = book.begin()
    f = book.mit
    f.tag = "off"
    fx.sounds.append("おまちください.wav")
    if ZEN in book.sheets:
        book.activate_sheet(ZEN)
    f.work, f.inflag = "ON", "flagon"
    if ZEN not in book.sheets:
        return fx
    ws = book.ws(ZEN)
    for i in range(7, 14, 2):
        if not sh.blank(ws.get(i, 2)) and (sh.blank(ws.get(i, 27)) or sh.blank(ws.get(i, 29))):
            fx.msg("ﾌﾞﾛｯｸNOの入力をしてください", flow.WARN, "UFMit:1534")
            return fx
    if f.kname == "" and f.ken_no == "":
        fx.msg("検査員名と検査番号を入力してください", source="UFMit:1540")
        return fx
    zen_transfer(book)
    job = genshi_print_setup(book)
    fx.sounds.append("印刷.wav")
    a = (answers or {}).get("zen_print")
    if a is None:
        return flow.Question("zen_print", "印刷しますか？", source="UFMit:1560", messages=list(fx.messages))
    if a:
        if book.has("転送済"):
            book.ws("転送済").visible = True
        if job:
            fx.prints.append(job)
        zen_save(book)
        for name in ("ｺﾋﾟ-用", "転送済", "全ヵ所検査保存", "全ヵ所作業中保存"):
            book.delete_sheet(name)
        row = sh.history_row(book, book.ws(ZEN), "三菱受託")
        if row is not None:
            fx.history.append(row)
        fx.sounds.append("完了.wav")
    else:
        fx.msg("ｷｬﾝｾﾙ処理", source="UFMit:1597")
    f.tag = ""
    return fx


def warn_only(book: Book, answers: Optional[dict] = None) -> Union[Effects, flow.Question]:
    """UFMitsubishi.CommandButton18_Click (UFMit:73-97)。検入高さ・公差外の警告だけ(後の処理なし)。"""
    fx = book.begin()
    if ZEN not in book.sheets:
        return fx
    if _height_missing(book):
        fx.msg("検入高さを入力してください。", flow.ERROR, "UFMit:83")
        return fx
    q = _red_x_question(book, answers, "zen_red_x")
    return q if isinstance(q, flow.Question) else fx


# ------------------------------------------------------------------
# 起動・終了
# ------------------------------------------------------------------
def open_mitubisi(book: Book, temp: str, *, line_name: str, worker_text: Optional[str] = None,
                  choice_input: Optional[str] = None, **lot_kw: Any) -> Union[Effects, flow.Question]:
    """Mitsubishi.OpenMitubisi (Mit:3-76) + UserForm_Initialize + Show(Activate)。

    `temp`: "全数"(全ヵ所) / "H494"。`worker_text`: UF1 の当日の検査員(T1/T2/T3 の Caption。"_" 区切りの最後の人)。
    シートからロット・検査No・作業選択・用途・高さ式を戻す。ロット欄を戻すと LotBox_Change → ロット読込が走る(`lot_kw`)。
    作業選択が空なら InputBox(`choice_input`。None なら `flow.Question(buttons="input")` を返す)。
    """
    book.begin()
    book.mit = sh.MitForm(line_name=line_name)
    f = book.mit
    f.pack_no_on_off = "off"
    worker = worker_text.split("_") if worker_text is not None else None
    name = ZEN if temp == "全数" else H494
    f.work_flag = "P1" if temp == "全数" else "P_H494"
    if name in book.sheets:
        ws = book.ws(name)
        book.activate_sheet(name)
        if worker is not None:
            h3 = sh.cell_str(ws.get(3, 8))
            new = worker[-1] + " " + h3 if h3 != "" else worker[-1]
            f.kname = remove_duplicates(new)
        lot = sh.cell_str(ws.get(2, 4))
        if temp == "全数":
            if f.lot_box != lot:
                f.lot_box = lot
                lot_box_change(book, **lot_kw)
            f.ken_no = sh.cell_str(ws.get(2, 8))
            _set_choice(book, sh.cell_str(ws.get(1, 22)))
            f.youto_no = sh.cell_str(ws.get(4, 8))
            f.takasa_formula = sh.cell_str(ws.get(1, 30))
            if ws.font_color.get((2, 4)) == (255, 0, 0):
                f.h494 = True
        else:
            if f.h494_lot != lot:
                f.h494_lot = lot
                from . import h494 as h
                h.h494_lot_change(book, **lot_kw)
            f.h494_ken_no = sh.cell_str(ws.get(2, 8))
            f.h494_youto = sh.cell_str(ws.get(4, 8))
            f.takasa_formula = sh.cell_str(ws.get(1, 30))
    res = activate(book, choice_input=choice_input)
    if isinstance(res, flow.Question):
        return res
    f.visible = True
    f.pack_no_on_off = "on"
    return book.fx


def activate(book: Book, *, choice_input: Optional[str] = None) -> Union[Effects, flow.Question]:
    """UFMitsubishi.UserForm_Activate (UFMit:2194-2268)。丸徳は「一般」固定。作業選択が空なら InputBox。"""
    fx = book.fx
    f = book.mit
    if f.tag == "off":
        return fx
    if f.work_flag == "P1":
        if f.line_name == "丸徳":
            _set_choice(book, "一般")
        if f.choice == "":
            if choice_input is None:
                return flow.Question("mit_choice", "作業名入力：三菱、5N、一般　のいずれかを入力", buttons="input",
                                     source="UFMit:2210", messages=list(fx.messages))
            temp = choice_input
            if temp == "５N":
                temp = "5N"
            if temp == "":
                raise flow.Stop("作業名が入力されていません", source="UFMit:2213", messages=list(fx.messages))
            _set_choice(book, temp)
            forn_select(book)
        else:
            forn_select(book)
        for name in (ZEN, "全ヵ所検査保存", "全ヵ所作業中保存"):
            if name in book.sheets or name in book.extra:
                book.ws(name).visible = True
        if ZEN in book.sheets:
            book.activate_sheet(ZEN)
    elif f.work_flag == "P_H494":
        for name in (H494, "H494検査保存", "H494作業中保存"):
            if name in book.sheets or name in book.extra:
                book.ws(name).visible = True
        if H494 not in book.sheets:
            zen_copy(book, H494)
        book.activate_sheet(H494)
    book.events_enabled = True
    return fx


def query_close(book: Book) -> Effects:
    """UFMitsubishi.UserForm_QueryClose (UFMit:2344-2366)。全ヵ所・H494 系のシートを隠す。"""
    fx = book.begin()
    for name in (ZEN, "全ヵ所検査保存", "全ヵ所作業中保存", H494, "H494検査保存", "H494作業中保存"):
        if name in book.sheets or name in book.extra:
            book.ws(name).visible = False
    book.mit.visible = False
    return fx
