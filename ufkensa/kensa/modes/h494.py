"""H494 検査(用途コード H494・10 枚)— UFMitsubishi(P_H494)+ H494検査シート

- シートのイベント: `H494_SH1`(= `H494_SH`)の Worksheet_Change / SelectionChange
- ロット読込は全ヵ所と共用(`zenkasho.combo_change`。WorkFlag = "P_H494" で H494検査 に書く)
- 転送(CB19)・印刷と保存(CB23)・新規発行(ws_out)

B30: H494 では判定ログの Lname が空(丸徳のログに出る)、印刷フッタの測定機器が AC3 を見る(機器は Y2 に書く)、
検査員が UFMitsubishi.Kname に入り H494_Kname が空のまま(印刷・保存の検査員欄が空)。既定は VBA どおり。
"""
from __future__ import annotations

from typing import Any, Optional, Union

from .. import flow, quirks, vbx
from . import sheet as sh
from . import zenkasho as zk
from .sheet import H494, Book, Call, Effects, LogCall, VbaError

_RAN = ((2, 2, 3, 6), (2, 11, 4, 15), (2, 17, 4, 19), (2, 21, 4, 26), (7, 2, 16, 2), (7, 4, 16, 4), (7, 6, 16, 6),
        (7, 8, 16, 8), (7, 11, 16, 11), (7, 13, 16, 13), (7, 15, 16, 15), (7, 17, 16, 17), (7, 19, 16, 19),
        (7, 21, 16, 22), (7, 24, 16, 28))


def worksheet_change(book: Book, target: sh.Sheet, r: int, c: int) -> None:
    """H494_SH1.Worksheet_Change (H494_SH:10-96)。On Error Resume Next の下で動く。"""
    ws = book.active
    befor = book.resume(lambda: sh.take_befor(book, book.mit, r, c), "H494_SH:20") or ""
    TR, TC = r, c
    if ws.name == H494 and 7 <= TR <= 16:
        if 2 <= TC <= 8:
            book.resume(lambda: book.activate(TR, TC + 1))
        elif TC == 9:
            if sh.blank(ws.get(TR, 11)):
                book.resume(lambda: book.activate(TR, 11))
            else:
                book.resume(lambda: book.activate(TR + 1, 2))
        elif 11 <= TC <= 20:
            book.resume(lambda: book.activate(TR, TC + 1))
        elif TC == 21:
            if not sh.blank(ws.get(TR, 21)):
                book.fx.sounds.append("音楽対角")
                book.resume(lambda: sh.uf_taikaku(book, TR, TC))
                book.events_enabled = False
                for cc in (22, 23, 24):
                    book.write(ws, TR, cc, "")
                book.events_enabled = True
                book.resume(lambda: book.activate(TR, TC + 1))
        elif TC == 22:
            book.events_enabled = False
            book.resume(lambda: sh.al_ter_range(book, TR, TC, befor), "AlTerRange")
            book.resume(lambda: sh.taikaku(book, TR, TC), "対角")
            book.events_enabled = True
            book.resume(lambda: book.activate(TR + 1, 2))
            return
        elif TC == 25:
            book.events_enabled = False
            v = ws.get(TR, 25)
            if not sh.is_numeric(v):
                pass
            elif v == "-":
                pass
            else:
                # 空でも UnitRecord(0) → "0.0" を書く(全ヵ所と違い空を飛ばさない)
                book.resume(lambda: book.write(ws, TR, 25, vbx.fmt(sh.unit_record(sh.val_cell(v)), "0.0")))
            book.resume(lambda: sh.lateralbow_z(book, TR, TC), "Lateralbow_Z")
            book.events_enabled = True
            book.resume(lambda: book.activate(TR, TC + 1))
            return
    if ws.name == H494:
        book.events_enabled = False
        book.resume(lambda: sh.al_ter_range(book, TR, TC, befor), "AlTerRange")
        if not sh.is_numeric(target.get(TR, TC)):
            if not quirks.reproduce("B79"):
                book.events_enabled = True           # 直した側: 数でない文字(外観「良」など)でもイベントを戻す
            return                                # B79: EnableEvents は False のまま(VBA どおり)
        book.events_enabled = True
        if not sh.blank(target.get(TR, TC)):
            book.resume(lambda: sh.kensa_hantei(book, TR, TC), "検査判定")
        book.events_enabled = False
        book.resume(lambda: sh.taikaku(book, TR, TC), "対角")
        book.events_enabled = True
        book.resume(lambda: sh.koshagai_kensaku(book, TR, TC), "公差外検索")


def worksheet_selection_change(book: Book, r: int, c: int) -> None:
    """H494_SH1.Worksheet_SelectionChange (H494_SH:98-157)。判定セルも含め、選んだセルを薄青にする。"""
    if book.mit.visible:
        book.mit.cnt = "0"
    ws = book.active
    if ws.name == H494:
        if c == 24:
            book.activate(r, 25)
        elif c == 26:
            book.events_enabled = False
            book.write(ws, 2, 8, book.mit.h494_ken_no)
            book.write(ws, 3, 8, book.mit.h494_kname)
            ws.font_size[(2, 8)] = ws.font_size[(3, 8)] = 6
            book.events_enabled = True
    if ws.name == H494 and (r, c) == (7, 4):
        book.events_enabled = False
        if not sh.blank(ws.get(7, 2)):
            book.write(ws, 1, 2, sh.fmt_date(book.now(), "yyyy/mm/dd h:mm"))
            ws.font_size[(1, 2)] = 5
            book.fx.calls.append(Call("Daily_Report", (1,)))
            book.fx.calls.append(Call("Daily_Report", (3,)))
        book.events_enabled = True
    if ws.name != H494:
        return
    # B81: ラテラルボー(Y 列)はセル自体を判定の色に塗るが、Ran(Y 列を含む)を白にするので直後に消える。
    # 直すと、判定の色(緑・赤)の Y 列のセルは白にも薄青にもしない
    keep = {}
    if not quirks.reproduce("B81"):
        keep = {(r, 25): ws.interior.get((r, 25)) for r in range(7, 17)
                if ws.interior.get((r, 25)) in (sh.COLOR_INDEX[10], sh.NG_RED)}
    for rect in _RAN:
        ws.fill(*rect, sh.COLOR_INDEX[2])
    ws.interior.update(keep)
    # NpRan は Nothing → 実行時エラー(On Error Resume Next で飛ばす)
    ar, ac = book.active_cell
    if (ar, ac) not in keep:
        ws.interior[(ar, ac)] = sh.LIGHT_BLUE
    ws.font_color[(ar, ac)] = (0, 0, 0)


sh.HANDLERS[H494] = (worksheet_change, worksheet_selection_change)


def h494_lot_change(book: Book, **kw: Any) -> Effects:
    """UFMitsubishi.H494_Lot_Change (UFMit:321-330)。7 桁で ComboBox1 に写す(→ ロット読込)。"""
    fx = book.fx
    book.events_enabled = False
    if len(book.mit.h494_lot) != 7:
        return fx
    if book.mit.combo != book.mit.h494_lot:
        book.mit.combo = book.mit.h494_lot
        zk.combo_change(book, **kw)
    book.events_enabled = True
    return fx


def transfer(book: Book) -> Effects:
    """UFMitsubishi.CommandButton19_Click (UFMit:99-124) H494 の転送。検査法(op1〜3)→ 最大最少 → Ainput_Startup("H494", 検査法)。"""
    fx = book.begin()
    f = book.mit
    if f.tag == "off":
        return fx
    if f.h494_method == "":
        fx.msg("検査法を選択してください", flow.ERROR, "UFMit:111")
        return fx
    book.events_enabled = False
    try:
        zk.max_min(book, H494)
    except VbaError as e:
        raise flow.Stop(f"実行時エラー '{e.number}':" + "\r\n\r\n" + e.description, source="Mit:全ヵ所用最大最少")
    fx.calls.append(Call("Ainput_Startup", ("H494", f.h494_method)))
    book.events_enabled = True
    return fx


def print_and_save(book: Book, answers: Optional[dict] = None) -> Union[Effects, flow.Question]:
    """UFMitsubishi.CommandButton23_Click (UFMit:138-230) H494 の印刷・保存。

    測定機器は Y2 に書き(`全ヵ所用機器転送(ZenSh, 2, 25, …, "H494")`)、印刷のフッタは AC3 を見る(B30)。
    """
    fx = book.begin()
    f = book.mit
    zs = book.ws(H494)
    atu, haba = sh.cell_str(zs.get(3, 2)), sh.cell_str(zs.get(3, 4))
    f.tag = "off"
    book.activate_sheet(H494)
    if f.kname == "" and f.h494_kname == "":
        fx.msg("検査員名と検査番号を入力してください", source="UFMit:149")
        return fx                                      # Tag は "off" のまま
    fx.sounds.append("おまちください.wav")
    zk.kiki_tensou(book, zs, 2, 25, atu, haba, "H494")
    job = zk.zen_print_setup(book, H494)
    fx.sounds.append("印刷.wav")
    a = (answers or {}).get("h494_print")
    if a is None:
        return flow.Question("h494_print", "印刷しますか？", source="UFMit:171", messages=list(fx.messages))
    if a:
        zs.visible = True
        if job:
            fx.prints.append(job)
        zk.zen_save(book, H494)
        kname = f.h494_kname if (quirks.reproduce("B30") or f.h494_kname != "") else f.kname
        zk._defect_print(book, kname, f.lot_box)
        for name in ("ｺﾋﾟ-用", "H494検査保存", "H494作業中保存"):
            book.delete_sheet(name)
        row = sh.history_row(book, zs, "H494")
        if row is not None:
            fx.history.append(row)
        fx.sounds.append("完了.wav")
    else:
        fx.msg("ｷｬﾝｾﾙ処理", source="UFMit:222")
    f.tag = ""
    return fx


def new_sheet(book: Book, answers: Optional[dict] = None) -> Union[Effects, flow.Question]:
    """UFMitsubishi.ws_out_Click (UFMit:2114-2151) H494 の新規発行。"""
    fx = book.begin()
    f = book.mit
    if f.tag == "off":
        return fx
    a = (answers or {}).get("h494_new")
    if a is None:
        return flow.Question("h494_new", "新規発行しますか？現在のｼｰﾄは削除されます。", source="UFMit:2118")
    fx.calls.append(Call("Register_Check/Daily_Create"))
    if a:
        zk.justincase_zn_save(book, H494)
        zk.zen_copy(book, H494)
        fx.logs.append(LogCall("*UFKen:" + "H494検査ｼｰﾄ発行", "", sh.log_dir(f.line_name)))
        dname = zk._defect_sheet_name(book)
        if book.has(dname):
            book.delete_sheet(dname)
        if f.h494_lot != "":
            f.h494_lot = ""
            h494_lot_change(book)
        f.h494_ken_no = ""
        if f.combo != "":
            f.combo = ""
            zk.combo_change(book)
        f.h494_youto = ""
    else:
        fx.msg("ｷｬﾝｾﾙ処理", source="UFMit:2149")
    return fx
