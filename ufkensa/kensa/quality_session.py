"""品質基準画面(VBA の UF品質基準)の組み立て ── 画面(app/routes/quality.py)はここだけを呼ぶ

状態は持たない(読み取り専用)。ロット 1 つについて、品質基準 DB の照合を `quality.py` / `quality_view.py`
(QP_DB / Quality_DB の移植)で行い、**規格ごとのカード**にまとめる。各カードには参照元(表名・規格番号)を添える。

VBA の流れ(UF1.CommandButton8_Click, UF1:3053-3111)と同じ部品を同じ引数で呼ぶ:
    DistArray(Lot, KenNo) → FormReflectDist(UF品質基準 表示中)        … 歪(平坦度)カード
    FlagAtu(製造厚) >= 8 なら DBdisplay外観_H                          … 外観(厚板)カード
    B1〜B11 の DBdisplay外観_S(ボタンを全部押した結果を表にする)       … 外観(薄板)カード
    手動検索(CommandButton1)の DBdisplay歪                             … 参考カード
検査で使う規格のうち UF品質基準 に出ないもの(BH・ラテラルボー・直角度・対角・ラテフリ直角)も、
検査画面と同じ部品(BhSearch / Lateralbow_J / Diagonal / DbReturnItem)で引いて並べる。

ロットの値は AccReflec(`lot.acc_reflec`。Lot情報履歴 は書かない)、引当は `lot.get_lot_info`(Lot情報履歴)、
無ければ SIKAHIKINOW の引当(HikiArrCreate)。通常検査の状態は**読むだけ**(いまのロット・DBflag・画面の公差)。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable, Optional

from . import datasource, judge, lines, lot as lotmod, quality, quality_view as qv, terminal, vbx, workers, wiring
from .flow import Message, Stop
from .store import LOT_HISTORY, LocalStore

#: UF品質基準 の薄板ページの異常ボタン(B1〜B11 の Caption。G6)
SURFACE_BUTTONS = ("トビコミ", "キズ,スリキズ", "共ズレ", "突上げ,突下げ", "汚れ,汚れ圧込み", "油寄り,レベラースジ",
                   "ロールキズ", "ロールマーク", "噛み込み跡", "切断面異常", "エッジ変色,油残り跡,油焼け")
#: 計算で起きうる VBA の実行時エラー(カード 1 枚の失敗にして、ほかのカードは出す)
ERRORS = (quality.VbaRuntimeError, vbx.TypeMismatch, IndexError, KeyError, ValueError, ZeroDivisionError)


@dataclass
class Card:
    key: str
    title: str
    sources: list[str] = field(default_factory=list)       # 参照元(表名・規格番号)
    status: str = "ok"                                       # ok / none(該当なし・不要)/ error
    rows: list[tuple[str, str]] = field(default_factory=list)
    table: Optional[dict[str, Any]] = None                   # {"head": [...], "rows": [[...]]}
    note: str = ""
    updated: str = ""                                        # 表の更新日(KousinL*)
    messages: list[Message] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {"key": self.key, "title": self.title, "sources": [s for s in self.sources if s], "status": self.status,
                "rows": [list(r) for r in self.rows], "table": self.table, "note": self.note, "updated": self.updated,
                "messages": [m.to_dict() for m in self.messages]}


def _s(v: Any) -> str:
    return quality.vstr(v) if v is not None else ""


def _f1(v: Any) -> str:
    """`Format(x, "0.0")`(数でなければ文字のまま)。"""
    return vbx.fmt(v, "0.0") if quality.isnum(v) else _s(v)


def _source_tokens(text: str) -> list[str]:
    """参照元の文字列(Sansyou1 や 平坦度上限値参照元)を空白で区切った語。重複は除く。"""
    return list(dict.fromkeys(t for t in text.replace("\r\n", " ").split(" ") if t))


def _guard(card: Card, fn: Callable[[], None]) -> Card:
    try:
        fn()
    except datasource.DbMissing as e:
        card.status, card.note = "error", str(e)
    except Stop as e:
        card.status, card.note = "error", e.text
    except ERRORS as e:                                      # VBA なら実行時エラーで止まる所
        card.status, card.note = "error", f"照合でエラーになりました({type(e).__name__}: {e})"
    return card


# ------------------------------------------------------------------
# ロットの値
# ------------------------------------------------------------------
@dataclass
class LotView:
    lot: str
    dbflag: str = ""
    reflect: Optional[lotmod.LotReflect] = None
    data: Optional[quality.LotData] = None                   # Arrbox(SIKALOTNOW) の Fields / Record
    ken_no: str = ""
    equipment: str = ""
    messages: list[Message] = field(default_factory=list)

    def row(self, name: str) -> str:
        return _s((self.reflect.row if self.reflect else {}).get(name))

    def form(self, name: str) -> str:
        return (self.reflect.form if self.reflect else {}).get(name, "")


def read_lot(lot_no: str, *, dbflag: str = "", facility: str = "", ken_no: str = "") -> LotView:
    """AccReflec(SaveLoad:56-345)と Arrbox(SIKALOTNOW / SIKAHIKINOW)。設備仕様は結果を見せるだけで止めない。"""
    lv = LotView(lot_no, dbflag)
    seen: list[str] = []

    def check_equipment(*args: Any) -> str:
        seen.append(quality.check_equipment_db(*args))
        return "可"
    screen = lotmod.LotScreen(lot=lot_no, cb_text=lot_no, dbflag=dbflag, facility=facility or workers.DEFAULT_FACILITY,
                              work_time="")
    lv.reflect = lotmod.acc_reflec(screen, check_equipment=check_equipment)
    lv.messages += lv.reflect.messages
    lv.equipment = seen[0] if seen else ""
    f = lotmod.arrbox("SIKALOTNOW", "板", "Fields", dbflag)
    r = lotmod.arrbox("SIKALOTNOW", lot_no, "Record", dbflag)
    lv.messages += f.messages + r.messages
    lv.data = quality.LotData(f.value, r.value)
    lv.ken_no = ken_no
    if lv.ken_no == "":
        # LSelect と同じ: SIKAHIKINOW の 検査番号、無ければ SIKALOTNOW の 検査番号(FormRelat / layout.lselect)
        hf = lotmod.arrbox("SIKAHIKINOW", "", "Fields", dbflag)
        hr = lotmod.arrbox("SIKAHIKINOW", lot_no, "Record", dbflag)
        if hr.value and hf.value:
            lv.ken_no = _s(quality.cell(hr.value[0], quality.col_of(hf.value, "検査番号")))
        elif lv.data.records and lv.data.fields:
            lv.ken_no = lv.data.value("検査番号")
    return lv


def allocations(store: LocalStore, lv: LotView) -> tuple[str, list[list[Any]]]:
    """引当の行 [納入先, 得意先, 用途, 材質, 調質, 厚, 巾, 丈, ﾛｯﾄNo, 検査No]。

    Lot情報履歴(GetLotInfo。検査画面の JudgeLotAllocations と同じ)にあればそれ、無ければ SIKAHIKINOW の引当
    (受注_納入先名称・受注_取引先名称)に仕掛の用途・材質・調質・製造寸法を合わせる。
    """
    rows = lotmod.get_lot_info(wiring.sheet_rows(store, LOT_HISTORY), lv.lot)
    if rows:
        return "Lot情報履歴", rows
    hiki = quality.hiki_arr_create(lv.lot) or []
    f = lv.form
    return "引当(SIKAHIKINOW)", [[n, t, f("Youto"), f("Zai"), f("Tyu"), f("ATU"), f("Hab"), f("Tak"), lv.lot, lv.ken_no]
                                 for n, t in hiki]


# ------------------------------------------------------------------
# カード
# ------------------------------------------------------------------
def card_distortion(lv: LotView) -> Card:
    """歪(平坦度)。DistArray → FormReflectDist(UF品質基準 表示中)。UF1:3102-3106。"""
    c = Card("distortion", "歪(平坦度)", ["品質基準: 平坦度一覧QP追加"])

    def run() -> None:
        qp = quality.dist_array(lv.lot, lv.ken_no, quality.hiki_arr_create(lv.lot), lv.data)
        c.messages += qp.messages
        if not qp.kousa:
            c.status, c.note = "none", "該当する歪公差がありません"
            return
        k = qp.kousa[0]
        r = quality.form_reflect_dist(qp.kousa, quality_visible=True, mitsubishi_visible=False)
        v = r.values
        c.sources = _source_tokens(_s(k[quality.K_SOURCE])) + c.sources
        c.rows = [("1M当りｹﾞｰｼﾞ", v.get("MG1", "")), ("巾歪", v.get("HaHiz1", "")), ("丈歪", v.get("TaHiz1", "")),
                  ("中歪", v.get("Naka1", "")), ("耳歪", v.get("Mimi1", "")), ("管理番号", _s(k[quality.K_NO])),
                  ("フラグ", _s(k[quality.K_FLAG]))]
        c.note = _s(k[quality.K_COMMENT]).rstrip("_").replace("。", "\n")
        c.updated = _s(qp.form.get("KousinL2.Caption", ""))
        if not any(x for _, x in c.rows[:5]):
            c.status = "none"                            # 判定不要(コメントに理由)
    return _guard(c, run)


def card_distortion_ref(lv: LotView) -> Card:
    """参考: UF品質基準 の手動検索の DBdisplay歪(DIM-122 系の表引き)。QDB:700-1384。"""
    c = Card("distortion_ref", "参考: 歪の表引き(手動検索)")

    def run() -> None:
        r = qv.display_distortion(lv.data, qv.QualityForm(), quiet=True)
        f = r.form
        c.messages += r.messages
        c.sources = _source_tokens(f.Sansyou1) or ["品質基準: Q_S_QA_DIM_122_A〜D・QA_022_01(参照元の追記なし)"]
        c.rows = [("1M当りｹﾞｰｼﾞ", f.MG1), ("巾歪", f.HaHiz1), ("丈歪", f.TaHiz1), ("中歪", f.Naka1), ("耳歪", f.Mimi1)]
        if f.Toukyu:
            c.rows += [("等級", f.Toukyu), ("中歪(2列目)", f.Naka2), ("耳歪(2列目)", f.Mimi2)]
        if f.VcFlag:
            c.rows.append(("VC", f.VcFlag))
        c.note = "検査の判定に使うのは「歪(平坦度)」(平坦度一覧QP追加)。これは VBA の手動検索と同じ表引きです。"
        if not any(v for _, v in c.rows):
            c.status = "none"
    return _guard(c, run)


def card_appearance(lv: LotView, flag_atu: str) -> Card:
    """外観。FlagAtu >= 8 なら DBdisplay外観_H(UF1:3110)、それ以外は薄板の DBdisplay外観_S(B1〜B11)。"""
    thick = vbx.val(flag_atu) >= 8 if flag_atu != "" else False
    if thick:
        c = Card("appearance", "外観(厚板)", ["Q-S-QA-124 別表1", "Q-S-QA-124 別表2"])

        def run() -> None:
            r = qv.display_appearance_h(lv.data, qv.QualityForm(), quiet=True)
            c.messages += r.messages
            c.updated = r.form.KousinL1
            ga = r.grade_array or []
            rows = [[_s(x[0]), _s(x[1]), _s(x[2])] for x in ga if x and x[0] is not None]
            if not rows:
                c.status, c.note = "none", "該当する外観グレードがありません"
                return
            c.table = {"head": ["欠陥名", "許容内容", "グレード"], "rows": rows}
        return _guard(c, run)

    c = Card("appearance", "外観(薄板)", ["Q-S-QA-SRF-128 別表1", "Q-S-QA-SRF-128 別表2"])

    def run_s() -> None:
        rows = []
        pattern = ""
        for name in SURFACE_BUTTONS:
            r = qv.display_appearance_s(lv.data, name, qv.QualityForm(), quiet=True)
            c.messages += [m for m in r.messages if all(m.text != x.text for x in c.messages)]
            f = r.form
            pattern = f.Upater
            c.updated = f.KousinL4 or c.updated
            if f.SelectO or f.SelectS or f.SelectK or f.SelectA:
                rows.append([name, f.SelectO, f.SelectS, f.SelectK, f.SelectA])
        c.rows = [("パターン", pattern)]
        if not rows:
            c.status, c.note = "none", pattern if not vbx.is_numeric(pattern) else "該当する基準がありません"
            return
        c.table = {"head": ["異常名", "大きさ", "触感", "数", "素洗い"], "rows": rows}
    return _guard(c, run_s)


def card_bh(lv: LotView) -> Card:
    """BH(バーコール硬度の下限)。QDB.BhSearch(QDB:1454-1490)。0 は不要(Judg.BhJudg)。"""
    c = Card("bh", "BH(バーコール硬度)", ["Q-S-KE-ALL-053"])

    def run() -> None:
        v = quality.bh_search(lv.form("Zai"), lv.form("Tyu"))
        if v == 0:
            c.status = "none"
            c.rows = [("下限値", "不要")]
        else:
            c.rows = [("下限値", str(v))]
        c.rows += [("材質・調質", f"{lv.form('Zai')} {lv.form('Tyu')}".strip())]
    return _guard(c, run)


def _dims(lv: LotView) -> tuple[float, float, float]:
    return (vbx.single(vbx.val(lv.form("Hab"))), vbx.single(vbx.val(lv.form("Tak"))),
            vbx.single(vbx.val(lv.form("ATU"))))


def card_diagonal(lv: LotView) -> Card:
    """対角(対角線差の公差)。指示_直角度 があればそれ、無ければ 直角度等級 4 → Diagonal、3 → ReDiagonal(SaveLoad:283-305)。"""
    c = Card("diagonal", "対角(対角線差)", ["品質基準: 製品規格直角度"])

    def run() -> None:
        r = lv.reflect
        grade = lv.form("TAIK")
        c.rows = [("直角度等級", grade)]
        if r is not None and r.taik_text is not None:
            c.rows += [("指示_直角度", r.shiji_taik), ("公差", r.taik_text)]
            c.sources = ["仕掛: 指示_直角度"]
            c.note = "技術指示の公差です(" + r.additional + ")"
            return
        hab, tak, _ = _dims(lv)
        msgs: list[Message] = []
        if grade == "4":
            v = vbx.fmt(judge.diagonal(hab, tak, messages=msgs), "0.0")
        elif grade == "3":
            v = vbx.fmt(judge.re_diagonal(hab, tak, messages=msgs), "0.0")
        else:
            v = ""
        c.messages += msgs
        c.rows.append(("公差", v))
        if v == "":
            c.status = "none"
    return _guard(c, run)


def card_lateral(lv: LotView) -> Card:
    """ラテラルボーの公差。指示_ﾗﾃﾗﾙﾎﾞｰ があればそれ、無ければ等級 4 → Lateralbow_J、3 → ReLateralbow(SaveLoad:306-327)。"""
    c = Card("lateral", "ラテラルボー", ["品質基準: 製品規格ラテラルボー"])

    def run() -> None:
        r = lv.reflect
        grade = lv.form("Lateral")
        c.rows = [("ラテラルボー等級", grade or "なし")]
        if r is not None and r.lateral_text is not None:
            c.rows += [("指示_ﾗﾃﾗﾙﾎﾞｰ", r.shiji_late), ("公差", r.lateral_text)]
            c.sources = ["仕掛: 指示_ﾗﾃﾗﾙﾎﾞｰ"]
            c.note = "技術指示の公差です(" + r.additional + ")"
            return
        hab, tak, atu = _dims(lv)
        msgs: list[Message] = []
        txt = ""
        if grade == "4":
            txt = vbx.fmt(judge.lateralbow_j(hab, tak, atu, messages=msgs), "0.0")
        elif grade == "3":
            txt = vbx.fmt(judge.re_lateralbow(hab, tak, atu, messages=msgs), "0.0")
        if txt != "" and vbx.u1_compare(txt, 0) == 0:
            txt = ""
        c.messages += msgs
        c.rows.append(("等級からの公差", txt))
        c.note = "ラテフリ直角 でラテラルボー「要」の引当は、その公差で上書きされます(下のカード)。"
        if txt == "":
            c.status = "none"
    return _guard(c, run)


def card_ratefuri(store: LocalStore, lv: LotView) -> Card:
    """ラテラルボー・直角度・フリー反りの要否(ラテフリ直角)。引当ごとに DbReturnItem(QP:1176-1456)。

    同じ 8 項目(納入先・得意先・用途・材質・調質・厚巾丈)の引当は 1 回だけ(FormRelat.JudgeLotAllocations)。
    """
    c = Card("ratefuri", "直角度・ラテラルボー・フリー反りの要否", ["品質基準: ラテフリ直角"])

    def run() -> None:
        src, rows = allocations(store, lv)
        c.rows = [("引当の出どころ", src)]
        out, done = [], set()
        used_179 = set()
        for row in rows:
            delivery = _s(row[0])
            if delivery == "":
                continue
            atu, hab, tak = quality.csng(row[5]), quality.csng(row[6]), quality.csng(row[7])
            key = (delivery, _s(row[1]), _s(row[2]), _s(row[3]), _s(row[4]), vbx.cstr(atu, is_single=True),
                   vbx.cstr(hab, is_single=True), vbx.cstr(tak, is_single=True))
            if key in done:
                continue
            done.add(key)
            r = quality.db_return_item(delivery, _s(row[1]), _s(row[2]), _s(row[3]), _s(row[4]), atu, hab, tak,
                                       additional_caption=lv.reflect.additional if lv.reflect else "",
                                       lateral_lot=lv.data, lateral_quiet=True)
            c.messages += r.messages
            items = r.items
            u = r.uf1
            if "LateralText.BackColor" in u:
                used_179.add("Q-S-QA-DIM-179")
            out.append([delivery, _s(row[1]),
                        "要" if "フリー反り" in items else "", "要" if "ラテラルボー" in items else "",
                        "要" if "直角度" in items else "",
                        _s(u.get("LateralText", "")), _s(u.get("HURIS", "")), _s(u.get("Additional.Caption", ""))])
        c.sources += sorted(used_179)
        if not out:
            c.status, c.note = "none", "引当がありません(Lot情報履歴・SIKAHIKINOW とも)"
            return
        c.table = {"head": ["納入先", "得意先", "ﾌﾘｰ反り", "ﾗﾃﾗﾙﾎﾞｰ", "直角度", "ﾗﾃﾗﾙ公差", "ﾌﾘｰ反り公差", "注記"],
                   "rows": out}
        if not any(r[2] or r[3] or r[4] for r in out):
            c.status = "none"
            c.note = "どの引当も 要 の項目はありません"
    return _guard(c, run)


# ------------------------------------------------------------------
# まとめ
# ------------------------------------------------------------------
def current_lot(kensa: Any = None) -> dict[str, str]:
    """通常検査でいま検査中のロット(読むだけ)。"""
    if kensa is None:
        from . import session
        kensa = session.get()
    st = kensa.state
    return {"lot": st.lot, "dbflag": st.dbflag, "line": st.lname}


def _kensa_values(st: Any) -> list[tuple[str, str]]:
    """検査画面にいま出ている規格(いまのロットのとき。読むだけ)。"""
    v = lambda n: st.v(n).value          # noqa: E731

    def join(*xs: str) -> str:
        return " / ".join(x or "-" for x in xs) if any(xs) else ""
    rows = [("厚", f"{st.atu}  +{st.ap}/-{st.am}" if st.atu else ""), ("巾", f"{st.hab}  +{st.hp}/-{st.hm}" if st.hab else ""),
            ("丈", f"{st.tak}  +{st.tp}/-{st.tm}" if st.tak else ""), ("対角", v("TaikText")),
            ("ラテラルボー", st.lateral_text), ("フリー反り", st.flatnes_text + (f" {st.huris}" if st.huris else "")),
            ("歪 ゲージ/巾/丈", join(v("MG"), v("HaHiz"), v("TaHiz"))), ("歪 中/耳", join(v("Naka"), v("Mimi"))),
            ("注記", st.additional)]
    return [(k, x.strip()) for k, x in rows if x.strip()]


def build(lot_no: str, *, store: LocalStore, kensa: Any = None, facility: Optional[str] = None) -> dict[str, Any]:
    """ロット 1 つの品質基準(読み取り専用)。lot_no が "" ならいま検査中のロット。"""
    if kensa is None:
        from . import session
        kensa = session.get()
    st = kensa.state
    lot_no = (lot_no or st.lot).strip().upper()
    out: dict[str, Any] = {"lot": lot_no, "current": bool(lot_no) and lot_no == st.lot, "kensa_lot": st.lot,
                           "found": False, "header": {}, "cards": [], "messages": [], "kensa": []}
    if lot_no == "":
        return out
    fac = facility if facility is not None else terminal.facility()
    current = out["current"]
    lv = read_lot(lot_no, dbflag=st.dbflag if current else "", facility=fac,
                  ken_no=st.ken_no_caption if current else "")
    out["messages"] = [m.to_dict() for m in lv.messages]
    r = lv.reflect
    if r is None or r.status != "ok":
        out["status"] = r.status if r else "error"
        return out
    out["found"] = True
    row = lv.row
    flag_atu = st.atu if current and st.atu else lv.form("ATU")      # UF品質基準.FlagAtu = UF1.ATU
    out["header"] = {
        "lot": lot_no, "order": row("ｵｰﾀﾞｰ番号"), "ken_no": lv.ken_no,
        "zai": lv.form("Zai"), "tyu": lv.form("Tyu"), "youto": lv.form("Youto"), "youto_name": row("用途名"),
        "size": " × ".join(x for x in (lv.form("ATU"), lv.form("Hab"), lv.form("Tak")) if x),
        "order_size": " × ".join(x for x in (row("ｵｰﾀﾞｰ板厚"), row("ｵｰﾀﾞｰ板幅"), row("ｵｰﾀﾞｰ板丈")) if x),
        "nou": row("納入先名"), "tori": row("取引先名"), "okuri": row("送り先名"),
        "grade": row("品質ｸﾞﾚｰﾄﾞ_表面処理"), "flat": lv.form("flatnes"), "taik": lv.form("TAIK"),
        "lateral": lv.form("Lateral"), "equipment": lv.equipment, "equipment_ok": lv.equipment in ("可", ""),
        "additional": r.additional,
        "line": lines.vba_name(fac) if fac else "",
    }
    out["cards"] = [c.to_dict() for c in (
        card_distortion(lv), card_appearance(lv, flag_atu), card_bh(lv), card_diagonal(lv), card_lateral(lv),
        card_ratefuri(kensa.store if store is None else store, lv), card_distortion_ref(lv))]
    if current:
        out["kensa"] = [list(x) for x in _kensa_values(st)]
    return out
