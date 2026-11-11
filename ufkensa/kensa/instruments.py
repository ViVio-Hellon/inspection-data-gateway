"""測定機器のチェックと帳票の機器名(VBA 【UFKen_06】MachineTransfer)

| VBA                                              | ここ                    |
|--------------------------------------------------|-------------------------|
| MachineTransfer.Micro_Juge (MT:165-214)          | `micro_juge`            |
| MachineTransfer.Machin_Text (MT:216-274)         | `machin_text`           |
| 補助.Maker (補助:393-406)                         | `maker`(共通化の候補)  |
| MachineTransfer.機器転送 (MT:276-404)             | `machine_transfer`      |
| MachineTransfer.DbReturnMicro (MT:406-501)       | `db_return_micro`       |
| MachineTransfer.CheckThickAndCode (MT:503-551)   | `check_thick_and_code`  |

移植しない: Micro_Juge_旧 (MT:4-163。呼び出し元なし、2024.2.9 に DB 方式へ切替済み)

【ライン名の照合(D13)】`測定機器ID` の `ライン` 列との照合は VBA の ADO Filter `[ライン]='<ライン>'` だが、
D13 により `lines.matches`(文字の幅だけ揃える。`ﾄｯﾄ` = `トット`)で照合する。
機器転送はバランサー・トットで "大板" を渡すので、マスタに "大板" の行が無い限り機器名は空になる(VBA どおり。B53)。
B53 を直した場合は、"大板" で引けなければ実際のライン(バランサー/トット)で引き直す。

【伝送用ファイルが無い・表が空】VBA は `Dim PackF(): PackF = GetFieldsArr(…)` で、Empty を代入すると
型エラー 13 になる。そのため PackingPath へ切り替える処理(MT:415-420)は到達しない(B51。直すと切り替える)。
ファイルが無いときは GetFieldsArr の中で End(→ `datasource.DbMissing`)。
"""
from __future__ import annotations

import unicodedata
from dataclasses import dataclass, replace
from typing import Any, Callable, Optional, Sequence

from . import ado_filter, datasource, flow, lines, match_rules, paths, size_thresholds, vbx
from .flow import WARN
from .lot import SubscriptOutOfRange

TABLE = "測定機器ID"
DENSO_FILE = "伝送用ファイル.sqlite3"
TRM = "測定するﾏｲｸﾛを正しく選定して下さい"


def _instr(s1: str, s2: str) -> int:
    """VBA `InStr(s1, s2)`。s1 が空なら 0、s2 が空なら 1(vbx.instr は InStr("", "") を 1 にする)。"""
    if s1 == "":
        return 0
    return vbx.instr(s1, s2)


def _val_single(x: Any) -> float:
    """`Dim t As Single: t = Val(文字列)`。"""
    return vbx.single(vbx.val(x))


def _csng(x: Any) -> float:
    """`Dim t As Single: t = UF1.ATU`(文字列を Single に。数にならなければ型エラー 13)。"""
    return vbx.single(vbx.cdbl(x))


# ------------------------------------------------------------------
# DbReturnMicro・CheckThickAndCode
# ------------------------------------------------------------------
def check_thick_and_code(lower: Any, upper: Any, youto_cond: Any, atu: float, yout: str) -> bool:
    """MachineTransfer.CheckThickAndCode (MT:503-551)。

    用途条件(カンマ区切り)のどれかに用途が当たり、かつ 厚下 <= 厚 <= 厚上 なら真。
    - `≠X`: 用途が X でなければ当たり / `*X`: 用途に X を含む / `X*`: 用途に X を含む / それ以外は完全一致
    厚下・厚上は Single(`Val`)、厚は `Val(Single)`(有効 7 桁)で比べる。
    B52: Single と Double の比較なので 28.01 などの境界ちょうどで外れる。直した場合は厚下・厚上も Double で比べる。
    """
    lower_t, upper_t = _bounds(lower, upper)
    found = youto_match(youto_cond, yout)
    if found:
        o = vbx.val_single(atu)
        return lower_t <= o and upper_t >= o
    return False


def youto_match(youto_cond: Any, yout: str) -> bool:
    """CheckThickAndCode の用途の部分(MT:509-545)。照合ルールの 含む条件式 もこれを呼ぶ。"""
    tmp = "" if youto_cond is None else str(youto_cond)
    items = tmp.split(",") if tmp != "" else []      # Split("", ",") は要素 0 個
    for item in items:
        if vbx.left(item, 1) == "≠":
            if vbx.mid(item, 2) != yout:
                return True
        elif vbx.left(item, 1) == "*":
            if _instr(yout, vbx.mid(item, 2)) > 0:
                return True
        elif vbx.right(item, 1) == "*":
            if _instr(yout, vbx.left(item, len(item) - 1)) > 0:
                return True
        else:
            if item == yout:
                return True
    return False


def _bounds(lower: Any, upper: Any) -> tuple[float, float]:
    """厚下・厚上(MT:505-506・489-490)。

    再現: `Dim LowerT As Single: LowerT = Val(…)`。厚は `Val(ATU)`(Double)なので、Single(28.01) > 28.01 で
    境界ちょうどの厚が外れる(B52)。直した場合は `Val` のまま(Double)にして厚と同じ精度で比べる。
    """
    return vbx.val(lower), vbx.val(upper)


def _master() -> Optional[tuple[list[str], list[list[str]], str]]:
    """`測定機器ID` の列名と全行(QualityPath の伝送用ファイル)。None は B51 を直した場合に両方とも空。"""
    file_path = paths.join(paths.get("quality"), DENSO_FILE)
    fields = datasource.get_fields(file_path, TABLE, "")
    if fields is None:
        # B51 を直す: PackingPath の伝送用ファイルへ切り替える(MT:415-420)。それも空なら None(Exit Function)
        file_path = paths.join(paths.get("packing"), DENSO_FILE)
        fields = datasource.get_fields(file_path, TABLE, "")
        if fields is None:
            return None
    _, rows = datasource.get_table(file_path, TABLE, "")
    return fields, rows, file_path


def db_return_micro(line_name: str, atu: float, yout: str, n: str = "", dep: str = "") -> str:
    """MachineTransfer.DbReturnMicro (MT:406-501)。`測定機器ID` から機器の出力テキスト(または機器名・ID)。

    1. ラインの行が無ければ ""
    2. `dep`(M1, N5, デプス など)があれば、ラインと機器名の**先頭行**の出力テキスト(`n` は見ない)
    3. 用途列に値のある行が 1 つでもあれば、用途条件と厚範囲に合う最初の行
    4. 無ければ厚範囲だけで最初の行。それも無ければ ""
    5. n="機器名"/"ID" ならその列
    `atu` は Single(VBA の引数の型)。
    """
    master = _master()
    if master is None:
        return ""                    # B51 を直した場合: どちらの伝送用ファイルも空(Exit Function で "")
    fields, rows, _ = master
    c = {k: datasource.column_index(fields, k)
         for k in ("ライン", "機器名", "ID", "出力テキスト", "厚下", "厚上", "用途")}

    def col(name: str) -> int:
        if c[name] < 0:
            raise SubscriptOutOfRange(f"測定機器ID に {name} の列が無い")
        return c[name]

    atu = vbx.single(atu)
    g = match_rules.group(match_rules.MICRO)
    if g.new_way:
        temp_i, hit = _micro_new(g, fields, rows, line_name, atu, yout, dep == "")
        if match_rules.shadow_on(g):
            match_rules.compare(match_rules.MICRO, _micro_old(fields, rows, col, line_name, atu, yout, dep == ""),
                                (temp_i, hit), repr((line_name, atu, yout, dep)))
    else:
        temp_i, hit = _micro_old(fields, rows, col, line_name, atu, yout, dep == "")
    temp = [rows[i] for i in temp_i]
    if not temp:
        return ""
    if dep != "":
        f = ado_filter.compile_filter("[機器名]= '" + ado_filter.quote(dep) + "'")
        hits = [r for r in temp if f({"機器名": r[col("機器名")]})]
        if not hits:
            return ""
        return hits[0][col("出力テキスト")]
    if hit is None:
        return ""
    temp, itti = rows, hit
    out = temp[itti][col("出力テキスト")]
    if n == "機器名":
        out = temp[itti][col("機器名")]
    elif n == "ID":
        out = temp[itti][col("ID")]
    return out


def _micro_old(fields: Sequence[str], rows: Sequence[Sequence[Any]], col: Callable[[str], int], line_name: str,
               atu: float, yout: str, want_hit: bool = True) -> tuple[list[int], Optional[int]]:
    """旧: DbReturnMicro の照合(MT:431-495)をコードで。(ラインの行の位置, 合った行の位置)。
    照合ルールの shadow で差が 0 件と確かめたら消す。"""
    temp = [i for i, r in enumerate(rows) if lines.matches(r[col("ライン")], line_name)]
    if not temp or not want_hit:
        return temp, None
    flag = any(rows[i][col("用途")] != "" for i in temp)
    if flag:
        for i in temp:
            r = rows[i]
            if check_thick_and_code(r[col("厚下")], r[col("厚上")], r[col("用途")], atu, yout):
                return temp, i
    o = vbx.val_single(atu)
    for i in temp:
        lower_t, upper_t = _bounds(rows[i][col("厚下")], rows[i][col("厚上")])     # B52
        if lower_t <= o and upper_t >= o:
            return temp, i
    return temp, None


def _micro_new(g: "match_rules.Group", fields: Sequence[str], rows: Sequence[Sequence[Any]], line_name: str,
               atu: float, yout: str, want_hit: bool = True) -> tuple[list[int], Optional[int]]:
    """新: 照合ルール「測定機器ID」で引く。戻りは旧と同じ形。

    1. 先に外すルール(ライン)で絞る。2. 含む条件式 のセルがどれかの行にあれば全部のルールで最初の行。
    3. 無ければ 含む条件式 のルールを外して最初の行。
    """
    from . import quality as q
    for r in g.rules:
        for c in r.columns:
            if r.enabled and datasource.column_index(fields, c) < 0:
                raise SubscriptOutOfRange(f"測定機器ID に {c} の列が無い")
    pre, rest = match_rules.split_pre(g)
    ctx = {"ライン": line_name, "厚": atu, "用途": yout}
    temp = match_rules.match(pre, rows, fields, ctx).matched
    if not temp or not want_hit:
        return temp, None
    sub = [rows[i] for i in temp]
    youto_rules = [r for r in rest.rules if r.enabled and r.compare == "含む条件式"]
    flag = any(q.vstr(q.cell(r, q.col_of(fields, x.columns[0]))) != "" for x in youto_rules for r in sub)
    if flag:
        res = match_rules.match(rest, sub, fields, ctx)
        if res.rows:
            return temp, temp[res.rows[0]]
    res = match_rules.match(replace(rest, rules=tuple(r for r in rest.rules if r.compare != "含む条件式")),
                            sub, fields, ctx)
    return temp, (temp[res.rows[0]] if res.rows else None)


# ------------------------------------------------------------------
# Maker・Machin_Text
# ------------------------------------------------------------------
def maker(trm: str) -> str:
    """補助.Maker (補助:393-406)。伝送文字列の頭 2 文字(大文字・半角にして)→ "DT" / "#L" / "#〆" / ""。"""
    if trm == "":
        return ""
    if len(trm) >= 3:
        head = unicodedata.normalize("NFKC", vbx.left(trm, 2).upper())
        if head in ("DT", "#L", "#〆"):
            return head
    return ""


def machin_text(trm: str, current_machin: str) -> str:
    """MachineTransfer.Machin_Text (MT:216-226)。"DT10002+0000021.712M" → "10002"(3 文字目から 5 文字)。
    DT 以外は今の UF1.Machin のまま。"""
    if maker(trm) == "DT":
        return vbx.mid(trm, 3, 5)
    return current_machin


# ------------------------------------------------------------------
# Micro_Juge
# ------------------------------------------------------------------
@dataclass
class MicroInput:
    """Micro_Juge が見る画面の値。"""
    st: str                         # 入力した座標名(S1, E12 など)
    machin: str                     # UF1.Machin.Caption(使用機器の出力テキスト)
    atu: str                        # UF1.ATU
    hab: str                        # UF1.Hab
    youto: str                      # UF1.Youto
    lname: str                      # UF1.LineName.Caption & UF1.LineNo.Caption
    l1work: Optional[bool] = None   # UF1.L1Work(None なら lname から判断)


@dataclass
class MicroJudge:
    """Micro_Juge の結果。`stop` があれば、`clear` の座標を "" にしてから `raise stop`
    (MsgError(…, Lname, 1, "END")。MsgError のログ書きは MsgError の担当)。"""
    ok: bool = True
    clear: str = ""
    stop: Optional[flow.Stop] = None


def micro_juge(m: MicroInput, *, db: Callable[..., str] = db_return_micro) -> MicroJudge:
    """MachineTransfer.Micro_Juge (MT:165-214)。選んだ機器が DB の正しい機器か。

    - UF1.Machin が空なら何もしない
    - 厚・幅は Single へ代入(空・数にならなければ型エラー 13 ── VBA と同じく最初に起きる)
    - それ以外: 用途と厚から引いた出力テキストと違えば NG
    """
    if m.machin == "":
        return MicroJudge()
    pran = m.machin
    atu = _csng(m.atu)
    _csng(m.hab)                    # 幅も Single へ代入する(数にならなければ型エラー 13。VBA どおり)
    use = m.youto
    lname = m.lname
    positive = db(lname, atu, use)
    mdevice = db(lname, atu, use, "機器名")

    if not (pran == positive):
        text = TRM + " " + mdevice + "を使用の事"
        return MicroJudge(False, m.st, flow.Stop(text, source="MachineTransfer:211", level=WARN))
    return MicroJudge()


# ------------------------------------------------------------------
# 機器転送
# ------------------------------------------------------------------
#: ライン(UF1 の OptionButton)→ (帳票のライン名, コンベックス)。L-1 のコンベックスは丈で決める
_LINE_CONB = {"HVC": ("HVC", " C11"), "LVC": ("LVC", " C10"), "バランサー": ("大板", " C12"), "トット": ("大板", " C11")}


@dataclass
class TransferInput:
    """機器転送の引数と、VBA が直接見ていた UF1 の値。"""
    x: int                          # 書く行
    y: int                          # 書く列
    atu: str                        # 引数 ATU(厚)
    haba: str                       # 引数 HABA(幅)
    yout: str = ""                  # 引数 Yout(用途)
    line: str = ""                  # 選ばれているライン(UF1.L1Work/HVC/LVC/Bala/Tott。どれでもなければ "")
    tak: str = ""                   # UF1.Tak.Text
    hab: str = ""                   # UF1.Hab
    cells: Callable[[int, int], Any] = lambda r, c: None   # 帳票シートのセル値(14〜18 行 × B/D/F/H 列を見る)


def machine_transfer(t: TransferInput, *, db: Callable[..., str] = db_return_micro) -> list[tuple[int, int, str]]:
    """MachineTransfer.機器転送 (MT:276-404)。帳票に書くセル `(行, 列, 値)` を書く順に返す。

    - (x, y) = ライン & " " & 機器名、(x, y+2) = ライン & コンベックス
    - L-1 のコンベックス: 丈 > 5500 → " C11"、それ以外 " C12"。HVC " C11"、LVC " C10"、
      バランサー " C12"・トット " C11"(ライン名は "大板")。それ以外のライン(丸徳・Test)は "" と ""
    - 厚がちょうど 300: シートの 14〜18 行 B/D/F/H 列の数に < 300 があれば N5、>= 300 があれば N6
    """
    ln = lines.find(t.line)
    official = ln.official if ln else ""
    line, conb = "", ""
    if official == "L-1":
        line = "L-1"
        conb_take = size_thresholds.get("conb_take")    # 保守 → 寸法のしきい値(既定 5500)
        if vbx.val(t.tak) > conb_take:
            conb = " C11"
        if vbx.val(t.tak) <= conb_take:
            conb = " C12"
    elif official in _LINE_CONB:
        line, conb = _LINE_CONB[official]

    out: list[tuple[int, int, str]] = []
    atu_val = _val_single(t.atu)
    mdevice = db(line, atu_val, t.yout, "機器名")
    if mdevice == "" and line == "大板":
        # B53 を直す: マスタに「大板」の行が無いので、帳票のライン名は「大板」のまま、
        # 機器は実際のライン(バランサー/トット)の行から引く。マスタに大板の行ができればそちらが先
        mdevice = db(official, atu_val, t.yout, "機器名")
    out.append((t.x, t.y, line + " " + mdevice))
    out.append((t.x, t.y + 2, line + conb))

    if vbx.val(t.atu) == 300.0:
        flag = koflag = False
        for i in range(14, 19):
            for a in range(2, 9, 2):
                v = t.cells(i, a)
                if vbx.is_numeric(v) and _ne_empty(v):
                    if vbx.val(v) < 300.0:
                        flag = True
                    elif vbx.val(v) >= 300.0:
                        koflag = True
        if flag and koflag:
            out.append((t.x, t.y, line + " N5,N6"))
        elif flag and not koflag:
            out.append((t.x, t.y, line + " N5"))
        elif koflag and not flag:
            out.append((t.x, t.y, line + " N6"))

    return out


def _ne_empty(v: Any) -> bool:
    """セル値 `<> ""`(Empty は ""、数は "" と等しくない)。"""
    if v is None:
        return False
    if isinstance(v, str):
        return v != ""
    return True


def apply_cells(writes: Sequence[tuple[int, int, str]]) -> dict[tuple[int, int], str]:
    """書く順に重ねた最終の値(同じセルは後勝ち)。"""
    out: dict[tuple[int, int], str] = {}
    for r, c, v in writes:
        out[(r, c)] = v
    return out
