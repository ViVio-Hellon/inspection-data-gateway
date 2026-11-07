"""包装仕様書の確認(VBA の PackingNoCheck・UFPack_Open・UF_PackNo フォーム)

ロットを読み込んだあと、`Lot情報履歴` の包装仕様NO と 使用パス(QualityPath)の `包装仕様書` フォルダの画像を
突き合わせ、一致すれば包装仕様書の画像と引当の一覧を出す(モーダル。「確認」を付けないと閉じられない)。
通常検査(UF1.CB_Change)・全ヵ所/H494(UFMitsubishi.ComboBox1_Change)・キャスク(UFCask.ComboBox1_Change)から使う。

| VBA | Python |
|---|---|
| FormRelat.PackingNoCheck (FormRelat:6619-6754) | `packing_no_check` |
| OutTolerance.UFPack_Open (O:520-543) + UF_PackNo.UserForm_Activate (UF_PackNo) | `pack_open` |
| FormRelat.PackSelect (FormRelat:6826-6875) | `pack_spec` |
| LabelCls の Hiki/Ord/Pac のクリック → Loop_Label + FormRelat.JPGLoad (FormRelat:6876-6915) | `pack_select` |
| UF_PackNo.CommandButton2_Click(閉じる。CB1「確認」が無いと MsgBox) | `pack_close` |

画像のパスは `paths.get("quality")`(VBA の QualityPath)から作る。直書きしない。
移植しないもの: CommandButton1(IE で包装仕様書のサイトを開く。D9)、拡大縮小の SpinButton(画面の CSS で行う)。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Optional, Sequence

from . import datasource, paths, quality, vbx
from .flow import ERROR, INFO, WARN, Message
from .lot import read_check

FOLDER = "包装仕様書"
#: UF_PackNo の引当の行は 5 行まで(Hiki1〜5 / Ord1〜5 / Pac1〜5)
MAX_ROWS = 5
NEED_CONFIRM = "確認後に押してください。"          # UF_PackNo.CommandButton2_Click


class PackingError(Exception):
    """VBA の実行時エラー(呼び手のエラー処理へ。ComboBox1_Change なら「DB読み込みエラー:」)。"""

    def __init__(self, number: int, description: str, where: str = "") -> None:
        super().__init__(f"{number}: {description}")
        self.number = number
        self.description = description
        self.where = where


def folder_path() -> str:
    """`QualityPath & "包装仕様書"`(使用パス「包装仕様書(画像)」。決めていなければ VBA と同じ場所)。"""
    if paths._saved_value("pack_spec") is None:
        return paths.join(paths.get("quality"), FOLDER)
    return paths.get("pack_spec").rstrip("\\/")


def picture_path(n: str) -> str:
    """`QualityPath & "包装仕様書\\" & n & ".JPG"`(O:524-525)。"""
    return paths.join(folder_path(), n + ".JPG")


def picture_file(n: str) -> Optional[Path]:
    """`ReadCheck(PackPath & PackPic, 1)` が真なら画像のファイル。Windows と同じく大文字小文字を区別しない。"""
    if n == "":
        return None
    target = picture_path(n)
    if read_check(target, 1):
        return paths.to_local(target)
    folder = paths.to_local(folder_path())
    want = (n + ".JPG").upper()
    try:
        for p in folder.iterdir():
            if p.is_file() and p.name.upper() == want:
                return p
    except OSError:
        return None
    return None


def _eq_str(v: Any, s: str) -> bool:
    """`配列の要素(Variant) = String 変数`(数のセルは CStr して文字で比べる)。"""
    return vbx.compare_typed(v, s, vbx.VAR, vbx.STR) == 0


def _columns(header: Sequence[Any], names: Sequence[str]) -> dict[str, int]:
    """`For i = 1 To UBound(Hist, 2): Select Case Hist(1, i)`(同じ見出しが 2 つなら右が勝つ。無ければ 0)。"""
    out = {n: 0 for n in names}
    for i, v in enumerate(header, start=1):
        if isinstance(v, str) and v in out:
            out[v] = i
    return out


def _at(row: Sequence[Any], c: int) -> Any:
    """`Hist(i, c)`。c = 0 は添字エラー 9(見出しが無い)。"""
    if c < 1:
        raise PackingError(9, "インデックスが有効範囲にありません。", "Hist(i, 0)")
    return row[c - 1] if c - 1 < len(row) else None


def _range_rows(hist: Sequence[Sequence[Any]]) -> list[list[Any]]:
    """`HisSH.Range(Cells(1, 1), Cells(UR(HisSH, 1), UC(HisSH, 1)))`。列は 1 行目の最後の列まで。"""
    if not hist:
        return []
    header = list(hist[0])
    width = 0
    for i, v in enumerate(header, start=1):
        if v is not None and v != "":
            width = i
    width = max(width, 1)
    return [(list(r) + [None] * width)[:width] for r in hist]


def list_files(folder: Path) -> list[str]:
    """`fso.GetFolder(FILE_PATH).Files` の順(NTFS の名前順 = 大文字にして並べた順)。フォルダが無ければエラー 76。"""
    try:
        names = [p.name for p in folder.iterdir() if p.is_file()]
    except OSError as e:
        raise PackingError(76, "パスが見つかりません。", "fso.GetFolder") from e
    return sorted(names, key=lambda s: (s.upper(), s))


def packing_no_check(hist: Optional[Sequence[Sequence[Any]]], lot_no: str, *,
                     messages: Optional[list[Message]] = None,
                     gaichu: Callable[..., str] = quality.gaichu_packing_no,
                     files: Optional[Callable[[Path], list[str]]] = None) -> str:
    """FormRelat.PackingNoCheck (FormRelat:6619-6754)。表示する包装仕様書No(無ければ "")。

    hist: `Lot情報履歴` の 1 行目(見出し)〜A 列の最終行(None はシートが無い)。
    1. 外注条件(QP.GaichuPackingNo)に当たればその No
    2. ロットの行の包装仕様NO のうち、`包装仕様書` フォルダにあるファイル名(拡張子の "." が 7 文字目なら左 6 文字、
       それ以外は名前のまま)と一致するもの。**フォルダの並びで最後に一致したもの**が返る(内側の Exit For だけ抜ける)

    VBA どおりの実行時エラー(`PackingError`): フォルダが無い = 76。Dir で数えたファイルが 0 件 = 9
    (`ReDim PacN(1 To 0)`)。フォルダの最初のファイルの "." が 7 文字目でない = 9(`PacN(0)`)。この2つは B84
    (直した場合はエラーにせず、空のフォルダは ""、名前はそのまま照合する)。
    引当NO での並び替え(merge_sort2)は結果に効かないので行わない。
    """
    msgs = messages if messages is not None else []
    if hist is None or lot_no == "":
        return ""
    rows = _range_rows(hist)
    if not rows:
        return ""
    col = _columns(rows[0], ("ﾛｯﾄNo", "ｵｰﾀﾞｰNo", "引当NO", "包装仕様NO", "材質", "調質", "ｺｰｽ", "厚", "用途ｺｰﾄﾞ"))
    r_lot = col["ﾛｯﾄNo"]
    if r_lot < 1:
        raise PackingError(9, "インデックスが有効範囲にありません。", "GaichuPackingNo Hist(i, R_Lot)")
    gno = gaichu(rows, lot_no, r_lot, msgs)
    if gno != "":
        return gno
    packs = [_at(r, col["包装仕様NO"]) for r in rows if _eq_str(_at(r, r_lot), lot_no)]
    if not packs:
        return ""
    folder = paths.to_local(folder_path())
    # B84: PacN はデバッグ用の配列。直した場合は PacN を作らない(空のフォルダ・"." が 7 文字目でない名前で止まらない)
    fix_b84 = True
    try:
        names = (files or list_files)(folder)
    except PackingError as e:
        if not fix_b84 or e.number != 76:
            raise
        # 直した側: 包装仕様書のフォルダが無くてもロット登録は止めない(画像を出さないだけ)
        msgs.append(Message(f"包装仕様書のフォルダが見つからないので、包装仕様書の画像は出しません: {folder}"
                            "(設定 → 使用パス の「包装仕様書」)", WARN, source="FormRelat.PackingNoCheck"))
        return ""
    file_int = len(names)                       # Dir(FILE_PATH & "\*") の数
    if file_int == 0:
        if fix_b84:
            return ""                           # 一致するファイルが無い
        raise PackingError(9, "インデックスが有効範囲にありません。", "ReDim PacN(1 To 0)")
    result = ""
    int_no = 0
    for name in names:
        pos = name.rfind(".") + 1               # InStrRev(Temp.Name, ".")
        if pos == 7:
            int_no += 1
            stem = vbx.left(name, 6)
        else:
            stem = name
        if not fix_b84 and not 1 <= int_no <= file_int:          # PacN(intNo) = …
            raise PackingError(9, "インデックスが有効範囲にありません。", "PacN(intNo)")
        for p in packs:
            if _eq_str(p, stem):
                result = stem
                break
    return result


# ------------------------------------------------------------------
# UF_PackNo(包装仕様書の画面)
# ------------------------------------------------------------------
@dataclass
class PackForm:
    """UF_PackNo の表示。rows は引当の行(hiki=引当NO, ord=ｵｰﾀﾞｰNo, pac=包装仕様NO の表示文字)。

    selected: 赤字の行(1〜5)。開いた直後はどの行も赤くない(0)。
    picture: 出している画像の No(`<No>.JPG`)。spec: L1〜L5 のキャプション(PackSelect)。
    """
    lot: str
    tb_load: str
    caption: str = ""
    rows: list[dict[str, str]] = field(default_factory=list)
    selected: int = 0
    picture: str = ""
    spec: list[str] = field(default_factory=lambda: [""] * 5)

    def to_dict(self) -> dict[str, Any]:
        return {"lot": self.lot, "tb_load": self.tb_load, "caption": self.caption, "rows": list(self.rows),
                "selected": self.selected, "picture": self.picture, "spec": list(self.spec)}


def _label(v: Any) -> str:
    """ラベルの Caption への代入(Empty は ""、数は CStr)。"""
    return vbx.cstr(v) if v is not None else ""


def _vlt(a: Any, b: Any) -> bool:
    """Variant どうしの `<`(merge_sort2)。数 < 文字列。Empty は相手に合わせて 0 / ""。"""
    num = lambda x: isinstance(x, (int, float)) and not isinstance(x, bool)  # noqa: E731
    if a is None:
        a = 0 if num(b) else ""
    if b is None:
        b = 0 if num(a) else ""
    if num(a) and num(b):
        return float(a) < float(b)
    if num(a) != num(b):
        return num(a)
    return str(a) < str(b)


def merge_sort2(rows: list[list[Any]], col: int) -> list[list[Any]]:
    """補助.merge_sort2 + merge2 (補助:932-1020) をそのまま。col 列(1 始まり)で昇順(同じ値は元の順)。

    比べ方は Variant の `<` / `<=`(`_vlt`)。数と文字が混ざる・Empty があるときも VBA の手順どおりの並びになる。
    """
    n = len(rows)
    if n == 0:
        return []
    key = [r[col - 1] for r in rows]
    irekae: list[Any] = [None] * (n + 1)          # 1 始まり
    indexer: list[int] = [0] * (n + 1)
    for i in range(1, n + 1, 2):
        if i + 1 > n:
            irekae[i], indexer[i] = key[i - 1], i
            break
        if _vlt(key[i], key[i - 1]):
            irekae[i], irekae[i + 1] = key[i], key[i - 1]
            indexer[i], indexer[i + 1] = i + 1, i
        else:
            irekae[i], irekae[i + 1] = key[i - 1], key[i]
            indexer[i], indexer[i + 1] = i, i + 1

    def merge2(st1: int, en1: int, st2: int, en2: int) -> None:
        tmp = {k: irekae[k] for k in range(st1, en2 + 1)}
        tmpi = {k: indexer[k] for k in range(st1, en2 + 1)}
        j, m = st1, st2
        while j < en1 + 1 or m < en2 + 1:
            if m >= en2 + 1:
                irekae[j + m - st2], indexer[j + m - st2] = tmp[j], tmpi[j]
                j += 1
            elif j < en1 + 1 and not _vlt(tmp[m], tmp[j]):        # tmpArr(j) <= tmpArr(n)
                irekae[j + m - st2], indexer[j + m - st2] = tmp[j], tmpi[j]
                j += 1
            else:
                irekae[j + m - st2], indexer[j + m - st2] = tmp[m], tmpi[m]
                m += 1

    i = 1
    en2 = 0
    while i * 2 <= n:
        i *= 2
        k = 0
        while en2 + i - 1 < n:
            k += 1
            st1 = i * 2 * (k - 1) + 1
            en1 = i * 2 * (k - 1) + i - 1 + 1
            st2 = en1 + 1
            en2 = n if st2 + i - 1 >= n else st2 + i - 1
            merge2(st1, en1, st2, en2)
        en2 = 0
    return [rows[indexer[k] - 1] for k in range(1, n + 1)]


def pack_spec(n: str) -> list[str]:
    """FormRelat.PackSelect (FormRelat:6826-6875)。L1〜L5(包装仕様書No・梱包形態・梱包指定・特記1・特記2)。

    伝送用ファイル(PackingPath)の PkgSpecData から。ファイルが無い・列が取れなければ全部 ""。
    該当の行が無いときは L1 だけ入る(Temp(1, …) で型エラー → 呼び手の On Error Resume Next で残りを飛ばす)。
    """
    out = [""] * 5
    db = paths.db_file("伝送用ファイル", base_key="packing")
    if not read_check(db, 1):
        return out
    try:
        temp = datasource.get_records(db, "PkgSpecData", "[包装仕様書№]='" + n.replace("'", "''") + "'")
        fields = datasource.get_fields(db, "PkgSpecData", "")
    except (datasource.DbMissing, OSError):
        return out
    if fields is None:
        return out
    c = {k: datasource.column_index(fields, k) for k in ("梱包形態", "包装仕様書№", "梱包指定別", "タイトル", "用途名")}
    out[0] = n
    if not temp:
        return out
    row = temp[0]
    for i, (key, head) in enumerate((("梱包形態", "梱包形態:"), ("梱包指定別", "梱包指定:"), ("タイトル", "特記1:"),
                                     ("用途名", "特記2:")), start=1):
        if c[key] < 0:
            return out                          # Temp(1, 0): 添字エラー → 以後は飛ばす
        out[i] = head + _label(row[c[key]])
    return out


def _upper_part_messages(ord_no: str, hist: Optional[Sequence[Sequence[Any]]], msgs: list[Message],
                         source: str) -> None:
    """上部緩衝材の MsgBox(UF_PackNo.Activate の最後 / JPGLoad)。"""
    r1 = quality.materials_upper_part(ord_no, hist)
    r2 = quality.materials_upper_part(ord_no, hist, True)
    if r1 != "" and r1 != "一致なし":
        msgs.append(Message(r1, INFO, source=source))
    if r1 == "アングル" and r2 != "" and r2 != "一致なし":
        msgs.append(Message(r2, INFO, source=source))


def pack_open(n: str, pc_lot: str, hist: Optional[Sequence[Sequence[Any]]], *,
              messages: Optional[list[Message]] = None) -> Optional[PackForm]:
    """OutTolerance.UFPack_Open (O:520-543) → UF_PackNo.UserForm_Activate。

    画像(`QualityPath\\包装仕様書\\<n>.JPG`)が無ければフォームは出ない(None)。
    引当の一覧はロットの行の先頭 5 行(6 行目以降は `TempHiki(6, …)` の添字エラーを On Error Resume Next で
    飛ばすので入らない)を引当NO で並べたもの。L1〜L5 は 1 行目の包装仕様NO(PackSelect)。
    """
    msgs = messages if messages is not None else []
    if picture_file(n) is None:
        return None
    form = PackForm(lot=pc_lot, tb_load=n, picture=n)
    if hist is None or n == "" or pc_lot == "":
        return form
    rows = _range_rows(hist)
    if not rows:
        return form
    col = _columns(rows[0], ("ﾛｯﾄNo", "ｵｰﾀﾞｰNo", "引当NO", "包装仕様NO"))
    if col["ﾛｯﾄNo"] < 1:
        return form                              # Hist(i, 0) の添字エラーで a = 0 → Exit Sub
    get = lambda r, k: (r[col[k] - 1] if 0 < col[k] <= len(r) else None)  # noqa: E731
    hits = [r for r in rows if _eq_str(get(r, "ﾛｯﾄNo"), pc_lot)][:MAX_ROWS]
    if not hits:
        return form
    temp = merge_sort2([[get(r, "ｵｰﾀﾞｰNo"), get(r, "包装仕様NO"), get(r, "引当NO"), get(r, "ﾛｯﾄNo")] for r in hits], 3)
    form.caption = _label(temp[0][3])
    form.rows = [{"hiki": _label(t[2]), "ord": _label(t[0]), "pac": _label(t[1])} for t in temp]
    form.spec = pack_spec(form.rows[0]["pac"])
    _upper_part_messages(_label(temp[0][0]), hist, msgs, "UF_PackNo.Activate")
    return form


def pack_select(form: PackForm, i: int, hist: Optional[Sequence[Sequence[Any]]], *,
                messages: Optional[list[Message]] = None) -> PackForm:
    """引当の行をクリック(LabelCls → Loop_Label(i) で i 行目を赤字 → JPGLoad)。

    画像が無ければ赤字だけ変わり、画像・L1〜L5 はそのまま。
    """
    msgs = messages if messages is not None else []
    if not 1 <= i <= MAX_ROWS:
        return form
    form.selected = i
    row = form.rows[i - 1] if i <= len(form.rows) else {"hiki": "", "ord": "", "pac": ""}
    n, a = row["pac"], row["ord"]
    if picture_file(n) is None:
        return form
    form.picture = n
    form.spec = pack_spec(n)
    _upper_part_messages(a, hist, msgs, "FormRelat.JPGLoad")
    return form


def pack_close(confirmed: bool) -> Optional[Message]:
    """UF_PackNo.CommandButton2_Click。確認(CB1)が無ければ MsgBox で閉じない(Message を返す)。"""
    if confirmed:
        return None
    return Message(NEED_CONFIRM, ERROR, source="UF_PackNo.CommandButton2_Click")


__all__ = ["PackingError", "PackForm", "packing_no_check", "pack_open", "pack_select", "pack_spec", "pack_close",
           "picture_file", "picture_path", "folder_path", "merge_sort2", "list_files", "NEED_CONFIRM"]
