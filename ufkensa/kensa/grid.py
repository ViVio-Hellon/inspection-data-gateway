"""検査画面(UF1)の状態モデル: 測定グリッドと画面の状態

VBA では UF1 のコントロール(Label の Caption・BackColor・Visible・Tag など)が状態そのものだった。
ここではそれを明示的なデータにする。**判定・チェック・保存・帳票の各担当もこのモデルを使う。**

- 1 つのコントロール = `Ctl`(値・背景色の状態・枠色・表示・有効・下線・Tag)
- 測定グリッド = `KensaState.cells`(位置 S/U/C/D/E × 行 1〜14、F は 1〜10。名前は VBA と同じ "S1"〜"E14"・"F1"〜"F10")
- 位置ごとの付随欄(対角・直角度・ラテラルボー・NG 表示・歪・BH・不良枚数)は位置文字 "S".."F" をキーにした dict
- 判定は `Ctl.color` に持つ(VBA と同じく **1 つの背景色の枠**。OK/NG の後に入力先の黄色を塗ると NG は消える。
  これは VBA と同じ結果にするため。判定状態だけを見たいときは `Ctl.judge`)
- 名前で引くとき(VBA の `UF1(st)`、`INP.Tag` に入っている名前)は `KensaState.ctl(name)`

色の値(`Ctl.color`):
    ""        既定色 &H80000005(ウィンドウの背景)
    "OK"      13434828 RGB(204,255,204) 薄い緑
    "NG"      8421631  RGB(255,128,128) コーラル
    "YELLOW"  RGB(255,255,0) 次の入力先・フォーカス中
    "WHITE"   RGB(255,255,255)
    "BUTTON"  &H8000000F(ボタン面の色)
    その他は VBA の数値を文字列で("26367" など)
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable, Iterator, Optional

from . import quirks, vbx

# ------------------------------------------------------------------
# 定数
# ------------------------------------------------------------------
#: 検査位置(列)。F は 4 山(6 回検査)のときだけ使う
POSITIONS: tuple[str, ...] = ("S", "U", "C", "D", "E", "F")
#: 位置ごとの行数(S..E は 1〜14、F は 1〜10。F11〜F14 は存在しない)
ROW_COUNT: dict[str, int] = {"S": 14, "U": 14, "C": 14, "D": 14, "E": 14, "F": 10}

#: 作業選択(OptionButton の名前)
WORKS: tuple[str, ...] = ("厚板", "薄板", "少枚数", "極厚", "スラブ")

#: FormRelat:3 `THIN_PLATE_ATU` 薄板判定のしきい値(mm)。これ以下で AtuFlag=True
THIN_PLATE_ATU = 6.0

COLOR_DEFAULT = ""
COLOR_OK = "OK"
COLOR_NG = "NG"
COLOR_YELLOW = "YELLOW"
COLOR_WHITE = "WHITE"
COLOR_BUTTON = "BUTTON"

#: 色の名前 → VBA の BackColor の値
VBA_COLOR: dict[str, int] = {
    COLOR_DEFAULT: -2147483643,   # &H80000005
    COLOR_OK: 13434828,
    COLOR_NG: 8421631,
    COLOR_YELLOW: 65535,          # RGB(255,255,0)
    COLOR_WHITE: 16777215,        # RGB(255,255,255)
    COLOR_BUTTON: -2147483633,    # &H8000000F
}

BORDER_DEFAULT = ""
BORDER_RED = "RED"       # RGB(255,0,0)
BORDER_BLACK = "BLACK"   # RGB(0,0,0)

#: U1(実機確認待ち): Variant の文字列(コントロールの値・遅延束縛の .Caption)と数値の比較。
#: 扱いは `vbx.U1_TEXT_VS_NUMBER_RAISES` 1 か所で決める(既定: 数にならない文字列は「数より大きい」でエラーにしない。
#: 本番で 厚板 の MinMax(S〜E)が動いていること ── U5・D5 は空のまま UnderBar で数値と比べている)。
#: ここは個別に上書きするときだけ使う(None = vbx に従う、True = エラーにしない、False = TypeMismatch)。
VARIANT_STRING_GT_NUMBER: Optional[bool] = None


def _string_gt_number() -> bool:
    if VARIANT_STRING_GT_NUMBER is None:
        return not vbx.U1_TEXT_VS_NUMBER_RAISES
    return VARIANT_STRING_GT_NUMBER


# ------------------------------------------------------------------
# コントロール
# ------------------------------------------------------------------
@dataclass
class Ctl:
    """UF1 の 1 コントロール。

    value     Label.Caption / TextBox.Text / ComboBox.Text(VBA は文字列。数値の代入は CStr で入る)
    color     BackColor の状態(上の色の値)
    border    BorderColor("" / "RED" / "BLACK")
    visible   .Visible
    enabled   .Enabled
    underline .Font.Underline(MinMax の最小・最大の印)
    tag       .Tag(測定セルでは NG_Pass の印。ダブルクリック消去で "" に戻す)
    """
    value: str = ""
    color: str = COLOR_DEFAULT
    border: str = BORDER_DEFAULT
    visible: bool = True
    enabled: bool = True
    underline: bool = False
    tag: str = ""

    @property
    def judge(self) -> str:
        """判定状態("OK" / "NG" / "")。色から読む(VBA は BackColor で判定している)。"""
        return self.color if self.color in (COLOR_OK, COLOR_NG) else ""

    @property
    def is_ng(self) -> bool:
        return self.color == COLOR_NG

    @property
    def is_ok(self) -> bool:
        return self.color == COLOR_OK


Cell = Ctl


@dataclass
class LSelectCache:
    """LSelect のキャッシュ(SaveLoad の Public 変数 g_Cache*、FormRelat の g_SkipLotNotFoundMsg)。

    lot_no     g_CacheLotNo
    ken_no     g_CacheKenNo(検査番号の頭 1 文字)
    cours      g_CacheCours(設計_設備ｺｰｽ)
    ord_atu    g_CacheOrdAtu(ｵｰﾀﾞｰ板厚。Single)
    c_course   g_CacheC_Course(実績_設備ｺｰｽ)
    skip_lot_not_found_msg  g_SkipLotNotFoundMsg(ClearLSelectCache 直後の 1 回だけ「該当ロットがありません」を出さない)
    """
    lot_no: str = ""
    ken_no: str = ""
    cours: str = ""
    ord_atu: float = 0.0
    c_course: str = ""
    skip_lot_not_found_msg: bool = False


def _cells() -> dict[str, Ctl]:
    d: dict[str, Ctl] = {}
    for p in POSITIONS:
        for r in range(1, ROW_COUNT[p] + 1):
            # 設計時: S..E の 1〜10 は表示、11〜14 は非表示。F は全部非表示(G6 UF1)
            d[f"{p}{r}"] = Ctl(visible=(p != "F" and r <= 10))
    return d


def _by_pos(hidden_f: bool = True, visible: bool = True, positions: tuple[str, ...] = POSITIONS) -> dict[str, Ctl]:
    return {p: Ctl(visible=(visible and not (hidden_f and p == "F"))) for p in positions}


#: 設計時に非表示の、グリッド以外のコントロール(G6 UF1)
_HIDDEN_VIEW = (
    "F", "TSF", "TEF", "FMF", "LC6", "LC12", "F_UpVisual", "F_UnderVisual", "OPS6", "OPE6", "CheckTF",
    "L7", "L8", "L10", "L11", "L12", "L13", "YamaLabel", "Loading", "ConductivL",
)
_VIEW_CAPTION = {
    "S": "", "U": "", "C": "", "D": "", "E": "", "F": "",
    "FS": "", "FU": "", "FC": "", "FD": "", "FE": "", "FF": "",
    "AL1": "厚1", "AL2": "厚2", "AL3": "厚3", "AL4": "厚4",
    "L1": "幅1", "L2": "幅2", "L3": "丈1", "L4": "丈2", "L5": "対角A", "L6": "対角B",
    "L9": "対角線差", "L10": "直角度", "L11": "ﾗﾃﾗﾙﾎﾞｰ",
    "LZS": "定盤:", "LZU": "定盤:", "LZC": "定盤:", "LZD": "定盤:", "LZE": "定盤:", "LZF": "定盤:",
    "LGS": "ｹﾞｰｼﾞ:", "LGU": "ｹﾞｰｼﾞ:", "LGC": "ｹﾞｰｼﾞ:", "LGD": "ｹﾞｰｼﾞ:", "LGE": "ｹﾞｰｼﾞ:", "LGF": "ｹﾞｰｼﾞ:",
    "OPS1": "S", "OPS2": "U", "OPS3": "2", "OPS4": "D", "OPS5": "E", "OPS6": "F",
    "OPE1": "S", "OPE2": "U", "OPE3": "2", "OPE4": "D", "OPE5": "E", "OPE6": "F",
}


#: グリッド以外で、表示を切り替えるコントロールの名前(位置ごと)
_VIEW_BY_POS = ("TS{}", "TE{}", "FM{}", "{}_UpVisual", "{}_UnderVisual", "CheckT{}")


def _view() -> dict[str, Ctl]:
    d = {k: Ctl(value=v) for k, v in _VIEW_CAPTION.items()}
    for pat in _VIEW_BY_POS:
        for p in POSITIONS:
            d.setdefault(pat.format(p), Ctl())
    for name in ("L7", "L8", "L12", "L13", "YamaLabel", "Loading", "ConductivL", "ZEB"):
        d.setdefault(name, Ctl())
    for k in _HIDDEN_VIEW:
        d.setdefault(k, Ctl()).visible = False
    return d


def _work_captions() -> dict[str, str]:
    return {"厚板": "厚板", "薄板": "薄板", "少枚数": "少枚数", "極厚": "耳付", "スラブ": "スラブ"}


# ------------------------------------------------------------------
# 状態
# ------------------------------------------------------------------
@dataclass
class KensaState:
    """UF1 の状態(1 端末 = 1 検査中ロット)。各フィールドの右は VBA のコントロール/変数。"""

    # --- 測定グリッド -------------------------------------------------
    #: 測定値ラベル S1..S14, U1..U14, C1..C14, D1..D14, E1..E14, F1..F10(行1〜4=厚、5〜=幅・丈・対角。§3.1)
    cells: dict[str, Ctl] = field(default_factory=_cells)
    #: 対角差 DiagS..DiagF(Label)
    diag: dict[str, Ctl] = field(default_factory=lambda: _by_pos())
    #: 直角度 DiagSKou..DiagFKou(TextBox。設計時は全部非表示)
    diag_kou: dict[str, Ctl] = field(default_factory=lambda: _by_pos(visible=False))
    #: ラテラルボー LateralS..LateralF(TextBox。設計時は全部非表示)
    lateral: dict[str, Ctl] = field(default_factory=lambda: _by_pos(visible=False))
    #: 位置ごとの NG 表示 NGS..NGF(Label。Caption = "NG" / "許容" / "保留" / "")
    ng: dict[str, Ctl] = field(default_factory=lambda: _by_pos())
    #: 歪(定盤 / 幅ｹﾞｰｼﾞ)ZS..ZF(TextBox)。表示中の面(strain_face)の値
    strain_z: dict[str, Ctl] = field(default_factory=lambda: _by_pos(hidden_f=False))
    #: 歪(ゲージ / 丈ｹﾞｰｼﾞ)GS..GF(TextBox)
    strain_g: dict[str, Ctl] = field(default_factory=lambda: _by_pos(hidden_f=False))
    #: 歪コンボ CBS1..CBF4(名前 "CBS1" をキー。CB?1 = 形状("耳"/"中" 等)、CB?2 = 上下、CB?3/CB?4 = ゲージ側)
    strain_cb: dict[str, Ctl] = field(default_factory=lambda: {f"CB{p}{i}": Ctl() for p in POSITIONS for i in range(1, 5)})
    #: A/B 面の切替 AorB.Caption("Now_A" / "Now_B")
    strain_face: str = "Now_A"
    #: 作業保存シートの歪の控え(O2:P7 = A 面、O8:P13 = B 面。ReflectStrain/LosdStrain)。キーは "ZS"/"GS" など
    strain_a: dict[str, str] = field(default_factory=dict)
    strain_b: dict[str, str] = field(default_factory=dict)
    #: BH SBH..EBH(TextBox。FBH は存在しない)
    bh: dict[str, Ctl] = field(default_factory=lambda: _by_pos(positions=("S", "U", "C", "D", "E")))
    #: 外観不良枚数 SBad..FBad(TextBox)
    bad: dict[str, Ctl] = field(default_factory=lambda: _by_pos())
    #: 検入高さ KenHBox1..KenHBox7(TextBox。7 = 少枚数の枚数用)
    ken_h: dict[int, Ctl] = field(default_factory=lambda: {i: Ctl(visible=False) for i in range(1, 8)})
    #: 少枚数の検入枚数 KenCount(TextBox)
    ken_count: str = ""
    #: 外観ボタン LC1..LC12(Label。LC1〜6 / LC7〜12 の 2 段)
    visual: dict[int, Ctl] = field(default_factory=lambda: {i: Ctl(visible=(i not in (6, 12))) for i in range(1, 13)})
    #: そのほかの表示だけのコントロール(列見出し S..F・FS..FF、行見出し AL1..AL4・L1..L13、OPS1..6/OPE1..6 の Caption、
    #: 歪フレーム FMS..FMF、LZ?/LG?、TSS..TEF、?_UpVisual/?_UnderVisual、CheckT?、YamaLabel、Loading など)。
    #: value = Caption
    view: dict[str, Ctl] = field(default_factory=_view)

    # --- ロット・規格(表示欄)---------------------------------------
    lot: str = ""                 # Lot(TextBox 7 桁)
    line_name: str = ""           # LineName.Caption(VBA の画面のライン名 L-1/LVC/HVC/ﾊﾞﾗﾝｻｰ/ﾄｯﾄ/Test/丸徳)
    line_no: str = ""             # LineNo.Caption(丸徳の号機)
    line_option: str = ""         # ライン OptionButton(L1Work/LVC/HVC/Tott/Bala/Test/Maru のどれが True か)
    ken_no_caption: str = ""      # KenNo.Caption(検査番号の全体)
    reien: str = ""               # Reien.Caption(冷延表示 "H_2HI" など。LSelect でだけ設定)
    comment: str = ""             # Comment.Caption
    youto: str = ""               # Youto(用途コード)
    zai: str = ""                 # Zai(材質)
    tyu: str = ""                 # Tyu(調質)
    sumisei: str = ""             # Sumisei.Caption(ラテラルボー要の納入先。旧: 住友精密)
    additional: str = ""          # Additional.Caption(技術指示。"ﾗﾃﾗﾙﾎﾞｰ検査不要" など)
    lateral_text: str = ""        # LateralText(ラテラルボー公差の表示)
    lateral_text_tag: str = ""    # LateralText.Tag("有" = 技術指示経路)
    lateral_text_colored: bool = False  # LateralText.BackColor = RGB(51,204,204) かつ BackStyle = 1(カードに公差あり)
    flatnes_text: str = ""        # flatnesText("ﾌﾘｰｿﾘ要")
    huris: str = ""               # HURIS
    conductiv_text: str = ""      # ConductivText.Caption
    atu: str = ""                 # ATU(厚)
    hab: str = ""                 # Hab(巾)
    tak: str = ""                 # Tak(丈)
    ap: str = ""                  # Ap(厚 +公差)
    am: str = ""                  # Am(厚 -公差)
    hp: str = ""                  # Hp
    hm: str = ""                  # Hm
    tp: str = ""                  # Tp
    tm: str = ""                  # Tm
    apk: str = ""                 # ApK(厚 上限)
    amk: str = ""                 # AmK(厚 下限)
    hpk: str = ""                 # HpK
    hmk: str = ""                 # HmK
    tpk: str = ""                 # TpK
    tmk: str = ""                 # TmK
    taik: str = ""                # TAIK(対角等級)
    lateral_grade: str = ""       # Lateral(ラテラル等級。TextBox "Lateral")
    amin: str = ""                # AMin(Label)
    amax: str = ""                # AMax
    hmin: str = ""                # HMin
    hmax: str = ""                # HMax
    tmin: str = ""                # TMin
    tmax: str = ""                # TMax
    dbflag: str = ""              # DBflag(TextBox。"" 以外で仕掛 DB を NextPath に・校正スキップ)
    # D28: ロットごとの「控え(第2)から読む」(DBflag の代わり。校正・Min/Max には効かない)
    lot_source: str = ""          # "" = いつもの読み方(モードどおり) / "second" = 控え(第2)だけから読む
    lot_source_lot: str = ""      # その選択をしたロット(このロットを読むときだけ効く)

    # --- 作業選択・検査回数 -------------------------------------------
    work: str = ""                # 厚板/薄板/少枚数/極厚/スラブ の OptionButton のうち True のもの("" = 未選択)
    work_caption: dict[str, str] = field(default_factory=_work_captions)   # 各 OptionButton の Caption(検番の頭が入る)
    work_visible: dict[str, bool] = field(default_factory=lambda: {w: True for w in WORKS})
    byousya: bool = False         # Byousya(強制描写 CheckBox)
    yama: str = ""                # Yama.Text(山数 "1"〜"7")
    yama_visible: bool = False    # Yama.Visible
    yama_tag: str = ""            # Yama.Tag("UF1" / "別フォーム")
    at_butt_caption: str = "厚2ヵ所"   # AtButt.Caption(次に切り替わる先。"厚2ヵ所" = いま 4ヵ所)
    at_butt_visible: bool = False
    at_count: str = ""            # AtCount.Caption("厚4ヵ所検査中" など)
    inspe_butt_caption: str = "検入≦25"  # InspeButt.Caption(次に切り替わる先。"検入≧26" = いま ≦25 = 2 回検査)
    inspe_butt_visible: bool = False
    ops: int = 0                  # OPS1..OPS6 のうち True の番号(0 = なし)。Min/Max の開始位置
    ope: int = 0                  # OPE1..OPE6。終了位置

    # --- 入力の流れ・イベント制御(VBA ではコントロールの Caption/Tag)----------
    load: str = "on"              # Load.Caption("off" でフォームのイベントを止める)
    form_tag: str = ""            # UF1.Tag("STOP" = 音声・MsgBox 停止)/ UFTag
    flag_label: int = 1           # FlagLabel(1 = INP_Change を受け付ける。0 = 処理中)
    auto: str = "on"              # Auto("on" = 次の入力欄を自動で決める。ラベルを手で選ぶと "off")
    inp: str = ""                 # INP(受信した伝送文字列)
    inp_tag: str = ""             # INP.Tag(いまの入力先のコントロール名 "S1" など)
    tagtext: str = ""             # Tagtext.Caption(入力先の表示)
    paked: str = ""               # PakeD(伝送の種別 "DT" / "#L" / "#〆")
    machin: str = ""              # Machin(測定機器名)
    mpex_tag: str = ""            # MPEX.Tag("Click" = 判定表示の抑止)
    picfm_tag: str = "0"          # PicFM.Tag(フォームクリックの回数)
    move_flag: str = ""           # MoveFlag
    proofread_flag: str = ""      # ProofreadFlag("ok" = 校正済み)
    work_button_tag: str = ""     # WorkButton.Tag("jump" = 校正スキップ)
    each_time: str = ""           # EachTime(2 度押し判定の時刻)
    each_value: str = ""          # EachValue(2 度押し判定の値)
    delete_p: str = ""            # DeleteP.Caption(最後に消した座標)
    delete_c: str = ""            # DeleteC.Caption(同じ座標を続けて消した回数)
    abno_tag: str = ""            # Abno.Tag("NG" = 異常値で検査停止)
    resize_tag: str = ""          # ReSize.Tag("有" = 公差の手修正中)
    blindfold_tag: str = ""       # blindfold.Tag("Set" = 歪コンボ追加済み)
    sfm_tag: str = ""             # SFM.Tag("Stop" = Min/Max の範囲が逆)
    blind_message: str = ""       # 目隠しの文言(UfText。"少々お待ちください" / "作業選択を選んでください")

    # --- モジュール変数 -------------------------------------------------
    cache: LSelectCache = field(default_factory=LSelectCache)
    sika_set_atu: float = 0.0     # SaveLoad の Public SikaSetAtu(Single)
    sika_atu_all: Optional[list[Any]] = None   # SaveLoad の Public SikaAtuAll

    # ------------------------------------------------------------------
    @property
    def lname(self) -> str:
        """VBA の `Lname = LineName.Caption & LineNo.Caption`。"""
        return self.line_name + self.line_no

    def is_work(self, name: str) -> bool:
        """`UF1.厚板` などの OptionButton の値。"""
        return self.work == name

    def ctl(self, name: str) -> Ctl:
        """VBA の `UF1(name)`。名前からコントロールを引く(無い名前は KeyError = VBA の実行時エラー)。"""
        if name in self.cells:
            return self.cells[name]
        if len(name) == 8 and name.startswith("Diag") and name.endswith("Kou") and name[4] in self.diag_kou:
            return self.diag_kou[name[4]]
        for prefix, group in (("Diag", self.diag), ("Lateral", self.lateral), ("NG", self.ng)):
            if name.startswith(prefix) and name[len(prefix):] in group:
                return group[name[len(prefix):]]
        if len(name) == 2 and name[0] == "Z" and name[1] in self.strain_z:
            return self.strain_z[name[1]]
        if len(name) == 2 and name[0] == "G" and name[1] in self.strain_g:
            return self.strain_g[name[1]]
        if name in self.strain_cb:
            return self.strain_cb[name]
        if len(name) == 3 and name.endswith("BH") and name[0] in self.bh:
            return self.bh[name[0]]
        if len(name) == 4 and name.endswith("Bad") and name[0] in self.bad:
            return self.bad[name[0]]
        if name.startswith("KenHBox") and name[7:].isdigit() and int(name[7:]) in self.ken_h:
            return self.ken_h[int(name[7:])]
        if name.startswith("LC") and name[2:].isdigit() and int(name[2:]) in self.visual:
            return self.visual[int(name[2:])]
        if name in self.view:
            return self.view[name]
        raise KeyError(f"UF1 にコントロール {name!r} がありません")

    def has_ctl(self, name: str) -> bool:
        try:
            self.ctl(name)
            return True
        except KeyError:
            return False

    def v(self, name: str) -> Ctl:
        """表示だけのコントロール(view)。無ければ既定値で作る。"""
        if name not in self.view:
            self.view[name] = Ctl()
        return self.view[name]

    def iter_cells(self) -> Iterator[tuple[str, Ctl]]:
        for p in POSITIONS:
            for r in range(1, ROW_COUNT[p] + 1):
                yield f"{p}{r}", self.cells[f"{p}{r}"]

    def set_header(self, pos: str, caption: str) -> None:
        """列見出し(Label "S".."F")の Caption。"""
        self.v(pos).value = caption

    def header(self, pos: str) -> str:
        return self.v(pos).value


# ------------------------------------------------------------------
# VBA の比較(コントロールの値は Variant)
# ------------------------------------------------------------------
def variant_cmp(a: Any, b: Any) -> int:
    """Variant(文字列)と数値の比較。VBA の `UF1.Hab < 300#`・`UF1.Yama = 3`・`UF1("S1").Caption = Min` に当たる。

    数にできる文字列は数で比べる。できない文字列(空文字を含む)は U1 のスイッチ(`vbx.U1_TEXT_VS_NUMBER_RAISES`)が
    偽なら「文字列のほうが大きい」(U1 実機確認待ち)、真なら vbx.vcompare と同じく TypeMismatch。
    """
    if _string_gt_number():
        num_a = isinstance(a, (int, float)) and not isinstance(a, bool)
        num_b = isinstance(b, (int, float)) and not isinstance(b, bool)
        if num_a and isinstance(b, str) and not vbx.is_numeric(b):
            return -1
        if num_b and isinstance(a, str) and not vbx.is_numeric(a):
            return 1
    return vbx.vcompare(a, b)


def variant_eq(a: Any, b: Any) -> bool:
    return variant_cmp(a, b) == 0


def typed_cmp(a: Any, b: Any) -> int:
    """型付きの String(`Label.Caption` を早期束縛で読んだもの・`.Text`)と数値の比較。数にならなければ TypeMismatch。"""
    return vbx.vcompare(a, b)


def yama_is(state: KensaState, n: int) -> bool:
    """`UF1.Yama = n`(ComboBox の Value は Variant)。空は偽(U1)。"""
    return variant_eq(state.yama, n)


def hab_lt_300(state: KensaState) -> bool:
    """`UF1.Hab < 300#`(TextBox の Value は Variant)。"""
    return variant_cmp(state.hab, 300.0) < 0


def ken_flags(state: KensaState, ken_no: str, *, blank_in_flag1: bool = False) -> dict[int, bool]:
    """検番フラグ(AllCheck:587-592 / PucketMove:739-744 / AtCountSelect:1989-1993 / InspectionSelect:2222-2226)。

    1: 冷延系(L/N/K/T/S/F)  2: 検番 H  3: 山数 3  4: flag1 かつ InspeButt="検入≧26"(25 枚以下)
    5: 検番なし  6: 山数 4
    `blank_in_flag1=True` は Look_S_Blank(1063-1065)の定義: flag1 に検番 "" を含め、flag5 は立てない。
    """
    f = {i: False for i in range(1, 7)}
    if ken_no in ("L", "N", "K", "T", "S", "F") or (blank_in_flag1 and ken_no == ""):
        f[1] = True
    if ken_no == "H":
        f[2] = True
    if yama_is(state, 3):
        f[3] = True
    if f[1] and state.inspe_butt_caption == "検入≧26":
        f[4] = True
    if not blank_in_flag1 and ken_no == "":
        f[5] = True
    if yama_is(state, 4):
        f[6] = True
    return f


# ------------------------------------------------------------------
# 共通の小さな処理
# ------------------------------------------------------------------
def remove_underline(state: KensaState) -> None:
    """FormRelat.RemoveUnderlineFromLabels (FormRelat:5629-5655)。S..E 1〜14、F 1〜10 の下線を外す。"""
    for _, c in state.iter_cells():
        c.underline = False


def fm_clear(state: KensaState, n: str) -> None:
    """FormRelat.Fm_Clear (FormRelat:1620-1829)。

    データに効くのは、Z? が数なら "0.00" に整えることと、最後の下線解除。
    フレームの幅・色(強調表示)は、`n` 以外の位置のフレームを既定に戻す(表示だけ)。
    """
    for p in POSITIONS:
        z = state.strain_z[p]
        if vbx.is_numeric(z.value):
            z.value = vbx.fmt(z.value, "0.00")
    for p in POSITIONS:
        if n in POSITIONS and p == n:
            continue
        for name in (f"FM{p}", f"LZ{p}", f"LG{p}"):
            state.v(name).color = COLOR_BUTTON
    remove_underline(state)


def n_count(state: KensaState) -> bool:
    """FormRelat.nCount (FormRelat:5423-5453)。測定値が 1 つでもあれば真(S..E は 1〜12 行、F は 1〜10)。

    B88: VBA は S..E の行 1〜12 しか見ない。直す: 極厚の丈3・4(行 13・14)も見る。
    """
    last = 12 if quirks.reproduce("B88") else 14
    for i in range(1, last + 1):
        for p in ("S", "U", "C", "D", "E"):
            if state.cells[f"{p}{i}"].value != "":
                return True
    for i in range(1, 11):
        if state.cells[f"F{i}"].value != "":
            return True
    return False


def small_count_th(state: KensaState, calculate_sheets: Callable[[KensaState, float], float], *,
                   start_from_ken_count: bool = False, by_ken_count_blank: bool = False) -> int:
    """少枚数の枚数 th(AllCheck:434-441、MinMax:5705-5712、UF1 735/757/1700/4397/5251)。

        If KenCount <> "" Then th = KenCount            ' AllCheck だけ(start_from_ken_count)
        If ATU <> "" Then
          If KenHBox7 = "" Then th = Format(Round(CalculateSheets(UF1, ATU), 1), "0")   ' UserForm_Click は KenCount = ""
          ElseIf KenCount <> "" Then th = KenCount
          If th >= 5 Then th = 5

    `calculate_sheets` は OutTolerance.CalculateSheets(Single を返す)。th は Long(代入は銀行丸め)。
    """
    th = 0
    if start_from_ken_count and state.ken_count != "":
        th = vbx.clng(state.ken_count)
    if state.atu != "":
        first = state.ken_count if by_ken_count_blank else state.ken_h[7].value
        if first == "":
            atu = vbx.single(vbx.cdbl(state.atu))
            th = vbx.clng(vbx.fmt(vbx.vround(calculate_sheets(state, atu), 1), "0"))
        elif state.ken_count != "":
            th = vbx.clng(state.ken_count)
        if th >= 5:
            th = 5
    return th


# ------------------------------------------------------------------
# 他の担当の処理(呼び出し口)
# ------------------------------------------------------------------
class Hooks:
    """画面の流れの中で呼ぶ、ほかの担当の VBA 手続き。

    値を返すもの(判定に効くもの)は既定で NotImplementedError。画面側・試験側で上書きする。
    値を返さないもの(判定・保存・ログ)は既定で何もしない。
    """

    # --- OutTolerance / Judg / SaveLoad / Worker ほか(値を返す) ----------
    def calculate_sheets(self, state: KensaState, atu: float) -> float:
        """OutTolerance.CalculateSheets(UF1, ATU)(検入高さ → 枚数。Single)。"""
        raise NotImplementedError("CalculateSheets")

    def register_check(self, state: KensaState) -> bool:
        """Worker.Register_Check(Format$(Now,"hh:mm:ss"), "")。真 = 検査員の登録が無い。"""
        raise NotImplementedError("Register_Check")

    def is_over_12_hours(self, state: KensaState) -> bool:
        """Worker.IsOver12Hours(B19)。"""
        raise NotImplementedError("IsOver12Hours")

    def get_lot_info(self, state: KensaState, lot: str) -> Optional[list[list[str]]]:
        """AccHIki.GetLotInfo("Lot情報履歴", Lot)。列 1..10 = 納入先,得意先,用途ｺｰﾄﾞ,材質,調質,厚,巾,丈,ﾛｯﾄNo,検査No。0 件は None。"""
        raise NotImplementedError("GetLotInfo")

    def db_return_item(self, state: KensaState, delivery: str, customer: str, youto: str, zai: str, tyu: str,
                       atu: float, hab: float, tak: float) -> str:
        """QP_DB.DbReturnItem(品質基準の照合結果。"ラテラルボー"/"直角度"/"フリー反り" を含む文字列)。"""
        raise NotImplementedError("DbReturnItem")

    def is_gauge_required(self, atu: float, hab: float, tak: float, zai: str, tyu: str, cours: str) -> bool:
        """QP_DB.IsGaugeRequired(ゲージ判定条件)。"""
        raise NotImplementedError("IsGaugeRequired")

    def sika_atu_create(self, state: KensaState, lot: str) -> Optional[list[Any]]:
        """SaveLoad.SikaAtuCreate(LotNo)。"""
        raise NotImplementedError("SikaAtuCreate")

    def arrbox(self, state: KensaState, target: str, sql_filter: str = "", use: str = "Record") -> Optional[list[Any]]:
        """AccHIki.Arrbox(target, filter, use, DBflag)。"Fields" は列名の list、"Record" は行(list[list[str]])の list。

        既定は datasource で直接読む簡易版(MainPath/NextPath の切替・TimeDesignation は lot.py の担当)。
        """
        from . import datasource, paths
        db = target[:-6] if target.endswith(".accdb") else target
        path = paths.db_file(db) if state.dbflag == "" else paths.db_file(db, base_key="next")
        if not datasource.exists(path):
            return None
        if sql_filter == "":
            flt = ""
        elif sql_filter == "板":
            flt = "[ｵｰﾀﾞｰ板丈]<>0"
        elif sql_filter == "コイル":
            flt = "[ｵｰﾀﾞｰ板丈]=0  AND  [BOX設計_設備名] = 'KEN' "
        else:
            flt = "[ﾛｯﾄ番号]='" + sql_filter + "'"
        if use == "Fields":
            return datasource.get_fields(path, "仕掛", "")
        return datasource.get_records(path, "仕掛", flt)

    def atu_hiki_records(self, state: KensaState, lot: str) -> tuple[bool, Optional[list[Any]]]:
        """Atu_OPselect の SIKAHIKINOW の読み (FormRelat:6227-6247)。Arrbox を通さず GetRecordsArr を直接呼ぶ。

        DBflag が空なら MainPath、MainPath が読めなければ(ReadCheck・TimeDesignation・FileDesignation)NextPath。
        NextPath も無いときは (False, None)(呼び手が「日稲:…がありません_ReadCheckAtu_OPselect」で抜ける)。
        Filter は Lot が空でも `[ﾛｯﾄ番号]=''`(Arrbox の「空なら絞り込み無し」とは違う)。
        既定は ReadCheck だけを見る簡易版(時刻窓・PLEASE WAIT は wiring.RealHooks)。
        """
        from . import ado_filter, datasource, paths
        main, nxt = paths.db_file("SIKAHIKINOW"), paths.db_file("SIKAHIKINOW", base_key="next")
        path = main if state.dbflag == "" else nxt
        if state.dbflag == "" and not datasource.exists(main):
            path = nxt
            if not datasource.exists(nxt):
                return False, None
        if not datasource.exists(path):
            return True, None
        return True, datasource.get_records(path, "仕掛", "[ﾛｯﾄ番号]='" + ado_filter.quote(lot) + "'")

    # --- 値を返さないもの ---------------------------------------------
    def loop_text(self, state: KensaState, st: str) -> None:
        """FormRelat.LoopText(st)(異常値除去→機器判定→TextJudg→TimeRecord→NumDig→CheckLabelsByCategory)。"""

    def posting(self, state: KensaState) -> None:
        """SaveLoad.Posting(作業保存シートへ書く)。"""

    def out_tolerance(self, state: KensaState) -> None:
        """OutTolerance.Out_Tolerance(公差外チェック)。"""

    def loop_judg(self, state: KensaState, n: float = 0) -> None:
        """Judg.LoopJudg。"""

    def hiz_loop_judg(self, state: KensaState, n: float = 0) -> None:
        """Judg.HizLoopJudg。"""

    def check_labels_by_category(self, state: KensaState, pos: str) -> None:
        """OutTolerance.CheckLabelsByCategory(pos)。"""

    def proofread_load(self, state: KensaState) -> None:
        """Proofreading.Proofread_Load(校正画面を開く)。"""

    def time_posting(self, state: KensaState, st: str, mode: Optional[str] = None) -> None:
        """SaveLoad.Time_Posting(st, Now[, ""])。mode="" は測定時刻のクリア。"""

    def confi_reason_open(self, state: KensaState, pos: str, num: int, lname: str, value: str, count: float) -> None:
        """OutTolerance.ConfiReason_Open(消去理由の入力画面)。"""

    def te_color(self, state: KensaState, name: str) -> None:
        """OutTolerance.TeColor(不良枚数の色)。"""

    def cb_add(self, state: KensaState, sika_set_atu: float) -> None:
        """OutTolerance.CBAdd(歪コンボの選択肢)。"""

    def write_log(self, state: KensaState, text: str) -> None:
        """WriteLog(text, "", HistoryPath & "<ライン>\\<ライン>_Log")。"""

    def initial_setting(self, state: KensaState, n: str = "") -> None:
        """FormRelat.InitialSetting(フォームの大きさ。表示だけ)。"""

    def user_form_click(self, state: KensaState) -> None:
        """UF1.UserForm_Click(再判定と次の入力欄)。"""

    def small_number_count(self, state: KensaState, th: int) -> None:
        """OutTolerance.SmallNumberCount(th)。"""

    def load_text(self, state: KensaState, n: str = "") -> None:
        """SaveLoad.Load_Text(作業保存の読み込み)。"""

    def worker_load(self, state: KensaState) -> None:
        """Worker.WorkerLoad。"""
