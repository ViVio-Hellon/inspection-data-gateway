# G6 全UserForm コントロール一覧（vbaProject.bin の f / o / x / VBFrame ストリームから抽出）

抽出方法: `scratchpad/tools/formctl.py`（MS-OFORMS 仕様に基づく自作パーサ、olefile 使用）で各フォームストレージの `f`（サイト=コントロール定義: 名前・Tag・ID・ClsidCacheIndex・表示フラグ）と `o`（各コントロールのプロパティ: Caption / Value 等）を読み、Frame / MultiPage / Page はサブストレージ `i<ID>` を再帰的に解析。MultiPage のページ表示名は MultiPage 内部 TabStrip の Items から取得（ページ並び順との対応は **推定**）。フォームのCaptionは `\x03VBFrame` ストリームのテキスト。その後 `scratchpad/tools/formreport.py` で各フォームのコード(`*.frm.bas`)中の `コントロール名_イベント名(` と突き合わせた。

注意:
- 「Caption / 初期値」は**設計時に保存された値**。実行時にコードで書き換えられるもの（例: `UF1.PrintB.Caption = "再印刷"`）は反映されない。
- 表示=○/× は設計時の Visible。
- 番号付きで Caption 空・イベント無し・Tag無しのコントロールが3個以上連続する場合は「A1 〜 A9（9個）」に圧縮。
- `LabelCls` / `clsDynLabel` / `clsBlinkLabel` 経由（WithEvents）で処理されるイベントはフォームコード上に現れない（UF1 の S1〜S14,U1〜U14,C1〜C14,D1〜D14,E1〜E14,F1〜F12 ラベル、CBS1〜CBF4 コンボ、UF_PackNo の Hiki1〜5/Ord1〜5/Pac1〜5、UFAinput の動的ラベル、UF1.AtButt / UFAinput.L1 / UFProgress.CommandButton1 の点滅）。
- ProgressBar / ListView 等の ActiveX（MSComctl）は SiteClassInfo 内 CLSID から判定。判定できないものは `ActiveX#n`。
- 抽出は完璧ではない（解析エラー0件だが、Caption が `None` になる Frame/Page は Caption 未設定または非対応箇所）。


## サマリ

| フォーム | 設計時Caption | コントロール数 | コード内イベント数 | 対応コントロール不在のイベント数 |
|---|---|---|---|---|
| Msg | "Microsoft Excel" | 13 | 5 | 0 |
| SF3 | "初期設定" | 10 | 7 | 0 |
| TEST | "UserForm1" | 100 | 51 | 1 |
| UF1 | "2026.10.14" | 586 | 400 | 2 |
| UFAinput | "2026.06.08" | 147 | 71 | 0 |
| UFCask | "2026.05.29_" | 67 | 29 | 1 |
| UFConfi | "2024.7.26_Ver" | 41 | 10 | 0 |
| UFMitsubishi | "2026.05.29" | 149 | 64 | 1 |
| UFMsgBox | (なし) | 8 | 5 | 0 |
| UFNG | (なし) | 1 | 2 | 0 |
| UFNameList | (なし) | 16 | 14 | 0 |
| UFOK | (なし) | 1 | 2 | 0 |
| UFProgress | (なし) | 6 | 5 | 0 |
| UFWorkerName | (なし) | 4 | 1 | 0 |
| UF_PackNo | (なし) | 35 | 7 | 0 |
| UFcount | "枚数表示" | 10 | 4 | 0 |
| UFdiagonal | "対角チェック" | 18 | 3 | 0 |
| UF品質基準 | "品質一覧" | 131 | 35 | 2 |
| UF対角 | (なし) | 8 | 1 | 0 |
| UF校正 | "品証検証.ｺﾝﾍﾞｯｸｽｶﾒﾗﾁｪｯｸ" | 30 | 16 | 0 |
| UF異常値発見 | "異常" | 11 | 5 | 0 |
| UF管理 | "2024.7.26_Ver" | 200 | 7 | 0 |


## 突き合わせで分かったこと（担当 D）

1. **フォームは 22 個**（PROJECT ストリームの `BaseClass=`）: Msg, SF3, TEST, UF1, UFAinput, UFCask, UFConfi, UFMitsubishi, UFMsgBox, UFNG, UFNameList, UFOK, UFProgress, UFWorkerName, UF_PackNo, UFcount, UFdiagonal, UF品質基準, UF対角, UF校正, UF異常値発見, UF管理。コントロール数は合計 1,592（名前付きのもの）、UF1 だけで 586。
2. **設計時の Caption にバージョン日付が入っているフォーム**: UF1 "2026.10.14"、UFAinput "2026.06.08"、UFCask "2026.05.29_"、UFMitsubishi "2026.05.29"、UFConfi / UF管理 "2024.7.26_Ver"。UF1 は実行時に設備名と「オンライン/オフライン」を後ろに付け足す（`UF1.frm.bas` Initialize）。
3. **UF1 の印刷関連コントロール**:
   - MP1 の「印刷」タブ（Page4）にある `PrintB`「印刷」、`CommandButton18`「保留」、`PrintBreset`「印刷履歴クリア」（設計時は非表示。印刷済みのときにコードで表示する）
   - 「ﾌｫｰﾑ」タブ（Page5）: ZEB「全数.三菱.5N」、KYB「ｷｬｽｸ」、CommandButton7「履歴」、CommandButton8「品質一覧」、H494B「H494」
   - `Conductiv`「導電率ﾌﾗｸﾞ」
   - 不良タブの `Cle`「ｸﾘｱ」、`DefectFactorLB`
   - 非表示の `TestButton`「Test」（印刷転送のテスト）
   - 検査形態の OptionButton: `厚板/薄板/少枚数/極厚(表示Caption「耳付」)/スラブ`
4. **MultiPage のタブ名**（内部 TabStrip の Items から）:
   - UF1.MP1 = 作業選択 / MinMax / 登録済検査員 / 印刷 / ﾌｫｰﾑ / 検査位置 / Log検出 / Test1 / Test2 / Special
   - UF1.MPEX = 検入高さ / 外観 / 歪 / BH / 不良 / ｵｰﾀﾞｰ検索 / Line / 避難場所
   - UFConfi.MP1 = 公差外 / 理由選択
   - UFMitsubishi = メイン / H494 / 検入高さ / 不良 /（空×3）
   - UF_PackNo = 包装仕様_Ver1.0.2 / Im隠し用
   - UFCask = ｷｬｽｸ専用校正
   - UF品質基準 = Q-S-QA-DIM-122,124,129 / Q-S-QA-SRF-128[別表1][別表2] / Test、品質写真
5. **コードにイベントがあるのに、そのコントロールがフォームに無いもの（削除済みコントロールの残骸。未使用コード候補）**:
   - TEST: `CommandButton13_Click`
   - UF1: `Frame1_Click`, `Frame7_Click`
   - UFCask: `MAXTex_KeyPress`
   - UFMitsubishi: `CommandButton14_DblClick`
   - UF品質基準: `FM1_Click`, `Fm2_Click`
   - （UF管理の `MIMI_AfterUpdate` は大文字小文字を区別しない照合で該当あり）
6. **ボタンはあるがイベントが無いもの**: TEST.CommandButton30、UFAinput.B1〜B5（動的にイベントを付けている可能性あり。要確認）、UFCask.CommandButton10、UFMitsubishi.ws_out、UFMsgBox.CommandButton2「キャンセル」（非表示）、UFProgress.CommandButton1（点滅表示専用）。
7. **設計時から非表示で、表示に切り替えるコードが見つからないもの**: UF_PackNo.CommandButton1「ﾌﾞﾗｳｻﾞで確認」（IE で包装仕様書 Web を開く唯一の入口）→ IE/WebDriver 機能は事実上使われていない（推測）。UF_PackNo.SpinButton1（画像の拡大縮小）も非表示。
8. **ActiveX**: UFProgress.ProgressBar1 = MSComctl ProgressBar。UFNameList.LV1 は CLSID を判定できず `ActiveX#2` と出たが、コード（`UFNameList.LV1.ListItems`）から MSComctl **ListView** と判断できる。どちらも Web 化では HTML の要素に置き換えになる。
9. **SF3**: 入力欄 LineChenge（設備名バーコード）、F1「モード設定」の OP1「オンライン」（初期値=1）/ OP2「オフライン」、CB1「実行」、CommandButton1「確認」、非表示の LB1。TextBox1 に設計時の値 'T716980265' があるが、使うコードは無い。
10. **UFConfi**: P1（公差外）に TB1 と「自動入力」×4（NewHB1 公差外 / NewHB2 許容 / NewHB3 再測定 / NewHB4 保留）、CommandButton1「送信」、List2（再測定理由。非表示）。P2（理由選択）に List1、その他（TextBox）、CommandButton2「決定」、KenName（現検査員）、MotoData（対象データ）。フォーム直下に、引き渡し用の Kpop / Knum / LineName / DelC / TR / TC。
11. **UFMsgBox**: L1（本文）、CommandButton1「OK」、CommandButton2「キャンセル」（非表示）、IM1〜IM4（!/×/?/i のアイコン）、FlLabe（"END"/"Ainput" のフラグ）。
12. Image コントロールの画像（アイコンや検査位置図）は `o` ストリーム内の StdPicture に入っている。今回は取り出していない（必要なら別途抽出する）。

---
# フォーム別コントロール一覧

## Msg（コード: `msg.frm.bas`）

- フォームCaption(設計時): "Microsoft Excel" / ClientWidth×Height(twip): 8430.001×3075 / StartUpPosition: 1  'オーナー フォームの中央 / ShowModal: (既定=True) / コントロール数: 13
- フォームイベント: Activate(L49), QueryClose(L53)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | F1 | Frame |  | ○ |  |  |
| 1 F1 | C1 | CommandButton | C1 | ○ |  | Click(L19) |
| 1 F1 | C2 | CommandButton | C2 | × |  | Click(L24) |
| 1 F1 | C3 | CommandButton | C3 | × |  | Click(L29) |
| 0 | ClickNo | Label | ClickNo | × |  |  |
| 0 | Icon | Image |  | ○ |  |  |
| 0 | Ic1 〜 Ic4（4個） | Image | (空) | × |  |  |
| 0 | CloseTime | Label | CloseTime | × |  |  |
| 0 | typeNo | Label | typeNo | × |  |  |
| 0 | L1 | Label | Msg | ○ |  |  |

## SF3（コード: `SF3.frm.bas`）

- フォームCaption(設計時): "初期設定" / ClientWidth×Height(twip): 4815×5685 / StartUpPosition: 1  'オーナー フォームの中央 / ShowModal: (既定=True) / コントロール数: 10
- フォームイベント: Activate(L68), Initialize(L72)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | Label1 | Label | 設備名 | ○ |  |  |
| 0 | CB1 | CommandButton | 実行 | ○ |  | Click(L26) |
| 0 | F1 | Frame | モード設定 | ○ |  |  |
| 1 F1 | OP1 | OptionButton | オンライン (初期値=1) | ○ |  | Click(L62) |
| 1 F1 | OP2 | OptionButton | オフライン (初期値=0) | ○ |  | Click(L65) |
| 0 | LineChenge | TextBox |  | ○ |  | Change(L52) |
| 0 | CommandButton1 | CommandButton | 確認 | ○ |  | Click(L48) |
| 0 | TextBox1 | TextBox | 値=T716980265 | ○ |  |  |
| 0 | Label2 | Label | 設備バーコードを読み取ってください | ○ |  |  |
| 0 | LB1 | ListBox |  | × |  |  |

## TEST（コード: `Test.frm.bas`）

- フォームCaption(設計時): "UserForm1" / ClientWidth×Height(twip): 18135×8970.001 / StartUpPosition: 1  'オーナー フォームの中央 / ShowModal: (既定=True) / コントロール数: 100
- フォームイベント: Initialize(L1151)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | CommandButton3 | CommandButton | Timeチェック | ○ |  | Click(L564) |
| 0 | T1 | TextBox | 値=VBATEST | ○ |  |  |
| 0 | btn0 | CommandButton | 動かせるボタン | ○ |  | Click(L204), MouseDown(L214), MouseMove(L221), MouseUp(L229) |
| 0 | L1 | Label | Label1 | ○ |  |  |
| 0 | L2 | Label | Label1 | ○ |  |  |
| 0 | CommandButton7 | CommandButton | Call Number_Count⏎ | ○ |  | Click(L761) |
| 0 | CommandButton8 | CommandButton | PDF | ○ |  | Click(L836) |
| 0 | CommandButton9 | CommandButton | DB元ファイル作成 | ○ |  | Click(L870) |
| 0 | TextTest | TextBox | 値=HOT,L-3,FS4,QF,2HI,ST2,ANF,TUR,PSW,KEN,≠NTT* | ○ |  | KeyDown(L985) |
| 0 | CommandButton11 | CommandButton | 日報close HiddenSheets | ○ |  | Click(L235) |
| 0 | TkNo | TextBox | 値=7P0106 | ○ |  |  |
| 0 | Wildcard | CommandButton | ≠　Wildcard | ○ |  | Click(L1157) |
| 0 | CommandButton14 | CommandButton | オーダー参照　あんぐるふた | ○ |  | Click(L311) |
| 0 | Test1 | TextBox | 値=10 | ○ |  |  |
| 0 | Test2 | TextBox | 値=1528 | ○ |  |  |
| 0 | Test3 | TextBox | 値=3055 | ○ |  |  |
| 0 | Test4 | TextBox | 値=12.0 | ○ |  |  |
| 0 | CommandButton15 | CommandButton | Driver.Shutdown | ○ |  | Click(L331) |
| 0 | CommandButton16 | CommandButton | ファイルの有無チェック | ○ |  | Click(L345) |
| 0 | LoopText | TextBox | 値=1 | ○ |  |  |
| 0 | ListBox1 | Label | 11.268 | ○ |  | MouseMove(L1002), BeforeDropOrPaste(L1110) |
| 0 | MoveFlag | TextBox |  | ○ |  |  |
| 0 | Tagtext | TextBox |  | ○ |  |  |
| 0 | CommandButton20 | CommandButton | Com半角 | ○ |  | Click(L400) |
| 0 | A1 | TextBox | 値=探す側 | ○ |  |  |
| 0 | CommandButton24 | CommandButton | DBソリ | ○ |  | Click(L433) |
| 0 | DistortionFrame | Frame | 歪公差 | ○ |  |  |
| 1 DistortionFrame | Label65 | Label | 耳歪 | ○ |  |  |
| 1 DistortionFrame | Label93 | Label | 1M当りｹﾞｰｼﾞ | ○ |  |  |
| 1 DistortionFrame | Mimi | TextBox |  | ○ |  |  |
| 1 DistortionFrame | Label94 | Label | 幅歪 | ○ |  |  |
| 1 DistortionFrame | TaHiz | TextBox |  | ○ |  |  |
| 1 DistortionFrame | HaHiz | TextBox |  | ○ |  |  |
| 1 DistortionFrame | Label95 | Label | 丈歪 | ○ |  |  |
| 1 DistortionFrame | MG | TextBox |  | ○ |  |  |
| 1 DistortionFrame | Label64 | Label | 中歪 | ○ |  |  |
| 1 DistortionFrame | Naka | TextBox |  | ○ |  |  |
| 1 DistortionFrame | NoTex | TextBox |  | ○ |  |  |
| 1 DistortionFrame | HizCome | Label |  | ○ |  |  |
| 0 | TextBox2 | TextBox | 値=R105S07 | ○ |  |  |
| 0 | CommandButton25 | CommandButton | コース検索 | ○ |  | Click(L465) |
| 0 | CourseTex | TextBox | 値=HOT L-3 FS4 QF 2HI ST2 ANF TUR PSW KEN ≠NTT* | ○ |  |  |
| 0 | CommandButton27 | CommandButton | DB品質 | ○ |  | Click(L497) |
| 0 | CommandButton4 | CommandButton | LoadPic | ○ |  | Click(L587) |
| 0 | HTbox | TextBox | 値=AL付着跡 | ○ |  |  |
| 0 | Frame4 | Frame | 検索用 | ○ |  |  |
| 1 Frame4 | T3 | TextBox |  | ○ |  | KeyPress(L901), Change(L923) |
| 1 Frame4 | T2 | TextBox | 値=H6435J0 | ○ |  | KeyPress(L890), AfterUpdate(L912) |
| 1 Frame4 | Label96 | Label | LotNo | ○ |  |  |
| 1 Frame4 | Label97 | Label | KenNo | ○ |  |  |
| 1 Frame4 | CommandButton34 | CommandButton | HikiC | ○ |  | Click(L58) |
| 1 Frame4 | CommandButton38 | CommandButton | 設備仕様 | ○ |  | Click(L186) |
| 0 | CommandButton28 | CommandButton | ページ追加 | ○ |  | Click(L524) |
| 0 | MP1 | MultiPage |  | ○ |  | Change(L886) |
| 1 MP1 | PTes | Page/Form | (タブ表示名)AL付着跡_1 | × |  |  |
| 2 PTes | Image2 | Image |  | ○ |  |  |
| 2 PTes | Frame6 | Frame | Test | ○ |  |  |
| 1 MP1 | Page8 | Page/Form | (タブ表示名)SSマーク_1 | × |  |  |
| 2 Page8 | Frame5 | Frame | Test | ○ |  |  |
| 2 Page8 | Image3 | Image |  | ○ |  |  |
| 1 MP1 | Page9 | Page/Form | (タブ表示名)KeyDow | × |  |  |
| 2 Page9 | Label5 | Label | KeyDownTest用 | ○ |  |  |
| 2 Page9 | PicL | Label |  | ○ |  |  |
| 2 Page9 | TestLabel | Label | aaaaa | ○ |  |  |
| 1 MP1 | Page10 | Page/Form | (タブ表示名)ﾚｼﾞｽﾄﾘ | ○ |  |  |
| 2 Page10 | CommandButton18 | CommandButton | AIM | ○ |  | Click(L362) |
| 2 Page10 | CommandButton22 | CommandButton | test取得 | ○ |  | Click(L408) |
| 2 Page10 | CommandButton23 | CommandButton | DeleteSetting | ○ |  | Click(L418) |
| 2 Page10 | CommandButton21 | CommandButton | HVC | ○ |  | Click(L404) |
| 2 Page10 | NowFocus | Label |  | ○ |  |  |
| 2 Page10 | ListBox2 | Label | 11.268 | ○ |  | MouseMove(L1018), BeforeDropOrPaste(L1124) |
| 2 Page10 | ListBox3 | Label |  | ○ |  | MouseMove(L1034), BeforeDropOrPaste(L1137) |
| 2 Page10 | CommandButton33 | CommandButton | 初期設定 | ○ |  | Click(L54) |
| 1 MP1 | Page11 | Page/Form | (タブ表示名)PageTest | × |  |  |
| 2 Page11 | CommandButton26 | CommandButton | Dir使い方 | ○ |  | Click(L479) |
| 2 Page11 | CommandButton29 | CommandButton | AddTest | ○ |  | Click(L537) |
| 2 Page11 | CommandButton30 | CommandButton | 保護材 | ○ |  |  |
| 2 Page11 | TextBox3 | TextBox | 値=51044352 | ○ |  |  |
| 2 Page11 | CommandButton32 | CommandButton | DB校正ｴｯﾁﾞ | ○ |  | Click(L576) |
| 1 MP1 | Page12 | Page/Form | (タブ表示名)LogLoad | × |  |  |
| 2 Page12 | Frame1 | Frame | Frame1 | ○ |  |  |
| 3 Frame1 | Label1 | Label | LotNo | ○ |  |  |
| 3 Frame1 | CommandButton6 | CommandButton | LogLoad | ○ |  | Click(L600) |
| 3 Frame1 | LoadLot | TextBox | 値=H6375C0 | ○ |  |  |
| 3 Frame1 | LoadWorker | TextBox | 値=TesMan | ○ |  |  |
| 3 Frame1 | Label2 | Label | 検査員 | ○ |  |  |
| 3 Frame1 | Label3 | Label | 参照ファイル名 | ○ |  |  |
| 3 Frame1 | LoadTextName | TextBox | 値=Access.log | ○ |  |  |
| 3 Frame1 | Reason | TextBox |  | ○ |  |  |
| 3 Frame1 | Label4 | Label | 理由 | ○ |  |  |
| 3 Frame1 | CommandButton19 | CommandButton | 文字化け | ○ |  | Click(L366) |
| 3 Frame1 | L3 | TextBox | 値=L-1 | ○ |  |  |
| 0 | CommandButton31 | CommandButton | MoveI | ○ |  | Click(L48) |
| 0 | TextBox4 | TextBox |  | ○ |  |  |
| 0 | TabST | TabStrip |  | ○ |  | Change(L198) |
| 0 | CommandButton35 | CommandButton | kinmu | ○ |  | Click(L143) |
| 0 | CommandButton36 | CommandButton | Wildcard | ○ |  | Click(L148) |
| 0 | CommandButton37 | CommandButton | まいくろ | ○ |  | Click(L168) |
| 0 | CommandButton39 | CommandButton | 時間経過 | ○ |  | Click(L193) |
| 0 | TextBox1 | TextBox | 値=51057370 | ○ |  |  |

- **コードにイベントがあるがフォーム上に同名コントロールが無いもの（未使用コード候補）**: commandbutton13_Click(L243)
- イベントプロシージャが無いCommandButton（他所から動的操作 or 未使用）: CommandButton30

## UF1（コード: `UF1.frm.bas`）

- フォームCaption(設計時): "2026.10.14" / ClientWidth×Height(twip): 17400×10515 / StartUpPosition: 2  '画面の中央 / ShowModal: (既定=True) / コントロール数: 586
- フォームイベント: Initialize(L173), Resize(L4202), Activate(L4218), Click(L4300), DblClick(L4417), Layout(L4421), QueryClose(L4432), Terminate(L4456), MouseMove(L4719)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | NGF | Label |  | × |  | DblClick(L360) |
| 0 | F1 〜 F10（10個） | Label | (空) | × |  |  |
| 0 | DiagF | Label |  | × |  |  |
| 0 | F | Label |  | × |  |  |
| 0 | TEF | TextBox |  | × |  |  |
| 0 | TSF | TextBox |  | × |  |  |
| 0 | LateralF | TextBox |  | × |  | Change(L826), MouseDown(L878), Enter(L1011), Exit(L1030), AfterUpdate(L1283), BeforeUpdate(L1429) |
| 0 | DiagFKou | TextBox |  | × |  | MouseDown(L917), Change(L979), Enter(L1055), Exit(L1075), AfterUpdate(L1094), BeforeUpdate(L1240) |
| 0 | Machin | Label |  | ○ |  |  |
| 0 | Frm2 | Frame | 厚公差 | ○ |  |  |
| 1 Frm2 | Ap | TextBox |  | ○ |  | Change(L547) |
| 1 Frm2 | Label36 | Label | － | ○ |  |  |
| 1 Frm2 | Am | TextBox |  | ○ |  | Change(L442) |
| 1 Frm2 | Label35 | Label | + | ○ |  |  |
| 1 Frm2 | ApK | TextBox |  | ○ |  |  |
| 1 Frm2 | AmK | TextBox |  | ○ |  |  |
| 1 Frm2 | Label41 | Label | ～ | ○ |  |  |
| 0 | Frm3 | Frame | 巾公差 | ○ |  |  |
| 1 Frm3 | Hm | TextBox |  | ○ |  | Change(L3203) |
| 1 Frm3 | Hp | TextBox |  | ○ |  | Change(L3208) |
| 1 Frm3 | Label38 | Label | － | ○ |  |  |
| 1 Frm3 | Label37 | Label | + | ○ |  |  |
| 1 Frm3 | HpK | TextBox |  | ○ |  |  |
| 1 Frm3 | HmK | TextBox |  | ○ |  |  |
| 1 Frm3 | Label42 | Label | ～ | ○ |  |  |
| 0 | Frm4 | Frame | 丈公差 | ○ |  |  |
| 1 Frm4 | Tm | TextBox |  | ○ |  | Change(L4164) |
| 1 Frm4 | Tp | TextBox |  | ○ |  | Change(L4169) |
| 1 Frm4 | Label40 | Label | － | ○ |  |  |
| 1 Frm4 | Label39 | Label | + | ○ |  |  |
| 1 Frm4 | TpK | TextBox |  | ○ |  |  |
| 1 Frm4 | TmK | TextBox |  | ○ |  |  |
| 1 Frm4 | Label43 | Label | ～ | ○ |  |  |
| 0 | S1 | Label |  | ○ |  |  |
| 0 | U1 | Label |  | ○ |  |  |
| 0 | C1 | Label |  | ○ |  |  |
| 0 | D1 | Label |  | ○ |  |  |
| 0 | E1 | Label |  | ○ |  |  |
| 0 | E2 | Label |  | ○ |  |  |
| 0 | S2 | Label |  | ○ |  |  |
| 0 | U2 | Label |  | ○ |  |  |
| 0 | C2 | Label |  | ○ |  |  |
| 0 | D2 | Label |  | ○ |  |  |
| 0 | E4 | Label |  | ○ |  |  |
| 0 | S3 | Label |  | ○ |  |  |
| 0 | U3 | Label |  | ○ |  |  |
| 0 | C3 | Label |  | ○ |  |  |
| 0 | D3 | Label |  | ○ |  |  |
| 0 | E3 | Label |  | ○ |  |  |
| 0 | S4 | Label |  | ○ |  |  |
| 0 | U4 | Label |  | ○ |  |  |
| 0 | C4 | Label |  | ○ |  |  |
| 0 | D4 | Label |  | ○ |  |  |
| 0 | S5 | Label |  | ○ |  |  |
| 0 | U5 | Label |  | ○ |  |  |
| 0 | C5 | Label |  | ○ |  |  |
| 0 | D5 | Label |  | ○ |  |  |
| 0 | E5 | Label |  | ○ |  |  |
| 0 | E6 | Label |  | ○ |  |  |
| 0 | S6 | Label |  | ○ |  |  |
| 0 | U6 | Label |  | ○ |  |  |
| 0 | C6 | Label |  | ○ |  |  |
| 0 | D6 | Label |  | ○ |  |  |
| 0 | E8 | Label |  | ○ |  |  |
| 0 | S7 | Label |  | ○ |  |  |
| 0 | U7 | Label |  | ○ |  |  |
| 0 | C7 | Label |  | ○ |  |  |
| 0 | D7 | Label |  | ○ |  |  |
| 0 | E7 | Label |  | ○ |  |  |
| 0 | S8 | Label |  | ○ |  |  |
| 0 | U8 | Label |  | ○ |  |  |
| 0 | C8 | Label |  | ○ |  |  |
| 0 | D8 | Label |  | ○ |  |  |
| 0 | S9 | Label |  | ○ |  |  |
| 0 | U9 | Label |  | ○ |  |  |
| 0 | C9 | Label |  | ○ |  |  |
| 0 | D9 | Label |  | ○ |  |  |
| 0 | E9 | Label |  | ○ |  |  |
| 0 | E10 | Label |  | ○ |  |  |
| 0 | S10 | Label |  | ○ |  |  |
| 0 | U10 | Label |  | ○ |  |  |
| 0 | C10 | Label |  | ○ |  |  |
| 0 | D10 | Label |  | ○ |  |  |
| 0 | E12 | Label |  | × |  |  |
| 0 | S11 | Label |  | × |  |  |
| 0 | U11 | Label |  | × |  |  |
| 0 | C11 | Label |  | × |  |  |
| 0 | D11 | Label |  | × |  |  |
| 0 | E11 | Label |  | × |  |  |
| 0 | S12 | Label |  | × |  |  |
| 0 | U12 | Label |  | × |  |  |
| 0 | C12 | Label |  | × |  |  |
| 0 | D12 | Label |  | × |  |  |
| 0 | DiagE | Label |  | ○ |  |  |
| 0 | DiagS | Label |  | ○ |  |  |
| 0 | DiagU | Label |  | ○ |  |  |
| 0 | DiagC | Label |  | ○ |  |  |
| 0 | DiagD | Label |  | ○ |  |  |
| 0 | NGE | Label |  | ○ |  | DblClick(L354) |
| 0 | NGS | Label |  | ○ |  | DblClick(L330) |
| 0 | NGU | Label |  | ○ |  | DblClick(L336) |
| 0 | NGC | Label |  | ○ |  | DblClick(L342) |
| 0 | NGD | Label |  | ○ |  | DblClick(L348) |
| 0 | Frm1 | Frame | ｻｲｽﾞ | ○ |  |  |
| 1 Frm1 | Atu | TextBox |  | ○ |  | Change(L551) |
| 1 Frm1 | Tak | TextBox |  | ○ |  | Change(L4161) |
| 1 Frm1 | Hab | TextBox |  | ○ |  | Change(L3199) |
| 1 Frm1 | Label2 | Label | 厚 | ○ |  |  |
| 1 Frm1 | Label3 | Label | 巾 | ○ |  |  |
| 1 Frm1 | Label4 | Label | 丈 | ○ |  |  |
| 0 | L10 | Label | 直角度 | × |  |  |
| 0 | L9 | Label | 対角線差 | ○ |  |  |
| 0 | L1 | Label | 幅1 | ○ |  |  |
| 0 | L2 | Label | 幅2 | ○ |  |  |
| 0 | L3 | Label | 丈1 | ○ |  |  |
| 0 | L4 | Label | 丈2 | ○ |  |  |
| 0 | L5 | Label | 対角A | ○ |  |  |
| 0 | L6 | Label | 対角B | ○ |  |  |
| 0 | L7 | Label |  | × |  |  |
| 0 | L8 | Label |  | × |  |  |
| 0 | AL1 | Label | 厚1 | ○ |  |  |
| 0 | AL2 | Label | 厚2 | ○ |  |  |
| 0 | AL3 | Label | 厚3 | ○ |  |  |
| 0 | AL4 | Label | 厚4 | ○ |  |  |
| 0 | Lng | Label | NG判定 | ○ |  |  |
| 0 | Frm5 | Frame | 直角度 | ○ |  |  |
| 1 Frm5 | Taik | TextBox |  | ○ |  | Change(L4173) |
| 1 Frm5 | Label51 | Label | 等級 | ○ |  |  |
| 1 Frm5 | TaikText | TextBox |  | ○ |  |  |
| 0 | Label18 | Label | ﾏｲｸﾛText: | ○ |  |  |
| 0 | S | Label |  | ○ |  |  |
| 0 | E | Label |  | ○ |  |  |
| 0 | U | Label |  | ○ |  |  |
| 0 | C | Label |  | ○ |  |  |
| 0 | D | Label |  | ○ |  |  |
| 0 | AtCount | Label |  | ○ |  |  |
| 0 | NewKen | CommandButton | 新規検査 | ○ |  | Click(L3359) |
| 0 | 表示 | Label |  | ○ |  | DblClick(L5329) |
| 0 | TSS | TextBox |  | ○ |  |  |
| 0 | TES | TextBox |  | ○ |  |  |
| 0 | TSU | TextBox |  | ○ |  |  |
| 0 | TSD | TextBox |  | ○ |  |  |
| 0 | TSC | TextBox |  | ○ |  |  |
| 0 | TEC | TextBox |  | ○ |  |  |
| 0 | TEU | TextBox |  | ○ |  |  |
| 0 | TEE | TextBox |  | ○ |  |  |
| 0 | TED | TextBox |  | ○ |  |  |
| 0 | TSE | TextBox |  | ○ |  |  |
| 0 | MP1 | MultiPage |  | ○ |  | Click(L1634), MouseMove(L1638), Change(L3354) |
| 1 MP1 | Page1 | Page/Form | (タブ表示名)作業選択 | ○ |  |  |
| 2 Page1 | SlectFM | Frame | 検査選択 | ○ |  | MouseMove(L2478) |
| 3 SlectFM | 厚板 | OptionButton | 厚板 (初期値=0) | ○ |  | Click(L5204) |
| 3 SlectFM | 薄板 | OptionButton | 薄板 (初期値=0) | ○ |  | Click(L5278) |
| 3 SlectFM | 少枚数 | OptionButton | 少枚数 (初期値=0) | ○ |  | Click(L5244) |
| 3 SlectFM | 極厚 | OptionButton | 耳付 (初期値=0) | ○ |  | Click(L5302) |
| 3 SlectFM | AtButt | CommandButton | 厚2ヵ所 | × |  | Click(L4469) |
| 3 SlectFM | InspeButt | CommandButton | 検入≦25 | × |  | Click(L4481) |
| 3 SlectFM | YamaLabel | Label | 山数 | × |  |  |
| 3 SlectFM | スラブ | OptionButton | スラブ (初期値=0) | ○ |  | Click(L5312) |
| 3 SlectFM | Yama | ComboBox |  | × |  | KeyPress(L2235), Change(L2246), MouseDown(L2462) |
| 3 SlectFM | Comment | Label |  | ○ |  |  |
| 2 Page1 | Abno | CommandButton | 異常値on | ○ |  | Click(L434) |
| 2 Page1 | ReSize | CommandButton | ｻｲｽﾞ修正 | ○ | 無 | Click(L4057) |
| 2 Page1 | CommandButton9 | CommandButton | ﾌｫｰﾑ閉 | ○ |  | Click(L3113) |
| 2 Page1 | Sumisei | Label |  | ○ |  |  |
| 2 Page1 | Reien | Label |  | ○ |  |  |
| 2 Page1 | Byousya | CheckBox | 強制描写 (初期値=0) | ○ |  | AfterUpdate(L554) |
| 2 Page1 | Additional | Label |  | ○ |  |  |
| 2 Page1 | ConductivText | Label |  | ○ |  |  |
| 2 Page1 | ConductivL | Label | 導電率写真確認No及びｺﾒﾝﾄ | × |  |  |
| 1 MP1 | Page2 | Page/Form | (タブ表示名)MinMax | × |  |  |
| 2 Page2 | CommandButton4 | CommandButton | MinMax | ○ |  | Click(L4459) |
| 2 Page2 | AMin | Label |  | ○ |  |  |
| 2 Page2 | TMax | Label |  | ○ |  |  |
| 2 Page2 | AMax | Label |  | ○ |  |  |
| 2 Page2 | HMax | Label |  | ○ |  |  |
| 2 Page2 | HMin | Label |  | ○ |  |  |
| 2 Page2 | TMin | Label |  | ○ |  |  |
| 2 Page2 | Label24 | Label | 丈Min | ○ |  |  |
| 2 Page2 | Label22 | Label | 厚Min | ○ |  |  |
| 2 Page2 | Label23 | Label | 巾Min | ○ |  |  |
| 2 Page2 | Label25 | Label | 厚Max | ○ |  |  |
| 2 Page2 | Label26 | Label | 丈Max | ○ |  |  |
| 2 Page2 | Label27 | Label | 巾Max | ○ |  |  |
| 2 Page2 | Frame14 | Frame | 検査法 | ○ |  |  |
| 3 Frame14 | Inspection1 | OptionButton | L1 (初期値=0) | ○ |  |  |
| 3 Frame14 | Inspection2 | OptionButton | L2 (初期値=0) | ○ |  |  |
| 3 Frame14 | Inspection3 | OptionButton | M6 (初期値=0) | ○ |  |  |
| 3 Frame14 | Inspection4 | OptionButton | M2 (初期値=0) | ○ |  |  |
| 2 Page2 | AMode | CommandButton | A_ Input | ○ |  | Click(L447) |
| 2 Page2 | EFM | Frame | ｴﾝﾄﾞ | ○ |  |  |
| 3 EFM | OPE1 | OptionButton | S (初期値=0) | ○ |  | Change(L3554) |
| 3 EFM | OPE5 | OptionButton | E (初期値=0) | ○ |  | Change(L3566) |
| 3 EFM | OPE2 | OptionButton | U (初期値=0) | ○ |  | Change(L3557) |
| 3 EFM | OPE3 | OptionButton | 2 (初期値=0) | ○ |  | Change(L3560) |
| 3 EFM | OPE4 | OptionButton | D (初期値=0) | ○ |  | Change(L3563) |
| 3 EFM | OPE6 | OptionButton | F (初期値=0) | × |  | Change(L3569) |
| 2 Page2 | SFM | Frame | ｽﾀｰﾄ | ○ |  |  |
| 3 SFM | OPS5 | OptionButton | E (初期値=0) | ○ |  | Change(L3547) |
| 3 SFM | OPS1 | OptionButton | S (初期値=0) | ○ |  | Change(L3535) |
| 3 SFM | OPS2 | OptionButton | U (初期値=0) | ○ |  | Change(L3538) |
| 3 SFM | OPS3 | OptionButton | 2 (初期値=0) | ○ |  | Change(L3541) |
| 3 SFM | OPS4 | OptionButton | D (初期値=0) | ○ |  | Change(L3544) |
| 3 SFM | OPS6 | OptionButton | F (初期値=0) | × |  | Change(L3550) |
| 1 MP1 | Page3 | Page/Form | (タブ表示名)登録済検査員 | × |  |  |
| 2 Page3 | T1 | TextBox |  | ○ |  | Change(L4099) |
| 2 Page3 | Label28 | Label | 1直登録 | ○ |  |  |
| 2 Page3 | Label29 | Label | 2直登録 | ○ |  |  |
| 2 Page3 | Label30 | Label | 3直登録 | ○ |  |  |
| 2 Page3 | T2 | TextBox |  | ○ |  | Change(L4121) |
| 2 Page3 | T3 | TextBox |  | ○ |  | Change(L4141) |
| 2 Page3 | Day1 | TextBox |  | ○ |  | Change(L3120) |
| 2 Page3 | Day3 | TextBox |  | ○ |  | Change(L3128) |
| 2 Page3 | Day2 | TextBox |  | ○ |  | Change(L3124) |
| 1 MP1 | Page4 | Page/Form | (タブ表示名)印刷 | × |  |  |
| 2 Page4 | PrintB | CommandButton | 印刷 | ○ |  | Click(L3619) |
| 2 Page4 | CommandButton18 | CommandButton | 保留 | ○ |  | Click(L120) |
| 2 Page4 | PrintBreset | CommandButton | 印刷履歴クリア | × |  | Click(L3599) |
| 1 MP1 | Page5 | Page/Form | (タブ表示名)ﾌｫｰﾑ | × |  |  |
| 2 Page5 | ZEB | CommandButton | 全数.三菱.5N | ○ |  | Click(L4874) |
| 2 Page5 | KYB | CommandButton | ｷｬｽｸ | ○ |  | Click(L3230) |
| 2 Page5 | CommandButton7 | CommandButton | 履歴 | ○ |  | Click(L3048) |
| 2 Page5 | CommandButton8 | CommandButton | 品質一覧 | ○ |  | Click(L3053) |
| 2 Page5 | H494B | CommandButton | H494 | ○ |  | Click(L604) |
| 1 MP1 | Page6 | Page/Form | (タブ表示名)検査位置 | × |  |  |
| 2 Page6 | L1IM | Image |  | × |  |  |
| 2 Page6 | LHIM | Image |  | × |  |  |
| 2 Page6 | TBIM | Image |  | × |  |  |
| 2 Page6 | GokuIM | Image |  | × |  |  |
| 2 Page6 | SurabuIM | Image |  | × |  |  |
| 1 MP1 | Test1 | Page/Form | (タブ表示名)Log検出 | × |  |  |
| 2 Test1 | Frame10 | Frame |  | ○ |  |  |
| 3 Frame10 | CB | ComboBox |  | ○ |  | Change(L4494) |
| 3 Frame10 | Label19 | Label | 開始時間(S1測定時間) | ○ |  |  |
| 3 Frame10 | EachValue | Label |  | ○ |  |  |
| 3 Frame10 | EachTime | Label |  | ○ |  |  |
| 3 Frame10 | StartTime | Label |  | ○ |  |  |
| 3 Frame10 | EndTime | Label |  | ○ |  |  |
| 3 Frame10 | Label20 | Label | 都度 | ○ |  |  |
| 3 Frame10 | Label21 | Label | 終了時間(印刷ﾀｲﾐﾝｸﾞ) | ○ |  |  |
| 3 Frame10 | StartTimeTY | Label |  | ○ |  |  |
| 3 Frame10 | EndTimeTY | Label |  | ○ |  |  |
| 3 Frame10 | CommandButton6 | CommandButton | 校正起動 | ○ |  | Click(L3029) |
| 3 Frame10 | PakeD | Label |  | ○ |  |  |
| 3 Frame10 | Label32 | Label | ﾊﾟｹｯﾄ | ○ |  |  |
| 3 Frame10 | CommandButton14 | CommandButton | ｼｰﾄ表示 | ○ |  | Click(L3024) |
| 3 Frame10 | Label78 | Label | Delete | ○ |  |  |
| 3 Frame10 | DeleteC | Label |  | ○ |  |  |
| 3 Frame10 | DeleteP | Label |  | ○ |  |  |
| 3 Frame10 | DBflag | TextBox |  | ○ |  |  |
| 3 Frame10 | Label58 | Label | 引当情報書込み | ○ |  |  |
| 3 Frame10 | Label44 | Label | DBflag:日稲鯖優先ﾌﾗｸﾞ | ○ |  |  |
| 3 Frame10 | TextBox2 | TextBox | 値=#L1528.2  dt10000+0000014.988M  #l3055.2  <社員カード番号> | ○ |  |  |
| 1 MP1 | Test2 | Page/Form | (タブ表示名)Test1 | × |  |  |
| 2 Test2 | Label8 | Label | Load | ○ |  |  |
| 2 Test2 | Label11 | Label | "STOP" 判定音.msgbox停止　ｲﾆｼｬﾗｲｽﾞでStop　Load_OPでSTOP | ○ |  |  |
| 2 Test2 | Load | Label | on | ○ |  |  |
| 2 Test2 | FlagLabel | Label | Flag1 | ○ |  |  |
| 2 Test2 | Auto | Label | on | ○ |  |  |
| 2 Test2 | Label6 | Label | INP＿Changeｲﾍﾞﾝﾄ用：１でﾙｰﾌﾟｽﾀｰﾄ | ○ |  |  |
| 2 Test2 | Label7 | Label | ｲﾆｼｬﾗｲｽﾞon→空白探し抜ける＊ﾃﾞﾌｫon.ﾗﾍﾞﾙｸﾘｯｸでoff | ○ |  |  |
| 2 Test2 | Label9 | Label | F1 | ○ |  |  |
| 2 Test2 | Label10 | Label | Auto | ○ |  |  |
| 2 Test2 | Label12 | Label | Abno.tag = ﾃﾞﾌｫﾙﾄ””　”NG”で検査停止 | ○ |  |  |
| 2 Test2 | UFTag | Label | UFTag | ○ |  |  |
| 2 Test2 | Label45 | Label | UF1.Tag  | ○ |  |  |
| 2 Test2 | Label55 | Label | ﾘｻｲｽﾞ有無 | ○ |  |  |
| 2 Test2 | ResizText | Label |  | ○ |  |  |
| 2 Test2 | Label57 | Label | 校正flag | ○ |  |  |
| 2 Test2 | ProofreadFlag | Label |  | ○ |  |  |
| 2 Test2 | SplitFlag | Label |  | ○ |  |  |
| 2 Test2 | Label76 | Label | 分割flag | ○ |  |  |
| 2 Test2 | Label5 | Label | UserForm_Activate　T1_Change　CB_Change停止用　MInMax　Print　⏎校正… | ○ |  |  |
| 2 Test2 | Label1 | Label | UF1.INP.Tagに座標を送り座標BOXにコンバート後の数値を送る | ○ |  |  |
| 2 Test2 | TagText | Label | ﾀｰｹﾞｯﾄ | ○ |  |  |
| 1 MP1 | Detect | Page/Form | (タブ表示名)Test2 | × |  |  |
| 2 Detect | CommandButton15 | CommandButton | Detection | ○ |  | Click(L579) |
| 2 Detect | Label47 | Label | LotNo | ○ |  |  |
| 2 Detect | LoadLot | TextBox |  | ○ |  | KeyPress(L1607) |
| 2 Detect | LoadWorker | TextBox |  | ○ |  | Change(L1618) |
| 2 Detect | Label48 | Label | 検査員 | ○ |  |  |
| 2 Detect | Label49 | Label | Desktop_参照ﾌｧｲﾙ名 | ○ |  |  |
| 2 Detect | LoadTextName | TextBox | 値=Access.log | ○ |  |  |
| 2 Detect | Reason | TextBox |  | ○ |  |  |
| 2 Detect | Label50 | Label | 理由 | ○ |  |  |
| 2 Detect | Label96 | Label | ※ANSI保存のファイルでないと文字化けして読み込めずLotNoを探せません。あとまぁまぁ重い。 | ○ |  |  |
| 1 MP1 | Test3 | Page/Form | (タブ表示名)Special | × |  |  |
| 2 Test3 | DebugTaikaku | TextBox |  | ○ |  |  |
| 2 Test3 | Label73 | Label | 対角 | ○ |  |  |
| 2 Test3 | DebugLate | TextBox |  | ○ |  |  |
| 2 Test3 | Label74 | Label | ラテ | ○ |  |  |
| 2 Test3 | Label59 | Label | MoveFlag.text＝関数ObjMoveTextで表示させたオブジェクト名　Test環境が必要⏎歪数値Tex… | ○ |  |  |
| 2 Test3 | Label60 | Label | MouseMove時にﾗﾍﾞﾙのTagに転記させた確認用　Test環境が必要 | ○ |  |  |
| 2 Test3 | MoveFlag | TextBox |  | ○ |  |  |
| 2 Test3 | LabeTag | TextBox |  | ○ |  |  |
| 2 Test3 | ErFM | Frame |  | ○ |  | MouseMove(L2497) |
| 2 Test3 | dataLoding | Label |  | × |  |  |
| 2 Test3 | Label97 | Label | 追記公差があった場合、等級での公差を表示 | ○ |  |  |
| 2 Test3 | Label98 | Label | 履歴呼び出しフラグ | ○ |  |  |
| 0 | WorkButton | CommandButton | 検査員登録 | ○ |  | Click(L4743) |
| 0 | Worker | TextBox |  | ○ |  | Change(L4788), Exit(L4867), Enter(L4870) |
| 0 | CommandButton2 | CommandButton | ﾌｫｰｶｽ | ○ |  | Click(L4537) |
| 0 | INP | TextBox |  | ○ |  | Change(L4565) |
| 0 | CommandButton3 | CommandButton | 手入力(要3桁) | ○ |  | Click(L4552) |
| 0 | Manual | TextBox |  | ○ |  | Change(L3351) |
| 0 | NekoIM1 〜 NekoIM13（13個） | Image | (空) | × |  |  |
| 0 | CommandButton10 | CommandButton | LOTNO登録 | ○ |  | Click(L2937) |
| 0 | NekoIM14 | Image |  | × |  |  |
| 0 | Label33 | Label | 現場Neko_15 | × |  |  |
| 0 | NekoIM15 | Image |  | × |  |  |
| 0 | VC_A | Image |  | × |  |  |
| 0 | VC_B | Image |  | × |  |  |
| 0 | VC_L_A | Label | A面 | × |  |  |
| 0 | VC_L_B | Label | B面 | × |  |  |
| 0 | TestButton | CommandButton | Test | × |  | Click(L1876) |
| 0 | VcNothing | TextBox |  | × |  |  |
| 0 | NekoText | TextBox | 値=対角A：測定 | × |  |  |
| 0 | StartTB | TextBox |  | ○ |  | KeyDown(L3295), AfterUpdate(L3307) |
| 0 | IzyouFl | CommandButton | 廃棄ﾌﾗｸﾞ | ○ |  | Click(L3213) |
| 0 | JudgBox1 | TextBox |  | ○ |  |  |
| 0 | JudgBox2 | TextBox |  | ○ |  |  |
| 0 | LZTKlabel | Label | ﾛｯﾄNo.材調質.用途.検査No | ○ |  |  |
| 0 | L11 | Label | ﾗﾃﾗﾙﾎﾞｰ | × |  |  |
| 0 | LateralS | TextBox |  | × |  | Change(L771), MouseDown(L838), Enter(L991), Exit(L1015), AfterUpdate(L1268), BeforeUpdate(L1287) |
| 0 | LateralU | TextBox |  | × |  | Change(L782), MouseDown(L846), Enter(L995), Exit(L1018), AfterUpdate(L1271), BeforeUpdate(L1317) |
| 0 | LateralD | TextBox |  | × |  | Change(L804), MouseDown(L862), Enter(L1003), Exit(L1024), AfterUpdate(L1277), BeforeUpdate(L1373) |
| 0 | LateralC | TextBox |  | × |  | Change(L793), MouseDown(L854), Enter(L999), Exit(L1021), AfterUpdate(L1274), BeforeUpdate(L1345) |
| 0 | LateralE | TextBox |  | × |  | Change(L815), MouseDown(L870), Enter(L1007), Exit(L1027), AfterUpdate(L1280), BeforeUpdate(L1401) |
| 0 | Frm6 | Frame | ﾗﾃﾗﾙﾎﾞｰ | ○ |  |  |
| 1 Frm6 | Lateral | TextBox |  | ○ |  | Change(L4186) |
| 1 Frm6 | Label53 | Label | 等級 | ○ |  |  |
| 1 Frm6 | LateralText | TextBox |  | ○ |  |  |
| 0 | Conductiv | CommandButton | 導電率ﾌﾗｸﾞ | ○ |  | Click(L591) |
| 0 | PicFM | Frame |  | ○ | 0 |  |
| 0 | Label72 | Label | 耳付き用丈3,4用((UF上にセット要) | × |  |  |
| 0 | E14 | Label |  | × |  |  |
| 0 | L12 | Label |  | × |  |  |
| 0 | L13 | Label |  | × |  |  |
| 0 | S13 | Label |  | × |  |  |
| 0 | U13 | Label |  | × |  |  |
| 0 | C13 | Label |  | × |  |  |
| 0 | D13 | Label |  | × |  |  |
| 0 | E13 | Label |  | × |  |  |
| 0 | S14 | Label |  | × |  |  |
| 0 | U14 | Label |  | × |  |  |
| 0 | C14 | Label |  | × |  |  |
| 0 | D14 | Label |  | × |  |  |
| 0 | Frm7 | Frame | ﾌﾗｯﾄﾈｽ | ○ |  |  |
| 1 Frm7 | flatnes | TextBox |  | ○ |  |  |
| 1 Frm7 | Label77 | Label | 等級 | ○ |  |  |
| 1 Frm7 | flatnesText | TextBox |  | ○ |  |  |
| 0 | MPEX | MultiPage |  | ○ |  | Change(L1642), Click(L1677), MouseMove(L1713) |
| 1 MPEX | 歪 | Page/Form | (タブ表示名)検入高さ | × |  |  |
| 2 歪 | AorB | CommandButton | Now_A | ○ | A | Click(L477) |
| 2 歪 | blindfold | Image |  | × |  |  |
| 2 歪 | HizCome | Label |  | ○ |  |  |
| 2 歪 | FMS | Frame |  | ○ |  | Click(L3168) |
| 3 FMS | CBS1 | ComboBox |  | ○ |  | Exit(L2625), Enter(L2775), Change(L2899) |
| 3 FMS | CBS2 | ComboBox |  | ○ |  | Exit(L2664), Enter(L2806) |
| 3 FMS | CBS3 | ComboBox |  | ○ |  | Exit(L2701), Enter(L2837), Change(L2918) |
| 3 FMS | CBS4 | ComboBox |  | ○ |  | Exit(L2738), Enter(L2868) |
| 3 FMS | LZS | Label | 定盤: | ○ |  | Click(L3246) |
| 3 FMS | LGS | Label | ｹﾞｰｼﾞ: | ○ |  |  |
| 3 FMS | ZS | TextBox |  | ○ |  | Exit(L2513), Enter(L2532), Change(L4901), MouseDown(L4985), AfterUpdate(L5047), BeforeUpdate(L5084) |
| 3 FMS | GS | TextBox |  | ○ |  | Exit(L2569), Enter(L2588), MouseDown(L5016), Change(L5115), AfterUpdate(L5136), BeforeUpdate(L5173) |
| 2 歪 | FMF | Frame |  | × |  | Click(L3193) |
| 3 FMF | CBF1 | ComboBox |  | ○ |  | Exit(L2657), Enter(L2800), Change(L2914) |
| 3 FMF | CBF2 | ComboBox |  | ○ |  | Exit(L2694), Enter(L2831) |
| 3 FMF | CBF3 | ComboBox |  | ○ |  | Exit(L2731), Enter(L2862), Change(L2933) |
| 3 FMF | CBF4 | ComboBox |  | ○ |  | Exit(L2768), Enter(L2893) |
| 3 FMF | LZF | Label | 定盤: | ○ |  |  |
| 3 FMF | LGF | Label | ｹﾞｰｼﾞ: | ○ |  |  |
| 3 FMF | ZF | TextBox |  | ○ |  | Exit(L2528), Enter(L2562), Change(L4947), MouseDown(L5010), AfterUpdate(L5077), BeforeUpdate(L5109) |
| 3 FMF | GF | TextBox |  | ○ |  | Exit(L2584), Enter(L2618), MouseDown(L5041), AfterUpdate(L5166), BeforeUpdate(L5198) |
| 2 歪 | FME | Frame |  | ○ |  | Click(L3188) |
| 3 FME | CBE1 | ComboBox |  | ○ |  | Exit(L2651), Enter(L2795), Change(L2911) |
| 3 FME | CBE2 | ComboBox |  | ○ |  | Exit(L2688), Enter(L2826) |
| 3 FME | CBE3 | ComboBox |  | ○ |  | Exit(L2725), Enter(L2857), Change(L2930) |
| 3 FME | CBE4 | ComboBox |  | ○ |  | Exit(L2762), Enter(L2888) |
| 3 FME | LZE | Label | 定盤: | ○ |  | Click(L3262) |
| 3 FME | LGE | Label | ｹﾞｰｼﾞ: | ○ |  |  |
| 3 FME | ZE | TextBox |  | ○ |  | Exit(L2525), Enter(L2556), Change(L4938), MouseDown(L5005), AfterUpdate(L5071), BeforeUpdate(L5104) |
| 3 FME | GE | TextBox |  | ○ |  | Exit(L2581), Enter(L2612), MouseDown(L5036), Change(L5131), AfterUpdate(L5160), BeforeUpdate(L5193) |
| 2 歪 | FMU | Frame |  | ○ |  | Click(L3173) |
| 3 FMU | CBU1 | ComboBox |  | ○ |  | Exit(L2633), Enter(L2780), Change(L2902) |
| 3 FMU | CBU2 | ComboBox |  | ○ |  | Exit(L2670), Enter(L2811) |
| 3 FMU | CBU3 | ComboBox |  | ○ |  | Exit(L2707), Enter(L2842), Change(L2921) |
| 3 FMU | CBU4 | ComboBox |  | ○ |  | Exit(L2744), Enter(L2873) |
| 3 FMU | LZU | Label | 定盤: | ○ |  | Click(L3250) |
| 3 FMU | LGU | Label | ｹﾞｰｼﾞ: | ○ |  |  |
| 3 FMU | ZU | TextBox |  | ○ |  | Exit(L2516), Enter(L2538), Change(L4911), MouseDown(L4990), AfterUpdate(L5053), BeforeUpdate(L5089) |
| 3 FMU | GU | TextBox |  | ○ |  | Exit(L2572), Enter(L2594), MouseDown(L5021), Change(L5119), AfterUpdate(L5142), BeforeUpdate(L5178) |
| 2 歪 | FMC | Frame |  | ○ |  | Click(L3178) |
| 3 FMC | CBC1 | ComboBox |  | ○ |  | Exit(L2639), Enter(L2785), Change(L2905) |
| 3 FMC | CBC2 | ComboBox |  | ○ |  | Exit(L2676), Enter(L2816) |
| 3 FMC | CBC3 | ComboBox |  | ○ |  | Exit(L2713), Enter(L2847), Change(L2924) |
| 3 FMC | CBC4 | ComboBox |  | ○ |  | Exit(L2750), Enter(L2878) |
| 3 FMC | LZC | Label | 定盤: | ○ |  | Click(L3254) |
| 3 FMC | LGC | Label | ｹﾞｰｼﾞ: | ○ |  |  |
| 3 FMC | ZC | TextBox |  | ○ |  | Exit(L2519), Enter(L2544), Change(L4920), MouseDown(L4995), AfterUpdate(L5059), BeforeUpdate(L5094) |
| 3 FMC | GC | TextBox |  | ○ |  | Exit(L2575), Enter(L2600), MouseDown(L5026), Change(L5123), AfterUpdate(L5148), BeforeUpdate(L5183) |
| 2 歪 | FMD | Frame |  | ○ |  | Click(L3183) |
| 3 FMD | CBD1 | ComboBox |  | ○ |  | Exit(L2645), Enter(L2790), Change(L2908) |
| 3 FMD | CBD2 | ComboBox |  | ○ |  | Exit(L2682), Enter(L2821) |
| 3 FMD | CBD3 | ComboBox |  | ○ |  | Exit(L2719), Enter(L2852), Change(L2927) |
| 3 FMD | CBD4 | ComboBox |  | ○ |  | Exit(L2756), Enter(L2883) |
| 3 FMD | LZD | Label | 定盤: | ○ |  | Click(L3258) |
| 3 FMD | LGD | Label | ｹﾞｰｼﾞ: | ○ |  |  |
| 3 FMD | ZD | TextBox |  | ○ |  | Exit(L2522), Enter(L2550), Change(L4929), MouseDown(L5000), AfterUpdate(L5065), BeforeUpdate(L5099) |
| 3 FMD | GD | TextBox |  | ○ |  | Exit(L2578), Enter(L2606), MouseDown(L5031), Change(L5127), AfterUpdate(L5154), BeforeUpdate(L5188) |
| 1 MPEX | 外観 | Page/Form | (タブ表示名)外観 | × |  |  |
| 2 外観 | Label79 | Label | 上面 | ○ |  |  |
| 2 外観 | Label80 | Label | 下面 | ○ |  |  |
| 2 外観 | S_UpVisual | Label |  | ○ |  |  |
| 2 外観 | E_UpVisual | Label |  | ○ |  |  |
| 2 外観 | U_UpVisual | Label |  | ○ |  |  |
| 2 外観 | C_UpVisual | Label |  | ○ |  |  |
| 2 外観 | D_UpVisual | Label |  | ○ |  |  |
| 2 外観 | LC1 | Label |  | ○ |  | Click(L1458), DblClick(L1507) |
| 2 外観 | LC5 | Label |  | ○ |  | Click(L1474), DblClick(L1523) |
| 2 外観 | LC2 | Label |  | ○ |  | Click(L1462), DblClick(L1511) |
| 2 外観 | LC3 | Label |  | ○ |  | Click(L1466), DblClick(L1515) |
| 2 外観 | LC4 | Label |  | ○ |  | Click(L1470), DblClick(L1519) |
| 2 外観 | E_UnderVisual | Label |  | ○ |  |  |
| 2 外観 | S_UnderVisual | Label |  | ○ |  |  |
| 2 外観 | U_UnderVisual | Label |  | ○ |  |  |
| 2 外観 | C_UnderVisual | Label |  | ○ |  |  |
| 2 外観 | D_UnderVisual | Label |  | ○ |  |  |
| 2 外観 | LC7 | Label |  | ○ |  | Click(L1482), DblClick(L1531) |
| 2 外観 | LC11 | Label |  | ○ |  | Click(L1498) |
| 2 外観 | LC8 | Label |  | ○ |  | Click(L1486), DblClick(L1535) |
| 2 外観 | LC9 | Label |  | ○ |  | Click(L1490), DblClick(L1539) |
| 2 外観 | LC10 | Label |  | ○ |  | Click(L1494), DblClick(L1543) |
| 2 外観 | Label88 | Label | 欠陥名選択 | ○ |  |  |
| 2 外観 | LC6 | Label |  | × |  | Click(L1478), DblClick(L1527) |
| 2 外観 | F_UpVisual | Label |  | × |  |  |
| 2 外観 | F_UnderVisual | Label |  | × |  |  |
| 2 外観 | LC12 | Label |  | × |  | Click(L1502) |
| 2 外観 | LBVisual | ListBox |  | ○ |  |  |
| 1 MPEX | BH | Page/Form | (タブ表示名)歪 | × |  |  |
| 2 BH | SBH | TextBox |  | ○ |  | Change(L1835), AfterUpdate(L1851) |
| 2 BH | UBH | TextBox |  | ○ |  | Change(L1838), AfterUpdate(L1856) |
| 2 BH | DBH | TextBox |  | ○ |  | Change(L1844), AfterUpdate(L1866) |
| 2 BH | CBH | TextBox |  | ○ |  | Change(L1841), AfterUpdate(L1861) |
| 2 BH | EBH | TextBox |  | ○ |  | Change(L1847), AfterUpdate(L1871) |
| 2 BH | LastHot | Label |  | ○ |  |  |
| 2 BH | QF | Label |  | ○ |  |  |
| 2 BH | NoteL | Label | 操業基準:Q-S-KE-ALL-053_バーコール硬度計による測定で指定された材調質のみ合否判定をする | ○ |  |  |
| 1 MPEX | OrderP | Page/Form | (タブ表示名)BH | × |  |  |
| 2 OrderP | PackingComment | Label |  | ○ |  |  |
| 2 OrderP | NoL | Label | オーダーNo.包装仕様No | ○ |  | Click(L1717) |
| 2 OrderP | OrdNoCB | ComboBox |  | ○ |  | Change(L1721) |
| 2 OrderP | PackingNo | Label |  | ○ |  |  |
| 2 OrderP | PackingForm | Label |  | ○ |  |  |
| 1 MPEX | LineSelect | Page/Form | (タブ表示名)不良 | × |  |  |
| 2 LineSelect | NoFrame | Frame | ﾗｲﾝNo | ○ |  |  |
| 3 NoFrame | LineNo | Label |  | ○ |  |  |
| 3 NoFrame | A1 | OptionButton | A22 (初期値=0) | ○ |  | Click(L367) |
| 3 NoFrame | A11 | OptionButton | B5 (初期値=0) | ○ |  | Click(L427) |
| 3 NoFrame | A7 | OptionButton | A16 (初期値=0) | ○ |  | Click(L403) |
| 3 NoFrame | A2 | OptionButton | A21 (初期値=0) | ○ |  | Click(L373) |
| 3 NoFrame | A4 | OptionButton | A19 (初期値=0) | ○ |  | Click(L385) |
| 3 NoFrame | A3 | OptionButton | A20 (初期値=0) | ○ |  | Click(L379) |
| 3 NoFrame | A6 | OptionButton | 場外 (初期値=0) | ○ |  | Click(L397) |
| 3 NoFrame | A5 | OptionButton | A18 (初期値=0) | ○ |  | Click(L391) |
| 3 NoFrame | A8 | OptionButton | A15 (初期値=0) | ○ |  | Click(L409) |
| 3 NoFrame | A9 | OptionButton | A14 (初期値=0) | ○ |  | Click(L415) |
| 3 NoFrame | A10 | OptionButton | A13 (初期値=0) | ○ |  | Click(L421) |
| 2 LineSelect | WorkFrame | Frame | 設備 | ○ |  |  |
| 3 WorkFrame | L1Work | OptionButton | L-1 (初期値=0) | ○ |  | Click(L3267) |
| 3 WorkFrame | LVC | OptionButton | LVC (初期値=0) | ○ |  | Click(L3277) |
| 3 WorkFrame | Tott | OptionButton | ﾄｯﾄ (初期値=0) | ○ |  | Click(L3315) |
| 3 WorkFrame | HVC | OptionButton | HVC (初期値=0) | ○ |  | Click(L3286) |
| 3 WorkFrame | Test | OptionButton | Test (初期値=0) | ○ |  | Click(L3333) |
| 3 WorkFrame | Bala | OptionButton | ﾊﾞﾗﾝｻｰ (初期値=0) | ○ |  | Click(L3324) |
| 3 WorkFrame | Maru | OptionButton | 丸徳 (初期値=0) | ○ |  | Click(L3342) |
| 3 WorkFrame | LineName | Label |  | ○ |  |  |
| 1 MPEX | 避難場所 | Page/Form | (タブ表示名)ｵｰﾀﾞｰ検索 | × |  |  |
| 1 MPEX | 検入高 | Page/Form | (タブ表示名)Line | ○ |  |  |
| 2 検入高 | KenHBox1 | TextBox |  | × |  | AfterUpdate(L696) |
| 2 検入高 | KenHL1 | Label | ｍｍ | × |  |  |
| 2 検入高 | KenHL2 | Label | ｍｍ | × |  |  |
| 2 検入高 | KenHBox2 | TextBox |  | × |  | AfterUpdate(L701) |
| 2 検入高 | KenHBox3 | TextBox |  | × |  | AfterUpdate(L707) |
| 2 検入高 | KenHL4 | Label | ｍｍ | × |  |  |
| 2 検入高 | KenHBox4 | TextBox |  | × |  | AfterUpdate(L712) |
| 2 検入高 | KenHL5 | Label | ｍｍ | × |  |  |
| 2 検入高 | KenHBox5 | TextBox |  | × |  | AfterUpdate(L717) |
| 2 検入高 | KenHBox6 | TextBox |  | × |  | AfterUpdate(L722) |
| 2 検入高 | HL1 〜 HL6（6個） | Label | (空) | ○ |  |  |
| 2 検入高 | Label87 | Label | 山数選択 | ○ |  |  |
| 2 検入高 | KenHBox7 | TextBox |  | × |  | AfterUpdate(L727), Change(L742) |
| 2 検入高 | KenHL7 | Label | 枚 | × |  |  |
| 2 検入高 | KenH | Label |  | ○ |  |  |
| 2 検入高 | 高さ計算式 | Label |  | ○ |  |  |
| 2 検入高 | KenCount | TextBox |  | × |  | AfterUpdate(L745), Change(L766) |
| 2 検入高 | KenHL10 | Label | 回 | × |  |  |
| 2 検入高 | KenHL9 | Label | 検査回数 | ○ |  |  |
| 2 検入高 | KenHL8 | Label | 検入 | ○ |  |  |
| 2 検入高 | KenHL3 | Label | ｍｍ | × |  |  |
| 2 検入高 | KenHL6 | Label | ｍｍ | × |  |  |
| 2 検入高 | YamaLB | ListBox |  | ○ |  | Click(L2508) |
| 1 MPEX | Badboard | Page/Form | (タブ表示名)避難場所 | × |  |  |
| 2 Badboard | Label89 | Label | 不良要因選択 | ○ |  |  |
| 2 Badboard | EBad | TextBox |  | ○ |  | Change(L1803), MouseDown(L1826) |
| 2 Badboard | SBad | TextBox |  | ○ |  | Change(L1791), MouseDown(L1810) |
| 2 Badboard | UBad | TextBox |  | ○ |  | Change(L1794), MouseDown(L1814) |
| 2 Badboard | DBad | TextBox |  | ○ |  | Change(L1800), MouseDown(L1822) |
| 2 Badboard | CBad | TextBox |  | ○ |  | Change(L1797), MouseDown(L1818) |
| 2 Badboard | ListReflect | CommandButton | 反映 | ○ |  | Click(L1548) |
| 2 Badboard | Label90 | Label | 単位(枚) | ○ |  |  |
| 2 Badboard | ViewCB | CommandButton | 確認View | ○ |  | Click(L1931) |
| 2 Badboard | Cle | CommandButton | ｸﾘｱ | ○ |  | Click(L567) |
| 2 Badboard | FBad | TextBox |  | × |  | Change(L1806), MouseDown(L1830) |
| 2 Badboard | DefectFactorLB | ListBox |  | ○ |  |  |
| 0 | Image1 | Image |  | × |  | Click(L630) |
| 0 | Frm8 | Frame |  | ○ |  |  |
| 1 Frm8 | Tyu | TextBox |  | ○ |  |  |
| 1 Frm8 | Zai | TextBox |  | ○ |  |  |
| 1 Frm8 | Youto | TextBox |  | ○ |  |  |
| 1 Frm8 | KenNo | Label |  | ○ |  |  |
| 0 | FU | Label |  | ○ |  |  |
| 0 | FS | Label |  | ○ |  |  |
| 0 | FC | Label |  | ○ |  |  |
| 0 | FD | Label |  | ○ |  |  |
| 0 | DiagSKou | TextBox |  | × |  | MouseDown(L887), Change(L924), Enter(L1035), Exit(L1060), AfterUpdate(L1079), BeforeUpdate(L1098) |
| 0 | DiagUKou | TextBox |  | × |  | MouseDown(L893), Change(L935), Enter(L1039), Exit(L1063), AfterUpdate(L1082), BeforeUpdate(L1126) |
| 0 | DiagDKou | TextBox |  | × |  | MouseDown(L905), Change(L957), Enter(L1047), Exit(L1069), AfterUpdate(L1088), BeforeUpdate(L1184) |
| 0 | DiagCKou | TextBox |  | × |  | MouseDown(L899), Change(L946), Enter(L1043), Exit(L1066), AfterUpdate(L1085), BeforeUpdate(L1156) |
| 0 | DiagEKou | TextBox |  | × |  | MouseDown(L911), Change(L968), Enter(L1051), Exit(L1072), AfterUpdate(L1091), BeforeUpdate(L1212) |
| 0 | FE | Label |  | ○ |  |  |
| 0 | DistortionFrame | Frame | 歪公差 | ○ |  |  |
| 1 DistortionFrame | Label65 | Label | 耳歪 | ○ |  |  |
| 1 DistortionFrame | Label93 | Label | 1M当りｹﾞｰｼﾞ | ○ |  |  |
| 1 DistortionFrame | Mimi | TextBox |  | ○ |  |  |
| 1 DistortionFrame | Label94 | Label | 幅歪 | ○ |  |  |
| 1 DistortionFrame | TaHiz | TextBox |  | ○ |  |  |
| 1 DistortionFrame | HaHiz | TextBox |  | ○ |  |  |
| 1 DistortionFrame | Label95 | Label | 丈歪 | ○ |  |  |
| 1 DistortionFrame | MG | TextBox |  | ○ |  |  |
| 1 DistortionFrame | Label64 | Label | 中歪 | ○ |  |  |
| 1 DistortionFrame | Naka | TextBox |  | ○ |  |  |
| 1 DistortionFrame | Label99 | Label | ﾌﾘｰ反り | ○ |  |  |
| 1 DistortionFrame | HURIS | TextBox |  | ○ |  |  |
| 0 | HiddenFrame | Frame |  | ○ |  |  |
| 1 HiddenFrame | CommandButton13 | CommandButton |  | ○ |  | Click(L3004) |
| 1 HiddenFrame | PasswordTex | TextBox |  | × |  |  |
| 0 | Loading | Label | Loading..... | × |  |  |
| 0 | blindfm | Frame |  | ○ |  | DblClick(L115) |
| 1 blindfm | UfText | TextBox |  | × |  |  |
| 0 | FF | Label |  | ○ |  |  |
| 0 | IMOK | Image |  | ○ |  |  |
| 0 | IMNG | Image |  | ○ |  |  |
| 0 | Lot | TextBox |  | ○ |  | Exit(L4652), Enter(L4657), Change(L4660), KeyPress(L4708) |
| 0 | CheckTF | TextBox |  | × |  |  |
| 0 | CheckTE | TextBox |  | ○ |  |  |
| 0 | CheckTD | TextBox |  | ○ |  |  |
| 0 | CheckTC | TextBox |  | ○ |  |  |
| 0 | CheckTU | TextBox |  | ○ |  |  |
| 0 | CheckTS | TextBox |  | ○ |  |  |

- **コードにイベントがあるがフォーム上に同名コントロールが無いもの（未使用コード候補）**: frame1_Click(L4559), frame7_Click(L4562)

## UFAinput（コード: `UFAinput.frm.bas`）

- フォームCaption(設計時): "2026.06.08" / ClientWidth×Height(twip): 11550×10785 / StartUpPosition: 1  'オーナー フォームの中央 / ShowModal: (既定=True) / コントロール数: 147
- フォームイベント: Click(L774), Layout(L782), Activate(L798), Initialize(L880), MouseMove(L895), QueryClose(L905)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | ExFrame | Frame |  | × |  | MouseMove(L309) |
| 1 ExFrame | Label34 | Label | EX風袋 | ○ |  |  |
| 1 ExFrame | ExGW | TextBox |  | ○ |  | Change(L288) |
| 1 ExFrame | Label33 | Label | EXGW | ○ |  |  |
| 1 ExFrame | ExMaterial1 | TextBox |  | ○ |  | Change(L291), MouseMove(L300) |
| 1 ExFrame | OutsideAtu | TextBox |  | ○ |  | Change(L466) |
| 1 ExFrame | Label35 | Label | 梱包外寸 | ○ |  |  |
| 1 ExFrame | OutsideHaba | TextBox |  | ○ |  | Change(L469) |
| 1 ExFrame | OutsideTake | TextBox |  | ○ |  | Change(L472) |
| 1 ExFrame | Label39 | Label | 高さ | ○ |  |  |
| 1 ExFrame | Label40 | Label | 奥行 | ○ |  |  |
| 1 ExFrame | Label41 | Label | 長さ | ○ |  |  |
| 1 ExFrame | ExCaseNo | TextBox |  | ○ |  | Change(L285) |
| 1 ExFrame | Label42 | Label | EXｹｰｽNO | ○ |  |  |
| 1 ExFrame | ExMaterialSpin | SpinButton |  | ○ |  | SpinUp(L443), SpinDown(L455) |
| 1 ExFrame | ExMaterial2 | TextBox |  | × |  | Change(L294), MouseMove(L303) |
| 1 ExFrame | ExMaterial3 | TextBox |  | × |  | Change(L297), MouseMove(L306) |
| 1 ExFrame | DigitLabel | Label |  | ○ |  |  |
| 0 | AinButt | CommandButton | Icas出力 | ○ |  | MouseMove(L925), Click(L932) |
| 0 | Frame2 | Frame |  | ○ |  |  |
| 1 Frame2 | TAKE | Label |  | ○ |  |  |
| 1 Frame2 | Label14 | Label | ロットNO | ○ |  |  |
| 1 Frame2 | ZAI | Label |  | ○ |  |  |
| 1 Frame2 | TYO | Label |  | ○ |  |  |
| 1 Frame2 | ATU | Label |  | ○ |  |  |
| 1 Frame2 | HABA | Label |  | ○ |  |  |
| 1 Frame2 | Label7 | Label | 検査NO | ○ |  |  |
| 1 Frame2 | Label9 | Label | 材質 | ○ |  |  |
| 1 Frame2 | Label10 | Label | 調質 | ○ |  |  |
| 1 Frame2 | Label11 | Label | 厚 | ○ |  |  |
| 1 Frame2 | Label12 | Label | 巾 | ○ |  |  |
| 1 Frame2 | Label13 | Label | 丈 | ○ |  |  |
| 1 Frame2 | LOT | Label |  | ○ |  |  |
| 1 Frame2 | KENNO | Label |  | ○ |  |  |
| 1 Frame2 | OrderNo | Label |  | ○ |  |  |
| 1 Frame2 | Label45 | Label | ｵｰﾀﾞｰNO | ○ |  |  |
| 1 Frame2 | AdjustLabel | Label | 調整NO | ○ |  |  |
| 1 Frame2 | AdjustNo | Label |  | ○ |  |  |
| 1 Frame2 | TMin | Label |  | ○ |  |  |
| 1 Frame2 | Label22 | Label | 厚Min | ○ |  |  |
| 1 Frame2 | AMin | Label |  | ○ |  |  |
| 1 Frame2 | TMax | Label |  | ○ |  |  |
| 1 Frame2 | AMax | Label |  | ○ |  |  |
| 1 Frame2 | HMax | Label |  | ○ |  |  |
| 1 Frame2 | HMin | Label |  | ○ |  |  |
| 1 Frame2 | Label24 | Label | 丈Min | ○ |  |  |
| 1 Frame2 | Label23 | Label | 巾Min | ○ |  |  |
| 1 Frame2 | Label25 | Label | 厚Max | ○ |  |  |
| 1 Frame2 | Label26 | Label | 丈Max | ○ |  |  |
| 1 Frame2 | Label27 | Label | 巾Max | ○ |  |  |
| 1 Frame2 | Label47 | Label | 測定ﾃﾞｰﾀ | ○ |  |  |
| 1 Frame2 | KenLaw | Label |  | ○ |  |  |
| 1 Frame2 | Label46 | Label | 検査法 | ○ |  |  |
| 0 | Label31 | Label | Auto | ○ |  |  |
| 0 | Frame3 | Frame |  | ○ |  | MouseMove(L579), Click(L583) |
| 1 Frame3 | Remain | TextBox | 値=0 | ○ |  | BeforeUpdate(L541), Change(L545), MouseMove(L574) |
| 1 Frame3 | COU1 | TextBox |  | ○ |  | Change(L279) |
| 1 Frame3 | Label6 | Label | 梱包NO | ○ |  |  |
| 1 Frame3 | MAILabel | Label | 梱包枚数 | ○ |  |  |
| 1 Frame3 | Label17 | Label | 包数 | ○ |  |  |
| 1 Frame3 | Label29 | Label | 引当NO | ○ |  |  |
| 1 Frame3 | KONNO | TextBox |  | ○ |  | BeforeUpdate(L375), Change(L379) |
| 1 Frame3 | MAI1 | TextBox |  | ○ |  | Change(L387) |
| 1 Frame3 | Label30 | Label | 出荷.ｽﾄｱ:0,1 | ○ |  |  |
| 1 Frame3 | PileMatch | TextBox |  | ○ |  | Change(L476), KeyPress(L519) |
| 1 Frame3 | PileMatchNo | TextBox |  | ○ |  | Change(L533) |
| 1 Frame3 | Label32 | Label | 積合せｽﾀｰﾄｴﾝﾄﾞ:1,9 | ○ |  |  |
| 1 Frame3 | Label28 | Label | 山数 | ○ |  |  |
| 1 Frame3 | Mount_Count | Label | 1 | ○ |  |  |
| 1 Frame3 | Material1 | TextBox |  | ○ |  | Change(L411) |
| 1 Frame3 | MaterialLabel | Label | 風袋 | ○ |  |  |
| 1 Frame3 | GW | TextBox |  | ○ |  | Change(L313) |
| 1 Frame3 | GWLabel | Label | 総重量 | ○ |  |  |
| 1 Frame3 | HikiNo | ComboBox |  | ○ |  | Change(L317) |
| 1 Frame3 | HIkiNoUnder | TextBox | 値=00 | × |  | Change(L368) |
| 1 Frame3 | COU2 | TextBox |  | × |  | Change(L282) |
| 1 Frame3 | MAI2 | TextBox |  | × |  | Change(L390) |
| 1 Frame3 | Material2 | TextBox |  | × |  | Change(L414) |
| 1 Frame3 | Material3 | TextBox |  | × |  | Change(L417) |
| 1 Frame3 | MAISpin | SpinButton |  | ○ |  | SpinUp(L393), SpinDown(L403) |
| 1 Frame3 | MaterialSpin | SpinButton |  | ○ |  | SpinUp(L420), SpinDown(L432) |
| 1 Frame3 | L1 | CommandButton | 伝送高さ測定 | × |  | Click(L598) |
| 1 Frame3 | L2 | Label |  | ○ |  | Click(L382) |
| 1 Frame3 | Inventory | Label |  | ○ |  |  |
| 1 Frame3 | K_Height_Label | Label |  | × |  |  |
| 1 Frame3 | K_Height_4_Label | Label |  | × |  |  |
| 1 Frame3 | K_Height_2_Label | Label |  | × |  |  |
| 1 Frame3 | K_Height_3_Label | Label |  | × |  |  |
| 1 Frame3 | Frame5 | Frame |  | ○ |  |  |
| 2 Frame5 | Ai | CheckBox | 合紙 (初期値=0) | ○ |  | BeforeUpdate(L73) |
| 2 Frame5 | Vab | CheckBox | VCAB (初期値=0) | ○ |  | BeforeUpdate(L98) |
| 2 Frame5 | Aisi | Label |  | ○ |  |  |
| 2 Frame5 | VCAB | Label |  | ○ |  |  |
| 2 Frame5 | CommandButton10 | CommandButton | 外注出荷 | ○ |  | Click(L107) |
| 2 Frame5 | GaityuuText | TextBox |  | × |  | BeforeUpdate(L49) |
| 2 Frame5 | Ga | CheckBox | 個装外装紙 (初期値=0) | ○ |  | BeforeUpdate(L88) |
| 2 Frame5 | Gaisousi | Label |  | ○ |  |  |
| 1 Frame3 | Label36 | Label | End区分:0,1,9 | × |  |  |
| 1 Frame3 | Kflag | TextBox |  | × |  | Change(L372) |
| 1 Frame3 | K_Height_6_Label | Label |  | × |  |  |
| 1 Frame3 | K_Height_5_Label | Label |  | × |  |  |
| 0 | Label37 | Label | Manual | ○ |  |  |
| 0 | ExportLabel | Label | Export | × |  |  |
| 0 | CommandButton2 | CommandButton | Reset | ○ |  | Click(L127), MouseMove(L178) |
| 0 | Frame4 | Frame |  | × |  |  |
| 1 Frame4 | MaiFlag | Label | 1 | ○ |  |  |
| 1 Frame4 | Label52 | Label | EX風袋Boxが１つ以上の時　”2”,”3” | ○ |  |  |
| 1 Frame4 | Label50 | Label | 風袋Boxが１つ以上の時　”2”,”3” | ○ |  |  |
| 1 Frame4 | Label51 | Label | 梱包数Boxが１つ以上の時　”2” | ○ |  |  |
| 1 Frame4 | ExMaterialFlag | Label | 1 | ○ |  |  |
| 1 Frame4 | MaterialFlag | Label | 1 | ○ |  |  |
| 1 Frame4 | TextBox60 | TextBox | 値=DT11001+0000037.000M | ○ |  |  |
| 1 Frame4 | LineName | Label |  | ○ |  |  |
| 1 Frame4 | Ngflag | Label |  | ○ |  |  |
| 1 Frame4 | Label54 | Label | 処理停止用 | ○ |  |  |
| 1 Frame4 | CommandButton3 | CommandButton | inputTest | ○ |  | Click(L185) |
| 1 Frame4 | CommandButton4 | CommandButton | resultTest | ○ |  | Click(L190) |
| 1 Frame4 | CommandButton5 | CommandButton | ｱﾌﾟﾘ起動 | ○ |  | Click(L194) |
| 1 Frame4 | DBflag | Label |  | ○ |  |  |
| 1 Frame4 | Label56 | Label | DB検索確認用 | ○ |  |  |
| 1 Frame4 | CommandButton1 | CommandButton | 日報Test | ○ |  | Click(L116) |
| 1 Frame4 | CommandButton7 | CommandButton | OutPutTest | ○ |  | Click(L202) |
| 1 Frame4 | Youto | Label | 用途ｺｰﾄﾞ | ○ |  |  |
| 1 Frame4 | CommandButton9 | CommandButton | 計算Test | ○ |  | Click(L250) |
| 0 | LoadTimeText | Label |  | ○ |  |  |
| 0 | B1 〜 B5（5個） | CommandButton | (空) | ○ |  |  |
| 0 | Re_hostsend | Label |  | ○ |  |  |
| 0 | OutCondition | Label |  | ○ |  |  |
| 0 | StartTB | TextBox |  | ○ |  | KeyDown(L697), Enter(L748), Exit(L752) |
| 0 | CommandButton6 | CommandButton | Close | ○ |  | Click(L198) |
| 0 | CommandButton8 | CommandButton | UFGW | ○ |  | Click(L220), MouseMove(L243) |
| 0 | GwOpen | Label |  | ○ |  |  |
| 0 | K_Height_2 | TextBox |  | × |  |  |
| 0 | K_Height_4 | TextBox |  | × |  |  |
| 0 | K_Height | TextBox |  | × |  |  |
| 0 | K_Height_3 | TextBox |  | × |  |  |
| 0 | Frame6 | Frame |  | × |  |  |
| 1 Frame6 | AiText | TextBox |  | ○ |  | Change(L62) |
| 1 Frame6 | GaText | TextBox |  | ○ |  | Change(L65) |
| 1 Frame6 | VcText | TextBox |  | ○ |  | Change(L69) |
| 0 | Frame7 | Frame |  | ○ |  |  |
| 0 | K_Height_6 | TextBox |  | × |  |  |
| 0 | K_Height_5 | TextBox |  | × |  |  |
- イベントプロシージャが無いCommandButton（他所から動的操作 or 未使用）: B1, B2, B3, B4, B5

## UFCask（コード: `UFCask.frm.bas`）

- フォームCaption(設計時): "2026.05.29_" / ClientWidth×Height(twip): 11565×5520 / StartUpPosition: 2  '画面の中央 / ShowModal: (既定=True) / コントロール数: 67
- フォームイベント: Activate(L777), Click(L811), Initialize(L817), QueryClose(L875)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | シート発行 | Frame | シート発行 | ○ |  |  |
| 1 シート発行 | CommandButton1 | CommandButton | ｷｬｽｸ用ｼｰﾄ発行 | ○ |  | Click(L41) |
| 0 | Frame1 | Frame | 印刷 & ｻｰﾊﾞｰ保存 | ○ |  |  |
| 1 Frame1 | CommandButton2 | CommandButton | 印刷 | ○ |  | Click(L110) |
| 0 | Frame2 | Frame | MINMAX | ○ |  |  |
| 1 Frame2 | Label1 | Label | Start板番号 | ○ |  |  |
| 1 Frame2 | Label2 | Label | End板番号 | ○ |  |  |
| 1 Frame2 | MINText | TextBox |  | ○ |  | KeyPress(L388) |
| 1 Frame2 | MAXText | TextBox |  | ○ |  |  |
| 1 Frame2 | CommandButton3 | CommandButton | 決定 | ○ |  | Click(L182) |
| 0 | Frame4 | Frame | LotNo | ○ |  |  |
| 1 Frame4 | LotBox | TextBox |  | ○ |  | KeyPress(L361), Change(L374) |
| 1 Frame4 | KenNo | TextBox |  | ○ |  |  |
| 1 Frame4 | Label5 | Label | 検査番号： | ○ |  |  |
| 0 | Frame5 | Frame | 処理 | ○ |  |  |
| 1 Frame5 | CommandButton12 | CommandButton | 枚数読上げ | ○ |  | Click(L85) |
| 1 Frame5 | Read | Label | OFF | ○ |  |  |
| 1 Frame5 | CommandButton9 | CommandButton | 巾丈測定  | ○ |  | Click(L329) |
| 1 Frame5 | Inspection | Label | OFF | ○ |  |  |
| 1 Frame5 | CommandButton8 | CommandButton | 処理再開 | ○ |  | Click(L324) |
| 1 Frame5 | CommandButton5 | CommandButton | Back space | ○ |  | Click(L245) |
| 1 Frame5 | CommandButton16 | CommandButton | 外観セルを広げる | ○ |  | Click(L103) |
| 0 | Frame10 | Frame | 入力必須欄 | ○ |  |  |
| 1 Frame10 | Frame9 | Frame | 対角公差 | ○ |  |  |
| 2 Frame9 | Tai | TextBox |  | ○ |  |  |
| 1 Frame10 | Frame6 | Frame | 検査員名 | ○ |  |  |
| 2 Frame6 | Kname | TextBox |  | ○ |  |  |
| 1 Frame10 | Frame7 | Frame | ﾗｲﾝ.ｺﾝﾍﾞｯｸｽNO | ○ |  |  |
| 2 Frame7 | Cno | TextBox |  | ○ |  |  |
| 1 Frame10 | Frame8 | Frame | ﾗｲﾝ.ﾏｲｸﾛNO.ﾏｲｸﾛｼﾘｱﾙNO | ○ |  |  |
| 2 Frame8 | Mno | TextBox |  | ○ |  |  |
| 2 Frame8 | SeriNo | TextBox |  | ○ |  | KeyPress(L716) |
| 2 Frame8 | Label3 | Label | ﾗｲﾝ　ﾏｲｸﾛNO | ○ |  |  |
| 2 Frame8 | Label4 | Label | シリアルNO(S/N不要) | ○ |  | MouseMove(L712) |
| 0 | ComboBox1 | ComboBox |  | ○ |  | KeyPress(L18), Change(L401) |
| 0 | Frame11 | Frame | 判定flag | ○ |  |  |
| 1 Frame11 | Hflag | Label | on | ○ |  |  |
| 0 | SB | SpinButton |  | ○ |  | SpinDown(L727), SpinUp(L731) |
| 0 | CommandButton10 | CommandButton | 印刷テスト | ○ |  |  |
| 0 | Frame24 | Frame | 異常値flag | ○ |  |  |
| 1 Frame24 | AbnorFlag | Label | 0 | ○ |  |  |
| 0 | CommandButton11 | CommandButton | 保存Test | ○ |  | Click(L81) |
| 0 | Frame3 | Frame | 保存 | ○ |  |  |
| 1 Frame3 | CommandButton4 | CommandButton | ｻｰﾊﾞｰ保存 | ○ |  | Click(L234) |
| 0 | CommandButton14 | CommandButton | 読上げ | ○ |  | Click(L94) |
| 0 | TextBox1 | TextBox | 値=DT10002+0000024.959M | ○ |  |  |
| 0 | TextBox2 | TextBox | 値=#L1528.22 | ○ |  |  |
| 0 | Frame27 | Frame | ﾁｪﾝｼﾞｶｳﾝﾄ | ○ |  |  |
| 1 Frame27 | Cnt | Label | 0 | ○ |  |  |
| 0 | Frame28 | Frame | Undo | ○ |  |  |
| 1 Frame28 | UndoCount | Label | 0 | ○ |  |  |
| 0 | LineName | Label |  | ○ |  |  |
| 0 | MultiPage1 | MultiPage |  | ○ |  |  |
| 1 MultiPage1 | Page1 | Page/Form | (タブ表示名)ｷｬｽｸ専用校正 | ○ |  |  |
| 2 Page1 | StartTB | TextBox |  | ○ |  | AfterUpdate(L736) |
| 2 Page1 | Convex | Label | 未 | ○ |  |  |
| 2 Page1 | Micro | Label | 未 | ○ |  |  |
| 2 Page1 | L2 | Label | ｺﾝﾍﾞｯｸｽｾﾞﾛ点 | ○ |  |  |
| 2 Page1 | L1 | Label | ﾏｲｸﾛｾﾞﾛ点 | ○ |  |  |
| 2 Page1 | Convex_1000 | Label | 未 | ○ |  |  |
| 2 Page1 | L3 | Label | ｺﾝﾍﾞｯｸｽ1000mm | ○ |  |  |
| 2 Page1 | CommandButton25 | CommandButton | UFsize | ○ |  | Click(L176) |
| 2 Page1 | CommandButton6 | CommandButton | 連番最終チェック | ○ |  | Click(L308) |
| 0 | CommandButton15 | CommandButton | 校正読み込み | ○ |  | Click(L99) |
| 0 | CountSt | Label |  | ○ |  |  |
| 0 | Frame29 | Frame | Loadflag | ○ |  |  |
| 1 Frame29 | Loadtext | Label | on | ○ |  |  |

- **コードにイベントがあるがフォーム上に同名コントロールが無いもの（未使用コード候補）**: maxtex_KeyPress(L29)
- イベントプロシージャが無いCommandButton（他所から動的操作 or 未使用）: CommandButton10

## UFConfi（コード: `UFConfi.frm.bas`）

- フォームCaption(設計時): "2024.7.26_Ver" / ClientWidth×Height(twip): 16245×11190 / StartUpPosition: 2  '画面の中央 / ShowModal: 0   'False / コントロール数: 41
- フォームイベント: Layout(L91), Activate(L101), Initialize(L144), QueryClose(L159)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | MP1 | MultiPage |  | ○ |  |  |
| 1 MP1 | P1 | Page/Form | (タブ表示名)公差外 | ○ |  |  |
| 2 P1 | Label1 | Label | 検査ﾃﾞｰﾀに公差外有り！ | ○ |  |  |
| 2 P1 | Label2 | Label | 公差外を廃棄する場合は | ○ |  |  |
| 2 P1 | Label3 | Label | 公差外 | ○ |  |  |
| 2 P1 | Label4 | Label | と入力して送信ボタン | ○ |  |  |
| 2 P1 | Label5 | Label | 再測定する場合は | ○ |  |  |
| 2 P1 | Label6 | Label | 再測定 | ○ |  |  |
| 2 P1 | TB1 | TextBox |  | ○ |  |  |
| 2 P1 | Label7 | Label | と入力して送信ボタン | ○ |  |  |
| 2 P1 | Label8 | Label | 公差外を梱包する場合は(出荷前提) | ○ |  |  |
| 2 P1 | Label9 | Label | 許容 | ○ |  |  |
| 2 P1 | Label10 | Label | と入力して送信ボタン | ○ |  |  |
| 2 P1 | NewHB1 | CommandButton | 自動入力 | ○ |  | Click(L274) |
| 2 P1 | NewHB2 | CommandButton | 自動入力 | ○ |  | Click(L277) |
| 2 P1 | NewHB3 | CommandButton | 自動入力 | ○ |  | Click(L283) |
| 2 P1 | CommandButton1 | CommandButton | 送信 | ○ |  | Click(L171) |
| 2 P1 | return_MotoData | Label |  | ○ |  |  |
| 2 P1 | Label13 | Label | 対象データ | ○ |  |  |
| 2 P1 | NewHB4 | CommandButton | 自動入力 | ○ |  | Click(L280) |
| 2 P1 | Label14 | Label | 公差外を保留とする場合は | ○ |  |  |
| 2 P1 | Label15 | Label | 保留 | ○ |  |  |
| 2 P1 | Label16 | Label | と入力して送信ボタン | ○ |  |  |
| 2 P1 | List2 | ListBox |  | × |  |  |
| 1 MP1 | P2 | Page/Form | (タブ表示名)理由選択 | × |  |  |
| 2 P2 | KenName | Label |  | ○ |  |  |
| 2 P2 | L1 | Label | 複数選択が可能 | ○ |  |  |
| 2 P2 | L2 | Label | 選択/選択解除はクリック | ○ |  |  |
| 2 P2 | Label11 | Label | 現検査員 | ○ |  |  |
| 2 P2 | その他 | TextBox |  | ○ |  |  |
| 2 P2 | CommandButton2 | CommandButton | 決定 | ○ |  | Click(L16) |
| 2 P2 | L3 | Label | 再測定及びその他の理由 | ○ |  |  |
| 2 P2 | MotoData | Label |  | ○ |  |  |
| 2 P2 | L4 | Label | 対象データ | ○ |  |  |
| 2 P2 | List1 | ListBox |  | ○ |  |  |
| 0 | Kpop | TextBox |  | ○ |  |  |
| 0 | LineName | Label |  | ○ |  |  |
| 0 | TR | Label | TR | ○ |  |  |
| 0 | TC | Label | TC | ○ |  |  |
| 0 | Knum | TextBox |  | ○ |  |  |
| 0 | DelC | Label |  | ○ |  |  |

## UFMitsubishi（コード: `UFMitsubishi.frm.bas`）

- フォームCaption(設計時): "2026.05.29" / ClientWidth×Height(twip): 10215×5865 / StartUpPosition: 2  '画面の中央 / ShowModal: (既定=True) / コントロール数: 149
- フォームイベント: Activate(L2194), Click(L2270), Initialize(L2294), MouseDown(L2337), QueryClose(L2344)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | Label4 | Label | flagonを印刷フラグに入力でテスト環境下になる | ○ |  |  |
| 0 | MLP | MultiPage |  | ○ |  | Click(L1833), MouseMove(L1894) |
| 1 MLP | P1 | Page/Form | (タブ表示名)メイン | ○ |  |  |
| 2 P1 | SB | SpinButton |  | ○ |  | SpinDown(L2183), SpinUp(L2188) |
| 2 P1 | Frame4 | Frame | LotNo | ○ |  | Click(L1773) |
| 3 Frame4 | Label3 | Label | 検査番号： | ○ |  |  |
| 3 Frame4 | KenNo | TextBox |  | ○ |  | KeyPress(L1799) |
| 3 Frame4 | LotBox | TextBox |  | ○ |  | KeyPress(L1811), Change(L1822) |
| 3 Frame4 | Label19 | Label | 用途ｺｰﾄﾞ： | ○ |  |  |
| 3 Frame4 | YoutoNo | TextBox |  | ○ |  |  |
| 2 P1 | Frame2 | Frame | MINMAX | ○ |  | Click(L276) |
| 3 Frame2 | Label1 | Label | Start板番号 | ○ |  |  |
| 3 Frame2 | Label2 | Label | End板番号 | ○ |  |  |
| 3 Frame2 | MINText | TextBox |  | ○ |  |  |
| 3 Frame2 | MAXText | TextBox |  | ○ |  |  |
| 3 Frame2 | CommandButton3 | CommandButton | 決定 | ○ |  | Click(L1605) |
| 3 Frame2 | Frame14 | Frame | 検査法 | ○ |  |  |
| 4 Frame14 | Inspection2 | OptionButton | L2 (初期値=0) | ○ |  |  |
| 4 Frame14 | Inspection3 | OptionButton | M6 (初期値=0) | ○ |  |  |
| 4 Frame14 | Inspection1 | OptionButton | L1 (初期値=0) | ○ |  |  |
| 2 P1 | シート発行 | Frame | シート発行 | ○ |  | Click(L2160) |
| 3 シート発行 | CommandButton1 | CommandButton | ｼｰﾄ発行 | ○ |  | Click(L1250) |
| 2 P1 | Frame5 | Frame | 処理 | ○ |  |  |
| 3 Frame5 | CommandButton16 | CommandButton | 処理再開 | ○ |  | Click(L1499) |
| 3 Frame5 | CommandButton17 | CommandButton | ｾﾙ拡大 | ○ |  | Click(L1504) |
| 2 P1 | Frame6 | Frame | 検査員名 | ○ |  | Click(L293) |
| 3 Frame6 | Kname | TextBox |  | ○ |  |  |
| 2 P1 | Frame8 | Frame | 保存　＆　印刷 | ○ |  |  |
| 3 Frame8 | Mitsubishi | Frame | 三菱専用印刷 | × |  |  |
| 4 Mitsubishi | CommandButton2 | CommandButton | 三菱印刷 | ○ |  | Click(L1515) |
| 3 Frame8 | Normal | Frame | 検査シート印刷 | × |  |  |
| 4 Normal | CommandButton7 | CommandButton | 検査印刷 | ○ |  | Click(L1654) |
| 3 Frame8 | Frame29 | Frame | 保存のみ | × |  |  |
| 4 Frame29 | CommandButton4 | CommandButton | ｻｰﾊﾞｰ保存 | ○ |  | Click(L1632) |
| 3 Frame8 | FiNi | Frame | 5N専用印刷 | × |  |  |
| 4 FiNi | CommandButton15 | CommandButton | 5N印刷 | ○ |  | Click(L1363) |
| 3 Frame8 | H494 | CheckBox | 用途ｺｰﾄﾞH494 (初期値=0) | ○ |  |  |
| 3 Frame8 | Label20 | Label | H494かつ幅300未満 | ○ |  |  |
| 2 P1 | FMH | Frame |  | ○ |  |  |
| 2 P1 | HindBo | CommandButton | 検査頻度確認 | ○ |  | Click(L1791) |
| 2 P1 | 一般 | OptionButton | 一般 (初期値=0) | ○ |  | Click(L2178) |
| 2 P1 | FiveOP | OptionButton | 5N (初期値=0) | ○ |  | Click(L271) |
| 2 P1 | CommandButton24 | CommandButton | 品質一覧 | ○ |  | Click(L232) |
| 2 P1 | CommandButton25 | CommandButton | UFsize | ○ |  | Click(L265) |
| 1 MLP | P2 | Page/Form | (タブ表示名)H494 | × |  |  |
| 2 P2 | Image1 | Image |  | ○ |  |  |
| 1 MLP | P3 | Page/Form | (タブ表示名)検入高さ | × |  |  |
| 2 P3 | Label5 | Label | PSW山毎の上下を板厚/巾/丈/対角/歪の測定　外観全数　 | ○ |  |  |
| 2 P3 | Label6 | Label | 異常発生時,詳細を検査票に記入し現品貼り付け | ○ |  |  |
| 2 P3 | Label8 | Label | PSW山毎に梱包 | ○ |  |  |
| 2 P3 | Label9 | Label | NO1の梱包に外装紙で包んだTP添付 | ○ |  |  |
| 2 P3 | Label10 | Label | 深さのある欠陥不可 | ○ |  |  |
| 2 P3 | Label11 | Label | 高さのあるバリ不可 | ○ |  |  |
| 2 P3 | Label12 | Label | 梱包仕様書1P1232-01 | ○ |  |  |
| 2 P3 | Label13 | Label | 切削屑以外引き取りの為廃棄しない | ○ |  |  |
| 2 P3 | Label14 | Label | 巾反り丈反り共に1.0以下 | ○ |  |  |
| 2 P3 | Label15 | Label | 巾公差±3.0　丈公差±4.0　対角公差4.0以下 | ○ |  |  |
| 1 MLP | P4 | Page/Form | (タブ表示名)不良 | × |  |  |
| 2 P4 | Image2 | Image |  | ○ |  |  |
| 1 MLP | P3 | Page/Form | (タブ表示名) | × |  |  |
| 2 P3 | ViewCB | CommandButton | 確認View | ○ |  | Click(L1939) |
| 2 P3 | Label89 | Label | 不良要因選択 | ○ |  |  |
| 2 P3 | EBad | TextBox |  | ○ |  | Change(L484), MouseDown(L504) |
| 2 P3 | SBad | TextBox |  | ○ |  | Change(L472), MouseDown(L488) |
| 2 P3 | UBad | TextBox |  | ○ |  | Change(L475), MouseDown(L492) |
| 2 P3 | DBad | TextBox |  | ○ |  | Change(L481), MouseDown(L500) |
| 2 P3 | CBad | TextBox |  | ○ |  | Change(L478), MouseDown(L496) |
| 2 P3 | ListReflect | CommandButton | 反映 | ○ |  | Click(L390) |
| 2 P3 | Label90 | Label | 単位(枚) | ○ |  |  |
| 2 P3 | Cle | CommandButton | ｸﾘｱ | ○ |  | Click(L59) |
| 2 P3 | NgIm | Image |  | ○ |  | Click(L445) |
| 2 P3 | FBad | TextBox |  | × |  |  |
| 2 P3 | DefectFactorLB | ListBox |  | ○ |  |  |
| 1 MLP | P2 | Page/Form | (タブ表示名) | × |  |  |
| 2 P2 | Label87 | Label | 山数選択 | ○ |  |  |
| 2 P2 | KenHBox1 | TextBox |  | × |  | AfterUpdate(L331) |
| 2 P2 | KenHL1 | Label | ｍｍ | × |  |  |
| 2 P2 | KenHL2 | Label | ｍｍ | × |  |  |
| 2 P2 | KenHBox2 | TextBox |  | × |  | AfterUpdate(L339) |
| 2 P2 | KenHL3 | Label | ｍｍ | × |  |  |
| 2 P2 | KenHBox3 | TextBox |  | × |  | AfterUpdate(L347) |
| 2 P2 | KenHL4 | Label | ｍｍ | × |  |  |
| 2 P2 | KenHBox4 | TextBox |  | × |  | AfterUpdate(L355) |
| 2 P2 | KenHL5 | Label | ｍｍ | × |  |  |
| 2 P2 | KenHBox5 | TextBox |  | × |  | AfterUpdate(L363) |
| 2 P2 | KenHL6 | Label | ｍｍ | × |  |  |
| 2 P2 | KenHBox6 | TextBox |  | × |  | AfterUpdate(L371) |
| 2 P2 | HL1 〜 HL6（6個） | Label | (空) | ○ |  |  |
| 2 P2 | KenHBox7 | TextBox |  | × |  | AfterUpdate(L379) |
| 2 P2 | KenHL7 | Label | 枚 | × |  |  |
| 2 P2 | 高さ計算式 | Label |  | ○ |  |  |
| 2 P2 | YamaLB | ListBox |  | ○ |  | Click(L2153) |
| 1 MLP | P_H494 | Page/Form | (タブ表示名) | × |  |  |
| 2 P_H494 | Frame30 | Frame | LotNo | ○ |  |  |
| 3 Frame30 | Label91 | Label | 検査番号： | ○ |  |  |
| 3 Frame30 | H494_KenNo | TextBox |  | ○ |  |  |
| 3 Frame30 | H494_Lot | TextBox |  | ○ |  |  |
| 3 Frame30 | Label92 | Label | 用途ｺｰﾄﾞ： | ○ |  |  |
| 3 Frame30 | H494_Youto | TextBox |  | ○ |  |  |
| 2 P_H494 | Frame31 | Frame | MINMAX | ○ |  |  |
| 3 Frame31 | Label93 | Label | Start板番号 | ○ |  |  |
| 3 Frame31 | Label94 | Label | End板番号 | ○ |  |  |
| 3 Frame31 | H494_MINText | TextBox |  | ○ |  |  |
| 3 Frame31 | H494_MAXText | TextBox |  | ○ |  |  |
| 3 Frame31 | CommandButton19 | CommandButton | 決定 | ○ |  | Click(L99) |
| 3 Frame31 | Frame32 | Frame | 検査法 | ○ |  |  |
| 4 Frame32 | op2 | OptionButton | L2 (初期値=0) | ○ |  |  |
| 4 Frame32 | op3 | OptionButton | M6 (初期値=0) | ○ |  |  |
| 4 Frame32 | op1 | OptionButton | L1 (初期値=0) | ○ |  |  |
| 2 P_H494 | Frame33 | Frame | シート発行 | ○ |  |  |
| 3 Frame33 | ws_out | CommandButton | ｼｰﾄ発行 | ○ |  |  |
| 2 P_H494 | Frame34 | Frame | 処理 | ○ |  |  |
| 3 Frame34 | CommandButton21 | CommandButton | 処理再開 | ○ |  | Click(L126) |
| 3 Frame34 | CommandButton22 | CommandButton | 外観ｾﾙ拡大 | ○ |  | Click(L131) |
| 2 P_H494 | Frame35 | Frame | 検査員名 | ○ |  |  |
| 3 Frame35 | H494_Kname | TextBox |  | ○ |  |  |
| 2 P_H494 | Frame36 | Frame | 検査シート印刷 | ○ |  |  |
| 3 Frame36 | CommandButton23 | CommandButton | 検査印刷 | ○ |  | Click(L138) |
| 0 | Frame37 | Frame |  | × |  |  |
| 1 Frame37 | ComboBox1 | ComboBox |  | ○ |  | KeyPress(L583), Change(L596) |
| 1 Frame37 | Frame9 | Frame | flag | ○ |  |  |
| 2 Frame9 | Work | Label | ON | ○ |  |  |
| 2 Frame9 | Label17 | Label | 印刷フラグ　flagon.”” | ○ |  |  |
| 2 Frame9 | Label16 | Label | 三菱　一般切替フラグ　ON.OFF | ○ |  |  |
| 2 Frame9 | Inflag | Label |  | ○ |  |  |
| 2 Frame9 | PackNo_on_off | Label | off | ○ |  |  |
| 1 Frame37 | Frame24 | Frame | 異常値flag | ○ |  |  |
| 2 Frame24 | AbnorFlag | Label | 0 | ○ |  |  |
| 1 Frame37 | TextBox1 | TextBox | 値=DT10002+0000010.512M | ○ |  |  |
| 1 Frame37 | TextBox2 | TextBox | 値=#L1528.22 | ○ |  |  |
| 1 Frame37 | Frame27 | Frame | ﾁｪﾝｼﾞｶｳﾝﾄ | ○ |  |  |
| 2 Frame27 | Cnt | Label | 0 | ○ |  |  |
| 1 Frame37 | Frame28 | Frame | Undo | ○ |  |  |
| 2 Frame28 | UndoCount | Label | 0 | ○ |  |  |
| 1 Frame37 | TextBox3 | TextBox |  | ○ |  |  |
| 1 Frame37 | LineName | Label |  | ○ |  |  |
| 1 Frame37 | CommandButton18 | CommandButton | 全テスト | ○ |  | Click(L73) |
| 1 Frame37 | Choice | TextBox |  | ○ |  | Change(L509) |
| 1 Frame37 | Label18 | Label | 検査作業 | ○ |  |  |
| 1 Frame37 | CommandButton5 | CommandButton | 印刷転送テスト | ○ |  | Click(L1643) |
| 1 Frame37 | CMB | CommandButton | 作業選択 | ○ |  | Click(L517) |
| 1 Frame37 | CommandButton6 | CommandButton | 新規保存 | ○ |  | Click(L1647) |
| 1 Frame37 | CommandButton10 | CommandButton | ５N転送テスト | ○ |  | Click(L1289) |
| 1 Frame37 | WorkFlag | Label |  | ○ |  |  |

- **コードにイベントがあるがフォーム上に同名コントロールが無いもの（未使用コード候補）**: commandbutton14_DblClick(L1298)
- イベントプロシージャが無いCommandButton（他所から動的操作 or 未使用）: ws_out

## UFMsgBox（コード: `UFMsgBox.frm.bas`）

- フォームCaption(設計時): (なし) / ClientWidth×Height(twip): 10020×1995 / StartUpPosition: 2  '画面の中央 / ShowModal: (既定=True) / コントロール数: 8
- フォームイベント: Layout(L16), Activate(L24), Initialize(L35), QueryClose(L46)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | L1 | Label | Label1 | ○ |  |  |
| 0 | CommandButton1 | CommandButton | OK | ○ |  | Click(L73) |
| 0 | CommandButton2 | CommandButton | キャンセル | × |  |  |
| 0 | IM1 〜 IM4（4個） | Image | (空) | × |  |  |
| 0 | FlLabe | Label |  | ○ |  |  |
- イベントプロシージャが無いCommandButton（他所から動的操作 or 未使用）: CommandButton2

## UFNG（コード: `UFNG.frm.bas`）

- フォームCaption(設計時): (なし) / ClientWidth×Height(twip): 9945.001×8355.001 / StartUpPosition: 2  '画面の中央 / ShowModal: (既定=True) / コントロール数: 1
- フォームイベント: Activate(L16), QueryClose(L32)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | Image1 | Image |  | ○ |  |  |

## UFNameList（コード: `UFNameList.frm.bas`）

- フォームCaption(設計時): (なし) / ClientWidth×Height(twip): 17190×6555 / StartUpPosition: 2  '画面の中央 / ShowModal: (既定=True) / コントロール数: 16
- フォームイベント: Initialize(L55), Layout(L295), Activate(L304), QueryClose(L327)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | CB1 | CommandButton | 不一致ﾘｽﾄ | ○ |  | Click(L349) |
| 0 | Label1 | Label | 登録検査員と測定者が一致しません。 | ○ |  |  |
| 0 | Label2 | Label | 不正の為検査判定を停止します。 | ○ |  |  |
| 0 | Label3 | Label | 「組長」に連絡して対処の事。 | ○ |  |  |
| 0 | 表示 | Label |  | ○ |  | DblClick(L379) |
| 0 | Content | Label |  | ○ |  |  |
| 0 | Frame1 | Frame |  | ○ |  |  |
| 1 Frame1 | LV1 | ActiveX#2 |  | ○ |  | Enter(L130), Click(L133), KeyUp(L251) |
| 0 | FM3 | Frame | 特殊操作 | ○ |  |  |
| 1 FM3 | CommandButton1 | CommandButton | 選択行削除 | ○ |  | Click(L254) |
| 1 FM3 | CommandButton2 | CommandButton | 反映 | ○ |  | Click(L261) |
| 0 | PasswordTex | TextBox |  | × |  |  |
| 0 | Frame2 | Frame |  | ○ |  |  |
| 1 Frame2 | TextBox1 | TextBox |  | ○ |  | Change(L122) |
| 0 | CommandButton3 | CommandButton | 再転記テスト | × |  | Click(L278) |
| 0 | OnOff | CommandButton | VBA停止 | ○ |  | Click(L36) |

## UFOK（コード: `UFOK.frm.bas`）

- フォームCaption(設計時): (なし) / ClientWidth×Height(twip): 10440×9600.001 / StartUpPosition: 2  '画面の中央 / ShowModal: (既定=True) / コントロール数: 1
- フォームイベント: Activate(L15), QueryClose(L33)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | Image1 | Image |  | ○ |  |  |

## UFProgress（コード: `UFProgress.frm.bas`）

- フォームCaption(設計時): (なし) / ClientWidth×Height(twip): 7950×2010 / StartUpPosition: 1  'オーナー フォームの中央 / ShowModal: (既定=True) / コントロール数: 6
- フォームイベント: Activate(L15), Click(L20), Initialize(L23), Terminate(L29), QueryClose(L32)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | ProgressBar1 | ProgressBar(MSComctl) |  | ○ |  |  |
| 0 | CommandButton1 | CommandButton | ﾃﾞｰﾀ書込中....。 | ○ |  |  |
| 0 | Label1 〜 Label3（3個） | Label | (空) | ○ |  |  |
| 0 | ProgressLabel | Label |  | × |  |  |
- イベントプロシージャが無いCommandButton（他所から動的操作 or 未使用）: CommandButton1

## UFWorkerName（コード: `UFWorkerName.frm.bas`）

- フォームCaption(設計時): (なし) / ClientWidth×Height(twip): 9540.001×2235 / StartUpPosition: ? / ShowModal: (既定=True) / コントロール数: 4
- フォームイベント: Click(L10)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | L1 | Label |  | ○ |  |  |
| 0 | TextBox1 | TextBox | 値=現　検査員 | ○ |  |  |
| 0 | TwiceTap | TextBox | 値=2度押しﾃﾞｰﾀ消去 | × |  |  |
| 0 | L2 | Label |  | ○ |  |  |

## UF_PackNo（コード: `UF_PackNo.frm.bas`）

- フォームCaption(設計時): (なし) / ClientWidth×Height(twip): 17160×12630 / StartUpPosition: 2  '画面の中央 / ShowModal: (既定=True) / コントロール数: 35
- フォームイベント: Initialize(L102), Activate(L118), QueryClose(L188)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | CommandButton1 | CommandButton | ﾌﾞﾗｳｻﾞで確認 | × |  | Click(L19) |
| 0 | L1 | Label | Label1 | ○ |  |  |
| 0 | L2 | Label | Label1 | ○ |  |  |
| 0 | L5 | Label | Label1 | ○ |  |  |
| 0 | L4 | Label | Label1 | ○ |  |  |
| 0 | L3 | Label | Label1 | ○ |  |  |
| 0 | TBLoad | TextBox |  | × |  |  |
| 0 | FM1 | Frame |  | ○ |  |  |
| 1 FM1 | CB1 | CheckBox | 確認 (初期値=0) | ○ |  |  |
| 1 FM1 | CommandButton2 | CommandButton | 閉じる | ○ |  | Click(L25) |
| 0 | Hiki1 〜 Hiki5（5個） | Label | (空) | ○ |  |  |
| 0 | Ord1 〜 Ord5（5個） | Label | (空) | ○ |  |  |
| 0 | Pac1 〜 Pac5（5個） | Label | (空) | ○ |  |  |
| 0 | PcLot | TextBox |  | × |  |  |
| 0 | Label1 | Label | 引当No | ○ |  |  |
| 0 | Label2 | Label | オーダーNo | ○ |  |  |
| 0 | Label3 | Label | 包装仕様書No | ○ |  |  |
| 0 | SpinButton1 | SpinButton |  | × |  | SpinDown(L51), SpinUp(L76) |
| 0 | MultiPage1 | MultiPage |  | ○ |  |  |
| 1 MultiPage1 | Page1 | Page/Form | (タブ表示名)包装仕様_Ver1.0.2 | ○ |  |  |
| 2 Page1 | FMBar | Frame |  | ○ |  |  |
| 1 MultiPage1 | Page2 | Page/Form | (タブ表示名)Im隠し用 | × |  |  |
| 2 Page2 | PacIM | Image |  | × |  |  |

## UFcount（コード: `UFcount.frm.bas`）

- フォームCaption(設計時): "枚数表示" / ClientWidth×Height(twip): 13260×6810 / StartUpPosition: ? / ShowModal: (既定=True) / コントロール数: 10
- フォームイベント: Activate(L18), QueryClose(L42)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | Frame1 | Frame | 判定板NO | ○ |  |  |
| 1 Frame1 | ItanoClone | Label |  | ○ |  |  |
| 0 | TR | TextBox |  | ○ |  |  |
| 0 | TC | TextBox |  | ○ |  |  |
| 0 | Itano | TextBox |  | ○ |  | Change(L10) |
| 0 | Location | TextBox |  | ○ |  | Change(L14) |
| 0 | Maru | Image |  | × |  |  |
| 0 | Batu | Image |  | × |  |  |
| 0 | Frame2 | Frame | 判定位置 | ○ |  |  |
| 1 Frame2 | LocationClone | Label |  | ○ |  |  |

## UFdiagonal（コード: `UFdiagonal.frm.bas`）

- フォームCaption(設計時): "対角チェック" / ClientWidth×Height(twip): 15330×6255 / StartUpPosition: 2  '画面の中央 / ShowModal: 0   'False / コントロール数: 18
- フォームイベント: Activate(L54), Initialize(L63)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | Label10 | Label | 3　or　4 | ○ |  |  |
| 0 | CommandButton1 | CommandButton | 送信 | ○ |  | Click(L11) |
| 0 | TAI | TextBox |  | ○ |  |  |
| 0 | Label11 | Label | 4級公差 | ○ |  |  |
| 0 | Label12 | Label | 製板カード上の「直角度等級」 | ○ |  |  |
| 0 | Label13 | Label | 製品規格：書式番号Z501 | ○ |  |  |
| 0 | Label14 | Label | を入力してください。 | ○ |  |  |
| 0 | TAISA | Label |  | ○ |  |  |
| 0 | Image1 | Image |  | ○ |  |  |
| 0 | Label15 | Label | から外れました。　　　　　　　　　　 | ○ |  |  |
| 0 | TAIKOU | Label |  | ○ |  |  |
| 0 | TR | Label | TR | ○ |  |  |
| 0 | TC | Label | TC | ○ |  |  |
| 0 | PosiA | Label |  | × |  |  |
| 0 | PosiB | Label |  | × |  |  |
| 0 | Terg | TextBox |  | ○ |  |  |
| 0 | NumB | Label |  | × |  |  |
| 0 | NumA | Label |  | × |  |  |

## UF品質基準（コード: `UF品質基準.frm.bas`）

- フォームCaption(設計時): "品質一覧" / ClientWidth×Height(twip): 21450×10260 / StartUpPosition: 2  '画面の中央 / ShowModal: (既定=True) / コントロール数: 131
- フォームイベント: QueryClose(L201), Activate(L210), Initialize(L213)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | MP | MultiPage |  | ○ |  |  |
| 1 MP | Page1 | Page/Form | (タブ表示名)Q-S-QA-DIM-122,124,129 | ○ |  |  |
| 2 Page1 | Zai | TextBox |  | ○ |  | KeyPress(L311) |
| 2 Page1 | Label19 | Label | ｻｲｽﾞ | ○ |  |  |
| 2 Page1 | Label20 | Label | 材調質 | ○ |  |  |
| 2 Page1 | Label9 | Label | ﾗ | ○ |  |  |
| 2 Page1 | Label10 | Label | ﾗ | ○ |  |  |
| 2 Page1 | Label13 | Label | - | ○ |  |  |
| 2 Page1 | Label14 | Label | 得意先 | ○ |  |  |
| 2 Page1 | Label15 | Label | 送り先 | ○ |  |  |
| 2 Page1 | Label16 | Label | 納入先 | ○ |  |  |
| 2 Page1 | Toku | Label |  | ○ |  |  |
| 2 Page1 | Okur | Label |  | ○ |  |  |
| 2 Page1 | Noun | Label |  | ○ |  |  |
| 2 Page1 | Yout | Label |  | × |  |  |
| 2 Page1 | Label17 | Label | 用途 | ○ |  |  |
| 2 Page1 | Frame2 | Frame | 平坦度規格 | ○ |  |  |
| 3 Frame2 | Label24 | Label | 単位mm | ○ |  |  |
| 3 Frame2 | Label21 | Label | 巾歪 | ○ |  |  |
| 3 Frame2 | TaHiz1 | TextBox |  | ○ |  | Change(L198) |
| 3 Frame2 | HaHiz1 | TextBox |  | ○ |  | Change(L159) |
| 3 Frame2 | Label22 | Label | 丈歪 | ○ |  |  |
| 3 Frame2 | MG1 | TextBox |  | ○ |  | Change(L162) |
| 3 Frame2 | KousinL2 | Label |  | ○ |  |  |
| 3 Frame2 | Label64 | Label | 中歪 | ○ |  |  |
| 3 Frame2 | Naka1 | TextBox |  | ○ |  | Change(L192) |
| 3 Frame2 | Label65 | Label | 耳歪 | ○ |  |  |
| 3 Frame2 | Mimi1 | TextBox |  | ○ |  | Change(L165) |
| 3 Frame2 | Sansyou1 | TextBox |  | ○ |  | Change(L195) |
| 3 Frame2 | Label73 | Label | 1M当りｹﾞｰｼﾞ | ○ |  |  |
| 2 Page1 | Atu | TextBox |  | ○ |  |  |
| 2 Page1 | Haba | TextBox |  | ○ |  |  |
| 2 Page1 | Take | TextBox |  | ○ |  |  |
| 2 Page1 | Tyo | TextBox |  | ○ |  | KeyPress(L322) |
| 2 Page1 | YoutC | TextBox |  | ○ |  | KeyPress(L300) |
| 2 Page1 | L1 | Label | 検索中..... | × |  |  |
| 2 Page1 | Flag2 | Label |  | ○ |  |  |
| 2 Page1 | Flag1 | Label |  | ○ |  |  |
| 2 Page1 | KousinL1 | Label |  | ○ |  |  |
| 2 Page1 | AbnoName | TextBox |  | ○ |  |  |
| 2 Page1 | TabAb | TabStrip |  | ○ |  | Change(L249) |
| 2 Page1 | gradeTB | TextBox |  | ○ |  |  |
| 2 Page1 | AbnoTB | TextBox |  | ○ |  |  |
| 2 Page1 | AbnoLabel | Label |  | ○ |  |  |
| 2 Page1 | Label74 | Label | グレード | ○ |  |  |
| 1 MP | Page2 | Page/Form | (タブ表示名)Q-S-QA-SRF-128[別表1][別表2]　　 | × |  |  |
| 2 Page2 | Utake | TextBox |  | ○ |  |  |
| 2 Page2 | Label28 | Label | ｻｲｽﾞ | ○ |  |  |
| 2 Page2 | Label29 | Label | 材調質 | ○ |  |  |
| 2 Page2 | Label33 | Label | ﾗ | ○ |  |  |
| 2 Page2 | Label34 | Label | ﾗ | ○ |  |  |
| 2 Page2 | Label37 | Label | - | ○ |  |  |
| 2 Page2 | Label38 | Label | 用途 | ○ |  |  |
| 2 Page2 | KousinL4 | Label |  | ○ |  |  |
| 2 Page2 | Label58 | Label | 表面処理ｺｰﾄﾞ | ○ |  |  |
| 2 Page2 | Label59 | Label | ﾊﾟﾀｰﾝ判定 | ○ |  |  |
| 2 Page2 | Upater | Label |  | ○ |  |  |
| 2 Page2 | B1 | CommandButton | トビコミ | ○ |  | Click(L11) |
| 2 Page2 | SelectA | Label |   | ○ |  |  |
| 2 Page2 | B2 | CommandButton | キズ,スリキズ | ○ |  | Click(L17) |
| 2 Page2 | B4 | CommandButton | 突上げ,突下げ | ○ |  | Click(L29) |
| 2 Page2 | B7 | CommandButton | ロールキズ | ○ |  | Click(L47) |
| 2 Page2 | B5 | CommandButton | 汚れ,汚れ圧込み | ○ |  | Click(L35) |
| 2 Page2 | B6 | CommandButton | 油寄り,レベラースジ | ○ |  | Click(L41) |
| 2 Page2 | B8 | CommandButton | ロールマーク | ○ |  | Click(L53) |
| 2 Page2 | B9 | CommandButton | 噛み込み跡 | ○ |  | Click(L59) |
| 2 Page2 | B10 | CommandButton | 切断面異常 | ○ |  | Click(L65) |
| 2 Page2 | B11 | CommandButton | エッジ変色,油残り跡,油焼け | ○ |  | Click(L71) |
| 2 Page2 | B3 | CommandButton | 共ズレ | ○ |  | Click(L23) |
| 2 Page2 | Label60 | Label | 大きさ | ○ |  |  |
| 2 Page2 | Label61 | Label | 触感 | ○ |  |  |
| 2 Page2 | Label62 | Label | 数 | ○ |  |  |
| 2 Page2 | Label63 | Label | 素洗い | ○ |  |  |
| 2 Page2 | SelectO | Label |   | ○ |  |  |
| 2 Page2 | SelectS | Label |   | ○ |  |  |
| 2 Page2 | SelectK | Label |   | ○ |  |  |
| 2 Page2 | UYoucode | TextBox |  | ○ |  | KeyPress(L267) |
| 2 Page2 | Uzai | TextBox |  | ○ |  | KeyPress(L278) |
| 2 Page2 | Utyou | TextBox |  | ○ |  | KeyPress(L256) |
| 2 Page2 | Uhaba | TextBox |  | ○ |  |  |
| 2 Page2 | Uatu | TextBox |  | ○ |  |  |
| 2 Page2 | Uhyou | TextBox |  | ○ |  | KeyPress(L289) |
| 2 Page2 | L2 | Label | 検索中..... | × |  |  |
| 2 Page2 | Frame5 | Frame | 平坦度規格 | ○ |  |  |
| 3 Frame5 | Label66 | Label | 単位mm | ○ |  |  |
| 3 Frame5 | Label67 | Label | 1M当りｹﾞｰｼﾞ | ○ |  |  |
| 3 Frame5 | Label68 | Label | 巾歪 | ○ |  |  |
| 3 Frame5 | TaHiz2 | TextBox |  | ○ |  |  |
| 3 Frame5 | HaHiz2 | TextBox |  | ○ |  |  |
| 3 Frame5 | Label69 | Label | 丈歪 | ○ |  |  |
| 3 Frame5 | MG2 | TextBox |  | ○ |  |  |
| 3 Frame5 | KousinL3 | Label |  | ○ |  |  |
| 3 Frame5 | Label71 | Label | 中歪 | ○ |  |  |
| 3 Frame5 | Naka2 | TextBox |  | ○ |  |  |
| 3 Frame5 | Label72 | Label | 耳歪 | ○ |  |  |
| 3 Frame5 | Mimi2 | TextBox |  | ○ |  |  |
| 3 Frame5 | Toukyu | Label |  | ○ |  |  |
| 3 Frame5 | Sansyou2 | TextBox |  | ○ |  |  |
| 1 MP | TestP | Page/Form | (タブ表示名)Test | × |  |  |
| 2 TestP | CommandButton1 | CommandButton | Test用 | ○ |  | Click(L79) |
| 2 TestP | T1 | TextBox |  | ○ |  |  |
| 2 TestP | PathText | TextBox | 値=\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\DB\品質\ | ○ |  |  |
| 2 TestP | G_Frame | Frame | 品質ｸﾞﾚｰﾄﾞ.基準 | × |  |  |
| 3 G_Frame | Tobi | TextBox | 値=あ⏎あ⏎あ⏎あ⏎あ | ○ |  |  |
| 3 G_Frame | Label1 | Label | 飛込み | ○ |  |  |
| 3 G_Frame | Label3 | Label | 鋳造ﾊｸﾘ.AL片圧込 | ○ |  |  |
| 3 G_Frame | Label4 | Label | AL付着跡 | ○ |  |  |
| 3 G_Frame | Label5 | Label | 汚れ(ﾎﾞﾀ落ち)黒・茶 | ○ |  |  |
| 3 G_Frame | Label6 | Label | ｶﾐｺﾐ跡.ｼｮｯｸ跡 | ○ |  |  |
| 3 G_Frame | Label2 | Label | ｷｽﾞ(ｽﾘｷｽﾞ) | ○ |  |  |
| 3 G_Frame | Kizu | TextBox |  | ○ |  |  |
| 3 G_Frame | Haku | TextBox |  | ○ |  |  |
| 3 G_Frame | ALhu | TextBox |  | ○ |  |  |
| 3 G_Frame | Bota | TextBox |  | ○ |  |  |
| 3 G_Frame | Kami | TextBox |  | ○ |  |  |
| 3 G_Frame | ToG | TextBox |  | ○ |  |  |
| 3 G_Frame | KiG | TextBox |  | ○ |  |  |
| 3 G_Frame | HaG | TextBox |  | ○ |  |  |
| 3 G_Frame | ALG | TextBox |  | ○ |  |  |
| 3 G_Frame | BoG | TextBox |  | ○ |  |  |
| 3 G_Frame | KaG | TextBox |  | ○ |  |  |
| 2 TestP | L3 | Label | 検索中..... | × |  |  |
| 2 TestP | LineName | Label |  | ○ |  |  |
| 2 TestP | FlagAtu | Label |  | ○ |  |  |
| 0 | VcFlag | Label |  | ○ |  |  |
| 0 | MPH | MultiPage |  | ○ |  | Change(L168) |
| 1 MPH | P1 | Page/Form | (タブ表示名)品質写真 | ○ |  |  |
| 2 P1 | TableTB | TextBox |  | ○ |  |  |
| 0 | CommandButton6 | CommandButton | Close | ○ |  | Click(L155) |
| 0 | CommandButton4 | CommandButton | 薄板 | ○ |  | Click(L136) |
| 0 | CommandButton3 | CommandButton | 厚板 | ○ |  | Click(L132) |

- **コードにイベントがあるがフォーム上に同名コントロールが無いもの（未使用コード候補）**: fm1_Click(L140), fm2_Click(L143)

## UF対角（コード: `UF対角.frm.bas`）

- フォームCaption(設計時): (なし) / ClientWidth×Height(twip): 8010×7110 / StartUpPosition: 1  'オーナー フォームの中央 / ShowModal: (既定=True) / コントロール数: 8
- フォームイベント: Activate(L13)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | DB1 | Image |  | × |  |  |
| 0 | TK | TextBox | 値=対角A測定 | ○ |  |  |
| 0 | DB2 〜 DB5（4個） | Image | (空) | × |  |  |
| 0 | TR | TextBox |  | ○ |  |  |
| 0 | TC | TextBox |  | ○ |  |  |

## UF校正（コード: `UF校正.frm.bas`）

- フォームCaption(設計時): "品証検証.ｺﾝﾍﾞｯｸｽｶﾒﾗﾁｪｯｸ" / ClientWidth×Height(twip): 9345.001×7350 / StartUpPosition: 2  '画面の中央 / ShowModal: (既定=True) / コントロール数: 30
- フォームイベント: Layout(L304), Activate(L314), Initialize(L342), QueryClose(L392)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | Gauge | Label | 1000mm直尺の全長を測定してください | ○ |  |  |
| 0 | SokuteiH | Label |  | ○ |  |  |
| 0 | CommandButton1 | CommandButton | 測定完了 | ○ |  | Click(L68) |
| 0 | CBcl | CommandButton | ｸﾘｱ＆ﾘｾｯﾄ | ○ |  | Click(L49) |
| 0 | Saiken | Label | ｸﾘｱを押してもう1度測定してください | × |  |  |
| 0 | NGvalue | Label |  | × |  |  |
| 0 | INP1000 | TextBox |  | ○ |  | AfterUpdate(L220), Enter(L241) |
| 0 | F1 | Label | Flag1 | ○ |  |  |
| 0 | Kname | Label |  | ○ |  |  |
| 0 | CommandButton3 | CommandButton | 閉じる | ○ |  | Click(L126) |
| 0 | Label3 | Label | 読み取りのﾊﾞﾗﾂｷ調査.ｶﾒﾗﾁｪｯｸ | ○ |  |  |
| 0 | INPG | TextBox |  | ○ |  | AfterUpdate(L250), Enter(L272), Exit(L281) |
| 0 | Rank1 | Label |  | × |  |  |
| 0 | Rank2 | Label |  | × |  |  |
| 0 | CommandButton4 | CommandButton | 品証順位 | ○ |  | Click(L129) |
| 0 | Rank3 〜 Rank6（4個） | Label | (空) | × |  |  |
| 0 | MaxLab | Label | ﾌﾟﾗｽ外れ | × |  |  |
| 0 | MinLab | Label | ﾏｲﾅｽ外れ | × |  |  |
| 0 | TukiLabe | Label |  | × |  |  |
| 0 | CommandButton5 | CommandButton | 送り | ○ |  | Click(L133) |
| 0 | CommandButton6 | CommandButton | 校正順位 | ○ |  | Click(L173) |
| 0 | Label5 | Label | ﾊﾟｹｯﾄ無しの場合””　判定もなし | ○ |  |  |
| 0 | Label6 | Label | 先 | ○ |  |  |
| 0 | Label7 | Label | 後 | ○ |  |  |
| 0 | TextBox1 | TextBox | 値=#L999.9 | ○ |  |  |
| 0 | LineName | Label |  | ○ |  |  |
| 0 | Ltrue | Label |  | ○ |  | DblClick(L298) |

## UF異常値発見（コード: `UF異常値発見.frm.bas`）

- フォームCaption(設計時): "異常" / ClientWidth×Height(twip): 13815×7230 / StartUpPosition: 2  '画面の中央 / ShowModal: (既定=True) / コントロール数: 11
- フォームイベント: Layout(L18), Activate(L28), Initialize(L40), QueryClose(L97)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | Label3 | Label | 異常値発見！ | ○ |  |  |
| 0 | IJO | Label |  | ○ |  |  |
| 0 | CommandButton1 | CommandButton | 閉じる | ○ |  | Click(L112) |
| 0 | TEIGI | Label | 異常値の定義　公差の倍とする | ○ |  |  |
| 0 | ATU | Label |  | ○ |  |  |
| 0 | HABA | Label |  | ○ |  |  |
| 0 | TAKE | Label |  | ○ |  |  |
| 0 | TAIK | Label |  | ○ |  |  |
| 0 | TEIGITAI | Label | 対角線　a^2＋b^2±6で比較 | ○ |  |  |
| 0 | Za | Label |  | ○ |  |  |
| 0 | TBname | Label |  | ○ |  |  |

## UF管理（コード: `UF管理.frm.bas`）

- フォームCaption(設計時): "2024.7.26_Ver" / ClientWidth×Height(twip): 21510×7170 / StartUpPosition: 2  '画面の中央 / ShowModal: (既定=True) / コントロール数: 200
- フォームイベント: Activate(L666), Initialize(L670)

| 階層 親 | 名前 | 種類 | Caption / 初期値 | 表示 | Tag | コード内イベント(行) |
|---|---|---|---|---|---|---|
| 0 | Frame1 | Frame | 保存データ | ○ |  |  |
| 1 Frame1 | HM | Label |  | ○ |  |  |
| 1 Frame1 | SE | Label |  | ○ |  |  |
| 1 Frame1 | CB1 | ComboBox |  | ○ |  | Change(L37) |
| 1 Frame1 | ATU | Label |  | ○ |  |  |
| 1 Frame1 | HABA | Label |  | ○ |  |  |
| 1 Frame1 | TAKE | Label |  | ○ |  |  |
| 1 Frame1 | AM | Label |  | ○ |  |  |
| 1 Frame1 | AP | Label |  | ○ |  |  |
| 1 Frame1 | HP | Label |  | ○ |  |  |
| 1 Frame1 | TM | Label |  | ○ |  |  |
| 1 Frame1 | TP | Label |  | ○ |  |  |
| 1 Frame1 | S1 〜 S12（12個） | Label | (空) | ○ |  |  |
| 1 Frame1 | U1 〜 U12（12個） | Label | (空) | ○ |  |  |
| 1 Frame1 | C1 〜 C12（12個） | Label | (空) | ○ |  |  |
| 1 Frame1 | D1 〜 D12（12個） | Label | (空) | ○ |  |  |
| 1 Frame1 | E1 〜 E12（12個） | Label | (空) | ○ |  |  |
| 1 Frame1 | A1 〜 A4（4個） | Label | (空) | ○ |  |  |
| 1 Frame1 | L1 〜 L8（8個） | Label | (空) | ○ |  |  |
| 1 Frame1 | S | Label | S | ○ |  |  |
| 1 Frame1 | U | Label | S1 | ○ |  |  |
| 1 Frame1 | C | Label | 2 | ○ |  |  |
| 1 Frame1 | D | Label | E1 | ○ |  |  |
| 1 Frame1 | E | Label | E | ○ |  |  |
| 1 Frame1 | Label88 | Label | ﾗ | ○ |  |  |
| 1 Frame1 | Label89 | Label | ﾗ | ○ |  |  |
| 1 Frame1 | Label91 | Label | 厚- | ○ |  |  |
| 1 Frame1 | Label92 | Label | 厚+ | ○ |  |  |
| 1 Frame1 | Label93 | Label | 巾- | ○ |  |  |
| 1 Frame1 | Label94 | Label | 巾+ | ○ |  |  |
| 1 Frame1 | Label95 | Label | 丈- | ○ |  |  |
| 1 Frame1 | Label96 | Label | 丈+ | ○ |  |  |
| 1 Frame1 | Count | Label |  | ○ |  |  |
| 1 Frame1 | Worker | Label |  | ○ |  |  |
| 1 Frame1 | SZH | Label |  | ○ |  |  |
| 1 Frame1 | SGH | Label |  | ○ |  |  |
| 1 Frame1 | UZH | Label |  | ○ |  |  |
| 1 Frame1 | UGH | Label |  | ○ |  |  |
| 1 Frame1 | CZH | Label |  | ○ |  |  |
| 1 Frame1 | CGH | Label |  | ○ |  |  |
| 1 Frame1 | DZH | Label |  | ○ |  |  |
| 1 Frame1 | DGH | Label |  | ○ |  |  |
| 1 Frame1 | EZH | Label |  | ○ |  |  |
| 1 Frame1 | EGH | Label |  | ○ |  |  |
| 1 Frame1 | H1 | Label |  | ○ |  |  |
| 1 Frame1 | H2 | Label |  | ○ |  |  |
| 1 Frame1 | CommandButton1 | CommandButton | ﾃﾞｰﾀ呼び出し | ○ |  | Click(L620) |
| 1 Frame1 | Elate | Label |  | ○ |  |  |
| 1 Frame1 | Slate | Label |  | ○ |  |  |
| 1 Frame1 | Ulate | Label |  | ○ |  |  |
| 1 Frame1 | Clate | Label |  | ○ |  |  |
| 1 Frame1 | Dlate | Label |  | ○ |  |  |
| 1 Frame1 | Llate | Label |  | ○ |  |  |
| 1 Frame1 | L9 | Label |  | ○ |  |  |
| 1 Frame1 | S13 | Label |  | ○ |  |  |
| 1 Frame1 | S14 | Label |  | ○ |  |  |
| 1 Frame1 | U13 | Label |  | ○ |  |  |
| 1 Frame1 | U14 | Label |  | ○ |  |  |
| 1 Frame1 | C13 | Label |  | ○ |  |  |
| 1 Frame1 | C14 | Label |  | ○ |  |  |
| 1 Frame1 | D13 | Label |  | ○ |  |  |
| 1 Frame1 | D14 | Label |  | ○ |  |  |
| 1 Frame1 | E13 | Label |  | ○ |  |  |
| 1 Frame1 | E14 | Label |  | ○ |  |  |
| 1 Frame1 | L10 | Label |  | ○ |  |  |
| 1 Frame1 | Mimi | CheckBox | 耳付き (初期値=0) | ○ |  | AfterUpdate(L641) |
| 1 Frame1 | WorkJudgment | Label |  | ○ |  |  |
| 1 Frame1 | LhizB | Label |  | ○ |  |  |
| 1 Frame1 | EGH_B | Label |  | ○ |  |  |
| 1 Frame1 | SZH_B | Label |  | ○ |  |  |
| 1 Frame1 | SGH_B | Label |  | ○ |  |  |
| 1 Frame1 | UZH_B | Label |  | ○ |  |  |
| 1 Frame1 | UGH_B | Label |  | ○ |  |  |
| 1 Frame1 | CZH_B | Label |  | ○ |  |  |
| 1 Frame1 | CGH_B | Label |  | ○ |  |  |
| 1 Frame1 | DZH_B | Label |  | ○ |  |  |
| 1 Frame1 | DGH_B | Label |  | ○ |  |  |
| 1 Frame1 | EZH_B | Label |  | ○ |  |  |
| 1 Frame1 | LhigB | Label |  | ○ |  |  |
| 1 Frame1 | Flate | Label |  | ○ |  |  |
| 1 Frame1 | F1 〜 F10（10個） | Label | (空) | ○ |  |  |
| 1 Frame1 | F | Label | F | ○ |  |  |
| 1 Frame1 | FZH | Label |  | ○ |  |  |
| 1 Frame1 | FGH | Label |  | ○ |  |  |
| 1 Frame1 | F13 | Label |  | ○ |  |  |
| 1 Frame1 | F14 | Label |  | ○ |  |  |
| 1 Frame1 | FGH_B | Label |  | ○ |  |  |
| 1 Frame1 | FZH_B | Label |  | ○ |  |  |
| 1 Frame1 | F11 | Label |  | ○ |  |  |
| 1 Frame1 | F12 | Label |  | ○ |  |  |
| 1 Frame1 | CommandButton6 | CommandButton | Close | ○ |  | Click(L637) |
| 1 Frame1 | NewCB | ComboBox |  | ○ |  | Change(L11) |
| 0 | SL | Label | S | ○ |  |  |
| 0 | UL | Label | S1 | ○ |  |  |
| 0 | CL | Label | 2 | ○ |  |  |
| 0 | DL | Label | E1 | ○ |  |  |
| 0 | EL | Label | E | ○ |  |  |
| 0 | Ae_Vis | Label |  | ○ |  |  |
| 0 | As_Vis | Label |  | ○ |  |  |
| 0 | Au_Vis | Label |  | ○ |  |  |
| 0 | Ac_Vis | Label |  | ○ |  |  |
| 0 | Ad_Vis | Label |  | ○ |  |  |
| 0 | Bs_Vis | Label |  | ○ |  |  |
| 0 | Be_Vis | Label |  | ○ |  |  |
| 0 | Bu_Vis | Label |  | ○ |  |  |
| 0 | Bc_Vis | Label |  | ○ |  |  |
| 0 | Bd_Vis | Label |  | ○ |  |  |
| 0 | Label107 | Label | A | ○ |  |  |
| 0 | Label108 | Label | B | ○ |  |  |
| 0 | E_BH | Label |  | ○ |  |  |
| 0 | S_BH | Label |  | ○ |  |  |
| 0 | U_BH | Label |  | ○ |  |  |
| 0 | C_BH | Label |  | ○ |  |  |
| 0 | D_BH | Label |  | ○ |  |  |
| 0 | Label114 | Label | BH | ○ |  |  |
| 0 | E_Angle | Label |  | ○ |  |  |
| 0 | S_Angle | Label |  | ○ |  |  |
| 0 | U_Angle | Label |  | ○ |  |  |
| 0 | C_Angle | Label |  | ○ |  |  |
| 0 | D_Angle | Label |  | ○ |  |  |
| 0 | Label120 | Label | 直角度 | ○ |  |  |
| 0 | F_Angle | Label |  | ○ |  |  |
| 0 | FL | Label | F | ○ |  |  |
| 0 | Af_Vis | Label |  | ○ |  |  |
| 0 | Bf_Vis | Label |  | ○ |  |  |
| 0 | F_BH | Label |  | ○ |  |  |
