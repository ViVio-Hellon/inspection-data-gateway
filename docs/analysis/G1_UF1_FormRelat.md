# G1 構造解析レポート: `UF1.frm.bas` / `【UFKen_02】FormRelat.bas.bas`

- 対象1: `vba/UF1.frm.bas`（5336行、メインUserForm UF1 のコードビハインド）
- 対象2: `vba/【UFKen_02】FormRelat.bas.bas`（7218行、UF1の画面制御・入力先決定・Min/Max・作業選択レイアウト等）
- 行番号は各ファイル内の物理行。「推測」と書いたものはコードから断定できない内容。
- コントロールのキャプション／配置（.frx / デザイナ情報）は入力に含まれていないため、**コード中に現れる名前・代入値からのみ**画面構成を復元している。

---

## 0. 移植で特に注意すべき点（先に要約）

1. **判定結果がコントロールの BackColor に保持されている**。OK=`13434828`(RGB 204,255,204)、NG=`8421631`(RGB 255,128,128)、フォーカス中=黄 RGB(255,255,0)。`NGrange_Check`/`BackColor_Check`/`NGErase`/`ResetBackColorIfNotRedOrGreen` などは色で OK/NG を判断している（FormRelat 1548-1600, 1844-1860, 5495-5600）。Python側では「判定状態」を明示的なデータとして持ち、色は表示に派生させる必要がある。ただし挙動は色ベースの判定と完全一致させること。
2. **イベント抑止・状態フラグがコントロールの Caption/Tag に分散**（`UF1.Load.Caption`="on"/"off"、`UF1.Tag`="STOP"、`FlagLabel`、`Auto`、`INP.Tag`=現在の入力先セル名 等）。詳細は §9。
3. **既存のバグ/癖（同等性維持のため要判断）**
   - `MinMax` の `If 0 < val(UF1.ATU) <= 74# Then`（FormRelat 5802）は VBA では `(0<x)` が True(-1)/False(0) になり、それと 74 を比較するため**常に真**。結果、厚 Min/Max は ATU>74 でも常に `"0.000"` 書式。ElseIf は死にコード。
   - `PrintB_Click` のラテラルボー必須判定 `UF1.Sumisei.Caption <> """"`（UF1 3687）は「ダブルクォート1文字と異なる」の意味で**ほぼ常に真**（旧版は `<> ""`、UF1 3909 のコメントアウト版参照）。
   - `INP_Change` では入力先決定に使う検番を `Left(UF1.Reien.Caption, 1)`（UF1 4585、冷延表示欄）から取っているが、`UserForm_Click` では `Left(UF1.KenNo.Caption, 1)`（UF1 4362）。Reien は 2HI/L-2/L-3/LL7 のときしか設定されない（FormRelat 2729-2768）ため、通常は空文字→`AllCheck`では flag(5)(=5回扱い)、`Look_S_Blank`では flag(1)(=3回扱い) になる。
   - `FBad_MouseDown` が `UF1.EBad = Format(UF1.FBad,"0")` と **E欄に書いている**（UF1 1831）。
   - `InspectionSelect` の2山時 `OPS5.Caption = "2-2-E"`（FormRelat 2273, 2457）。他箇所（LSelect / Yama_Change）は `"2-1-E"`。
   - `AtCountSelect` 検番T・6回時、F列は `.Visible=True` を2回設定し `.Enabled` を設定していない（FormRelat 2144, 2150）。
   - `Line_AutoSelection` で Lot情報履歴にロットが無いと `WorkerLoad` 後に `Exit Sub`（FormRelat 6531-6534）し、`UF1.Load="on"`/`UF1.Tag=""` に戻らない（イベント停止状態・音声停止状態が残る）。
4. **UF1 内の `Private SikaAtuAll / SikaSetAtu`（UF1 107-108）が `【UFKen_03】SaveLoad` の同名 Public 変数を隠蔽**している。UF1 内コードは UF1 私有変数、FormRelat の `LSelect` は SaveLoad の Public 変数を参照しており、別物。
5. **LSelect のキャッシュ使用時もオーダー厚はキャッシュ値にならない可能性**: `NewAtu = IIf(bUseCache, g_CacheOrdAtu, LotAcc(1, OrdAtu))`（FormRelat 2773）。VBA の IIf は両辺を評価するため、キャッシュ使用時（LotAcc=Empty）は `LotAcc(1, OrdAtu)` でエラー→`On Error Resume Next` により代入ごとスキップ→NewAtu=0→`val(UF1.ATU)` を使用（=6mm 判定は画面の厚で行われる）。移植時に「キャッシュ値を使う」と素直に書くと結果が変わりうる。
6. VBA の暗黙変換・丸めに依存: `th = Format(Round(CalculateSheets(...),1),"0")` を Long に代入（Round は銀行丸め、Format は四捨五入系）、`Format$("", "0.000")` は `""` を返す、Label.Caption(String) と Double の `=` 比較（UnderBar）、`Select Case` の `Case 1 To 6#` 範囲（境界の隙間あり）など。

---

## 1. モジュール概要

**UF1.frm.bas** は UF検査判定システムのメイン画面 UF1（モードレス）のイベントハンドラ群。起動時のライン設定読込（レジストリ）、検査員登録（社員カード13桁）、ロット登録（7桁英数字）、作業選択（厚板/薄板/少枚数/極厚/スラブ）、測定値入力（マイクロメータ/コンベックス等の伝送データ `DT…`/`#L…`/手入力 `#〆…` をバーコードリーダ的に StartTB→INP 経由で受信）、歪・ラテラルボー・直角度・BH・検入高さ・外観不良の入力、Min/Max計算、アイカスインプット起動、印刷・保存（シート→CSV→共有フォルダ）、新規検査（作業シート初期化）、各サブフォーム（全数/三菱/5N、H494、キャスク、品質基準、管理、校正）への遷移を担う。
**【UFKen_02】FormRelat** は UF1 から呼ばれる画面制御の共通モジュールで、フォームサイズ調整、測定入力のメイン処理（LoopText）、「次に入力すべきセル」の自動決定（PucketMove / Look_Blank 系）、未測定チェック（AllCheck）、作業選択別の表示レイアウト・検査回数（LSelect/AtCountSelect/InspectionSelect/Six_Display/SlabSet）、引当情報からのラテラルボー/直角度/フリー反り要否判定（JudgeLotAllocations）、Min/Max算出と下線表示（MinMax/UnderBar）、オーダー厚による作業選択ボタンの表示制御（Atu_OPselect）、包装仕様書No照合（PackingNoCheck）等を持つ。

---

## 2. Sub/Function 一覧

### 2.1 UF1.frm.bas（すべて Private）

#### (a) フォームライフサイクル
| 名前 | 行 | 役割 | 主な呼び出し先 |
|---|---|---|---|
| UserForm_Initialize | 173 | 起動処理（§4-1） | InitializePrintFlags, CleanOldHistory (Print), ProgressBar_count, GetSetting, Set_Shift, Visual_Add, GetZoom, ライン Option=True→Line_AutoSelection, CheckAndUpdateButtonColor, SikaAtuCreate, CBAdd, LabelCls登録, InitialSetting |
| UserForm_Activate | 4218 | アクティブ化時チェック（検査員登録有無、WorkerTime "OUT" 隠蔽対策、シート表示整理、校正済判定） | Register_Check, WriteLog, NameList_Open, VisibleSheets/HiddenSheets, FormSetting, 校正 |
| UserForm_Click | 4300 | フォームクリック＝再判定と次入力位置の再計算（先行データ票警告、LoopJudg等） | AdvanceCheck, MsgError, LoopJudg, HizLoopJudg, CheckLabelsByCategory×6, DefaultFormPosition, AllClose, LookingYellow, Look_*_Blank, WorkerName_Open, SmallNumberCount |
| UserForm_DblClick | 4417 | `Load.Caption="on"`（イベント再開の手動手段） | — |
| UserForm_Layout | 4421 | `Application.Visible=False` | — |
| UserForm_QueryClose | 4432 | 終了処理：音停止、Posting、UF管理/UFWorkerName閉、シートVeryHidden、Excel最大化表示、ThisWorkbook.Save | Posting, HiddenSheets("xlVeryHidden") |
| UserForm_Terminate | 4456 | BlinkLabel1 解放 | — |
| UserForm_Resize | 4202 | 中身全コメント（空） | — |
| UserForm_MouseMove | 4719 | 山数選択パネル(PicFM)を畳む、MPEX/MP1表示 | — |

#### (b) ライン選択・モード
| コントロール/名前 | 行 | 業務上の意味 |
|---|---|---|
| L1Work_Click / LVC_Click / HVC_Click / Tott_Click / Bala_Click / Test_Click / Maru_Click | 3267-3349 | ライン（設備）選択 OptionButton。`Load="off"`中は n="" それ以外 "表示" で Line_AutoSelection |
| A1_Click … A11_Click | 367-431 | 丸徳ラインの号機(LineNo)選択。丸徳時のみ Line_AutoSelection("表示") |
| 厚板_Click | 5204 | 作業選択=厚板。LSelect("厚板",True)（描画リセットのみ）→ Sumisei<>"" なら LSelect("住精") それ以外 LSelect("厚板") → OpForcedSelection |
| 薄板_Click | 5278 | 作業選択=薄板。LSelect("薄板",True)→ Sumisei<>"" なら LSelect("住精")。※Loading表示を消す処理が無い（Sumisei=""時 Loading.Visible=True のまま。LSelect 内でも消していない→推測: 表示残り） |
| 少枚数_Click | 5244 | 作業選択=少枚数（1〜5枚）。検番H なら SmallNumberCount(th) して終了 |
| 極厚_Click | 5302 | 作業選択=極厚（耳付き） |
| スラブ_Click | 5312 | 作業選択=スラブ |
| Byousya_AfterUpdate | 554 | 「強制描写」(Byousya)：作業選択を全解除し 6回目表示、ラテラル/直角度を強制表示 |
| Yama_Change | 2246 | 山数(1〜7)。1〜4 は UF1 で検査、5以上は「別フォームで検査してください」(Yama.Tag="別フォーム")。表示パターン切替、AtCountSelect/InspectionSelect、ログ「山数ﾁｪﾝｼﾞ」 |
| Yama_KeyPress | 2235 | 英数字のみ・小文字→大文字 |
| Yama_MouseDown / SlectFM_MouseMove / ErFM_MouseMove | 2462-2506 | 山数選択パネル(PicFM)の表示/非表示 |
| AtButt_Click | 4469 | 厚測定箇所 2ヵ所/4ヵ所 切替（Caption は「次に切り替わる先」を表示） |
| InspeButt_Click | 4481 | 検入枚数「検入≧26」/「検入≦25」切替（同じく次状態を表示） |
| OPS1..6_Change / OPE1..6_Change | 3535-3571 | Min/Max対象範囲（開始列/終了列）のOptionButton → TB_BackColor |
| AorB_Click | 477 | 歪入力の A面/B面切替（Now_A/Now_B）、ReflectStrain、HizLoopJudg |
| Abno_Click | 434 | 「異常値on/off」キャプショントグル（異常値除去 で参照と推測） |
| ReSize_Click | 4057 | サイズ・公差の手動編集を許可（ATU/Hab/Tak/Ap/Am/Hp/Hm/Tp/Tm/TAIK/Lateral ロック解除）→Posting |

#### (c) 検査員・ロット
| 名前 | 行 | 意味 |
|---|---|---|
| WorkButton_Click | 4743 | 検査員登録開始（バランサーは確認なし）。未登録なら Daily_Create（日報作成）、`C:\Windows\System32\ctfmon.exe` 起動、Worker にフォーカス |
| Worker_Change | 4788 | カード13桁で NameConv→直(TimeCheck)により Day1/T1(日勤・1直), Day2/T2(2直), Day3/T3(3直) へ "_" 連結で追記、ログ、PicFM.Tag=0、先行データ票警告、Daily_Report(1) |
| Worker_Enter/Exit | 4870/4867 | JudgBox1 緑/黒 |
| T1_Change / T2_Change / T3_Change | 4099/4121/4141 | WorkerSave、WorkerName_Open、当日登録かつ未校正なら 校正_Open（Excel最小化）+音声 |
| Day1/2/3_Change | 3120-3131 | 実質何もしない（Exit Sub のみ） |
| CommandButton10_Click | 2937 | **Lot登録**ボタン。確認→ロット分割確認2回（Yes で SplitFlag="Lot分割"）→Lotクリア＆入力可能化、関連表示クリア |
| Lot_KeyPress | 4708 | 英数字のみ・大文字化 |
| Lot_Change | 4660 | 7桁で `CB = Lot`（→CB_Change）、Lot登録ログ（ローカル、HistoryPath、Measured.log）、導電率チェック（Conductiv 赤） |
| Lot_Enter/Exit | 4657/4652 | JudgBox2 緑/黒 |
| CB_Change | 4494 | ロット確定処理（§4-3） |
| NoL_Click | 1717 | OrdNoOpen（オーダーNoをコンボへ） |
| OrdNoCB_Change | 1721 | オーダー選択→Lot情報履歴から包装仕様NO→伝送用ファイル.accdb PkgSpecData から梱包形態/タイトル表示、VC_Picture、MaterialsMsgBox |
| LoadLot_KeyPress / LoadWorker_Change | 1607/1618 | Accesslog読込用のロット/作業者入力（13桁カード→NameConv） |
| CommandButton15_Click | 579 | AccesslogLoad(LoadLot, LoadWorker, LoadTextName, Reason)＋ReDiag（Accesslogからの復元） |

#### (d) 測定入力
| 名前 | 行 | 意味 |
|---|---|---|
| StartTB_KeyDown | 3295 | Tab/Enterで `INP = StartTB`、StartTB クリア（スキャナ/伝送受信口） |
| StartTB_AfterUpdate | 3307 | 全コメント（空） |
| INP_Change | 4565 | **測定値入力のメイン**（§4-5） |
| CommandButton3_Click | 4552 | 手入力：`INP = "#〆" & Manual`（Len(INP)>=4 時）／`"#〆" & Format(Manual,"0.000")`（Len(INP)<=3 時） |
| Manual_Change | 3351 | 数値以外はクリア |
| CommandButton2_Click | 4537 | 再判定（LoopJudg/HizLoopJudg/Look_Blank） |
| CommandButton4_Click | 4459 | **MinMax** 実行（LoopJudg/HizLoopJudg→MinMax） |
| AMode_Click | 447 | **アイカスインプット起動**（Inspection1..4 で検査法選択必須。Test ライン or DBflag<>"" は無条件、それ以外は AMin/AMax/HMin/HMax/TMin/TMax がすべて非空のとき Ainput_Startup） |
| LateralS..F_Change / _BeforeUpdate / _AfterUpdate / _Enter / _Exit / _MouseDown | 771-1456 | ラテラルボー入力（6位置）。数値 or "-"。maxLength 3（"-"始まり4）。BeforeUpdate で UnitRecord(0.5ピッチ)→"0.0"→LateralbowJudgColor→Posting→Out_Tolerance |
| DiagSKou..DiagFKou_*（同上） | 887-1266 | 直角度入力。同様に UnitRecord→"0.0"→RightAngleJudgColor。※DiagUKou_BeforeUpdate のみ先頭で TagSet を呼ぶ（1129） |
| ZS..ZF_Change/BeforeUpdate/AfterUpdate/Enter/Exit/MouseDown | 2513-2567, 4901-5113 | 歪（定盤）入力。CB?1 が "耳"/"中" なら maxLength 5・非数値許容、他は数値のみ maxLength 4。BeforeUpdate で "0.00"、ReflectStrain(A/B,"保存")。AfterUpdate で HizumiJudg, LogOut(…,"AorB"), Out_Tolerance, CheckLabelsByCategory |
| GS..GF_*（同上） | 2569-2623, 5016-5201 | 歪（ゲージ）入力。数値のみ maxLength 4。処理は Z と同様 |
| CBS1..CBF4_Enter/Exit/Change | 2625-2935 | 歪コンボ（位置×4段）。Exit で Load="on"、値変化時 LogOutCombo。CB?1_Change→CB_UpDownAdd、CB?3_Change→CBG_UpDownAdd |
| FMS..FMF_Click, LZS..LZE_Click | 3168-3265 | 歪フレーム選択→FMCheck |
| SBH..EBH_Change/AfterUpdate | 1835-1875 | BH（板の何らかの硬さ/高さ測定と推測）入力。"0"書式→LogOut→simplyjudge(BH, BhJudg(BH, Zai, Tyu)) ※FBH は存在しない |
| KenHBox1..6_AfterUpdate | 696-726 | 検入高さ入力 "0.0"→LogOut→CalculateSheets(UF1, val(ATU)) |
| KenHBox7_AfterUpdate / _Change | 727/742 | "0"書式、th 算出のみ（SmallNumberCount 呼出はコメントアウト） |
| KenCount_AfterUpdate / _Change | 745/766 | 少枚数の検入枚数 "0"。厚板.Caption="H"かつ少枚数なら SmallNumberCount(th) |
| Atu/Hab/Tak/Ap/Am/Hp/Hm/Tp/Tm_Change | 442-553, 3199-3212, 4161-4172 | 数値以外クリア、Ap/Am/Hp/Hm/Tp/Tm は Form_Formula |
| Taik_Change | 4173 | 対角等級 "4"→Diagonal(Hab,Tak)、"3"→ReDiagonal、DebugTaikaku に "0.0" |
| Lateral_Change | 4186 | ラテラル等級 "4"→Lateralbow_J(Hab,Tak,ATU) "0"、"3"→ReLateralbow "0.0"、0なら空 |
| MPEX_Change | 1642 | タブ切替。0ページ=DrawListBoxContent、2ページ(歪)=重量により案内MsgBox・音声・HizLoopJudg |
| MPEX_Click / MPEX_MouseMove / MP1_Click / MP1_MouseMove / MP1_Change | 1634-1716, 3354 | MPEX.Tag="Click"（判定表示抑止用）、外観ボタン白化、少枚数 SmallNumberCount、MP1 ページ5で IM_Select、ページ1で PicFM.Tag=0 |
| YamaLB_Click | 2508 | DrawListBoxContent |

#### (e) 外観・不良・保留・削除
| 名前 | 行 | 意味 |
|---|---|---|
| LC1..LC12_Click | 1458-1505 | 外観ボタン LcColor(LCn,1)＋CheckLabelsByCategory（LC1=S, LC2=U … 対応は LC1〜LC6 / LC7〜LC12 の2段と推測） |
| LC1..LC10_DblClick | 1507-1546 | 色白化＋btn_allClea(LBVisual)（LC3〜LC10 は白化なし） |
| SBad..FBad_Change / _MouseDown | 1791-1834 | 外観不良枚数入力 "0"→TeColor |
| ListReflect_Click | 1548 | 不良要因リスト(DefectFactorLB)選択＋枚数→SelectionInput(UF1(TecReturn))。50枚超は分割要求（"ｽﾎﾟｯﾄ処理"は除外） |
| Cle_Click | 567 | 不良要因印刷_<Lname> の C3:X8, C10:X10, C11 をクリア |
| ViewCB_Click | 1931 | 不良要因印刷シート B1:X11 を画像化して Image1 に表示（クリップボード API・リトライ） |
| Image1_Click | 630 | 画像表示解除・コントロール再表示、JudgeLotAllocations→LSelectLateral/RightAngleDisplay、Additional により LSelectLateral |
| NGS..NGF_DblClick | 330-365 | 「まとめて削除」：ConfiReason_Open(位置,0,Lname,"<位置>板1枚",DeleteC) |
| CommandButton18_Click | 120 | **保留**処理：INP.Tag から位置(S/U/C/D/E/F)取得→確認→NG_Pass(位置,"保留") |
| IzyouFl_Click | 3213 | 異常内容印刷フラグ(異常ロット用_<Lname> A1=Lot)取消（2回確認でシート削除） |
| Conductiv_Click | 591 | 導電率測定用紙印刷（Conductiv 赤時のみ）→ConductivLoad→シート「導電率測定」PrintOut |

#### (f) 印刷・保存・新規・遷移
| 名前 | 行 | 意味 |
|---|---|---|
| PrintB_Click | 3619 | **印刷・保存**（§4-7） |
| PrintOnly | 3574 | 印刷済ロットの再印刷（印刷用_<Lname> D2="再印刷"） |
| PrintBreset_Click | 3599 | 印刷ボタンがオレンジ RGB(255,165,0) のとき印刷履歴削除(ClearPrinted) |
| NewKen_Click | 3359 | **新規検査**（§4-8） |
| H494B_Click | 604 | H494 検査へ（OpenMitubisi("H494")、UF1.Hide） |
| ZEB_Click | 4874 | 全数/三菱/5N 検査へ（OpenMitubisi("全数")、UF1.Hide）。キャプションは丸徳時「検入4山以上」、他「全数.三菱.5N」 |
| KYB_Click | 3230 | キャスク検査へ（UFCask.Show、UF1.Hide） |
| CommandButton8_Click | 3053 | 品質基準画面 UF品質基準（ATU>=4 なら Page0、未満 Page1、MPH へ写真ページ追加、DistArray→FormReflectDist、ATU>=8 なら DBdisplay外観_H） |
| CommandButton7_Click | 3048 | 管理画面 UF管理（UFsizeSave 後） |
| CommandButton6_Click | 3029 | 校正画面 UF校正（最新登録日の検査員名） |
| CommandButton13_Click | 3004 | パスワード：`weekend99`→隠しオブジェクト表示＋SF3、`jump`→校正スキップ(WorkButton.Tag="jump") |
| CommandButton14_Click | 3024 | Excel表示＋KenSheets |
| CommandButton9_Click | 3113 | Unload Me（閉じる） |
| 表示_DblClick | 5329 | Excel・全シート表示（保守用） |
| TestButton_Click | 1876 | テスト用：印刷転送→印刷用_Test 表示 |
| blindfm_DblClick | 115 | 目隠し解除 frameblindfold("通常") |
| Frame1_Click / Frame7_Click | 4559/4562 | StartTB へフォーカス |

### 2.2 【UFKen_02】FormRelat.bas.bas

| 名前 | 行 | Pub/Priv | 引数 | 役割 | 主な呼び出し元 → 呼び出し先 |
|---|---|---|---|---|---|
| GetScreenWidth / GetScreenHeight | 49/53 | Public | — | GetSystemMetrics(0/1) | 呼出元なし |
| GetZoom | 57 | Public | myForm | UfZoom/UfHeight/UfWidth を保存 | UF1 Init 他フォーム |
| InitialSetting | 64 | Public Func | myForm, n("サイズ保存"/"リセット"/"リスタート"/"初期値からスタート"/"") | Excelウィンドウサイズ比率でフォームの Zoom/幅/高さ/位置を決定。戻り値 Array(旧Zoom,旧H,旧W,新Zoom,新H,新W) | UF1 他多数 → DebugLog |
| FormSetting | 255 | Public | — | アクティブウィンドウに WS_THICKFRAME 付与（リサイズ可） | UF1 Activate 他 |
| auto_open | 269 | Public | — | 空（コメントのみ） | — |
| UF1_Close | 273 | Public | — | UF1.Hide（コメントに「Application.Run用」） | ブック内呼出なし（他ブックから Run の可能性：要確認） |
| LoopText | 278 | Public | st(セル名) | 測定入力のメイン：異常値除去→Machin_Text→Micro_Juge→TextJudg→TimeRecord→NumDig→CheckLabelsByCategory×6 | INP_Change |
| TimeRecord | 303 | Public | st | 測定時刻記録・2度押し検出（4秒以内同値でクリア）、Daily_Report(2)、Time_Posting | LoopText, Worker |
| NumDig | 345 | Public | — | 表示桁数整形（§5-2） | LoopText |
| AllCheck | 429 | Public Func | a,n,KenNo,Temp | 指定行範囲に未入力があれば True（作業/検番/山数で対象列が変わる） | PucketMove |
| PucketMove | 716 | Public | Pucket("DT"/"#L"/"#〆"), KenNo, Youto | 伝送データ種別から入力対象行範囲を決め、空セルを探して INP.Tag にセット | INP_Change |
| Look_S_Blank | 1018 | Public | a,n,KenNo | 薄板用「次の空セル」探索＋黄色表示 | UF1 |
| LookBlank_Color | 1272 | Public | a(セル名) | 次入力セルを赤枠黄背景（MPEX.Tag<>"" なら黒枠白） | Look_S_Blank, Smallnumber_Blank |
| Smallnumber_Blank | 1283 | Public | a,n,xa(枚数) | 少枚数用 空セル探索 | UF1 |
| Look_Blank | 1341 | Public | a,n | 厚板/スラブ/その他 空セル探索 | UF1, OutTolerance |
| ObjMoveText | 1502 | Public Func | Label | 名前を返す | LabelCls, Test |
| TagSet | 1505 | Public Func | TextBox | INP.Tag/Tagtext に名前セット | UF1 Enter系 |
| TextTag | 1510 | Public Func | Label | ラベルクリックで手動入力先指定（Auto="off"）、LoopJudg/HizLoopJudg | LabelCls |
| DiagonalErase | 1530 | Public | Label | 行9/10 消去時に Diag? クリア（極厚以外） | LabelCls |
| NGErase | 1548 | Public | Label | 位置列にNG色が無ければ NG? ラベルを空に | LabelCls, Judg |
| FMCheck | 1602 | Public | Frame | 歪フレーム強調（幅390・色39423） | UF1 |
| Fm_Clear | 1620 | Public | n | 歪フレーム幅/色リセット＋Z?を"0.00"書式 | UF1, Judg |
| FindYellowLabel | 1831 | **Private** Func | — | UF1上の黄色ラベル検索 | ResetBackColor |
| ChangeBackColor | 1844 | Public | ctrl | OK/NG色以外なら黄に | UF1 Enter系 |
| ResetBackColorIfNotRedOrGreen | 1852 | Public | ctrl | OK/NG色以外なら既定色に | ResetBackColor |
| ResetBackColor | 1862 | Public | n | 黄色ラベル解除＋歪/直角/ラテラル入力欄の色戻し | UF1 Exit系 |
| CBdateCheck | 1935 | Public Func | — | 歪入力が1か所以上あるか | 呼出なし（旧PrintBのみ） |
| AtCountSelect | 1950 | Public | n(2/4) | 薄板時の厚測定箇所(2/4)切替。DBから検査番号取得 | UF1, SaveLoad, Six_Display |
| InspectionSelect | 2183 | Public | n("表示"/"非表示") | 薄板時の検入枚数(≦25/≧26)による列表示切替 | UF1, SaveLoad |
| RightAngleDisplay | 2495 | Public | n(検番/"強制"/"新規"), Count | 直角度欄表示 | UF1, LSelect, SaveLoad |
| ClearLSelectCache | 2568 | Public | — | g_Cache* クリア、g_SkipLotNotFoundMsg=True | NewKen |
| LSelect | 2579 | Public | n(作業種別/"住精"), bSkipDBCheck | 作業選択別レイアウト、2HI等コース判定、ゲージ要否、引当判定（§5-5） | UF1, SaveLoad |
| SlabSet | 5104 | Public | — | スラブ用レイアウト | LSelect, SaveLoad |
| JudgeLotAllocations | 5259 | Public | TempArry, ByRef 検番, ByRef NeedLateral/NeedRight/NeedFlat | 引当（Lot情報履歴の複数行）ごとに DbReturnItem で要否判定し「1件でも要なら要」 | LSelect, Line_AutoSelection, UF1 Image1_Click, SaveLoad |
| LSelectLateral | 5335 | Public | n, Count | ラテラルボー欄表示（UF1.Width=1000） | 多数 |
| nCount | 5423 | Public Func | — | 測定値が1つでもあるか | Judg, Quality_DB, ListView |
| Load_OP | 5455 | Public | — | 作業保存シート T1 から作業選択復元 | **呼出なし** |
| TB_BackColor_Clear | 5476 | Public | OptionButton | Min/Max範囲ラベル色リセット | TB_BackColor |
| BackColor_Check | 5495 | Public Func | Label | NG色かつ NG?≠"許容" で True | NGrange_Check |
| BackColor_CheckTB | 5504 | Public Func | TextBox | 同上（TextBox版） | NGrange_Check |
| NGrange_Check | 5533 | Public Func | n,k(列範囲) | 範囲内に許容でないNGがあれば True | MinMax |
| SetSpecificTextBoxesBackColorToWhite | 5603 | Public | — | Z?/G? を白に | AorB_Click |
| IsInArray | 5618 | (Public) Func | Value, arr | 配列内存在判定 | 同上 |
| RemoveUnderlineFromLabels | 5629 | Public | — | S..E1-14, F1-10 の下線解除 | Fm_Clear, MinMax |
| MinMax | 5657 | Public | — | 指定範囲の厚/幅/丈 Min/Max 算出・下線（§5-6） | UF1 CommandButton4, Input, QP_DB |
| TB_BackColor | 5908 | Public | OP, n("S"/"E") | 範囲開始(26367)/終了(13421619)ラベル色 | UF1 OPS/OPE |
| OPcount | 5942 | Public | — | 開始>終了なら SFM.Tag="Stop" | TB_BackColor |
| UnderBar | 5961 | Public | buf("A"/"H"/"T"), n,k, Min, Max | Min/Max と一致するセルに下線 | MinMax |
| Form_Formula | 6169 | Public | — | ReSize.Tag="有" 時の上下限計算（§5-7） | UF1 |
| Atu_OPselect | 6185 | Public | ATU, WeightG | オーダー厚で作業選択ボタン表示制御、検査番号表示（§5-8） | SaveLoad |
| IM_Select | 6437 | Public | — | ライン別画像表示 | UF1 MP1_Change |
| VC_Clear | 6465 | Public | — | VC画像クリア・高さ545(4山565) | UF1, SaveLoad |
| Line_AutoSelection | 6485 | Public | n | ライン名/号機設定→Load_Text→引当判定→WorkerLoad→LoopJudg | UF1 |
| DefaultFormPosition | 6582 | Public | — | 画面中央へ（px↔pt 96/72） | UF1, 補助 |
| DistSelect | 6600 | Public | n(25/26/他) | 歪フレーム表示数（2回/3回/5回） | InspectionSelect, LSelect |
| PackingNoCheck | 6619 | Public Func | LotNo | 包装仕様書No（画像ファイル名）照合（§6） | UF1 CB_Change, UFCask, UFMitsubishi |
| Loop_Label / Label_ForeColor | 6767/6776 | Public | — | UF_PackNo の行強調 | LabelCls |
| PackSelect | 6826 | Public | n(仕様書No) | UF_PackNo へ PkgSpecData 表示 | UF_PackNo, JPGLoad |
| JPGLoad | 6876 | Public | — | 包装仕様書JPG表示＋MaterialsUpperPart メッセージ | LabelCls |
| EraseCount | 6911 | Public | Label | 消去座標と連続回数(DeleteP/DeleteC)→Posting | LabelCls, OutTolerance |
| EraseCount_F | 6936 | Public Func | n | n>=2 で True | **呼出なし** |
| frameblindfold | 6939 | Public | "目隠し"/"通常" | 「少々お待ちください」目隠し表示 | UF1, LSelect |
| OpForcedSelection | 6974 | Public | — | 作業未選択時は小型化し「作業選択を選んでください」 | UF1 |
| Six_Display | 7024 | Public | "on"/"off" | 4山(6回検査)レイアウト | UF1, LSelect, SaveLoad |
| UFsizeSave | 7213 | Public | myForm | UFsize* 保存 | UF1, OutTolerance |

外部関数の定義場所（確認のみ）: Judg系(LoopJudg/HizLoopJudg/TextJudg/HizumiJudg/UnitRecord/RightAngleJudgColor/LateralbowJudgColor/Diagonal/ReDiagonal/Lateralbow_J/ReLateralbow/BhJudg/simplyjudge/AdvanceCheck/CheckForLot/LookingYellow/MaterialsMsgBox/ReDiag)=`【UFKen_04】Judg`；OutTolerance系(Out_Tolerance/CheckLabelsByCategory/CalculateSheets/SmallNumberCount/CBAdd/Weight/WithorWithout_Gauge/Distortion_Check/Lateralbow_Check/Visual_Check/H_Check/Angle_Check/MsgError/NG_Pass/ConfiReason_Open/異常値除去/Conductiv_Check/ConductivLoad/SelectionInput/TecReturn/LcColor/TeColor/btn_allClea/GetSelectedItems/DrawListBoxContent/CB_UpDownAdd/CBG_UpDownAdd/WorkerName_Open/NameList_Open/UFPack_Open/Visual_Add)=`【UFKen_05】OutTolerance`；SaveLoad系(Posting/Load_Text/MainDataSave/AccReflec/ReflectStrain/LosdStrain/Split_Z_Tag_0/Time_Posting/UF_Array/SikaAtuCreate/OrdNoOpen/VC_Picture/AccesslogLoad)=`【UFKen_03】SaveLoad`；Worker系(Register_Check/NameConv/TimeCheck/DayCheck/YeaDayrCheck/WorkerSave/WorkerLoad/WorkeDayMax/Daily_Create/Daily_Report/EndTime_Register/Name_Match/Number_Count/IsOver12Hours/Set_Shift)=`【UFKen_07】Worker`；印刷(印刷転送/IsPrinted/SetPrinted/ClearPrinted/UpdatePrintButtonColor/CheckAndUpdateButtonColor/InitializePrintFlags/CleanOldHistory/Pesute)=`【UFKen_08】Print`；ExpCSV/履歴保存/Ainput_Startup=`【UFKen_09】Input`；校正/校正_Open/Proofread_Load=`【UFKen_10】Proofreading`；DistArray/FormReflectDist/IsGaugeRequired/DbReturnItem/GaichuPackingNo/MaterialsUpperPart=`【UFKen_11】QP_DB`；DBdisplay外観_H=`【UFKen_11】Quality_DB`；OpenMitubisi=`【UFKen_13】Mitsubishi`；GetRecordsArr/GetFieldsArr=`【UFKen_15】SQL_追加`；WriteLog=`【UFKen_16】WriteLog`；Arrbox/GetLotInfo=`【UFKen_01】AccHIki`；Machin_Text/Micro_Juge=`【UFKen_06】MachineTransfer`；その他補助(WsCheck/UR/UC/ReadCheck/AddWS/HiddenSheets/VisibleSheets/DisplaySheets/LogOut/LogOutCombo/Maker/Conv/MusicTrue/音楽を流す/音楽検査員登録/ProgressBar_count/ShowProgress/HideProgress/IsBookOpened/BookSaClose/AllClose/FindPriority/merge_sort2/TimeDesignation/FileDesignation/KenSheets/WindowMenuButtonOn/AddPagesForSubFolders/LoadImagesIntoMultiPageFromTemp/CopyRangeAsHighResolutionImage/CreatePictureFromClipboard/IsEmptyArray)=`【UFKen_17】補助`；DebugLog=`【◆】DebugOut`；パス変数=`初期設定`。

---

## 3. 画面（UF1）の構成

### 3.1 測定グリッド（Label 配列、LabelCls でクリック/ダブルクリックイベント）
- 列（検査位置）: **S / U / C / D / E / F**（F は 4山=6回検査時のみ）。列見出しラベル `S`,`U`,`C`,`D`,`E`,`F` と歪側 `FS`..`FF` のキャプションが山数・作業選択で変わる（§5-4）。
- 行（S1〜S14 等、F は F1〜F10）:
  - 行1〜4 = 厚（`AL1`〜`AL4` 見出し）
  - 通常(厚板/薄板/少枚数/スラブ): 行5,6=幅1,2、行7,8=丈1,2、行9,10=対角A,B（`L1`〜`L6` に "幅1","幅2","丈1","丈2","対角A","対角B"、FormRelat 2793-2799）。行11〜14 非表示。
  - 極厚: 行5〜10=幅1〜6、行11,12=丈1,2、行13,14=丈3,4（L1〜L6="幅1"〜"幅6"、L7/L8="丈1/2"、L12/L13="丈3/4"、2800-2810）。※行と L 番号の対応は推測（PucketMove の #L範囲と MinMax の範囲から整合）。
- 付随: `DiagS`〜`DiagF`（対角差）、`DiagSKou`〜`DiagFKou`（直角度、TextBox、見出し L10）、`LateralS`〜`LateralF`（ラテラルボー、見出し L11）、`NGS`〜`NGF`（位置ごとのNG/許容/保留表示）、`TSS..TSF/TES..TEF`（Min/Max範囲の色表示）。
### 3.2 歪（MPEX ページ2）
- 位置ごとフレーム `FMS`〜`FMF`、定盤値 `ZS`〜`ZF`、ゲージ値 `GS`〜`GF`、コンボ `CBS1..4`〜`CBF1..4`（CB?1 に "耳"/"中" 等）、ラベル `LZ?`（"定盤"/"幅ｹﾞｰｼﾞ"）、`LG?`（"ｹﾞｰｼﾞ"/"丈ｹﾞｰｼﾞ"）、A/B面切替 `AorB`（"Now_A"/"Now_B"）、`HizCome`。
### 3.3 入力系
- `Lot`（7桁）、`CB`（Lot 確定トリガ用）、`Worker`（カード13桁）、`T1/T2/T3`（直別検査員、"_"区切り）、`Day1/Day2/Day3`、`StartTB`→`INP`（測定データ受信）、`Manual`（手入力値）、`ATU/Hab/Tak`（厚/巾/丈）、`Ap/Am/Hp/Hm/Tp/Tm`（公差±）、`ApK..TmK`（上下限）、`TAIK`（対角等級）、`Lateral`（ラテラル等級）、`TaikText/LateralText`、`DebugTaikaku/DebugLate`、`Yama`（山数コンボ 1〜7）、`KenHBox1..7`（検入高さ）、`KenCount`、`SBH..EBH`、`SBad..FBad`、`DefectFactorLB`、`LBVisual`、`LC1..LC12`、`OrdNoCB`、`PasswordTex`、`DBflag`、`LoadLot/LoadWorker/LoadTextName/Reason`。
- 表示: `LineName`,`LineNo`,`KenNo`,`Reien`,`Comment`,`Youto`,`Zai`,`Tyu`,`Sumisei`,`Additional`,`PackingNo`,`PackingForm`,`PackingComment`,`flatnes`,`flatnesText`,`ConductivText`,`QF`,`LastHot`,`AMin..TMax`,`Tagtext`,`StartTime`,`EachTime`,`EachValue`,`Machin`,`PakeD`,`DeleteP`,`DeleteC`,`SplitFlag`,`AtCount`,`ResizText`,`HURIS`,`VC_A/VC_B/VC_L_A/VC_L_B`(VC画像)。
### 3.4 モード切替
- **ライン**: OptionButton `L1Work`(L-1)/`LVC`/`HVC`/`Bala`(バランサー)/`Tott`(トット)/`Test`/`Maru`(丸徳)。丸徳は号機 `A1`〜`A11`。起動時レジストリ「設備名」で自動選択（UF1 220-228）。
- **作業選択**: `厚板`/`薄板`/`少枚数`/`極厚`/`スラブ`（OptionButton、Caption に検番頭文字が入る＝`厚板.Caption`="H" 等、FormRelat 6402-6420）。
- **特殊検査（別フォームへ遷移）**: 全数/三菱受託/5N（ZEB→OpenMitubisi("全数")）、H494（H494B→OpenMitubisi("H494")）、キャスク（KYB→UFCask）。全ヵ所検査・キャスク検査シートが表示中なら UF1 は初期化/アクティブ化を拒否（UF1 174-180, 4219-4223）。
- **山数**: 1〜4 は UF1、5〜7 は「別フォームで検査」。
- **厚2/4ヵ所**（AtButt）、**検入≦25/≧26**（InspeButt）、**A/B面**（AorB）、**異常値on/off**（Abno）、**強制描写**（Byousya）。
### 3.5 画面遷移
UFProgress（進捗）、UFWorkerName（検査員表示・2度押し表示 TwiceTap）、UF校正、UF管理、UF品質基準、UFAinput（Ainput_Startup）、UFCask、UFMitsubishi（OpenMitubisi 経由と推測）、UF_PackNo（UFPack_Open）、UFConfi（ConfiReason_Open と推測）、UFNameList（NameList_Open と推測）、SF3（隠し/設定）。UF1 自身は H494/全数/キャスク遷移時 `UF1.Hide`。

---

## 4. 業務フロー（コード準拠）

1. **起動**（UserForm_Initialize 173-328）
   1. 全ヵ所検査/キャスク検査シートが Visible なら、UFCask/UFMitsubishi が非表示の場合「Loadを失敗しました…」→両シート非表示→終了。
   2. InitializePrintFlags、CleanOldHistory（コメント: 30日以上前の印刷履歴削除）。
   3. 高さ545・幅887.25、`Auto="on"`,`Tag="STOP"`,`Load="off"`,`FlagLabel=1`,`INP.Tag=""`,`PicFM.Tag=0`,`CB.Tag=""`。UFProgress 表示。
   4. レジストリ `GetSetting("UF検査判定システム","起動設定","設備名"/"モード","存在なし")`。キャプションに設備名と「オフライン」/「オンライン」。設備名=存在なし→「ライン設定がありません」→`End`（VBA全停止）。
   5. Set_Shift、Visual_Add、GetZoom。設備名でライン OptionButton を True（→ *_Click → Line_AutoSelection("")：Load_Text（作業保存読込）、GetLotInfo、JudgeLotAllocations（NeedFlat→flatnesText="ﾌﾘｰｿﾘ要"）、WorkerLoad、LoopJudg、HizLoopJudg、CheckLabelsByCategory、Load="on"）。
   6. CheckAndUpdateButtonColor、歪コンボ CBAdd（SikaAtuCreate の先頭値）。
   7. LabelCls 登録（S/U/C/D/E 1〜14、F 1〜10、CB?1〜4）、Yama リスト "1"〜"7"、BlinkLabel1(AtButt)、InitialSetting(UF1,"リスタート")、Load="on"。
2. **アクティブ化**（UserForm_Activate 4218）: 検査員未登録→WorkButton 赤。`WorkerTime_<Lname>!R1`="OUT"→「Excelを閉じても解除はできません 隠蔽工作Log出力」＋ログ＋NameList_Open(1)。Lot をロック、StartTB フォーカス、FormSetting、校正済なら ProofreadFlag="ok"。
3. **検査員登録**: WorkButton→Worker にカード13桁→直判定で T1/T2/T3 追記→T?_Change で WorkerSave・校正要求（当日・未校正なら UF校正）。
4. **ロット登録**: CommandButton10（確認・ロット分割確認）→Lot 入力（7桁）→Lot_Change（ログ、導電率）→`CB=Lot`→**CB_Change**（4494）:
   1. `CheckForLot(保存用, Lot, 2)` が True→「直近5Lot内に同じLotNoがあります」確認（No で中断）。
   2. AccReflec（引当取り込み）→CheckAndUpdateButtonColor→`CB.Tag="end"`なら終了。
   3. OrdNoOpen、PackingNoCheck→あれば UFPack_Open。
   4. DistArray(Lot, KenNo, DBflag, "")（歪/品質パラメータ）→OpForcedSelection（作業選択を促す）→FormReflectDist。
5. **作業選択**: 厚板/薄板/少枚数/極厚/スラブ→LSelect（レイアウト、2HI等判定、ゲージ要否、ラテラル/直角度/フリー反り）。薄板は山数選択。
6. **測定入力**（INP_Change 4565）:
   1. `FlagLabel<>1` or INP="" なら無視。山数未入力（Yama表示時）→警告。Yama.Tag="別フォーム"→拒否。
   2. FlagLabel=0、Fm_Clear("All")、`PakeD = Maker(INP)`（先頭2文字 "DT"/"#L"/"#〆"）。
   3. `KenNo = Left(Reien.Caption,1)`、Auto="on" なら PucketMove(PakeD, KenNo, Youto) で INP.Tag（入力先セル）決定。全部埋まっていれば「<種別>検査済」で Auto="off"。
   4. INP.Tag が空なら終了。種別 DT/#L/#〆 なら `UF1(st) = val(Conv(INP))`。
   5. 検査員未登録→「検査員登録をして下さい」・値クリア。作業未選択→「作業選択してください」。未校正(ProofreadFlag<>"ok")→ WorkButton.Tag="jump" または DBflag<>"" でなければ Proofread_Load・値クリア。
   6. LoopText(st)（異常値除去→機器判定→TextJudg→TimeRecord(2度押し)→NumDig）→Posting→Out_Tolerance。
   7. 次の空セル探索（薄板 Look_S_Blank、極厚 Look_Blank(1,14)、少枚数かつ KenCount<>"" Smallnumber_Blank、他 Look_Blank(1,10)）→Fm_Clear("Click")→INP=""、FlagLabel=1。
   - 手動: ラベルクリック→TextTag（Auto="off"、入力先固定）。ダブルクリック→LabelCls で消去＋EraseCount（同一座標2回連続で ConfiReason_Open）。
7. **歪/ラテラル/直角度/BH/検入高さ/外観** を各イベントで入力（§2.1(d)(e)）。
8. **Min/Max**（CommandButton4）→NG(許容以外)があれば「許容でない公差外が存在します」で中止。→**アイカスインプット**（AMode、Min/Max 済必須）。
9. **印刷・保存**（PrintB_Click 3619）:
   1. 2度押し防止（Enabled=False, Load="off"）。IsPrinted→PrintOnly（再印刷確認）。
   2. EndTime_Register。AdvanceCheck→「先行データ票はありますか？」(No で中止)。
   3. 必須チェック（順序どおり、最初の該当でエラー）: 定盤未入力→ゲージ未入力→歪AB(WithorWithout_Gauge)→ラテラル未入力（`ラテ<>"" And (Sumisei<>"""" And InStr(Additional,"ﾗﾃﾗﾙﾎﾞｰ検査不要")=0)`）→直角未入力（Youto="R115"）→外観未入力→検入高さor枚数(H_Check="入力なし")→山数未入力。
   4. UFProgress、音声、目隠し。Name_Match（登録検査員と実績の一致）、Number_Count（梱包枚数チェック）。PrintB.Tag="stop"なら中止。
   5. 印刷転送（印刷原紙へ出力）。PrintB.Tag="stop"→キャンセル。
   6. 不良要因印刷_<Lname> が無い/別ロットなら原紙からコピー作成。
   7. 「印刷する？」Yes→ログ（T1/T2/T3）→`印刷用_<Lname>` と `不良要因印刷_<Lname>` を PrintOut(1部)→MainDataSave("保存用")→履歴保存→CSV出力（§6）→ThisWorkbook.Save→SetPrinted→ボタン色→完了音。No→キャンセル。
   8. 最後に Application.Visible=False。
10. **新規検査**（NewKen_Click 3359）: ThisWorkbook.Save→Lot が保存用（無ければ耳用）シートB列に無ければ「該当LotNoが保存されていません」（MsgError "END"、詳細は MsgError 側）→確認→ログ「新規検査実行」→作業保存_<Lname>/WorkerTime_<Lname> 新規（AddWS）、印刷用/異常ロット用/不良要因印刷_<Lname> 削除、当日以外なら検査員_<Lname> 新規→ClearLSelectCache→歪ラベル/コンボ・Min/Max・各表示クリア→Split_Z_Tag_0(S..F)→A面に戻す→LosdStrain→Load_Text("ReLoad")→WorkerLoad→VC_Clear→AtCountSelect(4)→InspectionSelect("表示")→RightAngleDisplay("新規")→LSelectLateral("新規")→WorkerName_Open→InitialSetting("リセット")＋InitialSetting。※Lot="" の場合は確認なしで line: へ GoTo（3368）。
11. **終了**（QueryClose）: Posting、シート VeryHidden、Excel最大化表示、Save。

---

## 5. 判定ロジック・計算式・閾値・業務ルール（原文）

### 5-1. 定数・色
```vba
Public Const THIN_PLATE_ATU As Double = 6# '薄板判定閾値(mm)_6mm以下でAtuFlag=True   (FormRelat 3)
Dim Scol As Long: Scol = 13434828 'OK RGB(204,255,204)薄い緑
Dim Ncol As Long: Ncol = 8421631 'NG RGB(255,128,128)ｺｰﾗﾙ        (FormRelat 1845-1846 他)
```
`NG?` ラベルの Caption が `"許容"` の場合、NG色でも MinMax を止めない（FormRelat 5499, 5527）。保留は止める（UF1 3664-3666 コメント）。

### 5-2. 表示桁（NumDig, FormRelat 345-428）
- 厚(行1-4): `val(ATU) > 74#` → `"0.00"`、それ以外 `"0.000"`。
- 幅(行5-6): `UF1.Youto = "H494" And UF1.Hab < 300#` かつ `val(Hab)` が `1# To 299#` → `"0.00"`（ノギス測定、Q-S-QA-MSR-007-R04）、それ以外 `"0.0"`。
- 丈(行7-12, F は7-10): `"0.0"`。対角 Diag?・直角 Diag?Kou: `"0.0"`。

### 5-3. 入力先自動決定（PucketMove 716-1016, AllCheck 429-715）
- 検番フラグ（共通定義、AllCheck 587-592 / PucketMove 739-744 / AtCountSelect 1989-1993 / InspectionSelect 2222-2226）:
```vba
If KenNo = "L" Or KenNo = "N" Or KenNo = "K" Or KenNo = "T" Or KenNo = "S" Or KenNo = "F" Then flag(1) = True
If KenNo = "H" Then flag(2) = True
If UF1.Yama = 3 Then flag(3) = True
If flag(1) And UF1.InspeButt.Caption = "検入≧26" Then flag(4) = True '検入25枚以下
If KenNo = "" Then flag(5) = True
If UF1.Yama = 4 Then flag(6) = True
```
  （Look_S_Blank では flag(1) に `KenNo=""` も含み、flag(5) は未設定、1027-1031）
- 種別→対象行:
  - `"DT"`: AtButt.Caption `"厚2ヵ所"`→行1-4、`"厚4ヵ所"`→行1-2。`Youto="H494" And Hab<300#`→行1-6。（Caption がどちらでもなければ a=n=0→AllCheck=False→「DT検査済」：要確認）
  - `"#L"`: 極厚→行5-14、他→行5-10。H494かつHab<300→行7-10。
  - `"#〆"`: 極厚→1-14、他→1-10。他は「ﾊﾟｹｯﾄ無し」。
- AllCheck の未入力判定列（Temp="厚板"）: 行5-10 は S/C/E、他行は S/U/C/D/E。「少枚数」: th=5→S,U,C,D,E／4→(行5-10:S,C,D 他:S,U,C,D)／3→S,U,C／2→S,U／1→S。薄板（Temp=""）: flag6→6列、flag3 or flag5→5列、flag4→S,E、flag1→S,C,E（flag2のみは該当なし）。少枚数 or 極厚→5列。スラブ: 行1-4 S,C、行5-10 S,U,C,D。
- 厚板検番H×少枚数で Temp="少枚数"（787-790）。
- 探索順序は S→U→C→D→E→(F) の列優先・行昇順。厚板は行5-10 で U の代わりに C、D の代わりに E を探す（スラブは U/D 行5-8）。詳細は 869-1015 を逐語移植要。
- `Look_S_Blank` の「厚4ヵ所」分岐は `For` ループ内で `i = i + 2`（i=3→5）/`i = i + 1` によりループ変数を書き換える（1037-1047）。逐語移植要。

### 5-4. 位置ラベル（S,U,C,D,E,F 列見出し・OPS/OPE キャプション）
| 条件 | S | U | C | D | E | F |
|---|---|---|---|---|---|---|
| 厚板（住精厚板も同じ） | S | S1 | 2 | E1 | E | |
| 少枚数/極厚 | 1 | 2 | 3 | 4 | 5 | |
| 薄板 1山 | S | "" | 2 | "" | E | |
| 薄板 1山 検入≧26（≦25運用）| S | "" | "" | "" | E | |
| 薄板 2山 | 2-2-S | "" | 2-2-E | "" | 2-1-E | |
| 薄板 2山 検入≧26 | 2-2-S | "" | "" | "" | 2-1-E | |
| 薄板 3山 | 3-3-S | 3-3-E | 3-2-2 | 3-1-S | 3-1-E | |
| 4山（Six_Display "on"） | 4-4-S | 4-4-E | 4-3-2 | 4-2-2 | 4-1-S | 4-1-E |
| 薄板で Yama 非表示 | S | U | 2 | D | E | |
| スラブ | 1-上 | 1-下 | 2-上 | 2-下 | "" | (FS..: 1,"",2,"","") |
（出典: LSelect 2810-3007, Yama_Change 2280-2396, InspectionSelect 2235-2480, Six_Display 7065-7090, SlabSet 5190-5212。OPS5 の 2山 "2-2-E" 差異は §0）

### 5-5. LSelect の業務判定（2579-3864）
- `UF1.Byousya`（強制描写）なら DB 参照をスキップして `LineSkip:`（Case "薄板" ブロック内部 2897）へ GoTo。
- キャッシュ: `g_CacheLotNo = Lot And g_CacheCours<>"" And g_CacheOrdAtu<>0` なら g_CacheKenNo/g_CacheCours/g_CacheOrdAtu を使用。
- 非キャッシュ時: `Arrbox("SIKALOTNOW.accdb","板","Fields",DBflag)` で列位置（設計_設備ｺｰｽ/ｵｰﾀﾞｰ板厚/実績_設備ｺｰｽ）、`Arrbox("SIKALOTNOW.accdb", Lot, "Record", DBflag)`。無ければ（Load="on" かつ IsOver12Hours=False かつ g_SkipLotNotFoundMsg=False のとき）「SIKALOTNOWに該当ロットがありません」。検査番号は SIKAHIKINOW→無ければ SIKALOTNOW の「検査番号」先頭1文字。`KenNo="0"or""` なら `FindPriority(実績_設備ｺｰｽ)`（最終圧延: CM1→L, CM2→N, CM3→K, 2HI→T, 無ければ HOT→H。補助 1455-1490）。
- コース判定（設計_設備ｺｰｽ文字列）:
```vba
Case InStr(strCoursForCheck, "2HI") > 0: HIFlag = True: Reien = KenNo & "_2HI": Comment = "2HI材:厚測定強制4ヵ所"
Case InStr(strCoursForCheck, "L-2") > 0: FS4flag = True: Reien = KenNo & "_L-2" (& "_2HI")
Case InStr(strCoursForCheck, "L-3") > 0: FS4flag = True: Reien = KenNo & "_L-3" (& "_2HI")
If FS4flag Then If InStr(strCoursForCheck, "LL7") > 0 Then Reien = KenNo & "_LL7"
```
  `SSW` 判定（3540）は算出のみで未使用（旧ゲージ条件で使用）。
- オーダー厚 `NewAtu`（0なら val(ATU)）`<= 6` で AtuFlag。薄板: 2HI 未通過かつ 6mm超→AtButt 非表示、2HI 通過→AtButt 非表示（厚強制4か所）、検番T→InspeButt 非表示。
- ゲージ要否（作業選択5種すべて、3572）: `IsGaugeRequired(val(ATU), val(Hab), val(Tak), Zai, Tyu, コース)` True→LZ?="幅ｹﾞｰｼﾞ"、LG?="丈ｹﾞｰｼﾞ"、CBAdd(SikaSetAtu)、blindfold.Tag="Set"／False→"定盤"/"ｹﾞｰｼﾞ"。条件は品質基準.accdb「ゲージ判定条件」（コメント 3570-3571）。旧ハードコード条件（参考、3545-3560 コメント）: 重量>1300kg&丈<2500、重量>2160&2500≦丈<4000、重量>2500、幅<950&重量>1300、丈>5000、調質F&SSW。
- 直角度: `Youto <> "R115"` なら L10/Diag?Kou 非表示（QA-S-QA-DIM-154）。
- 引当判定（bSkipDBCheck=False 時）: GetLotInfo("Lot情報履歴", Lot)→JudgeLotAllocations→NeedLateral→LSelectLateral(検番)、NeedRight→RightAngleDisplay(検番)、NeedFlat→flatnesText="ﾌﾘｰｿﾘ要"。さらに `LateralText.BackColor = RGB(51,204,204) And LateralText.BackStyle = 1`（カードに公差あり、2024.8.27 家木指示）→LSelectLateral。
- JudgeLotAllocations（5259-5333）: TempArry 列 = 1納入先,2得意先,3用途ｺｰﾄﾞ,4材質,5調質,6厚,7巾,8丈,9ﾛｯﾄNo,10検査No（GetLotInfo）。納入先が空の行は無視、8項目キー重複は1回のみ DbReturnItem。戻り文字列に "ラテラルボー"/"直角度"/"フリー反り" を含めば各 Need=True。検番=**最終行**の検査No先頭1文字。Additional に "ﾗﾃﾗﾙﾎﾞ-追記公差あり" があれば Additional/LateralText を元に戻す、NeedLateral 時は "ﾗﾃﾗﾙﾎﾞｰ検査不要" を Additional から消す、NeedFlat なら HURIS 保持。
- LSelectLateral: 極厚→Additional="ﾗﾃﾗﾙﾎﾞｰ検査不要_耳付き"で終了。Additional に "ﾗﾃﾗﾙﾎﾞｰ検査不要"→終了。ATU="" →終了。検番H→U,D 非表示、N/K/T/S/F で山数1,2→U,D 非表示、F は4山のみ。
- RightAngleDisplay: Sumisei<>""→表示。検番 N/K/T/S/F：3山→DiagU/D 表示（R115 なら Kou も）、1,2山→非表示。

### 5-6. MinMax（5657-5906）
- 範囲: OPS1〜6 の最初の True=S_flag、OPE1〜6=E_flag。0 または S>E は中止。`NGrange_Check(S_flag,E_flag)`（行1〜8、極厚1〜12、Z?/G?、Sumisei<>"" ならラテラル、Sumisei<>"" かつ R115 なら直角）が True→「許容でない公差外が存在します」で中止。
- 厚 = 行1-4 の範囲内全値。幅: 極厚 行5-10、他 行5-6。丈: 極厚 行11-14、他 行7-8。
- 例外: 厚板で S=E=2(U)→幅丈は列1(S)、S=E=4(D)→列3(C)。少枚数 th=5 同様、th=4 は S=E=2→列1、S=E=4→列4。
- 書式: `If 0 < val(UF1.ATU) <= 74# Then` AMin/AMax `"0.000"`（※常に真、§0）、HMin..TMax `"0.0"`、H494&Hab<300&1〜299 は HMin/HMax `"0.00"`。値0は空文字。下線（UnderBar）は Caption と Min/Max の `=` 一致。
- 空白セルは空のまま Min/Max に渡す（WorksheetFunction.Min は空文字を無視：VBA/Excel 仕様依存）。

### 5-7. 計算式
```vba
.ApK = Format(val(.ATU) + val(.Ap), "0.000"); .AmK = Format(val(.ATU) - val(.Am), "0.000")
.HpK = Format(val(.Hab) + val(.Hp), "0.0");  .HmK = Format(val(.Hab) - val(.Hm), "0.0")
.TpK = Format(val(.Tak) + val(.Tp), "0.0");  .TmK = Format(val(.Tak) - val(.Tm), "0.0")   (Form_Formula 6169, ReSize.Tag="有"時のみ)
```
- 少枚数枚数: `th = Format(Round(CalculateSheets(UF1, UF1.ATU), 1), "0")`（KenHBox7="" のとき、それ以外 KenCount）、`If th >= 5 Then th = 5`（UF1 735/757/1700/4397/5251, FormRelat 434-441, 5705-5712）。※UserForm_Click(4396) のみ条件が `KenCount = ""`。
- 歪ページ: `TempW = Weight(val(ATU), val(Hab), val(Tak))`; `歪AB<>""` のとき `35 >= TempW`→「ゲージ不要 * 両面定盤歪測定必要」、`35 < TempW`→「ゲージによる測定必要」、B面かつ >35→「B面歪 入力画面です」（UF1 1651-1662）。
- 2度押し: `(UF1(st) = UF1.EachValue) And (DateDiff("s", CDate(EachTime), CDate(NowTim)) <= 4)` → 値クリア＋TwiceTap 1秒表示（FormRelat 317）。
- 消去カウント: 厚板/薄板/少枚数で行9/10 は同一座標 "?10" とみなす（6919-6922）。
- ラテラル/直角: `UnitRecord(val(x))`（0.5ピッチ変換）→ `"0.0"`。"-" は測定不要として保持。
- 不良枚数 >50 は入力拒否（"ｽﾎﾟｯﾄ処理"除く、UF1 1570-1574）。
- 対角/ラテラル基準値: 等級 "4"→Diagonal/Lateralbow_J（4級厳しい）、"3"→ReDiagonal/ReLateralbow（UF1 4173-4200）。

### 5-8. オーダー厚による作業選択ボタン表示（Atu_OPselect 6185-6435）
- オーダー厚 NewAtu（SIKALOTNOW ｵｰﾀﾞｰ板厚）と ATU が異なれば Comment="ｵｰﾀﾞｰ厚参照_" & Format(NewAtu,"0.000")。
- 検査番号: SIKAHIKINOW.accdb 仕掛 `[ﾛｯﾄ番号]='LotNo'` の検査番号→無ければ SIKALOTNOW。`KenNo.Caption` に全体、先頭1文字を検番に。"0"/"" なら FindPriority（見つかれば「最終圧延機から検査番号頭を生成します…ラテラルボーが必要の可能性があります Q-S-QA-ALL-088_検査番号の付け方」、無ければ「データ上にまだ検査番号がありません」）。
- 検番 "" → 終了。検番 "T" → 薄板のみ表示（Caption="T"）。
- その他（`Select Case NewAtu`）:
  - 厚板 非表示: `1 To 6#`（このとき 薄板.Caption=検番）、`200.1 To 299#` かつ WeightG<>0 かつ WeightG>1500#、`300# To 800#`
  - 薄板 非表示: `6.001 To 800#`（常に 厚板.Caption=検番）
  - 少枚数 非表示: `1 To 6#`
  - 極厚 非表示: `1 To 29.9`
  - スラブ 非表示: `1 To 199.9`
  - 境界の隙間（6〜6.001、29.9〜30、199.9〜200、200〜200.1、299〜300、<1、>800）は非表示にならない。

### 5-9. ハードコード業務ルール一覧
- 用途コード `"H494"`: 幅をノギス測定（Hab<300 かつ 1〜299）→幅2桁小数、DT で行1-6、#L で行7-10（NumDig/PucketMove/MinMax）。H494 専用検査フォーム（H494B）。
- 用途コード `"R115"`: 直角度必須（印刷時チェック、表示、NGrange_Check）。
- 検番頭: L/N/K/T/S/F=冷延系（3回検査系）、H=熱延（5回）、T=2HI材（薄板のみ・厚強制4か所・InspeButt非表示）。
- コース: "2HI","L-2","L-3","LL7","SSW"。
- 山数5以上は別フォーム。丸徳は ZEB キャプション「検入4山以上」。
- バランサーは Lot登録・検査員登録の確認ダイアログ省略。
- 極厚（耳付き）は保存先が「耳用」シート・CSV名 "_耳用履歴CSV"/"_耳用File"、ラテラル検査不要。
- 歪コンボ1段目 "耳"/"中" は Z 値に文字入力可（maxLength 5）。
- ロット7桁、カード13桁（コメント「ｶｰﾄﾞ16桁　-4000番」）。
- パスワード "weekend99"（隠し機能）、"jump"（校正スキップ）。
- 不良要因 "ｽﾎﾟｯﾄ処理" は50枚制限除外。
- 直近5ロット重複警告（CheckForLot 第3引数=2列目）。
- 先行データ票（AdvanceCheck）が必要なロット: 初回クリック・検査員登録時に「先行データ票が必要です。」、印刷時「先行データ票はありますか？」。
- 導電率（Conductiv_Check）要ロットは Conductiv ボタン赤。

---

## 6. 外部依存

### 6.1 レジストリ
`GetSetting("UF検査判定システム","起動設定","設備名")`（L-1/LVC/HVC/バランサー/トット/Test/丸徳）、`"モード"`（オンライン/オフライン）。UF1 199-200。

### 6.2 Access DB（GetRecordsArr/GetFieldsArr 経由。SQL は WHERE 句相当）
| DB | パス | テーブル | 条件 | 使用箇所 |
|---|---|---|---|---|
| SIKALOTNOW.accdb | Arrbox: DBflag="" なら MainPath（`\\nlmsrvngy03\Read\【New】仕掛\台帳\`）、ファイル無/TimeDesignation/FileDesignation=False なら NextPath（`…\梱包課共有\仕掛\`）、DBflag<>"" なら NextPath | 仕掛 | `[ﾛｯﾄ番号]='<Lot>'`／"板"指定は `[ｵｰﾀﾞｰ板丈]<>0`（Fields 取得用） | LSelect, AtCountSelect, InspectionSelect, Atu_OPselect |
| SIKAHIKINOW.accdb | 同上 | 仕掛 | `[ﾛｯﾄ番号]='<Lot>'` | 同上（検査番号） |
| 伝送用ファイル.accdb | PackingPath（`\\nlmfangyshrd\各課共有\0130_日軽稲沢\梱包課\AIM\【■】_参照用ファイル\`） | PkgSpecData | `[包装仕様書№]='<No>'` | OrdNoCB_Change, PackSelect |
| 品質基準.accdb（推測: QualityPath） | — | ゲージ判定条件 | IsGaugeRequired 内 | LSelect（コメント 3570） |
| DbReturnItem / DistArray / DBdisplay外観_H の DB | QP_DB/Quality_DB 側 | 不明（担当外） | | JudgeLotAllocations, CB_Change, CommandButton8 |
フィールド: 設計_設備ｺｰｽ, ｵｰﾀﾞｰ板厚, 実績_設備ｺｰｽ, 検査番号, ﾛｯﾄ番号, ｵｰﾀﾞｰ板丈 / 梱包形態, タイトル, 包装仕様書№, 梱包指定別, 用途名。

### 6.3 ファイル・フォルダ
- ログ（WriteLog）: ローカル（既定先は WriteLog 側）＋ `HistoryPath & "<ライン>\<ライン>_Log"`（ライン: L-1/LVC/HVC/バランサー/トット/Test/丸徳）、Lot登録は `Measured.log` にも。内容例: `"Lot登録:" & Lot & ":" & Lname`、`"*UFKen:山数ﾁｪﾝｼﾞ:" & n`、`"*UFKen:印刷 T1:… T2:… T3:…"`、`"UFKen*新規検査実行"`、`"*UFKen:検査員登録_" & Day & ":" & T`、`"*UFKen:UFNameList起動隠蔽対策:Excel再起動の形跡"`、`"*UFKen:LateralS_BeforeUpdate:<値>:<Lname>"` 等。
- CSV（ExpCSV、PrintB 3772-3802）:
  - ローカル: `Lname & "_履歴CSV"`（極厚 `"_耳用履歴CSV"`、出力先は ExpCSV 既定＝要確認）
  - 履歴: `HistoryPath & "<ライン>\<ライン>_履歴\" & FileNewName`（FileNewName = `Lname & "_File"` / 極厚 `Lname & "_耳用File"`）。Test は `Test\Test_履歴`、存在確認ファイル名は `Test_File.csv`。IsBookOpened で開いていないときのみ。
  - 工場共有: `FactoryyPath & "【梱包】L-1|LVC|HVC|大板(バランサー,トット)|中板(丸徳)"`（Test は出力なし）。
  - ※IsBookOpened のチェック対象名は常に `_File.csv`（耳用でも）。
- 包装仕様書画像: `QualityPath & "包装仕様書"` フォルダ全ファイル（拡張子位置7＝6文字名のものは先頭6文字、他はファイル名全体）と引当の包装仕様NO照合（PackingNoCheck）、表示は `n & ".JPG"`（JPGLoad）。※コメント上の旧パスは `\\nlmsrvngy03\工場内共有\検査データ\包装仕様書`。QualityPath は `…\品証-梱包やり取りフォルダ\Accdb\` なので実パスは `…\Accdb\包装仕様書`（要確認）。
- 品質写真: LoadImagesIntoMultiPageFromTemp（QualityPath & "品質写真\"）。
- 音声: `C:\VBAWAV\` の 歪の公差.wav, ありますか.wav, おまちください.wav, 印刷.wav, 完了.wav, 必要です.wav（音楽を流す/MusicTrue、UF1.Tag="STOP" で無音）。
- 実行: `C:\Windows\System32\ctfmon.exe`（WScript.Shell.Run, 待機）。
- 他ブック: `日報＆資材計算.xlsm`（BookSaClose / AllClose で閉じる）。
- Windows API: クリップボード(OpenClipboard 等)、OleCreatePictureIndirect、Sleep、GetSystemMetrics、Get/SetWindowLong、DrawMenuBar。
- プリンタ: PrintOut（既定プリンタ、1部）、`Application.PrintCommunication = False`。

### 6.4 ブック内シート
全ヵ所検査, キャスク検査, Lot情報履歴（ﾛｯﾄNo/ｵｰﾀﾞｰNo/包装仕様NO/引当NO/材質/調質/ｺｰｽ/厚/用途ｺｰﾄﾞ/納入先/得意先/巾/丈/検査No）, 保存用（B列ロット）, 耳用, 作業保存_<Lname>（O2:P7=A面歪, O8:P13=B面歪, T1=作業選択(Load_OP)）, WorkerTime_<Lname>（R1="OUT"、C1=日時[IsOver12Hours]）, 検査員_<Lname>, Daily_<Lname>, 印刷用_<Lname>（D2="再印刷"）, 不良要因印刷_<Lname>（B1=Lot、C3:X8/C10:X10/C11）, 不良要因印刷原紙, 異常ロット用_<Lname>（A1=Lot）, 導電率測定, 校正記録, 印刷用_Test。`<Lname>` = `LineName.Caption & LineNo.Caption`。

---

## 7. 入力データ / 出力データ

- **入力**: 設備名/モード（レジストリ）、社員カード13桁、ロットNo 7桁、測定伝送文字列（例コメント: `DT10002+0000021.712M`＝マイクロメータ等、`#L395.8`＝コンベックス、`#〆<値>`＝手入力）、歪（定盤/ゲージ、A/B面、コンボ位置）、ラテラルボー、直角度、BH、検入高さ1〜6/検入枚数、少枚数枚数、外観ボタン・不良要因・不良枚数、山数、厚2/4、検入≦25/≧26、作業選択、Min/Max範囲、オーダーNo選択、公差手動編集値、パスワード。
- **DB/シート由来**: オーダー厚、検査番号、設備コース、引当（納入先/得意先/用途/材質/調質/厚/巾/丈/検査No）、包装仕様、歪・品質パラメータ（DistArray）。
- **出力**: 画面判定（色）、印刷物（印刷用_<Lname>、不良要因印刷_<Lname>、導電率測定、再印刷時 D2="再印刷"）、保存用/耳用シート（MainDataSave）、履歴保存、CSV（ローカル・履歴・工場共有）、ログ、作業保存/WorkerTime/検査員/Daily シート（Posting/Time_Posting/WorkerSave/Daily_Report 経由）、印刷済フラグ（SetPrinted）、ThisWorkbook.Save。

---

## 8. エラー処理

- 構造化エラー処理はほぼ無く、`On Error Resume Next` を広範囲に使用（UF1 内 65か所、FormRelat: InitialSetting 65, LSelect 2585（**LSelect 全体**）, DefaultFormPosition 6583, OpForcedSelection 6978）。LSelect では `LotAcc` が Empty のまま `LotAcc(1, C_Course)`/`LotAcc(1, OrdAtu)` を参照する箇所（2703, 2773）があり、エラーは握りつぶされる（Resume Next ではエラー行の代入自体がスキップされるため NewAtu は 0 のまま→`val(UF1.ATU)` にフォールバック）。
- `On Error GoTo myError` はコメントアウト（FormRelat 279）。
- ユーザー通知は MsgBox と MsgError(Trm, Lname, n, Temp)（OutTolerance 側、n=1/2/4 の意味は担当外）。MsgError("…", Lname, 2, "END") は NewKen で使用。
- リトライ: ViewCB_Click の画像コピー（最大3回）・クリップボード取得（最大8回、Sleep 300/500ms）。
- `End` 文（設備名未設定時、UF1 206）で VBA 全体停止。
- 2度押し対策: PrintB.Enabled/Load.Caption、TimeRecord の4秒判定。
- エラー時の状態復帰漏れ: PrintB の各中断パスでは Load="on"・Enabled=True を戻しているが、`MsgError(…,"END")` 等で処理が止まった場合の挙動は MsgError 依存（要確認）。

---

## 9. グローバル変数・状態保持（長時間放置・日付変更に関係）

### 9.1 モジュール変数
- UF1: `Old_CBS1..Old_CBF4`（コンボ前回値、ログ差分用）、`LS/LU/LC/LD/LE/LF/ComS..ComF`（LabelCls）、`BlinkLabel1`、`Lname`（モジュール Dim。各Subでローカル再宣言され隠蔽）、`SikaAtuAll/SikaSetAtu`（Private、SaveLoad の Public を隠蔽）、`NowSizeH/W/MP`、`g_sng*SV/g_lngWindowState`（未使用）。
- FormRelat Public: `UfZoom/UfHeight/UfWidth`, `UFsizeWidth/Height/Zoom`, `g_FormHwnd`, `g_SkipLotNotFoundMsg`。参照のみ: SaveLoad の `g_CacheLotNo/g_CacheKenNo/g_CacheCours/g_CacheOrdAtu/g_CacheC_Course`、`SikaSetAtu/SikaAtuAll`、Quality_DB の `GradeArray`。
### 9.2 コントロールに持たせた状態（主要）
| 状態 | 値 | 意味 |
|---|---|---|
| `Load.Caption` | "on"/"off" | フォームイベント抑止。off のまま残ると多くの処理が Exit。UserForm_DblClick・MPEX_Change・CB?_Exit・各処理末尾で on |
| `UF1.Tag` / `UFTag` | "STOP"/"" | 音声停止・MsgBox抑止（Load_OP 等） |
| `FlagLabel` | 1/0 | Change イベント再入防止 |
| `Auto` | "on"/"off" | 次入力セル自動決定の有無（ラベル手動選択で off） |
| `INP.Tag` / `Tagtext` | セル名 | 現在の入力先 |
| `PicFM.Tag` | 数値 | UserForm_Click 回数（1回目のみ先行データ票警告・SmallNumberCount）。検査員登録・MP1ページ1でリセット |
| `MPEX.Tag` | "Click"/"" | 判定表示抑止（LateralbowJudgColor 対策） |
| `Loading.Tag` | "Loading" | AccessLog 停止 |
| `CB.Tag` | "end" | CB_Change 中断 |
| `Yama.Tag` | "UF1"/"別フォーム" | 山数5以上の入力拒否 |
| `WorkButton.Tag` | "jump" | 校正スキップ（校正() 呼出でクリア） |
| `ProofreadFlag` | "ok"/"" | 校正済。Activate 時と 校正() で更新 |
| `EachTime/EachValue` | 時刻/値 | 2度押し判定 |
| `DeleteP/DeleteC` | 座標/回数 | 連続消去 |
| `PrintB.Tag` | "stop" | 印刷中断 |
| `ReSize.Tag` | "有" | 公差編集モード |
| `blindfold.Tag` | "Set" | 歪コンボ追加済 |
| `SFM.Tag` | "Stop" | Min/Max 範囲不正 |
| `DBflag` | 任意文字列 | DB参照先 NextPath 切替・校正スキップ・AMode無条件（用途は要確認） |
| `Sumisei.Caption` | 納入先名等 | ラテラルボー要フラグ（旧: 住友精密） |
| `Additional.Caption` | "ﾗﾃﾗﾙﾎﾞｰ検査不要"等 | 技術指示 |
### 9.3 タイマー・ループ
- `Application.OnTime`・API タイマーは**どちらのモジュールにも無い**。
- DoEvents は UF1 内 36か所（進捗表示・フォーカス対策）。`Sleep` 使用: AorB_Click 1000ms、TimeRecord 1000ms、ViewCB 300〜1000ms。
- clsBlinkLabel の Blink（DoEvents＋Sleep500 の無限寄りループ）は `BlinkLabel1.Blink 2` がコメントアウト（UF1 4269）のため実質不使用（BlinkStop のみ呼ばれる）。
### 9.4 長時間放置・日付/直変更時に影響しうる点（推測含む）
- `ProofreadFlag` は Activate 時か T1〜T3 変更時（校正 経由）にしか再評価されない。UF1 を開いたまま直（シフト）が変わると、再アクティブ化されるまで "ok" が残り INP_Change の校正チェックをすり抜ける可能性（推測）。
- `Register_Check(Format$(Now,"hh:mm:ss"))` は毎回現在時刻で評価（INP_Change・Click・Activate）→直が変わると「検査員登録をして下さい」になる（Worker 側仕様）。
- 検査員は `TimeCheck` で 日勤/1直→T1、2直→T2、3直→T3 に振り分け、`DayCheck` で Day1〜3 に日付。夜勤の日付跨ぎは Worker 側（YeaDayrCheck）依存（コメント「夜勤0時がMaxとならない為停止」）。
- `IsOver12Hours`（名称に反し WorkerTime!C1 から **1時間以上**経過で True）が True の間は LSelect の「SIKALOTNOWに該当ロットがありません」を抑止。
- LSelect キャッシュ（g_Cache*）はロットNo一致で再利用、新規検査で ClearLSelectCache。DB更新（1日更新のNextPath等）後も同一ロットならキャッシュ値を使い続ける。
- `EachTime` は放置後でも4秒判定なので影響なし。
- Line_AutoSelection の早期 Exit で Load="off" が残ると、その後の多くのイベントが無反応（UserForm_DblClick 等で復帰）。

---

## 10. 未使用・死んでいそうなコード（候補、削除可否は断定しない）

| 対象 | 根拠 |
|---|---|
| FormRelat 3866-5102（約1240行） | 旧 `LSelect` 全体がコメントアウト（非コメント行は空行1行のみ） |
| UF1 3850-4055 | 旧 `PrintB_Click` がコメントアウト |
| `Load_OP`（5455） | 全モジュールで呼出なし。さらに作業保存シートの T1 はサンプル(`作業保存_Test.csv`)で "3"、作業選択文字列は AC1 にあり列位置が合わない |
| `CBdateCheck`（1935） | 呼出は旧PrintB（コメント内）のみ |
| `EraseCount_F`（6936） | 呼出なし |
| `GetScreenWidth/GetScreenHeight`（49/53） | 呼出なし |
| `auto_open`（269） | 本文が全コメント |
| `UF1_Close`（273） | ブック内呼出なし（Application.Run 用とコメント。他ブックからの呼出可能性あり） |
| `UserForm_Resize`, `StartTB_AfterUpdate` | 本文全コメント |
| `Day1/2/3_Change` | どの分岐でも Exit Sub のみ |
| `KenHBox7_AfterUpdate` の th 計算 | 結果未使用（SmallNumberCount はコメント） |
| `TestButton_Click` | テスト用（印刷用_Test 表示） |
| MinMax 5805-5807 の ElseIf | §0 のとおり到達不能 |
| LSelect `SSW_F`（3538-3541） | 算出後未使用 |
| `FlagColor` 等の一部変数、`FS4flag` 後の `LL7Flag`、`ResultT`の一部、`Sfl/AFl/x/p` 等 | 宣言のみ/未参照 |
| UF1 `g_sngTopSV` 等、FormRelat `g_sng*SV` | 未使用 |
| `PacN`（PackingNoCheck） | デバッグ用配列（コメント） |
| `Look_S_Blank` の flag(5) | 代入されないため常に False |
| Atu_OPselect 内 `UF1.LHIM.Visible=False` 重複等（IM_Select 6438-6443） | 冗長 |

---

## 11. 不明点・要確認事項

1. `INP_Change` の検番を `Reien.Caption` から取るのは意図か（UserForm_Click は KenNo.Caption）。Reien 空時に AllCheck=5回扱い・Look_S_Blank=3回扱いになる差異の業務的正否。
2. `PrintB_Click` の `Sumisei.Caption <> """"`（常に真）を移植で踏襲するか、`<> ""` に正すか。
3. `MinMax` の `0 < val(ATU) <= 74#` 常真バグ（ATU>74 でも厚 Min/Max 3桁小数）を踏襲するか。NumDig は 74 超で2桁なので表示とMin/Maxで桁が不一致になる。
4. `InspectionSelect` 2山の OPS5="2-2-E"（他は "2-1-E"）の正解。
5. `FBad_MouseDown` が EBad に書く件。
6. `AtCountSelect` 検番T 6回時の F 列 Enabled 未設定。
7. PucketMove "DT" で AtButt.Caption が "厚2ヵ所"/"厚4ヵ所" 以外（厚板等で AtButt 非表示時）のときの Caption 実値と挙動（a=n=0→「DT検査済」になるか）。デザイナ上の初期 Caption 不明。
8. 薄板で検番H・山数1/2のとき AllCheck が該当分岐なし（False＝検査済）になる点。
9. `Line_AutoSelection` 早期 Exit で Load="off"/Tag="STOP" が残る件の運用影響。
10. `DBflag` の業務上の意味（履歴参照モード？テストモード？）と、誰がいつ設定するか（UF1 コード内では代入なし）。
11. `Sumisei`（住精）フラグの現行意味（コメント上「ﾗﾃﾗﾙﾎﾞｰ測定の納入先」全般）、`LateralText.Tag="有"`（技術指示経路）の設定元。
12. `UF1_Close` を他ブック（日報＆資材計算.xlsm 等）が Application.Run で呼ぶか。
13. `IsGaugeRequired` / `DbReturnItem` / `DistArray` の参照DBファイル・テーブル・閾値（担当外モジュール）。
14. ExpCSV の既定出力先（ローカル `_履歴CSV`）と文字コード・区切り。
15. `MsgError` の n（1/2/4）と "END" 指定時の挙動（処理停止・End するか）。
16. PackingNoCheck の実フォルダ（QualityPath は `…\Accdb\` で終わるため `…\Accdb\包装仕様書` になる）が正しいか。
17. UF1 のコントロール初期値（Caption/Visible/Value）、`Inspection1..4`（検査法）のキャプション、`LC1..LC12` と位置の対応、`BH` の業務意味（板の何の測定か）。デザイナ情報が必要。
18. `CommandButton3_Click` が Manual の書式を `Len(INP)`（通常空）で判定している意図。
19. 少枚数 th 算出の条件が UserForm_Click だけ `KenCount=""` で他は `KenHBox7=""` である差異。
20. `Round`（銀行丸め）→`Format "0"`（四捨五入）→Long 代入の二段丸めを Python でどう再現するか（`decimal` で VBA 仕様を再現する必要）。
21. `Select Case NewAtu` の境界の隙間（6〜6.001 等）にあるオーダー厚の実在有無。
22. Atu_OPselect の `WeightG` を渡す呼出元（SaveLoad）での値の定義。
23. 2HI 判定は「設計_設備ｺｰｽ」、FindPriority は「実績_設備ｺｰｽ」を使う使い分けの正否。
24. LSelect の `IIf(bUseCache, g_CacheOrdAtu, LotAcc(1, OrdAtu))` がキャッシュ時に実質 val(ATU) を使う挙動（§0-5）を踏襲するか。
25. Excel の印刷レイアウト（印刷用/不良要因印刷原紙）を Python でどう再現するか（シート内容は sheets/ 側の担当範囲）。
