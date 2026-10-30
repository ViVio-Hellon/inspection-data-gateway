# G5 特殊検査モード 構造解析レポート（キャスク / 三菱受託 / H494 / 全ヵ所 / 5N）

対象: `【UFKen_12】Cask.bas.bas`(以下 **Cask**), `UFCask.frm.bas`(**UFCask**), `【UFKen_13】Mitsubishi.bas.bas`(**Mit**), `UFMitsubishi.frm.bas`(**UFMit**), シートモジュール `KySh/KySh1`, `ZenSh/ZenSh1/ZenSh2`, `H494_SH/H494_SH1`, `Mitsubishi.cls`, 関連シートCSV。
行番号は各ファイルの行番号（`cat -n` 準拠）。「推測」と明記したもの以外はコードから直接読み取った事実。

---

## 0. 全体像（要点）

| モード | 起動元 | フォーム | 作業シート（原紙→コピー） | 行範囲(板) | 主な出力 |
|---|---|---|---|---|---|
| キャスク | UF1 `KYB_Click`(UF1:3241) → `UFCask.Show` | UFCask | `キャスク原紙` → `キャスク検査` | 7–706 (No.1–700) | A4横印刷、`キャスク保存`+`品証ｺﾋﾟ-用` を xlsx保存（ローカル/サーバ/工場共有）、Access `キャスク梱包板番号` |
| 全ヵ所(共通) | UF1 → `OpenMitubisi("全数")`(UF1:4891) | UFMitsubishi (MLP Page `P1`, WorkFlag=`P1`) | `全ヵ所検査用原紙` → `全ヵ所検査` | 7–76 (No.1–70) ※一部処理は 7–36 のみ | 下記サブモードで分岐 |
| └ 一般 | `Choice`="一般" | 同上 | 同上 | | 全ヵ所検査シート印刷＋不良要因印刷、xlsx保存 |
| └ 5N | `Choice`="5N" | 同上 | 同上 | | 同上＋工場共有`【梱包】５N`へコピー、歪公差固定(幅2.0/丈3.0) |
| └ 三菱受託 | `Choice`="三菱" | 同上 | 同上＋`三菱受託原紙`→`転送済` | 8行(4ブロック×上下)を転記 | `転送済`(三菱様式)印刷、xlsx保存＋NPSサーバ保存 |
| H494 | UF1 → `OpenMitubisi("H494")`(UF1:621) | UFMitsubishi (Page `P_H494`, WorkFlag=`P_H494`) | `H494_原紙` → `H494検査` | 7–16 (No.1–10) | H494検査シート印刷＋不良要因印刷、xlsx保存 |

- 通常検査（UF1 側の検査フォーム）との違い: これら特殊モードは **シート上に直接測定値を入力**（キーボードウェッジでマイクロ/ノギス `DT…` パケット、コンベックス `#L…` パケット）し、**シートモジュールの `Worksheet_Change` が判定を駆動** する。判定記号は判定列に `○`/`×` を書き込み、色付けする。通常検査の帳票(印刷用原紙)・判定ロジック(Judg)とは別系統（ただし公差計算 `Diagonal`/`Lateralbow_J` 等、Ainput 転送 `Ainput_Startup`、日報 `Daily_Report` 等は共通モジュールを呼ぶ）。
- 判定の基本は **公差上下限との包含比較（境界含む＝合格）**。

---

## 1. 各検査モードの業務上の意味（読み取れる範囲）

### 1.1 キャスク検査
- 「キャスク」用製品の全数（最大700枚）の板厚測定。厚 <100mm は 3点（マイクロ）、≥100mm は 4点（ノギス）（`CountSelect`, UFCask:208-209, 459-460; Cask:1963-1993）。
  - 推測: 「キャスク」は使用済燃料輸送容器（キャスク）向けアルミ板。根拠はシート名のみ。
- 巾・丈・対角は「巾丈測定」ボタン(`Inspection`=ON)で任意に追加測定（UFCask:329-359、原紙B5「ﾌｫｰﾑ上のチェックﾎﾞｯｸｽ入で巾丈測定可能」）。
- ラテラルボー・幅歪・丈歪・ゲージ歪・外観の列はあるが **自動判定なし**（原紙K5「自動判定するのは、厚巾丈対角だけです」、AlTerRange は値変換のみ Cask:78-87）。
- 対角公差はフォーム入力必須（`UFCask.Tai`、未入力で判定不可 Cask:1209-1212）。原紙V5の「入力がない場合JIS4級で判定」は現行コードと不一致（4級テーブルはコメントアウト Cask:1136-1199）。
- 測定前の **ゼロ校正必須**（マイクロ/ノギス、コンベックス）。未校正だと判定しない（Cask:1039-1044）。1000mm コンベックス校正値は `品証検証` シートから自動読込（Cask:1939-1955）。
- 梱包単位（スタートNo〜エンドNo）を Access `伝送用ファイル.accdb` の `キャスク梱包板番号` に登録し、連番チェック・エンドNo照合を行う（2026.02.20 追加、UFCask:897-1190, Cask:2241-2530）。
- 保存物は品証用コピー（`品証ｺﾋﾟ-用`）を含み、工場共有 `【梱包】キャスク` へもコピー（Cask:665-677）。

### 1.2 全ヵ所検査（一般 / 5N / 三菱 の共通基盤）
- ロットの全板について、厚4点・幅2点・丈2点・対角(2本)・ラテラルボー・幅歪・丈歪・ゲージ・外観を測定（`全ヵ所検査用原紙` 6行目見出し）。
- サブモードは `UFMitsubishi.Choice`（"三菱"/"5N"/"一般"）と内部フラグ `Work`, `Inflag`, オプション `FiveOP`, `一般` で切替（Mit:1448-1513, UFMit:509-581）。シート `V1`(1,22) にモード名を記録。
- 設備「丸徳」では強制的に「一般」（UFMit:2204-2207）。
- 検入高さ（山数・枚数）の入力（KenHBox1-7, YamaLB）があり、シート `AD1:AD3`(30列) に保存（UFMit:1862-1889, 379-388）。印刷ヘッダに出力。
- 重量 = 厚×幅×丈×2.68/10^6 (kg, 小数3桁切捨て, OutTolerance:813-817) が 35kg 以下なら「ゲージ不要 * 両面定盤歪測定必要」、超なら「ゲージによる測定必要」とメッセージ（UFMit:287-290 等 6箇所）。
- 幅 1〜299 のロットは `H494` フラグON（ロットNoの文字色を赤）→ 幅をノギスで測定（DTパケットを幅列に受け付け、0.00表示）（UFMit:665-669, Cask:143-152）。

### 1.3 5N
- 推測: 5N（純度99.999%）アルミ材の全ヵ所検査。
- 歪判定公差を **固定（幅歪 2.0 / 丈歪 3.0）**（Cask:2017-2024、Mit:1475-1476 で M5/O5 に "2.0"/"3.0" 書込）。
- 保存時、工場共有 `FactoryyPath & "【梱包】５N\" & yyyy.mm`（「５」は全角 U+FF15）へコピー（Mit:1355-1373）。
- 旧来の専用様式 `5N原紙`→`5N` シートへの転記(`FiveN材転送`)は現在呼ばれていない（§12）。現行は全ヵ所検査シート自体を印刷（UFMit:1423-1429「歪公差外流出対策_2024.1.11」）。

### 1.4 三菱受託検査
- 三菱（顧客）向け受託品。`三菱受託原紙`（検査番号/スラブ番号/合金－質別 `MR664-H112`(原紙D4固定)/寸法/原板積載/切断ブロック/梱包番号/上下/厚#1-#4/幅1,2/長さ1,2/対角線差/歪(幅,長さ)/外観、承認・確認欄）へ転記した `転送済` シートを印刷（Mit:740-823）。
- 4ブロック×(上/下)=8行（原紙E10:E17 上/下、全ヵ所検査 AH列 ﾌﾞﾛｯｸ上/下）。
- 三菱原紙の公差表記: 幅±3.0、長さ±4.0、対角線差4.0mm以下、歪 1.0mm以下/製品幅・長さ（原紙J8,L8,N8,O8,P8 固定文字）。
- 保存時、`\\nlmsrvngy03\NPS\全社研修会トレーナー教育\第36回NPSトレーナー教育（名古屋）2017年\梱包実績の見える化\三菱受託検査_履歴\yyyy.mm\` にも保存（Mit:531-590）。
- **2023〜2024 の列追加後、三菱向けの列参照が旧レイアウトのまま残っている疑いが強い**（§4.10, §13）。

### 1.5 H494 検査
- 用途コード H494 向けの専用様式（10枚）。厚4点、幅3点（ノギス、0.00表示）、丈2点、対角、ラテラルボー、定盤、ゲージ、外観（`H494_原紙` 6行目）。
- 原紙B5「対角線差4級内＝○、対角線差3級内=△、対角線差3級外れ＝×」と記載があるが、現行コードは △ を出さない（○/×のみ）。
- ラテラルボー判定は 7〜10行目のみ（Cask:2070）。歪（定盤/ゲージ）は判定なし。

---

## 2. Sub/Function 一覧

### 2.1 `【UFKen_12】Cask.bas.bas`
| 名前(行) | 公開 | 引数 | 役割 | 呼び出し元 → 呼び出し先 |
|---|---|---|---|---|
| MoveTg(2) | Public Function→Range | n, sh | `n="A"`時、キャスク(7-706, B/D/F列)・全ヵ所(7-**36**, B/D/F/H列)で最初の空セルを返す。空き無→Nothing | UFCask(355,382,729,733), UFMit(2186,2191) |
| AlTerRange(43) | Public Function→Range | TempRange, BeforDate | 入力パケット判別(DT/#L)による値変換・振り分け（§4.1） | 各シートChange → Conv, Maker, Special_Move_DT/L |
| キャスクコピー(310) | Public Sub | – | `キャスク検査` 削除→`キャスク原紙` を先頭にコピー・改名、タブ色&H660066, Zoom90 | UFCask.CommandButton1 |
| キャスク用最大最小(336) | Public Sub | – | 指定範囲(MIN〜MAX)の厚/巾/丈 MAX/MIN を V3:X4 に記入・下線 | UFCask.CommandButton3 |
| キャスク専用保存(445) | Public Sub | – | `キャスク保存`+`品証ｺﾋﾟ-用` 作成→新ブック保存(パスワード・読取専用)→サーバ/工場共有にコピー | UFCask CB2/CB4/CB6 → GetCaskData, DisplaySheets, HiddenSheets, CreateFolder, NewFileName, WsCheck |
| JustinCase_KySave(683) | Public Sub | – | 新規発行前の万一保存（キャスク保存＋キャスク検査）→ `HistoryPath\JustinCaseSave` | UFCask CB1/CB11 → AddWS |
| ClearHeaderFooter(790) | Public Sub | ws | ヘッダ/フッタ消去（On Error Resume Next） | キャスク印刷範囲 |
| キャスク印刷範囲(807) | Public Sub | – | 印刷範囲A6:AB最終行、ヘッダ/フッタ、A4横 | UFCask CB2/CB6 → GetCaskData |
| HizumiJ(902) | Public Function→Bool | Ter, kou(String) | 歪 \|値\|≤kou で True | Hizumi_Z |
| LateralJ(912) | Public Function→Bool | Ter, kou(String) | ラテラル \|値\|≤kou で True | Lateralbow_Z |
| KenJudg(923) | Public Function→Bool | Ter, KP(上限), KM(下限) | KM≤値≤KP で True | 検査判定, ループ判定 |
| TaiJudg(932) | Public Function→Bool | Ter, kou(Single) | \|対角1−対角2\|(Single)≤kou | 対角, 対角再チェック |
| ループ判定(979) | Public Sub | – | キャスク検査全行の厚巾丈を再判定し○×書込 | UFCask CB2/CB9/Initialize |
| 検査判定(1028) | Public Sub | Ter | 厚巾丈の判定入口（3シート対応） | 各シートChange, Special_Move_* → Daily_Report, KenJudg, 判定付け |
| 対角(1117) | Public Sub | Ter | 対角判定入口 | 各シートChange, Special_Move_L → TaiJudg, 対角判定付け |
| 対角再チェック(1236) | Public Sub | Ter | 3級公差で△判定（全ヵ所のみ） | **呼出しなし**（コメントアウト Cask:1587） |
| 公差外検索(1335) | Public Sub | Ter | 名称と異なり、特定列入力時に作業中保存するだけ | 各シートChange → 作業毎保存 |
| 判定付け(1365) | Public Sub | Judg, Ter | ○×書込・色・ログ・音・フォーム表示 | 検査判定 → WriteLog, 音楽合格/不合格, UFcount起動, UFOK/UFNG |
| 対角判定付け(1473) | Public Sub | Judg, Ter, kou | 対角線差・○×書込 | 対角 |
| Special_Move_DT(1591) | Public Sub | TarR, BeforDate | DTパケットを次の空き厚セルへ移送し判定 | AlTerRange → 検査判定 |
| Special_Move_L(1715) | Public Sub | TarR, BeforDate | #Lパケットを次の空き巾丈対角セルへ移送し判定 | AlTerRange → 検査判定, 対角 |
| UFcount起動(1882) | Public Sub | Ter | キャスク: UFcount（板No/位置表示・読上げ）表示 | 判定付け系 |
| UF対角起動(1899) | Public Sub | Ter | UF対角（演出用）表示 | 各シートChange(対角1列) |
| 作業毎保存(1919) | Public Sub | – | 作業中保存シート（ｷｬｽｸ作業中保存/全ヵ所作業中保存/H494作業中保存）へ全セルコピー | 公差外検索 → AddWS, Pesute |
| ProofreadingSearch(1939) | Public Sub | ToTime, ToTyo | 品証検証 A5:K98 から当日・当直の1000mm校正値取得 | UFCask CB15/Activate |
| NumericConfi(1956) | Public Function→Bool | n("M"/"C") | 校正値閾値チェック | UFCask.StartTB_AfterUpdate |
| CountSelect(1963) | Public Sub | a(3/4) | 厚測定点数切替（H:I列表示/ロック） | UFCask(208,209,317,459,460) |
| Hizumi_Z(1994) | Public Sub | Ter | 全ヵ所 歪判定（5N固定/一般はM5,O5） | ZenSh Change(25,27列) → HizumiJ, Hizumi_判定付け |
| Lateralbow_Z(2040) | Public Sub | Ter | ラテラルボー判定（全ヵ所W列/H494 Y列7-10行） | ZenSh(23), H494_SH(25) → LateralJ, Lateralbow_判定付け |
| Lateralbow_判定付け(2077) | Public Sub | Judg, Ter | 全ヵ所はX列に○×、H494/キャスクはセル自体を着色 | Lateralbow_Z |
| Hizumi_判定付け(2159) | Public Sub | Judg, Ter | 全ヵ所 歪○×書込 | Hizumi_Z |
| GetCaskData(2251) | Public Function→Variant | strLot, strKen | キャスク梱包板番号を Lot+検査No で取得（スタートNo順） | 専用保存/印刷範囲/AddCaskLabels/CheckCaskEndNo → GetRecordsSQL |
| GetRecordsSQL(2273) | Public Function→Variant | FilePath, TableName, SQLfilter, orderBy | ADODB汎用SELECT→2次元配列、0件/エラーはEmpty | GetCaskData |
| AddCaskLabels(2367) | Public Sub | frm, strLot, strKen, left, top, w, h, gap, prefix | フォーム上に梱包範囲ラベル `No{n}:{start}～{end}` 生成 | UFCask CB3/Initialize |
| CheckCaskEndNo(2468) | Public Function→Bool | strLot, strKen | シートB列最終行のA列No と DB最終エンドNo の一致確認 | UFCask CB2/CB6 |

### 2.2 `UFCask.frm.bas`（すべて Private）
| 名前(行) | 役割 |
|---|---|
| ComboBox1_KeyPress(18), MAXTex_KeyPress(29), LotBox_KeyPress(361), MINText_KeyPress(388), SeriNo_KeyPress(716) | 英数字のみ・小文字→大文字。**MAXTex_KeyPress はコントロール名(MAXText)と不一致で発火しない**（推測: タイプミス） |
| CommandButton1_Click(41) | 新規発行（日報作成→確認→JustinCase_KySave→キャスクコピー→ログ→フォーム項目クリア） |
| CommandButton11_Click(81) | JustinCase_KySave |
| CommandButton12_Click(85) | 読上げ(Read) ON/OFF |
| CommandButton14_Click(94) | アクティブシートA7, A8 を Speak |
| CommandButton15_Click(99) | 1000mm校正値再読込 |
| CommandButton16_Click(103) | AB列AutoFit（保護解除したまま） |
| CommandButton2_Click(110) | 印刷（ループ判定→必須チェック→エンドNo照合→印刷範囲→印刷→専用保存→保存用に履歴） |
| CommandButton25_Click(176) | フォームサイズ復元 |
| CommandButton3_Click(182) | 転送（入力チェック→連番チェック→点数切替→DB保存→ラベル→最大最小→Ainput_Startup("Cask","L2")） |
| CommandButton4_Click(234) | 保存のみ（キャスク専用保存） |
| CommandButton5_Click(245) | 取消（アクティブ列に応じ直前入力をクリア） |
| CommandButton6_Click(308) | CountSelect(3)強制→エンドNo照合→印刷範囲→専用保存（試験用の疑い §12） |
| CommandButton8_Click(324) | EnableEvents/ScreenUpdating 復帰 |
| CommandButton9_Click(329) | 巾丈測定 ON/OFF（K7:AB706 ロック/灰色切替＋ループ判定） |
| LotBox_Change(374) | 7桁で ComboBox1 へ転記（→DB読込） |
| ComboBox1_Change(401) | SIKALOTNOW.accdb `仕掛` からロット情報→シート上部（寸法・公差）設定 |
| Label4_MouseMove(712) | 空 |
| SB_SpinDown/Up(727,731) | スクロール5行＋MoveTg |
| StartTB_AfterUpdate(736) | ゼロ校正入力（DT=マイクロ/ノギス, #L=コンベックス） |
| UserForm_Activate(777) | シート表示、校正値読込 |
| UserForm_Click(811) | Loadtext="on" |
| UserForm_Initialize(817) | シートからフォーム項目復元、ループ判定、ラベル生成 |
| UserForm_QueryClose(875) | キャスク系シート非表示・Main表示 |
| ValidateCaskInput(898) | Lot/検査No/スタートNo/エンドNo の未入力チェック |
| CheckCaskSequence(926) | 梱包範囲の連番チェック（DB） |
| SaveCaskToDB(1056) | キャスク梱包板番号 UPDATE/INSERT、スタートNo=1 で他レコード削除(DAO) |

### 2.3 `【UFKen_13】Mitsubishi.bas.bas`
| 名前(行) | 公開 | 引数 | 役割 | 呼び出し元 → 先 |
|---|---|---|---|---|
| OpenMitubisi(3) | Public Sub | Temp("全数"/"H494") | UFMitsubishi 起動前設定（ページ切替・シートから復元・検査員名合成） | UF1(621, 4891) → RemoveDuplicates, WsCheck |
| RemoveDuplicates(77) | (Public) Function | tMP, n | 区切り文字(既定半角空白)で語の重複削除 | OpenMitubisi |
| 全ヵ所用コピー(95) | Public Sub | Temp="全ヵ所検査" | 原紙コピーで作業シート新規作成（全ヵ所/H494）タブ色16764108, Zoom110 | UFMit CB1/ws_out/Activate |
| 全ヵ所用最大最少(134) | Public Sub | Temp | MAX/MIN 記入・下線（全ヵ所/H494） | UFMit CB3/CB19 |
| JustinCase_ZnSave(324) | Public Sub | Temp | 新規発行前の万一保存 | UFMit CB1/CB6/ws_out → AddWS, CreateFolder |
| 全ヵ所専用保存(427) | Public Sub | Temp | 保存シート作成→新ブック保存→(三菱時NPS保存)→サーバコピー | UFMit CB2/CB4/CB5/CB7/CB23 |
| 全ヵ所用転送(619) | Public Sub | – | `ｺﾋﾟ-用` 作成、機器文字列、不良要因シート、(三菱時)`転送済` 作成・転記・歪外観チェック | UFMit CB2/CB4/CB7/CB15 → 全ヵ所用機器転送, AddWS, Pesute, 不良要因印刷範囲 |
| 全ヵ所用印刷範囲(860) | Public Sub | Temp | 全ヵ所/H494 シート印刷設定・ヘッダ/フッタ | UFMit CB7/CB15/CB23 |
| 原紙印刷範囲(939) | Public Sub | – | `5N` or `転送済` 印刷設定 | UFMit CB2 |
| 全ヵ所用機器転送(986) | Public Sub | ws, x, y, ATU, HABA, Yout | 板厚・ライン別の測定機器名文字列をセルに記入（§4.9） | 全ヵ所用転送, FiveN材転送, UFMit CB23 |
| FiveN材転送(1058) | Public Sub | – | 5N原紙様式への転記 | **呼出しなし**（UFMit:1292,1402 はコメント） |
| FiveN専用保存(1232) | Public Sub | – | 5N用保存（列幅固定、工場共有【梱包】５N コピー） | UFMit CB4/CB10/CB15 |
| Judg_PorM(1379) | Public Function | ATU, KM, KP | 公差外の±方向 | **呼出しなし**（Cask:1458 コメント内のみ） |
| CheckRedXCells(1387) | (Public) Function | rangeA | 背景RGB(255,0,0)かつ"×"のセル数 | UFMit CB7/CB15/CB18 |
| ExtractNthText(1398) | (Public) Function | InputString, a | a番目の()内の文字列の"÷"より前 | ForN_Select |
| ExtractBefore(1426) | (Public) Function | Temp, n="÷" | n より前（無ければ""） | ExtractNthText, ForN_Select |
| CountParenthesesSets(1435) | (Public) Function | InputString | ()セット数 | **呼出しなし** |
| ForN_Select(1448) | Public Sub | – | Choice に応じフラグ/表示/シートV1設定、5N歪公差、検入高さ復元 | UFMit(273,512,2180,2218,2221) → DrawListBoxContent, ExtractNthText |

### 2.4 `UFMitsubishi.frm.bas`（すべて Private。主要なもの）
| 名前(行) | 役割 |
|---|---|
| Cle_Click(59) | 不良要因印刷_{Line} の C3:X7, C9:X9, C10 クリア |
| CommandButton18_Click(73) | 検入高さ必須・公差外(赤×)件数警告のみ（後続処理なし） |
| CommandButton19_Click(99) | H494 転送: 検査法(op1-3)必須→最大最少("H494検査")→Ainput_Startup("H494", 検査法) |
| CommandButton21/16_Click(126/1499) | EnableEvents=True, Excel表示 |
| CommandButton22_Click(131) | H494検査 AC列AutoFit |
| CommandButton23_Click(138) | H494 印刷・保存（機器転送H494→印刷範囲→印刷→全ヵ所専用保存("H494検査")→不良要因印刷→一時シート削除→保存用履歴"H494"） |
| CommandButton24_Click(232) | UF品質基準 表示（品証基準・写真・歪公差反映） |
| CommandButton25_Click(265) | サイズ復元 |
| FiveOP_Click(271) / 一般_Click(2178) | Choice="5N"/"一般" → ForN_Select |
| Frame2/Frame4/Frame6_Click, シート発行_Click, UserForm_Click, MLP_Click | 重量によるゲージ要否メッセージ（MLP_Click は検入高さ保存も） |
| H494_Lot_KeyPress/Change(310,321), LotBox_KeyPress/Change(1811,1822), KenNo_KeyPress(1799), ComboBox1_KeyPress(583) | 英数大文字化、7桁で ComboBox1 へ |
| KenHBox1-6_AfterUpdate(331-378) | 値を"0.0"整形、ログ、山数描画 |
| KenHBox7_AfterUpdate(379) | AD3 に「{KenHBox7}枚」。**但し整形・ログ対象が KenHBox6（バグ疑い）** |
| ListReflect_Click(390) | 不良要因の枚数反映（50枚超は拒否、ｽﾎﾟｯﾄ処理は除外） |
| NgIm_Click(445), ViewCB_Click(1939) | 不良要因シートの画像プレビュー表示/復帰 |
| S/U/C/D/EBad_Change, _MouseDown | 数値以外クリア、整形、色 |
| Choice_Change(509) | "５N"→"5N"、三菱/一般/5N 以外は警告、ForN_Select |
| CMB_Click(517) | Choice に応じフラグ設定（ForN_Select の簡易版） |
| ComboBox1_Change(596) | SIKALOTNOW.accdb `仕掛` からロット情報、公差、対角/ラテラル公差、Lot情報履歴、QP反映、包装仕様書チェック |
| CommandButton1_Click(1250) | 全ヵ所 新規発行 |
| CommandButton10_Click(1289) | FiveN専用保存 |
| CommandButton14_DblClick(1298) | モード循環 一般→三菱→5N→一般 |
| CommandButton15_Click(1363) | 5N 印刷・保存 |
| CommandButton17_Click(1504) | Y,AA,AD列 フォント8・AutoFit |
| CommandButton2_Click(1515) | 三菱 印刷・保存 |
| CommandButton3_Click(1605) | 全ヵ所 転送: 検査法(Inspection1-3)必須→最大最少→Ainput_Startup("Mitsubishi", 検査法) |
| CommandButton4_Click(1632) | 転送＋保存（印刷なし） |
| CommandButton5_Click(1643) | 全ヵ所専用保存のみ |
| CommandButton6_Click(1647) | JustinCase_ZnSave |
| CommandButton7_Click(1654) | 一般 印刷・保存 |
| HindBo_Click, MLP_MouseMove, UserForm_MouseDown | 表示サイズ調整 |
| ws_out_Click(2114) | H494 新規発行 |
| YamaLB_Click(2153) | 山数描画 |
| SB_SpinDown/Up(2183,2188) | 3行スクロール＋MoveTg("A","全ヵ所検査")（H494でも全ヵ所検査を対象） |
| UserForm_Activate(2194) | モード選択（InputBox）、シート表示、H494検査が無ければ作成 |
| UserForm_Initialize(2294) | レジストリから設備名/モード、YamaLB, 不良要因リスト |
| UserForm_QueryClose(2344) | 全ヵ所/H494系シート非表示、Main表示 |

### 2.5 シートモジュール
| モジュール | イベント | 内容 |
|---|---|---|
| KySh / KySh1 (同一内容) | Worksheet_Change(9), SelectionChange(93) | `キャスク検査` 用（§3.1） |
| ZenSh / ZenSh1 (同一内容) | Change(9), SelectionChange(127) | `全ヵ所検査` 用（§3.2） |
| ZenSh2 | Change(9), SelectionChange(131) | 旧版（三菱簡易歪判定1.0 等）。残骸の疑い（§12） |
| H494_SH / H494_SH1 (同一内容) | Change(10), SelectionChange(98) | `H494検査` 用（§3.3） |
| Mitsubishi.cls | なし（8行の空） | 推測: `三菱受託原紙` のコードネーム |

**どれが実シートに紐づくか（推測）**: 作業シートは毎回 `原紙.Copy` で作られる（Cask:323, Mit:110,113）。Excel はシートコピー時、元のコードネーム（例 `KySh`）に連番を付けた新コードネーム（`KySh1`）とモジュール内容の複製を作る。よって
- `KySh` = `キャスク原紙`、`KySh1` = `キャスク検査`（コピー）
- `ZenSh` = `全ヵ所検査用原紙`、`ZenSh1` = `全ヵ所検査`
- `H494_SH` = `H494_原紙`、`H494_SH1` = `H494検査`
- 内容が完全一致（diff で VB_Name 行のみ相違）であることがこの推測を支持。原紙は非表示で直接編集されないため、実際に発火するのは `*1` 側。
- `ZenSh2`: VB_Base が `{FCFB3D2A-…}`（ワークシート `{00020820-…}` ではない）、`VB_Customizable=False`、内容が 2023.12.20 / 2024.1.11 改修前の旧版 → シートに紐づかない残骸（イベント不発）と推測。
- 移植上は **`*1` と原紙側は同一ロジックとして1本化** してよい（どちらも同一コード）。

---

## 3. 各モードの画面項目・操作フロー

### 3.1 キャスク（UFCask + キャスク検査）

**フォーム項目（コードから参照されるもの）**
- 入力: `LotBox`(7桁英数), `ComboBox1`(ロット), `KenNo`(検査No), `Kname`(検査員), `Cno`(コンベックスNo), `Mno`(マイクロNo), `SeriNo`(S/N), `MINText`(スタートNo), `MAXText`(エンドNo), `Tai`(対角公差), `StartTB`(校正入力)
- 表示/状態: `Micro`(ゼロ校正値, Caption), `Convex`(コンベックス校正値), `Convex_1000`(1000mm校正値), `L1`(ﾏｲｸﾛ/ﾉｷﾞｽｾﾞﾛ点), `L2`,`L3`(表示フラグ), `CountSt`("3"/"4"), `Inspection`("ON"/"OFF" 巾丈測定), `Read`("ON"/"OFF" 読上げ), `LineName`, `Loadtext`("on"/"off"), `Cnt`, `UndoCount`, `Hflag`, `Tag`("off"で処理抑止), 動的ラベル `lblCask0..n`

**フロー**
1. UF1 で検査員登録済み→「キャスク」ボタン（UF1 `KYB_Click`）→ `UFCask.Show`、LineName設定。
2. Initialize: キャスク検査シートがあれば各項目をシートから復元（D2→LotBox, Y2→KenNo, Y3→Kname, Z2→Cno, Z3→Mno, Z4→SeriNo, AC2→Micro, AC3→Convex, AE3→Convex_1000, AG3→CountSt）、ループ判定、Inspection=OFF なら K6:AB706 灰色・K7:AB706 ロック、梱包ラベル生成（UFCask:833-868）。Activate: 1000mm校正読込。
3. 新規発行（CB1）: 「新規発行しますか？現在のｼｰﾄは削除されます。」→ JustinCase保存 → 原紙コピー → ログ「ｷｬｽｸ検査ｼｰﾄ発行」→ Inspection=OFF, LotBox/KenNo/Kname/ComboBox1/Convex/Micro をクリア（Cno/Mno/SeriNo/Convex_1000/Tai/MIN/MAX は保持）。
4. ロット入力（7桁）→ DB `仕掛` 検索 → D2=ロット, KenNo, B3=厚("0.000"), 点数切替, D3=巾("0.0"), F3=丈("0.0"), 公差 M/O 列、上下限式 Q/S 列。見つからない→上部クリア、「Text.データなし　手入力よろしく」、Q/S式のみ再設定。
5. ゼロ校正（StartTB）: DTパケット→厚≥100 は「ﾉｷﾞｽｾﾞﾛ点」("0.00")、未満「ﾏｲｸﾛｾﾞﾛ点」("0.000")。値≠0 なら「ノギス(マイクロ)校正を正しく実行してください！！」で消去。#L→コンベックス("0.0")、\|値\|>0.5 で「コンベックス校正を正しく実行してください！！」消去。`#L` 単独入力で L2 表示。
6. 測定: シートの厚セルにDTパケット入力→変換→判定（○/×）→次セルへ。巾丈測定ONなら #L パケットで巾・丈・対角。対角2入力で対角判定（対角公差必須）。
7. 転送（CB3）: スタートNo/エンドNo入力→連番チェック→DB保存→ラベル→MAX/MIN→Ainput 起動（検査法 "L2"）。
8. 印刷（CB2）: 必須項目（Cno, Mno, Kname, SeriNo）→エンドNo照合（不一致なら中断）→印刷→保存→`保存用` に履歴行。
9. 終了（×）: キャスク検査・ｷｬｽｸ作業中保存・キャスク保存 を非表示、Main 表示。

**シートイベント（KySh1）**
- Change（単一セル時のみ）:
  - `UFCask.Cnt="0"` なら Undo で変更前値 `Befor` を取得し再設定（KySh:19-27）。
  - 行7–706: 列2–8→右隣へ移動、列9（CountSt=4時）→ Inspection OFF:次行B / ON:同行K、列11–18→右隣、列19（対角1）→音・UF対角・T〜W列クリア・右へ、列20→右、列21→X列、列24–27→右（KySh:30-71）。※`Case 7` は `Case 2 To 8` に先取りされ到達不能。
  - その後: AlTerRange → 数値なら 検査判定 → 対角 → 公差外検索（KySh:73-84）。
- SelectionChange: Cnt="0"。列21選択→X列へ。列24選択→Y2:Z4, AB2:AE3, AG3 にフォーム値を書込（ヘッダ情報の保存タイミング）（KySh:104-126）。D7 選択時に B7 が入っていれば **B1 に開始日時**（毎回上書き）・Daily_Report(1),(3)（KySh:131-143）。入力セル薄青・判定セル以外白、Inspection OFF時 K6:AB706 灰色。

### 3.2 全ヵ所（UFMitsubishi P1 + 全ヵ所検査）
**フォーム項目**: `LotBox`, `ComboBox1`, `KenNo`, `Kname`, `YoutoNo`, `Choice`(三菱/5N/一般), `FiveOP`, `一般`(オプション), `Work`("ON"/"OFF"), `Inflag`("flagon"/""), `Normal`/`Mitsubishi`/`FiNi`(表示切替), `H494`(幅<300フラグ), `MINText`, `MAXText`, `Inspection1-3`(検査法 OptionButton), `KenHBox1-7`, `YamaLB`(1–6, 枚数), `高さ計算式`(Caption), `DefectFactorLB`, `SBad/UBad/CBad/DBad/EBad`(不良枚数), `NgIm`, `MLP`(MultiPage), `WorkFlag`(Caption "P1"/"P_H494"), `LineName`, `Cnt`, `UndoCount`, `PackNo_on_off`, `Tag`

**フロー**
1. UF1 →「全数」→ OpenMitubisi("全数"): P1 ページ表示、WorkFlag="P1"、検査員 = UF1 当日登録者（`WorkeDayMax` に応じ T1/T2/T3 を "_" 分割した最後の要素）＋シートH3 の既存名を空白連結して重複削除、シートから Lot/検査No/Choice(V1)/用途(H4)/高さ式(AD1) 復元、D2 文字色赤なら H494=True。
2. Activate: 丸徳→"一般"強制。Choice 空なら InputBox「作業名入力：三菱、5N、一般　のいずれかを入力」（キャンセル/空→`End`=全VBA停止）。ForN_Select。
3. 新規発行（CB1）: JustinCase_ZnSave → 全ヵ所用コピー → ログ「三菱受託検査ｼｰﾄ発行」→ 不良要因印刷_{Line} 削除（**DisplayAlerts未抑止で確認ダイアログが出る**）→ Work="ON", LotBox/KenNo/ComboBox1/YoutoNo クリア。
4. ロット入力 → ComboBox1_Change（§6.2）→ 寸法・公差・対角公差 S1・ラテラル公差 M1・(QP_DB による歪公差 M5/O5 等)。
5. 測定（シート入力、§3.2.1）。
6. 検入高さ（KenHBox, YamaLB）。MLP クリックで AD1=高さ計算式, AD2="{n}山"（最大番号の数値 KenHBox）, AD3="{n}枚"。
7. 転送（CB3）: 検査法選択必須 → 最大最少 → Ainput_Startup("Mitsubishi", 検査法)。
8. 印刷: 一般=CB7、5N=CB15（Choice="5N"でなければ何もせず終了）、三菱=CB2。
   - CB7/CB15 共通前処理: 検入高さ(AD2/AD3)いずれも空→「検入高さを入力してください。」中止。厚(B7:I76)・歪(Y7:AB76)の赤×件数を警告し続行確認（UFMit:1661-1683）。
9. 終了: 全ヵ所/H494系シート非表示、Main 表示。

#### 3.2.1 シートイベント（ZenSh1）
- Change（単一セル）: Undo で Befor 取得。行7–76:
  - 列2–8→右、列9→K列が空ならK、そうでなければ次行B、列11–18→右
  - 列19（対角1）: 値あり→音、UF対角起動（※名前誤記で表示されない）、**T〜W列(20–23)をクリア**（W=ラテラルボー列も消える）、右へ
  - 列20（対角2）: AlTerRange → 対角 → W列へ移動、終了
  - 列21,22→W列へ
  - 列23（ラテラルボー）: 数値かつ "-" 以外かつ非空 → `UnitRecord` で0.5ピッチ丸め("0.0") → Lateralbow_Z、終了
  - 列24→右、列25（幅歪）/27（丈歪）: 数値なら Hizumi_Z で終了（非数値・空は下の共通処理へ）、列26,28→右
  - 共通: AlTerRange → 数値なら 列2,4,6,8,11,13,15,17,19,20 のみ 検査判定 → 公差外検索 → シート保護（ZenSh:103-123）
- SelectionChange: 列2選択→H2=KenNo, H3=Kname, H4=YoutoNo（フォント5）。列22,24,26,28選択→右へ。D7選択で B7 が入っていれば B1=開始日時、Daily_Report(1),(3)。色付け。

### 3.3 H494（UFMitsubishi P_H494 + H494検査）
**フォーム項目**: `H494_Lot`, `H494_KenNo`, `H494_Kname`, `H494_Youto`, `H494_MINText`, `H494_MAXText`, `op1-3`(検査法), 他は全ヵ所と共用（ComboBox1 等）。
**フロー**: UF1 →「H494」→ OpenMitubisi("H494")（P_H494表示、シートから H494_Lot/KenNo/Youto/高さ式 復元、**検査員はUFMitsubishi.Kname に設定され H494_Kname ではない**）→ Activate で H494検査 が無ければ作成 → ロット入力（H494_Lot 7桁 → ComboBox1）→ 測定 → 転送 CB19 → 印刷 CB23 → 新規発行 ws_out。
**シートイベント（H494_SH1）**: 行7–16: 列2–8→右、列9→K空ならK/他は次行B、列11–20→右、列21（対角1）→音・UF対角・V〜X(22–24)クリア・右、列22（対角2）→AlTerRange・対角・次行Bへ、列25（ラテラル）→数値かつ"-"以外なら UnitRecord（**空でも0扱いで"0.0"を書く**）→Lateralbow_Z→右。共通: AlTerRange → 数値かつ非空なら 検査判定 → 対角 → 公差外検索。SelectionChange: 列24→Y列へ、列26選択→H2=H494_KenNo, H3=H494_Kname、D7で開始日時。

---

## 4. 判定ロジック・計算式・公差・閾値（原文・行番号付き）

### 4.1 入力パケット変換（AlTerRange, Conv, Maker）
`Maker`（補助:393-407）: 空なら""、長さ≥3 で先頭2文字を `StrConv(UCase(Left(Trm,2)), vbNarrow)` → "DT" / "#L" / "#〆"。
`Conv`（補助:380-391）:
```vb
Case "DT" 'DT10002+0000020.712M
 Conv = Mid(Trm, 9, 11)
Case "#L" '＃L395.8
 Conv = Mid(Trm, 3, Len(Trm))
Case "#〆"
 Conv = Mid(Trm, 3, Len(Trm))
Case Else
 Conv = Trm
```
- **DT の符号（8文字目）は捨てられる**（"DT10002-0000000.012M" → "0000000.012" = +0.012）。
- AlTerRange（Cask:43-308）の振り分け（tempStr=Maker(値)）:

| シート | 範囲 | DT | #L | その他(手入力等) |
|---|---|---|---|---|
| キャスク | 行7–706 列2,4,6,8(厚) | Conv で数値化 | Special_Move_L | **"" に消去** |
| キャスク | 列11,13,15,17,19,20(巾丈対角) | Special_Move_DT | Conv | **"" に消去** |
| キャスク | 列24–28(ﾗﾃ/歪/外観) | Special_Move_DT | Special_Move_L | Conv（手入力可） |
| キャスク | その他 | Undo→最初の空き厚セル(7行〜, CountSt 4→列8まで/3→列6まで)をActivate | Undo→最初の空き巾丈セル(列11–17)をActivate | Conv |
| 全ヵ所 | 行7–76 列2,4,6,8 | Conv | Special_Move_L | "" |
| 全ヵ所 | 列11,13 かつ H494=True | Conv＋表示形式 `0.00_);[赤](0.00)` | **Special_Move_L（手入力も含む）** | (同左) |
| 全ヵ所 | 列11,13,15,17,19,20 | Special_Move_DT | Conv＋表示形式 `0.0_);[赤](0.0)` | 何もしない（手入力保持） |
| 全ヵ所 | 列24,25,26 | Special_Move_DT | Special_Move_L | Conv |
| 全ヵ所 | 列23 | Special_Move_DT | Special_Move_L | 何もしない |
| 全ヵ所 | その他 | Undo→空き厚(列2–8)へ | Undo→空き巾丈(列11–17)へ | Conv |
| H494 | 行7–16 列2,4,6,8 | Conv | Special_Move_L | "" |
| H494 | 列11,13,15 | Conv＋`0.00` 形式 | Special_Move_L（手入力含む） | |
| H494 | 列17,19,21,22 | Special_Move_DT | Conv＋`0.0` 形式 | 何もしない |
| H494 | 列26,27 | Special_Move_DT | Special_Move_L | Conv |
| H494 | 列25 | Special_Move_DT | Special_Move_L | 何もしない |
| H494 | その他 | Undo→空き厚 | Undo→空き(列11–19) | Conv |

- Special_Move_DT（Cask:1591-1714）: **現在行 TR から下へ** 各行の厚列（キャスク: 列2..n step2, n=CountSt"4"→8/"3"→6、それ以外→ループなし; 全ヵ所: 列2–8 step2、H494フラグ時は続けて同行の列11,13; H494: 列2,4,6,8,11,13,15）で最初の空セルに Conv 値を書き、元セルを Befor に戻し、そのセルで 検査判定。空き無→「厚欄はすべて埋まっています」。
- Special_Move_L（Cask:1715-1881）: キャスクは Inspection="ON" 必須（OFF→「巾丈測定がONになっていません」）。行TRから下へ列11,13,15,17（→検査判定）・19,20（→対角）。全ヵ所: 列11,13(H494=Falseのみ),15,17→検査判定、19,20→対角。H494: 列17,19→検査判定、21,22→対角。空き無→「巾丈対角欄はすべて埋まっています」。
- 移送後 `UndoCount="1"`（Befor が空なら"0"）にし、元セルの再判定を1回スキップ（判定付け Cask:1398-1409）。

### 4.2 厚・巾・丈 判定（KenJudg）
```vb
Public Function KenJudg(Ter As Range, KP As Double, KM As Double) As Boolean   'Cask:923
If IsNumeric(Ter.Value) = True Then
  If ((Ter.Value) < KM) Or ((Ter.Value) > KP) Then
    KenJudg = False '×
  ElseIf ((Ter.Value) >= KM) And ((Ter.Value) <= KP) Then
    KenJudg = True '○
```
- 上下限（Cask:1072-1077）: `ATU1=val(Q2)`(下限), `ATU2=val(S2)`(上限), `haba1=val(Q3)`, `haba2=val(S3)`, `take1=val(Q4)`, `take2=val(S4)`。呼出し `KenJudg(Ter, ATU2, ATU1)`（上限, 下限）。
- シート数式（全3シート共通、UFCask:473-496, UFMit:683-706）:
  - Q2 `=B3-M2`（厚下限）、S2 `=B3+O2`（厚上限）
  - Q3 `=D3-M3`、S3 `=D3+O3`（巾）
  - Q4 `=F3-M4`、S4 `=F3+O4`（丈）
  - M列=マイナス公差、O列=プラス公差（値は正数で格納）
- 対象列（検査判定 Cask:1078-1112）:
  - キャスク(行7–706): 厚 列2,4,6,8 / 巾 11,13 / 丈 15,17
  - 全ヵ所(行7–76): 厚 2,4,6,8 / 巾 11,13 / 丈 15,17
  - H494(行7–16): 厚 2,4,6,8 / 巾 **11,13,15** / 丈 **17,19**
- 前提条件: キャスクでは `UFCask.Micro`・`Convex` が空なら「ゼロ校正を実行してください！！」で中止（Cask:1039-1044）。`If Not IsNumeric(.Cells(3,2)) And IsNumeric(.Cells(3,4)) And IsNumeric(.Cells(3,6)) Then Exit Sub`（Cask:1053、演算子優先で「B3非数値かつD3,F3数値」の時のみ中止）。値が文字列・空なら判定しない（Cask:1069）。
- 境界: 下限・上限ちょうどは ○。`val()` は数値→文字列(15桁)→数値の往復となる点に注意（移植では15有効桁丸めで再現推奨）。

### 4.3 公差の取得規則（DB）
UFCask:466-496 / UFMit:676-706（同一ロジック）:
```vb
If dOP = 0 Or dOM = 0 Then dOP = dMP: dOM = dMM
.Cells(2, 15) = Format(dOP, "0.000")   ' 厚+ (O2)
.Cells(2, 13) = Format(dOM, "0.000")   ' 厚- (M2)
```
- オーダー公差（±）の **どちらか一方でも 0** なら、両方とも製造公差を採用。巾・丈も同様（書式 "0.0"）。
- 厚 B3 = `Format(val(製造板厚), "0.000")`、巾 D3/丈 F3 = "0.0"。

### 4.4 判定付け（○×の出力）Cask:1365-1472
- `ws.Cells(TR, TC+1) = "○"` / `"×"`、色 ColorIndex 10(緑)/3(赤=RGB(255,0,0))。
- 次入力セル着色（"苦肉策"）: TC∈{2,4,6,11,13,15,17,20}→TC+2、TC=8→TC+3 を RGB(153,204,255)。
- ログ（WriteLog 2回）: `"*UFKen:" & ws.Name & "*入力:" & 値 & " *座標:" & "(" & TR & "," & TC & ")" & "：○"` と、末尾に `":" & Lname` を付けて `{Lpath}\Measured.log`。
- Lpath: L-1→`HistoryPath & "L-1\L-1_Log"`, LVC→`LVC\LVC_Log`, HVC→`HVC\HVC_Log`, ﾊﾞﾗﾝｻｰ→`バランサー\バランサー_Log`, ﾄｯﾄ→`トット\トット_Log`, Test→`Test\Test_Log`, それ以外→`丸徳\丸徳_Log`。**Lname は全ヵ所/キャスク可視時のみ設定され、H494検査可視時は "" → 丸徳ログに出る**（Cask:1373-1376）。
- 音（音楽合格/不合格）、キャスクは UFcount 表示、それ以外は UFOK/UFNG（モードレス）。
- ループ判定（Cask:979-1027）は色が RGB(0,128,0)/RGB(200,60,50) と異なる（ColorIndex 3 ではない）。

### 4.5 対角
- TaiJudg（Cask:932-977）: `Tai = Abs(val(S) - val(T))` を **Single** に格納、`kou >= Tai` で True。対象: キャスク 列19/20(行7–706)、全ヵ所 列19/20(7–76)、H494 列21/22(7–16)。両セル数値かつ非空でなければ False。
- 公差 kou（対角 Cask:1117-1235）:
  - キャスク: 列20のみ。`UFCask.Tai=""`→「対角公差が入力されていません！！」中止。`Inspection="ON"` かつ数値なら `kou=val(UFCask.Tai)`（それ以外は kou=0）。
  - 全ヵ所/H494（UFCask非表示時）: `kou = val(ZenSh.Cells(1,19))`（S1）。ZenSh は WorkFlag P1→全ヵ所検査、P_H494→H494検査。全ヵ所は列20、H494は列22で判定。
- S1 の設定（UFMit:720-733）: 品質ｸﾞﾚｰﾄﾞ_直角度(R1) "4"→`Format(Diagonal(幅,丈),"0.0")`、"3"→`ReDiagonal`、`指示_直角度` が空でなければ `Format(ConvertToHalf(指示),"0.0")` で上書きし R1="追記"。Diagonal/ReDiagonal は Judg モジュールの公差テーブル（LookupDiag、別担当）。
- 対角判定付け（Cask:1473-1590）:
  - キャスク: V列(22)= `Abs(Format(val(S)-val(T), "0.0"))`、U列(21)=○(RGB153,204,255)/×(RGB200,60,50)、×時 W列(23)=`"公差:[" & Format(UFCask.Tai,"0.0") & "]"`（Tai空なら `"4級公差:[" & Format(kou,"0.0") & "]"`）フォント8、○時 W列クリア。
  - 全ヵ所: U列(21)=差、V列(22)=○(RGB0,128,0)/×(RGB200,60,50)。
  - H494: X列(24)=差（`val(V)-val(U)`）、W列(23)=○(RGB153,204,255)/×。
- 注意: 判定は丸め前の差（Single）、表示は0.1丸め。表示"2.0"でも×があり得る。
- 対角再チェック（3級テーブル、△）: Cask:1248-1305（幅1–300/300.1–600/…/2100.1–2400 × 丈1–3700/3701–6000 → 1.5/2.0, 2.0/3.0, 3.0/4.0, 4.0/5.0, 5.0/6.0, 5/6, 5/6, 5/6）。**現在未使用**。

### 4.6 ラテラルボー
- UnitRecord（Judg:2421-2444）: 絶対値を 0→0, 0.1–0.5→0.5, 0.51–1.0→1.0, 1.01–1.5→1.5, …, 4.51–5.0→5.0, 5.01–999→そのまま、符号復元。範囲外（例 0<n<0.1, 0.5<n<0.51 等）は 0。引数は Single。
- 全ヵ所 列23: 非数値はそのまま、"-" は保持、空はスキップ、それ以外 `Format(UnitRecord(val(W)),"0.0")`（ZenSh:63-75）。H494 列25: 空でも UnitRecord(0) → "0.0"（H494_SH:62-73）。
- Lateralbow_Z（Cask:2040-2076）: `kou = val(.Cells(1,13))`（M1、String変数）。全ヵ所 行7–76 列23、H494 **行7–10** 列25。
- LateralJ（Cask:912-921）: kou が "0" または "" → **False（×）**。`val(kou) >= Abs(val(値))` で True。
- M1 の設定（UFMit:735-756）: `指示_ﾗﾃﾗﾙﾎﾞｰ` 非空→ConvertToHalf 値、L1="追記"。空→ 品質ｸﾞﾚｰﾄﾞ_ﾗﾃﾗﾙﾎﾞｰ(L1) "4"→Lateralbow_J(幅,丈,厚)、"3"→ReLateralbow、その他→Lateralbow_J、結果0なら""。さらに QP_DB(1426-1431) が M1 を上書きする場合あり（別担当）。
- 出力: 全ヵ所は X列(24)に○/×（ColorIndex 10/3）、H494・キャスクはセル自体を着色（○ColorIndex10, ×RGB(200,60,50)）。

### 4.7 歪（全ヵ所のみ）
- Hizumi_Z（Cask:1994-2039）:
```vb
If UFMitsubishi.FiveOP Then
  Case (TR >= 7 And TR <= 76) And (TC = 25): Call Hizumi_判定付け(HizumiJ(Ter, 2#), Ter) '5N 歪公差固定_2024.1.11
  Case (TR >= 7 And TR <= 76) And (TC = 27): Call Hizumi_判定付け(HizumiJ(Ter, 3#), Ter)
ElseIf UFMitsubishi.一般 Then
  SetH = sh.Cells(5, 13).Value : SetT = sh.Cells(5, 15).Value   ' M5 / O5
```
- HizumiJ（Cask:902-911）: kou "0"/"" → False(×)、`val(kou) >= Abs(val(値))` → ○。
- 出力: Z列(26)/AB列(28) に○/×（ColorIndex 10/3）。
- **三菱モード（FiveOP・一般 ともFalseの場合）は歪判定なし**。旧版 ZenSh2:76-103 には三菱時 歪 ≤1.0 で OK/NG 表示（○×書込なし）があった。
- M5/O5 は QP_DB(2051-2052) が設定（別担当）、5N選択時は "2.0"/"3.0" を書込（Mit:1475-1476）。

### 4.8 MAX/MIN（印刷帳票上の最大最小）
- キャスク（Cask:336-443）: 行 n=MINText+6 〜 k=MAXText+6。厚: CountSt"4"→B,D,F,H列、"3"→B,D,F列の Max/Min。巾: K,M、丈: O,Q。V3/V4, W3/W4, X3/X4 に記入、一致セルに下線（各行で最初の一致のみ）。
  - **バグ疑い（Cask:395-396）**: 厚MINの走査が `For i = x To k` / `For a = 2 To n Step 2`（x=列数6/8、n=開始行）と行・列変数が入れ替わっている。結果、MIN値と一致するセルが開始行より前の行や対象外の偶数列で探される／CountSt=4・MIN=1 の場合は行7・列8が探索されず **V4 が空のまま** になり得る。
- 全ヵ所/H494（Mit:134-322）: 厚 B,D,F,H 4列。MINatu=0 の場合の補正:
```vb
For i = 1 To 4
  Select Case i
  Case 1: AtuVar(1) = .Min(C2) ...
  Case 2: AtuVar(2) = .Min(C4) ...
  Case 6: ...   ' ← i=3,4 は一致しない
  Case 8: ...
MINatu = WorksheetFunction.Small((AtuVar), 2)
If MINatu = 0 Then MINatu = WorksheetFunction.Small((AtuVar), 3)
```
  - **Case 6/8 は到達不能**のため AtuVar(3),(4) は常に0 → 実質「列B と列D の最小値が両方非0ならその小さい方、そうでなければ0」。
  - 巾: P1→K,M / H494→K,M,O、丈: P1→O,Q / H494→Q,S。
  - 書込先: P1 W3/W4(厚), Y3/Y4(巾), AA3/AA4(丈)、H494 V,W,X 列。下線リセット範囲 B7:V{36 or 16}（全ヵ所は36行目まで）。

### 4.9 測定機器文字列（全ヵ所用機器転送 Mit:986-1056）
| ライン | line | Conb |
|---|---|---|
| L-1 | "L-1" | " C11" |
| HVC | "HVC" | " C11" |
| LVC | "LVC" | " C10" |
| ﾊﾞﾗﾝｻｰ | "ﾊﾞﾗﾝｻｰ" | " C11" |
| ﾄｯﾄ | "ﾄｯﾄ" | " C12" |
| その他(Test/丸徳) | "" | "" |

| 厚 ATU | 文字列（"　　"は全角空白2つ U+3000） |
|---|---|
| 0<ATU≤28 | `"ﾏｲｸﾛ:" & line & "M1" & "　　ｺﾝﾍﾞｯｸｽ:" & line & Conb` |
| 28<ATU≤54 | 同 "M2" |
| 54<ATU≤74 | 同 "M3" |
| 74<ATU<300 | `"ﾉｷﾞｽ:" & line & "N5" & "　　ｺﾝﾍﾞｯｸｽ:" & line & Conb` |
| ATU≥300 | 同 "N6"。ATU=300 ちょうどの時、ws の行14–18×列2,4,6,8 の数値を調べ <300 と ≥300 の両方→`"ﾉｷﾞｽ：" & line & "N5,N6" & "　　ｺﾝﾍﾞｯｸｽ:" & line & "C11"`、<300のみ→`"ﾉｷﾞｽ" & line & "N5" & …"C11"`、≥300のみ→`"ﾉｷﾞｽ" & line & "N6" & …"C11"`（コロン有無・全角コロン・C11固定が混在） |
| ATU≤0/非数値 | 書込なし |
| Yout="H494" かつ val(HABA) 1–299 | 上記を上書き `"ﾏｲｸﾛ:" & line & "M1" & "　　幅:ﾉｷﾞｽ:" & "N5" & "_丈:ｺﾝﾍﾞｯｸｽ:" & line & Conb` |
- 書込先: 全ヵ所検査 AC3(3,29)（全ヵ所用転送、Yout=H4用途コード）、転送済 G2(2,7)、H494検査 Y2(2,25)（CB23, Yout="H494"）、5N F2/全ヵ所Y2（FiveN材転送・未使用）。
- **H494 印刷フッタは AC3 を参照（Mit:922）するが機器文字列は Y2 に書かれる → 測定機器が空欄で印刷される疑い**。

### 4.10 三菱 転記・チェック（全ヵ所用転送 Mit:619-859）
- `ｺﾋﾟ-用`: 全ヵ所検査 A6:AD{最終行} を D1 から貼付け、列 F,H,J,L,M,O,Q,S,U:X,AA,AC,AE を削除 → 結果 D=No, E–H=厚1–4, I=幅1, J=幅2, K=丈1, L=丈2, **M=対角判定(V列)**, N=ラテラル, O=幅歪, P=丈歪, Q=ゲージ, R=外観。A1:B6 にラベル/値（LotNo, Size "厚×巾×丈", 測定機器, 検査員, 作業日 "B1～Now"）。
- `転送済`（三菱時のみ）: 原紙コピー＋図形、D2=KenNo, E5/G5/I5=厚/巾/丈, F8/J8/L8=`"' -{M}mm " & Chr(10) & "+{O}mm"`, Q19=作業日, Q20=検査員, G2=機器、B10,12,14,16 ← ｺﾋﾟ-用 Q2,Q4,Q6,Q8、C10.. ← S列（空）、D10,12,14,16 = 1,2,3,4（厚1に値がある場合）、F10:I17 ← ｺﾋﾟ-用 E2:H9（厚）、J10:Q17 ← ｺﾋﾟ-用 I2:P9。
  - **列ずれ疑い**: 三菱原紙は J,K=幅, L,M=長さ, N=対角線差, O,P=歪(幅,長さ), Q=外観 だが、実際は N←対角判定記号, O←ラテラル, P←幅歪, Q←丈歪 が入る。B列(原板積載)←ゲージ、C列(切断ブロック)←空。元板NO(AE)・切断ﾌﾞﾛｯｸNO(AG)はそもそもコピー範囲(A:AD)外。
  - 「数字変換」(Mit:661-672) は ｺﾋﾟ-用 S列(19) の 1–4 を ①–④ にするが S列は削除後空。
- 三菱時の入力チェック（Work="ON"）: 行7〜n で列24,25 が空→「歪入力がありません,はじめからやり直してください *[座標](i,a)」、列26 が空→「外観入力がありません…」→ ｺﾋﾟ-用/転送済/全ヵ所検査保存 を削除し `End`。**現レイアウトでは 24=X(ﾗﾃ判定), 25=Y(幅歪), 26=Z(幅歪判定)** であり、歪/外観の意図とずれている疑い。
- CB2 のブロックNOチェック（UFMit:1532-1537）: 行7,9,11,13 で B列値があり列27 または29 が空→「ﾌﾞﾛｯｸNOの入力をしてください」。**現レイアウトでは 27=AA(丈歪), 29=AC(ゲージ)**（元板NO=AE(31), 切断ﾌﾞﾛｯｸNO=AG(33)）。

### 4.11 その他の閾値
- 校正: マイクロ/ノギス ゼロ点は `val(Micro) <> 0` で NG（完全一致0のみ可）、コンベックスは `val(Convex) > 0.5 Or < -0.5` で NG（±0.5 含む OK）（Cask:1956-1962）。
- 点数切替: 厚 `100# > B3` → 3点、`100# <= B3` → 4点（UFCask:208-209）。ゼロ校正表示も `ATU >= 100` でノギス。
- 不良枚数: 各 >50 で拒否（ｽﾎﾟｯﾄ処理は除外）（UFMit:411-415）。
- H494 判定: 製造板幅 `Case 1# To 299#`（UFMit:665-669、Mit:1049-1051）。
- 重量 35kg 閾値（UFMit 各所）。
- 品質基準フォーム: FlagAtu ≥4 → Page1、<4 → Page2、>6 で DBdisplay外観_H（UFMit:243-262）。

---

## 5. 原紙シートのレイアウト

### 5.1 キャスク原紙 / キャスク検査（行7–706, A列=No 1–700）
| セル | 内容 |
|---|---|
| B1 | 開始日時（D7選択時）, C1/D1 は保存時「～」「保存日時」 |
| B2 "LotNo", D2 | ロットNo |
| B3 / D3 / F3（C3,E3 "×"） | 厚("0.000") / 巾 / 丈 |
| K2:K4 | 厚公差/巾公差/丈公差、L列"－" M列=マイナス値, N列"+" O列=プラス値, Q列=下限式, R列"～", S列=上限式 |
| U3 "MAX", U4 "MIN", V2:X2 厚/巾/丈 | V3:X4 = MAX/MIN 値 |
| Y2/Y3 | 検査No / 検査員（SelectionChange列24で記録） |
| Z2/Z3/Z4 | コンベックスNo / マイクロNo / S/N |
| AB2/AB3/AD3 | 「ｾﾞﾛ校正:ﾏｲｸﾛ or ﾉｷﾞｽ」「ｾﾞﾛ校正:ｺﾝﾍﾞｯｸｽ」「1000mm校正:ｺﾝﾍﾞｯｸｽ」 |
| AC2/AC3/AE3 | 各校正値 |
| AG3 | CountSt（3/4） |
| 6行目見出し | B厚1 C判定 D厚2 E判定 F厚3 G判定 H厚4 I判定 K幅1 L判定 M幅2 N判定 O丈1 P判定 Q丈2 R判定 S対角1 T対角2 U判定 V対角線差 (W=公差注記) Xﾗﾃﾗﾙﾎﾞｰ Y幅歪 Z丈歪 AAｹﾞｰｼﾞ歪 AB外観 |
| B4,B5,K5,V5,Z5 | 注意書き文言 |
- 印刷: A6:AB{B列最終行}、A4横、幅1ページ、余白 左右0/上0.8/下0.35/ヘッダ0.25/フッタ0.2 inch、水平中央（Cask:819-897）。
  - 左ヘッダ: `&"Meiryo UI"&8 *{D2}* 検査員：{Kname} ⏎ {厚 0.000}×{巾 0.0}×{丈 0.0} ⏎ *ｾﾞﾛ校正_{ﾉｷﾞｽ|ﾏｲｸﾛ}:{AC2 0.000} *ｾﾞﾛ校正_ｺﾝﾍﾞｯｸｽ:{AC3 0.0} *1000mm校正_ｺﾝﾍﾞｯｸｽ:{AE3 0.0}`
  - 右ヘッダ: `*厚公差：{x-Am 0.000}～{x+Ap 0.000} ⏎ *巾公差：{0.0}～{0.0} ⏎ *丈公差：{0.0}～{0.0}`
  - 左フッタ: `&"Meiryo UI"&10` + `梱包No{n}_{start}～{end}` を2空白区切り
  - 右フッタ: `*{B1 yyyy/mm/dd h:mm}～{Now} 印刷`
- 保存版 `キャスク保存`: 全セルコピー、下線・塗り解除、Y1"検査番号"/Z1, Y2"検査員", Y3"ｺﾝﾍﾞｯｸｽ", Y4"ﾏｲｸﾛ", Z4 `Mno & "  S/N" & SeriNo`, AB2:AE3 校正、C1"～", D1 保存日時、列幅 B,K,Z=8.5, V=18、保護（Cask:479-522）。
- `品証ｺﾋﾟ-用`: キャスク保存 A6:AB{n} を D1 へ、列 F,H,J,L:M,O,Q,S,U,V:X,Z 削除 → D=板番号(No), E–H=厚1–4, I,J=幅, K,L=丈, M=対角線差, N=ﾗﾃﾗﾙ, O=幅歪, P=丈歪, Q=ｹﾞｰｼﾞ歪, R=外観。A1:B10: LotNo / 検査No / Size / ｺﾝﾍﾞｯｸｽNo / ﾏｲｸﾛNo(`Mno & " S/N" & SeriNo`) / 検査員 / 作業日(`B1～D1` "yyyy/mm/dd  h:mm") / ｾﾞﾛ校正×2 / 1000mm校正。B12〜（12+i行, i=0起点）: `梱包No{i+1}:{start}～{end}`。Meiryo UI 10、罫線なし（Cask:534-597）。

### 5.2 全ヵ所検査用原紙 / 全ヵ所検査（行7–76, A列 No 1–70）
| セル | 内容 |
|---|---|
| B1 | 開始日時 |
| K1"ﾗﾃﾗﾙﾎﾞｰ等級", L1, M1 | 等級(4/3/追記), ラテラル公差 |
| O1 | 納入先＆材質（QP_DB） |
| Q1"直角度等級", R1, S1 | 等級(4/3/追記), 対角公差 |
| V1 | 作業選択名（三菱/5N/一般） |
| D2, H2, H3, H4 | LotNo, 検査No, 検査員, 用途コード |
| B3/D3/F3, K2:S4 | 寸法・公差（キャスクと同じ配置） |
| K5"歪公差", L5"幅", M5, N5"丈", O5 | 歪公差 |
| V3"MAX"/V4"MIN", W2厚 Y2巾 AA2丈 | W3:W4, Y3:Y4, AA3:AA4 |
| AC3 | 測定機器文字列 |
| AD1/AD2/AD3 | 高さ計算式 / "n山" / "n枚" |
| 6行目 | B厚1 C判定 D厚2 E判定 F厚3 G判定 H厚4 I判定 K幅1 L判定 M幅2 N判定 O丈1 P判定 Q丈2 R判定 S対角1 T対角2 U対角線差 V判定 Wﾗﾃﾗﾙﾎﾞｰ X判定 Y幅歪 Z判定 AA丈歪 AB判定 ACｹﾞｰｼﾞ AD外観 AE元板NO AG切断ﾌﾞﾛｯｸNO AH三菱 AI5N |
| AF奇数行 "-", AH 奇数行"ﾌﾞﾛｯｸ上"/偶数行"ﾌﾞﾛｯｸ下" | 原紙固定 |
- 印刷（Mit:860-938）: P1 A6:AD{n}、H494 A6:AB{n}、B7:AB76 塗りを白、A4横、1×1ページ、余白 左右0.5/上0.8/下0.5/ヘッダ・フッタ0.2、水平垂直中央。
  - 右ヘッダ: `*厚公差：…～… ⏎ *巾公差 ⏎ *丈公差 ⏎ *歪公差：幅:{M5 0.0}_丈:{O5 0.0} ⏎ *対角公差：{S1 0.0} ⏎ *ﾗﾃﾗﾙﾎﾞｰ公差：{M1 0.0}`
  - 左ヘッダ: `*{D2} ⏎ *{厚0.000}×{巾0.0}×{丈0.0} ⏎ *{AD1}_{AD2}`
  - 左フッタ: `*検査員：{Kname or H494_Kname} *測定機器：{AC3}`
  - 右フッタ: `*{B1 yyyy/mm/dd  h:mm}～{Now} 印刷`
- 保存版 `全ヵ所検査保存`: C1"～", D1保存日時, F1 `"検査員：" & Kname`(フォント7)。全ヵ所専用保存は AutoFit＋B,K,X(H494はB,K,Y)=7.7、FiveN専用保存は全30列8.75、列3,5,7,9,10,12,14,16,18,21,24,26,28=2.5。

### 5.3 H494_原紙 / H494検査（行7–16）
- 1行目 K1"ﾗﾃﾗﾙﾎﾞｰ等級"(L1,M1), Q1"対角公差"(R1,S1)。2–4行は全ヵ所と同配置、MAX/MIN は U3/U4 見出し・V,W,X 列。H2/H3=検査No/検査員、Y2=測定機器。
- 6行目: B厚1 C判定 D厚2 E判定 F厚3 G判定 H厚4 I判定 K幅1 L判定 M幅2 N判定 O幅3 P判定 Q丈1 R判定 S丈2 T判定 U対角1 V対角2 W判定 X対角線差 Yﾗﾃﾗﾙﾎﾞｰ Z定盤 AAｹﾞｰｼﾞ AB外観。

### 5.4 三菱受託原紙 → 転送済（印刷範囲 B2:R20、A4横 1×1、原紙印刷範囲 Mit:949-983 ／ 全ヵ所専用保存内では A3 設定 Mit:534-558）
- B2"検査番号" C2"NLM" F2"検査機器" Q2"承認" R2"確認" / B3"スラブ番号" C3"顧客" / B4"合金－質別" C4"顧客" D4"MR664-H112" / B5"寸法" C5"NLM" D5"厚(mm)" F5"幅(mm)" H5"長さ(mm)" / 7行: B原板積載 C切断ブロック D梱包番号 F厚さ J幅 L長さ N対角線差 O歪 Q外観 / 8行: E"公差" J"-3.0 +3.0" L"-4.0 +4.0" N"4.0mm以下" O"1.0mm以下/製品幅" P"1.0mm以下/製品長さ" / 9行: E"原板位置" F–I #1–#4, J,K 1,2, L,M 1,2, O幅 P長さ / E10–E17 上下交互 / P19"開始終了" P20"検査員"。
- 左ヘッダ: `&"Meiryo UI"&7{LineName} ⏎ *&"Meiryo UI"&7*{全ヵ所AD1}_{AD2}`。

### 5.5 5N原紙（未使用様式）
A2検査番号 E2測定器 K2承認 L2確認 / A3スラブ番号 / A4寸法 B4厚さ D4幅 F4長さ / A6板№ C6厚さ I6外観 K6備考 / C7公差 / A8測定位置 C8–F8 #1–#4 / I15開始終了 I16検査員。

---

## 6. 外部依存

### 6.1 パス（初期設定.bas:18-35, `Setting_Path`、レジストリ `UF検査判定システム\起動設定\モード`）
| 変数 | オンライン | オフライン |
|---|---|---|
| MainPath | `\\nlmsrvngy03\Read\【New】仕掛\台帳\` | ThisWorkbook.Path & "\" |
| QualityPath | `\\nlmsrvngy03\工場内共有\検査データ\品証-梱包やり取りフォルダ\Accdb\` | 同上 |
| HistoryPath | `\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\検査履歴\` | 同上 |
| FactoryyPath | `\\nlmsrvngy03\工場内共有\検査データ\` | 同上 |

### 6.2 Access
- `MainPath & "SIKALOTNOW.accdb"`（ACE.OLEDB.12.0, adOpenStatic/adLockReadOnly）
  - UFCask:426 / UFMit:629（同文）: `SELECT * FROM 仕掛 WHERE ﾛｯﾄ番号 = '{ロット}' AND ｵｰﾀﾞｰ板丈 <> '0'`（先頭1件のみ使用）
  - 使用フィールド（UFMit 表記）: ｵｰﾀﾞｰ番号, ﾛｯﾄ番号, 製造板厚, 製造板幅, 製造板丈, 検査番号, 製造材質, 製造調質, 用途ｺｰﾄﾞ, 設計_設備ｺｰｽ, 取引先名, 納入先名, 送り先名, BOX実績_枚本数, BOX実績_良品重量, 梱包_合紙, VC, 品質ｸﾞﾚｰﾄﾞ_ﾗﾃﾗﾙﾎﾞｰ, 品質ｸﾞﾚｰﾄﾞ_直角度, 指示_直角度, 指示_ﾗﾃﾗﾙﾎﾞｰ, 板厚/板幅/板丈公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ/ﾏｲﾅｽ, 板厚/板幅/板丈公差_製造_ﾌﾟﾗｽ/ﾏｲﾅｽ
  - **UFCask は一部フィールドを全角カナで参照**（"オーダー番号","ロット番号","用途コード","設計_設備コース"、UFCask:436-445）。UFMit は半角。どちらかがフィールド不一致でエラーになる可能性（§13）。
  - 取得結果は `HistorySave(PastHiki)`（Lot情報履歴、別担当）。
- `QualityPath & "伝送用ファイル.accdb"` テーブル `キャスク梱包板番号`（列: 管理番号, LotNo, 検査No, スタートNo, エンドNo, 作業時間）
  - GetRecordsSQL: `SELECT * FROM [キャスク梱包板番号] WHERE [LotNo]='{lot}' AND [検査No]='{ken}' ORDER BY スタートNo`（Cask:2258-2260, 2293-2299）。配列列位置 3=スタートNo, 4=エンドNo を前提（SELECT * の列順依存）。
  - CheckCaskSequence: `SELECT スタートNo, エンドNo FROM キャスク梱包板番号 WHERE LotNo='{lot}' AND 検査No='{ken}' AND スタートNo <> {start} ORDER BY スタートNo`（UFCask:976-980）
  - SaveCaskToDB STEP1: `SELECT * FROM キャスク梱包板番号 WHERE LotNo='{lot}' AND 検査No='{ken}' AND スタートNo={start}`（adOpenKeyset=1, adLockOptimistic=3）→ 有: エンドNo, 作業時間 更新 / 無: `SELECT MAX(管理番号) AS MaxNo FROM キャスク梱包板番号` で MAX+1（Null→1）採番して AddNew（UFCask:1087-1127）
  - STEP2（スタートNo=1 の時）: DAO.DBEngine.120 で `SELECT * FROM キャスク梱包板番号 WHERE LotNo='{lot}' AND 検査No='{ken}' AND スタートNo <> 1` を開き全件 Delete（UFCask:1136-1166）。トランザクションなし。
  - 作業時間 = CB3 クリック時の Now。
  - SQL は文字列連結（エスケープなし）。
- 他（別担当モジュール経由）: GetLotInfo("Lot情報履歴"), DbReturnItem, DistArray, FormReflectDist, DBdisplay外観_H, PackingNoCheck/UFPack_Open。

### 6.3 Excel ブック出力（いずれも新規ブック xlsx、`Password:=Lot & "." & Format(Date,"yyyy.mm.dd")`、保存後 `SetAttr vbReadOnly`）
| 処理 | 含むシート | 保存先 | 追加コピー |
|---|---|---|---|
| キャスク専用保存 | キャスク保存, 品証ｺﾋﾟ-用 | `C:\検査履歴\yyyy.mm\yyyymmdd_{Lot}\yyyymmdd_{Lot}.xlsx`（重複時 `.002`〜`.999`、999超→「連番をこれ以上付けることが出来ません。」→`End`） | `HistoryPath & "キャスク_履歴\" & yyyy.mm`（Testも同じ）、`FactoryyPath & "【梱包】キャスク\" & yyyy.mm`（Test: `HistoryPath & "Test\Test(工場共有)_履歴\キャスク_履歴"`）。コピー対象 = SetPath 内の最新ファイル(NewFileName) |
| JustinCase_KySave | キャスク保存, キャスク検査 | `HistoryPath & "JustinCaseSave\yyyymmdd_{Lot}\…"`（Test: `HistoryPath & "Test\Test_履歴\JustinCaseSave"`） | なし |
| JustinCase_ZnSave(Temp) | {Temp}保存, 全ヵ所検査（可視のもののみ） | 同上 JustinCaseSave | なし |
| 全ヵ所専用保存(Temp) | {Temp}保存, ｺﾋﾟ-用, 転送済, 5N, 不良要因印刷_{Line}（存在するもの） | `C:\検査履歴\…`（同形式） | 三菱(Inflag="flagon"): `\\nlmsrvngy03\NPS\…\三菱受託検査_履歴\yyyy.mm\yyyymmdd_{Lot}.xlsx`（UF1.Test時 `C:\Users\NLM Administrator\Desktop\保存用ﾌｫﾙﾀﾞ\三菱受託検査_履歴\…`）を **同じBKで再SaveAs**。サーバ: `HistoryPath & "{Temp}_履歴\" & yyyy.mm`（Test: `HistoryPath & "Test\ " & …` ※"\ "空白混入、丸徳等: ""） |
| FiveN専用保存 | 全ヵ所検査保存, ｺﾋﾟ-用, 転送済, 5N, 不良要因印刷_{Line} | `C:\検査履歴\…` | `HistoryPath & "全ヵ所検査_履歴\" & yyyy.mm`（Test: `Test\Test_履歴\全ヵ所検査_履歴\`）、5N条件(Inflag="" かつ Work="OFF" かつ Choice="5N")で `FactoryyPath & "【梱包】５N\" & yyyy.mm`（Test: `HistoryPath & "Test\5N_履歴\"`） |
- 作業中保存（ブック内シート）: `ｷｬｽｸ作業中保存`(タブ色10079487), `全ヵ所作業中保存`(6723891), `H494作業中保存`(&HCCFFCC)。
- `保存用` シートへの履歴行: A=`"{B1 yyyy/mm/dd  h:mm}～{Now}"`, B=Lot, C=厚, D=幅, E=丈, BY(77)="別ファイル参照", CE(83)=種別（"ｷｬｽｸ"/"三菱受託"/"全ヵ所"(一般・5N共通)/"H494"）。**キャスクは C列未記入・D列に厚→巾で上書き（UFCask:156-157 バグ疑い）**。

### 6.4 その他
- 印刷: `PrintOut Copies:=1`（キャスク検査 / 全ヵ所検査 / 転送済 / H494検査 / 不良要因印刷_{Line}）。
- 音: `MusicTrue("印刷.wav"/"完了.wav"/"おまちください.wav")`, 音楽合格/不合格/対角, `Range.Speak`。
- レジストリ: `GetSetting("UF検査判定システム","起動設定","設備名"/"モード")`。
- ログ: WriteLog（既定先＋`{Line}_Log\Measured.log`）、DebugLog（Cask/UFCask の DB処理）。
- Ainput_Startup("Cask","L2") / ("Mitsubishi", 検査法) / ("H494", 検査法) → 実績転送フォーム（別担当）。
- 日報: Register_Check→Daily_Create（新規発行時）、Daily_Report(1/2/3)。

---

## 7. ハードコードされた業務ルール（全列挙）

1. 厚判定・巾判定・丈判定: 下限≤値≤上限で ○（境界含む）。上限=基準+プラス公差、下限=基準−マイナス公差（シート数式）。
2. オーダー公差の±どちらかが0なら製造公差を両方採用（UFCask:470/481/492, UFMit:680/691/702）。
3. 厚書式 "0.000"、巾・丈 "0.0"、公差 厚"0.000"/巾丈"0.0"。
4. キャスク: 厚<100 → 3点、≥100 → 4点（ノギス）。
5. キャスク: ゼロ校正（マイクロ/ノギス=0ちょうど、コンベックス ±0.5以内）完了前は判定しない。
6. キャスク: 対角公差はフォーム入力必須。巾丈対角測定は Inspection=ON の時のみ。
7. キャスク: 自動判定は厚・巾・丈・対角のみ。ラテラル・歪・外観は記録のみ。
8. キャスク: 梱包範囲は連番必須（前エンドNo+1=次スタートNo）、スタートNo≤エンドNo、スタートNo=1 なら既存の他範囲を全削除、印刷時シート最終板No=DB最終エンドNo 必須。
9. 全ヵ所/H494: 対角公差 = 直角度等級4→4級テーブル, 3→3級テーブル, 指示_直角度があれば最優先。
10. 全ヵ所/H494: ラテラル公差 = 指示_ﾗﾃﾗﾙﾎﾞｰ優先、等級3→3級、それ以外(4含む)→4級、0なら空（→全て×）。
11. ラテラルボー測定値は 0.5mm ピッチ切上げ（5.0超はそのまま）、"-" は「測定不要」として保持。
12. 5N: 歪公差 幅2.0 / 丈3.0 固定。一般: M5/O5（品証DB由来）。三菱: 歪判定なし（現行）。
13. 幅 1〜299mm → H494 扱い（幅をノギス測定、0.00表示、ロットNo赤字）。用途コード "H494" かつ幅1–299 → 機器文字列「幅:ﾉｷﾞｽ:N5」。
14. 測定機器: 厚 ≤28 M1 / ≤54 M2 / ≤74 M3 / <300 ノギスN5 / ≥300 ノギスN6、ライン別コンベックス C10(LVC)/C11(L-1,HVC,ﾊﾞﾗﾝｻｰ)/C12(ﾄｯﾄ)。
15. 重量(厚×幅×丈×2.68/10^6, 小数3桁切捨て) ≤35kg → ゲージ不要・両面定盤歪測定、>35kg → ゲージ測定。
16. 設備「丸徳」は全ヵ所で「一般」固定。
17. 作業中保存タイミング: キャスク 列2,6(3点時)/4,8(4点時)/13,17,20,26、全ヵ所 列4,8,13,17,20,23、H494 列4,8,15,19,22,28 の入力時。
18. 開始日時: D7 選択時に B7 が入力済みなら B1=現在時刻（毎回上書き）。
19. 保存ファイル名 `yyyymmdd_{Lot}.xlsx`、重複時 `.002`〜`.999`、パスワード `{Lot}.{yyyy.mm.dd}`（保存日）、読取専用属性。
20. 月フォルダ `yyyy.mm`（保存日）。
21. 不良要因 枚数は1要因50枚以下（ｽﾎﾟｯﾄ処理除く）。
22. 印刷前警告: 厚列B:I・歪列Y:AB の「赤背景×」件数（一般/5N）。検入高さ未入力は印刷不可。
23. 三菱: 行7,9,11,13（測定ありの行）でブロックNO（現コードでは列27,29）必須、各行で列24,25,26 必須（§4.10 ずれ疑い）。
24. 全ヵ所 新規発行後 `Work="ON"`。
25. ロットNo・検査No等の入力は英数字のみ（小文字→大文字）、ロットNoは7桁で自動検索。
26. 1000mm校正値は `品証検証` A5:K98 の K列先頭10文字=日付(yyyy.mm.dd, 夜勤0時以降は前日)・B列=直 一致の C列（最後の一致行）、検査員未入力なら J列を採用。

---

## 8. 入出力項目

### 8.1 入力
- 測定入力（シート、キーボードウェッジ）: DT パケット `DT` + 5文字 + 符号 + 11文字数値 + `M`（例 `DT10002+0000150.712M`）、`#L{数値}`（コンベックス）、`#〆{数値}`（手入力扱い）、手入力数値（列により消去/保持 §4.1）。
- フォーム入力: §3 の各項目。
- DB: 仕掛（ロット情報）、キャスク梱包板番号、Lot情報履歴、品証DB(QP)。
- シート: 品証検証（1000mm校正）、不良要因印刷原紙（不良要因リスト C2:X2/C9:X9）、UF1 の検査員登録（T1/T2/T3, WorkeDayMax）。

### 8.2 出力
- 作業シートへの書込: 判定記号・色・差分・MAX/MIN・下線・ヘッダ情報（§5）。
- 印刷物（キャスク検査 / 全ヵ所検査 / 転送済 / H494検査 / 不良要因）。
- xlsx ファイル（§6.3）、Access 更新（キャスク梱包板番号）、`保存用` シート履歴行、ログファイル（Measured.log 他）、日報（Daily_Report）、Ainput 経由の実績転送。

---

## 9. エラー処理

| 箇所 | 方式 | 挙動 |
|---|---|---|
| 各シート Worksheet_Change | `On Error Resume Next`（KySh:18, ZenSh:18, H494_SH:19） | 判定処理中の実行時エラーは握り潰され次行へ。末尾で Err.Clear |
| ZenSh/H494 SelectionChange | `On Error Resume Next`（色付け部） | |
| GetRecordsSQL | `On Error GoTo ErrHandler` | MsgBox「GetRecordsSQL エラー：」+Description、Empty を返す |
| AddCaskLabels | ErrHandler | MsgBox「ラベル追加エラー：」 |
| CheckCaskEndNo | ErrHandler | MsgBox「チェックエラー：」→False（A列が数値でない等） |
| CheckCaskSequence | ErrHandler | MsgBox「連番チェックエラー：」→False（MIN/MAX 非数値の CLng 失敗含む） |
| SaveCaskToDB | ErrHandler | MsgBox「DB保存エラー：」→False（STEP1成功・STEP2失敗でも部分更新が残る） |
| UFCask/UFMit ComboBox1_Change | ErrHandler | MsgBox「DB読み込みエラー:」、イベント再有効化。UFMit は DB ファイル無で「DBファイルが見つかりません:」 |
| ClearHeaderFooter | On Error Resume Next | |
| ViewCB_Click | 部分的 Resume Next、リトライ（コピー3回・画像取得8回） | 失敗時 MsgBox |
| MLP_Click | On Error Resume Next | |
| UFCask.UserForm_Activate | シート無→「「キャスク検査」シートが見つかりません。処理を中断します。」Unload |
| 保存連番超過 | MsgBox→`End` | 全VBA停止（フォーム状態も消える） |
| 三菱 歪/外観未入力 | MsgBox→一時シート削除→`End` | 同上 |
| UFMit Activate InputBox キャンセル/空 | `End` | |
| その他多くの処理 | エラー処理なし | 実行時エラーダイアログ（例: MoveTg が Nothing の時 `.Activate`、全ヵ所専用保存の SrvPath="" 時 FileCopy、三菱時の NewFileName フォルダ不在 §13） |
- 業務メッセージ一覧（主なもの）: 「新規発行しますか？現在のｼｰﾄは削除されます。」「ｷｬﾝｾﾙ処理」「フォーム上で必要欄を記入してください」「ゼロ校正を実行してください！！」「対角公差が入力されていません！！」「厚欄はすべて埋まっています」「巾丈対角欄はすべて埋まっています」「巾丈測定がONになっていません」「Text.データなし　手入力よろしく」「LotNo が未入力です。」「検査No が未入力です。」「スタートNo が未入力です。」「エンドNo が未入力です。」「スタートNo がエンドNo より大きい値です。…」「番号が連番になっていない箇所があります。」(タイトル"連番チェック")「「キャスク検査」シートにデータがありません。」「DBにデータが見つかりません。LotNo：…　検査No：…」「エンドNoが一致しません。…」「キャスク検査を印刷しますか？」「印刷しますか？」「検査員名と検査番号を入力してください」「ﾌﾞﾛｯｸNOの入力をしてください」「検入高さを入力してください。」「公差外判定の板があります 厚:n箇所　歪:n箇所 処理を続行しますか？」「検査法を選択してください」「正しい作業選択名を入力してください」「5N専用」「保存ｼｰﾄがありません」「{Temp}がありません」「不良要因が選択されていません」「枚数入力がありません」「すみません、50枚以上は分けて入力してください」「ノギス/マイクロ/コンベックス校正を正しく実行してください！！」。

---

## 10. 状態保持・日付変更時の挙動

- **永続状態はブック内の作業シート**（キャスク検査/全ヵ所検査/H494検査）。フォームを閉じても作業シートは残り（非表示化）、再オープン時にシートから復元（UFCask Initialize, OpenMitubisi）。
- ヘッダ情報（検査No/検査員/機器No/校正値/CountSt）は **特定列のセル選択時** にシートへ書かれる（キャスク: 列24選択, 全ヵ所: 列2選択, H494: 列26選択）。それ以前にフォームを閉じると校正値等は失われ再校正が必要。
- フォーム上の一時フラグ: `Cnt`(Undo可否), `UndoCount`(移送後再判定抑止), `Tag`("off"で各処理抑止), `Inspection`, `CountSt`, `Work`, `Inflag`, `Choice`, `H494`, `FiveOP`/`一般`, `PackNo_on_off`, `Loadtext`。
  - UFCask CB2 でエンドNo不一致終了時、`Tag="off"`・ScreenUpdating=False・UFProgress 表示のまま終了（UFCask:130-135）→ CB1/CB3 が以後 Exit（検査判定で Tag が "" に戻るまで）。
  - CB3 で入力チェック失敗時 `EnableEvents=False` のまま終了（UFCask:184-201）。LotBox_Change/ComboBox1_Change も早期 Exit で EnableEvents=False のまま（UFCask:377-379, 404-406, 532; UFMit:598-602, 801, 809）→ シートの自動判定が止まる。復帰手段: CB8/CB16/CB21 等。
- 作業中保存: 特定列入力ごとにブック内の作業中保存シートへ全コピー（ブック自体の保存は別、推測: ブック保存は他モジュール）。
- 新規発行時に JustinCase 保存（サーバ）→シート再作成。
- **日付依存**: 保存フォルダ月・ファイル名日付・パスワードは「保存実行日」の Date（検査開始日ではない）。日付をまたいだ作業では検査開始日と異なるフォルダ/パスワードになる。開始日時 B1 は D7 選択のたびに上書き。
- 1000mm校正読込は YeaDayrCheck/TimeCheck（夜勤 0:00〜End3 は前日扱い、L-1/LVC/HVC/Test 以外は常に当日・"日勤"）。
- 新規発行時 Register_Check→Daily_Create（日報シート作成）、測定時 Daily_Report(2)、開始時 (1),(3)。

---

## 11. Python 移植時のテストケース候補（具体値）

### 11.1 パケット変換
| 入力 | Maker | Conv | 備考 |
|---|---|---|---|
| `DT10002+0000150.712M` | DT | `0000150.712` → 150.712 | |
| `DT10002-0000000.012M` | DT | `0000000.012` → 0.012 | 符号消失を再現 |
| `dt10002+0000010.512M` | DT | 10.512 | 大文字化 |
| `ＤＴ10002+0000010.512M` | DT | 10.512 | 全角→半角判定（Mid位置は元文字列） |
| `#L395.8` | #L | 395.8 | |
| `#〆12.3` | #〆 | 12.3 | |
| `DT` / `12` | "" | そのまま | 長さ<3 |
| `1528.2`（手入力） | "" | 1528.2 | キャスク厚列なら消去、全ヵ所巾列なら保持 |

### 11.2 厚判定（B3=36.000, M2=0.4, O2=0.4 → Q2=35.6, S2=36.4）
35.600→○, 35.599→×, 36.400→○, 36.401→×, 36.000→○。
（B3=10.000, M2=O2=0.2 → 10.200→○, 10.201→×, 9.800→○, 9.799→×）

### 11.3 公差フォールバック
オーダー(+0.3/−0.3)・製造(+0.2/−0.2) → 0.3/0.3。オーダー(+0/−0.3) → 製造 0.2/0.2。オーダー(+0.3/−0) → 製造。

### 11.4 点数切替
厚 99.999 → 3点（H:I列幅0.05・ロック）、100.000 → 4点。

### 11.5 校正
厚36: DT `…+0000000.000M` → Micro "0.000" OK。`…+0000000.001M` → NG（消去）。厚120: "0.00" 形式・ラベル「ﾉｷﾞｽｾﾞﾛ点」。#L0.5 OK、#L-0.5 OK、#L0.6 NG、#L-0.51 NG。

### 11.6 対角（Single 精度の再現が必須）
| S | T | 公差 | VBA(Single) | 素朴な倍精度 |
|---|---|---|---|---|
| 1530.4 | 1528.3 | 2.1 | **○** | × |
| 1000.3 | 1002.4 | 2.1 | **○** | × |
| 2000.0 | 2002.1 | 2.0 | × | × |
| 1528.3 | 1526.3 | 2.0 | ○ | ○ |
- キャスク: 2000.0/2002.1, Tai=2.0 → U=×, V=2.1, W="公差:[2.0]"。Tai 空 → 「対角公差が入力されていません！！」で判定なし。
- 全ヵ所: S1 空 → kou=0 → 差0のみ○。

### 11.7 ラテラルボー（M1=1.0）
入力 0.3→"0.5"→○, 1.0→"1.0"→○, 1.2→"1.5"→×, −0.7→"-1.0"→○, 0.05→"0.0"→○, "-"→保持（判定はスキップ：非数値）, 5.3→"5.3"→×。M1="" → 1.0 でも ×。H494 で空セル確定 → "0.0" が書かれ○（M1≠0時）。H494 行11以降は判定なし。

### 11.8 歪（全ヵ所）
5N: 幅歪 2.0→○, 2.1→×, −2.0→○; 丈歪 3.0→○, 3.01→×。一般 M5=0.6: 0.6→○, 0.7→×。M5 空→×。

### 11.9 MAX/MIN
- 全ヵ所: B7:B10={10.5,10.6,10.4,10.5}, D列同様、F,H 空 → .Min全体=0 → AtuVar=[10.4(B),min(D),0,0] → MIN=min(B,D)。B列のみ値・D空 → MIN=0（W4 空のまま）。
- キャスク: CountSt=4, MIN=1, MAX=3, 厚の最小が H7(列8) のみ → V4 空（バグ再現）。CountSt=3, MIN=1 で最小が B8 → V4=値, 下線あり。

### 11.10 連番チェック（CheckCaskSequence）
既存[1-10],[11-20]: 新[21-30]→OK, 新[22-30]→NG, 新[11-15]→OK(11-20を置換扱い), 既存に[21-30]もあり新[11-15]→NG, 新[1-5]→チェック不要でOK（保存後 1以外削除）, 新[30-20]→「スタートNo がエンドNo より大きい値です。」。
SaveCaskToDB: 管理番号 既存MAX=7 → 8、テーブル空 → 1。

### 11.11 エンドNo照合
B列最終入力行=56（A56=50）、DB最終(スタートNo最大)エンドNo=50 → True、49 → False＋メッセージ。

### 11.12 機器文字列（LVC）
厚 28 → `ﾏｲｸﾛ:LVCM1　　ｺﾝﾍﾞｯｸｽ:LVC C10`、28.001→M2、54→M2、74→M3、74.5→`ﾉｷﾞｽ:LVCN5…`、300（行14–18に 299.5 と 300.2）→`ﾉｷﾞｽ：LVCN5,N6　　ｺﾝﾍﾞｯｸｽ:LVCC11`、0→書込なし。Yout="H494", 幅150, 厚10 → `ﾏｲｸﾛ:LVCM1　　幅:ﾉｷﾞｽ:N5_丈:ｺﾝﾍﾞｯｸｽ:LVC C10`。Test ライン 厚10 → `ﾏｲｸﾛ:M1　　ｺﾝﾍﾞｯｸｽ:`。

### 11.13 文字列ユーティリティ
- RemoveDuplicates("田中 鈴木 田中") → "田中 鈴木"; **("田中 中") → "田中"**（部分一致で欠落、再現要否確認）。
- ExtractNthText("(10.0÷2)(20.5÷3)", 2) → "20.5"; ("(10.0)", 1) → ""（÷無）; (x, 0) → ""。
- UnitRecord(0.5)=0.5, (0.51)=1.0, (0.505)=0, (999.5)=0。

### 11.14 保存
Lot=H3084S0, 保存日 2026-10-30 → `C:\検査履歴\2026.10\20261030_H3084S0\20261030_H3084S0.xlsx`、2回目 `…\20261030_H3084S0.002.xlsx`、パスワード `H3084S0.2026.10.30`。

### 11.15 重量メッセージ
36×1528×3053 → 450.077kg → 「ゲージによる測定必要」。10×150×2502.5 → 10.06kg → 「ゲージ不要 * 両面定盤歪測定必要」。

---

## 12. 未使用コード候補（断定ではない）

| 対象 | 根拠 |
|---|---|
| `FiveN材転送` (Mit:1058) | 呼出しはコメントのみ（UFMit:1292, 1402） |
| `5N原紙` シート / `5N` シート / `原紙印刷範囲` の 5N 分岐 | 5N シートを作るのは FiveN材転送のみ。原紙印刷範囲は三菱 CB2 からのみ呼ばれ Choice="5N" では到達しない |
| `Judg_PorM` (Mit:1379) | 呼出しはコメント内（Cask:1458）のみ |
| `CountParenthesesSets` (Mit:1435) | 呼出しなし |
| `対角再チェック` (Cask:1236) | 呼出しはコメント（Cask:1587） |
| `UFCask.Hflag` | 書込のみ（UFCask:259,301）、参照なし |
| `MAXTex_KeyPress` (UFCask:29) | コントロール名 MAXText と不一致で発火しない（推測） |
| KySh `Case 7` (KySh:36-43) | `Case 2 To 8` が先に一致し到達不能 |
| `UF対角起動` の `"全ヵ所検査検査"` 分岐 (Cask:1908) | シート名誤記で一致しない |
| `ZenSh2` モジュール | VB_Base が FCFB3D2A、内容が旧版 → 残骸（イベント不発）と推測 |
| `KySh` / `ZenSh` / `H494_SH`（原紙側） | 原紙は非表示で編集されないため実質発火しない（推測）。コピー元として必要 |
| `Mitsubishi.cls`, `FiN.cls` 等の空モジュール | コードなし（推測: 三菱受託原紙・5N原紙のコードネーム） |
| UFCask `CommandButton6` | 点数を3に強制し、エンドNo照合結果に関係なく保存する → 試験用ボタンの疑い |
| UFMit `CommandButton18` | 警告表示のみで後続処理なし（確認用ボタンの可能性） |
| `UFCask.CommandButton14` | アクティブシート A7/A8 を読上げ（用途不明） |
| `MoveTg` の `If flag_Blank = False Then '空白なし` 空ブロック、AlTerRange 内の重複処理 | 実質無処理 |
| UFMit `Mitsubishi.bas` 内 三菱用 ｺﾋﾟ-用 S列の①〜④変換 (Mit:661-672) | 列削除後 S列は常に空 |
| 旧 ComboBox1_Change（UFCask:557-711, UFMit:877-1248 コメント） | 2026.05.29 ADO化前の旧実装（Arrbox） |
| 三菱関連処理全般（転送済、NPS保存、ブロックNOチェック） | 2024.1.11 の列追加に追随していない参照が多く、運用されていない可能性（推測、要確認） |

---

## 13. 不明点・要確認事項

### 高優先（結果・出力に影響）
1. **UFCask の DB フィールド名**（"オーダー番号","ロット番号","用途コード","設計_設備コース" 全角）と UFMit（半角）が異なる。`仕掛` テーブルの実フィールド名、ADO のフィールド名照合が全角/半角を同一視するか。片方は常に `DB読み込みエラー` になっていないか。
2. **三菱モードの列参照ずれ**: ブロックNOチェック(列27,29)、歪/外観必須チェック(列24,25,26)、転送済への転記（N←対角判定記号, O←ラテラル, P←幅歪, Q←丈歪, B←ゲージ, C←空）。現行運用で三菱モードは使われているか。VBA同等（ずれごと再現）か、意図通りに修正か。
3. 三菱時 `全ヵ所専用保存` で SetPath が NPS の月フォルダ＋DayLotName（フォルダではない）に上書きされた後 `NewFileName(SetPath & "\")` → 存在しないフォルダで実行時エラーの疑い（Mit:571-614）。サーバ `全ヵ所検査_履歴` へのコピーが実施されているか。
4. 丸徳等 LineName が L-1/LVC/HVC/ﾊﾞﾗﾝｻｰ/ﾄｯﾄ/Test 以外の時、全ヵ所/5N保存の SrvPath="" → FileCopy 先が "\{ファイル名}"。実際の挙動。
5. **キャスク MIN 計算の行/列変数入れ替え（Cask:395-396）**と全ヵ所 MIN 補正の到達不能 Case（Mit:197,203）。印刷物の MIN 値に影響。再現するか修正するか。
6. ループ判定 `KenJudg((Cells(i, a)), …)`（Cask:996,1004,1012）: 括弧付き引数で Range ではなく値が渡る → 実行時エラー(424/13)の可能性。UFCask 起動時・印刷時に実際エラーになっていないか（データがある状態で要実機確認）。
7. `保存用` シートのキャスク行で厚が D列に書かれ巾で上書き（UFCask:156-157）、C列空。意図。
8. H494 印刷フッタの測定機器が AC3 参照（Y2 に書いた値が出ない）。
9. H494 の判定ログが丸徳ログフォルダに出る（Lname 未設定）。
10. 全ヵ所 対角1(列19)入力時に W列(ラテラルボー値)まで消去される（ZenSh:49）。意図か旧レイアウト残りか。
11. 三菱モードの歪判定が現行コードでは行われない（旧版は1.0基準）。`FiveOP`/`一般` がOptionButton同グループか、三菱選択時に両方Falseになるか（フォーム定義が無いため不明）。5N→三菱 切替時に FiveOP=True が残ると5N固定公差で判定される可能性。
12. 新規発行（全ヵ所 CB1）で `Work="ON"` に固定 → 5N で再選択しないと工場共有【梱包】５N へのコピー条件(Work="OFF")を満たさない可能性。UserForm_Activate が再実行されるタイミング。
13. H494 モードで検査員が `Kname` に入り `H494_Kname` が空のまま → 印刷/保存の検査員欄が空になる可能性。
14. キャスク CountSt=3 時の厚3入力後、カーソルが非表示列 H へ移動する（KySh Case 2 To 8）。保護状態次第で H 列に4点目が入り得る。実運用の操作手順。
15. 対角公差 S1 がロット再入力時にクリアされず、等級が3/4以外かつ指示なしの場合は前ロットの値が残る可能性（UFMit:722-733）。
16. UFCask/UFMitsubishi の相互参照（例: 検査判定 `If UFCask.Visible`、対角 `UFCask.Visible`、UFMit:714）により非表示フォームが自動ロードされ Initialize の副作用（シートActivate・DBアクセス・非表示シートActivateでのエラー）が起きる可能性。シートイベントの On Error Resume Next で判定がスキップされていないか。
17. 全ヵ所の行範囲: シートは70行だが MoveTg・下線リセットは36行目まで。31枚以上の運用があるか。
18. `Application.Undo` による変更前値取得（各シート Change）の実効性（マクロ変更後は Undo 不可）。移植では「変更前値」を直接保持する方式で良いか。

### 中優先
19. Diagonal/ReDiagonal/Lateralbow_J/ReLateralbow のテーブル（Judg: LookupDiag/LookupLat、別担当）との整合。
20. 歪公差 M5/O5・ラテラル公差 M1 の QP_DB 側上書き順序（ComboBox1_Change 内 DistArray/FormReflectDist、別担当）。
21. `Ainput_Startup` の検査法（キャスク "L2"、全ヵ所/H494 は Inspection1-3/op1-3 の Caption）の値一覧（フォーム定義なしで不明）。
22. ブック自体の保存（作業中保存シートを含む xlsm の上書き保存）がどこで行われるか。
23. 不良要因印刷_{Line} の構造・印刷範囲（不良要因印刷範囲、別担当）。
24. キャスク原紙 V5「入力がない場合JIS4級で判定」、H494原紙 B5 の △ 判定 — 現行コードと不一致。帳票文言の扱い。
25. 手入力の扱い: キャスク/全ヵ所/H494 の厚列、キャスク巾丈列は手入力が消去される。全ヵ所/H494 の巾丈列は手入力保持。H494フラグ時の全ヵ所 巾列・H494の巾列への手入力は丈/対角の空きへ移送される。これを仕様として再現するか。

### 移植制約上の課題（Python標準ライブラリのみ）
26. パスワード付き xlsx の生成（標準ライブラリでは暗号化 OOXML 作成不可）・読取専用属性・印刷（A4/A3, ヘッダ/フッタ, FitToPages）・音声読上げ・WAV 再生の代替方針。
27. DAO による物理削除（STEP2）は VBScript+ADODB の `DELETE` 文で代替可能か（同一条件 `LotNo=… AND 検査No=… AND スタートNo<>1`）。
28. Single 精度比較（対角）、`val()` の15桁往復、VBA `Format` の丸め（四捨五入）を Python で厳密再現する必要。
