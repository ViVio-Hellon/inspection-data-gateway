# G2 構造解析レポート：判定・公差・入力・校正

対象：UF検査判定.xlsm（日軽稲沢・梱包課）
担当モジュール：
- `【UFKen_04】Judg.bas.bas`（2488行）
- `【UFKen_05】OutTolerance.bas.bas`（2954行）
- `【UFKen_09】Input.bas.bas`（1673行）
- `【UFKen_10】Proofreading.bas.bas`（654行）
- `UF校正.frm.bas`、`UFdiagonal.frm.bas`、`UF対角.frm.bas`、`UF異常値発見.frm.bas`、`UFNG.frm.bas`、`UFOK.frm.bas`、`UFcount.frm.bas`

表記ルール
- `J:123` は Judg の123行目、`O:` は OutTolerance、`I:` は Input、`P:` は Proofreading。フォームはファイル名と行番号で書く（例 `UF校正:220`）。他モジュールは名前をそのまま書く（例 `SaveLoad:256`）。
- 「推測」と書いたものはコードから断定できない内容。「要実機確認」は Excel/VBA の実機で挙動を確かめる必要があるもの。
- 業務上の結果を左右する VBA 特有の挙動（型変換・丸め・比較）は ⚠ を付けた。

---

## 0. 移植の前提として最初に押さえる VBA の挙動（全体共通）

Python 移植で結果が食い違う原因になる点を先にまとめる。

| # | 挙動 | 影響する箇所 | Python での扱い |
|---|---|---|---|
| V1 | **DB から取った値はすべて文字列**。`GetRecordsArr` → `adoRecords` が `Recordset.GetString(adClipString, …, vbTab, "<#(仮)改行#>")` で全体を1本の文字列にしてから Split している（`SQL_追加:419-441`）。Null は `""` になる。数値は ADO の文字列変換（15桁精度・ロケール依存）を通る | 公差テーブル、BH下限値、VC厚、注意_包装仕様、校正ストレートエッヂ など全部 | DB 値は「ADO が出す文字列」として受け取り、比較時に VBA と同じ規則で数値化する。Null→"" の置き換えも同じにする。値にタブが入っていると列がずれる不具合もそのまま残る |
| V2 | **ADO の `.Filter` は SQL の WHERE ではない**。`SELECT * FROM 表` で全件取ってから、クライアント側で `Recordset.Filter` をかけている（`SQL_追加:293-296`）。`LIKE '*52S*'` は ADO Filter の書き方 | BhSearch、PackagingSpecificationNo、VcConv、StraightEdge など | SQLite に移すときは Filter と同じ結果になるか確認する（ワイルドカードは `*`/`%`、文字列比較の大文字小文字の扱い） |
| V3 | **Single（単精度）の誤差**。`Diagonal(a As Single…)`、`UnitRecord(n As Single)`、`Height_Check` の変数はすべて Single | 公差参照、0.5刻み変換、高さチェック | `struct` で float32 を再現する。境界値（例 1.01 → Single では 1.00999999）の扱いが変わる（§3.8） |
| V4 | **`Val(Single)` は内部で `CStr` を通る**。CStr(Single) は有効7桁、CStr(Double) は有効15桁で丸めた文字列になる。`Val(kou) >= Val(Tai)` のような書き方は、実質「7桁に丸めた10進数どうしの比較」になる | DiagonalJudg、LateralJudg、Height_Check | 比較前に有効桁で丸めて揃える。`repr(float)` を使うと結果が変わる |
| V5 | **CStr(Double) は有効15桁**。例：`1.005*1000` は Python では `1004.9999999999999` だが、VBA の CStr では `"1005"` | StCon（Aインプット固定長） | 文字列化は `'%.15g'` 相当で行う |
| V6 | **Format 関数の丸めは「0から遠い方への四捨五入」**（銀行丸めではない）。`Format(x,"0.0")`、`"0.000"`、`"0"` を多用している。一方 `Round()` は銀行丸め | 公差上下限（ApK…）、Diag 表示、異常値しきい値、BH 入力値（`Format(SBH,"0")`） | Format 相当は `Decimal.quantize(ROUND_HALF_UP)`、Round 相当は ROUND_HALF_EVEN。ただし入力は2進数の double なので、`Decimal(str(x))` と `Decimal(x)` のどちらを使うかで .5 境界の結果が変わる（要実機確認） |
| V7 | **文字列と数値の比較**。`If 文字列 > 数値`（例 `Format(SAP,"0.00") > ErrorA`）は数値どうしの比較になる。数値にできない文字列（`""` など）を数値と比べると実行時エラー13（Type mismatch）**の可能性がある** | O:871 `Select Case Rig`、O:1989 `.ZS = 0`、J:212 `CDbl(TB) > ComparisonValue` | 要実機確認（§12 Q1） |
| V8 | **`=` と `Like` はバイナリ比較**（モジュールに Option Compare 指定なし）。ただし `StrComp(…, vbTextCompare)` と `InStr(…, vbTextCompare)` はテキスト比較（大文字小文字を区別しない。日本語環境では全角/半角・ひらがな/カタカナも同一視する可能性あり） | IsWildcardMatch（J:2403, 2409） | 完全一致・部分一致は文字種を区別しない比較、前方一致・後方一致は区別する比較、と分けて実装する |
| V9 | **Long への代入は銀行丸め**（CLng(2.5)=2、CLng(3.5)=4） | BhSearch の戻り値（As Long）、Height_Check の Mai（As Long） | ROUND_HALF_EVEN |
| V10 | **`-0`**：`UnitRecord` の負の処理 `Val(x)*(-1)` で、0 のとき -0.0 になる可能性がある。VBA の Format は "0.0" と表示する見込み、Python の `format(-0.0,'.1f')` は `'-0.0'` | ラテラルボー・直角度の 0 入力 | -0.0 は 0.0 に正規化する |

---

## 1. 各モジュール概要

### 1.1 【UFKen_04】Judg — 寸法判定の中心
- 厚/巾/丈の公差判定（`AHTJudg`）、対角線差の判定（`DiagonalJudg`）、ラテラルボーと直角度の絶対値判定（`LateralJudg`）、BH（バーコール硬度）判定（`BhJudg`）、歪（定盤・ゲージ・フリーソリ）判定（`HizumiJudg`）を持つ。
- 判定結果を UF1 のセル色（OK=薄緑 13434828、NG=コーラル 8421631）、効果音、OK/NG 画像の2秒表示（`Sleep 2000`）、`Measured.log` への書き込みで出す。
- `LoopJudg` / `HizLoopJudg`：画面全体の再判定（読み込み時、校正後、NG消去後に色を塗り直す）。
- 製品規格の公差表（直角度・ラテラルボー、4級/3級）を `品質基準.accdb` から読み込んでメモリに保持する（2026.09.17 にハードコードから DB 参照に変更。旧表はコメントで残っている J:1276-1775）。
- `MaterialsMsgBox` / `PackagingSpecificationNo`：包装仕様NO を条件に `梱包資材マスタ.accdb` の「注意_包装仕様」から注意文を出す。条件は用途コード、寸法、納入先、取引先（ワイルドカード・除外指定あり）。
- `AdvanceCheck`：先行データ票が必要かを判定する。
- `UnitRecord`：ラテラルボー・直角度の測定値を 0.5 刻みに変換する（Q-S-QA-ALL-153）。

### 1.2 【UFKen_05】OutTolerance — 公差外の処理・入力チェック・補助
- `Out_Tolerance`：NG 色のセルを見つけて確認フォーム（UFConfi）を出す。NG 消去（`NG_Back`）、許容/保留（`NG_Pass`）、NG 転記（`NG_Posting` → シート「異常ロット用_<ライン>」）。
- `異常値除去`：公差の2倍を外れた値を異常値とし、UF異常値発見 を出す。
- 印刷前の入力漏れチェック：`Distortion_Check`、`Lateralbow_Check`、`Angle_Check`、`Visual_Check`、`H_Check`、`WithorWithout_Gauge`（Q-S-QA-DIM-173 平坦度測定基準：重量 35kg を境にゲージ要否を決める）。
- 画面の入力状況をラベル色で示す：`CheckLabelsByCategory` / `GetLabelState` / `GaugeJudgment` / `IsValidFormat`。
- 少枚数検査の表示切り替え（`SmallNumberCount`）、ComboBox や ListBox の初期化（`CBAdd`、`Visual_Add`、`Reasons_Add`）、外観・不良要因の入力（`LcColor`、`SelectionInput`）。
- 導電率チェック（QA-021-08、75S/14S の焼入れ材）、重量計算 `Weight`、検入高さの枚数計算 `CalculateSheets`。
- 汎用メッセージ表示 `MsgError`（UFMsgBox を出し、ログを書き、UF1 を隠す）。

### 1.3 【UFKen_09】Input — Aインプット（アイカス／ホスト送信）と梱包高さチェック
- `Ainput_Startup`：UFAinput を起動し、Lot情報履歴（なければ SIKATWOYEAR.accdb）から値を入れる。
- `Height_Check`：梱包の山高さ実測値と「板厚×枚数＋合紙＋VC膜＋外装紙」の理論値を比べ、枚数の過不足を検出する。
- `FormArray` / `ZerCon` / `StCon` / `LinConv`：ホスト送信用の固定長電文（`KEN:AIN:0185…`）と、梱包課保存用 CSV の配列を作る。
- `A_Output` / `ExpCSV` / `A_CsvSave`：`C:\Tools\測定データ電子化\hostsend.txt` と、履歴 CSV を書き出す。
- `App_Activation` / `InterRuptText`：送信用アプリ `SokuteiDataTrans.exe` を起動し、`result.txt` の結果を読む。
- `履歴保存`：検査票を PDF にし、原本 xlsx をパスワード付き・読み取り専用で保存する（ローカル `C:\検査履歴` とサーバー）。
- 山高さ（最大15山）を動的ラベルで管理する関数群。

### 1.4 【UFKen_10】Proofreading と UF校正 — 直尺（スケール）の校正
- 直ごとに、測長器（伝送コンベックス）の校正が済んでいるかを判定する（`校正()`）。済んでいなければ測定入力をブロックし、UF校正 を起動する（`Proofread_Load`）。
- UF校正 は2段階：①1000mm 直尺の測定（品証検証用、999.5～1000.5）②ライン別ストレートエッヂの全長測定（DB「校正ストレートエッヂ」の下限～上限）。②の合格で「校正記録」シートに1行追加し、`ProofreadFlag="ok"` にする。
- 月ごとの不合格値ランキング（`校正ランク付け`、`品証ランク付け`）。

### 1.5 小さいフォーム
| フォーム | 役割 | 現状 |
|---|---|---|
| UFdiagonal | 対角線差の公差等級（3/4）を手で選んで再判定する | J:642-648, 701-707 の呼び出しがコメントアウト済み。**呼び出し元なし（推測：死んでいる）** |
| UF対角 | 対角の画像を表示し、2秒待って閉じる。`ActiveSheet` 名（"検査","少枚数検査","極検査"）と行番号で画像を選ぶ | Cask.bas:1903 から使用（キャスクのシート UI） |
| UF異常値発見 | 異常値のしきい値を表示する。ボタンでその値を消す（`NG_Back_TB`） | 使用中（`AbnoFormStart` から） |
| UFNG / UFOK | NG/OK 画像を3秒表示して閉じる | キャスク（Cask.bas）と全ヵ所検査（ZenSh2.cls）で使用中 |
| UFcount | 枚数と位置を読み上げ（`Speak`）、3秒で閉じる | `UFCask.Read` を参照。Cask.bas:1894 から使用 |

---

## 2. Sub/Function 一覧

凡例：Pub=Public、Pri=Private。「呼出元」は grep の結果（コメント行は除く）。

### 2.1 Judg
| 名前 | 種別 | 引数 | 戻り値 | 役割 | 主な呼出元 → 呼出先 |
|---|---|---|---|---|---|
| Judg_ImageSize | Pub Sub | n:"Y"/"N"/"Neko" | - | OK/NG/猫の画像を表示し、作業フレームを隠す（猫画像は乱数1～15） | JudgColor, TextJudg, LateralbowJudgColor, RightAngleJudgColor |
| HizumiJudg | Pub Sub | TB:TextBox, LoopJ | - | 歪1マス（ZS～ZF 定盤、GS～GF ゲージ）の判定と色付け | UF1 Z?/G?_AfterUpdate(UF1:5048-5167), HizLoopJudg → 音楽合格/不合格 |
| HizLoopJudg | Pub Sub | n=0, tMP="Loop" | - | 歪12マスすべてを再判定。空欄は白 | UF1 各所, UF校正:292, FormRelat, NG_Back |
| AdvanceCheck | Pub Func | Temp(LotNo) | Boolean | 先行データ票が必要か | UF1:3643,4317,4850, MaterialsMsgBox → Arrbox(SIKALOTNOW) |
| AHTJudg | Pub Func | Ter:Label, KP, KM As Double | Boolean | 厚/巾/丈 の上下限判定 | JudgColor, LoopJudg |
| ReDiag | Pub Func | - | (なし) | 再読み込み時に対角線差（Diag?）を再計算 | UF1:584, SaveLoad:4040,4181 |
| DiagonalJudg | Pub Func | Tai, kou As Single | Boolean | 対角線差 ≤ 公差 | JudgColor, LoopJudg, UFdiagonal |
| LateralJudg | Pub Func | a As Single, kou As String | Boolean | \|a\| ≤ 公差（ラテラルボー・直角度） | LoopJudg, LateralbowJudgColor, RightAngleJudgColor |
| BhJudg | Pub Func | Ter:TextBox, Zai, TYO | "OK"/"NG"/"不要" | BH 下限値判定 | UF1 ?BH_AfterUpdate(1854-1874), LoopJudg → BhSearch(Quality_DB) |
| simplyjudge | Pub Sub | TB, n | - | BH の色・音・ログ。NG なら工程停止などのメッセージ | UF1:1854-1874 → WriteLog, MsgError |
| JudgColor | Pub Sub | TB:Label, n:"A"/"H"/"T"/"D" | - | 1マスの判定・色・音・画像・ログ | TextJudg |
| TextJudg | Pub Sub | TB:Label | - | セル名から判定種別（厚/巾/丈/対角）を決めて振り分ける | FormRelat.LoopText:285 → JudgColor, NGErase |
| LoopJudg | Pub Sub | n=0, j=1, k=6 | - | 全セルを再判定（j～k 行：1=S…6=F） | UF1, UFConfi:65, UF校正:291, FormRelat, NG_Back |
| Diagonal / ReDiagonal | Pub Func | a=巾, b=丈 (Single) | Single | 直角度公差 4級 / 3級 | SaveLoad:285-291, UF1:4179-4181, UFMitsubishi, UFdiagonal → LookupDiag |
| Lateralbow_J / ReLateralbow | Pub Func | a=巾, b=丈, C=厚 | Single | ラテラルボー公差 4級 / 3級 | SaveLoad:305-318, UF1:4193-4196, UFMitsubishi → LookupLat |
| ReloadTolerance | Pub Sub | - | - | 公差キャッシュをクリアして再読み込み | **呼出元なし**（VBE から手動で実行する想定） |
| LookupDiag / LookupLat | Pri Func | grade, a, b(, C) | Single | メモリ上の表を検索（該当なしは0） | ↑ → EnsureLoaded, InRange |
| InRange | Pri Func | v, lo, hi | Boolean | lo ≤ v ≤ hi（Null/空欄は不一致） | Lookup* |
| EnsureLoaded | Pri Func | - | Boolean | 初回だけ DB から読み込む。失敗したら以後は再試行しない | Lookup* → LoadTable |
| LoadTable | Pri Sub | FilePath, TableName, colNames, data, idx | - | 列名から列番号を引き、全レコードを取る | EnsureLoaded → GetFieldsArr/GetRecordsArr |
| Test_Tolerance | Pub Sub | - | - | 動作確認用 MsgBox | **呼出元なし** |
| LateralbowJudgColor | Pub Sub | TB:TextBox | - | ラテラルボー判定・色・音・画像・ログ | UF1 Lateral?_BeforeUpdate(1305-1446) |
| RightAngleJudgColor | Pub Sub | TB:TextBox | - | 直角度（公差 1.0 固定）判定 | UF1 Diag?Kou_BeforeUpdate(1115-1257) |
| MaterialsMsgBox | Pub Sub | OrdNo | - | 包装仕様の注意文と先行データ票の案内を出す | UF1:1789, SaveLoad:443,449 → PackagingSpecificationNo, AdvanceCheck, MsgError |
| PackagingSpecificationNo | Pub Func | PackagingNO, 用途, 厚, 幅, 丈, 材質, 調質, 納入先, 取引先 | String "形状_コメント…" | 注意_包装仕様テーブルを条件で絞り込む | MaterialsMsgBox |
| IsMatchConditionField | Pri Func | FieldValue, target | Boolean | カンマ区切り OR、≠ は除外、* はワイルドカード | ↑ |
| IsWildcardMatch | Pri Func | Pattern, target | Boolean | 前方一致/後方一致/部分一致/完全一致 | ↑ |
| NflatRewrite | Pub Sub | - | - | ReSize.Tag="有" にして Posting | **呼出元なし** |
| UnitRecord | Pub Func | n As Single | Single | 0.5 刻みへの切り上げ変換 | UF1:1109-1440, ZenSh*, H494_SH* |
| LookingYellow | Pub Sub | - | - | 黄色背景と赤枠をリセット | UF1:4364 |
| CheckForLot | Pub Func | ws, Lot, Dcaunt | Boolean | シートの最終5行の Dcaunt 列に Lot が含まれるか | UF1:4502 |

### 2.2 OutTolerance
| 名前 | 種別 | 引数 | 戻り値 | 役割 | 主な呼出元 → 呼出先 |
|---|---|---|---|---|---|
| SmallNumberCount | Pub Sub | TB:検査回数(1-5) | - | 少枚数検査の表示切り替え、ラテラルボー/直角度欄の表示 | UF1(4), SaveLoad(1) → GetLotInfo, DbReturnItem, InitialSetting |
| UFPack_Open | Pub Sub | n(包装仕様NO), PcLot | - | 包装仕様書の画像（JPG）を表示 | UF1, UFCask, UFMitsubishi |
| NameList_Open | Pub Sub | n:1/2/3 | - | 検査員不一致／日報／梱包枚数の一覧を表示 | UF1, Worker, Test, UFNameList |
| WorkerName_Open | Pub Sub | - | - | 検査員名ラベルを表示 | UF1 T1-T3_Change |
| ConfiReason_Open | Pub Sub | Posi, n, Lname, Moto, delCount | - | 削除理由フォーム（UFConfi のページ1） | UF1, LabelCls |
| Confi_Open | Pub Sub | 同上 | - | 公差外確認フォーム（UFConfi のページ0）。n=96 ラテラルボー/97 直角度/98 定盤歪/99 ゲージ歪 | Out_Tolerance |
| MsgError | Pub Sub | Trm, Lname, n(アイコン), Temp | - | 警告・注意ボックス（カンマで改行）＋ログ＋UF1.Hide | 多数 |
| AbnoFormStart | Pub Sub | n(値), TB(名) | - | 異常値フォームをモーダルで出す | 異常値除去 |
| Era_flag | Pub Func | TB, KP, KM | Boolean | 異常値か（KM 未満または KP 超） | 異常値除去 |
| Weight | Pub Func | 厚, 幅, 長 | Double | 重量 kg（比重2.68、小数第3位で切り捨て） | 多数（Print, QP_DB, Mitsubishi 他） |
| Out_Tolerance | Pub Sub | Temp | - | NG セルを探して Confi_Open | UF1(25 箇所) |
| NG_Erase | Pub Sub | Posi, All | - | 作業保存シートのその位置の行を消して再読み込み | UFConfi |
| NG_Back | Pub Sub | Posi, n | - | NG 色のセルだけ消し、EraseCount に記録して再判定 | UFConfi |
| NG_Pass | Pub Sub | Posi, Temp="許容" | - | NG?ラベルに「許容」や「保留」を書く | UFConfi, UF1 |
| NG_Posting | Pub Sub | Posi | - | 異常ロット用シートへ転記（値＋NG セル色） | UFConfi(3), Print |
| 異常値除去 | Pub Sub | TB:Label | - | 公差×2 を外れたら異常値フォーム | FormRelat.LoopText |
| NG_Back_TB | Pub Sub | TB(名) | - | 1セルを消して Posting | UF異常値発見 |
| CBAdd | Pub Sub | ATU=6 | - | 歪形状 ComboBox を初期化 | UF1, FormRelat, SaveLoad, QP_DB |
| CB_UpDownAdd / CBG_UpDownAdd | Pub Sub | CB | - | 上/下、Ⅰ～Ⅲ型 の選択肢を入れる | UF1 |
| Visual_Add | Pub Sub | - | - | 外観異常名（DB）、山数、不良要因リストを入れる | UF1 |
| Reasons_Add | Pub Sub | - | - | 削除理由リストを入れる | UFConfi |
| LcColor | Pub Sub | La, n | - | 外観入力欄を選んで欠陥名を入れ、ログを書く | UF1(12) |
| GetSelectedItems | Pub Func | La:ListBox | String | 選ばれた項目をスペース区切りで返す（不良要因は 名_NO） | LcColor, SelectionInput, UF1, UFMitsubishi |
| btn_allClea | Pub Sub | La | - | 選択をすべて解除 | UF1 |
| DrawListBoxContent | Pub Sub | FormTemp, ThickTemp | - | 選んだ山数に応じて高さ入力欄を表示 | UF1, UFMitsubishi, SaveLoad, Mitsubishi → CalculateSheets |
| WithorWithout_Gauge | Pub Func | - | String（空=OK） | 平坦度測定基準の未入力チェック（旧版） | UF1:1654（歪ページ）, UF1:3664（印刷前） |
| Distortion_Check | Pub Func | n:"Z"/"G" | String | 歪3項目（形状・上下・値）の入力が揃っているか | UF1:3659-3660（印刷前） |
| Angle_Check | Pub Func | - | String(位置) | 直角度の未入力 | UF1:3665 |
| Lateralbow_Check | Pub Func | - | String(位置) | ラテラルボーの未入力 | UF1:3661, Print:1041 |
| Visual_Check | Pub Func | - | String(位置) | 外観の未入力 | UF1:3662 |
| BH_Check | Pub Func | Temp | String | BH の未入力 | **呼出元なし**（UF1:3878 はコメント） |
| H_Check | Pub Func | - | "入力なし"/"" | 検入高さ7欄がすべて空か | UF1:3663 |
| Conductiv_Check | Pub Func | - | Boolean | 導電率の測定が必要な材質・調質か | UF1:4702, SaveLoad:388,2971 |
| Conductiv_Coolingtime_Check | Pub Func | lottemp | String | QF 空冷時間チェックリストの写真確認No | SaveLoad:398,2972 |
| ConductivLoad | Pub Sub | - | - | 導電率測定結果記入表をコピーして記入 | UF1(1) |
| SelectionInput | Pub Sub | TE, n | - | 不良要因印刷シートに枚数を記入 | UF1, UFMitsubishi |
| TeColor / TecReturn | Pub | La, n | -/String | 不良要因欄の選択色管理 | UF1, UFMitsubishi |
| FindRightString | Func(暗黙Pub) | InputString | String | コース文字列から最終熱処理（ANI/ANF/CAI/CAF）を右から探す | SaveLoad, Test |
| CalculateSheets | Func(暗黙Pub) | FormTemp, ThicknessA | Single | Σ(高さ/厚) で枚数を出し、式を表示 | UF1(12), FormRelat, SaveLoad, DrawListBoxContent |
| UFtag_Array | Pub Func | x, y | Variant(x..y,1..2) | 歪 ComboBox/TextBox の Tag を「形状_上下_値」にする | SaveLoad(3), Print |
| allow_Check | Pub Func | - | String | 許容/保留の位置の一覧 | Print:609 |
| GetLabelState | Pub Func | FormName, labelName | "可"/"否"/"すべて非表示"/"無効" | コントロールの入力状態 | CheckLabelsByCategory |
| CheckLabelsByCategory | Pub Sub | Category(S..F) | - | CheckT? ラベルを緑/赤/白にする | UF1(42), FormRelat(12) |
| GaugeJudgment | Pub Func | Temp(S..F) | String | WithorWithout_Gauge の位置別の改良版 | CheckLabelsByCategory |
| IsValidFormat | Pub Func | inputStr | Boolean | 「文字_文字_数値」形式か | GaugeJudgment |
| IsValidNumericWithSlash | Pub Func | inputStr | Boolean | "/" を含み、除くと数値か | IsValidFormat |
| SingleUnitDistortion | Pub Func | n, Temp | String | Distortion_Check の位置別の改良版 | **呼出元なし** |

### 2.3 Input
| 名前 | 種別 | 引数 | 戻り値 | 役割 | 主な呼出元 → 呼出先 |
|---|---|---|---|---|---|
| GetDPI / GetActualWindowWidth / ResizeWindow | Pub | - / newWidthPt | Long / - | Win32 API でウィンドウのサイズ調整 | UFAinput |
| No_Search | Pub Sub | LotNo | - | 引当一覧を Hiki_Lot シートに作る（Lot情報履歴 → SIKACOMP.accdb） | Ainput_Startup |
| Ainput_Startup | Pub Sub | Temp:"UF1"/"Cask"/"Mitsubishi"/"H494", KenM | - | UFAinput を起動 | UF1(3), UFCask, UFMitsubishi |
| Number_Inp | Pub Sub | TB:Label | - | InputBox で 梱包NO/枚数/包み数/山数 を入力 | UFAinput |
| Height_Check | Pub Func | ATU1(min), ATU2(max), Hei, Mai, OrderNo | Boolean（**True=NG**） | 梱包高さの許容値チェック | UFAinput:268,1033 → VcConv, MsgError |
| Nz | Pri Func | val | String | Null → "" | Height_Check |
| VcConv | Pub Func | no(品名ｺｰﾄﾞ) | Single | VC 膜厚（伝送用ファイル.accdb の VC厚） | Height_Check |
| A_Output | Pub Sub | FormName, Lot, Ken | - | ホスト電文と保存 CSV の作成 | UFAinput, Worker |
| ExpCSV | Pub Sub | sh, fileName, FilePath, Sep, maxi, maxj | - | シートを CSV/TXT で上書き出力 | A_Output, A_CsvSave, UF1(15) |
| A_CsvSave | Pub Sub | ws, Lot, Ken | - | 月/Lot フォルダに yyyymmdd_検査No(.nn).csv を保存 | A_Output |
| A保存 | Pub Sub | ws, Lot, Ken | - | xlsx 保存（コメントに「使ってない」とある） | **呼出元なし** |
| 履歴保存 | Pub Sub | - | - | PDF 化と原本 xlsx の保存 | UF1(1) → PDFCREATE, SaveDateLocate, NewFileName |
| SaveDateLocate | Pub Func | DatePath, LineN, n | String | ライン\yyyy.mm\dd日 のフォルダを作り、パスを返す | 履歴保存, Test |
| CreateFolder | Func | FullFolderPath | Boolean | 多階層フォルダを一括で作る | 多数 |
| FormArray | Pub Func | FormName, n:1/2/3 | Variant | 1=梱包課用配列、2=ホスト電文（1要素）、3=項目名＋値の2列配列 | A_Output |
| StCon | Pub Func | Temp As Double, Pits:"厚"/"巾"/"丈" | String(7) | ×1000/×100/×10 して7桁にゼロ埋め | FormArray |
| LinConv | Pub Func | line | String | ライン名 → NIS コード | FormArray |
| ZerCon | Pub Func | DateNo, buf | String | 項目番号ごとの固定長ゼロ埋め | FormArray |
| HostCon | Pub Func | tMP | String | A000～A003 の説明文 | **呼出元なし** |
| ResultCon | Pub Func | tMP | String | result.txt の文言を短い文に変換 | InterRuptText |
| HostDataConfi | Pub Func | APPPATH, Ltext | String | 現在の hostsend.txt の引当No/梱包No | UFAinput |
| InterRuptText | Pub Sub | APPPATH, Ltext, WaitTime | - | result.txt の最終行を読み、成功/失敗を判定 | UFAinput |
| Loadtext | Pub Func | fileName, encoding | String | ADODB.Stream で UTF-8 として読む | HostDataConfi, UFCask |
| App_Activation | Pub Sub | APPPATH, App | - | SokuteiDataTrans.exe を起動 | UFAinput |
| MergeVariantArraysCol / Call_MergeArray_Row | Pub Func | 2次元配列×2 | Variant | 列方向 / 行方向に結合 | FormArray / Visual_Add 他 |
| Array2D | Pub Func | InArray, rc, n | Variant | 1次元 → 2次元 | **呼出元なし** |
| GetKHeight/SetKHeight/GetDynHeightLabel/CalcMountHeight/AllHeightFilled/NextEmptyHeightIndex/ClearAllKHeight | Pub | n 等 | | 山高さ（DynHL_1～15）の操作。合計は Σ Val(caption) | UFAinput, FormArray, Height_Check |

### 2.4 Proofreading / UF校正
| 名前 | 種別 | 引数 | 戻り値 | 役割 | 主な呼出元 → 呼出先 |
|---|---|---|---|---|---|
| 校正_Open | Pub Sub | wname, Lname | - | UF校正 を起動（ゲージの案内文をラインごとに設定） | UF1 T1-T3_Change, Proofread_Load |
| 校正 | Pub Func | STtime("hh:mm") | Boolean（**True=未校正**） | 校正記録の最終行の日付と直を照合 | UF1:4110,4132,4152,4296 → TimeCheck, YeaDayrCheck |
| SafeNum | Pub Func | S, success(ByRef) | Double | 数値文字列を安全に変換（カンマを除去、小数点2個以上は不可） | 測定完了 |
| 測定完了 | Pub Sub | - | - | ストレートエッヂ測定の合否と記録 | UF校正 INPG_AfterUpdate → StraightEdge |
| 品証測定 | Pub Sub | - | - | 1000mm 測定の合否と記録（品証検証シート） | UF校正 INP1000_AfterUpdate |
| 校正ランク付け | Pub Sub | - | - | 当月の赤字（不合格）値の上位・下位3位 | UF校正:174, 品証測定 |
| 品証ランク付け | Pub Sub | - | - | 品証検証の当月不合格値の上位・下位3位 | UF校正 Initialize, CommandButton4 |
| Proofread_Load | Pub Sub | - | - | 未校正のときメッセージを出して UF校正 を起動 | UF1:4619（測定入力時） |
| UF校正.TerConv | Pri Func | Ter | String | 伝送文字列 → 数値文字列 | INP1000/INPG_AfterUpdate |
| UF校正.TerMaker | Pub Func | Trm | "DT"/"#L"/"" | 伝送機器の種類を判別 | TerConv |
| UF校正 各イベント | Pri Sub | | | §7 で説明 | |

---

## 3. 判定ロジック・計算式・公差・閾値・丸め・境界値

### 3.0 セル（ラベル）名の割り当て（どのセルが何の項目か）
UF1 の測定ラベル名は「行文字＋番号」。行文字：S, U, C, D, E, F（F は4山のとき。番号は1～10まで）。画面の位置番号は `UF1.S.Caption` 等（少枚数では "1"～"5"）。

TextJudg（J:756-831）と 異常値除去（O:1350-1413）の割り当て：

| 番号 | 通常（薄板/厚板/少枚数/スラブ） | 極厚（`UF1.極厚`=True） |
|---|---|---|
| 1-4 | 厚 "A" | 厚 "A" |
| 5-6 | 巾 "H" | 巾 "H" |
| 7-8 | 丈 "T" | 巾 "H"（7,8,9 も巾） |
| 9-10 | 対角 "D"（9 と 10 の差が対角線差） | 9 は巾、10 も巾（Len=3 で "10"） |
| 11-14 | （非表示） | 丈 "T"（「耳付き」2022.06 で丈が2つ追加） |

規格値の出どころ（担当外のモジュールで設定される。参考として記載）
- `UF1.Am/Ap/Hm/Hp/Tm/Tp`（マイナス/プラス公差）、`ATU/Hab/Tak`（製造寸法）：SaveLoad:220-253。DB の「板幅公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ…」を使い、オーダー値が0なら製造値を使う。
- 上下限：`ApK = Format(Val(ATU)+Val(Ap),"0.000")`、`AmK = Format(Val(ATU)-Val(Am),"0.000")`、`HpK/HmK/TpK/TmK` は `"0.0"`（SaveLoad:256-261、FormRelat:6175-6180）。耳付き（G_flag）のときは巾・丈の上下限を空にする（SaveLoad:264-266）。
- `TaikText`（直角度＝対角線差の公差）：DB の「指示_直角度」があればそれ（全角→半角、`"0.0"`）。なければ等級 `.TAIK` が "4" なら `Diagonal(Hab,Tak)`、"3" なら `ReDiagonal`、`Format "0.0"`（SaveLoad:279-292）。
- `LateralText`：「指示_ﾗﾃﾗﾙﾎﾞｰ」があればそれ。なければ等級 `.Lateral` の4/3で `Lateralbow_J/ReLateralbow(Hab,Tak,ATU)`、`Format "0.0"`、0 なら ""（SaveLoad:295-322）。
  ⚠ UF1.Lateral_Change（UF1:4187-4201）では 4級を `Format(...,"0")`（整数丸め）にしているが、代入先はデバッグ表示用の `DebugLate` で、判定には使わない。
- `HaHiz/TaHiz/Naka/Mimi`（歪公差 幅/丈/中/耳）、`HURIS`（フリーソリ公差）、`flatnesText`：FormRelat で設定（担当外）。

### 3.1 厚・巾・丈：AHTJudg（J:362-369）
```vb
If IsNumeric(Ter.Caption) = False Then Exit Function          ' → False（NG扱い）
If (val(Ter.Caption) < KM) Or (val(Ter.Caption) > KP) Then AHTJudg = False
ElseIf (val(Ter.Caption) >= KM) And (val(Ter.Caption) <= KP) Then AHTJudg = True
```
- **境界は両端を含む**：下限 ≤ 値 ≤ 上限 で OK。
- 下限 KM＝`AmK`/`HmK`/`TmK`、上限 KP＝`ApK`/`HpK`/`TpK`（Format 済みの文字列 → 引数で Double に変換）。
- 数値でないとき False。Val を使うので "12.3abc" は 12.3 として判定される（ただし IsNumeric で先に弾かれる）。
- JudgColor の中で、マイナス公差の欄が空なら判定しない（J:533-535）：`If n="A" And UF1.Am="" Then GoTo SkipD`（H→Hm、T→Tm）。**プラス公差 Ap 等が空かどうかは見ていない**。
- LoopJudg では次のとおり。
  - `If (UF1.ApK = "") Or (UF1.AmK = "") Then Exit Sub`（J:846）：厚の上下限が空なら全体を再判定しない。
  - 厚：`UF1.Am=""` なら厚の判定をとばし、色は白（J:881-889, 934）。
  - 巾：`HpK=""` または `HmK=""` なら巾の判定をとばす（J:956）。
  - 丈7-8：`TpK=""` または `TmK=""` なら `GoTo SkipL`（J:983）。⚠ **このとき対角の判定（J:1006-1054）もとばされる**。
  - 丈11-14（極厚）：空欄チェックがない。⚠ `TpK=""` のまま `AHTJudg(…, UF1.TpK, …)` を呼ぶと、Double 引数への変換でエラー13 になる可能性がある（要実機確認）。

### 3.2 対角線差：JudgColor "D"（J:584-715）、DiagonalJudg（J:380-394）
- 計算：9 と 10 のどちらかを入れたときに、もう一方が入っていれば
  `Tai = Abs(val(UF1(X & 10)) - val(TB))`（Tai は Single）。
- `kou = .TaikText`（Single へ代入）。⚠ JudgColor は `On Error Resume Next`（J:483）なので、TaikText が "" だと代入でエラーになり kou=0 のまま → DiagonalJudg が False → **対角公差が空なら NG（×）になる**。
- DiagonalJudg：
```vb
If IsNumeric(Tai) = False Then Exit Function   ' Single なので常に数値
If kou = 0 Then Exit Function                    ' 公差0 → False（Tai=0 でも NG）
If Tai = 0 Then DiagonalJudg = True: Exit Function
If val(kou) >= val(Tai) Then True Else False     ' 差 ≤ 公差 で OK（等しいとき OK）
```
  ⚠ Val(Single) は CStr（有効7桁）を通るので、10進数に丸めて比べている（V4）。
- 判定後、`UF1("Diag" & 行) = Tai`（Single → Caption は CStr で7桁）。一方 ReDiag（J:370-379）は `Format(Abs(val(9)-val(10)),"0.0")` で、極厚でないときだけ再計算する。→ 表示の書式が2通りある。
- 9 を入れて 10 が空のときは、旧形式のログ `"*UFKen:" & 値 & ":(" & 名 & "):_:" & Lname` を書いて判定しない（J:587-592）。
- LoopJudg（J:1007-1054）：`IsNumeric(UF1.TaikText)` のときだけ。`DiagS` が空でなく `DiagonalJudg(DiagS, TaikText)` なら 9 と 10 を緑にする。それ以外は事前に塗ったコーラル（NG）のまま。
- 画面の「直角度」欄（Diag?Kou）とは別物。2023.11.24 に「対角公差から直角度に変更」（J:713-714、O:93-103）。

### 3.3 直角度：RightAngleJudgColor（J:1894-1985）
- 公差は**固定値 1.0**（Q-S-QA-DIM-154 対角線差及び直角度測定基準、J:1902-1904）：`Dim kou As String: kou = 1#` → 文字列 "1"。
- 判定：`LateralJudg(a, kou)` → **\|a\| ≤ 1.0 で OK**。
- 入力前処理（UF1:1104-1113）：数値なら `UnitRecord(Val(x))` → `Format(…,"0.0")` を書き戻してから判定する。"-" は「測定不要」としてそのまま残す。
- 直角度欄を表示するのは `UF1.Youto = "R115"`（用途コード）のときだけ（O:231-245, 266-273）。印刷前の未入力チェックも R115 のときだけ（UF1:3693）。
- LoopJudg（J:1090-1099）：`LateralJudg(Diag?Kou, "1.0")`。⚠ J:927-928 で E と D が入れ替わっている：
  ```vb
  If UF1.DiagEKou.Visible Then UF1("DiagDKou").BackColor = Ncol
  If UF1.DiagDKou.Visible Then UF1("DiagEKou").BackColor = Ncol
  ```
  （表示状態の判定と色を塗る欄が逆。4山などで D/E の表示が異なるときに色がずれる。不具合候補）

### 3.4 ラテラルボー：LateralbowJudgColor（J:1796-1892）、LateralJudg（J:395-409）
- 入力が "" または "-" なら何もしない。数値でなければ何もしない。
- `UF1.LateralText` が数値でなければ「ﾗﾃﾗﾙﾎﾞｰ公差がありません。(規定寸法範囲外)」を出し、`TB.Tag="範囲外"` にして終了（判定しない）。
- LateralJudg：
```vb
If IsNumeric(a) = False Then Exit Function
If kou = "0" Or kou = "" Then Exit Function      ' → False（NG）
If a = 0 Then LateralJudg = True: Exit Function
If val(kou) >= Abs(val(a)) Then True Else False  ' |a| ≤ 公差 で OK
```
  ⚠ kou が "0.0" のときは `kou = "0"` に一致しないので、`0 >= |a|` の比較になる（a≠0 なら NG）。
- 入力前処理は直角度と同じ：UnitRecord → Format "0.0"（UF1:1299-1300）。
- 公差の値：§3.0 の LateralText（4級/3級の表、または技術指示）。
- ラテラルボー欄の表示：`SmallNumberCount` の中で、`DbReturnItem(納入, 取引, 用途, 材質, 調質, 厚, 幅, 丈)` の戻り文字列に "ラテラルボー" が含まれ、かつ `Right(UF1.Additional.Caption,4) <> "検査不要"` のとき（O:198-229）。
- LoopJudg（J:1059-1069）：⚠ **`If IsNumeric(UF1.LateralText) = False Then Exit Sub`** なので、ラテラルボー公差が数値でないと、その後の BH、直角度、空欄を白に戻す処理（J:1070-1113）がすべてとばされる。事前にコーラルで塗った表示中のラテラルボー欄と直角度欄（J:911-929）は NG 色のまま残る。

### 3.5 0.5 刻み変換：UnitRecord（J:2421-2443）
```vb
If Sgn(n) = 1 Then flag = True    ' 正。0 と負は False
n = Abs(n)
Select Case n
Case 0:            0
Case 0.1 To 0.5:   0.5
Case 0.51 To 1#:   1.0
Case 1.01 To 1.5:  1.5
Case 1.51 To 2#:   2.0
Case 2.01 To 2.5:  2.5
Case 2.51 To 3#:   3.0
Case 3.01 To 3.5:  3.5
Case 3.51 To 4#:   4.0
Case 4.01 To 4.5:  4.5
Case 4.51 To 5#:   5.0
Case 5.01 To 999:  n（そのまま）
End Select                          ' どれにも当たらなければ 0
If Not flag Then UnitRecord = val(UnitRecord) * (-1)
```
- 0.5 刻みの「切り上げ」。5.0 を超えたら変換しない。
- ⚠ **すき間**：0 < n < 0.1、0.5 < n < 0.51 などの区間（2桁目以降の小数）と 999 超は **0 を返す**。
- ⚠ **Single の誤差**：n は Single。`CSng(0.51)=0.50999999`、1.01、1.51、2.01、2.51、3.01、3.51 も下限よりわずかに小さくなり、**区間から外れて 0 になる**（4.01, 4.51, 5.01 は Single で大きくなるので区間内）。入力は3文字に制限（maxLength=3、"-" で始まれば4）されているが、".51" は入力できる → 結果 0。
- 1桁小数の入力（例 2.6 → Single 2.5999999）は "2.51 To 3" に入って 3.0 になる。問題ない。
- 0 のときは `Val(0)*(-1)` で -0 になる可能性がある（V10）。

### 3.6 BH（バーコール硬度）：BhJudg（J:410-418）、simplyjudge（J:420-478）
```vb
Dim bh As Long: bh = BhSearch(Zai, TYO)      ' 下限値（Long に丸める）
If bh = 0 Or Ter.Text = "" Then BhJudg = "不要": Exit Function
If bh > val(Ter.Text) Then BhJudg = "NG"
If bh <= val(Ter.Text) Then BhJudg = "OK"   ' 下限値以上で OK
```
- BhSearch（Quality_DB:1454-1490、担当外）：`品質基準.accdb` の「Q_S_KE_ALL_053」。材質が4文字以上なら右3文字にして `[材質] LIKE '*xxx*'`（ADO Filter）。調質は完全一致。**一致した最後の行**の「バーコール下限値」を Long に変換（V9 の銀行丸め）。見つからなければ 0 → "不要"。
- 入力側（UF1:1852-1876）：`UF1.SBH = Format(UF1.SBH,"0")`（整数に四捨五入、V6）してから判定する。
- simplyjudge の色：OK=緑、NG=コーラル、不要=白（16777215）。
- simplyjudge の NG 時メッセージ（J:471-472）：最後に検査した位置を countL で決め、`Iti` に入れる。
  - S だけ検査済み（`countL(1)`）で SBH が NG → 「作業中止・保留・ロット異常連絡票発行」
  - E まで検査済み（`countL(5)`）で EBH が NG → 「梱包完了させてAインプット・保留・ロット異常連絡票発行」
  - それ以外（U/C/D 位置、4山以下）ではメッセージを出さない（`Iti=""`）。
- LoopJudg（J:1071-1088）：BH が OK なら**白**（緑ではない）、NG ならコーラル。

### 3.7 歪（平坦度）：HizumiJudg（J:134-291）
入力欄：`Z?`（定盤）と `G?`（ゲージ）。?=S,U,C,D,E,F。形状 ComboBox：`CB?1`（定盤：幅/丈/耳/中）、`CB?2`（上/下、Ⅰ～Ⅲ型）、`CB?3`（ゲージ：幅/丈、フリーソリなら上/下）、`CB?4`。

判定の順序
1. 空欄なら白にして終了。
2. **フリーソリ（2026.03.05 から先頭で独立して判定）**（J:155-176）：
   `LorR="G" And UF1.HURIS<>"" And IsNumeric(HURIS) And flatnesText="ﾌﾘｰｿﾘ要"` のとき、"/" より前の部分 `NumberT2` を取り、`CDbl(NumberT2) > CDbl(HURIS)` なら NG、それ以外（≤）は OK。
3. `CB?1` が空なら終了（**G 欄でも CB?1（定盤の形状）が必要**）。
4. 比較値：CB?1 が「幅」→ HaHiz、「丈」→ TaHiz、「中」→ Naka、「耳」→ Mimi。
5. **両方ゲージ測定のとき**（`LZS.Caption="幅ｹﾞｰｼﾞ" And LGS.Caption="丈ｹﾞｰｼﾞ"`、J:198-231）：Z 欄は CB?1 の値、G 欄は CB?3 の値で幅/丈の公差と比べる。`CDbl(TB) > 公差` なら NG、`≤` なら OK。
   ⚠ この分岐では "/" を処理せず `CDbl(TB)` を直接使う（"12/3" のような入力はエラー）。G 欄の公差 `ComValueG` が "" のときも、数値と文字列の比較でエラーになる可能性がある（V7）。
6. 通常のとき：比較値が "" なら CB?3 を見て、空でなければ SkipG へ。比較値が数値でなければ終了。
7. "xx/x山" 形式なら "/" より前を数値として使う（J:245-247）。
8. **Z 欄だけを判定**：`CDbl(NumberT) > ComparisonValue` なら NG、`<=` なら OK（公差ちょうどは OK）。G 欄は白（「ゲージの判定は不要だった為停止」J:273）。
- 音を鳴らすのは `LoopJ=""`（単独入力のとき）または `UF1.MoveFlag="on"`（MouseDown のとき）だけ。最後に `MoveFlag=""` に戻す。
- 公差の文字列と CDbl の比較は、文字列を数値に変換して比べる（V7）。

### 3.8 公差表の参照：LookupDiag/LookupLat（J:1153-1209）
- DB：`QualityPath & "品質基準.accdb"`
  - 「製品規格直角度」：管理番号, 等級, 巾_下限, 巾_上限, 丈_下限, 丈_上限, 公差
  - 「製品規格ラテラルボー」：管理番号, 等級, 厚_下限, 厚_上限, 巾_下限, 巾_上限, 丈_下限, 丈_上限, 公差
- 引数：Diagonal(a=巾, b=丈)、Lateralbow_J(a=巾, b=丈, C=厚)。
- 前処理：`w = Round(CDbl(a), 2)`（Single の誤差対策。Round は銀行丸め）。
- 条件：`CStr(等級) = "4"`（または "3"）かつ `InRange`：**下限 ≤ v ≤ 上限（両端を含む）**。下限か上限が Null/空なら不一致。
- 複数の行が当てはまるときは**管理番号が最小の行**を使う（`m_Diag(r, NO) < bestNo` は文字列と Double の比較で、数値として比べる）。
- 当てはまる行がなければ **0**。呼び出し側は 0 を「公差なし」として扱う（LateralText="" など）。
- 戻り値は `CSng(公差)`。
- 旧ハードコード表（J:1276-1775）との違い：旧表は `Select Case 1# To 300#, 300.1 To 600# …` で、300.05 のような値はどこにも当たらなかった。新方式の結果は DB 表の境界値の決め方しだい（要確認）。動作確認の期待値（J:1783-1786）：`Diagonal(500,2000)=1`、`ReDiagonal(1000,5000)=5`、`Lateralbow_J(500,2000,3.2)=1.5`、`ReLateralbow(300,4000,8)=15`。

旧表（DB の移行元。DB と照合するときの参照として転記）
- 直角度4級（巾 \| 丈 1-3700 / 3701-7000）：1-300: 0.8/1.0、300.1-600: 1.0/1.5、600.1-900: 1.5/2.0、900.1-1200: 2.0/2.5、1200.1-2400: 2.5/3.0
- 直角度3級：1-300: 1.5/2.0、300.1-600: 2.0/3.0、600.1-900: 3.0/4.0、900.1-1200: 4.0/5.0、1200.1-2400: 5.0/6.0
- ラテラルボー4級：厚・巾の区分によらず、丈 1-1500:1.0、-2300:1.5、-3000:2.0、-3800:3.5、-4500:4.0、-5500:5.0、-6000:6.0、-7000:6.0
- ラテラルボー3級：厚・巾によって 102/38/19/13… などに分かれる（J:1590-1775 を参照）

### 3.9 異常値の検出：異常値除去（O:1321-1415）
- `UF1.Abno.Caption = "異常値on"` のときだけ動く。
- しきい値（公差の2倍。UF異常値発見 の表示文「異常値の閾値　公差の2.0倍とする」）：
  ```
  iatu1 = Format(ATU - Am*2, "0.000")   iatu2 = Format(ATU + Ap*2, "0.000")
  ihaba1 = Format(HABA - Hm*2, "0.0")   ihaba2 = Format(HABA + Hp*2, "0.0")
  itake1 = Format(TAKE - Tm*2, "0.0")   itake2 = Format(TAKE + Tp*2, "0.0")
  itai1 = Format(Sqr(HABA^2 + TAKE^2) - 6, "0.0")   itai2 = ... + 6
  ```
  （HABA, TAKE は Single。Format の結果を Double に代入しなおす。）
- Era_flag：`Val(TB) < KM Or Val(TB) > KP` なら異常値。**境界値ちょうどは正常**。
- 項目ごとに、マイナス公差（Am/Hm/Tm）が空なら終了（対角も `UF1.Tm=""` で判定）。極厚では対角を判定しない。
- 異常値のとき：`UF1.Abno.Tag="NG"` にし、UF異常値発見 をモーダルで出す → ボタンで `NG_Back_TB`（値を消して Posting）。LoopText に戻ると TextJudg は空欄で終了する。
- UF異常値発見 の表示用の値は、作業保存シートの1行目 B～J（ATU, HABA, TAKE, Am, Ap, Hm, Hp, Tm, Tp）から同じ式で計算する（`UF異常値発見:47-82`）。

### 3.10 公差外確認のポップアップ：Out_Tolerance（O:818-977）
- `st = UF1.INP.Tag`（フォーカス位置のセル名。TagSet でコントロール名が入る）。
- `Rig`：st が3文字なら右2文字、2文字なら右1文字、それ以外は ""。
- 作業ごとの「確認を出すタイミング」（Rig がこの値のとき、その行の NG を全部確認する）：

| 作業 | Rig | 確認するセル | F 行 |
|---|---|---|---|
| 薄板 | 2,4,6,8,10 | 1-10 | あり |
| スラブ | 2,4,6,8,10 | 1-10 | なし |
| 厚板 または 少枚数 | 4,10 | 1-10 | あり |
| それ以外（極厚） | 4,7,10,12,14 | 1-14 | なし |

- その行の `NG?` ラベルが「許容」ならループを抜ける（「保留」なら抜けない）。
- NG 色（8421631）のセルごとに `Confi_Open(行, i, Lname, 値, Val(DeleteC))` を呼ぶ。⚠ UFConfi は1つしかなく、vbModeless で表示するのでループは止まらない → **実際に表示されるのは最後（i が最大）の NG セル**（推測。要実機確認）。
- 続けて、st が `Lateral?`/`Diag?Kou`/`Z?`/`G?` の場合：その欄が NG 色なら Confi_Open(行, 96/97/98/99, …) を呼ぶ（ラテラルボーと直角度は表示中のときだけ）。
- Confi_Open：厚板/薄板/少枚数で n=9 または 10 なら、`MotoData` を「線差 」＋ Diag? にする（O:689-696）。
- ⚠ Rig が "" や "S" のような数字でない文字列のとき（Lateral、Z、G の入力後）、`Select Case Rig` の `Case 2,…` で数値と比べるとエラー13 になる可能性がある（V7・要実機確認）。

### 3.11 重量：Weight（O:813-817）
```vb
Weight = Application.RoundDown((Thickness * Width * length * 2.68 / 10 ^ 6), 3)
```
- 単位は mm×mm×mm、比重 2.68 → kg。**小数第3位で切り捨て**（Excel の ROUNDDOWN）。
- どれかが数値でない、または "" なら 0。
- ⚠ Excel の ROUNDDOWN は15桁に丸めてから処理するので、Python の `math.floor(x*1000)/1000` とは結果が変わることがある（例：2.68×… が 12.345999999 のような値になる場合）。

### 3.12 平坦度測定基準（ゲージの要否）：WithorWithout_Gauge（O:1842-1983）、GaugeJudgment（O:2746-2838）
- 入力元：作業保存シート `O2:P7`（A面：O=定盤、P=ゲージ。行2-7=S,U,C,D,E,F）、`O8:P13`（B面）。値が "_" または "__" なら空欄とみなす。
- `TempW = Weight(ATU, Hab, Tak)`。0 なら判定しない（""）。
- 「厚の測定がある」＝ `S1..E1` のキャプションが空でない（F は `Yama="4"` のときだけ）。

WithorWithout_Gauge（旧版。印刷前に使う）。⚠ **後から代入したメッセージが優先される。途中の Exit Function は、それまでに入ったメッセージを返す。**
1. 厚の測定あり かつ A面定盤なし → "?のA面 定盤 未入力"（C は「**2**のA面…」と書かれている。誤記と思われるが表示はこのまま）
2. ≤35kg：A面あり・B面なし → "?のB面 定盤 未入力_ﾊﾟﾀｰﾝA"。B面あり・A面なし → "?のA面 定盤 未入力_ﾊﾟﾀｰﾝA"
3. >35kg：どこか1か所でも A・B 両方あれば即 Exit Function（それまでのメッセージを返す）
4. >35kg：A面のみでゲージAなし → "?のA面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"
5. >35kg：B面のみでゲージBなし → "?のB面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"。⚠ `For i = 1 To 5` なので F（Case 6）は判定されない。
- F の flagA(6) は Yama 条件なし、flagB(6) は `Yama="4"` が必要（O:1886-1887）。
- **境界：35.0kg ちょうどはパターンA（≤35）**。

GaugeJudgment（新版。CheckLabelsByCategory の色付けに使う）
- 位置を1か所指定。入力ありの判定は `<>"" And IsValidFormat(…)`。
- メッセージは全角の「パターンA/B」「ゲージ」。最初に当たった条件で Exit Function。

IsValidFormat（O:2840-2870）："_" で分けて3要素でなければ False。どこかが空欄のとき、3要素目が "0.00" でなければ False。数値/数値/文字 なら False。文字/文字/文字 なら IsValidNumericWithSlash(3要素目)。それ以外は True。

### 3.13 歪3項目の入力チェック：Distortion_Check（O:1984-2073）
- 先に、値が 0 なら形状 ComboBox を空にする：`If .ZS = 0 Then .CBS1 = ""`（ZF→CBF1 は**ない**。GS～GF は CB?3 を空にする）。⚠ TextBox が空のとき `"" = 0` を比べることになる（V7・要実機確認）。
- "Z"：形状/上下/値 のどれかが入っていて、かつどれかが空で、かつ値 `<> 0` → "定盤:?"。⚠ F 行は `.ZE <> 0` を見ている（O:2022 誤記）。
- "G"：flatnesText="" のときは CB?3/CB?4/G? の3項目で判定 → "ｹﾞｰｼﾞ:?"。"ﾌﾘｰｿﾘ要" のときは CB?3 と G? の2項目 → "ﾌﾘｰｿﾘ:?"。
- 複数該当なら最後（F 側）のメッセージが残る。

### 3.14 その他の印刷前チェック（すべて「最初に当たった位置」のキャプションを返す）
- Lateralbow_Check：`?1 <> "" And Lateral?.Visible And Lateral? = ""`。印刷を止めるのは `UF1.Sumisei.Caption <> """"`（ダブルクォート1文字と比べている）かつ Additional に「ﾗﾃﾗﾙﾎﾞｰ検査不要」がないとき（UF1:3688）。
- Angle_Check：Diag?Kou について同じ条件。止めるのは Youto="R115" のとき。
- Visual_Check：`?1 <> ""` かつ（上外観 = "" または 下外観 = ""）。
- H_Check：KenHBox1～7 がすべて空なら "入力なし"。
- 印刷前チェックの順序（UF1:3667-3714）：定盤 → ゲージ → 歪AB → ラテラルボー → 直角度 → 外観 → 検入高さ → 山数。

### 3.15 先行データ票：AdvanceCheck（J:334-361）
- `SIKALOTNOW.accdb`（Arrbox。DBflag が空なら MainPath、そうでなければ NextPath）の「仕掛」を `[ﾛｯﾄ番号]='Lot'` で引く。
- 条件：
  - flag1：`Val(製造板厚) <= 3#`（3mm 以下。**3.0 を含む**）
  - flag2：`品質ｸﾞﾚｰﾄﾞ_表面処理` が "S" でも "T" でもない
  - flag3：用途ｺｰﾄﾞの頭が "T"、または頭2文字が "D1"/"D2"
- **(flag1 And flag2) Or (flag2 And flag3)** なら True。
- 列が見つからないと列番号が0になり、添字エラーになる可能性がある（エラー処理なし）。

### 3.16 包装仕様の注意文：PackagingSpecificationNo（J:2188-2330）、MaterialsMsgBox（J:1987-2185）
- MaterialsMsgBox：`SIKAODRNOW.accdb`「仕掛」を `[受注番号]='OrdNo'` で引き、包装仕様NO、用途ｺｰﾄﾞ、受注材質、受注調質、納入先名称、取引先名称を取る。**厚・巾・丈は UF1 の製造寸法（ATU/Hab/Tak）**を使う（2024.8.28 変更）。
  - DB の場所：DBflag が "" なら MainPath。`ReadCheck` が失敗するか、TimeDesignation か、FileDesignation=False なら NextPath に切り替える。
- PackagingSpecificationNo：`PackingPath & "梱包資材マスタ.accdb"`「注意_包装仕様」を `[包装仕様NO]='…'` で引く。
  - レコード1件のとき（Cont=1）：フラグ "" なら無条件でコメント。"用途コード" なら完全一致。"寸法" なら3寸法とも `下 ≤ x ≤ 上`（Val で数値化。両端を含む）。"納入先" なら納入先≠"" かつ IsMatchConditionField。"寸法納入" は 納入先条件＋寸法。"寸法取引" は 取引先≠"" かつ 寸法、納入先条件、取引先条件。結果は `形状 & "_" & コメント`（条件に合わなければ "形状_"）。
  - 2件以上：合ったコメントを `Kari & "," & 既存` のように**前に追加していく**（DB の順番と逆になり、最後に "," が残る）。⚠ **"寸法納入" に合ったときだけ、前に追加せず上書きする**（J:2306。それまでのコメントが消える）。結果は `TempHiki(1,形状) & "_" & …`。
- MaterialsMsgBox の後処理：
  - 結果が ""（レコードなし）：先行データ票が必要なら「先行データ票が必要です。」を出し、必要です.wav を鳴らす。
  - `Split(Temp,"_")`：1要素目が "コイル" なら何も表示せずに終了。それ以外は**最後の要素**をコメントとして使う。⚠ コメント中の "_" より前は捨てられる。
  - カンマで分けた要素数 nCount=UBound：-1（例 "板_"）なら**先行データ票の案内も出さずに終了**（⚠ 結果が "" のときとふるまいが違う）。1以下なら1要素目だけ。2以上なら `Like "*要素*"` で部分一致を除いて連結（Like の特殊文字 `[ ] # ? *` が入ると誤判定）。
  - 先行が必要なら ",先行データ票が必要です。" を追加。"直角度の測定" を含めば 直角度.wav。Dictionary で重複を除き、`MsgError(…, 4)` で表示。
- IsMatchConditionField / IsWildcardMatch（J:2346-2411）：マスタ側の値を Trim。空、"＊"、"*" なら True。カンマ区切りで OR（除外指定どうしも OR）。先頭が "≠" なら「一致しなければ True」。`*xxx*` は InStr（テキスト比較）、`xxx*` は Left の **バイナリ比較**、`*xxx` は Right の**バイナリ比較**、`xxx` は StrComp（テキスト比較）。⚠ 例 "≠A,≠B" は、ほぼすべての値で True になる（コメントに説明あり）。

### 3.17 導電率（QA-021-08）：Conductiv_Check（O:2172-2183）
- 材質 "75S" で調質 "T651" または "T6"、材質 "14S" で調質 "T6" または "T651" なら True（完全一致）。N17S/17S/24S の条件はコメントアウト済み。
- Conductiv_Coolingtime_Check：`QualityPath\75S・14S焼入れチェック\QF空冷時間チェックリスト.xlsx` の Sheet1 を、別の Excel インスタンスで読み取り専用で開く。5行目から下、B列=LotNo、A列=写真確認No。**一致した最後の行**の A を返す（空なら ""）。一致しなければ "LotNo記載がない"。⚠ 呼び出し側（SaveLoad:409）は `"LotNoがない"` と比べているので、この分岐には入らず「何か文字列」の扱いで画面に表示される。

### 3.18 検入高さの枚数計算：CalculateSheets（O:2509-2564）
- `Quantity = Σ (KenHBox_i / 厚)`（i=1..6、IsNumeric のものだけ。Single で累積）。戻り値は `Format(Quantity,"0.0000")` → Single。
- 表示：`(h÷厚) + (h÷厚)=` & Quantity（Single の CStr）。
- 厚が 0、または数値の欄がひとつもなければ何もしない（UFMitsubishi のときは KenHBox7 の値を Tag に写す）。

### 3.19 梱包高さチェック：Height_Check（I:430-584）
引数：ATU1=AMin（測定厚の最小）、ATU2=AMax（最大）、Hei=Σ山高さ（Single）、Mai=M_Count（= MAI1×COU1（+MAI2×COU2）、Long）、OrderNo。
戻り値：**True = NG**（MsgError を出し、山高さをクリアする）。False = 合格。⚠ **Lot情報履歴シートがない、または OrderNo が見つからないときも False（合格扱い）**。OrderNo が見つからないときは MsgBox を出して UF1.Show（I:452-458）。

Lot情報履歴（2行目から）の列：2=ｵｰﾀﾞｰNo、17=合紙（Q）、18=VC（R）、19=VC_A（S）、20=VC_B（T）。

```
AtuMax = ATU2, AtuMin = ATU1
-- 合紙
unitAis = 0
If 合紙列 <> 0        → unitAis = 0.029 (AiText="通常")
ElseIf Remain = 1     → unitAis = 0.042 (AiText="ストア")
If unitAis > 0 Or (外注 And Ai=True) Then
   If 外注 And Ai=True Then unitAis = 0.029 : AiText="通常"
   A1 = unitAis
   extraMai = 1 + Val(COU1)(空でなければ) + Val(COU2)(空でなければ)
   Ais = Val(Mai + extraMai - 1) * A1        ' = (Mai + COU1 + COU2) × A1
   If 外注 And Ai=False → A1=0, Ais=0, AiText=0
-- VC
If VC列 = 1 Then
   If Remain <> 1 Or Youto ∈ {H191,H242,H241} Then
      V1 = VcConv(VC_A) + VcConv(VC_B)   （空ならその分は0）
      Vcs = V1 × Mai ;  If 外注 And Vab=True → V1=0, Vcs=0
Else  VCAB.Caption = 0
-- 外装紙
If Val(COU1)>1 Or Val(COU2)>1 Then Gai = (COU1(>1のとき) + COU2(>1のとき)) × (2×0.119 + 0.119)   ' 1包あたり 0.357
If 外注 And Ga=True → Gai = 0
-- 判定
Gosa = Ais + Vcs + Gai
GP = AtuMax × Mai + Gosa ;  GM = AtuMin × Mai + Gosa
ErrorA = Val(AtuMax) × 3/4 ;  ErrorB = Val(AtuMin) × 1/3
SAP = |Val(Hei) − Val(GP)| ;  SAM = |Val(Hei) − Val(GM)|
NG1: Val(Hei) > GP  And Format(SAP,"0.00") > ErrorA  → "[許容値ｵｰﾊﾞｰ]"   推定枚数 = (Hei−Gai)/(AtuMax+A1+V1)
NG2: Val(Hei) < GM  And Format(SAM,"0.00") > ErrorB  → "[枚数不足(?)]"   推定枚数 = (Hei−Gai)/(AtuMin+A1+V1)
```
- **境界**：差が ErrorA（または ErrorB）ちょうどなら合格。差は小数第2位に四捨五入（Format）してから比べる。
- メッセージ：`"[許容値ｵｰﾊﾞｰ] *測定高さ:" & Format(Hei,"0.00") & "*推定枚数:" & Format(…,"0.00") & " *入力枚数:" & Mai & "枚"`。MsgError の Lname は UFAinput.LineName、アイコンは 2、Temp は "Ainput"。
- ⚠ **コメントアウトされた旧版（I:592-751）とふるまいが違う**（コメント「元のSelect Caseの計算結果と一致」は正しくない）：
  1. 個装（COU）がないとき：旧版は `(Mai+1)×A1`、現行は `Mai×A1`。
  2. 外注で合紙ありのとき：旧版は Ai=True で 0、現行は Ai=True で 0.029 として計算し、**Ai=False で 0**（逆）。
  3. 外装紙：旧版は `COU1 + COU2×0.357`（演算子の優先順位の不具合）で、COU2="" の判定も違う。現行は `(COU1+COU2)×0.357`。
  4. VC：旧版はストアなら一律 0。現行はストアでも高精度用途（H191/H242/H241）なら加算する。VC=1・ストア・高精度以外のときは VCAB を設定しない。
  → **どちらを正とするか業務確認が必要**（§12）。移植では現行コードに合わせるのが原則。
- VcConv（I:753-803）：`伝送用ファイル.accdb`「VC厚」（Test のときは PackingPath、それ以外は QualityPath）を `[品名ｺｰﾄﾞ]='no'` で引き、「厚-係数」を返す。**登録がなければ「VC厚登録なし,担当者に連絡のこと」と出して 0.08406977**（登録VCの全平均厚、2022.12.19）。ファイルがなければ MsgBox を出して 0。

### 3.20 Aインプットの固定長電文：FormArray(n=2)、ZerCon、StCon、LinConv
電文 = `"KEN:AIN:" & "0185" & A3 & A4 & … & A40 & ライン`（区切りなしの1行。I:1256-1259）。

| 項目 | 内容 | 空欄のとき | 値があるとき |
|---|---|---|---|
| A3 | 引当No | "00000000" | **そのまま（桁を揃えない）** |
| A4 | 梱包No | "00" | Right("0"&buf,2) |
| A5 | 出荷ｽﾄｱ | "0" | そのまま |
| A6-A11 | 厚Min, 巾Min, 丈Min, 厚Max, 巾Max, 丈Max | **""（ZerCon を通らない。I:1225-1230）** | StCon：厚×1000→Right("0000"&CStr,7)、巾×100→Right("00000"&,7)、丈×10→Right("000000"&,7)。0 なら "" |
| A12, A14 | 枚数1,2 | "00000" | Right("0000"&buf,5) |
| A13, A15 | 包み数1,2 | "000" | Right("00"&buf,3) |
| A16 | GW | "00000" | 5桁 |
| A17, A19, A21 | 風袋1-3 | "00000" | 5桁 |
| A18, A20 | GW2/3 | "00000"（常に） | |
| A22, A23 | 正味M3, 風袋M3 | "0000"（常に） | |
| A24 | 梱包コード | "00000"（常に） | |
| A25 | EXｹｰｽNo | "     "（半角スペース5つ） | 5桁ゼロ埋め |
| A26-A29 | Ex風袋1-3, EXGW | "0000" | 4桁 |
| A30 | 積み合せコード | " " | そのまま |
| A31 | 積み合せNo | 10スペース | Left(buf & 9スペース, 10)（左詰め） |
| A32 | ｺｲﾙ長さ | "00000"（常に） | |
| A33-A35 | ｺｲﾙ内径/肉厚/本数 | "000"（常に） | |
| A36 | ｺｲﾙ巻数 | "0000"（常に） | |
| A37-A39 | EX外寸 | "00000" | 5桁 |
| A40 | × | "000000"（常に） | |
| line | LinConv | | L-1=05、LVC=03、HVC=04、ﾊﾞﾗﾝｻｰ=01、ﾄｯﾄ=01、LS4=06、NS1=07、Test=AA、それ以外="Maru"（4文字で桁が崩れる） |

- ⚠ Right(…,n) なので、桁あふれした値は**左側が切り捨てられる**（例：枚数 123456 → "23456"）。
- ⚠ StCon の CStr(Double)（V5）：`1.005×1000` は "1005"。小数が残る値（厚 0.2855→285.5）は "00285.5" になる。
- HostDataConfi：`Mid(buf,13,8)` を引当No、`Mid(buf,21,2)` を梱包No として表示する（A3 が8桁である前提）。

---

## 4. ハードコードされた業務ルール一覧

| 区分 | 内容 | 根拠の行 |
|---|---|---|
| 色 | OK=13434828（RGB 204,255,204）、NG=8421631（RGB 255,128,128）、白=16777215、選択=RGB(0,255,255)、黄色マーク=RGB(255,255,0)＋赤枠 | J:137-138, 2445-2448 |
| 直角度公差 | **1.0 固定**（Q-S-QA-DIM-154） | J:1904, 1092-1097 |
| 直角度欄の表示 | 用途コード **R115** のときだけ | O:231, 266, UF1:3693 |
| ラテラルボー | DbReturnItem の結果に "ラテラルボー"、かつ Additional の右4文字が "検査不要" でない | O:200, 208 |
| 0.5 刻み | 0.1-0.5→0.5 … 4.51-5.0→5.0、5.01-999 はそのまま（Q-S-QA-ALL-153） | J:2428-2441 |
| 公差等級 | "4"=厳しい、"3"=緩い。重なりは管理番号の小さい方 | J:1117-1135, 1150 |
| 異常値 | 公差の2.0倍、対角は √(巾²+丈²) ± 6 | O:1328-1347 |
| 歪の形状 | 厚 ≤ 6.0mm なら 幅/丈/耳/中、超えれば 幅/丈。中：Ⅰ/Ⅱ/Ⅲ型、耳：Ⅱ/Ⅲ型。フリーソリならゲージは 上/下 | O:1447-1501 |
| 平坦度 | 35kg 以下はパターンA（両面定盤）、超えればパターンB（ゲージ）、比重 2.68 | O:1842-1983, 816 |
| 先行データ票 | 製造板厚 ≤ 3.0 かつ 表面処理が S/T 以外、または 表面処理が S/T 以外 かつ 用途コードの頭が T/D1/D2 | J:354-359 |
| 導電率 | 75S:T651/T6、14S:T6/T651（QA-021-08） | O:2178-2180 |
| BH の NG メッセージ | 1か所目 NG：作業中止・保留・ロット異常連絡票発行。最終位置 NG：梱包完了させてAインプット・保留・ロット異常連絡票発行 | J:471-472 |
| 少枚数の検査位置 | 5回：U/D は厚のみ、ラテラルボーは S/C/E（2025.7、Q-S-CH-DIM-163）。4回：E 行なし、U は厚のみ、ラテラルボーは S/C/D。3回：D/E なし。2回：C/D/E なし。1回：S のみ | O:249-514 |
| 高さチェック | 合紙 0.029（通常）、0.042（ストア）、外装紙 0.119×3/包、VC の初期値 0.08406977、許容 +3/4×最大厚、−1/3×最小厚、高精度用途 H191/H242/H241 | I:476-555, 796 |
| 校正 1000mm | 999.5 ≤ x ≤ 1000.5（Test ラインは 50～3000） | P:281-287 |
| 校正ランクの中心値 | L-1:2053.2、LVC:2000.5、HVC:2003.8、ﾊﾞﾗﾝｻｰ/ﾄｯﾄ:2051.8（以上と未満で分ける） | P:417-465 |
| 品証ランク | 1000 以上と未満で分ける | P:558-568 |
| 校正の案内文 | L-1/LVC/HVC は「P2:1000mm直尺の全長」、大板（ﾊﾞﾗﾝｻｰ/ﾄｯﾄ）、Test。合格後：L-1=直尺「S-2」、LVC/HVC/大板=「S-1」 | P:11-16, 326-332 |
| 伝送フォーマット | "DT…"→Mid(14,6)、"#L…"→Mid(3)。7文字以上 | UF校正:24-47 |
| ライン → ログフォルダ | L-1、LVC、HVC、ﾊﾞﾗﾝｻｰ→バランサー、ﾄｯﾄ→トット、Test、それ以外→丸徳 | J:449-457 ほか多数 |
| ライン → 履歴の LN | L-1、LVC、HVC、Bala/Tott→「大板」、Maru→「中板」、それ以外→Test | I:1036-1044 |
| ライン → NIS コード | §3.20 | I:1320-1333 |
| 電文の固定値 | "KEN:AIN:"、"0185" | I:1233-1234 |
| ホスト結果コード | A000 正常、A001 テキスト不正（品質保証課）、A002 送信内容不正（NIS）、A003 登録失敗（NIS） | I:1373-1392 |
| 削除理由 | 入力位置間違い/測定位置間違い/測定ボタン押し間違い/再測定/0点確認/板廃棄/その他 | O:1618 |
| 山数 | "1"～"6"、"枚数"（UFAinput の山数は 1～15） | O:1584, I:7 |
| 不良要因 | 「不良要因印刷原紙」の行2（名前）と行9（コード）＋「ｽﾎﾟｯﾄ処理」。TP は22列目固定 | O:1593-1606, 2431 |
| 熱処理 | ANI/ANF/CAI/CAF | O:2501 |
| 保存パスワード | xlsx：`Lot & "." & yyyy.mm.dd`（履歴保存）、`Ken & "." & yyyy.mm.dd`（A保存） | I:1104, 976 |
| 連番の上限 | Ainput CSV は .02～.99（99 を超えると End）、履歴 xlsx は .002～.999 | I:908-917, 1107-1117 |
| 包装仕様の旧ハードコード | 1P0113 等 → DB 化済み（コメントのみ） | J:2124-2183 |

---

## 5. 外部依存

### 5.1 パス（初期設定.bas の Setting_Path。オンライン時の値）
- `MainPath` = `\\nlmsrvngy03\Read\【New】仕掛\台帳\`
- `NextPath` = `\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\仕掛\`
- `QualityPath` = `\\nlmsrvngy03\工場内共有\検査データ\品証-梱包やり取りフォルダ\Accdb\`
- `PackingPath` = `\\nlmfangyshrd\各課共有\0130_日軽稲沢\梱包課\AIM\【■】_参照用ファイル\`（サーバー名が他と違う）
- `HistoryPath` = `\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\検査履歴\`
- `FactoryyPath` = `\\nlmsrvngy03\工場内共有\検査データ\`
- `NisPath` = `C:\Tools\測定データ電子化`
- オフライン時は全部 `ThisWorkbook.path & "\"`

### 5.2 Access DB（ACE.OLEDB.12.0、`SELECT * FROM 表` ＋ Recordset.Filter）
| ファイル | 場所 | テーブル | 使用箇所 | Filter / 列 |
|---|---|---|---|---|
| 品質基準.accdb | QualityPath | 製品規格直角度 | EnsureLoaded | 全件。管理番号, 等級, 巾_下限, 巾_上限, 丈_下限, 丈_上限, 公差 |
| 〃 | 〃 | 製品規格ラテラルボー | 〃 | ＋厚_下限, 厚_上限 |
| 〃 | 〃 | 欠陥名 | Visual_Add | 全件。欠陥名 |
| 〃 | 〃 | Q_S_KE_ALL_053 | BhSearch（担当外） | `[材質] LIKE '*xxx*'`。材質, 調質, バーコール下限値 |
| 梱包資材マスタ.accdb | PackingPath | 注意_包装仕様 | PackagingSpecificationNo | `[包装仕様NO]='…'`。用途コード, 包装仕様NO, コメント, 備考, 受注材質, 受注調質, 厚下, 厚上, 板幅下, 板幅上, 板丈下, 板丈上, フラグ, 形状, 納入先, 取引先 |
| SIKAODRNOW.accdb | MainPath→NextPath | 仕掛 | MaterialsMsgBox | `[受注番号]='…'` |
| SIKALOTNOW.accdb | Arrbox | 仕掛 | AdvanceCheck | `[ﾛｯﾄ番号]='…'`。品質ｸﾞﾚｰﾄﾞ_表面処理, 製造板厚, 用途ｺｰﾄﾞ |
| SIKACOMP.accdb | NextPath | 仕掛 | No_Search | `[ﾛｯﾄ番号]='…'` |
| SIKATWOYEAR.accdb | NextPath | 仕掛 | Ainput_Startup | `[ﾛｯﾄ番号]='…'`。製造材質/調質/板厚/板幅/板丈, 検査番号 |
| 伝送用ファイル.accdb | QualityPath（Test のときは PackingPath） | VC厚 | VcConv | `[品名ｺｰﾄﾞ]='…'`。厚-係数 |
| 〃 | QualityPath→PackingPath | 校正ストレートエッヂ | StraightEdge（担当外） | `[ライン名]='設備名'`。長さ下（先頭行）、長さ上（最終行） |
| 〃 | 〃 | 時間用 | ShiftTime（担当外） | 開始, 終了, 直 → Start1-3/End1-3 |

⚠ Filter の値に `'` が入るとエラーになる（エスケープしていない）。

### 5.3 Excel ブック
- `QualityPath\75S・14S焼入れチェック\QF空冷時間チェックリスト.xlsx`（Sheet1、読み取り専用、別インスタンスで開く）
- `PackingPath\導電率測定結果記入表.xlsx`（シート「導電率測定」を自分のブックにコピー）

### 5.4 本ブックのシート
| シート | 読/書 | 箇所 |
|---|---|---|
| 校正記録 | 読み書き | 校正, 測定完了, 校正ランク付け（I:O 列をクリアして書く） |
| 品証検証 | 読み書き | 品証測定, 品証ランク付け（R:X 列をクリアして書く。P 列の日付に "2019." を足すデータ修正もする）、UF校正 CommandButton1 |
| 作業保存_<Lname> | 読み・消去 | NG_Erase（行2-7 の A:T、行10-15 の A:C、行8-13 の O:P）、WithorWithout_Gauge / GaugeJudgment（O2:P13）、UF異常値発見（1行目） |
| 異常ロット用_<Lname> | 書き | NG_Posting（A1=Lot、各行 1-18 列＋19 列=位置。NG セルは色付き。フォント メイリオ 12） |
| 不良要因印刷原紙 / 不良要因印刷_<Lname> | 読み書き | Visual_Add, SelectionInput（B1=Lot、B3:B8=位置、C3:X8=枚数、C11=スポット枚数） |
| Lot情報履歴 | 読み | SmallNumberCount（GetLotInfo）, No_Search, Ainput_Startup, Height_Check |
| Hiki_Lot | 書き | No_Search |
| WorkerTime_ / Daily_ / 梱包枚数計算_<Lname> | 読み | NameList_Open |
| 検査員_<Lname> | 読み | ConductivLoad |
| 導電率測定 | 書き | ConductivLoad（B4 日付, B5 Lot, B6 材質-調質, B7 厚×巾×丈, B8 ライン, B9 検査員） |
| Nitiina_ / A_Icas_<Lname> | 一時 | A_Output（作ったあと削除） |
| 印刷用_<Lname> | 読み | 履歴保存（C2 = Lot） |
| 高さチェック、ﾛｼﾞ、耳用、パレット種類_外径条件 | **担当コードからは参照なし**（高さチェックは手書きの検証データと思われ、テスト値として使える。§10） | |

### 5.5 ファイル入出力
- ログ：`WriteLog(msg, "", HistoryPath & "<ライン>\<ライン>_Log", "Measured.log")` と `WriteLog(msg, "")`（既定の出力先。WriteLog モジュールを参照）。判定ログの書式は `*UFKen|値|(セル名)|○|Lname|検査員`（2026.05.14 に ":" から "|" に変更）。⚠ LcColor のログは `"*UFKenl" & TB & "l" & …`（"|" ではなく小文字の L。O:1701-1702）。
- `C:\Tools\測定データ電子化\hostsend.txt`（ExpCSV。Print # で書くので ANSI（CP932）、末尾は CRLF）
- `C:\Tools\測定データ電子化\result.txt`（Line Input で最終行を読む。ANSI）
- `C:\Tools\測定データ電子化\SokuteiDataTrans\SokuteiDataTrans.exe`（Shell で起動）
- `HistoryPath\<ライン>\<ライン>_Ainput\yyyy.mm\<Lot>\yyyymmdd_<検査No>[.nn].csv`（"," 区切り、行末にも ","、引用符なし、CP932）
- `C:\検査履歴\<LN>\yyyy.mm\dd日\` と `FactoryyPath\【梱包】<LN>\…` に PDF。`C:\検査履歴\yyyy.mm\yyyymmdd_Lot\yyyymmdd_Lot[.nnn].xlsx`（パスワード付き・読み取り専用）→ `HistoryPath\<ライン>\<ライン>_履歴\yyyy.mm\` にコピー
- 包装仕様書の画像 `QualityPath\包装仕様書\<包装仕様NO>.JPG`

### 5.6 レジストリ・音声・その他
- `GetSetting("UF検査判定システム","起動設定","設備名","存在なし")`：測定完了（校正記録 G 列、ストレートエッヂ検索のキー）、TimeCheck/YeaDayrCheck（L-1/LVC/HVC/Test 以外は「日勤」）
- 音声：音楽合格、音楽不合格、音楽対角、音楽異常値、`音楽を流す("必要です.wav"/"直角度.wav")`
- Win32 API：Sleep、SetWindowPos、GetWindowRect、GetDC、GetDeviceCaps
- `Range.Speak`（UFcount の読み上げ）

---

## 6. 入力データ / 出力データ（項目レベル）

### 入力
- 画面（UF1）：S1-S14…F1-F10（厚・巾・丈・対角）、Lateral?、Diag?Kou、Z?/G?（歪）、CB?1-4（歪形状）、?BH、?_UpVisual/?_UnderVisual（外観）、KenHBox1-7（検入高さ）、Yama（山数）、T1-T3（検査員）、作業選択（薄板/厚板/極厚/少枚数/スラブ）、Abno（異常値の on/off）
- 規格（UF1 にセット済みの値）：ATU/Hab/Tak、Am/Ap/Hm/Hp/Tm/Tp、ApK/AmK/HpK/HmK/TpK/TmK、TaikText、LateralText、HaHiz/TaHiz/Naka/Mimi/HURIS、flatnesText、Zai/Tyu/Youto、QF、Additional、Sumisei
- UFAinput：HikiNo, KONNO, Remain, MAI1/2, COU1/2, Material1-3, GW, PileMatchNo/PileMatch, ExCaseNo, ExMaterial1-3, ExGW, OutsideAtu/Haba/Take, AMin/AMax/HMin/HMax/TMin/TMax, KenLaw, LineName, DynHL_1-15（山高さ）, GaityuuText, Ai/Ga/Vab（外注の除外チェック）, Youto, OrderNo
- UF校正：INP1000（1000mm）、INPG（ストレートエッヂ）、Kname、LineName
- 測定器：伝送文字列（"DT…"、"#L…"）

### 出力
- 画面の色、音、画像、UFMsgBox、UFConfi
- Measured.log（判定ごと）
- 校正記録：A 日付(yyyy.mm.dd)、B "夜勤日付処理"、C 時刻 "h時mm分ss秒"、D 直（合格時のみ）、E 測定値、F 検査員、G ライン（設備名）、H "〆"（不合格時）
- 品証検証：C 値、I ライン、J 検査員、K 時刻（合格時は該当日・直の行）。不合格時は M 値、N ライン、O 検査員、P 時刻を末尾に追加（赤字）
- 異常ロット用_<Lname>：1-14 測定値（F は 1-10）、15 直角度、16 ラテラルボー、17 定盤歪（"A:"/"B:"＋値）、18 ゲージ歪（"ｹﾞ:"＋値）、19 位置
- 不良要因印刷_<Lname>、導電率測定、Hiki_Lot
- hostsend.txt（固定長電文）、Ainput CSV（56行：項目名,値,）、PDF、xlsx

---

## 7. 校正（Proofreading / UF校正）の業務

### 7.1 何を校正するか
1. **1000mm 直尺**（品証用の「伝送コンベックス測定実績」。デジタル表示とスケール読み取りのばらつき調査）：合格範囲は 999.5 ≤ x ≤ 1000.5（P:285-296）。
2. **ライン別ストレートエッヂの全長**（例 L-1 = 2053.2±0.5 と旧コメントにある。現在は DB「校正ストレートエッヂ」の「長さ下」「長さ上」を使う。キーはレジストリの設備名。P:92-108）。合格範囲は `Val(SokuteiH) >= Koumin And <= Koumax`（**両端を含む**）。
   - SafeNum（2026.06.25）：DB の値にカンマ（桁区切り）があれば除去。小数点が2個以上、または数値でなければ失敗 → 「校正基準値の設定エラーです。管理者に連絡してください。」を出して終了（記録しない＝未校正のまま）。
   - StraightEdge が空（設備名が DB にない）なら `IsEmptyArray` で何も言わずに終了する（推測：IsEmptyArray("") の結果しだい）。

### 7.2 画面の流れ（UF校正）
- Initialize：背景は緑（&H80FF80）＝1000mm モード。`F1=1`。品証ランク付けを表示。
- INP1000_AfterUpdate（F1="1" のとき）→ TerConv → SokuteiH → ログ "1000mm測定:" → 品証測定
  - 合格：品証検証の当日（"d日"）・当直（日勤は "1直" に読み替え）の行に記録。背景を**白に変える**（＝ストレートエッヂ モード）。校正ランク付け、INPG にフォーカス。
  - 不合格：音を鳴らし、「999.5 ≦測定値≦ 1000.5 になっていません」、品証検証の末尾に赤字で記録。
- INPG_Enter：背景が緑のままなら INPG をクリアして INP1000 にフォーカス（**1000mm が合格するまでストレートエッヂを入力できない**）。
- INPG_AfterUpdate → TerConv → 測定完了 → ログ "ｹﾞｰｼﾞ測定:"
  - 合格：校正記録に追加。`F1="end"`、`UF1.ProofreadFlag="ok"`。
  - 不合格：校正記録に D 列（直）なし・H 列 "〆" で追加。NG 値を表示し、再検査。
- INPG_Exit（F1="end"）：UF1 を表示し、LoopJudg と HizLoopJudg で再判定してから Unload。
- ⚠ 抜け道：CommandButton5（背景を白にするだけ。1000mm 測定をとばせる）、`Ltrue_DblClick`（Excel を表示して Unload。管理者用と推測）、QueryClose の × は無効。
- TerConv：TerMaker で先頭2文字を UCase と半角化して判別。"DT" → `Mid(Ter,14,6)`、"#L" → `Mid(Ter,3)`、それ以外は「正しいデータが伝送できていません」で ""。7文字未満は ""。

### 7.3 どこに記録するか
§6 の「校正記録」「品証検証」を参照。合格の日付は `Format$(YeaDayrCheck(…),"yyyy.mm.dd")`。0:00～End3 なら B 列に "夜勤日付処理"。

### 7.4 未校正の判定とブロック
- `校正(STtime)`（P:28-61）：
  1. `UF1.WorkButton.Tag="jump"` または `UF1.DBflag<>""`（テストモード）なら False（済み扱い）。Tag はクリアする。
  2. 校正記録シートがなければ作って True。
  3. **A 列の最終行**で比べる：`A(最終) <> YeaDayrCheck(STtime)` なら True、`D(最終) <> TimeCheck(STtime)` なら True。どちらも一致すれば `ProofreadFlag="ok"`（関数の先頭で "" にしている）。
  4. ⚠ 不合格の行は D 列が空なので、**最後が不合格なら未校正**になる（再校正が必要）。
  5. ⚠ 比べるのは**最終行の1行だけ**。同じ直に別ラインの行があっても関係ない（1ブック＝1ライン PC の前提）。
- 呼び出されるタイミング：UF1 の初期化（UF1:4296。済みなら Flag=ok）、検査員の登録（T1/T2/T3_Change：Day? の日付が当日なら `校正()`。未校正なら UF校正 を起動。UF1:4099-4159）。
- **ブロック**：測定値の入力処理（UF1:4600-4627）で `UF1.ProofreadFlag <> "ok"`（jump/DBflag のときを除く）なら `Proofread_Load`（「校正が済んでいません。校正を起動します。」）→ 入力値を消して終了。
- ⚠ **日付・直が変わったとき**：ProofreadFlag は、UF1 の初期化、`校正()` の呼び出し（検査員登録時）、測定完了 でしか更新されない。**UF1 を開いたまま直が変わり、検査員も登録しなければ、Flag は "ok" のままでブロックされない**（推測。要確認）。

### 7.5 日付と直の判定（Worker モジュール。担当外だが校正の結果を左右する）
- Start1-3、End1-3 は DB「時間用」から取る（Set_Shift）。
- TimeCheck："1直" は `Start1 < t < End1`（両端を含まない）、"2直" は `Start2 ≤ t < End2`、それ以外は全部 "3直"。設備名が L-1/LVC/HVC/Test 以外なら「日勤」。
- YeaDayrCheck：`00:00 ≤ t ≤ End3` なら前日の日付、それ以外は当日（yyyy.mm.dd）。設備名が上記以外なら常に当日。
- `校正()` は "hh:mm"（分単位）で判定し、測定完了の夜勤マークは `Time <= CDate(End3)`（秒単位）で判定する。

### 7.6 ランク付け
- 校正ランク付け：E 列を `Format(Val,"0.0")` で書き直す（⚠ データを書き換える）。当月（A 列の先頭7文字 = yyyy.mm）の最初の行から、**A 列のフォント色が RGB(255,128,128) の行**をライン中心値の以上と未満に分けて I:K 列と M:O 列へ写す。Max/Large/Min/Small で上位と下位3位を Rank1-6 に表示する。
  ⚠ 現行の 測定完了 は**フォント色を設定していない**（旧コードではしていた P:243-247）→ **新しい記録はランクに出ない**（死んでいる可能性あり）。当月の行がないと Tuki="" になり、`For C = "" To j` でエラー → `On Error Resume Next` で先に進む。
- 品証ランク付け：品証検証の M 列（不合格値）を 1000 以上と未満に分ける。表示の書式は同様。

---

## 8. エラー処理の方式
- 大半は**エラー処理なし**（実行時エラーがそのまま出る）：HizumiJudg、LoopJudg、Out_Tolerance、Height_Check、AdvanceCheck、PackagingSpecificationNo、Distortion_Check など。
- `On Error Resume Next`：JudgColor（手続き全体。kou の代入失敗が NG 判定につながる）、CBAdd、GetLabelState、GetKHeight 系、Loadtext、CreateFolder、NameList_Open（途中から）、LateralbowJudgColor/RightAngleJudgColor（MPEX.Visible の前から最後まで）、校正ランク付け/品証ランク付け、UFdiagonal、UF校正 の Layout/Activate/Initialize。
- `On Error GoTo ErrHandler`：EnsureLoaded（MsgBox を出し、`m_LoadFailed=True` で以後は再試行しない → 公差が0 → 公差なしとして扱われる）、Visual_Add（どの段階かを MsgBox）、Test_Tolerance。
- 業務エラーは MsgBox か MsgError で知らせ、`Exit Sub` で終わる。`End`（VBA 全体の停止）を使う箇所：連番の上限（I:915, 930, 1115）、品証測定で品証検証シートがないとき（P:260）、adoRecordset で接続できないとき（SQL_追加:299）。
- 処理停止フラグ：`UFAinput.Ngflag="stop"`（App_Activation, InterRuptText）。

---

## 9. グローバル変数・状態、日付をまたぐときの挙動
| 状態 | 場所 | 内容・寿命 |
|---|---|---|
| m_Diag, m_Lat, m_DiagIdx, m_LatIdx, m_Loaded, m_LoadFailed | Judg モジュール | 公差表のキャッシュ。**Excel を閉じるか ReloadTolerance を実行するまで更新しない**。1度失敗すると、以後の公差はすべて0（「公差なし」）。DB を修正しても再起動までは反映されない |
| MAX_MOUNT=15 | Input（Const） | 山数の上限 |
| g_FormHwnd | FormRelat | ResizeWindow で使う |
| Start1-3/End1-3 | Worker | 直の時刻（DB） |
| g_lngWindowState（Proofreading, UF異常値発見）、g_sng*（UF校正）、i, a（UF校正のモジュール変数） | | 使っていない |
| UF1.ProofreadFlag | UF1 コントロール | 校正済みフラグ（§7.4） |
| UF1.MoveFlag | | "on" = MouseDown による入力（音を鳴らす） |
| UF1.Abno.Tag | | "NG" = 異常値フォーム表示中（TextJudg を止める） |
| UF1.MPEX.Tag | | "Click" = ページクリック中（画像表示の方法を切り替える） |
| UF1.INP.Tag | | 現在のフォーカス位置（セル名） |
| UF1.Tag="STOP" | | 音を止める |
| UF1.Load.Caption / UF1.Load | | "off" = フォームイベントを止める |
| UF1.NGS～NGF | | 「許容」「保留」 |
| UF1.DeleteC | | 削除回数 |
| UF1.WorkButton.Tag="jump"、UF1.DBflag | | 校正をとばす（テスト） |
| UF校正.F1 | | 1=入力待ち、0=処理中、2=品証 NG、"end"=完了 |
| UFAinput.Tag, Ngflag | | Aインプットの状態 |

日付をまたぐときの挙動
- 校正：§7.4、7.5。夜勤の 0:00～End3 は前日の日付・3直として扱う。
- `A_CsvSave` / `履歴保存` / `SaveDateLocate` のフォルダ名とファイル名は `Format(Date, …)`（**暦日**）を使い、夜勤の前日扱いをしない（YeaDayrCheck を使っていない）。夜勤の 0 時以降に作ったファイルは翌日の日付になる。
- `Conductiv…`、`ConductivLoad` の日付：`Format$(Now,"yyyy/mm/dd")`（暦日）。
- 品証測定：`Formay = Format$(Now,"d") & "日"`（**暦日**）と `x = TimeCheck`（3直）を組み合わせる。0:00～End3 の3直は暦日の翌日の行に書かれる（推測：YeaDayrCheck を使う校正記録とは日付がずれる）。
- 校正ランク付けの当月：`Format$(Now,"yyyy.mm")`（暦月）。

---

## 10. Python 移植時のテストケース候補

### 10.1 AHTJudg（ApK=25.450、AmK=24.550 のとき。作業保存_Test の値）
| 入力 | 期待 |
|---|---|
| 24.55 | OK（下限ちょうど） |
| 24.549 | NG |
| 25.45 | OK（上限ちょうど） |
| 25.451 | NG |
| "abc"、"" | False（NG） |
| Am="" のとき JudgColor "A" | 判定しない（色を変えない） |

### 10.2 対角（TaikText="1.5"）
| S9 | S10 | 期待 |
|---|---|---|
| 3145.9 | 3145.9 | Tai=0 → OK |
| 3145.9 | 3147.4 | Tai=1.5 → OK（等しい） |
| 3145.9 | 3147.5 | Tai=1.6 → NG |
| 2947.3 | 2946.4 | Tai=0.9（Double では 0.900000000000091、Single に変換すると 0.9）→ OK。Diag 表示は "0.9" |
| TaikText="" | | NG（kou=0） |
| TaikText="0.0" | 差 0 | DiagonalJudg は kou=0 で False → NG |

### 10.3 LateralJudg / UnitRecord（LateralText="2.0"）
| 入力 | UnitRecord | 判定 |
|---|---|---|
| 0 | 0（-0 に注意） | OK |
| 0.05 | 0 | OK |
| 0.3 | 0.5 | OK |
| 1.6 | 2.0 | OK |
| 2.1 | 2.5 | NG |
| -2.1 | -2.5 | NG（絶対値で判定） |
| -1.9 | -2.0 | OK |
| .51 | **0**（Single の境界。§3.5） | OK |
| 1.01 / 1.51 / 2.51 / 3.51 | **0** | OK |
| 4.01 | 4.5 | |
| 5.3 | 5.3 | |
| 1000 | 0（999 超） | |
| LateralText="0" | | NG（kou="0"） |
| LateralText="" | | LateralbowJudgColor は「公差がありません」 |
| 直角度 1.0 / 1.1→1.5 | | OK / NG |

### 10.4 BH（下限値 bh=40 のとき）
| 入力 | 期待 |
|---|---|
| 39 | NG |
| 40 | OK |
| "" | 不要（白） |
| 39.5 → Format "0" で 40 | OK（四捨五入） |
| 40.5 → "41" | OK |
| bh=0（DB に該当なし） | 不要 |
| DB 下限値 "35.5" → CLng 36 / "34.5" → 34 | 銀行丸めの確認 |

### 10.5 公差表の参照
- Test_Tolerance の4ケース（§3.8）。
- 巾 300 / 300.05 / 300.1 / 1500.1（Single）/ 丈 3700 / 3700.5 / 3701 / 7000 / 7000.1（該当なし → 0）。
- 同じ区分に重なる2行（管理番号 5 と 3）→ 3 の公差を使う。
- 等級が数値型か文字型か（CStr(…)="4"）。DB の下限/上限が Null の行は除く。
- DB に接続できないとき：MsgBox は1回だけ、以後の公差は 0。

### 10.6 歪（HaHiz="1.0"、TaHiz="2.0"、CBS1="幅"）
| 入力 ZS | 期待 |
|---|---|
| 1.0 | OK |
| 1.01 | NG |
| "0.8/3" | OK（"/" より前） |
| "abc" | 何もしない |
| CBS1="" | 何もしない |
| GS=0.5、CBS1="幅" | 白（ゲージは判定しない） |
| フリーソリ：HURIS="3"、flatnesText="ﾌﾘｰｿﾘ要"、GS=3 / 3.1 | OK / NG（CBS1 は空でもよい） |

### 10.7 重量と平坦度
- Weight(25, 1385, 2770.5) = RoundDown(25×1385×2770.5×2.68/1e6, 3)。Python と Excel ROUNDDOWN の結果が同じか確認する（float の floor 誤差）。
- 35kg ちょうどの例：Weight(1, 1000, 13059.70) ≈ 34.999…（境界付近の値を作って ≤35 と >35 を確認）。
- ≤35kg で A のみ → "?のB面 定盤 未入力_ﾊﾟﾀｰﾝA"。>35kg で S が A+B 両方、U が A のみ → Exit Function（"" または前のループのメッセージ）。C 行のメッセージが "2の…" になること。F 行は B ゲージを判定しないこと。

### 10.8 異常値（作業保存_Test：ATU=25、Am=Ap=0.45、HABA=1385、Hm=Hp=2、TAKE=2770.5、Tm=Tp=2.5）
- 厚：24.1 は正常（境界 24.100）、24.099 は異常。25.9 は正常、25.901 は異常。
- 巾：1381.0 は正常、1380.9 は異常。丈：2765.5 / 2775.5 が境界。
- 対角：√(1385²+2770.5²)=3097.401… → 3091.4 / 3103.4 が境界（Format "0.0"）。

### 10.9 高さチェック
- 例A（合紙なし、VC なし、個装なし、外注でない）：AMin=1.5、AMax=2.0、Mai=10 → GP=20、GM=15、ErrorA=1.5、ErrorB=0.5。
  - Hei=21.5 → OK（差1.50 > 1.5 は偽）。21.51 → NG「許容値ｵｰﾊﾞｰ」。21.504 → "1.50" で OK。21.506 → "1.51" で NG。
  - Hei=14.5 → OK。14.49 → NG「枚数不足」。
- 例B（合紙あり 列17=1、Mai=100、AMin=2.95、AMax=3.05）：Ais=100×0.029=2.9（現行。旧版なら 2.929）。GP=307.9、GM=297.9、ErrorA=2.2875、ErrorB=0.98333。Hei 310.1 → OK、310.2 → NG。297.0 → OK、296.9 → NG。
- 例C：COU1=3、COU2=""、合紙あり → Ais=(Mai+3)×0.029、Gai=3×0.357=1.071。
- 外注：Ai=True / False × 合紙あり/なし の4通り（現行の結果を正とする。旧版との差を記録しておく）。
- VC：VC=1、Remain=1、Youto=H191 → 加算される。Youto=H174 → 加算されない（VCAB も設定しない）。
- VC が DB に登録なし → 0.08406977。
- OrderNo が見つからない → False（合格扱い）＋MsgBox。
- `高さチェック` シート（MAX/MIN/枚/高さ/VC/合紙/判定）の実績値（例：2行目 2.001/1.979/100枚/210mm/B535+WVC/合紙無）を、VC 厚を DB から取った上で回帰テストに使える（H列「×」、I列「補正変更後〇」の意味は §12 で確認）。

### 10.10 電文と CSV
- StCon(1.005,"厚")="0001005"、StCon(1102.3,"巾")="0110230"、StCon(3053.3,"丈")="0030533"、StCon(0,…)=""、StCon(0.2855,"厚")="00285.5"。
- ZerCon(31,"AB")="AB        "（10桁）、ZerCon(25,"")="     "、ZerCon(12,"123456")="23456"。
- LinConv("丸徳")="Maru"（4文字）。
- A_CsvSave：同じ名前のファイルがあれば .02 を付ける。.99 まであれば End。CSV の各行の末尾に "," が付くこと。CP932 で書くこと。

### 10.11 校正
- ストレートエッヂ：下限 2052.7、上限 2053.7 → 2052.7 / 2053.7 は合格、2052.6 / 2053.8 は不合格。
- DB の値 "2,052.7" → 2052.7。"2053..5" → 設定エラー。"" → 設定エラー。
- 伝送文字列："DT10002+0000021.712M" → "21.712"、"#L395.8" → "395.8"、"ｄｔ…"（全角小文字）→ DT と判別、"ABC" → "" ＋警告。
- 校正() の判定：最終行が（当日、当直）→ 済み。（当日、他の直）→ 未校正。最終行が不合格行（D 空）→ 未校正。夜勤 0:30 の最終行が（前日の日付、3直）→ 済み。設備名がバランサー → 「日勤」、当日の日付。
- 1000mm：999.5 / 1000.5 は合格、999.4 は不合格。Test ラインは 50～3000。

### 10.12 注意文（PackagingSpecificationNo）
- 1件・フラグ "" → "板_コメント"。1件・寸法の条件に合わない → "板_" → **先行データ票の案内も出ない**。
- 2件以上で "寸法納入" を含むとき、前のコメントが上書きで消えること。
- 納入先の条件 "≠ﾆﾎﾝﾊﾂｼﾞﾖｳ*" と 実値 "ﾆﾎﾝﾊﾂｼﾞﾖｳ(ｶ" → False。"＊" → True。完全一致の大文字小文字を区別しないこと、前方一致は区別すること。
- コメントに "_" を含む → 最後の部分だけ表示されること。

---

## 11. 未使用・死んでいそうなコード候補（削除してよいとは断定しない）
| 対象 | 根拠 |
|---|---|
| Judg.NflatRewrite | 呼び出し元が見つからない |
| Judg.Test_Tolerance | 呼び出し元なし（「確認後は削除可」とコメント J:1777） |
| Judg.ReloadTolerance | 呼び出し元なし（手動で実行する前提。移植では管理機能にする候補） |
| Judg 旧公差表 J:1276-1775、旧注意喚起 J:2124-2183、HizumiJudg のゲージ判定 J:262-284 | コメントアウト済み |
| OutTolerance.BH_Check | 唯一の呼び出しが UF1:3878 でコメント |
| OutTolerance.SingleUnitDistortion | 呼び出し元なし |
| OutTolerance.WithorWithout_Gauge と GaugeJudgment の重複 | 用途が違う（印刷前チェックと色表示）。メッセージ文も違うので統合には注意 |
| OutTolerance.ConductivLoad | UF1 から1か所呼ばれる。SaveLoad 側の呼び出しは「2022.07_一旦停止」 |
| SmallNumberCount の `直角flag`（O:201） | 代入するだけで使っていない |
| Input.A保存 | 呼び出し元なし（コメント「CSVで保存しているから使ってない」） |
| Input.HostCon、Input.Array2D | 呼び出し元なし |
| Input の旧 Height_Check（I:592-751） | コメントアウト済み（ただし旧仕様の手がかりとして重要） |
| Proofreading の g_lngWindowState、UF校正 の g_sng* | 使っていない |
| 校正ランク付けの赤字抽出 | 現行の 測定完了 がフォント色を付けないので、新しい記録は抽出されない（推測） |
| UF校正.CommandButton1_Click | 品証検証に記録する別ルート（旧 UI と推測）。`UF校正.Show vbModal` で同じフォームを再表示する |
| UFdiagonal | 呼び出しがコメントアウト済み（J:642-648） |
| UFOK / UFNG / UF対角 / UFcount | **死んでいない**。キャスク検査（`Cask.bas` 1326-1571, 1886-1904）と全ヵ所検査シート（`ZenSh2.cls` 80-99）のシート UI から呼ばれている。UF1（板の通常検査）からは使われておらず、UF1 では Judg_ImageSize の画像表示がその役割を担う。UF対角 の `ws.Name` 分岐（"検査","少枚数検査","極検査"）は旧シート名で、該当シートが今のブックにあるかは未確認 |
| Input.ResizeWindow 系 | 使用中（UFAinput） |

---

## 12. 不明点・要確認事項
1. **【Q1・最重要】数値にできない文字列と数値の比較**：`Select Case Rig`（Rig=""/"S"）と `Case 2,4…`（O:871 ほか）、`.ZS = 0`（ZS が空）（O:1989-1999）、`CDbl(TB) > ComValueG`（ComValueG=""）。実機でエラー13 になるのか、何か規則で比べられるのかを Excel で確認する必要がある。本番で使えているので、エラーにならない何らかの理由がある可能性が高い。Python の再現方法はこの結果で決まる。
2. **Height_Check の旧版と現行版の差**（§3.19 の4点）：現行の計算（個装なしで Mai×0.029、外注 Ai の扱いが逆）が業務として正しいか。2026.05.29 の書き換えで意図せず変わった可能性がある。
3. Out_Tolerance で NG が複数あるとき、UFConfi は最後のセルだけを表示しているのか（実機確認）。移植で「全部出す」に変えてよいか。
4. LoopJudg J:927-928 の D/E の入れ替わり、Distortion_Check O:2022 の ZE/ZF、WithorWithout_Gauge の「2のA面」、F 行を判定しないループ（O:1973）：不具合として直すか、そのまま再現するか。
5. LoopJudg で LateralText が数値でないとき、BH/直角度の再判定がとばされ NG 色が残る仕様（J:1059）は意図どおりか。
6. 対角公差が空のとき、JudgColor は NG（On Error で kou=0）になる。LoopJudg は色を付けない（コーラルが残る）。業務として正しいか。
7. 極厚で TpK="" のとき LoopJudg の 11-14 でエラーになるか（2026.09.18 の耳付き修正の対象外）。
8. UnitRecord の Single の境界（.51 → 0 など）と、5.01-999 の範囲外（999 超 → 0）は意図どおりか。
9. 製品規格の公差表（品質基準.accdb）の境界値の決め方（旧表のすき間、例 3700～3701 をどう埋めたか）と、旧表との一致確認。
10. 校正：UF1 を開いたまま直が変わったときにブロックされないこと（§7.4）。品証測定の暦日と直の組み合わせ（§9）。校正ランクが実質機能していないこと（§7.6）。CommandButton5 で 1000mm をとばせること。
11. 伝送文字列 "DT…" の実際の形式：`Mid(14,6)` はコメントの例（21.712）では正しいが、2000mm 級の値（2053.2）ではどうなるか。高さチェックシートの E89 にある "DT11000+0000001.022M" では "01.022" になる。実機のフォーマットが知りたい。
12. Conductiv_Coolingtime_Check の戻り値 "LotNo記載がない" と、呼び出し側の "LotNoがない" の不一致（担当外のコードを直すかどうか）。
13. MaterialsMsgBox：結果が "形状_" のとき先行データ票の案内が出ないこと、"寸法納入" が上書きされること、コメント中の "_" が切れることは意図どおりか。
14. LinConv の "Maru"（4文字）は丸徳の電文として正しいか（固定長が崩れる）。電文の A3（引当No）と A6-A11 が空のときにゼロ埋めされない。インターフェース定義書（`\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\【4】_Ainput\システム間インターフェース定義.xlsm`）との照合が必要。
15. `高さチェック` シートの H 列「判定」と I 列「補正変更後」の意味（旧ロジックと新ロジックの比較記録か）。
16. 校正記録 A 列の日付 "2026.10.13" を Excel が日付型に自動変換していないか（変換されると CStr の書式が変わり、`TempD <> y` が常に真 → 常に未校正になる）。CSV の C2=0.4195… は、"h時mm分ss秒" の文字列が時刻のシリアル値に変換された跡に見える。
17. Format の .5 境界（例 `Format(2.675,"0.00")`）を VBA が 10進で丸めるのか2進で丸めるのか。上下限（ApK 等）の計算に影響する。
18. 外径・内径：担当モジュールには外径・内径の判定はない（パレット種類_外径条件 は Test.frm からだけ参照されている）。キャスク・三菱など、他の検査系統の担当範囲だと思われる。
19. JudgColor/LateralbowJudgColor は WriteLog を2回呼ぶ（Measured.log と既定の出力先）。既定の出力先はどこか（WriteLog モジュールの担当者に確認）。

---

## 要約
- 判定はすべて「OK=下限≤値≤上限（両端を含む）」「差・絶対値≤公差（等しいとき OK）」で、`Val` と Single を使った比較。
- 直角度は 1.0 固定（R115 のときだけ）。ラテラルボーと対角（直角度の等級）の公差は、品質基準.accdb の表を管理番号最小の優先で引く（キャッシュは Excel を閉じるまで保持、失敗すると公差0）。
- 0.5 刻み変換（UnitRecord）は Single の誤差で .51、1.01 などが 0 になるすき間がある。
- 異常値は公差×2（対角は √(巾²+丈²)±6）。重量は比重2.68で切り捨て、35kg を境に平坦度のパターン A/B。
- 梱包高さチェックは +3/4×最大厚、−1/3×最小厚。現行コードはコメントの旧版と、合紙・外注・外装紙・VC の計算が違う（要業務確認）。
- 校正は、校正記録の最終行の（日付、直）が当日・当直と一致すれば済み。測定入力時に ProofreadFlag でブロックするが、直が変わっても自動では再判定しない。
- DB の値はすべて文字列（GetString）で、Filter は ADO 方式。CStr/Format/Round/Long の丸め方の違いが、移植で結果が食い違う主な原因になる。
- 実機確認が必要な最大の点は、数値にできない文字列と数値の比較（Select Case Rig、.ZS=0）がエラーになるかどうか。
