# G4 構造解析レポート：品質基準・品証DB・管理・補助

対象（行番号は各ファイルの 1 始まり行番号。`Attribute VB_Name` 行 = 1 行目）:

| 略称 | ファイル | 行数 |
|---|---|---|
| Q | `【UFKen_11】Quality.bas.bas` | 2272 |
| QDB | `【UFKen_11】Quality_DB.bas.bas` | 1696 |
| QP | `【UFKen_11】QP_DB.bas.bas` | 2552 |
| FQ | `UF品質基準.frm.bas` | 333 |
| FK | `UF管理.frm.bas` | 682 |
| H | `【UFKen_17】補助.bas.bas` | 1807 |

参照した外部定義: `初期設定.bas.bas`（パス）、`【UFKen_15】SQL_追加`（ADO 読込）、`【UFKen_01】AccHIki`（Arrbox）、`【UFKen_07】Worker`（IsOver12Hours）、`【UFKen_05】OutTolerance`（Weight）、`【UFKen_16】WriteLog`、呼び出し側の `UF1.frm` / `UFMitsubishi.frm`。

> 表記: 「推測」は根拠がコードにないもの。「要確認」は実機・実DBで確かめないと決められないもの。

---

## 0. 移植で一番大事な注意点（先に読む）

1. **「SQLfilter」は SQL の WHERE ではない。** 全部 `SELECT * FROM <テーブル>` を発行し、その後 **ADO の `Recordset.Filter`** で絞り込む（SQL_追加 L293, L296）。LIKE のワイルドカードは `*`。大小文字の扱い、Null と `''` の扱いは ADO Filter の仕様に従う（§5.4, §12）。
2. **取得した値はすべて文字列になる。** `adoRecords` は `GetString(adClipString, …, vbTab, "<#(仮)改行#>")` で読んで Split している（SQL_追加 L428-439）。数値も日付も Yes/No も文字列で、Null は `""` になる。後の処理は `Val` / `CDbl` / `CBool` / Variant 比較で型を変換している。
3. **ファイルが無いと VBA は異常終了する。** `adoConnection` は `PathCheck` が偽なら Nothing を返し（SQL_追加 L201-204）、`adoRecordset` は `.ActiveConnection = adoCn` の後で `If adoCn Is Nothing Then End`（L292-300）。つまり実行時エラーになるか、`End` で全状態がリセットされる（推測: どちらになるかは要確認）。`GetFieldsArr` が Empty を返して `PackingPath` に切り替えるフォールバック（QP L1195 など）は、**ファイルはあるが、テーブルが空か Filter に1件も該当しない**ときだけ働く。
4. **`GetFieldsArr` は 0 件のテーブルに対して Empty を返す**（`If adoRs.EOF Then Exit Function` が先にある。SQL_追加 L462）。「フィールド取得失敗」と「テーブルが空」は区別されない。
5. **行の順序に依存している。** ORDER BY が無いため、「最初に一致した行」「最後に一致した行」が Access の返却順で決まる（§12）。
6. **判定に効かない条件・バグ的な挙動が多数ある**（例: DBQP では送り先コードと巾取数の条件が最終判定に入っていない。QP L328-330）。「VBA と完全に同等」を目標にするなら、バグも再現対象になる。§4 と §14 に全部挙げた。

---

## 1. 各モジュール概要

### 1.1 Quality（Q）— ハードコードされた旧品質基準（現在は呼び出されていない）

- 品質規格書の内容を VBA に直書きした旧版。
  - `品質グレード`（厚板の外観グレード。Q-S-QA-124-R02 別表1/2）
  - `平坦度`（Q-S-QA-122-R05 / Q-S-QA-129）
  - `中歪` / `耳歪`
  - `表面品質基準`（薄板。Q-S-QA-SRF-128）、`PatternJudge`、`GradeJudge`
  - フォーム補助の `BoColor` / `SelectAdd`
- 呼び出し元の UF1 `CommandButton8_Click` 側のコードはコメントアウト済み（UF1.frm L3068-3079、コメントは「2023.4.12_品証管理がしやすいようDB関数作成」）。**現在コードから到達できるのは `BoColor` だけ**（FQ の B1〜B11 から）。
- 入力元は、ローカルの「Lot情報履歴」シート、`UF1` のコントロール、`SIKALOTNOW.accdb`（`表面品質基準`）。

### 1.2 Quality_DB（QDB）— 品証DB（品質基準.accdb）を参照する UF品質基準 表示用ロジック

- `DBdisplay外観_H`: 厚板の外観グレード（Q_S_QA_124_別表1/2）。
- `DBdisplay外観_S`: 薄板の表面パターン（Q_S_QA_SRF_128_別表1/2）。
- `DBdisplay歪`: 歪公差を表示（耳歪/中歪 4級・3級、QA_022_01、Q_S_QA_DIM_122_A/B/C/D、ソリ3級/4級）。**UF品質基準の手動検索ボタン（FQ L122-124）からだけ呼ばれる。**
- 共通のサイズ表引き `LoopArray`。
- `BhSearch`: バーコール硬度の下限値（Q_S_KE_ALL_053）。Judg の `BhJudg` から呼ばれる。
- `DBdisplayラテラル`: ラテラルボーの公差（Q_S_QA_DIM_179_日車 / Q_S_QA_DIM_179_白銅滑川）。
- Public 変数 `GradeArray`（L3）: 外観グレード配列を保持する。

### 1.3 QP_DB（QP）— 平坦度（歪）公差の本番エンジンと各種マスタ判定

- `DBQP`: 品質基準.accdb「平坦度一覧QP追加」（QP＝品証が保守するマスタ）と等級表（中歪/耳歪/ソリ 3級・4級）から歪公差配列 `Kousa(n,9)` を作る。
- `DistArray`: 1つのロットに複数の引当があるとき DBQP を引当ごとに実行し、厳しい方の公差で上書きする。
- `FormReflectDist`: 公差を UF1 / UF品質基準 / 全ヵ所検査シートに反映する。**UF1 と UFMitsubishi から実際に使われている主経路。**
- 汎用のマスタ照合エンジン: `FindInList` / `CheckCondition` / `CheckThickness` / `CheckWidth` / `CheckWithin` / `CheckSize` / `FindIndices` / `SoriArray` / `ExtractNums` / `CompareStrings`。QDB からも使われる。
- その他のマスタ判定:
  - `DbReturnItem`: ラテフリ直角。フリー反り・ラテラルボー・直角度の要否。
  - `MaterialsUpperPart`: 梱包資材マスタ「梱包保護材」。
  - `ShiftTime`: 時間用（直の開始・終了時刻）。
  - `StraightEdge`: 校正ストレートエッヂ。
  - `HikiArrCreate`: SIKAHIKINOW の引当一覧。
  - `CheckEquipmentDb`: 設備仕様。
  - `GaichuPackingNo`: 外注条件・外注条件明細。
  - `IsGaugeRequired`: ゲージ判定条件。

### 1.4 3モジュールの関係

```
UF1.CommandButton8 / UFMitsubishi ──► UF品質基準.Show
   ├─► DistArray (QP) ─► DBQP (QP) ─► FindIndices/LoopArray(QDB)/SoriArray/CheckXxx(QP)
   │                └─► FormReflectDist (QP) ─► UF1 / UF品質基準 / 全ヵ所検査シート
   └─► (FlagAtu>=8 or >6) GradeArray = DBdisplay外観_H (QDB) ─► FalseSearch_L, AddPagesArrN(H), UF_DbDisplay
UF品質基準.CommandButton1(手動検索) ─► DBdisplay外観_H + DBdisplay歪 (QDB)
UF品質基準.B1..B11 ─► BoColor (Q) + DBdisplay外観_S (QDB) ─► FalseSearch_S, 基準反映
Quality (Q) の 品質グレード/平坦度/表面品質基準 … 呼び出し元なし（旧版）
```

QDB と QP は互いに依存している（QDB → `CheckCondition` 等の QP 関数。QP → `LoopArray` 等の QDB 関数）。

### 1.5 UF品質基準（FQ）

品質基準を表示するモードレスフォーム。主な構成:
- 厚板/薄板のページ（MultiPage `MP`）
- 品質写真ページ `MPH`
- 外観異常タブ `TabAb`
- 異常ボタン B1〜B11
- 歪公差欄（`MG1`, `HaHiz1`, `TaHiz1`, `Naka1`, `Mimi1` と、その2系列目）
- 参照元 `Sansyou1`

**マスタを編集する機能は無い。**

### 1.6 UF管理（FK）

ローカルシート「保存用」（通常品）/「耳用」（耳付き）に保存された検査履歴を**閲覧する**フォーム。「呼び出し」ボタンで `HistoryLoad`（SaveLoad）を呼び、UF1 に復元する。**パスワードもマスタ編集も無い**（§6）。

### 1.7 補助（H）

- Win32 API 操作（マウスクリック、ウィンドウ操作、クリップボードの画像）
- 効果音（`C:\VBAWAV\`）
- 計測器入力の文字列変換 `Conv` / `Maker`
- ファイル・シートの存在確認、最終行・最終列
- シートの表示・非表示
- プログレス表示
- PDF 出力
- 安定マージソート
- DB 更新中の判定 `TimeDesignation` / `FileDesignation`
- 計測ログ出力 `LogOut`
- 検査番号の頭の決定 `FindPriority`
- 写真ページの生成など

---

## 2. Sub/Function 一覧

### 2.1 Quality（Q）

| 名前 | 行 | 公開 | 引数 | 戻り値 | 役割 | 呼び出し元 → 呼び出し先 |
|---|---|---|---|---|---|---|
| 品質グレード | 4 | Public Sub | なし | – | Lot情報履歴から UF1.Lot の行を探し、用途・材質・板厚から厚板の外観グレードと文言を UF品質基準 に書く | 元: なし（UF1 L3077 はコメント）／先: WsCheck, UR, UC |
| 平坦度 | 446 | Public Sub | ATU, HABA, TAKE As Double, Zai, Tyu, YoutC As String | – | 用途・材質・厚・幅・丈から ゲージ/巾歪/丈歪/中/耳 公差を UF品質基準 に書く | 元: なし（UF1 L3073, L3078 はコメント）／先: 中歪, 耳歪, WsCheck, UR, UC |
| 中歪 | 1638 | Public Function | ATU, HABA As Double, Optional grade As Long=4 | Single | 中歪公差（4級/3級） | 元: 平坦度 |
| 耳歪 | 1683 | Public Function | 同上 | Single | 耳歪公差 | 元: 平坦度 |
| 表面品質基準 | 1728 | Public Sub | LotNo As String | – | SIKALOTNOW → パターン → グレード表を UF品質基準 の TB1..E4 に表示 | 元: なし（UF1 L3072 はコメント）／先: ReadCheck, TimeDesignation, FileDesignation, GetRecordsArr, GetFieldsArr, PatternJudge, GradeJudge |
| PatternJudge | 1830 | Public Function | Hyou, Zai, Yout As String, ATU As Single | Long(0=該当なし) | 早見表パターン番号 1〜6 | 元: 表面品質基準 |
| GradeJudge | 1909 | Public Function | Abno As String, Pattern As Long | Variant(0始まり配列4要素 O,S,K,A) | 異常名×パターン → 4つの条件文言 | 元: 表面品質基準, SelectAdd |
| BoColor | 2235 | Public Sub | Bot As CommandButton | – | B1〜B11 を白に戻し、押したボタンを黄にする。SelectO/S/K/A をクリア | 元: FQ B1〜B11_Click |
| SelectAdd | 2248 | Public Sub | Bot As CommandButton | – | ボタン名から GradeJudge を引いて SelectO/S/K/A に表示 | 元: なし（FQ 内はコメント） |

### 2.2 Quality_DB（QDB）

| 名前 | 行 | 公開 | 引数 | 戻り値 | 役割 | 呼び出し元 → 呼び出し先 |
|---|---|---|---|---|---|---|
| (変数) GradeArray | 3 | Public | – | Variant | 外観グレード配列。TabAb_Change で使う | 書: UF1 L3110, UFMitsubishi L262／読: FQ L252 |
| FalseSearch_S | 5 | Public Function | TempArr(別表2 結合配列), Temp As Long(パターン) | Variant(1 To 5, 1 To n) | パターンに一致する行（O/S/K/A）× 異常名の表を作る | 元: DBdisplay外観_S |
| FalseSearch_L | 52 | Public Function | TempArr(別表1 結合配列), Resul(欠陥名,グレード) | Variant(n,3) 欠陥名/許容内容/グレード | 厚板の欠陥グレードから許容内容を引く | 元: DBdisplay外観_H |
| DBdisplay外観_H | 86 | Public Function | ロット, Optional TempPath | Variant (TempSize_L) | 厚板（オーダー厚 ≥ 8）の外観基準。TabAb にタブを追加して表示 | 元: UF1 L3110, UFMitsubishi L262, FQ L122／先: IsOver12Hours, ReadCheck, TimeDesignation, FileDesignation, Arrbox, GetFieldsArr, GetRecordsArr, Call_MergeArray_Row, CheckCondition, CheckThickness, CheckWidth, FalseSearch_L, AddPagesArrN, UF_DbDisplay |
| UF_DbDisplay | 411 | Public Sub | TempArr, tabCaption | – | タブ名と一致する行を AbnoLabel/gradeTB/AbnoTB に表示 | 元: DBdisplay外観_H, FQ TabAb_Change |
| DBdisplay外観_S | 462 | Public Sub | ロット, OutName(ボタンcaption), Optional TempPath | – | 薄板（オーダー厚 < 8）。パターンを決めて SelectO/S/K/A を表示 | 元: FQ B1〜B11／先: Arrbox, GetFieldsArr, GetRecordsArr, CheckCondition, FalseSearch_S, 基準反映 |
| 基準反映 | 676 | Public Sub | Temp(異常名), TempArr | – | FalseSearch_S の結果から O/S/K/A の文言を表示 | 元: DBdisplay外観_S |
| DBdisplay歪 | 700 | Public Sub | ロット, Optional TempPath | – | 歪公差を UF品質基準 に表示（DIM_122 系） | 元: FQ L124／先: Arrbox, GetFieldsArr, GetRecordsArr, LoopArray |
| LoopArray | 1386 | Public Function | TempA(レコード), tempB(フィールド), ReturnSize As Single, ATU As Single | Single | 「列名が数値」のサイズ表から、厚の範囲に入る最初の行の該当サイズ列を返す | 元: QDB 多数, QP DBQP |
| BhSearch | 1454 | Public Function | tempZ(材質), TempT(調質) | Long | バーコール下限値（無ければ 0） | 元: Judg L411 BhJudg |
| DBdisplayラテラル | 1491 | Public Function | ロット, TargetText("日車"/"白銅滑川"), Optional TempPath | Single | ラテラルボー公差（製造丈で表引き） | 元: QP DbReturnItem L1359/1363/1429/1431 |

### 2.3 QP_DB（QP）

| 名前 | 行 | 公開 | 引数 | 戻り値 | 役割 | 呼び出し元 → 呼び出し先 |
|---|---|---|---|---|---|---|
| C_TARGET_HOSONOS | 3 | Public Const | – | String | `"1P0001,1P0104,1P0110,1P0115,1P0116,1P0118,1P0119,1P0120"`（蓋アングルの絡む包装仕様書No） | IsTargetHosoNo |
| ClearMaterialsCache | 64 | Public Sub | – | – | 梱包保護材のキャッシュを消す | 元: **なし** |
| DBQP | 69 | Public Function | ロットNo, 検査No, delivery(納入先), Good(取引先), Optional TempPath, Optional Execution | Variant Kousa(1..n,1..9) または Empty | 歪公差の決定（§4.3） | 元: DistArray／先: ReadCheck, TimeDesignation, FileDesignation, Arrbox, FindPriority, GetFieldsArr, GetRecordsArr, FindIndices, CheckCondition, CheckThickness, CheckWidth, CheckSize, ExtractNums, CompareStrings, LoopArray, SoriArray |
| FindInList | 804 | (Public) Function | tMP(DB側), List(探す値), Optional BlankFlag | Boolean | カンマ区切りの条件式（*、≠ を含む）を評価（§4.4） | 元: CheckCondition, Test.frm |
| CheckCondition | 916 | (Public) Function | row, Column, Criteria, ArrQuality, Optional BlankFlag | Boolean | DB セルを半角化し、「ｰ」→「-」に置換してから FindInList | 元: QP / QDB 多数 |
| CheckThickness | 928 | (Public) Function | row, Col1(下), Col2(上), Col3(値列), OrderThickness, ArrQuality | Boolean | 3列とも空なら True。それ以外は 下 ≤ x ≤ 上 | 多数 |
| CheckWidth | 947 | (Public) Function | 同上 | Boolean | 同じロジック（幅・丈） | 多数 |
| CheckWithin | 966 | (Public) Function | row, Col1, Col2, Order, ArrQuality | Boolean | 空チェック無しで 下 ≤ x ≤ 上 | CheckEquipmentDb |
| CheckSize | 974 | (Public) Function | row, Col1, OrderSize, ArrQuality | Boolean | 空なら True。それ以外は Val(OrderSize) ≥ Val(DB) | DBQP（結果は最終判定に未使用） |
| FindIndices | 985 | (Public) Function | Arry(フィールド), WidthLength("H"/"T"), ReturnSize | Variant(1 To 2): (1)=列index, (2)=9999 ならサイズ超過 | 「H500」「T2000」形式の列を選ぶ | DBQP, Test.frm |
| ExtractNumber | 1051 | (Public) Function | str | Variant | 以上/以下があれば最初の数値 | 元: **なし** |
| CompareStrings | 1093 | (Public) Function | tMP, Lot厚 As Single, DB厚 As Single | Boolean | 「以上」なら Lot ≥ DB、「以下」なら Lot ≤ DB | DBQP |
| IsHalfWidth | 1103 | (Public) Function | S | Boolean | 全文字が AscW ≤ &HFF | 元: **なし** |
| ExtractNums | 1115 | (Public) Function | tMP, Condition | String | カンマで分け、Condition を含む部分から数値を抽出し「, 」で連結 | DBQP |
| ExtractNumFromPart | 1139 | (Public) Function | part, Condition | String | 最初の連続した数字と「.」1個 | ExtractNums |
| DbReturnItem | 1176 | Public Function | Nou, Tori, YoutC, Zai, Tyou As String, ATU, HABA, TAKE As Single | String（",フリー反り,ラテラルボー,直角度" の部分集合） | ラテフリ直角マスタ判定と、UF1 / シートへの表示 | 元: UFMitsubishi L842, FormRelat L5292, OutTolerance L198／先: GetFieldsArr, GetRecordsArr, CheckCondition, CheckThickness, CheckWidth, DBdisplayラテラル, CBAdd, WsCheck |
| SoriArray | 1458 | Public Function | TempA, tempB, ATU, HabSize, TakSize As Single, FindIndices(列idx) | Single | 厚の範囲に入る最初の行の指定列 | DBQP, Test.frm |
| IsTargetHosoNo | 1618 | Public Function | HosoNo | Boolean | C_TARGET_HOSONOS に含まれるか | MaterialsUpperPart |
| EnsureCache | 1633 | Private Function | OrdNo | Boolean | 梱包保護材の読込キャッシュ | MaterialsUpperPart |
| MaterialsUpperPart | 1674 | Public Function | OrdNo, Optional CheckBikou As Boolean=False | String（使用保護材 / "一致なし" / ""） | 緩衝材（上部保護材）判定 | 元: UF_PackNo L176-177, FormRelat L6900-6901, SaveLoad L1624-1625, Test.frm |
| ShiftTime | 1786 | Public Function | – | Variant(1..4,1..3) 直/開始/終了 または "" | 直の時刻表 | 元: Worker L15, Test.frm |
| StraightEdge | 1831 | Public Function | LineName | Variant(1..2,1..1) 長さ下/長さ上 または "" | 校正ストレートエッヂの範囲 | 元: Proofreading L92, Test.frm |
| HikiArrCreate | 1867 | Public Function | LotNo, Optional n(未使用) | Variant(1..H,1..2) 納入先名称/取引先名称 | 引当一覧 | DistArray, Test.frm |
| DistArray | 1911 | Public Function | LotNo, 検査No, Optional TempPath, Optional Execution | Variant(Kousa) | 引当ごとに DBQP を実行し、最小公差で上書き | 元: UF1 L3102, L4525, UFMitsubishi L256, L846, Test.frm |
| FormReflectDist | 1986 | Public Sub | Kousa, Optional Execution(未使用) | – | 公差をフォーム・シートに反映 | 元: UF1 L3106, L4533, UFMitsubishi L260, L850 |
| UpdateQpArryD | 2057 | (Public) Function | QpArryD, QpArryG | Variant | 小さい方で上書き＋コメント | 元: **なし**（呼び出し側はコメント） |
| CheckEquipmentDb | 2084 | Public Function | line, Zai, ATU, HABA, TAKE, Wei As Single | String | 設備仕様との適合判定 | 元: SaveLoad L165, Test.frm |
| GaichuPackingNo | 2250 | Public Function | Hist() (Lot情報履歴の配列, ByRef), LotNo, R_Lot | String（包装仕様書NO/""） | 外注条件を判定 | 元: FormRelat L6701 |
| MatchCondSet | 2272 | Private | Hist, LotNo, R_Lot, headNo | Boolean | 条件番号どうしは AND | GaichuPackingNo |
| MatchCondNo | 2287 | Private | …, condNo | Boolean | 同じ番号内は OR | MatchCondSet |
| JudgeOne | 2317 | Private | cellVal, method, judgeVal | Boolean | = / Like / 含む / >= / <= | MatchCondNo |
| FindHistCol | 2341 | Private | Hist, target | Long | 見出しの列番号 | MatchCondNo |
| GetCondNos | 2349 | Private | headNo | Variant(Double配列)/Empty | 有効な条件番号の一覧 | MatchCondSet |
| ReloadGaichu | 2383 | Public Sub | – | – | 外注条件の再読込（手動マクロ） | 元: **コードからの呼び出し無し** |
| EnsureLoaded | 2396 | Private | – | Boolean | 外注条件の初回読込（失敗するとその後も再試行しない） | GaichuPackingNo |
| LoadTable | 2432 | Private Sub | FilePath, TableName, colNames, ByRef data, ByRef idx() | – | 必須列の確認と読込 | EnsureLoaded, EnsureLoadedGauge |
| IsGaugeRequired | 2473 | Public Function | 厚, 幅, 丈 As Double, 材質, 調質, コース As String | Boolean（True=ゲージ検査、False=定盤検査） | ゲージ判定 | 元: FormRelat L3572, SaveLoad L487／先: Weight |
| HasVal | 2514 | Private | v | Boolean | Null・Empty・空白のみなら False | IsGaugeRequired |
| ReloadGauge | 2520 | Public Sub | – | – | 再読込（手動マクロ） | 元: **なし** |
| EnsureLoadedGauge | 2528 | Private | – | Boolean | 初回読込（失敗するとその後も再試行しない） | IsGaugeRequired |

（「(Public)」は修飾子が省略されたもので、VBA の既定により Public。）

### 2.4 UF品質基準（FQ）

| イベント | 行 | 内容 |
|---|---|---|
| B1〜B11_Click | 11-76 | `BoColor(Bn)` → `DBdisplay外観_S(Me.T1, Me.Bn.Caption, UF1.DBflag.Text)` |
| CommandButton1_Click（手動検索） | 79-129 | 表示欄をクリア（ToG〜KaG、Tobi〜Kami、MG1/HaHiz1/TaHiz1/Naka1/Mimi1、ATU/HABA/TAKE/Zai/TYO/YoutC、Uatu〜Uhyou、Sansyou1）→ L3 を表示 → `DBdisplay外観_H(T1, PathText)` → `DBdisplay歪(T1, PathText)` |
| CommandButton3/4 | 132-138 | MP.Value = 0 / 1 に切り替え |
| FM1_Click | 140 | UF1.Show vbModeless |
| Fm2_Click | 143 | B1〜B11 を白に戻し、SelectO/S/K/A をクリア |
| CommandButton6 | 155 | Unload Me |
| HaHiz1/MG1/Mimi1/Naka1/Sansyou1/TaHiz1 _Change | 159-200 | 2系列目（HaHiz2/MG2/Mimi2/Naka2/Sansyou2/TaHiz2）へコピー |
| MPH_Change | 168 | LoadImagesIntoMultiPageFromTemp。選択中ページの caption を「　★…★　」で強調し、他ページから★を除去 |
| UserForm_QueryClose | 201 | 全ヵ所検査シートが表示中なら UFMitsubishi を、そうでなければ UF1 を Show（Cancel はしない） |
| UserForm_Activate | 210 | FormSetting |
| UserForm_Initialize | 213 | GetZoom, InitialSetting。一部の欄（ToG〜KaG、Tobi〜Kami、MG1/HaHiz1/TaHiz1、Uatu〜Uhyou、Upater）を空にし、B1〜B11 を白にする |
| TabAb_Change | 249 | 選択タブの caption で `UF_DbDisplay(GradeArray, caption)` |
| Utyou/UYoucode/Uzai/Uhyou/YoutC/Zai/Tyo _KeyPress | 256-332 | 入力制限: 0-9 と A-Z は許可。a-z は大文字に変換（-32）。それ以外は 0（入力拒否） |

### 2.5 UF管理（FK）

| イベント | 行 | 内容 |
|---|---|---|
| NewCB_Change | 11 | 同じ LotNo が複数あるとき、NewCB（作業時間）と CB1 の Lot が一致する行に `CB1.ListIndex = i-2` を設定 |
| CB1_Change | 37 | 選択 Lot の履歴行を表示（§6） |
| CommandButton1_Click | 620 | UF1 表示中で UF1.Lot が空でなければ MsgBox「別Lotデータがあります、新規登録処理を行ってください」で中止。それ以外は `HistoryLoad(CB1.Text[, NewCB.Text])` → `InitialSetting(UF1,"サイズ保存")` |
| CommandButton6 | 637 | Unload |
| MIMI_AfterUpdate | 641 | 耳付きチェックの切替。CB1 に「耳用」または「保存用」の B 列（2行目〜最終行）を読み込む |
| UserForm_Activate / Initialize | 666/670 | FormSetting。Initialize では CB1 に「保存用」の B 列を読み込む |

### 2.6 補助（H）

§7 に一覧を載せた。

---

## 3. 品質基準（規格値・マスタ）の取得元・構造・キー

### 3.1 取得元の全体像

| 基準 | 実体 | 場所 | キー（照合項目） |
|---|---|---|---|
| 厚板外観グレード（旧） | **ハードコード** | Q L19-59, L100-441 | 用途コード(H191/H207/H197/K21S)、材質、得意先/納入先/送り先（ﾅﾒｶﾜ判定）、板厚（UF1.ATU） |
| 平坦度（旧） | **ハードコード** | Q L497-1635 | 用途コード、材質、ATU/HABA/TAKE、VC_A/VC_B（Lot情報履歴） |
| 中歪/耳歪（旧） | **ハードコード** | Q L1638-1727 | 厚、幅、等級 |
| 表面パターン/グレード（旧） | **ハードコード** | Q L1830-2234 | 表面処理(B/H/K/S/T/C)、材質（Like）、用途コード、厚 |
| 厚板外観（新） | **Access 品質基準.accdb** `Q_S_QA_124_別表2`（グレード）＋`Q_S_QA_124_別表1`（行列入替の許容表） | QualityPath | 用途コード(LIKE)、材質、厚（製造厚）、幅/丈（オーダー）、納入先 |
| 薄板表面（新） | 同 `Q_S_QA_SRF_128_別表1`（パターン判別）＋`_別表2`（グレード表） | QualityPath | 用途コード(LIKE)、材質、表面処理コード(Like)、厚フラグ/板厚1（オーダー厚） |
| 歪（UF品質基準 表示） | 同 `耳歪4級` `中歪4級` `耳歪3級` `中歪3級` `QA_022_01` `Q_S_QA_DIM_122_A/B/C/D` `ソリ4級` `ソリ3級` | QualityPath | §4.2 |
| 歪（本番：UF1 反映） | 同 `平坦度一覧QP追加` ＋等級表 | QualityPath | §4.3（検査番号の頭、用途、材質、調質、厚、幅、丈、納入先、取引先、送り先 等） |
| バーコール下限 | 同 `Q_S_KE_ALL_053` | QualityPath | 材質（右3文字 LIKE）、調質（完全一致） |
| ラテラル公差 | 同 `Q_S_QA_DIM_179_日車` / `_白銅滑川` | QualityPath | 材質、製造厚、製造幅 → 製造丈で表引き |
| フリー反り/ラテラル/直角度の要否 | 同 `ラテフリ直角`（無ければ PackingPath の同名ファイル） | QualityPath→PackingPath | 納入先、取引先、用途、製造材質、製造調質、厚/幅/丈 範囲 |
| 外注梱包 | 同 `外注条件`＋`外注条件明細` | QualityPath | 明細で指定した Lot情報履歴の任意列 |
| ゲージ判定 | 同 `ゲージ判定条件` | QualityPath | 重量・丈・幅・材質・調質・コース |
| 上部保護材 | `梱包資材マスタ.accdb` `梱包保護材` | PackingPath | 包装仕様書、材質、調質、用途、厚/幅/丈、備考="当板用" |
| 直の時刻 | `伝送用ファイル.accdb` `時間用` | QualityPath→PackingPath | なし（全件の先頭4行） |
| ストレートエッジ | `伝送用ファイル.accdb` `校正ストレートエッヂ` | 同上 | ライン名 |
| 設備仕様 | `伝送用ファイル.accdb` `設備仕様` | 同上 | 設備名、製造材質、厚/幅/丈/単重 MIN-MAX |
| ロット情報 | `SIKALOTNOW.accdb` `仕掛`（Arrbox。表面品質基準のみ `SIKATWOYEAR.accdb` に再試行） | MainPath→NextPath | [ﾛｯﾄ番号] |
| VC 情報 | `SIKAODRNOW.accdb` `仕掛` | DBdisplay歪 で決まった Main/Next | [受注番号]（SIKALOT ｵｰﾀﾞｰ番号） |
| 引当 | `SIKAHIKINOW.accdb` `仕掛` | MainPath→NextPath | [ﾛｯﾄ番号] |

パス（初期設定 L18-35）: レジストリ「モード」が "オンライン" のとき、
- MainPath = `\\nlmsrvngy03\Read\【New】仕掛\台帳\`
- NextPath = `\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\仕掛\`
- QualityPath = `\\nlmsrvngy03\工場内共有\検査データ\品証-梱包やり取りフォルダ\Accdb\`
- PackingPath = `\\nlmfangyshrd\各課共有\0130_日軽稲沢\梱包課\AIM\【■】_参照用ファイル\`
- HistoryPath = `\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\検査履歴\`

"オフライン" のときは全部 `ThisWorkbook.path & "\"`。**どちらでもない（"存在なし"）ときは全パスが空文字**になり、相対パスとして扱われる（要確認）。

### 3.2 SIKALOTNOW（仕掛）の使用列

`ｵｰﾀﾞｰ板厚` `ｵｰﾀﾞｰ板幅` `ｵｰﾀﾞｰ板丈` `ｵｰﾀﾞｰ材質` `ｵｰﾀﾞｰ調質` `製造板厚` `製造板幅` `製造板丈` `製造材質` `製造調質` `品質ｸﾞﾚｰﾄﾞ_表面処理` `品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ` `納入先名` `納入先ｺｰﾄﾞ` `取引先名` `取引先ｺｰﾄﾞ` `送り先名` `送り先ｺｰﾄﾞ` `用途ｺｰﾄﾞ` `用途名` `ｵｰﾀﾞｰ番号` `VC` `実績_設備ｺｰｽ` `BOX設計_横割数(板)` `指示_中歪_高さ` `指示_耳歪_高さ` `指示_そり巾_方向高さ` `指示_そり丈_方向高さ`（QP L117-151, QDB L122-139/L742-757/L1531-1545）。

`Arrbox(...,"Record")` で複数行が返ったときは先頭行 `LotAcc(1, …)` だけを使う。

### 3.3 Lot情報履歴シートの列（sheets/Lot情報履歴.csv 1行目）

A 登録日 / B ｵｰﾀﾞｰNo / C ﾛｯﾄNo / D 厚 / E 巾 / F 丈 / G 検査No / H 材質 / I 調質 / J 用途ｺｰﾄﾞ / K ｺｰｽ / L 得意先 / M 納入先 / N 送り先 / O 前工程実績枚数 / P 前工程実績重量 / Q 合紙 / R VC / S VC_A / T VC_B / U 単重 / V 引当NO / W 引当調整NO / X 引当数量 / Y 引当ﾌｧｲﾙ登録 / Z 包装仕様NO

- `品質グレード` は列番号を固定で読む（3=Lot, 4〜6=寸法, 8=材質, 9=調質, 10=用途, 12=得意先, 13=納入先, 14=送り先）。正しい。
- `平坦度` は Tok=11(ｺｰｽ), Nou=12(得意先), Oku=13(納入先) と **1列ずれている**（Q L480-482）。ただし平坦度の中で Tok/Nou/Oku は使われないので結果には影響しない。VCA=19, VCB=20 は正しい。
- `MaterialsUpperPart` と `GaichuPackingNo` は見出し名で列を探す。

### 3.4 ハードコード vs DB の結論

- 実運用の判定・表示（UF1 への反映）は **DB（品質基準.accdb）だけ**。
- ハードコードが残っているもの:
  - Q モジュール全体（未使用）
  - VC の銘柄コード（`208PRNF` / `V325NNF` / `V311PNF`。Q L489-490, QDB L844-845）
  - 包装仕様書No の一覧（QP L3）
  - 調質 "O" / "Hx8" の判定不要ルール（QP L643-662）
  - FindPriority のコース → 頭文字の対応（H L1458-1488）
  - ラテラル特例の納入先（QP L1276-1306）
  - しきい値の厚 8/6/4（§8）

---

## 4. 判定・計算ロジック（原文・行番号）

### 4.1 Quality（旧・未使用。移植するなら全部再現が必要）

#### 4.1.1 品質グレード（Q L4-445）

- 文言の定義（L19-59。原文）:
  - G1〜G5 = "ｸﾞﾚｰﾄﾞ1"〜"ｸﾞﾚｰﾄﾞ5"
  - To1 = "- (切削品)"、To2 = "*大きさ:3～5mmΦ *数:3個/枚以下"、To3 = "*大きさ:2～3mmΦ *数:3個/枚以下"、To4 = "*大きさ:1～2mmΦ *数:1個/枚以下"、To5 = "無き事"
  - Ki1 = "-"、Ki2 = "*触感:爪に「かかる」 *大きさ:100mm以下"、Ki3 = "…「かかる」 *大きさ:50mm以下"、Ki4 = "*触感:爪に「かからない」 *大きさ:100mm以下"、Ki5 = "無き事"
  - Ha1 = "触感無き事(ﾌﾟﾘﾝﾄ)"、Ha2 = "*ﾌﾟﾘﾝﾄ *大きさ:200mm以下"、Ha3 = "…50mm以下"、Ha4 = "…20mm以下"、Ha5 = "無き事"
  - AL1 = "-"、AL2 = "*触感:あり *大きさ:300mm以下 *数:2個/枚以下"、AL3 = "…150mm以下 *数:2個/枚以下"、AL4 = "…50mm以下 *数:1個/枚以下"、Al5 = "無き事"
  - Bo1 = "-"、Bo2 = "*大きさ:5～10mmΦ以下 *数:3個/枚以下"、Bo3 = "*大きさ:3～5mmΦ以下 *数:3個/枚以下"、Bo4 = "*大きさ:3～5mmΦ以下 *数:1個/枚以下"、Bo5 = "無き事"
  - Ka1 = "-"、Ka2 = "*触感:あり *大きさ:200mm以下 *数:5個/枚以下"、Ka3 = "…100mm以下 *数:3個/枚以下"、Ka4 = "…100mm以下 *数:1個/枚以下"、Ka5 = "触感無き事"
- Lot情報履歴の2行目以降で C 列 = UF1.Lot の最初の行を使う（L69-75）。無ければ何もしない。
- Keyword = UF1.ATU（Double）。Select Case の範囲は両端を含む。
- 判定表（To/Ki/Ha/AL/Bo/Ka のグレード。文言は同じ番号を使う）:

| 用途 | 材質 | 8–10 | 12–30 | 30.1–55 | 55.1–75 | 75.1–300 |
|---|---|---|---|---|---|---|
| H191（高精度）L100 / H207（中精度）L170（両者同一） | 52S, F52S | 4,4,4,4,4,"" | 4,4,4,3,3,"" | 3,3,3,2,2,"" | 2,2,1,1,1,"" | 1,1,1,1,1,"" |
| H197 かつ ﾅﾒｶﾜ L241（Flag1="ﾅﾒｶﾜｱﾙﾐ"） | 52S,F52S,NP56,14S,17S,N17S,61S,R61S | 4,4,4,4,4,5 | 4,4,3,3,3,5 | 4,4,3,3,2,5 | 55.1–300: 4,4,3,3,2,5 | （同左） |
| H197 その他 L300 | 同上 | 3,3,4,4,4,"" | 3,3,3,3,3,"" | 2,2,2,2,2,"" | 1×5,"" | 1×5,"" |
| K21S（鉄道ボディ）L371 | U74S, NP56, 75S | 2,2,2,1,2,"" | 2,2,2,1,2,"" | 2,2,2,1,2,"" | 1×5,"" | 1×5,"" |

- ﾅﾒｶﾜ判定（L241、原文）: `Tok = "ﾅﾒｶﾜｱﾙﾐ(ｶ" Or Nou = "ﾅﾒｶﾜｱﾙﾐ(ｶ" Or Oku = "ﾅﾒｶﾜｱﾙﾐ(ｶ" Or Oku = "ﾅﾒｶﾜｱﾙﾐ(ｶ ｲﾀｺｺｳｼﾞﾖｳ"`。Flag1 は材質が一覧に無くても設定される。
- **範囲の隙間**: <8、(10,12)、(30,30.1)、(55,55.1)、(75,75.1)、>300 では何も書かない。そのため**前回の表示値が残る**（Initialize 以外でクリアしない）。

#### 4.1.2 平坦度（Q L446-1637）

- Lot情報履歴に一致行があれば、引数を履歴の値で上書きする（L473-491。ATU は Format "0.000"、幅/丈は "0.0" の文字列を Double に戻す）。
- VcFlag = VCA または VCB が `208PRNF` / `V325NNF` / `V311PNF` のどれか（L489-490）。
- 幅の区分（厚板。H 系）: (0,500], (500,750], (750,1004], (1004,1254], (1254,1529], (1529,2004], (2004,2504], (2504,3054]。
  - 丈も同じ区分で、**丈 > 3054 は最後の区分と同じ値**。
  - **幅 > 3054 は何も書かない**（U74S を除く）。
  - 幅 ≤ 0 も書かない。
- K21S の区分: 500 / 750 / 1000 / 1250 / 1525 / 2000 / 2500 / 3050。
- 中歪/耳歪は `If ATU <= 6# Then .Naka1 = Format(中歪(ATU, HABA), "0.0")`（同様に .Mimi1）。中耳を書くのは、各表で「中耳」と記した Case だけ。

値は8区分の順に列挙する（丈も同じ値で、>上限は最終値）:

| 用途コード（行） | 材質 | 厚 Case | MG1 | 巾/丈歪（8区分） | 中耳 |
|---|---|---|---|---|---|
| H191, H242（L500） | 52S,NP56,F52S,5083 | 6–50 | "0.2以下" | 0.1,0.2,0.2,0.3,0.3,0.4,0.4,0.4 | ATU≤6 |
| H196,H299,H207,H234,H237,H238,H239,H241,H272,H255（L546） | 同上 | 4–11.999 | "0.5以下" | 0.3,0.4,0.5,0.6,0.8,1.0,1.0,1.0 | ATU≤6 |
| 〃 | 〃 | 12–50 | "0.4以下" | 0.2,0.3,0.4,0.5,0.6,0.7,0.7,0.7 | – |
| 〃 | 〃 | 50.1–100 | "1.2以下" | 0.6,0.9,1.2,1.5,2.0,2.5,3.0,3.5 | – |
| H205, H251（L674） | 同上 | 4–11.999 | "0.5以下" | 0.3,0.4,0.5,0.6,0.8,1.0,1.0,1.0 | ATU≤6 |
| 〃 | 〃 | 12–30 | "0.4以下" | 0.2,0.3,0.4,0.5,0.6,0.7,0.7,0.7 | – |
| 〃 | 〃 | 30.1–50 | "1.0以下" | 0.5,0.8,1.0,1.3,1.8,2.0,2.0,2.0 | – |
| 〃 | 〃 | 50.1–100 | "1.2以下" | 0.6,0.9,1.2,1.5,2.0,2.5,3.0,3.5 | – |
| H197 **かつ VcFlag**（L836。材質制限なし。Flag2="H197 and 指定VC "） | 全材質 | 4–11.999 / 12–50 / 50.1–100 | 0.5 / 0.4 / 1.2 以下 | H196 群と同じ3行 | 4–11.999 のみ ATU≤6 |
| H197,H19S,H19B,H19T,H271,H248,H249（L956。一般材） | 52S,NP56,F52S,5083 | 4–50 | "1.0以下" | 0.5,0.8,1.0,1.3,1.8,2.0,2.0,2.0 | ATU≤6 |
| 〃 | 〃 | 50.1–100 | "1.2以下" | 0.6,0.9,1.2,1.5,2.0,2.5,3.0,3.5 | – |
| 〃 | 〃 | 100.1–150 | "1.4以下" | 0.7,1.1,1.4,1.8,2.0,2.5,3.0,3.5 | – |
| 〃 | 〃 | 150.1–200 | "2.0以下" | 1.0,1.5,2.0,2.0,2.0,3.0,4.0,4.0 | – |
| 〃 | 〃 | 200.1–500 | "3.0以下" | 巾 3.0,3.0,3.0,5.0,5.0,5.0,5.0,5.0 ／ 丈 3.0,**"3.."**,3.0,5.0,5.0,5.0,5.0,5.0 | – |
| K21S（L1161）MG1/中耳なし | 5005 | 0.1–5 | – | 1.5,2.0,2.0,2.0,2.0,3.0,3.0,3.0 | – |
| 〃 | 52S,F52S | 0.1–5 | – | 1.5,2.0,2.0,2.0,2.0,3.0,3.0,3.0 | – |
| 〃 | 〃 | 5.01–75 | – | 1.5,2.0,2.0,2.0,2.0,3.0,3.5,3.5 | – |
| 〃 | 〃 | 75.1–200 | – | 巾 2.0,2.0,2.5,3.0,3.0 と、6〜8区分は **.TaHiz1 に代入**（L1293/1295/1297 の誤記: 3.5,4.0,4.0）／丈 2.0,2.0,2.5,3.0,3.0,3.5,4.0,4.0 | – |
| 〃 | 〃 | 200.1–500 | – | 2.0,2.5,3.0,3.0,3.5,4.0,4.0,4.5 | – |
| 〃 | NP56,5083 | 5.01–75 | – | 1.5,2.0,2.0,2.5,2.5,2.5,3.0,3.5 | – |
| 〃 | 〃 | 75.1–120 | – | 2.0,2.0,2.5,2.5,2.5,3.5,3.5,3.5 | – |
| 〃 | 〃 | 120.1–200 | – | 2.0,2.5,2.5,3.0,3.0,3.5,3.5,4.0 | – |
| 〃 | 〃 | 200.1–500 | – | 2.0,2.5,3.0,3.5,3.5,4.0,4.0,4.5 | – |
| 〃 | U74S | 4–50 | – | 2.0,2.5,2.5,2.5,2.5,3.0,3.5,3.5（**幅 >3050 も 3.5**） | – |
| 〃 | 〃 | 50.1–75 / 75.1–120 | – | 2.0,2.5,2.5,3.0,3.0,3.5,4.0,5.0 | – |

- 誤記の原文（L1141）: `.TaHiz1 = "3.."`。
- L1293 の誤記で、K21S/52S の 75.1–200 で幅 >1525 のとき **HaHiz1 は書かれず前回値が残る**（TaHiz1 は直後の丈ブロックで上書きされる）。
- 厚の隙間（11.999〜12、50〜50.1、5〜5.01 など）は書き込みなし。

#### 4.1.3 中歪 / 耳歪（Q L1638-1727）

幅 ≤500 / 幅 ≥501 の値。**500 < 幅 < 501 のときは 0 を返す**。

| ATU | 中歪4級 | 耳歪4級 | 中歪3級 / 耳歪3級（同一） |
|---|---|---|---|
| 0–0.5 | 1.5 / 2 | 1.5（Format "0.0" を Single に代入）/ 2 | 2 / 4 |
| 0.51–2 | 1 / 2 | 1 / 2 | 2 / 4 |
| 2.01–5 | 1 / 2 | 1.5 / 2 | 3 / 4 |
| 5.01–12 | 1 / 2 | 1.5 / 2 | 3 / 4 |
| 12.01–30 | 1.5 / 1.5 | 1.5 / 1.5 | 3 / 3 |
| 30.01–50 | 1.5 / 1.5 | 1.5 / 1.5 | 3 / 3 |

隙間（0.5〜0.51 など）と ATU >50 のときは 0。

#### 4.1.4 PatternJudge（Q L1830-1908）

- 先頭の判定: `If Hyou = "" Or Zai = "" Or Yout = "" Or IsNumeric(ATU) = False Then Exit Function`（0 を返す）。
- `Select Case True` は**最初に一致した Case だけ**が評価される。

| 表面 | 材質条件（Like） | 用途 | パターン |
|---|---|---|---|
| B,H,K | `*52S` | H115,H127,U19B,P191,H19B,K21B,H196,U194,H228,H299,H248,H249 | ATU 0–2.5→1、2.51–7.999→2 |
| B,H,K | `= "NK01"` | T238,T23B,T23H,T23K | 4 |
| B,H,K | `*995` | U19B,H19B,T271 → 4 ／ F191,F61B → 6 | |
| B,H,K | `*17S` or `*61S` or `*75S` or `*24S` | H115,R191,U19B,P197,H19B,P181 | 2 |
| S | `*52S` | K32S,K19S,H197,H19S,K194,C21S,H193,H211,K21S,J19S,K197,T322,T323,T324,T321 | 0–2.5→1、2.51–7.999→2 |
| S | `*52S` | "T323","T32S","T321"（**T323/T321 は直前の Case が先に一致するので到達しない**。実質 T32S のみ） | 5 |
| S | `= "NP56"` | H197,H19S,C218,C217,K21S → 2 ／ C215,C216,C219 → 5 | |
| S | `*995` | R292,H19S → 4 ／ N192,F61S → 6 | |
| S | `*17S/*61S/*75S/*24S` | H197,H19S | 2 |
| S | `= "5154"` | N193 | 6 |
| T,C | `*52S` | U19T,F97T,H19T,K21T | 3 |
| T,C | `= "NK01"` | T23T | 3 |
| T,C | `*995` | U19T,H31C | 3 |

#### 4.1.5 GradeJudge（Q L1909-2234）

戻り値は Array(O, S, K, A)。O=大きさ、S=触感、K=数、A=素洗い。原文の値を `/`（＝"／"）区切りで並べた。並びはパターン 1 / 2 / 3 / 4 / 5 / 6。

| 異常 | O | S | K | A |
|---|---|---|---|---|
| トビコミ | 2mm以下×5, 無き事 | 爪にかからない/爪に軽くかかる/爪に軽くかかる/爪にかからない/爪にかからない/無き事 | ﾋﾟｯﾁ300mm以上/ﾋﾟｯﾁ100mm以上/ﾋﾟｯﾁ100mm以上/ﾋﾟｯﾁ300mm以上/ﾋﾟｯﾁ300mm以上/無き事 | ／×6 |
| キズ | 長さ30～50mm×5, 無き事 | （トビコミと同じ） | 3個まで×5, 無き事 | ／×6 |
| 突 | ／×5, 無き事 | （トビコミと同じ） | ／×5, 無き事 | ／×6 |
| トモズレ | ／×5, 無き事 | 爪にかからない/爪にかからない/／/爪にかからない/爪にかからない/無き事 | ／×5, 無き事 | ／×6 |
| 汚れ | 長さ10mm以下/長さ30mm以下/／/長さ10mm以下/長さ10mm以下/無き事 | ／×6 | 3個まで/3個まで/／/1個まで/3個まで/無き事 | ／×6 |
| 油より | 目立たない事/／/／/目立たない事/／/／ | ／×6 | ／/1本まで/／/／/無き事/無き事 | ／/1分で消える事/／/／/／/／ |
| ロールキズ | 3mm以下/5mm以下/／/3mm以下/3mm以下/無き事 | ／×6 | ﾋﾟｯﾁ500mm以上/ﾋﾟｯﾁ500mm以上/／/ﾋﾟｯﾁ500mm以上/ﾋﾟｯﾁ1000mm以上/無き事 | 1分で消える事/／/／/1分で消える事/1分で消える事/／ |
| ロールマーク | 20mm以下/50mm以下/／/10mm以下/10mm以下/無き事 | ／×6 | 1M当たり1個まで/1M当たり2個まで/／/1M当たり1個まで/1M当たり1個まで/無き事 | ／×6 |
| カミコミ | 無き事/全巾に無き事/／/無き事/無き事/無き事 | ／×6 | ／×6 | ／×6 |
| 切断面 | ／×5, 無き事 | ビビり無き事/段差無き事/段差無き事/ビビり無き事/ビビり無き事/無き事 | ／×6 | ／×6 |
| 変色 | 境界目立たない事/全丈に無き事/／/境界目立たない事/／/／ | ／×6 | ／×6 | ／/／/／/1分で消える事/／/／ |

表面品質基準（L1816-1828）では、"TB","K","TA","TM","Y","A","RK","RM","KM","S","E" & i（i=1..4）に配列の要素 i-1 を表示する。

### 4.2 Quality_DB の判定

#### 4.2.1 共通の前処理（DBdisplay外観_H/S, DBdisplay歪, DBdisplayラテラル。QP の DBQP も同じ）

```
TempPath = IIf(TempPath="", MainPath, NextPath)
If TempPath = MainPath Then
  If ReadCheck(MainPath&"SIKALOTNOW.accdb",1)=False Or TimeDesignation Or FileDesignation=False Then
    TempPath = NextPath : If ReadCheck(NextPath&target,1)=False Then MsgBox ... : Exit
TempH = IIf(TempPath=NextPath, UF1.DBflag.Text, "")   ' DBQP は常に ""（QP L97、コメント化）
HiAcc = Arrbox("SIKALOTNOW.accdb","板","Fields",TempH) ; LotAcc = Arrbox(...,ロット,"Record",TempH)
```

- Arrbox の内部でも同じ Main/Next の判定をもう一度行う（AccHIki L19-37）。
- `path<>""` のときは NextPath に固定され、ReadCheck もしない。
- メッセージは `IsOver12Hours = False` のときだけ出す。この関数は WorkerTime_<ライン> シートの C1 から DateDiff("h") ≥ 1 なら True（Worker L328-349）。

#### 4.2.2 DBdisplay外観_H（QDB L86-410）

1. 厚板の判定: `If オーダー厚 >= 8# Then AtuFlag = True`（L184。2026.03.27 修正後）。False なら何もせず Empty を返す。
2. `Q_S_QA_124_別表2` のフィールド名を取得し、列番号を探す（L221-247）。条件列（管理番号, 材質, 用途コード, 板厚下, 板厚上, 厚, 更新日, 備考, 調質, 納入先, 板幅下, 板幅上, 幅, 板丈下, 板丈上, 丈）**以外のすべての列を「欠陥名」**として Arry_L(n+1,2) に入れる（L249-272）。
3. レコードは `[用途コード] LIKE '*<用途>*'` で絞る（L274）。
4. 各行の条件（L298-303）:
   - C1 = CheckCondition(材質列, 材質)
   - C2 = C1 And CheckCondition(**材質列, 材質**) ← コメントは「製造調質のチェック」だが材質を再判定しているだけで、**調質は判定されない**
   - C3 = CheckThickness(板厚下, 板厚上, 厚, **製造厚**)
   - C4 = CheckWidth(板幅下, 板幅上, 幅, **オーダー幅**)
   - C5 = CheckWidth(板丈下, 板丈上, 丈, **オーダー丈**)
   - C6 = CheckCondition(納入先, 納入)
   - 一致した行の管理番号を順に集める。
5. 一致したすべての管理番号について、結合配列から欠陥名の列の値（グレード）を Arry_L(,2) に入れる（L328-342）。**複数一致したときは後の番号で上書き**されるので、最後に一致した番号が残る。
6. `FalseSearch_L(別表1 結合配列, Arry_L)`: 別表1 の [欠陥名] 列 = 欠陥名 の行で、列見出し = グレード値 の列の内容を集め、(欠陥名, 内容, グレード) の配列にする（L52-83）。
7. `AddPagesArrN(TabAb, 結果)` で各欠陥名のタブを**追加**する（クリアしない）。AbnoName = 「異常グレード確認一覧」+一覧。
8. KousinL1 = `Q_S_QA_124_別表2` 全件の**先頭行の更新日**（L399-402）。
9. 先頭の欠陥を UF_DbDisplay で表示し、配列を返す。

#### 4.2.3 DBdisplay外観_S（QDB L462-675）

1. 薄板の判定: `If オーダー厚 < 8# Then AtuFlag = True`（L555）。False なら Upater = "該当板厚なし"。
2. `Q_S_QA_SRF_128_別表1` を `[用途コード] LIKE '*用途*'` で絞る。0件なら Upater = "該当用途ｺｰﾄﾞなし"。
3. 行を順に見て、**最初の一致**で決定する（L629-654）:
   - `CheckCondition(材質)`
   - かつ `ArrQuality(i, 表面処理コード) Like "*" & グレード & "*"`（グレード = 品質ｸﾞﾚｰﾄﾞ_表面処理）
   - 厚フラグ <> "" のとき: `Val(板厚1) >= Val(オーダー厚)` ならパターン1、`<` ならパターン2
   - 厚フラグが空ならパターン1
   - `Arry_S` は Long 型で、パターン列が数値でないと型不一致エラーになる（推測）
4. `Upater = Arry_S`。続いて FalseSearch_S を**先に呼んでから** `If Arry_S = 0` → "該当ﾊﾟﾀｰﾝなし"。
5. `基準反映(OutName, 表)`: 1行目（異常名）がボタン caption と一致する列について、2行目以降の **1列目** が "O"/"S"/"K"/"A" の行の値を SelectO/S/K/A に表示する（L683-695）。
   - FalseSearch_S が除外するのは管理番号・パターン・備考・更新日だけ（L10-11）。「条件」列が除外後の1列目にあることが前提（推測: 別表2 の列順は 管理番号, パターン, 条件, …）。
   - FalseSearch_S は `ReDim TempSize(1 To 5, …)` なので、**同じパターンの行が5行以上あると添字エラー**。
6. KousinL4 = 絞り込み後の先頭行の更新日。Sansyou1 に " Q-S-QA-SRF-128" が無ければ追加。
7. L609-618: `Exit Sub` の後に `L2.Visible=False` があり、実行されない（ローディング表示が残る）。

#### 4.2.4 DBdisplay歪（QDB L700-1384）。UF品質基準 手動検索だけで使う

- `AtuFlag = オーダー厚 <= 6#`（L803）。`VcFlag = Val(LotAcc(,VC)) = 1`（L805）。
- VcFlag のとき: SIKAODRNOW.accdb `仕掛` を `[受注番号]='<ｵｰﾀﾞｰ番号>'` で引く。VC_表/VC_裏 が 208PRNF/V325NNF/V311PNF のどれかなら useFlag（L814-846）。
- **A. AtuFlag**（L880-940）:
  - Toukyu = "4級"
  - Mimi1 = LoopArray(耳歪4級 全件, 耳歪4級 フィールド, オーダー板丈, オーダー厚)。結果 0 なら ""。
  - Naka1 = 同様に 中歪4級
  - Sansyou1 += " ﾌﾗｯﾄﾈｽ製品規格中耳歪(ｵｰﾀﾞｰ丈判定)"
  - `フラットネス = 3` のとき: Mimi2 = 耳歪3級、Naka2 = 中歪3級（**フィールド配列は中歪4級のものを流用**）、Toukyu = "3級"
  - Sansyou1 の空白区切りの語を Dictionary で重複除去
- **SkipA → useFlag**（L944-994）:
  - QA_022_01 全件で `板厚下 <= 製造厚 <= 板厚上` の**最初の行**の [1m当たり平坦度上限] を MG1 に表示
  - HaHiz1 = LoopArray(QA_022_01, …, オーダー板幅, オーダー厚)、TaHiz1 = 同（オーダー板丈）
  - VcFlag.Caption = "あり"、Sansyou1 += " QA-022-01" → Exit
- **B. Q_S_QA_DIM_122_B**（L997-1104）:
  - `[用途コード] LIKE '*用途*'` が0件なら skip122A へ
  - **絞り込み後の先頭行**の用途コードを "," で分け、Trim した値が用途と完全一致すれば VariableFound
  - VariableFound のとき:
    - 材質が4文字以上なら右3文字にする（L1059）
    - `[材質] LIKE '*材質*' AND [調質] LIKE '*調質*'` で絞る。0件なら **Exit Sub**
    - `板厚下<=オーダー厚<=板厚上` かつ `板幅下<=オーダー板幅<=板幅上` の最初の行で、HaHiz1 = [板幅]、TaHiz1 = [板丈]
    - Sansyou1 += " Q-S-QA-DIM-122_B"、中耳4欄 = "不問"
    - 一致の有無にかかわらず Exit
- **C. skip122A: Q_S_QA_DIM_122_A**（L1108-1205）:
  - `[用途コード] LIKE '*用途*' AND [調質] LIKE '*調質*'` が0件なら skipO へ
  - 材質を右3文字にする
  - 白銅の優先: `ArrQuality(材質) Like "*材質*"`、`板厚下<=製造厚<=板厚上`、`納入フラグ="白銅"`、`納入 Like "*ﾊｸﾄﾞｳ*"` の最初の行を ArrQ_H に入れる。MG1 = 1m当たり平坦度上限、Noun の背景色 = RGB(175,15,5)
  - 白銅でない場合: 同じ条件で `納入フラグ <> "白銅"` の**全行**を ArrQ_H に追加。MG1 は**最後の行**の値、Noun = 白
  - HaHiz1 = LoopArray(ArrQ_H, …, **製造幅, 製造厚**)、TaHiz1 = LoopArray(…, 製造丈, 製造厚)
  - 両方とも "" なら skipO へ。それ以外は Sansyou1 += " Q-S-QA-DIM-122" → Exit
- **D. skipO: Q_S_QA_DIM_122_C**（L1211-1291）:
  - `[調質] LIKE '*調質*'` が0件なら skipHX8 へ
  - 材質を右3文字にする
  - 全行を見る。条件: `製造材質 Like *材質*`、`用途コード Like *用途*`、`納入 Like "*" & DB納入先名 & "*"`、`Val(板厚下)<=Val(オーダー厚)<=Val(板厚上)`
  - パターン=1 なら DB の [オーダー板幅] ≥ lot のオーダー板幅、=2 なら ≤、空なら無条件
  - 一致したら HaHiz1 = 幅公差、TaHiz1 = 丈公差、Noun = RGB(255,255,153)
  - **最後に一致した行が採用される**
  - 中耳4欄 = "不問"、Sansyou1 += " Q-S-QA-DIM-122_O"（**一致がなくても付く**）→ Exit
- **E. skipHX8: Q_S_QA_DIM_122_D**（L1293-1378）:
  - `[製造材質] LIKE '*材質*'` が0件なら skipother（何もしないで終了）
  - 条件: `製造調質 Like Left(調質,1) & "*" & Right(調質,1)`、`用途コード Like *用途*`、`納入 Like *DB納入先名*`、`Val(オーダー板厚)=Val(オーダー厚)`
  - 等級列が空: 幅公差/丈公差を表示。Noun = RGB(0,102,204)、Sansyou1 += " Q-S-QA-DIM-122_D"、終了
  - 等級列に値あり: フラットネス 4/3 で ソリ4級/ソリ3級 を読み、**ArrQuality をループの途中で差し替える**（L1342/1349）。HaHiz1/TaHiz1 = LoopArray(ソリ, …, オーダー板幅/丈, オーダー厚)。Ittiflag は立たないので、差し替えた配列でループが続く（バグ）
  - さらに、ソリ表の列名は DBQP で "H…"/"T…" 形式（FindIndices）であり数値ではない。そのため LoopArray の `ReDim TempSize(1 To 0)` で**エラーになる可能性がある**（推測、要確認）
  - 最後に中耳4欄 = "不問"

#### 4.2.5 LoopArray（QDB L1386-1453）。サイズ表の引き方（DB 歪公差の中核）

- フィールド名が `IsNumeric` の列を、**列の順番どおり**に TempSize(1..n) に入れる（並べ替えない。列順が昇順である前提）。
- 行を順に見て、`Val(板厚下) <= ATU <= Val(板厚上)` の**最初の行**で、a を n から 1 へ降順に回す:
  - a=1（n=0）: `TempSize(1) >= ReturnSize And 0 < ReturnSize` なら最小列の値
  - a>1: `TempSize(a) >= ReturnSize And TempSize(a-1) < ReturnSize` なら列 a の値。区間は (前列, 当列] で、上端を含む
  - 当てはまらないまま a=1 を過ぎた場合（**ReturnSize > 最大列、または ReturnSize ≤ 0**）、`GoTo skipErr` で a = UBound となり、**最大列の値を返す**（L1433-1436。2023.8.22「DBにないサイズ(4500mm超え)」）
- 戻り値は `Format(値,"0.0")` を Single に代入したもの。セルが ""（Null 由来）なら型不一致エラーになる可能性がある（要確認）。
- 厚が範囲に入る行が無ければ 0。

#### 4.2.6 BhSearch（QDB L1454-1490）

材質を右3文字にし、`[材質] LIKE '*材質*'` で絞る。`調質 = TempT`（完全一致）の行の [バーコール下限値] を返す。**ループは途中で抜けないので最後の一致行**が採用される。一致がなければ 0。BhJudg は 0 なら "不要"。

#### 4.2.7 DBdisplayラテラル（QDB L1491-1696）

- テーブル: "日車" → `Q_S_QA_DIM_179_日車`、"白銅滑川" → `Q_S_QA_DIM_179_白銅滑川`。それ以外は TableName が "" になり、SQL エラーになる。
- 全件を読む（Filter なし。L1621）。
- **ArrQuality が Empty でも確認せずに UBound を取る**（L1625）ので、0件だとエラー。
- 条件: CheckCondition(材質)、CheckThickness(板厚下, 板厚上, 厚, **製造厚**)、CheckWidth(板幅下, 板幅上, 幅, **製造幅**)。
- 一致した行を HitArry に入れ、`LoopArray(HitArry, フィールド, 製造丈, 製造厚)` を返す。

### 4.3 DBQP（QP L69-801）。本番の歪公差

戻り値は `Kousa(行,1..9)`。列の意味:

| 列 | 内容 |
|---|---|
| 1 | ゲージ（1m当たり平坦度上限） |
| 2 | 幅歪（H列） |
| 3 | 丈歪（T列） |
| 4 | 中歪（中歪・耳歪 列 → 後で数値化） |
| 5 | 耳歪 |
| 6 | コメント（表示コメント） |
| 7 | 参照元（平坦度上限値参照元 + 等級表名） |
| 8 | 管理番号 |
| 9 | フラグ |

1. SIKALOTNOW の確認と Arrbox（TempH は常に ""）。ロットが無ければ MsgBox「SIKALOTNOWに該当ロットがありません_DBQP」で Empty を返す。
2. **検査番号の頭**（L166-177）:
   - `KenA = Left(検査No,1)`
   - "0" または "" なら `FindPriority(実績_設備ｺｰｽ)` を使う
   - それも "" なら MsgBox「検査Noがないファイルです_DBQP」で Empty を返す
3. `平坦度一覧QP追加` を Filter `"[検査番号の頭] LIKE '*" & KenA & "*' OR [検査番号の頭] =  ''"` で読む（L283）。UpdateDay = 先頭行の [更新日]。
4. 幅・丈の列: `HitHaba = FindIndices(フィールド,"H",オーダー板幅)`、`HitTake = FindIndices(…,"T",オーダー板丈)`。
5. 条件（L308-325。各条件は前の条件との AND の連鎖）:

| # | 判定 | DB列 | 比較値 |
|---|---|---|---|
| 1 | CheckCondition | 用途コード | 用途(SIKALOT) |
| 2 | 〃 | 製造材質 | 製造材質 |
| 3 | 〃 | オーダー材質 | オーダー材質 |
| 4 | 〃 | 製造調質 | 製造調質 |
| 5 | 〃 | オーダー調質 | オーダー調質 |
| 6 | CheckThickness | 板厚下/板厚上/オーダー板厚 | オーダー厚 |
| 7 | CheckWidth | 板幅下/板幅上/オーダー板幅 | オーダー板幅 |
| 8 | CheckWidth | 板丈下/板丈上/オーダー板丈 | オーダー板丈（2026.05.12 追加） |
| 9 | CheckCondition | 納入先 | **引数 delivery**（引当の納入先名称） |
| 10 | 〃 | 納入先コード | SIKALOT 納入先ｺｰﾄﾞ |
| 11 | 〃 | 取引先 | **引数 Good**（引当の取引先名称） |
| 12 | 〃 | 取引先コード | SIKALOT 取引先ｺｰﾄﾞ |
| 13 | 〃 | 送り先 | SIKALOT 送り先名 |
| 14 | 〃 | 送り先コード | SIKALOT 送り先ｺｰﾄﾞ |
| 15 | CheckSize | 巾取数幅分割数のこと | BOX設計_横割数(板) |

   **最終判定は `Conditions(1)…(13)` だけ**（L328-330）。14（送り先コード）と 15（巾取数）は**判定に効いていない**。原文:
   ```
   If Conditions(1) And Conditions(2) And Conditions(3) And Conditions(4) And Conditions(5) And _
   Conditions(6) And Conditions(7) And Conditions(8) And Conditions(9) And Conditions(10) _
   And Conditions(11) And Conditions(12) And Conditions(13) Then
   ```
   一致したときは管理番号（Long 変換）を MatchingArray に入れる。
6. 一致が0件なら **SkipQP** へ（下記 10）。
7. **優先**（L352-376）: 絞り込み後の行順で、一致した行のうち **納入先・納入先コード・取引先・取引先コード・送り先・送り先コード・巾取数幅分割数のこと のどれかが空でない最初の行**の番号で `NewArray(1)` を上書きする。元の NewArray(1) は失われ、優先行が重複して入ることもある。
8. Kousa(a,1..9) を設定する。(a,2) = `ArrQuality(i, HitHaba(1))`、(a,3) = `ArrQuality(i, HitTake(1))`、(a,5) = ""。
   - 2行以上あるとき: フラグに "重複可" を含む行のコメントを、1行目のコメントに含まれていなければ vbCrLf で連結する（L401-426）。
   - `HitHaba(2) = 9999`（幅超過）かつ参照元に "別表A" を含むなら、コメントに vbCrLf & "3054mm越えは3050mmと同様の公差" を追加（L432-437）。
9. `フラットネス` が 3 でも 4 でもなければ MsgBox「歪等級が3,4以外です」で **Exit Function（Empty を返す）**（L449-452）。
   中耳列 `Kousa(1,4)` の処理（L455-587）:
   - "不問" / "不問とする。" → SkipNoQuestionsAsked
   - `= 0` または "無き事" → SkipEndA（中=0, 耳=0）
   - "" → SkipT
   - それ以外の文字列 tMP:
     - `UpFlag`: ExtractNums(tMP,"以上") が "" でも " " でもない
     - `DowFlag`: `WhatDow <> "" And WhatUp <> " "`（L483。**WhatUp を見ている誤記**）
     - Result = Val(以上の数値) → DowFlag なら Val(以下の数値) で上書き
     - has3Grade/has4Grade = InStr("3級"/"4級")
     - 3級 かつ (Up or Dow) かつ CompareStrings(tMP, オーダー厚, Result) → 3級表 / 4級も同様 / 3級のみでカンマなし → 3級表 / 4級のみでカンマなし → 4級表
     - どれにも該当しなければ Kousa(1,4) = ""
     - **Up も Dow も無い**ときは、`中` / `耳` を含む部分から数値を取り出して Kousa(1,4)/(1,5) に入れ、skipEnd へ（L521-530）。**この分岐が等級表の参照より先にある**ため、「4級」とだけ書かれた行は表を引かず、中耳は "" になる（後で SIKALOT の指示値で補われる）。挙動として要確認。
     - 以上/以下あり・等級指定なし・CompareStrings 真 → フラットネスの等級で中耳表を LoopArray（オーダー板丈, オーダー厚）
   - TableN <> "" なら中耳表を LoopArray で引いて (1,4)/(1,5) に入れ、参照元に SearchS（"Q-E-001_品質規格:等級3" / "…等級4"）を追加する。
   - SkipT / SkipNoQuestionsAsked: (1,2) または (1,3) が "フラットネス品質等級" なら、ソリ3級/4級の FindIndices と SoriArray（オーダー厚）で (1,2)/(1,3) を上書きし、参照元に追加する（L599-629）。
10. **SkipQP**（QP の行に該当なし。L638-760）:
    - Kousa を (1,9) に作り直す
    - `製造調質 = "O"` → (1,6) = "別表Cに記載のないO材_判定不要"、(1,7) = "Q-S-QA-DIM-122_別表C"。Execution="" かつ UF1 表示中なら UF1.HizCome / DistortionFrame に表示して返す
    - `Len(製造調質)=3 And Left="H" And Right="8"` → "別表Dに記載のないHx8材_判定不要" / "Q-S-QA-DIM-122_別表D"
    - それ以外はフラットネス 3/4 で表を選ぶ。**3/4 以外だと TableName は L229 の "平坦度一覧QP追加" のまま、TableN/TableM は "" のまま**で処理が進む（要確認）
    - ソリ: FindIndices + SoriArray で (1,2)/(1,3)、参照元に追加
    - `6# >= オーダー厚` のとき: SIKALOT の指示 中歪 または 耳歪 が空なら、中耳表を LoopArray（オーダー板丈, オーダー厚）で引く
    - オーダー厚 > 6 のとき: (1,4)/(1,5) = ""、ソリを再計算（同じ結果）
11. **skipEnd**（L769-781）:
    - (1,1) "不問とする。" → ""
    - (1,4) "不問とする。"/"不問" → ""
    - (1,5) "不問" → ""
    - `(1,4) = ""` かつ SIKALOT `指示_中歪_高さ` <> "" → その値（耳も同様。**SIKALOT の指示値は最後の補完**）
    - Execution = "" なら UF品質基準 の Uatu(製造厚 "0.000")/Uhaba/Utake("0.0")/Uzai/Utyou/UYoucode/Uhyou/Noun(=delivery)/KousinL2/L3(=UpdateDay) を書く。**UF品質基準 が表示されていなくても暗黙にロードされる**

### 4.4 照合関数の正確な意味（QP L804-981）

**FindInList(tMP=DB側, List=探す値)**

- `tMP = ""` → True（空欄は全件該当）。
- tMP を "," で分けて各要素を Trim し、**左から順に**評価する。肯定条件は一致すると True にするだけで、途中で抜けない（後勝ちの OR）。≠ 条件で不一致になると即 False で終了。
  - `≠X`: X ≠ List なら **True を設定**（それまでの肯定条件が不一致でも True になる）。X = List なら False で終了。続けて X の形に応じて:
    - `*Y*` → List に Y を含めば False
    - `*Y` → 後方一致なら False
    - `Y*` → 前方一致なら False
    - 完全一致 → False
  - `*Y*` は部分一致、`*Y` は後方一致、`Y*` は前方一致、それ以外は完全一致。
  - 比較は Option Compare Binary（モジュールに指定が無い）なので、**大小文字・全角半角を区別**する。
- 例: `"52S,≠F52S"` に `"5083"` を渡すと True（≠ 要素が True にする）。Python でも同じにすること。

**CheckCondition**: `TmpString = ArrQuality(row,Column)`。空でなければ `StrConv(vbNarrow)`（全角 → 半角。カナを含む）、続けて `Replace("ｰ","-")`。**Criteria 側は変換しない。**

**CheckThickness / CheckWidth**: 下・上・値列の3つが全部 "" なら True。それ以外は `Val(下) <= Val(x) <= Val(上)`。上限が空だと Val("")=0 となり、x>0 では False。値列（Col3）は空判定だけに使う（2026.05.13 に条件から削除）。

**CheckWithin**: 空判定なしで `Val(下) <= x <= Val(上)`。両方空だと 0..0。

**CheckSize**: 空なら True。それ以外は `Val(OrderSize) >= Val(DB)`（Long）。

**FindIndices**（L985-1049）:
- 1行目のフィールド名で、先頭の文字が "H"/"T" かつ残りが数値の列を集める（Double）。
- 昇順に並べ替えるが、**入れ替え用の変数 tempIndex が Long** なので小数は切り捨てられる。
- a を上から下へ回す:
  - a=1: `Temp(0) >= R And 0 < R` → 最小列 ／ `Temp(1) >= R And Temp(0) < R` → 列1
  - a>1: `Temp(a) >= R And Temp(a-1) < R`
  - a=0（n=-1）に達したら、最大列＋ `(2)=9999`（範囲外。**R ≤ 0 のときも該当**）
- 一致した数値を `WidthLength & Temp(Amatch)` という文字列にして元の列名を探す（"H0500" のような書式だと一致しない）。
- 該当列が1つだけのときは常に 9999 になる。該当列が無いと Temp が未確保で実行時エラー。

**SoriArray**: `If FindIndices = 0 Then Exit Function`（0 を返す）。それ以外は `板厚下 <= ATU <= 板厚上` の最初の行の指定列を Format "0.0"。

**ExtractNumFromPart**: Condition を含む部分の、最初に現れる「数字の連続と小数点1個」。`IsNumeric(1文字)` は全角数字でも True になる（要確認）。

**CompareStrings**: "以上" があれば Lot厚 ≥ DB厚、なければ "以下" があれば Lot厚 ≤ DB厚、どちらも無ければ False。

### 4.5 DistArray / FormReflectDist（QP L1911-2055）

**DistArray**
- `HikiArrCreate(LotNo)`: SIKAHIKINOW `仕掛` `[ﾛｯﾄ番号]='LotNo'` の（受注_納入先名称, 受注_取引先名称）一覧。MainPath で取れなければ NextPath。
- 納入先名称が全部同じとき: 1件目の（納入, 取引）で DBQP を1回だけ実行する（**取引先が違っても1件目の値**）。
- 違う納入先があるとき: 引当ごとに DBQP を実行し、WidthLength(i) = (1,2) & "_" & (1,3)。
  - 公差の文字列が異なるなら Val して最小値をとる。**最小値が 0 なら SecondSmallest**（ソート後に最小値と異なる次の値。全部同じなら `CVErr(xlErrNA)` で、Double への代入時にエラー）。
  - 1件目の結果の (1,2)/(1,3) を最小値で上書きし、(1,6) に vbCrLf & "納入先による公差違い有,厳しい公差で上書き" を追加。
  - **ゲージ・中・耳は1件目の値のまま。**
  - 1件目の DBQP が Empty だと TopFormerArr の参照でエラー（要確認）。

**FormReflectDist**
- UF品質基準・UFMitsubishi がどちらも非表示のとき（UF1 に反映）:
  - MG: 数値なら Format "0.0"、そうでなければ ""。その後 `If UF1.MG = 0 Then ""`
  - HaHiz/TaHiz: 数値なら "0.0"、**数値でなければ文字列をそのまま表示**（"フラットネス品質等級" など）。0 なら ""
  - Naka/Mimi: 数値なら "0.0"、そうでなければ ""。0 なら ""
  - HizCome = コメント。末尾が "_" なら除き、"。" を vbCrLf に置換
  - DistortionFrame = "参照元:" & 参照元
- UF品質基準 表示中: MG1/HaHiz1/TaHiz1/Naka1/Mimi1（数値でなければ文字列のまま）、Sansyou1 = 参照元 & "_" & コメント（末尾の "_" を除去）。
- UF1 非表示・全ヵ所検査シートあり・UFMitsubishi 表示中: シートの保護を解除し、セル (5,13) = 幅歪、(5,15) = 丈歪。ロックしてフォントサイズ 8。
- `UF1.MG = 0` のような「テキストボックスの値 = 0」の比較は、値が "" のとき VBA の Variant 比較規則でどうなるか（エラーか False か）を**実機で確認する必要がある**。

### 4.6 DbReturnItem（QP L1176-1456）

- 前処理: UF1.HURIS = ""。
- `ラテフリ直角` を全件読む（QualityPath。テーブルが空なら PackingPath）。
- 条件（すべて AND）: 納入先 = Nou、取引先 = Tori、用途コード = YoutC、製造材質 = Zai、製造調質 = Tyou（いずれも CheckCondition）、CheckThickness(厚下, 厚上, 製造板厚, ATU)、CheckWidth(幅下, 幅上, 製造板幅, HABA)、CheckWidth(丈下, 丈上, 製造板丈, TAKE)。
- 該当行ごとに:
  - [フリー反り]="要" → ",フリー反り" を追加し、UF1.HURIS = [公差]
  - [ラテラルボー]="要" → ",ラテラルボー"
  - [直角度]="要" → ",直角度"
- 該当0件のとき:
  - `Nou Like "*ﾆﾎﾝｼﾔﾘﾖｳｾｲｿﾞｳ*"` かつ Additional に "ﾗﾃﾗﾙﾎﾞ-追記公差あり" が無い: **最後に評価した行の Conditions** から、
    - C4 False → "ﾗﾃﾗﾙﾎﾞｰ検査不要_製造材質"
    - C4 And C5 False → "…_ｻｲｽﾞ幅"
    - C4 And C6 False → "…_ｻｲｽﾞ丈"
    - 4つ目の分岐（"_ｻｲｽﾞ厚"）は到達しない
    - 条件番号と文言が対応していない（C5=調質、C6=厚）
  - `Tori Like "ﾅﾒｶﾜｱﾙﾐ(ｶ*"` のとき: `Nou Like "*K.T.N. CO. LTD*"` → "K.T.N ﾗﾃﾗﾙﾎﾞｰ検査不要"、`Nou Like "ﾆﾎﾝﾊﾂｼﾞﾖｳ*"` → "ﾆﾎﾝﾊﾂｼﾞﾖｳ ﾗﾃﾗﾙﾎﾞｰ検査不要"。LateralText = ""
- ラテラルボーの表示（UFMitsubishi 非表示時）:
  - Sumisei = Nou
  - Additional に "ﾗﾃﾗﾙﾎﾞ-追記公差あり" が無ければ、**先頭の該当行**の [公差] で決める: 数値なら Format "0.0"、"別表B" なら DBdisplayラテラル(日車)、"別表A" なら DBdisplayラテラル(白銅滑川)（背景をオレンジ RGB(255,153,0) に）、"等級" なら何もしない
- フリー反り: HizCome.Tag = 先頭該当行の [公差] & "mm以下"、flatnesText = "ﾌﾘｰｿﾘ要"、LGS/LGU/LGC/LGD/LGE/LGF.Caption = "ﾌﾘｰｿﾘ"、CBAdd(ATU)、HURIS = 先頭行の公差、HizCome.Caption = "ﾌﾘｰ反り" & Tag。
- UFMitsubishi 表示中: シート（WorkFlag "P1" → 全ヵ所検査、"P_H494" → H494検査）。保護を解除し、(1,15) = Nou & Zai、(1,13) = 公差（または別表）、フリー反りなら (6,29) = "ﾌﾘｰｿﾘ"。

### 4.7 その他のマスタ判定

- **MaterialsUpperPart**（QP L1674-1783）:
  - Lot情報履歴（**見出し行を含む** 1行目から）で [ｵｰﾀﾞｰNo] = OrdNo（"'" を除去）の最初の行
  - 包装仕様NO が C_TARGET_HOSONOS に無ければ ""
  - 梱包保護材の各行で:
    - `(備考="当板用") = CheckBikou` でなければ次の行へ
    - 包装仕様書・材質・調質・用途コード（CheckCondition）、板厚・板幅・板丈（範囲）
  - 最初の一致の [使用保護材] を返す。無ければ "一致なし"
- **CheckEquipmentDb**（L2084-2190）:
  - `[設備名] LIKE 'line'`（ワイルドカードなし）が0件なら "可"
  - 条件: CheckCondition(製造材質)、CheckWithin（厚・幅・丈・単重の MIN–MAX）
  - 0件一致のとき、**最後に評価した行**の Conditions で: C2 False → "オーダー厚 範囲外"（材質不一致でもこれになる）／ C3 → "C_オーダー幅 範囲外"（"C_" も原文どおり）／ C4 → "オーダー丈 範囲外"／ C5 → "単重 範囲外"
  - 一致したら、空でない [作業可否] のうち**最後のもの**。全部空なら ""
- **GaichuPackingNo**（L2250-2269）:
  - 外注条件の行で 有効=True のもの
  - 明細の条件番号ごとに OR、番号どうしは AND
  - 明細が無い条件セットは不成立
  - 成立したセットのうち優先順位が最小のもの（**同値なら先の行**。比較は `<`）の [戻り値]
  - JudgeOne:
    - "=": 両方が数値なら Round(,3) で比較、それ以外は文字列の完全一致
    - "Like": VBA Like
    - "含む": InStr（判定値が空なら False）
    - ">=" / "<=": 両方が数値のときだけ
    - それ以外の方法は False
  - 対象行は Hist の中で `CStr(Hist(i,R_Lot)) = LotNo` の**どれか1行**で成立すればよい
- **IsGaugeRequired**（L2473-2511）:
  - 重量 = `Weight(厚,幅,丈)` = `Application.RoundDown(厚*幅*丈*2.68/10^6, 3)`（OutTolerance L813-817）
  - 有効な行で、値が入っている条件をすべて満たす: 重量_超（>）、丈_以上（>=）、丈_未満（<）、丈_超（>）、幅_未満（<）、材質（=）、調質（=）、コース含む（InStr）
  - nCond > 0 の行が1つでも成立すれば True
- **ShiftTime**: `時間用` の全件の**先頭4行**を (直, 開始, 終了) で返す。4行未満なら添字エラー。
- **StraightEdge**: `[ライン名]='LineName'` の、先頭行の [長さ下] と最終行の [長さ上]。

### 4.8 補助の判定・変換（H）

- **TimeDesignation**（L1027-1040）。原文:
  ```
  If (5945 < 分 & 秒) Then ... ElseIf (1 > 分) And (45 >= 秒) Then
  ```
  文字列連結を数値と比較しているので、**毎時 :59:46〜:59:59 と :00:00〜:00:45 に True** になる（DB 再生成中とみなし NextPath を使う）。True のたびに NextPath\Access.log へ `分:秒_時間ﾁｪｯｸ` を書く。
- **FileDesignation**（L1042-1053）: `MainPath & "PLEASE WAIT.txt"` があれば False（ログ "PLEASE WAIT.txt_ﾁｪｯｸ*在る" を書く）、無ければ True。
- **Maker / Conv**（L380-407）。計測器の文字列解析:
  - `Maker`: `Len>=3` のとき `StrConv(UCase(Left(Trm,2)),vbNarrow)` が "DT" → "DT"、"#L" → "#L"、"#〆" → "#〆"、それ以外は ""
  - `Conv`: "DT" → `Mid(Trm,9,11)`（例 "DT10002+0000020.712M" → "0000020.712"）、"#L" と "#〆" → `Mid(Trm,3)`、それ以外はそのまま（元の文字列。全角は維持）
- **FindPriority**（L1455-1491）:
  - コース文字列を空白で分け、**右から**見て CM1/CM2/CM3/2HI のどれかに完全一致した最初の語
  - 無ければ全体に "HOT" があるか
  - 対応: CM1→L、CM2→N、CM3→K、2HI→T、HOT→H。どれも無ければ ""
- **numExtract**: "/" があればその前だけを対象に、`[0-9.]` の文字だけを連結する（複数の "." も残る）。**ConvertToHalf** は全角数字を半角にしてから numExtract。
- **CircledNumber**: 1–20 → ChrW(9311+n)（①〜⑳）、21–35 → ChrW(12860+n)（㉑〜㉟）、36–50 → ChrW(12941+n)（㊱〜㊿）。範囲外は数値の文字列。
- **SecondSmallest**: ソートして、隣と異なる最初の要素（2番目に小さい異なる値）。無ければ CVErr(xlErrNA)。
- **AllElementsSame**: 先頭要素の文字列と全要素を `<>` で比較。
- **merge_sort2**: 指定列で安定ソート（昇順）。比較は Variant の `<` / `<=`。

---

## 5. 外部依存

### 5.1 Access ファイルとテーブル（このグループのコードで使うもの）

| ファイル | 場所 | テーブル | 使用列 | 呼び出し元 |
|---|---|---|---|---|
| SIKALOTNOW.accdb | MainPath→NextPath | 仕掛 | §3.2 | Arrbox 経由: QDB 外観_H/S, 歪, ラテラル, QP DBQP／直接: Q 表面品質基準 |
| SIKATWOYEAR.accdb | NextPath | 仕掛 | 品質ｸﾞﾚｰﾄﾞ_表面処理, 製造板厚/幅/丈, 製造材質, 製造調質, 用途ｺｰﾄﾞ | Q 表面品質基準（SIKALOTNOW が空のとき） |
| SIKAODRNOW.accdb | 決定した TempPath | 仕掛 | 受注番号, VC_表, VC_裏, 納入先名称, 包装仕様NO | QDB DBdisplay歪 |
| SIKAHIKINOW.accdb | MainPath→NextPath | 仕掛 | ﾛｯﾄ番号, 受注_納入先名称, 受注_取引先名称 | QP HikiArrCreate |
| 品質基準.accdb | QualityPath | Q_S_QA_124_別表2 | 材質, 用途コード, 板厚下, 板厚上, 厚, 調質, 納入先, 板幅下, 板幅上, 幅, 板丈下, 板丈上, 丈, 備考, 更新日, 管理番号, ＋欠陥名の列（飛込み, キズ, 鋳造剥離AL圧込, AL付着跡, 汚れボタ落ち, カミコミ跡ショック跡 …） | 外観_H |
| 〃 | 〃 | Q_S_QA_124_別表1 | 欠陥名, グレード名の列 | 外観_H |
| 〃 | 〃 | Q_S_QA_SRF_128_別表1 | 表面処理コード, 材質, 用途名, 用途コード, 厚フラグ, 板厚1, パターン1, パターン2, 備考, 更新日 | 外観_S |
| 〃 | 〃 | Q_S_QA_SRF_128_別表2 | 管理番号, パターン, 条件, 備考, 更新日, ＋異常名の列 | 外観_S |
| 〃 | 〃 | 耳歪4級 / 中歪4級 / 耳歪3級 / 中歪3級 | 板厚下, 板厚上, 平坦度上限値参照元, 用途コード, 備考, **数値名の列（サイズ）** | 歪, DBQP |
| 〃 | 〃 | ソリ4級 / ソリ3級 | 板厚下, 板厚上, **"H<数値>" / "T<数値>" の列** | DBQP, 歪(D) |
| 〃 | 〃 | QA_022_01 | VC, 1m当たり平坦度上限, 板厚下, 板厚上, 厚, 備考, 数値名の列 | 歪 |
| 〃 | 〃 | Q_S_QA_DIM_122_A | 平坦度上限値参照元, 用途コード, 板厚下, 板厚上, 厚, 1m当たり平坦度上限, 備考, 納入フラグ, 材質, 調質(Filter), 数値名の列 | 歪 |
| 〃 | 〃 | Q_S_QA_DIM_122_B | 材質, 調質, 用途コード, 板厚下, 板厚上, 厚, 板幅下, 板幅上, 幅, 板幅, 板丈, 備考 | 歪 |
| 〃 | 〃 | Q_S_QA_DIM_122_C | 用途名, 製造材質, 用途コード, 板厚下, 板厚上, オーダー板厚T, オーダー板幅, 幅公差, 丈公差, 備考, 納入先名, 調質, パターン | 歪 |
| 〃 | 〃 | Q_S_QA_DIM_122_D | 用途名, 製造材質, 製造調質, 用途コード, オーダー板厚, 幅公差, 丈公差, 備考, 納入先名, フラットネス品質等級 | 歪 |
| 〃 | 〃 | 平坦度一覧QP追加 | 平坦度上限値参照元, 用途コード, 板厚下, 板厚上, 1m当たり平坦度上限, 納入先, 納入先コード, 送り先, 送り先コード, 取引先, 取引先コード, 巾取数幅分割数のこと, 検査番号の頭, 板幅下, 板幅上, 板丈下, 板丈上, オーダー材質, オーダー調質, 製造調質, 製造材質, 用途名, 表示コメント, 管理番号, オーダー板厚, オーダー板幅, オーダー板丈, 中歪・耳歪, フラグ, 更新日, **"H…"/"T…" の列** | DBQP |
| 〃 | 〃 | Q_S_KE_ALL_053 | 材質, 調質, バーコール下限値, 引張強さ(Mpa) | BhSearch |
| 〃 | 〃 | Q_S_QA_DIM_179_日車 / Q_S_QA_DIM_179_白銅滑川 | 材質, 調質, 納入先名, 取引先名, 用途コード, 板厚下, 板厚上, 厚, 板幅下, 板幅上, 幅, 備考, 管理番号, 数値名の列 | ラテラル |
| 〃（無ければ PackingPath） | QualityPath→PackingPath | ラテフリ直角 | 管理番号, 納入先, 取引先, 用途コード, 製造材質, 製造調質, 厚下, 厚上, 幅下, 幅上, 丈下, 丈上, フリー反り, ラテラルボー, 直角度, 製造板厚, 製造板幅, 製造板丈, 公差 | DbReturnItem |
| 〃 | QualityPath | 外注条件 | 管理番号, 条件名, 戻り値, 優先順位, 有効 | GaichuPackingNo |
| 〃 | 〃 | 外注条件明細 | 条件管理番号, 条件番号, 対象項目, 判定方法, 判定値, 有効 | 〃 |
| 〃 | 〃 | ゲージ判定条件 | 有効, 重量_超, 丈_以上, 丈_未満, 丈_超, 幅_未満, 材質, 調質, コース含む | IsGaugeRequired |
| 梱包資材マスタ.accdb | PackingPath | 梱包保護材 | 管理番号, 包装仕様書, 用途コード, 材質, 調質, 板厚下, 板厚上, 板厚, 板幅下, 板幅上, 板幅, 板丈下, 板丈上, 板丈, 使用保護材, 備考 | MaterialsUpperPart |
| 伝送用ファイル.accdb | QualityPath→PackingPath | 時間用 | 直, 開始, 終了 | ShiftTime |
| 〃 | 〃 | 校正ストレートエッヂ | ライン名, 長さ下, 長さ上 | StraightEdge |
| 〃 | 〃 | 設備仕様 | 管理番号, 設備名, 製造材質, オーダー厚MIN/MAX, オーダー幅MIN/MAX, オーダー丈MIN/MAX, 単重MIN/MAX, 作業可否, 備考 | CheckEquipmentDb |

### 5.2 SQL 文（すべて）

実際に発行される SQL は全部同じ形（SQL_追加 L293）:

```
SELECT * FROM <TableName>
```

絞り込みはすべて ADO の `Recordset.Filter`（L296）。下表に原文を載せる（`&` は連結）。

| 行 | テーブル | Filter（原文） |
|---|---|---|
| AccHIki L43 | 仕掛 | `"[ｵｰﾀﾞｰ板丈]<>0"`（Arrbox "板" — ただし Fields 取得では使われない） |
| AccHIki L45 | 仕掛 | `"[ﾛｯﾄ番号]='" & LotNo & "'"` |
| Q L1757 | 仕掛 | `"[ﾛｯﾄ番号]='" & LotNo & "'"` |
| QDB L274 / L599 / L1017 | 別表2 / SRF128別表1 / DIM_122_B | `"[用途コード] LIKE " & "'*" & 用途 & "*'"` |
| QDB L280, 399, 614, 897, 908, 918, 925, 961, 1337, 1621 | 各テーブル | `""`（全件） |
| QDB L835 | 仕掛(SIKAODRNOW) | `"[受注番号]='" & 受注番号 & "'"` |
| QDB L1060 | DIM_122_B | `"[材質] LIKE " & "'*" & 材質 & "*' AND  [調質] LIKE " & " '*" & 調質 & "*'"` |
| QDB L1130 | DIM_122_A | `"[用途コード] LIKE " & "'*" & 用途 & "*' AND  [調質] LIKE " & " '*" & 調質 & "*'"` |
| QDB L1234 | DIM_122_C | `"[調質] LIKE " & "'*" & 調質 & "*'"` |
| QDB L1314 | DIM_122_D | `"[製造材質] LIKE " & "'*" & 材質 & "*'"` |
| QDB L1480 | Q_S_KE_ALL_053 | `"[材質] LIKE " & "'*" & tempZ & "*'"` |
| QP L283 | 平坦度一覧QP追加 | `"[検査番号の頭] LIKE " & "'*" & KenA & "*' OR [検査番号の頭] = " & " '" & "" & "'"` |
| QP L469, 552, 615, 685, 707, 745, 1233, 1808 | 各 | `""` |
| QP L1853 | 校正ストレートエッヂ | `"[ライン名]='" & LineName & "'"` |
| QP L1890 | 仕掛(SIKAHIKINOW) | `"[ﾛｯﾄ番号]='" & LotNo & "'"` |
| QP L2121 | 設備仕様 | `"[設備名] LIKE " & "'" & line & "'"` |
| QP L2435 / L1641 / L1666 | 外注条件・明細・ゲージ・梱包保護材 | `""` |

- 値は文字列連結で埋め込んでいるので、値に `'` を含むと Filter が壊れる（エスケープなし）。
- **書き込み SQL（INSERT/UPDATE/DELETE）はこのグループには無い。** Access は読み取り専用（`adOpenStatic, adLockReadOnly`、クライアントカーソル）。

### 5.3 接続の開閉・ロック・トランザクション

- `GetAdoArr`（SQL_追加 L458-470）は、**1回の呼び出しごとに Connection を開き → Recordset を開き → 閉じる**。GetFieldsArr と GetRecordsArr を続けて呼ぶと2回接続する。
- DBQP 1回でおおよそ 10〜20 回の開閉になる（推測）。引当が複数あるとその倍数。
- 接続文字列: Provider `Microsoft.ACE.OLEDB.12.0`、`Jet OLEDB:Database Password = "nlm"`（L250-255）、`CursorLocation = adUseClient`。
- トランザクションもロック指定も無い（ReadOnly だけ）。
- EOF で Exit Function するときは Close を呼ばない（オブジェクトがスコープを抜けて解放される）。

### 5.4 ADO Filter 依存（Python 移植で注意）

- `LIKE '*x*'` は前後の `*` だけが許される（ADO の制約）。
- 大小文字の区別は ADO Filter 仕様による（推測: 区別しない。要確認）。
- `[検査番号の頭] = ''` が **Null の行に一致するか**は要確認。Access の空欄は通常 Null で、一致しない可能性がある。そうだとすると、空欄の行は L283 で除外されている可能性がある。
- `AND` と `OR` を混ぜると制約がある。L283 は OR だけなので問題ない。

### 5.5 Excel シート・CSV・ファイル

| 対象 | 読/書 | 箇所 |
|---|---|---|
| Lot情報履歴 シート | 読 | Q L62-75, L457-470、QP MaterialsUpperPart L1676-1716、GaichuPackingNo（引数 Hist） |
| 保存用 / 耳用 シート | 読 | FK 全体 |
| 全ヵ所検査 / H494検査 シート | 書（セル (1,13), (1,15), (6,29), (5,13), (5,15)。Unprotect する） | QP L1413-1455, L2043-2054 |
| WorkerTime_<ライン> シート C1 | 読 | IsOver12Hours |
| Sheet1 | 読書 | H 半角へ変換（Test.frm からだけ呼ばれる） |
| `C:\VBAWAV\*.wav` | 再生 | H 音楽系（異常値.wav, chimes.wav, chime2.wav, 確認作業.wav, ナンバーNO.wav, tada.wav） |
| `QualityPath\品質写真\<サブフォルダ>\*.jpg` | 読 | H AddPagesForSubFolders / LoadImagesIntoMultiPageFromTemp |
| `MainPath\PLEASE WAIT.txt` | 存在確認 | FileDesignation |
| `NextPath\Access.log` | 追記 | TimeDesignation / FileDesignation |
| `HistoryPath\<ライン>\<ライン>_Log\Measured.log` と `ThisWorkbook.path\Access.log` | 追記 | LogOut / LogOutCombo |
| PDF 出力先 | 書（既存のファイルは Kill してから） | PDFCREATE |

品証検証.csv と Test.csv は、このグループでは**シートの表示・非表示の名前一覧**（H L586, L599, L630）に出てくるだけで、判定ロジックからは参照されない。

---

## 6. 管理画面（UF管理）の機能

- **誰でも開ける。** UF1 の CommandButton7（`UFsizeSave(UF1)` の後に `UF管理.Show vbModeless`。UF1.frm L3048-3051。UF管理.Show は L3050）。パスワード・権限確認は無い。
- **マスタの編集機能は無い。** 読むのはブック内の「保存用」「耳用」シートだけで、書き込みは `HistoryLoad`（SaveLoad 側）経由で UF1 を復元すること。
- 機能:
  1. CB1 に保存済みの LotNo（B列）を並べる（Initialize。耳付きチェック `MIMI` の切替で耳用/保存用を切り替える）。
  2. 同じ LotNo が複数あるとき（Count>1）: NewCB に該当行の A列（検査開始終了の時刻）を並べ、赤枠・黄背景にして MsgBox「同LotNoあり、作業時間を選択してください」。NewCB で選ぶと CB1.ListIndex を合わせる。
  3. 選択行の表示（通常品。保存用シートの列番号）:

| 列 | 表示先 |
|---|---|
| 1 | 時間 |
| 3〜5 | 寸法 |
| 6〜11 | 公差 |
| 12〜21 | S 測定 10 項目 |
| 24〜33 | U |
| 36〜45 | C |
| 48〜57 | D |
| 60〜69 | E |
| 22/23, 34/35, 46/47, 58/59, 70/71 | 定盤/ゲージ歪（"__" なら表示しない） |
| 110/111 | F の歪（**FGH の判定が列110を見ている誤記**。L232） |
| 72〜76 | ラテラル |
| 77 | 検査員 |
| 83 | 検査頻度（WorkJudgment） |
| 84 | "〆" ならフリー反り表示（H2="フリー反り"） |
| 85〜89 | 外観（"A_…,B_…" を分解） |
| 113 | 6回目の外観 |
| 90〜94 | BH |
| 95〜99 | 直角度 |
| 100〜109 / 112 / 114 | 6回目（F） |

     - 歪の値 "A:x,B:y" を A面/B面 に分解する（2023.6.14）。
     - 検査番号の表示（列83 の先頭・末尾の文字）: 「H…数」→ 1〜5、「H…H」→ S/S1/2/E1/E、「末尾1」→ S/2/E、「末尾2」→ 2-S/2-E/1-E、「末尾3」→ 3-S/3-E/2-2/1-S/1-E、「末尾4」→ 4-S/4-E/3-2/2-2/1-S/1-E（6回目あり）。
  4. 耳用: 巾1〜6・丈1〜4 のレイアウト。列: 92 = 検査員、98 = 作業選択、26/27, 42/43, 58/59, 74/75, 90/91 = 歪、99〜103 = 外観、104〜108 = BH。98 が "極厚" または空なら S/U/2/D/E 表示。
  5. 呼び出し（CommandButton1）: UF1 に別の Lot が入っていたら中止。それ以外は HistoryLoad で UF1 に復元する。
- 行の対応は `b + 2`（b = CB1.ListIndex、1行目は見出し）。シートの行順と CB1 の並びが一致していることが前提。

---

## 7. 補助モジュールの汎用関数（Python で共通ユーティリティにすべきもの）

参照数は、補助モジュール以外での呼び出し行数（コメント行を除いた概数）。

| 関数 | 行 | 参照数 | 内容 | Python 化の方針 |
|---|---|---|---|---|
| WsCheck | 442 | 186行/28ファイル | ブック内のシートの存在 | データストア（テーブル/ファイル）の存在確認 |
| AddWS / ShJuge | 466 / 485 | 66 / 4 | シートを作り直して非表示（既存は**削除**して左端に追加） | 一時テーブルの作り直し（既存データ消去の意味に注意） |
| UR / UC | 500 / 520 | 65 / 34 | 最終行・最終列（n=0: SpecialCells(xlLastCell)、副作用で `sh.Unprotect`。n>0: End(xlUp)/End(xlToLeft)。空なら 1） | シート相当データの行数・列数 |
| ReadCheck | 454 | 63/15 | ファイル(1)・フォルダ(2)の存在（On Error Resume Next） | `os.path.isfile/isdir` ＋例外を False に |
| PathCheck / GetFileSystemType | 409 / 420 | 5 | 同上（FSO） | 同上 |
| TimeDesignation / FileDesignation | 1027 / 1042 | 各15/8 | DB 再生成中の判定（§4.8） | **秒単位で同じ時間窓を再現**（ログ出力も） |
| IsEmptyArray | 1132 | 37/12 | UBound でエラーなら空、UBound(a,1) < 0 なら空 | `not seq` 相当（Empty と 0 要素を区別しない） |
| Conv / Maker | 380 / 393 | 49/6、7/5 | 計測器入力の文字列解析 | **判定結果に直結。厳密に移植** |
| ConvertToHalf / numExtract | 1105 / 1117 | 4 | 全角数字を半角にし、数字と "." だけ残す | 共通の文字列ユーティリティ |
| FindPriority | 1455 | 5/3 | コース → 検査番号の頭（L/N/K/T/H） | 共通 |
| LogOut / LogOutCombo | 1147 / 1182 | 32 / 24 | 計測ログ（"*UFKen\|名前\|値\|ライン\|検査員"。A/B面は "[A] "） | ログ仕様の互換（区切り文字 "\|" は 2026.05.14 から） |
| ShowProgress / HideProgress / ProgressBar_count | 667 / 685 / 694 | 26 / 8 / 54 | 進捗表示 | UI 層（業務ロジックには不要） |
| 音楽*（MusicTrue 21, 音楽合格 22, 音楽不合格 23, 音楽対角 10 …） | 244-379 | – | 効果音。UF1.Tag="STOP" で鳴らさない。全ヵ所/キャスク/校正のときは STOP を無視（Line2） | クライアント側の音声（STOP の扱いを含めて再現） |
| WindowMenuButtonDel / On | 231 / 238 | 16 / 3 | フォームの × ボタンを消す・戻す | UI |
| HiddenSheets / DisplaySheets / KenSheets / VisibleSheets | 579-659 | 14 / 6 / 1 / 4 | シートの表示制御。HiddenSheets("x") は "Daily_" "WorkerTime_" "検査員_" "作業保存_" シートを**削除** | Web では不要（削除処理の業務的な意味は要確認） |
| PDFCREATE | 717 | 6 | シートの1ページ目を PDF にし、最大5秒待つ | 帳票出力（別設計） |
| NewFileName / GetNewFileName | 743 / 753 | 6 | フォルダ内で更新日時が最新のファイル名 | 共通 |
| merge_sort2 | 932 | 3 | 2次元配列の列で安定ソート | `sorted(..., key)`。Variant 比較の互換に注意 |
| IsBookOpened / BookSaClose / AllClose | 826 / 804 / 815 | 13 / 3 / 2 | ブックが開いているか（Append でオープンを試す）・保存して閉じる | ファイルロックの確認（要再設計） |
| CopyRangeAsHighResolutionImage / CreatePictureFromClipboard | 1328 / 1361 | 5 / 4 | セル範囲の画像化 | 帳票の画像化（別設計） |
| LoadImagesIntoMultiPageFromTemp / AddPagesForSubFolders / AddPagesArrN | 1493 / 1531 / 1569 | 5 / 3 / 2 | 品質写真ページ・異常タブの追加（**ページは追加するだけでクリアしない**） | UI |
| AllElementsSame / SplitWidthLength / SecondSmallest | 1623 / 1636 / 1653 | 2 / 1 / 3 | DistArray の補助 | 公差計算ユーティリティ |
| CircledNumber | 1205 | 2 | 丸数字 | 共通 |
| GetOneDimData | 538 | 1 | バイナリで読み **ANSI(CP932) を Unicode に変換**し、vbCrLf で分割（`ReDim buf(LOF(1))` はファイル番号1固定のバグ、末尾に余分な1バイト） | CP932 読込ユーティリティ |
| ReadFile | 1424 | 1（Test） | FSO ReadAll → ADODB.Stream（Charset 指定）で再エンコード | – |
| Any_Click | 186 | 1（UF校正） | 座標をクリック（1920×1080 なら 939,596、それ以外は 778,508） | Web では不要 |

QP の照合エンジン（FindInList / CheckCondition / CheckThickness / CheckWidth / CheckWithin / CheckSize / FindIndices / LoopArray / SoriArray / ExtractNums / CompareStrings）も QDB と QP で共通に使われている。**「マスタ照合ライブラリ」として1つにまとめ、単体テストで VBA と同じ結果になることを保証すべき**。

---

## 8. ハードコードされた業務ルール（全列挙）

1. 厚のしきい値:
   - 外観 厚板 = オーダー厚 ≥ 8（QDB L184）、薄板 = < 8（L555）
   - 中耳歪の対象 = オーダー厚 ≤ 6（QDB L803、QP L704）、旧版は ATU ≤ 6（Q）
   - UF品質基準 のページ選択 = UF1.ATU ≥ 4 で厚板ページ（UF1 L3082、UFMitsubishi L243）
   - UF1 から外観_H を呼ぶのは FlagAtu ≥ 8（UF1 L3110）、UFMitsubishi は > 6（L262）
   - **しきい値が場所ごとにばらばら。**
2. VC 指定銘柄 `208PRNF`, `V325NNF`, `V311PNF` で QA-022-01 を適用（Q L489, QDB L844）。SIKALOT の [VC] = 1 で VC 材（QDB L805）。
3. 材質が4文字以上なら右3文字にしてあいまい検索（F52S→52S）（QDB L1059, L1143, L1239, L1313, L1479）。
4. 白銅: DIM_122_A の 納入フラグ="白銅" かつ 納入先 Like "\*ﾊｸﾄﾞｳ\*" を優先（QDB L1150）。
5. 調質 "O" で QP に該当なしなら「別表Cに記載のないO材_判定不要」、"H?8"（3文字）なら「別表Dに記載のないHx8材_判定不要」（QP L643-662）。
6. 歪の等級は 3 または 4 だけ。それ以外は MsgBox「歪等級が3,4以外です」（QP L449-452, L602, L732）。
7. 検査番号の頭: 検査No の1文字目。"0" や空のときはコース最終圧延から CM1→L, CM2→N, CM3→K, 2HI→T, HOT→H（H L1458-1488）。規格 Q-S-QA-ALL-088 に基づく。
8. QP の優先: 納入先・取引先・送り先とそのコード、巾取数のどれかが記入された行を最優先（QP L358-361）。
9. 「重複可」フラグの行はコメントを連結（QP L408）。
10. 幅超過（FindIndices 9999）かつ参照元 "別表A" → 「3054mm越えは3050mmと同様の公差」（QP L432-436）。
11. 中耳の文言: "不問"、"不問とする。"、"無き事"、"0"（QP L457-465, L773-776）。
12. 等級表を示す文字列: "3級"、"4級"、"以上"、"以下"、"中"、"耳"（QP L480-528）。参照元の文言: "Q-E-001_品質規格:等級3" / "…等級4"。
13. "フラットネス品質等級" という文字列が幅歪・丈歪の列にあれば、ソリ等級表で上書き（QP L600）。
14. SIKALOT の指示 中歪/耳歪 は最後の補完（QP L778-779）。
15. 複数引当で納入先が違うとき、幅・丈の厳しい公差で上書き。0 は無視して次に小さい値（QP L1953-1974）。
16. ラテラル公差: [公差] が数値なら直接、"別表B" → 日車表、"別表A" → 白銅滑川表、"等級" → 何もしない（QP L1355-1367）。"ﾗﾃﾗﾙﾎﾞ-追記公差あり" があれば上書きしない。
17. ラテラル不要の特例: 納入先に "ﾆﾎﾝｼﾔﾘﾖｳｾｲｿﾞｳ" を含む、取引先が "ﾅﾒｶﾜｱﾙﾐ(ｶ…" で始まる（K.T.N. CO. LTD / ﾆﾎﾝﾊﾂｼﾞﾖｳ）（QP L1276-1306）。
18. フリー反り要 → 中歪ラベル6つを "ﾌﾘｰｿﾘ" に、公差 & "mm以下"（QP L1393-1409）。
19. 蓋アングルの包装仕様書: 1P0001, 1P0104, 1P0110, 1P0115, 1P0116, 1P0118, 1P0119, 1P0120（QP L3）。備考 "当板用" で分岐（QP L1757）。
20. 重量 = 厚×幅×丈×2.68/10^6 を小数3桁で切り捨て（OutTolerance L816。IsGaugeRequired で使う）。
21. DB 再生成の時間窓: 毎時 :59:46〜:00:45、フラグファイル "PLEASE WAIT.txt"（H L1033-1052）。
22. 設備仕様: 設備名が登録されていなければ "可"（QP L2124-2126）。
23. 旧版の厚板グレード・平坦度・表面パターン・グレード表（Q 全体。§4.1）。ﾅﾒｶﾜ判定の文字列 "ﾅﾒｶﾜｱﾙﾐ(ｶ"、"ﾅﾒｶﾜｱﾙﾐ(ｶ ｲﾀｺｺｳｼﾞﾖｳ"。
24. ライン名 → ログフォルダ: L-1, LVC, HVC, ﾊﾞﾗﾝｻｰ→バランサー, ﾄｯﾄ→トット, Test, その他→丸徳（H L1155-1163）。
25. 効果音のファイル名（H L266-379）。
26. UF品質基準 の入力制限（英数字のみ・大文字化）（FQ L256-332）。

---

## 9. 入出力項目

### 9.1 入力

- UF1: Lot, KenNo.Caption, DBflag.Text（"" なら MainPath、値があれば NextPath に切替）, ATU, Hab, Tak, Zai, Tyu, LineName / LineNo, AorB, Tag("STOP"), Additional.Caption, HizCome, Test などのラインのチェック。
- UFMitsubishi: LotBox, KenNo, WorkFlag.Caption("P1"/"P_H494")。
- UF品質基準: T1（Lot）, PathText, LineName, FlagAtu, B1〜B11.Caption（異常名: 例 "トビコミ" など。caption の実値は frx 内で不明）。
- SIKALOTNOW / SIKAHIKINOW / SIKAODRNOW の各列（§3.2, §5.1）。
- 品質基準.accdb の各マスタ。
- ローカルシート: Lot情報履歴, 保存用, 耳用, WorkerTime_<ライン>。

### 9.2 出力

- DBQP / DistArray の戻り値 Kousa: §4.3 の 9 列。
- UF1:
  - MG, HaHiz, TaHiz, Naka, Mimi（Format "0.0"、0 なら空）
  - HizCome.Caption（コメント。"。"→改行）
  - DistortionFrame.Caption（"参照元:" & …）
  - LateralText（＋背景オレンジ）, Sumisei.Caption, Additional.Caption, HURIS, flatnesText, LGS〜LGF.Caption, CB 群（CBAdd）
- UF品質基準:
  - MG1 / HaHiz1 / TaHiz1 / Naka1 / Mimi1 / Mimi2 / Naka2 / Toukyu / VcFlag.Caption
  - Sansyou1（参照元）, Noun（＋背景色の意味: 白銅=RGB(175,15,5)、O材表=RGB(255,255,153)、HX8表=RGB(0,102,204)、通常=白、QP 反映時=&H80000000）
  - Uatu…Uhyou, Upater, SelectO/S/K/A, AbnoName/AbnoLabel/gradeTB/AbnoTB, TabAb のタブ, KousinL1〜L4（更新日）
  - ATU/HABA/TAKE/Zai/TYO/YoutC/Noun（外観_H）
  - 旧版: ToG〜KaG, Tobi〜Kami, Flag1, Flag2, TB1〜E4
- シート: 全ヵ所検査/H494検査 の (1,13) 公差, (1,15) 納入+材質, (6,29) "ﾌﾘｰｿﾘ", (5,13)/(5,15) 幅/丈歪公差。
- 文字列の戻り値:
  - DbReturnItem（",フリー反り,ラテラルボー,直角度"）
  - MaterialsUpperPart（保護材名/"一致なし"/""）
  - CheckEquipmentDb（"可"/作業可否/"…範囲外"/""）
  - GaichuPackingNo（包装仕様書NO/""）
  - IsGaugeRequired（Bool）, BhSearch（Long）, ShiftTime, StraightEdge
- ログ: Access.log（時間チェック・PLEASE WAIT）、Measured.log。

---

## 10. エラー処理方式

- 担当モジュールの主要なロジック（Q / QDB / QP の大半）は **On Error を一切使っていない**。On Error があるのは QP の EnsureLoaded / EnsureLoadedGauge だけ（QP L2404, L2535）。
- 異常時の方式:
  1. `MsgBox …, vbCritical` を出して `Exit Function/Sub`（戻り値は Empty）。例: QP L92, L162, L174, L450。
  2. `IsOver12Hours = False` のときだけ MsgBox を出す（作業開始から1時間以内だけ警告し、以後は黙って Exit）。QDB L103, L148, L525, L769, L1510, L1554。
  3. `If IsEmpty(...) Then Exit`。GetFieldsArr/GetRecordsArr の空は「ファイル無し/テーブル空/該当なし」をまとめて表す。
  4. 想定外のデータ（列が無く index 0、空の配列、数値でない値を Long/Single に代入、Null 由来の "" を Single に、ArrQuality が Empty のまま UBound など）では**実行時エラーになり、VBA のデバッグダイアログで停止する**。
  5. ファイルが無いと adoRecordset の `End`（全モジュール変数・フォームのリセット）またはエラー（§0-3）。
  6. 外注条件・ゲージ判定は失敗すると MsgBox を出して m_LoadFailed = True。**セッション中は ReloadGaichu/ReloadGauge を手動で実行するまで再試行しない**。ゲージは False（定盤）として動く。
- 補助: ReadCheck / IsBookOpen / IsBookOpened は On Error Resume Next。IsEmptyArray / Get_ArrayDim はエラー発生を判定に使う。WriteLog は On Error Resume Next（ログ失敗は無視）。
- Python 側の方針（提案）: 「Empty を返して黙って続ける」と「MsgBox を出して中断」の違いは業務フローに影響するので、**結果型（OK / 該当なし / エラー＋メッセージ）として明示的に再現**する。実行時エラーになる経路は、VBA ではその時点で処理が止まるので、同等性の定義（エラーとして中断する）を決める必要がある（§14）。

---

## 11. 状態保持・長時間起動・日付変更・マスタ更新後の古い値

| 状態 | 場所 | リスク |
|---|---|---|
| `GradeArray`（Public） | QDB L3 | UF1/UFMitsubishi から開いたときだけ更新される。手動検索（FQ CommandButton1）では戻り値を代入しないので、**TabAb_Change は古いロットの配列を表示する**。 |
| `m_PackQuality` / `m_CacheOrdNo` | QP L5-23 | 梱包保護材マスタのキャッシュ。**キーがオーダーNo**なので、同じオーダーNo の間は再読込しない。`ClearMaterialsCache` は誰も呼ばない。マスタを更新しても、オーダーが切り替わるまで反映されない。`IsEmpty(配列)` は常に False で、空チェックは効いていない。 |
| `m_Head` / `m_Det` / `m_Loaded` / `m_LoadFailed` | QP L37-42 | 外注条件は**ブックを開いている間ずっと初回の読込のまま**（コメントに「フォームを開き直しても反映」とあるが、モジュール変数はフォームを閉じても残る。推測: Excel を再起動するか ReloadGaichu を実行するまで残る。要確認）。読込に失敗すると再試行しない。 |
| `m_Rule` / `m_GaugeLoaded` / `m_GaugeLoadFailed` | QP L58-61 | 同上（ゲージ判定条件）。 |
| `MMBuf/ColBuf/RowBuf` | H L11-13 | 宣言だけで未使用（推測）。 |
| `Static fso` | WriteLog L22 | 問題なし。 |
| UF品質基準 の表示値 | FQ | モードレスで、閉じるまで値が残る。Initialize がクリアするのは一部だけ（Mimi1, Naka1, Mimi2, Naka2, Toukyu, VcFlag, KousinL*, Noun の色, Sansyou1 はクリアしない）。DBdisplay歪 は Sansyou1 に**追記**する（重複語は除去）。範囲外の厚では**前回の値が残る**（旧版と DB 版の両方）。TabAb / MPH は**ページが毎回追加され増え続ける**（AddPagesArrN / AddPagesForSubFolders。クリア処理なし）。 |
| `IsOver12Hours` | Worker | WorkerTime シートの C1 からの `DateDiff("h")`。これは時の境界を何回またいだかで数える（10:59→11:00 でも 1）。経過時間ではない。 |
| `TimeDesignation` | H | 毎時の時間窓。長時間起動でも状態は持たないが、その時間帯は NextPath（梱包課コピー。コメントでは「1日更新」）を読む。**古いデータで判定する可能性**がある。 |
| 日付変更 | – | このグループには日付切替の処理は無い。ShiftTime は時刻表を返すだけ（使う側は Worker）。 |
| Lot情報履歴 シート | – | 旧版の Q と MaterialsUpperPart / GaichuPackingNo が参照するローカルのスナップショット。SIKALOT 側の変更は、シートが更新されるまで反映されない。 |

---

## 12. Access → SQLite 移行時の注意点（このコードに即して）

1. **Filter の意味をそろえる**: `LIKE '*x*'` は `LIKE '%x%'` になる。ただし SQLite の LIKE は ASCII だけ大小文字を区別しない。半角カナや全角の扱い、ADO Filter の大小文字（要確認）に合わせる必要がある。前後一致の Like は、VBA 側の照合（FindInList / Like 演算子）と DB 側の絞り込みが二重に行われているので、**DB 側は広めに取り、最終判定はアプリ側の VBA 互換関数で行う**のが安全。
2. **Null と '' の区別**: GetString で Null は "" になる。CheckThickness などは "" を「空欄（条件なし）」として扱う。SQLite では NULL と '' を同じ扱いに正規化する。ただし QP L283 の `= ''` が Access で Null 行に一致していなかったなら、それも再現が必要（要確認）。
3. **型**: すべて文字列として扱われていた。数値の文字列表現（例 1.5 / 1.50、指数表記）、日付（更新日をそのまま表示。書式は地域設定依存）、Yes/No（"True"/"False" を CBool）を再現する。SQLite は動的型なので、保存時の型と、取り出して文字列にするときの書式を固定する。
4. **行の順序**: ORDER BY が無く、Access の返却順に依存している:
   - 「最初の一致」: 外観_S、QA_022_01、DIM_122_B/D、LoopArray、SoriArray、MaterialsUpperPart、DBQP の優先行・NewArray の順・UpdateDay、StraightEdge
   - 「最後の一致」: 外観_H の管理番号、DIM_122_A の MG1、DIM_122_C、BhSearch、CheckEquipmentDb
   - 「先頭4行」: ShiftTime
   - SQLite では `ORDER BY` を明示する。Access の実際の順序（主キー順か物理順か）を移行前に実データで確認すること（要確認）。
5. **列の順序と列名**:
   - LoopArray は「列名が数値の列」を**列の並び順で**区間として扱う（並べ替えない）。
   - FindIndices は "H500" / "T2000" 形式の列名から区間を作る。
   - FalseSearch_S は除外後の1列目を「条件」とみなす。
   - 外観_H は条件列以外の**全列を欠陥名**とみなす。
   - SQLite でも列順・列名を厳密に保つか、横持ちから縦持ち（サイズ, 値）に正規化する場合も**区間の境界（上端を含む、範囲外は最大列、0以下は最大列）を同じにする**。
6. **識別子**: テーブル名・列名に日本語、半角カナ（ｵｰﾀﾞｰ板厚）、記号（`中歪・耳歪`、`引張強さ(Mpa)`、`1m当たり平坦度上限`）、数字だけの列名がある。SQLite では二重引用符で囲む。
7. **パスワード** "nlm" による保護は SQLite には無い。共有フォルダの権限で代替する。
8. **同時アクセス**: 現状は読み取り専用で、1回ごとに開閉している。品証が Access で編集中に読む構成。SQLite をネットワーク共有に置くのはロックの問題があるので非推奨。品証側の編集手段（現在は Access のテーブルを直接編集。QP L2197-2239 の説明コメント）も代わりを用意する必要がある。
9. **ファイルの切替**: MainPath が更新中のとき NextPath に切り替える仕組み（TimeDesignation / PLEASE WAIT）。SQLite に移行するなら、生成側（SIKALOTNOW を作る処理）との整合が必要。
10. **値の埋め込み**: Filter に値を連結していて、`'` を含む値で壊れる。SQLite ではパラメータバインドにするが、**壊れていた挙動との差**（VBA では該当なし・エラー）を許容するか決める。

---

## 13. 未使用コード候補（根拠付き。断定しない）

| 対象 | 根拠 |
|---|---|
| Q: 品質グレード, 平坦度, 表面品質基準, SelectAdd | 呼び出し箇所は全部コメント（UF1 L3072-3078、FQ L15 等）。ボタンの OnAction から呼ばれる可能性は frx/シートを見ていないので要確認 |
| Q: 中歪, 耳歪, PatternJudge, GradeJudge | 呼び出し元が上の未使用関数だけ |
| QP: ClearMaterialsCache | 呼び出しが見つからない |
| QP: UpdateQpArryD | 呼び出し側（UF1 L3104, UFMitsubishi L258 など）はコメント |
| QP: ExtractNumber, IsHalfWidth | 呼び出しなし |
| QP: ReloadGaichu, ReloadGauge | コードからの呼び出しなし（MsgBox で「実行してください」と案内する手動マクロ） |
| QP: HikiArrCreate の引数 n、FormReflectDist の引数 Execution | 関数内で使っていない |
| QP L1499-1612 旧 MaterialsUpperPart | 全行コメント |
| QDB L345-387, L1677-1693 | コメント |
| H: RightClick, UF1_Click, AppActiveWindow, CheckBookOpened, ListBoxDBArr, setListBoxArr, GetListBoxDataStr, Get_ArrayDim（setListBoxArr からだけ）, CreateHighQualityReport, CreateSection | 補助の外から参照されていない（CreateHighQualityReport は 2026.5.15 の報告書を作るだけの一時的なコード） |
| H: Pos_Check, IMEOFF, Numlock_onoff, FileExistence | 参照は全部コメント（UF1 L4415, L4226, L4116、SaveLoad L62, L581） |
| H: 半角へ変換, ReadFile | Test.frm（テスト用フォーム）からだけ |
| H: MMBuf, ColBuf, RowBuf | 参照が見つからない |
| H: EmptyRangeJudge | 中身が全部コメントで、何もしない |
| FQ: CommandButton3/4, Fm2 など | 実際に使われているかは frx を見ないと分からない |
| Test.frm 経由の参照（FindInList, FindIndices, SoriArray, ShiftTime など） | テスト用ハーネス。本番の呼び出しではない |

---

## 14. 不明点・要確認事項

1. **ADO Filter の仕様**: 大小文字・全角半角の扱い、`[検査番号の頭] = ''` が Null の行に一致するか（QP L283）。実 DB で、検査番号の頭が空欄の行が DBQP の候補に入るかを確認する。
2. **Access の返却順**（ORDER BY なし）。最初/最後の一致に依存する全箇所（§12-4）。実データで順序を記録して、同じ順を再現する。
3. **VBA の Variant 比較**: `UF1.MG = 0` のようにテキストボックスの値 "" と数値を比べたとき（QP L1994 ほか）、`Kousa(1,4) = 0` で Kousa が数値でない文字列のとき（QP L462）、`ArrQuality(i,オーダー板幅) >= オーダー板幅` のように文字列と Single を比べたとき（QDB L1251）。エラーか False か、数値比較になるか。Python の比較ヘルパは実機テストに合わせて作る。
4. **ファイル無しの挙動**: adoConnection が Nothing のときの `.ActiveConnection = Nothing` → `End`。実行時エラーになるか End になるか。品質基準.accdb や伝送用ファイル.accdb の共有先が落ちたときの業務上の期待動作。
5. **バグ挙動を再現するか、修正するか**（業務部門の判断が必要）:
   - a. DBQP の最終判定に Conditions(14) 送り先コード・(15) 巾取数が入っていない（QP L328-330）
   - b. 外観_H で調質を判定せず、材質を2回判定している（QDB L299）
   - c. DowFlag の誤記（QP L483）
   - d. 中耳列が「3級」「4級」だけのときに等級表を引かない（QP L505-530）
   - e. Q 平坦度の "3.." と、TaHiz1 への誤代入（Q L1141, L1293-1297）
   - f. DIM_122_D で ArrQuality をループ中に差し替え、ソリ表に LoopArray を使っている（QDB L1338-1354）
   - g. DIM_122_C で一致がなくても "不問" と参照元を付ける（QDB L1276-1282）
   - h. DbReturnItem の該当なしメッセージの条件対応のずれ（QP L1279-1291）
   - i. CheckEquipmentDb の "C_オーダー幅 範囲外" の文言と、材質不一致でも "オーダー厚 範囲外" になること
   - j. UF管理 FGH の列誤り（FK L232）
   - k. FindIndices の Long 切捨てと列名の一致方法（QP L1003, L1042）
   - l. DistArray で SecondSmallest が CVErr を返す経路
   - m. DBdisplayラテラル が 0 件のときの UBound エラー（QDB L1625）
6. **厚のしきい値の違い**（4 / 6 / 8 / >6 / ≥8）は意図したものか（§8-1）。
7. **フラットネスが 3/4 以外のとき**、SkipQP 経路で TableName が "平坦度一覧QP追加" のまま処理が進む（QP L671-689）。実データでそういうロットがあり得るか。
8. **ソリ表・中耳表・QA_022_01・DIM_179 の実際の列名**（数値名か、"H…"/"T…" 形式か）と、その列順が昇順かどうか。LoopArray / FindIndices の前提を満たしているか。
9. `GetString` が返す数値・日付・Yes/No の文字列書式（PC の地域設定に依存）。UpdateDay / KousinL の表示形式。
10. B1〜B11 の Caption（異常名）の実値。`基準反映` の一致キーになるが、frx はこのグループの資料に含まれていない。同様に `UF品質基準` / `UF管理` のコントロール型（TextBox か Label か）。
11. Q（旧版）を Python に移植する必要があるか。コード上は到達できないが、運用で手動実行されている可能性。
12. 外注条件・ゲージ判定の「フォームを開き直しても反映」というコメント（QP L2239）と実際の挙動（モジュール変数が残る）の食い違い。
13. MaterialsUpperPart が Lot情報履歴の**見出し行を含めて**オーダーNo を探していること（QP L1682。実害は無いと推測）。キャッシュキーがオーダーNo である業務上の理由。
14. レジストリの「モード」がオンラインでもオフラインでもないとき、パスが空になる挙動（初期設定 L18-35）。
15. `IsOver12Hours` の名前と実際の判定（1時の境界をまたいだら True）の違い。メッセージを抑止する仕様として正しいか。
16. 品証検証 / Test シートはこのグループのロジックでは使われていない。他グループで使われているか確認する。

---

## 要約

- Quality は品質規格をハードコードした旧版で、現在は `BoColor` 以外に呼び出し元が無い。本番の基準は全部 品質基準.accdb（QualityPath）から読む。QDB は UF品質基準 の表示（外観・歪・バーコール・ラテラル）、QP は本番の歪公差エンジン（DBQP / DistArray / FormReflectDist）と各マスタ判定を担う。2つは照合関数を互いに使っている。
- Access 読込は毎回 `SELECT * FROM 表` を発行し、ADO Filter で絞る。値はすべて文字列（Null→""）。1回ごとに接続を開閉し、読み取り専用。トランザクションと書き込みは無い。DB パスワードは "nlm"。
- 照合の要は FindInList（カンマ区切り、`*` と `≠`、後勝ちの OR）、CheckThickness/Width（3列とも空なら真）、LoopArray/FindIndices（上端を含む区間、範囲外と0以下は最大列）。行順・列順に依存しているので、SQLite 移行時は ORDER BY と列構成を固定する必要がある。
- バグ的な挙動が多い。代表例: DBQP で送り先コード・巾取数が判定外、外観_H で調質を判定していない、等級だけの中耳表を引かない、"3.." の誤記。再現するか修正するかは業務の判断が必要。
- 状態: GradeArray が古いまま残る、梱包保護材・外注条件・ゲージ判定のキャッシュがマスタ更新を反映しない、フォームの値やタブが残り続ける。TimeDesignation により毎時 :59:46〜:00:45 は NextPath（古い可能性のあるコピー）を読む。
- UF管理 はローカルシート（保存用/耳用）の履歴を閲覧・復元するだけで、パスワードもマスタ編集も無い。
