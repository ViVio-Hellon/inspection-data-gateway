# G3 構造解析レポート：データ入出力・作業者（SaveLoad / AccHIki / SQL_追加 / MachineTransfer / WriteLog / UFAinput / Worker / UFNameList / UFWorkerName）

- 対象: `UF検査判定.xlsm` の VBA（UTF-8 化済みエクスポート）
- 行番号は各 `.bas` ファイルの行番号（1 始まり）。「L123」は当該モジュールの 123 行目を表す。他モジュールは `モジュール名:L行` で表記する。
- **推測**には「推測」と書いた。断定できない事項は §14 に「要確認」としてまとめた。
- 実装はしていない。解析だけを行った。

---

## 0. 前提：パス変数（`初期設定.bas`）

`Setting_Path`（初期設定:L18-48）がレジストリ `GetSetting("UF検査判定システム","起動設定","モード")` を読み、以下のグローバル変数を決める。

| 変数 | オンライン | オフライン |
|---|---|---|
| MainPath | `\\nlmsrvngy03\Read\【New】仕掛\台帳\` | `ThisWorkbook.path & "\"` |
| NextPath | `\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\仕掛\` | 同上 |
| QualityPath | `\\nlmsrvngy03\工場内共有\検査データ\品証-梱包やり取りフォルダ\Accdb\` | 同上 |
| PackingPath | `\\nlmfangyshrd\各課共有\0130_日軽稲沢\梱包課\AIM\【■】_参照用ファイル\`（サーバー名が **nlmfangyshrd** で他と違う） | 同上 |
| HistoryPath | `\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\検査履歴\` | 同上 |
| FactoryyPath | `\\nlmsrvngy03\工場内共有\検査データ\` | 同上 |
| NisPath | `C:\Tools\測定データ電子化`（末尾に `\` なし） | **未設定（空文字）** |

- モードが「オンライン」「オフライン」のどちらでもない場合（例: 未設定で「存在なし」）、全変数は空文字のままになる。
- 末尾の `\` の有無が変数ごとに違う。NisPath だけ末尾 `\` がない。

---

## 1. 各モジュール概要

| モジュール | 行数 | 概要 |
|---|---|---|
| 【UFKen_03】SaveLoad | 4570 | ロット引当情報のフォーム反映（`AccReflec`）、Lot情報履歴シートへの書込（`HistorySave`）、作業中データのシート保存・復元（`Posting`/`Load_Text`）、印刷時の確定データ保存（`MainDataSave`→シート `保存用`/`耳用` と Access `梱包検査ﾃﾞｰﾀ.accdb`）、印刷済み履歴の呼出し（`HistoryLoad`）、ログファイルからの復元（`AccesslogLoad` ほか）、VC 画像表示、作業者・時刻記録（`Time_Posting`）。L574-1256 は旧 `AccReflec` と旧「公差表示」で、**全行コメントアウト**されている（L1760-1853 も同様）。 |
| 【UFKen_01】AccHIki | 705 | 仕掛 DB（`SIKA*NOW.accdb`）の読込ラッパー（`Arrbox`）、旧方式の一覧シート更新（`AccUpdate` 系。呼出し元なし）、固定長テキスト `sikaodr*.txt` の展開（呼出し元なし）、Lot情報履歴からのロット配列抽出（`GetLotInfo`）。`Option Explicit` はコメントアウトされている（L2）。 |
| 【UFKen_15】SQL_追加 | 808 | ADO/ADOX の汎用ラッパー。**システム全体の Access アクセスはほぼすべてここの `GetRecordsArr`/`GetFieldsArr` を経由する**（呼出し箇所: QP_DB 20/22、Quality_DB 20/16 など）。書込みは `adoSQL`（ADDNEW/UPDATE/DELETE）と `AddAccdb`。`Option Private Module`（L3）。 |
| 【UFKen_06】MachineTransfer | 546 | 測定器（マイクロメータ・ノギス・デプス）の選定チェック（`Micro_Juge`）、印刷シートへの使用機器名の転記（`機器転送`）、`伝送用ファイル.accdb` の `測定機器ID` テーブルから機器を引く `DbReturnMicro`。「設備への転送」ではなく「**測定機器情報の転記**」のモジュールである。 |
| 【UFKen_16】WriteLog | 40 | CSV 形式の追記ログを書く `WriteLog`。全モジュールから合計約 350 箇所で呼ばれる。 |
| UFAinput（フォーム） | 1497 | 梱包実績（A インプット）の入力画面。高さ測定、ホスト（ｱｲｶｽ）送信用 `hostsend.txt` の出力、送信アプリ `SokuteiDataTrans.exe` の起動、`result.txt` の結果読込、日報への記録。主な処理本体は【UFKen_09】Input（担当外）にある。 |
| 【UFKen_07】Worker | 1021 | 直（シフト）の時間定義、日付・直の判定（夜勤の日付またぎ）、検査員カード番号から氏名への変換（`NameConv`）、検査員シートの保存と読込、日報シート（`Daily_*`）の作成と記入、梱包枚数と検入枚数の照合、検査員の不一致チェック（`Name_Match`）。 |
| UFNameList（フォーム） | 384 | 日報閲覧・編集、検査員不一致リスト表示、VBA 停止（パスワード）を兼ねるリスト画面。 |
| UFWorkerName（フォーム） | 12 | クリックすると `UF1.Show vbModeless` を呼ぶだけ（L10-12）。 |

---

## 2. Sub / Function 一覧

### 2.1 【UFKen_03】SaveLoad

| 名前 | 公開 | 引数 | 役割 | 呼出し元 → 呼出し先（主なもの） |
|---|---|---|---|---|
| SikaAtuCreate (L16) | Public Function | LotNo | SIKALOTNOW `仕掛` からロットの「ｵｰﾀﾞｰ板厚」を 1 次元配列で返す | UF1:L250, FormRelat:L2589 → GetFieldsArr/GetRecordsArr |
| AccReflec (L56) | Public Sub | WorkTime="S" | **ロット引当の中核処理**。SIKALOTNOW を SQL で直接読み、フォームへ反映し、設備仕様をチェックし、HistorySave を呼び、VC と注意喚起を出し、ゲージ判定を行う | UF1:L4511（ロット選択時） → CheckEquipmentDb, MsgError, HistorySave, LSelectLateral, VC_Clear/VC_Picture, Atu_OPselect, Conductiv_*, Arrbox, MaterialsMsgBox, IsGaugeRequired, CBAdd, GetLotInfo, JudgeLotAllocations |
| HistorySave (L1258) | Public Sub | PastHiki(0..18) | `Lot情報履歴` シートで同じロットの行を削除してから追記する（SIKAHIKINOW × SIKAODRNOW の受注番号で結合）。LSelect 用のキャッシュも設定する | AccReflec:L273, UFCask:L535, UFMitsubishi:L804 → GetRecordsArr/GetFieldsArr, IsOver12Hours |
| VC_Picture (L1554) | Public Sub | OrderNo="" | VC（化粧紙）画像を表示する。`伝送用ファイル.accdb` の `VC厚` から系統・備考を取る | AccReflec:L381, Load_Text:L2958, UF1:L1788, Judg:L743 → MaterialsUpperPart, GetRecordsArr |
| ReflectStrain (L1855) | Public Sub | n, DefaultB | 歪（定盤・ゲージ）の A 面/B 面を `作業保存_<ライン>` の O:P 列に退避・入替する | UF1（多数）, Posting:L2049, HistoryLoad:L3775 → UF_Array, UFtag_Array, LosdStrain |
| LosdStrain (L1903) | Public Sub | Temp("A"/"B"), Aran, Bran | シートの歪文字列 `CB1_CB2_Z` を分解してフォームと `.Tag` に設定する | ReflectStrain, Load_Text, FormReflectionEx, UF1:L3501 → Split_Z/G(_0) |
| Posting (L1981) | Public Sub | なし | **作業中保存**。UF1 の内容を `作業保存_<ライン>` シートへ書く | UF1（16 箇所。入力のたび）, UFConfi, UFdiagonal, FormRelat, Judg, OutTolerance → UF_Array, ReflectStrain, GetLotInfo, JudgeLotAllocations |
| UF_Array (L2135) | Public Function | n,k(列範囲),x,y(行範囲) | UF1 の測定値（S/U/C/D/E/F × 1-20 項目）を 2 次元配列にする | FormRelat（27）, Print（3）, 本モジュール |
| UF_SaveArray (L2236) | Public Function | 同上 | 耳付きを除く保存用（1-10 測定、11 定盤歪、12 ゲージ歪） | MainDataSave |
| Time_Posting (L2288) | Public Sub | Position, EachT, Temp="1" | 測定位置ごとの作業者・測定時刻・PakeD を `WorkerTime_<ライン>` に記録する | LabelCls（6）, FormRelat:L333-342 → WorkeDayMax |
| ApplyWorkerToTimeSheet (L2362) | Public Sub | Position, WorkerName, timeStamp | Time_Posting の作業者書込部分だけを再現する（ログ復元用） | ApplyWorkerFromLog |
| Split_Z_0 / Split_G_0 / Split_Z / Split_G / Split_Z_Tag / Split_G_Tag / Split_Z_Tag_0 / Split_G_Tag_0 (L2394-2450) | Public Sub | n, a, b | 歪コントロール（CB{n}1..4, Z{n}, G{n}）とその Tag へ値を配る | LosdStrain, HistoryLoad, UF1, OutTolerance（Split_Z_Tag_0）。Split_G_Tag_0 は呼出し元なし |
| OrdNoOpen (L2452) | Public Sub | LotNo | Lot情報履歴からオーダーNo を集め、引当No 順に並べ替えて UF1.OrdNoCB に入れる | UF1:L1718, L4518 → merge_sort2 |
| Load_Text (L2501) | Public Sub | n=""/"表示"/"ReLoad" | **作業再開**。`作業保存_<ライン>` からフォームを復元する | UF1:L3503, FormRelat:L6527, OutTolerance:L1043 → LosdStrain, Atu_OPselect, LSelect, LSelectLateral, Six_Display, AtCountSelect, InspectionSelect, SlabSet, CalculateSheets, SmallNumberCount, DrawListBoxContent, VC_Picture, InitialSetting, Conductiv_* |
| AccesslogLoad (L3005) | Public Sub | Lot, Worke, logName, Reas | ローカルのログファイル（Shift_JIS）からロットの測定値を抽出して `Accesslog` シートへ書き、フォームへ反映する | UF1:L583（CommandButton15） → ApplyVisualStrainFromLog, FormReflection, FormReflectionEx, ApplyWorkerFromLog |
| GetPosRow (L3172) | Public Function | posChar | S..F を 1..6 に変換する | 本モジュール |
| ApplyVisualStrainFromLog (L3185) | Public Sub | LineBuff(), SameCount, EndCount | ログから外観（→Accesslog U:V 列）と歪（→作業保存 O:P 列）を復元する | AccesslogLoad |
| ApplyWorkerFromLog (L3361) | Public Sub | 同上 | ログ末尾の検査員名を WorkerTime シートへ復元する | AccesslogLoad |
| HistoryLoad (L3409) | Public Sub | HistLot, TempDate="" | 印刷済みデータ（シート `保存用` または `耳用`）を UF1 に呼び出す | UF管理:L628/631（Print:L591 の一致はコメント内の文字列で、呼出しではない） → Split_*, ReflectStrain, LSelect, LSelectLateral, RightAngleDisplay, Six_Display, InitialSetting, ReDiag |
| FormReflection (L4052) | Public Sub | Temp="" | `Accesslog` シートの値をフォームへ反映する | AccesslogLoad |
| FormReflectionEx (L4192) | Public Sub | なし | 外観と歪の反映 | AccesslogLoad |
| MainDataSave (L4216) | Public Sub | TemWs="保存用", Temp="耳用" | **検査結果の確定保存**。`保存用`（通常）または `耳用`（極厚＝耳付き）シートへ 1 行追記し、その行を `梱包検査ﾃﾞｰﾀ.accdb` へ ADDNEW する | UF1:L3776（印刷 Yes 後） → UF_Array, UF_SaveArray, UFtag_Array, WorkeDayMax, adoSQL |

### 2.2 【UFKen_01】AccHIki

| 名前 | 公開 | 引数 | 役割 | 呼出し元 → 呼出し先 |
|---|---|---|---|---|
| Arrbox (L14) | Public Function | target, SQLfilter, Use="Record"/"Fields", path | MainPath（path="" の場合）または NextPath の `target`（SIKA*.accdb）の `仕掛` テーブルを読む。SQLfilter に "板"/"コイル"/ロットNo などのキーワードを渡す | FormRelat（22）, Quality_DB（8）, SaveLoad（4）, Judg（2）, QP_DB（2） → ReadCheck, TimeDesignation, FileDesignation, IsOver12Hours, GetRecordsArr/GetFieldsArr |
| Lotflag (L94) | Public Function | target, shName | ファイル更新日時とシートの記録日時を比較する | **呼出し元なし**（UF1:L3372 はコメント） |
| AccUpdate (L124) | Public Sub | なし | SIKALOTNOW を `LOT一覧` シートへ展開する | **呼出し元なし** |
| AccUpdate_Etc (L183) | Public Sub | target, shName | フィルタなしで指定シートへ展開する | **呼出し元なし** |
| Acc_Etc (L233) | Public Sub | target, shName | コイル用。L264 で未定義変数 `HIK` を参照するため、実行すると実行時エラーになる | **呼出し元なし**（ZZ:L201 はコメント） |
| Order_Update (L277) | Public Sub | n | `\\Nlmfanago02d\仕掛ﾃﾞｰﾀ\sikaodr<yyyymmdd>.txt` を `VC一覧` シートへ展開する（6:35 で日付を切り替え） | **呼出し元なし** |
| オーダー更新 (L335) / オーダー処理 (L346) | Public Sub | fileName, … | 固定長テキストの展開 | Order_Update からのみ |
| Conv (L576) | Private Function | Char, FuncNo, FormatNo, PosiNo, j, title | 固定長フィールドの変換 | オーダー処理からのみ。`i` が暗黙の Variant（Option Explicit なし）のため L606 の `If i = 1` は常に偽になる（推測） |
| GetLotInfo (L644) | Public Function | SheetName, lottemp | Lot情報履歴から該当ロットの [納入先, 得意先, 用途ｺｰﾄﾞ, 材質, 調質, 厚, 巾, 丈, ﾛｯﾄNo, 検査No] を (n×10) 配列で返す | UF1:L673, UFMitsubishi:L808, FormRelat:L3844/L6530, SaveLoad:L554/L2117, OutTolerance:L158 |

### 2.3 【UFKen_15】SQL_追加

| 名前 | 公開 | 引数 | 役割 | 呼出し元 |
|---|---|---|---|---|
| fldAccdb (L64) | Public Sub | tgtDb, TableName, fileName, FilePath, newflag | ADOX で DB とテーブルを作成する（全列 `adVarWChar`） | ExpAccdb |
| AddAccdb (L126) | Public Sub | tgtDb, TableName, fileName, FilePath | 2 次元配列の 2 行目以降を AddNew で一括登録する（トランザクションあり） | ExpAccdb |
| AdoOpenCode (L194) | Private Function | — | `"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="` を返す | fldAccdb, AddAccdb |
| adoConnection (L199) | Public Function | FilePath | 拡張子から Provider を選び、パスワード `nlm` を付けて接続し、`CursorLocation=adUseClient` にする | 本モジュール全体 |
| adoTableDelete (L272) | Public Sub | FilePath, TableName | ADOX でテーブルを削除する | **呼出し元なし** |
| FileExtensionType (L278) | Private Function | tgtPath | 拡張子（小文字） | adoConnection |
| adoRecordset (L284) | Public Function | TableName, adoCn, SQLfilter, CursorType, LockType | `SELECT * FROM <table>` に **ADO の Filter プロパティ**でフィルタを掛けて開く。adoCn が Nothing なら **`End`（VBA 全停止）** | 本モジュール |
| adoCommand (L316) / adoDic (L322) | Private Function | — | Command / Dictionary を生成する | 本モジュール |
| adoFieldkeyNo (L326) | Public Function | adoCn, adoRs, keyFieldName, jcnt | フィールド名から 0 始まりのインデックスを返す | adoSQL |
| GetFieldDic (L339) | Public Function | FilePath, TableName, ColLim | フィールド名 → 列番号の Dictionary | **呼出し元なし** |
| adoFieldDic (L361) | Public Function | adoRs | 同上（Recordset 版） | **呼出し元なし** |
| AdoArr (L375) | Public Function | FilePath, TableName, SQLfilter, ColLim | {"fields","record"} の Dictionary | adoRead のみ |
| adoFields (L407) | Public Function | adoRs, jcnt | フィールド名配列 (1×n) | 本モジュール |
| adoRecords (L419) | Public Function | adoRs, jcnt, PageCnt, pageNo | `GetString` で全行を **文字列化**して (行×列) 配列にする | 本モジュール |
| GetFieldsArr (L444) | Public Function | FilePath, TableName, SQLfilter, ColLim | フィールド名配列 | 全モジュール |
| GetRecordsArr (L448) | Public Function | 同上 | レコード配列（全値が文字列） | 全モジュール |
| GetRecordsAdo (L451) | Public Function | adoCn, TableName, SQLfilter | 既存接続でレコード配列を返す | **呼出し元なし** |
| GetAdoArr (L458) | Private Function | …, GetType | 接続を開き、読み、閉じる共通処理 | GetFieldsArr/GetRecordsArr |
| adoRead (L473) | Public Function | TFctrl, … | ListView に表示する | **呼出し元なし** |
| adoReadPageCtrl (L478) | Public Function | TF, …, PageCnt, pageNo | ページング表示 | **呼出し元なし** |
| adoSQL (L525) | Public Sub | FilePath, adoTABLENAME, MODE, tgtInput(0..), keyFieldName, skipmsgflag, keySQLfilter | ADDNEW/UPDATE/DELETE（トランザクションあり） | MainDataSave:L4567/4568, Test:L561 |
| TxtToAccdb (L707) | Public Sub | FilePath, shName, TableName | 未完成（`Temp()` が空のまま ExpAccdb を呼ぶ） | **呼出し元なし** |
| XlsToAccdb (L722) | Public Sub | ArrSheetName, ArrTableName, targetFilename, accdbFilename | Excel から accdb を作り直す | **呼出し元なし** |
| getAddress (L759) | Private Function | — | パスを分解する | XlsToAccdb |
| ExpAccdb (L771) | Public Sub | Temp, TableName, fileName, FilePath, newflag | fldAccdb と AddAccdb を続けて呼ぶ | Test:L868（デバッグ用）, XlsToAccdb, TxtToAccdb |
| OpenXlsWs (L776) | Private Function | — | ブックを開いてシートを配列化する | **呼出し元なし** |
| tgtFileDelete (L802) | Public Sub | fileName, FilePath | ファイルを削除する | XlsToAccdb |

### 2.4 【UFKen_06】MachineTransfer

| 名前 | 公開 | 引数 | 役割 | 呼出し元 |
|---|---|---|---|---|
| Micro_Juge_旧 (L4) | Public Sub | st | 旧方式（機器 ID をハードコード） | **呼出し元なし** |
| Micro_Juge (L165) | Public Sub | st（座標名 S1 等） | 選択中の機器（UF1.Machin）が DB 上の正しい機器でなければ、測定値をクリアしてエラーを出す | FormRelat:L284 |
| Machin_Text (L216) | Public Function | Trm | 測定器の伝送文字列 `DT10002+0000021.712M` から 5 桁の機器番号 `Mid(Trm,3,5)` を取り出す | FormRelat:L283 |
| 機器転送 (L276) | Public Sub | ws, x, y, ATU, HABA, Yout | 印刷シートの (x,y) に「ライン 機器名」、(x,y+2) に「ライン コンベックス番号」を書く | Print:L544, UFMitsubishi:L156, Mitsubishi（5 箇所） |
| DbReturnMicro (L406) | Public Function | LineName, ATU, Yout, n="", Dep="" | `測定機器ID` テーブルから、厚さと用途に合う機器の出力テキスト・機器名・ID を返す | Micro_Juge, 機器転送, UFAinput.Height_Machin, Test:L174 |
| CheckThickAndCode (L503) | Public Function（修飾子なし） | row, Col1, Col2, Col3, OrderT, Yout, ArrQuality | 用途条件（`≠X`, `*X`, `X*`, 完全一致、カンマ区切り）と厚さ範囲を判定する | DbReturnMicro |

### 2.5 【UFKen_16】WriteLog

| 名前 | 公開 | 引数 | 役割 |
|---|---|---|---|
| WriteLog (L6) | Public Sub | msg1, Sepl=",", TargetPath="", TargetFile="" | `<TargetPath>\<TargetFile or Access.log>` に 1 行追記する。TargetPath="" の場合は `ThisWorkbook.path\Access.log` |

### 2.6 【UFKen_07】Worker

| 名前 | 公開 | 引数 | 役割 | 呼出し元 |
|---|---|---|---|---|
| Set_Shift (L12) | Public Sub | — | `ShiftTime()`（QP_DB。`伝送用ファイル.accdb` の `時間用`）から Start1..3/昼 と End1..3/昼 を設定する | UF1:L225（起動時）, Test |
| NameConv (L37) | Public Function | no（13 桁のカード番号） | `伝送用ファイル.accdb` の `名簿` テーブル [No] → [名前] | UF1:L1624/1629/4797 |
| YeaDayrCheck (L143) | Public Function | DTime | 直を考慮した日付 `yyyy.mm.dd`（夜勤の 0 時以降は前日） | UF1, UFCask, UF校正, Proofreading, Register_Check, Daily_Create ほか |
| DayCheck (L204) | Public Function | DTime | 日付と時刻 `yyyy.mm.dd  hh:nn`（夜勤の 0 時以降は前日の「24+h 時」表記） | UF1:L4808-4829 |
| TimeCheck (L251) | Public Function | DTime | "日勤"/"1直"/"2直"/"3直" | UF1, UFCask, UF校正, Proofreading, 本モジュール |
| EndTime_Register (L312) | Public Sub | — | UF1.EndTime と `WorkerTime_<ライン>`!C1 を設定する | UF1:L3641（印刷）, Print:L190 |
| IsOver12Hours (L328) | Public Function | — | `WorkerTime_<ライン>`!C1（終了時刻）から **DateDiff("h") ≥ 1** なら True。名前は「12 時間」だが実際の閾値は 1 時間で、しかも経過時間ではなく「時」の境界をまたいだ数で判定する | Arrbox, HistorySave, FormRelat, Quality_DB, Test |
| Name_Match (L352) | Public Sub | — | WorkerTime 上の測定者と登録検査員（T1/T2/T3）を照合する。不一致があれば UFNameList を開く。**ThisWorkbook.Save を実行する** | UF1:L3718（印刷前） |
| Register_Check (L456) | Public Function | STtime, n="" | 検査員登録が当日・当直で行われているか（True なら要登録） | UF1（多数）, UFCask, UFMitsubishi |
| Epual (L498) | Public Function | Wor | 配列の全要素が同じなら True | WorkerSave |
| ReWorker (L509) | Public Function | Wor, n="_" | Split のラッパー | WorkerSave |
| WorkerSave (L513) | Public Sub | — | UF1 の Day1..3 と T1..3 を `検査員_<ライン>` の A1:B3 に保存する | UF1:L4105/4127/4147（T1-3_Change） |
| WorkerLoad (L566) | Public Sub | — | `検査員_<ライン>` からフォームへ戻す | UF1:L3504, FormRelat:L6533/6563 |
| WorkeDayMax (L585) | Public Function | — | 検査員シート上で最も遅い登録の直番号 "1"/"2"/"3" | 多数 |
| GetWorkerName (L707) | Public Function | — | 最新の直の登録者（`_` 区切りの末尾） | Judg, OutTolerance, 補助（LogOut 等） |
| Daily_Create (L719) | Public Sub | — | `Daily_<ライン>` シートを（直が変わったら作り直して）準備する | UF1:L4765, UFCask, UFMitsubishi |
| Daily_Report (L745) | Public Sub | n(1-4) | 日報への記入（1: ヘッダ、2: ロット情報、3: 開始時刻、4: 終了時刻と梱包） | UFAinput:L125/L1082, UF1:L4861, FormRelat:L313, Cask, シートモジュール（ZenSh/KySh/H494_SH 等） |
| Daily_Report4_Execution (L919) | Public Sub | — | 梱包No の連番チェックをしてから Daily_Report(4) を呼ぶ | **呼出し元なし**（UFAinput:L117/L1055 はコメント） |
| Number_Count (L948) | Public Sub | Lotname | 日報の梱包枚数（枚数×包数）と検入枚数を照合する。超過していれば印刷を止める | UF1:L3719（印刷前） |

### 2.7 UFAinput / UFNameList / UFWorkerName

| 名前 | 公開 | 役割 |
|---|---|---|
| UFAinput.AinButt_Click (L932) | Private | **A インプット送信**の本体（§6.5） |
| UFAinput.CommandButton2_Click (L127) | Private | 入力をクリアし、`hostsend.txt` を再出力してから送信アプリを起動する（再送信） |
| UFAinput.CommandButton3_Click (L185) | Private | `A_Output` だけを実行する |
| UFAinput.CommandButton4_Click (L190) | Private | `result.txt` の読込（InterRuptText） |
| UFAinput.CommandButton5_Click (L194) | Private | 送信アプリの起動 |
| UFAinput.CommandButton7_Click (L202) | Private | `Nitiina_<ライン>` を作り、**ハードコードされた Lot "A111119" と検番 "B22229"** で A_CsvSave を呼ぶ（テスト用と推測） |
| UFAinput.CommandButton1_Click (L116) | Private | `Daily_Test` シートを表示して Daily_Report(4) を呼ぶ（テスト用と推測） |
| UFAinput.CommandButton8_Click (L220) | Private | `ThisWorkbook.path\日報＆資材計算.xlsm` を開いてマクロ `GW_Open` を実行する |
| UFAinput.CommandButton9_Click (L250) | Private | 外注時の高さ再チェック（Height_Check） |
| UFAinput.HikiNo_Change (L317) | Private | `Hiki_Lot` シートから、選択された引当No に対応するオーダー・ロット・検番・調整No・引当数量を表示する。オーダーNo の 3 桁目が 1 なら輸出欄を表示する |
| UFAinput.UserForm_Activate (L798) | Private | `Hiki_Lot` の A 列を読み、数値昇順のバブルソートをして ComboBox に入れる。検査法（KenLaw）が M2/M6 の場合の表示切替 |
| UFAinput.UserForm_Initialize (L880) | Private | `Lot情報履歴` シートがなければ **`End`** |
| UFAinput.Height_Machin (L610) | Public Function | デプスゲージの機器番号を照合する（DbReturnMicro(…,"デプス")） |
| UFAinput.StartTB_KeyDown (L697) | Private | Tab キーで高さ測定値を確定する（CapsLock を強制 OFF） |
| UFAinput.RefreshDynHeightLabels (L1340) / ShowHeightControls (L1473) | Public | 山数（1-15）に応じて動的ラベルを生成する |
| その他 `_Change` 系 | Private | 数値以外を消す入力制限、PileMatch（ストアカード）の頭文字チェックなど |
| UFNameList.UserForm_Activate (L304) | Private | Excel を非表示にして、起動ログを WriteLog する |
| UFNameList.CommandButton2_Click (L261) | Private | ListView の内容を反映（ListPosting）し、`Daily_<ライン>` を再表示する |
| UFNameList.CommandButton3_Click (L278) | Private | ListView を **`Daily_Test` シート（ライン名がハードコード）** に書き戻す |
| UFNameList.CB1_Click (L349) | Private | WorkerTime の不一致リスト（O:Q 列）と全リスト（A:L 列）を切り替える |
| UFNameList.OnOff_Click (L36) | Private | パスワード `weekend99` で VBA 停止（`End`） |
| UFNameList.UserForm_QueryClose (L327) | Private | Content が "検査員不一致" なら UF1 を表示し、"VBA停止" なら End |
| UFWorkerName.UserForm_Click (L10) | Private | UF1 を表示する |

---

## 3. 全データストア一覧

### 3.1 Access DB ファイル

| # | ファイル | 場所（優先 → 代替） | テーブル | R/W | 担当コードでの使用箇所 |
|---|---|---|---|---|---|
| A1 | SIKALOTNOW.accdb | MainPath → NextPath（§3.1.1 の条件で切替） | 仕掛 | **R** | AccReflec（SQL 直書き）, SikaAtuCreate, (AccUpdate) |
| A2 | SIKAHIKINOW.accdb | 同上 | 仕掛 | **R** | HistorySave, AccReflec→Arrbox, (AccUpdate_Etc) |
| A3 | SIKAODRNOW.accdb | HistorySave で決めた FPath → NextPath | 仕掛 | **R**（全件） | HistorySave |
| A4 | 伝送用ファイル.accdb | QualityPath → PackingPath（UF1.Test の場合は PackingPath を優先：NameConv, VC_Picture） | 測定機器ID / 名簿 / VC厚 / 時間用 | **R** | DbReturnMicro, NameConv, VC_Picture, ShiftTime(QP_DB) |
| A5 | 梱包検査ﾃﾞｰﾀ.accdb | `FactoryyPath & "【測定DB】梱包\梱包検査ﾃﾞｰﾀ.accdb"` | 検査ﾃﾞｰﾀ / 耳付検査ﾃﾞｰﾀ | **W**（追記。条件一致時は UPDATE） | MainDataSave → adoSQL |
| A6 | SIKACoil.accdb | ThisWorkbook.path → NextPath | 仕掛 | R | Acc_Etc（未使用・実行不能） |
| A7 | （任意名） | ThisWorkbook.path | 任意 | W（新規作成） | fldAccdb/AddAccdb/ExpAccdb（Test フォームのデバッグ用途のみ） |

※ 担当外モジュールが SQL_追加 経由で使う DB：`品質基準.accdb`、`梱包資材マスタ.accdb`、`SIKACOMP.accdb`（NextPath。Input:L148）、`SIKATWOYEAR.accdb`、`パレット種類_外径条件.accdb`。

#### 3.1.1 SIKA*NOW の参照先切替ロジック（3 系統あり、条件が異なる）

1. **Arrbox / AccUpdate 系**（AccHIki:L19-37）
   - `path=""` なら MainPath、それ以外（= `UF1.DBflag.Text` が空でない）なら NextPath。
   - MainPath の場合、`ReadCheck(MainPath+target)=False` **または** `TimeDesignation()=True`（毎時 59:46-59:59 と 00:00-00:45）**または** `FileDesignation()=False`（MainPath に `PLEASE WAIT.txt` がある）なら NextPath に切り替える。
   - NextPath にもファイルがなければ、`IsOver12Hours=False` のときだけ MsgBox「日稲:<target>がありません_ReadCheck_Arrbox」を出して終了する（戻り値は Empty）。
   - Use="Record"/"Fields" は L49-58 で即 `Exit Function` するため、L60 以降（結果が空のときに NextPath を再参照するフォールバック）は**到達不能**。
2. **AccReflec**（SaveLoad:L101-112）
   - `Dir(MainPath & "PLEASE WAIT.txt")<>""` → NextPath。それ以外で `UF1.DBflag.Text<>""` → NextPath。それ以外は MainPath。
   - **TimeDesignation は見ない**。ファイルがなければ MsgBox「DBファイルが見つかりません」を出して終了する。
3. **HistorySave**（SaveLoad:L1294-1332）
   - PLEASE WAIT.txt → NextPath。「全ヵ所検査」または「キャスク検査」シートが Visible → MainPath。`UF1.DBflag<>""` → NextPath。それ以外は MainPath。
   - さらに L1319 で ReadCheck / TimeDesignation / FileDesignation による切替を行う。
   - 該当ロットのレコードが空なら NextPath で再検索する（L1336-1358）。
4. **SikaAtuCreate**：MainPath でフィールドもレコードも空なら NextPath（ReadCheck や時刻窓は見ない）。

> TimeDesignation（補助:L1027）は `(5945 < 分 & 秒)` という**文字列連結**で比較している。分=59・秒=46-59 なら真になるが、秒 0-9 では "590" 等の 3 桁になり偽になる。L1033 は 分<1 かつ 秒≤45 で真。真の場合は毎回 `WriteLog(分 & ":" & 秒 & "_時間ﾁｪｯｸ", "", NextPath)` で NextPath\Access.log に書く。FileDesignation も PLEASE WAIT.txt がある間は毎回 NextPath\Access.log に書く。

#### 3.1.2 テーブルと列（コード中の参照から抽出。型は不明）

**A1 SIKALOTNOW.accdb / 仕掛**（AccReflec L120-338、SikaAtuCreate L33）
`ｵｰﾀﾞｰ番号, ﾛｯﾄ番号, ｵｰﾀﾞｰ板厚, ｵｰﾀﾞｰ板幅, ｵｰﾀﾞｰ板丈, 製造板厚, 製造板幅, 製造板丈, BOX実績_板厚, BOX実績_板幅, BOX実績_板丈, 検査番号, 製造材質, 製造調質, 用途ｺｰﾄﾞ, 設計_設備ｺｰｽ, 実績_設備ｺｰｽ, 取引先名, 納入先名, 送り先名, BOX実績_枚本数, BOX実績_良品重量, 梱包_合紙, VC, 品質ｸﾞﾚｰﾄﾞ_ﾗﾃﾗﾙﾎﾞｰ, 品質ｸﾞﾚｰﾄﾞ_直角度, 品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ, 製品単重, 板厚公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ, 板厚公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ, 板厚公差_製造_ﾌﾟﾗｽ, 板厚公差_製造_ﾏｲﾅｽ, 板幅公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ, 板幅公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ, 板幅公差_製造_ﾌﾟﾗｽ, 板幅公差_製造_ﾏｲﾅｽ, 板丈公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ, 板丈公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ, 板丈公差_製造_ﾌﾟﾗｽ, 板丈公差_製造_ﾏｲﾅｽ, 指示_直角度, 指示_ﾗﾃﾗﾙﾎﾞｰ`
フィルタ用（AccUpdate 系）: `残仕掛設備ｺｰｽ, BOX設計_設備名, 受注板丈`

**A2 SIKAHIKINOW.accdb / 仕掛**（HistorySave L1369-1381、AccReflec L437）
`受注番号, 引当番号, ﾛｯﾄ番号, 検査番号, 引当数量, 引当調整NO, 引当ﾌｧｲﾙ登録年月日時分秒, 出荷日`（出荷日は列位置を取るだけで未使用）

**A3 SIKAODRNOW.accdb / 仕掛**（HistorySave L1433-1445）
`受注番号, VC_表, VC_裏, 製品単重, 合紙, 取引先名称, 納入先名称, 送り先名称, 包装仕様NO`

**A4 伝送用ファイル.accdb**
- `測定機器ID`（MachineTransfer L422-434）: `ライン, 機器名, ID, 出力テキスト, 厚下, 厚上, 用途, 備`（コード上の列名は「備」。備考の誤記の可能性がある → §14）。値の例（推測）: 機器名 = M1/M2/M3/M4/N5/N6/デプス、出力テキスト = 5 桁の機器番号。
- `名簿`（Worker L57-72）: `No, 名前`
- `VC厚`（SaveLoad L1649-1666）: `品名ｺｰﾄﾞ, 系統, 備考`
- `時間用`（QP_DB:L1786-1823）: `直, 開始, 終了`。**先頭 4 行を固定で読む**（行が 4 行未満なら添字エラー）。

**A5 梱包検査ﾃﾞｰﾀ.accdb**
- `検査ﾃﾞｰﾀ`: 列は**位置で対応づけている**（adoSQL L628-630 `Fields(j-1)=CStr(tgtInput(j-1))`）。Field0 = ライン名（`UF1.LineName.Caption & UF1.LineNo.Caption`）、Field1..Field114 = シート `保存用` の A..DJ 列（§9.6）。キー列名として `[検査開始終了]` を参照している（MainDataSave L4565）。
- `耳付検査ﾃﾞｰﾀ`: Field0 = ライン名、Field1..Field108 = シート `耳用` の A..DD 列。
- テーブル側の列数が配列（0..EndC）より多いと添字エラーになり、ロールバックされる。

#### 3.1.3 SQL 原文（担当モジュール内のすべて）

| 場所 | 種別 | 原文 |
|---|---|---|
| SaveLoad L120 | SELECT（直書き） | `"SELECT * FROM 仕掛 WHERE ﾛｯﾄ番号 = '" & UF1.CB.Text & "' AND ｵｰﾀﾞｰ板丈 <> '0'"` （ｵｰﾀﾞｰ板丈 を**文字列 '0'** と比較している。ロット値は UF1.Lot ではなく **UF1.CB.Text**） |
| SQL_追加 L144 | SELECT | `"SELECT * FROM " & TableName` （AddAccdb） |
| SQL_追加 L293 | SELECT | `"SELECT * FROM " & TableName` （adoRecordset。絞り込みは WHERE ではなく `.filter`） |
| SQL_追加 L542 | SELECT | `"SELECT * FROM " & adoTABLENAME` （adoSQL。`.filter` を併用） |
| SQL_追加 L652-665 | UPDATE（動的） | `MODE & " " & adoTABLENAME & " SET " & "[f1]='v1', [f2]='v2', ... " & "WHERE [key]='val'"` または `"WHERE " & keySQLfilter`。**SET の末尾と WHERE の間に空白がない**（`…'vN'WHERE …`）。全値を `'…'` で囲む（全列テキスト前提）。エスケープ処理はない（値に `'` が含まれると構文エラー）。Field0 は SET に含めない（j=1..jcnt-1）。 |

**ADO Filter 文字列**（SQL ではなく ADO の Recordset.Filter 構文。§12 で詳述）

| 場所 | Filter |
|---|---|
| AccHIki L43 | `[ｵｰﾀﾞｰ板丈]<>0` |
| AccHIki L44 | `[ｵｰﾀﾞｰ板丈]=0  AND  [BOX設計_設備名] = 'KEN' ` |
| AccHIki L45, SaveLoad L39/L1333 | `[ﾛｯﾄ番号]='<LotNo>'` |
| AccHIki L143 | `[ｵｰﾀﾞｰ板丈]<>0 AND [残仕掛設備ｺｰｽ] LIKE '*KEN*' AND  [BOX設計_設備名] = 'KEN' ` |
| AccHIki L246 | `[受注板丈]=0` |
| MachineTransfer L436 | `[ライン]='<LineName>'` |
| MachineTransfer L445 | `[ライン]='<LineName>' AND  [機器名]= '<Dep>'` |
| Worker L57 | `[No]='<no>'` |
| SaveLoad L1649/1650 | `[品名ｺｰﾄﾞ]='<VC_A>'` / `[品名ｺｰﾄﾞ]='<VC_B>'` |
| SaveLoad L4565 | `[検査開始終了] ='<A列の値> '` （**閉じ引用符の前に半角空白**がある → §14） |
| SQL_追加 L546 | `[<keyFieldName>]='<tgtInput(keyNo)>'`（keySQLfilter が空の場合） |

### 3.2 ファイル（Access 以外）

| ファイル | パス | 形式・文字コード | R/W | 書く処理・タイミング |
|---|---|---|---|---|
| Access.log（ローカル） | `ThisWorkbook.path\Access.log` | テキスト。FSO の既定 = ANSI（日本語 Windows では Shift_JIS/cp932 と推測）。CRLF。新規作成時は 1 行目に `管理ﾌｧｲﾙ使用状況` | W（追記）/ R（AccesslogLoad） | WriteLog（TargetPath なし）。LogOut（補助）の測定値ログ、Lot登録ログ（UF1:L4673）など |
| ライン別ログ | `HistoryPath & "<ﾌｫﾙﾀﾞ>\<ﾌｫﾙﾀﾞ>_Log\Access.log"`（ﾌｫﾙﾀﾞ = L-1 / LVC / HVC / バランサー / トット / Test / 丸徳） | 同上 | W | UFAinput（送信時・計算高さ操作時）、UFNameList（起動時）、UF1 の印刷時など |
| Measured.log | `HistoryPath & "<ﾌｫﾙﾀﾞ>\<ﾌｫﾙﾀﾞ>_Log\Measured.log"` | 同上 | W | LogOut / LogOutCombo（補助）。担当外 |
| NextPath\Access.log | `NextPath & "\" & "Access.log"`（`\\` が二重になる） | 同上 | W | TimeDesignation / FileDesignation（補助）が Arrbox/HistorySave から呼ばれたとき |
| PLEASE WAIT.txt | `MainPath` | 存在確認のみ | R | 上流の DB 更新中フラグ（推測） |
| VC 画像 | `PackingPath & "VC種類画像\" & <品名ｺｰﾄﾞ> & ".JPG"` | JPEG | R | VC_Picture（LoadPicture） |
| hostsend.txt | `NisPath & "\hostsend.txt"`（オフライン時は NisPath="" のため `ThisWorkbook.path\hostsend.txt`） | `Print #` による ANSI（推測）。ExpCSV で fileName="hostsend" の場合は 1 セルの文字列を 1 行出力 | W（上書き） | A_Output, UFAinput.CommandButton2 |
| result.txt | `C:\Tools\測定データ電子化\result.txt`（既定引数でハードコード） | `Line Input`（ANSI）。最終行だけを使う | R | InterRuptText。値が `"正常に送信が行われました。"` なら成功 |
| SokuteiDataTrans.exe | `C:\Tools\測定データ電子化\SokuteiDataTrans\SokuteiDataTrans.exe` | 実行 | 起動 | App_Activation（Shell） |
| A インプット CSV | `HistoryPath & "<ﾌｫﾙﾀﾞ>\<ﾌｫﾙﾀﾞ>_Ainput\" & yyyy.mm & "\" & <Lot> & "\" & yyyymmdd & "_" & <検番> & [".NN"] & ".csv"` | カンマ区切り。**各行の末尾にもカンマ**が付く。ANSI | W | A_CsvSave（Input）。同名ファイルがあれば .02～.99 を付け、99 を超えると MsgBox の後 `End` |
| 日報＆資材計算.xlsm | `ThisWorkbook.path` | Excel | 開く・マクロ実行 | UFAinput.CommandButton8 |
| sikaodr<yyyymmdd>.txt | `\\Nlmfanago02d\仕掛ﾃﾞｰﾀ`（Test: `C:\Users\NLM Administrator\Desktop`） | 固定長（`Mid(Buff,2,8)` 等） | R | Order_Update（未使用） |
| 任意のログファイル | `ThisWorkbook.path & "\" & logName` | **ADODB.Stream, Charset="Shift_JIS"**、行単位 | R | AccesslogLoad |
| ThisWorkbook 自身 | xlsm | — | **W（上書き保存）** | Name_Match（Worker:L449）。UF1 でも 3 箇所 |

### 3.3 ブック内シート（保存領域）

`<L>` = `UF1.LineName.Caption & UF1.LineNo.Caption`（例: `Test`）。サンプル CSV はすべて `_Test`。

| シート | 書く処理 | 読む処理 | 備考 |
|---|---|---|---|
| 作業保存_<L> | Posting（毎入力）, ReflectStrain, ApplyVisualStrainFromLog | Load_Text, LosdStrain, FormReflectionEx | 作業中データ（§9.2） |
| WorkerTime_<L> | Time_Posting, ApplyWorkerToTimeSheet, ApplyWorkerFromLog, EndTime_Register(C1), Name_Match(N:R), HistoryLoad(A1="履歴呼び出し") | IsOver12Hours(C1), Load_Text(B1), Name_Match, UFNameList.CB1, Print | ロット切替時に AddWS で再作成（Time_Posting L2299） |
| 検査員_<L> | WorkerSave | WorkerLoad, WorkeDayMax, Register_Check | A1:B3 |
| Daily_<L> | Daily_Create（直が変わったら再作成）, Daily_Report, UFNameList | Number_Count, UFNameList, Daily_Report4_Execution | 日報 |
| 梱包枚数計算_<L> | Number_Count（毎回再作成） | — | 検証用 |
| Lot情報履歴 | HistorySave（同じロットの行を削除して追記） | GetLotInfo, VC_Picture, OrdNoOpen, Daily_Report, Load_Text, Micro_Juge_旧, UFAinput.Initialize, Input(No_Search) | 26 列（§9.1） |
| 保存用 | MainDataSave（通常） | HistoryLoad, UF1:L4502 CheckForLot, UF1 印刷後の ExpCSV（`<L>_履歴CSV.csv`） | 114 列（§9.6） |
| 耳用 | MainDataSave（極厚） | HistoryLoad（UF管理.Mimi） | 108 列（§9.7） |
| Accesslog | AccesslogLoad, ApplyVisualStrainFromLog(U:V) | FormReflection, FormReflectionEx | §9.8 |
| Hiki_Lot | Input.No_Search（担当外） | UFAinput.Activate, HikiNo_Change | A=引当番号, B=受注番号, C=ﾛｯﾄ, D=検査番号, E=引当調整NO, F=引当数量, G=引当ﾌｧｲﾙ登録, H=包装仕様NO（Input:L112-119） |
| LOT一覧 / VC一覧 | AccUpdate 系 / Order_Update | Lotflag | いずれも未使用 |
| A_Icas_<L> / Nitiina_<L> | A_Output, UFAinput.CommandButton2/7 | ExpCSV / A_CsvSave | A_Output では出力後に削除 |
| 異常ロット用_<L> | （担当外） | Load_Text L2966（A1 が Lot と一致したら赤表示） | |
| 全ヵ所検査 / キャスク検査 / H494検査 | （担当外） | Visible 判定と D2（Lot）の参照（Worker:L747-765, SaveLoad:L1304, UFAinput:L785ほか） | |
| 全ヵ所検査保存 / 全ヵ所作業中保存 / H494作業中保存 / H494検査保存 | **担当モジュールは読み書きしない**（UFMitsubishi / Print / Cask / Mitsubishi / 補助 が使用） | — | |

---

## 4. 接続方式

### 4.1 接続文字列

1. **adoConnection（SQL_追加:L199-270）** — 全汎用 API で使用
   - `Provider`: 拡張子が csv/accdb/xlsx/xlsm/xlsb なら `Microsoft.ACE.OLEDB.12.0`、mdb/xls なら `Microsoft.Jet.OLEDB.4.0`。
   - `Extended Properties`: csv=`Text;FMT=Delimited;HDR=NO`、accdb=""、xls=`Excel 8.0;IMEX=1;Readonly=False`、xlsx=`Excel 12.0 Xml`、xlsm=`Excel 12.0 Macro`、xlsb=`Excel 12.0`。
   - **`Jet OLEDB:Database Password` = `"nlm"`（ハードコード。L250、L255）**。すべてのファイルに設定する。
   - `.Open FilePath` の後に `.CursorLocation = adUseClient(3)`（L257）。
   - `PathCheck(FilePath)=False` なら Nothing を返す（L201-204）。
2. **AccReflec（SaveLoad:L122-123）**: `"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strPath & ";"`。**パスワード指定なし**。Recordset は `rs.Open strSQL, db, 3, 1`（adOpenStatic, adLockReadOnly）で、カーソル位置は既定（サーバー側）。
3. **AddAccdb（L136-140）**: `AdoOpenCode & FilePath & "\" & fileName & ";"`（パスワードなし）。Recordset は CursorType=1（Keyset）、LockType=2（**Pessimistic**）。
4. **fldAccdb**: `ADOX.Catalog.Create AdoOpenCode & MakefilePath`（新規 DB）。その後 adoConnection で接続する。

### 4.2 接続を開く・閉じるタイミング

- **毎回開いて毎回閉じる（開きっぱなしにしない）**。GetAdoArr（L458-470）は 1 回の呼出しで Connection.Open → Recordset.Open → 読込 → Close → Close を行う。
- 1 回の GetFieldsArr と GetRecordsArr でそれぞれ別の接続を張る。例えば HistorySave 1 回の中で、SIKAHIKINOW を 2～3 回、SIKAODRNOW を 2～4 回開閉する。
- **例外**: GetAdoArr は EOF（0 件）のとき L462 で `Exit Function` するので、**rs と cn を Close しない**（オブジェクトのスコープが抜けたときの解放に任せている）。L461 の rs Is Nothing の場合も同じ。
- adoSQL は Finally で Close する。ただし接続オープン前後のエラーにはハンドラがない。
- AccReflec は正常終了時と ErrHandler の中で明示的に Close する。

### 4.3 カーソル / ロック

| 処理 | CursorLocation | CursorType | LockType |
|---|---|---|---|
| GetAdoArr / AdoArr / GetFieldDic / adoReadPageCtrl | Client（接続側で設定） | adOpenStatic(3) | adLockReadOnly(1) |
| adoRecordset の既定 | Client | ForwardOnly(0)（Client カーソルでは実質 Static） | ReadOnly(1) |
| adoSQL（列数取得） | Client | Static | ReadOnly |
| adoSQL（更新本体 L543-544） | Client | **adOpenKeyset(1)**（Client カーソルでは Static に置き換わる。ADO の仕様） | **adLockOptimistic(3)** |
| AddAccdb | 既定（Server） | Keyset(1) | **Pessimistic(2)** |
| AccReflec | 既定（Server） | Static(3) | ReadOnly(1) |

- **Client カーソル + Filter の組合せのため、毎回 `SELECT * FROM テーブル` の全件をネットワーク越しに取得してから、クライアント側でフィルタしている**（L291-301。SIKALOTNOW の仕掛テーブルにはコメントで「ﾃﾞｰﾀｻｲｽﾞ大きいため」とある）。
- adoSQL は、列数を知るためだけに `検査ﾃﾞｰﾀ` 全件を ReadOnly で一度読み込む（L529）。その後もう一度全件を開いて Filter を掛ける（L538-555）。そのため、テーブルが大きくなるほど遅くなる。

### 4.4 トランザクション

- **adoSQL**: `adoCn.BeginTrans`（L586）→ 処理 → `CommitTrans`（L679）。`On Error GoTo catch`（L583）で `RollbackTrans`（L691）。ADDNEW（AddNew/Update）、UPDATE（Command.Execute）、DELETE（Delete/Update）すべてを対象とする。
- **AddAccdb**: BeginTrans/CommitTrans（L156/L170）はあるが、`On Error GoTo catch` が**コメントアウト**（L153）されているため、ロールバックは実行されない（エラーは実行時エラーとして上位へ伝わる）。
- 読込側にトランザクションはない。

### 4.5 Access 固有機能の使用状況

| 機能 | 使用 | 箇所 |
|---|---|---|
| ADOX | **あり** | fldAccdb（Catalog.Create, Tables.Append, Columns.Append 型 adVarWChar=202）、adoTableDelete |
| DAO | なし | — |
| Compact/Repair | なし | — |
| SQL 内の IIf / Format / Nz | なし（VBA 側の IIf/Format は多用） | — |
| 日付 `#...#` リテラル | なし | — |
| オートナンバー | **使用していない**（作成テーブルは全列テキスト。ADDNEW は全フィールドに位置で値を代入する → 書込先テーブルにオートナンバー列はないと推測） | — |
| DB パスワード | **あり**（`nlm`） | adoConnection |
| ADO Recordset.Filter（`LIKE '*KEN*'` などの `*` ワイルドカード） | **あり** | adoRecordset, adoSQL |
| `Recordset.GetString(adClipString, n, vbTab, "<#(仮)改行#>")` | **あり**（全値を文字列化） | adoRecords L428 |
| PageSize / AbsolutePage | あり | adoRecords |
| 角括弧による列名引用 `[…]` | あり | 全フィルタ |
| AccReflec の WHERE で列名を引用していない（`ﾛｯﾄ番号`, `ｵｰﾀﾞｰ板丈`） | あり | SaveLoad L120 |

---

## 5. 複数 PC 同時アクセスの扱い

| 観点 | 実装 | 根拠 |
|---|---|---|
| 上流 DB 更新中の検出 | MainPath の `PLEASE WAIT.txt` の有無で NextPath（梱包課側のコピー）に切り替える。時刻窓（毎時 59:46～00:45）でも NextPath に切り替える（Arrbox, HistorySave） | SaveLoad L102, L1299; AccHIki L29; 補助 L1027-1053 |
| `FileExistence`（PLEASE WAIT が消えるまで 100ms×最大 200 回待つ） | 存在するが、AccReflec からの呼出しはコメントアウトされている（L60-62） | 補助 L1055 |
| ファイルロックの検出 | ReadCheck は存在とサイズだけを確認する（ロックは確認しない）。ログ読込前の MsgBox は「開かれているか、存在しません」だが、実際は存在しか確認していない | SaveLoad L3012 |
| 書込み競合（梱包検査ﾃﾞｰﾀ.accdb） | Optimistic ロック + トランザクション。**リトライなし・待機なし**。失敗したら Rollback して `msg.Box` でエラー番号を表示するだけ。シート `保存用` には既に追記済みのため、**シートと DB が不一致になる** | SQL_追加 L583-702, SaveLoad L4454-4568 |
| キー重複 | ADDNEW で Filter が一致すると UPDATE に切り替わる（L641-644）。キーは `[検査開始終了]`（"開始～終了" の文字列）。末尾に空白が入るため、一致しないと推測 → §14 | |
| ログ（共有フォルダ）への同時追記 | FSO の OpenTextFile(ForAppending)。`On Error Resume Next` のため、開けなければ**黙ってそのログを捨てる** | WriteLog L10 |
| Lot情報履歴 / 作業保存などのシート | ブックごとのローカル状態。PC 間では共有しない（ThisWorkbook.Save で保存） | Worker L449 |
| 上流 DB の読取 | ReadOnly だが、ACE は読取時にも `.laccdb` を作成しようとする。`\\nlmsrvngy03\Read\...` が書込み不可の共有の場合の挙動は要確認（推測: 読取専用モードで開くか、排他で開く） | §14 |

---

## 6. 業務フロー

### 6.1 ロット引当（ロット選択 → AccReflec）

UF1 で CB（ロットコンボ）を確定する（UF1:L4500-4520）。
1. シート `保存用` の直近 5 件に同じロットがあれば確認する（CheckForLot。担当外）。
2. **AccReflec**（SaveLoad:L56）
   1. UF1 の各種表示をクリアする（L64-77）。`UF1.Lot=""` なら終了。
   2. 参照パスを決める（§3.1.1-2）。ファイルがなければ MsgBox。
   3. `SELECT * FROM 仕掛 WHERE ﾛｯﾄ番号='<CB.Text>' AND ｵｰﾀﾞｰ板丈 <> '0'`（L120）。**最初の 1 行だけ**を使う（ORDER BY なし）。
   4. EOF の場合（L127-131）: SkipNoItti で UF1 の寸法・公差などを空にして終了する（L347-376）。
   5. **外注判定**: `設計_設備ｺｰｽ` に "GFS"/"GCT"/"GSS" が含まれれば（この順で判定）G_flag=True とし、メッセージを用意する（L139-149）。
   6. **設備仕様チェック**: `CheckEquipmentDb(設備名レジストリ, 製造材質, Order厚, Order幅, Order丈, 製品単重)`。G_flag の場合は BOX実績_板厚/幅/丈 を使う。結果が "可" 以外なら MsgError(…,"END") を出し、CB.Tag="end" にして終了する（L164-173）。
   7. PastHiki(0..17) を作る（L176-199。PastHiki(0) はオーダー番号の頭に `'` を付ける）。
   8. UF1 に反映する（L202-328）。
      - 材質・調質・用途・品質グレード（直角度/ラテラルボー/フラットネス）を設定する。TAIK が数値でなければ "4"、Lateral・flatnes が数値でなければ "" にする。
      - SikaSetAtu = ｵｰﾀﾞｰ板厚。
      - 寸法: 製造板厚/幅/丈（G_flag の場合は BOX実績）を "0.000"/"0.0"/"0.0" で表示する。
      - 公差: オーダー公差の＋と－の**どちらかが 0 なら**製造公差を使う（L234/243/252）。判定値 ApK = 厚+Ap、AmK = 厚-Am など。G_flag の場合は幅・丈の判定値を空にする（L264-266）。
      - **WorkTime="S" のとき HistorySave(PastHiki) を呼ぶ**（L272-274）。
      - 指示_直角度 / 指示_ﾗﾃﾗﾙﾎﾞｰ があれば、それを全角→半角に変換して表示し、色を付ける。なければ等級から計算する（Diagonal/ReDiagonal/Lateralbow_J/ReLateralbow）。
   9. キャッシュ（g_CacheOrdAtu, g_CacheC_Course）を設定して、rs と db を閉じる。
   10. VC が 0 なら VC_Clear、1 なら VC_Picture（L378-381）。
   11. Atu_OPselect（作業選択）、導電率チェック（Conductiv_Check → Conductiv_Coolingtime_Check）。
   12. メッセージの重複を Dictionary で除いてから MsgError(…,4) を出す（L419-429）。
   13. Arrbox("SIKAHIKINOW.accdb", Lot) の先頭行の受注番号で MaterialsMsgBox を呼ぶ。0 件なら PastHiki(0) で呼ぶ（L432-449）。
   14. **ゲージ判定**: IsGaugeRequired(厚,幅,丈,材質,調質,実績コース)（`品質基準.accdb`）の結果で、LZ*/LG* のキャプションを「幅ｹﾞｰｼﾞ/丈ｹﾞｰｼﾞ」か「定盤/ｹﾞｰｼﾞ」にする。どちらの場合も CBAdd(SikaSetAtu) を呼ぶ（L487-516）。
   15. GetLotInfo → JudgeLotAllocations（引当が混在する場合の判定。担当外）。
3. UF1 側に戻り、OrdNoOpen（オーダーコンボ）、PackingNoCheck、DistArray などを実行する。

**HistorySave**（SaveLoad:L1258）
1. `Lot情報履歴` で同じロットが最初に現れる行 x と最後に現れる行 y を探し、**x:y の行をまとめて削除する**（間に別ロットの行があっても消える）。
2. SIKAHIKINOW を `[ﾛｯﾄ番号]='Lot'` で絞り込んで取得する。0 件なら（Test 条件を除き）NextPath で再取得し、それでも 0 件なら MsgBox を出して終了する。
3. SIKAODRNOW の `仕掛` を**全件**取得する（VC と VCR）。
4. VCR の全行 × TempHiki の全行について、受注番号が一致した組合せごとに 1 行追記する（列は §9.1）。追記順は VCR（SIKAODRNOW の自然順）が外側、TempHiki が内側。
5. ループの各回で、キャッシュ g_CacheLotNo、g_CacheCours、g_CacheKenNo（検査番号の頭 1 文字）を設定する。
6. B 列と V 列を文字列書式にする。フォントを「メイリオ」10 にして AutoFit する。

### 6.2 作業中保存 / 再開

- **保存（Posting）**: 測定値の入力・削除のたびに UF1 などから呼ばれる。毎回 `AddWS("作業保存_<L>")` で**シートを削除して作り直し**、全量を書き直す（§9.2）。ReflectStrain(…,"保存") で歪の A/B を上書きする。Lot情報履歴の判定で直角度が不要なら R2:R7 をクリアする。
- **再開（Load_Text）**: UF1 の起動時やライン選択時（UF1:L3503, FormRelat:L6527）に `作業保存_<L>` から全項目をフォームへ戻す。
  - 数値は `Format(val(x), "0.000")` などで整形する。値が 0 の場合は空にして背景を白に戻す。
  - H494 で幅 1-299 の場合は幅を "0.00" で表示する（L2688-2704）。
  - 検査頻度（厚板/薄板/少枚数/極厚/スラブ）、検番の頭文字による厚板/薄板キャプション、住精フラグ、山数、厚 2/4 ヵ所、検入 ≧26/≦25 を復元する。
  - 最後に Lot情報履歴の VC フラグで VC 表示、WorkerTime!B1 から StartTime、異常ロット、導電率を復元する。
  - `UF1.Loading.Tag="Loading"` の間は AccessLog 出力を止める（L2520、L2988 で再開）。
- **作業者・時刻（Time_Posting）**: 測定値が確定するたびに呼ばれる（LabelCls）。WorkerTime を（SplitFlag が空で、A1 が現在のロットと違う場合に）再作成し、`(行=位置+2, 列=座標番号)` に最新直の登録者の**末尾の名前**を、`(位置+8)` に時刻 EachT を、`(位置+14)` に UF1.PakeD を書く。

### 6.3 検査結果保存（印刷時）

UF1 の印刷処理（UF1:L3630-3790）
1. IsPrinted なら PrintOnly で終了する（担当外）。
2. **EndTime_Register**（終了時刻）。
3. 先行データ票などの確認（担当外）。
4. **Name_Match**（検査員照合。ブックを保存する）→ **Number_Count**（梱包枚数照合。NG なら PrintB.Tag="stop" で中断）。
5. 印刷転送（Print。ここで `機器転送` を呼ぶ）。
6. 「印刷する？」が Yes なら WriteLog（ライン別ログ）→ PrintOut ×2 → `UF1.dataLoding.Caption=""` → **MainDataSave("保存用")** → 履歴保存（担当外）→ `保存用` を ExpCSV で `ThisWorkbook.path\<L>_履歴CSV.csv` に出力する。

**MainDataSave**（SaveLoad:L4216）
1. 外観 A/B を "A_xxx,B_yyy" 形式に結合する（L4246-4294）。
   - A だけの場合は "A_xxx,"、B だけの場合は " ,B_yyy"（**C だけは L4268 で "" ,B_yyy"（先頭空白なし）**）。
   - **F 面の分岐 L4290 は a_E/b_E を判定している**（コピーミス。F の A だけの場合でも、E の状態によって空になる）。
2. 歪は、表示中の面（AorB.Caption が Now_A/Now_B）と Tag 側の面を組み合わせて "A:<z>,B:<z>" にする。値が "__"（未入力）の場合は空にする。
3. 極厚なら `耳用`、それ以外なら `保存用` の最終行の次の行に書く（§9.6/9.7）。シート保護を解除する（再保護はコメントアウト）。
4. 書いた最終行（A 列基準の UR）を 0..EndC の配列にし、`AdoTemp(0)=ライン名` として adoSQL で `検査ﾃﾞｰﾀ` または `耳付検査ﾃﾞｰﾀ` に ADDNEW（skipmsgflag=True）。

### 6.4 「設備（MachineTransfer）」— 実体は測定機器チェックと機器名の転記

- **Micro_Juge**（測定値入力時。FormRelat:L284）
  - 選択中の機器 `UF1.Machin.Caption`（Machin_Text で 5 桁の機器番号）と、`DbReturnMicro(ライン名&No, 厚, 用途)` の出力テキストを比較する。違えば該当セルを "" にして MsgError(…,"END") を出す。
  - L-1 で用途 H494 かつ幅 <300 の場合は、座標番号 5-6（幅）だけ N5（ノギス）と比較し、それ以外の座標はチェックせずに終了する。
- **DbReturnMicro**
  1. フィールドを取得する。QualityPath の `伝送用ファイル.accdb` が空（0 件）なら PackingPath。
  2. `[ライン]='<ライン>'` で全行を取得する。
  3. Dep 指定（N5、デプス等）があれば `[機器名]` でも絞り込み、**先頭行**の出力テキストを返す。
  4. 指定がなければ、用途列に値がある行が 1 つでもあれば、CheckThickAndCode で「用途条件に一致し、かつ 厚下 ≤ 厚 ≤ 厚上」の**最初の行**を探す。
  5. 見つからなければ厚さ範囲だけで最初の行を探す。
  6. n="機器名"/"ID" ならその列を返す。
- **機器転送**（印刷シート作成時）
  - ライン: L-1 / HVC / LVC / バランサー→"大板" / トット→"大板"。
  - コンベックス: L-1 は丈 >5500 なら " C11"、それ以外は " C12"（2024.11.25 から）。HVC " C11"、LVC " C10"、バランサー " C12"、トット " C11"。
  - 書込み: `(x,y) = line & " " & 機器名`、`(x,y+2) = line & Conb`。
  - 厚が**ちょうど 300** の場合、シートの B/D/F/H 列 14-18 行の測定値に <300 があれば N5、≥300 があれば N6 として、"N5,N6"/"N5"/"N6" に置き換える。
  - H494 かつ幅 1-299: `(x-1,y+2) = "幅:ﾉｷﾞｽ , 丈:ｺﾝﾍﾞｯｸｽ"`、`(x,y+2) = line & "N5," & Conb`。

### 6.5 A インプット（UFAinput.AinButt_Click）

1. DBflag.Caption="true"（Lot情報履歴になく、SIKACOMP から引いたロット）なら送信不可。HikiNo が空なら終了。Tag="Load" なら再入を防ぐために終了する。
2. 枚数 M_Count = 枚数1×包数1（MaiFlag="2" の場合は ＋ 枚数2×包数2）。山高さ = CalcMountHeight(山数)（担当外）。
3. **1 回目の押下**（L1 が非表示、または全山が未入力）: 1 山あたりの高さ = 厚×枚数1×包数1÷山数。≤300 なら「Depth測定」、>300 なら「ｺﾝﾍﾞｯｸｽ測定」と表示し、山数ぶんの入力ラベルを出す（最大 15）。
   - 高さ入力（StartTB で Tab）: ≤300 の場合は Height_Machin（デプスの機器番号が DB と一致）が必須。>300 の場合は `#L` で始まる値（コンベックス）が必須。
4. **2 回目の押下**（全山入力済み）
   1. 入力チェック: 高さ・山数・引当No・枚数・包数。輸出（ExFrame）の場合はストア以外でケースNo が必須。外寸は 1-9 の値を禁止し、空も禁止する。ストア材（Remain=1）ならストアカードNo が必須。
   2. Height_Check(AMin, AMax, 山高さ, M_Count, オーダーNo) が **False**（許容内と推測）なら「ｱｲｶｽ転送を行いますか？」を出す。
   3. Yes の場合: WriteLog（ライン別。`*UFKen:<枚>枚_<高さ>mm * T?:<登録者>`。丸徳だけ `枚\`）→ A_Output（`Nitiina_<L>` を A_CsvSave で CSV 保存、`A_Icas_<L>` を `hostsend.txt` に出力、両シートを削除）→ App_Activation（SokuteiDataTrans.exe を Shell で起動）→ 1 秒待つ → InterRuptText（5 秒待って result.txt の最終行を確認。正常でなければ Ngflag="stop"）→ **Daily_Report(4)** → 高さ入力欄と梱包No をクリアする。
5. ストアカード（PileMatch）: 7 桁で、頭文字がロットの頭文字と一致すること。ロットの頭が S/F の場合は "L" でも可（L476-517）。積合せ（PileMatchNo=1）なら引当No&梱包No を自動設定する。

### 6.6 ログ書込（WriteLog）

- 行の形式: `Format(Now,"YYYY/MM/DD hh:mm:ss") & " " & Sepl & COMPUTERNAME & " " & Sepl & " " & msg1 & " "`
  - Sepl の既定値は ","。ただし呼出しのほぼすべてが **Sepl=""** を渡すため、実際の行は `2026/10/30 12:34:56 PCNAME  *UFKen|24.868|(S1)|○|L-1|阿部 `（先頭は空白 1 つ、PC 名の後は空白 2 つ、末尾に空白）。
- フォルダは CreateFolder で毎回作成を試みる（エラーは無視）。FSO は `Static` で保持する。
- AccesslogLoad はこの形式を読む。`|` を含む行は新形式で、値は buf(1)。`:` 区切りは旧形式で、値は `buf(UBound-3)`。ロットを含む行から次の "Lot登録" を含む行の手前までを対象にする。

---

## 7. 作業者（Worker）管理

### 7.1 検査員の登録とカード変換

- UF1 の Worker テキストに **13 桁**のカード番号が入ると（UF1:L4794）NameConv が `名簿.[No]='<13桁>'` → `名前` に変換する。未登録なら `音楽未登録` を鳴らして "" を返す。
- 直ごとに `UF1.T1/T2/T3`（名前を `_` で連結。複数人を登録できる）と `Day1/2/3`（登録日時。"yyyy.mm.dd  hh:mm"、空白 2 つ）を持つ。
- **WorkerSave**（T?_Change のたび）: `検査員_<L>`!A1:B3 = [Day1,T1],[Day2,T2],[Day3,T3]。T の全員が同じ名前なら 1 人にまとめる（Epual）。
- **WorkerLoad**: A1:B3 → Day は `Format(…,"yyyy.mm.dd  hh:mm")`、T はそのまま。FlagLabel を 0 にしてから 1 に戻す。
- **WorkeDayMax**: 検査員シートの A 列の先頭 10 文字（日付）と 13-14 文字目（時）を比較して、最も遅い直の番号を返す。**分は比較しない**。同じ値の場合は Select Case の先頭（T1 → T2 → T3 の順）が優先される。On Error Resume Next。
- **GetWorkerName**: WorkeDayMax の直の T を `_` で分けた**末尾**。
- **Register_Check**
  - n="新規": 検査員シートの A1:A3 のどれかが当日（YeaDayrCheck）でなければ True。
  - 通常: TimeCheck(STtime) の直の T または Day が空なら True。Day の先頭 10 文字が当日でなければ True。
  - シートがなければ True。

### 7.2 直（シフト）判定

- Start1..3/End1..3 は DB の `時間用` から読む（文字列）。比較は `DTime(String) > CDate(Start1)` という String と Date の比較（VBA が暗黙に型変換する）。
- 設備名（レジストリ）が L-1/LVC/HVC/Test の**いずれでもない**場合（バランサー、トット、丸徳など）は常に「日勤」で、日付はその日のまま。
- 判定（YeaDayrCheck / DayCheck / TimeCheck 共通）:
  - `DTime > Start1 And DTime < End1` → 1直（Start1 と**同じ時刻は含まない**）
  - `>= Start2 And < End2` → 2直
  - `>= Start3 And <= "23:59"` → 3直（当日）
  - `>= "00:00" And <= End3` → 3直（**前日**扱い）
  - それ以外 → 3直（当日）
- DayCheck の前日扱いの表示は `前日 yyyy.mm.dd & "  " & (現在の時+24) & ":nn"`（例 "2026.10.29  27:15"）。
- **EndTime_Register は DB を使わず 7:00 固定**: 0:00 ≤ Time ≤ 7:00 なら `Now-1日` を "yyyy/mm/dd hh:mm:ss" で記録する（**時刻はそのまま、日付だけ 1 日戻す**）。それ以外は Now。

### 7.3 作業時間・作業者の記録（WorkerTime_<L>）

| 範囲 | 内容 |
|---|---|
| A1 | ロットNo（HistoryLoad 後は "履歴呼び出し"） |
| B1 | 開始時刻 `UF1.StartTime`（yyyy/mm/dd hh:mm:ss） |
| C1 | 終了時刻（EndTime_Register） |
| A2:J2（極厚の場合は A2:L2） | 見出し: 厚1-4, 幅1, 幅2, 丈1, 丈2, 対角A, 対角B（極厚: 厚1-4, 幅1-6, 丈1, 丈2） |
| 3-8 行 × 1-14 列 | 位置 S/U/C/D/E/F（行 = 位置番号+2）× 座標番号の**測定者名** |
| 9-14 行 | 測定時刻（EachT） |
| 15-20 行 | UF1.PakeD（内容は担当外で未確認） |
| N 列以降（Name_Match） | N=番号, O=作業済検査員履歴（O1 見出し）, P=登録検査員, Q=一致無し検査員, **R1="OUT"** |

- **Name_Match は 3-7 行（S～E）× 1-12 列だけを読み、F 面（8 行目）と 13-14 列は対象外**（L357）。N:Q をクリアするが、**R1 の "OUT" はクリアしない**（L362）。
- 照合方向: 測定者（WorkerTime）が登録検査員（T1+T2+T3 の全員）に含まれなければ不一致とし、UFNameList(1) を開く。
- 最後に `ThisWorkbook.Save`（Name_Match の不一致の有無にかかわらず）。

### 7.4 日報（Daily_<L>）

- **Daily_Create**: シートがない場合、または A1 が `YeaDayrCheck & "_" & TimeCheck` と違う場合（直の変わり目）は、**シートを作り直す（前の直の日報は消える）**。2 行目に 25 列の見出し（§9.4）を書き、Daily_Report(1)、(3) を呼ぶ。
- **Daily_Report**
  - 対象行 LR = D 列の最終行+1。対象ロットが Lot情報履歴にない場合は何もしない。
  - (1): A1 = 日付_直、B1 = ライン名、C1 = 当直の T。
  - (2): ロット情報を書く（A=Lot, B=材-調質, C="厚0.000×巾0.0×丈0.0", E=前工程実績枚数, K=合紙 0→無/1→有, Y=単重 "0.00", N=VC（VC=1 の場合 "A:xx_B:yy" など）, O="キャスク"（キャスク時））。
  - (3): (1) と (2)（A が空の場合）を実行し、開始時刻を F（時）と G（分）に `'` 付きの 2 桁で書く。**分が 1 桁の場合は F 列に書く（時を上書きする。L852 の不具合）**。
  - (4): 終了時刻（H, I）、梱包No（D）、枚数（L）、包数（M）を書き、次の行に Lot と次の開始時刻を書く。MAI2/COU2 がある場合は 2 行に分ける。
  - 全ヵ所/キャスク/H494 のシートが表示中のときは Lname="" になり、`Daily_` シートがなければ何もしない（L747-750。キャスク・三菱モードでは実質何もしないと推測）。
- **Number_Count**（印刷前）: 日報の 3 行目以降で、A 列の空欄を直前のロットで埋める。ロットごとに Σ(枚数×包数) を求め、検入枚数（そのロットが最初に現れた行の E 列）と比較する。梱包数 > 検入数 なら E 列を赤くし、MsgBox を出して `PrintB.Tag="stop"`、UFNameList(3) を開く。結果を `梱包枚数計算_<L>` に書く。

---

## 8. ハードコードされた業務ルール・パス・ライン別分岐

### 8.1 パス・ファイル名・接続情報
- §0 の全パス、`伝送用ファイル.accdb`、`SIKALOTNOW/SIKAHIKINOW/SIKAODRNOW/SIKACoil.accdb`、`【測定DB】梱包\梱包検査ﾃﾞｰﾀ.accdb`、テーブル `仕掛 / 測定機器ID / 名簿 / VC厚 / 時間用 / 検査ﾃﾞｰﾀ / 耳付検査ﾃﾞｰﾀ`。
- DB パスワード **`nlm`**（SQL_追加 L250）。VBA 停止パスワード **`weekend99`**（UFNameList L40）。
- `C:\Tools\測定データ電子化\`、`SokuteiDataTrans\SokuteiDataTrans.exe`、`hostsend.txt`、`result.txt`、成功文字列 `正常に送信が行われました。`。
- `\\Nlmfanago02d\仕掛ﾃﾞｰﾀ\sikaodr` と `C:\Users\NLM Administrator\Desktop`（Order_Update。未使用）。
- `PackingPath & "VC種類画像\"`。エラーメッセージ内に `\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\VC種類画像\` が固定文字列で残っている（SaveLoad L1699 ほか。実際の参照先 PackingPath（nlmfangyshrd）と一致しない）。
- `日報＆資材計算.xlsm` とマクロ `GW_Open`。テスト値 `A111119`, `B22229`（UFAinput L217）。
- ログのフォルダ名: `L-1\L-1_Log`, `LVC\LVC_Log`, `HVC\HVC_Log`, `バランサー\バランサー_Log`, `トット\トット_Log`, `Test\Test_Log`, `丸徳\丸徳_Log`。A インプットは `_Ainput\yyyy.mm\`。

### 8.2 ライン別分岐
| 箇所 | 内容 |
|---|---|
| WriteLog の振り分け（UFAinput L52-60, L1040-1048, UFNameList L311-319） | UF1.L1Work / LVC / HVC / Bala / Tott / Test / その他=丸徳（UFAinput.AinButt は LineName キャプションの "L-1","LVC","HVC","ﾊﾞﾗﾝｻｰ","ﾄｯﾄ","Test"） |
| 直判定（Worker L150, L217, L258） | 設備名が L-1/LVC/HVC/Test の場合だけ 3 交替。それ以外は日勤 |
| 機器転送（MT L282-299） | ライン名と C10/C11/C12。L-1 は丈 5500 で分岐。バランサーとトットは "大板" |
| Micro_Juge（MT L192-207） | L-1 だけ H494 のノギス規則 |
| Micro_Juge_旧（未使用） | 機器 ID の固定値: HVC M1=10001, M2=10000, M3=10004, M4=10002, N5=10005／LVC M1=10000, M2=10001／L-1 M1=10000, M2=10001, M3=10003, N5=10005_ﾉｷﾞｽ, N6=10002_ﾉｷﾞｽ／ﾊﾞﾗﾝｻｰ M2=10000／ﾄｯﾄ M1=11001。厚さの境界 28/54/74/300、H207/H191 で 25 未満は M4 など |
| Daily_Report(1)（Worker L779-821） | ライン名の表記 L-1/LVC/HVC/ﾊﾞﾗﾝｻｰ/ﾄｯﾄ/Test。キャスクは L-1/ﾊﾞﾗﾝｻｰ/ﾄｯﾄ のみ |
| Height_Machin（UFAinput L625-633） | L-1/LVC/HVC/ﾊﾞﾗﾝｻｰ/ﾄｯﾄ/Test は DB のデプス機器番号と比較。それ以外（丸徳）は機器番号が空なら True |
| HistorySave（SL L1303-1313, L1337-1339） | 全ヵ所検査・キャスク検査の表示時は MainPath 固定。Test ラインでは該当なしのメッセージを抑止 |
| NameConv / VC_Picture | UF1.Test の場合は PackingPath を優先 |

### 8.3 業務ルール（数値・コード）
- 外注コース: `GFS`, `GCT`, `GSS`（設計_設備ｺｰｽ に含まれる）。外注時は BOX実績サイズを使い、幅・丈の判定値を空にする。
- 実績コースに `QF` を含む → "QF有り"。`SSW` → SSW_F（変数を設定するだけで、現在は使用していない。L453-456）。
- 公差: オーダー公差の＋/－のどちらかが 0 なら製造公差を使う。
- 直角度等級が数値でなければ "4"（厳しい方）。
- 指示_直角度 / 指示_ﾗﾃﾗﾙﾎﾞｰ の指示値で等級公差を置き換える（色 RGB(51,204,204)）。住精フラグの色は RGB(255,153,0)。
- H494 かつ幅 <300（1-299）: 幅はノギスで測定し、表示は "0.00"。
- 厚 300 の N5/N6 判定（<300 → N5、≥300 → N6）。
- L-1 のコンベックス: 丈 >5500 → C11、≤5500 → C12。
- A インプット: 1 山あたりの高さ ≤300 → デプス、>300 → コンベックス（`#L` 入力）。山数は 1-15。オーダーNo の 3 桁目が 1 → 輸出。ストアカードは 7 桁で頭文字ルールあり。積合せフラグは 1 か 9 のみ。梱包No が 1 桁なら先頭に 0 を付ける。0 は空にする。Remain（ストア）は 0/1。EX 外寸に 1-9 は不可。ストアの場合は EX 外寸を 0,0,0 にする。
- 少枚数は最大 5 枚（Load_Text L2922）。
- 検査頻度コード（保存用 CE 列）: 厚板=検番の頭（H/T）、薄板="<頭>:<山数>"、少枚数="<厚板キャプション>:少枚数"、それ以外="ｽﾗﾌﾞ"。
- HistoryLoad の検番判定: 左 1 文字が H かつ右 1 文字が "数" → 少枚数、H と H → 厚板、H 以外で 1/2/3 → 薄板、H 以外で 4 → 6 面表示。
- メッセージ抑止: `IsOver12Hours`（終了時刻から 1 時間以上の境界をまたいだら True）の間は、DB なしのメッセージを出さない。
- 時刻窓: 毎時 59:46-59:59 と 00:00-00:45 は NextPath を使う（Arrbox/HistorySave）。
- 夜勤の日付またぎ: Order_Update は 6:35（未使用）、EndTime_Register は 7:00、それ以外は DB の End3。

---

## 9. 入出力項目（列レベル）

### 9.1 シート `Lot情報履歴`（HistorySave が書く。1 行 = 1 引当）
| 列 | 見出し | 値の出所 |
|---|---|---|
| A | 登録日 | Now |
| B | ｵｰﾀﾞｰNo | SIKAHIKINOW.受注番号（PastHiki(0) を上書き） |
| C | ﾛｯﾄNo | PastHiki(1) |
| D/E/F | 厚/巾/丈 | 製造板厚/幅/丈（外注の場合は BOX実績） |
| G | 検査No | 検査番号 |
| H/I/J | 材質/調質/用途ｺｰﾄﾞ | 製造材質/製造調質/用途ｺｰﾄﾞ |
| K | ｺｰｽ | 設計_設備ｺｰｽ |
| L/M/N | 得意先/納入先/送り先 | SIKAODRNOW の取引先名称/納入先名称/送り先名称（上書き） |
| O/P | 前工程実績枚数/重量 | BOX実績_枚本数/良品重量 |
| Q | 合紙 | SIKAODRNOW.合紙（上書き） |
| R | VC | SIKALOTNOW.VC（0/1） |
| S/T | VC_A/VC_B | VC_表/VC_裏（**PastHiki(17)=ﾗﾃﾗﾙﾎﾞｰ等級を上書きして消す**） |
| U | 単重 | 製品単重（SIKAODRNOW） |
| V | 引当NO | 引当番号 |
| W/X/Y | 引当調整NO/引当数量/引当ﾌｧｲﾙ登録 | SIKAHIKINOW |
| Z | 包装仕様NO | SIKAODRNOW |

### 9.2 シート `作業保存_<L>`（Posting が書き、Load_Text が読む）
| 位置 | 内容 |
|---|---|
| 1 行目 A..AB（1-28） | Lot, 厚, 巾, 丈, Am, Ap, Hm, Hp, Tm, Tp, AmK, ApK, HmK, HpK, TmK, TpK, 材質, 調質, 用途, 直角度等級(TAIK), 機器(Machin), ReSize.Tag, ﾗﾃﾗﾙﾎﾞｰ等級, TaikText, LateralText, ﾌﾗｯﾄﾈｽ等級, Sumisei, Additional |
| AC(29) | 検査頻度（厚板/薄板/少枚数/極厚/スラブ） |
| AD(30) | 薄板かつ AtButt 表示中: AtCount の先頭 4 文字（"厚2ヵ所"/"厚4ヵ所"） |
| AE(31) | 薄板かつ InspeButt 表示中: そのキャプション（"検入≧26"/"検入≦25"） |
| AF(32) | 薄板かつ山数あり: 山数 |
| AG..AM(33-39) | MG, HaHiz, TaHiz, Naka, Mimi, DistortionFrame, HizCome（歪公差） |
| AN(40) | KenCount（検査回数） |
| AO(41) | HURIS（フリー反り公差） |
| AP(42) | LateralText.Tag（"有" = 技術指示ラテ） |
| 2-7 行 A..N（1-14） | 面 S..F × 測定 1-14（F 面の 11-14 は ""） |
| 2-7 行 O,P（15,16） | A 面の歪 "CB1_CB2_Z" / "CB3_CB4_G" |
| 8-13 行 O,P | B 面の歪 |
| 2-7 行 Q..T（17-20） | 対角線差, 直角度(DiagKou), 公差外判定ﾌﾗｸﾞ(NG*), ﾗﾃﾗﾙﾎﾞｰ（直角度が不要なら R 列をクリア） |
| A8 / B8 | DeleteP（削除座標）/ DeleteC（削除カウント） |
| A9 / B9 / C9..I9 / J9 | 検番 / 検入山数(YamaLB) / 検入高さ 1-7 / 高さ計算式 |
| A10..A15 / B10..B15 / C10..C14 | 外観上面 S..F / 外観下面 S..F / BH S..E |
| A16 | 前工程実績（KenH） |
| A17 / B17 | 最終熱処理工程 / QF の有無 |

※ サンプル `作業保存_Test.csv` の値と一致することを確認した（例: AB1=対角追記公差あり、AC1=薄板、O2=丈_上_0.30）。

### 9.3 シート `検査員_<L>`: A1:B3 = [Day1, T1], [Day2, T2], [Day3, T3]

### 9.4 シート `Daily_<L>`
- A1 = "yyyy.mm.dd_N直"、B1 = ライン、C1 = 検査員
- 2 行目の見出し（25 列）: `ﾛｯﾄNo, 材・調質, 製品寸法, 梱包No, 検入枚数, 開始時, 開始分, 終了時, 終了分, 作業人数, 合紙, 枚数, 包数, VC種別, etc, " 停止①記号", " 停止①時間", " 停止②記号", " 停止②時間", " 停止③記号", " 停止③時間", " 実績枚数", " 実績重量", " 作業時間", " 単重"`（P 列以降は**先頭に半角空白**がある）
- 担当コードが値を書くのは A, B, C, D, E, F, G, H, I, K, L, M, N, O, Y 列。J, P-X 列は書かない（UFNameList で手入力する想定と推測）。

### 9.5 シート `梱包枚数計算_<L>`: A=LotNo, B=梱包枚数, C=LotNo, D=検入枚数

### 9.6 シート `保存用`（MainDataSave → 検査ﾃﾞｰﾀ Field1..114）
A 検査開始終了 ("yyyy/mm/dd hh:mm～yyyy/mm/dd hh:mm"), B LOTNO, C-E 厚/幅/丈, F-K 厚-/厚+/幅-/幅+/丈-/丈+（**Am, Ap, Hm, Hp, Tm, Tp の順**）, L-W S 面（厚1-4, 巾1-2, 丈1-2, 対角A/B, 定盤歪, ゲージ歪）, X-AI U 面, AJ-AU C 面, AV-BG D 面, BH-BS E 面, BT-BX ﾗﾃﾗﾙﾎﾞｰ S-E, BY 検査員, BZ 公差修正(ReSize.Tag), CA 材質, CB 調質, CC 用途コード, CD LoadFile（履歴呼出し時 "〆"）, CE 検査頻度, CF フリー反り（"〆"）, CG-CK 外観 S-E, **CL-CP BH（MainDataSave はコメントアウトされていて書かない）**, CQ-CU 直角度 S-E, CV-DE F 面（厚1-4, 巾1-2, 丈1-2, 対角A/B）, DF F 定盤歪, DG F ゲージ歪, DH F ﾗﾃﾗﾙ, DI F 外観, DJ F 直角度。
- 各面の 12 項目は `UF_SaveArray(1,12,1,5)` の行を **`Join(" ")` → `Split(" ")`** で 1 次元にしたもの（値に半角空白が含まれると列がずれる）。その後、歪の 2 列を "A:…,B:…" で上書きする。

### 9.7 シート `耳用`（極厚。→ 耳付検査ﾃﾞｰﾀ Field1..108）
A, B, C-K は保存用と同じ。L-AA S 面（厚1-4, 巾1-6, 丈1-4, 定盤歪, ゲージ歪。16 項目）、AB-AQ U 面、AR-BG C 面、BH-BW D 面、BX-CM E 面、CN 検査員、CO 公差修正、CP 材質、CQ 調質、CR 用途、CS LoadFile、CT 検査頻度（"耳付き"）、CU-CY 外観 S-E、CZ-DD BH S-E。歪は "CB1_CB2_Z" のまま（A:/B: を付けない）。

### 9.8 シート `Accesslog`（AccesslogLoad）
A=Lot、B..O=S1..S14（行 n）、U・C・D・E 面は n+1..n+4 行、F 面は n+5 行の B..K。P=直角度、Q=ﾗﾃﾗﾙ、R=作業者、S=ログファイル名、T=理由。U1:U6 = 外観上面 S..F、V1:V6 = 外観下面（ApplyVisualStrainFromLog）。n は A 列の最終行（UR）で、**既存の最終行を上書きする**（+1 しない）。
- FormReflection は B1:Q6 を固定で読む（n=1 を前提としている）。

### 9.9 hostsend.txt / A インプット CSV
FormArray(UFAinput, 2/3)（Input:L1178。担当外）が列構成を決める。HostDataConfi は hostsend.txt の 13-20 文字目を引当No、21-22 文字目を梱包No として表示する（Input:L1399）。

---

## 10. エラー処理方式

| 箇所 | 方式 |
|---|---|
| AccReflec | `On Error GoTo ErrHandler_AccReflec`（L95）。rs と db を閉じ、MsgBox "DB読み込みエラー:"＋Err.Description。**DB 読込の後に呼ぶ HistorySave、VC_Picture、IsGaugeRequired、JudgeLotAllocations 等の実行時エラーもここに来て「DB読み込みエラー」と表示される** |
| adoSQL | BeginTrans 後に `On Error GoTo catch` → Rollback → `msg.Box`（独自フォーム `msg`）→ Resume Finally。接続の確立や列数取得（L527-533）は保護されていない |
| fldAccdb | catch で -2147217897 なら「既に存在」、それ以外は番号と内容を表示 → `Resume Finally` で**テーブル追加を続行する** |
| AddAccdb | ハンドラが無効（L153 がコメント）→ 実行時エラーになる |
| adoRecordset | 接続が Nothing なら **`End`**（L297-300）。全モジュール変数とフォームが破棄される。**ファイルがない場合（PathCheck=False）は、呼出し側の「空ならフォールバック」ロジックに到達せず、ここで終了する**（例: DbReturnMicro で QualityPath にファイルがなければ PackingPath へは行かずに End になると推測） |
| GetAdoArr | ハンドラなし。Open に失敗すると呼出し側へ伝わる |
| WriteLog | `On Error Resume Next`（全体）。ログが失われても気づけない |
| ReadCheck / IsBookOpen / Lotflag / WorkeDayMax / Ai_BeforeUpdate 系 / UserForm_Layout/Activate/QueryClose（UFAinput）/ L1_Click / CB1_Click | `On Error Resume Next` |
| RefreshDynHeightLabels | `On Error GoTo LblError/BindError` → DebugLog に出力して Exit |
| HistorySave / Posting / Load_Text / MainDataSave / Worker 系 | ハンドラなし（AccReflec から呼ばれた場合を除き、VBA の実行時エラーダイアログになる） |
| 業務エラーの通知 | MsgBox、MsgError（OutTolerance。"END" 指定あり）、`msg.Box`、音（音楽未登録 等）、DebugLog（【◆】DebugOut） |
| 強制終了 `End` | adoRecordset L299、UFAinput.Initialize L881、UFNameList（VBA 停止）L44/L345、A_CsvSave（副番 99 超） |
| 処理停止フラグ | `UFAinput.Ngflag="stop"`、`UF1.PrintB.Tag="stop"`、`UF1.CB.Tag="end"`、`UFAinput.Tag="Load"`（再入防止） |

---

## 11. グローバル変数・状態保持、日付変更時の挙動

### 11.1 グローバル / モジュール変数
| 変数 | 場所 | 用途 |
|---|---|---|
| MainPath ほか 7 個 | 初期設定 | §0 |
| SikaAtuAll, SikaSetAtu | SaveLoad L3-4 | オーダー板厚（歪コンボ CBAdd に使う） |
| g_CacheLotNo, g_CacheLotAcc（未設定）, g_CacheKenNo, g_CacheCours, g_CacheOrdAtu, g_CacheC_Course | SaveLoad L9-14 | LSelect 用キャッシュ（HistorySave と AccReflec が設定する） |
| Sika(1..10000,1..6), skipFlag, listflag, Add1, Add2, MaxRow, Time1, ws1 | AccHIki L3-12 | 固定長展開用（skipFlag と listflag の参照先は担当外） |
| RecFlag | SQL_追加 L4 | 直近の Recordset が EOF か（設定するだけで参照していない） |
| Start1..3, Start昼, End1..3, End昼 | Worker L3-10 | 直の時刻（起動時 Set_Shift） |
| GLB_INT_COL / GLB_INT_ROW | UFNameList | ListView の編集セル |
| BlinkLabel1, DynLabelEvents(1..15) | UFAinput | UI |
| Static fso | WriteLog | FSO の再利用 |

**フォームのコントロールを状態フラグとして使っている（移植時は明示的な状態に置き換える必要がある）**: `UF1.Load.Caption`("on"/"off" = イベント抑止), `UF1.FlagLabel`(0/1), `UF1.Loading.Tag`("Loading"), `UF1.dataLoding.Caption`("Load" = 履歴呼出し後), `UF1.DBflag`（空以外 = NextPath を使う）, `UF1.SplitFlag.Caption`（ロット分割）, `UF1.CB.Tag`("end"), `UF1.PrintB.Tag`("stop"), `UF1.ReSize.Tag`, `UF1.LateralText.Tag`("有"), `UF1.AorB.Caption`("Now_A"/"Now_B"), 歪コントロールの `.Tag`（非表示面の値）, `UF1.VC_L_A/B.Tag`("on"), `UFAinput.Tag`("Load"), `UFAinput.Ngflag`, `UFAinput.DBflag.Caption`("true"), `UFAinput.MaiFlag / MaterialFlag / ExMaterialFlag`, `UFNameList.Content`（モード）。

### 11.2 日付・時刻に依存する挙動
| 処理 | 日付の基準 |
|---|---|
| 直の日付（YeaDayrCheck） | 夜勤の 0 時から End3 までは前日。Daily シートの A1、検査員登録の当日判定、Register_Check に使う |
| Daily_<L> の作り直し | 「日付_直」が変わった最初の Daily_Create で、**前の直の日報シートを削除して作り直す** |
| EndTime_Register | 0:00-7:00 は日付だけ 1 日戻す |
| A インプット CSV | **暦日**（Format(Date)）。月フォルダ yyyy.mm、ファイル名 yyyymmdd（直の日付ではない） |
| ログのタイムスタンプ | Now |
| HistorySave の A 列 | Now |
| Order_Update（未使用） | 6:35 で日付を切り替える `sikaodr<yyyymmdd>.txt` |
| TimeDesignation | 毎時の時刻窓 |
| IsOver12Hours | WorkerTime!C1 との時差（"h" の境界数） |
| WorkerTime の作り直し | 日付ではなく、ロットが変わったとき（Time_Posting） |

---

## 12. Access → SQLite 移行時の具体的な注意点（このコードに即して）

### 12.1 SQL 方言・クエリ構造
1. **ADO Filter を WHERE に置き換えるときの意味の違い**: 現状は `SELECT *` で全件を取得し、Filter で絞り込んでいる。WHERE に置き換える場合、次の点が同じ結果になるかを検証する必要がある。
   - 文字列比較の大文字・小文字の区別: ADO Filter と Access の `=` はどちらも区別しない（推測。要確認）。SQLite の `=` は既定で BINARY（区別する）。ロットNo（例 `N456J52`）や機器名に小文字が入りうるなら `COLLATE NOCASE` が必要。
   - 全角と半角、カナの区別（日本語照合順序）: Access の Japanese 照合順序が幅を区別するかどうかは要確認。列名と値には半角カナ（`ﾛｯﾄ番号`, `ﾊﾞﾗﾝｻｰ`）が多用されている。
   - 末尾空白: MainDataSave のキーは末尾に空白がある。比較時に空白を無視するかどうかで ADDNEW と UPDATE の分岐が変わる。
   - `LIKE '*KEN*'`（AccHIki L143。未使用関数）: SQLite では `LIKE '%KEN%'`（ASCII の大文字・小文字は区別しない）。
2. **AccReflec L120 の `ｵｰﾀﾞｰ板丈 <> '0'`**（文字列比較）と、Arrbox の `[ｵｰﾀﾞｰ板丈]<>0`（数値比較）が混在している。列の型によって結果が変わる。SQLite の型アフィニティでは、TEXT 列の '0.0' や ' 0' は '0' と等しくならない。元の Access の列型を確認する必要がある。
3. **ORDER BY がない**: 「先頭行」を採用する箇所の結果が、DB の物理順に依存する。
   - AccReflec（複数行のうち 1 行目）、Arrbox の HikiAcc の先頭行の受注番号（MaterialsMsgBox）、DbReturnMicro の「最初に一致した行」（**どの測定器を正とするかの業務結果に直結する**）、ShiftTime の先頭 4 行、NameConv の先頭行、VC厚 の先頭行。
   - HistorySave で Lot情報履歴に書く行の順序（SIKAODRNOW の順 × SIKAHIKINOW の順）は、OrdNoOpen（ソートあり）以外の VC_Picture・Daily_Report・GetLotInfo の結果に影響する。
   - SQLite に移すなら rowid（取込順）を明示的に保存し、`ORDER BY rowid` で同じ順序を再現する必要がある。**Access 側の物理順の定義（主キー順か取込順か）は要確認**。
4. **列名**: 半角カナ、`№`、`_`、全角文字を含む。SQLite では `"…"` または `[…]` で引用すれば使える（`[]` も可）。AccReflec のように引用せずに書いた識別子（`ﾛｯﾄ番号`）も SQLite では通る可能性があるが、すべて引用することを推奨する。
5. **動的 UPDATE**（adoSQL L652-668）: 値の埋込み、`'` のエスケープなし、`WHERE` の前に空白なし。移植ではパラメータバインドが望ましい。ただし「全値を CStr で文字列化して保存する」挙動は維持する必要がある（§12.2）。
6. **列の位置で対応づけている**: 検査ﾃﾞｰﾀ / 耳付検査ﾃﾞｰﾀ は「テーブルの列順 = シートの列順」が前提。SQLite のテーブル定義でも列順を厳密に一致させる必要がある（`INSERT INTO t VALUES (...)` を列名なしで行う方式と同じ意味）。

### 12.2 データ型と値の表現
1. **読込値はすべて文字列**: adoRecords は `GetString` を使うため、数値・日付・NULL をすべて文字列で返す。NULL は ""。
   - 数値: COM の既定書式（例 `16.26`, `187.53235`。Single 型の列なら 7 桁の有効数字）。
   - 日付: Windows ロケールの短い日付と時刻（例 `2026/05/24 18:00:03`。時刻が 00:00:00 なら日付だけ）。
   - 後続処理は `val()`、`Format()`、文字列比較（例 `TempArr(i,C_用途) <> ""`、`PastHiki(16) = 0`）でこれを使っている。
   - SQLite から読むときは、**Access と同じ文字列表現を再現する変換層**が必要になる。特に浮動小数（Single の 10.512 を Double で持つと `10.512000083923340` になる）と日付書式。
   - Hiki_Lot!G1 が `10/29/2026  9:55:23 AM` という米国形式になっている → PC ごとのロケール差の可能性がある。要確認。
2. **値の中のタブ**: GetString はタブで列を区切る。値にタブが含まれると列がずれる（現行の潜在バグ）。SQLite 版で「直す」と結果が変わる可能性がある。
3. **書込みはすべて CStr の文字列**: adoSQL と AddAccdb は全フィールドに `CStr(値)` を代入する。テーブルも fldAccdb の作り方からすると全列 Text と推測。
   - Empty は ""（長さ 0 の文字列）。Access の Text 列で AllowZeroLength=No だとエラーになるはずなので、現行テーブルの設定は要確認。
   - SQLite では TEXT 型で '' を保存する。NULL ではない。
4. **AccReflec の rs.Fields(...).Value**（L136 ほか）は型付きの値（Null の可能性あり）を `String` 変数に代入している。Null を代入すると実行時エラー 94 → ErrHandler に入る。SQLite 版でも NULL の扱い（エラーにするか "" にするか）を合わせる必要がある。要確認。
5. **`val()` の意味**: VBA の `val` は先頭から数値として読める部分だけを解釈する（`"12abc"` → 12、`"1,000"` → 1）。Python の float とは違う。

### 12.3 Access 特有機能
- **DB パスワード `nlm`**（adoConnection）: SQLite には同等の機能がない（SEE や SQLCipher は標準ライブラリにない）。パスワードに依存した運用があるかは要確認。
- **ADOX（fldAccdb, adoTableDelete）**: SQLite では `CREATE TABLE` / `DROP TABLE`。本番フローからは使われていない（Test のデバッグ用途のみ）。
- **ADO Filter と Client カーソル**: 上記の通り。
- **`.laccdb` ロックファイルとページ単位のロック**: SQLite はファイル全体をロックする。
- **Provider のビット数**: cscript から ACE OLE DB を使う場合、Office（ACE）のビット数と cscript のビット数（`C:\Windows\SysWOW64\cscript.exe` が 32bit）を合わせる必要がある。

### 12.4 複数 PC からの同時書込み・ロック・トランザクション
1. **書込み対象は実質 `梱包検査ﾃﾞｰﾀ.accdb` の 2 テーブルだけ**（共有フォルダ上で、複数ラインの PC が追記する）。SIKA*NOW.accdb は読取専用（上流システムが毎時再生成すると推測。PLEASE WAIT.txt による更新中フラグ）。
2. **SQLite を SMB 共有フォルダ上に置くことは公式に非推奨**（ネットワークファイルシステムではファイルロックの信頼性が低く、破損のリスクがある。WAL モードはネットワーク FS では使えない）。現行の Access でも同じリスクはあるが、ACE はこの使い方を想定した実装になっている。移行の選択肢（各 PC のローカル SQLite と集約処理、単一の書込みサービス、Access のまま維持、など）を検討する必要がある。
3. **リトライ**: 現行はリトライしない。失敗時はロールバックしてメッセージを出すが、`保存用` シートには追記済みのまま残る。SQLite の `SQLITE_BUSY` に対して busy_timeout を設定すると「待つ」挙動になり、現行の挙動（即時エラー）と違ってくる。同等性の観点でどちらを正とするか決める必要がある。
4. **トランザクション範囲**: adoSQL は 1 回の呼出しで 1 トランザクション（1 行）。SQLite でも `BEGIN IMMEDIATE` と `COMMIT` を 1 行ずつ行えば同じになる。
5. **ADDNEW → UPDATE の自動切替**（Filter に一致する行があれば UPDATE）: 一致判定の文字列比較規則（末尾空白、大文字・小文字）を Access と合わせないと、重複行の発生や上書きの有無が変わる。
6. **接続の保持**: 現行は呼出しごとに Connection を開閉する（1 回のロット引当で 10 回以上）。cscript + VBScript 方式で 1 呼出し = 1 プロセスにすると遅くなるため、まとめて呼ぶ設計が必要。SQLite（Python の sqlite3）なら接続を保持できるが、共有ファイルを長時間ロックしないよう、短命な接続にするのが安全。
7. **全件読込の性能**: `SELECT *` で全件を取得してクライアント側で絞り込んでいる（SIKALOTNOW / SIKAODRNOW / 検査ﾃﾞｰﾀ）。SQLite で WHERE 句と索引を使えば速くなるが、§12.1-1 の比較規則の差に注意が必要。adoSQL の「列数を知るための全件読込」は `PRAGMA table_info` で置き換えられる。

### 12.5 フォールバックの意味の違い
- `GetFieldsArr` は**テーブルが 0 件だと Empty を返す**（EOF で Exit するため、列名すら取得できない）。呼出し側の「Fields が空なら代替パスへ」は「テーブルが空」の意味になる。SQLite で `PRAGMA table_info` を使って列名を取得すると、空のテーブルでも列名が返るので分岐が変わる。
- ファイルがない場合は adoRecordset の `End` で終了する（§10）。SQLite で `sqlite3.connect` を使うと、存在しないファイルは**新規作成されてしまう**。存在確認を必ず先に行う必要がある（URI の `mode=ro` / `mode=rw` を使う方法もある）。

---

## 13. 未使用コード候補（断定はしない。根拠は grep 結果）

| 候補 | 根拠 |
|---|---|
| AccHIki: Lotflag, AccUpdate, AccUpdate_Etc, Acc_Etc, Order_Update, オーダー更新, オーダー処理, Conv | 呼出しはモジュール内か、コメント行（UF1:L3372-3373, ZZ:L201）だけ。Acc_Etc は未定義の `HIK` を参照するため実行できない |
| Arrbox L60-91 | Use="Record"/"Fields" で L49-58 から必ず抜けるため到達しない（他の Use 値で呼ぶ箇所は見つからない） |
| SaveLoad L574-1256, L1760-1853 | 全行がコメント（旧 AccReflec、旧 公差表示） |
| SaveLoad: Split_G_Tag_0 | 呼出し元なし |
| SaveLoad: SSW_F（L453-456）, QfCheck の一部, g_CacheLotAcc | 設定するだけで参照していない（g_CacheLotAcc は代入もない） |
| SQL_追加: adoTableDelete, GetFieldDic, adoFieldDic, GetRecordsAdo, adoRead, adoReadPageCtrl, TxtToAccdb, XlsToAccdb, OpenXlsWs, getAddress, tgtFileDelete, AdoArr（adoRead からのみ）, RecFlag | 呼出し元なし、または未使用関数からだけ呼ばれる |
| SQL_追加: fldAccdb, AddAccdb, ExpAccdb | Test フォームの `DB元ファイル作成`（デバッグ）からだけ呼ばれる |
| MachineTransfer: Micro_Juge_旧 | 呼出し元なし（2024.2.9 に DB 方式へ切替） |
| Worker: Daily_Report4_Execution | 呼出しはコメントだけ（UFAinput L117, L1055） |
| Worker: NameConv 内の L82-140 | コメント（旧 名簿シート方式） |
| UFAinput: CommandButton1（Daily_Test 表示）, CommandButton7（固定 Lot "A111119"） | テスト用と推測（実際のフォームに配置されているかは要確認） |
| UFAinput: Frame3_Click の動的ラベルのループ（L586-595） | 中身が空 |
| UFNameList: subFrmSclSet, subLvwIniSet | 中身がすべてコメント |
| UFWorkerName | Click で UF1 を表示するだけ。フォームの表示元（WorkerName_Open）は担当外 |

---

## 14. 不明点・要確認事項

1. **SIKA*NOW.accdb の各列の型**（特に `ｵｰﾀﾞｰ板丈` が Text か Number か）。AccReflec は '0' と、Arrbox は 0 と比較している。
2. **SIKALOTNOW.accdb と梱包検査ﾃﾞｰﾀ.accdb のパスワードの有無**。AccReflec はパスワードなしで接続し、adoConnection は `nlm` を付ける。どの DB に `nlm` が設定されているか。
3. **`測定機器ID` の列名「備」**（MachineTransfer L432）は実在する列名か（「備考」なら C_備考 は常に 0。使っていないので影響はない）。`ライン` 列の値の表記（`ﾊﾞﾗﾝｻｰ` の半角/全角、`大板` と UF1.LineName の値の対応）。Micro_Juge は `LineName&LineNo` を、機器転送は "L-1"/"HVC"/"LVC"/"大板" を渡しており、キーの体系が二重になっている。
4. **MainDataSave のキー `"[検査開始終了] ='" & 値 & " '"` の末尾空白**は意図したものか。ADO Filter で末尾空白を無視して一致するかどうか（一致しないなら常に ADDNEW になる）。重複行が許容されているか。
5. **検査ﾃﾞｰﾀ / 耳付検査ﾃﾞｰﾀ の実際の列数と列名**。Field0 がライン名であること、列数が保存用（114）+1 = 115、耳用（108）+1 = 109 であること。テーブル側の列が多い場合はエラーになる。
6. **ACE が書込み不可の共有（`\\nlmsrvngy03\Read\…`）で `.laccdb` を作れない場合の挙動**と、他の PC への影響（排他で開くことがあるか）。
7. **`TimeDesignation` の時刻窓**（59:46-59:59 / 00:00-00:45）は上流 DB の更新時刻に合わせたものか。文字列連結による比較の不具合（秒が 1 桁のとき偽）を意図どおりとみなしてよいか。
8. **IsOver12Hours の閾値**（コード上は 1 時間。DateDiff は時の境界数で判定）。名前と実装の違いは仕様どおりか。
9. **HistorySave の削除範囲**: 同じロットが連続していない場合、間の別ロットの行も消える。仕様どおりか。
10. **HistorySave L1403 の不具合**（フィールド取得のフォールバック時に `VC = GetRecordsArr(...)` とレコードを代入している）。このとき後続で添字エラー → AccReflec の ErrHandler に入る。
11. **Lot情報履歴の S 列**: PastHiki(17)（ﾗﾃﾗﾙﾎﾞｰ等級）が VC_A で上書きされる。仕様か不具合か。
12. **MainDataSave の外観結合の不整合**（C 面だけ先頭空白なし（L4268）、F 面で E を判定（L4290））と、BH 列（CL-CP）を書かない（コメント）点。移植で直すか、そのまま再現するか。
13. **HistoryLoad L3561 と L3562**: F 面のラテラルの有無を 111 列（F ゲージ歪）で、直角度の有無を 113 列（F 外観）で判定している（112 列と 114 列の誤りの可能性）。
14. **Daily_Report(3) L852**: 分が 1 桁のとき F 列（時）を上書きする。
15. **Name_Match が WorkerTime の F 面（8 行目）と 13-14 列を照合しない点**、R1 の "OUT" が残り続ける点。
16. **UFNameList.CommandButton3 が `Daily_Test` に固定で書き戻す点**（他のラインでは誤ったシートになる）。
17. **ThisWorkbook（xlsm）はローカルに置かれているか、共有フォルダに置かれているか**。作業保存などのシートを PC ごとに持つ前提でよいか。`ThisWorkbook.path\Access.log` や `日報＆資材計算.xlsm` の場所もこれで決まる。
18. **PackingPath のサーバー名**（`nlmfangyshrd`）が他（`nlmsrvngy03`）と違うのは正しいか。VC 画像エラーメッセージ内のパスと一致しない。
19. **NisPath がオフライン時に未設定**の場合の hostsend.txt の出力先（ThisWorkbook.path になる）は意図どおりか。
20. **Height_Check の戻り値の意味**（False = 許容内として送信確認を出す、と読める）。担当外の Input モジュールで要確認。
21. **ログ文字コード**: WriteLog（FSO の既定 = ANSI）、AccesslogLoad（Shift_JIS と明示）、Loadtext（UTF-8 既定。Input）が混在している。PC のシステムロケールが日本語（cp932）であることが前提か。
22. **`時間用` テーブルの行数と順序**（ShiftTime は 4 行固定）。昼の行の意味（Start昼/End昼 は設定されるが、担当コードでは使っていない）。
23. **AccReflec で Null の列があった場合**（String 変数への代入でエラー 94 → 「DB読み込みエラー」）が実際に起きているか。
24. **ADO Filter が大文字・小文字、全角・半角を区別するか**（ロットNo、カード No の照合結果に影響する）。
25. **Daily_Report がキャスク・三菱モードで実質何もしないこと**（Lname="" になる）は意図どおりか。
26. **AccesslogLoad の Accesslog シート**: 最終行 n を上書きする一方、FormReflection は 1-6 行を固定で読む。複数回呼び出したときの挙動。
27. **UF1 の印刷フローで、MainDataSave の DB 書込みに失敗してもシートと CSV の保存は続行される**。業務上許容されているか（再送の手順はあるか）。
