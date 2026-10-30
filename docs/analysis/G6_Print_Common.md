# G6 構造解析レポート：印刷・帳票 / 一覧・Web / 小フォーム・部品クラス

対象: UF検査判定.xlsm（日軽稲沢・梱包課）
担当範囲: A（印刷・帳票）、B（一覧・Web・その他）、C（小フォーム・部品クラス・起動系）
根拠表記: `ファイル名:行番号`（行番号は `vba_src/*.bas` の行。UF1 等の担当外モジュールも、印刷フローを説明するのに必要な箇所だけ参照）
推測は **（推測）**、未確定は **（要確認）** と明記する。

---

## 0. 結論サマリ（最初に読む）

1. **ラベル印刷やバーコード印刷のコードは、このブックには存在しない。** `LabelCls` は「印刷ラベル」ではなく、**UserForm の Label/ComboBox コントロールのイベントを受け取る WithEvents クラス**である（`LabelCls.cls.bas:31-62`）。"バーコード" という語が出るのは、SF3 で**設備名バーコードを読み取る**箇所だけ（`SF3.frm.bas:30,43,95`）。梱包ラベルはアイカス（A_input）側が出しているとみられる（Main シートの改修履歴 L67「ラベル出力ができない」は A_input/アイカス側の話）**（推測）**。
2. 本体の帳票（UF1 の「印刷」ボタンで出るもの）は次の **2 枚。どちらも 1 部ずつ、Windows の通常使うプリンタに出す**（`UF1.frm.bas:3770,3772`）。プリンタ名を指定するコードは無い。ブックに保存されたプリンタ設定は全シートとも `EPSON LP-S180DN`（A4 モノクロレーザー）。
   - ① **検査成績票**: シート `印刷用_<ライン名>`。原紙は `印刷用原紙`。A4 横向き。
   - ② **不良要因票**: シート `不良要因印刷_<ライン名>`。原紙は `不良要因印刷原紙`。A4 横向き。不良の入力が無くても**必ず印刷される**。
   - 同じ 2 シートを PDF にもして、`C:\検査履歴\<LN>\yyyy.mm\dd日\` と `<FactoryyPath>【梱包】<LN>\yyyy.mm\dd日\` に保存する（`【UFKen_09】Input.bas.bas:1029-1058`。担当外モジュールなので概要のみ）。
3. **印刷履歴**は隠しシート `印刷履歴管理`（A=LOT番号, B=印刷日時, C=ライン名）に残る。キーは **LOT 番号だけで、ラインは区別しない**。UF1 を起動したときに **4 日以上前の行を削除**する（コメントには「30日」とあるが、コードは `DateDiff("d") > 3`。`Print.bas:176`、`UF1.frm.bas:184-185`）。印刷済みの LOT はボタンがオレンジ色の「再印刷」になり、押すと**既存のシートをそのまま再印刷するだけ**（データの再生成・保存・CSV・PDF は行わない）。
4. **Web/IE/WebDriver の対象は、社内の「包装仕様書 現場閲覧」Web だけ**（`http://nlmfangyweb1a:9084/PkgSpecWeb/#/hososhiyosho/genba/etsuran`）。包装仕様書 No を検索欄 `pkgSpecNo` に入れて検索ボタンを押す。ただし呼び出し元のボタン `UF_PackNo.CommandButton1`「ﾌﾞﾗｳｻﾞで確認」は設計時から Visible=False で、表示に切り替えるコードも無い。Edge 版（WebDriver）は呼び出しがコメントアウトされている。**事実上使われていない（推測）**。
5. 移植で影響が大きい VBA 側の不具合・癖（帳票の中身に関わるもの。§10 に詳細）:
   - `allow_Check` は、許容/保留の位置が 1 か所だけだと戻り値に「許容」「保留」の文字が入らない。そのため成績票の B10「許容公差外在り…」が**印刷されない**（`OutTolerance.bas:2611-2623` と `Print.bas:612-620`）。
   - 再印刷（PrintOnly）は、そのときシートに残っている内容を印刷する。**別 LOT のデータが残っていてもそれを印刷する**可能性がある。
   - `PageSetup.Zoom` の扱いが帳票ごとに違い、実際の印刷倍率は 100%／115% になっているはず **（要実機確認）**。
   - 不良要因の合計行は、`印刷転送` を 2 回実行すると消える（囲み数字に変換した後に、もう一度合計し直すため）。

---

## 1. 各モジュール概要

### A. 印刷・帳票

| モジュール | 種別 | 概要 |
|---|---|---|
| `【UFKen_08】Print.bas.bas`（1460行） | 標準 | ①印刷履歴（隠しシート `印刷履歴管理`）の管理と UF1 の印刷ボタンの色制御、②検査成績票シート `印刷用_<Lname>` の作成（`印刷転送`）、③罫線・見出しの検査形態別の調整（`印刷用罫線`）、④ページ設定（`印刷範囲`/`印刷範囲2`/`不良要因印刷範囲`）、⑤原紙コピー（`Pesute`、作業中保存シート作成にも共用）、⑥不良要因の合計と囲み数字への変換、⑦外観欄のフォント縮小（`FitTextToCell`）。**印刷そのもの（PrintOut）は UF1 側が行う。** |
| `LabelCls.cls.bas`（815行） | クラス | **ラベル印刷とは無関係。** UF1 の測定値ラベル（S/U/C/D/E/F 行）と ComboBox（CBS1〜CBF4）、UF_PackNo の Hiki/Ord/Pac ラベルのイベントを WithEvents で受ける。クリック＝測定位置の選択、ダブルクリック＝値の消去と理由入力、ドラッグ＆ドロップ（制限付き）、ComboBox 変更時のログ。 |
| シート `印刷用原紙` | 隠し | 検査成績票の原紙（§3.2）。図形（板の測定位置図）を 1 グループ含む（`drawing3.xml`、L11 付近）。 |
| シート `不良要因印刷原紙` | 隠し | 不良要因票の原紙（§3.3）。UF1 の不良要因 ListBox の選択肢の元データにもなる（`OutTolerance.bas:1591-1605`）。 |
| シート `印刷履歴管理` | 隠し | A1=LOT番号, B1=印刷日時, C1=ライン名。現物には `L346A20 / "2026.06.22"（文字列）/ TEST` が 1 行ある。 |
| シート `図` | 隠し | B2="ﾄｯﾄﾊﾞﾗﾝｻｰ"。図形（平行四辺形・番号1〜4/A/B の円・矢印・画像 image3.png）が 1 グループ。公差外データがあるとき成績票の L11 に貼り付ける（`Print.bas:651-659`）。 |
| シート `Main` | 表示 | 起動ボタン（ActiveX: CommandButton1「ﾒｲﾝﾌｫｰﾑ起動」、CommandButton2「CommandButton2」）、A1="CLOSE"（閉じてよいかのフラグ）、A2=ライン名、**I/J/L 列＝改修履歴（2020.12.16〜2026.10.14、219件）**、B 列＝開発者メモ（§3.6）。 |

### B. 一覧・Web・その他

| モジュール | 種別 | 概要 |
|---|---|---|
| `【UFKen_14】ListView.bas.bas`（521行） | 標準 | MSComctl ListView を表示するための共通処理。文字幅の計測に GDI API（`GetDC/CreateFont/DrawText`）を使う。現役で使われているのは `CreateField`/`CreateList`（UFNameList の日報一覧と OutTolerance）、`ListPosting`（ListView の内容を `Daily_<Lname>` シートに書き戻す）。 |
| `【UFKen_18】IE.bas.bas`（159行） | 標準 | 包装仕様書 Web（PkgSpecWeb）を IE または Edge(WebDriver) で開いて検索する。IE プロセスを taskkill する。**Option Explicit なし**。 |
| `WebDriver.cls.bas` / `WebElement.cls.bas` / `Capabilities.cls.bas` | クラス | **外部 OSS「TinySeleniumVBA v0.1.2/0.1.3」(uezo, MIT)**。msedgedriver（localhost:9515）と HTTP+JSON でやり取りする。 |
| `JsonConverter.bas.bas`（1124行） | 標準 | **外部 OSS「VBA-JSON v2.3.1」(Tim Hall, MIT)**。WebDriver/Capabilities からしか使われていない。 |
| `Test.frm.bas`（フォーム名 `TEST`、1161行） | フォーム | 開発者用のテストハーネス（約 35 ボタン）。隠しシート `Test`（コード名 ZZ）のボタン「テスト用」から開く。 |
| `TabletTest.bas.bas`（248行） | 標準 | タブレット向けにフォームを拡大縮小する試作。**呼び出し元が無い**。 |
| `【◆】DebugOut.bas.bas`（143行） | 標準 | デバッグ用のファイルログ `ThisWorkbook.Path\DebugPrint\<設備名>\DebugLog_yyyymmdd_hhmmss.txt`。**本番の処理からも呼ばれている**（UFAinput/UFCask/FormRelat/SaveLoad/Input/Cask/UF_PackNo で 100 か所以上）。 |

### C. 小フォーム・部品クラス・起動系

| モジュール | 種別 | 概要 |
|---|---|---|
| `msg.frm.bas`（フォーム名 `Msg`） | フォーム | 自作 MsgBox 互換（`Msg.Box`）。**呼び出しはすべてコメントアウトされており、使われていない**。 |
| `UFMsgBox.frm.bas` | フォーム | 共通の警告ダイアログ。`MsgError`（`OutTolerance.bas:716`）から表示する。閉じたときの後処理（UF1 を再表示する、`End` で VBA を終了する）を持つ。 |
| `UFProgress.frm.bas` | フォーム | 進捗バー（MSComctl ProgressBar）と点滅する「ﾃﾞｰﾀ書込中....。」ボタン。 |
| `UF_PackNo.frm.bas` | フォーム | 包装仕様書の画像（JPG）ビューア。引当No/オーダーNo/包装仕様書No を最大 5 行表示する。 |
| `UFConfi.frm.bas` | フォーム | 公差外のときの処理選択（公差外/許容/再測定/保留）と、測定値を消した理由の選択。 |
| `SF3.frm.bas` | フォーム | 初期設定。設備名バーコードとオンライン/オフラインをレジストリに保存する。 |
| `clsBlinkLabel.cls.bas` | クラス | CommandButton の文字を点滅させる（Sleep と DoEvents のループ）。AddinBox 由来のコード。 |
| `clsDynLabel.cls.bas` | クラス | UFAinput で動的に作る高さ表示ラベルのクリック/ダブルクリックを処理する。 |
| `ControlInfo.cls.bas` | クラス | コントロールの位置・サイズ・フォントを保持する入れ物。TabletTest からしか使われていない。 |
| `Capabilities.cls.bas` | クラス | TinySeleniumVBA（上記 B）。 |
| `ZZ.cls.bas` | シートモジュール | 隠しシート **`Test`** のコード（コード名 ZZ）。ボタン SikaHiki（Access の引当とオーダーの照会）、テスト用（TEST フォームを表示）、高さチェック（`Height_Checker` を表示するが、**そのフォームはプロジェクトに無い**）。 |
| `StringBuilder.cls.bas` | クラス | 文字列を連結するクラス（HAYs/dev-clips 由来）。**使われていない**。 |
| `Logic.cls.bas` | シートモジュール | シート `ﾛｼﾞ` のモジュール。中身は空。 |
| `ThisWorkbook.cls.bas` | ブックモジュール | Open（パス設定、他ブックの開閉を監視）、BeforeClose（Main!A1="CLOSE" のときだけ閉じてよい）、他ブックの開閉をログに残す。 |
| `ThisWorkbook1.cls.bas` / `Main1.cls.bas` | **クラスモジュール** | PROJECT ストリーム上は `Class=`。ブック/シートのモジュールをエクスポートして再インポートした残骸。インスタンスを作る箇所が無く、イベントも発生しない **→ 使われていない**。 |
| `Main.cls.bas` | シートモジュール | Main シートのボタン。CommandButton1 → `初期設定値確認` → `UF1.Show vbModeless`。CommandButton2 → `Form_1.Show`（**Form_1 はプロジェクトに無い**）。 |
| `初期設定.bas.bas` | 標準 | ネットワークパスの決定（オンライン/オフライン）、レジストリ設定の読み書き、SF3 の表示。 |

---

## 2. Sub/Function 一覧

### 2.1 `【UFKen_08】Print.bas.bas`

モジュール変数: `Public PrintedLots As Object`（Scripting.Dictionary。キー=LOT番号、値=印刷日時）、`Public Const FLAG_SHEET_NAME = "印刷履歴管理"`（L3-4）

| 名前（行） | 公開 | 引数 | 役割 | 呼び出し元 → 呼び出し先 |
|---|---|---|---|---|
| InitializePrintFlags (8) | Public | なし | Dictionary を作り、履歴シートを読み込む | UF1.Initialize(`UF1:182`)、IsPrinted、SetPrinted、ClearPrinted → LoadPrintHistory |
| CreateFlagSheet (18) | Private | なし | `印刷履歴管理` が無ければ作り（xlSheetVeryHidden）、見出し A1:C1 を書く | Load/Save/CleanOldHistory → WsCheck, AddWS |
| LoadPrintHistory (31) | Public | なし | 2 行目以降の A 列（LOT）と B 列（日時）を Dictionary に入れる | InitializePrintFlags |
| SavePrintHistory (56) | Public | LotNo, Lname | 既存の LOT なら B 列を `Now` に更新、無ければ最終行の次に A=LOT, B=Now, C=Lname を追記 | SetPrinted |
| IsPrinted (96) | Public Function→Boolean | LotNo | 毎回シートから読み直したうえで、Dictionary に LOT があるかを返す | UF1.PrintB_Click(`UF1:3634`)、UpdatePrintButtonColor |
| UpdatePrintButtonColor (102) | Public | LotNo | 印刷済みなら PrintB をオレンジ RGB(255,165,0)・「再印刷」にし、PrintBreset を表示。未印刷ならグレー RGB(192,192,192)・「印刷」にし、PrintBreset を非表示 | CheckAndUpdateButtonColor、UF1.PrintBreset_Click |
| CheckAndUpdateButtonColor (119) | Public | なし | UF1.Lot.Text が空でなければ上を呼ぶ | UF1.Initialize(`:247`)、印刷完了後(`:3833`)、Lot 確定時(`:4514`) |
| SetPrinted (126) | Public | LotNo, Lname | Dictionary に Now を入れ、SavePrintHistory を呼ぶ | UF1.PrintB_Click(`:3831`)、PrintOnly(`:3590`) |
| ClearPrinted (133) | Public | LotNo | Dictionary とシートから、その LOT の行を 1 行削除する | UF1.PrintBreset_Click(`:3609`) |
| CleanOldHistory (160) | Public | なし | B 列が日付で、`DateDiff("d", 印刷日, Date) > 3` の行を削除する | UF1.Initialize(`:185`) |
| 印刷転送 (184) | Public | なし | **検査成績票シート `印刷用_<Lname>` を作る本体**（§3.2）。不良要因シートの用意と、その印刷設定も行う | UF1.PrintB_Click(`:3729`)、UF1.TestButton_Click(`:1886`、非表示ボタン) → EndTime_Register, AddWS, Pesute, UF_Array, UFtag_Array, FitTextToCell, 印刷用罫線, 機器転送, WorkeDayMax, IsEmptyArray, MsgError, WsCheck, allow_Check, UR, 印刷範囲/印刷範囲2, 不良要因印刷範囲 |
| 印刷用罫線 (736) | Public(既定) | n: "厚板"/"薄板"/"少枚数"/"極厚"/"スラブ" | 検査形態ごとに行見出し・列見出し・罫線・ラテラルボー/直角度/差の列の配置を変える（§3.2.4） | 印刷転送 → Lateralbow_Check |
| 印刷範囲 (1137) | Public | Wor（検査員文字列） | 公差外データ**あり**のときのページ設定（A1:O(n+1)、極厚は A1:P(n+1)） | 印刷転送 |
| 印刷範囲2 (1168) | Public | Wor | 公差外データ**なし**のときのページ設定（B1:L(n+1)、Zoom=115、水平中央） | 印刷転送 |
| Pesute (1202) | Public | n=原紙名, m=貼付先シート名 | 原紙の全セルをコピーし（`xlPasteAllUsingSourceTheme`）、背景を塗りなし・文字色を黒に戻し、シートごとにズーム・枠線・原紙の保護を設定する | 印刷転送, UF1(:2010,2013,3745,3748), UFMitsubishi(10か所), OutTolerance.SelectionInput(:2377,2380), Cask(3), Mitsubishi(2) |
| 不良要因印刷範囲 (1278) | Public | Wor, Optional LotN="" | 不良要因票のページ設定、合計計算、囲み数字への変換。LotN が空なら UF1 用、空でなければ UFMitsubishi 用（フッターなし） | 印刷転送(:726), UFMitsubishi(:186,1441), Mitsubishi.bas(:737,1224) → CalculateColumnSums, ConvertNumbersToString |
| CalculateColumnSums (1337) | Public | ws | C〜X 列（3〜24）について、行 3〜8 の数値を合計して行 10 に書く | 不良要因印刷範囲 |
| ConvertNumbersToString (1361) | Public | ws | C3:AA8 と C10:AA10 の数値を CircledNumber（①〜㊿）に置き換える。行 10 の 0 は空にする | 不良要因印刷範囲 → CircledNumber(`補助:1205`) |
| Call_BordersLineStyleWeightCopyPaste (1397) | Public Function | CopyRange, PasteRange | 罫線をコピーする | **呼び出し元なし（未使用候補）** |
| FitTextToCell (1412) | Public | TargetCell | 一時テキストボックスで文字の大きさを測り、外観欄のフォントを決める（§3.2.6） | 印刷転送(:506-508) |

関連する UF1 側（担当外。印刷フローを理解するのに必要な範囲だけ）: `PrintOnly`(UF1:3574)、`PrintBreset_Click`(3599)、`PrintB_Click`(3619-3846)、`Conductiv_Click`(591、導電率測定用紙の印刷)、`Cle_Click`(567、不良要因の入力クリア)。

### 2.2 `LabelCls.cls.bas`

WithEvents フィールド: ObjS/ObjU/ObjC/ObjD/ObjE/ObjF（Label）、ObjCbS〜ObjCbF（ComboBox）、ObjHiki/ObjOrde/ObjPack（UF_PackNo の Label）。`Cont`, `PacCont`（インデックス。代入されるだけで参照されない）。

| 名前（行） | 公開 | 役割 | 呼び出し先 |
|---|---|---|---|
| EvHiki/EvOrde/EvPack (65-76) | Public | UF_PackNo のラベルを登録する | — |
| ObjHiki/Orde/Pack_Click, _DblClick (78-101) | Public | その行を赤字で選択表示し、包装仕様書 JPG を読み込む | Loop_Label(`FormRelat:6767`), JPGLoad(`FormRelat:6876`) |
| EvS〜EvF (103-126) | Public | UF1 の測定ラベルを登録する（UF1:273 付近で S1〜S14,U,C,D,E 1〜14、F 1〜12） | — |
| ObjS〜ObjF_Click (128-163) | Public | 測定位置を選ぶ（黄色背景・赤枠、INP.Tag にラベル名を入れ、LoopJudg/HizLoopJudg を実行）。その後 Manual→StartTB にフォーカスを移す | TextTag(`FormRelat:1510`) |
| ObjS〜ObjF_DblClick (166-341) | Private | **測定値を消す**: 時間と作業者の転写をクリア、対角差をリセット、Caption/Tag を空にする、消した回数を数える（`EraseCount`）、背景を白に戻す、NG フラグを消す。消す前の値が空でなく `UF1.DeleteC >= 1`（EraseCount の後は常に 1 以上）なら、**理由入力フォーム UFConfi を開く**（品証規定 Q-S-QA-DIM-174、コメント L191-193） | Time_Posting(`SaveLoad:2288`), DiagonalErase, EraseCount(`FormRelat:6911`), NGErase, ConfiReason_Open(`OutTolerance:646`) |
| ObjS〜ObjE_MouseMove (345-421) | Private | ドラッグ開始。**`UF1.DBflag` が空なら何もしない**（「品証許可出ず_2022.04.25」）。F 行には無い | ObjMoveText |
| ObjS〜ObjE_BeforeDragOver (423-526) | Private | ドラッグ元の Caption を消し、ログを書く（DBflag 条件付き） | WriteLog |
| ObjS〜ObjE_BeforeDropOrPaste (528-623) | Private | 「Dataをドロップしますか？」で Yes ならログを書いてから値を入れる（DBflag 条件付き） | WriteLog |
| EvComS〜EvComF (626-649) | Public | UF1 の ComboBox（CBS1〜4 等）を登録する | — |
| ObjCbX_DropButtonClick (652-669) | Private | `UF1.Load.Caption="off"`（他のイベントを抑止） | — |
| ObjCbX_Change (671-814) | Private | 履歴の読み込み中（`Loading.Tag="Loading"`）と `FlagLabel=0` のときは何もしない。値が空でなければ `[A/B] CBS1:値:ライン` を **Measured.log（ライン別）と Access.log（ブックのフォルダ）の 2 か所**に書く | WriteLog |

### 2.3 `【UFKen_14】ListView.bas.bas`

| 名前（行） | 公開 | 役割 | 呼び出し元 |
|---|---|---|---|
| MeasureTextWidth (77) | Public Function | GDI で文字列の幅（px）を測る。`font_height=1` で呼ばれるので、実際にはほぼ下限値になる **（推測）** | LvRec, CreateList |
| LvField (95) | Public | 列見出しを作る | CreateAll（下記） |
| CreateAll (131) | Public | 見出しと明細を一括で作る | `SQL_追加:475,518`（adoRead/adoReadPageCtrl。**どちらも呼び出し元なし**）→ 実質未使用 |
| CreateField (145) | Public | 列見出しを作る（高さ/幅を 0 にして高速化） | UFNameList(:271,365,374), OutTolerance(:570) |
| LvRec (198) | Public | 明細を作る | CreateAll |
| CreateList (282) | Public | 明細を作る（`On Error Resume Next`）。列幅=最大文字幅×7.95（6〜30 の範囲に丸める）。幅が 12 を超える列は左寄せ、それ以外は中央寄せ。Tag に件数と処理時間を入れる | UFNameList, OutTolerance |
| LvHide (382) / ViewSelNumArr (389) / GetLvData (432) / GetList (473) / ListPatch (497) / GetWidth (414, Private) / IsHalf (424, Private) | — | 列の非表示、選択中のインデックス取得、選択データの取得、フォルダの一覧、表示ずれ対策など | **外部からの呼び出しなし**（ListPatch は CreateAll からのみ） |
| ListPosting (504) | Public | UFNameList.LV1 の全行を `Daily_<Lname>` の 3 行目以降に書き戻す（先に A3 から 20+件数行×列数をクリア） | UFNameList.CommandButton2(:265)「日報ﾃﾞｰﾀを修正しますか？」 |

### 2.4 `【UFKen_18】IE.bas.bas`

モジュール変数: `objShell`, `objExec`, `Public AppName/Path_webDriver/Path_target As String`, `Public Driver As New WebDriver`。`objIE` と `evt` は宣言されていない（Option Explicit が無いので暗黙の Variant になる）。

| 名前（行） | 公開 | 役割 | 呼び出し元 |
|---|---|---|---|
| EdgeSiteSetTextData (17) | Public | msedgedriver_97.0.1072.69.exe を起動し、PkgSpecWeb を開き、`name=pkgSpecNo` に包装仕様書 No を入れて、class `btn btn-info c-btn-size` の 2 つ目をクリックする。最後に driver の exe を Shell でもう一度起動する（コメントでは「再読み込みの代わり」） | **呼び出しはコメントアウト**（`UF_PackNo:23`「NIS都合 edge使用不可**2022.02末」） |
| IeSiteSetTextData (56) | Public | IE（InternetExplorerMedium）で PkgSpecWeb を開く。3 秒待ってから `getElementById("pkgSpecNo").Value` に値を入れ、change イベントを発火し、class `btn btn-info c-btn-size` の 1 つ目をクリックする | UF_PackNo.CommandButton1_Click(:22)。ボタンは設計時から非表示 |
| TextDataLoad (93) | Public | 未定義変数 `ie` を参照している（実行すればエラー） | 呼び出しなし |
| ieCheck (100) / WaitIE (124) / Call_IE_WaitTime (131) | Public | IE の読み込み待ち（ieCheck は 20 秒ごとに Refresh。他の 2 つはタイムアウトなし） | 呼び出しなし |
| ie_kill (142) | Public | Shell.Application の窓に HTMLDocument があれば `taskkill.exe /F /IM iexplore.exe` を実行する | UF_PackNo.CommandButton2「閉じる」(:33)、Test.frm |
| IsExecutingIE (151) | Private | 上の判定 | ie_kill |

### 2.5 WebDriver / WebElement / Capabilities / JsonConverter（外部ライブラリ）

- **TinySeleniumVBA**（https://github.com/uezo/TinySeleniumVBA、MIT）。`WebDriver.Edge/Chrome(driverPath, "http://localhost:9515")` で driver の exe を `Shell(driverPath, vbMinimizedNoFocus)` で起動し（`WebDriver.cls.bas:208`）、W3C WebDriver の REST を `MSXML2.ServerXMLHTTP`（:568）で呼ぶ。`Enum By {ID, TagName, ClassName, Name, CssSelector, XPath}`（:172）。
- **VBA-JSON v2.3.1**（Tim Hall、MIT）。使っているのは WebDriver/Capabilities だけ。
- このシステム固有の改変は見つからなかった（ヘッダーのバージョン表記だけ確認。差分の照合はしていない）。**移植は不要（推測）**。

### 2.6 `Test.frm.bas`（TEST）

開発者用。主なボタン: Access/DB の照会テスト（CommandButton34 で歪公差 DistArray、24 でソリ4級、27 で品質基準、37/38 で機器 DB）、レジストリ操作（18=AIM 設定、21=HVC 設定、22=表示して SF3 を開く、23=削除）、Access.log を `過去Log`/`Accesslog` シートに再展開（6, 19）、PDF 出力テスト（8）、梱包枚数計算テスト（7）、IE テスト（Main/CreateIE で google.co.jp を開く、15=ie_kill）、ドラッグ移動 UI の実験（btn0, ListBox1-3）など。`UserForm_Initialize` で `Application.Visible=True` にし、`Setting_Path` を呼ぶ（:1151-1155）。

### 2.7 `TabletTest.bas.bas`

`VC描写有無`(7)、`AdjustFormToScreenSize`(13、Application.Width/Height に合わせて最大 2 倍に Zoom。VC 表示時は高さ+158、新規時は高さ=545)、`AdjustFrameSize`(68, Private)、`AdjustControlSize`(84, Private)、`SaveOriginalControlInfo`(143)、`ResizeAndRepositionControls`(159)、`PadWithSpaces`(242, Private)。**どれも外部からの呼び出しなし。**

### 2.8 `【◆】DebugOut.bas.bas`

| 名前（行） | 役割 |
|---|---|
| DebugLog(message, useWorkbookPath)(11) | 初めて呼ばれたときに Initialize する。`yyyy/mm/dd hh:mm:ss | message` を追記し、Debug.Print にも出す。On Error Resume Next |
| InitializeDebugLog(15→47) | レジストリの設備名を使い `ThisWorkbook.Path\DebugPrint\<設備名>\DebugLog_yyyymmdd_hhmmss.txt` を作り、ヘッダーを書く。**Excel セッション中は 1 ファイル**（`End` 実行や Reset で変数が消えると新しいファイルになる） |
| CloseDebugLog(112) / ResetDebugLog(141) | **呼び出しなし** |

### 2.9 C 群のフォームとクラス

| モジュール | プロシージャ（行） | 内容 |
|---|---|---|
| Msg | C1/C2/C3_Click(19-33)、NumSet(35)、UserForm_QueryClose(53)、**Box(msg, ModeNo, title, FS)(70)**、BottomSetting(173) | 外部呼び出しはすべてコメントアウト（`SQL_追加:202,298,645`）。BottomSetting で `Case 2` が重複しているため、3 ボタン表示が実行されない不具合あり（:190,197） |
| UFMsgBox | UserForm_Layout(16)、Activate(24)、Initialize(35)、**QueryClose(46)**、CommandButton1_Click(73) | 表示側は `MsgError`。QueryClose: Tag/FlLabe が "END" なら Excel を最大化・表示し、`Unload Me` → **`End`（VBA の全状態をリセット）**。それ以外は Unload したうえで、"Ainput" 以外なら `UF1.Show vbModeless` |
| UFProgress | Activate(15: Tag≠"1" なら Blink 2)、Click(20: BlinkStop)、Initialize(23: ×ボタン削除、CommandButton1 を点滅対象に登録)、Terminate(29)、QueryClose(32、中身なし)。`Public IsCancel`(14) は**未使用** | 進捗の更新は `ProgressBar_count(x, Count)`(`補助:694`、Max=1000、Label1 に「n%完了」) |
| UF_PackNo | CommandButton1_Click(19: IE で検索)、CommandButton2_Click(25: 「確認」CB1 にチェックが無いと「確認後に押してください。」。ある場合は ie_kill → Unload → UFMitsubishi/UFCask が表示中でなく、かつ UFMsgBox が非表示なら UF1.Show)、SpinButton1_SpinDown/Up(51/76: 画像を 10% ずつ縮小/拡大。ただしローカル変数を毎回初期化するので、実際は 1 段階しか変わらない **（推測）**。SpinButton は設計時から非表示)、Initialize(102: LabelCls を 15 個登録、GetZoom、InitialSetting、DebugLog)、**Activate(118)**: `Lot情報履歴` から LOT が一致する行（最大 5 行）を集めて引当No でソートし、Hiki/Ord/Pac の各行に表示 → PackSelect(1行目) → MaterialsUpperPart で上部緩衝材の MsgBox（「アングル」のときは当板の有無も）、QueryClose(188: ×ボタンでは閉じない) | 表示側は `UFPack_Open`(`OutTolerance:520`、`QualityPath & "包装仕様書\<No>.JPG"` があるときだけ vbModal で表示) |
| UFConfi | CommandButton2_Click(16: 理由一覧 List1 を確定。「再測定」「その他」を選んだのに理由欄が空ならエラー。「板廃棄」なら消した値を戻して → LoopJudg → NG_Posting → NG_Erase。ログ「<位置><番号>_<元値>:理由選択:<理由>:<検査員>」)、Layout/Activate(91/101: Excel を非表示、HVC/Test のときだけ自動入力ボタンを表示、起動ログ)、Initialize(144: Reasons_Add)、QueryClose(159: Posting + UF1 のサイズ保存)、**CommandButton1_Click(171)「送信」**: TB1 の文字列で分岐。公差外 → NG_Posting + NG_Erase、許容 → NG_Pass、再測定 → DelC≥2 なら理由 List2 を表示、それ以外は NG_Back、保留 → NG_Pass(,"保留")、それ以外は「指定文字列を入力してください」。NewHB1〜4_Click(274-285: TB1 に 公差外/許容/再測定/保留 を入れる) | 表示側は `ConfiReason_Open`(値を消したとき、P2 ページ) と `Confi_Open`(NG 判定時、P1 ページ)（`OutTolerance:646-715`） |
| SF3 | CB1_Click(26)、CommandButton1_Click(48)、LineChenge_Change(52: 10 文字になったら LineConv)、OP1/OP2_Click(62/65)、Activate(68)、Initialize(72)、**LineConv(88)** | §3.5 |
| clsBlinkLabel | Property Let cmd(53)、**Blink(BlinkCount, Color1, Color2, Buzzer, Sound)(64)**、BlinkStop(181)、mctlLabel_Click(188) | 点滅対象は MSForms.CommandButton。ForeColor を「背景色⇔元の色」で切り替え、0.5 秒×2 を 1 回として BlinkCount 回繰り返す。BlinkCount≤1 なら何もしない |
| clsDynLabel | Init(16)、m_lbl_Click(24: Caption と Tag を空にする)、m_lbl_DblClick(31: Tag を空にして UFAinput.StartTB にフォーカス) | UFAinput:1418 で 1〜15 個を作る |
| ControlInfo | Left/Top/Width/Height/FontSize/IsMultiPage/PageCount の Get/Let、InitFromControl(69) | TabletTest からだけ |
| StringBuilder | Capacity/Length、Clear/Append/Insert/ToString（いずれも Friend） | 未使用 |
| ZZ（シート`Test`） | CommandButton1_Click(SikaHiki: T1 の LOT で `SIKAHIKINOW.accdb` → `SIKAODRNOW.accdb` を照会し、受注番号・VC_表・VC_裏・梱包単位_梱包形状_輸出先国名 を H5:?8 に書く)、CommandButton2_Click(テスト用: Excel を表示 → TEST.Show → VisibleSheets)、CommandButton3_Click(高さチェック: `Height_Checker.Show` → **存在しないフォーム**)、T1_KeyPress(英数字だけ受け付け、大文字にする) | |
| ThisWorkbook | Workbook_Open(12)、xlApp_WorkbookOpen(30)、xlApp_WorkbookBeforeClose(46)、Workbook_BeforeClose(62) | §5, §8 |
| ThisWorkbook1 / Main1 | ThisWorkbook と Main の古い複製（Main1 には CommandButton3 → `UserForm1.Show` もあるが、そのフォームは存在しない） | 未使用 |
| Main | CommandButton1_Click(11)、CommandButton2_Click(26: `Form_1.Show` → 存在しない) | |
| 初期設定 | Setting_Path(18)、AIM(49)、HVC(53)、test取得(56)、DelSet(59)、初期設定値確認(66) | §3.5 |

---

## 3. 印刷・帳票

### 3.1 帳票一覧（いつ・何を・どこへ）

| # | 帳票（シート） | いつ | 出力先 | 部数 | 原紙 | 作るコード |
|---|---|---|---|---|---|---|
| 1 | **検査成績票** `印刷用_<Lname>` | UF1「印刷」ボタン（`PrintB_Click`）で入力チェックを通過し、「印刷する？」に Yes | 通常使うプリンタ（指定なし）。保存されている設定は EPSON LP-S180DN。続けて PDF 2 か所 | 1 | `印刷用原紙` | `印刷転送`、`印刷用罫線`、`印刷範囲`/`印刷範囲2` |
| 2 | **不良要因票** `不良要因印刷_<Lname>` | 1 の直後に**必ず**出す（不良の入力が無くても） | 同上。PDF も 2 か所（ファイル名 `<LOT>_不良要因.pdf`） | 1 | `不良要因印刷原紙` | 入力は `SelectionInput`(OutTolerance:2356)。ページ設定と集計は `不良要因印刷範囲` |
| 1' / 2' | 再印刷 | 印刷済み LOT（`印刷履歴管理` にある）で「再印刷」を押し、「このLOTは印刷済みです。再印刷しますか？」に Yes | 同上（**PDF・保存・CSV は無し**） | 各1 | — | `PrintOnly`(UF1:3574)。成績票の D2 に「再印刷」と書いてから印刷 |
| 3 | 導電率測定用紙 `導電率測定` | UF1「導電率ﾌﾗｸﾞ」ボタンが赤のとき、「導電率測定用紙を印刷しますか？」に Yes | 通常使うプリンタ | 1 | 外部ブック KenBook からシートをコピー（`OutTolerance:2299-2334`、担当外） | `ConductivLoad` |
| 4 | 全ヵ所検査 / H494検査 / 5N / 転送済 / キャスク | UFMitsubishi / UFCask の各画面 | 同上 | 1 | `全ヵ所検査用原紙`/`H494_原紙`/`5N原紙`/`三菱受託原紙`/`キャスク原紙` | 担当外（`UFMitsubishi.frm.bas:173,190,1429,1444,1562,1714,1729`、`UFCask.frm.bas:143`）。不良要因票は `不良要因印刷範囲(<検査員>, ...)` を共用し、LotN が非空なのでフッターは出ない |

- `PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False`。直前に `Application.PrintCommunication = False` にしており、**PrintB_Click/PrintOnly の中では True に戻していない**（`UF1:3768,3580`）。印刷範囲を設定する関数の中では True に戻している。
- 印刷の順序は「成績票 → 不良要因票」。

### 3.2 検査成績票 `印刷用_<Lname>` の作成（`印刷転送` Print.bas:184-734）

#### 3.2.1 全体の流れ（PrintB_Click → 印刷転送）

1. `UF1.PrintB_Click`（UF1:3619）
   - 2 度押し対策: `PrintB.Enabled=False`、`Load.Caption="off"`
   - `IsPrinted(Lot)` が真 → `PrintOnly` へ進んで終了
   - `EndTime_Register`（終了時刻。§3.4）
   - `AdvanceCheck(Lot)` が真なら「先行データ票はありますか？」（ありますか.wav）。No なら中断
   - 入力チェック（順に判定し、最初に該当したものを `MsgError(..., Lname, 1)` で表示して中断）:
     - 定盤歪の未入力 → 「<位置>入力なし」
     - ゲージ歪の未入力 → 同様
     - 歪 A/B の不整合 → WithorWithout_Gauge のメッセージ
     - ラテラルボーの未入力 → 「<位置>:ﾗﾃﾗﾙﾎﾞｰ入力なし」（条件: `UF1.Sumisei.Caption <> """"` かつ「ﾗﾃﾗﾙﾎﾞｰ検査不要」でない）
     - 直角度の未入力 → 「:直角度入力なし」（用途 R115 のときだけ）
     - 外観の未入力 → 「:外観入力なし」
     - 検入高さ → 「検入高さor枚数　入力なし」
     - 山数 → 「山数の入力がありません」（Yama が表示されていて空のとき）
   - UFProgress を表示、おまちください.wav、画面を「少々お待ちください」で隠す（frameblindfold）
   - `Name_Match`、`Number_Count`。`PrintB.Tag="stop"` なら中断
   - **`印刷転送`**。`Tag="stop"` なら「ｷｬﾝｾﾙ処理」
   - 不良要因シートが無い、または B1≠Lot なら、作り直す（`AddWS`+`Pesute`）
   - 印刷.wav →「印刷する？」に Yes
     - ログ `*UFKen:印刷 T1:.. T2:.. T3:..`
     - **PrintOut ×2**
     - `dataLoding.Caption=""`
     - `MainDataSave("保存用")`
     - `履歴保存`（PDF と原本の保存）
     - CSV（`ExpCSV`。耳付きは `耳用` シートを使う）を、ローカル・`HistoryPath<ライン>\<ライン>_履歴`・`FactoryyPath【梱包】<LN>` に出す
     - `sh.Visible=False`
     - `ThisWorkbook.Save`
     - `SetPrinted`、ボタン色の更新、完了.wav
   - 最後に **`Application.Visible = False`**（UF1:3846）
2. `印刷転送` の中の手順:
   - `EndTime_Register` をもう一度呼ぶ（:190）
   - `AddWS("印刷用_<Lname>")`。同名のシートがあれば**削除してから**先頭に新規作成し、非表示にする（`補助:466-484`）
   - `Pesute("印刷用原紙", …)`
   - セルを書き込む → 罫線 → 機器 → 検査員 → ライン → 許容/保留 → 書式 → 公差外データ → ページ設定 → 不良要因票の準備とページ設定 → 改行コードの置換（C12:F16 の vbCrLf を vbLf に。PDF で改行記号が出るのを防ぐ。:729-730）

`Lname = UF1.LineName.Caption & UF1.LineNo.Caption`（:188）。シート名の接尾辞であり、ヘッダー左にも印字される。

#### 3.2.2 原紙 `印刷用原紙` の固定文字（CSV より）

B2=LotNo, F2=検入高さ(mm) / C4=厚, E4=巾, G4=丈, I4=歪 / C5..H5=ﾏｲﾅｽ,ﾌﾟﾗｽ（×3）, I5=ｹﾞｰｼﾞ, J5=幅, K5=丈 / C6..H6=ﾏｲﾅｽ実,ﾌﾟﾗｽ実（×3）/ B8=検査員, E8=マイクロ, G8=コンベックス, Q8=中歪, R8=耳歪 / C11=外観, E11=定盤歪, F11=ゲージ歪, G11..J11=厚1..厚4 / B12..B16=1..5 / C18..H18=幅1..幅6, I18..L18=丈1..丈4, N18=差, P18=ﾗﾃﾗﾙﾎﾞｰ, R18=直角度 / B19..B23=1..5。
セル結合: B9:D9、C12:D12〜C16:D16。列幅: A=9, B=7.25, C〜L=11.625, M=9, N〜P=9.375。図形: 板の測定位置図のグループ（円弧・番号1〜6の円・角丸四角1〜4・矢印。アンカーは L11 付近、`drawing3.xml`）。**`Pesute`（Cells.Copy → PasteSpecial）でこの図形がコピーされるかは要確認**（Excel の「セルと一緒にオブジェクトをコピー」設定に依存）。

> 注意: C5:H5、C6:H6 の「ﾏｲﾅｽ/ﾌﾟﾗｽ」等の見出しは、`印刷転送` が**公差の値で上書きする**ので、印刷物には残らない。

#### 3.2.3 セル書き込みマップ（`sh` = 印刷用_<Lname>）

| セル | 内容 | 元データ（UF1 コントロール等） | 根拠 |
|---|---|---|---|
| F1 | 高さ計算式の表示 | `UF1.高さ計算式.Caption` | :318 |
| B1 | "ｻｲｽﾞ修正" | `UF1.ReSize.Tag = "有"` のときだけ | :426 |
| C2 | LOT番号 | `UF1.Lot.Text` | :427 |
| D2 | "再印刷" | PrintOnly のときだけ | UF1:3581 |
| G2 | 検入高さ合計（Single。`Val(KenHBox1)+…+Val(KenHBox6)`）。**KenHBox7 が数値**なら、代わりに "<KenHBox7>枚" と書き、F2 を "検入枚数" に変える | KenHBox1〜7 | :302-306, 511-514 |
| H2 | 山数 "6山"〜"1山"（KenHBox6→1 の順で、最初に数値だったものの番号）。KenHBox7 が数値のときは書かない | KenHBox1〜6 | :307-316, 516 |
| B3 | "履歴呼び出しから印刷" | 履歴から読み込んでいるとき | :602 |
| B4 | Lname | | :607 |
| D4 / F4 / H4 | 厚 / 幅 / 丈（サイズ） | ATU / Hab / Tak | :428-430 |
| C5 / E5 / G5 | `'-` & 厚/幅/丈 のマイナス公差（先頭の `'` は文字列にするため） | Am / Hm / Tm | :431-433 |
| D5 / F5 / H5 | `'+` & プラス公差 | Ap / Hp / Tp | :434-436 |
| C6..H6 | 実寸公差（−厚, +厚, −幅, +幅, −丈, +丈） | AmK, ApK, HmK, HpK, TmK, TpK | :437-442 |
| I3 | 歪公差の参照元（フォント 6pt） | `DistortionFrame.Caption` | :487-488 |
| I6 / J6 / K6 | ゲージ / 幅歪 / 丈歪 公差 | MG / HaHiz / TaHiz | :446-448 |
| Q9 / R9 | 中歪 / 耳歪 公差 | Naka / Mimi | :449-450 |
| J7:K8 | `UF1.Naka<>""` のとき Q8:R9（見出しと値）をここにコピー | | :763 |
| L4 / L6 | "対角" / 対角公差 `Format(Val(UF1.TaikText),"0.0")`。**空なら "0.0"**。L4:L6 は太枠で、下罫線は Mimi が空でなければ二重線、空なら太線（※HURIS ではなく Mimi で判定している） | TaikText | :467-482 |
| L7 / L8 | "フリー反り" / `Format(Val(HURIS),"0.0")`（HURIS が空でないときだけ）。L7:L8 は太枠で上辺が二重線 | HURIS | :454-465 |
| B9（B9:D9 結合） | **検査員**（縮小して全体を表示） | 次の段落 | :546-604 |
| B10 | "許容公差外在り:" & 位置、または "公差外 保留処理:" & 位置 | `allow_Check`（§10 の不具合に注意） | :609-620 |
| E9 / G9 | マイクロメータ機器名 / コンベックス機器名 | `機器転送(sh,9,5,厚,幅,用途)` | :541-544、`MachineTransfer:276-` |
| C12:C17 | 外観（6 行。A 面と B 面を合わせたもの） | S〜F_UpVisual / _UnderVisual | :199-293, 501 |
| E12:F17 | 定盤歪 / ゲージ歪（A/B 面）。`_` を消す | UF_Array(15,16)、UFtag_Array | :386-423, 490, 496-497 |
| G12:J17 | 厚1〜4 | UF_Array 1〜4（S1..F4 の Caption） | :374, 492 |
| C19:L24 | 幅/丈（UF_Array 5〜14） | S5..E14、F5..F10 | :375, 493 |
| N19:N24 | 対角差 | DiagS..DiagF | :379, 494 |
| P19:P24 | ラテラルボー（値 & "/L"） | UF_Array 20 | :381, 495 |
| R19:R24 | 直角度（**用途 R115 のときだけ**） | UF_Array 18 | :380, 499 |
| L11 | シート `図` の図形を貼り付け（公差外データがあるときだけ） | | :651-659 |
| A26〜 | 公差外データ表（§3.2.5） | `異常ロット用_<Lname>` | :660-711 |

**検査員の決め方**（:546-604）:
- `WorkeDayMax`（1/2/3 直のうち登録時刻が最も遅いもの）に対応する `UF1.T1/T2/T3` を `_` で分割する。
- 空なら `MsgError("実行時エラー！検査員登録をやり直してください。(変数なし)")` → 中断。
- `WorkerTime_<Lname>` シートが**無ければ**、分割した最後の名前を入れる。ただし直後に同シートを `Set` しているので、**シートが無いと実行時エラーになる**（:558-561）。
- シートがあれば、`WorkerTime_<Lname>!A3:L7` を行優先で走査し、重複を除いた名前を出現順に `_` で連結して B9 に入れる。
- 履歴から読み込んでいない（`dataLoding≠"Load"` かつ `WorkerTime!A1≠"履歴呼び出し"`）のに空なら、`MsgError("検査員が空白です。WorkerTime上に履歴がない可能性有。 担当者に連絡してください。")` → 中断。
- 履歴から読み込んでいる場合は、UFWorkerName が表示中ならその L1（現検査員）を入れる。非表示なら空になる。

**外観の文字列の作り方**（:199-289）:
- 各面の値 v について:
  - v が空でなく「欠陥ﾅｼ」を含まない → "△"&v
  - v がちょうど「欠陥ﾅｼ」 → "〇欠陥ﾅｼ"
  - 「欠陥ﾅｼ」を含むがそれと一致しない場合は、記号を付けない
- A 面（Up）と B 面（Under）の組み合わせ:
  - 両方あり → `"A:"&a & vbCrLf & "B:"&b`
  - A だけ → `"A:"&a & vbCrLf`
  - B だけ → `vbCrLf & "B:"&b`
  - 両方なし → ""

**歪の文字列の作り方**（:386-423, 496-497）:
- 定盤は `CBx1 & "_" & CBx2 & "_" & Zx`、ゲージは `CBx3 & "_" & CBx4 & "_" & Gx`（`SaveLoad:2172-2195`）。
- いま表示している面（`UF1.AorB.Caption` が "Now_A" か "Now_B"）の値と、反対の面の値（各コントロールの `.Tag`）を A/B に割り当てる。
- "__"（全部空）でなければ "A:" / "B:" を前に付け、`A & vbCrLf & B` を作る。
- 書き込んだ後、E12:F17 の "_" を全部消す。
- `UF1.LZS.Caption="幅ｹﾞｰｼﾞ"` なら、E11="幅ゲージ"、F11="丈ゲージ"。
- そうでなく `flatnesText="ﾌﾘｰｿﾘ要"` なら、F11="ﾌﾘｰｿﾘ" とし、F12:F16 の "A:"/"B:" を消す（:1127-1134）。

#### 3.2.4 `印刷用罫線`：検査形態別のレイアウト（:736-1136）

形態の判定は `UF1.厚板/薄板/少枚数/極厚/スラブ` の OptionButton（`極厚` の表示 Caption は「**耳付**」）。

| 形態 | 行見出し B12..B16 / B19..B23 | 列見出しの変更 | その他 |
|---|---|---|---|
| 厚板 | S,S1,2,E1,E / S,"",2,"",E | E18=丈1, F18=丈2, G18=対角A, H18=対角B, I18=差。N18:N24 を I18:I24 にコピー | 20 行と 22 行の罫線を引き直す |
| 少枚数 | 原紙のまま（1〜5） | 同上 | |
| 薄板 | S,U,2,D,E（上下とも） | 同上 | 山数 `Yama` で上書き: 4 山=4-S,4-E,3-2,2-2,1-S,1-E（17/24 行も使う。6 回検査）、3 山=3-S,3-E,2-2,1-S,1-E、2 山=2-S,"",2-E,"",1-E、1 山=S,"",2,"",E。InspeButt が "検入≧26" で表示中なら 3 回検査用に罫線を変える（skipEnd へ）。3 山 / 1・2 山 / それ以外で罫線を分ける |
| スラブ | 1,"",2,"","" / 1上,1下,2上,2下,"" | 厚板と同じ | J18:L23 を消し、罫線を整理 |
| 極厚（耳付） | 原紙のまま | 原紙のまま（幅1〜6、丈1〜4） | I 列の左に中太の罫線 |

**ラテラルボー / 直角度の列**（:1041-1123）:
- 表示する条件:
  - `(Sumisei<>"" かつ 「ﾗﾃﾗﾙﾎﾞｰ検査不要」でない) または FlagColor（LateralText の背景が RGB(51,204,204) かつ BackStyle=1。薄板 4 山の場合のみ評価） または LateralText.Tag="有"`
  - かつ、極厚でもスラブでもない
- 表示する場合の配置:
  - R115: P 列（ラテラル）→ K 列、R 列（直角度）→ J 列
  - それ以外: P 列 → J 列
  - 次に、`Lateralbow_Check`（OutTolerance:2090。表示中のラテラル欄に未入力があればその位置名を、無ければ "" を返す）が "" なら J18:N24 を消す（:1053-1057。コメント上の意図は「測定不要のときに罫線を消す」）。印刷前チェックで未入力は弾かれているので、**印刷時は常に "" になり、消去が実行される**。
  - その後、R115 以外なら P → J に再びコピーして K:N を消す（:1059-1063）。
  - **結果として R115 では、先にコピーしたラテラル/直角度の列が消える**（コードを順に追った結果。要確認。§10-21）
- 表示しない場合: J18:N23 を消す（極厚は N18:P23）。
  - ただし LateralText の色フラグがあり、R115 以外・極厚以外・スラブ以外なら、P → J にコピーし直す（:1104-1115）。
  - 「ﾗﾃﾗﾙﾎﾞｰ検査不要」かつ薄板なら J:N を消す（:1118-1122）。

#### 3.2.5 公差外データ（:640-717）

- 条件: シート `異常ロット用_<Lname>` があり、かつ その A1 = 成績票 C2（LOT）。
- 処理（`図` シートがあれば、その図形を L11 に貼る）:
  - A26="公差外ﾃﾞｰﾀ"。2024.12.24 の改修で開始行を 26 に固定（:663）。
  - 極厚: 見出し B26.. = 厚1-4, 幅1-6, 丈1-4（14 列 B..O）。データは異常シートの 2 行目〜n 行目、列 1-14 → B27。
  - それ以外: 見出し = 厚1-4, 幅1, 幅2, 丈1, 丈2, 対角A, 対角B, 直角度, ﾗﾃﾗﾙﾎﾞｰ, 定盤, ｹﾞ-ｼﾞ（B..O）。データは列 1-10 → B27、列 15-16 → L27、列 17-18 → N27。※コメントには「ﾗﾃﾗﾙﾎﾞｰ 直角度の順」とあり、見出しの順と食い違う **（要確認）**。
  - 位置（列 19）→ A27。
  - 表全体: フォント 8pt、左寄せ、上下中央。
  - 数値書式: 厚が 74 を超える → A..E 列="0.00_);[赤](0.00)"、74 以下 → "0.000_);[赤](0.00)"。F..K 列は "0.0_);[赤](0.0)"。H494 で幅<300 のときは下記の特例。
  - ページ設定は `印刷範囲`。
- シートはあるが LOT が違うとき: **比較しているのは D4（厚）と 異常シート A1**（:712、C2 ではない）。実際には常に「違う」になり、`印刷範囲2` が呼ばれる。万一一致すると、**ページ設定が呼ばれない**。
- シートが無いとき → `印刷範囲2`。

#### 3.2.6 書式と文字サイズ

- G12:J16（厚。**17 行目は対象外**）: 厚 D4 > 74 → `"0.00_);[赤](0.00)"`、それ以外 → `"0.000_);[赤](0.00)"`（:625-629）。
- 用途 H494 かつ `UF1.Hab < 300`: Val(幅) が 1〜299 → C19:D23 を `"0.00_);[赤](0.00)"`、それ以外 → `"0.00_);[赤](0.0)"`（:632-639。品証規定 Q-S-QA-MSR-007「H494 は幅をノギスで測る」）。
- 外観 C12:C17 の `FitTextToCell`（:1412-1460）:
  - 8pt から 0.5pt ずつ下げ、一時テキストボックスの BoundWidth/Height がセルに収まるまで縮める。
  - 結果が **4pt 以上なら 8pt に戻す**。ちょうど 3.5pt なら 6pt。それ未満はその値（下限 1pt）。
  - つまり「8pt のままか、極端に長い場合だけ 6pt 以下」。Python では同じ計測ができないため、**同等の判定方法を決める必要あり（要確認）**。
- B9 は `ShrinkToFit=True`。I3 は 6pt。

#### 3.2.7 ページ設定（ヘッダー・フッター）

| 項目 | 印刷範囲（公差外あり） | 印刷範囲2（公差外なし） |
|---|---|---|
| PrintArea | 極厚: `$A$1:$P$(n+1)`、他: `$A$1:$O$(n+1)`（n=A 列の最終行） | `$B$1:$L$(n+1)`（n=B 列の最終行。通常 24〜25）。※極厚でも B:L なので、N 列（差）などは印刷されない |
| 用紙・向き | A4 横 | A4 横 |
| 余白 | 左右下・ヘッダー・フッター 0、上 0.4inch | 同左 |
| 拡大縮小 | FitToPagesWide/Tall=1。ただし Zoom は設定しておらず、既定の 100 のままなので、**Fit は効かない可能性あり（要実機確認）** | FitToPages=1 の後に **Zoom=115**（Excel の仕様上、Zoom が数値なら Fit は無視される）。水平中央 |
| 左ヘッダー | `&"Meiryo UI"&8` + Lname | 同左 |
| 右ヘッダー | `&"Meiryo UI"&8` + C2(LOT) + " " + 検査員 | 同左 |
| 右フッター | `&"Meiryo UI"&8` + " " + `Format(StartTime,"yyyy/mm/dd  h:mm")` + "～" + `Format(EndTime,"yyyy/mm/dd  h:mm")` + " 印刷"（日付と時刻の間は**半角スペース 2 個**） | 同左 |

### 3.3 不良要因票 `不良要因印刷_<Lname>`

原紙（`不良要因印刷原紙`）:
- 1 行目（区分）: C1:D1=鋳造、E1:H1=熱延、I1:M1=冷延、N1:V1=仕上（結合）。
- 2 行目（不良名、C..X）: フクレ, (空), ムシレ, 汚, リックアップ, (空), 圧込, 汚, フラットネス, 厚, (空), 汚, 飛込, 当り, 擢, 付着油, 切断面, 巾丈, 特寸, (全角空白), 他, TP。
- 9 行目（不良コード、C..W）: 01,02,11,12,13,14,21,22,23,24,25,31,32,33,34,35,36,37,38,39,41。X9 は空（TP 用）。
- B10="不良合計"、B11="スポット"。
- 列幅: A=3, B=10.25, C〜W=5.875, X=6, Y=3.375。

入力（`SelectionInput` OutTolerance:2356-2436。担当外だが帳票の中身に関わるので記載）:
- B1=LOT。B3..B8 = `'` & UF1.S/U/C/D/E/F.Caption（検査位置名）。
- 選択した不良「名前_コード」のコードと 9 行目が一致する列の、位置行（S=3 … F=8）に数量を入れる。
- 「TP」は 22 列目（X 列）。「ｽﾎﾟｯﾄ処理」は C11 = 数量 & "枚"。
- 一致する列が無ければ、シートを削除して作り直し、やり直す。
- UF1 の「ｸﾘｱ」ボタン（`Cle_Click` UF1:567）で C3:X8、C10:X10、C11 を消す。

印刷前の処理（`不良要因印刷範囲` Print.bas:1278-1335）:
- PrintArea `$A$1:$Y$(n+1)`（n=B 列の最終行。通常 11 → A1:Y12）。
- A4 横。余白はすべて 0（上も 0）。Fit 1×1 の後に **Zoom=100**。水平・垂直とも中央。
- ヘッダー: 左 = Lname、右 = LOT & " " & 検査員（どちらも Meiryo UI 8pt）。
- **右フッター（UF1 経由のときだけ）**: `StartA ～ EndT 印刷`。
  - `StartA` = UF1.StartTime（UFMitsubishi が表示中なら `全ヵ所検査!B1`、UFCask なら `キャスク検査!B1`）。
  - `EndT` = 現在時刻。ただし時刻が 0:00〜`End3`（3 直の終了。ShiftTime DB 由来）なら**前日の日付**にする（:1296-1305）。
- `CalculateColumnSums`: C..X 列、行 3〜8 の数値を合計して行 10 に書く（Single）。
- `ConvertNumbersToString`:
  - C3:AA8 と C10:AA10 の数値を `CircledNumber` に置き換える（1〜20 → ①〜⑳ U+2460〜、21〜35 → ㉑〜㉟ U+3251〜、36〜50 → ㊱〜㊿ U+32B1〜、範囲外はそのまま）。値は Long に代入する（小数は銀行丸め）。
  - **数値でないセルは、配列を書き戻すと空になる**（ConvArr に入らないため）。ただし書き戻すのは、数値が 1 つ以上あったときだけ。
  - 行 10 の 0 は ""。
  - 50 を超える数は数字のまま（:1357 のコメント）。

### 3.4 時刻の扱い（成績票のフッター）

- `EndTime_Register`（`Worker.bas:312-325`）:
  - 現在時刻が 7:00 より後かつ 23:59 以下 → EndTime = 今。
  - 0:00〜7:00 → **前日の日付で、時刻は現在時刻**（夜勤は前日扱い）。
  - EndTime は `WorkerTime_<Lname>!C1` にも書く。
  - 7:00 は固定値。不良要因票の方は `Start1/End3`（DB の直時間）を使っており、**判定方法が統一されていない**。
- StartTime は UF1 側で決まる（「開始時間(S1測定時間)」。担当外）。

### 3.5 ラベル・バーコード

- **ラベルを出力するコードは無い**（PrintOut は §3.1 の帳票だけ。プリンタ制御コード・ZPL・ラベルソフト連携も見つからない）。
- **バーコード**は入力にだけ使う:
  - SF3 で 10 桁の設備名バーコードを読む → `伝送用ファイル.accdb` のテーブル `設備名バーコード` を `[No]='<コード>'` で引き、`名前` を得る（`SF3.frm.bas:52-60, 88-143`）。
  - 検査員バーコードは NameConv（`Worker.bas:37`、担当外）。
- 帳票の「バーコード」欄も無い。

### 3.6 Main シート（Main.csv、使用行 1〜482）

| 範囲 | 内容 |
|---|---|
| A1 | "CLOSE"（固定値。BeforeClose で閉じてよいかの判定に使う。**この値を書き換えるコードは無い**） |
| A2 | ライン名（"Test"）。SF3 で設定する。ThisWorkbook がログの出力先フォルダを選ぶのに使う |
| G2 | "H3082Z0"（LOT 番号らしい文字列。参照するコードなし。メモと推測） |
| B8〜B46 | 開発者メモ: フリー反りの条件コード（B9、B10「2026.2.6 フリー反り測定基準が発行されたら条件停止」）、Text ファイルの文字コード（UTF-8 は文字化けする、ANSI なら読める。B37-41）、**許容/保留の業務定義**（B44「許容 全スルー 許容フラグが立った後は別の検査でNGが出ても処理を聞かれない、MinMax、印刷前チェックもスルー」、B45「保留 別の検査でNGが出た場合再度処理を聞く、MinMaxでSTOP」、B46「印刷前チェックでNGチェックをなくす」） |
| I 列 / J 列 / L 列 | **改修履歴**（I=日付 2020.12.16〜2026.10.14、J=問題、L=対応内容とモジュール名。最終行付近 L431-482 にはコードの断片）。印刷関連の主な項目: 2020.12.18 検査員を履歴から全員拾う / 2023.5.12 H494 桁数 / 2023.5.29 囲み数字 / 2023.7.11 許容公差外を印刷 / 2025.01.17 不良要因も PDF / 2025.01.29 外観の文字サイズ / **2025.12.15 再印刷機能（理由「保留処理で印刷用紙がいるため」）** / 2026.02.25 薄板のラテラル罫線 / 2026.03.05 フリー反り・公差の印刷レイアウト |
| K259 | "AccReflec_フリー反り以外の場合もCBAdd実行" |
| ActiveX | CommandButton1「ﾒｲﾝﾌｫｰﾑ起動」（B1:E7 付近）、CommandButton2「CommandButton2」（E23:G31 付近。`Form_1` が存在しないので押すとエラー **（推測）**）。図形: 注意書き「本作業標準にない操作は予期しない処理が行われる場合があり…」、マニュアルへのハイパーリンク `\\nlmfangyshrd\各課共有\0130_日軽稲沢\梱包課\0-2-22_検査判定システム操作.pdf`、EX/ストア梱包時のアイカス入力の注意書き |
| ページ設定（保存値） | A4 横、Zoom 115、上余白 0.4、ヘッダー `&RHX370D0`、フッター `&R2020.12.15 19:26:15～2020/12/16  9:41 印刷`（過去に成績票の設定を試した名残 **（推測）**） |

---

## 4. Web / IE / WebDriver

| 項目 | 内容 |
|---|---|
| 対象システム | 社内の包装仕様書 Web「PkgSpecWeb」の**現場閲覧画面** |
| URL（本番） | `http://nlmfangyweb1a:9084/PkgSpecWeb/#/hososhiyosho/genba/etsuran`（`IE.bas:24,60`） |
| URL（開発・コメント） | `http://nlmfangydev2:9083/PkgSpecWeb/#/hososhiyosho/genba/etsuran` |
| 操作 | 検索欄（id/name=`pkgSpecNo`）に包装仕様書 No を入れる → IE 版は `change` イベントを発火（2022.03.11 追加） → class `btn btn-info c-btn-size` の検索ボタンをクリック。表示するだけで、値の取得や書き込みは無い |
| 呼び出し条件 | UF_PackNo の「ﾌﾞﾗｳｻﾞで確認」ボタン（`UF_PackNo.CommandButton1`）。設計時から **Visible=False** で、表示に切り替えるコードは見つからない → **実行されない（推測）** |
| Edge/WebDriver | driver: `\\nlmfangyshrd\各課共有\0130_日軽稲沢\梱包課\AIM\検査\詰め合わせパック\msedgedriver_97.0.1072.69.exe`、`http://localhost:9515`。呼び出しはコメントアウト（「NIS都合 edge使用不可**2022.02末」） |
| 後片付け | UF_PackNo「閉じる」で `ie_kill` → IE の窓があれば `taskkill.exe /F /IM iexplore.exe`（**利用者が開いた IE も強制終了する**） |
| TEST フォーム | google.co.jp を IE で開くテストだけ |

移植方針の候補: 包装仕様書 No を URL のリンクとして表示するだけで足りる可能性が高い。ただし SPA（`#/...`）にクエリで値を渡せるかは不明 **（要確認）**。

---

## 5. 外部依存

### 5.1 ネットワーク・ローカルのパス（`初期設定.bas:18-48`、レジストリのモードで切り替え）

| 変数 | オンライン | オフライン |
|---|---|---|
| MainPath | `\\nlmsrvngy03\Read\【New】仕掛\台帳\` | ThisWorkbook.Path & "\" |
| NextPath | `\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\仕掛\` | 同上 |
| QualityPath | `\\nlmsrvngy03\工場内共有\検査データ\品証-梱包やり取りフォルダ\Accdb\`（包装仕様書の JPG は `QualityPath & "包装仕様書\"`） | 同上 |
| PackingPath | `\\nlmfangyshrd\各課共有\0130_日軽稲沢\梱包課\AIM\【■】_参照用ファイル\` | 同上 |
| HistoryPath | `\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\検査履歴\`（ログ・CSV） | 同上 |
| FactoryyPath | `\\nlmsrvngy03\工場内共有\検査データ\`（工場共有の PDF/CSV。`【梱包】L-1/LVC/HVC/大板/中板`） | 同上 |
| NisPath | `C:\Tools\測定データ電子化` | **未設定（空）** |

- 設定するタイミングは `Workbook_Open`、`Workbook_BeforeClose`、`初期設定値確認`（Main ボタン）。**初回起動でレジストリが無いとパスは空のまま**になり、SF3 で設定した後もブックを開き直すか、Main ボタンを押し直すまで空 **（推測）**。`End` が実行されると Public 変数が消えて空になる。
- その他のハードコードされたパス:
  - `C:\検査履歴`（PDF のローカル保存先。Input:1029）
  - `C:\VBAWAV\`（音声）
  - `\\nlmfangyshrd\…\詰め合わせパック\`（WebDriver）
  - `C:\Users\xfa-ngyinazawa\Desktop\DB\品質\`（TEST だけ）
  - `ThisWorkbook.Path\DebugPrint\…`
  - `ThisWorkbook.Path\Access.log`

### 5.2 DB（Access .accdb。ADO 経由の `GetRecordsArr`/`GetFieldsArr` は SQL_追加 にある＝担当外）

- `伝送用ファイル.accdb`（QualityPath → 無ければ PackingPath）:
  - テーブル `設備名バーコード`（No→名前）: SF3
  - テーブル `PkgSpecData`（包装仕様書№、梱包形態、梱包指定別、タイトル、用途名）: PackSelect（UF_PackNo の L1〜L5 表示）
  - テーブル `名簿`: TEST/NameConv
- `SIKAHIKINOW.accdb` / `SIKAODRNOW.accdb`（テーブル「仕掛」）: ZZ（Test シート）と TEST。
- DbReturnMicro（機器名。担当外）、ShiftTime（直時間。担当外）。

### 5.3 音声（`C:\VBAWAV\`、`mciSendString "Play <file>"`。ファイルが無ければ何もしない）

印刷の流れで使うもの: `おまちください.wav`、`印刷.wav`（「印刷する？」・再印刷の前）、`完了.wav`、`ありますか.wav`（先行データ票。`音楽を流す` は `UF1.Tag="STOP"` なら鳴らさない）。SF3 で `chime1.wav`。その他（担当外）: 必要です / 確認作業 / ナンバーNO / 直角度 / 異常値 / 歪の公差 / お待ちください / tada / chimes / chime2.wav。clsBlinkLabel は Buzzer 指定時に `%WINDIR%\Media\Ding/Chord/Notify.wav` を鳴らすか Beep するが、呼び出し側は Buzzer を指定していない。

### 5.4 レジストリ

`HKCU\Software\VB and VBA Program Settings\UF検査判定システム\起動設定\`
- `設備名`: L-1 / LVC / HVC / バランサー / トット / Test / 丸徳。SF3 は "AIM" も受け付ける。
- `モード`: オンライン / オフライン。

保守用のマクロ: `AIM()`（設備名=Test、モード=オンライン）、`HVC()`、`DelSet()`（両方削除）、`test取得()`。

### 5.5 Shell・外部プロセス・COM

- `taskkill.exe /F /IM iexplore.exe`（ie_kill, TEST）
- msedgedriver の起動（WebDriver:208, IE:46）
- `InternetExplorerMedium` / `Shell.Application` / `WScript.Shell`
- `MSXML2.ServerXMLHTTP`
- `Scripting.FileSystemObject` / `Scripting.Dictionary`
- GDI/USER32 API（ListView の文字幅、WindowMenuButtonDel の FindWindow/SetWindowLong）
- `kernel32.Sleep`
- `winmm` の PlaySound / mciSendString
- MSComctl の ListView / ProgressBar（ActiveX）

### 5.6 ログファイル（`WriteLog` `【UFKen_16】WriteLog.bas:6-40`）

- 書式: `Format(Now,"YYYY/MM/DD hh:mm:ss") & " " & Sepl & COMPUTERNAME & " " & Sepl & " " & msg & " "`（Sepl は全呼び出しで ""）。
- 新しくファイルを作ったときは、1 行目に `管理ﾌｧｲﾙ使用状況` を書く。
- ファイル名は既定 `Access.log`（`Measured.log` も使う）。パス未指定なら `ThisWorkbook.Path`。フォルダは `CreateFolder` で作る。
- `On Error Resume Next` なので、失敗しても何も起きない。
- 出力先: `HistoryPath<ライン>\<ライン>_Log\Access.log`。

担当範囲のモジュールが出すメッセージ:

| メッセージ | 出す場所 |
|---|---|
| `*UFKen:印刷 T1:<T1> T2:<T2> T3:<T3>` | UF1:3755 |
| `*UFKen:MsgError:<文> T1:.. T2:.. T3:..` | MsgError |
| `UFConfi起動 T1:..` | UFConfi |
| `<位置><番号>_<元値>:理由選択:<理由>:<検査員>` | UFConfi |
| `公差外:<位置><番号>_<元値>:理由選択:再測定,<理由>:<検査員>` | UFConfi |
| `*UFConfi.TB1入力:<文字>` | UFConfi |
| `*UFKen:DragOver:座標_<名> Data_<値>` / `*UFKen:DropOrPaste:移動先_<名>:元data_<旧> 後data_<新>` | LabelCls |
| `*UFKen:[A/B] CBS1:<値>:<ライン>` | LabelCls（Measured.log とブックフォルダの Access.log） |
| `*<ライン>_*<ブック名>_Open/_Close` / `<ライン>_ﾌﾞｯｸ閉` | ThisWorkbook |

ライン名とログフォルダの対応:
- UF1 系: Lname "L-1"/"LVC"/"HVC"/"ﾊﾞﾗﾝｻｰ"/"ﾄｯﾄ"/"Test"、それ以外は**丸徳**フォルダ。
- ThisWorkbook 系: Main!A2 が全角の "バランサー"/"トット"。

---

## 6. ハードコードされた業務ルール・設定値（担当範囲）

1. 印刷履歴の保持期間: `DateDiff("d") > 3`（4 日目以降に削除。Print:176）。削除するのは UF1 起動時だけ。
2. 印刷済みの判定キー: LOT 番号だけ（ライン横断）。
3. ボタンの色: 印刷済み RGB(255,165,0)「再印刷」、未印刷 RGB(192,192,192)「印刷」。
4. 厚さの表示桁: 厚 > 74 → 小数 2 桁、≤ 74 → 小数 3 桁（成績票 G12:J16、公差外表の A..E）。公差外表の F..K は小数 1 桁。
5. H494 かつ幅 < 300: 幅（C19:D23 と公差外表）を小数 2 桁（Q-S-QA-MSR-007）。
6. 対角公差・フリー反り公差は `"0.0"` 形式。対角公差が空でも "0.0" を印字する。
7. 外観の記号: 欠陥あり → "△"、ちょうど「欠陥ﾅｼ」 → "〇"。A/B 面は "A:"、"B:" を前に付けて改行で区切る。
8. 歪の A/B: "A:"、"B:"。全部空は "__"。"_" は削除。フリー反りのときは A:/B: を外す。
9. 検入: KenHBox7 が数値なら「枚数表示（<n>枚）」、そうでなければ合計高さと「n山」。
10. 直角度は用途コード **R115** のときだけ。ラテラルボーの表示条件は §3.2.4（「ﾗﾃﾗﾙﾎﾞｰ検査不要」の文字列、LateralText の色 RGB(51,204,204)、Tag "有"、Sumisei）。極厚とスラブでは表示しない（2024.11.15）。
11. 検査形態ごとの位置名（S/S1/2/E1/E、薄板の山数別 4-S…、スラブ 1上/1下…）。§3.2.4 のとおり。
12. 公差外表の開始行は 26 で固定。
13. 不良要因: コード表（01〜41, TP）、囲み数字は 1〜50、スポットは「n枚」。
14. 夜勤の日付またぎ: 成績票の終了時刻は 7:00 が境界（固定）、不良要因票は `End3`（DB）が境界。
15. ヘッダー/フッターのフォントは Meiryo UI 8pt。日付書式は `yyyy/mm/dd  h:mm`。
16. 再測定の理由入力: 同じ座標を 2 回以上消したら、理由リスト（入力位置間違い/測定位置間違い/測定ボタン押し間違い/0点確認/その他）を表示（UFConfi:239-253）。値を消したときの理由（List1）: 入力位置間違い/測定位置間違い/測定ボタン押し間違い/再測定/0点確認/板廃棄/その他（`OutTolerance:1616-1625`）。「再測定」「その他」は自由記述が必須。
17. 板廃棄: 消した値を戻して判定し直し → NG_Posting → NG_Erase（2024.2.5。150kg の条件は 2024.4.18 に廃止）。
18. UFConfi の自動入力ボタン（公差外/許容/再測定/保留）は HVC と Test のときだけ表示する。
19. 機器名（機器転送。担当外の概要）: コンベックスは L-1 なら 丈>5500 で C11、それ以外 C12。HVC=C11、LVC=C10、バランサー=大板 C12、トット=大板 C11。丸徳/Test は空。厚=300 のときの N5/N6 判定もある。
20. 包装仕様書の画像: `<QualityPath>包装仕様書\<No>.JPG`。UF_PackNo の表示は最大 5 行（引当No 順）。
21. 有効な設備名: L-1/LVC/HVC/バランサー/トット/AIM/Test/丸徳（SF3:27-34）。設備名バーコードは 10 桁。
22. ListView の列幅: 6〜30 文字 × 7.95。12 文字を超えたら左寄せ。

---

## 7. エラー処理と共通メッセージ

- **Print モジュールには `On Error` が無い**。シートが無い（WorkerTime_、印刷用_、不良要因印刷_）、`Start1/End3` が空で `CDate("")` になる、などのときは VBA の実行時エラーダイアログで止まる。
- 業務上のエラーは `MsgError(Trm, Lname, n, Temp)`（OutTolerance:716-786）を使う。
  - `UFMsgBox` に表示する。アイコンは n=1「!」IM1、2「×」IM2、3「?」IM3、4「i」IM4（Caption は 4 なら "注意"、それ以外は "警告"）。
  - Trm を "," で区切って改行する。文字数からフォーム幅を計算する（`LenB(StrConv(..,vbFromUnicode))*8+120`）。
  - **モードレスで表示 → ログを書く → UF1.Hide**。
  - OK か × で閉じると: Temp="END" なら Excel を最大化・表示して **`End`**（VBA 全体の変数リセット）。"Ainput" ならそのまま。それ以外は UF1.Show。
  - 印刷側は MsgError の後に `PrintB.Tag="stop"` にして処理を抜ける。
- 通常の `MsgBox` もよく使う（「印刷する？」「ｷｬﾝｾﾙ処理」「再印刷が完了しました」など）。
- `Msg.Box`（自作 MsgBox）は使われていない。
- `On Error Resume Next` を使っている箇所: WriteLog、DebugLog、IE.bas 全体、ListView.CreateList、clsBlinkLabel.Blink、UFMsgBox の Activate/QueryClose、UF_PackNo の Activate と CommandButton2、ControlInfo.InitFromControl、TabletTest の Resize。
  - **IE.bas:79 の `Do While objIE.Busy`**: objIE が取れないと、エラーで次の文へ進み続けて**無限ループ**になる可能性がある **（推測）**。

---

## 8. 状態保持・待ち・放置時の挙動

| 仕組み | 場所 | 挙動 |
|---|---|---|
| `Application.OnTime` / API タイマー | — | **プロジェクト全体で使われていない**（grep で確認） |
| 点滅 `clsBlinkLabel.Blink` | UFProgress.Activate (Blink 2)、UFAinput:1495 (Blink 3)、UF1（呼び出しはコメントアウト） | `Sleep 500` → DoEvents → `Sleep 500` を BlinkCount 回繰り返す（2 回で約 2 秒）。その間 UI は DoEvents でしか動かない。クリックで止まる（mctlLabel_Click / UFProgress のフォームクリック） |
| 進捗 `ProgressBar_count` | 印刷フロー（0→250→375→550→650→880→1000） | ループの中で DoEvents。時間待ちではない |
| `Sleep` | IE（3000/1000ms）、初期設定（500ms、SF3 表示前） | ブロックする |
| DoEvents ループ（タイムアウトなし） | IE.WaitIE / Call_IE_WaitTime / IeSiteSetTextData:79 | IE が応答しないと永久に待つ。ieCheck は 20 秒ごとに Refresh を繰り返す（終わらない可能性あり） |
| `Application.Wait` | UFNG/UFOK/UFcount/UF対角（担当外） | 表示時間を待つ |
| モジュール変数 | `PrintedLots`（Print）、`g_logFilePath`（DebugOut）、`Static fso`（WriteLog）、パス 7 個（初期設定）、`Driver As New WebDriver`（IE） | `End`（MsgError "END"、UF1.Initialize で設備名なし）や VBA のリセットで消える。PrintedLots は IsPrinted のたびにシートから読み直すので影響は小さい |
| フォームの状態 | UF1 の Label.Caption/Tag/BackColor、`Load.Caption`（"on"/"off"）、`PrintB.Tag`（"stop"）、`dataLoding.Caption`（"Load"）、`Loading.Tag`（"Loading"）、`DeleteC/DeleteP`、`DBflag` | 状態はすべてコントロールのプロパティにある。Web 化ではサーバー側のセッション状態として持つ必要がある |

長時間放置したときの挙動:
- **印刷履歴の古い行は UF1 を再起動しないと消えない**。UF1 を何日も開いたままだと、4 日以上前の LOT でも「再印刷」扱いのまま **（推測）**。
- DebugLog のファイルは Excel セッション中ずっと同じファイルに追記され、大きくなる。WriteLog の各ログは無期限に追記される（ローテーションなし）。
- 開始時刻から終了時刻（印刷時）までの差がそのままフッターに出る。夜勤をまたいで放置すると、日付補正ルール（§3.4）に従った表示になる。
- `PrintB` の 2 度押し対策（`Enabled=False`/`Load="off"`）は途中で `Exit Sub` するとき戻しているが、**実行時エラーで落ちた場合は戻らず、ボタンが無効のまま残る（推測）**。
- 印刷後は `Application.Visible=False` になり、Excel 本体は見えない（UF1 だけ表示）。

---

## 9. 未使用コードの候補（断定はしない）

| 候補 | 根拠 |
|---|---|
| `ThisWorkbook1.cls`, `Main1.cls` | PROJECT ストリーム上は `Class=`（Document ではない）。インスタンスを作るコードが無いので、Workbook_Open 等は発生しない。Main1 は存在しない `UserForm1` を参照 |
| `TabletTest.bas` 全体と `ControlInfo.cls` | 外部からの呼び出しが 0（grep で確認） |
| `StringBuilder.cls` | `New StringBuilder` が 0 件 |
| `Msg` フォーム | `Msg.Box` の呼び出しがすべてコメントアウト（SQL_追加:202,298,645） |
| `Logic.cls`（シート ﾛｼﾞ） | 中身が空 |
| `WebDriver`/`WebElement`/`Capabilities`/`JsonConverter` | EdgeSiteSetTextData の呼び出しがコメントアウト（UF_PackNo:23）。他に使っている箇所なし |
| `IE.bas` の IeSiteSetTextData | 呼び出し元のボタン「ﾌﾞﾗｳｻﾞで確認」が設計時から非表示で、表示にするコードも無い。TextDataLoad/ieCheck/WaitIE/Call_IE_WaitTime は呼び出しなし（TextDataLoad はそもそも未定義変数を参照）。**ie_kill だけは UF_PackNo の「閉じる」で毎回呼ばれる** |
| `TEST` フォーム・`ZZ`（Test シート） | 開発者用。隠しシート Test のボタンからしか開けない。ZZ.CommandButton3 は存在しない `Height_Checker` を参照 |
| `Main.CommandButton2_Click`（Form_1） | `Form_1` がプロジェクトに無い。ボタンの Caption も既定値の "CommandButton2" |
| Print: `Call_BordersLineStyleWeightCopyPaste` | 呼び出しなし |
| Print: `Pesute` の "作業中保存" ケース | シート名が接尾辞なしの "作業中保存" で呼ばれる箇所があるかは担当外のため未確認（"ｷｬｽｸ作業中保存"、"全ヵ所作業中保存"、"H494作業中保存" は Cask/Mitsubishi が使う **（推測）**） |
| ListView: `LvField/LvRec/CreateAll/LvHide/ViewSelNumArr/GetLvData/GetList/ListPatch/GetWidth/IsHalf` | 外部からの呼び出しなし。CreateAll の呼び出し元 adoRead/adoReadPageCtrl も呼び出しなし |
| DebugOut: `CloseDebugLog`/`ResetDebugLog` | 呼び出しなし |
| UFProgress: `IsCancel` | 参照なし |
| LabelCls の DragDrop 系 | `UF1.DBflag`（「日稲鯖優先ﾌﾗｸﾞ」テキストボックス）が空なら動かない。コメントは「品証許可出ず」。**通常運用では無効（推測）** |
| UF1.TestButton（Caption "Test"、Visible=False） | 印刷転送を単体で呼ぶテスト用 |
| `初期設定` の AIM/HVC/test取得/DelSet | 保守用（TEST フォームのボタンからだけ使う） |

---

## 10. 不明点・要確認事項（優先度順）

1. **【帳票の中身】B10「許容公差外在り/公差外 保留処理」が出ない不具合**: `allow_Check`（OutTolerance:2611-2623）は、該当が 1 か所だけだと「S」のような位置名だけを返し、「許容」「保留」の文字が入らない。そのため `印刷転送:612-620` の InStr 判定に引っかからず、B10 が空になる。2 か所以上だと、2 つ目以降にだけ「 許容」が付く（例: `S, U 許容` → 置換後 `B10="許容公差外在り:S, U "`）。**現行のこの挙動を再現するか、意図どおりに直すか**の判断が必要。
2. **【帳票の中身】再印刷が別 LOT の内容を出すおそれ**: PrintOnly は `印刷用_<Lname>` と `不良要因印刷_<Lname>` をそのまま印刷する。LOT A を印刷 → LOT B を印刷 → A を読み込み直して「再印刷」とすると、**B の内容に「再印刷」と書いたものが出る**可能性がある（シートはライン単位で 1 枚しかないため）。運用上起きるか、確認が必要。
3. **【帳票の中身】不良要因の合計が消える**: 「印刷する？」で No を選んだ後にもう一度印刷すると、`印刷転送` が再び `不良要因印刷範囲` を呼ぶ。囲み数字は数値でないため合計が 0 になり、行 10 が空になる。また、その間に追加した数値だけが合計される。
4. **印刷倍率**: 成績票（公差外なし）は Zoom=115、不良要因は Zoom=100、成績票（公差外あり）は Fit 指定だけで Zoom 未指定。Excel の仕様（Zoom が数値なら Fit は無視）どおり動いているか、実際の印刷物で確認が必要。Python 版の出力（PDF/HTML 印刷）の縮尺の基準にする。
5. **図形のコピー**: `Pesute` で原紙の測定位置図（L11）がコピーされるか。公差外のときに `図` シートの図形を同じ L11 に重ねて貼っているので、どちらが見えるのかを**実際の印刷物で確認**する必要がある。
6. L4:L6 の下罫線を `HizKousa(4)`（耳歪 Mimi）で判定している。フリー反り（HizKousa(6)）の誤りではないか（:475）。
7. 公差外表の列順: 見出しは「直角度, ﾗﾃﾗﾙﾎﾞｰ」だが、コメントは「ﾗﾃﾗﾙﾎﾞｰ 直角度の順」（:677-680）。`異常ロット用_` シートの列 15/16 の意味を担当側（OutTolerance.NG_Posting）で確認する。
8. 公差外の判定 `sh.Cells(4,4) <> IzSh.Cells(1,1)`（:712）は、C2（LOT）と比べるべきところを D4（厚）と比べている。現状は実害なしの不具合 **（推測）**。
9. `FitTextToCell` の判定（4pt 以上なら 8pt に戻す）を、Python でどう再現するか（文字幅を計測する手段がない）。
10. `印刷履歴管理` の B2 が文字列 "2026.06.22" になっている。`IsDate("2026.06.22")` が日本語ロケールの VBA で True になるかによって、削除の挙動が変わる。また、PrintOnly/ClearPrinted/CleanOldHistory の後は `ThisWorkbook.Save` されないので、Excel を保存せずに閉じると変更が消える。
11. 終了時刻の日付またぎの境界が、成績票（7:00 固定）と不良要因票（`End3`）で違う。どちらに統一するか。
12. `WorkerTime_<Lname>` シートが無いと実行時エラーになる（:558-561）。これを前提にしてよいか。
13. ライン名の表記ゆれ: UF1 の Lname（半角「ﾊﾞﾗﾝｻｰ」「ﾄｯﾄ」）と Main!A2（全角「バランサー」「トット」）。丸徳と AIM は Case Else 扱いになる。印刷ヘッダーの文字列は Lname（`LineName.Caption & LineNo.Caption`）。実際の値は UF1 側の担当で確認する。
14. プリンタの指定: 現状は Windows の既定プリンタ（EPSON LP-S180DN の想定）。Web 化した後、ブラウザ印刷にするか、サーバーから印刷するか。
15. `C:\検査履歴` と `FactoryyPath【梱包】<LN>\yyyy.mm\dd日\` への PDF 保存（Input:1029-1058）は、帳票の出力の一部として移植が必要（担当 G?）。
16. IE/WebDriver（包装仕様書 Web）は本当に不要か（現場がボタンを使っていないか）。`ie_kill` が利用者の IE を巻き込んで終了させる挙動は再現しない予定でよいか。
17. Main シートの CommandButton2（Form_1）と、Test シートの「高さチェック」（Height_Checker）は壊れたボタン。削除してよいか。
18. `DebugLog` はデバッグ目的なのに、本番環境でブックのフォルダにファイルを作っている。移植では通常のログに統合してよいか。
19. SF3 の TextBox1 に設計時の値 'T716980265' が入っている（参照するコードなし。用途不明）。
20. 機器転送の 厚=300 のときの N5/N6 判定は、成績票の B14:H18 付近（現在のレイアウトでは外観や歪の欄）を数値チェックしており、旧レイアウトの名残の可能性がある（担当外。要確認）。
21. **R115 のラテラルボー/直角度の列が成績票から消える疑い**（§3.2.4）: `印刷用罫線:1046-1057` の順に処理すると、R115 用にコピーした J/K 列を、直後の `ラテ=""` の処理が消す。過去の R115 の印刷物で、この 2 列が出ているかを確認する必要がある。
