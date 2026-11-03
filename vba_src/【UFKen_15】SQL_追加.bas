Attribute VB_Name = "【UFKen_15】SQL_追加"
Option Explicit
Option Private Module 'マクロメニューから隠す
Private RecFlag As Boolean 'ﾚｺｰﾄﾞの有無
Private Const adClipString As Long = 2
'
Private Const adOpenForwardOnly As Long = 0     '前方スクロール専用カーソル（既定値）
Private Const adOpenKeyset As Long = 1          'キーセットカーソル（他ユーザーの追加・削除レコードへのアクセス不可）
Private Const adOpenDynamic As Long = 2         '動的カーソル（他ユーザーによる追加・更新・削除を反映）
Private Const adOpenStatic As Long = 3          '静的カーソル（他ユーザーによる追加・更新・削除は反映されない）
Private Const adOpenUnspecified As Long = -1     'カーソルタイプ非指定
'
Private Const adLockReadOnly As Long = 1        '読み取り専用
Private Const adLockPessimistic As Long = 2     'レコード単位の排他的ロック
Private Const adLockOptimistic As Long = 3      'レコード単位の共有的ロック(Updateメソッドでのみロックする)
Private Const adLockBatchOptimistic As Long = 4 '共有的バッチ更新
Private Const adLockUnspecified As Long = -1     'ロックタイプ非指定
'
Private Const adCmdText As Long = 1             'SQL 文
Private Const adCmdTable As Long = 2            'テーブル名
Private Const adCmdStoredProc As Long = 4       'ストアド プロシージャ名
Private Const adCmdUnknown As Long = 8          '不明（既定値）
Private Const adCmdFile As Long = 256           '永続的に保存された Recordset のファイル名
Private Const adCmdTableDirect As Long = 512    '列がすべて返されるテーブル名
Private Const adCmdUnspecified As Long = -1     'コマンドタイプ引数を指定しません

Private Const adBigInt As Long = 20           '8 バイトの符号付き整数
Private Const adBinary As Long = 128          'バイナリ値
Private Const adBoolean As Long = 11          'ブール型 (Boolean) 値
Private Const adBSTR As Long = 8              '終端が null の文字列 (Unicode)
Private Const adChar As Long = 129            '文字列型 (String) 値
Private Const adCurrency As Long = 6          '通貨値 (8 バイトの符号付き整数、桁数 10,000)
Private Const adDate As Long = 7              '日付型 (Date) 値
Private Const adDBDate As Long = 133          '日付値 (yyyymmdd)
Private Const adDBTime As Long = 134          '時刻値 (hhmmss)
Private Const adDBTimeStamp As Long = 135     '日付時刻スタンプ (yyyymmddhhmmss + 10 億分の 1 桁までの小数)
Private Const adDecimal As Long = 14          '固定の精度および桁数を持つ正確な数値
Private Const adDouble As Long = 5            '倍精度浮動小数点値
Private Const adEmpty As Long = 0             '指定された値なし
Private Const adError As Long = 10            '32 ビットのエラー コード
Private Const adGUID As Long = 72             'グローバル一意識別子 (GUID)
Private Const adIDispatch As Long = 9         'OLE オブジェクトの IDispatch インターフェイスを指すポインタ
Private Const adInteger As Long = 3           '4 バイトの符号付き整数
Private Const adIUnknown As Long = 13         'OLE オブジェクトの IUnknown インターフェイスを指すポインタ
Private Const adLongVarBinary As Long = 205   'ロング バイナリ値 (Parameter オブジェクトのみ)
Private Const adLongVarChar As Long = 201     '長文字列型 (String) 値 (Parameter オブジェクトのみ)
Private Const adLongVarWChar As Long = 203    '終端が null の長文字列値 (Parameter オブジェクトのみ)
Private Const adNumeric As Long = 131         '固定の精度および桁数を持つ正確な数値
Private Const adSingle As Long = 4            '単精度浮動小数点値
Private Const adSmallInt As Long = 2          '2 バイトの符号付き整数
Private Const adTinyInt As Long = 16          '1 バイトの符号付き整数
Private Const adUnsignedBigInt As Long = 21   '8 バイトの符号なし整数
Private Const adUnsignedInt As Long = 19      '4 バイトの符号なし整数
Private Const adUnsignedSmallInt As Long = 18 '2 バイトの符号なし整数
Private Const adUnsignedTinyInt As Long = 17  '1 バイトの符号なし整数
Private Const adUserDefined As Long = 132     'ユーザー定義変数
Private Const adUseClient As Long = 3         'クライアントサイドカーソル
Private Const adVarBinary As Long = 204       'バイナリ値 (Parameter オブジェクトのみ)
Private Const adVarChar As Long = 200         '文字列型 (String) 値 (Parameter オブジェクトのみ)
Private Const adVariant As Long = 12          'OLE オートメーション バリアント型 (Variant)
Private Const adVarWChar As Long = 202        '終端が null の Unicode 文字列 (Parameter オブジェクトのみ)
Private Const adWChar As Long = 130           '終端が null の Unicode 文字列
'accdbﾌｨｰﾙﾄﾞ作成(警告なしにファイルを更新)
Public Sub fldAccdb(tgtDb As Variant, TableName As String, fileName As String, Optional FilePath As String = "", Optional newflag As Boolean = True)

If FilePath = "" Then FilePath = ThisWorkbook.path
Dim MakefilePath As String: MakefilePath = FilePath & "\" & fileName

'エラートラップ
On Error GoTo catch

  Dim cat As Object: Set cat = CreateObject("ADOX.Catalog")
  Dim tbl As Object: Set tbl = CreateObject("ADOX.Table")

  'データベースから作る場合とテーブルのみ作る場合を分ける
  Select Case newflag
  Case True 'データベース＋テーブル
    cat.Create AdoOpenCode & MakefilePath ' データベース作成
  Case False 'テーブルのみ
  End Select
  
  Dim adoCn As Object: Set adoCn = adoConnection(MakefilePath)
  cat.ActiveConnection = adoCn 'データベース接続
  Set tbl.ParentCatalog = cat
    
  Dim j As Long
  
  'フィールド（列）の設定
  Dim FieldData As String
  For j = 1 To UBound(tgtDb, 2)
    FieldData = Replace(tgtDb(1, j), vbLf, "") 'セル内改行はエラーとなる
'    Select Case True
'      Case IsDate(FieldData)
'        tbl.Columns.Append FieldData, adDate '日付型
'      Case Len(FieldData) > 255
'        tbl.Columns.Append FieldData, adLongVarChar '文字列型(長いテキスト)
'      Case Else
         tbl.Columns.Append FieldData, adVarWChar '文字列型
'    End Select
  Next j

Finally:

      tbl.Name = TableName
      cat.Tables.Append tbl 'テーブルの作成
      Set cat = Nothing
      Set tbl = Nothing
    
    Exit Sub

catch:

  Select Case Err.Number
    Case -2147217897 'ﾃﾞｰﾀﾍﾞｰｽは既に存在している
      MsgBox "ﾃﾞｰﾀﾍﾞｰｽは既に存在している"
    Case Else
      MsgBox "エラー番号  :" & Err.Number & vbCrLf & _
             "エラーの内容:" & Err.Description, vbExclamation
  End Select

  On Error GoTo 0
  Resume Finally 'Resume:エラー発生時のみ

End Sub
'accdbにレコード追加
Public Sub AddAccdb(tgtDb As Variant, TableName As String, fileName As String, Optional FilePath As String = "")

If FilePath = "" Then FilePath = ThisWorkbook.path

Dim MakefilePath As String: MakefilePath = FilePath & "\" & fileName

'テータベースへ登録
Dim i As Long
Dim j As Long
Dim jcnt As Long
Dim adoCn As Object: Set adoCn = CreateObject("ADODB.Connection")
Dim adoRs As Object: Set adoRs = CreateObject("ADODB.Recordset")
Dim adoCd As Object: Set adoCd = CreateObject("ADODB.Command")
Dim tgtFieldStr() As String
adoCn.Open AdoOpenCode & FilePath & "\" & fileName & ";"  'Accessファイルに接続
adoCd.ActiveConnection = adoCn

With adoRs
  .Source = "SELECT * FROM " & TableName
  .ActiveConnection = adoCn
    jcnt = .Fields.Count
  .CursorType = 1 'adOpenKeyset
  .LockType = 2 'adLockPessimistic
  .Open
End With

'エラートラップ
'On Error GoTo Catch

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
adoCn.BeginTrans 'トランザクション開始

With adoRs
  
  For i = 2 To UBound(tgtDb, 1)
  .AddNew
    For j = 1 To jcnt
      adoRs.Fields(j - 1) = CStr(tgtDb(i, j))
    Next j
  .Update
  Next i
  
End With

adoCn.CommitTrans 'トランザクション終了
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Finally:

    adoRs.Close: Set adoRs = Nothing
    Set adoCd = Nothing
    adoCn.Close: Set adoCn = Nothing
    
    Exit Sub

catch:

  'エラーが発生したのでデータを処理開始前に戻します
  adoCn.RollbackTrans

  MsgBox "エラー番号  :" & Err.Number & vbCrLf & _
         "エラーの内容:" & Err.Description, vbExclamation

  On Error GoTo 0
  Resume Finally 'Resume:エラー発生時のみ
 
End Sub
'Accdb用：汎用化したら消去
Private Function AdoOpenCode() As String
AdoOpenCode = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
End Function
'ADOConnection:参照設定(Microsoft ActiveX Data Objects x.x Library)**************************************************************************
'Microsoft Excel の接続をﾃﾞｰﾀ ﾌﾟﾛﾊﾞｲﾀﾞｰに公開して､Excel が使用するのと同じｾｯｼｮﾝのｺﾝﾃｷｽﾄ内でﾕｰｻﾞｰがｺｰﾄﾞを記述
Public Function adoConnection(FilePath As String) As Object

If Not PathCheck(FilePath) Then
'  Msg.Box "ファイルが参照できません。", vbExclamation
  Exit Function
End If

Dim Ext As String: Ext = FileExtensionType(FilePath)
Dim ConProv As String
Dim ExtProp As String

Select Case Ext
'  Case "mdb"
'    ConProv = "Microsoft.ACE.OLEDB.16.0"
  Case "csv", "accdb", "xlsx", "xlsm", "xlsb"
    ConProv = "Microsoft.ACE.OLEDB.12.0" 'Excel2007以降で作成したブック
  Case "mdb", "xls"
    ConProv = "Microsoft.Jet.OLEDB.4.0" 'Excel2002/2003で作成したブック
  Case "" 'MySql
  Case "" 'SQLServer
  Case "" 'Oracle
  Case "" 'ODBC
End Select

'Excelの接続文字列におけるIMEXについて******************************************
'IMEX(ImportMixedType) 列の中に数字と文字など違うデータ型が混在している場合、
'テキストとして解釈するか。省略すると自動的に判断。
'Imex=1 テキスト　Imex=2 リンクモード IMEX=0はExport
'組合せによる表示状態 (※テキストボックスに表示した場合)
'*******************************************************************************

Select Case Ext
  Case "csv"
    ExtProp = "Text;FMT=Delimited;HDR=NO" 'HDR=YES:先頭行を列名とみなす
  Case "accdb"
    ExtProp = ""
  Case "xls"
    ExtProp = "Excel 8.0;IMEX=1;Readonly=False"
  Case "xlsx"
    ExtProp = "Excel 12.0 Xml"
  Case "xlsm"
    ExtProp = "Excel 12.0 Macro"
  Case "xlsb"
    ExtProp = "Excel 12.0"
  Case "" 'MySql
  Case "" 'SQLServer
  Case "" 'Oracle
  Case "" 'ODBC
End Select

Dim tempConnection As Object: Set tempConnection = CreateObject("ADODB.Connection")
Dim pas As String: pas = "nlm"

With tempConnection
  .Provider = ConProv
  .Properties("Extended Properties") = ExtProp
  .Properties("Jet OLEDB:Database Password") = pas ''.value
  .Open FilePath 'ファイルに接続
  .CursorLocation = adUseClient 'adUseClient=3'クライアントサイドカーソルに変更して、RecordCountを取得
End With

'サーバーサイドカーソルとは取得したデータをサーバーにもつ。
'クライアントサイドカーソルとは取得したデータをクライアント側にもつ｡
'サーバーサイドにすれば、サーバーで処理してデータだけクライアントに返すのでデータ通信量が少なくなるが、
'サーバー側に負荷がかかるし、クライアント側のRecordCountが設定されず-1なんてこともある
'クライアントサイドにすれば、結果セット全体をサーバーからもらって以降の処理はクライアント側で行うので、
'サーバーの負荷が減らせるし、クライアント側の処理能力に比例して、以降の処理速度が速くなるメリットがあるかわりに、
'結果セット全体をダウンロードするのにネットワークが重くなる。

Set adoConnection = tempConnection

End Function
'対象ﾃﾞｰﾀﾍﾞｰｽのﾃｰﾌﾞﾙ削除
Public Sub adoTableDelete(FilePath As String, TableName As String)
With CreateObject("ADOX.Catalog")
  .ActiveConnection = adoConnection(FilePath)
  .Tables.Delete TableName
End With
End Sub
Private Function FileExtensionType(tgtPath As String) As String
Dim Cnt As Long: Cnt = InStrRev(tgtPath, ".")
FileExtensionType = LCase(Mid(tgtPath, Cnt + 1))
End Function
'ADORecordset:参照設定(Microsoft ActiveX Data Objects x.x Library)******************************************************************************
'基になるﾃｰﾌﾞﾙのﾚｺｰﾄﾞｾｯﾄ全体､またはｺﾏﾝﾄﾞの実行によって返された結果を表す
Public Function adoRecordset(TableName As String, adoCn As Object, Optional SQLfilter As String = "", Optional CursorType As Long = 99, Optional LockType As Long = 99) As Object

Dim tempRecordset As Object: Set tempRecordset = CreateObject("ADODB.Recordset")

If CursorType = 99 Then CursorType = adOpenForwardOnly
If LockType = 99 Then LockType = adLockReadOnly

With tempRecordset
  .ActiveConnection = adoCn
  .Source = "SELECT * FROM " & TableName 'SELECT フィールド名 FROM テーブル名　' * : すべて
  .CursorType = CursorType
  .LockType = LockType
  .filter = SQLfilter
  If adoCn Is Nothing Then
'    Msg.Box "adoConnection接続に失敗しました。", vbInformation
    End
  End If
  .Open
End With

'Sheet1のセルA1からZ50までの範囲を指定する場合
'SELECT * FROM [Sheet1$A1:Z50] WHERE [F1]=1;
'Sheet1でセル範囲は１行目から表全体を指定する場合
'SELECT * FROM [Sheet1$] WHERE [F1]=1;

'対象レコードがなくても戻り値を返す
RecFlag = (tempRecordset.EOF) '抽出したRSにデータが存在しているかどうか
Set adoRecordset = tempRecordset

End Function
'ADOCommand:参照設定(Microsoft ActiveX Data Objects x.x Library)******************************************************************************
'Recordsetｵﾌﾞｼﾞｪｸﾄのﾚｺｰﾄﾞを取得したり､一括操作を実行したり､ﾃﾞｰﾀﾍﾞｰｽの構造を操作したりできる
Private Function adoCommand(adoCn As Object) As Object
Dim tempCommand As Object: Set tempCommand = CreateObject("ADODB.Command")
tempCommand.ActiveConnection = adoCn
Set adoCommand = tempCommand
End Function
'Dictionary:参照設定(Microsoft Scripting Runtime)**********************************************************************************************
Private Function adoDic() As Object
Dim tempDic As Object: Set tempDic = CreateObject("Scripting.Dictionary")
Set adoDic = tempDic
End Function
Public Function adoFieldkeyNo(adoCn As Object, adoRs As Object, keyFieldName As String, jcnt As Long) As Long
Dim j As Long
Dim tgtFieldStr As Variant
ReDim tgtFieldStr(0 To jcnt - 1) As String
For j = 1 To jcnt
  tgtFieldStr(j - 1) = adoRs.Fields(j - 1).Name 'フィールド名取得（Source　&　ActiveConnectionの後）
  If keyFieldName = adoRs.Fields(j - 1).Name Then
    adoFieldkeyNo = j - 1
    Exit Function
  End If
Next j
End Function
'Field名をDictionaryｵﾌﾞｼﾞｪｸﾄに変換する
Public Function GetFieldDic(FilePath As String, TableName As String, Optional ColLim As Long = 0) As Object

Dim adoCn As Object: Set adoCn = adoConnection(FilePath)
Dim adoRs As Object: Set adoRs = adoRecordset(TableName, adoCn, "", adOpenStatic, adLockReadOnly)

Dim jcnt As Long: jcnt = IIf(ColLim = 0, adoRs.Fields.Count, ColLim)
Dim tempDicj As Object: Set tempDicj = adoDic
Dim tempFieldsName As String
Dim j As Long
For j = 1 To jcnt
  tempFieldsName = adoRs.Fields(j - 1).Name
  tempDicj(Application.Clean(Trim(tempFieldsName))) = j
Next j

adoRs.Close: Set adoRs = Nothing
adoCn.Close: Set adoCn = Nothing

Set GetFieldDic = tempDicj
Set tempDicj = Nothing
    
End Function
'ﾌｨｰﾙﾄﾞﾃﾞｰﾀのみを配列にして取得する
Public Function adoFieldDic(adoRs As Object) As Object
If adoRs.EOF Then Exit Function
Dim tempDicj As Object: Set tempDicj = adoDic
Dim tempFieldsName As String
Dim j As Long
For j = 1 To adoRs.Fields.Count
  tempFieldsName = adoRs.Fields(j - 1).Name
  tempDicj(Application.Clean(Trim(tempFieldsName))) = j
Next j
Set adoFieldDic = tempDicj
Set tempDicj = Nothing
End Function

'ADOで対象ﾌｧｲﾙを2次元配列に変換しDictionaryｵﾌﾞｼﾞｪｸﾄに変換する
Public Function AdoArr(FilePath As String, TableName As String, Optional SQLfilter As String = "", Optional ColLim As Long = 0) As Object

'Debug.Print filepath
Dim adoCn As Object: Set adoCn = adoConnection(FilePath)
Dim adoRs As Object: Set adoRs = adoRecordset(TableName, adoCn, SQLfilter, adOpenStatic, adLockReadOnly)

If adoRs Is Nothing Then Exit Function

'レコードセットの内容を変数に格納
Dim jcnt As Long: jcnt = IIf(ColLim = 0, adoRs.Fields.Count, ColLim)

Dim adoCheckBool As Boolean: adoCheckBool = adoRs.EOF 'ﾚｺｰﾄﾞ無し

Dim tempFields() As Variant: tempFields = adoFields(adoRs)
Dim tempRecord() As Variant

If Not adoCheckBool Then
  tempRecord = adoRecords(adoRs, jcnt)
End If

adoRs.Close: Set adoRs = Nothing
adoCn.Close: Set adoCn = Nothing

Dim TempArr As Object: Set TempArr = adoDic
TempArr("fields") = tempFields
TempArr("record") = tempRecord

Set AdoArr = TempArr
Set TempArr = Nothing

End Function
'ﾌｨｰﾙﾄﾞﾃﾞｰﾀのみを配列にして取得する
Public Function adoFields(adoRs As Object, Optional jcnt As Long = 0) As Variant
If adoRs Is Nothing Then Exit Function
Dim j As Long
Dim tempFields() As Variant
jcnt = IIf(jcnt = 0, adoRs.Fields.Count, jcnt)
ReDim tempFields(1 To 1, 1 To jcnt) As Variant
For j = 1 To jcnt
  tempFields(1, j) = adoRs.Fields(j - 1).Name
Next j
adoFields = tempFields
End Function
'ﾚｺｰﾄﾞﾃﾞｰﾀのみを配列にして取得する
Public Function adoRecords(adoRs As Object, Optional jcnt As Long = 0, Optional PageCnt As Long = 0, Optional pageNo As Long = 1) As Variant
If adoRs Is Nothing Then Exit Function
Dim icnt As Long: icnt = adoRs.RecordCount
If icnt = -1 Then Exit Function
jcnt = IIf(jcnt = 0, adoRs.Fields.Count, jcnt)
adoRs.PageSize = IIf(PageCnt = 0, icnt, PageCnt) 'Recordset ｵﾌﾞｼﾞｪｸﾄの1ページあたりのﾚｺｰﾄﾞ数
adoRs.AbsolutePage = pageNo
'レコードの内容を文字列で取得
Dim oneD As Variant
oneD = Split(adoRs.GetString(adClipString, adoRs.PageSize, vbTab, "<#(仮)改行#>"), "<#(仮)改行#>")
Dim tempRecord() As Variant 'WorksheetFunction.Transpose(adoRs.GetRows)'一括配列化
ReDim tempRecord(1 To icnt, 1 To jcnt) As Variant
Dim colArr As Variant
Dim i As Long
Dim j As Long
For i = 1 To icnt
  If oneD(i - 1) = "" Then Exit For
  colArr = Split(oneD(i - 1), vbTab)
  For j = 1 To jcnt
    tempRecord(i, j) = colArr(j - 1)
  Next j
Next i
adoRecords = tempRecord
End Function
'指定accdbファイルのfield名を配列に取得
Public Function GetFieldsArr(FilePath As String, TableName As String, Optional SQLfilter As String = "", Optional ColLim As Long = 0) As Variant
GetFieldsArr = GetAdoArr(FilePath, TableName, SQLfilter, ColLim, "Fields")
End Function
'指定accdbファイルのRecordを配列に取得
Public Function GetRecordsArr(FilePath As String, TableName As String, Optional SQLfilter As String = "", Optional ColLim As Long = 0) As Variant
GetRecordsArr = GetAdoArr(FilePath, TableName, SQLfilter, ColLim, "Records")
End Function
Public Function GetRecordsAdo(adoCn As Object, TableName As String, SQLfilter As String) As Variant
Dim adoRs As Object: Set adoRs = adoRecordset(TableName, adoCn, SQLfilter)
If adoRs Is Nothing Then Exit Function
If adoRs.EOF Then Exit Function
GetRecordsAdo = adoRecords(adoRs)
adoRs.Close: Set adoRs = Nothing
End Function
Private Function GetAdoArr(FilePath As String, TableName As String, Optional SQLfilter As String = "", Optional ColLim As Long = 0, Optional GetType As String = "Fields") As Variant
Dim adoCn As Object: Set adoCn = adoConnection(FilePath)
Dim adoRs As Object: Set adoRs = adoRecordset(TableName, adoCn, SQLfilter, adOpenStatic, adLockReadOnly)
If adoRs Is Nothing Then Exit Function
If adoRs.EOF Then Exit Function
Dim jcnt As Long: jcnt = IIf(ColLim = 0, adoRs.Fields.Count, ColLim)
Select Case GetType
  Case "Fields": GetAdoArr = adoFields(adoRs)
  Case "Records": GetAdoArr = adoRecords(adoRs, jcnt)
End Select
adoRs.Close: Set adoRs = Nothing
adoCn.Close: Set adoCn = Nothing
End Function

'ListViewに表示する
Public Function adoRead(TFctrl As Control, FilePath As String, TableName As String, Optional SQLfilter As String = "", Optional ColLim As Long = 0) As Object
Dim adoData As Object: Set adoData = AdoArr(FilePath, TableName, SQLfilter, ColLim)
Call CreateAll(TFctrl, adoData("fields"), adoData("record"), True)
Set adoRead = adoData
End Function
Public Function adoReadPageCtrl(TF As UserForm, FilePath As String, TableName As String, Optional SQLfilter As String = "", Optional PageCnt As Long = 0, Optional pageNo As Long = 1, Optional ColLim As Long = 0) As Variant

Dim adoCn As Object: Set adoCn = adoConnection(FilePath)
Dim adoRs As Object: Set adoRs = adoRecordset(TableName, adoCn, SQLfilter, adOpenStatic, adLockReadOnly)

If adoRs Is Nothing Then
  TF.ListView1.ListItems.clear
  Exit Function
End If

'レコードセットの内容を変数に格納
Dim TrimColFlag As Boolean: TrimColFlag = Not (ColLim = 0)
Dim jcnt As Long
Select Case TrimColFlag
  Case True: jcnt = ColLim
  Case Else: jcnt = adoRs.Fields.Count
End Select
Dim tempFields() As Variant: tempFields = adoFields(adoRs)
Dim tempRecord() As Variant: tempRecord = adoRecords(adoRs, jcnt, PageCnt, pageNo)
Dim icnt As Long: icnt = adoRs.RecordCount
'ページ設定
'***********************************************************************
Dim tgtctrl As Control: Set tgtctrl = TF.SB1
If Not tgtctrl Is Nothing Then
  If PageCnt = 0 Then
    tgtctrl.Max = 1
  Else
    tgtctrl.Max = Application.WorksheetFunction.RoundUp(icnt / PageCnt, 0)  'ページ制限
    If IsEmpty(tempRecord) Then
      TF.ListView1.ListItems.clear
      Exit Function '抽出したRSにデータが存在しているかどうかをチェック
    End If
  End If
End If
'***********************************************************************

'TF.Cnt = iCnt & " 件"

adoRs.Close: Set adoRs = Nothing
adoCn.Close: Set adoCn = Nothing
Call CreateAll(TF.ListView1, tempFields, tempRecord, TrimColFlag)

TF.page.Tag = icnt
TF.page = TF.SB1.Value & " / " & TF.SB1.Max

End Function

Public Sub adoSQL(FilePath As String, adoTABLENAME As String, MODE As String, tgtInput As Variant, keyFieldName As String, Optional skipmsgflag As Boolean = False, Optional keySQLfilter As String = "")

Dim adoCn As Object: Set adoCn = adoConnection(FilePath)
Dim adoCd As Object: Set adoCd = adoCommand(adoCn)
Dim adoRs As Object: Set adoRs = adoRecordset(adoTABLENAME, adoCn, "", adOpenStatic, adLockReadOnly)

If adoRs Is Nothing Then Exit Sub
Dim jcnt As Long: jcnt = adoRs.Fields.Count
Dim keyNo As Long: keyNo = adoFieldkeyNo(adoCn, adoRs, keyFieldName, jcnt)
Dim easyFlag As Boolean: easyFlag = (keySQLfilter = "")


With adoRs
  .Close
  .ActiveConnection = adoCn
  
  '.Source = "SELECT * FROM " & adoTABLENAME & " WHERE [" & keyFieldName & "]='" & CStr(tgtInput(keyNo)) & "'"
  .Source = "SELECT * FROM " & adoTABLENAME 'SELECT フィールド名 FROM テーブル名　' * : すべて
  .CursorType = adOpenKeyset
  .LockType = adLockOptimistic
  
  .filter = IIf(easyFlag, "[" & keyFieldName & "]='" & CStr(tgtInput(keyNo)) & "'", keySQLfilter)
  'keySQLfilter = "[BOX設計_設備名] LIKE '%" & tmpFaci & "%'" 'AND [検査番号] <> '0000000000'"  'AND [ｺﾝﾏ5本ｶｰﾄﾞ区分] <> '3'" ' AND [ﾛｯﾄ番号] = 'L5140Z0'
  
  
  'adLockReadOnly 読み取り専用(データの更新は不可)(既定値)
  'adLockPessimistic レコードごとの排他的ロック(データソースでロック)
  'adLockOptimistic レコードごとの共有的ロック(Updateメソッドを呼び出した場合のみ共有的ロック)
  'adLockBatchOptimistic 共有的バッチ更新
  
  .Open
  'icnt = .RecordCount 'データ数をチェック複数同時編集でない限り
End With

Dim msgStr(1 To 5) As String
Select Case True
    Case MODE = "ADDNEW"
        msgStr(1) = "登録がキャンセルされました。"
        msgStr(2) = "新規登録しますか？"
        msgStr(3) = "既に登録されているデータがあります。" & vbCrLf & "上書きしますか？"
        msgStr(4) = "新規登録しました！"
        msgStr(5) = "キーが重複しているため、登録がキャンセルされました。"
    Case MODE = "UPDATE"
        msgStr(1) = "登録がキャンセルされました。"
        msgStr(2) = "上書き登録しますか？"
        msgStr(3) = "登録データがありません。新規登録しますか？"
        msgStr(4) = "更新しました！"
        msgStr(5) = ""
    Case MODE = "DELETE"
        msgStr(1) = "削除がキャンセルされました。"
        msgStr(2) = "登録データを削除しますか？"
        msgStr(3) = "登録されていないか、既に削除済みです。"
        msgStr(4) = "削除しました！"
        msgStr(5) = ""
    Case Else
End Select

'エラートラップ
On Error GoTo catch

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
adoCn.BeginTrans 'トランザクション開始

Select Case MODE
  Case "ADDNEW", "UPDATE"
  Case "DELETE"
    If adoRs.EOF Then '抽出したRSにデータが存在しているかどうかをチェック
      'レコードが存在しない
      If Not skipmsgflag Then msg.Box msgStr(3), vbExclamation
    Else
      'レコードが存在している
      If skipmsgflag Then GoTo skipDELETE
      If msg.Box(msgStr(2), vbYesNo + vbQuestion) = vbYes Then
skipDELETE:
        With adoRs
          .Delete
          .Update
        End With
        If Not skipmsgflag Then msg.Box msgStr(4), vbInformation
      Else
        msg.Box msgStr(1), vbExclamation
      End If
    End If
    GoTo skipStep
  Case Else
    '指定外モード(想定外エラー)
    Call Err.Raise(5, "NoMode", "正しい引数(モード)が入力されていません。")
    GoTo skipStep
End Select

Dim j As Long
'レコードが存在しない=新規>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Dim TempMsg As String: TempMsg = ""
If adoRs.EOF Then '抽出したRSにデータが存在しているかどうかをチェック
  Select Case MODE
  Case "ADDNEW": TempMsg = msgStr(2)
  Case "UPDATE": TempMsg = msgStr(3)
  End Select
  
  If skipmsgflag Then GoTo skipADDNEW
  If msg.Box(TempMsg, vbYesNo + vbQuestion) = vbYes Then
skipADDNEW:
    adoRs.AddNew
      For j = 1 To jcnt
        adoRs.Fields(j - 1) = CStr(tgtInput(j - 1))
      Next j
    adoRs.Update
    If Not skipmsgflag Then msg.Box msgStr(4), vbInformation
  Else
    msg.Box msgStr(1), vbExclamation
  End If
  GoTo skipStep
End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'レコードが存在している<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Select Case MODE
  Case "ADDNEW"
    MODE = "UPDATE"
    GoTo skipMODEupdate '更新に切り替え(メッセージは新規登録扱い)
    'If Not skipmsgflag Then Msg.Box msgStr(5), vbInformation
    
  Case "UPDATE"
skipMODEupdate:
    If skipmsgflag Then GoTo skipUPDATE
    If msg.Box(msgStr(2), vbYesNo + vbQuestion) = vbYes Then
skipUPDATE:
      Dim sbSQL As String: sbSQL = MODE & " " & adoTABLENAME & " SET "
      '更新用のデータ読み込み
      For j = 1 To jcnt - 1
        sbSQL = sbSQL & "[" & adoRs.Fields(j).Name & "]='" & CStr(tgtInput(j)) & "'"
        If j <> jcnt - 1 Then sbSQL = sbSQL & ", "
      Next j
      
      '更新対象
      If easyFlag Then
        sbSQL = sbSQL & "WHERE [" & adoRs.Fields(keyNo).Name & "]='" & CStr(tgtInput(keyNo)) & "'"
      Else
        'ﾌｨﾙﾀｰ使用時は対象行に対して行う
        sbSQL = sbSQL & "WHERE " & keySQLfilter
      End If
      
      adoCd.CommandText = sbSQL
      adoCd.Execute '処理を実行
      If Not skipmsgflag Then msg.Box msgStr(4), vbInformation
    Else
      msg.Box msgStr(1), vbExclamation
    End If
    
End Select
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

skipStep:

    adoCn.CommitTrans 'トランザクション終了
    
Finally:
    
    Set adoCd = Nothing
    adoRs.Close: Set adoRs = Nothing
    adoCn.Close: Set adoCn = Nothing
    Exit Sub

catch:

  'エラーが発生したのでデータを処理開始前に戻します
  adoCn.RollbackTrans
  
  'BeginTrans   :   処理前に実行
  'CommitTrans  :   正常終了の場合､変更を保存しトランザクションを終了
  'RollbackTrans:   変更をキャンセルしトランザクションを終了
  
  msg.Box "エラー番号  :" & Err.Number & vbCrLf & _
         "エラーの内容:" & Err.Description, vbExclamation
  
  On Error GoTo 0
  
  Resume Finally 'Resume:エラー発生時のみ

End Sub

'指定txtﾌｧｲﾙをAccdbﾌｧｲﾙに変換する
Public Sub TxtToAccdb(FilePath As String, Optional shName As String = "", Optional TableName As String = "")

If Not PathCheck(FilePath) Then Exit Sub

Dim OneDimTemp As Variant
OneDimTemp = GetOneDimData(FilePath) 'TextFileFunc

Dim Temp() As Variant
'Dim Con As Variant: Con = Range(Ms1.Cells(1, 1), Ms1.Cells(5, UC(Ms1))) 'Ms1が汎用性なし
'temp = ConvDataExt(OneDimTemp, Con)

Call ExpAccdb(Temp(), TableName)

End Sub
'指定EXCELﾌｧｲﾙをAccdbﾌｧｲﾙに変換する
Public Sub XlsToAccdb(ArrSheetName As Variant, ArrTableName As Variant, targetFilename As String, accdbFilename As String)

  Application.ScreenUpdating = False
  
  Dim servPathDic As Object: Set servPathDic = getAddress(targetFilename) '指定ﾊﾟｽの場合は引数で指定
  Dim savePathDic As Object: Set savePathDic = getAddress(accdbFilename)
  
  Dim ServPath As String: ServPath = servPathDic("folder")
  Dim SavePath As String: SavePath = savePathDic("folder")
  Dim FilePath As String: FilePath = servPathDic("file")
  
  If Not PathCheck(ServPath) Then Exit Sub '参照先チェック
  If Not PathCheck(SavePath) Then Exit Sub '保存先チェック
  
  Dim sourceWB As Workbook: Set sourceWB = xlsFileOpenWb(FilePath)
  
  Call tgtFileDelete(accdbFilename, SavePath)
  
  Dim i As Long
  Dim shCnt As Long: shCnt = UBound(ArrSheetName)
  Dim wsData() As Variant
  
  Dim flag As Boolean: flag = True
  
  'ExpAccdb関数：1回目でﾌｧｲﾙ作成&ﾃｰﾌﾞﾙ作成、2回目以降はﾃｰﾌﾞﾙ作成のみ
  For i = 1 To shCnt
    wsData = shDataGet(CStr(ArrSheetName(i)), 1)
    Call ExpAccdb(wsData, CStr(ArrTableName(i)), accdbFilename, SavePath, flag) '2つ目以降はFlag=False
    Erase wsData
    flag = False
  Next i
  
  Call xlsFileClose(sourceWB, False)
  
  Application.ScreenUpdating = True
  
End Sub
Private Function getAddress(tempFileName As String, Optional TempFolderPath As String = "") As Object
Dim tempDic As Object: Set tempDic = CreateObject("Scripting.Dictionary")
Dim TempPath As String
Select Case TempFolderPath
  Case "":    TempPath = ThisWorkbook.path
  Case Else:  TempPath = TempFolderPath
End Select
tempDic("folder") = TempPath
tempDic("file") = TempPath & "\" & tempFileName
Set getAddress = tempDic
End Function
'CSV出力(警告なしにファイルを更新)
Public Sub ExpAccdb(Temp As Variant, TableName As String, Optional fileName As String = "", Optional FilePath As String = "", Optional newflag As Boolean = True)
Call fldAccdb(Temp, TableName, fileName, FilePath, newflag)
Call AddAccdb(Temp, TableName, fileName, FilePath)
End Sub
'対象ﾌｧｲﾙﾊﾟｽの指定ｼｰﾄ名ｼｰﾄﾃﾞｰﾀを配列に格納する
Private Function OpenXlsWs(FilePath As String, Optional shName As String = "") As Variant

Application.ScreenUpdating = False

  Dim actWb As Workbook
  Set actWb = xlsFileOpenWb(FilePath) 'ReadOnly
  
  Dim DbWs As Worksheet
  If shName = "" Then
    shName = actWb.ActiveSheet.Name
    Set DbWs = actWb.ActiveSheet
  Else
    Set DbWs = actWb.Worksheets(shName)
  End If
  
  OpenXlsWs = UseWs(DbWs)
  
  'Application.DisplayAlerts = False
  Workbooks(GetFileName(FilePath)).Close
  'Application.DisplayAlerts = True
  
Application.ScreenUpdating = True

End Function

'対象ﾌｧｲﾙがあったら削除
Public Sub tgtFileDelete(fileName As String, Optional FilePath As String = "")
  Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim tgtPath As String: tgtPath = FilePath & "\" & fileName
  If fso.fileExists(tgtPath) Then Call fso.DeleteFile(tgtPath)
  Set fso = Nothing
End Sub

