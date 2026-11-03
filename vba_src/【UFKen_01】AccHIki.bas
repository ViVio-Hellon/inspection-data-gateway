Attribute VB_Name = "【UFKen_01】AccHIki"
'Option Explicit
Private Sika(1 To 10000, 1 To 6)   '2次元配列
'不必要なﾃﾞｰﾀ処理を飛ばす
Public skipFlag As Byte
Public listflag As Byte
Private Add1 As Long
Private Add2 As Long
Private MaxRow As Long  '読込ﾃｷｽﾄの行数取得用
Private Time1 As Date '時間計測用
'Private Buff() As Byte 'ﾃﾞｰﾀ格納用 ReDim用
Private ws1 As Worksheet

Public Function Arrbox(target As String, Optional SQLfilter As String = "", Optional Use As String = "Record", Optional path As String = "") As Variant '2021.11.25_SQL配列呼び出しやすく
'Arrbox("SIKALOTNOW.accdb", "板", "Fields",path)
Dim FPath As String
'On Error Resume Next
'If UF1.Load.Caption = "off" Then Exit Function 'ﾌｫｰﾑｲﾍﾞﾝﾄ停止
If path = "" Then
 FPath = MainPath
Else
 FPath = NextPath
End If

Dim ResultT As Boolean: ResultT = IsOver12Hours '2024.9.3

Dim FilePath As String: FilePath = FPath & target
If FPath = MainPath Then
  If ReadCheck(FilePath, 1) = False Or TimeDesignation Or FileDesignation = False Then   '1：ファイル,2：フォルダ
    FPath = NextPath
    FilePath = FPath & target
    If ReadCheck(NextPath & target, 1) = False Then '梱包課ﾌｧｲﾙを使う*1日更新
      If ResultT = False Then MsgBox "日稲:" & target & "がありません" & "_ReadCheck_Arrbox", vbCritical
      Exit Function
    End If
  End If
End If
Dim TableName As String: TableName = "仕掛"
Dim LotNo As String:     LotNo = SQLfilter

Select Case SQLfilter
Case "":       SQLfilter = "" 'ﾌｨﾙﾀ不要
Case "板":     SQLfilter = "[ｵｰﾀﾞｰ板丈]<>0" '"" 'RNE変更後フィルターをかけると検索しなくなった(検索数が減った)_2022.09.14"[ｵｰﾀﾞｰ板丈]<>0 AND [残仕掛設備ｺｰｽ] LIKE '*KEN*' AND  [BOX設計_設備名] = 'KEN' "
Case "コイル": SQLfilter = "[ｵｰﾀﾞｰ板丈]=0  AND  [BOX設計_設備名] = 'KEN' "
Case Else:     SQLfilter = "[ﾛｯﾄ番号]='" & LotNo & "'" 'ﾛｯﾄNo検索
End Select

'IsEmptyArray(a As Variant) As Boolean 'trueの時配列が空
Select Case Use
Case "Record"
  Arrbox = GetRecordsArr(FilePath, TableName, SQLfilter)
  SQLfilter = ""
  Exit Function
Case "Fields"
  Arrbox = GetFieldsArr(FilePath, TableName, "")
  SQLfilter = ""
  Exit Function
End Select

If IsEmptyArray(Arrbox) Then '空の時***ファイル不備のため日稲側も参照させる
  Select Case True
  Case FPath = NextPath
'    MsgBox "\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\仕掛\" & Target & "が空です" & "_Arrbox", vbCritical 'ここでは処理は止めない_2021.06.23
    SQLfilter = ""
    Exit Function
  Case FPath = MainPath
    FPath = NextPath
    FilePath = FPath & target
    If ReadCheck(FilePath, 1) = False Then '1：ファイル,2：フォルダ
'      MsgBox filepath & "がありません" & "_Arrbox", vbCritical '台帳も日軽稲沢もSikaHikiがない
      SQLfilter = ""
      Exit Function
    End If
  End Select
  
    Select Case Use
    Case "Record"
      Arrbox = GetRecordsArr(FilePath, TableName, SQLfilter)
      SQLfilter = ""
      Exit Function
    Case "Fields"
      Arrbox = GetFieldsArr(FilePath, TableName, SQLfilter)
      SQLfilter = ""
      Exit Function
    End Select
  
End If
If IsEmpty(Arrbox) Then '該当ロットなし***続行
  SQLfilter = ""
  Exit Function
End If
End Function

Public Function Lotflag(target As String, shName As String) As Boolean '更新ﾌﾗｸﾞ 一覧
Dim a As Long, i As Long, k As Long, x As Long
Dim FPath As String
Select Case True
Case UF1.Test:  FPath = MainPath 'ThisWorkbook.Path & "\"
Case Else:      FPath = MainPath
End Select
Dim FilePath As String: FilePath = FPath & target
On Error Resume Next
If WsCheck(shName) Then
  Dim Lots As Worksheet: Set Lots = ThisWorkbook.Sheets(shName)
  Dim Date1 As Date: Date1 = CDate(Lots.Cells(1, UC(Lots, 1)))
  Dim Date2 As Date: Date2 = FileDateTime(FilePath)
  If Date2 > Date1 Then '更新ﾌｧｲﾙのほうが新しい
    Lotflag = True
  End If
  
  If Lotflag = False Then 'ﾌｧｲﾙがない場合(鯖繋がっていない場合)
    FPath = NextPath
    FilePath = FPath & target
    Date2 = FileDateTime(FilePath)
    If Date2 > Date1 Then '更新ﾌｧｲﾙのほうが新しい
      Lotflag = True
    End If
  End If
  
Else 'ｼｰﾄなし
    Lotflag = True
End If
End Function
Public Sub AccUpdate() 'LOT
Application.ScreenUpdating = False
Dim FPath As String
Select Case True
Case UF1.Test:  FPath = MainPath 'ThisWorkbook.Path & "\" '"C:\Users\NLM Administrator\Desktop\"
Case Else:      FPath = MainPath
End Select
Dim target As String: target = "SIKALOTNOW.accdb"
Dim FilePath As String: FilePath = FPath & target
If ReadCheck(FilePath, 1) = False Or TimeDesignation Or FileDesignation = False Then  '1：ファイル,2：フォルダ
  FPath = NextPath
  FilePath = FPath & target
  If ReadCheck(NextPath & target, 1) = False Then '梱包課ﾌｧｲﾙを使う*1日更新
    MsgBox "日稲:" & target & "がありません" & "_ReadCheck_AccUpdate", vbCritical
    Exit Sub
  End If
End If
Dim TableName As String: TableName = "仕掛"
Dim SQLfilter As String
SQLfilter = "[ｵｰﾀﾞｰ板丈]<>0 AND [残仕掛設備ｺｰｽ] LIKE '*KEN*' AND  [BOX設計_設備名] = 'KEN' " 'ﾃﾞｰﾀｻｲｽﾞ大きいため、ﾌｨﾙﾀ必須ﾚﾍﾞﾙ
Dim Temp As Variant: Temp = GetRecordsArr(FilePath, TableName, SQLfilter)
Dim Ftemp As Variant: Ftemp = GetFieldsArr(FilePath, TableName, "")
Dim FileTime As String: FileTime = FileDateTime(FilePath)


If (IsEmpty(Temp)) Or (Format$(FileTime, "yyyy/mm/dd") <> Format$(Now, "yyyy/mm/dd")) Then 'ﾌｧｲﾙが読めても配列が不完全だった場合'ファイル更新日が当日でない場合_2021.11.12
  FPath = NextPath
  FilePath = FPath & target
  If ReadCheck(FilePath, 1) = False Then '1：ファイル,2：フォルダ
      MsgBox "日稲:" & target & "がありません" & "_ReadCheck_AccUpdate", vbCritical
      Exit Sub
  End If
  Temp = GetRecordsArr(FilePath, TableName, SQLfilter)
  Ftemp = GetFieldsArr(FilePath, TableName, "")
End If
If IsEmpty(Temp) Then Exit Sub 'それでもEmptyの時は停止

Dim LR As Long: LR = UBound(Temp, 1)
Dim LC As Long: LC = UBound(Temp, 2)
Call AddWS("LOT一覧")
Dim HIK As Worksheet: Set HIK = ThisWorkbook.Sheets("LOT一覧")
HIK.Columns("A:E").NumberFormatLocal = "@" '文字列

Range(HIK.Cells(2, 1), HIK.Cells(LR + 1, LC)) = Temp 'ﾌｨｰﾙﾄﾞ分+1
Range(HIK.Cells(1, 1), HIK.Cells(1, LC)) = Ftemp

HIK.Cells(1, UBound(Temp, 2) + 1) = FileDateTime(FilePath)
With HIK.Cells.Font
.Name = "Meiryo UI"
.Size = 8
.TintAndShade = 0
HIK.Cells.WrapText = False '2021.02.05_追加
HIK.Cells.EntireColumn.AutoFit
End With
HIK.Visible = False
'MsgBox "Hiki実行"
Application.ScreenUpdating = True
End Sub

Public Sub AccUpdate_Etc(target As String, shName As String) 'ﾌｨﾙﾀ不要の時のみ "SIKAODRNOW.accdb" "SIKAHIKINOW.accdb"
Application.ScreenUpdating = False
Dim FPath As String
Select Case True
Case UF1.Test:  FPath = MainPath 'ThisWorkbook.Path & "\"
Case Else:      FPath = MainPath
End Select

Dim FilePath As String: FilePath = FPath & target
If ReadCheck(FilePath, 1) = False Or TimeDesignation Or FileDesignation = False Then  '1：ファイル,2：フォルダ
  FPath = NextPath
  FilePath = FPath & target
  If ReadCheck(NextPath & target, 1) = False Then '梱包課ﾌｧｲﾙを使う*1日更新
    MsgBox "日稲:" & target & "がありません" & "_ReadCheck_AccUpdate_Etc", vbCritical
    Exit Sub
  End If
End If
Dim TableName As String: TableName = "仕掛"
Dim SQLfilter As String
SQLfilter = "" 'ﾌｨﾙﾀ不要
Dim Temp As Variant: Temp = GetRecordsArr(FilePath, TableName, SQLfilter)
Dim Ftemp As Variant: Ftemp = GetFieldsArr(FilePath, TableName, "")
Dim FileTime As String: FileTime = FileDateTime(FilePath)

If (IsEmpty(Temp)) Or (Format$(FileTime, "yyyy/mm/dd") <> Format$(Now, "yyyy/mm/dd")) Then 'ﾌｧｲﾙが読めても配列が不完全だった場合'ファイル更新日が当日でない場合_2021.11.12
  FPath = NextPath
  FilePath = FPath & target
  Temp = GetRecordsArr(FilePath, TableName, SQLfilter)
  Ftemp = GetFieldsArr(FilePath, TableName, "")
End If
If IsEmpty(Temp) Then Exit Sub 'それでもEmptyの時は停止

Dim LR As Long: LR = UBound(Temp, 1)
Dim LC As Long: LC = UBound(Temp, 2)
Call AddWS(shName)
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets(shName)
sh.Cells.NumberFormatLocal = "@"   '文字列
Range(sh.Cells(2, 1), sh.Cells(LR + 1, LC)) = Temp
Range(sh.Cells(1, 1), sh.Cells(1, LC)) = Ftemp

sh.Cells(1, UBound(Temp, 2) + 1) = FileDateTime(FilePath)
With sh.Cells.Font
.Name = "Meiryo UI"
.Size = 8
.TintAndShade = 0
sh.Cells.EntireColumn.AutoFit
End With
sh.Visible = False
Application.ScreenUpdating = True
End Sub
Public Sub Acc_Etc(target As String, shName As String) '主にcoil用_2021.01.28
Application.ScreenUpdating = False
Dim FPath As String: FPath = ThisWorkbook.path & "\"

Dim FilePath As String: FilePath = FPath & target
If ReadCheck(FilePath, 1) = False Or TimeDesignation Or FileDesignation = False Then  '1：ファイル,2：フォルダ
  If ReadCheck(NextPath & target, 1) = False Then '梱包課ﾌｧｲﾙを使う*1日更新
    MsgBox "日稲:" & target & "がありません" & "_ReadCheck_AccUpdate_Etc", vbCritical
    Exit Sub
  End If
End If
Dim TableName As String: TableName = "仕掛"
Dim SQLfilter As String
SQLfilter = "[受注板丈]=0"
Dim Temp As Variant: Temp = GetRecordsArr(FilePath, TableName, SQLfilter)
Dim Ftemp As Variant: Ftemp = GetFieldsArr(FilePath, TableName, "")

If IsEmpty(Temp) Then  'ﾌｧｲﾙが読めても配列が不完全だった場合
  FPath = NextPath
  FilePath = FPath & target
  Temp = GetRecordsArr(FilePath, TableName, SQLfilter)
  Ftemp = GetFieldsArr(FilePath, TableName, "")
End If
If IsEmpty(Temp) Then Exit Sub 'それでもEmptyの時は停止

Dim LR As Long: LR = UBound(Temp, 1)
Dim LC As Long: LC = UBound(Temp, 2)
Call AddWS(shName)
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets(shName)
sh.Cells.NumberFormatLocal = "@"   '文字列
Range(sh.Cells(2, 1), sh.Cells(LR, LC)) = Temp
Range(HIK.Cells(1, 1), HIK.Cells(1, LC)) = Ftemp

sh.Cells(1, UBound(Temp, 2) + 1) = FileDateTime(FilePath)
With sh.Cells.Font
.Name = "Meiryo UI"
.Size = 8
.TintAndShade = 0
sh.Cells.EntireColumn.AutoFit
End With
sh.Visible = False
Application.ScreenUpdating = True
End Sub

Public Sub Order_Update(n As String)
Application.ScreenUpdating = False
Dim FileTime As String, FilePath As String
Dim ToDate As String: ToDate = Format$(Date, "yyyymmdd;@") '当日
Dim dt As String: dt = DateAdd("d", -1, Date) '夜勤の日付またぎ前日扱い
Dim YeDate As String: YeDate = Format$(dt, "yyyymmdd;@") '夜勤用1日戻し
Dim fm As String: fm = "sikaodr" 'ﾃｷｽﾄ名
Dim a As Long, i As Long, Last As Long
Dim flag As Boolean: flag = False
Select Case n
Case "Test":    FilePath = "C:\Users\NLM Administrator\Desktop"
Case Else:      FilePath = "\\Nlmfanago02d\仕掛ﾃﾞｰﾀ"
End Select
If WsCheck("VC一覧") Then
  flag = True
  Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("VC一覧")
  Last = ws.Cells(1, Columns.Count).End(xlToLeft).Column
End If
  If Format$(Now, "hh:mm") >= CDate("6:35") And Format$(Now, "hh:mm") <= CDate("23:59") Then
    If ReadCheck((FilePath & "\" & fm & ToDate & ".txt"), 1) = True Then
      FileTime = FileDateTime(FilePath & "\" & fm & ToDate & ".txt") '最後に変更された日付と時刻
      If flag = False Then
        Call オーダー更新(FilePath & "\" & fm & ToDate & ".txt", FileTime)
      Else
        If FileTime <> ws.Cells(1, Last) Then Call オーダー更新(FilePath & "\" & fm & ToDate & ".txt", FileTime)
      End If
    Else
      MsgBox "当日引当ﾌｧｲﾙがない、もしくはﾈｯﾄﾜｰｸにつながっていません", vbCritical
      Exit Sub
    End If
  ElseIf Format$(Now, "hh:mm") >= CDate("00:00") And Format$(Now, "hh:mm") <= CDate("6:34") Then
    If ReadCheck((FilePath & "\" & fm & YeDate & ".txt"), 1) = True Then
      FileTime = FileDateTime(FilePath & "\" & fm & YeDate & ".txt") '夜勤の日付またぎ前日扱い
      If flag = False Then
        Call オーダー更新(FilePath & "\" & fm & YeDate & ".txt", FileTime)
      Else
        If FileTime <> ws.Cells(1, Last) Then Call オーダー更新(FilePath & "\" & fm & YeDate & ".txt", FileTime)
      End If
    Else
      MsgBox "当日引当ﾌｧｲﾙがない、もしくはﾈｯﾄﾜｰｸにつながっていません", vbCritical
      Exit Sub
    End If
  Else '23:59ｼﾞｬｽﾄ
    If ReadCheck((FilePath & "\" & fm & ToDate & ".txt"), 1) = True Then
      FileTime = FileDateTime(FilePath & "\" & fm & ToDate & ".txt")
      If flag = False Then
        Call オーダー更新(FilePath & "\" & fm & ToDate & ".txt", FileTime)
      Else
        If FileTime <> ws.Cells(1, Last) Then Call オーダー更新(FilePath & "\" & fm & ToDate & ".txt", FileTime)
      End If
    Else
      MsgBox "当日引当ﾌｧｲﾙがない、もしくはﾈｯﾄﾜｰｸにつながっていません", vbCritical
      Exit Sub
    End If
  End If
Application.ScreenUpdating = True
End Sub

Public Sub オーダー更新(fileName As String, FileTime As String)
Call オーダー処理(fileName, "VC一覧")
Dim Last As Long: Last = ThisWorkbook.Sheets("VC一覧").Cells(1, Columns.Count).End(xlToLeft).Column
ThisWorkbook.Sheets("VC一覧").Cells(1, Last + 1) = FileTime
With ThisWorkbook.Sheets("VC一覧").Cells(1, Last + 1)
  .HorizontalAlignment = xlLeft
  .VerticalAlignment = xlCenter
  .EntireColumn.AutoFit
End With
ThisWorkbook.Sheets("VC一覧").Visible = False
End Sub
Public Sub オーダー処理(fileName As String, WsName As String)
Dim FileTime As String, Count As Long, Percent As Long
Dim Buff
Dim Uni As String
Dim OldStatusBar As Boolean 'ｽﾃｰﾀｽﾊﾞｰ表示状態保存
OldStatusBar = Application.DisplayStatusBar
Application.StatusBar = True '強制表示
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.StatusBar = "ﾌｧｲﾙにｱｸｾｽを開始しました。"
Time1 = Timer
Call AddWS(WsName) 'ﾜｰｸｼｰﾄを追加(初期化)
Set ws1 = Sheets(WsName) '転記先シート名
With ws1.Cells
  '.ClearContents 'ﾃﾞｰﾀｸﾘｱ
  .Font.Size = 8
  .Font.Name = "メイリオ"
  .NumberFormatLocal = "@"
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlCenter
End With

Open fileName For Input As #1
i = 0
Do Until EOF(1) '最大値チェック
Line Input #1, Buff
i = i + 1
Loop
Count = i
Close #1
UFProgress.Tag = "1"
UFProgress.Show vbModeless
UFProgress.ProgressBar1.Min = 0
UFProgress.ProgressBar1.Max = Count

i = 1 'ｶｳﾝﾀ初期化
 Open fileName For Input As #1
 Do Until EOF(1)
     
    Line Input #1, Buff
'Buffから文字列を抽出しながら、ｺﾝﾊﾞｰﾄ処理
  'Conv(Char As String, FuncNo As Long, FormatNo As Long, PosiNo As long, j As Long, Title As String) As String
  'Char = 読み込んだ8桁の文字列
  'FuncNo = 処理番号　               1：日付変換、2：無変換
  'FormatNo = 書式番号(XY)       (X) 1～9(全9ﾊﾟﾀｰﾝ)
  '                              (Y) 1：文字列型、2：整数、3：小数点1桁、4：小数点2桁、5：小数点3桁、6：日付型
  'PosiNo = 書式位置番号             1：Center、  2：Left  3：Right
  'j = 配列内の格納番号(配列：LineBufの該当番号)
  'Title = 表題（i=1のときのみ有効）
    Sika(i, 1) = Conv("'" & Mid(Buff, 2, 8), 2, 1, 11, 1, "ｵｰﾀﾞｰ" & vbCrLf & "№") 'ｵｰﾀﾞｰ№
    Sika(i, 2) = Conv(Mid(Buff, 251, 7), 2, 81, 1, 2, "VC" & vbCrLf & "A名称") 'VC A面
    Sika(i, 3) = Conv(Mid(Buff, 645, 7), 2, 81, 1, 3, "VC" & vbCrLf & "B名称")  'VC B面
    Sika(i, 4) = Conv(Mid(Buff, 4664, 3), 2, 86, 1, 4, "輸出国") '輸出国
    Sika(i, 5) = Conv(Mid(Buff, 678, 30), 2, 86, 1, 5, "取引先")  '取引
    Sika(i, 6) = Conv(Mid(Buff, 665, 8), 2, 11, 1, 6, "合紙")  '合紙
'    Sika(i, 2) = Conv(Mid(Buff, 126, 8), 2, 86, 1, 2, "用途" & vbCrLf & "コード")    '用途コード
'    Sika(i, 3) = Conv(Mid(Buff, 619, 20), 2, 86, 1, 3, "用途名")    '用途名

'    Sika(i, 4) = Conv(Mid(Buff, 45, 5), 2, 11, 1, 4, "ｵｰﾀﾞｰ" & vbCrLf & "材質")     'ｵｰﾀﾞｰ材質
'    Sika(i, 5) = Conv(Mid(Buff, 50, 6), 2, 11, 1, 5, "ｵｰﾀﾞｰ" & vbCrLf & "調質")    'ｵｰﾀﾞｰ調質
'    Sika(i, 6) = Conv(Mid(Buff, 58, 4), 2, 15, 3, 6, "ｵｰﾀﾞｰ" & vbCrLf & "板厚")  'ｵｰﾀﾞｰ板厚
'    Sika(i, 7) = Conv(Mid(Buff, 67, 5), 2, 13, 3, 7, "ｵｰﾀﾞｰ" & vbCrLf & "板幅")    'ｵｰﾀﾞｰ板幅
'    Sika(i, 8) = Conv(Mid(Buff, 77, 6), 2, 13, 3, 8, "ｵｰﾀﾞｰ" & vbCrLf & "板丈")    'ｵｰﾀﾞｰ板丈
    
'    Sika(i, 9) = Conv(Mid(Buff, 251, 7), 2, 81, 1, 9, "VC" & vbCrLf & "A名称") 'VC A面
'    Sika(i, 10) = Conv(Mid(Buff, 645, 7), 2, 81, 1, 10, "VC" & vbCrLf & "B名称")  'VC B面
'    Sika(i, 11) = Conv(Mid(Buff, 665, 8), 2, 11, 1, 11, "合紙")  '合紙
    
'    Sika(i, 12) = Conv(Mid(Buff, 346, 11), 2, 86, 1, 12, "コメント")  'コメント
'    Sika(i, 13) = Conv(Mid(Buff, 357, 11), 2, 86, 1, 13, "コメント")  'コメント
'    Sika(i, 14) = Conv(Mid(Buff, 368, 11), 2, 86, 1, 14, "コメント")  'コメント
'
'    Sika(i, 15) = Conv(Mid(Buff, 379, 11), 2, 86, 1, 15, "コメント")  'コメント
'    Sika(i, 16) = Conv(Mid(Buff, 390, 11), 2, 86, 1, 16, "コメント")  'コメント
'    Sika(i, 17) = Conv(Mid(Buff, 401, 11), 2, 86, 1, 17, "コメント")  'コメント
    
'    Sika(i, 18) = Conv(Mid(Buff, 519, 8), 2, 86, 1, 18, "梱包" & vbCrLf & "仕様書") '仕様書
'
'    Sika(i, 19) = Conv(Mid(Buff, 678, 30), 2, 86, 1, 19, "取引先")  '得意
'    Sika(i, 20) = Conv(Mid(Buff, 708, 30), 2, 86, 1, 20, "納入先")  '納入
'    Sika(i, 21) = Conv(Mid(Buff, 738, 30), 2, 86, 1, 21, "送り先") '送り
'    Sika(i, 22) = Conv(Mid(Buff, 768, 30), 2, 86, 1, 22, "送り先住所") '送り住所
'
'    Sika(i, 23) = Conv(Mid(Buff, 4601, 6), 2, 3, 1, 23, "梱包" & vbCrLf & "単位重量")  '梱包単位重量
'    Sika(i, 24) = Conv(Mid(Buff, 4607, 8), 2, 2, 1, 24, "梱包" & vbCrLf & "単位枚数")  '梱包単位枚数
'    Sika(i, 25) = Conv(Mid(Buff, 4616, 8), 2, 3, 1, 25, "許容上限")  '許容上限
'    Sika(i, 26) = Conv(Mid(Buff, 4624, 8), 2, 3, 1, 26, "許容下限")  '許容下限
'    Sika(i, 27) = Conv(Mid(Buff, 4632, 5), 2, 2, 1, 27, "制限山数")  '制限山数
'    Sika(i, 28) = Conv(Mid(Buff, 4637, 5), 2, 3, 1, 28, "制限製品高さ")  '制限製品高さ
'    Sika(i, 29) = Conv(Mid(Buff, 4642, 5), 2, 3, 1, 29, "制限総高さ") '制限総高さ
'
'    Sika(i, 30) = Conv(Mid(Buff, 4654, 5), 2, 13, 1, 30, "外装巾")  '外装巾
'    Sika(i, 31) = Conv(Mid(Buff, 4659, 5), 2, 13, 1, 31, "外装丈")  '外装丈
'    Sika(i, 32) = Conv(Mid(Buff, 4664, 3), 2, 86, 1, 32, "輸出国") '輸出国
'
'    Sika(i, 33) = Conv(Mid(Buff, 4805, 3), 2, 11, 1, 33, "ケース開始No")  'ケース開始No
'    Sika(i, 34) = Conv(Mid(Buff, 4808, 2), 2, 11, 1, 34, "ケース数")  'ケース数
'    Sika(i, 35) = Conv(Mid(Buff, 4810, 1), 2, 11, 1, 35, "貴注No" & vbCrLf & "表示有無")  '貴注No表示有無
'    Sika(i, 36) = Conv(Mid(Buff, 4811, 5), 2, 11, 1, 36, "荷印")  '荷印
'    Sika(i, 37) = Conv(Mid(Buff, 4816, 20), 2, 86, 1, 37, "備考")  '備考
'    Sika(i, 38) = Conv(Mid(Buff, 4836, 5), 2, 13, 1, 38, "梱包指示" & vbCrLf & "ケースNo")  '梱包指示ケースNo
'    Sika(i, 39) = Conv(Mid(Buff, 4800, 5), 2, 1, 1, 39, "事務担当") '事務担当
'
'    Sika(i, 40) = Conv(Mid(Buff, 3592, 6), 2, 13, 1, 40, "検査持込" & vbCrLf & "予定量")  '検査持込み予定量
'    Sika(i, 41) = Conv(Mid(Buff, 3598, 6), 2, 13, 1, 41, "検査良品" & vbCrLf & "予定量") '検査良品予定量
'
'    Sika(i, 42) = Conv(Mid(Buff, 89, 8), 2, 2, 1, 42, "契約枚数")  '契約枚数
'    Sika(i, 43) = Conv(Mid(Buff, 97, 9), 2, 13, 1, 43, "契約重量")  '契約重量
'    Sika(i, 44) = Conv(Mid(Buff, 107, 8), 2, 13, 1, 44, "単価") '単価
'    'オーダー
'    Sika(i, 51) = Conv(Mid(Buff, 1411, 6), 2, 4, 11, 51, "オーダー厚+")  '厚公差+
'     If Sika(i, 51) = "40.404" Or Sika(i, 51) = "0" Or Sika(i, 51) = "40.4" Then
'      Sika(i, 51) = ""
'     End If
'    Sika(i, 52) = Conv(Mid(Buff, 1417, 6), 2, 4, 11, 52, "オーダー厚-")  '厚公差-
'     If Sika(i, 52) = "40.404" Or Sika(i, 52) = "0" Or Sika(i, 52) = "40.4" Then
'      Sika(i, 52) = ""
'     End If
'    Sika(i, 53) = Conv(Mid(Buff, 1447, 4), 2, 4, 11, 53, "オーダー幅+")  '幅公差+
'     If Sika(i, 53) = "40.404" Or Sika(i, 53) = "0" Or Sika(i, 53) = "40.4" Then
'      Sika(i, 53) = ""
'     End If
'    Sika(i, 54) = Conv(Mid(Buff, 1451, 4), 2, 4, 11, 54, "オーダー幅-")  '幅公差-
'     If Sika(i, 54) = "40.404" Or Sika(i, 54) = "0" Or Sika(i, 54) = "40.4" Then
'      Sika(i, 54) = ""
'     End If
'    Sika(i, 55) = Conv(Mid(Buff, 1471, 4), 2, 4, 11, 55, "オーダー丈+")  '丈公差+
'     If Sika(i, 55) = "40.404" Or Sika(i, 55) = "0" Or Sika(i, 55) = "40.4" Then
'      Sika(i, 55) = ""
'     End If
'    Sika(i, 56) = Conv(Mid(Buff, 1475, 4), 2, 4, 11, 56, "オーダー丈-")  '丈公差-
'     If Sika(i, 56) = "40.404" Or Sika(i, 56) = "0" Or Sika(i, 56) = "40.4" Then
'      Sika(i, 56) = ""
'     End If
'    '製造
'    Sika(i, 45) = Conv(Mid(Buff, 1423, 6), 2, 4, 13, 45, "製造厚+")  '厚公差+
'    Sika(i, 46) = Conv(Mid(Buff, 1429, 6), 2, 4, 13, 46, "製造厚-")  '厚公差-
'
'    Sika(i, 47) = Conv(Mid(Buff, 1455, 4), 2, 4, 13, 47, "製造幅+")  '幅公差+
'    Sika(i, 48) = Conv(Mid(Buff, 1459, 4), 2, 4, 13, 48, "製造幅-")  '幅公差-
'
'    Sika(i, 49) = Conv(Mid(Buff, 1479, 4), 2, 4, 13, 49, "製造丈+")  '丈公差+
'    Sika(i, 50) = Conv(Mid(Buff, 1483, 4), 2, 4, 13, 50, "製造丈-")  '丈公差-
    
    '製造目標
'    Sika(i, 57) = Conv(Mid(Buff, 1441, 6), 2, 13, 4, 57, "目標厚-")  '厚-
'    Sika(i, 58) = Conv(Mid(Buff, 1435, 6), 2, 13, 4, 58, "目標厚+")  '厚+
'
'    Sika(i, 59) = Conv(Mid(Buff, 1467, 4), 2, 13, 4, 59, "目標巾-")  '幅-
'    Sika(i, 60) = Conv(Mid(Buff, 1463, 4), 2, 13, 4, 60, "目標巾+")  '幅+
'
'    Sika(i, 61) = Conv(Mid(Buff, 1491, 4), 2, 13, 4, 61, "目標丈-")  '丈-
'    Sika(i, 62) = Conv(Mid(Buff, 1487, 4), 2, 13, 4, 62, "目標丈+")  '丈+
    
'    Cells(i, 36) = Mid(Buff, 2487, 4)    '検査持込み
'    Cells(i, 37) = Mid(Buff, 2490, 4)   '検査良品予定
'
'
'    Cells(i, 38) = Mid(Buff, 3081, 4)   '梱包単位重量
'    Cells(i, 39) = Mid(Buff, 3085, 4)   '梱包単位枚数
'    Cells(i, 40) = Mid(Buff, 3090, 4)   '許容範囲+
'    Cells(i, 41) = Mid(Buff, 3094, 4)   '許容範囲-
'
'    Cells(i, 42) = Mid(Buff, 3098, 3)   '山数
'
'    Cells(i, 43) = Mid(Buff, 3101, 3)   '梱包高さ
'    Cells(i, 44) = Mid(Buff, 3104, 3)   '総高さ
'
'    Cells(i, 45) = Mid(Buff, 3112, 3)   '許容外装巾
'    Cells(i, 46) = Mid(Buff, 3115, 3)   '許容外装丈
'
'    Cells(i, 47) = Mid(Buff, 3094, 3)   '輸出先国名
'
'    Cells(i, 48) = Mid(Buff, 3148, 5)   '試験番号
'
'    Cells(i, 49) = Mid(Buff, 1105, 2)   '歪巾公差
'    Cells(i, 50) = Mid(Buff, 1103, 2)   '％
'    Cells(i, 51) = Mid(Buff, 1109, 2)   '歪丈公差
'    Cells(i, 52) = Mid(Buff, 1107, 2)   '％
'
'    Cells(i, 53) = Mid(Buff, 1168, 8)   'ラテラルボー
'    Cells(i, 54) = Mid(Buff, 1176, 8)   '直角度

'    If i >= 2 Then
'      Call データ変換 'フラグの変換や空欄処理
'    End If
'
'    If (i Mod 12) = 0 Then
'       Application.StatusBar = FileName & " 　" & i - 1 & "行目　読取中"
'    End If
    i = i + 1
    If UFProgress.ProgressBar1.Min < i And UFProgress.ProgressBar1.Max >= i Then
      Percent = CInt(i / Count * 100)
      UFProgress.Label1.Caption = fileName & " 　" & i & "行　展開中"
      UFProgress.Label2.Caption = Percent & "%完了"
      Time1 = Timer - Time1
      UFProgress.Label3.Caption = "実行時間" & Format(Time1, "0.0") & "秒 "
      UFProgress.ProgressBar1.Value = i
      DoEvents
    End If
   Loop

 Close #1
' Application.StatusBar = "EXCELｼｰﾄへﾃﾞｰﾀを書込中"
 Range(ws1.Cells(1, 1), ws1.Cells(i, UBound(Sika, 2))) = Sika
 ws1.Cells.EntireColumn.AutoFit
 UFProgress.Tag = "1"
 Unload UFProgress
 Erase Sika
 
'  ws1.Activate
'  ActiveWindow.ScrollRow = 1
'  ActiveWindow.ScrollColumn = 1
'  Rows("2:2").Select
'  ActiveWindow.FreezePanes = True '固定
'  Application.ScreenUpdating = True
'  Application.Calculation = xlCalculationAutomatic
'  Application.StatusBar = OldStatusBar '戻す
'
'  With ActiveWorkbook.Sheets("VC一覧").Tab
'    .Color = &HC0C000
'    .TintAndShade = 0
'  End With

'  Time1 = Timer - Time1
'  Set ws1 = Nothing  'ｵﾌﾞｼﾞｪｸﾄの解放
'  MsgBox "更新が完了しました。" & vbCrLf & "実行時間は" & Format(Time1, "0.0") & "秒 ", vbInformation
  
End Sub
'表題・書式設定、日付ﾃﾞｰﾀ変換等の処理(データの基本処理)
Private Function Conv(Char As String, FuncNo As Long, FormatNo As Long, PosiNo As Long, j As Long, title As String) As String
'FuncNo = 処理番号　               1：日付変換、2：無変換
'FormatNo = 書式番号(XY)       (X) 1～9(全9ﾊﾟﾀｰﾝ)
'                              (Y) 1：文字列型、2：整数、3：小数点1桁、4：小数点2桁、5：小数点3桁、6：日付型
'PosiNo = 書式位置番号             1：Center、  2：Left  3：Right
'Char = 読み込んだ8桁の文字列
'j = 配列内の格納番号
'Title = 表題
'Conv = 出力するデータ
Select Case FuncNo
  '日付格納データの展開処理（8桁のﾃﾞｰﾀを日付の表示の形に変換する）
  Case Is = 1 '日付
    If val(Char) = 0 Then
      Conv = ""
    Else
      If Mid(Char, 1, 2) = "98" Or Mid(Char, 1, 2) = "99" Then  '19**年処理
        Conv = "19" & Mid(Char, 1, 2) & "/" & Mid(Char, 3, 2) & "/" & Mid(Char, 5, 2)
      Else  '20**年処理
        Conv = "20" & Mid(Char, 1, 2) & "/" & Mid(Char, 3, 2) & "/" & Mid(Char, 5, 2)
      End If

      If j >= 175 Then '進度情報(175列以降)は時刻データ追加
        Conv = Conv & " " & Mid(Char, 10, 2) & ":" & Mid(Char, 12, 2) & ":" & Mid(Char, 14, 2)
      End If
    End If
  Case Is = 2
    Conv = Trim(Char)
End Select

'表題・書式設定
If i = 1 Then

   Conv = title '表題

  With Range(ws1.Cells(1, j), ws1.Cells(1, j))
    Select Case Left(FormatNo, 1)
      Case Is = 1: .Interior.Color = RGB(250, 240, 230)
      Case Is = 2: .Interior.Color = RGB(245, 230, 250)
      Case Is = 3: .Interior.Color = RGB(230, 240, 220)
      Case Is = 4: .Interior.Color = RGB(240, 250, 210)
      Case Is = 5: .Interior.Color = RGB(235, 200, 245)
      Case Is = 6: .Interior.Color = RGB(250, 250, 230)
      Case Is = 7: .Interior.Color = RGB(255, 240, 220)
      Case Is = 8: .Interior.Color = RGB(235, 230, 250)
      Case Is = 9: .Interior.Color = RGB(240, 250, 245)
    End Select
  End With
   
ElseIf i = 2084 Then '最終行処理時のみ一括書式設定 '最終行判定難しい。
  With Range(ws1.Cells(2, j), ws1.Cells(i, j))
    Select Case Right(FormatNo, 1)
      Case Is = 1: .NumberFormatLocal = "@" '文字列
      Case Is = 2: .NumberFormatLocal = "0" '整数
      Case Is = 3: .NumberFormatLocal = "0.0" '1桁
      Case Is = 4: .NumberFormatLocal = "0.00" '2桁
      Case Is = 5: .NumberFormatLocal = "0.000" '3桁
      Case Is = 6: .NumberFormatLocal = "yyyy/mm/dd" '日付
      Case Is = 7: .NumberFormatLocal = "0.00000"  '5桁（分析結果用）
    End Select
    Select Case PosiNo
      Case Is = 1: .HorizontalAlignment = xlCenter 'センター合わせ
      Case Is = 2: .HorizontalAlignment = xlLeft '左寄せ
      Case Is = 3: .HorizontalAlignment = xlRight '右寄せ
    End Select
  End With
End If
End Function
'ヒットしたロットの配列を作成
Public Function GetLotInfo(SheetName As String, lottemp As String) As Variant
If WsCheck(SheetName) = False Then Exit Function
Dim HistoryLot As Worksheet: Set HistoryLot = ThisWorkbook.Sheets(SheetName)
Dim EndR As Long:            EndR = UR(HistoryLot, 1)
Dim EndC As Long:            EndC = UC(HistoryLot, 1)
Dim HistoryRan As Range:     Set HistoryRan = Range(HistoryLot.Cells(1, 1), HistoryLot.Cells(EndR, EndC))
Dim HisArry():               HisArry() = HistoryRan

Dim C_納入先 As Long, C_用途コード As Long, C_ロットNo As Long, C_材質 As Long, C_得意先 As Long, _
C_調質 As Long, C_厚 As Long, C_幅 As Long, C_丈 As Long, i As Long, a As Long, C_検査No As Long
For i = LBound(HisArry, 2) To UBound(HisArry, 2)
  Select Case HisArry(1, i)
  Case "ﾛｯﾄNo":        C_ロットNo = i
  Case "納入先":       C_納入先 = i
  
  Case "得意先":       C_得意先 = i '2024.8.5
  
  Case "用途ｺｰﾄﾞ":     C_用途コード = i
  Case "材質":         C_材質 = i
  Case "調質":         C_調質 = i
  Case "厚":           C_厚 = i
  Case "巾":           C_幅 = i
  Case "丈":           C_丈 = i
  Case "検査No":       C_検査No = i
  End Select
Next i
If C_ロットNo = 0 Or C_納入先 = 0 Or C_得意先 = 0 Or C_用途コード = 0 Or C_材質 = 0 Or C_調質 = 0 Or C_厚 = 0 Or C_幅 = 0 Or C_丈 = 0 Or C_検査No = 0 Then Exit Function

Dim DestArry() As Variant ' コピー先の配列
Dim destRowIndex As Long: destRowIndex = 0

For i = LBound(HisArry, 1) To UBound(HisArry, 1)
  If HisArry(i, C_ロットNo) = lottemp Then
    destRowIndex = destRowIndex + 1
  End If
Next i

If destRowIndex = 0 Then Exit Function

ReDim DestArry(1 To destRowIndex, 1 To 10)
destRowIndex = 0
'2024.4.22_修正　ループが入れ子でおかしかった
For i = LBound(HisArry, 1) To UBound(HisArry, 1)
  If HisArry(i, C_ロットNo) = lottemp Then
    destRowIndex = destRowIndex + 1
    DestArry(destRowIndex, 1) = HisArry(i, C_納入先)
    DestArry(destRowIndex, 2) = HisArry(i, C_得意先) '2024.8.5
    DestArry(destRowIndex, 3) = HisArry(i, C_用途コード)
    DestArry(destRowIndex, 4) = HisArry(i, C_材質)
    DestArry(destRowIndex, 5) = HisArry(i, C_調質)
    DestArry(destRowIndex, 6) = HisArry(i, C_厚)
    DestArry(destRowIndex, 7) = HisArry(i, C_幅)
    DestArry(destRowIndex, 8) = HisArry(i, C_丈)
    DestArry(destRowIndex, 9) = HisArry(i, C_ロットNo)
    DestArry(destRowIndex, 10) = HisArry(i, C_検査No)
  End If
Next i
'納入、取引、用途コード、材質、調質、厚、幅、丈、ロットNo、検査No、ｵｰﾀﾞｰNo、包装仕様NO
GetLotInfo = DestArry

End Function

