Attribute VB_Name = "【UFKen_13】Mitsubishi"
Option Explicit
Public Sub OpenMitubisi(Temp As String)
Dim Worker, sh As Worksheet
Dim NewWork As String

Select Case WorkeDayMax
Case "1": Worker = Split(UF1.T1, "_") '複数人の登録があった場合後の登録者を記録
Case "2": Worker = Split(UF1.T2, "_")
Case "3": Worker = Split(UF1.T3, "_")
End Select

UFMitsubishi.PackNo_on_off = "off"
Select Case Temp
Case "全数"
  UFMitsubishi.MLP.P1.Visible = True
  UFMitsubishi.MLP.P_H494.Visible = False
  UFMitsubishi.MLP.Value = 0
  
  UFMitsubishi.WorkFlag.Caption = "P1"
  
  If WsCheck("全ヵ所検査") Then
    Set sh = ThisWorkbook.Worksheets("全ヵ所検査")
    sh.Activate
    '検査員
    If IsEmpty(Worker) = False Then
      If sh.Cells(3, 8) <> "" Then NewWork = Worker(UBound(Worker)) & " " & sh.Cells(3, 8)
      If sh.Cells(3, 8) = "" Then NewWork = Worker(UBound(Worker))
      
      NewWork = RemoveDuplicates(NewWork) '重複削除
      UFMitsubishi.Kname = NewWork
    End If
    
    UFMitsubishi.LotBox = sh.Cells(2, 4)
    UFMitsubishi.KenNo = sh.Cells(2, 8)
  '  UFMitsubishi.Kname = SH.Cells(3, 8)
    UFMitsubishi.Choice = sh.Cells(1, 22) '作業選択名_2021.09.15
    UFMitsubishi.YoutoNo = sh.Cells(4, 8)
    UFMitsubishi.高さ計算式.Caption = sh.Cells(1, 30) '2023.8.28
    '2023.5.16_特別用途コード
    If sh.Cells(2, 4).Font.Color = RGB(255, 0, 0) Then UFMitsubishi.H494 = True
  End If
Case "H494"
  UFMitsubishi.MLP.P1.Visible = False
  UFMitsubishi.MLP.P_H494.Visible = True
  UFMitsubishi.MLP.Value = 1
  
  UFMitsubishi.WorkFlag.Caption = "P_H494"
 
  If WsCheck("H494検査") Then
    Set sh = ThisWorkbook.Worksheets("H494検査")
    sh.Activate
    
    '検査員
    If IsEmpty(Worker) = False Then
      If sh.Cells(3, 8) <> "" Then NewWork = Worker(UBound(Worker)) & " " & sh.Cells(3, 8)
      If sh.Cells(3, 8) = "" Then NewWork = Worker(UBound(Worker))
      
      NewWork = RemoveDuplicates(NewWork) '重複削除
      UFMitsubishi.Kname = NewWork
    End If
    
    UFMitsubishi.H494_Lot = sh.Cells(2, 4)
    UFMitsubishi.H494_KenNo = sh.Cells(2, 8)
  '  UFMitsubishi.H494_Kname = SH.Cells(3, 8)
    UFMitsubishi.H494_Youto = sh.Cells(4, 8)
    UFMitsubishi.高さ計算式.Caption = sh.Cells(1, 30) '2023.8.28
  End If
End Select

UFMitsubishi.LineName.Caption = UF1.LineName.Caption
Load UFMitsubishi

UFMitsubishi.Show vbModeless
UFMitsubishi.PackNo_on_off = "on"
End Sub
Function RemoveDuplicates(tMP As String, Optional n As String = "") As String '重複削除
Dim Words() As String
Dim UniqueWords As String
Dim Word As Variant
If n <> "" Then
  Words = Split(tMP, n) 'nで区切られた単語を配列に分割
Else
  Words = Split(tMP, " ") '半角スペースで区切られた単語を配列に分割
  n = " "
End If
' 重複を削除して新しい文字列を構築
For Each Word In Words
    If InStr(1, UniqueWords, Word & n, vbTextCompare) = 0 Then
        UniqueWords = UniqueWords & Word & n 'まだ追加されていない単語の場合は新しい文字列に追加
    End If
Next Word
RemoveDuplicates = Trim(UniqueWords) ' 末尾のスペースを削除して結果を返す
End Function
Public Sub 全ヵ所用コピー(Optional Temp As String = "全ヵ所検査") '引数追加_2023.10.16
Dim sh As String: sh = Temp
Dim ws As Worksheet, flag As Boolean
Application.ScreenUpdating = False
For Each ws In ThisWorkbook.Worksheets
  If ws.Name = sh Then flag = True
Next ws
If flag = True Then
  Application.DisplayAlerts = False
  ThisWorkbook.Sheets(sh).Delete
  Application.DisplayAlerts = True
End If
Select Case Temp
Case "全ヵ所検査"
  ThisWorkbook.Sheets("全ヵ所検査用原紙").Visible = True
  ThisWorkbook.Sheets("全ヵ所検査用原紙").Copy before:=Worksheets(1)
Case "H494検査"
  ThisWorkbook.Sheets("H494_原紙").Visible = True
  ThisWorkbook.Sheets("H494_原紙").Copy before:=Worksheets(1)
End Select
  
Set ws = ActiveSheet ' コピーしたシートがアクティブ
' そのシートに対して編集
ws.Name = sh
With ws.Tab
    .Color = 16764108
    .TintAndShade = 0
End With
Select Case Temp
Case "全ヵ所検査"
  ThisWorkbook.Sheets("全ヵ所検査用原紙").Visible = False
  ActiveWindow.Zoom = 110
Case "H494検査"
  ThisWorkbook.Sheets("H494_原紙").Visible = False
  ActiveWindow.Zoom = 110
End Select
  
Application.ScreenUpdating = True
End Sub
Public Sub 全ヵ所用最大最少(Optional Temp As String = "全ヵ所検査") 'H494引数追加_2023.10.18
Select Case UFMitsubishi.WorkFlag.Caption
Case "P1"
  If UFMitsubishi.MINText = "" Then Exit Sub
  If UFMitsubishi.MAXText = "" Then Exit Sub
Case "P_H494"
  If UFMitsubishi.H494_MINText = "" Then Exit Sub
  If UFMitsubishi.H494_MAXText = "" Then Exit Sub
End Select

Dim i As Long, a As Long, n As Long, k As Long, Erow As Long, Ecol As Long
Dim C2 As Range, C4 As Range, C6 As Range, C8 As Range
Dim C11 As Range, C13 As Range, C15 As Range, C17 As Range, C19 As Range
Dim MAXatu As Single, MINatu As Single
Dim MAXhaba As Single, MINhaba As Single
Dim MAXtake As Single, MINtake As Single
Dim shName As String

Select Case UFMitsubishi.WorkFlag.Caption
Case "P1"
  n = val(UFMitsubishi.MINText) + 6
  k = val(UFMitsubishi.MAXText) + 6
  Erow = 36
  Ecol = 13
Case "P_H494"
  n = val(UFMitsubishi.H494_MINText) + 6
  k = val(UFMitsubishi.H494_MAXText) + 6
  Erow = 16
  Ecol = 15
End Select

With ThisWorkbook.Sheets(Temp) '列毎の設定
  Set C2 = .Range(.Cells(n, 2), (.Cells(k, 2)))
  Set C4 = .Range(.Cells(n, 4), (.Cells(k, 4)))
  Set C6 = .Range(.Cells(n, 6), (.Cells(k, 6)))
  Set C8 = .Range(.Cells(n, 8), (.Cells(k, 8)))
  Set C11 = .Range(.Cells(n, 11), (.Cells(k, 11)))
  Set C13 = .Range(.Cells(n, 13), (.Cells(k, 13)))
  Set C15 = .Range(.Cells(n, 15), (.Cells(k, 15)))
  Set C17 = .Range(.Cells(n, 17), (.Cells(k, 17)))
  Set C19 = .Range(.Cells(n, 19), (.Cells(k, 19)))
End With
With Application.WorksheetFunction
'4列での最大最小 (厚)
    MAXatu = .Max(.Max(C2), .Max(C4), .Max(C6), .Max(C8))
    MINatu = .Min(.Min(C2), .Min(C4), .Min(C6), .Min(C8))
    '4列内に空白がある場合の処理(最小が測定していない箇所になってしまう)
    If MINatu = 0 Then '2022.01.27**厚2か所検査の場合数字を拾わない対策
      Dim AtuVar(1 To 4) As Single
      For i = 1 To 4
        Select Case i
        Case 1
          If .Min(C2) = 0 Then
            AtuVar(1) = 0
          Else
            AtuVar(1) = .Min(C2)
          End If
        Case 2
          If .Min(C4) = 0 Then
            AtuVar(2) = 0
          Else
            AtuVar(2) = .Min(C4)
          End If
        Case 6
          If .Min(C6) = 0 Then
            AtuVar(3) = 0
          Else
            AtuVar(3) = .Min(C6)
          End If
        Case 8
          If .Min(C8) = 0 Then
            AtuVar(4) = 0
          Else
            AtuVar(4) = .Min(C8)
          End If
        End Select
      Next i
     MINatu = WorksheetFunction.Small((AtuVar), 2) '一番小さいのは0なので次に小さい数字
     If MINatu = 0 Then MINatu = WorksheetFunction.Small((AtuVar), 3) '2か所検査の場合おそらくここまで
'     Stop
    End If
  Select Case UFMitsubishi.WorkFlag.Caption
  Case "P1" '2列での最大最小 (巾)(丈)
    MAXhaba = .Max(.Max(C11), .Max(C13))
    MINhaba = .Min(.Min(C11), .Min(C13))
    MAXtake = .Max(.Max(C15), .Max(C17))
    MINtake = .Min(.Min(C15), .Min(C17))
  Case "P_H494" '3列での最大最小 (巾)
    MAXhaba = .Max(.Max(C11), .Max(C13), .Max(C15))
    MINhaba = .Min(.Min(C11), .Min(C13), .Min(C15))
    MAXtake = .Max(.Max(C17), .Max(C19))
    MINtake = .Min(.Min(C17), .Min(C19))
  End Select
End With

With ThisWorkbook.Sheets(Temp)
.Unprotect 'シート保護解除
'リセット**************歪公差外流出によりレイアウト変更_23024.1.11
Dim TempA As Long, TempH As Long, TempT As Long

Select Case UFMitsubishi.WorkFlag.Caption
Case "P1"
  .Cells(3, 23) = ""
  .Cells(4, 23) = ""
  TempA = 23
  .Cells(3, 25) = ""
  .Cells(4, 25) = ""
  TempH = 25
  .Cells(3, 27) = ""
  .Cells(4, 27) = ""
  TempT = 27
Case "P_H494"
  .Cells(3, 22) = ""
  .Cells(4, 22) = ""
  TempA = 22
  .Cells(3, 23) = ""
  .Cells(4, 23) = ""
  TempH = 23
  .Cells(3, 24) = ""
  .Cells(4, 24) = ""
  TempT = 24
End Select
.Range(.Cells(7, 2), .Cells(Erow, 22)).Font.Underline = xlUnderlineStyleNone
'Range("B7:U36").Font.Underline = xlUnderlineStyleNone

'指定の数字にアンダーバー
 '-----------------------------------------------------厚
    For i = n To k
      For a = 2 To 8 Step 2
        If .Cells(i, a) = MAXatu Then '同じ数字の時
           .Cells(i, a).Font.Underline = xlUnderlineStyleSingle
           .Cells(3, TempA) = MAXatu
        Exit For
        End If
      Next a
    Next i
    For i = n To k
      For a = 2 To 8 Step 2
        If .Cells(i, a) = MINatu Then '同じ数字の時
           .Cells(i, a).Font.Underline = xlUnderlineStyleSingle
           .Cells(4, TempA) = MINatu
        Exit For
        End If
      Next a
    Next i
 '-----------------------------------------------------巾
    For i = n To k
      For a = 11 To Ecol Step 2
        If .Cells(i, a) = MAXhaba Then '同じ数字の時
           .Cells(i, a).Font.Underline = xlUnderlineStyleSingle
           .Cells(3, TempH) = MAXhaba
        Exit For
        End If
      Next a
    Next i
    For i = n To k
      For a = 11 To Ecol Step 2
        If .Cells(i, a) = MINhaba Then '同じ数字の時
           .Cells(i, a).Font.Underline = xlUnderlineStyleSingle
           .Cells(4, TempH) = MINhaba
        Exit For
        End If
      Next a
    Next i
 '-----------------------------------------------------丈
 Dim colsmall As Long: colsmall = Ecol + 2
 Dim colBig As Long:   colBig = Ecol + 4
    For i = n To k
      For a = colsmall To colBig Step 2
        If .Cells(i, a) = MAXtake Then '同じ数字の時
           .Cells(i, a).Font.Underline = xlUnderlineStyleSingle
           .Cells(3, TempT) = MAXtake
        Exit For
        End If
      Next a
    Next i
    For i = n To k
      For a = colsmall To colBig Step 2
        If .Cells(i, a) = MINtake Then '同じ数字の時
           .Cells(i, a).Font.Underline = xlUnderlineStyleSingle
           .Cells(4, TempT) = MINtake
        Exit For
        End If
      Next a
    Next i
.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'シート保護
End With

End Sub

Public Sub JustinCase_ZnSave(Optional Temp As String = "全ヵ所検査") '2019.11.23

If WsCheck(Temp) = False Then Exit Sub
Dim ws As Worksheet, flag As Boolean
Dim sh As Worksheet
Dim fCount As Integer, Lcount As Integer
Dim DayLotName As String, MonthName As String, Lotname As String
Dim SetPath As String, MonthPath As String
Dim wkArray()
Dim idx, Cnt As Integer
Dim n As Long
Dim BK As Workbook

Application.ScreenUpdating = False
ThisWorkbook.Activate
'************************************************保存シートの初期化
    Call AddWS(Temp & "保存")
    With ThisWorkbook.Sheets(Temp & "保存").Tab
      .Color = &H99FFFF
    End With
    ActiveWindow.Zoom = 90
  '------------------------------------------------コピー処理 保存閲覧用
          ThisWorkbook.Sheets(Temp).Unprotect
          ThisWorkbook.Sheets(Temp).Cells.Copy
          Worksheets(Temp & "保存").Cells.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          Application.CutCopyMode = False
          With Worksheets(Temp & "保存").Cells.Interior '色塗りをなくす
              .Pattern = xlNone
              .TintAndShade = 0
              .PatternTintAndShade = 0
          End With
          With ThisWorkbook.Sheets(Temp & "保存")
            .Cells(1, 3) = "～"
            .Cells(1, 4) = Format$(Now, "yyyy/mm/dd h:mm")
            .Range(.Cells(1, 3), .Cells(1, 4)).Font.Size = 5
            .Cells.EntireColumn.AutoFit
           If UFMitsubishi.WorkFlag.Caption = "P1" Then .Range("B:B,K:K,X:X").ColumnWidth = 7.7
           .Cells(1, 6) = "検査員：" & UFMitsubishi.Kname
           .Cells(1, 6).Font.Size = 7
           .Cells.Locked = True
           .Cells.FormulaHidden = False
           .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'シート保護
          End With
     '------------------------------------------------
      '必要データのみのシートをｻｰﾊﾞｰ保存する
      ThisWorkbook.Sheets(Temp & "保存").Visible = True
      For idx = 1 To ThisWorkbook.Worksheets.Count
        If ThisWorkbook.Worksheets(idx).Visible Then
         If (ThisWorkbook.Worksheets(idx).Name) = Temp & "保存" Or (ThisWorkbook.Worksheets(idx).Name) = "全ヵ所検査" Then
            Cnt = Cnt + 1
            ReDim Preserve wkArray(1 To Cnt)
            wkArray(Cnt) = ThisWorkbook.Worksheets(idx).Name
         End If
        End If
      Next idx
      If Cnt = 0 Then
      Exit Sub
      Else
      ThisWorkbook.Worksheets(wkArray).Copy
      End If
     '------------------------------------------------2019.08.28 アクティブがうつって保存できないエラーが出た
      Set BK = ActiveWorkbook
     '------------------------------------------------
    Application.DisplayAlerts = False
      '************************************月毎にフォルダを作成してその中に格納する2018.12.19
        MonthName = Format(Date, "yyyy.mm")
'        MonthPath = "C:\Users\NLM Administrator\Desktop\JustinCaseSave\" & MonthName '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\キャスク 履歴\" & MonthName 'テスト用パスの定義
        Select Case UFMitsubishi.LineName
        Case "Test":    MonthPath = HistoryPath & "Test\Test_履歴\JustinCaseSave"  '"C:\Users\NLM Administrator\Desktop\保存用ﾌｫﾙﾀﾞ\JustinCaseSave\" & MonthName
        Case Else:      MonthPath = HistoryPath & "JustinCaseSave"
        End Select

        Lotname = ThisWorkbook.Sheets(Temp & "保存").Cells(2, 4) 'ロットNO
        DayLotName = Format(Date, "yyyymmdd") & "_" & Sheets(Temp & "保存").Cells(2, 4) '日付とロットNOのフォルダを作成　2019.08.22
        SetPath = MonthPath & "\" & DayLotName
        
        Call CreateFolder(SetPath) 'ﾌｫﾙﾀﾞを作成
        
        Dim FilePath As String:     FilePath = SetPath & "\" & DayLotName & ".xlsx"
        
        If Dir(FilePath, vbDirectory) = "" Then
            '2019.08.28 , FileFormat:=xlOpenXMLWorkbook
            BK.SaveAs (FilePath), FileFormat:=xlOpenXMLWorkbook, Password:=Lotname & "." & Format(Date, "yyyy.mm.dd")    'pass LotNo.西暦.月.日
            SetAttr (FilePath), vbReadOnly
        Else
          For fCount = 2 To 999
              If Dir(SetPath & "\" & DayLotName & "." & Right("00" & CStr(fCount), 3) & ".xlsx", vbDirectory) = "" Then '同じ場合副番を付けて999まで保存する
              BK.SaveAs (SetPath & "\" & DayLotName & "." & Right("00" & CStr(fCount), 3) & ".xlsx"), FileFormat:=xlOpenXMLWorkbook, Password:=Lotname & "." & Format(Date, "yyyy.mm.dd")    'pass LotNo.西暦.月.日
              SetAttr (SetPath & "\" & DayLotName & "." & Right("00" & CStr(fCount), 3) & ".xlsx"), vbReadOnly
          Exit For
          End If
        If fCount = 999 Then
            MsgBox ("連番をこれ以上付けることが出来ません。"), vbCritical
        End
        End If
        Next
        End If
  '-------------------------------------------------------------------------------------------
    BK.Close
    ThisWorkbook.Sheets(Temp).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True  '保護設定
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
Public Sub 全ヵ所専用保存(Optional Temp As String = "全ヵ所検査") '2019.08.15
If WsCheck(Temp) = False Then
  MsgBox Temp & "がありません", vbCritical
  Exit Sub
End If
Dim ws As Worksheet, flag As Boolean
Dim sh As Worksheet, Kname As String, FileN As String
Dim fCount As Integer, Lcount As Integer
Dim DayLotName As String, MonthName As String, Lotname As String
Dim SetPath As String, MonthPath As String
Dim Lname As String: Lname = UFMitsubishi.LineName.Caption
Dim wkArray()
Dim idx, Cnt As Long, n As Long
Dim BK As Workbook
Application.Visible = True
Application.ScreenUpdating = False
Call AddWS(Temp & "保存")
Dim ZnSh As Worksheet: Set ZnSh = ThisWorkbook.Worksheets(Temp & "保存")
Dim KeSh As Worksheet: Set KeSh = ThisWorkbook.Worksheets(Temp)
  '------------------------------------------------コピー処理 保存閲覧用
  
  Select Case UFMitsubishi.WorkFlag.Caption
  Case "P1"
    Kname = UFMitsubishi.Kname
  Case "P_H494"
    Kname = UFMitsubishi.H494_Kname
  End Select
  
  
  KeSh.Unprotect
  KeSh.Cells.Copy
  ZnSh.Cells.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  With ZnSh.Cells.Interior '色塗りをなくす
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  ZnSh.Cells(1, 3) = "～"
  ZnSh.Cells(1, 4) = Format$(Now, "yyyy/mm/dd h:mm")
  Range(ZnSh.Cells(1, 3), ZnSh.Cells(1, 4)).Font.Size = 5
  ZnSh.Cells.EntireColumn.AutoFit
  If UFMitsubishi.WorkFlag.Caption = "P1" Then ZnSh.Range("B:B,K:K,X:X").ColumnWidth = 7.7
  If UFMitsubishi.WorkFlag.Caption = "P_H494" Then ZnSh.Range("B:B,K:K,Y:Y").ColumnWidth = 7.7
  ZnSh.Cells(1, 6) = "検査員：" & Kname
  ZnSh.Cells(1, 6).Font.Size = 7
  ZnSh.Cells.Locked = True
  ZnSh.Cells.FormulaHidden = False
'  ZnSh.Protect Password:=""
  ZnSh.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'シート保護
  '------------------------------------------------
  '必要データのみのシートをｻｰﾊﾞｰ保存する
  Call DisplaySheets
  With ThisWorkbook
    For idx = 1 To .Worksheets.Count
      If (.Worksheets(idx).Name) = Temp & "保存" Or (.Worksheets(idx).Name) = "ｺﾋﾟ-用" _
      Or (.Worksheets(idx).Name) = "転送済" Or (.Worksheets(idx).Name) = "5N" Or (.Worksheets(idx).Name) = "不良要因印刷" & "_" & Lname Then
        Cnt = Cnt + 1
        ReDim Preserve wkArray(1 To Cnt)
        wkArray(Cnt) = .Worksheets(idx).Name
      End If
    Next idx
    If Cnt = 0 Then
    Exit Sub
    Else
    .Worksheets(wkArray).Copy
    End If
  End With
  Call HiddenSheets
  ThisWorkbook.Sheets(Temp).Visible = True
  '------------------------------------------------2019.08.28 アクティブがうつって保存できないエラーが出た
  Set BK = ActiveWorkbook
  '------------------------------------------------
  Application.DisplayAlerts = False
    '月毎にフォルダを作成してその中に格納する2018.12.19 '名前を付けて保存　名前が被れば001～数字を付ける
      MonthName = Format(Date, "yyyy.mm")
      MonthPath = "C:\検査履歴" & "\" & MonthName '-----------------------ﾛｰｶﾙ保存_2021.01.25
      Lotname = ZnSh.Cells(2, 4) 'ロットNO
      DayLotName = Format(Date, "yyyymmdd") & "_" & ZnSh.Cells(2, 4) '日付とロットNOのフォルダを作成　2019.08.22
      SetPath = MonthPath & "\" & DayLotName
      
      Call CreateFolder(SetPath) 'ﾌｫﾙﾀﾞを作成
      
      Dim FilePath As String:     FilePath = SetPath & "\" & DayLotName & ".xlsx"
      
      If Dir(FilePath, vbDirectory) = "" Then
          '2019.08.28 , FileFormat:=xlOpenXMLWorkbook
          BK.SaveAs (FilePath), FileFormat:=xlOpenXMLWorkbook, Password:=Lotname & "." & Format(Date, "yyyy.mm.dd")    'pass LotNo.西暦.月.日
          SetAttr (FilePath), vbReadOnly
      Else
        For fCount = 2 To 999
          If Dir(SetPath & "\" & DayLotName & "." & Right("00" & CStr(fCount), 3) & ".xlsx", vbDirectory) = "" Then '同じ場合副番を付けて999まで保存する
            BK.SaveAs (SetPath & "\" & DayLotName & "." & Right("00" & CStr(fCount), 3) & ".xlsx"), FileFormat:=xlOpenXMLWorkbook, Password:=Lotname & "." & Format(Date, "yyyy.mm.dd")    'pass LotNo.西暦.月.日
            SetAttr (SetPath & "\" & DayLotName & "." & Right("00" & CStr(fCount), 3) & ".xlsx"), vbReadOnly
            Exit For
          End If
          If fCount = 999 Then
            MsgBox ("連番をこれ以上付けることが出来ません。"), vbCritical
            End
          End If
        Next
      End If
  '品証技術用に別に保存する　\\nlmsrvngy03\NPS\全社研修会トレーナー教育\第36回NPSトレーナー教育（名古屋）2017年\梱包実績の見える化\
  '-------------------------------------------------------------------------------------------
    If UFMitsubishi.Inflag = "" Then '検査印刷をおしたときのﾌﾗｸﾞ
    ElseIf UFMitsubishi.Inflag = "flagon" Then '三菱印刷をおしたときのﾌﾗｸﾞ
        Application.PrintCommunication = False
          ThisWorkbook.Sheets("転送済").PageSetup.PrintArea = "$B$2:$R$20" 'A3印刷範囲 2020.01.09 列+1範囲修正
        With ThisWorkbook.Sheets("転送済").PageSetup
          .LeftHeader = ""
          .CenterHeader = ""
          .RightHeader = ""
          .LeftFooter = ""
          .CenterFooter = ""
          .LeftFooter = ""
          .CenterFooter = ""
          .RightFooter = ""
        End With
        With ThisWorkbook.Sheets("転送済").PageSetup '"&""Meiryo UI""03" &
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .LeftMargin = Application.InchesToPoints(0#)
        .RightMargin = Application.InchesToPoints(0#)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0#)
        .CenterHorizontally = True
        .CenterVertically = True
        .PaperSize = xlPaperA3 'サイズ
        .Orientation = xlLandscape  '縦xlPortrait　横 xllandscape
        .FitToPagesWide = 1 '1ページで納める
        .FitToPagesTall = 1
        End With
        Application.PrintCommunication = True
        MonthName = Format(Date, "yyyy.mm")
        Select Case True
        Case UF1.Test:  MonthPath = "C:\Users\NLM Administrator\Desktop\保存用ﾌｫﾙﾀﾞ\三菱受託検査_履歴\" & MonthName
        Case Else:      MonthPath = "\\nlmsrvngy03\NPS\全社研修会トレーナー教育\第36回NPSトレーナー教育（名古屋）2017年\梱包実績の見える化\三菱受託検査_履歴\" & MonthName
        End Select
        '"C:\Users\NLM Administrator\Desktop\三菱受託検査 履歴\" & MonthName '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\三菱受託検査 履歴\" & MonthName 'テスト用パスの定義
        If Dir(MonthPath, vbDirectory) = "" Then 'vbDirectory必要　指定すると引数setPathで指定したフォルダがあればフォルダ名を返し,なければ長さ0の文字列「""」を返します。
            MkDir MonthPath '月毎のフォルダを作成
        Else
        End If
      '************************************名前を付けて保存　名前が被れば001～数字を付ける
        Lotname = ZnSh.Cells(2, 4) 'ロットNO
        DayLotName = Format(Date, "yyyymmdd") & "_" & ZnSh.Cells(2, 4) '日付とロットNOのフォルダを作成　2019.08.22
        SetPath = MonthPath & "\" & DayLotName
        If Dir(SetPath & ".xlsx", vbDirectory) = "" Then
            BK.SaveAs (SetPath & ".xlsx"), FileFormat:=xlOpenXMLWorkbook, Password:=Lotname & "." & Format(Date, "yyyy.mm.dd")  'pass LotNo.西暦.月.日
            SetAttr (SetPath & ".xlsx"), vbReadOnly
        Else
          For fCount = 2 To 999
            If Dir(SetPath & "." & Right("00" & CStr(fCount), 3) & ".xlsx", vbDirectory) = "" Then '同じ場合副番を付けて999まで保存する
              BK.SaveAs (SetPath & "." & Right("00" & CStr(fCount), 3) & ".xlsx"), FileFormat:=xlOpenXMLWorkbook, Password:=Lotname & "." & Format(Date, "yyyy.mm.dd")  'pass LotNo.西暦.月.日
              SetAttr (SetPath & "." & Right("00" & CStr(fCount), 3) & ".xlsx"), vbReadOnly
              Exit For
            End If
            If fCount = 999 Then
              MsgBox ("連番をこれ以上付けることが出来ません。"), vbCritical
              End
            End If
          Next
        End If
    End If
    
  '-------------------------------------------------------------------------------------------
  BK.Close
  ZnSh.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True  '保護設定
  '--------------------ﾌｧｲﾙｺﾋﾟｰ_処理ｽﾋﾟｰﾄﾞ速い_2021.01.25
  
  FileN = Temp & "_履歴\"
  Dim SrvPath As String
  Select Case UFMitsubishi.LineName.Caption
  Case "L-1":    SrvPath = HistoryPath & FileN & MonthName
  Case "LVC":    SrvPath = HistoryPath & FileN & MonthName
  Case "HVC":    SrvPath = HistoryPath & FileN & MonthName
  Case "ﾊﾞﾗﾝｻｰ": SrvPath = HistoryPath & FileN & MonthName
  Case "ﾄｯﾄ":    SrvPath = HistoryPath & FileN & MonthName
  Case "Test":   SrvPath = HistoryPath & "Test\ " & FileN & MonthName '"C:\Users\NLM Administrator\Desktop\全ヵ所検査_履歴\" & MonthName
  Case Else:     SrvPath = ""
  End Select
  
  Call CreateFolder(SrvPath)
  
  If Dir(SrvPath, vbDirectory) = "" Then MkDir SrvPath 'vbDirectory必要　指定すると引数setPathで指定したフォルダがあればフォルダ名を返し,なければ長さ0の文字列「""」を返します。
  
  Dim NewFiles As String: NewFiles = NewFileName(SetPath & "\")
  FileCopy SetPath & "\" & NewFiles, SrvPath & "\" & NewFiles
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Visible = True
End Sub
Public Sub 全ヵ所用転送() '2019.08.20
'専用フォーマットに転送する
Dim ws As Worksheet
Dim flag As Boolean, Flag2 As Boolean, Koflag As Boolean
Dim i As Long, a As Long, n As Long
Dim SheetsName As String, SheetsName2 As String, sh As String
Dim ShN As Worksheet
Dim ATU As String:  ATU = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 2)
Dim HABA As String: HABA = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 4)


Application.ScreenUpdating = False
  SheetsName = "ｺﾋﾟ-用"
  SheetsName2 = "転送済"
  For Each ws In ThisWorkbook.Worksheets
    If ws.Name = SheetsName Then flag = True
    If ws.Name = SheetsName2 Then Flag2 = True
  Next ws
  If flag = True Then    '同じ名前のシートあれば消す
      Application.DisplayAlerts = False
      ThisWorkbook.Sheets(SheetsName).Delete
      Application.DisplayAlerts = True
  End If
'------------------------------------------------整理用シート
  If UFMitsubishi.Kname = "" And UFMitsubishi.KenNo = "" Then
   MsgBox "検査員名と検査番号を入力してください"
    Exit Sub
  Else
       ThisWorkbook.Sheets.Add before:=Worksheets(1) 'Sheets(Sheets.Count) '新たにシートを追加する
       ActiveSheet.Name = "ｺﾋﾟ-用"
       With ActiveWorkbook.Sheets("ｺﾋﾟ-用").Tab
           .Color = 128
           .TintAndShade = 0
       End With
       ActiveWindow.Zoom = 100
     'コピー処理　2024.01.11　列追加 Z→AD
        n = ThisWorkbook.Sheets("全ヵ所検査").Cells(Rows.Count, 2).End(xlUp).row '1～30列
        ThisWorkbook.Sheets("全ヵ所検査").Range("A6:AD" & n).Copy _
        Destination:=Sheets("ｺﾋﾟ-用").Cells(1, 4)
     '不要列削除
        ThisWorkbook.Sheets("ｺﾋﾟ-用").Range("F:F,H:H,J:J,L:L,M:M,O:O,Q:Q,S:S,U:X,AA:AA,AC:AC,AE:AE").Delete 'ラテラルボー追加により削除列修正_2023.5.17_2024.1.11**歪判定列追加
     '数字変換
     For i = 2 To 8
      Select Case True
      Case ThisWorkbook.Sheets("ｺﾋﾟ-用").Cells(i, 19) = 1
        ThisWorkbook.Sheets("ｺﾋﾟ-用").Cells(i, 19) = "①"
      Case Sheets("ｺﾋﾟ-用").Cells(i, 19) = 2
        ThisWorkbook.Sheets("ｺﾋﾟ-用").Cells(i, 19) = "②"
      Case Sheets("ｺﾋﾟ-用").Cells(i, 19) = 3
        ThisWorkbook.Sheets("ｺﾋﾟ-用").Cells(i, 19) = "③"
      Case Sheets("ｺﾋﾟ-用").Cells(i, 19) = 4
        ThisWorkbook.Sheets("ｺﾋﾟ-用").Cells(i, 19) = "④"
      End Select
     Next i
     
   Call 全ヵ所用機器転送(ThisWorkbook.Sheets("全ヵ所検査"), 3, 29, ATU, HABA, ThisWorkbook.Sheets("全ヵ所検査").Cells(4, 8)) 'マイクロ
   With ThisWorkbook.Sheets("全ヵ所検査")
    .Unprotect 'シート保護解除
    .Cells(3, 29).Cells.Font.Name = "Meiryo UI"
    .Cells(3, 29).Cells.Font.Size = "8"
    .Cells(3, 29).HorizontalAlignment = xlLeft
    .Cells(3, 29).VerticalAlignment = xlCenter
    
    .Cells.Font.Underline = xlUnderlineStyleNone 'アンダーバー解除
    
    .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True 'シート保護
   End With
     '必要データ添付
        With ThisWorkbook.Sheets("ｺﾋﾟ-用")
            .Unprotect
            .Cells(1, 1) = "LotNo"
            .Cells(1, 2) = ThisWorkbook.Sheets("全ヵ所検査").Cells(2, 4)
            .Cells(2, 1) = "Size"
            .Cells(2, 2) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 2) & "×" & ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 4) & "×" & ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 6)
            .Cells(3, 1) = "測定機器"
            .Cells(3, 2) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 29)
            .Cells(5, 1) = "検査員"
            .Cells(5, 2) = UFMitsubishi.Kname
            .Cells(6, 1) = "作業日"
            .Cells(6, 2) = Format$(ThisWorkbook.Sheets("全ヵ所検査").Cells(1, 2), "yyyy/mm/dd  h:mm") & "～" & Format$(Now, "yyyy/mm/dd h:mm")
        End With
        Dim Clea1 As Borders: Set Clea1 = ThisWorkbook.Sheets("ｺﾋﾟ-用").Range(Cells(1, 5), Cells(71, 19)).Borders
        Clea1.LineStyle = xlLineStyleNone
        With ThisWorkbook.Sheets("ｺﾋﾟ-用").Cells
            .Font.Name = "Meiryo UI"
            .Font.Size = "10"
            .Interior.Pattern = xlNone '塗りつぶし
            .Interior.TintAndShade = 0
            .Interior.PatternTintAndShade = 0
            .Cells.EntireColumn.AutoFit
            .Cells.Locked = True 'セルロック
            .Cells.FormulaHidden = False
            
            .Font.Underline = xlUnderlineStyleNone 'アンダーバー解除
            
            .Borders(xlDiagonalDown).LineStyle = xlNone '枠線消し
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
            
            'Sheets("ｺﾋﾟ-用").Protect Password:=""
        ThisWorkbook.Sheets("ｺﾋﾟ-用").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True
        End With
  End If
  
  '---------------------------------------------------------------------------------------------------不良追加_2023.8.23
  Dim Lname As String: Lname = UFMitsubishi.LineName.Caption
  If (WsCheck("不良要因印刷" & "_" & Lname) = False) Then
    Call AddWS("不良要因印刷" & "_" & Lname)
    Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
  End If
  
  ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Cells(1, 2) = UFMitsubishi.LotBox
  
  If WsCheck("不良要因印刷" & "_" & Lname) Then Call 不良要因印刷範囲(UFMitsubishi.Kname)   '2023.5.29
  
  
'------------------------------------------------転送用シート　原紙　品証技術用
  '2019.09.11 押したボタンで三菱の判定を行う
  If UFMitsubishi.Inflag = "flagon" Then '三菱印刷用フラグ 通常は""
    If Flag2 = True Then    '同じ名前のシートあれば消す
        Application.DisplayAlerts = False
        Sheets(SheetsName2).Delete
        Application.DisplayAlerts = True
    End If
    ThisWorkbook.Sheets.Add before:=Worksheets(1) 'Sheets(Sheets.Count) '新たにシートを追加する
    ActiveSheet.Name = "転送済"
    With ThisWorkbook.Sheets("転送済").Tab
        .Color = 128
        .TintAndShade = 0
    End With
    ActiveWindow.Zoom = 50
'--------------------------------------------------------------------------原紙コピー
        If ThisWorkbook.Sheets("全ヵ所検査").Cells(2, 4) <> "" Then 'ロット登録がある場合
            ThisWorkbook.Sheets("三菱受託原紙").Visible = True
            ThisWorkbook.Sheets("三菱受託原紙").Cells.Copy
            ThisWorkbook.Sheets("転送済").Cells.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End If
'---------------------------------------------------------------------------図形挿入
      Dim myshape As Variant
      For Each myshape In ThisWorkbook.Sheets("三菱受託原紙").Shapes
       myshape.Copy
       ThisWorkbook.Sheets("転送済").Cells(2, 11).Select
       ThisWorkbook.Sheets("転送済").Paste
      Next
'---------------------------------------------------------------
      ThisWorkbook.Sheets("三菱受託原紙").Visible = False
      Application.CutCopyMode = False
  ElseIf UFMitsubishi.Inflag = "" Then '通常印刷用フラグ
  End If
  '2019.09.11 押したボタンで三菱の判定を行う
  If UFMitsubishi.Inflag = "flagon" Then '三菱印刷用フラグ
   '*********************************************************転送 2020.01.09 列追加により修正
    With ThisWorkbook.Sheets("転送済")
          .Cells(2, 4) = UFMitsubishi.KenNo   '検査番号
          .Cells(5, 5) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 2) '厚
          .Cells(5, 7) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 4) '幅
          .Cells(5, 9) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 6) '丈
          .Cells(8, 6) = "' " & "-" & ThisWorkbook.Sheets("全ヵ所検査").Cells(2, 13) & "mm " & Chr(10) & "+" & ThisWorkbook.Sheets("全ヵ所検査").Cells(2, 15) & "mm" '厚公差
          .Cells(8, 10) = "' " & "-" & ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 13) & "mm " & Chr(10) & "+" & ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 15) & "mm" '巾公差
          .Cells(8, 12) = "' " & "-" & ThisWorkbook.Sheets("全ヵ所検査").Cells(4, 13) & "mm " & Chr(10) & "+" & ThisWorkbook.Sheets("全ヵ所検査").Cells(4, 15) & "mm" '丈公差
          .Cells(19, 17) = ThisWorkbook.Sheets("ｺﾋﾟ-用").Cells(6, 2) '開始終了
          .Cells(19, 17).ShrinkToFit = True
          .Cells(20, 17) = UFMitsubishi.Kname '検査員
          
        Call 全ヵ所用機器転送(ThisWorkbook.Sheets("転送済"), 2, 7, ATU, HABA, ThisWorkbook.Sheets("全ヵ所検査").Cells(4, 8)) 'マイクロ
      '*********************************************************原板積載NO　2020.01.09
         For i = 10 To 16 Step 2
          .Cells(i, 2) = ThisWorkbook.Sheets("ｺﾋﾟ-用").Cells(i - 8, 17)
         Next i
      '*********************************************************切断ﾌﾞﾛｯｸNO　2020.01.09
         For i = 10 To 16 Step 2
          .Cells(i, 3) = ThisWorkbook.Sheets("ｺﾋﾟ-用").Cells(i - 8, 19)
         Next i
      '*********************************************************梱包NO　2020.01.09
         a = 0
         For i = 2 To 8 Step 2
         a = a + 1
          If ThisWorkbook.Sheets("ｺﾋﾟ-用").Cells(i, 5) <> "" Then '測定値がある場合のみ梱包番号記載
           .Cells(i + 8, 4) = a
          End If
         Next i
      '*********************************************************配列 2020.01.09 列追加により修正
          Dim rangeA As Range, RangeHT As Range, RangeSA As Range, RangeSHT As Range
          Dim KenA() As Variant, KenHT() As Variant
          Set rangeA = Range(ThisWorkbook.Sheets("ｺﾋﾟ-用").Cells(2, 5), ThisWorkbook.Sheets("ｺﾋﾟ-用").Cells(9, 8)) '元　厚
          KenA = rangeA
          Set RangeHT = Range(ThisWorkbook.Sheets("ｺﾋﾟ-用").Cells(2, 9), ThisWorkbook.Sheets("ｺﾋﾟ-用").Cells(9, 16)) '元　巾丈対角歪外観
          KenHT = RangeHT
    
          Set RangeSA = Range(.Cells(10, 6), .Cells(17, 9)) '先 厚
          Set RangeSHT = Range(.Cells(10, 10), .Cells(17, 17)) '先　巾丈対角歪外観
          '-----------------------------
           RangeSHT = KenHT '配列はきだし
           RangeSA = KenA
      '--------------------------------------------------------------------------セルロック.シート保護
        .Cells.Locked = True
        .Cells.FormulaHidden = False
    '    .Protect Password:=""
        .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'シート保護
    End With
  '***************************************歪、外観の中身チェック
    If UFMitsubishi.Work = "ON" Then '三菱受託用フラグにチェックが入っている時
      For i = 7 To (n)
        For a = 24 To 25
          If Sheets("全ヵ所検査").Cells(i, a) = "" Then
            MsgBox "歪入力がありません,はじめからやり直してください" & " " & " *[座標](" & i & "," & a & ")", vbExclamation
            For Each ShN In Sheets '入力がない場合やり直し
              If (ShN.Name = "ｺﾋﾟ-用" Or ShN.Name = "転送済" Or ShN.Name = "全ヵ所検査保存") Then
                Application.DisplayAlerts = False
                 ShN.Delete
                Application.DisplayAlerts = True
              End If
            Next ShN
            End 'シート削除後ループ停止
          End If
        Next a
        If ThisWorkbook.Sheets("全ヵ所検査").Cells(i, a) = "" Then Exit For
      Next i

      For i = 7 To (n)
        If ThisWorkbook.Sheets("全ヵ所検査").Cells(i, 26) = "" Then
          MsgBox "外観入力がありません,はじめからやり直してください" & " " & " *[座標](" & i & "," & "26" & ")", vbExclamation
            For Each ShN In Sheets '入力がない場合やり直し
              If (ShN.Name = "ｺﾋﾟ-用" Or ShN.Name = "転送済" Or ShN.Name = "全ヵ所検査保存") Then
                Application.DisplayAlerts = False
                 ShN.Delete
                Application.DisplayAlerts = True
              End If
            Next ShN
        End 'シート削除後ループ停止
        End If
      Next i
    End If
  End If
Application.ScreenUpdating = True
End Sub
Public Sub 全ヵ所用印刷範囲(Optional Temp As String = "全ヵ所検査") '2023.10.17_H494用　引数追加
'検査シートの印刷範囲
Dim n As Long, Kname As String

If WsCheck(Temp) = False Then Exit Sub
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets(Temp)

  Application.PrintCommunication = False
  n = sh.Cells(Rows.Count, 2).End(xlUp).row
  
  Select Case UFMitsubishi.WorkFlag.Caption
  Case "P1"
    sh.PageSetup.PrintArea = "$A$6:$AD$" & n '印刷範囲**********2024.1.11_変更
    Kname = UFMitsubishi.Kname
  Case "P_H494"
    sh.PageSetup.PrintArea = "$A$6:$AB$" & n '印刷範囲
    Kname = UFMitsubishi.H494_Kname
  End Select
  
  sh.Activate
  sh.Unprotect 'シート保護解除
  sh.Range("B7:AB76").Interior.Color = RGB(255, 255, 255) '色クリア
  sh.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True 'シート保護
   With sh.PageSetup
     .LeftHeader = ""
     .CenterHeader = ""
     .RightHeader = ""
     .LeftFooter = ""
     .CenterFooter = ""
     .LeftFooter = ""
     .CenterFooter = ""
     .RightFooter = ""
   End With
   Dim x As Double, y As Double, z As Double
   Dim Am As Double, Ap As Double
   Dim Hm As Double, Hp As Double
   Dim Tm As Double, Tp As Double
   With sh
     x = val(.Cells(3, 2))
     y = val(.Cells(3, 4))
     z = val(.Cells(3, 6))
     Am = val(.Cells(2, 13))
     Ap = val(.Cells(2, 15))
     Hm = val(.Cells(3, 13))
     Hp = val(.Cells(3, 15))
     Tm = val(.Cells(4, 13))
     Tp = val(.Cells(4, 15))
   End With
   With sh.PageSetup '"&""Meiryo UI""03" &
     .RightFooter = "&""Meiryo UI""&8" & "*" & Format$(sh.Cells(1, 2), "yyyy/mm/dd  h:mm") & "～" & Format$(Now, "yyyy/mm/dd  h:mm") & " " & "印刷"
     
     .RightHeader = "&""Meiryo UI""&8" & "*" & "厚公差：" & Format((x - Am), "0.000") & "～" & Format((x + Ap), "0.000") & vbCr & "*" _
     & "巾公差：" & Format((y - Hm), "0.0") & "～" & Format((y + Hp), "0.0") & vbCr & "*" _
     & "丈公差：" & Format((z - Tm), "0.0") & "～" & Format((z + Tp), "0.0") & vbCr & "*" _
     & "歪公差：" & "幅:" & Format(val(sh.Cells(5, 13)), "0.0") & "_" & "丈:" & Format(val(sh.Cells(5, 15)), "0.0") & vbCr & "*" _
     & "対角公差：" & Format(val(sh.Cells(1, 19)), "0.0") & vbCr & "*" _
     & "ﾗﾃﾗﾙﾎﾞｰ公差：" & Format(val(sh.Cells(1, 13)), "0.0")

     .LeftHeader = "&""Meiryo UI""&8" & "*" & sh.Cells(2, 4) & vbCr & "*" _
     & Format(val(sh.Cells(3, 2)), "0.000") & "×" & Format(val(sh.Cells(3, 4)), "0.0") & "×" & Format(val(sh.Cells(3, 6)), "0.0") & vbCr & "*" _
     & sh.Cells(1, 30) & "_" & sh.Cells(2, 30)
     
     .LeftFooter = "&""Meiryo UI""&8" & "*検査員：" & Kname & " " & "*測定機器：" & sh.Cells(3, 29)
     
     .HeaderMargin = Application.InchesToPoints(0.2) '2023.7.21_0をなくした*印刷が全範囲の場合FooterHeaderの情報が消えていた
     .FooterMargin = Application.InchesToPoints(0.2)
     .LeftMargin = Application.InchesToPoints(0.5)
     .RightMargin = Application.InchesToPoints(0.5)
     .TopMargin = Application.InchesToPoints(0.8)
     .BottomMargin = Application.InchesToPoints(0.5)
     .CenterHorizontally = True
     .CenterVertically = True
     .PaperSize = xlPaperA4 'サイズ
     .Orientation = xlLandscape  '縦xlPortrait　横 xllandscape
     .FitToPagesWide = 1 '1ページで納める 巾方向
     .FitToPagesTall = 1 '丈方向は何ページになってもよい
   End With
   Application.PrintCommunication = True
End Sub
Public Sub 原紙印刷範囲()
'専用フォーマットの印刷範囲
Dim sh As Worksheet, n As Long
Application.PrintCommunication = False
If UFMitsubishi.Choice = "5N" Then
  If WsCheck("5N") = False Then Exit Sub
  Set sh = ThisWorkbook.Sheets("5N")
  sh.PageSetup.PrintArea = ""
  n = sh.Cells(Rows.Count, 3).End(xlUp).row + 3
  sh.PageSetup.PrintArea = "$A$1:$L$" & n '印刷範追加　2020.06.15
ElseIf (UFMitsubishi.Choice = "三菱") Or (UFMitsubishi.Choice = "一般") Then
 If WsCheck("転送済") = False Then Exit Sub
 Set sh = ThisWorkbook.Sheets("転送済")
 sh.PageSetup.PrintArea = ""
 sh.PageSetup.PrintArea = "$B$2:$R$20" '印刷範囲 2020.01.09列+1修正
End If
With sh.PageSetup
 .LeftHeader = "&""Meiryo UI""&7" & UFMitsubishi.LineName.Caption & vbCr & "*" _
 & "&""Meiryo UI""&7" & "*" & ThisWorkbook.Sheets("全ヵ所検査").Cells(1, 30) & "_" & ThisWorkbook.Sheets("全ヵ所検査").Cells(2, 30)
 .CenterHeader = ""
 .RightHeader = ""
 .LeftFooter = ""
 .CenterFooter = ""
 .LeftFooter = ""
 .CenterFooter = ""
 .RightFooter = ""
' .RightFooter = "&""Meiryo UI""&8" & "*" & Format$(Now, "yyyy/mm/dd  h:mm") & " " & "印刷" 'フォント名は「&」の後ろに「"」（ダブルクォーテーション）で括って指定
'もともと印刷されている内容なので停止_2023.11.21
 .HeaderMargin = Application.InchesToPoints(0)
 .FooterMargin = Application.InchesToPoints(0)
 .LeftMargin = Application.InchesToPoints(0#)
 .RightMargin = Application.InchesToPoints(0#)
 .TopMargin = Application.InchesToPoints(0.5)
 .BottomMargin = Application.InchesToPoints(0#)
 .CenterHorizontally = True '横
' .CenterVertically = True'縦
 .PaperSize = xlPaperA4 'サイズ
 .Orientation = xlLandscape  '縦xlPortrait　横 xllandscape
' If UFMitsubishi.Choice = "5N" Then
 .FitToPagesWide = 1
' .FitToPagesTall = False '丈方向は何ページになってもよい
' Else
 .FitToPagesTall = 1 '1ページに収める
' End If
End With
Application.PrintCommunication = True
End Sub
Public Sub 全ヵ所用機器転送(ws As Worksheet, x As Long, y As Long, ATU As String, HABA As String, Optional Yout As String = "") 'ｼｰﾄ、座標行、座標列,板厚
'2023.5.17_引数追加
Dim flag As Boolean, Flag2 As Boolean, Koflag As Boolean
Dim i As Long, a As Long
'*********************************************************マイクロ
With ws
ws.Unprotect 'シート保護解除
Dim line As String, Conb As String
Select Case UFMitsubishi.LineName.Caption
Case "L-1"
  line = "L-1"
  Conb = " C11"
Case "HVC"
  line = "HVC"
  Conb = " C11"
Case "LVC"
  line = "LVC"
  Conb = " C10"
Case "ﾊﾞﾗﾝｻｰ"
  line = "ﾊﾞﾗﾝｻｰ"
  Conb = " C11"
Case "ﾄｯﾄ"
  line = "ﾄｯﾄ"
  Conb = " C12"
End Select
If 0 < val(ATU) And val(ATU) <= 28# Then   'ライン外がL-1に変更2018.12.24
  .Cells(x, y) = "ﾏｲｸﾛ:" & line & "M1" & "　　ｺﾝﾍﾞｯｸｽ:" & line & Conb
ElseIf 28 < val(ATU) And val(ATU) <= 54# Then
  .Cells(x, y) = "ﾏｲｸﾛ:" & line & "M2" & "　　ｺﾝﾍﾞｯｸｽ:" & line & Conb
ElseIf 54# < val(ATU) And val(ATU) <= 74# Then
  .Cells(x, y) = "ﾏｲｸﾛ:" & line & "M3" & "　　ｺﾝﾍﾞｯｸｽ:" & line & Conb
ElseIf 74# < val(ATU) And val(ATU) < 300 Then      'ノギス
  .Cells(x, y) = "ﾉｷﾞｽ:" & line & "N5" & "　　ｺﾝﾍﾞｯｸｽ:" & line & Conb
ElseIf 300# <= val(ATU) Then  'ノギス
  .Cells(x, y) = "ﾉｷﾞｽ:" & line & "N6" & "　　ｺﾝﾍﾞｯｸｽ:" & line & Conb
  '300
  If val(ATU) = 300# Then
    For i = 14 To 18 '測定値のﾁｪｯｸ
      For a = 2 To 8 Step 2
        If IsNumeric(.Cells(i, a)) = True And (.Cells(i, a)) <> "" Then '数字
          If val(.Cells(i, a)) < 300# Then 'N5
           flag = True
          ElseIf val(.Cells(i, a)) >= 300# Then 'N6
           Koflag = True
          ElseIf (.Cells(i, a)) = "" Then '空白
          End If
        End If
      Next a
    Next i
    '書き込み内容
    If flag = True And Koflag = True Then
      .Cells(x, y) = "ﾉｷﾞｽ：" & line & "N5,N6" & "　　ｺﾝﾍﾞｯｸｽ:" & line & "C11"
    ElseIf flag = True And Koflag = False Then
      .Cells(x, y) = "ﾉｷﾞｽ" & line & "N5" & "　　ｺﾝﾍﾞｯｸｽ:" & line & "C11"
    ElseIf Koflag = True And flag = False Then
      .Cells(x, y) = "ﾉｷﾞｽ" & line & "N6" & "　　ｺﾝﾍﾞｯｸｽ:" & line & "C11"
    End If
  End If
End If

'特殊用途コード***2023.5.11
'最後に実行して通常分を書き換え
  If Yout = "H494" Then '用途コードH494に限り幅測定をノギスで行う
    Select Case val(HABA) '幅寸法が範囲内の時
    Case 1# To 299#
      .Cells(x, y) = "ﾏｲｸﾛ:" & line & "M1" & "　　幅:ﾉｷﾞｽ:" & "N5" & "_丈:ｺﾝﾍﾞｯｸｽ:" & line & Conb
    End Select
  End If

End With
End Sub

Public Sub FiveN材転送() '2020.06.12
'専用フォーマットに転送する
Dim ws As Worksheet
Dim flag As Boolean, Flag2 As Boolean, Koflag As Boolean
Dim i As Long, a As Long, n As Long
Dim SheetsName As String, SheetsName2 As String, sh As String
Dim ShN As Worksheet
Call AddWS("ｺﾋﾟ-用")
Dim CpySh As Worksheet: Set CpySh = ThisWorkbook.Sheets("ｺﾋﾟ-用")
Dim ZenSh As Worksheet: Set ZenSh = ThisWorkbook.Sheets("全ヵ所検査")
Dim ATU As String:  ATU = ZenSh.Cells(3, 2)
Dim HABA As String: HABA = ZenSh.Cells(3, 4)
'------------------------------------------------整理用シート
  If UFMitsubishi.Kname = "" And UFMitsubishi.KenNo = "" Then
   MsgBox "検査員名と検査番号を入力してください", vbCritical
  Exit Sub
  Else
    ZenSh.Unprotect 'シート保護解除
    'コピー処理　2020.01.09　列追加 Z→AC
    n = ZenSh.Cells(Rows.Count, 2).End(xlUp).row '1～26列
    ZenSh.Range("A6:AA" & n).Copy _
    Destination:=CpySh.Cells(1, 4)
    '不要列削除
    CpySh.Range("F:F,H:H,J:J,L:L,M:M,O:O,Q:Q,S:S,U:X").Delete '2023.11.21_修正ラテラルボーとゲージの部分が必要
  
   Call 全ヵ所用機器転送(ZenSh, 2, 25, ATU, HABA) '*********************************************************マイクロ
   With ZenSh
    .Cells(2, 25).Cells.Font.Name = "Meiryo UI"
    .Cells(2, 25).Cells.Font.Size = "8"
    .Cells(2, 25).HorizontalAlignment = xlLeft
    .Cells(2, 25).VerticalAlignment = xlCenter
    .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True 'シート保護
   End With
     '必要データ添付
      CpySh.Cells(1, 1) = "LotNo"
      CpySh.Cells(1, 2) = ZenSh.Cells(2, 4)
      CpySh.Cells(2, 1) = "Size"
      CpySh.Cells(2, 2) = ZenSh.Cells(3, 2) & "×" & ZenSh.Cells(3, 4) & "×" & ZenSh.Cells(3, 6)
      CpySh.Cells(3, 1) = "測定機器"
      CpySh.Cells(3, 2) = ZenSh.Cells(2, 25)
      CpySh.Cells(5, 1) = "検査員"
      CpySh.Cells(5, 2) = UFMitsubishi.Kname
      CpySh.Cells(6, 1) = "作業日"
      CpySh.Cells(6, 2) = Format$(Sheets("全ヵ所検査").Cells(1, 2), "yyyy/mm/dd  h:mm") & "～" & Format$(Now, "yyyy/mm/dd h:mm")
      Dim Clea1 As Borders: Set Clea1 = Range(CpySh.Cells(1, 5), CpySh.Cells(31, 19)).Borders
      Clea1.LineStyle = xlLineStyleNone
      With CpySh.Cells
        .Font.Name = "Meiryo UI"
        .Font.Size = "10"
        .Interior.Pattern = xlNone '塗りつぶし
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Cells.EntireColumn.AutoFit
        .Cells.Locked = True 'セルロック
        .Cells.FormulaHidden = False
        
        .Font.Underline = xlUnderlineStyleNone 'アンダーバー解除
        
        'CpySh.Protect Password:=""
      CpySh.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True
      End With
  End If
'--------------------------------------------------転送用シート　原紙　品証技術用
If Not UFMitsubishi.Choice = "5N" Then
MsgBox "5N専用"
Exit Sub
End If

  Call AddWS("5N")
  Dim FivSh As Worksheet: Set FivSh = ThisWorkbook.Sheets("5N")
  FivSh.Tab.Color = 128
  '--------------------------------------------------原紙コピー
  If ZenSh.Cells(2, 4) <> "" Then 'ロット登録がある場合
  If WsCheck("5N原紙") = False Then Exit Sub
  Dim GSh As Worksheet: Set GSh = ThisWorkbook.Sheets("5N原紙")
  GSh.Visible = True
  GSh.Cells.Copy
  FivSh.Cells.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  '--------------------------------------------------行高さ調整
  Dim myBaseRng As Range, myNewRng As Range
  Dim myRow As Long, myClm As Long
  Set myBaseRng = Range(GSh.Cells(2, 1), GSh.Cells(16, 12))
  Set myNewRng = FivSh.Range("A2")
  With myBaseRng
    For myClm = 1 To .Columns.Count
        myNewRng.Columns(myClm).ColumnWidth = .Columns(myClm).ColumnWidth
    Next
    For myRow = 1 To .Rows.Count
        myNewRng.Rows(myRow).RowHeight = .Rows(myRow).RowHeight
    Next
  End With
  Set myBaseRng = Nothing
  Set myNewRng = Nothing
  End If
  '--------------------------------------------------図形挿入
  Dim myshape As Variant
  For Each myshape In ThisWorkbook.Sheets("5N原紙").Shapes
  myshape.Copy
  FivSh.Activate
  FivSh.Cells(2, 9).Select
  FivSh.Paste
  Next
  '---------------------------------------------------------------
  ThisWorkbook.Sheets("5N原紙").Visible = False
  Application.CutCopyMode = False
  '--------------------------------------------------転送 2020.01.09 列追加により修正
  FivSh.Cells(2, 2) = UFMitsubishi.KenNo   '検査番号
  FivSh.Cells(4, 3) = ZenSh.Cells(3, 2) '厚
  FivSh.Cells(4, 5) = ZenSh.Cells(3, 4) '幅
  FivSh.Cells(4, 7) = ZenSh.Cells(3, 6) '丈
  FivSh.Cells(7, 4) = "’" & "-" & ZenSh.Cells(2, 13) & "mm " & Chr(10) & "+" & ZenSh.Cells(2, 15) & "mm" '厚公差
  FivSh.Cells(15, 11) = CpySh.Cells(6, 2) '開始終了
  FivSh.Cells(15, 11).ShrinkToFit = True '列幅に収まるように自動的に文字列を縮小
'  FivSh.Cells(15, 11).Font.Size = 7
  FivSh.Cells(16, 11) = UFMitsubishi.Kname '検査員
  Call 全ヵ所用機器転送(FivSh, 2, 6, ATU, HABA) 'マイクロ
  '--------------------------------------------------行数調整　検査員.開始終了入力後調整する
  Dim j As Long: j = CpySh.Cells(Rows.Count, 4).End(xlUp).row '番号列最終行
  Dim Ins As Long, InEnd As Long
  If val(CpySh.Cells(j, 4)) > 5 Then '板番号の最大値が5枚以上(公差外検査がある)
    Ins = val(CpySh.Cells(j, 4)) - 5 '最大数から-5
    InEnd = 14 + Ins - 1 '5Nｼｰﾄの5枚目以降の行番号が14行目
    FivSh.Rows(14 & ":" & InEnd).Insert Shift:=xlDown
'    FivSh.Cells(14, 2) = "公差外検査"
     FivSh.Cells(14, 2).Font.Size = 9
    '--------------------------------------------------罫線
    Dim Bord As Borders: Set Bord = Range(FivSh.Cells(14, 3), FivSh.Cells(InEnd, 6)).Borders
    Dim Bord2 As Borders: Set Bord2 = Range(FivSh.Cells(14, 9), FivSh.Cells(InEnd, 12)).Borders
    Dim BordRange As Range: Set BordRange = Range(FivSh.Cells(14, 3), FivSh.Cells(InEnd, 6))
    Dim BordRange2 As Range: Set BordRange2 = Range(FivSh.Cells(14, 9), FivSh.Cells(InEnd, 12))
    Bord.LineStyle = xlContinuous '格子
    Bord2.LineStyle = xlContinuous '格子
    BordRange.Borders(xlEdgeTop).LineStyle = xlDot '点線
    BordRange.Borders(xlEdgeTop).Weight = xlHairline '細かく
    BordRange2.Borders(xlEdgeTop).LineStyle = xlDot '点線
    BordRange2.Borders(xlEdgeTop).Weight = xlHairline '細かく
    BordRange.Borders(xlInsideHorizontal).LineStyle = xlDot '点線
    BordRange.Borders(xlInsideHorizontal).Weight = xlHairline '細かく
    BordRange2.Borders(xlInsideHorizontal).LineStyle = xlDot '点線
    BordRange2.Borders(xlInsideHorizontal).Weight = xlHairline '細かく
    Range(FivSh.Cells(14, 10), FivSh.Cells(InEnd, 11)).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
    Range(FivSh.Cells(14, 10), FivSh.Cells(InEnd, 11)).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
'    Range(FivSh.Cells(14, 1), FivSh.Cells(14, 2)).Borders(xlEdgeTop).LineStyle = xlDouble
'    BordRange.Borders(xlEdgeTop).LineStyle = xlDouble '良品区分　二重線
  End If
  '--------------------------------------------------配列 2020.01.09 列追加により修正
  Dim rangeA As Range, RangeHT As Range, RangeSA As Range, RangeSHT As Range
  Dim KenA() As Variant, KenHT() As Variant
  Dim k As Long: k = CpySh.Cells(Rows.Count, 5).End(xlUp).row
  Set rangeA = Range(CpySh.Cells(2, 5), CpySh.Cells(k, 8)) '元　厚
  KenA = rangeA
  Set RangeHT = Range(CpySh.Cells(2, 18), CpySh.Cells(k, 18)) '元　外観
  KenHT = RangeHT
  Set RangeSA = Range(FivSh.Cells(9, 3), FivSh.Cells(8 + UBound(KenA, 1), 6)) '先 厚
  Set RangeSHT = Range(FivSh.Cells(9, 9), FivSh.Cells(8 + UBound(KenHT, 1), 9)) '先　巾丈対角歪外観
  RangeSHT = KenHT '配列はきだし
  RangeSA = KenA
  '---------------------------------------------------------------------------------------------------不良追加_2023.8.23
  Dim Lname As String: Lname = UFMitsubishi.LineName.Caption
  If (WsCheck("不良要因印刷" & "_" & Lname) = False) Then
    Call AddWS("不良要因印刷" & "_" & Lname)
    Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
  End If
  
  ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Cells(1, 2) = UFMitsubishi.LotBox
  
  If WsCheck("不良要因印刷" & "_" & Lname) Then Call 不良要因印刷範囲(UFMitsubishi.Kname)   '2023.5.29
  
  '--------------------------------------------------セルロック.シート保護
  FivSh.Cells.Locked = True
  FivSh.Cells.FormulaHidden = False
  '    FivSh.Protect Password:=""
  FivSh.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'シート保護
End Sub
Public Sub FiveN専用保存() '2020.06.12
If WsCheck("全ヵ所検査") = False Then
  MsgBox "保存ｼｰﾄがありません", vbCritical
  Exit Sub
End If
Dim ws As Worksheet, flag As Boolean
Dim sh As Worksheet
Dim fCount As Integer, Lcount As Integer
Dim DayLotName As String, MonthName As String, Lotname As String
Dim SetPath As String, MonthPath As String
Dim wkArray()
Dim idx, Cnt As Long, n As Long
Dim BK As Workbook
Dim Lname As String: Lname = UFMitsubishi.LineName.Caption
Application.Visible = True
Application.ScreenUpdating = False
Call AddWS("全ヵ所検査保存")
Dim ZnSh As Worksheet: Set ZnSh = ThisWorkbook.Worksheets("全ヵ所検査保存")
Dim KeSh As Worksheet: Set KeSh = ThisWorkbook.Worksheets("全ヵ所検査")
  '------------------------------------------------コピー処理 保存閲覧用
  KeSh.Unprotect
  KeSh.Cells.Copy
  ZnSh.Cells.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  With ZnSh.Cells.Interior '色塗りをなくす
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  ZnSh.Cells(1, 3) = "～"
  ZnSh.Cells(1, 4) = Format$(Now, "yyyy/mm/dd h:mm")
  Range(ZnSh.Cells(1, 3), ZnSh.Cells(1, 4)).Font.Size = 5
  ZnSh.Cells.EntireColumn.AutoFit
  '列幅調整
  Dim i As Long, columnsToSetWidth As Variant, col As Variant
  For i = 1 To 30
    ZnSh.Columns(i).ColumnWidth = 8.75
  Next i
  columnsToSetWidth = Array(3, 5, 7, 9, 10, 12, 14, 16, 18, 21, 24, 26, 28)
  For Each col In columnsToSetWidth
      ZnSh.Columns(col).ColumnWidth = 2.5
  Next col
  
  ZnSh.Cells(1, 6) = "検査員：" & UFMitsubishi.Kname
  ZnSh.Cells(1, 6).Font.Size = 7
  ZnSh.Cells.Locked = True
  ZnSh.Cells.FormulaHidden = False
'  ZnSh.Protect Password:=""
  ZnSh.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'シート保護
  '------------------------------------------------
  '必要データのみのシートをｻｰﾊﾞｰ保存する
  Call DisplaySheets
  With ThisWorkbook
    For idx = 1 To .Worksheets.Count
      If (.Worksheets(idx).Name) = "全ヵ所検査保存" Or (.Worksheets(idx).Name) = "ｺﾋﾟ-用" _
      Or (.Worksheets(idx).Name) = "転送済" Or (.Worksheets(idx).Name) = "5N" Or (.Worksheets(idx).Name) = "不良要因印刷" & "_" & Lname Then
        Cnt = Cnt + 1
        ReDim Preserve wkArray(1 To Cnt)
        wkArray(Cnt) = .Worksheets(idx).Name
      End If
    Next idx
    If Cnt = 0 Then
    Exit Sub
    Else
    .Worksheets(wkArray).Copy
    End If
  End With
  Call HiddenSheets
  ThisWorkbook.Sheets("全ヵ所検査").Visible = True
  '------------------------------------------------2019.08.28 アクティブがうつって保存できないエラーが出た
  Set BK = ActiveWorkbook
  '------------------------------------------------
  Application.DisplayAlerts = False
'月毎にフォルダを作成してその中に格納する2018.12.19'名前を付けて保存　名前が被れば001～数字を付ける
  MonthName = Format(Date, "yyyy.mm")
  MonthPath = "C:\検査履歴" & "\" & MonthName '-----------------------ﾛｰｶﾙ保存_2021.01.25
  Lotname = ZnSh.Cells(2, 4) 'ロットNO
  DayLotName = Format(Date, "yyyymmdd") & "_" & ZnSh.Cells(2, 4) '日付とロットNOのフォルダを作成　2019.08.22
  SetPath = MonthPath & "\" & DayLotName
  
  Call CreateFolder(SetPath) 'ﾌｫﾙﾀﾞを作成
  
  Dim FilePath As String:     FilePath = SetPath & "\" & DayLotName & ".xlsx"
  
  If Dir(FilePath, vbDirectory) = "" Then
      '2019.08.28 , FileFormat:=xlOpenXMLWorkbook
      BK.SaveAs (FilePath), FileFormat:=xlOpenXMLWorkbook, Password:=Lotname & "." & Format(Date, "yyyy.mm.dd")     'pass LotNo.西暦.月.日
      SetAttr (FilePath), vbReadOnly
  Else
  For fCount = 2 To 999
    If Dir(SetPath & "\" & DayLotName & "." & Right("00" & CStr(fCount), 3) & ".xlsx", vbDirectory) = "" Then '同じ場合副番を付けて999まで保存する
      BK.SaveAs (SetPath & "\" & DayLotName & "." & Right("00" & CStr(fCount), 3) & ".xlsx"), FileFormat:=xlOpenXMLWorkbook, Password:=Lotname & "." & Format(Date, "yyyy.mm.dd")    'pass LotNo.西暦.月.日
      SetAttr (SetPath & "\" & DayLotName & "." & Right("00" & CStr(fCount), 3) & ".xlsx"), vbReadOnly
      Exit For
    End If
    If fCount = 999 Then
      MsgBox ("連番をこれ以上付けることが出来ません。"), vbCritical
      End
    End If
  Next
  End If
  BK.Close
  ZnSh.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True  '保護設定
  
  '--------------------ﾌｧｲﾙｺﾋﾟｰ_処理ｽﾋﾟｰﾄﾞ速い_2021.01.25
  Dim SrvPath As String
  Select Case UFMitsubishi.LineName.Caption
  Case "L-1":    SrvPath = HistoryPath & "全ヵ所検査_履歴\" & MonthName
  Case "LVC":    SrvPath = HistoryPath & "全ヵ所検査_履歴\" & MonthName
  Case "HVC":    SrvPath = HistoryPath & "全ヵ所検査_履歴\" & MonthName
  Case "ﾊﾞﾗﾝｻｰ": SrvPath = HistoryPath & "全ヵ所検査_履歴\" & MonthName
  Case "ﾄｯﾄ":    SrvPath = HistoryPath & "全ヵ所検査_履歴\" & MonthName
  Case "Test":   SrvPath = HistoryPath & "Test\Test_履歴\全ヵ所検査_履歴\" & MonthName '"C:\Users\NLM Administrator\Desktop\全ヵ所検査_履歴\" & MonthName
  Case Else:     SrvPath = ""
  End Select
  
  Call CreateFolder(SrvPath) 'ﾌｫﾙﾀﾞを作成
  
  Dim NewFiles As String: NewFiles = NewFileName(SetPath & "\")
  FileCopy SetPath & "\" & NewFiles, SrvPath & "\" & NewFiles
  
  '品証技術用に別に保存する　\\nlmsrvngy03\NPS\全社研修会トレーナー教育\第36回NPSトレーナー教育（名古屋）2017年\梱包実績の見える化\
  '-------------------------------------------------------------------------------------------
  If UFMitsubishi.Inflag = "" And UFMitsubishi.Work = "OFF" And UFMitsubishi.Choice = "5N" Then '5N
      MonthName = Format(Date, "yyyy.mm")
      Dim Kouzyou As String:      Kouzyou = FactoryyPath & "【梱包】５N\"   '2021.01.22_パス変更
      Dim KouzyouPath As String
      Select Case UFMitsubishi.LineName.Caption '"\\nlmsrvngy03\NPS\全社研修会トレーナー教育\第36回NPSトレーナー教育（名古屋）2017年\梱包実績の見える化\5N_履歴\"
      Case "L-1":    KouzyouPath = Kouzyou & MonthName
      Case "LVC":    KouzyouPath = Kouzyou & MonthName
      Case "HVC":    KouzyouPath = Kouzyou & MonthName
      Case "ﾊﾞﾗﾝｻｰ": KouzyouPath = Kouzyou & MonthName
      Case "ﾄｯﾄ":    KouzyouPath = Kouzyou & MonthName
      Case "Test":   KouzyouPath = HistoryPath & "Test\5N_履歴\" & MonthName ' '"C:\Users\NLM Administrator\Desktop\5N_履歴\" & MonthName
      Case Else:     KouzyouPath = ""
      End Select
      
      Call CreateFolder(KouzyouPath) 'ﾌｫﾙﾀﾞを作成
      
      '"C:\Users\NLM Administrator\Desktop\三菱受託検査 履歴\" & MonthName '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\三菱受託検査 履歴\" & MonthName 'テスト用パスの定義
      FileCopy SetPath & "\" & NewFiles, KouzyouPath & "\" & NewFiles
  End If
  '-------------------------------------------------------------------------------------------
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Visible = True
End Sub
Public Function Judg_PorM(ATU As Double, KM As Double, KP As Double) As Boolean '公差外で+-どちらにはずれたか
If ATU > KP Then
  Judg_PorM = True 'P
ElseIf ATU < KM Then
  Judg_PorM = False 'M
End If
End Function
'指定した範囲（RangeA）内でRGB(255, 0, 0)の背景色かつ文字列が "×" のセルがある場合に加算_2024.1.12
Function CheckRedXCells(rangeA As Range) As Long
Dim Cell As Range
Dim Count As Long

For Each Cell In rangeA
  If Cell.Interior.Color = RGB(255, 0, 0) And Cell.Value = "×" Then
    Count = Count + 1
  End If
Next Cell
CheckRedXCells = Count
End Function
Function ExtractNthText(InputString As String, a As Long) As String  '指定の()内の文字列を返す
Dim StartPos As Long, EndPos As Long, i As Long
' Aが1より小さい場合は空の文字列を返す
If a < 1 Then
  ExtractNthText = ""
  Exit Function
End If
' A番目の"()"セットを探し、その位置を取得
For i = 1 To a
  StartPos = InStr(StartPos + 1, InputString, "(") ' 最初の"("の位置を取得
  If StartPos = 0 Then ' "("が見つからない場合は空の文字列を返す
    ExtractNthText = ""
    Exit Function
  End If

  EndPos = InStr(StartPos, InputString, ")") ' ")"の位置を取得
  
  If EndPos = 0 Then ' ")"が見つからない場合は空の文字列を返す
    ExtractNthText = ""
    Exit Function
  End If
Next i
If StartPos > 0 And EndPos > StartPos Then
  ExtractNthText = ExtractBefore(Mid(InputString, StartPos + 1, EndPos - StartPos - 1))  '"÷"より前だけ抜き取る
Else
  ExtractNthText = ""
End If
End Function
Function ExtractBefore(Temp As String, Optional n As String = "÷") As String  'nが見つかった場合、その前の文字列を抜き取る
Dim pos As Long
pos = InStr(1, Temp, n) 'nの位置を取得
If pos > 0 Then 'その前の文字列を抜き取る
  ExtractBefore = Left(Temp, pos - 1)
Else
  ExtractBefore = ""  'nが見つからない場合は""を返す
End If
End Function
Function CountParenthesesSets(InputString As String) As Long ' ()のセットの数を返す
Dim i As Long, Count As Long, InsideParentheses As Boolean
' 文字列を一文字ずつ検査して()のセットの数を数える
For i = 1 To Len(InputString)
  If Mid(InputString, i, 1) = "(" Then
    InsideParentheses = True ' "("を見つけたらinsideParenthesesをTrueに設定
  ElseIf Mid(InputString, i, 1) = ")" And InsideParentheses Then
    Count = Count + 1 ' insideParenthesesがTrueでかつ")"を見つけたらセットとしてカウント
    InsideParentheses = False ' insideParenthesesをFalseに戻す
  End If
Next i
CountParenthesesSets = Count
End Function
Public Sub ForN_Select() 'UF起動時に処理させる
If UFMitsubishi.Tag = "off" Then Exit Sub
Dim ws As Worksheet, i As Long, n As Long, tMP As String, myUserForm As Object
If WsCheck("全ヵ所検査") = False Then Exit Sub
Set ws = ThisWorkbook.Worksheets("全ヵ所検査")
ws.Activate
With UFMitsubishi
  Select Case UFMitsubishi.Choice.Text '切替
  Case "三菱"
  '受託品作業
    .Work = "ON" '三菱受託用フラグ
    .Inflag = "flagon" '三菱印刷用フラグ
    .Normal.Visible = False
    .Mitsubishi.Visible = True
    .FiNi.Visible = False
    ws.Unprotect '保護解除
    ws.Cells(1, 22) = "三菱"
  Case "5N"
  '5N作業
    .Work = "OFF"
    .Inflag = ""
    .Normal.Visible = False
    .Mitsubishi.Visible = False
    .FiNi.Visible = True
    ws.Unprotect '保護解除
    ws.Cells(1, 22) = "5N"
    .FiveOP = True
    ws.Cells(5, 13) = "2.0"
    ws.Cells(5, 15) = "3.0"
  Case "一般"
  '一般作業
    .Work = "OFF"
    .Inflag = ""
    .Normal.Visible = True
    .Mitsubishi.Visible = False
    .FiNi.Visible = False
    ws.Unprotect '保護解除
    ws.Cells(1, 22) = "一般"
    .一般 = True
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
  End Select
  
  '2024.1.26*****ごちゃごちゃうるさいから関数作成、リロード後フォームに反映するように修正
  '検入高さセット
  If ws.Cells(3, 30) <> "" Then
    If Right(ws.Cells(3, 30), 1) = "枚" Then
      UFMitsubishi.YamaLB.ListIndex = UFMitsubishi.YamaLB.ListCount - 1
      If (ws.Cells(3, 2) <> "") Then Call DrawListBoxContent(UFMitsubishi, val(ws.Cells(3, 2)))
      UFMitsubishi.KenHBox7 = ExtractBefore(ws.Cells(3, 30), "枚")
    End If
  End If
  If ws.Cells(2, 30) <> "" Then UFMitsubishi.YamaLB.ListIndex = val(Left(ws.Cells(2, 30), 1)) - 1
  If (ws.Cells(3, 2) <> "") Then Call DrawListBoxContent(UFMitsubishi, val(ws.Cells(3, 2)))

  '検入高さリロード
  If ws.Cells(2, 30) <> "" Then
    n = val(Left(ws.Cells(2, 30), 1))
    ' nが0になるまでループ
    Do While n > 0
      tMP = ExtractNthText(ws.Cells(1, 30), n)
      UFMitsubishi("KenHBox" & n).Value = tMP
      n = n - 1 ' n
    Loop
  End If
End With
End Sub

