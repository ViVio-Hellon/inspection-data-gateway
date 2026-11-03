Attribute VB_Name = "【UFKen_12】Cask"
Public Function MoveTg(n As String, sh As String) As Range
Dim i As Long, j As Long
Dim flag_Blank As Boolean: flag_Blank = False
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(sh)
ws.Activate
Select Case n
Case "A"
  Select Case ws.Name
  Case "キャスク検査"
    For i = 7 To 706
      For j = 2 To 6 Step 2
        If flag_Blank = True Then Exit For
          If ws.Cells(i, j) = "" Then
            Set MoveTg = ws.Cells(i, j)
            flag_Blank = True
            Exit For
          End If
      Next j
    If flag_Blank = True Then Exit For
    Next i
    If flag_Blank = False Then '空白なし
    End If
  Case "全ヵ所検査"
    For i = 7 To 36
      For j = 2 To 8 Step 2
        If flag_Blank = True Then Exit For
          If ws.Cells(i, j) = "" Then
            Set MoveTg = ws.Cells(i, j)
            flag_Blank = True
            Exit For
          End If
      Next j
    If flag_Blank = True Then Exit For
    Next i
    If flag_Blank = False Then '空白なし
    End If
  End Select
End Select
End Function
'キャスク.全ヵ所検査用
'ﾒﾓﾘ解放　2019.12.25
Public Function AlTerRange(TempRange As Range, Optional BeforDate As String = "") As Range
Dim flag_Blank As Boolean
Dim i As Long, j As Long
Dim TR As Long: TR = TempRange.row
Dim TC As Long: TC = TempRange.Column

Dim tempStr As String: tempStr = Maker(TempRange.Value)
Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet

  Select Case ws.Name
  Case "キャスク検査"
  UFCask.Cnt = "1" 'Application.Undo対策　2019.12.24
    Select Case True
    Case (TR >= 7 And TR <= 706) And (TC = 2 Or TC = 4 Or TC = 6 Or TC = 8)  '厚座標内_2021.07.09_TC = 8追加
        Select Case tempStr
        Case "DT" 'DT10002+0000150.712M
          Set AlTerRange = TempRange
          AlTerRange = Conv(AlTerRange.Value) '数値変換
        Case "#L" '#L395.8 #〆
          Call Special_Move_L(TempRange, BeforDate)
        Case Else
         Set AlTerRange = TempRange
         AlTerRange = "" 'Conv(AlTerRange.Value)
        End Select
    Case (TR >= 7 And TR <= 706) And (TC = 11 Or TC = 13 Or TC = 15 Or TC = 17 Or TC = 19 Or TC = 20)  '巾丈対角座標内
        Select Case tempStr
        Case "DT" 'DT10002+0000021.712M
          Call Special_Move_DT(TempRange, BeforDate)
        Case "#L" '#L395.8
          Set AlTerRange = TempRange
          AlTerRange = Conv(AlTerRange.Value) '数値変換
        Case Else
         Set AlTerRange = TempRange
         AlTerRange = "" ' Conv(AlTerRange.Value)
        End Select
    Case (TR >= 7 And TR <= 706) And (TC = 24 Or TC = 25 Or TC = 26 Or TC = 27 Or TC = 28) 'ﾗﾃﾗﾙﾎﾞｰ歪座標内
        Select Case tempStr
        Case "DT" 'DT10002+0000021.712M
          Call Special_Move_DT(TempRange, BeforDate)
        Case "#L" '#L395.8
          Call Special_Move_L(TempRange, BeforDate)
        Case Else
         Set AlTerRange = TempRange
         AlTerRange = Conv(AlTerRange.Value)
        End Select
    Case Else 'DT#Lの時に限りアクティブ移動
        Select Case tempStr
        Case "DT" 'DT10002+0000021.712M
          Sheets(ws.Name).Unprotect
            flag_Blank = False
            For i = 7 To 706
            Dim n As Long
            If UFCask.CountSt.Caption = "4" Then n = 8 '2021.07.09_検査箇所変動用追加
            If UFCask.CountSt.Caption = "3" Then n = 6
              For j = 2 To n Step 2
                If flag_Blank = True Then Exit For
                  If ws.Cells(i, j) = "" Then
                    Application.Undo
                    ws.Cells(i, j).Activate '空白欄ｱｸﾃｨﾌﾞ
                    AppActivate Application.Caption
                    flag_Blank = True
                    Exit For
                  End If
              Next j
            Next i
        Case "#L" '#L395.8
          Sheets(ws.Name).Unprotect
          flag_Blank = False
          For i = 7 To 706
            For j = 11 To 17 Step 2
              If flag_Blank = True Then Exit For
                If ws.Cells(i, j) = "" Then
                  Application.Undo
                  ws.Cells(i, j).Activate '空白欄ｱｸﾃｨﾌﾞ
                  AppActivate Application.Caption
                  flag_Blank = True
                  Exit For
                End If
            Next j
          If flag_Blank = True Then Exit For
          Next i
        Case Else
         Set AlTerRange = TempRange
         AlTerRange = Conv(AlTerRange.Value)
        End Select
    End Select
  Case "全ヵ所検査"
  UFMitsubishi.Cnt = "1" 'Application.Undo対策　2019.12.24
    Select Case True
    Case (TR >= 7 And TR <= 76) And (TC = 2 Or TC = 4 Or TC = 6 Or TC = 8)  '厚座標内
        Select Case tempStr
        Case "DT" 'DT10002+0000021.712M
          Set AlTerRange = TempRange
          AlTerRange = Conv(AlTerRange.Value) '数値変換
        Case "#L" '#L395.8
          Call Special_Move_L(TempRange, BeforDate)
        Case Else
         Set AlTerRange = TempRange
         AlTerRange = "" 'Conv(AlTerRange.Value)
        End Select
    Case (TR >= 7 And TR <= 76) And (TC = 11 Or TC = 13) And (UFMitsubishi.H494)  '特別用途
    '2023.6.20_修正
      Select Case tempStr
      Case "DT" 'DT10002+0000021.712M
        Set AlTerRange = TempRange
        AlTerRange = Conv(AlTerRange.Value) '数値変換
        AlTerRange.NumberFormatLocal = "0.00_);[赤](0.00)" 'ノギス下二けた表示
      Case Else
        Call Special_Move_L(TempRange, BeforDate)
      End Select
    Case (TR >= 7 And TR <= 76) And (TC = 11 Or TC = 13 Or TC = 15 Or TC = 17 Or TC = 19 Or TC = 20)  '巾丈対角座標内
    '2023.5.16_修正
        Select Case tempStr
        Case "DT" 'DT10002+0000021.712M
         Call Special_Move_DT(TempRange, BeforDate)
        Case "#L" '#L395.8
           Set AlTerRange = TempRange
           AlTerRange = Conv(AlTerRange.Value) '数値変換
           AlTerRange.NumberFormatLocal = "0.0_);[赤](0.0)" 'ノギス下二けた表示だった場合元に戻す
        End Select
    Case (TR >= 7 And TR <= 76) And (TC = 24 Or TC = 25 Or TC = 26)   '歪座標内
        Select Case tempStr
        Case "DT" 'DT10002+0000021.712M
          Call Special_Move_DT(TempRange, BeforDate)
        Case "#L" '#L395.8
          Call Special_Move_L(TempRange, BeforDate)
        Case Else
         Set AlTerRange = TempRange
         AlTerRange = Conv(AlTerRange.Value)
        End Select
    Case (TR >= 7 And TR <= 76) And (TC = 23)   'ラテラルボー
        Select Case tempStr
        Case "DT" 'DT10002+0000021.712M
          Call Special_Move_DT(TempRange, BeforDate)
        Case "#L" '#L395.8
          Call Special_Move_L(TempRange, BeforDate)
        End Select
    Case Else 'DT#Lの時に限りアクティブ移動
        Select Case tempStr
        Case "DT" 'DT10002+0000021.712M
          Sheets(ws.Name).Unprotect
            flag_Blank = False
            For i = 7 To 76
              For j = 2 To 8 Step 2
                If flag_Blank = True Then Exit For
                  If ws.Cells(i, j) = "" Then
                    Application.Undo
                    ws.Cells(i, j).Activate '空白欄ｱｸﾃｨﾌﾞ
                    AppActivate Application.Caption
                    flag_Blank = True
                    Exit For
                  End If
              Next j
            Next i
        Case "#L" '#L395.8
          Sheets(ws.Name).Unprotect
          flag_Blank = False
          For i = 7 To 76
            For j = 11 To 17 Step 2
              If flag_Blank = True Then Exit For
                If ws.Cells(i, j) = "" Then
                  Application.Undo
                  ws.Cells(i, j).Activate '空白欄ｱｸﾃｨﾌﾞ
                  AppActivate Application.Caption
                  flag_Blank = True
                  Exit For
                End If
            Next j
          If flag_Blank = True Then Exit For
          Next i
        Case Else
         Set AlTerRange = TempRange
         AlTerRange = Conv(AlTerRange.Value)
        End Select
    End Select
  
  Case "H494検査"
  UFMitsubishi.Cnt = "1" 'Application.Undo対策　2019.12.24
    Select Case True
    Case (TR >= 7 And TR <= 16) And (TC = 2 Or TC = 4 Or TC = 6 Or TC = 8)  '厚座標内
        Select Case tempStr
        Case "DT" 'DT10002+0000021.712M
          Set AlTerRange = TempRange
          AlTerRange = Conv(AlTerRange.Value) '数値変換
        Case "#L" '#L395.8
          Call Special_Move_L(TempRange, BeforDate)
        Case Else
         Set AlTerRange = TempRange
         AlTerRange = "" 'Conv(AlTerRange.Value)
        End Select
    Case (TR >= 7 And TR <= 16) And (TC = 11 Or TC = 13 Or TC = 15)  '特別用途
      Select Case tempStr
      Case "DT" 'DT10002+0000021.712M
        Set AlTerRange = TempRange
        AlTerRange = Conv(AlTerRange.Value) '数値変換
        AlTerRange.NumberFormatLocal = "0.00_);[赤](0.00)" 'ノギス下二けた表示
      Case Else
        Call Special_Move_L(TempRange, BeforDate)
      End Select
    Case (TR >= 7 And TR <= 16) And (TC = 17 Or TC = 19 Or TC = 21 Or TC = 22)  '丈対角座標内
        Select Case tempStr
        Case "DT" 'DT10002+0000021.712M
         Call Special_Move_DT(TempRange, BeforDate)
        Case "#L" '#L395.8
           Set AlTerRange = TempRange
           AlTerRange = Conv(AlTerRange.Value) '数値変換
           AlTerRange.NumberFormatLocal = "0.0_);[赤](0.0)" 'ノギス下二けた表示だった場合元に戻す
        End Select
    Case (TR >= 7 And TR <= 16) And (TC = 26 Or TC = 27)   '歪座標内
        Select Case tempStr
        Case "DT" 'DT10002+0000021.712M
          Call Special_Move_DT(TempRange, BeforDate)
        Case "#L" '#L395.8
          Call Special_Move_L(TempRange, BeforDate)
        Case Else
         Set AlTerRange = TempRange
         AlTerRange = Conv(AlTerRange.Value)
        End Select
    Case (TR >= 7 And TR <= 16) And (TC = 25)   'ラテラルボー
        Select Case tempStr
        Case "DT" 'DT10002+0000021.712M
          Call Special_Move_DT(TempRange, BeforDate)
        Case "#L" '#L395.8
          Call Special_Move_L(TempRange, BeforDate)
        End Select
    Case Else 'DT#Lの時に限りアクティブ移動
        Select Case tempStr
        Case "DT" 'DT10002+0000021.712M
          Sheets(ws.Name).Unprotect
            flag_Blank = False
            For i = 7 To 16
              For j = 2 To 8 Step 2
                If flag_Blank = True Then Exit For
                  If ws.Cells(i, j) = "" Then
                    Application.Undo
                    ws.Cells(i, j).Activate '空白欄ｱｸﾃｨﾌﾞ
                    AppActivate Application.Caption
                    flag_Blank = True
                    Exit For
                  End If
              Next j
            Next i
        Case "#L" '#L395.8
          Sheets(ws.Name).Unprotect
          flag_Blank = False
          For i = 7 To 16
            For j = 11 To 19 Step 2
              If flag_Blank = True Then Exit For
                If ws.Cells(i, j) = "" Then
                  Application.Undo
                  ws.Cells(i, j).Activate '空白欄ｱｸﾃｨﾌﾞ
                  AppActivate Application.Caption
                  flag_Blank = True
                  Exit For
                End If
            Next j
          If flag_Blank = True Then Exit For
          Next i
        Case Else
         Set AlTerRange = TempRange
         AlTerRange = Conv(AlTerRange.Value)
        End Select
    End Select
  End Select
'MsgBox ActiveCell.Address
End Function

Public Sub キャスクコピー() '2019.05.26
Dim sh As String: sh = "キャスク検査"
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
  ThisWorkbook.Sheets("キャスク原紙").Visible = True
  ThisWorkbook.Sheets("キャスク原紙").Copy before:=Worksheets(1)
  Set ws = ActiveSheet ' コピーしたシートがアクティブ
  ' そのシートに対して編集
  ws.Name = sh
  With ws.Tab
      .Color = &H660066
      .TintAndShade = 0
  End With
  ActiveWindow.Zoom = 90
  ThisWorkbook.Sheets("キャスク原紙").Visible = False
Application.ScreenUpdating = True
End Sub

Public Sub キャスク用最大最小() '2019.05.26
Dim i As Long, a As Long
Dim n As Long, k As Long
Dim C2 As Range, C4 As Range, C6 As Range
Dim C11 As Range, C13 As Range, C15 As Range, C17 As Range
Dim MAXatu As Single, MINatu As Single
Dim MAXhaba As Single, MINhaba As Single
Dim MAXtake As Single, MINtake As Single
Dim shName As String
n = val(UFCask.MINText) + 6
k = val(UFCask.MAXText) + 6
With ThisWorkbook.Sheets("キャスク検査") '列毎の設定
    Set C2 = .Range(.Cells(n, 2), (.Cells(k, 2)))
    Set C4 = .Range(.Cells(n, 4), (.Cells(k, 4)))
    Set C6 = .Range(.Cells(n, 6), (.Cells(k, 6)))
    Set C8 = .Range(.Cells(n, 8), (.Cells(k, 8))) '2021.07.09_追加
    Set C11 = .Range(.Cells(n, 11), (.Cells(k, 11)))
    Set C13 = .Range(.Cells(n, 13), (.Cells(k, 13)))
    Set C15 = .Range(.Cells(n, 15), (.Cells(k, 15)))
    Set C17 = .Range(.Cells(n, 17), (.Cells(k, 17)))
End With
With Application.WorksheetFunction
  If UFCask.CountSt.Caption = "4" Then '4列での最大最少(厚)_2021.07.09
    MAXatu = .Max(.Max(C2), .Max(C4), .Max(C6), .Max(C8))
    MINatu = .Min(.Min(C2), .Min(C4), .Min(C6), .Min(C8))
  ElseIf UFCask.CountSt.Caption = "3" Then '3列での最大最少(厚)
    MAXatu = .Max(.Max(C2), .Max(C4), .Max(C6))
    MINatu = .Min(.Min(C2), .Min(C4), .Min(C6))
  End If
'2列での最大最少 (巾)
    MAXhaba = .Max(.Max(C11), .Max(C13))
    MINhaba = .Min(.Min(C11), .Min(C13))
'2列での最大最少 (丈)
    MAXtake = .Max(.Max(C15), .Max(C17))
    MINtake = .Min(.Min(C15), .Min(C17))
End With
With ThisWorkbook.Sheets("キャスク検査")
.Unprotect
'リセット
.Cells(3, 22) = ""
.Cells(4, 22) = ""
.Cells(3, 23) = ""
.Cells(4, 23) = ""
.Cells(3, 24) = ""
.Cells(4, 24) = ""
.Range("B7:AB706").Font.Underline = xlUnderlineStyleNone
'指定の数字にアンダーバー
    For i = n To k
    Dim x As Long
    If UFCask.CountSt.Caption = "4" Then x = 8 '2021.07.09_検査箇所変動用追加
    If UFCask.CountSt.Caption = "3" Then x = 6
      For a = 2 To x Step 2
        If .Cells(i, a) = MAXatu Then '同じ数字の時
           .Cells(i, a).Font.Underline = xlUnderlineStyleSingle
           .Cells(3, 22) = MAXatu
        Exit For
        End If
      Next a
    Next i
    For i = x To k
      For a = 2 To n Step 2
        If .Cells(i, a) = MINatu Then '同じ数字の時
           .Cells(i, a).Font.Underline = xlUnderlineStyleSingle
           .Cells(4, 22) = MINatu
        Exit For
        End If
      Next a
    Next i
 '-----------------------------------------------------巾
    For i = n To k
      For a = 11 To 13 Step 2
        If .Cells(i, a) = MAXhaba Then '同じ数字の時
           .Cells(i, a).Font.Underline = xlUnderlineStyleSingle
           .Cells(3, 23) = MAXhaba
        Exit For
        End If
      Next a
    Next i
    For i = n To k
      For a = 11 To 13 Step 2
        If .Cells(i, a) = MINhaba Then '同じ数字の時
           .Cells(i, a).Font.Underline = xlUnderlineStyleSingle
           .Cells(4, 23) = MINhaba
        Exit For
        End If
      Next a
    Next i
 '-----------------------------------------------------丈
    For i = n To k
      For a = 15 To 17 Step 2
        If .Cells(i, a) = MAXtake Then '同じ数字の時
           .Cells(i, a).Font.Underline = xlUnderlineStyleSingle
           .Cells(3, 24) = MAXtake
        Exit For
        End If
      Next a
    Next i
    For i = n To k
      For a = 15 To 17 Step 2
        If .Cells(i, a) = MINtake Then '同じ数字の時
           .Cells(i, a).Font.Underline = xlUnderlineStyleSingle
           .Cells(4, 24) = MINtake
        Exit For
        End If
      Next a
    Next i
End With
End Sub

Public Sub キャスク専用保存() '2019.05.26
Dim ws As Worksheet, flag As Boolean, Flag2 As Boolean
Dim sh As Worksheet
Dim fCount As Integer, Lcount As Integer
Dim Lotname As String, MonthName As String, DayLotName As String
Dim SetPath As String, MonthPath As String
Dim wkArray()
Dim idx, Cnt As Integer
Dim BK As Workbook
Application.Visible = True
Application.ScreenUpdating = False
ThisWorkbook.Activate
'*************************************************シートの有無
     For Each ws In ThisWorkbook.Worksheets
      If ws.Name = "キャスク保存" Then
       flag = True
      End If
      If ws.Name = "品証ｺﾋﾟ-用" Then
       Flag2 = True
      End If
     Next ws
'************************************************保存シートの初期化
  Application.DisplayAlerts = False
  If flag Then ThisWorkbook.Sheets("キャスク保存").Delete
  If Flag2 Then ThisWorkbook.Sheets("品証ｺﾋﾟ-用").Delete
  Application.DisplayAlerts = True
  '------------------------------------------------追加シート処理
    ThisWorkbook.Sheets.Add before:=Worksheets(1) 'Sheets(Sheets.Count) '新たにシートを追加する
    ActiveSheet.Name = "キャスク保存"
    With ThisWorkbook.Sheets("キャスク保存").Tab
        .Color = &H99FFFF
        .TintAndShade = 0
    End With
    ActiveWindow.Zoom = 100
  '------------------------------------------------コピー処理　保存閲覧用
    ThisWorkbook.Sheets("キャスク検査").Unprotect
    ThisWorkbook.Sheets("キャスク検査").Cells.Copy
    ThisWorkbook.Sheets("キャスク保存").Cells.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    ThisWorkbook.Sheets("キャスク保存").Cells.Font.Underline = xlUnderlineStyleNone
    With ThisWorkbook.Sheets("キャスク保存").Cells.Interior '黒塗りをなくす
      .Pattern = xlNone
      .TintAndShade = 0
      .PatternTintAndShade = 0
    End With
    With ThisWorkbook.Sheets("キャスク保存")
      .Cells(1, 25) = "検査番号"
      .Cells(1, 26) = UFCask.KenNo
      .Cells(2, 25) = "検査員"
      .Cells(3, 25) = "ｺﾝﾍﾞｯｸｽ"
      .Cells(4, 25) = "ﾏｲｸﾛ"
      .Cells(2, 26) = UFCask.Kname
      .Cells(3, 26) = UFCask.Cno
      .Cells(4, 26) = UFCask.Mno & "  S/N" & UFCask.SeriNo
      Dim SokuSt As String
      If val(.Cells(3, 2)) >= 100# Then SokuSt = "ﾉｷﾞｽ"
      If val(.Cells(3, 2)) < 100# Then SokuSt = "ﾏｲｸﾛ"
      .Cells(2, 28) = "ｾﾞﾛ校正:" & SokuSt
      .Cells(3, 28) = "ｾﾞﾛ校正:" & "ｺﾝﾍﾞｯｸｽ"
      .Cells(3, 30) = "1000mm校正:" & "ｺﾝﾍﾞｯｸｽ"
      .Cells(2, 29) = UFCask.Micro
      .Cells(3, 29) = UFCask.Convex
      .Cells(3, 31) = UFCask.Convex_1000
      
      .Range(.Cells(1, 25), .Cells(4, 31)).Font.Size = 8
      .Range(.Cells(1, 25), .Cells(4, 31)).HorizontalAlignment = xlLeft
      .Range(.Cells(1, 25), .Cells(4, 31)).VerticalAlignment = xlCenter
      .Cells(1, 3) = "～"
      .Cells(1, 4) = Format$(Now, "yyyy/mm/dd h:mm")
      .Range(.Cells(1, 3), .Cells(1, 4)).Font.Size = 5
      .Cells.EntireColumn.AutoFit
      .Range("B:B,K:K,Z:Z").ColumnWidth = 8.5
      .Range("V:V").ColumnWidth = 18
      .Cells.Locked = True 'セルロック
      .Cells.FormulaHidden = False
      '.Protect Password:=""
      .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'シート保護
    End With
'------------------------------------------------品証用に整理する　品証コピーペースト用
'追加シート処理
  If UFCask.Cno = "" Or UFCask.Mno = "" Or UFCask.Kname = "" Or UFCask.SeriNo = "" Then
   MsgBox "フォーム上で必要欄を記入してください"
   Exit Sub
  Else
  ''2026.02.20_梱包板Noの保存追加
        Dim strLot As String: strLot = UFCask.LotBox
        Dim strKen As String: strKen = UFCask.KenNo
        Dim Temp As Variant: Temp = GetCaskData(strLot, strKen)
  
       ThisWorkbook.Sheets.Add before:=Worksheets(1) 'Sheets(Sheets.Count) '新たにシートを追加する
       ActiveSheet.Name = "品証ｺﾋﾟ-用"
       With ThisWorkbook.Sheets("品証ｺﾋﾟ-用").Tab
           .Color = 128
           .TintAndShade = 0
       End With
       ActiveWindow.Zoom = 100
     'コピー処理
        n = ThisWorkbook.Sheets("キャスク保存").Cells(Rows.Count, 2).End(xlUp).row
        ThisWorkbook.Sheets("キャスク保存").Range("A6:AB" & n).Copy _
        Destination:=ThisWorkbook.Sheets("品証ｺﾋﾟ-用").Cells(1, 4)
     '不要列削除
        ThisWorkbook.Sheets("品証ｺﾋﾟ-用").Range("F:F,H:H,J:J,L:M,O:O,Q:Q,S:S,U:U,V:X,Z:Z").Delete '2021.07.09_厚4の追加により削除部分変更
     '必要データ添付
        With ThisWorkbook.Sheets("品証ｺﾋﾟ-用")
            .Cells(1, 1) = "LotNo"
            .Cells(1, 2) = ThisWorkbook.Sheets("キャスク保存").Cells(2, 4)
            .Cells(2, 1) = "検査No"
            .Cells(2, 2) = ThisWorkbook.Sheets("キャスク保存").Cells(1, 26)
            .Cells(3, 1) = "Size"
            .Cells(3, 2) = ThisWorkbook.Sheets("キャスク保存").Cells(3, 2) & "×" & Sheets("キャスク保存").Cells(3, 4) & "×" & Sheets("キャスク保存").Cells(3, 6)
            .Cells(4, 1) = "ｺﾝﾍﾞｯｸｽNo"
            .Cells(4, 2) = UFCask.Cno
            .Cells(5, 1) = "ﾏｲｸﾛNo"
            .Cells(5, 2) = UFCask.Mno & " S/N" & UFCask.SeriNo
            .Cells(6, 1) = "検査員"
            .Cells(6, 2) = UFCask.Kname
            .Cells(1, 4) = "板番号"
            .Cells(7, 1) = "作業日"
            .Cells(7, 2) = Format$(ThisWorkbook.Sheets("キャスク保存").Cells(1, 2), "yyyy/mm/dd  h:mm") & "～" & Format$(ThisWorkbook.Sheets("キャスク保存").Cells(1, 4), "yyyy/mm/dd  h:mm")
            .Cells(8, 1) = "ｾﾞﾛ校正:" & SokuSt
            .Cells(8, 2) = UFCask.Micro
            .Cells(9, 1) = "ｾﾞﾛ校正:" & "ｺﾝﾍﾞｯｸｽ"
            .Cells(9, 2) = UFCask.Convex
            .Cells(10, 1) = "1000mm校正:" & "ｺﾝﾍﾞｯｸｽ"
            .Cells(10, 2) = UFCask.Convex_1000
            
            If IsEmpty(Temp) Then GoTo CaskLine
        
            For i = LBound(Temp, 1) To UBound(Temp, 1)
                .Cells(12 + i, 2).Value = "梱包No" & i + 1 & ":" & Temp(i, 3) & "～" & Temp(i, 4)
            Next i
        End With
       
CaskLine:
       
        With ThisWorkbook.Sheets("品証ｺﾋﾟ-用").Cells
          .Font.Name = "Meiryo UI"
          .Font.Size = "10"
          .Font.Underline = xlUnderlineStyleNone
          .Borders(xlDiagonalDown).LineStyle = xlNone
          .Borders(xlDiagonalUp).LineStyle = xlNone
          .Borders(xlEdgeLeft).LineStyle = xlNone
          .Borders(xlEdgeTop).LineStyle = xlNone
          .Borders(xlEdgeBottom).LineStyle = xlNone
          .Borders(xlEdgeRight).LineStyle = xlNone
          .Borders(xlInsideVertical).LineStyle = xlNone
          .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With
       ThisWorkbook.Sheets("品証ｺﾋﾟ-用").Cells.EntireColumn.AutoFit
       ThisWorkbook.Sheets("品証ｺﾋﾟ-用").Cells.Locked = True
       ThisWorkbook.Sheets("品証ｺﾋﾟ-用").Cells.FormulaHidden = False
'       ThisWorkbook.Sheets("品証ｺﾋﾟ-用").Protect Password:=""
       ThisWorkbook.Sheets("品証ｺﾋﾟ-用").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'シート保護
  'キャスク検査のコードを消して必要データのみのシートをｻｰﾊﾞｰ保存する
      Call DisplaySheets
      For idx = 1 To Worksheets.Count
        If (ThisWorkbook.Worksheets(idx).Name) = "キャスク保存" Or (ThisWorkbook.Worksheets(idx).Name) = "品証ｺﾋﾟ-用" Then
           Cnt = Cnt + 1
           ReDim Preserve wkArray(1 To Cnt)
           wkArray(Cnt) = ThisWorkbook.Worksheets(idx).Name
        End If
      Next idx
      If Cnt = 0 Then
        Exit Sub '保存シートがないとエラーになる2018.12.19
      Else
        ThisWorkbook.Worksheets(wkArray).Copy
      End If
      Call HiddenSheets
      ThisWorkbook.Sheets("キャスク検査").Visible = True
     '------------------------------------------------2019.08.28 アクティブがうつって保存できないエラーが出た
      Set BK = ActiveWorkbook
      '************************************月毎にフォルダを作成してその中に格納する2018.12.19
        MonthName = Format(Date, "yyyy.mm")
        MonthPath = "C:\検査履歴" & "\" & MonthName '-----------------------ﾛｰｶﾙ保存_2021.01.25
        Lotname = ThisWorkbook.Sheets("キャスク保存").Cells(2, 4) 'ロットNO
        DayLotName = Format(Date, "yyyymmdd") & "_" & ThisWorkbook.Sheets("キャスク保存").Cells(2, 4) '日付とロットNOのﾌｧｲﾙを作成　2019.08.22
        SetPath = MonthPath & "\" & DayLotName
        
        Call CreateFolder(SetPath) 'ﾌｫﾙﾀﾞを作成
        
        Dim FilePath As String:     FilePath = SetPath & "\" & DayLotName & ".xlsx"
        
        If Dir(FilePath, vbDirectory) = "" Then
            BK.SaveAs (FilePath), FileFormat:=xlOpenXMLWorkbook, Password:=Lotname & "." & Format(Date, "yyyy.mm.dd")   'pass LotNo.西暦.月.日
            SetAttr (FilePath), vbReadOnly
        Else
          For fCount = 2 To 999
            If Dir(SetPath & "\" & DayLotName & "." & Right("00" & CStr(fCount), 3) & ".xlsx", vbDirectory) = "" Then  '同じ場合副番を付けて999まで保存する
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
    ThisWorkbook.Activate
    Application.DisplayAlerts = False
      If WsCheck("品証ｺﾋﾟ-用") Then ThisWorkbook.Sheets("品証ｺﾋﾟ-用").Delete
      If WsCheck("キャスク保存") Then ThisWorkbook.Sheets("キャスク保存").Delete
      If WsCheck("ｷｬｽｸ作業中保存") Then ThisWorkbook.Sheets("ｷｬｽｸ作業中保存").Delete
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("キャスク検査").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True  '保護設定
      '--------------------ﾌｧｲﾙｺﾋﾟｰ_処理ｽﾋﾟｰﾄﾞ速い_2021.01.25
      
      MonthName = Format(Date, "yyyy.mm")
      Dim SrvPath As String
      Select Case UFCask.LineName
      Case "Test":    SrvPath = HistoryPath & "キャスク_履歴" & "\" & MonthName  ' HistoryPath & "Test\Test_履歴\キャスク_履歴"
      Case Else:      SrvPath = HistoryPath & "キャスク_履歴" & "\" & MonthName
      End Select
      
      Call CreateFolder(SrvPath) 'ﾌｫﾙﾀﾞを作成
      
      Dim NewFiles As String: NewFiles = NewFileName(SetPath & "\")
      FileCopy SetPath & "\" & NewFiles, SrvPath & "\" & NewFiles
      
      '品証用に別に保存する　\\nlmsrvngy03\NPS\全社研修会トレーナー教育\第36回NPSトレーナー教育（名古屋）2017年\梱包実績の見える化\
      '-------------------------------------------------------------------------------------------
      Dim Kouzyou As String: Kouzyou = FactoryyPath & "【梱包】キャスク\" '2021.01.22_パス変更
      
      Dim KouzyouPath As String
      Select Case UFCask.LineName '"\\nlmsrvngy03\NPS\全社研修会トレーナー教育\第36回NPSトレーナー教育（名古屋）2017年\梱包実績の見える化\キャスク_履歴\"
      Case "Test":    KouzyouPath = HistoryPath & "Test\Test(工場共有)_履歴\キャスク_履歴"
      Case Else:      KouzyouPath = Kouzyou & MonthName
      End Select
      
      Call CreateFolder(KouzyouPath) 'ﾌｫﾙﾀﾞを作成
      
      FileCopy SetPath & "\" & NewFiles, KouzyouPath & "\" & NewFiles
      '-------------------------------------------------------------------------------------------
    Application.ScreenUpdating = True
  End If
Application.Visible = True
End Sub
Public Sub JustinCase_KySave() '2019.11.23
If WsCheck("キャスク検査") = False Then Exit Sub
Dim ws As Worksheet, flag As Boolean
Dim sh As Worksheet
Dim fCount As Integer, Lcount As Integer
Dim Lotname As String, MonthName As String, DayLotName As String
Dim SetPath As String, MonthPath As String
Dim wkArray()
Dim idx, Cnt As Integer
Dim BK As Workbook
Application.ScreenUpdating = False
ThisWorkbook.Activate
'*************************************************シートの有無
    Call AddWS("キャスク保存")
    With ActiveWorkbook.Sheets("キャスク保存").Tab
      .Color = &H99FFFF
      .TintAndShade = 0
    End With
    ActiveWindow.Zoom = 90
  '------------------------------------------------コピー処理　保存閲覧用
    ThisWorkbook.Sheets("キャスク検査").Unprotect
    ThisWorkbook.Sheets("キャスク検査").Cells.Copy
    Worksheets("キャスク保存").Cells.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    With Worksheets("キャスク保存").Cells.Interior '黒塗りをなくす
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With ThisWorkbook.Sheets("キャスク保存")
      .Cells(1, 25) = "検査番号"
      .Cells(1, 26) = UFCask.KenNo
      .Cells(2, 25) = "検査員"
      .Cells(3, 25) = "ｺﾝﾍﾞｯｸｽ"
      .Cells(4, 25) = "ﾏｲｸﾛ"
      .Cells(2, 26) = UFCask.Kname
      .Cells(3, 26) = UFCask.Cno
      .Cells(4, 26) = UFCask.Mno & "  S/N" & UFCask.SeriNo
       Range(.Cells(1, 25), .Cells(4, 26)).Font.Size = 8
       Range(.Cells(1, 25), .Cells(4, 26)).HorizontalAlignment = xlLeft
       Range(.Cells(1, 25), .Cells(4, 26)).VerticalAlignment = xlCenter
      .Cells(1, 3) = "～"
      .Cells(1, 4) = Format$(Now, "yyyy/mm/dd h:mm")
       Range(.Cells(1, 3), .Cells(1, 4)).Font.Size = 5
      .Cells.EntireColumn.AutoFit
      .Range("B:B,K:K,Z:Z").ColumnWidth = 8.5
      .Range("V:V").ColumnWidth = 18
      .Cells.Locked = True 'セルロック
      .Cells.FormulaHidden = False
      '.Protect Password:=""
      .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'シート保護
    End With
  'キャスク検査のコードを消して必要データのみのシートをｻｰﾊﾞｰ保存する
     ThisWorkbook.Sheets("キャスク保存").Visible = True
     For idx = 1 To Worksheets.Count
      If (ThisWorkbook.Worksheets(idx).Name) = "キャスク保存" Or (ThisWorkbook.Worksheets(idx).Name) = "キャスク検査" Then
        Cnt = Cnt + 1
        ReDim Preserve wkArray(1 To Cnt)
        wkArray(Cnt) = ThisWorkbook.Worksheets(idx).Name
      End If
     Next idx
     If Cnt = 0 Then
     Exit Sub '保存シートがないとエラーになる2018.12.19
     Else
     ThisWorkbook.Worksheets(wkArray).Copy
     End If
    '------------------------------------------------2019.08.28 アクティブがうつって保存できないエラーが出た
     Set BK = ActiveWorkbook
    '------------------------------------------------
    Application.DisplayAlerts = False
      '************************************月毎にフォルダを作成してその中に格納する2018.12.19
        MonthName = Format(Date, "yyyy.mm")
        Select Case UFCask.LineName
        Case "Test":    MonthPath = HistoryPath & "Test\Test_履歴\JustinCaseSave"  '"C:\Users\NLM Administrator\Desktop\保存用ﾌｫﾙﾀﾞ\JustinCaseSave\" & MonthName
        Case Else:      MonthPath = HistoryPath & "JustinCaseSave"
        End Select
        Lotname = ThisWorkbook.Sheets("キャスク保存").Cells(2, 4) 'ロットNO
        DayLotName = Format(Date, "yyyymmdd") & "_" & Sheets("キャスク保存").Cells(2, 4) '日付とロットNOのﾌｧｲﾙを作成　2019.08.22
        SetPath = MonthPath & "\" & DayLotName
        
        Call CreateFolder(SetPath) 'ﾌｫﾙﾀﾞを作成
        
        Dim FilePath As String:     FilePath = SetPath & "\" & DayLotName & ".xlsx"
        
        If Dir(FilePath, vbDirectory) = "" Then
            BK.SaveAs (FilePath), FileFormat:=xlOpenXMLWorkbook, Password:=Lotname & "." & Format(Date, "yyyy.mm.dd")   'pass LotNo.西暦.月.日
            SetAttr (FilePath), vbReadOnly
        Else
          For fCount = 2 To 999
            If Dir(SetPath & "\" & DayLotName & "." & Right("00" & CStr(fCount), 3) & ".xlsx", vbDirectory) = "" Then '同じ場合副番を付けて999まで保存する
              BK.SaveAs (SetPath & "\" & DayLotName & "." & Right("00" & CStr(fCount), 3) & ".xlsx"), FileFormat:=xlOpenXMLWorkbook, Password:=Lotname & "." & Format(Date, "yyyy.mm.dd")   'pass LotNo.西暦.月.日
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
    ThisWorkbook.Activate
    ThisWorkbook.Sheets("キャスク検査").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True  '保護設定
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Public Sub ClearHeaderFooter(ws As Worksheet)
On Error Resume Next
    ' 一時的に保護解除（パスワードがあるなら "" の中に入れてください）
    ws.Unprotect
    
    Application.PrintCommunication = False
    With ws.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
    End With
    Application.PrintCommunication = True
End Sub

Public Sub キャスク印刷範囲()
    Dim SheetsName As String: SheetsName = "キャスク検査"
    Dim ws As Worksheet, sh As Worksheet
    Dim flag As Boolean
    Dim n As Long
    
    '--- 1. シートの存在確認 ---
    If WsCheck(SheetsName) = False Then Exit Sub
    
    Set sh = ThisWorkbook.Sheets(SheetsName)
    Call ClearHeaderFooter(sh)
    
    '--- 2. 印刷範囲と書式の設定 ---
    Application.PrintCommunication = False
    
    n = sh.Cells(sh.Rows.Count, 2).End(xlUp).row
    sh.PageSetup.PrintArea = "$A$6:$AB$" & n ' 印刷範囲設定
    
    ' 書式クリア（範囲を固定で指定）
    With sh.Range("B7:AB706")
        .Interior.Color = RGB(255, 255, 255)
        .Font.Underline = xlUnderlineStyleNone
    End With
    
    '--- 3. データ取得とフッター文字列作成 ---
    Dim strLot As String: strLot = UFCask.LotBox
    Dim strKen As String: strKen = UFCask.KenNo
    Dim Temp As Variant:  Temp = GetCaskData(strLot, strKen)
    Dim footerText As String
    Dim i As Long
    
    '2026.02.20_梱包板Noの保存追加
    If Not IsEmpty(Temp) Then
        For i = LBound(Temp, 1) To UBound(Temp, 1)
            If footerText <> "" Then footerText = footerText & "  "
            footerText = footerText & "梱包No" & i + 1 & "_" & Temp(i, 3) & "～" & Temp(i, 4)
        Next i
    End If

    '--- 4. 数値データ等の取得 ---
    Dim x As Double, y As Double, z As Double
    Dim Am As Double, Ap As Double, Hm As Double, Hp As Double, Tm As Double, Tp As Double
    Dim SokuSt As String
    
    With sh
        x = val(.Cells(3, 2)):  y = val(.Cells(3, 4)):  z = val(.Cells(3, 6))
        Am = val(.Cells(2, 13)): Ap = val(.Cells(2, 15))
        Hm = val(.Cells(3, 13)): Hp = val(.Cells(3, 15))
        Tm = val(.Cells(4, 13)): Tp = val(.Cells(4, 15))
        SokuSt = IIf(x >= 100, "ﾉｷﾞｽ", "ﾏｲｸﾛ")
    End With

    '--- 5. ページ設定（ヘッダー・フッター） ---
    With sh.PageSetup
        ' ヘッダー・フッターのクリア
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        
        ' 左ヘッダー
        .LeftHeader = "&""Meiryo UI""&8" & "*" & sh.Cells(2, 4) & "*" & " 検査員：" & UFCask.Kname & " " & vbCr & _
                      Format(x, "0.000") & "×" & Format(y, "0.0") & "×" & Format(z, "0.0") & vbCr & _
                      "*ｾﾞﾛ校正_" & SokuSt & ":" & Format(sh.Cells(2, 29), "0.000") & _
                      " *ｾﾞﾛ校正_ｺﾝﾍﾞｯｸｽ:" & Format(sh.Cells(3, 29), "0.0") & _
                      " *1000mm校正_ｺﾝﾍﾞｯｸｽ:" & Format(sh.Cells(3, 31), "0.0")
        
        ' 右ヘッダー（公差情報）
        .RightHeader = "&""Meiryo UI""&8" & "*厚公差：" & Format((x - Am), "0.000") & "～" & Format((x + Ap), "0.000") & vbCr & _
                       "*巾公差：" & Format((y - Hm), "0.0") & "～" & Format((y + Hp), "0.0") & vbCr & _
                       "*丈公差：" & Format((z - Tm), "0.0") & "～" & Format((z + Tp), "0.0")
        
        ' フッター設定
        .LeftFooter = "&""Meiryo UI""&10" & footerText
        .RightFooter = "&""Meiryo UI""&8" & "*" & Format(sh.Cells(1, 2), "yyyy/mm/dd h:mm") & "～" & Format(Now, "yyyy/mm/dd h:mm") & " 印刷"
        
        ' 余白・レイアウト設定
        .HeaderMargin = Application.InchesToPoints(0.25)
        .FooterMargin = Application.InchesToPoints(0.2)
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0.8)
        .BottomMargin = Application.InchesToPoints(0.35)
        .CenterHorizontally = True
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
    Application.PrintCommunication = True
End Sub

Public Function HizumiJ(Ter As Range, kou As String) As Boolean
If IsNumeric(Ter.Value) = False Then Exit Function
If kou = "0" Or kou = "" Then Exit Function
If Ter.Value = 0 Then HizumiJ = True '合格
If val(kou) >= Abs(val(Ter.Value)) Then
 HizumiJ = True '○
ElseIf val(kou) < Abs(val(Ter.Value)) Then
 HizumiJ = False '×
End If
End Function
Public Function LateralJ(Ter As Range, kou As String) As Boolean '絶対値評価
If IsNumeric(Ter.Value) = False Then Exit Function
If kou = "0" Or kou = "" Then Exit Function
If Ter.Value = 0 Then LateralJ = True '合格
If val(kou) >= Abs(val(Ter.Value)) Then
 LateralJ = True '○
ElseIf val(kou) < Abs(val(Ter.Value)) Then
 LateralJ = False '×
End If
End Function

Public Function KenJudg(Ter As Range, KP As Double, KM As Double) As Boolean
If IsNumeric(Ter.Value) = True Then
  If ((Ter.Value) < KM) Or ((Ter.Value) > KP) Then
    KenJudg = False '×
  ElseIf ((Ter.Value) >= KM) And ((Ter.Value) <= KP) Then
    KenJudg = True '○
  End If
End If
End Function
Public Function TaiJudg(Ter As Range, kou As Single) As Boolean
Dim Tai As Single
Dim TR As Long: TR = Ter.row
Dim TC As Long: TC = Ter.Column
Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
Select Case ws.Name
Case "キャスク検査"
  Select Case True
  Case (TR >= 7 And TR <= 706) And (TC = 19 Or TC = 20)  '指定座標内
    If (IsNumeric(ws.Cells(TR, 19)) And IsNumeric(ws.Cells(TR, 20))) And (ws.Cells(TR, 19) <> "" And ws.Cells(TR, 20) <> "") Then 'この条件から外れるとfalse
    Tai = Abs(val(ws.Cells(TR, 19).Value) - val(ws.Cells(TR, 20).Value))
      If kou >= Tai Then
       TaiJudg = True
      ElseIf kou < Tai Then
       TaiJudg = False
      End If
    End If
  End Select
Case "全ヵ所検査"
  Select Case True
  Case (TR >= 7 And TR <= 76) And (TC = 19 Or TC = 20)  '指定座標内
    If (IsNumeric(ws.Cells(TR, 19)) And IsNumeric(ws.Cells(TR, 20))) And (ws.Cells(TR, 19) <> "" And ws.Cells(TR, 20) <> "") Then 'この条件から外れるとfalse
    Tai = Abs(val(ws.Cells(TR, 19).Value) - val(ws.Cells(TR, 20).Value))
      If kou >= Tai Then
       TaiJudg = True
       'MsgBox kou
      ElseIf kou < Tai Then
       TaiJudg = False
      End If
    End If
  End Select
Case "H494検査"
  Select Case True
  Case (TR >= 7 And TR <= 16) And (TC = 21 Or TC = 22)  '指定座標内
    If (IsNumeric(ws.Cells(TR, 21)) And IsNumeric(ws.Cells(TR, 22))) And (ws.Cells(TR, 21) <> "" And ws.Cells(TR, 22) <> "") Then 'この条件から外れるとfalse
    Tai = Abs(val(ws.Cells(TR, 21).Value) - val(ws.Cells(TR, 22).Value))
      If kou >= Tai Then
       TaiJudg = True
       'MsgBox kou
      ElseIf kou < Tai Then
       TaiJudg = False
      End If
    End If
  End Select
End Select
End Function

Public Sub ループ判定() '2019.05.26
 Dim a As Long, i As Long, b As Long, C As Long
Application.ScreenUpdating = False
Application.EnableEvents = False
With ThisWorkbook.Sheets("キャスク検査")
.Unprotect
  Dim ATU1 As Double:  ATU1 = val(.Cells(2, 17))
  Dim haba1 As Double: haba1 = val(.Cells(3, 17))
  Dim take1 As Double: take1 = val(.Cells(4, 17))
  Dim ATU2 As Double:  ATU2 = val(.Cells(2, 19))
  Dim haba2 As Double: haba2 = val(.Cells(3, 19))
  Dim take2 As Double: take2 = val(.Cells(4, 19))
For i = 7 To 706
  For a = 2 To 17
    If Not .Cells(i, a) = "" Then
      Select Case a
      Case 2, 4, 6, 8 '2021.07.09_8列目追加
        If KenJudg((Cells(i, a)), ATU2, ATU1) Then
          .Cells(i, a + 1) = "○"
          .Cells(i, a + 1).Interior.Color = RGB(0, 128, 0)
        Else
          .Cells(i, a + 1) = "×"
          .Cells(i, a + 1).Interior.Color = RGB(200, 60, 50)
        End If
      Case 11, 13
        If KenJudg((Cells(i, a)), haba2, haba1) Then
          .Cells(i, a + 1) = "○"
          .Cells(i, a + 1).Interior.Color = RGB(0, 128, 0)
        Else
          .Cells(i, a + 1) = "×"
          .Cells(i, a + 1).Interior.Color = RGB(200, 60, 50)
        End If
      Case 15, 17
        If KenJudg((Cells(i, a)), take2, take1) Then
          .Cells(i, a + 1) = "○"
          .Cells(i, a + 1).Interior.Color = RGB(0, 128, 0)
        Else
          .Cells(i, a + 1) = "×"
          .Cells(i, a + 1).Interior.Color = RGB(200, 60, 50)
        End If
      End Select
    End If
  Next a
Next i
.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True 'シート保護
End With
Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub
Public Sub 検査判定(Ter As Range) '2018.02.20　公差外があった場合別保存してその板をクリアする
'2019.10.16 log吐出し追加
'2019.11.02 大改修　3種ｼｰﾄ対応
'2019.11.11 ｷｬｽｸ,全ヵ所検査追加
'2020.07.10 UFVerに移行
Dim Lname As String
Select Case True
Case ThisWorkbook.Sheets("全ヵ所検査").Visible:   Lname = UFMitsubishi.LineName.Caption
Case ThisWorkbook.Sheets("キャスク検査").Visible: Lname = UFCask.LineName.Caption
Case ThisWorkbook.Sheets("H494検査").Visible:     Lname = UFMitsubishi.LineName.Caption
End Select
If UFCask.Visible = True Then
  If UFCask.Micro.Caption = "" Or UFCask.Convex.Caption = "" Then
    MsgBox "ゼロ校正を実行してください！！", vbExclamation '2021.07.09_校正前は検査不可とする
    Exit Sub
  End If
End If
Dim ATU1 As Double, haba1 As Double, take1 As Double '-
Dim ATU2 As Double, haba2 As Double, take2 As Double
Dim i As Long, a As Long, b As Long, C As Long
Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
Dim TR As Long: TR = Ter.row
Dim TC As Long: TC = Ter.Column
'------------------------------公差
With ThisWorkbook.Sheets(ws.Name)
If Not IsNumeric(.Cells(3, 2)) And IsNumeric(.Cells(3, 4)) And IsNumeric(.Cells(3, 6)) Then Exit Sub
If Not IsNumeric(Ter) Then Exit Sub '判定○×スルー
If WsCheck("Daily" & "_" & Lname) Then '日報情報
  Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Daily" & "_" & Lname)
  Dim LR As Long: LR = UR(sh, 1)
  Dim WorkLot As String
  Select Case True
  Case ThisWorkbook.Sheets("全ヵ所検査").Visible:   WorkLot = ThisWorkbook.Sheets("全ヵ所検査").Cells(2, 4)
  Case ThisWorkbook.Sheets("キャスク検査").Visible: WorkLot = ThisWorkbook.Sheets("キャスク検査").Cells(2, 4)
  Case ThisWorkbook.Sheets("H494検査").Visible:     WorkLot = ThisWorkbook.Sheets("H494検査").Cells(2, 4)
  End Select
'  If Sh.Cells(LR, 1) <> WorkLot Then Call Daily_Report(2)   '日報情報(Lot情報)  'LotNoの一致が条件
  Call Daily_Report(2) '---------------検査測定のﾀｲﾐﾝｸﾞなので無条件でLot情報書き込み_2021.04.22
' If st = "S1" Then Call Daily_Report(3) '日報情報(検査開始)　＊開始時間を個別に行うと繋がらなくなる為Worker_Changeのみに実行
End If
'指定座標以外は判定しない
 If VarType(Ter) = vbString Or (Ter.Value) = "" Then   '文字　空白の時は何もしない
' Exit Sub
 Else
  ATU1 = val(.Cells(2, 17)) '-
  ATU2 = val(.Cells(2, 19))
  haba1 = val(.Cells(3, 17)) '-
  haba2 = val(.Cells(3, 19))
  take1 = val(.Cells(4, 17)) '-
  take2 = val(.Cells(4, 19))
  Select Case ws.Name
  Case "キャスク検査"
    UFCask.Tag = "off"
    Select Case True
    Case (TR >= 7 And TR <= 706) And (TC = 2 Or TC = 4 Or TC = 6 Or TC = 8)  '指定座標内_2021.07.09_TC=8追加
       Call 判定付け(KenJudg(Ter, ATU2, ATU1), Ter)
    Case (TR >= 7 And TR <= 706) And (TC = 11 Or TC = 13)   '指定座標内
       Call 判定付け(KenJudg(Ter, haba2, haba1), Ter)
    Case (TR >= 7 And TR <= 706) And (TC = 15 Or TC = 17)   '指定座標内
       Call 判定付け(KenJudg(Ter, take2, take1), Ter)
    End Select
    UFCask.Tag = ""
  Case "全ヵ所検査"
    UFMitsubishi.Tag = "off"
    Select Case True
    Case (TR >= 7 And TR <= 76) And (TC = 2 Or TC = 4 Or TC = 6 Or TC = 8)
       Call 判定付け(KenJudg(Ter, ATU2, ATU1), Ter)
    Case (TR >= 7 And TR <= 76) And (TC = 11 Or TC = 13)
       Call 判定付け(KenJudg(Ter, haba2, haba1), Ter)
    Case (TR >= 7 And TR <= 76) And (TC = 15 Or TC = 17)
       Call 判定付け(KenJudg(Ter, take2, take1), Ter)
    End Select
    UFMitsubishi.Tag = ""
  Case "H494検査"
    UFMitsubishi.Tag = "off"
    Select Case True
    Case (TR >= 7 And TR <= 16) And (TC = 2 Or TC = 4 Or TC = 6 Or TC = 8)
       Call 判定付け(KenJudg(Ter, ATU2, ATU1), Ter)
    Case (TR >= 7 And TR <= 16) And (TC = 11 Or TC = 13 Or TC = 15)
       Call 判定付け(KenJudg(Ter, haba2, haba1), Ter)
    Case (TR >= 7 And TR <= 16) And (TC = 17 Or TC = 19)
       Call 判定付け(KenJudg(Ter, take2, take1), Ter)
    End Select
    UFMitsubishi.Tag = ""
  End Select
 End If
End With
End Sub

Public Sub 対角(Ter As Range) '2018.02.26
'2019.11.06 大改修　2種ｼｰﾄ対応
'2019.11.11 ｷｬｽｸ,全ヵ所検査追加
  Dim HABA As Single, TAKE As Single, kou As Single, Tai As Single
  Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
  Dim ZenSh As Worksheet
  Select Case UFMitsubishi.WorkFlag.Caption
  Case "P1"
    Set ZenSh = ThisWorkbook.Sheets("全ヵ所検査")
  Case "P_H494"
    Set ZenSh = ThisWorkbook.Sheets("H494検査")
  End Select
  
  Dim TR As Long: TR = Ter.row
  Dim TC As Long: TC = Ter.Column
Application.ScreenUpdating = False

If UFCask.Visible Then GoTo CaskSkip '2023.6.16_重たい修正

'------------------------------------------列20でしか判定しない
'Select Case True
'Case TC = 20
'  Haba = Val(ws.Cells(3, 4))
'  Take = Val(ws.Cells(3, 6))
''------------------------------公差等級.4級(厳しい)製品規格：書式番号Z501
' Select Case Haba
'  Case 1# To 300#
'    Select Case Take
'     Case 1# To 3700#
'      kou = 0.8
'     Case 3701# To 6000#
'      kou = 1#
'    End Select
'  Case 300.1 To 600#
'    Select Case Take
'     Case 1# To 3700#
'      kou = 1#
'     Case 3701# To 6000#
'      kou = 1.5
'    End Select
'  Case 600.1 To 900#
'    Select Case Take
'     Case 1# To 3700#
'      kou = 1.5
'     Case 3701# To 6000#
'      kou = 2#
'    End Select
'  Case 900.1 To 1200#
'    Select Case Take
'     Case 1# To 3700#
'      kou = 2#
'     Case 3701# To 6000#
'      kou = 2.5
'    End Select
'  Case 1200.1 To 1500#
'    Select Case Take
'     Case 1# To 3700#
'      kou = 2.5
'     Case 3701# To 6000#
'      kou = 3#
'    End Select
'  Case 1500.1 To 1800#
'    Select Case Take
'     Case 1# To 3700#
'      kou = 2.5
'     Case 3701# To 6000#
'      kou = 3#
'    End Select
'  Case 1800.1 To 2100#
'    Select Case Take
'     Case 1# To 3700#
'      kou = 2.5
'     Case 3701# To 6000#
'      kou = 3#
'    End Select
'  Case 2100.1 To 2400#
'    Select Case Take
'     Case 1# To 3700#
'      kou = 2.5
'     Case 3701# To 6000#
'      kou = 3#
'    End Select
' End Select
'公差追加表示_2023.4.13
'ZenSh.Unprotect 'シート保護解除
kou = val(ZenSh.Cells(1, 19))
 
'---------------------------------------------------------------------------------------------------------------------------------------2023.6.16_重たい修正
CaskSkip:
  If ws.Name = "キャスク検査" Then
    Select Case True
    Case TC = 20
      If UFCask.Tai = "" Then
      MsgBox "対角公差が入力されていません！！", vbExclamation '2020.06.09 ｷｬｽｸ公差のほうが厳しいことがある必須項目とする
      Exit Sub
      End If
      If UFCask.Inspection = "ON" Then 'フォーム上で巾丈測定にチェックが入っている時
        If UFCask.Tai <> "" Then 'フォーム上に数字がある時
         If IsNumeric(UFCask.Tai) = True Then 'かつ数字の時
           kou = val(UFCask.Tai) '公差はフォーム入力されている数字
         End If
        End If
      End If
      Call 対角判定付け(TaiJudg(Ter, kou), Ter, kou)
    End Select
  End If
  
  If ws.Name = "全ヵ所検査" And TC = 20 Then Call 対角判定付け(TaiJudg(Ter, kou), Ter, kou)
  If ws.Name = "H494検査" And TC = 22 Then Call 対角判定付け(TaiJudg(Ter, kou), Ter, kou)
  
'------------------------------
'2018.03.27 色の付けるセルを○×のセルのみに変更
'2019.06.06 対角の判定を都度行う様に修正
'2019.11.06 Function TaiJudg作成
'2023.6.20_公差はシートに先に反映　その公差を読む
'Call 対角判定付け(TaiJudg(Ter, kou), Ter, kou)
'End Select
Application.ScreenUpdating = True
End Sub
Public Sub 対角再チェック(Ter As Range)
'再チェックに限り　フォーム上のTR、TCから数字をもらう
'2019.11.06 大改修　2種ｼｰﾄ対応
'2019.11.11 全ヵ所検査追加
  Dim HABA As Single, TAKE As Single, kou As Single, Tai As Single
  Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
  Dim TR As Long: TR = Ter.row
  Dim TC As Long: TC = Ter.Column
'  MsgBox TR & ":" & TC 'ﾃﾞﾊﾞｯｸﾞ
  HABA = val(ws.Cells(3, 4))
  TAKE = val(ws.Cells(3, 6))
'------------------------------公差等級.3級(緩い)製品規格：書式番号Z501
 Select Case HABA
  Case 1# To 300#
    Select Case TAKE
     Case 1# To 3700#
      kou = 1.5
     Case 3701# To 6000#
      kou = 2#
    End Select
  Case 300.1 To 600#
    Select Case TAKE
     Case 1# To 3700#
      kou = 2#
     Case 3701# To 6000#
      kou = 3#
    End Select
  Case 600.1 To 900#
    Select Case TAKE
     Case 1# To 3700#
      kou = 3#
     Case 3701# To 6000#
      kou = 4#
    End Select
  Case 900.1 To 1200#
    Select Case TAKE
     Case 1# To 3700#
      kou = 4#
     Case 3701# To 6000#
      kou = 5#
    End Select
  Case 1200.1 To 1500#
    Select Case TAKE
     Case 1# To 3700#
      kou = 5#
     Case 3701# To 6000#
      kou = 6#
    End Select
  Case 1500.1 To 1800#
    Select Case TAKE
     Case 1# To 3700#
      kou = 5#
     Case 3701# To 6000#
      kou = 6#
    End Select
  Case 1800.1 To 2100#
    Select Case TAKE
     Case 1# To 3700#
      kou = 5#
     Case 3701# To 6000#
      kou = 6#
    End Select
  Case 2100.1 To 2400#
    Select Case TAKE
     Case 1# To 3700#
      kou = 5#
     Case 3701# To 6000#
      kou = 6#
    End Select
 End Select
'------------------------------
'2018.03.27 色の付けるセルを○×のセルのみに変更 再チェックは×か空白の時2018.04.10
'2019.06.06 位置フラグ追加　再判定の条件の×を外した
'2019.11.06 ファンクション作成
'  MsgBox "対角再チェック" & "*" & kou
  Select Case ws.Name
  Case "全ヵ所検査"
    Select Case True
    Case (TR >= 7 And TR <= 76)  '指定座標内
      If (IsNumeric(ws.Cells(TR, 19)) And IsNumeric(ws.Cells(TR, 20))) And (ws.Cells(TR, 19) <> "" And ws.Cells(TR, 20) <> "") Then 'この条件がないとfalse扱い
        If TaiJudg(Ter, kou) Then
          ws.Cells(TR, 21) = "△"
          ws.Cells(TR, 21).Interior.Color = RGB(153, 204, 255)
          Call WriteLog("*UFKen:" & "*対角再ﾁｪｯｸ:" & Cells(TR, TC).Value & " *位置:" & TR & "," & TC & "：△", "", "")
          Call 音楽合格
        ElseIf TaiJudg(Ter, kou) = False Then '公差外
          ws.Cells(TR, 21) = "×"
          ws.Cells(TR, 21).Interior.Color = RGB(200, 60, 50)
          Call WriteLog("*UFKen:" & "*対角再ﾁｪｯｸ:" & Cells(TR, TC).Value & " *位置:" & TR & "," & TC & "：×", "", "")
          Call 音楽不合格
          UFNG.Show vbModeless
          ws.Cells(TR, 23) = ""
          ws.Cells(TR, 23) = "3級公差" & ":" & "[" & Format(kou, "0.0") & "]"
          ws.Cells(TR, 23).Font.Size = 8
        End If
      End If
    End Select
  End Select
End Sub
Public Sub 公差外検索(Ter As Range) '全カ所測定後セル内に×のあるものを探す、2018.03.06　耳付き検査の時は耳付公差外検索
'2017.04.14　作業毎保存削除　セルの位置が大きく変わって処理が難しくなった
'2019.06.06  次の行で反映させていたものを削除　対角セルで反映させることで不要となった　対角セル用はｸﾗｽﾓｼﾞｭｰﾙに追加
'2019.10.30  大改修　整理
Dim TR As Long: TR = Ter.row
Dim TC As Long: TC = Ter.Column
Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
'公差外は探さないが途中保存はする
Select Case ws.Name
Case "キャスク検査" '細かく保存 2019.12.11
  Select Case TC
  Case 2, 6 '2021.07.09_検査厚で個別に設定
    If UFCask.CountSt.Caption = "3" Then Call 作業毎保存
  Case 4, 8
    If UFCask.CountSt.Caption = "4" Then Call 作業毎保存
  Case 13, 17, 20, 26
    Call 作業毎保存
  End Select
Case "全ヵ所検査"
  Select Case TC
  Case 4, 8, 13, 17, 20, 23
    Call 作業毎保存
  End Select
Case "H494検査"
  Select Case TC
  Case 4, 8, 15, 19, 22, 28
    Call 作業毎保存
  End Select
End Select
End Sub
Public Sub 判定付け(Judg As Boolean, Ter As Range)
If Not IsNumeric(Ter) Then Exit Sub
Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
Dim TR As Long: TR = Ter.row
Dim TC As Long: TC = Ter.Column
Dim YColor As Long, NColor As Long

Dim Lname As String
Select Case True
Case ThisWorkbook.Sheets("全ヵ所検査").Visible:   Lname = UFMitsubishi.LineName.Caption
Case ThisWorkbook.Sheets("キャスク検査").Visible: Lname = UFCask.LineName.Caption
End Select

Dim fname As String: fname = "Measured.log"
Dim path As String: path = HistoryPath
Dim Lpath As String
Select Case Lname
Case "L-1":    Lpath = path & "L-1\L-1_Log"
Case "LVC":    Lpath = path & "LVC\LVC_Log"
Case "HVC":    Lpath = path & "HVC\HVC_Log"
Case "ﾊﾞﾗﾝｻｰ": Lpath = path & "バランサー\バランサー_Log"
Case "ﾄｯﾄ":    Lpath = path & "トット\トット_Log"
Case "Test":   Lpath = path & "Test\Test_Log"
Case Else:     Lpath = path & "丸徳\丸徳_Log"
End Select

'-------------------------------------------------判定色設定
Select Case ws.Name
Case "キャスク検査", "全ヵ所検査", "H494検査"
 YColor = 10
 NColor = 3
End Select
'-------------------------------------------------判定有無
Select Case True 'Undo使用後の不要な判定を止める　2019.12.24
Case ThisWorkbook.Sheets("キャスク検査").Visible
  If UFCask.UndoCount = "1" Then
    UFCask.UndoCount = "0" 'ﾘｾｯﾄしてExit
    Exit Sub
  End If
Case ThisWorkbook.Sheets("全ヵ所検査").Visible, ThisWorkbook.Sheets("H494検査").Visible
  If UFMitsubishi.UndoCount = "1" Then
    UFMitsubishi.UndoCount = "0"
    Exit Sub
  End If
End Select
'-------------------------------------------------判定処理
  If Judg Then
    ws.Cells(TR, TC + 1) = "○"
    ws.Cells(TR, TC + 1).Interior.ColorIndex = YColor 'RGB(0, 128, 0)
    Call WriteLog("*UFKen:" & ws.Name & "*入力:" & ws.Cells(TR, TC).Value & " *座標:" & "(" & TR & "," & TC & ")" & "：○", "", "") 'シート名追加_2023.11.14
    Call WriteLog("*UFKen:" & ws.Name & "*入力:" & ws.Cells(TR, TC).Value & " *座標:" & "(" & TR & "," & TC & ")" & "：○:" & Lname, "", Lpath, fname)
    
      Call 音楽合格
    '苦肉策_2023.12.04
    Select Case TC
    Case 2, 4, 6, 11, 13, 15, 17, 20: ws.Cells(TR, TC + 2).Interior.Color = RGB(153, 204, 255)
    Case 8:                           ws.Cells(TR, TC + 3).Interior.Color = RGB(153, 204, 255)
    End Select
      
      Select Case True 'キャスクはフォームを出さない
      Case ThisWorkbook.Sheets("キャスク検査").Visible
       UFcount.Maru.Visible = True
       Call UFcount起動(Cells(TR, TC))
      Case Else
       UFOK.Show vbModeless
      End Select
  ElseIf Judg = False Then
    ws.Cells(TR, TC + 1) = "×"
    ws.Cells(TR, TC + 1).Interior.ColorIndex = NColor
    Call WriteLog("*UFKen:" & ws.Name & "*入力:" & ws.Cells(TR, TC).Value & " *座標:" & "(" & TR & "," & TC & ")" & "：×", "", "")
    Call WriteLog("*UFKen:" & ws.Name & "*入力:" & ws.Cells(TR, TC).Value & " *座標:" & "(" & TR & "," & TC & ")" & "：×:" & Lname, "", Lpath, fname)
    
     Call 音楽不合格
     
    '苦肉策_2023.12.04
    Select Case TC
    Case 2, 4, 6, 11, 13, 15, 17, 20: ws.Cells(TR, TC + 2).Interior.Color = RGB(153, 204, 255)
    Case 8:                           ws.Cells(TR, TC + 3).Interior.Color = RGB(153, 204, 255)
    End Select
     
     Select Case True 'キャスクはフォームを出さない
     Case ThisWorkbook.Sheets("キャスク検査").Visible
      UFcount.Batu.Visible = True
      Call UFcount起動(Cells(TR, TC))
     Case Else
      UFNG.Show vbModeless
      '2021.07.09_一旦停止(全数検査指示のため)
'      If UFMitsubishi.Choice = "5N" Then 'ﾏｲﾅｽ外れの時最終行+1へｺﾋﾟｰ
'       Dim Sh As Worksheet: Set Sh = ThisWorkbook.Sheets("全ヵ所検査")
'       Dim LR As Long: LR = UR(Sh, 2)
'       Dim CopyARan As Range:  Set CopyARan = Range(Sh.Cells(TR, 2), Sh.Cells(TR, 9))
'       Dim CopyHTRan As Range: Set CopyHTRan = Range(Sh.Cells(TR, 11), Sh.Cells(TR, 19))
'       Dim CopyETRan As Range: Set CopyETRan = Range(Sh.Cells(TR, 24), Sh.Cells(TR, 26))
'       If Judg_PorM(Val(Ter.Value), Val(Sh.Cells(2, 17)), Val(Sh.Cells(2, 19))) = False Then '-外れ
'         Application.EnableEvents = False
'         CopyARan.Copy Destination:=Sh.Cells(LR + 1, 2)
'         CopyHTRan.Copy Destination:=Sh.Cells(LR + 1, 11)
'         CopyETRan.Copy Destination:=Sh.Cells(LR + 1, 24)
'         CopyARan.ClearContents
'         CopyHTRan.ClearContents
'         CopyETRan.ClearContents
'         Range(Sh.Cells(TR, 2), Sh.Cells(TR, 26)).Interior.Pattern = xlNone
'         Application.EnableEvents = True
'       End If
'      End If
     End Select
  End If
End Sub
Public Sub 対角判定付け(Judg As Boolean, Ter As Range, kou As Single)
Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
Dim TR As Long: TR = Ter.row
Dim TC As Long: TC = Ter.Column
Dim Tai As Single

Dim Lname As String
Select Case True
Case ThisWorkbook.Sheets("全ヵ所検査").Visible:   Lname = UFMitsubishi.LineName.Caption
Case ThisWorkbook.Sheets("キャスク検査").Visible: Lname = UFCask.LineName.Caption
End Select

Dim fname As String: fname = "Measured.log"
Dim path As String: path = HistoryPath
Dim Lpath As String
Select Case Lname
Case "L-1":    Lpath = path & "L-1\L-1_Log"
Case "LVC":    Lpath = path & "LVC\LVC_Log"
Case "HVC":    Lpath = path & "HVC\HVC_Log"
Case "ﾊﾞﾗﾝｻｰ": Lpath = path & "バランサー\バランサー_Log"
Case "ﾄｯﾄ":    Lpath = path & "トット\トット_Log"
Case "Test":   Lpath = path & "Test\Test_Log"
Case Else:     Lpath = path & "丸徳\丸徳_Log"
End Select

  Select Case ws.Name
  Case "キャスク検査" '外れた時の再チェックは行わない　フォームに公差入力がある場合はそちらを優先
    If (IsNumeric(ws.Cells(TR, 19)) And IsNumeric(ws.Cells(TR, 20))) And (ws.Cells(TR, 19) <> "" And ws.Cells(TR, 20) <> "") Then 'この条件がないとfalse扱い
      ws.Cells(TR, 22) = Abs(Format(val(ws.Cells(TR, 19)) - val(ws.Cells(TR, 20)), "0.0"))
      If Judg Then
        ws.Cells(TR, 21) = "○"
        ws.Cells(TR, 21).Interior.Color = RGB(153, 204, 255) '反転色のとき
        Call WriteLog("*UFKen:" & ws.Name & "*入力:" & ws.Cells(TR, TC).Value & " *位置:" & TR & "," & TC & "：○", "", "") 'シート名追加_2023.11.14
        Call WriteLog("*UFKen:" & ws.Name & "*入力:" & ws.Cells(TR, TC).Value & " *位置:" & TR & "," & TC & "：○:" & Lname, "", Lpath, fname)
        
        Call 音楽合格
        Select Case True 'キャスクはUFOKフォームを出さない
        Case Sheets("キャスク検査").Visible
         UFcount.Maru.Visible = True
        End Select
        ws.Cells(TR, 23) = ""
      ElseIf Judg = False Then
        ws.Cells(TR, 21) = "×"
        ws.Cells(TR, 21).Interior.Color = RGB(200, 60, 50)
        Call WriteLog("*UFKen:" & "*入力:" & ws.Cells(TR, TC).Value & " *位置:" & TR & "," & TC & "：×", "", "")
        Call WriteLog("*UFKen:" & "*入力:" & ws.Cells(TR, TC).Value & " *位置:" & TR & "," & TC & "：×:" & Lname, "", Lpath, fname)
        
        Call 音楽不合格
      Select Case True 'キャスクはUFNGフォームを出さない
        Case Sheets("キャスク検査").Visible
         UFcount.Batu.Visible = True
        End Select
        Tai = Abs(Format(val(ws.Cells(TR, 19)) - val(ws.Cells(TR, 20)), "0.0")) 'ﾃﾞﾊﾞｯｸﾞ
        ws.Cells(TR, 23) = ""
        If UFCask.Tai <> "" Then
          ws.Cells(TR, 23) = "公差" & ":" & "[" & Format(UFCask.Tai, "0.0") & "]"
          ws.Cells(TR, 23).Font.Size = 8
        ElseIf UFCask.Tai = "" Then
          ws.Cells(TR, 23) = "4級公差" & ":" & "[" & Format(kou, "0.0") & "]"
          ws.Cells(TR, 23).Font.Size = 8
        End If
      End If
    End If
  Case "全ヵ所検査"
    ws.Cells(TR, 21) = Abs(Format(val(ws.Cells(TR, 19)) - val(ws.Cells(TR, 20)), "0.0"))
    If (IsNumeric(ws.Cells(TR, 19)) And IsNumeric(ws.Cells(TR, 20))) And (ws.Cells(TR, 19) <> "" And ws.Cells(TR, 20) <> "") Then 'この条件がないとfalse扱い
      If Judg Then
        ws.Cells(TR, 22) = "○"
        ws.Cells(TR, 22).Interior.Color = RGB(0, 128, 0) '反転色のとき RGB(0, 128, 0) RGB(153, 204, 255)
        Call WriteLog("*UFKen:" & "*入力:" & ws.Cells(TR, TC).Value & " *位置:" & TR & "," & TC & "：○", "", "") 'シート名追加_2023.11.14
        Call WriteLog("*UFKen:" & "*入力:" & ws.Cells(TR, TC).Value & " *位置:" & TR & "," & TC & "：○:" & Lname, "", Lpath, fname)
        Call 音楽合格
        UFOK.Show vbModeless
      ElseIf Judg = False Then
       ws.Cells(TR, 22).Interior.Color = RGB(200, 60, 50)
       ws.Cells(TR, 22) = "×"
       Call WriteLog("*UFKen:" & "*入力:" & ws.Cells(TR, TC).Value & " *位置:" & TR & "," & TC & "：×", "", "")
       Call WriteLog("*UFKen:" & "*入力:" & ws.Cells(TR, TC).Value & " *位置:" & TR & "," & TC & "：×:" & Lname, "", Lpath, fname)
       Call 音楽不合格
       UFNG.Show vbModeless
      End If
    End If
  Case "H494検査"
    ws.Cells(TR, 24) = Abs(Format(val(ws.Cells(TR, 22)) - val(ws.Cells(TR, 21)), "0.0"))
    If (IsNumeric(ws.Cells(TR, 21)) And IsNumeric(ws.Cells(TR, 22))) And (ws.Cells(TR, 21) <> "" And ws.Cells(TR, 22) <> "") Then 'この条件がないとfalse扱い
      If Judg Then
        ws.Cells(TR, 23) = "○"
        ws.Cells(TR, 23).Interior.Color = RGB(153, 204, 255) '反転色のとき
        Call WriteLog("*UFKen:" & "*入力:" & ws.Cells(TR, TC).Value & " *位置:" & TR & "," & TC & "：○", "", "")
        Call WriteLog("*UFKen:" & "*入力:" & ws.Cells(TR, TC).Value & " *位置:" & TR & "," & TC & "：○:" & Lname, "", Lpath, fname)
        Call 音楽合格
        UFOK.Show vbModeless
      ElseIf Judg = False Then
       ws.Cells(TR, 23).Interior.Color = RGB(200, 60, 50)
       ws.Cells(TR, 23) = "×"
       Call WriteLog("*UFKen:" & "*入力:" & ws.Cells(TR, TC).Value & " *位置:" & TR & "," & TC & "：×", "", "")
       Call WriteLog("*UFKen:" & "*入力:" & ws.Cells(TR, TC).Value & " *位置:" & TR & "," & TC & "：×:" & Lname, "", Lpath, fname)
       Call 音楽不合格
       UFNG.Show vbModeless
      End If
    End If
'2023.6.20_停止　公差反映修正
'        TAI = Abs(Format(Val(ws.Cells(TR, 19)) - Val(ws.Cells(TR, 20)), "0.0")) 'ﾃﾞﾊﾞｯｸﾞ
'          If UFMitsubishi.Work = "ON" Then
'            ws.Cells(TR, 23) = "三菱公差." & "[" & Format(Val(kou), "0.0") & "]" '2020.01.09変更
'            ws.Cells(TR, 23).Font.Size = "8"
'            Call 音楽不合格
'          End If
'       MsgBox Tai

'2023.4.12****再チェック不要とした
'        If UFMitsubishi.Work = "OFF" Then
'          ws.Cells(TR, 23) = "4級公差." & "[" & Format(Val(kou), "0.0") & "]"
'          ws.Cells(TR, 23).Font.Size = "8"
'          Call 対角再チェック(Ter) '2020.01.09追加
'        End If
  End Select
End Sub
Public Sub Special_Move_DT(TarR As Range, Optional BeforDate As String = "")
'ｷｬｽｸ.三菱用
Dim i As Long, j As Long, nj As Long
Dim ChBefor As String, ChAfter As String
Dim flag_Blank As Boolean
Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
Dim TR As Long: TR = TarR.row
Dim TC As Long: TC = TarR.Column
'MsgBox (TarR.Address) 'ﾃﾞﾊﾞｯｸﾞ
Application.EnableEvents = False 'イベント停止のまま続行
Sheets(ws.Name).Unprotect
  ChAfter = TarR '変更後
  ChBefor = BeforDate '変更前
'  MsgBox ChAfter & "*" & ChBefor
Select Case True 'Undo使用後の不要な判定を止める　2019.12.24
Case ThisWorkbook.Sheets("キャスク検査").Visible
  UFCask.UndoCount = "0"
Case Sheets("全ヵ所検査").Visible
  UFMitsubishi.UndoCount = "0"
End Select
Select Case True
Case ws.Name = "キャスク検査"
    flag_Blank = False
    For i = TR To 706 '-----------------------------------------ｱｸﾃｨﾌﾞｾﾙ以降の空白
    Dim n As Long
    If UFCask.CountSt.Caption = "4" Then n = 8 '2021.07.09_検査箇所変動用追加
    If UFCask.CountSt.Caption = "3" Then n = 6
      For j = 2 To n Step 2
        If flag_Blank = True Then Exit For
          If ws.Cells(i, j) = "" Then
            ws.Cells(i, j) = Conv(ChAfter) '数値変換
            TarR = ChBefor
            Application.EnableEvents = True 'イベント開始
            Call 検査判定(ws.Cells(i, j))
            Application.EnableEvents = False 'イベント停止
            AppActivate Application.Caption
            flag_Blank = True
            Exit For
          End If
      Next j
    If flag_Blank = True Then Exit For
    Next i
    If flag_Blank = False Then
      MsgBox "厚欄はすべて埋まっています", vbCritical
      TarR = ChBefor
    End If
Case ws.Name = "全ヵ所検査"
    flag_Blank = False
    For i = TR To 76 '-----------------------------------------ｱｸﾃｨﾌﾞｾﾙ以降の空白
      For j = 2 To 8 Step 2
        If flag_Blank = True Then Exit For
          If ws.Cells(i, j) = "" Then
            ws.Cells(i, j) = Conv(ChAfter) '数値変換
            TarR = ChBefor
            Application.EnableEvents = True 'イベント開始
            Call 検査判定(ws.Cells(i, j))
            Application.EnableEvents = False 'イベント停止
            AppActivate Application.Caption
            flag_Blank = True
            Exit For
          End If
      Next j
    '2023.5.16_追加
    If UFMitsubishi.H494 Then 'H494の時はDTパケット
        For nj = 11 To 13 Step 2
          If flag_Blank = True Then Exit For
            If ws.Cells(i, nj) = "" Then
              ws.Cells(i, nj) = Conv(ChAfter) '数値変換
              TarR = ChBefor
              Application.EnableEvents = True 'イベント開始
              Call 検査判定(ws.Cells(i, nj))
              Application.EnableEvents = False 'イベント停止
              AppActivate Application.Caption
              flag_Blank = True
              Exit For
            End If
        Next nj
    End If
      
    If flag_Blank = True Then Exit For
    Next i
    If flag_Blank = False Then
      MsgBox "厚欄はすべて埋まっています", vbCritical
      TarR = ChBefor
    End If
Case ws.Name = "H494検査" '2023.10.16
    flag_Blank = False
    For i = TR To 16 '-----------------------------------------ｱｸﾃｨﾌﾞｾﾙ以降の空白
      For j = 2 To 22
        Select Case j
        Case 2, 4, 6, 8, 11, 13, 15
          If ws.Cells(i, j) = "" Then
            ws.Cells(i, j) = Conv(ChAfter) '数値変換
            TarR = ChBefor
            Application.EnableEvents = True 'イベント開始
            Call 検査判定(ws.Cells(i, j))
            Application.EnableEvents = False 'イベント停止
            AppActivate Application.Caption
            flag_Blank = True
            Exit For
          End If
        End Select
        If flag_Blank = True Then Exit For
      Next j
      If flag_Blank = True Then Exit For
    Next i
    If flag_Blank = False Then
      MsgBox "厚欄はすべて埋まっています", vbCritical
      TarR = ChBefor
    End If
End Select
Select Case True 'Undo使用後の不要な判定を止める　2019.12.24
Case ThisWorkbook.Sheets("キャスク検査").Visible
  UFCask.UndoCount = "1"
  If ChBefor = "" Then
  UFCask.UndoCount = "0"
  End If
Case ThisWorkbook.Sheets("全ヵ所検査").Visible, ThisWorkbook.Sheets("H494検査").Visible
  UFMitsubishi.UndoCount = "1"
  If ChBefor = "" Then
  UFMitsubishi.UndoCount = "0"
  End If
End Select
End Sub
Public Sub Special_Move_L(TarR As Range, Optional BeforDate As String = "")
'ｷｬｽｸ.三菱用
Dim i As Long, j As Long
Dim ChBefor As String, ChAfter As String
Dim flag_Blank As Boolean
Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
Dim TR As Long: TR = TarR.row
Dim TC As Long: TC = TarR.Column
'MsgBox Cells(TR, TC) 'ﾃﾞﾊﾞｯｸﾞ
Application.EnableEvents = False 'イベント停止のまま続行
Sheets(ws.Name).Unprotect
  ChAfter = TarR '変更後
  ChBefor = BeforDate '変更前
'  MsgBox ChAfter & "*" & ChBefor
Select Case True 'Undo使用後の不要な判定を止める　2019.12.24
Case ThisWorkbook.Sheets("キャスク検査").Visible
  UFCask.UndoCount = "0"
Case ThisWorkbook.Sheets("全ヵ所検査").Visible
  UFMitsubishi.UndoCount = "0"
End Select
  Select Case ws.Name
  Case "キャスク検査"
  If UFCask.Inspection = "ON" Then
        flag_Blank = False
        For i = TR To 706 '-----------------------------------------ｱｸﾃｨﾌﾞｾﾙ以降の空白
          For j = 11 To 20
            Select Case j
            Case 11, 13, 15, 17
            If flag_Blank = True Then Exit For
              If ws.Cells(i, j) = "" Then
                ws.Cells(i, j) = Conv(ChAfter) '数値変換
                TarR = ChBefor
                  Application.EnableEvents = True 'イベント開始
                  Call 検査判定(ws.Cells(i, j))
                  Application.EnableEvents = False 'イベント停止
                AppActivate Application.Caption
                flag_Blank = True
                Exit For
              End If
            Case 19, 20
            If flag_Blank = True Then Exit For
              If ws.Cells(i, j) = "" Then
                ws.Cells(i, j) = Conv(ChAfter) '数値変換
                TarR = ChBefor
                  Application.EnableEvents = True 'イベント開始
                  Call 対角(ws.Cells(i, j))
                  Application.EnableEvents = False 'イベント停止
                AppActivate Application.Caption
                flag_Blank = True
                Exit For
              End If
            End Select
          Next j
        If flag_Blank = True Then Exit For
        Next i
    If flag_Blank = False Then
      MsgBox "巾丈対角欄はすべて埋まっています", vbCritical
      TarR = ChBefor
    End If
  Else
    If flag_Blank = False Then
      MsgBox "巾丈測定がONになっていません", vbCritical
      TarR = ChBefor
    End If
  End If
  Case "全ヵ所検査"
        flag_Blank = False
        For i = TR To 76 '-----------------------------------------ｱｸﾃｨﾌﾞｾﾙ以降の空白
          For j = 11 To 20
            Select Case j
            Case 11, 13
            If UFMitsubishi.H494 = False Then 'H494の時はDTパケット
              If flag_Blank = True Then Exit For
                If ws.Cells(i, j) = "" Then
                  ws.Cells(i, j) = Conv(ChAfter) '数値変換
                  TarR = ChBefor
                    Application.EnableEvents = True 'イベント開始
                    Call 検査判定(ws.Cells(i, j))
                    Application.EnableEvents = False 'イベント停止
                  AppActivate Application.Caption
                  flag_Blank = True
                  Exit For
                End If
            End If
            Case 15, 17
            If flag_Blank = True Then Exit For
              If ws.Cells(i, j) = "" Then
                ws.Cells(i, j) = Conv(ChAfter) '数値変換
                TarR = ChBefor
                  Application.EnableEvents = True 'イベント開始
                  Call 検査判定(ws.Cells(i, j))
                  Application.EnableEvents = False 'イベント停止
                AppActivate Application.Caption
                flag_Blank = True
                Exit For
              End If
            Case 19, 20
            If flag_Blank = True Then Exit For
              If ws.Cells(i, j) = "" Then
                ws.Cells(i, j) = Conv(ChAfter) '数値変換
                TarR = ChBefor
                  Application.EnableEvents = True 'イベント開始
                  Call 対角(ws.Cells(i, j))
                  Application.EnableEvents = False 'イベント停止
                AppActivate Application.Caption
                flag_Blank = True
                Exit For
              End If
            End Select
          Next j
        If flag_Blank = True Then Exit For
        Next i
    If flag_Blank = False Then
      MsgBox "巾丈対角欄はすべて埋まっています", vbCritical
      TarR = ChBefor
    End If
  Case "H494検査" '2023.10.16
        flag_Blank = False
        For i = TR To 16 '-----------------------------------------ｱｸﾃｨﾌﾞｾﾙ以降の空白
          For j = 11 To 22
            Select Case j
            Case 17, 19
            If flag_Blank = True Then Exit For
              If ws.Cells(i, j) = "" Then
                ws.Cells(i, j) = Conv(ChAfter) '数値変換
                TarR = ChBefor
                  Application.EnableEvents = True 'イベント開始
                  Call 検査判定(ws.Cells(i, j))
                  Application.EnableEvents = False 'イベント停止
                AppActivate Application.Caption
                flag_Blank = True
                Exit For
              End If
            Case 21, 22
            If flag_Blank = True Then Exit For
              If ws.Cells(i, j) = "" Then
                ws.Cells(i, j) = Conv(ChAfter) '数値変換
                TarR = ChBefor
                  Application.EnableEvents = True 'イベント開始
                  Call 対角(ws.Cells(i, j))
                  Application.EnableEvents = False 'イベント停止
                AppActivate Application.Caption
                flag_Blank = True
                Exit For
              End If
            End Select
          Next j
        If flag_Blank = True Then Exit For
        Next i
    If flag_Blank = False Then
      MsgBox "巾丈対角欄はすべて埋まっています", vbCritical
      TarR = ChBefor
    End If
  End Select
Select Case True 'Undo使用後の不要な判定を止める　2019.12.24
Case ThisWorkbook.Sheets("キャスク検査").Visible
  UFCask.UndoCount = "1"
  If ChBefor = "" Then
  UFCask.UndoCount = "0"
  End If
Case ThisWorkbook.Sheets("全ヵ所検査").Visible, ThisWorkbook.Sheets("H494検査").Visible
  UFMitsubishi.UndoCount = "1"
  If ChBefor = "" Then
  UFMitsubishi.UndoCount = "0"
  End If
End Select
End Sub
Public Sub UFcount起動(Ter As Range)
Dim TR As Long:      TR = Ter.row
Dim TC As Long:      TC = Ter.Column
Dim ws As Worksheet: Set ws = ActiveSheet
UFcount.TR = TR
UFcount.TC = TC
Select Case TR
Case 1 To 6
Case 7 To 706
  Select Case TC
  Case 2, 4, 6, 8, 11, 13, 15, 17, 19, 20 '8追加_2021.07.09検査箇所変動
   If ws.Cells(TR, TC).Value <> "" Then
    UFcount.Show vbModeless 'vbModal
   End If
  End Select
End Select
End Sub
Public Sub UF対角起動(Ter As Range)
Dim TR As Long:      TR = Ter.row
Dim TC As Long:      TC = Ter.Column
Dim ws As Worksheet: Set ws = ActiveSheet
UF対角.TR = TR
UF対角.TC = TC

Select Case ws.Name
Case "キャスク検査":   If TC = 19 And ws.Cells(TR, 19).Value <> "" Then UF対角.Show vbModeless
Case "全ヵ所検査検査": If TC = 19 And ws.Cells(TR, 19).Value <> "" Then UF対角.Show vbModeless
Case "H494検査":       If TC = 21 And ws.Cells(TR, 21).Value <> "" Then UF対角.Show vbModeless
End Select

'Select Case TC
'Case 19
' If ws.Cells(TR, 19).value <> "" Then
'  UF対角.Show vbModeless 'vbModal
' End If
'End Select
End Sub
Public Sub 作業毎保存()
Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
Application.ScreenUpdating = False
Select Case ws.Name
Case "全ヵ所検査"
   Call AddWS("全ヵ所作業中保存") '追加シート処理
   ThisWorkbook.Sheets("全ヵ所作業中保存").Tab.Color = 6723891
   Call Pesute(ws.Name, "全ヵ所作業中保存") 'コピー処理
Case "キャスク検査"
   Call AddWS("ｷｬｽｸ作業中保存") '追加シート処理
   ThisWorkbook.Sheets("ｷｬｽｸ作業中保存").Tab.Color = 10079487
   Call Pesute(ws.Name, "ｷｬｽｸ作業中保存") 'コピー処理
Case "H494検査"
   Call AddWS("H494作業中保存") '追加シート処理
   ThisWorkbook.Sheets("H494作業中保存").Tab.Color = &HCCFFCC
   Call Pesute(ws.Name, "H494作業中保存") 'コピー処理
End Select
Application.ScreenUpdating = True
End Sub
 
Public Sub ProofreadingSearch(ToTime As String, ToTyo As String) '日付一致から校正データを読み込む_2021.06.24
'YeaDayrCheck(Format$(Now, "hh:mm")), TimeCheck(Format$(Now, "hh:mm:ss"))_日付,直を引数
If WsCheck("品証検証") = False Then Exit Sub
Dim saveSh As Worksheet: Set saveSh = ThisWorkbook.Sheets("品証検証")
Dim SaveR As Range:      Set SaveR = Range(saveSh.Cells(5, 1), saveSh.Cells(98, 11))
Dim SaveTemp():          SaveTemp = SaveR
Dim i As Long
For i = LBound(SaveTemp, 1) To UBound(SaveTemp, 1)
  If Left(SaveTemp(i, 11), 10) = ToTime Then '日付一致
    If SaveTemp(i, 2) = ToTyo Then '直一致
      UFCask.Convex_1000 = SaveTemp(i, 3) '測定値
      If UFCask.Kname = "" Then UFCask.Kname = SaveTemp(i, 10) '測定者
      UFCask.L3.Visible = True
    End If
  End If
Next i
End Sub
Public Function NumericConfi(n As String) As Boolean '校正値が閾値内かのチェック_2021.06.24
NumericConfi = True '閾値内の場合_True
Select Case n
Case "M": If val(UFCask.Micro) <> 0 Then NumericConfi = False
Case "C": If val(UFCask.Convex) > 0.5 Or val(UFCask.Convex) < -0.5 Then NumericConfi = False '未満
End Select
End Function
Public Sub CountSelect(a As Long) '検査箇所設定
If WsCheck("キャスク検査") Then
  Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("キャスク検査")
  ws.Activate 'ﾃﾞﾌｫﾙﾄ
  ws.Unprotect '保護解除
  ws.Range("H7:I706").Locked = True '狭めるロック設定
  ws.Columns("H:I").ColumnWidth = 0.05
  ws.Range(Cells(6, 8), Cells(6, 9)).Font.ThemeColor = xlThemeColorDark1
  ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
  UFCask.CountSt.Caption = "3"
Else
  Exit Sub
End If
With ws
  ws.Unprotect '保護解除
  Select Case a
  Case 3
    ws.Range("H7:I706").Locked = True '狭めるロック設定
    ws.Columns("H:I").ColumnWidth = 0.05
    UFCask.CountSt.Caption = "3"
  Case 4
    ws.Range("H7:I706").Locked = False '広げる
    ws.Columns("H:H").ColumnWidth = 8.88
    ws.Columns("I:I").ColumnWidth = 2.63
    ws.Range(Cells(6, 8), Cells(6, 9)).Font.ThemeColor = xlThemeColorLight1
    UFCask.CountSt.Caption = "4"
  End Select
  ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
  
End With
End Sub
Public Sub Hizumi_Z(Ter As Range) '歪判定
Dim Lname As String, sh As Worksheet
Select Case True
Case ThisWorkbook.Sheets("全ヵ所検査").Visible
  Lname = UFMitsubishi.LineName.Caption
  Set sh = ThisWorkbook.Sheets("全ヵ所検査")
Case ThisWorkbook.Sheets("キャスク検査").Visible
  Lname = UFCask.LineName.Caption
End Select

Dim i As Long, a As Long, b As Long, C As Long, kou As String
Dim ws As Worksheet:  Set ws = ThisWorkbook.ActiveSheet
Dim TR As Long:       TR = Ter.row
Dim TC As Long:       TC = Ter.Column
'------------------------------公差
With ThisWorkbook.Sheets(ws.Name)
  If Not IsNumeric(Ter) Then Exit Sub '判定○×スルー
  '指定座標以外は判定しない
  If VarType(Ter) = vbString Or (Ter.Value) = "" Then   '文字　空白の時は何もしない
  If (Ter.Value) = "" Then Exit Sub
  ' Exit Sub
  Else
    UFMitsubishi.Tag = "off"
    If UFMitsubishi.FiveOP Then
      Select Case ws.Name
        Case "全ヵ所検査"
        Select Case True
          Case (TR >= 7 And TR <= 76) And (TC = 25): Call Hizumi_判定付け(HizumiJ(Ter, 2#), Ter) '5N 歪公差固定_2024.1.11
          Case (TR >= 7 And TR <= 76) And (TC = 27): Call Hizumi_判定付け(HizumiJ(Ter, 3#), Ter)
        End Select
      End Select
    ElseIf UFMitsubishi.一般 Then
      Dim SetH As String: SetH = sh.Cells(5, 13).Value
      Dim SetT As String: SetT = sh.Cells(5, 15).Value
      Select Case ws.Name
        Case "全ヵ所検査"
        Select Case True
          Case (TR >= 7 And TR <= 76) And (TC = 25): Call Hizumi_判定付け(HizumiJ(Ter, SetH), Ter) '2024.4.26_追加
          Case (TR >= 7 And TR <= 76) And (TC = 27): Call Hizumi_判定付け(HizumiJ(Ter, SetT), Ter)
        End Select
      End Select
    End If
    UFMitsubishi.Tag = ""
  End If
End With
End Sub
Public Sub Lateralbow_Z(Ter As Range) 'ラテラルボー判定
Dim Lname As String
Select Case True
Case ThisWorkbook.Sheets("全ヵ所検査").Visible:   Lname = UFMitsubishi.LineName.Caption
Case ThisWorkbook.Sheets("キャスク検査").Visible: Lname = UFCask.LineName.Caption
End Select
Dim i As Long, a As Long, b As Long, C As Long, kou As String
Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
Dim TR As Long: TR = Ter.row
Dim TC As Long: TC = Ter.Column
'------------------------------公差
With ThisWorkbook.Sheets(ws.Name)
  If Not IsNumeric(.Cells(3, 2)) And IsNumeric(.Cells(3, 4)) And IsNumeric(.Cells(3, 6)) Then Exit Sub
  If Not IsNumeric(Ter) Then Exit Sub '判定○×スルー
  '指定座標以外は判定しない
   If VarType(Ter) = vbString Or (Ter.Value) = "" Then   '文字　空白の時は何もしない
   If (Ter.Value) = "" Then Exit Sub
  ' Exit Sub
   Else
      kou = val(.Cells(1, 13))
      Select Case ws.Name
      Case "全ヵ所検査"
        UFMitsubishi.Tag = "off"
        Select Case True
        Case (TR >= 7 And TR <= 76) And (TC = 23): Call Lateralbow_判定付け(LateralJ(Ter, kou), Ter)  '指定座標内
        End Select
        UFMitsubishi.Tag = ""
      Case "H494検査"
        UFMitsubishi.Tag = "off"
        Select Case True
        Case (TR >= 7 And TR <= 10) And (TC = 25): Call Lateralbow_判定付け(LateralJ(Ter, kou), Ter)  '指定座標内
        End Select
        UFMitsubishi.Tag = ""
      End Select
   End If
End With
End Sub
Public Sub Lateralbow_判定付け(Judg As Boolean, Ter As Range)
If Not IsNumeric(Ter) Then Exit Sub
Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
Dim TR As Long: TR = Ter.row
Dim TC As Long: TC = Ter.Column
Dim YColor As Long, NColor As Long

Dim Lname As String
Select Case True
Case ThisWorkbook.Sheets("全ヵ所検査").Visible:   Lname = UFMitsubishi.LineName.Caption
Case ThisWorkbook.Sheets("キャスク検査").Visible: Lname = UFCask.LineName.Caption
End Select

Dim fname As String: fname = "Measured.log"
Dim path As String: path = HistoryPath
Dim Lpath As String
Select Case Lname
Case "L-1":    Lpath = path & "L-1\L-1_Log"
Case "LVC":    Lpath = path & "LVC\LVC_Log"
Case "HVC":    Lpath = path & "HVC\HVC_Log"
Case "ﾊﾞﾗﾝｻｰ": Lpath = path & "バランサー\バランサー_Log"
Case "ﾄｯﾄ":    Lpath = path & "トット\トット_Log"
Case "Test":   Lpath = path & "Test\Test_Log"
Case Else:     Lpath = path & "丸徳\丸徳_Log"
End Select

'-------------------------------------------------判定色設定
Select Case ws.Name
Case "キャスク検査", "全ヵ所検査", "H494検査"
 YColor = 10
 NColor = 3
End Select
'-------------------------------------------------判定有無
Select Case True 'Undo使用後の不要な判定を止める　2019.12.24
Case ThisWorkbook.Sheets("全ヵ所検査").Visible
  If UFMitsubishi.UndoCount = "1" Then
    UFMitsubishi.UndoCount = "0"
    Exit Sub
  End If
End Select
'-------------------------------------------------判定処理
  If Judg Then
    Select Case ws.Name
    Case "全ヵ所検査"
      ws.Cells(TR, TC + 1) = "○"
      ws.Cells(TR, TC + 1).Interior.ColorIndex = YColor 'RGB(0, 128, 0)
    Case "キャスク検査", "H494検査"
      ws.Cells(TR, TC).Interior.ColorIndex = YColor 'RGB(0, 128, 0)
    End Select
    Call WriteLog("*UFKen:" & ws.Name & "*入力:" & ws.Cells(TR, TC).Value & " *座標:" & "(" & TR & "," & TC & ")" & "：○", "", "")  'シート名追加_2023.11.14
    Call WriteLog("*UFKen:" & ws.Name & "*入力:" & ws.Cells(TR, TC).Value & " *座標:" & "(" & TR & "," & TC & ")" & "：○:" & Lname, "", Lpath, fname)
    Call 音楽合格
    
    Select Case True 'キャスクはフォームを出さない
    Case ThisWorkbook.Sheets("キャスク検査").Visible
     UFcount.Maru.Visible = True
     Call UFcount起動(Cells(TR, TC))
    Case Else
     UFOK.Show vbModeless
    End Select
  ElseIf Judg = False Then
    Select Case ws.Name
    Case "全ヵ所検査"
      ws.Cells(TR, TC + 1) = "×"
      ws.Cells(TR, TC + 1).Interior.ColorIndex = NColor
    Case "キャスク検査", "H494検査"
      ws.Cells(TR, TC).Interior.Color = RGB(200, 60, 50)
    End Select
    Call WriteLog("*UFKen:" & ws.Name & "*入力:" & ws.Cells(TR, TC).Value & " *座標:" & "(" & TR & "," & TC & ")" & "：×", "", "")
    Call WriteLog("*UFKen:" & ws.Name & "*入力:" & ws.Cells(TR, TC).Value & " *座標:" & "(" & TR & "," & TC & ")" & "：×:" & Lname, "", Lpath, fname)
    Call 音楽不合格
    
   Select Case True 'キャスクはフォームを出さない
   Case ThisWorkbook.Sheets("キャスク検査").Visible
    UFcount.Batu.Visible = True
    Call UFcount起動(Cells(TR, TC))
   Case Else
    UFNG.Show vbModeless
   End Select
  End If
  
End Sub
Public Sub Hizumi_判定付け(Judg As Boolean, Ter As Range)
'If Not UFMitsubishi.FiveOP Then Exit Sub
'5N以外も判定を行う_2024.5.7
If Not IsNumeric(Ter) Then Exit Sub
Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
Dim TR As Long: TR = Ter.row
Dim TC As Long: TC = Ter.Column
Dim YColor As Long, NColor As Long

Dim Lname As String
Select Case True
Case ThisWorkbook.Sheets("全ヵ所検査").Visible:   Lname = UFMitsubishi.LineName.Caption
Case ThisWorkbook.Sheets("キャスク検査").Visible: Lname = UFCask.LineName.Caption
End Select

Dim fname As String: fname = "Measured.log"
Dim path As String: path = HistoryPath
Dim Lpath As String
Select Case Lname
Case "L-1":    Lpath = path & "L-1\L-1_Log"
Case "LVC":    Lpath = path & "LVC\LVC_Log"
Case "HVC":    Lpath = path & "HVC\HVC_Log"
Case "ﾊﾞﾗﾝｻｰ": Lpath = path & "バランサー\バランサー_Log"
Case "ﾄｯﾄ":    Lpath = path & "トット\トット_Log"
Case "Test":   Lpath = path & "Test\Test_Log"
Case Else:     Lpath = path & "丸徳\丸徳_Log"
End Select

'-------------------------------------------------判定色設定
Select Case ws.Name
Case "全ヵ所検査"
  YColor = 10
  NColor = 3
End Select
'-------------------------------------------------判定有無
Select Case True 'Undo使用後の不要な判定を止める　2019.12.24
Case ThisWorkbook.Sheets("全ヵ所検査").Visible
  If UFMitsubishi.UndoCount = "1" Then
    UFMitsubishi.UndoCount = "0"
    Exit Sub
  End If
End Select
'-------------------------------------------------判定処理
If Judg Then '5N選択時
  Select Case ws.Name
  Case "全ヵ所検査"
    ws.Cells(TR, TC + 1) = "○"
    ws.Cells(TR, TC + 1).Interior.ColorIndex = YColor 'RGB(0, 128, 0)
    Call WriteLog("*UFKen:" & ws.Name & "*入力:" & ws.Cells(TR, TC).Value & " *座標:" & "(" & TR & "," & TC & ")" & "：○", "", "")  'シート名追加_2023.11.14
    Call WriteLog("*UFKen:" & ws.Name & "*入力:" & ws.Cells(TR, TC).Value & " *座標:" & "(" & TR & "," & TC & ")" & "：○:" & Lname, "", Lpath, fname)
    Call 音楽合格
    
    Select Case True 'キャスクはフォームを出さない
    Case ThisWorkbook.Sheets("キャスク検査").Visible
     UFcount.Maru.Visible = True
     Call UFcount起動(Cells(TR, TC))
    Case Else
     UFOK.Show vbModeless
    End Select
    
  End Select
ElseIf Judg = False Then
  Select Case ws.Name
  Case "全ヵ所検査"
    ws.Cells(TR, TC + 1) = "×"
    ws.Cells(TR, TC + 1).Interior.ColorIndex = NColor
    Call WriteLog("*UFKen:" & ws.Name & "*入力:" & ws.Cells(TR, TC).Value & " *座標:" & "(" & TR & "," & TC & ")" & "：×", "", "")
    Call WriteLog("*UFKen:" & ws.Name & "*入力:" & ws.Cells(TR, TC).Value & " *座標:" & "(" & TR & "," & TC & ")" & "：×:" & Lname, "", Lpath, fname)
    Call 音楽不合格
    
   Select Case True 'キャスクはフォームを出さない
   Case ThisWorkbook.Sheets("キャスク検査").Visible
    UFcount.Batu.Visible = True
    Call UFcount起動(Cells(TR, TC))
   Case Else
    UFNG.Show vbModeless
   End Select
   
  End Select
End If
End Sub

'-----------------------------------------------------------
' 戻り値：Variant型 2次元配列（SELECT * の列順）
'   arr(i, 0) = 管理番号
'   arr(i, 1) = LotNo
'   arr(i, 2) = 検査No
'   arr(i, 3) = スタートNo
'   arr(i, 4) = エンドNo
'   arr(i, 5) = 作業時間
'   データなし時は Empty を返す
'-----------------------------------------------------------
Public Function GetCaskData(ByVal strLot As String, _
                              ByVal strKen As String) As Variant

    Dim FilePath  As String
    Dim SQLfilter As String

    FilePath = QualityPath & "伝送用ファイル.accdb"
    SQLfilter = "[LotNo]='" & strLot & "' AND [検査No]='" & strKen & "'"

    GetCaskData = GetRecordsSQL(FilePath, "キャスク梱包板番号", SQLfilter, "スタートNo")

End Function
'-----------------------------------------------------------
' 引数：
'   filePath  ：Accessファイルのフルパス
'   TableName ：テーブル名
'   SQLfilter ：WHERE句（例："[LotNo]='H473C51' AND [検査No]='H451076'"）
'               空文字を渡すと全件取得
'   orderBy   ：ORDER BY句（例："スタートNo"）省略可
' 戻り値：
'   Variant型 2次元配列 arr(行, 列)
'   データなし・エラー時は Empty を返す------------------
Public Function GetRecordsSQL(ByVal FilePath As String, _
                                ByVal TableName As String, _
                                ByVal SQLfilter As String, _
                       Optional ByVal orderBy As String = "") As Variant

    Dim db     As Object
    Dim rs     As Object
    Dim strSQL As String
    Dim colCnt As Long
    Dim rowCnt As Long
    Dim r      As Long
    Dim C      As Long
    Dim tMP()  As Variant   '一時1次元配列（行ごと格納）
    Dim arr()  As Variant   '最終2次元配列

    GetRecordsSQL = Empty

    On Error GoTo ErrHandler

    '--- SQL組み立て ---
    strSQL = "SELECT * FROM [" & TableName & "]"
    If Trim(SQLfilter) <> "" Then
        strSQL = strSQL & " WHERE " & SQLfilter
    End If
    If Trim(orderBy) <> "" Then
        strSQL = strSQL & " ORDER BY " & orderBy
    End If
    DebugLog "[GetRecordsSQL] SQL=" & strSQL

    '--- DB接続 ---
    Set db = CreateObject("ADODB.Connection")
    db.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FilePath & ";"
    DebugLog "[GetRecordsSQL] DB接続OK"

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, db, 3, 1   ' adOpenStatic, adLockReadOnly
    DebugLog "[GetRecordsSQL] RS取得OK / EOF=" & rs.EOF & " / RecordCount=" & rs.RecordCount & " / Fields.Count=" & rs.Fields.Count

    If rs.EOF Then GoTo CleanUp

    '--- フィールド数・レコード数を確定してから配列を一括確保 ---
    '    ReDim Preserve は最終次元しか変更不可のため
    '    先にRecordCountで行数を確定し arr(行, 列) を一度に確保する
    colCnt = rs.Fields.Count - 1
    rowCnt = rs.RecordCount - 1
    DebugLog "[GetRecordsSQL] colCnt=" & colCnt & " / rowCnt=" & rowCnt

    ReDim arr(rowCnt, colCnt)   '一括確保（ReDim Preserve 不要）

    r = 0
    Do While Not rs.EOF
        For C = 0 To colCnt
            arr(r, C) = rs.Fields(C).Value
        Next C
        r = r + 1
        rs.MoveNext
    Loop
    DebugLog "[GetRecordsSQL] 配列格納完了 行数=" & r

    GetRecordsSQL = arr

CleanUp:
    If Not rs Is Nothing Then
        If rs.state = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then
        db.Close: Set db = Nothing
    End If
    Exit Function

ErrHandler:
    DebugLog "[GetRecordsSQL] ★エラー Err=" & Err.Number & " / " & Err.Description
    MsgBox "GetRecordsSQL エラー：" & vbCrLf & Err.Description, vbCritical
    Resume CleanUp

End Function


'-----------------------------------------------------------
' 引数：
'   frm       ：対象UserForm
'   strLot    ：LotNo
'   strKen    ：検査No
'   sngLeft   ：最初のラベルのLeft座標（pt）
'   sngTop    ：ラベルのTop座標（pt）
'   sngWidth  ：各ラベルの幅（pt）
'   sngHeight ：各ラベルの高さ（pt）
'   sngGap    ：ラベル間の間隔（pt）
'   strPrefix ：ラベル名のプレフィックス（例："lblCask"）
'
' ラベル表示内容：スタートNo - エンドNo（改行）作業時間
' 既存の同プレフィックスラベルは追加前にすべて削除する
'-----------------------------------------------------------
Public Sub AddCaskLabels(ByVal frm As Object, _
                           ByVal strLot As String, _
                           ByVal strKen As String, _
                           ByVal sngLeft As Single, _
                           ByVal sngTop As Single, _
                           ByVal sngWidth As Single, _
                           ByVal sngHeight As Single, _
                           ByVal sngGap As Single, _
                  Optional ByVal strPrefix As String = "lblCask")

    Dim varData  As Variant
    Dim ctrl     As Object
    Dim lbl      As Object
    Dim i        As Long
    Dim dataCnt  As Long
    Dim lblName  As String

    On Error GoTo ErrHandler

'    DebugLog "[AddCaskLabels] ===== 開始 ====="
'    DebugLog "[AddCaskLabels] LotNo=" & strLot & " 検査No=" & strKen & " Prefix=" & strPrefix

    '--- STEP1：既存の同プレフィックスラベルをすべて非表示にする ---
    '    Controls.Remove は実行中フォームでは失敗するため
    '    削除せず非表示にして後で再利用する
    For Each ctrl In frm.Controls
        If Left(ctrl.Name, Len(strPrefix)) = strPrefix Then
            ctrl.Visible = False
'            DebugLog "[AddCaskLabels] 非表示 Name=" & ctrl.Name
        End If
    Next ctrl

    '--- STEP2：データ取得 ---
'    DebugLog "[AddCaskLabels] GetCaskData 呼び出し"
    varData = GetCaskData(strLot, strKen)
'    DebugLog "[AddCaskLabels] GetCaskData 完了 IsEmpty=" & IsEmpty(varData)

    If IsEmpty(varData) Then
'        DebugLog "[AddCaskLabels] データなし → 終了"
        Exit Sub
    End If

    dataCnt = UBound(varData, 1)
'    DebugLog "[AddCaskLabels] レコード件数=" & (dataCnt + 1)

    '--- STEP3：件数分ラベルを設定（既存再利用 or 新規追加）---
    For i = 0 To dataCnt

        lblName = strPrefix & i

        '既存ラベルがあれば再利用、なければ新規追加
        On Error Resume Next
        Set lbl = frm.Controls(lblName)
        On Error GoTo ErrHandler

        If lbl Is Nothing Then
            DebugLog "[AddCaskLabels] 新規追加 Name=" & lblName
            Set lbl = frm.Controls.Add("Forms.Label.1", lblName, True)
        Else
            DebugLog "[AddCaskLabels] 既存再利用 Name=" & lblName
        End If

        With lbl
            .Left = sngLeft + i * (sngWidth + sngGap)
            .tOP = sngTop
            .Width = sngWidth
            .Height = sngHeight
            .Caption = "No" & (i + 1) & ":" & varData(i, 3) & "～" & varData(i, 4)
            .TextAlign = 2
            .BorderStyle = 1
            .BackColor = &HFFFFFF
            .ForeColor = &H0
            .Font.Name = "Meiryo UI"
            .Font.Size = 12
            .WordWrap = True
            .Visible = True       '★ 必ず表示に戻す
        End With

        DebugLog "[AddCaskLabels] 設定完了 i=" & i & " Caption=" & lbl.Caption
        Set lbl = Nothing

    Next i

    DebugLog "[AddCaskLabels] ===== 正常終了 ====="
    Exit Sub

ErrHandler:
    DebugLog "[AddCaskLabels] ★エラー Err=" & Err.Number & " / " & Err.Description
    MsgBox "ラベル追加エラー：" & vbCrLf & Err.Description, vbCritical

End Sub

'-----------------------------------------------------------
' 処理概要：
'   シート「キャスク検査」のB列最終行のA列値と
'   GetCaskData で取得したレコードの最大エンドNo を比較する
'
' 戻り値：
'   True  → 一致（正常）
'   False → 不一致 or データなし（呼び出し元で対処）
'-----------------------------------------------------------
Public Function CheckCaskEndNo(ByVal strLot As String, _
                                  ByVal strKen As String) As Boolean

    Dim ws       As Worksheet
    Dim lastRow  As Long
    Dim shEndNo  As Long   'シートのA列値
    Dim dbEndNo  As Long   'DBのエンドNo最大値
    Dim varData  As Variant

    CheckCaskEndNo = False

    On Error GoTo ErrHandler

    '--- シート「キャスク検査」を取得 ---
    Set ws = ThisWorkbook.Worksheets("キャスク検査")

    '--- B列の最終行を取得 ---
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    DebugLog "[CheckCaskEndNo] シート最終行=" & lastRow

    If lastRow < 1 Then
        DebugLog "[CheckCaskEndNo] シートにデータなし"
        MsgBox "「キャスク検査」シートにデータがありません。", vbExclamation
        Exit Function
    End If

    '--- B列最終行のA列値を取得 ---
    shEndNo = CLng(ws.Cells(lastRow, "A").Value)
    DebugLog "[CheckCaskEndNo] シートA列値=" & shEndNo

    '--- DBからデータ取得 ---
    varData = GetCaskData(strLot, strKen)

    If IsEmpty(varData) Then
        DebugLog "[CheckCaskEndNo] DBデータなし"
        MsgBox "DBにデータが見つかりません。" & vbCrLf & _
               "LotNo：" & strLot & "　検査No：" & strKen, vbExclamation
        Exit Function
    End If

    '--- DBの最終レコード（スタートNo昇順）のエンドNoを取得 ---
    dbEndNo = CLng(varData(UBound(varData, 1), 4))
    DebugLog "[CheckCaskEndNo] DBエンドNo（最終）=" & dbEndNo

    '--- 照合 ---
    If shEndNo = dbEndNo Then
        DebugLog "[CheckCaskEndNo] 一致 OK"
        CheckCaskEndNo = True
    Else
        CheckCaskEndNo = False
        DebugLog "[CheckCaskEndNo] 不一致 シート=" & shEndNo & " DB=" & dbEndNo
        MsgBox "エンドNoが一致しません。" & vbCrLf & vbCrLf & _
               "シート最終行（A列）：" & shEndNo & vbCrLf & _
               "DB エンドNo（最終）：" & dbEndNo, vbExclamation
    End If

    Exit Function

ErrHandler:
    DebugLog "[CheckCaskEndNo] ★エラー Err=" & Err.Number & " / " & Err.Description
    MsgBox "チェックエラー：" & vbCrLf & Err.Description, vbCritical

End Function

