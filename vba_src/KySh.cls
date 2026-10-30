Attribute VB_Name = "KySh"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal target As Range)
Dim After As String, Befor As String
Dim TR As Long:          TR = target.row
Dim TC As Long:          TC = target.Column
Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
Application.ScreenUpdating = False
'変換判定処理
  If target.Count <> 1 Then '2以上のChange
  Else
  On Error Resume Next
    If UFCask.Cnt = "0" Then
      After = target '変更後
      Application.EnableEvents = False
       Application.Undo
      Befor = target '変更前
      target = After
      Application.EnableEvents = True
'       MsgBox "*前:" & Befor & "*後:" & After
    End If
    Select Case ws.Name
    Case "キャスク検査"
      Select Case TR '-------------------------------行
      Case 1 To 6
      Case 7 To 706
        Select Case TC '------------------------------列
        Case 2 To 8
         Cells(TR, TC + 1).Activate
        Case 7
          If UFCask.CountSt.Caption = "3" Then
            If UFCask.Inspection = "OFF" Then '厚のみ
             Cells(target.row + 1, 2).Activate
            ElseIf UFCask.Inspection = "ON" Then '巾丈も測定
             Cells(target.row, 11).Activate
            End If
          End If
        Case 9 '100mm以上限定
          If UFCask.CountSt.Caption = "4" Then
            If UFCask.Inspection = "OFF" Then '厚のみ
             Cells(target.row + 1, 2).Activate
            ElseIf UFCask.Inspection = "ON" Then '巾丈も測定
             Cells(target.row, 11).Activate
            End If
          End If
        Case 11 To 18
         Cells(TR, TC + 1).Activate
        Case 19
           If Cells(TR, 19).Value <> "" Then
            Call 音楽対角
            Call UF対角起動(target)
            Application.EnableEvents = False
              Range(Cells(TR, 20), Cells(TR, 23)) = ""
            Application.EnableEvents = True
            Cells(TR, TC + 1).Activate
           End If
        Case 20
         Cells(TR, TC + 1).Activate
        Case 21
         Cells(TR, 24).Activate
        Case 24 To 27
         Cells(TR, TC + 1).Activate
        Case 28
        End Select
      End Select
    End Select
    Select Case ws.Name 'シート毎に判定ﾓｼﾞｭｰﾙが違う
    Case "キャスク検査"
       Application.EnableEvents = False
        Call AlTerRange(target, Befor) 'ﾊﾟｹﾃﾞｰﾀの入力でイベントが動き、変換後にもう一度ｲﾍﾞﾝﾄ発生してしまう
        If Not IsNumeric(target) Then Exit Sub
       Application.EnableEvents = True
        Call 検査判定(target)
       Application.EnableEvents = False
        Call 対角(target)
       Application.EnableEvents = True
        Call 公差外検索(target)
    End Select
    'ｸﾘｱ
    If Err.Number <> 0 Then
     Err.clear
    End If
  End If
Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)
Dim TR As Long:          TR = target.row
Dim TC As Long:          TC = target.Column
Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
UFCask.Cnt = "0" '移動後はクリア　Application.Undo可　2019.12.24
 Application.ScreenUpdating = False
  Select Case ws.Name
  Case "キャスク検査"
    Select Case TC
    Case 21
      Cells(target.row, 24).Activate
    Case 24
  '*****************************************名前検番の登録
      Application.EnableEvents = False
      ws.Unprotect 'シート保護解除
      ws.Cells(2, 25) = UFCask.KenNo
      ws.Cells(3, 25) = UFCask.Kname
      ws.Cells(2, 26) = UFCask.Cno
      ws.Cells(3, 26) = UFCask.Mno
      ws.Cells(4, 26) = UFCask.SeriNo
      Dim SokuSt As String
      If val(ws.Cells(3, 2)) >= 100# Then SokuSt = "ﾉｷﾞｽ"
      If val(ws.Cells(3, 2)) < 100# Then SokuSt = "ﾏｲｸﾛ"
      ws.Cells(2, 28) = "ｾﾞﾛ校正:" & SokuSt
      ws.Cells(3, 28) = "ｾﾞﾛ校正:" & "ｺﾝﾍﾞｯｸｽ"
      ws.Cells(3, 30) = "1000mm校正:" & "ｺﾝﾍﾞｯｸｽ"
      ws.Cells(2, 29) = UFCask.Micro
      ws.Cells(3, 29) = UFCask.Convex
      ws.Cells(3, 31) = UFCask.Convex_1000
      ws.Cells(3, 33) = UFCask.CountSt
      Range(ws.Cells(2, 25), ws.Cells(4, 31)).Font.Size = 8
      Range(ws.Cells(2, 25), ws.Cells(4, 31)).HorizontalAlignment = xlLeft
      ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True 'シート保護
      Application.EnableEvents = True
    End Select
  End Select

'*****************************************作業開始判定
  Select Case ws.Name
  Case "キャスク検査"
    If target.Address = "$D$7" Then
    Application.EnableEvents = False
      If Cells(7, 2).Value <> "" Then
        Cells(1, 2) = Format$(Now, "yyyy/mm/dd h:mm")
        Cells(1, 2).Font.Size = 5
        Call Daily_Report(1) '日報情報
        Call Daily_Report(3) '日報情報(検査開始)
      End If
    Application.EnableEvents = True
    End If
  End Select
'*****************************************色
  Dim HanteiCell As Range
  Select Case ws.Name
  Case "キャスク検査"
    Set Ran = Union(Range(Cells(2, 2), Cells(3, 6)), Range(Cells(2, 11), Cells(4, 15)), Range(Cells(2, 17), Cells(4, 19)), Range(Cells(2, 21), Cells(4, 26)) _
    , Range(Cells(7, 2), Cells(706, 2)), Range(Cells(7, 4), Cells(706, 4)), Range(Cells(7, 6), Cells(706, 6)), Range(Cells(7, 8), Cells(706, 8)) _
    , Range(Cells(7, 11), Cells(706, 11)), Range(Cells(7, 13), Cells(706, 13)), Range(Cells(7, 15), Cells(706, 15)), Range(Cells(7, 17), Cells(706, 17)) _
    , Range(Cells(7, 19), Cells(706, 20)), Range(Cells(7, 22), Cells(706, 22)), Range(Cells(7, 24), Cells(706, 28)))
    Set HanteiCell = Union(Range(Cells(7, 3), Cells(706, 3)), Range(Cells(7, 5), Cells(706, 5)), Range(Cells(7, 7), Cells(706, 7)), Range(Cells(7, 9), Cells(706, 9)) _
    , Range(Cells(7, 12), Cells(706, 12)), Range(Cells(7, 14), Cells(706, 14)), Range(Cells(7, 16), Cells(706, 16)), Range(Cells(7, 18), Cells(706, 18)), Range(Cells(7, 21), Cells(706, 21)))
  End Select

' セルがHanteiCellの範囲内であれば何もせず
  If Not Intersect(target, HanteiCell) Is Nothing Then
    Ran.Interior.Color = RGB(255, 255, 255) '判定セルに移動した場合判定セル以外を白
    If UFCask.Inspection = "OFF" Then
       ws.Range("K6:AB706").Interior.Color = RGB(128, 128, 128)
    ElseIf UFCask.Inspection = "ON" Then
    End If
    Exit Sub ' 何もしない
  End If
  ' セルがRanの範囲内であれば背景色変更
  If Not Intersect(target, Ran) Is Nothing Then
      Select Case ws.Name
      Case "キャスク検査"
        Ran.Interior.Color = RGB(255, 255, 255) ' 白
        If UFCask.Inspection = "OFF" Then
          Ran.Interior.Color = RGB(255, 255, 255)
          target.Interior.Color = RGB(153, 204, 255) 'セル薄い青
          target.Font.Color = RGB(0, 0, 0) '文字黒
           ws.Range("K6:AB706").Interior.Color = RGB(128, 128, 128)
        ElseIf UFCask.Inspection = "ON" Then
          Ran.Interior.Color = RGB(255, 255, 255)
          target.Interior.Color = RGB(153, 204, 255) 'セル薄い青
          target.Font.Color = RGB(0, 0, 0) '文字黒
        End If
      End Select
  End If

'  On Error Resume Next
'  With ActiveCell
'      Select Case ws.Name
'      Case "キャスク検査"
'        Ran.Interior.ColorIndex = 2 '白
'        If UFCask.Inspection = "OFF" Then
'         Ran.Interior.Color = RGB(255, 255, 255)
'         .Interior.Color = RGB(153, 204, 255) 'セル薄い青
'        .Font.Color = RGB(0, 0, 0) '文字黒
'          ws.Range("K6:AB706").Interior.Color = RGB(128, 128, 128)
'        ElseIf UFCask.Inspection = "ON" Then
'         Ran.Interior.Color = RGB(255, 255, 255)
'         .Interior.Color = RGB(153, 204, 255) 'セル薄い青
'        .Font.Color = RGB(0, 0, 0) '文字黒
'        End If
'      End Select
'  End With
 Application.ScreenUpdating = True
End Sub

