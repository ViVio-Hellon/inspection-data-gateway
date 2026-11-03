Attribute VB_Name = "H494_SH"
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
    If UFMitsubishi.Cnt = "0" Then
      After = target '変更後
      Application.EnableEvents = False
       Application.Undo
      Befor = target '変更前
      target = After
      Application.EnableEvents = True
'       MsgBox "*前:" & Befor & "*後:" & After
    End If
    Select Case ws.Name
    Case "H494検査"
      Select Case TR '-------------------------------行
      Case 1 To 6
      Case 7 To 16
        Select Case TC '------------------------------列
        Case 2 To 8
         Cells(TR, TC + 1).Activate
        Case 9
          If Cells(TR, 11) = "" Then
            Cells(TR, 11).Activate
          Else
            Cells(TR + 1, 2).Activate
          End If
        Case 11 To 20
          Cells(TR, TC + 1).Activate
        Case 21
           If Cells(TR, 21).Value <> "" Then
            Call 音楽対角
            Call UF対角起動(target)
            Application.EnableEvents = False
            Range(Cells(TR, 22), Cells(TR, 24)) = ""
            Application.EnableEvents = True
            Cells(TR, TC + 1).Activate
           End If
        Case 22
          Application.EnableEvents = False
          Call AlTerRange(target, Befor)
          Call 対角(target)
          Application.EnableEvents = True
          Application.ScreenUpdating = True
          Cells(TR + 1, 2).Activate
          Exit Sub
        Case 25 '毎回のイベントは不要
          Application.EnableEvents = False
          If Not IsNumeric(Cells(TR, 25)) Then
          ElseIf Cells(TR, 25) = "-" Then '残す　測定不要のパターン
          Else
            Cells(TR, 25) = Format(UnitRecord(val(Cells(TR, 25))), "0.0") 'ラテラルボー測定値を0.5ピッチに変換
          End If
          Call Lateralbow_Z(target)
          Application.EnableEvents = True
          Application.ScreenUpdating = True
          Cells(TR, TC + 1).Activate
          Exit Sub
        End Select
      End Select
    End Select
    Select Case ws.Name 'シート毎に判定ﾓｼﾞｭｰﾙが違う
    Case "H494検査"
       Application.EnableEvents = False
        Call AlTerRange(target, Befor) 'ﾊﾟｹﾃﾞｰﾀの入力でイベントが動き、変換後にもう一度ｲﾍﾞﾝﾄ発生してしまう
        If Not IsNumeric(target) Then Exit Sub
       Application.EnableEvents = True
        If target.Value <> "" Then Call 検査判定(target)
       Application.EnableEvents = False
        Call 対角(target)
'        Call 異常値除去(Target)’判定付けの中で行う
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
Dim Ran As Range, NpRan As Range
If UFMitsubishi.Visible Then UFMitsubishi.Cnt = "0" '移動後はクリア　Application.Undo可　2019.12.24
 Application.ScreenUpdating = False
  Select Case ws.Name
  Case "H494検査"
    Select Case TC
    Case 24
      Cells(TR, 25).Activate
    Case 26
    '*****************************************名前検番の登録
      Application.EnableEvents = False
      ws.Unprotect 'シート保護解除
      ws.Cells(2, 8) = UFMitsubishi.H494_KenNo
      ws.Cells(3, 8) = UFMitsubishi.H494_Kname
      Range(ws.Cells(2, 8), ws.Cells(3, 8)).Font.Size = 6
      ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True 'シート保護
      Application.EnableEvents = True
    End Select
  End Select

'*****************************************作業開始判定
  Select Case ws.Name
  Case "H494検査"
    If target.Address = "$D$7" Then
    Application.EnableEvents = False
      If ws.Cells(7, 2).Value <> "" Then
        ws.Cells(1, 2) = Format$(Now, "yyyy/mm/dd h:mm")
        ws.Cells(1, 2).Font.Size = 5
        Call Daily_Report(1) '日報情報
        Call Daily_Report(3) '日報情報(検査開始)
      End If
    Application.EnableEvents = True
    End If
  End Select
'*****************************************色
  Select Case ws.Name
  Case "H494検査"
    Set Ran = Union(Range(Cells(2, 2), Cells(3, 6)), Range(Cells(2, 11), Cells(4, 15)), Range(Cells(2, 17), Cells(4, 19)), Range(Cells(2, 21), Cells(4, 26)) _
    , Range(Cells(7, 2), Cells(16, 2)), Range(Cells(7, 4), Cells(16, 4)), Range(Cells(7, 6), Cells(16, 6)), Range(Cells(7, 8), Cells(16, 8)) _
    , Range(Cells(7, 11), Cells(16, 11)), Range(Cells(7, 13), Cells(16, 13)), Range(Cells(7, 15), Cells(16, 15)), Range(Cells(7, 17), Cells(16, 17)), Range(Cells(7, 19), Cells(16, 19)) _
    , Range(Cells(7, 21), Cells(16, 22)), Range(Cells(7, 24), Cells(16, 28)))
  End Select

  On Error Resume Next
  With ActiveCell
      Select Case ws.Name
      Case "H494検査"
         Ran.Interior.ColorIndex = 2 '白
         NpRan.Interior.ColorIndex = 27 '黄
         If UFMitsubishi.Work = "OFF" Then NpRan.Interior.ColorIndex = 15 '灰
         .Interior.Color = RGB(153, 204, 255) 'セル薄い青
        .Font.Color = RGB(0, 0, 0) '文字黒
      End Select
  End With
 Application.ScreenUpdating = True
End Sub
