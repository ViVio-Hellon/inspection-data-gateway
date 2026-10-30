Attribute VB_Name = "ZenSh2"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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
    Case "全ヵ所検査"
      Select Case TR '-------------------------------行
      Case 1 To 6
      Case 7 To 76
        Select Case TC '------------------------------列
        Case 2 To 8
         Cells(TR, TC + 1).Activate
        Case 9
          If Cells(TR, 11) = "" Then
            Cells(TR, 11).Activate
          Else
            Cells(TR + 1, 2).Activate
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
        Case 20 '毎回のイベントは不要なので修正_2023.6.20
          Application.EnableEvents = False
          Call AlTerRange(target, Befor)
          Call 対角(target)
          Application.EnableEvents = True
          Application.ScreenUpdating = True
          Cells(TR + 1, 2).Activate
          Exit Sub
        Case 21
          Cells(TR, 24).Activate
        Case 23 '毎回のイベントは不要なので修正_2023.6.20
          Application.EnableEvents = False
          If Not IsNumeric(Cells(TR, 23)) Then
          ElseIf Cells(TR, 23) = "-" Then '残す　測定不要のパターン
          Else
            Cells(TR, 23) = Format(UnitRecord(val(Cells(TR, 23))), "0.0") 'ラテラルボー測定値を0.5ピッチに変換
          End If
          Call Lateralbow_Z(target)
          Application.EnableEvents = True
          Application.ScreenUpdating = True
          Cells(TR, TC + 1).Activate
          Exit Sub
        Case 24 '簡易判定 公差1.0に修正2019.12.24
        If UFMitsubishi.Work = "ON" Then
         If IsNumeric(Cells(target.row, 24).Value) = True Then '入力が数字である
          If 1# >= val(Cells(TR, 24).Value) Then  '巾の単位をMにして1M当たりの公差を出す
              Call 音楽合格
              UFOK.Show vbModeless
          ElseIf 1# < val(Cells(TR, 24).Value) Then
              Call 音楽不合格
              UFNG.Show vbModeless
          End If
         End If
        End If
          Application.EnableEvents = True
          Application.ScreenUpdating = True
          Cells(TR, TC + 1).Activate
          Exit Sub
        Case 25  '公差1.0に修正2019.12.24
        If UFMitsubishi.Work = "ON" Then
         If IsNumeric(Cells(target.row, 24).Value) = True Then '入力が数字である
          If 1# >= val(Cells(TR, 25).Value) Then  '丈の単位をMにして1M当たりの公差を出す
              Call 音楽合格
              UFOK.Show vbModeless
          ElseIf 1# < val(Cells(TR, 25).Value) Then
              Call 音楽不合格
              UFNG.Show vbModeless
          End If
         End If
        End If
          Cells(TR, TC + 1).Activate
        Case 26
          Cells(TR, TC + 1).Activate
        Case 27
        End Select
      End Select
    End Select
    Select Case ws.Name 'シート毎に判定ﾓｼﾞｭｰﾙが違う
    Case "全ヵ所検査"
       Application.EnableEvents = False
        Call AlTerRange(target, Befor) 'ﾊﾟｹﾃﾞｰﾀの入力でイベントが動き、変換後にもう一度ｲﾍﾞﾝﾄ発生してしまう
        If Not IsNumeric(target) Then Exit Sub
       Application.EnableEvents = True
        Call 検査判定(target)
       Application.EnableEvents = False

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
UFMitsubishi.Cnt = "0" '移動後はクリア　Application.Undo可　2019.12.24
 Application.ScreenUpdating = False
  Select Case ws.Name
  Case "全ヵ所検査"
    Select Case TC
    Case 21
      Cells(TR, 24).Activate
    Case 24
    '*****************************************名前検番の登録
      Application.EnableEvents = False
      ws.Unprotect 'シート保護解除
      ws.Cells(2, 8) = UFMitsubishi.KenNo
      ws.Cells(3, 8) = UFMitsubishi.Kname
      ws.Cells(4, 8) = UFMitsubishi.YoutoNo
      Range(ws.Cells(2, 8), ws.Cells(4, 8)).Font.Size = 5
      ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True 'シート保護
      Application.EnableEvents = True
    End Select
  End Select

'*****************************************作業開始判定
  Select Case ws.Name
  Case "全ヵ所検査"
    If target.Address = "$D$7" Then
    Application.EnableEvents = False
      If ws.Cells(7, 2).Value <> "" Then
        ws.Unprotect 'シート保護解除
        ws.Cells(1, 2) = Format$(Now, "yyyy/mm/dd h:mm")
        ws.Cells(1, 2).Font.Size = 5
        Call Daily_Report(1) '日報情報
        Call Daily_Report(3) '日報情報(検査開始)
      End If
    Application.EnableEvents = True
    End If
  End Select
'*****************************************色　　Range(Cells(1, 11), Cells(1, 15)),Range(Cells(1, 17), Cells(1, 18)), Range(Cells(7, 22), Cells(76, 23))*2023.4.13追加
  Select Case ws.Name
  Case "全ヵ所検査"
    Set Ran = Union(Range(Cells(2, 2), Cells(3, 6)), Range(Cells(1, 11), Cells(1, 15)), Range(Cells(1, 17), Cells(1, 18)), Range(Cells(2, 11), Cells(4, 15)), Range(Cells(2, 17), Cells(4, 19)), Range(Cells(2, 21), Cells(4, 26)) _
    , Range(Cells(7, 2), Cells(76, 2)), Range(Cells(7, 4), Cells(76, 4)), Range(Cells(7, 6), Cells(76, 6)), Range(Cells(7, 8), Cells(76, 8)) _
    , Range(Cells(7, 11), Cells(76, 11)), Range(Cells(7, 13), Cells(76, 13)), Range(Cells(7, 15), Cells(76, 15)), Range(Cells(7, 17), Cells(76, 17)) _
    , Range(Cells(7, 19), Cells(76, 20)), Range(Cells(7, 22), Cells(76, 23)), Range(Cells(7, 24), Cells(76, 27)))
    Set NpRan = Union(Range(Cells(7, 28), Cells(7, 30)), Range(Cells(9, 28), Cells(9, 30)), Range(Cells(11, 28), Cells(11, 29)), Range(Cells(13, 28), Cells(13, 30)))
  End Select

  On Error Resume Next
  With ActiveCell
      Select Case ws.Name
      Case "全ヵ所検査"
         Ran.Interior.ColorIndex = 2 '白
         NpRan.Interior.ColorIndex = 27 '黄
'         If UFMitsubishi.Work = "OFF" Then NpRan.Interior.ColorIndex = 15 '灰
         .Interior.Color = RGB(153, 204, 255) 'セル薄い青
        .Font.Color = RGB(0, 0, 0) '文字黒
      End Select
  End With
 Application.ScreenUpdating = True
End Sub


