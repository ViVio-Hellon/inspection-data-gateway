Attribute VB_Name = "UF校正"
Attribute VB_Base = "0{BBF1104B-019D-4D94-B7DB-86BFBCC364E1}{CE0DC945-8054-469A-B489-581173ECC4DF}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
#If VBA7 And Win64 Then
  '64ビット版
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
  '32ビット版
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Dim i As Long, a As Long
'Private g_lngWindowState As XlWindowState  ' ウィンドウ状態
' Excelウィンドウ状態退避用
Private g_sngTopSV As Single ' 上端位置
Private g_sngLeftSV As Single ' 左端位置
Private g_sngWidthSV As Single ' 幅
Private g_sngHeightSV As Single ' 高さ
Private Function TerConv(Ter As String) As String
If Ter = "" Then Exit Function
  Select Case TerMaker(Ter)
  Case "DT" 'DT10002+0000021.712M
   TerConv = Mid(Ter, 14, 6)
  Case "#L" '＃L395.8
   TerConv = Mid(Ter, 3, Len(Ter))
  Case Else
    TerConv = "" 'Ter
    MsgBox "正しいデータが伝送できていません", vbCritical
  End Select
End Function
Public Function TerMaker(Trm As String) As String 'DT #L選別
  TerMaker = ""
  If Trm = "" Then Exit Function
  If Len(Trm) >= 7 And Trm <> "" Then
    Select Case StrConv(UCase(Left(Trm, 2)), vbNarrow) '強制大文字変換
    Case "DT" 'DT10002+0000021.712M
     TerMaker = "DT"
    Case "#L" '#L395.8
     TerMaker = "#L"
    End Select
  End If
End Function

Private Sub CBcl_Click()
UF校正.F1 = 1 'フラグを１にする
If UF校正.BackColor = &H80FF80 Then
 UF校正.INP1000 = ""
 UF校正.SokuteiH = ""
 UF校正.NGvalue = ""
 UF校正.INP1000.SetFocus 'インプットボックスに初めから数字が入るようにする
ElseIf UF校正.BackColor = &HFFFFFF Then
 UF校正.INPG = ""
 UF校正.SokuteiH = ""
 UF校正.NGvalue = ""
 UF校正.INPG.SetFocus 'インプットボックスに初めから数字が入るようにする
End If
If UF校正.INP1000 = "" Then
 UF校正.Saiken.Visible = False
 UF校正.NGvalue = ""
End If
End Sub

Private Sub CommandButton1_Click()
Dim ws As Worksheet
Dim flag As Boolean, Flag2 As Boolean
Dim SheetsName As String, strDate As String, StrDate2 As String, dt As String, line As String
Dim k As Long: k = ThisWorkbook.Sheets("品証検証").Cells(Rows.Count, 3).End(xlUp).row + 1
'*****************************ﾗｲﾝの設定
  Select Case UF校正.LineName
  Case "L-1":    line = "L-1"
  Case "LVC":    line = "LVC"
  Case "HVC":    line = "HVC"
  Case "ﾊﾞﾗﾝｻｰ": line = "ﾊﾞﾗﾝｻｰ"
  Case "ﾄｯﾄ":    line = "ﾄｯﾄ"
  Case "Test":   line = "Test"
  Case Else: '丸徳用
  End Select
'********************************************
Application.ScreenUpdating = False
Dim x As String: x = TimeCheck(Format$(Now, "hh:mm")) '直を返す
Dim y As String: y = YeaDayrCheck(Format$(Now, "hh:mm")) '日付を返す
 If UF校正.SokuteiH = "" Then
  UF校正.Saiken.Visible = True
  UF校正.NGvalue = ""
  UF校正.INP1000 = ""
  UF校正.SokuteiH = ""
  UF校正.INP1000.SetFocus 'インプットボックスに初めから数字が入るようにする
  Exit Sub
 Else
  For i = 6 To 96
   If ThisWorkbook.Sheets("品証検証").Cells(i, 1) = y Then
       'MsgBox x & y & strDate & "1"
     If Sheets("品証検証").Cells(i, 2) = x Then
       'MsgBox x & y & strDate & "1"
      ThisWorkbook.Sheets("品証検証").Cells(i, 3) = UF校正.SokuteiH '値
      ThisWorkbook.Sheets("品証検証").Cells(i, 9) = line 'ライン
      ThisWorkbook.Sheets("品証検証").Cells(i, 10) = UF校正.Kname '検査員
      ThisWorkbook.Sheets("品証検証").Cells(i, 11) = Format$(Now, "mm/dd h時mm分ss秒") '時間
     ElseIf ThisWorkbook.Sheets("品証検証").Cells(i + 1, 2) = x Then
       'MsgBox x & y & strDate & "2"
      ThisWorkbook.Sheets("品証検証").Cells(i + 1, 3) = UF校正.SokuteiH '値
      ThisWorkbook.Sheets("品証検証").Cells(i + 1, 9) = line 'ライン
      ThisWorkbook.Sheets("品証検証").Cells(i + 1, 10) = UF校正.Kname '検査員
      ThisWorkbook.Sheets("品証検証").Cells(i + 1, 11) = Format$(Now, "mm/dd h時mm分ss秒") '時間
     ElseIf ThisWorkbook.Sheets("品証検証").Cells(i + 2, 2) = x Then
       'MsgBox x & y & strDate & "3"
      ThisWorkbook.Sheets("品証検証").Cells(i + 2, 3) = UF校正.SokuteiH '値
      ThisWorkbook.Sheets("品証検証").Cells(i + 2, 9) = line 'ライン
      ThisWorkbook.Sheets("品証検証").Cells(i + 2, 10) = UF校正.Kname '検査員
      ThisWorkbook.Sheets("品証検証").Cells(i + 2, 11) = Format$(Now, "mm/dd h時mm分ss秒") '時間
     End If
   End If
  Next i
 End If
   UF校正.Kname = ""
   UF校正.Saiken.Visible = False
   Unload Me
Application.ScreenUpdating = True
UF校正.Show vbModal '1000mm測定後　通常の校正ﾌｫｰﾑを起動させる
End Sub
Private Sub CommandButton3_Click()
Unload Me
End Sub
Private Sub CommandButton4_Click()
Call 品証ランク付け
End Sub

Private Sub CommandButton5_Click()
With UF校正
  .Saiken.Visible = False
  .NGvalue = ""
  .INP1000 = ""
  .SokuteiH = ""
  Select Case UF校正.LineName
  Case "L-1":    .Gauge = "[L-1]:直尺「S-2」の全長を測定してください"
  Case "LVC":    .Gauge = "[LVC]:直尺「S-1」の全長を測定してください"
  Case "HVC":    .Gauge = "[HVC]:直尺「S-1」の全長を測定してください"
  Case "ﾊﾞﾗﾝｻｰ": .Gauge = "[大板]:直尺「S-1」の全長を測定してください"
  Case "ﾄｯﾄ":    .Gauge = "[大板]:直尺「S-1」の全長を測定してください"
  Case "Test":   .Gauge = "[TEST]:直尺「S-1」の全長を測定してください"
  Case Else: '丸徳用
  End Select
  .Rank1 = ""
  .Rank2 = ""
  .Rank3 = ""
  .Rank4 = ""
  .Rank5 = ""
  .Rank6 = ""
'     Call 校正ランク付け
  .Rank1.BackColor = &HFFFFFF '白
  .Rank2.BackColor = &HFFFFFF
  .Rank3.BackColor = &HFFFFFF
  .Rank4.BackColor = &HFFFFFF
  .Rank5.BackColor = &HFFFFFF
  .Rank6.BackColor = &HFFFFFF
  .MaxLab.BackColor = &HFFFFFF
  .MinLab.BackColor = &HFFFFFF
  .TukiLabe.BackColor = &HFFFFFF
  .BackColor = &HFFFFFF
  .Label3.BackColor = &HFFFFFF
  .Gauge.BackColor = &HFFFFFF
  .NGvalue.BackColor = &HFFFFFF
  .Saiken.BackColor = &HFFFFFF
  .SokuteiH.BackColor = &HFFFFFF
  .INPG.SetFocus
End With
End Sub
Private Sub CommandButton6_Click()
Call 校正ランク付け
End Sub

'Private Sub INP1000_AfterUpdate()
'INP1000 = TerConv(INP1000.Value) '変換
'UF校正.SokuteiH = UF校正.INP1000
'If UF校正.BackColor = &HFFFFFF Then '白
' UF校正.INP1000 = ""
' UF校正.INP1000.SetFocus
' UF校正.INPG.SetFocus 'インプットボックスに初めから数字が入るようにする
'Exit Sub
'End If
'If Len(UF校正.INP1000) <> 6 Then   'xxxx.x  'xxx.x <> 6
'  If Len(UF校正.INP1000) = 5 And Mid(UF校正.INP1000, 1, 1) = "9" Then
'   UF校正.SokuteiH = UF校正.INP1000
'   Call 品証測定
'  Else
'    UF校正.INP1000 = ""
'    Call 音楽不合格
'    Exit Sub
'  End If
'Exit Sub
'Else
' UF校正.SokuteiH = UF校正.INP1000
' Call 品証測定
'End If
'End Sub
'Private Sub INPG_AfterUpdate()
'INPG = TerConv(INPG.Value) '変換
'If UF校正.BackColor = &H80FF80 Then '緑
' UF校正.INPG = ""
' UF校正.INPG.SetFocus
' UF校正.INP1000.SetFocus 'インプットボックスに初めから数字が入るようにする
'Exit Sub
'End If
'If Not IsNumeric(UF校正.INPG.Text) Then UF校正.INPG.Text = "" '数字以外非表示
'If Len(UF校正.INPG) <> 6 Then 'xxxx.x　<>6
'  UF校正.INPG = ""
'  Call 音楽不合格
'  Exit Sub
'Else
'  UF校正.SokuteiH = UF校正.INPG
'  If UF校正.F1 = "1" Then Call 測定完了
'End If
'End Sub

Private Sub INP1000_AfterUpdate()
If UF校正.F1 <> "1" Then Exit Sub
UF校正.F1 = 0
UF校正.INP1000 = TerConv(UF校正.INP1000.Value) '変換
UF校正.SokuteiH = UF校正.INP1000
If UF校正.SokuteiH <> "" Then
  Dim path As String: path = HistoryPath
  Dim Temp As String: Temp = "1000mm測定:" & UF校正.SokuteiH
  Select Case UF校正.LineName
  Case "L-1":    Call WriteLog("*L-1_" & Temp, "", path & "L-1\L-1_Log")
  Case "LVC":    Call WriteLog("*LVC_" & Temp, "", path & "LVC\LVC_Log")
  Case "HVC":    Call WriteLog("*HVC_" & Temp, "", path & "HVC\HVC_Log")
  Case "ﾊﾞﾗﾝｻｰ": Call WriteLog("*バランサー_" & Temp, "", path & "バランサー\バランサー_Log")
  Case "ﾄｯﾄ":    Call WriteLog("*トット_" & Temp, "", path & "トット\トット_Log")
  Case "Test":   Call WriteLog("*Test_" & Temp, "", path & "Test\Test_Log")
  Case Else:     Call WriteLog("*丸徳_" & Temp, "", path & "丸徳\丸徳_Log")
  End Select
  Call 品証測定  '変換後空白でない時のみ測定判定_2021.01.26ログ出力より後にしないとクリアされてログ出力されない
End If
UF校正.F1 = 1 'フラグをもどす
End Sub
Private Sub INP1000_Enter()
If UF校正.BackColor = &HFFFFFF Then '白
  UF校正.INP1000 = ""
  UF校正.INP1000.SetFocus
  UF校正.INPG.SetFocus 'インプットボックスに初めから数字が入るようにする
  Exit Sub
End If
End Sub

Private Sub INPG_AfterUpdate()
If UF校正.F1 <> "1" Then Exit Sub
UF校正.F1 = 0
UF校正.INPG = TerConv(UF校正.INPG.Value) '変換
UF校正.SokuteiH = UF校正.INPG
If UF校正.SokuteiH <> "" Then
  Call 測定完了 '変換後空白でない時のみ測定判定
  Dim path As String: path = HistoryPath
  Dim Temp As String: Temp = "ｹﾞｰｼﾞ測定:" & UF校正.SokuteiH
  Select Case UF校正.LineName
  Case "L-1":    Call WriteLog("*L-1_" & Temp, "", path & "L-1\L-1_Log")
  Case "LVC":    Call WriteLog("*LVC_" & Temp, "", path & "LVC\LVC_Log")
  Case "HVC":    Call WriteLog("*HVC_" & Temp, "", path & "HVC\HVC_Log")
  Case "ﾊﾞﾗﾝｻｰ": Call WriteLog("*バランサー_" & Temp, "", path & "バランサー\バランサー_Log")
  Case "ﾄｯﾄ":    Call WriteLog("*トット_" & Temp, "", path & "トット\トット_Log")
  Case "Test":   Call WriteLog("*Test_" & Temp, "", path & "Test\Test_Log")
  Case Else:     Call WriteLog("*丸徳_" & Temp, "", path & "丸徳\丸徳_Log")
  End Select
Else
  UF校正.F1 = 1 'フラグをもどす
End If
End Sub
Private Sub INPG_Enter()
If UF校正.BackColor = &H80FF80 Then '緑
 UF校正.INPG = ""
 UF校正.INPG.SetFocus
 UF校正.INP1000.SetFocus 'インプットボックスに初めから数字が入るようにする
Exit Sub
End If
End Sub

Private Sub INPG_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If UF校正.F1 = "end" Then '測定完了ﾌﾗｸﾞがあるときのみ閉じる
  DoEvents
  
  UF1.Show vbModeless
  
  Sleep 5
  
  '校正起動前は停止させているので再検査
  UF1.Load.Caption = "off"
  Call LoopJudg
  Call HizLoopJudg
  UF1.Load.Caption = "on"
  Unload UF校正
End If
End Sub

Private Sub Ltrue_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Application.Visible = True 'Excel表示

Unload Me
End Sub

Private Sub UserForm_Layout()
On Error Resume Next
' Excelウィンドウ位置を同期させる
'With Application
'  If .WindowState = xlNormal Then
'    .Top = Me.Top + 1
'    .Left = Me.Left + 1
'  End If
'End With
End Sub
Private Sub UserForm_Activate()
Application.Visible = False 'Excel非表示
Call WindowMenuButtonDel(UF校正)
'SendKeys "{TAB}"
'Call Numlock_onoff
On Error Resume Next
Call Any_Click
' フォームの後ろにExcelウィンドウを隠す
'With Application
'  .Top = Me.Top + 1
'  .Left = Me.Left + 1
'  .Width = Me.Width - 2
'  .Height = Me.Height - 2
'End With

Dim path As String: path = HistoryPath
Dim Temp As String: Temp = "UF校正起動"
Select Case UF校正.LineName
Case "L-1":    Call WriteLog("*L-1_" & Temp, "", path & "L-1\L-1_Log")
Case "LVC":    Call WriteLog("*LVC_" & Temp, "", path & "LVC\LVC_Log")
Case "HVC":    Call WriteLog("*HVC_" & Temp, "", path & "HVC\HVC_Log")
Case "ﾊﾞﾗﾝｻｰ": Call WriteLog("*バランサー_" & Temp, "", path & "バランサー\バランサー_Log")
Case "ﾄｯﾄ":    Call WriteLog("*トット_" & Temp, "", path & "トット\トット_Log")
Case "Test":   Call WriteLog("*Test_" & Temp, "", path & "Test\Test_Log")
Case Else:     Call WriteLog("*丸徳_" & Temp, "", path & "丸徳\丸徳_Log")
End Select
End Sub

Private Sub UserForm_Initialize()
On Error Resume Next
' 現在のExcelウィンドウ状態を保管
'With Application
'  .WindowState = xlNormal
'  g_sngTopSV = .Top
'  g_sngLeftSV = .Left
'  g_sngWidthSV = .Width
'  g_sngHeightSV = .Height
'End With
' Excelウィンドウを最小化
'Application.WindowState = xlMinimized
'AppActivate Application.Caption

With UF校正
  .Rank1.Visible = True
  .Rank2.Visible = True
  .Rank3.Visible = True
  .Rank4.Visible = True
  .Rank5.Visible = True
  .Rank6.Visible = True
  .MaxLab.Visible = True
  .MinLab.Visible = True
  .TukiLabe.Visible = True
  .TukiLabe = Format$(Now, "m") & "月度"
  .Rank1.BackColor = &H80FF80 '緑
  .Rank2.BackColor = &H80FF80
  .Rank3.BackColor = &H80FF80
  .Rank4.BackColor = &H80FF80
  .Rank5.BackColor = &H80FF80
  .Rank6.BackColor = &H80FF80
  .MaxLab.BackColor = &H80FF80
  .MinLab.BackColor = &H80FF80
  .TukiLabe.BackColor = &H80FF80
  .BackColor = &H80FF80
  .Label3.BackColor = &H80FF80
  .Gauge.BackColor = &H80FF80
  .Saiken.BackColor = &H80FF80
  .NGvalue.BackColor = &H80FF80
  .SokuteiH.BackColor = &H80FF80
  .Ltrue.BackColor = &H80FF80
  .Saiken.Visible = False '再検査の非表示
  .NGvalue.Visible = False 'NG値の非表示
  .NGvalue = ""
  .F1 = 1 'フラグを１にする
  .INP1000 = ""
  .INPG.SetFocus '後側　SendKeys "{TAB}"でHへ移動
End With
Call 品証ランク付け
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then
  Cancel = True
End If
End Sub

