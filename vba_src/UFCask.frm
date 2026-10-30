Attribute VB_Name = "UFCask"
Attribute VB_Base = "0{08C13520-58A7-4DC8-B4DD-E440824B13DC}{4295D2DF-C13D-4A08-A59F-46FBCAA99D2A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False







Option Explicit
Private g_lngWindowState As XlWindowState                       ' ウィンドウ状態
Private Sub ComboBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'入力制限
Select Case KeyAscii.Value
Case 48 To 57 '0~9
Case 65 To 90 'A~Z
Case 97 To 122 'a~z
KeyAscii.Value = KeyAscii.Value - 32 '大文字強制変換
Case Else
KeyAscii.Value = 0
End Select
End Sub
Private Sub MAXTex_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'入力制限
Select Case KeyAscii.Value
Case 48 To 57 '0~9
Case 65 To 90 'A~Z
Case 97 To 122 'a~z
KeyAscii.Value = KeyAscii.Value - 32 '大文字強制変換
Case Else
KeyAscii.Value = 0
End Select
End Sub

Private Sub CommandButton1_Click()
If UFCask.Tag = "off" Then Exit Sub
If Register_Check(Format$(Now, "hh:mm:ss")) Then Call Daily_Create '日報作成
If MsgBox("新規発行しますか？現在のｼｰﾄは削除されます。", vbQuestion + vbYesNo) = vbYes Then
  UFProgress.Show vbModeless
  UFProgress.CommandButton1.Caption = "NewDateCreate...."
  Call ProgressBar_count(0, 230)
  DoEvents
  Call JustinCase_KySave
  Call ProgressBar_count(230, 460)
  DoEvents
  Call キャスクコピー
  Select Case UFCask.LineName.Caption
  Case "L-1":    Call WriteLog("*UFKen:" & "ｷｬｽｸ検査ｼｰﾄ発行", "", HistoryPath & "L-1\L-1_Log")
  Case "LVC":    Call WriteLog("*UFKen:" & "ｷｬｽｸ検査ｼｰﾄ発行", "", HistoryPath & "LVC\LVC_Log")
  Case "HVC":    Call WriteLog("*UFKen:" & "ｷｬｽｸ検査ｼｰﾄ発行", "", HistoryPath & "HVC\HVC_Log")
  Case "ﾊﾞﾗﾝｻｰ": Call WriteLog("*UFKen:" & "ｷｬｽｸ検査ｼｰﾄ発行", "", HistoryPath & "バランサー\バランサー_Log")
  Case "ﾄｯﾄ":    Call WriteLog("*UFKen:" & "ｷｬｽｸ検査ｼｰﾄ発行", "", HistoryPath & "トット\トット_Log")
  Case "Test":   Call WriteLog("*UFKen:" & "ｷｬｽｸ検査ｼｰﾄ発行", "", HistoryPath & "Test\Test_Log")
  Case Else:     Call WriteLog("*UFKen:" & "ｷｬｽｸ検査ｼｰﾄ発行", "", HistoryPath & "丸徳\丸徳_Log")
  End Select
  'ｸﾘｱ
  UFCask.Inspection = "OFF"
  UFCask.LotBox = ""
  UFCask.KenNo = ""
  UFCask.Kname = ""
  UFCask.ComboBox1 = ""
  UFCask.Convex = ""
  UFCask.Micro = ""
  
  Call ProgressBar_count(460, 1000)
  DoEvents
  Unload UFProgress
Else
  Unload UFProgress
  MsgBox "ｷｬﾝｾﾙ処理", vbInformation
End If
Application.EnableEvents = True
End Sub

Private Sub CommandButton11_Click()
Call JustinCase_KySave
End Sub

Private Sub CommandButton12_Click()
If UFCask.Read = "ON" Then
  UFCask.Read = "OFF"
ElseIf UFCask.Read = "OFF" Then
  UFCask.Read = "ON"
End If
Application.EnableEvents = True
End Sub

Private Sub CommandButton14_Click()
  Cells(7, 1).Speak '枚数読み上げ
  Cells(8, 1).Speak '位置読み上げ
End Sub

Private Sub CommandButton15_Click()
Call ProofreadingSearch(YeaDayrCheck(Format$(Now, "hh:mm")), TimeCheck(Format$(Now, "hh:mm:ss")))
End Sub

Private Sub CommandButton16_Click()
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("キャスク検査")
ws.Unprotect
ws.Range("AB:AB").EntireColumn.AutoFit
'ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'シート保護
End Sub

Private Sub CommandButton2_Click()
Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("キャスク検査")
Dim n As Long
UFCask.Tag = "off"

ws.Activate
Call ループ判定

If UFCask.Cno = "" Or UFCask.Mno = "" Or UFCask.Kname = "" Or UFCask.SeriNo = "" Then
    MsgBox "フォーム上で必要欄を記入してください", vbExclamation
    Exit Sub
Else
  UFProgress.Show vbModeless
  Call ProgressBar_count(0, 330)
  DoEvents
  
    Application.ScreenUpdating = False
    '2026.02.20_梱包板Noの保存追加
    Dim strLot As String: strLot = Me.LotBox
    Dim strKen As String: strKen = Me.KenNo
    Dim endflag As Boolean: endflag = CheckCaskEndNo(strLot, strKen) 'エンドチェック
    If endflag Then
        Call キャスク印刷範囲
    Else
        Exit Sub
    End If
    Application.ScreenUpdating = True
  
  Call MusicTrue("印刷.wav")
  If MsgBox("キャスク検査を印刷しますか？", vbQuestion + vbYesNo) = vbYes Then
    Call ProgressBar_count(330, 765)
    DoEvents
    ws.Visible = True
    ws.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
    Call キャスク専用保存
    Call ProgressBar_count(765, 888)
    DoEvents
  '----------------------------------------------履歴を残す　2019.09.13　これ以前のものは履歴なし
    If WsCheck("保存用") Then
      With ThisWorkbook.Sheets("保存用")
        n = .Cells(Rows.Count, 1).End(xlUp).row + 1
        .Unprotect
        .Cells.Locked = False
        .Cells.FormulaHidden = False
        .Cells(n, 2) = ws.Cells(2, 4) 'ロット
        .Cells(n, 83) = "ｷｬｽｸ"
        .Cells(n, 4) = ws.Cells(3, 2) '厚
        .Cells(n, 4) = ws.Cells(3, 4) '幅
        .Cells(n, 5) = ws.Cells(3, 6) '丈
        .Cells(n, 1) = Format$(ws.Cells(1, 2), "yyyy/mm/dd  h:mm") & "～" & Format$(Now, "yyyy/mm/dd  h:mm")
        .Cells(n, 77) = "別ファイル参照" '検査員
      End With
    End If
    Call ProgressBar_count(888, 1000)
    Call MusicTrue("完了.wav")
    DoEvents
    Unload UFProgress
  Else
    Unload UFProgress
    MsgBox "ｷｬﾝｾﾙ処理", vbInformation
  End If
End If
Application.WindowState = xlMaximized
UFCask.Tag = ""
End Sub

Private Sub CommandButton25_Click()
Me.Width = 582.75
Me.Height = 339.75
Me.Zoom = 90
End Sub

Private Sub CommandButton3_Click()
If UFCask.Tag = "off" Then Exit Sub
Application.EnableEvents = False
UFProgress.Show vbModeless
UFProgress.CommandButton1.Caption = "転送中...."

'2026.02.20_梱包板Noの保存追加
Dim dtWork As Date
dtWork = Now()   'クリック時刻を取得

'入力チェック
If Not ValidateCaskInput() Then
    Unload UFProgress
    Exit Sub
End If
'連番チェック
If Not CheckCaskSequence() Then
    Unload UFProgress
    Exit Sub
End If

Dim strLot As String: strLot = Me.LotBox
Dim strKen As String: strKen = Me.KenNo

If ThisWorkbook.Sheets("キャスク検査").Cells(3, 2) <> "" Then
    Select Case True
    Case 100# > val(ThisWorkbook.Sheets("キャスク検査").Cells(3, 2)):  Call CountSelect(3)
    Case 100# <= val(ThisWorkbook.Sheets("キャスク検査").Cells(3, 2)): Call CountSelect(4)
    End Select
End If

'DB保存
If SaveCaskToDB(dtWork) Then
    Call AddCaskLabels(Me, strLot, strKen, 18, 300, 90, 20, 6) 'ラベル生成
'    MsgBox "DB保存が完了しました。" & vbCrLf & _
'           "作業時間：" & Format(dtWork, "yyyy/mm/dd hh:mm:ss"), vbInformation

    Call ProgressBar_count(1, 730)
    Call キャスク用最大最小
    Call ProgressBar_count(730, 1000)
    Unload UFProgress
    Call Ainput_Startup("Cask", "L2") '起動ﾌｫｰﾑ、検査法
End If

'Call ProgressBar_count(1, 730)
'Call キャスク用最大最小
'Call ProgressBar_count(730, 1000)
'Unload UFProgress
'Call Ainput_Startup("Cask", "L2") '起動ﾌｫｰﾑ、検査法
Application.EnableEvents = True
End Sub

Private Sub CommandButton4_Click()
UFCask.Tag = "off"
If UFCask.Cno = "" Or UFCask.Mno = "" Or UFCask.Kname = "" Or UFCask.SeriNo = "" Then
 MsgBox "フォーム上で必要欄を記入してください"
 Exit Sub
Else
  Call キャスク専用保存
End If
UFCask.Tag = ""
End Sub

Private Sub CommandButton5_Click()
Dim SN As String, Color As Long: Color = 2 '白
Dim ws As Worksheet, flag As Boolean
Dim AR As Long: AR = ActiveCell.row
Dim ac As Long: ac = ActiveCell.Column
Dim i As Long
Application.ScreenUpdating = False
Application.EnableEvents = False
SN = "キャスク検査"
For Each ws In Worksheets
  If ws.Name = SN Then flag = True
Next ws
ThisWorkbook.Sheets("キャスク検査").Unprotect '保護解除
'----------------------------------------------------
UFCask.Hflag = "off" '判定を止める
'----------------------------------------------------.Interior.Color = RGB(153, 204, 255)
If flag = True Then
  ThisWorkbook.Sheets("キャスク検査").Activate
  Select Case ac
  Case 4, 6, 13, 15, 17, 19, 22
   Range(Cells(AR, ac - 2), Cells(AR, ac - 1)).ClearContents
   Range(Cells(AR, ac - 2), Cells(AR, ac)).Interior.ColorIndex = Color
   Cells(AR, ac - 2).Activate
   Cells(AR, ac - 2).Interior.Color = RGB(153, 204, 255)
   AppActivate Application.Caption
  Case 11
   For i = 6 To 7
    Cells(AR, i).ClearContents
   Next i
   Range(Cells(AR, 7), Cells(AR, ac)).Interior.ColorIndex = Color
   Cells(AR, 6).Activate
   Cells(AR, 6).Interior.Color = RGB(153, 204, 255)
   AppActivate Application.Caption
  Case 20
   Cells(AR, 19).ClearContents
   Range(Cells(AR, 19), Cells(AR, ac)).Interior.ColorIndex = Color
   Cells(AR, 19).Activate
   Cells(AR, 19).Interior.Color = RGB(153, 204, 255)
   AppActivate Application.Caption
  Case 24
   For i = 19 To 23
    Cells(AR, i).ClearContents
   Next i
   Range(Cells(AR, 19), Cells(AR, ac)).Interior.ColorIndex = Color
   Cells(AR, 19).Activate
   Cells(AR, 19).Interior.Color = RGB(153, 204, 255)
   AppActivate Application.Caption
  Case 25 To 28
   Cells(AR, ac).Interior.ColorIndex = Color
   Cells(AR, ac - 1).ClearContents
   Cells(AR, ac - 1).Activate
   Cells(AR, ac - 1).Interior.Color = RGB(153, 204, 255)
   AppActivate Application.Caption
  End Select
End If
'----------------------------------------------------
UFCask.Hflag = "on" '判定再開
'----------------------------------------------------
ThisWorkbook.Sheets("キャスク検査").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
Application.ScreenUpdating = True
Application.EnableEvents = True
End Sub

Private Sub CommandButton6_Click()
Dim strLot As String: strLot = Me.LotBox
Dim strKen As String: strKen = Me.KenNo
'Dim temp As Variant: temp = GetCaskData(strLot, strKen)
''
''Stop
'
'Call AddCaskLabels(Me, strLot, strKen, 18, 312, 115, 20, 6)

Call CountSelect(3)
Dim endflag As Boolean: endflag = CheckCaskEndNo(strLot, strKen) 'エンドチェック
If endflag Then Call キャスク印刷範囲

Call キャスク専用保存

End Sub
Private Sub CommandButton8_Click()
Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton9_Click()
Dim ws As Worksheet, FL As Boolean
Dim sh As String
Application.ScreenUpdating = False
sh = "キャスク検査"
For Each ws In ThisWorkbook.Worksheets
    If ws.Name = sh Then FL = True
Next ws
If FL = False Then
Exit Sub
Else
  If Me.Inspection = "OFF" Then
    Me.Inspection = "ON"
    ThisWorkbook.Sheets("キャスク検査").Unprotect '保護解除
    ThisWorkbook.Sheets("キャスク検査").Range("K7:AB706").Locked = False 'ロック設定
    ThisWorkbook.Sheets("キャスク検査").Range("K6:AB706").Interior.Color = RGB(255, 255, 255)
    ThisWorkbook.Sheets("キャスク検査").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
    Call ループ判定
  ElseIf Me.Inspection = "ON" Then
    Me.Inspection = "OFF"
    ThisWorkbook.Sheets("キャスク検査").Unprotect '保護解除
    ThisWorkbook.Sheets("キャスク検査").Range("K6:AB706").Interior.Color = RGB(128, 128, 128)
    ThisWorkbook.Sheets("キャスク検査").Range("K7:AB706").Locked = True 'ロック設定
    ThisWorkbook.Sheets("キャスク検査").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
  End If
End If
MoveTg("A", "キャスク検査").Activate
AppActivate Application.Caption
Application.ScreenUpdating = True
Application.EnableEvents = True
End Sub

Private Sub LotBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'入力制限
Select Case KeyAscii.Value
Case 48 To 57 '0~9
Case 65 To 90 'A~Z
Case 97 To 122 'a~z
KeyAscii.Value = KeyAscii.Value - 32 '大文字強制変換
Case Else
KeyAscii.Value = 0
End Select

End Sub
Private Sub LotBox_Change()
If Not ThisWorkbook.Sheets("キャスク検査").Visible = True Then Exit Sub
Application.ScreenUpdating = False
Application.EnableEvents = False
If Len(Me.LotBox.Text) <> 7 Then
Exit Sub
Else
  Me.ComboBox1 = Me.LotBox
  MoveTg("A", "キャスク検査").Activate
  AppActivate Application.Caption
End If
Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub
Private Sub MINText_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'入力制限
Select Case KeyAscii.Value
Case 48 To 57 '0~9
Case 65 To 90 'A~Z
Case 97 To 122 'a~z
KeyAscii.Value = KeyAscii.Value - 32 '大文字強制変換
Case Else
KeyAscii.Value = 0
End Select
End Sub

'2026.05.29_読み方変更　ADO 化
Private Sub ComboBox1_Change()
Application.EnableEvents = False
Application.ScreenUpdating = False '処理高速化のための画面更新停止
If WsCheck("キャスク検査") = False Then Exit Sub
If Len(UFCask.ComboBox1.Text) <> 7 Then Exit Sub
If UFCask.Loadtext.Caption = "off" Then Exit Sub
Dim b As String: b = UFCask.ComboBox1.ListIndex
Dim i As Long, a As Long, z As Long, k As Long, j As Long
Dim Order As Long, Lot As Long, ATU As Long, HABA As Long, TAKE As Long, KenNo As Long, HikiNo As Long
Dim Zai As Long, Tyou As Long, UseNo As Long, cours As Long, Good As Long, Deliver As Long, Sendi As Long
Dim 製厚P As Long, 製厚M As Long, O厚P As Long, O厚M As Long
Dim 製巾P As Long, 製巾M As Long, O巾P As Long, O巾M As Long
Dim 製丈P As Long, 製丈M As Long, O丈P As Long, O丈M As Long, 対角 As Long
Dim 実績枚 As Long, 実績重 As Long, 合紙 As Long, VC As Long

  Dim db       As Object
  Dim rs       As Object
  Dim strPath  As String
  Dim strSQL   As String
  Dim NoFl     As Boolean: NoFl = False
  Dim PastHiki(0 To 17)

  On Error GoTo ErrHandler_ComboChange

  strPath = MainPath & "SIKALOTNOW.accdb"
  strSQL = "SELECT * FROM 仕掛 WHERE ﾛｯﾄ番号 = '" & Me.ComboBox1.Text & "' AND ｵｰﾀﾞｰ板丈 <> '0'"
  Set db = CreateObject("ADODB.Connection")
  db.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strPath & ";"
  Set rs = CreateObject("ADODB.Recordset")
  rs.Open strSQL, db, 3, 1
  
  If Not rs.EOF Then
    NoFl = True
    With ThisWorkbook.Sheets("キャスク検査")
      .Unprotect
      PastHiki(0) = "'" & rs.Fields("オーダー番号").Value
      PastHiki(1) = rs.Fields("ロット番号").Value
      PastHiki(2) = rs.Fields("製造板厚").Value
      PastHiki(3) = rs.Fields("製造板幅").Value
      PastHiki(4) = rs.Fields("製造板丈").Value
      PastHiki(5) = rs.Fields("検査番号").Value
      PastHiki(6) = rs.Fields("製造材質").Value
      PastHiki(7) = rs.Fields("製造調質").Value
      PastHiki(8) = rs.Fields("用途コード").Value
      PastHiki(9) = rs.Fields("設計_設備コース").Value
      PastHiki(10) = rs.Fields("取引先名").Value
      PastHiki(11) = rs.Fields("納入先名").Value
      PastHiki(12) = rs.Fields("送り先名").Value
      PastHiki(13) = rs.Fields("BOX実績_枚本数").Value
      PastHiki(14) = rs.Fields("BOX実績_良品重量").Value
      PastHiki(15) = rs.Fields("梱包_合紙").Value
      PastHiki(16) = rs.Fields("VC").Value

      .Cells(2, 4) = Me.ComboBox1.Text
      UFCask.KenNo = rs.Fields("検査番号").Value

      .Cells(3, 2) = Format(val(rs.Fields("製造板厚").Value), "0.000")
      Select Case True
        Case 100# > .Cells(3, 2): Call CountSelect(3)
        Case 100# <= .Cells(3, 2): Call CountSelect(4)
      End Select
      .Cells(3, 4) = Format(val(rs.Fields("製造板幅").Value), "0.0")
      .Cells(3, 6) = Format(val(rs.Fields("製造板丈").Value), "0.0")
      
      ' 板厚公差
      Dim dOP As Double: dOP = val(rs.Fields("板厚公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ").Value)
      Dim dOM As Double: dOM = val(rs.Fields("板厚公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ").Value)
      Dim dMP As Double: dMP = val(rs.Fields("板厚公差_製造_ﾌﾟﾗｽ").Value)
      Dim dMM As Double: dMM = val(rs.Fields("板厚公差_製造_ﾏｲﾅｽ").Value)
      If dOP = 0 Or dOM = 0 Then dOP = dMP: dOM = dMM
      .Cells(2, 15) = Format(dOP, "0.000")
      .Cells(2, 13) = Format(dOM, "0.000")
      .Cells(2, 19) = "=R[1]C[-17]+RC[-4]"
      .Cells(2, 17) = "=R[1]C[-15]-RC[-4]"

      ' 板幅公差
      Dim wOP As Double: wOP = val(rs.Fields("板幅公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ").Value)
      Dim wOM As Double: wOM = val(rs.Fields("板幅公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ").Value)
      Dim wMP As Double: wMP = val(rs.Fields("板幅公差_製造_ﾌﾟﾗｽ").Value)
      Dim wMM As Double: wMM = val(rs.Fields("板幅公差_製造_ﾏｲﾅｽ").Value)
      If wOP = 0 Or wOM = 0 Then wOP = wMP: wOM = wMM
      .Cells(3, 15) = Format(wOP, "0.0")
      .Cells(3, 13) = Format(wOM, "0.0")
      .Cells(3, 19) = "=RC[-15]+RC[-4]"
      .Cells(3, 17) = "=RC[-13]-RC[-4]"

      ' 板丈公差
      Dim tOP As Double: tOP = val(rs.Fields("板丈公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ").Value)
      Dim tOM As Double: tOM = val(rs.Fields("板丈公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ").Value)
      Dim tMP As Double: tMP = val(rs.Fields("板丈公差_製造_ﾌﾟﾗｽ").Value)
      Dim tMM As Double: tMM = val(rs.Fields("板丈公差_製造_ﾏｲﾅｽ").Value)
      If tOP = 0 Or tOM = 0 Then tOP = tMP: tOM = tMM
      .Cells(4, 15) = Format(tOP, "0.0")
      .Cells(4, 13) = Format(tOM, "0.0")
      .Cells(4, 19) = "=R[-1]C[-13]+RC[-4]"
      .Cells(4, 17) = "=R[-1]C[-11]-RC[-4]"
      .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True
    End With
  End If

  rs.Close: Set rs = Nothing
  db.Close: Set db = Nothing

'*************************************************************
  With ThisWorkbook.Sheets("キャスク検査")
      If NoFl = False Then
        .Unprotect
        .Cells(3, 2) = ""
        .Cells(3, 4) = ""
        .Cells(3, 6) = ""
        .Cells(2, 13) = ""
        .Cells(2, 15) = ""
        .Cells(2, 17) = ""
        .Cells(2, 19) = ""
        .Cells(3, 13) = ""
        .Cells(3, 15) = ""
        .Cells(3, 17) = ""
        .Cells(3, 19) = ""
        .Cells(4, 13) = ""
        .Cells(4, 15) = ""
        .Cells(4, 17) = ""
        .Cells(4, 19) = ""
        UFCask.KenNo = ""
        If UFCask.ComboBox1.Tag <> "Load" Then MsgBox "Text.データなし　手入力よろしく"
        Range("Q2").FormulaR1C1 = "=R[1]C[-15]-RC[-4]" '厚
        Range("S2").FormulaR1C1 = "=R[1]C[-17]+RC[-4]"
        Range("Q3").FormulaR1C1 = "=RC[-13]-RC[-4]" '巾
        Range("S3").FormulaR1C1 = "=RC[-15]+RC[-4]"
        Range("Q4").FormulaR1C1 = "=R[-1]C[-11]-RC[-4]" '丈
        Range("S4").FormulaR1C1 = "=R[-1]C[-13]+RC[-4]"
        .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True
      Exit Sub
      End If
  End With
Call HistorySave(PastHiki)
ThisWorkbook.Sheets("キャスク検査").Activate

If Not PackingNoCheck(Me.LotBox) = "" Then Call UFPack_Open(PackingNoCheck(Me.LotBox), Me.LotBox) '包装仕様書Noチェック

  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Exit Sub

ErrHandler_ComboChange:
  If Not rs Is Nothing Then
    If rs.state = 1 Then rs.Close
    Set rs = Nothing
  End If
  If Not db Is Nothing Then
    db.Close: Set db = Nothing
  End If
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  MsgBox "DB読み込みエラー:" & vbCrLf & Err.Description, vbCritical
End Sub

'
'Private Sub ComboBox1_Change()
'Application.EnableEvents = False
'Application.ScreenUpdating = False '処理高速化のための画面更新停止
'If WsCheck("キャスク検査") = False Then Exit Sub
'If Len(UFCask.ComboBox1.Text) <> 7 Then Exit Sub
'If UFCask.Loadtext.Caption = "off" Then Exit Sub
'Dim b As String: b = UFCask.ComboBox1.ListIndex
'Dim i As Long, a As Long, z As Long, k As Long, j As Long
'Dim Order As Long, LOT As Long, Atu As Long, Haba As Long, Take As Long, KENNO As Long, HikiNo As Long
'Dim Zai As Long, Tyou As Long, UseNo As Long, Cours As Long, Good As Long, Deliver As Long, Sendi As Long
'Dim 製厚P As Long, 製厚M As Long, O厚P As Long, O厚M As Long
'Dim 製巾P As Long, 製巾M As Long, O巾P As Long, O巾M As Long
'Dim 製丈P As Long, 製丈M As Long, O丈P As Long, O丈M As Long, 対角 As Long
'Dim 実績枚 As Long, 実績重 As Long, 合紙 As Long, VC As Long
'
'Dim HiAcc: HiAcc = Arrbox("SIKALOTNOW.accdb", "板", "Fields")
'  For i = 1 To UBound(HiAcc, 2) '製造'ｵｰﾀﾞｰ公差確認後片方でも0の場合は製造公差を見ることとする
'    Select Case HiAcc(1, i)
'    Case "ｵｰﾀﾞｰ番号":    Order = i
'    Case "ﾛｯﾄ番号":      LOT = i
'    Case "製造板厚":     Atu = i
'    Case "製造板幅":     Haba = i
'    Case "製造板丈":     Take = i
'    Case "検査番号":     KENNO = i
'    Case "製造材質":     Zai = i
'    Case "製造調質":     Tyou = i
'    Case "用途ｺｰﾄﾞ":     UseNo = i
'    Case "設計_設備ｺｰｽ": Cours = i
'    Case "取引先名":     Good = i
'    Case "納入先名":     Deliver = i
'    Case "送り先名":     Sendi = i
'    Case "板厚公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ": O厚P = i
'    Case "板厚公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ": O厚M = i
'    Case "板厚公差_製造_ﾌﾟﾗｽ":  製厚P = i
'    Case "板厚公差_製造_ﾏｲﾅｽ":  製厚M = i
'    Case "板幅公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ": O巾P = i
'    Case "板幅公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ": O巾M = i
'    Case "板幅公差_製造_ﾌﾟﾗｽ":  製巾P = i
'    Case "板幅公差_製造_ﾏｲﾅｽ":  製巾M = i
'    Case "板丈公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ": O丈P = i
'    Case "板丈公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ": O丈M = i
'    Case "板丈公差_製造_ﾌﾟﾗｽ":  製丈P = i
'    Case "板丈公差_製造_ﾏｲﾅｽ":  製丈M = i
'    Case "品質ｸﾞﾚｰﾄﾞ_直角度":   対角 = i
'    Case "BOX実績_枚本数":      実績枚 = i
'    Case "BOX実績_良品重量":    実績重 = i
'    Case "梱包_合紙":           合紙 = i
'    Case "VC":                  VC = i
'    End Select
'  Next i
'  Dim PastHiki(0 To 17)
'  Dim NoFl As Boolean: NoFl = False
'  Dim LotAcc: LotAcc = Arrbox("SIKALOTNOW.accdb", "板", "Record")
'
'  For i = 1 To UBound(LotAcc, 1)
'  With ThisWorkbook.Sheets("キャスク検査")
'    If LotAcc(i, LOT) = Me.ComboBox1.Text Then 'LOTNO一致がある
'        NoFl = True
'        PastHiki(0) = "'" & LotAcc(i, Order)
'        PastHiki(1) = LotAcc(i, LOT)
'        PastHiki(2) = LotAcc(i, Atu)
'        PastHiki(3) = LotAcc(i, Haba)
'        PastHiki(4) = LotAcc(i, Take)
'        PastHiki(5) = LotAcc(i, KENNO) '検番
'        PastHiki(6) = LotAcc(i, Zai)
'        PastHiki(7) = LotAcc(i, Tyou)
'        PastHiki(8) = LotAcc(i, UseNo) '用途コード
'        PastHiki(9) = LotAcc(i, Cours) 'ｺｰｽ
'        PastHiki(10) = LotAcc(i, Good) '得意先
'        PastHiki(11) = LotAcc(i, Deliver) '納入先
'        PastHiki(12) = LotAcc(i, Sendi) '送り先
'        PastHiki(13) = LotAcc(i, 実績枚) '実績枚
'        PastHiki(14) = LotAcc(i, 実績重) '実績重
'        PastHiki(15) = LotAcc(i, 合紙) '合紙
'        PastHiki(16) = LotAcc(i, VC) 'VC
'        .Cells(2, 4) = Me.ComboBox1.Text
'        UFCask.KENNO = LotAcc(i, KENNO) '検査番号
'        .Cells(3, 2) = Format(LotAcc(i, Atu), "0.000") '厚　製造サイズとする
'        Select Case True
'        Case 100# > .Cells(3, 2):  Call CountSelect(3)
'        Case 100# <= .Cells(3, 2): Call CountSelect(4)
'        End Select
'        .Cells(3, 4) = Format(LotAcc(i, Haba), "0.0")  '巾
'        .Cells(3, 6) = Format(LotAcc(i, Take), "0.0") '丈
'       '*************************************************************厚公差
'        .Cells(2, 15) = Format(LotAcc(i, O厚P), "0.000") '+　　マイナスを左側
'        .Cells(2, 13) = Format(LotAcc(i, O厚M), "0.000") '-
'          If (.Cells(2, 13) = "0" Or .Cells(2, 15) = "0") Or .Cells(2, 13) = "" Or .Cells(2, 15) = "" Then '片方でも0の場合は製造公差を見る
'          .Cells(2, 15) = Format(LotAcc(i, 製厚P), "0.000") '+　　マイナスを左側
'          .Cells(2, 13) = Format(LotAcc(i, 製厚M), "0.000") '-
'          End If
'       '--------------------------厚公差計算後
'        .Cells(2, 19) = "=R[1]C[-17]+RC[-4]" '+' Format(Val(.Cells(3, 2)) + Val(.Cells(2, 15)), "0.000") 'マイナスを左側
'        .Cells(2, 17) = "=R[1]C[-15]-RC[-4]" '-'Format(Val(.Cells(3, 2)) - Val(.Cells(2, 13)), "0.000") '- 厚
'       '*************************************************************巾公差
'         .Cells(3, 15) = Format(LotAcc(i, O巾P), "0.0") '+
'         .Cells(3, 13) = Format(LotAcc(i, O巾M), "0.0") '-
'        If .Cells(3, 13) = "" Or .Cells(3, 13) = "0" Or .Cells(3, 15) = "" Or .Cells(3, 15) = "0" Then '片方でも0の場合は製造公差を見る
'         .Cells(3, 15) = Format(LotAcc(i, 製巾P), "0.0") '+
'         .Cells(3, 13) = Format(LotAcc(i, 製巾M), "0.0") '-
'        End If
'       '--------------------------巾公差計算後
'        .Cells(3, 19) = "=RC[-15]+RC[-4]" '+'Format(Val(.Cells(3, 4)) + Val(.Cells(3, 15)), "0.0") 'マイナスを左側
'        .Cells(3, 17) = "=RC[-13]-RC[-4]" '-'Format(Val(.Cells(3, 4)) - Val(.Cells(3, 13)), "0.0") '- 巾
'       '*************************************************************丈公差
'         .Cells(4, 15) = Format(LotAcc(i, O丈P), "0.0") '+
'         .Cells(4, 13) = Format(LotAcc(i, O丈M), "0.0") '-
'        If .Cells(4, 13) = "" Or .Cells(4, 13) = "0" Or .Cells(4, 15) = "" Or .Cells(4, 15) = "0" Then '片方でも0の場合は製造公差を見る
'         .Cells(4, 15) = Format(LotAcc(i, 製丈P), "0.0") '+
'         .Cells(4, 13) = Format(LotAcc(i, 製丈M), "0.0") '-
'        End If
'       '--------------------------丈公差計算後
'        .Cells(4, 19) = "=R[-1]C[-13]+RC[-4]" '+'Format(Val(.Cells(3, 6)) + Val(.Cells(4, 15)), "0.0") 'マイナスを左側
'        .Cells(4, 17) = "=R[-1]C[-11]-RC[-4]" '-' Format(Val(.Cells(3, 6)) - Val(.Cells(4, 13)), "0.0") '- 丈
'    End If
'  End With
'  Next i
''*************************************************************
'  With ThisWorkbook.Sheets("キャスク検査")
'      If NoFl = False Then
'        .Cells(3, 2) = ""
'        .Cells(3, 4) = ""
'        .Cells(3, 6) = ""
'        .Cells(2, 13) = ""
'        .Cells(2, 15) = ""
'        .Cells(2, 17) = ""
'        .Cells(2, 19) = ""
'        .Cells(3, 13) = ""
'        .Cells(3, 15) = ""
'        .Cells(3, 17) = ""
'        .Cells(3, 19) = ""
'        .Cells(4, 13) = ""
'        .Cells(4, 15) = ""
'        .Cells(4, 17) = ""
'        .Cells(4, 19) = ""
'        UFCask.KENNO = ""
'        If UFCask.ComboBox1.Tag <> "Load" Then MsgBox "Text.データなし　手入力よろしく"
'        Range("Q2").FormulaR1C1 = "=R[1]C[-15]-RC[-4]" '厚
'        Range("S2").FormulaR1C1 = "=R[1]C[-17]+RC[-4]"
'        Range("Q3").FormulaR1C1 = "=RC[-13]-RC[-4]" '巾
'        Range("S3").FormulaR1C1 = "=RC[-15]+RC[-4]"
'        Range("Q4").FormulaR1C1 = "=R[-1]C[-11]-RC[-4]" '丈
'        Range("S4").FormulaR1C1 = "=R[-1]C[-13]+RC[-4]"
'      Exit Sub
'      End If
'  End With
'Call HistorySave(PastHiki)
'ThisWorkbook.Sheets("キャスク検査").Activate
'
'If Not PackingNoCheck(Me.LotBox) = "" Then Call UFPack_Open(PackingNoCheck(Me.LotBox), Me.LotBox) '包装仕様書Noチェック
'
'Application.ScreenUpdating = True '処理高速化のための画面更新停止
'Application.EnableEvents = True
'End Sub
Private Sub Label4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'UFpic.Show vbmoderes
End Sub

Private Sub SeriNo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'入力制限
Select Case KeyAscii.Value
Case 48 To 57 '0~9
Case 65 To 90 'A~Z
Case 97 To 122 'a~z
KeyAscii.Value = KeyAscii.Value - 32 '大文字強制変換
Case Else
KeyAscii.Value = 0
End Select
End Sub
Private Sub SB_SpinDown()
  ActiveWindow.SmallScroll down:=5
  MoveTg("A", "キャスク検査").Activate
End Sub
Private Sub SB_SpinUp()
  ActiveWindow.SmallScroll up:=5
  MoveTg("A", "キャスク検査").Activate
End Sub

Private Sub StartTB_AfterUpdate()
Dim PrNo As String
If UFCask.StartTB = "#L" Then
  UFCask.L2.Visible = True
  Exit Sub
End If

Dim Temp As String: Temp = Maker(UFCask.StartTB)
Dim MicroNo As String, ConvexNo As String
Dim ATU As String: ATU = ThisWorkbook.Sheets("キャスク検査").Cells(3, 2)
Select Case Temp
Case "DT"
  PrNo = Conv(UFCask.StartTB)
  If ATU = "" Then ATU = "0"
  If ATU >= 100 Then
    UFCask.L1.Caption = "ﾉｷﾞｽｾﾞﾛ点"
    UFCask.Micro = Format(val(PrNo), "0.00")
    If NumericConfi("M") = False Then
      MsgBox "ノギス校正を正しく実行してください！！", vbExclamation
      UFCask.Micro.Caption = ""
    End If
  Else
    UFCask.L1.Caption = "ﾏｲｸﾛｾﾞﾛ点"
    UFCask.Micro = Format(val(PrNo), "0.000")
    If NumericConfi("M") = False Then
      MsgBox "マイクロ校正を正しく実行してください！！", vbExclamation
      UFCask.Micro.Caption = ""
    End If
  End If
Case "#L"
  PrNo = Conv(UFCask.StartTB)
  UFCask.Convex = Format(val(PrNo), "0.0")
  If NumericConfi("C") = False Then
    MsgBox "コンベックス校正を正しく実行してください！！", vbExclamation
    UFCask.Convex.Caption = ""
  End If
Case Else
  UFCask.StartTB = ""
End Select
UFCask.StartTB = ""
End Sub
Private Sub UserForm_Activate()
Dim a As Long, i As Long, k As Long
Application.Visible = True
Application.WindowState = xlMaximized '最大化
DoEvents

' --- 1. シートの可視化処理 ---
With ThisWorkbook
    For i = 1 To .Sheets.Count
       If .Sheets(i).Name = "キャスク検査" Or .Sheets(i).Name = "ｷｬｽｸ作業中保存" Or .Sheets(i).Name = "キャスク保存" Then
           .Sheets(i).Visible = True
       End If
    Next i
End With

' --- 2. シートの存在チェック ---
Dim ws As Worksheet
On Error Resume Next ' エラーが発生しても次の行に進む
Set ws = ThisWorkbook.Worksheets("キャスク検査")
On Error GoTo 0 ' エラーハンドリングを元に戻す

' シートが見つからなかった場合（wsが空っぽのとき）
If ws Is Nothing Then
    MsgBox "「キャスク検査」シートが見つかりません。" & vbCrLf & _
           "処理を中断します。", vbCritical, "エラー"
    Unload Me ' フォームを閉じる
    Exit Sub  ' 処理を抜ける
End If

' --- 3. シートが存在する場合の継続処理 ---
Call ProofreadingSearch(YeaDayrCheck(Format$(Now, "hh:mm")), TimeCheck(Format$(Now, "hh:mm:ss"))) '校正読み込み
ws.Activate
End Sub

Private Sub UserForm_Click()
Application.Visible = True
Application.WindowState = xlMaximized '最大化
UFCask.Loadtext.Caption = "on"
End Sub

Private Sub UserForm_Initialize()
If UFCask.Tag = "off" Then Exit Sub
UFCask.ComboBox1.Tag = "Load"
Dim a As Long, i As Long, k As Long
Dim SheetsName As String
Application.EnableEvents = True
Application.ScreenUpdating = False
'UFCask.ComboBox1.Clear
'k = ThisWorkbook.Sheets("LOT一覧").Cells(Rows.Count, 1).End(xlUp).Row
'For a = 2 To k '空白は含まない
'  If ThisWorkbook.Sheets("LOT一覧").Cells(a, 3) = "" Then
'  Else
'   UFCask.ComboBox1.AddItem ThisWorkbook.Sheets("LOT一覧").Cells(a, 3)
'  End If
'Next a
Application.ScreenUpdating = True
  If WsCheck("キャスク検査") Then
  UFCask.Loadtext.Caption = "off"
    With ThisWorkbook.Sheets("キャスク検査")
     .Activate
      UFCask.LotBox = .Cells(2, 4)
      UFCask.KenNo = .Cells(2, 25)
      UFCask.Kname = .Cells(3, 25)
      UFCask.Cno = .Cells(2, 26)
      UFCask.Mno = .Cells(3, 26)
      UFCask.SeriNo = .Cells(4, 26)
      
      UFCask.Micro = .Cells(2, 29) '2021.07.09_校正追加
      UFCask.Convex = .Cells(3, 29)
      UFCask.Convex_1000 = .Cells(3, 31)
      
      UFCask.CountSt = .Cells(3, 33) '2022.11.18_板厚検査回数カウント
      Call ループ判定
      If Me.Inspection = "OFF" Then
         .Unprotect '保護解除
         .Range("K6:AB706").Interior.Color = RGB(128, 128, 128)
         .Range("K7:AB706").Locked = True 'ロック設定
         .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
      End If
    End With
  End If
Application.ScreenUpdating = False
Call VisibleSheets 'シート表示
Call HiddenSheets '隠す

Call GetZoom(Me)
Dim Temp As Variant:  Temp = InitialSetting(Me)


Dim strLot As String: strLot = Me.LotBox
Dim strKen As String: strKen = Me.KenNo
Call AddCaskLabels(Me, strLot, strKen, 18, 300, 90, 20, 6) 'ラベル生成 梱包板番呼び出し

Application.ScreenUpdating = True
UFCask.ComboBox1.Tag = ""
UFCask.Loadtext.Caption = "on"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim i As Long
Application.Visible = True
Application.WindowState = xlMaximized '最大化
Application.ScreenUpdating = False
With ThisWorkbook
For i = 1 To .Sheets.Count
   If .Sheets(i).Name = "キャスク検査" Or .Sheets(i).Name = "ｷｬｽｸ作業中保存" Or .Sheets(i).Name = "キャスク保存" Then
   .Sheets(i).Visible = False
   End If
Next i
End With
With ThisWorkbook
For i = 1 To .Sheets.Count
   If .Sheets(i).Name = "Main" Then
   .Sheets(i).Visible = True
   End If
Next i
End With
Application.ScreenUpdating = True
End Sub

'（入力チェック）_2026.02.20追加
Private Function ValidateCaskInput() As Boolean

    Dim chk(1 To 4) As String
    Dim msg(1 To 4) As String
    Dim i            As Integer

    chk(1) = UFCask.LotBox.Value: msg(1) = "LotNo が未入力です。"
    chk(2) = UFCask.KenNo.Value: msg(2) = "検査No が未入力です。"
    chk(3) = UFCask.MINText.Value: msg(3) = "スタートNo が未入力です。"
    chk(4) = UFCask.MAXText.Value: msg(4) = "エンドNo が未入力です。"

    For i = 1 To 4
        If Trim(chk(i)) = "" Then
            MsgBox msg(i), vbExclamation
            ValidateCaskInput = False
            Exit Function
        End If
    Next i

    ValidateCaskInput = True

End Function

'-----------------------------------------------------------
' 同じ LotNo + 検査No のDB既存レコードに,入力分を加えてスタートNo順に並べエンドNo + 1 = 次のスタートNo になっているか確認する
'つまり連番チェック
' 同じスタートNoのレコードは上書き対象のため除外
'-----------------------------------------------------------
Private Function CheckCaskSequence() As Boolean

    Dim db       As Object
    Dim rs       As Object
    Dim strPath  As String
    Dim strSQL   As String
    Dim strLot   As String
    Dim strKen   As String
    Dim lngStart As Long
    Dim lngEnd   As Long

    '--- 動的配列（件数制限なし）---
    Dim arrS() As Long   'スタートNo
    Dim arrE() As Long   'エンドNo
    Dim Cnt    As Long
    Dim i         As Long
    Dim j         As Long
    Dim tmpS      As Long
    Dim tmpE      As Long
    Dim strMsg    As String

    CheckCaskSequence = True   'デフォルトOK

    On Error GoTo ErrHandler

    strLot = UFCask.LotBox.Value
    strKen = UFCask.KenNo.Value
    lngStart = CLng(UFCask.MINText.Value)
    lngEnd = CLng(UFCask.MAXText.Value)

    '--- スタートNo > エンドNo の単体チェック ---
    If lngStart > lngEnd Then
        MsgBox "スタートNo がエンドNo より大きい値です。" & vbCrLf & _
               "スタートNo：" & lngStart & "　エンドNo：" & lngEnd, vbExclamation
        CheckCaskSequence = False
        Exit Function
    End If

    '--- スタートNo=1 の場合は他レコードを全削除するため連番チェック不要 ---
    If lngStart = 1 Then
        CheckCaskSequence = True
        Exit Function
    End If

    '--- DB接続 ---
    strPath = QualityPath & "伝送用ファイル.accdb"
    Set db = CreateObject("ADODB.Connection")
    db.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strPath & ";"

    '--- 同じLotNo+検査No で 今回スタートNoと異なるレコードを取得 ---
    strSQL = "SELECT スタートNo, エンドNo FROM キャスク梱包板番号 " & _
             "WHERE LotNo='" & strLot & "' " & _
             "AND 検査No='" & strKen & "' " & _
             "AND スタートNo <> " & lngStart & " " & _
             "ORDER BY スタートNo"

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, db, 3, 1   ' adOpenStatic, adLockReadOnly

    '--- 既存レコードを配列に格納 ---
    Cnt = 0
    ReDim arrS(0)
    ReDim arrE(0)
    Do While Not rs.EOF
        ReDim Preserve arrS(Cnt)
        ReDim Preserve arrE(Cnt)
        arrS(Cnt) = CLng(rs.Fields("スタートNo").Value)
        arrE(Cnt) = CLng(rs.Fields("エンドNo").Value)
        Cnt = Cnt + 1
        rs.MoveNext
    Loop

    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing

    '--- 今回入力分を追加 ---
    ReDim Preserve arrS(Cnt)
    ReDim Preserve arrE(Cnt)
    arrS(Cnt) = lngStart
    arrE(Cnt) = lngEnd
    Cnt = Cnt + 1

    '--- レコードが1件のみなら連番チェック不要 ---
    If Cnt <= 1 Then
        CheckCaskSequence = True
        Exit Function
    End If

    '--- バブルソート（スタートNo 昇順）---
    For i = 0 To Cnt - 2
        For j = 0 To Cnt - 2 - i
            If arrS(j) > arrS(j + 1) Then
                tmpS = arrS(j): arrS(j) = arrS(j + 1): arrS(j + 1) = tmpS
                tmpE = arrE(j): arrE(j) = arrE(j + 1): arrE(j + 1) = tmpE
            End If
        Next j
    Next i

    '--- 連番チェック：エンドNo + 1 = 次のスタートNo ---
    strMsg = ""
    For i = 0 To Cnt - 2
        If arrE(i) + 1 <> arrS(i + 1) Then
            strMsg = strMsg & "　[" & arrS(i) & "-" & arrE(i) & "] → [" & _
                      arrS(i + 1) & "-" & arrE(i + 1) & "]" & vbCrLf
        End If
    Next i

    If strMsg <> "" Then
        MsgBox "番号が連番になっていない箇所があります。", vbExclamation, "連番チェック"
        CheckCaskSequence = False
        Exit Function
    End If

    CheckCaskSequence = True
    Exit Function

ErrHandler:
    MsgBox "連番チェックエラー：" & vbCrLf & Err.Description, vbCritical
    CheckCaskSequence = False
    If Not rs Is Nothing Then
        If rs.state = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then
        db.Close: Set db = Nothing
    End If

End Function

'（DB保存）
Private Function SaveCaskToDB(ByVal dtWork As Date) As Boolean
    Dim db       As Object
    Dim rs       As Object
    Dim daoEng   As Object
    Dim daoDB    As Object
    Dim daoRS    As Object
    Dim strPath  As String
    Dim strSQL   As String
    Dim strLot   As String
    Dim strKen   As String
    Dim strStart As String
    Dim strEnd   As String

    SaveCaskToDB = False
    On Error GoTo ErrHandler

    strLot = UFCask.LotBox.Value
    strKen = UFCask.KenNo.Value
    strStart = UFCask.MINText.Value
    strEnd = UFCask.MAXText.Value
    strPath = QualityPath & "伝送用ファイル.accdb"

    DebugLog "[SaveCaskToDB] ===== 開始 ====="
    DebugLog "[SaveCaskToDB] LotNo=" & strLot & " 検査No=" & strKen & " スタートNo=" & strStart & " エンドNo=" & strEnd

    '============================================================
    ' STEP1：ADODB で UPDATE / INSERT
    '============================================================
    Set db = CreateObject("ADODB.Connection")
    db.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strPath & ";"

    strSQL = "SELECT * FROM キャスク梱包板番号 " & _
             "WHERE LotNo='" & strLot & "' " & _
             "AND 検査No='" & strKen & "' " & _
             "AND スタートNo=" & strStart

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, db, 1, 3

    DebugLog "[SaveCaskToDB] STEP1 RS EOF=" & rs.EOF & " RecordCount=" & rs.RecordCount

    If Not rs.EOF Then
        DebugLog "[SaveCaskToDB] STEP1 → UPDATE"
        rs.Fields("エンドNo").Value = strEnd
        rs.Fields("作業時間").Value = dtWork
        rs.Update
    Else
        DebugLog "[SaveCaskToDB] STEP1 → INSERT"

        '--- 管理番号を手動採番（MAX+1）---
        '    オートナンバーは削除後に欠番になるため手動で連番を維持する
        Dim rsMax    As Object
        Dim lngNextNo As Long
        Set rsMax = CreateObject("ADODB.Recordset")
        rsMax.Open "SELECT MAX(管理番号) AS MaxNo FROM キャスク梱包板番号", db, 3, 1
        If IsNull(rsMax.Fields("MaxNo").Value) Then
            lngNextNo = 1
        Else
            lngNextNo = CLng(rsMax.Fields("MaxNo").Value) + 1
        End If
        rsMax.Close: Set rsMax = Nothing
        DebugLog "[SaveCaskToDB] STEP1 管理番号採番=" & lngNextNo

        rs.AddNew
        rs.Fields("管理番号").Value = lngNextNo
        rs.Fields("LotNo").Value = strLot
        rs.Fields("検査No").Value = strKen
        rs.Fields("スタートNo").Value = strStart
        rs.Fields("エンドNo").Value = strEnd
        rs.Fields("作業時間").Value = dtWork
        rs.Update
    End If

    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    DebugLog "[SaveCaskToDB] STEP1 完了"

    '============================================================
    ' STEP2：スタートNo=1 のとき DAO で物理削除
    '============================================================
    If CLng(strStart) = 1 Then

        DebugLog "[SaveCaskToDB] STEP2 DAO削除開始"

        strSQL = "SELECT * FROM キャスク梱包板番号 " & _
                 "WHERE LotNo='" & strLot & "' " & _
                 "AND 検査No='" & strKen & "' " & _
                 "AND スタートNo <> 1"

        Set daoEng = CreateObject("DAO.DBEngine.120")
        Set daoDB = daoEng.OpenDatabase(strPath)
        Set daoRS = daoDB.OpenRecordset(strSQL, 2)

        Dim delCount As Long
        delCount = 0
        Do While Not daoRS.EOF
            DebugLog "[SaveCaskToDB] STEP2 削除 スタートNo=" & daoRS.Fields("スタートNo").Value
            daoRS.Delete
            daoRS.MoveNext
            delCount = delCount + 1
        Loop

        DebugLog "[SaveCaskToDB] STEP2 削除完了 件数=" & delCount

        daoRS.Close: Set daoRS = Nothing
        daoDB.Close: Set daoDB = Nothing
        Set daoEng = Nothing

    Else
        DebugLog "[SaveCaskToDB] STEP2 スキップ"
    End If

    SaveCaskToDB = True
    DebugLog "[SaveCaskToDB] ===== 正常終了 ====="
    Exit Function

ErrHandler:
    DebugLog "[SaveCaskToDB] ★エラー Err=" & Err.Number & " / " & Err.Description
    MsgBox "DB保存エラー：" & vbCrLf & Err.Description, vbCritical
    If Not rs Is Nothing Then
        If rs.state = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then
        db.Close: Set db = Nothing
    End If
    If Not daoRS Is Nothing Then
        daoRS.Close: Set daoRS = Nothing
    End If
    If Not daoDB Is Nothing Then
        daoDB.Close: Set daoDB = Nothing
    End If
    Set daoEng = Nothing

End Function
