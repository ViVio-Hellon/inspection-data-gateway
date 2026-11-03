Attribute VB_Name = "UFMitsubishi"
Attribute VB_Base = "0{D8D3F607-E22C-4B54-BBA0-21C027DFC553}{A2F8B868-06F5-4F14-8EAB-8C8E351ED11F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Option Explicit
Private g_lngWindowState As XlWindowState                       ' ウィンドウ状態
' Windows API定数定義
Private Const CF_BITMAP = 2
Private Const CF_PALETTE = 9
Private Const CF_ENHMETAFILE = 14

' Windows API関数宣言
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PictDesc, riid As GUID, ByVal fOwn As Long, ipic As IPicture) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PictDesc, riid As GUID, ByVal fOwn As Long, ipic As IPicture) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' 構造体定義
#If VBA7 Then
    Private Type PictDesc
        cbSizeofStruct As Long
        picType As Long
        hImage As LongPtr
        Option1 As LongPtr
    End Type
#Else
    Private Type PictDesc
        cbSizeofStruct As Long
        picType As Long
        hImage As Long
        Option1 As Long
    End Type
#End If

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Sub Cle_Click()
Dim Lname As String: Lname = UFMitsubishi.LineName.Caption
If (WsCheck("不良要因印刷" & "_" & Lname) = False) Then Exit Sub

Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname)
Dim bRange As Range: Set bRange = Range(sh.Cells(3, 3), sh.Cells(7, 24))
Dim aRange As Range: Set aRange = Range(sh.Cells(9, 3), sh.Cells(9, 24))
Dim cRange As Range: Set cRange = sh.Cells(10, 3)
bRange.ClearContents
aRange.ClearContents
cRange.ClearContents
End Sub


Private Sub CommandButton18_Click()
'注意表示---------------------------------------------------------------------------------------------
If WsCheck("全ヵ所検査") = False Then Exit Sub
Dim result1 As Long, result2 As Long, myRange1 As Range, myRange2 As Range, myRange3 As Range, myRange4 As Range, WSZ As Worksheet: Set WSZ = ThisWorkbook.Worksheets("全ヵ所検査")
Set myRange1 = Range(WSZ.Cells(7, 2), WSZ.Cells(76, 9))
Set myRange2 = Range(WSZ.Cells(7, 25), WSZ.Cells(76, 28))
Set myRange3 = Range(WSZ.Cells(2, 30))
Set myRange4 = Range(WSZ.Cells(3, 30))

If myRange3 = "" And myRange4 = "" Then
  MsgBox "検入高さを入力してください。", vbCritical
  Exit Sub
End If

result1 = CheckRedXCells(myRange1)
result2 = CheckRedXCells(myRange2)
If result1 > 0 And result2 > 0 Then
  If MsgBox("公差外判定の板があります " & "厚:" & result1 & "箇所　" & "歪:" & result2 & "箇所" & vbCrLf & "処理を続行しますか？", vbQuestion + vbYesNo) = vbNo Then Exit Sub
ElseIf result1 > 0 And result2 = 0 Then
  If MsgBox("公差外判定の板があります " & "厚:" & result1 & "箇所　" & vbCrLf & "処理を続行しますか？", vbQuestion + vbYesNo) = vbNo Then Exit Sub
ElseIf result1 = 0 And result2 > 0 Then
  If MsgBox("公差外判定の板があります " & "歪:" & result2 & "箇所" & vbCrLf & "処理を続行しますか？", vbQuestion + vbYesNo) = vbNo Then Exit Sub
End If
'------------------------------------------------------------------------------------------------------
End Sub

Private Sub CommandButton19_Click()
If UFMitsubishi.Tag = "off" Then Exit Sub

Dim i As Long, FL As Boolean, KenM As String
For i = 1 To 3
  If UFMitsubishi("op" & i) = True Then
    KenM = UFMitsubishi("op" & i).Caption
    FL = True
    Exit For
  End If
Next i
If FL = False Then
  MsgBox "検査法を選択してください", vbCritical
  Exit Sub
End If

Application.EnableEvents = False
UFProgress.Show vbModeless
UFProgress.CommandButton1.Caption = "転送中...."
Call ProgressBar_count(1, 730)
Call 全ヵ所用最大最少("H494検査")
Call ProgressBar_count(730, 1000)
Unload UFProgress
Call Ainput_Startup("H494", KenM) '起動ﾌｫｰﾑ、検査法
Application.EnableEvents = True
End Sub

Private Sub CommandButton21_Click()
Application.EnableEvents = True
Application.Visible = True 'Excel表示
End Sub

Private Sub CommandButton22_Click()
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("H494検査")
ws.Unprotect
ws.Range("AC:AC").EntireColumn.AutoFit
ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'シート保護
End Sub

Private Sub CommandButton23_Click()
Dim shName As String, ws As Worksheet, n As Long
Dim Lname As String: Lname = UFMitsubishi.LineName.Caption
Dim ZenSh As Worksheet: Set ZenSh = ThisWorkbook.Sheets("H494検査")

Dim ATU As String:  ATU = ZenSh.Cells(3, 2)
Dim HABA As String: HABA = ZenSh.Cells(3, 4)
  
UFMitsubishi.Tag = "off"
ZenSh.Activate
If UFMitsubishi.Kname = "" And UFMitsubishi.H494_Kname = "" Then
 MsgBox "検査員名と検査番号を入力してください"
Exit Sub
Else
UFMitsubishi.Tag = "off"
Call MusicTrue("おまちください.wav")
  UFProgress.Show vbModeless
  Call ProgressBar_count(0, 200)
  Call 全ヵ所用機器転送(ZenSh, 2, 25, ATU, HABA, "H494") 'H494指定
  DoEvents
  Application.ScreenUpdating = False
  With ActiveSheet.PageSetup
      .LeftHeader = ""
      .CenterHeader = ""
      .RightHeader = ""
      .LeftFooter = ""
      .CenterFooter = ""
      .LeftFooter = ""
      .CenterFooter = ""
      .RightFooter = ""
  End With
    Call 全ヵ所用印刷範囲("H494検査")
    Call MusicTrue("印刷.wav")
    If MsgBox("印刷しますか？", vbQuestion + vbYesNo) = vbYes Then
        ZenSh.Visible = True
        ZenSh.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
        Call ProgressBar_count(200, 450)
        DoEvents
        Call 全ヵ所専用保存("H494検査")
        Call ProgressBar_count(450, 880)
        '-------------------------------------------------------------------不良要因
        If (WsCheck("不良要因印刷" & "_" & Lname) = False) Then 'シートなし
          Call AddWS("不良要因印刷" & "_" & Lname)
          Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
        ElseIf ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Cells(1, 2) <> UFMitsubishi.LotBox Then '不良の入力がなければこの状態
          Call AddWS("不良要因印刷" & "_" & Lname)
          Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname)
        End If
        Call 不良要因印刷範囲(UFMitsubishi.H494_Kname)
        Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname)
        
        sh.Visible = True
        sh.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
        sh.Visible = False
        
        '-------------------------------------------------------------------
        ThisWorkbook.Activate
          Application.DisplayAlerts = False '保存後シートは消す
            If WsCheck("ｺﾋﾟ-用") Then ThisWorkbook.Sheets("ｺﾋﾟ-用").Delete
            If WsCheck("H494検査保存") Then ThisWorkbook.Sheets("H494検査保存").Delete
            If WsCheck("H494作業中保存") Then ThisWorkbook.Sheets("H494作業中保存").Delete
          Application.DisplayAlerts = True
    '----------------------------------------------履歴を残す　2019.09.13　これ以前のものは履歴なし
        If WsCheck("保存用") Then
          With ThisWorkbook.Sheets("保存用")
            n = .Cells(Rows.Count, 1).End(xlUp).row + 1
            .Unprotect
            .Cells.Locked = False
            .Cells.FormulaHidden = False
            .Cells(n, 2) = ThisWorkbook.Sheets("H494検査").Cells(2, 4) 'ロット
            .Cells(n, 83) = "H494"
            .Cells(n, 3) = ThisWorkbook.Sheets("H494検査").Cells(3, 2) '厚
            .Cells(n, 4) = ThisWorkbook.Sheets("H494検査").Cells(3, 4) '幅
            .Cells(n, 5) = ThisWorkbook.Sheets("H494検査").Cells(3, 6) '丈
            .Cells(n, 1) = Format$(ThisWorkbook.Sheets("H494検査").Cells(1, 2), "yyyy/mm/dd  h:mm") & "～" & Format$(Now, "yyyy/mm/dd  h:mm")
            .Cells(n, 77) = "別ファイル参照" '検査員
          End With
        End If
      Call ProgressBar_count(880, 1000)
      Call MusicTrue("完了.wav")
      DoEvents
      Unload UFProgress
    Else
      Unload UFProgress
      MsgBox "ｷｬﾝｾﾙ処理", vbInformation
    End If
End If
Application.ScreenUpdating = True
Application.WindowState = xlMaximized
Application.Visible = True

UFMitsubishi.Tag = ""
End Sub

Private Sub CommandButton24_Click()
UF品質基準.LineName = UFMitsubishi.LineName
UF品質基準.T1 = UFMitsubishi.LotBox
UF品質基準.FlagAtu = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 2)
UF品質基準.MP.Page1.Visible = True
UF品質基準.MP.Page2.Visible = True
'厚板,薄板の判定を行う
'Page1=0　Page2=1
If UF品質基準.FlagAtu = "" Then Exit Sub
  '---------------------------------------------------------------------------------------
'2023.4.12_品証管理がしやすいようDB関数作成
If UF品質基準.FlagAtu >= 4 Then
  UF品質基準.MP.Value = 0
Else
  UF品質基準.MP.Value = 1
End If

Dim SetTable As String: SetTable = AddPagesForSubFolders(UF品質基準.MPH) 'ページ追加
UF品質基準.TableTB = "「追加ページ一覧」" & vbCrLf & vbCrLf & SetTable

Call LoadImagesIntoMultiPageFromTemp(UF品質基準.MPH) '写真追加

UF品質基準.Show vbModeless '歪検索より先に表示

Dim QpArry As Variant:         If UFMitsubishi.LotBox <> "" Then QpArry = DistArray(UFMitsubishi.LotBox, UFMitsubishi.KenNo, "", "")  'ゲージ、幅歪、丈歪、中、耳、コメント、参照元、番号、フラグ
'Dim QpArryG As Variant:        If UFMitsubishi.LotBox <> "" Then QpArryG = DistArray(UFMitsubishi.LotBox, UFMitsubishi.KenNo, "取引", "", "")
'Dim UpdatedQpArryD As Variant: If UFMitsubishi.LotBox <> "" Then UpdatedQpArryD = UpdateQpArryD(QpArry, QpArryG)

If IsEmptyArray(QpArry) = False Then Call FormReflectDist(QpArry, "") 'フォーム反映

If UF品質基準.FlagAtu > 6 Then GradeArray = DBdisplay外観_H(UFMitsubishi.LotBox, "")
End Sub

Private Sub CommandButton25_Click()
Me.Width = 527.25
Me.Height = 332.25
Me.Zoom = 100
End Sub

Private Sub FiveOP_Click()
UFMitsubishi.Choice = "5N"
Call ForN_Select
End Sub

Private Sub Frame2_Click()
Dim ws As Worksheet
If WorkFlag.Caption = "P1" Then
  If WsCheck("全ヵ所検査") Then ThisWorkbook.Worksheets("全ヵ所検査").Activate
  Set ws = ThisWorkbook.Worksheets("全ヵ所検査")
End If
If WorkFlag.Caption = "P_H494" Then
  If WsCheck("H494検査") Then ThisWorkbook.Worksheets("H494検査").Activate
  Set ws = ThisWorkbook.Worksheets("H494検査")
End If

Dim TempW As Double: TempW = Weight(val(ws.Cells(3, 2)), val(ws.Cells(3, 4)), val(ws.Cells(3, 6)))
If TempW = 0 Then Exit Sub '重量なし処理停止
If 35 >= TempW Then MsgBox "ゲージ不要 * 両面定盤歪測定必要"
If 35 < TempW Then MsgBox "ゲージによる測定必要"
End Sub

Private Sub Frame6_Click()
Dim ws As Worksheet
If WorkFlag.Caption = "P1" Then
  If WsCheck("全ヵ所検査") Then ThisWorkbook.Worksheets("全ヵ所検査").Activate
  Set ws = ThisWorkbook.Worksheets("全ヵ所検査")
End If
If WorkFlag.Caption = "P_H494" Then
  If WsCheck("H494検査") Then ThisWorkbook.Worksheets("H494検査").Activate
  Set ws = ThisWorkbook.Worksheets("H494検査")
End If

Dim TempW As Double: TempW = Weight(val(ws.Cells(3, 2)), val(ws.Cells(3, 4)), val(ws.Cells(3, 6)))
If TempW = 0 Then Exit Sub '重量なし処理停止
If 35 >= TempW Then MsgBox "ゲージ不要 * 両面定盤歪測定必要"
If 35 < TempW Then MsgBox "ゲージによる測定必要"
End Sub

Private Sub H494_Lot_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Private Sub H494_Lot_Change()
Application.EnableEvents = False
If Len(Me.H494_Lot.Text) <> 7 Then
Exit Sub
Else
 Me.ComboBox1 = Me.H494_Lot
End If
Application.EnableEvents = True
ActiveWindow.SmallScroll up:=99
End Sub
Private Sub KenHBox1_AfterUpdate()
UFMitsubishi.KenHBox1 = Format((UFMitsubishi.KenHBox1), "0.0")
Call LogOut(KenHBox1, "", "三菱")
If WsCheck("全ヵ所検査") = False Then Exit Sub
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("全ヵ所検査")
If ws.Cells(3, 2) = "" Then Exit Sub
Call DrawListBoxContent(UFMitsubishi, val(ws.Cells(3, 2)))
End Sub
Private Sub KenHBox2_AfterUpdate()
UFMitsubishi.KenHBox2 = Format((UFMitsubishi.KenHBox2), "0.0")
Call LogOut(KenHBox2, "", "三菱")
If WsCheck("全ヵ所検査") = False Then Exit Sub
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("全ヵ所検査")
If ws.Cells(3, 2) = "" Then Exit Sub
Call DrawListBoxContent(UFMitsubishi, val(ws.Cells(3, 2)))
End Sub
Private Sub KenHBox3_AfterUpdate()
UFMitsubishi.KenHBox3 = Format((UFMitsubishi.KenHBox3), "0.0")
Call LogOut(KenHBox3, "", "三菱")
If WsCheck("全ヵ所検査") = False Then Exit Sub
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("全ヵ所検査")
If ws.Cells(3, 2) = "" Then Exit Sub
Call DrawListBoxContent(UFMitsubishi, val(ws.Cells(3, 2)))
End Sub
Private Sub KenHBox4_AfterUpdate()
UFMitsubishi.KenHBox4 = Format((UFMitsubishi.KenHBox4), "0.0")
Call LogOut(KenHBox4, "", "三菱")
If WsCheck("全ヵ所検査") = False Then Exit Sub
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("全ヵ所検査")
If ws.Cells(3, 2) = "" Then Exit Sub
Call DrawListBoxContent(UFMitsubishi, val(ws.Cells(3, 2)))
End Sub
Private Sub KenHBox5_AfterUpdate()
UFMitsubishi.KenHBox5 = Format((UFMitsubishi.KenHBox5), "0.0")
Call LogOut(KenHBox5, "", "三菱")
If WsCheck("全ヵ所検査") = False Then Exit Sub
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("全ヵ所検査")
If ws.Cells(3, 2) = "" Then Exit Sub
Call DrawListBoxContent(UFMitsubishi, val(ws.Cells(3, 2)))
End Sub
Private Sub KenHBox6_AfterUpdate()
UFMitsubishi.KenHBox6 = Format((UFMitsubishi.KenHBox6), "0.0")
Call LogOut(KenHBox6, "", "三菱")
If WsCheck("全ヵ所検査") = False Then Exit Sub
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("全ヵ所検査")
If ws.Cells(3, 2) = "" Then Exit Sub
Call DrawListBoxContent(UFMitsubishi, val(ws.Cells(3, 2)))
End Sub
Private Sub KenHBox7_AfterUpdate()
UFMitsubishi.KenHBox6 = Format((UFMitsubishi.KenHBox6), "0")
Call LogOut(KenHBox6, "", "三菱")
If WsCheck("全ヵ所検査") = False Then Exit Sub
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("全ヵ所検査")
If ws.Cells(3, 2) = "" Then Exit Sub
ws.Unprotect
ws.Cells(3, 30) = Me.KenHBox7 & "枚"
ws.Protect Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
End Sub

Private Sub ListReflect_Click()
Dim Temps As String: Temps = GetSelectedItems(UFMitsubishi.DefectFactorLB)

If UFMitsubishi.DefectFactorLB.ListIndex = -1 Then
  MsgBox "不良要因が選択されていません", vbCritical
  Exit Sub
End If

Dim FL As Boolean: If UFMitsubishi("SBad") = "" And UFMitsubishi("UBad") = "" And UFMitsubishi("CBad") = "" And UFMitsubishi("DBad") = "" And UFMitsubishi("EBad") = "" Then FL = True
If FL Then
  MsgBox "枚数入力がありません", vbCritical
  Exit Sub
End If

Dim S As Long: If IsNumeric(UFMitsubishi("SBad")) Then S = val(UFMitsubishi("SBad"))
Dim U As Long: If IsNumeric(UFMitsubishi("UBad")) Then U = val(UFMitsubishi("UBad"))
Dim C As Long: If IsNumeric(UFMitsubishi("CBad")) Then C = val(UFMitsubishi("CBad"))
Dim D As Long: If IsNumeric(UFMitsubishi("DBad")) Then D = val(UFMitsubishi("DBad"))
Dim E As Long: If IsNumeric(UFMitsubishi("EBad")) Then E = val(UFMitsubishi("EBad"))

'スポット処理は囲い文字処理なし_2024.1.10
If Temps = "ｽﾎﾟｯﾄ処理" Then GoTo SkipPrint
If S > 50 Or U > 50 Or C > 50 Or D > 50 Or E > 50 Then
  MsgBox "すみません、50枚以上は分けて入力してください" '囲い文字処理の関係で50まで
  Exit Sub
End If

SkipPrint:
Dim n As String: n = TecReturn
Call SelectionInput(UFMitsubishi(n), "全数用") '出力

'リセット
Dim i As Long
For i = 1 To 5
  Select Case i
  Case 1
    UFMitsubishi("SBad") = ""
    UFMitsubishi("SBad").BackColor = RGB(255, 255, 255)
  Case 2
    UFMitsubishi("UBad") = ""
    UFMitsubishi("UBad").BackColor = RGB(255, 255, 255)
  Case 3
    UFMitsubishi("CBad") = ""
    UFMitsubishi("CBad").BackColor = RGB(255, 255, 255)
  Case 4
    UFMitsubishi("DBad") = ""
    UFMitsubishi("DBad").BackColor = RGB(255, 255, 255)
  Case 5
    UFMitsubishi("EBad") = ""
    UFMitsubishi("EBad").BackColor = RGB(255, 255, 255)
  End Select
Next i
End Sub


Private Sub NgIm_Click()
Dim i As Long
With UFMitsubishi
  .NgIm.Height = 1
  .NgIm.Width = 1
  .NgIm.Visible = False
  .DefectFactorLB.Visible = True
  
  .Height = 325.5
  .Width = 525
  
  .MLP.Height = 282
  .MLP.Width = 492
  .Repaint
End With

Call GetZoom(Me)
Dim MiTemp As Variant:  MiTemp = InitialSetting(Me)
End Sub

'三菱ﾌﾟﾚｰﾄ
'UFMitsubishi.Work = "ON" '三菱受託用フラグ
'UFMitsubishi.Inflag = "flagon" '三菱印刷用フラグ
'一般
'UFMitsubishi.Work = "OFF" '印刷用フラグ
'UFMitsubishi.Inflag = "" '印刷用フラグ
'5N
Private Sub SBad_Change()
If IsNumeric(UFMitsubishi.SBad) = False Then UFMitsubishi.SBad = ""
End Sub
Private Sub UBad_Change()
If IsNumeric(UFMitsubishi.UBad) = False Then UFMitsubishi.UBad = ""
End Sub
Private Sub CBad_Change()
If IsNumeric(UFMitsubishi.CBad) = False Then UFMitsubishi.CBad = ""
End Sub
Private Sub DBad_Change()
If IsNumeric(UFMitsubishi.DBad) = False Then UFMitsubishi.DBad = ""
End Sub
Private Sub EBad_Change()
If IsNumeric(UFMitsubishi.EBad) = False Then UFMitsubishi.EBad = ""
End Sub

Private Sub SBad_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UFMitsubishi.SBad = Format((UFMitsubishi.SBad), "0")
Call TeColor(SBad)
End Sub
Private Sub UBad_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UFMitsubishi.UBad = Format((UFMitsubishi.UBad), "0")
Call TeColor(UBad)
End Sub
Private Sub CBad_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UFMitsubishi.CBad = Format((UFMitsubishi.CBad), "0")
Call TeColor(CBad)
End Sub
Private Sub DBad_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UFMitsubishi.DBad = Format((UFMitsubishi.DBad), "0")
Call TeColor(DBad)
End Sub
Private Sub EBad_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UFMitsubishi.EBad = Format((UFMitsubishi.EBad), "0")
Call TeColor(EBad)
End Sub

Private Sub Choice_Change()
If UFMitsubishi.Choice.Text = "５N" Then UFMitsubishi.Choice.Text = "5N"
Select Case UFMitsubishi.Choice.Text
Case "三菱", "一般", "5N": Call ForN_Select
Case Else: MsgBox "正しい作業選択名を入力してください", vbExclamation
End Select
End Sub

Private Sub CMB_Click()
If UFMitsubishi.Tag = "off" Then Exit Sub
Dim ws As Worksheet
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
'  Range(ws.Cells(7, 27), ws.Cells(7, 29)).Locked = False 'ロック設定
'  Range(ws.Cells(9, 27), ws.Cells(9, 29)).Locked = False
'  Range(ws.Cells(11, 27), ws.Cells(11, 29)).Locked = False
'  Range(ws.Cells(13, 27), ws.Cells(13, 29)).Locked = False
'  Union(Range(ws.Cells(7, 27), ws.Cells(7, 29)), ws.Range(ws.Cells(9, 27), ws.Cells(9, 29)), _
'  Range(ws.Cells(11, 27), ws.Cells(11, 29)), Range(ws.Cells(13, 27), ws.Cells(13, 29))).Interior.Color = RGB(255, 255, 0)
'  Range(ws.Cells(6, 30), ws.Cells(14, 30)).Font.Color = RGB(0, 0, 0)
'  Range(ws.Cells(6, 31), ws.Cells(11, 31)).Font.Color = RGB(255, 255, 255)
'  ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
Case "5N"
'5N作業
  .Work = "OFF"
  .Inflag = ""
  .Normal.Visible = False
  .Mitsubishi.Visible = False
  .FiNi.Visible = True
  ws.Unprotect '保護解除
  ws.Cells(1, 22) = "5N"
'  Range(ws.Cells(7, 27), Cells(7, 29)).Locked = True 'ロック設定
'  Range(ws.Cells(9, 27), Cells(9, 29)).Locked = True
'  Range(ws.Cells(11, 27), ws.Cells(11, 29)).Locked = True
'  Range(ws.Cells(13, 27), ws.Cells(13, 29)).Locked = True
'  Union(Range(ws.Cells(7, 27), ws.Cells(7, 29)), Range(ws.Cells(9, 27), ws.Cells(9, 29)), _
'  Range(ws.Cells(11, 27), ws.Cells(11, 29)), Range(ws.Cells(13, 27), ws.Cells(13, 29))).Interior.Color = RGB(192, 192, 192)
'  Range(ws.Cells(6, 30), ws.Cells(14, 30)).Font.Color = RGB(255, 255, 255)
'  Range(ws.Cells(6, 31), ws.Cells(11, 31)).Font.Color = RGB(0, 0, 0)
'  ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
Case "一般"
'一般作業
  .Work = "OFF"
  .Inflag = ""
  .Normal.Visible = True
  .Mitsubishi.Visible = False
  .FiNi.Visible = False
  ws.Unprotect '保護解除
  ws.Cells(1, 22) = "一般"
'  Range(ws.Cells(7, 27), Cells(7, 29)).Locked = True 'ロック設定
'  Range(ws.Cells(9, 27), Cells(9, 29)).Locked = True
'  Range(ws.Cells(11, 27), ws.Cells(11, 29)).Locked = True
'  Range(ws.Cells(13, 27), ws.Cells(13, 29)).Locked = True
'  Union(Range(ws.Cells(7, 27), ws.Cells(7, 29)), Range(ws.Cells(9, 27), ws.Cells(9, 29)), _
'  Range(ws.Cells(11, 27), ws.Cells(11, 29)), Range(ws.Cells(13, 27), ws.Cells(13, 29))).Interior.Color = RGB(192, 192, 192)
'  Range(ws.Cells(6, 30), ws.Cells(14, 30)).Font.Color = RGB(255, 255, 255)
'  Range(ws.Cells(6, 31), ws.Cells(11, 31)).Font.Color = RGB(255, 255, 255)
  ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
End Select
End With
End Sub

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

'ADO 化_2026.05.29
Private Sub ComboBox1_Change()
If UFMitsubishi.ComboBox1.Tag = "Load" Then Exit Sub
Application.EnableEvents = False
Application.ScreenUpdating = False '処理高速化のための画面更新停止
If WsCheck("全ヵ所検査") = False Then Exit Sub
If Len(UFMitsubishi.ComboBox1.Text) <> 7 Then Exit Sub
If UFMitsubishi.ComboBox1.Tag = "Load" Then Exit Sub

Dim i As Long, j As Long
Dim PastHiki(0 To 18)
Dim NoFl     As Boolean: NoFl = False

Dim db       As Object
Dim rs       As Object
Dim strPath  As String
Dim strSQL   As String

Dim ZenSh As Worksheet
If UFMitsubishi.WorkFlag.Caption = "P1" Then Set ZenSh = ThisWorkbook.Sheets("全ヵ所検査")
If UFMitsubishi.WorkFlag.Caption = "P_H494" Then Set ZenSh = ThisWorkbook.Sheets("H494検査")

ZenSh.Unprotect
ZenSh.Cells(1, 15) = ""
ZenSh.Cells(2, 4).Font.Color = RGB(0, 0, 0)

On Error GoTo ErrHandler_MitCombo

strPath = MainPath & "SIKALOTNOW.accdb"
If Dir(strPath) = "" Then
    MsgBox "DBファイルが見つかりません:" & vbCrLf & strPath, vbCritical
    GoTo ErrHandler_MitCombo
End If

strSQL = "SELECT * FROM 仕掛 WHERE ﾛｯﾄ番号 = '" & Me.ComboBox1.Text & "' AND ｵｰﾀﾞｰ板丈 <> '0'"

Set db = CreateObject("ADODB.Connection")
db.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strPath & ";"
Set rs = CreateObject("ADODB.Recordset")
rs.Open strSQL, db, 3, 1

If Not rs.EOF Then
    NoFl = True

    PastHiki(0) = "'" & rs.Fields("ｵｰﾀﾞｰ番号").Value
    PastHiki(1) = rs.Fields("ﾛｯﾄ番号").Value
    PastHiki(2) = rs.Fields("製造板厚").Value
    PastHiki(3) = rs.Fields("製造板幅").Value
    PastHiki(4) = rs.Fields("製造板丈").Value
    PastHiki(5) = rs.Fields("検査番号").Value
    PastHiki(6) = rs.Fields("製造材質").Value
    PastHiki(7) = rs.Fields("製造調質").Value
    PastHiki(8) = rs.Fields("用途ｺｰﾄﾞ").Value
    PastHiki(9) = rs.Fields("設計_設備ｺｰｽ").Value
    PastHiki(10) = rs.Fields("取引先名").Value
    PastHiki(11) = rs.Fields("納入先名").Value
    PastHiki(12) = rs.Fields("送り先名").Value
    PastHiki(13) = rs.Fields("BOX実績_枚本数").Value
    PastHiki(14) = rs.Fields("BOX実績_良品重量").Value
    PastHiki(15) = rs.Fields("梱包_合紙").Value
    PastHiki(16) = rs.Fields("VC").Value
    PastHiki(17) = rs.Fields("品質ｸﾞﾚｰﾄﾞ_ﾗﾃﾗﾙﾎﾞｰ").Value

    ZenSh.Cells(2, 4) = Me.ComboBox1.Text
    UFMitsubishi.KenNo = rs.Fields("検査番号").Value
    UFMitsubishi.H494_KenNo = rs.Fields("検査番号").Value

    UFMitsubishi.YoutoNo = rs.Fields("用途ｺｰﾄﾞ").Value
    UFMitsubishi.H494_Youto = rs.Fields("用途ｺｰﾄﾞ").Value

    Select Case val(rs.Fields("製造板幅").Value)
    Case 1# To 299#
        UFMitsubishi.H494 = True
        ZenSh.Cells(2, 4).Font.Color = RGB(255, 0, 0)
    End Select

    ZenSh.Cells(3, 2) = Format(val(rs.Fields("製造板厚").Value), "0.000")
    ZenSh.Cells(3, 4) = Format(val(rs.Fields("製造板幅").Value), "0.0")
    ZenSh.Cells(3, 6) = Format(val(rs.Fields("製造板丈").Value), "0.0")

    ' 板厚公差
    Dim dOP As Double: dOP = val(rs.Fields("板厚公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ").Value)
    Dim dOM As Double: dOM = val(rs.Fields("板厚公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ").Value)
    Dim dMP As Double: dMP = val(rs.Fields("板厚公差_製造_ﾌﾟﾗｽ").Value)
    Dim dMM As Double: dMM = val(rs.Fields("板厚公差_製造_ﾏｲﾅｽ").Value)
    If dOP = 0 Or dOM = 0 Then dOP = dMP: dOM = dMM
    ZenSh.Cells(2, 15) = Format(dOP, "0.000")
    ZenSh.Cells(2, 13) = Format(dOM, "0.000")
    ZenSh.Cells(2, 19) = "=R[1]C[-17]+RC[-4]"
    ZenSh.Cells(2, 17) = "=R[1]C[-15]-RC[-4]"

    ' 板幅公差
    Dim wOP As Double: wOP = val(rs.Fields("板幅公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ").Value)
    Dim wOM As Double: wOM = val(rs.Fields("板幅公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ").Value)
    Dim wMP As Double: wMP = val(rs.Fields("板幅公差_製造_ﾌﾟﾗｽ").Value)
    Dim wMM As Double: wMM = val(rs.Fields("板幅公差_製造_ﾏｲﾅｽ").Value)
    If wOP = 0 Or wOM = 0 Then wOP = wMP: wOM = wMM
    ZenSh.Cells(3, 15) = Format(wOP, "0.0")
    ZenSh.Cells(3, 13) = Format(wOM, "0.0")
    ZenSh.Cells(3, 19) = "=RC[-15]+RC[-4]"
    ZenSh.Cells(3, 17) = "=RC[-13]-RC[-4]"

    ' 板丈公差
    Dim tOP As Double: tOP = val(rs.Fields("板丈公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ").Value)
    Dim tOM As Double: tOM = val(rs.Fields("板丈公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ").Value)
    Dim tMP As Double: tMP = val(rs.Fields("板丈公差_製造_ﾌﾟﾗｽ").Value)
    Dim tMM As Double: tMM = val(rs.Fields("板丈公差_製造_ﾏｲﾅｽ").Value)
    If tOP = 0 Or tOM = 0 Then tOP = tMP: tOM = tMM
    ZenSh.Cells(4, 15) = Format(tOP, "0.0")
    ZenSh.Cells(4, 13) = Format(tOM, "0.0")
    ZenSh.Cells(4, 19) = "=R[-1]C[-13]+RC[-4]"
    ZenSh.Cells(4, 17) = "=R[-1]C[-11]-RC[-4]"

    ZenSh.Cells(1, 12) = rs.Fields("品質ｸﾞﾚｰﾄﾞ_ﾗﾃﾗﾙﾎﾞｰ").Value
    ZenSh.Cells(1, 13) = ""
    ZenSh.Cells(1, 18) = rs.Fields("品質ｸﾞﾚｰﾄﾞ_直角度").Value

    ' 対角処理（元コードのまま）
    Dim LotNo As String
    If UFCask.Visible Then LotNo = UFCask.LotBox
    Select Case UFMitsubishi.WorkFlag.Caption
    Case "P1":     LotNo = UFMitsubishi.LotBox
    Case "P_H494": LotNo = UFMitsubishi.H494_Lot
    End Select

    Dim habaT As Single: habaT = PastHiki(3)
    Dim takeT As Single: takeT = PastHiki(4)
    Select Case ZenSh.Cells(1, 18)
    Case "4": ZenSh.Cells(1, 19) = Format(Diagonal(habaT, takeT), "0.0")
    Case "3": ZenSh.Cells(1, 19) = Format(ReDiagonal(habaT, takeT), "0.0")
    End Select

    Dim strShiji As String: strShiji = rs.Fields("指示_直角度").Value
    If strShiji <> "" Then
        Dim kou As Single: kou = Format(ConvertToHalf(strShiji), "0.0")
        ZenSh.Unprotect
        ZenSh.Cells(1, 19) = kou
        ZenSh.Cells(1, 18) = "追記"
    End If

    ' ラテラルボー処理（元コードのまま）
    Dim ラテflag As Boolean: ラテflag = False
    Dim strShijiLate As String: strShijiLate = rs.Fields("指示_ﾗﾃﾗﾙﾎﾞｰ").Value
    If strShijiLate <> "" Then
        ZenSh.Cells(1, 13) = ConvertToHalf(strShijiLate)
        ラテflag = True
    Else
        If ZenSh.Cells(1, 13) = "" Then
            Select Case ZenSh.Cells(1, 12)
            Case "4"
                ZenSh.Cells(1, 13) = Format(Lateralbow_J(ZenSh.Cells(3, 4), ZenSh.Cells(3, 6), ZenSh.Cells(3, 2)), "0.0")
                If ZenSh.Cells(1, 13) = 0 Then ZenSh.Cells(1, 13) = ""
            Case "3"
                ZenSh.Cells(1, 13) = Format(ReLateralbow(ZenSh.Cells(3, 4), ZenSh.Cells(3, 6), ZenSh.Cells(3, 2)), "0.0")
                If ZenSh.Cells(1, 13) = 0 Then ZenSh.Cells(1, 13) = ""
            Case Else
                ZenSh.Cells(1, 13) = Format(Lateralbow_J(ZenSh.Cells(3, 4), ZenSh.Cells(3, 6), ZenSh.Cells(3, 2)), "0.0")
                If ZenSh.Cells(1, 13) = 0 Then ZenSh.Cells(1, 13) = ""
            End Select
        End If
    End If
    If ラテflag Then ZenSh.Cells(1, 12) = "追記"

End If

rs.Close: Set rs = Nothing
db.Close: Set db = Nothing
  
'*************************************************************
    If NoFl = False Then
      ZenSh.Cells(1, 12) = ""
      ZenSh.Cells(1, 13) = ""
      ZenSh.Cells(1, 14) = ""
      ZenSh.Cells(1, 15) = ""
      
      ZenSh.Cells(1, 18) = ""
      ZenSh.Cells(1, 19) = ""
      
      ZenSh.Cells(1, 15) = ""
      ZenSh.Cells(3, 2) = ""
      ZenSh.Cells(3, 4) = ""
      ZenSh.Cells(3, 6) = ""
      ZenSh.Cells(2, 13) = ""
      ZenSh.Cells(2, 15) = ""
      ZenSh.Cells(2, 17) = ""
      ZenSh.Cells(2, 19) = ""
      ZenSh.Cells(3, 13) = ""
      ZenSh.Cells(3, 15) = ""
      ZenSh.Cells(3, 17) = ""
      ZenSh.Cells(3, 19) = ""
      ZenSh.Cells(4, 13) = ""
      ZenSh.Cells(4, 15) = ""
      ZenSh.Cells(4, 17) = ""
      ZenSh.Cells(4, 19) = ""
      '2024.1.16
      ZenSh.Cells(5, 13) = ""
      ZenSh.Cells(5, 15) = ""
      
      UFMitsubishi.KenNo = ""
      If UFMitsubishi.ComboBox1.Tag <> "Load" Then MsgBox "Text.データなし　手入力よろしく"
      Range("Q2").FormulaR1C1 = "=R[1]C[-15]-RC[-4]" '厚
      Range("S2").FormulaR1C1 = "=R[1]C[-17]+RC[-4]"
      Range("Q3").FormulaR1C1 = "=RC[-13]-RC[-4]" '巾
      Range("S3").FormulaR1C1 = "=RC[-15]+RC[-4]"
      Range("Q4").FormulaR1C1 = "=R[-1]C[-11]-RC[-4]" '丈
      Range("S4").FormulaR1C1 = "=R[-1]C[-13]+RC[-4]"
    Exit Sub
    End If
    
Call HistorySave(PastHiki)

'------------------------------------------------------------------------------------------------
Dim LotText As String:   LotText = UFMitsubishi.LotBox.Text
Dim TempArry As Variant: TempArry = GetLotInfo("Lot情報履歴", LotText)
If IsEmpty(TempArry) Then Exit Sub
Dim TempDelivery As String
'納入、取引、用途コード、材質、調質、厚、幅、丈、ロットNo、検査No、ｵｰﾀﾞｰNo、包装仕様NO
Dim 納入 As Variant, 取引 As Variant, 用途コード As String, 材質 As String, 調質 As String
Dim 厚 As Single, 幅 As Single, 丈 As Single, 検番 As String
Dim ArrCount As Long: ArrCount = UBound(TempArry, 1)
ReDim 納入(1 To ArrCount) As Variant
ReDim 取引(1 To ArrCount) As Variant '2024.8.5
j = 0
  For i = LBound(TempArry, 1) To UBound(TempArry, 1) '引当数分チェック....1回でよくない？ ***あかん！滑川は複数引き当ての時に特殊検査のない引き当てと混ざってくる_2024.4.22
    j = j + 1
    納入(j) = TempArry(i, 1)
    取引(j) = TempArry(i, 2)
    用途コード = TempArry(i, 3)
    材質 = TempArry(i, 4)
    調質 = TempArry(i, 5)
    厚 = TempArry(i, 6)
    幅 = TempArry(i, 7)
    丈 = TempArry(i, 8)
    検番 = TempArry(i, 10) '2026.05.08_参照がずっとLotNoだったTempArry(i, 9)　検番参照に修正TempArry(i, 10)
  Next i
  
  Dim 納入Check As String, 取引Check As String
  Dim 前回納入Check As String: 前回納入Check = ""
  For i = LBound(納入, 1) To UBound(納入, 1)
    If 納入(i) <> "" Then
      納入Check = 納入(i)
      取引Check = 取引(i)
      
      '納入先が前回と同じなら DbReturnItem をスキップ
      If 納入Check = 前回納入Check Then GoTo NextLoop
      前回納入Check = 納入Check
      
      TempDelivery = DbReturnItem(納入Check, 取引Check, 用途コード, 材質, 調質, 厚, 幅, 丈)
    End If
NextLoop:
Next i
Dim QpArry As Variant:         If UFMitsubishi.LotBox <> "" Then QpArry = DistArray(UFMitsubishi.LotBox, UFMitsubishi.KenNo, "", "") 'ゲージ、幅歪、丈歪、中、耳、コメント、参照元、番号、フラグ
'Dim QpArryG As Variant:        If UFMitsubishi.LotBox <> "" Then QpArryG = DistArray(UFMitsubishi.LotBox, UFMitsubishi.KenNo, "取引", "", "")
'Dim UpdatedQpArryD As Variant: If UFMitsubishi.LotBox <> "" Then UpdatedQpArryD = UpdateQpArryD(QpArry, QpArryG)

If IsEmptyArray(QpArry) = False Then Call FormReflectDist(QpArry, "") 'フォーム反映

ZenSh.Activate
Application.ScreenUpdating = True '処理高速化のための画面更新停止
Application.EnableEvents = True

If UFMitsubishi.PackNo_on_off <> "off" Then
  If Not PackingNoCheck(LotNo) = "" Then Call UFPack_Open(PackingNoCheck(LotNo), LotNo) '包装仕様書Noチェック
End If

ZenSh.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True
Exit Sub

ErrHandler_MitCombo:
  If Not rs Is Nothing Then
    If rs.state = 1 Then rs.Close
    Set rs = Nothing
  End If
  If Not db Is Nothing Then
    db.Close: Set db = Nothing
  End If
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  ZenSh.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True
  MsgBox "DB読み込みエラー:" & vbCrLf & Err.Description, vbCritical
End Sub

'
'Private Sub ComboBox1_Change()
'If UFMitsubishi.ComboBox1.Tag = "Load" Then Exit Sub
'Application.EnableEvents = False
'Application.ScreenUpdating = False '処理高速化のための画面更新停止
'If WsCheck("全ヵ所検査") = False Then Exit Sub
'If Len(UFMitsubishi.ComboBox1.Text) <> 7 Then Exit Sub
'If UFMitsubishi.ComboBox1.Tag = "Load" Then Exit Sub
'Dim b As String: b = UFMitsubishi.ComboBox1.ListIndex
'Dim i As Long, a As Long, z As Long, k As Long, j As Long
'Dim Order As Long, LOT As Long, Atu As Long, Haba As Long, Take As Long, KenNo As Long, HikiNo As Long
'Dim Zai As Long, Tyou As Long, UseNo As Long, Cours As Long, Good As Long, Deliver As Long, Sendi As Long
'Dim 製厚P As Long, 製厚M As Long, O厚P As Long, O厚M As Long
'Dim 製巾P As Long, 製巾M As Long, O巾P As Long, O巾M As Long
'Dim 製丈P As Long, 製丈M As Long, O丈P As Long, O丈M As Long, 対角 As Long
'Dim 実績枚 As Long, 実績重 As Long, 合紙 As Long, VC As Long, LateR As Long '2021.07.13_ﾗﾃﾗﾙﾎﾞｰ
'Dim 指示歪巾 As Long, 指示歪丈 As Long, 指示ラテ As Long, 指示対角 As Long '2022.07
'
'Dim ZenSh As Worksheet
'If UFMitsubishi.WorkFlag.Caption = "P1" Then Set ZenSh = ThisWorkbook.Sheets("全ヵ所検査")
'If UFMitsubishi.WorkFlag.Caption = "P_H494" Then Set ZenSh = ThisWorkbook.Sheets("H494検査")
'
'ZenSh.Unprotect 'シート保護解除
'ZenSh.Cells(1, 15) = ""
'ZenSh.Cells(2, 4).Font.Color = RGB(0, 0, 0)
'
'
'Dim HiAcc: HiAcc = Arrbox("SIKALOTNOW.accdb", "板", "Fields", "")
'If IsEmpty(HiAcc) Then Exit Sub
'  For i = 1 To UBound(HiAcc, 2) '製造'ｵｰﾀﾞｰ公差確認後片方でも0の場合は製造公差を見ることとする
'    Select Case HiAcc(1, i)
'    Case "ｵｰﾀﾞｰ番号":           Order = i
'    Case "ﾛｯﾄ番号":             LOT = i
'    Case "製造板厚":            Atu = i
'    Case "製造板幅":            Haba = i
'    Case "製造板丈":            Take = i
'    Case "検査番号":            KenNo = i
'    Case "製造材質":            Zai = i
'    Case "製造調質":            Tyou = i
'    Case "用途ｺｰﾄﾞ":            UseNo = i
'    Case "設計_設備ｺｰｽ":        Cours = i
'    Case "取引先名":            Good = i
'    Case "納入先名":            Deliver = i
'    Case "送り先名":            Sendi = i
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
'    Case "品質ｸﾞﾚｰﾄﾞ_ﾗﾃﾗﾙﾎﾞｰ":  LateR = i
'    Case "指示_ﾗﾃﾗﾙﾎﾞｰ":        指示ラテ = i
'    Case "指示_直角度":         指示対角 = i
'    End Select
'  Next i
'  Dim PastHiki(0 To 18) '2020.07.22******2021.07.13_ラテラルボー追加
'
'  Dim NoFl As Boolean: NoFl = False
'  Dim LotAcc: LotAcc = Arrbox("SIKALOTNOW.accdb", "板", "Record", "")
'
'  For i = 2 To UBound(LotAcc, 1)
'    If LotAcc(i, LOT) = Me.ComboBox1.Text Then 'LOTNO一致がある
'
'         NoFl = True
'
'           PastHiki(0) = "'" & LotAcc(i, Order)
'           PastHiki(1) = LotAcc(i, LOT)
'           PastHiki(2) = LotAcc(i, Atu)
'           PastHiki(3) = LotAcc(i, Haba)
'           PastHiki(4) = LotAcc(i, Take)
'           PastHiki(5) = LotAcc(i, KenNo) '検番
'           PastHiki(6) = LotAcc(i, Zai)
'           PastHiki(7) = LotAcc(i, Tyou)
'           PastHiki(8) = LotAcc(i, UseNo) '用途コード
'           PastHiki(9) = LotAcc(i, Cours) 'ｺｰｽ
'           PastHiki(10) = LotAcc(i, Good) '得意先
'           PastHiki(11) = LotAcc(i, Deliver) '納入先
'           PastHiki(12) = LotAcc(i, Sendi) '送り先
'           PastHiki(13) = LotAcc(i, 実績枚) '実績枚
'           PastHiki(14) = LotAcc(i, 実績重) '実績重
'           PastHiki(15) = LotAcc(i, 合紙) '合紙
'           PastHiki(16) = LotAcc(i, VC) 'VC
'           PastHiki(17) = LotAcc(i, LateR) 'ﾗﾃﾗﾙﾎﾞｰ
'
'         ZenSh.Cells(2, 4) = Me.ComboBox1.Text
'         UFMitsubishi.KenNo = LotAcc(i, KenNo) '検査番号
'         UFMitsubishi.H494_KenNo = LotAcc(i, KenNo)
'
'         '2023.05.16_追加***検査頻度不明のためこちらにも追加
'         UFMitsubishi.YoutoNo = LotAcc(i, UseNo) '用途コード
'         UFMitsubishi.H494_Youto = LotAcc(i, UseNo)
'         Select Case val(LotAcc(i, Haba)) '幅寸法が範囲内の時
'         Case 1# To 299#
'          UFMitsubishi.H494 = True
'          ZenSh.Cells(2, 4).Font.Color = RGB(255, 0, 0)
'         End Select
'
'         ZenSh.Cells(3, 2) = Format(LotAcc(i, Atu), "0.000") '厚　製造サイズとする
'         ZenSh.Cells(3, 4) = Format(LotAcc(i, Haba), "0.0")  '巾
'         ZenSh.Cells(3, 6) = Format(LotAcc(i, Take), "0.0") '丈
'        '*************************************************************厚公差
'         ZenSh.Cells(2, 15) = Format(LotAcc(i, O厚P), "0.000") '+　　マイナスを左側
'         ZenSh.Cells(2, 13) = Format(LotAcc(i, O厚M), "0.000") '-
'           If (ZenSh.Cells(2, 13) = "0" Or ZenSh.Cells(2, 15) = "0") Or ZenSh.Cells(2, 13) = "" Or ZenSh.Cells(2, 15) = "" Then '片方でも0の場合は製造公差を見る
'           ZenSh.Cells(2, 15) = Format(LotAcc(i, 製厚P), "0.000") '+　　マイナスを左側
'           ZenSh.Cells(2, 13) = Format(LotAcc(i, 製厚M), "0.000") '-
'           End If
'        '--------------------------厚公差計算後
'         ZenSh.Cells(2, 19) = "=R[1]C[-17]+RC[-4]" '+' Format(Val(.Cells(3, 2)) + Val(.Cells(2, 15)), "0.000") 'マイナスを左側
'         ZenSh.Cells(2, 17) = "=R[1]C[-15]-RC[-4]" '-'Format(Val(.Cells(3, 2)) - Val(.Cells(2, 13)), "0.000") '- 厚
'        '*************************************************************巾公差
'          ZenSh.Cells(3, 15) = Format(LotAcc(i, O巾P), "0.0") '+
'          ZenSh.Cells(3, 13) = Format(LotAcc(i, O巾M), "0.0") '-
'         If ZenSh.Cells(3, 13) = "" Or ZenSh.Cells(3, 13) = "0" Or ZenSh.Cells(3, 15) = "" Or ZenSh.Cells(3, 15) = "0" Then '片方でも0の場合は製造公差を見る
'          ZenSh.Cells(3, 15) = Format(LotAcc(i, 製巾P), "0.0") '+
'          ZenSh.Cells(3, 13) = Format(LotAcc(i, 製巾M), "0.0") '-
'         End If
'        '--------------------------巾公差計算後
'         ZenSh.Cells(3, 19) = "=RC[-15]+RC[-4]" '+'Format(Val(.Cells(3, 4)) + Val(.Cells(3, 15)), "0.0") 'マイナスを左側
'         ZenSh.Cells(3, 17) = "=RC[-13]-RC[-4]" '-'Format(Val(.Cells(3, 4)) - Val(.Cells(3, 13)), "0.0") '- 巾
'        '*************************************************************丈公差
'          ZenSh.Cells(4, 15) = Format(LotAcc(i, O丈P), "0.0") '+
'          ZenSh.Cells(4, 13) = Format(LotAcc(i, O丈M), "0.0") '-
'         If ZenSh.Cells(4, 13) = "" Or ZenSh.Cells(4, 13) = "0" Or ZenSh.Cells(4, 15) = "" Or ZenSh.Cells(4, 15) = "0" Then '片方でも0の場合は製造公差を見る
'          ZenSh.Cells(4, 15) = Format(LotAcc(i, 製丈P), "0.0") '+
'          ZenSh.Cells(4, 13) = Format(LotAcc(i, 製丈M), "0.0") '-
'         End If
'        '--------------------------丈公差計算後
'         ZenSh.Cells(4, 19) = "=R[-1]C[-13]+RC[-4]" '+'Format(Val(.Cells(3, 6)) + Val(.Cells(4, 15)), "0.0") 'マイナスを左側
'         ZenSh.Cells(4, 17) = "=R[-1]C[-11]-RC[-4]" '-' Format(Val(.Cells(3, 6)) - Val(.Cells(4, 13)), "0.0") '- 丈
'
'
'          ZenSh.Cells(1, 12) = LotAcc(i, LateR)  'ﾗﾃﾗﾙﾎﾞｰ'2023.4.13_追加
'          ZenSh.Cells(1, 13) = ""
'          ZenSh.Cells(1, 18) = LotAcc(i, 対角)   '対角'2023.6.20_追加
'
'         '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>対角
'          '追加公差対策***2023.4.12
'            Dim LotNo As String
'            If UFCask.Visible Then LotNo = UFCask.LotBox
'
'            Select Case UFMitsubishi.WorkFlag.Caption
'            Case "P1"
'              LotNo = UFMitsubishi.LotBox
'            Case "P_H494"
'              LotNo = UFMitsubishi.H494_Lot
'            End Select
'
'            Dim habaT As Single: habaT = PastHiki(3)
'            Dim takeT As Single: takeT = PastHiki(4)
'            Select Case ZenSh.Cells(1, 18)
'            Case "4"
'              ZenSh.Cells(1, 19) = Format(Diagonal(habaT, takeT), "0.0")  '4級(厳しい)
'            Case "3"
'              ZenSh.Cells(1, 19) = Format(ReDiagonal(habaT, takeT), "0.0")
'            End Select
'
'            If IsEmptyArray(LotAcc) Then GoTo skiplate '2022.03.14_空ファイル対策
'
'            Select Case True
'            Case LotAcc(i, LOT) = LotNo 'LOTNO一致がある　Lot情報履歴用配列
'              If LotAcc(i, 指示対角) <> "" Then '技術指示がある場合を追加**2022.07
'                '追記公差分
'                  Dim kou As Single: kou = Format(ConvertToHalf(LotAcc(i, 指示対角)), "0.0") '数値のみを半角へ変換後スラッシュ前の数値を抜き取る'専用関数
'                  '公差追加表示_2023.4.13
'                  ZenSh.Unprotect 'シート保護解除
'                  ZenSh.Cells(1, 19) = kou
'                  ZenSh.Cells(1, 18) = "追記"
'              End If
'            End Select
'
'skiplate:
'         '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>ﾗﾃﾗﾙﾎﾞｰ
'         '公差に関してはフォーマット関数を使用せずそのままの桁数で表示させることとする
'         Dim 対flag As Boolean:   対flag = False
'         Dim ラテflag As Boolean: ラテflag = False
'If 指示ラテ = 0 Then GoTo LineSkip2
'            If LotAcc(i, 指示ラテ) <> "" Then '技術指示がある場合を追加**2022.07
'            '追記公差分
'                ZenSh.Cells(1, 13) = ConvertToHalf(LotAcc(i, 指示ラテ))
'                ラテflag = True
'            ElseIf LotAcc(i, 指示ラテ) = "" Then
'LineSkip2:
'                '追記公差なし
'                If ZenSh.Cells(1, 13) = "" Then '追記公差がなければ公差記入がないので等級から公差を割り出し記入
'                  Select Case ZenSh.Cells(1, 12)
'                  Case "4"
'                    ZenSh.Cells(1, 13) = Format(Lateralbow_J(ZenSh.Cells(3, 4), ZenSh.Cells(3, 6), ZenSh.Cells(3, 2)), "0.0")  '4級(厳しい)
'                   If ZenSh.Cells(1, 13) = 0 Then ZenSh.Cells(1, 13) = ""
'                  Case "3"
'                   ZenSh.Cells(1, 13) = Format(ReLateralbow(ZenSh.Cells(3, 4), ZenSh.Cells(3, 6), ZenSh.Cells(3, 2)), "0.0")
'                   If ZenSh.Cells(1, 13) = 0 Then ZenSh.Cells(1, 13) = ""
'                  Case Else '数値がない場合
'                    ZenSh.Cells(1, 13) = Format(Lateralbow_J(ZenSh.Cells(3, 4), ZenSh.Cells(3, 6), ZenSh.Cells(3, 2)), "0.0")  '4級(厳しい)
'                   If ZenSh.Cells(1, 13) = 0 Then ZenSh.Cells(1, 13) = ""
'                  End Select
'                End If
'            End If
'
'          If ラテflag Then ZenSh.Cells(1, 12) = "追記"
'
''          '白銅追加*****2022.08.24
''          '**************納入が白銅の時ﾗﾃﾗﾙﾎﾞｰ公差を書き換える**************
''          If LotAcc(i, Deliver) Like "ﾊｸﾄﾞｳ*" Then '白銅を含む納入先
''            Select Case LotAcc(i, ATU) '指定の板厚**指示書_QA-022-03
''            Case 0.3 To 302.5
''              If LotAcc(i, ZAI) = "995" Or LotAcc(i, ZAI) = "NP56" Or LotAcc(i, ZAI) Like "*52S*" Then
''                  ZenSh.Cells(1, 13) = Format(2, "0.0")
''                  ZenSh.Cells(1, 15) = "白銅" & LotAcc(i, ZAI)
''              ElseIf LotAcc(i, ZAI) Like "*17S*" Or LotAcc(i, ZAI) Like "*61S*" Then
''                  ZenSh.Cells(1, 13) = Format(3, "0.0")
''                  ZenSh.Cells(1, 15) = "白銅" & LotAcc(i, ZAI)
''              End If
''            End Select
''          End If
''          '日本車両追加*****2023.09.05
''          '**************納入が日本車両の時ﾗﾃﾗﾙﾎﾞｰ公差を書き換える**************
''          Dim LotNoN As String: LotNoN = PastHiki(1)
''          If LotAcc(i, Deliver) Like "*ﾆﾎﾝｼﾔﾘﾖｳｾｲｿﾞｳ(ｶ*" Then '日車を含む納入先
''            ZenSh.Cells(1, 13) = Format(DBdisplayラテラル(LotNoN), "0.0")
''            ZenSh.Cells(1, 15) = "日本車両" & LotAcc(i, ZAI)
''          End If
''          '滑川-Q-S-QA-DIM-179_板材曲がり(ラテラルボー)検査基準 発行日2024年3月5日
''          '**************納入が滑川の時ﾗﾃﾗﾙﾎﾞｰ公差を書き換える**************
''          If LotAcc(i, Deliver) Like "ﾅﾒｶﾜｱﾙﾐ*" Then
''          Select Case LotAcc(i, ATU)
''          Case 0.3 To 302.5
''            If LotAcc(i, ZAI) = "NP56" Or LotAcc(i, ZAI) Like "*52S*" Or LotAcc(i, ZAI) Like "*995" Then
''                ZenSh.Cells(1, 13) = Format(2, "0.0")
''                ZenSh.Cells(1, 15) = "滑川" & LotAcc(i, ZAI)
''            ElseIf LotAcc(i, ZAI) Like "*17S*" Or LotAcc(i, ZAI) Like "*61S*" Then
''                ZenSh.Cells(1, 13) = Format(3, "0.0")
''                ZenSh.Cells(1, 15) = "滑川" & LotAcc(i, ZAI)
''            End If
''          End Select
''        End If
'    End If
'  Next i
'
'     '指定サイズのみ先行して記入欄を埋める　2020.01.09　*2019.12~2020.08納入分まで
'      On Error Resume Next
''      If (Val(ZenSh.Cells(3, 2)) = 20# And Val(ZenSh.Cells(3, 4)) = 580# And Val(ZenSh.Cells(3, 6)) = 2220#) Then
''      For i = 7 To 13 Step 2 'ｸﾘｱ
''       ZenSh.Cells(i, 27) = ""
''       ZenSh.Cells(i, 29) = ""
''      Next i
''       ZenSh.Cells(7, 27) = 2
''       ZenSh.Cells(7, 29) = 2
''       ZenSh.Cells(9, 27) = 1
''       ZenSh.Cells(9, 29) = 4
''      MsgBox "三菱受託対象Size", vbExclamation
''      ElseIf (Val(ZenSh.Cells(3, 2)) = 20# And Val(ZenSh.Cells(3, 4)) = 720# And Val(ZenSh.Cells(3, 6)) = 2220#) Then
''      For i = 7 To 13 Step 2 'ｸﾘｱ
''       ZenSh.Cells(i, 27) = ""
''       ZenSh.Cells(i, 29) = ""
''      Next i
''       ZenSh.Cells(7, 27) = 2
''       ZenSh.Cells(7, 29) = 1
''       ZenSh.Cells(9, 27) = 1
''       ZenSh.Cells(9, 29) = 3
''      MsgBox "三菱受託対象Size", vbExclamation
''      ElseIf (Val(ZenSh.Cells(3, 2)) = 12# And Val(ZenSh.Cells(3, 4)) = 640# And Val(ZenSh.Cells(3, 6)) = 2150#) _
''      Or (Val(ZenSh.Cells(3, 2)) = 20# And Val(ZenSh.Cells(3, 4)) = 650# And Val(ZenSh.Cells(3, 6)) = 2220#) Then
''      For i = 7 To 13 Step 2 'ｸﾘｱ
''       ZenSh.Cells(i, 27) = ""
''       ZenSh.Cells(i, 29) = ""
''      Next i
''      MsgBox "三菱受託対象Size" & vbCrLf & "ﾌﾞﾛｯｸNOの記入をする事", vbExclamation
''      End If
''*************************************************************
'    If NoFl = False Then
'      ZenSh.Cells(1, 12) = ""
'      ZenSh.Cells(1, 13) = ""
'      ZenSh.Cells(1, 14) = ""
'      ZenSh.Cells(1, 15) = ""
'
'      ZenSh.Cells(1, 18) = ""
'      ZenSh.Cells(1, 19) = ""
'
'      ZenSh.Cells(1, 15) = ""
'      ZenSh.Cells(3, 2) = ""
'      ZenSh.Cells(3, 4) = ""
'      ZenSh.Cells(3, 6) = ""
'      ZenSh.Cells(2, 13) = ""
'      ZenSh.Cells(2, 15) = ""
'      ZenSh.Cells(2, 17) = ""
'      ZenSh.Cells(2, 19) = ""
'      ZenSh.Cells(3, 13) = ""
'      ZenSh.Cells(3, 15) = ""
'      ZenSh.Cells(3, 17) = ""
'      ZenSh.Cells(3, 19) = ""
'      ZenSh.Cells(4, 13) = ""
'      ZenSh.Cells(4, 15) = ""
'      ZenSh.Cells(4, 17) = ""
'      ZenSh.Cells(4, 19) = ""
'      '2024.1.16
'      ZenSh.Cells(5, 13) = ""
'      ZenSh.Cells(5, 15) = ""
'
'      UFMitsubishi.KenNo = ""
'      If UFMitsubishi.ComboBox1.Tag <> "Load" Then MsgBox "Text.データなし　手入力よろしく"
'      Range("Q2").FormulaR1C1 = "=R[1]C[-15]-RC[-4]" '厚
'      Range("S2").FormulaR1C1 = "=R[1]C[-17]+RC[-4]"
'      Range("Q3").FormulaR1C1 = "=RC[-13]-RC[-4]" '巾
'      Range("S3").FormulaR1C1 = "=RC[-15]+RC[-4]"
'      Range("Q4").FormulaR1C1 = "=R[-1]C[-11]-RC[-4]" '丈
'      Range("S4").FormulaR1C1 = "=R[-1]C[-13]+RC[-4]"
'    Exit Sub
'    End If
'
'
'Call HistorySave(PastHiki)
'
''------------------------------------------------------------------------------------------------
'Dim LotText As String: LotText = UFMitsubishi.LotBox.Text
'Dim TempArry As Variant: TempArry = GetLotInfo("Lot情報履歴", LotText)
'If IsEmpty(TempArry) Then Exit Sub
'Dim TempDelivery As String
''納入、取引、用途コード、材質、調質、厚、幅、丈、ロットNo、検査No、ｵｰﾀﾞｰNo、包装仕様NO
'Dim 納入 As Variant, 取引 As Variant, 用途コード As String, 材質 As String, 調質 As String
'Dim 厚 As Single, 幅 As Single, 丈 As Single, 検番 As String
'Dim ArrCount As Long: ArrCount = UBound(TempArry, 1)
'ReDim 納入(1 To ArrCount) As Variant
'ReDim 取引(1 To ArrCount) As Variant '2024.8.5
'j = 0
'  For i = LBound(TempArry, 1) To UBound(TempArry, 1) '引当数分チェック....1回でよくない？ ***あかん！滑川は複数引き当ての時に特殊検査のない引き当てと混ざってくる_2024.4.22
'    j = j + 1
'    納入(j) = TempArry(i, 1)
'    取引(j) = TempArry(i, 2)
'    用途コード = TempArry(i, 3)
'    材質 = TempArry(i, 4)
'    調質 = TempArry(i, 5)
'    厚 = TempArry(i, 6)
'    幅 = TempArry(i, 7)
'    丈 = TempArry(i, 8)
'    検番 = TempArry(i, 10) '2026.05.08_参照がずっとLotNoだったTempArry(i, 9)　検番参照に修正TempArry(i, 10)
'  Next i
'
'  Dim 納入Check As String, 取引Check As String
'  For i = LBound(納入, 1) To UBound(納入, 1)
'    If 納入(i) <> "" Then
'      納入Check = 納入(i)
'      取引Check = 取引(i)
'      TempDelivery = DbReturnItem(納入Check, 取引Check, 用途コード, 材質, 調質, 厚, 幅, 丈)
'    End If
'Next i
'Dim QpArry As Variant:         If UFMitsubishi.LotBox <> "" Then QpArry = DistArray(UFMitsubishi.LotBox, UFMitsubishi.KenNo, "", "") 'ゲージ、幅歪、丈歪、中、耳、コメント、参照元、番号、フラグ
''Dim QpArryG As Variant:        If UFMitsubishi.LotBox <> "" Then QpArryG = DistArray(UFMitsubishi.LotBox, UFMitsubishi.KenNo, "取引", "", "")
''Dim UpdatedQpArryD As Variant: If UFMitsubishi.LotBox <> "" Then UpdatedQpArryD = UpdateQpArryD(QpArry, QpArryG)
'
'If IsEmptyArray(QpArry) = False Then Call FormReflectDist(QpArry, "") 'フォーム反映
'
'ZenSh.Activate
'Application.ScreenUpdating = True '処理高速化のための画面更新停止
'Application.EnableEvents = True
'
'If UFMitsubishi.PackNo_on_off <> "off" Then
'  If Not PackingNoCheck(LotNo) = "" Then Call UFPack_Open(PackingNoCheck(LotNo), LotNo) '包装仕様書Noチェック
'End If
'ZenSh.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True
'End Sub

Private Sub CommandButton1_Click()
Dim Lname As String: Lname = UFMitsubishi.LineName.Caption
If UFMitsubishi.Tag = "off" Then Exit Sub
If Register_Check(Format$(Now, "hh:mm:ss")) Then Call Daily_Create '日報作成
If MsgBox("新規発行しますか？現在のｼｰﾄは削除されます。", vbQuestion + vbYesNo) = vbYes Then
  UFProgress.Show vbModeless
  UFProgress.CommandButton1.Caption = "NewDataCreate...."
  Call ProgressBar_count(0, 350)
  DoEvents
  Call JustinCase_ZnSave
  Call ProgressBar_count(350, 560)
  DoEvents
  Call 全ヵ所用コピー
  Select Case UFMitsubishi.LineName.Caption
  Case "L-1":    Call WriteLog("*UFKen:" & "三菱受託検査ｼｰﾄ発行", "", HistoryPath & "L-1\L-1_Log")
  Case "LVC":    Call WriteLog("*UFKen:" & "三菱受託検査ｼｰﾄ発行", "", HistoryPath & "LVC\LVC_Log")
  Case "HVC":    Call WriteLog("*UFKen:" & "三菱受託検査ｼｰﾄ発行", "", HistoryPath & "HVC\HVC_Log")
  Case "ﾊﾞﾗﾝｻｰ": Call WriteLog("*UFKen:" & "三菱受託検査ｼｰﾄ発行", "", HistoryPath & "バランサー\バランサー_Log")
  Case "ﾄｯﾄ":    Call WriteLog("*UFKen:" & "三菱受託検査ｼｰﾄ発行", "", HistoryPath & "トット\トット_Log")
  Case "Test":   Call WriteLog("*UFKen:" & "三菱受託検査ｼｰﾄ発行", "", HistoryPath & "Test\Test_Log")
  Case Else:     Call WriteLog("*UFKen:" & "三菱受託検査ｼｰﾄ発行", "", HistoryPath & "丸徳\丸徳_Log")
  End Select
  
  If WsCheck("不良要因印刷" & "_" & Lname) Then ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Delete '2023.10.25_追加
  
  UFMitsubishi.Work = "ON"
  UFMitsubishi.LotBox = ""
  UFMitsubishi.KenNo = ""
  UFMitsubishi.ComboBox1 = ""
  UFMitsubishi.YoutoNo = ""
  Call ProgressBar_count(560, 1000)
  DoEvents
  Unload UFProgress
Else
  Unload UFProgress
  MsgBox "ｷｬﾝｾﾙ処理", vbInformation
End If
End Sub

Private Sub CommandButton10_Click()
Call FiveN専用保存
'Application.ScreenUpdating = False
'Call FiveN材転送
'Call 原紙印刷範囲
'ThisWorkbook.Sheets("5N").Visible = True '保存用　印刷用_Test 不良要因印刷_Test
'Application.ScreenUpdating = True
End Sub

Private Sub CommandButton14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim ws As Worksheet
If WsCheck("全ヵ所検査") = False Then Exit Sub
Set ws = ThisWorkbook.Worksheets("全ヵ所検査")
ws.Activate
With UFMitsubishi
Select Case UFMitsubishi.Choice.Text '切替
Case "一般"
'受託品作業
  .Choice = "三菱"
  .Work = "ON" '三菱受託用フラグ
  .Inflag = "flagon" '三菱印刷用フラグ
  .Normal.Visible = False
  .Mitsubishi.Visible = True
  .FiNi.Visible = False
  ws.Unprotect '保護解除
'  Range(ws.Cells(7, 27), ws.Cells(7, 29)).Locked = False 'ロック設定
'  Range(ws.Cells(9, 27), ws.Cells(9, 29)).Locked = False
'  Range(ws.Cells(11, 27), ws.Cells(11, 29)).Locked = False
'  Range(ws.Cells(13, 27), ws.Cells(13, 29)).Locked = False
'  Union(Range(ws.Cells(7, 27), ws.Cells(7, 29)), ws.Range(ws.Cells(9, 27), ws.Cells(9, 29)), _
'  Range(ws.Cells(11, 27), ws.Cells(11, 29)), Range(ws.Cells(13, 27), ws.Cells(13, 29))).Interior.Color = RGB(255, 255, 0)
'  Range(ws.Cells(6, 30), ws.Cells(14, 30)).Font.Color = RGB(0, 0, 0)
'  Range(ws.Cells(6, 31), ws.Cells(11, 31)).Font.Color = RGB(255, 255, 255)
  ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
Case "三菱"
'5N作業
  .Choice = "5N"
  .Work = "OFF"
  .Inflag = ""
  .Normal.Visible = False
  .Mitsubishi.Visible = False
  .FiNi.Visible = True
  ws.Unprotect '保護解除
'  Range(ws.Cells(7, 27), Cells(7, 29)).Locked = True 'ロック設定
'  Range(ws.Cells(9, 27), Cells(9, 29)).Locked = True
'  Range(ws.Cells(11, 27), ws.Cells(11, 29)).Locked = True
'  Range(ws.Cells(13, 27), ws.Cells(13, 29)).Locked = True
'  Union(Range(ws.Cells(7, 27), ws.Cells(7, 29)), Range(ws.Cells(9, 27), ws.Cells(9, 29)), _
'  Range(ws.Cells(11, 27), ws.Cells(11, 29)), Range(ws.Cells(13, 27), ws.Cells(13, 29))).Interior.Color = RGB(192, 192, 192)
'  Range(ws.Cells(6, 30), ws.Cells(14, 30)).Font.Color = RGB(255, 255, 255)
'  Range(ws.Cells(6, 31), ws.Cells(11, 31)).Font.Color = RGB(0, 0, 0)
  ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
Case "5N"
'一般作業
  .Choice = "一般"
  .Work = "OFF"
  .Inflag = ""
  .Normal.Visible = True
  .Mitsubishi.Visible = False
  .FiNi.Visible = False
  ws.Unprotect '保護解除
'  Range(ws.Cells(7, 27), Cells(7, 29)).Locked = True 'ロック設定
'  Range(ws.Cells(9, 27), Cells(9, 29)).Locked = True
'  Range(ws.Cells(11, 27), ws.Cells(11, 29)).Locked = True
'  Range(ws.Cells(13, 27), ws.Cells(13, 29)).Locked = True
'  Union(Range(ws.Cells(7, 27), ws.Cells(7, 29)), Range(ws.Cells(9, 27), ws.Cells(9, 29)), _
'  Range(ws.Cells(11, 27), ws.Cells(11, 29)), Range(ws.Cells(13, 27), ws.Cells(13, 29))).Interior.Color = RGB(192, 192, 192)
'  Range(ws.Cells(6, 30), ws.Cells(14, 30)).Font.Color = RGB(255, 255, 255)
'  Range(ws.Cells(6, 31), ws.Cells(11, 31)).Font.Color = RGB(255, 255, 255)
  ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
End Select
End With
End Sub

Private Sub CommandButton15_Click()
Dim ws As Worksheet
Dim FL As Boolean, flag As Boolean
Dim n As Long, i As Long
Dim Lname As String: Lname = UFMitsubishi.LineName.Caption
'注意表示---------------------------------------------------------------------------------------------
If WsCheck("全ヵ所検査") = False Then Exit Sub
Dim result1 As Long, result2 As Long, myRange1 As Range, myRange2 As Range, myRange3 As Range, myRange4 As Range, WSZ As Worksheet: Set WSZ = ThisWorkbook.Worksheets("全ヵ所検査")
Set myRange1 = Range(WSZ.Cells(7, 2), WSZ.Cells(76, 9))
Set myRange2 = Range(WSZ.Cells(7, 25), WSZ.Cells(76, 28))
Set myRange3 = (WSZ.Cells(2, 30))
Set myRange4 = (WSZ.Cells(3, 30))

If myRange3 = "" And myRange4 = "" Then
  MsgBox "検入高さを入力してください。", vbCritical
  Exit Sub
End If

result1 = CheckRedXCells(myRange1)
result2 = CheckRedXCells(myRange2)
If result1 > 0 And result2 > 0 Then
  If MsgBox("公差外判定の板があります " & "厚:" & result1 & "箇所　" & "歪:" & result2 & "箇所" & vbCrLf & "処理を続行しますか？", vbQuestion + vbYesNo) = vbNo Then Exit Sub
ElseIf result1 > 0 And result2 = 0 Then
  If MsgBox("公差外判定の板があります " & "厚:" & result1 & "箇所　" & vbCrLf & "処理を続行しますか？", vbQuestion + vbYesNo) = vbNo Then Exit Sub
ElseIf result1 = 0 And result2 > 0 Then
  If MsgBox("公差外判定の板があります " & "歪:" & result2 & "箇所" & vbCrLf & "処理を続行しますか？", vbQuestion + vbYesNo) = vbNo Then Exit Sub
End If
'------------------------------------------------------------------------------------------------------
If Not UFMitsubishi.Choice = "5N" Then Exit Sub '5N用フラグ
ThisWorkbook.Worksheets("全ヵ所検査").Activate
If UFMitsubishi.Kname = "" And UFMitsubishi.KenNo = "" Then
  MsgBox "検査員名と検査番号を入力してください"
  Exit Sub
End If
UFMitsubishi.Tag = "off"
Call MusicTrue("おまちください.wav")
  UFProgress.Show vbModeless
  Call ProgressBar_count(0, 250)
  DoEvents
'  Call FiveN材転送
  
  Call 全ヵ所用転送
  
  Application.ScreenUpdating = False
  With ActiveSheet.PageSetup
      .LeftHeader = ""
      .CenterHeader = ""
      .RightHeader = ""
      .LeftFooter = ""
      .CenterFooter = ""
      .LeftFooter = ""
      .CenterFooter = ""
      .RightFooter = ""
  End With
'    Call 原紙印刷範囲 'Sheets("転送済").PageSetup.PrintArea = "$B$2:$Q$20"
'    Call MusicTrue("印刷.wav")
'    If MsgBox("印刷しますか？", vbQuestion + vbYesNo) = vbYes Then
'        ThisWorkbook.Sheets("5N").Visible = True
'        ThisWorkbook.Sheets("5N").PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False

'************************歪公差外流出対策_2024.1.11**********************************

    Call 全ヵ所用印刷範囲 'Sheets("全ヵ所検査").PageSetup.PrintArea = "$A$6:$AA$" & n
    Call MusicTrue("印刷.wav")
    If MsgBox("印刷しますか？", vbQuestion + vbYesNo) = vbYes Then
        ThisWorkbook.Sheets("全ヵ所検査").Visible = True
        ThisWorkbook.Sheets("全ヵ所検査").PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
        Call ProgressBar_count(250, 485)
        '-------------------------------------------------------------------不良要因
        If (WsCheck("不良要因印刷" & "_" & Lname) = False) Then 'シートなし
          Call AddWS("不良要因印刷" & "_" & Lname)
          Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
        ElseIf ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Cells(1, 2) <> UFMitsubishi.LotBox Then '不良の入力がなければこの状態
          Call AddWS("不良要因印刷" & "_" & Lname)
          Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname)
        End If
        Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname)
        
        Call 不良要因印刷範囲(UFMitsubishi.Kname)
        
        sh.Visible = True
        sh.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
        sh.Visible = False
        '-------------------------------------------------------------------
        
        Call FiveN専用保存
        Call ProgressBar_count(485, 780)
        DoEvents
          Application.DisplayAlerts = False '保存後シートは消す
            If WsCheck("ｺﾋﾟ-用") Then ThisWorkbook.Sheets("ｺﾋﾟ-用").Delete
            If WsCheck("5N") Then ThisWorkbook.Sheets("5N").Delete
            If WsCheck("全ヵ所検査保存") Then ThisWorkbook.Sheets("全ヵ所検査保存").Delete
            If WsCheck("全ヵ所作業中保存") Then ThisWorkbook.Sheets("全ヵ所作業中保存").Delete
          Application.DisplayAlerts = True
    '----------------------------------------------履歴を残す　2019.09.13　これ以前のものは履歴なし
        If WsCheck("保存用") = False Then Exit Sub
          With ThisWorkbook.Sheets("保存用")
            n = .Cells(Rows.Count, 1).End(xlUp).row + 1
            .Unprotect
            .Cells.Locked = False
            .Cells.FormulaHidden = False
            .Cells(n, 2) = ThisWorkbook.Sheets("全ヵ所検査").Cells(2, 4) 'ロット
            .Cells(n, 83) = "全ヵ所"
            .Cells(n, 3) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 2) '厚
            .Cells(n, 4) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 4) '幅
            .Cells(n, 5) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 6) '丈
            .Cells(n, 1) = Format$(ThisWorkbook.Sheets("全ヵ所検査").Cells(1, 2), "yyyy/mm/dd  h:mm") & "～" & Format$(Now, "yyyy/mm/dd  h:mm")
            .Cells(n, 77) = "別ファイル参照" '検査員
          End With
'        If WsCheck("保存用") = False Then Exit Sub
'          With ThisWorkbook.Sheets("保存用")
'            n = .Cells(Rows.Count, 1).End(xlUp).Row + 1
'            .Unprotect
'            .Cells.Locked = False
'            .Cells.FormulaHidden = False
'            .Cells(n, 2) = ThisWorkbook.Sheets("全ヵ所検査").Cells(2, 4) 'ロット
'            .Cells(n, 83) = "5N"
'            .Cells(n, 3) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 2) '厚
'            .Cells(n, 4) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 4) '幅
'            .Cells(n, 5) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 6) '丈
'            .Cells(n, 1) = Format$(ThisWorkbook.Sheets("全ヵ所検査").Cells(1, 2), "yyyy/mm/dd  h:mm") & "～" & Format$(Now, "yyyy/mm/dd  h:mm")
'            .Cells(n, 77) = "別ファイル参照" '検査員
'          End With
    Call ProgressBar_count(780, 1000)
    Call MusicTrue("完了.wav")
    DoEvents
    Unload UFProgress
    Else
      Unload UFProgress
      MsgBox "ｷｬﾝｾﾙ処理", vbInformation
    End If
Application.ScreenUpdating = True
Application.WindowState = xlMaximized
UFMitsubishi.Tag = ""
End Sub

Private Sub CommandButton16_Click()
Application.EnableEvents = True
Application.Visible = True 'Excel表示
End Sub

Private Sub CommandButton17_Click()
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("全ヵ所検査")
ws.Unprotect

ws.Range("Y:Y,AA:AA,AD:AD").Font.Size = 8
ws.Range("Y:Y, AA:AA, AD:AD").EntireColumn.AutoFit

ActiveWindow.Zoom = 100
ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'シート保護
End Sub

Private Sub CommandButton2_Click()
Dim shName As String, sh As String
Dim ws As Worksheet
Dim FL As Boolean, flag As Boolean
Dim n As Long, i As Long
UFMitsubishi.Tag = "off"
Call MusicTrue("おまちください.wav")
ThisWorkbook.Worksheets("全ヵ所検査").Activate
UFMitsubishi.Work = "ON" '三菱受託用フラグ
UFMitsubishi.Inflag = "flagon" '三菱印刷用フラグ
'----------------------------ﾌﾞﾛｯｸNO判定
sh = "全ヵ所検査"
For Each ws In ThisWorkbook.Worksheets
If ws.Name = sh Then flag = True
Next ws
If flag = False Then Exit Sub
With ThisWorkbook.Sheets("全ヵ所検査")
  For i = 7 To 13 Step 2 '2020.01.09 追加
    If (.Cells(i, 2) <> "" And (.Cells(i, 27) = "" Or .Cells(i, 29) = "")) Then '測定値がありﾌﾞﾛｯｸNOの入力がない場合
     MsgBox "ﾌﾞﾛｯｸNOの入力をしてください", vbExclamation
     Exit Sub
    End If
  Next i
End With
If UFMitsubishi.Kname = "" And UFMitsubishi.KenNo = "" Then
 MsgBox "検査員名と検査番号を入力してください"
Exit Sub
Else
  UFProgress.Show vbModeless
  Call ProgressBar_count(0, 350)
  DoEvents
  Call 全ヵ所用転送
  Application.ScreenUpdating = False
    With ActiveSheet.PageSetup
      .LeftHeader = ""
      .CenterHeader = ""
      .RightHeader = ""
      .LeftFooter = ""
      .CenterFooter = ""
      .LeftFooter = ""
      .CenterFooter = ""
      .RightFooter = ""
    End With
    Call 原紙印刷範囲 'Sheets("転送済").PageSetup.PrintArea = "$B$2:$Q$20"
    Call MusicTrue("印刷.wav")
    If MsgBox("印刷しますか？", vbQuestion + vbYesNo) = vbYes Then
        ThisWorkbook.Sheets("転送済").Visible = True
        ThisWorkbook.Sheets("転送済").PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
        Call ProgressBar_count(350, 650)
        DoEvents
        Call 全ヵ所専用保存
        Call ProgressBar_count(650, 880)
        DoEvents
        ThisWorkbook.Activate
          Application.DisplayAlerts = False '保存後シートは消す
            If WsCheck("ｺﾋﾟ-用") Then ThisWorkbook.Sheets("ｺﾋﾟ-用").Delete
            If WsCheck("転送済") Then ThisWorkbook.Sheets("転送済").Delete
            If WsCheck("全ヵ所検査保存") Then ThisWorkbook.Sheets("全ヵ所検査保存").Delete
            If WsCheck("全ヵ所作業中保存") Then ThisWorkbook.Sheets("全ヵ所作業中保存").Delete
          Application.DisplayAlerts = True
    '----------------------------------------------履歴を残す　2019.09.13　これ以前のものは履歴なし
        If WsCheck("保存用") Then
          With ThisWorkbook.Sheets("保存用")
            n = .Cells(Rows.Count, 1).End(xlUp).row + 1
            .Unprotect
            .Cells.Locked = False
            .Cells.FormulaHidden = False
            .Cells(n, 2) = ThisWorkbook.Sheets("全ヵ所検査").Cells(2, 4) 'ロット
            .Cells(n, 83) = "三菱受託"
            .Cells(n, 3) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 2) '厚
            .Cells(n, 4) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 4) '幅
            .Cells(n, 5) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 6) '丈
            .Cells(n, 1) = Format$(ThisWorkbook.Sheets("全ヵ所検査").Cells(1, 2), "yyyy/mm/dd  h:mm") & "～" & Format$(Now, "yyyy/mm/dd  h:mm")
            .Cells(n, 77) = "別ファイル参照" '検査員
          End With
        End If
    Call ProgressBar_count(880, 1000)
    Call MusicTrue("完了.wav")
    DoEvents
    Unload UFProgress
    Else
      Unload UFProgress
      MsgBox "ｷｬﾝｾﾙ処理", vbInformation
    End If
End If
Application.ScreenUpdating = True
Application.WindowState = xlMaximized
UFMitsubishi.Tag = ""
End Sub

Private Sub CommandButton3_Click()
If UFMitsubishi.Tag = "off" Then Exit Sub

Dim i As Long, FL As Boolean, KenM As String
For i = 1 To 3
  If UFMitsubishi("Inspection" & i) = True Then
    KenM = UFMitsubishi("Inspection" & i).Caption
    FL = True
    Exit For
  End If
Next i
If FL = False Then
  MsgBox "検査法を選択してください", vbCritical
  Exit Sub
End If

Application.EnableEvents = False
UFProgress.Show vbModeless
UFProgress.CommandButton1.Caption = "転送中...."
Call ProgressBar_count(1, 730)
Call 全ヵ所用最大最少
Call ProgressBar_count(730, 1000)
Unload UFProgress
Call Ainput_Startup("Mitsubishi", KenM) '起動ﾌｫｰﾑ、検査法
Application.EnableEvents = True
End Sub

Private Sub CommandButton4_Click()
If UFMitsubishi.Choice <> "5N" Then
Call 全ヵ所用転送
Call 全ヵ所専用保存
Else
Call 全ヵ所用転送
Call FiveN専用保存
End If
'コピー用と転送済はここでは消さずに残す　印刷ボタンでは消す
End Sub

Private Sub CommandButton5_Click()
Call 全ヵ所専用保存
End Sub

Private Sub CommandButton6_Click()
'Call 原紙印刷範囲 'Sheets("転送済").PageSetup.PrintArea = "$B$2:$Q$20"
'ThisWorkbook.Sheets("5N").PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False

Call JustinCase_ZnSave
End Sub

Private Sub CommandButton7_Click()
Dim shName As String
Dim ws As Worksheet
Dim FL As Boolean
Dim n As Long
Dim Lname As String: Lname = UFMitsubishi.LineName.Caption
  
'注意表示---------------------------------------------------------------------------------------------
If WsCheck("全ヵ所検査") = False Then Exit Sub
ThisWorkbook.Worksheets("全ヵ所検査").Activate
Dim result1 As Long, result2 As Long, myRange1 As Range, myRange2 As Range, myRange3 As Range, myRange4 As Range, WSZ As Worksheet: Set WSZ = ThisWorkbook.Worksheets("全ヵ所検査")
Set myRange1 = Range(WSZ.Cells(7, 2), WSZ.Cells(76, 9))
Set myRange2 = Range(WSZ.Cells(7, 25), WSZ.Cells(76, 28))
Set myRange3 = (WSZ.Cells(2, 30))
Set myRange4 = (WSZ.Cells(3, 30))

If myRange3 = "" And myRange4 = "" Then
  MsgBox "検入高さを入力してください。", vbCritical
  Exit Sub
End If

result1 = CheckRedXCells(myRange1)
result2 = CheckRedXCells(myRange2)
If result1 > 0 And result2 > 0 Then
  If MsgBox("公差外判定の板があります " & "厚:" & result1 & "箇所　" & "歪:" & result2 & "箇所" & vbCrLf & "処理を続行しますか？", vbQuestion + vbYesNo) = vbNo Then Exit Sub
ElseIf result1 > 0 And result2 = 0 Then
  If MsgBox("公差外判定の板があります " & "厚:" & result1 & "箇所　" & vbCrLf & "処理を続行しますか？", vbQuestion + vbYesNo) = vbNo Then Exit Sub
ElseIf result1 = 0 And result2 > 0 Then
  If MsgBox("公差外判定の板があります " & "歪:" & result2 & "箇所" & vbCrLf & "処理を続行しますか？", vbQuestion + vbYesNo) = vbNo Then Exit Sub
End If
'------------------------------------------------------------------------------------------------------
  
UFMitsubishi.Tag = "off"
UFMitsubishi.Work = "OFF" '三菱作業フラグ
UFMitsubishi.Inflag = "" '印刷用フラグ
If UFMitsubishi.Kname = "" And UFMitsubishi.KenNo = "" Then
 MsgBox "検査員名と検査番号を入力してください"
Exit Sub
Else
UFMitsubishi.Tag = "off"
Call MusicTrue("おまちください.wav")
  UFProgress.Show vbModeless
  Call 全ヵ所用転送
  Call ProgressBar_count(0, 200)
  DoEvents
  Application.ScreenUpdating = False
  With ActiveSheet.PageSetup
      .LeftHeader = ""
      .CenterHeader = ""
      .RightHeader = ""
      .LeftFooter = ""
      .CenterFooter = ""
      .LeftFooter = ""
      .CenterFooter = ""
      .RightFooter = ""
  End With
    Call 全ヵ所用印刷範囲 'Sheets("全ヵ所検査").PageSetup.PrintArea = "$A$6:$AA$" & n
    Call MusicTrue("印刷.wav")
    If MsgBox("印刷しますか？", vbQuestion + vbYesNo) = vbYes Then
        ThisWorkbook.Sheets("全ヵ所検査").Visible = True
        ThisWorkbook.Sheets("全ヵ所検査").PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
        Call ProgressBar_count(200, 450)
        DoEvents
        Call 全ヵ所専用保存
        Call ProgressBar_count(450, 880)
        '-------------------------------------------------------------------不良要因
        If (WsCheck("不良要因印刷" & "_" & Lname) = False) Then 'シートなし
          Call AddWS("不良要因印刷" & "_" & Lname)
          Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
        ElseIf ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Cells(1, 2) <> UFMitsubishi.LotBox Then '不良の入力がなければこの状態
          Call AddWS("不良要因印刷" & "_" & Lname)
          Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname)
        End If
        Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname)
        sh.Visible = True
        sh.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
        sh.Visible = False
        '-------------------------------------------------------------------
        ThisWorkbook.Activate
          Application.DisplayAlerts = False '保存後シートは消す
            If WsCheck("ｺﾋﾟ-用") Then ThisWorkbook.Sheets("ｺﾋﾟ-用").Delete
            If WsCheck("全ヵ所検査保存") Then ThisWorkbook.Sheets("全ヵ所検査保存").Delete
            If WsCheck("全ヵ所作業中保存") Then ThisWorkbook.Sheets("全ヵ所作業中保存").Delete
          Application.DisplayAlerts = True
    '----------------------------------------------履歴を残す　2019.09.13　これ以前のものは履歴なし
        If WsCheck("保存用") Then
          With ThisWorkbook.Sheets("保存用")
            n = .Cells(Rows.Count, 1).End(xlUp).row + 1
            .Unprotect
            .Cells.Locked = False
            .Cells.FormulaHidden = False
            .Cells(n, 2) = ThisWorkbook.Sheets("全ヵ所検査").Cells(2, 4) 'ロット
            .Cells(n, 83) = "全ヵ所"
            .Cells(n, 3) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 2) '厚
            .Cells(n, 4) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 4) '幅
            .Cells(n, 5) = ThisWorkbook.Sheets("全ヵ所検査").Cells(3, 6) '丈
            .Cells(n, 1) = Format$(ThisWorkbook.Sheets("全ヵ所検査").Cells(1, 2), "yyyy/mm/dd  h:mm") & "～" & Format$(Now, "yyyy/mm/dd  h:mm")
            .Cells(n, 77) = "別ファイル参照" '検査員
'            .Cells.Locked = True
'            .Cells.FormulaHidden = False
''            .Protect Password:=""
'            .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'シート保護
          End With
        End If
      Call ProgressBar_count(880, 1000)
      Call MusicTrue("完了.wav")
      DoEvents
      Unload UFProgress
    Else
      Unload UFProgress
      MsgBox "ｷｬﾝｾﾙ処理", vbInformation
    End If
End If
Application.ScreenUpdating = True
Application.WindowState = xlMaximized
Application.Visible = True
UFMitsubishi.Tag = ""
End Sub

Private Sub Frame4_Click()
Application.Visible = True 'Excel表示
Dim ws As Worksheet
If WorkFlag.Caption = "P1" Then
  If WsCheck("全ヵ所検査") Then ThisWorkbook.Worksheets("全ヵ所検査").Activate
  Set ws = ThisWorkbook.Worksheets("全ヵ所検査")
End If
If WorkFlag.Caption = "P_H494" Then
  If WsCheck("H494検査") Then ThisWorkbook.Worksheets("H494検査").Activate
  Set ws = ThisWorkbook.Worksheets("H494検査")
End If

Dim TempW As Double: TempW = Weight(val(ws.Cells(3, 2)), val(ws.Cells(3, 4)), val(ws.Cells(3, 6)))
If TempW = 0 Then Exit Sub '重量なし処理停止
If 35 >= TempW Then MsgBox "ゲージ不要 * 両面定盤歪測定必要"
If 35 < TempW Then MsgBox "ゲージによる測定必要"
End Sub

Private Sub HindBo_Click()
Me.FMH.Height = 205
Me.FMH.Width = 655
Me.MLP.Width = 650
UFMitsubishi.Width = 670
Me.HindBo.BackColor = RGB(204, 255, 255)
End Sub

Private Sub KenNo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Application.EnableEvents = False
If Len(Me.LotBox.Text) <> 7 Then
Exit Sub
Else
 Me.ComboBox1 = Me.LotBox
End If
Application.EnableEvents = True
ActiveWindow.SmallScroll up:=99
End Sub

Private Sub MLP_Click(ByVal Index As Long)
On Error Resume Next
Application.Visible = True 'Excel表示

If WorkFlag.Caption = "P1" Then
  If WsCheck("全ヵ所検査") Then ThisWorkbook.Worksheets("全ヵ所検査").Activate
End If
If WorkFlag.Caption = "P_H494" Then
  If WsCheck("H494検査") Then ThisWorkbook.Worksheets("H494検査").Activate
End If

UFMitsubishi.MLP.Width = 500
UFMitsubishi.Width = 520
Dim ws As Worksheet
Select Case UFMitsubishi.WorkFlag.Caption
Case "P1"
  Set ws = ThisWorkbook.Sheets("全ヵ所検査")
Case "P_H494"
  Set ws = ThisWorkbook.Sheets("H494検査")
End Select

Dim TempW As Double: TempW = Weight(val(ws.Cells(3, 2)), val(ws.Cells(3, 4)), val(ws.Cells(3, 6)))
If TempW = 0 Then Exit Sub '重量なし処理停止
If 35 >= TempW Then MsgBox "ゲージ不要 * 両面定盤歪測定必要"
If 35 < TempW Then MsgBox "ゲージによる測定必要"

If ws.Cells(3, 2) = "" Then Exit Sub
Call DrawListBoxContent(UFMitsubishi, val(ws.Cells(3, 2)))

Dim Yamasu As String
Select Case True
Case IsNumeric(UFMitsubishi.KenHBox6): Yamasu = "6山"
Case IsNumeric(UFMitsubishi.KenHBox5): Yamasu = "5山"
Case IsNumeric(UFMitsubishi.KenHBox4): Yamasu = "4山"
Case IsNumeric(UFMitsubishi.KenHBox3): Yamasu = "3山"
Case IsNumeric(UFMitsubishi.KenHBox2): Yamasu = "2山"
Case IsNumeric(UFMitsubishi.KenHBox1): Yamasu = "1山"
End Select
With UFMitsubishi
  If (.KenHBox7 = "") And (.KenHBox6 = "" And .KenHBox5 = "" And .KenHBox4 = "" And .KenHBox3 = "" And .KenHBox2 = "" And .KenHBox1 = "") Then
    ws.Unprotect 'シート保護解除
    ws.Cells(1, 30) = ""
    ws.Cells(2, 30) = ""
    ws.Cells(3, 30) = ""
    ws.Protect Contents:=True, Scenarios:=True, AllowFormattingCells:=True 'シート保護
  ElseIf (.KenHBox7 = "") Then
    ws.Unprotect
    ws.Cells(3, 30) = ""
    ws.Protect Contents:=True, Scenarios:=True, AllowFormattingCells:=True
  End If
End With
ws.Unprotect
ws.Cells(1, 30) = UFMitsubishi.高さ計算式.Caption
ws.Cells(2, 30) = Yamasu
Range(ws.Cells(1, 30), ws.Cells(3, 30)).Font.Size = 8
Range(ws.Cells(1, 30), ws.Cells(3, 30)).HorizontalAlignment = xlLeft
ws.Protect Contents:=True, Scenarios:=True, AllowFormattingCells:=True

UFMitsubishi.YamaLB.SetFocus
End Sub

Private Sub MLP_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'Me.MLP.width = 500
'UFMitsubishi.width = 520

Me.FMH.Width = 1
Me.FMH.Height = 1
Me.HindBo.BackColor = &HFFFFFF
End Sub

'Private Sub ViewCB_Click()
'Dim Lname As String: Lname = UFMitsubishi.LineName.Caption
''入力なし、シートなしの場合からの確認時対策
'If (WsCheck("不良要因印刷" & "_" & Lname) = False) Then
'  Call AddWS("不良要因印刷" & "_" & Lname)
'  Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
'ElseIf ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Cells(1, 2) <> UFMitsubishi.LotBox Then
'  Call AddWS("不良要因印刷" & "_" & Lname)
'  Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
'End If
'Dim WS As Worksheet: Set WS = ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname)
''Range(WS.Cells(1, 2), WS.Cells(10, 24)).Copy
''WS.Visible = xlSheetVisible
'
'' 高解像度でコピーするための設定
'Call CopyRangeAsHighResolutionImage(WS, WS.Range(WS.Cells(1, 2), WS.Cells(10, 24)))
'
'Me.Zoom = 100
'
'With UFMitsubishi
'  .NgIm.Visible = True
'  .DefectFactorLB.Visible = False
'  .Height = 535
'  .Width = 900
'
'  .MLP.Height = 535
'  .MLP.Width = 900
'
'
'  Set .NgIm.Picture = CreatePictureFromClipboard
'  Sleep 8
'  .Repaint
'  Set .NgIm.Picture = CreatePictureFromClipboard
'  Sleep 8
'End With
'End Sub
Private Sub ViewCB_Click()
Dim Lname As String: Lname = UFMitsubishi.LineName.Caption
Dim ws As Worksheet
Dim tempPicture As StdPicture
Dim retryCount As Integer
Dim targetWidth As Single, targetHeight As Single

' プログレス表示開始
Call ShowProgress("データ準備中...")

'入力なし、シートなしの場合からの確認時対策
If (WsCheck("不良要因印刷" & "_" & Lname) = False) Then
    Call AddWS("不良要因印刷" & "_" & Lname)
    Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
ElseIf ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Cells(1, 2) <> UFMitsubishi.LotBox Then
    Call AddWS("不良要因印刷" & "_" & Lname)
    Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
End If

Set ws = ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname)

Call ShowProgress("画像データをコピー中...")

' クリップボードをクリア
Application.CutCopyMode = False
DoEvents
Sleep 300

' 高解像度コピーまたは通常のコピー処理
retryCount = 0
Do While retryCount < 3
    ' 高解像度コピー関数が存在する場合はそれを使用、そうでなければ通常のコピー
    On Error Resume Next
    Call CopyRangeAsHighResolutionImage(ws, ws.Range(ws.Cells(1, 2), ws.Cells(11, 24)))
    If Err.Number <> 0 Then
        ' 高解像度コピー関数がない場合は通常のコピー
        Err.clear
        ws.Range(ws.Cells(1, 2), ws.Cells(10, 24)).Copy
    End If
    On Error GoTo 0
    
    DoEvents
    Sleep 500 ' コピー完了を確実に待つ
    
    ' クリップボードにデータがあるか確認
    If IsClipboardFormatAvailable(CF_BITMAP) <> 0 Or IsClipboardFormatAvailable(CF_ENHMETAFILE) <> 0 Then
        Exit Do
    End If
    
    retryCount = retryCount + 1
    Call ShowProgress("画像データをコピー中... (再試行 " & retryCount & "/3)")
Loop

Call ShowProgress("画像を取得中...")

' UFMitsubishiフォームの設定
Me.Zoom = 100

With UFMitsubishi
    ' 画像取得の試行
    retryCount = 0
    Do While retryCount < 8
        Set tempPicture = CreatePictureFromClipboard
        If Not tempPicture Is Nothing Then Exit Do
        
        retryCount = retryCount + 1
        Call ShowProgress("画像を取得中... (試行 " & retryCount & "/8)")
        DoEvents
        Sleep 300
        
        ' 3回目以降は再コピーも試行
        If retryCount >= 3 Then
            On Error Resume Next
            Call CopyRangeAsHighResolutionImage(ws, ws.Range(ws.Cells(1, 2), ws.Cells(11, 24)))
            If Err.Number <> 0 Then
                Err.clear
                ws.Range(ws.Cells(1, 2), ws.Cells(11, 24)).Copy
            End If
            On Error GoTo 0
            DoEvents
            Sleep 300
        End If
    Loop
    
    ' 画像取得失敗時のエラーハンドリング
    If tempPicture Is Nothing Then
        Call HideProgress
        MsgBox "画像の取得に失敗しました。再度お試しください。", vbExclamation
        Exit Sub
    End If
    
    Call ShowProgress("画面を更新中...")
    
    ' フォームサイズとコントロールの設定
    .Height = 535
    .Width = 900
    .MLP.Height = 535
    .MLP.Width = 900
    
    ' NgIm画像コントロールの設定（サイズ固定化）
    .NgIm.Visible = False ' 一旦非表示にして設定
    .NgIm.AutoSize = False ' AutoSizeを無効化
    .NgIm.PictureSizeMode = 1 ' fmPictureSizeModeStretch
    
    ' 固定サイズを設定（フォームサイズに合わせて調整）
    targetWidth = 800  ' 目標幅
    targetHeight = 450 ' 目標高さ
    
    .NgIm.Width = targetWidth
    .NgIm.Height = targetHeight
    
    ' UI要素の制御
    .DefectFactorLB.Visible = False
    
    Call ShowProgress("画像を設定中...")
    
    ' 画像の設定
    Set .NgIm.Picture = tempPicture
    DoEvents
    Sleep 200
    
    ' NgImを表示
    .NgIm.Visible = True
    
    ' フォーム更新
    .Repaint
    DoEvents
    Sleep 200
    
    ' 最終確認と修正
    If .NgIm.Picture Is Nothing Then
        Call ShowProgress("画像を再設定中...")
        
        ' 最終リトライ
        On Error Resume Next
        Call CopyRangeAsHighResolutionImage(ws, ws.Range(ws.Cells(1, 2), ws.Cells(11, 24)))
        If Err.Number <> 0 Then
            Err.clear
            ws.Range(ws.Cells(1, 2), ws.Cells(11, 24)).Copy
        End If
        On Error GoTo 0
        DoEvents
        Sleep 500
        
        Set tempPicture = CreatePictureFromClipboard
        If Not tempPicture Is Nothing Then
            Set .NgIm.Picture = tempPicture
            .Repaint
            DoEvents
        Else
            Call HideProgress
            MsgBox "画像の表示に失敗しました。", vbExclamation
            Exit Sub
        End If
    End If
    
    ' サイズの最終調整（画像が設定された後）
    .NgIm.Width = targetWidth
    .NgIm.Height = targetHeight
    .Repaint
    
End With

' 完了表示
Call ShowProgress("完了しました")
Sleep 1000
Call HideProgress

' クリーンアップ
Application.CutCopyMode = False
Set tempPicture = Nothing
Set ws = Nothing
End Sub


Private Sub ws_out_Click()
Dim Lname As String: Lname = UFMitsubishi.LineName.Caption
If UFMitsubishi.Tag = "off" Then Exit Sub
If Register_Check(Format$(Now, "hh:mm:ss")) Then Call Daily_Create '日報作成
If MsgBox("新規発行しますか？現在のｼｰﾄは削除されます。", vbQuestion + vbYesNo) = vbYes Then
  UFProgress.Show vbModeless
  UFProgress.CommandButton1.Caption = "NewDataCreate...."
  Call ProgressBar_count(0, 350)
  DoEvents
  Call JustinCase_ZnSave("H494検査")
  Call ProgressBar_count(350, 560)
  DoEvents
  Call 全ヵ所用コピー("H494検査")
  
  Select Case UFMitsubishi.LineName.Caption
  Case "L-1":    Call WriteLog("*UFKen:" & "H494検査ｼｰﾄ発行", "", HistoryPath & "L-1\L-1_Log")
  Case "LVC":    Call WriteLog("*UFKen:" & "H494検査ｼｰﾄ発行", "", HistoryPath & "LVC\LVC_Log")
  Case "HVC":    Call WriteLog("*UFKen:" & "H494検査ｼｰﾄ発行", "", HistoryPath & "HVC\HVC_Log")
  Case "ﾊﾞﾗﾝｻｰ": Call WriteLog("*UFKen:" & "H494検査ｼｰﾄ発行", "", HistoryPath & "バランサー\バランサー_Log")
  Case "ﾄｯﾄ":    Call WriteLog("*UFKen:" & "H494検査ｼｰﾄ発行", "", HistoryPath & "トット\トット_Log")
  Case "Test":   Call WriteLog("*UFKen:" & "H494検査ｼｰﾄ発行", "", HistoryPath & "Test\Test_Log")
  Case Else:     Call WriteLog("*UFKen:" & "H494検査ｼｰﾄ発行", "", HistoryPath & "丸徳\丸徳_Log")
  End Select
  
  If WsCheck("不良要因印刷" & "_" & Lname) Then ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Delete '2023.10.25_追加
  
  UFMitsubishi.H494_Lot = ""
  UFMitsubishi.H494_KenNo = ""
  UFMitsubishi.ComboBox1 = ""
  UFMitsubishi.H494_Youto = ""
  Call ProgressBar_count(560, 1000)
  DoEvents
  Unload UFProgress
Else
  Unload UFProgress
  MsgBox "ｷｬﾝｾﾙ処理", vbInformation
End If
End Sub

Private Sub YamaLB_Click()
If WsCheck("全ヵ所検査") = False Then Exit Sub
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("全ヵ所検査")
If ws.Cells(3, 2) = "" Then Exit Sub
Call DrawListBoxContent(UFMitsubishi, val(ws.Cells(3, 2)))
End Sub

Private Sub シート発行_Click()
Application.Visible = True 'Excel表示
Dim ws As Worksheet
If WorkFlag.Caption = "P1" Then
  If WsCheck("全ヵ所検査") Then ThisWorkbook.Worksheets("全ヵ所検査").Activate
  Set ws = ThisWorkbook.Worksheets("全ヵ所検査")
End If
If WorkFlag.Caption = "P_H494" Then
  If WsCheck("H494検査") Then ThisWorkbook.Worksheets("H494検査").Activate
  Set ws = ThisWorkbook.Worksheets("H494検査")
End If

Dim TempW As Double: TempW = Weight(val(ws.Cells(3, 2)), val(ws.Cells(3, 4)), val(ws.Cells(3, 6)))
If TempW = 0 Then Exit Sub '重量なし処理停止
If 35 >= TempW Then MsgBox "ゲージ不要 * 両面定盤歪測定必要"
If 35 < TempW Then MsgBox "ゲージによる測定必要"
End Sub

Private Sub 一般_Click()
UFMitsubishi.Choice = "一般"
 Call ForN_Select
End Sub

Private Sub SB_SpinDown()
If UFMitsubishi.Tag = "off" Then Exit Sub
  ActiveWindow.SmallScroll down:=3
  MoveTg("A", "全ヵ所検査").Activate
End Sub
Private Sub SB_SpinUp()
If UFMitsubishi.Tag = "off" Then Exit Sub
  ActiveWindow.SmallScroll up:=3
  MoveTg("A", "全ヵ所検査").Activate
End Sub

Private Sub UserForm_Activate()
Dim a As Long, i As Long, k As Long
If UFMitsubishi.Tag = "off" Then Exit Sub
Application.Visible = True
Application.WindowState = xlMaximized '最大化

Dim ws As Worksheet
Dim Temp
If WorkFlag.Caption = "P1" Then
    If WsCheck("全ヵ所検査用原紙") = False Then Exit Sub
    If UFMitsubishi.LineName = "丸徳" Then
      Temp = "一般"
      UFMitsubishi.Choice = Temp
    End If
    '2021.09.15_作業選択強制
    If UFMitsubishi.Choice = "" Then
      Temp = Application.InputBox("作業名入力：三菱、5N、一般　のいずれかを入力")  '作業名入力
      If Temp = "５N" Then Temp = "5N"
      If Temp = False Then
        End
      ElseIf Temp = "" Then
        End
      Else
        UFMitsubishi.Choice = Temp
        Call ForN_Select
      End If
    Else
      Call ForN_Select
    End If
    Application.ScreenUpdating = False
    With ThisWorkbook
      For i = 1 To .Sheets.Count
         If .Sheets(i).Name = "全ヵ所検査" Or .Sheets(i).Name = "全ヵ所検査保存" Or .Sheets(i).Name = "全ヵ所作業中保存" Then
         .Sheets(i).Visible = True
         End If
      Next i
    End With
    If WsCheck("全ヵ所検査") Then ThisWorkbook.Worksheets("全ヵ所検査").Activate
ElseIf WorkFlag.Caption = "P_H494" Then
    If WsCheck("H494_原紙") = False Then Exit Sub
    Application.ScreenUpdating = False
    With ThisWorkbook
      For i = 1 To .Sheets.Count
         If .Sheets(i).Name = "H494検査" Or .Sheets(i).Name = "H494検査保存" Or .Sheets(i).Name = "H494作業中保存" Then
         .Sheets(i).Visible = True
         End If
      Next i
    End With
    If WsCheck("H494検査") = False Then 全ヵ所用コピー ("H494検査")
    If WsCheck("H494検査") Then ThisWorkbook.Worksheets("H494検査").Activate
End If

Application.EnableEvents = True
Application.ScreenUpdating = True

'With UFMitsubishi
''一般作業
'  .Choice = "一般"
'  .Work = "OFF"
'  .Inflag = ""
'  .Normal.Visible = True
'  .Mitsubishi.Visible = False
'  .FiNi.Visible = False
'  WS.Unprotect '保護解除
'  Range(WS.Cells(7, 27), Cells(7, 29)).Locked = True 'ロック設定
'  Range(WS.Cells(9, 27), Cells(9, 29)).Locked = True
'  Range(WS.Cells(11, 27), WS.Cells(11, 29)).Locked = True
'  Range(WS.Cells(13, 27), WS.Cells(13, 29)).Locked = True
'  Union(Range(WS.Cells(7, 27), WS.Cells(7, 29)), Range(WS.Cells(9, 27), WS.Cells(9, 29)), _
'  Range(WS.Cells(11, 27), WS.Cells(11, 29)), Range(WS.Cells(13, 27), WS.Cells(13, 29))).Interior.Color = RGB(192, 192, 192)
'  Range(WS.Cells(6, 30), WS.Cells(14, 30)).Font.Color = RGB(255, 255, 255)
'  Range(WS.Cells(6, 31), WS.Cells(11, 31)).Font.Color = RGB(255, 255, 255)
'  WS.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
'End With
End Sub

Private Sub UserForm_Click()
If UFMitsubishi.Tag = "off" Then Exit Sub
Application.EnableEvents = True
Application.Visible = True
Application.WindowState = xlMaximized '最大化

Dim ws As Worksheet
If WorkFlag.Caption = "P1" Then
  If WsCheck("全ヵ所検査") Then ThisWorkbook.Worksheets("全ヵ所検査").Activate
  Set ws = ThisWorkbook.Worksheets("全ヵ所検査")
End If
If WorkFlag.Caption = "P_H494" Then
  If WsCheck("H494検査") Then ThisWorkbook.Worksheets("H494検査").Activate
  Set ws = ThisWorkbook.Worksheets("H494検査")
End If

Dim TempW As Double: TempW = Weight(val(ws.Cells(3, 2)), val(ws.Cells(3, 4)), val(ws.Cells(3, 6)))
If TempW = 0 Then Exit Sub '重量なし処理停止
If 35 >= TempW Then MsgBox "ゲージ不要 * 両面定盤歪測定必要"
If 35 < TempW Then MsgBox "ゲージによる測定必要"


End Sub

Private Sub UserForm_Initialize()
If UFMitsubishi.Tag = "off" Then Exit Sub
UFMitsubishi.ComboBox1.Tag = "Load"
Dim i As Long, a As Long

Dim tmpFaci As String: tmpFaci = GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")
Dim tmpMode As String: tmpMode = GetSetting("UF検査判定システム", "起動設定", "モード", "存在なし")
If tmpMode = "オフライン" Then Me.Caption = Me.Caption & tmpFaci & "  " & "「オフライン」"
If tmpMode = "オンライン" Then Me.Caption = Me.Caption & tmpFaci & "  " & "「オンライン」"


Dim Comb As Variant: Comb = Array("1", "2", "3", "4", "5", "6", "枚数")
UFMitsubishi.YamaLB.List = Comb
 
'--------------------------------------------------------------------------------------------------不良用
Dim Comb_data As Variant, List_Data1() As Variant, List_Data2() As Variant


If WsCheck("不良要因印刷原紙") = False Then Exit Sub
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("不良要因印刷原紙")
Dim rangeA As Range: Set rangeA = ws.Range(ws.Cells(2, 3), ws.Cells(2, 24))
Dim rangeB As Range: Set rangeB = ws.Range(ws.Cells(9, 3), ws.Cells(9, 24))

Dim TempA(): TempA = rangeA
Dim tempB(): tempB = rangeB

Dim Temp():        Temp = Call_MergeArray_Row(TempA, tempB) '二次元配列を行方向に結合
Dim ReplaceTemp(): ReplaceTemp = WorksheetFunction.Transpose(Temp) 'テーブルデータを縦横変換

Call GetZoom(Me)
Dim MiTemp As Variant:  MiTemp = InitialSetting(Me)

With UFMitsubishi.DefectFactorLB
  .ColumnCount = 2 '表示列数
  .ColumnWidths = "78;15"  '列幅
  .List = ReplaceTemp
  .AddItem "ｽﾎﾟｯﾄ処理" '追加_2023.6.16
End With
'--------------------------------------------------------------------------------------------------
UFMitsubishi.ComboBox1.Tag = ""
Application.EnableEvents = True
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Me.MLP.Width = 500
UFMitsubishi.Width = 520
Me.FMH.Width = 1
Me.FMH.Height = 1
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim i As Long
Application.Visible = True
Application.WindowState = xlMaximized '最大化
Application.ScreenUpdating = False
With ThisWorkbook
For i = 1 To .Sheets.Count
   If .Sheets(i).Name = "全ヵ所検査" Or .Sheets(i).Name = "全ヵ所検査保存" _
   Or .Sheets(i).Name = "全ヵ所作業中保存" Or .Sheets(i).Name = "H494検査" _
   Or .Sheets(i).Name = "H494検査保存" Or .Sheets(i).Name = "H494作業中保存" Then
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

