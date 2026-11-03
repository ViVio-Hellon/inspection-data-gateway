Attribute VB_Name = "UFAinput"
Attribute VB_Base = "0{A806C58A-77C5-44B1-A9D0-DE7BD65E5160}{D9B964FE-60F7-4F40-8148-8AF09D887686}"
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

#If VBA7 And Win64 Then
    Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
#Else
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#End If
Private Const VK_CAPITAL  As Long = &H14
Private Const KEYEVENTF_KEYUP As Long = &H2

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If

'点滅用
Private BlinkLabel1 As clsBlinkLabel

' 動的ラベルイベント管理　clsDynLabel の配列宣言
Private DynLabelEvents(1 To 15) As clsDynLabel

' Excelウィンドウ状態退避用
Private g_sngTopSV As Single ' 上端位置
Private g_sngLeftSV As Single ' 左端位置
Private g_sngWidthSV As Single ' 幅
Private g_sngHeightSV As Single ' 高さ
' ウィンドウ状態
Private g_lngWindowState As XlWindowState

Private Sub GaityuuText_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
Dim path As String: path = HistoryPath
Dim Temp As String: Temp = "*UFKen:計算高さ操作"
Select Case True
Case UF1.L1Work:   Call WriteLog(Temp, "", path & "L-1\L-1_Log")
Case UF1.LVC:      Call WriteLog(Temp, "", path & "LVC\LVC_Log")
Case UF1.HVC:      Call WriteLog(Temp, "", path & "HVC\HVC_Log")
Case UF1.Bala:     Call WriteLog(Temp, "", path & "バランサー\バランサー_Log")
Case UF1.Tott:     Call WriteLog(Temp, "", path & "トット\トット_Log")
Case UF1.Test:     Call WriteLog(Temp, "", path & "Test\Test_Log")
Case Else:         Call WriteLog(Temp, "", path & "丸徳\丸徳_Log")
End Select
End Sub
Private Sub AiText_Change() 'イベントでラベルに文字列を転写
UFAinput.Aisi.Caption = UFAinput.AiText
End Sub
Private Sub GaText_Change() 'イベントでラベルに文字列を転写
UFAinput.Gaisousi.Caption = UFAinput.GaText
End Sub

Private Sub VcText_Change() 'イベントでラベルに文字列を転写
UFAinput.VCAB.Caption = UFAinput.VcText
End Sub

Private Sub Ai_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean) 'チェックを消したときtextから元の文字列を復旧させる
On Error Resume Next
'>>>>>>>>>>>>>>>>>>>>>Height_Checkも修正要
'正規引当(外注戻り)　個装の場合、積上げ出荷であってもAｲﾝﾌﾟｯﾄはL1検査法にて個装で入力を行う
'その為個装での本当での積高さは偽装するしか方法がなく、このｼｽﾃﾑが使用できるのは正規引当も積上げの場合の時だけである
'Vの有無、合紙の有無、個装外装紙の有無を操作しても正規が個装の場合個装としての積高さを用意できない
If UFAinput.Aisi.Caption = "" Then Exit Sub
Dim BackText As String: BackText = UFAinput.AiText
If UFAinput.GaityuuText = "外注" And UFAinput.Ai Then
  If Not UFAinput.Aisi.Caption = 0 Then UFAinput.Aisi.Caption = "なし"
Else
  UFAinput.Aisi.Caption = BackText
End If
End Sub

Private Sub Ga_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
Dim BackText As String: BackText = UFAinput.GaText
If UFAinput.GaityuuText = "外注" And UFAinput.Ga Then
  UFAinput.Gaisousi.Caption = "なし"
Else
  UFAinput.Gaisousi.Caption = BackText
End If
End Sub

Private Sub Vab_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
Dim BackText As String: BackText = UFAinput.VcText
If UFAinput.GaityuuText = "外注" And UFAinput.Vab Then
  UFAinput.VCAB.Caption = "なし"
Else
  UFAinput.VCAB.Caption = BackText
End If
End Sub
Private Sub CommandButton10_Click()
If UFAinput.Ai.Locked Then
  UFAinput.Ai.Locked = False
  UFAinput.Ga.Locked = False
  UFAinput.Vab.Locked = False
  UFAinput.GaityuuText.Visible = True
  MsgBox "合紙,VCの有無を変更したい場合" & vbCrLf & "赤BOXに 外注 と入力してください。", vbInformation
End If
End Sub
Private Sub CommandButton1_Click()
'Call Daily_Report4_Execution
Application.Visible = True 'Excel表示
Dim i As Long
With ThisWorkbook
  .Sheets("Daily_Test").Visible = xlSheetVisible
End With
'Call Daily_Report(2) 'ロット情報
'Call Daily_Report(3) '開始時間
Call Daily_Report(4) 'Aインプット
End Sub
Private Sub CommandButton2_Click()
With UFAinput
  .HikiNo = ""
  .HIkiNoUnder = ""
  .Remain = ""
  .KONNO = ""
  .MAI1 = ""
  .COU1 = ""
  .MAI2 = ""
  .COU2 = ""
  .Material1 = ""
  .Material2 = ""
  .Material3 = ""
  
  .GW = ""
  .PileMatchNo = ""
  .PileMatch = ""
  .Mount_Count = ""
  
  Call ClearAllKHeight '2026.05.29
  
  .Kflag = ""
  .K_Height.Visible = False
  .K_Height_2.Visible = False
  .K_Height_3.Visible = False
  .K_Height_4.Visible = False
  .L1.Visible = False
  .L2.Visible = False
  
  .ExCaseNo = ""
  .ExMaterial1 = ""
  .ExMaterial2 = ""
  .ExMaterial3 = ""
  .ExGW = ""
  .OutsideAtu = ""
  .OutsideHaba = ""
  .OutsideTake = ""
End With
Dim A_Icas As Variant: A_Icas = FormArray(UFAinput, 2) 'ｱｲｶｽ用
If IsEmpty(A_Icas) Then Exit Sub
Dim Lname As String: Lname = UFAinput.LineName
Call AddWS("A_Icas" & "_" & Lname)
Dim SHIcas As Worksheet: Set SHIcas = ThisWorkbook.Sheets("A_Icas" & "_" & Lname) 'ｱｲｶｽ用
SHIcas.Cells.NumberFormatLocal = "@" '文字列
Dim Past_Inp As Range: Set Past_Inp = SHIcas.Cells(1, 1) 'ｱｲｶｽ用
Past_Inp = A_Icas 'ｱｲｶｽ用
Call ExpCSV(SHIcas, "hostsend", NisPath, vbTab) 'CSV保存---保存先でAｲﾝﾌﾟｯﾄ用出力になる　*出力ﾌｧｲﾙ,ﾌｧｲﾙ名,ﾊﾟｽ,出力形式
Sleep 1000
Call App_Activation 'ｱﾌﾟﾘ起動
End Sub

Private Sub CommandButton2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Me.Re_hostsend.Visible = True
Me.OutCondition.Visible = False
Me.GwOpen.Visible = False
Me.Re_hostsend.Caption = HostDataConfi
End Sub

Private Sub CommandButton3_Click()
Call A_Output(UFAinput, Me.Lot, Me.KenNo)
Application.Visible = True 'Excel表示
End Sub

Private Sub CommandButton4_Click()
Call InterRuptText(, , 5000)
End Sub

Private Sub CommandButton5_Click()
Call App_Activation   '("C:\Users\", "BouyomiChan.exe")
End Sub

Private Sub CommandButton6_Click()
Unload Me
End Sub

Private Sub CommandButton7_Click()
Dim A_niti As Variant: A_niti = FormArray(UFAinput, 3) '梱包用_変換無しに変更*2021.04.01
If IsEmpty(A_niti) Then Exit Sub
Dim Lname As String: Lname = UFAinput.LineName

Call AddWS("Nitiina" & "_" & Lname)
Dim SHIna As Worksheet:  Set SHIna = ThisWorkbook.Sheets("Nitiina" & "_" & Lname) '梱包用

SHIna.Cells.NumberFormatLocal = "@" '文字列
Dim Past_Ni As Range:  Set Past_Ni = Range(SHIna.Cells(LBound(A_niti, 1) + 1, 1), SHIna.Cells(UBound(A_niti, 1) + 1, UBound(A_niti, 2))) '梱包用
Application.Visible = True
SHIna.Visible = xlSheetVisible
Past_Ni = A_niti
'Stop
'Past_Ni = Application.WorksheetFunction.Transpose(A_niti) '行列入れ替え'梱包保存用*FormArray(UFAinput, 3)以外で使用
Call A_CsvSave(SHIna.Name, "A111119", "B22229") 'CSV保存
End Sub

Private Sub CommandButton8_Click()
Dim Wb_Path As String: Wb_Path = ThisWorkbook.path & "\日報＆資材計算.xlsm"
'2025.5.14
If ReadCheck(Wb_Path, 1) Then '有無のチェック
  If IsBookOpen(Wb_Path) = False Then '開いてない
    '実行しているエクセルとは、違う新しいアプリケーションでエクセルを開く
    Dim ExcelApp As New Application
    ExcelApp.Visible = False
    ExcelApp.DisplayAlerts = False
    ' UpdateLinks:=0, ReadOnly:=True, CorruptLoad:=xlRepairFile) Workbooks.Open(Wb_Path)
    Dim wb As Workbook: Set wb = Workbooks.Open(Wb_Path) ' ExcelApp.Workbooks.Open(Wb_Path, , True)  '読取り専用
    UFAinput.Left = UFAinput.Left - 5
    
    Application.Run "'日報＆資材計算.xlsm'!GW_Open"
  Else
    Application.Run "'日報＆資材計算.xlsm'!GW_Open"
  End If
Else
'  MsgBox "日報＆資材計算ファイル無し"
  Exit Sub
End If
End Sub

Private Sub CommandButton8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Me.Re_hostsend.Visible = False
Me.OutCondition.Visible = False
Me.GwOpen.Visible = True
Me.GwOpen.Caption = "日報＆資材計算が開いていることが起動条件"
End Sub

Private Sub CommandButton9_Click()
'クリア
UFAinput.AiText = ""
UFAinput.GaText = ""
UFAinput.VcText = ""
Dim Ainput As Variant
Dim Workers As String, path As String: path = HistoryPath
With UFAinput
  '枚数
  Dim M_Count As Long
  Select Case .MaiFlag.Caption '包み枚数の表示ﾌﾗｸﾞ
  Case "1": M_Count = val(.MAI1) * val(.COU1)
  Case "2": M_Count = (val(.MAI1) * val(.COU1)) + (val(.MAI2) * val(.COU2))
  End Select
  '山高さ
  Dim Mount_Height As Single
  Mount_Height = CalcMountHeight(val(.Mount_Count.Caption)) '_2026.05.29
  
  If Height_Check(val(UFAinput.AMin.Caption), val(UFAinput.AMax.Caption), val(Mount_Height), val(M_Count), UFAinput.OrderNo.Caption) = False Then

End If
End With
UFAinput.GaityuuText = ""
UFAinput.Ai.Locked = True
UFAinput.Ga.Locked = True
UFAinput.Vab.Locked = True
UFAinput.GaityuuText.Visible = False
End Sub

Private Sub COU1_Change()
If IsNumeric(UFAinput.COU1) = False Then UFAinput.COU1 = ""
End Sub
Private Sub COU2_Change()
If IsNumeric(UFAinput.COU2) = False Then UFAinput.COU2 = ""
End Sub
Private Sub ExCaseNo_Change()
If IsNumeric(UFAinput.ExCaseNo) = False Then UFAinput.ExCaseNo = ""
End Sub
Private Sub ExGW_Change()
If IsNumeric(UFAinput.ExGW) = False Then UFAinput.ExGW = ""
End Sub
Private Sub ExMaterial1_Change()
If IsNumeric(UFAinput.ExMaterial1) = False Then UFAinput.ExMaterial1 = ""
End Sub
Private Sub ExMaterial2_Change()
If IsNumeric(UFAinput.ExMaterial2) = False Then UFAinput.ExMaterial2 = ""
End Sub
Private Sub ExMaterial3_Change()
If IsNumeric(UFAinput.ExMaterial3) = False Then UFAinput.ExMaterial3 = ""
End Sub
Private Sub ExMaterial1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Me.DigitLabel = "3桁までの入力"
End Sub
Private Sub ExMaterial2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Me.DigitLabel = "3桁までの入力"
End Sub
Private Sub ExMaterial3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Me.DigitLabel = "3桁までの入力"
End Sub
Private Sub ExFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Me.DigitLabel = ""
End Sub

Private Sub GW_Change()
If IsNumeric(UFAinput.GW) = False Then UFAinput.GW = ""
End Sub

Private Sub HikiNo_Change() 'ｵｰﾀﾞｰNOの頭3桁目が1のとき輸出
UFAinput.OrderNo = ""
UFAinput.AdjustLabel.Visible = True
UFAinput.AdjustNo.Visible = True
'ソート処理を入れたためコンボボックスの並び順とシートの行順がズレる、選択値で Hiki_Lot を検索する方式に変更_2026.06.08
Dim b As Long: b = UFAinput.HikiNo.ListIndex
If Not b = -1 Then
  ' ソート後のListIndexとシート行がズレるため選択値でシートを検索
  Dim hikiSH As Worksheet: Set hikiSH = ThisWorkbook.Sheets("Hiki_Lot")
  Dim selectedVal As String: selectedVal = UFAinput.HikiNo.Value
  Dim lastRow As Long: lastRow = hikiSH.Cells(Rows.Count, 1).End(xlUp).row
  Dim rowIdx As Long: rowIdx = 0
  Dim k As Long
  For k = 1 To lastRow
      If CStr(hikiSH.Cells(k, 1)) = selectedVal Then
          rowIdx = k
          Exit For
      End If
  Next k
  If rowIdx = 0 Then Exit Sub  ' 一致なし
  UFAinput.OrderNo = hikiSH.Cells(rowIdx, 2)
  If UFAinput.DBflag.Caption = "true" Then UFAinput.Lot = hikiSH.Cells(rowIdx, 3)
  If UFAinput.DBflag.Caption = "true" Then UFAinput.KenNo = hikiSH.Cells(rowIdx, 4)
  UFAinput.AdjustLabel.Caption = "調整No"
  UFAinput.AdjustNo = hikiSH.Cells(rowIdx, 5)
  If Not UFAinput.AdjustNo = "0" Then
  ElseIf UFAinput.AdjustNo = "0" Then
    UFAinput.AdjustLabel.Caption = "引当数量"
    UFAinput.AdjustNo = hikiSH.Cells(rowIdx, 6)
  End If
  If UFAinput.AdjustNo = "0" Then
    UFAinput.AdjustLabel.Visible = False
    UFAinput.AdjustNo.Visible = False
  End If
  
  Dim ExFlag As Long: ExFlag = Mid(UFAinput.OrderNo.Caption, 3, 1) '頭3桁目の確認
  If ExFlag = 1 Then
    UFAinput.ExportLabel.Visible = True
    UFAinput.ExFrame.Visible = True
    UFAinput.Height = 620 '560
  Else
    UFAinput.ExportLabel.Visible = False
    UFAinput.ExFrame.Visible = False
    UFAinput.Height = 490 '437
  End If
End If

'Dim AinTemp As Variant:  AinTemp = InitialSetting(UFAinput)
UFAinput.tOP = (Application.Height - UFAinput.Height) / 2
UFAinput.Left = (Application.Width - UFAinput.Width) / 2
End Sub
Private Sub HIkiNoUnder_Change()
If IsNumeric(UFAinput.HIkiNoUnder) = False Then UFAinput.HIkiNoUnder = ""
End Sub

Private Sub Kflag_Change()
If IsNumeric(UFAinput.Kflag) = False Then UFAinput.Kflag = ""
End Sub
Private Sub KONNO_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If (Len(UFAinput.KONNO) = 1 And UFAinput.KONNO <> 0) Then UFAinput.KONNO = "0" & UFAinput.KONNO
If UFAinput.KONNO = 0 Then UFAinput.KONNO = ""
End Sub
Private Sub KONNO_Change()
If IsNumeric(UFAinput.KONNO) = False Then UFAinput.KONNO = ""
End Sub
Private Sub L2_Click()
On Error Resume Next
UFAinput.L1.SetFocus
UFAinput.StartTB.SetFocus
End Sub
Private Sub MAI1_Change()
If IsNumeric(UFAinput.MAI1) = False Then UFAinput.MAI1 = ""
End Sub
Private Sub MAI2_Change()
If IsNumeric(UFAinput.MAI2) = False Then UFAinput.MAI2 = ""
End Sub
Private Sub MAISpin_SpinUp()
Select Case UFAinput.MaiFlag
Case "2"
  UFAinput.MAI2.Visible = False
  UFAinput.COU2.Visible = False
  UFAinput.MAI2 = ""
  UFAinput.COU2 = ""
  UFAinput.MaiFlag = "1"
End Select
End Sub
Private Sub MAISpin_SpinDown()
Select Case UFAinput.MaiFlag
Case "1"
  UFAinput.MAI2.Visible = True
  UFAinput.COU2.Visible = True
  UFAinput.MaiFlag = "2"
End Select
End Sub
Private Sub Material1_Change()
If IsNumeric(UFAinput.Material1) = False Then UFAinput.Material1 = ""
End Sub
Private Sub Material2_Change()
If IsNumeric(UFAinput.Material2) = False Then UFAinput.Material2 = ""
End Sub
Private Sub Material3_Change()
If IsNumeric(UFAinput.Material3) = False Then UFAinput.Material3 = ""
End Sub
Private Sub MaterialSpin_SpinUp()
Select Case UFAinput.MaterialFlag
Case "3"
  UFAinput.Material3.Visible = False
  UFAinput.Material3 = ""
  UFAinput.MaterialFlag = "2"
Case "2"
  UFAinput.Material2.Visible = False
  UFAinput.Material2 = ""
  UFAinput.MaterialFlag = "1"
End Select
End Sub
Private Sub MaterialSpin_SpinDown()
Select Case UFAinput.MaterialFlag
Case "1"
  UFAinput.Material2.Visible = True
  UFAinput.MaterialFlag = "2"
Case "2"
  UFAinput.Material3.Visible = True
  UFAinput.MaterialFlag = "3"
End Select
End Sub

Private Sub ExMaterialSpin_SpinUp()
Select Case UFAinput.ExMaterialFlag
Case "3"
  UFAinput.ExMaterial3.Visible = False
  UFAinput.ExMaterial3 = ""
  UFAinput.ExMaterialFlag = "2"
Case "2"
  UFAinput.ExMaterial2.Visible = False
  UFAinput.ExMaterial2 = ""
  UFAinput.ExMaterialFlag = "1"
End Select
End Sub
Private Sub ExMaterialSpin_SpinDown()
Select Case UFAinput.ExMaterialFlag
Case "1"
  UFAinput.ExMaterial2.Visible = True
  UFAinput.ExMaterialFlag = "2"
Case "2"
  UFAinput.ExMaterial3.Visible = True
  UFAinput.ExMaterialFlag = "3"
End Select
End Sub

Private Sub OutsideAtu_Change()
If IsNumeric(UFAinput.OutsideAtu) = False Then UFAinput.OutsideAtu = ""
End Sub
Private Sub OutsideHaba_Change()
If IsNumeric(UFAinput.OutsideHaba) = False Then UFAinput.OutsideHaba = ""
End Sub
Private Sub OutsideTake_Change()
If IsNumeric(UFAinput.OutsideTake) = False Then UFAinput.OutsideTake = ""
End Sub

Private Sub PileMatch_Change()
Dim flag As Boolean: flag = False
Me.PileMatch.maxLength = 10
'ストアの有無***2023.2.15_丸徳ストカードNo間違い対策
  If Me.Remain.Text = "1" Then
  Me.PileMatch.maxLength = 7
    If Len(Me.PileMatch) <> 7 Then '7桁のストアNo
    Exit Sub
    Else
      Dim Store As String: Store = Left((Me.PileMatch.Text), 1)  'ストアカードNo頭
      Dim LotH As String:  LotH = Left((Me.Lot.Caption), 1) 'LotNoの頭
      If Store = LotH Then flag = True 'LotNoの頭とストアカードNoの頭が一致
      Select Case LotH 'S,Fは別扱い
      Case "S", "F"
        If Store = "L" Then flag = True
      End Select
      
'    MsgBox Len(Me.PileMatch)
    
    End If
    
    If flag = False Then
      MsgBox "ストアカードの頭文字が違います", vbCritical
      Me.PileMatch = ""
      Exit Sub
    End If
    
  ElseIf Me.Remain.Text = "0" And Me.PileMatchNo = "" Then
  Me.PileMatch.maxLength = 10
    Me.PileMatch = ""
  '  MsgBox "ストア、積合せ　どちらのフラグもありません", vbCritical
    Exit Sub
  End If
  

'積み合わせの有無
Select Case True
Case Me.PileMatchNo <> ""
  If IsNumeric(UFAinput.PileMatch) = False Then UFAinput.PileMatch = ""
Case Me.PileMatchNo = ""
End Select
End Sub

Private Sub PileMatch_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'入力制限
If Me.Remain = 1 Then
  Select Case KeyAscii.Value
  Case 48 To 57 '0~9
  Case 65 To 90 'A~Z
  Case 97 To 122 'a~z
  KeyAscii.Value = KeyAscii.Value - 32 '大文字強制変換
  Case Else
  KeyAscii.Value = 0
  End Select
End If
End Sub

Private Sub PileMatchNo_Change()
Me.PileMatch.maxLength = 10
If IsNumeric(UFAinput.PileMatchNo) = False Then UFAinput.PileMatchNo = ""
If Not (UFAinput.PileMatchNo = 9 Or UFAinput.PileMatchNo = 1) Then UFAinput.PileMatchNo = ""
If Not UFAinput.PileMatchNo = 1 Then UFAinput.PileMatch = ""
If UFAinput.PileMatchNo = 1 Then UFAinput.PileMatch = UFAinput.HikiNo & UFAinput.KONNO
End Sub

Private Sub Remain_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If Not (UFAinput.Remain = 0 Or UFAinput.Remain = 1) Then UFAinput.Remain = ""
End Sub

Private Sub Remain_Change()
If IsNumeric(UFAinput.Remain) = False Then UFAinput.Remain = ""
With UFAinput
  If .ExFrame.Visible = True Then 'EXとｽﾄｱの表示_2021.06.17
    Select Case .Remain
    Case 1 'EXｹｰｽNo、EXGW、EX風袋の入力が不要＊＊EX外寸の0,0,0は必要_2021.07.21
      .ExCaseNo.Visible = False
'      .ExMaterial1.Visible = False
'      .ExMaterial2.Visible = False
'      .ExMaterial3.Visible = False
'      .ExGW.Visible = False
'      .ExMaterialSpin.Visible = False
      .OutsideAtu = 0
      .OutsideHaba = 0
      .OutsideTake = 0
    Case 0
      .ExCaseNo.Visible = True
      .ExMaterial1.Visible = True
      .ExMaterial2.Visible = True
      .ExMaterial3.Visible = True
      .ExGW.Visible = True
      .ExMaterialSpin.Visible = True
      .OutsideAtu = ""
      .OutsideHaba = ""
      .OutsideTake = ""
    End Select
  End If
End With
End Sub
Private Sub Remain_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UFAinput.Inventory.Visible = True
UFAinput.Inventory.Caption = "高さ計算方法*出荷:引当通り_ｽﾄｱ:合紙有,用途高精度の時VC有"
End Sub

Private Sub Frame3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UFAinput.Inventory.Visible = False
End Sub

Private Sub Frame3_Click()
BlinkLabel1.BlinkStop '点滅ｽﾄｯﾌﾟ
' 動的ラベル上クリック時のクリア処理
Dim i As Long
Dim mc As Long: mc = val(UFAinput.Mount_Count.Caption)
For i = 7 To mc
    Dim lbl As MSForms.Label
    Set lbl = GetDynHeightLabel(i)
    If Not lbl Is Nothing Then
        ' MouseoverやClickは取れないためFrame3_Clickで全体クリアの案内のみ
        ' 個別クリアはラベルタグを利用した別途UI検討推奨
    End If
Next i
End Sub

Private Sub L1_Click()
On Error Resume Next
Dim L_height As Single: L_height = (val(UFAinput.ATU) * val(UFAinput.MAI1) * val(UFAinput.COU1)) / val(UFAinput.Mount_Count.Caption)
Select Case True
Case L_height <= 300#: UFAinput.L1.Caption = "Depth測定"  'Heightｹﾞｰｼﾞ使用可能
Case L_height > 300#: UFAinput.L1.Caption = "ｺﾝﾍﾞｯｸｽ測定"
End Select
BlinkLabel1.BlinkStop
UFAinput.L1.SetFocus
UFAinput.StartTB.SetFocus
End Sub

Public Function Height_Machin(Trm As String) As Boolean '測定機器選別 Heightｹﾞｰｼﾞ使用でtrue
Dim Temp As String: Temp = Maker(Trm)
Dim MachinNo As String
Select Case Temp
Case "DT"
 MachinNo = Mid(Trm, 3, 5) 'DT 10002 +0000021.712M
End Select

Dim SetAtu As Single
If Me.ATU <> "" Then SetAtu = val(Me.ATU)
If Me.ATU = "" Then SetAtu = 1#  '適当

Dim Lname As String:   Lname = Me.LineName.Caption
Dim Mdevice As String: Mdevice = DbReturnMicro(Lname, SetAtu, " ", "", "デプス") '機器指定ならライン名以外は引数を省ける

Select Case UFAinput.LineName
Case "L-1":    If MachinNo = Mdevice Then Height_Machin = True
Case "LVC":    If MachinNo = Mdevice Then Height_Machin = True
Case "HVC":    If MachinNo = Mdevice Then Height_Machin = True
Case "ﾊﾞﾗﾝｻｰ": If MachinNo = Mdevice Then Height_Machin = True
Case "ﾄｯﾄ":    If MachinNo = Mdevice Then Height_Machin = True
Case "Test":   If MachinNo = Mdevice Then Height_Machin = True
Case Else:     If MachinNo = "" Then Height_Machin = True
End Select
End Function

'
''使用高さ測定器選定　コンバート関数選定
'Private Sub StartTB_AfterUpdate()
'With UFAinput
'    If .StartTB.Tag <> "" Then Exit Sub
'    If .Mount_Count.Caption = "" Then
'        .StartTB.Tag = ""
'        Exit Sub
'    End If
'    .StartTB.Tag = "1"
'    Dim L_height As Single
'    L_height = (val(.Atu) * val(.MAI1) * val(.COU1)) / val(.Mount_Count.Caption)
'    Dim mc As Long: mc = val(.Mount_Count.Caption)
'    Dim idx As Long: idx = NextEmptyHeightIndex(mc)
'    If idx = -1 Then
'        .StartTB = ""
'        .StartTB.Tag = ""
'        .StartTB.SetFocus
'        Exit Sub
'    End If
'    ' 値チェックと変換
'    Dim valStr As String: valStr = ""
'    Select Case True
'    Case L_height <= 300#
'        If Height_Machin(.StartTB) = False Then
'            .StartTB = ""
'            .StartTB.Tag = ""
'            .StartTB.SetFocus
'            Exit Sub
'        End If
'        valStr = CStr(val(Conv(.StartTB)))
'        Call DebugLog("音楽対角呼び出し前 ActiveControl=" & UFAinput.ActiveControl.Name)
'        Call 音楽対角
'        Call DebugLog("音楽対角呼び出し後 ActiveControl=" & UFAinput.ActiveControl.Name)
'    Case L_height > 300#
'        If Maker(.StartTB) = "#L" Then
'            valStr = CStr(val(Conv(.StartTB)))
'            Call DebugLog("音楽対角呼び出し前 ActiveControl=" & UFAinput.ActiveControl.Name)
'            Call 音楽対角
'            Call DebugLog("音楽対角呼び出し後 ActiveControl=" & UFAinput.ActiveControl.Name)
'        Else
'            .StartTB = ""
'            .StartTB.Tag = ""
'            .StartTB.SetFocus
'            Exit Sub
'        End If
'    End Select
'    ' 動的ラベルに格納
'    Call SetKHeight(idx, valStr)
'    Call DebugLog("SetKHeight完了 idx=" & idx & " valStr=" & valStr)
'    .StartTB = ""
'    .StartTB.Tag = ""
'    AppActivate UFAinput.Caption
'    DoEvents
'    .StartTB.SetFocus
'    Call DebugLog("SetFocus呼び出し後 ActiveControl=" & UFAinput.ActiveControl.Name)
'End With
'
'End Sub


Private Sub StartTB_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode <> 9 Then Exit Sub  ' Tab以外は無視
    KeyCode = 0  ' Tabキーを消費してフォーカス移動を防ぐ
    ' CapsLock強制OFF
    If GetKeyState(vbKeyCapital) <> 0 Then
        keybd_event VK_CAPITAL, 0, 0, 0           ' キーダウン
        keybd_event VK_CAPITAL, 0, KEYEVENTF_KEYUP, 0 ' キーアップ
    End If
    
    With UFAinput
        If .StartTB.Tag <> "" Then Exit Sub
        If .Mount_Count.Caption = "" Then Exit Sub
        
        .StartTB.Tag = "1"
        Dim L_height As Single
        L_height = (val(.ATU) * val(.MAI1) * val(.COU1)) / val(.Mount_Count.Caption)
        Dim mc As Long: mc = val(.Mount_Count.Caption)
        Dim idx As Long: idx = NextEmptyHeightIndex(mc)
        If idx = -1 Then
            .StartTB = ""
            .StartTB.Tag = ""
            Exit Sub
        End If
        Dim valStr As String: valStr = ""
        Select Case True
        Case L_height <= 300#
            If Height_Machin(.StartTB) = False Then
                .StartTB = ""
                .StartTB.Tag = ""
                Exit Sub
            End If
            valStr = CStr(val(Conv(.StartTB)))
            Call 音楽対角
        Case L_height > 300#
            If Maker(.StartTB) = "#L" Then
                valStr = CStr(val(Conv(.StartTB)))
                Call 音楽対角
            Else
                .StartTB = ""
                .StartTB.Tag = ""
                Exit Sub
            End If
        End Select
        Call SetKHeight(idx, valStr)
        .StartTB = ""
        .StartTB.Tag = ""
        .AinButt.SetFocus    ' 一度別コントロールへ
        .StartTB.SetFocus    ' StartTBへ戻す
    End With
End Sub

Private Sub StartTB_Enter()
If Me.L2.Visible = True Then Me.L2.Caption = "測定できます。"
End Sub

Private Sub StartTB_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim mc As Long:    mc = val(UFAinput.Mount_Count.Caption)
    Dim filled As Long: filled = 0
    Dim i As Long
    
    ' 入力済み山数カウント
    For i = 1 To mc
        If GetKHeight(i) <> "" Then filled = filled + 1
    Next i
    
    If Me.L2.Visible = False Then Exit Sub
    
    Select Case True
    Case filled = 0
        Me.L2.Caption = "☆1測定の場合ｸﾘｯｸ"
    Case filled = mc
        Me.L2.Caption = "☆再測定の場合不要部消してからｸﾘｯｸ"
    Case filled < mc
        Me.L2.Caption = "☆" & (filled + 1) & "測定の場合ｸﾘｯｸ"
    End Select
End Sub

Private Sub UserForm_Click()
'Application.Visible = True
If UFAinput.Tag = "Load" Then Exit Sub
BlinkLabel1.BlinkStop '点滅ｽﾄｯﾌﾟ
'Me.L1.BackColor = &H0&
'Me.L1.ForeColor = &HC000&
UFAinput.AinButt.SetFocus
End Sub
Private Sub UserForm_Layout()
On Error Resume Next
Select Case True
Case ThisWorkbook.Sheets("キャスク検査").Visible
Case ThisWorkbook.Sheets("全ヵ所検査").Visible
Case Else
Application.Visible = False 'Excel非表示
' Excelウィンドウ位置を同期させる
'With Application
'  If .WindowState = xlNormal Then
'    .Top = Me.Top + 1
'    .Left = Me.Left + 5
'  End If
'End With
End Select
End Sub
Private Sub UserForm_Activate()
Call FormSetting

On Error Resume Next
UFAinput.HikiNo.clear '引当No初期化
Select Case True
Case ThisWorkbook.Sheets("キャスク検査").Visible
Case ThisWorkbook.Sheets("全ヵ所検査").Visible
Case Else
Application.Visible = False 'Excel非表示
End Select
Dim i As Long
If WsCheck("Hiki_Lot") = False Then Exit Sub

'ｺﾝﾎﾞﾎﾞｯｸｽ反映
' 一旦配列に格納してソート後にAddItem
Dim hikiSH As Worksheet: Set hikiSH = ThisWorkbook.Sheets("Hiki_Lot")
Dim lastRow As Long: lastRow = hikiSH.Cells(Rows.Count, 1).End(xlUp).row
Dim hikiArr() As String
Dim hikiCnt As Long: hikiCnt = 0

' 空白除外で配列に格納
For i = 1 To lastRow
    If hikiSH.Cells(i, 1) <> "" Then
        ReDim Preserve hikiArr(hikiCnt)
        hikiArr(hikiCnt) = CStr(hikiSH.Cells(i, 1))
        hikiCnt = hikiCnt + 1
    End If
Next i

' バブルソート（数値昇順）
Dim j As Long, tmpStr As String
For i = 0 To hikiCnt - 2
    For j = 0 To hikiCnt - 2 - i
        Dim isNumJ As Boolean:  isNumJ = IsNumeric(hikiArr(j))
        Dim isNumJ1 As Boolean: isNumJ1 = IsNumeric(hikiArr(j + 1))
        Dim doSwap As Boolean:  doSwap = False
        Select Case True
        Case isNumJ And isNumJ1:         doSwap = val(hikiArr(j)) > val(hikiArr(j + 1))
        Case isNumJ And Not isNumJ1:     doSwap = False
        Case Not isNumJ And isNumJ1:     doSwap = True
        Case Else:                       doSwap = hikiArr(j) > hikiArr(j + 1)
        End Select
        If doSwap Then
            tmpStr = hikiArr(j)
            hikiArr(j) = hikiArr(j + 1)
            hikiArr(j + 1) = tmpStr
        End If
    Next j
Next i

' ソート済みをAddItem
For i = 0 To hikiCnt - 1
    UFAinput.HikiNo.AddItem hikiArr(i)
Next i

'入力Box反映
Select Case UFAinput.KenLaw
Case "M2"
  UFAinput.MAILabel.Caption = "個装高さ"
  UFAinput.Material1.Visible = True
  UFAinput.GW.Visible = True
  UFAinput.MaterialLabel.Visible = True
  UFAinput.GWLabel.Visible = True
Case "M6"
  UFAinput.MAILabel.Caption = "梱包枚数"
  UFAinput.Material1.Visible = True
  UFAinput.GW.Visible = True
  UFAinput.MaterialLabel.Visible = True
  UFAinput.GWLabel.Visible = True
Case Else
  UFAinput.MAILabel.Caption = "梱包枚数"
  UFAinput.Material1.Visible = False
  UFAinput.Material2.Visible = False
  UFAinput.Material3.Visible = False
  UFAinput.GW.Visible = False
  UFAinput.MaterialLabel.Visible = False
  UFAinput.GWLabel.Visible = False
  UFAinput.MaterialSpin.Visible = False
End Select
Call WindowMenuButtonOn(Me) '2021.03.25
End Sub
Private Sub UserForm_Initialize()
If WsCheck("Lot情報履歴") = False Then End

Set BlinkLabel1 = New clsBlinkLabel '点滅ラベルの登録
BlinkLabel1.cmd = Me.L1

Call GetZoom(UFAinput)
Dim Temp As Variant:  Temp = InitialSetting(UFAinput)

Select Case True
Case ThisWorkbook.Sheets("キャスク検査").Visible
Case ThisWorkbook.Sheets("全ヵ所検査").Visible
End Select
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Me.GwOpen.Visible = True
Me.Re_hostsend.Visible = True
Me.OutCondition.Visible = True
Me.GwOpen.Caption = ""
Me.Re_hostsend.Caption = ""
Me.OutCondition.Caption = ""
Me.LoadTimeText.Caption = ""
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '×では閉じないようにする
If UFAinput.Tag = "Load" Then Exit Sub
On Error Resume Next
If UFMsgBox.Visible = True Then Unload UFMsgBox
Select Case True
Case ThisWorkbook.Sheets("キャスク検査").Visible
  UFCask.Show vbModeless
Case ThisWorkbook.Sheets("全ヵ所検査").Visible, ThisWorkbook.Sheets("H494検査").Visible
  UFMitsubishi.Show vbModeless
Case Else
'  With Application
'  .Top = g_sngTopSV
'  .Left = g_sngLeftSV
'  .Width = g_sngWidthSV
'  .Height = g_sngHeightSV
'  End With
  UF1.Show vbModeless
End Select
End Sub

Private Sub AinButt_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Me.Re_hostsend.Visible = False
Me.OutCondition.Visible = True
Me.GwOpen.Visible = False
Me.OutCondition.Caption = "全て高さCheck要" '"梱包枚数が1枚　or　板厚が100mm以上の場合高さﾁｪｯｸ不要"
End Sub

Private Sub AinButt_Click()
Dim Ainput As Variant
If UFAinput.DBflag.Caption = "true" Then
  MsgBox "Lot情報履歴に履歴が無いため" & vbCrLf & "A_input転送は使用できません。" & vbCrLf & "_AinButt:DBflag", vbExclamation
  Exit Sub
End If
If UFAinput.HikiNo = "" Then
  MsgBox "必要項目の入力がありません" & "(引当No)", vbCritical
  Exit Sub '高さ測定山数入力無し
End If

If UFAinput.Tag = "Load" Then Exit Sub
'クリア
  UFAinput.AiText = ""
  UFAinput.GaText = ""
  UFAinput.VcText = ""
UFAinput.Tag = "Load" '*******************************これ以降Exitする際はtag=""必須_2021.05.20
'1度目のﾎﾞﾀﾝｸﾘｯｸで山数入力値から伝送高さﾎﾞｯｸｽを表示
Dim Workers As String, path As String: path = HistoryPath
With UFAinput
  '枚数
  Dim M_Count As Long
  Select Case .MaiFlag.Caption '包み枚数の表示ﾌﾗｸﾞ
  Case "1": M_Count = val(.MAI1) * val(.COU1)
  Case "2": M_Count = (val(.MAI1) * val(.COU1)) + (val(.MAI2) * val(.COU2))
  End Select
  
  
  
  '山高さ'全山高さ加算
  Dim Mount_Height As Single: Mount_Height = CalcMountHeight(val(.Mount_Count.Caption)) '_2026.05.29
  Dim ainMc As Long: ainMc = val(.Mount_Count.Caption)

 '--------------------高さﾁｪｯｸなし*2021.04.09_廃止、すべて高さ測定を行う
  
 '--------------------高さﾁｪｯｸあり
  Select Case True
  Case .L1.Visible And AllHeightFilled(ainMc) '伝送高さ測定ﾗﾍﾞﾙ表示あり 'AllHeightFilled 全山入力済みで True を返す_2026.05.29
      '入力チェック
      If GetKHeight(1) = "" Or .Mount_Count = "" Or .HikiNo = "" Then
        .Tag = ""
        MsgBox "必要項目の入力がありません" & "(高さ測定値,山数,引当No)", vbCritical
        Exit Sub '高さ測定山数入力無し
      End If
      Select Case .MaiFlag.Caption '包み枚数の表示ﾌﾗｸﾞ
      Case "1"
        If .MAI1 = "" Or .COU1 = "" Then
          .Tag = ""
          MsgBox "必要項目の入力がありません" & "(枚数1,包数1)", vbCritical
          Exit Sub   '入力無し
        End If
      Case "2"
        If .MAI1 = "" Or .COU1 = "" Or .MAI2 = "" Or .COU2 = "" Then
          .Tag = ""
          MsgBox "必要項目の入力がありません" & "(枚数1,包数1,枚数2,包数2)", vbCritical
          Exit Sub   '入力無し
        End If
      End Select
      
      If UFAinput.ExFrame.Visible Then 'EXの時_2021.05.10
        If UFAinput.Remain <> 1 Then '2021.06.17_ｽﾄｱの時はチェック不要
          If UFAinput.ExCaseNo = "" Then '2021.05.20_EX風袋EXGWの入力ﾁｪｯｸは一旦なし
            MsgBox "ｹｰｽNoの入力がありません。", vbExclamation
            UFAinput.Tag = ""
            Exit Sub
          End If
        End If
        Dim FlEX1 As Boolean, FlEX2 As Boolean
        Select Case UFAinput.OutsideAtu '入力数値ﾁｪｯｸ_.textをつけると文字列扱い　ここでは数列チェック
        Case 1 To 9: FlEX1 = True
        End Select
        Select Case UFAinput.OutsideHaba
        Case 1 To 9: FlEX1 = True
        End Select
        Select Case UFAinput.OutsideTake
        Case 1 To 9: FlEX1 = True
        End Select
        
        Select Case True '入力ﾁｪｯｸ
        Case (UFAinput.OutsideAtu = "" Or UFAinput.OutsideHaba = "" Or UFAinput.OutsideTake = ""): FlEX2 = True
        End Select
        If FlEX1 Then 'ﾌﾗｸﾞﾁｪｯｸ
          MsgBox "EX外寸に在り得ない数値の入力があります。", vbExclamation
          UFAinput.Tag = ""
          Exit Sub
        End If
        If FlEX2 Then
          MsgBox "EX外寸の入力がありません｡", vbExclamation
          UFAinput.Tag = ""
          Exit Sub
        End If
      End If
       
      If UFAinput.Remain = "1" Then 'ｽﾄｱ材の時_2022.02.04
        If UFAinput.PileMatch = "" Then
          MsgBox "ｽﾄｱｶｰﾄﾞNo入力がありません。", vbExclamation  'ｽﾄｱNo入力なし
          UFAinput.Tag = ""
          Exit Sub
        End If
      End If
      
      If Height_Check(val(.AMin.Caption), val(.AMax.Caption), val(Mount_Height), val(M_Count), .OrderNo.Caption) = False Then '高さ許容値ﾁｪｯｸ(ATU1 As Single, ATU2 As Single, Hei As Single, MAI As Long, OrderNo As String)
        If MsgBox("ｱｲｶｽ転送を行いますか？", vbYesNo) = vbYes Then
          Select Case WorkeDayMax '登録時間が一番遅い
          Case "1": Workers = " T1:" & UF1.T1.Text
          Case "2": Workers = " T2:" & UF1.T2.Text
          Case "3": Workers = " T3:" & UF1.T3.Text
          End Select
          Select Case UFAinput.LineName
          Case "L-1":    Call WriteLog("*UFKen:" & M_Count & "枚_" & Mount_Height & "mm *" & Workers, "", path & "L-1\L-1_Log")
          Case "LVC":    Call WriteLog("*UFKen:" & M_Count & "枚_" & Mount_Height & "mm *" & Workers, "", path & "LVC\LVC_Log")
          Case "HVC":    Call WriteLog("*UFKen:" & M_Count & "枚_" & Mount_Height & "mm *" & Workers, "", path & "HVC\HVC_Log")
          Case "ﾊﾞﾗﾝｻｰ": Call WriteLog("*UFKen:" & M_Count & "枚_" & Mount_Height & "mm *" & Workers, "", path & "バランサー\バランサー_Log")
          Case "ﾄｯﾄ":    Call WriteLog("*UFKen:" & M_Count & "枚_" & Mount_Height & "mm *" & Workers, "", path & "トット\トット_Log")
          Case "Test":   Call WriteLog("*UFKen:" & M_Count & "枚_" & Mount_Height & "mm *" & Workers, "", path & "Test\Test_Log")
          Case Else:     Call WriteLog("*UFKen:" & M_Count & "枚\" & Mount_Height & "mm *" & Workers, "", path & "丸徳\丸徳_Log")
          End Select
        
          Call MusicTrue("お待ちください.wav")
          UFProgress.Show vbModeless
          UFProgress.CommandButton1.Caption = "Data送信中..."
          Call ProgressBar_count(1, 350)
          DoEvents
  '        Call Daily_Report4_Execution '梱包No　終了時間　枚数　次開始時間__*2021.03.08一旦停止
  '         If UFAinput.Ngflag = "stop" Then '連番でないとき
  '            UFAinput.Tag = ""
  '            Exit Sub
  '         End If
  
          Call A_Output(UFAinput, .Lot, .KenNo) '格納値Output
          
          
          Call ProgressBar_count(350, 750)
          DoEvents
          Call App_Activation 'ｱﾌﾟﾘ起動
          If UFAinput.Ngflag = "stop" Then
             Unload UFProgress
             UFAinput.Tag = ""
             UFAinput.Repaint
             Exit Sub
          End If
          Sleep 1000
          Call InterRuptText(, , 5000) '結果読み込み
          If UFAinput.Ngflag = "stop" Then
             Unload UFProgress
             UFAinput.Tag = ""
             UFAinput.Repaint
             Exit Sub
          End If
          
          Call Daily_Report(4) '日報登録
          Call ProgressBar_count(750, 1000)
          DoEvents
          Unload UFProgress
          Call MusicTrue("完了.wav")
          .K_Height = "" 'ﾘｾｯﾄ
          .K_Height_2 = ""
          .K_Height_3 = ""
          .K_Height_4 = ""
          .KONNO = ""
        Else
          .Tag = ""
          Exit Sub
        End If
      Else 'Height_Check=Falseのときunload処理の為不要
        Exit Sub
      End If
  Case Else '*******************伝送高さ測定ﾗﾍﾞﾙ表示なし
      If .Mount_Count.Caption <> "" Then
      .L1.Visible = True
      .L2.Visible = True
      .L1.Font.Size = 8
      
        Dim L_height As Single: L_height = (val(UFAinput.ATU) * val(UFAinput.MAI1) * val(UFAinput.COU1)) / val(UFAinput.Mount_Count.Caption)
        Select Case True
        Case L_height <= 300#: .L1.Caption = "Depth測定"  'Heightｹﾞｰｼﾞ使用可能
        Case L_height > 300#:   .L1.Caption = "ｺﾝﾍﾞｯｸｽ測定"
        End Select
        
      '冗長部分修正_2026.05.29
        Dim mc As Long: mc = val(.Mount_Count.Caption)
        If mc >= 1 And mc <= MAX_MOUNT Then
            Call ShowHeightControls(mc)
  '          BlinkLabel1.Blink 3
            .L1.BackColor = RGB(0, 0, 0)
            .L1.ForeColor = &HC000&
            .StartTB.SetFocus
        Else
            .L1.Visible = False
            .L2.Visible = False
        End If
        
      Else
       .Tag = ""
       MsgBox "山数入力無し", vbCritical
       Exit Sub
      End If
  End Select
End With
UFAinput.Tag = ""
'UFAinput.GaityuuText = ""
UFAinput.Ai.Locked = True
UFAinput.Ga.Locked = True
UFAinput.Vab.Locked = True
'UFAinput.GaityuuText.Visible = False
End Sub


'動的生成ラベルにすべて変更_2026.05.29　イベントすべてコメントアウト

'Private Sub K_Height_Label_Click()
'Me.K_Height = ""
'End Sub
'Private Sub K_Height_2_Label_Click()
'Me.K_Height_2 = ""
'End Sub
'Private Sub K_Height_3_Label_Click()
'Me.K_Height_3 = ""
'End Sub
'Private Sub K_Height_4_Label_Click()
'Me.K_Height_4 = ""
'End Sub
'Private Sub K_Height_5_Label_Click()
'Me.K_Height_5 = ""
'End Sub
'Private Sub K_Height_6_Label_Click()
'Me.K_Height_6 = ""
'End Sub

'Private Sub K_Height_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'UFAinput.K_Height.Tag = ""
'UFAinput.K_Height.SetFocus
'UFAinput.StartTB.SetFocus
'End Sub
'Private Sub K_Height_2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'UFAinput.K_Height.Tag = ""
'UFAinput.K_Height_2.SetFocus
'UFAinput.StartTB.SetFocus
'End Sub
'Private Sub K_Height_3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'UFAinput.K_Height.Tag = ""
'UFAinput.K_Height_3.SetFocus
'UFAinput.StartTB.SetFocus
'End Sub
'Private Sub K_Height_4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'UFAinput.K_Height.Tag = ""
'UFAinput.K_Height_4.SetFocus
'UFAinput.StartTB.SetFocus
'End Sub
'Private Sub K_Height_5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'UFAinput.K_Height.Tag = ""
'UFAinput.K_Height_5.SetFocus
'UFAinput.StartTB.SetFocus
'End Sub
'Private Sub K_Height_6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'UFAinput.K_Height.Tag = ""
'UFAinput.K_Height_6.SetFocus
'UFAinput.StartTB.SetFocus
'End Sub

''2021.02.17_ Heightｹﾞｰｼﾞﾁｪｯｸ追加_2021.03.09_Changeに変更　StartTBから送る
'Private Sub K_Height_Change() 'Change_AfterUpdate
'If UFAinput.K_Height.Tag <> "" Then Exit Sub
'On Error Resume Next
'UFAinput.K_Height.Tag = "1"
'Dim L_height As Single: L_height = (val(UFAinput.Atu) * val(UFAinput.MAI1) * val(UFAinput.COU1)) / val(UFAinput.Mount_Count.Caption)
'Select Case True
'Case L_height <= 300#  'Heightｹﾞｰｼﾞ使用可能
'  If Height_Machin(UFAinput.K_Height) = False Then UFAinput.K_Height = ""
'  If Height_Machin(UFAinput.K_Height) Then
'    UFAinput.K_Height = val(Conv(UFAinput.K_Height))
'    Call 音楽対角
'  End If
'Case L_height > 300#
'  If Maker(UFAinput.K_Height) = "#L" Then
'    UFAinput.K_Height = val(Conv(UFAinput.K_Height))
'    Call 音楽対角
'  End If
'End Select
'UFAinput.K_Height_Label.Caption = UFAinput.K_Height
'UFAinput.K_Height.Tag = ""
'UFAinput.K_Height.SetFocus
'UFAinput.StartTB.SetFocus
'End Sub
'Private Sub K_Height_2_Change()
'If UFAinput.K_Height_2.Tag <> "" Then Exit Sub
'On Error Resume Next
'UFAinput.K_Height_2.Tag = "1"
'Dim L_height As Single: L_height = (val(UFAinput.Atu) * val(UFAinput.MAI1) * val(UFAinput.COU1)) / val(UFAinput.Mount_Count.Caption)
'Select Case True
'Case L_height <= 300#  'Heightｹﾞｰｼﾞ使用可能
'  If Height_Machin(UFAinput.K_Height_2) = False Then UFAinput.K_Height_2 = ""
'  If Height_Machin(UFAinput.K_Height_2) Then
'    UFAinput.K_Height_2 = val(Conv(UFAinput.K_Height_2))
'    Call 音楽対角
'  End If
'Case L_height > 300#
'  If Maker(UFAinput.K_Height_2) = "#L" Then
'    UFAinput.K_Height_2 = val(Conv(UFAinput.K_Height_2))
'    Call 音楽対角
'  End If
'End Select
'UFAinput.K_Height_2_Label.Caption = UFAinput.K_Height_2
'UFAinput.K_Height_2.Tag = ""
'UFAinput.K_Height.SetFocus
'UFAinput.StartTB.SetFocus
'End Sub
'Private Sub K_Height_3_Change()
'If UFAinput.K_Height_3.Tag <> "" Then Exit Sub
'On Error Resume Next
'UFAinput.K_Height_3.Tag = "1"
'Dim L_height As Single: L_height = (val(UFAinput.Atu) * val(UFAinput.MAI1) * val(UFAinput.COU1)) / val(UFAinput.Mount_Count.Caption)
'Select Case True
'Case L_height <= 300#  'Heightｹﾞｰｼﾞ使用可能
'  If Height_Machin(UFAinput.K_Height_3) = False Then UFAinput.K_Height_3 = ""
'  If Height_Machin(UFAinput.K_Height_3) Then
'    UFAinput.K_Height_3 = val(Conv(UFAinput.K_Height_3))
'    Call 音楽対角
'  End If
'Case L_height > 300#
'  If Maker(UFAinput.K_Height_3) = "#L" Then
'    UFAinput.K_Height_3 = val(Conv(UFAinput.K_Height_3))
'    Call 音楽対角
'  End If
'End Select
'UFAinput.K_Height_3_Label.Caption = UFAinput.K_Height_3
'UFAinput.K_Height_3.Tag = ""
'UFAinput.K_Height.SetFocus
'UFAinput.StartTB.SetFocus
'End Sub
'Private Sub K_Height_4_Change()
'If UFAinput.K_Height_4.Tag <> "" Then Exit Sub
'On Error Resume Next
'UFAinput.K_Height_4.Tag = "1"
'Dim L_height As Single: L_height = (val(UFAinput.Atu) * val(UFAinput.MAI1) * val(UFAinput.COU1)) / val(UFAinput.Mount_Count.Caption)
'Select Case True
'Case L_height <= 300#  'Heightｹﾞｰｼﾞ使用可能
'  If Height_Machin(UFAinput.K_Height_4) = False Then UFAinput.K_Height_4 = ""
'  If Height_Machin(UFAinput.K_Height_4) Then
'    UFAinput.K_Height_4 = val(Conv(UFAinput.K_Height_4))
'    Call 音楽対角
'  End If
'Case L_height > 300#
'  If Maker(UFAinput.K_Height_4) = "#L" Then
'    UFAinput.K_Height_4 = val(Conv(UFAinput.K_Height_4))
'    Call 音楽対角
'  End If
'End Select
'UFAinput.K_Height_4_Label.Caption = UFAinput.K_Height_4
'UFAinput.K_Height_4.Tag = ""
'UFAinput.K_Height.SetFocus
'UFAinput.StartTB.SetFocus
'End Sub
'Private Sub K_Height_5_Change()
'If UFAinput.K_Height_5.Tag <> "" Then Exit Sub
'On Error Resume Next
'UFAinput.K_Height_5.Tag = "1"
'Dim L_height As Single: L_height = (val(UFAinput.Atu) * val(UFAinput.MAI1) * val(UFAinput.COU1)) / val(UFAinput.Mount_Count.Caption)
'Select Case True
'Case L_height <= 300#  'Heightｹﾞｰｼﾞ使用可能
'  If Height_Machin(UFAinput.K_Height_5) = False Then UFAinput.K_Height_5 = ""
'  If Height_Machin(UFAinput.K_Height_5) Then
'    UFAinput.K_Height_5 = val(Conv(UFAinput.K_Height_5))
'    Call 音楽対角
'  End If
'Case L_height > 300#
'  If Maker(UFAinput.K_Height_5) = "#L" Then
'    UFAinput.K_Height_5 = val(Conv(UFAinput.K_Height_5))
'    Call 音楽対角
'  End If
'End Select
'UFAinput.K_Height_5_Label.Caption = UFAinput.K_Height_5
'UFAinput.K_Height_5.Tag = ""
'UFAinput.K_Height.SetFocus
'UFAinput.StartTB.SetFocus
'End Sub
'Private Sub K_Height_6_Change()
'If UFAinput.K_Height_6.Tag <> "" Then Exit Sub
'On Error Resume Next
'UFAinput.K_Height_6.Tag = "1"
'Dim L_height As Single: L_height = (val(UFAinput.Atu) * val(UFAinput.MAI1) * val(UFAinput.COU1)) / val(UFAinput.Mount_Count.Caption)
'Select Case True
'Case L_height <= 300#  'Heightｹﾞｰｼﾞ使用可能
'  If Height_Machin(UFAinput.K_Height_6) = False Then UFAinput.K_Height_6 = ""
'  If Height_Machin(UFAinput.K_Height_6) Then
'    UFAinput.K_Height_6 = val(Conv(UFAinput.K_Height_6))
'    Call 音楽対角
'  End If
'Case L_height > 300#
'  If Maker(UFAinput.K_Height_6) = "#L" Then
'    UFAinput.K_Height_6 = val(Conv(UFAinput.K_Height_6))
'    Call 音楽対角
'  End If
'End Select
'UFAinput.K_Height_6_Label.Caption = UFAinput.K_Height_6
'UFAinput.K_Height_6.Tag = ""
'UFAinput.K_Height.SetFocus
'UFAinput.StartTB.SetFocus
'End Sub

Private Sub Mount_Count_Click()
If UFAinput.Tag = "Load" Then Exit Sub
' 動的ラベル更新_2026.05.29
Call Number_Inp(Mount_Count)
End Sub

' 動的ラベル（1～15）の表示/非表示切り替えとフォームリサイズ
' mountCount: Mount_Count.Captionの値
Public Sub RefreshDynHeightLabels(mountCount As Long)
    Const LBL_TOP        As Single = 162
    Const LBL_HEIGHT     As Single = 25.5
    Const LBL_WIDTH      As Single = 66
    Const LBL_LEFT_START As Single = 66
    Const LEFT_STEP      As Single = 66

    Call DebugLog("RefreshDynHeightLabels 開始 mountCount=" & mountCount)
'    Call DebugLog("Frame3 Top=" & Me.Frame3.Top & " Left=" & Me.Frame3.Left & " Width=" & Me.Frame3.Width & " Height=" & Me.Frame3.Height)
'    Call DebugLog("フォーム InsideWidth=" & Me.InsideWidth & " InsideHeight=" & Me.InsideHeight)
'
'    Call DebugLog("K_Height_Label Visible=" & Me.K_Height_Label.Visible & " Top=" & Me.K_Height_Label.Top & " Left=" & Me.K_Height_Label.Left)
    
    Dim i As Long
    Dim lbl As MSForms.Label
    Dim targetName As String
    Dim ctrl As MSForms.Control

    ' 既存TB全非表示（全動的ラベル化のため）使用しないが影響考えるのめんどいから残す
    Me.K_Height.Visible = False
    Me.K_Height_2.Visible = False
    Me.K_Height_3.Visible = False
    Me.K_Height_4.Visible = False
    Me.K_Height_5.Visible = False
    Me.K_Height_6.Visible = False

    ' 動的ラベル全非表示 同名ラベルがあるか探す 既存の DynHL_ ラベルを全部非表示
    For Each ctrl In Me.Frame3.Controls
        If Left(ctrl.Name, 6) = "DynHL_" Then
            ctrl.Visible = False
        End If
    Next ctrl

    If mountCount >= 1 Then
        For i = 1 To mountCount
            targetName = "DynHL_" & i
            Set lbl = Nothing

            For Each ctrl In Me.Frame3.Controls
                If ctrl.Name = targetName Then
                    Set lbl = ctrl
'                    Call DebugLog("既存ラベル発見: " & targetName)
                    Exit For
                End If
            Next ctrl
'なければ 新規生成
            If lbl Is Nothing Then
                Call DebugLog("ラベル新規生成開始: " & targetName)
                On Error GoTo LblError
                Set lbl = Me.Frame3.Controls.Add("Forms.Label.1", targetName, True)
                On Error GoTo 0
'                Call DebugLog("ラベル生成成功: " & targetName)
                With lbl
                    .tOP = LBL_TOP
                    .Left = LBL_LEFT_START + (i - 1) * LEFT_STEP
                    .Height = LBL_HEIGHT
                    .Width = LBL_WIDTH
                    .Caption = ""
                    .Tag = ""
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = &H0&
                    .ForeColor = &HC000&
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = &H80000006
                    .Font.Name = "Meiryo UI"
                    .Font.Size = 15
                    .TextAlign = fmTextAlignCenter
                    .Visible = True
                End With
'                Call DebugLog("ラベル色確認: BackColor=" & lbl.BackColor & " BorderColor=" & lbl.BorderColor & " BorderStyle=" & lbl.BorderStyle)
'                Call DebugLog("ラベルプロパティ設定完了: Top=" & lbl.Top & " Left=" & lbl.Left)
            Else
                lbl.Visible = True
'                Call DebugLog("既存ラベル再表示: " & targetName)
            End If

            ' イベントバインド
            On Error GoTo BindError
            Set DynLabelEvents(i) = New clsDynLabel
            DynLabelEvents(i).Init lbl, i
            On Error GoTo 0
            Call DebugLog("イベントバインド完了: index=" & i)
        Next i
'    Else
'        Call DebugLog("mountCount<=6 の為動的ラベル生成スキップ")
    End If

    ' フォーム横幅調整_2026.05.29' ラベル数から必要幅を直接計算
    ' Frame3内：左余白66 + ラベル幅66 * 山数 + 右余白12
    Dim newFrameWidth As Single
    Dim newFormWidth  As Single
    
    '山数が少なくても Frame3.Width は最低 546
    newFrameWidth = Application.WorksheetFunction.Max(546, LBL_LEFT_START + (LBL_WIDTH * mountCount) + 12)
    ' フォーム幅：Frame3のLeft18 + Frame3幅 + 右余白12
    newFormWidth = Me.Frame3.Left + newFrameWidth + 12

    Me.Frame3.Width = newFrameWidth
    DoEvents
    ' 最低サイズ保証
    If Me.Width < 592.5 Then Me.Width = 592.5
    If Me.Height < 568.5 Then Me.Height = 568.5
    Call ResizeWindow(Me.Frame3.Left + newFrameWidth + 30)
    
    ' InsideWidthとMe.Widthの差分を計算してオフセットとして使う
    Dim widthOffset As Single: widthOffset = Me.Width - Me.InsideWidth
    Dim targetInsideWidth As Single: targetInsideWidth = Me.Frame3.Left + newFrameWidth + 60
    Dim targetWidthPt As Single:     targetWidthPt = targetInsideWidth + widthOffset
    Call ResizeWindow(targetWidthPt)
    DoEvents
    
    Me.Width = targetWidthPt
    
    Dim dpi As Long: dpi = GetDPI()
    Me.Left = (GetSystemMetrics32(0) * 72# / dpi - Me.Width) / 2
    Me.tOP = (GetSystemMetrics32(1) * 72# / dpi - Me.Height) / 2
    DoEvents
    Call DebugLog("フォーム幅設定完了: Width=" & Me.Width & " Frame3.Width=" & Me.Frame3.Width & " InsideWidth=" & Me.InsideWidth)
    Call DebugLog("実ウィンドウ幅(px)=" & GetActualWindowWidth())
    Call DebugLog("RefreshDynHeightLabels 終了")
    
    Exit Sub

LblError:
    Call DebugLog("ラベル生成エラー: " & targetName & " / " & Err.Number & ":" & Err.Description)
    On Error GoTo 0
    Exit Sub

BindError:
    Call DebugLog("バインドエラー: index=" & i & " / " & Err.Number & ":" & Err.Description)
    On Error GoTo 0
End Sub

Public Sub ShowHeightControls(mountCount As Long)
    With Me
        ' 既存TBは全非表示（全動的ラベル化）
        .K_Height.Visible = False
        .K_Height_2.Visible = False
        .K_Height_3.Visible = False
        .K_Height_4.Visible = False
        .K_Height_5.Visible = False
        .K_Height_6.Visible = False
        Dim lh As Single
        lh = (val(.ATU) * val(.MAI1) * val(.COU1)) / val(.Mount_Count.Caption)
        Select Case True
        Case lh <= 300#: .L1.Caption = "Depth測定"
        Case lh > 300#:  .L1.Caption = "ｺﾝﾍﾞｯｸｽ測定"
        Case Else:       .L1.Caption = "伝送高さ測定"
        End Select
        .L1.Visible = True
        .L2.Visible = True
        .L1.Font.Size = 8
        .L1.BackColor = RGB(255, 0, 0)
        .L1.ForeColor = RGB(255, 255, 255)
    End With
    BlinkLabel1.Blink 3
    Call RefreshDynHeightLabels(mountCount)
End Sub
