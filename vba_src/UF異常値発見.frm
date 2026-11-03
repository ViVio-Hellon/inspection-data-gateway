Attribute VB_Name = "UF異常値発見"
Attribute VB_Base = "0{94379F2F-026B-4D4F-A8D2-F700C210028C}{C2E64533-0D1D-4F1E-8AE0-C2467EBD40D9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit
' Excelウィンドウ状態退避用
'Private g_sngTopSV As Single ' 上端位置
'Private g_sngLeftSV As Single ' 左端位置
'Private g_sngWidthSV As Single ' 幅
'Private g_sngHeightSV As Single ' 高さ
Private g_lngWindowState As XlWindowState                       ' ウィンドウ状態

Private Sub UserForm_Layout()
' Excelウィンドウ位置を同期させる
On Error Resume Next
'With Application
'  If .WindowState = xlNormal Then
'    .Top = Me.Top + 1
'    .Left = Me.Left + 1
'  End If
'End With
End Sub
Private Sub UserForm_Activate()
Call 音楽異常値
Call WindowMenuButtonDel(UF異常値発見)
' フォームの後ろにExcelウィンドウを隠す
On Error Resume Next
'With Application
'  .Top = Me.Top + 1
'  .Left = Me.Left + 1
'  .Width = Me.Width - 2
'  .Height = Me.Height - 2
'End With
End Sub
Private Sub UserForm_Initialize()
' 現在のExcelウィンドウ状態を保管
On Error Resume Next
' Excelウィンドウを最小化
'Application.WindowState = xlMinimized
'AppActivate Application.Caption
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("作業保存" & "_" & Lname) = False Then Exit Sub '-------------------------ｼｰﾄなし
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("作業保存" & "_" & Lname)
Dim ShRange As Range: Set ShRange = Range(sh.Cells(1, 2), sh.Cells(1, UC(sh, 1)))
Dim SaveDate(): SaveDate = ShRange
Dim ATU As Double '厚
Dim HABA As Single '巾
Dim TAKE As Single '丈
Dim i As Long, a As Long
Dim Am As Double: Am = val(SaveDate(1, 4)) * 2
Dim Ap As Double: Ap = val(SaveDate(1, 5)) * 2
Dim Hm As Double: Hm = val(SaveDate(1, 6)) * 2
Dim Hp As Double: Hp = val(SaveDate(1, 7)) * 2
Dim Tm As Double: Tm = val(SaveDate(1, 8)) * 2
Dim Tp As Double: Tp = val(SaveDate(1, 9)) * 2
Dim iatu1 As Double, ihaba1 As Double, itake1 As Double, itai1 As Double '-
Dim iatu2 As Double, ihaba2 As Double, itake2 As Double, itai2 As Double '+
ATU = val(SaveDate(1, 1))
HABA = val(SaveDate(1, 2))
TAKE = val(SaveDate(1, 3))
'****************************************************************************************************
'異常値の定義　2018.04.14　公差の倍以上、倍以下とする
   UF異常値発見.TEIGI = "異常値の閾値　公差の2.0倍とする"
   UF異常値発見.TEIGITAI = "対角線　a^2＋b^2±6で比較"
   iatu1 = Format(ATU - Am, "0.000")
   iatu2 = Format(ATU + Ap, "0.000")
   UF異常値発見.ATU = "厚" & " " & Format(iatu1, "0.000") & "未満" & " " & Format(iatu2, "0.000") & "超え"
   ihaba1 = Format(HABA - Hm, "0.0")
   ihaba2 = Format(HABA + Hp, "0.0")
   UF異常値発見.HABA = "巾" & " " & Format(ihaba1, "0.0") & "未満" & " " & Format(ihaba2, "0.0") & "超え"
   itake1 = Format(TAKE - Tm, "0.0")
   itake2 = Format(TAKE + Tp, "0.0")
   UF異常値発見.TAKE = "丈" & " " & Format(itake1, "0.0") & "未満" & " " & Format(itake2, "0.0") & "超え"
 '3平方の定理　a^2＋b^2＝c^2  Sqr関数＝数値の平方根を返す
   itai1 = Format(Sqr(val(HABA) ^ 2 + val(TAKE) ^ 2) - 6, "0.0") '-
   itai2 = Format(Sqr(val(HABA) ^ 2 + val(TAKE) ^ 2) + 6, "0.0") '+
   UF異常値発見.TAIK = "対角線" & " " & Format(itai1, "0.0") & "未満" & " " & Format(itai2, "0.0") & "超え"
'****************************************************************************************************

Call GetZoom(Me)
Dim Temp As Variant:  Temp = InitialSetting(Me)

'With Application
'  .WindowState = xlNormal
'  g_sngTopSV = .Top
'  g_sngLeftSV = .Left
'  g_sngWidthSV = .Width
'  g_sngHeightSV = .Height
'End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '×では閉じないようにする
If CloseMode = vbFormControlMenu Then
  Cancel = True
End If
' 保管情報を元にExcelウィンドウを元に戻す
On Error Resume Next
UF1.Show vbModeless
DoEvents
'With Application
'  .Top = g_sngTopSV
'  .Left = g_sngLeftSV
'  .Width = g_sngWidthSV
'  .Height = g_sngHeightSV
'End With
End Sub
Private Sub CommandButton1_Click()
Dim st As String: st = UF異常値発見.TBname
Unload Me
Call NG_Back_TB(st)
End Sub

