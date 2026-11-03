Attribute VB_Name = "UFMsgBox"
Attribute VB_Base = "0{177B3272-DC62-48D9-BE14-55E3F394BF5D}{CA7A40E2-C677-4265-84E7-3D03B2CB5634}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit
' Excelウィンドウ状態退避用
Private g_sngTopSV As Single ' 上端位置
Private g_sngLeftSV As Single ' 左端位置
Private g_sngWidthSV As Single ' 幅
Private g_sngHeightSV As Single ' 高さ
Private Sub UserForm_Layout()
'With Application ' Excelウィンドウ位置を同期させる
'  If .WindowState = xlNormal Then
'    .Top = Me.Top + 1
'    .Left = Me.Left + 1
'  End If
'End With
End Sub
Private Sub UserForm_Activate()
On Error Resume Next
Call WindowMenuButtonDel(UFMsgBox)
If UFAinput.Visible = True Then Me.tOP = 0
'With Application ' フォームの後ろにExcelウィンドウを隠す
'  .Top = Me.Top + 1
'  .Left = Me.Left + 1
'  .Width = Me.Width - 2
'  .Height = Me.Height - 2
'End With
End Sub
Private Sub UserForm_Initialize()
Call WindowMenuButtonDel(UFMsgBox)
' 現在のExcelウィンドウ状態を保管
'With Application
'  .WindowState = xlNormal
'  g_sngTopSV = .Top
'  g_sngLeftSV = .Left
'  g_sngWidthSV = .Width
'  g_sngHeightSV = .Height
'End With
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '×では閉じないようにする
 On Error Resume Next
' 保管情報を元にExcelウィンドウを元に戻す
'With Application
'  .Top = g_sngTopSV
'  .Left = g_sngLeftSV
'  .Width = g_sngWidthSV
'  .Height = g_sngHeightSV
'End With
If (UFMsgBox.Tag = "END") Or (UFMsgBox.FlLabe = "END") Then
  UFMsgBox.Tag = ""
  UFMsgBox.FlLabe = ""
  Application.WindowState = xlMaximized '最大化
  Application.Visible = True
  Unload Me
  End '終了
Else
  Unload Me
End If

If (UFMsgBox.Tag = "Ainput") Or (UFMsgBox.FlLabe = "Ainput") Then
Else
  UF1.Show vbModeless
  DoEvents
End If
End Sub

Private Sub CommandButton1_Click()
'If (UFMsgBox.Tag = "END") Or (UFMsgBox.FlLabe = "END") Then
'  UFMsgBox.Tag = ""
'  UFMsgBox.FlLabe = ""
'  Application.WindowState = xlMaximized '最大化
'  Application.Visible = True
  Unload Me '2021.05.20_QueryCloseが動くのでこちらの記述不要
'  End '終了
'Else
'  Unload Me
'End If
'
'If (UFMsgBox.Tag = "Ainput") Or (UFMsgBox.FlLabe = "Ainput") Then
'Else
'  UF1.Show vbModeless
'  DoEvents
'End If
End Sub

