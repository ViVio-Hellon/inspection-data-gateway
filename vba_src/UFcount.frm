Attribute VB_Name = "UFcount"
Attribute VB_Base = "0{4C57F798-F0F3-46B6-A1DC-E2F2D32DFFB8}{8E10735F-5B46-4BFE-8239-6C8F1F4351AA}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Sub Itano_Change()
Me.ItanoClone = Me.Itano
End Sub

Private Sub Location_Change()
Me.LocationClone = Me.Location
End Sub

Private Sub UserForm_Activate()
Dim TR As Long:          TR = UFcount.TR
Dim TC As Long:          TC = UFcount.TC
Dim ws As Worksheet: Set ws = ActiveSheet

Select Case TC
Case 2, 4, 6, 8, 11, 13, 15, 17, 19, 20 '2021.07.09_検査箇所追加対応
 If ws.Cells(TR, TC).Value <> "" Then
  UFcount.Itano = ws.Cells(TR, 1).Value
  UFcount.Location = ws.Cells(6, TC)
  If UFCask.Read = "ON" Then
    ws.Cells(TR, 1).Speak '枚数読み上げ
    ws.Cells(6, TC).Speak '位置読み上げ
  End If
 End If
End Select

'2019.08.19 追加
If UFOK.Visible = False And UFNG.Visible = False And UF対角.Visible = False Then
  Dim MyWaitTime As Date: MyWaitTime = TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + 3)
  Application.Wait MyWaitTime
  Unload UFcount
End If
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '×では閉じないようにする
  Application.EnableEvents = True
End Sub
