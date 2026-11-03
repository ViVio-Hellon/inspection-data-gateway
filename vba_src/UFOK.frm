Attribute VB_Name = "UFOK"
Attribute VB_Base = "0{8E30CD6E-125E-4E20-AAF9-832D0223D584}{39898B39-DAFD-4A04-A4EC-8C18FF938805}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" Alias "sleep" (ByVal ms As Long)
#End If

Private Sub UserForm_Activate()
Dim MyWaitTime As Date

  If UFcount.Visible = True Then
   Unload UFcount
  End If
  
  MyWaitTime = TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + 3)
  Application.Wait MyWaitTime

'   Sleep 1000 'waitを先に走らせないとピクチャーが表示されない

'  Cells(UFOK.AcCellrow, UFOK.AcCellcol).Select
'  AppActivate Application.Caption
  '2019.08.19 追加
  Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   Sleep 3

'  Cells(UFOK.AcCellrow, UFOK.AcCellcol).Select
'  AppActivate Application.Caption
'If UFOK.Visible = False Then
'  Cells(UFOK.AcCellrow, UFOK.AcCellcol).Select
'  AppActivate Application.Caption
'ElseIf UFOK.Visible = True Then
'   Sleep 1500
'  Cells(UFOK.AcCellrow, UFOK.AcCellcol).Select
'  AppActivate Application.Caption
'Else
'End If

End Sub

