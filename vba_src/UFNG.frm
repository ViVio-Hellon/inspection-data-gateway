Attribute VB_Name = "UFNG"
Attribute VB_Base = "0{50E3F41E-3DBB-4571-93F2-F2737D6C9B93}{F59631C7-D6CF-416A-BAC6-96410B83F16B}"
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

  If UFcount.Visible = True Then
   Unload UFcount
  End If
  MyWaitTime = TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + 3)
  Application.Wait MyWaitTime

'   Sleep 1000 'waitを先に走らせないとピクチャーが表示されない

'  Cells(UFNG.AcCellrow, UFOK.AcCellcol).Select
'  AppActivate Application.Caption
  '2019.08.19 追加
  Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'  Cells(UFNG.AcCellrow, UFNG.AcCellcol).Select
'  AppActivate Application.Caption
End Sub

