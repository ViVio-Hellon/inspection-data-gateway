Attribute VB_Name = "UF対角"
Attribute VB_Base = "0{CB4DAB17-9732-43BF-969B-6DA6CDEC4E87}{0B565251-1934-44D5-B037-029353376DB1}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



Option Explicit
Private Sub UserForm_Activate()
Dim MyWaitTime As Date
Dim i As Long, x As Long: x = Int(5 * Rnd(1)) + 1 '(最大値-最小値+1)*Rnd+最小値
Dim TR As Long:          TR = UF対角.TR
Dim TC As Long:          TC = UF対角.TC
Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
'リセット
For i = 1 To 5
 UF対角.Controls("DB" & i).Visible = False
Next i
Select Case ws.Name
Case "検査", "少枚数検査", "極検査"
  Select Case TR
  Case 14
   UF対角.DB1.Visible = True
  Case 15
   UF対角.DB2.Visible = True
  Case 16
   UF対角.DB3.Visible = True
  Case 17
   UF対角.DB4.Visible = True
  Case 18
   UF対角.DB5.Visible = True
  End Select
Case Else
 Randomize
 UF対角.Controls("DB" & x).Visible = True
End Select
'2019.11.15 追加
If UFOK.Visible = True Then Unload UFOK
If UFNG.Visible = True Then Unload UFNG
'2019.08.19 追加
If UFcount.Visible = True Then Unload UFcount

MyWaitTime = TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + 2)
Application.Wait MyWaitTime

Unload Me
  
End Sub

