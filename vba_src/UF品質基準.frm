Attribute VB_Name = "UF品質基準"
Attribute VB_Base = "0{B0D8107F-B331-4133-B60F-2C3145D39645}{633EECE9-A043-410D-96ED-6EAF4875F984}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub B1_Click()
Call BoColor(B1)
Call DBdisplay外観_S(Me.T1, Me.B1.Caption, UF1.DBflag.Text)
  '旧バージョン
'Call SelectAdd(B1)
End Sub
Private Sub B2_Click()
Call BoColor(B2)
Call DBdisplay外観_S(Me.T1, Me.B2.Caption, UF1.DBflag.Text)
  '旧バージョン
'Call SelectAdd(B2)
End Sub
Private Sub B3_Click()
Call BoColor(B3)
Call DBdisplay外観_S(Me.T1, Me.B3.Caption, UF1.DBflag.Text)
  '旧バージョン
'Call SelectAdd(B3)
End Sub
Private Sub B4_Click()
Call BoColor(B4)
Call DBdisplay外観_S(Me.T1, Me.B4.Caption, UF1.DBflag.Text)
  '旧バージョン
'Call SelectAdd(B4)
End Sub
Private Sub B5_Click()
Call BoColor(B5)
Call DBdisplay外観_S(Me.T1, Me.B5.Caption, UF1.DBflag.Text)
  '旧バージョン
'Call SelectAdd(B5)
End Sub
Private Sub B6_Click()
Call BoColor(B6)
Call DBdisplay外観_S(Me.T1, Me.B6.Caption, UF1.DBflag.Text)
  '旧バージョン
'Call SelectAdd(B6)
End Sub
Private Sub B7_Click()
Call BoColor(B7)
Call DBdisplay外観_S(Me.T1, Me.B7.Caption, UF1.DBflag.Text)
  '旧バージョン
'Call SelectAdd(B7)
End Sub
Private Sub B8_Click()
Call BoColor(B8)
Call DBdisplay外観_S(Me.T1, Me.B8.Caption, UF1.DBflag.Text)
  '旧バージョン
'Call SelectAdd(B8)
End Sub
Private Sub B9_Click()
Call BoColor(B9)
Call DBdisplay外観_S(Me.T1, Me.B9.Caption, UF1.DBflag.Text)
  '旧バージョン
'Call SelectAdd(B9)
End Sub
Private Sub B10_Click()
Call BoColor(B10)
Call DBdisplay外観_S(Me.T1, Me.B10.Caption, UF1.DBflag.Text)
  '旧バージョン
'Call SelectAdd(B10)
End Sub
Private Sub B11_Click()
Call BoColor(B11)
Call DBdisplay外観_S(Me.T1, Me.B11.Caption, UF1.DBflag.Text)
  '旧バージョン
'Call SelectAdd(B11)
End Sub


Private Sub CommandButton1_Click()
With UF品質基準
 .ToG = ""
 .KiG = ""
 .HaG = ""
 .ALG = ""
 .BoG = ""
 .KaG = ""
 
 .Tobi = ""
 .Kizu = ""
 .Haku = ""
 .ALhu = ""
 .Bota = ""
 .Kami = ""
 
 .MG1 = ""
 .HaHiz1 = ""
 .TaHiz1 = ""
 .Naka1 = ""
 .Mimi1 = ""
 
 .ATU = ""
 .HABA = ""
 .TAKE = ""
 .Zai = ""
 .TYO = ""
 .YoutC = ""
 
 .Uatu = ""
 .Uhaba = ""
 .Utake = ""
 .Uzai = ""
 .Utyou = ""
 .UYoucode = ""
 .Uhyou = ""
 
 .Sansyou1 = ""
End With

UF品質基準.L3.Visible = True
UF品質基準.Repaint

Call DBdisplay外観_H(Me.T1, Me.PathText)
 
Call DBdisplay歪(Me.T1, Me.PathText)

UF品質基準.L3.Visible = False
UF品質基準.Repaint

End Sub


Private Sub CommandButton3_Click()
UF品質基準.MP.Value = 0
End Sub

Private Sub CommandButton4_Click()
UF品質基準.MP.Value = 1
End Sub

Private Sub FM1_Click()
UF1.Show vbModeless
End Sub
Private Sub Fm2_Click()
Dim i As Long
For i = 1 To 11
  UF品質基準("B" & i).BackColor = &HFFFFFF
Next i

UF品質基準.SelectO = ""
UF品質基準.SelectS = ""
UF品質基準.SelectK = ""
UF品質基準.SelectA = ""
End Sub

Private Sub CommandButton6_Click()
Unload Me
End Sub

Private Sub HaHiz1_Change()
Me.HaHiz2 = Me.HaHiz1
End Sub
Private Sub MG1_Change()
Me.MG2 = Me.MG1
End Sub
Private Sub Mimi1_Change()
Me.Mimi2 = Me.Mimi1
End Sub
Private Sub MPH_Change()
Call LoadImagesIntoMultiPageFromTemp(Me.MPH)

'ページ強調
Dim MP As MultiPage, MARK As String: MARK = "♪♪"
Set MP = Me.MPH ' UF品質基準.MPHはマルチページコントロールへの参照

If MP.SelectedItem Is Nothing Then Exit Sub

Dim SelectedPageIndex As Long: SelectedPageIndex = MP.SelectedItem.Index ' 選択されたページのインデックスを取得
Dim Tgt_PageCaption As String: Tgt_PageCaption = MP.Pages(SelectedPageIndex).Caption

Dim p As Object
  For Each p In MP.Pages
  If p.Caption = Tgt_PageCaption Then
    p.Caption = "　★" & p.Caption & "★　" '"　★" & P.Caption & "★　"'MARK & P.Caption
  Else
    p.Caption = Replace(p.Caption, MARK, vbNullString)
    p.Caption = Replace(p.Caption, "　★", vbNullString)
    p.Caption = Replace(p.Caption, "★　", vbNullString)
  End If
Next

End Sub
Private Sub Naka1_Change()
Me.Naka2 = Me.Naka1
End Sub
Private Sub Sansyou1_Change()
Me.Sansyou2 = Me.Sansyou1
End Sub
Private Sub TaHiz1_Change()
Me.TaHiz2 = Me.TaHiz1
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If WsCheck("全ヵ所検査") Then
  If ThisWorkbook.Sheets("全ヵ所検査").Visible Then
    UFMitsubishi.Show vbModeless
  Else
    UF1.Show vbModeless
  End If
End If
End Sub
Private Sub UserForm_Activate()
Call FormSetting
End Sub
Private Sub UserForm_Initialize()
Call GetZoom(Me)
Dim Temp As Variant:  Temp = InitialSetting(Me)
With UF品質基準
  .ToG = ""
  .KiG = ""
  .HaG = ""
  .ALG = ""
  .BoG = ""
  .KaG = ""
  .Tobi = ""
  .Kizu = ""
  .Haku = ""
  .ALhu = ""
  .Bota = ""
  .Kami = ""
  .MG1 = ""
  .HaHiz1 = ""
  .TaHiz1 = ""
  
  .Uatu = ""
  .Uhaba = ""
  .Utake = ""
  .Uzai = ""
  .Utyou = ""
  .UYoucode = ""
  .Uhyou = ""
  .Upater = ""
End With

Dim i As Long
For i = 1 To 11
  UF品質基準("B" & i).BackColor = &HFFFFFF
Next i

End Sub
Private Sub TabAb_Change()
Dim SelectedTabIndex As Long:  SelectedTabIndex = Me.TabAb.SelectedItem.Index
Dim SelectedCaption As String: SelectedCaption = Me.TabAb.Tabs(SelectedTabIndex).Caption
Dim SetArr As Variant:         SetArr = GradeArray
Call UF_DbDisplay(SetArr, SelectedCaption)
End Sub

Private Sub Utyou_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Private Sub UYoucode_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Private Sub Uzai_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Private Sub Uhyou_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Private Sub YoutC_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Private Sub Zai_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Private Sub Tyo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

