Attribute VB_Name = "UFdiagonal"
Attribute VB_Base = "0{5AEC67C2-B871-45A5-A93C-02207FD54AA7}{6A3C189D-90DA-4477-BFA2-3614DDC24A1D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit
Private Sub CommandButton1_Click()
On Error Resume Next
Dim kou As Single: kou = ReDiagonal(UF1.Hab, UF1.Tak) '緩い公差
Dim Scol As Long: Scol = 13434828 '合格
Dim Ncol As Long: Ncol = 8421631
Dim TaiA As String: TaiA = UFdiagonal.PosiA
Dim TaiB As String: TaiB = UFdiagonal.PosiB
Dim Tai As Single: Tai = Abs(val(Me.NumA) - val(Me.NumB))
Application.ScreenUpdating = False
 If Me.Tai = "3" Or Me.Tai = "３" Then
    If DiagonalJudg(Tai, kou) Then
      UF1(TaiA).BackColor = Scol
      UF1(TaiB).BackColor = Scol
      Call 音楽合格
      UF1("Diag" & Left(UFdiagonal.PosiA, 1) & "Kou") = kou
      UF1.Show vbModeless
      Call Posting
      Unload Me
    ElseIf DiagonalJudg(Tai, kou) = False Then
      UF1(TaiA).BackColor = Ncol
      UF1(TaiB).BackColor = Ncol
      Call 音楽不合格
      '公差外処理
      UF1("Diag" & Left(UFdiagonal.PosiA, 1) & "Kou") = kou
      UF1.Show vbModeless
      Call Posting
      Unload Me
    End If
 ElseIf Me.Tai = "4" Or Me.Tai = "４" Then
    UF1(TaiA).BackColor = Ncol
    UF1(TaiB).BackColor = Ncol
    Call 音楽不合格
    '公差外処理
    UF1.Show vbModeless
    Unload Me
 Else
  MsgBox "3 か 4 "
 End If
If Err.Number <> 0 Then
 Err.clear
End If
Application.ScreenUpdating = True
End Sub
Private Sub UserForm_Activate()
Dim TaiA As String: TaiA = UFdiagonal.PosiA
Dim TaiB As String: TaiB = UFdiagonal.PosiB
UFdiagonal.Terg.SetFocus
UFdiagonal.Tai.SetFocus 'インプットボックスに初めから数字が入るようにする
UFdiagonal.TAISA = "絶対値" & " " & "(" & "[対角A]" & " " & Format(UF1(TaiA), "0.0") & " " & "-" _
& "[対角B]" & " " & Format(UF1(TaiB), "0.0") & ")" & "=" & Format(val(UF1("Diag" & Left((TaiA), 1))), "0.0")
End Sub

Private Sub UserForm_Initialize()
Call WindowMenuButtonDel(Me) 'ｳｨﾝﾄﾞｳ×ﾎﾞﾀﾝを消す。
End Sub
