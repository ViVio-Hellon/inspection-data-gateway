Attribute VB_Name = "clsDynLabel"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
'2026.05.29_Ainput用動的ラベルで使用

Private WithEvents m_lbl As MSForms.Label
Attribute m_lbl.VB_VarHelpID = -1
Private m_index As Long

Public Sub Init(lbl As MSForms.Label, idx As Long)
    Set m_lbl = lbl
    m_index = idx
End Sub

' StartTBからの値受け取りはSetKHeight経由で行う
' ここではクリック（クリア）とダブルクリック（再入力許可）を担当

Private Sub m_lbl_Click()
    If m_lbl Is Nothing Then Exit Sub
    ' クリア
    m_lbl.Caption = ""
    m_lbl.Tag = ""
End Sub

Private Sub m_lbl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If m_lbl Is Nothing Then Exit Sub
    ' ダブルクリックで再入力：TagをクリアしてStartTBへフォーカス
    m_lbl.Tag = ""
    UFAinput.StartTB.SetFocus
End Sub

