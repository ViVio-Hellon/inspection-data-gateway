Attribute VB_Name = "UFProgress"
Attribute VB_Base = "0{5962C8CF-20E8-49AD-ADB3-05BDC67F8F78}{23B551ED-5B33-4B2B-A18B-788496AC3072}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit
'点滅用
Private BlinkLabel1 As clsBlinkLabel
'キャンセル処理用フラグ
Public IsCancel As Boolean
Private Sub UserForm_Activate()
If Not UFProgress.Tag = "1" Then
  BlinkLabel1.Blink 2 '点滅回数登録
End If
End Sub
Private Sub UserForm_Click()
BlinkLabel1.BlinkStop '点滅ｽﾄｯﾌﾟ
End Sub
Private Sub UserForm_Initialize()
'Application.Speech.Speak Text:="少々お待ちください"
Call WindowMenuButtonDel(UFProgress)
Set BlinkLabel1 = New clsBlinkLabel '点滅ラベルの登録
BlinkLabel1.cmd = Me.CommandButton1
End Sub
Private Sub UserForm_Terminate()
Set BlinkLabel1 = Nothing
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Application.Speech.Speak Text:="完了しました"
End Sub

