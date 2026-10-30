Attribute VB_Name = "Main1"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CommandButton1_Click()
'Application.Visible = True 'Excel表示
'2022.11.14_初期設定追加
Call 初期設定値確認
Dim tmpFaci As String: tmpFaci = GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")
If tmpFaci <> "存在なし" Then
  
If WsCheck("全ヵ所検査") Then If ThisWorkbook.Sheets("全ヵ所検査").Visible Then ThisWorkbook.Sheets("全ヵ所検査").Visible = False
If WsCheck("キャスク検査") Then If ThisWorkbook.Sheets("キャスク検査").Visible Then ThisWorkbook.Sheets("キャスク検査").Visible = False

UF1.Show vbModeless
End If

End Sub

Private Sub CommandButton2_Click()
Form_1.Show vbModeless
End Sub

Private Sub CommandButton3_Click()
UserForm1.Show vbModeless
End Sub
