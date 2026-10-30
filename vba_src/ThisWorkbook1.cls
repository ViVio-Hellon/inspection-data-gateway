Attribute VB_Name = "ThisWorkbook1"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Sub Workbook_Open()
Application.ScreenUpdating = False
If WsCheck("キャスク検査") = False Then Exit Sub
If WsCheck("全ヵ所検査") = False Then Exit Sub
Dim i As Long
With ThisWorkbook
  For i = 1 To .Sheets.Count
    If (.Sheets("キャスク検査").Visible = True) Or (.Sheets("全ヵ所検査").Visible = True) Then
      Application.Visible = True
    End If
  Next i
End With
Application.ScreenUpdating = True
Application.WindowState = xlMaximized '最大化
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
If ThisWorkbook.Sheets("Main").Cells(1, 1) = "CLOSE" Then
Else
  Cancel = 1
End If
End Sub

