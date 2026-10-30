Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Dim WithEvents xlApp As Application '不要エクセル対策_2021.01.18
Attribute xlApp.VB_VarHelpID = -1
Private Sub Workbook_Open()
Call Setting_Path
Set xlApp = Application '自ブック
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

Private Sub xlApp_WorkbookOpen(ByVal wb As Workbook)
If Not wb Is ThisWorkbook Then
'  MsgBox Wb.Name & "開きましたね"
    Dim path As String: path = HistoryPath
    Dim Temp As String: Temp = "*" & wb.Name & "_Open"
    Select Case True
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "L-1":        Call WriteLog("*L-1_" & Temp, "", path & "L-1\L-1_Log")
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "LVC":        Call WriteLog("*LVC_" & Temp, "", path & "LVC\LVC_Log")
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "HVC":        Call WriteLog("*HVC_" & Temp, "", path & "HVC\HVC_Log")
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "バランサー": Call WriteLog("*バランサー_" & Temp, "", path & "バランサー\バランサー_Log")
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "トット":     Call WriteLog("*トット_" & Temp, "", path & "トット\トット_Log")
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "Test":       Call WriteLog("*Test_" & Temp, "", path & "Test\Test_Log") 'path &"Test\Test_Log"
    Case Else:                                                   Call WriteLog("*丸徳_" & Temp, "", path & "丸徳\丸徳_Log")
    End Select
End If
End Sub
Private Sub xlApp_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
If Not wb Is ThisWorkbook Then
'  MsgBox Wb.Name & "閉じたね"
    Dim path As String: path = HistoryPath
    Dim Temp As String: Temp = "*" & wb.Name & "_Close"
    Select Case True
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "L-1":        Call WriteLog("*L-1_" & Temp, "", path & "L-1\L-1_Log")
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "LVC":        Call WriteLog("*LVC_" & Temp, "", path & "LVC\LVC_Log")
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "HVC":        Call WriteLog("*HVC_" & Temp, "", path & "HVC\HVC_Log")
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "バランサー": Call WriteLog("*バランサー_" & Temp, "", path & "バランサー\バランサー_Log")
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "トット":     Call WriteLog("*トット_" & Temp, "", path & "トット\トット_Log")
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "Test":       Call WriteLog("*Test_" & Temp, "", path & "Test\Test_Log") 'path &"Test\Test_Log"
    Case Else:                                                   Call WriteLog("*丸徳_" & Temp, "", path & "丸徳\丸徳_Log")
    End Select
End If
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
Call Setting_Path
'Select Case True
'Case UF1.L1Work:   ThisWorkbook.Sheets("Main").Cells(2, 1) = "L-1"
'Case UF1.LVC:      ThisWorkbook.Sheets("Main").Cells(2, 1) = "LVC"
'Case UF1.HVC:      ThisWorkbook.Sheets("Main").Cells(2, 1) = "HVC"
'Case UF1.Bala:     ThisWorkbook.Sheets("Main").Cells(2, 1) = "バランサー"
'Case UF1.Tott:     ThisWorkbook.Sheets("Main").Cells(2, 1) = "トット"
'Case UF1.Test:     ThisWorkbook.Sheets("Main").Cells(2, 1) = "Test"
'Case Else:         ThisWorkbook.Sheets("Main").Cells(2, 1) = "丸徳"
'End Select
Dim path As String: path = HistoryPath
If ThisWorkbook.Sheets("Main").Cells(1, 1) = "CLOSE" Then
    Select Case True
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "L-1":        Call WriteLog("L-1_" & "ﾌﾞｯｸ閉", "", path & "L-1\L-1_Log")
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "LVC":        Call WriteLog("LVC_" & "ﾌﾞｯｸ閉", "", path & "LVC\LVC_Log")
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "HVC":        Call WriteLog("HVC_" & "ﾌﾞｯｸ閉", "", path & "HVC\HVC_Log")
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "バランサー": Call WriteLog("バランサー_" & "ﾌﾞｯｸ閉", "", path & "バランサー\バランサー_Log")
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "トット":     Call WriteLog("トット_" & "ﾌﾞｯｸ閉", "", path & "トット\トット_Log")
    Case ThisWorkbook.Sheets("Main").Cells(2, 1) = "Test":       Call WriteLog("Test_" & "ﾌﾞｯｸ閉", "", path & "Test\Test_Log") 'path &"Test\Test_Log"
    Case Else:                                                   Call WriteLog("丸徳_" & "ﾌﾞｯｸ閉", "", path & "丸徳\丸徳_Log")
    End Select
    DoEvents
Else
  Cancel = 1
End If

End Sub

