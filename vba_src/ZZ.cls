Attribute VB_Name = "ZZ"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "CommandButton3, 6, 0, MSForms, CommandButton"
Attribute VB_Control = "CommandButton2, 3, 1, MSForms, CommandButton"
Attribute VB_Control = "T1, 2, 2, MSForms, TextBox"
Attribute VB_Control = "CommandButton1, 1, 3, MSForms, CommandButton"
Attribute VB_Control = "T2, 4, 4, MSForms, TextBox"
Option Explicit

Private Sub CommandButton1_Click()
Application.ScreenUpdating = False
Range(Cells(5, 8), Cells(8, 12)).ClearContents
Dim TableName As String: TableName = "仕掛"
Dim FPath As String:  FPath = MainPath 'ThisWorkbook.Path & "\" 'Fpath = "\\nlmsrvngy03\Read\【New】仕掛\台帳\"
'ﾌｧｲﾙ名
Dim target As String:       target = "SIKALOTNOW.accdb"
Dim Target_Order As String: Target_Order = "SIKAODRNOW.accdb"
Dim Target_Hiki As String:  Target_Hiki = "SIKAHIKINOW.accdb"
'ﾌｧｲﾙﾊﾟｽ
Dim FilePath As String: FilePath = FPath & target 'Lot
Dim filepath_Order As String: filepath_Order = FPath & Target_Order 'ｵｰﾀﾞｰ
Dim FilePath_Hiki As String:  FilePath_Hiki = FPath & Target_Hiki '引当
'検索項目
Dim SQLfilter As String: SQLfilter = "[ﾛｯﾄ番号]='" & T1 & "'" 'ﾛｯﾄNo検索
Dim TempHiki As Variant: TempHiki = GetRecordsArr(FilePath_Hiki, TableName, SQLfilter) 'All,fields,record
If IsEmpty(TempHiki) Then 'emptyの場合停止
  MsgBox "SIKAHIKINOW:該当ロットなし"
  Exit Sub
End If
'Stop
'ﾌｨｰﾙﾄﾞ名ﾁｪｯｸ　*RNE更新に対応
Dim fieldsfilter As String: fieldsfilter = ""
Dim Hikifields As Variant:  Hikifields = GetFieldsArr(FilePath_Hiki, TableName, fieldsfilter)
Dim Orderfields As Variant: Orderfields = GetFieldsArr(filepath_Order, TableName, fieldsfilter)
Select Case True 'emptyの場合停止
Case IsEmpty(Hikifields)
  MsgBox "Hikifieldなし"
  Exit Sub
Case IsEmpty(Orderfields)
  MsgBox "Orderfieldなし"
  Exit Sub
End Select
'Stop
Dim i As Long, H_受注番号 As Long 'Hiki
For i = LBound(Hikifields, 2) To UBound(Hikifields, 2)
  If Hikifields(1, i) = "受注番号" Then
    H_受注番号 = i
    Exit For
  End If
Next i

Dim O_受注番号 As Long, O_VC_表 As Long, O_VC_裏 As Long, O_梱包単位_梱包形状_輸出先国名 As Long 'Order
For i = LBound(Orderfields, 2) To UBound(Orderfields, 2)
  Select Case Orderfields(1, i)
  Case "受注番号":                     O_受注番号 = i
  Case "VC_表":                        O_VC_表 = i
  Case "VC_裏":                        O_VC_裏 = i
  Case "梱包単位_梱包形状_輸出先国名": O_梱包単位_梱包形状_輸出先国名 = i
  End Select
Next i

Dim Orderfilter As String, TempOrder As Variant
For i = LBound(TempHiki, 1) To UBound(TempHiki, 1)
  Orderfilter = "[受注番号]='" & TempHiki(i, H_受注番号) & "'" 'ｵｰﾀﾞｰNo検索
  TempOrder = GetRecordsArr(filepath_Order, TableName, Orderfilter)
  If IsEmpty(TempOrder) = False Then
    Sheets("Test").Cells(5, 7 + i) = TempOrder(LBound(TempOrder, 1), O_受注番号)
    Sheets("Test").Cells(6, 7 + i) = TempOrder(LBound(TempOrder, 1), O_VC_表)
    Sheets("Test").Cells(7, 7 + i) = TempOrder(LBound(TempOrder, 1), O_VC_裏)
    Sheets("Test").Cells(8, 7 + i) = TempOrder(LBound(TempOrder, 1), O_梱包単位_梱包形状_輸出先国名)
  Else
    MsgBox "Order側受注Noなし" 'emptyの場合停止
    Exit Sub
  End If
Next i
'Stop
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton2_Click()
Application.ScreenUpdating = False
'Dim Wb_Path As String: Wb_Path = ThisWorkbook.Path & "\日報＆資材計算.xlsm"
'If IsBookOpened(Wb_Path) Then
'  Application.Run "'日報＆資材計算.xlsm'!GW_Open"
'End If

'If ReadCheck("C:\Windows\System32\ctfmon.exe", 1) Then
'  With CreateObject("Wscript.Shell")
'    Dim ret As Long: ret = .Run("C:\Windows\System32\ctfmon.exe", 5, True)
'  End With
'  If ret <> 0 Then
'  ' MsgBox "失敗"
'  End If
'End If

Application.Visible = True 'Excel表示
'Call HiddenSheets

Test.Show vbModeless

Call VisibleSheets
Application.ScreenUpdating = True
'Call HiddenSheets
'Call KenSheets

'ThisWorkbook.Sheets("保存用").Visible = xlSheetVisible
'Dim SaveSh As Worksheet: Set SaveSh = ThisWorkbook.Sheets("保存用")
'Dim Kouzyou As String: Kouzyou = "\\nlmsrvngy03\工場内共有\検査データ"
'If IsBookOpened(Kouzyou & "\" & "Test_File.csv") Then MsgBox "開いています"
'Call ExpCSV(SaveSh, "Test_File", Kouzyou)

'Call BookSaClose(ThisWorkbook.Path, "日報＆資材計算.xlsm") 'ﾊﾟｽ,閉じたいbook名

'Dim wTime As Date
'Dim Keika As Date: Keika = Now - wTime
'Dim hh As Long: hh = CLng(Month(Keika)) * 24
'Dim buf: buf = Format(Now, "h") + hh
''--------------------------------------------
'Dim DT As String: DT = DateAdd("d", -1, Now) '夜勤の日付またぎ前日扱い
'Dim StrDate As String: StrDate = Format$(Now, "yyyy.mm.dd  hh:nn")
'Dim StrDate2 As String: StrDate2 = Format$(DT, "yyyy.mm.dd") & "  " & buf & Format$(DT, ":nn")
'MsgBox StrDate2


'Dim f As String: f = SaveDateLocate(ThisWorkbook.Path, "(^^)")   'ﾊﾟｽ,ﾌｧｲﾙﾗｲﾝ名_必要ﾌｫﾙﾀﾞ生成とﾌｧｲﾙﾊﾟｽ
'Dim NewFiles As String: NewFiles = NewFileName("C:\検査履歴" & "\" & Format(Date, "yyyy.mm"))
'MsgBox NewFiles

'Dim TableName As String: TableName = "VC厚"
'Dim Fpath As String
'Fpath = "\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\検査\" 'ThisWorkbook.Path & "\"
''ﾌｧｲﾙ名
'Dim Target As String:   Target = "VC厚" & ".accdb"
''ﾌｧｲﾙﾊﾟｽ
'Dim FilePath As String: FilePath = Fpath & Target
'  If ReadCheck(FilePath, 1) = False Then '1：ファイル,2：フォルダ
'    MsgBox FilePath & "がありません" & "_VcConv", vbCritical
'    Exit Sub
'  End If
''検索項目
'Dim SQLfilter As String: SQLfilter = "[品名ｺｰﾄﾞ]='" & T1 & "'" 'VC検索
'Dim Temp As Variant: Temp = GetRecordsArr(FilePath, TableName, SQLfilter) 'All,fields,record
''ﾌｨｰﾙﾄﾞ名ﾁｪｯｸ　*更新に対応
'Dim fieldsfilter As String: fieldsfilter = ""
'Dim fields As Variant:  fields = GetFieldsArr (FilePath, TableName, fieldsfilter)
'If IsEmpty(fields) Then 'emptyの場合停止
'  MsgBox "fieldsなし" & "_VcConv"
'  Exit Sub
'End If
'Dim i As Long, C_厚係数 As Long
'For i = LBound(fields, 2) To UBound(fields, 2)
'  If fields(1, i) = "厚-係数" Then
'    C_厚係数 = i
'    Exit For
'  End If
'Next i
'If IsEmpty(Temp) Then 'emptyの場合停止
'  MsgBox "VC登録なし"
'  Exit Sub
'Else
'  Dim a As Double: a = Temp(1, C_厚係数)
'End If
'T2 = a


'Call No_Search(T1.Text)

'Call Ainput_Startup("UF1", "M6")

'UF品質基準.Show vbModeless
'UF品質基準.LineName = "Test"
'UF品質基準.FlagAtu = "8.0"
'UF品質基準.MP.Page1.Visible = True
'UF品質基準.MP.Page2.Visible = True
''厚板,薄板の判定を行う
''Page1=0　Page2=1
'If UF品質基準.FlagAtu <= 7.999 Then
'  UF品質基準.MP.Value = 1
'  UF品質基準.MP.Page1.Visible = False
'  Call 表面品質基準(UF品質基準.T1)
'Else
'  UF品質基準.MP.Value = 0
'  UF品質基準.MP.Page2.Visible = False
'  Call 品質グレード
'  Call 平坦度
'End If

'Application.ScreenUpdating = True
'Load UF校正
'UF校正.Kname = "Testﾏﾝ"
'UF校正.LineName = "Test"
'Call WindowMenuButtonDel(UF校正)
'UF校正.Show vbModeless 'vbModeless'vbModal

'Call Acc_Etc("SIKACoil.accdb", "SikaCoil")

'Call WindowMenuButtonDel(UF校正)
'UF校正.Show vbModeless

'Load UFAinput
'UFAinput.KenLaw = "M6"
'UFAinput.Show vbModeless


'Height_Checker.Show vbModeless
'Call WriteLog("TestBB_", "", "\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\検査履歴\Test\Test_Log")

'Call Daily_Report(3)
'Call NameList_Open(2)
'Sheets("Hiki_Lot").Visible = xlSheetVisible
'Call Name_Match

'Dim AppPath As String: AppPath = "C:\Tools\測定データ電子化\"
'Dim Ltext As String: Ltext = "result.txt"
'Dim LineBuff
'Open AppPath & Ltext For Input As #1 '文字化けする
'  Do Until EOF(1)
'    Line Input #1, LineBuff
'  Loop
'Close #1
'MsgBox LineBuff

'Call HostDataConfi
End Sub

Private Sub CommandButton3_Click()
Height_Checker.Show vbModeless
End Sub

Private Sub T1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

