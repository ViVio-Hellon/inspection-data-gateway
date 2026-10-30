Attribute VB_Name = "【UFKen_09】Input"
Option Explicit
'========================================
' 山高さ動的管理 関数群
' 既存K_Height_1～6の仕組みを15まで拡張
'========================================
Public Const MAX_MOUNT As Long = 15  ' 最大山数

#If VBA7 And Win64 Then
  '64ビット版
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
  '32ビット版
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
#Else
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If

Private Type RECT
    Left   As Long
    tOP    As Long
    Right  As Long
    Bottom As Long
End Type

Private Const SWP_NOMOVE   As Long = &H2
Private Const SWP_NOZORDER As Long = &H4
Private Const LOGPIXELSX   As Long = 88

Public Function GetDPI() As Long
    #If VBA7 And Win64 Then
        Dim hdc As LongPtr
    #Else
        Dim hdc As Long
    #End If
    hdc = GetDC(0)
    GetDPI = GetDeviceCaps(hdc, LOGPIXELSX)
    ReleaseDC 0, hdc
End Function

Public Function GetActualWindowWidth() As Long
    Dim rc As RECT
    GetWindowRect g_FormHwnd, rc
    GetActualWindowWidth = rc.Right - rc.Left
End Function

Public Sub ResizeWindow(newWidthPt As Single)
    If g_FormHwnd = 0 Then
        Call DebugLog("ResizeWindow: hwnd=0 スキップ")
        Exit Sub
    End If
    Dim dpi As Long:    dpi = GetDPI()
    Dim rc As RECT
    GetWindowRect g_FormHwnd, rc
    Dim currentHeight As Long: currentHeight = rc.Bottom - rc.tOP
    Dim newWidthPx As Long:    newWidthPx = CLng(newWidthPt * dpi / 72)
    Dim screenW As Long: screenW = GetSystemMetrics32(0)
    Dim screenH As Long: screenH = GetSystemMetrics32(1)
    Dim newLeft As Long: newLeft = (screenW - newWidthPx) / 2
    Dim newTop  As Long: newTop = (screenH - currentHeight) / 2
    SetWindowPos g_FormHwnd, 0, newLeft, newTop, newWidthPx, currentHeight, SWP_NOZORDER
    Call DebugLog("ResizeWindow: dpi=" & dpi & " newWidthPx=" & newWidthPx & " newLeft=" & newLeft & " newTop=" & newTop)
End Sub

Public Sub No_Search(LotNo As String) 'Aｲﾝﾌﾟｯﾄﾌｫｰﾑの引当ｺﾝﾎﾞﾎﾞｯｸｽ用
'引当番号,受注番号,ﾛｯﾄ番号,検査番号,引当調整NO,引当数量,引当ﾌｧｲﾙ登録年月日時分秒,出荷日
'---------------------Lot情報履歴検索
If WsCheck("Lot情報履歴") = False Then Exit Sub  '-------------------------ｼｰﾄなし処理停止
If LotNo = "" Then Exit Sub
Dim HisSH As Worksheet:   Set HisSH = ThisWorkbook.Sheets("Lot情報履歴")
Dim HisRan As Range:      Set HisRan = Range(HisSH.Cells(1, 1), HisSH.Cells(UR(HisSH, 1), UC(HisSH, 1)))
Dim Hist(): Hist = HisRan
Dim i As Long, a As Long, n As Long: n = 0
Dim R_Lot As Long, R_Ord As Long, R_Hik As Long, R_Ken As Long, R_Tyo As Long, R_Sun As Long, R_Tim As Long
Dim R_Packing As Long '2022.01.27追加
For i = 1 To UBound(Hist, 2) 'ﾌｨｰﾙﾄﾞ名
  Select Case Hist(1, i)
  Case "ﾛｯﾄNo":        R_Lot = i
  Case "引当NO":       R_Hik = i
  Case "ｵｰﾀﾞｰNo":      R_Ord = i
  Case "検査No":       R_Ken = i
  Case "引当調整NO":   R_Tyo = i
  Case "引当数量":     R_Sun = i
  Case "引当ﾌｧｲﾙ登録": R_Tim = i
  Case "包装仕様NO":   R_Packing = i
  End Select
Next i
Dim FL As Boolean: FL = False
Dim TempHiki()
For i = LBound(Hist, 1) To UBound(Hist, 1) '1次元ﾁｪｯｸ
  If Hist(i, R_Lot) = LotNo Then a = a + 1
Next i

If a = 0 Then GoTo DBstart

ReDim TempHiki(1 To a, 1 To 8)
For i = LBound(Hist, 1) To UBound(Hist, 1)
  If Hist(i, R_Lot) = LotNo Then '一致の時
    FL = True
    n = n + 1
    TempHiki(n, 1) = Hist(i, R_Hik)
    TempHiki(n, 2) = Hist(i, R_Ord)
    TempHiki(n, 3) = Hist(i, R_Lot)
    TempHiki(n, 4) = Hist(i, R_Ken)
    TempHiki(n, 5) = Hist(i, R_Tyo)
    TempHiki(n, 6) = Hist(i, R_Sun)
    TempHiki(n, 7) = Hist(i, R_Tim)
    TempHiki(n, 8) = Hist(i, R_Packing) '20222.01.27_追加
  End If
Next i

If FL Then
  If IsEmpty(TempHiki) = False Then
      Call AddWS("Hiki_Lot")
      Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Hiki_Lot")
      sh.Cells.NumberFormatLocal = "@" '文字列
      Dim HiRA As Range: Set HiRA = Range(sh.Cells(1, 1), sh.Cells(UBound(TempHiki, 1), UBound(TempHiki, 2))) '範囲指定
      HiRA = TempHiki '貼り付け
      With sh.Cells.Font
        .Name = "Meiryo UI"
        .Size = 8
        .TintAndShade = 0
        sh.Cells.EntireColumn.AutoFit
      End With
  End If
End If
'---------------------DBによる検索*Lot完了後参照できない_過去dataﾃｽﾄ用_2021.02.18
DBstart:
If FL = False Then
'  MsgBox "Lot情報履歴：履歴該当なし" & "DB検索start" & "_No_Search"
  UFAinput.DBflag.Caption = "true"
  Dim TableName As String: TableName = "仕掛"
  Dim FPath As String
  Select Case True
  Case UF1.Test:  FPath = NextPath
  Case Else:      FPath = NextPath
  End Select
  Dim target As String:    target = "SIKACOMP.accdb" '*SIKACOMP:引当完了.RNE*
  Dim FilePath As String:  FilePath = FPath & target 'Lot
  Dim SQLfilter As String: SQLfilter = "[ﾛｯﾄ番号]='" & LotNo & "'" 'ﾛｯﾄNo検索
    If ReadCheck(FilePath, 1) = False Then '1：ファイル,2：ﾌｫﾙﾀﾞ
        MsgBox "日稲:" & target & "がありません" & "_ReadCheck_No_Search", vbCritical
        Exit Sub
    End If
  Dim COMPHiki As Variant: COMPHiki = GetRecordsArr(FilePath, TableName, SQLfilter)
  If IsEmpty(COMPHiki) Then 'emptyの場合停止
    MsgBox "SIKACOMP：該当ロットなし" & "_No_Search:empty"
    Call AddWS("Hiki_Lot")
    Exit Sub
  End If
  
  Dim fieldsfilter As String: fieldsfilter = ""
  Dim Hikifields As Variant:  Hikifields = GetFieldsArr(FilePath, TableName, fieldsfilter)
  R_Lot = 0
  R_Hik = 0
  R_Ord = 0
  R_Ken = 0
  R_Tyo = 0
  R_Sun = 0
  For i = LBound(Hikifields, 2) To UBound(Hikifields, 2)
    Select Case Hikifields(1, i) 'ﾌｨｰﾙﾄﾞ名
    Case "ﾛｯﾄ番号":     R_Lot = i
    Case "引当番号":    R_Hik = i
    Case "受注番号":    R_Ord = i
    Case "検査番号":    R_Ken = i
    Case "引当調整NO":  R_Tyo = i
    Case "引当数量":    R_Sun = i
    End Select
  Next i
  
  For i = LBound(COMPHiki, 1) To UBound(COMPHiki, 1) '1次元ﾁｪｯｸ
    If COMPHiki(i, R_Lot) = LotNo Then a = a + 1
  Next i
  If a = 0 Then
    MsgBox "SIKACOMP：該当Lotなし" & "_No_Search:LotNo"
    Call AddWS("Hiki_Lot")
    Exit Sub
  End If
  ReDim NewTemp(1 To a, 1 To 6)
  For i = LBound(COMPHiki, 1) To UBound(COMPHiki, 1)
    If COMPHiki(i, R_Lot) = LotNo Then '一致の時
      FL = True
      n = n + 1
      NewTemp(n, 1) = COMPHiki(i, R_Hik)
      NewTemp(n, 2) = COMPHiki(i, R_Ord)
      NewTemp(n, 3) = COMPHiki(i, R_Lot)
      NewTemp(n, 4) = COMPHiki(i, R_Ken)
      NewTemp(n, 5) = COMPHiki(i, R_Tyo)
      NewTemp(n, 6) = COMPHiki(i, R_Sun)
    End If
  Next i
  If FL = False Then
    MsgBox "SIKACOMP：該当Lotなし" & "_No_Search:NewTemp"
    Call AddWS("Hiki_Lot")
    Exit Sub
  End If
  If IsEmpty(NewTemp) = False Then
    Call AddWS("Hiki_Lot")
    Set sh = ThisWorkbook.Sheets("Hiki_Lot")
    sh.Cells.NumberFormatLocal = "@" '文字列
    Dim COMPRA As Range: Set COMPRA = Range(sh.Cells(1, 1), sh.Cells(UBound(NewTemp, 1), UBound(NewTemp, 2))) '範囲指定
    COMPRA = NewTemp '貼り付け
    With sh.Cells.Font
      .Name = "Meiryo UI"
      .Size = 8
      .TintAndShade = 0
      sh.Cells.EntireColumn.AutoFit
    End With
  End If
End If
End Sub

Public Sub Ainput_Startup(Temp As String, KenM As String) '起動ﾌｫｰﾑ、検査法
If WsCheck("Lot情報履歴") = False Then
  MsgBox "Lot情報履歴ｼｰﾄがありません" & "_Ainput_Startup", vbCritical
  Exit Sub
End If
Dim HistorySH As Worksheet: Set HistorySH = ThisWorkbook.Sheets("Lot情報履歴")
Dim LR As Long: LR = UR(HistorySH, 1)
Dim LC As Long: LC = UC(HistorySH, 1)
Dim HistoryR As Range: Set HistoryR = Range(HistorySH.Cells(2, 1), HistorySH.Cells(LR, LC))
Dim TempHiki(): TempHiki = HistoryR
'Dim CaskSh As Worksheet:  If WsCheck("キャスク検査") Then Set CaskSh = ThisWorkbook.Sheets("キャスク検査")
'Dim MitSh As Worksheet:   If WsCheck("全ヵ所検査") Then Set MitSh = ThisWorkbook.Sheets("全ヵ所検査")
'Dim H494Sh As Worksheet:  If WsCheck("H494検査") Then Set H494Sh = ThisWorkbook.Sheets("全ヵ所検査")

Dim TmpSh As Worksheet
Dim i As Long, x As Long, matchFlag As Boolean, WorkingLot As String

Dim Lname As String
Select Case True
Case UF1.Visible:          Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Case UFCask.Visible:       Lname = UFCask.LineName.Caption
Case UFMitsubishi.Visible: Lname = UFMitsubishi.LineName.Caption
End Select

Select Case Temp
Case "UF1":        WorkingLot = UF1.Lot
Case "Cask":       Set TmpSh = ThisWorkbook.Sheets("キャスク検査")
Case "Mitsubishi": Set TmpSh = ThisWorkbook.Sheets("全ヵ所検査")
Case "H494":       Set TmpSh = ThisWorkbook.Sheets("H494検査")
End Select

If Temp <> "UF1" Then WorkingLot = TmpSh.Cells(2, 4)

For i = 1 To UBound(TempHiki, 1)
  If TempHiki(i, 3) = WorkingLot Then 'LOTNO一致
    x = i
    matchFlag = True
    Exit For
  End If
Next i
Load UFAinput

UFAinput.LineName = Lname 'ﾗｲﾝ名
With UFAinput
  If matchFlag Then
    .Lot = WorkingLot
    .KenNo = TempHiki(x, 7)
    .Zai = TempHiki(x, 8)
    .TYO = TempHiki(x, 9)
    .ATU = Format(TempHiki(x, 4), "0.000")
    .HABA = Format(TempHiki(x, 5), "0.0")
    .TAKE = Format(TempHiki(x, 6), "0.0")
    .KenLaw = KenM '検査法
    .Youto = TempHiki(x, 10) '*2021.05.10_ｽﾄｱのVC判定で使用
  ElseIf matchFlag = False Then '---------------------DBによる検索_過去dataﾃｽﾄ用_2021.02.18
    Dim stTimer As Date: stTimer = Timer()
    If UF1.Visible = True Then
      UFAinput.DBflag.Caption = "true"
      MsgBox "Lot情報履歴：履歴該当なし。" & vbCrLf & "*DB検索start(15秒程かかります)" & vbCrLf & "_Ainput_Startup"
    Else
      Exit Sub
    End If
    Dim TableName As String: TableName = "仕掛"
    Dim FPath As String
    Select Case True
    Case UF1.Test:  FPath = NextPath
    Case Else:      FPath = NextPath
    End Select
    Dim target As String:    target = "SIKATWOYEAR.accdb" '*SIKATWOYEAR:仕掛_完了2年.RNE *SIKACOMP:引当完了.RNE*
    Dim FilePath As String:  FilePath = FPath & target 'Lot
    Dim SQLfilter As String: SQLfilter = "[ﾛｯﾄ番号]='" & WorkingLot & "'" 'ﾛｯﾄNo検索
      If ReadCheck(FilePath, 1) = False Then '1：ファイル,2：ﾌｫﾙﾀﾞ
          MsgBox "日稲" & target & "がありません" & "_ReadCheck_Ainput_Startup", vbCritical
          Exit Sub
      End If
    Dim COMPHiki As Variant: COMPHiki = GetRecordsArr(FilePath, TableName, SQLfilter)
    If IsEmpty(COMPHiki) Then 'emptyの場合停止
      MsgBox "SIKACOMP：該当ロットなし" & "_Ainput_Startup"
      Exit Sub
    End If
    Dim R_Lot As Long, R_Ken As Long, R_zai As Long, R_Tyo As Long, R_atu As Long, R_hab As Long, R_tak As Long
    Dim fieldsfilter As String: fieldsfilter = ""
    Dim Hikifields As Variant:  Hikifields = GetFieldsArr(FilePath, TableName, fieldsfilter)
    For i = LBound(Hikifields, 2) To UBound(Hikifields, 2)
      Select Case Hikifields(1, i) 'ﾌｨｰﾙﾄﾞ名
      Case "ﾛｯﾄ番号":    R_Lot = i
      Case "検査番号":   R_Ken = i
      Case "製造材質":   R_zai = i
      Case "製造調質":   R_Tyo = i
      Case "製造板厚":   R_atu = i
      Case "製造板幅":   R_hab = i
      Case "製造板丈":   R_tak = i
      End Select
    Next i
    .Lot = WorkingLot
    .KenNo = COMPHiki(1, R_Ken)
    .Zai = COMPHiki(1, R_zai)
    .TYO = COMPHiki(1, R_Tyo)
    .ATU = Format(COMPHiki(1, R_atu), "0.000")
    .HABA = Format(COMPHiki(1, R_hab), "0.0")
    .TAKE = Format(COMPHiki(1, R_tak), "0.0")
    .KenLaw = KenM '検査法
  End If
  
  Dim Workers As String
  Select Case Temp
  Case "UF1"
    .AMin = UF1.AMin
    .HMin = UF1.HMin
    .TMin = UF1.TMin
    .AMax = UF1.AMax
    .HMax = UF1.HMax
    .TMax = UF1.TMax
    '極力UF1は干渉させない_2021.09.21
    Select Case WorkeDayMax '登録時間が一番遅い
    Case "1": Workers = " T1:" & UF1.T1.Text
    Case "2": Workers = " T2:" & UF1.T2.Text
    Case "3": Workers = " T3:" & UF1.T3.Text
    End Select
  Case "Mitsubishi" '座標変更_2024.1.11
    If TmpSh.Cells(4, 23) <> 0 Then .AMin = Format(TmpSh.Cells(4, 23), "0.000")
    If TmpSh.Cells(4, 25) <> 0 Then .HMin = Format(TmpSh.Cells(4, 25), "0.0")
    If TmpSh.Cells(4, 27) <> 0 Then .TMin = Format(TmpSh.Cells(4, 27), "0.0")
    If TmpSh.Cells(3, 23) <> 0 Then .AMax = Format(TmpSh.Cells(3, 23), "0.000")
    If TmpSh.Cells(3, 25) <> 0 Then .HMax = Format(TmpSh.Cells(3, 25), "0.0")
    If TmpSh.Cells(3, 27) <> 0 Then .TMax = Format(TmpSh.Cells(3, 27), "0.0")
  Case "Cask", "Mitsubishi", "H494"
    If TmpSh.Cells(4, 22) <> 0 Then .AMin = Format(TmpSh.Cells(4, 22), "0.000")
    If TmpSh.Cells(4, 23) <> 0 Then .HMin = Format(TmpSh.Cells(4, 23), "0.0")
    If TmpSh.Cells(4, 24) <> 0 Then .TMin = Format(TmpSh.Cells(4, 24), "0.0")
    If TmpSh.Cells(3, 22) <> 0 Then .AMax = Format(TmpSh.Cells(3, 22), "0.000")
    If TmpSh.Cells(3, 23) <> 0 Then .HMax = Format(TmpSh.Cells(3, 23), "0.0")
    If TmpSh.Cells(3, 24) <> 0 Then .TMax = Format(TmpSh.Cells(3, 24), "0.0")
  End Select
  
  Select Case Temp '極力UF1は干渉させない_2021.09.21
  Case "Cask":       Workers = "UFｷｬｽｸ:" & UFCask.Kname
  Case "Mitsubishi": Workers = "UF三菱:" & UFMitsubishi.Kname
  Case "H494":       Workers = "UF三菱:" & UFMitsubishi.H494_Kname
  End Select
End With

Dim path As String: path = HistoryPath
Select Case Lname '2021.03.30*UFからﾓｼﾞｭｰﾙへ変更
Case "L-1":    Call WriteLog("*UFKen:" & "UFAinput起動" & Workers, "", path & "L-1\L-1_Log")
Case "LVC":    Call WriteLog("*UFKen:" & "UFAinput起動" & Workers, "", path & "LVC\LVC_Log")
Case "HVC":    Call WriteLog("*UFKen:" & "UFAinput起動" & Workers, "", path & "HVC\HVC_Log")
Case "ﾊﾞﾗﾝｻｰ": Call WriteLog("*UFKen:" & "UFAinput起動" & Workers, "", path & "バランサー\バランサー_Log")
Case "ﾄｯﾄ":    Call WriteLog("*UFKen:" & "UFAinput起動" & Workers, "", path & "トット\トット_Log")
Case "Test":   Call WriteLog("*UFKen:" & "UFAinput起動" & Workers, "", path & "Test\Test_Log")
Case Else:     Call WriteLog("*UFKen:" & "UFAinput起動" & Workers, "", path & "丸徳\丸徳_Log")
End Select

Select Case Temp
Case "UF1":                UF1.Hide
Case "Cask":               UFCask.Hide
Case "Mitsubishi", "H494": UFMitsubishi.Hide
End Select

Call No_Search(WorkingLot) '引当からLotNo検索

UFAinput.ExportLabel.Visible = False
UFAinput.ExFrame.Visible = False
UFAinput.Height = 437

DoEvents

UFAinput.Show vbModeless

Sleep 1000
'MsgBox Format(Timer() - stTimer, "0.0000000秒")
'イニシャライズだとなんかあかんわ
Call GetZoom(UFAinput)
Dim AinTemp As Variant:  AinTemp = InitialSetting(UFAinput)

End Sub

Public Sub Number_Inp(TB As MSForms.Label) 'ｲﾝﾌﾟｯﾄﾎﾞｯｸｽ入力
    Dim Score As Variant
    Select Case TB.Name
    Case "KONNO":       Score = Application.InputBox("梱包NO入力", "梱包NO")
    Case "MAI":         Score = Application.InputBox("枚数入力", "梱包枚数")
    Case "COU":         Score = Application.InputBox("包み数入力", "梱包数")
    Case "Mount_Count": Score = Application.InputBox("梱包山数入力", "山数")
    End Select
    If Score = False Then Exit Sub
    ' Mount_Countのとき動的ラベル更新をここで実行_2026.05.29
    If TB.Name = "Mount_Count" Then
        Dim mc As Long: mc = val(Score)
        If mc < 1 Or mc > MAX_MOUNT Then
            MsgBox "山数は1～" & MAX_MOUNT & "で入力してください。", vbExclamation
            Exit Sub
        End If
    End If
    UFAinput(TB.Name) = LCase(Score)
    If TB.Name = "Mount_Count" Then
        Dim mc2 As Long: mc2 = val(UFAinput.Mount_Count.Caption)
        Call DebugLog("Number_Inp Mount_Count Caption=" & UFAinput.Mount_Count.Caption & " mc=" & mc)
        If mc >= 1 And mc <= MAX_MOUNT Then
            UFAinput.ShowHeightControls mc2
        End If
        UFAinput.StartTB.SetFocus
    End If
End Sub

Public Function Height_Check(ATU1 As Single, ATU2 As Single, Hei As Single, Mai As Long, OrderNo As String) As Boolean
    Call DebugLog("Height_Check 開始 ATU1=" & ATU1 & " ATU2=" & ATU2 & " Hei=" & Hei & " Mai=" & Mai & " OrderNo=" & OrderNo)
    
    '--- 1. 初期チェックとデータ抽出 ---
    If Not WsCheck("Lot情報履歴") Then Exit Function
    Height_Check = False
    
    Dim HistorySH As Worksheet: Set HistorySH = ThisWorkbook.Sheets("Lot情報履歴")
    Dim LR As Long: LR = UR(HistorySH, 1)
    Dim LC As Long: LC = UC(HistorySH, 1)
    Dim TempHiki(): TempHiki = Range(HistorySH.Cells(2, 1), HistorySH.Cells(LR, LC))
    
    Dim i As Long, x As Long, matchFlag As Boolean: matchFlag = False
    For i = LBound(TempHiki, 1) To UBound(TempHiki, 1)
        If TempHiki(i, 2) = OrderNo Then
            x = i
            matchFlag = True
            Exit For
        End If
    Next i
    
    ' 該当データなし
    If Not matchFlag Then
        Call DebugLog("Height_Check: OrderNo一致なし")
        MsgBox "Lot情報履歴内にｵｰﾀﾞｰNO無し"
        UF1.Show vbModeless
        DoEvents
        Exit Function
    End If
    Call DebugLog("Height_Check: OrderNo一致 x=" & x)

    '--- 2. 計算用変数の定義 ---
    Dim Ais As Single: Ais = 0   ' 合紙（間紙）の総厚
    Dim Vcs As Single: Vcs = 0   ' VC皮膜の総厚
    Dim Gai As Single: Gai = 0   ' 外装紙の総厚
    Dim A1 As Single:  A1 = 0    ' 合紙1枚あたりの厚み
    Dim V1 As Single:  V1 = 0    ' VC1枚あたりの厚み
    
    ' 板厚のMax/Min（薄板の判定漏れを防ぐため厚い方をMaxとする）
    Dim AtuMax As Single: AtuMax = ATU2
    Dim AtuMin As Single: AtuMin = ATU1
    Call DebugLog("AtuMax=" & AtuMax & " AtuMin=" & AtuMin)

    '--- 3. 合紙(Ais)の計算 ---
    Dim unitAis As Single: unitAis = 0
    If (TempHiki(x, 17) <> 0) Then
        unitAis = 0.029 ' 通常
        UFAinput.AiText = "通常"
    ElseIf (UFAinput.Remain = 1) Then
        unitAis = 0.042 ' ストア用
        UFAinput.AiText = "ストア"
    End If
    Call DebugLog("unitAis=" & unitAis & " TempHiki(x,17)=" & TempHiki(x, 17) & " Remain=" & UFAinput.Remain)

    If unitAis > 0 Or (UFAinput.GaityuuText = "外注" And UFAinput.Ai = True) Then
        ' 外注かつ合紙ありの場合は、ストア判定等に関わらず0.029で計算し直すロジックを継承
        If UFAinput.GaityuuText = "外注" And UFAinput.Ai = True Then unitAis = 0.029: UFAinput.AiText = "通常"
        
        A1 = unitAis
        Dim extraMai As Long: extraMai = 1 ' 基本は枚数+1
        If UFAinput.COU1 <> "" Then extraMai = extraMai + val(UFAinput.COU1)
        If UFAinput.COU2 <> "" Then extraMai = extraMai + val(UFAinput.COU2)
        
        Ais = val(Mai + extraMai - 1) * A1 ' 元のSelect Caseの計算結果と一致
        UFAinput.AiText = UFAinput.AiText & Ais
        
        ' 外注出荷時のリセット処理
        If UFAinput.GaityuuText = "外注" And UFAinput.Ai = True And UFAinput.AiText <> "通常" & Ais Then
            ' 外注ロジック特有の挙動を維持
        ElseIf UFAinput.GaityuuText = "外注" And UFAinput.Ai = True Then
            ' 何もしない（上の計算を維持）
        ElseIf UFAinput.GaityuuText = "外注" Then
            A1 = 0: Ais = 0: UFAinput.AiText = 0
        End If
    End If
    Call DebugLog("Ais=" & Ais & " A1=" & A1)
    Dim VCA As Single: VCA = 0
    Dim VCB As Single: VCB = 0

    '--- 4. VC皮膜(Vcs)の計算 ---
    If (TempHiki(x, 18) = 1) Then
        ' 高精度等の条件チェック
        If (UFAinput.Remain <> 1) Or (UFAinput.Youto = "H191" Or UFAinput.Youto = "H242" Or UFAinput.Youto = "H241") Then
        '空チェックをしてから VcConv を呼ぶ
            If Nz(TempHiki(x, 19)) <> "" Then VCA = VcConv(CStr(TempHiki(x, 19)))
            If Nz(TempHiki(x, 20)) <> "" Then VCB = VcConv(CStr(TempHiki(x, 20)))
            V1 = VCA + VCB
            Vcs = V1 * val(Mai)
            UFAinput.VcText = Vcs
            ' 外注かつVCチェックありならリセット
            If UFAinput.GaityuuText = "外注" And UFAinput.Vab = True Then
                V1 = 0: Vcs = 0: UFAinput.VcText = 0
            End If
        End If
    Else
        UFAinput.VCAB.Caption = 0
    End If
    Call DebugLog("Vcs=" & Vcs & " V1=" & V1 & " TempHiki(x,18)=" & TempHiki(x, 18))

    '--- 5. 外装紙(Gai)の計算 ---
    ' 個装設定(COU1, COU2)がある場合のみ加算
    If val(UFAinput.COU1) > 1 Or val(UFAinput.COU2) > 1 Then
        Dim cntGai As Long: cntGai = 0
        If val(UFAinput.COU1) > 1 Then cntGai = cntGai + val(UFAinput.COU1)
        If val(UFAinput.COU2) > 1 Then cntGai = cntGai + val(UFAinput.COU2)
        Gai = cntGai * (2 * 0.119 + 0.119)
    End If
    UFAinput.GaText = Gai
    If UFAinput.GaityuuText = "外注" And UFAinput.Ga = True Then Gai = 0: UFAinput.GaText = 0
    Call DebugLog("Gai=" & Gai)

    '--- 6. 高さ判定ロジック ---
    ' 【Gosa】板以外の副資材（合紙 + VC膜 + 外装紙）の合計厚み
    Dim Gosa As Single: Gosa = Ais + Vcs + Gai
    
    ' 【GP / GM】理論上の最大高さと最小高さ
    ' GP: (最大板厚 × 枚数) + 副資材
    ' GM: (最小板厚 × 枚数) + 副資材
    Dim GP As Single: GP = (AtuMax * Mai) + Gosa
    Dim GM As Single: GM = (AtuMin * Mai) + Gosa
    
    ' 【ErrorA / ErrorB】許容できる誤差のしきい値
    ' ErrorA: 板厚の75%まで（上方向）
    ' ErrorB: 板厚の33%まで（下方向）
    Dim ErrorA As Single: ErrorA = (val(AtuMax) * 3 / 4)
    Dim ErrorB As Single: ErrorB = (val(AtuMin) * 1 / 3)
    
    ' 【SAP / SAM】理論値と実測値の差（絶対値）
    Dim SAP As Single: SAP = Abs(val(Hei) - val(GP))
    Dim SAM As Single: SAM = Abs(val(Hei) - val(GM))
    Call DebugLog("Gosa=" & Gosa & " GP=" & GP & " GM=" & GM)
    Call DebugLog("ErrorA=" & ErrorA & " ErrorB=" & ErrorB & " SAP=" & SAP & " SAM=" & SAM)
    
    Dim Count_1 As Single ' 1セット（板+合紙+VC）あたりの厚み
    
    ' --- NG判定1：実測が理論最大値(GP)を超え、かつ誤差が許容値(ErrorA)以上の時 ---
    If (val(Hei) > GP) And (Format(SAP, "0.00") > ErrorA) Then
        Count_1 = AtuMax + A1 + V1
        Height_Check = True
        Call DebugLog("NG判定1: 許容値オーバー Count_1=" & Count_1)
        Call MsgError("[許容値ｵｰﾊﾞｰ]" & " *測定高さ:" & Format(Hei, "0.00") & "*推定枚数:" & Format((Hei - Gai) / Count_1, "0.00") & " *入力枚数:" & Mai & "枚", UFAinput.LineName, 2, "Ainput")
        Call ClearAllKHeight: UFAinput.Tag = "": Exit Function
    End If
    
    ' --- NG判定2：実測が理論最小値(GM)を下回り、かつ誤差が許容値(ErrorB)以上の時 ---
    If (val(Hei) < GM) And (Format(SAM, "0.00") > ErrorB) Then
        Count_1 = AtuMin + A1 + V1
        Height_Check = True
        Call DebugLog("NG判定2: 枚数不足 Count_1=" & Count_1)
        Call MsgError("[枚数不足(?)]" & " *測定高さ:" & Format(Hei, "0.00") & "*推定枚数:" & Format((Hei - Gai) / Count_1, "0.00") & " *入力枚数:" & Mai & "枚", UFAinput.LineName, 2, "Ainput")
        Call ClearAllKHeight: UFAinput.Tag = "": Exit Function
    End If

    Call DebugLog("Height_Check 正常終了 Height_Check=" & Height_Check)
End Function

' 補助関数：Null対策
Private Function Nz(val As Variant) As String
    If IsNull(val) Then Nz = "" Else Nz = CStr(val)
End Function


'Public Function Height_Check(ATU1 As Single, ATU2 As Single, Hei As Single, Mai As Long, OrderNo As String) As Boolean '高さﾁｪｯｸ
''厚1、厚2、高さ、枚数、オーダーNo
'If WsCheck("Lot情報履歴") = False Then Exit Function '--------------------ｼｰﾄなし
'Height_Check = False
'Dim HistorySH As Worksheet: Set HistorySH = ThisWorkbook.Sheets("Lot情報履歴")
'Dim LR As Long: LR = UR(HistorySH, 1)
'Dim LC As Long: LC = UC(HistorySH, 1)
'Dim HistoryR As Range: Set HistoryR = Range(HistorySH.Cells(2, 1), HistorySH.Cells(LR, LC))
'Dim TempHiki(): TempHiki = HistoryR
'Dim i As Long, x As Long, matchFlag As Boolean: matchFlag = False
'For i = LBound(TempHiki, 1) To UBound(TempHiki, 1)
'  If TempHiki(i, 2) = OrderNo Then 'ｵｰﾀﾞｰNO一致 引数盛り込み2020.09.22
'    x = i
'    matchFlag = True
'    Exit For
'  End If
'Next i
'If matchFlag = False Then GoTo LineEnd Else GoTo LineStert '分岐
'LineStert: '---------------------------一致あり
'  Dim Vcs As Single: Vcs = 0
'  Dim Ais As Single: Ais = 0
'  '厚平均
'  Dim AtuMax As Single: AtuMax = ATU2 'ATU = (ATU1 + ATU2) / 2’厚い方のみ使用でないと薄板はほぼ外れる　2020.08.12
'  Dim AtuMin As Single: AtuMin = ATU1 '2021.01.20
'  '合紙
'  Dim A1 As Single
'  If (TempHiki(x, 17) <> 0) Then '2021.05.10_ｽﾄｱの場合を追加
'    Select Case True '一番下の分で枚数+1'厚0.22~0.36
'    Case (UFAinput.COU1 <> "" And UFAinput.COU2 <> ""): Ais = val((Mai) * 0.029) + (val(UFAinput.COU1) * 0.029) + (val(UFAinput.COU2) * 0.029) '個装時は個装ごとに+1_2021.04.16
'    Case (UFAinput.COU1 <> "" And UFAinput.COU2 = ""):  Ais = val((Mai) * 0.029) + (val(UFAinput.COU1) * 0.029)
'    Case (UFAinput.COU1 = "" And UFAinput.COU2 <> ""):  Ais = val((Mai) * 0.029) + (val(UFAinput.COU2) * 0.029)
'    Case Else:                                          Ais = val((Mai) + 1) * 0.029
'    End Select
'    A1 = 0.029 '1枚分
'    UFAinput.AiText = "通常" & Ais
'
'    '外注出荷_2021.09.08
'    If UFAinput.GaityuuText = "外注" And UFAinput.Ai = True Then
'      A1 = 0
'      Ais = 0
'      UFAinput.AiText = 0
'    End If
'  ElseIf (UFAinput.Remain = 1) Then 'ストア用合紙_2021.07.21
'    Select Case True '一番下の分で枚数+1'厚0.36~0.47
'    Case (UFAinput.COU1 <> "" And UFAinput.COU2 <> ""): Ais = val((Mai) * 0.042) + (val(UFAinput.COU1) * 0.042) + (val(UFAinput.COU2) * 0.042)
'    Case (UFAinput.COU1 <> "" And UFAinput.COU2 = ""):  Ais = val((Mai) * 0.042) + (val(UFAinput.COU1) * 0.042)
'    Case (UFAinput.COU1 = "" And UFAinput.COU2 <> ""):  Ais = val((Mai) * 0.042) + (val(UFAinput.COU2) * 0.042)
'    Case Else:                                          Ais = val((Mai) + 1) * 0.042
'    End Select
'    A1 = 0.042 '1枚分
'    UFAinput.AiText = "ストア" & Ais
'
'    '外注出荷_2021.09.08
'    If UFAinput.GaityuuText = "外注" And UFAinput.Ai = True Then
'      A1 = 0
'      Ais = 0
'      UFAinput.AiText = 0
'    End If
'  Else
'    A1 = 0
'    Ais = 0
'    UFAinput.AiText = 0
'    If UFAinput.GaityuuText = "外注" And UFAinput.Ai = True Then
'      Select Case True '一番下の分で枚数+1'厚0.22~0.36
'      Case (UFAinput.COU1 <> "" And UFAinput.COU2 <> ""): Ais = val((Mai) * 0.029) + (val(UFAinput.COU1) * 0.029) + (val(UFAinput.COU2) * 0.029) '個装時は個装ごとに+1_2021.04.16
'      Case (UFAinput.COU1 <> "" And UFAinput.COU2 = ""):  Ais = val((Mai) * 0.029) + (val(UFAinput.COU1) * 0.029)
'      Case (UFAinput.COU1 = "" And UFAinput.COU2 <> ""):  Ais = val((Mai) * 0.029) + (val(UFAinput.COU2) * 0.029)
'      Case Else:                                          Ais = val((Mai) + 1) * 0.029
'      End Select
'      A1 = 0.029 '1枚分
'      UFAinput.AiText = "通常" & Ais
'    End If
'  End If
'  'VC
'  Dim V1 As Single '1枚分
'  If (TempHiki(x, 18) <> 1) Or (UFAinput.Remain = 1) Then '引当上なしかｽﾄｱの場合
'    Vcs = 0
'    V1 = 0
'    UFAinput.VCAB.Caption = 0
'  ElseIf (TempHiki(x, 18) = 1) Or (TempHiki(x, 18) = 1 And UFAinput.Remain = 1 And (UFAinput.Youto = "H191" Or UFAinput.Youto = "H242" Or UFAinput.Youto = "H241")) Then '2021.05.10:ｽﾄｱでかつ高精度の場合VC有
'    Dim VCA As String, VCB As String
'    If TempHiki(x, 19) <> "" Then VCA = TempHiki(x, 19)
'    If TempHiki(x, 20) <> "" Then VCB = TempHiki(x, 20)
'    Select Case True
'    Case VCA <> "" And VCB <> "":  Vcs = (VcConv(VCA) + VcConv(VCB)) * val(Mai) 'AB
'    Case VCA <> "" And VCB = "":   Vcs = VcConv(VCA) * val(Mai) 'A
'    Case VCA = "" And VCB <> "":   Vcs = VcConv(VCB) * val(Mai) 'B
'    Case VCA = "" And VCB = "":    Vcs = 0 * val(Mai)
'    End Select
'    '1枚分
'    Select Case True
'    Case VCA <> "" And VCB <> "":  V1 = (VcConv(VCA) + VcConv(VCB)) 'AB
'    Case VCA <> "" And VCB = "":   V1 = VcConv(VCA)  'A
'    Case VCA = "" And VCB <> "":   V1 = VcConv(VCB)  'B
'    Case VCA = "" And VCB = "":    V1 = 0
'    End Select
'    UFAinput.VcText = Vcs
'
'    '外注出荷_2021.09.08
'    If UFAinput.GaityuuText = "外注" And UFAinput.Vab = True Then
'      V1 = 0
'      Vcs = 0
'      UFAinput.VcText = 0
'    End If
'  End If
'  '外装紙
'  Dim Gai As Single
'  Select Case True '上下+折り返し*2021.05.10_<> ""で分岐させていたが積上げ時でも1入力があるため ">"へ修正
'  Case (val(UFAinput.COU1) > 1 And val(UFAinput.COU2) > 1): Gai = val(UFAinput.COU1) + val(UFAinput.COU2) * (2 * 0.119 + 0.119) '高精度ﾏｲｸﾛ0.081_通常ﾏｲｸﾛ0.119*修正2021.04.16
'  Case (val(UFAinput.COU1) > 1 And UFAinput.COU2 = ""):     Gai = val(UFAinput.COU1) * (2 * 0.119 + 0.119) '()が無かったので0.119が最後に+されている状態だった_2021.04.16
'  Case (UFAinput.COU1 = "" And val(UFAinput.COU2) > 1):     Gai = val(UFAinput.COU2) * (2 * 0.119 + 0.119)
'  Case Else:                                                Gai = 0 '2021.05.07_積上げの場合中を測る為0*_2021.04.16*梱包しているから0ではない
'  End Select
'  UFAinput.GaText = Gai
'
'  '外注出荷_2021.09.08
'  If UFAinput.GaityuuText = "外注" And UFAinput.Ga = True Then
'    Gai = 0
'    UFAinput.GaText = 0
'  End If
'
'  Dim Gosa As Single: Gosa = Ais + Vcs + Gai '枚数分合紙　+　枚位数分VC膜厚 + 個装時の外装紙_2021.03.10
'  Dim GP As Single:   GP = (AtuMax * Mai) + Gosa '板max厚*枚数　+合紙+VC+外装紙
'  Dim GM As Single:   GM = (AtuMin * Mai) + Gosa '板min厚*枚数　+合紙+VC+外装紙
'  Dim ErrorA As Single: ErrorA = (val(AtuMax) * 3 / 4) '許容誤差　厚*3/4 2020.08.17_ﾌﾟﾗｽ誤差で使用
'  Dim ErrorB As Single: ErrorB = (val(AtuMin) * 1 / 3) '許容誤差　厚*1/3 2021.01.12_ﾏｲﾅｽ誤差で使用
'
'  Dim SAP As Single: SAP = Abs((val(Hei)) - val(GP)) '実測高さから計算値を引く
'  Dim SAM As Single: SAM = Abs((val(Hei)) - val(GM))
'  'Stop
'  Dim Count_1 As Single
'  'NG定義
'  If (val(Hei) > GP) And (Format((SAP), "0.00") > ErrorA) Then '実測高さが計算値より大きく、その差が許容誤差より大きい
'    Count_1 = A1 + V1 + (AtuMax)  '2021.04.01_表示をシンプルにする*外装紙分は不要_2021.04.09
'    Height_Check = True
'    Call MsgError("[許容値ｵｰﾊﾞｰ]" & " *測定高さ:" & Format(Hei, "0.00") & "*測定高さから計算した枚数:" & Format((Hei - Gai) / Count_1, "0.00") & " *入力枚数" & Mai & "枚", UFAinput.LineName, 2, "Ainput") '& " *計算高さ:" & Format(GP, "0.0") & " *ﾌﾟﾗｽ許容誤差:" & Format(ErrorA, "0.000"), UFAinput.LineName, 2, "Ainput")
'    DoEvents
'
'    Call ClearAllKHeight '2026.05.29
'
'    UFAinput.Tag = ""
''    Unload UFAinput
'    Exit Function
'  End If
'  If (val(Hei) < GM) And (Format((SAM), "0.00") > ErrorB) Then   '高さ測定値がﾏｲﾅｽ計算値より小さい場合(枚数不足)_2021.01.12
'    Count_1 = A1 + V1 + (AtuMin)
'    Height_Check = True
'    Call MsgError("[枚数不足(?)]" & " *測定高さ:" & Format(Hei, "0.00") & "*測定高さから計算した枚数:" & Format((Hei - Gai) / Count_1, "0.00") & " *入力枚数" & Mai & "枚", UFAinput.LineName, 2, "Ainput") '& " *計算高さ:" & Format(GM, "0.0") & " *ﾏｲﾅｽ許容誤差:" & Format(ErrorB, "0.000"), UFAinput.LineName, 2, "Ainput")
'
'    Call ClearAllKHeight '2026.05.29
'
'    UFAinput.Tag = ""
'    Exit Function
'  End If
'  Exit Function
'LineEnd: '---------------------------一致なし
'MsgBox "Lot情報履歴内にｵｰﾀﾞｰNO無し"
'UF1.Show vbModeless
'DoEvents
'End Function

Public Function VcConv(no As String) As Single 'VC検索 膜厚
'------------------------------DBから読み取る
Dim TableName As String: TableName = "VC厚"
Dim FPath As String

Select Case True
Case UF1.Test:  FPath = PackingPath 'ThisWorkbook.Path & "\"
Case Else:      FPath = QualityPath '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\検査\"
End Select

'ﾌｧｲﾙ名
Dim target As String:   target = "伝送用ファイル" & ".accdb" 'ファイル名変更_2024.2.12

Call DebugLog("VcConv 開始 No=[" & no & "] filePath=" & FPath & target)

'ﾌｧｲﾙﾊﾟｽ
Dim FilePath As String: FilePath = FPath & target
  If ReadCheck(FilePath, 1) = False Then '1：ファイル,2：ﾌｫﾙﾀﾞ
    MsgBox "日稲:" & target & "がありません" & "_ReadCheck_VcConv", vbCritical
    Exit Function
  End If
'検索項目
Dim SQLfilter As String: SQLfilter = "[品名ｺｰﾄﾞ]='" & no & "'" 'VC検索
Dim Temp As Variant: Temp = GetRecordsArr(FilePath, TableName, SQLfilter) 'All,fields,record

Call DebugLog("VcConv GetRecordsArr結果 IsEmpty=" & IsEmpty(Temp))

'ﾌｨｰﾙﾄﾞ名ﾁｪｯｸ　*更新に対応
Dim fieldsfilter As String: fieldsfilter = ""
Dim Fields As Variant:  Fields = GetFieldsArr(FilePath, TableName, fieldsfilter)
If IsEmpty(Fields) Then 'emptyの場合停止
  MsgBox "fieldsなし" & "_VcConv"
  Exit Function
End If
Dim i As Long, C_厚係数 As Long
For i = LBound(Fields, 2) To UBound(Fields, 2)
  If Fields(1, i) = "厚-係数" Then
    C_厚係数 = i
    Exit For
  End If
Next i
If IsEmpty(Temp) Then 'emptyの場合停止
  MsgBox "VC厚登録なし,担当者に連絡のこと"
  VcConv = 0.08406977 '登録VC全平均厚_2022.12.19
  Exit Function
Else
  VcConv = Temp(1, C_厚係数)
End If
'Stop

End Function

Public Sub A_Output(FormName As Object, Lot As String, Ken As String)
Dim A_niti As Variant: A_niti = FormArray(FormName, 3) '梱包用_変換無しに変更*2021.04.01
If IsEmpty(A_niti) Then Exit Sub
Dim A_Icas As Variant: A_Icas = FormArray(FormName, 2) 'ｱｲｶｽ用
If IsEmpty(A_Icas) Then Exit Sub
Dim Lname As String: Lname = UFAinput.LineName

Call AddWS("Nitiina" & "_" & Lname)
Call AddWS("A_Icas" & "_" & Lname)
Dim SHIna As Worksheet:  Set SHIna = ThisWorkbook.Sheets("Nitiina" & "_" & Lname) '梱包用
Dim SHIcas As Worksheet: Set SHIcas = ThisWorkbook.Sheets("A_Icas" & "_" & Lname) 'ｱｲｶｽ用

SHIna.Cells.NumberFormatLocal = "@" '文字列
SHIcas.Cells.NumberFormatLocal = "@" '文字列
Dim Past_Ni As Range:  Set Past_Ni = Range(SHIna.Cells(LBound(A_niti, 1), 1), SHIna.Cells(UBound(A_niti, 1), UBound(A_niti, 2))) '梱包用
Dim Past_Inp As Range: Set Past_Inp = SHIcas.Cells(1, 1) 'ｱｲｶｽ用
Past_Ni = A_niti
Past_Inp = A_Icas
'Past_Ni = Application.WorksheetFunction.Transpose(A_niti) '行列入れ替え'梱包保存用*FormArray(UFAinput, 1)で使用
'-------------------Excel出力用
'Application.Visible = True
'A保存(SHSa.Name, Lot, ken)*梱包課用保存Excelｼｰﾄ
'Application.Visible = False
'-------------------
Call A_CsvSave(SHIna.Name, Lot, Ken) 'CSV保存
Call ExpCSV(SHIcas, "hostsend", NisPath, vbTab) 'CSV保存---保存先でAｲﾝﾌﾟｯﾄ用出力になる　*出力ﾌｧｲﾙ,ﾌｧｲﾙ名,ﾊﾟｽ,出力形式
Application.DisplayAlerts = False '保存後シートは消す
  If WsCheck("Nitiina" & "_" & Lname) Then ThisWorkbook.Sheets("Nitiina" & "_" & Lname).Delete
  If WsCheck("A_Icas" & "_" & Lname) Then ThisWorkbook.Sheets("A_Icas" & "_" & Lname).Delete
Application.DisplayAlerts = True
End Sub

'CSV出力(警告なしにファイルを上書き)
Public Sub ExpCSV(sh As Worksheet, Optional fileName As String = "", Optional FilePath As String = "", Optional Sep As String = ",", Optional maxi As Long = 0, Optional maxj As Long = 0)
Dim i As Long
Dim j As Long
If maxi = 0 Then maxi = UR(sh)
If maxj = 0 Then maxj = UC(sh)
Dim ShDb As Variant: ShDb = Range(sh.Cells(1, 1), sh.Cells(maxi, maxj))
If FilePath = "" Then FilePath = ThisWorkbook.path
If fileName = "" Then fileName = "ExportFile"
Dim ExpName As String
'character-separated values
Select Case Sep
  Case "," 'ｶﾝﾏ
    ExpName = ".csv"
  Case vbTab 'ﾀﾌﾞ
    ExpName = ".txt" '".tsv"
  Case Else
    ExpName = ".txt"
End Select
Dim MakefilePath As String: MakefilePath = FilePath & "\" & fileName & ExpName
Dim StrLink As String: StrLink = ""
If fileName = "hostsend" Then '1つ分の文字列扱い_2020.11.25
   Open MakefilePath For Output As #1
      StrLink = ShDb
   Print #1, StrLink
   Close #1
Else
   Open MakefilePath For Output As #1
   For i = 1 To maxi
     For j = 1 To maxj
       StrLink = StrLink & ShDb(i, j) & Sep '","以外にも対応"
     Next j
     Print #1, StrLink
     StrLink = ""
   Next i
   Close #1
End If
End Sub

Public Sub A_CsvSave(ws As String, Lot As String, Ken As String) 'CSV
Dim SetPath As String, MonthPath As String, fCount As Long
Application.ScreenUpdating = False
If WsCheck(ws) = False Then Exit Sub 'ｼｰﾄなし
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets(ws)
Application.ScreenUpdating = False
'************************************月毎にﾌｫﾙﾀﾞを作成してその中に格納する2018.12.19
Dim MonthName As String: MonthName = Format(Date, "yyyy.mm")
Select Case True
Case UF1.L1Work: MonthPath = HistoryPath & "L-1\L-1_Ainput\" & MonthName
Case UF1.LVC:    MonthPath = HistoryPath & "LVC\LVC_Ainput\" & MonthName
Case UF1.HVC:    MonthPath = HistoryPath & "HVC\HVC_Ainput\" & MonthName
Case UF1.Bala:   MonthPath = HistoryPath & "バランサー\バランサー_Ainput\" & MonthName
Case UF1.Tott:   MonthPath = HistoryPath & "トット\トット_Ainput\" & MonthName
Case UF1.Maru:   MonthPath = HistoryPath & "丸徳\丸徳_Ainput\" & MonthName
Case Else:       MonthPath = HistoryPath & "Test\Test_Ainput\" & MonthName 'ThisWorkbook.Path & "\" & MonthName '"C:\Users\NLM Administrator\Desktop\保存用ﾌｫﾙﾀﾞ\L-1 Ainput\" & MonthName 'ThisWorkbook.Path& MonthName
End Select

Call CreateFolder(MonthPath) 'ﾌｫﾙﾀﾞを作成_2024.5.27 不要

  If Dir(MonthPath, vbDirectory) = "" Then 'vbDirectory必要　指定すると引数setPathで指定したﾌｫﾙﾀﾞがあればﾌｫﾙﾀﾞ名を返し,なければ長さ0の文字列「""」を返します。
    MkDir MonthPath '月のﾌｫﾙﾀﾞを作成
  End If
'************************************名前を付けて保存　名前が被れば001～数字を付ける
  Dim DayKenName As String: DayKenName = Format(Date, "yyyymmdd") & "_" & Ken
  SetPath = MonthPath & "\" & Lot
  '西暦月(.)ﾌｫﾙﾀﾞ￥LotNoﾌｫﾙﾀﾞ￥西暦日付(.なし)＿検査No　2019.08.27
  If Dir(SetPath, vbDirectory) = "" Then
      MkDir SetPath 'ロットNOのﾌｫﾙﾀﾞを作成
      If Dir(SetPath & "\" & DayKenName & ".csv", vbDirectory) = "" Then 'ﾌｫﾙﾀﾞ内にNOのブックを保存
          Call ExpCSV(sh, DayKenName, SetPath)
      Else
        For fCount = 2 To 99
          If Dir(SetPath & "\" & DayKenName & "." & Right("0" & CStr(fCount), 2) & ".csv", vbDirectory) = "" Then '同じNOがある場合副番を付けて99まで保存する
            Call ExpCSV(sh, DayKenName & "." & Right("0" & CStr(fCount), 2), SetPath)
            Exit For
          End If
          If fCount = 99 Then
              MsgBox ("連番をこれ以上付けることが出来ません。"), vbCritical
            End
          End If
        Next
      End If
  Else
    If Dir(SetPath & "\" & DayKenName & ".csv", vbDirectory) = "" Then 'ﾌｫﾙﾀﾞ内にNOのブックを保存
        Call ExpCSV(sh, DayKenName, SetPath)
    Else
      For fCount = 2 To 99
        If Dir(SetPath & "\" & DayKenName & "." & Right("0" & CStr(fCount), 2) & ".csv", vbDirectory) = "" Then '同じNOがある場合副番を付けて99まで保存する
            Call ExpCSV(sh, DayKenName & "." & Right("0" & CStr(fCount), 2), SetPath)
          Exit For
        End If
        If fCount = 99 Then
            MsgBox ("連番をこれ以上付けることが出来ません。"), vbCritical
          End
        End If
      Next
    End If
  End If
Application.ScreenUpdating = True
End Sub

Public Sub A保存(ws As String, Lot As String, Ken As String) 'Excel****CSVで保存しているから使ってない
'2018.03.27　インプットデータシートを保存  AインプットにCall
'2019.08.19 読み取り専用で保存に変更
'2019.10.08 パスワード設定
'2020.07.13 保存内容変更 UF用ｶｽﾀﾏｲｽﾞ
Dim SetPath As String, MonthPath As String
Dim idx, Cnt As Long, Lcount As Long, fCount As Long
Dim wkArray()
Application.ScreenUpdating = False
Application.Visible = True 'Excel表示
ThisWorkbook.Activate
If WsCheck(ws) = False Then Exit Sub '-------------------------ｼｰﾄなし
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets(ws)
sh.Visible = True
Application.ScreenUpdating = False
sh.Copy
 Dim BK As Workbook: Set BK = ActiveWorkbook
'************************************月毎にﾌｫﾙﾀﾞを作成してその中に格納する2018.12.19
  Dim MonthName As String: MonthName = Format(Date, "yyyy.mm")
  Select Case True
  Case UF1.L1Work: MonthPath = HistoryPath & "L-1\L-1_Ainput\" & MonthName
  Case UF1.LVC:    MonthPath = HistoryPath & "LVC\LVC_Ainput\" & MonthName
  Case UF1.HVC:    MonthPath = HistoryPath & "HVC\HVC_Ainput\" & MonthName
  Case UF1.Bala:   MonthPath = HistoryPath & "バランサー\バランサー_Ainput\" & MonthName
  Case UF1.Tott:   MonthPath = HistoryPath & "トット\トット_Ainput\" & MonthName
  Case UF1.Maru:   MonthPath = HistoryPath & "丸徳\丸徳_Ainput\" & MonthName
  Case Else:       MonthPath = HistoryPath & "Test\Test_Ainput\" & MonthName 'ThisWorkbook.Path & "\" & MonthName '"C:\Users\NLM Administrator\Desktop\保存用ﾌｫﾙﾀﾞ\L-1 Ainput\" & MonthName 'ThisWorkbook.Path& MonthName
  End Select
  If Dir(MonthPath, vbDirectory) = "" Then 'vbDirectory必要　指定すると引数setPathで指定したﾌｫﾙﾀﾞがあればﾌｫﾙﾀﾞ名を返し,なければ長さ0の文字列「""」を返します。
    MkDir MonthPath '月のﾌｫﾙﾀﾞを作成
  End If
'************************************名前を付けて保存　名前が被れば001～数字を付ける
  Dim DayKenName As String: DayKenName = Format(Date, "yyyymmdd") & "_" & Ken '2019.08.27 ファイル名変更　パスワード設定に備える
  SetPath = MonthPath & "\" & Lot
  '西暦月(.)ﾌｫﾙﾀﾞ￥LotNoﾌｫﾙﾀﾞ￥西暦日付(.なし)＿検査No　2019.08.27
  If Dir(SetPath, vbDirectory) = "" Then 'vbDirectory必要　指定すると引数setPathで指定したﾌｫﾙﾀﾞがあればﾌｫﾙﾀﾞ名を返し,なければ長さ0の文字列「""」を返します。
      MkDir SetPath 'ロットNOのﾌｫﾙﾀﾞを作成
      If Dir(SetPath & "\" & DayKenName & ".xlsx", vbDirectory) = "" Then 'ﾌｫﾙﾀﾞ内にNOのブックを保存
          BK.SaveAs (SetPath & "\" & DayKenName & ".xlsx"), Password:=Ken & "." & Format(Date, "yyyy.mm.dd")  ', Password:=""'pass SH.Cells(1, 1).西暦.月.日
          SetAttr (SetPath & "\" & DayKenName & ".xlsx"), vbReadOnly
      Else
        For fCount = 2 To 99
          If Dir(SetPath & "\" & DayKenName & "." & Right("0" & CStr(fCount), 2) & ".xlsx", vbDirectory) = "" Then '同じNOがある場合副番を付けて99まで保存する
            BK.SaveAs (SetPath & "\" & DayKenName & "." & Right("0" & CStr(fCount), 2) & ".xlsx"), Password:=Ken & "." & Format(Date, "yyyy.mm.dd")    ', Password:=""'pass 検査No.西暦.月.日
            SetAttr (SetPath & "\" & DayKenName & "." & Right("0" & CStr(fCount), 2) & ".xlsx"), vbReadOnly
            Exit For
          End If
          If fCount = 99 Then
              MsgBox ("連番をこれ以上付けることが出来ません。"), vbCritical
            End
          End If
        Next
      End If
  Else
    If Dir(SetPath & "\" & DayKenName & ".xlsx", vbDirectory) = "" Then 'ﾌｫﾙﾀﾞ内にNOのブックを保存
        BK.SaveAs (SetPath & "\" & DayKenName & ".xlsx"), Password:=sh.Cells(1, 1) & "." & Format(Date, "yyyy.mm.dd")  ', Password:=""'pass SH.Cells(1, 1).西暦.月.日
        SetAttr (SetPath & "\" & DayKenName & ".xlsx"), vbReadOnly
    Else
      For fCount = 2 To 99
        If Dir(SetPath & "\" & DayKenName & "." & Right("0" & CStr(fCount), 2) & ".xlsx", vbDirectory) = "" Then '同じNOがある場合副番を付けて99まで保存する
          BK.SaveAs (SetPath & "\" & DayKenName & "." & Right("0" & CStr(fCount), 2) & ".xlsx"), Password:=Ken & "." & Format(Date, "yyyy.mm.dd")   ', Password:=""'pass 検査No.西暦.月.日
          SetAttr (SetPath & "\" & DayKenName & "." & Right("0" & CStr(fCount), 2) & ".xlsx"), vbReadOnly
          Exit For
        End If
        If fCount = 99 Then
            MsgBox ("連番をこれ以上付けることが出来ません。"), vbCritical
          End
        End If
      Next
    End If
  End If
 '************************************
  BK.Close
Application.ScreenUpdating = True
End Sub

Public Sub 履歴保存()
'印刷にCall2018.06.26'2019.08.19 読み取り専用で保存に変更2019.10.08 パスワード設定
Dim ws As Worksheet, flag As Boolean
Dim Lotname As String, MonthName As String, DayLotName As String
Dim SetPath As String, MonthPath As String
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Application.ScreenUpdating = False
Application.Visible = True 'Excel表示
Dim wkArray()
Dim idx, Cnt As Long, fCount As Long

Call DisplaySheets

'2021.01.22_ﾛｰｶﾙ,工場共有にPDF保存
Dim KouPath As String: KouPath = "C:\検査履歴" '"\\nlmsrvngy03\工場内共有\検査データ"

If Dir(KouPath, vbDirectory) = "" Then MkDir KouPath

Dim LN As String
Dim saveSh As Worksheet:    Set saveSh = ThisWorkbook.Sheets("印刷用" & "_" & Lname)
Dim Defective As Worksheet: Set Defective = ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname)

Select Case True
Case UF1.L1Work: LN = "L-1"
Case UF1.LVC:    LN = "LVC"
Case UF1.HVC:    LN = "HVC"
Case UF1.Bala:   LN = "大板"
Case UF1.Tott:   LN = "大板"
Case UF1.Maru:   LN = "中板"
Case Else:       LN = "Test"
End Select
Dim F As String:        F = SaveDateLocate(KouPath & "\", LN)   'ﾊﾟｽ,ﾌｧｲﾙﾗｲﾝ名_必要ﾌｫﾙﾀﾞ生成とﾌｧｲﾙﾊﾟｽ
Dim Kouf As String:     Kouf = SaveDateLocate(FactoryyPath, "【梱包】" & LN)
'印刷帳票
Call PDFCREATE(F, saveSh, saveSh.Cells(2, 3)) 'ﾊﾟｽ,保存ｼｰﾄ,保存ﾌｧｲﾙ名
Call PDFCREATE(Kouf, saveSh, saveSh.Cells(2, 3)) 'ﾌｧｲﾙｺﾋﾟｰやめた_2025.1.17

'Dim NewFiles As String: NewFiles = NewFileName(F) '生成後一番新しいﾌｧｲﾙ名
'FileCopy F & "\" & NewFiles, Kouf & "\" & NewFiles 'ﾌｧｲﾙｺﾋﾟｰ

'不良要因_2025.1.17*品証依頼不良要因もPDF化
Call PDFCREATE(F, Defective, Defective.Cells(1, 2) & "_不良要因")
'Debug.Print "1つ目のPDF作成: " & F & "\" & fileName
DoEvents ' ← ここで一度待つ
Call PDFCREATE(Kouf, Defective, Defective.Cells(1, 2) & "_不良要因")
'Debug.Print "2つ目のPDF作成: " & Kouf & "\" & fileName

'Dim NewFiles2 As String:  NewFiles2 = NewFileName(F) '生成後一番新しいﾌｧｲﾙ名
'Dim fileP As String: fileP = F & "\" & NewFiles2
'' ファイルが作成されるまで待機
'Do While Dir(fileP) = ""
'  DoEvents ' 他の処理を実行
'Loop
'FileCopy F & "\" & NewFiles2, Kouf & "\" & NewFiles2

'原本(Excelﾌｧｲﾙ)保存
With ThisWorkbook
  For idx = 1 To .Worksheets.Count '2021.01.25_校正記録,品証検証,"Daily" & "_" & Lname 除く
    If (.Worksheets(idx).Name) = "印刷用" & "_" & Lname Or (.Worksheets(idx).Name) = "Lot情報履歴" Or _
    (.Worksheets(idx).Name) = "WorkerTime" & "_" & Lname Or _
    (.Worksheets(idx).Name) = "作業保存" & "_" & Lname Or _
    (.Worksheets(idx).Name) = "Daily" & "_" & Lname Or _
    (.Worksheets(idx).Name) = "不良要因印刷" & "_" & Lname Then  '2021.07.22_作業保存ｼｰﾄ追加保存**2023.05.229_不良要因印刷追加保存
      Cnt = Cnt + 1
      ReDim Preserve wkArray(1 To Cnt)
      wkArray(Cnt) = .Worksheets(idx).Name
    End If
  Next idx
End With
If Cnt = 0 Then '保存シートがないとエラーになる2018.12.19
  Exit Sub
Else
  Application.ScreenUpdating = False
  ThisWorkbook.Worksheets(wkArray).Copy
End If
Dim BK As Workbook: Set BK = ActiveWorkbook '新規ブック側
'月毎にﾌｫﾙﾀﾞを作成してその中に格納する2018.12.19
'名前を付けて保存　名前が被れば001～数字を付ける
  MonthName = Format(Date, "yyyy.mm")
  MonthPath = "C:\検査履歴" & "\" & MonthName '-----------------------ﾛｰｶﾙ保存_2021.01.25
  Lotname = ThisWorkbook.Sheets("印刷用" & "_" & Lname).Cells(2, 3) '"検査履歴"
  DayLotName = Format(Date, "yyyymmdd") & "_" & Lotname
  SetPath = MonthPath & "\" & DayLotName '2019.08.27 ファイル名変更　パスワード設定に備える
  
  Call CreateFolder(SetPath) 'ﾌｫﾙﾀﾞを作成_2024.5.27 不要
  
  Dim FilePath As String:     FilePath = SetPath & "\" & DayLotName & ".xlsx"
  
  '西暦月(.)ﾌｫﾙﾀﾞ￥西暦日付(.なし)_LotNo　2019.08.27_2021.01.22,ﾊﾟｽﾜｰﾄﾞ変更
  If Dir(FilePath, vbDirectory) = "" Then
      BK.SaveAs (FilePath), Password:=Lotname & "." & Format(Date, "yyyy.mm.dd") ', Password:=""'pass LotNo.西暦.月.日
      SetAttr (FilePath), vbReadOnly
  Else
    For fCount = 2 To 999
      If Dir(SetPath & "\" & DayLotName & "." & Right("00" & CStr(fCount), 3) & ".xlsx", vbDirectory) = "" Then  '同じ場合副番を付けて999まで保存する
        BK.SaveAs (SetPath & "\" & DayLotName & "." & Right("00" & CStr(fCount), 3) & ".xlsx"), Password:=Lotname & "." & Format(Date, "yyyy.mm.dd")  ', Password:=""'pass LotNo.西暦.月.日
        SetAttr (SetPath & "\" & DayLotName & "." & Right("00" & CStr(fCount), 3) & ".xlsx"), vbReadOnly
        Exit For
      End If
      If fCount = 999 Then
        MsgBox ("連番をこれ以上付けることが出来ません。"), vbCritical
        End
      End If
    Next
  End If
 '************************************
BK.Close
ThisWorkbook.Activate

Call HiddenSheets

'--------------------ﾌｧｲﾙｺﾋﾟｰ_処理ｽﾋﾟｰﾄﾞ速い_2021.01.25
Dim SrvPath As String
Select Case True
Case UF1.L1Work: SrvPath = HistoryPath & "L-1\L-1_履歴\" & MonthName
Case UF1.LVC:    SrvPath = HistoryPath & "LVC\LVC_履歴\" & MonthName
Case UF1.HVC:    SrvPath = HistoryPath & "HVC\HVC_履歴\" & MonthName
Case UF1.Bala:   SrvPath = HistoryPath & "バランサー\バランサー_履歴\" & MonthName
Case UF1.Tott:   SrvPath = HistoryPath & "トット\トット_履歴\" & MonthName
Case UF1.Maru:   SrvPath = HistoryPath & "丸徳\丸徳_履歴\" & MonthName
Case Else:       SrvPath = HistoryPath & "Test\Test_履歴\" & MonthName 'ThisWorkbook.Path & "\" & MonthName '"C:\Users\NLM Administrator\Desktop\保存用ﾌｫﾙﾀﾞ\L-1 履歴\" & MonthName
End Select

Call CreateFolder(SrvPath) 'ﾌｫﾙﾀﾞを作成

Dim NewFiles As String: NewFiles = NewFileName(SetPath & "\") '生成後一番新しいﾌｧｲﾙ名 フォルダパスまでを渡す
FileCopy SetPath & "\" & NewFiles, SrvPath & "\" & NewFiles
Application.ScreenUpdating = True
End Sub
Public Function SaveDateLocate(DatePath As String, Optional LineN As String = "Datafolda", Optional n As String = "_") As String 'ﾌｫﾙﾀﾞ生成とﾊﾟｽ吐出し'ﾊﾟｽ,ﾌｧｲﾙﾗｲﾝ名
'ﾌｫﾙﾀﾞ処理_2021.01.26
If LineN = "" Then Exit Function
Dim SetPath As String: SetPath = DatePath & LineN
  If Dir(SetPath, vbDirectory) = "" Then 'vbDirectory必要　指定すると引数setPathで指定したﾌｫﾙﾀﾞがあればﾌｫﾙﾀﾞ名を返し,なければ長さ0の文字列「""」を返します。
      MkDir SetPath 'ラインﾌｫﾙﾀﾞを作成
  End If
Dim MonthName As String: MonthName = Format(Date, "yyyy.mm")
Dim MonthPath As String: MonthPath = SetPath & "\" & MonthName
  If Dir(MonthPath, vbDirectory) = "" Then
      MkDir MonthPath '月ﾌｫﾙﾀﾞを作成
  End If
Dim DayName As String: DayName = Format(Date, "dd日")
Dim DayPath As String: DayPath = MonthPath & "\" & DayName
  If Dir(DayPath, vbDirectory) = "" Then
      MkDir DayPath '日ﾌｫﾙﾀﾞを作成
  End If
SaveDateLocate = DayPath
End Function
'フォルダ構成からフォルダを一括生成
Function CreateFolder(FullFolderPath As String) As Boolean
Dim FolderNames() As String: FolderNames = Split(FullFolderPath, "\")

On Error Resume Next
Dim FolderPath As String
FolderPath = FolderNames(0)
Dim i As Long
For i = 1 To UBound(FolderNames)
  FolderPath = FolderPath & "\" & FolderNames(i)
  MkDir FolderPath
Next
On Error GoTo 0
If Dir(FullFolderPath, vbDirectory) <> "" Then CreateFolder = True
End Function

Public Function FormArray(FormName As Object, n As Long) As Variant 'ﾌｫｰﾑ値を配列化 ﾌｧｲﾙﾚｲｱｳﾄNoを変数と合わせる'配列作成  Aｲﾝﾌﾟｯﾄ出力用配列
Select Case FormName.Name
Case "UFAinput"
  With FormName
    Dim A0 As String:   A0 = .KenLaw.Caption '検査法
    Dim A3 As String:   A3 = ZerCon(3, .HikiNo.Text) '引当No
    Dim A4 As String:   A4 = ZerCon(4, .KONNO.Text) '梱包No
    Dim A5 As String:   A5 = ZerCon(5, .Remain.Text) '出荷ｽﾄｱ
    
    Dim A12 As String:  A12 = ZerCon(12, .MAI1.Text) '枚数
    Dim A14 As String:  A14 = ZerCon(14, .MAI2.Text)
    Dim A13 As String:  A13 = ZerCon(13, .COU1.Text) '包み数
    Dim A15 As String:  A15 = ZerCon(15, .COU2.Text)
    
    Dim A17 As String:  A17 = ZerCon(17, .Material1.Text) '風袋
    Dim A19 As String:  A19 = ZerCon(19, .Material2.Text)
    Dim A21 As String:  A21 = ZerCon(21, .Material3.Text)
    
    Dim A16 As String:  A16 = ZerCon(16, .GW.Text) 'GW
    
    Dim A18 As String:  A18 = ZerCon(18, "") 'GW2
    Dim A20 As String:  A20 = ZerCon(20, "") 'GW3
    
    Dim A22 As String:  A22 = ZerCon(22, "") '正味M3
    Dim A23 As String:  A23 = ZerCon(23, "") '風袋M3
    Dim A24 As String:  A24 = ZerCon(24, "") '梱包ｺｰﾄﾞ
    
    Dim A32 As String:  A32 = ZerCon(32, "") 'ｺｲﾙ長さ
    Dim A33 As String:  A33 = ZerCon(33, "") 'ｺｲﾙ内径
    Dim A34 As String:  A34 = ZerCon(34, "") 'ｺｲﾙ肉厚
    Dim A35 As String:  A35 = ZerCon(35, "") 'ｺｲﾙ本数
    Dim A36 As String:  A36 = ZerCon(36, "") 'ｺｲﾙ巻数
    
    Dim A40 As String:  A40 = ZerCon(40, "") '×
    
    Dim A30 As String:  A30 = ZerCon(30, .PileMatchNo.Text) '積み合せ1,9
    Dim A31 As String:  A31 = ZerCon(31, .PileMatch.Text) '積み合せNo
    
    Dim A25 As String:  A25 = ZerCon(25, .ExCaseNo.Text) 'EXｹｰｽNo
    Dim A26 As String:  A26 = ZerCon(26, .ExMaterial1.Text) 'Ex風袋
    Dim A27 As String:  A27 = ZerCon(27, .ExMaterial2.Text)
    Dim A28 As String:  A28 = ZerCon(28, .ExMaterial3.Text)
    Dim A29 As String:  A29 = ZerCon(29, .ExGW.Text) 'EXGW
    Dim A37 As String:  A37 = ZerCon(37, .OutsideAtu.Text) 'EX外寸
    Dim A38 As String:  A38 = ZerCon(38, .OutsideHaba.Text)
    Dim A39 As String:  A39 = ZerCon(39, .OutsideTake.Text)
    
    Dim A6 As String:   If Not .AMin.Caption = "" Then A6 = StCon(.AMin.Caption, "厚") 'MinMax　0埋め変換
    Dim A9 As String:   If Not .AMax.Caption = "" Then A9 = StCon(.AMax.Caption, "厚")
    Dim A7 As String:   If Not .HMin.Caption = "" Then A7 = StCon(.HMin.Caption, "巾")
    Dim A10 As String:  If Not .HMax.Caption = "" Then A10 = StCon(.HMax.Caption, "巾")
    Dim A8 As String:   If Not .TMin.Caption = "" Then A8 = StCon(.TMin.Caption, "丈")
    Dim A11 As String:  If Not .TMax.Caption = "" Then A11 = StCon(.TMax.Caption, "丈")
    
    Dim line As String: line = LinConv(.LineName.Caption) 'ﾗｲﾝ名 NIS指定に変換
    Dim NIS1 As String: NIS1 = "KEN:AIN:" '固定値
    Dim NIS2 As String: NIS2 = "0185"
    
'    Dim Height1 As String: Height1 = .K_Height.Text '2021.04.01_追加
'    Dim Height2 As String: Height2 = .K_Height_2.Text
'    Dim Height3 As String: Height3 = .K_Height_3.Text
'    Dim Height4 As String: Height4 = .K_Height_4.Text
    '↑の修正版_2026.05.29' 高さ配列を動的取得
      Dim hIdx As Long
      Dim HeightVals(1 To MAX_MOUNT) As String
      Dim HeightKeys(1 To MAX_MOUNT) As String
      For hIdx = 1 To MAX_MOUNT
          HeightVals(hIdx) = GetKHeight(hIdx)
          HeightKeys(hIdx) = CStr(41 + hIdx) & "_梱包高さ" & hIdx
      Next hIdx
  End With
  Select Case n
  Case 1 '梱包課用
    FormArray = Array(NIS1, NIS2, A3, A4, A5, A6, A7, A8, A9, A10, _
    A11, A12, A13, A14, A15, A16, A17, A18, A19, A20, _
    A21, A22, A23, A24, A25, A26, A27, A28, A29, A30, _
    A31, A32, A33, A34, A35, A36, A37, A38, A39, A40, line)
  Case 2 '配列作成  Aｲﾝﾌﾟｯﾄ出力用配列　ここで必要なもの追加*ﾌｧｲﾙﾚｲｱｳﾄ参照**\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\【4】_Ainput\システム間インターフェース定義.xlsm
    FormArray = Array(NIS1 & NIS2 & A3 & A4 & A5 & A6 & A7 & A8 & A9 & A10 _
    & A11 & A12 & A13 & A14 & A15 & A16 & A17 & A18 & A19 & A20 _
    & A21 & A22 & A23 & A24 & A25 & A26 & A27 & A28 & A29 & A30 _
    & A31 & A32 & A33 & A34 & A35 & A36 & A37 & A38 & A39 & A40 & line)
  Case 3 '変換なし
  Dim a, b, Fields, Temp
    With FormName
      a = Array(NIS1, NIS2, .HikiNo.Text, .KONNO.Text, .Remain.Text, .AMin.Caption, .HMin.Caption, .TMin.Caption, .AMax.Caption, .HMax.Caption, .TMax.Caption, _
        .MAI1.Text, .COU1.Text, .MAI2.Text, .COU2.Text, .GW.Text, .Material1.Text, "0", .Material2.Text, "0", .Material3.Text, _
        "0", "0", "0", .ExCaseNo.Text, .ExMaterial1.Text, .ExMaterial2.Text, .ExMaterial3.Text, .ExGW.Text, A30, A31, _
        "0", "0", "0", "0", "0", .OutsideAtu.Text, .OutsideHaba.Text, .OutsideTake.Text, "0", line, _
        HeightVals(1), HeightVals(2), HeightVals(3), HeightVals(4), HeightVals(5), HeightVals(6), _
        HeightVals(7), HeightVals(8), HeightVals(9), HeightVals(10), HeightVals(11), HeightVals(12), _
        HeightVals(13), HeightVals(14), HeightVals(15))
      
      b = Array("1_固定値", "2_固定値", "3_引当No", "4_梱包No", "5_出荷ｽﾄｱ", "6_厚Min", "7_巾Min", "8_丈Min", "9_厚Max", "10_巾Max" _
        , "11_丈Max", "12_枚数1", "13_包数1", "14_枚数2", "15_包数2", "16_GW1", "17_風袋1", "18_GW2_x", "19_風袋2", "20_GW3_x" _
        , "21_風袋3", "22_正味_x", "23_風袋_x", "24_梱包ｺｰﾄﾞ_x", "25_EXｹｰｽNo", "26_Ex風袋1", "27_Ex風袋2", "28_Ex風袋3", "29_EXGW", "30_積み合せｺｰﾄﾞ" _
        , "31_積み合せNo", "32_ｺｲﾙ長さ_x", "33_ｺｲﾙ内径_x", "34_ｺｲﾙ肉厚_x", "35_ｺｲﾙ本数_x", "36_ｺｲﾙ巻数_x", "37_EX外寸高さ", "38_EX外寸奥", "39_EX外寸長さ", "40_x" _
        , "41_ﾗｲﾝ名" _
        , "42_梱包高さ1", "43_梱包高さ2", "44_梱包高さ3", "45_梱包高さ4", "46_梱包高さ5" _
        , "47_梱包高さ6", "48_梱包高さ7", "49_梱包高さ8", "50_梱包高さ9", "51_梱包高さ10" _
        , "52_梱包高さ11", "53_梱包高さ12", "54_梱包高さ13", "55_梱包高さ14", "56_梱包高さ15") '2021.04.01*高さ検算記入をなくすためこちらで記録する’2026.05.29_動的取得
      
      
'      Temp = Array2D(a)
'      fields = Array2D(b) '2次元変換’0配列0が消える為下記に変更_2021.09.16

      Dim Last_a As Long: Last_a = UBound(a) + 1
      Dim Last_b As Long: Last_b = UBound(b) + 1
      ReDim Temp(1 To Last_a, 1) '2次元変換*配列合体のために必要
      ReDim Fields(1 To Last_b, 1)
      Dim i As Long
      For i = LBound(a) To UBound(a)
        Temp(i + 1, 1) = a(i)
      Next i
      For i = LBound(b) To UBound(b)
        Fields(i + 1, 1) = b(i)
      Next i
      
      FormArray = MergeVariantArraysCol(Fields, Temp) '列方向配列合体
    End With
  End Select
End Select
End Function

Public Function StCon(Temp As Double, Pits As String) As String '0埋めの7桁に変換
If Temp = 0 Then Exit Function
Dim Cal As Double, Ket As String, buf
Select Case Pits
Case "厚"
  Cal = Temp * 1000
  Ket = "0000"
Case "巾"
  Cal = Temp * 100
  Ket = "00000"
Case "丈"
  Cal = Temp * 10
  Ket = "000000"
End Select
buf = Cal
StCon = Right(Ket & buf, 7)
End Function

Public Function LinConv(line As String) As String 'ﾗｲﾝ名変換_別紙**\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\【4】_Ainput\システム間インターフェース定義.xlsm
If line = "" Then Exit Function
Select Case line
Case "L-1":    LinConv = "05"
Case "LVC":    LinConv = "03"
Case "HVC":    LinConv = "04"
Case "ﾊﾞﾗﾝｻｰ": LinConv = "01" '2021.03.29*01と02入れ替え
Case "ﾄｯﾄ":    LinConv = "01" '"02"'*2021.05.11_1excelに対し1ｱｲｶｽとの仕様の為ﾄｯﾄとﾊﾞﾗﾝｻｰを統合
Case "LS4":    LinConv = "06"
Case "NS1":    LinConv = "07"
Case "Test":   LinConv = "AA" '"TestNo"
Case Else:     LinConv = "Maru"
End Select
End Function

Public Function ZerCon(DateNo As Long, buf As String) As String '0埋め　計算不要
Select Case buf
Case ""
  Select Case DateNo
  Case 3:                                             ZerCon = "00000000" '8
  Case 4, 41:                                         ZerCon = "00" '2
  Case 5:                                             ZerCon = "0" '1
  Case 6 To 11:                                       ZerCon = "0000000" '7
  Case 12, 14:                                        ZerCon = "00000"  '5
  Case 16 To 21:                                      ZerCon = "00000" '5
  
  Case 24:                                            ZerCon = "00000" '5
  
  Case 25:                                            ZerCon = "     " '半角スペース EXｹｰｽ_2025.1.7
  
  Case 32, 37, 38, 39:                                ZerCon = "00000" '5
  Case 13, 15, 33, 34, 35:                            ZerCon = "000" '3
  Case 22, 23, 26, 27, 28, 29, 36:                    ZerCon = "0000" '4
  Case 30:                                            ZerCon = " " '積み合せｺｰﾄﾞ1
  Case 31:                                            ZerCon = "          " '積み合せｺｰﾄﾞ10
  Case 40:                                            ZerCon = "000000" '6
  End Select
Case Else
  Select Case DateNo
  Case 3:                                             ZerCon = buf 'Right("0000000" & buf, 8) '8
  Case 4, 41:                                         ZerCon = Right("0" & buf, 2) '2
  Case 5:                                             ZerCon = buf '1
  Case 12, 14, 16, 17, 18, 19, 20, 21, 24, 25, 32:    ZerCon = Right("0000" & buf, 5) '5
  Case 37 To 39:                                      ZerCon = Right("0000" & buf, 5) '5
  Case 13, 15, 33, 34, 35:                            ZerCon = Right("00" & buf, 3) '3
  Case 22, 23, 26, 27, 28, 29, 36:                    ZerCon = Right("000" & buf, 4) '4
  Case 30:                                            ZerCon = buf '1
  Case 31:                                            ZerCon = Left(buf & "         ", 10) '10 左詰め
  Case 40:                                            ZerCon = Right("00000" & buf, 6) '6
  End Select
End Select
End Function

Public Function HostCon(tMP As String) As String  'ﾎｽﾄ処理結果変換
'8,4,4　「KEN:RES:」固定_「0004」固定
Dim HostNo As String: HostNo = Right(tMP, 4)
Select Case HostNo
Case "A000": HostCon = "正常に送信が行われました。"
Case "A001": HostCon = "作成されたテキストの内容が正しくありません。品質保証課へ連絡。" '品質保証課
Case "A002": HostCon = "ホストへの送信内容が正しくありません。NISへ連絡。" 'NIS"
Case "A003": HostCon = "ホストにてAインプットデータの登録に失敗しました。NISへ連絡。" 'NIS
End Select
End Function

Public Function ResultCon(tMP As Variant) As String  'ﾎｽﾄ処理結果変換_2021.09.21送信失敗コメント変更
Select Case tMP
Case "正常に送信が行われました。":                                      ResultCon = tMP '"A000"
Case "作成されたテキストの内容が正しくありません。品質保証課へ連絡してください。":  ResultCon = "作成されたﾃｷｽﾄの内容が正しくありません。担当者へ連絡して下さい。" & vbCrLf & "A001" '品質保証課"A001"
Case "ホストへの送信内容が正しくありません。NISへ連絡してください。":               ResultCon = "ﾎｽﾄへの送信内容が正しくありません。担当者へ連絡して下さい。" & vbCrLf & "A002" 'NIS"A002"
Case "ホストへのAインプットデータの送信に失敗しました。NISへ連絡してください。":    ResultCon = "ﾎｽﾄにてAｲﾝﾌﾟｯﾄﾃﾞｰﾀの送信に失敗しました。担当者へ連絡して下さい。" & vbCrLf & "A003" 'NIS"A003"
Case Else:                                                                          ResultCon = tMP
End Select
End Function

Public Function HostDataConfi(Optional APPPATH As String = "C:\Tools\測定データ電子化\", Optional Ltext As String = "hostsend.txt") As String '現状出力確認
Dim LineBuff, buf As String
If ReadCheck(APPPATH & Ltext, 1) = False Then Exit Function
LineBuff = Loadtext(APPPATH & Ltext) '関数化 文字化け対策
buf = LineBuff
HostDataConfi = "現添付Data【" & "引当No_" & Mid(buf, 13, 8) & " " & "梱包No_" & Mid(buf, 21, 2) & "】"
End Function

Public Sub InterRuptText(Optional APPPATH As String = "C:\Tools\測定データ電子化\", Optional Ltext As String = "result.txt", Optional WaitTime As Long = 0) 'ホスト取り込み結果出力
Dim LineBuff, buf As String
'Application.Wait (Now + TimeValue("0:00:0" & WaitTime)) '送信タイムアウト時間
UFAinput.LoadTimeText.Caption = "ｱｲｶｽ返答待ち....."
UFAinput.Repaint
Select Case WaitTime
Case 5000
  Sleep 1000
  UFAinput.B1.BackColor = RGB(51, 102, 255)
  UFAinput.Repaint
  Sleep 1000
  UFAinput.B2.BackColor = RGB(51, 102, 255)
  UFAinput.Repaint
  Sleep 1000
  UFAinput.B3.BackColor = RGB(51, 102, 255)
  UFAinput.Repaint
  Sleep 1000
  UFAinput.B4.BackColor = RGB(51, 102, 255)
  UFAinput.Repaint
  Sleep 1000
  UFAinput.B5.BackColor = RGB(51, 102, 255)
  UFAinput.Repaint
Case Else
  Sleep WaitTime
End Select
If ReadCheck(APPPATH & Ltext, 1) = False Then Exit Sub
'LineBuff = ResultCon(LoadText(AppPath & Ltext))   '関数化 文字化け対策
Open APPPATH & Ltext For Input As #1 'エンコード依存で文字化けする
  Do Until EOF(1)
    Line Input #1, LineBuff
  Loop
Close #1
Select Case LineBuff
Case "正常に送信が行われました。"
  UFAinput.LoadTimeText.Caption = LineBuff
  UFAinput.B1.BackColor = RGB(255, 255, 255)
  UFAinput.B2.BackColor = RGB(255, 255, 255)
  UFAinput.B3.BackColor = RGB(255, 255, 255)
  UFAinput.B4.BackColor = RGB(255, 255, 255)
  UFAinput.B5.BackColor = RGB(255, 255, 255)
  UFAinput.Repaint
  Sleep 1000
Case Else
  MsgBox ResultCon(LineBuff), vbExclamation '2021.03.19
  If UFProgress.Visible Then Unload UFProgress
  UFAinput.Ngflag = "stop" '-----------------------------処理停止フラグ
  UFAinput.B1.BackColor = RGB(255, 255, 255)
  UFAinput.B2.BackColor = RGB(255, 255, 255)
  UFAinput.B3.BackColor = RGB(255, 255, 255)
  UFAinput.B4.BackColor = RGB(255, 255, 255)
  UFAinput.B5.BackColor = RGB(255, 255, 255)
  UFAinput.LoadTimeText.Caption = ""
  UFAinput.Repaint
  Exit Sub
End Select
'UFAinput.LoadTimeText.Caption = ""'2021.06.04_MouseMoveで実行に変更
'UFAinput.Repaint
End Sub

Public Function Loadtext(fileName, Optional encoding = "UTF-8") As String '文字化け対策 参照設定　Microsoft ActiveX Date Objectx.x Library
On Error Resume Next
With CreateObject("ADODB.Stream")
  .Open
  .Charset = encoding
  .LoadFromFile fileName
  Loadtext = .ReadText
  .Close
End With
If Err.Number <> 0 Then
  Loadtext = ""
End If
End Function

Public Sub App_Activation(Optional APPPATH As String = "C:\Tools\測定データ電子化\SokuteiDataTrans\", Optional App As String = "SokuteiDataTrans.exe") '呼び出す送信用アプリ名
Dim rc As Long
If ReadCheck(APPPATH & App, 1) = False Then
  MsgBox APPPATH & App & "を開くことができません。" & vbCrLf & "担当者に連絡してください", vbCritical
  UFAinput.Ngflag = "stop" '-----------------------------処理停止フラグ
  Exit Sub
End If
rc = Shell(APPPATH & App, vbNormalFocus)
If rc = 0 Then
MsgBox "失敗"
UFAinput.Ngflag = "stop"
End If
End Sub

'--- 2つのVariant配列を結合する（列方向） ---'
Public Function MergeVariantArraysCol(vArray1 As Variant, vArray2 As Variant) As Variant
'結合する配列のサイズ
Dim vArray1_row As Long
Dim vArray1_col As Long
Dim vArray2_row As Long
Dim vArray2_col As Long
vArray1_row = UBound(vArray1, 1)
vArray1_col = UBound(vArray1, 2)
vArray2_row = UBound(vArray2, 1)
vArray2_col = UBound(vArray2, 2)

'結合後の配列のサイズ
Dim newArray_row As Long
Dim newArray_col As Long
newArray_row = Application.WorksheetFunction.Max(vArray1_row, vArray2_row)
newArray_col = vArray1_col + vArray2_col

'結合後の配列
Dim NewArray As Variant
ReDim NewArray(1 To newArray_row, 1 To newArray_col)

'配列を結合する
Dim i As Long
Dim j As Long
For j = 1 To newArray_col
    If (j <= vArray1_col) Then
        For i = 1 To newArray_row
            If (i <= vArray1_row) Then
                NewArray(i, j) = vArray1(i, j)
            Else
                NewArray(i, j) = Empty
            End If
        Next i
    Else
        For i = 1 To newArray_row
            If (i <= vArray2_row) Then
                NewArray(i, j) = vArray2(i, j - vArray1_col)
            Else
                NewArray(i, j) = Empty
            End If
        Next i
    End If
Next j
MergeVariantArraysCol = NewArray
End Function

'■2個の二次元配列を行方向(縦方向)に結合(マージ)する
Public Function Call_MergeArray_Row(arr1 As Variant, arr2 As Variant) As Variant
     
    '■結合(マージ)後の配列サイズ
    '■■行方向(縦)に結合、列方向(横)は二次元配列の大きい方に合わせる。
    Dim ROW_NEW As Long
    Dim COL_NEW As Long
    ROW_NEW = UBound(arr1, 1) + UBound(arr2, 1)
    COL_NEW = Application.WorksheetFunction.Max(UBound(arr1, 2), UBound(arr2, 2))
     
    '■結合(マージ)後の二次元配列
    Dim NewArr As Variant
    ReDim NewArr(1 To ROW_NEW, 1 To COL_NEW)
     
    '■二次元配列を結合処理
    Dim i As Long
    Dim j As Long
    For i = 1 To ROW_NEW
        If i <= UBound(arr1, 1) Then
            For j = 1 To COL_NEW
                If j <= UBound(arr1, 2) Then
                    NewArr(i, j) = arr1(i, j)
                Else
                    NewArr(i, j) = Empty
                End If
            Next j
        Else
            For j = 1 To COL_NEW
                If j <= UBound(arr2, 2) Then
                    NewArr(i, j) = arr2(i - UBound(arr1, 1), j)
                Else
                    NewArr(i, j) = Empty
                End If
            Next j
        End If
    Next i
     
    Call_MergeArray_Row = NewArr
     
End Function

'1次元配列の強制2次元配列化(0ｽﾀｰﾄ配列の場合、0値は消滅する)
Public Function Array2D(InArray As Variant, Optional rc As String = "R", Optional n As Long = 0) As Variant
'N:ListViewﾀｲﾄﾙ設定用ｵﾌﾟｼｮﾝ, 0:0値消滅
Dim Temp() As Variant

Select Case rc
  Case "R"
    Dim i As Long
    Dim maxi As Long: maxi = UBound(InArray)
    ReDim Temp(1 To maxi + n, 1 To 1) As Variant
    For i = 1 To maxi + n
      Temp(i, 1) = InArray(i - n)
    Next i
  Case "C"
    Dim j As Long
    Dim maxj As Long: maxj = UBound(InArray)
    ReDim Temp(1 To 1, 1 To maxj + n) As Variant
    For j = 1 To maxj + n
      Temp(1, j) = InArray(j - n)
    Next j
End Select

Array2D = Temp

End Function

' K_Height系テキストボックスの値を番号指定で取得
Public Function GetKHeight(n As Long) As String
    On Error Resume Next
    Dim lbl As MSForms.Label
    Set lbl = GetDynHeightLabel(n)
    If Not lbl Is Nothing Then GetKHeight = lbl.Caption
End Function

' K_Height系テキストボックスの値を番号指定でセット
Public Sub SetKHeight(n As Long, val As String)
    On Error Resume Next
    Dim lbl As MSForms.Label
    Set lbl = GetDynHeightLabel(n)
    If Not lbl Is Nothing Then lbl.Caption = val
End Sub

' 動的ラベルの取得_Ainputの山高さ
Public Function GetDynHeightLabel(n As Long) As MSForms.Label
    On Error Resume Next
    Dim ctrl As MSForms.Control
    Dim targetName As String: targetName = "DynHL_" & n
    For Each ctrl In UFAinput.Frame3.Controls
        If ctrl.Name = targetName Then
            Set GetDynHeightLabel = ctrl
            Exit Function
        End If
    Next ctrl
End Function

' 動的領域含む合計山高さ計算（既存Mount_Heightロジック置き換え用）
Public Function CalcMountHeight(mountCount As Long) As Single
    Dim total As Single: total = 0
    Dim i As Long
    For i = 1 To mountCount
        total = total + val(GetKHeight(i))
    Next i
    CalcMountHeight = total
End Function

' 動的領域含む「全Height空白チェック」
' 全て埋まっていればTrue
Public Function AllHeightFilled(mountCount As Long) As Boolean
    Dim i As Long
    For i = 1 To mountCount
        If GetKHeight(i) = "" Then
            AllHeightFilled = False
            Exit Function
        End If
    Next i
    AllHeightFilled = True
End Function

' 動的領域含む「次の空きインデックス取得」 -1なら全埋まり
Public Function NextEmptyHeightIndex(mountCount As Long) As Long
    Dim i As Long
    For i = 1 To mountCount
        If GetKHeight(i) = "" Then
            NextEmptyHeightIndex = i
            Exit Function
        End If
    Next i
    NextEmptyHeightIndex = -1
End Function

' 全Heightクリア（1?MAX_MOUNT）
Public Sub ClearAllKHeight()
    Dim i As Long
    For i = 1 To MAX_MOUNT
        SetKHeight i, ""
    Next i
End Sub
