Attribute VB_Name = "【UFKen_02】FormRelat"
Option Explicit
Public Const THIN_PLATE_ATU As Double = 6# '薄板判定閾値(mm)_6mm以下でAtuFlag=True

'Windows API宣言
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000

#If Win64 Then
  Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
  Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long
  Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
#Else
  Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
  Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Private Declare Function GetActiveWindow Lib "user32" () As Long
  Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
#End If
 
' Excelウィンドウ状態退避用
Private g_sngTopSV As Single ' 上端位置
Private g_sngLeftSV As Single ' 左端位置
Private g_sngWidthSV As Single ' 幅
Private g_sngHeightSV As Single ' 高さ

#If VBA7 And Win64 Then
  '64ビット版
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
  '32ビット版
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public UfZoom As Long
Public UfHeight As Long
Public UfWidth As Long

Public UFsizeWidth As Long
Public UFsizeHeight As Long
Public UFsizeZoom As Long
Public g_FormHwnd As Long

Public g_SkipLotNotFoundMsg As Boolean 'キャッシュクリア直後の1回だけMsgBoxを抑制_2026.08.06

Public Function GetScreenWidth() As Long
    GetScreenWidth = GetSystemMetrics32(0)
End Function

Public Function GetScreenHeight() As Long
    GetScreenHeight = GetSystemMetrics32(1)
End Function

Public Sub GetZoom(myForm As Object)
UfZoom = myForm.Zoom
UfHeight = myForm.Height
UfWidth = myForm.Width
End Sub

'フォームサイズをいい感じにする_2024.5.31　返しはサイズ、なくてもいい
Public Function InitialSetting(myForm As Object, Optional n As String = "") As Variant ', FormZoom As Long, FormHeight As Long
On Error Resume Next
Application.WindowState = xlMaximized
Dim NowWidth As Long:  NowWidth = myForm.Width
Dim NowHeight As Long: NowHeight = myForm.Height
Dim NowZoom As Long:   NowZoom = myForm.Zoom
Dim UfName As String:  UfName = myForm.Name
 
If n = "サイズ保存" Then  '2024.8.5
  myForm.Zoom = UFsizeZoom
  myForm.Width = UFsizeWidth
  myForm.Height = UFsizeHeight
  GoTo SkipA
End If

If n = "リセット" Then
  myForm.Zoom = 95
  myForm.Height = 545
  myForm.Width = 887
  GoTo SkipA
End If
 
If (UfHeight < 545) Or (UfWidth < 887) And (NowZoom >= 115 And UfName = "UF1") Then   'リセットUF1限定
 UfZoom = 95
 UfHeight = 545
 UfWidth = 887
End If

Dim VcHeight As Long, flag As Boolean

Const HeightRate As Double = 0.8
Const WidthRate As Double = 0.78

Dim ActiveZoom As Long:   ActiveZoom = UfZoom * ((Application.Height * HeightRate) / UfHeight)
If UfName = "UF1" And NowHeight <= 300 Then ActiveZoom = ActiveZoom + 8

Dim ActiveHeight As Long: ActiveHeight = UfHeight * ((Application.Height * HeightRate) / UfHeight)
Dim ActiveWidth As Long:  ActiveWidth = UfWidth * ((Application.Width * WidthRate) / UfWidth)

myForm.Zoom = ActiveZoom 'デフォ

Dim 目隠し As Boolean
'目隠し中は不要
If UF1.UfText.Visible = False And UfName = "UF1" Then
  myForm.Height = ActiveHeight 'デフォ
  myForm.Width = ActiveWidth 'デフォ
  目隠し = True
End If

If UfName = "UF1" Then
  If UF1.厚板 = False And UF1.薄板 = False And UF1.少枚数 = False And UF1.極厚 = False And UF1.スラブ = False Then
    If UF管理.Visible Then '履歴呼び出し中
      myForm.Width = UfWidth * ((Application.Width * WidthRate) / UfWidth)
      InitialSetting = Array(NowZoom, NowHeight, NowWidth, myForm.Zoom, myForm.Height, myForm.Width)
    Else
      myForm.tOP = (Application.Height - myForm.Height) / 2
      myForm.Left = (Application.Width - myForm.Width) / 2
      InitialSetting = Array(NowZoom, NowHeight, NowWidth, myForm.Zoom, myForm.Height, myForm.Width)
      
    End If
    Exit Function
  End If
  
  
  If UF1.VC_L_A.Tag <> "" Or UF1.VC_L_B.Tag <> "" Then 'VCあり
    VcHeight = UfHeight + 145
    flag = True
  End If
  
  If flag Then
    If UF1.Yama.Text <> "" Then
      myForm.Zoom = UfZoom * ((Application.Height * 0.77) / UfHeight)
      myForm.Width = UfWidth * ((Application.Width * 0.72) / UfWidth)
      myForm.Height = Application.Height - 20
      Exit Function
    End If
    
    If (UF1.少枚数 Or UF1.厚板) And UF1.厚板.Caption = "H" Then '少枚数用にサイズが変わる
    myForm.Height = Application.Height - 20
      If n = "リスタート" Then
        ActiveZoom = UfZoom * ((Application.Height * 0.94) / VcHeight)
        myForm.Zoom = ActiveZoom
      Else
        ActiveZoom = UfZoom * ((Application.Height * 0.92) / VcHeight)
        myForm.Zoom = 100
        myForm.Width = NowWidth * ((Application.Width * 0.7) / NowWidth)
      End If
    Else
      ActiveZoom = UfZoom * ((Application.Height * 0.95) / NowHeight)
      myForm.Zoom = ActiveZoom
      myForm.Width = NowWidth * ((Application.Width * 0.72) / NowWidth)
    End If
  Else
    If n = "初期値からスタート" Then
      myForm.Zoom = NowZoom
      myForm.Width = NowWidth * ((Application.Width * 0.75) / NowWidth)
      myForm.Height = NowHeight * ((Application.Height * 0.83) / NowHeight)
      
      If UF1.極厚 Then myForm.Width = UfWidth * ((Application.Width * 0.93) / UfWidth)
    Else
      If 目隠し Then
        myForm.Width = NowWidth * ((Application.Width * 0.74) / NowWidth)
        myForm.Height = ActiveHeight
        
        If UF1.極厚 Then myForm.Width = UfWidth * ((Application.Width * 0.83) / UfWidth)
        
        GoTo SkipA
      End If
      
      myForm.Zoom = NowZoom * ((Application.Height * HeightRate) / NowHeight)
      myForm.Width = NowWidth * ((Application.Width * 0.73) / NowWidth)
      myForm.Height = ActiveHeight
      If UF1.Sumisei.Caption <> "" Then
        myForm.Zoom = NowZoom * ((Application.Height * HeightRate) / NowHeight)
        myForm.Width = NowWidth * ((Application.Width * 0.8) / NowWidth)
        myForm.Height = ActiveHeight
      End If
    End If
  End If
  
  If UF1.極厚 Then myForm.Width = UfWidth * ((Application.Width * 0.83) / UfWidth)
Else
  If UfName = "UFConfi" Then
    myForm.Zoom = NowZoom * ((Application.Height * HeightRate) / NowHeight)
    myForm.Width = UfWidth * ((Application.Width * 0.6) / UfWidth)
    myForm.Height = ActiveHeight
  End If
  If UfName = "UF管理" Then
    myForm.Width = UfWidth * ((Application.Width * 0.89) / UfWidth)
    myForm.Height = UfHeight * ((Application.Height * 0.52) / UfHeight)
  End If
  If UfName = "UF品質基準" Then
    myForm.Width = UfWidth * ((Application.Width * 0.83) / UfWidth)
    myForm.Height = UfHeight * ((Application.Height * 0.75) / UfHeight)
  End If
  
  '2026.05.29
  If UfName = "UFAinput" Then
    Dim ainMc As Long: ainMc = val(UFAinput.Mount_Count.Caption)
    If ainMc >= 7 Then
        ' 動的ラベル表示中はRefreshDynHeightLabelsで設定済みのサイズを維持
    Else
        ' UFAinput自身のUfWidth/UfHeightで計算
        Dim ainZoom  As Long:   ainZoom = 100 * ((Application.Height * 0.8) / 568)
        Dim ainWidth As Single: ainWidth = 596 * ((Application.Width * 0.45) / 596)
        Dim ainHeight As Single: ainHeight = 568 * ((Application.Height * 0.8) / 568)
        myForm.Zoom = ainZoom
        myForm.Width = ainWidth
        myForm.Height = ainHeight
    End If
    If myForm.Width < 595.5 Then myForm.Width = 595.5
    If myForm.Height < 568.5 Then myForm.Height = 568.5
    Call DebugLog("UFAinput確定 Zoom=" & myForm.Zoom & " Width=" & myForm.Width & " Height=" & myForm.Height)
    Call DebugLog("UfZoom=" & UfZoom & " UfWidth=" & UfWidth & " UfHeight=" & UfHeight)
    Call DebugLog("Application.Width=" & Application.Width & " Application.Height=" & Application.Height)
    Call DebugLog("ActiveZoom=" & ActiveZoom & " ActiveHeight=" & ActiveHeight & " ActiveWidth=" & ActiveWidth)
  End If
  
  If UfName = "UF異常値発見" Then
    myForm.Width = ActiveWidth
    myForm.Height = ActiveHeight
  End If
  
  If UfName = "UF_PackNo" Then
    myForm.Zoom = 100
    myForm.Width = 870
    myForm.Height = 660.75
  End If
  
  If UfName = "UFCask" Then
    myForm.Zoom = ActiveZoom * 0.65
    myForm.Width = UfWidth * ((Application.Width * 0.36) / UfWidth)
    myForm.Height = UfHeight * ((Application.Height * 0.35) / UfHeight)
  End If
  If UfName = "UFMitsubishi" Then
    myForm.Zoom = ActiveZoom * 0.65
    myForm.Width = UfWidth * ((Application.Width * 0.31) / UfWidth)
    myForm.Height = UfHeight * ((Application.Height * 0.32) / UfHeight)
  End If
  
End If

SkipA:
Call DebugLog("SkipA時点 " & myForm.Name & " Width=" & myForm.Width & " Height=" & myForm.Height)
myForm.tOP = (Application.Height - myForm.Height) / 2
myForm.Left = (Application.Width - myForm.Width) / 2

InitialSetting = Array(NowZoom, NowHeight, NowWidth, myForm.Zoom, myForm.Height, myForm.Width)
End Function

' フォームをリサイズ可能にするための設定
Public Sub FormSetting()
Dim Result As Long
Dim hwnd As Long
Dim Wnd_STYLE As Long
hwnd = GetActiveWindow()
g_FormHwnd = hwnd '2026.05.29

Wnd_STYLE = GetWindowLong(hwnd, GWL_STYLE)
Wnd_STYLE = Wnd_STYLE Or WS_THICKFRAME Or &H30000

Result = SetWindowLong(hwnd, GWL_STYLE, Wnd_STYLE)
Result = DrawMenuBar(hwnd)
End Sub

Public Sub auto_open()
'ThisWorkbook.Sheets("Main").Select
'UF1.Show vbModeless '_オートメーションエラー対策のため停止
End Sub
Public Sub UF1_Close() '2021.03.15_Application.Run用
If UF1.Visible = True Then UF1.Hide
'Unload UF1
End Sub

Public Sub LoopText(st As String) 'ﾒｲﾝ処理
'On Error GoTo myError
If st = "" Then Exit Sub
UF1.AtButt.Locked = True 'なぜかfocusがここに来る
Call 異常値除去(UF1(st))
UF1.Machin = Machin_Text(UF1.INP)
Call Micro_Juge(st) '機器ｼﾞｬｯｼﾞ ﾏｲｸﾛ判定を先に行う_2020.12.24
Call TextJudg(UF1(st))
Call TimeRecord(st) '2度押しﾁｪｯｸ
UF1.Abno.Tag = ""
Call NumDig '桁数ｾｯﾄ

'2025.1.23
Call CheckLabelsByCategory("S")
Call CheckLabelsByCategory("U")
Call CheckLabelsByCategory("C")
Call CheckLabelsByCategory("D")
Call CheckLabelsByCategory("E")
Call CheckLabelsByCategory("F")

UF1.AtButt.Locked = False
Exit Sub
'myError:
' Call MsgError("ｲﾝﾌﾟｯﾄﾀｰｹﾞｯﾄをして下さい。", Lname, 1)
End Sub
Public Sub TimeRecord(st As String) '測定時間 2度押しﾁｪｯｸ*LoopText
If st = "" Then Exit Sub
If UF1.Abno.Tag = "NG" Then Exit Sub
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim NowTim As Date: NowTim = Format$(Now, "yyyy/mm/dd hh:mm:ss")
If st = "S1" Then UF1.StartTime = Format$(Now, "yyyy/mm/dd hh:mm:ss") 'ｽﾀｰﾄ時間
If WsCheck("Daily" & "_" & Lname) Then '日報情報
 Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Daily" & "_" & Lname)
 Dim LR As Long: LR = UR(sh, 1)
' If Sh.Cells(LR, 1) <> UF1.Lot Then Call Daily_Report(2)   '日報情報(Lot情報)  'LotNoの一致が条件
Call Daily_Report(2) '---------------検査測定のﾀｲﾐﾝｸﾞなので無条件でLot情報書き込み_2021.04.22
' If st = "S1" Then Call Daily_Report(3) '日報情報(検査開始)　＊開始時間を個別に行うと繋がらなくなる為Worker_Changeのみに実行
End If
If (UF1.EachTime <> "") And (UF1.EachValue <> "") Then
    If (UF1(st) = UF1.EachValue) And (DateDiff("s", CDate(UF1.EachTime), CDate(NowTim)) <= 4) Then '値の一致と時間が4秒以内
'        MsgBox "2度押し" & "*" & DateDiff("s", CDate(UF1.EachTime), CDate(NowTim)) & "秒"
        UFWorkerName.TwiceTap.tOP = 36
        UFWorkerName.TwiceTap.Left = 6
        UFWorkerName.TwiceTap.Visible = True
        UFWorkerName.Repaint '描写対策
        Sleep 1000
        UFWorkerName.TwiceTap.Visible = False
        UF1.Repaint '描写対策
        
        UF1.FlagLabel = 0 'フラグを0にしてINP_Changeを抜ける
        UF1(st) = "" 'ｸﾘｱ
        UF1.FlagLabel = 1
        UF1(st).BackColor = &H80000005
        UF1.EachTime = NowTim
        UF1.EachValue = val(Conv(UF1.INP)) '値を残す
        Call Time_Posting(st, NowTim, "") 'WorkerTimeｸﾘｱ
    Else
        UF1.EachTime = NowTim
        UF1.EachValue = val(Conv(UF1.INP)) '値を残す
        Call Time_Posting(st, NowTim) 'WorkerTime上書き
    End If
ElseIf (UF1.EachTime = "") And (UF1.EachValue = "") Then
    UF1.EachTime = NowTim
    UF1.EachValue = val(Conv(UF1.INP))
    Call Time_Posting(st, NowTim)
End If
End Sub
Public Sub NumDig() '桁数ｾｯﾄ
Dim i As Long, a As Long
If UF1.ATU = "" Then Exit Sub
'厚
If val(UF1.ATU) > 74# Then '表示関係：桁ｾｯﾄ UF1("Diag" & Left(UFdiagonal.PosiA, 1) & "Kou") =
  For i = 1 To 4
    UF1("S" & i) = Format$(UF1("S" & i), "0.00")
    UF1("U" & i) = Format$(UF1("U" & i), "0.00")
    UF1("C" & i) = Format$(UF1("C" & i), "0.00")
    UF1("D" & i) = Format$(UF1("D" & i), "0.00")
    UF1("E" & i) = Format$(UF1("E" & i), "0.00")
    UF1("F" & i) = Format$(UF1("F" & i), "0.00")
  Next i
Else
  For i = 1 To 4
    UF1("S" & i) = Format$(UF1("S" & i), "0.000")
    UF1("U" & i) = Format$(UF1("U" & i), "0.000")
    UF1("C" & i) = Format$(UF1("C" & i), "0.000")
    UF1("D" & i) = Format$(UF1("D" & i), "0.000")
    UF1("E" & i) = Format$(UF1("E" & i), "0.000")
    UF1("F" & i) = Format$(UF1("F" & i), "0.000")
  Next i
End If
'幅
'2023.5.12_表示桁数品証確認
If UF1.Youto = "H494" And UF1.Hab < 300# Then '用途コードH494に限り幅測定をノギスで行う
  Select Case val(UF1.Hab) '幅寸法が範囲内の時
  Case 1# To 299#
    For i = 5 To 6
      UF1("S" & i) = Format$(UF1("S" & i), "0.00")
      UF1("U" & i) = Format$(UF1("U" & i), "0.00")
      UF1("C" & i) = Format$(UF1("C" & i), "0.00")
      UF1("D" & i) = Format$(UF1("D" & i), "0.00")
      UF1("E" & i) = Format$(UF1("E" & i), "0.00")
      UF1("F" & i) = Format$(UF1("F" & i), "0.00")
    Next i
  Case Else
    For i = 5 To 6
       UF1("S" & i) = Format$(UF1("S" & i), "0.0")
       UF1("U" & i) = Format$(UF1("U" & i), "0.0")
       UF1("C" & i) = Format$(UF1("C" & i), "0.0")
       UF1("D" & i) = Format$(UF1("D" & i), "0.0")
       UF1("E" & i) = Format$(UF1("E" & i), "0.0")
       UF1("F" & i) = Format$(UF1("F" & i), "0.0")
    Next i
  End Select
Else
  For i = 5 To 6
    UF1("S" & i) = Format$(UF1("S" & i), "0.0")
    UF1("U" & i) = Format$(UF1("U" & i), "0.0")
    UF1("C" & i) = Format$(UF1("C" & i), "0.0")
    UF1("D" & i) = Format$(UF1("D" & i), "0.0")
    UF1("E" & i) = Format$(UF1("E" & i), "0.0")
    UF1("F" & i) = Format$(UF1("F" & i), "0.0")
  Next i
End If

'丈
For i = 7 To 12
   UF1("S" & i) = Format$(UF1("S" & i), "0.0")
   UF1("U" & i) = Format$(UF1("U" & i), "0.0")
   UF1("C" & i) = Format$(UF1("C" & i), "0.0")
   UF1("D" & i) = Format$(UF1("D" & i), "0.0")
   UF1("E" & i) = Format$(UF1("E" & i), "0.0")
Next i
For i = 7 To 10
   UF1("F" & i) = Format$(UF1("F" & i), "0.0")
Next i

'対角
UF1.DiagS = Format$(UF1.DiagS, "0.0")
UF1.DiagU = Format$(UF1.DiagU, "0.0")
UF1.DiagC = Format$(UF1.DiagC, "0.0")
UF1.DiagD = Format$(UF1.DiagD, "0.0")
UF1.DiagE = Format$(UF1.DiagE, "0.0")
UF1.DiagF = Format$(UF1.DiagF, "0.0")
UF1.DiagSKou = Format$(UF1.DiagSKou, "0.0")
UF1.DiagUKou = Format$(UF1.DiagUKou, "0.0")
UF1.DiagCKou = Format$(UF1.DiagCKou, "0.0")
UF1.DiagDKou = Format$(UF1.DiagDKou, "0.0")
UF1.DiagEKou = Format$(UF1.DiagEKou, "0.0")
UF1.DiagFKou = Format$(UF1.DiagFKou, "0.0")

End Sub
Public Function AllCheck(a As Long, n As Long, KenNo As String, Optional Temp As String = "") As Boolean '未測定ﾁｪｯｸ　ﾌﾗｸﾞでﾁｪｯｸヵ所判定_空があったらtrue
If a = 0 Or n = 0 Then Exit Function
Dim i As Long, CountA As Long
Dim th As Long: If UF1.KenCount <> "" Then th = UF1.KenCount
'少枚数用の枚数を確定させる_2023.8.21**枚数の入力がない場合Exit Functionしていたので高さ入力していた場合検査が不可となっていた
If UF1.ATU <> "" Then
  If UF1.KenHBox7 = "" Then
    th = Format(Round(CalculateSheets(UF1, UF1.ATU), 1), "0")
  ElseIf UF1.KenCount <> "" Then
    th = UF1.KenCount
  End If
  If th >= 5 Then th = 5 '少枚数は5枚まで
End If

Select Case Temp
Case "厚板" 'UとDを除く
    For i = a To n
      Select Case i
      Case 5 To 10
        Select Case True
        Case UF1("S" & i) = ""
          AllCheck = True
          Exit For
        Case UF1("C" & i) = ""
          AllCheck = True
          Exit For
        Case UF1("E" & i) = ""
          AllCheck = True
          Exit For
        End Select
      Case Else
        Select Case True
        Case UF1("S" & i) = ""
          AllCheck = True
          Exit For
        Case UF1("U" & i) = ""
          AllCheck = True
          Exit For
        Case UF1("C" & i) = ""
          AllCheck = True
          Exit For
        Case UF1("D" & i) = ""
          AllCheck = True
          Exit For
        Case UF1("E" & i) = ""
          AllCheck = True
          Exit For
        End Select
      End Select
    Next i
Case "少枚数" '2023.6.21_追加
'If UF1.KenCount = "" Then Exit Function'不要停止_2023.8.21
  Select Case th
  Case 5
    For i = a To n
      Select Case True
      Case UF1("S" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("U" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("C" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("D" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("E" & i) = ""
        AllCheck = True
        Exit For
      End Select
    Next i
  Case 4
    For i = a To n
      Select Case i
      Case 5 To 10
        Select Case True
        Case UF1("S" & i) = ""
          AllCheck = True
          Exit For
        Case UF1("C" & i) = ""
          AllCheck = True
          Exit For
        Case UF1("D" & i) = ""
          AllCheck = True
          Exit For
        End Select
      Case Else
        Select Case True
        Case UF1("S" & i) = ""
          AllCheck = True
          Exit For
        Case UF1("U" & i) = ""
          AllCheck = True
          Exit For
        Case UF1("C" & i) = ""
          AllCheck = True
          Exit For
        Case UF1("D" & i) = ""
          AllCheck = True
          Exit For
        End Select
      End Select
    Next i
  Case 3
    For i = a To n
      Select Case True
      Case UF1("S" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("U" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("C" & i) = ""
        AllCheck = True
        Exit For
      End Select
    Next i
  Case 2
    For i = a To n
      Select Case True
      Case UF1("S" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("U" & i) = ""
        AllCheck = True
        Exit For
      End Select
    Next i
  Case 1
    For i = a To n
      Select Case True
      Case UF1("S" & i) = ""
        AllCheck = True
        Exit For
      End Select
    Next i
  End Select
Case "" '操業基準追加***Q-S-QA-ALL-163_検査回数変動***2021.11.29
  If UF1.薄板 Then '------------------------------------------------------------------------
    Dim flag(1 To 6) As Boolean
'    If UF1.InspeButt.Visible Then Flag(1) = True '検入ボタン表示あり(冷延あり 6mm以下)
'    If UF1.Yama.Visible Then Flag(2) = True '山数選択表示あり(冷延あり　6mm以下)
'    If Flag(1) And UF1.InspeButt.Caption = "検入≧26" Then Flag(3) = True '検入25枚以下
'    If UF1.Yama = 3 Then Flag(4) = True    '山数が3
'    If Flag(1) And UF1.InspeButt.Caption = "検入≦25" Then Flag(5) = True '検入26枚以上
    '条件再定義****2022.01.11
'    If UF1.AtButt.Visible And UF1.InspeButt.Visible And UF1.Yama.Visible Then flag(1) = True '冷延あり 6mm以下
'    If UF1.AtButt.Visible And UF1.InspeButt.Visible = False And UF1.Yama.Visible = False Then flag(2) = True '冷延あり 6mm超え
'    If UF1.Yama = 3 Then flag(3) = True      '山数が3
'    If flag(1) And UF1.InspeButt.Caption = "検入≧26" Then flag(4) = True '検入25枚以下
'    If UF1.AtButt.Visible = False And UF1.InspeButt.Visible = False And UF1.Yama.Visible = False Then flag(5) = True '冷延なし
    
'2023.4.17****検番による分岐に変更_定義のコース部分は不要
'2023.4.17_検査No検索

'条件定義****2022.4.18
If KenNo = "L" Or KenNo = "N" Or KenNo = "K" Or KenNo = "T" Or KenNo = "S" Or KenNo = "F" Then flag(1) = True
If KenNo = "H" Then flag(2) = True
If UF1.Yama = 3 Then flag(3) = True  '山数が3*********山数3を選択時県有枚数選択は消える為flag1は条件に入れることができない
If flag(1) And UF1.InspeButt.Caption = "検入≧26" Then flag(4) = True '検入25枚以下
If KenNo = "" Then flag(5) = True '検査番号がなかった場合
If UF1.Yama = 4 Then flag(6) = True  '山数が4


  If flag(6) Then '6回
    For i = a To n
    Select Case True
      Case UF1("S" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("U" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("C" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("D" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("E" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("F" & i) = ""
        AllCheck = True
        Exit For
      End Select
    Next i
  ElseIf flag(3) Or flag(5) Then '5回'冷延なしの薄板　or　冷延あり　かつ　6mm超え　or　冷延あり　かつ　山数が3＿flag(5) Or flag(2) Or flag(3)
    For i = a To n
    Select Case True
      Case UF1("S" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("U" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("C" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("D" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("E" & i) = ""
        AllCheck = True
        Exit For
      End Select
    Next i
  ElseIf flag(4) Then '2回'冷延あり　かつ　検入25枚以下
    For i = a To n
    Select Case True
      Case UF1("S" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("E" & i) = ""
        AllCheck = True
        Exit For
      End Select
    Next i
  ElseIf flag(1) Then '3回'冷延あり 6mm以下
    For i = a To n
    Select Case True
      Case UF1("S" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("C" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("E" & i) = ""
        AllCheck = True
        Exit For
      End Select
    Next i
  End If
ElseIf UF1.少枚数 Or UF1.極厚 Then '極追加
  For i = a To n
  Select Case True
    Case UF1("S" & i) = ""
      AllCheck = True
      Exit For
    Case UF1("U" & i) = ""
      AllCheck = True
      Exit For
    Case UF1("C" & i) = ""
      AllCheck = True
      Exit For
    Case UF1("D" & i) = ""
      AllCheck = True
      Exit For
    Case UF1("E" & i) = ""
      AllCheck = True
      Exit For
    End Select
  Next i
ElseIf UF1.スラブ Then
  For i = a To n
    Select Case i
    Case 1 To 4
      Select Case True
      Case UF1("S" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("C" & i) = ""
        AllCheck = True
        Exit For
      End Select
    Case 5 To 10
      Select Case True
      Case UF1("S" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("U" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("C" & i) = ""
        AllCheck = True
        Exit For
      Case UF1("D" & i) = ""
        AllCheck = True
        Exit For
      End Select
    End Select
  Next i
End If '------------------------------------------------------------------------
End Select
End Function
Public Sub PucketMove(Pucket As String, KenNo As String, Optional Youto As String = "")  '伝送機器判定用

If UF1.厚板 = False And UF1.薄板 = False And UF1.少枚数 = False And UF1.極厚 = False And UF1.スラブ = False Then Exit Sub

Dim a As Long, b As Long, i As Long, n As Long, Temp As String: Temp = ""
'2022.01.11_条件定義
Dim flag(1 To 6) As Boolean
'If UF1.AtButt.Visible And UF1.InspeButt.Visible And UF1.Yama.Visible Then flag(1) = True '冷延あり 6mm以下
'If UF1.AtButt.Visible And UF1.InspeButt.Visible = False And UF1.Yama.Visible = False Then flag(2) = True '冷延あり 6mm超え
'If UF1.Yama = 3 Then flag(3) = True      '山数が3
'If flag(1) And UF1.InspeButt.Caption = "検入≧26" Then flag(4) = True '検入25枚以下
'If UF1.AtButt.Visible = False And UF1.InspeButt.Visible = False And UF1.Yama.Visible = False Then flag(5) = True '冷延なし

'2023.4.17****検番による分岐に変更_定義のコース部分は不要
'2023.4.17_検査No検索

'-----------------------------------------------------------------------------
'Q-S-QA-MSR-007-R04にて用途コードH494に限り幅測定をノギスで行う_2023.5.10
'-----------------------------------------------------------------------------

If UF1.厚板 = False And UF1.薄板 = False And UF1.少枚数 = False And UF1.極厚 = False And UF1.スラブ = False Then Exit Sub

'条件定義****2022.4.18
If KenNo = "L" Or KenNo = "N" Or KenNo = "K" Or KenNo = "T" Or KenNo = "S" Or KenNo = "F" Then flag(1) = True
If KenNo = "H" Then flag(2) = True
If UF1.Yama = 3 Then flag(3) = True  '山数が3*********山数3を選択時県有枚数選択は消える為flag1は条件に入れることができない
If flag(1) And UF1.InspeButt.Caption = "検入≧26" Then flag(4) = True '検入25枚以下
If KenNo = "" Then flag(5) = True '検査番号がなかった場合
If UF1.Yama = 4 Then flag(6) = True  '山数が4

Select Case Pucket '移動先選定
Case "DT" 'DT10002+0000021.712M
  Select Case UF1.AtButt.Caption
  Case "厚2ヵ所" 'ボタン表示が2回の時は4回検査の状態
    a = 1
    n = 4
  Case "厚4ヵ所" 'ボタン表示が4回の時は2回検査の状態
    a = 1
    n = 2
  End Select
  
  '2023.5.10_用途コードH494に限り幅測定をノギスで行う
  If Youto = "H494" And UF1.Hab < 300# Then
    a = 1
    n = 6
  End If
Case "#L"  '＃L395.8
  If UF1.極厚 Then
    a = 5
    n = 14 '2022.06_丈3,4分追加+2
  Else
    a = 5
    n = 10
  End If
  
  '2023.5.10_用途コードH494に限り幅測定をノギスで行うためコンベックスで測定する箇所は丈のみとなるためずれる
  If Youto = "H494" And UF1.Hab < 300# Then
    a = 7
    n = 10
  End If
Case "#〆" '手入力_2021.02.23 @を#に変更
  If UF1.極厚 Then '巾丈に移行
   a = 1
   n = 14
  Else
   a = 1
   n = 10
  End If
Case Else
MsgBox "ﾊﾟｹｯﾄ無し:Exit Sub" & "_PucketMove", vbExclamation
Exit Sub
End Select

If UF1.厚板 Then Temp = "厚板" 'AllCheck用ﾌﾗｸﾞ
'2023.6.21_少枚数用
Dim Kno As String: Kno = UF1.厚板.Caption
Dim Sflag As Boolean: If UF1.少枚数 Then Sflag = True
If Kno = "H" And Sflag Then Temp = "少枚数"

Dim CheckFlag As Boolean: CheckFlag = AllCheck(a, n, KenNo, Temp)

If CheckFlag = False Then '全部埋まっている状態
  MsgBox Pucket & "検査済", vbExclamation
  UF1.Auto = "off"
  UF1.INP = ""
  UF1.INP.Tag = ""
  UF1.FlagLabel = 1
  Exit Sub
End If

For i = a To n
  If UF1("S" & i) = "" Then '----------------------------------------------S
    UF1.INP.Tag = "S" & i 'UF1.INP.Tag
    UF1.Tagtext.Caption = UF1.INP.Tag '可視化
    Exit Sub
  End If
Next i
If UF1.厚板 Then '----------------------------------------------U
  For i = a To n
    Select Case i
    Case 1 To 4
      If UF1("U" & i) = "" Then
        UF1.INP.Tag = "U" & i
        UF1.Tagtext.Caption = UF1.INP.Tag
        Exit Sub
      End If
    Case 5 To 10
      For b = 5 To 10
        If UF1("C" & b) = "" Then
          UF1.INP.Tag = "C" & b
          UF1.Tagtext.Caption = UF1.INP.Tag
          Exit Sub
        End If
      Next b
    End Select
  Next i
ElseIf UF1.スラブ Then '2回_S,C
  For i = a To n
    Select Case i
    Case 1 To 4
      For b = 5 To 8
        If UF1("U" & b) = "" Then
          UF1.INP.Tag = "U" & b
          UF1.Tagtext.Caption = UF1.INP.Tag
          Exit Sub
        End If
      Next b
    Case 5 To 8
      If UF1("U" & i) = "" Then
        UF1.INP.Tag = "U" & i
        UF1.Tagtext.Caption = UF1.INP.Tag
        Exit Sub
      End If
    End Select
  Next i
ElseIf UF1.薄板 Then
  If flag(2) Or flag(5) Or flag(3) Or flag(6) Then '5回
    For i = a To n
      If UF1("U" & i) = "" Then
        UF1.INP.Tag = "U" & i 'UF1.INP.Tag
        UF1.Tagtext.Caption = UF1.INP.Tag
        Exit Sub
      End If
    Next i
  ElseIf flag(4) Then '2回
    For i = a To n
      Select Case i
      Case 1 To 10
        If UF1("E" & i) = "" Then
          UF1.INP.Tag = "E" & i
          UF1.Tagtext.Caption = UF1.INP.Tag
          Exit Sub
        End If
      End Select
    Next i
  ElseIf flag(1) Then  'L4HI1,2山,26枚以上_3回
    For i = a To n
      Select Case i
      Case 1 To 10
        If UF1("C" & i) = "" Then
          UF1.INP.Tag = "C" & i
          UF1.Tagtext.Caption = UF1.INP.Tag
          Exit Sub
        End If
      End Select
    Next i
  End If
Else
  For i = a To n
    If UF1("U" & i) = "" And UF1("U" & i).Visible Then '2023.6.21_少枚数用
      UF1.INP.Tag = "U" & i 'UF1.INP.Tag
      UF1.Tagtext.Caption = UF1.INP.Tag
      Exit Sub
    End If
  Next i
End If
If Not UF1.薄板 Then '----------------------------------------------C
  For i = a To n
    If UF1("C" & i) = "" And UF1("C" & i).Visible Then
      UF1.INP.Tag = "C" & i
      UF1.Tagtext.Caption = UF1.INP.Tag
      Exit Sub
    End If
  Next i
'(UF1.薄板 And flag(1)) Or (UF1.薄板 And flag(2)) Or (UF1.薄板 And flag(3)) Or (UF1.薄板 And flag(5))
ElseIf flag(1) Or flag(2) Or flag(3) Or flag(5) Or flag(6) Then  '(冷延あり　かつ6mm超え)　　か　(冷延なし　かつ　8mm未満)**2022.01.05
  For i = a To n
    If UF1("C" & i) = "" Then
      UF1.INP.Tag = "C" & i
      UF1.Tagtext.Caption = UF1.INP.Tag
      Exit Sub
    End If
  Next i
ElseIf UF1.薄板 And flag(4) Then  '記述忘れ_2021.12.23
    For i = a To n
      Select Case i
      Case 1 To 10
        If UF1("E" & i) = "" Then
          UF1.INP.Tag = "E" & i
          UF1.Tagtext.Caption = UF1.INP.Tag
          Exit Sub
        End If
      End Select
    Next i
End If
If UF1.厚板 Then '----------------------------------------------D
  For i = a To n
    Select Case i
    Case 1 To 4
      If UF1("D" & i) = "" Then
        UF1.INP.Tag = "D" & i
        UF1.Tagtext.Caption = UF1.INP.Tag
        Exit Sub
      End If
    Case 5 To 10
      For b = 5 To 10
        If UF1("E" & b) = "" Then
          UF1.INP.Tag = "E" & b
          UF1.Tagtext.Caption = UF1.INP.Tag
          Exit Sub
        End If
      Next b
    End Select
  Next i
ElseIf UF1.スラブ Then
  For i = a To n
    Select Case i
    Case 1 To 4
      For b = 5 To 8
        If UF1("D" & b) = "" Then
          UF1.INP.Tag = "D" & b
          UF1.Tagtext.Caption = UF1.INP.Tag
          Exit Sub
        End If
      Next b
    Case 5 To 8
      If UF1("D" & i) = "" Then
        UF1.INP.Tag = "D" & i
        UF1.Tagtext.Caption = UF1.INP.Tag
        Exit Sub
      End If
    End Select
  Next i
ElseIf UF1.薄板 Then
  If flag(2) Or flag(5) Or flag(3) Or flag(6) Then
    For i = a To n
      If UF1("D" & i) = "" Then
        UF1.INP.Tag = "D" & i 'UF1.INP.Tag
        UF1.Tagtext.Caption = UF1.INP.Tag
        Exit Sub
      End If
    Next i
  ElseIf flag(4) Then
    For i = a To n
      Select Case i
      Case 1 To 10
        If UF1("E" & i) = "" Then
          UF1.INP.Tag = "E" & i
          UF1.Tagtext.Caption = UF1.INP.Tag
          Exit Sub
        End If
      End Select
    Next i
  ElseIf flag(1) Then 'L4HI1,2山,26枚以上
    For i = a To n
      Select Case i
      Case 1 To 10
        If UF1("E" & i) = "" Then
          UF1.INP.Tag = "E" & i
          UF1.Tagtext.Caption = UF1.INP.Tag
          Exit Sub
        End If
      End Select
    Next i
  End If
Else
  For i = a To n
    If UF1("D" & i) = "" And UF1("D" & i).Visible Then '2023.6.21_少枚数用
      UF1.INP.Tag = "D" & i 'UF1.INP.Tag
      UF1.Tagtext.Caption = UF1.INP.Tag
      Exit Sub
    End If
  Next i
End If
For i = a To n
  If UF1("E" & i) = "" And UF1("E" & i).Visible Then  '----------------------------------------------E
    UF1.INP.Tag = "E" & i
    UF1.Tagtext.Caption = UF1.INP.Tag
    Exit Sub
  End If
Next i
                                                      
If flag(6) Then
For i = a To n
  If UF1("F" & i) = "" And UF1("F" & i).Visible Then '----------------------------------------------F
    UF1.INP.Tag = "F" & i
    UF1.Tagtext.Caption = UF1.INP.Tag
    Exit Sub
  End If
Next i
End If
End Sub
Public Sub Look_S_Blank(a As Long, n As Long, KenNo As String)  '空白探し_薄板用*2021.04.12****2023.4.20_引数追加
'If UF1.Yama.Text = "" Then Exit Sub
Dim flag(1 To 6) As Boolean
'条件定義****2022.01.11
'If UF1.AtButt.Visible And UF1.InspeButt.Visible And UF1.Yama.Visible Then flag(1) = True '冷延あり 6mm以下
'If UF1.AtButt.Visible And UF1.InspeButt.Visible = False And UF1.Yama.Visible = False Then flag(2) = True '冷延あり 6mm超え
'If UF1.Yama = 3 Then flag(3) = True      '山数が3
'If flag(1) And UF1.InspeButt.Caption = "検入≧26" Then flag(4) = True '検入25枚以下
'If UF1.AtButt.Visible = False And UF1.InspeButt.Visible = False And UF1.Yama.Visible = False Then flag(5) = True '冷延なし

Dim i As Long, b As Long
If Not UF1.Auto = "on" Then Exit Sub '直接選択の場合抜ける
Dim AFl As Boolean: AFl = False

'2023.4.17****検番による分岐に変更_定義のコース部分は不要
'2023.4.17_検査No検索
'Dim SIKAHIKAcc: SIKAHIKAcc = Arrbox("SIKAHIKINOW.accdb", "", "Fields", UF1.DBflag.text) 'SIKAHIKINOW.accdbのフィールドに[ｵｰﾀﾞｰ板丈]がないのでエラーになる
'Dim No As Long
'If Not IsEmpty(SIKAHIKAcc) Then
'  For i = 1 To UBound(SIKAHIKAcc, 2)
'    Select Case SIKAHIKAcc(1, i)
'    Case "検査番号": No = i
'    End Select
'  Next i
'End If
'Dim KenAcc: KenAcc = Arrbox("SIKAHIKINOW.accdb", UF1.LOT, "Record", UF1.DBflag.text)
'
'If IsEmpty(KenAcc) Then 'SIKALOTNOWでやりなおし
'    SIKAHIKAcc = Arrbox("SIKALOTNOW.accdb", "", "Fields", UF1.DBflag.text)
'    If Not IsEmpty(SIKAHIKAcc) Then
'      For i = 1 To UBound(SIKAHIKAcc, 2)
'        Select Case SIKAHIKAcc(1, i)
'        Case "検査番号": No = i
'        End Select
'      Next i
'    End If
'    KenAcc = Arrbox("SIKALOTNOW.accdb", UF1.LOT, "Record", UF1.DBflag.text)
'End If
'
'Dim KENNO As String
'If IsEmpty(KenAcc) = False Then KENNO = Left(KenAcc(1, No), 1)  '検査No'頭1桁

'条件定義****2022.4.18
If KenNo = "L" Or KenNo = "N" Or KenNo = "K" Or KenNo = "T" Or KenNo = "S" Or KenNo = "F" Or KenNo = "" Then flag(1) = True 'KenNo = ""検番拾えなかったとき用追加_2023.6.27
If KenNo = "H" Then flag(2) = True
If UF1.Yama = 3 Then flag(3) = True  '山数が3*********山数3を選択時県有枚数選択は消える為flag1は条件に入れることができない
If flag(1) And UF1.InspeButt.Caption = "検入≧26" Then flag(4) = True '検入25枚以下
If UF1.Yama = 4 Then flag(6) = True  '山数が4

Select Case True
Case UF1.薄板 And UF1.AtButt.Caption = "厚4ヵ所"   'ボタン表示が4回の時は2回検査の状態
'----------------------------------------------S
  For i = a To n
    Select Case i
    Case 1, 2, 5, 6, 7, 8, 9, 10
      If UF1("S" & i) = "" Then: AFl = True
    Case 3
      i = i + 2
      If UF1("S" & i) = "" Then: AFl = True
    Case 4
      i = i + 1
      If UF1("S" & i) = "" Then: AFl = True
    End Select
    If AFl Then
      UF1.INP.Tag = "S" & i 'UF1.INP.Tag
      Call LookBlank_Color(UF1.INP.Tag)
      Exit Sub
    End If
  Next i
If flag(4) Then '2022.3.23_追加************先に実行必須_2回検査
  For i = a To n
    Select Case i
    Case 1, 2, 5, 6, 7, 8, 9, 10
      If UF1("E" & i) = "" Then: AFl = True
    Case 3
      i = i + 2
      If UF1("E" & i) = "" Then: AFl = True
    Case 4
      i = i + 1
      If UF1("E" & i) = "" Then: AFl = True
    End Select
    If AFl Then
      UF1.INP.Tag = "E" & i 'UF1.INP.Tag
      Call LookBlank_Color(UF1.INP.Tag)
      Exit Sub
    End If
  Next i
End If
'----------------------------------------------U_操業基準追加_UD検査削除**Q-S-QA-ALL-163**2021.11.29
  If flag(2) Or flag(3) Or flag(5) Or flag(6) Then '2021.11.30_4HI材以外の場合 山数3の時
    For i = a To n
      Select Case i
      Case 1, 2, 5, 6, 7, 8, 9, 10
        If UF1("U" & i) = "" And UF1("U" & i).Visible Then: AFl = True
      Case 3
        i = i + 2
        If UF1("U" & i) = "" And UF1("U" & i).Visible Then: AFl = True
      Case 4
        i = i + 1
        If UF1("U" & i) = "" And UF1("U" & i).Visible Then: AFl = True
      End Select
      If AFl Then
        UF1.INP.Tag = "U" & i 'UF1.INP.Tag
        Call LookBlank_Color(UF1.INP.Tag)
        Exit Sub
      End If
    Next i
  End If
'----------------------------------------------C*****冷延ありで6mm超えの場合追加(25，26表示なし、山表示なし)
  If flag(1) Or flag(2) Or flag(3) Or flag(5) Or flag(6) Then '2021.12.23**2022.01.05_追加 ''flag(1) Or flag(2) Or flag(3) Or flag(5)
    For i = a To n
      Select Case i
      Case 1, 2, 5, 6, 7, 8, 9, 10
        If UF1("C" & i) = "" And UF1("C" & i).Visible Then: AFl = True
      Case 3
        i = i + 2
        If UF1("C" & i) = "" And UF1("C" & i).Visible Then: AFl = True
      Case 4
        i = i + 1
        If UF1("C" & i) = "" And UF1("C" & i).Visible Then: AFl = True
      End Select
      If AFl Then
        UF1.INP.Tag = "C" & i 'UF1.INP.Tag
        Call LookBlank_Color(UF1.INP.Tag)
        Exit Sub
      End If
    Next i
  End If
'----------------------------------------------D
  If flag(2) Or flag(3) Or flag(5) Or flag(6) Then
    For i = a To n
      Select Case i
      Case 1, 2, 5, 6, 7, 8, 9, 10
        If UF1("D" & i) = "" And UF1("D" & i).Visible Then: AFl = True
      Case 3
        i = i + 2
        If UF1("D" & i) = "" And UF1("D" & i).Visible Then: AFl = True
      Case 4
        i = i + 1
        If UF1("D" & i) = "" And UF1("D" & i).Visible Then: AFl = True
      End Select
      If AFl Then
        UF1.INP.Tag = "D" & i 'UF1.INP.Tag
        Call LookBlank_Color(UF1.INP.Tag)
        Exit Sub
      End If
    Next i
  End If
'----------------------------------------------E
  For i = a To n
    Select Case i
    Case 1, 2, 5, 6, 7, 8, 9, 10
      If UF1("E" & i) = "" And UF1("E" & i).Visible Then: AFl = True
    Case 3
      i = i + 2
      If UF1("E" & i) = "" And UF1("E" & i).Visible Then: AFl = True
    Case 4
      i = i + 1
      If UF1("E" & i) = "" And UF1("E" & i).Visible Then: AFl = True
    End Select
    If AFl Then
      UF1.INP.Tag = "E" & i 'UF1.INP.Tag
      Call LookBlank_Color(UF1.INP.Tag)
      Exit Sub
    End If
  Next i
'----------------------------------------------F
  If flag(6) Then
    For i = a To n
      Select Case i
      Case 1, 2, 5, 6, 7, 8, 9, 10
        If UF1("F" & i) = "" And UF1("F" & i).Visible Then: AFl = True
      Case 3
        i = i + 2
        If UF1("F" & i) = "" And UF1("F" & i).Visible Then: AFl = True
      Case 4
        i = i + 1
        If UF1("F" & i) = "" And UF1("F" & i).Visible Then: AFl = True
      End Select
      If AFl Then
        UF1.INP.Tag = "F" & i 'UF1.INP.Tag
        Call LookBlank_Color(UF1.INP.Tag)
        Exit Sub
      End If
    Next i
  End If
Case UF1.薄板 And UF1.AtButt.Caption = "厚2ヵ所"   'ボタン表示が2回の時は4回検査の状態
  '----------------------------------------------S
  For i = a To n
    If UF1("S" & i) = "" Then
      UF1.INP.Tag = "S" & i 'UF1.INP.Tag
      Call LookBlank_Color(UF1.INP.Tag)
      Exit Sub
    End If
  Next i
  If flag(4) Then '2022.3.23_追加************先に実行必須_2回検査
    For i = a To n
      If UF1("E" & i) = "" Then
        UF1.INP.Tag = "E" & i
        Call LookBlank_Color(UF1.INP.Tag)
        Exit Sub
      End If
    Next i
  End If
  '----------------------------------------------U
  If flag(2) Or flag(3) Or flag(5) Or flag(6) Then '2021.11.30_4HI材以外の場合
    For i = a To n
      If UF1("U" & i) = "" Then
        UF1.INP.Tag = "U" & i 'UF1.INP.Tag
        Call LookBlank_Color(UF1.INP.Tag)
        Exit Sub
      End If
    Next i
  End If
  '----------------------------------------------C**冷延ありで6mm超えの場合追加(25，26表示なし、山表示なし)
  If flag(1) Or flag(2) Or flag(3) Or flag(5) Or flag(6) Then '2021.12.23**2022.01.05_追加
    For i = a To n
      If UF1("C" & i) = "" Then
        UF1.INP.Tag = "C" & i
        Call LookBlank_Color(UF1.INP.Tag)
        Exit Sub
      End If
    Next i
  End If
  '----------------------------------------------D
  If flag(2) Or flag(3) Or flag(5) Or flag(6) Then
    For i = a To n
      If UF1("D" & i) = "" Then
        UF1.INP.Tag = "D" & i 'UF1.INP.Tag
        Call LookBlank_Color(UF1.INP.Tag)
        Exit Sub
      End If
    Next i
  End If
   '----------------------------------------------E
  For i = a To n
    If UF1("E" & i) = "" Then
      UF1.INP.Tag = "E" & i
      Call LookBlank_Color(UF1.INP.Tag)
      Exit Sub
    End If
  Next i
End Select
   '----------------------------------------------F
If flag(6) Then
  For i = a To n
    If UF1("F" & i) = "" Then
      UF1.INP.Tag = "F" & i
      Call LookBlank_Color(UF1.INP.Tag)
      Exit Sub
    End If
  Next i
End If
End Sub
Public Sub LookBlank_Color(a As String) '*Look_S_Blankで使用
UF1.Tagtext.Caption = a '可視化
'フォーム上部に
If UF1.MPEX.Tag = "" Then '2024.08.20 RGB(255,255,255)
    UF1(a).BorderColor = RGB(255, 0, 0)
    UF1(a).BackColor = RGB(255, 255, 0)
Else
    UF1(a).BorderColor = RGB(0, 0, 0)
    UF1(a).BackColor = RGB(255, 255, 255)
End If
End Sub
Public Sub Smallnumber_Blank(a As Long, n As Long, xa As Long) '少枚数用空白探し  UF1.INP.Tagに座標**********2023.6.21
Dim i As Long, b As Long, x As Long, p As String
Dim AFl As Boolean: AFl = False
Dim Sfl As Boolean: Sfl = False
If Not UF1.Auto = "on" Then Exit Sub '直接選択の場合抜ける

Select Case xa
Case 5, 4, 3, 2, 1
  For i = a To n
    If UF1("S" & i) = "" Then '----------------------------------------------S
      UF1.INP.Tag = "S" & i 'UF1.INP.Tag
      UF1.Tagtext.Caption = UF1.INP.Tag '可視化
      Call LookBlank_Color(UF1.INP.Tag)
      Exit Sub
    End If
  Next i
  For i = a To n
    If UF1("U" & i) = "" And UF1("U" & i).Visible Then  '----------------------------------------------U
      UF1.INP.Tag = "U" & i
      UF1.Tagtext.Caption = UF1.INP.Tag
      Call LookBlank_Color(UF1.INP.Tag)
      Exit Sub
    End If
  Next i
  For i = a To n
    If UF1("C" & i) = "" And UF1("C" & i).Visible Then  '----------------------------------------------C
      UF1.INP.Tag = "C" & i
      UF1.Tagtext.Caption = UF1.INP.Tag
      Call LookBlank_Color(UF1.INP.Tag)
      Exit Sub
    End If
  Next i
  For i = a To n
    If UF1("D" & i) = "" And UF1("D" & i).Visible Then '----------------------------------------------D
      UF1.INP.Tag = "D" & i
      UF1.Tagtext.Caption = UF1.INP.Tag
      Call LookBlank_Color(UF1.INP.Tag)
      Exit Sub
    End If
  Next i
  For i = a To n
    If UF1("E" & i) = "" And UF1("E" & i).Visible Then  '----------------------------------------------E
      UF1.INP.Tag = "E" & i
      UF1.Tagtext.Caption = UF1.INP.Tag
      Call LookBlank_Color(UF1.INP.Tag)
      Exit Sub
    End If
  Next i
  For i = a To n
    If UF1("F" & i) = "" And UF1("F" & i).Visible Then  '----------------------------------------------F
      UF1.INP.Tag = "F" & i
      UF1.Tagtext.Caption = UF1.INP.Tag
      Call LookBlank_Color(UF1.INP.Tag)
      Exit Sub
    End If
  Next i
End Select
End Sub
Public Sub Look_Blank(a As Long, n As Long)  '空白探し  UF1.INP.Tagに座標
Dim i As Long, b As Long, x As Long, p As String
Dim AFl As Boolean: AFl = False
Dim Sfl As Boolean: Sfl = False

If Not UF1.Auto = "on" Then Exit Sub '直接選択の場合抜ける
For i = a To n
  If UF1("S" & i) = "" Then '----------------------------------------------S
    UF1.INP.Tag = "S" & i 'UF1.INP.Tag
    UF1.Tagtext.Caption = UF1.INP.Tag '可視化
    UF1("S" & i).BorderColor = RGB(255, 0, 0)
    UF1("S" & i).BackColor = RGB(255, 255, 0)
    Exit Sub
  End If
Next i
If UF1.厚板 Then '----------------------------------------------U
  For i = a To n
    Select Case i
    Case 1 To 4
      If UF1("U" & i) = "" Then
        UF1.INP.Tag = "U" & i
        UF1.Tagtext.Caption = UF1.INP.Tag
        UF1("U" & i).BorderColor = RGB(255, 0, 0)
        UF1("U" & i).BackColor = RGB(255, 255, 0)
        Exit Sub
      End If
    Case 5 To 10
      For b = 1 To 10
        If UF1("C" & b) = "" Then
          UF1.INP.Tag = "C" & b
          UF1.Tagtext.Caption = UF1.INP.Tag
          UF1("C" & b).BorderColor = RGB(255, 0, 0)
          UF1("C" & b).BackColor = RGB(255, 255, 0)
          Exit Sub
        End If
      Next b
    End Select
  Next i
ElseIf UF1.スラブ Then
  For i = a To n
    Select Case i
    Case 1 To 4
      For b = 5 To 8
        If UF1("U" & b) = "" Then
          UF1.INP.Tag = "U" & b
          UF1.Tagtext.Caption = UF1.INP.Tag
          UF1("U" & b).BorderColor = RGB(255, 0, 0)
          UF1("U" & b).BackColor = RGB(255, 255, 0)
          Exit Sub
        End If
      Next b
    Case 5 To 8
      If UF1("U" & i) = "" Then
        UF1.INP.Tag = "U" & i
        UF1.Tagtext.Caption = UF1.INP.Tag
        UF1("U" & i).BorderColor = RGB(255, 0, 0)
        UF1("U" & i).BackColor = RGB(255, 255, 0)
        Exit Sub
      End If
    End Select
  Next i
Else
  For i = a To n
    If UF1("U" & i) = "" Then
      UF1.INP.Tag = "U" & i 'UF1.INP.Tag
      UF1.Tagtext.Caption = UF1.INP.Tag
      UF1("U" & i).BorderColor = RGB(255, 0, 0)
      UF1("U" & i).BackColor = RGB(255, 255, 0)
      Exit Sub
    End If
  Next i
End If
For i = a To n
  If UF1("C" & i) = "" Then '----------------------------------------------C
    UF1.INP.Tag = "C" & i
    UF1.Tagtext.Caption = UF1.INP.Tag
    UF1("C" & i).BorderColor = RGB(255, 0, 0)
    UF1("C" & i).BackColor = RGB(255, 255, 0)
    Exit Sub
  End If
Next i
If UF1.厚板 Then '----------------------------------------------D
  For i = a To n
    Select Case i
    Case 1 To 4
      If UF1("D" & i) = "" Then
        UF1.INP.Tag = "D" & i
        UF1.Tagtext.Caption = UF1.INP.Tag
        UF1("D" & i).BorderColor = RGB(255, 0, 0)
        UF1("D" & i).BackColor = RGB(255, 255, 0)
        Exit Sub
      End If
    Case 5 To 10
      For b = 1 To 10
        If UF1("E" & b) = "" Then
          UF1.INP.Tag = "E" & b
          UF1.Tagtext.Caption = UF1.INP.Tag
          UF1("E" & b).BorderColor = RGB(255, 0, 0)
          UF1("E" & b).BackColor = RGB(255, 255, 0)
          Exit Sub
        End If
      Next b
    End Select
  Next i
ElseIf UF1.スラブ Then
  For i = a To n
    Select Case i
    Case 1 To 4
      For b = 5 To 8
        If UF1("D" & b) = "" Then
          UF1.INP.Tag = "D" & b
          UF1.Tagtext.Caption = UF1.INP.Tag
          UF1("D" & b).BorderColor = RGB(255, 0, 0)
          UF1("D" & b).BackColor = RGB(255, 255, 0)
          Exit Sub
        End If
      Next b
    Case 5 To 8
      If UF1("D" & i) = "" Then
        UF1.INP.Tag = "D" & i
        UF1.Tagtext.Caption = UF1.INP.Tag
        UF1("D" & i).BorderColor = RGB(255, 0, 0)
        UF1("D" & i).BackColor = RGB(255, 255, 0)
        Exit Sub
      End If
    End Select
  Next i
Else
  For i = a To n
    If UF1("D" & i) = "" Then
      UF1.INP.Tag = "D" & i 'UF1.INP.Tag
      UF1.Tagtext.Caption = UF1.INP.Tag
      UF1("D" & i).BorderColor = RGB(255, 0, 0)
      UF1("D" & i).BackColor = RGB(255, 255, 0)
      Exit Sub
    End If
  Next i
End If

For i = a To n
  If UF1("E" & i) = "" Then '----------------------------------------------E
    UF1.INP.Tag = "E" & i
    UF1.Tagtext.Caption = UF1.INP.Tag
    UF1("E" & i).BorderColor = RGB(255, 0, 0)
    UF1("E" & i).BackColor = RGB(255, 255, 0)
    Exit Sub
  End If
Next i

If UF1.Yama = 4 Then
  For i = a To n
    If UF1("F" & i) = "" Then '----------------------------------------------F
      UF1.INP.Tag = "F" & i
      UF1.Tagtext.Caption = UF1.INP.Tag
      UF1("F" & i).BorderColor = RGB(255, 0, 0)
      UF1("F" & i).BackColor = RGB(255, 255, 0)
      Exit Sub
    End If
  Next i
End If
End Sub
Public Function ObjMoveText(TB As MSForms.Label) As String 'ﾗﾍﾞﾙｸﾘｯｸでの処理　ｸﾗｽ_2022.04.12
ObjMoveText = TB.Name
End Function
Public Function TagSet(TB As MSForms.TextBox) As String 'ｸﾘｯｸでの処理
TagSet = TB.Name
UF1.INP.Tag = TagSet
UF1.Tagtext.Caption = UF1.INP.Tag
End Function
Public Function TextTag(TB As MSForms.Label) As String 'ﾗﾍﾞﾙｸﾘｯｸでの処理　ｸﾗｽ
UF1.Auto = "off" '空白探し停止
UF1.Tag = "STOP" '音楽ｽﾄｯﾌﾟ

'-----2021.03.04_Focus対策
UF1.Worker.Enabled = True
UF1.INP.Enabled = True
UF1.Manual.Enabled = True
UF1.StartTB.Enabled = True

Call LoopJudg
Call HizLoopJudg

TextTag = TB.Name
UF1(TB.Name).BorderColor = RGB(255, 0, 0)
UF1(TB.Name).BackColor = RGB(255, 255, 0)
UF1.INP.Tag = TextTag
UF1.Tagtext.Caption = UF1.INP.Tag
UF1.Tag = "" '音楽ｽﾀｰﾄ
End Function
Public Sub DiagonalErase(TB As MSForms.Label) '対角差ののﾘｾｯﾄ　ｸﾗｽ
If Not UF1.極厚 Then
  Select Case Right((TB.Name), 1)
  Case 9
'    UF1(Left(TB.Name, 1) & 10) = ""
'    UF1("Diag" & Left(TB.Name, 1) & "Kou") = ""
    UF1("Diag" & Left(TB.Name, 1)) = ""
    UF1(Left(TB.Name, 1) & 10).BackColor = &H80000005
    UF1(Left(TB.Name, 1) & 9).BackColor = &H80000005
  Case 0
'    UF1(Left(TB.Name, 1) & 9) = ""
'    UF1("Diag" & Left(TB.Name, 1) & "Kou") = ""
    UF1("Diag" & Left(TB.Name, 1)) = ""
    UF1(Left(TB.Name, 1) & 10).BackColor = &H80000005
    UF1(Left(TB.Name, 1) & 9).BackColor = &H80000005
  End Select
End If
End Sub
Public Sub NGErase(TB As MSForms.Label)  '各層でNGがあるかのチェック　ｸﾗｽ
Dim i As Long, FL As Boolean: FL = False
Dim Scol As Long: Scol = 13434828 'OK
Dim Ncol As Long: Ncol = 8421631 'NG
Dim TmpTB As String: TmpTB = Left(TB.Name, 1)
Dim Count As Long

If TmpTB = "F" Then Count = 10
If TmpTB <> "F" Then Count = 12

  For i = 1 To Count
    If UF1(Left(TB.Name, 1) & i) <> "" Then
      If UF1(Left(TB.Name, 1) & i).BackColor = Ncol Then
        FL = True '1つでもNGがある
        Exit For
      End If
    End If
  Next i

'ラテラルボー
Select Case TB.Name
Case "LateralS":  If UF1.LateralS.BackColor = Ncol Then FL = True
Case "LateralU":  If UF1.LateralU.BackColor = Ncol Then FL = True
Case "LateralC":  If UF1.LateralC.BackColor = Ncol Then FL = True
Case "LateralD":  If UF1.LateralD.BackColor = Ncol Then FL = True
Case "LateralE":  If UF1.LateralE.BackColor = Ncol Then FL = True
Case "LateralF":  If UF1.LateralF.BackColor = Ncol Then FL = True
'直角度
Case "DiagSKou":  If UF1.DiagSKou.BackColor = Ncol Then FL = True
Case "DiagUKou":  If UF1.DiagUKou.BackColor = Ncol Then FL = True
Case "DiagCKou":  If UF1.DiagCKou.BackColor = Ncol Then FL = True
Case "DiagDKou":  If UF1.DiagDKou.BackColor = Ncol Then FL = True
Case "DiagEKou":  If UF1.DiagEKou.BackColor = Ncol Then FL = True
Case "DiagFKou":  If UF1.DiagFKou.BackColor = Ncol Then FL = True
'歪A
Case "ZS":        If UF1.ZS.BackColor = Ncol Then FL = True
Case "ZU":        If UF1.ZU.BackColor = Ncol Then FL = True
Case "ZC":        If UF1.ZC.BackColor = Ncol Then FL = True
Case "ZD":        If UF1.ZD.BackColor = Ncol Then FL = True
Case "ZE":        If UF1.ZE.BackColor = Ncol Then FL = True
Case "ZF":        If UF1.ZF.BackColor = Ncol Then FL = True
'歪ｹﾞｰｼﾞ
Case "GS":        If UF1.GS.BackColor = Ncol Then FL = True
Case "GU":        If UF1.GU.BackColor = Ncol Then FL = True
Case "GC":        If UF1.GC.BackColor = Ncol Then FL = True
Case "GD":        If UF1.GD.BackColor = Ncol Then FL = True
Case "GE":        If UF1.GE.BackColor = Ncol Then FL = True
Case "GF":        If UF1.GF.BackColor = Ncol Then FL = True
End Select

'NGがすべて消えたらフラグも消す
If FL = False Then UF1("NG" & Left(TB.Name, 1)) = ""
End Sub

Public Sub FMCheck(fm As MSForms.frame) 'ｻｲｽﾞ
UF1.MPEX.Tag = "Click"
If UF1.Load.Caption = "off" Then Exit Sub
Dim Colo As String: Colo = 39423
UF1("FM" & Right((fm.Name), 1)).Width = 390
UF1("FM" & Right((fm.Name), 1)).BackColor = Colo
UF1("LZ" & Right((fm.Name), 1)).BackColor = Colo
UF1("LG" & Right((fm.Name), 1)).BackColor = Colo
Select Case Right((fm.Name), 1)
Case "S": Call Fm_Clear("S")
Case "U": Call Fm_Clear("U")
Case "C": Call Fm_Clear("C")
Case "D": Call Fm_Clear("D")
Case "E": Call Fm_Clear("E")
Case "F": Call Fm_Clear("F")
End Select
'UF1.MP1.Visible = False
End Sub
Public Sub Fm_Clear(n As String) 'ﾘｾｯﾄ
Dim Colo As String: Colo = &H8000000F
Dim Siz As Long: Siz = 390 '197.95 ' 半分用
Dim SizS As Long: SizS = 197.95

If IsNumeric(UF1.ZS) Then UF1.ZS = Format((UF1.ZS), "0.00") '2020.12.17_Exitと同じ処理
If IsNumeric(UF1.ZU) Then UF1.ZU = Format((UF1.ZU), "0.00")
If IsNumeric(UF1.ZC) Then UF1.ZC = Format((UF1.ZC), "0.00")
If IsNumeric(UF1.ZD) Then UF1.ZD = Format((UF1.ZD), "0.00")
If IsNumeric(UF1.ZE) Then UF1.ZE = Format((UF1.ZE), "0.00")
If IsNumeric(UF1.ZF) Then UF1.ZF = Format((UF1.ZF), "0.00") '2024.5.22

With UF1
Select Case n
Case "S"
.FMU.Width = Siz
.FMC.Width = Siz
.FMD.Width = Siz
.FME.Width = Siz
.FMF.Width = Siz
.FMU.BackColor = Colo
.FMC.BackColor = Colo
.FMD.BackColor = Colo
.FME.BackColor = Colo
.FMF.BackColor = Colo
.LZU.BackColor = Colo
.LGU.BackColor = Colo
.LZC.BackColor = Colo
.LGC.BackColor = Colo
.LZD.BackColor = Colo
.LGD.BackColor = Colo
.LZE.BackColor = Colo
.LGE.BackColor = Colo
.LZF.BackColor = Colo
.LGF.BackColor = Colo
Case "U"
.FMS.Width = Siz
.FMC.Width = Siz
.FMD.Width = Siz
.FME.Width = Siz
.FMF.Width = Siz
.FMS.BackColor = Colo
.FMC.BackColor = Colo
.FMD.BackColor = Colo
.FME.BackColor = Colo
.FMF.BackColor = Colo
.LZS.BackColor = Colo
.LGS.BackColor = Colo
.LZC.BackColor = Colo
.LGC.BackColor = Colo
.LZD.BackColor = Colo
.LGD.BackColor = Colo
.LZE.BackColor = Colo
.LGE.BackColor = Colo
.LZF.BackColor = Colo
.LGF.BackColor = Colo
Case "C"
.FMS.Width = Siz
.FMU.Width = Siz
.FMD.Width = Siz
.FME.Width = Siz
.FMF.Width = Siz
.FMS.BackColor = Colo
.FMU.BackColor = Colo
.FMD.BackColor = Colo
.FME.BackColor = Colo
.FMF.BackColor = Colo
.LZS.BackColor = Colo
.LGS.BackColor = Colo
.LZU.BackColor = Colo
.LGU.BackColor = Colo
.LZD.BackColor = Colo
.LGD.BackColor = Colo
.LZE.BackColor = Colo
.LGE.BackColor = Colo
.LZF.BackColor = Colo
.LGF.BackColor = Colo
Case "D"
.FMS.Width = Siz
.FMU.Width = Siz
.FMC.Width = Siz
.FME.Width = Siz
.FMF.Width = Siz
.FMS.BackColor = Colo
.FMU.BackColor = Colo
.FMC.BackColor = Colo
.FME.BackColor = Colo
.FMF.BackColor = Colo
.LZS.BackColor = Colo
.LGS.BackColor = Colo
.LZU.BackColor = Colo
.LGU.BackColor = Colo
.LZC.BackColor = Colo
.LGC.BackColor = Colo
.LZE.BackColor = Colo
.LGE.BackColor = Colo
.LZF.BackColor = Colo
.LGF.BackColor = Colo
Case "E"
.FMS.Width = Siz
.FMU.Width = Siz
.FMC.Width = Siz
.FMD.Width = Siz
.FMF.Width = Siz
.FMS.BackColor = Colo
.FMU.BackColor = Colo
.FMC.BackColor = Colo
.FMD.BackColor = Colo
.FMF.BackColor = Colo
.LZS.BackColor = Colo
.LGS.BackColor = Colo
.LZU.BackColor = Colo
.LGU.BackColor = Colo
.LZC.BackColor = Colo
.LGC.BackColor = Colo
.LZD.BackColor = Colo
.LGD.BackColor = Colo
.LZF.BackColor = Colo
.LGF.BackColor = Colo
Case "F"
.FMS.Width = Siz
.FMU.Width = Siz
.FMC.Width = Siz
.FMD.Width = Siz
.FME.Width = Siz
.FMS.BackColor = Colo
.FMU.BackColor = Colo
.FMC.BackColor = Colo
.FMD.BackColor = Colo
.FME.BackColor = Colo
.LZS.BackColor = Colo
.LGS.BackColor = Colo
.LZU.BackColor = Colo
.LGU.BackColor = Colo
.LZC.BackColor = Colo
.LGC.BackColor = Colo
.LZD.BackColor = Colo
.LGD.BackColor = Colo
.LZE.BackColor = Colo
.LGE.BackColor = Colo
Case "Click"
.FMS.Width = Siz
.FMU.Width = Siz
.FMC.Width = Siz
.FMD.Width = Siz
.FME.Width = Siz
.FMF.Width = Siz
.FMS.BackColor = Colo
.FMU.BackColor = Colo
.FMC.BackColor = Colo
.FMD.BackColor = Colo
.FME.BackColor = Colo
.FMF.BackColor = Colo
.LZS.BackColor = Colo
.LGS.BackColor = Colo
.LZU.BackColor = Colo
.LGU.BackColor = Colo
.LZC.BackColor = Colo
.LGC.BackColor = Colo
.LZD.BackColor = Colo
.LGD.BackColor = Colo
.LZE.BackColor = Colo
.LGE.BackColor = Colo
.LZF.BackColor = Colo
.LGF.BackColor = Colo
Case "All"
.FMS.Width = SizS
.FMU.Width = SizS
.FMC.Width = SizS
.FMD.Width = SizS
.FME.Width = SizS
.FMF.Width = SizS
.FMS.BackColor = Colo
.FMU.BackColor = Colo
.FMC.BackColor = Colo
.FMD.BackColor = Colo
.FME.BackColor = Colo
.FMF.BackColor = Colo
.LZS.BackColor = Colo
.LGS.BackColor = Colo
.LZU.BackColor = Colo
.LGU.BackColor = Colo
.LZC.BackColor = Colo
.LGC.BackColor = Colo
.LZD.BackColor = Colo
.LGD.BackColor = Colo
.LZE.BackColor = Colo
.LGE.BackColor = Colo
.LZF.BackColor = Colo
.LGF.BackColor = Colo
End Select
End With


'2021バージョンで実行するとtrueになる.....2024.2.28
'Dim i As Long, a As Long
'For i = 1 To 12
'  For a = 1 To 5
'    Select Case a
'    Case 1:  UF1("S" & i).Font.Underline = False
'    Case 2:  UF1("U" & i).Font.Underline = False
'    Case 3:  UF1("C" & i).Font.Underline = False
'    Case 4:  UF1("D" & i).Font.Underline = False
'    Case 5:  UF1("E" & i).Font.Underline = False
'    End Select
'  Next a
'Next i
'ﾘｾｯﾄ>>>>>>.Font.Underline = Falseをループで回すと2021バージョンだとすべてTrue処理される....関数を別途作って対応_2024.2.29
Call RemoveUnderlineFromLabels

End Sub
Private Function FindYellowLabel() As MSForms.Label ' 黄色のラベルを見つける関数_2024.8.20
Dim ctrl As Control
For Each ctrl In UF1.Controls
  If TypeOf ctrl Is MSForms.Label Then
    If ctrl.BackColor = RGB(255, 255, 0) Then ' 黄色
      Set FindYellowLabel = ctrl
      Exit Function
    End If
  End If
Next ctrl
Set FindYellowLabel = Nothing
End Function
' 選択されたコントロールの背景色を黄色にする_2024.8.20 歪コンボ、テキスト
Public Sub ChangeBackColor(ctrl As Control)
Dim Scol As Long: Scol = 13434828 'OK RGB(204,255,204)薄い緑
Dim Ncol As Long: Ncol = 8421631 'NG RGB(255,128,128)ｺｰﾗﾙ
If Not (ctrl.BackColor = Scol Or ctrl.BackColor = Ncol) Then
'  ResetBackColor ' すべてのコントロールの背景色をリセット
  ctrl.BackColor = RGB(255, 255, 0) ' フォーカスがあるコントロールの背景色を黄色にする
End If
End Sub
Public Sub ResetBackColorIfNotRedOrGreen(ctrl As Control)
Dim Scol As Long: Scol = 13434828 'OK RGB(204,255,204)薄い緑
Dim Ncol As Long: Ncol = 8421631 'NG RGB(255,128,128)ｺｰﾗﾙ

If Not (ctrl.BackColor = Scol Or ctrl.BackColor = Ncol) Then
  ctrl.BackColor = &H80000005 ' デフォルト色
End If
End Sub

' すべてのコントロールの背景色をデフォルトに戻す_2024.8.20
Public Sub ResetBackColor(Optional n As String = "")
If n = "" Then
  Dim YellowLabel As MSForms.Label
  Set YellowLabel = FindYellowLabel() ' 黄色のラベルを見つける
  If Not YellowLabel Is Nothing Then
    YellowLabel.BackColor = &H80000005 ' 黄色のラベルが見つかった場合、デフォルト色に戻す
  End If
End If

ResetBackColorIfNotRedOrGreen UF1.CBS1
ResetBackColorIfNotRedOrGreen UF1.CBS2
ResetBackColorIfNotRedOrGreen UF1.CBS3
ResetBackColorIfNotRedOrGreen UF1.CBS4

ResetBackColorIfNotRedOrGreen UF1.CBU1
ResetBackColorIfNotRedOrGreen UF1.CBU2
ResetBackColorIfNotRedOrGreen UF1.CBU3
ResetBackColorIfNotRedOrGreen UF1.CBU4

ResetBackColorIfNotRedOrGreen UF1.CBC1
ResetBackColorIfNotRedOrGreen UF1.CBC2
ResetBackColorIfNotRedOrGreen UF1.CBC3
ResetBackColorIfNotRedOrGreen UF1.CBC4

ResetBackColorIfNotRedOrGreen UF1.CBD1
ResetBackColorIfNotRedOrGreen UF1.CBD2
ResetBackColorIfNotRedOrGreen UF1.CBD3
ResetBackColorIfNotRedOrGreen UF1.CBD4

ResetBackColorIfNotRedOrGreen UF1.CBE1
ResetBackColorIfNotRedOrGreen UF1.CBE2
ResetBackColorIfNotRedOrGreen UF1.CBE3
ResetBackColorIfNotRedOrGreen UF1.CBE4

ResetBackColorIfNotRedOrGreen UF1.CBF1
ResetBackColorIfNotRedOrGreen UF1.CBF2
ResetBackColorIfNotRedOrGreen UF1.CBF3
ResetBackColorIfNotRedOrGreen UF1.CBF4

ResetBackColorIfNotRedOrGreen UF1.ZS
ResetBackColorIfNotRedOrGreen UF1.GS

ResetBackColorIfNotRedOrGreen UF1.ZU
ResetBackColorIfNotRedOrGreen UF1.GU

ResetBackColorIfNotRedOrGreen UF1.ZC
ResetBackColorIfNotRedOrGreen UF1.GC

ResetBackColorIfNotRedOrGreen UF1.ZD
ResetBackColorIfNotRedOrGreen UF1.GD

ResetBackColorIfNotRedOrGreen UF1.ZE
ResetBackColorIfNotRedOrGreen UF1.GE

ResetBackColorIfNotRedOrGreen UF1.ZF
ResetBackColorIfNotRedOrGreen UF1.GF

ResetBackColorIfNotRedOrGreen UF1.DiagSKou
ResetBackColorIfNotRedOrGreen UF1.DiagUKou
ResetBackColorIfNotRedOrGreen UF1.DiagCKou
ResetBackColorIfNotRedOrGreen UF1.DiagDKou
ResetBackColorIfNotRedOrGreen UF1.DiagEKou
ResetBackColorIfNotRedOrGreen UF1.DiagFKou

ResetBackColorIfNotRedOrGreen UF1.LateralS
ResetBackColorIfNotRedOrGreen UF1.LateralU
ResetBackColorIfNotRedOrGreen UF1.LateralC
ResetBackColorIfNotRedOrGreen UF1.LateralD
ResetBackColorIfNotRedOrGreen UF1.LateralE
ResetBackColorIfNotRedOrGreen UF1.LateralF
End Sub


Public Function CBdateCheck() As Boolean '歪入力の有無ﾁｪｯｸ_2021.09.20測定値0の場合はスルー
Dim a As Long
Dim S As Boolean, U As Boolean, C As Boolean, D As Boolean, E As Boolean
  For a = 1 To 5
    Select Case a
    Case 1: If (UF1.CBS1 <> "" And UF1.CBS2 <> "" And UF1.ZS <> "") Or (UF1.ZS = 0) Then S = True
    Case 2: If (UF1.CBU1 <> "" And UF1.CBU2 <> "" And UF1.ZU <> "") Or (UF1.ZU = 0) Then U = True
    Case 3: If (UF1.CBC1 <> "" And UF1.CBC2 <> "" And UF1.ZC <> "") Or (UF1.ZC = 0) Then C = True
    Case 4: If (UF1.CBD1 <> "" And UF1.CBD2 <> "" And UF1.ZD <> "") Or (UF1.ZD = 0) Then D = True
    Case 5: If (UF1.CBE1 <> "" And UF1.CBE2 <> "" And UF1.ZE <> "") Or (UF1.ZE = 0) Then E = True
    End Select
  Next a
If S Or U Or C Or D Or E Then CBdateCheck = True '歪の測定が最低でも1カ所ある
End Function

Public Sub AtCountSelect(n As Long) '薄板時の厚検査回数ｾｯﾄ
Dim i As Long
Dim flag(1 To 6) As Boolean
''条件定義****2022.01.11
'If UF1.AtButt.Visible And UF1.InspeButt.Visible And UF1.Yama.Visible Then flag(1) = True '冷延あり 6mm以下
'If UF1.AtButt.Visible And UF1.InspeButt.Visible = False And UF1.Yama.Visible = False Then flag(2) = True '冷延あり 6mm超え
'If UF1.Yama = 3 Then flag(3) = True  '山数が3*********山数3を選択時県有枚数選択は消える為flag1は条件に入れることができない
'If flag(1) And UF1.InspeButt.Caption = "検入≧26" Then flag(4) = True '検入25枚以下
'If UF1.AtButt.Visible = False And UF1.InspeButt.Visible = False And UF1.Yama.Visible = False Then flag(5) = True '冷延なし

'2023.4.17****検番による分岐に変更_定義のコース部分は不要
'2023.4.17_検査No検索
Dim SIKAHIKAcc: SIKAHIKAcc = Arrbox("SIKAHIKINOW.accdb", "", "Fields", UF1.DBflag.Text) 'SIKAHIKINOW.accdbのフィールドに[ｵｰﾀﾞｰ板丈]がないのでエラーになる
Dim no As Long
If Not IsEmpty(SIKAHIKAcc) Then
  For i = 1 To UBound(SIKAHIKAcc, 2)
    Select Case SIKAHIKAcc(1, i)
    Case "検査番号": no = i
    End Select
  Next i
End If
Dim KenAcc: KenAcc = Arrbox("SIKAHIKINOW.accdb", UF1.Lot, "Record", UF1.DBflag.Text)

If IsEmpty(KenAcc) Then 'SIKALOTNOWでやりなおし
    SIKAHIKAcc = Arrbox("SIKALOTNOW.accdb", "", "Fields", UF1.DBflag.Text)
    If Not IsEmpty(SIKAHIKAcc) Then
      For i = 1 To UBound(SIKAHIKAcc, 2)
        Select Case SIKAHIKAcc(1, i)
        Case "検査番号": no = i
        End Select
      Next i
    End If
    KenAcc = Arrbox("SIKALOTNOW.accdb", UF1.Lot, "Record", UF1.DBflag.Text)
End If

Dim KenNo As String
If IsEmpty(KenAcc) = False Then KenNo = Left(KenAcc(1, no), 1)  '検査No'頭1桁

'条件定義****2022.4.18
If KenNo = "L" Or KenNo = "N" Or KenNo = "K" Or KenNo = "T" Or KenNo = "S" Or KenNo = "F" Then flag(1) = True
If KenNo = "H" Then flag(2) = True
If UF1.Yama = 3 Then flag(3) = True  '山数が3*********山数3を選択時県有枚数選択は消える為flag1は条件に入れることができない
If flag(1) And UF1.InspeButt.Caption = "検入≧26" Then flag(4) = True '検入25枚以下
If UF1.Yama = 4 Then flag(6) = True  '山数が4

Select Case n
Case 2
  UF1.AtCount = "厚2ヵ所検査中"
  UF1.AtButt.Caption = "厚4ヵ所"
  For i = 3 To 4
    UF1("AL" & i).Visible = False
    If flag(6) Then '6回_2024.5.27　優先
      UF1("S" & i).Visible = False
      UF1("U" & i).Visible = False
      UF1("C" & i).Visible = False
      UF1("D" & i).Visible = False
      UF1("E" & i).Visible = False
      UF1("F" & i).Visible = False
      UF1("S" & i).Enabled = False
      UF1("U" & i).Enabled = False
      UF1("C" & i).Enabled = False
      UF1("D" & i).Enabled = False
      UF1("E" & i).Enabled = False
      UF1("F" & i).Enabled = False
    ElseIf flag(4) Then '2
      UF1("S" & i).Visible = False
      UF1("E" & i).Visible = False
      UF1("S" & i).Enabled = False
      UF1("E" & i).Enabled = False
    ElseIf flag(3) Then '5
      UF1("S" & i).Visible = False
      UF1("U" & i).Visible = False
      UF1("C" & i).Visible = False
      UF1("D" & i).Visible = False
      UF1("E" & i).Visible = False
      UF1("S" & i).Enabled = False
      UF1("U" & i).Enabled = False
      UF1("C" & i).Enabled = False
      UF1("D" & i).Enabled = False
      UF1("E" & i).Enabled = False
'    ElseIf flag(5) Then '5'冷延なしの場合強制表示
'      UF1("AL" & i).Visible = True
'      UF1.AtCount = "厚4ヵ所検査中"
'      UF1.AtButt.Caption = "厚2ヵ所"
'      UF1("S" & i).Visible = True
'      UF1("U" & i).Visible = True
'      UF1("C" & i).Visible = True
'      UF1("D" & i).Visible = True
'      UF1("E" & i).Visible = True
'      UF1("S" & i).Enabled = True
'      UF1("U" & i).Enabled = True
'      UF1("C" & i).Enabled = True
'      UF1("D" & i).Enabled = True
'      UF1("E" & i).Enabled = True
    ElseIf flag(1) Then '3
      UF1("S" & i).Visible = False
      UF1("C" & i).Visible = False
      UF1("E" & i).Visible = False
      UF1("S" & i).Enabled = False
      UF1("C" & i).Enabled = False
      UF1("E" & i).Enabled = False
    ElseIf flag(2) Then '5
      UF1("S" & i).Visible = False
      UF1("U" & i).Visible = False
      UF1("C" & i).Visible = False
      UF1("D" & i).Visible = False
      UF1("E" & i).Visible = False
      UF1("S" & i).Enabled = False
      UF1("U" & i).Enabled = False
      UF1("C" & i).Enabled = False
      UF1("D" & i).Enabled = False
      UF1("E" & i).Enabled = False
    End If
  Next i
Case 4
  UF1.AtCount = "厚4ヵ所検査中"
  UF1.AtButt.Caption = "厚2ヵ所"
  For i = 3 To 4
    UF1("AL" & i).Visible = True
    If flag(6) Then '6回　flag(1)より優先
      UF1("S" & i).Visible = True
      UF1("U" & i).Visible = True
      UF1("C" & i).Visible = True
      UF1("D" & i).Visible = True
      UF1("E" & i).Visible = True
      UF1("F" & i).Visible = True
      UF1("S" & i).Enabled = True
      UF1("U" & i).Enabled = True
      UF1("C" & i).Enabled = True
      UF1("D" & i).Enabled = True
      UF1("E" & i).Enabled = True
      UF1("F" & i).Enabled = True
    ElseIf flag(4) Then '2
      UF1("S" & i).Visible = True
      UF1("E" & i).Visible = True
      UF1("S" & i).Enabled = True
      UF1("E" & i).Enabled = True
    ElseIf flag(3) Then '5
      UF1("S" & i).Visible = True
      UF1("U" & i).Visible = True
      UF1("C" & i).Visible = True
      UF1("D" & i).Visible = True
      UF1("E" & i).Visible = True
      UF1("S" & i).Enabled = True
      UF1("U" & i).Enabled = True
      UF1("C" & i).Enabled = True
      UF1("D" & i).Enabled = True
      UF1("E" & i).Enabled = True
  '    ElseIf flag(5) Then '5'冷延なしの場合強制表示
  '      UF1("AL" & i).Visible = True
  '      UF1.AtCount = "厚4ヵ所検査中"
  '      UF1.AtButt.Caption = "厚2ヵ所"
  '      UF1("S" & i).Visible = True
  '      UF1("U" & i).Visible = True
  '      UF1("C" & i).Visible = True
  '      UF1("D" & i).Visible = True
  '      UF1("E" & i).Visible = True
  '      UF1("S" & i).Enabled = True
  '      UF1("U" & i).Enabled = True
  '      UF1("C" & i).Enabled = True
  '      UF1("D" & i).Enabled = True
  '      UF1("E" & i).Enabled = True
    ElseIf flag(1) Then '3
      UF1("S" & i).Visible = True
      UF1("C" & i).Visible = True
      UF1("E" & i).Visible = True
      UF1("S" & i).Enabled = True
      UF1("C" & i).Enabled = True
      UF1("E" & i).Enabled = True
    ElseIf flag(2) Then '5
      UF1("S" & i).Visible = True
      UF1("U" & i).Visible = True
      UF1("C" & i).Visible = True
      UF1("D" & i).Visible = True
      UF1("E" & i).Visible = True
      UF1("S" & i).Enabled = True
      UF1("U" & i).Enabled = True
      UF1("C" & i).Enabled = True
      UF1("D" & i).Enabled = True
      UF1("E" & i).Enabled = True
    End If
  Next i
End Select

'If InStr(UF1.Reien.Caption, "2HI") > 0 Then '2HI通過時強制4か所表示
If KenNo = "T" Then
  For i = 3 To 4
    If flag(6) Then '6
      UF1("AL" & i).Visible = True
      UF1("S" & i).Visible = True
      UF1("U" & i).Visible = True
      UF1("C" & i).Visible = True
      UF1("D" & i).Visible = True
      UF1("E" & i).Visible = True
      UF1("F" & i).Visible = True
      UF1("S" & i).Enabled = True
      UF1("U" & i).Enabled = True
      UF1("C" & i).Enabled = True
      UF1("D" & i).Enabled = True
      UF1("E" & i).Enabled = True
      UF1("F" & i).Visible = True
    ElseIf flag(4) Then '2
      UF1("AL" & i).Visible = True
      UF1("S" & i).Visible = True
      UF1("E" & i).Visible = True
      UF1("S" & i).Enabled = True
      UF1("E" & i).Enabled = True
    ElseIf flag(3) Then '5
      UF1("AL" & i).Visible = True
      UF1("S" & i).Visible = True
      UF1("U" & i).Visible = True
      UF1("C" & i).Visible = True
      UF1("D" & i).Visible = True
      UF1("E" & i).Visible = True
      UF1("S" & i).Enabled = True
      UF1("U" & i).Enabled = True
      UF1("C" & i).Enabled = True
      UF1("D" & i).Enabled = True
      UF1("E" & i).Enabled = True
    Else '3
      UF1("AL" & i).Visible = True
      UF1("S" & i).Visible = True
      UF1("C" & i).Visible = True
      UF1("E" & i).Visible = True
      UF1("S" & i).Enabled = True
      UF1("C" & i).Enabled = True
      UF1("E" & i).Enabled = True
    End If
  Next i
End If

End Sub

Public Sub InspectionSelect(n As String)    '薄板時の検入枚数_2021.11.29
Dim i As Long
'2022.01.11_条件定義
Dim flag(1 To 6) As Boolean
'If UF1.AtButt.Visible And UF1.InspeButt.Visible And UF1.Yama.Visible Then flag(1) = True '冷延あり 6mm以下
'If UF1.AtButt.Visible And UF1.InspeButt.Visible = False And UF1.Yama.Visible = False Then flag(2) = True '冷延あり 6mm超え
'If UF1.Yama = 3 Then flag(3) = True      '山数が3
'If flag(1) And UF1.InspeButt.Caption = "検入≧26" Then flag(4) = True '検入25枚以下
'If UF1.AtButt.Visible = False And UF1.InspeButt.Visible = False And UF1.Yama.Visible = False Then flag(5) = True '冷延なし

'2023.4.17****検番による分岐に変更_定義のコース部分は不要
'2023.4.17_検査No検索
Dim SIKAHIKAcc: SIKAHIKAcc = Arrbox("SIKAHIKINOW.accdb", "", "Fields", UF1.DBflag.Text) 'SIKAHIKINOW.accdbのフィールドに[ｵｰﾀﾞｰ板丈]がないのでエラーになる
Dim no As Long
If Not IsEmpty(SIKAHIKAcc) Then
  For i = 1 To UBound(SIKAHIKAcc, 2)
    Select Case SIKAHIKAcc(1, i)
    Case "検査番号": no = i
    End Select
  Next i
End If
Dim KenAcc: KenAcc = Arrbox("SIKAHIKINOW.accdb", UF1.Lot, "Record", UF1.DBflag.Text)

If IsEmpty(KenAcc) Then 'SIKALOTNOWでやりなおし
    SIKAHIKAcc = Arrbox("SIKALOTNOW.accdb", "", "Fields", UF1.DBflag.Text)
    If Not IsEmpty(SIKAHIKAcc) Then
      For i = 1 To UBound(SIKAHIKAcc, 2)
        Select Case SIKAHIKAcc(1, i)
        Case "検査番号": no = i
        End Select
      Next i
    End If
    KenAcc = Arrbox("SIKALOTNOW.accdb", UF1.Lot, "Record", UF1.DBflag.Text)
End If

Dim KenNo As String
If IsEmpty(KenAcc) = False Then KenNo = Left(KenAcc(1, no), 1)  '検査No'頭1桁

'条件定義****2022.4.18
If KenNo = "L" Or KenNo = "N" Or KenNo = "K" Or KenNo = "T" Or KenNo = "S" Or KenNo = "F" Then flag(1) = True
If KenNo = "H" Then flag(2) = True
If UF1.Yama = 3 Then flag(3) = True  '山数が3*********山数3を選択時県有枚数選択は消える為flag1は条件に入れることができない
If flag(1) And UF1.InspeButt.Caption = "検入≧26" Then flag(4) = True '検入25枚以下
If UF1.Yama = 4 Then flag(6) = True  '山数が4

Select Case n
Case "非表示"
  UF1.InspeButt.Caption = "検入≧26"
  For i = 1 To 10
    UF1("C" & i).Visible = False
    UF1("C" & i).Enabled = False
  Next i
  UF1("DiagC").Visible = False
'  UF1("DiagCKou").Visible = False
  UF1.FC.Caption = ""
  UF1.C.Caption = ""
  
  Select Case UF1.Yama '2022.03.22**追加Load時うまく読み込めない対策
  Case 1
    UF1.S = "S"
    UF1.C = ""
    UF1.E = "E"
    UF1.FS = "S"
    UF1.FC = ""
    UF1.FE = "E"
    UF1.OPS1.Caption = "S"
    UF1.OPS3.Caption = ""
    UF1.OPS5.Caption = "E"
    UF1.OPE1.Caption = "S"
    UF1.OPE3.Caption = ""
    UF1.OPE5.Caption = "E"
    UF1.U = ""
    UF1.FU = ""
    UF1.D = ""
    UF1.FD = ""
  Case 2
    UF1.S = "2-2-S"
    UF1.U = ""
    UF1.C = ""
    UF1.D = ""
    UF1.E = "2-1-E"
    UF1.FS = "2-2-S"
    UF1.FU = ""
    UF1.FC = ""
    UF1.FD = ""
    UF1.FE = "2-1-E"
    UF1.OPS1.Caption = "2-2-S"
    UF1.OPS2.Caption = ""
    UF1.OPS3.Caption = "2-2-E"
    UF1.OPS4.Caption = ""
    UF1.OPS5.Caption = "2-2-E"
    UF1.OPE1.Caption = "2-2-S"
    UF1.OPE2.Caption = ""
    UF1.OPE3.Caption = "2-2-E"
    UF1.OPE4.Caption = ""
    UF1.OPE5.Caption = "2-1-E"
    UF1.U = ""
    UF1.FU = ""
    UF1.D = ""
    UF1.FD = ""
  Case 3
    UF1.S = "3-3-S"
    UF1.U = ""
    UF1.C = ""
    UF1.D = ""
    UF1.E = "3-1-E"
    UF1.FS = "3-3-S"
    UF1.FU = ""
    UF1.FC = ""
    UF1.FD = ""
    UF1.FE = "3-1-E"
    UF1.OPS1.Caption = "3-3-S"
    UF1.OPS2.Caption = ""
    UF1.OPS3.Caption = ""
    UF1.OPS4.Caption = ""
    UF1.OPS5.Caption = "3-1-E"
    UF1.OPE1.Caption = "3-3-S"
    UF1.OPE2.Caption = ""
    UF1.OPE3.Caption = ""
    UF1.OPE4.Caption = ""
    UF1.OPE5.Caption = "3-1-E"
  End Select
  
  Call DistSelect(25) '*****************歪表示2回
  
    For i = 2 To 4 '外観表示
      UF1("LC" & i).Visible = False
    Next i
    For i = 8 To 10
      UF1("LC" & i).Visible = False
    Next i
    
    UF1.U_UpVisual.Visible = False
    UF1.C_UpVisual.Visible = False
    UF1.D_UpVisual.Visible = False
    UF1.U_UnderVisual.Visible = False
    UF1.C_UnderVisual.Visible = False
    UF1.D_UnderVisual.Visible = False
  
    'BH_2023.5.22
    UF1.UBH.Visible = False
    UF1.CBH.Visible = False
    UF1.DBH.Visible = False
    
  '2023.12.07_検査2回
  If UF1.Sumisei.Caption <> "" Or UF1.LateralText.Tag = "有" Then '2026.07.13_技術指示経路も対象に追加
    UF1.LateralU.Visible = False
    UF1.LateralC.Visible = False
    UF1.LateralD.Visible = False
  End If
Case "表示"
  UF1.InspeButt.Caption = "検入≦25"
'  If InStr(UF1.Reien.Caption, "2HI") > 0 Then '2023.3.16**2Hiの条件追加により厚測定2，4選択ボタンの有無の条件もここに追加'2HI通過時強制4か所表示
  If KenNo = "T" Then
    For i = 1 To 10
      UF1("C" & i).Visible = True
      UF1("C" & i).Enabled = True
    Next i
  ElseIf UF1.AtButt.Caption = "厚2ヵ所" Then
    For i = 1 To 10
      UF1("C" & i).Visible = True
      UF1("C" & i).Enabled = True
    Next i
  ElseIf UF1.AtButt.Caption = "厚4ヵ所" Then
    For i = 1 To 10
      Select Case i
      Case 1, 2
        UF1("C" & i).Visible = True
        UF1("C" & i).Enabled = True
      Case 3, 4
        UF1("C" & i).Visible = False
        UF1("C" & i).Enabled = False
      Case 5 To 10
        UF1("C" & i).Visible = True
        UF1("C" & i).Enabled = True
      End Select
    Next i
  End If
'  If flag(5) Then Call DistSelect(5) '歪表示5回


  If flag(6) Then '歪表示6回
  '別関数処理
  ElseIf flag(1) Then
    Call DistSelect(26) '歪表示3回
    
    '外観表示
    UF1("LC" & 2).Visible = False
    UF1("LC" & 4).Visible = False
    UF1("LC" & 3).Visible = True
    
    UF1("LC" & 8).Visible = False
    UF1("LC" & 10).Visible = False
    UF1("LC" & 9).Visible = True
    
    UF1.U_UpVisual.Visible = False
    UF1.D_UpVisual.Visible = False
    UF1.C_UpVisual.Visible = True
    UF1.U_UnderVisual.Visible = False
    UF1.D_UnderVisual.Visible = False
    UF1.C_UnderVisual.Visible = True
    
    'BH_2023.5.22
    UF1.UBH.Visible = False
    UF1.CBH.Visible = True
    UF1.DBH.Visible = False
    '2023.12.07_検査3回
    If UF1.Sumisei.Caption <> "" Or UF1.LateralText.Tag = "有" Then '2026.07.13_技術指示経路も対象に追加
      UF1.LateralU.Visible = False
      UF1.LateralC.Visible = True
      UF1.LateralD.Visible = False
    End If
  End If
'  If InStr(UF1.Reien.Caption, "2HI") > 0 Then Call DistSelect(26)   '歪表示3回**薄板2HI材_2023.3.16
  If KenNo = "T" Then Call DistSelect(26)    '歪表示3回*2HI材
  If flag(3) Or flag(2) Then '山3の場合どんな条件でも5回検査 '厚板の場合どんな条件でも5回検査
    Call DistSelect(5) '歪表示5回
  
    '外観表示
    For i = 2 To 4 '外観表示
      UF1("LC" & i).Visible = True
    Next i
    For i = 8 To 10
      UF1("LC" & i).Visible = True
    Next i
    
    UF1.U_UpVisual.Visible = True
    UF1.C_UpVisual.Visible = True
    UF1.D_UpVisual.Visible = True
    UF1.U_UnderVisual.Visible = True
    UF1.C_UpVisual.Visible = True
    UF1.D_UnderVisual.Visible = True
    
    'BH_2023.5.22
    UF1.UBH.Visible = True
    UF1.CBH.Visible = True
    UF1.DBH.Visible = True
  End If
  
'  If Flag(2) Then Call DistSelect(5)
  
  Select Case UF1.Yama '2022.03.22**追加Load時うまく読み込めない対策
  Case 1
    UF1.S = "S"
    UF1.C = "2"
    UF1.E = "E"
    UF1.FS = "S"
    UF1.FC = "2"
    UF1.FE = "E"
    UF1.OPS1.Caption = "S"
    UF1.OPS3.Caption = "2"
    UF1.OPS5.Caption = "E"
    UF1.OPE1.Caption = "S"
    UF1.OPE3.Caption = "2"
    UF1.OPE5.Caption = "E"
    UF1.U = ""
    UF1.FU = ""
    UF1.D = ""
    UF1.FD = ""
  Case 2
    UF1.S = "2-2-S"
    UF1.U = ""
    UF1.C = "2-2-E"
    UF1.D = ""
    UF1.E = "2-1-E"
    UF1.FS = "2-2-S"
    UF1.FU = ""
    UF1.FC = "2-2-E"
    UF1.FD = ""
    UF1.FE = "2-1-E"
    UF1.OPS1.Caption = "2-2-S"
    UF1.OPS2.Caption = ""
    UF1.OPS3.Caption = "2-2-E"
    UF1.OPS4.Caption = ""
    UF1.OPS5.Caption = "2-2-E"
    UF1.OPE1.Caption = "2-2-S"
    UF1.OPE2.Caption = ""
    UF1.OPE3.Caption = "2-2-E"
    UF1.OPE4.Caption = ""
    UF1.OPE5.Caption = "2-1-E"
    UF1.U = ""
    UF1.FU = ""
    UF1.D = ""
    UF1.FD = ""
  Case 3
    UF1.S = "3-3-S"
    UF1.U = "3-3-E"
    UF1.C = "3-2-2"
    UF1.D = "3-1-S"
    UF1.E = "3-1-E"
    UF1.FS = "3-3-S"
    UF1.FU = "3-3-E"
    UF1.FC = "3-2-2"
    UF1.FD = "3-1-S"
    UF1.FE = "3-1-E"
    UF1.OPS1.Caption = "3-3-S"
    UF1.OPS2.Caption = "3-3-E"
    UF1.OPS3.Caption = "3-2-2"
    UF1.OPS4.Caption = "3-1-S"
    UF1.OPS5.Caption = "3-1-E"
    UF1.OPE1.Caption = "3-3-S"
    UF1.OPE2.Caption = "3-3-E"
    UF1.OPE3.Caption = "3-2-2"
    UF1.OPE4.Caption = "3-1-S"
    UF1.OPE5.Caption = "3-1-E"
  End Select
  
  UF1("DiagC").Visible = True
'  UF1("DiagCKou").Visible = True
End Select
End Sub

Public Sub RightAngleDisplay(n As String, Optional Count As String = "") '直角度の表示
'Atu_OPselectでデフォルト表示が決まっているので変える場合はそっちも参照
'対角公差から直角度へ変更_2023.11.24
Dim LeftTemp As Long: LeftTemp = 804
Dim LeftTemp2 As Long: LeftTemp2 = 852
Dim LeftOut As Long: LeftOut = 906
Dim LeftOut2 As Long: LeftOut2 = 972
'位置調整 ｻｲｽﾞ
UF1.L10.Left = LeftTemp2
UF1.DiagSKou.Left = LeftTemp2
UF1.DiagUKou.Left = LeftTemp2
UF1.DiagCKou.Left = LeftTemp2
UF1.DiagDKou.Left = LeftTemp2
UF1.DiagEKou.Left = LeftTemp2
  
If UF1.Sumisei.Caption <> "" Then
'  If UF1.Youto = "R115" Then 'QA-S-QA-DIM-154_対角線差及び直角度の測定基準
    UF1.L10.Visible = True
    UF1.DiagSKou.Visible = True
    UF1.DiagUKou.Visible = True
    UF1.DiagCKou.Visible = True
    UF1.DiagDKou.Visible = True
    UF1.DiagEKou.Visible = True
'  End If
End If
'検査回数

If UF1.Yama = 4 Then UF1.DiagFKou.Visible = True
If UF1.Yama <> 4 Then UF1.DiagFKou.Visible = False
If Count <> "" Then UF1.DiagFKou.Visible = True

Select Case n
Case "N", "K", "T", "S", "F"
  Select Case UF1.Yama
  Case 3
    If UF1.Youto = "R115" Then
      UF1.DiagUKou.Visible = True
      UF1.DiagDKou.Visible = True
    End If
    '対角線
    UF1.DiagU.Visible = True
    UF1.DiagD.Visible = True
    UF1.DiagU.Enabled = True
    UF1.DiagD.Enabled = True
  Case 2, 1
    If UF1.Youto = "R115" Then
      UF1.DiagUKou.Visible = False
      UF1.DiagDKou.Visible = False
    End If
    '対角線
    UF1.DiagU.Visible = False
    UF1.DiagD.Visible = False
    UF1.DiagU.Enabled = False
    UF1.DiagD.Enabled = False
  End Select
Case "H"
Case "強制"
  UF1.L10.Visible = True
  UF1.DiagSKou.Visible = True
  UF1.DiagUKou.Visible = True
  UF1.DiagCKou.Visible = True
  UF1.DiagDKou.Visible = True
  UF1.DiagEKou.Visible = True
Case "新規"
  UF1.L10.Visible = False
  UF1.DiagSKou.Visible = False
  UF1.DiagUKou.Visible = False
  UF1.DiagCKou.Visible = False
  UF1.DiagDKou.Visible = False
  UF1.DiagEKou.Visible = False
End Select
End Sub

Public Sub ClearLSelectCache()
    g_CacheLotNo = ""
    g_CacheKenNo = ""
    g_CacheCours = ""
    g_CacheOrdAtu = 0
    g_CacheC_Course = ""
    g_SkipLotNotFoundMsg = True '追加：次回のLSelectでMsgBoxを1回だけ抑制
    Call DebugLog("[ClearLSelectCache] キャッシュクリア完了")
End Sub

'2027.07.10_引数追加
Public Sub LSelect(n As String, Optional bSkipDBCheck As Boolean = False) '検査選択　ラベル位置　フォームサイズ設定'住精追加_2021.07.13
Call DebugLog("[LSelect] 呼び出し開始 / n=" & n & " / UF1.Lot=[" & UF1.Lot.Text & "]")
Call frameblindfold("目隠し")

'耳付きについて丈の下部3,4の追加あり_2022.06.02
'スラブ追加_2022.06.07
On Error Resume Next
'オーダー厚のセット_2024.5.14
Dim LotNo As String: LotNo = UF1.Lot.Text
If SikaSetAtu <> 0 Then
  SikaAtuAll = SikaAtuCreate(LotNo)
  If Not IsEmpty(SikaAtuAll) Then SikaSetAtu = SikaAtuAll(LBound(SikaAtuAll))
End If

If UF1.Byousya Then GoTo LineSkip   '強制描写_2022.06*****LotNo検索に引っかからないと描写できないため

'LotNo,コース検索
'キャッシュ有効チェック_2026.06.03
'------------------------------------------------------
Dim i       As Long
Dim FS4flag As Boolean: FS4flag = False
Dim AtuFlag As Boolean: AtuFlag = False
Dim HIFlag  As Boolean: HIFlag = False
Dim LotAcc  As Variant
Dim KenNo   As String
Dim cours   As Long
Dim OrdAtu  As Long
Dim C_Course As Long

'------------------------------------------------------
'キャッシュ有効チェック_2026.06.03
'------------------------------------------------------
Dim bUseCache As Boolean: bUseCache = False
If g_CacheLotNo = UF1.Lot.Text _
    And g_CacheCours <> "" _
    And g_CacheOrdAtu <> 0 Then
    bUseCache = True
End If

Call DebugLog("[LSelect] キャッシュ有効チェック" & _
    " / bUseCache=" & bUseCache & _
    " / UF1.Lot=" & UF1.Lot.Text & _
    " / g_CacheLotNo=" & g_CacheLotNo & _
    " / g_CacheCours=[" & g_CacheCours & "]" & _
    " / g_CacheOrdAtu=" & g_CacheOrdAtu)

If bUseCache Then
    '--- キャッシュ使用 ---
    KenNo = g_CacheKenNo
    'KenNo が空または "0" の場合は FindPriority で補完
    If KenNo = "0" Or KenNo = "" Then
        Dim Result_C As String
        Result_C = FindPriority(g_CacheCours)
        If Result_C <> "" Then KenNo = Result_C
    End If
    
    Call DebugLog("[LSelect] KenNo確定(キャッシュ使用)" & _
        " / KenNo=" & KenNo & _
        " / g_CacheKenNo=" & g_CacheKenNo)
Else
    '--- キャッシュなし：従来通り Arrbox で取得 ---
    Dim ResultT As Boolean: ResultT = IsOver12Hours

    Dim HiAcc: HiAcc = Arrbox("SIKALOTNOW.accdb", "板", "Fields", UF1.DBflag.Text)
    If Not IsEmpty(HiAcc) Then
      For i = 1 To UBound(HiAcc, 2)
        Select Case HiAcc(1, i)
        Case "設計_設備ｺｰｽ": cours = i
        Case "ｵｰﾀﾞｰ板厚":    OrdAtu = i
        Case "実績_設備ｺｰｽ": C_Course = i
        End Select
      Next i
    ElseIf IsEmpty(HiAcc) Then
      If n = "住精" Then GoTo LineSkip
    End If

    If UF1.Lot.Text = "" Then '2026.08.06 UF1.Lot.Text = "" の場合は Arrbox の Record 取得を呼ばないようにガードを追加
        Call DebugLog("[LSelect] UF1.Lotが空のためArrbox(Record)をスキップ")
        LotAcc = Empty
    Else
        LotAcc = Arrbox("SIKALOTNOW.accdb", UF1.Lot, "Record", UF1.DBflag.Text)
    End If
    
    Call DebugLog("[LSelect] Arrbox取得結果 / UF1.Lot=[" & UF1.Lot.Text & "]" & _
        " / IsEmpty(LotAcc)=" & IsEmpty(LotAcc))
    
    If IsEmpty(LotAcc) Then
      If UF1.Load.Caption = "on" Then
        If ResultT = False And g_SkipLotNotFoundMsg = False Then
            MsgBox "SIKALOTNOWに該当ロットがありません" & "LSelect", vbCritical
        End If
      End If
      If n = "住精" Then GoTo LineSkip
    End If
    
    'フラグは1回使ったら必ずリセット_2026.08.06
    g_SkipLotNotFoundMsg = False

    Dim SIKAHIKAcc: SIKAHIKAcc = Arrbox("SIKAHIKINOW.accdb", "", "Fields", UF1.DBflag.Text)
    Dim no As Long
    If Not IsEmpty(SIKAHIKAcc) Then
      For i = 1 To UBound(SIKAHIKAcc, 2)
        Select Case SIKAHIKAcc(1, i)
        Case "検査番号": no = i
        End Select
      Next i
    End If
    Dim KenAcc: KenAcc = Arrbox("SIKAHIKINOW.accdb", UF1.Lot, "Record", UF1.DBflag.Text)

    If IsEmpty(KenAcc) Then
        SIKAHIKAcc = Arrbox("SIKALOTNOW.accdb", "", "Fields", UF1.DBflag.Text)
        If Not IsEmpty(SIKAHIKAcc) Then
          For i = 1 To UBound(SIKAHIKAcc, 2)
            Select Case SIKAHIKAcc(1, i)
            Case "検査番号": no = i
            End Select
          Next i
        End If
        KenAcc = Arrbox("SIKALOTNOW.accdb", UF1.Lot, "Record", UF1.DBflag.Text)
    End If

    If IsEmpty(KenAcc) = False Then KenNo = Left(KenAcc(1, no), 1)

    If KenNo = "0" Or KenNo = "" Then
        Dim InputString As String: InputString = LotAcc(1, C_Course)
        Dim Result_C2 As String:   Result_C2 = FindPriority(InputString)
        If Result_C2 <> "" Then KenNo = Result_C2
    End If

    Call DebugLog("[LSelect] KenNo確定(Arrbox使用)" & _
        " / KenNo=" & KenNo)
End If
'------------------------------------------------------

'コース表示
'条件定義
'追加_2023.03

Call DebugLog("[LSelect] 2HI判定直前 / bUseCache=" & bUseCache & _
    " / IsEmpty(LotAcc)=" & IsEmpty(LotAcc) & _
    " / g_CacheCours=[" & g_CacheCours & "]")

UF1.Reien.Caption = ""
UF1.Comment.Caption = ""  '追加：前回値の残留防止_2027.07.10

Dim strCoursForCheck As String
If bUseCache Then
    strCoursForCheck = g_CacheCours
ElseIf Not IsEmpty(LotAcc) Then
    strCoursForCheck = LotAcc(1, cours)
Else
    strCoursForCheck = ""
End If

Select Case True
Case InStr(strCoursForCheck, "2HI") > 0
  HIFlag = True
  UF1.Reien.Caption = KenNo & "_2HI"
  UF1.Comment.Caption = "2HI材:厚測定強制4ヵ所"
End Select

Call DebugLog("[LSelect] 2HI判定直後 / HIFlag=" & HIFlag & _
    " / Reien=" & UF1.Reien.Caption & _
    " / Comment=" & UF1.Comment.Caption)

Select Case True
Case InStr(strCoursForCheck, "L-2") > 0
  FS4flag = True
  If HIFlag Then UF1.Reien.Caption = KenNo & "_L-2" & "_2HI"
  If HIFlag = False Then UF1.Reien.Caption = KenNo & "_L-2"
Case InStr(strCoursForCheck, "L-3") > 0
  FS4flag = True
  If HIFlag Then UF1.Reien.Caption = KenNo & "_L-3" & "_2HI"
  If HIFlag = False Then UF1.Reien.Caption = KenNo & "_L-3"
End Select


'LL7対応_2022.01.28
'多分消すはず**
'If FS4flag Then
'  If InStr(LotAcc(1, Cours), "LL7") > 0 Then
'    UF1.Reien.Caption = KenNo & "_LL7"
'    Dim LL7Flag As Boolean: LL7Flag = True
'  End If
'End If
'変更後
If FS4flag Then
  If InStr(strCoursForCheck, "LL7") > 0 Then
    UF1.Reien.Caption = KenNo & "_LL7"
    Dim LL7Flag As Boolean: LL7Flag = True
  End If
End If


Dim NewAtu As Single: NewAtu = IIf(bUseCache, g_CacheOrdAtu, LotAcc(1, OrdAtu)) 'ｵｰﾀﾞｰ厚参照の控え_2022.01.11

If NewAtu = 0 Then
  If UF1.ATU <> "" Then
    NewAtu = val(UF1.ATU)
  End If
End If
If val(NewAtu) <= THIN_PLATE_ATU Then AtuFlag = True 'THIN_PLATE_ATU以下でTrue
Call DebugLog("[LSelect] NewAtu確定" & _
    " / NewAtu=" & NewAtu & _
    " / 取得元=" & IIf(bUseCache, "キャッシュ", "Arrbox") & _
    " / AtuFlag(" & THIN_PLATE_ATU & "mm以下)=" & (val(NewAtu) <= THIN_PLATE_ATU))

'ﾗﾍﾞﾙﾌｨｰﾙﾄﾞ
With UF1
Select Case n
Case "厚板", "薄板", "少枚数"
    .L1 = "幅1"
    .L2 = "幅2"
    .L3 = "丈1"
    .L4 = "丈2"
    .L5 = "対角A"
    .L6 = "対角B"
Case "極厚"
    .L1 = "幅1"
    .L2 = "幅2"
    .L3 = "幅3"
    .L4 = "幅4"
    .L5 = "幅5"
    .L6 = "幅6"
    .L7 = "丈1"
    .L8 = "丈2"
    .L12 = "丈3" '2022.06.02
    .L13 = "丈4"
End Select

'ボタン描写、検査位置描写
Select Case n
Case "厚板"
    .AtCount = ""
    .AtButt.Visible = False '2回4回
    .InspeButt.Visible = False '検入枚数25以上以下
    
    .Yama.Visible = False
    Call Six_Display("off") 'yamaの表示とセット_2024.5.28

    .YamaLabel.Visible = False
    .SlectFM.Width = 205
    .S = "S"
    .U = "S1"
    .C = "2"
    .D = "E1"
    .E = "E"
    .FS = "S"
    .FU = "S1"
    .FC = "2"
    .FD = "E1"
    .FE = "E"
    .OPS1.Caption = "S" '2021.09.06
    .OPS2.Caption = "S1"
    .OPS3.Caption = "2"
    .OPS4.Caption = "E1"
    .OPS5.Caption = "E"
    .OPE1.Caption = "S"
    .OPE2.Caption = "S1"
    .OPE3.Caption = "2"
    .OPE4.Caption = "E1"
    .OPE5.Caption = "E"
Case "少枚数", "極厚"
    .AtCount = ""
    .AtButt.Visible = False
    .InspeButt.Visible = False '検入枚数25以上以下
    .Yama.Visible = False
    Call Six_Display("off") 'yamaの表示とセット_2024.5.28
    
    .YamaLabel.Visible = False
    .SlectFM.Width = 205
    .S = "1"
    .U = "2"
    .C = "3"
    .D = "4"
    .E = "5"
    .FS = "1"
    .FU = "2"
    .FC = "3"
    .FD = "4"
    .FE = "5"
    .OPS1.Caption = "1" '2021.09.06
    .OPS2.Caption = "2"
    .OPS3.Caption = "3"
    .OPS4.Caption = "4"
    .OPS5.Caption = "5"
    .OPE1.Caption = "1"
    .OPE2.Caption = "2"
    .OPE3.Caption = "3"
    .OPE4.Caption = "4"
    .OPE5.Caption = "5"
'2023.03_条件見直し
'冷延条件が不要になる…？
Case "薄板"
'2024.6.7_追加
    If HIFlag = False Then '2HI通ってない→2か所検査ok　’HIFlag = Trueで2hi通過
      If AtuFlag = False Then '厚6mm以上あるならば4か所検査(検査番号Tのとき薄板を使用するため)　'AtuFlag6mm以下でTrue
        .AtButt.Visible = False '2hiを通過していなくても厚が6mm以上あればボタンは表示しない
      End If
    End If
    
    If HIFlag Then .AtButt.Visible = False '2HI通過は厚強制4か所検査　厚選択非表示
    
    .InspeButt.Visible = True '検入枚数25以上以下
    .Yama.Visible = True '山数
    .YamaLabel.Visible = True
    
    If .Yama.Text = "" Then Call Six_Display("off") 'yamaの表示とセット_2024.5.28
    
    If KenNo = "T" Then .InspeButt.Visible = False '2HI材は枚数ボタン非表示_2023.4.19

'  If LL7Flag Then '5回*****冷延あり　かつ　LL7通過
'  ElseIf (FS4flag) And (AtuFlag = False) Then '5回*****冷延あり　かつ　6mm超え＊＊＊＊＊＊＊＊＊＊＊＊＊この条件はAtu_OPselectで弾かれる*2023.03
'  ElseIf (FS4flag) And (AtuFlag) Then '3回
'  ElseIf (FS4flag) Then '5回*****冷延あり
'  ElseIf FS4flag = False Then '5回*****冷延なし＊＊＊＊＊＊＊＊＊＊＊＊＊この条件はAtu_OPselectで弾かれる*2023.03
'  End If
  
LineSkip:
  .SlectFM.Width = 420
  .AtButt.Caption = "厚2ヵ所"
  .AtCount = "厚4ヵ所検査中"
  If KenNo <> "T" Then .AtButt.Visible = True
  '検査位置ラベル表示
  Select Case UF1.Yama.Text
  Case "1"
    .S = "S"
    .C = "2"
    .E = "E"
    .FS = "S"
    .FC = "2"
    .FE = "E"
    .OPS1.Caption = "S" '2021.09.06
    .OPS3.Caption = "2"
    .OPS5.Caption = "E"
    .OPE1.Caption = "S"
    .OPE3.Caption = "2"
    .OPE5.Caption = "E"
    .U = ""
    .FU = ""
    .D = ""
    .FD = ""
    
    If KenNo <> "T" Then .AtButt.Visible = True '検番T
    
'    If FS4flag And HIFlag = False Then .AtButt.Visible = True
    If .Yama.Visible Then .InspeButt.Visible = True 'ﾃﾞﾌｫﾙﾄ値が1の為、上記で表示を消しても再表示されてしまう**山選択がある場合のみ1山で25，26の選択
    If KenNo = "T" Then .InspeButt.Visible = False
  Case "2"
    .S = "2-2-S"
    .U = ""
    .C = "2-2-E"
    .D = ""
    .E = "2-1-E"
    .FS = "2-2-S"
    .FU = ""
    .FC = "2-2-E"
    .FD = ""
    .FE = "2-1-E"
    .OPS1.Caption = "2-2-S"
    .OPS2.Caption = ""
    .OPS3.Caption = "2-2-E"
    .OPS4.Caption = ""
    .OPS5.Caption = "2-1-E"
    .OPE1.Caption = "2-2-S"
    .OPE2.Caption = ""
    .OPE3.Caption = "2-2-E"
    .OPE4.Caption = ""
    .OPE5.Caption = "2-1-E"
    .U = ""
    .FU = ""
    .D = ""
    .FD = ""
    
    If KenNo <> "T" Then .AtButt.Visible = True
    
'    If FS4flag And HIFlag = False Then .AtButt.Visible = True
    .InspeButt.Visible = True
    If KenNo = "T" Then .InspeButt.Visible = False
  Case "3"
      .S = "3-3-S"
      .U = "3-3-E"
      .C = "3-2-2"
      .D = "3-1-S"
      .E = "3-1-E"
      .FS = "3-3-S"
      .FU = "3-3-E"
      .FC = "3-2-2"
      .FD = "3-1-S"
      .FE = "3-1-E"
      .OPS1.Caption = "3-3-S"
      .OPS2.Caption = "3-3-E"
      .OPS3.Caption = "3-2-2"
      .OPS4.Caption = "3-1-S"
      .OPS5.Caption = "3-1-E"
      .OPE1.Caption = "3-3-S"
      .OPE2.Caption = "3-3-E"
      .OPE3.Caption = "3-2-2"
      .OPE4.Caption = "3-1-S"
      .OPE5.Caption = "3-1-E"
      
    If KenNo <> "T" Then .AtButt.Visible = True
      
'    If FS4flag And HIFlag = False Then .AtButt.Visible = True
    .InspeButt.Visible = False  '3山ある場合おそらく26枚以上あるため表示を消す
    If KenNo = "T" Then .InspeButt.Visible = False
  End Select
  If .Yama.Visible = False Then
    .S = "S"
    .U = "U"
    .C = "2"
    .D = "D"
    .E = "E"
    .FS = "S"
    .FU = "U"
    .FC = "2"
    .FD = "D"
    .FE = "E"
    .OPS1.Caption = "S" '2021.09.06
    .OPS2.Caption = "U"
    .OPS3.Caption = "2"
    .OPS4.Caption = "D"
    .OPS5.Caption = "E"
    .OPE1.Caption = "S"
    .OPE2.Caption = "U"
    .OPE3.Caption = "2"
    .OPE4.Caption = "D"
    .OPE5.Caption = "E"
  End If
End Select

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>強制描写
'描写
UF1.TSE.Visible = True
UF1.TEE.Visible = True
UF1.E.Visible = True
UF1.NGE.Visible = True

UF1.LC2.Visible = True
UF1.LC8.Visible = True
'少枚数時描写を消すのでリセット用_2024.6.13
For i = 1 To 4
  UF1("S" & i).Visible = True
  UF1("U" & i).Visible = True
  UF1("C" & i).Visible = True
  UF1("D" & i).Visible = True
  UF1("E" & i).Visible = True
  UF1("S" & i).Enabled = True
  UF1("U" & i).Enabled = True
  UF1("C" & i).Enabled = True
  UF1("D" & i).Enabled = True
  UF1("E" & i).Enabled = True
Next i


UF1.FMS.Visible = True '歪
UF1.FMU.Visible = True
UF1.FMC.Visible = True
UF1.FMD.Visible = True
UF1.FME.Visible = True

UF1.L11.Visible = False
UF1.LateralS.Visible = False 'ﾗﾃﾗﾙﾎﾞｰ
UF1.LateralU.Visible = False
UF1.LateralC.Visible = False
UF1.LateralD.Visible = False
UF1.LateralE.Visible = False
'外観
For i = 1 To 5
  UF1("LC" & i).Visible = True
Next i
For i = 7 To 11
  UF1("LC" & i).Visible = True
Next i

UF1.S_UpVisual.Visible = True
UF1.U_UpVisual.Visible = True
UF1.C_UpVisual.Visible = True
UF1.D_UpVisual.Visible = True
UF1.E_UpVisual.Visible = True
UF1.S_UnderVisual.Visible = True
UF1.U_UnderVisual.Visible = True
UF1.C_UnderVisual.Visible = True
UF1.D_UnderVisual.Visible = True
UF1.E_UnderVisual.Visible = True
'BH
UF1.SBH.Visible = True
UF1.UBH.Visible = True
UF1.CBH.Visible = True
UF1.DBH.Visible = True
UF1.EBH.Visible = True
'直角度
If UF1.Youto <> "R115" Then 'QA-S-QA-DIM-154_対角線差及び直角度の測定基準
  UF1.L10.Visible = False
  UF1.DiagSKou.Visible = False
  UF1.DiagUKou.Visible = False
  UF1.DiagCKou.Visible = False
  UF1.DiagDKou.Visible = False
  UF1.DiagEKou.Visible = False
End If

Select Case n
Case "厚板", "薄板", "少枚数" '耳付きで描写するものを非表示
    For i = 7 To 14
      Select Case i
      Case 7, 8, 12, 13
        UF1("L" & i).Visible = False
      End Select
    Next i
    '範囲が被るから分けた
    For i = 7 To 14
      Select Case i
      Case 11, 12, 13, 14
        UF1("S" & i).Visible = False
        UF1("U" & i).Visible = False
        UF1("C" & i).Visible = False
        UF1("D" & i).Visible = False
        UF1("E" & i).Visible = False
        UF1("S" & i).Enabled = False
        UF1("U" & i).Enabled = False
        UF1("C" & i).Enabled = False
        UF1("D" & i).Enabled = False
        UF1("E" & i).Enabled = False
      End Select
    Next i
Case "住精" '2022.11.10******耳を描写後に再選択する場合耳の描写が残るためリセット
    For i = 13 To 14
      Select Case i
      Case 13, 14
        UF1("S" & i).Visible = False
        UF1("U" & i).Visible = False
        UF1("C" & i).Visible = False
        UF1("D" & i).Visible = False
        UF1("E" & i).Visible = False
        UF1("S" & i).Enabled = False
        UF1("U" & i).Enabled = False
        UF1("C" & i).Enabled = False
        UF1("D" & i).Enabled = False
        UF1("E" & i).Enabled = False
      End Select
    Next i
End Select
'2022.06.02_耳付き丈追加あり
Select Case n
Case "極厚"
    For i = 7 To 14
    Select Case i
    Case 7, 8, 12, 13
        UF1("L" & i).Visible = True
    Case 11, 12, 13, 14
        UF1("S" & i).Visible = True
        UF1("U" & i).Visible = True
        UF1("C" & i).Visible = True
        UF1("D" & i).Visible = True
        UF1("E" & i).Visible = True
        UF1("S" & i).Enabled = True
        UF1("U" & i).Enabled = True
        UF1("C" & i).Enabled = True
        UF1("D" & i).Enabled = True
        UF1("E" & i).Enabled = True
    End Select
    Next i
End Select

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>厚
Select Case n
Case "厚板", "少枚数", "極厚"
  For i = 1 To 4 '2021.11.29***薄板描写の厚1,2を消す挙動をリセット
      UF1("AL" & i).Visible = True
      UF1("S" & i).Visible = True
      UF1("U" & i).Visible = True
      UF1("C" & i).Visible = True
      UF1("D" & i).Visible = True
      UF1("E" & i).Visible = True
      UF1("S" & i).Enabled = True
      UF1("U" & i).Enabled = True
      UF1("C" & i).Enabled = True
      UF1("D" & i).Enabled = True
      UF1("E" & i).Enabled = True
  Next i
End Select


If n = "住精" Then GoTo skipSmisei


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>巾丈
Select Case n
Case "厚板"
    For i = 5 To 10 '巾丈3回
        UF1("U" & i).Visible = False
        UF1("D" & i).Visible = False
        UF1("U" & i).Enabled = False
        UF1("D" & i).Enabled = False
        UF1("C" & i).Visible = True 'Cが消されていた時再描写
        UF1("C" & i).Enabled = True
        UF1("E" & i).Visible = True 'Eが消されていた時再描写
        UF1("E" & i).Enabled = True
    Next i
    For i = 1 To 5
      UF1("LC" & i).Visible = True
    Next i
    For i = 7 To 11
      UF1("LC" & i).Visible = True
    Next i
    UF1.U_UpVisual.Visible = True
    UF1.C_UpVisual.Visible = True
    UF1.D_UpVisual.Visible = True
    UF1.U_UnderVisual.Visible = True
    UF1.C_UnderVisual.Visible = True
    UF1.D_UnderVisual.Visible = True
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>薄板専用
Case "薄板" '操業基準追加***Q-S-QA-ALL-163***2021.11.29
  If UF1.Yama.Text = "4" Then
      Call Six_Display("on")
  ElseIf (UF1.Yama.Text = "3") Then '5回*****3山
    For i = 1 To 10
      UF1("S" & i).Visible = True
      UF1("U" & i).Visible = True
      UF1("C" & i).Visible = True
      UF1("D" & i).Visible = True
      UF1("E" & i).Visible = True
      
      UF1("S" & i).Enabled = True
      UF1("U" & i).Enabled = True
      UF1("C" & i).Enabled = True
      UF1("D" & i).Enabled = True
      UF1("E" & i).Enabled = True
    Next i
    
    For i = 1 To 5
      UF1("LC" & i).Visible = True
    Next i
    For i = 7 To 11
      UF1("LC" & i).Visible = True
    Next i
    
    UF1.LC6.Visible = False
    UF1.LC12.Visible = False
    
    Call DistSelect(5) '5回
  Else '3回
    For i = 1 To 10
      UF1("U" & i).Visible = False
      UF1("D" & i).Visible = False
      UF1("U" & i).Enabled = False
      UF1("D" & i).Enabled = False
    Next i
      '外観_2023.5.22
      UF1.LC2.Visible = False
      UF1.LC4.Visible = False
      UF1.LC8.Visible = False
      UF1.LC10.Visible = False
      
      UF1.LC6.Visible = False
      UF1.LC12.Visible = False
        
      UF1.U_UpVisual.Visible = False
      UF1.D_UpVisual.Visible = False
      UF1.U_UnderVisual.Visible = False
      UF1.D_UnderVisual.Visible = False
      
      'BH_2023.5.22
      UF1.UBH.Visible = False
      UF1.DBH.Visible = False
    
    Call DistSelect(26) '歪3回********************************歪の描写_2021.12.01
  End If

'  If LL7Flag Then '5回*****冷延あり　かつ　LL7通過
'  ElseIf (FS4flag) And (UF1.Yama = 3) Then  '5回****冷延あり　かつ　3山
'  ElseIf (FS4flag) And (AtuFlag = False) Then '5回****冷延あり　かつ　6.0mm超え****この条件はAtu_OPselectで弾かれる*2023.03
'  ElseIf (FS4flag) And (UF1.Yama = 1 Or UF1.Yama = 2) Then '3回****冷延あり　かつ　1,2山
'  ElseIf (FS4flag) Then  '3回****冷延あり　かつ　山入力なし
'  ElseIf FS4flag = False Then '5回****冷延なし***条件として最後
'  End If

'2023.6.20_住精の場合ここに一気にスキップしてここだけ見るように修正
skipSmisei:
Case "少枚数"
    For i = 5 To 10
        UF1("S" & i).Visible = True
        UF1("U" & i).Visible = True
        UF1("C" & i).Visible = True
        UF1("D" & i).Visible = True
        UF1("E" & i).Visible = True
        
        UF1("S" & i).Enabled = True
        UF1("U" & i).Enabled = True
        UF1("C" & i).Enabled = True
        UF1("D" & i).Enabled = True
        UF1("E" & i).Enabled = True
    Next i
    For i = 1 To 5
      UF1("LC" & i).Visible = True
    Next i
    For i = 7 To 11
      UF1("LC" & i).Visible = True
    Next i
    UF1.U_UpVisual.Visible = True
    UF1.C_UpVisual.Visible = True
    UF1.D_UpVisual.Visible = True
    UF1.U_UnderVisual.Visible = True
    UF1.C_UnderVisual.Visible = True
    UF1.D_UnderVisual.Visible = True
Case "極厚"
    For i = 5 To 14
        UF1("S" & i).Visible = True
        UF1("U" & i).Visible = True
        UF1("C" & i).Visible = True
        UF1("D" & i).Visible = True
        UF1("E" & i).Visible = True
        
        UF1("S" & i).Enabled = True
        UF1("U" & i).Enabled = True
        UF1("C" & i).Enabled = True
        UF1("D" & i).Enabled = True
        UF1("E" & i).Enabled = True
    Next i
    For i = 1 To 5
      UF1("LC" & i).Visible = True
    Next i
    For i = 7 To 11
      UF1("LC" & i).Visible = True
    Next i
    UF1.U_UpVisual.Visible = True
    UF1.C_UpVisual.Visible = True
    UF1.D_UpVisual.Visible = True
    UF1.U_UnderVisual.Visible = True
    UF1.C_UnderVisual.Visible = True
    UF1.D_UnderVisual.Visible = True
End Select

Select Case n
Case "住精" '2021.07.13
  Select Case True
  Case UF1.厚板
  Call Six_Display("off")
      For i = 5 To 10 '巾丈3回
          UF1("U" & i).Visible = False
          UF1("D" & i).Visible = False
          UF1("U" & i).Enabled = False
          UF1("D" & i).Enabled = False
          UF1("C" & i).Visible = True 'Cが消されていた時再描写
          UF1("C" & i).Enabled = True
          UF1("E" & i).Visible = True 'Eが消されていた時再描写
          UF1("E" & i).Enabled = True
      Next i
      UF1.U_UpVisual.Visible = True
      UF1.C_UpVisual.Visible = True
      UF1.D_UpVisual.Visible = True
      UF1.U_UnderVisual.Visible = True
      UF1.C_UnderVisual.Visible = True
      UF1.D_UnderVisual.Visible = True
      '別の検査に切り替えてもいいように追加
      UF1.DiagS.Visible = True
      UF1.DiagU.Visible = False
      UF1.DiagU.Enabled = False
      UF1.DiagC.Visible = True
      UF1.DiagD.Visible = False
      UF1.DiagD.Enabled = False
      UF1.DiagE.Visible = True
      
      For i = 1 To 5
        UF1("LC" & i).Visible = True
      Next i
      For i = 7 To 11
        UF1("LC" & i).Visible = True
      Next i
      
      UF1.S = "S"
      UF1.U = "S1"
      UF1.C = "2"
      UF1.D = "E1"
      UF1.E = "E"
      UF1.FS = "S"
      UF1.FU = "S1"
      UF1.FC = "2"
      UF1.FD = "E1"
      UF1.FE = "E"
      UF1.OPS1.Caption = "S" '2021.09.06
      UF1.OPS2.Caption = "S1"
      UF1.OPS3.Caption = "2"
      UF1.OPS4.Caption = "E1"
      UF1.OPS5.Caption = "E"
      UF1.OPE1.Caption = "S"
      UF1.OPE2.Caption = "S1"
      UF1.OPE3.Caption = "2"
      UF1.OPE4.Caption = "E1"
      UF1.OPE5.Caption = "E"
      
  Case UF1.薄板 '操業基準追加***Q-S-QA-ALL-163***2021.11.29
    If UF1.Yama.Text = "4" Then
        Call Six_Display("on")
    ElseIf (UF1.Yama.Text = "3") Then '5回*****3山'6回*****4山
      For i = 1 To 10
        UF1("S" & i).Visible = True
        UF1("U" & i).Visible = True
        UF1("C" & i).Visible = True
        UF1("D" & i).Visible = True
        UF1("E" & i).Visible = True
        
        UF1("S" & i).Enabled = True
        UF1("U" & i).Enabled = True
        UF1("C" & i).Enabled = True
        UF1("D" & i).Enabled = True
        UF1("E" & i).Enabled = True
      Next i
      '対角
      UF1.DiagS.Visible = True
      UF1.DiagS.Enabled = True
      UF1.DiagU.Visible = True
      UF1.DiagU.Enabled = True
      UF1.DiagC.Visible = True
      UF1.DiagC.Enabled = True
      UF1.DiagD.Visible = True
      UF1.DiagD.Enabled = True
      UF1.DiagE.Visible = True
      UF1.DiagE.Enabled = True
      
      For i = 1 To 5
        UF1("LC" & i).Visible = True
      Next i
      For i = 7 To 11
        UF1("LC" & i).Visible = True
      Next i
      
      UF1.LC6.Visible = False
      UF1.LC12.Visible = False
      Call DistSelect(5) '5回
    Else '3回
      For i = 1 To 10
        UF1("U" & i).Visible = False
        UF1("D" & i).Visible = False
        UF1("U" & i).Enabled = False
        UF1("D" & i).Enabled = False
      Next i
        '対角
        UF1.DiagU.Visible = False
        UF1.DiagU.Enabled = False
        UF1.DiagD.Visible = False
        UF1.DiagD.Enabled = False
        '外観_2023.5.22
        UF1.LC2.Visible = False
        UF1.LC4.Visible = False
        UF1.LC8.Visible = False
        UF1.LC10.Visible = False
        
        UF1.LC6.Visible = False
        UF1.LC12.Visible = False
          
        UF1.U_UpVisual.Visible = False
        UF1.D_UpVisual.Visible = False
        UF1.U_UnderVisual.Visible = False
        UF1.D_UnderVisual.Visible = False
        'BH_2023.5.22
        UF1.UBH.Visible = False
        UF1.DBH.Visible = False
      
      Call DistSelect(26) '歪3回********************************歪の描写_2021.12.01
      
    End If
    
    Select Case UF1.Yama.Text
    Case "1"
      .S = "S"
      .C = "2"
      .E = "E"
      .FS = "S"
      .FC = "2"
      .FE = "E"
      .OPS1.Caption = "S" '2021.09.06
      .OPS2.Caption = ""
      .OPS3.Caption = "2"
      .OPS4.Caption = ""
      .OPS5.Caption = "E"
      .OPE1.Caption = "S"
      .OPE2.Caption = ""
      .OPE3.Caption = "2"
      .OPE4.Caption = ""
      .OPE5.Caption = "E"
      .U = ""
      .FU = ""
      .D = ""
      .FD = ""
      
      If KenNo <> "T" Then .AtButt.Visible = True '検番T
      
  '    If FS4flag And HIFlag = False Then .AtButt.Visible = True
      If .Yama.Visible Then .InspeButt.Visible = True 'ﾃﾞﾌｫﾙﾄ値が1の為、上記で表示を消しても再表示されてしまう**山選択がある場合のみ1山で25，26の選択
      If KenNo = "T" Then .InspeButt.Visible = False
    Case "2"
      .S = "2-2-S"
      .U = ""
      .C = "2-2-E"
      .D = ""
      .E = "2-1-E"
      .FS = "2-2-S"
      .FU = ""
      .FC = "2-2-E"
      .FD = ""
      .FE = "2-1-E"
      .OPS1.Caption = "2-2-S"
      .OPS2.Caption = ""
      .OPS3.Caption = "2-2-E"
      .OPS4.Caption = ""
      .OPS5.Caption = "2-1-E"
      .OPE1.Caption = "2-2-S"
      .OPE2.Caption = ""
      .OPE3.Caption = "2-2-E"
      .OPE4.Caption = ""
      .OPE5.Caption = "2-1-E"
      .U = ""
      .FU = ""
      .D = ""
      .FD = ""
      
      If KenNo <> "T" Then .AtButt.Visible = True
      
  '    If FS4flag And HIFlag = False Then .AtButt.Visible = True
      .InspeButt.Visible = True
      If KenNo = "T" Then .InspeButt.Visible = False
    Case "3"
        .S = "3-3-S"
        .U = "3-3-E"
        .C = "3-2-2"
        .D = "3-1-S"
        .E = "3-1-E"
        .FS = "3-3-S"
        .FU = "3-3-E"
        .FC = "3-2-2"
        .FD = "3-1-S"
        .FE = "3-1-E"
        .OPS1.Caption = "3-3-S"
        .OPS2.Caption = "3-3-E"
        .OPS3.Caption = "3-2-2"
        .OPS4.Caption = "3-1-S"
        .OPS5.Caption = "3-1-E"
        .OPE1.Caption = "3-3-S"
        .OPE2.Caption = "3-3-E"
        .OPE3.Caption = "3-2-2"
        .OPE4.Caption = "3-1-S"
        .OPE5.Caption = "3-1-E"
        
      If KenNo <> "T" Then .AtButt.Visible = True
        
  '    If FS4flag And HIFlag = False Then .AtButt.Visible = True
      .InspeButt.Visible = False '3山ある場合おそらく26枚以上あるため表示を消す
      If KenNo = "T" Then .InspeButt.Visible = False
    Case "4"
      .Height = 565
      If KenNo <> "T" Then .AtButt.Visible = True
      .InspeButt.Visible = False '3山ある場合おそらく26枚以上あるため表示を消す
      If KenNo = "T" Then .InspeButt.Visible = False
    End Select
  End Select
End Select

Dim SSW_F As Boolean
Select Case True
Case InStr(strCoursForCheck, "SSW") > 0: SSW_F = True
End Select

Select Case n '2020.12.25_追加＊極厚時は定盤検査がないので両方ｹﾞｰｼﾞ表示とする
Case "極厚", "厚板", "薄板", "少枚数", "スラブ"
'Case Else '2021.03.30*修正
'修正時AccReflec_LSelect_Line_AutoSelectionすべて修正要
'  Dim flag(0 To 7) As Boolean
'  If Weight(val(UF1.atu), val(UF1.hab), val(UF1.tak)) > 1300# Then flag(0) = True '1300kg以上_新5点作業不可
'  If val(UF1.tak) < 2500# Then flag(1) = True '丈が2500mm未満_6点作業不可
'  If Weight(val(UF1.atu), val(UF1.hab), val(UF1.tak)) > 2160# Then flag(2) = True  '2160kg以上_6点作業不可
'  If val(UF1.tak) >= 2500# And val(UF1.tak) < 4000# Then flag(3) = True  '2500mm以上かつ丈が4000mm未満_10点作業不可
'  If Weight(val(UF1.atu), val(UF1.hab), val(UF1.tak)) > 2500# Then flag(4) = True   '2500kg以上_10点作業不可
'  If val(UF1.hab) < 950# Then flag(5) = True   '幅が950mm未満_10点作業不可
'  If val(UF1.tak) > 5000 Then flag(6) = True     '丈が5000mm超え_2022.02.16*****5000に変更(定盤サイズ変更)_2023.10.24
'  If UF1.tyu = "F" And SSW_F Then flag(7) = True 'SSW通過かつ75Sで調質F_表面切削なしバキューム不可*****2023.10.26材質条件削除
'  If flag(0) And flag(1) Or _
'  flag(2) And flag(3) Or _
'  flag(4) Or _
'  flag(5) And flag(0) Or _
'  flag(6) Or _
'  flag(7) Then
  '5点で吊れない(重量オーバー)かつ6点も使えない(丈が短い)
  '6点不可(重量オーバー)かつ10点も使えない(丈が短い)
  '10点不可(重量オーバー)
  '6,10点不可(幅が短い)かつ5点で吊れない(重量オーバー)
  '重量を問わず丈が6000mmを超える場合(定盤に乗せる意味がない)
  '表面切削なしバキューム不可***2023.10.05
  
  '2026.09.21_バキューム条件DB化
'************************
'ゲージ判定の条件は品質基準.accdb の「ゲージ判定条件」で管理（LSelect と共通）
  If IsGaugeRequired(val(UF1.ATU), val(UF1.Hab), val(UF1.Tak), CStr(UF1.Zai), CStr(UF1.Tyu), strCoursForCheck) Then
      UF1.LZS.Caption = "幅ｹﾞｰｼﾞ"
      UF1.LZU.Caption = "幅ｹﾞｰｼﾞ"
      UF1.LZC.Caption = "幅ｹﾞｰｼﾞ"
      UF1.LZD.Caption = "幅ｹﾞｰｼﾞ"
      UF1.LZE.Caption = "幅ｹﾞｰｼﾞ"
      UF1.LZF.Caption = "幅ｹﾞｰｼﾞ" '2024.5.28_追加
      
      UF1.LGS.Caption = "丈ｹﾞｰｼﾞ"
      UF1.LGU.Caption = "丈ｹﾞｰｼﾞ"
      UF1.LGC.Caption = "丈ｹﾞｰｼﾞ"
      UF1.LGD.Caption = "丈ｹﾞｰｼﾞ"
      UF1.LGE.Caption = "丈ｹﾞｰｼﾞ"
      UF1.LGF.Caption = "丈ｹﾞｰｼﾞ"
      Call CBAdd(SikaSetAtu)
      UF1.blindfold.Tag = "Set"
    Else
      UF1.LZS.Caption = "定盤"
      UF1.LZU.Caption = "定盤"
      UF1.LZC.Caption = "定盤"
      UF1.LZD.Caption = "定盤"
      UF1.LZE.Caption = "定盤"
      UF1.LZF.Caption = "定盤"
      UF1.LGS.Caption = "ｹﾞｰｼﾞ"
      UF1.LGU.Caption = "ｹﾞｰｼﾞ"
      UF1.LGC.Caption = "ｹﾞｰｼﾞ"
      UF1.LGD.Caption = "ｹﾞｰｼﾞ"
      UF1.LGE.Caption = "ｹﾞｰｼﾞ"
      UF1.LGF.Caption = "ｹﾞｰｼﾞ"
  End If '2021.10.22_修正
  
'If (UF1.ATU) <> "" Then '2021.01.28
'  If (Weight(Val(UF1.ATU), Val(UF1.Hab), Val(UF1.Tak)) > 1300# And Val(UF1.Tak) < 2500#) Or _
'  (Weight(Val(UF1.ATU), Val(UF1.Hab), Val(UF1.Tak)) > 2500#) Then 'ｽﾘﾝｸﾞ検査の時_2021.01.13
'    UF1("LZS").Caption = "ｹﾞｰｼﾞ"
'    UF1("LZU").Caption = "ｹﾞｰｼﾞ"
'    UF1("LZC").Caption = "ｹﾞｰｼﾞ"
'    UF1("LZD").Caption = "ｹﾞｰｼﾞ"
'    UF1("LZE").Caption = "ｹﾞｰｼﾞ"
'    UF1("LGS").Caption = "ｹﾞｰｼﾞ"
'    UF1("LGU").Caption = "ｹﾞｰｼﾞ"
'    UF1("LGC").Caption = "ｹﾞｰｼﾞ"
'    UF1("LGD").Caption = "ｹﾞｰｼﾞ"
'    UF1("LGE").Caption = "ｹﾞｰｼﾞ"
'    Call CBAdd(SikaSetAtu)
'  End If
'End If
  
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
''Q-S-QA-ALL-163_板材抜取検査頻度・板厚測定位置基準**板厚3mm以下かつフラットネス4級の際に測定
''Q-S-QA-DIM-022_フリー反り測定基準_発行日 2024年2月29日_対象がサイズ納入用途に変更_2024.3.5
'2026.03.05_フリー反り条件変更
'いったん戻す　糞品証
'If (UF1.flatnes = 4) And (val(NewAtu) <= 3#) Then '3mm以下ﾌﾗｯﾄﾈｽ等級4級
''  If UF1.HizCome.Tag <> "" Then
'  UF1.flatnesText = "ﾌﾘｰｿﾘ要"
'  UF1.LGS.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGU.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGC.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGD.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGE.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGF.Caption = "ﾌﾘｰｿﾘ"
'  Call CBAdd(SikaSetAtu)
'  UF1.blindfold.Tag = "Set" 'UF1イニシャライズ用フラグ
'End If

End Select

Dim LeftTemp As Long: LeftTemp = 804
Dim LeftTemp2 As Long: LeftTemp2 = 852
Dim LeftOut As Long: LeftOut = 906
Dim LeftOut2 As Long: LeftOut2 = 972
'位置調整 ｻｲｽﾞ
Select Case n
Case "厚板", "薄板", "少枚数"
    UF1.Width = 887.25
    UF1.L9.Left = LeftTemp
'    UF1.L10.Left = LeftTemp2
    UF1("DiagS").Left = LeftTemp
    UF1("DiagU").Left = LeftTemp
    UF1("DiagC").Left = LeftTemp
    UF1("DiagD").Left = LeftTemp
    UF1("DiagE").Left = LeftTemp
    UF1.L7.Left = LeftOut
    UF1.L8.Left = LeftOut2
    For i = 11 To 12
    Select Case i
    Case 11
      UF1("S" & i).Left = LeftOut
      UF1("U" & i).Left = LeftOut
      UF1("C" & i).Left = LeftOut
      UF1("D" & i).Left = LeftOut
      UF1("E" & i).Left = LeftOut
    Case 12
      UF1("S" & i).Left = LeftOut2
      UF1("U" & i).Left = LeftOut2
      UF1("C" & i).Left = LeftOut2
      UF1("D" & i).Left = LeftOut2
      UF1("E" & i).Left = LeftOut2
    End Select
    Next i
Case "極厚"
    UF1.Width = 1044.75 '2024.1.19_変更
    Dim LeftTemp3 As Long: LeftTemp3 = 870
    UF1.L7.Left = LeftTemp
    UF1.L8.Left = LeftTemp3
    '2022.06.02_丈3,4追加
    Dim LeftTemp4 As Long: LeftTemp4 = 936
    Dim LeftTemp5 As Long: LeftTemp5 = 1002
    Dim Top1 As Long:   Top1 = 390
    Dim Top2 As Single: Top2 = 415.25
    Dim Top3 As Long:   Top3 = 440.05
    Dim Top4 As Single: Top4 = 465.7
    Dim Top5 As Single: Top5 = 490.95
    
    UF1.L12.tOP = 366
    UF1.L13.tOP = 366
    UF1.L12.Left = 936
    UF1.L13.Left = 1002
    
    For i = 11 To 14
      Select Case i
      Case 11
        UF1("S" & i).Left = LeftTemp
        UF1("U" & i).Left = LeftTemp
        UF1("C" & i).Left = LeftTemp
        UF1("D" & i).Left = LeftTemp
        UF1("E" & i).Left = LeftTemp
      Case 12
        UF1("S" & i).Left = LeftTemp3
        UF1("U" & i).Left = LeftTemp3
        UF1("C" & i).Left = LeftTemp3
        UF1("D" & i).Left = LeftTemp3
        UF1("E" & i).Left = LeftTemp3
      Case 13
        UF1("S" & i).tOP = Top1
        UF1("U" & i).tOP = Top2
        UF1("C" & i).tOP = Top3
        UF1("D" & i).tOP = Top4
        UF1("E" & i).tOP = Top5
        
        UF1("S" & i).Left = LeftTemp4
        UF1("U" & i).Left = LeftTemp4
        UF1("C" & i).Left = LeftTemp4
        UF1("D" & i).Left = LeftTemp4
        UF1("E" & i).Left = LeftTemp4
      Case 14
        UF1("S" & i).tOP = Top1
        UF1("U" & i).tOP = Top2
        UF1("C" & i).tOP = Top3
        UF1("D" & i).tOP = Top4
        UF1("E" & i).tOP = Top5
      
        UF1("S" & i).Left = LeftTemp5
        UF1("U" & i).Left = LeftTemp5
        UF1("C" & i).Left = LeftTemp5
        UF1("D" & i).Left = LeftTemp5
        UF1("E" & i).Left = LeftTemp5
      End Select
    Next i
    .L9.Visible = False
    UF1("DiagS").Visible = False
    UF1("DiagU").Visible = False
    UF1("DiagC").Visible = False
    UF1("DiagD").Visible = False
    UF1("DiagE").Visible = False
    '対角公差から直角度へ変更_2023.11.24
    .L10.Visible = False
    UF1.DiagSKou.Visible = False
    UF1.DiagUKou.Visible = False
    UF1.DiagCKou.Visible = False
    UF1.DiagDKou.Visible = False
    UF1.DiagEKou.Visible = False
End Select

Select Case n
Case "住精" '2021.07.13
'HistorySaveで個別呼び出し_2024.3.6
'--------------------------------------------------------------------------------------------------------------------2023.8.22ﾗﾃﾗﾙﾎﾞｰ個別対応
'  Call LSelectLateral(KenNo)
''--------------------------------------------------------------------------------------------------------------------2023.11.27直角度表示対応
'  Call RightAngleDisplay(KenNo)
Case "スラブ"
  Call SlabSet
End Select

'対角表示
Select Case n
Case "住精"
  If UF1.Yama.Visible And UF1.Yama.Text = "1" Then
    UF1("DiagS").Visible = True
    UF1("DiagU").Visible = False
    UF1("DiagU").Enabled = False
    UF1("DiagC").Visible = True
    UF1("DiagD").Visible = False
    UF1("DiagD").Enabled = False
    UF1("DiagE").Visible = True
  ElseIf UF1.Yama.Visible And UF1.Yama.Text = "2" Then
    UF1("DiagS").Visible = True
    UF1("DiagU").Visible = False
    UF1("DiagU").Enabled = False
    UF1("DiagC").Visible = True
    UF1("DiagD").Visible = False
    UF1("DiagD").Enabled = False
    UF1("DiagE").Visible = True
  ElseIf UF1.Yama.Visible And UF1.Yama.Text = "3" Then
    UF1("DiagU").Visible = True
    UF1("DiagU").Enabled = True
    UF1("DiagD").Visible = True
    UF1("DiagD").Enabled = True
  End If
Case "厚板"
    UF1.L9.Visible = True
'    UF1.L10.Visible = True
    UF1("DiagS").Visible = True
    UF1("DiagU").Visible = False
    UF1("DiagU").Enabled = False
    UF1("DiagC").Visible = True
    UF1("DiagD").Visible = False
    UF1("DiagD").Enabled = False
    UF1("DiagE").Visible = True
Case "薄板"
  If (UF1.Yama.Text = "3") Then  '5回****　3山
    UF1.L9.Visible = True
'    UF1.L10.Visible = True
    UF1("DiagS").Visible = True
    UF1("DiagU").Visible = True
    UF1("DiagU").Enabled = True
    UF1("DiagC").Visible = True
    UF1("DiagD").Visible = True
    UF1("DiagD").Enabled = True
    UF1("DiagE").Visible = True
  Else '3回
    UF1.L9.Visible = True
'    UF1.L10.Visible = True
    UF1("DiagS").Visible = True
    UF1("DiagU").Visible = False
    UF1("DiagU").Enabled = False
    UF1("DiagC").Visible = True
    UF1("DiagD").Visible = False
    UF1("DiagD").Enabled = False
    UF1("DiagE").Visible = True
  End If

'  If LL7Flag Then '5回*****冷延あり　かつ　LL7通過
'  ElseIf (FS4flag) And (UF1.Yama = 3) Then  '5回****冷延あり　かつ　3山
'  ElseIf (FS4flag) And (AtuFlag = False) Then '5回****冷延あり　かつ　6.0mm超え
'  ElseIf (FS4flag) And (UF1.Yama = 1 Or UF1.Yama = 2) Then '3回****冷延あり　かつ　1,2山
'  ElseIf (FS4flag) Then  '3回****冷延あり　かつ　山入力なし
'  ElseIf FS4flag = False Then '5回****冷延なし***条件として最後
'  End If

Case "少枚数"
    UF1.L9.Visible = True
'    UF1.L10.Visible = True
    UF1("DiagS").Visible = True
    UF1("DiagS").Enabled = True
    UF1("DiagU").Visible = True
    UF1("DiagU").Enabled = True
    UF1("DiagC").Visible = True
    UF1("DiagC").Enabled = True
    UF1("DiagD").Visible = True
    UF1("DiagD").Enabled = True
    UF1("DiagE").Visible = True
    UF1("DiagE").Enabled = True
End Select
End With
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
If bSkipDBCheck Then GoTo skipEnd 'リセット目的の呼び出しはDBチェック不要


Dim TempArry As Variant: TempArry = GetLotInfo("Lot情報履歴", UF1.Lot.Text)
If IsEmpty(TempArry) Then GoTo skipEnd

'2026.10.14_引当混在対策：共通判定（要が1件でもあれば要）
Dim 検番 As String
Dim NeedLateral As Boolean, NeedRight As Boolean, NeedFlat As Boolean
Call JudgeLotAllocations(TempArry, 検番, NeedLateral, NeedRight, NeedFlat)

If NeedLateral Then Call LSelectLateral(検番)
If NeedRight Then Call RightAngleDisplay(検番)
If NeedFlat Then UF1.flatnesText = "ﾌﾘｰｿﾘ要"

'カードに公差がある場合は欄を出す’2024.8.27　家木指示
Dim FlagColor As Boolean: FlagColor = False
If (UF1.LateralText.BackColor = RGB(51, 204, 204)) And (UF1.LateralText.BackStyle = 1) Then FlagColor = True
If FlagColor Then Call LSelectLateral(検番) '------------------------------2024.8.27

skipEnd:
Call frameblindfold("通常")

End Sub

'
'Public Sub LSelect(n As String) '検査選択　ラベル位置　フォームサイズ設定'住精追加_2021.07.13
'Call frameblindfold("目隠し")
'
''耳付きについて丈の下部3,4の追加あり_2022.06.02
''スラブ追加_2022.06.07
'On Error Resume Next
''オーダー厚のセット_2024.5.14
'Dim LotNo As String: LotNo = UF1.Lot.Text
'If SikaSetAtu <> 0 Then
'  SikaAtuAll = SikaAtuCreate(LotNo)
'  If Not IsEmpty(SikaAtuAll) Then SikaSetAtu = SikaAtuAll(LBound(SikaAtuAll))
'End If
'
'If UF1.Byousya Then GoTo LineSkip   '強制描写_2022.06*****LotNo検索に引っかからないと描写できないため
'
''LotNo,コース検索
'Dim HiAcc: HiAcc = Arrbox("SIKALOTNOW.accdb", "板", "Fields", UF1.DBflag.Text) '2022.07_引数追加
'Dim Cours As Long, i As Long, OrdAtu As Long, C_Course  As Long, FS4flag As Boolean: FS4flag = False
'Dim AtuFlag As Boolean: AtuFlag = False
'Dim HIFlag As Boolean:  HIFlag = False
'If Not IsEmpty(HiAcc) Then
'  For i = 1 To UBound(HiAcc, 2)
'    Select Case HiAcc(1, i)
'    Case "設計_設備ｺｰｽ": Cours = i
'    Case "ｵｰﾀﾞｰ板厚":    OrdAtu = i
'    Case "実績_設備ｺｰｽ": C_Course = i
'    End Select
'  Next i
'ElseIf IsEmpty(HiAcc) Then
'  If n = "住精" Then GoTo LineSkip
'End If
'
'Dim ResultT As Boolean: ResultT = IsOver12Hours '2024.9.3
'
'Dim LotAcc: LotAcc = Arrbox("SIKALOTNOW.accdb", UF1.Lot, "Record", UF1.DBflag.Text)
'If IsEmpty(LotAcc) Then
'  If UF1.Load.Caption = "on" Then
'    If ResultT = False Then MsgBox "SIKALOTNOWに該当ロットがありません", vbCritical
'  End If
'  If n = "住精" Then GoTo LineSkip
'End If
'
''<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
''厚板の定義変更_2023.03
''0.001 To 7.999以外が厚板→0.001 To 6.000以外が厚板
''2HI通過チェック追加
''LL7の条件を外す***2023.3.16
''<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'
''2023.4.17_検査No検索
'Dim SIKAHIKAcc: SIKAHIKAcc = Arrbox("SIKAHIKINOW.accdb", "", "Fields", UF1.DBflag.Text) 'SIKAHIKINOW.accdbのフィールドに[ｵｰﾀﾞｰ板丈]がないのでエラーになる
'Dim No As Long
'If Not IsEmpty(SIKAHIKAcc) Then
'  For i = 1 To UBound(SIKAHIKAcc, 2)
'    Select Case SIKAHIKAcc(1, i)
'    Case "検査番号": No = i
'    End Select
'  Next i
'End If
'Dim KenAcc: KenAcc = Arrbox("SIKAHIKINOW.accdb", UF1.Lot, "Record", UF1.DBflag.Text)
'
'If IsEmpty(KenAcc) Then 'SIKALOTNOWでやりなおし
'    SIKAHIKAcc = Arrbox("SIKALOTNOW.accdb", "", "Fields", UF1.DBflag.Text)
'    If Not IsEmpty(SIKAHIKAcc) Then
'      For i = 1 To UBound(SIKAHIKAcc, 2)
'        Select Case SIKAHIKAcc(1, i)
'        Case "検査番号": No = i
'        End Select
'      Next i
'    End If
'    KenAcc = Arrbox("SIKALOTNOW.accdb", UF1.Lot, "Record", UF1.DBflag.Text)
'End If
'
'Dim KenNo As String
'If IsEmpty(KenAcc) = False Then KenNo = Left(KenAcc(1, No), 1)  '検査No'頭1桁
'
'If KenNo = "0" Or KenNo = "" Then
''Q-S-QA-ALL-088_検査番号の付け方に基づき作る
'Dim InputString As String: InputString = LotAcc(1, C_Course)
'Dim Result_C As String:    Result_C = FindPriority(InputString) 'コース最終圧延検索
'  If Result_C <> "" Then KenNo = Result_C
'End If
'
''コース表示
''条件定義
''追加_2023.03
'UF1.Reien.Caption = ""
'Select Case True
'Case InStr(LotAcc(1, Cours), "2HI") > 0
'  HIFlag = True '2HI検索
'  UF1.Reien.Caption = KenNo & "_2HI"
'  UF1.Comment.Caption = "2HI材:厚測定強制4ヵ所"
'End Select
'
'Select Case True
'Case InStr(LotAcc(1, Cours), "L-2") > 0
'  FS4flag = True 'L4HI検索
'
'  If HIFlag Then UF1.Reien.Caption = KenNo & "_L-2" & "_2HI"
'  If HIFlag = False Then UF1.Reien.Caption = KenNo & "_L-2"
'Case InStr(LotAcc(1, Cours), "L-3") > 0
'  FS4flag = True 'L4HI検索
'
'  If HIFlag Then UF1.Reien.Caption = KenNo & "_L-3" & "_2HI"
'  If HIFlag = False Then UF1.Reien.Caption = KenNo & "_L-3"
'End Select
'
'
''LL7対応_2022.01.28
''多分消すはず**
'If FS4flag Then
'  If InStr(LotAcc(1, Cours), "LL7") > 0 Then
'    UF1.Reien.Caption = KenNo & "_LL7"
'    Dim LL7Flag As Boolean: LL7Flag = True
'  End If
'End If
'
'
'Dim NewAtu As Single: NewAtu = LotAcc(1, OrdAtu) 'ｵｰﾀﾞｰ厚参照の控え_2022.01.11
'
'If NewAtu = 0 Then
'  If UF1.Atu <> "" Then
'    NewAtu = val(UF1.Atu)
'  End If
'End If
'If val(NewAtu) <= 6# Then AtuFlag = True '6mm以下でTrue
'
''ﾗﾍﾞﾙﾌｨｰﾙﾄﾞ
'With UF1
'Select Case n
'Case "厚板", "薄板", "少枚数"
'    .L1 = "幅1"
'    .L2 = "幅2"
'    .L3 = "丈1"
'    .L4 = "丈2"
'    .L5 = "対角A"
'    .L6 = "対角B"
'Case "極厚"
'    .L1 = "幅1"
'    .L2 = "幅2"
'    .L3 = "幅3"
'    .L4 = "幅4"
'    .L5 = "幅5"
'    .L6 = "幅6"
'    .L7 = "丈1"
'    .L8 = "丈2"
'    .L12 = "丈3" '2022.06.02
'    .L13 = "丈4"
'End Select
'
''ボタン描写、検査位置描写
'Select Case n
'Case "厚板"
'    .AtCount = ""
'    .AtButt.Visible = False '2回4回
'    .InspeButt.Visible = False '検入枚数25以上以下
'
'    .Yama.Visible = False
'    Call Six_Display("off") 'yamaの表示とセット_2024.5.28
'
'    .YamaLabel.Visible = False
'    .SlectFM.Width = 205
'    .S = "S"
'    .U = "S1"
'    .C = "2"
'    .D = "E1"
'    .E = "E"
'    .FS = "S"
'    .FU = "S1"
'    .FC = "2"
'    .FD = "E1"
'    .FE = "E"
'    .OPS1.Caption = "S" '2021.09.06
'    .OPS2.Caption = "S1"
'    .OPS3.Caption = "2"
'    .OPS4.Caption = "E1"
'    .OPS5.Caption = "E"
'    .OPE1.Caption = "S"
'    .OPE2.Caption = "S1"
'    .OPE3.Caption = "2"
'    .OPE4.Caption = "E1"
'    .OPE5.Caption = "E"
'Case "少枚数", "極厚"
'    .AtCount = ""
'    .AtButt.Visible = False
'    .InspeButt.Visible = False '検入枚数25以上以下
'    .Yama.Visible = False
'    Call Six_Display("off") 'yamaの表示とセット_2024.5.28
'
'    .YamaLabel.Visible = False
'    .SlectFM.Width = 205
'    .S = "1"
'    .U = "2"
'    .C = "3"
'    .D = "4"
'    .E = "5"
'    .FS = "1"
'    .FU = "2"
'    .FC = "3"
'    .FD = "4"
'    .FE = "5"
'    .OPS1.Caption = "1" '2021.09.06
'    .OPS2.Caption = "2"
'    .OPS3.Caption = "3"
'    .OPS4.Caption = "4"
'    .OPS5.Caption = "5"
'    .OPE1.Caption = "1"
'    .OPE2.Caption = "2"
'    .OPE3.Caption = "3"
'    .OPE4.Caption = "4"
'    .OPE5.Caption = "5"
''2023.03_条件見直し
''冷延条件が不要になる…？
'Case "薄板"
''2024.6.7_追加
'    If HIFlag = False Then '2HI通ってない→2か所検査ok　’HIFlag = Trueで2hi通過
'      If AtuFlag = False Then '厚6mm以上あるならば4か所検査(検査番号Tのとき薄板を使用するため)　'AtuFlag6mm以下でTrue
'        .AtButt.Visible = False '2hiを通過していなくても厚が6mm以上あればボタンは表示しない
'      End If
'    End If
'
'    If HIFlag Then .AtButt.Visible = False '2HI通過は厚強制4か所検査　厚選択非表示
'
'    .InspeButt.Visible = True '検入枚数25以上以下
'    .Yama.Visible = True '山数
'    .YamaLabel.Visible = True
'
'    If .Yama.Text = "" Then Call Six_Display("off") 'yamaの表示とセット_2024.5.28
'
'    If KenNo = "T" Then .InspeButt.Visible = False '2HI材は枚数ボタン非表示_2023.4.19
'
''  If LL7Flag Then '5回*****冷延あり　かつ　LL7通過
''  ElseIf (FS4flag) And (AtuFlag = False) Then '5回*****冷延あり　かつ　6mm超え＊＊＊＊＊＊＊＊＊＊＊＊＊この条件はAtu_OPselectで弾かれる*2023.03
''  ElseIf (FS4flag) And (AtuFlag) Then '3回
''  ElseIf (FS4flag) Then '5回*****冷延あり
''  ElseIf FS4flag = False Then '5回*****冷延なし＊＊＊＊＊＊＊＊＊＊＊＊＊この条件はAtu_OPselectで弾かれる*2023.03
''  End If
'
'LineSkip:
'  .SlectFM.Width = 420
'  .AtButt.Caption = "厚2ヵ所"
'  .AtCount = "厚4ヵ所検査中"
'  If KenNo <> "T" Then .AtButt.Visible = True
'  '検査位置ラベル表示
'  Select Case UF1.Yama.Text
'  Case "1"
'    .S = "S"
'    .C = "2"
'    .E = "E"
'    .FS = "S"
'    .FC = "2"
'    .FE = "E"
'    .OPS1.Caption = "S" '2021.09.06
'    .OPS3.Caption = "2"
'    .OPS5.Caption = "E"
'    .OPE1.Caption = "S"
'    .OPE3.Caption = "2"
'    .OPE5.Caption = "E"
'    .U = ""
'    .FU = ""
'    .D = ""
'    .FD = ""
'
'    If KenNo <> "T" Then .AtButt.Visible = True '検番T
'
''    If FS4flag And HIFlag = False Then .AtButt.Visible = True
'    If .Yama.Visible Then .InspeButt.Visible = True 'ﾃﾞﾌｫﾙﾄ値が1の為、上記で表示を消しても再表示されてしまう**山選択がある場合のみ1山で25，26の選択
'    If KenNo = "T" Then .InspeButt.Visible = False
'  Case "2"
'    .S = "2-2-S"
'    .U = ""
'    .C = "2-2-E"
'    .D = ""
'    .E = "2-1-E"
'    .FS = "2-2-S"
'    .FU = ""
'    .FC = "2-2-E"
'    .FD = ""
'    .FE = "2-1-E"
'    .OPS1.Caption = "2-2-S"
'    .OPS2.Caption = ""
'    .OPS3.Caption = "2-2-E"
'    .OPS4.Caption = ""
'    .OPS5.Caption = "2-1-E"
'    .OPE1.Caption = "2-2-S"
'    .OPE2.Caption = ""
'    .OPE3.Caption = "2-2-E"
'    .OPE4.Caption = ""
'    .OPE5.Caption = "2-1-E"
'    .U = ""
'    .FU = ""
'    .D = ""
'    .FD = ""
'
'    If KenNo <> "T" Then .AtButt.Visible = True
'
''    If FS4flag And HIFlag = False Then .AtButt.Visible = True
'    .InspeButt.Visible = True
'    If KenNo = "T" Then .InspeButt.Visible = False
'  Case "3"
'      .S = "3-3-S"
'      .U = "3-3-E"
'      .C = "3-2-2"
'      .D = "3-1-S"
'      .E = "3-1-E"
'      .FS = "3-3-S"
'      .FU = "3-3-E"
'      .FC = "3-2-2"
'      .FD = "3-1-S"
'      .FE = "3-1-E"
'      .OPS1.Caption = "3-3-S"
'      .OPS2.Caption = "3-3-E"
'      .OPS3.Caption = "3-2-2"
'      .OPS4.Caption = "3-1-S"
'      .OPS5.Caption = "3-1-E"
'      .OPE1.Caption = "3-3-S"
'      .OPE2.Caption = "3-3-E"
'      .OPE3.Caption = "3-2-2"
'      .OPE4.Caption = "3-1-S"
'      .OPE5.Caption = "3-1-E"
'
'    If KenNo <> "T" Then .AtButt.Visible = True
'
''    If FS4flag And HIFlag = False Then .AtButt.Visible = True
'    .InspeButt.Visible = False  '3山ある場合おそらく26枚以上あるため表示を消す
'    If KenNo = "T" Then .InspeButt.Visible = False
'  End Select
'  If .Yama.Visible = False Then
'    .S = "S"
'    .U = "U"
'    .C = "2"
'    .D = "D"
'    .E = "E"
'    .FS = "S"
'    .FU = "U"
'    .FC = "2"
'    .FD = "D"
'    .FE = "E"
'    .OPS1.Caption = "S" '2021.09.06
'    .OPS2.Caption = "U"
'    .OPS3.Caption = "2"
'    .OPS4.Caption = "D"
'    .OPS5.Caption = "E"
'    .OPE1.Caption = "S"
'    .OPE2.Caption = "U"
'    .OPE3.Caption = "2"
'    .OPE4.Caption = "D"
'    .OPE5.Caption = "E"
'  End If
'End Select
'
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>強制描写
''描写
'UF1.TSE.Visible = True
'UF1.TEE.Visible = True
'UF1.E.Visible = True
'UF1.NGE.Visible = True
'
'UF1.LC2.Visible = True
'UF1.LC8.Visible = True
''少枚数時描写を消すのでリセット用_2024.6.13
'For i = 1 To 4
'  UF1("S" & i).Visible = True
'  UF1("U" & i).Visible = True
'  UF1("C" & i).Visible = True
'  UF1("D" & i).Visible = True
'  UF1("E" & i).Visible = True
'  UF1("S" & i).Enabled = True
'  UF1("U" & i).Enabled = True
'  UF1("C" & i).Enabled = True
'  UF1("D" & i).Enabled = True
'  UF1("E" & i).Enabled = True
'Next i
'
'
'UF1.FMS.Visible = True '歪
'UF1.FMU.Visible = True
'UF1.FMC.Visible = True
'UF1.FMD.Visible = True
'UF1.FME.Visible = True
'
'UF1.L11.Visible = False
'UF1.LateralS.Visible = False 'ﾗﾃﾗﾙﾎﾞｰ
'UF1.LateralU.Visible = False
'UF1.LateralC.Visible = False
'UF1.LateralD.Visible = False
'UF1.LateralE.Visible = False
''外観
'For i = 1 To 5
'  UF1("LC" & i).Visible = True
'Next i
'For i = 7 To 11
'  UF1("LC" & i).Visible = True
'Next i
'
'UF1.S_UpVisual.Visible = True
'UF1.U_UpVisual.Visible = True
'UF1.C_UpVisual.Visible = True
'UF1.D_UpVisual.Visible = True
'UF1.E_UpVisual.Visible = True
'UF1.S_UnderVisual.Visible = True
'UF1.U_UnderVisual.Visible = True
'UF1.C_UnderVisual.Visible = True
'UF1.D_UnderVisual.Visible = True
'UF1.E_UnderVisual.Visible = True
''BH
'UF1.SBH.Visible = True
'UF1.UBH.Visible = True
'UF1.CBH.Visible = True
'UF1.DBH.Visible = True
'UF1.EBH.Visible = True
''直角度
'If UF1.Youto <> "R115" Then 'QA-S-QA-DIM-154_対角線差及び直角度の測定基準
'  UF1.L10.Visible = False
'  UF1.DiagSKou.Visible = False
'  UF1.DiagUKou.Visible = False
'  UF1.DiagCKou.Visible = False
'  UF1.DiagDKou.Visible = False
'  UF1.DiagEKou.Visible = False
'End If
'
'Select Case n
'Case "厚板", "薄板", "少枚数" '耳付きで描写するものを非表示
'    For i = 7 To 14
'      Select Case i
'      Case 7, 8, 12, 13
'        UF1("L" & i).Visible = False
'      End Select
'    Next i
'    '範囲が被るから分けた
'    For i = 7 To 14
'      Select Case i
'      Case 11, 12, 13, 14
'        UF1("S" & i).Visible = False
'        UF1("U" & i).Visible = False
'        UF1("C" & i).Visible = False
'        UF1("D" & i).Visible = False
'        UF1("E" & i).Visible = False
'        UF1("S" & i).Enabled = False
'        UF1("U" & i).Enabled = False
'        UF1("C" & i).Enabled = False
'        UF1("D" & i).Enabled = False
'        UF1("E" & i).Enabled = False
'      End Select
'    Next i
'Case "住精" '2022.11.10******耳を描写後に再選択する場合耳の描写が残るためリセット
'    For i = 13 To 14
'      Select Case i
'      Case 13, 14
'        UF1("S" & i).Visible = False
'        UF1("U" & i).Visible = False
'        UF1("C" & i).Visible = False
'        UF1("D" & i).Visible = False
'        UF1("E" & i).Visible = False
'        UF1("S" & i).Enabled = False
'        UF1("U" & i).Enabled = False
'        UF1("C" & i).Enabled = False
'        UF1("D" & i).Enabled = False
'        UF1("E" & i).Enabled = False
'      End Select
'    Next i
'End Select
''2022.06.02_耳付き丈追加あり
'Select Case n
'Case "極厚"
'    For i = 7 To 14
'    Select Case i
'    Case 7, 8, 12, 13
'        UF1("L" & i).Visible = True
'    Case 11, 12, 13, 14
'        UF1("S" & i).Visible = True
'        UF1("U" & i).Visible = True
'        UF1("C" & i).Visible = True
'        UF1("D" & i).Visible = True
'        UF1("E" & i).Visible = True
'        UF1("S" & i).Enabled = True
'        UF1("U" & i).Enabled = True
'        UF1("C" & i).Enabled = True
'        UF1("D" & i).Enabled = True
'        UF1("E" & i).Enabled = True
'    End Select
'    Next i
'End Select
'
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>厚
'Select Case n
'Case "厚板", "少枚数", "極厚"
'  For i = 1 To 4 '2021.11.29***薄板描写の厚1,2を消す挙動をリセット
'      UF1("AL" & i).Visible = True
'      UF1("S" & i).Visible = True
'      UF1("U" & i).Visible = True
'      UF1("C" & i).Visible = True
'      UF1("D" & i).Visible = True
'      UF1("E" & i).Visible = True
'      UF1("S" & i).Enabled = True
'      UF1("U" & i).Enabled = True
'      UF1("C" & i).Enabled = True
'      UF1("D" & i).Enabled = True
'      UF1("E" & i).Enabled = True
'  Next i
'End Select
'
'
'If n = "住精" Then GoTo skipSmisei
'
'
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>巾丈
'Select Case n
'Case "厚板"
'    For i = 5 To 10 '巾丈3回
'        UF1("U" & i).Visible = False
'        UF1("D" & i).Visible = False
'        UF1("U" & i).Enabled = False
'        UF1("D" & i).Enabled = False
'        UF1("C" & i).Visible = True 'Cが消されていた時再描写
'        UF1("C" & i).Enabled = True
'        UF1("E" & i).Visible = True 'Eが消されていた時再描写
'        UF1("E" & i).Enabled = True
'    Next i
'    For i = 1 To 5
'      UF1("LC" & i).Visible = True
'    Next i
'    For i = 7 To 11
'      UF1("LC" & i).Visible = True
'    Next i
'    UF1.U_UpVisual.Visible = True
'    UF1.C_UpVisual.Visible = True
'    UF1.D_UpVisual.Visible = True
'    UF1.U_UnderVisual.Visible = True
'    UF1.C_UnderVisual.Visible = True
'    UF1.D_UnderVisual.Visible = True
'
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>薄板専用
'Case "薄板" '操業基準追加***Q-S-QA-ALL-163***2021.11.29
'  If UF1.Yama.Text = "4" Then
'      Call Six_Display("on")
'  ElseIf (UF1.Yama.Text = "3") Then '5回*****3山
'    For i = 1 To 10
'      UF1("S" & i).Visible = True
'      UF1("U" & i).Visible = True
'      UF1("C" & i).Visible = True
'      UF1("D" & i).Visible = True
'      UF1("E" & i).Visible = True
'
'      UF1("S" & i).Enabled = True
'      UF1("U" & i).Enabled = True
'      UF1("C" & i).Enabled = True
'      UF1("D" & i).Enabled = True
'      UF1("E" & i).Enabled = True
'    Next i
'
'    For i = 1 To 5
'      UF1("LC" & i).Visible = True
'    Next i
'    For i = 7 To 11
'      UF1("LC" & i).Visible = True
'    Next i
'
'    UF1.LC6.Visible = False
'    UF1.LC12.Visible = False
'
'    Call DistSelect(5) '5回
'  Else '3回
'    For i = 1 To 10
'      UF1("U" & i).Visible = False
'      UF1("D" & i).Visible = False
'      UF1("U" & i).Enabled = False
'      UF1("D" & i).Enabled = False
'    Next i
'      '外観_2023.5.22
'      UF1.LC2.Visible = False
'      UF1.LC4.Visible = False
'      UF1.LC8.Visible = False
'      UF1.LC10.Visible = False
'
'      UF1.LC6.Visible = False
'      UF1.LC12.Visible = False
'
'      UF1.U_UpVisual.Visible = False
'      UF1.D_UpVisual.Visible = False
'      UF1.U_UnderVisual.Visible = False
'      UF1.D_UnderVisual.Visible = False
'
'      'BH_2023.5.22
'      UF1.UBH.Visible = False
'      UF1.DBH.Visible = False
'
'    Call DistSelect(26) '歪3回********************************歪の描写_2021.12.01
'  End If
'
''  If LL7Flag Then '5回*****冷延あり　かつ　LL7通過
''  ElseIf (FS4flag) And (UF1.Yama = 3) Then  '5回****冷延あり　かつ　3山
''  ElseIf (FS4flag) And (AtuFlag = False) Then '5回****冷延あり　かつ　6.0mm超え****この条件はAtu_OPselectで弾かれる*2023.03
''  ElseIf (FS4flag) And (UF1.Yama = 1 Or UF1.Yama = 2) Then '3回****冷延あり　かつ　1,2山
''  ElseIf (FS4flag) Then  '3回****冷延あり　かつ　山入力なし
''  ElseIf FS4flag = False Then '5回****冷延なし***条件として最後
''  End If
'
''2023.6.20_住精の場合ここに一気にスキップしてここだけ見るように修正
'skipSmisei:
'Case "少枚数"
'    For i = 5 To 10
'        UF1("S" & i).Visible = True
'        UF1("U" & i).Visible = True
'        UF1("C" & i).Visible = True
'        UF1("D" & i).Visible = True
'        UF1("E" & i).Visible = True
'
'        UF1("S" & i).Enabled = True
'        UF1("U" & i).Enabled = True
'        UF1("C" & i).Enabled = True
'        UF1("D" & i).Enabled = True
'        UF1("E" & i).Enabled = True
'    Next i
'    For i = 1 To 5
'      UF1("LC" & i).Visible = True
'    Next i
'    For i = 7 To 11
'      UF1("LC" & i).Visible = True
'    Next i
'    UF1.U_UpVisual.Visible = True
'    UF1.C_UpVisual.Visible = True
'    UF1.D_UpVisual.Visible = True
'    UF1.U_UnderVisual.Visible = True
'    UF1.C_UnderVisual.Visible = True
'    UF1.D_UnderVisual.Visible = True
'Case "極厚"
'    For i = 5 To 14
'        UF1("S" & i).Visible = True
'        UF1("U" & i).Visible = True
'        UF1("C" & i).Visible = True
'        UF1("D" & i).Visible = True
'        UF1("E" & i).Visible = True
'
'        UF1("S" & i).Enabled = True
'        UF1("U" & i).Enabled = True
'        UF1("C" & i).Enabled = True
'        UF1("D" & i).Enabled = True
'        UF1("E" & i).Enabled = True
'    Next i
'    For i = 1 To 5
'      UF1("LC" & i).Visible = True
'    Next i
'    For i = 7 To 11
'      UF1("LC" & i).Visible = True
'    Next i
'    UF1.U_UpVisual.Visible = True
'    UF1.C_UpVisual.Visible = True
'    UF1.D_UpVisual.Visible = True
'    UF1.U_UnderVisual.Visible = True
'    UF1.C_UnderVisual.Visible = True
'    UF1.D_UnderVisual.Visible = True
'End Select
'
'Select Case n
'Case "住精" '2021.07.13
'  Select Case True
'  Case UF1.厚板
'  Call Six_Display("off")
'      For i = 5 To 10 '巾丈3回
'          UF1("U" & i).Visible = False
'          UF1("D" & i).Visible = False
'          UF1("U" & i).Enabled = False
'          UF1("D" & i).Enabled = False
'          UF1("C" & i).Visible = True 'Cが消されていた時再描写
'          UF1("C" & i).Enabled = True
'          UF1("E" & i).Visible = True 'Eが消されていた時再描写
'          UF1("E" & i).Enabled = True
'      Next i
'      UF1.U_UpVisual.Visible = True
'      UF1.C_UpVisual.Visible = True
'      UF1.D_UpVisual.Visible = True
'      UF1.U_UnderVisual.Visible = True
'      UF1.C_UnderVisual.Visible = True
'      UF1.D_UnderVisual.Visible = True
'      '別の検査に切り替えてもいいように追加
'      UF1.DiagS.Visible = True
'      UF1.DiagU.Visible = False
'      UF1.DiagU.Enabled = False
'      UF1.DiagC.Visible = True
'      UF1.DiagD.Visible = False
'      UF1.DiagD.Enabled = False
'      UF1.DiagE.Visible = True
'
'      For i = 1 To 5
'        UF1("LC" & i).Visible = True
'      Next i
'      For i = 7 To 11
'        UF1("LC" & i).Visible = True
'      Next i
'
'      UF1.S = "S"
'      UF1.U = "S1"
'      UF1.C = "2"
'      UF1.D = "E1"
'      UF1.E = "E"
'      UF1.FS = "S"
'      UF1.FU = "S1"
'      UF1.FC = "2"
'      UF1.FD = "E1"
'      UF1.FE = "E"
'      UF1.OPS1.Caption = "S" '2021.09.06
'      UF1.OPS2.Caption = "S1"
'      UF1.OPS3.Caption = "2"
'      UF1.OPS4.Caption = "E1"
'      UF1.OPS5.Caption = "E"
'      UF1.OPE1.Caption = "S"
'      UF1.OPE2.Caption = "S1"
'      UF1.OPE3.Caption = "2"
'      UF1.OPE4.Caption = "E1"
'      UF1.OPE5.Caption = "E"
'
'  Case UF1.薄板 '操業基準追加***Q-S-QA-ALL-163***2021.11.29
'    If UF1.Yama.Text = "4" Then
'        Call Six_Display("on")
'    ElseIf (UF1.Yama.Text = "3") Then '5回*****3山'6回*****4山
'      For i = 1 To 10
'        UF1("S" & i).Visible = True
'        UF1("U" & i).Visible = True
'        UF1("C" & i).Visible = True
'        UF1("D" & i).Visible = True
'        UF1("E" & i).Visible = True
'
'        UF1("S" & i).Enabled = True
'        UF1("U" & i).Enabled = True
'        UF1("C" & i).Enabled = True
'        UF1("D" & i).Enabled = True
'        UF1("E" & i).Enabled = True
'      Next i
'      '対角
'      UF1.DiagS.Visible = True
'      UF1.DiagS.Enabled = True
'      UF1.DiagU.Visible = True
'      UF1.DiagU.Enabled = True
'      UF1.DiagC.Visible = True
'      UF1.DiagC.Enabled = True
'      UF1.DiagD.Visible = True
'      UF1.DiagD.Enabled = True
'      UF1.DiagE.Visible = True
'      UF1.DiagE.Enabled = True
'
'      For i = 1 To 5
'        UF1("LC" & i).Visible = True
'      Next i
'      For i = 7 To 11
'        UF1("LC" & i).Visible = True
'      Next i
'
'      UF1.LC6.Visible = False
'      UF1.LC12.Visible = False
'      Call DistSelect(5) '5回
'    Else '3回
'      For i = 1 To 10
'        UF1("U" & i).Visible = False
'        UF1("D" & i).Visible = False
'        UF1("U" & i).Enabled = False
'        UF1("D" & i).Enabled = False
'      Next i
'        '対角
'        UF1.DiagU.Visible = False
'        UF1.DiagU.Enabled = False
'        UF1.DiagD.Visible = False
'        UF1.DiagD.Enabled = False
'        '外観_2023.5.22
'        UF1.LC2.Visible = False
'        UF1.LC4.Visible = False
'        UF1.LC8.Visible = False
'        UF1.LC10.Visible = False
'
'        UF1.LC6.Visible = False
'        UF1.LC12.Visible = False
'
'        UF1.U_UpVisual.Visible = False
'        UF1.D_UpVisual.Visible = False
'        UF1.U_UnderVisual.Visible = False
'        UF1.D_UnderVisual.Visible = False
'        'BH_2023.5.22
'        UF1.UBH.Visible = False
'        UF1.DBH.Visible = False
'
'      Call DistSelect(26) '歪3回********************************歪の描写_2021.12.01
'
'    End If
'
'    Select Case UF1.Yama.Text
'    Case "1"
'      .S = "S"
'      .C = "2"
'      .E = "E"
'      .FS = "S"
'      .FC = "2"
'      .FE = "E"
'      .OPS1.Caption = "S" '2021.09.06
'      .OPS2.Caption = ""
'      .OPS3.Caption = "2"
'      .OPS4.Caption = ""
'      .OPS5.Caption = "E"
'      .OPE1.Caption = "S"
'      .OPE2.Caption = ""
'      .OPE3.Caption = "2"
'      .OPE4.Caption = ""
'      .OPE5.Caption = "E"
'      .U = ""
'      .FU = ""
'      .D = ""
'      .FD = ""
'
'      If KenNo <> "T" Then .AtButt.Visible = True '検番T
'
'  '    If FS4flag And HIFlag = False Then .AtButt.Visible = True
'      If .Yama.Visible Then .InspeButt.Visible = True 'ﾃﾞﾌｫﾙﾄ値が1の為、上記で表示を消しても再表示されてしまう**山選択がある場合のみ1山で25，26の選択
'      If KenNo = "T" Then .InspeButt.Visible = False
'    Case "2"
'      .S = "2-2-S"
'      .U = ""
'      .C = "2-2-E"
'      .D = ""
'      .E = "2-1-E"
'      .FS = "2-2-S"
'      .FU = ""
'      .FC = "2-2-E"
'      .FD = ""
'      .FE = "2-1-E"
'      .OPS1.Caption = "2-2-S"
'      .OPS2.Caption = ""
'      .OPS3.Caption = "2-2-E"
'      .OPS4.Caption = ""
'      .OPS5.Caption = "2-1-E"
'      .OPE1.Caption = "2-2-S"
'      .OPE2.Caption = ""
'      .OPE3.Caption = "2-2-E"
'      .OPE4.Caption = ""
'      .OPE5.Caption = "2-1-E"
'      .U = ""
'      .FU = ""
'      .D = ""
'      .FD = ""
'
'      If KenNo <> "T" Then .AtButt.Visible = True
'
'  '    If FS4flag And HIFlag = False Then .AtButt.Visible = True
'      .InspeButt.Visible = True
'      If KenNo = "T" Then .InspeButt.Visible = False
'    Case "3"
'        .S = "3-3-S"
'        .U = "3-3-E"
'        .C = "3-2-2"
'        .D = "3-1-S"
'        .E = "3-1-E"
'        .FS = "3-3-S"
'        .FU = "3-3-E"
'        .FC = "3-2-2"
'        .FD = "3-1-S"
'        .FE = "3-1-E"
'        .OPS1.Caption = "3-3-S"
'        .OPS2.Caption = "3-3-E"
'        .OPS3.Caption = "3-2-2"
'        .OPS4.Caption = "3-1-S"
'        .OPS5.Caption = "3-1-E"
'        .OPE1.Caption = "3-3-S"
'        .OPE2.Caption = "3-3-E"
'        .OPE3.Caption = "3-2-2"
'        .OPE4.Caption = "3-1-S"
'        .OPE5.Caption = "3-1-E"
'
'      If KenNo <> "T" Then .AtButt.Visible = True
'
'  '    If FS4flag And HIFlag = False Then .AtButt.Visible = True
'      .InspeButt.Visible = False '3山ある場合おそらく26枚以上あるため表示を消す
'      If KenNo = "T" Then .InspeButt.Visible = False
'    Case "4"
'      .Height = 565
'      If KenNo <> "T" Then .AtButt.Visible = True
'      .InspeButt.Visible = False '3山ある場合おそらく26枚以上あるため表示を消す
'      If KenNo = "T" Then .InspeButt.Visible = False
'    End Select
'  End Select
'End Select
'
'Dim SSW_F As Boolean
'Select Case True
'Case InStr(LotAcc(1, Cours), "SSW") > 0: SSW_F = True
'End Select
'
'Select Case n '2020.12.25_追加＊極厚時は定盤検査がないので両方ｹﾞｰｼﾞ表示とする
'Case "極厚", "厚板", "薄板", "少枚数", "スラブ"
''Case Else '2021.03.30*修正
''修正時AccReflec_LSelect_Line_AutoSelectionすべて修正要
'  Dim flag(0 To 7) As Boolean
'  If Weight(val(UF1.Atu), val(UF1.Hab), val(UF1.Tak)) > 1300# Then flag(0) = True '1300kg以上_新5点作業不可
'  If val(UF1.Tak) < 2500# Then flag(1) = True '丈が2500mm未満_6点作業不可
'  If Weight(val(UF1.Atu), val(UF1.Hab), val(UF1.Tak)) > 2160# Then flag(2) = True  '2160kg以上_6点作業不可
'  If val(UF1.Tak) >= 2500# And val(UF1.Tak) < 4000# Then flag(3) = True  '2500mm以上かつ丈が4000mm未満_10点作業不可
'  If Weight(val(UF1.Atu), val(UF1.Hab), val(UF1.Tak)) > 2500# Then flag(4) = True   '2500kg以上_10点作業不可
'  If val(UF1.Hab) < 950# Then flag(5) = True   '幅が950mm未満_10点作業不可
'  If val(UF1.Tak) > 5000 Then flag(6) = True     '丈が5000mm超え_2022.02.16*****5000に変更(定盤サイズ変更)_2023.10.24
'  If UF1.Tyu = "F" And SSW_F Then flag(7) = True 'SSW通過かつ75Sで調質F_表面切削なしバキューム不可*****2023.10.26材質条件削除
'
'  If flag(0) And flag(1) Or _
'  flag(2) And flag(3) Or _
'  flag(4) Or _
'  flag(5) And flag(0) Or _
'  flag(6) Or _
'  flag(7) Then
'  '5点で吊れない(重量オーバー)かつ6点も使えない(丈が短い)
'  '6点不可(重量オーバー)かつ10点も使えない(丈が短い)
'  '10点不可(重量オーバー)
'  '6,10点不可(幅が短い)かつ5点で吊れない(重量オーバー)
'  '重量を問わず丈が6000mmを超える場合(定盤に乗せる意味がない)
'  '表面切削なしバキューム不可***2023.10.05
'      UF1.LZS.Caption = "幅ｹﾞｰｼﾞ"
'      UF1.LZU.Caption = "幅ｹﾞｰｼﾞ"
'      UF1.LZC.Caption = "幅ｹﾞｰｼﾞ"
'      UF1.LZD.Caption = "幅ｹﾞｰｼﾞ"
'      UF1.LZE.Caption = "幅ｹﾞｰｼﾞ"
'      UF1.LZF.Caption = "幅ｹﾞｰｼﾞ" '2024.5.28_追加
'
'      UF1.LGS.Caption = "丈ｹﾞｰｼﾞ"
'      UF1.LGU.Caption = "丈ｹﾞｰｼﾞ"
'      UF1.LGC.Caption = "丈ｹﾞｰｼﾞ"
'      UF1.LGD.Caption = "丈ｹﾞｰｼﾞ"
'      UF1.LGE.Caption = "丈ｹﾞｰｼﾞ"
'      UF1.LGF.Caption = "丈ｹﾞｰｼﾞ"
'      Call CBAdd(SikaSetAtu)
'      UF1.blindfold.Tag = "Set"
'    Else
'      UF1.LZS.Caption = "定盤"
'      UF1.LZU.Caption = "定盤"
'      UF1.LZC.Caption = "定盤"
'      UF1.LZD.Caption = "定盤"
'      UF1.LZE.Caption = "定盤"
'      UF1.LZF.Caption = "定盤"
'      UF1.LGS.Caption = "ｹﾞｰｼﾞ"
'      UF1.LGU.Caption = "ｹﾞｰｼﾞ"
'      UF1.LGC.Caption = "ｹﾞｰｼﾞ"
'      UF1.LGD.Caption = "ｹﾞｰｼﾞ"
'      UF1.LGE.Caption = "ｹﾞｰｼﾞ"
'      UF1.LGF.Caption = "ｹﾞｰｼﾞ"
'  End If '2021.10.22_修正
'
''If (UF1.ATU) <> "" Then '2021.01.28
''  If (Weight(Val(UF1.ATU), Val(UF1.Hab), Val(UF1.Tak)) > 1300# And Val(UF1.Tak) < 2500#) Or _
''  (Weight(Val(UF1.ATU), Val(UF1.Hab), Val(UF1.Tak)) > 2500#) Then 'ｽﾘﾝｸﾞ検査の時_2021.01.13
''    UF1("LZS").Caption = "ｹﾞｰｼﾞ"
''    UF1("LZU").Caption = "ｹﾞｰｼﾞ"
''    UF1("LZC").Caption = "ｹﾞｰｼﾞ"
''    UF1("LZD").Caption = "ｹﾞｰｼﾞ"
''    UF1("LZE").Caption = "ｹﾞｰｼﾞ"
''    UF1("LGS").Caption = "ｹﾞｰｼﾞ"
''    UF1("LGU").Caption = "ｹﾞｰｼﾞ"
''    UF1("LGC").Caption = "ｹﾞｰｼﾞ"
''    UF1("LGD").Caption = "ｹﾞｰｼﾞ"
''    UF1("LGE").Caption = "ｹﾞｰｼﾞ"
''    Call CBAdd(SikaSetAtu)
''  End If
''End If
'
'''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''Q-S-QA-ALL-163_板材抜取検査頻度・板厚測定位置基準**板厚3mm以下かつフラットネス4級の際に測定
'''Q-S-QA-DIM-022_フリー反り測定基準_発行日 2024年2月29日_対象がサイズ納入用途に変更_2024.3.5
''2026.03.05_フリー反り条件変更
''いったん戻す　糞品証
'If (UF1.flatnes = 4) And (val(NewAtu) <= 3#) Then '3mm以下ﾌﾗｯﾄﾈｽ等級4級
''  If UF1.HizCome.Tag <> "" Then
'  UF1.flatnesText = "ﾌﾘｰｿﾘ要"
'  UF1.LGS.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGU.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGC.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGD.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGE.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGF.Caption = "ﾌﾘｰｿﾘ"
'  Call CBAdd(SikaSetAtu)
'  UF1.blindfold.Tag = "Set" 'UF1イニシャライズ用フラグ
'End If
'
'End Select
'
'Dim LeftTemp As Long: LeftTemp = 804
'Dim LeftTemp2 As Long: LeftTemp2 = 852
'Dim LeftOut As Long: LeftOut = 906
'Dim LeftOut2 As Long: LeftOut2 = 972
''位置調整 ｻｲｽﾞ
'Select Case n
'Case "厚板", "薄板", "少枚数"
'    UF1.Width = 887.25
'    UF1.L9.Left = LeftTemp
''    UF1.L10.Left = LeftTemp2
'    UF1("DiagS").Left = LeftTemp
'    UF1("DiagU").Left = LeftTemp
'    UF1("DiagC").Left = LeftTemp
'    UF1("DiagD").Left = LeftTemp
'    UF1("DiagE").Left = LeftTemp
'    UF1.L7.Left = LeftOut
'    UF1.L8.Left = LeftOut2
'    For i = 11 To 12
'    Select Case i
'    Case 11
'      UF1("S" & i).Left = LeftOut
'      UF1("U" & i).Left = LeftOut
'      UF1("C" & i).Left = LeftOut
'      UF1("D" & i).Left = LeftOut
'      UF1("E" & i).Left = LeftOut
'    Case 12
'      UF1("S" & i).Left = LeftOut2
'      UF1("U" & i).Left = LeftOut2
'      UF1("C" & i).Left = LeftOut2
'      UF1("D" & i).Left = LeftOut2
'      UF1("E" & i).Left = LeftOut2
'    End Select
'    Next i
'Case "極厚"
'    UF1.Width = 1044.75 '2024.1.19_変更
'    Dim LeftTemp3 As Long: LeftTemp3 = 870
'    UF1.L7.Left = LeftTemp
'    UF1.L8.Left = LeftTemp3
'    '2022.06.02_丈3,4追加
'    Dim LeftTemp4 As Long: LeftTemp4 = 936
'    Dim LeftTemp5 As Long: LeftTemp5 = 1002
'    Dim Top1 As Long:   Top1 = 390
'    Dim Top2 As Single: Top2 = 415.25
'    Dim Top3 As Long:   Top3 = 440.05
'    Dim Top4 As Single: Top4 = 465.7
'    Dim Top5 As Single: Top5 = 490.95
'
'    UF1.L12.tOP = 366
'    UF1.L13.tOP = 366
'    UF1.L12.Left = 936
'    UF1.L13.Left = 1002
'
'    For i = 11 To 14
'      Select Case i
'      Case 11
'        UF1("S" & i).Left = LeftTemp
'        UF1("U" & i).Left = LeftTemp
'        UF1("C" & i).Left = LeftTemp
'        UF1("D" & i).Left = LeftTemp
'        UF1("E" & i).Left = LeftTemp
'      Case 12
'        UF1("S" & i).Left = LeftTemp3
'        UF1("U" & i).Left = LeftTemp3
'        UF1("C" & i).Left = LeftTemp3
'        UF1("D" & i).Left = LeftTemp3
'        UF1("E" & i).Left = LeftTemp3
'      Case 13
'        UF1("S" & i).tOP = Top1
'        UF1("U" & i).tOP = Top2
'        UF1("C" & i).tOP = Top3
'        UF1("D" & i).tOP = Top4
'        UF1("E" & i).tOP = Top5
'
'        UF1("S" & i).Left = LeftTemp4
'        UF1("U" & i).Left = LeftTemp4
'        UF1("C" & i).Left = LeftTemp4
'        UF1("D" & i).Left = LeftTemp4
'        UF1("E" & i).Left = LeftTemp4
'      Case 14
'        UF1("S" & i).tOP = Top1
'        UF1("U" & i).tOP = Top2
'        UF1("C" & i).tOP = Top3
'        UF1("D" & i).tOP = Top4
'        UF1("E" & i).tOP = Top5
'
'        UF1("S" & i).Left = LeftTemp5
'        UF1("U" & i).Left = LeftTemp5
'        UF1("C" & i).Left = LeftTemp5
'        UF1("D" & i).Left = LeftTemp5
'        UF1("E" & i).Left = LeftTemp5
'      End Select
'    Next i
'    .L9.Visible = False
'    UF1("DiagS").Visible = False
'    UF1("DiagU").Visible = False
'    UF1("DiagC").Visible = False
'    UF1("DiagD").Visible = False
'    UF1("DiagE").Visible = False
'    '対角公差から直角度へ変更_2023.11.24
'    .L10.Visible = False
'    UF1.DiagSKou.Visible = False
'    UF1.DiagUKou.Visible = False
'    UF1.DiagCKou.Visible = False
'    UF1.DiagDKou.Visible = False
'    UF1.DiagEKou.Visible = False
'End Select
'
'Select Case n
'Case "住精" '2021.07.13
''HistorySaveで個別呼び出し_2024.3.6
''--------------------------------------------------------------------------------------------------------------------2023.8.22ﾗﾃﾗﾙﾎﾞｰ個別対応
''  Call LSelectLateral(KenNo)
'''--------------------------------------------------------------------------------------------------------------------2023.11.27直角度表示対応
''  Call RightAngleDisplay(KenNo)
'Case "スラブ"
'  Call SlabSet
'End Select
'
''対角表示
'Select Case n
'Case "住精"
'  If UF1.Yama.Visible And UF1.Yama.Text = "1" Then
'    UF1("DiagS").Visible = True
'    UF1("DiagU").Visible = False
'    UF1("DiagU").Enabled = False
'    UF1("DiagC").Visible = True
'    UF1("DiagD").Visible = False
'    UF1("DiagD").Enabled = False
'    UF1("DiagE").Visible = True
'  ElseIf UF1.Yama.Visible And UF1.Yama.Text = "2" Then
'    UF1("DiagS").Visible = True
'    UF1("DiagU").Visible = False
'    UF1("DiagU").Enabled = False
'    UF1("DiagC").Visible = True
'    UF1("DiagD").Visible = False
'    UF1("DiagD").Enabled = False
'    UF1("DiagE").Visible = True
'  ElseIf UF1.Yama.Visible And UF1.Yama.Text = "3" Then
'    UF1("DiagU").Visible = True
'    UF1("DiagU").Enabled = True
'    UF1("DiagD").Visible = True
'    UF1("DiagD").Enabled = True
'  End If
'Case "厚板"
'    UF1.L9.Visible = True
''    UF1.L10.Visible = True
'    UF1("DiagS").Visible = True
'    UF1("DiagU").Visible = False
'    UF1("DiagU").Enabled = False
'    UF1("DiagC").Visible = True
'    UF1("DiagD").Visible = False
'    UF1("DiagD").Enabled = False
'    UF1("DiagE").Visible = True
'Case "薄板"
'  If (UF1.Yama.Text = "3") Then  '5回****　3山
'    UF1.L9.Visible = True
''    UF1.L10.Visible = True
'    UF1("DiagS").Visible = True
'    UF1("DiagU").Visible = True
'    UF1("DiagU").Enabled = True
'    UF1("DiagC").Visible = True
'    UF1("DiagD").Visible = True
'    UF1("DiagD").Enabled = True
'    UF1("DiagE").Visible = True
'  Else '3回
'    UF1.L9.Visible = True
''    UF1.L10.Visible = True
'    UF1("DiagS").Visible = True
'    UF1("DiagU").Visible = False
'    UF1("DiagU").Enabled = False
'    UF1("DiagC").Visible = True
'    UF1("DiagD").Visible = False
'    UF1("DiagD").Enabled = False
'    UF1("DiagE").Visible = True
'  End If
'
''  If LL7Flag Then '5回*****冷延あり　かつ　LL7通過
''  ElseIf (FS4flag) And (UF1.Yama = 3) Then  '5回****冷延あり　かつ　3山
''  ElseIf (FS4flag) And (AtuFlag = False) Then '5回****冷延あり　かつ　6.0mm超え
''  ElseIf (FS4flag) And (UF1.Yama = 1 Or UF1.Yama = 2) Then '3回****冷延あり　かつ　1,2山
''  ElseIf (FS4flag) Then  '3回****冷延あり　かつ　山入力なし
''  ElseIf FS4flag = False Then '5回****冷延なし***条件として最後
''  End If
'
'Case "少枚数"
'    UF1.L9.Visible = True
''    UF1.L10.Visible = True
'    UF1("DiagS").Visible = True
'    UF1("DiagS").Enabled = True
'    UF1("DiagU").Visible = True
'    UF1("DiagU").Enabled = True
'    UF1("DiagC").Visible = True
'    UF1("DiagC").Enabled = True
'    UF1("DiagD").Visible = True
'    UF1("DiagD").Enabled = True
'    UF1("DiagE").Visible = True
'    UF1("DiagE").Enabled = True
'End Select
'End With
''------------------------------------------------------------------------------------------------
'Dim TempArry As Variant: TempArry = GetLotInfo("Lot情報履歴", UF1.Lot.Text) 'ヒットしたロットの配列を作成
'If IsEmpty(TempArry) Then
'  GoTo skipEnd
'  Exit Sub
'End If
'Dim TempDelivery As String
'
''納入、取引、用途コード、材質、調質、厚、幅、丈、ロットNo、検査No、ｵｰﾀﾞｰNo、包装仕様NO
'Dim 納入 As Variant, 取引 As Variant, 用途コード As String, 材質 As String, 調質 As String
'Dim 厚 As Single, 幅 As Single, 丈 As Single, 検番 As String
'Dim ArrCount As Long: ArrCount = UBound(TempArry, 1)
'ReDim 納入(1 To ArrCount) As Variant
'ReDim 取引(1 To ArrCount) As Variant  '2024.8.5
'Dim j As Long: j = 0
'  For i = LBound(TempArry, 1) To UBound(TempArry, 1) '引当数分チェック....1回でよくない？ ***あかん！滑川は複数引き当ての時に特殊検査のない引き当てと混ざってくる_2024.4.22
'    j = j + 1
'    納入(j) = TempArry(i, 1)
'    取引(j) = TempArry(i, 2)
'    用途コード = TempArry(i, 3)
'    材質 = TempArry(i, 4)
'    調質 = TempArry(i, 5)
'    厚 = TempArry(i, 6)
'    幅 = TempArry(i, 7)
'    丈 = TempArry(i, 8)
'    検番 = TempArry(i, 10) '2026.05.08_参照がずっとLotNoだったTempArry(i, 9)　検番参照に修正TempArry(i, 10)
'  Next i
'
'検番 = Left(検番, 1)
'
'Dim 納入Check As String, 取引Check As String
'For i = LBound(納入, 1) To UBound(納入, 1)
'  If 納入(i) <> "" Then
'    納入Check = 納入(i)
'    取引Check = 取引(i)
'    TempDelivery = DbReturnItem(納入Check, 取引Check, 用途コード, 材質, 調質, 厚, 幅, 丈)
'
'    If InStr(TempDelivery, "ラテラルボー") > 0 Then Call LSelectLateral(検番)
'    If InStr(TempDelivery, "直角度") > 0 Then Call RightAngleDisplay(検番)
'    If InStr(TempDelivery, "フリー反り") > 0 Then UF1.flatnesText = "ﾌﾘｰｿﾘ要" 'DbReturnItemで直接実行
'  End If
'Next i
'
''カードに公差がある場合は欄を出す’2024.8.27　家木指示
'Dim FlagColor As Boolean: FlagColor = False
'If (UF1.LateralText.BackColor = RGB(51, 204, 204)) And (UF1.LateralText.BackStyle = 1) Then FlagColor = True
'If FlagColor Then Call LSelectLateral(検番) '------------------------------2024.8.27
'
'skipEnd:
'Call frameblindfold("通常")
'
''------------------------------------------------------------------------------------------------
''With Application ' フォームの後ろにExcelウィンドウを隠す
''  .WindowState = xlNormal
''  .Top = UF1.Top
''  .Left = UF1.Left
''  .Width = UF1.Width - 3
''  .Height = UF1.Height - 3
''End With
'End Sub

Public Sub SlabSet() 'スラブ描写用
Dim i As Long
Call Six_Display("off")
With UF1
  For i = 1 To 4
    UF1("AL" & i).Visible = True
    UF1("S" & i).Visible = True
    UF1("U" & i).Visible = False
    UF1("C" & i).Visible = True
    UF1("D" & i).Visible = False
    UF1("E" & i).Visible = False
    UF1("S" & i).Enabled = True
    UF1("U" & i).Enabled = False
    UF1("C" & i).Enabled = True
    UF1("D" & i).Enabled = False
    UF1("E" & i).Enabled = False
  Next i
  
  For i = 5 To 10
    Select Case i
    Case 5 To 8
      UF1("S" & i).Visible = True
      UF1("U" & i).Visible = True
      UF1("C" & i).Visible = True
      UF1("D" & i).Visible = True
      UF1("E" & i).Visible = False
      UF1("S" & i).Enabled = True
      UF1("U" & i).Enabled = True
      UF1("C" & i).Enabled = True
      UF1("D" & i).Enabled = True
      UF1("E" & i).Enabled = False
    Case 9, 10
      UF1("S" & i).Visible = True
      UF1("U" & i).Visible = False
      UF1("C" & i).Visible = True
      UF1("D" & i).Visible = False
      UF1("E" & i).Visible = False
      UF1("S" & i).Enabled = True
      UF1("U" & i).Enabled = False
      UF1("C" & i).Enabled = True
      UF1("D" & i).Enabled = False
      UF1("E" & i).Enabled = False
    End Select
  Next i
  For i = 1 To 5
    UF1("LC" & i).Visible = True
  Next i
  For i = 7 To 11
    UF1("LC" & i).Visible = True
  Next i
  
  For i = 7 To 14
    Select Case i
    Case 7, 8, 12, 13
      UF1("L" & i).Visible = False
    Case 9, 10
      UF1("U" & i).Visible = False
      UF1("D" & i).Visible = False
      UF1("U" & i).Enabled = False
      UF1("D" & i).Enabled = False
    End Select
  Next i
  For i = 11 To 14
    Select Case i
    Case 11, 12, 13, 14
      UF1("S" & i).Visible = False
      UF1("U" & i).Visible = False
      UF1("C" & i).Visible = False
      UF1("D" & i).Visible = False
      UF1("E" & i).Visible = False
      UF1("S" & i).Enabled = False
      UF1("U" & i).Enabled = False
      UF1("C" & i).Enabled = False
      UF1("D" & i).Enabled = False
      UF1("E" & i).Enabled = False
    End Select
  Next i
    
    .L1 = "幅1"
    .L2 = "幅2"
    .L3 = "丈1"
    .L4 = "丈2"
    .L5 = "対角A"
    .L6 = "対角B"
    .AtCount = ""
    .AtButt.Visible = False
    .InspeButt.Visible = False '検入枚数25以上以下
    .Yama.Visible = False
    
    .YamaLabel.Visible = False
    .SlectFM.Width = 205
    .S = "1-上"
    .U = "1-下"
    .C = "2-上"
    .D = "2-下"
    .E = ""
    .FS = "1"
    .FU = ""
    .FC = "2"
    .FD = ""
    .FE = ""
    .OPS1.Caption = "1" '2021.09.06
    .OPS2.Caption = "1"
    .OPS3.Caption = "2"
    .OPS4.Caption = "2"
    .OPS5.Caption = ""
    .OPE1.Caption = "1"
    .OPE2.Caption = "1"
    .OPE3.Caption = "2"
    .OPE4.Caption = "2"
    .OPE5.Caption = ""
    
    .S_UpVisual.Visible = True
    .C_UpVisual.Visible = True
    .S_UnderVisual.Visible = True
    .C_UnderVisual.Visible = True
    
    
    .FMS.Visible = True '歪
    .FMC.Visible = True
    
    .FMU.Visible = False
    .FMD.Visible = False
    .FME.Visible = False
    
    .U_UpVisual.Visible = False
    .U_UnderVisual.Visible = False
    .D_UpVisual.Visible = False
    .D_UnderVisual.Visible = False
    .E_UpVisual.Visible = False
    .E_UnderVisual.Visible = False
    
    .LC2.Visible = False
    .LC8.Visible = False
    For i = 4 To 6
      UF1("LC" & i).Visible = False
      UF1("LC" & i + 6).Visible = False
    Next i
    
    .TSE.Visible = False
    .TEE.Visible = False
    .E.Visible = False
    .NGE.Visible = False
    
    .DiagU.Visible = False
    .DiagU.Enabled = False
    .DiagD.Visible = False
    .DiagD.Enabled = False
    .DiagE.Visible = False
    .DiagE.Enabled = False
End With
End Sub
    
'2026.10.14_引当混在対策：全引当を判定し「要」が1件でもあれば要で確定させる
'LSelect / Line_AutoSelection / AccReflec 共通
Public Sub JudgeLotAllocations(TempArry As Variant, ByRef 検番 As String, _
                               ByRef NeedLateral As Boolean, ByRef NeedRight As Boolean, ByRef NeedFlat As Boolean)
Dim i As Long
Dim 納入Check As String, 取引Check As String, 用途コード As String, 材質 As String, 調質 As String
Dim 厚 As Single, 幅 As Single, 丈 As Single
Dim TempDelivery As String, CheckKey As String
Dim DoneKeys As String: DoneKeys = vbLf
Dim KeepLateralText As String, KeepSumisei As String, KeepHURIS As String
Dim OrgLateralText As String: OrgLateralText = UF1.LateralText  '等級公差など事前セット分
Dim OrgAdditional As String:  OrgAdditional = UF1.Additional.Caption  '技術指示（追記公差）の有無

NeedLateral = False: NeedRight = False: NeedFlat = False
検番 = ""

For i = LBound(TempArry, 1) To UBound(TempArry, 1)
  検番 = TempArry(i, 10)                 '従来どおり最終行の検番
  納入Check = TempArry(i, 1)
  If 納入Check <> "" Then
    '最終行ではなく、その引当行自身の値で判定
    取引Check = TempArry(i, 2)
    用途コード = TempArry(i, 3)
    材質 = TempArry(i, 4)
    調質 = TempArry(i, 5)
    厚 = TempArry(i, 6)
    幅 = TempArry(i, 7)
    丈 = TempArry(i, 8)

    '同じ条件の引当は1回だけ判定
    CheckKey = 納入Check & vbTab & 取引Check & vbTab & 用途コード & vbTab & 材質 & vbTab & _
               調質 & vbTab & 厚 & vbTab & 幅 & vbTab & 丈
    If InStr(DoneKeys, vbLf & CheckKey & vbLf) = 0 Then
      DoneKeys = DoneKeys & CheckKey & vbLf

      TempDelivery = DbReturnItem(納入Check, 取引Check, 用途コード, 材質, 調質, 厚, 幅, 丈)

      Call DebugLog("[引当判定] 納入=" & 納入Check & " / 取引=" & 取引Check & _
                    " / 用途=" & 用途コード & " / 材質=" & 材質 & " / 調質=" & 調質 & _
                    " / 厚=" & 厚 & " / 幅=" & 幅 & " / 丈=" & 丈 & _
                    " / 結果=[" & TempDelivery & "] / Additional=[" & UF1.Additional.Caption & "]")

      If InStr(TempDelivery, "ラテラルボー") > 0 Then
        NeedLateral = True
        KeepLateralText = UF1.LateralText
        KeepSumisei = UF1.Sumisei.Caption
      End If
      If InStr(TempDelivery, "直角度") > 0 Then NeedRight = True
      If InStr(TempDelivery, "フリー反り") > 0 Then
        NeedFlat = True
        KeepHURIS = UF1.HURIS
      End If
    End If
  End If
Next i

検番 = Left(検番, 1)

'---- 集約 ----
If InStr(OrgAdditional, "ﾗﾃﾗﾙﾎﾞ-追記公差あり") > 0 Then
  '技術指示（追記公差）が最優先：滑川の不要判定で消された表示と公差を戻す
  UF1.Additional.Caption = OrgAdditional
  UF1.LateralText = OrgLateralText
  If NeedLateral And KeepSumisei <> "" Then UF1.Sumisei.Caption = KeepSumisei
ElseIf NeedLateral Then
  '1件でも「要」があれば要で確定
  If KeepLateralText = "" Then KeepLateralText = OrgLateralText
  UF1.LateralText = KeepLateralText
  UF1.Sumisei.Caption = KeepSumisei
  '不要表示は出所（どの処理で書かれたか）に関係なく消す
  If InStr(UF1.Additional.Caption, "ﾗﾃﾗﾙﾎﾞｰ検査不要") > 0 Then UF1.Additional.Caption = ""
End If
If NeedFlat Then UF1.HURIS = KeepHURIS

Call DebugLog("[引当判定] 集約結果 NeedLateral=" & NeedLateral & " / NeedRight=" & NeedRight & _
              " / NeedFlat=" & NeedFlat & " / Additional=[" & UF1.Additional.Caption & "]")
End Sub
    
Public Sub LSelectLateral(n As String, Optional Count As String = "") 'ﾗﾃﾗﾙﾎﾞｰ用*2023.8.22*今後の検査頻度、検査の有無が不明確なので個別に設定できるようにしておく
'対象外****2023.12.21_K215  ﾃﾝｼﾞﾖｳﾊﾟﾈﾙｲﾀ  AFX5
'2024.5.15_条件テーブルに≠K215追加
'2024.8.16_新基準書で不要条件となった_Q-S-QA-DIM-179
'If UF1.Youto = "K215" Then
'  UF1.Additional.Caption = "ﾗﾃﾗﾙﾎﾞｰ検査不要_用途K215"
'  Exit Sub '日車であってもこの用途の場合はスルー
'End If

If UF1.極厚 Then
  UF1.Additional.Caption = "ﾗﾃﾗﾙﾎﾞｰ検査不要_耳付き"
  Exit Sub
End If

If InStr(UF1.Additional.Caption, "ﾗﾃﾗﾙﾎﾞｰ検査不要") > 0 Then Exit Sub '2025.6.4_追加

If UF1.ATU = "" Then Exit Sub

'Q-S-QA-DIM-179_板材曲がり(ラテラルボー)検査基準_指定の範囲以外は測定不要(技術)_2024.3.5*****納入指定測定
'Q-S-CH-DIM-163_板材抜取検査頻度.寸法測定位置基準_2025.7*ラテラルボー検査回数変更
Dim ATU As Single:   ATU = UF1.ATU
Dim HABA As Single:  HABA = UF1.Hab
Dim TAKE As Single:  TAKE = UF1.Tak
Dim YoutC As String: YoutC = UF1.Youto
Dim Zai As String:   Zai = UF1.Zai
Dim Tyou As String:  Tyou = UF1.Tyu

'描写
UF1.L11.Visible = True
UF1.LateralS.Visible = True
UF1.LateralU.Visible = True
UF1.LateralC.Visible = True
UF1.LateralD.Visible = True
UF1.LateralE.Visible = True
'位置調整 ｻｲｽﾞ
Dim LeftTemp As Long: LeftTemp = 804
Dim LeftTemp2 As Long: LeftTemp2 = 852
Dim LeftOut As Long: LeftOut = 906
Dim LeftOut2 As Long: LeftOut2 = 972

UF1.Width = 1000
UF1.L11.Left = LeftOut
UF1.LateralS.Left = LeftOut
UF1.LateralU.Left = LeftOut
UF1.LateralC.Left = LeftOut
UF1.LateralD.Left = LeftOut
UF1.LateralE.Left = LeftOut
UF1.LateralF.Left = LeftOut
UF1.L11.tOP = 366
UF1.LateralS.tOP = 390
UF1.LateralU.tOP = 415.25
UF1.LateralC.tOP = 440.05
UF1.LateralD.tOP = 465.7
UF1.LateralE.tOP = 490.95
UF1.LateralF.tOP = 516.1

'2024.5.31_6回検査用
If UF1.Yama.Text = "4" Then UF1.LateralF.Visible = True
If UF1.Yama.Text <> "4" Then UF1.LateralF.Visible = False
If Count <> "" Then UF1.LateralF.Visible = True

Select Case n
Case "H"
  UF1.LateralU.Visible = False
  UF1.LateralD.Visible = False
Case "N", "K", "T", "S", "F"
  Select Case UF1.Yama
  Case 2, 1
    UF1.LateralU.Visible = False
    UF1.LateralD.Visible = False
  End Select
Case "強制"
  UF1.L11.Visible = True
  UF1.LateralS.Visible = True
  UF1.LateralU.Visible = True
  UF1.LateralC.Visible = True
  UF1.LateralD.Visible = True
  UF1.LateralE.Visible = True
Case "新規"
  UF1.L11.Visible = False
  UF1.LateralS.Visible = False
  UF1.LateralU.Visible = False
  UF1.LateralC.Visible = False
  UF1.LateralD.Visible = False
  UF1.LateralE.Visible = False
End Select
End Sub

Public Function nCount() As Boolean '測定値の有無
Dim i As Long, a As Long
For i = 1 To 12
  If UF1("S" & i) <> "" Then
    nCount = True
    Exit For
  End If
  If UF1("U" & i) <> "" Then
    nCount = True
    Exit For
  End If
  If UF1("C" & i) <> "" Then
    nCount = True
    Exit For
  End If
  If UF1("D" & i) <> "" Then
    nCount = True
    Exit For
  End If
  If UF1("E" & i) <> "" Then
    nCount = True
    Exit For
  End If
Next i
For i = 1 To 10
  If UF1("F" & i) <> "" Then
    nCount = True
    Exit For
  End If
Next i
End Function

Public Sub Load_OP() '検査頻度OP選択
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("作業保存" & "_" & Lname) = False Then Exit Sub '-------------------------ｼｰﾄなし
UF1.Tag = "STOP" 'Msgbox停止
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("作業保存" & "_" & Lname)
Dim LastCol As Long: LastCol = UC(sh, 1)
Select Case sh.Cells(1, 20)
Case "厚板":   UF1.厚板 = True
Case "薄板"
 UF1.薄板 = True
 Select Case sh.Cells(1, LastCol)
 Case "厚2ヵ所"
  Call AtCountSelect(2)
 Case "厚4ヵ所"
  Call AtCountSelect(4)
 End Select
Case "少枚数": UF1.少枚数 = True
Case "極厚":   UF1.極厚 = True
End Select
UF1.Tag = "" 'Msgbox再開
End Sub
Public Sub TB_BackColor_Clear(OP As MSForms.OptionButton) '色ﾘｾｯﾄ
Dim DColo As String: DColo = "&H80000005"
Select Case Mid((OP.Name), 3, 1)
Case "S"
  UF1.TSS.BackColor = DColo
  UF1.TSU.BackColor = DColo
  UF1.TSC.BackColor = DColo
  UF1.TSD.BackColor = DColo
  UF1.TSE.BackColor = DColo
  UF1.TSF.BackColor = DColo
Case "E"
  UF1.TES.BackColor = DColo
  UF1.TEU.BackColor = DColo
  UF1.TEC.BackColor = DColo
  UF1.TED.BackColor = DColo
  UF1.TEE.BackColor = DColo
  UF1.TEF.BackColor = DColo
End Select
End Sub
Public Function BackColor_Check(TB As MSForms.Label) As Boolean 'NG色のときtrue Label
'保留は通過させない
Dim Ncol As Long: Ncol = 8421631  'NG RGB(255,128,128)ｺｰﾗﾙ
If UF1(TB.Name).Caption <> "" Then
  If UF1(TB.Name).BackColor = Ncol Then
    If UF1("NG" & Left((TB.Name), 1)).Caption <> "許容" Then BackColor_Check = True
  End If
End If
End Function
Public Function BackColor_CheckTB(TB As MSForms.TextBox) As Boolean 'NG色のときtrue TextBox_2024.4.16
'保留は通過させない
Dim Ncol As Long: Ncol = 8421631  'NG RGB(255,128,128)ｺｰﾗﾙ
Dim Temp As String


'2024.4.17_右から抜いた文字列が数字の場合は不要なので左から抜きなおす　欲しいのはSUCDEF
If Len((TB.Name)) = 2 Then
  Temp = Right((TB.Name), 1) 'ZS GS S1 C1
  If IsNumeric(Temp) Then
    Temp = Left((TB.Name), 1)
  End If
ElseIf Len((TB.Name)) = 3 Then
  Temp = Right((TB.Name), 2) 'ZS GS S1 C1
  If IsNumeric(Temp) Then
    Temp = Left((TB.Name), 1)
  End If
End If


If InStr((TB.Name), "Lateral") > 0 Then Temp = Right((TB.Name), 1) 'LateralS
If InStr((TB.Name), "Diag") > 0 Then Temp = Mid((TB.Name), 5, 1) 'DiagSKou

If UF1(TB.Name) <> "" Then
  If UF1(TB.Name).BackColor = Ncol Then
    If UF1("NG" & Temp).Caption <> "許容" Then BackColor_CheckTB = True
  End If
End If
End Function
Public Function NGrange_Check(n As Long, k As Long) As Boolean '指定範囲に"許容"ﾌﾗｸﾞなしのNGがあるときtrue
NGrange_Check = False
Dim x As Long, i As Long, a As Long
Select Case True '対角は不要
Case UF1.極厚:                       x = 12
Case UF1.厚板, UF1.薄板, UF1.少枚数: x = 8
Case Else:                           x = 8
End Select
For i = 1 To x
  For a = n To k
    Select Case a
    Case 1:  If BackColor_Check(UF1("S" & i)) Then NGrange_Check = True
    Case 2:  If BackColor_Check(UF1("U" & i)) Then NGrange_Check = True
    Case 3:  If BackColor_Check(UF1("C" & i)) Then NGrange_Check = True
    Case 4:  If BackColor_Check(UF1("D" & i)) Then NGrange_Check = True
    Case 5:  If BackColor_Check(UF1("E" & i)) Then NGrange_Check = True
    Case 6:  If BackColor_Check(UF1("F" & i)) Then NGrange_Check = True
    End Select
  Next a
Next i
'別関数使用
'歪A_2024.4.12
For a = n To k
  Select Case a
  Case 1:  If BackColor_CheckTB(UF1("ZS")) Then NGrange_Check = True
  Case 2:  If BackColor_CheckTB(UF1("ZU")) Then NGrange_Check = True
  Case 3:  If BackColor_CheckTB(UF1("ZC")) Then NGrange_Check = True
  Case 4:  If BackColor_CheckTB(UF1("ZD")) Then NGrange_Check = True
  Case 5:  If BackColor_CheckTB(UF1("ZE")) Then NGrange_Check = True
  Case 6:  If BackColor_CheckTB(UF1("ZF")) Then NGrange_Check = True
  End Select
Next a
'歪ｹﾞｰｼﾞ
For a = n To k
  Select Case a
  Case 1:  If BackColor_CheckTB(UF1("GS")) Then NGrange_Check = True
  Case 2:  If BackColor_CheckTB(UF1("GU")) Then NGrange_Check = True
  Case 3:  If BackColor_CheckTB(UF1("GC")) Then NGrange_Check = True
  Case 4:  If BackColor_CheckTB(UF1("GD")) Then NGrange_Check = True
  Case 5:  If BackColor_CheckTB(UF1("GE")) Then NGrange_Check = True
  Case 6:  If BackColor_CheckTB(UF1("GF")) Then NGrange_Check = True
  End Select
Next a

If UF1.Sumisei.Caption <> "" Then '2021.07.20**2022.8.17*****ﾗﾃﾗﾙﾎﾞｰ測定の納入先が追加される為"住友精密"から<>に変更
  For a = n To k
    Select Case a
    Case 1:  If BackColor_CheckTB(UF1("LateralS")) Then NGrange_Check = True
    Case 2:  If BackColor_CheckTB(UF1("LateralU")) Then NGrange_Check = True
    Case 3:  If BackColor_CheckTB(UF1("LateralC")) Then NGrange_Check = True
    Case 4:  If BackColor_CheckTB(UF1("LateralD")) Then NGrange_Check = True
    Case 5:  If BackColor_CheckTB(UF1("LateralE")) Then NGrange_Check = True
    Case 6:  If BackColor_CheckTB(UF1("LateralF")) Then NGrange_Check = True
    End Select
  Next a
End If

If UF1.Sumisei.Caption <> "" And UF1.Youto = "R115" Then '22023.11.27_直角度
  For a = n To k
    Select Case a
    Case 1:  If BackColor_CheckTB(UF1("DiagSKou")) Then NGrange_Check = True
    Case 2:  If BackColor_CheckTB(UF1("DiagUKou")) Then NGrange_Check = True
    Case 3:  If BackColor_CheckTB(UF1("DiagCKou")) Then NGrange_Check = True
    Case 4:  If BackColor_CheckTB(UF1("DiagDKou")) Then NGrange_Check = True
    Case 5:  If BackColor_CheckTB(UF1("DiagEKou")) Then NGrange_Check = True
    Case 6:  If BackColor_CheckTB(UF1("DiagFKou")) Then NGrange_Check = True
    End Select
  Next a
End If
End Function
Public Sub SetSpecificTextBoxesBackColorToWhite() '特定のテキストボックスのみを対象として、バックカラーを白に変更
Dim ctrl As Control
Dim TargetTextBoxes As Variant
TargetTextBoxes = Array("ZS", "ZU", "ZC", "ZD", "ZE", "ZF", "GS", "GU", "GC", "GD", "GE", "GF")

For Each ctrl In UF1.Controls
  If TypeOf ctrl Is MSForms.TextBox Then
    If IsInArray(ctrl.Name, TargetTextBoxes) Then
      If ctrl.BackColor <> RGB(255, 255, 255) Then ' RGB(255, 255, 255) は白色を表す
        ctrl.BackColor = RGB(255, 255, 255)
      End If
    End If
  End If
Next ctrl
End Sub
Function IsInArray(ByVal Value As Variant, ByVal arr As Variant) As Boolean '指定された値が配列内に存在するかどうかを確認
Dim element As Variant
For Each element In arr
  If Value = element Then
    IsInArray = True
    Exit Function
  End If
Next element
IsInArray = False
End Function

Public Sub RemoveUnderlineFromLabels() 'アンダーラインが設定されている場合はずす
Dim ctrl As Control
Dim LabelPrefixes As Variant
Dim prefix As Variant
Dim i As Long
' ラベルの接頭辞
LabelPrefixes = Array("S", "U", "C", "D", "E")
For Each prefix In LabelPrefixes
  For i = 1 To 14
    Set ctrl = UF1.Controls(prefix & i)
    If ctrl.Font.Underline = True Then 'アンダーラインが設定されている場合
      ctrl.Font.Underline = False
    End If
  Next i
Next prefix
'数が違うから分けて処理
LabelPrefixes = Array("F")
For Each prefix In LabelPrefixes
  For i = 1 To 10
    Set ctrl = UF1.Controls(prefix & i)
    If ctrl.Font.Underline = True Then 'アンダーラインが設定されている場合
      ctrl.Font.Underline = False
    End If
  Next i
Next prefix

End Sub

Public Sub MinMax()
'印刷に関しては保留許容ともに可とする(データ保存のため)
'MinMaxでNGrange_Checkにて保留を通過させない→ｱｲｶｽｲﾝﾌﾟｯﾄに行けない
Dim i As Long, a As Long, th As Long
Dim S_flag As Long: S_flag = 0
Dim E_flag As Long: E_flag = 0
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Select Case True
Case UF1("OPS1"): S_flag = 1
Case UF1("OPS2"): S_flag = 2
Case UF1("OPS3"): S_flag = 3
Case UF1("OPS4"): S_flag = 4
Case UF1("OPS5"): S_flag = 5
Case UF1("OPS6"): S_flag = 6
End Select
Select Case True
Case UF1("OPE1"): E_flag = 1
Case UF1("OPE2"): E_flag = 2
Case UF1("OPE3"): E_flag = 3
Case UF1("OPE4"): E_flag = 4
Case UF1("OPE5"): E_flag = 5
Case UF1("OPE6"): E_flag = 6
End Select
If S_flag = 0 Or E_flag = 0 Then Exit Sub 'ﾁｪｯｸがない
If S_flag > E_flag Then Exit Sub 'ｴﾝﾄﾞのほうが小さい
'許容はスルー　保留は止める**2024.4.16
If NGrange_Check(S_flag, E_flag) Then 'NGがある場合表示させない
  Call MsgError("許容でない公差外が存在します", Lname, 1)
  Exit Sub
End If

 '------------------------------厚
Dim Aran: Aran = UF_Array(1, 4, S_flag, E_flag) '列,行
For i = LBound(Aran, 2) To UBound(Aran, 2)
  For a = LBound(Aran, 1) To UBound(Aran, 1)
    If Aran(a, i) <> "" Then Aran(a, i) = val(Aran(a, i)) '空白は空白のまま処理
  Next a
Next i
 '------------------------------巾
Dim HRan
Select Case True
Case UF1.極厚
  HRan = UF_Array(5, 10, S_flag, E_flag)
Case UF1.薄板
  HRan = UF_Array(5, 6, S_flag, E_flag)
Case UF1.厚板
  If S_flag = 2 And E_flag = 2 Then 'U_巾丈測定無しの為S参照_2021.01.20
    HRan = UF_Array(5, 6, 1, 1)
  ElseIf S_flag = 4 And E_flag = 4 Then 'D_巾丈測定無しの為C参照
    HRan = UF_Array(5, 6, 3, 3)
  Else
    HRan = UF_Array(5, 6, S_flag, E_flag)
  End If
Case UF1.少枚数 '1~5枚検査用_2023.9.13
'枚数を確定させる
  If UF1.ATU <> "" Then
    If UF1.KenHBox7 = "" Then
      th = Format(Round(CalculateSheets(UF1, UF1.ATU), 1), "0")
    ElseIf UF1.KenCount <> "" Then
      th = UF1.KenCount
    End If
    If th >= 5 Then th = 5 '少枚数は5枚まで
  End If

  If th = 5 Then '5枚検査*巾丈3回
    If S_flag = 2 And E_flag = 2 Then 'U_巾丈測定無しの為S参照_2021.01.20
      HRan = UF_Array(5, 6, 1, 1)
    ElseIf S_flag = 4 And E_flag = 4 Then 'D_巾丈測定無しの為C参照
      HRan = UF_Array(5, 6, 3, 3)
    Else
      HRan = UF_Array(5, 6, S_flag, E_flag)
    End If
  ElseIf th = 4 Then '4枚検査*巾丈3回
    If S_flag = 2 And E_flag = 2 Then
      HRan = UF_Array(5, 6, 1, 1) '2巾丈測定無しの為1参照
    ElseIf S_flag = 4 And E_flag = 4 Then
      HRan = UF_Array(5, 6, 4, 4)
    Else
      HRan = UF_Array(5, 6, S_flag, E_flag)
    End If
  Else
    HRan = UF_Array(5, 6, S_flag, E_flag)
  End If
Case Else
  HRan = UF_Array(5, 6, S_flag, E_flag)
End Select
For i = LBound(HRan, 2) To UBound(HRan, 2)
  For a = LBound(HRan, 1) To UBound(HRan, 1)
    If HRan(a, i) <> "" Then HRan(a, i) = val(HRan(a, i))
  Next a
Next i
 '------------------------------丈
Dim TRan
Select Case True
Case UF1.極厚
  TRan = UF_Array(11, 14, S_flag, E_flag) '2022.06_耳付き丈3,4追加
Case UF1.薄板
  TRan = UF_Array(7, 8, S_flag, E_flag)
Case UF1.厚板
  If S_flag = 2 And E_flag = 2 Then 'U_巾丈測定無しの為S参照_2021.01.20
    TRan = UF_Array(7, 8, 1, 1)
  ElseIf S_flag = 4 And E_flag = 4 Then 'D_巾丈測定無しの為C参照
    TRan = UF_Array(7, 8, 3, 3)
  Else
    TRan = UF_Array(7, 8, S_flag, E_flag)
  End If
Case UF1.少枚数  '1~5枚検査用_2023.9.13
'枚数を確定させる
  If UF1.ATU <> "" Then
    If UF1.KenHBox7 = "" Then
      th = Format(Round(CalculateSheets(UF1, UF1.ATU), 1), "0")
    ElseIf UF1.KenCount <> "" Then
      th = UF1.KenCount
    End If
    If th >= 5 Then th = 5 '少枚数は5枚まで
  End If
  
  If th = 5 Then '5枚検査*巾丈3回
    If S_flag = 2 And E_flag = 2 Then 'U_巾丈測定無しの為S参照_2021.01.20
      TRan = UF_Array(7, 8, 1, 1)
    ElseIf S_flag = 4 And E_flag = 4 Then 'D_巾丈測定無しの為C参照
      TRan = UF_Array(7, 8, 3, 3)
    Else
      TRan = UF_Array(7, 8, S_flag, E_flag)
    End If
  ElseIf th = 4 Then '4枚検査*巾丈3回
    If S_flag = 2 And E_flag = 2 Then
      TRan = UF_Array(7, 8, 1, 1) '2巾丈測定無しの為1参照
    ElseIf S_flag = 4 And E_flag = 4 Then
      TRan = UF_Array(7, 8, 4, 4)
    Else
      TRan = UF_Array(7, 8, S_flag, E_flag)
    End If
  Else
    TRan = UF_Array(7, 8, S_flag, E_flag)
  End If
Case Else
  TRan = UF_Array(7, 8, S_flag, E_flag)
End Select
For i = LBound(TRan, 2) To UBound(TRan, 2)
  For a = LBound(TRan, 1) To UBound(TRan, 1)
    If TRan(a, i) <> "" Then TRan(a, i) = val(TRan(a, i))
  Next a
Next i
''------------------------------反映_桁
If 0 < val(UF1.ATU) <= 74# Then
  UF1.AMin = Format(Application.WorksheetFunction.Min(Aran), "0.000")
  UF1.AMax = Format(Application.WorksheetFunction.Max(Aran), "0.000")
ElseIf val(UF1.ATU) > 74 Then
  UF1.AMin = Format(Application.WorksheetFunction.Min(Aran), "0.00")
  UF1.AMax = Format(Application.WorksheetFunction.Max(Aran), "0.00")
End If
UF1.HMin = Format(Application.WorksheetFunction.Min(HRan), "0.0")
UF1.HMax = Format(Application.WorksheetFunction.Max(HRan), "0.0")
UF1.TMin = Format(Application.WorksheetFunction.Min(TRan), "0.0")
UF1.TMax = Format(Application.WorksheetFunction.Max(TRan), "0.0")

'2023.9.13_ノギス測定の条件を忘れていたため追加
If UF1.Youto = "H494" And UF1.Hab < 300# Then '用途コードH494に限り幅測定をノギスで行う
  Select Case val(UF1.Hab) '幅寸法が範囲内の時
  Case 1# To 299#
    UF1.HMin = Format(Application.WorksheetFunction.Min(HRan), "0.00")
    UF1.HMax = Format(Application.WorksheetFunction.Max(HRan), "0.00")
  Case Else
    UF1.HMin = Format(Application.WorksheetFunction.Min(HRan), "0.0")
    UF1.HMax = Format(Application.WorksheetFunction.Max(HRan), "0.0")
  End Select
End If

If UF1.AMin = 0 Then UF1.AMin = ""
If UF1.AMax = 0 Then UF1.AMax = ""
If UF1.HMin = 0 Then UF1.HMin = ""
If UF1.HMax = 0 Then UF1.HMax = ""
If UF1.TMin = 0 Then UF1.TMin = ""
If UF1.TMax = 0 Then UF1.TMax = ""

'ﾘｾｯﾄ>>>>>>.Font.Underline = Falseをループで回すと2021バージョンだとすべてTrue処理される....関数を別途作って対応_2024.2.29
Call RemoveUnderlineFromLabels

If UF1.AMin <> "" And UF1.AMax <> "" Then Call UnderBar("A", S_flag, E_flag, UF1.AMin, UF1.AMax) '厚

If UF1.HMin <> "" And UF1.HMax <> "" Then
  Select Case True
  Case UF1.厚板
    If S_flag = 2 And E_flag = 2 Then 'U_巾丈測定無しの為S参照_2021.01.20
      Call UnderBar("H", 1, 1, UF1.HMin, UF1.HMax) '巾
    ElseIf S_flag = 4 And E_flag = 4 Then
      Call UnderBar("H", 3, 3, UF1.HMin, UF1.HMax) '巾
    Else
      Call UnderBar("H", S_flag, E_flag, UF1.HMin, UF1.HMax) '巾
    End If
  Case UF1.少枚数
    If th = 5 Then '5枚検査*巾丈3回
      If S_flag = 2 And E_flag = 2 Then 'U_巾丈測定無しの為S参照_2021.01.20
        Call UnderBar("H", 1, 1, UF1.HMin, UF1.HMax) '巾
      ElseIf S_flag = 4 And E_flag = 4 Then
        Call UnderBar("H", 3, 3, UF1.HMin, UF1.HMax) '巾
      Else
        Call UnderBar("H", S_flag, E_flag, UF1.HMin, UF1.HMax) '巾
      End If
    ElseIf th = 4 Then '4枚検査*巾丈3回
      If S_flag = 2 And E_flag = 2 Then 'U_巾丈測定無しの為S参照_2021.01.20
        Call UnderBar("H", 1, 1, UF1.HMin, UF1.HMax) '巾
      ElseIf S_flag = 4 And E_flag = 4 Then
        Call UnderBar("H", 4, 4, UF1.HMin, UF1.HMax) '巾
      Else
        Call UnderBar("H", S_flag, E_flag, UF1.HMin, UF1.HMax) '巾
      End If
    Else
      Call UnderBar("H", S_flag, E_flag, UF1.HMin, UF1.HMax)
    End If
  Case Else
    Call UnderBar("H", S_flag, E_flag, UF1.HMin, UF1.HMax)
  End Select
End If

If UF1.TMin <> "" And UF1.TMax <> "" Then
  Select Case True
  Case UF1.厚板
    If S_flag = 2 And E_flag = 2 Then 'U_巾丈測定無しの為S参照_2021.01.20
      Call UnderBar("T", 1, 1, UF1.TMin, UF1.TMax) '丈
    ElseIf S_flag = 4 And E_flag = 4 Then
      Call UnderBar("T", 3, 3, UF1.TMin, UF1.TMax) '丈
    Else
      Call UnderBar("T", S_flag, E_flag, UF1.TMin, UF1.TMax) '丈
    End If
  Case UF1.少枚数
    If th = 5 Then '5枚検査*巾丈3回
      If S_flag = 2 And E_flag = 2 Then 'U_巾丈測定無しの為S参照_2021.01.20
        Call UnderBar("T", 1, 1, UF1.TMin, UF1.TMax) '丈
      ElseIf S_flag = 4 And E_flag = 4 Then
        Call UnderBar("T", 3, 3, UF1.TMin, UF1.TMax) '丈
      Else
        Call UnderBar("T", S_flag, E_flag, UF1.TMin, UF1.TMax) '丈
      End If
    ElseIf th = 4 Then '4枚検査*巾丈3回
      If S_flag = 2 And E_flag = 2 Then 'U_巾丈測定無しの為S参照_2021.01.20
        Call UnderBar("T", 1, 1, UF1.TMin, UF1.TMax) '丈
      ElseIf S_flag = 4 And E_flag = 4 Then
        Call UnderBar("T", 4, 4, UF1.TMin, UF1.TMax) '丈
      Else
        Call UnderBar("T", S_flag, E_flag, UF1.TMin, UF1.TMax) '丈
      End If
    Else
      Call UnderBar("T", S_flag, E_flag, UF1.TMin, UF1.TMax)
    End If
  Case Else
    Call UnderBar("T", S_flag, E_flag, UF1.TMin, UF1.TMax) '丈
  End Select
End If
End Sub
Public Sub TB_BackColor(OP As MSForms.OptionButton, n As String) 'OPtrueでTBｶﾗｰ
If Not OP Then Exit Sub 'trueの時だけ実行
Dim SColo As Long: SColo = 26367
Dim EColo As Long: EColo = 13421619
Dim DColo As String: DColo = "&H80000005"
Call TB_BackColor_Clear(OP)

Call OPcount 'SがEより大きいとき停止フラグを出す
If Not UF1.SFM.Tag = "" Then
  UF1.SFM.Tag = ""
  Exit Sub
End If

Select Case n
Case "S"
  Select Case Right((OP.Name), 1)
  Case 1: UF1.TSS.BackColor = SColo
  Case 2: UF1.TSU.BackColor = SColo
  Case 3: UF1.TSC.BackColor = SColo
  Case 4: UF1.TSD.BackColor = SColo
  Case 5: UF1.TSE.BackColor = SColo
  Case 6: UF1.TSF.BackColor = SColo
  End Select
Case "E"
  Select Case Right((OP.Name), 1)
  Case 1: UF1.TES.BackColor = EColo
  Case 2: UF1.TEU.BackColor = EColo
  Case 3: UF1.TEC.BackColor = EColo
  Case 4: UF1.TED.BackColor = EColo
  Case 5: UF1.TEE.BackColor = EColo
  Case 6: UF1.TEF.BackColor = EColo
  End Select
End Select
End Sub
Public Sub OPcount() 'OPSとOPEｽﾀｰﾄｴﾝﾄﾞ
Dim nos As Long, NoE As Long, i As Long
Dim Sfl As Boolean, Efl As Boolean
For i = 1 To 6
 If UF1("OPS" & i) Then
 Sfl = True
 nos = i
 Exit For
 End If
Next i
For i = 1 To 6
 If UF1("OPE" & i) Then
 Efl = True
 NoE = i
 Exit For
 End If
Next i
If Sfl And Efl Then If nos > NoE Then UF1.SFM.Tag = "Stop"
End Sub
Public Sub UnderBar(buf As String, n As Long, k As Long, Min As Double, Max As Double) 'ｽﾀｰﾄOP、ｴﾝﾄﾞOP、min値、max値
Dim S As String
Dim E As String
Dim i As Long, a As Long, x As Long, y As Long

Select Case buf
Case "A"
  x = 1
  y = 4
Case "H"
  Select Case True
  Case UF1.極厚
    x = 5
    y = 10
  Case UF1.厚板, UF1.薄板, UF1.少枚数
    x = 5
    y = 6
  Case Else
    x = 5
    y = 6
  End Select
Case "T"
  Select Case True
  Case UF1.極厚
    x = 11
    y = 14 '2022.06_耳付き丈3,4追加
  Case UF1.厚板, UF1.薄板, UF1.少枚数
    x = 7
    y = 8
  Case Else
    x = 7
    y = 8
  End Select
End Select

Select Case n
Case 1: S = "S"
Case 2: S = "U"
Case 3: S = "C"
Case 4: S = "D"
Case 5: S = "E"
Case 6: S = "F"
End Select
Select Case k
Case 1: E = "S"
Case 2: E = "U"
Case 3: E = "C"
Case 4: E = "D"
Case 5: E = "E"
Case 6: E = "F"
End Select

For i = x To y
  Select Case S
  Case "S"
    Select Case E
    Case "S"
      If UF1("S" & i).Caption = Min Then UF1("S" & i).Font.Underline = True
      If UF1("S" & i).Caption = Max Then UF1("S" & i).Font.Underline = True
    Case "U"
      If UF1("S" & i).Caption = Min Then UF1("S" & i).Font.Underline = True
      If UF1("U" & i).Caption = Min Then UF1("U" & i).Font.Underline = True
      If UF1("S" & i).Caption = Max Then UF1("S" & i).Font.Underline = True
      If UF1("U" & i).Caption = Max Then UF1("U" & i).Font.Underline = True
    Case "C"
      If UF1("S" & i).Caption = Min Then UF1("S" & i).Font.Underline = True
      If UF1("U" & i).Caption = Min Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Min Then UF1("C" & i).Font.Underline = True
      If UF1("S" & i).Caption = Max Then UF1("S" & i).Font.Underline = True
      If UF1("U" & i).Caption = Max Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Max Then UF1("C" & i).Font.Underline = True
    Case "D"
      If UF1("S" & i).Caption = Min Then UF1("S" & i).Font.Underline = True
      If UF1("U" & i).Caption = Min Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Min Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Min Then UF1("D" & i).Font.Underline = True
      If UF1("S" & i).Caption = Max Then UF1("S" & i).Font.Underline = True
      If UF1("U" & i).Caption = Max Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Max Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Max Then UF1("D" & i).Font.Underline = True
    Case "E"
      If UF1("S" & i).Caption = Min Then UF1("S" & i).Font.Underline = True
      If UF1("U" & i).Caption = Min Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Min Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Min Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Min Then UF1("E" & i).Font.Underline = True
      If UF1("S" & i).Caption = Max Then UF1("S" & i).Font.Underline = True
      If UF1("U" & i).Caption = Max Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Max Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Max Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Max Then UF1("E" & i).Font.Underline = True
    Case "F"
      If UF1("S" & i).Caption = Min Then UF1("S" & i).Font.Underline = True
      If UF1("U" & i).Caption = Min Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Min Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Min Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Min Then UF1("E" & i).Font.Underline = True
      If UF1("F" & i).Caption = Min Then UF1("F" & i).Font.Underline = True
      If UF1("S" & i).Caption = Max Then UF1("S" & i).Font.Underline = True
      If UF1("U" & i).Caption = Max Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Max Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Max Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Max Then UF1("E" & i).Font.Underline = True
      If UF1("F" & i).Caption = Max Then UF1("F" & i).Font.Underline = True
    End Select
  Case "U"
    Select Case E
    Case "U"
      If UF1("U" & i).Caption = Min Then UF1("U" & i).Font.Underline = True
      If UF1("U" & i).Caption = Max Then UF1("U" & i).Font.Underline = True
    Case "C"
      If UF1("U" & i).Caption = Min Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Min Then UF1("C" & i).Font.Underline = True
      If UF1("U" & i).Caption = Max Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Max Then UF1("C" & i).Font.Underline = True
    Case "D"
      If UF1("U" & i).Caption = Min Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Min Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Min Then UF1("D" & i).Font.Underline = True
      If UF1("U" & i).Caption = Max Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Max Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Max Then UF1("D" & i).Font.Underline = True
    Case "E"
      If UF1("U" & i).Caption = Min Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Min Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Min Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Min Then UF1("E" & i).Font.Underline = True
      If UF1("U" & i).Caption = Max Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Max Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Max Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Max Then UF1("E" & i).Font.Underline = True
    Case "F"
      If UF1("U" & i).Caption = Min Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Min Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Min Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Min Then UF1("E" & i).Font.Underline = True
      If UF1("F" & i).Caption = Min Then UF1("F" & i).Font.Underline = True
      If UF1("U" & i).Caption = Max Then UF1("U" & i).Font.Underline = True
      If UF1("C" & i).Caption = Max Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Max Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Max Then UF1("E" & i).Font.Underline = True
      If UF1("F" & i).Caption = Max Then UF1("F" & i).Font.Underline = True
    End Select
  Case "C"
    Select Case E
    Case "C"
      If UF1("C" & i).Caption = Min Then UF1("C" & i).Font.Underline = True
      If UF1("C" & i).Caption = Max Then UF1("C" & i).Font.Underline = True
    Case "D"
      If UF1("C" & i).Caption = Min Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Min Then UF1("D" & i).Font.Underline = True
      If UF1("C" & i).Caption = Max Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Max Then UF1("D" & i).Font.Underline = True
    Case "E"
      If UF1("C" & i).Caption = Min Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Min Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Min Then UF1("E" & i).Font.Underline = True
      If UF1("C" & i).Caption = Max Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Max Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Max Then UF1("E" & i).Font.Underline = True
    Case "F"
      If UF1("C" & i).Caption = Min Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Min Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Min Then UF1("E" & i).Font.Underline = True
      If UF1("F" & i).Caption = Min Then UF1("F" & i).Font.Underline = True
      If UF1("C" & i).Caption = Max Then UF1("C" & i).Font.Underline = True
      If UF1("D" & i).Caption = Max Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Max Then UF1("E" & i).Font.Underline = True
      If UF1("F" & i).Caption = Max Then UF1("F" & i).Font.Underline = True
    End Select
  Case "D"
    Select Case E
    Case "D"
      If UF1("D" & i).Caption = Min Then UF1("D" & i).Font.Underline = True
      If UF1("D" & i).Caption = Max Then UF1("D" & i).Font.Underline = True
    Case "E"
      If UF1("D" & i).Caption = Min Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Min Then UF1("E" & i).Font.Underline = True
      If UF1("D" & i).Caption = Max Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Max Then UF1("E" & i).Font.Underline = True
    Case "F"
      If UF1("D" & i).Caption = Min Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Min Then UF1("E" & i).Font.Underline = True
      If UF1("F" & i).Caption = Min Then UF1("F" & i).Font.Underline = True
      If UF1("D" & i).Caption = Max Then UF1("D" & i).Font.Underline = True
      If UF1("E" & i).Caption = Max Then UF1("E" & i).Font.Underline = True
      If UF1("F" & i).Caption = Max Then UF1("F" & i).Font.Underline = True
    End Select
  Case "E"
    Select Case E
    Case "E"
      If UF1("E" & i).Caption = Min Then UF1("E" & i).Font.Underline = True
      If UF1("E" & i).Caption = Max Then UF1("E" & i).Font.Underline = True
    Case "F"
      If UF1("E" & i).Caption = Min Then UF1("E" & i).Font.Underline = True
      If UF1("F" & i).Caption = Min Then UF1("F" & i).Font.Underline = True
      If UF1("E" & i).Caption = Max Then UF1("E" & i).Font.Underline = True
      If UF1("F" & i).Caption = Max Then UF1("F" & i).Font.Underline = True
    End Select
  Case "F"
    Select Case E
    Case "F"
      If UF1("F" & i).Caption = Min Then UF1("F" & i).Font.Underline = True
      If UF1("F" & i).Caption = Max Then UF1("F" & i).Font.Underline = True
    End Select
  End Select
Next i
End Sub
Public Sub Form_Formula()
If Not UF1.ReSize.Tag = "有" Then Exit Sub
Dim a As Long
With UF1
For a = 1 To 6
  Select Case a
  Case 1: If (.ATU <> "") And (.Ap <> "") Then .ApK = Format(val(.ATU) + val(.Ap), "0.000")
  Case 2: If (.ATU <> "") And (.Am <> "") Then .AmK = Format(val(.ATU) - val(.Am), "0.000") '- 厚
  Case 3: If (.Hab <> "") And (.Hp <> "") Then .HpK = Format(val(.Hab) + val(.Hp), "0.0")
  Case 4: If (.Hab <> "") And (.Hm <> "") Then .HmK = Format(val(.Hab) - val(.Hm), "0.0")  '-巾
  Case 5: If (.Tak <> "") And (.Tp <> "") Then .TpK = Format(val(.Tak) + val(.Tp), "0.0")
  Case 6: If (.Tak <> "") And (.Tm <> "") Then .TmK = Format(val(.Tak) - val(.Tm), "0.0") '-丈
  End Select
Next a
End With
End Sub
Public Sub Atu_OPselect(ATU As Double, Optional WeightG As Double = 0)  'OP厚自動判定
'Q-S-QA-063_厚板材の板厚測定基準
'Q-S-QA-ALL-163_板材抜取検査頻度
If ATU = 0 Then GoTo SkipA
'---------------------------------------------------------------------------------------------------------------------------
'オーダー厚で厚板薄板の定義を行う***2023.03.16
Dim HiAcc: HiAcc = Arrbox("SIKALOTNOW.accdb", "板", "Fields", UF1.DBflag.Text)
Dim OrdAtu As Long, i As Long, No1 As Long, No2 As Long, C_Course As Long
If Not IsEmpty(HiAcc) Then
  For i = 1 To UBound(HiAcc, 2)
    Select Case HiAcc(1, i)
    Case "ｵｰﾀﾞｰ板厚":    OrdAtu = i
    Case "検査番号":     No1 = i
    Case "実績_設備ｺｰｽ": C_Course = i
    End Select
  Next i
End If
Dim LotNo As String: LotNo = UF1.Lot
Dim LotAcc: LotAcc = Arrbox("SIKALOTNOW.accdb", LotNo, "Record", UF1.DBflag.Text)

Dim NewAtu As Double
If IsEmpty(LotAcc) = False Then NewAtu = LotAcc(1, OrdAtu) 'ｵｰﾀﾞｰ厚参照


'厚の再定義フォーム数値からオーダー数値へ
If (val(ATU) <> val(NewAtu)) Then UF1.Comment.Caption = "ｵｰﾀﾞｰ厚参照_" & Format(NewAtu, "0.000")
ATU = NewAtu
'---------------------------------------------------------------------------------------------------------------------------
'2023.4.17_検査No検索

Dim SIKAHIKAcc: SIKAHIKAcc = Arrbox("SIKAHIKINOW.accdb", "", "Fields", UF1.DBflag.Text) 'SIKAHIKINOW.accdbのフィールドに[ｵｰﾀﾞｰ板丈]がないのでエラーになる
If Not IsEmpty(SIKAHIKAcc) Then
  For i = 1 To UBound(SIKAHIKAcc, 2)
    Select Case SIKAHIKAcc(1, i)
    Case "検査番号": No2 = i
    End Select
  Next i
End If

Dim FPath As String
Dim target As String: target = "SIKAHIKINOW.accdb"
If UF1.DBflag.Text = "" Then
 FPath = MainPath
Else
 FPath = NextPath
End If

Dim FilePath As String: FilePath = FPath & target
If FPath = MainPath Then
  If ReadCheck(FilePath, 1) = False Or TimeDesignation Or FileDesignation = False Then   '1：ファイル,2：フォルダ
    FPath = NextPath
    FilePath = FPath & target
    If ReadCheck(NextPath & target, 1) = False Then '梱包課ﾌｧｲﾙを使う*1日更新
      MsgBox "日稲:" & target & "がありません" & "_ReadCheckAtu_OPselect", vbCritical
      UF1.Comment.Caption = ""
      Exit Sub
    End If
  End If
End If

Dim TableName As String: TableName = "仕掛"
Dim SQLfilter As String: SQLfilter = "[ﾛｯﾄ番号]='" & LotNo & "'"
Dim KenAcc: KenAcc = GetRecordsArr(FilePath, TableName, SQLfilter)

Dim KenNo As String
If IsEmpty(KenAcc) = False Then
  KenNo = Left(KenAcc(1, No2), 1)  '検査No'頭1桁
  UF1.KenNo.Caption = KenAcc(1, No2)
Else
  If IsEmpty(LotAcc) = False Then
    KenNo = Left(LotAcc(1, No1), 1)
    UF1.KenNo.Caption = LotAcc(1, No1)
  End If
End If

If IsEmpty(LotAcc) = False Then
  If KenNo = "0" Or KenNo = "" Then
  'Q-S-QA-ALL-088_検査番号の付け方に基づき作る
  Dim InputString As String: InputString = LotAcc(1, C_Course)
  Dim Result As String:      Result = FindPriority(InputString) 'コース最終圧延検索
    If Result <> "" Then
      MsgBox "最終圧延機から検査番号頭を生成します" & vbCrLf & "オーダーが付いていないので納入先が不明" & _
      vbCrLf & "ラテラルボーが必要の可能性があります" & vbCrLf & "Q-S-QA-ALL-088_検査番号の付け方", vbInformation
      KenNo = Result
    End If
  
    If Result = "" Then
      MsgBox "データ上にまだ検査番号がありません" & vbCrLf & "(一部動作しない機能があります)", vbInformation
      UF1.Comment.Caption = ""
    End If
  End If
End If
SkipA:

'***********************************リセット
Dim MyCtrl As Control 'ｵｰﾄﾒｰｼｮﾝｴﾗｰ対策_UF1("S" & i).Visible = True で行うとﾌﾞｯｸ開閉を繰り返したときｴﾗｰ Controlsで管理する**2022.06
For Each MyCtrl In UF1.Controls
  If (TypeName(MyCtrl) = "Label") Then 'ラベルのみ
    Select Case Left(MyCtrl.Name, 1) '名前先頭に指定の5文字
    Case "S", "U", "C", "D", "E", "F" '検査データ用ラベル
      If Len(MyCtrl.Name) = 3 Then '番号チェック　3桁
        If Right(MyCtrl.Name, 1) = "0" Then '10の時
          MyCtrl.Visible = True
          MyCtrl.Enabled = True
        Else '10以外　11,12,13,14
          MyCtrl.Visible = False
          MyCtrl.Enabled = False
        End If
      ElseIf Len(MyCtrl.Name) = 2 Then 'FS~
        If IsNumeric(Right(MyCtrl.Name, 1)) Then '2024.5.29_追加
          MyCtrl.Visible = True
          MyCtrl.Enabled = True
        Else
          MyCtrl.Visible = True
          MyCtrl.Enabled = False
        End If
      ElseIf Len(MyCtrl.Name) = 1 Then
        MyCtrl.Visible = True
        MyCtrl.Enabled = False
      Else '
        MyCtrl.Visible = True
        MyCtrl.Enabled = True
      End If
    End Select

    If Left(MyCtrl.Name, 2) = "AL" Then MyCtrl.Visible = True '厚用表示ﾗﾍﾞﾙ

    If Left(MyCtrl.Name, 1) = "L" Then '巾丈用表示ﾗﾍﾞﾙ
      If Len(MyCtrl.Name) = 3 Then
        Select Case Right(MyCtrl.Name, 2)
        Case "11" 'ﾗﾃﾗﾙﾎﾞｰ用
      MyCtrl.Visible = False
      If UF1.Sumisei.Caption <> "" Or UF1.LateralText.Tag = "有" Then MyCtrl.Visible = True '2026.07.13_技術指示経路も対象に追加
        Case "12", "13"
          MyCtrl.Visible = False
        Case 10 '直角度に変更のため　デフォルト非表示_2023.11.27
'          MyCtrl.Visible = False
        End Select
      ElseIf Len(MyCtrl.Name) = 2 Then '1　to　9
        Select Case Right(MyCtrl.Name, 1)
        Case "7", "8"
          MyCtrl.Visible = False
        Case Else
          MyCtrl.Visible = True
        End Select
      End If
    End If

    If Left(MyCtrl.Name, 4) = "Diag" Then '検査対角用
      MyCtrl.Visible = True
      MyCtrl.Enabled = True
    End If

    If MyCtrl.Name = "ConductivL" Then MyCtrl.Visible = False '導電率用
  End If
  If (TypeName(MyCtrl) = "Frame") Then 'フレームのみ
    If Left(MyCtrl.Name, 2) = "FM" Then MyCtrl.Visible = True '歪フレーム
  End If
Next MyCtrl

'trueをすべて外してから再表示_2021.02.08
For Each MyCtrl In UF1.Controls
  If (TypeName(MyCtrl) = "OptionButton") Then
  'DefaultFormPositionの設定のため別OptionButtonをクリアすると位置ずれする
    Select Case MyCtrl.Name
    Case "厚板"
      MyCtrl = False
      MyCtrl.Visible = True
    Case "薄板"
      MyCtrl = False
      MyCtrl.Visible = True
    Case "少枚数"
      MyCtrl = False
      MyCtrl.Visible = True
    Case "極厚"
      MyCtrl = False
      MyCtrl.Visible = True
    Case "スラブ"
      MyCtrl = False
      MyCtrl.Visible = True
    End Select
  End If
Next MyCtrl

UF1.AtButt.Visible = False '2021.05.18_ﾘｾｯﾄ
UF1.InspeButt.Visible = False '2021.11.29

UF1.Yama.Visible = False
UF1.YamaLabel.Visible = False

UF1.ConductivText.Caption = ""

'************スラブ追加************2023.4.17_条件定義を検査No頭一桁目に変更
'不要のOPを非表示にする
'2023.05.17_いったん板厚で判定後表示させ、表示されたOptionButtonを検番に変えるに変更(厚板薄板のCaptionのみ書き換えたいため)

If KenNo = "" Then
  UF1.Comment.Caption = ""
  Exit Sub
End If

'特殊条件_検査番号頭T'2023.6.7
If KenNo = "T" Then
  UF1.薄板.Caption = KenNo
  UF1.薄板.Visible = True
  UF1.厚板.Visible = False
  UF1.極厚.Visible = False
  UF1.スラブ.Visible = False
  UF1.少枚数.Visible = False
Exit Sub
End If

For Each MyCtrl In UF1.Controls
  If (TypeName(MyCtrl) = "OptionButton") Then
    Select Case MyCtrl.Name
    Case "厚板"
      Select Case NewAtu
      Case 1 To 6#    '1 To 7.999*2023.03_条件見直し
        MyCtrl.Visible = False
        UF1.薄板.Caption = KenNo
      Case 200.1 To 299#
        '2022.10.03_小板で表示したい場合がある　重量で別チェックをかける
          If WeightG <> 0 Then
            If WeightG > 1500# Then
              MyCtrl.Visible = False
            End If
          End If
      Case 300# To 800#
        MyCtrl.Visible = False
      End Select
    Case "薄板"
      Select Case NewAtu
      Case 6.001 To 800#:  MyCtrl.Visible = False '8 To 30 *2023.03_条件見直し
      End Select
      UF1.厚板.Caption = KenNo
    Case "少枚数"
      Select Case NewAtu
      Case 1 To 6#:        MyCtrl.Visible = False
      End Select
    Case "極厚"
      Select Case NewAtu
      Case 1 To 29.9:      MyCtrl.Visible = False
      End Select
    Case "スラブ"
      Select Case NewAtu
      Case 1 To 199.9:     MyCtrl.Visible = False
      End Select
    End Select
  End If
Next MyCtrl

End Sub
Public Sub IM_Select()  '画像選択_2021.01.07追加
UF1.L1IM.Visible = False
UF1.LHIM.Visible = False
UF1.LHIM.Visible = False
UF1.TBIM.Visible = False
UF1.TBIM.Visible = False
UF1.GokuIM.Visible = False

Select Case True
Case UF1.L1Work: UF1.L1IM.Visible = True
Case UF1.LVC:    UF1.LHIM.Visible = True
Case UF1.HVC:    UF1.LHIM.Visible = True
Case UF1.Bala:   UF1.TBIM.Visible = True
Case UF1.Tott:   UF1.TBIM.Visible = True
Case Else:
End Select
If UF1.L1Work Then
  If UF1.極厚 Then
    UF1.L1IM.Visible = False
    UF1.SurabuIM.Visible = False
    UF1.GokuIM.Visible = True
  ElseIf UF1.スラブ Then
    UF1.L1IM.Visible = False
    UF1.GokuIM.Visible = False
    UF1.SurabuIM.Visible = True
  End If
End If
End Sub
Public Sub VC_Clear() 'VC画像がクリアされたとき
Application.ScreenUpdating = False
UF1.VC_A.Picture = LoadPicture("") '2023.9.11_画像リセットも追加
UF1.VC_B.Picture = LoadPicture("")
UF1.VC_L_A.Visible = False
UF1.VC_L_B.Visible = False
UF1.VC_A.Visible = False
UF1.VC_B.Visible = False

UF1.Height = 545
If UF1.Yama.Text = "4" Then UF1.Height = 565

With Application
  If .WindowState = xlNormal Then
    .Height = 565
  End If
End With
Application.ScreenUpdating = True
End Sub

Public Sub Line_AutoSelection(Optional n As String = "") '設備読み込み n=""でﾌﾟﾛｸﾞﾚｽﾊﾞｰ非表示
'-----2021.03.04_Focus対策
UF1.Worker.Enabled = True
UF1.INP.Enabled = True
UF1.Manual.Enabled = True
UF1.StartTB.Enabled = True
Select Case True
Case UF1.L1Work: UF1.LineName = UF1.L1Work.Caption
Case UF1.LVC:    UF1.LineName = UF1.LVC.Caption
Case UF1.HVC:    UF1.LineName = UF1.HVC.Caption
Case UF1.Tott:   UF1.LineName = UF1.Tott.Caption
Case UF1.Bala:   UF1.LineName = UF1.Bala.Caption
Case UF1.Test:   UF1.LineName = UF1.Test.Caption
Case UF1.Maru:   UF1.LineName = UF1.Maru.Caption
End Select

Dim i As Long
Select Case True
Case UF1.Maru '丸徳の時のみ表示_2021.10.07ライン追加
  For i = 1 To 11
   If UF1.Controls("A" & i).Value = True Then UF1.LineNo = UF1("A" & i).Caption
  Next i
Case Else
  UF1.LineNo = "" 'ｸﾘｱ
  For i = 1 To 11
    UF1.Controls("A" & i).Value = False
  Next i
End Select

'2021.12.15_丸徳用
Select Case Mid(UF1.LineName, 1, 2)
Case "丸徳": UF1.ZEB.Caption = "検入4山以上" '5回検査であっても薄板フォームセットが検査回数3回になってしまうため
Case Else:   UF1.ZEB.Caption = "全数.三菱.5N"
End Select

UF1.Tag = "STOP" '音楽ｽﾄｯﾌﾟ
UF1.UFTag = UF1.Tag '目視用
UF1.Load = "off" 'ﾌｫｰﾑｲﾍﾞﾝﾄ停止


'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'-----------------------------------自動での読み込み
Call Load_Text(n) '作業保存読み込み　n=""でﾌﾟﾛｸﾞﾚｽﾊﾞｰ非表示
If UF1.Lot = "" Then GoTo SkipLoad 'LotNoがないなら不要
'------------------------------------------------------------------------------------------------
Dim TempArry As Variant: TempArry = GetLotInfo("Lot情報履歴", UF1.Lot.Text)

If IsEmpty(TempArry) Then
  Call WorkerLoad '検査員読み込み
  Exit Sub
End If

'2026.10.14_引当混在対策：LSelectと共通の判定を使う（ここではﾗﾃﾗﾙﾎﾞｰ欄の描写はしない）
Dim 検番 As String
Dim NeedLateral As Boolean, NeedRight As Boolean, NeedFlat As Boolean
Call JudgeLotAllocations(TempArry, 検番, NeedLateral, NeedRight, NeedFlat)
If NeedFlat Then UF1.flatnesText = "ﾌﾘｰｿﾘ要"
  
'------------------------------------------------------------------------------------------------
'デフォルト表示は不要
'If (UF1.ATU) <> "" Then '2021.01.28
'  If (Weight(Val(UF1.ATU), Val(UF1.Hab), Val(UF1.Tak)) > 1300# And Val(UF1.Tak) < 2500#) Or _
'  (Weight(Val(UF1.ATU), Val(UF1.Hab), Val(UF1.Tak)) > 2500#) Then 'ｽﾘﾝｸﾞ検査の時_2021.01.13
'    UF1("LZS").Caption = "ｹﾞｰｼﾞ"
'    UF1("LZU").Caption = "ｹﾞｰｼﾞ"
'    UF1("LZC").Caption = "ｹﾞｰｼﾞ"
'    UF1("LZD").Caption = "ｹﾞｰｼﾞ"
'    UF1("LZE").Caption = "ｹﾞｰｼﾞ"
'    UF1("LGS").Caption = "ｹﾞｰｼﾞ"
'    UF1("LGU").Caption = "ｹﾞｰｼﾞ"
'    UF1("LGC").Caption = "ｹﾞｰｼﾞ"
'    UF1("LGD").Caption = "ｹﾞｰｼﾞ"
'    UF1("LGE").Caption = "ｹﾞｰｼﾞ"
'    Call CBAdd(SikaSetAtu)
'  End If
'End If

SkipLoad:
Call WorkerLoad '検査員読み込み
Call LoopJudg '検査判定
Call HizLoopJudg

'2025.1.23
Call CheckLabelsByCategory("S")
Call CheckLabelsByCategory("U")
Call CheckLabelsByCategory("C")
Call CheckLabelsByCategory("D")
Call CheckLabelsByCategory("E")
Call CheckLabelsByCategory("F")

'------------------------------------
UF1.Load = "on" 'ﾌｫｰﾑｲﾍﾞﾝﾄ停止
UF1.Tag = "" '音楽ｽﾀｰﾄ
UF1.UFTag = UF1.Tag

End Sub

Public Sub DefaultFormPosition() 'ｸﾘｯｸさせる前にﾌｫｰﾑの位置を戻す
On Error Resume Next

'Application.WindowState = xlMaximized
'UF1.Top = (Application.Height - UF1.Height) / 2 '画面の真ん中に合わせる
'UF1.Left = (Application.Width - UF1.Width) / 2 + 48 '画面の真ん中に合わせる
'2026.05.29_↑を↓にした
Dim screenW As Long: screenW = GetSystemMetrics32(0)
Dim screenH As Long: screenH = GetSystemMetrics32(1)
Dim formWpx As Long: formWpx = CLng(UF1.Width * 96 / 72)
Dim formHpx As Long: formHpx = CLng(UF1.Height * 96 / 72)
Dim newLeft As Long: newLeft = (screenW - formWpx) / 2
Dim newTop  As Long: newTop = (screenH - formHpx) / 2
' ピクセルをポイントに戻してフォームに設定
UF1.Left = newLeft * 72 / 96
UF1.tOP = newTop * 72 / 96
End Sub

Public Sub DistSelect(n As Long) '薄板時の歪表示
Select Case n
Case 25
  UF1.FMU.Visible = False
  UF1.FMC.Visible = False
  UF1.FMD.Visible = False
Case 26
  UF1.FMU.Visible = False
  UF1.FMC.Visible = True
  UF1.FMD.Visible = False
Case Else
  UF1.FMS.Visible = True
  UF1.FMU.Visible = True
  UF1.FMC.Visible = True
  UF1.FMD.Visible = True
  UF1.FME.Visible = True
End Select
End Sub

Public Function PackingNoCheck(LotNo As String) As String '包装仕様書Noチェック 仕様書Noを返す
'1P0113'裸
'1P0114'裸脚
'1P0117'裸2山以上
'1P0118'縦バンド
'1P0119'縦バンド
'1P0121'裸1山
'1P1189'北海道
'1P1231'NTC裸
'1P1232'三菱
If WsCheck("Lot情報履歴") = False Then Exit Function '-------------------------ｼｰﾄなし処理停止
If LotNo = "" Then Exit Function
Dim HisSH As Worksheet:   Set HisSH = ThisWorkbook.Sheets("Lot情報履歴")
Dim HisRan As Range:      Set HisRan = HisSH.Range(HisSH.Cells(1, 1), HisSH.Cells(UR(HisSH, 1), UC(HisSH, 1)))
Dim Hist(): Hist = HisRan
Dim i As Long, a As Long, n As Long: n = 0
Dim R_Lot As Long, R_Ord As Long, R_Packing As Long, R_Hik As Long, R_材質 As Long, R_調質 As Long, R_厚 As Long, R_コース As Long, R_用途コード As Long
For i = 1 To UBound(Hist, 2) 'ﾌｨｰﾙﾄﾞ名
  Select Case Hist(1, i)
  Case "ﾛｯﾄNo":        R_Lot = i
  Case "ｵｰﾀﾞｰNo":      R_Ord = i
  Case "引当NO":       R_Hik = i
  Case "包装仕様NO":   R_Packing = i
  Case "材質":         R_材質 = i
  Case "調質":         R_調質 = i
  Case "ｺｰｽ":          R_コース = i
  Case "厚":           R_厚 = i
  Case "用途ｺｰﾄﾞ":     R_用途コード = i
  End Select
Next i

'-------------------------------------------------------------------------
''1P0122の外注梱包に限り強制表示*引当通りでないため_2023.07.7
'Dim flag(1 To 5) As Boolean
'For i = LBound(Hist, 1) To UBound(Hist, 1)
'  If Hist(i, R_Lot) = LotNo Then '一致の時
'    If Hist(i, R_材質) Like "*75S" Then flag(1) = True
'    If Hist(i, R_調質) Like "T6*" Then flag(2) = True
'    If Hist(i, R_用途コード) = "K434" Or Hist(i, R_用途コード) = "K435" Then '2024.5.23_用途コード追加
'      flag(3) = True
'    End If
'    If Hist(i, R_厚) = 6.75 Then flag(4) = True
'    Select Case True
'    Case InStr(Hist(i, R_コース), "GCT") > 0: flag(5) = True
'    Case InStr(Hist(i, R_コース), "GFS") > 0: flag(5) = True
'    Case InStr(Hist(i, R_コース), "GSS") > 0: flag(5) = True '2024.9.17
'    End Select
'  End If
'Next i
'
'If flag(1) And flag(2) And flag(3) And flag(4) And flag(5) Then
'  PackingNoCheck = "1P0122" '外注梱包　通常引き当てでは表示できないのでここでexit
'  Exit Function
'End If
'
''2024.12.10_裸梱包
''2026.09.18　厚変わってやがる　くそが
'For i = LBound(Hist, 1) To UBound(Hist, 1)
'  If Hist(i, R_Lot) = LotNo Then
'    If Hist(i, R_材質) Like "*52S" Then flag(1) = True
'    If Hist(i, R_調質) = "R" Then flag(2) = True
'    If Hist(i, R_用途コード) = "H19S" Then
'      flag(3) = True
'    End If
''    If Hist(i, R_厚) >= 350 Then flag(4) = True
'    If Hist(i, R_厚) >= 322.5 Then flag(4) = True
'    Select Case True
'    Case InStr(Hist(i, R_コース), "GCT") > 0: flag(5) = True
'    Case InStr(Hist(i, R_コース), "GFS") > 0: flag(5) = True
'    Case InStr(Hist(i, R_コース), "GSS") > 0: flag(5) = True
'    End Select
'  End If
'Next i
'If flag(1) And flag(2) And flag(3) And flag(4) And flag(5) Then
'  PackingNoCheck = "1P0113" '外注梱包　通常引き当てでは表示できないのでここでexit
'  Exit Function
'End If
'-------------------------------------------------------------------------2026.09.18 DB読み込みに変更
'外注梱包　通常引き当てでは表示できないのでここでexit
'条件は品質基準.accdb の 外注条件 / 外注条件明細 で管理
'  1P0122 … 引当通りでないため強制表示(2023.07.07)
'  1P0113 … 裸梱包(2024.12.10)
Dim GaichuNo As String: GaichuNo = GaichuPackingNo(Hist, LotNo, R_Lot)
If GaichuNo <> "" Then
  PackingNoCheck = GaichuNo
  Exit Function
End If
'-------------------------------------------------------------------------


Dim FL As Boolean: FL = False
Dim TempHiki()
For i = LBound(Hist, 1) To UBound(Hist, 1) '1次元ﾁｪｯｸ
  If Hist(i, R_Lot) = LotNo Then a = a + 1
Next i

If a = 0 Then Exit Function

ReDim TempHiki(1 To a, 1 To 4)
For i = LBound(Hist, 1) To UBound(Hist, 1)
  If Hist(i, R_Lot) = LotNo Then '一致の時
    n = n + 1
    TempHiki(n, 1) = Hist(i, R_Lot)
    TempHiki(n, 2) = Hist(i, R_Ord)
    TempHiki(n, 3) = Hist(i, R_Packing)
    TempHiki(n, 4) = Hist(i, R_Hik) '並び替えのために配列に入れる(引当1から表示させる)
  End If
Next i

Call merge_sort2(TempHiki, 4) '配列内の並び替え_2022.02.01***引当No基準ｿｰﾄ
'Stop

Dim FILE_PATH As String: FILE_PATH = QualityPath & "包装仕様書" '"\\nlmsrvngy03\工場内共有\検査データ\包装仕様書" '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\検査\包装仕様"
Dim fso As Object:       Set fso = New FileSystemObject
Dim target As Files:     Set target = fso.GetFolder(FILE_PATH).Files
Dim Temp As Object
Dim fileName As String:  fileName = "*" '"*.jpg" 'ファイル名をワイルドカードと拡張を指定する
Dim FileInt As Long:     FileInt = 0 'ファイル数を一度0にする
Dim SetPath As String:   SetPath = Dir(FILE_PATH & "\" & fileName)
Do While SetPath <> "" 'ファイル名が取得出来なくなるまでループ
  FileInt = FileInt + 1
  SetPath = Dir()
Loop
Dim strFileExExt As String '拡張子を除いたパス（ファイル名）を格納
ReDim PacN(1 To FileInt) 'デバック用
Dim posExt As Long, intNo As Long: intNo = 0
For Each Temp In target
  posExt = InStrRev(Temp.Name, ".") '文字列の中から別の文字列を後方から検索
  If (posExt = 7) Then '6P0101.jpgとかなので7が返ってくるもの** (0 < posExt)から修正_2023.6.26
    intNo = intNo + 1 '**ｺﾒﾝﾄｱｳﾄでもいいよ
    strFileExExt = Left(Temp.Name, 6) '6P0101.jpgとかなので必要な文字列は左から6文字**Left(temp.Name, posExt - 1)から修正_2023.6.26
  Else
    strFileExExt = Temp.Name
  End If
  PacN(intNo) = strFileExExt 'デバック用に別配列に突っ込む**ｺﾒﾝﾄｱｳﾄでもいいよ
  For i = LBound(TempHiki, 1) To UBound(TempHiki, 1)
    If TempHiki(i, 3) = strFileExExt Then '一致の時
'    MsgBox "一致＿" & strFileExExt
      PackingNoCheck = strFileExExt
      Exit For
    End If
  Next i
Next

'Stop

End Function

Public Function Loop_Label(La As MSForms.Label) 'ラベル行ごとに変更させる
Select Case (Right(La.Name, 1))
Case 1: Call Label_ForeColor(1)
Case 2: Call Label_ForeColor(2)
Case 3: Call Label_ForeColor(3)
Case 4: Call Label_ForeColor(4)
Case 5: Call Label_ForeColor(5)
End Select
End Function
Public Sub Label_ForeColor(n As Long) '文字色変更
Dim i As Long
Select Case n
Case 1
  For i = 1 To 5
    UF_PackNo("Hiki" & i).ForeColor = RGB(0, 0, 0)
    UF_PackNo("Ord" & i).ForeColor = RGB(0, 0, 0)
    UF_PackNo("Pac" & i).ForeColor = RGB(0, 0, 0)
  Next i
  UF_PackNo.Hiki1.ForeColor = RGB(255, 0, 0)
  UF_PackNo.Ord1.ForeColor = RGB(255, 0, 0)
  UF_PackNo.Pac1.ForeColor = RGB(255, 0, 0)
Case 2
  For i = 1 To 5
    UF_PackNo("Hiki" & i).ForeColor = RGB(0, 0, 0)
    UF_PackNo("Ord" & i).ForeColor = RGB(0, 0, 0)
    UF_PackNo("Pac" & i).ForeColor = RGB(0, 0, 0)
  Next i
  UF_PackNo.Hiki2.ForeColor = RGB(255, 0, 0)
  UF_PackNo.Ord2.ForeColor = RGB(255, 0, 0)
  UF_PackNo.Pac2.ForeColor = RGB(255, 0, 0)
Case 3
  For i = 1 To 5
    UF_PackNo("Hiki" & i).ForeColor = RGB(0, 0, 0)
    UF_PackNo("Ord" & i).ForeColor = RGB(0, 0, 0)
    UF_PackNo("Pac" & i).ForeColor = RGB(0, 0, 0)
  Next i
  UF_PackNo.Hiki3.ForeColor = RGB(255, 0, 0)
  UF_PackNo.Ord3.ForeColor = RGB(255, 0, 0)
  UF_PackNo.Pac3.ForeColor = RGB(255, 0, 0)
Case 4
  For i = 1 To 5
    UF_PackNo("Hiki" & i).ForeColor = RGB(0, 0, 0)
    UF_PackNo("Ord" & i).ForeColor = RGB(0, 0, 0)
    UF_PackNo("Pac" & i).ForeColor = RGB(0, 0, 0)
  Next i
  UF_PackNo.Hiki4.ForeColor = RGB(255, 0, 0)
  UF_PackNo.Ord4.ForeColor = RGB(255, 0, 0)
  UF_PackNo.Pac4.ForeColor = RGB(255, 0, 0)
Case 5
  For i = 1 To 5
    UF_PackNo("Hiki" & i).ForeColor = RGB(0, 0, 0)
    UF_PackNo("Ord" & i).ForeColor = RGB(0, 0, 0)
    UF_PackNo("Pac" & i).ForeColor = RGB(0, 0, 0)
  Next i
  UF_PackNo.Hiki5.ForeColor = RGB(255, 0, 0)
  UF_PackNo.Ord5.ForeColor = RGB(255, 0, 0)
  UF_PackNo.Pac5.ForeColor = RGB(255, 0, 0)
End Select
End Sub
Public Sub PackSelect(n As String)
UF_PackNo.L1.Caption = ""
UF_PackNo.L2.Caption = ""
UF_PackNo.L3.Caption = ""
UF_PackNo.L4.Caption = ""
UF_PackNo.L5.Caption = ""
'仕様書ｺﾒﾝﾄ表示
Dim TableName As String: TableName = "PkgSpecData"
Dim FPath As String: FPath = PackingPath
'ﾌｧｲﾙ名
Dim target As String:   target = "伝送用ファイル.accdb"
'ﾌｧｲﾙﾊﾟｽ
Dim FilePath As String: FilePath = FPath & target
  If ReadCheck(FilePath, 1) = False Then Exit Sub '1：ファイル,2：フォルダ
'検索項目
Dim SQLfilter As String: SQLfilter = "[包装仕様書№]='" & n & "'" '仕様書No検索
Dim Temp As Variant:     Temp = GetRecordsArr(FilePath, TableName, SQLfilter) 'All,fields,record
'ﾌｨｰﾙﾄﾞ名ﾁｪｯｸ　*更新に対応
Dim fieldsfilter As String: fieldsfilter = ""
Dim Fields As Variant:      Fields = GetFieldsArr(FilePath, TableName, fieldsfilter)
If IsEmpty(Fields) Then Exit Sub 'emptyの場合停止
Dim C_包装仕様書№ As Long, C_梱包形態 As Long, C_梱包指定別 As Long, C_タイトル As Long, C_用途名 As Long, i As Long
For i = LBound(Fields, 2) To UBound(Fields, 2)
  Select Case Fields(1, i)
  Case "梱包形態":      C_梱包形態 = i
  Case "包装仕様書№":  C_包装仕様書№ = i
  Case "梱包指定別":    C_梱包指定別 = i
  Case "タイトル":      C_タイトル = i
  Case "用途名":        C_用途名 = i
  End Select
Next i
UF_PackNo.L1.Caption = n
UF_PackNo.L2.Caption = "梱包形態:" & Temp(1, C_梱包形態)
UF_PackNo.L3.Caption = "梱包指定:" & Temp(1, C_梱包指定別)
UF_PackNo.L4.Caption = "特記1:" & Temp(1, C_タイトル)
UF_PackNo.L5.Caption = "特記2:" & Temp(1, C_用途名)

Dim iLog As Long
For iLog = 1 To 5
'    Call DebugLog("Pac" & iLog & _
'        " Font.Size=" & UF_PackNo("Pac" & iLog).Font.Size & _
'        " Height=" & UF_PackNo("Pac" & iLog).Height & _
'        " Top=" & UF_PackNo("Pac" & iLog).tOP & _
'        " Width=" & UF_PackNo("Pac" & iLog).Width)
Next iLog


UF_PackNo.Repaint
End Sub

Public Sub JPGLoad()
Dim i As Long, Cont As Long
For i = 1 To 5
  If UF_PackNo("Hiki" & i).ForeColor = RGB(255, 0, 0) Then
    Cont = i
    Exit For
  End If
Next i
Dim n As String: n = UF_PackNo("Pac" & Cont).Caption
Dim a As String: a = UF_PackNo("Ord" & Cont).Caption
'”\”忘れないこと
Dim PackPath As String: PackPath = QualityPath & "包装仕様書\" '"\\nlmsrvngy03\工場内共有\検査データ\包装仕様書\" '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\検査\包装仕様\"
Dim PackPic As String:  PackPic = n & ".JPG"
If ReadCheck((PackPath & PackPic), 1) = True Then 'ﾈｯﾄﾜｰｸにつながっていないor登録がない
  Dim pic As Object: Set pic = LoadPicture(PackPath & PackPic)
  UF_PackNo.PacIM.Picture = pic
  With UF_PackNo.FMBar
    .ScrollHeight = UF_PackNo.PacIM.Height
    .ScrollWidth = UF_PackNo.PacIM.Width
    .Picture = pic
  End With
  Call PackSelect(n)
  
  '上部緩衝材の表示_2022.02.04
  Dim result1 As String: result1 = MaterialsUpperPart(a)          ' ザラ板/アングル
  Dim result2 As String: result2 = MaterialsUpperPart(a, True)    ' 当板有無
  
    If result1 <> "" And result1 <> "一致なし" Then
      MsgBox result1
    End If
    If result1 = "アングル" Then
        If result2 <> "" And result2 <> "一致なし" Then MsgBox result2
    End If
End If
End Sub
Public Sub EraseCount(TB As MSForms.Label) '座標と数を残す*2024.4.24
'対角線外れの場合消去対象が2つのためカウント数が必ずリセットされてしまう_2023.9.28 修正
Dim Tbdigit As Long
Dim Temp As String:         Temp = TB.Name
Dim Tempalphabet As String: Tempalphabet = Left(TB.Name, 1) '座標アルファベット部分
If Len(TB.Name) = 3 Then Tbdigit = Right(TB.Name, 2) '数字部分
If Len(TB.Name) = 2 Then Tbdigit = Right(TB.Name, 1)

If UF1.厚板 Or UF1.薄板 Or UF1.少枚数 Then 'スラブ、極厚以外の検査
'9,10が対角線どちらが来ても10にすることでカウントを増やす(DeleteC)
  If Tbdigit = 9 Or Tbdigit = 10 Then
    Temp = Tempalphabet & "10"
  End If
End If

Dim i As Long: i = 1
If UF1.DeleteP.Caption = Temp Then '同じ座標
  UF1.DeleteC.Caption = val(UF1.DeleteC.Caption) + i '加算
Else
  UF1.DeleteC.Caption = 1 '違う座標ならば1に戻す
  UF1.DeleteP.Caption = Temp
End If

Call Posting
End Sub
Public Function EraseCount_F(n As Long) As Boolean '決めた数字以上でフラグtrue
If n >= 2 Then EraseCount_F = True
End Function
Public Sub frameblindfold(tMP As String)
If tMP = "目隠し" Then
  UF1.PicFM.Visible = False
  UF1.MP1.Visible = False
  UF1.MPEX.Visible = False
  UF1.blindfm.Left = 5
  UF1.blindfm.Height = UF1.Height
  UF1.blindfm.Width = UF1.Width + 25
  UF1.blindfm.Visible = True
  With UF1.UfText
    .Visible = True
    .Left = 200
    .tOP = 180
    
'    .Top = (UF1.Height) / 2
'    .Left = (UF1.Width) / 2
    
    .Font.Size = 55
    .ForeColor = RGB(255, 0, 0)
    .TabStop = False
    .Locked = False
    .Text = "少々お待ちください"
    .AutoSize = True
  End With
  UF1.Repaint
ElseIf tMP = "通常" Then
  UF1.MP1.Visible = True
  UF1.MPEX.Visible = True
  UF1.blindfm.Height = 8
  UF1.blindfm.Width = 8
  UF1.UfText.Visible = False
  UF1.blindfm.Visible = False
  UF1.Repaint
End If
End Sub
Public Sub OpForcedSelection() 'サイズをコンパクトにし作業選択を強調する_2024.5.10
If UF1.Load = "off" Then Exit Sub
If UF1.Lot = "" Then Exit Sub

On Error Resume Next
Dim Temp As Variant
If UF1.厚板 = False And UF1.薄板 = False And UF1.少枚数 = False And UF1.極厚 = False And UF1.スラブ = False Then
  UF1.Height = 280
  UF1.Width = 1030
  UF1.blindfm.Height = 195
  UF1.blindfm.Width = 465
  UF1.blindfm.tOP = 132
  UF1.blindfm.Left = 0
  UF1.blindfm.Visible = True
  With UF1.UfText
    .Visible = True
    .Left = 10
    .tOP = 30
    .Font.Size = 35
    .TabStop = False
    .Locked = False
    .Text = "作業選択を選んでください"
    .AutoSize = True
    Temp = InitialSetting(UF1, "初期値からスタート")
  End With
  UF1.Repaint
Else
  If UF1.VC_L_A.Visible Or UF1.VC_L_B.Visible Then 'VCありなし
    UF1.Height = 700
    Temp = InitialSetting(UF1, "初期値からスタート")
  Else
    UF1.Height = 545
    If (UF1.Yama.Visible) Then
      If (UF1.Yama.Text = "4") Then UF1.Height = 562
'    Else’描写が全部同じになっちゃうからいらない
'      Call Six_Display("off")
    End If
    Temp = InitialSetting(UF1, "初期値からスタート")
  End If
  UF1.blindfm.Height = 8
  UF1.blindfm.Width = 8
  UF1.blindfm.tOP = 6
  UF1.blindfm.Left = 0
  UF1.UfText.Visible = False
  UF1.blindfm.Visible = False
  UF1.Repaint
End If

End Sub
'2024.5.21_検査箇所6回目の表示非表示
Public Sub Six_Display(n As String)
Dim i As Long

With UF1
If n <> "on" Then GoTo SkipA '描写リセットをスキップする
'リセット
  For i = 1 To 10
    UF1("S" & i).Visible = True
    UF1("U" & i).Visible = True
    UF1("C" & i).Visible = True
    UF1("D" & i).Visible = True
    UF1("E" & i).Visible = True
    
    UF1("S" & i).Enabled = True
    UF1("U" & i).Enabled = True
    UF1("C" & i).Enabled = True
    UF1("D" & i).Enabled = True
    UF1("E" & i).Enabled = True
  Next i

  .DiagS.Visible = True '対角
  .DiagU.Visible = True
  .DiagC.Visible = True
  .DiagD.Visible = True
  .DiagE.Visible = True
  
  .DiagS.Enabled = True
  .DiagU.Enabled = True
  .DiagC.Enabled = True
  .DiagD.Enabled = True
  .DiagE.Enabled = True

  .FMS.Visible = True '歪
  .FMU.Visible = True
  .FMC.Visible = True
  .FMD.Visible = True
  .FME.Visible = True
  
  '外観
  For i = 1 To 5
    UF1("LC" & i).Visible = True
  Next i
  For i = 7 To 11
    UF1("LC" & i).Visible = True
  Next i
  .S_UpVisual.Visible = True
  .U_UpVisual.Visible = True
  .C_UpVisual.Visible = True
  .D_UpVisual.Visible = True
  .E_UpVisual.Visible = True
  .S_UnderVisual.Visible = True
  .U_UnderVisual.Visible = True
  .C_UnderVisual.Visible = True
  .D_UnderVisual.Visible = True
  .E_UnderVisual.Visible = True
  '直角度
  If .Youto <> "R115" Then 'QA-S-QA-DIM-154_対角線差及び直角度の測定基準
    .L10.Visible = False
    .DiagSKou.Visible = False
    .DiagUKou.Visible = False
    .DiagCKou.Visible = False
    .DiagDKou.Visible = False
    .DiagEKou.Visible = False
  End If

'描写リセットをスキップする
SkipA:
  If n = "on" Then
    UF1.Height = 562
    For i = 1 To 10
      UF1("F" & i).Visible = True
    Next i
    .OPS6.Visible = True
    .OPE6.Visible = True
    .FBad.Visible = True
    .LC6.Visible = True
    .LC12.Visible = True
    .F_UnderVisual.Visible = True
    .F_UpVisual.Visible = True
    .FMF.Visible = True
    .FF.Visible = True
    .DiagF.Visible = True
    .DiagF.Enabled = True
    .NGF.Visible = True
    .F.Visible = True
    .TSF.Visible = True
    .TEF.Visible = True
    .CheckTF.Visible = True '2025.1.23
    
    .InspeButt.Visible = False '25枚以下検査ボタン　不要
    
    .S = "4-4-S"
    .U = "4-4-E"
    .C = "4-3-2"
    .D = "4-2-2"
    .E = "4-1-S"
    .F = "4-1-E"
    
    .FS = "4-4-S"
    .FU = "4-4-E"
    .FC = "4-3-2"
    .FD = "4-2-2"
    .FE = "4-1-S"
    .FF = "4-1-E"
    
    .OPS1.Caption = "4-4-S"
    .OPS2.Caption = "4-4-E"
    .OPS3.Caption = "4-3-2"
    .OPS4.Caption = "4-2-2"
    .OPS5.Caption = "4-1-S"
    .OPS6.Caption = "4-1-E"
    
    .OPE1.Caption = "4-4-S"
    .OPE2.Caption = "4-4-E"
    .OPE3.Caption = "4-3-2"
    .OPE4.Caption = "4-2-2"
    .OPE5.Caption = "4-1-S"
    .OPE6.Caption = "4-1-E"
    
    
    Select Case UF1.AtButt.Caption
    Case "厚2ヵ所": Call AtCountSelect(4)
    Case "厚4ヵ所": Call AtCountSelect(2)
    End Select
    
  Else
    For i = 1 To 10
      UF1("F" & i).Visible = False
    Next i
    .OPS6.Visible = False
    .OPE6.Visible = False
    .FBad.Visible = False
    .LC6.Visible = False
    .LC12.Visible = False
    .F_UnderVisual.Visible = False
    .F_UpVisual.Visible = False
    .FMF.Visible = False
    .FF.Visible = False
    .DiagF.Visible = False
    .DiagF.Enabled = False
    .NGF.Visible = False
    .F.Visible = False
    .TSF.Visible = False
    .TEF.Visible = False
    .CheckTF.Visible = False '2025.1.23
    .LateralF.Visible = False
    .DiagFKou.Visible = False
    
    If UF1.極厚 Then
      .DiagS.Visible = False '対角
      .DiagU.Visible = False
      .DiagC.Visible = False
      .DiagD.Visible = False
      .DiagE.Visible = False
      
      .DiagS.Enabled = False
      .DiagU.Enabled = False
      .DiagC.Enabled = False
      .DiagD.Enabled = False
      .DiagE.Enabled = False
      .Loading.Visible = False
    Exit Sub
    End If
    
'    .S = "3-3-S"
'    .U = "3-3-E"
'    .C = "3-2-2"
'    .D = "3-1-S"
'    .E = "3-1-E"
'    .FS = "3-3-S"
'    .FU = "3-3-E"
'    .FC = "3-2-2"
'    .FD = "3-1-S"
'    .FE = "3-1-E"
'    .OPS1.Caption = "3-3-S"
'    .OPS2.Caption = "3-3-E"
'    .OPS3.Caption = "3-2-2"
'    .OPS4.Caption = "3-1-S"
'    .OPS5.Caption = "3-1-E"
'    .OPE1.Caption = "3-3-S"
'    .OPE2.Caption = "3-3-E"
'    .OPE3.Caption = "3-2-2"
'    .OPE4.Caption = "3-1-S"
'    .OPE5.Caption = "3-1-E"
  End If
  .Loading.Visible = False
End With
End Sub

Public Sub UFsizeSave(myForm As Object) '2024.8.5
UFsizeWidth = myForm.Width
UFsizeHeight = myForm.Height
UFsizeZoom = myForm.Zoom
End Sub

