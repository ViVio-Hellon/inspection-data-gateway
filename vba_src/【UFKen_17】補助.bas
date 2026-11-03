Attribute VB_Name = "【UFKen_17】補助"
Option Explicit
#If Win64 Then
  Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
  Private Declare PtrSafe Function APIBeep Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
#Else
  Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
  Private Declare Function APIBeep Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
#End If

Public MMBuf(6) As Variant    'MIN,MAX値
Public ColBuf(6) As Variant  '列番号
Public RowBuf(6) As Variant  '行番号

'強制的に最前面にさせる************************************************************************************************
#If Win64 Then
  Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
  Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
#Else
  Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
  Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
#End If

'ﾕｰｻﾞｰﾌｫｰﾑの×ﾎﾞﾀﾝを非表示にする************************************************************************************************
Public Const GWL_STYLE = -16
#If Win64 Then
  Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal WinHandle As Long, ByVal nIndex As Long) As Long
  Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal WinHandle As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal WinHandle As Long) As Long
#Else
  Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal WinHandle As Long, ByVal nIndex As Long) As Long
  Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal WinHandle As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Public Declare Function DrawMenuBar Lib "user32" (ByVal WinHandle As Long) As Long
#End If

'ｶｰｿﾙの形状変更**************************************************************************************************************************
#If Win64 Then
  Public Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
  Public Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#Else
  Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
  Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#End If

Public Const IDC_HAND = 32649& '32649& 'ハンドカーソル　32648&’ 禁止カーソル
Public mCursor As Long 'マウスカーソルの戻り値

Private Type Position
x As Long
y As Long
End Type
#If Win64 Then
  Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
  Declare PtrSafe Sub mouse_event Lib "user32" ( _
  ByVal dwFlags As Long, _
  Optional ByVal dx As Long = 0, _
  Optional ByVal dy As Long = 0, _
  Optional ByVal dwDate As Long = 0, _
  Optional ByVal dwExtraInfo As Long = 0)
  Declare PtrSafe Function GetCursorPos Lib "user32" _
  (lpPoint As Position) As Long
#Else
  Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
  Declare Sub mouse_event Lib "user32" ( _
  ByVal dwFlags As Long, _
  Optional ByVal dx As Long = 0, _
  Optional ByVal dy As Long = 0, _
  Optional ByVal dwDate As Long = 0, _
  Optional ByVal dwExtraInfo As Long = 0)
  Declare Function GetCursorPos Lib "user32" _
  (lpPoint As Position) As Long
#End If

#If VBA7 And Win64 Then
  '64ビット版
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
  '32ビット版
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' // 標準モジュール

''olepro32をoleaut32へ変更_2024.3.11
'#If VBA7 And Win64 Then
'  '64ビット版
'  ' // Declareations --------------------------------------------------
'  Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "User32.dll" (ByVal wFormat As Long) As Long
'  Private Declare PtrSafe Function OpenClipboard Lib "User32.dll" (ByVal hwnd As Long) As Long
'  Private Declare PtrSafe Function GetClipboardData Lib "User32.dll" (ByVal wFormat As Long) As Long
'  Private Declare PtrSafe Function CloseClipboard Lib "User32.dll" () As Long
'  Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef lpPictDesc As PictDesc, _
'         ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Long, ByRef IPic As IPicture) As Long
'#Else
'  '32ビット版
'  ' // Declareations --------------------------------------------------
'  Private Declare Function IsClipboardFormatAvailable Lib "User32.dll" (ByVal wFormat As Long) As Long
'  Private Declare Function OpenClipboard Lib "User32.dll" (ByVal hwnd As Long) As Long
'  Private Declare Function GetClipboardData Lib "User32.dll" (ByVal wFormat As Long) As Long
'  Private Declare Function CloseClipboard Lib "User32.dll" () As Long
'  Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef lpPictDesc As PictDesc, _
'         ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Long, ByRef IPic As IPicture) As Long
'#End If
'
'' // Types ----------------------------------------------------------
'Private Type PictDesc
'       cbSizeofStruct As Long
'       picType        As Long
'       hImage         As Long
'       Option1        As Long
'       Option2        As Long
'End Type
'Private Type GUID
'       Data1          As Long
'       Data2          As Integer
'       Data3          As Integer
'       Data4(7)       As Byte
'End Type
'' // Constants ------------------------------------------------------
'Private Const CF_BITMAP      As Long = 2
'Private Const CF_PALETTE     As Long = 9
'' Windows API定数の宣言
'Private Const CF_ENHMETAFILE As Long = 14

' Windows API定数定義
Private Const CF_BITMAP = 2
Private Const CF_PALETTE = 9
Private Const CF_ENHMETAFILE = 14

' Windows API関数宣言
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PictDesc, riid As GUID, ByVal fOwn As Long, ipic As IPicture) As Long
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PictDesc, riid As GUID, ByVal fOwn As Long, ipic As IPicture) As Long
#End If

' 構造体定義
#If VBA7 Then
    Private Type PictDesc
        cbSizeofStruct As Long
        picType As Long
        hImage As LongPtr
        Option1 As LongPtr
    End Type
#Else
    Private Type PictDesc
        cbSizeofStruct As Long
        picType As Long
        hImage As Long
        Option1 As Long
    End Type
#End If

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


Public Sub LeftClick()
mouse_event 2
mouse_event 4
End Sub
Public Sub RightClick()
mouse_event 8
mouse_event 16
End Sub
Public Sub Pos_Check()
Dim pos As Position
Call GetCursorPos(pos)
Debug.Print pos.x, pos.y
MsgBox pos.x & ", " & pos.y
End Sub
Public Sub Any_Click() '校正用
Select Case UF校正.LineName
Case "Test", "ﾄｯﾄ", "ﾊﾞﾗﾝｻｰ", "L-1", "HVC", "LVC": SetCursorPos 939, 596 '解像度1920*1080
Case Else:                                         SetCursorPos 778, 508 '解像度1600*900
End Select
LeftClick '右クリックしたいときは、LeftClickの代わりにRightClick
End Sub
Public Sub UF1_Click()
If UF1.Visible = True Then 'UF1の表示がある
  Call DefaultFormPosition 'ﾃﾞﾌｫﾙﾄ位置に戻す

  Select Case True
  Case UF1.Test, UF1.Tott, UF1.Bala, UF1.L1Work, UF1.HVC, UF1.LVC '解像度1920 * 1080_2021.07.26
    Select Case UF1.Height 'VC表示の有無
    Case 700:  SetCursorPos 1111, 487
    Case Else: SetCursorPos 990, 556
    End Select
  Case Else '解像度1600*900
    Select Case UF1.Height 'VC表示の有無
    Case 635:  SetCursorPos 760, 205
    Case Else: SetCursorPos 760, 316
    End Select
  End Select
  
  LeftClick
  
End If
End Sub
'指定したｳｨﾝﾄﾞｳをｱｸﾃｨﾌﾞにする。
Public Sub AppActiveWindow(AppTitle As String)
Dim objWindow, objShell
Dim MyName As String * 128
Dim ret As Long
Set objShell = CreateObject("Shell.Application")
For Each objWindow In objShell.Windows()
  MyName = ""
  ret = GetWindowText(objWindow.hwnd, MyName, Len(MyName))
  'Debug.Print MyName
  If MyName Like AppTitle Then
    SetForegroundWindow objWindow.hwnd
    Exit For
  End If
Next
End Sub

Public Sub WindowMenuButtonDel(UF As UserForm)
Dim WinHandle As Long: WinHandle = FindWindow("ThunderDFrame", UF.Caption) ' ウインドウハンドル取得(Excel2000)
Dim lngRet As Long: lngRet = GetWindowLong(WinHandle, GWL_STYLE) ' ウインドウ属性取得
lngRet = SetWindowLong(WinHandle, GWL_STYLE, lngRet And Not "&H80000") ' ウインドウ属性変更(システムメニュー削除)
DrawMenuBar WinHandle
End Sub

Public Sub WindowMenuButtonOn(UF As UserForm)
Dim WinHandle As Long: WinHandle = FindWindow("ThunderDFrame", UF.Caption) ' ウインドウハンドル取得(Excel2000)
Dim lngRet As Long: lngRet = GetWindowLong(WinHandle, GWL_STYLE) ' ウインドウ属性取得
lngRet = SetWindowLong(WinHandle, GWL_STYLE, lngRet Or "&H80000")  ' 戻す
DrawMenuBar WinHandle
End Sub
Public Sub MusicTrue(Temp As String) 'Cドライブにファイルがあることを前提とする
Dim Spath As String: Spath = "C:\VBAWAV\"
Dim SoundFile As String: SoundFile = Spath & Temp
If ReadCheck(SoundFile, 1) = False Then Exit Sub
Dim rc As Long: rc = mciSendString("Play " & SoundFile, "", 0, 0)
End Sub

Public Sub 音楽を流す(Temp As String)
Dim Spath As String, Sfile As String: Sfile = Temp
Select Case True
Case UF1.Test:  Spath = "C:\VBAWAV\"
Case Else:      Spath = "C:\VBAWAV\"  '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\VBAWAV\"
End Select
Dim SoundFile As String, rc As Long
If UF1.Tag = "STOP" Then Exit Sub
SoundFile = Spath & Sfile
If ReadCheck(SoundFile, 1) = False Then
'MsgBox SoundFile & vbCrLf & "がありません。", vbExclamation
Exit Sub
End If
rc = mciSendString("Play " & SoundFile, "", 0, 0)
End Sub
Public Sub 音楽異常値() 'SoundFile = "C:\VBAWAV\異常値.wav"
Dim Spath As String, Sfile As String: Sfile = "異常値.wav"
Select Case True
Case UF1.Test:  Spath = "C:\VBAWAV\"
Case Else:      Spath = "C:\VBAWAV\"  '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\VBAWAV\"
End Select
Dim SoundFile As String, rc As Long
If UF1.Tag = "STOP" Then Exit Sub
SoundFile = Spath & Sfile
If ReadCheck(SoundFile, 1) = False Then
'MsgBox SoundFile & vbCrLf & "がありません。", vbExclamation
Exit Sub
End If
rc = mciSendString("Play " & SoundFile, "", 0, 0)
End Sub

Public Sub 音楽対角() 'SoundFile = "C:\VBAWAV\chimes.wav"
Dim Spath As String, Sfile As String: Sfile = "chimes.wav"
Select Case True
Case UF1.Test:  Spath = "C:\VBAWAV\"
Case Else:      Spath = "C:\VBAWAV\"  '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\VBAWAV\"
End Select
Dim SoundFile As String, rc As Long
If ThisWorkbook.Sheets("全ヵ所検査").Visible = True Or ThisWorkbook.Sheets("キャスク検査").Visible = True Then GoTo Line2 Else GoTo Line1
Line1:
  If UF1.Tag = "STOP" Then Exit Sub
  SoundFile = Spath & Sfile
  If ReadCheck(SoundFile, 1) = False Then
  Exit Sub
  End If
  rc = mciSendString("Play " & SoundFile, "", 0, 0)
Line2:
  SoundFile = Spath & Sfile
  If ReadCheck(SoundFile, 1) = False Then
  Exit Sub
  End If
  rc = mciSendString("Play " & SoundFile, "", 0, 0)
End Sub

Public Sub 音楽合格() 'UF1以外のﾌｫｰﾑ起動時UF1を起動させないようにする
Dim Spath As String, Sfile As String: Sfile = "chime2.wav"
Select Case True
Case UF校正.Visible:  Spath = "C:\VBAWAV\"
Case UF1.Test:        Spath = "C:\VBAWAV\"
Case Else:            Spath = "C:\VBAWAV\"  '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\VBAWAV\"
End Select
Dim SoundFile As String, rc As Long

If ThisWorkbook.Sheets("全ヵ所検査").Visible = True Or ThisWorkbook.Sheets("キャスク検査").Visible = True Or UF校正.Visible = True Then GoTo Line2 Else GoTo Line1
Line1:
  If UF1.Tag = "STOP" Then Exit Sub
  SoundFile = Spath & Sfile
  If ReadCheck(SoundFile, 1) = False Then
  Exit Sub
  End If
  rc = mciSendString("Play " & SoundFile, "", 0, 0)
Line2:
  SoundFile = Spath & Sfile
  If ReadCheck(SoundFile, 1) = False Then
  Exit Sub
  End If
  rc = mciSendString("Play " & SoundFile, "", 0, 0)
End Sub
Public Sub 音楽検査員登録()
Dim Spath As String, Sfile As String: Sfile = "確認作業.wav" '"chime1.wav"
Select Case True
Case UF1.Test:  Spath = "C:\VBAWAV\"
Case Else:      Spath = "C:\VBAWAV\"  '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\VBAWAV\"
End Select
Dim SoundFile As String, rc As Long
If UF1.Tag = "STOP" Then Exit Sub
SoundFile = Spath & Sfile
If ReadCheck(SoundFile, 1) = False Then
Exit Sub
End If
rc = mciSendString("Play " & SoundFile, "", 0, 0)
End Sub
Public Sub 音楽未登録()
Dim Spath As String, Sfile As String: Sfile = "ナンバーNO.wav"
Select Case True
Case UF1.Test:  Spath = "C:\VBAWAV\"
Case Else:      Spath = "C:\VBAWAV\"  '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\VBAWAV\"
End Select
Dim SoundFile As String, rc As Long
If UF1.Tag = "STOP" Then Exit Sub
SoundFile = Spath & Sfile
If ReadCheck(SoundFile, 1) = False Then
Exit Sub
End If
rc = mciSendString("Play " & SoundFile, "", 0, 0)
End Sub
Public Sub 音楽不合格() 'SoundFile = "C:\VBAWAV\tada.wav"
Dim Spath As String, Sfile As String: Sfile = "tada.wav"
Select Case True
Case UF校正.Visible:  Spath = "C:\VBAWAV\"
Case UF1.Test:        Spath = "C:\VBAWAV\"
Case Else:            Spath = "C:\VBAWAV\" '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\VBAWAV\"
End Select
Dim SoundFile As String, rc As Long
If ThisWorkbook.Sheets("全ヵ所検査").Visible = True Or ThisWorkbook.Sheets("キャスク検査").Visible = True Or UF校正.Visible = True Then GoTo Line2 Else GoTo Line1
Line1:
  If UF1.Tag = "STOP" Then Exit Sub
  SoundFile = Spath & Sfile
  If ReadCheck(SoundFile, 1) = False Then
  Exit Sub
  End If
  rc = mciSendString("Play " & SoundFile, "", 0, 0)
Line2:
  SoundFile = Spath & Sfile
  If ReadCheck(SoundFile, 1) = False Then
  Exit Sub
  End If
  rc = mciSendString("Play " & SoundFile, "", 0, 0)
End Sub
Public Function Conv(Trm As String) As String '変換
Dim a As String, b As String
  Select Case Maker(Trm)
  Case "DT" 'DT10002+0000020.712M
   Conv = Mid(Trm, 9, 11)
  Case "#L" '＃L395.8
   Conv = Mid(Trm, 3, Len(Trm))
  Case "#〆" ''手入力_2021.02.23 @を#に変更
   Conv = Mid(Trm, 3, Len(Trm))
  Case Else
   Conv = Trm
  End Select
End Function
Public Function Maker(Trm As String) As String 'DT #L選別
  Maker = ""
  If Trm = "" Then Exit Function
  If Len(Trm) >= 3 And Trm <> "" Then
  Dim TnmCon As String: TnmCon = StrConv(UCase(Left(Trm, 2)), vbNarrow) '強制大文字変換
    Select Case TnmCon
    Case "DT" 'DT10002+0000021.712M
     Maker = "DT"
    Case "#L" '#L395.8
     Maker = "#L"
    Case "#〆" '〆手入力
     Maker = "#〆"
    End Select
  End If
End Function
'ﾌｫﾙﾀﾞ・ﾌｧｲﾙの存在確認'※ﾈｯﾄﾜｰｸ上のﾌｧｲﾙ･ﾌｫﾙﾀﾞをﾁｪｯｸする際、Dir関数では存在しない場合ｴﾗｰが発生する。
Public Function PathCheck(path As String) As Boolean
'Dim fso As New FileSystemObject '(参照設定)Microsoft Scripting Runtime
Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
Select Case GetFileSystemType(path) 'ﾁｪｯｸｶﾃｺﾞﾘ(1:ﾌｧｲﾙ, 2:ﾌｫﾙﾀﾞ)
  Case 1: If FileLen(path) >= 0 Then PathCheck = True 'ある(ﾌｧｲﾙｻｲｽﾞ追加)
  Case 2: PathCheck = True 'ある
  Case 3: PathCheck = False 'ない
End Select
Set fso = Nothing
End Function
'ﾌｫﾙﾀﾞ・ﾌｧｲﾙを見分ける
Public Function GetFileSystemType(path As String) As Long
'Dim fso As New FileSystemObject '(参照設定)Microsoft Scripting Runtime
Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
Select Case True
  Case fso.fileExists(path):   GetFileSystemType = 1
  Case fso.FolderExists(path): GetFileSystemType = 2
  Case Else: GetFileSystemType = 3
End Select
Set fso = Nothing
End Function
'指定のﾌﾞｯｸが開かれているか確認する
Public Function IsBookOpen(FilePath As String) As Boolean
On Error Resume Next
Open FilePath For Append As #1
Close #1
If Err.Number > 0 Then
 IsBookOpen = True
Else
 IsBookOpen = False
End If
End Function
'任意のﾜｰｸｼｰﾄの存在ﾁｪｯｸ
Public Function WsCheck(ByRef WsName As String) As Boolean
Dim ws As Worksheet 'ﾜｰｸｼｰﾄ
WsCheck = False
For Each ws In ThisWorkbook.Worksheets
  If ws.Name = WsName Then
    WsCheck = True
    Exit For
  End If
Next ws
Set ws = Nothing
End Function
'ﾌｫﾙﾀﾞ・ﾌｧｲﾙの存在確認
Public Function ReadCheck(ReadTarget As String, n As Long) As Boolean
On Error Resume Next
'※ﾈｯﾄﾜｰｸ上のﾌｧｲﾙ･ﾌｫﾙﾀﾞをﾁｪｯｸする際、Dir関数では存在しない場合ｴﾗｰが発生する。
Dim fso As New FileSystemObject '参照設定必須
ReadCheck = False 'ない(初期化)
Select Case True 'ﾁｪｯｸｶﾃｺﾞﾘ(1:ﾌｧｲﾙ, 2:ﾌｫﾙﾀﾞ)
  Case n = 1 And fso.fileExists(ReadTarget)
    If FileLen(ReadTarget) >= 0 Then ReadCheck = True 'ある(ﾌｧｲﾙｻｲｽﾞ追加)
  Case n = 2 And fso.FolderExists(ReadTarget): ReadCheck = True 'ある
End Select
End Function
'指定の名前のワークシートを追加
Public Sub AddWS(ByRef WsName As String)
Application.DisplayAlerts = False
Dim ws As Worksheet 'ﾜｰｸｼｰﾄ
For Each ws In ThisWorkbook.Worksheets
  If ws.Name = WsName Then
    ThisWorkbook.Sheets(WsName).Visible = xlSheetVisible '_BeforeUpdate直後にUserForm_QueryCloseが実行されるとエラー_2021.10.26
    ThisWorkbook.Worksheets(WsName).Delete
    Exit For
  End If
Next ws
ThisWorkbook.Worksheets.Add(before:=Worksheets(1)).Name = WsName 'before:=Worksheets(1)左端に追加　'Sheets(Sheets.Count)'右端に追加
Select Case WsName
Case "A_Input":    ThisWorkbook.Sheets(WsName).Visible = True
Case "A_InputSa":  ThisWorkbook.Sheets(WsName).Visible = True
Case Else:         ThisWorkbook.Sheets(WsName).Visible = False
End Select
Set ws = Nothing 'ｵﾌﾞｼﾞｪｸﾄの解放
Application.DisplayAlerts = True
End Sub
Public Function ShJuge(Temp As String) As Worksheet 'なければ作り非表示
If WsCheck(Temp) = False Then
 Call AddWS(Temp)
End If
'ThisWorkbook.Sheets(Temp).Visible = False
Set ShJuge = ThisWorkbook.Sheets(Temp)
End Function
'選択した範囲は空かどうか。データが無い場合は中断
Public Sub EmptyRangeJudge(Optional tgtRange As Range, Optional msg As String = "空のWorkSheet")
If Application.WorksheetFunction.CountA(tgtRange) = 0 Then
'  MsgBox msg & "を参照しています。中断します。", vbExclamation
'  End
End If
End Sub
'最終行の取得
Public Function UR(Optional sh As Worksheet, Optional n As Long = 0) As Long
If sh Is Nothing Then Set sh = ActiveSheet
Dim tgtRange As Range
Call EmptyRangeJudge(sh.UsedRange)
Dim lDummy As Long '仕様上のバグ対策(xlLastCellがﾘｾｯﾄされない)
Select Case n
  Case 0 '(含む：書式設定、非表示行)
  
    Application.EnableEvents = False 'ｲﾍﾞﾝﾄSelectChange抑制
    lDummy = sh.UsedRange.row 'UsedRangeを無駄に使用して初期化
    sh.Unprotect
    UR = sh.Cells(1, 1).SpecialCells(xlLastCell).row
    Application.EnableEvents = True
  Case Else '(含まない：書式設定、非表示行)
    UR = sh.Cells(Rows.Count, n).End(xlUp).row
    Set tgtRange = Range(sh.Cells(1, n), sh.Cells(UR, n))
    Call EmptyRangeJudge(tgtRange, "空の列")
End Select
End Function
'最終列の取得
Public Function UC(Optional sh As Worksheet, Optional n As Long = 0) As Long
If sh Is Nothing Then Set sh = ActiveSheet
Call EmptyRangeJudge(sh.UsedRange)
Dim lDummy As Long '仕様上のバグ対策(xlLastCellがﾘｾｯﾄされない)
Dim tgtRange As Range
Select Case n
  Case 0 '(含む：書式設定、非表示列)
    Application.EnableEvents = False 'ｲﾍﾞﾝﾄSelectChange抑制
    lDummy = sh.UsedRange.Column 'UsedRangeを無駄に使用して初期化
    UC = sh.Cells(1, 1).SpecialCells(xlLastCell).Column
    Application.EnableEvents = True
  Case Else '(含まない：書式設定、非表示列)
    UC = sh.Cells(n, Columns.Count).End(xlToLeft).Column
    Set tgtRange = Range(sh.Cells(1, 1), sh.Cells(n, UC))
    Call EmptyRangeJudge(tgtRange, "空の行")
End Select
End Function
'ﾌｧｲﾙ配列変換（1次元配列）
Public Function GetOneDimData(fileName As String) As Variant
Dim FreeNum As Long: FreeNum = FreeFile
Dim buf() As Byte 'Binary読込用全ﾃﾞｰﾀ
If Not PathCheck(fileName) Then Exit Function 'ﾌｧｲﾙｻｲｽﾞ確認
Open fileName For Binary As #FreeNum 'ﾊﾞｲﾅﾘ形式でﾌｧｲﾙを一括読み込み
ReDim buf(LOF(1))
Get #FreeNum, , buf
Close #FreeNum
'Unicodeに変換し改行ｺｰﾄﾞ毎に区切って配列化（ﾃﾞｰﾀを1行毎に区切る）'ﾃﾞｰﾀを配列で返す。
GetOneDimData = Split(StrConv(buf(), vbUnicode), vbCrLf, , vbBinaryCompare)
End Function
'ファイルを開く
Public Function xlsFileOpenWb(FilePath As String, Optional MODE As Boolean = True) As Object '"ReadOnly"
Dim fileName As String: fileName = GetFileName(FilePath)
If IsBookOpen(fileName) Then Exit Function
Workbooks.Open FilePath, ReadOnly:=MODE
Set xlsFileOpenWb = ActiveWorkbook
End Function
'ファイルを閉じる
Public Sub xlsFileClose(tgtWB As Workbook, Optional saveflag As Boolean = False) '"ReadOnly"
If Not IsBookOpen(tgtWB.Name) Then Exit Sub
Application.DisplayAlerts = False
If saveflag Then tgtWB.Save
tgtWB.Close
Application.DisplayAlerts = True
End Sub
'ﾌﾙﾊﾟｽからﾌｧｲﾙ名を抜き出す
Public Function GetFileName(PathName As String) As String
Dim Temp As Variant: Temp = Split(PathName, "\")
GetFileName = Temp(UBound(Temp))
End Function
'指定WSデータの2次元配列化
Public Function shDataGet(shName As String, Optional k As Long = 0) As Variant
Dim sh As Worksheet: Set sh = Worksheets(shName)
shDataGet = Range(sh.Cells(1, 1), sh.Cells(UR(sh), UC(sh, k)))
End Function
'指定WSデータの2次元配列化
Public Function UseWs(sh As Worksheet, Optional k As Long = 0) As Variant
UseWs = Range(sh.Cells(1, 1), sh.Cells(UR(sh), UC(sh, k)))
End Function

Public Sub DisplaySheets() 'ｼｰﾄ全表示
Dim i As Long
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
With ThisWorkbook
For i = 1 To .Sheets.Count
  If (.Sheets(i).Name) = "保存用" Or _
  (.Sheets(i).Name) = "印刷用" Or (.Sheets(i).Name) = "校正記録" Or _
  (.Sheets(i).Name) = "Lot情報履歴" Or (.Sheets(i).Name) = "品証検証" Or _
  (.Sheets(i).Name) = "WorkerTime" & "_" & Lname Or (.Sheets(i).Name) = "Daily" & "_" & Lname Or _
  (.Sheets(i).Name) = "全ヵ所検査保存" Or (.Sheets(i).Name) = "キャスク保存" Or _
  (.Sheets(i).Name) = "ｺﾋﾟ-用" Or (.Sheets(i).Name) = "転送済" Or (.Sheets(i).Name) = "作業保存" & "_" & Lname Or _
  (.Sheets(i).Name) = "導電率測定" Or (.Sheets(i).Name) = "不良要因印刷" & "_" & Lname Then .Sheets(i).Visible = xlSheetVisible '導電率測定_2021.11.25追加
Next i
End With
End Sub
Public Sub KenSheets() '検査履歴系のみ全表示
Dim i As Long
With ThisWorkbook
For i = 1 To .Sheets.Count
  If Not ((.Sheets(i).Name) = "保存用" Or (.Sheets(i).Name) = "印刷用" Or (.Sheets(i).Name) = "校正記録" _
  Or (.Sheets(i).Name) = "Lot情報履歴" Or (.Sheets(i).Name) = "品証検証" _
  Or (.Sheets(i).Name) = "全ヵ所検査保存" Or (.Sheets(i).Name) = "キャスク保存" _
  Or (.Sheets(i).Name) = "VC一覧" Or (.Sheets(i).Name) = "LOT一覧" Or (.Sheets(i).Name) = "ｷｬｽｸ作業中保存" _
  Or (.Sheets(i).Name) = "VC厚" Or (.Sheets(i).Name) = "5N" _
  Or (.Sheets(i).Name) = "キャスク検査" Or (.Sheets(i).Name) = "キャスク原紙" Or (.Sheets(i).Name) = "三菱受託原紙" _
  Or (.Sheets(i).Name) = "全ヵ所検査" Or (.Sheets(i).Name) = "全ヵ所検査用原紙" Or (.Sheets(i).Name) = "5N原紙" _
  Or (.Sheets(i).Name) = "印刷用原紙" Or (.Sheets(i).Name) = "図" Or (.Sheets(i).Name) = "名簿" _
  Or (.Sheets(i).Name) = "Lot情報履歴" Or (.Sheets(i).Name) = "不良要因印刷原紙" _
  Or (.Sheets(i).Name) = "H494_原紙" Or (.Sheets(i).Name) = "H494検査") Then '不良要因印刷原紙_2023.5.26追加 _2023.10.16 H494_原紙 H494検査
  .Sheets(i).Visible = xlSheetVisible
  End If
Next i
End With
End Sub
Public Sub HiddenSheets(Optional n As String = "") 'ｼｰﾄ非表示_2021.03.15非表示しなくてよいものﾋﾟｯｸｱｯﾌﾟ
Dim i As Long
With ThisWorkbook
  Select Case n
  Case "" 'フォーム起動後
    For i = 1 To .Sheets.Count
      If Not (.Sheets(i).Name = "Main") Then .Sheets(i).Visible = False
    Next i
  Case Else 'フォームを閉じる前_2021.03.17
    Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
    Application.DisplayAlerts = False '不要シートは消す_2021.03.17
      If WsCheck("Daily_") Then ThisWorkbook.Sheets("Daily_").Delete
      If WsCheck("WorkerTime_") Then ThisWorkbook.Sheets("WorkerTime_").Delete
      If WsCheck("検査員_") Then ThisWorkbook.Sheets("検査員_").Delete
      If WsCheck("作業保存_") Then ThisWorkbook.Sheets("作業保存_").Delete
    Application.DisplayAlerts = True
    For i = 1 To .Sheets.Count '
      If Not (.Sheets(i).Name = "Main" Or .Sheets(i).Name = "Test" Or _
       .Sheets(i).Name = "高さチェック" Or .Sheets(i).Name = "ﾛｼﾞ" Or _
      (.Sheets(i).Name) = "Lot情報履歴" Or (.Sheets(i).Name) = "VC一覧" Or (.Sheets(i).Name) = "LOT一覧") Then .Sheets(i).Visible = xlVeryHidden '見えなくする(解除はコードでのみ)
    Next i
    For i = 1 To .Sheets.Count
      If (.Sheets(i).Name = "Main" Or .Sheets(i).Name = "Test" Or _
      .Sheets(i).Name = "高さチェック" Or .Sheets(i).Name = "ﾛｼﾞ" Or _
      (.Sheets(i).Name) = "Lot情報履歴" Or (.Sheets(i).Name) = "VC一覧" Or (.Sheets(i).Name) = "LOT一覧" Or _
      (.Sheets(i).Name) = "キャスク検査" Or (.Sheets(i).Name) = "全ヵ所検査" Or (.Sheets(i).Name) = "H494検査") Then .Sheets(i).Visible = False '非表示
    Next i
    
  '  For i = 1 To .Sheets.Count
  '    If (.Sheets(i).Name = "Main" Or .Sheets(i).Name = "Test" Or _
  '    .Sheets(i).Name = "5" Or .Sheets(i).Name = "高さチェック" Or .Sheets(i).Name = "ﾛｼﾞ" Or _
  '    (.Sheets(i).Name) = "Lot情報履歴" Or (.Sheets(i).Name) = "全ヵ所検査保存" Or (.Sheets(i).Name) = "キャスク保存" Or _
  '    (.Sheets(i).Name) = "VC一覧" Or (.Sheets(i).Name) = "LOT一覧" Or (.Sheets(i).Name) = "ｷｬｽｸ作業中保存" Or (.Sheets(i).Name) = "5N" Or _
  '    (.Sheets(i).Name) = "キャスク検査" Or (.Sheets(i).Name) = "キャスク原紙" Or (.Sheets(i).Name) = "三菱受託原紙" Or _
  '    (.Sheets(i).Name) = "全ヵ所検査" Or (.Sheets(i).Name) = "全ヵ所検査用原紙" Or (.Sheets(i).Name) = "5N原紙") Then .Sheets(i).Visible = False'下＊＊＊＊’.Sheets(i).Visible = xlVeryHidden’上
  '  Next i
  End Select
End With
End Sub
Public Sub VisibleSheets() 'ｼｰﾄ表示
Dim i As Long
With ThisWorkbook
  For i = 1 To .Sheets.Count
    .Sheets(i).Visible = xlSheetVisible
  Next i
End With
End Sub
Public Sub Numlock_onoff() 'Numlock対策
Dim Wshshell
Set Wshshell = CreateObject("WScript.Shell")
Wshshell.SendKeys "{NUMLOCK}"
Set Wshshell = Nothing
End Sub
'2025.7.23
Public Sub ShowProgress(message As String)
UFProgress.Show vbModeless
UFProgress.ProgressBar1.Visible = False
UFProgress.Label1.Visible = False
UFProgress.ProgressLabel.Visible = True
' プログレス表示用のラベルがある場合
If UFProgress.Controls.Count > 0 Then
    On Error Resume Next
    UFProgress.Controls("ProgressLabel").Caption = message
    UFProgress.Controls("ProgressLabel").Visible = True
    On Error GoTo 0
End If

' ステータスバーにも表示
Application.StatusBar = message
DoEvents
End Sub
'2025.7.23
Public Sub HideProgress()
On Error Resume Next
UFProgress.Controls("ProgressLabel").Visible = False
On Error GoTo 0
Application.StatusBar = False

Unload UFProgress
End Sub

Public Sub ProgressBar_count(x As Long, Count As Long) 'CountをMaxになった時100%
If UFProgress.Visible = False Then Exit Sub
Dim 秒 As Long: 秒 = Second(Now)
Dim a As Long, Percent As Long
UFProgress.ProgressBar1.Min = 1 'プログレスバーの最小値を設定
UFProgress.ProgressBar1.Max = 1000  'プログレスバーの最大値を設定
UFProgress.ProgressBar1.Value = 1 'プログレスバーの現在値を設定
  For a = x To Count
    If UFProgress.ProgressBar1.Min < a And UFProgress.ProgressBar1.Max >= a Then 'プログレスバーの値表示を更新
      Percent = CInt(a / 1000 * 100) 'プログレスバーのLabel表示を更新
      UFProgress.Label1.Caption = Percent & "%完了" '& "_" & 秒 & "秒"
      UFProgress.ProgressBar1.Value = a 'プログレスバーの値を更新
      DoEvents '滞留処理を実行
    End If
  Next a
End Sub
'PDFファイルを作成する ｼｰﾄの印刷範囲設定が反映される
'Public Sub PDFCREATE(PDFFileName As String, sh1 As Worksheet, fname As String) 'パス,保存シート,保存ファイル名
'Dim fileName As String '保存先フォルダパス＆ファイル名
'fileName = PDFFileName & "\" & fname & ".pdf"
'sh1.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName, Quality:=xlQualityStandard, From:=1, To:=1
''MsgBox ThisWorkbook.Path & vbCrLf & "PDFファイルに保存しました。", vbInformation
'End Sub
Public Sub PDFCREATE(PDFFileName As String, sh1 As Worksheet, fname As String)
Dim FilePath As String
FilePath = PDFFileName & "\" & fname & ".pdf"

' 既存ファイルがある場合削除（競合回避）
If Dir(FilePath) <> "" Then Kill FilePath

' PDF作成処理
On Error Resume Next
sh1.ExportAsFixedFormat Type:=xlTypePDF, fileName:=FilePath, Quality:=xlQualityStandard, From:=1, To:=1
On Error GoTo 0

' PDFが完全に作成されるまで待機（最大5秒）
Dim i As Integer
For i = 1 To 50 ' 100ミリ秒 × 50 = 5秒
  If Dir(FilePath) <> "" Then Exit For
  Sleep 100 ' 100ミリ秒（0.1秒）待機
  DoEvents ' 他の処理を実行
Next i

' PDFが作成されたか最終確認
If Dir(FilePath) = "" Then
  MsgBox "PDFファイルが作成されませんでした: " & FilePath, vbExclamation
End If
End Sub

Public Function NewFileName(SetPath As String) As String
Dim F As String, FileTime As String
'SetPath = "C:\検査履歴\2021.01\"
F = GetNewFileName(SetPath, Array("*.txt", "*.csv", "*.doc", "*.xlsx", "*.pdf"))
FileTime = FileDateTime(SetPath & "\" & F)
NewFileName = F
'MsgBox SetPath & "  " & f & "  " & FileTime
End Function

'指定フォルダ中の指定ファイル(ワイルドカード名、配列可)から最新のファイル名を返す、無ければ""
Public Function GetNewFileName(TargetPath As String, Fnames As Variant) As String
Dim fso As Object 'FileSystemObject
Dim FL As Object 'Folder
Dim F As Object 'File
Dim myFile As Object 'File
Dim myDate As Date '更新日時の最新値
Dim Flg As Boolean 'Fnamesに一致したかどうかのフラグ
Dim Fn As Variant

Set fso = CreateObject("Scripting.FileSystemObject")
Set FL = fso.GetFolder(TargetPath)

For Each F In FL.Files
  '指定ファイルか確認
  Flg = False
  '配列で指定されている時
    If IsArray(Fnames) Then
      For Each Fn In Fnames
          If LCase(F.Name) Like LCase(Fn) Then
            Flg = True
            Exit For
          End If
      Next
    Else
      '単一指定の時
        If LCase(F.Name) Like LCase(Fnames) Then
          Flg = True
        End If
    End If
  '更新日が最新のファイルを抽出（作成日なら、.DateCreated）
    If Flg Then
      If myFile Is Nothing Then
        Set myFile = F
        myDate = myFile.DateLastModified
      ElseIf F.DateLastModified > myDate Then
        Set myFile = F
        myDate = myFile.DateLastModified
      End If
    End If
Next
  If Not myFile Is Nothing Then
    GetNewFileName = myFile.Name
  Else
    GetNewFileName = ""
  End If
Set myFile = Nothing
Set F = Nothing
Set FL = Nothing
Set fso = Nothing

End Function
Public Sub BookSaClose(BkPath As String, BkName As String) '拡張子含む_2021.01.29
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Dim path As String: path = BkPath & "\" & BkName
If IsBookOpened(path) Then
  Dim wb As Workbook: Set wb = Workbooks(BkName)
  wb.Close SaveChanges:=True '保存して閉じる
End If
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Public Sub AllClose(Optional Aname As String = "") '指定した名前以外のブックを閉じる
Dim BK As Workbook
Dim Bname As String: Bname = ThisWorkbook.Name
For Each BK In Workbooks
  If Not (BK.Name = Aname Or BK.Name = Bname) Then
    Application.DisplayAlerts = False
    BK.Close SaveChanges:=True
    Application.DisplayAlerts = True
  End If
Next
End Sub
Function IsBookOpened(FilePath) As Boolean
On Error Resume Next
Open FilePath For Append As #1 '保存済みのブックか判定
Close #1

  If Err.Number > 0 Then '既に開かれている場合
    IsBookOpened = True
  Else ' 開かれていない場合
    IsBookOpened = False
  End If
End Function
Public Sub CheckBookOpened()
Dim FilePath
Dim ret As Boolean
'filepath = OfficeList.Terget & OfficeList.Pass
ret = IsBookOpened(FilePath) ' 開いているかチェック
On Error Resume Next
  If (ret = True) Then
    Debug.Print "開いている"
    Exit Sub
  Else
    Debug.Print "開いていない"
  End If
End Sub
Public Sub IMEOFF()
Application.ScreenUpdating = False
Call AddWS("IMEOFF")
Dim ImeSh As Worksheet: Set ImeSh = ThisWorkbook.Sheets("IMEOFF")
  With ImeSh.Range("A1").Validation
   .Add Type:=xlValidateInputOnly
   .IMEMode = xlIMEModeOff
  End With
Application.DisplayAlerts = False
  ImeSh.Delete
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
'LISTBOX：ListBoxの配列データを2次元配列化
Public Function ListBoxDBArr(objDb As MSForms.ListBox, Optional i_Max As Long = 0, Optional j_Max As Long = 0) As Variant

i_Max = IIf(i_Max <> 0, i_Max, objDb.ListCount)
j_Max = IIf(j_Max <> 0, j_Max, objDb.ColumnCount)

Dim NewArr() As Variant
ReDim NewArr(0 To i_Max, 0 To j_Max) As Variant '新規データ枠

Dim i As Long, j As Long
For i = 1 To i_Max
  For j = 1 To j_Max
    NewArr(i, j) = objDb.List(i, j)
  Next j
Next i

ListBoxDBArr = NewArr

End Function
'LISTBOX：2次元配列をListBoxに格納
Public Sub setListBoxArr(tmpLis As MSForms.ListBox, tMP As Variant)
Dim i As Long: i = Get_ArrayDim(tMP)
Select Case i
Case 2
  Dim LR As Long: LR = UBound(tMP, 1)
  Dim LC As Long: LC = UBound(tMP, 2)
Case 1
  LC = 1
Case Else
End Select
With tmpLis
  .clear
  .ColumnCount = LC
  .List = tMP
  .ColumnWidths = "35;35;35;35;35;35;35;35;35;35;35;35;35;35;35;35;35;35;35;35;35;35;35;35;35;35;35"
End With
End Sub
'LISTBOX：ListBoxの配列データをセパレート入り文字列に変換
Public Function GetListBoxDataStr(tgtctrl As MSForms.ListBox, Optional stNum As Long = 0, Optional Spl1st As String = "|", Optional Spl2nd As String = ";") As String
Dim i As Long
Dim k As Long
Dim buf As Variant: buf = tgtctrl.List
Dim tMP As String: tMP = ""
For k = stNum To UBound(buf, 1) '行
  tMP = tMP & Spl1st
  For i = stNum To UBound(buf, 2) '列
    tMP = tMP & Spl2nd & buf(k, i)
  Next i
Next k
GetListBoxDataStr = tMP
End Function
' 配列の次元数を求める
' 引数1 : vlArray / 対象の変数' 戻り値 : [Integer] / -1:配列ではない, 0:初期化されていないか要素無し,1 以上: 配列の次元数
Function Get_ArrayDim(ByVal vlArray As Variant) As Long
Dim i As Long
Dim varTmp As Variant
i = 0
On Error GoTo ErrLine
If Not IsArray(vlArray) Then GoTo ErrLine
  Do
    i = i + 1
    varTmp = UBound(vlArray, i)
    If varTmp = -1 Then GoTo ErrLine
  Loop
ErrLine:
If Err.Number <> 0 Then Err.clear
Get_ArrayDim = i - 1
End Function
'二次元配列に並べ替え
Public Sub merge_sort2(ByRef arr As Variant, ByVal col As Long) 'うまく書かれたクイックソートの1.5倍程度処理時間がかかる、メモリ使用量はクイックソートの3倍程度
  Dim irekae As Variant
  Dim indexer As Variant
  Dim Tmp1() As Variant
  Dim Tmp2() As Variant
  Dim i As Long
  ReDim irekae(LBound(arr, 1) To UBound(arr, 1))
  ReDim indexer(LBound(arr, 1) To UBound(arr, 1))
  ReDim Tmp1(LBound(arr, 1) To UBound(arr, 1))
  ReDim Tmp2(LBound(arr, 1) To UBound(arr, 1))
  For i = LBound(arr, 1) To UBound(arr, 1) Step 2
      If i + 1 > UBound(arr, 1) Then
          irekae(i) = arr(i, col)
          indexer(i) = i
          Exit For
      End If
      If arr(i + 1, col) < arr(i, col) Then
          irekae(i) = arr(i + 1, col)
          irekae(i + 1) = arr(i, col)
          indexer(i) = i + 1
          indexer(i + 1) = i
      Else
          irekae(i) = arr(i, col)
          irekae(i + 1) = arr(i + 1, col)
          indexer(i) = i
          indexer(i + 1) = i + 1
      End If
  Next
  Dim st1 As Long
  Dim en1 As Long
  Dim st2 As Long
  Dim en2 As Long
  Dim n As Long
  i = 1
  Do While i * 2 <= UBound(arr, 1)
      i = i * 2
      n = 0
      Do While en2 + i - 1 < UBound(arr, 1)
          n = n + 1
          st1 = i * 2 * (n - 1) + LBound(arr, 1)
          en1 = i * 2 * (n - 1) + i - 1 + LBound(arr, 1)
          st2 = en1 + 1
          en2 = IIf(st2 + i - 1 >= UBound(arr, 1), UBound(arr, 1), st2 + i - 1)
          Call merge2(irekae, indexer, Tmp1, Tmp2, st1, en1, st2, en2)
      Loop
      en2 = 0
  Loop
  Dim ret As Variant
  ReDim ret(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))
  For i = LBound(arr, 1) To UBound(arr, 1)
      For n = LBound(arr, 2) To UBound(arr, 2)
          If IsObject(arr(indexer(i), n)) Then
              Set ret(i, n) = arr(indexer(i), n)
          Else
              ret(i, n) = arr(indexer(i), n)
          End If
      Next
  Next
  arr = ret
End Sub

Public Sub merge2(ByRef irekae As Variant, _
ByRef indexer As Variant, _
ByRef tmpArr() As Variant, _
ByRef tmpIndexer() As Variant, _
ByVal st1 As Long, _
ByVal en1 As Long, _
ByVal st2 As Long, _
ByVal en2 As Long)
    Dim j As Long
    Dim n As Long
    Dim i As Long
    For i = st1 To en2
        tmpArr(i) = irekae(i)
        tmpIndexer(i) = indexer(i)
    Next
    j = st1
    n = st2
    Do While (j < en1 + 1 Or n < en2 + 1)
        If n >= en2 + 1 Then
            irekae(j + n - st2) = tmpArr(j)
            indexer(j + n - st2) = tmpIndexer(j)
            j = j + 1
        ElseIf j < en1 + 1 And tmpArr(j) <= tmpArr(n) Then
            irekae(j + n - st2) = tmpArr(j)
            indexer(j + n - st2) = tmpIndexer(j)
            j = j + 1
        Else
            irekae(j + n - st2) = tmpArr(n)
            indexer(j + n - st2) = tmpIndexer(n)
            n = n + 1
        End If
    Loop
End Sub

Public Function TimeDesignation() As Boolean
'trueで指定の時間内
Dim path As String: path = NextPath
Dim 時 As Long: 時 = Hour(Now)
Dim 分 As Long: 分 = Minute(Now)
Dim 秒 As Long: 秒 = Second(Now)
If (5945 < 分 & 秒) Then   ' And (145 > 分 & 秒)
  TimeDesignation = True
  If TimeDesignation Then Call WriteLog(分 & ":" & 秒 & "_時間ﾁｪｯｸ", "", path)
ElseIf (1 > 分) And (45 >= 秒) Then
  TimeDesignation = True
  If TimeDesignation Then Call WriteLog(分 & ":" & 秒 & "_時間ﾁｪｯｸ", "", path)
End If
End Function

Public Function FileDesignation() As Boolean
'trueで更新中Flagファイルが消えた状態
Dim fileName As String: fileName = "PLEASE WAIT.txt"
Dim FilePath As String: FilePath = MainPath

Dim path As String: path = NextPath
Dim i As Long, flag As Boolean
i = 0
If ReadCheck(FilePath & fileName, 1) Then flag = True '在る
If flag Then Call WriteLog(fileName & "_ﾁｪｯｸ*在る", "", path)  '消えていない
If Not flag Then FileDesignation = True
End Function

Public Function FileExistence(fileName As String, FilePath As String, Lname As String) As Boolean 'テキストファイルがなくなるまでチェック
'trueで更新中Flagファイルが消えた状態
Dim path As String: path = NextPath
Dim i As Long
Dim flag As Boolean
i = 0
UFProgress.Show vbModeless
UFProgress.CommandButton1.Caption = "DB更新中"
UFProgress.Tag = "1"
Do Until ReadCheck(FilePath & fileName, 1) = False '条件を満たす迄
'1000 が 1秒
  Sleep 100
  i = i + 5
  If i = 1000 Then '20秒
    flag = True
    Exit Do
  End If
  Call ProgressBar_count(i, i + 5)
Loop
If flag Then
  MsgBox "元ファイル形成中、1分後リトライしてください。"
  If flag Then Call WriteLog(fileName & "_" & Lname, "", path)
End If
If Not flag Then FileExistence = True
Unload UFProgress
End Function

Public Sub 半角へ変換() '数値のみを半角へ変換
Dim i As Long, n As Long, LR As Long, LC As Long
Dim saveSh As Worksheet: Set saveSh = ThisWorkbook.Sheets("Sheet1")
LR = UR(saveSh, 1)
LC = UC(saveSh, 1)
Dim SaveR As Range:      Set SaveR = Range(saveSh.Cells(2, 1), saveSh.Cells(LR, 2))
Dim SaveTemp():          SaveTemp = SaveR 'シートを配列へ

For n = LBound(SaveTemp, 1) To UBound(SaveTemp, 1)
  For i = 1 To Len(SaveTemp(n, 1)) '数値のみを半角へ変換
    If Mid(SaveTemp(n, 1), i, 1) Like "[０-９]" Then
      Mid(SaveTemp(n, 1), i, 1) = StrConv(Mid(SaveTemp(n, 1), i, 1), vbNarrow)
    End If
  Next i
  SaveTemp(n, 1) = numExtract(SaveTemp(n, 1)) '数字だけ抽出
  If SaveTemp(n, 1) <> "" Then saveSh.Cells(n + 1, 3) = val(SaveTemp(n, 1))
  saveSh.Cells(n + 1, 3).NumberFormatLocal = "0.0_ "
  SaveTemp(n, 2) = numExtract(SaveTemp(n, 2))
  If SaveTemp(n, 2) <> "" Then saveSh.Cells(n + 1, 4) = val(SaveTemp(n, 2))
  saveSh.Cells(n + 1, 4).NumberFormatLocal = "0.0_ "
Next n
'Stop
End Sub
Public Function ConvertToHalf(Temp As Variant) As String '数値のみを半角へ変換
Dim i As Long
For i = 1 To Len(Temp)
  If Mid(Temp, i, 1) Like "[０-９]" Then
    Mid(Temp, i, 1) = StrConv(Mid(Temp, i, 1), vbNarrow)
  End If
Next i
ConvertToHalf = numExtract(Temp) '数字だけ抽出
'Stop
End Function

'文字列から数字だけ抽出
Public Function numExtract(StringValue As Variant) As String
 Dim i As Integer
 Dim numText As String
 Dim Temp
 If InStr(StringValue, "/") > 0 Then '1.0/2M スラッシュ前の数字を反映
  Temp = Split(StringValue, "/")
  StringValue = Temp(LBound(Temp))
 End If
' Stop
 For i = 1 To Len(StringValue) '1 to 文字列（最終文字）まで For Nextで影響させる
    numText = Mid(StringValue, i, 1) 'Midで文字列を左から順にnumTextに格納
    If numText Like "[0-9.]" Then: numExtract = numExtract & numText 'numText[0-9]に該当する場合 変数numExtractに格納していく**小数点含む
 Next i
End Function
'配列が空かどうかのジャッジ
Public Function IsEmptyArray(a As Variant) As Boolean 'trueの時配列が空 *エラー発生=配列が空 ※変数なしチェック_0配列でﾃﾞｰﾀ1個の時はfalseになる *=追加でtrueにする様にした
'(例)_ a = Split(Cells(1, 3), "_") cellがからの場合変数なしになる
On Error GoTo ERROR_
   If (0 <= UBound(a, 1)) Then
      IsEmptyArray = False
   Else
      IsEmptyArray = True
   End If
   Exit Function
ERROR_:
'エラー発生時
 IsEmptyArray = True
End Function

'2023.5.25_Log　　2026.05.14_:を|に変更
Public Sub LogOut(La As MSForms.TextBox, Optional AorB As String = "", Optional Mitsu As String = "") '歪は引数2に文字列でA,B面かわかるようにできる’引数3に三菱かの文字列
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim fname As String: fname = "Measured.log"
Dim path As String: path = HistoryPath
Dim Lpath As String, TB As String

Dim WorkerName As String: WorkerName = GetWorkerName()

Select Case Lname
Case "L-1":    Lpath = path & "L-1\L-1_Log"
Case "LVC":    Lpath = path & "LVC\LVC_Log"
Case "HVC":    Lpath = path & "HVC\HVC_Log"
Case "ﾊﾞﾗﾝｻｰ": Lpath = path & "バランサー\バランサー_Log"
Case "ﾄｯﾄ":    Lpath = path & "トット\トット_Log"
Case "Test":   Lpath = path & "Test\Test_Log"
Case Else:     Lpath = path & "丸徳\丸徳_Log"
End Select

Dim Temp As String: Temp = Right(UF1.AorB.Caption, 1)
If Mitsu = "三菱" Then
  Call WriteLog("*UF全数|" & La.Name & "|" & La.Value & "|" & Lname & "|" & WorkerName, "", Lpath, fname)
  Call WriteLog("*UF全数|" & La.Name & "|" & La.Value & "|" & Lname & "|" & WorkerName, "")
Exit Sub
End If

If AorB = "" Then
  Call WriteLog("*UFKen|" & La.Name & "|" & La.Value & "|" & Lname & "|" & WorkerName, "", Lpath, fname)
  Call WriteLog("*UFKen|" & La.Name & "|" & La.Value & "|" & Lname & "|" & WorkerName, "")
Else
  Call WriteLog("*UFKen|" & "[" & Temp & "] " & La.Name & "|" & La.Value & "|" & Lname & "|" & WorkerName, "", Lpath, fname)
  Call WriteLog("*UFKen|" & "[" & Temp & "] " & La.Name & "|" & La.Value & "|" & Lname & "|" & WorkerName, "")
End If
End Sub

'コンボボックス用ログ出力
Public Sub LogOutCombo(Co As MSForms.ComboBox, Optional n As String = "")
Dim Lname      As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim fname      As String: fname = "Measured.log"
Dim path       As String: path = HistoryPath
Dim Lpath      As String
Dim WorkerName As String: WorkerName = GetWorkerName()
Dim Temp       As String: Temp = Right(UF1.AorB.Caption, 1) '追加: "A" or "B"

Select Case Lname
Case "L-1":    Lpath = path & "L-1\L-1_Log"
Case "LVC":    Lpath = path & "LVC\LVC_Log"
Case "HVC":    Lpath = path & "HVC\HVC_Log"
Case "ﾊﾞﾗﾝｻｰ": Lpath = path & "バランサー\バランサー_Log"
Case "ﾄｯﾄ":    Lpath = path & "トット\トット_Log"
Case "Test":   Lpath = path & "Test\Test_Log"
Case Else:     Lpath = path & "丸徳\丸徳_Log"
End Select

Call WriteLog("*UFKen|[" & Temp & "] " & Co.Name & "|" & Co.Value & "|" & Lname & "|" & WorkerName, "", Lpath, fname)
Call WriteLog("*UFKen|[" & Temp & "] " & Co.Name & "|" & Co.Value & "|" & Lname & "|" & WorkerName, "")
End Sub

'丸付き数字に変換する
Public Function CircledNumber(num As Long) As String
    CircledNumber = num
    If num < 1 Or num > 50 Then Exit Function
  
    Dim i As Long
    Select Case num
        Case Is <= 20: i = 9311
        Case Is <= 35: i = 12860
        Case Else: i = 12941
    End Select
    CircledNumber = ChrW(i + num)
End Function
' ' // クリップボードのビットマップデータから Picture オブジェクトを作成
'Public Function CreatePictureFromClipboard() As StdPicture’**************初期コード
'
'Dim hImg      As Long
'Dim hPalette As Long
'Dim uPictDesc As PictDesc
'Dim uGUID     As GUID
'Dim retryCount As Integer
'
'Set CreatePictureFromClipboard = Nothing
'
'' 終了条件:: クリップボードに該当データが無い
'If IsClipboardFormatAvailable(CF_BITMAP) = 0 Then Exit Function
'' 終了条件:: クリップボードからイメージハンドルが取得できない
'
'' リトライカウンター初期化
'retryCount = 0
'
'Do While retryCount < 5
'  ' クリップボードからイメージハンドルを取得
'  If OpenClipboard(0&) <> 0 Then
'      hImg = GetClipboardData(CF_BITMAP)
'      hPalette = GetClipboardData(CF_PALETTE)
'      Call CloseClipboard
'  End If
'
'  ' クリップボードからイメージハンドルが取得できた場合
'  If hImg <> 0 Then Exit Do
'
'  ' 少し待ってからリトライ
'  retryCount = retryCount + 1
'  DoEvents ' 他のプロセスに時間を与える
'  Application.Wait Now + TimeValue("00:00:01") ' 1秒待つ
'Loop
'
'
'If hImg = 0 Then Exit Function
'
'With uPictDesc
'   .cbSizeofStruct = Len(uPictDesc)
'   .picType = 1
'   .hImage = hImg
'   .Option1 = hPalette
'End With
'With uGUID
'   .Data1 = &H20400
'   .Data4(0) = &HC0
'   .Data4(7) = &H46
'End With
'Call OleCreatePictureIndirect(uPictDesc, uGUID, 0&, CreatePictureFromClipboard)
'
'End Function
' より高解像度対応のCreatePictureFromClipboard関数
'Public Function CreatePictureFromClipboard() As StdPicture’**************改良
'Dim hImg As Long
'Dim hPalette As Long
'Dim uPictDesc As PictDesc
'Dim uGUID As GUID
'Dim retryCount As Integer
'
'Set CreatePictureFromClipboard = Nothing
'
'' 終了条件:: クリップボードに該当データが無い
'If IsClipboardFormatAvailable(CF_BITMAP) = 0 And IsClipboardFormatAvailable(CF_ENHMETAFILE) = 0 Then Exit Function
'
'' リトライカウンター初期化
'retryCount = 0
'Do While retryCount < 10 ' リトライ回数を増やす
'  ' クリップボードからイメージハンドルを取得
'  If OpenClipboard(0&) <> 0 Then
'      ' まずはEnhanced Metafile（高解像度）を試す
'      hImg = GetClipboardData(CF_ENHMETAFILE)
'      If hImg <> 0 Then
'          uPictDesc.picType = 4 ' vbPicTypeEMetafile
'      Else
'          ' Bitmapにフォールバック
'          hImg = GetClipboardData(CF_BITMAP)
'          hPalette = GetClipboardData(CF_PALETTE)
'          uPictDesc.picType = 1 ' vbPicTypeBitmap
'      End If
'      Call CloseClipboard
'  End If
'
'  ' クリップボードからイメージハンドルが取得できた場合
'  If hImg <> 0 Then Exit Do
'
'  ' 少し待ってからリトライ
'  retryCount = retryCount + 1
'  DoEvents
'  Application.Wait Now + TimeValue("00:00:01")
'Loop
'
'If hImg = 0 Then Exit Function
'
'With uPictDesc
'  .cbSizeofStruct = Len(uPictDesc)
'  .hImage = hImg
'  If uPictDesc.picType = 1 Then .Option1 = hPalette
'End With
'
'With uGUID
'  .Data1 = &H20400
'  .Data4(0) = &HC0
'  .Data4(7) = &H46
'End With
'
'Call OleCreatePictureIndirect(uPictDesc, uGUID, 0&, CreatePictureFromClipboard)
'End Function


' 高解像度でセル範囲をコピーする新しいサブルーチン
Public Sub CopyRangeAsHighResolutionImage(ws As Worksheet, rng As Range)
Dim originalZoom As Integer
Dim originalView As Integer
Dim originalActiveSheet As Worksheet

' 現在のアクティブシートを保存
If UFMitsubishi.Visible Then
  Set originalActiveSheet = ThisWorkbook.Sheets("全ヵ所検査")
Else
  Set originalActiveSheet = ws
End If
' 対象シートをアクティブにする
ws.Activate

' 現在のズーム値を保存
originalZoom = ActiveWindow.Zoom
originalView = ActiveWindow.View

' 高解像度化のための設定
ActiveWindow.View = xlNormalView
ActiveWindow.Zoom = 200

' Selectを使わずに直接CopyPicture
Application.CutCopyMode = False
rng.CopyPicture Appearance:=xlScreen, Format:=xlBitmap

' 元の設定に戻す
ActiveWindow.Zoom = originalZoom
ActiveWindow.View = originalView

' 元のアクティブシートに戻す
If UFMitsubishi.Visible Then originalActiveSheet.Activate
End Sub
Public Function CreatePictureFromClipboard() As StdPicture '**************改良2
#If VBA7 Then
  Dim hImg As LongPtr
  Dim hPalette As LongPtr
#Else
  Dim hImg As Long
  Dim hPalette As Long
#End If

Dim uPictDesc As PictDesc
Dim uGUID As GUID
Dim retryCount As Integer

Set CreatePictureFromClipboard = Nothing

' 終了条件:: クリップボードに該当データが無い
If IsClipboardFormatAvailable(CF_BITMAP) = 0 And IsClipboardFormatAvailable(CF_ENHMETAFILE) = 0 Then Exit Function

' リトライカウンター初期化
retryCount = 0

Do While retryCount < 15 ' リトライ回数を増やす
  ' クリップボードからイメージハンドルを取得
  If OpenClipboard(0&) <> 0 Then
      ' まずはEnhanced Metafile（高解像度）を試す
      hImg = GetClipboardData(CF_ENHMETAFILE)
      If hImg <> 0 Then
          uPictDesc.picType = 4 ' vbPicTypeEMetafile
      Else
          ' Bitmapにフォールバック
          hImg = GetClipboardData(CF_BITMAP)
          hPalette = GetClipboardData(CF_PALETTE)
          uPictDesc.picType = 1 ' vbPicTypeBitmap
      End If
      Call CloseClipboard
  End If
  
  ' クリップボードからイメージハンドルが取得できた場合
  If hImg <> 0 Then Exit Do
  
  ' 少し待ってからリトライ
  retryCount = retryCount + 1
  DoEvents
  Sleep 200 ' 待機時間を短縮
Loop

If hImg = 0 Then Exit Function

With uPictDesc
  .cbSizeofStruct = Len(uPictDesc)
  .hImage = hImg
  If uPictDesc.picType = 1 Then .Option1 = hPalette
End With

With uGUID
  .Data1 = &H20400
  .Data4(0) = &HC0
  .Data4(7) = &H46
End With

Call OleCreatePictureIndirect(uPictDesc, uGUID, 0&, CreatePictureFromClipboard)
End Function

Public Function ReadFile(FilePath As String, Optional encoding As String = "UTF-8") As Variant 'UTF-8 エンコーディングで読み込む
Dim objFSO As Object
Dim objFile As Object
Dim stream As Object
Dim fileContent As String
Set objFSO = CreateObject("Scripting.FileSystemObject") ' ファイル オブジェクトを作成
Set objFile = objFSO.OpenTextFile(FilePath, 1, False, TristateUseDefault)
fileContent = objFile.ReadAll ' ファイル内容を読み込む

objFile.Close ' ファイルを閉じる

' オブジェクトを解放
Set objFile = Nothing
Set objFSO = Nothing

Set stream = CreateObject("ADODB.Stream") ' ストリーム オブジェクトを作成

' ストリームをUTF-8で扱うように設定
stream.Charset = encoding '"UTF-8"
stream.Open
stream.WriteText fileContent
stream.Position = 0

' UTF-8 に変換した内容を返す
ReadFile = stream.ReadText

' ストリームを閉じる
stream.Close
Set stream = Nothing
'Stop
End Function
Public Function FindPriority(str As String) As String '最終圧延から検査番号の頭を作る
Dim Keywords As Variant, Key As Variant, Words As Variant, Word As Variant
Dim priority As String, i As Long
Keywords = Array("CM1", "CM2", "CM3", "2HI") '検索対象のキーワード
priority = "" 'デフォルトの優先
Words = Split(str, " ") ' 文字列をスペースで分割
For i = UBound(Words) To LBound(Words) Step -1 '分割された単語を反転してループ
  Word = Words(i)
  For Each Key In Keywords
    If Key = Word Then
      priority = Key
      Exit For
    End If
  Next Key
  If priority <> "" Then Exit For ' 優先が見つかった場合はループを抜ける
Next i

' 優先度が見つからない場合は再度"HOT"を検索
If priority = "" Then
  For Each Word In Words
    If Word = "HOT" Then
      priority = "HOT"
      Exit For
    End If
  Next Word
End If
'検査番号の頭に振替
Select Case priority
Case "CM1": priority = "L"
Case "CM2": priority = "N"
Case "CM3": priority = "K"
Case "2HI": priority = "T"
Case "HOT": priority = "H"
End Select

FindPriority = priority
End Function
'マルチページ画像貼り付け
Public Sub LoadImagesIntoMultiPageFromTemp(MultiP As MSForms.MultiPage, Optional FolderN As String = "品質写真") 'La As MSForms.ListBox
Dim FolderPath As String: FolderPath = QualityPath & FolderN & "\" '"C:\Users\xfa-ngyinazawa\Desktop\品質写真\"
Dim PageName As String, ImagePath As String
Dim ctl As Control
Dim frm As frame
Dim img As Image
Dim IntIndex As Long: IntIndex = MultiP.Value
Dim pic As Object, KariPic As Image

If ReadCheck(FolderPath, 2) = False Then Exit Sub

PageName = MultiP.Pages(IntIndex).Caption ' ページの名前を取得

For Each ctl In MultiP.Pages(IntIndex).Controls ' ページ内のすべてのコントロール
  If TypeOf ctl Is Image Then ' Imageコントロールであるか確認
    ImagePath = Dir(FolderPath & PageName & "\*.jpg", vbDirectory) 'サブフォルダ内の.JPGファイルを処理
    If ImagePath <> "" Then
      Set img = ctl
      Set pic = LoadPicture(FolderPath & PageName & "\" & ImagePath) ' 画像を読み込んで表示
      img.Picture = pic
      Set KariPic = img 'ここでImageにしておかないと正しいサイズが読めない
    End If
  End If
Next ctl
'フレームを探す
For Each ctl In MultiP.Pages(IntIndex).Controls ' ページ内のすべてのコントロール
  If TypeOf ctl Is frame Then
'  MsgBox Ctl.Name
    If ImagePath <> "" Then
      Set frm = ctl
      frm.ScrollHeight = KariPic.Height 'Imageのサイズ
      frm.ScrollWidth = KariPic.Width
      frm.Picture = pic
    End If
  End If
Next ctl
End Sub
'指定パスフォルダ数分ページ追加とページ内にイメージとフレームの追加
Public Function AddPagesForSubFolders(MP As MSForms.MultiPage, Optional FolderN As String = "品質写真") As String
'Dim MP As MultiPage:        Set MP = Test.MP1  'UF品質基準.MPH 'MP = Test.MP1
Dim FolderPath As String:   FolderPath = QualityPath & FolderN & "\"  '"C:\Users\xfa-ngyinazawa\Desktop\品質写真\" ' 指定されたフォルダパス
Dim fso As Object:          Set fso = CreateObject("Scripting.FileSystemObject")
Dim newPage As Object
Dim SubFolder As Object

If ReadCheck(FolderPath, 2) = False Then Exit Function
'フォルダの数分追加
Dim i As Long
Dim SubCount As Long:   SubCount = fso.GetFolder(FolderPath).SubFolders.Count
Dim NumPages  As Long:  NumPages = MP.Pages.Count
Dim AddCount As Long:   AddCount = SubCount
If AddCount < 1 Then Exit Function

Dim img As Image
Dim frm As frame
Dim SubFolderName As String

For Each SubFolder In fso.GetFolder(FolderPath).SubFolders
  SubFolderName = SubFolder.Name
  AddPagesForSubFolders = AddPagesForSubFolders & SubFolderName & vbCrLf
  Set newPage = MP.Pages.Add 'ページオブジェクトを追加
  newPage.Caption = SubFolder.Name
  Set img = newPage.Controls.Add("Forms.Image.1", "Image1") 'イメージオブジェクトを追加
  img.AutoSize = True ' イメージのサイズをオートサイズに設定
  img.Visible = False '
  Set frm = newPage.Controls.Add("Forms.Frame.1", "Frame1") 'フレームオブジェクトを追加
  frm.ScrollBars = fmScrollBarsBoth 'スクロールバーセット
  frm.tOP = 0 '上端をページの上端に配置
  frm.Left = 0 '左端をページの左端に配置
  frm.Width = MP.Width - 10 '幅をページの幅に合わせる
  frm.Height = MP.Height - 20 'の高さをページの高さに合わせる
Next SubFolder
Set fso = Nothing
Set newPage = Nothing
End Function
'配列内文字列分Tab追加とページ内にTextBoxとラベルの追加
Public Function AddPagesArrN(Mpage As MSForms.TabStrip, n As Variant) As String
Dim MlPage As Object
If IsEmptyArray(n) Then Exit Function
Dim narryC As Long: narryC = UBound(n, 1)

Dim i As Long
Dim NumPages  As Long:  NumPages = Mpage.Tabs.Count
Dim AddCount As Long:   AddCount = narryC
If AddCount < 1 Then Exit Function

Dim lbl As MSForms.Label
Dim txtBox As Object
Dim txtAbn As Object
Dim SubFolderName As String

'欠陥名、欠陥許容値、グレードの配列
Dim a As Long: a = LBound(n, 2)
Dim k As Long: k = LBound(n, 2) + 1
Dim j As Long: j = UBound(n, 2)
Dim RepeT As String
For i = LBound(n, 1) To UBound(n, 1)
  AddPagesArrN = n(i, LBound(n, 2)) & vbCrLf
  Set MlPage = Mpage.Tabs.Add 'ページオブジェクトを追加
  MlPage.Caption = n(i, a)
  RepeT = RepeT & n(i, a) & vbCrLf
  
'  Set lbl = MlPage.Controls.Add("Forms.Label.1", "lbl" & i)
'  lbl.Caption = n(i, a)
'  lbl.AutoSize = True
'  lbl.Top = 20 '上端をページの上端+20に配置
'  lbl.Left = 0 '左端をページの左端に配置
  
'  Set txtBox = MlPage.Controls.Add("Forms.TextBox.1", "txtBox" & i)
'  txtBox.Top = 20
'  txtBox.Left = lbl.width + 15
'  txtBox.text = n(i, k)
'  txtBox.AutoSize = True
'  txtBox.Locked = True
'  txtBox.TabStop = False
'
'  Set txtAbn = MlPage.Controls.Add("Forms.TextBox.1", "txtBox" & i)
'  txtAbn.Top = 20
'  txtAbn.Left = txtBox.width + lbl.width + 25
'  txtAbn.width = Mpage.width - txtBox.width - lbl.width - 55
'  txtAbn.Height = txtBox.Height * 5
'  txtAbn.MultiLine = True
'  txtAbn.text = n(i, j)
'  txtAbn.Locked = True
'  txtAbn.TabStop = False
Next i
AddPagesArrN = RepeT
Set MlPage = Nothing
End Function
'配列の中身がすべて同じ文字列かどうかをチェック
Function AllElementsSame(arr As Variant) As Boolean 'arrは1次元
Dim i As Long
Dim checkString As String
checkString = arr(LBound(arr)) ' 配列の最初の要素を取得してチェック文字列として設定
For i = LBound(arr) To UBound(arr) ' 配列の各要素がチェック文字列と一致するか確認
  If arr(i) <> checkString Then
    AllElementsSame = False
    Exit Function
  End If
Next i
AllElementsSame = True ' すべての要素が一致した場合はTrueを返す
End Function
' 配列内の各要素を _ で分割し、Width 配列と Length 配列にそれぞれ格納
Public Sub SplitWidthLength(WidthLength As Variant, ByRef WidthArray As Variant, ByRef LengthArray As Variant)
Dim i As Long, Size As Long
Dim parts() As String
Size = UBound(WidthLength, 1) - LBound(WidthLength, 1) + 1

'分割された要素を格納するための2つの2次元配列をリサイズ
ReDim WidthArray(1 To Size)
ReDim LengthArray(1 To Size)

For i = LBound(WidthLength, 1) To UBound(WidthLength, 1)
  If WidthLength(i) <> "" Then
    parts = Split(WidthLength(i), "_")
    WidthArray(i) = parts(0) ' WidthとLengthの2次元配列に分割された要素を格納
    LengthArray(i) = parts(1)
  End If
Next i
End Sub
Public Function SecondSmallest(Ayy As Variant) As Variant  '配列内比較
Dim tempArray() As Variant
Dim i As Long, j As Long
Dim Temp As Variant
' 配列を一時的な配列にコピー
ReDim tempArray(LBound(Ayy) To UBound(Ayy))
For i = LBound(Ayy) To UBound(Ayy)
    tempArray(i) = Ayy(i)
Next i
' 配列を昇順でソート
For i = LBound(tempArray) To UBound(tempArray) - 1
    For j = i + 1 To UBound(tempArray)
        If tempArray(i) > tempArray(j) Then
            Temp = tempArray(i)
            tempArray(i) = tempArray(j)
            tempArray(j) = Temp
        End If
    Next j
Next i
' 重複を除外して2番目に小さい数値を返す
For i = LBound(tempArray) To UBound(tempArray) - 1
    If tempArray(i) <> tempArray(i + 1) Then
        SecondSmallest = tempArray(i + 1)
        Exit Function
    End If
Next i
' 2番目に小さい数値が見つからない場合はエラーを返す
SecondSmallest = CVErr(xlErrNA)
End Function
Sub CreateHighQualityReport()
    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = Sheets.Add
    ws.Name = "改善報告書_正式版"

    ' --- 1. 全体設定 ---
    With ws
        .Cells.Font.Name = "Meiryo UI"
        .Cells.Font.Size = 10
        .Columns("A").ColumnWidth = 5   ' No用
        .Columns("B").ColumnWidth = 15  ' 項目用
        .Columns("C").ColumnWidth = 20  ' 内容1
        .Columns("D").ColumnWidth = 30  ' 内容2
        .Columns("E").ColumnWidth = 15  ' 備考/サイン
        
        ' 作成者情報 (右寄せ)
        .Range("D1:E1").Merge
        .Range("D1").Value = "作成日：2026年5月15日　作成者：開発担当"
        .Range("D1").HorizontalAlignment = xlRight
        .Range("D1").Font.Size = 9

        ' タイトル (中央揃え・太字)
        .Range("A2:E2").Merge
        With .Range("A2")
            .Value = "不具合改善報告書"
            .Font.Size = 16
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With

        ' --- 2. ヘッダー情報の表 ---
        Dim headerItems As Variant
        headerItems = Array( _
            Array("件　名", "ログ呼び出し（AccesslogLoad）後の印刷実行時に検査員エラーで処理停止する事象の改善"), _
            Array("対象システム", "検査判定システム（UF1）／印刷転送モジュール"), _
            Array("発生状況", "非定常作業（フォーム値焼失等）でログから検査データを呼び出した後、印刷を実行すると処理が停止する") _
        )
        Dim i As Long
        For i = 0 To 2
            .Range("A" & (4 + i)).Value = headerItems(i)(0)
            .Range("B" & (4 + i) & ":E" & (4 + i)).Merge
            .Range("B" & (4 + i)).Value = headerItems(i)(1)
            ' 書式設定
            With .Range("A" & (4 + i))
                .Interior.Color = RGB(232, 232, 232)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
        Next i
        .Range("A4:E6").Borders.LineStyle = xlContinuous

        ' --- 3. セクション作成関数（共通処理） ---
        Call CreateSection(ws, 8, "1. 事象の詳細")
        .Range("A9:E10").Merge
        .Range("A9").Value = "通常運用では検査員名は WorkerTime シートに記録されるが、旧バージョンのログには検査員名フィールドが存在しなかった。そのためログ呼び出し時にシートが更新されず、印刷転送時の空白チェックで処理が強制停止（Tag='stop'）していた。"
        .Range("A9").WrapText = True
        .Range("A9").VerticalAlignment = xlTop

        Call CreateSection(ws, 12, "2. 原因まとめ")
        ' 原因テーブル
        Dim causeData As Variant
        causeData = Array( _
            Array("直接原因", "旧ログに検査員名フィールドが存在せず、WorkerTimeシートが空のままだった"), _
            Array("根本原因", "AccesslogLoad がWorkerTimeシートへの反映機能を持っていなかった"), _
            Array("副要因", "ログ区切り文字（:）により、フィールド取得のオフセット計算が不安定だった") _
        )
        For i = 0 To 2
            .Cells(13 + i, 1).Value = causeData(i)(0)
            .Range("B" & (13 + i) & ":E" & (13 + i)).Merge
            .Range("B" & (13 + i)).Value = causeData(i)(1)
            .Cells(13 + i, 1).Interior.Color = RGB(240, 240, 240)
            .Cells(13 + i, 1).Font.Bold = True
        Next i
        .Range("A13:E15").Borders.LineStyle = xlContinuous

        Call CreateSection(ws, 17, "3. 実施した改善内容")
        ' 改善内容テーブル
        .Range("A18").Value = "No": .Range("B18").Value = "改善項目": .Range("C18:E18").Merge: .Range("C18").Value = "内容"
        .Range("A18:E18").Interior.Color = RGB(85, 85, 85): .Range("A18:E18").Font.Color = vbWhite
        
        Dim impData As Variant
        impData = Array( _
            Array(1, "ログ区切り文字変更", "「:」から「|」に変更。新旧ログの自動判別機能を実装。"), _
            Array(2, "検査員名ログ出力", "WriteLog呼び出し時に GetWorkerName() で取得した名を末尾に追加。"), _
            Array(3, "WorkerTime反映", "新関数 ApplyWorkerFromLog 等によりログからシートを自動復元。"), _
            Array(4, "高速化・文字化け", "ADODB.Stream 1行読みに変更。Shift-JIS統一で文字化け解消。"), _
            Array(5, "外観・歪データ復元", "外観テキストおよび歪測定値をログから復元可能にした。") _
        )
        For i = 0 To 4
            .Cells(19 + i, 1).Value = impData(i)(0)
            .Cells(19 + i, 2).Value = impData(i)(1)
            .Range("C" & (19 + i) & ":E" & (19 + i)).Merge
            .Range("C" & (19 + i)).Value = impData(i)(2)
            .Cells(19 + i, 1).HorizontalAlignment = xlCenter
        Next i
        .Range("A18:E23").Borders.LineStyle = xlContinuous

        Call CreateSection(ws, 25, "4. 改善後の動作")
        .Range("A26:E27").Merge
        .Range("A26").Value = "ログ呼び出し後にWorkerTimeシートへ検査員名が正しく書き込まれるため、印刷処理が正常に完了する。また外観・歪の値も復元されるため、非定常フローでの作業再開が容易になった。"
        .Range("A26").WrapText = True

        ' --- 4. サイン欄 (右下) ---
        Dim signRow As Long: signRow = 29
        .Cells(signRow, 4).Value = "確認": .Cells(signRow, 5).Value = "承認"
        .Range(.Cells(signRow, 4), .Cells(signRow, 5)).Interior.Color = RGB(232, 232, 232)
        .Range(.Cells(signRow, 4), .Cells(signRow, 5)).HorizontalAlignment = xlCenter
        .Range(.Cells(signRow, 4), .Cells(signRow + 1, 5)).Borders.LineStyle = xlContinuous
        .Range(.Cells(signRow + 1, 4), .Cells(signRow + 1, 5)).RowHeight = 35
    End With
    
    MsgBox "画像に基づいた正式な報告書を作成しました。"
End Sub

' 見出し作成用補助関数
Sub CreateSection(ws As Worksheet, rowNum As Long, title As String)
    With ws.Range("A" & rowNum & ":E" & rowNum)
        .Merge
        .Value = title
        .Interior.Color = RGB(51, 51, 51)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
End Sub

