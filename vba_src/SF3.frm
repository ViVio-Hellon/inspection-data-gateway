Attribute VB_Name = "SF3"
Attribute VB_Base = "0{12FECAB6-4CAF-4CA1-AB9B-850913F25F92}{6C41790A-1744-486F-8BB5-7525D494F31F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
#If VBA7 And Win64 Then
  '64ビット版
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
  '32ビット版
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
#If Win64 Then
  Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
  Private Declare PtrSafe Function APIBeep Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
#Else
  Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
  Private Declare Function APIBeep Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
#End If


Private Sub CB1_Click()
If (Me.LineChenge.Text <> "L-1") And (Me.LineChenge.Text <> "LVC") And (Me.LineChenge.Text <> "HVC") _
  And (Me.LineChenge.Text <> "バランサー") And (Me.LineChenge.Text <> "トット") And (Me.LineChenge.Text <> "AIM") And (Me.LineChenge.Text <> "Test") _
  And (Me.LineChenge.Text <> "丸徳") Then 'いずれでもない
    MsgBox "正しい設備名バーコードを読ませてください。", vbCritical
    Exit Sub
ElseIf (Me.LineChenge.Text = "L-1") Or (Me.LineChenge.Text = "LVC") Or (Me.LineChenge.Text = "HVC") _
  Or (Me.LineChenge.Text = "バランサー") Or (Me.LineChenge.Text = "トット") Or (Me.LineChenge.Text = "AIM") Or (Me.LineChenge.Text = "Test") _
  Or (Me.LineChenge.Text = "丸徳") Then 'どれかにヒット
    SaveSetting "UF検査判定システム", "起動設定", "設備名", Me.LineChenge
    SaveSetting "UF検査判定システム", "起動設定", "モード", Me.F1.Tag
    
    If WsCheck("Main") Then ThisWorkbook.Sheets("Main").Cells(2, 1) = Me.LineChenge
    
    Unload Me
    MsgBox "設定しました。" & vbCrLf & "フォームが起動するまでお待ちください。", vbInformation
Else
    MsgBox "設備名バーコードを読ませてください。", vbCritical
    Exit Sub
End If
End Sub

Private Sub CommandButton1_Click()
MsgBox GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")
End Sub

Private Sub LineChenge_Change()
Dim Temp As String
If Len(Me.LineChenge.Text) <> 10 Then  'ｶｰﾄﾞ10桁
  Exit Sub
Else
  Dim WorLine As String: WorLine = LineConv(Me.LineChenge.Value) '2021.04.16_13桁の時だけ変換
  Me.LineChenge = WorLine
End If
End Sub

Private Sub OP1_Click()
Me.F1.Tag = "オンライン"
End Sub
Private Sub OP2_Click()
Me.F1.Tag = "オフライン"
End Sub
Private Sub UserForm_Activate()
Call FormSetting ' フォームをリサイズ可能にするための設定
End Sub

Private Sub UserForm_Initialize()
Call WindowMenuButtonDel(Me)
Me.LB1.List = Array("L-1", "LVC", "HVC", "バランサー", "トット", "Test", "丸徳")
Me.F1.Tag = "オンライン"
End Sub
'Public Sub 初期設定値確認(Optional Flag As Boolean = False)
'Dim tmpFaci As String: tmpFaci = GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")
'Dim tmpMode As String: tmpMode = GetSetting("UF検査判定システム", "起動設定", "モード", "存在なし")
'Select Case True
'Case tmpFaci = "存在なし" Or Flag '初期設定モード
'  DoEvents
'  Sleep 500
'  SF3.Show
'Case Else
'End Select
'End Sub
Public Function LineConv(no As String) As String
'------------------------------DBから読み取る
If no = "" Then Exit Function

Dim tmpFaci As String: tmpFaci = GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")
Dim tmpMode As String: tmpMode = GetSetting("UF検査判定システム", "起動設定", "モード", "存在なし")

Dim TableName As String:   TableName = "設備名バーコード"
Dim FPath As String:       FPath = QualityPath '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\検査\"
'初期起動時パス設定がないのでライン名が読めない
If FPath = "" Then
  If tmpMode = "存在なし" Then
    FPath = "\\nlmsrvngy03\工場内共有\検査データ\品証-梱包やり取りフォルダ\Accdb\" ' ThisWorkbook.path & "\"
  End If
End If

Dim target As String:      target = "伝送用ファイル" & ".accdb" 'ﾌｧｲﾙ名
Dim FilePath As String:    FilePath = FPath & target 'ﾌｧｲﾙﾊﾟｽ
If ReadCheck(FilePath, 1) = False Then '1：ファイル,2：フォルダ
  FPath = PackingPath
  FilePath = FPath & target
  If ReadCheck(FilePath, 1) = False Then
    MsgBox target & "がありません" & "_LineConv", vbCritical
    Exit Function
  End If
End If
'検索項目
Dim SQLfilter As String: SQLfilter = "[No]='" & no & "'" 'No検索
Dim Temp As Variant: Temp = GetRecordsArr(FilePath, TableName, SQLfilter) 'All,fields,record
'ﾌｨｰﾙﾄﾞ名ﾁｪｯｸ　*更新に対応
Dim fieldsfilter As String: fieldsfilter = ""
Dim Fields As Variant:  Fields = GetFieldsArr(FilePath, TableName, fieldsfilter)
  If IsEmpty(Fields) Then 'emptyの場合停止
  MsgBox "fieldsなし" & "_NameConv"
  Exit Function
End If
Dim i As Long, C_名前 As Long
For i = LBound(Fields, 2) To UBound(Fields, 2)
  If Fields(1, i) = "名前" Then
    C_名前 = i
    Exit For
  End If
Next i
If IsEmpty(Temp) Then 'emptyの場合停止
  LineConv = ""
  Exit Function
Else
  Dim Spath As String: Spath = "C:\VBAWAV\"
  Dim Sfile As String: Sfile = "chime1.wav"
  Dim SoundFile As String, rc As Long
  SoundFile = Spath & Sfile
  If Not ReadCheck(SoundFile, 1) = False Then rc = mciSendString("Play " & SoundFile, "", 0, 0)
  
  LineConv = Temp(1, C_名前)
End If
End Function
