Attribute VB_Name = "TEST"
Attribute VB_Base = "0{16CCE326-5DA4-4F9F-A09D-24D914520AC9}{701F7DE7-502B-49D2-B9C2-D1085DA7AE35}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit
'Private Start1 As String
'Private Start2 As String
'Private Start3 As String
'Private Start昼 As String
'Private End1 As String
'Private End2 As String
'Private End3 As String
'Private End昼 As String
'
'Private MainPath As String
'Private NextPath As String
'Private QualityPath As String
'Private PackingPath As String
'Private HistoryPath As String
'Private FactoryyPath As String

Dim objShell As Object
Dim objExec As Object

Private mx As Single
Private my As Single
Private IsClick As Boolean
#If Win64 Then
  Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
  Private Declare PtrSafe Function APIBeep Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
#Else
  Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
  Private Declare Function APIBeep Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
#End If
#If VBA7 And Win64 Then
  '64ビット版
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
  '32ビット版
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Sub CommandButton31_Click()
Dim Temp As Variant: Temp = Array(5, 1, 2, 1, 1, 3, 5, 6, 10)
Dim n As Variant:    n = SecondSmallest(Temp)
Stop
End Sub

Private Sub CommandButton33_Click()
Call 初期設定値確認
End Sub

Private Sub CommandButton34_Click()
Call Setting_Path
Dim LotNo As String: LotNo = Me.T2

Dim TableName As String: TableName = "仕掛"
Dim FPath As String:     FPath = NextPath
'ﾌｧｲﾙ名
Dim Target_Hiki As String:  Target_Hiki = "SIKAHIKINOW.accdb"
'ﾌｧｲﾙﾊﾟｽ
Dim FilePath_Hiki As String:  FilePath_Hiki = FPath & Target_Hiki '引当
'検索項目
Dim SQLfilter As String:    SQLfilter = "[ﾛｯﾄ番号]='" & LotNo & "'" 'ﾛｯﾄNo検索
Dim TempHiki As Variant:    TempHiki = GetRecordsArr(FilePath_Hiki, TableName, SQLfilter) 'All,fields,record
Dim Hikifields As Variant:  Hikifields = GetFieldsArr(FilePath_Hiki, TableName, "")
If IsEmpty(TempHiki) Then Exit Sub 'emptyの場合停止

Dim i As Long, H_検査番号 As Long, H_ロット番号 As Long 'Hiki
For i = LBound(Hikifields, 2) To UBound(Hikifields, 2)
  If Hikifields(1, i) = "検査番号" Then
    H_検査番号 = i
  End If
  If Hikifields(1, i) = "ﾛｯﾄ番号" Then
    H_ロット番号 = i
  End If
Next i

Me.T2 = TempHiki(1, H_ロット番号)
Me.T3 = TempHiki(1, H_検査番号)

Dim KenNo As String: KenNo = Me.T3

Dim TempD As Variant: TempD = HikiArrCreate(LotNo) '納入
Dim a As Long, 納入 As String, 取引 As String
Dim QpArryD As Variant, QpArryG As Variant
納入 = TempD(1, 1)
取引 = TempD(1, 2)
QpArryD = DistArray(LotNo, KenNo, "Test", "Test")


Me.MG = Format(QpArryD(LBound(QpArryD, 1), 1), "0.0")
Me.HaHiz = Format(QpArryD(LBound(QpArryD, 1), 2), "0.0")
Me.TaHiz = Format(QpArryD(LBound(QpArryD, 1), 3), "0.0")

Me.Naka = Format(QpArryD(LBound(QpArryD, 1), 4), "0.0")
Me.Mimi = Format(QpArryD(LBound(QpArryD, 1), 5), "0.0")

Me.HizCome = QpArryD(LBound(QpArryD, 1), 6)
Me.DistortionFrame.Caption = QpArryD(LBound(QpArryD, 1), 7)
Me.NoTex = QpArryD(LBound(QpArryD, 1), 8)

MsgBox "ロード済"

'Stop
End Sub
Sub 値の入力(ByVal targetRange As Range)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim n As Integer
    Dim Cell As Range

    n = 1 ' 最初の値
    k = 4 ' 同じ値の数

    For Each Cell In targetRange
        For j = 1 To k
            Cell.Value = n
            Set Cell = Cell.Offset(1, 0)
        Next j
        Set Cell = Cell.Offset(-1, 0) ' 直前のセルに戻る
        If n = 1 Or n = 2 Then
            Set Cell = Cell.Offset(1, 0) ' 休憩を挿入するために次のセルに移動
            Cell.Value = "休"
            Set Cell = Cell.Offset(1, 0)
        ElseIf n = 3 Then
            Set Cell = Cell.Offset(1, 0)
            Cell.Value = "休"
            Set Cell = Cell.Offset(1, 0)
            Cell.Value = "休"
            Set Cell = Cell.Offset(1, 0)
        End If
        n = n + 1
        If n > 3 Then Exit For ' 4以上の値は不要
    Next Cell
End Sub
Private Sub CommandButton35_Click()

MsgBox Application.Height
End Sub

Private Sub CommandButton36_Click()

Dim tMP As String: tMP = Me.A1

Select Case True
Case Me.TextTest Like "*" & tMP & "*"
  MsgBox "aru"
'Case Else
'  MsgBox "nai"
End Select

If Left(tMP, 1) = "≠" Then
  tMP = Replace(tMP, "≠", "")
  Select Case True
  Case Not Me.TextTest Like "" & tMP & ""
    MsgBox "not aru"
  End Select
End If
End Sub

Private Sub CommandButton37_Click()
''2021.03.05 L-1:A_input導入 2021.03.12 LVC:A_input導入 2021.03.29 ﾊﾞﾗﾝｻｰ:A_input導入 2021.05.11 HVC:A_input導入  2021.05.11 ﾄｯﾄ:A_input導入　※ﾄｯﾄﾊﾞﾗﾝｻｰにｱｲｶｽが１台の為統合(1excel1ｱｲｶｽ)
'Dim Tmp As String:         Tmp = "2021.03.05 L-1:A_input導入。2021.03.12 LVC:A_input導入 2021.03.29 ﾊﾞﾗﾝｻｰ:A_input導入。2021.05.11 HVC:A_input導入。2021.05.11 ﾄｯﾄ:A_input導入。※ﾄｯﾄﾊﾞﾗﾝｻｰにｱｲｶｽが１台の為統合(1excel1ｱｲｶｽ)。"
'Dim ModifiedTmp As String: ModifiedTmp = Replace(Tmp, "。", vbCrLf)
'Me.HizCome.Caption = ModifiedTmp

Dim Mdevice As String:   Mdevice = DbReturnMicro("LVC", 5, "H162", "機器名")  '機器名出力
'MsgBox Mdevice


If Not "M1" = Mdevice Then
  MsgBox (Mdevice & "を使用の事")
Else
  MsgBox Mdevice
End If

End Sub

Private Sub CommandButton38_Click()
Call Setting_Path
Dim Temp As String: Temp = CheckEquipmentDb("LVC", "5NB", 0.49, 730, 1500, 225)
MsgBox Temp
'If Temp <> "可" Then Call MsgError("設備仕様を満たしていません,終了します。", "LVC", 2, "END")
End Sub

Private Sub CommandButton39_Click()
Dim ResultT As Boolean: ResultT = IsOver12Hours
MsgBox ResultT
End Sub

Private Sub TabST_Change()
Dim SelectedTabIndex As Long: SelectedTabIndex = Me.TabST.SelectedItem.Index
' タブのキャプションを取得して表示
MsgBox "選択されたタブのキャプションは: " & Me.TabST.Tabs(SelectedTabIndex).Caption
End Sub

Private Sub btn0_Click()
'Me.ListBox1 = ""
'Me.ListBox2 = ""
'Me.ListBox3 = ""
'Me.ListBox1 = Me.T1


Me.ListBox1.Tag = Me.ListBox1.Caption
Me.MoveFlag.Text = Me.ListBox1.Tag
End Sub
Private Sub btn0_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
mx = x
my = y
IsClick = True
Me.L1.Caption = mx
Me.L2.Caption = my
End Sub
Private Sub btn0_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
If IsClick = True Then
Me.btn0.Left = btn0.Left - (mx - x)
btn0.tOP = btn0.tOP - (my - y)
Me.L1.Caption = x
Me.L2.Caption = y
End If
End Sub
Private Sub btn0_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
IsClick = False
mx = 0
my = 0
End Sub

Private Sub CommandButton11_Click()
Call AllClose("日報＆資材計算.xlsm")
Application.ScreenUpdating = False
Call HiddenSheets '隠す
Application.ScreenUpdating = True
End Sub


Private Sub CommandButton13_Click()

'MsgBox Len("\\nlmsrvngy03\Read\【New】仕掛\台帳\")

Dim fname As String: fname = "PLEASE WAIT.txt"
Dim FPath As String: FPath = "\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\仕掛\"
'If FileDesignation(Fname, fPath, "Test") = False Then Exit Sub

If TimeDesignation Or FileDesignation = False Then MsgBox "台帳侵入禁止"

End Sub
Sub Main()
    
    Dim objIE As InternetExplorerMedium
    
    'IEプロセスを削除'
    IeProcessKill
    
    'Googleトップページを開く'
    CreateIE objIE, "https://www.google.co.jp/"
    
    '1秒待機'
    Sleep 1000
    
    'IEを閉じる'
    objIE.Quit
    Set objIE = Nothing
    
    '処理完了メッセージ表示'
    MsgBox "処理完了"
    
End Sub
 
'IEのプロセスを削除する関数'
Public Sub IeProcessKill()
    
    Dim objShell As Object
    Dim objExec As Object
    
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("taskkill.exe /F /IM iexplore.exe")
    Sleep 1000
    
End Sub
 
'Internet Explorer作成関数'
Sub CreateIE(ByRef objIE As InternetExplorerMedium, ByVal url As String)
    
    Dim timeout As Date
    timeout = Now + TimeSerial(0, 0, 10)
    
    If objIE Is Nothing Then
        Set objIE = New InternetExplorerMedium
    End If
    
    objIE.Visible = True
    objIE.Navigate url
    
    Do While objIE.Busy = True Or objIE.readyState <> 4
        Sleep 100
        DoEvents
        If Now > timeout Then
            objIE.Refresh
            timeout = Now + TimeSerial(0, 0, 10)
        End If
    Loop
End Sub

Private Sub CommandButton14_Click()
'PackingNoCheck'OrdNoOpen

' Call UFPack_Open(Me.TkNo, Me.TkNo)

'Call EdgeSiteSetTextData("1P1185")

Dim result1 As String: result1 = MaterialsUpperPart(Me.TextBox1)    ' ザラ板/アングル
Dim result2 As String: result2 = MaterialsUpperPart(Me.TextBox1, True)   ' 当板有無

    If result1 <> "" And result1 <> "一致なし" Then
      MsgBox result1
    End If
    If result1 = "アングル" Then
        If result2 <> "" And result2 <> "一致なし" Then MsgBox result2
    End If
    
'Call IeSiteSetTextData("1P1177")
End Sub

Private Sub CommandButton15_Click()
On Error Resume Next
'Dim winshell As Object: Set winshell = CreateObject("Shell.Application")
'Dim win As Object
'For Each win In winshell.Windows
'  If win.Name = "Internet Explorer" Then
'    Set objIE = win
'    Exit For
'  End If
'Next
'objIE.Quit
Call ie_kill
End Sub

Private Sub CommandButton16_Click()
Dim FPath As String:     FPath = ThisWorkbook.path
Dim target As String:    target = "SIKAHIKINOW.accdb" '---------------------------Lot一覧ｼｰﾄからLoTNo_Accessから引当情報
Dim FilePath As String:  FilePath = FPath & "\" & target
If FPath = ThisWorkbook.path Then
  If ReadCheck(FilePath, 1) = False Then
   MsgBox "通過1"
    If TimeDesignation Then
     MsgBox "通過2"
      If FileDesignation = False Then   '1：ファイル,2：フォルダ
        MsgBox "通過3"
      End If
    End If
  End If
End If
End Sub

Private Sub CommandButton18_Click()
Call AIM
End Sub

Private Sub CommandButton19_Click()

Application.ScreenUpdating = False
Call HiddenSheets '隠す
AddWS ("過去Log")
ThisWorkbook.Sheets("過去Log").Visible = xlSheetVisible
ThisWorkbook.Sheets("過去Log").Visible = True
ThisWorkbook.Sheets("Test").Visible = True
Application.ScreenUpdating = True

Application.Visible = True 'Excel表示


Dim TextName As String: TextName = Me.LoadTextName '"Access.log"_要拡張子
Dim FilePath As String: FilePath = ThisWorkbook.path & "\" & TextName
Dim i As Long, j As Long

Dim fileContent As String, fileLines() As String

' テキストファイルをUTF-8で読み込む
fileContent = ReadFile(FilePath, "UTF-8")
' 改行文字で分割して配列に格納
fileLines = Split(fileContent, vbCrLf)
    
    
    
' 配列の内容を表示（デバッグ用）
For i = LBound(fileLines) To UBound(fileLines)
  Debug.Print fileLines(i)
Next i
Stop

End Sub

Private Sub CommandButton20_Click()
Call 半角へ変換
End Sub

Private Sub CommandButton21_Click()
Call HVC
End Sub

Private Sub CommandButton22_Click()
'Call test取得
'DeleteSetting "UF検査判定システム", "起動設定", "モード"
Dim tmpFaci As String: tmpFaci = GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")
Dim tmpMode As String: tmpMode = GetSetting("UF検査判定システム", "起動設定", "モード", "存在なし")
MsgBox "設備名_" & tmpFaci & "  " & "モード_" & tmpMode

SF3.Show vbModeless
End Sub

Private Sub CommandButton23_Click()
Call DelSet
End Sub

Sub TestFindRightString()
    Dim InputString As String
    Dim Result As String
    
    InputString = "HOT L-3 L-3 L-3 DL2 ANI L-3 FS4 ANF TUL CAI CAF ANI CAI ANF"
    
    Result = FindRightString(InputString)
    
    MsgBox Result
End Sub

Private Sub CommandButton24_Click()
Dim Target2 As String:       Target2 = "品質基準.accdb"
Dim TempPath2 As String:     TempPath2 = "C:\Users\xfa-ngyinazawa\Desktop\DB\品質\" 'QualityPath  '"C:\Users\xfa-ngykonpo\Desktop\DB\品質\"
Dim TableName As String
Dim fieldsfilter As String:  fieldsfilter = ""
Dim SQLfilter As String

Dim FilePath As String:      FilePath = TempPath2 & Target2 '再定義
Dim ArrQuality As Variant
Dim FieldsQuality As Variant

TableName = "ソリ4級"
FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter)
If IsEmpty(FieldsQuality) Then Exit Sub
ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter)
If IsEmpty(ArrQuality) Then Exit Sub


Dim オーダー厚 As Single:    オーダー厚 = Me.Test1
Dim オーダー板幅 As Single:  オーダー板幅 = Me.Test2
Dim オーダー板丈 As Single:  オーダー板丈 = Me.Test3

'幅丈のインデックス
Dim HitHaba As Variant: HitHaba = FindIndices(FieldsQuality, "H", オーダー板幅)
Dim HitTake As Variant: HitTake = FindIndices(FieldsQuality, "T", オーダー板丈)

Dim 幅反り公差 As Single: 幅反り公差 = SoriArray(ArrQuality, FieldsQuality, オーダー厚, オーダー板幅, オーダー板丈, HitHaba(1))
Dim 丈反り公差 As Single: 丈反り公差 = SoriArray(ArrQuality, FieldsQuality, オーダー厚, オーダー板幅, オーダー板丈, HitTake(1))

MsgBox "幅" & Format(幅反り公差, "0.0") & " " & "丈" & Format(丈反り公差, "0.0")

End Sub
Private Sub CommandButton25_Click()
Dim InputString As String
Dim Result As String

' テスト用文字列
InputString = Me.CourseTex.Text

' 関数呼び出し
Result = FindPriority(InputString)

' 結果を表示
MsgBox Result
End Sub

Private Sub CommandButton26_Click()
'Dim PageName As String:   PageName = "噛込み跡ショック跡"
'Dim folderPath As String: folderPath = "C:\Users\xfa-ngyinazawa\Desktop\品質写真\" ' ここを適切なパスに変更してください
'Dim ImagePath As String:  ImagePath = Dir(folderPath & PageName, vbDirectory)
'
'If Dir(folderPath & PageName, vbDirectory) = "" Then 'vbDirectory必要　指定すると引数setPathで指定したフォルダがあればフォルダ名を返し,なければ長さ0の文字列「""」を返します。
'    MkDir folderPath & PageName '月毎のフォルダを作成
'End If
'
'MsgBox ImagePath
'---------------------------------------------------------------------------

'Dim Kouf As String: Kouf = SaveDateLocate(ThisWorkbook.path & "\", "【梱包】" & "Test")
Dim MonthName As String: MonthName = Format(Date, "yyyy.mm")
Dim SrvPath As String:   SrvPath = ThisWorkbook.path & "\" & "L-1\L-1_履歴\" & MonthName
Call CreateFolder(SrvPath)
End Sub

Private Sub CommandButton27_Click()
Dim TableName As String: TableName = "Q_S_QA_124_別表1"
Dim FPath As String
FPath = "C:\Users\xfa-ngyinazawa\Desktop\DB\品質\"
Dim target As String:    target = "品質基準.accdb"
Dim FilePath As String:  FilePath = FPath & target 'Lot
Dim SQLfilter As String: SQLfilter = "[欠陥名]='" & Me.HTbox & "'"
Dim TempHiki As Variant: TempHiki = GetRecordsArr(FilePath, TableName, SQLfilter)


Dim tMP As String:       tMP = TempHiki(1, 4)
Dim FinalText As String: FinalText = ""
Dim i As Long

' コンマがある場所で改行を挿入
For i = 1 To Len(tMP)
  If Mid(tMP, i, 1) = "," Then
      FinalText = FinalText & vbCrLf
  Else
      FinalText = FinalText & Mid(tMP, i, 1)
  End If
Next i

Dim HasNewLine As Long: HasNewLine = InStr(FinalText, vbCrLf)
Me.A1.MultiLine = True
Me.A1 = FinalText
End Sub
Private Sub CommandButton28_Click()
Dim Arry_L(1 To 3, 1 To 2) As Variant
Arry_L(1, 1) = "AAAA"
Arry_L(2, 1) = "BBBBBB"
Arry_L(3, 1) = "CCCCCCCCC"
Arry_L(1, 2) = "まじ"
Arry_L(2, 2) = "まんじ"
Arry_L(3, 2) = "てｘｃｔ"
Dim Temp As String: Temp = AddPagesArrN(Me.TabST, Arry_L)

Dim SetTable As String: SetTable = AddPagesForSubFolders(Me.MP1) 'ページ追加
End Sub

Private Sub CommandButton29_Click()
Dim i As Long, FPath As String: FPath = ThisWorkbook.path & "\【測定DB】梱包\梱包検査ﾃﾞｰﾀ.accdb"
Dim sh As Worksheet:            Set sh = ThisWorkbook.Sheets("保存用")

Dim n As Long:         n = UR(sh, 1) + 1
Dim EndR As Long:      EndR = UR(sh, 1)
Dim EndC As Long:      EndC = UC(sh, 1)
Dim AdoRange As Range: Set AdoRange = Range(sh.Cells(EndR, 1), sh.Cells(EndR, EndC))

Dim AdoVar():          AdoVar = AdoRange
Dim Count As Long:     Count = UBound(AdoVar, 2) '最終列

ReDim AdoTemp(0 To Count) As Variant
AdoTemp(0) = "Test"
Dim MinL As Long:      MinL = LBound(AdoVar, 1)
Dim MaxL As Long:      MaxL = UBound(AdoVar, 2)

For i = MinL To MaxL
  AdoTemp(i) = AdoVar(1, i)
Next i

Dim LotNoT As String:  LotNoT = AdoTemp(1)
Dim Fil As String:     Fil = "[検査開始終了] ='" & LotNoT & " '"
'Stop
Call adoSQL(FPath, "検査ﾃﾞｰﾀ", "UPDATE", AdoTemp, "", True, Fil) '検査開始終了LOTNO
End Sub

Private Sub CommandButton3_Click()
'Dim tmpFaci As String: tmpFaci = GetSetting("UF検査判定システム", "起動設定", "設備名", "") ' "存在なし"
Call Setting_Path
Dim TempTime: TempTime = ShiftTime()
Call Set_Shift

Dim TimeOne As String: TimeOne = "07:59" 'Format$(Now, "hh:mm")

MsgBox YeaDayrCheck(TimeOne)

End Sub

Private Sub CommandButton32_Click()
Dim Temp As Variant: Temp = StraightEdge("AAA")
If IsEmptyArray(Temp) Then
MsgBox "配列なし"
Exit Sub
End If

MsgBox Temp(1, 1) & "_" & Temp(2, 1)

End Sub

Private Sub CommandButton4_Click()
Call LoadImagesIntoMultiPageFromTemp(Me.MP1)
End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub deCle(flag As Boolean)
Select Case True
Case flag
 Me.T1 = ""
' Call NewKen_Click
End Select
End Sub

Private Sub CommandButton6_Click()
AddWS ("過去Log")
ThisWorkbook.Sheets("過去Log").Visible = xlSheetVisible
Dim Lname As String: Lname = Me.L3 'Replace(UF1.LineName.Caption & UF1.LineNo.Caption, " ", "") '半角スペースを消す_2021.12.15
Dim LotNo As String:    LotNo = Me.LoadLot
Dim Worker As String:   Worker = Me.LoadWorker
Dim TextName As String: TextName = Me.LoadTextName '"Access.log"_要拡張子
Dim Rea As String:      Rea = Me.Reason
'If ReadCheck(ThisWorkbook.path & "\" & TextName, 1) = False Then
'  MsgBox ThisWorkbook.path & "\" & TextName & "が開かれているか、存在しません。" & vbCrLf & "ファイルを確認してください。", vbCritical
'  Exit Sub
'End If

Dim LineBuff, tMP, i As Long
Open ThisWorkbook.path & "\" & TextName For Input As #1 '最大行チェック
Do Until EOF(1)
  Line Input #1, tMP
  i = i + 1
Loop
Close #1

Dim LotBuf
Dim I_count As Long: I_count = i
ReDim LineBuff(1 To I_count)
i = 0

Open (ThisWorkbook.path & "\" & TextName) For Input As #1
  Do Until EOF(1)
    Line Input #1, tMP
    i = i + 1
    ThisWorkbook.Sheets("過去Log").Cells(i, 1) = tMP
    LineBuff(i) = tMP
  Loop
Close #1


'str = Replace(対象文字列, 置換前文字, 置換後文字)
Dim a As Long, SameCount As Long, LineFlag As Boolean
For a = 1 To UBound(LineBuff)
 If (InStr(LineBuff(a), LotNo) <> 0) Then 'LotNoの有無
  LotBuf = Split(LineBuff(a), ":")
  Dim NowLine As String: NowLine = Replace(LotBuf(UBound(LotBuf)), " ", "")
  If NowLine = Lname Then
    SameCount = a  'LotNo一致行 かつ　ライン一致_2021.11.29　　　And (LotBuf(UBound(LotBuf)) = Lname)
    LineFlag = True 'ライン一致でフラグを立てる_2021.12.15
    Exit For
  End If
 End If
Next a

If SameCount = 0 Then '2021.04.12*エラー用
  MsgBox "一致するLotNoがありません", vbCritical
  Exit Sub
ElseIf LineFlag = False Then
  MsgBox "一致するﾗｲﾝがありません", vbCritical
  Exit Sub
End If
a = 1
Dim buf
Dim S(1 To 14) '耳付き丈3,4追加の為+2
Dim U(1 To 14)
Dim C(1 To 14)
Dim D(1 To 14)
Dim E(1 To 14)
Dim Slate
Dim Ulate
Dim Clate
Dim Dlate
Dim Elate

i = 0
Dim EndCount As Long
For a = SameCount + 1 To UBound(LineBuff)
 If InStr(LineBuff(a), "Lot登録") <> 0 Then
  EndCount = a '別ロット検索まで*測定値が別ロットのものになる可能性
  Exit For
 End If
Next a
a = 1
'*******検索ロットの登録が何個もあるとエラーになる、ほしい範囲以外のLot登録は消すこと
If EndCount = 0 Then EndCount = UBound(LineBuff) '別ロット検索がなければ最終行まで
For a = SameCount To EndCount 'Lot一致から別Lot検索まで
  For i = 1 To 14 '耳付き丈3,4追加の為+2
  Dim LineName As String
  Select Case True
  Case InStr(LineBuff(a), "(S" & i & ")") <> 0 '検査位置検索
    buf = Split(LineBuff(a), ":") '文字列分割
    LineName = Replace(buf(UBound(buf)), " ", "")
    If LineName = Lname Then S(i) = buf(3)
  Case InStr(LineBuff(a), "(U" & i & ")") <> 0
    buf = Split(LineBuff(a), ":")
    LineName = Replace(buf(UBound(buf)), " ", "")
    If LineName = Lname Then U(i) = buf(3)
  Case InStr(LineBuff(a), "(C" & i & ")") <> 0
    buf = Split(LineBuff(a), ":")
    LineName = Replace(buf(UBound(buf)), " ", "")
    If LineName = Lname Then C(i) = buf(3)
  Case InStr(LineBuff(a), "(D" & i & ")") <> 0
    buf = Split(LineBuff(a), ":")
    LineName = Replace(buf(UBound(buf)), " ", "")
    If LineName = Lname Then D(i) = buf(3)
  Case InStr(LineBuff(a), "(E" & i & ")") <> 0
    buf = Split(LineBuff(a), ":")
    LineName = Replace(buf(UBound(buf)), " ", "")
    If LineName = Lname Then E(i) = buf(3)
  End Select
  Next i
  
  Select Case True '2021.07.16_ﾗﾃﾗﾙﾎﾞｰ
  Case InStr(LineBuff(a), "(LateralS)") <> 0
    buf = Split(LineBuff(a), ":")
    LineName = Replace(buf(UBound(buf)), " ", "")
    If LineName = Lname Then Slate = buf(3)
  Case InStr(LineBuff(a), "(LateralU)") <> 0
    buf = Split(LineBuff(a), ":")
    LineName = Replace(buf(UBound(buf)), " ", "")
    If LineName = Lname Then Ulate = buf(3)
  Case InStr(LineBuff(a), "(LateralC)") <> 0
    buf = Split(LineBuff(a), ":")
    LineName = Replace(buf(UBound(buf)), " ", "")
    If LineName = Lname Then Clate = buf(3)
  Case InStr(LineBuff(a), "(LateralD)") <> 0
    buf = Split(LineBuff(a), ":")
    LineName = Replace(buf(UBound(buf)), " ", "")
    If LineName = Lname Then Dlate = buf(3)
  Case InStr(LineBuff(a), "(LateralE)") <> 0
    buf = Split(LineBuff(a), ":")
    LineName = Replace(buf(UBound(buf)), " ", "")
    If LineName = Lname Then Elate = buf(3)
  End Select
Next a
Call AddWS("Accesslog")
Dim sh As Worksheet: Set sh = ShJuge("Accesslog")
Dim n As Long: n = UR(sh, 1)
'---------------------------------------------------------------転送処理
sh.Cells(n, 1) = LotNo 'LotNo
sh.Cells(n, 17) = Worker '検査員
sh.Cells(n, 18) = TextName '"Access.log添付" 'ラテラルボー追加のため移動_2021.07.16'2022.06_耳付き対応
sh.Cells(n, 19) = Rea 'Load理由

'Stop

Dim Sran As Range: Set Sran = Range(sh.Cells(n, 2), sh.Cells(n, 1 + UBound(S))) '耳付きレイアウトに合わせる
Dim Uran As Range: Set Uran = Range(sh.Cells(n + 1, 2), sh.Cells(n + 1, 1 + UBound(U)))
Dim Cran As Range: Set Cran = Range(sh.Cells(n + 2, 2), sh.Cells(n + 2, 1 + UBound(C)))
Dim Dran As Range: Set Dran = Range(sh.Cells(n + 3, 2), sh.Cells(n + 3, 1 + UBound(D)))
Dim Eran As Range: Set Eran = Range(sh.Cells(n + 4, 2), sh.Cells(n + 4, 1 + UBound(E)))
Sran = S
Uran = U
Cran = C
Dran = D
Eran = E
'ﾗﾃﾗﾙﾎﾞｰ
sh.Cells(n, 16) = Slate
sh.Cells(n + 1, 16) = Ulate
sh.Cells(n + 2, 16) = Clate
sh.Cells(n + 3, 16) = Dlate
sh.Cells(n + 4, 16) = Elate
ThisWorkbook.Sheets("Accesslog").Visible = True
End Sub

Private Sub CommandButton7_Click()
Dim Lname As String: Lname = "Test"
Dim Lotname As String: Lotname = "B"
If WsCheck("Daily" & "_" & Lname) = False Then Exit Sub
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Daily" & "_" & Lname)
Call AddWS("梱包枚数計算" & "_" & Lname)
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("梱包枚数計算" & "_" & Lname)
Range(sh.Cells(3, 5), sh.Cells(UR(sh, 5), 5)).Interior.Pattern = xlNone '色ｸﾘｱ
Dim Dran As Range: Set Dran = Range(sh.Cells(3, 1), sh.Cells(UR(sh, 1), 13))
Dim i As Long, x As Long, y As Long, n As Long: n = 0
Dim total(), BadFl As Boolean
Dim CountDic As Object: Set CountDic = CreateObject("Scripting.Dictionary") '検入枚数連想配列
total = Dran '配列
Dim Kara As String: Kara = "" 'LotNoで埋める
For i = 1 To UBound(total, 1)
 If total(i, 1) = "" Then
   total(i, 1) = Kara
 Else 'LoTNo入力があるとき
  Kara = total(i, 1)
  If CountDic.exists(total(i, 1)) = False Then '検入枚数とLOTNO格納
    CountDic(total(i, 1)) = total(i, 5)
  End If
 End If
Next i
Dim LotDic As Object: Set LotDic = CreateObject("Scripting.Dictionary") '梱包枚数連想配列
For i = LBound(total, 1) To UBound(total, 1)
  If LotDic.exists(total(i, 1)) = False Then '一致でない時格納
    x = val(total(i, 12)) * val(total(i, 13)) 'LotNoと梱包枚数
    LotDic(total(i, 1)) = x
  Else
    y = val(total(i, 12)) * val(total(i, 13))
    LotDic(total(i, 1)) = val(LotDic(total(i, 1))) + y 'LotNoが重複時梱包枚数を加算
  End If
Next i
Dim VarName As Variant, j As Long: j = 2
ws.Cells(1, 1) = "LotNo"
ws.Cells(1, 2) = "梱包枚数"
For Each VarName In LotDic '梱包数とLotNo書き込み
  ws.Cells(j, 1).Value = VarName
  ws.Cells(j, 2).Value = LotDic.Item(VarName)
  j = j + 1
Next
Dim VarName2 As Variant
ws.Cells(1, 3) = "LotNo"
ws.Cells(1, 4) = "検入枚数"
j = 2
For Each VarName2 In CountDic '検入枚数とLotNo書き込み
  ws.Cells(j, 3).Value = VarName2
  ws.Cells(j, 4).Value = CountDic.Item(VarName2)
  j = j + 1
Next
Dim LotNo As Variant, Posi As Long
If CountDic.exists(Lotname) Then '梱包枚数連想配列側にあるLotNoを検入枚数連想配列側でチェック ’有る時
  If val(LotDic.Item(Lotname)) > val(CountDic.Item(Lotname)) Then '梱包枚数が登録検入枚数より大きい
    For i = 1 To UBound(total, 1)
      If total(i, 1) = Lotname Then 'シート配列から探す
        If total(i, 5) <> "" Then '検入枚数入力がある
          Posi = i '位置格納
          Exit For
        End If
      End If
    Next i
    sh.Cells(Posi + 2, 5).Interior.Color = RGB(255, 128, 128)
    BadFl = True
  End If
End If
Set LotDic = Nothing
Set CountDic = Nothing
If BadFl Then '検入枚数より大きい
  MsgBox "梱包数と検入枚数が一致しません", vbCritical
'  UF1.PrintB.Tag = "stop" '停止フラグ　一致しない場合は印刷停止　2020.11.24
  Call NameList_Open(3)
End If
End Sub

Private Sub CommandButton8_Click()
Dim saveSh As Worksheet:    Set saveSh = ThisWorkbook.Sheets("印刷用" & "_" & "Test")
Dim Defective As Worksheet: Set Defective = ThisWorkbook.Sheets("不良要因印刷" & "_" & "Test")
Dim KouPath As String:  KouPath = "C:\検査履歴"

Dim LM As String:       LM = "Test"
'印刷帳票
Dim F As String:        F = SaveDateLocate(KouPath & "\", LM)   'ﾊﾟｽ,ﾌｧｲﾙﾗｲﾝ名_必要ﾌｫﾙﾀﾞ生成とﾌｧｲﾙﾊﾟｽ
Call PDFCREATE(F, saveSh, saveSh.Cells(2, 3)) 'ﾊﾟｽ,保存ｼｰﾄ,保存ﾌｧｲﾙ名
Dim NewFiles As String: NewFiles = NewFileName(F) '生成後一番新しいﾌｧｲﾙ名
Dim Kouf As String:     Kouf = SaveDateLocate(FactoryyPath, "【梱包】" & LM)
FileCopy F & "\" & NewFiles, Kouf & "\" & NewFiles 'ﾌｧｲﾙｺﾋﾟｰ

'不良要因'2025.1.17
Call PDFCREATE(F, Defective, Defective.Cells(1, 2) & "_不良要因")
NewFiles = NewFileName(F) '生成後一番新しいﾌｧｲﾙ名

Dim FilePath As String: FilePath = F & "\" & NewFiles
' ファイルが作成されるまで待機
Do While Dir(FilePath) = ""
  DoEvents ' 他の処理を実行
Loop

FileCopy F & "\" & NewFiles, Kouf & "\" & NewFiles

End Sub

Public Sub DB元ファイル作成() 'デバッグ
Dim SavePath As String: SavePath = ThisWorkbook.path
Dim DbFileName As String: DbFileName = "パレット種類_外径条件.accdb"
Dim TmpSh As Worksheet: Set TmpSh = Worksheets("パレット種類_外径条件")
Dim wsData As Variant: wsData = Range(TmpSh.Cells(1, 1), TmpSh.Cells(UR(TmpSh, 1), UC(TmpSh)))
Call ExpAccdb(wsData, "パレット種類_外径条件", DbFileName, SavePath, True) '2つ目以降はFlag=False
End Sub
Private Sub CommandButton9_Click()
Call DB元ファイル作成

'Dim fPath As String: fPath = ThisWorkbook.path & "\Lot情報履歴.accdb"
'Dim a(1 To 10) As Variant, i As Long
''a(1, 1) = "AIM"
''a(1, 2) = "2021.10.22_Test5"
''a(2, 1) = "テスト中_1回目入ります！！！！"
''For i = 3 To 10
'' a(1, i) = "X"
''Next i
'Dim Fil As String: Fil = "[LOTNO] ='" & a(2) & " '"
'Call adoSQL(fPath, "Lot情報履歴", "UPDATE", a, "", True, "") '(fPath, "検査ﾃﾞｰﾀ", "DELETE", a, "ﾛｯﾄ番号") 'UPDATE,ADDNEW

End Sub

Private Sub MP1_Change()
Call LoadImagesIntoMultiPageFromTemp(Me.MP1)
End Sub

Private Sub T2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Private Sub T3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Private Sub T2_AfterUpdate()
'Dim maxLength As Integer: maxLength = 3
'Dim MinLength As Integer: MinLength = 5
'
'With Me.T2
'  If Len(.text) > maxLength Then 'Or Len(.text) < MinLength Then
'      MsgBox maxLength & "文字以下で入力して下さい。"
'      If Len(.text) >= 10 Then .text = Mid(.text, 1, 10)
'  End If
'End With
End Sub
Private Sub T3_Change()
'Select Case Len(Me.T3)
'Case 7: Me.T3 = Me.T3
'Case 8: Me.T3 = Mid(Me.T3, 1, 7)
'End Select
'If Len(Me.T3) <> 7 Then Exit Sub
'Debug.Print Len(Me.T3)
End Sub
'＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞
Private Function NameLoad(no As String) As String
If no = "" Then Exit Function
Dim TableName As String: TableName = "名簿"
Dim FPath As String: FPath = PackingPath
'ﾌｧｲﾙ名
Dim target As String:   target = "伝送用ファイル" & ".accdb"
'ﾌｧｲﾙﾊﾟｽ
Dim FilePath As String: FilePath = FPath & target
  If ReadCheck(FilePath, 1) = False Then '1：ファイル,2：フォルダ
    MsgBox FilePath & "がありません" & "_NameConv", vbCritical
    Exit Function
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

Dim Sfile As String: Sfile = "chime1.wav"
Dim Spath As String: Spath = "C:\VBAWAV\"
Dim SoundFile As String: SoundFile = Spath & Sfile
Dim rc As Long

If IsEmpty(Temp) Then 'emptyの場合停止
  Sfile = "ナンバーNO.wav"
  SoundFile = Spath & Sfile
  If ReadCheck(SoundFile, 1) = False Then
  Exit Function
  End If
  rc = mciSendString("Play " & SoundFile, "", 0, 0)
  NameLoad = ""
  Exit Function

Else
  If ReadCheck(SoundFile, 1) = False Then
  Exit Function
  End If
  rc = mciSendString("Play " & SoundFile, "", 0, 0)
  NameLoad = Temp(1, C_名前)
End If
End Function
Private Sub TextTest_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyTab Then
  Me.TestLabel.Caption = ""
  Dim Temp As String
  If Len(Me.TextTest.Text) <> 13 Then
  Exit Sub
  Else
  Dim WorName As String: WorName = NameLoad(Me.TextTest.Value)
  End If
  Me.TestLabel.Caption = WorName
  Me.TextTest = ""
  KeyCode = 0 'tab無効化
'  Me.NowFocus.Caption = ActiveControl.Name
End If
End Sub
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub ListBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim D As DataObject
If Button <> 1 Then Exit Sub 'マウス左ボタンのドラッグ時に対応
Set D = New DataObject 'データオブジェクトに現在の選択地を格納
If Me.ListBox1.Caption <> "" Then

  Me.ListBox1.Tag = Me.ListBox1.Caption
  Dim n As String:  n = Me.ListBox1.Tag
  Dim TB As String: TB = ObjMoveText(Me.ListBox1)
  Me.MoveFlag.Text = TB
  Me.Tagtext.Text = n
  
  D.SetText ListBox1.Caption
  D.StartDrag
End If
End Sub
Private Sub ListBox2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim D As DataObject
If Button <> 1 Then Exit Sub
Set D = New DataObject
If Me.ListBox2.Caption <> "" Then

  Me.ListBox2.Tag = Me.ListBox2.Caption
  Dim n As String:  n = Me.ListBox2.Tag
  Dim TB As String: TB = ObjMoveText(Me.ListBox2)
  Me.MoveFlag.Text = TB
  Me.Tagtext.Text = n

  D.SetText ListBox2.Caption
  D.StartDrag
End If
End Sub
Private Sub ListBox3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim D As DataObject
If Button <> 1 Then Exit Sub
Set D = New DataObject
If Me.ListBox3.Caption <> "" Then

  Me.ListBox3.Tag = Me.ListBox3.Caption
  Dim n As String:  n = Me.ListBox3.Tag
  Dim TB As String: TB = ObjMoveText(Me.ListBox3)
  Me.MoveFlag.Text = TB
  Me.Tagtext.Text = n

  D.SetText ListBox3.Caption
  D.StartDrag
End If
End Sub

'○ドラッグ（In)
Private Sub ListBox1_BeforeDragOver _
    (ByVal Cancel As MSForms.ReturnBoolean, _
    ByVal data As MSForms.DataObject, _
    ByVal x As Single, ByVal y As Single, _
    ByVal DragState As MSForms.fmDragState, _
    ByVal Effect As MSForms.ReturnEffect, _
    ByVal Shift As Integer)
'マウスが入った時のイベント
'Cancel=TrueでDrag&Drop継続
Dim Temp: Temp = data.GetText 'ドラッグ中のデータ
Dim st As String: st = Me.MoveFlag.Text
If Me.ListBox1.Caption = Temp Then
  Cancel = False '運ぶ先のデータが同じとき
'  Me(ST).Caption = Me(ST).Tag
End If
If Me.ListBox1.Caption <> Temp Then
  Cancel = True
  Me(st).Caption = ""
End If
End Sub
Private Sub ListBox2_BeforeDragOver _
    (ByVal Cancel As MSForms.ReturnBoolean, _
    ByVal data As MSForms.DataObject, _
    ByVal x As Single, ByVal y As Single, _
    ByVal DragState As MSForms.fmDragState, _
    ByVal Effect As MSForms.ReturnEffect, _
    ByVal Shift As Integer)
Dim Temp: Temp = data.GetText 'ドラッグ中のデータ
Dim st As String: st = Me.MoveFlag.Text
If Me.ListBox2.Caption = Temp Then
  Cancel = False '運ぶ先のデータが同じとき
'  Me(ST).Caption = Me(ST).Tag
End If
If Me.ListBox2.Caption <> Temp Then
  Cancel = True
  Me(st).Caption = ""
End If
End Sub
Private Sub ListBox3_BeforeDragOver _
    (ByVal Cancel As MSForms.ReturnBoolean, _
    ByVal data As MSForms.DataObject, _
    ByVal x As Single, ByVal y As Single, _
    ByVal DragState As MSForms.fmDragState, _
    ByVal Effect As MSForms.ReturnEffect, _
    ByVal Shift As Integer)
Dim Temp: Temp = data.GetText 'ドラッグ中のデータ
Dim st As String: st = Me.MoveFlag.Text
If Me.ListBox3.Caption = Temp Then
  Cancel = False '運ぶ先のデータが同じとき
  Me(st).Caption = Me(st).Tag
End If
If Me.ListBox3.Caption <> Temp Then
  Cancel = True
  Me(st).Caption = ""
End If
End Sub

'○ドロップ
Private Sub ListBox1_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, _
        ByVal Action As MSForms.fmAction, _
        ByVal data As MSForms.DataObject, ByVal x As Single, _
        ByVal y As Single, ByVal Effect As MSForms.ReturnEffect, _
        ByVal Shift As Integer)

'ドラッグ時のみドラッグされたデータを追加
If Action = fmActionDragDrop Then
  If MsgBox("dataをドロップしますか？", vbYesNo + vbQuestion, "確認") = vbYes Then
      Me.ListBox1.Caption = data.GetText
  End If
End If
data.clear 'DataObjectのデータクリア
End Sub
Private Sub ListBox2_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, _
        ByVal Action As MSForms.fmAction, _
        ByVal data As MSForms.DataObject, ByVal x As Single, _
        ByVal y As Single, ByVal Effect As MSForms.ReturnEffect, _
        ByVal Shift As Integer)

If Action = fmActionDragDrop Then
  If MsgBox("dataをドロップしますか？", vbYesNo + vbQuestion, "確認") = vbYes Then
      Me.ListBox2.Caption = data.GetText
  End If
End If
data.clear 'DataObjectのデータクリア
End Sub
Private Sub ListBox3_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, _
        ByVal Action As MSForms.fmAction, _
        ByVal data As MSForms.DataObject, ByVal x As Single, _
        ByVal y As Single, ByVal Effect As MSForms.ReturnEffect, _
        ByVal Shift As Integer)

If Action = fmActionDragDrop Then
  If MsgBox("dataをドロップしますか？", vbYesNo + vbQuestion, "確認") = vbYes Then
      Me.ListBox3.Caption = data.GetText
  End If
End If
data.clear 'DataObjectのデータクリア
End Sub

Private Sub UserForm_Initialize()
Application.Visible = True 'Excel表示

Call Setting_Path
End Sub

Private Sub Wildcard_Click()

MsgBox FindInList(Me.TextTest, Me.A1.Text, "")

End Sub
