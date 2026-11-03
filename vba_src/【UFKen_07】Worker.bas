Attribute VB_Name = "【UFKen_07】Worker"
Option Explicit
Public Start1 As String
Public Start2 As String
Public Start3 As String
Public Start昼 As String
Public End1 As String
Public End2 As String
Public End3 As String
Public End昼 As String

Public Sub Set_Shift() '直の開始終了時間の定義
'ShiftTimeで時間の配列
Dim i As Long, a As Long
Dim TempTime: TempTime = ShiftTime()

If IsEmptyArray(TempTime) Then Exit Sub

For a = LBound(TempTime, 1) To UBound(TempTime, 1)
  Select Case TempTime(a, 1)
  Case "1"
    Start1 = TempTime(a, 2)
    End1 = TempTime(a, 3)
  Case "2"
    Start2 = TempTime(a, 2)
    End2 = TempTime(a, 3)
  Case "3"
    Start3 = TempTime(a, 2)
    End3 = TempTime(a, 3)
  Case "昼"
    Start昼 = TempTime(a, 2)
    End昼 = TempTime(a, 3)
  End Select
Next a
End Sub

Public Function NameConv(no As String) As String 'ﾊﾞｰｺｰﾄﾞ変換
'------------------------------DBから読み取る
If no = "" Then Exit Function
Dim TableName As String:  TableName = "名簿"
Dim FPath As String
Select Case True
Case UF1.Test:  FPath = PackingPath
Case Else:      FPath = QualityPath '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\検査\"
End Select
Dim target As String:     target = "伝送用ファイル" & ".accdb" 'ﾌｧｲﾙ名
Dim FilePath As String:   FilePath = FPath & target 'ﾌｧｲﾙﾊﾟｽ
If ReadCheck(FilePath, 1) = False Then '1：ファイル,2：フォルダ
  If FPath = QualityPath Then FPath = PackingPath
  FilePath = FPath & target
  If ReadCheck(FilePath, 1) = False Then
    MsgBox "日稲:" & target & "がありません" & "_ReadCheck_NameConv", vbCritical
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
  Call 音楽未登録
  NameConv = ""
  Exit Function
Else
  NameConv = Temp(1, C_名前)
End If
'Stop

'------------------------------他ﾌﾞｯｸから必要シートを読み取る
'Application.ScreenUpdating = False  '画面更新停止
'Application.DisplayAlerts = False
'Dim ListPass As String
'  Select Case True
'  Case UF1.Test:  ListPass = ThisWorkbook.Path & "\"
'  Case Else:      ListPass = "\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\検査\"
'  End Select
'Dim ListName As String: ListName = "UF検査_原紙類.xlsm"
'If ReadCheck(ListPass & ListName, 1) = False Then
'  MsgBox ListPass & ListName & "がありません" & "_NameConv", vbCritical
'  Exit Function
'End If
'Dim KenBook As Workbook: Set KenBook = Workbooks.Open(ListPass & ListName, UpdateLinks:=3)
'Dim WS As Worksheet 'ﾜｰｸｼｰﾄ
'Dim WsCheck As Boolean: WsCheck = False 'ｼｰﾄを探す
'For Each WS In KenBook.Worksheets
'  If WS.Name = "名簿" Then
'    WsCheck = True
'    Exit For
'  End If
'Next WS
'Set WS = Nothing
'If WsCheck Then
'   Dim NameDic As Object: Set NameDic = CreateObject("Scripting.Dictionary") '連想配列
'   Dim SH As Worksheet: Set SH = KenBook.Sheets("名簿")
'   Dim i As Long, LR As Long: LR = UR(SH, 1) 'ｶｰﾄﾞNO最終行
'   Dim NameList(), ListRange As Range: Set ListRange = Range(SH.Cells(1, 1), SH.Cells(LR, 2)) 'ｶｰﾄﾞNOと名前
'   NameList() = ListRange '配列化
'Else
'   MsgBox ListPass & "「名簿」ｼｰﾄがありません。" & vbCrLf & "やり直してください。", vbExclamation
'   Exit Function
'End If
'KenBook.Close
'Application.DisplayAlerts = True
'Application.ScreenUpdating = True

''------------------------------ブック内にあるシートから必要データを読み取る
'If WsCheck("名簿") = False Then Exit Function '-------------------------ｼｰﾄなし
'Dim NameDic As Object: Set NameDic = CreateObject("Scripting.Dictionary") '連想配列
'Dim SH As Worksheet: Set SH = ThisWorkbook.Sheets("名簿")
'Dim i As Long, LR As Long: LR = UR(SH, 1) 'ｶｰﾄﾞNO最終行
'Dim NameList(), ListRange As Range: Set ListRange = Range(SH.Cells(1, 1), SH.Cells(LR, 2)) 'ｶｰﾄﾞNOと名前
'NameList() = ListRange '配列化
'For i = LBound(NameList, 1) To UBound(NameList, 1)
'  If NameDic.Exists(NameList(i, 1)) = False Then 'NOが一致でない時格納
'     NameDic(NameList(i, 1)) = NameList(i, 2)
'  End If
'Next i
'Dim a() As Variant, b() As Variant
'a() = NameDic.Keys()
'b() = NameDic.Items()
'If NameDic.Exists(NO) = True Then '一致がある *存在する場合は、真 (True) を返します
'  NameConv = NameDic.Item(NO) '名簿から数列とセットの名前を吐き出す
'Else '名簿登録なし
'  Call 音楽未登録
'  NameConv = ""
'End If
'Set NameDic = Nothing
End Function

Public Function YeaDayrCheck(DTime As String) As String  '日付を返す   引数=時間
Dim dt As String:       dt = DateAdd("d", -1, Now) '夜勤の日付またぎ前日扱い
Dim strDate As String:  strDate = Format$(Now, "yyyy.mm.dd")
Dim StrDate2 As String: StrDate2 = Format$(dt, "yyyy.mm.dd")

Dim tmpFaci As String:   tmpFaci = GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")

If Not (tmpFaci = "L-1" Or tmpFaci = "LVC" Or tmpFaci = "HVC" Or tmpFaci = "Test") Then
  YeaDayrCheck = (strDate)
Else
  If DTime > CDate(Start1) And DTime < CDate(End1) Then
    YeaDayrCheck = (strDate)
  ElseIf DTime >= CDate(Start2) And DTime < CDate(End2) Then
    YeaDayrCheck = (strDate)
  ElseIf DTime >= CDate(Start3) And DTime <= CDate("23:59") Then
    YeaDayrCheck = (strDate)
  ElseIf DTime >= CDate("00:00") And DTime <= CDate(End3) Then
    YeaDayrCheck = (StrDate2)
  Else
    YeaDayrCheck = (strDate) '23:59のタイミングをひっかけられない
  End If
End If

'If UF校正.Visible = True Then
'  Select Case UF校正.LineName
'  Case "ﾊﾞﾗﾝｻｰ", "ﾄｯﾄ"
'    YeaDayrCheck = (strDate)
'  Case Else
'    If DTime > CDate("7:00") And DTime < CDate("15:00") Then
'      YeaDayrCheck = (strDate)
'    ElseIf DTime >= CDate("15:00") And DTime < CDate("22:50") Then
'      YeaDayrCheck = (strDate)
'    ElseIf DTime >= CDate("22:50") And DTime <= CDate("23:59") Then
'      YeaDayrCheck = (strDate)
'    ElseIf DTime >= CDate("00:00") And DTime <= CDate("7:00") Then
'      YeaDayrCheck = (StrDate2)
'    Else
'      YeaDayrCheck = (strDate) '23:59のタイミングをひっかけられない
'    End If
'  End Select
'Exit Function
'End If
'
'If UF1.Tott = True Then
'  YeaDayrCheck = (strDate)
'ElseIf UF1.Bala = True Then
'  YeaDayrCheck = (strDate)
'Else
'  If DTime > CDate("7:00") And DTime < CDate("15:00") Then
'    YeaDayrCheck = (strDate)
'  ElseIf DTime >= CDate("15:00") And DTime < CDate("22:50") Then
'    YeaDayrCheck = (strDate)
'  ElseIf DTime >= CDate("22:50") And DTime <= CDate("23:59") Then
'    YeaDayrCheck = (strDate)
'  ElseIf DTime >= CDate("00:00") And DTime <= CDate("7:00") Then
'    YeaDayrCheck = (StrDate2)
'  Else
'    YeaDayrCheck = (strDate) '23:59のタイミングをひっかけられない
'  End If
'End If
End Function
Public Function DayCheck(DTime As String) As String  '日付時間を返す
'24時以降を表示するための定義_2021.01.19
Dim wTime As Date
Dim Keika As Date: Keika = Now - wTime
Dim hh As Long:    hh = 24 'CLng(Month(Keika)) * 12'*単純計算で良さそう2021.03.22
Dim buf:           buf = Format(Now, "h") + hh
'--------------------------------------------
Dim dt As String:        dt = DateAdd("d", -1, Now) '夜勤の日付またぎ前日扱い
Dim strDate As String:   strDate = Format$(Now, "yyyy.mm.dd  hh:nn")
Dim StrDate2 As String:  StrDate2 = Format$(dt, "yyyy.mm.dd") & "  " & buf & Format$(dt, ":nn")

Dim tmpFaci As String:   tmpFaci = GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")

If Not (tmpFaci = "L-1" Or tmpFaci = "LVC" Or tmpFaci = "HVC" Or tmpFaci = "Test") Then
  DayCheck = (strDate)
Else
  If DTime > CDate(Start1) And DTime < CDate(End1) Then
    DayCheck = (strDate)
  ElseIf DTime >= CDate(Start2) And DTime < CDate(End2) Then
    DayCheck = (strDate)
  ElseIf DTime >= CDate(Start3) And DTime <= CDate("23:59") Then
    DayCheck = (strDate)
  ElseIf DTime >= CDate("00:00") And DTime <= CDate(End3) Then
    DayCheck = (StrDate2)
  Else
    DayCheck = (strDate) '23:59のタイミングをひっかけられない
  End If
End If

'If UF1.Tott = True Then
'  DayCheck = (strDate)
'ElseIf UF1.Bala = True Then
'  DayCheck = (strDate)
'Else
'  If DTime > CDate("7:00") And DTime < CDate("15:00") Then
'    DayCheck = (strDate)
'  ElseIf DTime >= CDate("15:00") And DTime < CDate("22:50") Then
'    DayCheck = (strDate)
'  ElseIf DTime >= CDate("22:50") And DTime <= CDate("23:59") Then
'    DayCheck = (strDate)
'  ElseIf DTime >= CDate("00:00") And DTime <= CDate("7:00") Then
'    DayCheck = (StrDate2)
'  Else
'    DayCheck = (strDate) '23:59のタイミングをひっかけられない
'  End If
'End If
End Function
Public Function TimeCheck(DTime As String) As String '直を返す
Dim dt As String:       dt = DateAdd("d", -1, Now) '夜勤の日付またぎ前日扱い
Dim strDate As String:  strDate = Format$(Now, "yyyy""年""m""月""d""日"";@")
Dim StrDate2 As String: StrDate2 = Format$(dt, "yyyy""年""m""月""d""日"";@")

Dim tmpFaci As String:   tmpFaci = GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")

If Not (tmpFaci = "L-1" Or tmpFaci = "LVC" Or tmpFaci = "HVC" Or tmpFaci = "Test") Then
  TimeCheck = "日勤"
Else
  If DTime > CDate(Start1) And DTime < CDate(End1) Then
    TimeCheck = "1直"
  ElseIf DTime >= CDate(Start2) And DTime < CDate(End2) Then
    TimeCheck = "2直"
  ElseIf DTime >= CDate(Start3) And DTime <= CDate("23:59") Then
    TimeCheck = "3直"
  ElseIf DTime >= CDate("00:00") And DTime <= CDate(End3) Then
    TimeCheck = "3直"
  Else
     TimeCheck = "3直" '23:59のタイミングをひっかけられない
  End If
End If

'If UF校正.Visible = True Then
'  Select Case UF校正.LineName
'  Case "ﾊﾞﾗﾝｻｰ", "ﾄｯﾄ"
'    TimeCheck = "日勤"
'  Case Else
'    If DTime > CDate("7:00") And DTime < CDate("15:00") Then
'      TimeCheck = "1直"
'    ElseIf DTime >= CDate("15:00") And DTime < CDate("22:50") Then
'      TimeCheck = "2直"
'    ElseIf DTime >= CDate("22:50") And DTime <= CDate("23:59") Then
'      TimeCheck = "3直"
'    ElseIf DTime >= CDate("00:00") And DTime <= CDate("7:00") Then
'      TimeCheck = "3直"
'    Else
'       TimeCheck = "3直" '23:59のタイミングをひっかけられない
'    End If
'  End Select
'Exit Function
'End If
'
'If UF1.Tott = True Then
'  TimeCheck = "日勤"
'ElseIf UF1.Bala = True Then
'  TimeCheck = "日勤"
'Else
'  If DTime > CDate("7:00") And DTime < CDate("15:00") Then
'    TimeCheck = "1直"
'  ElseIf DTime >= CDate("15:00") And DTime < CDate("22:50") Then
'    TimeCheck = "2直"
'  ElseIf DTime >= CDate("22:50") And DTime <= CDate("23:59") Then
'    TimeCheck = "3直"
'  ElseIf DTime >= CDate("00:00") And DTime <= CDate("7:00") Then
'    TimeCheck = "3直"
'  Else
'     TimeCheck = "3直" '23:59のタイミングをひっかけられない
'  End If
'End If
End Function
Public Sub EndTime_Register() '2019.05.03 印刷ボタンにcall　印刷時が終了時間としている
'6:50を7:00に修正　2019.05.29
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim dt As String:       dt = DateAdd("d", -1, Now) '夜勤の日付またぎ前日扱い
Dim strDate As String:  strDate = Format$(dt, "yyyy/mm/dd hh:mm:ss")
Dim NowTim As Date: NowTim = Format$(Now, "yyyy/mm/dd hh:mm:ss")
If Time > CDate("7:00") And Time <= CDate("23:59") Then '0時前は押したときの日付
  UF1.EndTime = NowTim
ElseIf Time >= CDate("00:00") And Time <= CDate("7:00") Then '0時～7:00まで1日戻して処理する
  UF1.EndTime = Format$(strDate, "yyyy/mm/dd hh:mm:ss")
Else
  UF1.EndTime = NowTim
End If
If WsCheck("WorkerTime" & "_" & Lname) Then ThisWorkbook.Sheets("WorkerTime" & "_" & Lname).Cells(1, 3) = UF1.EndTime
End Sub
' 12時間以上経過しているかを判定_2024.9.3 経過していたらTrue
Public Function IsOver12Hours() As Boolean
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption

Dim TargetSheet As Worksheet
If WsCheck("WorkerTime" & "_" & Lname) Then Set TargetSheet = ThisWorkbook.Sheets("WorkerTime" & "_" & Lname) '2024.10.11_シートの有無チェックがなかった
If WsCheck("WorkerTime" & "_" & Lname) = False Then
  IsOver12Hours = False
  Exit Function
End If

If TargetSheet.Cells(1, 3) = "" Then
  IsOver12Hours = False
  Exit Function
End If
Dim TargetDate As Date:       TargetDate = CDate(TargetSheet.Cells(1, 3).Value)
Dim CurrentDate As Date:      CurrentDate = Now

If DateDiff("h", TargetDate, CurrentDate) >= 1 Then '12時間もいらなかった_2024.9.12_2024.12.12　3時間もいらんぽい
  IsOver12Hours = True
Else
  IsOver12Hours = False
End If
End Function

Public Sub Name_Match() '検査員一致判定
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("WorkerTime" & "_" & Lname) = False Then Exit Sub '-------------------------ｼｰﾄなし
Dim i As Long, a As Long, n As Long: n = 0
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("WorkerTime" & "_" & Lname) 'ｼｰﾄ
Dim WorkeRange As Range: Set WorkeRange = sh.Range(sh.Cells(3, 1), sh.Cells(7, 12)) '作業者範囲
Dim Worker(): Worker = WorkeRange '検査履歴格納
Dim NameDic As Object: Set NameDic = CreateObject("Scripting.Dictionary") '連想配列
Dim DoneDic As Object: Set DoneDic = CreateObject("Scripting.Dictionary") '連想配列
Dim OutFl As Boolean: OutFl = False
sh.Columns("N:Q").ClearContents
'------------------------------------------------------------------------------ｼｰﾄ上の測定者履歴を格納 重複を無くす
For i = LBound(Worker, 2) To UBound(Worker, 2) '列
  For a = LBound(Worker, 1) To UBound(Worker, 1) '行
    If Worker(a, i) <> "" Then '検査済みだけ
      If NameDic.exists(Worker(a, i)) = False Then '一致でない時格納
         n = n + 1 '番号
         NameDic(Worker(a, i)) = n
      End If
    End If
  Next a
Next i
'------------------------------------------------------------------------------貼り付け
Dim VarName As Variant, WorList() As Variant, j As Long: j = 1
ReDim WorList(1 To NameDic.Count + 1, 1 To 2) '測定者履歴ﾘｽﾄ
WorList(1, 2) = "作業済検査員履歴"
For Each VarName In NameDic '重複なしの作業者履歴
  j = j + 1
  WorList(j, 1) = NameDic(VarName)
  WorList(j, 2) = VarName
Next
Dim Temp As Range: Set Temp = sh.Range(sh.Cells(1, 14), sh.Cells(UBound(WorList, 1), 13 + UBound(WorList, 2))) '貼り付け範囲
Temp = WorList '貼り付け
'------------------ﾌｫｰﾑ登録者---------------複数人の場合一つの配列にまとめる
'2020.12.16_最終登録者チェックのみとする＊全ﾁｪｯｸでは前直までチェックしてしまい夜勤者でひっかかる
'Dim Tyoku As String: Tyoku = WorkeDayMax '最終登録直チェック
'Dim WorkerAll
'Select Case Tyoku
'Case "1": WorkerAll = Split(UF1.T1, "_")
'Case "2": WorkerAll = Split(UF1.T2, "_")
'Case "3": WorkerAll = Split(UF1.T3, "_")
'End Select
'Dim OneWorker: OneWorker = WorkerAll(UBound(WorkerAll))

Dim Worker1:  Worker1 = Split(UF1.T1, "_")
Dim Worker2:  Worker2 = Split(UF1.T2, "_")
Dim Worker3:  Worker3 = Split(UF1.T3, "_")
Dim OneWorker: OneWorker = Split(Join(Worker1) & " " & Join(Worker2) & " " & Join(Worker3)) '登録検査員ﾘｽﾄ:0配列
'------------------------------------------------------------------------------登録検査員を格納　重複を無くす
n = 0
For i = LBound(OneWorker) To UBound(OneWorker)
  If OneWorker(i) <> "" Then
    If DoneDic.exists(OneWorker(i)) = False Then
      n = n + 1  '番号
      DoneDic(OneWorker(i)) = n '空白は除く
    End If
  End If
Next i
'------------------------------------------------------------------------------貼り付け
Dim VarOneWork As Variant, OneWorkList() As Variant, R_Count As Long: R_Count = 1
ReDim OneWorkList(DoneDic.Count + 1) '1次元
OneWorkList(0) = "登録検査員"
For Each VarOneWork In DoneDic '重複なしの登録検査員
  OneWorkList(R_Count) = VarOneWork
  R_Count = R_Count + 1
Next
Dim OneWorktemp As Range: Set OneWorktemp = sh.Range(sh.Cells(1, 16), sh.Cells(UBound(OneWorkList, 1), 16)) '貼り付け範囲
OneWorktemp = Application.WorksheetFunction.Transpose(OneWorkList) '1次元入れ替え貼り付け
'Stop
'------------------検索------------------登録検査員を測定者履歴(NameDic)から探す
'※登録にミスがあった時、1,2,3直に登録があり3直のみで検査をした場合1,2直の登録者が履歴にない_2020.12.16_停止
Dim k As Long: k = 1
sh.Cells(1, 17) = "一致無し検査員"
'For i = 1 To UBound(OneWorkList) 'OneWorkList(0)はﾗﾍﾞﾙの為LBound不使用
'  If OneWorkList(i) <> "" Then '空白は除く
'    If NameDic.Exists(OneWorkList(i)) = False Then  '測定者履歴(検査履歴)に一致がない
'       k = k + 1
'       SH.Cells(k, 17) = OneWorkList(i)
'       SH.Cells(1, 18) = "OUT" '閉じたとき用ﾌﾗｸﾞ
'       OutFl = True
'    End If
'  End If
'Next i
'------------------検索------------------測定者履歴ﾘｽﾄ(重複無し作業者)を登録検査員から探す
For i = 2 To UBound(WorList, 1)
  If WorList(i, 2) <> "" Then '空白は除く
    If DoneDic.exists(WorList(i, 2)) = False Then '登録検査員に一致がない'停止処理の事
       k = k + 1
       sh.Cells(k, 17) = WorList(i, 2)
       sh.Cells(1, 18) = "OUT" '閉じたとき用ﾌﾗｸﾞ
       OutFl = True
    End If
  End If
Next i
'Stop

Application.DisplayAlerts = False
ThisWorkbook.Save '上書き保存
Application.DisplayAlerts = True

If OutFl Then Call NameList_Open(1) 'ﾌｫｰﾑ起動
Set NameDic = Nothing
Set DoneDic = Nothing
End Sub
Public Function Register_Check(STtime As String, Optional n As String = "") As Boolean '検査前,INP_Changeで使用
Dim Lname As String:      Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim DayReturn As String:  DayReturn = YeaDayrCheck(Format$(Now, "hh:mm"))
If n = "新規" Then GoTo Line1 Else GoTo Line2 '新規_Trueでシート発行処理
Line1:
  If WsCheck("検査員" & "_" & Lname) = False Then 'ｼｰﾄがない
    Register_Check = True
    Exit Function
  End If
  Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("検査員" & "_" & Lname)
  Dim i As Long
  For i = 1 To 3 '3行(Day1,Day2,Day3)に当日以外の登録がある
    If sh.Cells(i, 1) <> "" Then
     If Left(sh.Cells(i, 1), 10) <> DayReturn Then
      Register_Check = True
      Exit Function
     End If
    End If
  Next i
  Exit Function
Line2:
'------------------------------------------------------------検査員登録がない場合true
  If WsCheck("検査員" & "_" & Lname) = False Then 'ｼｰﾄがない
    Register_Check = True
    Exit Function
  End If
  
  Dim SetTime As String:   SetTime = TimeCheck(STtime)
  
  Select Case SetTime '該当直　'その時間に登録が無い
  Case "日勤": If (UF1.T1 = "") Or (UF1.Day1 = "") Then Register_Check = True
  Case "1直":  If (UF1.T1 = "") Or (UF1.Day1 = "") Then Register_Check = True
  Case "2直":  If (UF1.T2 = "") Or (UF1.Day2 = "") Then Register_Check = True
  Case "3直":  If (UF1.T3 = "") Or (UF1.Day3 = "") Then Register_Check = True
  End Select
  Select Case SetTime '登録はあるが当日登録でない
  Case "日勤": If Mid((UF1.Day1), 1, 10) <> DayReturn Then Register_Check = True
  Case "1直":  If Mid((UF1.Day1), 1, 10) <> DayReturn Then Register_Check = True
  Case "2直":  If Mid((UF1.Day2), 1, 10) <> DayReturn Then Register_Check = True
  Case "3直":  If Mid((UF1.Day3), 1, 10) <> DayReturn Then Register_Check = True
  End Select
End Function
Public Function Epual(Wor As Variant) As Boolean  '全同作業者判定*2021.04.15 trueで全一致
Dim i As Long
For i = 1 To UBound(Wor) 'Split関数_添え字が「0」
  Select Case True
  Case Wor(i) <> Wor(i - 1)
    Epual = False
    Exit Function
  End Select
Next i
Epual = True
End Function
Public Function ReWorker(Wor As String, Optional n As String = "_") As Variant '2021.04.15
ReWorker = Split(Wor, n)
If IsEmptyArray(ReWorker) Then ReWorker = "" 'trueで空
End Function
Public Sub WorkerSave() '検査員保存
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("検査員" & "_" & Lname) = False Then
  Call AddWS("検査員" & "_" & Lname)
End If
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("検査員" & "_" & Lname)
Dim i As Long, a As Long
Dim TY1R As Range, TY2R As Range, TY3R As Range
Dim T_1 As String, T_2 As String, T_3 As String
Dim ReT1, ReT2, ReT3
With UF1
  Dim 直1(): 直1 = Array(.T1.Text)
  Dim 直2(): 直2 = Array(.T2.Text)
  Dim 直3(): 直3 = Array(.T3.Text)
  T_1 = 直1(0)
  T_2 = 直2(0)
  T_3 = 直3(0)
  '2021.04.15_全て同じ作業者の場合1人にして貼り付け
  If T_1 <> "" Then
    ReT1 = ReWorker(T_1) 'Split使用関数で配列を返す
    If Epual(ReT1) Then T_1 = ReT1(UBound(ReT1)) 'trueで全一致_全一致の時だけ文字列書き換え
  ElseIf T_1 = "" Then
    T_1 = ""
  End If
  If T_2 <> "" Then
    ReT2 = ReWorker(T_2)
    If Epual(ReT2) Then T_2 = ReT2(UBound(ReT2))
  ElseIf T_2 = "" Then
    T_2 = ""
  End If
  If T_3 <> "" Then
    ReT3 = ReWorker(T_3)
    If Epual(ReT3) Then T_3 = ReT3(UBound(ReT3))
  ElseIf T_3 = "" Then
    T_3 = ""
  End If
  直1 = Array(.Day1.Text, T_1)
  直2 = Array(.Day2.Text, T_2)
  直3 = Array(.Day3.Text, T_3)
End With
Set TY1R = sh.Range(sh.Cells(1, 1), sh.Cells(1, UBound(直1) + 1))
Set TY2R = sh.Range(sh.Cells(2, 1), sh.Cells(2, UBound(直2) + 1))
Set TY3R = sh.Range(sh.Cells(3, 1), sh.Cells(3, UBound(直3) + 1))
TY1R = 直1
TY2R = 直2
TY3R = 直3
With sh.Cells
  .Font.Name = "メイリオ"
  .Font.Size = 12
  .ColumnWidth = 10.25
  .EntireColumn.AutoFit
End With
End Sub
Public Sub WorkerLoad() '検査員読み込み
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("検査員" & "_" & Lname) = False Then Exit Sub '-------------------------ｼｰﾄなし
Dim Worker(), WorkRan As Range
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("検査員" & "_" & Lname)
Set WorkRan = sh.Range(sh.Cells(1, 1), sh.Cells(3, 2))
Worker() = WorkRan
UF1.FlagLabel = 0 'ﾁｪﾝｼﾞが1から動いていくため0で止める
With UF1
  .Day1 = Format(Worker(1, 1), "yyyy.mm.dd  hh:mm")
  .T1 = Worker(1, 2)
  .Day2 = Format(Worker(2, 1), "yyyy.mm.dd  hh:mm")
  .T2 = Worker(2, 2)
  .Day3 = Format(Worker(3, 1), "yyyy.mm.dd  hh:mm")
  .T3 = Worker(3, 2)
End With
UF1.FlagLabel = 1 '出そろってからWorkeDayMax処理
End Sub

Public Function WorkeDayMax() As String '登録日の遅いもの
On Error Resume Next
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("検査員" & "_" & Lname) = False Then Exit Function '-------------------------ｼｰﾄなし
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("検査員" & "_" & Lname)
'UFではﾁｪﾝｼﾞｲﾍﾞﾝﾄが1から順に起こって2,3の時間が拾えないのでｼｰﾄから拾う '最終日付ﾁｪｯｸ
Dim Day1str As String, Day2str As String, Day3str As String
'日付抜取
Day1str = Mid(sh.Cells(1, 1), 1, 10)
Day2str = Mid(sh.Cells(2, 1), 1, 10)
Day3str = Mid(sh.Cells(3, 1), 1, 10)
Dim flag(1 To 7) As Boolean
Dim T1 As Long, T2 As Long, T3 As Long
Dim Day1 As Date, Day2 As Date, Day3 As Date '変換用
'ﾌﾗｸﾞ作成_2021.02.11
If Day1str <> "" And Day2str = "" And Day3str = "" Then flag(1) = True '1のみ
If Day1str = "" And Day2str <> "" And Day3str = "" Then flag(2) = True '2のみ
If Day1str = "" And Day2str = "" And Day3str <> "" Then flag(3) = True '3のみ
If Day1str <> "" And Day2str <> "" And Day3str = "" Then flag(4) = True '1,2
If Day1str <> "" And Day2str = "" And Day3str <> "" Then flag(5) = True '1,3
If Day1str = "" And Day2str <> "" And Day3str <> "" Then flag(6) = True '2,3
If Day1str <> "" And Day2str <> "" And Day3str <> "" Then flag(7) = True '1,2,3
'時間抜取
If sh.Cells(1, 1) <> "" Then T1 = Mid(sh.Cells(1, 1), 13, 2) 'スペース2つ
If sh.Cells(2, 1) <> "" Then T2 = Mid(sh.Cells(2, 1), 13, 2)
If sh.Cells(3, 1) <> "" Then T3 = Mid(sh.Cells(3, 1), 13, 2)
'日付型変換
If Day1str <> "" Then Day1 = Replace(Day1str, ".", "/")
If Day2str <> "" Then Day2 = Replace(Day2str, ".", "/")
If Day3str <> "" Then Day3 = Replace(Day3str, ".", "/")

Select Case True
Case flag(1)
  WorkeDayMax = "1"
Case flag(2)
  WorkeDayMax = "2"
Case flag(3)
  WorkeDayMax = "3"
Case flag(4) '1,2
  If Day1str = Day2str Then  '同日_時間で比較
    Select Case Application.WorksheetFunction.Max(T1, T2) '登録時間が一番遅い
    Case T1: WorkeDayMax = "1"
    Case T2: WorkeDayMax = "2"
    End Select
  Else
    Select Case Application.WorksheetFunction.Max(Day1, Day2)
    Case Day1: WorkeDayMax = "1"
    Case Day2: WorkeDayMax = "2"
    End Select
  End If
Case flag(5) '1,3
  If Day1str = Day3str Then
    Select Case Application.WorksheetFunction.Max(T1, T3)
    Case T1: WorkeDayMax = "1"
    Case T3: WorkeDayMax = "3"
    End Select
  Else
    Select Case Application.WorksheetFunction.Max(Day1, Day3)
    Case Day1: WorkeDayMax = "1"
    Case Day3: WorkeDayMax = "3"
    End Select
  End If
Case flag(6) '2,3
  If Day2str = Day3str Then
    Select Case Application.WorksheetFunction.Max(T2, T3)
    Case T2: WorkeDayMax = "2"
    Case T3: WorkeDayMax = "3"
    End Select
  Else
    Select Case Application.WorksheetFunction.Max(Day2, Day3)
    Case Day2: WorkeDayMax = "2"
    Case Day3: WorkeDayMax = "3"
    End Select
  End If
Case flag(7) '1,2,3
  If Day1str = Day2str And Day1str = Day3str Then '1,2,3_全て同日は時間比較のみ
    Select Case Application.WorksheetFunction.Max(T1, T2, T3)
    Case T1: WorkeDayMax = "1"
    Case T2: WorkeDayMax = "2"
    Case T3: WorkeDayMax = "3"
    End Select
  ElseIf Day1str = Day2str Then '1,2_2つが同日は日付比較後に時間比較
    Select Case Application.WorksheetFunction.Max(Day1, Day2, Day3)
    Case Day3
      WorkeDayMax = "3"
    Case Else
      Select Case Application.WorksheetFunction.Max(T1, T2)
      Case T1: WorkeDayMax = "1"
      Case T2: WorkeDayMax = "2"
      End Select
    End Select
  ElseIf Day1str = Day3str Then '1,3
    Select Case Application.WorksheetFunction.Max(Day1, Day2, Day3)
    Case Day2
      WorkeDayMax = "2"
    Case Else
      Select Case Application.WorksheetFunction.Max(T1, T3)
      Case T1: WorkeDayMax = "1"
      Case T3: WorkeDayMax = "3"
      End Select
    End Select
  ElseIf Day2str = Day3str Then '2,3
    Select Case Application.WorksheetFunction.Max(Day1, Day2, Day3)
    Case Day1
      WorkeDayMax = "1"
    Case Else
      Select Case Application.WorksheetFunction.Max(T2, T3)
      Case T2: WorkeDayMax = "2"
      Case T3: WorkeDayMax = "3"
      End Select
    End Select
  Else 'すべて違う日付_2021.02.15修正
    Select Case Application.WorksheetFunction.Max(Day1, Day2, Day3)
    Case Day1: WorkeDayMax = "1"
    Case Day2: WorkeDayMax = "2"
    Case Day3: WorkeDayMax = "3"
    End Select
  End If
End Select
End Function

'WorkeDayMax に応じて T1/T2/T3 から最後の登録者名を取得する
Public Function GetWorkerName() As String
    Dim Worker As Variant
    Select Case WorkeDayMax
    Case "1": Worker = Split(UF1.T1, "_")
    Case "2": Worker = Split(UF1.T2, "_")
    Case "3": Worker = Split(UF1.T3, "_")
    Case Else: Exit Function
    End Select
    Dim w As String: w = Worker(UBound(Worker))
    If w <> "" Then GetWorkerName = w
End Function

Public Sub Daily_Create() '日報ｼｰﾄ作成　直毎 UF1検査員登録 ｷｬｽｸ三菱は新規発行のﾀｲﾐﾝｸﾞ
Dim Lname As String
Select Case True
Case UF1.Visible:          Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Case UFCask.Visible:       Lname = UFCask.LineName.Caption
Case UFMitsubishi.Visible: Lname = UFMitsubishi.LineName.Caption
End Select

Dim DayReturn As String:  DayReturn = YeaDayrCheck(Format$(Now, "hh:mm"))
Dim SetTime As String:    SetTime = TimeCheck(Format$(Now, "hh:mm:ss"))

If WsCheck("Daily" & "_" & Lname) = False Then Call AddWS("Daily" & "_" & Lname)
If ThisWorkbook.Sheets("Daily" & "_" & Lname).Cells(1, 1) <> DayReturn & "_" & SetTime Then Call AddWS("Daily" & "_" & Lname)
'If ThisWorkbook.Sheets("Daily" & "_" & Lname).Cells(1, 1) <> YeaDayrCheck(Format$(Now, "hh:mm")) & "_" & TimeCheck(Format$(Now, "hh:mm:ss")) Then Call AddWS("Daily" & "_" & Lname)
'ﾌｨｰﾙﾄﾞ登録
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Daily" & "_" & Lname)
Dim LField As Variant, FieldRan As Range
LField = Array("ﾛｯﾄNo", "材・調質", "製品寸法", "梱包No", "検入枚数", "開始時", "開始分", "終了時", "終了分", "作業人数", "合紙", "枚数", "包数", "VC種別", _
"etc", " 停止①記号", " 停止①時間", " 停止②記号", " 停止②時間", " 停止③記号", " 停止③時間", " 実績枚数", " 実績重量", " 作業時間", " 単重")
Set FieldRan = sh.Range(sh.Cells(2, 1), sh.Cells(2, UBound(LField) + 1)) '0配列
FieldRan = LField
FieldRan.Interior.Color = RGB(255, 255, 204)
Call Daily_Report(1) '日報情報_検査員,ライン名
Call Daily_Report(3) '日報情報(検査開始)
End Sub

Public Sub Daily_Report(n As Long) '日報データ挿入
Dim Lname As String
If ThisWorkbook.Sheets("全ヵ所検査").Visible = False And ThisWorkbook.Sheets("キャスク検査").Visible = False And ThisWorkbook.Sheets("H494検査").Visible = False Then
  Lname = UF1.LineName.Caption & UF1.LineNo.Caption
End If
If WsCheck("Daily" & "_" & Lname) = False Then Exit Sub
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Daily" & "_" & Lname)
Dim LR As Long: LR = UR(sh, 4) + 1
If WsCheck("Lot情報履歴") = False Then Call AddWS("Lot情報履歴")
Dim HistorySH As Worksheet: Set HistorySH = ThisWorkbook.Sheets("Lot情報履歴")
Dim HisLR As Long: HisLR = UR(HistorySH, 1)
Dim HisLC As Long: HisLC = UC(HistorySH, 1)
Dim HistoryR As Range: Set HistoryR = Range(HistorySH.Cells(2, 1), HistorySH.Cells(HisLR, HisLC))
Dim TempHiki(): TempHiki = HistoryR
Dim i As Long, x As Long, matchFlag As Boolean: matchFlag = False
Dim WorkingLot As String

Select Case True
Case ThisWorkbook.Sheets("全ヵ所検査").Visible:    WorkingLot = ThisWorkbook.Sheets("全ヵ所検査").Cells(2, 4)
Case ThisWorkbook.Sheets("キャスク検査").Visible:  WorkingLot = ThisWorkbook.Sheets("キャスク検査").Cells(2, 4)
Case ThisWorkbook.Sheets("H494検査").Visible:      WorkingLot = ThisWorkbook.Sheets("H494検査").Cells(2, 4)
Case Else:                                         WorkingLot = UF1.Lot
End Select
For i = 1 To UBound(TempHiki, 1)
  If TempHiki(i, 3) = WorkingLot Then 'LOTNO一致
    x = i
    matchFlag = True
    Exit For
  End If
Next i
If matchFlag = False Then Exit Sub '---------------------------一致なし
Select Case n
Case 1 '日付,ﾗｲﾝ,検査員　Worker_Changeのタイミング(検査員読み込み後)　ｷｬｽｸ三菱は検査1枚目のﾀｲﾐﾝｸﾞ
  sh.Cells(1, 1) = YeaDayrCheck(Format$(Now, "hh:mm")) & "_" & TimeCheck(Format$(Now, "hh:mm:ss"))
  Select Case True
  Case ThisWorkbook.Sheets("全ヵ所検査").Visible
    Select Case UFMitsubishi.LineName.Caption
    Case "L-1":    sh.Cells(1, 2) = "L-1"
    Case "LVC":    sh.Cells(1, 2) = "LVC"
    Case "HVC":    sh.Cells(1, 2) = "HVC"
    Case "ﾊﾞﾗﾝｻｰ": sh.Cells(1, 2) = "ﾊﾞﾗﾝｻｰ"
    Case "ﾄｯﾄ":    sh.Cells(1, 2) = "ﾄｯﾄ"
    Case Else:     sh.Cells(1, 2) = "Test"
    End Select
    sh.Cells(1, 3) = UFMitsubishi.Kname
  Case ThisWorkbook.Sheets("H494検査").Visible
    Select Case UFMitsubishi.LineName.Caption
    Case "L-1":    sh.Cells(1, 2) = "L-1"
    Case "LVC":    sh.Cells(1, 2) = "LVC"
    Case "HVC":    sh.Cells(1, 2) = "HVC"
    Case "ﾊﾞﾗﾝｻｰ": sh.Cells(1, 2) = "ﾊﾞﾗﾝｻｰ"
    Case "ﾄｯﾄ":    sh.Cells(1, 2) = "ﾄｯﾄ"
    Case Else:     sh.Cells(1, 2) = "Test"
    End Select
    sh.Cells(1, 3) = UFMitsubishi.Kname
  Case ThisWorkbook.Sheets("キャスク検査").Visible
    Select Case UFCask.LineName.Caption
    Case "L-1":    sh.Cells(1, 2) = "L-1"
    Case "ﾊﾞﾗﾝｻｰ": sh.Cells(1, 2) = "ﾊﾞﾗﾝｻｰ"
    Case "ﾄｯﾄ":    sh.Cells(1, 2) = "ﾄｯﾄ"
    End Select
    sh.Cells(1, 3) = UFCask.Kname
  Case Else
    Select Case True
    Case UF1.L1Work: sh.Cells(1, 2) = "L-1"
    Case UF1.LVC:    sh.Cells(1, 2) = "LVC"
    Case UF1.HVC:    sh.Cells(1, 2) = "HVC"
    Case UF1.Bala:   sh.Cells(1, 2) = "ﾊﾞﾗﾝｻｰ"
    Case UF1.Tott:   sh.Cells(1, 2) = "ﾄｯﾄ"
    Case Else:       sh.Cells(1, 2) = "Test"
    End Select
    Select Case TimeCheck(Format$(Now, "hh:mm"))
    Case "日勤": sh.Cells(1, 3) = UF1.T1
    Case "1直":  sh.Cells(1, 3) = UF1.T1
    Case "2直":  sh.Cells(1, 3) = UF1.T2
    Case "3直":  sh.Cells(1, 3) = UF1.T3
    End Select
  End Select
Case 2 'Lot情報　TimeRecordのタイミング   ｷｬｽｸ三菱は検査判定
  sh.Cells(LR, 1) = TempHiki(x, 3) 'LotNo
  sh.Cells(LR, 2) = TempHiki(x, 8) & "-" & TempHiki(x, 9) '材調質
  sh.Cells(LR, 3) = Format(TempHiki(x, 4), "0.000") & "×" & Format(TempHiki(x, 5), "0.0") & "×" & Format(TempHiki(x, 6), "0.0") 'ｻｲｽﾞ
  sh.Cells(LR, 5) = TempHiki(x, 15) '検入枚数
  Select Case TempHiki(x, 17) '合紙
  Case 0: sh.Cells(LR, 11) = "無"
  Case 1: sh.Cells(LR, 11) = "有"
  End Select
  sh.Cells(LR, 25) = Format(TempHiki(x, 21), "0.00")
  If TempHiki(x, 18) = 1 Then
    Select Case True
    Case TempHiki(x, 19) <> "" And TempHiki(x, 20) <> "": sh.Cells(LR, 14) = "A" & ":" & TempHiki(x, 19) & "_" & "B" & ":" & TempHiki(x, 20) 'AB
    Case TempHiki(x, 19) <> "" And TempHiki(x, 20) = "":  sh.Cells(LR, 14) = "A" & ":" & TempHiki(x, 19)  'A
    Case TempHiki(x, 19) = "" And TempHiki(x, 20) <> "":  sh.Cells(LR, 14) = "B" & ":" & TempHiki(x, 20)  'B
    Case TempHiki(x, 19) = "" And TempHiki(x, 20) = "":   sh.Cells(LR, 14) = "" 'なし
    End Select
  End If
  Select Case True
  Case ThisWorkbook.Sheets("キャスク検査").Visible: sh.Cells(LR, 15) = "キャスク"
  End Select
Case 3 '開始時間情報 Daily_Create＊一番初めだけで良い　以降の開始はCase 4で行う
  Call Daily_Report(1)
  If sh.Cells(LR, 1) = "" Then Call Daily_Report(2) '検査員登録後検査が一度もないままAｲﾝﾌﾟｯﾄした場合
  Select Case Len(Hour(Now)) '時
  Case 1: sh.Cells(LR, 6) = "'" & "0" & Hour(Now)
  Case 2: sh.Cells(LR, 6) = "'" & Hour(Now)
  End Select
  Select Case Len(Minute(Now)) '分
  Case 1: sh.Cells(LR, 6) = "'" & "0" & Minute(Now)
  Case 2: sh.Cells(LR, 7) = "'" & Minute(Now)
  End Select
Case 4 '終了時間情報 minmaxのタイミング　A_Output前
  '終了時間
  If sh.Cells(LR, 1) = "" Then Call Daily_Report(2) '検査員登録後検査が一度もないままAｲﾝﾌﾟｯﾄした場合_2021.07.27追加*検入枚数未登録のまま続行できてしまうため追加
  sh.Cells(LR + 1, 1) = UFAinput.Lot
  sh.Cells(LR, 4) = "'" & UFAinput.KONNO '梱包No
  Select Case True
  Case (UFAinput.COU2) <> ""  'Box入力の判定
    Select Case Len(Hour(Now)) '時
    Case 1: sh.Cells(LR + 1, 8) = "'" & "0" & Hour(Now)
    Case 2: sh.Cells(LR + 1, 8) = "'" & Hour(Now)
    End Select
    Select Case Len(Minute(Now)) '分
    Case 1: sh.Cells(LR + 1, 9) = "'" & "0" & Minute(Now)
    Case 2: sh.Cells(LR + 1, 9) = "'" & Minute(Now)
    End Select
    '***********************2021.11.05_梱包入力Boxを2段使用した場合の記述を追加(分けて入力する)
    '梱包数
    sh.Cells(LR, 12) = val(UFAinput.MAI1) '枚数
    sh.Cells(LR, 13) = val(UFAinput.COU1) '包み
    
    sh.Cells(LR + 2, 1) = UFAinput.Lot
    sh.Cells(LR + 1, 4) = "'" & UFAinput.KONNO '梱包No_梱包が2つあるので次行にも同じ梱包Noをつける
    sh.Cells(LR + 1, 12) = val(UFAinput.MAI2) '枚数
    sh.Cells(LR + 1, 13) = val(UFAinput.COU2) '包み
    '次梱包の開始
    Select Case Len(Hour(Now)) '時
    Case 1: sh.Cells(LR + 2, 6) = "'" & "0" & Hour(Now)
    Case 2: sh.Cells(LR + 2, 6) = "'" & Hour(Now)
    End Select
    Select Case Len(Minute(Now)) '分
    Case 1: sh.Cells(LR + 2, 7) = "'" & "0" & Minute(Now)
    Case 2: sh.Cells(LR + 2, 7) = "'" & Minute(Now)
    End Select
  Case Else
    sh.Cells(LR + 1, 1) = UFAinput.Lot
    sh.Cells(LR, 4) = "'" & UFAinput.KONNO '梱包No
    Select Case Len(Hour(Now)) '時
    Case 1: sh.Cells(LR, 8) = "'" & "0" & Hour(Now)
    Case 2: sh.Cells(LR, 8) = "'" & Hour(Now)
    End Select
    Select Case Len(Minute(Now)) '分
    Case 1: sh.Cells(LR, 9) = "'" & "0" & Minute(Now)
    Case 2: sh.Cells(LR, 9) = "'" & Minute(Now)
    End Select
    
    sh.Cells(LR, 12) = val(UFAinput.MAI1) '枚数
    sh.Cells(LR, 13) = val(UFAinput.COU1) '包み
    '次行の開始
    Select Case Len(Hour(Now)) '時
    Case 1: sh.Cells(LR + 1, 6) = "'" & "0" & Hour(Now)
    Case 2: sh.Cells(LR + 1, 6) = "'" & Hour(Now)
    End Select
    Select Case Len(Minute(Now)) '分
    Case 1: sh.Cells(LR + 1, 7) = "'" & "0" & Minute(Now)
    Case 2: sh.Cells(LR + 1, 7) = "'" & Minute(Now)
    End Select
  End Select
End Select
With sh.Cells
  .Font.Name = "メイリオ"
  .Font.Size = 10
  .EntireColumn.AutoFit
End With
End Sub
Public Sub Daily_Report4_Execution() 'Daily_Report4専用
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("Daily" & "_" & Lname) = False Then Exit Sub
UFAinput.Ngflag = ""
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Daily" & "_" & Lname)
Dim LR As Long: LR = UR(sh, 1)
Dim EndR As Long: EndR = UR(sh, 4)
Dim K_No As Long: K_No = val(sh.Cells(EndR, 4)) + 1 '最終行の次の番号

 If UFAinput.KONNO <> "" Then '梱包Noの入力があるか
  If sh.Cells(EndR, 4) = "梱包No" Then
   If sh.Cells(LR, 1) = UFAinput.Lot Then Call Daily_Report(4) 'LotNoの一致が条件
   Exit Sub
  End If
   If K_No = val(UFAinput.KONNO) Then '梱包Noが連番になっているか
    If sh.Cells(LR, 1) = UFAinput.Lot Then Call Daily_Report(4) 'LotNoの一致が条件
    Exit Sub
   Else
    If (sh.Cells(LR, 4) = "") And (sh.Cells(LR - 1, 1) <> UFAinput.Lot) Then '別ロット作業の初回_2021.03.05追加
     Call Daily_Report(4)
    Else
     MsgBox "連番でない" & "やり直してください", vbCritical
     If UFProgress.Visible Then Unload UFProgress
     UFAinput.Ngflag = "stop"
     Exit Sub
    End If
   End If
 End If
End Sub
Public Sub Number_Count(Lotname As String) '梱包枚数計算
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("Daily" & "_" & Lname) = False Then Exit Sub
Dim sh As Worksheet:   Set sh = ThisWorkbook.Sheets("Daily" & "_" & Lname)
Call AddWS("梱包枚数計算" & "_" & Lname)
Dim ws As Worksheet:   Set ws = ThisWorkbook.Sheets("梱包枚数計算" & "_" & Lname)
sh.Range(sh.Cells(3, 5), sh.Cells(UR(sh, 5), 5)).Interior.Pattern = xlNone  '色ｸﾘｱ
Dim Dran As Range:     Set Dran = sh.Range(sh.Cells(3, 1), sh.Cells(UR(sh, 1), 13))
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
  UF1.PrintB.Tag = "stop" '停止フラグ　一致しない場合は印刷停止　2020.11.24
  Call NameList_Open(3)
End If
End Sub

