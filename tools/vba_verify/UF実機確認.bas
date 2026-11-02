Attribute VB_Name = "UF実機確認"
'=====================================================================
' UF検査判定 Python 移植の「実機確認」(docs/04 の U1～U12)
'
' 使い方:
'   1. UF検査判定.xlsm(の写し)を開き、VBE で このファイルをインポート
'   2. 下の DB_FOLDER を、品質基準.accdb・伝送用ファイル.accdb・SIKALOTNOW.accdb がある
'      フォルダに書き換える(QualityPath / MainPath のどちらかに揃っていなくてもよい。
'      無いファイルの項目は「ファイルなし」と出るだけ)
'   3. マクロ「実機確認_全部」を実行
'   4. 新しいシート「実機確認」ができるので、そのシートを CSV で保存して送ってください
'
' 何も書き換えません(DB は読むだけ・ブックはシートを1枚足すだけ)。
'=====================================================================
Option Explicit

Private Const DB_FOLDER As String = "C:\UF実機確認\"
Private Const DB_PASSWORD As String = "nlm"

Private ws As Worksheet
Private r As Long

Public Sub 実機確認_全部()
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    On Error Resume Next
    ws.Name = "実機確認"
    On Error GoTo 0
    ws.Range("A1:E1").Value = Array("番号", "確かめること", "式", "結果", "エラー番号")
    r = 1
    Call U1_比較
    Call U5_丸め
    Call U12_日付
    Call 文字列関数
    Call U3_U4_U9_DB
    ws.Columns("A:E").AutoFit
    MsgBox "終わりました。シート「実機確認」を CSV で保存して送ってください。", vbInformation
End Sub

Private Sub Rec(no As String, what As String, expr As String, result As Variant, Optional errNo As Long = 0)
    r = r + 1
    ws.Cells(r, 1).Value = no
    ws.Cells(r, 2).Value = what
    ws.Cells(r, 3).Value = "'" & expr
    ws.Cells(r, 4).Value = "'" & CStr(result)
    ws.Cells(r, 5).Value = errNo
End Sub

'--- U1: 数にならない文字列と数の比較(OutTolerance O:871 / O:1989、Judg J:212) -------------
Private Sub U1_比較()
    Dim v As Variant, s As String, b As Boolean
    ' テキストボックスの値と同じ形(Variant に入った文字列)
    v = ""
    On Error Resume Next
    Err.Clear: b = (v = 0): Rec "U1-1", "Variant の空文字 = 0", "v="""": v = 0", b, Err.Number
    Err.Clear: b = (v > 0): Rec "U1-2", "Variant の空文字 > 0", "v="""": v > 0", b, Err.Number
    v = "S"
    Err.Clear: b = (v = 2): Rec "U1-3", "Variant の ""S"" = 2", "v=""S"": v = 2", b, Err.Number
    s = ""
    Err.Clear: b = (s = 0): Rec "U1-4", "String 変数の空文字 = 0", "s="""": s = 0", b, Err.Number
    v = ""
    Err.Clear
    Dim hit As String: hit = "どれでもない"
    Select Case v
        Case 2, 4: hit = "Case 2,4"
        Case Else: hit = "Case Else"
    End Select
    Rec "U1-5", "Select Case に空文字(Case 2, 4)", "Select Case """"", hit, Err.Number
    v = "S"
    Err.Clear: hit = "どれでもない"
    Select Case v
        Case 2, 4: hit = "Case 2,4"
        Case Else: hit = "Case Else"
    End Select
    Rec "U1-6", "Select Case に ""S""(Case 2, 4)", "Select Case ""S""", hit, Err.Number
    Dim vn As Variant: vn = 0
    v = ""
    Err.Clear: b = (v > vn): Rec "U1-10", "Variant の空文字 > Variant の 0", "v="""": vn=0: v > vn", b, Err.Number
    s = ""
    Err.Clear: b = (s = vn): Rec "U1-11", "String の空文字 = Variant の 0", "s="""": s = vn", b, Err.Number
    s = "10"
    vn = 9
    Err.Clear: b = (s < vn): Rec "U1-12", "String の ""10"" < Variant の 9(文字で比べるなら True)", "s=""10"": s < vn", b, Err.Number
    Err.Clear: b = (CDbl("1.5") > ""): Rec "U1-7", "CDbl(""1.5"") > """"", "CDbl(TB) > """"", b, Err.Number
    ' On Error Resume Next の中で If の条件がエラーになったとき、Then の中に入るか
    Err.Clear: hit = "Then に入らない"
    v = ""
    If v > 0 Then hit = "Then に入った"
    Rec "U1-8", "Resume Next 中、条件がエラーの If", "If """" > 0 Then", hit, Err.Number
    Dim f As Variant
    Err.Clear: f = Empty: Dim PackF(): PackF = f: Rec "U1-9", "Dim PackF(): PackF = Empty", "PackF = Empty", "代入した", Err.Number
    On Error GoTo 0
End Sub

'--- U5: Format / Round の .5 の境界 -------------------------------------------------------
Private Sub U5_丸め()
    Dim x As Variant
    For Each x In Array(2.675, 0.125, 1.005, 2.5, 3.5, 0.045, 1.0049999999, 999.95)
        Rec "U5-F", "Format(x, ""0.00"")", "x=" & CStr(x), Format(x, "0.00")
        Rec "U5-F1", "Format(x, ""0.0"")", "x=" & CStr(x), Format(x, "0.0")
        Rec "U5-R", "Round(x, 2)", "x=" & CStr(x), Round(x, 2)
        Rec "U5-R0", "Round(x)", "x=" & CStr(x), Round(x)
        Rec "U5-L", "CLng(x)", "x=" & CStr(x), CLng(x)
    Next
    Dim sg As Single
    For Each x In Array(1.01, 2.51, 0.05, 2053.25, 1.005)
        sg = CSng(x)
        Rec "U5-S", "Single: CStr / Format 0.00 / Val", "x=" & CStr(x), CStr(sg) & " / " & Format(sg, "0.00") & " / " & Val(sg)
    Next
    Rec "U5-D", "CStr(1.005 * 1000)", "1.005*1000", CStr(1.005 * 1000)
    Rec "U5-W", "WorksheetFunction.RoundDown(12.3456, 2)", "RoundDown", Application.WorksheetFunction.RoundDown(12.3456, 2)
    Rec "U5-N", "Format(-0.04, ""0.0"")", "-0.04", Format(-0.04, "0.0")
End Sub

'--- U12: IsDate(印刷履歴の日付の書き方) ---------------------------------------------------
Private Sub U12_日付()
    Rec "U12-1", "IsDate(""2026.06.22"")", "IsDate", IsDate("2026.06.22")
    Rec "U12-2", "IsDate(""2026.10.13"")(校正記録)", "IsDate", IsDate("2026.10.13")
    On Error Resume Next
    Err.Clear: Rec "U12-3", "CDate(""2026.06.22"")", "CDate", CDate("2026.06.22"), Err.Number
    On Error GoTo 0
End Sub

'--- 文字列関数 ---------------------------------------------------------------------------
Private Sub 文字列関数()
    Rec "S-1", "Trim が全角空白も消すか", "Trim(""　A　"")", "[" & Trim("　A　") & "]"
    On Error Resume Next
    Dim t As String
    Err.Clear: t = Left("", -1): Rec "S-2", "Left("""", -1)", "Left", t, Err.Number
    On Error GoTo 0
    Rec "S-3", "StrComp(""ﾄｯﾄ"",""トット"",vbTextCompare)", "StrComp", StrComp("ﾄｯﾄ", "トット", vbTextCompare)
    Rec "S-4", "InStr(1,""ABC"",""b"",vbTextCompare)", "InStr", InStr(1, "ABC", "b", vbTextCompare)
    Rec "S-5", """ﾄｯﾄ"" = ""トット""", "=", ("ﾄｯﾄ" = "トット")
End Sub

'--- U3 / U4 / U9: Access(ADO Filter・列名の全角半角・返却順) ------------------------------
Private Function OpenDb(name As String) As Object
    Dim p As String: p = DB_FOLDER & name
    If Dir(p) = "" Then Exit Function
    Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    On Error Resume Next
    cn.Properties("Jet OLEDB:Database Password") = DB_PASSWORD
    cn.Open p
    If Err.Number <> 0 Then
        Err.Clear
        Set cn = CreateObject("ADODB.Connection")
        cn.Provider = "Microsoft.ACE.OLEDB.12.0"
        cn.Open p
    End If
    On Error GoTo 0
    Set OpenDb = cn
End Function

Private Function FilterCount(cn As Object, table As String, flt As String) As String
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    On Error Resume Next
    rs.CursorLocation = 3
    rs.Open "SELECT * FROM " & table, cn, 3, 1
    rs.Filter = flt
    If Err.Number <> 0 Then
        FilterCount = "エラー " & Err.Number & " " & Err.Description
    Else
        FilterCount = CStr(rs.RecordCount)
    End If
    rs.Close
    On Error GoTo 0
End Function

Private Sub U3_U4_U9_DB()
    Dim cn As Object
    Set cn = OpenDb("伝送用ファイル.accdb")
    If cn Is Nothing Then
        Rec "U3", "伝送用ファイル.accdb", DB_FOLDER, "ファイルなし"
    Else
        Rec "U3-1", "Filter [ライン]='ﾄｯﾄ'(表はトットと全角)", "測定機器ID", FilterCount(cn, "測定機器ID", "[ライン]='ﾄｯﾄ'")
        Rec "U3-2", "Filter [ライン]='トット'", "測定機器ID", FilterCount(cn, "測定機器ID", "[ライン]='トット'")
        Rec "U3-3", "Filter [ライン]='ﾊﾞﾗﾝｻｰ'", "測定機器ID", FilterCount(cn, "測定機器ID", "[ライン]='ﾊﾞﾗﾝｻｰ'")
        Rec "U3-4", "Filter [ライン]='l-1'(小文字)", "測定機器ID", FilterCount(cn, "測定機器ID", "[ライン]='l-1'")
        Rec "U3-5", "Filter [ライン]='L-1 '(末尾空白)", "測定機器ID", FilterCount(cn, "測定機器ID", "[ライン]='L-1 '")
        Rec "U3-6", "Filter [用途]=''(空欄の行)", "測定機器ID", FilterCount(cn, "測定機器ID", "[用途]=''")
        Rec "U3-7", "Filter [ライン]='ＬＶＣ'(全角英字)", "測定機器ID", FilterCount(cn, "測定機器ID", "[ライン]='ＬＶＣ'")
        Rec "U9-1", "時間用の返却順(直)", "SELECT *", FirstColumnValues(cn, "時間用", "直")
        cn.Close
    End If
    Set cn = OpenDb("品質基準.accdb")
    If cn Is Nothing Then
        Rec "U9", "品質基準.accdb", DB_FOLDER, "ファイルなし"
    Else
        Rec "U9-2", "製品規格ラテラルボーの返却順(管理番号の先頭20件)", "SELECT *", FirstColumnValues(cn, "製品規格ラテラルボー", "管理番号", 20)
        Rec "U9-3", "平坦度一覧QP追加の返却順(先頭20件)", "SELECT *", FirstColumnValues(cn, "平坦度一覧QP追加", "フラグ", 20)
        cn.Close
    End If
    Set cn = OpenDb("SIKALOTNOW.accdb")
    If cn Is Nothing Then
        Rec "U4", "SIKALOTNOW.accdb", DB_FOLDER, "ファイルなし"
    Else
        Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
        rs.Open "SELECT TOP 1 * FROM 仕掛", cn, 3, 1
        Dim v As Variant
        On Error Resume Next
        Err.Clear: v = rs.Fields("オーダー番号").Value: Rec "U4-1", "rs.Fields(""オーダー番号"")(全角。実際の列は半角)", "UFCask:436", v, Err.Number
        Err.Clear: v = rs.Fields("ｵｰﾀﾞｰ番号").Value: Rec "U4-2", "rs.Fields(""ｵｰﾀﾞｰ番号"")(半角)", "比較用", v, Err.Number
        Err.Clear: v = rs.Fields("設計_設備コース").Value: Rec "U4-3", "rs.Fields(""設計_設備コース"")(全角)", "UFCask:445", v, Err.Number
        On Error GoTo 0
        rs.Close
        cn.Close
    End If
End Sub

Private Function FirstColumnValues(cn As Object, table As String, col As String, Optional n As Long = 10) As String
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    Dim out As String, i As Long
    On Error Resume Next
    rs.Open "SELECT * FROM " & table, cn, 3, 1
    Do Until rs.EOF Or i >= n
        out = out & IIf(out = "", "", ",") & CStr(rs.Fields(col).Value & "")
        rs.MoveNext
        i = i + 1
    Loop
    If Err.Number <> 0 Then out = out & " (エラー " & Err.Number & ")"
    rs.Close
    On Error GoTo 0
    FirstColumnValues = out
End Function
