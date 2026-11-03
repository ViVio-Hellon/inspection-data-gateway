Attribute VB_Name = "【UFKen_03】SaveLoad"
Option Explicit
Public SikaAtuAll As Variant
Public SikaSetAtu As Single

'======================================================
'LSelect用キャッシュ変数_HistorySaveで格納_LSelectで参照
'======================================================
Public g_CacheLotNo    As String  'キャッシュ対象ロットNo
Public g_CacheLotAcc   As Variant 'SIKALOTNOW 該当ロットレコード(1行目)
Public g_CacheKenNo    As String  '検査番号頭1桁
Public g_CacheCours    As String  '設計_設備コース値
Public g_CacheOrdAtu   As Single  'オーダー板厚値
Public g_CacheC_Course As String  '実績_設備コース値

Public Function SikaAtuCreate(LotNo As String) As Variant '1次元
If LotNo = "" Then Exit Function
Dim TableName As String:       TableName = "仕掛"
Dim FPath As String:           FPath = MainPath
Dim Target_Hiki As String:     Target_Hiki = "SIKALOTNOW.accdb"
Dim FilePath_Hiki As String:   FilePath_Hiki = FPath & Target_Hiki
Dim fieldsfilter As String:    fieldsfilter = ""
Dim Hikifields As Variant:     Hikifields = GetFieldsArr(FilePath_Hiki, TableName, fieldsfilter)
If IsEmpty(Hikifields) Then
  FPath = NextPath
  FilePath_Hiki = FPath & Target_Hiki
  Hikifields = GetFieldsArr(FilePath_Hiki, TableName, fieldsfilter)
  If IsEmpty(Hikifields) Then Exit Function
End If

Dim i As Long, H_オーダー板厚 As Long
For i = LBound(Hikifields, 2) To UBound(Hikifields, 2)
  If Hikifields(1, i) = "ｵｰﾀﾞｰ板厚" Then
    H_オーダー板厚 = i
    Exit For
  End If
Next i

Dim SQLfilter As String:   SQLfilter = "[ﾛｯﾄ番号]='" & LotNo & "'" 'ﾛｯﾄNo検索
Dim TempHiki As Variant:   TempHiki = GetRecordsArr(FilePath_Hiki, TableName, SQLfilter)
If IsEmpty(TempHiki) Then
  FPath = NextPath
  FilePath_Hiki = FPath & Target_Hiki
  TempHiki = GetRecordsArr(FilePath_Hiki, TableName, SQLfilter)
  If IsEmpty(TempHiki) Then Exit Function
End If

Dim OrderAtu As Variant, H_count As Long: H_count = UBound(TempHiki, 1)
ReDim OrderAtu(1 To H_count)
For i = LBound(TempHiki, 1) To UBound(TempHiki, 1)
  OrderAtu(i) = TempHiki(i, H_オーダー板厚)
Next i
SikaAtuCreate = OrderAtu
End Function

Public Sub AccReflec(Optional WorkTime As String = "S") '2020.06.19 '2020.09.03 WorkTime="S"　or "T"
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption

'>>>>>>>>>>>>>>>>>>元ファイル更新中に.accdbに干渉しないようにする_2022.03.17
'Dim Fname As String: Fname = "PLEASE WAIT.txt"
'Dim fPath As String: fPath = "\\nlmsrvngy03\Read\【New】仕掛\台帳\"
'If FileExistence(Fname, fPath, Lname) = False Then Exit Sub

With UF1
  .TaikText.BackStyle = 0
  .LateralText.BackStyle = 0
  .Sumisei.Caption = "" 'ﾌﾗｸﾞﾘｾｯﾄ
  .TaikText = ""
  .LateralText = ""
  .LateralText.Tag = "" '2026.07.13_strShijiLateﾌﾗｸﾞ再読込用もリセット
  .DebugTaikaku = ""
  .DebugLate = ""
  .Additional.Caption = ""
  .Comment.Caption = ""
  .Reien.Caption = ""
  .KenH.Caption = ""
End With

'Acc展開用
If UF1.Lot = "" Then Exit Sub
Dim i As Long, a As Long, z As Long, k As Long, j As Long

Dim db      As Object
Dim rs      As Object
Dim strPath As String
Dim strSQL  As String

Dim PastHiki(0 To 18)
Dim NoFl    As Boolean: NoFl = False
Dim tempE   As String
Dim HCours  As String
Dim G_flag  As Boolean: G_flag = False
Dim itti    As Boolean: itti = False

On Error GoTo ErrHandler_AccReflec

'------------------------------------------------------
'パス決定ロジック_2026.06.03
'優先順位：PLEASE WAIT.txt > DBflag > MainPath
'------------------------------------------------------
Dim strBase As String
If Dir(MainPath & "PLEASE WAIT.txt") <> "" Then
    strBase = NextPath
    Call DebugLog("[AccReflec] PLEASE WAIT.txt検出 NextPath優先")
ElseIf UF1.DBflag.Text <> "" Then
    strBase = NextPath
    Call DebugLog("[AccReflec] DBflag検出 NextPath使用 DBflag=" & UF1.DBflag.Text)
Else
    strBase = MainPath
    Call DebugLog("[AccReflec] MainPath使用")
End If
strPath = strBase & "SIKALOTNOW.accdb"
'------------------------------------------------------

If Dir(strPath) = "" Then
    MsgBox "DBファイルが見つかりません:" & vbCrLf & strPath, vbCritical
    Exit Sub
End If

strSQL = "SELECT * FROM 仕掛 WHERE ﾛｯﾄ番号 = '" & UF1.CB.Text & "' AND ｵｰﾀﾞｰ板丈 <> '0'"

Set db = CreateObject("ADODB.Connection")
db.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strPath & ";"
Set rs = CreateObject("ADODB.Recordset")
rs.Open strSQL, db, 3, 1

If rs.EOF Then
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    GoTo SkipNoItti
End If

NoFl = True

'外注チェック
Dim g_Cours As String: g_Cours = rs.Fields("設計_設備ｺｰｽ").Value
HCours = rs.Fields("実績_設備ｺｰｽ").Value

Select Case True
Case InStr(g_Cours, "GFS") > 0
    tempE = "GFSｺｰｽ有、進度にあった製品ｻｲｽﾞ入力をお願いします。"
    G_flag = True
Case InStr(g_Cours, "GCT") > 0
    tempE = "GCTｺｰｽ有、進度にあった製品ｻｲｽﾞ入力をお願いします。"
    G_flag = True
Case InStr(g_Cours, "GSS") > 0
    tempE = "GSSｺｰｽ有、進度にあった製品ｻｲｽﾞ入力をお願いします。"
    G_flag = True
End Select

'設備仕様チェック
Dim 製造Zai  As String: 製造Zai = rs.Fields("製造材質").Value
Dim Order厚  As Single: Order厚 = val(rs.Fields("ｵｰﾀﾞｰ板厚").Value)
Dim Order幅  As Single: Order幅 = val(rs.Fields("ｵｰﾀﾞｰ板幅").Value)
Dim Order丈  As Single: Order丈 = val(rs.Fields("ｵｰﾀﾞｰ板丈").Value)
Dim WeightC  As Single: WeightC = val(rs.Fields("製品単重").Value)

If G_flag Then
    Order厚 = val(rs.Fields("BOX実績_板厚").Value)
    Order幅 = val(rs.Fields("BOX実績_板幅").Value)
    Order丈 = val(rs.Fields("BOX実績_板丈").Value)
End If

Dim tmpFaci  As String: tmpFaci = GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")
Dim LineCheck As String: LineCheck = CheckEquipmentDb(tmpFaci, 製造Zai, Order厚, Order幅, Order丈, WeightC)

If LineCheck <> "可" Then
    Call MsgError("設備仕様から外れています。" & "," & LineCheck & "," & "終了します。", tmpFaci, 2, "END")
    UF1.CB.Tag = "end"
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    Exit Sub
End If

'PastHiki 格納
PastHiki(0) = "'" & rs.Fields("ｵｰﾀﾞｰ番号").Value
PastHiki(1) = rs.Fields("ﾛｯﾄ番号").Value
If G_flag = False Then
    PastHiki(2) = rs.Fields("製造板厚").Value
    PastHiki(3) = rs.Fields("製造板幅").Value
    PastHiki(4) = rs.Fields("製造板丈").Value
Else
    PastHiki(2) = rs.Fields("BOX実績_板厚").Value
    PastHiki(3) = rs.Fields("BOX実績_板幅").Value
    PastHiki(4) = rs.Fields("BOX実績_板丈").Value
End If
PastHiki(5) = rs.Fields("検査番号").Value
PastHiki(6) = rs.Fields("製造材質").Value
PastHiki(7) = rs.Fields("製造調質").Value
PastHiki(8) = rs.Fields("用途ｺｰﾄﾞ").Value
PastHiki(9) = rs.Fields("設計_設備ｺｰｽ").Value
PastHiki(10) = rs.Fields("取引先名").Value
PastHiki(11) = rs.Fields("納入先名").Value
PastHiki(12) = rs.Fields("送り先名").Value
PastHiki(13) = rs.Fields("BOX実績_枚本数").Value
PastHiki(14) = rs.Fields("BOX実績_良品重量").Value
PastHiki(15) = rs.Fields("梱包_合紙").Value
PastHiki(16) = rs.Fields("VC").Value
PastHiki(17) = rs.Fields("品質ｸﾞﾚｰﾄﾞ_ﾗﾃﾗﾙﾎﾞｰ").Value

'フォーム反映
With UF1
    .Zai = rs.Fields("製造材質").Value
    .Tyu = rs.Fields("製造調質").Value
    .Youto = rs.Fields("用途ｺｰﾄﾞ").Value
    .TAIK = rs.Fields("品質ｸﾞﾚｰﾄﾞ_直角度").Value
    .Lateral = rs.Fields("品質ｸﾞﾚｰﾄﾞ_ﾗﾃﾗﾙﾎﾞｰ").Value
    .flatnes = rs.Fields("品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ").Value
    SikaSetAtu = val(rs.Fields("ｵｰﾀﾞｰ板厚").Value)
    .KenH.Caption = "前工程実績" & rs.Fields("BOX実績_枚本数").Value & "枚"
    If FindRightString(HCours) <> "" Then .LastHot.Caption = "最終熱処理工程：" & FindRightString(HCours)
    Select Case True
    Case InStr(HCours, "QF") > 0: .QF.Caption = "QF有り"
    End Select
    If IsNumeric(.TAIK) = False Then .TAIK = "4"
    If IsNumeric(.Lateral) = False Then .Lateral = ""
    If IsNumeric(.flatnes) = False Then .flatnes = ""

    If G_flag = False Then
        .ATU = Format(val(rs.Fields("製造板厚").Value), "0.000")
        .Hab = Format(val(rs.Fields("製造板幅").Value), "0.0")
        .Tak = Format(val(rs.Fields("製造板丈").Value), "0.0")
    Else
        .ATU = Format(val(rs.Fields("BOX実績_板厚").Value), "0.000")
        .Hab = Format(val(rs.Fields("BOX実績_板幅").Value), "0.0")
        .Tak = Format(val(rs.Fields("BOX実績_板丈").Value), "0.0")
    End If

    ' 板厚公差
    Dim dOP As Double: dOP = val(rs.Fields("板厚公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ").Value)
    Dim dOM As Double: dOM = val(rs.Fields("板厚公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ").Value)
    Dim dMP As Double: dMP = val(rs.Fields("板厚公差_製造_ﾌﾟﾗｽ").Value)
    Dim dMM As Double: dMM = val(rs.Fields("板厚公差_製造_ﾏｲﾅｽ").Value)
    If dOP = 0 Or dOM = 0 Then dOP = dMP: dOM = dMM
    .Ap = Format(dOP, "0.000")
    .Am = Format(dOM, "0.000")

    ' 板幅公差
    Dim wOP As Double: wOP = val(rs.Fields("板幅公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ").Value)
    Dim wOM As Double: wOM = val(rs.Fields("板幅公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ").Value)
    Dim wMP As Double: wMP = val(rs.Fields("板幅公差_製造_ﾌﾟﾗｽ").Value)
    Dim wMM As Double: wMM = val(rs.Fields("板幅公差_製造_ﾏｲﾅｽ").Value)
    If wOP = 0 Or wOM = 0 Then wOP = wMP: wOM = wMM
    .Hp = Format(wOP, "0.0")
    .Hm = Format(wOM, "0.0")

    ' 板丈公差
    Dim tOP As Double: tOP = val(rs.Fields("板丈公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ").Value)
    Dim tOM As Double: tOM = val(rs.Fields("板丈公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ").Value)
    Dim tMP As Double: tMP = val(rs.Fields("板丈公差_製造_ﾌﾟﾗｽ").Value)
    Dim tMM As Double: tMM = val(rs.Fields("板丈公差_製造_ﾏｲﾅｽ").Value)
    If tOP = 0 Or tOM = 0 Then tOP = tMP: tOM = tMM
    .Tp = Format(tOP, "0.0")
    .Tm = Format(tOM, "0.0")

    .ApK = Format(val(.ATU) + val(.Ap), "0.000")
    .AmK = Format(val(.ATU) - val(.Am), "0.000")
    .HpK = Format(val(.Hab) + val(.Hp), "0.0")
    .HmK = Format(val(.Hab) - val(.Hm), "0.0")
    .TpK = Format(val(.Tak) + val(.Tp), "0.0")
    .TmK = Format(val(.Tak) - val(.Tm), "0.0")
    
  '寸法まで消すと後々面倒　公差だけ消す_2026.09.18
    If G_flag Then
        .HpK = "": .HmK = "": .TpK = "": .TmK = ""
    End If

    ' 対角・ラテラルボー処理
    Dim 対flag  As Boolean: 対flag = False
    Dim ラテflag As Boolean: ラテflag = False

    If WorkTime = "S" Then
        If NoFl Then Call HistorySave(PastHiki)
    End If

    Dim strShiji As String:    strShiji = rs.Fields("指示_直角度").Value
    Dim strShijiLate As String: strShijiLate = rs.Fields("指示_ﾗﾃﾗﾙﾎﾞｰ").Value

    If strShiji <> "" Then
        .TaikText = Format(ConvertToHalf(strShiji), "0.0")
        .TaikText.BackStyle = 1
        .TaikText.BackColor = RGB(51, 204, 204)
        対flag = True
        Select Case .TAIK
        Case "4": .DebugTaikaku = Format(Diagonal(.Hab, .Tak), "0.0")
        Case "3": .DebugTaikaku = Format(ReDiagonal(.Hab, .Tak), "0.0")
        End Select
    Else
        Select Case .TAIK
        Case "4": .TaikText = Format(Diagonal(.Hab, .Tak), "0.0")
        Case "3": .TaikText = Format(ReDiagonal(.Hab, .Tak), "0.0")
        End Select
    End If

    If strShijiLate <> "" Then
        .LateralText = ConvertToHalf(strShijiLate)
        .LateralText.BackStyle = 1
        .LateralText.BackColor = RGB(51, 204, 204)
        ラテflag = True
        .LateralText.Tag = "有" '2026.07.13_技術指示による表示状態を保存用に記録
        Dim kenNoTemp As String: kenNoTemp = PastHiki(5)
        Call LSelectLateral(kenNoTemp)
        Select Case .Lateral
        Case "4"
            .DebugLate = Format(Lateralbow_J(.Hab, .Tak, .ATU), "0.0")
            If .DebugLate = 0 Then .DebugLate = ""
        Case "3"
            .DebugLate = Format(ReLateralbow(.Hab, .Tak, .ATU), "0.0")
            If .DebugLate = 0 Then .DebugLate = ""
        End Select
    Else
        If .Lateral <> "" Then
            Select Case .Lateral
            Case "4"
                .LateralText = Format(Lateralbow_J(.Hab, .Tak, .ATU), "0.0")
                If .LateralText = 0 Then .LateralText = ""
            Case "3"
                .LateralText = Format(ReLateralbow(.Hab, .Tak, .ATU), "0.0")
                If .LateralText = 0 Then .LateralText = ""
            End Select
        End If
    End If

    Select Case True
    Case 対flag And ラテflag: .Additional.Caption = "対角ﾗﾃﾗﾙﾎﾞ-追記公差あり"
    Case 対flag:              .Additional.Caption = "対角追記公差あり"
    Case ラテflag:            .Additional.Caption = "ﾗﾃﾗﾙﾎﾞ-追記公差あり"
    End Select
'2026.03.05_フリー反り条件変更 '2026.06.05 Q-S-CH-DIM-022-R03 テーブルを見る方式に変更
'    If (UF1.flatnes = 4) And (val(rs.Fields("ｵｰﾀﾞｰ板厚").Value) <= 3#) Then itti = True

End With

'------------------------------------------------------
'LSelect用キャッシュ格納（AccReflec担当分）_2026.06.03
'------------------------------------------------------
g_CacheOrdAtu = val(rs.Fields("ｵｰﾀﾞｰ板厚").Value)
g_CacheC_Course = CStr(rs.Fields("実績_設備ｺｰｽ").Value)
Call DebugLog("[AccReflec] キャッシュ格納完了 / g_CacheOrdAtu=" & g_CacheOrdAtu & " / g_CacheC_Course=" & g_CacheC_Course)
'------------------------------------------------------

rs.Close: Set rs = Nothing
db.Close: Set db = Nothing

SkipNoItti:
  
  With UF1
    If NoFl = False Then 'LOTNO一致がない
      .Zai = "" '材質
      .Tyu = "" '調質
      .Youto = "" '用途
      .ATU = ""  '厚　製造サイズとする
      .Hab = ""  '巾
      .Tak = ""  '丈
      .Ap = ""  '+マイナスを左側にしたい
      .Am = "" '-
      .Hp = ""  '+マイナスを左側にしたい
      .Hm = "" '-
      .Tp = ""  '+マイナスを左側にしたい
      .Tm = "" '-
      .ApK = "" 'マイナスを左側にしたい
      .AmK = "" '- 厚
      .HpK = ""
      .HmK = "" '-巾
      .TpK = ""
      .TmK = "" '-丈
      .TAIK = "" '対角
      .MG = ""
      .HaHiz = ""
      .TaHiz = ""
      .Naka = ""
      .Mimi = ""
      .HURIS = "" 'フリー反り
    Exit Sub
    End If
  End With
  '------------------------------------------------
  If PastHiki(16) = 0 Then Call VC_Clear
  '2026.06.26_意図的に呼び出しVC_Picture内にMaterialsUpperPartがあり走るにはオーダーNOが必要 意図的に止める
  Dim OderTempNo As String: OderTempNo = PastHiki(0)
  If PastHiki(16) = 1 Then Call VC_Picture     'VC画像

'MsgError用コメントの設定
If UF1.ATU <> "" Then Call Atu_OPselect(val(UF1.ATU), Weight(val(UF1.ATU), val(UF1.Hab), val(UF1.Tak))) '作業選択
Dim QfCheck As String

Select Case True '2021.11.24_導電率チェック
Case Conductiv_Check
  UF1.Conductiv.BackColor = RGB(255, 0, 0)
'  Call ConductivLoad'2022.07_一旦停止

  If tempE = "" Then
    tempE = "ハンドの場合、導電率測定が必要です。注意願います。"
  ElseIf tempE <> "" Then
    tempE = tempE & "," & "ハンドの場合、導電率測定が必要です。注意願います。"
  End If
  
  QfCheck = Conductiv_Coolingtime_Check(UF1.Lot) '変数に修正_2024.5.13
  
  'Call MsgError("ハンドの場合、導電率測定が必要です。注意願います。", Lname, 4) '2021.01.27_追加
  If IsNumeric(QfCheck) Then '数字の時**検査可能
    UF1.ConductivL.Visible = True
    UF1.ConductivText.Caption = QfCheck
  ElseIf QfCheck = "" Then '空白_検査不可
    UF1.ConductivL.Visible = False
    MsgBox "指定ファイルに写真確認Noに入力がない" & vbCrLf & "品証連絡", vbExclamation
  ElseIf QfCheck = "LotNoがない" Then '検査不可
    UF1.ConductivL.Visible = False
    MsgBox "指定ファイルにLotNo入力がない" & vbCrLf & "品証連絡", vbExclamation
  ElseIf QfCheck <> "" Then '何か文字列
    UF1.ConductivL.Visible = True
    UF1.ConductivText.Caption = QfCheck
  End If
End Select


'MsgErrorを改行表示させて1回だけの表示にする_2023.7.7
'重複削除---------------------------------------------------------------------------
Dim arr() As String
arr = Split(tempE, ",") ' 文字列をカンマで分割して配列に格納
Dim Dict As Object: Set Dict = CreateObject("Scripting.Dictionary")

For i = LBound(arr) To UBound(arr)
  Dict(arr(i)) = 0 ' 辞書オブジェクトを使用して重複を削除
Next i

Dim Result As String: Result = Join(Dict.Keys, ",") ' 辞書オブジェクトのキーをカンマで結合

If Result <> "" Then Call MsgError(Result, Lname, 4)
'------------------------------------------------------------------------------------
Dim C_受注番号 As Long, n As String
Dim HikiField: HikiField = Arrbox("SIKAHIKINOW.accdb", "板", "Fields", UF1.DBflag.Text)
Dim HikiAcc: HikiAcc = Arrbox("SIKAHIKINOW.accdb", UF1.Lot.Text, "Record", UF1.DBflag.Text)
If Not IsEmpty(HikiField) Then
  For i = LBound(HikiField, 2) To UBound(HikiField, 2)
    Select Case HikiField(1, i)
    Case "受注番号": C_受注番号 = i
    End Select
  Next i
End If
If IsEmptyArray(HikiAcc) Then GoTo SkipMaterials
n = HikiAcc(LBound(HikiAcc, 1), C_受注番号)
Call MaterialsMsgBox(n) 'SikaLotのオーダーの場合すでに完了している場合がある対策_2023.10.11
GoTo skipS

'引当なしの時用、あまり意味はない
SkipMaterials:
n = Replace(PastHiki(0), "'", "")
Call MaterialsMsgBox(n) '注意喚起_2022.02.04


skipS:
Dim SSW_F As Boolean
Select Case True
Case InStr(HCours, "SSW") > 0: SSW_F = True
End Select

'************************
'修正時AccReflec_LSelect_Line_AutoSelectionすべて修正要
'  Dim flag(0 To 7) As Boolean
'  If Weight(val(UF1.atu), val(UF1.hab), val(UF1.tak)) > 1300# Then flag(0) = True '1300kg以上_新5点作業不可
'  If val(UF1.tak) < 2500# Then flag(1) = True '丈が2500mm未満_6点作業不可
'  If Weight(val(UF1.atu), val(UF1.hab), val(UF1.tak)) > 2160# Then flag(2) = True  '2160kg以上_6点作業不可
'  If val(UF1.tak) >= 2500# And val(UF1.tak) < 4000# Then flag(3) = True  '2500mm以上かつ丈が4000mm未満_10点作業不可
'  If Weight(val(UF1.atu), val(UF1.hab), val(UF1.tak)) > 2500# Then flag(4) = True   '2500kg以上_10点作業不可
'  If val(UF1.hab) < 950# Then flag(5) = True   '幅が950mm未満_10点作業不可
'  If val(UF1.tak) > 5000# Then flag(6) = True     '丈が6000mm超え_2022.02.16*****5000に変更(定盤サイズ変更)_2023.10.24
'  If UF1.zai = "75S" And UF1.tyu = "F" And SSW_F Then flag(7) = True 'SSW通過かつ75Sで調質F_表面切削なしバキューム不可
'
'
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
  If IsGaugeRequired(val(UF1.ATU), val(UF1.Hab), val(UF1.Tak), CStr(UF1.Zai), CStr(UF1.Tyu), HCours) Then
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
    Call CBAdd(SikaSetAtu) 'ﾌﾘｰｿﾘ→無し　の場合ﾌﾘｰｿﾘの判定のままになる_2026.1.12
End If '2021.10.22_修正

'If (UF1.ATU) <> "" Then '2021.01.28
'  If (Weight(Val(UF1.ATU), Val(UF1.Hab), Val(UF1.Tak)) > 1300# And Val(UF1.Tak) < 2500#) Or _
'  (Weight(Val(UF1.ATU), Val(UF1.Hab), Val(UF1.Tak)) > 2500#) Then 'ｽﾘﾝｸﾞ検査の時_2021.01.13
'    UF1.LZS.Caption = "ｹﾞｰｼﾞ"
'    UF1.LZU.Caption = "ｹﾞｰｼﾞ"
'    UF1.LZC.Caption = "ｹﾞｰｼﾞ"
'    UF1.LZD.Caption = "ｹﾞｰｼﾞ"
'    UF1.LZE.Caption = "ｹﾞｰｼﾞ"
'    UF1.LGS.Caption = "ｹﾞｰｼﾞ"
'    UF1.LGU.Caption = "ｹﾞｰｼﾞ"
'    UF1.LGC.Caption = "ｹﾞｰｼﾞ"
'    UF1.LGD.Caption = "ｹﾞｰｼﾞ"
'    UF1.LGE.Caption = "ｹﾞｰｼﾞ"
'    Call CBAdd(Val(SikaSetAtu))
'  End If
'End If

''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
''Q-S-QA-ALL-163_板材抜取検査頻度・板厚測定位置基準**板厚3mm以下かつフラットネス4級の際に測定
''Q-S-QA-DIM-022_フリー反り測定基準_発行日 2024年2月29日_対象がサイズ納入用途に変更_2024.3.5
'2026.03.05_フリー反り条件変更 '2026.06.05 Q-S-CH-DIM-022-R03 テーブルを見る方式に変更
'戻す　糞品証
'If itti Then
'  UF1.flatnesText = "ﾌﾘｰｿﾘ要"
'  UF1.LGS.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGU.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGC.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGD.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGE.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGF.Caption = "ﾌﾘｰｿﾘ"
'  Call CBAdd(SikaSetAtu)
'
'  UF1.HizCome.Caption = "ﾌﾘｰ反り" & UF1.HizCome.Tag 'フリーそりがあった場合だけHizCome差し替え
'End If


Dim TempArry As Variant: TempArry = GetLotInfo("Lot情報履歴", UF1.Lot.Text)
If IsEmpty(TempArry) Then Exit Sub

'2026.10.14_引当混在対策：共通判定（ﾗﾃﾗﾙﾎﾞｰ欄の描写はLSelect側で行う）
'フリー反りはDbReturnItem内でレイアウト反映
Dim 検番 As String
Dim NeedLateral As Boolean, NeedRight As Boolean, NeedFlat As Boolean
Call JudgeLotAllocations(TempArry, 検番, NeedLateral, NeedRight, NeedFlat)
Exit Sub

ErrHandler_AccReflec:
  If Not rs Is Nothing Then
    If rs.state = 1 Then rs.Close
    Set rs = Nothing
  End If
  If Not db Is Nothing Then
    db.Close: Set db = Nothing
  End If
  MsgBox "DB読み込みエラー:" & vbCrLf & Err.Description, vbCritical
End Sub
'
'Public Sub AccReflec(Optional WorkTime As String = "S") '2020.06.19 '2020.09.03 WorkTime="S"　or "T"
'Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
'
''>>>>>>>>>>>>>>>>>>元ファイル更新中に.accdbに干渉しないようにする_2022.03.17
''Dim Fname As String: Fname = "PLEASE WAIT.txt"
''Dim fPath As String: fPath = "\\nlmsrvngy03\Read\【New】仕掛\台帳\"
''If FileExistence(Fname, fPath, Lname) = False Then Exit Sub
'
'With UF1
'  .TaikText.BackStyle = 0
'  .LateralText.BackStyle = 0
'  .Sumisei.Caption = "" 'ﾌﾗｸﾞﾘｾｯﾄ
'  .TaikText = ""
'  .LateralText = ""
'  .DebugTaikaku = ""
'  .DebugLate = ""
'  .Additional.Caption = ""
'  .Comment.Caption = ""
'  .Reien.Caption = ""
'  .KenH.Caption = ""
'End With
'
''Acc展開用
'If UF1.LOT = "" Then Exit Sub
'Dim b As String: b = UF1.CB.ListIndex
'Dim i As Long, a As Long, z As Long, k As Long, j As Long
'Dim Order As Long, OrderAtu As Long, LOT As Long, Atu As Long, Haba As Long, Take As Long, KenNo As Long, HikiNo As Long
'Dim Zai As Long, Tyou As Long, UseNo As Long, Cours As Long, Good As Long, Deliver As Long, Sendi As Long
'Dim 製厚P As Long, 製厚M As Long, O厚P As Long, O厚M As Long
'Dim 製巾P As Long, 製巾M As Long, O巾P As Long, O巾M As Long
'Dim 製丈P As Long, 製丈M As Long, O丈P As Long, O丈M As Long, 対角 As Long
'Dim 実績枚 As Long, 実績重 As Long, 合紙 As Long, VC As Long, LateR As Long '2021.07.13_ﾗﾃﾗﾙﾎﾞｰ
'Dim 指示歪巾 As Long, 指示歪丈 As Long, 指示ラテ As Long, 指示対角 As Long, フラットネス As Long, Z_Cours As Long
'Dim C_オーダー板幅 As Long, C_オーダー板丈 As Long, C_製品単重 As Long, C_BOX実績_板厚 As Long, C_BOX実績_板幅 As Long, C_BOX実績_板丈 As Long
'
'Dim HiAcc: HiAcc = Arrbox("SIKALOTNOW.accdb", "", "Fields", UF1.DBflag.Text) '2022.02.28_引数追加
'
'If Not IsEmpty(HiAcc) Then
'  For i = 1 To UBound(HiAcc, 2) '製造'ｵｰﾀﾞｰ公差確認後片方でも0の場合は製造公差を見ることとする
'    Select Case HiAcc(1, i)
'    Case "ｵｰﾀﾞｰ番号":           Order = i
'    Case "ｵｰﾀﾞｰ板厚":           OrderAtu = i
'
'    Case "ｵｰﾀﾞｰ板幅":           C_オーダー板幅 = i '2024.8.19
'    Case "ｵｰﾀﾞｰ板丈":           C_オーダー板丈 = i '2024.8.19
'    Case "製品単重":            C_製品単重 = i '2024.8.19
'
'    Case "ﾛｯﾄ番号":             LOT = i
'    Case "製造板厚":            Atu = i
'    Case "製造板幅":            Haba = i
'    Case "製造板丈":            Take = i
'    Case "検査番号":            KenNo = i
'    Case "製造材質":            Zai = i
'    Case "製造調質":            Tyou = i
'    Case "用途ｺｰﾄﾞ":            UseNo = i
'    Case "設計_設備ｺｰｽ":        Cours = i
'    Case "実績_設備ｺｰｽ":        Z_Cours = i '2023.6.8
'    Case "取引先名":            Good = i
'    Case "納入先名":            Deliver = i
'    Case "送り先名":            Sendi = i
'    Case "板厚公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ": O厚P = i
'    Case "板厚公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ": O厚M = i
'    Case "板厚公差_製造_ﾌﾟﾗｽ":  製厚P = i
'    Case "板厚公差_製造_ﾏｲﾅｽ":  製厚M = i
'    Case "板幅公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ": O巾P = i
'    Case "板幅公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ": O巾M = i
'    Case "板幅公差_製造_ﾌﾟﾗｽ":  製巾P = i
'    Case "板幅公差_製造_ﾏｲﾅｽ":  製巾M = i
'    Case "板丈公差_ｵｰﾀﾞｰ_ﾌﾟﾗｽ": O丈P = i
'    Case "板丈公差_ｵｰﾀﾞｰ_ﾏｲﾅｽ": O丈M = i
'    Case "板丈公差_製造_ﾌﾟﾗｽ":  製丈P = i
'    Case "板丈公差_製造_ﾏｲﾅｽ":  製丈M = i
'    Case "品質ｸﾞﾚｰﾄﾞ_直角度":   対角 = i
'    Case "BOX実績_枚本数":      実績枚 = i
'    Case "BOX実績_良品重量":    実績重 = i
'    Case "梱包_合紙":           合紙 = i
'    Case "VC":                  VC = i
'    Case "品質ｸﾞﾚｰﾄﾞ_ﾗﾃﾗﾙﾎﾞｰ":  LateR = i
'    Case "指示_ﾗﾃﾗﾙﾎﾞｰ":        指示ラテ = i
'    Case "指示_直角度":         指示対角 = i
'    Case "品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ":   フラットネス = i
'
'    '2024.9.17_最重実績設備のサイズ
'    Case "BOX実績_板厚":        C_BOX実績_板厚 = i
'    Case "BOX実績_板幅":        C_BOX実績_板幅 = i
'    Case "BOX実績_板丈":        C_BOX実績_板丈 = i
'    End Select
'  Next i
'  Dim PastHiki(0 To 18) '2020.07.22******2021.07.13_ラテラルボー追加
'  Dim NoFl As Boolean: NoFl = False
'  Dim LotAcc: LotAcc = Arrbox("SIKALOTNOW.accdb", "板", "Record", UF1.DBflag.Text) '2022.02.28_引数追加
'
'  If IsEmptyArray(LotAcc) Then Exit Sub '2022.03.14_空ファイル対策
'
'  '外注チェック
'  Dim tempE As String, HCours As String
'  Dim Keyword As String:  Keyword = "GCT"
'  Dim Keyword2 As String: Keyword2 = "GFS"
'  Dim Keyword3 As String: Keyword2 = "GSS"
''  Dim Keyword4 As String: Keyword2 = "RSW"
'  Dim G_Cours As String
'
'  Dim IttiTemp As Long
'  For i = LBound(LotAcc, 1) To UBound(LotAcc, 1)
'    If LotAcc(i, LOT) = UF1.CB.Text Then 'LOTNO一致がある　Lot情報履歴用配列
'      HCours = LotAcc(i, Z_Cours) '実績コース
'      G_Cours = LotAcc(i, Cours) '設計ｺｰｽ
'      IttiTemp = i
'    End If
'  Next i
'
'  'MsgError用コメントの設定
'  Dim G_flag As Boolean: G_flag = False
'  Select Case True
'  Case InStr(G_Cours, "GFS") > 0
'    tempE = "GFSｺｰｽ有、進度にあった製品ｻｲｽﾞ入力をお願いします。" 'Call MsgError("GFSｺｰｽ有、進度にあった製品ｻｲｽﾞ入力をお願いします。", Lname, 4)
'    G_flag = True
'  Case InStr(G_Cours, "GCT") > 0
'    tempE = "GCTｺｰｽ有、進度にあった製品ｻｲｽﾞ入力をお願いします。"
'    G_flag = True
'  Case InStr(G_Cours, "GSS") > 0
'    tempE = "GSSｺｰｽ有、進度にあった製品ｻｲｽﾞ入力をお願いします。"
'    G_flag = True
''  Case InStr(G_Cours, "RSW") > 0
''    tempE = "RSWｺｰｽ有、進度にあった製品ｻｲｽﾞ入力をお願いします。"
''    G_flag = True
'  End Select
'
'  If IttiTemp = 0 Then GoTo SkipNoItti
'
'  '2024.8.19_設備仕様チェック
'  Dim 製造Zai As String: 製造Zai = LotAcc(IttiTemp, Zai)
'  Dim Order厚 As Single: Order厚 = LotAcc(IttiTemp, OrderAtu)
'  Dim Order幅 As Single: Order幅 = LotAcc(IttiTemp, C_オーダー板幅)
'  Dim Order丈 As Single: Order丈 = LotAcc(IttiTemp, C_オーダー板丈)
'  Dim WeightC As Single: WeightC = LotAcc(IttiTemp, C_製品単重)
'
'  If G_flag Then '外注ありの場合は最終実績コースのサイズ_2024.9.17
'    Order厚 = LotAcc(IttiTemp, C_BOX実績_板厚)
'    Order幅 = LotAcc(IttiTemp, C_BOX実績_板幅)
'    Order丈 = LotAcc(IttiTemp, C_BOX実績_板丈)
'  End If
'
'  Dim tmpFaci As String:   tmpFaci = GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")
'  Dim LineCheck As String: LineCheck = CheckEquipmentDb(tmpFaci, 製造Zai, Order厚, Order幅, Order丈, WeightC)
''    CheckEquipmentDb = "オーダー厚 範囲外"
''    CheckEquipmentDb = "C_オーダー幅 範囲外"
''    CheckEquipmentDb = "オーダー丈 範囲外"
''    CheckEquipmentDb = "単重 範囲外"
'
'  If LineCheck <> "可" Then
'  Call MsgError("設備仕様から外れています。" & "," & LineCheck & "," & "終了します。", tmpFaci, 2, "END")
'  UF1.CB.Tag = "end"
'  Exit Sub
'  End If
'
'SkipNoItti:
'  '------------------------------------------------------------------------------------------------------------
'  Dim itti As Boolean: itti = False
'
'  For i = LBound(LotAcc, 1) To UBound(LotAcc, 1)
'    Select Case True
'    Case LotAcc(i, LOT) = UF1.CB.Text 'LOTNO一致がある　Lot情報履歴用配列
'        NoFl = True
'
'        HCours = LotAcc(i, Z_Cours) '実績コース
'
'        PastHiki(0) = "'" & LotAcc(i, Order)
'        PastHiki(1) = LotAcc(i, LOT)
'
'        '2024.9.17
'        If G_flag = False Then '外注なし
'          PastHiki(2) = LotAcc(i, Atu)
'          PastHiki(3) = LotAcc(i, Haba)
'          PastHiki(4) = LotAcc(i, Take)
'        Else
'          PastHiki(2) = LotAcc(i, C_BOX実績_板厚)
'          PastHiki(3) = LotAcc(i, C_BOX実績_板幅)
'          PastHiki(4) = LotAcc(i, C_BOX実績_板丈)
'        End If
'
'        PastHiki(5) = LotAcc(i, KenNo) '検番
'        PastHiki(6) = LotAcc(i, Zai)
'        PastHiki(7) = LotAcc(i, Tyou)
'        PastHiki(8) = LotAcc(i, UseNo) '用途コード
'        PastHiki(9) = LotAcc(i, Cours) 'ｺｰｽ
'        PastHiki(10) = LotAcc(i, Good) '得意先
'        PastHiki(11) = LotAcc(i, Deliver) '納入先
'        PastHiki(12) = LotAcc(i, Sendi) '送り先
'        PastHiki(13) = LotAcc(i, 実績枚) '実績枚
'        PastHiki(14) = LotAcc(i, 実績重) '実績重
'        PastHiki(15) = LotAcc(i, 合紙) '合紙
'        PastHiki(16) = LotAcc(i, VC) 'VC
'        PastHiki(17) = LotAcc(i, LateR) 'ﾗﾃﾗﾙﾎﾞｰ
'      With UF1 '*************************************フォーム内
'        .Zai = LotAcc(i, Zai) '材質
'        .Tyu = LotAcc(i, Tyou) '調質
'        .Youto = LotAcc(i, UseNo) '用途_2021.01.26
'        .Taik = LotAcc(i, 対角) '対角
'        .Lateral = LotAcc(i, LateR) 'ﾗﾃﾗﾙﾎﾞｰ'2021.07.13_追加
'        .flatnes = LotAcc(i, フラットネス)
'
'        'オーダー厚のセット_2024.5.14
'        SikaSetAtu = val(LotAcc(i, OrderAtu))
'
'        .KenH.Caption = "前工程実績" & LotAcc(i, 実績枚) & "枚"
'
'        '最終熱処理工程選別_2023.06.08
'        If FindRightString(HCours) <> "" Then .LastHot.Caption = "最終熱処理工程：" & FindRightString(HCours)
'        'QFの通過があるか
'        Select Case True
'        Case InStr(HCours, "QF") > 0: .QF.Caption = "QF" & "有り"
'        End Select
'
'        If IsNumeric(.Taik) = False Then .Taik = "4"   '----------------------------------------------------------ｸﾞﾚｰﾄﾞ表示が数字でない時は4級(厳しい)表示 *"."のことがある
'        If IsNumeric(.Lateral) = False Then .Lateral = ""
'        If IsNumeric(.flatnes) = False Then .flatnes = ""
'
'        '2024.9.17
'        If G_flag = False Then '外注なし
'          .Atu = Format(LotAcc(i, Atu), "0.000")  '厚　製造サイズとする
'          .Hab = Format(LotAcc(i, Haba), "0.0")  '巾
'          .Tak = Format(LotAcc(i, Take), "0.0")  '丈
'        Else
'          .Atu = Format(LotAcc(i, C_BOX実績_板厚), "0.000")
'          .Hab = Format(LotAcc(i, C_BOX実績_板幅), "0.0")
'          .Tak = Format(LotAcc(i, C_BOX実績_板丈), "0.0")
'        End If
'
'        .Ap = Format(LotAcc(i, O厚P), "0.000")  '+　先にオーダー公差を入れる_マイナスを左側にしたい
'        .Am = Format(LotAcc(i, O厚M), "0.000") '-
'        If (.Ap = 0 Or .Ap = "") Or (.Am = 0 Or .Am = "") Then 'オーダー公差が片方でも0の場合は製造公差を見ることとする2018.03.30　andからorに変更
'          .Ap = Format(LotAcc(i, 製厚P), "0.000") '+　　マイナスを左側にしたい
'          .Am = Format(LotAcc(i, 製厚M), "0.000") '-
'        End If
'        .Hp = Format(LotAcc(i, O巾P), "0.0") '+
'        .Hm = Format(LotAcc(i, O巾M), "0.0") '-
'        If (.Hp = "" Or .Hp = "0.0") Or (.Hm = "" Or .Hm = "0.0") Then '片方でも0の場合は製造公差を見ることとする2018.03.30　andからorに変更
'          .Hp = Format(LotAcc(i, 製巾P), "0.0") '+
'          .Hm = Format(LotAcc(i, 製巾M), "0.0") '-
'        End If
'        .Tp = Format(LotAcc(i, O丈P), "0.0") '+
'        .Tm = Format(LotAcc(i, O丈M), "0.0") '-
'        If (.Tp = "" Or .Tp = "0.0") Or (.Tm = "" Or .Tm = "0.0") Then '片方でも0の場合は製造公差を見ることとする2018.03.30　andからorに変更
'          .Tp = Format(LotAcc(i, 製丈P), "0.0") '+
'          .Tm = Format(LotAcc(i, 製丈M), "0.0") '-
'        End If
'        .ApK = Format(val(.Atu) + val(.Ap), "0.000") 'マイナスを左側にしたい
'        .AmK = Format(val(.Atu) - val(.Am), "0.000") '- 厚
'        .HpK = Format(val(.Hab) + val(.Hp), "0.0")
'        .HmK = Format(val(.Hab) - val(.Hm), "0.0") '-巾
'        .TpK = Format(val(.Tak) + val(.Tp), "0.0")
'        .TmK = Format(val(.Tak) - val(.Tm), "0.0") '-丈
'
'         '2025.8.8_外注時公差停止　測定だけで良い
'         '2026.04.14_築山課長確認　板厚の測定判定は必要　寸法は測定必要、判定不要　歪は
'        If G_flag Then '外注
'          .Hp = "": .Hm = "": .Tp = "": .Tm = "":  .HpK = "": .HmK = "": .TpK = "": .TmK = ""
'        End If
'       '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>対角ﾗﾃﾗﾙﾎﾞｰ
'       '公差に関してはフォーマット関数を使用せずそのままの桁数で表示させることとする
'       Dim 対flag As Boolean:   対flag = False
'       Dim ラテflag As Boolean: ラテflag = False
'
'      '--------------------------------作業Lot情報履歴_2020.09.03
'      If WorkTime = "S" Then
'        If NoFl Then Call HistorySave(PastHiki) 'PastHiki(0 To 17) 'LOTNO一致がある
'      End If
'
'    '-------------------------------------------------------------------------------------------------------カード記載による公差
'
'If 指示対角 = 0 Then GoTo LineSkip
'          If LotAcc(i, 指示対角) <> "" Then '技術指示がある場合を追加**2022.07
'            '追記公差分
'              .TaikText = Format(ConvertToHalf(LotAcc(i, 指示対角)), "0.0") '数値のみを半角へ変換後スラッシュ前の数値を抜き取る
'              .TaikText.BackStyle = 1
'              .TaikText.BackColor = RGB(51, 204, 204) '特別指示のある場合変色
'              対flag = True
'              Select Case .Taik '等級でみた公差を確認用に載せる
'              Case "4"
'               .DebugTaikaku = Format(Diagonal(.Hab, .Tak), "0.0")   '4級(厳しい)
'              Case "3"
'               .DebugTaikaku = Format(ReDiagonal(.Hab, .Tak), "0.0")
'              End Select
'          ElseIf LotAcc(i, 指示対角) = "" Then
'LineSkip:
'            '追記公差なし
'              Select Case .Taik
'              Case "4"
'               .TaikText = Format(Diagonal(.Hab, .Tak), "0.0")  '4級(厳しい)
'              Case "3"
'               .TaikText = Format(ReDiagonal(.Hab, .Tak), "0.0")
'              End Select
'          End If
'If 指示ラテ = 0 Then GoTo LineSkip2
'          If LotAcc(i, 指示ラテ) <> "" Then '技術指示がある場合を追加**2022.07
'          '追記公差分
'              .LateralText = ConvertToHalf(LotAcc(i, 指示ラテ))
'              .LateralText.BackStyle = 1
'              .LateralText.BackColor = RGB(51, 204, 204) '特別指示のある場合変色
'              ラテflag = True
'
'              Dim kenNoTemp As String: kenNoTemp = PastHiki(5)
'              Call LSelectLateral(kenNoTemp) '--------------------------------------------------------------------------2024.8.27
'
'                Select Case .Lateral '等級でみた公差を確認用に載せる
'                Case "4"
'                 .DebugLate = Format(Lateralbow_J(.Hab, .Tak, .Atu), "0.0")  '4級(厳しい)
'                 If .DebugLate = 0 Then .DebugLate = ""
'                Case "3"
'                 .DebugLate = Format(ReLateralbow(.Hab, .Tak, .Atu), "0.0")
'                 If .DebugLate = 0 Then .DebugLate = ""
'                End Select
'          ElseIf LotAcc(i, 指示ラテ) = "" Then
'LineSkip2:
'              '追記公差なし
'              If .Lateral <> "" Then
'                Select Case .Lateral
'                Case "4"
'                 .LateralText = Format(Lateralbow_J(.Hab, .Tak, .Atu), "0.0")  '4級(厳しい)
'                 If .LateralText = 0 Then .LateralText = ""
'                Case "3"
'                 .LateralText = Format(ReLateralbow(.Hab, .Tak, .Atu), "0.0")
'                 If .LateralText = 0 Then .LateralText = ""
'                End Select
'              End If
'          End If
'        Select Case True
'        Case 対flag And ラテflag: .Additional.Caption = "対角ﾗﾃﾗﾙﾎﾞ-追記公差あり"
'        Case 対flag:              .Additional.Caption = "対角追記公差あり"
'        Case ラテflag:            .Additional.Caption = "ﾗﾃﾗﾙﾎﾞ-追記公差あり"
'        End Select
'
'
''    '---------------Q-S-QA-DIM-179_板材曲がり(ラテラルボー)検査基準 発行日2024年3月5日------------------------------------納入先での処理
''       '白銅追加*****2022.08.24'納入が白銅の時ﾗﾃﾗﾙﾎﾞｰ公差を書き換える先に等級による公差が反映)
''       '**************HistorySaveでラテラルボー必須の納入先かのチェックあり**************
''        If LotAcc(i, Deliver) Like "ﾊｸﾄﾞｳ*" Then '白銅を含む納入先
''          Select Case LotAcc(i, Atu) '指定の板厚**指示書_QA-022-03
''          Case 0.3 To 302.5
''            If LotAcc(i, Zai) = "NP56" Or LotAcc(i, Zai) Like "*52S*" Or LotAcc(i, Zai) Like "*995" Then
''                .LateralText = Format(2, "0.0")
''                .LateralText.BackStyle = 1
''                .LateralText.BackColor = RGB(255, 153, 0) '特別指示のある場合変色
''            ElseIf LotAcc(i, Zai) Like "*17S*" Or LotAcc(i, Zai) Like "*61S*" Then
''                .LateralText = Format(3, "0.0")
''                .LateralText.BackStyle = 1
''                .LateralText.BackColor = RGB(255, 153, 0)
''            End If
''          End Select
''        ElseIf LotAcc(i, Deliver) Like "ﾅﾒｶﾜｱﾙﾐ*" Then '滑川を含む納入先
''          Select Case LotAcc(i, Atu)
''          Case 0.3 To 302.5
''            If LotAcc(i, Zai) = "NP56" Or LotAcc(i, Zai) Like "*52S*" Or LotAcc(i, Zai) Like "*995" Then
''                .LateralText = Format(2, "0.0")
''                .LateralText.BackStyle = 1
''                .LateralText.BackColor = RGB(255, 153, 0) '特別指示のある場合変色
''            ElseIf LotAcc(i, Zai) Like "*17S*" Or LotAcc(i, Zai) Like "*61S*" Then
''                .LateralText = Format(3, "0.0")
''                .LateralText.BackStyle = 1
''                .LateralText.BackColor = RGB(255, 153, 0)
''            End If
''          End Select
''        End If
''       '日本車両追加*****2023.09.05
''        '**************納入が日本車両の時ﾗﾃﾗﾙﾎﾞｰ公差を書き換える**************
''        If LotAcc(i, Deliver) Like "*ﾆﾎﾝｼﾔﾘﾖｳｾｲｿﾞｳ(ｶ*" Then '日本車両を含む納入先
''          Select Case LotAcc(i, Atu)
''          Case 0.8 To 240#
''            If LotAcc(i, Zai) = "5005" Or LotAcc(i, Zai) = "NP56" Or LotAcc(i, Zai) = "U74S" Or LotAcc(i, Zai) Like "*52S*" Then
''              .LateralText = Format(DBdisplayラテラル(UF1.Lot, UF1.DBflag.text), "0")
''              .LateralText.BackStyle = 1
''              .LateralText.BackColor = RGB(255, 153, 0) '特別指示のある場合変色
''            End If
''          End Select
''        End If
''
'
'
'
'
''        '---------------Q-S-QA-DIM-022_フリー反り測定基準_発行日 2024年2月29日_対象がサイズ納入用途に変更---------------
''2026.03.05_フリー反りの検査条件変更
''この修正のせいでほかの修正が生かせないのでまた戻す　糞品証がっ
'       If (UF1.flatnes = 4) And (val(LotAcc(i, OrderAtu)) <= 3#) Then itti = True '3mm以下ﾌﾗｯﾄﾈｽ等級4級_フリー反りフラグ
'
'
''
''        If (LotAcc(i, Deliver) Like "*ｶ)ﾀｶﾁﾃﾞﾝｷｺｳｷﾞﾖｳ*") And (LotAcc(i, Tyou) = "H24") Then
''          Select Case LotAcc(i, Atu)
''          Case 1# To 3#
''            If LotAcc(i, UseNo) = "U19B" Or LotAcc(i, UseNo) = "H19B" Then
''              If LotAcc(i, Zai) = "995" Or LotAcc(i, Zai) = "A995" Or LotAcc(i, Zai) = "G995" Then
''                Select Case LotAcc(i, Haba)
''                Case 1000# To 1020#
''                  If LotAcc(i, Take) = 2000# Then
''                  itti = True
''                  UF1.HizCome.Tag = "20mm以下"
''                  End If
''                End Select
''              End If
''            End If
''          End Select
''        ElseIf (LotAcc(i, Deliver) Like "*ﾆﾂﾎﾟﾝﾌﾙﾊ-ﾌ(ｶ*") And (LotAcc(i, Tyou) = "H18") Then
''          Select Case LotAcc(i, Atu)
''          Case 0.6 To 0.65
''            If LotAcc(i, UseNo) = "F975" Or LotAcc(i, UseNo) = "F97T" Then
''              If LotAcc(i, Zai) = "52S" Or LotAcc(i, Zai) = "A52N" Or LotAcc(i, Zai) = "F52S" Then 'A52「N」であってます
''                Select Case LotAcc(i, Haba)
''                Case 1217# To 1220#
''                  Select Case LotAcc(i, Take)
''                  Case 2438# To 2730#
''                    itti = True
''                    UF1.HizCome.Tag = "40mm以下"
''                  End Select
''                End Select
''              End If
''            End If
''          End Select
''        ElseIf (LotAcc(i, Deliver) Like "*ｺｼｲﾓｸｻﾞｲｺｳｷﾞﾖｳ(ｶ*") And (LotAcc(i, Tyou) = "H18") Then
''          Select Case LotAcc(i, Atu)
''          Case 0.6
''            If LotAcc(i, UseNo) = "F975" Or LotAcc(i, UseNo) = "F97T" Then
''              If LotAcc(i, Zai) = "52S" Or LotAcc(i, Zai) = "A52N" Or LotAcc(i, Zai) = "F52S" Then 'A52「N」であってます
''                Select Case LotAcc(i, Haba)
''                Case 1220#
''                  Select Case LotAcc(i, Take)
''                  Case 2058# To 2440#
''                    itti = True
''                    UF1.HizCome.Tag = "40mm以下"
''                  End Select
''                End Select
''              End If
''            End If
''          End Select
''        End If
'
'      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'      End With
'      Exit For
'    End Select
'  Next i
'
'
'  With UF1
'    If NoFl = False Then 'LOTNO一致がない
'      .Zai = "" '材質
'      .Tyu = "" '調質
'      .Youto = "" '用途
'      .Atu = ""  '厚　製造サイズとする
'      .Hab = ""  '巾
'      .Tak = ""  '丈
'      .Ap = ""  '+マイナスを左側にしたい
'      .Am = "" '-
'      .Hp = ""  '+マイナスを左側にしたい
'      .Hm = "" '-
'      .Tp = ""  '+マイナスを左側にしたい
'      .Tm = "" '-
'      .ApK = "" 'マイナスを左側にしたい
'      .AmK = "" '- 厚
'      .HpK = ""
'      .HmK = "" '-巾
'      .TpK = ""
'      .TmK = "" '-丈
'      .Taik = "" '対角
'      .MG = ""
'      .HaHiz = ""
'      .TaHiz = ""
'      .Naka = ""
'      .Mimi = ""
'      .HURIS = "" 'フリー反り
'    Exit Sub
'    End If
'  End With
'    '------------------------------------------------
'    If PastHiki(16) = 0 Then Call VC_Clear
'    If PastHiki(16) = 1 Then Call VC_Picture 'VC画像
'End If
'
'
''MsgError用コメントの設定
'If UF1.Atu <> "" Then Call Atu_OPselect(val(UF1.Atu), Weight(val(UF1.Atu), val(UF1.Hab), val(UF1.Tak))) '作業選択
'Dim QfCheck As String
'
'Select Case True '2021.11.24_導電率チェック
'Case Conductiv_Check
'  UF1.Conductiv.BackColor = RGB(255, 0, 0)
''  Call ConductivLoad'2022.07_一旦停止
'
'  If tempE = "" Then
'    tempE = "ハンドの場合、導電率測定が必要です。注意願います。"
'  ElseIf tempE <> "" Then
'    tempE = tempE & "," & "ハンドの場合、導電率測定が必要です。注意願います。"
'  End If
'
'  QfCheck = Conductiv_Coolingtime_Check(UF1.LOT) '変数に修正_2024.5.13
'
'  'Call MsgError("ハンドの場合、導電率測定が必要です。注意願います。", Lname, 4) '2021.01.27_追加
'  If IsNumeric(QfCheck) Then '数字の時**検査可能
'    UF1.ConductivL.Visible = True
'    UF1.ConductivText.Caption = QfCheck
'  ElseIf QfCheck = "" Then '空白_検査不可
'    UF1.ConductivL.Visible = False
'    MsgBox "指定ファイルに写真確認Noに入力がない" & vbCrLf & "品証連絡", vbExclamation
'  ElseIf QfCheck = "LotNoがない" Then '検査不可
'    UF1.ConductivL.Visible = False
'    MsgBox "指定ファイルにLotNo入力がない" & vbCrLf & "品証連絡", vbExclamation
'  ElseIf QfCheck <> "" Then '何か文字列
'    UF1.ConductivL.Visible = True
'    UF1.ConductivText.Caption = QfCheck
'  End If
'End Select
'
'
''MsgErrorを改行表示させて1回だけの表示にする_2023.7.7
''重複削除---------------------------------------------------------------------------
'Dim arr() As String
'arr = Split(tempE, ",") ' 文字列をカンマで分割して配列に格納
'Dim Dict As Object: Set Dict = CreateObject("Scripting.Dictionary")
'
'For i = LBound(arr) To UBound(arr)
'  Dict(arr(i)) = 0 ' 辞書オブジェクトを使用して重複を削除
'Next i
'
'Dim Result As String: Result = Join(Dict.Keys, ",") ' 辞書オブジェクトのキーをカンマで結合
'
'If Result <> "" Then Call MsgError(Result, Lname, 4)
''------------------------------------------------------------------------------------
'Dim C_受注番号 As Long, n As String
'Dim HikiField: HikiField = Arrbox("SIKAHIKINOW.accdb", "板", "Fields", UF1.DBflag.Text)
'Dim HikiAcc: HikiAcc = Arrbox("SIKAHIKINOW.accdb", UF1.LOT.Text, "Record", UF1.DBflag.Text)
'If Not IsEmpty(HikiField) Then
'  For i = LBound(HikiField, 2) To UBound(HikiField, 2)
'    Select Case HikiField(1, i)
'    Case "受注番号": C_受注番号 = i
'    End Select
'  Next i
'End If
'If IsEmptyArray(HikiAcc) Then GoTo SkipMaterials
'n = HikiAcc(LBound(HikiAcc, 1), C_受注番号)
'Call MaterialsMsgBox(n) 'SikaLotのオーダーの場合すでに完了している場合がある対策_2023.10.11
'GoTo skipS
'
''引当なしの時用、あまり意味はない
'SkipMaterials:
'n = Replace(PastHiki(0), "'", "")
'Call MaterialsMsgBox(n) '注意喚起_2022.02.04
'
'
'skipS:
'Dim SSW_F As Boolean
'Select Case True
'Case InStr(HCours, "SSW") > 0: SSW_F = True
'End Select
'
''************************
''修正時AccReflec_LSelect_Line_AutoSelectionすべて修正要
'  Dim flag(0 To 7) As Boolean
'  If Weight(val(UF1.Atu), val(UF1.Hab), val(UF1.Tak)) > 1300# Then flag(0) = True '1300kg以上_新5点作業不可
'  If val(UF1.Tak) < 2500# Then flag(1) = True '丈が2500mm未満_6点作業不可
'  If Weight(val(UF1.Atu), val(UF1.Hab), val(UF1.Tak)) > 2160# Then flag(2) = True  '2160kg以上_6点作業不可
'  If val(UF1.Tak) >= 2500# And val(UF1.Tak) < 4000# Then flag(3) = True  '2500mm以上かつ丈が4000mm未満_10点作業不可
'  If Weight(val(UF1.Atu), val(UF1.Hab), val(UF1.Tak)) > 2500# Then flag(4) = True   '2500kg以上_10点作業不可
'  If val(UF1.Hab) < 950# Then flag(5) = True   '幅が950mm未満_10点作業不可
'  If val(UF1.Tak) > 5000# Then flag(6) = True     '丈が6000mm超え_2022.02.16*****5000に変更(定盤サイズ変更)_2023.10.24
'  If UF1.Zai = "75S" And UF1.Tyu = "F" And SSW_F Then flag(7) = True 'SSW通過かつ75Sで調質F_表面切削なしバキューム不可
'
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
'
'    UF1.LZS.Caption = "幅ｹﾞｰｼﾞ"
'    UF1.LZU.Caption = "幅ｹﾞｰｼﾞ"
'    UF1.LZC.Caption = "幅ｹﾞｰｼﾞ"
'    UF1.LZD.Caption = "幅ｹﾞｰｼﾞ"
'    UF1.LZE.Caption = "幅ｹﾞｰｼﾞ"
'    UF1.LZF.Caption = "幅ｹﾞｰｼﾞ" '2024.5.28_追加
'
'    UF1.LGS.Caption = "丈ｹﾞｰｼﾞ"
'    UF1.LGU.Caption = "丈ｹﾞｰｼﾞ"
'    UF1.LGC.Caption = "丈ｹﾞｰｼﾞ"
'    UF1.LGD.Caption = "丈ｹﾞｰｼﾞ"
'    UF1.LGE.Caption = "丈ｹﾞｰｼﾞ"
'    UF1.LGF.Caption = "丈ｹﾞｰｼﾞ"
'    Call CBAdd(SikaSetAtu)
'  Else
'    UF1.LZS.Caption = "定盤"
'    UF1.LZU.Caption = "定盤"
'    UF1.LZC.Caption = "定盤"
'    UF1.LZD.Caption = "定盤"
'    UF1.LZE.Caption = "定盤"
'    UF1.LZF.Caption = "定盤"
'    UF1.LGS.Caption = "ｹﾞｰｼﾞ"
'    UF1.LGU.Caption = "ｹﾞｰｼﾞ"
'    UF1.LGC.Caption = "ｹﾞｰｼﾞ"
'    UF1.LGD.Caption = "ｹﾞｰｼﾞ"
'    UF1.LGE.Caption = "ｹﾞｰｼﾞ"
'    UF1.LGF.Caption = "ｹﾞｰｼﾞ"
'    Call CBAdd(SikaSetAtu) 'ﾌﾘｰｿﾘ→無し　の場合ﾌﾘｰｿﾘの判定のままになる_2026.1.12
'End If '2021.10.22_修正
'
''If (UF1.ATU) <> "" Then '2021.01.28
''  If (Weight(Val(UF1.ATU), Val(UF1.Hab), Val(UF1.Tak)) > 1300# And Val(UF1.Tak) < 2500#) Or _
''  (Weight(Val(UF1.ATU), Val(UF1.Hab), Val(UF1.Tak)) > 2500#) Then 'ｽﾘﾝｸﾞ検査の時_2021.01.13
''    UF1.LZS.Caption = "ｹﾞｰｼﾞ"
''    UF1.LZU.Caption = "ｹﾞｰｼﾞ"
''    UF1.LZC.Caption = "ｹﾞｰｼﾞ"
''    UF1.LZD.Caption = "ｹﾞｰｼﾞ"
''    UF1.LZE.Caption = "ｹﾞｰｼﾞ"
''    UF1.LGS.Caption = "ｹﾞｰｼﾞ"
''    UF1.LGU.Caption = "ｹﾞｰｼﾞ"
''    UF1.LGC.Caption = "ｹﾞｰｼﾞ"
''    UF1.LGD.Caption = "ｹﾞｰｼﾞ"
''    UF1.LGE.Caption = "ｹﾞｰｼﾞ"
''    Call CBAdd(Val(SikaSetAtu))
''  End If
''End If
'
'''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''Q-S-QA-ALL-163_板材抜取検査頻度・板厚測定位置基準**板厚3mm以下かつフラットネス4級の際に測定
'''Q-S-QA-DIM-022_フリー反り測定基準_発行日 2024年2月29日_対象がサイズ納入用途に変更_2024.3.5
''2026.03.05_フリー反り条件変更
''戻す　糞品証
'If itti Then
'  UF1.flatnesText = "ﾌﾘｰｿﾘ要"
'  UF1.LGS.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGU.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGC.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGD.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGE.Caption = "ﾌﾘｰｿﾘ"
'  UF1.LGF.Caption = "ﾌﾘｰｿﾘ"
'  Call CBAdd(SikaSetAtu)
'
'  UF1.HizCome.Caption = "ﾌﾘｰ反り" & UF1.HizCome.Tag 'フリーそりがあった場合だけHizCome差し替え
'End If
'
'
'Dim TempArry As Variant: TempArry = GetLotInfo("Lot情報履歴", UF1.LOT.Text)
'If IsEmpty(TempArry) Then Exit Sub
'Dim TempDelivery As String
'
''納入、取引、用途コード、材質、調質、厚、幅、丈、ロットNo、検査No、ｵｰﾀﾞｰNo、包装仕様NO
'Dim 納入 As Variant, 取引 As Variant, 用途コード As String, 材質 As String, 調質 As String
'Dim 厚 As Single, 幅 As Single, 丈 As Single, 検番 As String
'Dim ArrCount As Long: ArrCount = UBound(TempArry, 1)
'ReDim 納入(1 To ArrCount) As Variant
'ReDim 取引(1 To ArrCount) As Variant  '2024.8.5
'j = 0
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
'Dim 納入Check As String, 取引Check As String
'For i = LBound(納入, 1) To UBound(納入, 1)
''  If 納入(i) <> "" Then
'    納入Check = 納入(i)
'    取引Check = 取引(i)
'    TempDelivery = DbReturnItem(納入Check, 取引Check, 用途コード, 材質, 調質, 厚, 幅, 丈) '公差関表示
''  End If
'Next i
'
''Stop
'End Sub

Public Sub HistorySave(PastHiki As Variant) '作業Lot情報履歴_2020.09.03 'fields追加 2020.11.23
If WsCheck("Lot情報履歴") = False Then Call AddWS("Lot情報履歴")
Dim HistoryLot As Worksheet: Set HistoryLot = ThisWorkbook.Sheets("Lot情報履歴")
Dim EndR As Long:   EndR = UR(HistoryLot, 1)
Dim EndC As Long:   EndC = UC(HistoryLot, 1)
Dim HistoryRan As Range
Dim i As Long, a As Long, x As Long, y As Long, FL As Boolean: FL = False

Dim ResultT As Boolean: ResultT = IsOver12Hours '2024.9.3

'同じLotNoがあった場合消す
Dim BufRan As Range: Set BufRan = HistoryLot.Range(HistoryLot.Cells(1, 1), HistoryLot.Cells(EndR, EndC))
Dim buf(): buf = BufRan
For i = 2 To UBound(buf, 1)
  If buf(i, 3) = PastHiki(1) Then '履歴に同LotNoがある_PastHiki,0配列
    x = i '上から
    FL = True
    Exit For
  End If
Next i
For i = UBound(buf, 1) To 2 Step -1 '複数ｵｰﾀﾞｰ時,1行ではないので下からも検索する
  If buf(i, 3) = PastHiki(1) Then
    y = i '下から
    Exit For
  End If
Next i
If FL Then
  HistoryLot.Rows(x & ":" & y).Delete  '行削除
  EndR = UR(HistoryLot, 1) '更新
  EndC = UC(HistoryLot, 1)
End If

'LotNo引当No一致チェック
Dim LotNo As String:     LotNo = PastHiki(1)
Dim TableName As String: TableName = "仕掛"

Dim FPath As String
'------------------------------------------------------
'パス決定ロジック_2026.06.03
'優先順位：PLEASE WAIT.txt > DBflag > シート/フォーム判定 > MainPath
'------------------------------------------------------
If Dir(MainPath & "PLEASE WAIT.txt") <> "" Then
    FPath = NextPath
    Call DebugLog("[HistorySave] PLEASE WAIT.txt検出 NextPath優先")
Else
    Select Case True
    Case ThisWorkbook.Sheets("全ヵ所検査").Visible
        FPath = MainPath
    Case ThisWorkbook.Sheets("キャスク検査").Visible
        FPath = MainPath
    Case UF1.DBflag <> ""
        FPath = NextPath
        Call DebugLog("[HistorySave] DBflag検出 NextPath使用 DBflag=" & UF1.DBflag.Text)
    Case Else
        FPath = MainPath
    End Select
End If
'----------------------------------------------------

Dim target As String:    target = "SIKAHIKINOW.accdb" '---------------------------Lot一覧ｼｰﾄからLoTNo_Accessから引当情報
Dim FilePath As String:  FilePath = FPath & target 'Lot
  If ReadCheck(FilePath, 1) = False Or TimeDesignation Or FileDesignation = False Then  '1：ファイル,2：フォルダ
    Select Case True
    Case FPath = NextPath
      If ResultT = False Then MsgBox "日稲:" & target & "がありません" & "_ReadCheck_HistorySave", vbCritical 'ここでは処理は止めない_2021.06.23
      Exit Sub
    Case FPath = MainPath
      FPath = NextPath
      FilePath = FPath & target
      If ReadCheck(FilePath, 1) = False Then '1：ファイル,2：フォルダ
        If ResultT = False Then MsgBox "日稲:" & target & "がありません" & "_ReadCheck_HistorySave", vbCritical '台帳も日軽稲沢もSikaHikiがない
        Exit Sub
      End If
    End Select
  End If
Dim SQLfilter As String:   SQLfilter = "[ﾛｯﾄ番号]='" & LotNo & "'" 'ﾛｯﾄNo検索
Dim TempHiki As Variant:   TempHiki = GetRecordsArr(FilePath, TableName, SQLfilter)

If IsEmpty(TempHiki) Then 'emptyの場合梱包課作成分を確認
  If ThisWorkbook.Sheets("全ヵ所検査").Visible And UFMitsubishi.LineName.Caption = "Test" Then Exit Sub 'Test環境下の時は確認不要
  If ThisWorkbook.Sheets("キャスク検査").Visible And UFCask.LineName.Caption = "Test" Then Exit Sub 'ｷｬｽｸ中はUF1を干渉させない_2021.07.09
  If Not ThisWorkbook.Sheets("キャスク検査").Visible Then If UF1.Test Then Exit Sub

  Select Case True
  Case FPath = NextPath
    If ResultT = False Then MsgBox target & "に" & LotNo & "ﾃﾞｰﾀがありません" & "_HistorySave", vbCritical 'SikaHikiにLotNoがない
    Exit Sub
  Case FPath = MainPath
    FPath = NextPath
    FilePath = FPath & target
    If ReadCheck(FilePath, 1) = False Then '1：ファイル,2：フォルダ
      If ResultT = False Then MsgBox "日稲:" & target & "がありません" & "_ReadCheck_HistorySave", vbCritical '台帳も日軽稲沢もSikaHikiがない
      Exit Sub
    End If
  End Select
  TempHiki = GetRecordsArr(FilePath, TableName, SQLfilter)
  If IsEmpty(TempHiki) Then
    MsgBox "LOT一覧：LOTNO有_" & "HIKIDB：LOTNO無" & "_HistorySave", vbCritical 'それでもEmptyの時は停止
    Exit Sub
  End If
End If
'ﾌｨｰﾙﾄﾞ名ﾁｪｯｸ　*RNE更新に対応
Dim fieldsfilter As String:  fieldsfilter = ""
Dim Hikifields As Variant:   Hikifields = GetFieldsArr(FilePath, TableName, fieldsfilter)
Select Case True 'emptyの場合停止
Case IsEmpty(Hikifields)
  If IsOver12Hours = False Then MsgBox "Hikifieldなし" & "_HistorySave", vbCritical
  Exit Sub
End Select

'2022.10.03_ファイルの並び変更の伴いフィールドチェック追加
Dim H_受注番号 As Long, H_引当番号 As Long, H_ロット番号 As Long, H_検査番号 As Long, H_引当数量 As Long, H_引当調整NO As Long, H_引当ファイル登録 As Long, H_出荷日 As Long 'Hiki
For i = LBound(Hikifields, 2) To UBound(Hikifields, 2)
    Select Case Hikifields(1, i)
    Case "受注番号":                  H_受注番号 = i
    Case "引当番号":                  H_引当番号 = i
    Case "ﾛｯﾄ番号":                   H_ロット番号 = i
    Case "検査番号":                  H_検査番号 = i
    Case "引当数量":                  H_引当数量 = i
    Case "引当調整NO":                H_引当調整NO = i
    Case "引当ﾌｧｲﾙ登録年月日時分秒":  H_引当ファイル登録 = i
    Case "出荷日":                    H_出荷日 = i
    End Select
Next i

'Call No_Search(LotNo)'ｼｰﾄ書込み

'VC一覧ｵｰﾀﾞｰとSIKAHIKINOWのｵｰﾀﾞｰの一致チェック
SQLfilter = ""
target = "SIKAODRNOW.accdb"
FilePath = FPath & target
Dim VC(): VC = GetFieldsArr(FilePath, TableName, fieldsfilter) 'Arrbox("SIKAODRNOW.accdb", "", "Fields")
If IsEmpty(VC) Then
  Select Case True
  Case FPath = NextPath
    If ResultT = False Then MsgBox target & "に" & "VCFieldﾃﾞｰﾀがありません" & "_HistorySave", vbCritical  'SikaHikiにLotNoがない
    Exit Sub
  Case FPath = MainPath
    FPath = NextPath
    FilePath = FPath & target
    If ReadCheck(FilePath, 1) = False Then '1：ファイル,2：フォルダ
      If ResultT = False Then MsgBox "日稲:" & target & "がありません" & "_ReadCheck_HistorySave", vbCritical '台帳も日軽稲沢もSikaHikiがない
      Exit Sub
    End If
  End Select
  VC = GetRecordsArr(FilePath, TableName, SQLfilter)
  If IsEmpty(VC) Then
    If ResultT = False Then MsgBox "VCFieldData無" & "_HistorySave", vbCritical  'それでもEmptyの時は停止
    Exit Sub
  End If
End If

Dim VCR(): VCR = GetRecordsArr(FilePath, TableName, SQLfilter) 'Arrbox("SIKAODRNOW.accdb", "", "Record")
If IsEmpty(VCR) Then
  Select Case True
  Case FPath = NextPath
    If ResultT = False Then MsgBox target & "に" & "VCRecordﾃﾞｰﾀがありません" & "_HistorySave", vbCritical  'SikaHikiにLotNoがない
    Exit Sub
  Case FPath = MainPath
    FPath = NextPath
    FilePath = FPath & target
    If ReadCheck(FilePath, 1) = False Then '1：ファイル,2：フォルダ
      If ResultT = False Then MsgBox "日稲:" & target & "がありません" & "_ReadCheck_HistorySave", vbCritical '台帳も日軽稲沢もSikaHikiがない
      Exit Sub
    End If
  End Select
  VCR = GetRecordsArr(FilePath, TableName, SQLfilter)
  If IsEmpty(VCR) Then
    If ResultT = False Then MsgBox "VCRecordData無" & "_HistorySave", vbCritical  'それでもEmptyの時は停止
    Exit Sub
  End If
End If

Dim R_VCA As Long, R_VCB As Long, R_Tan As Long, R_Ais As Long, R_Good As Long, R_Deliver As Long, R_Sendi As Long, R_Vcodr As Long
Dim R_Packing As Long '2022.01.27追加
For i = 1 To UBound(VC, 2) 'ﾌｨｰﾙﾄﾞ名
  Select Case VC(1, i)
  Case "受注番号":    R_Vcodr = i
  Case "VC_表":       R_VCA = i
  Case "VC_裏":       R_VCB = i
  Case "製品単重":    R_Tan = i
  Case "合紙":        R_Ais = i
  Case "取引先名称":  R_Good = i
  Case "納入先名称":  R_Deliver = i
  Case "送り先名称":  R_Sendi = i
  Case "包装仕様NO":  R_Packing = i
  End Select
Next i
HistoryLot.Range("B:B,V:V").NumberFormatLocal = "@" '文字列
Dim n As Long: n = 0

'各引当に指定の納入先があるかのチェック※Lotのみの情報では紐づいた初めのオーダー分しか解らない為
Dim Deliv As Long: Deliv = 0

For i = LBound(VCR, 1) To UBound(VCR, 1)
  For a = LBound(TempHiki, 1) To UBound(TempHiki, 1)
    If TempHiki(a, H_受注番号) = VCR(i, R_Vcodr) Then 'ｵｰﾀﾞｰNoの一致の時
      n = n + 1
      Set HistoryRan = HistoryLot.Range(HistoryLot.Cells(EndR + n, 2), HistoryLot.Cells(EndR + n, UBound(PastHiki) + 2)) '最終行範囲
      HistoryRan = PastHiki 'サイズ貼り付け(製造)
      HistoryLot.Cells(EndR + n, 1) = Now
      HistoryLot.Cells(EndR + n, 19) = VCR(i, R_VCA)
      HistoryLot.Cells(EndR + n, 20) = VCR(i, R_VCB)
      HistoryLot.Cells(EndR + n, 21) = VCR(i, R_Tan)
      HistoryLot.Cells(EndR + n, 22) = TempHiki(a, H_引当番号) '引当No 追加
      HistoryLot.Cells(EndR + n, 2) = TempHiki(a, H_受注番号) 'ｵｰﾀﾞｰ 書換え_VCR(i, 1)
      HistoryLot.Cells(EndR + n, 17) = VCR(i, R_Ais)  '合紙 書換え
      HistoryLot.Cells(EndR + n, 12) = VCR(i, R_Good) '得意先 書換え
      HistoryLot.Cells(EndR + n, 13) = VCR(i, R_Deliver) '納入先 書換え
      Deliv = Deliv + 1
      HistoryLot.Cells(EndR + n, 14) = VCR(i, R_Sendi) '送り先 書換え
      HistoryLot.Cells(EndR + n, 23) = TempHiki(a, H_引当調整NO) '引当調整NO
      HistoryLot.Cells(EndR + n, 24) = TempHiki(a, H_引当数量) '引当数量
      HistoryLot.Cells(EndR + n, 25) = TempHiki(a, H_引当ファイル登録) '引当ﾌｧｲﾙ登録
      HistoryLot.Cells(EndR + n, 26) = VCR(i, R_Packing) '包装仕様NO
      
        '------------------------------------------------------
        'LSelect用キャッシュ格納（HistorySave担当分）_2026.06.03
        '------------------------------------------------------
        If n > 0 Then
            g_CacheLotNo = PastHiki(1)
            g_CacheCours = CStr(PastHiki(9))  'PastHiki(9) = 設計_設備コース
        
            'SIKAHIKINOW から KenNo 取得
            Dim strCacheKen As String: strCacheKen = ""
            Dim iC As Long
            For iC = LBound(TempHiki, 1) To UBound(TempHiki, 1)
                If TempHiki(iC, H_ロット番号) = PastHiki(1) Then
                    strCacheKen = Left(CStr(TempHiki(iC, H_検査番号)), 1)
                    Exit For
                End If
            Next iC
            g_CacheKenNo = strCacheKen
            
            Call DebugLog("[HistorySave] キャッシュ格納完了" & _
                " / g_CacheLotNo=" & g_CacheLotNo & _
                " / g_CacheKenNo=" & g_CacheKenNo & _
                " / g_CacheCours=" & g_CacheCours)
            
        End If
        '------------------------------------------------------
      
    End If
  Next a
Next i

'-------------------------------------------------------------------------------------------------------------
'**2022.8.17*****ﾗﾃﾗﾙﾎﾞｰ測定の納入先が追加される全オーダーをチェックする
'配列に突っ込む
'ReDim TempDeliver(1 To Deliv) As String
'n = 0
'For i = LBound(VCR, 1) To UBound(VCR, 1)
'  For a = LBound(TempHiki, 1) To UBound(TempHiki, 1)
'    If TempHiki(a, H_受注番号) = VCR(i, R_Vcodr) Then 'ｵｰﾀﾞｰNoの一致の時
'      n = n + 1
'      TempDeliver(n) = VCR(i, R_Deliver) '納入先 書換え
'    End If
'  Next a
'Next i

'Stop
'-------------------------------------------------------------------------------------------------------------
''フィルタデータを格納し元の配列にフィルタをかける***2022.08.17
'Dim flagLat As Boolean: flagLat = False
'Dim strSubNames As Variant
'strSubNames = filter(TempDeliver, "ｽﾐﾄﾓｾｲﾐﾂｺｳｷﾞﾖｳ(ｶ")
'If UBound(strSubNames) > -1 Then 'UBound値が-1より大きい場合、値が見つかったことになる
'  UF1.Sumisei.Caption = "住友精密" 'ｽﾐﾄﾓﾌﾗｸﾞ 2021.07.13_追加
'  flagLat = True
'End If
'strSubNames = filter(TempDeliver, "ﾊｸﾄﾞｳ")
'If UBound(strSubNames) > -1 Then
'  UF1.Sumisei.Caption = "白銅"
'  flagLat = True
'End If
'strSubNames = filter(TempDeliver, "ﾆﾎﾝｼﾔﾘﾖｳｾｲｿﾞｳ(ｶ")
'If UBound(strSubNames) > -1 Then
'  UF1.Sumisei.Caption = "日本車両"
'  flagLat = True
'End If
'strSubNames = filter(TempDeliver, "ﾅﾒｶﾜｱﾙﾐ") '2024.3.5_追加
'If UBound(strSubNames) > -1 Then
'  UF1.Sumisei.Caption = "滑川"
'  flagLat = True
'End If

'If flagLat Then Call LSelect("住精")
'-------------------------------------------------------------------------------------------------------------
'Stop
With HistoryLot.Cells
   .Font.Name = "メイリオ"
   .Font.Size = 10
   .EntireColumn.AutoFit
End With
End Sub

Public Sub VC_Picture(Optional OrderNo As String = "") 'VC画像表示 .JPG"
Dim DefL As Long: DefL = 18
Dim DefT As Long: DefT = 558 '2024.1.19_変更
UF1.VC_L_A.Visible = False
UF1.VC_L_B.Visible = False
UF1.VC_L_A.Caption = "A面"
UF1.VC_L_B.Caption = "B面"
UF1.VC_A.Visible = False
UF1.VC_B.Visible = False
UF1.VcNothing.Visible = False
'位置ﾘｾｯﾄ
UF1.VC_L_A.Left = DefL
UF1.VC_L_A.tOP = DefT
UF1.VC_L_B.Left = DefL
UF1.VC_L_B.tOP = DefT
UF1.VC_A.Left = DefL
UF1.VC_A.tOP = DefT + 20
UF1.VC_B.Left = DefL
UF1.VC_B.tOP = DefT + 20

Dim VCPath As String
Dim AFl As Boolean, BFl As Boolean
Select Case True
Case UF1.Test:  VCPath = PackingPath & "VC種類画像\" 'ThisWorkbook.Path & "\"
Case Else:      VCPath = PackingPath & "VC種類画像\" '梱包課で管理のため両方　PackingPath
End Select

Dim HistorySH As Worksheet: Set HistorySH = ThisWorkbook.Sheets("Lot情報履歴")
Dim LR As Long: LR = UR(HistorySH, 1)
Dim LC As Long: LC = UC(HistorySH, 1)
Dim HistoryR As Range: Set HistoryR = HistorySH.Range(HistorySH.Cells(2, 1), HistorySH.Cells(LR, LC))
Dim TempHiki(): TempHiki = HistoryR
Dim i As Long, x As Long, matchFlag As Boolean: matchFlag = False
For i = 1 To UBound(TempHiki, 1)
  If TempHiki(i, 3) = UF1.Lot Then 'LOTNO一致
    x = i
    matchFlag = True
    Exit For
  End If
Next i
If matchFlag = False Then Exit Sub


Dim VCA As String, VCB As String
If OrderNo = "" Then
    If TempHiki(x, 18) = 1 Then
        If TempHiki(x, 19) <> "" Then
          VCA = TempHiki(x, 19) & ".JPG"
          AFl = True
        End If
        If TempHiki(x, 20) <> "" Then
          VCB = TempHiki(x, 20) & ".JPG"
          BFl = True
        End If
    End If
Else '2022.01.27_追加
    For i = LBound(TempHiki, 1) To UBound(TempHiki, 1) '引当件数分ループ
      If TempHiki(i, 2) = OrderNo Then
          If TempHiki(i, 19) <> "" Then
            VCA = TempHiki(i, 19) & ".JPG"
            AFl = True
          End If
          If TempHiki(i, 20) <> "" Then
            VCB = TempHiki(i, 20) & ".JPG"
            BFl = True
          End If
      End If
    Next i
  
  '上部緩衝材の表示_2022.02.04*引数でオーダーナンバーがある場合のみ
    Dim result1 As String: result1 = MaterialsUpperPart(OrderNo)         ' ザラ板/アングル
    Dim result2 As String: result2 = MaterialsUpperPart(OrderNo, True)   ' 当板有無
    
    If result1 <> "" And result1 <> "一致なし" Then
      MsgBox result1
    End If
    If result1 = "アングル" Then
        If result2 <> "" And result2 <> "一致なし" Then MsgBox result2
    End If
End If
'-------------------------------------------VC確認*2021.04.20
Dim TableName As String: TableName = "VC厚"
Dim FPath As String
Select Case True
Case UF1.Test:  FPath = PackingPath 'ThisWorkbook.Path & "\"
Case Else:      FPath = QualityPath
End Select
'ﾌｧｲﾙ名
Dim target As String:   target = "伝送用ファイル" & ".accdb" 'ファイル名変更_2024.212
'ﾌｧｲﾙﾊﾟｽ
Dim FilePath As String: FilePath = FPath & target
  If ReadCheck(FilePath, 1) = False Then '1：ファイル,2：ﾌｫﾙﾀﾞ
    Dim Fl_filepath As Boolean: Fl_filepath = True
  End If
'検索項目
Dim SQLfilterA As String: SQLfilterA = "[品名ｺｰﾄﾞ]='" & TempHiki(x, 19) & "'" 'VC検索
Dim SQLfilterB As String: SQLfilterB = "[品名ｺｰﾄﾞ]='" & TempHiki(x, 20) & "'"
Dim TempA As Variant: TempA = GetRecordsArr(FilePath, TableName, SQLfilterA)          'All,fields,record
Dim tempB As Variant: tempB = GetRecordsArr(FilePath, TableName, SQLfilterB)

'ﾌｨｰﾙﾄﾞ名ﾁｪｯｸ　*更新に対応
Dim fieldsfilter As String: fieldsfilter = ""
Dim Fields As Variant:  Fields = GetFieldsArr(FilePath, TableName, fieldsfilter)
If IsEmpty(Fields) Then 'emptyの場合
  Dim Fl_fields As Boolean: Fl_fields = True
End If
Dim C_系統 As Long, C_備考 As Long
For i = LBound(Fields, 2) To UBound(Fields, 2)
  Select Case Fields(1, i)
  Case "系統":     C_系統 = i
  Case "備考":     C_備考 = i
  End Select
Next i
Dim Name_A As String, Name_B As String
If Fl_filepath Or Fl_fields Then
  Name_A = ""
  Name_B = ""
End If
If IsEmpty(TempA) Then
  Name_A = "" 'emptyの場合
Else
  Name_A = TempA(1, C_系統) & "_" & TempA(1, C_備考)
End If
If IsEmpty(tempB) Then
  Name_B = ""
Else
  Name_B = tempB(1, C_系統) & "_" & tempB(1, C_備考)
End If

'位置調整     Call VC_Clear
If AFl Or BFl Then UF1.Height = 703.5 '画像有のときのUFのHeight
Dim AutoW As Long
If AFl And BFl Then 'AB
  UF1.VC_L_A.Tag = "on"
  UF1.VC_L_B.Tag = "on"
  UF1.VC_L_A.Visible = True
  UF1.VC_L_B.Visible = True
  UF1.VC_A.Visible = True
  UF1.VC_B.Visible = True
  UF1.VC_L_A.Caption = "A面" & ":" & Name_A
  UF1.VC_L_B.Caption = "B面" & ":" & Name_B
  If ReadCheck((VCPath & VCA), 1) = True Then 'ﾈｯﾄﾜｰｸにつながっていないor登録がない
    UF1.VC_A.Picture = LoadPicture(VCPath & VCA)
  Else
    UF1.VcNothing.Visible = True
    UF1.VcNothing = VCA & "写真登録がないかﾈｯﾄﾜｰｸに繋がっていません：\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\VC種類画像\"
    UF1.VC_L_A.Visible = False
    UF1.VC_L_B.Visible = False
    UF1.VC_A.Visible = False
    UF1.VC_B.Visible = False
  End If
  AutoW = UF1.VC_A.Width + 30
  UF1.VC_B.Left = AutoW
  UF1.VC_L_B.Left = AutoW
  If ReadCheck((VCPath & VCB), 1) = True Then
    UF1.VC_B.Picture = LoadPicture(VCPath & VCB)
  Else
    UF1.VcNothing.Visible = True
    UF1.VcNothing = VCB & "写真登録がないかﾈｯﾄﾜｰｸに繋がっていません：\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\VC種類画像\"
    UF1.VC_L_A.Visible = False
    UF1.VC_L_B.Visible = False
    UF1.VC_A.Visible = False
    UF1.VC_B.Visible = False
  End If
ElseIf AFl And BFl = False Then 'A
  UF1.VC_L_A.Tag = "on"
  UF1.VC_L_A.Visible = True
  UF1.VC_L_A.Caption = "A面" & ":" & Name_A
  UF1.VC_A.Visible = True
  If ReadCheck((VCPath & VCA), 1) = True Then
    UF1.VC_A.Picture = LoadPicture(VCPath & VCA)
    UF1.VC_L_B.Visible = False 'B表示は消しておく
    UF1.VC_B.Visible = False
  Else
    UF1.VcNothing.Visible = True
    UF1.VcNothing = VCA & "写真登録がないかﾈｯﾄﾜｰｸに繋がっていません：\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\VC種類画像\"
    UF1.VC_L_A.Visible = False
    UF1.VC_L_B.Visible = False
    UF1.VC_A.Visible = False
    UF1.VC_B.Visible = False
  End If
ElseIf AFl = False And BFl Then 'B
  UF1.VC_L_B.Tag = "on"
  UF1.VC_L_B.Visible = True
  UF1.VC_L_B.Caption = "B面" & ":" & Name_B
  UF1.VC_B.Visible = True
  If ReadCheck((VCPath & VCB), 1) = True Then
    UF1.VC_B.Picture = LoadPicture(VCPath & VCB)
    UF1.VC_L_A.Visible = False 'A表示は消しておく
    UF1.VC_A.Visible = False
  Else
    UF1.VcNothing.Visible = True
    UF1.VcNothing = VCB & "写真登録がないかﾈｯﾄﾜｰｸに繋がっていません：\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\VC種類画像\"
    UF1.VC_L_A.Visible = False
    UF1.VC_L_B.Visible = False
    UF1.VC_A.Visible = False
    UF1.VC_B.Visible = False
  End If
  AutoW = UF1.VC_B.Width + 30
  UF1.VC_B.Left = AutoW
  UF1.VC_L_B.Left = AutoW
ElseIf AFl = False Or BFl = False Then 'VC=1にもかかわらずVCがない
    UF1.VcNothing.Visible = True
    UF1.VcNothing = TempHiki(x, 2) & ":受注番号一致なし(進度ﾌｧｲﾙ異常)"
End If
End Sub
'Public Sub 公差表示()  '2018.03.19　記入項目追加'2018.03.30片方でも0の場合製造公差を見るよう修正
''Text展開用
'Dim b As String: b = UF1.CB.ListIndex
'Dim i As Long, a As Long, z As Long, k As Long, j As Long
'Dim Obj As Range
'Dim Keyword As String, keyword2 As String
'If WsCheck("LOT一覧") = False Then Exit Sub
'If WsCheck("LOT一覧") Then k = ThisWorkbook.Sheets("LOT一覧").Cells(Rows.Count, 3).End(xlUp).Row
'Dim HIKISH As Worksheet: Set HIKISH = ThisWorkbook.Sheets("LOT一覧")
'Dim SH As Worksheet:     Set SH = ThisWorkbook.Sheets("一覧")
'Dim Lname As String:     Lname = UF1.LineName.Caption & UF1.LineNo.Caption
'With SH
' For i = 2 To k
'   If ThisWorkbook.Sheets("LOT一覧").Cells(i, 3) = UF1.CB.text Then
'     z = z + 1
'     .Cells(z, 1) = HIKISH.Cells(i, 2) 'オーダーNO書き写し
'     .Cells(z, 2) = HIKISH.Cells(i, 3) 'ロットNO
'     .Cells(z, 3) = Format$(HIKISH.Cells(i, 75), "0.000") '厚
'     .Cells(z, 4) = Format$(HIKISH.Cells(i, 76), "0.0") '巾
'     .Cells(z, 5) = Format$(HIKISH.Cells(i, 77), "0.0") '丈
'     .Cells(z, 6) = HIKISH.Cells(i, 4) '検査NO
'     .Cells(z, 7) = HIKISH.Cells(i, 1) '引当NO
'     .Cells(z, 8) = HIKISH.Cells(i, 73) '材質　2019.12.05 オーダー材質から製造材質へ変更
'     .Cells(z, 9) = HIKISH.Cells(i, 74) '調質
'     .Cells(z, 10) = HIKISH.Cells(i, 7) '用途
'     .Cells(z, 11) = HIKISH.Cells(i, 79) 'コース
'     .Cells(z, 12) = HIKISH.Cells(i, 105) '得意先
'     .Cells(z, 13) = HIKISH.Cells(i, 106) '納入先
'     .Cells(z, 14) = HIKISH.Cells(i, 107) '送り先
'     .Cells(z, 15) = HIKISH.Cells(i, 6) '用途コード
'   End If
' Next i
'End With
'With SH.Cells
'   .Font.Name = "メイリオ"
'   .Font.Size = 12
'   .EntireColumn.AutoFit
'End With
''*************************************フォーム内
'   With UF1
'    '125 126 オーダー　127 128 製造
'      .ZAI = HIKISH.Cells(b + 2, 73) '材質
'      .Tyu = HIKISH.Cells(b + 2, 74) '用途
'      .ATU = Format(HIKISH.Cells(b + 2, 75), "0.000") '厚　製造サイズとする
'      .Hab = Format(HIKISH.Cells(b + 2, 76), "0.0") '巾
'      .Tak = Format(HIKISH.Cells(b + 2, 77), "0.0") '丈
'      .AP = Format(HIKISH.Cells(b + 2, 126), "0.000") '+　　マイナスを左側にしたい
'      .AM = Format(HIKISH.Cells(b + 2, 125), "0.000") '-
'     If .AP = "0.000" Or .AM = "0.000" Then '片方でも0の場合は製造公差を見ることとする2018.03.30　andからorに変更
'      .AP = Format(HIKISH.Cells(b + 2, 128), "0.000") '+　　マイナスを左側にしたい
'      .AM = Format(HIKISH.Cells(b + 2, 127), "0.000") '-
'     End If
'      .HP = Format(HIKISH.Cells(b + 2, 130), "0.0") '+
'      .HM = Format(HIKISH.Cells(b + 2, 129), "0.0") '-
'     If .HP = "" Or .HP = "0.0" Or .HM = "" Or .HM = "0.0" Then '片方でも0の場合は製造公差を見ることとする2018.03.30　andからorに変更
'      .HP = Format(HIKISH.Cells(b + 2, 132), "0.0") '+
'      .HM = Format(HIKISH.Cells(b + 2, 131), "0.0") '-
'     End If
'      .TP = Format(HIKISH.Cells(b + 2, 134), "0.0") '+
'      .TM = Format(HIKISH.Cells(b + 2, 133), "0.0") '-
'     If .TP = "" Or .TP = "0.0" Or .TM = "" Or .TM = "0.0" Then '片方でも0の場合は製造公差を見ることとする2018.03.30　andからorに変更
'      .TP = Format(HIKISH.Cells(b + 2, 136), "0.0") '+
'      .TM = Format(HIKISH.Cells(b + 2, 135), "0.0") '-
'     End If
'      .ApK = Format(Val(.ATU) + Val(.AP), "0.000") 'マイナスを左側にしたい
'      .AmK = Format(Val(.ATU) - Val(.AM), "0.000") '- 厚
'      .HpK = Format(Val(.Hab) + Val(.HP), "0.0")
'      .HmK = Format(Val(.Hab) - Val(.HM), "0.0") '-巾
'      .TpK = Format(Val(.Tak) + Val(.TP), "0.0")
'      .TmK = Format(Val(.Tak) - Val(.TM), "0.0") '-丈
'   End With
'    '外注チェック
'    Keyword = "GCT"
'    Dim Course As Range: Set Course = SH.Cells(1, 10)
'    Set Obj = Course.Find(Keyword, LookAt:=xlPart) '部分一致　xlPart　完全一致　xlWhole
'
'    If Obj Is Nothing Then
'    keyword2 = "GFS"
'    Set Obj = Course.Find(keyword2, LookAt:=xlPart)
'    If Obj Is Nothing Then
'    Else
'    Call MsgError("GFSｺｰｽ有、検入Size入力をお願いします。", Lname)
'    End If
'    Else
'    Call MsgError("GCTｺｰｽ有、進度にあったSize入力をお願いします。", Lname)
'    End If
'    '桁表示
'    Dim Thick As String: Thick = UF1.ATU
'    If 0 < Val(Thick) <= 74# Then
'    Range("B14:B18,D14:D18,F14:F18,H14:H18").NumberFormatLocal = "0.000_);[赤](0.000)"
'    ElseIf Val(Thick) > 74 Then
'    Range("B14:B18,D14:D18,F14:F18,H14:H18").NumberFormatLocal = "0.00_);[赤](0.00)"
'    End If
'End Sub

Public Sub ReflectStrain(n As String, Optional DefaultB As String = "") '歪のみの写し
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("作業保存" & "_" & Lname) = False Then Exit Sub '-------------------------ｼｰﾄなし処理停止
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("作業保存" & "_" & Lname)

Dim hizumi_Arry(1 To 6, 1 To 2) As String, i As Long, a As Long
Dim hizumi(): hizumi = UF_Array(15, 16, 1, 6) '格納_2024.5.23*6行面(F)追加

For i = 1 To 6 '配列入れ直し
  For a = 1 To 2
    hizumi_Arry(i, a) = hizumi(i, 14 + a) 'hizumi_Arryには今表示されている数値が入る
  Next a
Next i

Dim A_range As Range: Set A_range = sh.Range(sh.Cells(2, 15), sh.Cells(7, 16)) 'A面
Dim B_range As Range: Set B_range = sh.Range(sh.Cells(8, 15), sh.Cells(13, 16)) 'B面
Dim tagHizumi()

'hizumi_Arry*****tagからデータ取り出す
'Split_Zy********tagにデータ保存

If DefaultB = "" Then
  'ボタンを押した時のイベント(ボタン表示逆反応)
  If n = "B" Then 'AだったときBに切り替えるので残っているデータはA
    A_range = hizumi_Arry 'A保存
    Call LosdStrain("B", A_range, B_range) 'B呼び出し
  ElseIf n = "A" Then
    B_range = hizumi_Arry 'B保存
    Call LosdStrain("A", A_range, B_range) 'A呼び出し
  End If
ElseIf DefaultB = "保存" Then
  'ボタンを押さずにするイベント
  If n = "B" Then 'そのままB保存　呼び出しなし
    B_range = hizumi_Arry 'B保存
    
    tagHizumi = UFtag_Array(1, 6) 'Aﾃﾞｰﾀ

    A_range = tagHizumi
  ElseIf n = "A" Then 'そのままA保存　呼び出しなし
    A_range = hizumi_Arry  'A保存
    
    tagHizumi = UFtag_Array(1, 6) 'Bﾃﾞｰﾀ

    B_range = tagHizumi
  End If
End If
End Sub

Public Sub LosdStrain(Temp As String, Aran As Range, Bran As Range) '歪のみ
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("作業保存" & "_" & Lname) = False Then Exit Sub '-------------------------ｼｰﾄなし処理停止
Dim StraA(): StraA = Aran
Dim StraB(): StraB = Bran
Dim ZAdata, GAdata, データ, 反対面, a As Long, i As Long
Dim ZBdata, GBdata

Application.EnableEvents = False
If Temp = "B" Then 'B呼び出し
  データ = StraB
  反対面 = StraA
End If
If Temp = "A" Then 'A呼び出し
  データ = StraA
  反対面 = StraB
End If

For i = 1 To 2
  For a = 1 To 6
    ZAdata = Split(データ(a, i), "_") 'Split関数0配列
    ZBdata = Split(反対面(a, i), "_")
    GAdata = Split(データ(a, i), "_")
    GBdata = Split(反対面(a, i), "_")
    
    Select Case i
    Case 1 '定盤歪--------------3つの入力がない場合Loadなし
      Select Case a
      Case 1 To 6
        If IsEmptyArray(ZAdata) Then ' 配列が空の場合
          Select Case a
          Case 1: Split_Z_0 ("S")
          Case 2: Split_Z_0 ("U")
          Case 3: Split_Z_0 ("C")
          Case 4: Split_Z_0 ("D")
          Case 5: Split_Z_0 ("E")
          Case 6: Split_Z_0 ("F")
          End Select
        ElseIf UBound(ZAdata) = 2 Then ' 添え字上限が3である場合
          Select Case a
          Case 1: Call Split_Z("S", ZAdata, ZBdata)
          Case 2: Call Split_Z("U", ZAdata, ZBdata)
          Case 3: Call Split_Z("C", ZAdata, ZBdata)
          Case 4: Call Split_Z("D", ZAdata, ZBdata)
          Case 5: Call Split_Z("E", ZAdata, ZBdata)
          Case 6: Call Split_Z("F", ZAdata, ZBdata)
          End Select
        End If
      End Select
    Case 2 'ｹﾞｰｼﾞ歪
        Select Case a
        Case 1 To 6
        If IsEmptyArray(GAdata) Then ' 配列が空の場合
          Select Case a
          Case 1: Split_G_0 ("S")
          Case 2: Split_G_0 ("U")
          Case 3: Split_G_0 ("C")
          Case 4: Split_G_0 ("D")
          Case 5: Split_G_0 ("E")
          Case 6: Split_G_0 ("F")
          End Select
        ElseIf UBound(GAdata) = 2 Then ' 添え字上限が3である場合
          Select Case a
          Case 1: Call Split_G("S", GAdata, GBdata)
          Case 2: Call Split_G("U", GAdata, GBdata)
          Case 3: Call Split_G("C", GAdata, GBdata)
          Case 4: Call Split_G("D", GAdata, GBdata)
          Case 5: Call Split_G("E", GAdata, GBdata)
          Case 6: Call Split_G("F", GAdata, GBdata)
          End Select
        End If
      End Select
    End Select
  Next a
Next i
Application.EnableEvents = True
End Sub

Public Sub Posting() '写し
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Call AddWS("作業保存" & "_" & Lname)
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("作業保存" & "_" & Lname)
Dim i As Long, a As Long
Dim ReSize As String: ReSize = UF1.ReSize.Tag 'サイズ修正の有無_2021.01.13
Dim TemRange As Range, DataRange As Range, TextBoxRange As Range
With UF1
Dim 検査 As Variant: 検査 = Array(.Lot.Text, .ATU.Text, .Hab.Text, _
.Tak.Text, .Am.Text, .Ap.Text, .Hm.Text, .Hp.Text, .Tm.Text, .Tp.Text, _
.AmK.Text, .ApK.Text, .HmK.Text, .HpK.Text, .TmK.Text, .TpK.Text, _
.Zai.Text, .Tyu.Text, .Youto.Text, .TAIK.Text, .Machin.Caption, ReSize, .Lateral.Text, .TaikText.Text, .LateralText.Text, .flatnes.Text, _
.Sumisei.Caption, .Additional.Caption) 'Lot情報格納　2021.07.13_ﾗﾃﾗﾙﾎﾞｰ,住精ﾌﾗｸﾞ追加27ｺ_ReSizeのみtag_2022.07*追記公差ﾌﾗｸﾞ追加**2023.4.18_ﾌﾗｯﾄﾈｽ等級追加

'1～28列目（A～AB列）: 変数 検査（Lot情報など28項目）の貼り付け。
'29～32列目（AC～AF列）: 検査頻度（厚板・薄板など）、検査回数、検入枚数、検入山数。
'33～39列目（AG～AM列）: TextBoxArray（歪の公差関連：ゲ幅、丈、中、耳、コメ、コメの7個）。
'40列目（AN列）: UF1.KenCount.Text（検査回数）。※2025.12.17の修正箇所

'1 行目: 基本情報､検査頻度､歪公差などのヘッダー的情報｡
'2～7行目: データA（測定値など）および データB（対角線差、直角度など）。
'8 行目: 削除座標､削除カウント｡
'9行目: 検番、YamaLB、検入高さ（7項目）、高さ計算式。
'10～15行目: 外観（表面・裏面）、BH。
'16 行目: 前工程実績｡
'17 行目: 最終熱処理工程､QFの有無｡

'歪の公差***2024.2.6
Dim TextBoxArray As Variant: TextBoxArray = Array(.MG.Text, .HaHiz.Text, .TaHiz.Text, .Naka.Text, .Mimi.Text, .DistortionFrame.Caption, .HizCome.Caption)
Set TextBoxRange = sh.Range(sh.Cells(1, 33), sh.Cells(1, 39)) '32は山数がくるのでずらす　ゲ幅丈中耳コメコメの7個
TextBoxRange = TextBoxArray

End With '検査~28ｺ_TextBoxArray~5ｺ

'Stop
Set TemRange = sh.Range(sh.Cells(1, 1), sh.Cells(1, UBound(検査) + 1)) '0配列-----28ｺ,(OPﾎﾞﾀﾝ内容回数含め28ｺ)
TemRange = 検査 '貼り付け
Dim LastCol As Long: LastCol = UBound(検査) + 1 + 1 'UC(SH, 1) + 1
Select Case True '検査頻度選択
Case UF1.厚板:   sh.Cells(1, LastCol) = "厚板"
Case UF1.薄板:   sh.Cells(1, LastCol) = "薄板"
Case UF1.少枚数: sh.Cells(1, LastCol) = "少枚数"
Case UF1.極厚:   sh.Cells(1, LastCol) = "極厚"
Case UF1.スラブ: sh.Cells(1, LastCol) = "スラブ"
End Select
'*2023.4.18_ﾌﾗｯﾄﾈｽ等級追加
If sh.Cells(1, LastCol) = "薄板" And UF1.AtButt.Visible Then sh.Cells(1, LastCol + 1) = Mid(UF1.AtCount, 1, 4)   '29***厚検査回数(薄板のみ)_2021.01.13 最終列修正
If sh.Cells(1, LastCol) = "薄板" And UF1.InspeButt.Visible Then sh.Cells(1, LastCol + 2) = UF1.InspeButt.Caption '30**2021.11.29_薄板検入枚数
If (sh.Cells(1, LastCol) = "薄板") And (UF1.Yama.Text <> "") Then sh.Cells(1, LastCol + 3) = UF1.Yama.Text '31**2021.12.15_薄板検入山数’2021.12.20_3山の時だけになっていたので修正
'Dim データ(1 To 5, 1 To 17) As Variant
Dim データA(), データB()
データA = UF_Array(1, 14, 1, 6) '格納**********************2022.06_耳付き丈3,4追加分+2　2024.5.23*6行面(F)追加
データB = UF_Array(17, 20, 1, 6) '対角線差、直角度、公差外判定ﾌﾗｸ、ﾞﾗﾃﾗﾙﾎﾞｰ
'Stop
Dim LR As Long: LR = UR(sh, 1)
Set DataRange = sh.Range(sh.Cells(2, 1), sh.Cells(7, 14))
DataRange = データA '貼り付け

Dim Temp_Arry(1 To 6, 1 To 4) As String
For i = 1 To 6 '配列入れ直し
  For a = 1 To 4
    Temp_Arry(i, a) = データB(i, 16 + a)
  Next a
Next i
Dim Temp_range As Range: Set Temp_range = sh.Range(sh.Cells(2, 17), sh.Cells(7, 20))
Temp_range = Temp_Arry

'歪部分上書き_2023.6.13
Call ReflectStrain(Right(UF1.AorB.Caption, 1), "保存")
'-------------------------------------------------------------2024.5.22_1行ずれます

'2023.4.24_消した座標と数の保存*固定座標　検査結果の下
sh.Cells(8, 1) = UF1.DeleteP.Caption '座標
sh.Cells(8, 2) = UF1.DeleteC.Caption 'delカウント

'2023.5.12_追加
sh.Cells(9, 1) = UF1.KenNo.Caption '検番
'2023.5.22_追加
 sh.Cells(16, 1) = UF1.KenH.Caption '前工程実績
'2023.6.8_追加
 sh.Cells(17, 1) = UF1.LastHot.Caption '最終熱処理工程
'2023.6.9_追加
 sh.Cells(17, 2) = UF1.QF.Caption 'QFの有無

'2023.5.19_追加
With UF1
 '外観
  Dim VisItemA As Variant: VisItemA = Array(.S_UpVisual.Caption, .U_UpVisual.Caption, .C_UpVisual.Caption, .D_UpVisual.Caption, .E_UpVisual.Caption, .F_UpVisual.Caption)
  Dim VisItemB As Variant: VisItemB = Array(.S_UnderVisual.Caption, .U_UnderVisual.Caption, .C_UnderVisual.Caption, .D_UnderVisual.Caption, .E_UnderVisual.Caption, .F_UnderVisual.Caption)
  Dim VIA(1 To 6, 1 To 1), VIB(1 To 6, 1 To 1)
  For a = 0 To 5
    VIA(a + 1, 1) = VisItemA(a)
    VIB(a + 1, 1) = VisItemB(a)
  Next a

  'BH
  Dim BHItem As Variant: BHItem = Array(.SBH, .UBH, .CBH, .DBH, .EBH)
  Dim bh(1 To 5, 1 To 1)
  For a = 0 To 4
    bh(a + 1, 1) = BHItem(a)
  Next a

  Dim VisualA As Range:  Set VisualA = sh.Range(sh.Cells(10, 1), sh.Cells(15, 1)) '外観
  Dim VisualB As Range:  Set VisualB = sh.Range(sh.Cells(10, 2), sh.Cells(15, 2))
  Dim BHR As Range:      Set BHR = sh.Range(sh.Cells(10, 3), sh.Cells(14, 3)) 'BH
  VisualA = VIA
  VisualB = VIB
  BHR = bh
  '検入高さ
  Dim KenHItem As Variant: KenHItem = Array(.KenHBox1, .KenHBox2, .KenHBox3, .KenHBox4, .KenHBox5, .KenHBox6, .KenHBox7)
  Dim KenH(1 To 1, 1 To 7)
  For a = 0 To 6
    KenH(1, a + 1) = KenHItem(a) '列方向
  Next a
  sh.Cells(9, 2) = UF1.YamaLB.Text '.List(YamaLB.ListIndex)
  Dim KenHR As Range:  Set KenHR = sh.Range(sh.Cells(9, 3), sh.Cells(9, 9)) '検入高さ
  KenHR = KenH
  
'2024.10.30
'  SH.Cells(1, 18) = UF1.KenCount.Text '検査回数
  sh.Cells(1, 40) = UF1.KenCount.Text '2025.12.17_バッティングが起きていた
  
'2026.03.05_フリー反り公差
  sh.Cells(1, 41) = UF1.HURIS.Text

'2026.07.13_ﾗﾃﾗﾙﾎﾞｰ欄表示状態(strShijiLate経路)復元用フラグ
  sh.Cells(1, 42) = UF1.LateralText.Tag
  
  
'2023.6.8_追加
  sh.Cells(9, 10) = UF1.高さ計算式.Caption '計算式
End With

'列18_直角度****2023.11.27
Dim ClearRange As Range:  Set ClearRange = sh.Range(sh.Cells(2, 18), sh.Cells(7, 18))

Dim TempArry As Variant: TempArry = GetLotInfo("Lot情報履歴", UF1.Lot.Text) 'ヒットしたロットの配列を作成
If IsEmpty(TempArry) Then Exit Sub

'2026.10.14_引当混在対策：共通判定
'直角度も「1件でも要なら残す」に変更（従来は不要の引当が1件でもあるとクリアされていた）
Dim 検番 As String
Dim NeedLateral As Boolean, NeedRight As Boolean, NeedFlat As Boolean
Call JudgeLotAllocations(TempArry, 検番, NeedLateral, NeedRight, NeedFlat)
If Not NeedRight Then ClearRange.ClearContents '直角度測定が不要ならばクリア

With sh.Cells
  .Font.Name = "メイリオ"
  .Font.Size = 12
  .ColumnWidth = 10.25
  .EntireColumn.AutoFit
End With
End Sub

Public Function UF_Array(n As Long, k As Long, x As Long, y As Long) As Variant   '検査結果格納 *(列,行)並びが逆になってしまった....
'"許容"ﾌﾗｸﾞなしの公差外は格納しない　2020.08.08　BackColor_Check
Dim i As Long, a As Long
ReDim データ(x To y, n To k) As Variant
For i = n To k '検査結果格納
  For a = x To y
    Select Case i
    Case 1 To 4 '厚
      Select Case a
      Case 1: データ(1, i) = UF1("S" & i).Caption
      Case 2: データ(2, i) = UF1("U" & i).Caption
      Case 3: データ(3, i) = UF1("C" & i).Caption
      Case 4: データ(4, i) = UF1("D" & i).Caption
      Case 5: データ(5, i) = UF1("E" & i).Caption
      Case 6: データ(6, i) = UF1("F" & i).Caption
      End Select
    Case 5 To 10 '巾丈***********2022.06_耳付き丈3,4追加
      Select Case a
      Case 1: データ(1, i) = UF1("S" & i).Caption
      Case 2: データ(2, i) = UF1("U" & i).Caption
      Case 3: データ(3, i) = UF1("C" & i).Caption
      Case 4: データ(4, i) = UF1("D" & i).Caption
      Case 5: データ(5, i) = UF1("E" & i).Caption
      Case 6: データ(6, i) = UF1("F" & i).Caption 'Fは10までしかない
      End Select
    Case 11 To 14 '巾丈***********2022.06_耳付き丈3,4追加
      Select Case a
      Case 1: データ(1, i) = UF1("S" & i).Caption
      Case 2: データ(2, i) = UF1("U" & i).Caption
      Case 3: データ(3, i) = UF1("C" & i).Caption
      Case 4: データ(4, i) = UF1("D" & i).Caption
      Case 5: データ(5, i) = UF1("E" & i).Caption
      Case 6: データ(6, i) = "" 'Fなし
      End Select
    Case 15 '定盤歪_2021.04.08*If .ZS <> "" Then*削除ﾌｫｰｶｽを移す前に閉じる人が多かった
    With UF1
      Select Case a
      Case 1: データ(1, i) = .CBS1 & "_" & .CBS2 & "_" & .ZS
      Case 2: データ(2, i) = .CBU1 & "_" & .CBU2 & "_" & .ZU
      Case 3: データ(3, i) = .CBC1 & "_" & .CBC2 & "_" & .ZC
      Case 4: データ(4, i) = .CBD1 & "_" & .CBD2 & "_" & .ZD
      Case 5: データ(5, i) = .CBE1 & "_" & .CBE2 & "_" & .ZE
      Case 6: データ(6, i) = .CBF1 & "_" & .CBF2 & "_" & .ZF
      End Select
    End With
    Case 16 'ｹﾞｰｼﾞ歪
    With UF1
      Select Case a
      Case 1: データ(1, i) = .CBS3 & "_" & .CBS4 & "_" & .GS
      Case 2: データ(2, i) = .CBU3 & "_" & .CBU4 & "_" & .GU
      Case 3: データ(3, i) = .CBC3 & "_" & .CBC4 & "_" & .GC
      Case 4: データ(4, i) = .CBD3 & "_" & .CBD4 & "_" & .GD
      Case 5: データ(5, i) = .CBE3 & "_" & .CBE4 & "_" & .GE
      Case 6: データ(6, i) = .CBF3 & "_" & .CBF4 & "_" & .GF
      End Select
    End With
    Case 17 '対角線差
      Select Case a
      Case 1: データ(1, i) = UF1.DiagS.Caption
      Case 2: データ(2, i) = UF1.DiagU.Caption
      Case 3: データ(3, i) = UF1.DiagC.Caption
      Case 4: データ(4, i) = UF1.DiagD.Caption
      Case 5: データ(5, i) = UF1.DiagE.Caption
      Case 6: データ(6, i) = UF1.DiagF.Caption
      End Select
    Case 18 '対角公差**********2023.11.24_対角公差から直角度に変更
'      If UF1.Youto = "R115" Then
        Select Case a
        Case 1: データ(1, i) = UF1.DiagSKou.Text
        Case 2: データ(2, i) = UF1.DiagUKou.Text
        Case 3: データ(3, i) = UF1.DiagCKou.Text
        Case 4: データ(4, i) = UF1.DiagDKou.Text
        Case 5: データ(5, i) = UF1.DiagEKou.Text
        Case 6: データ(6, i) = UF1.DiagFKou.Text
        End Select
'      End If
    Case 19 '公差外判定ﾌﾗｸﾞ
      Select Case a
      Case 1: データ(1, i) = UF1.NGS.Caption
      Case 2: データ(2, i) = UF1.NGU.Caption
      Case 3: データ(3, i) = UF1.NGC.Caption
      Case 4: データ(4, i) = UF1.NGD.Caption
      Case 5: データ(5, i) = UF1.NGE.Caption
      Case 6: データ(6, i) = UF1.NGF.Caption
      End Select
    Case 20 'ﾗﾃﾗﾙﾎﾞｰ
      Select Case a
      Case 1: データ(1, i) = UF1.LateralS.Text
      Case 2: データ(2, i) = UF1.LateralU.Text
      Case 3: データ(3, i) = UF1.LateralC.Text
      Case 4: データ(4, i) = UF1.LateralD.Text
      Case 5: データ(5, i) = UF1.LateralE.Text
      Case 6: データ(6, i) = UF1.LateralF.Text
      End Select
    End Select
  Next a
Next i
UF_Array = データ
End Function
'シート保存用_2022.06
'耳付き除去用
Public Function UF_SaveArray(n As Long, k As Long, x As Long, y As Long) As Variant
Dim i As Long, a As Long
ReDim データ(x To y, n To k) As Variant
For i = n To k '検査結果格納
  For a = x To y
    Select Case i
    Case 1 To 4 '厚
      Select Case a
      Case 1: データ(1, i) = UF1("S" & i).Caption
      Case 2: データ(2, i) = UF1("U" & i).Caption
      Case 3: データ(3, i) = UF1("C" & i).Caption
      Case 4: データ(4, i) = UF1("D" & i).Caption
      Case 5: データ(5, i) = UF1("E" & i).Caption
      Case 6: データ(6, i) = UF1("F" & i).Caption
      End Select
    Case 5 To 10 '巾丈_耳付きを除く
      Select Case a
      Case 1: データ(1, i) = UF1("S" & i).Caption
      Case 2: データ(2, i) = UF1("U" & i).Caption
      Case 3: データ(3, i) = UF1("C" & i).Caption
      Case 4: データ(4, i) = UF1("D" & i).Caption
      Case 5: データ(5, i) = UF1("E" & i).Caption
      Case 6: データ(6, i) = UF1("F" & i).Caption
      End Select
    Case 11 '定盤歪_2021.04.08*If .ZS <> "" Then*削除ﾌｫｰｶｽを移す前に閉じる人が多かった
    With UF1
      Select Case a
      Case 1: データ(1, i) = .CBS1 & "_" & .CBS2 & "_" & .ZS
      Case 2: データ(2, i) = .CBU1 & "_" & .CBU2 & "_" & .ZU
      Case 3: データ(3, i) = .CBC1 & "_" & .CBC2 & "_" & .ZC
      Case 4: データ(4, i) = .CBD1 & "_" & .CBD2 & "_" & .ZD
      Case 5: データ(5, i) = .CBE1 & "_" & .CBE2 & "_" & .ZE
      Case 6: データ(6, i) = .CBF1 & "_" & .CBF2 & "_" & .ZF
      End Select
    End With
    Case 12 'ｹﾞｰｼﾞ歪
    With UF1
      Select Case a
      Case 1: データ(1, i) = .CBS3 & "_" & .CBS4 & "_" & .GS
      Case 2: データ(2, i) = .CBU3 & "_" & .CBU4 & "_" & .GU
      Case 3: データ(3, i) = .CBC3 & "_" & .CBC4 & "_" & .GC
      Case 4: データ(4, i) = .CBD3 & "_" & .CBD4 & "_" & .GD
      Case 5: データ(5, i) = .CBE3 & "_" & .CBE4 & "_" & .GE
      Case 6: データ(6, i) = .CBF3 & "_" & .CBF4 & "_" & .GF
      End Select
    End With
    End Select
  Next a
Next i
UF_SaveArray = データ
End Function

Public Sub Time_Posting(Position As String, EachT As Date, Optional Temp As String = "1") '作業者,作業時間写し *消すときはTemp=""
If Position = "" Then Exit Sub
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("WorkerTime" & "_" & Lname) = False Then Call AddWS("WorkerTime" & "_" & Lname)

'Lot分割時の処理追記_2023.02.27
If UF1.SplitFlag.Caption <> "" Then 'LotNo登録時空白でないときにボタンを押すと分割かの確認_2023.2.27
'*******フラグとしてフォームに残すだけなのでフォームを閉じるとリセットさせる
    'クリアせず前ロットデータから続行
ElseIf UF1.SplitFlag.Caption = "" Then
    'フォームのLotNoとシートのLotNoが違う場合シート再発行をしてクリア
  If ThisWorkbook.Sheets("WorkerTime" & "_" & Lname).Cells(1, 1) <> UF1.Lot Then Call AddWS("WorkerTime" & "_" & Lname)
End If

Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("WorkerTime" & "_" & Lname)
Dim i As Long, x As Long
'LOT,ﾗﾍﾞﾙ
sh.Cells(1, 1) = UF1.Lot
sh.Cells(1, 2) = Format$((UF1.StartTime), "yyyy/mm/dd hh:mm:ss") 'start
Dim LField As Variant, FieldRan As Range
Select Case True
Case UF1.極厚: LField = Array("厚1", "厚2", "厚3", "厚4", "幅1", "幅2", "幅3", "幅4", "幅5", "幅6", "丈1", "丈2")
Case Else:     LField = Array("厚1", "厚2", "厚3", "厚4", "幅1", "幅2", "丈1", "丈2", "対角A", "対角B")
End Select
Set FieldRan = sh.Range(sh.Cells(2, 1), sh.Cells(2, UBound(LField) + 1)) '0配列
FieldRan = LField
'座標決定
Select Case Left(Position, 1) '行決定
Case "S": i = 1
Case "U": i = 2
Case "C": i = 3
Case "D": i = 4
Case "E": i = 5
Case "F": i = 6
End Select
Select Case Len(Position) '座標列
Case 2: x = Right(Position, 1) 'S1
Case 3: x = Mid(Position, 2, 2) 'S10
End Select
'作業者,時間転写
Dim Worker1: Worker1 = Split(UF1.T1, "_") '複数人の登録があった場合後の登録者を記録する
Dim Worker2: Worker2 = Split(UF1.T2, "_")
Dim Worker3: Worker3 = Split(UF1.T3, "_")

'F追加で座標修正_2024.605
If Temp = "1" Then 'Tempのﾃﾞﾌｫﾙﾄ値　転写
  Select Case WorkeDayMax '登録時間が一番遅い_2021.02.23:WorkeDayMaxに修正　0時超えに関してはWorkeDayMax使用
  Case "1"
    If IsEmpty(Worker1) = False Then sh.Cells(i + 2, x) = Worker1(UBound(Worker1)) '最後の登録者
    sh.Cells(i + 8, x) = EachT
    sh.Cells(i + 14, x) = UF1.PakeD
  Case "2"
    If IsEmpty(Worker2) = False Then sh.Cells(i + 2, x) = Worker2(UBound(Worker2))
    sh.Cells(i + 8, x) = EachT
    sh.Cells(i + 14, x) = UF1.PakeD
  Case "3"
    If IsEmpty(Worker3) = False Then sh.Cells(i + 2, x) = Worker3(UBound(Worker3))
    sh.Cells(i + 8, x) = EachT
    sh.Cells(i + 14, x) = UF1.PakeD
  End Select
Else 'Tempが空白の時作業者、測定時間をｸﾘｱ
  sh.Cells(i + 2, x) = ""
  sh.Cells(i + 8, x) = ""
  sh.Cells(i + 14, x) = ""
End If
With sh.Cells
  .Font.Name = "メイリオ"
  .Font.Size = 10
  .ColumnWidth = 10.25
  .EntireColumn.AutoFit
End With
End Sub

'Time_Posting の検査員書き込み部分（sh.Cells(i + 2, x)）のみを再現　WorkerTime シートが存在しない場合は何もしない
Public Sub ApplyWorkerToTimeSheet(Position As String, WorkerName As String, Optional timeStamp As String = "")
    If Position = "" Or WorkerName = "" Then Exit Sub
    Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
    If WsCheck("WorkerTime" & "_" & Lname) = False Then Call AddWS("WorkerTime" & "_" & Lname)

    Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("WorkerTime" & "_" & Lname)
    Dim i As Long, x As Long

    Select Case Left(Position, 1)
    Case "S": i = 1
    Case "U": i = 2
    Case "C": i = 3
    Case "D": i = 4
    Case "E": i = 5
    Case "F": i = 6
    Case Else: Exit Sub
    End Select

    Select Case Len(Position)
    Case 2: x = val(Right(Position, 1))
    Case 3: x = val(Mid(Position, 2, 2))
    Case Else: Exit Sub
    End Select

    If x = 0 Then Exit Sub

    sh.Cells(i + 2, x) = WorkerName              ' Time_Posting と同オフセット
    If timeStamp <> "" Then sh.Cells(i + 8, x) = timeStamp  ' 測定時刻
End Sub
'Split_Z(n, a, b)
'  a → CB{n}1.Value / CB{n}2.Value / Z{n}.Value  ← 表示中の面
'  b → CB{n}1.Tag  / CB{n}2.Tag  / Z{n}.Tag     ← 隠れている面
Public Sub Split_Z_0(n As String, Optional b As Variant = "") '空の時****AなしBありに対応_2023.6.15
UF1("CB" & n & "1") = ""
UF1("CB" & n & "2") = ""
UF1("Z" & n) = ""
If IsEmptyArray(b) Then Exit Sub
Call Split_Z_Tag(n, b) 'Split_Zで保持する面とは逆の面をとりこむ
End Sub
Public Sub Split_G_0(n As String, Optional b As Variant = "") '空の時****AなしBありに対応_2023.6.15
UF1("CB" & n & "3") = ""
UF1("CB" & n & "4") = ""
UF1("G" & n) = ""
If IsEmptyArray(b) Then Exit Sub
Call Split_G_Tag(n, b) 'Split_Gで保持する面とは逆の面をとりこむ
End Sub
Public Sub Split_Z(n As String, a As Variant, Optional b As Variant = "") 'Splitで分けたデータを振り分け
If IsEmptyArray(a) Then GoTo SkipB
UF1("CB" & n & "1") = a(LBound(a))
UF1("CB" & n & "2") = a(1)
UF1("Z" & n) = a(UBound(a))
SkipB:
If IsEmptyArray(b) Then Exit Sub
Call Split_Z_Tag(n, b) 'Split_Zで保持する面とは逆の面をとりこむ
End Sub
Public Sub Split_G(n As String, a As Variant, Optional b As Variant = "") 'Splitで分けたデータを振り分け
If IsEmptyArray(a) Then GoTo SkipB
UF1("CB" & n & "3") = a(LBound(a))
UF1("CB" & n & "4") = a(1)
UF1("G" & n) = a(UBound(a))
SkipB:
If IsEmptyArray(b) Then Exit Sub
Call Split_G_Tag(n, b) 'Split_Gで保持する面とは逆の面をとりこむ
End Sub

Public Sub Split_Z_Tag(n As String, a As Variant) 'Split_Zで保持する面とは逆の面をとりこむ
UF1("CB" & n & "1").Tag = a(LBound(a))
UF1("CB" & n & "2").Tag = a(1)
UF1("Z" & n).Tag = a(UBound(a))
End Sub
Public Sub Split_G_Tag(n As String, a As Variant) 'Split_Gで保持する面とは逆の面をとりこむ
UF1("CB" & n & "3").Tag = a(LBound(a))
UF1("CB" & n & "4").Tag = a(1)
UF1("G" & n).Tag = a(UBound(a))
End Sub

Public Sub Split_Z_Tag_0(n As String) 'tagデータを消す
UF1("CB" & n & "1").Tag = ""
UF1("CB" & n & "2").Tag = ""
UF1("CB" & n & "3").Tag = ""
UF1("CB" & n & "4").Tag = ""
UF1("G" & n).Tag = ""
UF1("Z" & n).Tag = ""
End Sub
Public Sub Split_G_Tag_0(n As String) 'tagデータを消す
UF1("CB" & n & "3").Tag = ""
UF1("CB" & n & "4").Tag = ""
UF1("G" & n).Tag = ""
End Sub

Public Sub OrdNoOpen(LotNo As String) 'ｵｰﾀﾞｰNOをコンボボックスに入れる_2022.01.27
'---------------------Lot情報履歴検索
If WsCheck("Lot情報履歴") = False Then Exit Sub  '-------------------------ｼｰﾄなし処理停止
If LotNo = "" Then Exit Sub
Dim HisSH As Worksheet:   Set HisSH = ThisWorkbook.Sheets("Lot情報履歴")
Dim HisRan As Range:      Set HisRan = HisSH.Range(HisSH.Cells(1, 1), HisSH.Cells(UR(HisSH, 1), UC(HisSH, 1)))
Dim Hist(): Hist = HisRan
Dim i As Long, a As Long, n As Long: n = 0
Dim R_Ord As Long, R_Packing As Long, R_Lot As Long, R_Hik As Long
For i = 1 To UBound(Hist, 2) 'ﾌｨｰﾙﾄﾞ名
  Select Case Hist(1, i)
  Case "ﾛｯﾄNo":        R_Lot = i
  Case "ｵｰﾀﾞｰNo":      R_Ord = i
  Case "引当NO":       R_Hik = i
  Case "包装仕様NO":   R_Packing = i
  End Select
Next i
Dim FL As Boolean: FL = False
Dim TempHiki()
For i = LBound(Hist, 1) To UBound(Hist, 1) '1次元ﾁｪｯｸ
  If Hist(i, R_Lot) = LotNo Then a = a + 1
Next i

If a = 0 Then Exit Sub

ReDim TempHiki(1 To a, 1 To 3)
For i = LBound(Hist, 1) To UBound(Hist, 1)
  If Hist(i, R_Lot) = LotNo Then '一致の時
    FL = True
    n = n + 1
    TempHiki(n, 1) = Hist(i, R_Ord)
    TempHiki(n, 2) = Hist(i, R_Packing)
    TempHiki(n, 3) = (Hist(i, R_Hik))
  End If
Next i

Call merge_sort2(TempHiki, 3) '配列内の並び替え_2022.02.01***引当No基準ｿｰﾄ
'Stop
'コンボボックス反映
UF1.OrdNoCB.clear
UF1.PackingNo = ""
UF1.PackingForm = ""
UF1.PackingComment = ""
For i = 1 To UBound(TempHiki, 1)
  If Not TempHiki(i, 1) = "" Then
    UF1.OrdNoCB.AddItem TempHiki(i, 1)
  End If
Next i
End Sub
Public Sub Load_Text(Optional n As String = "") 'n=""でﾌﾟﾛｸﾞﾚｽﾊﾞｰ非表示
'2022.06*****耳付き丈3,4追加分修正
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("作業保存" & "_" & Lname) = False Then Exit Sub '-------------------------ｼｰﾄなし処理停止
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("作業保存" & "_" & Lname)
If n = "ReLoad" Then GoTo Next1
If n = "表示" Then GoTo Line1 Else GoTo LineOut
Line1:
  UFProgress.Show vbModeless
  UFProgress.CommandButton1.Caption = "SaveDate_Loading...."
  Call ProgressBar_count(0, 620)
  DoEvents
LineOut:
'If SH.Cells(1, 1) = "" Then
'   Unload UFProgress
'   Exit Sub '----------------------------------中身なし
'End If
Next1:

UF1.Loading.Tag = "Loading" 'AccessLog停止_2023.8.25

UF1.FlagLabel = 0 'ループ停止
Dim 検査(), TemRange As Range
Dim データ(), TextBoxArray(), DateRange As Range, TextBoxRange As Range
Dim i As Long, a As Long
Dim EndCol As Long: EndCol = UC(sh, 1)

Set TemRange = sh.Range(sh.Cells(1, 1), sh.Cells(1, 28)) 'EndCol使用時未検査でﾏｲｸﾛNOがないときｴﾗｰ 2023.04.18 29ｺ
検査() = TemRange

Dim EndRow As Long: EndRow = UR(sh, 1) '途中検査が抜けている時エラー

Set DateRange = sh.Range(sh.Cells(2, 1), sh.Cells(9, 20)) 'ﾃﾞｰﾀ 2023.5.12_Row+1***(検番)
データ() = DateRange

'歪の公差***2024.2.6
Set TextBoxRange = sh.Range(sh.Cells(1, 33), sh.Cells(1, 39))
TextBoxArray() = TextBoxRange

'2023.05.19_追加*外観_BH_検入高さ
With UF1
  Dim Visual As Range:  Set Visual = sh.Range(sh.Cells(10, 1), sh.Cells(15, 2)) '外観
  Dim BHR As Range:     Set BHR = sh.Range(sh.Cells(10, 3), sh.Cells(14, 3)) 'BH
  
  Dim VI(): VI() = Visual
  Dim bh(): bh() = BHR

  UF1.S_UpVisual = VI(1, 1)
  UF1.S_UnderVisual = VI(1, 2)
  UF1.SBH = bh(1, 1)
  UF1.U_UpVisual = VI(2, 1)
  UF1.U_UnderVisual = VI(2, 2)
  UF1.UBH = bh(2, 1)
  UF1.C_UpVisual = VI(3, 1)
  UF1.C_UnderVisual = VI(3, 2)
  UF1.CBH = bh(3, 1)
  UF1.D_UpVisual = VI(4, 1)
  UF1.D_UnderVisual = VI(4, 2)
  UF1.DBH = bh(4, 1)
  UF1.E_UpVisual = VI(5, 1)
  UF1.E_UnderVisual = VI(5, 2)
  UF1.EBH = bh(5, 1)
  UF1.F_UpVisual = VI(6, 1)
  UF1.F_UnderVisual = VI(6, 2)

  '2023.05.19_追加検入高さ
  Dim KenHR As Range:  Set KenHR = sh.Range(sh.Cells(9, 3), sh.Cells(9, 9)) '検入高さ
  Dim KenH(): KenH() = KenHR
  For a = 1 To 6
    UF1("KenHBox" & a) = Format(KenH(1, a), "0.0")
  Next a
    UF1("KenHBox" & 7) = KenH(1, 7)
    
  '検入山数選択
  Dim Yh As Long
  If sh.Cells(9, 2) <> "" And IsNumeric(sh.Cells(9, 2)) Then '山数
    Yh = val(sh.Cells(9, 2)) - 1 '先頭行無し設定
    UF1.YamaLB.ListIndex = Yh
'    UF1.YamaLB.Selected(Yh) = True'空白時-1でえらーになる
  ElseIf sh.Cells(9, 2) <> "" And IsNumeric(sh.Cells(9, 2)) = False Then '空白でないかつ数字ではない("枚数")
    UF1.YamaLB.ListIndex = UF1.YamaLB.ListCount - 1
  End If
  
  '2024.10.30
'   UF1.KenCount = SH.Cells(1, 18) '検査回数
   UF1.KenCount = sh.Cells(1, 40) '2025.12.17_バッティングが起きていた
  
  '2026.03.05_フリー反り公差
    UF1.HURIS.Text = sh.Cells(1, 41)
    
  '2026.07.13_ﾗﾃﾗﾙﾎﾞｰ欄表示状態(strShijiLate経路)復元用フラグ
    UF1.LateralText.Tag = sh.Cells(1, 42)
    
  '2023.6.8_追加
    UF1.高さ計算式.Caption = sh.Cells(9, 10) '計算式
End With

'2023.5.22_追加
UF1.KenH.Caption = sh.Cells(16, 1) '前工程実績
'2023.6.8_追加
UF1.LastHot.Caption = sh.Cells(17, 1) '最終熱処理工程
'2023.6.9_追加
UF1.QF.Caption = sh.Cells(17, 2) 'QFの有無

With UF1
  .Lot = 検査(1, 1)
  .ATU = Format(検査(1, 2), "0.000")
  .Hab = Format(検査(1, 3), "0.0")
  .Tak = Format(検査(1, 4), "0.0")
  .Am = Format(検査(1, 5), "0.000")
  .Ap = Format(検査(1, 6), "0.000")
  .Hm = Format(検査(1, 7), "0.0")
  .Hp = Format(検査(1, 8), "0.0")
  .Tm = Format(検査(1, 9), "0.0")
  .Tp = Format(検査(1, 10), "0.0")
  .AmK = Format(検査(1, 11), "0.000")
  .ApK = Format(検査(1, 12), "0.000")
  .HmK = Format(検査(1, 13), "0.0")
  .HpK = Format(検査(1, 14), "0.0")
  .TmK = Format(検査(1, 15), "0.0")
  .TpK = Format(検査(1, 16), "0.0")
  .Zai = 検査(1, 17)
  .Tyu = 検査(1, 18)
  .Youto = 検査(1, 19) '2021.01.26
  .TAIK = 検査(1, 20)
  .Machin = 検査(1, 21)
  .Lateral = 検査(1, 23) '2021.07.13
  .TaikText.Text = Format(検査(1, 24), "0.0") '2021.07.13
  .LateralText.Text = Format(検査(1, 25), "0.0") '2021.07.13'等級で見るときは下1桁***2022.09.20
  
  .flatnes.Text = 検査(1, 26)
  
  .Sumisei.Caption = 検査(1, 27)    '---------------ｽﾐｾｲﾌﾗｸﾞ2021.07.13
  
  '歪の公差
  .MG = Format(TextBoxArray(1, 1), "0.0") 'ｹﾞｰｼﾞ
  .HaHiz = Format(TextBoxArray(1, 2), "0.0") '幅
  .TaHiz = Format(TextBoxArray(1, 3), "0.0") '丈
  .Naka = TextBoxArray(1, 4) '中
  .Mimi = TextBoxArray(1, 5) '耳
  If IsNumeric(.Naka) Then .Naka = Format((.Naka), "0.0")
  If IsNumeric(.Mimi) Then .Mimi = Format((.Mimi), "0.0")
  .DistortionFrame.Caption = TextBoxArray(1, 6)
  .HizCome.Caption = TextBoxArray(1, 7)
  
  '<<<<<<<<<<<<<<<<<<<<<<<<<<<<特別指示のある場合変色処理
  '2023.9.5_日本車両を追加
  If 検査(1, 27) <> "" Then '= "白銅" Or 検査(1, 27) = "日本車両" Or 検査(1, 27) = "滑川" Then 'HistorySaveで処理命名****2022.08.26
    .LateralText.BackStyle = 1
    .LateralText.BackColor = RGB(255, 153, 0)
'    .LateralText.text = Format(検査(1, 25), "0") '白銅で見るときは小数点なし***2022.09.20****Q-S-QA-DIM-179_板材曲がり(ラテラルボー)検査基準にて改訂
  End If
  If 検査(1, 28) <> "" Then '追記公差フラグの有無
    .Additional.Caption = 検査(1, 28) '---------------追加公差ﾌﾗｸﾞ2022.07.01
    Select Case Mid(検査(1, 28), 1, 3)
    Case "対角ﾗ" '両方
      .TaikText.BackStyle = 1
      .TaikText.BackColor = RGB(51, 204, 204)
      .LateralText.BackStyle = 1
      .LateralText.BackColor = RGB(51, 204, 204)
    Case "対角追" '対角のみ
      .TaikText.BackStyle = 1
      .TaikText.BackColor = RGB(51, 204, 204)
    Case "ﾗﾃﾗ" 'ラテのみ
      .LateralText.BackStyle = 1
      .LateralText.BackColor = RGB(51, 204, 204)
    End Select
  End If
  
  Select Case n '2021.01.13_追加 ﾘｻｲｽﾞの有無を明確にする　一度空白でﾎﾟｽﾄされると永遠空白で読み込んでしまう
  Case "ReLoad": .ReSize.Tag = "無"
  Case Else:     .ReSize.Tag = 検査(1, 22)
  End Select
  .ResizText.Caption = .ReSize.Tag
For i = 1 To 20 '---------------------------------------------検査結果格納 桁揃え**************2022.06_耳付き丈3,4追加により+2
  For a = 1 To 6
    Select Case i
    Case 1 To 4
      UF1("S" & i) = Format(val(データ(1, i)), "0.000")
      UF1("U" & i) = Format(val(データ(2, i)), "0.000")
      UF1("C" & i) = Format(val(データ(3, i)), "0.000")
      UF1("D" & i) = Format(val(データ(4, i)), "0.000")
      UF1("E" & i) = Format(val(データ(5, i)), "0.000")
      UF1("F" & i) = Format(val(データ(6, i)), "0.000")
    Case 5 To 6
      '幅
      '2023.5.12_表示桁数品証確認
      If 検査(1, 19) = "H494" And 検査(1, 3) < 300# Then    '用途コードH494に限り幅測定をノギスで行う
        Select Case val(検査(1, 3)) '幅寸法が範囲内の時
        Case 1# To 299#
          UF1("S" & i) = Format(val(データ(1, i)), "0.00")
          UF1("U" & i) = Format(val(データ(2, i)), "0.00")
          UF1("C" & i) = Format(val(データ(3, i)), "0.00")
          UF1("D" & i) = Format(val(データ(4, i)), "0.00")
          UF1("E" & i) = Format(val(データ(5, i)), "0.00")
          UF1("F" & i) = Format(val(データ(6, i)), "0.00")
        Case Else
          UF1("S" & i) = Format(val(データ(1, i)), "0.0")
          UF1("U" & i) = Format(val(データ(2, i)), "0.0")
          UF1("C" & i) = Format(val(データ(3, i)), "0.0")
          UF1("D" & i) = Format(val(データ(4, i)), "0.0")
          UF1("E" & i) = Format(val(データ(5, i)), "0.0")
          UF1("F" & i) = Format(val(データ(6, i)), "0.0")
        End Select
      Else
        UF1("S" & i) = Format(val(データ(1, i)), "0.0")
        UF1("U" & i) = Format(val(データ(2, i)), "0.0")
        UF1("C" & i) = Format(val(データ(3, i)), "0.0")
        UF1("D" & i) = Format(val(データ(4, i)), "0.0")
        UF1("E" & i) = Format(val(データ(5, i)), "0.0")
        UF1("F" & i) = Format(val(データ(6, i)), "0.0")
      End If
    Case 7 To 10
      '丈
      UF1("S" & i) = Format(val(データ(1, i)), "0.0")
      UF1("U" & i) = Format(val(データ(2, i)), "0.0")
      UF1("C" & i) = Format(val(データ(3, i)), "0.0")
      UF1("D" & i) = Format(val(データ(4, i)), "0.0")
      UF1("E" & i) = Format(val(データ(5, i)), "0.0")
      UF1("F" & i) = Format(val(データ(6, i)), "0.0")
    Case 11 To 14
      '丈
      UF1("S" & i) = Format(val(データ(1, i)), "0.0")
      UF1("U" & i) = Format(val(データ(2, i)), "0.0")
      UF1("C" & i) = Format(val(データ(3, i)), "0.0")
      UF1("D" & i) = Format(val(データ(4, i)), "0.0")
      UF1("E" & i) = Format(val(データ(5, i)), "0.0")
    Case 15 To 16
        '歪
    Case 17 '対角線差
      .DiagS = Format(val(データ(1, i)), "0.0")
      .DiagU = Format(val(データ(2, i)), "0.0")
      .DiagC = Format(val(データ(3, i)), "0.0")
      .DiagD = Format(val(データ(4, i)), "0.0")
      .DiagE = Format(val(データ(5, i)), "0.0")
      .DiagF = Format(val(データ(6, i)), "0.0")
    Case 18 '対角公差**********2023.11.24_対角公差から直角度に変更
'      If 検査(1, 19) = "R115" Then
        Select Case a
        Case 1
          If データ(1, i) <> "" And IsNumeric(データ(1, i)) Then
            .DiagSKou = Format(val(データ(1, i)), "0.0") '測定値があるときのみ
          ElseIf データ(1, i) = "" Then
            .DiagSKou = ""
          End If
        Case 2
          If データ(2, i) <> "" And IsNumeric(データ(2, i)) Then
            .DiagUKou = Format(val(データ(2, i)), "0.0")
          ElseIf データ(2, i) = "" Then
            .DiagUKou = ""
          End If
        Case 3
          If データ(3, i) <> "" And IsNumeric(データ(3, i)) Then
            .DiagCKou = Format(val(データ(3, i)), "0.0")
          ElseIf データ(3, i) = "" Then
            .DiagCKou = ""
          End If
        Case 4
          If データ(4, i) <> "" And IsNumeric(データ(4, i)) Then
            .DiagDKou = Format(val(データ(4, i)), "0.0")
          ElseIf データ(4, i) = "" Then
            .DiagDKou = ""
          End If
        Case 5
          If データ(5, i) <> "" And IsNumeric(データ(5, i)) Then
            .DiagEKou = Format(val(データ(5, i)), "0.0")
          ElseIf データ(5, i) = "" Then
            .DiagEKou = ""
          End If
        Case 6
          If データ(6, i) <> "" And IsNumeric(データ(6, i)) Then
            .DiagFKou = Format(val(データ(6, i)), "0.0")
          ElseIf データ(6, i) = "" Then
            .DiagFKou = ""
          End If
        End Select
'      End If
    Case 19 '公差外判定ﾌﾗｸﾞ
      .NGS.Caption = データ(1, i)
      .NGU.Caption = データ(2, i)
      .NGC.Caption = データ(3, i)
      .NGD.Caption = データ(4, i)
      .NGE.Caption = データ(5, i)
       .NGF.Caption = データ(6, i)
    Case 20 'ﾗﾃﾗﾙﾎﾞｰ
      .LateralS = Format(データ(1, i), "0.0")
      .LateralU = Format(データ(2, i), "0.0")
      .LateralC = Format(データ(3, i), "0.0")
      .LateralD = Format(データ(4, i), "0.0")
      .LateralE = Format(データ(5, i), "0.0")
      .LateralF = Format(データ(6, i), "0.0")
    End Select
  Next a
Next i

'歪用_2023.6.14
Dim A_range As Range: Set A_range = sh.Range(sh.Cells(2, 15), sh.Cells(7, 16)) 'A面
Dim B_range As Range: Set B_range = sh.Range(sh.Cells(8, 15), sh.Cells(13, 16)) 'B面
Call LosdStrain("A", A_range, B_range) 'tagに数値落とし込み

 '0→""書き換え
For i = 1 To 14
  If UF1("S" & i) = 0 Then
    UF1("S" & i) = ""
    UF1("S" & i).BackColor = &H80000005 '白　ﾃﾞﾌｫﾙﾄ色
  End If
  If UF1("U" & i) = 0 Then
    UF1("U" & i) = ""
    UF1("U" & i).BackColor = &H80000005
  End If
  If UF1("C" & i) = 0 Then
    UF1("C" & i) = ""
    UF1("C" & i).BackColor = &H80000005
  End If
  If UF1("D" & i) = 0 Then
    UF1("D" & i) = ""
    UF1("D" & i).BackColor = &H80000005
  End If
  If UF1("E" & i) = 0 Then
    UF1("E" & i) = ""
    UF1("E" & i).BackColor = &H80000005
  End If
Next i
For i = 1 To 10
  If UF1("F" & i) = 0 Then
    UF1("F" & i) = ""
    UF1("F" & i).BackColor = &H80000005
  End If
Next i
'対角測定がある場合は空白にしない_2023.12.05
 If .S10.Caption = "" And .DiagS = 0 Then .DiagS = ""
 If .U10.Caption = "" And .DiagU = 0 Then .DiagU = ""
 If .C10.Caption = "" And .DiagC = 0 Then .DiagC = ""
 If .D10.Caption = "" And .DiagD = 0 Then .DiagD = ""
 If .E10.Caption = "" And .DiagE = 0 Then .DiagE = ""
 If .F10.Caption = "" And .DiagF = 0 Then .DiagF = ""
End With

Call Atu_OPselect(val(sh.Cells(1, 2))) '描写リセットあり

Dim SetTempRange As Range: Set SetTempRange = sh.Range(sh.Cells(1, 29), sh.Cells(1, 32))
Dim KenItem():             KenItem = SetTempRange

Dim TempAtSl As String:   TempAtSl = KenItem(1, 1) '作業選択
Dim Two_Four As String:   Two_Four = KenItem(1, 2) '厚の2か所検査、4か所検査
Dim AboveBelow As String: AboveBelow = KenItem(1, 3) '25枚以上以下
Dim TempYama As String:   TempYama = KenItem(1, 4) '山数
Dim KenNo As String:      KenNo = Left(データ(8, 1), 1)  '検査No'頭1桁

'2022.07*追記公差ﾌﾗｸﾞ追加分+1ずれる
UF1.Yama.Text = TempYama '山数**************先に登録しないとLSelectで正しく描写できない_2022.04.18
Dim LastCol As Long: LastCol = 32 'UC(SH, 1)'2021.12.15_薄板検入山数追加***2023.4.18ﾌﾗｯﾄﾈｽ等級追加

Select Case TempAtSl '薄板時測定カ所で最終列が変わる
Case "厚板"
  UF1.厚板 = True
  Select Case KenNo
  Case "H", "T"
    UF1.厚板.Caption = KenNo
  End Select
Case "薄板"
  UF1.薄板 = True
  Select Case KenNo
  Case "L", "N", "K", "T", "S", "F"
    UF1.薄板.Caption = KenNo
  End Select
Case "少枚数": UF1.少枚数 = True
Case "極厚":   UF1.極厚 = True
Case "スラブ": UF1.スラブ = True
End Select

'2024.11.15_耳付きスラブの際はラテラルボーは不要
If UF1.Sumisei.Caption <> "" Then
  If UF1.極厚 = False And UF1.スラブ = False Then Call LSelect("住精")  '2021.07.13*******2022.8.17*****ﾗﾃﾗﾙﾎﾞｰ測定の納入先が追加される為"住友精密"から<>に変更
End If

'2026.07.13_ﾗﾃﾗﾙﾎﾞｰ欄復元_strShijiLate経路(Tagフラグ) Or DbReturnItem経路(Sumisei.Caption)
If UF1.LateralText.Tag = "有" Or UF1.Sumisei.Caption <> "" Then
  Call LSelectLateral(KenNo)
End If

If UF1.Yama.Visible Then
  If UF1.Yama.Text = "4" Then
    Call Six_Display("on") '2024.5.23_4山での表示
  Else
    Call Six_Display("off")
  End If
Else
  Call Six_Display("off")
End If

'LSelect("住精")より後に実行
Select Case Two_Four 'ｽﾐｾｲﾌﾗｸﾞがないとき空白になるため指定値
Case "厚2ヵ所"
 Call AtCountSelect(2)
Case "厚4ヵ所"
 Call AtCountSelect(4)
End Select

Select Case AboveBelow
Case "検入≧26":  Call InspectionSelect("非表示")
Case "検入≦25":  Call InspectionSelect("表示")
End Select

If UF1.スラブ Then Call SlabSet '2024.6.6_追加

UF1.KenNo.Caption = データ(8, 1) '検番

'2023.4.24_消した座標と数の保存*固定座標　検査結果の下
UF1.DeleteP.Caption = データ(7, 1)
UF1.DeleteC.Caption = データ(7, 2)

'少枚数用の枚数を確定させる_2023.8.24
'2024.10.30_検査回数と検入を分ける
Dim th As Long
If UF1.ATU <> "" Then
  If UF1.KenHBox7 = "" Then
    th = Format(Round(CalculateSheets(UF1, UF1.ATU), 1), "0")
  ElseIf UF1.KenCount <> "" Then
    th = UF1.KenCount
  End If
  If th >= 5 Then th = 5 '少枚数は5枚まで
End If

Dim Kno As String:    Kno = UF1.厚板.Caption
Dim flag As Boolean:  If UF1.少枚数 Then flag = True
If Kno = "H" And flag Then Call SmallNumberCount(th)

'検入山　高さ
'検入山　高さ
If Not (UF1.ATU = "") Then Call DrawListBoxContent(UF1, UF1.ATU) '板厚表示の後に実行

UF1.Repaint '2020.12.24_追加

'SH="作業保存"
If WsCheck("Lot情報履歴") = False Then Exit Sub
Dim HistorySH As Worksheet: Set HistorySH = ThisWorkbook.Sheets("Lot情報履歴")
Dim LR As Long:             LR = UR(HistorySH, 1)
Dim LC As Long:             LC = UC(HistorySH, 1)
Dim HistoryR As Range:      Set HistoryR = HistorySH.Range(HistorySH.Cells(2, 1), HistorySH.Cells(LR, LC))
Dim TempHiki():             TempHiki = HistoryR
Dim x As Long, matchFlag As Boolean: matchFlag = False
For i = 1 To UBound(TempHiki, 1)
  If TempHiki(i, 3) = UF1.Lot Then 'LOTNO一致
    x = i
    matchFlag = True
    Exit For
  End If
Next i

If matchFlag = False Then
   Call ProgressBar_count(620, 1000)
   DoEvents
   Unload UFProgress
   Exit Sub
End If
If TempHiki(x, 18) = 0 Then Call VC_Clear
If TempHiki(x, 18) = 1 Then Call VC_Picture  'VC画像

'2024.8.5
Dim Retemp As Variant
If UFsizeWidth <> 0 Then Retemp = InitialSetting(UF1, "サイズ保存")


If WsCheck("WorkerTime" & "_" & Lname) Then UF1.StartTime = Format$((ThisWorkbook.Sheets("WorkerTime" & "_" & Lname).Cells(1, 2)), "yyyy/mm/dd hh:mm:ss")
If WsCheck("異常ロット用" & "_" & Lname) Then If ThisWorkbook.Sheets("異常ロット用" & "_" & Lname).Cells(1, 1) = UF1.Lot Then UF1.IzyouFl.BackColor = RGB(255, 0, 0) '2021.01.21_異常ロット管理明確

Dim QfCheck As String

Select Case True '2021.11.24_導電率チェック
Case Conductiv_Check
  QfCheck = Conductiv_Coolingtime_Check(UF1.Lot) '変数に修正_2024.5.13
  UF1.Conductiv.BackColor = RGB(255, 0, 0)
  If IsNumeric(QfCheck) Then '数字の時**検査可
    UF1.ConductivL.Visible = True
    UF1.ConductivText.Caption = QfCheck
  ElseIf QfCheck <> "" Then '何か文字列
    UF1.ConductivL.Visible = True
    UF1.ConductivText.Caption = QfCheck
    '検査不可フラグはLoad時はスルーする_2022.07
  End If
Case Else
  UF1.ConductivL.Visible = False
  UF1.Conductiv.BackColor = &H8000000F
End Select

UF1.FlagLabel = 1
UF1.Loading.Tag = "" ' 'AccessLog再開
Call ProgressBar_count(620, 1000)
DoEvents
Unload UFProgress
End Sub

'ログから呼び出し
'1 LotNo
'2～15  測定値 S/U/C/D/E/F × 14箇所
'16  直角度 (DiagKou)
'17  ラテラルボー
'18 Print 20; Worker / logName / 理由
'21  UpVisual (上面外観)
'22  UnderVisual (下面外観)
'歪　作業保存シート col 15-16 に CB1_CB2_Z 形式で直接書く　LosdStrain をそのまま呼べる
'2022.06_専用シートに転記するよう修正'読み取ったデータをFormReflectionでフォームに反映させる
'2026.05.15_ReadFile（全文字列）→ ADODB.Stream の ReadText(-2)（1行ずつ）重すぎたので改修
Public Sub AccesslogLoad(Lot As String, Worke As String, logName As String, Reas As String)
If Not (Lot <> "" And Worke <> "" And logName <> "" And Reas <> "") Then
    MsgBox "未記入TextBoxがあります。すべて記入してください。", vbExclamation
    Exit Sub
End If

Dim FilePath As String: FilePath = ThisWorkbook.path & "\" & logName
If ReadCheck(FilePath, 1) = False Then
    MsgBox FilePath & "が開かれているか、存在しません。" & vbCrLf & _
           "ファイルを確認してください。", vbCritical
    Exit Sub
End If

'── ADODB.Stream で UTF-8 全取得 ────────────────────────────────────
Call ShowProgress("ログファイル読込中...")

Dim st          As Object: Set st = CreateObject("ADODB.Stream")
st.Type = 2
st.Charset = "Shift_JIS"   ' UTF-8 → Shift_JIS
st.Open
st.LoadFromFile FilePath

Call ShowProgress("LotNo検索中...")
'── データ格納変数 ───────────────────────────────────────────────────
Dim i As Long
Dim S(1 To 14) As String, U(1 To 14) As String, C(1 To 14) As String
Dim D(1 To 14) As String, E(1 To 14) As String, F(1 To 10) As String
Dim Slate  As String, Ulate  As String, Clate  As String
Dim Dlate  As String, Elate  As String, Flate  As String
Dim Sangle As String, Uangle As String, Cangle As String
Dim Dangle As String, Eangle As String, Fangle As String

'ApplyVisualStrainFromLog / ApplyWorkerFromLog 用 Lot範囲行
Dim lotLines() As String
Dim lotCount   As Long: lotCount = 0
ReDim lotLines(0 To 999)

'── 1パス処理ループ ──────────────────────────────────────────────────
Dim a        As Long
Dim lineStr  As String
Dim buf      As Variant
Dim Hitcount As Long
Dim inLot    As Boolean: inLot = False
Dim found    As Boolean: found = False

Do While Not st.EOS
    lineStr = st.ReadText(-2)          ' -2 = adReadLine（1行ずつ）
    lineStr = Replace(lineStr, vbCr, "")
    
    If Not inLot Then
        '── LotNo を探す ────────────────────────────────────────────
        If InStr(lineStr, Lot) <> 0 Then
            inLot = True
            found = True
            If lotCount > UBound(lotLines) Then ReDim Preserve lotLines(0 To UBound(lotLines) + 500)
            lotLines(lotCount) = lineStr    '行[0] = Lot登録行
            lotCount = lotCount + 1
        End If

    Else
        '── 次の Lot登録 → 即終了（後半を処理しない） ───────────────
        If InStr(lineStr, "Lot登録") <> 0 Then Exit Do

        '行収集
        If lotCount > UBound(lotLines) Then ReDim Preserve lotLines(0 To UBound(lotLines) + 500)
        lotLines(lotCount) = lineStr
        lotCount = lotCount + 1

        If lineStr = "" Then GoTo ContinueA

        '── フォーマット判別 ─────────────────────────────────────────
        If InStr(lineStr, "|") > 0 Then
            buf = Split(lineStr, "|")
            Hitcount = 1
        Else
            buf = Split(lineStr, ":")
            Hitcount = UBound(buf) - 3
        End If
        If Hitcount < 0 Or Hitcount > UBound(buf) Then GoTo ContinueA

        '── 測定値抽出（S/U/C/D/E: 1-14） ───────────────────────────
        For i = 1 To 14
            Select Case True
            Case InStr(lineStr, "(S" & i & ")") <> 0: S(i) = buf(Hitcount)
            Case InStr(lineStr, "(U" & i & ")") <> 0: U(i) = buf(Hitcount)
            Case InStr(lineStr, "(C" & i & ")") <> 0: C(i) = buf(Hitcount)
            Case InStr(lineStr, "(D" & i & ")") <> 0: D(i) = buf(Hitcount)
            Case InStr(lineStr, "(E" & i & ")") <> 0: E(i) = buf(Hitcount)
            End Select
        Next i

        '── 測定値抽出（F: 1-10） ────────────────────────────────────
        For i = 1 To 10
            If InStr(lineStr, "(F" & i & ")") <> 0 Then F(i) = buf(Hitcount)
        Next i

        '── ラテラルボー ─────────────────────────────────────────────
        Select Case True
        Case InStr(lineStr, "(LateralS)") <> 0: Slate = buf(Hitcount)
        Case InStr(lineStr, "(LateralU)") <> 0: Ulate = buf(Hitcount)
        Case InStr(lineStr, "(LateralC)") <> 0: Clate = buf(Hitcount)
        Case InStr(lineStr, "(LateralD)") <> 0: Dlate = buf(Hitcount)
        Case InStr(lineStr, "(LateralE)") <> 0: Elate = buf(Hitcount)
        Case InStr(lineStr, "(LateralF)") <> 0: Flate = buf(Hitcount)
        End Select

        '── 直角度 ───────────────────────────────────────────────────
        Select Case True
        Case InStr(lineStr, "(DiagSKou)") <> 0: Sangle = buf(Hitcount)
        Case InStr(lineStr, "(DiagUKou)") <> 0: Uangle = buf(Hitcount)
        Case InStr(lineStr, "(DiagCKou)") <> 0: Cangle = buf(Hitcount)
        Case InStr(lineStr, "(DiagDKou)") <> 0: Dangle = buf(Hitcount)
        Case InStr(lineStr, "(DiagEKou)") <> 0: Eangle = buf(Hitcount)
        Case InStr(lineStr, "(DiagFKou)") <> 0: Fangle = buf(Hitcount)
        End Select
    End If
ContinueA:
Loop

st.Close
Set st = Nothing

If Not found Then
    MsgBox "一致するLotNoがありません", vbCritical
    Exit Sub
End If

ReDim Preserve lotLines(0 To lotCount - 1)

'── Accesslogシートへ書込 ────────────────────────────────────────────
Call ShowProgress("Accesslogシートへ書込中...")

Call AddWS("Accesslog")
Dim sh As Worksheet: Set sh = ShJuge("Accesslog")
Dim n  As Long:      n = UR(sh, 1)

sh.Cells(n, 1) = Lot
sh.Cells(n, 18) = Worke
sh.Cells(n, 19) = logName
sh.Cells(n, 20) = Reas

Dim Sran As Range: Set Sran = sh.Range(sh.Cells(n, 2), sh.Cells(n, 15))
Dim Uran As Range: Set Uran = sh.Range(sh.Cells(n + 1, 2), sh.Cells(n + 1, 15))
Dim Cran As Range: Set Cran = sh.Range(sh.Cells(n + 2, 2), sh.Cells(n + 2, 15))
Dim Dran As Range: Set Dran = sh.Range(sh.Cells(n + 3, 2), sh.Cells(n + 3, 15))
Dim Eran As Range: Set Eran = sh.Range(sh.Cells(n + 4, 2), sh.Cells(n + 4, 15))
Dim Fran As Range: Set Fran = sh.Range(sh.Cells(n + 5, 2), sh.Cells(n + 5, 11))
Sran = S:  Uran = U:  Cran = C:  Dran = D:  Eran = E:  Fran = F

sh.Cells(n, 16) = Sangle: sh.Cells(n, 17) = Slate
sh.Cells(n + 1, 16) = Uangle: sh.Cells(n + 1, 17) = Ulate
sh.Cells(n + 2, 16) = Cangle: sh.Cells(n + 2, 17) = Clate
sh.Cells(n + 3, 16) = Dangle: sh.Cells(n + 3, 17) = Dlate
sh.Cells(n + 4, 16) = Eangle: sh.Cells(n + 4, 17) = Elate
sh.Cells(n + 5, 16) = Fangle: sh.Cells(n + 5, 17) = Flate

Call ShowProgress("準備完了")
Call HideProgress
If MsgBox("結果を反映しますか？", vbYesNo) = vbYes Then
    Call ApplyVisualStrainFromLog(lotLines, 1, lotCount - 1)
    Call FormReflection
    Call FormReflectionEx
    Call ApplyWorkerFromLog(lotLines, 1, lotCount - 1)
End If
End Sub

'位置文字 S/U/C/D/E/F → Accesslogシート行番号
Public Function GetPosRow(posChar As String) As Long
Select Case posChar
Case "S": GetPosRow = 1
Case "U": GetPosRow = 2
Case "C": GetPosRow = 3
Case "D": GetPosRow = 4
Case "E": GetPosRow = 5
Case "F": GetPosRow = 6
Case Else: GetPosRow = 0
End Select
End Function

'ログ行を精査してAccesslogシートの列21～24に書き込む　buf(1) に Visual を含む → 外観 / [ で始まる → 歪
Public Sub ApplyVisualStrainFromLog(LineBuff() As String, SameCount As Long, EndCount As Long)
Call ShowProgress("ログ解析 準備中...")
Call DebugLog("ApplyVisualStrainFromLog 開始 / SameCount=" & SameCount & " EndCount=" & EndCount)

If WsCheck("Accesslog") = False Then
    Call DebugLog("Accesslogシートなし → 終了")
    Exit Sub
End If

Dim accSh  As Worksheet: Set accSh = ThisWorkbook.Sheets("Accesslog")
Dim Lname  As String:    Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim posNames As Variant: posNames = Array("", "S", "U", "C", "D", "E", "F") '1-based用

Application.ScreenUpdating = False
Application.EnableEvents = False

accSh.Range(accSh.Cells(1, 21), accSh.Cells(6, 22)).ClearContents

'歪一時格納 (位置1-6, 面A=1/B=2, slot1-3)
Dim tmpZ(1 To 6, 1 To 2, 1 To 3) As String
Dim tmpG(1 To 6, 1 To 2, 1 To 3) As String

Dim a        As Long
Dim buf      As Variant
Dim ctrlName As String
Dim logVal   As String
Dim face     As String
Dim faceIdx  As Long
Dim ctrlBase As String
Dim posChar  As String
Dim posIdx   As Long
Dim slotNum  As Long

Dim totalLines As Long: totalLines = EndCount - SameCount + 1
Dim prog       As Long

For a = SameCount To EndCount

    '── 進捗表示（200行ごと） ─────────────────────────────────────────
    prog = a - SameCount + 1
    If prog Mod 200 = 1 Then
        Call ShowProgress("ログ解析中... " & prog & " / " & totalLines & " 行")
    End If

    '── 前処理フィルタ（不要行を極力早く弾く） ───────────────────────
    If LineBuff(a) = "" Then GoTo ContinueA
    If InStr(LineBuff(a), "|") = 0 Then GoTo ContinueA
    If InStr(LineBuff(a), "*UFKen") = 0 Then GoTo ContinueA

    buf = Split(LineBuff(a), "|")
    If UBound(buf) < 3 Then GoTo ContinueA

    ctrlName = Trim(buf(1))
    logVal = Trim(buf(2))

    '── 外観: "Visual" を含む ─────────────────────────────
    '   例: *UFKen|C_UnderVisual|ｶﾐｺﾐ跡|L-1
    If InStr(ctrlName, "Visual") > 0 Then
        posChar = Left(ctrlName, 1)
        posIdx = GetPosRow(posChar)
        If posIdx = 0 Then GoTo ContinueA

        If InStr(ctrlName, "UpVisual") > 0 Then
            accSh.Cells(posIdx, 21) = logVal
            Call DebugLog("外観上面 / pos=" & posChar & " val=[" & logVal & "]")

        ElseIf InStr(ctrlName, "UnderVisual") > 0 Then
            accSh.Cells(posIdx, 22) = logVal
            Call DebugLog("外観下面 / pos=" & posChar & " val=[" & logVal & "]")
        End If

    '── 歪: "[A]" or "[B]" で始まる ──────────────────────
    '   例: *UFKen|[A] CBU1|丈|L-1
    '       *UFKen|[B] ZU|2.60|L-1
    ElseIf Left(ctrlName, 1) = "[" Then
        If Len(ctrlName) < 5 Then GoTo ContinueA

        face = Mid(ctrlName, 2, 1)              '"A" or "B"
        faceIdx = IIf(face = "A", 1, 2)
        ctrlBase = Trim(Mid(ctrlName, 5))        '"CBU1" / "ZU" / "GU" ("[A] "=4文字)
        If Len(ctrlBase) < 2 Then GoTo ContinueA

        '── CB コンボ: CB{pos}{1-4} ──────────────────────
        If Left(ctrlBase, 2) = "CB" Then
            posChar = Mid(ctrlBase, 3, 1)
            posIdx = GetPosRow(posChar)
            If posIdx = 0 Then GoTo ContinueA
            slotNum = val(Right(ctrlBase, 1))
            Select Case slotNum
            Case 1: tmpZ(posIdx, faceIdx, 1) = logVal
            Case 2: tmpZ(posIdx, faceIdx, 2) = logVal
            Case 3: tmpG(posIdx, faceIdx, 1) = logVal
            Case 4: tmpG(posIdx, faceIdx, 2) = logVal
            Case Else: GoTo ContinueA
            End Select
            Call DebugLog("歪CB / face=" & face & " " & ctrlBase & " pos=" & posChar & " slot=" & slotNum & " val=[" & logVal & "]")

        '── Z 定盤測定値: Z{pos} ─────────────────────────
        ElseIf Left(ctrlBase, 1) = "Z" Then
            posChar = Right(ctrlBase, 1)
            posIdx = GetPosRow(posChar)
            If posIdx = 0 Then GoTo ContinueA
            tmpZ(posIdx, faceIdx, 3) = logVal
            Call DebugLog("歪Z / face=" & face & " pos=" & posChar & " val=[" & logVal & "]")

        '── G ゲージ測定値: G{pos} ───────────────────────
        ElseIf Left(ctrlBase, 1) = "G" Then
            posChar = Right(ctrlBase, 1)
            posIdx = GetPosRow(posChar)
            If posIdx = 0 Then GoTo ContinueA
            tmpG(posIdx, faceIdx, 3) = logVal
            Call DebugLog("歪G / face=" & face & " pos=" & posChar & " val=[" & logVal & "]")
        End If
        
    '── 歪CB（面プレフィックスなし旧フォーマット）: "CB" で始まる ──
    '   例: *UFKen|CBS1|幅|Test|<氏名>
    '   → 面不明のためA面として扱う
    ElseIf Left(ctrlName, 2) = "CB" Then
        posChar = Mid(ctrlName, 3, 1)              'S/U/C/D/E/F
        posIdx = GetPosRow(posChar)
        If posIdx = 0 Then GoTo ContinueA
        slotNum = val(Right(ctrlName, 1))          '1/2/3/4
        Select Case slotNum
        Case 1: tmpZ(posIdx, 1, 1) = logVal        'A面固定
        Case 2: tmpZ(posIdx, 1, 2) = logVal
        Case 3: tmpG(posIdx, 1, 1) = logVal
        Case 4: tmpG(posIdx, 1, 2) = logVal
        Case Else: GoTo ContinueA
        End Select
        Call DebugLog("歪CB(旧形式A面固定) / " & ctrlName & " slot=" & slotNum & " val=[" & logVal & "]")
    End If
ContinueA:
Next a

'── 作業保存シートへ書込 ─────────────────────────────────────────────
Call ShowProgress("作業保存シートへ書込中...")

If WsCheck("作業保存" & "_" & Lname) = False Then Call AddWS("作業保存" & "_" & Lname)
Dim saveSh As Worksheet: Set saveSh = ThisWorkbook.Sheets("作業保存" & "_" & Lname)

Dim i As Long, zStr As String, gStr As String
For i = 1 To 6
    'A面 (rows 2-7)
    zStr = tmpZ(i, 1, 1) & "_" & tmpZ(i, 1, 2) & "_" & tmpZ(i, 1, 3)
    gStr = tmpG(i, 1, 1) & "_" & tmpG(i, 1, 2) & "_" & tmpG(i, 1, 3)
    If tmpZ(i, 1, 1) <> "" Or tmpZ(i, 1, 3) <> "" Then
        saveSh.Cells(i + 1, 15) = zStr
        Call DebugLog("保存A面Z [" & posNames(i) & "] row=" & (i + 1) & " → " & zStr)
    End If
    If tmpG(i, 1, 1) <> "" Or tmpG(i, 1, 3) <> "" Then
        saveSh.Cells(i + 1, 16) = gStr
        Call DebugLog("保存A面G [" & posNames(i) & "] row=" & (i + 1) & " → " & gStr)
    End If

    'B面 (rows 8-13)
    zStr = tmpZ(i, 2, 1) & "_" & tmpZ(i, 2, 2) & "_" & tmpZ(i, 2, 3)
    gStr = tmpG(i, 2, 1) & "_" & tmpG(i, 2, 2) & "_" & tmpG(i, 2, 3)
    If tmpZ(i, 2, 1) <> "" Or tmpZ(i, 2, 3) <> "" Then
        saveSh.Cells(i + 7, 15) = zStr
        Call DebugLog("保存B面Z [" & posNames(i) & "] row=" & (i + 7) & " → " & zStr)
    End If
    If tmpG(i, 2, 1) <> "" Or tmpG(i, 2, 3) <> "" Then
        saveSh.Cells(i + 7, 16) = gStr
        Call DebugLog("保存B面G [" & posNames(i) & "] row=" & (i + 7) & " → " & gStr)
    End If
Next i

Application.ScreenUpdating = True
Application.EnableEvents = True

Call DebugLog("ApplyVisualStrainFromLog 完了")
Call ShowProgress("解析完了")
Call HideProgress
End Sub

'配列を受け取り、（UBound(buf) >= 7）のみ末尾の検査員名を抽出して ApplyWorkerToTimeSheet に渡す
Public Sub ApplyWorkerFromLog(LineBuff() As String, SameCount As Long, EndCount As Long)
    Dim a As Long, i As Long
    Dim buf As Variant
    Dim wname As String, timeStamp As String
    Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption

    '── WorkerTime シート取得（なければ作成） ────────────────────────────
    If WsCheck("WorkerTime" & "_" & Lname) = False Then Call AddWS("WorkerTime" & "_" & Lname)
    Dim wSh As Worksheet: Set wSh = ThisWorkbook.Sheets("WorkerTime" & "_" & Lname)

    '── Lot登録行 (LineBuff(0)) からLotNo・タイムスタンプを書く ──────────
    wSh.Cells(1, 1) = UF1.Lot
    If LBound(LineBuff) = 0 Then
        Dim lotRegTime As String: lotRegTime = Left(LineBuff(0), 19)
        If lotRegTime <> "" Then
          wSh.Cells(1, 1) = UF1.Lot
          wSh.Cells(1, 2) = lotRegTime
        End If
    End If

    '── 測定値行 ─────────────────────────────────────────────────────────
    For a = SameCount To EndCount
        If LineBuff(a) = "" Then GoTo ContinueA
        If InStr(LineBuff(a), "|") = 0 Then GoTo ContinueA
        buf = Split(LineBuff(a), "|")
        If UBound(buf) < 5 Then GoTo ContinueA
        wname = Trim(buf(5))
        If wname = "" Then GoTo ContinueA

        timeStamp = Left(LineBuff(a), 19)   '先頭19文字 = "yyyy/mm/dd hh:mm:ss"

        For i = 1 To 14
            Select Case True
            Case InStr(LineBuff(a), "(S" & i & ")") <> 0: Call ApplyWorkerToTimeSheet("S" & i, wname, timeStamp)
            Case InStr(LineBuff(a), "(U" & i & ")") <> 0: Call ApplyWorkerToTimeSheet("U" & i, wname, timeStamp)
            Case InStr(LineBuff(a), "(C" & i & ")") <> 0: Call ApplyWorkerToTimeSheet("C" & i, wname, timeStamp)
            Case InStr(LineBuff(a), "(D" & i & ")") <> 0: Call ApplyWorkerToTimeSheet("D" & i, wname, timeStamp)
            Case InStr(LineBuff(a), "(E" & i & ")") <> 0: Call ApplyWorkerToTimeSheet("E" & i, wname, timeStamp)
            End Select
        Next i
        For i = 1 To 10
            If InStr(LineBuff(a), "(F" & i & ")") <> 0 Then Call ApplyWorkerToTimeSheet("F" & i, wname, timeStamp)
        Next i
ContinueA:
    Next a
End Sub

'印刷済み呼び出し
Public Sub HistoryLoad(HistLot As String, Optional TempDate As String = "") 'UF管理
'2022.06_耳付き丈3,4追加
'2023.5.25_外観BH追加
If WsCheck("保存用") = False Then Exit Sub  '-------------------------ｼｰﾄなし
Dim saveSh As Worksheet
Dim LR As Long, LC As Long
Dim SaveR As Range
Dim SaveTemp()
Dim ZdataA() As String, GdataA() As String, ZdataB() As String, GdataB() As String, splTemp() As String
Dim av(5), bv(5)
Dim i As Long, a As Long, x As Long, Count As Long, matchFlag As Boolean: matchFlag = False

'2024.11.06_同LotNo対策
Dim Randnum As Long
If TempDate = "" Then Randnum = 2
If TempDate <> "" Then Randnum = 1

UF1.Load.Caption = "off" 'イベント停止
'LoadingはUF1のラベル(Loadingの赤文字)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> '通常品検索
If Not UF管理.Mimi Then
    Set saveSh = ThisWorkbook.Sheets("保存用")
    LR = UR(saveSh, 1)
    LC = UC(saveSh, 1)
    Set SaveR = saveSh.Range(saveSh.Cells(2, 1), saveSh.Cells(LR, LC))
    SaveTemp = SaveR 'シートを配列へ
    
    For i = 1 To UBound(SaveTemp, 1)
      If TempDate = "" Then
        If SaveTemp(i, Randnum) = HistLot Then 'LOTNO一致
          x = i
          matchFlag = True
          Exit For
        End If
      ElseIf TempDate <> "" Then
        If SaveTemp(i, Randnum) = TempDate Then 'か日付の一致***2024.11.06
          x = i
          matchFlag = True
          Exit For
        End If
      End If
    Next i
    
    
    If matchFlag = False Then Exit Sub '一致無し
    Dim SixSplTemp() As String
    
    'リセット,公差反映
    With UF1
      .Lot = SaveTemp(x, 2)
      .ATU = SaveTemp(x, 3)
      .Hab = SaveTemp(x, 4)
      .Tak = SaveTemp(x, 5)
      .Am = SaveTemp(x, 6)
      .Ap = SaveTemp(x, 7)
      .Hm = SaveTemp(x, 8)
      .Hp = SaveTemp(x, 9)
      .Tm = SaveTemp(x, 10)
      .Tp = SaveTemp(x, 11)
      .Zai = ""
      .Tyu = ""
      .Youto = ""
      .TAIK = ""
      .Machin = ""
      .Lateral = ""
      .ReSize.Tag = ""
      .Sumisei.Caption = ""
      
      .flatnesText = ""
      .flatnes = ""
      .ConductivText.Caption = ""
      .Comment.Caption = ""
      .Reien.Caption = ""
      
      .KenNo.Caption = ""
      .KenH.Caption = ""
      
      .Additional.Caption = ""
      .TaikText.BackColor = &H80000006
      .LateralText.BackColor = &H80000006
      .Conductiv.BackColor = &H8000000F
      .IzyouFl.BackColor = &H8000000F
      
      .ApK = Format(val(.ATU) + val(.Ap), "0.000") 'マイナスを左側にしたい
      .AmK = Format(val(.ATU) - val(.Am), "0.000") '- 厚
      .HpK = Format(val(.Hab) + val(.Hp), "0.0")
      .HmK = Format(val(.Hab) - val(.Hm), "0.0") '-巾
      .TpK = Format(val(.Tak) + val(.Tp), "0.0")
      .TmK = Format(val(.Tak) - val(.Tm), "0.0") '-丈
      .Load.Caption = "on"
    
      '外観データAB分割
      If Not IsEmpty(SaveTemp(x, 85)) Or Not IsEmpty(SaveTemp(x, 86)) Or Not IsEmpty(SaveTemp(x, 87)) Or Not IsEmpty(SaveTemp(x, 88)) Or Not IsEmpty(SaveTemp(x, 89)) Then
        For i = 85 To 89
          splTemp = Split(SaveTemp(x, i), ",")
          
          If Not IsEmptyArray(splTemp) Then '配列が空かどうかのチェック
            av(i - 85) = splTemp(LBound(splTemp, 1))
            bv(i - 85) = splTemp(UBound(splTemp, 1))
          End If
          
          If Not IsEmpty(SaveTemp(x, 113)) Then
            SixSplTemp = Split(SaveTemp(x, 113), ",") '6回目の検査がある場合の外観
            If Not IsEmptyArray(SixSplTemp) Then
              av(UBound(av)) = SixSplTemp(LBound(SixSplTemp, 1))
              bv(UBound(bv)) = SixSplTemp(UBound(SixSplTemp, 1))
            End If
          End If
          
        Next i
'        Stop
        .S_UpVisual.Caption = Replace(av(0), "A_", "")
        .U_UpVisual.Caption = Replace(av(1), "A_", "")
        .C_UpVisual.Caption = Replace(av(2), "A_", "")
        .D_UpVisual.Caption = Replace(av(3), "A_", "")
        .E_UpVisual.Caption = Replace(av(4), "A_", "")
        .F_UpVisual.Caption = Replace(av(5), "A_", "")
    
        .S_UnderVisual.Caption = Replace(bv(0), "B_", "")
        .U_UnderVisual.Caption = Replace(bv(1), "B_", "")
        .C_UnderVisual.Caption = Replace(bv(2), "B_", "")
        .D_UnderVisual.Caption = Replace(bv(3), "B_", "")
        .E_UnderVisual.Caption = Replace(bv(4), "B_", "")
        .F_UnderVisual.Caption = Replace(bv(5), "B_", "")
      End If
    End With
    
    If SaveTemp(x, 84) = "〆" Then
      UF1.LGS.Caption = "ﾌﾘｰｿﾘ"
      UF1.LGU.Caption = "ﾌﾘｰｿﾘ"
      UF1.LGC.Caption = "ﾌﾘｰｿﾘ"
      UF1.LGD.Caption = "ﾌﾘｰｿﾘ"
      UF1.LGE.Caption = "ﾌﾘｰｿﾘ"
      UF1.LGF.Caption = "ﾌﾘｰｿﾘ"
    End If
    
    '対角再計算_2023.5.25
    Dim Tai(1 To 6) As Single
    Tai(1) = Abs(val(SaveTemp(x, 20)) - val(SaveTemp(x, 21)))
    Tai(2) = Abs(val(SaveTemp(x, 32)) - val(SaveTemp(x, 33)))
    Tai(3) = Abs(val(SaveTemp(x, 44)) - val(SaveTemp(x, 45)))
    Tai(4) = Abs(val(SaveTemp(x, 56)) - val(SaveTemp(x, 57)))
    Tai(5) = Abs(val(SaveTemp(x, 68)) - val(SaveTemp(x, 69)))
    Tai(6) = Abs(val(SaveTemp(x, 108)) - val(SaveTemp(x, 109)))
    
  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>6回目の検査がある場合
    Dim KenNoL As String: KenNoL = Left(SaveTemp(x, 83), 1)
    Dim KenNoR As String: KenNoR = Right(SaveTemp(x, 83), 1)
      If KenNoR = "4" Then
        For i = 1 To 10
            UF1("F" & i) = SaveTemp(x, i + 99)
        Next i
        If SaveTemp(x, 111) <> "" Then UF1("LateralF") = Format(val(SaveTemp(x, 112)), "0.0") 'ﾗﾃﾗﾙﾎﾞｰ
        If SaveTemp(x, 113) <> "" Then UF1.DiagFKou = Format(val(SaveTemp(x, 114)), "0.0") '直角度
      End If
    
    For i = 1 To 17 '---------------------------------------------検査結果反映 桁揃え
      Select Case i
      Case 1 To 4
        UF1("S" & i) = Format(val(SaveTemp(x, i + 11)), "0.000")
        UF1("U" & i) = Format(val(SaveTemp(x, i + 23)), "0.000")
        UF1("C" & i) = Format(val(SaveTemp(x, i + 35)), "0.000")
        UF1("D" & i) = Format(val(SaveTemp(x, i + 47)), "0.000")
        UF1("E" & i) = Format(val(SaveTemp(x, i + 59)), "0.000")
        UF1("F" & i) = Format(val(UF1("F" & i)), "0.000")
      Case 5 To 10
        UF1("S" & i) = Format(val(SaveTemp(x, i + 11)), "0.0")
        UF1("U" & i) = Format(val(SaveTemp(x, i + 23)), "0.0")
        UF1("C" & i) = Format(val(SaveTemp(x, i + 35)), "0.0")
        UF1("D" & i) = Format(val(SaveTemp(x, i + 47)), "0.0")
        UF1("E" & i) = Format(val(SaveTemp(x, i + 59)), "0.0")
        UF1("F" & i) = Format(val(UF1("F" & i)), "0.0")
      Case 11  '対角_測定がある場合だけ_2024.8.5
        If UF1("S" & 10) <> "" Then UF1.DiagS = Format(Tai(1), "0.0")
        If UF1("U" & 10) <> "" Then UF1.DiagU = Format(Tai(2), "0.0")
        If UF1("C" & 10) <> "" Then UF1.DiagC = Format(Tai(3), "0.0")
        If UF1("D" & 10) <> "" Then UF1.DiagD = Format(Tai(4), "0.0")
        If UF1("E" & 10) <> "" Then UF1.DiagE = Format(Tai(5), "0.0")
        If UF1("F" & 10) <> "" Then UF1.DiagF = Format(Tai(6), "0.0")
      Case 12 '直角度_2023.11.29
        If SaveTemp(x, 95) <> "" Then UF1.DiagSKou = Format(val(SaveTemp(x, 95)), "0.0")
        If SaveTemp(x, 96) <> "" Then UF1.DiagUKou = Format(val(SaveTemp(x, 96)), "0.0")
        If SaveTemp(x, 97) <> "" Then UF1.DiagCKou = Format(val(SaveTemp(x, 97)), "0.0")
        If SaveTemp(x, 98) <> "" Then UF1.DiagDKou = Format(val(SaveTemp(x, 98)), "0.0")
        If SaveTemp(x, 99) <> "" Then UF1.DiagEKou = Format(val(SaveTemp(x, 99)), "0.0")
      Case 15 'ﾗﾃﾗﾙﾎﾞｰ
        If SaveTemp(x, 72) <> "" Then UF1("LateralS") = Format(val(SaveTemp(x, 72)), "0.0")
        If SaveTemp(x, 73) <> "" Then UF1("LateralU") = Format(val(SaveTemp(x, 73)), "0.0")
        If SaveTemp(x, 74) <> "" Then UF1("LateralC") = Format(val(SaveTemp(x, 74)), "0.0")
        If SaveTemp(x, 75) <> "" Then UF1("LateralD") = Format(val(SaveTemp(x, 75)), "0.0")
        If SaveTemp(x, 76) <> "" Then UF1("LateralE") = Format(val(SaveTemp(x, 76)), "0.0")
      Case 16 'BH
        If SaveTemp(x, 90) <> "" Then UF1("SBH") = Format(val(SaveTemp(x, 90)), "0")
        If SaveTemp(x, 91) <> "" Then UF1("UBH") = Format(val(SaveTemp(x, 91)), "0")
        If SaveTemp(x, 92) <> "" Then UF1("CBH") = Format(val(SaveTemp(x, 92)), "0")
        If SaveTemp(x, 93) <> "" Then UF1("DBH") = Format(val(SaveTemp(x, 93)), "0")
        If SaveTemp(x, 94) <> "" Then UF1("EBH") = Format(val(SaveTemp(x, 94)), "0")
      End Select
    Next i
    
'  Dim TmpSplit As Variant
  '歪AB対応_2023.6.14
   Application.EnableEvents = False
   
    Dim SplitArrZ, SplitArrG, TempA, tempB, tempC, TempD
    Dim hitA As Long: hitA = 0
    '定盤
    For Count = 22 To 70 Step 12
    hitA = hitA + 1
    If SaveTemp(x, Count) <> "" Then
      If InStr(SaveTemp(x, Count), ",") > 0 Then 'AB面のデータなのか
        SplitArrZ = Split(SaveTemp(x, Count), ",") 'Z
        If InStr(SplitArrZ(LBound(SplitArrZ, 1)), "A:") > 0 Then
          TempA = Replace(SplitArrZ(LBound(SplitArrZ, 1)), "A:", "") '表示
        Else 'AデータなしBデータが2つ目に入っている
          If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then TempA = "" 'Aの表示不要
        End If

        If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then tempB = Replace(SplitArrZ(UBound(SplitArrZ, 1)), "B:", "")  'tag

        ZdataA() = Split(TempA, "_") 'Split関数0配列
        ZdataB() = Split(tempB, "_")
          If UBound(ZdataA) = 2 Then ' 添え字上限が3である場合
            Select Case hitA
            Case 1: Call Split_Z("S", ZdataA, ZdataB)
            Case 2: Call Split_Z("U", ZdataA, ZdataB)
            Case 3: Call Split_Z("C", ZdataA, ZdataB)
            Case 4: Call Split_Z("D", ZdataA, ZdataB)
            Case 5: Call Split_Z("E", ZdataA, ZdataB)
            End Select
          ElseIf IsEmptyArray(ZdataA) Then
            Select Case hitA
            Case 1: Call Split_Z_0("S", ZdataB)
            Case 2: Call Split_Z_0("U", ZdataB)
            Case 3: Call Split_Z_0("C", ZdataB)
            Case 4: Call Split_Z_0("D", ZdataB)
            Case 5: Call Split_Z_0("E", ZdataB)
            End Select
          End If
      Else 'Aのみ
        ZdataA() = Split(Replace(SaveTemp(x, Count), "A:", ""), "_")
        If UBound(ZdataA) = 2 Then ' 添え字上限が3である場合
          Select Case hitA
          Case 1: Call Split_Z("S", ZdataA)
          Case 2: Call Split_Z("U", ZdataA)
          Case 3: Call Split_Z("C", ZdataA)
          Case 4: Call Split_Z("D", ZdataA)
          Case 5: Call Split_Z("E", ZdataA)
          End Select
        ElseIf IsEmptyArray(ZdataA) Then
          Select Case hitA
          Case 1: Call Split_Z_0("S")
          Case 2: Call Split_Z_0("U")
          Case 3: Call Split_Z_0("C")
          Case 4: Call Split_Z_0("D")
          Case 5: Call Split_Z_0("E")
          End Select
        End If
      End If
    End If
    Next Count
'    '6回目用_2024.6.04
      If InStr(SaveTemp(x, 110), ",") > 0 Then 'AB面のデータなのか
        SplitArrZ = Split(SaveTemp(x, 110), ",") 'Z
        If InStr(SplitArrZ(LBound(SplitArrZ, 1)), "A:") > 0 Then
          TempA = Replace(SplitArrZ(LBound(SplitArrZ, 1)), "A:", "") '表示
        Else 'AデータなしBデータが2つ目に入っている
          If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then TempA = "" 'Aの表示不要
        End If
        If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then tempB = Replace(SplitArrZ(UBound(SplitArrZ, 1)), "B:", "")  'tag
        ZdataA() = Split(TempA, "_") 'Split関数0配列
        ZdataB() = Split(tempB, "_")
          If UBound(ZdataA) = 2 Then ' 添え字上限が3である場合
            Call Split_Z("F", ZdataA, ZdataB)
          ElseIf IsEmptyArray(ZdataA) Then
            Call Split_Z_0("F", ZdataB)
          End If
      Else 'Aのみ
        ZdataA() = Split(Replace(SaveTemp(x, 110), "A:", ""), "_")
        If UBound(ZdataA) = 2 Then ' 添え字上限が3である場合
          Call Split_Z("F", ZdataA)
        ElseIf IsEmptyArray(ZdataA) Then
           Call Split_Z_0("F")
        End If
      End If


    'ゲージ
    hitA = 0
    For Count = 23 To 71 Step 12
    hitA = hitA + 1
      If InStr(SaveTemp(x, Count), ",") > 0 Then
        SplitArrG = Split(SaveTemp(x, Count), ",") 'G
        If InStr(SplitArrG(LBound(SplitArrG, 1)), "A:") > 0 Then
          tempC = Replace(SplitArrG(LBound(SplitArrG, 1)), "A:", "") '表示
        Else
          If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then tempC = ""
        End If

        If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then TempD = Replace(SplitArrG(UBound(SplitArrG, 1)), "B:", "") 'tag
        GdataA() = Split(tempC, "_") 'Split関数0配列
        GdataB() = Split(TempD, "_")
        If UBound(GdataA) = 2 Then ' 添え字上限が3である場合 Split_G_0 ("S")
          Select Case hitA
          Case 1: Call Split_G("S", GdataA, GdataB)
          Case 2: Call Split_G("U", GdataA, GdataB)
          Case 3: Call Split_G("C", GdataA, GdataB)
          Case 4: Call Split_G("D", GdataA, GdataB)
          Case 5: Call Split_G("E", GdataA, GdataB)
          End Select
        ElseIf IsEmptyArray(GdataA) Then
          Select Case hitA
          Case 1: Call Split_G_0("S")
          Case 2: Call Split_G_0("U")
          Case 3: Call Split_G_0("C")
          Case 4: Call Split_G_0("D")
          Case 5: Call Split_G_0("E")
          End Select
        End If
      Else
      GdataA() = Split(Replace(SaveTemp(x, Count), "A:", ""), "_")
        If UBound(GdataA) = 2 Then ' 添え字上限が3である場合
          Select Case hitA
          Case 1: Call Split_G("S", GdataA)
          Case 2: Call Split_G("U", GdataA)
          Case 3: Call Split_G("C", GdataA)
          Case 4: Call Split_G("D", GdataA)
          Case 5: Call Split_G("E", GdataA)
          End Select
        ElseIf IsEmptyArray(GdataA) Then
          Select Case hitA
          Case 1: Call Split_G_0("S", GdataB)
          Case 2: Call Split_G_0("U", GdataB)
          Case 3: Call Split_G_0("C", GdataB)
          Case 4: Call Split_G_0("D", GdataB)
          Case 5: Call Split_G_0("E", GdataB)
          End Select
        End If
      End If
    Next Count
    '6回目用_2024.6.04
      If InStr(SaveTemp(x, 111), ",") > 0 Then
        SplitArrG = Split(SaveTemp(x, 111), ",") 'G
        If InStr(SplitArrG(LBound(SplitArrG, 1)), "A:") > 0 Then
          tempC = Replace(SplitArrG(LBound(SplitArrG, 1)), "A:", "") '表示
        Else
          If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then tempC = ""
        End If

        If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then TempD = Replace(SplitArrG(UBound(SplitArrG, 1)), "B:", "") 'tag
        GdataA() = Split(tempC, "_") 'Split関数0配列
        GdataB() = Split(TempD, "_")
        If UBound(GdataA) = 2 Then ' 添え字上限が3である場合 Split_G_0 ("S")
          Call Split_G("F", GdataA, GdataB)
        ElseIf IsEmptyArray(GdataA) Then
          Call Split_G_0("F")
        End If
      Else
      GdataA() = Split(Replace(SaveTemp(x, 111), "A:", ""), "_")
        If UBound(GdataA) = 2 Then ' 添え字上限が3である場合
          Call Split_G("F", GdataA)
        ElseIf IsEmptyArray(GdataA) Then
          Call Split_G_0("F", GdataB)
        End If
      End If

   Call ReflectStrain(Right(UF1.AorB.Caption, 1), "保存") 'これがないとBが読み込めない
   Application.EnableEvents = True

'    Stop
    
    For i = 1 To 10 '書き換え
      If UF1("S" & i) = 0 Then
        UF1("S" & i) = ""
        UF1("S" & i).BackColor = &H80000005 '白　ﾃﾞﾌｫﾙﾄ色
      End If
      If UF1("U" & i) = 0 Then
        UF1("U" & i) = ""
        UF1("U" & i).BackColor = &H80000005
      End If
      If UF1("C" & i) = 0 Then
        UF1("C" & i) = ""
        UF1("C" & i).BackColor = &H80000005
      End If
      If UF1("D" & i) = 0 Then
        UF1("D" & i) = ""
        UF1("D" & i).BackColor = &H80000005
      End If
      If UF1("E" & i) = 0 Then
        UF1("E" & i) = ""
        UF1("E" & i).BackColor = &H80000005
      End If
      If UF1("F" & i) = 0 Then
        UF1("F" & i) = ""
        UF1("F" & i).BackColor = &H80000005
      End If
    Next i

    With UF1
      If KenNoL = "H" And KenNoR = "数" Then Call LSelect("少枚数")
      If KenNoL = "H" And KenNoR = "H" Then Call LSelect("厚板")
      If KenNoL <> "H" And KenNoR = "1" Then Call LSelect("薄板")
      If KenNoL <> "H" And KenNoR = "2" Then Call LSelect("薄板")
      If KenNoL <> "H" And KenNoR = "3" Then Call LSelect("薄板")
      If (KenNoL <> "H" And KenNoR = "4") Then Call Six_Display("on")
      If .LateralS <> "" Or .LateralU <> "" Or .LateralC <> "" Or .LateralD <> "" Or .LateralE <> "" Then Call LSelectLateral("強制")  '2021.07.16
      If .DiagSKou <> "" Or .DiagUKou <> "" Or .DiagCKou <> "" Or .DiagDKou <> "" Or .DiagEKou <> "" Then Call RightAngleDisplay("強制") '2023.11.29
      If .LateralF <> "" Then Call LSelectLateral("強制", "LateralF")
      If .DiagFKou <> "" Then Call RightAngleDisplay("強制", "DiagFKou")
      
      If UF管理.Visible Then
        .S = UF管理.S.Caption
        .U = UF管理.U.Caption
        .C = UF管理.C.Caption
        .D = UF管理.D.Caption
        .E = UF管理.E.Caption
        .F = UF管理.F.Caption
        .FS = UF管理.S.Caption
        .FU = UF管理.U.Caption
        .FC = UF管理.C.Caption
        .FD = UF管理.D.Caption
        .FE = UF管理.E.Caption
        .FF = UF管理.F.Caption
      End If
    End With
    
    UF1.VC_L_A.Tag = ""
    UF1.VC_L_B.Tag = ""
    Dim Temp As Variant:  Temp = InitialSetting(UF1)

    MsgBox "Load完了,印刷する場合はフォームを閉じないでください。", vbInformation
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> '耳付き
Else
    Set saveSh = ThisWorkbook.Sheets("耳用")
    LR = UR(saveSh, 1)
    LC = UC(saveSh, 1)
    Set SaveR = saveSh.Range(saveSh.Cells(2, 1), saveSh.Cells(LR, LC))
    SaveTemp = SaveR 'シートを配列へ
    
    For i = 1 To UBound(SaveTemp, 1)
      If SaveTemp(i, 2) = HistLot Then 'LOTNO一致
        x = i
        matchFlag = True
        Exit For
      End If
    Next i
    If matchFlag = False Then Exit Sub '一致無し
    With UF1
      .Lot = SaveTemp(x, 2)
      .ATU = SaveTemp(x, 3)
      .Hab = SaveTemp(x, 4)
      .Tak = SaveTemp(x, 5)
      .Am = SaveTemp(x, 6)
      .Ap = SaveTemp(x, 7)
      .Hm = SaveTemp(x, 8)
      .Hp = SaveTemp(x, 9)
      .Tm = SaveTemp(x, 10)
      .Tp = SaveTemp(x, 11)
      .Zai = ""
      .Tyu = ""
      .Youto = ""
      .TAIK = ""
      .Machin = ""
      .Lateral = ""
      .ReSize.Tag = ""
      .Sumisei.Caption = ""
      
      .flatnesText = ""
      .flatnes = ""
      .ConductivText.Caption = ""
      .Comment.Caption = ""
      .Reien.Caption = ""
      
      .KenNo.Caption = ""
      .KenH.Caption = ""
      
      .Additional.Caption = ""
      .TaikText.BackColor = &H80000006
      .LateralText.BackColor = &H80000006
      .Conductiv.BackColor = &H8000000F
      .IzyouFl.BackColor = &H8000000F
      
      .ApK = Format(val(.ATU) + val(.Ap), "0.000") 'マイナスを左側にしたい
      .AmK = Format(val(.ATU) - val(.Am), "0.000") '- 厚
      .HpK = Format(val(.Hab) + val(.Hp), "0.0")
      .HmK = Format(val(.Hab) - val(.Hm), "0.0") '-巾
      .TpK = Format(val(.Tak) + val(.Tp), "0.0")
      .TmK = Format(val(.Tak) - val(.Tm), "0.0") '-丈
      .Load.Caption = "on"
       
      '外観データAB分割
      If Not IsEmpty(SaveTemp(x, 99)) Or Not IsEmpty(SaveTemp(x, 100)) Or Not IsEmpty(SaveTemp(x, 101)) Or Not IsEmpty(SaveTemp(x, 102)) Or Not IsEmpty(SaveTemp(x, 103)) Then
        For i = 99 To 103
          splTemp = Split(SaveTemp(x, i), ",")
          
          If Not IsEmptyArray(splTemp) Then '配列が空かどうかのチェック
            av(i - 99) = splTemp(LBound(splTemp, 1))
            bv(i - 99) = splTemp(UBound(splTemp, 1))
          End If
        Next i
        .S_UpVisual.Caption = Replace(av(0), "A_", "")
        .U_UpVisual.Caption = Replace(av(1), "A_", "")
        .C_UpVisual.Caption = Replace(av(2), "A_", "")
        .D_UpVisual.Caption = Replace(av(3), "A_", "")
        .E_UpVisual.Caption = Replace(av(4), "A_", "")
    
        .S_UnderVisual.Caption = Replace(bv(0), "B_", "")
        .U_UnderVisual.Caption = Replace(bv(1), "B_", "")
        .C_UnderVisual.Caption = Replace(bv(2), "B_", "")
        .D_UnderVisual.Caption = Replace(bv(3), "B_", "")
        .E_UnderVisual.Caption = Replace(bv(4), "B_", "")
      End If
    End With
    
    For i = 1 To 17 '---------------------------------------------検査結果反映 桁揃え
      For a = 1 To 5
        Select Case i
        Case 1 To 4
          Select Case a
          Case 1:  UF1("S" & i) = Format(val(SaveTemp(x, i + 11)), "0.000")
          Case 2:  UF1("U" & i) = Format(val(SaveTemp(x, i + 27)), "0.000")
          Case 3:  UF1("C" & i) = Format(val(SaveTemp(x, i + 43)), "0.000")
          Case 4:  UF1("D" & i) = Format(val(SaveTemp(x, i + 59)), "0.000")
          Case 5:  UF1("E" & i) = Format(val(SaveTemp(x, i + 75)), "0.000")
          End Select
        Case 5 To 14
          Select Case a
          Case 1:  UF1("S" & i) = Format(val(SaveTemp(x, i + 11)), "0.0")
          Case 2:  UF1("U" & i) = Format(val(SaveTemp(x, i + 27)), "0.0")
          Case 3:  UF1("C" & i) = Format(val(SaveTemp(x, i + 43)), "0.0")
          Case 4:  UF1("D" & i) = Format(val(SaveTemp(x, i + 59)), "0.0")
          Case 5:  UF1("E" & i) = Format(val(SaveTemp(x, i + 75)), "0.0")
          End Select
        Case 15 '定盤歪--------------3つの入力がない場合Loadなし
            Select Case a
            Case 1
              ZdataA() = Split(SaveTemp(x, 26), "_") 'Split関数0配列
              If UBound(ZdataA) = 2 Then Call Split_Z("S", ZdataA) '添え字上限でないとｴﾗｰ
            Case 2
              ZdataA() = Split(SaveTemp(x, 42), "_")
              If UBound(ZdataA) = 2 Then Call Split_Z("U", ZdataA) '
            Case 3
              ZdataA() = Split(SaveTemp(x, 58), "_")
              If UBound(ZdataA) = 2 Then Call Split_Z("C", ZdataA) '
            Case 4
              ZdataA() = Split(SaveTemp(x, 74), "_")
              If UBound(ZdataA) = 2 Then Call Split_Z("D", ZdataA) '
            Case 5
              ZdataA() = Split(SaveTemp(x, 90), "_")
              If UBound(ZdataA) = 2 Then Call Split_Z("E", ZdataA) '
            End Select
        Case 16 'ｹﾞｰｼﾞ歪
            Select Case a
            Case 1
              GdataA() = Split(SaveTemp(x, 27), "_")
              If UBound(GdataA) = 2 Then Call Split_G("S", GdataA)
            Case 2
              GdataA() = Split(SaveTemp(x, 43), "_")
              If UBound(GdataA) = 2 Then Call Split_G("U", GdataA)
            Case 3
              GdataA() = Split(SaveTemp(x, 59), "_")
              If UBound(GdataA) = 2 Then Call Split_G("C", GdataA)
            Case 4
              GdataA() = Split(SaveTemp(x, 75), "_")
              If UBound(GdataA) = 2 Then Call Split_G("D", GdataA)
            Case 5
              GdataA() = Split(SaveTemp(x, 91), "_")
              If UBound(GdataA) = 2 Then Call Split_G("E", GdataA)
            End Select
        Case 17 'BH
            Select Case a
            Case 1: If SaveTemp(x, 104) <> "" Then UF1("SBH") = Format(val(SaveTemp(x, 104)), "0")
            Case 2: If SaveTemp(x, 105) <> "" Then UF1("UBH") = Format(val(SaveTemp(x, 105)), "0")
            Case 3: If SaveTemp(x, 106) <> "" Then UF1("CBH") = Format(val(SaveTemp(x, 106)), "0")
            Case 4: If SaveTemp(x, 107) <> "" Then UF1("DBH") = Format(val(SaveTemp(x, 107)), "0")
            Case 5: If SaveTemp(x, 108) <> "" Then UF1("EBH") = Format(val(SaveTemp(x, 108)), "0")
            End Select
        End Select
      Next a
    Next i
    'Stop
    For i = 1 To 14 '書き換え
      For a = 1 To 5
        Select Case a
        Case 1
          If UF1("S" & i) = 0 Then
            UF1("S" & i) = ""
            UF1("S" & i).BackColor = &H80000005 '白　ﾃﾞﾌｫﾙﾄ色
          End If
        Case 2
          If UF1("U" & i) = 0 Then
            UF1("U" & i) = ""
            UF1("U" & i).BackColor = &H80000005
          End If
        Case 3
          If UF1("C" & i) = 0 Then
            UF1("C" & i) = ""
            UF1("C" & i).BackColor = &H80000005
          End If
        Case 4
          If UF1("D" & i) = 0 Then
            UF1("D" & i) = ""
            UF1("D" & i).BackColor = &H80000005
          End If
        Case 5
          If UF1("E" & i) = 0 Then
            UF1("E" & i) = ""
            UF1("E" & i).BackColor = &H80000005
          End If
        End Select
      Next a
    Next i
    
    MsgBox "Load完了,印刷する場合はフォームを閉じないでください。", vbInformation
    
End If
  
'Call Atu_OPselect(Val(SaveTemp(x, 3)))'2023.11.29_表示がうまくできない　一旦停止
    
Dim n As String
If Not UF管理.Mimi Then n = SaveTemp(x, 83)
If UF管理.Mimi Then n = "耳付き"

'Select Case n
'Case "厚板":     UF1.厚板 = True
'Case "5回検査":  UF1.薄板 = True
'Case "少枚数":   UF1.少枚数 = True
'Case "耳付き":   UF1.極厚 = True
'End Select

Call ReDiag '対角再計算(線差)
UF1.Load.Caption = "on" 'イベント再開
UF1.dataLoding.Caption = "Load"


Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("WorkerTime" & "_" & Lname) = False Then Call AddWS("WorkerTime" & "_" & Lname)
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("WorkerTime" & "_" & Lname)
sh.Cells(1, 1) = "履歴呼び出し"
End Sub

'Accesslogシートcol 23-26→作業保存シートrows 2-7 / 8-13, col 15-16→LosdStrain("A", ...)→Z/G コントロール (A面)+ Tag (B面)
Public Sub FormReflection(Optional Temp As String = "") 'AccesslogLoadで読み取ったデータをフォームに反映
If WsCheck("Accesslog") = False Then Exit Sub  '-------------------------ｼｰﾄなし
Dim saveSh As Worksheet: Set saveSh = ThisWorkbook.Sheets("Accesslog") '専用シート
Dim LR As Long: LR = UR(saveSh, 2)
Dim LC As Long: LC = UC(saveSh, 1)
'2023.6.15_ちょっと修正可読性上げた
Dim SaveR As Range: Set SaveR = saveSh.Range(saveSh.Cells(1, 2), saveSh.Cells(6, 17))
Dim SaveTemp():     SaveTemp = SaveR 'シートを配列へ
Dim i As Long, a As Long
Dim Usually As Boolean: Usually = False
For i = LBound(SaveTemp, 1) To UBound(SaveTemp, 1) '通常か耳付き検査かを見分ける***耳付きの場合11,12,13,14
  If IsEmpty(SaveTemp(i, 11)) = False Then Usually = True 'empty値ではないとき(数字がある)
Next i
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<耳付き
If Usually Then
    UF1.Lot = saveSh.Cells(1, 1)
    For i = LBound(SaveTemp, 2) To UBound(SaveTemp, 2) '---------------------------------------------検査結果反映 桁揃え
      For a = LBound(SaveTemp, 1) To UBound(SaveTemp, 1)
        Select Case i
        Case 1 To 4
          Select Case a
          Case 1:  UF1("S" & i) = Format(val(SaveTemp(a, i)), "0.000")
          Case 2:  UF1("U" & i) = Format(val(SaveTemp(a, i)), "0.000")
          Case 3:  UF1("C" & i) = Format(val(SaveTemp(a, i)), "0.000")
          Case 4:  UF1("D" & i) = Format(val(SaveTemp(a, i)), "0.000")
          Case 5:  UF1("E" & i) = Format(val(SaveTemp(a, i)), "0.000")
          End Select
        Case 5 To 14
          Select Case a
          Case 1:  UF1("S" & i) = Format(val(SaveTemp(a, i)), "0.0")
          Case 2:  UF1("U" & i) = Format(val(SaveTemp(a, i)), "0.0")
          Case 3:  UF1("C" & i) = Format(val(SaveTemp(a, i)), "0.0")
          Case 4:  UF1("D" & i) = Format(val(SaveTemp(a, i)), "0.0")
          Case 5:  UF1("E" & i) = Format(val(SaveTemp(a, i)), "0.0")
          End Select
        End Select
      Next a
    Next i
    Call LSelect("極厚")
'    Stop
    MsgBox "Load完了", vbInformation
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<通常検査
Else
    UF1.Lot = saveSh.Cells(1, 1)
    For i = LBound(SaveTemp, 2) To UBound(SaveTemp, 2) '---------------------------------------------検査結果反映 桁揃え
      For a = LBound(SaveTemp, 1) To UBound(SaveTemp, 1)
        Select Case i
        Case 1 To 4
          Select Case a
          Case 1:  UF1("S" & i) = Format(val(SaveTemp(a, i)), "0.000")
          Case 2:  UF1("U" & i) = Format(val(SaveTemp(a, i)), "0.000")
          Case 3:  UF1("C" & i) = Format(val(SaveTemp(a, i)), "0.000")
          Case 4:  UF1("D" & i) = Format(val(SaveTemp(a, i)), "0.000")
          Case 5:  UF1("E" & i) = Format(val(SaveTemp(a, i)), "0.000")
          Case 6:  UF1("F" & i) = Format(val(SaveTemp(a, i)), "0.000")
          End Select
        Case 5 To 10
          Select Case a
          Case 1:  UF1("S" & i) = Format(val(SaveTemp(a, i)), "0.0")
          Case 2:  UF1("U" & i) = Format(val(SaveTemp(a, i)), "0.0")
          Case 3:  UF1("C" & i) = Format(val(SaveTemp(a, i)), "0.0")
          Case 4:  UF1("D" & i) = Format(val(SaveTemp(a, i)), "0.0")
          Case 5:  UF1("E" & i) = Format(val(SaveTemp(a, i)), "0.0")
          Case 6:  UF1("F" & i) = Format(val(SaveTemp(a, i)), "0.0")
          End Select
        End Select
      Next a
    Next i
    'Stop
End If

'直角
UF1.DiagSKou = SaveTemp(1, 15)
UF1.DiagUKou = SaveTemp(2, 15)
UF1.DiagCKou = SaveTemp(3, 15)
UF1.DiagDKou = SaveTemp(4, 15)
UF1.DiagEKou = SaveTemp(5, 15)
UF1.DiagFKou = SaveTemp(6, 15)
'ラテ
UF1.LateralS = SaveTemp(1, 16)
UF1.LateralU = SaveTemp(2, 16)
UF1.LateralC = SaveTemp(3, 16)
UF1.LateralD = SaveTemp(4, 16)
UF1.LateralE = SaveTemp(5, 16)
UF1.LateralF = SaveTemp(6, 16)

For i = 1 To 14 '書き換え
  If UF1("S" & i) = 0 Then
    UF1("S" & i) = ""
    UF1("S" & i).BackColor = &H80000005 '白　ﾃﾞﾌｫﾙﾄ色
  End If
  If UF1("U" & i) = 0 Then
    UF1("U" & i) = ""
    UF1("U" & i).BackColor = &H80000005
  End If
  If UF1("C" & i) = 0 Then
    UF1("C" & i) = ""
    UF1("C" & i).BackColor = &H80000005
  End If
  If UF1("D" & i) = 0 Then
    UF1("D" & i) = ""
    UF1("D" & i).BackColor = &H80000005
  End If
  If UF1("E" & i) = 0 Then
    UF1("E" & i) = ""
    UF1("E" & i).BackColor = &H80000005
  End If
Next i
For i = 1 To 10
  If UF1("F" & i) = 0 Then
    UF1("F" & i) = ""
    UF1("F" & i).BackColor = &H80000005
  End If
Next i

If UF1.DiagSKou <> "" Then UF1.DiagSKou = Format(UF1.DiagSKou, "0.0")
If UF1.DiagUKou <> "" Then UF1.DiagUKou = Format(UF1.DiagUKou, "0.0")
If UF1.DiagCKou <> "" Then UF1.DiagCKou = Format(UF1.DiagCKou, "0.0")
If UF1.DiagDKou <> "" Then UF1.DiagDKou = Format(UF1.DiagDKou, "0.0")
If UF1.DiagEKou <> "" Then UF1.DiagEKou = Format(UF1.DiagEKou, "0.0")
If UF1.DiagFKou <> "" Then UF1.DiagFKou = Format(UF1.DiagFKou, "0.0")

If UF1.LateralS <> "" Then UF1.LateralS = Format(UF1.LateralS, "0.0")
If UF1.LateralU <> "" Then UF1.LateralU = Format(UF1.LateralU, "0.0")
If UF1.LateralC <> "" Then UF1.LateralC = Format(UF1.LateralC, "0.0")
If UF1.LateralD <> "" Then UF1.LateralD = Format(UF1.LateralD, "0.0")
If UF1.LateralE <> "" Then UF1.LateralE = Format(UF1.LateralE, "0.0")
If UF1.LateralF <> "" Then UF1.LateralF = Format(UF1.LateralF, "0.0")

Call ReDiag '対角再計算(線差)
'If SaveTemp(UBound(SaveTemp, 1), 1) <> "" And SaveTemp(UBound(SaveTemp, 1), 2) Then
'  Call Six_Display("on")
'Else
'  Call LSelect("少枚数")
'End If
MsgBox "Load完了", vbInformation

End Sub

'外観はそのまま反映。歪は作業保存シートの col 15-16 (A面: rows 2-7 / B面: rows 8-13) に転記後、 LosdStrain("A", A_range, B_range) を呼ぶ。Load_Text と完全に同パターン
Public Sub FormReflectionEx()
    If WsCheck("Accesslog") = False Then Exit Sub
    Dim accSh As Worksheet: Set accSh = ThisWorkbook.Sheets("Accesslog")
    Dim Lname As String:    Lname = UF1.LineName.Caption & UF1.LineNo.Caption

    Dim i As Long
    Dim posArr As Variant: posArr = Array("S", "U", "C", "D", "E", "F")

    '── 外観 ───────────────────────────────────────────────────────────────
    Dim pos As String
    For i = 1 To 6
        pos = posArr(i - 1)
        If accSh.Cells(i, 21).Value <> "" Then UF1(pos & "_UpVisual") = accSh.Cells(i, 21).Value
        If accSh.Cells(i, 22).Value <> "" Then UF1(pos & "_UnderVisual") = accSh.Cells(i, 22).Value
    Next i

    '── 歪: 作業保存が ApplyVisualStrainFromLog で整備済のため常に LosdStrain ──
    If WsCheck("作業保存" & "_" & Lname) = False Then Exit Sub
    Dim saveSh As Worksheet: Set saveSh = ThisWorkbook.Sheets("作業保存" & "_" & Lname)
    Dim A_range As Range: Set A_range = saveSh.Range(saveSh.Cells(2, 15), saveSh.Cells(7, 16))
    Dim B_range As Range: Set B_range = saveSh.Range(saveSh.Cells(8, 15), saveSh.Cells(13, 16))
    Call LosdStrain("A", A_range, B_range)
End Sub

Public Sub MainDataSave(Optional TemWs As String = "保存用", Optional Temp As String = "耳用") 'シート内保存
Dim i As Long, a As Long
Dim OneKen() As String
Dim Sken(), Uken(), Cken(), Dken(), Eken(), Fken()
Dim sh As Worksheet: Set sh = ShJuge(TemWs)
Dim ws As Worksheet: Set ws = ShJuge(Temp)
Dim n As Long

Dim KenNo As String: KenNo = Left(UF1.Reien.Caption, 1)

With UF1
'Size,公差
  Dim HIKIdata As Variant: HIKIdata = Array(.ATU, .Hab, .Tak, .Am, .Ap, .Hm, .Hp, .Tm, .Tp) '配列化
  
  Dim a_S As String: a_S = .S_UpVisual.Caption
  Dim a_U As String: a_U = .U_UpVisual.Caption
  Dim a_C As String: a_C = .C_UpVisual.Caption
  Dim a_D As String: a_D = .D_UpVisual.Caption
  Dim a_E As String: a_E = .E_UpVisual.Caption
  Dim a_F As String: a_F = .F_UpVisual.Caption
  
  Dim b_S As String: b_S = .S_UnderVisual.Caption
  Dim b_U As String: b_U = .U_UnderVisual.Caption
  Dim b_C As String: b_C = .C_UnderVisual.Caption
  Dim b_D As String: b_D = .D_UnderVisual.Caption
  Dim b_E As String: b_E = .E_UnderVisual.Caption
  Dim b_F As String: b_F = .F_UnderVisual.Caption
  
  '外観AB合体
  '履歴でデータを取りやすいように,をいれる
  Dim ab_S As String
    If (a_S <> "") And (b_S <> "") Then
      ab_S = "A_" & a_S & ",B_" & b_S
    ElseIf (a_S <> "") And (b_S = "") Then
      ab_S = "A_" & a_S & ","
    ElseIf (a_S = "") And (b_S <> "") Then
      ab_S = " " & ",B_" & b_S
    End If
  Dim ab_U As String
    If (a_U <> "") And (b_U <> "") Then
      ab_U = "A_" & a_U & ",B_" & b_U
    ElseIf (a_U <> "") And (b_U = "") Then
      ab_U = "A_" & a_U & ","
    ElseIf (a_U = "") And (b_U <> "") Then
      ab_U = " " & ",B_" & b_U
    End If
  Dim ab_C As String
    If (a_C <> "") And (b_C <> "") Then
      ab_C = "A_" & a_C & ",B_" & b_C
    ElseIf (a_C <> "") And (b_C = "") Then
      ab_C = "A_" & a_C & ","
    ElseIf (a_C = "") And (b_C <> "") Then
      ab_C = "" & ",B_" & b_C
    End If
  Dim ab_D As String
    If (a_D <> "") And (b_D <> "") Then
      ab_D = "A_" & a_D & ",B_" & b_D
    ElseIf (a_D <> "") And (b_D = "") Then
      ab_D = "A_" & a_D & ","
    ElseIf (a_D = "") And (b_D <> "") Then
      ab_D = " " & ",B_" & b_D
    End If
  Dim ab_E As String
    If (a_E <> "") And (b_E <> "") Then
      ab_E = "A_" & a_E & ",B_" & b_E
    ElseIf (a_E <> "") And (b_E = "") Then
      ab_E = "A_" & a_E & ","
    ElseIf (a_E = "") And (b_E <> "") Then
      ab_E = " " & ",B_" & b_E
    End If
    
  Dim ab_F As String
    If (a_F <> "") And (b_F <> "") Then
      ab_F = "A_" & a_F & ",B_" & b_F
    ElseIf (a_E <> "") And (b_E = "") Then
      ab_F = "A_" & a_F & ","
    ElseIf (a_F = "") And (b_F <> "") Then
      ab_F = " " & ",B_" & b_F
    End If
  
  Dim VisItem As Variant: VisItem = Array(ab_S, ab_U, ab_C, ab_D, ab_E)
  'BH
  Dim BHdata As Variant:  BHdata = Array(.SBH, .UBH, .CBH, .DBH, .EBH)
  
  '歪AB面対応_2023.6.14
  Dim hizumi_Arry(1 To 6, 1 To 2) As String
  Dim hizumAB(1 To 6, 1 To 2) As String
  Dim hizumi():    hizumi = UF_Array(15, 16, 1, 6) '格納　2024.5.23*6行面(F)追加
  Dim tagHizumi(): tagHizumi = UFtag_Array(1, 6) 'tag(現状表示されていない面の歪数値)
  For i = 1 To 6 '配列入れ直し
    For a = 1 To 2
      hizumi_Arry(i, a) = hizumi(i, 14 + a) 'hizumi_Arryには今表示されている数値が入る
    Next a
  Next i
  
  Dim hizumA As Variant, hizumB As Variant
  
  If UF1.AorB.Caption = "Now_A" Then 'A
    hizumA = hizumi_Arry
    hizumB = tagHizumi
  ElseIf UF1.AorB.Caption = "Now_B" Then 'B
    hizumA = tagHizumi
    hizumB = hizumi_Arry
  End If
  
  For i = LBound(hizumA, 1) To UBound(hizumA, 1) ' Aの各要素に"A:"を追加する
    For a = 1 To 2
      If hizumA(i, a) <> "__" Then
        hizumA(i, a) = "A:" & hizumA(i, a)
      Else
        hizumA(i, a) = ""
      End If
    Next a
  Next i
  
  For i = LBound(hizumB, 1) To UBound(hizumB, 1) ' Bの各要素に"B:"を追加する
    For a = 1 To 2
      If hizumB(i, a) <> "__" Then
        hizumB(i, a) = ",B:" & hizumB(i, a)
      Else
        hizumB(i, a) = ""
      End If
    Next a
  Next i
  '合体
  For i = LBound(hizumAB, 1) To UBound(hizumAB, 1)
    For a = LBound(hizumAB, 2) To UBound(hizumAB, 2)
      hizumAB(i, a) = hizumA(i, a) & hizumB(i, a)
    Next a
  Next i
End With

'検査データ収集
Dim KenDate()
Dim AllKen()
If UF1.極厚 Then
  KenDate = UF_Array(1, 16, 1, 5) '格納:対角線公差異常ﾌﾗｸﾞを除く
Else
  KenDate = UF_SaveArray(1, 12, 1, 5) '通常検査用
End If
AllKen = UF_Array(1, 20, 1, 5) '全て

Dim Lateralbow(1 To 5)
For i = 1 To 5
  Lateralbow(i) = AllKen(i, 20) 'ラテラルボーのみ抜き取る
Next i
Dim RightAngle(1 To 5)
For i = 1 To 5
  RightAngle(i) = AllKen(i, 18) '直角度(旧対角公差)のみ抜き取る
Next i

'6回目だけの検査データを抜く
Dim SixKen:            SixKen = UF_SaveArray(1, 12, 6, 6) '厚幅丈
Dim SixArr:            SixArr = UF_Array(1, 20, 6, 6)
Dim SixLateral(0):     SixLateral(0) = SixArr(6, 20) 'ラテ
Dim SixRightAngle(0):  SixRightAngle(0) = SixArr(6, 18) '直角度
Dim SixVis As Variant: SixVis = Array(ab_F) '外観
Dim Sixhiz(1)
Sixhiz(0) = hizumAB(6, 1) '歪
Sixhiz(1) = hizumAB(6, 2)

Dim SixAll(1 To 1, 1 To 15)
For i = 1 To 15 'UF_Arrayで格納されている順番
  Select Case i
  Case 1 To 10:   SixAll(1, i) = SixKen(6, i) '厚幅丈
  Case 11:        SixAll(1, i) = Sixhiz(0) '定盤歪
  Case 12:        SixAll(1, i) = Sixhiz(1) 'ゲージ歪
  Case 13:        SixAll(1, i) = SixLateral(0) 'ラテ
  Case 14:        SixAll(1, i) = SixVis(0) '外観
  Case 15:        SixAll(1, i) = SixRightAngle(0) '直角度
  End Select
Next i

'Stop

'1次元に変換
Sken = WorksheetFunction.Index(KenDate, 1) 'Sのみ抜き取る
Uken = WorksheetFunction.Index(KenDate, 2) 'Uのみ抜き取る
Cken = WorksheetFunction.Index(KenDate, 3) 'Cのみ抜き取る
Dken = WorksheetFunction.Index(KenDate, 4) 'Dのみ抜き取る
Eken = WorksheetFunction.Index(KenDate, 5) 'Eのみ抜き取る
'配列内ﾃﾞｰﾀを繋げて文字列化し(Join(Sken))、その複数ﾃﾞｰﾀを繋げてSplitで再配列化する(配列のまま行うと型ｴﾗｰ)
OneKen = Split(Join(Sken) & " " & Join(Uken) & " " & Join(Cken) & " " & Join(Dken) & " " & Join(Eken)) '区切り文字省略で" "になる
'Stop

Dim HikiRange As Range 'Size'公差
Dim PastRan As Range '検査データ
Dim PastLate As Range 'ラテラルボー
Dim VisRange As Range '外観
Dim BhRange  As Range 'BH
Dim PastAngle As Range '直角度
Dim PastSix As Range '6回目の検査

Dim Worker
Select Case WorkeDayMax '登録時間が一番遅い
Case "1": Worker = Split(UF1.T1, "_")
Case "2": Worker = Split(UF1.T2, "_")
Case "3": Worker = Split(UF1.T3, "_")
End Select
'---------------------------------------------------------------転送処理
If UF1.極厚 Then
  With ws
    .Unprotect '---------------------------------保護解除
    .Cells.Locked = False
    .Cells.FormulaHidden = False
    n = UR(ws, 1) + 1
    .Cells(n, 2) = UF1.Lot 'ロット
    Set HikiRange = .Range(.Cells(n, 3), .Cells(n, 3 + UBound(HIKIdata)))
    HikiRange = HIKIdata
    Set PastRan = .Range(.Cells(n, 12), .Cells(n, 12 + UBound(OneKen)))
    PastRan = OneKen
    Select Case True
    Case UF1.極厚:  .Cells(n, 98) = "耳付き"
    End Select
    .Cells(n, 1) = Format$(UF1.StartTime, "yyyy/mm/dd hh:mm") & "～" & Format$(UF1.EndTime, "yyyy/mm/dd hh:mm")  '検査開始から印刷された時間
    .Cells(n, 92) = Worker(UBound(Worker)) '検査員
    .Cells(n, 93) = UF1.ReSize.Tag '公差修正
    .Cells(n, 94) = UF1.Zai '基本データ追加_2021.01.22
    .Cells(n, 95) = UF1.Tyu
    .Cells(n, 96) = UF1.Youto
    
    '外観
    Set VisRange = .Range(.Cells(n, 99), .Cells(n, 99 + UBound(VisItem)))
    VisRange = VisItem
    'BH
    Set BhRange = .Range(.Cells(n, 104), .Cells(n, 104 + UBound(BHdata)))
    BhRange = BHdata
    
    'LoadingはUF1のラベル(Loadingの赤文字)
    '関数HistoryLoadで履歴を呼び出して印刷した場合のフラグ_2023.6.7*追加
    If UF1.dataLoding.Caption = "Load" Then .Cells(n, 97) = "〆"
    
    .Cells.Font.Name = "メイリオ"
    .Cells.Font.Size = 9
    .Cells.EntireColumn.AutoFit
    .Cells.FormulaHidden = False
  End With
Else
  With sh
    .Unprotect
    .Cells.Locked = False
    .Cells.FormulaHidden = False
    n = UR(sh, 1) + 1
    .Cells(n, 2) = UF1.Lot 'ロット
    'Size'公差
    Set HikiRange = .Range(.Cells(n, 3), .Cells(n, 3 + UBound(HIKIdata)))
    HikiRange = HIKIdata
    '検査データ
    Set PastRan = .Range(.Cells(n, 12), .Cells(n, 12 + UBound(OneKen)))
    PastRan = OneKen
    'ラテラルボー
    Set PastLate = .Range(.Cells(n, 72), .Cells(n, 72 + UBound(Lateralbow) - 1))
    PastLate = Lateralbow
    '直角度
    Set PastAngle = .Range(.Cells(n, 95), .Cells(n, 95 + UBound(RightAngle) - 1))
    PastAngle = RightAngle
    '6回目検査_2024.5.31
    Set PastSix = .Range(.Cells(n, 100), .Cells(n, 100 + UBound(SixAll, 2) - 1))
    PastSix = SixAll
    
    Select Case True '2023.6.21_修正
    Case UF1.厚板:   .Cells(n, 83) = UF1.厚板.Caption
    Case UF1.薄板:   .Cells(n, 83) = UF1.薄板.Caption & ":" & UF1.Yama.Text
    Case UF1.少枚数: .Cells(n, 83) = UF1.厚板.Caption & ":" & "少枚数"
    Case Else:       .Cells(n, 83) = "ｽﾗﾌﾞ"
    End Select
    
'    '2023.6.7*追加*****上記で修正した2023.6.21
'    If KenNo = "T" Then .Cells(n, 83) = "2HI" '2023.6.7***特殊条件*2HIの場合薄板選択となっているため(しかし板厚は6mm超えの場合は当然ある)
    
    .Cells(n, 1) = Format$(UF1.StartTime, "yyyy/mm/dd hh:mm") & "～" & Format$(UF1.EndTime, "yyyy/mm/dd hh:mm")  '検査開始から印刷された時間
    .Cells(n, 77) = Worker(UBound(Worker)) '検査員
    .Cells(n, 78) = UF1.ReSize.Tag '公差修正
    .Cells(n, 79) = UF1.Zai '基本データ追加_2021.01.22
    .Cells(n, 80) = UF1.Tyu
    .Cells(n, 81) = UF1.Youto
    
    'フリー反り追加_2023.5.15
    If UF1("LGS").Caption = "ﾌﾘｰｿﾘ" Then .Cells(n, 84) = "〆"
    
    '外観
    Set VisRange = .Range(.Cells(n, 85), .Cells(n, 85 + UBound(VisItem)))
    VisRange = VisItem
'    'BH
'    Set BhRange = .Range(.Cells(n, 90), .Cells(n, 90 + UBound(BHdata)))
'    BhRange = BHdata
    
    
    '歪AB面データ(上書き対応)***印刷を押したタイミングで表示されている歪の数値が記載されているのでAB対応済のもので上書き
    Dim hizS As Range: Set hizS = .Range(.Cells(n, 22), .Cells(n, 23))
    Dim hizU As Range: Set hizU = .Range(.Cells(n, 34), .Cells(n, 35))
    Dim hizC As Range: Set hizC = .Range(.Cells(n, 46), .Cells(n, 47))
    Dim hizD As Range: Set hizD = .Range(.Cells(n, 58), .Cells(n, 59))
    Dim hizE As Range: Set hizE = .Range(.Cells(n, 70), .Cells(n, 71))
    hizS = Array(hizumAB(1, 1), hizumAB(1, 2))
    hizU = Array(hizumAB(2, 1), hizumAB(2, 2))
    hizC = Array(hizumAB(3, 1), hizumAB(3, 2))
    hizD = Array(hizumAB(4, 1), hizumAB(4, 2))
    hizE = Array(hizumAB(5, 1), hizumAB(5, 2))
    
    'LoadingはUF1のラベル(Loadingの赤文字)
    '関数HistoryLoadで履歴を呼び出して印刷した場合のフラグ_2023.6.7*追加
    If UF1.dataLoding.Caption = "Load" Then .Cells(n, 82) = "〆"
    
    .Cells.Font.Name = "メイリオ"
    .Cells.Font.Size = 9
    .Cells.EntireColumn.AutoFit
    .Cells.FormulaHidden = False
'    .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'シート保護
  End With
End If


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>工場共有にDB追記_2021.10.19
'"\\nlmsrvngy03\工場内共有\検査データ\【測定DB】梱包\ファイルレイアウト変更前\2023.5.16" & "\梱包検査ﾃﾞｰﾀ.accdb"
Dim FPath As String: FPath = FactoryyPath & "【測定DB】梱包" & "\梱包検査ﾃﾞｰﾀ.accdb"
Dim EndR As Long
Dim EndC As Long
Dim AdoRange As Range
If UF1.極厚 Then
  EndR = UR(ws, 1)
  EndC = UC(ws, 1)
  Set AdoRange = ws.Range(ws.Cells(EndR, 1), ws.Cells(EndR, EndC))
Else
  EndR = UR(sh, 1)
  EndC = UC(sh, 1)
  Set AdoRange = sh.Range(sh.Cells(EndR, 1), sh.Cells(EndR, EndC))
End If

Dim AdoVar():      AdoVar = AdoRange
Dim Count As Long: Count = EndC '最終列

'If UF1.極厚 Then
'  Count = 108
'Else
'  Count = 99 '2023.11.27_+5
'End If

ReDim AdoTemp(0 To Count) As Variant

AdoTemp(0) = UF1.LineName.Caption & UF1.LineNo.Caption
Dim MinL As Long: MinL = LBound(AdoVar, 1)
Dim MaxL As Long: MaxL = UBound(AdoVar, 2)
For i = MinL To MaxL
  AdoTemp(i) = AdoVar(1, i)
Next i
'Stop

Dim LotNoT As String:    LotNoT = AdoTemp(1) '上書きしてしまっていたので修正　キーを重複しないものにした'並び→検査開始終了, LOTNO, 厚,,
Dim Fil As String:       Fil = "[検査開始終了] ='" & LotNoT & " '"

If Not UF1.極厚 Then Call adoSQL(FPath, "検査ﾃﾞｰﾀ", "ADDNEW", AdoTemp, "", True, Fil) '(fPath, "検査ﾃﾞｰﾀ", "DELETE", a, "ﾛｯﾄ番号") 'UPDATE,ADDNEW
If UF1.極厚 Then Call adoSQL(FPath, "耳付検査ﾃﾞｰﾀ", "ADDNEW", AdoTemp, "", True, Fil)
End Sub

