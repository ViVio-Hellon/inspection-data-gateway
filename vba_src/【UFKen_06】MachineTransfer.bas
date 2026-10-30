Attribute VB_Name = "【UFKen_06】MachineTransfer"
Option Explicit
'Q-S-QA-MSR-007*操業基準_検査.測定機器の使用基準
Public Sub Micro_Juge_旧(st As String) '使用ﾏｲｸﾛと操業基準ﾏｲｸﾛの比較 [Machin_Text]で実行********2021.09.22マイクロ選択が正しくない場合数値消去を追加
'UF1のラベルcaptionが引数
'機器ﾊﾟｹｯﾄ 機器交換時更新必要
If UF1.Machin = "" Then Exit Sub
Dim HM1 As String:  HM1 = "10001_M1" 'HVC_M12021.07.12交換
Dim HM2 As String:  HM2 = "10000_M2"
Dim HM3 As String:  HM3 = "10004_M3"
Dim HM4 As String:  HM4 = "10002_M4"
Dim HN5 As String:  HN5 = "10005_N5"
Dim LM1 As String:  LM1 = "10000_M1" 'LVC
Dim LM2 As String:  LM2 = "10001_M2"
Dim L1M1 As String: L1M1 = "10000_M1" 'L-1
Dim L1M2 As String: L1M2 = "10001_M2"
Dim L1M3 As String: L1M3 = "10003_M3"
'Q-S-QA-MSR-007-R04にて用途コードH494に限り幅測定をノギスで行う_2023.5.10
'------------------------------------------------------------
Dim L1N5 As String: L1N5 = "10005_ﾉｷﾞｽ" '2022.05.11_交換
Dim L1N6 As String: L1N6 = "10002_ﾉｷﾞｽ" '2022.05.19_交換
Dim BM2 As String:  BM2 = "10000_M2" 'ﾊﾞﾗﾝｻｰ
Dim TM1 As String:  TM1 = "11001_M1" 'ﾄｯﾄ　DT11001

Dim Pran As String: Pran = UF1.Machin.Caption '使用機器
Dim ATU As Single:  ATU = UF1.ATU '厚
Dim HABA As Single: HABA = UF1.Hab '幅
Dim Use As String '用途ｺｰﾄﾞ

Dim HistorySH As Worksheet: Set HistorySH = ThisWorkbook.Sheets("Lot情報履歴")
Dim LR As Long: LR = UR(HistorySH, 1)
Dim LC As Long: LC = UC(HistorySH, 1)
Dim HistoryR As Range: Set HistoryR = HistorySH.Range(HistorySH.Cells(1, 1), HistorySH.Cells(LR, LC))
Dim TempHiki(): TempHiki = HistoryR
Dim i As Long, x As Long, matchFlag As Boolean: matchFlag = False
Dim itti As Long
For i = 1 To UBound(TempHiki, 2) 'ﾌｨｰﾙﾄﾞ検索
  Select Case TempHiki(1, i)
  Case "用途ｺｰﾄﾞ": itti = i
  End Select
Next i
For i = 1 To UBound(TempHiki, 1)
  If TempHiki(i, 3) = UF1.Lot Then 'LOTNO一致
    x = i
    matchFlag = True
    Exit For
  End If
Next i
If matchFlag = False Then Exit Sub '---------------------------一致なし

Use = TempHiki(x, itti) '用途ｺｰﾄﾞ

Dim CountSt As String
If Len(st) = 3 Then
  CountSt = Right(st, 2)
Else
  CountSt = Right(st, 1)
End If

Dim Trm As String: Trm = "測定するﾏｲｸﾛを正しく選定して下さい"
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Select Case True
Case UF1.L1Work
'特殊用途コード
  If Use = "H494" And UF1.Hab < 300# Then  '用途コードH494に限り幅測定をノギスで行う
    Select Case HABA '幅寸法が範囲内の時
    Case 1# To 299#
      Select Case CountSt '幅の座標の時
      Case 5 To 6
        If Not Pran = L1N5 Then
          UF1(st) = "" 'ｸﾘｱ
          Call MsgError(("用途コード_" & Use & "の幅測定はノギスで行ってください") & " *Not:" & Lname & "N5", Lname, 1, "END")
        End If
      End Select
    End Select
    Exit Sub '抜ける
  End If

  If 0 < val(ATU) And val(ATU) <= 28# Then
   If Not Pran = L1M1 Then
    UF1(st) = "" 'ｸﾘｱ
    Call MsgError((Trm) & " *Not:" & Lname & "M1", Lname, 1, "END")
   End If
  ElseIf 28# < val(ATU) And val(ATU) <= 54# Then
   If Not Pran = L1M2 Then
    UF1(st) = "" 'ｸﾘｱ
    Call MsgError((Trm) & " *Not:" & Lname & "M2", Lname, 1, "END")
   End If
  ElseIf 54# < val(ATU) And val(ATU) <= 74# Then
   If Not Pran = L1M3 Then
    UF1(st) = "" 'ｸﾘｱ
    Call MsgError((Trm) & " *Not:" & Lname & "M3", Lname, 1, "END")
   End If
  ElseIf 74# < val(ATU) And val(ATU) < 300# Then 'ノギス
   If Not Pran = L1N5 Then
    UF1(st) = "" 'ｸﾘｱ
    Call MsgError((Trm) & " *Not:" & Lname & "N5", Lname, 1, "END")
   End If
  ElseIf 300# <= val(ATU) Then  'ノギス
   If Not Pran = L1N6 Then
    UF1(st) = "" 'ｸﾘｱ
    Call MsgError((Trm) & " *Not:" & Lname & "N6", Lname, 1, "END")
   End If
  End If
Case UF1.HVC
  If 0 < val(ATU) And val(ATU) <= 28# Then
   If (val(ATU) < 25#) And (Use = "H207" Or Use = "H191") Then '2022.07.14******************品証黒木そのままで良いとのこと　’<=　を<で25未満になり解決
    If Not Pran = HM4 Then
      UF1(st) = "" 'ｸﾘｱ
      Call MsgError((Trm) & " *Not:" & Lname & "M4" & "(M4は25mm以上の測定はできません。)", Lname, 1, "END")
    End If
   Else
    If Not Pran = HM1 Then
      UF1(st) = "" 'ｸﾘｱ
      Call MsgError((Trm) & " *Not:" & Lname & "M1", Lname, 1, "END")
    End If
   End If
  ElseIf 28# < val(ATU) And val(ATU) <= 54# Then
   If Not Pran = HM2 Then
    UF1(st) = "" 'ｸﾘｱ
    Call MsgError((Trm) & " *Not:" & Lname & "M2", Lname, 1, "END")
   End If
  ElseIf 54# < val(ATU) And val(ATU) <= 74# Then
   If Not Pran = HM3 Then
    UF1(st) = "" 'ｸﾘｱ
    Call MsgError((Trm) & " *Not:" & Lname & "M3", Lname, 1, "END")
   End If
  ElseIf 74# < val(ATU) And val(ATU) < 300# Then 'ノギス
   If Not Pran = HN5 Then
    UF1(st) = "" 'ｸﾘｱ
    Call MsgError((Trm) & " *Not:" & Lname & "N5", Lname, 1, "END")
   End If
  End If
Case UF1.LVC
  If 0 < val(ATU) And val(ATU) <= 25# Then
   If (val(ATU) <= 25#) And (Use = "H207" Or Use = "H191") Then
    If Not Pran = LM2 Then
      UF1(st) = "" 'ｸﾘｱ
      Call MsgError((Trm) & " *Not:" & Lname & "M2", Lname, 1, "END")
    End If
   Else
    If (Pran <> LM1) And (Pran <> LM2) Then
      UF1(st) = "" 'ｸﾘｱ
      Call MsgError((Trm) & " *Not:" & Lname & "M1" & "or" & Lname & "M2", Lname, 1, "END")
    End If
   End If
  End If
Case UF1.Bala
  If 0 < val(ATU) And val(ATU) <= 25# Then
    If Not Pran = BM2 Then
      UF1(st) = "" 'ｸﾘｱ
      Call MsgError((Trm) & " *Not:" & Lname & "M2", Lname, 1, "END")
    End If
  End If
Case UF1.Tott
  If 0 < val(ATU) And val(ATU) <= 25# Then
    If Not Pran = TM1 Then
      UF1(st) = "" 'ｸﾘｱ
      Call MsgError((Trm) & " *Not:" & Lname & "M1", Lname, 1, "END")
    End If
  End If
End Select
End Sub

Public Sub Micro_Juge(st As String) 'DB読み込みに変更_2024.2.9
'UF1のラベルcaptionが引数'機器ﾊﾟｹｯﾄ 機器交換時DB更新必要
If UF1.Machin = "" Then Exit Sub

Dim Pran As String: Pran = UF1.Machin.Caption '使用機器
Dim ATU As Single:  ATU = UF1.ATU '厚
Dim HABA As Single: HABA = UF1.Hab '幅
Dim Use As String:  Use = UF1.Youto '用途ｺｰﾄﾞ

'If UF1.LineName.Caption = "Test" Then Exit Sub

'ﾊﾞﾗﾝｻｰが半角になっている注意....直さねば...

Dim Trm As String:    Trm = "測定するﾏｲｸﾛを正しく選定して下さい"
Dim Lname As String:  Lname = UF1.LineName.Caption & UF1.LineNo.Caption
'ライン名、厚、用途コード、見たい列フィールド名、デプスゲージ
Dim PositiveS As String: PositiveS = DbReturnMicro(Lname, ATU, Use)         '出力Text
Dim Mdevice As String:   Mdevice = DbReturnMicro(Lname, ATU, Use, "機器名") '機器名出力

Dim CountSt As String
If Len(st) = 3 Then
  CountSt = Right(st, 2)
Else
  CountSt = Right(st, 1)
End If

'特殊用途コード
Select Case True
Case UF1.L1Work
  If Use = "H494" And UF1.Hab < 300# Then  '用途コードH494に限り幅測定をノギスで行う
    Select Case HABA '幅寸法が範囲内の時
    Case 1# To 299#
      Select Case CountSt '幅の座標の時
      Case 5 To 6
        If Not Pran = DbReturnMicro(Lname, ATU, Use, "", "N5") Then
          UF1(st) = "" 'ｸﾘｱ
          Call MsgError(("用途コード_" & Use & "の幅測定はノギスで行ってください") & " *Not:" & Lname & "N5", Lname, 1, "END")
        End If
      End Select
    End Select
    Exit Sub '抜ける
  End If
End Select

If Not Pran = PositiveS Then
 UF1(st) = "" 'ｸﾘｱ
 Call MsgError((Trm) & " " & Mdevice & "を使用の事", Lname, 1, "END")
End If

End Sub

Public Function Machin_Text(Trm As String) As String '測定機器選別
Dim Temp As String: Temp = Maker(Trm)
Machin_Text = ""
Select Case Temp
Case "DT"
 Machin_Text = Mid(Trm, 3, 5) 'DT 10002 +0000021.712M
Case Else
 Machin_Text = UF1.Machin
End Select

'******************************************************
'DBから読み込む方式に変更*2024.2.9
'数値部分のみ出力とする
'******************************************************

'変換_2021.01.20
'Select Case True
'Case UF1.L1Work 'L-1
'  Select Case Machin_Text
'  Case "10000": Machin_Text = "10000_M1"
'  Case "10001": Machin_Text = "10001_M2"
'  Case "10003": Machin_Text = "10003_M3"
''Q-S-QA-MSR-007-R04にて用途コードH494に限り幅測定をノギスで行う_2023.5.10
'  Case "10005": Machin_Text = "10005_ﾉｷﾞｽ" '2022.05.11_交換
'  Case "10002": Machin_Text = "10002_ﾉｷﾞｽ" '2022.05.19_交換
'  End Select
'Case UF1.LVC 'LVC
'  Select Case Machin_Text
'  Case "10000": Machin_Text = "10000_M1"
'  Case "10001": Machin_Text = "10001_M2"
'  End Select
'Case UF1.HVC 'HVC
'  Select Case Machin_Text
'  Case "10001": Machin_Text = "10001_M1" '2021.07.12交換DT10001+0000021.471M
'  Case "10000": Machin_Text = "10000_M2"
'  Case "10004": Machin_Text = "10004_M3"
'  Case "10002": Machin_Text = "10002_M4"
'  Case "10005": Machin_Text = "10005_N5"
'  End Select
'Case UF1.Bala 'ﾊﾞﾗﾝｻｰ　50320
'  Select Case Machin_Text
'  Case "10000": Machin_Text = "10000_M2"
'  End Select
'Case UF1.Tott 'ﾄｯﾄ　50319
'  Select Case Machin_Text
'  Case "11001": Machin_Text = "11001_M1"
'  End Select
'Case Else
'  Select Case Machin_Text
'  Case "10001": Machin_Text = "10001_Test"
'  Case "10000": Machin_Text = "10000_Test"
'  Case "10004": Machin_Text = "10004_Test"
'  Case "10002": Machin_Text = "10002_Test"
'  Case "10005": Machin_Text = "10005_N_Test"
'  Case "11001": Machin_Text = "11001_N_Test"
'  Case "11000": Machin_Text = "11000_Test"
'  End Select
'End Select
End Function

Public Sub 機器転送(ws As Worksheet, x As Long, y As Long, ATU As String, HABA As String, Optional Yout As String = "") 'ｼｰﾄ、座標行、座標列,板厚
Dim flag As Boolean, Flag2 As Boolean, Koflag As Boolean
Dim i As Long, a As Long
Dim line As String, Conb As String

With ws
  Select Case True
  Case UF1.L1Work
    line = "L-1"
    If val(UF1.Tak.Text) > 5500# Then Conb = " C11" '2024.11.25_コンベックスが壊れやすい2個もちにする
    If val(UF1.Tak.Text) <= 5500# Then Conb = " C12"
  Case UF1.HVC
    line = "HVC"
    Conb = " C11"
  Case UF1.LVC
    line = "LVC"
    Conb = " C10"
  Case UF1.Bala '50320
    line = "大板" '2024.2.8_ﾊﾞﾗﾝｻｰから変更
    Conb = " C12"
  Case UF1.Tott '50319
    line = "大板" '2024.2.8_ﾊﾞﾗﾝｻｰから変更
    Conb = " C11"
  End Select
 
 '追加修正_202411.18
 Dim AtuVal As Single:     AtuVal = val(ATU)
  Dim Mdevice As String:   Mdevice = DbReturnMicro(line, AtuVal, Yout, "機器名")  '機器名出力
 .Cells(x, y) = line & " " & Mdevice
 .Cells(x, y + 2) = line & Conb
 
    If val(ATU) = 300# Then
      For i = 14 To 18 '測定値のﾁｪｯｸ
        For a = 2 To 8 Step 2
          If IsNumeric(.Cells(i, a)) = True And (.Cells(i, a)) <> "" Then '数字
            If val(.Cells(i, a)) < 300# Then 'N5
             flag = True
            ElseIf val(.Cells(i, a)) >= 300# Then 'N6
             Koflag = True
            ElseIf (.Cells(i, a)) = "" Then '空白
            End If
          End If
        Next a
      Next i
      '書き込み内容
      If flag = True And Koflag = True Then
        .Cells(x, y) = line & " N5,N6"
      ElseIf flag = True And Koflag = False Then
        .Cells(x, y) = line & " N5"
      ElseIf Koflag = True And flag = False Then
        .Cells(x, y) = line & " N6"
      End If
    End If
 
 
'  If 0 < Val(Atu) And Val(Atu) <= 20# Then   'ライン外がL-1に変更2018.12.24
'     If (UF1.Youto = "H207" And line = "LVC") Then '用途分岐あり_2021.02.22　中精度,高精度時もM1になっていた為修正
'       .Cells(x, y) = line & " M2"
'     ElseIf (UF1.Youto = "H207" And line = "HVC") Then
'       .Cells(x, y) = line & " M4"
'     Else
'       If UF1.Bala Then 'ﾊﾞﾗﾝｻｰのみM2*2021.04.13
'        .Cells(x, y) = line & " M2"
'       Else '用途分岐なし
'        .Cells(x, y) = line & " M1"
'       End If
'     End If
'     .Cells(x, y + 2) = line & Conb
'  ElseIf 20 < Val(Atu) And Val(Atu) < 25 Then  '2021.02.22以降2022.07.14以前のデータはM1になっていると思われる
'     If (UF1.Youto = "H207" And line = "LVC") Then
'       .Cells(x, y) = line & " M2"
'     ElseIf (UF1.Youto = "H207" And line = "HVC") Then 'M4の使用限度が25未満ではあるが一旦続行_********************2022.07.14　’修正時Micro_Jugeも修正要
'       .Cells(x, y) = line & " M4"
'     Else '用途分岐なし
'       .Cells(x, y) = line & " M1"
'     End If
'     .Cells(x, y + 2) = line & Conb
'  ElseIf 25 <= Val(Atu) And Val(Atu) <= 28# Then
'    .Cells(x, y) = line & " M1"
'    .Cells(x, y + 2) = line & Conb
'  ElseIf 28 < Val(Atu) And Val(Atu) <= 54# Then
'    .Cells(x, y) = line & " M2"
'    .Cells(x, y + 2) = line & Conb
'  ElseIf 54# < Val(Atu) And Val(Atu) <= 74# Then
'    .Cells(x, y) = line & " M3"
'    .Cells(x, y + 2) = line & Conb
'  ElseIf 74# < Val(Atu) And Val(Atu) < 300 Then
'    .Cells(x - 1, y) = "ﾉｷﾞｽ" 'ノギス
'    .Cells(x, y) = line & " N5"
'    .Cells(x, y + 2) = line & Conb
'  ElseIf 300# <= Val(Atu) Then
'    .Cells(x - 1, y) = "ﾉｷﾞｽ" 'ノギス
'    .Cells(x, y) = line & " N6"
'    .Cells(x, y + 2) = line & Conb
'    '300
'    If Val(Atu) = 300# Then
'      For i = 14 To 18 '測定値のﾁｪｯｸ
'        For a = 2 To 8 Step 2
'          If IsNumeric(.Cells(i, a)) = True And (.Cells(i, a)) <> "" Then '数字
'            If Val(.Cells(i, a)) < 300# Then 'N5
'             flag = True
'            ElseIf Val(.Cells(i, a)) >= 300# Then 'N6
'             Koflag = True
'            ElseIf (.Cells(i, a)) = "" Then '空白
'            End If
'          End If
'        Next a
'      Next i
'      '書き込み内容
'      If flag = True And Koflag = True Then
'        .Cells(x, y) = line & " N5,N6"
'      ElseIf flag = True And Koflag = False Then
'        .Cells(x, y) = line & " N5"
'      ElseIf Koflag = True And flag = False Then
'        .Cells(x, y) = line & " N6"
'      End If
'    End If
'  End If
  
'特殊用途コード***2023.5.11
'最後に実行して通常分を書き換え
  If Yout = "H494" And UF1.Hab < 300# Then   '用途コードH494に限り幅測定をノギスで行う
    Select Case val(HABA) '幅寸法が範囲内の時
    Case 1# To 299#
      .Cells(x - 1, y + 2) = "幅:ﾉｷﾞｽ , 丈:ｺﾝﾍﾞｯｸｽ"
      .Cells(x, y + 2) = line & "N5," & Conb
    End Select
  End If
End With
End Sub
Public Function DbReturnMicro(LineName As String, ATU As Single, Yout As String, Optional n As String = "", Optional Dep As String = "") As String 'DBから出力テキストを返す
'ライン名、厚、用途コード、見たい列フィールド名、デプスゲージ
Dim FPath As String:         FPath = QualityPath '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\検査\"
Dim TableName As String:     TableName = "測定機器ID"
Dim target As String:        target = "伝送用ファイル.accdb"
Dim FilePath As String:      FilePath = FPath & target
Dim fieldsfilter As String:  fieldsfilter = ""
Dim PackF(): PackF = GetFieldsArr(FilePath, TableName, fieldsfilter)

If IsEmpty(PackF) Then
  FPath = PackingPath
  FilePath = FPath & target
  PackF = GetFieldsArr(FilePath, TableName, fieldsfilter)
  If IsEmpty(PackF) Then Exit Function
End If

Dim C_ライン As Long, C_機器名 As Long, C_ID As Long, C_出力テキスト As Long, C_厚下 As Long, C_厚上 As Long, C_用途 As Long, C_備考 As Long, i As Long
For i = LBound(PackF, 2) To UBound(PackF, 2)
  Select Case PackF(1, i)
  Case "ライン":       C_ライン = i
  Case "機器名":       C_機器名 = i
  Case "ID":           C_ID = i
  Case "出力テキスト": C_出力テキスト = i
  Case "厚下":         C_厚下 = i
  Case "厚上":         C_厚上 = i
  Case "用途":         C_用途 = i
  Case "備":           C_備考 = i
  End Select
Next i

Dim SQLfilter As String: SQLfilter = "[ライン]='" & LineName & "'"
Dim TempArr As Variant:  TempArr = GetRecordsArr(FilePath, TableName, SQLfilter)

If IsEmpty(TempArr) Then
  DbReturnMicro = ""
  Exit Function
End If
'機器指定ならここで終わり
If Dep <> "" Then 'M1,M2,M3,M4,N5,N6,デプス
  SQLfilter = "[ライン]='" & LineName & "' AND  [機器名]=" & " '" & Dep & "'"
  TempArr = GetRecordsArr(FilePath, TableName, SQLfilter)
  If IsEmpty(TempArr) Then
    DbReturnMicro = ""
    Exit Function
  End If
  DbReturnMicro = TempArr(LBound(TempArr, 1), C_出力テキスト)
  Exit Function
End If
'--------------------------------------------------------------------------------
Dim flag As Boolean: flag = False
For i = LBound(TempArr, 1) To UBound(TempArr, 1)
  If TempArr(i, C_用途) <> "" Then
    flag = True
    Exit For
  End If
Next i

Dim itti As Long: itti = 0

If flag Then
  For i = LBound(TempArr, 1) To UBound(TempArr, 1)
    If CheckThickAndCode(i, C_厚下, C_厚上, C_用途, ATU, Yout, TempArr) Then
      itti = i
      Exit For
    End If
  Next i
End If

Dim LowerT As Single, UpperT As Single

If itti = 0 Then '用途フラグあり　かつ　一致がなかった
  For i = LBound(TempArr, 1) To UBound(TempArr, 1)
    LowerT = val(TempArr(i, C_厚下))
    UpperT = val(TempArr(i, C_厚上))
    If (LowerT <= val(ATU) And UpperT >= val(ATU)) Then
      itti = i
      Exit For
    End If
  Next i
End If

If itti = 0 Then
  DbReturnMicro = ""
  Exit Function
End If

If itti <> 0 Then DbReturnMicro = TempArr(itti, C_出力テキスト)

If n <> "" Then
  Select Case n
  Case "機器名": DbReturnMicro = TempArr(itti, C_機器名)
  Case "ID":     DbReturnMicro = TempArr(itti, C_ID)
  End Select
End If

End Function
' 厚さの範囲と用途チェックする関数_測定器版
Function CheckThickAndCode(row As Long, Col1 As Long, Col2 As Long, Col3 As Long, OrderT As Variant, Yout As String, ArrQuality As Variant) As Boolean
'Row、厚下列番号、厚上列番号、用途列番号、板厚、用途、DB配列
Dim LowerT As Single: LowerT = val(ArrQuality(row, Col1))
Dim UpperT As Single: UpperT = val(ArrQuality(row, Col2))
Dim flag As Boolean:  flag = False
Dim itti As Long:     itti = 0
Dim i As Long
 
Dim FindInL As Boolean, TmpArray() As String, Item As Variant
Dim TmpString As String: TmpString = ArrQuality(row, Col3)

TmpArray = Split(TmpString, ",") ' Tmpを`,`で区切って配列にする

' 配列をループしてYoutを探す
For Each Item In TmpArray
  If Left(Item, 1) = "≠" Then ' ≠がついている場合
    If Mid(Item, 2) <> Yout Then ' ≠以降の文字列以外であればTrueを返す
      FindInL = True
      Exit For
    End If
  ElseIf Left(Item, 1) = "*" Then ' あたま*がついている場合
    If InStr(Yout, Mid(Item, 2)) > 0 Then ' *以降の文字列がYoutに含まれていればTrueを返す
      FindInL = True
      Exit For
    End If
  ElseIf Right(Item, 1) = "*" Then  ' *が右から1文字目についている場合
    If InStr(Yout, Left(Item, Len(Item) - 1)) > 0 Then ' *以前の文字列がYoutに含まれていればTrueを返す
      FindInL = True
      Exit For
    End If
  Else ' それ以外の場合はそのまま比較
    If Item = Yout Then
      FindInL = True
      Exit For
    End If
  End If
Next Item

If FindInL Then
  If (LowerT <= val(OrderT) And UpperT >= val(OrderT)) Then CheckThickAndCode = True
  Exit Function
End If
End Function

