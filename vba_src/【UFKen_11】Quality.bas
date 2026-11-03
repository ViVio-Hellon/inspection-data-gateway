Attribute VB_Name = "【UFKen_11】Quality"
Option Explicit

Public Sub 品質グレード()  '2019.10.03 '厚板
Dim ws As Worksheet, flag As Boolean, Flag2 As Boolean, Flag3 As Boolean
Dim Keyword As Double
Dim Zai As String, You As String, Tok As String, Nou As String, Oku As String, YouC As String
Dim To1 As String, To2 As String, To3 As String, To4 As String, To5 As String
Dim Ki1 As String, Ki2 As String, Ki3 As String, Ki4 As String, Ki5 As String
Dim Ha1 As String, Ha2 As String, Ha3 As String, Ha4 As String, Ha5 As String
Dim AL1 As String, AL2 As String, AL3 As String, AL4 As String, Al5 As String
Dim Bo1 As String, Bo2 As String, Bo3 As String, Bo4 As String, Bo5 As String
Dim Ka1 As String, Ka2 As String, Ka3 As String, Ka4 As String, Ka5 As String
Dim G1 As String, G2 As String, G3 As String, G4 As String, G5 As String

UF品質基準.Flag1 = ""

'-------------------------------------------------基準一覧　Q-S-QA-124-R02[別表1]　2019.10.02改訂版
G1 = "ｸﾞﾚｰﾄﾞ1"
G2 = "ｸﾞﾚｰﾄﾞ2"
G3 = "ｸﾞﾚｰﾄﾞ3"
G4 = "ｸﾞﾚｰﾄﾞ4"
G5 = "ｸﾞﾚｰﾄﾞ5"

To1 = "-" & " " & "(切削品)"
To2 = "*大きさ:3～5mmΦ" & " " & "*数:3個/枚以下"
To3 = "*大きさ:2～3mmΦ" & " " & "*数:3個/枚以下"
To4 = "*大きさ:1～2mmΦ" & " " & "*数:1個/枚以下"
To5 = "無き事"

Ki1 = "-"
Ki2 = "*触感:爪に「かかる」" & " " & "*大きさ:100mm以下"
Ki3 = "*触感:爪に「かかる」" & " " & "*大きさ:50mm以下"
Ki4 = "*触感:爪に「かからない」" & " " & "*大きさ:100mm以下"
Ki5 = "無き事"

Ha1 = "触感無き事(ﾌﾟﾘﾝﾄ)"
Ha2 = "*ﾌﾟﾘﾝﾄ" & " " & "*大きさ:200mm以下"
Ha3 = "*ﾌﾟﾘﾝﾄ" & " " & "*大きさ:50mm以下"
Ha4 = "*ﾌﾟﾘﾝﾄ" & " " & "*大きさ:20mm以下"
Ha5 = "無き事"

AL1 = "-"
AL2 = "*触感:あり" & " " & "*大きさ:300mm以下" & " " & "*数:2個/枚以下"
AL3 = "*触感:あり" & " " & "*大きさ:150mm以下" & " " & "*数:2個/枚以下"
AL4 = "*触感:あり" & " " & "*大きさ:50mm以下" & " " & "*数:1個/枚以下"
Al5 = "無き事"

Bo1 = "-"
Bo2 = "*大きさ:5～10mmΦ以下" & " " & "*数:3個/枚以下"
Bo3 = "*大きさ:3～5mmΦ以下" & " " & "*数:3個/枚以下"
Bo4 = "*大きさ:3～5mmΦ以下" & " " & "*数:1個/枚以下"
Bo5 = "無き事"

Ka1 = "-"
Ka2 = "*触感:あり" & " " & "*大きさ:200mm以下" & " " & "*数:5個/枚以下"
Ka3 = "*触感:あり" & " " & "*大きさ:100mm以下" & " " & "*数:3個/枚以下"
Ka4 = "*触感:あり" & " " & "*大きさ:100mm以下" & " " & "*数:1個/枚以下"
Ka5 = "触感無き事"
'--------------------------------------------------------
Application.ScreenUpdating = False
If WsCheck("Lot情報履歴") = False Then Exit Sub
Dim HistorySH As Worksheet: Set HistorySH = ThisWorkbook.Sheets("Lot情報履歴")
Dim LR As Long: LR = UR(HistorySH, 1)
Dim LC As Long: LC = UC(HistorySH, 1)
Dim HistoryR As Range: Set HistoryR = Range(HistorySH.Cells(2, 1), HistorySH.Cells(LR, LC))
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
  Keyword = UF1.ATU
  Zai = TempHiki(x, 8) '材質
  You = TempHiki(x, 10) '用途ｺｰﾄﾞ
  Tok = TempHiki(x, 12) '得意先
  Nou = TempHiki(x, 13) '納入先
  Oku = TempHiki(x, 14) '送り先
  YouC = TempHiki(x, 10) '用途コード
'用途.材質.板厚.行先(ﾅﾒｶﾜのみ)　グレード一覧　Q-S-QA-124-R02[別表2]　2019.10.02改訂版
  With UF品質基準
    '対象をフォーム転記
    .ATU = Format(val(TempHiki(x, 4)), "0.000")
    .HABA = Format(val(TempHiki(x, 5)), "0.0")
    .TAKE = Format(val(TempHiki(x, 6)), "0.0")
    .Zai = Zai
    .TYO = TempHiki(x, 9)
    .Yout = You
    .Toku = Tok
    .Noun = Nou
    .Okur = Oku
    .YoutC = YouC
 '----------------------------------------選別開始
    If Zai = "" Then 'データなし
    Else
      If You = "H191" Then '-------------------------------------------用途高精度
        If Zai = "52S" Or Zai = "F52S" Then
         Select Case Keyword
         Case 8# To 10#
          .ToG = G4
          .KiG = G4
          .HaG = G4
          .ALG = G4
          .BoG = G4
          .KaG = ""
          .Tobi = To4
          .Kizu = Ki4
          .Haku = Ha4
          .ALhu = AL4
          .Bota = Bo4
          .Kami = ""
         Case 12# To 30#
          .ToG = G4
          .KiG = G4
          .HaG = G4
          .ALG = G3
          .BoG = G3
          .KaG = ""
          .Tobi = To4
          .Kizu = Ki4
          .Haku = Ha4
          .ALhu = AL3
          .Bota = Bo3
          .Kami = ""
         Case 30.1 To 55#
          .ToG = G3
          .KiG = G3
          .HaG = G3
          .ALG = G2
          .BoG = G2
          .KaG = ""
          .Tobi = To3
          .Kizu = Ki3
          .Haku = Ha3
          .ALhu = AL2
          .Bota = Bo2
          .Kami = ""
         Case 55.1 To 75#
          .ToG = G2
          .KiG = G2
          .HaG = G1
          .ALG = G1
          .BoG = G1
          .KaG = ""
          .Tobi = To2
          .Kizu = Ki2
          .Haku = Ha1
          .ALhu = AL1
          .Bota = Bo1
          .Kami = ""
         Case 75.1 To 300
          .ToG = G1
          .KiG = G1
          .HaG = G1
          .ALG = G1
          .BoG = G1
          .KaG = ""
          .Tobi = To1
          .Kizu = Ki1
          .Haku = Ha1
          .ALhu = AL1
          .Bota = Bo1
          .Kami = ""
         End Select
        End If
      ElseIf You = "H207" Then '-------------------------------------------用途中精度
        If Zai = "52S" Or Zai = "F52S" Then
         Select Case Keyword
         Case 8# To 10#
          .ToG = G4
          .KiG = G4
          .HaG = G4
          .ALG = G4
          .BoG = G4
          .KaG = ""
          .Tobi = To4
          .Kizu = Ki4
          .Haku = Ha4
          .ALhu = AL4
          .Bota = Bo4
          .Kami = ""
         Case 12# To 30#
          .ToG = G4
          .KiG = G4
          .HaG = G4
          .ALG = G3
          .BoG = G3
          .KaG = ""
          .Tobi = To4
          .Kizu = Ki4
          .Haku = Ha4
          .ALhu = AL3
          .Bota = Bo3
          .Kami = ""
         Case 30.1 To 55#
          .ToG = G3
          .KiG = G3
          .HaG = G3
          .ALG = G2
          .BoG = G2
          .KaG = ""
          .Tobi = To3
          .Kizu = Ki3
          .Haku = Ha3
          .ALhu = AL2
          .Bota = Bo2
          .Kami = ""
         Case 55.1 To 75#
          .ToG = G2
          .KiG = G2
          .HaG = G1
          .ALG = G1
          .BoG = G1
          .KaG = ""
          .Tobi = To2
          .Kizu = Ki2
          .Haku = Ha1
          .ALhu = AL1
          .Bota = Bo1
          .Kami = ""
         Case 75.1 To 300
          .ToG = G1
          .KiG = G1
          .HaG = G1
          .ALG = G1
          .BoG = G1
          .KaG = ""
          .Tobi = To1
          .Kizu = Ki1
          .Haku = Ha1
          .ALhu = AL1
          .Bota = Bo1
          .Kami = ""
         End Select
        End If
      ElseIf You = "H197" Then '-------------------------------------------用途シャーシ　シャーシのみ行先分岐有
       If Tok = "ﾅﾒｶﾜｱﾙﾐ(ｶ" Or Nou = "ﾅﾒｶﾜｱﾙﾐ(ｶ" Or Oku = "ﾅﾒｶﾜｱﾙﾐ(ｶ" Or Oku = "ﾅﾒｶﾜｱﾙﾐ(ｶ ｲﾀｺｺｳｼﾞﾖｳ" Then
       '確認用フラグ表示
       UF品質基準.Flag1 = "ﾅﾒｶﾜｱﾙﾐ"
        If Zai = "52S" Or Zai = "F52S" Or Zai = "NP56" Or Zai = "14S" Or Zai = "17S" Or Zai = "N17S" Or Zai = "61S" Or Zai = "R61S" Then
         Select Case Keyword
         Case 8# To 10#
          .ToG = G4
          .KiG = G4
          .HaG = G4
          .ALG = G4
          .BoG = G4
          .KaG = G5
          .Tobi = To4
          .Kizu = Ki4
          .Haku = Ha4
          .ALhu = AL4
          .Bota = Bo4
          .Kami = Ka5
         Case 12# To 30#
          .ToG = G4
          .KiG = G4
          .HaG = G3
          .ALG = G3
          .BoG = G3
          .KaG = G5
          .Tobi = To4
          .Kizu = Ki4
          .Haku = Ha3
          .ALhu = AL3
          .Bota = Bo3
          .Kami = Ka5
         Case 30.1 To 55#
          .ToG = G4
          .KiG = G4
          .HaG = G3
          .ALG = G3
          .BoG = G2
          .KaG = G5
          .Tobi = To4
          .Kizu = Ki4
          .Haku = Ha3
          .ALhu = AL3
          .Bota = Bo2
          .Kami = Ka5
         Case 55.1 To 300
          .ToG = G4
          .KiG = G4
          .HaG = G3
          .ALG = G3
          .BoG = G2
          .KaG = G5
          .Tobi = To4
          .Kizu = Ki4
          .Haku = Ha3
          .ALhu = AL3
          .Bota = Bo2
          .Kami = Ka5
         End Select
        End If
       Else
        If Zai = "52S" Or Zai = "F52S" Or Zai = "NP56" Or Zai = "14S" Or Zai = "17S" Or Zai = "N17S" Or Zai = "61S" Or Zai = "R61S" Then
         Select Case Keyword
         Case 8# To 10#
          .ToG = G3
          .KiG = G3
          .HaG = G4
          .ALG = G4
          .BoG = G4
          .KaG = ""
          .Tobi = To3
          .Kizu = Ki3
          .Haku = Ha4
          .ALhu = AL4
          .Bota = Bo4
          .Kami = ""
         Case 12# To 30#
          .ToG = G3
          .KiG = G3
          .HaG = G3
          .ALG = G3
          .BoG = G3
          .KaG = ""
          .Tobi = To3
          .Kizu = Ki3
          .Haku = Ha3
          .ALhu = AL3
          .Bota = Bo3
          .Kami = ""
         Case 30.1 To 55#
          .ToG = G2
          .KiG = G2
          .HaG = G2
          .ALG = G2
          .BoG = G2
          .KaG = ""
          .Tobi = To2
          .Kizu = Ki2
          .Haku = Ha2
          .ALhu = AL2
          .Bota = Bo2
          .Kami = ""
         Case 55.1 To 75#
          .ToG = G1
          .KiG = G1
          .HaG = G1
          .ALG = G1
          .BoG = G1
          .KaG = ""
          .Tobi = To1
          .Kizu = Ki1
          .Haku = Ha1
          .ALhu = AL1
          .Bota = Bo1
          .Kami = ""
         Case 75.1 To 300
          .ToG = G1
          .KiG = G1
          .HaG = G1
          .ALG = G1
          .BoG = G1
          .KaG = ""
          .Tobi = To1
          .Kizu = Ki1
          .Haku = Ha1
          .ALhu = AL1
          .Bota = Bo1
          .Kami = ""
         End Select
        End If
       End If
      ElseIf You = "K21S" Then '-------------------------------------------鉄道ボディ
        If Zai = "U74S" Or Zai = "NP56" Or Zai = "75S" Then
         Select Case Keyword
         Case 8# To 10#
          .ToG = G2
          .KiG = G2
          .HaG = G2
          .ALG = G1
          .BoG = G2
          .KaG = ""
          .Tobi = To2
          .Kizu = Ki2
          .Haku = Ha2
          .ALhu = AL1
          .Bota = Bo2
          .Kami = ""
         Case 12# To 30#
          .ToG = G2
          .KiG = G2
          .HaG = G2
          .ALG = G1
          .BoG = G2
          .KaG = ""
          .Tobi = To2
          .Kizu = Ki2
          .Haku = Ha2
          .ALhu = AL1
          .Bota = Bo2
          .Kami = ""
         Case 30.1 To 55#
          .ToG = G2
          .KiG = G2
          .HaG = G2
          .ALG = G1
          .BoG = G2
          .KaG = ""
          .Tobi = To2
          .Kizu = Ki2
          .Haku = Ha2
          .ALhu = AL1
          .Bota = Bo2
          .Kami = ""
         Case 55.1 To 75#
          .ToG = G1
          .KiG = G1
          .HaG = G1
          .ALG = G1
          .BoG = G1
          .KaG = ""
          .Tobi = To1
          .Kizu = Ki1
          .Haku = Ha1
          .ALhu = AL1
          .Bota = Bo1
          .Kami = ""
         Case 75.1 To 300
          .ToG = G1
          .KiG = G1
          .HaG = G1
          .ALG = G1
          .BoG = G1
          .KaG = ""
          .Tobi = To1
          .Kizu = Ki1
          .Haku = Ha1
          .ALhu = AL1
          .Bota = Bo1
          .Kami = ""
         End Select
        End If
      End If
    End If
  End With
Application.ScreenUpdating = True
End Sub
Public Sub 平坦度(ATU As Double, HABA As Double, TAKE As Double, Zai As String, Tyu As String, YoutC As String)    '2019.10.04　厚板’2023.1.31引数追加_VCなしでも公差を拾える場合は拾う
'2022.02鉄道ボディー  厚120~200追加
'2022.05.27_Q-S-QA-122-R05 平坦度追加更新
Dim ws As Worksheet, flag As Boolean, Flag2 As Boolean, Flag3 As Boolean
'Dim ATU As Double, HABA As Double, TAKE As Double
Dim You As String, Tok As String, Nou As String, Oku As String
'Dim Zai As String, You As String, Tok As String, Nou As String, Oku As String, YoutC As String
Dim VCA As String, VCB As String '作業指示_QA-022-01用*2022.07.05
Application.ScreenUpdating = False
UF品質基準.Flag2 = ""

If WsCheck("Lot情報履歴") = False Then Exit Sub
Dim HistorySH As Worksheet: Set HistorySH = ThisWorkbook.Sheets("Lot情報履歴")
Dim LR As Long: LR = UR(HistorySH, 1)
Dim LC As Long: LC = UC(HistorySH, 1)
Dim HistoryR As Range: Set HistoryR = Range(HistorySH.Cells(2, 1), HistorySH.Cells(LR, LC))
Dim TempHiki(): TempHiki = HistoryR
Dim i As Long, x As Long, matchFlag As Boolean: matchFlag = False
For i = 1 To UBound(TempHiki, 1)
  If TempHiki(i, 3) = UF1.Lot Then 'LOTNO一致
    x = i
    matchFlag = True
    Exit For
  End If
Next i

'VCデータが必要な歪み基準がある為引当データが必要、引当データがない場合寸法と材調質だけで基準を拾える場合は拾う
If matchFlag Then  '引当データがある場合はLot情報履歴からデータをとって上書き
  If IsEmpty(TempHiki) Then Exit Sub
  ATU = Format(val(TempHiki(x, 4)), "0.000")
  HABA = Format(val(TempHiki(x, 5)), "0.0")
  TAKE = Format(val(TempHiki(x, 6)), "0.0")
  Zai = TempHiki(x, 8)
  You = TempHiki(x, 10)
  Tok = TempHiki(x, 11)
  Nou = TempHiki(x, 12)
  Oku = TempHiki(x, 13)
  YoutC = TempHiki(x, 10)
  VCA = TempHiki(x, 19)
  VCB = TempHiki(x, 20)
  
  '作業指示_QA-022-01用ﾌﾗｸﾞ設定**2022.07.05
  Dim VcFlag As Boolean: VcFlag = False
  If (VCA = "208PRNF" Or VCA = "V325NNF" Or VCA = "V311PNF") Or _
  (VCB = "208PRNF" Or VCB = "V325NNF" Or VCB = "V311PNF") Then VcFlag = True
End If
  
  
 '----------------------------------------厚板材.Q-S-QA-122-R04[別表]
 '2022.05.27_Q-S-QA-122-R05 平坦度追加更新
 '中歪耳歪については厚6mm以下限定
  With UF品質基準
    If Zai = "" Then 'データなし
    Else
      If YoutC = "H191" Or YoutC = "H242" Then  '-------------------------------------------NFLAT-FINE
        If Zai = "52S" Or Zai = "NP56" Or Zai = "F52S" Or Zai = "5083" Then '----------------------材質限定
         Select Case ATU
         Case 6# To 50#
          .MG1 = "0.2以下" '--------------------------------ｹﾞｰｼﾞ
          If ATU <= 6# Then .Naka1 = Format(中歪(ATU, HABA), "0.0") '--------------------------------中歪
          If ATU <= 6# Then .Mimi1 = Format(耳歪(ATU, HABA), "0.0")  '--------------------------------耳歪
          If 0 < HABA And HABA <= 500# Then '-----------------------巾歪
          .HaHiz1 = "0.1"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "0.2"
          ElseIf 750# < HABA And HABA <= 1004# Then
          .HaHiz1 = "0.2"
          ElseIf 1004# < HABA And HABA <= 1254# Then
          .HaHiz1 = "0.3"
          ElseIf 1254# < HABA And HABA <= 1529# Then
          .HaHiz1 = "0.3"
          ElseIf 1529# < HABA And HABA <= 2004# Then   '2022.05.23_巾2005~3055追加
          .HaHiz1 = "0.4"
          ElseIf 2004# < HABA And HABA <= 2504# Then
          .HaHiz1 = "0.4"
          ElseIf 2504# < HABA And HABA <= 3054# Then
          .HaHiz1 = "0.4"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "0.1"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "0.2"
          ElseIf 750# < TAKE And TAKE <= 1004# Then
          .TaHiz1 = "0.2"
          ElseIf 1004# < TAKE And TAKE <= 1254# Then
          .TaHiz1 = "0.3"
          ElseIf 1254# < TAKE And TAKE <= 1529# Then
          .TaHiz1 = "0.3"
          ElseIf 1529# < TAKE And TAKE <= 2004# Then
          .TaHiz1 = "0.4"
          ElseIf 2004# < TAKE And TAKE <= 2504# Then
          .TaHiz1 = "0.4"
          ElseIf 2504# < TAKE And TAKE <= 3054# Then
          .TaHiz1 = "0.4"
          ElseIf 3054# < TAKE Then
          .TaHiz1 = "0.4"
          End If
         End Select
        End If
      ElseIf YoutC = "H196" Or YoutC = "H299" Or _
      YoutC = "H207" Or YoutC = "H234" Or _
      YoutC = "H237" Or YoutC = "H238" Or _
      YoutC = "H239" Or YoutC = "H241" Or _
      YoutC = "H272" Or YoutC = "H255" Then '-------------------------------------------NFLAT納入仕様書
        If Zai = "52S" Or Zai = "NP56" Or Zai = "F52S" Or Zai = "5083" Then '----------------------材質限定
         Select Case ATU
         Case 4# To 11.999
          .MG1 = "0.5以下" '--------------------------------ｹﾞｰｼﾞ
          If ATU <= 6# Then .Naka1 = Format(中歪(ATU, HABA), "0.0") '--------------------------------中歪
          If ATU <= 6# Then .Mimi1 = Format(耳歪(ATU, HABA), "0.0")  '--------------------------------耳歪
          If 0 < HABA And HABA <= 500# Then '-----------------------巾歪
          .HaHiz1 = "0.3"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "0.4"
          ElseIf 750# < HABA And HABA <= 1004# Then
          .HaHiz1 = "0.5"
          ElseIf 1004# < HABA And HABA <= 1254# Then
          .HaHiz1 = "0.6"
          ElseIf 1254# < HABA And HABA <= 1529# Then
          .HaHiz1 = "0.8"
          ElseIf 1529# < HABA And HABA <= 2004# Then   '2022.05.23_巾2005~3055追加
          .HaHiz1 = "1.0"
          ElseIf 2004# < HABA And HABA <= 2504# Then
          .HaHiz1 = "1.0"
          ElseIf 2504# < HABA And HABA <= 3054# Then
          .HaHiz1 = "1.0"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "0.3"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "0.4"
          ElseIf 750# < TAKE And TAKE <= 1004# Then
          .TaHiz1 = "0.5"
          ElseIf 1004# < TAKE And TAKE <= 1254# Then
          .TaHiz1 = "0.6"
          ElseIf 1254# < TAKE And TAKE <= 1529# Then
          .TaHiz1 = "0.8"
          ElseIf 1529# < TAKE And TAKE <= 2004# Then
          .TaHiz1 = "1.0"
          ElseIf 2004# < TAKE And TAKE <= 2504# Then
          .TaHiz1 = "1.0"
          ElseIf 2504# < TAKE And TAKE <= 3054# Then
          .TaHiz1 = "1.0"
          ElseIf 3054# < TAKE Then
          .TaHiz1 = "1.0"
          End If
         Case 12# To 50#
          .MG1 = "0.4以下" '--------------------------------ｹﾞｰｼﾞ
          If 0 < HABA And HABA <= 500# Then  '-----------------------巾歪
          .HaHiz1 = "0.2"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "0.3"
          ElseIf 750# < HABA And HABA <= 1004# Then
          .HaHiz1 = "0.4"
          ElseIf 1004# < HABA And HABA <= 1254# Then
          .HaHiz1 = "0.5"
          ElseIf 1254# < HABA And HABA <= 1529# Then
          .HaHiz1 = "0.6"
          ElseIf 1529# < HABA And HABA <= 2004# Then   '2022.05.23_巾2005~3055追加
          .HaHiz1 = "0.7"
          ElseIf 2004# < HABA And HABA <= 2504# Then
          .HaHiz1 = "0.7"
          ElseIf 2504# < HABA And HABA <= 3054# Then
          .HaHiz1 = "0.7"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then  '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "0.2"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "0.3"
          ElseIf 750# < TAKE And TAKE <= 1004# Then
          .TaHiz1 = "0.4"
          ElseIf 1004# < TAKE And TAKE <= 1254# Then
          .TaHiz1 = "0.5"
          ElseIf 1254# < TAKE And TAKE <= 1529# Then
          .TaHiz1 = "0.6"
          ElseIf 1529# < TAKE And TAKE <= 2004# Then
          .TaHiz1 = "0.7"
          ElseIf 2004# < TAKE And TAKE <= 2504# Then
          .TaHiz1 = "0.7"
          ElseIf 2504# < TAKE And TAKE <= 3054# Then
          .TaHiz1 = "0.7"
          ElseIf 3054# < TAKE Then
          .TaHiz1 = "0.7"
          End If
         Case 50.1 To 100#
          .MG1 = "1.2以下" '--------------------------------ｹﾞｰｼﾞ
          If 0 < HABA And HABA <= 500# Then '-----------------------巾歪
          .HaHiz1 = "0.6"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "0.9"
          ElseIf 750# < HABA And HABA <= 1004# Then
          .HaHiz1 = "1.2"
          ElseIf 1004# < HABA And HABA <= 1254# Then
          .HaHiz1 = "1.5"
          ElseIf 1254# < HABA And HABA <= 1529# Then
          .HaHiz1 = "2.0"
          ElseIf 1529# < HABA And HABA <= 2004# Then   '2022.05.23_巾2005~3055追加
          .HaHiz1 = "2.5"
          ElseIf 2004# < HABA And HABA <= 2504# Then
          .HaHiz1 = "3.0"
          ElseIf 2504# < HABA And HABA <= 3054# Then
          .HaHiz1 = "3.5"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then  '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "0.6"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "0.9"
          ElseIf 750# < TAKE And TAKE <= 1004# Then
          .TaHiz1 = "1.2"
          ElseIf 1004# < TAKE And TAKE <= 1254# Then
          .TaHiz1 = "1.5"
          ElseIf 1254# < TAKE And TAKE <= 1529# Then
          .TaHiz1 = "2.0"
          ElseIf 1529# < TAKE And TAKE <= 2004# Then
          .TaHiz1 = "2.5"
          ElseIf 2004# < TAKE And TAKE <= 2504# Then
          .TaHiz1 = "3.0"
          ElseIf 2504# < TAKE And TAKE <= 3054# Then
          .TaHiz1 = "3.5"
          ElseIf 3054# < TAKE Then
          .TaHiz1 = "3.5"
          End If
         End Select
        End If
      ElseIf YoutC = "H205" Or YoutC = "H251" Then  '-------------------------------------------NFLATカタログ_2022.05.27
        If Zai = "52S" Or Zai = "NP56" Or Zai = "F52S" Or Zai = "5083" Then '----------------------材質限定
         Select Case ATU
         Case 4# To 11.999
          .MG1 = "0.5以下" '--------------------------------ｹﾞｰｼﾞ
          If ATU <= 6# Then .Naka1 = Format(中歪(ATU, HABA), "0.0") '--------------------------------中歪
          If ATU <= 6# Then .Mimi1 = Format(耳歪(ATU, HABA), "0.0") '--------------------------------耳歪
          If 0 < HABA And HABA <= 500# Then '-----------------------巾歪
          .HaHiz1 = "0.3"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "0.4"
          ElseIf 750# < HABA And HABA <= 1004# Then
          .HaHiz1 = "0.5"
          ElseIf 1004# < HABA And HABA <= 1254# Then
          .HaHiz1 = "0.6"
          ElseIf 1254# < HABA And HABA <= 1529# Then
          .HaHiz1 = "0.8"
          ElseIf 1529# < HABA And HABA <= 2004# Then
          .HaHiz1 = "1.0"
          ElseIf 2004# < HABA And HABA <= 2504# Then
          .HaHiz1 = "1.0"
          ElseIf 2504# < HABA And HABA <= 3054# Then
          .HaHiz1 = "1.0"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪
          .TaHiz1 = "0.3"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "0.4"
          ElseIf 750# < TAKE And TAKE <= 1004# Then
          .TaHiz1 = "0.5"
          ElseIf 1004# < TAKE And TAKE <= 1254# Then
          .TaHiz1 = "0.6"
          ElseIf 1254# < TAKE And TAKE <= 1529# Then
          .TaHiz1 = "0.8"
          ElseIf 1529# < TAKE And TAKE <= 2004# Then
          .TaHiz1 = "1.0"
          ElseIf 2004# < TAKE And TAKE <= 2504# Then
          .TaHiz1 = "1.0"
          ElseIf 2504# < TAKE And TAKE <= 3054# Then
          .TaHiz1 = "1.0"
          ElseIf 3054# < TAKE Then
          .TaHiz1 = "1.0"
          End If
         Case 12# To 30#
          .MG1 = "0.4以下" '--------------------------------ｹﾞｰｼﾞ
          If 0 < HABA And HABA <= 500# Then  '-----------------------巾歪
          .HaHiz1 = "0.2"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "0.3"
          ElseIf 750# < HABA And HABA <= 1004# Then
          .HaHiz1 = "0.4"
          ElseIf 1004# < HABA And HABA <= 1254# Then
          .HaHiz1 = "0.5"
          ElseIf 1254# < HABA And HABA <= 1529# Then
          .HaHiz1 = "0.6"
          ElseIf 1529# < HABA And HABA <= 2004# Then
          .HaHiz1 = "0.7"
          ElseIf 2004# < HABA And HABA <= 2504# Then
          .HaHiz1 = "0.7"
          ElseIf 2504# < HABA And HABA <= 3054# Then
          .HaHiz1 = "0.7"
          End If
          If 0 < TAKE And TAKE <= 500# Then  '----------------------丈歪
          .TaHiz1 = "0.2"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "0.3"
          ElseIf 750# < TAKE And TAKE <= 1004# Then
          .TaHiz1 = "0.4"
          ElseIf 1004# < TAKE And TAKE <= 1254# Then
          .TaHiz1 = "0.5"
          ElseIf 1254# < TAKE And TAKE <= 1529# Then
          .TaHiz1 = "0.6"
          ElseIf 1529# < TAKE And TAKE <= 2004# Then
          .TaHiz1 = "0.7"
          ElseIf 2004# < TAKE And TAKE <= 2504# Then
          .TaHiz1 = "0.7"
          ElseIf 2504# < TAKE And TAKE <= 3054# Then
          .TaHiz1 = "0.7"
          ElseIf 3054# < TAKE Then
          .TaHiz1 = "0.7"
          End If
         Case 30.1 To 50#
          .MG1 = "1.0以下" '--------------------------------ｹﾞｰｼﾞ
          If 0 < HABA And HABA <= 500# Then '-----------------------巾歪
          .HaHiz1 = "0.5"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "0.8"
          ElseIf 750# < HABA And HABA <= 1004# Then
          .HaHiz1 = "1.0"
          ElseIf 1004# < HABA And HABA <= 1254# Then
          .HaHiz1 = "1.3"
          ElseIf 1254# < HABA And HABA <= 1529# Then
          .HaHiz1 = "1.8"
          ElseIf 1529# < HABA And HABA <= 2004# Then
          .HaHiz1 = "2.0"
          ElseIf 2004# < HABA And HABA <= 2504# Then
          .HaHiz1 = "2.0"
          ElseIf 2504# < HABA And HABA <= 3054# Then
          .HaHiz1 = "2.0"
          End If
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪
          .TaHiz1 = "0.5"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "0.8"
          ElseIf 750# < TAKE And TAKE <= 1004# Then
          .TaHiz1 = "1.0"
          ElseIf 1004# < TAKE And TAKE <= 1254# Then
          .TaHiz1 = "1.3"
          ElseIf 1254# < TAKE And TAKE <= 1529# Then
          .TaHiz1 = "1.8"
          ElseIf 1529# < TAKE And TAKE <= 2004# Then
          .TaHiz1 = "2.0"
          ElseIf 2004# < TAKE And TAKE <= 2504# Then
          .TaHiz1 = "2.0"
          ElseIf 2504# < TAKE And TAKE <= 3054# Then
          .TaHiz1 = "2.0"
          ElseIf 3054# < TAKE Then
          .TaHiz1 = "2.0"
          End If
         Case 50.1 To 100#
          .MG1 = "1.2以下" '--------------------------------ｹﾞｰｼﾞ
          If 0 < HABA And HABA <= 500# Then '-----------------------巾歪
          .HaHiz1 = "0.6"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "0.9"
          ElseIf 750# < HABA And HABA <= 1004# Then
          .HaHiz1 = "1.2"
          ElseIf 1004# < HABA And HABA <= 1254# Then
          .HaHiz1 = "1.5"
          ElseIf 1254# < HABA And HABA <= 1529# Then
          .HaHiz1 = "2.0"
          ElseIf 1529# < HABA And HABA <= 2004# Then
          .HaHiz1 = "2.5"
          ElseIf 2004# < HABA And HABA <= 2504# Then
          .HaHiz1 = "3.0"
          ElseIf 2504# < HABA And HABA <= 3054# Then
          .HaHiz1 = "3.5"
          End If
          If 0 < TAKE And TAKE <= 500# Then  '----------------------丈歪
          .TaHiz1 = "0.6"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "0.9"
          ElseIf 750# < TAKE And TAKE <= 1004# Then
          .TaHiz1 = "1.2"
          ElseIf 1004# < TAKE And TAKE <= 1254# Then
          .TaHiz1 = "1.5"
          ElseIf 1254# < TAKE And TAKE <= 1529# Then
          .TaHiz1 = "2.0"
          ElseIf 1529# < TAKE And TAKE <= 2004# Then
          .TaHiz1 = "2.5"
          ElseIf 2004# < TAKE And TAKE <= 2504# Then
          .TaHiz1 = "3.0"
          ElseIf 2504# < TAKE And TAKE <= 3054# Then
          .TaHiz1 = "3.5"
          ElseIf 3054# < TAKE Then
          .TaHiz1 = "3.5"
          End If
         End Select
        End If
      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> QA-022-01
      '用途コードH197かつ指定のビニール使用時
      ElseIf YoutC = "H197" And VcFlag Then
      UF品質基準.Flag2 = "H197 and 指定VC "
        Select Case ATU '4㎜以上100㎜以下
        Case 4# To 11.999
          .MG1 = "0.5以下" '--------------------------------ｹﾞｰｼﾞ
          If ATU <= 6# Then .Naka1 = Format(中歪(ATU, HABA), "0.0") '--------------------------------中歪
          If ATU <= 6# Then .Mimi1 = Format(耳歪(ATU, HABA), "0.0")  '--------------------------------耳歪
          If 0 < HABA And HABA <= 500# Then '-----------------------巾歪
            .HaHiz1 = "0.3"
          ElseIf 500# < HABA And HABA <= 750# Then
            .HaHiz1 = "0.4"
          ElseIf 750# < HABA And HABA <= 1004# Then
            .HaHiz1 = "0.5"
          ElseIf 1004# < HABA And HABA <= 1254# Then
            .HaHiz1 = "0.6"
          ElseIf 1254# < HABA And HABA <= 1529# Then
            .HaHiz1 = "0.8"
          ElseIf 1529# < HABA And HABA <= 2004# Then   '2022.05.23_巾2005~3055追加
            .HaHiz1 = "1.0"
          ElseIf 2004# < HABA And HABA <= 2504# Then
            .HaHiz1 = "1.0"
          ElseIf 2504# < HABA And HABA <= 3054# Then
            .HaHiz1 = "1.0"
          End If
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
            .TaHiz1 = "0.3"
          ElseIf 500# < TAKE And TAKE <= 750# Then
            .TaHiz1 = "0.4"
          ElseIf 750# < TAKE And TAKE <= 1004# Then
            .TaHiz1 = "0.5"
          ElseIf 1004# < TAKE And TAKE <= 1254# Then
            .TaHiz1 = "0.6"
          ElseIf 1254# < TAKE And TAKE <= 1529# Then
            .TaHiz1 = "0.8"
          ElseIf 1529# < TAKE And TAKE <= 2004# Then
            .TaHiz1 = "1.0"
          ElseIf 2004# < TAKE And TAKE <= 2504# Then
            .TaHiz1 = "1.0"
          ElseIf 2504# < TAKE And TAKE <= 3054# Then
            .TaHiz1 = "1.0"
          ElseIf 3054# < TAKE Then
            .TaHiz1 = "1.0"
          End If
        Case 12# To 50#
          .MG1 = "0.4以下" '--------------------------------ｹﾞｰｼﾞ
          If 0 < HABA And HABA <= 500# Then  '-----------------------巾歪
            .HaHiz1 = "0.2"
          ElseIf 500# < HABA And HABA <= 750# Then
            .HaHiz1 = "0.3"
          ElseIf 750# < HABA And HABA <= 1004# Then
            .HaHiz1 = "0.4"
          ElseIf 1004# < HABA And HABA <= 1254# Then
            .HaHiz1 = "0.5"
          ElseIf 1254# < HABA And HABA <= 1529# Then
            .HaHiz1 = "0.6"
          ElseIf 1529# < HABA And HABA <= 2004# Then   '2022.05.23_巾2005~3055追加
            .HaHiz1 = "0.7"
          ElseIf 2004# < HABA And HABA <= 2504# Then
            .HaHiz1 = "0.7"
          ElseIf 2504# < HABA And HABA <= 3054# Then
            .HaHiz1 = "0.7"
          End If
          If 0 < TAKE And TAKE <= 500# Then  '----------------------丈歪'2022.05.23_丈500~1530追加
            .TaHiz1 = "0.2"
          ElseIf 500# < TAKE And TAKE <= 750# Then
            .TaHiz1 = "0.3"
          ElseIf 750# < TAKE And TAKE <= 1004# Then
            .TaHiz1 = "0.4"
          ElseIf 1004# < TAKE And TAKE <= 1254# Then
            .TaHiz1 = "0.5"
          ElseIf 1254# < TAKE And TAKE <= 1529# Then
            .TaHiz1 = "0.6"
          ElseIf 1529# < TAKE And TAKE <= 2004# Then
            .TaHiz1 = "0.7"
          ElseIf 2004# < TAKE And TAKE <= 2504# Then
            .TaHiz1 = "0.7"
          ElseIf 2504# < TAKE And TAKE <= 3054# Then
            .TaHiz1 = "0.7"
          ElseIf 3054# < TAKE Then
            .TaHiz1 = "0.7"
          End If
        Case 50.1 To 100#
          .MG1 = "1.2以下" '--------------------------------ｹﾞｰｼﾞ
          If 0 < HABA And HABA <= 500# Then '-----------------------巾歪
            .HaHiz1 = "0.6"
          ElseIf 500# < HABA And HABA <= 750# Then
            .HaHiz1 = "0.9"
          ElseIf 750# < HABA And HABA <= 1004# Then
            .HaHiz1 = "1.2"
          ElseIf 1004# < HABA And HABA <= 1254# Then
            .HaHiz1 = "1.5"
          ElseIf 1254# < HABA And HABA <= 1529# Then
            .HaHiz1 = "2.0"
          ElseIf 1529# < HABA And HABA <= 2004# Then   '2022.05.23_巾2005~3055追加
            .HaHiz1 = "2.5"
          ElseIf 2004# < HABA And HABA <= 2504# Then
            .HaHiz1 = "3.0"
          ElseIf 2504# < HABA And HABA <= 3054# Then
            .HaHiz1 = "3.5"
          End If
          If 0 < TAKE And TAKE <= 500# Then  '----------------------丈歪'2022.05.23_丈500~1530追加
            .TaHiz1 = "0.6"
          ElseIf 500# < TAKE And TAKE <= 750# Then
            .TaHiz1 = "0.9"
          ElseIf 750# < TAKE And TAKE <= 1004# Then
            .TaHiz1 = "1.2"
          ElseIf 1004# < TAKE And TAKE <= 1254# Then
            .TaHiz1 = "1.5"
          ElseIf 1254# < TAKE And TAKE <= 1529# Then
            .TaHiz1 = "2.0"
          ElseIf 1529# < TAKE And TAKE <= 2004# Then
            .TaHiz1 = "2.5"
          ElseIf 2004# < TAKE And TAKE <= 2504# Then
            .TaHiz1 = "3.0"
          ElseIf 2504# < TAKE And TAKE <= 3054# Then
            .TaHiz1 = "3.5"
          ElseIf 3054# < TAKE Then
            .TaHiz1 = "3.5"
          End If
        End Select
      ElseIf YoutC = "H197" Or YoutC = "H19S" Or _
      YoutC = "H19B" Or YoutC = "H19T" Or _
      YoutC = "H271" Or YoutC = "H248" Or _
      YoutC = "H249" Then '-------------------------------------------一般材　納入仕様書
        If Zai = "52S" Or Zai = "NP56" Or Zai = "F52S" Or Zai = "5083" Then '----------------------材質限定
         Select Case ATU
         Case 4# To 50#
            .MG1 = "1.0以下" '--------------------------------ｹﾞｰｼﾞ
            If ATU <= 6# Then .Naka1 = Format(中歪(ATU, HABA), "0.0") '--------------------------------中歪
            If ATU <= 6# Then .Mimi1 = Format(耳歪(ATU, HABA), "0.0") '--------------------------------耳歪
            If 0 < HABA And HABA <= 500# Then '-----------------------巾歪
            .HaHiz1 = "0.5"
            ElseIf 500# < HABA And HABA <= 750# Then
            .HaHiz1 = "0.8"
            ElseIf 750# < HABA And HABA <= 1004# Then
            .HaHiz1 = "1.0"
            ElseIf 1004# < HABA And HABA <= 1254# Then
            .HaHiz1 = "1.3"
            ElseIf 1254# < HABA And HABA <= 1529# Then
            .HaHiz1 = "1.8"
            ElseIf 1529# < HABA And HABA <= 2004# Then   '2022.05.23_巾2005~3055追加
            .HaHiz1 = "2.0"
            ElseIf 2004# < HABA And HABA <= 2504# Then
            .HaHiz1 = "2.0"
            ElseIf 2504# < HABA And HABA <= 3054# Then
            .HaHiz1 = "2.0"
            End If
            
            If 0 < TAKE And TAKE <= 500# Then  '----------------------丈歪'2022.05.23_丈500~1530追加
            .TaHiz1 = "0.5"
            ElseIf 500# < TAKE And TAKE <= 750# Then
            .TaHiz1 = "0.8"
            ElseIf 750# < TAKE And TAKE <= 1004# Then
            .TaHiz1 = "1.0"
            ElseIf 1004# < TAKE And TAKE <= 1254# Then
            .TaHiz1 = "1.3"
            ElseIf 1254# < TAKE And TAKE <= 1529# Then
            .TaHiz1 = "1.8"
            ElseIf 1529# < TAKE And TAKE <= 2004# Then
            .TaHiz1 = "2.0"
            ElseIf 2004# < TAKE And TAKE <= 2504# Then
            .TaHiz1 = "2.0"
            ElseIf 2504# < TAKE And TAKE <= 3054# Then
            .TaHiz1 = "2.0"
            ElseIf 3054# < TAKE Then
            .TaHiz1 = "2.0"
            End If
         Case 50.1 To 100#
            .MG1 = "1.2以下" '--------------------------------ｹﾞｰｼﾞ
            If 0 < HABA And HABA <= 500# Then '-----------------------巾歪
            .HaHiz1 = "0.6"
            ElseIf 500# < HABA And HABA <= 750# Then
            .HaHiz1 = "0.9"
            ElseIf 750# < HABA And HABA <= 1004# Then
            .HaHiz1 = "1.2"
            ElseIf 1004# < HABA And HABA <= 1254# Then
            .HaHiz1 = "1.5"
            ElseIf 1254# < HABA And HABA <= 1529# Then
            .HaHiz1 = "2.0"
            ElseIf 1529# < HABA And HABA <= 2004# Then   '2022.05.23_巾2005~3055追加
            .HaHiz1 = "2.5"
            ElseIf 2004# < HABA And HABA <= 2504# Then
            .HaHiz1 = "3.0"
            ElseIf 2504# < HABA And HABA <= 3054# Then
            .HaHiz1 = "3.5"
            End If
            
            If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
            .TaHiz1 = "0.6"
            ElseIf 500# < TAKE And TAKE <= 750# Then
            .TaHiz1 = "0.9"
            ElseIf 750# < TAKE And TAKE <= 1004# Then
            .TaHiz1 = "1.2"
            ElseIf 1004# < TAKE And TAKE <= 1254# Then
            .TaHiz1 = "1.5"
            ElseIf 1254# < TAKE And TAKE <= 1529# Then
            .TaHiz1 = "2.0"
            ElseIf 1529# < TAKE And TAKE <= 2004# Then
            .TaHiz1 = "2.5"
            ElseIf 2004# < TAKE And TAKE <= 2504# Then
            .TaHiz1 = "3.0"
            ElseIf 2504# < TAKE And TAKE <= 3054# Then
            .TaHiz1 = "3.5"
            ElseIf 3054# < TAKE Then
            .TaHiz1 = "3.5"
            End If
         Case 100.1 To 150#
             .MG1 = "1.4以下" '--------------------------------ｹﾞｰｼﾞ
             If 0 < HABA And HABA <= 500# Then '-----------------------巾歪
             .HaHiz1 = "0.7"
             ElseIf 500# < HABA And HABA <= 750# Then
             .HaHiz1 = "1.1"
             ElseIf 750# < HABA And HABA <= 1004# Then
             .HaHiz1 = "1.4"
             ElseIf 1004# < HABA And HABA <= 1254# Then
             .HaHiz1 = "1.8"
             ElseIf 1254# < HABA And HABA <= 1529# Then
             .HaHiz1 = "2.0"
             ElseIf 1529# < HABA And HABA <= 2004# Then   '2022.05.23_巾2005~3055追加
             .HaHiz1 = "2.5"
             ElseIf 2004# < HABA And HABA <= 2504# Then
             .HaHiz1 = "3.0"
             ElseIf 2504# < HABA And HABA <= 3054# Then
             .HaHiz1 = "3.5"
             End If
             
             If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
             .TaHiz1 = "0.7"
             ElseIf 500# < TAKE And TAKE <= 750# Then
             .TaHiz1 = "1.1"
             ElseIf 750# < TAKE And TAKE <= 1004# Then
             .TaHiz1 = "1.4"
             ElseIf 1004# < TAKE And TAKE <= 1254# Then
             .TaHiz1 = "1.8"
             ElseIf 1254# < TAKE And TAKE <= 1529# Then
             .TaHiz1 = "2.0"
             ElseIf 1529# < TAKE And TAKE <= 2004# Then
             .TaHiz1 = "2.5"
             ElseIf 2004# < TAKE And TAKE <= 2504# Then
             .TaHiz1 = "3.0"
             ElseIf 2504# < TAKE And TAKE <= 3054# Then
             .TaHiz1 = "3.5"
             ElseIf 3054# < TAKE Then
             .TaHiz1 = "3.5"
             End If
            Case 150.1 To 200#
             .MG1 = "2.0以下" '--------------------------------ｹﾞｰｼﾞ
             If 0 < HABA And HABA <= 500# Then '-----------------------巾歪
             .HaHiz1 = "1.0"
             ElseIf 500# < HABA And HABA <= 750# Then
             .HaHiz1 = "1.5"
             ElseIf 750# < HABA And HABA <= 1004# Then
             .HaHiz1 = "2.0"
             ElseIf 1004# < HABA And HABA <= 1254# Then
             .HaHiz1 = "2.0"
             ElseIf 1254# < HABA And HABA <= 1529# Then
             .HaHiz1 = "2.0"
             ElseIf 1529# < HABA And HABA <= 2004# Then
             .HaHiz1 = "3.0"
             ElseIf 2004# < HABA And HABA <= 2504# Then
             .HaHiz1 = "4.0"
             ElseIf 2504# < HABA And HABA <= 3054# Then
             .HaHiz1 = "4.0"
             End If
             If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪
             .TaHiz1 = "1.0"
             ElseIf 500# < TAKE And TAKE <= 750# Then
             .TaHiz1 = "1.5"
             ElseIf 750# < TAKE And TAKE <= 1004# Then
             .TaHiz1 = "2.0"
             ElseIf 1004# < TAKE And TAKE <= 1254# Then
             .TaHiz1 = "2.0"
             ElseIf 1254# < TAKE And TAKE <= 1529# Then
             .TaHiz1 = "2.0"
             ElseIf 1529# < TAKE And TAKE <= 2004# Then
             .TaHiz1 = "3.0"
             ElseIf 2004# < TAKE And TAKE <= 2504# Then
             .TaHiz1 = "4.0"
             ElseIf 2504# < TAKE And TAKE <= 3054# Then
             .TaHiz1 = "4.0"
             ElseIf 3054# < TAKE Then
             .TaHiz1 = "4.0"
             End If
         Case 200.1 To 500# '2022.05.30_追加修正
            .MG1 = "3.0以下" '--------------------------------ｹﾞｰｼﾞ
            If 0 < HABA And HABA <= 500# Then '-----------------------巾歪
            .HaHiz1 = "3.0"
            ElseIf 500# < HABA And HABA <= 750# Then
            .HaHiz1 = "3.0"
            ElseIf 750# < HABA And HABA <= 1004# Then
            .HaHiz1 = "3.0"
            ElseIf 1004# < HABA And HABA <= 1254# Then
            .HaHiz1 = "5.0"
            ElseIf 1254# < HABA And HABA <= 1529# Then
            .HaHiz1 = "5.0"
            ElseIf 1529# < HABA And HABA <= 2004# Then
            .HaHiz1 = "5.0"
            ElseIf 2004# < HABA And HABA <= 2504# Then
            .HaHiz1 = "5.0"
            ElseIf 2504# < HABA And HABA <= 3054# Then
            .HaHiz1 = "5.0"
            End If
            If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪
            .TaHiz1 = "3.0"
            ElseIf 500# < TAKE And TAKE <= 750# Then
            .TaHiz1 = "3.."
            ElseIf 750# < TAKE And TAKE <= 1004# Then
            .TaHiz1 = "3.0"
            ElseIf 1004# < TAKE And TAKE <= 1254# Then
            .TaHiz1 = "5.0"
            ElseIf 1254# < TAKE And TAKE <= 1529# Then
            .TaHiz1 = "5.0"
            ElseIf 1529# < TAKE And TAKE <= 2004# Then
            .TaHiz1 = "5.0"
            ElseIf 2004# < TAKE And TAKE <= 2504# Then
            .TaHiz1 = "5.0"
            ElseIf 2504# < TAKE And TAKE <= 3054# Then
            .TaHiz1 = "5.0"
            ElseIf 3054# < TAKE Then
            .TaHiz1 = "5.0"
            End If
         End Select
        End If
      '鉄道ﾎﾞﾃﾞｨ.Q-S-QA-129-01[別表]　2017.02.24作成版
      '--------------------------------------------------------------------------------用途鉄道ボディ限定_2022.02.09板厚120^200追加更新
      ElseIf YoutC = "K21S" Then
        If Zai = "5005" Then '----------------------------------------------------------5005
         Select Case ATU
         Case 0.1 To 5# '------------------------------------------5mm以下
          If 0 < HABA And HABA <= 500# Then '巾歪
          .HaHiz1 = "1.5"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "2.0"
          ElseIf 750# < HABA And HABA <= 1000# Then
          .HaHiz1 = "2.0"
          ElseIf 1000# < HABA And HABA <= 1250# Then
          .HaHiz1 = "2.0"
          ElseIf 1250# < HABA And HABA <= 1525# Then
          .HaHiz1 = "2.0"
          ElseIf 1525# < HABA And HABA <= 2000# Then '2022.05.23_巾2005~3055追加
          .HaHiz1 = "3.0"
          ElseIf 2000# < HABA And HABA <= 2500# Then
          .HaHiz1 = "3.0"
          ElseIf 2500# < HABA And HABA <= 3050# Then
          .HaHiz1 = "3.0"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "1.5"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "2.0"
          ElseIf 750# < TAKE And TAKE <= 1000# Then
          .TaHiz1 = "2.0"
          ElseIf 1000# < TAKE And TAKE <= 1250# Then
          .TaHiz1 = "2.0"
          ElseIf 1250# < TAKE And TAKE <= 1525# Then
          .TaHiz1 = "2.0"
          ElseIf 1525# < TAKE And TAKE <= 2000# Then
          .TaHiz1 = "3.0"
          ElseIf 2000# < TAKE And TAKE <= 2500# Then
          .TaHiz1 = "3.0"
          ElseIf 2500# < TAKE And TAKE <= 3050# Then
          .TaHiz1 = "3.0"
          ElseIf 3050# < TAKE Then
          .TaHiz1 = "3.0"
          End If
         End Select
        ElseIf Zai = "52S" Or Zai = "F52S" Then  '----------------------------------------------------------52S
         Select Case ATU
         Case 0.1 To 5# '------------------------------------------5mm以下
          If 0 < HABA And HABA <= 500# Then '巾歪
          .HaHiz1 = "1.5"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "2.0"
          ElseIf 750# < HABA And HABA <= 1000# Then
          .HaHiz1 = "2.0"
          ElseIf 1000# < HABA And HABA <= 1250# Then
          .HaHiz1 = "2.0"
          ElseIf 1250# < HABA And HABA <= 1525# Then
          .HaHiz1 = "2.0"
          ElseIf 1525# < HABA And HABA <= 2000# Then '2022.05.23_巾2005~3055追加
          .HaHiz1 = "3.0"
          ElseIf 2000# < HABA And HABA <= 2500# Then
          .HaHiz1 = "3.0"
          ElseIf 2500# < HABA And HABA <= 3050# Then
          .HaHiz1 = "3.0"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "1.5"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "2.0"
          ElseIf 750# < TAKE And TAKE <= 1000# Then
          .TaHiz1 = "2.0"
          ElseIf 1000# < TAKE And TAKE <= 1250# Then
          .TaHiz1 = "2.0"
          ElseIf 1250# < TAKE And TAKE <= 1525# Then
          .TaHiz1 = "2.0"
          ElseIf 1525# < TAKE And TAKE <= 2000# Then
          .TaHiz1 = "3.0"
          ElseIf 2000# < TAKE And TAKE <= 2500# Then
          .TaHiz1 = "3.0"
          ElseIf 2500# < TAKE And TAKE <= 3050# Then
          .TaHiz1 = "3.0"
          ElseIf 3050# < TAKE Then
          .TaHiz1 = "3.0"
          End If
         Case 5.01 To 75# '------------------------------------------5～75
          If 0 < HABA And HABA <= 500# Then '巾歪
          .HaHiz1 = "1.5"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "2.0"
          ElseIf 750# < HABA And HABA <= 1000# Then
          .HaHiz1 = "2.0"
          ElseIf 1000# < HABA And HABA <= 1250# Then
          .HaHiz1 = "2.0"
          ElseIf 1250# < HABA And HABA <= 1525# Then
          .HaHiz1 = "2.0"
          ElseIf 1525# < HABA And HABA <= 2000# Then '2022.05.23_巾2005~3055追加
          .HaHiz1 = "3.0"
          ElseIf 2000# < HABA And HABA <= 2500# Then
          .HaHiz1 = "3.5"
          ElseIf 2500# < HABA And HABA <= 3050# Then
          .HaHiz1 = "3.5"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "1.5"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "2.0"
          ElseIf 750# < TAKE And TAKE <= 1000# Then
          .TaHiz1 = "2.0"
          ElseIf 1000# < TAKE And TAKE <= 1250# Then
          .TaHiz1 = "2.0"
          ElseIf 1250# < TAKE And TAKE <= 1525# Then
          .TaHiz1 = "2.0"
          ElseIf 1525# < TAKE And TAKE <= 2000# Then
          .TaHiz1 = "3.0"
          ElseIf 2000# < TAKE And TAKE <= 2500# Then
          .TaHiz1 = "3.5"
          ElseIf 2500# < TAKE And TAKE <= 3050# Then
          .TaHiz1 = "3.5"
          ElseIf 3050# < TAKE Then
          .TaHiz1 = "3.5"
          End If
         Case 75.1 To 200# '------------------------------------------75～200
          If 0 < HABA And HABA <= 500# Then '巾歪
          .HaHiz1 = "2.0"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "2.0"
          ElseIf 750# < HABA And HABA <= 1000# Then
          .HaHiz1 = "2.5"
          ElseIf 1000# < HABA And HABA <= 1250# Then
          .HaHiz1 = "3.0"
          ElseIf 1250# < HABA And HABA <= 1525# Then
          .HaHiz1 = "3.0"
          ElseIf 1525# < HABA And HABA <= 2000# Then '2022.05.23_巾2005~3055追加
          .TaHiz1 = "3.5"
          ElseIf 2000# < HABA And HABA <= 2500# Then
          .TaHiz1 = "4.0"
          ElseIf 2500# < HABA And HABA <= 3050# Then
          .TaHiz1 = "4.0"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "2.0"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "2.0"
          ElseIf 750# < TAKE And TAKE <= 1000# Then
          .TaHiz1 = "2.5"
          ElseIf 1000# < TAKE And TAKE <= 1250# Then
          .TaHiz1 = "3.0"
          ElseIf 1250# < TAKE And TAKE <= 1525# Then
          .TaHiz1 = "3.0"
          ElseIf 1525# < TAKE And TAKE <= 2000# Then
          .TaHiz1 = "3.5"
          ElseIf 2000# < TAKE And TAKE <= 2500# Then
          .TaHiz1 = "4.0"
          ElseIf 2500# < TAKE And TAKE <= 3050# Then
          .TaHiz1 = "4.0"
          ElseIf 3050# < TAKE Then
          .TaHiz1 = "4.0"
          End If
         Case 200.1 To 500#  '------------------------------------------200～
          If 0 < HABA And HABA <= 500# Then '巾歪
          .HaHiz1 = "2.0"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "2.5"
          ElseIf 750# < HABA And HABA <= 1000# Then
          .HaHiz1 = "3.0"
          ElseIf 1000# < HABA And HABA <= 1250# Then
          .HaHiz1 = "3.0"
          ElseIf 1250# < HABA And HABA <= 1525# Then
          .HaHiz1 = "3.5"
          ElseIf 1525# < HABA And HABA <= 2000# Then '2022.05.23_巾2005~3055追加
          .HaHiz1 = "4.0"
          ElseIf 2000# < HABA And HABA <= 2500# Then
          .HaHiz1 = "4.0"
          ElseIf 2500# < HABA And HABA <= 3050# Then
          .HaHiz1 = "4.5"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "2.0"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "2.5"
          ElseIf 750# < TAKE And TAKE <= 1000# Then
          .TaHiz1 = "3.0"
          ElseIf 1000# < TAKE And TAKE <= 1250# Then
          .TaHiz1 = "3.0"
          ElseIf 1250# < TAKE And TAKE <= 1525# Then
          .TaHiz1 = "3.5"
          ElseIf 1525# < TAKE And TAKE <= 2000# Then
          .TaHiz1 = "4.0"
          ElseIf 2000# < TAKE And TAKE <= 2500# Then
          .TaHiz1 = "4.0"
          ElseIf 2500# < TAKE And TAKE <= 3050# Then
          .TaHiz1 = "4.5"
          ElseIf 3050# < TAKE Then
          .TaHiz1 = "4.5"
          End If
         End Select
        ElseIf Zai = "NP56" Or Zai = "5083" Then '----------------------------------------------------------NP56
         Select Case ATU
         Case 5.01 To 75# '------------------------------------------5～75
          If 0 < HABA And HABA <= 500# Then '巾歪
          .HaHiz1 = "1.5"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "2.0"
          ElseIf 750# < HABA And HABA <= 1000# Then
          .HaHiz1 = "2.0"
          ElseIf 1000# < HABA And HABA <= 1250# Then
          .HaHiz1 = "2.5"
          ElseIf 1250# < HABA And HABA <= 1525# Then
          .HaHiz1 = "2.5"
          ElseIf 1525# < HABA And HABA <= 2000# Then '2022.05.23_巾2005~3055追加
          .HaHiz1 = "2.5"
          ElseIf 2000# < HABA And HABA <= 2500# Then
          .HaHiz1 = "3.0"
          ElseIf 2500# < HABA And HABA <= 3050# Then
          .HaHiz1 = "3.5"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "1.5"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "2.0"
          ElseIf 750# < TAKE And TAKE <= 1000# Then
          .TaHiz1 = "2.0"
          ElseIf 1000# < TAKE And TAKE <= 1250# Then
          .TaHiz1 = "2.5"
          ElseIf 1250# < TAKE And TAKE <= 1525# Then
          .TaHiz1 = "2.5"
          ElseIf 1525# < TAKE And TAKE <= 2000# Then
          .TaHiz1 = "2.5"
          ElseIf 2000# < TAKE And TAKE <= 2500# Then
          .TaHiz1 = "3.0"
          ElseIf 2500# < TAKE And TAKE <= 3050# Then
          .TaHiz1 = "3.5"
          ElseIf 3050# < TAKE Then
          .TaHiz1 = "3.5"
          End If
         Case 75.1 To 120#  '------------------------------------------75～120
          If 0 < HABA And HABA <= 500# Then '巾歪
          .HaHiz1 = "2.0"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "2.0"
          ElseIf 750# < HABA And HABA <= 1000# Then
          .HaHiz1 = "2.5"
          ElseIf 1000# < HABA And HABA <= 1250# Then
          .HaHiz1 = "2.5"
          ElseIf 1250# < HABA And HABA <= 1525# Then
          .HaHiz1 = "2.5"
          ElseIf 1525# < HABA And HABA <= 2000# Then  '2022.05.23_巾2005~3055追加
          .HaHiz1 = "3.5"
          ElseIf 2000# < HABA And HABA <= 2500# Then
          .HaHiz1 = "3.5"
          ElseIf 2500# < HABA And HABA <= 3050# Then
          .HaHiz1 = "3.5"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "2.0"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "2.0"
          ElseIf 750# < TAKE And TAKE <= 1000# Then
          .TaHiz1 = "2.5"
          ElseIf 1000# < TAKE And TAKE <= 1250# Then
          .TaHiz1 = "2.5"
          ElseIf 1250# < TAKE And TAKE <= 1525# Then
          .TaHiz1 = "2.5"
          ElseIf 1525# < TAKE And TAKE <= 2000# Then
          .TaHiz1 = "3.5"
          ElseIf 2000# < TAKE And TAKE <= 2500# Then
          .TaHiz1 = "3.5"
          ElseIf 2500# < TAKE And TAKE <= 3050# Then
          .TaHiz1 = "3.5"
          ElseIf 3050# < TAKE Then
          .TaHiz1 = "3.5"
          End If
         Case 120.1 To 200#  '------------------------------------------120～200***2022.02追加部分
          If 0 < HABA And HABA <= 500# Then '巾歪
          .HaHiz1 = "2.0"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "2.5"
          ElseIf 750# < HABA And HABA <= 1000# Then
          .HaHiz1 = "2.5"
          ElseIf 1000# < HABA And HABA <= 1250# Then
          .HaHiz1 = "3.0"
          ElseIf 1250# < HABA And HABA <= 1525# Then
          .HaHiz1 = "3.0"
          ElseIf 1525# < HABA And HABA <= 2000# Then '2022.05.23_巾2005~3055追加
          .HaHiz1 = "3.5"
          ElseIf 2000# < HABA And HABA <= 2500# Then
          .HaHiz1 = "3.5"
          ElseIf 2500# < HABA And HABA <= 3050# Then
          .HaHiz1 = "4.0"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "2.0"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "2.5"
          ElseIf 750# < TAKE And TAKE <= 1000# Then
          .TaHiz1 = "2.5"
          ElseIf 1000# < TAKE And TAKE <= 1250# Then
          .TaHiz1 = "3.0"
          ElseIf 1250# < TAKE And TAKE <= 1525# Then
          .TaHiz1 = "3.0"
          ElseIf 1525# < TAKE And TAKE <= 2000# Then
          .TaHiz1 = "3.5"
          ElseIf 2000# < TAKE And TAKE <= 2500# Then
          .TaHiz1 = "3.5"
          ElseIf 2500# < TAKE And TAKE <= 3050# Then
          .TaHiz1 = "4.0"
          ElseIf 3050# < TAKE Then
          .TaHiz1 = "4.0"
          End If
         Case 200.1 To 500#   '------------------------------------------200～
          If 0 < HABA And HABA <= 500# Then '巾歪
          .HaHiz1 = "2.0"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "2.5"
          ElseIf 750# < HABA And HABA <= 1000# Then
          .HaHiz1 = "3.0"
          ElseIf 1000# < HABA And HABA <= 1250# Then
          .HaHiz1 = "3.5"
          ElseIf 1250# < HABA And HABA <= 1525# Then
          .HaHiz1 = "3.5"
          ElseIf 1525# < HABA And HABA <= 2000# Then '2022.05.23_巾2005~3055追加
          .HaHiz1 = "4.0"
          ElseIf 2000# < HABA And HABA <= 2500# Then
          .HaHiz1 = "4.0"
          ElseIf 2500# < HABA And HABA <= 3050# Then
          .HaHiz1 = "4.5"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "2.0"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "2.5"
          ElseIf 750# < TAKE And TAKE <= 1000# Then
          .TaHiz1 = "3.0"
          ElseIf 1000# < TAKE And TAKE <= 1250# Then
          .TaHiz1 = "3.5"
          ElseIf 1250# < TAKE And TAKE <= 1525# Then
          .TaHiz1 = "3.5"
          ElseIf 1525# < TAKE And TAKE <= 2000# Then
          .TaHiz1 = "4.0"
          ElseIf 2000# < TAKE And TAKE <= 2500# Then
          .TaHiz1 = "4.0"
          ElseIf 2500# < TAKE And TAKE <= 3050# Then
          .TaHiz1 = "4.5"
          ElseIf 3050# < TAKE Then
          .TaHiz1 = "4.5"
          End If
         End Select
        ElseIf Zai = "U74S" Then '----------------------------------------------------------U74S
         Select Case ATU
         Case 4# To 50#   '------------------------------------------4～50
          If 0 < HABA And HABA <= 500# Then '巾歪
          .HaHiz1 = "2.0"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "2.5"
          ElseIf 750# < HABA And HABA <= 1000# Then
          .HaHiz1 = "2.5"
          ElseIf 1000# < HABA And HABA <= 1250# Then
          .HaHiz1 = "2.5"
          ElseIf 1250# < HABA And HABA <= 1525# Then
          .HaHiz1 = "2.5"
          ElseIf 1525# < HABA And HABA <= 2000# Then '2022.05.23_巾2005~3055追加
          .HaHiz1 = "3.0"
          ElseIf 2000# < HABA And HABA <= 2500# Then
          .HaHiz1 = "3.5"
          ElseIf 2500# < HABA And HABA <= 3050# Then
          .HaHiz1 = "3.5"
          ElseIf 3050# < HABA Then
          .HaHiz1 = "3.5"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "2.0"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "2.5"
          ElseIf 750# < TAKE And TAKE <= 1000# Then
          .TaHiz1 = "2.5"
          ElseIf 1000# < TAKE And TAKE <= 1250# Then
          .TaHiz1 = "2.5"
          ElseIf 1250# < TAKE And TAKE <= 1525# Then
          .TaHiz1 = "2.5"
          ElseIf 1525# < TAKE And TAKE <= 2000# Then
          .TaHiz1 = "3.0"
          ElseIf 2000# < TAKE And TAKE <= 2500# Then
          .TaHiz1 = "3.5"
          ElseIf 2500# < TAKE And TAKE <= 3050# Then
          .TaHiz1 = "3.5"
          ElseIf 3050# < TAKE Then
          .TaHiz1 = "3.5"
          End If
         Case 50.1 To 75#  '------------------------------------------50～75
          If 0 < HABA And HABA <= 500# Then '巾歪
          .HaHiz1 = "2.0"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "2.5"
          ElseIf 750# < HABA And HABA <= 1000# Then
          .HaHiz1 = "2.5"
          ElseIf 1000# < HABA And HABA <= 1250# Then
          .HaHiz1 = "3.0"
          ElseIf 1250# < HABA And HABA <= 1525# Then
          .HaHiz1 = "3.0"
          ElseIf 1525# < HABA And HABA <= 2000# Then  '2022.05.23_巾2005~3055追加
          .HaHiz1 = "3.5"
          ElseIf 2000# < HABA And HABA <= 2500# Then
          .HaHiz1 = "4.0"
          ElseIf 2500# < HABA And HABA <= 3050# Then
          .HaHiz1 = "5.0"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then  '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "2.0"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "2.5"
          ElseIf 750# < TAKE And TAKE <= 1000# Then
          .TaHiz1 = "2.5"
          ElseIf 1000# < TAKE And TAKE <= 1250# Then
          .TaHiz1 = "3.0"
          ElseIf 1250# < TAKE And TAKE <= 1525# Then
          .TaHiz1 = "3.0"
          ElseIf 1525# < TAKE And TAKE <= 2000# Then
          .TaHiz1 = "3.5"
          ElseIf 2000# < TAKE And TAKE <= 2500# Then
          .TaHiz1 = "4.0"
          ElseIf 2500# < TAKE And TAKE <= 3050# Then
          .TaHiz1 = "5.0"
          ElseIf 3050# < TAKE Then
          .TaHiz1 = "5.0"
          End If
         Case 75.1 To 120 '------------------------------------------75～120
          If 0 < HABA And HABA <= 500# Then '巾歪
          .HaHiz1 = "2.0"
          ElseIf 500# < HABA And HABA <= 750# Then
          .HaHiz1 = "2.5"
          ElseIf 750# < HABA And HABA <= 1000# Then
          .HaHiz1 = "2.5"
          ElseIf 1000# < HABA And HABA <= 1250# Then
          .HaHiz1 = "3.0"
          ElseIf 1250# < HABA And HABA <= 1525# Then
          .HaHiz1 = "3.0"
          ElseIf 1525# < HABA And HABA <= 2000# Then '2022.05.23_巾2005~3055追加
          .HaHiz1 = "3.5"
          ElseIf 2000# < HABA And HABA <= 2500# Then
          .HaHiz1 = "4.0"
          ElseIf 2500# < HABA And HABA <= 3050# Then
          .HaHiz1 = "5.0"
          End If
          
          If 0 < TAKE And TAKE <= 500# Then '----------------------丈歪'2022.05.23_丈500~1530追加
          .TaHiz1 = "2.0"
          ElseIf 500# < TAKE And TAKE <= 750# Then
          .TaHiz1 = "2.5"
          ElseIf 750# < TAKE And TAKE <= 1000# Then
          .TaHiz1 = "2.5"
          ElseIf 1000# < TAKE And TAKE <= 1250# Then
          .TaHiz1 = "3.0"
          ElseIf 1250# < TAKE And TAKE <= 1525# Then
          .TaHiz1 = "3.0"
          ElseIf 1525# < TAKE And TAKE <= 2000# Then
          .TaHiz1 = "3.5"
          ElseIf 2000# < TAKE And TAKE <= 2500# Then
          .TaHiz1 = "4.0"
          ElseIf 2500# < TAKE And TAKE <= 3050# Then
          .TaHiz1 = "5.0"
          ElseIf 3050# < TAKE Then
          .TaHiz1 = "5.0"
          End If
         End Select
        End If
      End If
    End If
  End With
Application.ScreenUpdating = True
End Sub
Public Function 中歪(ATU As Double, HABA As Double, Optional grade As Long = 4) As Single '4級でセット厚板表示用_2022.05
If grade = 4 Then
  Select Case ATU
  Case 0 To 0.5
    If HABA <= 500# Then 中歪 = 1.5
    If HABA >= 501# Then 中歪 = 2#
  Case 0.51 To 2#
    If HABA <= 500# Then 中歪 = 1#
    If HABA >= 501# Then 中歪 = 2#
  Case 2.01 To 5#
    If HABA <= 500# Then 中歪 = 1#
    If HABA >= 501# Then 中歪 = 2#
  Case 5.01 To 12#
    If HABA <= 500# Then 中歪 = 1#
    If HABA >= 501# Then 中歪 = 2#
  Case 12.01 To 30#
    If HABA <= 500# Then 中歪 = 1.5
    If HABA >= 501# Then 中歪 = 1.5
  Case 30.01 To 50#
    If HABA <= 500# Then 中歪 = 1.5
    If HABA >= 501# Then 中歪 = 1.5
  End Select
ElseIf grade = 3 Then '緩い
  Select Case ATU
  Case 0 To 0.5
    If HABA <= 500# Then 中歪 = 2#
    If HABA >= 501# Then 中歪 = 4#
  Case 0.51 To 2#
    If HABA <= 500# Then 中歪 = 2#
    If HABA >= 501# Then 中歪 = 4#
  Case 2.01 To 5#
    If HABA <= 500# Then 中歪 = 3#
    If HABA >= 501# Then 中歪 = 4#
  Case 5.01 To 12#
    If HABA <= 500# Then 中歪 = 3#
    If HABA >= 501# Then 中歪 = 4#
  Case 12.01 To 30#
    If HABA <= 500# Then 中歪 = 3#
    If HABA >= 501# Then 中歪 = 3#
  Case 30.01 To 50#
    If HABA <= 500# Then 中歪 = 3#
    If HABA >= 501# Then 中歪 = 3#
  End Select
End If
End Function
Public Function 耳歪(ATU As Double, HABA As Double, Optional grade As Long = 4) As Single '4級でセット厚板表示用_2022.05
If grade = 4 Then
  Select Case ATU
  Case 0 To 0.5
    If HABA <= 500# Then 耳歪 = Format(1.5, "0.0")
    If HABA >= 501# Then 耳歪 = 2#
  Case 0.51 To 2#
    If HABA <= 500# Then 耳歪 = 1#
    If HABA >= 501# Then 耳歪 = 2#
  Case 2.01 To 5#
    If HABA <= 500# Then 耳歪 = 1.5
    If HABA >= 501# Then 耳歪 = 2#
  Case 5.01 To 12#
    If HABA <= 500# Then 耳歪 = 1.5
    If HABA >= 501# Then 耳歪 = 2#
  Case 12.01 To 30#
    If HABA <= 500# Then 耳歪 = 1.5
    If HABA >= 501# Then 耳歪 = 1.5
  Case 30.01 To 50#
    If HABA <= 500# Then 耳歪 = 1.5
    If HABA >= 501# Then 耳歪 = 1.5
  End Select
ElseIf grade = 3 Then '緩い
  Select Case ATU
  Case 0 To 0.5
    If HABA <= 500# Then 耳歪 = 2#
    If HABA >= 501# Then 耳歪 = 4#
  Case 0.51 To 2#
    If HABA <= 500# Then 耳歪 = 2#
    If HABA >= 501# Then 耳歪 = 4#
  Case 2.01 To 5#
    If HABA <= 500# Then 耳歪 = 3#
    If HABA >= 501# Then 耳歪 = 4#
  Case 5.01 To 12#
    If HABA <= 500# Then 耳歪 = 3#
    If HABA >= 501# Then 耳歪 = 4#
  Case 12.01 To 30#
    If HABA <= 500# Then 耳歪 = 3#
    If HABA >= 501# Then 耳歪 = 3#
  Case 30.01 To 50#
    If HABA <= 500# Then 耳歪 = 3#
    If HABA >= 501# Then 耳歪 = 3#
  End Select
End If
End Function
Public Sub 表面品質基準(LotNo As String) '薄板_履歴からの表示は不可とする*作業中のみ(DB参照)_2021.012.05*手動検索追加
'-------------------------------------------------品質基準ﾊﾟﾀｰﾝ　早見表　Q-S-QA-128-R00[別表2]　2015.08.19_最終更新*****2022.02.09_更新
Dim TableName As String: TableName = "仕掛"
Dim FPath As String

Select Case UF品質基準.LineName
Case "Test":  FPath = MainPath 'ThisWorkbook.Path & "\"
Case Else:    FPath = MainPath
End Select

Dim target As String:    target = "SIKALOTNOW.accdb"
Dim FilePath As String:  FilePath = FPath & target 'Lot

If FPath = MainPath Then
  If ReadCheck(FilePath, 1) = False Or TimeDesignation Or FileDesignation = False Then   '1：ファイル,2：フォルダ
    FPath = NextPath
    FilePath = FPath & target
    If ReadCheck(NextPath & target, 1) = False Then '梱包課ﾌｧｲﾙを使う*1日更新
      MsgBox FilePath & "SIKALOTNOW_ﾌｧｲﾙがありません", vbCritical
      Exit Sub
    End If
  End If
End If

'If TimeDesignation Or FileDesignation = False Then
'  MsgBox "元ファイル形成中、1分後リトライしてください。"
'  Exit Sub
'End If

Dim SQLfilter As String: SQLfilter = "[ﾛｯﾄ番号]='" & LotNo & "'" 'ﾛｯﾄNo検索
Dim TempLot As Variant: TempLot = GetRecordsArr(FilePath, TableName, SQLfilter)
If IsEmpty(TempLot) Then  'ﾌｧｲﾙが読めても配列が不完全だった場合
  Select Case UF品質基準.LineName
  Case "Test":    FPath = NextPath 'ThisWorkbook.Path & "\"
  Case Else:      FPath = NextPath
  End Select
  target = "SIKATWOYEAR.accdb" '専用
  FilePath = FPath & target
  TempLot = GetRecordsArr(FilePath, TableName, SQLfilter)
End If

If IsEmpty(TempLot) Then Exit Sub  'それでもEmptyの時は停止

'ﾌｨｰﾙﾄﾞ名ﾁｪｯｸ　*RNE更新に対応
Dim fieldsfilter As String: fieldsfilter = ""
Dim Lotfields As Variant:  Lotfields = GetFieldsArr(FilePath, TableName, fieldsfilter)
If IsEmpty(Lotfields) Then Exit Sub 'Emptyの時は停止

Dim i As Long, R_表面処理 As Long, R_厚 As Long, R_巾 As Long, R_丈 As Long, R_材 As Long, R_調 As Long, R_用途 As Long
For i = LBound(Lotfields, 2) To UBound(Lotfields, 2)
  Select Case Lotfields(1, i)
  Case "品質ｸﾞﾚｰﾄﾞ_表面処理": R_表面処理 = i
  Case "製造板厚":            R_厚 = i
  Case "製造板幅":            R_巾 = i
  Case "製造板丈":            R_丈 = i
  Case "製造材質":            R_材 = i
  Case "製造調質":            R_調 = i
  Case "用途ｺｰﾄﾞ":            R_用途 = i
  End Select
Next i
'ﾌｫｰﾑ転記
With UF品質基準
  .Uatu = Format(TempLot(LBound(TempLot, 1), R_厚), "0.000")
  .Uhaba = Format(TempLot(LBound(TempLot, 1), R_巾), "0.0")
  .Utake = Format(TempLot(LBound(TempLot, 1), R_丈), "0.0")
  .Uzai = TempLot(LBound(TempLot, 1), R_材)
  .Utyou = TempLot(LBound(TempLot, 1), R_調)
  .UYoucode = TempLot(LBound(TempLot, 1), R_用途)
  .Uhyou = TempLot(LBound(TempLot, 1), R_表面処理)
  .Upater = PatternJudge(.Uhyou, .Uzai, .UYoucode, .Uatu)
  If .Upater = 0 Then
    MsgBox "該当ﾊﾟﾀｰﾝなし"
    Exit Sub
  End If
'早見表読み込み
  Dim トビコミ:     トビコミ = GradeJudge("トビコミ", .Upater) '0配列
  Dim キズ:         キズ = GradeJudge("キズ", .Upater)
  Dim 突:           突 = GradeJudge("突", .Upater)
  Dim トモズレ:     トモズレ = GradeJudge("トモズレ", .Upater)
  Dim 汚れ:         汚れ = GradeJudge("汚れ", .Upater)
  Dim 油より:       油より = GradeJudge("油より", .Upater)
  Dim ロールキズ:   ロールキズ = GradeJudge("ロールキズ", .Upater)
  Dim ロールマーク: ロールマーク = GradeJudge("ロールマーク", .Upater)
  Dim カミコミ:     カミコミ = GradeJudge("カミコミ", .Upater)
  Dim 切断面:       切断面 = GradeJudge("切断面", .Upater)
  Dim 変色:         変色 = GradeJudge("変色", .Upater)
End With
'反映
For i = 1 To 4
  UF品質基準("TB" & i) = トビコミ(i - 1)
  UF品質基準("K" & i) = キズ(i - 1)
  UF品質基準("TA" & i) = 突(i - 1)
  UF品質基準("TM" & i) = トモズレ(i - 1)
  UF品質基準("Y" & i) = 汚れ(i - 1)
  UF品質基準("A" & i) = 油より(i - 1)
  UF品質基準("RK" & i) = ロールキズ(i - 1)
  UF品質基準("RM" & i) = ロールマーク(i - 1)
  UF品質基準("KM" & i) = カミコミ(i - 1)
  UF品質基準("S" & i) = 切断面(i - 1)
  UF品質基準("E" & i) = 変色(i - 1)
Next i
End Sub
Public Function PatternJudge(Hyou As String, Zai As String, Yout As String, ATU As Single) As Long 'パターン早見表

'2023.1.20_材質検索をあいまい検索に修正(材質をもろもろいれていた)

If Hyou = "" Or Zai = "" Or Yout = "" Or IsNumeric(ATU) = False Then Exit Function
'-------------------------------------------------品質基準ﾊﾟﾀｰﾝ　早見表　Q-S-QA-SRF-128-R00[別表1]　2015.08.19_最終更新*****2022.02.09_更新_2022.12.21_更新*R02
Select Case Hyou
Case "B", "H", "K" '-------------------------------------------------------表面
  Select Case True
  Case Zai Like "*52S"
    Select Case Yout
    Case "H115", "H127", "U19B", "P191", "H19B", "K21B", "H196", "U194", "H228", "H299", "H248", "H249" 'H299, H248, H249追加_2022.12.21
      Select Case ATU
      Case 0 To 2.5:                                           PatternJudge = 1
      Case 2.51 To 7.999:                                      PatternJudge = 2
      End Select
    End Select
  Case Zai = "NK01"
    Select Case Yout
    Case "T238", "T23B", "T23H", "T23K":                       PatternJudge = 4
    End Select
  Case Zai Like "*995"
    Select Case Yout
    Case "U19B", "H19B", "T271":                               PatternJudge = 4
    Case "F191", "F61B":                                       PatternJudge = 6
    End Select
  Case (Zai Like "*17S") Or (Zai Like "*61S") Or (Zai Like "*75S") Or (Zai Like "*24S")
    Select Case Yout
    Case "H115", "R191", "U19B", "P197", "H19B", "P181":       PatternJudge = 2
    End Select
  End Select
Case "S" '-----------------------------------------------------------------表面
  Select Case True
  Case Zai Like "*52S"
    Select Case Yout
    Case "K32S", "K19S", "H197", "H19S", "K194", "C21S", "H193", "H211", "K21S", "J19S", "K197", "T322", "T323", "T324", "T321" 'K197追加_2022.02.09**2023.1.20追加_T322,T323,T324,T321
      Select Case ATU
      Case 0 To 2.5:                                           PatternJudge = 1
      Case 2.51 To 7.999:                                      PatternJudge = 2
      End Select
    Case "T323", "T32S", "T321"
                                                               PatternJudge = 5
    End Select
  Case Zai = "NP56"
    Select Case Yout
    Case "H197", "H19S", "C218", "C217", "K21S":               PatternJudge = 2 'ﾀﾝｸｶｶﾞﾐｲﾀ→ﾀﾝｸｶｶﾞﾐｽｷﾝP_2022.02.09
    Case "C215", "C216", "C219":                               PatternJudge = 5 'C219追加_2022.02.09
    End Select
  Case Zai Like "*995"
    Select Case Yout
    Case "R292", "H19S":                                       PatternJudge = 4
    Case "N192", "F61S":                                       PatternJudge = 6
    End Select
  Case (Zai Like "*17S") Or (Zai Like "*61S") Or (Zai Like "*75S") Or (Zai Like "*24S")
    Select Case Yout
    Case "H197", "H19S":                                       PatternJudge = 2
    End Select
  Case Zai = "5154"
    Select Case Yout
    Case "N193":                                               PatternJudge = 6 '2023.1.20追加
    End Select
  End Select
Case "T", "C" '------------------------------------------------------------表面
  Select Case True
  Case Zai Like "*52S"
    Select Case Yout
    Case "U19T", "F97T", "H19T", "K21T":                       PatternJudge = 3
    End Select
  Case Zai = "NK01"
    Select Case Yout
    Case "T23T":                                               PatternJudge = 3
    End Select
  Case Zai Like "*995"
    Select Case Yout
    Case "U19T", "H31C":                                       PatternJudge = 3
    End Select
  End Select
End Select
End Function
Public Function GradeJudge(Abno As String, Pattern As Long) As Variant '早見表 判定結果***グレード表
'-------------------------------------------------品質基準ﾊﾟﾀｰﾝ　早見表　Q-S-QA-SRF-128-R00[別表1]　2015.08.19_最終更新*****2022.02.09_更新*R01
Dim O1 As String, O2 As String, O3 As String, O4 As String, O5 As String, O6 As String
Dim S1 As String, S2 As String, S3 As String, S4 As String, S5 As String, S6 As String
Dim K1 As String, K2 As String, K3 As String, K4 As String, K5 As String, K6 As String
Dim A1 As String, A2 As String, A3 As String, A4 As String, A5 As String, A6 As String
'表に統一　不要　許容を／に変更
Select Case Abno
Case "トビコミ" '---------------------------------------------------------------------------
  O1 = "2mm以下"
  O2 = "2mm以下"
  O3 = "2mm以下"
  O4 = "2mm以下"
  O5 = "2mm以下"
  O6 = "無き事"
  
  S1 = "爪にかからない"
  S2 = "爪に軽くかかる"
  S3 = "爪に軽くかかる"
  S4 = "爪にかからない"
  S5 = "爪にかからない"
  S6 = "無き事"
  
  K1 = "ﾋﾟｯﾁ300mm以上"
  K2 = "ﾋﾟｯﾁ100mm以上"
  K3 = "ﾋﾟｯﾁ100mm以上"
  K4 = "ﾋﾟｯﾁ300mm以上"
  K5 = "ﾋﾟｯﾁ300mm以上"
  K6 = "無き事"
  
  A1 = "／"
  A2 = "／"
  A3 = "／"
  A4 = "／"
  A5 = "／"
  A6 = "／"
Case "キズ" '---------------------------------------------------------------------------
  O1 = "長さ30～50mm"
  O2 = "長さ30～50mm"
  O3 = "長さ30～50mm"
  O4 = "長さ30～50mm"
  O5 = "長さ30～50mm"
  O6 = "無き事"
  
  S1 = "爪にかからない"
  S2 = "爪に軽くかかる"
  S3 = "爪に軽くかかる"
  S4 = "爪にかからない"
  S5 = "爪にかからない"
  S6 = "無き事"
  
  K1 = "3個まで"
  K2 = "3個まで"
  K3 = "3個まで"
  K4 = "3個まで"
  K5 = "3個まで"
  K6 = "無き事"
  
  A1 = "／"
  A2 = "／"
  A3 = "／"
  A4 = "／"
  A5 = "／"
  A6 = "／"
Case "突" '---------------------------------------------------------------------------
  O1 = "／"
  O2 = "／"
  O3 = "／"
  O4 = "／"
  O5 = "／"
  O6 = "無き事"
  
  S1 = "爪にかからない"
  S2 = "爪に軽くかかる"
  S3 = "爪に軽くかかる"
  S4 = "爪にかからない"
  S5 = "爪にかからない"
  S6 = "無き事"
  
  K1 = "／"
  K2 = "／"
  K3 = "／"
  K4 = "／"
  K5 = "／"
  K6 = "無き事"
  
  A1 = "／"
  A2 = "／"
  A3 = "／"
  A4 = "／"
  A5 = "／"
  A6 = "／"
Case "トモズレ" '---------------------------------------------------------------------------
  O1 = "／"
  O2 = "／"
  O3 = "／"
  O4 = "／"
  O5 = "／"
  O6 = "無き事"
  
  S1 = "爪にかからない"
  S2 = "爪にかからない" '爪に軽くかかる　→　爪にかからないへ変更
  S3 = "／"
  S4 = "爪にかからない"
  S5 = "爪にかからない"
  S6 = "無き事"
  
  K1 = "／"
  K2 = "／"
  K3 = "／"
  K4 = "／"
  K5 = "／"
  K6 = "無き事"
  
  A1 = "／"
  A2 = "／"
  A3 = "／"
  A4 = "／"
  A5 = "／"
  A6 = "／"
Case "汚れ" '---------------------------------------------------------------------------
  O1 = "長さ10mm以下"
  O2 = "長さ30mm以下"
  O3 = "／"
  O4 = "長さ10mm以下"
  O5 = "長さ10mm以下"
  O6 = "無き事"
  
  S1 = "／"
  S2 = "／"
  S3 = "／"
  S4 = "／"
  S5 = "／"
  S6 = "／"
  
  K1 = "3個まで"
  K2 = "3個まで"
  K3 = "／"
  K4 = "1個まで"
  K5 = "3個まで"
  K6 = "無き事"
  
  A1 = "／"
  A2 = "／"
  A3 = "／"
  A4 = "／"
  A5 = "／"
  A6 = "／"
Case "油より" '---------------------------------------------------------------------------
  O1 = "目立たない事"
  O2 = "／"
  O3 = "／"
  O4 = "目立たない事"
  O5 = "／"
  O6 = "／"
  
  S1 = "／"
  S2 = "／"
  S3 = "／"
  S4 = "／"
  S5 = "／"
  S6 = "／"
  
  K1 = "／"
  K2 = "1本まで"
  K3 = "／"
  K4 = "／"
  K5 = "無き事"
  K6 = "無き事"
  
  A1 = "／"
  A2 = "1分で消える事"
  A3 = "／"
  A4 = "／"
  A5 = "／"
  A6 = "／"
Case "ロールキズ" '---------------------------------------------------------------------------
  O1 = "3mm以下"
  O2 = "5mm以下"
  O3 = "／"
  O4 = "3mm以下"
  O5 = "3mm以下"
  O6 = "無き事"
  
  S1 = "／"
  S2 = "／"
  S3 = "／"
  S4 = "／"
  S5 = "／"
  S6 = "／"
  
  K1 = "ﾋﾟｯﾁ500mm以上"
  K2 = "ﾋﾟｯﾁ500mm以上"
  K3 = "／"
  K4 = "ﾋﾟｯﾁ500mm以上"
  K5 = "ﾋﾟｯﾁ1000mm以上"
  K6 = "無き事"
  
  A1 = "1分で消える事"
  A2 = "／"
  A3 = "／"
  A4 = "1分で消える事"
  A5 = "1分で消える事"
  A6 = "／"
Case "ロールマーク" '---------------------------------------------------------------------------
  O1 = "20mm以下"
  O2 = "50mm以下"
  O3 = "／"
  O4 = "10mm以下"
  O5 = "10mm以下"
  O6 = "無き事"
  
  S1 = "／"
  S2 = "／"
  S3 = "／"
  S4 = "／"
  S5 = "／"
  S6 = "／"
  
  K1 = "1M当たり1個まで"
  K2 = "1M当たり2個まで"
  K3 = "／"
  K4 = "1M当たり1個まで"
  K5 = "1M当たり1個まで"
  K6 = "無き事"
  
  A1 = "／"
  A2 = "／"
  A3 = "／"
  A4 = "／"
  A5 = "／"
  A6 = "／"
Case "カミコミ" '---------------------------------------------------------------------------
  O1 = "無き事"
  O2 = "全巾に無き事"
  O3 = "／"
  O4 = "無き事"
  O5 = "無き事"
  O6 = "無き事"
  
  S1 = "／"
  S2 = "／"
  S3 = "／"
  S4 = "／"
  S5 = "／"
  S6 = "／"
  
  K1 = "／"
  K2 = "／"
  K3 = "／"
  K4 = "／"
  K5 = "／"
  K6 = "／"
  
  A1 = "／"
  A2 = "／"
  A3 = "／"
  A4 = "／"
  A5 = "／"
  A6 = "／"
Case "切断面" '---------------------------------------------------------------------------
  O1 = "／"
  O2 = "／"
  O3 = "／"
  O4 = "／"
  O5 = "／"
  O6 = "無き事"
  
  S1 = "ビビり無き事"
  S2 = "段差無き事"
  S3 = "段差無き事"
  S4 = "ビビり無き事"
  S5 = "ビビり無き事"
  S6 = "無き事"
  
  K1 = "／"
  K2 = "／"
  K3 = "／"
  K4 = "／"
  K5 = "／"
  K6 = "／"
  
  A1 = "／"
  A2 = "／"
  A3 = "／"
  A4 = "／"
  A5 = "／"
  A6 = "／"
Case "変色" '---------------------------------------------------------------------------
  O1 = "境界目立たない事"
  O2 = "全丈に無き事"
  O3 = "／"
  O4 = "境界目立たない事"
  O5 = "／"
  O6 = "／"
  
  S1 = "／"
  S2 = "／"
  S3 = "／"
  S4 = "／"
  S5 = "／"
  S6 = "／"
  
  K1 = "／"
  K2 = "／"
  K3 = "／"
  K4 = "／"
  K5 = "／"
  K6 = "／"
  
  A1 = "／"
  A2 = "／"
  A3 = "／"
  A4 = "1分で消える事"
  A5 = "／"
  A6 = "／"
End Select
Select Case Pattern
Case 1: GradeJudge = Array(O1, S1, K1, A1)
Case 2: GradeJudge = Array(O2, S2, K2, A2)
Case 3: GradeJudge = Array(O3, S3, K3, A3)
Case 4: GradeJudge = Array(O4, S4, K4, A4)
Case 5: GradeJudge = Array(O5, S5, K5, A5)
Case 6: GradeJudge = Array(O6, S6, K6, A6)
End Select
End Function
Public Sub BoColor(Bot As MSForms.CommandButton)
Dim i As Long
For i = 1 To 11
  UF品質基準("B" & i).BackColor = &HFFFFFF
Next i

UF品質基準.SelectO = ""
UF品質基準.SelectS = ""
UF品質基準.SelectK = ""
UF品質基準.SelectA = ""

Bot.BackColor = &H80FFFF
End Sub
Public Sub SelectAdd(Bot As MSForms.CommandButton)
If UF品質基準.Upater = "" Then Exit Sub
Dim 異常
With UF品質基準
  Select Case Bot.Name '早見表読み込み
  Case "B1":  異常 = GradeJudge("トビコミ", .Upater) '0配列
  Case "B2":  異常 = GradeJudge("キズ", .Upater)
  Case "B3":  異常 = GradeJudge("トモズレ", .Upater)
  Case "B4":  異常 = GradeJudge("突", .Upater)
  Case "B5":  異常 = GradeJudge("汚れ", .Upater)
  Case "B6":  異常 = GradeJudge("油より", .Upater)
  Case "B7":  異常 = GradeJudge("ロールキズ", .Upater)
  Case "B8":  異常 = GradeJudge("ロールマーク", .Upater)
  Case "B9":  異常 = GradeJudge("カミコミ", .Upater)
  Case "B10": 異常 = GradeJudge("切断面", .Upater)
  Case "B11": 異常 = GradeJudge("変色", .Upater)
  End Select
End With
If IsEmpty(異常) Then Exit Sub '表示前に停止
UF品質基準.SelectO = 異常(0)
UF品質基準.SelectS = 異常(1)
UF品質基準.SelectK = 異常(2)
UF品質基準.SelectA = 異常(3)
End Sub

