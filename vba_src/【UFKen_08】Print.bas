Attribute VB_Name = "【UFKen_08】Print"
Option Explicit
Public PrintedLots As Object ' Dictionary
Public Const FLAG_SHEET_NAME As String = "印刷履歴管理" ' 隠しシート名

' ========== 初期化処理 ==========
' UserForm_Initializeで呼び出す
Public Sub InitializePrintFlags()
If PrintedLots Is Nothing Then
    Set PrintedLots = CreateObject("Scripting.Dictionary")
End If

' 隠しシートから印刷履歴を読み込む
Call LoadPrintHistory
End Sub

' ========== 隠しシートの作成 ==========
Private Sub CreateFlagSheet()
If WsCheck(FLAG_SHEET_NAME) = False Then
    Call AddWS(FLAG_SHEET_NAME)
    With ThisWorkbook.Sheets(FLAG_SHEET_NAME)
        .Visible = xlSheetVeryHidden ' 非表示
        .Cells(1, 1).Value = "LOT番号"
        .Cells(1, 2).Value = "印刷日時"
        .Cells(1, 3).Value = "ライン名"
    End With
End If
End Sub

' ========== 印刷履歴の読み込み ==========
Public Sub LoadPrintHistory()
Call CreateFlagSheet ' シートがなければ作成

PrintedLots.RemoveAll ' 一旦クリア

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets(FLAG_SHEET_NAME)

Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

If lastRow > 1 Then
    Dim i As Long
    For i = 2 To lastRow
        Dim LotNo As String
        LotNo = ws.Cells(i, 1).Value
        If LotNo <> "" Then
            ' LOT番号をキーに、印刷日時を値として格納
            PrintedLots(LotNo) = ws.Cells(i, 2).Value
        End If
    Next i
End If
End Sub

' ========== 印刷履歴の保存 ==========
Public Sub SavePrintHistory(LotNo As String, Lname As String)
Debug.Print "SavePrintHistory実行: " & LotNo & " / " & Lname
Call CreateFlagSheet

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets(FLAG_SHEET_NAME)
Debug.Print "シート取得成功"

' 既存のLOT番号があるか確認
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

Dim existingRow As Long
existingRow = 0

If lastRow > 1 Then
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = LotNo Then
            existingRow = i
            Exit For
        End If
    Next i
End If

' 新規または更新
If existingRow = 0 Then
    ' 新規追加
    Dim newRow As Long
    newRow = lastRow + 1
    ws.Cells(newRow, 1).Value = LotNo
    ws.Cells(newRow, 2).Value = Now
    ws.Cells(newRow, 3).Value = Lname
Else
    ' 既存を更新（再印刷の日時を更新）
    ws.Cells(existingRow, 2).Value = Now
End If
End Sub

' ========== 印刷済みかチェック ==========
Public Function IsPrinted(LotNo As String) As Boolean
Call InitializePrintFlags
IsPrinted = PrintedLots.exists(LotNo)
End Function

' ========== 印刷ボタンの色を更新 ==========
Public Sub UpdatePrintButtonColor(LotNo As String)
If IsPrinted(LotNo) Then
    ' 印刷済み = オレンジ色
    UF1.PrintB.BackColor = RGB(255, 165, 0) ' オレンジ
    UF1.PrintB.Caption = "再印刷"
    UF1.PrintBreset.Visible = True
Else
    ' 未印刷 = デフォルト色（または緑色）
    UF1.PrintB.BackColor = RGB(192, 192, 192) ' グレー（デフォルト）
    ' または RGB(144, 238, 144) ' ライトグリーン
    UF1.PrintB.Caption = "印刷"
    UF1.PrintBreset.Visible = False
End If
End Sub

' ========== LOT番号が確定したタイミングでボタン色をチェック ==========
' 例：LOT番号を読み込んだ後や、フォームにデータをロードした後に呼び出す
Public Sub CheckAndUpdateButtonColor()
If UF1.Lot.Text <> "" Then
    Call UpdatePrintButtonColor(UF1.Lot.Text)
End If
End Sub

' ========== 印刷済みフラグを設定 ==========
Public Sub SetPrinted(LotNo As String, Lname As String)
Call InitializePrintFlags
PrintedLots(LotNo) = Now ' メモリ上のDictionaryに追加
Call SavePrintHistory(LotNo, Lname) ' 隠しシートに保存
End Sub

' ========== 印刷済みフラグをクリア（個別） ==========
Public Sub ClearPrinted(LotNo As String)
Call InitializePrintFlags

' メモリから削除
If PrintedLots.exists(LotNo) Then
    PrintedLots.Remove LotNo
End If

' 隠しシートから削除
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets(FLAG_SHEET_NAME)

Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

If lastRow > 1 Then
    Dim i As Long
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 1).Value = LotNo Then
            ws.Rows(i).Delete
            Exit For
        End If
    Next i
End If
End Sub

' ========== 古い履歴の自動削除（3日以上前） ==========残しておく意味がない
Public Sub CleanOldHistory()
Call CreateFlagSheet

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets(FLAG_SHEET_NAME)

Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

If lastRow > 1 Then
    Dim i As Long
    For i = lastRow To 2 Step -1
        Dim printDate As Date
        If IsDate(ws.Cells(i, 2).Value) Then
            printDate = ws.Cells(i, 2).Value
            ' 3日以上前のデータを削除
            If DateDiff("d", printDate, Date) > 3 Then
                ws.Rows(i).Delete
            End If
        End If
    Next i
End If
End Sub

Public Sub 印刷転送() '印刷原紙に合わせて出力
'2024.2.15_BH停止
'＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞罫線を修正するときはCopy処理に注意
Dim i As Long, a As Long, n As Long
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Application.ScreenUpdating = False
Call EndTime_Register '終了時間登録
Call AddWS("印刷用" & "_" & Lname) '追加シート処理
Call Pesute("印刷用原紙", "印刷用" & "_" & Lname) '(原紙名、付ける名前)
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("印刷用" & "_" & Lname)

With UF1
  Dim LotDate As Variant:   LotDate = Array(.Lot.Text, .ATU.Text, .Hab.Text, .Tak.Text, .Am.Text, .Ap.Text, .Hm.Text, .Hp.Text, .Tm.Text, .Tp.Text, .AmK.Text, .ApK.Text, .HmK.Text, .HpK.Text, .TmK.Text, .TpK.Text) 'ロット情報格納
  Dim HizKousa As Variant:  HizKousa = Array(.MG.Text, .HaHiz.Text, .TaHiz.Text, .Naka.Text, .Mimi.Text, .DistortionFrame.Caption, .HURIS.Text) '2024.2.8_追加**2026.03.05フリー反り公差追加
  '外観に記号を追加　〇△
  Dim a_S As String: a_S = .S_UpVisual.Caption
  Dim a_U As String: a_U = .U_UpVisual.Caption
  Dim a_C As String: a_C = .C_UpVisual.Caption
  Dim a_D As String: a_D = .D_UpVisual.Caption
  Dim a_E As String: a_E = .E_UpVisual.Caption
  Dim a_F As String: a_F = .F_UpVisual.Caption
  If (a_S <> "") And (Not a_S Like "*欠陥ﾅｼ*") Then a_S = "△" & a_S
  If (a_U <> "") And (Not a_U Like "*欠陥ﾅｼ*") Then a_U = "△" & a_U
  If (a_C <> "") And (Not a_C Like "*欠陥ﾅｼ*") Then a_C = "△" & a_C
  If (a_D <> "") And (Not a_D Like "*欠陥ﾅｼ*") Then a_D = "△" & a_D
  If (a_E <> "") And (Not a_E Like "*欠陥ﾅｼ*") Then a_E = "△" & a_E
  If (a_F <> "") And (Not a_F Like "*欠陥ﾅｼ*") Then a_F = "△" & a_F
  
  If a_S = "欠陥ﾅｼ" Then a_S = "〇" & a_S
  If a_U = "欠陥ﾅｼ" Then a_U = "〇" & a_U
  If a_C = "欠陥ﾅｼ" Then a_C = "〇" & a_C
  If a_D = "欠陥ﾅｼ" Then a_D = "〇" & a_D
  If a_E = "欠陥ﾅｼ" Then a_E = "〇" & a_E
  If a_F = "欠陥ﾅｼ" Then a_F = "〇" & a_F
  
  Dim b_S As String: b_S = .S_UnderVisual.Caption
  Dim b_U As String: b_U = .U_UnderVisual.Caption
  Dim b_C As String: b_C = .C_UnderVisual.Caption
  Dim b_D As String: b_D = .D_UnderVisual.Caption
  Dim b_E As String: b_E = .E_UnderVisual.Caption
  Dim b_F As String: b_F = .F_UnderVisual.Caption
  If (b_S <> "") And (Not b_S Like "*欠陥ﾅｼ*") Then b_S = "△" & b_S
  If (b_U <> "") And (Not b_U Like "*欠陥ﾅｼ*") Then b_U = "△" & b_U
  If (b_C <> "") And (Not b_C Like "*欠陥ﾅｼ*") Then b_C = "△" & b_C
  If (b_D <> "") And (Not b_D Like "*欠陥ﾅｼ*") Then b_D = "△" & b_D
  If (b_E <> "") And (Not b_E Like "*欠陥ﾅｼ*") Then b_E = "△" & b_E
  If (b_F <> "") And (Not b_F Like "*欠陥ﾅｼ*") Then b_F = "△" & b_F
  
  If b_S = "欠陥ﾅｼ" Then b_S = "〇" & b_S
  If b_U = "欠陥ﾅｼ" Then b_U = "〇" & b_U
  If b_C = "欠陥ﾅｼ" Then b_C = "〇" & b_C
  If b_D = "欠陥ﾅｼ" Then b_D = "〇" & b_D
  If b_E = "欠陥ﾅｼ" Then b_E = "〇" & b_E
  If b_F = "欠陥ﾅｼ" Then b_F = "〇" & b_F
  
  '外観AB合体
  Dim ab_S As String
    If (a_S <> "") And (b_S <> "") Then
      ab_S = "A:" & a_S & vbCrLf & "B:" & b_S
    ElseIf (a_S <> "") And (b_S = "") Then
      ab_S = "A:" & a_S & vbCrLf & ""
    ElseIf (a_S = "") And (b_S <> "") Then
      ab_S = "" & vbCrLf & "B:" & b_S
    End If
  Dim ab_U As String
    If (a_U <> "") And (b_U <> "") Then
      ab_U = "A:" & a_U & vbCrLf & "B:" & b_U
    ElseIf (a_U <> "") And (b_U = "") Then
      ab_U = "A:" & a_U & vbCrLf & ""
    ElseIf (a_U = "") And (b_U <> "") Then
      ab_U = "" & vbCrLf & "B:" & b_U
    End If
  Dim ab_C As String
    If (a_C <> "") And (b_C <> "") Then
      ab_C = "A:" & a_C & vbCrLf & "B:" & b_C
    ElseIf (a_C <> "") And (b_C = "") Then
      ab_C = "A:" & a_C & vbCrLf & ""
    ElseIf (a_C = "") And (b_C <> "") Then
      ab_C = "" & vbCrLf & "B:" & b_C
    End If
  Dim ab_D As String
    If (a_D <> "") And (b_D <> "") Then
      ab_D = "A:" & a_D & vbCrLf & "B:" & b_D
    ElseIf (a_D <> "") And (b_D = "") Then
      ab_D = "A:" & a_D & vbCrLf & ""
    ElseIf (a_D = "") And (b_D <> "") Then
      ab_D = "" & vbCrLf & "B:" & b_D
    End If
  Dim ab_E As String
    If (a_E <> "") And (b_E <> "") Then
      ab_E = "A:" & a_E & vbCrLf & "B:" & b_E
    ElseIf (a_E <> "") And (b_E = "") Then
      ab_E = "A:" & a_E & vbCrLf & ""
    ElseIf (a_E = "") And (b_E <> "") Then
      ab_E = "" & vbCrLf & "B:" & b_E
    End If
  Dim ab_F As String
    If (a_F <> "") And (b_F <> "") Then
      ab_F = "A:" & a_F & vbCrLf & "B:" & b_F
    ElseIf (a_F <> "") And (b_F = "") Then
      ab_F = "A:" & a_F & vbCrLf & ""
    ElseIf (a_F = "") And (b_F <> "") Then
      ab_F = "" & vbCrLf & "B:" & b_F
    End If
  
  Dim VisItem As Variant: VisItem = Array(ab_S, ab_U, ab_C, ab_D, ab_E, ab_F)
  Dim VI(1 To 6, 1 To 1)
  For a = 0 To 5
    VI(a + 1, 1) = VisItem(a)
  Next a
  'BH
'  Dim BHItem As Variant: BHItem = Array(.SBH, .UBH, .CBH, .DBH, .EBH)
'  Dim bh(1 To 5, 1 To 1)
'  For a = 0 To 4
'    bh(a + 1, 1) = BHItem(a)
'  Next a
  
  '検入高さ
  Dim KenHItem As Variant: KenHItem = Array(.KenHBox1, .KenHBox2, .KenHBox3, .KenHBox4, .KenHBox5, .KenHBox6)
  Dim KenH As Single
  For a = 0 To 5
    KenH = val(KenHItem(a)) + (KenH)
  Next a
  Dim Yamasu As String
  Select Case True
  Case IsNumeric(.KenHBox7): Yamasu = .KenHBox7 & "枚"
  Case IsNumeric(.KenHBox6): Yamasu = "6山"
  Case IsNumeric(.KenHBox5): Yamasu = "5山"
  Case IsNumeric(.KenHBox4): Yamasu = "4山"
  Case IsNumeric(.KenHBox3): Yamasu = "3山"
  Case IsNumeric(.KenHBox2): Yamasu = "2山"
  Case IsNumeric(.KenHBox1): Yamasu = "1山"
  End Select
  
  sh.Cells(1, 6) = .高さ計算式.Caption '2023.6.9
End With

'原紙基準
Dim LotR As Range:    Set LotR = sh.Cells(2, 3)
Dim AR As Range:      Set AR = sh.Cells(4, 4) 'サイズ
Dim HR As Range:      Set HR = sh.Cells(4, 6)
Dim TR As Range:      Set TR = sh.Cells(4, 8)
Dim AmR As Range:     Set AmR = sh.Cells(5, 3) '-公差
Dim HmR As Range:     Set HmR = sh.Cells(5, 5)
Dim TmR As Range:     Set TmR = sh.Cells(5, 7)
Dim ApR As Range:     Set ApR = sh.Cells(5, 4) '+公差
Dim HpR As Range:     Set HpR = sh.Cells(5, 6)
Dim TpR As Range:     Set TpR = sh.Cells(5, 8)
Dim AmkR As Range:    Set AmkR = sh.Cells(6, 3) '-実寸公差
Dim HmkR As Range:    Set HmkR = sh.Cells(6, 5)
Dim TmkR As Range:    Set TmkR = sh.Cells(6, 7)
Dim ApkR As Range:    Set ApkR = sh.Cells(6, 4) '+実寸公差
Dim HpkR As Range:    Set HpkR = sh.Cells(6, 6)
Dim TpkR As Range:    Set TpkR = sh.Cells(6, 8)
'2024.2.8_追加
Dim KouG As Range:    Set KouG = sh.Cells(6, 9) 'ゲージ公差
Dim KouHa As Range:   Set KouHa = sh.Cells(6, 10) '幅歪公差
Dim KouTa As Range:   Set KouTa = sh.Cells(6, 11) '丈歪公差
Dim KouNa As Range:   Set KouNa = sh.Cells(9, 17) '中歪公差
Dim KouMi As Range:   Set KouMi = sh.Cells(9, 18) '耳歪公差
Dim Resou As Range:   Set Resou = sh.Cells(3, 9) '歪公差参照元
'2026.03.05_フリー反り,対角
Dim Huri As Range:    Set Huri = sh.Cells(8, 12) 'フリー反り公差
Dim Taikak As Range:  Set Taikak = sh.Cells(6, 12) '対角公差

Dim Visual As Range:  Set Visual = sh.Range(sh.Cells(12, 3), sh.Cells(17, 3)) '外観
'Dim BHR As Range:     Set BHR = Range(SH.Cells(12, 11), SH.Cells(16, 11)) 'BH
Dim KHR As Range:     Set KHR = sh.Cells(2, 7) '検入高さ
Dim KYama As Range:   Set KYama = sh.Cells(2, 8) '検入山数

Dim FlZR As Range:    Set FlZR = sh.Range(sh.Cells(12, 5), sh.Cells(17, 5)) '定盤歪
Dim FlGR As Range:    Set FlGR = sh.Range(sh.Cells(12, 6), sh.Cells(17, 6)) 'ｹﾞｰｼﾞ歪

Dim FlZGR As Range:   Set FlZGR = sh.Range(sh.Cells(12, 5), sh.Cells(17, 6)) '定盤ｹﾞｰｼﾞ歪

Dim AtuR As Range:    Set AtuR = sh.Range(sh.Cells(12, 7), sh.Cells(17, 10)) '厚
Dim HTR As Range:     Set HTR = sh.Range(sh.Cells(19, 3), sh.Cells(24, 12)) '巾丈
Dim TaiR As Range:    Set TaiR = sh.Range(sh.Cells(19, 14), sh.Cells(24, 14)) '対角
Dim LateR As Range:   Set LateR = sh.Range(sh.Cells(19, 16), sh.Cells(24, 16)) 'ラテラルボー_2021.07.13
Dim WorkerR As Range: Set WorkerR = sh.Cells(9, 2) '検査員

Dim AngleR As Range:  Set AngleR = sh.Range(sh.Cells(19, 18), sh.Cells(24, 18)) '直角度_2023.11.27


Dim AllDate(), AT(1 To 6, 1 To 4) As Variant, HT(1 To 6, 1 To 10) As Variant, Tai(1 To 6, 0) As Variant, Late(1 To 6, 0) As Variant, FlZ(1 To 6, 0) As Variant, Flg(1 To 6, 0) As Variant
Dim AngleArr(1 To 6, 0) As Variant
AllDate = UF_Array(1, 20, 1, 6) '全格納_1~20'******2022.06耳付き丈追加+2
For i = 1 To 20 'UF_Arrayで格納されている順番
  For a = 1 To 6
    Select Case i
    Case 1 To 4:   AT(a, i) = AllDate(a, i) '厚格納
    Case 5 To 14:  HT(a, i - 4) = AllDate(a, i) '巾丈格納　******2次元
    
'    Case 15:       FlZ(a, i - 15) = AllDate(a, i) '定盤歪格納****1次元
'    Case 16:       Flg(a, i - 16) = AllDate(a, i) 'ｹﾞｰｼﾞ歪格納***1次元
    Case 17:       Tai(a, i - 17) = AllDate(a, i) '対角格納　 ***1次元
    Case 18:       If AllDate(a, i) <> "" Then AngleArr(a, i - 18) = AllDate(a, i)
    Case 20:       If AllDate(a, i) <> "" Then Late(a, i - 20) = AllDate(a, i) & "/L" 'ラテラルボー格納***1次元
    End Select
  Next a
Next i
'------------------------------------------------------------------------------------------------------------
'歪AB面対応_2023.6.14
Dim hizumi_Arry(1 To 6, 1 To 2) As String
Dim hizumAB(1 To 6, 1 To 2) As String
Dim hizumi(): hizumi = UF_Array(15, 16, 1, 6) '格納
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
    If hizumA(i, a) <> "__" Then hizumA(i, a) = "A:" & hizumA(i, a)
  Next a
Next i

For i = LBound(hizumB, 1) To UBound(hizumB, 1) ' Bの各要素に"B:"を追加する
  For a = 1 To 2
    If hizumB(i, a) <> "__" Then hizumB(i, a) = "B:" & hizumB(i, a)
  Next a
Next i
'合体
For i = LBound(hizumAB, 1) To UBound(hizumAB, 1)
  For a = LBound(hizumAB, 2) To UBound(hizumAB, 2)
    hizumAB(i, a) = hizumA(i, a) & vbCrLf & hizumB(i, a)
  Next a
Next i

'----------------------------貼り付け
If UF1.ReSize.Tag = "有" Then sh.Cells(1, 2) = "ｻｲｽﾞ修正"
  LotR = LotDate(0)
  AR = LotDate(1)  'サイズ
  HR = LotDate(2)
  TR = LotDate(3)
  AmR = "'-" & LotDate(4) '-公差'2021.0.27_"’"追加
  HmR = "'-" & LotDate(6)
  TmR = "'-" & LotDate(8)
  ApR = "'+" & LotDate(5) '+公差
  HpR = "'+" & LotDate(7)
  TpR = "'+" & LotDate(9)
  AmkR = LotDate(10) '-実寸公差
  HmkR = LotDate(12)
  TmkR = LotDate(14)
  ApkR = LotDate(11)  '+実寸公差
  HpkR = LotDate(13)
  TpkR = LotDate(15)
'  FlZR = FlZ '定盤歪
'  FlGR = Flg 'ｹﾞｰｼﾞ歪

  KouG = HizKousa(0)  '歪公差
  KouHa = HizKousa(1)
  KouTa = HizKousa(2)
  KouNa = HizKousa(3)
  KouMi = HizKousa(4)
'  HizKousa(5)参照元
  '2026.03.05　　Q-S-CH-DIM-022-R03　正式に反映_2026.06.05
  'フリー反り
  If HizKousa(6) <> "" Then
    Huri = Format(val(HizKousa(6)), "0.0")
     sh.Cells(7, 12) = "フリー反り"
    Huri.Borders.LineStyle = xlContinuous
    With sh.Range(sh.Cells(7, 12), sh.Cells(8, 12))
        .Borders(xlEdgeTop).LineStyle = xlDouble      ' 上：2重線
        .Borders(xlEdgeRight).Weight = xlThick        ' 右：太線
        .Borders(xlEdgeBottom).Weight = xlThick       ' 下：太線
        .Borders(xlEdgeLeft).Weight = xlThick       ' 左：太線
        .NumberFormatLocal = "0.0"
    End With
  End If
  '対角線公差
  Taikak = Format(val(UF1.TaikText), "0.0")
  sh.Cells(4, 12) = "対角"
  sh.Cells(5, 12).Borders(xlEdgeTop).LineStyle = xlDouble
    With sh.Range(sh.Cells(4, 12), sh.Cells(6, 12))
        .Borders(xlEdgeTop).Weight = xlThick          ' 上：太線
        .Borders(xlEdgeRight).Weight = xlThick        ' 右：太線
        .Borders(xlEdgeLeft).Weight = xlThick         ' 左：太線
        .NumberFormatLocal = "0.0"
        If HizKousa(4) <> "" Then
            ' 下：2重線
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        Else
            ' 下：太線
            .Borders(xlEdgeBottom).Weight = xlThick
        End If
    End With
  
'  Dim dTmp As String:         dTmp = HizKousa(5)
'  Dim fiedTmp As String:      fiedTmp = Replace(HizKousa(5), " ", vbCrLf) '改行に差し替え
'  HizKousa(5) = fiedTmp
  Resou = HizKousa(5) '参照元
  Resou.Font.Size = 6
  
  FlZGR = hizumAB '定盤ゲージAB_2023.6.14
  
  AtuR = AT '厚
  HTR = HT '巾丈
  TaiR = Tai '対角
  LateR = Late 'ラテラルボー
  FlZR.Replace "_", "", xlPart
  FlGR.Replace "_", "", xlPart
  
  If UF1.Youto = "R115" Then AngleR = AngleArr '直角度_2023.11.27
  
  Visual = VI '外観
'  Visual.ShrinkToFit = True
  
  Dim Cell As Range
  ' 範囲内の各セルに対してフォントサイズを調整
  For Each Cell In Visual
      FitTextToCell Cell
  Next Cell
  
'  BHR = bh 'BH
  KHR = KenH '検入高さ
  If IsNumeric(UF1.KenHBox7) Then
    sh.Cells(2, 6) = "検入枚数"
    KHR = Yamasu
  Else '枚数ではないとき
    KYama = Yamasu '検入山数
  End If
  
'最終熱処理工程_2023.6.8********6.9_削除予定のため停止 めんどくさ
'Dim StrArray() As String
'If UF1.LastHot.Caption <> "" Then
'  StrArray = Split(UF1.LastHot.Caption, "：") '文字列分割
'  SH.Cells(12, 12) = StrArray(UBound(StrArray))
'End If
  
'----------------------------罫線
Select Case True
Case UF1.厚板
  Call 印刷用罫線("厚板")
Case UF1.薄板
  Call 印刷用罫線("薄板")
Case UF1.少枚数
  Call 印刷用罫線("少枚数")
Case UF1.極厚
  Call 印刷用罫線("極厚")
Case UF1.スラブ
  Call 印刷用罫線("スラブ")
End Select
  
'----------------------------測定機器
Dim Hiki_Atu As String: Hiki_Atu = LotDate(1)
Dim Hiki_Hab As String: Hiki_Hab = LotDate(2) '2023.5.11
Dim YoutC As String: YoutC = UF1.Youto '2023.5.11_用途コードH494に限り幅測定をノギスで行う
Call 機器転送(sh, 9, 5, Hiki_Atu, Hiki_Hab, YoutC)
'----------------------------検査員
Dim Worker
Select Case WorkeDayMax '登録時間が一番遅い
Case "1": Worker = Split(UF1.T1, "_")
Case "2": Worker = Split(UF1.T2, "_")
Case "3": Worker = Split(UF1.T3, "_")
End Select

If IsEmptyArray(Worker) Then 'trueで配列が空
  Call MsgError("実行時エラー！検査員登録をやり直してください。(変数なし)", Lname, 1)
  UF1.PrintB.Tag = "stop"
  Exit Sub
Else
  If WsCheck("WorkerTime" & "_" & Lname) = False Then WorkerR = Worker(UBound(Worker))  '検査履歴から全員ひろう_2020.12.18
  Dim NameDic As Object: Set NameDic = CreateObject("Scripting.Dictionary") '連想配列
  Dim WorkerTime(), WorSh As Worksheet, WorRa As Range
  Set WorSh = ThisWorkbook.Sheets("WorkerTime" & "_" & Lname)
  Set WorRa = WorSh.Range(WorSh.Cells(3, 1), WorSh.Cells(7, 12))
  WorkerTime = WorRa
  n = 0
  For i = LBound(WorkerTime, 1) To UBound(WorkerTime, 1)
    For a = LBound(WorkerTime, 2) To UBound(WorkerTime, 2)
      If WorkerTime(i, a) <> "" Then '検査済みだけ
        If NameDic.exists(WorkerTime(i, a)) = False Then '一致でない時格納
           n = n + 1 '番号
           NameDic(WorkerTime(i, a)) = n
        End If
      End If
    Next a
  Next i
  Dim VarName As Variant, WorList() As Variant, j As Long: j = 0
  If Not NameDic.Count = 0 Then '2021.01.27　WorkerTimeｼｰﾄ上の検査員が0の場合※ありえないがｼｰﾄがｸﾘｱされた場合
    ReDim WorList(1 To NameDic.Count) '測定者履歴ﾘｽﾄ
    For Each VarName In NameDic '重複なしの作業者履歴
      j = j + 1
      WorList(j) = VarName
    Next
    Dim PastWor: PastWor = Join(WorList, "_") '貼り付け用に"_"追加
    WorkerR = PastWor
  End If
End If

'WorSh.Cells(1, 1) = "履歴呼び出し"_2024.6.5 履歴呼び出しに追加

'LoadingはUF1のラベル(Loadingの赤文字)  If UF1.Loading.Tag = "HistoryLoad" Then
'UF1.Loading.Tag = "HistoryLoad" からUF1.dataLoding.Caption = "Load"に変更_2023.8.28
If (UF1.dataLoding.Caption <> "Load") And (WorSh.Cells(1, 1) <> "履歴呼び出し") Then    'HistoryLoadで履歴呼び出しをしていないのに検査員なしの時は止める_2023.05.25
  If WorkerR = "" Then
    Call MsgError("検査員が空白です。WorkerTime上に履歴がない可能性有。 担当者に連絡してください。", Lname, 1)
    UF1.PrintB.Tag = "stop"
    Exit Sub
  End If
ElseIf UF1.dataLoding.Caption = "Load" Or WorSh.Cells(1, 1) = "履歴呼び出し" Then
'履歴から数値をUF1に表示させたときの検査員はその時の現検査員とする_2023.6.15  UFWorkerName.L1
  Dim KarKensa As Variant
  If UFWorkerName.Visible Then KarKensa = UFWorkerName.L1.Caption
  WorkerR = KarKensa
  sh.Cells(3, 2) = "履歴呼び出しから印刷"
End If
WorkerR.ShrinkToFit = True

'ライン**2023.6.7
 sh.Cells(4, 2) = Lname
'異常許容の有無*****2023.8_なかったので追加
Dim NgString As String: NgString = allow_Check

'2024.4.12
If NgString <> "" Then
  If InStr(NgString, "許容") > 0 Then
   NgString = Replace(NgString, "許容", "")
   sh.Cells(10, 2) = "許容公差外在り" & ":" & NgString
  ElseIf InStr(NgString, "保留") > 0 Then
   NgString = Replace(NgString, "保留", "")
   sh.Cells(10, 2) = "公差外 保留処理" & ":" & NgString
  End If
End If

'----------------------------桁設定
Dim x As Long, y As Long
Dim Miflag As Boolean, Koflag As Boolean
If val(sh.Cells(4, 4)) > 74# Then
    sh.Range(sh.Cells(12, 7), sh.Cells(16, 10)).NumberFormatLocal = "0.00_);[赤](0.00)"
ElseIf val(sh.Cells(4, 4)) <= 74# Then
    sh.Range(sh.Cells(12, 7), sh.Cells(16, 10)).NumberFormatLocal = "0.000_);[赤](0.00)"
End If

'2023.5.12_表示桁数品証確認
If UF1.Youto = "H494" And UF1.Hab < 300# Then    '用途コードH494に限り幅測定をノギスで行う
  Select Case val(LotDate(2)) '幅寸法が範囲内の時
  Case 1# To 299#
    sh.Range(sh.Cells(19, 3), sh.Cells(23, 4)).NumberFormatLocal = "0.00_);[赤](0.00)"
  Case Else
    sh.Range(sh.Cells(19, 3), sh.Cells(23, 4)).NumberFormatLocal = "0.00_);[赤](0.0)"
  End Select
End If
'----------------------------異常データがあるときコピーする
Dim ws As Worksheet, IzSh As Worksheet
Dim FL As Boolean, Fl2 As Boolean
For Each ws In ThisWorkbook.Sheets
 If ws.Name = "図" Then FL = True
 If ws.Name = "異常ロット用" & "_" & Lname Then Fl2 = True
Next ws
  If Fl2 Then
    Set IzSh = ThisWorkbook.Sheets("異常ロット用" & "_" & Lname)
    n = UR(IzSh, 1)
    If sh.Cells(2, 3) = IzSh.Cells(1, 1) Then 'ロット番号が同じのときが条件
      If FL Then
        Dim myshape As Variant
        For Each myshape In Sheets("図").Shapes
         myshape.Copy
         sh.Activate
         sh.Cells(11, 12).Select
         sh.Paste
        Next
      End If
    Dim FieldDate As Variant, FieldRange As Range
    
    '2024.12.24_検査形態で最終行が曖昧過ぎるので固定値に修正
    Dim LR As Long: LR = 26 'LR = UR(SH, 3) + 1 'データがある列
    
    sh.Cells(LR, 1) = "公差外ﾃﾞｰﾀ"
    Select Case True
    Case UF1.極厚
      FieldDate = Array("厚1", "厚2", "厚3", "厚4", "幅1", "幅2", "幅3", "幅4", "幅5", "幅6", "丈1", "丈2", "丈3", "丈4")
      IzSh.Range(IzSh.Cells(2, 1), IzSh.Cells(n, 14)).Copy Destination:=sh.Cells(LR + 1, 2)
'    Case UF1.厚板 Or UF1.薄板 Or UF1.少枚数 Or UF1.スラブ  '2022.06_スラブ追加
'      FieldDate = Array("厚1", "厚2", "厚3", "厚4", "幅1", "幅2", "丈1", "丈2", "対角A", "対角B", "直角度", "ﾗﾃﾗﾙﾎﾞｰ", "定盤", "ｹﾞ-ｼﾞ")
'      '印刷の並びに合わせてある ﾗﾃﾗﾙﾎﾞｰ 直角度の順
'      Range(IzSh.Cells(2, 1), IzSh.Cells(n, 10)).Copy Destination:=SH.Cells(25, 2)
'      Range(IzSh.Cells(2, 15), IzSh.Cells(n, 16)).Copy Destination:=SH.Cells(25, 12) 'ﾗﾃﾗﾙﾎﾞｰ 直角度
'      Range(IzSh.Cells(2, 17), IzSh.Cells(n, 18)).Copy Destination:=SH.Cells(25, 14) '定歪　ｹﾞｰｼﾞ
    Case Else 'OP=trueがない場合も
      FieldDate = Array("厚1", "厚2", "厚3", "厚4", "幅1", "幅2", "丈1", "丈2", "対角A", "対角B", "直角度", "ﾗﾃﾗﾙﾎﾞｰ", "定盤", "ｹﾞ-ｼﾞ")
      '印刷の並びに合わせてある ﾗﾃﾗﾙﾎﾞｰ 直角度の順
      IzSh.Range(IzSh.Cells(2, 1), IzSh.Cells(n, 10)).Copy Destination:=sh.Cells(LR + 1, 2)
      IzSh.Range(IzSh.Cells(2, 15), IzSh.Cells(n, 16)).Copy Destination:=sh.Cells(LR + 1, 12) 'ﾗﾃﾗﾙﾎﾞｰ 直角度
      IzSh.Range(IzSh.Cells(2, 17), IzSh.Cells(n, 18)).Copy Destination:=sh.Cells(LR + 1, 14) '定歪　ｹﾞｰｼﾞ
    End Select
    
    Set FieldRange = sh.Range(sh.Cells(LR, 2), sh.Cells(LR, UBound(FieldDate) + 1 + 1))
    FieldRange = FieldDate
    
    IzSh.Range(IzSh.Cells(2, 19), IzSh.Cells(n, 19)).Copy Destination:=sh.Cells(LR + 1, 1) '位置****固定値　NG_Postingで変えたらここも変わる
    
    sh.Range(sh.Cells(LR, 1), sh.Cells(UR(sh, 1), UBound(FieldDate) + 1 + 1)).Font.Size = 8 '異常データサイズ
    sh.Range(sh.Cells(LR, 1), sh.Cells(UR(sh, 1), UBound(FieldDate) + 1 + 1)).HorizontalAlignment = xlLeft
    sh.Range(sh.Cells(LR, 1), sh.Cells(UR(sh, 1), UBound(FieldDate) + 1 + 1)).VerticalAlignment = xlCenter
    
    If val(sh.Cells(4, 4)) > 74# Then
        sh.Range(sh.Cells(LR + 1, 1), sh.Cells(UR(sh, 1), 5)).NumberFormatLocal = "0.00_);[赤](0.00)"
        sh.Range(sh.Cells(LR + 1, 6), sh.Cells(UR(sh, 1), 11)).NumberFormatLocal = "0.0_);[赤](0.0)"
    ElseIf val(sh.Cells(4, 4)) <= 74# Then
        sh.Range(sh.Cells(LR + 1, 1), sh.Cells(UR(sh, 1), 5)).NumberFormatLocal = "0.000_);[赤](0.00)"
        sh.Range(sh.Cells(LR + 1, 6), sh.Cells(UR(sh, 1), 11)).NumberFormatLocal = "0.0_);[赤](0.0)"
    End If
    If UF1.Youto = "H494" And UF1.Hab < 300# Then    '用途コードH494に限り幅測定をノギスで行う
      Select Case val(LotDate(2)) '幅寸法が範囲内の時
      Case 1# To 299#
        sh.Range(sh.Cells(LR + 1, 1), sh.Cells(UR(sh, 1), 5)).NumberFormatLocal = "0.00_);[赤](0.00)"
        sh.Range(sh.Cells(LR + 1, 6), sh.Cells(UR(sh, 1), 11)).NumberFormatLocal = "0.0_);[赤](0.0)"
      Case Else
        sh.Range(sh.Cells(LR + 1, 1), sh.Cells(UR(sh, 1), 5)).NumberFormatLocal = "0.00_);[赤](0.0)"
        sh.Range(sh.Cells(LR + 1, 6), sh.Cells(UR(sh, 1), 11)).NumberFormatLocal = "0.0_);[赤](0.0)"
      End Select
    End If
    
      Call 印刷範囲(WorkerR.Value)   '異常データをコピーしてから範囲設定
    ElseIf sh.Cells(4, 4) <> IzSh.Cells(1, 1) Then '異常ロット用と印刷用のロット番号が違うとき
      Call 印刷範囲2(WorkerR.Value) '範囲設定してから
    End If
  Else
      Call 印刷範囲2(WorkerR.Value) '範囲設定してから
  End If
  
If (WsCheck("不良要因印刷" & "_" & Lname) = False) Then
  Call AddWS("不良要因印刷" & "_" & Lname)
  Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
End If

ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Cells(1, 2) = UF1.Lot

If WsCheck("不良要因印刷" & "_" & Lname) Then Call 不良要因印刷範囲(WorkerR.Value)   '2023.5.29

'PDF改行記号対策_2023.8.21
Dim RewriteRange As Range: Set RewriteRange = sh.Range(sh.Cells(12, 3), sh.Cells(16, 6))
RewriteRange.Replace vbCrLf, vbLf, xlPart '部分一致

Set NameDic = Nothing
Application.ScreenUpdating = True
End Sub

Sub 印刷用罫線(n As String) 'ﾃﾞﾌｫﾙﾄは極厚用
'******対角、ラテラルボーは転送後コピーで転写している
'コピーがフォーマットのみになっている箇所があったので修正**2022.07.12
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("印刷用" & "_" & Lname) = False Then Exit Sub
ThisWorkbook.Sheets("印刷用" & "_" & Lname).Visible = True
ThisWorkbook.Sheets("印刷用" & "_" & Lname).Activate

'2026.2.25_文字列設定ミス修正　不要でも印刷に描写されてしまっていた
Dim NgTemp As String
If InStr(UF1.Additional.Caption, "ﾗﾃﾗﾙﾎﾞｰ検査不要") > 0 Then NgTemp = "検査不要"

With ThisWorkbook.Sheets("印刷用" & "_" & Lname)
  Dim Clea1 As Borders: Set Clea1 = .Range(Cells(20, 2), Cells(20, 10)).Borders
  Dim Clea2 As Borders: Set Clea2 = .Range(Cells(22, 2), Cells(22, 10)).Borders
  Dim Top1 As border: Set Top1 = .Range(Cells(20, 2), Cells(20, 9)).Borders(xlEdgeTop)
  Dim Top2 As border: Set Top2 = .Range(Cells(22, 2), Cells(22, 9)).Borders(xlEdgeTop)
  Dim Bot1 As border: Set Bot1 = .Range(Cells(20, 2), Cells(20, 9)).Borders(xlEdgeBottom)
  Dim Bot2 As border: Set Bot2 = .Range(Cells(22, 2), Cells(22, 9)).Borders(xlEdgeBottom)
  Dim Clea3 As Borders: Set Clea3 = .Range(Cells(13, 3), Cells(13, 6)).Borders
  Dim Clea4 As Borders: Set Clea4 = .Range(Cells(15, 3), Cells(15, 6)).Borders
  Dim Top3 As border: Set Top3 = .Range(Cells(13, 2), Cells(13, 6)).Borders(xlEdgeTop)
  Dim Top4 As border: Set Top4 = .Range(Cells(13, 2), Cells(13, 6)).Borders(xlEdgeBottom)
  Dim Bot3 As border: Set Bot3 = .Range(Cells(15, 2), Cells(15, 6)).Borders(xlEdgeTop)
  Dim Bot4 As border: Set Bot4 = .Range(Cells(15, 2), Cells(15, 6)).Borders(xlEdgeBottom)
  Dim i As Long
  
  If UF1.Naka <> "" Then .Range(.Cells(8, 17), .Cells(9, 18)).Copy Range(.Cells(7, 10), .Cells(8, 11)) '中,耳歪の公差がある場合_2024.2.8
  
  Select Case n
  Case "厚板"
    Clea1.LineStyle = xlLineStyleNone
    Clea2.LineStyle = xlLineStyleNone
'    Clea3.LineStyle = xlLineStyleNone
'    Clea4.LineStyle = xlLineStyleNone
    .Cells(12, 2) = "S"
    .Cells(13, 2) = "S1"
    .Cells(14, 2) = "2"
    .Cells(15, 2) = "E1"
    .Cells(16, 2) = "E"
    .Cells(19, 2) = "S"
    .Cells(20, 2) = ""
    .Cells(21, 2) = "2"
    .Cells(22, 2) = ""
    .Cells(23, 2) = "E"
    .Cells(18, 5) = "丈1"
    .Cells(18, 6) = "丈2"
    .Cells(18, 7) = "対角A"
    .Cells(18, 8) = "対角B"
    .Cells(18, 9) = "差"
    .Range(.Cells(18, 14), .Cells(24, 14)).Copy .Range(.Cells(18, 9), .Cells(24, 9)) '数字もコピー
    .Cells(20, 9).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
    .Cells(20, 9).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
    .Cells(22, 9).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
    .Cells(22, 9).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
    Top1.LineStyle = xlContinuous
    Bot1.LineStyle = xlContinuous
    Top2.LineStyle = xlContinuous
    Bot2.LineStyle = xlContinuous
    Top3.LineStyle = xlContinuous
    Bot3.LineStyle = xlContinuous
    Top4.LineStyle = xlContinuous
    Bot4.LineStyle = xlContinuous
  Case "少枚数"
    .Cells(18, 5) = "丈1"
    .Cells(18, 6) = "丈2"
    .Cells(18, 7) = "対角A"
    .Cells(18, 8) = "対角B"
    .Cells(18, 9) = "差"
    .Range(.Cells(18, 14), .Cells(24, 14)).Copy .Range(.Cells(18, 9), .Cells(24, 9))
  Case "薄板" '2021.11.29_UD削除
      .Cells(12, 2) = "S" '*****検査位置がデフォになっていた
      .Cells(13, 2) = "U"
      .Cells(14, 2) = "2"
      .Cells(15, 2) = "D"
      .Cells(16, 2) = "E"
      .Cells(19, 2) = "S"
      .Cells(20, 2) = "U"
      .Cells(21, 2) = "2"
      .Cells(22, 2) = "D"
      .Cells(23, 2) = "E"
      .Cells(18, 5) = "丈1"
      .Cells(18, 6) = "丈2"
      .Cells(18, 7) = "対角A"
      .Cells(18, 8) = "対角B"
      .Cells(18, 9) = "差"
      .Range(.Cells(18, 14), .Cells(24, 14)).Copy .Range(.Cells(18, 9), .Cells(24, 9)) '対角消えてた_2023.9.8修正
    Select Case UF1.Yama.Text '検査位置設定
    Case 4
      .Cells(12, 2) = "4-S"
      .Cells(13, 2) = "4-E"
      .Cells(14, 2) = "'3-2"
      .Cells(15, 2) = "'2-2"
      .Cells(16, 2) = "1-S"
      .Cells(17, 2) = "1-E"
      
      .Cells(19, 2) = "4-S"
      .Cells(20, 2) = "4-E"
      .Cells(21, 2) = "'3-2"
      .Cells(22, 2) = "'2-2"
      .Cells(23, 2) = "1-S"
      .Cells(24, 2) = "1-E"
    Case 3
      .Cells(12, 2) = "3-S"
      .Cells(13, 2) = "3-E"
      .Cells(14, 2) = "'2-2"
      .Cells(15, 2) = "1-S"
      .Cells(16, 2) = "1-E"
      .Cells(19, 2) = "3-S"
      .Cells(20, 2) = "3-E"
      .Cells(21, 2) = "'2-2"
      .Cells(22, 2) = "1-S"
      .Cells(23, 2) = "1-E"
    Case 2
      .Cells(12, 2) = "2-S"
      .Cells(13, 2) = ""
      .Cells(14, 2) = "2-E"
      .Cells(15, 2) = ""
      .Cells(16, 2) = "1-E"
      .Cells(19, 2) = "2-S"
      .Cells(20, 2) = ""
      .Cells(21, 2) = "2-E"
      .Cells(22, 2) = ""
      .Cells(23, 2) = "1-E"
    Case 1
      .Cells(12, 2) = "S"
      .Cells(13, 2) = ""
      .Cells(14, 2) = "2"
      .Cells(15, 2) = ""
      .Cells(16, 2) = "E"
      .Cells(19, 2) = "S"
      .Cells(20, 2) = ""
      .Cells(21, 2) = "2"
      .Cells(22, 2) = ""
      .Cells(23, 2) = "E"
    End Select
    
    If UF1.Yama.Text = "4" Then '罫線設定****4山は6回検査_2024.5.29追加
      'カードに公差がある場合は欄を出す’2024.8.27　家木指示
      Dim FlagColor As Boolean: FlagColor = False
      If (UF1.LateralText.BackColor = RGB(51, 204, 204)) And (UF1.LateralText.BackStyle = 1) Then FlagColor = True
      
      Dim Lascol As Long
      If UF1.Sumisei = "" Then Lascol = 9
      If UF1.Sumisei <> "" And NgTemp <> "検査不要" Then Lascol = 10
      If FlagColor Then Lascol = 10
    
      Dim rangeE As Range: Set rangeE = .Range(.Cells(16, 2), .Cells(16, Lascol))
      Dim rangeF As Range: Set rangeF = .Range(.Cells(17, 2), .Cells(17, Lascol))
      
      rangeE.Copy
      rangeF.PasteSpecial Paste:=xlPasteFormats '書式のみ
      Application.CutCopyMode = False
      rangeF.Borders(xlEdgeTop).LineStyle = xlDot
      
      Dim rangeEunder As Range: Set rangeEunder = .Range(.Cells(23, 2), .Cells(23, Lascol))
      Dim rangeFunder As Range: Set rangeFunder = .Range(.Cells(24, 2), .Cells(24, Lascol))
      
      rangeEunder.Copy
      rangeFunder.PasteSpecial Paste:=xlPasteFormats '書式のみ
      Application.CutCopyMode = False
      rangeFunder.Borders(xlEdgeTop).LineStyle = xlDot
      .Rows(24).RowHeight = .Rows(23).RowHeight
    End If
    
    '検査回数2回
      If (UF1.InspeButt.Visible = True) And (UF1.InspeButt.Caption = "検入≧26") Then '2023.6.9_条件記載漏れ追記
        .Cells(14, 2) = ""
        .Cells(21, 2) = ""
        .Range(.Cells(18, 14), .Cells(23, 14)).Copy Range(.Cells(18, 10), .Cells(23, 10))
        .Range(.Cells(13, 2), .Cells(15, 11)).Borders.LineStyle = xlNone
        .Range(.Cells(20, 2), .Cells(22, 9)).Borders.LineStyle = xlNone
        
        .Range(.Cells(12, 2), .Cells(12, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(15, 2), .Cells(15, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(19, 2), .Cells(19, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(23, 2), .Cells(23, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
        
        .Range(.Cells(13, 11), .Cells(15, 11)).Borders(xlEdgeTop).LineStyle = xlLineStyleNone
        .Range(.Cells(13, 11), .Cells(15, 11)).Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
      GoTo skipEnd
      End If
    
    If UF1.薄板.Caption = "T" Then '*******************************************2023.6.9_条件追加
      If UF1.Yama.Text = "3" Then '罫線設定****3山は5回検査
        .Range(.Cells(18, 12), .Cells(23, 12)).Borders.LineStyle = xlLineStyleNone
        .Range(.Cells(18, 12), .Cells(23, 12)).ClearContents
        .Range(.Cells(18, 14), .Cells(23, 14)).Copy Range(.Cells(18, 9), .Cells(23, 9))
      ElseIf UF1.Yama.Text = "2" Or UF1.Yama.Text = "1" Or UF1.InspeButt.Visible Then   '1,2山は3回検査**2023.08.28
      
        Clea1.LineStyle = xlLineStyleNone
        Clea2.LineStyle = xlLineStyleNone
        Dim Clea5 As Borders: Set Clea5 = .Range(Cells(13, 2), Cells(13, 10)).Borders
        Dim Clea6 As Borders: Set Clea6 = .Range(Cells(15, 2), Cells(15, 10)).Borders
        Clea5.LineStyle = xlLineStyleNone
        Clea6.LineStyle = xlLineStyleNone
        Dim Top5 As border: Set Top5 = .Range(Cells(13, 2), Cells(13, 10)).Borders(xlEdgeTop)
        Dim Top6 As border: Set Top6 = .Range(Cells(15, 2), Cells(15, 10)).Borders(xlEdgeTop)
        Dim Bot5 As border: Set Bot5 = .Range(Cells(13, 2), Cells(13, 10)).Borders(xlEdgeBottom)
        Dim Bot6 As border: Set Bot6 = .Range(Cells(15, 2), Cells(15, 10)).Borders(xlEdgeBottom)
        Top5.LineStyle = xlContinuous
        Bot5.LineStyle = xlContinuous
        Top6.LineStyle = xlContinuous
        Bot6.LineStyle = xlContinuous
    '    Range(.Cells(18, 12), .Cells(23, 12)).Copy Range(.Cells(18, 9), .Cells(23, 9))
        .Range(.Cells(18, 14), .Cells(23, 14)).Copy .Range(.Cells(18, 9), .Cells(23, 9))
        .Cells(20, 9).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
        .Cells(20, 9).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
        .Cells(22, 9).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
        .Cells(22, 9).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
        Top1.LineStyle = xlContinuous
        Bot1.LineStyle = xlContinuous
        Top2.LineStyle = xlContinuous
        Bot2.LineStyle = xlContinuous
        Top3.LineStyle = xlContinuous
        Bot3.LineStyle = xlContinuous
        Top4.LineStyle = xlContinuous
        Bot4.LineStyle = xlContinuous
      ElseIf Not UF1.InspeButt.Visible Then
  '      Range(.Cells(18, 14), .Cells(23, 14)).Copy
  '      Range(.Cells(18, 9), .Cells(23, 9)).PasteSpecial (xlPasteFormats)
        
        .Range(.Cells(18, 14), .Cells(23, 14)).Copy .Range(.Cells(18, 9), .Cells(23, 9))
      End If
    End If
  Case "スラブ" '2022.06追加
    .Cells(12, 2) = "1"
    .Cells(13, 2) = ""
    .Cells(14, 2) = "2"
    .Cells(15, 2) = ""
    .Cells(16, 2) = ""
    .Cells(19, 2) = "1上"
    .Cells(20, 2) = "1下"
    .Cells(21, 2) = "2上"
    .Cells(22, 2) = "2下"
    .Cells(23, 2) = ""
    .Cells(18, 5) = "丈1"
    .Cells(18, 6) = "丈2"
    .Cells(18, 7) = "対角A"
    .Cells(18, 8) = "対角B"
    .Cells(18, 9) = "差"
    .Range(.Cells(18, 14), .Cells(23, 14)).Copy .Range(.Cells(18, 9), .Cells(23, 9)) '.PasteSpecial (xlPasteFormats)
    .Cells(20, 8).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
    .Cells(22, 8).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
    .Cells(20, 9).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
    .Cells(20, 9).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
    .Cells(22, 9).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
    .Cells(22, 9).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
    
    
    '順番変えると消せない線がある
    .Range(.Cells(15, 2), .Cells(16, 10)).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
    .Range(.Cells(15, 2), .Cells(16, 10)).Borders(xlInsideVertical).LineStyle = xlLineStyleNone
    .Range(.Cells(15, 2), .Cells(16, 10)).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
    .Range(.Cells(15, 2), .Cells(16, 10)).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
    .Range(.Cells(15, 2), .Cells(16, 10)).Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
    .Range(.Cells(14, 2), .Cells(14, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Range(.Cells(14, 2), .Cells(14, 10)).Borders(xlEdgeBottom).Weight = xlMedium
    
    .Range(.Cells(23, 2), .Cells(23, 9)).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
    .Range(.Cells(23, 2), .Cells(23, 9)).Borders(xlInsideVertical).LineStyle = xlLineStyleNone
    .Range(.Cells(23, 2), .Cells(23, 9)).Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
    
    .Range(.Cells(20, 7), .Cells(20, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(.Cells(22, 7), .Cells(22, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(.Cells(20, 7), .Cells(20, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Range(.Cells(22, 7), .Cells(22, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    .Range(.Cells(22, 7), .Cells(22, 9)).Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
    .Range(.Cells(23, 2), .Cells(23, 9)).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
    .Range(.Cells(23, 2), .Cells(23, 9)).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
    .Range(.Cells(22, 2), .Cells(22, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Range(.Cells(22, 2), .Cells(22, 6)).Borders(xlEdgeBottom).Weight = xlMedium
    
    '不要部分
    Dim rngClear As Range:   Set rngClear = .Range(.Cells(18, 10), .Cells(23, 12))
    rngClear.ClearContents
    Dim rngBorder1 As Range: Set rngBorder1 = .Range(.Cells(18, 9), .Cells(19, 9))
    Dim rngBorder2 As Range: Set rngBorder2 = .Cells(21, 9)
    
    ' 罫線をクリア
    Dim border As Variant
    For Each border In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
      rngClear.Borders(border).LineStyle = xlNone
    Next border
    
    With rngBorder1.Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .Weight = xlThick
      .ColorIndex = xlAutomatic
    End With
    With rngBorder2.Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .Weight = xlThick
      .ColorIndex = xlAutomatic
    End With
    
  Case Else '耳付き
    .Range(Cells(18, 9), Cells(23, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range(Cells(18, 9), Cells(23, 9)).Borders(xlEdgeLeft).Weight = xlMedium
  End Select
skipEnd:
    '********************不要文字消去
    '2021.07.13_ﾗﾃﾗﾙﾎﾞｰ追加
    
    Dim ラテ As String:    ラテ = Lateralbow_Check '2021.07.14
     '白銅追加によりフラグ変更　= "住友精密"から<>""**2022.08****'2024.10.08**2026.07.13_FlagColorが薄板4山限定でしか定義されないためTagでも判定'
    If (UF1.Sumisei <> "" And NgTemp <> "検査不要") Or (FlagColor) Or (UF1.LateralText.Tag = "有") Then
      If UF1.極厚 = False And UF1.スラブ = False Then '2024.11.15_耳付き、スラブの際も不要
         '直角度の追加_2023.11.27
        If UF1.Youto = "R115" Then 'ラテラルボーと直角度
          .Range(.Cells(18, 16), .Cells(24, 16)).Copy Range(.Cells(18, 11), .Cells(24, 11))
          .Range(.Cells(18, 18), .Cells(24, 18)).Copy Range(.Cells(18, 10), .Cells(24, 10))
          .Range(.Cells(18, 12), .Cells(24, 14)).Borders.LineStyle = xlLineStyleNone
          .Range(.Cells(18, 12), .Cells(24, 14)).ClearContents
        End If
        
        If ラテ = "" Then '2024.7.1_ラテラルボー測定不要も罫線が出ていたため修正_必ず先に実行
          .Range(.Cells(18, 10), .Cells(24, 14)).ClearContents
          .Range(.Cells(18, 10), .Cells(23, 14)).Borders.LineStyle = xlLineStyleNone
          .Range(.Cells(18, 10), .Cells(23, 14)).ClearContents
        End If
        
        If UF1.Youto <> "R115" Then 'R115の時消してしまうとコピーできない'ラテラルボーのみ
          .Range(.Cells(18, 16), .Cells(24, 16)).Copy Range(.Cells(18, 10), .Cells(24, 10))
          .Range(.Cells(18, 11), .Cells(24, 14)).Borders.LineStyle = xlLineStyleNone
          .Range(.Cells(18, 11), .Cells(24, 14)).ClearContents
        End If
        
        If UF1.Yama.Text = "4" Then
          Set rangeEunder = .Range(.Cells(19, 9), .Cells(19, 11))
          Set rangeFunder = .Range(.Cells(24, 9), .Cells(24, 11))
  
          rangeEunder.Copy
          rangeFunder.PasteSpecial Paste:=xlPasteFormats '書式のみ
          Application.CutCopyMode = False
          rangeFunder.Borders(xlEdgeTop).LineStyle = xlDot
          rangeFunder.Borders(xlEdgeBottom).Weight = xlMedium
          .Rows(24).RowHeight = .Rows(23).RowHeight
          If UF1.Youto <> "R115" Then .Cells(24, 11).Borders(xlEdgeTop).LineStyle = xlNone
          If UF1.Youto <> "R115" Then .Cells(24, 11).Borders(xlEdgeBottom).LineStyle = xlNone
        End If
        
        Select Case n
        Case "厚板"
          .Range(Cells(20, 2), Cells(20, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
          .Range(Cells(22, 2), Cells(22, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
          .Range(Cells(20, 2), Cells(20, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Range(Cells(22, 2), Cells(22, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        End Select
      End If
    Else
        Select Case n '丈3，4追加
        Case "極厚"
          .Range(.Cells(18, 14), .Cells(23, 16)).Borders.LineStyle = xlLineStyleNone
          .Range(.Cells(18, 14), .Cells(23, 16)).ClearContents
        Case Else
          If UF1.Yama.Text = "4" Then
            .Range(.Cells(18, 10), .Cells(24, 14)).Borders.LineStyle = xlLineStyleNone
            .Range(.Cells(18, 10), .Cells(24, 14)).ClearContents
          End If
          .Range(.Cells(18, 10), .Cells(23, 14)).Borders.LineStyle = xlLineStyleNone
          .Range(.Cells(18, 10), .Cells(23, 14)).ClearContents
  '        .Cells(24, 9).Borders(xlEdgeTop).LineStyle = xlDot
        End Select
        
        '納入先captionがない場合でもラテラルボーが必要になった_2024.8.27
        '後に実行　先にやると消しちゃう
        If (UF1.LateralText.BackColor = RGB(51, 204, 204)) And (UF1.LateralText.BackStyle = 1) Then
          If ラテ = "" And UF1.Youto <> "R115" Then
            If UF1.極厚 = False And UF1.スラブ = False Then '2024.11.15
              .Range(.Cells(18, 10), .Cells(24, 14)).ClearContents
              .Range(.Cells(18, 10), .Cells(23, 14)).Borders.LineStyle = xlLineStyleNone
              .Range(.Cells(18, 10), .Cells(23, 14)).ClearContents
              .Range(.Cells(18, 16), .Cells(24, 16)).Copy Range(.Cells(18, 10), .Cells(24, 10))
              .Range(.Cells(18, 11), .Cells(24, 14)).Borders.LineStyle = xlLineStyleNone
              .Range(.Cells(18, 11), .Cells(24, 14)).ClearContents
            End If
          End If
        End If
        
        'AfX5用　薄板は上記を通過しないのでここに追加_2026.02.25
        If (NgTemp = "検査不要") And (UF1.薄板) Then
          .Range(.Cells(18, 10), .Cells(24, 14)).ClearContents
          .Range(.Cells(18, 10), .Cells(23, 14)).Borders.LineStyle = xlLineStyleNone
          .Range(.Cells(18, 10), .Cells(23, 14)).ClearContents
        End If
    End If
    
    
  Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("印刷用" & "_" & Lname)
  If UF1("LZS").Caption = "幅ｹﾞｰｼﾞ" Then
    .Cells(11, 5) = "幅ゲージ" '定盤測定なし_2021.01.14
    .Cells(11, 6) = "丈ゲージ" '2023.03.09_ゲージを幅丈明確にした  LGS
  ElseIf UF1.flatnesText = "ﾌﾘｰｿﾘ要" Then  '2023.4.18_フリー反り追加
    .Cells(11, 6) = "ﾌﾘｰｿﾘ"
    .Range(sh.Cells(12, 6), sh.Cells(16, 6)).Replace "A:", "", xlPart 'フリーそり時はAB面不要_2023.911
    .Range(sh.Cells(12, 6), sh.Cells(16, 6)).Replace "B:", "", xlPart
  End If
End With
End Sub
Public Sub 印刷範囲(Wor As String) '公差外データ記載用
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("印刷用" & "_" & Lname) = False Then Exit Sub
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("印刷用" & "_" & Lname)
Dim n As Long: n = ws.Cells(Rows.Count, 1).End(xlUp).row
Application.PrintCommunication = False

Select Case True '検査頻度選択
Case UF1.極厚: ws.PageSetup.PrintArea = "$A$1:$P$" & n + 1 '2022.06_丈3,4追加により範囲を広げる
Case Else:     ws.PageSetup.PrintArea = "$A$1:$O$" & n + 1
End Select
With ws.PageSetup
  .LeftHeader = "&""Meiryo UI""&8" & Lname
  .CenterHeader = ""
  .RightHeader = "&""Meiryo UI""&8" & ThisWorkbook.Sheets("印刷用" & "_" & Lname).Cells(2, 3) & " " & Wor
  .LeftFooter = ""
  .CenterFooter = ""
  .RightFooter = "&""Meiryo UI""&8" & " " & Format$(UF1.StartTime, "yyyy/mm/dd  h:mm") & "～" & Format$(UF1.EndTime, "yyyy/mm/dd  h:mm") & " " & "印刷"
  .LeftMargin = Application.InchesToPoints(0) '余白をすべて０へ
  .RightMargin = Application.InchesToPoints(0)
  .TopMargin = Application.InchesToPoints(0.4)
  .BottomMargin = Application.InchesToPoints(0)
  .HeaderMargin = Application.InchesToPoints(0)
  .FooterMargin = Application.InchesToPoints(0)
  .PaperSize = xlPaperA4 'サイズ
  .Orientation = xlLandscape '縦xlPortrait　横 xllandscape
  .FitToPagesWide = 1 '1ページで納める
  .FitToPagesTall = 1
End With
Application.PrintCommunication = True
End Sub
Public Sub 印刷範囲2(Wor As String) '異常ロットデータなし
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("印刷用" & "_" & Lname) = False Then Exit Sub
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("印刷用" & "_" & Lname)
Dim n As Long: n = ws.Cells(Rows.Count, 2).End(xlUp).row
Application.PrintCommunication = False
Select Case True '検査頻度選択
Case UF1.極厚: ws.PageSetup.PrintArea = "$B$1:$L$" & n + 1 '2022.06_丈3,4追加により範囲を広げる
Case Else:     ws.PageSetup.PrintArea = "$B$1:$L$" & n + 1 '2026.03.05_公差表示追加
End Select
With ws.PageSetup
  .LeftHeader = "&""Meiryo UI""&8" & Lname
  .CenterHeader = ""
  .RightHeader = "&""Meiryo UI""&8" & ThisWorkbook.Sheets("印刷用" & "_" & Lname).Cells(2, 3) & " " & Wor
  .LeftFooter = ""
  .CenterFooter = ""
  .RightFooter = "&""Meiryo UI""&8" & " " & Format$(UF1.StartTime, "yyyy/mm/dd  h:mm") & "～" & Format$(UF1.EndTime, "yyyy/mm/dd  h:mm") & " " & "印刷"
  .LeftMargin = Application.InchesToPoints(0) '余白をすべて０へ
  .RightMargin = Application.InchesToPoints(0)
  .TopMargin = Application.InchesToPoints(0.4)
  .BottomMargin = Application.InchesToPoints(0)
  .HeaderMargin = Application.InchesToPoints(0)
  .FooterMargin = Application.InchesToPoints(0)
  .PaperSize = xlPaperA4 'サイズ
  .Orientation = xlLandscape '縦xlPortrait　横 xllandscape
  .FitToPagesWide = 1 '1ページで納める
  .FitToPagesTall = 1
  .CenterHorizontally = True '水平
'  .CenterVertically = True '垂直
  .Zoom = 115
End With
Application.PrintCommunication = True
End Sub
'指定ｼｰﾄのコピー　追加シートは AddWSでしておく
Public Sub Pesute(n As String, m As String) '(原紙名、付ける名前)
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
ThisWorkbook.Sheets(n).Unprotect '保護の解除
ThisWorkbook.Sheets(n).Cells.Copy
ThisWorkbook.Worksheets(m).Cells.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
'ﾃﾞﾌｫに戻す
  With ThisWorkbook.Worksheets(m).Cells.Interior '黒塗りをなくす
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With ThisWorkbook.Worksheets(m).Cells.Font '文字黒
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0
  End With
Select Case m
Case "作業中保存"
  ThisWorkbook.Worksheets(m).Activate 'ﾍﾟｰｽﾄ側
  ActiveWindow.Zoom = 120
  ActiveWindow.GridlineColorIndex = 1 '枠線黒
  ActiveWindow.DisplayGridlines = False
  Worksheets(n).Activate 'ｺﾋﾟｰ側
  ActiveWindow.ScrollRow = 1
  ActiveWindow.ScrollColumn = 1
  Worksheets(n).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
Case "ｷｬｽｸ作業中保存"
  ThisWorkbook.Worksheets(m).Activate 'ﾍﾟｰｽﾄ側
  ThisWorkbook.Worksheets(m).Range("B:B,D:D,F:F").ColumnWidth = 9
  ThisWorkbook.Worksheets(m).Range("K:K,M:M,O:O,Q:Q,S:T").ColumnWidth = 9
  ThisWorkbook.Worksheets(m).Range("V:V").ColumnWidth = 18
  ActiveWindow.Zoom = 80
  ActiveWindow.GridlineColorIndex = 1 '枠線黒
  ActiveWindow.DisplayGridlines = False
  Worksheets(n).Activate 'ｺﾋﾟｰ側
  Worksheets(n).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
Case "全ヵ所作業中保存"
'  ThisWorkbook.Worksheets(m).Activate 'ﾍﾟｰｽﾄ側
'  Cells.EntireColumn.AutoFit
''
''  Dim i As Long, columnsToSetWidth As Variant, col As Variant
''  For i = 1 To 30
''    ThisWorkbook.Worksheets(m).Columns(i).ColumnWidth = 8.75
''  Next i
''  columnsToSetWidth = Array(3, 5, 7, 9, 10, 12, 14, 16, 18, 21, 24, 26, 28)
''  For Each col In columnsToSetWidth
''      ThisWorkbook.Worksheets(m).Columns(col).ColumnWidth = 2.5
''  Next col
''
'  ActiveWindow.Zoom = 85
'  ActiveWindow.GridlineColorIndex = 1 '枠線黒
'  ActiveWindow.DisplayGridlines = False
'  Worksheets(n).Activate 'ｺﾋﾟｰ側
'  Worksheets(n).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
  ThisWorkbook.Worksheets(m).Activate 'ﾍﾟｰｽﾄ側
  ActiveWindow.Zoom = 85
  ActiveWindow.GridlineColorIndex = 1 '枠線黒
  ActiveWindow.DisplayGridlines = False
  Worksheets(n).Activate 'ｺﾋﾟｰ側
  Worksheets(n).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
Case "H494作業中保存"
  ThisWorkbook.Worksheets(m).Activate 'ﾍﾟｰｽﾄ側
  ActiveWindow.Zoom = 100
  ActiveWindow.GridlineColorIndex = 1 '枠線黒
  ActiveWindow.DisplayGridlines = False
  Worksheets(n).Activate 'ｺﾋﾟｰ側
  Worksheets(n).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True '保護設定
Case "印刷用" & "_" & Lname
  ThisWorkbook.Worksheets(m).Activate 'ﾍﾟｰｽﾄ側
  ActiveWindow.Zoom = 100
Case "不良要因印刷" & "_" & Lname
  ThisWorkbook.Worksheets(m).Activate 'ﾍﾟｰｽﾄ側
  ActiveWindow.Zoom = 70
End Select
End Sub

Public Sub 不良要因印刷範囲(Wor As String, Optional LotN As String = "")
Dim Lname As String, LotN_C As String
If LotN = "" Then
  Lname = UF1.LineName.Caption & UF1.LineNo.Caption
  LotN_C = UF1.Lot
Else
  Lname = UFMitsubishi.LineName.Caption
  LotN_C = UFMitsubishi.LotBox
End If
If WsCheck("不良要因印刷" & "_" & Lname) = False Then Exit Sub
Dim ws As Worksheet:    Set ws = ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname)
'時間定義----UF1使用でないときに必要
Dim StartA As Date
If UF1.Visible And UF1.StartTime.Caption <> "" Then StartA = Format$(UF1.StartTime.Caption, "yyyy/mm/dd h:mm")
If UFMitsubishi.Visible Then StartA = Format$(ThisWorkbook.Sheets("全ヵ所検査").Cells(1, 2), "yyyy/mm/dd h:mm")
If UFCask.Visible Then StartA = Format$(ThisWorkbook.Sheets("キャスク検査").Cells(1, 2), "yyyy/mm/dd h:mm")

Dim EndT As Date
Dim dt As String:       dt = DateAdd("d", -1, Now) '夜勤の日付またぎ前日扱い
Dim strDate As String:  strDate = Format$(dt, "yyyy/mm/dd hh:mm:ss")
Dim NowTim As Date:     NowTim = Format$(Now, "yyyy/mm/dd hh:mm:ss")
If Time > CDate(Start1) And Time <= CDate("23:59") Then '0時前は押したときの日付
  EndT = NowTim
ElseIf Time >= CDate("00:00") And Time <= CDate(End3) Then '0時～夜勤終了まで1日戻して処理する
  EndT = Format$(strDate, "yyyy/mm/dd hh:mm:ss")
Else
  EndT = NowTim
End If

Dim n As Long: n = ws.Cells(Rows.Count, 2).End(xlUp).row
Application.PrintCommunication = False
ws.PageSetup.PrintArea = "$A$1:$Y$" & n + 1
With ws.PageSetup
  .LeftHeader = "&""Meiryo UI""&8" & Lname
  .CenterHeader = ""
  .RightHeader = "&""Meiryo UI""&8" & LotN_C & " " & Wor
  .LeftFooter = ""
  .CenterFooter = ""
  If LotN = "" Then .RightFooter = "&""Meiryo UI""&8" & " " & Format$(StartA, "yyyy/mm/dd  h:mm") & "～" & Format$(EndT, "yyyy/mm/dd  h:mm") & " " & "印刷"
  .LeftMargin = Application.InchesToPoints(0) '余白をすべて０へ
  .RightMargin = Application.InchesToPoints(0)
  .TopMargin = Application.InchesToPoints(0)
  .BottomMargin = Application.InchesToPoints(0)
  .HeaderMargin = Application.InchesToPoints(0)
  .FooterMargin = Application.InchesToPoints(0)
  .PaperSize = xlPaperA4 'サイズ
  .Orientation = xlLandscape '縦xlPortrait　横 xllandscape
  .FitToPagesWide = 1 '1ページで納める
  .FitToPagesTall = 1
  .CenterHorizontally = True '水平
  .CenterVertically = True '垂直
  .Zoom = 100
End With
Application.PrintCommunication = True

Call CalculateColumnSums(ws) '不良要因合計計算
Call ConvertNumbersToString(ws) '数字を囲い文字数字に変換
End Sub
'不良要因合計計算
Public Sub CalculateColumnSums(ws As Worksheet)
If WsCheck(ws.Name) = False Then Exit Sub
Dim lastRow As Long: lastRow = 7
    
Dim Cols As Long
Dim Rows As Long
Dim SumValue As Single


' 列3から列24までの行3から行7までの合計を計算
For Cols = 3 To 24
  SumValue = 0
  For Rows = 3 To 8
    If IsNumeric(ws.Cells(Rows, Cols).Value) Then
      SumValue = SumValue + ws.Cells(Rows, Cols).Value
    End If
  Next Rows
  ws.Cells(10, Cols).Value = SumValue ' 合計を最終行に記入
Next Cols

'**********50を超えた数字は囲い文字にならずにそのままとなる

End Sub
'数字を囲い文字数字に変換****印刷の前にのみ実行
Public Sub ConvertNumbersToString(ws As Worksheet)
If WsCheck(ws.Name) = False Then Exit Sub
Dim TempR1 As Range: Set TempR1 = ws.Range(ws.Cells(3, 3), ws.Cells(8, 27))
Dim TempArr(): TempArr = TempR1
Dim TempR2 As Range: Set TempR2 = ws.Range(ws.Cells(10, 3), ws.Cells(10, 27))
Dim TempArrG(): TempArrG = TempR2
Dim i As Long, a As Long, num As Long
Dim Flag1 As Boolean: Flag1 = False
Dim Flag2 As Boolean: Flag2 = False

ReDim ConvArr(LBound(TempArr, 1) To UBound(TempArr, 1), LBound(TempArr, 2) To UBound(TempArr, 2))

For i = LBound(TempArr, 2) To UBound(TempArr, 2)
  For a = LBound(TempArr, 1) To UBound(TempArr, 1)
    If (TempArr(a, i) <> "") And (IsNumeric(TempArr(a, i))) Then
      num = TempArr(a, i)
      ConvArr(a, i) = CircledNumber(num) '書き換え
      Flag1 = True
    End If
  Next a
Next i
If Flag1 Then TempR1 = ConvArr '貼り付け

ReDim ConvArr(LBound(TempArrG, 1) To UBound(TempArrG, 1), LBound(TempArrG, 2) To UBound(TempArrG, 2))
For i = LBound(TempArrG, 2) To UBound(TempArrG, 2)
  For a = LBound(TempArrG, 1) To UBound(TempArrG, 1)
    If (TempArrG(a, i) <> "") And (IsNumeric(TempArrG(a, i))) Then
      num = TempArrG(a, i)
      ConvArr(a, i) = CircledNumber(num)
      Flag2 = True
    End If
    If TempArrG(a, i) = 0 Then ConvArr(a, i) = ""
  Next a
Next i
If Flag2 Then TempR2 = ConvArr
End Sub
Public Function Call_BordersLineStyleWeightCopyPaste(CopyRange As Range, PasteRange As Range) 'セルの罫線をコピーアンドペーストする
PasteRange.Borders(xlEdgeTop).LineStyle = CopyRange.Borders(xlEdgeTop).LineStyle
PasteRange.Borders(xlEdgeTop).Weight = CopyRange.Borders(xlEdgeTop).Weight
'PasteRange.Borders(xlEdgeBottom).LineStyle = CopyRange.Borders(xlEdgeBottom).LineStyle
'PasteRange.Borders(xlEdgeBottom).Weight = CopyRange.Borders(xlEdgeBottom).Weight
PasteRange.Borders(xlEdgeLeft).LineStyle = CopyRange.Borders(xlEdgeLeft).LineStyle
PasteRange.Borders(xlEdgeLeft).Weight = CopyRange.Borders(xlEdgeLeft).Weight
PasteRange.Borders(xlInsideHorizontal).LineStyle = CopyRange.Borders(xlInsideHorizontal).LineStyle
PasteRange.Borders(xlInsideHorizontal).Weight = CopyRange.Borders(xlInsideHorizontal).Weight
'PasteRange.Borders(xlInsideVertical).LineStyle = CopyRange.Borders(xlInsideVertical).LineStyle
'PasteRange.Borders(xlInsideVertical).Weight = CopyRange.Borders(xlInsideVertical).Weight
PasteRange.Borders(xlEdgeRight).LineStyle = CopyRange.Borders(xlEdgeRight).LineStyle
PasteRange.Borders(xlEdgeRight).Weight = CopyRange.Borders(xlEdgeRight).Weight
End Function
'デフォルトサイズ8でフォントサイズを調整
Public Sub FitTextToCell(TargetCell As Range)
Dim FontSize As Double
Dim CellWidth As Double, CellHeight As Double
Dim TempShape As Shape
' セルの幅と高さを取得
CellWidth = TargetCell.Width
CellHeight = TargetCell.Height

FontSize = 8 ' デフォルトフォントサイズを設定

' 一時的なShapeを作成して文字列のサイズを測定
Set TempShape = TargetCell.Worksheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 100, 100)
TempShape.TextFrame2.TextRange.Text = TargetCell.Value
TempShape.line.Visible = msoFalse ' テキストボックスの枠線を非表示にする

' フォントサイズを調整（幅と高さの両方を考慮）
Do
  With TempShape.TextFrame2.TextRange.Font
      .Name = TargetCell.Font.Name
      .Size = FontSize
  End With

  ' テキストボックスの幅と高さを計算
  TempShape.LockAspectRatio = msoFalse
  TempShape.Width = CellWidth
  TempShape.Height = CellHeight

  ' 幅または高さのどちらかがセルを超える場合のみ、フォントサイズを縮小
  If TempShape.TextFrame2.TextRange.BoundWidth > CellWidth Or TempShape.TextFrame2.TextRange.BoundHeight > CellHeight Then
      FontSize = FontSize - 0.5 ' フォントサイズを縮小
  Else
      ' 収まっている場合はサイズ縮小を止める
      Exit Do
  End If
Loop While TempShape.TextFrame2.TextRange.BoundWidth > CellWidth Or TempShape.TextFrame2.TextRange.BoundHeight > CellHeight

If FontSize >= 4# Then ' フォントサイズが4.5以下の場合、そのままにし、4.5以上なら8に戻す
  FontSize = 8
ElseIf FontSize = 3.5 Then
  FontSize = 6
End If

If FontSize < 1 Then ' フォントサイズが1以下にならないように制限
    FontSize = 1
End If

TargetCell.Font.Size = FontSize ' フォントサイズをセルに適用
TempShape.Delete ' 一時オブジェクトを削除
End Sub
