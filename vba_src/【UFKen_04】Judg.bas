Attribute VB_Name = "【UFKen_04】Judg"
Option Explicit
#If VBA7 And Win64 Then
  '64ビット版
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
  '32ビット版
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'========================================2026.09.17
' 製品規格 公差参照（品質基準.accdb）
'   呼出側互換：Diagonal / ReDiagonal / Lateralbow_J / ReLateralbow
'   初回呼出時に GetFieldsArr / GetRecordsArr でテーブルを読込み、以後はメモリ上で検索
'========================================
Private Const TBL_DIAG As String = "製品規格直角度"
Private Const TBL_LAT As String = "製品規格ラテラルボー"

'--- 製品規格直角度：列名リスト内の位置
Private Const D_NO As Long = 0, D_GRADE As Long = 1
Private Const D_W_LO As Long = 2, D_W_HI As Long = 3
Private Const D_L_LO As Long = 4, D_L_HI As Long = 5
Private Const D_TOL As Long = 6

'--- 製品規格ラテラルボー：列名リスト内の位置
Private Const L_NO As Long = 0, L_GRADE As Long = 1
Private Const L_T_LO As Long = 2, L_T_HI As Long = 3
Private Const L_W_LO As Long = 4, L_W_HI As Long = 5
Private Const L_L_LO As Long = 6, L_L_HI As Long = 7
Private Const L_TOL As Long = 8

Private m_Diag As Variant          'GetRecordsArr の結果 (行, 列)
Private m_Lat As Variant
Private m_DiagIdx() As Long        '上の位置 → 実際の列番号
Private m_LatIdx() As Long
Private m_Loaded As Boolean
Private m_LoadFailed As Boolean



Public Sub Judg_ImageSize(n As String) '判定画像　Randomize 画像数を変える度に修正要
Dim x As Long: x = Int(15 * Rnd(1)) + 1 '(最大値-最小値+1)*Rnd+最小値
If 0 >= x Or x >= 16 Then Exit Sub
Dim Hsize As Long: Hsize = 500
Dim Wsize As Long: Wsize = 500
Dim Tpos As Long: Tpos = 90
Dim Lpos As Long: Lpos = 245

If UF1.VC_A.Visible Or UF1.VC_B.Visible Then Tpos = Tpos + 55

UF1.VC_A.Visible = False
UF1.VC_B.Visible = False
UF1.VC_L_A.Visible = False
UF1.VC_L_B.Visible = False
UF1.MP1.Width = 0

Select Case n '位置とサイズ調整
Case "Y"
  UF1.IMOK.Height = Hsize
  UF1.IMOK.Width = Wsize
  UF1.IMOK.tOP = Tpos
  UF1.IMOK.Left = Lpos
  UF1.IMOK.Visible = True
  UF1.WorkFrame.Visible = False
  UF1.NoFrame.Visible = False
  UF1.MP1.Width = 11 '438---2021.02.15_薄板選択時非表示にできない、原因不明のためサイズ修正
  '2022.01.28_追加
  UF1.NoL.Visible = False
  UF1.PackingForm.Visible = False
  UF1.PackingComment.Visible = False
  
  If UF1.MPEX.Tag <> "Click" Then '2023.8.18_エラー対策
    UF1.MPEX.Visible = False '2023.5.17_追加
    UF1.MP1.Visible = False
  Else
    UF1.MPEX.Height = 1
    UF1.MPEX.Width = 1
  End If
  
Case "N"
  UF1.IMNG.Height = Hsize
  UF1.IMNG.Width = Wsize
  UF1.IMNG.tOP = Tpos
  UF1.IMNG.Left = Lpos
  UF1.IMNG.Visible = True
  UF1.MP1.Width = 1 '438
  UF1.WorkFrame.Visible = False
  UF1.NoFrame.Visible = False
  '2022.01.28_追加
  UF1.NoL.Visible = False
  UF1.PackingForm.Visible = False
  UF1.PackingComment.Visible = False
  
  If UF1.MPEX.Tag <> "Click" Then '2023.8.18_エラー対策
    UF1.MPEX.Visible = False '2023.5.17_追加
    UF1.MP1.Visible = False
  Else
    UF1.MPEX.Height = 1
    UF1.MPEX.Width = 1
  End If
  
Case "Neko"
  UF1("NekoText").tOP = 240
  UF1("NekoText").Left = 645
  UF1("NekoText").Font.Size = 35
  UF1("NekoText").AutoSize = True
  UF1("NekoText").Visible = True
  Randomize
  UF1("NekoIM" & x).Height = Hsize
  UF1("NekoIM" & x).Width = 350
  UF1("NekoIM" & x).tOP = Tpos
  UF1("NekoIM" & x).Left = Lpos
  UF1("NekoIM" & x).Visible = True
  UF1.MP1.Width = 0 '438
  UF1.WorkFrame.Visible = False
  UF1.NoFrame.Visible = False
  '2022.01.28_追加
  UF1.NoL.Visible = False
  UF1.PackingForm.Visible = False
  UF1.PackingComment.Visible = False
  
  If UF1.MPEX.Tag <> "Click" Then '2023.8.18_エラー対策
    UF1.MPEX.Visible = False '2023.5.17_追加
    UF1.MP1.Visible = False
  Else
    UF1.MPEX.Height = 1
    UF1.MPEX.Width = 1
  End If
End Select

UF1.Repaint
End Sub

Public Sub HizumiJudg(TB As MSForms.TextBox, Optional LoopJ As String = "") '歪の判定
'TB.Tagは歪のAB切り替えで使用しているので安易に使うと死ぬ
'UF1.MoveFlag = TB.Tag & "_" & TB.Name '確認用　必要なければコメントアウト
Dim Scol As Long: Scol = 13434828 'OK RGB(204,255,204)薄い緑
Dim Ncol As Long: Ncol = 8421631 'NG RGB(255,128,128)ｺｰﾗﾙ

If TB = "" Then
  TB.BackColor = RGB(255, 255, 255) '白
  Exit Sub
End If

Dim ComparisonValue As String, ComValueG As String
Dim Index As Long, NumberT As String
' コントロールの取得
Dim LorR As String:              LorR = Left(TB.Name, 1)  'Z、G
Dim Temp As String:              Temp = Right(TB.Name, 1) 'S、U、C、D、E、F
Dim TargetComboxZ As Object:     Set TargetComboxZ = UF1("CB" & Temp & "1") '定盤'ComboBox"CBS1"
Dim TargetComboxG As Object:     Set TargetComboxG = UF1("CB" & Temp & "3") '定盤'ComboBox"CBS3"
Dim FrameControl As Object:      Set FrameControl = UF1("FM" & Temp) 'Frame"FMS"

'2026.03.05_追加（フリー反り判定をTargetComboxZ未選択でも実行できるよう先頭に独立させる）
If LorR = "G" And UF1.HURIS <> "" And IsNumeric(UF1.HURIS.Value) And UF1.flatnesText = "ﾌﾘｰｿﾘ要" Then
  Dim NumberT2 As String, Index2 As Long
  Index2 = InStr(TB, "/")
  If Index2 > 0 Then
    NumberT2 = Left(TB, Index2 - 1)
  Else
    NumberT2 = TB
  End If
  If IsNumeric(NumberT2) Then
    Dim HurisVal As Double
    HurisVal = CDbl(UF1.HURIS.Value)
    If CDbl(NumberT2) > HurisVal Then
      TB.BackColor = Ncol
      If LoopJ = "" Or UF1.MoveFlag = "on" Then Call 音楽不合格
    Else
      TB.BackColor = Scol
      If LoopJ = "" Or UF1.MoveFlag = "on" Then Call 音楽合格
    End If
  End If
  UF1.MoveFlag = ""
  Exit Sub
End If

If TargetComboxZ.Value = "" Then Exit Sub

' 比較値の取得
If TargetComboxZ.Value = "幅" Then
  ComparisonValue = UF1.HaHiz.Value
ElseIf TargetComboxZ.Value = "丈" Then
  ComparisonValue = UF1.TaHiz.Value
ElseIf TargetComboxZ.Value = "中" Then
  ComparisonValue = UF1.Naka.Value
ElseIf TargetComboxZ.Value = "耳" Then
  ComparisonValue = UF1.Mimi.Value
End If
'--------------------------------------------------------------------------------------------------------------------------------------------------------
'5点で吊れない(重量オーバー)かつ6点も使えない(丈が短い)
'6点不可(重量オーバー)かつ10点も使えない(丈が短い)
'10点不可(重量オーバー)
'6,10点不可(幅が短い)かつ5点で吊れない(重量オーバー)
'重量を問わず丈が6000mmを超える場合(定盤に乗せる意味がない)
'表面切削なしバキューム不可
'上記条件時はゲージ測定を定盤公差と比較させる
If UF1.LZS.Caption = "幅ｹﾞｰｼﾞ" And UF1.LGS.Caption = "丈ｹﾞｰｼﾞ" Then '両方ゲージ測定だが比較対象は幅丈
  If TargetComboxZ.Value = "幅" Then
    ComparisonValue = UF1.HaHiz.Value
  ElseIf TargetComboxZ.Value = "丈" Then
    ComparisonValue = UF1.TaHiz.Value
  End If
  
  If TargetComboxG.Value = "幅" Then
    ComValueG = UF1.HaHiz.Value
  ElseIf TargetComboxG.Value = "丈" Then
    ComValueG = UF1.TaHiz.Value
  End If
  
  If LorR = "Z" And IsNumeric(ComparisonValue) Then '定盤側に入力→比較対象幅
    If CDbl(TB) > ComparisonValue Then
      TB.BackColor = Ncol  'NG
      If LoopJ = "" Or UF1.MoveFlag = "on" Then Call 音楽不合格 'UF1.MoveFlag = "on"はMouseDownイベントで都度実行
    ElseIf CDbl(TB) <= ComparisonValue Then
      TB.BackColor = Scol 'OK
      If LoopJ = "" Or UF1.MoveFlag = "on" Then Call 音楽合格
    End If
  ElseIf LorR = "G" Then 'ゲージ側に入力→比較対象丈
If TargetComboxG.Value = "" Then Exit Sub
    If CDbl(TB) > ComValueG Then
      TB.BackColor = Ncol  'NG
      If LoopJ = "" Or UF1.MoveFlag = "on" Then Call 音楽不合格 'UF1.MoveFlag = "on"はMouseDownイベントで都度実行
    ElseIf CDbl(TB) <= ComValueG Then
      TB.BackColor = Scol 'OK
      If LoopJ = "" Or UF1.MoveFlag = "on" Then Call 音楽合格
    End If
  End If
UF1.MoveFlag = "" 'リセット
Exit Sub
End If
'--------------------------------------------------------------------------------------------------------------------------------------------------------

'定盤入力なしでゲージから行った場合
If ComparisonValue = "" Then
  Set TargetComboxZ = UF1("CB" & Temp & "3")
  If TargetComboxZ.Value <> "" Then GoTo SkipG 'スキップ
End If

If IsNumeric(ComparisonValue) = False Then Exit Sub

SkipG:

'xx/x山の場合/より前を抜き取る
Index = InStr(TB, "/")
If Index > 0 Then NumberT = Left(TB, Index - 1)
If Index = 0 Then NumberT = TB

If IsNumeric(NumberT) = False Then Exit Sub
If NumberT = "" Then Exit Sub

If LorR = "Z" And IsNumeric(ComparisonValue) Then '定盤判定
  '該当Box色の変更
  If CDbl(NumberT) > ComparisonValue Then
    TB.BackColor = Ncol  'NG
    If LoopJ = "" Or UF1.MoveFlag = "on" Then Call 音楽不合格 'UF1.MoveFlag = "on"はMouseDownイベントで都度実行
  ElseIf CDbl(NumberT) <= ComparisonValue Then
    TB.BackColor = Scol 'OK
    If LoopJ = "" Or UF1.MoveFlag = "on" Then Call 音楽合格
  End If
'2026.03.05_追加
'2026.06.05_追加（先頭側に独立移動済みのためここでは判定しない）
'ElseIf LorR = "G" And UF1.HURIS <> "" And IsNumeric(UF1.HURIS.Value) And UF1.flatnesText = "ﾌﾘｰｿﾘ要" Then 'ゲージ判定(フリーソリ)
'  HurisVal = CDbl(UF1.HURIS.Value)
'  If CDbl(NumberT) > HurisVal Then
'    TB.BackColor = Ncol  'NG
'    If LoopJ = "" Or UF1.MoveFlag = "on" Then Call 音楽不合格
'  ElseIf CDbl(NumberT) <= HurisVal Then
'    TB.BackColor = Scol  'OK
'    If LoopJ = "" Or UF1.MoveFlag = "on" Then Call 音楽合格
'  End If
  
  'ゲージの判定は不要だった為停止
'ElseIf LorR = "G" Then 'ゲージ判定
'  ComparisonValue = UF1.MG.value
'  If ComparisonValue <> "" Then '公差がある時だけ
'    If CDbl(NumberT) > ComparisonValue Then
'      TB.BackColor = Ncol
'      If LoopJ = "" Or UF1.MoveFlag = "on" Then Call 音楽不合格
'    ElseIf CDbl(NumberT) <= ComparisonValue Then
'      TB.BackColor = Scol
'      If LoopJ = "" Or UF1.MoveFlag = "on" Then Call 音楽合格
'    End If
'  End If
Else
  TB.BackColor = RGB(255, 255, 255) '白
End If

UF1.MoveFlag = "" 'リセット

End Sub

Public Sub HizLoopJudg(Optional n As Long = 0, Optional tMP As String = "Loop")
If n >= 2 Then Exit Sub ' nのリセットはUF1のイニシャライズのみで行う
If UF1.Lot = "" Then Exit Sub

Dim i As Long, a As Long
Dim Scol As Long: Scol = 13434828 'OK RGB(204,255,204)薄い緑
Dim Ncol As Long: Ncol = 8421631  'NG RGB(255,128,128)ｺｰﾗﾙ
If (UF1.HaHiz = "") And (UF1.TaHiz = "") Then Exit Sub
If UF校正.Visible = True Then Exit Sub
With UF1
  If .ZS <> "" Then Call HizumiJudg(.ZS, tMP)
  If .ZU <> "" Then Call HizumiJudg(.ZU, tMP)
  If .ZC <> "" Then Call HizumiJudg(.ZC, tMP)
  If .ZD <> "" Then Call HizumiJudg(.ZD, tMP)
  If .ZE <> "" Then Call HizumiJudg(.ZE, tMP)
  If .ZF <> "" Then Call HizumiJudg(.ZF, tMP)
  
  If .GS <> "" Then Call HizumiJudg(.GS, tMP)
  If .GU <> "" Then Call HizumiJudg(.GU, tMP)
  If .GC <> "" Then Call HizumiJudg(.GC, tMP)
  If .GD <> "" Then Call HizumiJudg(.GD, tMP)
  If .GE <> "" Then Call HizumiJudg(.GE, tMP)
  If .GF <> "" Then Call HizumiJudg(.GF, tMP)
  
  If .ZS = "" Then .ZS.BackColor = RGB(255, 255, 255) '白
  If .ZU = "" Then .ZU.BackColor = RGB(255, 255, 255)
  If .ZC = "" Then .ZC.BackColor = RGB(255, 255, 255)
  If .ZD = "" Then .ZD.BackColor = RGB(255, 255, 255)
  If .ZE = "" Then .ZE.BackColor = RGB(255, 255, 255)
  If .ZF = "" Then .ZF.BackColor = RGB(255, 255, 255)
  
  If .GS = "" Then .GS.BackColor = RGB(255, 255, 255)
  If .GU = "" Then .GU.BackColor = RGB(255, 255, 255)
  If .GC = "" Then .GC.BackColor = RGB(255, 255, 255)
  If .GD = "" Then .GD.BackColor = RGB(255, 255, 255)
  If .GE = "" Then .GE.BackColor = RGB(255, 255, 255)
  If .GF = "" Then .GF.BackColor = RGB(255, 255, 255)
End With

End Sub

Public Function AdvanceCheck(Temp As String) As Boolean '2023.12.22_先行データが必要な場合True
If Temp = "" Then Exit Function
Dim SikaFields:   SikaFields = Arrbox("SIKALOTNOW.accdb", "", "Fields", UF1.DBflag.Text) 'UF1.DBflag.text
If IsEmpty(SikaFields) Then Exit Function
Dim SikaAcc:      SikaAcc = Arrbox("SIKALOTNOW.accdb", Temp, "Record", UF1.DBflag.Text)
If IsEmpty(SikaAcc) Then Exit Function

Dim i As Long, C_品質グレード_表面処理 As Long, C_製造板厚 As Long, C_用途 As Long
For i = LBound(SikaFields, 2) To UBound(SikaFields, 2)
  Select Case SikaFields(1, i)
  Case "品質ｸﾞﾚｰﾄﾞ_表面処理":    C_品質グレード_表面処理 = i
  Case "製造板厚":               C_製造板厚 = i
  Case "用途ｺｰﾄﾞ":               C_用途 = i
  End Select
Next i
Dim flag(1 To 3) As Boolean
Dim 厚 As Single:      厚 = val(SikaAcc(LBound(SikaAcc, 1), C_製造板厚))
Dim 表面 As String:    表面 = SikaAcc(LBound(SikaAcc, 1), C_品質グレード_表面処理)
Dim 用途 As String:    用途 = SikaAcc(LBound(SikaAcc, 1), C_用途)

If 厚 <= 3# Then flag(1) = True '3mm以下
If 表面 <> "S" And 表面 <> "T" Then flag(2) = True '表面処理S,T以外
If (Left(用途, 1) = "T" Or Left(用途, 2) = "D1" Or Left(用途, 2) = "D2") Then flag(3) = True '用途ｺｰﾄﾞ頭T,D1,D2

If flag(1) And flag(2) Then AdvanceCheck = True
If flag(2) And flag(3) Then AdvanceCheck = True

End Function
Public Function AHTJudg(Ter As MSForms.Label, KP As Double, KM As Double) As Boolean
If IsNumeric(Ter.Caption) = False Then Exit Function
If (val(Ter.Caption) < KM) Or (val(Ter.Caption) > KP) Then
  AHTJudg = False '×
ElseIf (val(Ter.Caption) >= KM) And (val(Ter.Caption) <= KP) Then
  AHTJudg = True '○
End If
End Function
Public Function ReDiag() '2022.02.01_データ再読み込みの時の再計算
If Not UF1.極厚 Then
  If (UF1("S" & 9) <> "") And (UF1("S" & 10) <> "") Then UF1.DiagS = Format(Abs(val(UF1("S" & 9)) - val(UF1("S" & 10))), "0.0")
  If (UF1("U" & 9) <> "") And (UF1("U" & 10) <> "") Then UF1.DiagU = Format(Abs(val(UF1("U" & 9)) - val(UF1("U" & 10))), "0.0")
  If (UF1("C" & 9) <> "") And (UF1("C" & 10) <> "") Then UF1.DiagC = Format(Abs(val(UF1("C" & 9)) - val(UF1("C" & 10))), "0.0")
  If (UF1("D" & 9) <> "") And (UF1("D" & 10) <> "") Then UF1.DiagD = Format(Abs(val(UF1("D" & 9)) - val(UF1("D" & 10))), "0.0")
  If (UF1("E" & 9) <> "") And (UF1("E" & 10) <> "") Then UF1.DiagE = Format(Abs(val(UF1("E" & 9)) - val(UF1("E" & 10))), "0.0")
  If (UF1("F" & 9) <> "") And (UF1("F" & 10) <> "") Then UF1.DiagF = Format(Abs(val(UF1("F" & 9)) - val(UF1("F" & 10))), "0.0")
End If
End Function
Public Function DiagonalJudg(Tai As Single, kou As Single) As Boolean
If IsNumeric(Tai) = False Then Exit Function
If kou = 0 Then Exit Function

If Tai = 0 Then
  DiagonalJudg = True '合格
  Exit Function
End If

If val(kou) >= val(Tai) Then
 DiagonalJudg = True '○
ElseIf val(kou) < val(Tai) Then
 DiagonalJudg = False '×
End If
End Function
Public Function LateralJudg(a As Single, kou As String) As Boolean '絶対値評価
If IsNumeric(a) = False Then Exit Function
If kou = "0" Or kou = "" Then Exit Function

If a = 0 Then
  LateralJudg = True '合格
  Exit Function
End If

If val(kou) >= Abs(val(a)) Then
 LateralJudg = True '○
ElseIf val(kou) < Abs(val(a)) Then
 LateralJudg = False '×
End If
End Function
Public Function BhJudg(Ter As MSForms.TextBox, Zai As String, TYO As String) As String '_2023.6.27
Dim bh As Long: bh = BhSearch(Zai, TYO) '下限値
If bh = 0 Or Ter.Text = "" Then 'DB上にない材質
  BhJudg = "不要"
  Exit Function
End If
If bh > val(Ter.Text) Then BhJudg = "NG" '下限値より下NG
If bh <= val(Ter.Text) Then BhJudg = "OK"
End Function

Public Sub simplyjudge(TB As MSForms.TextBox, n As String) '合格不合格の判定色　音　画像だけを行う_2023.6.27
Dim Scol As Long: Scol = 13434828 '合格
Dim Ncol As Long: Ncol = 8421631
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim Iti As String, i As Long
Dim WorkerName As String: WorkerName = GetWorkerName()

Dim countL(1 To 5) As Boolean
With UF1 '最終検査位置の選定
  For i = 1 To 5
    Select Case i
    Case 1: If (.S1.Caption <> "") And (.U1.Caption = "") And (.C1.Caption = "") And (.D1.Caption = "") And (.E1.Caption = "") Then countL(1) = True
    Case 2: If (.U1.Caption <> "") And (.C1.Caption = "") And (.D1.Caption = "") And (.E1.Caption = "") Then countL(2) = True
    Case 3: If (.C1.Caption <> "") And (.D1.Caption = "") And (.E1.Caption = "") Then countL(3) = True
    Case 4: If (.D1.Caption <> "") And (.E1.Caption = "") Then countL(4) = True
    Case 5: If .E1.Caption <> "" Then countL(5) = True
    End Select
  Next i
  Select Case True
  Case countL(1)
    If (TB.Name = "SBH") Then Iti = .S.Caption
  Case countL(5)
    If (TB.Name = "EBH") Then Iti = .E.Caption
   End Select
End With

Dim fname As String:  fname = "Measured.log"
Dim path As String:   path = HistoryPath
Dim Lpath As String
Select Case Lname
Case "L-1":    Lpath = path & "L-1\L-1_Log"
Case "LVC":    Lpath = path & "LVC\LVC_Log"
Case "HVC":    Lpath = path & "HVC\HVC_Log"
Case "ﾊﾞﾗﾝｻｰ": Lpath = path & "バランサー\バランサー_Log"
Case "ﾄｯﾄ":    Lpath = path & "トット\トット_Log"
Case "Test":   Lpath = path & "Test\Test_Log"
Case Else:     Lpath = path & "丸徳\丸徳_Log"
End Select

'2026.05.14_WriteLog引数追加
Select Case n
Case "OK"
  TB.BackColor = Scol
  Call 音楽合格
  Call WriteLog("*BH|" & TB & "|(" & TB.Name & ")" & "|○|" & Lname & "|" & WorkerName, "", Lpath, fname)
  Call WriteLog("*BH|" & TB & "|(" & TB.Name & ")" & "|○|" & Lname & "|" & WorkerName, "")
Case "NG"
  TB.BackColor = Ncol
  Call 音楽不合格
  Call WriteLog("*BH|" & TB & "|(" & TB.Name & ")" & "|×|" & Lname & "|" & WorkerName, "", Lpath, fname)
  Call WriteLog("*BH|" & TB & "|(" & TB.Name & ")" & "|×|" & Lname & "|" & WorkerName, "")
  If Iti = UF1.S.Caption Then Call MsgError("作業中止・保留・ロット異常連絡票発行", Lname, 1)
  If Iti = UF1.E.Caption Then Call MsgError("梱包完了させてAインプット・保留・ロット異常連絡票発行", Lname, 1)
Case "不要"
  TB.BackColor = 16777215
Exit Sub
End Select
UF1.Repaint
End Sub

'2026.05.14_2022/11/14 08:53:39 NLM-FMV50291  *UFKen|12.001|(C3)|○|L-1|検査太郎
'タイムスタンプ自体が 08:53:39 : を含む　|に変更
Public Sub JudgColor(TB As MSForms.Label, n As String) '判定色　音　画像
On Error Resume Next
Dim Scol As Long: Scol = 13434828 '合格
Dim Ncol As Long: Ncol = 8421631
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim WorkerName As String: WorkerName = GetWorkerName()

Dim KP As Double, KM As Double, kou As Single, Tai As Single
Dim MyWaitTime As Date: MyWaitTime = TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + 2)
With UF1
  Select Case n
  Case "A"
    KP = .ApK
    KM = .AmK
  Case "H"
    KP = .HpK
    KM = .HmK
  Case "T"
    KP = .TpK
    KM = .TmK
  Case "D" '公差の追記を優先させるため3,4等級削除_2022.07
  ' Select Case .Taik
  ' Case "4"
    kou = .TaikText 'Diagonal(.Hab, .Tak) '4級(厳しい)
  ' Case "3"
  '  kou = .TaikText 'ReDiagonal(.Hab, .Tak)
  ' End Select
  End Select
End With


Dim fname As String: fname = "Measured.log"
Dim path As String:  path = HistoryPath
Dim Lpath As String
Select Case Lname
Case "L-1":    Lpath = path & "L-1\L-1_Log"
Case "LVC":    Lpath = path & "LVC\LVC_Log"
Case "HVC":    Lpath = path & "HVC\HVC_Log"
Case "ﾊﾞﾗﾝｻｰ": Lpath = path & "バランサー\バランサー_Log"
Case "ﾄｯﾄ":    Lpath = path & "トット\トット_Log"
Case "Test":   Lpath = path & "Test\Test_Log"
Case Else:     Lpath = path & "丸徳\丸徳_Log"
End Select

Dim NowSizeH As Long:  NowSizeH = UF1.MPEX.Height '2024.5.30_追加
Dim NowSizeW As Long:  NowSizeW = UF1.MPEX.Width
Dim NowSizeMP As Long: NowSizeMP = UF1.MP1.Width
'2025.8.8_公差数値無しの場合スルー
Select Case n
Case "A", "H", "T" '厚巾丈

  If n = "A" And UF1.Am = "" Then GoTo SkipD
  If n = "H" And UF1.Hm = "" Then GoTo SkipD
  If n = "T" And UF1.Tm = "" Then GoTo SkipD
  
'2026.05.14_WriteLog引数追加
'buf(0) = タイムスタンプ+PC+マーカー　buf(1) = 測定値　buf(2) = 位置　buf(3) = ○/×　buf(4) = ライン名　buf(5) = 検査員名（省略時 UBound=4）
  If AHTJudg(TB, KP, KM) Then
    Call WriteLog("*UFKen|" & TB.Caption & "|(" & TB.Name & ")" & "|○|" & Lname & "|" & WorkerName, "", Lpath, fname)
    Call WriteLog("*UFKen|" & TB.Caption & "|(" & TB.Name & ")" & "|○|" & Lname & "|" & WorkerName, "")
    TB.BackColor = Scol
    Call 音楽合格
    Call Judg_ImageSize("Y")
    UF1.Repaint '描写対策_2020.12.18
'    Application.Wait MyWaitTime'-----PC負荷的にSleepに変更_2020.12.18

    Sleep 2000

    UF1.IMOK.Visible = False
    UF1.Repaint
    UF1.MP1.Visible = True
    UF1.MP1.Width = NowSizeMP '2021.2.15_薄板の時非表示にできない....ｻｲｽﾞ修正で逃
    
    UF1.MPEX.Height = NowSizeH
    UF1.MPEX.Width = NowSizeW
    UF1.MPEX.Visible = True
     
    UF1.Repaint
  ElseIf AHTJudg(TB, KP, KM) = False Then
    Call WriteLog("*UFKen|" & TB.Caption & "|(" & TB.Name & ")" & "|×|" & Lname & "|" & WorkerName, "", Lpath, fname)
    Call WriteLog("*UFKen|" & TB.Caption & "|(" & TB.Name & ")" & "|×|" & Lname & "|" & WorkerName, "")
    
    TB.BackColor = Ncol
    Call 音楽不合格
    Call Judg_ImageSize("N")
    UF1.Repaint
'    Application.Wait MyWaitTime

    Sleep 2000

    UF1.IMNG.Visible = False
    UF1.Repaint
    UF1.MP1.Visible = True
    UF1.MP1.Width = NowSizeMP '2021.2.15_薄板の時非表示にできない....ｻｲｽﾞ修正で逃

    
    UF1.MPEX.Height = NowSizeH
    UF1.MPEX.Width = NowSizeW
    UF1.MPEX.Visible = True
     
    UF1.Repaint
  End If
Case "D" '対角---------------TextJudgで一つ目の対角測定
  Select Case Right((TB.Name), 1)
  Case 9
    If UF1(Left(TB.Name, 1) & 10) = "" Then
      Call WriteLog("*UFKen:" & TB.Caption & ":(" & TB.Name & ")" & ":_:" & Lname, "", Lpath, fname) '判定前　2020.12.17追加修正
      Call WriteLog("*UFKen:" & TB.Caption & ":(" & TB.Name & ")" & ":_:" & Lname, "")
      GoTo SkipD
      Exit Sub
    End If
    Tai = Abs(val(UF1(Left(TB.Name, 1) & 10)) - val(TB))
    If DiagonalJudg(Tai, kou) Then
      Call WriteLog("*UFKen|" & TB.Caption & "|(" & TB.Name & ")" & "|○|" & Lname & "|" & WorkerName, "", Lpath, fname)
      Call WriteLog("*UFKen|" & TB.Caption & "|(" & TB.Name & ")" & "|○|" & Lname & "|" & WorkerName, "")
      UF1(Left(TB.Name, 1) & 10).BackColor = Scol
      TB.BackColor = Scol
      Call 音楽合格
      Call Judg_ImageSize("Y")
      UF1.Repaint
  '    Application.Wait MyWaitTime
  
      Sleep 2000
  
      UF1.IMOK.Visible = False
      UF1.Repaint
      UF1.MP1.Visible = True
      UF1.MP1.Width = NowSizeMP '2021.2.15_薄板の時非表示にできない....ｻｲｽﾞ修正で逃
  
      
      UF1.MPEX.Height = NowSizeH
      UF1.MPEX.Width = NowSizeW
      UF1.MPEX.Visible = True
       
      UF1.Repaint
      
    ElseIf DiagonalJudg(Tai, kou) = False Then
      Call WriteLog("*UFKen|" & TB.Caption & "|(" & TB.Name & ")" & "|×|" & Lname & "|" & WorkerName, "", Lpath, fname)
      Call WriteLog("*UFKen|" & TB.Caption & "|(" & TB.Name & ")" & "|×|" & Lname & "|" & WorkerName, "")
      UF1(Left(TB.Name, 1) & 10).BackColor = Ncol
      TB.BackColor = Ncol
      Call 音楽不合格
      Call Judg_ImageSize("N")
      UF1.Repaint
  '    Application.Wait MyWaitTime
  
      Sleep 2000
  
      UF1.IMNG.Visible = False
      UF1.Repaint
      UF1.MP1.Visible = True
      UF1.MP1.Width = NowSizeMP '2021.2.15_薄板の時非表示にできない....ｻｲｽﾞ修正で逃
  
      
      UF1.MPEX.Height = NowSizeH
      UF1.MPEX.Width = NowSizeW
      UF1.MPEX.Visible = True
       
      UF1.Repaint
      '------------------------------------------品質ｸﾞﾚｰﾄﾞ_直角度で3,4の手入力が不要になった　2020.06.23
'      UFdiagonal.PosiA = UF1.INP.Tag 'ﾌｫｰｶｽ位置
'      UFdiagonal.PosiB = UF1(Left(TB.Name, 1) & 10).Name
'      UFdiagonal.NumA = TB
'      UFdiagonal.NumB = UF1(Left(TB.Name, 1) & 10)
'      UFdiagonal.TAIKOU = Kou
'      UFdiagonal.Show vbModeless
'      UF1.Hide
    End If
  Case 0
    If UF1(Left(TB.Name, 1) & 9) = "" Then Exit Sub
    Tai = Abs(val(UF1(Left(TB.Name, 1) & 9)) - val(TB))
    If DiagonalJudg(Tai, kou) Then
      Call WriteLog("*UFKen|" & TB.Caption & "|(" & TB.Name & ")" & "|○|" & Lname & "|" & WorkerName, "", Lpath, fname)
      Call WriteLog("*UFKen|" & TB.Caption & "|(" & TB.Name & ")" & "|○|" & Lname & "|" & WorkerName, "")
      
      UF1(Left(TB.Name, 1) & 9).BackColor = Scol
      TB.BackColor = Scol
      Call 音楽合格
      Call Judg_ImageSize("Y")
      UF1.Repaint
  '    Application.Wait MyWaitTime
  
      Sleep 2000
  
      UF1.IMOK.Visible = False
      UF1.Repaint
      UF1.MP1.Visible = True
      UF1.MP1.Width = NowSizeMP '2021.2.15_薄板の時非表示にできない....ｻｲｽﾞ修正で逃
  
      
      UF1.MPEX.Height = NowSizeH
      UF1.MPEX.Width = NowSizeW
      UF1.MPEX.Visible = True
       
      UF1.Repaint
      
    ElseIf DiagonalJudg(Tai, kou) = False Then
      Call WriteLog("*UFKen|" & TB.Caption & "|(" & TB.Name & ")" & "|×|" & Lname & "|" & WorkerName, "", Lpath, fname)
      Call WriteLog("*UFKen|" & TB.Caption & "|(" & TB.Name & ")" & "|×|" & Lname & "|" & WorkerName, "")
      
      UF1(Left(TB.Name, 1) & 9).BackColor = Ncol
      TB.BackColor = Ncol
      Call 音楽不合格
      Call Judg_ImageSize("N")
      UF1.Repaint
  '    Application.Wait MyWaitTime
  
      Sleep 2000
  
      UF1.IMNG.Visible = False
      UF1.Repaint
      UF1.MP1.Visible = True
      UF1.MP1.Width = NowSizeMP '2021.2.15_薄板の時非表示にできない....ｻｲｽﾞ修正で逃
      
      UF1.MPEX.Height = NowSizeH
      UF1.MPEX.Width = NowSizeW
      UF1.MPEX.Visible = True
       
      UF1.Repaint
'      UFdiagonal.PosiA = UF1(Left(TB.Name, 1) & 9).Name
'      UFdiagonal.PosiB = UF1.INP.Tag 'ﾌｫｰｶｽ位置
'      UFdiagonal.NumA = UF1(Left(TB.Name, 1) & 9)
'      UFdiagonal.NumB = TB
'      UFdiagonal.TAIKOU = Kou
'      UFdiagonal.Show vbModeless
'      UF1.Hide
    End If
  End Select
  
UF1("Diag" & Left(TB.Name, 1)) = Tai

'2023.11.24_対角公差から直角度に変更
'UF1("Diag" & Left(TB.Name, 1) & "Kou") = kou
End Select

SkipD:

UF1.MPEX.Visible = True
UF1.Repaint

If UF1.Test Then
  UF1.WorkFrame.Visible = True
  UF1.NoFrame.Visible = True
End If

'2022.01.28_追加
UF1.NoL.Visible = True
UF1.PackingForm.Visible = True
UF1.PackingComment.Visible = True


'2023.5.17_追加
If UF1.VC_L_A.Tag <> "" Then
  UF1.VC_A.Visible = True
  UF1.VC_L_A.Visible = True
End If
If UF1.VC_L_B.Tag <> "" Then
  UF1.VC_B.Visible = True
  UF1.VC_L_B.Visible = True
End If

If UF1.VC_L_A.Tag = "" And UF1.VC_L_B.Tag = "" Then Call VC_Picture

UF1.MPEX.Visible = True
UF1.Repaint
End Sub
Public Sub TextJudg(TB As MSForms.Label) '検査判定用
If UFMsgBox.Visible = True Then Exit Sub '2020.12.24_追加
Dim MyWaitTime As Date: MyWaitTime = TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + 2)
Dim i As Long
If TB.Caption = "" Then Exit Sub
If UF1.Abno.Tag = "NG" Then Exit Sub
With UF1
'厚
If (Len(TB.Name) = 2) Then
  Select Case Right((TB.Name), 1)
  Case 1, 2, 3, 4
   Call JudgColor(TB, "A")
  End Select
End If
'巾
If .極厚 Then
  If (Len(TB.Name) = 2) Then
    Select Case Right((TB.Name), 1)
    Case 5 To 9
      Call JudgColor(TB, "H")
    End Select
  ElseIf (Len(TB.Name) = 3) Then
    Select Case Right((TB.Name), 2)
    Case 10
      Call JudgColor(TB, "H")
    End Select
  End If
Else
  If (Len(TB.Name) = 2) Then
    Select Case Right((TB.Name), 1)
    Case 5, 6
      Call JudgColor(TB, "H")
    End Select
  End If
End If
'丈
If .極厚 Then
  If (Len(TB.Name) = 3) Then
    Select Case Right((TB.Name), 2)
    Case 11, 12, 13, 14 '丈追加+2_2022.06
      Call JudgColor(TB, "T")
    End Select
  End If
Else
  If (Len(TB.Name) = 2) Then
    Select Case Right((TB.Name), 1)
    Case 7, 8
      Call JudgColor(TB, "T")
    End Select
  End If
End If
'対角
If Not .極厚 Then
  If (Len(TB.Name) = 3) Then
    Select Case Right((TB.Name), 2)
    Case 10
      Call JudgColor(TB, "D")
    End Select
  ElseIf (Len(TB.Name) = 2) Then
    Select Case Right((TB.Name), 1)
    Case 9
      Call 音楽対角
      Call Judg_ImageSize("Neko")
      UF1.Repaint
      Sleep 2000
      For i = 1 To 15 '画像を増やす場合要修正
        If UF1("NekoIM" & i).Visible = True Then
        UF1("NekoIM" & i).Visible = False
        Exit For
        End If
      Next i
      UF1("NekoText").Visible = False '猫画像がどちらでもtextは消す
      UF1.Repaint
      UF1.MP1.Visible = True
      UF1.MP1.Width = 438
      If UF1.Test Then
        UF1.WorkFrame.Visible = True
        UF1.NoFrame.Visible = True
      End If
      UF1.Repaint
      Call JudgColor(TB, "D") '2020.12.17_追加修正 9のログだけ残っていなかった
    End Select
  End If
End If
Call NGErase(TB)
End With
End Sub

Public Sub LoopJudg(Optional n As Long = 0, Optional j As Long = 1, Optional k As Long = 6) '引数追加_2023.6.27  2024.5.17_引数追加
If n >= 2 Then Exit Sub '2023.6.27 nのリセットはUF1のイニシャライズのみで行う

Dim i As Long, a As Long, x As Long, y As Long
Dim Scol As Long:     Scol = 13434828 'OK RGB(204,255,204)薄い緑
Dim Ncol As Long:     Ncol = 8421631 'NG RGB(255,128,128)ｺｰﾗﾙ
Dim whiteT As Long:   whiteT = 16777215
Dim BorCol As String: BorCol = RGB(0, 0, 0)
Dim BacCol As String: BacCol = RGB(255, 255, 255)

If (UF1.ApK = "") Or (UF1.AmK = "") Then Exit Sub
If UF校正.Visible = True Then Exit Sub

'枠線ﾘｾｯﾄ
For i = 1 To 14 '丈追加+2_2022.06
  UF1("S" & i).BorderColor = BorCol
  UF1("S" & i).BackColor = BacCol
  UF1("U" & i).BorderColor = BorCol
  UF1("U" & i).BackColor = BacCol
  UF1("C" & i).BorderColor = BorCol
  UF1("C" & i).BackColor = BacCol
  UF1("D" & i).BorderColor = BorCol
  UF1("D" & i).BackColor = BacCol
  UF1("E" & i).BorderColor = BorCol
  UF1("E" & i).BackColor = BacCol
Next i
For i = 1 To 10
  UF1("F" & i).BorderColor = BorCol
  UF1("F" & i).BackColor = BacCol
  If UF1("F" & i) <> "" Then UF1("F" & i).BackColor = Ncol
Next i

'正以外の色付けをする
For i = 1 To 14 '丈追加+2_2022.06
  If UF1("S" & i) <> "" Then UF1("S" & i).BackColor = Ncol
  If UF1("U" & i) <> "" Then UF1("U" & i).BackColor = Ncol
  If UF1("C" & i) <> "" Then UF1("C" & i).BackColor = Ncol
  If UF1("D" & i) <> "" Then UF1("D" & i).BackColor = Ncol
  If UF1("E" & i) <> "" Then UF1("E" & i).BackColor = Ncol
Next i
For i = 1 To 10
  If UF1("F" & i) <> "" Then UF1("F" & i).BackColor = Ncol
Next i

'2025.8.8
If UF1.Am = "" Then
  For i = 1 To 4
    If UF1("S" & i) <> "" Then UF1("S" & i).BackColor = whiteT
    If UF1("U" & i) <> "" Then UF1("U" & i).BackColor = whiteT
    If UF1("C" & i) <> "" Then UF1("C" & i).BackColor = whiteT
    If UF1("D" & i) <> "" Then UF1("D" & i).BackColor = whiteT
    If UF1("E" & i) <> "" Then UF1("E" & i).BackColor = whiteT
  Next i
End If
If UF1.Hm = "" Then
  For i = 5 To 6
    If UF1("S" & i) <> "" Then UF1("S" & i).BackColor = whiteT
    If UF1("U" & i) <> "" Then UF1("U" & i).BackColor = whiteT
    If UF1("C" & i) <> "" Then UF1("C" & i).BackColor = whiteT
    If UF1("D" & i) <> "" Then UF1("D" & i).BackColor = whiteT
    If UF1("E" & i) <> "" Then UF1("E" & i).BackColor = whiteT
  Next i
End If
If UF1.Tm = "" Then
  For i = 7 To 8
    If UF1("S" & i) <> "" Then UF1("S" & i).BackColor = whiteT
    If UF1("U" & i) <> "" Then UF1("U" & i).BackColor = whiteT
    If UF1("C" & i) <> "" Then UF1("C" & i).BackColor = whiteT
    If UF1("D" & i) <> "" Then UF1("D" & i).BackColor = whiteT
    If UF1("E" & i) <> "" Then UF1("E" & i).BackColor = whiteT
  Next i
End If

'ラテラルボーに限りﾃﾞﾌｫ色(UFConfiが起動する)_2021.07.16
'2023.11.28_ラテ直角度のUFConfi起動判定追加(Out_Tolerance)
If UF1.LateralS.Visible Then UF1("LateralS").BackColor = Ncol
If UF1.LateralU.Visible Then UF1("LateralU").BackColor = Ncol
If UF1.LateralC.Visible Then UF1("LateralC").BackColor = Ncol
If UF1.LateralD.Visible Then UF1("LateralD").BackColor = Ncol
If UF1.LateralE.Visible Then UF1("LateralE").BackColor = Ncol
If UF1.LateralF.Visible Then UF1("LateralF").BackColor = Ncol

'UF1("SBH").BackColor = &H80000005 'BF_2023.6.27
'UF1("UBH").BackColor = &H80000005
'UF1("CBH").BackColor = &H80000005
'UF1("DBH").BackColor = &H80000005
'UF1("EBH").BackColor = &H80000005

If UF1.DiagSKou.Visible Then UF1("DiagSKou").BackColor = Ncol '直角度_2023.11.27
If UF1.DiagUKou.Visible Then UF1("DiagUKou").BackColor = Ncol
If UF1.DiagCKou.Visible Then UF1("DiagCKou").BackColor = Ncol
If UF1.DiagEKou.Visible Then UF1("DiagDKou").BackColor = Ncol
If UF1.DiagDKou.Visible Then UF1("DiagEKou").BackColor = Ncol
If UF1.DiagFKou.Visible Then UF1("DiagFKou").BackColor = Ncol
'ここから判定し正の色付けをする
'厚
For i = 1 To 4
  For a = j To k
  If UF1.Am = "" Then GoTo SkipH '2025.8.8
    Select Case a
    Case 1: If AHTJudg(UF1("S" & i), UF1.ApK, UF1.AmK) Then UF1("S" & i).BackColor = Scol
    Case 2: If AHTJudg(UF1("U" & i), UF1.ApK, UF1.AmK) Then UF1("U" & i).BackColor = Scol
    Case 3: If AHTJudg(UF1("C" & i), UF1.ApK, UF1.AmK) Then UF1("C" & i).BackColor = Scol
    Case 4: If AHTJudg(UF1("D" & i), UF1.ApK, UF1.AmK) Then UF1("D" & i).BackColor = Scol
    Case 5: If AHTJudg(UF1("E" & i), UF1.ApK, UF1.AmK) Then UF1("E" & i).BackColor = Scol
    Case 6: If AHTJudg(UF1("F" & i), UF1.ApK, UF1.AmK) Then UF1("F" & i).BackColor = Scol
    End Select
  Next a
Next i

SkipH:

'巾
If UF1.極厚 Then
 x = 10
Else
 x = 6
End If
For i = 5 To x
  For a = j To k
  If UF1.HpK = "" Or UF1.HmK = "" Then GoTo SkipT  '2025.8.8　2026.09.18_耳付きの時エラーになる　条件変更
    Select Case a
    Case 1: If AHTJudg(UF1("S" & i), UF1.HpK, UF1.HmK) Then UF1("S" & i).BackColor = Scol
    Case 2: If AHTJudg(UF1("U" & i), UF1.HpK, UF1.HmK) Then UF1("U" & i).BackColor = Scol
    Case 3: If AHTJudg(UF1("C" & i), UF1.HpK, UF1.HmK) Then UF1("C" & i).BackColor = Scol
    Case 4: If AHTJudg(UF1("D" & i), UF1.HpK, UF1.HmK) Then UF1("D" & i).BackColor = Scol
    Case 5: If AHTJudg(UF1("E" & i), UF1.HpK, UF1.HmK) Then UF1("E" & i).BackColor = Scol
    Case 6: If AHTJudg(UF1("F" & i), UF1.HpK, UF1.HmK) Then UF1("F" & i).BackColor = Scol
    End Select
  Next a
Next i

SkipT:

'丈
If UF1.極厚 Then
  x = 14 '丈追加+2_2022.06
  y = 11
Else
  x = 8
  y = 7
End If

For i = y To x
  Select Case i
  Case 7 To 8
    For a = j To k
    If UF1.TpK = "" Or UF1.TmK = "" Then GoTo SkipL  '2025.8.8　　2026.09.18_耳付きの時エラーになる　条件変更
      Select Case a
      Case 1: If AHTJudg(UF1("S" & i), UF1.TpK, UF1.TmK) Then UF1("S" & i).BackColor = Scol
      Case 2: If AHTJudg(UF1("U" & i), UF1.TpK, UF1.TmK) Then UF1("U" & i).BackColor = Scol
      Case 3: If AHTJudg(UF1("C" & i), UF1.TpK, UF1.TmK) Then UF1("C" & i).BackColor = Scol
      Case 4: If AHTJudg(UF1("D" & i), UF1.TpK, UF1.TmK) Then UF1("D" & i).BackColor = Scol
      Case 5: If AHTJudg(UF1("E" & i), UF1.TpK, UF1.TmK) Then UF1("E" & i).BackColor = Scol
      Case 6: If AHTJudg(UF1("F" & i), UF1.TpK, UF1.TmK) Then UF1("F" & i).BackColor = Scol
      End Select
    Next a
  Case 11 To 14
    For a = j To k
      Select Case a
      Case 1: If AHTJudg(UF1("S" & i), UF1.TpK, UF1.TmK) Then UF1("S" & i).BackColor = Scol
      Case 2: If AHTJudg(UF1("U" & i), UF1.TpK, UF1.TmK) Then UF1("U" & i).BackColor = Scol
      Case 3: If AHTJudg(UF1("C" & i), UF1.TpK, UF1.TmK) Then UF1("C" & i).BackColor = Scol
      Case 4: If AHTJudg(UF1("D" & i), UF1.TpK, UF1.TmK) Then UF1("D" & i).BackColor = Scol
      Case 5: If AHTJudg(UF1("E" & i), UF1.TpK, UF1.TmK) Then UF1("E" & i).BackColor = Scol
      End Select
    Next a
  End Select
Next i

'対角
If IsNumeric(UF1.TaikText) Then '2022.02.01_対角公差参照変更
  For a = j To k
    Select Case a
    Case 1
    If (UF1.DiagS <> "") Then
     If DiagonalJudg(UF1.DiagS, UF1.TaikText) Then
      UF1("S" & 9).BackColor = Scol
      UF1("S" & 10).BackColor = Scol
     End If
    End If
    Case 2
    If (UF1.DiagU <> "") Then
     If DiagonalJudg(UF1.DiagU, UF1.TaikText) Then
      UF1("U" & 9).BackColor = Scol
      UF1("U" & 10).BackColor = Scol
     End If
    End If
    Case 3
    If (UF1.DiagC <> "") Then
     If DiagonalJudg(UF1.DiagC, UF1.TaikText) Then
      UF1("C" & 9).BackColor = Scol
      UF1("C" & 10).BackColor = Scol
     End If
    End If
    Case 4
    If (UF1.DiagD <> "") Then
     If DiagonalJudg(UF1.DiagD, UF1.TaikText) Then
      UF1("D" & 9).BackColor = Scol
      UF1("D" & 10).BackColor = Scol
     End If
    End If
    Case 5
    If (UF1.DiagE <> "") Then
     If DiagonalJudg(UF1.DiagE, UF1.TaikText) Then
      UF1("E" & 9).BackColor = Scol
      UF1("E" & 10).BackColor = Scol
     End If
    End If
    Case 6
    If (UF1.DiagF <> "") Then
     If DiagonalJudg(UF1.DiagF, UF1.TaikText) Then
      UF1("F" & 9).BackColor = Scol
      UF1("F" & 10).BackColor = Scol
     End If
    End If
    End Select
  Next a
End If

SkipL:

'ﾗﾃﾗﾙﾎﾞｰ_2021.07.13
If IsNumeric(UF1.LateralText) = False Then Exit Sub
For a = j To k
  Select Case a
  Case 1: If IsNumeric(UF1.LateralS) Then If LateralJudg(UF1.LateralS, UF1.LateralText) Then UF1.LateralS.BackColor = Scol
  Case 2: If IsNumeric(UF1.LateralU) Then If LateralJudg(UF1.LateralU, UF1.LateralText) Then UF1.LateralU.BackColor = Scol
  Case 3: If IsNumeric(UF1.LateralC) Then If LateralJudg(UF1.LateralC, UF1.LateralText) Then UF1.LateralC.BackColor = Scol
  Case 4: If IsNumeric(UF1.LateralD) Then If LateralJudg(UF1.LateralD, UF1.LateralText) Then UF1.LateralD.BackColor = Scol
  Case 5: If IsNumeric(UF1.LateralE) Then If LateralJudg(UF1.LateralE, UF1.LateralText) Then UF1.LateralE.BackColor = Scol
  Case 6: If IsNumeric(UF1.LateralF) Then If LateralJudg(UF1.LateralF, UF1.LateralText) Then UF1.LateralF.BackColor = Scol
  End Select
Next a
'2023.6.27_BH
For a = j To k
  Select Case a
  Case 1: If IsNumeric(UF1.SBH) Then If BhJudg(UF1.SBH, UF1.Zai, UF1.Tyu) = "OK" Then UF1.SBH.BackColor = RGB(255, 255, 255)
  Case 2: If IsNumeric(UF1.UBH) Then If BhJudg(UF1.UBH, UF1.Zai, UF1.Tyu) = "OK" Then UF1.UBH.BackColor = RGB(255, 255, 255)
  Case 3: If IsNumeric(UF1.CBH) Then If BhJudg(UF1.CBH, UF1.Zai, UF1.Tyu) = "OK" Then UF1.CBH.BackColor = RGB(255, 255, 255)
  Case 4: If IsNumeric(UF1.DBH) Then If BhJudg(UF1.DBH, UF1.Zai, UF1.Tyu) = "OK" Then UF1.DBH.BackColor = RGB(255, 255, 255)
  Case 5: If IsNumeric(UF1.EBH) Then If BhJudg(UF1.EBH, UF1.Zai, UF1.Tyu) = "OK" Then UF1.EBH.BackColor = RGB(255, 255, 255)
  End Select
Next a
For a = j To k
  Select Case a
  Case 1: If IsNumeric(UF1.SBH) Then If BhJudg(UF1.SBH, UF1.Zai, UF1.Tyu) = "NG" Then UF1.SBH.BackColor = Ncol
  Case 2: If IsNumeric(UF1.UBH) Then If BhJudg(UF1.UBH, UF1.Zai, UF1.Tyu) = "NG" Then UF1.UBH.BackColor = Ncol
  Case 3: If IsNumeric(UF1.CBH) Then If BhJudg(UF1.CBH, UF1.Zai, UF1.Tyu) = "NG" Then UF1.CBH.BackColor = Ncol
  Case 4: If IsNumeric(UF1.DBH) Then If BhJudg(UF1.DBH, UF1.Zai, UF1.Tyu) = "NG" Then UF1.DBH.BackColor = Ncol
  Case 5: If IsNumeric(UF1.EBH) Then If BhJudg(UF1.EBH, UF1.Zai, UF1.Tyu) = "NG" Then UF1.EBH.BackColor = Ncol
  End Select
Next a
'直角度_R115の指定公差*1.0
For a = j To k
  Select Case a
  Case 1: If IsNumeric(UF1.DiagSKou) Then If LateralJudg(UF1.DiagSKou, "1.0") Then UF1.DiagSKou.BackColor = Scol
  Case 2: If IsNumeric(UF1.DiagUKou) Then If LateralJudg(UF1.DiagUKou, "1.0") Then UF1.DiagUKou.BackColor = Scol
  Case 3: If IsNumeric(UF1.DiagCKou) Then If LateralJudg(UF1.DiagCKou, "1.0") Then UF1.DiagCKou.BackColor = Scol
  Case 4: If IsNumeric(UF1.DiagDKou) Then If LateralJudg(UF1.DiagDKou, "1.0") Then UF1.DiagDKou.BackColor = Scol
  Case 5: If IsNumeric(UF1.DiagEKou) Then If LateralJudg(UF1.DiagEKou, "1.0") Then UF1.DiagEKou.BackColor = Scol
  Case 6: If IsNumeric(UF1.DiagFKou) Then If LateralJudg(UF1.DiagFKou, "1.0") Then UF1.DiagFKou.BackColor = Scol
  End Select
Next a

If UF1.LateralS = "" Or UF1.LateralS = "-" Then UF1.LateralS.BackColor = RGB(255, 255, 255)
If UF1.LateralU = "" Or UF1.LateralU = "-" Then UF1.LateralU.BackColor = RGB(255, 255, 255)
If UF1.LateralC = "" Or UF1.LateralC = "-" Then UF1.LateralC.BackColor = RGB(255, 255, 255)
If UF1.LateralD = "" Or UF1.LateralD = "-" Then UF1.LateralD.BackColor = RGB(255, 255, 255)
If UF1.LateralE = "" Or UF1.LateralE = "-" Then UF1.LateralE.BackColor = RGB(255, 255, 255)
If UF1.LateralF = "" Or UF1.LateralF = "-" Then UF1.LateralF.BackColor = RGB(255, 255, 255)

If UF1.DiagSKou = "" Or UF1.DiagSKou = "-" Then UF1.DiagSKou.BackColor = RGB(255, 255, 255)
If UF1.DiagUKou = "" Or UF1.DiagUKou = "-" Then UF1.DiagUKou.BackColor = RGB(255, 255, 255)
If UF1.DiagCKou = "" Or UF1.DiagCKou = "-" Then UF1.DiagCKou.BackColor = RGB(255, 255, 255)
If UF1.DiagDKou = "" Or UF1.DiagDKou = "-" Then UF1.DiagDKou.BackColor = RGB(255, 255, 255)
If UF1.DiagEKou = "" Or UF1.DiagEKou = "-" Then UF1.DiagEKou.BackColor = RGB(255, 255, 255)
If UF1.DiagFKou = "" Or UF1.DiagFKou = "-" Then UF1.DiagFKou.BackColor = RGB(255, 255, 255)

End Sub

'------------------------------ 公差等級.4級(厳しい) 直角度
Public Function Diagonal(a As Single, b As Single) As Single '(巾,丈)公差
    Diagonal = LookupDiag("4", a, b)
End Function

'------------------------------ 公差等級.3級(緩い) 直角度
Public Function ReDiagonal(a As Single, b As Single) As Single '(巾,丈)公差
    ReDiagonal = LookupDiag("3", a, b)
End Function

'------------------------------ 公差等級.4級(厳しい) ラテラルボー
Public Function Lateralbow_J(a As Single, b As Single, C As Single) As Single '(巾,丈,厚)=公差
    Lateralbow_J = LookupLat("4", a, b, C)
End Function

'------------------------------ 公差等級.3級(緩い) ラテラルボー
Public Function ReLateralbow(a As Single, b As Single, C As Single) As Single '(巾,丈,厚)=公差
    ReLateralbow = LookupLat("3", a, b, C)
End Function

'--- テーブル修正後に読み直す
Public Sub ReloadTolerance()
    m_Loaded = False
    m_LoadFailed = False
    m_Diag = Empty
    m_Lat = Empty
    If EnsureLoaded() Then
        MsgBox "公差テーブルを読み込み直しました", vbInformation
    End If
End Sub

'========================================
' 検索
'   範囲が重なる行が複数あれば、管理番号の小さい行を採用
'   （ラテラルボーの巾900：前区分を優先＝旧コードと同じ）
'========================================
Private Function LookupDiag(ByVal grade As String, ByVal a As Single, ByVal b As Single) As Single
    If Not EnsureLoaded() Then Exit Function
    If IsEmpty(m_Diag) Then Exit Function

    'Single の誤差対策（1500.1 → 1500.0999… で区分から漏れる問題）
    Dim w As Double: w = Round(CDbl(a), 2)
    Dim l As Double: l = Round(CDbl(b), 2)
    Dim r As Long, bestNo As Double, hit As Boolean

    For r = LBound(m_Diag, 1) To UBound(m_Diag, 1)
        If CStr(m_Diag(r, m_DiagIdx(D_GRADE)) & "") = grade Then
            If InRange(w, m_Diag(r, m_DiagIdx(D_W_LO)), m_Diag(r, m_DiagIdx(D_W_HI))) Then
                If InRange(l, m_Diag(r, m_DiagIdx(D_L_LO)), m_Diag(r, m_DiagIdx(D_L_HI))) Then
                    If Not hit Or m_Diag(r, m_DiagIdx(D_NO)) < bestNo Then
                        hit = True
                        bestNo = m_Diag(r, m_DiagIdx(D_NO))
                        LookupDiag = CSng(m_Diag(r, m_DiagIdx(D_TOL)))
                    End If
                End If
            End If
        End If
    Next r
    '該当なし → 0（従来どおり）
End Function

Private Function LookupLat(ByVal grade As String, ByVal a As Single, ByVal b As Single, ByVal C As Single) As Single
    If Not EnsureLoaded() Then Exit Function
    If IsEmpty(m_Lat) Then Exit Function

    Dim w As Double: w = Round(CDbl(a), 2)
    Dim l As Double: l = Round(CDbl(b), 2)
    Dim t As Double: t = Round(CDbl(C), 2)
    Dim r As Long, bestNo As Double, hit As Boolean

    For r = LBound(m_Lat, 1) To UBound(m_Lat, 1)
        If CStr(m_Lat(r, m_LatIdx(L_GRADE)) & "") = grade Then
            If InRange(t, m_Lat(r, m_LatIdx(L_T_LO)), m_Lat(r, m_LatIdx(L_T_HI))) Then
                If InRange(w, m_Lat(r, m_LatIdx(L_W_LO)), m_Lat(r, m_LatIdx(L_W_HI))) Then
                    If InRange(l, m_Lat(r, m_LatIdx(L_L_LO)), m_Lat(r, m_LatIdx(L_L_HI))) Then
                        If Not hit Or m_Lat(r, m_LatIdx(L_NO)) < bestNo Then
                            hit = True
                            bestNo = m_Lat(r, m_LatIdx(L_NO))
                            LookupLat = CSng(m_Lat(r, m_LatIdx(L_TOL)))
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Function

'--- 下限以上・上限以下（Null・空欄は不一致扱い）
Private Function InRange(ByVal v As Double, ByVal lo As Variant, ByVal hi As Variant) As Boolean
    If IsNull(lo) Or IsNull(hi) Then Exit Function
    If Len(lo & "") = 0 Or Len(hi & "") = 0 Then Exit Function
    InRange = (v >= CDbl(lo) And v <= CDbl(hi))
End Function

'========================================
' 読込（初回のみ）
'========================================
Private Function EnsureLoaded() As Boolean
    If m_Loaded Then EnsureLoaded = True: Exit Function
    If m_LoadFailed Then Exit Function      '失敗済みなら再試行しない（入力のたびにMsgBoxが出るのを防ぐ）

    Dim FilePath As String
    Dim step As String
    Dim errNo As Long, errDesc As String

    On Error GoTo ErrHandler

    FilePath = QualityPath & "品質基準.accdb"

    step = TBL_DIAG & " 読込"
    Call LoadTable(FilePath, TBL_DIAG, _
        Array("管理番号", "等級", "巾_下限", "巾_上限", "丈_下限", "丈_上限", "公差"), _
        m_Diag, m_DiagIdx)

    step = TBL_LAT & " 読込"
    Call LoadTable(FilePath, TBL_LAT, _
        Array("管理番号", "等級", "厚_下限", "厚_上限", "巾_下限", "巾_上限", "丈_下限", "丈_上限", "公差"), _
        m_Lat, m_LatIdx)

    m_Loaded = True
    EnsureLoaded = True
    Exit Function

ErrHandler:
    errNo = Err.Number: errDesc = Err.Description   '★先に退避
    m_LoadFailed = True
    MsgBox "公差テーブルを読み込めませんでした" & vbCrLf & _
           "段階: " & step & vbCrLf & _
           "内容: " & errNo & " " & errDesc & vbCrLf & vbCrLf & _
           "公差は表示されません。原因を解消後 ReloadTolerance を実行してください。", vbCritical
End Function

'--- 1テーブル分：列名から列番号を求め、レコードを取得
Private Sub LoadTable(ByVal FilePath As String, ByVal TableName As String, ByVal colNames As Variant, _
                      ByRef data As Variant, ByRef idx() As Long)
    Dim fieldsfilter As String:  fieldsfilter = ""
    Dim SQLfilter As String:     SQLfilter = ""
    Dim flds As Variant
    Dim hdr As Long, i As Long, k As Long

    flds = GetFieldsArr(FilePath, TableName, fieldsfilter)
    If IsEmpty(flds) Then Err.Raise vbObjectError + 1, , "「" & TableName & "」のフィールドを取得できません"

    hdr = LBound(flds, 1)
    ReDim idx(LBound(colNames) To UBound(colNames))
    For k = LBound(colNames) To UBound(colNames)
        idx(k) = -1
        For i = LBound(flds, 2) To UBound(flds, 2)
            If flds(hdr, i) = colNames(k) Then idx(k) = i: Exit For
        Next i
        If idx(k) = -1 Then
            Err.Raise vbObjectError + 2, , "「" & TableName & "」に「" & colNames(k) & "」フィールドがありません"
        End If
    Next k

    data = GetRecordsArr(FilePath, TableName, SQLfilter)
End Sub


'Public Function Diagonal(a As Single, b As Single) As Single '(巾,丈)公差
''------------------------------公差等級.4級(厳しい)製品規格：書式番号Z501
''3701~いくつまでという記述なし_7000mmまで反映させる**2022.02.16
'Select Case a
'Case 1# To 300#
'  Select Case b
'  Case 1# To 3700#
'    Diagonal = 0.8
'  Case 3701# To 7000#
'    Diagonal = 1#
'  End Select
'Case 300.1 To 600#
'  Select Case b
'  Case 1# To 3700#
'    Diagonal = 1#
'  Case 3701# To 7000#
'    Diagonal = 1.5
'  End Select
'Case 600.1 To 900#
'  Select Case b
'  Case 1# To 3700#
'    Diagonal = 1.5
'  Case 3701# To 7000#
'    Diagonal = 2#
'  End Select
'Case 900.1 To 1200#
'  Select Case b
'  Case 1# To 3700#
'    Diagonal = 2#
'  Case 3701# To 7000#
'    Diagonal = 2.5
'  End Select
'Case 1200.1 To 1500#
'  Select Case b
'  Case 1# To 3700#
'    Diagonal = 2.5
'  Case 3701# To 7000#
'    Diagonal = 3#
'  End Select
'Case 1500.1 To 1800#
'  Select Case b
'  Case 1# To 3700#
'    Diagonal = 2.5
'  Case 3701# To 7000#
'    Diagonal = 3#
'  End Select
'Case 1800.1 To 2100#
'  Select Case b
'  Case 1# To 3700#
'    Diagonal = 2.5
'  Case 3701# To 7000#
'    Diagonal = 3#
'  End Select
'Case 2100.1 To 2400#
'  Select Case b
'  Case 1# To 3700#
'    Diagonal = 2.5
'  Case 3701# To 7000#
'    Diagonal = 3#
'  End Select
'End Select
'End Function
'
'Public Function ReDiagonal(a As Single, b As Single) As Single  '(巾,丈)公差
''------------------------------公差等級.3級(緩い)製品規格：書式番号Z501
''3701~いくつまでという記述なし_7000mmまで反映させる**2022.02.16
'Select Case a
'Case 1# To 300#
'  Select Case b
'  Case 1# To 3700#
'   ReDiagonal = 1.5
'  Case 3701# To 7000#
'   ReDiagonal = 2#
'  End Select
'Case 300.1 To 600#
'  Select Case b
'  Case 1# To 3700#
'   ReDiagonal = 2#
'  Case 3701# To 7000#
'   ReDiagonal = 3#
'  End Select
'Case 600.1 To 900#
'  Select Case b
'  Case 1# To 3700#
'   ReDiagonal = 3#
'  Case 3701# To 7000#
'   ReDiagonal = 4#
'  End Select
'Case 900.1 To 1200#
'  Select Case b
'  Case 1# To 3700#
'   ReDiagonal = 4#
'  Case 3701# To 7000#
'   ReDiagonal = 5#
'  End Select
'Case 1200.1 To 1500#
'  Select Case b
'  Case 1# To 3700#
'   ReDiagonal = 5#
'  Case 3701# To 7000#
'   ReDiagonal = 6#
'  End Select
'Case 1500.1 To 1800#
'  Select Case b
'  Case 1# To 3700#
'   ReDiagonal = 5#
'  Case 3701# To 7000#
'   ReDiagonal = 6#
'  End Select
'Case 1800.1 To 2100#
'  Select Case b
'  Case 1# To 3700#
'   ReDiagonal = 5#
'  Case 3701# To 7000#
'   ReDiagonal = 6#
'  End Select
'Case 2100.1 To 2400#
'  Select Case b
'  Case 1# To 3700#
'   ReDiagonal = 5#
'  Case 3701# To 7000#
'   ReDiagonal = 6#
'  End Select
'End Select
'End Function


'Public Function Lateralbow_J(a As Single, b As Single, C As Single) As Single   '(巾,丈,厚)=公差
''------------------------------公差等級.4級(厳しい)製品規格：書式番号Z501_2021.07.26公差更新
''表は6000まで*特記で6000超えについての記述あり_2022.02.16
'Select Case C '厚
'Case 0.1 To 3.2
'  Select Case a '巾
'  Case 1# To 100#
'    Select Case b '丈
'    Case 1 To 1500#
'      Lateralbow_J = 1#
'    Case 1500.1 To 2300#
'      Lateralbow_J = 1.5
'    Case 2300.1 To 3000#
'      Lateralbow_J = 2#
'    Case 3000.1 To 3800#
'      Lateralbow_J = 3.5
'    Case 3800.1 To 4500#
'      Lateralbow_J = 4#
'    Case 4500.1 To 5500#
'      Lateralbow_J = 5#
'    Case 5500.1 To 6000#
'      Lateralbow_J = 6#
'    Case 6000.1 To 7000#
'      Lateralbow_J = 6#
'    End Select
'  Case 100.1 To 250#
'    Select Case b '丈
'    Case 1 To 1500#
'      Lateralbow_J = 1#
'    Case 1500.1 To 2300#
'      Lateralbow_J = 1.5
'    Case 2300.1 To 3000#
'      Lateralbow_J = 2#
'    Case 3000.1 To 3800#
'      Lateralbow_J = 3.5
'    Case 3800.1 To 4500#
'      Lateralbow_J = 4#
'    Case 4500.1 To 5500#
'      Lateralbow_J = 5#
'    Case 5500.1 To 6000#
'      Lateralbow_J = 6#
'    Case 6000.1 To 7000#
'      Lateralbow_J = 6#
'    End Select
'  Case 250.1 To 900#
'    Select Case b '丈
'    Case 1 To 1500#
'      Lateralbow_J = 1#
'    Case 1500.1 To 2300#
'      Lateralbow_J = 1.5
'    Case 2300.1 To 3000#
'      Lateralbow_J = 2#
'    Case 3000.1 To 3800#
'      Lateralbow_J = 3.5
'    Case 3800.1 To 4500#
'      Lateralbow_J = 4#
'    Case 4500.1 To 5500#
'      Lateralbow_J = 5#
'    Case 5500.1 To 6000#
'      Lateralbow_J = 6#
'    Case 6000.1 To 7000#
'      Lateralbow_J = 6#
'    End Select
'  Case 900# To 2000#
'    Select Case b '丈
'    Case 1 To 1500#
'      Lateralbow_J = 1#
'    Case 1500.1 To 2300#
'      Lateralbow_J = 1.5
'    Case 2300.1 To 3000#
'      Lateralbow_J = 2#
'    Case 3000.1 To 3800#
'      Lateralbow_J = 3.5
'    Case 3800.1 To 4500#
'      Lateralbow_J = 4#
'    Case 4500.1 To 5500#
'      Lateralbow_J = 5#
'    Case 5500.1 To 6000#
'      Lateralbow_J = 6#
'    Case 6000.1 To 7000#
'      Lateralbow_J = 6#
'    End Select
'  End Select
'Case 3.21 To 6.5
'  Select Case a '巾
'  Case 101 To 900
'    Select Case b '丈
'    Case 1 To 1500#
'      Lateralbow_J = 1#
'    Case 1500.1 To 2300#
'      Lateralbow_J = 1.5
'    Case 2300.1 To 3000#
'      Lateralbow_J = 2#
'    Case 3000.1 To 3800#
'      Lateralbow_J = 3.5
'    Case 3800.1 To 4500#
'      Lateralbow_J = 4#
'    Case 4500.1 To 5500#
'      Lateralbow_J = 5#
'    Case 5500.1 To 6000#
'      Lateralbow_J = 6#
'    Case 6000.1 To 7000#
'      Lateralbow_J = 6#
'    End Select
'  Case 901# To 2000#
'    Select Case b '丈
'    Case 1 To 1500#
'      Lateralbow_J = 1#
'    Case 1500.1 To 2300#
'      Lateralbow_J = 1.5
'    Case 2300.1 To 3000#
'      Lateralbow_J = 2#
'    Case 3000.1 To 3800#
'      Lateralbow_J = 3.5
'    Case 3800.1 To 4500#
'      Lateralbow_J = 4#
'    Case 4500.1 To 5500#
'      Lateralbow_J = 5#
'    Case 5500.1 To 6000#
'      Lateralbow_J = 6#
'    Case 6000.1 To 7000#
'      Lateralbow_J = 6#
'    End Select
'  End Select
'Case 6.51 To 203#
'  Select Case a '巾
'  Case 1 To 250#
'    Select Case b '丈
'    Case 1 To 1500#
'      Lateralbow_J = 1#
'    Case 1500.1 To 2300#
'      Lateralbow_J = 1.5
'    Case 2300.1 To 3000#
'      Lateralbow_J = 2#
'    Case 3000.1 To 3800#
'      Lateralbow_J = 3.5
'    Case 3800.1 To 4500#
'      Lateralbow_J = 4#
'    Case 4500.1 To 5500#
'      Lateralbow_J = 5#
'    Case 5500.1 To 6000#
'      Lateralbow_J = 6#
'    Case 6000.1 To 7000#
'      Lateralbow_J = 6#
'    End Select
'  Case 251# To 450#
'    Select Case b '丈
'    Case 1 To 1500#
'      Lateralbow_J = 1#
'    Case 1500.1 To 2300#
'      Lateralbow_J = 1.5
'    Case 2300.1 To 3000#
'      Lateralbow_J = 2#
'    Case 3000.1 To 3800#
'      Lateralbow_J = 3.5
'    Case 3800.1 To 4500#
'      Lateralbow_J = 4#
'    Case 4500.1 To 5500#
'      Lateralbow_J = 5#
'    Case 5500.1 To 6000#
'      Lateralbow_J = 6#
'    Case 6000.1 To 7000#
'      Lateralbow_J = 6#
'    End Select
'  Case 451# To 2000#
'    Select Case b '丈
'    Case 1 To 1500#
'      Lateralbow_J = 1#
'    Case 1500.1 To 2300#
'      Lateralbow_J = 1.5
'    Case 2300.1 To 3000#
'      Lateralbow_J = 2#
'    Case 3000.1 To 3800#
'      Lateralbow_J = 3.5
'    Case 3800.1 To 4500#
'      Lateralbow_J = 4#
'    Case 4500.1 To 5500#
'      Lateralbow_J = 5#
'    Case 5500.1 To 6000#
'      Lateralbow_J = 6#
'    Case 6000.1 To 7000#
'      Lateralbow_J = 6#
'    End Select
'  End Select
'End Select
'End Function
'
'Public Function ReLateralbow(a As Single, b As Single, C As Single) As Single '(巾,丈,厚)=公差
''------------------------------公差等級.3級(緩い)製品規格：書式番号Z501_2021.07.26公差更新
''6000超えについては6000の許容差を適用_2022.02.16
'Select Case C '厚
'Case 0.1 To 3.2
'  Select Case a '巾
'  Case 1# To 100#
'    Select Case b '丈
'    Case 1 To 1500#
'      ReLateralbow = 6#
'    Case 1500.1 To 2300#
'      ReLateralbow = 14#
'    Case 2300.1 To 3000#
'      ReLateralbow = 25#
'    Case 3000.1 To 3800#
'      ReLateralbow = 39#
'    Case 3800.1 To 4500#
'      ReLateralbow = 57#
'    Case 4500.1 To 5500#
'      ReLateralbow = 76#
'    Case 5500.1 To 6000#
'      ReLateralbow = 102#
'    Case 6000.1 To 7000#
'      ReLateralbow = 102#
'    End Select
'  Case 100.1 To 250#
'    Select Case b '丈
'    Case 1 To 1500#
'      ReLateralbow = 3#
'    Case 1500.1 To 2300#
'      ReLateralbow = 6#
'    Case 2300.1 To 3000#
'      ReLateralbow = 10#
'    Case 3000.1 To 3800#
'      ReLateralbow = 14#
'    Case 3800.1 To 4500#
'      ReLateralbow = 22#
'    Case 4500.1 To 5500#
'      ReLateralbow = 29#
'    Case 5500.1 To 6000#
'      ReLateralbow = 38#
'    Case 6000.1 To 7000#
'      ReLateralbow = 38#
'    End Select
'  Case 250.1 To 900#
'    Select Case b '丈
'    Case 1 To 1500#
'      ReLateralbow = 2#
'    Case 1500.1 To 2300#
'      ReLateralbow = 3#
'    Case 2300.1 To 3000#
'      ReLateralbow = 5#
'    Case 3000.1 To 3800#
'      ReLateralbow = 6#
'    Case 3800.1 To 4500#
'      ReLateralbow = 10#
'    Case 4500.1 To 5500#
'      ReLateralbow = 13#
'    Case 5500.1 To 6000#
'      ReLateralbow = 19#
'    Case 6000.1 To 7000#
'      ReLateralbow = 19#
'    End Select
'  Case 900# To 2000#
'    Select Case b '丈
'    Case 1 To 1500#
'      ReLateralbow = 1#
'    Case 1500.1 To 2300#
'      ReLateralbow = 2#
'    Case 2300.1 To 3000#
'      ReLateralbow = 3#
'    Case 3000.1 To 3800#
'      ReLateralbow = 5#
'    Case 3800.1 To 4500#
'      ReLateralbow = 6#
'    Case 4500.1 To 5500#
'      ReLateralbow = 10#
'    Case 5500.1 To 6000#
'      ReLateralbow = 13#
'    Case 6000.1 To 7000#
'      ReLateralbow = 13#
'    End Select
'  End Select
'Case 3.21 To 6.5
'  Select Case a '巾
'  Case 101 To 900#
'    Select Case b '丈
'    Case 1 To 1500#
'      ReLateralbow = 2#
'    Case 1500.1 To 2300#
'      ReLateralbow = 3#
'    Case 2300.1 To 3000#
'      ReLateralbow = 6#
'    Case 3000.1 To 3800#
'      ReLateralbow = 9#
'    Case 3800.1 To 4500#
'      ReLateralbow = 14#
'    Case 4500.1 To 5500#
'      ReLateralbow = 19#
'    Case 5500.1 To 6000#
'      ReLateralbow = 25#
'    Case 6000.1 To 7000#
'      ReLateralbow = 25#
'    End Select
'  Case 901# To 2000#
'    Select Case b '丈
'    Case 1 To 1500#
'      ReLateralbow = 1#
'    Case 1500.1 To 2300#
'      ReLateralbow = 2#
'    Case 2300.1 To 3000#
'      ReLateralbow = 3#
'    Case 3000.1 To 3800#
'      ReLateralbow = 5#
'    Case 3800.1 To 4500#
'      ReLateralbow = 6#
'    Case 4500.1 To 5500#
'      ReLateralbow = 10#
'    Case 5500.1 To 6000#
'      ReLateralbow = 13#
'    Case 6000.1 To 7000#
'      ReLateralbow = 13#
'    End Select
'  End Select
'Case 6.51 To 203#
'  Select Case a '巾
'  Case 1 To 250#
'    Select Case b '丈
'    Case 1 To 1500#
'      ReLateralbow = 6#
'    Case 1500.1 To 2300#
'      ReLateralbow = 14#
'    Case 2300.1 To 3000#
'      ReLateralbow = 25#
'    Case 3000.1 To 3800#
'      ReLateralbow = 39#
'    Case 3800.1 To 4500#
'      ReLateralbow = 57#
'    Case 4500.1 To 5500#
'      ReLateralbow = 76#
'    Case 5500.1 To 6000#
'      ReLateralbow = 102#
'    Case 6000.1 To 7000#
'      ReLateralbow = 102#
'    End Select
'  Case 251# To 450#
'    Select Case b '丈
'    Case 1 To 1500#
'      ReLateralbow = 2#
'    Case 1500.1 To 2300#
'      ReLateralbow = 3#
'    Case 2300.1 To 3000#
'      ReLateralbow = 6#
'    Case 3000.1 To 3800#
'      ReLateralbow = 10#
'    Case 3800.1 To 4500#
'      ReLateralbow = 15#
'    Case 4500.1 To 5500#
'      ReLateralbow = 20#
'    Case 5500.1 To 6000#
'      ReLateralbow = 25#
'    Case 6000.1 To 7000#
'      ReLateralbow = 25#
'    End Select
'  Case 451# To 2000#
'    Select Case b '丈
'    Case 1 To 1500#
'      ReLateralbow = 1#
'    Case 1500.1 To 2300#
'      ReLateralbow = 2#
'    Case 2300.1 To 3000#
'      ReLateralbow = 3#
'    Case 3000.1 To 3800#
'      ReLateralbow = 5#
'    Case 3800.1 To 4500#
'      ReLateralbow = 8#
'    Case 4500.1 To 5500#
'      ReLateralbow = 11#
'    Case 5500.1 To 6000#
'      ReLateralbow = 14#
'    Case 6000.1 To 7000#
'      ReLateralbow = 14#
'    End Select
'  End Select
'End Select
'End Function

'--- 動作確認用（確認後は削除可）
Public Sub Test_Tolerance()
    Dim msg As String
    On Error GoTo ErrHandler

    msg = "QualityPath = [" & QualityPath & "]" & vbCrLf & vbCrLf
    msg = msg & "Diagonal(500, 2000)          = " & Diagonal(500, 2000) & "　（期待値 1）" & vbCrLf
    msg = msg & "ReDiagonal(1000, 5000)       = " & ReDiagonal(1000, 5000) & "　（期待値 5）" & vbCrLf
    msg = msg & "Lateralbow_J(500, 2000, 3.2) = " & Lateralbow_J(500, 2000, 3.2) & "　（期待値 1.5）" & vbCrLf
    msg = msg & "ReLateralbow(300, 4000, 8)   = " & ReLateralbow(300, 4000, 8) & "　（期待値 15）"

    MsgBox msg, vbInformation, "公差テスト"
    Exit Sub

ErrHandler:
    MsgBox "エラー: " & Err.Number & " " & Err.Description, vbCritical, "公差テスト"
End Sub


Public Sub LateralbowJudgColor(TB As MSForms.TextBox) 'ラテラルボー判定　色　音　画像
If TB.Text = "" Then Exit Sub
If TB.Text = "-" Then Exit Sub

If IsNumeric(TB.Text) = False Then Exit Sub
Dim Scol As Long: Scol = 13434828 '合格
Dim Ncol As Long: Ncol = 8421631
Dim a As Single: a = TB.Text
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If IsNumeric(UF1.LateralText) Then '公差に数字がある
  Dim kou As String: kou = UF1.LateralText
Else
  MsgBox "ﾗﾃﾗﾙﾎﾞｰ公差がありません。(規定寸法範囲外)", vbInformation
  UF1.MPEX.Tag = "" 'フラグリセット
  TB.Tag = "範囲外"
  Exit Sub
End If

Dim fname As String: fname = "Measured.log"
Dim path As String: path = HistoryPath
Dim Lpath As String
Select Case Lname
Case "L-1":    Lpath = path & "L-1\L-1_Log"
Case "LVC":    Lpath = path & "LVC\LVC_Log"
Case "HVC":    Lpath = path & "HVC\HVC_Log"
Case "ﾊﾞﾗﾝｻｰ": Lpath = path & "バランサー\バランサー_Log"
Case "ﾄｯﾄ":    Lpath = path & "トット\トット_Log"
Case "Test":   Lpath = path & "Test\Test_Log"
Case Else:     Lpath = path & "丸徳\丸徳_Log"
End Select

Dim NowSizeH As Long:  NowSizeH = UF1.MPEX.Height '2024.5.30_追加
Dim NowSizeW As Long:  NowSizeW = UF1.MPEX.Width
Dim NowSizeMP As Long: NowSizeMP = UF1.MP1.Width
Dim WorkerName As String: WorkerName = GetWorkerName()

Call Fm_Clear("All")

If UF1.MPEX.Tag <> "Click" Then '2023.8.18_エラー対策
    On Error Resume Next
  UF1.MPEX.Visible = False '2023.5.17_追加
  UF1.MP1.Visible = False
Else
  UF1.MPEX.Height = 1
  UF1.MPEX.Width = 1
End If

'2026.05.14_WriteLog引数追加
  If LateralJudg(a, kou) Then
    Call WriteLog("*UFKen|" & TB.Text & "|(" & TB.Name & ")" & "|○|" & Lname & "|" & WorkerName, "", Lpath, fname)
    Call WriteLog("*UFKen|" & TB.Text & "|(" & TB.Name & ")" & "|○|" & Lname & "|" & WorkerName, "")
    TB.BackColor = Scol
    Call 音楽合格
    Call Judg_ImageSize("Y")
    
  ElseIf LateralJudg(a, kou) = False Then
    Call WriteLog("*UFKen|" & TB.Text & "|(" & TB.Name & ")" & "|×|" & Lname & "|" & WorkerName, "", Lpath, fname)
    Call WriteLog("*UFKen|" & TB.Text & "|(" & TB.Name & ")" & "|×|" & Lname & "|" & WorkerName, "")
    TB.BackColor = Ncol
    Call 音楽不合格
    Call Judg_ImageSize("N")
    
  End If
  
Sleep 2000

UF1.IMNG.Visible = False
UF1.IMOK.Visible = False
'2022.01.28_追加
UF1.NoL.Visible = True
UF1.PackingForm.Visible = True
UF1.PackingComment.Visible = True
  
UF1.MP1.Visible = True
UF1.MPEX.Visible = True '2023.5.17_追加
Call Fm_Clear("Click")

UF1.MP1.Visible = True
UF1.MP1.Width = NowSizeMP '2021.2.15_薄板の時非表示にできない....ｻｲｽﾞ修正で逃
UF1.MPEX.Height = NowSizeH
UF1.MPEX.Width = NowSizeW

UF1.Repaint
If UF1.Test Then
  UF1.WorkFrame.Visible = True
  UF1.NoFrame.Visible = True
End If

'2023.8.21_追加
UF1.VC_A.Visible = True
UF1.VC_B.Visible = True
UF1.VC_L_A.Visible = True
UF1.VC_L_B.Visible = True

UF1.MPEX.Tag = "" 'フラグリセット

End Sub

Public Sub RightAngleJudgColor(TB As MSForms.TextBox) '直角度判定　色　音　画像*****2023.11.27
If TB.Text = "" Then Exit Sub
If IsNumeric(TB.Text) = False Then Exit Sub
Dim Scol As Long: Scol = 13434828 '合格
Dim Ncol As Long: Ncol = 8421631

Dim a As Single: a = TB.Text
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
'Q-S-QA-DIM-154_対角線差及び直角度測定基準
'指定公差　1.0
Dim kou As String: kou = 1#

Dim fname As String:  fname = "Measured.log"
Dim path As String:   path = HistoryPath
Dim Lpath As String
Select Case Lname
Case "L-1":    Lpath = path & "L-1\L-1_Log"
Case "LVC":    Lpath = path & "LVC\LVC_Log"
Case "HVC":    Lpath = path & "HVC\HVC_Log"
Case "ﾊﾞﾗﾝｻｰ": Lpath = path & "バランサー\バランサー_Log"
Case "ﾄｯﾄ":    Lpath = path & "トット\トット_Log"
Case "Test":   Lpath = path & "Test\Test_Log"
Case Else:     Lpath = path & "丸徳\丸徳_Log"
End Select

Dim NowSizeH As Long:  NowSizeH = UF1.MPEX.Height '2024.5.30_追加
Dim NowSizeW As Long:  NowSizeW = UF1.MPEX.Width
Dim NowSizeMP As Long: NowSizeMP = UF1.MP1.Width
Dim WorkerName As String: WorkerName = GetWorkerName()

Call Fm_Clear("All")

If UF1.MPEX.Tag <> "Click" Then '2023.8.18_エラー対策
    On Error Resume Next
  UF1.MPEX.Visible = False '2023.5.17_追加
  UF1.MP1.Visible = False
Else
  UF1.MPEX.Height = 1
  UF1.MPEX.Width = 1
End If

'判定関数はラテラルボーと同じ
  If LateralJudg(a, kou) Then
    Call WriteLog("*UFKen|" & TB.Text & "|(" & TB.Name & ")" & "|○|" & Lname & "|" & WorkerName, "", Lpath, fname)
    Call WriteLog("*UFKen|" & TB.Text & "|(" & TB.Name & ")" & "|○|" & Lname & "|" & WorkerName, "")
    
    TB.BackColor = Scol
    Call 音楽合格
    Call Judg_ImageSize("Y")
    
  ElseIf LateralJudg(a, kou) = False Then
    Call WriteLog("*UFKen|" & TB.Text & "|(" & TB.Name & ")" & "|×|" & Lname & "|" & WorkerName, "", Lpath, fname)
    Call WriteLog("*UFKen|" & TB.Text & "|(" & TB.Name & ")" & "|×|" & Lname & "|" & WorkerName, "")
    
    TB.BackColor = Ncol
    Call 音楽不合格
    Call Judg_ImageSize("N")
    
  End If
  
Sleep 2000
UF1.IMNG.Visible = False
UF1.IMOK.Visible = False
'2022.01.28_追加
UF1.NoL.Visible = True
UF1.PackingForm.Visible = True
UF1.PackingComment.Visible = True
  
UF1.MP1.Visible = True
UF1.MPEX.Visible = True '2023.5.17_追加
Call Fm_Clear("Click")

UF1.MP1.Visible = True
UF1.MP1.Width = NowSizeMP '2021.2.15_薄板の時非表示にできない....ｻｲｽﾞ修正で逃
UF1.MPEX.Height = NowSizeH
UF1.MPEX.Width = NowSizeW

UF1.Repaint
If UF1.Test Then
  UF1.WorkFrame.Visible = True
  UF1.NoFrame.Visible = True
End If

'2023.8.21_追加
UF1.VC_A.Visible = True
UF1.VC_B.Visible = True
UF1.VC_L_A.Visible = True
UF1.VC_L_B.Visible = True

UF1.MPEX.Tag = "" 'フラグリセット

End Sub

Public Sub MaterialsMsgBox(OrdNo As String)  '注意喚起用
If UF1.Lot = "" Then Exit Sub
'オーダーNo検索
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption

'シート検索からDB検索に変更_2023.4.18
Dim FPath As String
If UF1.DBflag = "" Then
  FPath = MainPath
Else 'DBflagに文字列_'2021.03.11
  FPath = NextPath
End If

Dim target As String:    target = "SIKAODRNOW.accdb"
Dim FilePath As String:  FilePath = FPath & target
  If ReadCheck(FilePath, 1) = False Or TimeDesignation Or FileDesignation = False Then  '1：ファイル,2：フォルダ
    Select Case True
    Case FPath = NextPath
      MsgBox "日稲:" & target & "がありません" & "_MaterialsMsgBox", vbCritical 'ここでは処理は止めない_2021.06.23
      Exit Sub
    Case FPath = MainPath
      FPath = NextPath
      FilePath = FPath & target
      If ReadCheck(FilePath, 1) = False Then '1：ファイル,2：フォルダ
        MsgBox "日稲:" & target & "がありません" & "_MaterialsMsgBox", vbCritical   '台帳も日軽稲沢もSikaHikiがない
        Exit Sub
      End If
    End Select
  End If
  
Dim TableName  As String:    TableName = "仕掛"
Dim fieldsfilter As String:  fieldsfilter = ""
Dim SQLfilter As String:     SQLfilter = "[受注番号]='" & OrdNo & "'" 'ｵｰﾀﾞｰNo検索

Dim i As Long, a As Long
Dim R_Ord As Long, R_Packing As Long, R_Lot As Long, R_Hik As Long, R_厚 As Long, _
R_巾 As Long, R_丈 As Long, R_材質 As Long, R_調質 As Long, R_用途コード As Long, R_合紙 As Long, R_納入先名称 As Long, R_取引先名称 As Long

Dim SIKAHIKAcc: SIKAHIKAcc = GetFieldsArr(FilePath, TableName, fieldsfilter)
Dim no As Long
If Not IsEmpty(SIKAHIKAcc) Then
  For i = 1 To UBound(SIKAHIKAcc, 2) 'ﾌｨｰﾙﾄﾞ名
    Select Case SIKAHIKAcc(1, i)
    Case "受注番号":     R_Ord = i
    Case "包装仕様NO":   R_Packing = i
    Case "受注板厚":     R_厚 = i
    Case "受注板幅":     R_巾 = i
    Case "受注板丈":     R_丈 = i
    Case "受注材質":     R_材質 = i
    Case "受注調質":     R_調質 = i
    Case "用途ｺｰﾄﾞ":     R_用途コード = i
    Case "合紙":         R_合紙 = i
    Case "納入先名称":   R_納入先名称 = i
    
    Case "取引先名称":   R_取引先名称 = i '2026.07.10_取引必要追加
    End Select
  Next i
End If
Dim KenAcc: KenAcc = GetRecordsArr(FilePath, TableName, SQLfilter)

If IsEmpty(KenAcc) Then Exit Sub

'置き換え
Dim 包装仕様NO As String: 包装仕様NO = KenAcc(1, R_Packing)
Dim 用途コード As String: 用途コード = KenAcc(1, R_用途コード)

'オーダーサイズではなく製造を見るように変更_2024.8.28(ラテラルボーが製造を見るに決まった)
Dim 厚 As Single:         If UF1.ATU <> "" Then 厚 = val(UF1.ATU) 'KenAcc(1, R_厚)
Dim 巾 As Single:         If UF1.Hab <> "" Then 巾 = val(UF1.Hab) 'KenAcc(1, R_巾)
Dim 丈 As Single:         If UF1.Tak <> "" Then 丈 = val(UF1.Tak) 'KenAcc(1, R_丈)

Dim 材質 As String:       材質 = KenAcc(1, R_材質)
Dim 調質 As String:       調質 = KenAcc(1, R_調質)
Dim 納入先 As String:     納入先 = KenAcc(1, R_納入先名称)
Dim 取引先 As String:     取引先 = KenAcc(1, R_取引先名称)

Dim Dict As Object: Set Dict = CreateObject("Scripting.Dictionary")
Dim ArrDic() As String, Result As String

Dim LotNo As String: LotNo = UF1.Lot.Text
Dim flag As Boolean: flag = AdvanceCheck(LotNo)
Dim 先行 As String:  先行 = "先行データ票が必要です。"

Dim Temp As String: Temp = PackagingSpecificationNo(包装仕様NO, 用途コード, 厚, 巾, 丈, 材質, 調質, 納入先, 取引先)
If Temp = "" Then
  '注意喚起がない場合でも先行必要時は実行_2023.12.22
  If flag Then
    Temp = 先行
    Call MsgError(Temp, Lname, 4)
    Call 音楽を流す("必要です.wav")
  End If
Exit Sub
End If

Dim arr: arr = Split(Temp, "_") '形状とコメントを分ける　ここは"_"

If arr(LBound(arr, 1)) <> "コイル" Then
  Dim SpiltArr: SpiltArr = Split(arr(UBound(arr, 1)), ",") '複数ある場合
  Dim nCount As Long: nCount = UBound(SpiltArr, 1)
  Temp = ""
  If nCount >= 2 Then '2つ以上あるとき
    For i = LBound(SpiltArr, 1) To nCount
      If Not (Temp Like "*" & SpiltArr(i) & "*") Then Temp = Temp & "," & SpiltArr(i) '","で改行表示に修正_2023.7.7
    Next i
  ElseIf nCount <= 1 Then
    If nCount = -1 Then
      Exit Sub '変数なし対策
    End If
    Temp = SpiltArr(LBound(SpiltArr, 1))
  End If
  
  If flag Then '先行が必要な場合追記_2023.12.22
    Temp = Temp & "," & 先行
    Call 音楽を流す("必要です.wav")
  End If
  
  '2024.1.18
  Dim containsSubstring As Boolean: containsSubstring = InStr(1, Temp, "直角度の測定") > 0
  If containsSubstring Then Call 音楽を流す("直角度.wav")
 
  
  ArrDic = Split(Temp, ",")   ' 文字列を" "で分割して配列に格納
  For i = LBound(ArrDic) To UBound(ArrDic)
    Dict(ArrDic(i)) = 0 ' 辞書オブジェクトを使用して重複を削除
  Next i
  
  Result = Join(Dict.Keys, ",")
  Temp = Result
  
  Set Dict = Nothing
  
  Call MsgError(Temp, Lname, 4) '板の時だけBox表示
  Exit Sub
End If

'Stop

'<<<<<<<<<<<<<<<<<<<<<旧方式>>>>>>>>>>>>>>>>>>>>>>>>
'If Hist(a, R_Packing) = "1P0113" Or Hist(a, R_Packing) = "1P0117" Or Hist(a, R_Packing) = "1P0121" _
'Or Hist(a, R_Packing) = "1P1231" Or Hist(a, R_Packing) = "3P1167" Then Call MsgError("裸梱包です。注意願います。", Lname, 4) '裸
'
'If Hist(a, R_Packing) = "1P1202" Then Call MsgError("白銅:裸梱包です。" & "ラテラルボー測定が必要です。" & "注意願います。", Lname, 4)  '裸
'
'If Hist(a, R_Packing) = "1P1177" Then Call MsgError("ラテラルボー測定が必要です。" & "注意願います。", Lname, 4) 'IKボードの条件はオーダー寸法の為製造しか拾ってないのでここでは判定不可
'
'If Hist(a, R_Packing) = "1P1178" Then Call MsgError("ラテラルボー測定が必要です。" & "注意願います。", Lname, 4)
'
'If Hist(a, R_Packing) = "1P1179" Then Call MsgError("ラテラルボー測定が必要です。" & "注意願います。", Lname, 4)
'
'If Hist(a, R_Packing) = "1P1216" Then Call MsgError("ｺｰﾐ金属:裸梱包です。注意願います。", Lname, 4)  '裸
'
'If Hist(a, R_Packing) = "1P1234" Then Call MsgError("ﾊｰﾄﾞﾎﾞｰﾄﾞ向き指定「有」です。注意願います。", Lname, 4)  'HB　日鉄物産
'
'If Hist(a, R_Packing) = "1P0118" Or Hist(a, R_Packing) = "1P0119" _
'Or Hist(a, R_Packing) = "1P1189" Then Call MsgError("縦バンドが「有」です。注意願います。", Lname, 4) '縦
'
'If Hist(a, R_Packing) = "1P1147" Then Call MsgError("4側面ﾊｰﾄﾞﾎﾞｰﾄﾞ保護です。注意願います。", Lname, 4) '日本無線
'
'If Hist(a, R_Packing) = "1P1230" Then Call MsgError("検査票に「両面VC」表記が必要です。注意願います。", Lname, 4) 'アルナ輸送
'
'If Hist(a, R_Packing) = "1P1230" Then Call MsgError("委託圧延の為、TP添付が必要です。注意願います。", Lname, 4) '三菱委託圧延
'
'If Hist(a, R_Packing) = "1P1239" Then Call MsgError("4隅ｾﾛﾃｰﾌﾟ2重保護です。注意願います。", Lname, 4) 'ﾊﾞﾗﾂｷGFS
'
''If Hist(a, R_Packing) = "1P1118" Then Call MsgError("ｽﾌﾟﾚｰﾏｰｷﾝｸﾞが必要です", Lname, 4) 'ｽﾐｾｲ
''住精でJIS、ASMEでマーキング有無が変わる為包装仕様書Noでの判定ではなく用途コードでの判定に修正_2022.09.16
''Q-S-KE-089_ASME規格　住友精密工業向け板側面マーキング作業について
'If Hist(a, R_用途コード) = "R112" Or Hist(a, R_用途コード) = "R113" _
'Or Hist(a, R_用途コード) = "R115" Then Call MsgError("用途ｺｰﾄﾞ:" & Hist(a, R_用途コード) & "_" & "ｽﾌﾟﾚｰﾏｰｷﾝｸﾞが必要です。" & "  台紙ｻﾝﾌﾟﾙが2枚必要です。", Lname, 4)
'
'If Hist(a, R_用途コード) = "G283" Or Hist(a, R_用途コード) = "G284" _
'Or Hist(a, R_用途コード) = "R114" Then Call MsgError("用途ｺｰﾄﾞ:" & Hist(a, R_用途コード) & "_" & "不要です。", Lname, 4)
'
'
'If Hist(a, R_Packing) = "1P1116" Then Call MsgError("探傷試験が必要です。注意願います。", Lname, 4)  'ﾌｼﾞｱｾﾁﾚﾝｺｳｷﾞﾖｳ(ｶ
'
''2022.09_白銅に強化パレット使用クレーム
'If Hist(a, R_Packing) = "1P1202" Or Hist(a, R_Packing) = "1P1177" Or Hist(a, R_Packing) = "1P1178" Or Hist(a, R_Packing) = "1P1179" Then
'  If Val(Hist(a, R_厚)) <= 13# Then
'    If Val(Hist(a, R_巾)) >= 1490# And Val(Hist(a, R_巾)) <= 1540# Then
'      If Val(Hist(a, R_丈)) >= 2700# And Val(Hist(a, R_丈)) <= 3090# Then
'        Call MsgError("白銅強度UP使用「不可」！！", Lname, 4) '白銅に限り了承が無いため*Q-S-QA-ALL-167_国内向け製品板厚13mm以下強度UP5*10パレット使用基準
'        Exit Sub
'      End If
'    End If
'  End If
'End If
'
'If Hist(a, R_Packing) = "1P0001" Then
'  If Val(Hist(a, R_厚)) <= 13# Then
'    If Val(Hist(a, R_巾)) >= 1490# And Val(Hist(a, R_巾)) <= 1540# Then
'      If Val(Hist(a, R_丈)) >= 2700# And Val(Hist(a, R_丈)) <= 3090# Then
'        Call MsgError("桁角材高さ90mm、強度UP5×10ﾊﾟﾚｯﾄ使用の事。注意願います。", Lname, 4) '強度UP5*10
'      End If
'    End If
'  End If
'End If

End Sub

'DBから条件抽出して表示させる
Public Function PackagingSpecificationNo(PackagingNO As String, 用途 As String, 厚 As Single, 幅 As Single, 丈 As Single, 材質 As String, 調質 As String, Optional 納入先 As String = "", Optional 取引先 As String = "") As String
'If PackagingNO = "" Then Exit Function
Dim FPath As String:         FPath = PackingPath
Dim TableName As String:     TableName = "注意_包装仕様"
Dim target As String:        target = "梱包資材マスタ.accdb"
Dim FilePath As String:      FilePath = FPath & target
Dim fieldsfilter As String:  fieldsfilter = ""
Dim Hikifields As Variant:   Hikifields = GetFieldsArr(FilePath, TableName, fieldsfilter)
FilePath = FPath & target
Dim PackF(): PackF = GetFieldsArr(FilePath, TableName, fieldsfilter)
If IsEmpty(PackF) Then Exit Function
Dim C_用途コード As Long, C_包装仕様NO As Long, C_コメント As Long, C_備考 As Long, i As Long '＊
Dim C_受注材質 As Long, C_受注調質 As Long, _
C_厚下 As Long, C_厚上 As Long, C_板幅下 As Long, C_板幅上 As Long, C_板丈下 As Long, C_板丈上 As Long, _
C_フラグ As Long, C_形状 As Long, C_納入先 As Long, C_取引先 As Long
For i = LBound(PackF, 2) To UBound(PackF, 2)
  Select Case PackF(1, i)
  Case "用途コード":   C_用途コード = i
  Case "包装仕様NO":   C_包装仕様NO = i
  Case "コメント":     C_コメント = i
  Case "備考":         C_備考 = i
  Case "受注材質":     C_受注材質 = i
  Case "受注調質":     C_受注調質 = i
  Case "厚下":         C_厚下 = i
  Case "厚上":         C_厚上 = i
  Case "板幅下":       C_板幅下 = i
  Case "板幅上":       C_板幅上 = i
  Case "板丈下":       C_板丈下 = i
  Case "板丈上":       C_板丈上 = i
  Case "フラグ":       C_フラグ = i
  Case "形状":         C_形状 = i
  Case "納入先":       C_納入先 = i
  Case "取引先":       C_取引先 = i '2026.07.10
  End Select
Next i

Dim SQLfilter As String: SQLfilter = "[包装仕様NO]='" & PackagingNO & "'"
Dim TempHiki As Variant: TempHiki = GetRecordsArr(FilePath, TableName, SQLfilter)
If IsEmpty(TempHiki) Then
  PackagingSpecificationNo = ""
  Exit Function
End If

Dim Cont As Long: Cont = UBound(TempHiki, 1)
Dim Kari As String
'2023.9.11_納入先フラグ追加
'2024.3.7_納入先を=からLikeに変更
'2026.07.10_取引追加
If Cont = 1 Then
  If TempHiki(1, C_フラグ) = "" Then PackagingSpecificationNo = TempHiki(1, C_コメント) '表示条件なし
  If TempHiki(1, C_フラグ) <> "" Then '表示条件あり
    Select Case TempHiki(1, C_フラグ)
    Case "用途コード"
      If 用途 = TempHiki(1, C_用途コード) Then PackagingSpecificationNo = TempHiki(1, C_コメント)
    Case "寸法"
      If (厚 >= val(TempHiki(1, C_厚下)) And 厚 <= val(TempHiki(1, C_厚上))) _
      And (幅 >= val(TempHiki(1, C_板幅下)) And 幅 <= val(TempHiki(1, C_板幅上))) _
      And (丈 >= val(TempHiki(1, C_板丈下)) And 丈 <= val(TempHiki(1, C_板丈上))) Then PackagingSpecificationNo = TempHiki(1, C_コメント)
    Case "納入先"
      If 納入先 <> "" Then '欠陥引当？納入が空のことがあるその場合すべてラテラルボー扱いになる_2024.3.11
        If IsMatchConditionField(CStr(TempHiki(1, C_納入先)), 納入先) Then PackagingSpecificationNo = TempHiki(1, C_コメント)
      End If
    '2024.8.28 日車対策
    Case "寸法納入"
      If 納入先 <> "" Then
        If IsMatchConditionField(CStr(TempHiki(1, C_納入先)), 納入先) Then
          If (厚 >= val(TempHiki(1, C_厚下)) And 厚 <= val(TempHiki(1, C_厚上))) _
          And (幅 >= val(TempHiki(1, C_板幅下)) And 幅 <= val(TempHiki(1, C_板幅上))) _
          And (丈 >= val(TempHiki(1, C_板丈下)) And 丈 <= val(TempHiki(1, C_板丈上))) Then PackagingSpecificationNo = TempHiki(1, C_コメント)
        End If
      End If
    '2026.07.10_取引先・除外・部分一致対応
    Case "寸法取引"
      If 取引先 <> "" Then
        If (厚 >= val(TempHiki(1, C_厚下)) And 厚 <= val(TempHiki(1, C_厚上))) _
        And (幅 >= val(TempHiki(1, C_板幅下)) And 幅 <= val(TempHiki(1, C_板幅上))) _
        And (丈 >= val(TempHiki(1, C_板丈下)) And 丈 <= val(TempHiki(1, C_板丈上))) _
        And IsMatchConditionField(CStr(TempHiki(1, C_納入先)), 納入先) _
        And IsMatchConditionField(CStr(TempHiki(1, C_取引先)), 取引先) Then
          PackagingSpecificationNo = TempHiki(1, C_コメント)
        End If
      End If
    End Select
  End If
  PackagingSpecificationNo = TempHiki(1, C_形状) & "_" & PackagingSpecificationNo '板orコイル
ElseIf Cont >= 2 Then '複数の注意項目がある包装仕様No
  For i = 1 To Cont
    If TempHiki(i, C_フラグ) = "" Then
      Kari = TempHiki(i, C_コメント)
      PackagingSpecificationNo = Kari & "," & PackagingSpecificationNo
    End If
    If TempHiki(i, C_フラグ) <> "" Then
      Select Case TempHiki(i, C_フラグ)
      Case "用途コード"
        If 用途 = TempHiki(i, C_用途コード) Then
          Kari = TempHiki(i, C_コメント)
          PackagingSpecificationNo = Kari & "," & PackagingSpecificationNo 'コメントを重ねていく
        End If
      Case "寸法"
        If (厚 >= val(TempHiki(i, C_厚下)) And 厚 <= val(TempHiki(i, C_厚上))) _
        And (幅 >= val(TempHiki(i, C_板幅下)) And 幅 <= val(TempHiki(i, C_板幅上))) _
        And (丈 >= val(TempHiki(i, C_板丈下)) And 丈 <= val(TempHiki(i, C_板丈上))) Then
          Kari = TempHiki(i, C_コメント)
          PackagingSpecificationNo = Kari & "," & PackagingSpecificationNo 'コメントを重ねていく
        End If
      Case "納入先"
        If 納入先 <> "" Then '欠陥引当？納入が空のことがあるその場合すべてラテラルボー扱いになる_2024.3.11
          If IsMatchConditionField(CStr(TempHiki(i, C_納入先)), 納入先) Then
            Kari = TempHiki(i, C_コメント)
            PackagingSpecificationNo = Kari & "," & PackagingSpecificationNo
          End If
        End If
      '2024.8.28 日車対策
      Case "寸法納入"
        If 納入先 <> "" Then
          If IsMatchConditionField(CStr(TempHiki(i, C_納入先)), 納入先) Then
            If (厚 >= val(TempHiki(i, C_厚下)) And 厚 <= val(TempHiki(i, C_厚上))) _
            And (幅 >= val(TempHiki(i, C_板幅下)) And 幅 <= val(TempHiki(i, C_板幅上))) _
            And (丈 >= val(TempHiki(i, C_板丈下)) And 丈 <= val(TempHiki(i, C_板丈上))) Then PackagingSpecificationNo = TempHiki(i, C_コメント)
          End If
        End If
      '2026.07.10_取引先・除外・部分一致対応
      Case "寸法取引"
        If 取引先 <> "" Then
          If (厚 >= val(TempHiki(i, C_厚下)) And 厚 <= val(TempHiki(i, C_厚上))) _
          And (幅 >= val(TempHiki(i, C_板幅下)) And 幅 <= val(TempHiki(i, C_板幅上))) _
          And (丈 >= val(TempHiki(i, C_板丈下)) And 丈 <= val(TempHiki(i, C_板丈上))) _
          And IsMatchConditionField(CStr(TempHiki(i, C_納入先)), 納入先) _
          And IsMatchConditionField(CStr(TempHiki(i, C_取引先)), 取引先) Then
            Kari = TempHiki(i, C_コメント)
            PackagingSpecificationNo = Kari & "," & PackagingSpecificationNo
          End If
        End If
      End Select
    End If
  Next i
  
  PackagingSpecificationNo = TempHiki(1, C_形状) & "_" & PackagingSpecificationNo '形状(板orコイル)とコメント
  '末尾の，を消すと複数表示されない･････
'  If Right(PackagingSpecificationNo, 1) = "," Then PackagingSpecificationNo = Left(PackagingSpecificationNo, Len(PackagingSpecificationNo) - 1) '最後のコンマを削除
End If

End Function

'============================================================
' マスタ側の条件文字列（カンマ区切り、OR結合）と実値を比較する
' 例: "≠K.T.N. CO. LTD,≠ﾆﾎﾝﾊﾂｼﾞﾖｳ*"
'   → 「K.T.N. CO. LTDと一致しない」または「ﾆﾎﾝﾊﾂｼﾞﾖｳで始まらない」のいずれかでTrue
' ルール:
'   ・FieldValue（マスタ側）が空 or "＊" のときは常にTrue（条件指定なし）
'   ・カンマ区切りは常にOR結合（除外条件同士でもOR）
'   ・各トークン先頭の「≠」は除外指定（一致しなければTrue）
'   ・各トークン内の「*」位置でワイルドカード判定
'       xxx*  → 前方一致（xxxで始まる）
'       *xxx  → 後方一致（xxxで終わる）
'       *xxx* → 部分一致（xxxを含む）
'       xxx   → 完全一致
'============================================================
Private Function IsMatchConditionField(FieldValue As String, target As String) As Boolean
Dim CondStr As String: CondStr = Trim(FieldValue)

'マスタ側が空 or ＊のみ → 条件指定なし扱いで常にTrue
If CondStr = "" Or CondStr = "＊" Or CondStr = "*" Then
  IsMatchConditionField = True
  Exit Function
End If

Dim Tokens() As String: Tokens = Split(CondStr, ",")
Dim t As Long, Token As String
Dim IsExclude As Boolean, Pattern As String, IsHit As Boolean

For t = LBound(Tokens) To UBound(Tokens)
  Token = Trim(Tokens(t))
  If Token <> "" Then
    IsExclude = False
    If Left(Token, 1) = "≠" Then
      IsExclude = True
      Token = Mid(Token, 2)
    End If

    IsHit = IsWildcardMatch(Token, target)

    If IsExclude Then
      If Not IsHit Then
        IsMatchConditionField = True
        Exit Function
      End If
    Else
      If IsHit Then
        IsMatchConditionField = True
        Exit Function
      End If
    End If
  End If
Next t

IsMatchConditionField = False
End Function

'============================================================
' Pattern中の「*」位置により、前方一致／後方一致／部分一致／完全一致を判定
'============================================================
Private Function IsWildcardMatch(Pattern As String, target As String) As Boolean
Dim HasHead As Boolean, HasTail As Boolean
Dim Core As String
Dim p As String: p = Pattern

HasHead = (Left(p, 1) = "*")   '先頭が* → 後方一致判定に使う目印
HasTail = (Right(p, 1) = "*")  '末尾が* → 前方一致判定に使う目印

Core = p
If HasHead Then Core = Mid(Core, 2)
If HasTail Then Core = Left(Core, Len(Core) - 1)

If HasHead And HasTail Then
  IsWildcardMatch = (InStr(1, target, Core, vbTextCompare) > 0)          '*xxx* 部分一致
ElseIf HasTail Then
  IsWildcardMatch = (Left(target, Len(Core)) = Core)                     'xxx*  前方一致
ElseIf HasHead Then
  IsWildcardMatch = (Right(target, Len(Core)) = Core)                    '*xxx  後方一致
Else
  IsWildcardMatch = (StrComp(target, Core, vbTextCompare) = 0)           'xxx   完全一致
End If
End Function

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'2022.04.11_公差の書き換え***VC_208PRNFを貼る製品全てをH207(中精度)公差へ書き換える
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Public Sub NflatRewrite()
UF1.ReSize.Tag = "有" '履歴として残す
Call Posting
End Sub

Public Function UnitRecord(n As Single) As Single 'ラテラルボー測定値を0.5ピッチに変換**Q-S-QA-ALL-153
'ラテラルボー入力を3桁に制限(小数点を含めば小数点第1位までの入力になる)***2022.08.25**数字の丸めが2回入る対策
'Sgn関数 正(マイナスでない)」の場合「1」が返る
'2023.11.27_直角度追加
Dim flag As Boolean
If Sgn(n) = 1 Then flag = True '正
n = Abs(n)
Select Case n
Case 0:            UnitRecord = 0
Case 0.1 To 0.5:   UnitRecord = 0.5
Case 0.51 To 1#:   UnitRecord = 1#
Case 1.01 To 1.5:  UnitRecord = 1.5
Case 1.51 To 2#:   UnitRecord = 2#
Case 2.01 To 2.5:  UnitRecord = 2.5
Case 2.51 To 3#:   UnitRecord = 3#
Case 3.01 To 3.5:  UnitRecord = 3.5
Case 3.51 To 4#:   UnitRecord = 4#
Case 4.01 To 4.5:  UnitRecord = 4.5
Case 4.51 To 5#:   UnitRecord = 5#
Case 5.01 To 999:  UnitRecord = n
End Select
If Not flag Then UnitRecord = val(UnitRecord) * (-1) '負の場合
End Function
Public Sub LookingYellow() '黄色のバックカラーがあったら消す
Dim BorderCol As String: BorderCol = RGB(255, 0, 0)
Dim BackCol As String:   BackCol = RGB(255, 255, 0)
Dim ResetBor As String:  ResetBor = RGB(0, 0, 0)
Dim ResetBac As String:  ResetBac = RGB(255, 255, 255)

Dim i As Long
For i = 1 To 14
  If UF1("S" & i).BorderColor = BorderCol Then UF1("S" & i).BorderColor = ResetBor
  If UF1("S" & i).BackColor = BackCol Then UF1("S" & i).BackColor = ResetBac
  If UF1("U" & i).BorderColor = BorderCol Then UF1("U" & i).BorderColor = ResetBor
  If UF1("U" & i).BackColor = BackCol Then UF1("U" & i).BackColor = ResetBac
  If UF1("C" & i).BorderColor = BorderCol Then UF1("C" & i).BorderColor = ResetBor
  If UF1("C" & i).BackColor = BackCol Then UF1("C" & i).BackColor = ResetBac
  If UF1("D" & i).BorderColor = BorderCol Then UF1("D" & i).BorderColor = ResetBor
  If UF1("D" & i).BackColor = BackCol Then UF1("D" & i).BackColor = ResetBac
  If UF1("E" & i).BorderColor = BorderCol Then UF1("E" & i).BorderColor = ResetBor
  If UF1("E" & i).BackColor = BackCol Then UF1("E" & i).BackColor = ResetBac
Next i
For i = 1 To 10
  If UF1("F" & i).BorderColor = BorderCol Then UF1("F" & i).BorderColor = ResetBor
  If UF1("F" & i).BackColor = BackCol Then UF1("F" & i).BackColor = ResetBac
Next i
End Sub
Public Function CheckForLot(ws As Worksheet, Lot As String, Dcaunt As Long) As Boolean 'シート名,探す文字列,探す列_2024.11.06
If Lot = "" Then
  CheckForLot = False
  Exit Function
End If

Dim lastRow As Long:   lastRow = UR(ws, 1) ' シートの最終行を取得
Dim StartRow As Long:  StartRow = Application.WorksheetFunction.Max(1, lastRow - 4) ' 最終行から上に5行分の開始行を決定
Dim i As Long
Dim DataArray():       DataArray = ws.Range(ws.Cells(StartRow, 1), ws.Cells(lastRow, Dcaunt)) ' データを配列に格納

' 配列内に"Lot"があるかをチェック
CheckForLot = False
For i = 1 To UBound(DataArray, 1)
  If InStr(DataArray(i, Dcaunt), Lot) > 0 Then
    CheckForLot = True
    Exit Function
  End If
Next i
End Function

