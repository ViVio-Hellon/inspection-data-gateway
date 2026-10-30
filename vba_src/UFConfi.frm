Attribute VB_Name = "UFConfi"
Attribute VB_Base = "0{DE507CF8-8A7B-48F1-BC04-C6CCB1CB3D6D}{E27BF97C-2DC4-4367-BDE8-EC9B0628F557}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
' Excelウィンドウ状態退避用
Private g_sngTopSV As Single ' 上端位置
Private g_sngLeftSV As Single ' 左端位置
Private g_sngWidthSV As Single ' 幅
Private g_sngHeightSV As Single ' 高さ

Private Sub CommandButton2_Click()
Dim Rows As String, i As Long
Dim RowPosi As Long

Dim Temp As String: Temp = Me.MotoData.Caption
  With Me.List1
    For i = 0 To .ListCount - 1 'リストボックスの行数
      If .Selected(i) Then '選択されていた場合
        If Rows = "" Then
          Rows = Me.List1.List(i) '1こめのとき
        Else
          Rows = Rows & ", " & Me.List1.List(i) '複数あるとき
        End If
      End If
    Next i
  End With
  
  If Rows = "" Then
    MsgBox "選択してください", vbExclamation
    Exit Sub
  ElseIf InStr(Rows, "再測定") > 0 Or InStr(Rows, "その他") > 0 Then
    If Me.その他 = "" Then
      MsgBox "理由を入力してください", vbExclamation
      Exit Sub
    Else
      Rows = Rows & ", " & Me.その他
    End If
  End If
  
If Not (Rows Like "*" & Me.その他 & "*") Then Rows = Rows & ", " & Me.その他
'If Me.その他 <> "" Then Rows = Rows & ", " & Me.その他

'板廃棄の場合任意の検査箇所全消し**2024.2.5'まとめて削除処理:Logには位置0_:理由選択:板廃棄:検査員名で出力
If InStr(Rows, "板廃棄") > 0 Then
'150Kg関係なく残す_2024.4.18 家木より鈴木課長メールの確認
'------------------------------測定値を消したことで起動している際の処理_2024.5.17
  Dim Posi As String:      Posi = UFConfi.Kpop
  Dim nNo As Long:         nNo = UFConfi.Knum
  
  If nNo <> 0 Then '位置0は不要
    Dim EraseData As String: EraseData = UFConfi.MotoData.Caption
    UF1(Posi & nNo) = EraseData '転記する際は消したデータを戻す(印刷用)
    Select Case Posi
    Case "S": RowPosi = 1
    Case "U": RowPosi = 2
    Case "C": RowPosi = 3
    Case "D": RowPosi = 4
    Case "E": RowPosi = 5
    End Select
    Call LoopJudg(, RowPosi, RowPosi) '戻した数値の検査判定が必要
  End If
  
  Call NG_Posting(Me.Kpop) '転記処理
  Call NG_Erase(Me.Kpop, "板廃棄") '消去_NG_Eraseに引数追加、2つ目に文字列入力でUF1.Showはこちらで処理
End If

Dim k As String: k = Me.KenName.Caption
Dim path As String: path = HistoryPath
Select Case Me.LineName
Case "L-1":     Call WriteLog(Me.Kpop & Me.Knum & "_" & Temp & ":" & "理由選択:" & Rows & ":" & k, "", path & "L-1\L-1_Log")
Case "LVC":     Call WriteLog(Me.Kpop & Me.Knum & "_" & Temp & ":" & "理由選択:" & Rows & ":" & k, "", path & "LVC\LVC_Log")
Case "HVC":     Call WriteLog(Me.Kpop & Me.Knum & "_" & Temp & ":" & "理由選択:" & Rows & ":" & k, "", path & "HVC\HVC_Log")
Case "ﾊﾞﾗﾝｻｰ":  Call WriteLog(Me.Kpop & Me.Knum & "_" & Temp & ":" & "理由選択:" & Rows & ":" & k, "", path & "バランサー\バランサー_Log")
Case "ﾄｯﾄ":     Call WriteLog(Me.Kpop & Me.Knum & "_" & Temp & ":" & "理由選択:" & Rows & ":" & k, "", path & "トット\トット_Log")
Case "Test":    Call WriteLog(Me.Kpop & Me.Knum & "_" & Temp & ":" & "理由選択:" & Rows & ":" & k, "", path & "Test\Test_Log")
Case Else:      Call WriteLog(Me.Kpop & Me.Knum & "_" & Temp & ":" & "理由選択:" & Rows & ":" & k, "", path & "丸徳\丸徳_Log")
End Select
    
UF1.Show vbModeless
DoEvents
Unload UFConfi
    
End Sub


Private Sub UserForm_Layout()
Application.Visible = False 'Excel非表示
' Excelウィンドウ位置を同期させる
'With Application
'  If .WindowState = xlNormal Then
'    .Top = Me.Top + 1
'    .Left = Me.Left + 1
'  End If
'End With
End Sub
Private Sub UserForm_Activate()
Dim Temp As Variant:  Temp = InitialSetting(Me)
Call FormSetting
'MsgBox Temp(0) & "_" & Temp(1) & "_" & Temp(2) & "_" & Temp(3) & "_" & Temp(4) & "_" & Temp(5)

Application.Visible = False 'Excel非表示
Call WindowMenuButtonDel(UFConfi)
'With Application ' フォームの後ろにExcelウィンドウを隠す
'  .Top = Me.Top + 1
'  .Left = Me.Left + 1
'  .Width = Me.Width - 2
'  .Height = Me.Height - 2
'End With

Select Case UFConfi.LineName
Case "HVC", "Test"
  Me.NewHB1.Visible = True
  Me.NewHB2.Visible = True
  Me.NewHB3.Visible = True
  Me.NewHB4.Visible = True '2024.4.12_追加
Case Else
  Me.NewHB1.Visible = False
  Me.NewHB2.Visible = False
  Me.NewHB3.Visible = False
  Me.NewHB4.Visible = False
End Select
'Call WindowMenuButtonDel(Me) 'ｳｨﾝﾄﾞｳ×ﾎﾞﾀﾝを消す。
If Me.MP1.P1.Visible Then 'ページ表示されている場合のみ***2023.4.19
  UFConfi.Kpop.SetFocus
  UFConfi.TB1.SetFocus
End If

Dim path As String: path = HistoryPath
Select Case UFConfi.LineName
Case "L-1":     Call WriteLog("UFConfi起動" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", path & "L-1\L-1_Log")
Case "LVC":     Call WriteLog("UFConfi起動" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", path & "LVC\LVC_Log")
Case "HVC":     Call WriteLog("UFConfi起動" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", path & "HVC\HVC_Log")
Case "ﾊﾞﾗﾝｻｰ":  Call WriteLog("UFConfi起動" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", path & "バランサー\バランサー_Log")
Case "ﾄｯﾄ":     Call WriteLog("UFConfi起動" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", path & "トット\トット_Log")
Case "Test":    Call WriteLog("UFConfi起動" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", path & "Test\Test_Log")
Case Else:      Call WriteLog("UFConfi起動" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", path & "丸徳\丸徳_Log")
End Select
End Sub
Private Sub UserForm_Initialize()
Call GetZoom(Me)

Call WindowMenuButtonDel(UFConfi)
Call Reasons_Add

' 現在のExcelウィンドウ状態を保管
'With Application
'  .WindowState = xlNormal
'  g_sngTopSV = .Top
'  g_sngLeftSV = .Left
'  g_sngWidthSV = .Width
'  g_sngHeightSV = .Height
'End With
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) '×では閉じないようにする
' 保管情報を元にExcelウィンドウを元に戻す
'With Application
'  .Top = g_sngTopSV
'  .Left = g_sngLeftSV
'  .Width = g_sngWidthSV
'  .Height = g_sngHeightSV
'End With
Call Posting
Dim Temp As Variant:  Temp = InitialSetting(UF1, "サイズ保存")
End Sub

Private Sub CommandButton1_Click()
Dim Lname As String: Lname = Me.LineName
Dim Temp As String:  Temp = Me.MotoData.Caption
Dim k As String:     k = Me.KenName.Caption
Dim path As String:  path = HistoryPath

If UFConfi("List2").Visible = False Then GoTo SkipA
'List2が表示されている状態
Dim Rows As String, i As Long

  With Me.List2
    For i = 0 To .ListCount - 1 'リストボックスの行数
      If .Selected(i) Then '選択されていた場合
        If Rows = "" Then
          Rows = Me.List2.List(i) '1こめのとき
        Else
          Rows = Rows & ", " & Me.List2.List(i) '複数あるとき
        End If
      End If
    Next i
  End With
  
  If Rows = "" Then
    MsgBox "選択してください", vbExclamation
    Exit Sub
  End If
  
If Not (Rows Like "*" & Me.その他 & "*") Then Rows = Rows & ", " & Me.その他


'---------------------------------NG判定があるから起動した場合の処理
'許容、保留はNGEraseでNG色が消えた時クリアされる
Select Case UFConfi.TB1
Case "公差外"
'  If Weight(Val(UF1.ATU), Val(UF1.Hab), Val(UF1.Tak)) >= 150 Then
'150Kg関係なく残す_2024.4.18 家木より鈴木課長メールの確認
  Call NG_Posting(Me.Kpop) '転記処理
  Call NG_Erase(Me.Kpop) '消去
Case "許容":     Call NG_Pass(Me.Kpop)
Case "再測定": Call NG_Back(Me.Kpop, Me.Knum)
Case "保留":   Call NG_Pass(Me.Kpop, "保留") '2024.4.12_追加
End Select
'NG判定後再測定を入力した場合、理由選択後出力
Select Case Lname
Case "L-1":     Call WriteLog("公差外:" & Me.Kpop & Me.Knum & "_" & Temp & ":" & "理由選択:" & "再測定," & Rows & ":" & k, "", path & "L-1\L-1_Log")
Case "LVC":     Call WriteLog("公差外:" & Me.Kpop & Me.Knum & "_" & Temp & ":" & "理由選択:" & "再測定," & Rows & ":" & k, "", path & "LVC\LVC_Log")
Case "HVC":     Call WriteLog("公差外:" & Me.Kpop & Me.Knum & "_" & Temp & ":" & "理由選択:" & "再測定," & Rows & ":" & k, "", path & "HVC\HVC_Log")
Case "ﾊﾞﾗﾝｻｰ":  Call WriteLog("公差外:" & Me.Kpop & Me.Knum & "_" & Temp & ":" & "理由選択:" & "再測定," & Rows & ":" & k, "", path & "バランサー\バランサー_Log")
Case "ﾄｯﾄ":     Call WriteLog("公差外:" & Me.Kpop & Me.Knum & "_" & Temp & ":" & "理由選択:" & "再測定," & Rows & ":" & k, "", path & "トット\トット_Log")
Case "Test":    Call WriteLog("公差外:" & Me.Kpop & Me.Knum & "_" & Temp & ":" & "理由選択:" & "再測定," & Rows & ":" & k, "", path & "Test\Test_Log")
Case Else:      Call WriteLog("公差外:" & Me.Kpop & Me.Knum & "_" & Temp & ":" & "理由選択:" & "再測定," & Rows & ":" & k, "", path & "丸徳\丸徳_Log")
End Select

Exit Sub

SkipA:

Select Case UFConfi.TB1
Case "公差外"
'  If Weight(Val(UF1.ATU), Val(UF1.Hab), Val(UF1.Tak)) >= 150 Then
'150Kg関係なく残す2024.4.18 家木より鈴木課長メールの確認
  Call NG_Posting(Me.Kpop) '転記処理
  Call NG_Erase(Me.Kpop) '消去
Case "許容":     Call NG_Pass(Me.Kpop)
'---------------------------------------------------------------------------------------------------------------
Case "再測定"
'Q-S-QA-DIM-174_寸法検査結果の取り扱い基準*3回を超えて規格外れとなる場合は直ちに異常報告を行うこと_2023.7.19
'ここでは創業基準より厳しい2回で管理を行う_2023.8.15
  If val(Me.DelC.Caption) >= 2 Then '同座標を消した回数が2回以上
    UFConfi("List2").Visible = True
    MsgBox "理由を選択してください", vbExclamation

    Dim Com_data As Variant, CMB As Variant
    UFConfi("List2").clear
    Com_data = Array("入力位置間違い", "測定位置間違い", "測定ボタン押し間違い", "0点確認", "その他")
    
    With UFConfi.List2
      .ColumnCount = 1 '表示列数
      .ColumnWidths = "75"  '列幅
      .List = Com_data
    End With
    
  Exit Sub
  Else
    Call NG_Back(Me.Kpop, Me.Knum)
  End If
'---------------------------------------------------------------------------------------------------------------
Case "保留":   Call NG_Pass(Me.Kpop, "保留")
Case Else: MsgBox "指定文字列を入力してください", vbCritical
End Select
'NG判定後再測定以外を入力した場合
Select Case Lname '2023.4.20_出力Log変更UFKen削除
Case "L-1":     Call WriteLog("*UFConfi.TB1入力:" & Me.TB1.Value, "", path & "L-1\L-1_Log")
Case "LVC":     Call WriteLog("*UFConfi.TB1入力:" & Me.TB1.Value, "", path & "LVC\LVC_Log")
Case "HVC":     Call WriteLog("*UFConfi.TB1入力:" & Me.TB1.Value, "", path & "HVC\HVC_Log")
Case "ﾊﾞﾗﾝｻｰ":  Call WriteLog("*UFConfi.TB1入力:" & Me.TB1.Value, "", path & "バランサー\バランサー_Log")
Case "ﾄｯﾄ":     Call WriteLog("*UFConfi.TB1入力:" & Me.TB1.Value, "", path & "トット\トット_Log")
Case "Test":    Call WriteLog("*UFConfi.TB1入力:" & Me.TB1.Value, "", path & "Test\Test_Log")
Case Else:      Call WriteLog("*UFConfi.TB1入力:" & Me.TB1.Value, "", path & "丸徳\丸徳_Log")
End Select

UFConfi("List2").Visible = False
End Sub
Private Sub NewHB1_Click()
UFConfi.TB1 = "公差外"
End Sub
Private Sub NewHB2_Click()
UFConfi.TB1 = "許容"
End Sub
Private Sub NewHB4_Click()
UFConfi.TB1 = "保留"
End Sub
Private Sub NewHB3_Click()
UFConfi.TB1 = "再測定"
End Sub

