Attribute VB_Name = "【UFKen_10】Proofreading"
Option Explicit
Private g_lngWindowState As XlWindowState ' ウィンドウ状態

Public Sub 校正_Open(wname As String, Lname As String) '検査員,ライン名
'g_lngWindowState = Application.WindowState ' WindowStateを退避

UF校正.Show vbModeless
Unload UF校正
Select Case True
Case UF1.L1Work: UF校正.Gauge = "[L-1] P2:1000mm直尺の全長を測定してください"
Case UF1.LVC:    UF校正.Gauge = "[LVC] P2:1000mm直尺の全長を測定してください"
Case UF1.HVC:    UF校正.Gauge = "[HVC] P2:1000mm直尺の全長を測定してください"
Case UF1.Bala:   UF校正.Gauge = "[大板] P2:1000mm直尺の全長を測定してください"
Case UF1.Tott:   UF校正.Gauge = "[大板] P2:1000mm直尺の全長を測定してください"
Case UF1.Test:   UF校正.Gauge = "[Test] xx:1000mm直尺の全長を測定してください"
End Select
UF校正.Kname = wname
UF校正.LineName = Lname
Load UF校正
Call WindowMenuButtonDel(UF校正)
DoEvents
UF1.Hide
UF校正.Show vbModeless 'vbModeless'vbModal
'Call Any_Click
'Application.WindowState = g_lngWindowState ' WindowStateを復旧
End Sub
Public Function 校正(STtime As String) As Boolean '未校正の場合true'uf1.Day_Change
If (UF1.WorkButton.Tag = "jump") Or (UF1.DBflag <> "") Then '校正停止_testmode_2021.03.11
  校正 = False
  UF1.WorkButton.Tag = ""
  Exit Function
End If

If WsCheck("校正記録") = False Then 'シートがないとき
  Call AddWS("校正記録")
  校正 = True
  Exit Function
End If

Dim ws As Worksheet:  Set ws = ThisWorkbook.Sheets("校正記録")
Dim x As String:      x = TimeCheck(STtime)  '直を返す
Dim y As String:      y = YeaDayrCheck(STtime) '日付を返す
Dim k As Long:        k = ws.Cells(Rows.Count, 1).End(xlUp).row
Dim TempD As String:  TempD = ws.Cells(k, 1)
Dim TempT As String:  TempT = ws.Cells(k, 4)
UF1.ProofreadFlag = ""
With ws
  If TempD <> y Then '最終行日付をみる'Format$(Now, "yyyy年(ge年) mm/dd h時mm分ss秒")
    校正 = True
    Exit Function
  Else '最終行日付が一致したとき直を見る
    If TempT <> x Then '直のﾁｪｯｸ
      校正 = True
      Exit Function
    Else '校正済
      UF1.ProofreadFlag = "ok"
    End If
  End If
End With
End Function

'2026.06.25　テーブルの値を入れ間違い時通過できてしまっていた_修正
'数値文字列を安全にDoubleへ変換。変換できなければ success を False で返す
Public Function SafeNum(ByVal S As Variant, ByRef success As Boolean) As Double
  success = False
  SafeNum = 0
  If IsNull(S) Then Exit Function
  If IsEmpty(S) Then Exit Function
  Dim t As String: t = Trim$(CStr(S))
  If t = "" Then Exit Function
  t = Replace$(t, ",", "")        '桁区切りカンマを除去（2,052.7 → 2052.7）
  '小数点が2個以上ある異常値（2053..5 等）を弾く
  If Len(t) - Len(Replace$(t, ".", "")) > 1 Then Exit Function
  If Not IsNumeric(t) Then Exit Function
  SafeNum = CDbl(t)
  success = True
End Function

Public Sub 測定完了() '2024.4.9　DB読み取りに修正
If WsCheck("校正記録") = False Then 'シートがないとき
  Call AddWS("校正記録")
  Exit Sub
End If
'L-1 LVC HVC バランサー トット Test 丸徳
Dim tmpFaci As String:  tmpFaci = GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")
Dim ws As Worksheet:    Set ws = ThisWorkbook.Sheets("校正記録")
Dim x As String:        x = TimeCheck(Format$(Now, "hh:mm"))  '直を返す
Dim y As String:        y = YeaDayrCheck(Format$(Now, "hh:mm")) '日付を返す
Dim k As Long:          k = ws.Cells(Rows.Count, 1).End(xlUp).row + 1

Dim Temp As Variant:   Temp = StraightEdge(tmpFaci) 'ストレートエッヂの許容値下と上を返す
If IsEmptyArray(Temp) Then Exit Sub
Dim Koumin As Double, Koumax As Double '2018.12.18　Longから修正
'*****************************直尺の設定
Dim okLo As Boolean, okHi As Boolean
Koumin = SafeNum(Temp(LBound(Temp, 1), 1), okLo)
Koumax = SafeNum(Temp(UBound(Temp, 1), 1), okHi)

'2026.06.25　テーブルの値を入れ間違い時通過できてしまっていた_修正
If (Not okLo) Or (Not okHi) Then
  MsgBox "校正基準値の設定エラーです。管理者に連絡してください。", vbCritical
  Exit Sub
End If

'End3　夜勤終了時間変数
Application.ScreenUpdating = False
 If val(UF校正.SokuteiH) >= val(Koumin) And val(UF校正.SokuteiH) <= val(Koumax) Then
    If Time >= CDate("00:00") And Time <= CDate(End3) Then
      ws.Cells(k, 1) = Format$(y, "yyyy.mm.dd")
      ws.Cells(k, 2) = "夜勤日付処理" '夜勤処理があったかどうか
    Else
      ws.Cells(k, 1) = Format$(y, "yyyy.mm.dd")
    End If
    With ws
      .Cells(k, 3) = Format$(Now, "h時mm分ss秒") '時間
      .Cells(k, 4) = x '直のﾁｪｯｸ　更新後は要注意　call校正で直の一致が条件
      .Cells(k, 5) = UF校正.SokuteiH '値
      .Cells(k, 6) = UF校正.Kname '検査員
      .Cells(k, 7) = tmpFaci 'ライン
    End With
    UF校正.Kname = ""
    UF校正.Saiken.Visible = False
    Call 音楽合格
    UF校正.F1 = 0 'フラグを0にしてChangeを抜ける 'ChangeをExitへ変更フォーカス移動の記述を消す 2019.06.05'2019.06.18 KeyDownへ変更Exitイベントは2回処理してしまう
    UF校正.F1 = "end" 'フラグ  '2019.11.13　Exit イベントでendをひろってUnloadする　ここでUnloadするとオートメーションエラー
    UF1.ProofreadFlag = "ok" '2021.07.15
 Else
   UF校正.NGvalue.Visible = True 'NG値の表示
    If UF校正.INPG <> "" Then
    UF校正.NGvalue = "測定値" & " " & UF校正.INPG & " " & "mm"
    Else
    UF校正.NGvalue = ""
    End If
    With ws
    If Time >= CDate("00:00") And Time <= CDate(End3) Then
      .Cells(k, 1) = y
      .Cells(k, 2) = "夜勤日付処理" '夜勤処理があったかどうか
    Else
      .Cells(k, 1) = y
    End If
    .Cells(k, 3) = Format$(Now, "h時mm分ss秒") '時間
    .Cells(k, 5) = UF校正.SokuteiH '値
    .Cells(k, 6) = UF校正.Kname '検査員
    .Cells(k, 7) = tmpFaci 'ライン
    .Cells(k, 8) = "〆"   '印
    End With
   Call 音楽不合格
   UF校正.INPG = ""
   UF校正.SokuteiH = ""
   UF校正.F1 = 0 'フラグを0にしてChangeを抜ける
   UF校正.INPG.SetFocus 'インプットボックスに初めから数字が入るようにする
   UF校正.Saiken.Visible = True '再検査の表示
 End If
Application.ScreenUpdating = True

''カメラ校正用
'Dim x As String: x = TimeCheck(Format$(Now, "hh:mm")) '直を返す
'Dim y As String: y = YeaDayrCheck(Format$(Now, "hh:mm")) '日付を返す
'Dim WS As Worksheet
'Dim Line As String
'Dim Koumin As Double, Koumax As Double '2018.12.18　Longから修正
'Dim k As Long: k = ThisWorkbook.Sheets("校正記録").Cells(Rows.Count, 1).End(xlUp).Row + 1
' '*****************************直尺の設定
' '閾値変更時はランク付けも変更必要
'  Select Case UF校正.LineName
'  Case "L-1" '2053.2 ± 0.5
'    Line = "L-1"
'    Koumin = 2052.7
'    Koumax = 2053.7
'  Case "LVC" '2000.5 ± 0.5
'    Line = "LVC"
'    Koumin = 2000#
'    Koumax = 2001#
'  Case "HVC" '2003.8 ± 0.5'2019.01.21 閾値修正
'    Line = "HVC"
'    Koumin = 2003.3
'    Koumax = 2004.3
'  Case "ﾊﾞﾗﾝｻｰ" '2051.8 ± 0.5
'    Line = "ﾊﾞﾗﾝｻｰ"
'    Koumin = 2051.3
'    Koumax = 2052.3
'  Case "ﾄｯﾄ" '2051.8 ± 0.5
'    Line = "ﾄｯﾄ"
'    Koumin = 2051.3
'    Koumax = 2052.3
'  Case "Test" 'テスト用 2019.11.12
'    Line = "Test"
'    Koumin = 30#
'    Koumax = 3000#
'  Case Else '丸徳用
'  End Select
 '********************************************
'Application.ScreenUpdating = False
' If Val(UF校正.SokuteiH) >= Val(Koumin) And Val(UF校正.SokuteiH) <= Val(Koumax) Then '校正の閾値 2019.04.01 =　追加
'    If Time >= CDate("00:00") And Time <= CDate("7:00") Then
'      ThisWorkbook.Sheets("校正記録").Cells(k, 1) = Format$(y, "yyyy.mm.dd")
'      ThisWorkbook.Sheets("校正記録").Cells(k, 2) = "夜勤日付処理" '夜勤処理があったかどうか
'    Else
'      ThisWorkbook.Sheets("校正記録").Cells(k, 1) = Format$(y, "yyyy.mm.dd")
'    End If
'    With ThisWorkbook.Sheets("校正記録")
'      .Cells(k, 3) = Format$(Now, "h時mm分ss秒") '時間
'      .Cells(k, 4) = x   '直のﾁｪｯｸ　更新後は要注意　call校正で直の一致が条件
'      .Cells(k, 5) = UF校正.SokuteiH '値
'      .Cells(k, 6) = UF校正.Kname '検査員
'      .Cells(k, 7) = Line 'ライン
'      .Cells(k, 1).Font.Color = RGB(0, 0, 0)
'      .Cells(k, 3).Font.Color = RGB(0, 0, 0)
'      .Cells(k, 5).Font.Color = RGB(0, 0, 0)
'      .Cells(k, 6).Font.Color = RGB(0, 0, 0)
'      .Cells(k, 7).Font.Color = RGB(0, 0, 0)
'    End With
'    UF校正.Kname = ""
'    UF校正.Saiken.Visible = False
'    Call 音楽合格
'    UF校正.F1 = 0 'フラグを0にしてChangeを抜ける
'   'ChangeをExitへ変更フォーカス移動の記述を消す 2019.06.05
'   'Exit イベントは、同一フォーム上にある別のコントロールにフォーカスを移す直前に発生
'   '2019.06.18 KeyDownへ変更Exitイベントは2回処理してしまう
'   '2019.11.13　Exit イベントでendをひろってUnloadする　ここでUnloadするとオートメーションエラー
'    UF校正.F1 = "end" 'フラグ
'    UF1.ProofreadFlag = "ok" '2021.07.15
' Else
'   UF校正.NGvalue.Visible = True 'NG値の表示
'    If UF校正.INPG <> "" Then
'    UF校正.NGvalue = "測定値" & " " & UF校正.INPG & " " & "mm"
'    Else
'    UF校正.NGvalue = ""
'    End If
'    With ThisWorkbook.Sheets("校正記録")
'    If Time >= CDate("00:00") And Time <= CDate("7:00") Then
'      .Cells(k, 1) = y
'      .Cells(k, 2) = "夜勤日付処理" '夜勤処理があったかどうか
'    Else
'      .Cells(k, 1) = y
'    End If
'    .Cells(k, 3) = Format$(Now, "h時mm分ss秒") '時間
'    .Cells(k, 5) = UF校正.SokuteiH '値
'    .Cells(k, 6) = UF校正.Kname '検査員
'    .Cells(k, 7) = Line 'ライン
'    .Cells(k, 8) = "〆"   '印
'    .Cells(k, 1).Font.Color = RGB(255, 128, 128)
'    .Cells(k, 3).Font.Color = RGB(255, 128, 128)
'    .Cells(k, 5).Font.Color = RGB(255, 128, 128)
'    .Cells(k, 6).Font.Color = RGB(255, 128, 128)
'    .Cells(k, 7).Font.Color = RGB(255, 128, 128)
'    End With
'   Call 音楽不合格
'   UF校正.INPG = ""
'   UF校正.SokuteiH = ""
'   UF校正.F1 = 0 'フラグを0にしてChangeを抜ける
'   UF校正.INPG.SetFocus 'インプットボックスに初めから数字が入るようにする
'   UF校正.Saiken.Visible = True '再検査の表示
' End If
'Application.ScreenUpdating = True
End Sub
Public Sub 品証測定() '2019.02.19
'品証用　1000mm 閾値チェック有り
If WsCheck("品証検証") = False Then End
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("品証検証")
Dim i As Long, k As Long: k = ThisWorkbook.Sheets("品証検証").Cells(Rows.Count, 13).End(xlUp).row + 1
Dim Koumin As Double, Koumax As Double, line As String
'*****************************ﾗｲﾝの設定
  Select Case UF校正.LineName
  Case "L-1":    line = "L-1"
  Case "LVC":    line = "LVC"
  Case "HVC":    line = "HVC"
  Case "ﾊﾞﾗﾝｻｰ": line = "ﾊﾞﾗﾝｻｰ"
  Case "ﾄｯﾄ":    line = "ﾄｯﾄ"
  Case "Test":   line = "Test"
  Case Else:     line = "丸徳"
  End Select
'********************************************時間定義
Dim x As String:       x = TimeCheck(Format$(Now, "hh:mm")) '直を返す
If x = "日勤" Then x = "1直" '2021.06.15_大板の記録がなかった
Dim y As String:       y = YeaDayrCheck(Format$(Now, "hh:mm")) '日付を返す
Dim Formay As String:  Formay = Format$(Now, "d") & "日" '変換
'********************************************閾値設定
 '閾値変更時はランク付けも変更必要
  If UF校正.LineName = "Test" Then
    Koumin = 50 '999.5
    Koumax = 3000 '2000 '1000.5
  Else
    Koumin = 999.5
    Koumax = 1000.5
  End If
'********************************************
Application.ScreenUpdating = False
With UF校正
  If Not IsNumeric(.INP1000.Text) Then
    .INP1000.Text = "" '数字以外非表示
    .SokuteiH = .INP1000
    Exit Sub
  End If
  If val(.SokuteiH) >= val(Koumin) And val(.SokuteiH) <= val(Koumax) Then '校正の閾値 2019.04.01 =　追加
    Call 音楽合格
    .Saiken.Visible = False
    '********************************************2020.02.04 西暦を足す
'    Stop
    For i = 6 To 96
      If sh.Cells(i, 1) = Formay Then
        If sh.Cells(i, 2) = x Then
          sh.Cells(i, 3) = .SokuteiH '値
          sh.Cells(i, 9) = line 'ライン
          sh.Cells(i, 10) = .Kname '検査員
          sh.Cells(i, 11) = Format$(Now, "yyyy.mm.dd h時mm分ss秒") '時間
        ElseIf sh.Cells(i + 1, 2) = x Then
          sh.Cells(i + 1, 3) = .SokuteiH '値
          sh.Cells(i + 1, 9) = line 'ライン
          sh.Cells(i + 1, 10) = .Kname '検査員
          sh.Cells(i + 1, 11) = Format$(Now, "yyyy.mm.dd h時mm分ss秒") '時間
        ElseIf sh.Cells(i + 2, 2) = x Then
          sh.Cells(i + 2, 3) = .SokuteiH '値
          sh.Cells(i + 2, 9) = line 'ライン
          sh.Cells(i + 2, 10) = .Kname '検査員
          sh.Cells(i + 2, 11) = Format$(Now, "yyyy.mm.dd h時mm分ss秒") '時間
        End If
      End If
    Next i
    .Saiken.Visible = False
    .NGvalue = ""
    .INP1000 = ""
    .SokuteiH = ""
      Select Case UF校正.LineName
      Case "L-1":    .Gauge = "[L-1]:直尺「S-2」の全長を測定してください"
      Case "LVC":    .Gauge = "[LVC]:直尺「S-1」の全長を測定してください"
      Case "HVC":    .Gauge = "[HVC]:直尺「S-1」の全長を測定してください"
      Case "ﾊﾞﾗﾝｻｰ": .Gauge = "[大板]:直尺「S-1」の全長を測定してください"
      Case "ﾄｯﾄ":    .Gauge = "[大板]:直尺「S-1」の全長を測定してください"
      Case "Test":   .Gauge = "[Test]:直尺「xxx」の全長を測定してください"
      Case Else:     .Gauge = "[中板]:直尺「xxx」の全長を測定してください"
      End Select
    .Rank1 = ""
    .Rank2 = ""
    .Rank3 = ""
    .Rank4 = ""
    .Rank5 = ""
    .Rank6 = ""
    
    Call 校正ランク付け
    .Rank1.BackColor = &HFFFFFF '白
    .Rank2.BackColor = &HFFFFFF
    .Rank3.BackColor = &HFFFFFF
    .Rank4.BackColor = &HFFFFFF
    .Rank5.BackColor = &HFFFFFF
    .Rank6.BackColor = &HFFFFFF
    .MaxLab.BackColor = &HFFFFFF
    .MinLab.BackColor = &HFFFFFF
    .TukiLabe.BackColor = &HFFFFFF
    .BackColor = &HFFFFFF
    .Label3.BackColor = &HFFFFFF
    .Gauge.BackColor = &HFFFFFF
    .NGvalue.BackColor = &HFFFFFF
    .Saiken.BackColor = &HFFFFFF
    .SokuteiH.BackColor = &HFFFFFF
    .Ltrue.BackColor = &HFFFFFF
    .INPG.SetFocus '----------------------------------------------
    .F1 = 0 'フラグを0にしてChangeを抜ける
  Else
    Call 音楽不合格
    .SokuteiH = .INP1000
    .Saiken.Visible = True
    .NGvalue.Visible = True
    .NGvalue = "999.5 ≦" & "測定値" & "≦ 1000.5 になっていません"
    .F1 = 2
    If .INP1000 = "" Then
      .Saiken.Visible = False
      .NGvalue = ""
    End If
    sh.Cells(k, 16) = Format$(Now, "yyyy.mm.dd h時mm分ss秒") '時間
    sh.Cells(k, 13) = .SokuteiH '値
    sh.Cells(k, 15) = .Kname '検査員
    sh.Cells(k, 14) = line 'ライン
    sh.Cells(k, 13).Font.Color = RGB(255, 128, 128)
    sh.Cells(k, 14).Font.Color = RGB(255, 128, 128)
    sh.Cells(k, 15).Font.Color = RGB(255, 128, 128)
    sh.Cells(k, 16).Font.Color = RGB(255, 128, 128)
    .INP1000 = ""
    .F1 = 0 'フラグを0にしてChangeを抜ける
    .INPG.SetFocus
    .INP1000.SetFocus 'インプットボックスに初めから数字が入るようにする
  End If
End With
Application.ScreenUpdating = True
End Sub

Public Sub 校正ランク付け()
Dim MAXatu1 As Double, MAXatu2 As Double, MAXatu3 As Double, MINatu1 As Double, MINatu2 As Double, MINatu3 As Double
Dim i As Long, a As Long, b As Long, C As Long, D As Long
Dim Ran1 As Range, Ran2 As Range
Dim Tuki As String
Dim k As Long: k = ThisWorkbook.Sheets("校正記録").Cells(Rows.Count, 10).End(xlUp).row
Dim j As Long: j = ThisWorkbook.Sheets("校正記録").Cells(Rows.Count, 1).End(xlUp).row
Dim l As Long: l = ThisWorkbook.Sheets("校正記録").Cells(Rows.Count, 14).End(xlUp).row
Dim n As Long: n = 0
Dim m As Long: m = 0
  'リセット
  ThisWorkbook.Sheets("校正記録").Columns("I:O").ClearContents
  On Error Resume Next
  With ThisWorkbook.Sheets("校正記録")
    '文字列を数字に直す
    For i = 2 To j
     .Cells(i, 5) = Format(val(.Cells(i, 5)), "0.0")
     .Cells(i, 5).NumberFormatLocal = "0.0_ "
    Next i
    '月のﾁｪｯｸ 当月
    For b = j To 2 Step -1
     If Mid(.Cells(b, 1), 1, 7) = Format$(Now, "yyyy.mm") Then '2020.01.23 月のみのﾁｪｯｸから西暦.月のﾁｪｯｸに変更
      Tuki = b
     End If
    Next b
'    MsgBox Tuki & ":" & Mid(.Cells(Tuki, 1), 1, 7) '確認用
    '目的月の赤字のみを書き写す
    For C = Tuki To j
     If .Cells(C, 1).Font.Color = RGB(255, 128, 128) Then
        If UF校正.LineName = "L-1" Then '2053.2 ± 0.5
         If val(.Cells(C, 5)) >= 2053.2 Then
          n = n + 1
          .Cells(n + 1, 9) = .Cells(C, 1)
          .Cells(n + 1, 10) = .Cells(C, 5)
          .Cells(n + 1, 11) = .Cells(C, 6)
         ElseIf val(.Cells(C, 5)) < 2053.2 Then
          m = m + 1
          .Cells(m + 1, 13) = .Cells(C, 1)
          .Cells(m + 1, 14) = .Cells(C, 5)
          .Cells(m + 1, 15) = .Cells(C, 6)
         End If
        ElseIf UF校正.LineName = "LVC" Then '2000.5 ± 0.5
         If val(.Cells(C, 5)) >= 2000.5 Then
          n = n + 1
          .Cells(n + 1, 9) = .Cells(C, 1)
          .Cells(n + 1, 10) = .Cells(C, 5)
          .Cells(n + 1, 11) = .Cells(C, 6)
         ElseIf val(.Cells(C, 5)) < 2000.5 Then
          m = m + 1
          .Cells(m + 1, 13) = .Cells(C, 1)
          .Cells(m + 1, 14) = .Cells(C, 5)
          .Cells(m + 1, 15) = .Cells(C, 6)
         End If
        ElseIf UF校正.LineName = "HVC" Then '2003.8 ± 0.5'2019.01.21 閾値修正
         If val(.Cells(C, 5)) >= 2003.8 Then
          n = n + 1
          .Cells(n + 1, 9) = .Cells(C, 1)
          .Cells(n + 1, 10) = .Cells(C, 5)
          .Cells(n + 1, 11) = .Cells(C, 6)
         ElseIf val(.Cells(C, 5)) < 2003.8 Then
          m = m + 1
          .Cells(m + 1, 13) = .Cells(C, 1)
          .Cells(m + 1, 14) = .Cells(C, 5)
          .Cells(m + 1, 15) = .Cells(C, 6)
         End If
        ElseIf UF校正.LineName = "ﾊﾞﾗﾝｻｰ" Or UF校正.LineName = "ﾄｯﾄ" Then '2051.8 ± 0.5
         If val(.Cells(C, 5)) >= 2051.8 Then
          n = n + 1
          .Cells(n + 1, 9) = .Cells(C, 1)
          .Cells(n + 1, 10) = .Cells(C, 5)
          .Cells(n + 1, 11) = .Cells(C, 6)
         ElseIf val(.Cells(C, 5)) < 2051.8 Then
          m = m + 1
          .Cells(m + 1, 13) = .Cells(C, 1)
          .Cells(m + 1, 14) = .Cells(C, 5)
          .Cells(m + 1, 15) = .Cells(C, 6)
         End If
        End If
     End If
    Next C
    
    Set Ran1 = Range(.Cells(2, 10), .Cells(k, 10))
    Set Ran2 = Range(.Cells(2, 14), .Cells(l, 14))
  End With
  
  '.CountIf(範囲, xx) 重複個数を返す
  With Application.WorksheetFunction
      On Error Resume Next
      MAXatu1 = .Max((Ran1)) '最大値
      MAXatu2 = .Large((Ran1), val(.CountIf(Ran1, MAXatu1)) + 1)
      MAXatu3 = .Large((Ran1), val(.CountIf(Ran1, MAXatu1)) + 1 + val(.CountIf(Ran1, MAXatu2)) + 1)
      MINatu1 = .Min((Ran2)) '最小値
      MINatu2 = .Small((Ran2), val(.CountIf(Ran2, MINatu1)) + 1)
      MINatu3 = .Small((Ran2), val(.CountIf(Ran2, MINatu1)) + 1 + val(.CountIf(Ran2, MINatu2)) + 1)
  End With

'      UF校正.Rank1 = Val(MAXatu1)
'      UF校正.Rank2 = Val(MAXatu2)
'      UF校正.Rank3 = Val(MAXatu3)
'
'      UF校正.Rank4 = Val(MINatu1)
'      UF校正.Rank5 = Val(MINatu2)
'      UF校正.Rank6 = Val(MINatu3)

    '大小3位までをフォームに記す
  With ThisWorkbook.Sheets("校正記録")
    For a = 2 To k
    On Error Resume Next
       Select Case .Cells(a, 10)
       Case MAXatu1
        UF校正.Rank1 = "1位：　" & .Cells(a, 9) & "　値：" & Format(MAXatu1, "0.0") & "　作業者：" & .Cells(a, 11)
       Case MAXatu2
        UF校正.Rank2 = "2位：　" & .Cells(a, 9) & "　値：" & Format(MAXatu2, "0.0") & "　作業者：" & .Cells(a, 11)
       Case MAXatu3
        UF校正.Rank3 = "3位：　" & .Cells(a, 9) & "　値：" & Format(MAXatu3, "0.0") & "　作業者：" & .Cells(a, 11)
      End Select
    Next a
    
    For D = 2 To l
       Select Case .Cells(D, 14)
       Case MINatu1
        UF校正.Rank4 = "1位：　" & .Cells(D, 13) & "　値：" & Format(MINatu1, "0.0") & "　作業者：" & .Cells(D, 15)
       Case MINatu2
        UF校正.Rank5 = "2位：　" & .Cells(D, 13) & "　値：" & Format(MINatu2, "0.0") & "　作業者：" & .Cells(D, 15)
       Case MINatu3
        UF校正.Rank6 = "3位：　" & .Cells(D, 13) & "　値：" & Format(MINatu3, "0.0") & "　作業者：" & .Cells(D, 15)
      End Select
    Next D
  End With
End Sub
Public Sub 品証ランク付け()
Dim MAXatu1 As Double, MAXatu2 As Double, MAXatu3 As Double, MINatu1 As Double, MINatu2 As Double, MINatu3 As Double
Dim i As Long, a As Long, b As Long, C As Long, D As Long
Dim k As Long: k = ThisWorkbook.Sheets("品証検証").Cells(Rows.Count, 18).End(xlUp).row
Dim j As Long: j = ThisWorkbook.Sheets("品証検証").Cells(Rows.Count, 16).End(xlUp).row
Dim l As Long: l = ThisWorkbook.Sheets("品証検証").Cells(Rows.Count, 22).End(xlUp).row
Dim Ran1 As Range, Ran2 As Range
Dim Tuki As String
Dim n As Long: n = 5
Dim m As Long: m = 5
  'リセット
  ThisWorkbook.Sheets("品証検証").Columns("R:X").ClearContents
  On Error Resume Next
  With ThisWorkbook.Sheets("品証検証")
    For i = 6 To j
     If .Cells(i, 16) = "" Then
     ElseIf Mid(.Cells(i, 16), 1, 2) <> "20" Then '2020.02.04 西暦追記
     .Cells(i, 16) = "2019." & Replace(.Cells(i, 16), "/", ".") 'yyyy.mm.dd h時mm分ss秒
     End If
    Next i
    '文字列を数字に直す
    For i = 6 To j
     .Cells(i, 13) = Format(val(.Cells(i, 13)), "0.0")
    .Cells(i, 13).NumberFormatLocal = "0.0_ "
    Next i
    '月のﾁｪｯｸ 当月
    For b = j To 6 Step -1
     If Mid(.Cells(b, 16), 1, 7) = Format$(Now, "yyyy.mm") Then '2020.02.04西暦をたす
      Tuki = b
     ElseIf Mid(.Cells(b, 16), 1, 7) <> Format$(Now, "yyyy.mm") Then '最終行が当月でない(測定ミスがない)
      Exit For
     End If
    Next b
'  MsgBox Tuki '確認用
  
    '目的月のみを書き写す
    If .Cells(6, 13) = "" Then '異常測定値なし
     Exit Sub
    End If
    For C = Tuki To j
      If val(.Cells(C, 13)) >= 1000 Then '1000以上の異常値
        n = n + 1
        .Cells(n, 18) = .Cells(C, 13)
        .Cells(n, 19) = .Cells(C, 15)
        .Cells(n, 20) = .Cells(C, 16)
      ElseIf val(.Cells(C, 13)) < 1000 Then '1000未満の異常値
        m = m + 1
        .Cells(m, 22) = .Cells(C, 13)
        .Cells(m, 23) = .Cells(C, 15)
        .Cells(m, 24) = .Cells(C, 16)
      End If
    Next C
    
    Set Ran1 = Range(.Cells(6, 18), .Cells(k, 18))
    Set Ran2 = Range(.Cells(6, 22), .Cells(l, 22))
    
  End With
  
  '.CountIf(Ran1, xx) 重複個数を返す
  With Application.WorksheetFunction
      On Error Resume Next
      MAXatu1 = .Max((Ran1)) '最大値
      MAXatu2 = .Large((Ran1), val(.CountIf(Ran1, MAXatu1)) + 1)
      MAXatu3 = .Large((Ran1), val(.CountIf(Ran1, MAXatu1)) + 1 + val(.CountIf(Ran1, MAXatu2)) + 1)
      MINatu1 = .Min((Ran2)) '最小値
      MINatu2 = .Small((Ran2), val(.CountIf(Ran2, MINatu1)) + 1)
      MINatu3 = .Small((Ran2), val(.CountIf(Ran2, MINatu1)) + 1 + val(.CountIf(Ran2, MINatu2)) + 1)
      'デバック用
'      Sheets("品証検証").Range("M1:Q3").ClearContents
'      Sheets("品証検証").Cells(1, 13) = MAXatu1
'      Sheets("品証検証").Cells(2, 13) = MAXatu2
'      Sheets("品証検証").Cells(3, 13) = MAXatu3
'
'      Sheets("品証検証").Cells(1, 14) = MAXatu1
'      Sheets("品証検証").Cells(2, 14) = Val(.CountIf(Ran1, MAXatu1))
'      Sheets("品証検証").Cells(3, 14) = Val(.CountIf(Ran1, MAXatu2))
'
'      Sheets("品証検証").Cells(1, 16) = MINatu1
'      Sheets("品証検証").Cells(2, 16) = MINatu2
'      Sheets("品証検証").Cells(3, 16) = MINatu3
'
'      Sheets("品証検証").Cells(1, 17) = MINatu1
'      Sheets("品証検証").Cells(2, 17) = Val(.CountIf(Ran1, MINatu1))
'      Sheets("品証検証").Cells(3, 17) = Val(.CountIf(Ran1, MINatu2))
  End With
'      UF校正.Rank1 = MAXatu1
'      UF校正.Rank2 = MAXatu2
'      UF校正.Rank3 = MAXatu3
'
'      UF校正.Rank4 = MINatu1
'      UF校正.Rank5 = MINatu2
'      UF校正.Rank6 = MINatu3
 '大小3位までをフォームに記す
 '2019.04.03　1000以上と以下をわけて表示した
  With ThisWorkbook.Sheets("品証検証")
    For a = 6 To k
    On Error Resume Next
       Select Case .Cells(a, 18)
       Case MAXatu1
        UF校正.Rank1 = "1位：　" & .Cells(a, 20) & "　値：" & Format(MAXatu1, "0.0") & "　作業者：" & .Cells(a, 19)
       Case MAXatu2
        UF校正.Rank2 = "2位：　" & .Cells(a, 20) & "　値：" & Format(MAXatu2, "0.0") & "　作業者：" & .Cells(a, 19)
       Case MAXatu3
        UF校正.Rank3 = "3位：　" & .Cells(a, 20) & "　値：" & Format(MAXatu3, "0.0") & "　作業者：" & .Cells(a, 19)
      End Select
    Next a
    
    For D = 6 To l
       Select Case .Cells(D, 22)
       Case MINatu1
        UF校正.Rank4 = "1位：　" & .Cells(D, 24) & "　値：" & Format(MINatu1, "0.0") & "　作業者：" & .Cells(D, 23)
       Case MINatu2
        UF校正.Rank5 = "2位：　" & .Cells(D, 24) & "　値：" & Format(MINatu2, "0.0") & "　作業者：" & .Cells(D, 23)
       Case MINatu3
        UF校正.Rank6 = "3位：　" & .Cells(D, 24) & "　値：" & Format(MINatu3, "0.0") & "　作業者：" & .Cells(D, 23)
      End Select
    Next D
  End With
End Sub

Public Sub Proofread_Load() '未校正の時、最終登録直、作業者のチェック後校正起動_2021.07.15
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If (UF1.WorkButton.Tag = "jump") Or (UF1.DBflag <> "") Or UF1.ProofreadFlag = "ok" Then '校正停止Flag
  Exit Sub
End If
Dim Workers
Select Case WorkeDayMax '登録時間が一番遅い
Case "1": If UF1.T1.Text <> "" Then Workers = Split(UF1.T1.Text, "_")
Case "2": If UF1.T2.Text <> "" Then Workers = Split(UF1.T2.Text, "_")
Case "3": If UF1.T3.Text <> "" Then Workers = Split(UF1.T3.Text, "_")
End Select
Dim EndWork As String: EndWork = ""
If IsEmpty(Workers) Then EndWork = ""
If IsEmpty(Workers) = False Then EndWork = Workers(UBound(Workers))
MsgBox "校正が済んでいません。校正を起動します。", vbCritical
Call 校正_Open(EndWork, Lname)
End Sub
