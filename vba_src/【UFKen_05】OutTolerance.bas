Attribute VB_Name = "【UFKen_05】OutTolerance"
Option Explicit
Public Sub SmallNumberCount(TB As Long)
'2024.9.26_ラテラルボー検査不要の時描写停止を追加
Dim NgTemp As String: NgTemp = Right(UF1.Additional.Caption, 4)
'少枚数設定後に検査回数で表示を変える
Dim i As Long
  For i = 11 To 14 '耳→少にすると描写が残る
    UF1("S" & i).Visible = False
    UF1("U" & i).Visible = False
    UF1("C" & i).Visible = False
    UF1("D" & i).Visible = False
    UF1("E" & i).Visible = False
    UF1("S" & i).Enabled = False
    UF1("U" & i).Enabled = False
    UF1("C" & i).Enabled = False
    UF1("D" & i).Enabled = False
    UF1("E" & i).Enabled = False
  Next i

'少枚数かつ枚数入力かつ検番頭H
'ﾘｾｯﾄ5回用を表示<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
With UF1
  .S = "1"
  .U = "2"
  .C = "3"
  .D = "4"
  .E = "5"
  .FS = "1"
  .FU = "2"
  .FC = "3"
  .FD = "4"
  .FE = "5"
  .OPS1.Caption = "1"
  .OPS2.Caption = "2"
  .OPS3.Caption = "3"
  .OPS4.Caption = "4"
  .OPS5.Caption = "5"
  .OPE1.Caption = "1"
  .OPE2.Caption = "2"
  .OPE3.Caption = "3"
  .OPE4.Caption = "4"
  .OPE5.Caption = "5"

'歪
  .FMS.Visible = True
  .FMU.Visible = True
  .FMC.Visible = True
  .FMD.Visible = True
  .FME.Visible = True
'外観
  For i = 1 To 5
    UF1("LC" & i).Visible = True
  Next i
  For i = 7 To 11
    UF1("LC" & i).Visible = True
  Next i

  .S_UpVisual.Visible = True
  .U_UpVisual.Visible = True
  .C_UpVisual.Visible = True
  .D_UpVisual.Visible = True
  .E_UpVisual.Visible = True
  .S_UnderVisual.Visible = True
  .U_UnderVisual.Visible = True
  .C_UnderVisual.Visible = True
  .D_UnderVisual.Visible = True
  .E_UnderVisual.Visible = True
'厚幅丈
  For i = 1 To 10
    UF1("S" & i).Visible = True
    UF1("U" & i).Visible = True
    UF1("C" & i).Visible = True
    UF1("D" & i).Visible = True
    UF1("E" & i).Visible = True
    UF1("S" & i).Enabled = True
    UF1("U" & i).Enabled = True
    UF1("C" & i).Enabled = True
    UF1("D" & i).Enabled = True
    UF1("E" & i).Enabled = True
  Next i
'対角
    .DiagS.Visible = True
    .DiagS.Enabled = True
    .DiagU.Visible = True
    .DiagU.Enabled = True
    .DiagC.Visible = True
    .DiagC.Enabled = True
    .DiagD.Visible = True
    .DiagD.Enabled = True
    .DiagE.Visible = True
    .DiagE.Enabled = True
    '**********2023.11.24_対角公差から直角度に変更
'    .DiagSKou.Visible = True
'    .DiagSKou.Enabled = True
'    .DiagUKou.Visible = True
'    .DiagUKou.Enabled = True
'    .DiagCKou.Visible = True
'    .DiagCKou.Enabled = True
'    .DiagDKou.Visible = True
'    .DiagDKou.Enabled = True
'    .DiagEKou.Visible = True
'    .DiagEKou.Enabled = True
'BH
  .SBH.Visible = True
  .UBH.Visible = True
  .CBH.Visible = True
  .DBH.Visible = True
  .EBH.Visible = True
'チェック用
.TSS.Visible = True
.TSU.Visible = True
.TSC.Visible = True
.TSD.Visible = True
.TSE.Visible = True
.TES.Visible = True
.TEU.Visible = True
.TEC.Visible = True
.TED.Visible = True
.TEE.Visible = True
'位置調整 ｻｲｽﾞ
    Dim LeftTemp As Long: LeftTemp = 804
    Dim LeftTemp2 As Long: LeftTemp2 = 852
    Dim LeftOut As Long: LeftOut = 906
    Dim LeftOut2 As Long: LeftOut2 = 972
    .Width = 895
    .L9.Left = LeftTemp
    .L10.Left = LeftTemp2
    .DiagS.Left = LeftTemp
    .DiagU.Left = LeftTemp
    .DiagC.Left = LeftTemp
    .DiagD.Left = LeftTemp
    .DiagE.Left = LeftTemp
    .DiagSKou.Left = LeftTemp2
    .DiagUKou.Left = LeftTemp2
    .DiagCKou.Left = LeftTemp2
    .DiagDKou.Left = LeftTemp2
    .DiagEKou.Left = LeftTemp2
    .L7.Left = LeftOut
    .L8.Left = LeftOut2
    For i = 11 To 12
    Select Case i
    Case 11
      UF1("S" & i).Left = LeftOut
      UF1("U" & i).Left = LeftOut
      UF1("C" & i).Left = LeftOut
      UF1("D" & i).Left = LeftOut
      UF1("E" & i).Left = LeftOut
    Case 12
      UF1("S" & i).Left = LeftOut2
      UF1("U" & i).Left = LeftOut2
      UF1("C" & i).Left = LeftOut2
      UF1("D" & i).Left = LeftOut2
      UF1("E" & i).Left = LeftOut2
    End Select
    Next i
    
    Dim TempArry As Variant: TempArry = GetLotInfo("Lot情報履歴", UF1.Lot.Text) 'ヒットしたロットの配列を作成
    If IsEmpty(TempArry) Then
      GoTo skipEnd
      Exit Sub
    End If
    Dim TempDelivery As String
    
    '納入、取引、用途コード、材質、調質、厚、幅、丈、ロットNo、検査No、ｵｰﾀﾞｰNo、包装仕様NO
    Dim 納入 As Variant, 取引 As Variant, 用途コード As String, 材質 As String, 調質 As String
    Dim 厚 As Single, 幅 As Single, 丈 As Single, 検番 As String
    Dim ArrCount As Long: ArrCount = UBound(TempArry, 1)
    ReDim 納入(1 To ArrCount) As Variant
    ReDim 取引(1 To ArrCount) As Variant  '2024.8.5
    Dim j As Long: j = 0
      For i = LBound(TempArry, 1) To UBound(TempArry, 1) '引当数分チェック....1回でよくない？ ***あかん！滑川は複数引き当ての時に特殊検査のない引き当てと混ざってくる_2024.4.22
        j = j + 1
        納入(j) = TempArry(i, 1)
        取引(j) = TempArry(i, 2)
        用途コード = TempArry(i, 3)
        材質 = TempArry(i, 4)
        調質 = TempArry(i, 5)
        厚 = TempArry(i, 6)
        幅 = TempArry(i, 7)
        丈 = TempArry(i, 8)
        検番 = TempArry(i, 10) '2026.05.08_参照がずっとLotNoだったTempArry(i, 9)　検番参照に修正TempArry(i, 10)
      Next i
      
    検番 = Left(検番, 1)
      
    Dim 納入Check As String, 取引Check As String
    Dim 前回納入Check As String: 前回納入Check = ""
    For i = LBound(納入, 1) To UBound(納入, 1)
      If 納入(i) <> "" Then
        納入Check = 納入(i)
        取引Check = 取引(i)
        
        '納入先が前回と同じなら DbReturnItem をスキップ
        If 納入Check = 前回納入Check Then GoTo NextLoop
        前回納入Check = 納入Check
        
        TempDelivery = DbReturnItem(納入Check, 取引Check, 用途コード, 材質, 調質, 厚, 幅, 丈)
        Dim ラテflag As Boolean, 直角flag As Boolean
        If InStr(TempDelivery, "ラテラルボー") > 0 Then ラテflag = True
        If InStr(TempDelivery, "直角度") > 0 Then 直角flag = True
      End If
NextLoop:
    Next i
    
skipEnd:
    '.Sumisei.Caption <> ""をフラグと変更_2024.11.15
    If ラテflag And NgTemp <> "検査不要" Then     'ﾗﾃﾗﾙﾎﾞｰ有　 And NgTemp <> "検査不要"
      .L11.Visible = True
      .LateralS.Visible = True
      .LateralU.Visible = True
      .LateralC.Visible = True
      .LateralD.Visible = True
      .LateralE.Visible = True
      
      .Width = 987.75
      .L11.Left = LeftOut
      .LateralS.Left = LeftOut
      .LateralU.Left = LeftOut
      .LateralC.Left = LeftOut
      .LateralD.Left = LeftOut
      .LateralE.Left = LeftOut
      .L11.tOP = 366
      .LateralS.tOP = 390
      .LateralU.tOP = 415.55
      .LateralC.tOP = 441.05
      .LateralD.tOP = 466.55
      .LateralE.tOP = 492.05
    End If
    
    If UF1.Youto = "R115" Then 'QA-S-QA-DIM-154_対角線差及び直角度の測定基準
      UF1.L10.Visible = True
      UF1.DiagSKou.Visible = True
      UF1.DiagUKou.Visible = True
      UF1.DiagCKou.Visible = True
      UF1.DiagDKou.Visible = True
      UF1.DiagEKou.Visible = True
      
      UF1.L10.Left = LeftTemp2
      UF1.DiagSKou.Left = LeftTemp2
      UF1.DiagUKou.Left = LeftTemp2
      UF1.DiagCKou.Left = LeftTemp2
      UF1.DiagDKou.Left = LeftTemp2
      UF1.DiagEKou.Left = LeftTemp2
    End If
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'Q-S-CH-DIM-163_板材抜取検査頻度.寸法測定位置基準_2025.7*ラテラルボー検査回数変更
Select Case TB
Case 5
  If ラテflag And NgTemp <> "検査不要" Then  'ﾗﾃﾗﾙﾎﾞｰ測定の納入先が追加される為"住友精密"から<>に変更
    .LateralS.Visible = True
    .LateralU.Visible = False '2025.7
    .LateralC.Visible = True
    .LateralD.Visible = False '2025.7
    .LateralE.Visible = True
  End If
  
  For i = 5 To 10
    UF1("U" & i).Visible = False
    UF1("D" & i).Visible = False
    UF1("U" & i).Enabled = False
    UF1("D" & i).Enabled = False
  Next i
  
  If UF1.Youto = "R115" Then 'QA-S-QA-DIM-154_対角線差及び直角度の測定基準
    UF1.L10.Visible = True
    UF1.DiagSKou.Visible = True
    UF1.DiagUKou.Visible = True
    UF1.DiagCKou.Visible = True
    UF1.DiagDKou.Visible = True
    UF1.DiagEKou.Visible = True
  End If
  
    .DiagU.Visible = False
    .DiagU.Enabled = False
    .DiagD.Visible = False
    .DiagD.Enabled = False
Case 4
  .E = ""
  .FE = ""
  .OPS5.Caption = ""
  .OPE5.Caption = ""
  .FME.Visible = False
  UF1("LC" & 5).Visible = False
  UF1("LC" & 11).Visible = False
  .E_UpVisual.Visible = False
  .E_UnderVisual.Visible = False
  For i = 1 To 10
    UF1("E" & i).Visible = False
    UF1("E" & i).Enabled = False
  Next i
  For i = 5 To 10
    UF1("U" & i).Visible = False
    UF1("U" & i).Enabled = False
  Next i
  .DiagU.Enabled = False
  .DiagU.Visible = False
  .DiagE.Enabled = False
  .DiagE.Visible = False
  
  If UF1.Youto = "R115" Then
    .DiagEKou.Enabled = False
    .DiagEKou.Visible = False
  End If
  
  .EBH.Visible = False
  .TSE.Visible = False
  .TEE.Visible = False
    If ラテflag And NgTemp <> "検査不要" Then    'ﾗﾃﾗﾙﾎﾞｰ有
      .LateralU.Visible = False '2025.7
      .LateralE.Visible = False
    End If
Case 3
  .D = ""
  .E = ""
  .FD = ""
  .FE = ""
  .OPS4.Caption = ""
  .OPS5.Caption = ""
  .OPE4.Caption = ""
  .OPE5.Caption = ""
  .FMD.Visible = False
  .FME.Visible = False
  UF1("LC" & 5).Visible = False
  UF1("LC" & 11).Visible = False
  UF1("LC" & 4).Visible = False
  UF1("LC" & 10).Visible = False
  .D_UpVisual.Visible = False
  .E_UpVisual.Visible = False
  .D_UnderVisual.Visible = False
  .E_UnderVisual.Visible = False
  For i = 1 To 10
    UF1("D" & i).Visible = False
    UF1("E" & i).Visible = False
    UF1("D" & i).Enabled = False
    UF1("E" & i).Enabled = False
  Next i
  .DiagD.Enabled = False
  .DiagD.Visible = False
  
  If UF1.Youto = "R115" Then
    .DiagDKou.Enabled = False
    .DiagDKou.Visible = False
    .DiagEKou.Enabled = False
    .DiagEKou.Visible = False
  End If
  
  .DiagE.Enabled = False
  .DiagE.Visible = False
  .DBH.Visible = False
  .EBH.Visible = False
  .TSD.Visible = False
  .TSE.Visible = False
  .TED.Visible = False
  .TEE.Visible = False
    If ラテflag And NgTemp <> "検査不要" Then  'ﾗﾃﾗﾙﾎﾞｰ有
      .LateralD.Visible = False
      .LateralE.Visible = False
    End If
Case 2
  .C = ""
  .D = ""
  .E = ""
  .FC = ""
  .FD = ""
  .FE = ""
  .OPS3.Caption = ""
  .OPS4.Caption = ""
  .OPS5.Caption = ""
  .OPE3.Caption = ""
  .OPE4.Caption = ""
  .OPE5.Caption = ""
  .FMC.Visible = False
  .FMD.Visible = False
  .FME.Visible = False
  UF1("LC" & 5).Visible = False
  UF1("LC" & 11).Visible = False
  UF1("LC" & 4).Visible = False
  UF1("LC" & 10).Visible = False
  UF1("LC" & 3).Visible = False
  UF1("LC" & 9).Visible = False
  .C_UpVisual.Visible = False
  .D_UpVisual.Visible = False
  .E_UpVisual.Visible = False
  .C_UnderVisual.Visible = False
  .D_UnderVisual.Visible = False
  .E_UnderVisual.Visible = False
  For i = 1 To 10
    UF1("C" & i).Visible = False
    UF1("D" & i).Visible = False
    UF1("E" & i).Visible = False
    UF1("C" & i).Enabled = False
    UF1("D" & i).Enabled = False
    UF1("E" & i).Enabled = False
  Next i
  .DiagC.Enabled = False
  .DiagC.Visible = False
  
  If UF1.Youto = "R115" Then
    .DiagCKou.Enabled = False
    .DiagCKou.Visible = False
    .DiagDKou.Enabled = False
    .DiagDKou.Visible = False
    .DiagEKou.Enabled = False
    .DiagEKou.Visible = False
  End If
  
  .DiagD.Enabled = False
  .DiagD.Visible = False
  .DiagE.Enabled = False
  .DiagE.Visible = False
  .CBH.Visible = False
  .DBH.Visible = False
  .EBH.Visible = False
  .TSC.Visible = False
  .TSD.Visible = False
  .TSE.Visible = False
  .TEC.Visible = False
  .TED.Visible = False
  .TEE.Visible = False
    If ラテflag And NgTemp <> "検査不要" Then  'ﾗﾃﾗﾙﾎﾞｰ有
      .LateralC.Visible = False
      .LateralD.Visible = False
      .LateralE.Visible = False
    End If
Case 1
  .U = ""
  .C = ""
  .D = ""
  .E = ""
  .FU = ""
  .FC = ""
  .FD = ""
  .FE = ""
  .OPS2.Caption = ""
  .OPS3.Caption = ""
  .OPS4.Caption = ""
  .OPS5.Caption = ""
  .OPE2.Caption = ""
  .OPE3.Caption = ""
  .OPE4.Caption = ""
  .OPE5.Caption = ""
  .FMU.Visible = False
  .FMC.Visible = False
  .FMD.Visible = False
  .FME.Visible = False
  UF1("LC" & 5).Visible = False
  UF1("LC" & 11).Visible = False
  UF1("LC" & 4).Visible = False
  UF1("LC" & 10).Visible = False
  UF1("LC" & 3).Visible = False
  UF1("LC" & 9).Visible = False
  UF1("LC" & 2).Visible = False
  UF1("LC" & 8).Visible = False
  .U_UpVisual.Visible = False
  .C_UpVisual.Visible = False
  .D_UpVisual.Visible = False
  .E_UpVisual.Visible = False
  .U_UnderVisual.Visible = False
  .C_UnderVisual.Visible = False
  .D_UnderVisual.Visible = False
  .E_UnderVisual.Visible = False
  For i = 1 To 10
    UF1("U" & i).Visible = False
    UF1("C" & i).Visible = False
    UF1("D" & i).Visible = False
    UF1("E" & i).Visible = False
    UF1("U" & i).Enabled = False
    UF1("C" & i).Enabled = False
    UF1("D" & i).Enabled = False
    UF1("E" & i).Enabled = False
  Next i
  If UF1.Youto = "R115" Then
    .DiagUKou.Enabled = False
    .DiagUKou.Visible = False
    .DiagCKou.Enabled = False
    .DiagCKou.Visible = False
    .DiagDKou.Enabled = False
    .DiagDKou.Visible = False
    .DiagEKou.Enabled = False
    .DiagEKou.Visible = False
  End If
  
  .DiagU.Enabled = False
  .DiagU.Visible = False
  .DiagC.Enabled = False
  .DiagC.Visible = False
  .DiagD.Enabled = False
  .DiagD.Visible = False
  .DiagE.Enabled = False
  .DiagE.Visible = False
  .CBH.Visible = False
  .DBH.Visible = False
  .EBH.Visible = False
  .UBH.Visible = False
  .CBH.Visible = False
  .DBH.Visible = False
  .EBH.Visible = False
  .TSU.Visible = False
  .TSC.Visible = False
  .TSD.Visible = False
  .TSE.Visible = False
  .TEU.Visible = False
  .TEC.Visible = False
  .TED.Visible = False
  .TEE.Visible = False
    If ラテflag And NgTemp <> "検査不要" Then     'ﾗﾃﾗﾙﾎﾞｰ有　　<> "" And NgTemp <> "検査不要"
      .LateralU.Visible = False
      .LateralC.Visible = False
      .LateralD.Visible = False
      .LateralE.Visible = False
    End If
End Select
End With

Dim Temp As Variant:  Temp = InitialSetting(UF1)

End Sub
Public Sub UFPack_Open(n As String, PcLot As String) 'UFPackﾌｫｰﾑ
If UFProgress.Visible = True Then Unload UFProgress
Load UF_PackNo
'”\”忘れないこと
Dim PackPath As String: PackPath = QualityPath & "包装仕様書\" '"\\nlmsrvngy03\工場内共有\検査データ\包装仕様書\" '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\検査\包装仕様\"
Dim PackPic As String:  PackPic = n & ".JPG"
If ReadCheck((PackPath & PackPic), 1) = True Then 'ﾈｯﾄﾜｰｸにつながっていないor登録がない

UF1.Hide '最前面に来ない場合がある対策_2023.10.10

  Dim pic As Object: Set pic = LoadPicture(PackPath & PackPic)
  UF_PackNo.PacIM.Picture = pic
  With UF_PackNo.FMBar
    .ScrollHeight = UF_PackNo.PacIM.Height
    .ScrollWidth = UF_PackNo.PacIM.Width
    .Picture = pic
  End With
  UF_PackNo.PcLot = PcLot
  UF_PackNo.TBLoad = n
  
  UF_PackNo.Show vbModal
  
End If
End Sub

Public Sub NameList_Open(n As Long) 'UFNameListﾌｫｰﾑ
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim sh As Worksheet, Temp1() As Variant, TemP2() As Variant
If UFProgress.Visible = True Then Unload UFProgress
Select Case n
Case 1 '検査員不一致
  Select Case True
  Case UF1.L1Work:  Call WriteLog("*UFKen:" & "UFNameList起動" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "L-1\L-1_Log")
  Case UF1.LVC:     Call WriteLog("*UFKen:" & "UFNameList起動" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "LVC\LVC_Log")
  Case UF1.HVC:     Call WriteLog("*UFKen:" & "UFNameList起動" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "HVC\HVC_Log")
  Case UF1.Bala:    Call WriteLog("*UFKen:" & "UFNameList起動" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "バランサー\バランサー_Log")
  Case UF1.Tott:    Call WriteLog("*UFKen:" & "UFNameList起動" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "トット\トット_Log")
  Case UF1.Maru:    Call WriteLog("*UFKen:" & "UFNameList起動" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "丸徳\丸徳_Log")
  Case Else:        Call WriteLog("*UFKen:" & "UFNameList起動" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "Test\Test_Log")
  End Select
  If WsCheck("WorkerTime" & "_" & Lname) = False Then Exit Sub '-------------------------ｼｰﾄなし
  UFNameList.OnOff.Visible = True
  UFNameList.OnOff.Caption = "VBA停止"
  Load UFNameList
  UFNameList.FM3.Visible = False '特殊操作ﾎﾞﾀﾝ
  UFNameList.Content = "検査員不一致"
  On Error Resume Next
  Set sh = ThisWorkbook.Sheets("WorkerTime" & "_" & Lname)
  Temp1 = sh.Range(sh.Cells(2, 1), sh.Cells(2, 12))
  TemP2 = sh.Range(sh.Cells(3, 1), sh.Cells(12, 12)) 'UR(SH, 1):最終行 ﾊﾟｹｯﾄﾃﾞｰﾀを除くため最終行は使わない
  Call CreateField(UFNameList.LV1, Temp1)
  Call CreateList(UFNameList.LV1, TemP2)
  Call WindowMenuButtonDel(UFNameList)
  DoEvents
  UF1.Hide
  UFNameList.Show vbModal
Case 2 '日報
  Load UFNameList
  UFNameList.Content = "日報閲覧"
  UFNameList.CB1.Visible = False
  UFNameList.OnOff.Visible = False
  UFNameList.Label1.Visible = False
  UFNameList.Label2.Visible = False
  UFNameList.Label3.Visible = False
  Set sh = ThisWorkbook.Sheets("Daily" & "_" & Lname)
  Temp1 = sh.Range(sh.Cells(2, 1), sh.Cells(2, 25))
  TemP2 = sh.Range(sh.Cells(3, 1), sh.Cells(UR(sh, 6), 25))
  
'  Call CreateAll(UFNameList.LV1, Temp1, Temp2)
  
  Call CreateField(UFNameList.LV1, Temp1)
  Call CreateList(UFNameList.LV1, TemP2)
  Dim i As Long, x As Long, y As Long
  x = UFNameList.LV1.ColumnHeaders.Item(4).Width
  UFNameList.LV1.ColumnHeaders.Item(4).Width = x + 70
  y = UFNameList.LV1.ColumnHeaders.Item(15).Width
  UFNameList.LV1.ColumnHeaders.Item(15).Width = x + 70
  For i = 16 To 25 '2からｽﾀｰﾄ
   UFNameList.LV1.ColumnHeaders.Item(i).Width = 0
  Next i
  DoEvents
  UFNameList.Show vbModal
Case 3 '梱包枚数
  Load UFNameList
  UFNameList.Content = "梱包枚数"
  UFNameList.OnOff.Visible = True
  UFNameList.OnOff.Caption = "編集起動"
  UFNameList.FM3.Visible = False '特殊操作ﾎﾞﾀﾝ
  UFNameList.CB1.Visible = False
  UFNameList.Label1.Visible = False
  UFNameList.Label2.Visible = False
  UFNameList.Label3.Visible = False
  Set sh = ThisWorkbook.Sheets("梱包枚数計算" & "_" & Lname)
  Temp1 = sh.Range(sh.Cells(1, 1), sh.Cells(1, UC(sh, 1)))
  Dim Erow As Long: Erow = UR(sh, 1)
  TemP2 = sh.Range(sh.Cells(2, 1), sh.Cells(Erow, 4)) '最終行まで表示させないと初めのLotしか表示しない_2021.05.07
  Call CreateField(UFNameList.LV1, Temp1)
  Call CreateList(UFNameList.LV1, TemP2)
  UFNameList.Width = 440
  x = UFNameList.LV1.ColumnHeaders.Item(2).Width
  UFNameList.LV1.ColumnHeaders.Item(2).Width = x + 70
  y = UFNameList.LV1.ColumnHeaders.Item(4).Width
  UFNameList.LV1.ColumnHeaders.Item(4).Width = x + 70
  DoEvents
  UFNameList.Show vbModal
End Select
End Sub
Public Sub WorkerName_Open() '検査員確認
UFWorkerName.Show vbModeless
UFWorkerName.tOP = 0
UFWorkerName.Left = 0
If Not UF1.Visible = True Then Exit Sub
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim Worker
Select Case WorkeDayMax
Case "1": Worker = Split(UF1.T1, "_") '複数人の登録があった場合後の登録者を記録
Case "2": Worker = Split(UF1.T2, "_")
Case "3": Worker = Split(UF1.T3, "_")
End Select
If IsEmpty(Worker) Then UFWorkerName.L1 = "           "
If IsEmpty(Worker) = False Then UFWorkerName.L1 = Worker(UBound(Worker))
If UF1.Visible Then UFWorkerName.L2 = Lname '2021.10.25
'StrConv関数で文字列をいったんANSI形式に変換し、変換後の文字列をもとにLenB関数でバイト数を取得
UFWorkerName.L1.Width = LenB(StrConv(UFWorkerName.L1.Caption, vbFromUnicode)) * 25
UFWorkerName.Width = 16 + (LenB(StrConv(UFWorkerName.L1.Caption, vbFromUnicode)) * 25) + 35
End Sub
Public Sub ConfiReason_Open(Posi As String, n As Long, Lname As String, Moto As String, Optional delCount As Long = 0) '確認ﾌｫｰﾑ**2023.4.21_引数追加
Call UFsizeSave(UF1) '2024.8.5
'NG判定以外を消した場合に出現
UFConfi.Kpop = Posi
UFConfi.Knum = n
UFConfi.LineName = Lname
UFConfi.MotoData.Caption = Moto
UFConfi.DelC.Caption = delCount

Call WindowMenuButtonDel(UFConfi)
DoEvents
UF1.Hide
'2023.4.19_マルチ追加
UFConfi.MP1.Value = 1
UFConfi.MP1.P1.Visible = False

Dim Worker
Select Case WorkeDayMax
Case "1": Worker = Split(UF1.T1, "_") '複数人の登録があった場合後の登録者を記録
Case "2": Worker = Split(UF1.T2, "_")
Case "3": Worker = Split(UF1.T3, "_")
End Select

If IsEmpty(Worker) = False Then UFConfi.KenName = Worker(UBound(Worker))

UFConfi.Show vbModeless
End Sub
Public Sub Confi_Open(Posi As String, n As Long, Lname As String, Moto As String, Optional delCount As Long = 0) '確認ﾌｫｰﾑ**2023.4.21_引数追加
Call UFsizeSave(UF1) '2024.8.5

'NG判定うけた場合に指定の座標で出現_Out_Tolerance
'n=96,97,98,99 *ﾗﾃﾗﾙﾎﾞｰ、直角度、定歪、ｹﾞｰｼﾞ歪
UFConfi.Kpop = Posi
UFConfi.Knum = n
UFConfi.LineName = Lname
UFConfi.MotoData.Caption = Moto
UFConfi.return_MotoData.Caption = Moto '2023.9.28追加

UFConfi.DelC.Caption = delCount

Call WindowMenuButtonDel(UFConfi)

'公差外が対角線差の場合差の数値を出力する*2023.9.28
If UF1.厚板 Or UF1.薄板 Or UF1.少枚数 Then 'スラブ、極厚以外の検査
Dim Temp As String: Temp = "Diag" & Posi
'9,10が対角線
  If n = 9 Or n = 10 Then
    UFConfi.MotoData.Caption = "線差 " & UF1(Temp).Caption
    UFConfi.return_MotoData.Caption = "線差 " & UF1(Temp).Caption
  End If
End If


DoEvents
UF1.Hide
'2023.4.19_マルチ追加
UFConfi.MP1.Value = 0
UFConfi.MP1.P2.Visible = False

Dim Worker
Select Case WorkeDayMax
Case "1": Worker = Split(UF1.T1, "_") '複数人の登録があった場合後の登録者を記録
Case "2": Worker = Split(UF1.T2, "_")
Case "3": Worker = Split(UF1.T3, "_")
End Select

If IsEmpty(Worker) = False Then UFConfi.KenName = Worker(UBound(Worker))

UFConfi.Show vbModeless
End Sub
Public Sub MsgError(Trm As String, Lname As String, Optional n As Long = 0, Optional Temp As String = "")
Dim i As Long
'警告文,ﾗｲﾝ名,ｱｲｺﾝ,end処理　*temp="END"でend処理*temp="Ainput"でUF1をhide
Select Case n
Case 1: UFMsgBox.IM1.Visible = True '!
Case 2: UFMsgBox.IM2.Visible = True '×
Case 3: UFMsgBox.IM3.Visible = True '？
Case 4: UFMsgBox.IM4.Visible = True 'i
End Select

UFMsgBox.Caption = "警告"
If n = 4 Then UFMsgBox.Caption = "注意"

UFMsgBox.Tag = Temp
UFMsgBox.FlLabe = UFMsgBox.Tag

Dim linesCount As Long: linesCount = Len(Trm) - Len(Replace(Trm, ",", "")) + 1 ' 改行の数をカウント
Dim lineHeight As Long: lineHeight = 24 ' 1行あたりの高さを設定
Dim formattedText As String
Dim TempArr() As String: TempArr = Split(Trm, ",") ' コンマで文字列を分割
Dim maxLength As Long, textLength As Long
Dim maxElement As String

' 都度文字数をカウントし、最大文字数の要素を記録
For i = LBound(TempArr) To UBound(TempArr)
  textLength = Len(Trim(TempArr(i))) ' 空白を除いた文字数を取得
  
  If textLength > maxLength Then
      maxLength = textLength ' 最大文字数の更新
      maxElement = TempArr(i) ' 最大文字数の要素を記録
  End If
Next i

If linesCount = 1 Then
UFMsgBox.L1 = Trm
GoTo skipLno '改行なし
End If

UFMsgBox.L1.Height = linesCount * lineHeight ' ラベルの高さを調整

' 都度改行して表示
For i = LBound(TempArr) To UBound(TempArr)
  If Trim(TempArr(i)) <> "" Then
    formattedText = formattedText & TempArr(i) & vbCrLf
  End If
Next i

formattedText = Left(formattedText, Len(formattedText) - Len(vbCrLf)) ' 最後の改行を削除
UFMsgBox.L1 = formattedText ' ラベルのキャプションに設定
UFMsgBox.Height = UFMsgBox.L1.Height + 100 'フォームサイズ修正、適当に+
UFMsgBox.CommandButton1.tOP = 33 * linesCount - 5 'ボタンの高さを調整　改行数にボタン初期Top値付近を掛け算

skipLno:

'StrConv関数で文字列をいったんANSI形式に変換し、変換後の文字列をもとにLenB関数でバイト数を取得
UFMsgBox.L1.Width = LenB(StrConv(maxElement, vbFromUnicode)) * 8 + 120
UFMsgBox.Width = 25 + (LenB(StrConv(maxElement, vbFromUnicode)) * 8) + 120
UFMsgBox.CommandButton1.Left = ((16 + (LenB(StrConv(maxElement, vbFromUnicode)) * 8) + 75) / 2) - 30

'UFMsgBox.L1.width = 16 + (Len(Trm)) * 15  '微調整 記号がある場合広くなる
'UFMsgBox.Width = 16 + ((Len(Trm)) * 15) + 35 '両サイド分
Call WindowMenuButtonDel(UFMsgBox)

UF1.Repaint
Load UFMsgBox
UFMsgBox.Show vbModeless

Dim path As String: path = HistoryPath
Select Case Lname
Case "L-1":     Call WriteLog("*UFKen:" & "MsgError:" & Trm & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", path & "L-1\L-1_Log")
Case "LVC":     Call WriteLog("*UFKen:" & "MsgError:" & Trm & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", path & "LVC\LVC_Log")
Case "HVC":     Call WriteLog("*UFKen:" & "MsgError:" & Trm & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", path & "HVC\HVC_Log")
Case "ﾊﾞﾗﾝｻｰ":  Call WriteLog("*UFKen:" & "MsgError:" & Trm & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", path & "バランサー\バランサー_Log")
Case "ﾄｯﾄ":     Call WriteLog("*UFKen:" & "MsgError:" & Trm & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", path & "トット\トット_Log")
Case "Test":    Call WriteLog("*UFKen:" & "MsgError:" & Trm & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", path & "Test\Test_Log")
Case Else:      Call WriteLog("*UFKen:" & "MsgError:" & Trm & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", path & "丸徳\丸徳_Log")
End Select
UF1.Hide
End Sub

Public Sub AbnoFormStart(n As String, TB As String) '異常値ﾌｫｰﾑ
UF1.Abno.Tag = "NG"
DoEvents

Load UF異常値発見
Call WindowMenuButtonDel(UF異常値発見)
UF異常値発見.IJO = n
UF異常値発見.TBname = TB
UF異常値発見.Show vbModal 'vbModeless '
End Sub
'異常値除去で使用
Public Function Era_flag(TB As MSForms.Label, KP As Double, KM As Double) As Boolean
  If (val(TB) < KM) Or (val(TB) > KP) Then
    Era_flag = True
  End If
End Function
'重量計算
Public Function Weight(Thickness, Width, length) As Double
If (IsNumeric(Thickness) = False Or IsNumeric(Width) = False Or IsNumeric(length) = False) Then Exit Function
If (Thickness = "" Or Width = "" Or length = "") Then Exit Function
Weight = Application.RoundDown((Thickness * Width * length * 2.68 / 10 ^ 6), 3) '下3切り捨て
End Function
Public Sub Out_Tolerance(Optional Temp As String = "") '公差外検索
If UF1.INP.Tag = "" Then Exit Sub
Dim Scol As Long: Scol = 13434828 '合格
Dim Ncol As Long: Ncol = 8421631

Dim st As String:  st = UF1.INP.Tag
If Temp <> "" Then st = Temp

Dim i As Long
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption

Dim TempAlp As String

'2024.4.17_右から抜いた文字列が数字の場合は不要なので左から抜きなおす　欲しいのはSUCDEF
If Len(st) = 2 Then
  TempAlp = Right(st, 1) 'ZS GS S1 C1
  If IsNumeric(TempAlp) Then
    TempAlp = Left(st, 1)
  End If
ElseIf Len(st) = 3 Then
  TempAlp = Right(st, 2) 'ZS GS S1 C1
  If IsNumeric(TempAlp) Then
    TempAlp = Left(st, 1)
  End If
End If

If InStr(st, "Lateral") > 0 Then TempAlp = Right(st, 1) 'LateralS
If InStr(st, "Diag") > 0 Then TempAlp = Mid(st, 5, 1) 'DiagSKou

Dim Rig As String
If Len((st)) = 3 Then '2022.06_耳付き用条件追加
  Rig = Right((st), 2)
ElseIf Len((st)) = 2 Then
  Rig = Right((st), 1)
End If

Dim Posi As String, HizPosi As String
Dim HizNGPosi As String, NGPosi As String

Posi = TempAlp
NGPosi = UF1("NG" & Posi)

'If IsNumeric(Rig) Then 'S1  U2  C1
'  Posi = Left((st), 1) '寸法
'  NGPosi = UF1("NG" & Posi)
'Else 'ZS GS
'  HizPosi = Right((st), 1) '歪
'  HizNGPosi = UF1("NG" & HizPosi)
'End If

'>>>>>>>>>条件分岐させる_2022.04.08**厚板でのデータ残しをしやすいようにする]]
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>11,12は耳以外必要ない9,10が対角線
If UF1.薄板 Then
  Select Case Rig
  Case 2, 4, 6, 8, 10 'この座標の時だけ****2がない場合薄板の時スルーされる
    For i = 1 To 10
    If UF1("NG" & Posi) = "許容" Then Exit For
      Select Case Posi
      Case "S": If UF1("S" & i).BackColor = Ncol Then Call Confi_Open("S", i, Lname, UF1("S" & i).Caption, val(UF1.DeleteC.Caption))
      Case "U": If UF1("U" & i).BackColor = Ncol Then Call Confi_Open("U", i, Lname, UF1("U" & i).Caption, val(UF1.DeleteC.Caption))
      Case "C": If UF1("C" & i).BackColor = Ncol Then Call Confi_Open("C", i, Lname, UF1("C" & i).Caption, val(UF1.DeleteC.Caption))
      Case "D": If UF1("D" & i).BackColor = Ncol Then Call Confi_Open("D", i, Lname, UF1("D" & i).Caption, val(UF1.DeleteC.Caption))
      Case "E": If UF1("E" & i).BackColor = Ncol Then Call Confi_Open("E", i, Lname, UF1("E" & i).Caption, val(UF1.DeleteC.Caption))
      Case "F": If UF1("F" & i).BackColor = Ncol Then Call Confi_Open("F", i, Lname, UF1("F" & i).Caption, val(UF1.DeleteC.Caption))
      End Select
    Next i
  End Select
ElseIf UF1.スラブ Then '2022.06_下側に公差外があった場合厚板と同じ条件だとスルーしてしまう
  Select Case Rig
  Case 2, 4, 6, 8, 10
    For i = 1 To 10
    If UF1("NG" & Posi) = "許容" Then Exit For
      Select Case Posi
      Case "S": If UF1("S" & i).BackColor = Ncol Then Call Confi_Open("S", i, Lname, UF1("S" & i).Caption, val(UF1.DeleteC.Caption))
      Case "U": If UF1("U" & i).BackColor = Ncol Then Call Confi_Open("U", i, Lname, UF1("U" & i).Caption, val(UF1.DeleteC.Caption))
      Case "C": If UF1("C" & i).BackColor = Ncol Then Call Confi_Open("C", i, Lname, UF1("C" & i).Caption, val(UF1.DeleteC.Caption))
      Case "D": If UF1("D" & i).BackColor = Ncol Then Call Confi_Open("D", i, Lname, UF1("D" & i).Caption, val(UF1.DeleteC.Caption))
      Case "E": If UF1("E" & i).BackColor = Ncol Then Call Confi_Open("E", i, Lname, UF1("E" & i).Caption, val(UF1.DeleteC.Caption))
      End Select
    Next i
  End Select
ElseIf UF1.厚板 Or UF1.少枚数 Then '保留をかける機会が多いサイズなので聞く回数を減らす
  Select Case Rig
  Case 4, 10
    For i = 1 To 10
    If UF1("NG" & Posi) = "許容" Then Exit For
      Select Case Posi
      Case "S": If UF1("S" & i).BackColor = Ncol Then Call Confi_Open("S", i, Lname, UF1("S" & i).Caption, val(UF1.DeleteC.Caption))
      Case "U": If UF1("U" & i).BackColor = Ncol Then Call Confi_Open("U", i, Lname, UF1("U" & i).Caption, val(UF1.DeleteC.Caption))
      Case "C": If UF1("C" & i).BackColor = Ncol Then Call Confi_Open("C", i, Lname, UF1("C" & i).Caption, val(UF1.DeleteC.Caption))
      Case "D": If UF1("D" & i).BackColor = Ncol Then Call Confi_Open("D", i, Lname, UF1("D" & i).Caption, val(UF1.DeleteC.Caption))
      Case "E": If UF1("E" & i).BackColor = Ncol Then Call Confi_Open("E", i, Lname, UF1("E" & i).Caption, val(UF1.DeleteC.Caption))
      Case "F": If UF1("F" & i).BackColor = Ncol Then Call Confi_Open("F", i, Lname, UF1("F" & i).Caption, val(UF1.DeleteC.Caption))
      End Select
    Next i
  End Select
Else '極
  Select Case Rig
  Case 4, 7, 10, 12, 14
    For i = 1 To 14 '2022.06_耳付き用条件追加********11,12,13,14が丈
    If UF1("NG" & Posi) = "許容" Then Exit For
      Select Case Posi
      Case "S": If UF1("S" & i).BackColor = Ncol Then Call Confi_Open("S", i, Lname, UF1("S" & i).Caption, val(UF1.DeleteC.Caption))
      Case "U": If UF1("U" & i).BackColor = Ncol Then Call Confi_Open("U", i, Lname, UF1("U" & i).Caption, val(UF1.DeleteC.Caption))
      Case "C": If UF1("C" & i).BackColor = Ncol Then Call Confi_Open("C", i, Lname, UF1("C" & i).Caption, val(UF1.DeleteC.Caption))
      Case "D": If UF1("D" & i).BackColor = Ncol Then Call Confi_Open("D", i, Lname, UF1("D" & i).Caption, val(UF1.DeleteC.Caption))
      Case "E": If UF1("E" & i).BackColor = Ncol Then Call Confi_Open("E", i, Lname, UF1("E" & i).Caption, val(UF1.DeleteC.Caption))
      End Select
    Next i
  End Select
End If

'NG_Backに影響する引数2(("S", 96,)
'n=96,97,98,99 *ﾗﾃﾗﾙﾎﾞｰ、直角度、定歪、ｹﾞｰｼﾞ歪
With UF1
  If NGPosi <> "許容" Then
    'ラテラルボー
    Select Case st
    Case "LateralS": If .LateralS.Visible Then If .LateralS.BackColor = Ncol Then Call Confi_Open("S", 96, Lname, .LateralS.Text, val(.DeleteC.Caption))
    Case "LateralU": If .LateralU.Visible Then If .LateralU.BackColor = Ncol Then Call Confi_Open("U", 96, Lname, .LateralU.Text, val(.DeleteC.Caption))
    Case "LateralC": If .LateralC.Visible Then If .LateralC.BackColor = Ncol Then Call Confi_Open("C", 96, Lname, .LateralC.Text, val(.DeleteC.Caption))
    Case "LateralD": If .LateralD.Visible Then If .LateralD.BackColor = Ncol Then Call Confi_Open("D", 96, Lname, .LateralD.Text, val(.DeleteC.Caption))
    Case "LateralE": If .LateralE.Visible Then If .LateralE.BackColor = Ncol Then Call Confi_Open("E", 96, Lname, .LateralE.Text, val(.DeleteC.Caption))
    Case "LateralF": If .LateralF.Visible Then If .LateralF.BackColor = Ncol Then Call Confi_Open("F", 96, Lname, .LateralF.Text, val(.DeleteC.Caption))
    End Select
    '直角度
    Select Case st
    Case "DiagSKou": If .DiagSKou.Visible Then If .DiagSKou.BackColor = Ncol Then Call Confi_Open("S", 97, Lname, .DiagSKou.Text, val(.DeleteC.Caption))
    Case "DiagUKou": If .DiagUKou.Visible Then If .DiagUKou.BackColor = Ncol Then Call Confi_Open("U", 97, Lname, .DiagUKou.Text, val(.DeleteC.Caption))
    Case "DiagCKou": If .DiagCKou.Visible Then If .DiagCKou.BackColor = Ncol Then Call Confi_Open("C", 97, Lname, .DiagCKou.Text, val(.DeleteC.Caption))
    Case "DiagDKou": If .DiagDKou.Visible Then If .DiagDKou.BackColor = Ncol Then Call Confi_Open("D", 97, Lname, .DiagDKou.Text, val(.DeleteC.Caption))
    Case "DiagEKou": If .DiagEKou.Visible Then If .DiagEKou.BackColor = Ncol Then Call Confi_Open("E", 97, Lname, .DiagEKou.Text, val(.DeleteC.Caption))
    Case "DiagFKou": If .DiagFKou.Visible Then If .DiagFKou.BackColor = Ncol Then Call Confi_Open("F", 97, Lname, .DiagFKou.Text, val(.DeleteC.Caption))
    End Select
  End If
  If NGPosi <> "許容" Then
    '歪A
    Select Case st
    Case "ZS": If .ZS.BackColor = Ncol Then Call Confi_Open("S", 98, Lname, .ZS.Text, val(.DeleteC.Caption))
    Case "ZU": If .ZU.BackColor = Ncol Then Call Confi_Open("U", 98, Lname, .ZU.Text, val(.DeleteC.Caption))
    Case "ZC": If .ZC.BackColor = Ncol Then Call Confi_Open("C", 98, Lname, .ZC.Text, val(.DeleteC.Caption))
    Case "ZD": If .ZD.BackColor = Ncol Then Call Confi_Open("D", 98, Lname, .ZD.Text, val(.DeleteC.Caption))
    Case "ZE": If .ZE.BackColor = Ncol Then Call Confi_Open("E", 98, Lname, .ZE.Text, val(.DeleteC.Caption))
    Case "ZF": If .ZF.BackColor = Ncol Then Call Confi_Open("F", 98, Lname, .ZF.Text, val(.DeleteC.Caption))
    End Select
    '歪ｹﾞｰｼﾞ
    Select Case st
    Case "GS": If .GS.BackColor = Ncol Then Call Confi_Open("S", 99, Lname, .GS.Text, val(.DeleteC.Caption))
    Case "GU": If .GU.BackColor = Ncol Then Call Confi_Open("U", 99, Lname, .GU.Text, val(.DeleteC.Caption))
    Case "GC": If .GC.BackColor = Ncol Then Call Confi_Open("C", 99, Lname, .GC.Text, val(.DeleteC.Caption))
    Case "GD": If .GD.BackColor = Ncol Then Call Confi_Open("D", 99, Lname, .GD.Text, val(.DeleteC.Caption))
    Case "GE": If .GE.BackColor = Ncol Then Call Confi_Open("E", 99, Lname, .GE.Text, val(.DeleteC.Caption))
    Case "GF": If .GF.BackColor = Ncol Then Call Confi_Open("F", 99, Lname, .GF.Text, val(.DeleteC.Caption))
    End Select
  End If
End With

Call UFsizeSave(UF1) '2024.8.5

End Sub
Public Sub NG_Erase(Posi As String, Optional All As String = "") '全消し*座標修正_2023.11.28**範囲拡大_2024.2.5
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("作業保存" & "_" & Lname)
'厚,幅,丈,歪A,ﾗﾃ,直角度
Dim Sran As Range: Set Sran = sh.Range(sh.Cells(2, 1), sh.Cells(2, 20))
Dim Uran As Range: Set Uran = sh.Range(sh.Cells(3, 1), sh.Cells(3, 20))
Dim Cran As Range: Set Cran = sh.Range(sh.Cells(4, 1), sh.Cells(4, 20))
Dim Dran As Range: Set Dran = sh.Range(sh.Cells(5, 1), sh.Cells(5, 20))
Dim Eran As Range: Set Eran = sh.Range(sh.Cells(6, 1), sh.Cells(6, 20))
Dim Fran As Range: Set Fran = sh.Range(sh.Cells(7, 1), sh.Cells(7, 20))
'外観BH
Dim S外BH As Range: Set S外BH = sh.Range(sh.Cells(10, 1), sh.Cells(10, 3))
Dim U外BH As Range: Set U外BH = sh.Range(sh.Cells(11, 1), sh.Cells(11, 3))
Dim C外BH As Range: Set C外BH = sh.Range(sh.Cells(12, 1), sh.Cells(12, 3))
Dim D外BH As Range: Set D外BH = sh.Range(sh.Cells(13, 1), sh.Cells(13, 3))
Dim E外BH As Range: Set E外BH = sh.Range(sh.Cells(14, 1), sh.Cells(14, 3))
Dim F外BH As Range: Set F外BH = sh.Range(sh.Cells(15, 1), sh.Cells(15, 3))
'歪B
Dim SHiZ_B As Range: Set SHiZ_B = sh.Range(sh.Cells(8, 15), sh.Cells(8, 16))
Dim UHiZ_B As Range: Set UHiZ_B = sh.Range(sh.Cells(9, 15), sh.Cells(9, 16))
Dim CHiZ_B As Range: Set CHiZ_B = sh.Range(sh.Cells(10, 15), sh.Cells(10, 16))
Dim DHiZ_B As Range: Set DHiZ_B = sh.Range(sh.Cells(11, 15), sh.Cells(11, 16))
Dim EHiZ_B As Range: Set EHiZ_B = sh.Range(sh.Cells(12, 15), sh.Cells(12, 16))
Dim FHiZ_B As Range: Set FHiZ_B = sh.Range(sh.Cells(13, 15), sh.Cells(13, 16))

Select Case Posi
Case "S"
  Sran.ClearContents
  S外BH.ClearContents
  SHiZ_B.ClearContents
Case "U"
  Uran.ClearContents
  U外BH.ClearContents
  UHiZ_B.ClearContents
Case "C"
  Cran.ClearContents
  C外BH.ClearContents
  CHiZ_B.ClearContents
Case "D"
  Dran.ClearContents
  D外BH.ClearContents
  DHiZ_B.ClearContents
Case "E"
  Eran.ClearContents
  E外BH.ClearContents
  EHiZ_B.ClearContents
Case "F"
  Fran.ClearContents
  F外BH.ClearContents
  FHiZ_B.ClearContents
End Select
  
'tagデータ消去
UF1.Loading.Tag = "Loading"
Select Case Posi
Case "S": Call Split_Z_Tag_0("S")
Case "U": Call Split_Z_Tag_0("U")
Case "C": Call Split_Z_Tag_0("C")
Case "D": Call Split_Z_Tag_0("D")
Case "E": Call Split_Z_Tag_0("E")
Case "F": Call Split_Z_Tag_0("F")
End Select
UF1.Loading.Tag = "" '
  
UF1.Load = "off" 'ﾌｫｰﾑｲﾍﾞﾝﾄ停止
Call Load_Text '作業保存読み込み
UF1.Load = "on"

If All <> "" Then Exit Sub

If UFConfi.Visible Then '2024.2.5_追加
  UF1.Show vbModeless
  Unload UFConfi
End If
End Sub
Public Sub NG_Back(Posi As String, n As Long) 'BackSpace　NG消し
'座標の保存とカウント**2023.4.24_追加
Dim Ncol As Long: Ncol = 8421631 'NG
Dim i As Long
UF1.FlagLabel = 0
For i = 1 To 14
  Select Case Posi
  Case "S"
    If UF1("S" & i).BackColor = Ncol Then
      UF1("S" & i) = ""
      Call EraseCount(UF1("S" & i)) '座標の保存とカウント
    End If
  Case "U"
    If UF1("U" & i).BackColor = Ncol Then
      UF1("U" & i) = ""
      Call EraseCount(UF1("U" & i))
    End If
  Case "C"
    If UF1("C" & i).BackColor = Ncol Then
      UF1("C" & i) = ""
      Call EraseCount(UF1("C" & i))
    End If
  Case "D"
    If UF1("D" & i).BackColor = Ncol Then
      UF1("D" & i) = ""
      Call EraseCount(UF1("D" & i))
    End If
  Case "E"
    If UF1("E" & i).BackColor = Ncol Then
      UF1("E" & i) = ""
      Call EraseCount(UF1("E" & i))
    End If
  End Select
Next i

For i = 1 To 10
  Select Case Posi
  Case "F"
    If UF1("F" & i).BackColor = Ncol Then
      UF1("F" & i) = ""
      Call EraseCount(UF1("F" & i))
    End If
  End Select
Next i

'n=96,97,98,99 *ﾗﾃﾗﾙﾎﾞｰ、直角度、定歪、ｹﾞｰｼﾞ歪
'EraseCountなし
If n = 96 Then 'ラテラルボー
  Select Case Posi
  Case "S": UF1.LateralS = ""
  Case "U": UF1.LateralU = ""
  Case "C": UF1.LateralC = ""
  Case "D": UF1.LateralD = ""
  Case "E": UF1.LateralE = ""
  Case "F": UF1.LateralF = ""
  End Select
ElseIf n = 97 Then '直角度
  Select Case Posi
  Case "S": UF1.DiagSKou = ""
  Case "U": UF1.DiagUKou = ""
  Case "C": UF1.DiagCKou = ""
  Case "D": UF1.DiagDKou = ""
  Case "E": UF1.DiagEKou = ""
  Case "F": UF1.DiagFKou = ""
  End Select
ElseIf n = 98 Then '定歪
  Select Case Posi
  Case "S": UF1.ZS = ""
  Case "U": UF1.ZU = ""
  Case "C": UF1.ZC = ""
  Case "D": UF1.ZD = ""
  Case "E": UF1.ZE = ""
  Case "F": UF1.ZF = ""
  End Select
ElseIf n = 99 Then 'ｹﾞｰｼﾞ歪
  Select Case Posi
  Case "S": UF1.GS = ""
  Case "U": UF1.GU = ""
  Case "C": UF1.GC = ""
  Case "D": UF1.GD = ""
  Case "E": UF1.GE = ""
  Case "F": UF1.GF = ""
  End Select
End If

UF1.FlagLabel = 1
UF1.Tag = "STOP" '音楽ｽﾄｯﾌﾟ

UF1.Show vbModeless
DoEvents
Unload UFConfi
Call LoopJudg
Call HizLoopJudg(0)

If Not UF1.極厚 Then Call Look_Blank(1, 10)
If UF1.極厚 Then Call Look_Blank(1, 14)

UF1.Manual.SetFocus
UF1.StartTB.SetFocus

UF1.Tag = "" '音楽ｽﾀｰﾄ

End Sub

Public Sub NG_Pass(Posi As String, Optional Temp As String = "許容") '許容'保留
'許容、保留はNGEraseでNG色が消えた時クリアされる
Select Case Posi
Case "S": UF1("NGS") = Temp
Case "U": UF1("NGU") = Temp
Case "C": UF1("NGC") = Temp
Case "D": UF1("NGD") = Temp
Case "E": UF1("NGE") = Temp
Case "F": UF1("NGF") = Temp
End Select
UF1.Show vbModeless
DoEvents
Unload UFConfi
End Sub
Public Sub NG_Posting(Posi As String) 'NG転記
Dim Lname As String:  Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("異常ロット用" & "_" & Lname) = False Then Call AddWS("異常ロット用" & "_" & Lname)
If ThisWorkbook.Sheets("異常ロット用" & "_" & Lname).Cells(1, 1) <> UF1.Lot Then Call AddWS("異常ロット用" & "_" & Lname)
Dim sh As Worksheet:   Set sh = ThisWorkbook.Sheets("異常ロット用" & "_" & Lname)
Dim Ncol As Long:      Ncol = 8421631
Dim i As Long, a As Long, LR As Long: LR = UR(sh, 1) + 1
Dim NG(17), NGran As Range '0配列**2024.2.14_歪分追加

sh.Cells(1, 1) = UF1.Lot
a = 0
For i = 1 To 14 '耳付き丈3,4追加+2_2022.06
  Select Case Posi
  Case "S"
    NG(a) = UF1("S" & i)
    a = a + 1 '0配列
  Case "U"
    NG(a) = UF1("U" & i)
    a = a + 1
  Case "C"
    NG(a) = UF1("C" & i)
    a = a + 1
  Case "D"
    NG(a) = UF1("D" & i)
    a = a + 1
  Case "E"
    NG(a) = UF1("E" & i)
    a = a + 1
  End Select
Next i
For i = 1 To 10
  Select Case Posi
  Case "F"
    NG(a) = UF1("F" & i)
    a = a + 1
  End Select
Next i

With UF1
  '直角度,ラテの順
  Select Case Posi
  Case "S"
    If .GS.Text <> "" Then NG(17) = "ｹﾞ:" & .GS.Text
    If .AorB.Caption = "Now_A" And .ZS.Text <> "" Then NG(16) = "A:" & .ZS.Text
    If .AorB.Caption = "Now_B" And .ZS.Text <> "" Then NG(16) = "B:" & .ZS.Text
    
    NG(14) = .DiagSKou.Text
    NG(15) = .LateralS.Text
  Case "U"
    If .GU.Text <> "" Then NG(17) = "ｹﾞ:" & .GU.Text
    If .AorB.Caption = "Now_A" And .ZU.Text <> "" Then NG(16) = "A:" & .ZU.Text
    If .AorB.Caption = "Now_B" And .ZU.Text <> "" Then NG(16) = "B:" & .ZU.Text
    
    NG(14) = .DiagUKou.Text
    NG(15) = .LateralU.Text
  Case "C"
    If .GC.Text <> "" Then NG(17) = "ｹﾞ:" & .GC.Text
    If .AorB.Caption = "Now_A" And .ZC.Text <> "" Then NG(16) = "A:" & .ZC.Text
    If .AorB.Caption = "Now_B" And .ZC.Text <> "" Then NG(16) = "B:" & .ZC.Text
    
    NG(14) = .DiagCKou.Text
    NG(15) = .LateralC.Text
  Case "D"
    If .GD.Text <> "" Then NG(17) = "ｹﾞ:" & .GD.Text
    If .AorB.Caption = "Now_A" And .ZD.Text <> "" Then NG(16) = "A:" & .ZD.Text
    If .AorB.Caption = "Now_B" And .ZD.Text <> "" Then NG(16) = "B:" & .ZD.Text
    
    NG(14) = .DiagDKou.Text
    NG(15) = .LateralD.Text
  Case "E"
    If .GE.Text <> "" Then NG(17) = "ｹﾞ:" & .GE.Text
    If .AorB.Caption = "Now_A" And .ZE.Text <> "" Then NG(16) = "A:" & .ZE.Text
    If .AorB.Caption = "Now_B" And .ZE.Text <> "" Then NG(16) = "B:" & .ZE.Text
    
    NG(14) = .DiagEKou.Text
    NG(15) = .LateralE.Text
  Case "F"
    If .GF.Text <> "" Then NG(17) = "ｹﾞ:" & .GF.Text
    If .AorB.Caption = "Now_A" And .ZF.Text <> "" Then NG(16) = "A:" & .ZF.Text
    If .AorB.Caption = "Now_B" And .ZF.Text <> "" Then NG(16) = "B:" & .ZF.Text
    
    NG(14) = .DiagFKou.Text
    NG(15) = .LateralF.Text
  End Select
  
  Set NGran = sh.Range(sh.Cells(LR, 1), sh.Cells(LR, UBound(NG) + 1))
  NGran = NG

  For i = 1 To 14 '耳付き丈3,4追加+2_2022.06
    Select Case Posi
    Case "S": If UF1("S" & i).BackColor = Ncol Then sh.Cells(LR, i).Interior.Color = Ncol
    Case "U": If UF1("U" & i).BackColor = Ncol Then sh.Cells(LR, i).Interior.Color = Ncol
    Case "C": If UF1("C" & i).BackColor = Ncol Then sh.Cells(LR, i).Interior.Color = Ncol
    Case "D": If UF1("D" & i).BackColor = Ncol Then sh.Cells(LR, i).Interior.Color = Ncol
    Case "E": If UF1("E" & i).BackColor = Ncol Then sh.Cells(LR, i).Interior.Color = Ncol
    End Select
  Next i
  For i = 1 To 10
    Select Case Posi
    Case "F": If UF1("F" & i).BackColor = Ncol Then sh.Cells(LR, i).Interior.Color = Ncol
    End Select
  Next i

  '直角度ラテ_2023.11.28*2024.2.14_歪
  Select Case Posi
  Case "S"
    If .GS.BackColor = Ncol Then sh.Cells(LR, 18).Interior.Color = Ncol
    If .ZS.BackColor = Ncol Then sh.Cells(LR, 17).Interior.Color = Ncol
    If .LateralS.BackColor = Ncol Then sh.Cells(LR, 16).Interior.Color = Ncol
    If .DiagSKou.BackColor = Ncol Then sh.Cells(LR, 15).Interior.Color = Ncol
  Case "U"
    If .GU.BackColor = Ncol Then sh.Cells(LR, 18).Interior.Color = Ncol
    If .ZU.BackColor = Ncol Then sh.Cells(LR, 17).Interior.Color = Ncol
    If .LateralU.BackColor = Ncol Then sh.Cells(LR, 16).Interior.Color = Ncol
    If .DiagUKou.BackColor = Ncol Then sh.Cells(LR, 15).Interior.Color = Ncol
  Case "C"
    If .GC.BackColor = Ncol Then sh.Cells(LR, 18).Interior.Color = Ncol
    If .ZC.BackColor = Ncol Then sh.Cells(LR, 17).Interior.Color = Ncol
    If .LateralC.BackColor = Ncol Then sh.Cells(LR, 16).Interior.Color = Ncol
    If .DiagCKou.BackColor = Ncol Then sh.Cells(LR, 15).Interior.Color = Ncol
  Case "D"
    If .GD.BackColor = Ncol Then sh.Cells(LR, 18).Interior.Color = Ncol
    If .ZD.BackColor = Ncol Then sh.Cells(LR, 17).Interior.Color = Ncol
    If .LateralD.BackColor = Ncol Then sh.Cells(LR, 16).Interior.Color = Ncol
    If .DiagDKou.BackColor = Ncol Then sh.Cells(LR, 15).Interior.Color = Ncol
  Case "E"
    If .GE.BackColor = Ncol Then sh.Cells(LR, 18).Interior.Color = Ncol
    If .ZE.BackColor = Ncol Then sh.Cells(LR, 17).Interior.Color = Ncol
    If .LateralE.BackColor = Ncol Then sh.Cells(LR, 16).Interior.Color = Ncol
    If .DiagEKou.BackColor = Ncol Then sh.Cells(LR, 15).Interior.Color = Ncol
  Case "F"
    If .GF.BackColor = Ncol Then sh.Cells(LR, 18).Interior.Color = Ncol
    If .ZF.BackColor = Ncol Then sh.Cells(LR, 17).Interior.Color = Ncol
    If .LateralF.BackColor = Ncol Then sh.Cells(LR, 16).Interior.Color = Ncol
    If .DiagFKou.BackColor = Ncol Then sh.Cells(LR, 15).Interior.Color = Ncol
  End Select
End With

sh.Cells(LR, 19) = Posi '耳付き丈3,4追加+2_2022.06****2024.2.14_17→19

With sh.Cells
  .Font.Name = "メイリオ"
  .Font.Size = 12
  .EntireColumn.AutoFit
End With
'UF1.Show vbModeless
If WsCheck("異常ロット用" & "_" & Lname) Then UF1.IzyouFl.BackColor = RGB(255, 0, 0)
'DoEvents
'Unload UFConfi
End Sub
Public Sub 異常値除去(TB As MSForms.Label)
If Not UF1.Abno.Caption = "異常値on" Then Exit Sub

Dim ATU As Double '厚
Dim HABA As Single '巾
Dim TAKE As Single '丈
Dim i As Long, a As Long
Dim Am As Double: Am = val(UF1.Am) * 2
Dim Ap As Double: Ap = val(UF1.Ap) * 2
Dim Hm As Double: Hm = val(UF1.Hm) * 2
Dim Hp As Double: Hp = val(UF1.Hp) * 2
Dim Tm As Double: Tm = val(UF1.Tm) * 2
Dim Tp As Double: Tp = val(UF1.Tp) * 2
Dim iatu1 As Double, ihaba1 As Double, itake1 As Double, itai1 As Double '-
Dim iatu2 As Double, ihaba2 As Double, itake2 As Double, itai2 As Double '+
ATU = val(UF1.ATU)
HABA = val(UF1.Hab)
TAKE = val(UF1.Tak)
iatu1 = Format(ATU - Am, "0.000")
iatu2 = Format(ATU + Ap, "0.000")
ihaba1 = Format(HABA - Hm, "0.0")
ihaba2 = Format(HABA + Hp, "0.0")
itake1 = Format(TAKE - Tm, "0.0")
itake2 = Format(TAKE + Tp, "0.0")
'3平方の定理　a^2＋b^2＝c^2  Sqr関数＝数値の平方根を返す
itai1 = Format(Sqr(val(HABA) ^ 2 + val(TAKE) ^ 2) - 6, "0.0") '-
itai2 = Format(Sqr(val(HABA) ^ 2 + val(TAKE) ^ 2) + 6, "0.0") '+
If TB.Caption = "" Then Exit Sub '文字　空白の時は何もしない
'厚
If (Len(TB.Name) = 2) Then
  Select Case Right((TB.Name), 1)
  Case 1, 2, 3, 4
    If UF1.Am = "" Then Exit Sub '2025.8.8
    If Era_flag(TB, iatu2, iatu1) Then Call AbnoFormStart(TB.Caption, TB.Name) 'UF異常値発見起動
  End Select
End If
'巾
If UF1.極厚 Then
  If (Len(TB.Name) = 2) Then
    Select Case Right((TB.Name), 1)
    Case 5 To 9
        If UF1.Hm = "" Then Exit Sub '2025.8.8
        If Era_flag(TB, ihaba2, ihaba1) Then Call AbnoFormStart(TB.Caption, TB.Name)
    End Select
  ElseIf (Len(TB.Name) = 3) Then
    Select Case Right((TB.Name), 2)
    Case 10
        If UF1.Hm = "" Then Exit Sub '2025.8.8
        If Era_flag(TB, ihaba2, ihaba1) Then Call AbnoFormStart(TB.Caption, TB.Name)
    End Select
  End If
Else
  If (Len(TB.Name) = 2) Then
    Select Case Right((TB.Name), 1)
    Case 5, 6
        If UF1.Hm = "" Then Exit Sub '2025.8.8
        If Era_flag(TB, ihaba2, ihaba1) Then Call AbnoFormStart(TB.Caption, TB.Name)
    End Select
  End If
End If
'丈
If UF1.極厚 Then
  If (Len(TB.Name) = 3) Then
    Select Case Right((TB.Name), 2)
    Case 11, 12, 13, 14
        If UF1.Tm = "" Then Exit Sub '2025.8.8
        If Era_flag(TB, itake2, itake1) Then Call AbnoFormStart(TB.Caption, TB.Name)
    End Select
  End If
Else
  If (Len(TB.Name) = 2) Then
    Select Case Right((TB.Name), 1)
    Case 7, 8
        If UF1.Tm = "" Then Exit Sub '2025.8.8
        If Era_flag(TB, itake2, itake1) Then Call AbnoFormStart(TB.Caption, TB.Name)
    End Select
  End If
End If
If UF1.極厚 Then Exit Sub '対角の異常値判定が不要_2021.02.11
'対角
If (Len(TB.Name) = 3) Then
  Select Case Right((TB.Name), 2)
  Case 10
        If UF1.Tm = "" Then Exit Sub '2025.8.8
        If Era_flag(TB, itai2, itai1) Then Call AbnoFormStart(TB.Caption, TB.Name)
  End Select
ElseIf (Len(TB.Name) = 2) Then
  Select Case Right((TB.Name), 1)
  Case 9
        If UF1.Tm = "" Then Exit Sub '2025.8.8
        If Era_flag(TB, itai2, itai1) Then Call AbnoFormStart(TB.Caption, TB.Name)
  End Select
End If
 
End Sub
Public Sub NG_Back_TB(TB As String) 'BackSpace
UF1.FlagLabel = 0
UF1(TB).Caption = ""
UF1.INP = ""
UF1(TB).BackColor = &H80000005
Call Posting '消した値の反映
UF1.FlagLabel = 1
UF1.Manual.SetFocus
UF1.StartTB.SetFocus
End Sub

Public Sub CBAdd(Optional ATU As Single = 6#)   'CB登録
On Error Resume Next
'2022.08.16_中歪追加
Application.EnableEvents = False
UF1.CBS1.clear
UF1.CBU1.clear
UF1.CBC1.clear
UF1.CBD1.clear
UF1.CBE1.clear
UF1.CBF1.clear

UF1.CBS3.clear
UF1.CBU3.clear
UF1.CBC3.clear
UF1.CBD3.clear
UF1.CBE3.clear
UF1.CBF3.clear

Application.EnableEvents = True
Dim Comb_data As Variant, CombG_data As Variant, Comb As Variant
If ATU <= 6# Then '2024.4.18_6.00mm超えは "耳", "中"不要
  Comb_data = Array("幅", "丈", "耳", "中")
Else
  Comb_data = Array("幅", "丈")
End If

For Each Comb In Comb_data
 UF1.CBS1.AddItem Comb
 UF1.CBU1.AddItem Comb
 UF1.CBC1.AddItem Comb
 UF1.CBD1.AddItem Comb
 UF1.CBE1.AddItem Comb
 UF1.CBF1.AddItem Comb
Next

'2023.4.18_ﾌﾘｰｿﾘ追加
Dim i As Long, a As Boolean
If UF1.LGS.Caption = "ﾌﾘｰｿﾘ" Then a = True
If UF1.LGU.Caption = "ﾌﾘｰｿﾘ" Then a = True
If UF1.LGC.Caption = "ﾌﾘｰｿﾘ" Then a = True
If UF1.LGD.Caption = "ﾌﾘｰｿﾘ" Then a = True
If UF1.LGE.Caption = "ﾌﾘｰｿﾘ" Then a = True
If UF1.LGE.Caption = "ﾌﾘｰｿﾘ" Then a = True
If UF1.LGF.Caption = "ﾌﾘｰｿﾘ" Then a = True

If a Then
  CombG_data = Array("上", "下")
Else
  CombG_data = Array("幅", "丈")
End If

For Each Comb In CombG_data
  UF1.CBS3.AddItem Comb
  UF1.CBU3.AddItem Comb
  UF1.CBC3.AddItem Comb
  UF1.CBD3.AddItem Comb
  UF1.CBE3.AddItem Comb
  UF1.CBF3.AddItem Comb
Next
End Sub
Public Sub CB_UpDownAdd(CB As MSForms.ComboBox) '定盤上下型登録
'2022.08.16_中歪追加
UF1(Mid((CB.Name), 1, 3) & "2").clear
Select Case CB.Text
Case "幅", "丈"
    UF1(Mid((CB.Name), 1, 3) & "2").AddItem "上"
    UF1(Mid((CB.Name), 1, 3) & "2").AddItem "下"
Case "中"
    UF1(Mid((CB.Name), 1, 3) & "2").AddItem "Ⅰ型"
    UF1(Mid((CB.Name), 1, 3) & "2").AddItem "Ⅱ型"
    UF1(Mid((CB.Name), 1, 3) & "2").AddItem "Ⅲ型"
Case "耳"
    UF1(Mid((CB.Name), 1, 3) & "2").AddItem "Ⅱ型"
    UF1(Mid((CB.Name), 1, 3) & "2").AddItem "Ⅲ型"
End Select
End Sub

Public Sub CBG_UpDownAdd(CB As MSForms.ComboBox) 'ｹﾞｰｼﾞ上下型登録
UF1(Mid((CB.Name), 1, 3) & "4").clear
Select Case CB.Text
Case "幅", "丈"
    UF1(Mid((CB.Name), 1, 3) & "4").AddItem "上"
    UF1(Mid((CB.Name), 1, 3) & "4").AddItem "下"
End Select
End Sub

Public Sub Visual_Add() '外観異常名登録
'Dim Comb_data As Variant, Comb As Variant, List_Data1() As Variant, List_Data2() As Variant

''外観タブのリスト　ハードコード過ぎたのでDB読み取りに変更_2026.09.17
'Comb_data = Array("", "欠陥ﾅｼ", "当り跡ﾞ", "当りｷｽﾞ", "扱いｷｽﾞ", "圧込み", "圧込み状ﾑｼﾚ", "油圧込み", "油残り跡", _
'"油焼け", "油汚れ", "油寄り", "AL粉圧込み", "AL片圧込み", "AL付着跡", "板厚外れ", "異物巻込み", "異物圧込み", "ｴｯｼﾞﾞ酸化", _
'"折れ", "外注欠陥", "ｶﾐｺﾐ跡", "ｷｽﾞ", "ｺｰﾃｨﾝｸﾞﾑﾗ", "承台跡", "ｼｮｯｸ跡", "ｽｼﾞ", "ｽﾄﾘｰｸﾑｼﾚ", "ｽﾍﾟｰｻｰｷｽﾞ", "ｽﾘｷｽﾞ", _
'"ｽﾘｯﾌﾟ跡", "ｽﾘｰﾛｰﾙｽﾘｯﾌﾟ跡", "ｽﾘｯﾌﾟｷｽﾞ", "寸法外れ", "切断面異常", "段", "ﾁｬﾀﾘﾝｸﾞ", "突上げ", "突下げ", "ﾂﾔﾑﾗ", "ﾃｰﾌﾞﾙｽﾘ跡", "ﾃｰﾙﾏｰｸ", _
'"ﾄﾋﾞｺﾐ", "共ｽﾞﾚ", "共ｽﾞﾚ圧込み", "刃跡", "ﾊﾟｲﾗｰｷｽﾞ", "肌荒れ", "歪", "ﾌｸﾚ", "腐食", "ﾍｺ", "ﾍﾘﾝｸﾞﾎﾞｰﾝ", "偏析", _
'"補助ﾛｰﾙｷｽﾞ", "Hot共ｽﾞﾚ", "巻付段", "ﾑｼﾚ", "焼刃", "汚れ", "QF汚れ", "QF黒ﾎﾟﾁ汚れ", "QF白ﾎﾟﾁ汚れ", "QF湯流れ", "ﾚﾍﾞﾗｰｷｽﾞ", "ﾚﾍﾞﾗｰｽｼﾞ", "ﾛｰﾙｷｽﾞ", "ﾛｰﾙﾏｰｸ", "割れ", _
'"ﾋﾞﾋﾞﾘ段", "巻付跡", "ﾎﾞﾀ落ち汚れ", "ﾊﾟｯﾄｽｼﾞ", "鋳造剥離", "SSﾏｰｸ", "ｶｯﾌﾟﾙ跡") '2023.9.8_ｽﾘｷｽﾞがスキリズになっていた　2024.3.14_SSﾏｰｸ追加　2024.4.3_AL付着跡追加
''2025.01.14_欠陥名追加・QF汚れ・QFの黒ポチ汚れ・QFの白ポチ汚れ・QF湯流れ・スリーロールスリップ跡
''2026.09.17_欠陥名追加・ｶｯﾌﾟﾙ跡
'UF1.LBVisual.List = Comb_data

Dim step As String        'エラー時にどの段階かを表示する
On Error GoTo ErrHandler

Dim Comb As Variant, List_Data1() As Variant, List_Data2() As Variant

'----------------------------------------
' 1. 外観異常名：品質基準.accdb の「欠陥名」テーブルから読込
'----------------------------------------
step = "欠陥名フィールド取得"
Dim TableName As String:        TableName = "欠陥名"
Dim FPath As String:            FPath = QualityPath
Dim Target_D As String:         Target_D = "品質基準.accdb"
Dim FilePath As String:         FilePath = FPath & Target_D

Dim fieldsfilter As String:     fieldsfilter = ""
Dim Hikifields As Variant:      Hikifields = GetFieldsArr(FilePath, TableName, fieldsfilter)
If IsEmpty(Hikifields) Then Exit Sub

step = "欠陥名フィールド検索"
Dim i As Long
Dim H_欠陥名 As Long:           H_欠陥名 = -1
Dim Rcount As Long:             Rcount = LBound(Hikifields, 1)
For i = LBound(Hikifields, 2) To UBound(Hikifields, 2)
    Select Case Hikifields(Rcount, i)
        Case "欠陥名": H_欠陥名 = i
    End Select
Next i
If H_欠陥名 = -1 Then
'        MsgBox "「" & TableName & "」テーブルに「欠陥名」フィールドがありません", vbExclamation
    Exit Sub
End If

step = "欠陥名レコード取得"
Dim SQLfilter As String:        SQLfilter = ""
Dim TempD As Variant:           TempD = GetRecordsArr(FilePath, TableName, SQLfilter)
If IsEmpty(TempD) Then Exit Sub

step = "欠陥名リスト作成"
Dim VisualList() As String
Dim r As Long, n As Long
ReDim VisualList(0 To UBound(TempD, 1) - LBound(TempD, 1) + 1)
VisualList(0) = ""                               '先頭は空白行
n = 0
For r = LBound(TempD, 1) To UBound(TempD, 1)
    If Len(TempD(r, H_欠陥名) & "") > 0 Then      'Null・空文字は除外
        n = n + 1
        VisualList(n) = CStr(TempD(r, H_欠陥名))
    End If
Next r
ReDim Preserve VisualList(0 To n)
UF1.LBVisual.List = VisualList

'----------------------------------------
' 2. 山
'----------------------------------------
step = "山リスト設定"
Comb = Array("1", "2", "3", "4", "5", "6", "枚数")
UF1.YamaLB.List = Comb

'----------------------------------------
' 3. 不良要因
'----------------------------------------
step = "不良要因リスト設定"
If WsCheck("不良要因印刷原紙") = False Then Exit Sub
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("不良要因印刷原紙")
Dim rangeA As Range: Set rangeA = ws.Range(ws.Cells(2, 3), ws.Cells(2, 24))
Dim rangeB As Range: Set rangeB = ws.Range(ws.Cells(9, 3), ws.Cells(9, 24))

Dim TempA(): TempA = rangeA
Dim tempB(): tempB = rangeB

Dim Temp(): Temp = Call_MergeArray_Row(TempA, tempB)            '二次元配列を行方向に結合
Dim ReplaceTemp(): ReplaceTemp = WorksheetFunction.Transpose(Temp) 'テーブルデータを縦横変換

With UF1.DefectFactorLB
    .ColumnCount = 2          '表示列数
    .ColumnWidths = "78;15"   '列幅
    .List = ReplaceTemp
    .AddItem "ｽﾎﾟｯﾄ処理"      '追加_2023.6.16
End With
Exit Sub

ErrHandler:
MsgBox "外観異常名登録でエラーが発生しました" & vbCrLf & _
       "段階: " & step & vbCrLf & _
       "内容: " & Err.Number & " " & Err.Description, vbCritical
End Sub

Public Sub Reasons_Add() '削除理由登録
Dim List_data As Variant
List_data = Array("入力位置間違い", "測定位置間違い", "測定ボタン押し間違い", "再測定", "0点確認", "板廃棄", "その他")

With UFConfi.List1
  .ColumnCount = 1 '表示列数
  .ColumnWidths = "75"  '列幅
  .List = List_data
End With

End Sub

Public Sub LcColor(La As MSForms.Label, n As Long)
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim fname As String: fname = "Measured.log"
Dim path As String: path = HistoryPath
Dim Lpath As String, TB As String
Select Case Lname
Case "L-1":    Lpath = path & "L-1\L-1_Log"
Case "LVC":    Lpath = path & "LVC\LVC_Log"
Case "HVC":    Lpath = path & "HVC\HVC_Log"
Case "ﾊﾞﾗﾝｻｰ": Lpath = path & "バランサー\バランサー_Log"
Case "ﾄｯﾄ":    Lpath = path & "トット\トット_Log"
Case "Test":   Lpath = path & "Test\Test_Log"
Case Else:     Lpath = path & "丸徳\丸徳_Log"
End Select

Dim WorkerName As String: WorkerName = GetWorkerName()

'色付けし隣のラベルに文字列表記
Dim i As Long
UF1(La.Name).Caption = ""
' LC1からLC10までのラベルをループ処理
Select Case n
Case 1, 7
  For i = 1 To 12
    UF1("LC" & i).BackColor = RGB(255, 255, 255) ' 白色' クリックされたラベル以外は白色にする
  Next i
End Select

UF1(La.Name).BackColor = RGB(0, 255, 255) ' 水色' クリックされたラベルに水色をつける
Dim Temp As String: Temp = GetSelectedItems(UF1.LBVisual)

Select Case True
Case UF1("LC1").BackColor = RGB(0, 255, 255)
  UF1("S_UpVisual").Caption = Temp
  TB = "S_UpVisual"
Case UF1("LC2").BackColor = RGB(0, 255, 255)
  UF1("U_UpVisual").Caption = Temp
  TB = "U_UpVisual"
Case UF1("LC3").BackColor = RGB(0, 255, 255)
  UF1("C_UpVisual").Caption = Temp
  TB = "C_UpVisual"
Case UF1("LC4").BackColor = RGB(0, 255, 255)
  UF1("D_UpVisual").Caption = Temp
  TB = "D_UpVisual"
Case UF1("LC5").BackColor = RGB(0, 255, 255)
  UF1("E_UpVisual").Caption = Temp
  TB = "E_UpVisual"
Case UF1("LC6").BackColor = RGB(0, 255, 255)
  UF1("F_UpVisual").Caption = Temp
  TB = "F_UpVisual"
End Select

Select Case True
Case UF1("LC7").BackColor = RGB(0, 255, 255)
  UF1("S_UnderVisual").Caption = Temp
  TB = "S_UnderVisual"
Case UF1("LC8").BackColor = RGB(0, 255, 255)
  UF1("U_UnderVisual").Caption = Temp
  TB = "U_UnderVisual"
Case UF1("LC9").BackColor = RGB(0, 255, 255)
  UF1("C_UnderVisual").Caption = Temp
  TB = "C_UnderVisual"
Case UF1("LC10").BackColor = RGB(0, 255, 255)
  UF1("D_UnderVisual").Caption = Temp
  TB = "D_UnderVisual"
Case UF1("LC11").BackColor = RGB(0, 255, 255)
  UF1("E_UnderVisual").Caption = Temp
  TB = "E_UnderVisual"
Case UF1("LC12").BackColor = RGB(0, 255, 255)
  UF1("F_UnderVisual").Caption = Temp
  TB = "F_UnderVisual"
End Select

Call WriteLog("*UFKenl" & TB & "l" & Temp & "l" & Lname & "|" & WorkerName, "", Lpath, fname)
Call WriteLog("*UFKenl" & TB & "l" & Temp & "l" & Lname & "|" & WorkerName, "")

End Sub
Public Function GetSelectedItems(La As MSForms.ListBox) As String 'リストボックスから選んだ項目を抜き取る
Dim SelectedItems As String, i As Long
'完全一致では指定した文字列と判定する文字列を=
'部分一致では指定した文字列と判定する文字列を =ではなくLike
If La.Name <> "DefectFactorLB" Then
  With La
    For i = 0 To .ListCount - 1 'リストボックスの行数
      If .Selected(i) Then '選択されていた場合
        If SelectedItems = "" Then
          If Not (.List(i) = SelectedItems) Then SelectedItems = .List(i) '1こめのとき
        Else
          If Not (.List(i) = SelectedItems) Then SelectedItems = SelectedItems & " " & .List(i)  '複数あるとき
        End If
      End If
    Next i
  End With
Else
  With La
    For i = 0 To .ListCount - 1 'リストボックスの行数
      If .Selected(i) Then '選択されていた場合
        If SelectedItems = "" Then
          If Not (.List(i) = SelectedItems) Then SelectedItems = .List(i, 0) & "_" & .List(i, 1) '1こめのとき
        Else
          If Not (.List(i) = SelectedItems) Then SelectedItems = SelectedItems & " " & .List(i, 0) & "_" & .List(i, 1) '複数あるとき
        End If
      End If
    Next i
  End With
End If
If SelectedItems = "" Then Exit Function

If Len(SelectedItems) = 0 Then GetSelectedItems = "" ' 選択された項目がない場合は空を返す
If Right(SelectedItems, 1) = "_" Then SelectedItems = Left(SelectedItems, Len(SelectedItems) - 1) '最後が_の時は消す_2024.1.10
GetSelectedItems = SelectedItems
End Function

Public Sub btn_allClea(La As MSForms.ListBox) 'リスト件数分ループしてSelectedをFalseにする
Dim i As Long
For i = 0 To La.ListCount - 1
    La.Selected(i) = False
Next
End Sub
Public Sub DrawListBoxContent(FormTemp As Object, ThickTemp As Single) '選択した山数で描写をする
If FormTemp.YamaLB.Value = "" Then Exit Sub
Dim SelectedValue As Variant: SelectedValue = FormTemp.YamaLB.Value ' 選択された値を取得

'FormTemp.Yama.text = SelectedValue’山数自動セット**重くなるから停止

FormTemp.高さ計算式.Caption = "" 'リセット
Dim i As Long
'リセット
  For i = 1 To 6
    FormTemp("HL" & i).Visible = False
  Next i
  For i = 1 To 7
    FormTemp("KenHBox" & i).Visible = False
  Next i
  
  '2024.10.30追加
  Dim TemA As Long
  If FormTemp.Name = "UFMitsubishi" Then TemA = 7
  If FormTemp.Name <> "UFMitsubishi" Then TemA = 10
  For i = 1 To TemA
    FormTemp("KenHL" & i).Visible = False
  Next i
  
  If FormTemp.Name <> "UFMitsubishi" Then FormTemp.KenCount.Visible = False
  
  
  Select Case SelectedValue
  Case "1"
    FormTemp.HL1.Visible = True
    FormTemp.HL1.Caption = "1"
    FormTemp.KenHBox1.Visible = True
    FormTemp.KenHL1.Visible = True
  Case "2"
    For i = 1 To 2
      FormTemp("HL" & i).Visible = True
      FormTemp.HL1.Caption = "2-2"
      FormTemp.HL2.Caption = "2-1"
      
      FormTemp("KenHBox" & i).Visible = True
      FormTemp("KenHL" & i).Visible = True
    Next i
  Case "3"
    For i = 1 To 3
      FormTemp("HL" & i).Visible = True
      FormTemp.HL1.Caption = "3-3"
      FormTemp.HL2.Caption = "3-2"
      FormTemp.HL3.Caption = "3-1"
      FormTemp("KenHBox" & i).Visible = True
      FormTemp("KenHL" & i).Visible = True
    Next i
  Case "4"
    For i = 1 To 4
      FormTemp("HL" & i).Visible = True
      FormTemp.HL1.Caption = "4-4"
      FormTemp.HL2.Caption = "4-3"
      FormTemp.HL3.Caption = "4-2"
      FormTemp.HL4.Caption = "4-1"
      FormTemp("KenHBox" & i).Visible = True
      FormTemp("KenHL" & i).Visible = True
    Next i
  Case "5"
    For i = 1 To 5
      FormTemp("HL" & i).Visible = True
      FormTemp.HL1.Caption = "5-5"
      FormTemp.HL2.Caption = "5-4"
      FormTemp.HL3.Caption = "5-3"
      FormTemp.HL4.Caption = "5-2"
      FormTemp.HL5.Caption = "5-1"
      FormTemp("KenHBox" & i).Visible = True
      FormTemp("KenHL" & i).Visible = True
    Next i
  Case "6"
    For i = 1 To 6
      FormTemp("HL" & i).Visible = True
      FormTemp.HL1.Caption = "6-6"
      FormTemp.HL2.Caption = "6-5"
      FormTemp.HL3.Caption = "6-4"
      FormTemp.HL4.Caption = "6-3"
      FormTemp.HL5.Caption = "6-2"
      FormTemp.HL6.Caption = "6-1"
      FormTemp("KenHBox" & i).Visible = True
      FormTemp("KenHL" & i).Visible = True
    Next i
  Case "枚数"
    FormTemp.KenHBox7.Visible = True
    If FormTemp.Name <> "UFMitsubishi" Then FormTemp.KenCount.Visible = True
    For i = 7 To TemA '2024.10.30追加
      FormTemp("KenHL" & i).Visible = True
    Next i
  End Select
  Call CalculateSheets(FormTemp, val(ThickTemp)) '再計算
  FormTemp.Repaint
End Sub

Public Function WithorWithout_Gauge() As String 'ゲージの有無判定　Q-S-QA-DIM-173*平坦度測定基準
'Q-S-QA-DIM-173*平坦度測定基準_パターンA　35kg以下のときゲージ不要＊両面定盤歪測定必要
'Q-S-QA-DIM-173*平坦度測定基準_パターンB　35kg超えのときゲージ要＊両面歪測定時でもゲージによる測定要
'Q-S-QA-DIM-173*平坦度測定基準_パターンC　ゲージによる測定のみ＊定盤に載せれない
'F行追加_2024.5.23
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim TempW As Double: TempW = Weight(val(UF1.ATU), val(UF1.Hab), val(UF1.Tak))
If TempW = 0 Then Exit Function '重量なし処理停止
If WsCheck("作業保存" & "_" & Lname) = False Then Exit Function 'ｼｰﾄなし処理停止

'歪AB面のチェック
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("作業保存" & "_" & Lname)
Dim A_range As Range: Set A_range = sh.Range(sh.Cells(2, 15), sh.Cells(7, 16)) 'A面
Dim B_range As Range: Set B_range = sh.Range(sh.Cells(8, 15), sh.Cells(13, 16)) 'B面

Dim A_Dis(): A_Dis = A_range
Dim B_Dis(): B_Dis = B_range
'整理
Dim i As Long
For i = LBound(A_Dis, 1) To UBound(A_Dis, 1)
  If A_Dis(i, 1) = "_" Or A_Dis(i, 1) = "__" Then A_Dis(i, 1) = ""
  If A_Dis(i, 2) = "_" Or A_Dis(i, 2) = "__" Then A_Dis(i, 2) = ""
Next i
For i = LBound(B_Dis, 1) To UBound(B_Dis, 1)
  If B_Dis(i, 1) = "_" Or B_Dis(i, 1) = "__" Then B_Dis(i, 1) = ""
  If B_Dis(i, 2) = "_" Or B_Dis(i, 2) = "__" Then B_Dis(i, 2) = ""
Next i

Dim flagA(1 To 6) As Boolean
Dim flagB(1 To 6) As Boolean
Dim flagAG(1 To 6) As Boolean
Dim flagBG(1 To 6) As Boolean
Dim AtuF(1 To 6) As Boolean
'定盤歪入力がある時True
If A_Dis(1, 1) <> "" Then flagA(1) = True 'S
If B_Dis(1, 1) <> "" Then flagB(1) = True 'S
If A_Dis(2, 1) <> "" Then flagA(2) = True 'U
If B_Dis(2, 1) <> "" Then flagB(2) = True 'U
If A_Dis(3, 1) <> "" Then flagA(3) = True 'C
If B_Dis(3, 1) <> "" Then flagB(3) = True 'C
If A_Dis(4, 1) <> "" Then flagA(4) = True 'D
If B_Dis(4, 1) <> "" Then flagB(4) = True 'D
If A_Dis(5, 1) <> "" Then flagA(5) = True 'E
If B_Dis(5, 1) <> "" Then flagB(5) = True 'E
If A_Dis(6, 1) <> "" Then flagA(6) = True 'F
If UF1.Yama.Text = "4" And B_Dis(6, 1) <> "" Then flagB(6) = True 'F
'ゲージ歪入力がある時True
If A_Dis(1, 2) <> "" Then flagAG(1) = True 'S
If B_Dis(1, 2) <> "" Then flagBG(1) = True 'S
If A_Dis(2, 2) <> "" Then flagAG(2) = True 'U
If B_Dis(2, 2) <> "" Then flagBG(2) = True 'U
If A_Dis(3, 2) <> "" Then flagAG(3) = True 'C
If B_Dis(3, 2) <> "" Then flagBG(3) = True 'C
If A_Dis(4, 2) <> "" Then flagAG(4) = True 'D
If B_Dis(4, 2) <> "" Then flagBG(4) = True 'D
If A_Dis(5, 2) <> "" Then flagAG(5) = True 'E
If B_Dis(5, 2) <> "" Then flagBG(5) = True 'E
If UF1.Yama.Text = "4" And A_Dis(6, 2) <> "" Then flagAG(6) = True 'F
If UF1.Yama.Text = "4" And B_Dis(6, 2) <> "" Then flagBG(6) = True 'F
'厚の測定がある時True
If UF1.S1.Caption <> "" Then AtuF(1) = True 'S
If UF1.U1.Caption <> "" Then AtuF(2) = True 'U
If UF1.C1.Caption <> "" Then AtuF(3) = True 'C
If UF1.D1.Caption <> "" Then AtuF(4) = True 'D
If UF1.E1.Caption <> "" Then AtuF(5) = True 'E
If UF1.Yama.Text = "4" And UF1.F1.Caption <> "" Then AtuF(6) = True 'F

'そもそも測定がない
For i = 1 To 6
  Select Case i
  Case 1:  If AtuF(1) And flagA(1) = False Then WithorWithout_Gauge = "SのA面 定盤 未入力"
  Case 2:  If AtuF(2) And flagA(2) = False Then WithorWithout_Gauge = "UのA面 定盤 未入力"
  Case 3:  If AtuF(3) And flagA(3) = False Then WithorWithout_Gauge = "2のA面 定盤 未入力"
  Case 4:  If AtuF(4) And flagA(4) = False Then WithorWithout_Gauge = "DのA面 定盤 未入力"
  Case 5:  If AtuF(5) And flagA(5) = False Then WithorWithout_Gauge = "EのA面 定盤 未入力"
  Case 6:  If AtuF(6) And flagA(6) = False Then WithorWithout_Gauge = "FのA面 定盤 未入力"
  End Select
Next i

'Q-S-QA-DIM-173*平坦度測定基準_パターンA　35kg以下のときゲージ不要＊両面定盤歪測定必要
For i = 1 To 6
  Select Case i
  Case 1:  If AtuF(1) And TempW <= 35# And flagA(1) And flagB(1) = False Then WithorWithout_Gauge = "SのB面 定盤 未入力_ﾊﾟﾀｰﾝA"
  Case 2:  If AtuF(2) And TempW <= 35# And flagA(2) And flagB(2) = False Then WithorWithout_Gauge = "UのB面 定盤 未入力_ﾊﾟﾀｰﾝA"
  Case 3:  If AtuF(3) And TempW <= 35# And flagA(3) And flagB(3) = False Then WithorWithout_Gauge = "2のB面 定盤 未入力_ﾊﾟﾀｰﾝA"
  Case 4:  If AtuF(4) And TempW <= 35# And flagA(4) And flagB(4) = False Then WithorWithout_Gauge = "DのB面 定盤 未入力_ﾊﾟﾀｰﾝA"
  Case 5:  If AtuF(5) And TempW <= 35# And flagA(5) And flagB(5) = False Then WithorWithout_Gauge = "EのB面 定盤 未入力_ﾊﾟﾀｰﾝA"
  Case 6:  If AtuF(6) And TempW <= 35# And flagA(6) And flagB(6) = False Then WithorWithout_Gauge = "FのB面 定盤 未入力_ﾊﾟﾀｰﾝA"
  End Select
Next i
For i = 1 To 6
  Select Case i
  Case 1:  If AtuF(1) And TempW <= 35# And flagB(1) And flagA(1) = False Then WithorWithout_Gauge = "SのA面 定盤 未入力_ﾊﾟﾀｰﾝA"
  Case 2:  If AtuF(2) And TempW <= 35# And flagB(2) And flagA(2) = False Then WithorWithout_Gauge = "UのA面 定盤 未入力_ﾊﾟﾀｰﾝA"
  Case 3:  If AtuF(3) And TempW <= 35# And flagB(3) And flagA(3) = False Then WithorWithout_Gauge = "2のA面 定盤 未入力_ﾊﾟﾀｰﾝA"
  Case 4:  If AtuF(4) And TempW <= 35# And flagB(4) And flagA(4) = False Then WithorWithout_Gauge = "DのA面 定盤 未入力_ﾊﾟﾀｰﾝA"
  Case 5:  If AtuF(5) And TempW <= 35# And flagB(5) And flagA(5) = False Then WithorWithout_Gauge = "EのA面 定盤 未入力_ﾊﾟﾀｰﾝA"
  Case 6:  If AtuF(6) And TempW <= 35# And flagB(6) And flagA(6) = False Then WithorWithout_Gauge = "FのA面 定盤 未入力_ﾊﾟﾀｰﾝA"
  End Select
Next i

'フリーそりの条件は後日/////2023.11.23


'Q-S-QA-DIM-173*平坦度測定基準_パターンB　35kg超えのときゲージ要＊両面歪測定時でもゲージによる測定要
'定盤歪みの入力があるのにかかわらずゲージがない場合にのみ未入力判定をする

'35kg超えであってもABの定盤測定がある場合はスルー
For i = 1 To 6
  Select Case i
  Case 1:  If AtuF(1) And TempW > 35# And flagA(1) And flagB(1) Then Exit Function
  Case 2:  If AtuF(2) And TempW > 35# And flagA(2) And flagB(2) Then Exit Function
  Case 3:  If AtuF(3) And TempW > 35# And flagA(3) And flagB(3) Then Exit Function
  Case 4:  If AtuF(4) And TempW > 35# And flagA(4) And flagB(4) Then Exit Function
  Case 5:  If AtuF(5) And TempW > 35# And flagA(5) And flagB(5) Then Exit Function
  Case 6:  If AtuF(6) And TempW > 35# And flagA(6) And flagB(6) Then Exit Function
  End Select
Next i

'定盤Aのみの測定の場合ゲージはいる
For i = 1 To 6
  Select Case i
  Case 1:  If AtuF(1) And TempW > 35# And flagA(1) And flagB(1) = False And flagAG(1) = False Then WithorWithout_Gauge = "SのA面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"
  Case 2:  If AtuF(2) And TempW > 35# And flagA(2) And flagB(2) = False And flagAG(2) = False Then WithorWithout_Gauge = "UのA面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"
  Case 3:  If AtuF(3) And TempW > 35# And flagA(3) And flagB(3) = False And flagAG(3) = False Then WithorWithout_Gauge = "2のA面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"
  Case 4:  If AtuF(4) And TempW > 35# And flagA(4) And flagB(4) = False And flagAG(4) = False Then WithorWithout_Gauge = "DのA面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"
  Case 5:  If AtuF(5) And TempW > 35# And flagA(5) And flagB(5) = False And flagAG(5) = False Then WithorWithout_Gauge = "EのA面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"
  Case 6:  If AtuF(6) And TempW > 35# And flagA(6) And flagB(6) = False And flagAG(6) = False Then WithorWithout_Gauge = "FのA面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"
  End Select
Next i
'定盤Bのみの測定の場合ゲージはいる(定盤のチェックで引っかかる)
For i = 1 To 5
  Select Case i
  Case 1:  If AtuF(1) And TempW > 35# And flagB(1) And flagA(1) = False And flagBG(1) = False Then WithorWithout_Gauge = "SのB面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"
  Case 2:  If AtuF(2) And TempW > 35# And flagB(2) And flagA(2) = False And flagBG(2) = False Then WithorWithout_Gauge = "UのB面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"
  Case 3:  If AtuF(3) And TempW > 35# And flagB(3) And flagA(3) = False And flagBG(3) = False Then WithorWithout_Gauge = "2のB面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"
  Case 4:  If AtuF(4) And TempW > 35# And flagB(4) And flagA(4) = False And flagBG(4) = False Then WithorWithout_Gauge = "DのB面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"
  Case 5:  If AtuF(5) And TempW > 35# And flagB(5) And flagA(5) = False And flagBG(5) = False Then WithorWithout_Gauge = "EのB面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"
  Case 6:  If AtuF(6) And TempW > 35# And flagB(6) And flagA(6) = False And flagBG(6) = False Then WithorWithout_Gauge = "FのB面 ｹﾞｰｼﾞ 未入力_ﾊﾟﾀｰﾝB"
  End Select
Next i
End Function
Public Function Distortion_Check(n As String) As String '歪3種入力ﾁｪｯｸ
Distortion_Check = ""
'測定値0の時形状をｸﾘｱさせる_2021.01.14
'F行追加_2024.5.23
With UF1
  If .ZS = 0 Then .CBS1 = ""
  If .ZU = 0 Then .CBU1 = ""
  If .ZC = 0 Then .CBC1 = ""
  If .ZD = 0 Then .CBD1 = ""
  If .ZE = 0 Then .CBE1 = ""
  If .GS = 0 Then .CBS3 = ""
  If .GU = 0 Then .CBU3 = ""
  If .GC = 0 Then .CBC3 = ""
  If .GD = 0 Then .CBD3 = ""
  If .GE = 0 Then .CBE3 = ""
  If .GF = 0 Then .CBF3 = ""
  Select Case n '測定値0の時はスルー
  Case "Z"
    Dim Z_S_Tr As Boolean, Z_U_Tr As Boolean, Z_C_Tr As Boolean, Z_D_Tr As Boolean, Z_E_Tr As Boolean, Z_F_Tr As Boolean
    If .CBS1 <> "" Or .CBS2 <> "" Or .ZS <> "" Then Z_S_Tr = True '一つでも入力がある
    If .CBU1 <> "" Or .CBU2 <> "" Or .ZU <> "" Then Z_U_Tr = True
    If .CBC1 <> "" Or .CBC2 <> "" Or .ZC <> "" Then Z_C_Tr = True
    If .CBD1 <> "" Or .CBD2 <> "" Or .ZD <> "" Then Z_D_Tr = True
    If .CBE1 <> "" Or .CBE2 <> "" Or .ZE <> "" Then Z_E_Tr = True
    If .CBF1 <> "" Or .CBF2 <> "" Or .ZF <> "" Then Z_F_Tr = True
    Dim Z_S_Fa As Boolean, Z_U_Fa As Boolean, Z_C_Fa As Boolean, Z_D_Fa As Boolean, Z_E_Fa As Boolean, Z_F_Fa As Boolean
    If .CBS1 = "" Or .CBS2 = "" Or .ZS = "" Then Z_S_Fa = True '3つのうちどれか未入力がある
    If .CBU1 = "" Or .CBU2 = "" Or .ZU = "" Then Z_U_Fa = True
    If .CBC1 = "" Or .CBC2 = "" Or .ZC = "" Then Z_C_Fa = True
    If .CBD1 = "" Or .CBD2 = "" Or .ZD = "" Then Z_D_Fa = True
    If .CBE1 = "" Or .CBE2 = "" Or .ZE = "" Then Z_E_Fa = True
    If .CBF1 = "" Or .CBF2 = "" Or .ZF = "" Then Z_F_Fa = True
    'どこかに入力がありかつ未入力がある　にもかかわらず数値が0ではない
    If Z_S_Tr And Z_S_Fa And .ZS <> 0 Then Distortion_Check = "定盤:S"
    If Z_U_Tr And Z_U_Fa And .ZU <> 0 Then Distortion_Check = "定盤:U"
    If Z_C_Tr And Z_C_Fa And .ZC <> 0 Then Distortion_Check = "定盤:C"
    If Z_D_Tr And Z_D_Fa And .ZD <> 0 Then Distortion_Check = "定盤:D"
    If Z_E_Tr And Z_E_Fa And .ZE <> 0 Then Distortion_Check = "定盤:E"
    If Z_F_Tr And Z_F_Fa And .ZE <> 0 Then Distortion_Check = "定盤:F"
  Case "G"
    Dim G_S_Tr As Boolean, G_U_Tr As Boolean, G_C_Tr As Boolean, G_D_Tr As Boolean, G_E_Tr As Boolean, G_F_Tr As Boolean
    Dim G_S_Fa As Boolean, G_U_Fa As Boolean, G_C_Fa As Boolean, G_D_Fa As Boolean, G_E_Fa As Boolean, G_F_Fa As Boolean
    '2023.8.22_フリーそり時必ず引っかけてしまう修正
    If UF1.flatnesText = "" Then
      If .CBS3 <> "" Or .CBS4 <> "" Or .GS <> "" Then G_S_Tr = True
      If .CBU3 <> "" Or .CBU4 <> "" Or .GU <> "" Then G_U_Tr = True
      If .CBC3 <> "" Or .CBC4 <> "" Or .GC <> "" Then G_C_Tr = True
      If .CBD3 <> "" Or .CBD4 <> "" Or .GD <> "" Then G_D_Tr = True
      If .CBE3 <> "" Or .CBE4 <> "" Or .GE <> "" Then G_E_Tr = True
      If .CBF3 <> "" Or .CBF4 <> "" Or .GF <> "" Then G_F_Tr = True
      
      If .CBS3 = "" Or .CBS4 = "" Or .GS = "" Then G_S_Fa = True
      If .CBU3 = "" Or .CBU4 = "" Or .GU = "" Then G_U_Fa = True
      If .CBC3 = "" Or .CBC4 = "" Or .GC = "" Then G_C_Fa = True
      If .CBD3 = "" Or .CBD4 = "" Or .GD = "" Then G_D_Fa = True
      If .CBE3 = "" Or .CBE4 = "" Or .GE = "" Then G_E_Fa = True
      If .CBF3 = "" Or .CBF4 = "" Or .GF = "" Then G_F_Fa = True
      
      If G_S_Tr And G_S_Fa And .GS <> 0 Then Distortion_Check = "ｹﾞｰｼﾞ:S"
      If G_U_Tr And G_U_Fa And .GU <> 0 Then Distortion_Check = "ｹﾞｰｼﾞ:U"
      If G_C_Tr And G_C_Fa And .GC <> 0 Then Distortion_Check = "ｹﾞｰｼﾞ:C"
      If G_D_Tr And G_D_Fa And .GD <> 0 Then Distortion_Check = "ｹﾞｰｼﾞ:D"
      If G_E_Tr And G_E_Fa And .GE <> 0 Then Distortion_Check = "ｹﾞｰｼﾞ:E"
      If G_F_Tr And G_F_Fa And .GF <> 0 Then Distortion_Check = "ｹﾞｰｼﾞ:F"
      
    ElseIf UF1.flatnesText = "ﾌﾘｰｿﾘ要" Then
      If .CBS3 <> "" Or .GS <> "" Then G_S_Tr = True
      If .CBU3 <> "" Or .GU <> "" Then G_U_Tr = True
      If .CBC3 <> "" Or .GC <> "" Then G_C_Tr = True
      If .CBD3 <> "" Or .GD <> "" Then G_D_Tr = True
      If .CBE3 <> "" Or .GE <> "" Then G_E_Tr = True
      If .CBF3 <> "" Or .GF <> "" Then G_F_Tr = True
      
      If .CBS3 = "" Or .GS = "" Then G_S_Fa = True
      If .CBU3 = "" Or .GU = "" Then G_U_Fa = True
      If .CBC3 = "" Or .GC = "" Then G_C_Fa = True
      If .CBD3 = "" Or .GD = "" Then G_D_Fa = True
      If .CBE3 = "" Or .GE = "" Then G_E_Fa = True
      If .CBF3 = "" Or .GF = "" Then G_F_Fa = True
      
      If G_S_Tr And G_S_Fa And .GS <> 0 Then Distortion_Check = "ﾌﾘｰｿﾘ:S"
      If G_U_Tr And G_U_Fa And .GU <> 0 Then Distortion_Check = "ﾌﾘｰｿﾘ:U"
      If G_C_Tr And G_C_Fa And .GC <> 0 Then Distortion_Check = "ﾌﾘｰｿﾘ:C"
      If G_D_Tr And G_D_Fa And .GD <> 0 Then Distortion_Check = "ﾌﾘｰｿﾘ:D"
      If G_E_Tr And G_E_Fa And .GE <> 0 Then Distortion_Check = "ﾌﾘｰｿﾘ:E"
      If G_F_Tr And G_F_Fa And .GF <> 0 Then Distortion_Check = "ﾌﾘｰｿﾘ:F"
    End If
  End Select
End With
End Function
Public Function Angle_Check() As String '直角度入力チェック_2023.11.30
'厚の1か所目の測定がある(検査あり)かつ直角度の表示があり入力がない場合
'F行追加_2024.5.23
Angle_Check = ""
With UF1
  Select Case True
  Case .S1.Caption <> "" And .DiagSKou.Visible And .DiagSKou = "": Angle_Check = .S.Caption
  Case .U1.Caption <> "" And .DiagUKou.Visible And .DiagUKou = "": Angle_Check = .U.Caption
  Case .C1.Caption <> "" And .DiagCKou.Visible And .DiagCKou = "": Angle_Check = .C.Caption
  Case .D1.Caption <> "" And .DiagDKou.Visible And .DiagDKou = "": Angle_Check = .D.Caption
  Case .E1.Caption <> "" And .DiagEKou.Visible And .DiagEKou = "": Angle_Check = .E.Caption
  Case .Yama.Text = "4" And .F1.Caption <> "" And .DiagFKou.Visible And .DiagFKou = "": Angle_Check = .F.Caption
  End Select
End With
End Function

Public Function Lateralbow_Check() As String 'ラテラルボー入力チェック_2021.07.14
'厚の1か所目の測定がある(検査あり)かつラテラルボーの入力がない場合　に修正_2021.11.17 *****表示ありを条件に追加_2023.11.30
'2023.5.22_不要部削除
'F行追加_2024.5.23
Lateralbow_Check = ""
With UF1
  Select Case True
  Case .S1.Caption <> "" And .LateralS.Visible And .LateralS = "": Lateralbow_Check = .S.Caption
  Case .U1.Caption <> "" And .LateralU.Visible And .LateralU = "": Lateralbow_Check = .U.Caption
  Case .C1.Caption <> "" And .LateralC.Visible And .LateralC = "": Lateralbow_Check = .C.Caption
  Case .D1.Caption <> "" And .LateralD.Visible And .LateralD = "": Lateralbow_Check = .D.Caption
  Case .E1.Caption <> "" And .LateralE.Visible And .LateralE = "": Lateralbow_Check = .E.Caption
  Case .Yama.Text = "4" And .F1.Caption <> "" And .LateralF.Visible And .LateralF = "": Lateralbow_Check = .F.Caption
  End Select
End With
End Function
Public Function Visual_Check() As String '外観検査入力チェック_2023.5.22
'厚の1か所目の測定がある(検査あり)かつ外観検査の入力がない場合
'F行追加_2024.5.23
Visual_Check = ""
With UF1
  Select Case True
  Case (.S1.Caption <> "") And (.S_UpVisual.Caption = "" Or .S_UnderVisual.Caption = ""): Visual_Check = .S.Caption
  Case (.U1.Caption <> "") And (.U_UpVisual.Caption = "" Or .U_UnderVisual.Caption = ""): Visual_Check = .U.Caption
  Case (.C1.Caption <> "") And (.C_UpVisual.Caption = "" Or .C_UnderVisual.Caption = ""): Visual_Check = .C.Caption
  Case (.D1.Caption <> "") And (.D_UpVisual.Caption = "" Or .D_UnderVisual.Caption = ""): Visual_Check = .D.Caption
  Case (.E1.Caption <> "") And (.E_UpVisual.Caption = "" Or .E_UnderVisual.Caption = ""): Visual_Check = .E.Caption
  Case (.Yama.Text = "4") And (.F1.Caption <> "") And (.F_UpVisual.Caption = "" Or .F_UnderVisual.Caption = ""): Visual_Check = .F.Caption
  End Select
End With
End Function
Public Function BH_Check(Optional Temp As String = "") As String 'BH入力チェック_2023.6.9***操業基準ﾊﾞｰｺｰﾙ硬度計による測定見直し_2023.6.22
'厚の1か所目の測定がある(検査あり)かつ　LotNoの頭がH,R　か　QF通過　かつ　入力がない場合
'QF通過　かつ　入力がない場合に修正_2023.6.22
BH_Check = ""
Dim i As Long
Dim flag As Boolean: flag = False
If InStr(UF1.QF.Caption, "QF") > 0 Then flag = True
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
    If (flag) And (.SBH = "") Then BH_Check = .S.Caption
  Case countL(2)
    If (flag) And (.SBH = "") Then BH_Check = .S.Caption
    If (flag) And (.UBH = "") Then BH_Check = .U.Caption
  Case countL(3)
    If (flag) And (.SBH = "") Then BH_Check = .S.Caption
    If (flag) And (.CBH = "") Then BH_Check = .C.Caption
  Case countL(4)
    If (flag) And (.SBH = "") Then BH_Check = .S.Caption
    If (flag) And (.DBH = "") Then BH_Check = .D.Caption
  Case countL(5)
    If (flag) And (.SBH = "") Then BH_Check = .S.Caption
    If (flag) And (.EBH = "") Then BH_Check = .E.Caption
   End Select
End With
End Function
Public Function H_Check() As String  '検入高さ入力チェック_2023.6.9
'厚の1か所目の測定がある(検査あり)かつすべてのBOXに入力がないの場合(1つでもあればスルー)
Dim i As Long, flag As Boolean: flag = False
H_Check = ""
For i = 1 To 7
  If UF1("KenHBox" & i) <> "" Then 'どれか一つでも入力がある
    flag = True
    H_Check = ""
    Exit Function
  End If
Next i
If flag = False Then H_Check = "入力なし"
End Function

Public Function Conductiv_Check() As Boolean '導電率チェック trueで検査要
'作業指示書:QA-021-08　対象焼き入れ材の導電率測定指示
With UF1
  Select Case .Zai
'  Case "N17S": If .Tyu = "T351" Then Conductiv_Check = True’QA-021-08-R02で調質T3**削除
'  Case "17S":  If .Tyu = "T3" Then Conductiv_Check = True
  Case "75S":  If .Tyu = "T651" Or .Tyu = "T6" Then Conductiv_Check = True
'  Case "24S":  If .Tyu = "T3" Or .Tyu = "T351" Then Conductiv_Check = True
  Case "14S":  If .Tyu = "T6" Or .Tyu = "T651" Then Conductiv_Check = True
  End Select
End With
End Function
Public Function Conductiv_Coolingtime_Check(lottemp As String) As String '導電率空冷時間チェックチェック trueで検査要
'作業指示書:QA-021-08　対象焼き入れ材の導電率測定指示
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Application.ScreenUpdating = False  '画面更新停止
Application.DisplayAlerts = False

Dim ListPass As String: ListPass = QualityPath & "75S・14S焼入れチェック\"
Dim ListName As String: ListName = "QF空冷時間チェックリスト.xlsx" '"QF空冷時間チェックリスト.xlsx"
If ReadCheck(ListPass & ListName, 1) = False Then
  MsgBox ListPass & ListName & "がありません" & "_Conductiv_Coolingtime_Check", vbCritical
  Exit Function
End If

'実行しているエクセルとは、違う新しいアプリケーションでエクセルを開く_2024.5.13
Dim ExcelApp As New Application
ExcelApp.Visible = False
ExcelApp.DisplayAlerts = False
'Dim KenBook As Workbook: Set KenBook = Workbooks.Open(ListPass & ListName) ', UpdateLinks:=0, ReadOnly:=True, CorruptLoad:=xlRepairFile)
Dim wb As Workbook: Set wb = ExcelApp.Workbooks.Open(ListPass & ListName, , True) '読取り専用

Dim ws As Worksheet 'ﾜｰｸｼｰﾄ
Dim WsC As Boolean: WsC = False 'ｼｰﾄを探す
For Each ws In wb.Worksheets
  If ws.Name = "Sheet1" Then
    WsC = True
    Exit For
  End If
Next ws

'Dim WS As Worksheet 'ﾜｰｸｼｰﾄ
'Dim WsC As Boolean: WsC = False 'ｼｰﾄを探す
'For Each WS In KenBook.Worksheets
'  If WS.Name = "Sheet1" Then
'    WsC = True
'    Exit For
'  End If
'Next WS

Set ws = Nothing
Dim ConductivLot As Worksheet
Dim i As Long, flag As Boolean: flag = False

If WsC Then
  Set ConductivLot = wb.Sheets("Sheet1") '品証指定のシートを配列化
  Dim EndR As Long:   EndR = ConductivLot.Cells(Rows.Count, 2).End(xlUp).row
  Dim EndC As Long:   EndC = ConductivLot.Cells(4, Columns.Count).End(xlToLeft).Column
  Dim HistoryRan As Range
  'ConductivLot.Range(ConductivLot.Cells(X,X),ConductivLot.Cells(X,X)) Cellsの前にいちいちシートオブジェクトをつけることでエラーを回避
  Dim BufRan As Range: Set BufRan = ConductivLot.Range(ConductivLot.Cells(5, 1), ConductivLot.Cells(EndR, EndC))
  Dim buf(): buf = BufRan
  For i = LBound(buf, 1) To UBound(buf, 1) 'LotNo検索
    If buf(i, 2) = lottemp Then 'Lot一致
    flag = True
      If buf(i, 1) <> "" Then '写真確認が空白でない
        If IsNumeric(buf(i, 1)) Then
          Conductiv_Coolingtime_Check = buf(i, 1) 'かつ数字
        Else
          Conductiv_Coolingtime_Check = buf(i, 1)  '文字列
        End If
      ElseIf buf(i, 1) = "" Then
        Conductiv_Coolingtime_Check = "" '"LotNoは在るが写真確認Noがない"
      End If
    End If
  Next i
  
  If flag = False Then Conductiv_Coolingtime_Check = "LotNo記載がない"
Else
  MsgBox "指定シートがありません" & "_Conductiv_Coolingtime_Check", vbCritical
  Conductiv_Coolingtime_Check = ""
  
  wb.Close
  ExcelApp.Quit '必須
  
  Set ExcelApp = Nothing
  Exit Function
End If
  
wb.Close
ExcelApp.Quit '必須
Set ExcelApp = Nothing

'If WsC Then
'  Set ConductivLot = KenBook.Sheets("Sheet1") '品証指定のシートを配列化
'  Dim EndR As Long:   EndR = UR(ConductivLot, 2)
'  Dim EndC As Long:   EndC = UC(ConductivLot, 4)
'  Dim HistoryRan As Range
'  Dim BufRan As Range: Set BufRan = Range(ConductivLot.Cells(5, 1), ConductivLot.Cells(EndR, EndC))
'  Dim buf(): buf = BufRan
'  KenBook.Close
'  For i = LBound(buf, 1) To UBound(buf, 1) 'LotNo検索
'    If buf(i, 2) = lottemp Then 'Lot一致
'    flag = True
'      If buf(i, 1) <> "" Then '写真確認が空白でない
'        If IsNumeric(buf(i, 1)) Then
'          Conductiv_Coolingtime_Check = buf(i, 1) 'かつ数字
'        Else
'          Conductiv_Coolingtime_Check = buf(i, 1)  '文字列
'        End If
'      ElseIf buf(i, 1) = "" Then
'        Conductiv_Coolingtime_Check = "" '"LotNoは在るが写真確認Noがない"
'      End If
'    End If
'  Next i
'  If flag = False Then Conductiv_Coolingtime_Check = "LotNo記載がない"
'Else
'  KenBook.Close
'  MsgBox "指定シートがありません" & "_Conductiv_Coolingtime_Check", vbCritical
'  Conductiv_Coolingtime_Check = ""
'  Exit Function
'End If

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Function

Public Sub ConductivLoad() '導電率用紙読み込みと書き込み
'作業指示書:QA-021-08　対象焼き入れ材の導電率測定指示**2021.11.24
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Application.Visible = True 'Excel表示****非表示のままだと空枠のエクセルが発生する
Application.ScreenUpdating = False  '画面更新停止
Application.DisplayAlerts = False
Dim ListPass As String: ListPass = PackingPath
Dim ListName As String: ListName = "導電率測定結果記入表.xlsx"
If ReadCheck(ListPass & ListName, 1) = False Then
  MsgBox ListPass & ListName & "がありません" & "_NameConv", vbCritical
  Exit Sub
End If
Dim KenBook As Workbook: Set KenBook = Workbooks.Open(ListPass & ListName, UpdateLinks:=3)
Dim ws As Worksheet 'ﾜｰｸｼｰﾄ
Dim WsC As Boolean: WsC = False 'ｼｰﾄを探す
For Each ws In KenBook.Worksheets
  If ws.Name = "導電率測定" Then
    WsC = True
    Exit For
  End If
Next ws
Set ws = Nothing

Call DisplaySheets 'シート削除のためいったん表示

If WsC Then
  If WsCheck("導電率測定") Then ThisWorkbook.Sheets("導電率測定").Delete
  KenBook.Sheets("導電率測定").Copy After:=ThisWorkbook.Worksheets(Worksheets.Count)
Else
  MsgBox ListPass & "「導電率測定結果記入表」ｼｰﾄがありません。" & vbCrLf & "やり直してください。", vbExclamation
  Exit Sub
End If
KenBook.Close
Application.Visible = False 'Excel非表示

Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("導電率測定")
sh.Cells(4, 2) = Format$(Now, "yyyy/mm/dd")
sh.Cells(5, 2) = UF1.Lot
sh.Cells(6, 2) = UF1.Zai & "-" & UF1.Tyu
sh.Cells(7, 2) = UF1.ATU & "×" & UF1.Hab & "×" & UF1.Tak
sh.Cells(8, 2) = Lname

Dim wSh As Worksheet: Set wSh = ThisWorkbook.Sheets("検査員" & "_" & Lname) 'フォームリロード中の場合対策
Dim Worker
Select Case WorkeDayMax '登録時間が一番遅い
Case "1": Worker = Split(wSh.Cells(1, 2), "_")
Case "2": Worker = Split(wSh.Cells(2, 2), "_")
Case "3": Worker = Split(wSh.Cells(3, 2), "_")
End Select

sh.Cells(9, 2) = Worker(UBound(Worker)) '検査員
sh.Range("B4:B9").HorizontalAlignment = xlLeft

Call HiddenSheets
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Public Sub SelectionInput(TE As MSForms.TextBox, Optional n As String = "") '不良要因の入力
If n = "" Then
  If UF1.DefectFactorLB.Value = "" Then Exit Sub
Else
  If UFMitsubishi.DefectFactorLB.Value = "" Then Exit Sub
End If

Dim Lname As String, LotN_C As String, SelectedValue As String
If n = "" Then
  SelectedValue = GetSelectedItems(UF1.DefectFactorLB) ' 選択された値を取得
  Lname = UF1.LineName.Caption & UF1.LineNo.Caption
  LotN_C = UF1.Lot
Else
  SelectedValue = GetSelectedItems(UFMitsubishi.DefectFactorLB) ' 選択された値を取得
  Lname = UFMitsubishi.LineName.Caption
  LotN_C = UFMitsubishi.LotBox
End If

SkipRet:
If (WsCheck("不良要因印刷" & "_" & Lname) = False) Then
  Call AddWS("不良要因印刷" & "_" & Lname)
  Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
ElseIf ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Cells(1, 2) <> LotN_C Then
  Call AddWS("不良要因印刷" & "_" & Lname)
  Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
End If

'検査位置の転記
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname)
If n = "" Then
  sh.Cells(1, 2) = LotN_C
  sh.Cells(3, 2) = "'" & UF1.S.Caption
  sh.Cells(4, 2) = "'" & UF1.U.Caption
  sh.Cells(5, 2) = "'" & UF1.C.Caption
  sh.Cells(6, 2) = "'" & UF1.D.Caption
  sh.Cells(7, 2) = "'" & UF1.E.Caption
  sh.Cells(8, 2) = "'" & UF1.F.Caption
Else
  sh.Cells(1, 2) = LotN_C
End If


Dim bRange As Range: Set bRange = sh.Range(sh.Cells(3, 3), sh.Cells(9, 24))
Dim b_Arr(): b_Arr = bRange

'List_Data1 = Array("ﾌｸﾚ", " ", "ﾑｼﾚ", "汚", "ﾘｯｸｱｯﾌﾟ", " ", "圧込", "汚", "ﾌﾗｯﾄﾈｽ", "厚", " ", "汚", "飛込", "当り", _
'"擢", "付着油", "切断面", "巾丈", "特寸", "　", "他1", "TP")
'List_Data2 = Array("01", "02", "11", "12", "13", "14", "21", "22", "23", "24", "25", "31", "32", "33", _
'"34", "35", "36", "37", "38", "39", "41", "41")

If SelectedValue = "ｽﾎﾟｯﾄ処理" Then '2023.6.16_追加’2024.1.10_削除
  sh.Cells(11, 3) = TE & "枚"
  Exit Sub
End If

Dim Temp As String, SplitTemp: SplitTemp = Split(SelectedValue, "_") 'リストボックス内の不良要因とNOを分ける
If Not IsEmptyArray(SplitTemp) Then Temp = SplitTemp(UBound(SplitTemp)) '配列が空かチェック_NOだけ抜き取り

Dim a As Long, i As Long, itti As Long
Select Case Left(TE.Name, 1)
Case "S": a = 1
Case "U": a = 2
Case "C": a = 3
Case "D": a = 4
Case "E": a = 5
Case "F": a = 6
End Select

For i = LBound(b_Arr, 2) To UBound(b_Arr, 2) '列方向検索
  If b_Arr(UBound(b_Arr, 1), i) = Temp Then  'NO検索
    itti = i
    Exit For
  End If
Next i

If SelectedValue = "TP" Then '2024.1.11_TPの時はNOがないので指定の22(TPは41だが他41とわけている)
  b_Arr(a, 22) = TE
  bRange = b_Arr
Else
  b_Arr(a, itti) = TE
  bRange = b_Arr
End If

If SelectedValue = "TP" Then Exit Sub

If itti = 0 Then '必要なレイアウトが消えた状態なので消してやり直す_2025.7.9
  Application.DisplayAlerts = False
  sh.Delete
  Application.DisplayAlerts = True
  GoTo SkipRet
  Exit Sub
End If

End Sub

Public Sub TeColor(La As MSForms.TextBox, Optional n As String = "")
If n = "" Then
  UF1.SBad.BackColor = RGB(255, 255, 255)
  UF1.UBad.BackColor = RGB(255, 255, 255)
  UF1.CBad.BackColor = RGB(255, 255, 255)
  UF1.DBad.BackColor = RGB(255, 255, 255)
  UF1.EBad.BackColor = RGB(255, 255, 255)
  UF1.FBad.BackColor = RGB(255, 255, 255)
Else
  UFMitsubishi.SBad.BackColor = RGB(255, 255, 255)
  UFMitsubishi.UBad.BackColor = RGB(255, 255, 255)
  UFMitsubishi.CBad.BackColor = RGB(255, 255, 255)
  UFMitsubishi.DBad.BackColor = RGB(255, 255, 255)
  UFMitsubishi.EBad.BackColor = RGB(255, 255, 255)
  UFMitsubishi.FBad.BackColor = RGB(255, 255, 255)
End If

UF1(La.Name).BackColor = RGB(0, 255, 255) ' 水色' クリックされたラベルに水色をつける

End Sub
Public Function TecReturn(Optional n As String = "") As String '水色のTextBoxを返す
Dim i As Long
For i = 1 To 6
  If n = "" Then
    Select Case i
    Case 1: If UF1("SBad").BackColor = RGB(0, 255, 255) Then TecReturn = "SBad"
    Case 2: If UF1("UBad").BackColor = RGB(0, 255, 255) Then TecReturn = "UBad"
    Case 3: If UF1("CBad").BackColor = RGB(0, 255, 255) Then TecReturn = "CBad"
    Case 4: If UF1("DBad").BackColor = RGB(0, 255, 255) Then TecReturn = "DBad"
    Case 5: If UF1("EBad").BackColor = RGB(0, 255, 255) Then TecReturn = "EBad"
    Case 6: If UF1("FBad").BackColor = RGB(0, 255, 255) Then TecReturn = "FBad"
    End Select
  Else
    Select Case i
    Case 1: If UFMitsubishi("SBad").BackColor = RGB(0, 255, 255) Then TecReturn = "SBad"
    Case 2: If UFMitsubishi("UBad").BackColor = RGB(0, 255, 255) Then TecReturn = "UBad"
    Case 3: If UFMitsubishi("CBad").BackColor = RGB(0, 255, 255) Then TecReturn = "CBad"
    Case 4: If UFMitsubishi("DBad").BackColor = RGB(0, 255, 255) Then TecReturn = "DBad"
    Case 5: If UFMitsubishi("EBad").BackColor = RGB(0, 255, 255) Then TecReturn = "EBad"
    Case 6: If UFMitsubishi("FBad").BackColor = RGB(0, 255, 255) Then TecReturn = "FBad"
    End Select
  End If
Next i
End Function
Function FindRightString(InputString As String) As String
'右からANF,ANI,CAI,CAFを探して最終熱処理工程を割り出す
Dim StrArray() As String
Dim i As Long
StrArray = Split(InputString, " ") '文字列分割
For i = UBound(StrArray) To LBound(StrArray) Step -1 '右にある文字列から探す
  If StrArray(i) = "ANI" Or StrArray(i) = "ANF" Or StrArray(i) = "CAI" Or StrArray(i) = "CAF" Then
    FindRightString = StrArray(i) 'ヒット
    Exit Function
  End If
Next i
FindRightString = "" '該当する文字列が見つからなかった場合は空文字列を返す
End Function

Function CalculateSheets(FormTemp As Object, ThicknessA As Single) As Single '検入高さ計算
If ThicknessA = 0 Then Exit Function
Dim Quantity As Single
Dim Formula As String
Dim i As Long
' テキストボックスの値を取得
Dim TextBoxValues(1 To 6) As Variant
For i = 1 To 6
  TextBoxValues(i) = FormTemp("KenHBox" & i)
Next i

Dim Count As Long: Count = 0
' 数値が入力されたテキストボックスの数をカウント
For i = 1 To 6
  If IsNumeric(TextBoxValues(i)) Then
    Count = Count + 1
  End If
Next i

' 数値が入力されたテキストボックスの数が0の場合
If Count = 0 Then
  If FormTemp.Name = "UFMitsubishi" Then
    If FormTemp("KenHBox" & 7) <> "" Then
      FormTemp("KenHBox" & 7).Tag = FormTemp("KenHBox" & 7) 'Tag転記
    End If
    Exit Function
  Else
    Exit Function
  End If
End If

' 厚さAで割り算し枚数を算出
Quantity = 0

For i = 1 To 6
  If IsNumeric(TextBoxValues(i)) Then
    Quantity = Quantity + (TextBoxValues(i) / ThicknessA)
  End If
Next i

CalculateSheets = Format(Quantity, "0.0000")

' 計算式を作成
Formula = ""

For i = 1 To 6
  If IsNumeric(TextBoxValues(i)) Then
    Formula = Formula & "(" & TextBoxValues(i) & "÷" & ThicknessA & ")" & " + "
  End If
Next i

Formula = Left(Formula, Len(Formula) - 3) ' 最後の "+ " を削除
' 計算式をラベルAに設定
FormTemp.高さ計算式.Caption = Formula & "=" & Quantity

End Function
'tagデータ取り出し
Public Function UFtag_Array(x As Long, y As Long) As Variant
Dim i As Long, n As Long
ReDim DataArr(x To y, 1 To 2) As Variant
  With UF1
    For i = 1 To 2 '検査結果格納
      For n = x To y
        Select Case i
        Case 1
          Select Case n
          Case 1: DataArr(1, i) = .CBS1.Tag & "_" & .CBS2.Tag & "_" & .ZS.Tag
          Case 2: DataArr(2, i) = .CBU1.Tag & "_" & .CBU2.Tag & "_" & .ZU.Tag
          Case 3: DataArr(3, i) = .CBC1.Tag & "_" & .CBC2.Tag & "_" & .ZC.Tag
          Case 4: DataArr(4, i) = .CBD1.Tag & "_" & .CBD2.Tag & "_" & .ZD.Tag
          Case 5: DataArr(5, i) = .CBE1.Tag & "_" & .CBE2.Tag & "_" & .ZE.Tag
          Case 6: DataArr(6, i) = .CBF1.Tag & "_" & .CBF2.Tag & "_" & .ZF.Tag
          End Select
        Case 2
          Select Case n
          Case 1: DataArr(1, i) = .CBS3.Tag & "_" & .CBS4.Tag & "_" & .GS.Tag
          Case 2: DataArr(2, i) = .CBU3.Tag & "_" & .CBU4.Tag & "_" & .GU.Tag
          Case 3: DataArr(3, i) = .CBC3.Tag & "_" & .CBC4.Tag & "_" & .GC.Tag
          Case 4: DataArr(4, i) = .CBD3.Tag & "_" & .CBD4.Tag & "_" & .GD.Tag
          Case 5: DataArr(5, i) = .CBE3.Tag & "_" & .CBE4.Tag & "_" & .GE.Tag
          Case 6: DataArr(6, i) = .CBF3.Tag & "_" & .CBF4.Tag & "_" & .GF.Tag
          End Select
        End Select
      Next n
    Next i
  End With
  UFtag_Array = DataArr
'Stop
End Function
Public Function allow_Check() As String '許容フラグありの時位置を返す
'Dim a As Long
'For a = 1 To 5
'  Select Case a
'  Case 1: If UF1.NGS = "許容" Then allow_Check = UF1.S.Caption
'  Case 2: If UF1.NGU = "許容" Then allow_Check = UF1.U.Caption
'  Case 3: If UF1.NGC = "許容" Then allow_Check = UF1.C.Caption
'  Case 4: If UF1.NGD = "許容" Then allow_Check = UF1.D.Caption
'  Case 5: If UF1.NGE = "許容" Then allow_Check = UF1.E.Caption
'  End Select
'Next a

Dim Result As String, lblName As Variant, lbl As Control
Dim lblNames(): lblNames = Array("NGS", "NGU", "NGC", "NGD", "NGE", "NGF") ' 対象とするラベル名の配列を定義
For Each lblName In lblNames ' 配列をループして対象のラベルを調べる
    Set lbl = UF1.Controls(lblName)
    If InStr(lbl.Caption, "許容") > 0 Then
      If Result = "" Then ' 該当するラベルの名前をresultに追加
        Result = lbl.Name
      Else
        Result = Result & ", " & lbl.Name & " " & "許容"
      End If
    ElseIf InStr(lbl.Caption, "保留") > 0 Then
      If Result = "" Then
        Result = lbl.Name
      Else
        Result = Result & ", " & lbl.Name & " " & "保留"
      End If
    End If
Next lblName

Result = Replace(Result, "NGS", UF1.S.Caption)
Result = Replace(Result, "NGU", UF1.U.Caption)
Result = Replace(Result, "NGC", UF1.C.Caption)
Result = Replace(Result, "NGD", UF1.D.Caption)
Result = Replace(Result, "NGE", UF1.E.Caption)
Result = Replace(Result, "NGF", UF1.F.Caption)

allow_Check = Result ' 結果を返す

End Function
'2025.1.23
Public Function GetLabelState(FormName As UserForm, labelName As String) As String
'ラベルが Visible = True かつ Caption = "" の場合："否"
'ラベルが Visible = False の場合:"可"
'ラベルが Visible = True かつ Caption ≠ "" の場合:"可"
Dim ctrl As Control
On Error Resume Next
' 指定された名前のコントロールを取得
Set ctrl = FormName.Controls(labelName)
If Not ctrl Is Nothing Then
    If TypeName(ctrl) = "Label" Or TypeName(ctrl) = "TextBox" Then   'ラベルかテキストボックスであるか確認
        If ctrl.Visible = True Then
          If TypeName(ctrl) = "Label" Then
            If ctrl.Caption = "" Then
                GetLabelState = "否"
            Else
                GetLabelState = "可"
            End If
          ElseIf TypeName(ctrl) = "TextBox" Then
            If ctrl.Text = "" Then
                GetLabelState = "否"
            Else
                GetLabelState = "可"
            End If
          End If
        Else
            GetLabelState = "すべて非表示" 'Visible = False の場合
        End If
    Else
        GetLabelState = "無効" 'ラベル以外のコントロールの場合（補足）
    End If
Else
    GetLabelState = "無効" 'コントロールが存在しない場合
End If
On Error GoTo 0
End Function
'2025.1.23
Public Sub CheckLabelsByCategory(Category As String) '指定された入力欄を GetLabelState でチェック
Dim i As Long
Dim labelNames As Variant
Dim labelName As Variant
Dim state As String, Temp As String
Dim flag As Boolean: flag = False


Select Case Category
Case "S", "U", "C", "D", "E", "F"
  For i = 1 To 14 ' 番号付きラベル（例：S1～S14、U1～U14など）
    Temp = Category & i ' ラベル名を生成
    state = GetLabelState(UF1, Temp) ' ラベルの状態を取得
'    Debug.Print Temp & " の状態: " & state ' デバッグ出力
    If state = "可" Then
      UF1("CheckT" & Category).BackColor = RGB(0, 255, 0)
      flag = True
    ElseIf state = "すべて非表示" Then
      UF1("CheckT" & Category).BackColor = RGB(255, 255, 255)
    Else ' 否、無効が見つかったら処理を中断
    'F11で止まる
      If Category = "F" And i = 11 Then GoTo skipRun '2025.9.2
    
'      Debug.Print "処理を中断しました: " & Temp
      UF1("CheckT" & Category).BackColor = RGB(255, 0, 0)
      flag = False
      Exit Sub
    End If
  Next i
End Select


skipRun:
' 番号がつかないラベル（例：LateralS, DiagSKouなど）
Select Case Category
Case "S", "U", "C", "D", "E", "F"
  labelNames = Array("Lateral" & Category, "Diag" & Category & "Kou", Category & "_UpVisual", Category & "_UnderVisual")
Case Else
' MsgBox "無効なカテゴリが指定されました: " & Category, vbExclamation
'  Debug.Print "無効なカテゴリが指定されました: " & Category
  UF1("CheckT" & Category).BackColor = RGB(255, 0, 0)
  Exit Sub
End Select

' 配列のラベルを順番にチェック
For Each labelName In labelNames
    Temp = labelName
    state = GetLabelState(UF1, Temp) ' ラベルの状態を取得
'      Debug.Print Temp & " の状態: " & state ' デバッグ出力
    If state = "可" Then
      UF1("CheckT" & Category).BackColor = RGB(0, 255, 0)
      flag = True
    ElseIf state = "すべて非表示" Then
      UF1("CheckT" & Category).BackColor = RGB(255, 255, 255)
    Else
      UF1("CheckT" & Category).BackColor = RGB(255, 0, 0)
      flag = False
      Exit Sub
    End If
Next labelName

'歪の入力
Dim Ganswer As String: Ganswer = GaugeJudgment(Category)
If Ganswer = "" Then
  If flag Then UF1("CheckT" & Category).BackColor = RGB(0, 255, 0) '上記条件かつ歪入力あり
ElseIf Ganswer <> "" Then
  UF1("CheckT" & Category).BackColor = RGB(255, 0, 0)
End If
End Sub
'WithorWithout_Gaugeの改良版_2025.1.23
Public Function GaugeJudgment(Temp As String) As String
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim TempW As Double: TempW = Weight(val(UF1.ATU), val(UF1.Hab), val(UF1.Tak))
If TempW = 0 Then Exit Function ' 重量なし処理停止
If WsCheck("作業保存" & "_" & Lname) = False Then Exit Function ' シートなし処理停止

' ワークシート参照
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("作業保存" & "_" & Lname)
Dim A_range As Range: Set A_range = sh.Range(sh.Cells(2, 15), sh.Cells(7, 16)) ' A面
Dim B_range As Range: Set B_range = sh.Range(sh.Cells(8, 15), sh.Cells(13, 16)) ' B面

Dim A_Dis(): A_Dis = A_range.Value
Dim B_Dis(): B_Dis = B_range.Value

' 空データを整理
Dim i As Long
For i = LBound(A_Dis, 1) To UBound(A_Dis, 1)
    If A_Dis(i, 1) = "_" Or A_Dis(i, 1) = "__" Then A_Dis(i, 1) = ""
    If A_Dis(i, 2) = "_" Or A_Dis(i, 2) = "__" Then A_Dis(i, 2) = ""
Next i
For i = LBound(B_Dis, 1) To UBound(B_Dis, 1)
    If B_Dis(i, 1) = "_" Or B_Dis(i, 1) = "__" Then B_Dis(i, 1) = ""
    If B_Dis(i, 2) = "_" Or B_Dis(i, 2) = "__" Then B_Dis(i, 2) = ""
Next i

' 判定フラグ
Dim flagA As Boolean, flagB As Boolean
Dim flagAG As Boolean, flagBG As Boolean
Dim AtuF As Boolean: AtuF = False

' Temp に基づいて判定フラグを設定
Select Case Temp
Case "S": i = 1
Case "U": i = 2
Case "C": i = 3
Case "D": i = 4
Case "E": i = 5
Case "F": i = 6
Case Else: Exit Function ' 不正な引数の場合は終了
End Select

' フラグ設定
flagA = A_Dis(i, 1) <> "" And IsValidFormat(A_Dis(i, 1)) '文字_文字_数値　の判定
flagB = B_Dis(i, 1) <> "" And IsValidFormat(B_Dis(i, 1))
flagAG = A_Dis(i, 2) <> "" And IsValidFormat(A_Dis(i, 2))
flagBG = B_Dis(i, 2) <> "" And IsValidFormat(B_Dis(i, 2))

' 厚の測定フラグ設定
Select Case Temp
Case "S": AtuF = UF1.S1.Caption <> ""
Case "U": AtuF = UF1.U1.Caption <> ""
Case "C": AtuF = UF1.C1.Caption <> ""
Case "D": AtuF = UF1.D1.Caption <> ""
Case "E": AtuF = UF1.E1.Caption <> ""
Case "F": If UF1.Yama.Text = "4" Then AtuF = UF1.F1.Caption <> ""
End Select

' そもそも測定がない場合
If AtuF And Not flagA Then
    GaugeJudgment = Temp & "のA面 定盤 未入力"
    Exit Function
End If

' Q-S-QA-DIM-173*平坦度測定基準_パターンA（35kg以下のとき）
If AtuF And TempW <= 35 Then
    If flagA And Not flagB Then
        GaugeJudgment = Temp & "のB面 定盤 未入力_パターンA"
        Exit Function
    End If
    If flagB And Not flagA Then
        GaugeJudgment = Temp & "のA面 定盤 未入力_パターンA"
        Exit Function
    End If
End If

' Q-S-QA-DIM-173*平坦度測定基準_パターンB（35kg超えのとき）
If AtuF And TempW > 35 Then
    ' 定盤歪み測定が両面ある場合はスルー
    If flagA And flagB Then Exit Function

    ' 定盤Aのみ測定でゲージがない場合
    If flagA And Not flagB And Not flagAG Then
        GaugeJudgment = Temp & "のA面 ゲージ 未入力_パターンB"
        Exit Function
    End If

    ' 定盤Bのみ測定でゲージがない場合
    If flagB And Not flagA And Not flagBG Then
        GaugeJudgment = Temp & "のB面 ゲージ 未入力_パターンB"
        Exit Function
    End If
End If
End Function

Public Function IsValidFormat(ByVal inputStr As String) As Boolean '文字列_文字列_数値という形式であるかを判定
Dim parts() As String: parts = Split(inputStr, "_") ' _で文字列を分割

If UBound(parts) <> 2 Then ' 部分文字列が3つでなければFalse
   IsValidFormat = False
   Exit Function
End If

' どこかに入力があり、かつ未入力がある場合
If (parts(0) <> "" Or parts(1) <> "" Or parts(2) <> "") And _
  (parts(0) = "" Or parts(1) = "" Or parts(2) = "") Then
   ' 数値が"0.00"でない場合はFalse
   If parts(2) <> "0.00" Then
       IsValidFormat = False
       Exit Function
   End If
End If

' 各部分の条件を判定
Dim isNumPart0 As Boolean: isNumPart0 = IsNumeric(parts(0))
Dim isNumPart1 As Boolean: isNumPart1 = IsNumeric(parts(1))
Dim isNumPart2 As Boolean: isNumPart2 = IsNumeric(parts(2))

If isNumPart0 And isNumPart1 And Not isNumPart2 Then ' 数列, 数列, 文字列
   IsValidFormat = False
ElseIf Not isNumPart0 And Not isNumPart1 And Not isNumPart2 Then ' 文字列, 文字列, 文字列
   IsValidFormat = IsValidNumericWithSlash(parts(2)) ' "/" が含まれており、かつ "/" を除いた部分が数値であるかを判定
Else
   IsValidFormat = True
End If
End Function

Public Function IsValidNumericWithSlash(ByVal inputStr As String) As Boolean ' "/" が含まれており、かつ "/" を除いた部分が数値であるかを判定
' "/" が含まれているか確認
If InStr(inputStr, "/") = 0 Then
  IsValidNumericWithSlash = False
  Exit Function
End If
' "/" を取り除いた文字列を取得
Dim numericPart As String
numericPart = Replace(inputStr, "/", "")
' 残りの部分が数値かどうかを判定
If IsNumeric(numericPart) Then
  IsValidNumericWithSlash = True
Else
  IsValidNumericWithSlash = False
End If
End Function
'Distortion_Checkの改良_2025.1.23
Public Function SingleUnitDistortion(n As String, Temp As String) As String
SingleUnitDistortion = ""
With UF1
    ' 測定値0の時はスルー
    Select Case n
    Case "Z"
        Dim Value As String ' 数値以外も保持するため文字列型に変更
        Dim CB1 As String, CB2 As String

        ' Temp による判定対象の切り替え
        Select Case Temp
        Case "S"
            Value = CStr(.ZS): CB1 = .CBS1: CB2 = .CBS2
        Case "U"
            Value = CStr(.ZU): CB1 = .CBU1: CB2 = .CBU2
        Case "C"
            Value = CStr(.ZC): CB1 = .CBC1: CB2 = .CBC2
        Case "D"
            Value = CStr(.ZD): CB1 = .CBD1: CB2 = .CBD2
        Case "E"
            Value = CStr(.ZE): CB1 = .CBE1: CB2 = .CBE2
        Case "F"
            Value = CStr(.ZF): CB1 = .CBF1: CB2 = .CBF2
        End Select

        ' 判定処理 (空白を考慮)
        Dim Z_Tr As Boolean, Z_Fa As Boolean
        Z_Tr = (CB1 <> "" Or CB2 <> "" Or (Value <> "" And Value <> "0")) ' 入力がある
        Z_Fa = (CB1 = "" Or CB2 = "" Or Value = "")                      ' 未入力がある

        If Z_Tr And Z_Fa Then
            SingleUnitDistortion = "定盤:" & Temp
        End If

    Case "G"
        Dim GValue As String
        Dim GCB1 As String, GCB2 As String

        ' Temp による判定対象の切り替え
        Select Case Temp
        Case "S"
            GValue = CStr(.GS): GCB1 = .CBS3: GCB2 = .CBS4
        Case "U"
            GValue = CStr(.GU): GCB1 = .CBU3: GCB2 = .CBU4
        Case "C"
            GValue = CStr(.GC): GCB1 = .CBC3: GCB2 = .CBC4
        Case "D"
            GValue = CStr(.GD): GCB1 = .CBD3: GCB2 = .CBD4
        Case "E"
            GValue = CStr(.GE): GCB1 = .CBE3: GCB2 = .CBE4
        Case "F"
            GValue = CStr(.GF): GCB1 = .CBF3: GCB2 = .CBF4
        End Select

        ' 判定処理 (空白を考慮)
        Dim G_Tr As Boolean, G_Fa As Boolean
        G_Tr = (GCB1 <> "" Or GCB2 <> "" Or (GValue <> "" And GValue <> "0")) ' 入力がある
        G_Fa = (GCB1 = "" Or GCB2 = "" Or GValue = "")                        ' 未入力がある

        If G_Tr And G_Fa Then
            SingleUnitDistortion = "ｹﾞｰｼﾞ:" & Temp
        End If
    End Select
End With
End Function

