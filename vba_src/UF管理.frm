Attribute VB_Name = "UF管理"
Attribute VB_Base = "0{9B361EDC-AE3F-4125-B058-643404D8EAB5}{8702AB7A-BFA2-4BEF-8303-F0467A5BDA27}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
'ロットNO被り対策
Private Sub NewCB_Change() '2024.8.7
Dim i As Long, a As Long
Dim ListC As Long:          ListC = UF管理.NewCB.ListCount
Dim b As Long:              b = UF管理.NewCB.ListIndex
Dim sh As Worksheet

If Me.Mimi = False Then
  Set sh = ThisWorkbook.Sheets("保存用")
ElseIf Me.Mimi Then
  Set sh = ThisWorkbook.Sheets("耳用")
End If

Dim TargetText As String:   TargetText = Me.NewCB.Text
Dim EndR As Long:           EndR = UR(sh, 1)
Dim TargetRang As Range:    Set TargetRang = sh.Range(sh.Cells(1, 1), sh.Cells(EndR, 2))
Dim tMP():                  tMP = TargetRang
Me.CB1.Tag = "ループ"
'Me.CB1 = "" 'チェンジイベント発生のためリセット
For i = LBound(tMP, 1) To UBound(tMP, 1)
  If tMP(i, 1) = TargetText And tMP(i, 2) = Me.CB1.Text Then
    Me.CB1.ListIndex = i - 2
    Exit For
  End If
Next i
End Sub

Private Sub CB1_Change() '保存の確認

Dim i As Long, a As Long
Dim ListC As Long: ListC = UF管理.CB1.ListCount
Dim b As Long:     b = UF管理.CB1.ListIndex
Dim saveSh As Worksheet
Dim sh As Worksheet
Dim av(5), bv(5), VisRange As Range
Dim TempVis()
Dim splTemp() As String

Dim SearchString As String: SearchString = Me.CB1.Text
Dim Count As Long:          Count = 0

' コンボボックスのアイテムを検索
For i = 0 To Me.CB1.ListCount - 1
  If Me.CB1.List(i) = SearchString Then
    Count = Count + 1
  End If
Next i

If Count = 1 Then
 '2024.8.7
  Me.NewCB.clear
  Me.NewCB.BorderColor = &H80000006
  Me.NewCB.BackColor = &H80000005
End If

'2024.8.7
If Me.Mimi = False Then
  Set sh = ThisWorkbook.Sheets("保存用")
ElseIf Me.Mimi Then
  Set sh = ThisWorkbook.Sheets("耳用")
End If

Dim k As Long: k = sh.Cells(Rows.Count, 2).End(xlUp).row
If Count > 1 And Me.CB1.Tag <> "ループ" Then  '同LotNoあり
Me.NewCB.clear
  For a = 2 To k
    If sh.Cells(a, 2) = SearchString Then
      Me.NewCB.AddItem sh.Cells(a, 1) '時間
    End If
  Next a
End If

'リセット
  For i = 1 To 14
    UF管理("S" & i) = ""
    UF管理("U" & i) = ""
    UF管理("C" & i) = ""
    UF管理("D" & i) = ""
    UF管理("E" & i) = ""
    UF管理("F" & i) = ""
  Next i
  
  With UF管理
    .Slate = ""
    .Ulate = ""
    .Clate = ""
    .Dlate = ""
    .Elate = ""
    .Flate = ""
    
    .As_Vis = ""
    .Au_Vis = ""
    .Ac_Vis = ""
    .Ad_Vis = ""
    .Ae_Vis = ""
    .Af_Vis = ""
    
    .Bs_Vis = ""
    .Bu_Vis = ""
    .Bc_Vis = ""
    .Bd_Vis = ""
    .Be_Vis = ""
    .Bf_Vis = ""
    
    .S_BH = ""
    .U_BH = ""
    .C_BH = ""
    .D_BH = ""
    .E_BH = ""
    .F_BH = ""
    
    .SZH = ""
    .SGH = ""
    .UZH = ""
    .UGH = ""
    .CZH = ""
    .CGH = ""
    .DZH = ""
    .DGH = ""
    .EZH = ""
    .EGH = ""
    .FZH = ""
    .FGH = ""
    
    .SZH_B = ""
    .SGH_B = ""
    .UZH_B = ""
    .UGH_B = ""
    .CZH_B = ""
    .CGH_B = ""
    .DZH_B = ""
    .DGH_B = ""
    .EZH_B = ""
    .EGH_B = ""
    .FZH_B = ""
    .FGH_B = ""
    
    .S_Angle = ""
    .U_Angle = ""
    .C_Angle = ""
    .D_Angle = ""
    .E_Angle = ""
    .F_Angle = ""
    
  End With
If Count > 1 And Me.CB1.Tag <> "ループ" Then '2024.8.7
  MsgBox "同LotNoあり、作業時間を選択してください", vbExclamation
'  UF管理.Worker = ""
'  UF管理.Count = ""
'  UF管理.WorkJudgment = ""
'  UF管理.ATU = ""
'  UF管理.HABA = ""
'  UF管理.TAKE = ""
'  UF管理.AM = ""
'  UF管理.AP = ""
'  UF管理.HM = ""
'  UF管理.HP = ""
'  UF管理.TM = ""
'  UF管理.TP = ""
  Me.NewCB.BorderColor = RGB(255, 0, 0)
  Me.NewCB.BackColor = RGB(255, 255, 0)
  Me.CB1.Tag = ""
'  Exit Sub
End If

If Not Me.Mimi Then '通常品検索
  If Not WsCheck("保存用") Then
  Me.CB1.Tag = ""
  Exit Sub
  End If
  
  With UF管理
    .H1 = "A_定盤歪"
    .H2 = "A_ｹﾞｰｼﾞ歪"
    '2023.6.14
    .LhizB = "B_定盤歪"
    .LhigB = "B_ｹﾞｰｼﾞ歪"
    
    .A1 = "厚1"
    .A2 = "厚2"
    .A3 = "厚3"
    .A4 = "厚4"
    .L1 = "巾1"
    .L2 = "巾2"
    .L3 = "丈1"
    .L4 = "丈2"
    .L5 = "対角A"
    .L6 = "対角B"
    .L7 = ""
    .L8 = ""
    .Llate = "ﾗﾃﾗﾙﾎﾞｰ"
    
    If sh.Cells(b + 2, 84) = "〆" Then
      .H2 = "フリー反り"
      .LhigB = ""
    End If
    
    .SE = sh.Cells(b + 2, 1) '時間
    .Worker = sh.Cells(b + 2, 77) '検査員
    .Count = "作業数" & ListC & "件" '登録件数
    .WorkJudgment = sh.Cells(b + 2, 83) '作業選択
    .ATU = sh.Cells(b + 2, 3)
    .HABA = sh.Cells(b + 2, 4)
    .TAKE = sh.Cells(b + 2, 5)
    .Am = sh.Cells(b + 2, 6)
    .Ap = sh.Cells(b + 2, 7)
    .Hm = sh.Cells(b + 2, 8)
    .Hp = sh.Cells(b + 2, 9)
    .Tm = sh.Cells(b + 2, 10)
    .Tp = sh.Cells(b + 2, 11)
    
    If sh.Cells(b + 2, 22) <> "__" Then .SZH = sh.Cells(b + 2, 22)
    If sh.Cells(b + 2, 23) <> "__" Then .SGH = sh.Cells(b + 2, 23)
    If sh.Cells(b + 2, 34) <> "__" Then .UZH = sh.Cells(b + 2, 34)
    If sh.Cells(b + 2, 35) <> "__" Then .UGH = sh.Cells(b + 2, 35)
    If sh.Cells(b + 2, 46) <> "__" Then .CZH = sh.Cells(b + 2, 46)
    If sh.Cells(b + 2, 47) <> "__" Then .CGH = sh.Cells(b + 2, 47)
    If sh.Cells(b + 2, 58) <> "__" Then .DZH = sh.Cells(b + 2, 58)
    If sh.Cells(b + 2, 59) <> "__" Then .DGH = sh.Cells(b + 2, 59)
    If sh.Cells(b + 2, 70) <> "__" Then .EZH = sh.Cells(b + 2, 70)
    If sh.Cells(b + 2, 71) <> "__" Then .EGH = sh.Cells(b + 2, 71)
    If sh.Cells(b + 2, 110) <> "__" Then .FZH = sh.Cells(b + 2, 110)
    If sh.Cells(b + 2, 110) <> "__" Then .FGH = sh.Cells(b + 2, 111)
    
    '2023.6.14_歪AB対応
    Dim SplitArrZ, SplitArrG
    If InStr(.SZH.Caption, ",") > 0 Then
      SplitArrZ = Split(.SZH.Caption, ",")
      If InStr(SplitArrZ(LBound(SplitArrZ, 1)), "A:") > 0 Then
        .SZH = Replace(SplitArrZ(LBound(SplitArrZ, 1)), "A:", "")
      Else 'Aが空白でBがある時
        If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then .SZH = "" 'A面なし
      End If
      If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then .SZH_B = Replace(SplitArrZ(UBound(SplitArrZ, 1)), "B:", "")
      
      
      SplitArrG = Split(.SGH.Caption, ",")
      If InStr(SplitArrG(LBound(SplitArrG, 1)), "A:") > 0 Then
        .SGH = Replace(SplitArrG(LBound(SplitArrG, 1)), "A:", "")
      Else
        If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then .SGH = ""
      End If
      If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then .SGH_B = Replace(SplitArrG(UBound(SplitArrG, 1)), "B:", "")
    End If
    If InStr(.UZH.Caption, ",") > 0 Then
      SplitArrZ = Split(.UZH.Caption, ",")
      If InStr(SplitArrZ(LBound(SplitArrZ, 1)), "A:") > 0 Then
        .UZH = Replace(SplitArrZ(LBound(SplitArrZ, 1)), "A:", "")
      Else
        If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then .UZH = ""
      End If
      If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then .UZH_B = Replace(SplitArrZ(UBound(SplitArrZ, 1)), "B:", "")
      
      
      SplitArrG = Split(.UGH.Caption, ",")
      If InStr(SplitArrG(LBound(SplitArrG, 1)), "A:") > 0 Then
        .UGH = Replace(SplitArrG(LBound(SplitArrG, 1)), "A:", "")
      Else
        If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then .UGH = ""
      End If
      If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then .UGH_B = Replace(SplitArrG(UBound(SplitArrG, 1)), "B:", "")
    End If
    If InStr(.CZH.Caption, ",") > 0 Then
      SplitArrZ = Split(.CZH.Caption, ",")
      If InStr(SplitArrZ(LBound(SplitArrZ, 1)), "A:") > 0 Then
        .CZH = Replace(SplitArrZ(LBound(SplitArrZ, 1)), "A:", "")
      Else
        If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then .CZH = ""
      End If
      If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then .CZH_B = Replace(SplitArrZ(UBound(SplitArrZ, 1)), "B:", "")
      
      
      SplitArrG = Split(.CGH.Caption, ",")
      If InStr(SplitArrG(LBound(SplitArrG, 1)), "A:") > 0 Then
        .CGH = Replace(SplitArrG(LBound(SplitArrG, 1)), "A:", "")
      Else
        If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then .CGH = ""
      End If
      If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then .CGH_B = Replace(SplitArrG(UBound(SplitArrG, 1)), "B:", "")
    End If
    If InStr(.DZH.Caption, ",") > 0 Then
      SplitArrZ = Split(.DZH.Caption, ",")
      If InStr(SplitArrZ(LBound(SplitArrZ, 1)), "A:") > 0 Then
        .DZH = Replace(SplitArrZ(LBound(SplitArrZ, 1)), "A:", "")
      Else
        If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then .DZH = ""
      End If
      If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then .DZH_B = Replace(SplitArrZ(UBound(SplitArrZ, 1)), "B:", "")
      
      
      SplitArrG = Split(.DGH.Caption, ",")
      If InStr(SplitArrG(LBound(SplitArrG, 1)), "A:") > 0 Then
        .DGH = Replace(SplitArrG(LBound(SplitArrG, 1)), "A:", "")
      Else
        If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then .DGH = ""
      End If
      If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then .DGH_B = Replace(SplitArrG(UBound(SplitArrG, 1)), "B:", "")
    End If
    If InStr(.EZH.Caption, ",") > 0 Then
      SplitArrZ = Split(.EZH.Caption, ",")
      If InStr(SplitArrZ(LBound(SplitArrZ, 1)), "A:") > 0 Then
        .EZH = Replace(SplitArrZ(LBound(SplitArrZ, 1)), "A:", "")
      Else
        If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then .EZH = ""
      End If
      If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then .EZH_B = Replace(SplitArrZ(UBound(SplitArrZ, 1)), "B:", "")
      
      
      SplitArrG = Split(.EGH.Caption, ",")
      If InStr(SplitArrG(LBound(SplitArrG, 1)), "A:") > 0 Then
        .EGH = Replace(SplitArrG(LBound(SplitArrG, 1)), "A:", "")
      Else
        If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then .EGH = ""
      End If
      If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then .EGH_B = Replace(SplitArrG(UBound(SplitArrG, 1)), "B:", "")
    End If
    
    For i = 1 To 10
      For a = 1 To 5
        Select Case a
        Case 1: UF管理("S" & i) = sh.Cells(b + 2, i + 11)
        Case 2: UF管理("U" & i) = sh.Cells(b + 2, i + 23)
        Case 3: UF管理("C" & i) = sh.Cells(b + 2, i + 35)
        Case 4: UF管理("D" & i) = sh.Cells(b + 2, i + 47)
        Case 5: UF管理("E" & i) = sh.Cells(b + 2, i + 59)
        End Select
      Next a
    Next i
    UF管理.Slate = sh.Cells(b + 2, 72)
    UF管理.Ulate = sh.Cells(b + 2, 73)
    UF管理.Clate = sh.Cells(b + 2, 74)
    UF管理.Dlate = sh.Cells(b + 2, 75)
    UF管理.Elate = sh.Cells(b + 2, 76)
    
    Set VisRange = sh.Range(sh.Cells(b + 2, 85), sh.Cells(b + 2, 89)) '外観Range
    TempVis = VisRange
    
    Dim SixViRange As Range: Set SixViRange = sh.Cells(b + 2, 113) '外観Range　6回目の検査がある場合
    Dim SixTempVis(1):       SixTempVis(1) = SixViRange
    Dim SixSplTemp() As String
    
    If Not IsEmpty(TempVis(1, 1)) Or Not IsEmpty(TempVis(1, 2)) Or Not IsEmpty(TempVis(1, 3)) Or Not IsEmpty(TempVis(1, 4)) Or Not IsEmpty(TempVis(1, 5)) Then
      For i = LBound(TempVis, 2) To UBound(TempVis, 2)
        splTemp = Split(TempVis(1, i), ",")
        
        If Not IsEmptyArray(splTemp) Then '配列が空かどうかのチェック
          av(i - 1) = splTemp(LBound(splTemp, 1))
          bv(i - 1) = splTemp(UBound(splTemp, 1))
        End If
      Next i
      
      If Not IsEmpty(SixTempVis(1)) Then
        SixSplTemp = Split(SixTempVis(1), ",")
        If Not IsEmptyArray(SixSplTemp) Then
          av(UBound(av)) = SixSplTemp(LBound(SixSplTemp, 1))
          bv(UBound(bv)) = SixSplTemp(UBound(SixSplTemp, 1))
        End If
      End If
      
'      Stop
  
      UF管理.As_Vis = Replace(av(0), "A_", "")
      UF管理.Au_Vis = Replace(av(1), "A_", "")
      UF管理.Ac_Vis = Replace(av(2), "A_", "")
      UF管理.Ad_Vis = Replace(av(3), "A_", "")
      UF管理.Ae_Vis = Replace(av(4), "A_", "")
  
      UF管理.Bs_Vis = Replace(bv(0), "B_", "")
      UF管理.Bu_Vis = Replace(bv(1), "B_", "")
      UF管理.Bc_Vis = Replace(bv(2), "B_", "")
      UF管理.Bd_Vis = Replace(bv(3), "B_", "")
      UF管理.Be_Vis = Replace(bv(4), "B_", "")
  
      UF管理.S_BH = sh.Cells(b + 2, 90)
      UF管理.U_BH = sh.Cells(b + 2, 91)
      UF管理.C_BH = sh.Cells(b + 2, 92)
      UF管理.D_BH = sh.Cells(b + 2, 93)
      UF管理.E_BH = sh.Cells(b + 2, 94)
      
      UF管理.S_Angle = sh.Cells(b + 2, 95)
      UF管理.U_Angle = sh.Cells(b + 2, 96)
      UF管理.C_Angle = sh.Cells(b + 2, 97)
      UF管理.D_Angle = sh.Cells(b + 2, 98)
      UF管理.E_Angle = sh.Cells(b + 2, 99)
    End If
    
    Dim KenNoL As String: KenNoL = Left(sh.Cells(b + 2, 83), 1)
    Dim KenNoR As String: KenNoR = Right(sh.Cells(b + 2, 83), 1)
    
    If KenNoL = "H" And KenNoR = "数" Then
    .S = "1"
    .U = "2"
    .C = "3"
    .D = "4"
    .E = "5"
    .F = ""
    .SL = "1"
    .UL = "2"
    .CL = "3"
    .DL = "4"
    .EL = "5"
    .FL = ""
    ElseIf KenNoL = "H" And KenNoR = "H" Then
    .S = "S"
    .U = "S1"
    .C = "2"
    .D = "E1"
    .E = "E"
    .F = ""
    .SL = "S"
    .UL = "S1"
    .CL = "2"
    .DL = "E1"
    .EL = "E"
    .FL = ""
    ElseIf KenNoL <> "H" And KenNoR = "1" Then
    .S = "S"
    .U = ""
    .C = "2"
    .D = ""
    .E = "E"
    .F = ""
    .SL = "S"
    .UL = ""
    .CL = "2"
    .DL = ""
    .EL = "E"
    .FL = ""
    ElseIf KenNoL <> "H" And KenNoR = "2" Then
    .S = "2-S"
    .U = ""
    .C = "2-E"
    .D = ""
    .E = "1-E"
    .F = ""
    .SL = "S"
    .UL = ""
    .CL = "2-E"
    .DL = ""
    .EL = "1-E"
    .FL = ""
    ElseIf KenNoL <> "H" And KenNoR = "3" Then
    .S = "3-S"
    .U = "3-E"
    .C = "2-2"
    .D = "1-S"
    .E = "1-E"
    .F = ""
    .SL = "3-S"
    .UL = "3-E"
    .CL = "2-2"
    .DL = "1-S"
    .EL = "1-E"
    .FL = ""
    ElseIf KenNoL <> "H" And KenNoR = "4" Then
    .S = "4-S"
    .U = "4-E"
    .C = "3-2"
    .D = "2-2"
    .E = "1-S"
    .F = "1-E"
    .SL = "4-S"
    .UL = "4-E"
    .CL = "3-2"
    .DL = "2-2"
    .EL = "1-S"
    .FL = "1-E"
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>6回目がある用_2024.6.04
    If KenNoR = "4" Then
    Dim SixRange As Range: Set SixRange = sh.Range(sh.Cells(b + 2, 100), sh.Cells(b + 2, 114))
    '厚4，幅2,丈2,対角2,歪2,ラテ,外観,直角度
      For i = 1 To 10
          UF管理("F" & i) = sh.Cells(b + 2, i + 99)
      Next i
      UF管理.Flate = sh.Cells(b + 2, 112) 'ラテ
      UF管理.F_Angle = sh.Cells(b + 2, 114) '直角度
      UF管理.Af_Vis = Replace(av(5), "A_", "") '外観
      UF管理.Bf_Vis = Replace(bv(5), "B_", "")
      If InStr(.FZH.Caption, ",") > 0 Then
        SplitArrZ = Split(.FZH.Caption, ",")
        If InStr(SplitArrZ(LBound(SplitArrZ, 1)), "A:") > 0 Then
          .FZH = Replace(SplitArrZ(LBound(SplitArrZ, 1)), "A:", "")
        Else
          If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then .FZH = ""
        End If
        If InStr(SplitArrZ(UBound(SplitArrZ, 1)), "B:") > 0 Then .FZH_B = Replace(SplitArrZ(UBound(SplitArrZ, 1)), "B:", "")
        
        
        SplitArrG = Split(.FGH.Caption, ",")
        If InStr(SplitArrG(LBound(SplitArrG, 1)), "A:") > 0 Then
          .FGH = Replace(SplitArrG(LBound(SplitArrG, 1)), "A:", "")
        Else
          If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then .FGH = ""
        End If
        If InStr(SplitArrG(UBound(SplitArrG, 1)), "B:") > 0 Then .FGH_B = Replace(SplitArrG(UBound(SplitArrG, 1)), "B:", "")
      End If
    End If
  End With
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Else
  If Not WsCheck("耳用") Then
  Me.CB1.Tag = ""
  Exit Sub
  End If
  
  Set sh = ThisWorkbook.Sheets("耳用")
  Set VisRange = Range(sh.Cells(b + 2, 99), sh.Cells(b + 2, 103))
  TempVis = VisRange
  
  With UF管理
    .H1 = "定盤歪"
    .H2 = "ｹﾞｰｼﾞ歪"
    .A1 = "厚1"
    .A2 = "厚2"
    .A3 = "厚3"
    .A4 = "厚4"
    .L1 = "巾1"
    .L2 = "巾2"
    .L3 = "巾3"
    .L4 = "巾4"
    .L5 = "巾5"
    .L6 = "巾6"
    .L7 = "丈1"
    .L8 = "丈2"
    .L9 = "丈3"
    .L10 = "丈4"
    .SE = sh.Cells(b + 2, 1)
    .Worker = sh.Cells(b + 2, 92)
    .Count = "作業数" & ListC & "件"
    .WorkJudgment = sh.Cells(b + 2, 98) '作業選択
    .ATU = sh.Cells(b + 2, 3)
    .HABA = sh.Cells(b + 2, 4)
    .TAKE = sh.Cells(b + 2, 5)
    .Am = sh.Cells(b + 2, 6)
    .Ap = sh.Cells(b + 2, 7)
    .Hm = sh.Cells(b + 2, 8)
    .Hp = sh.Cells(b + 2, 9)
    .Tm = sh.Cells(b + 2, 10)
    .Tp = sh.Cells(b + 2, 11)
    
    .SZH = sh.Cells(b + 2, 26)
    .SGH = sh.Cells(b + 2, 27)
    .UZH = sh.Cells(b + 2, 42)
    .UGH = sh.Cells(b + 2, 43)
    .CZH = sh.Cells(b + 2, 58)
    .CGH = sh.Cells(b + 2, 59)
    .DZH = sh.Cells(b + 2, 74)
    .DGH = sh.Cells(b + 2, 75)
    .EZH = sh.Cells(b + 2, 90)
    .EGH = sh.Cells(b + 2, 91)
    For i = 1 To 14
      For a = 1 To 5
        Select Case a
        Case 1: UF管理("S" & i) = sh.Cells(b + 2, i + 11)
        Case 2: UF管理("U" & i) = sh.Cells(b + 2, i + 27)
        Case 3: UF管理("C" & i) = sh.Cells(b + 2, i + 43)
        Case 4: UF管理("D" & i) = sh.Cells(b + 2, i + 59)
        Case 5: UF管理("E" & i) = sh.Cells(b + 2, i + 75)
        End Select
      Next a
    Next i
  
    If Not IsEmpty(TempVis(1, 1)) Or Not IsEmpty(TempVis(1, 2)) Or Not IsEmpty(TempVis(1, 3)) Or Not IsEmpty(TempVis(1, 4)) Or Not IsEmpty(TempVis(1, 5)) Then
      For i = LBound(TempVis, 2) To UBound(TempVis, 2)
        splTemp = Split(TempVis(1, i), ",")
        If Not IsEmptyArray(splTemp) Then
          av(i - 1) = splTemp(LBound(splTemp, 1))
          bv(i - 1) = splTemp(UBound(splTemp, 1))
        End If
      Next i
  '    Stop
      UF管理("As_Vis") = Replace(av(0), "A_", "")
      UF管理("Au_Vis") = Replace(av(1), "A_", "")
      UF管理("Ac_Vis") = Replace(av(2), "A_", "")
      UF管理("Ad_Vis") = Replace(av(3), "A_", "")
      UF管理("Ae_Vis") = Replace(av(4), "A_", "")
  
      UF管理("Bs_Vis") = Replace(bv(0), "B_", "")
      UF管理("Bu_Vis") = Replace(bv(1), "B_", "")
      UF管理("Bc_Vis") = Replace(bv(2), "B_", "")
      UF管理("Bd_Vis") = Replace(bv(3), "B_", "")
      UF管理("Be_Vis") = Replace(bv(4), "B_", "")
  
      UF管理("S_BH") = sh.Cells(b + 2, 104)
      UF管理("U_BH") = sh.Cells(b + 2, 105)
      UF管理("C_BH") = sh.Cells(b + 2, 106)
      UF管理("D_BH") = sh.Cells(b + 2, 107)
      UF管理("E_BH") = sh.Cells(b + 2, 108)
    End If
    
    If (sh.Cells(b + 2, 98) = "極厚") Or (sh.Cells(b + 2, 98) = "") Then
    .S = "1"
    .U = "2"
    .C = "3"
    .D = "4"
    .E = "5"
    .SL = "S"
    .UL = "U"
    .CL = "2"
    .DL = "D"
    .EL = "E"
    End If
  End With
End If

Me.CB1.Tag = ""

End Sub

Private Sub CommandButton1_Click()
 If UF1.Visible And UF1.Lot <> "" Then
 MsgBox "別Lotデータがあります、新規登録処理を行ってください", vbCritical
Exit Sub
End If

'2024.11.06_同LotNo対策
Dim TempDate As String
If Me.NewCB.Text = "" Then Call HistoryLoad(Me.CB1.Text)
If Me.NewCB.Text <> "" Then
  TempDate = Me.NewCB.Text
  Call HistoryLoad(Me.CB1.Text, TempDate)
End If

Dim Temp As Variant:  Temp = InitialSetting(UF1, "サイズ保存")
End Sub

Private Sub CommandButton6_Click()
Unload Me
End Sub

Private Sub MIMI_AfterUpdate()
Dim a As Long
Dim k As Long
ThisWorkbook.Activate
If Me.Mimi Then
  UF管理.CB1.clear
  If WsCheck("耳用") Then
    k = ThisWorkbook.Sheets("耳用").Cells(Rows.Count, 2).End(xlUp).row
    For a = 2 To k
        UF管理.CB1.AddItem ThisWorkbook.Sheets("耳用").Cells(a, 2)
    Next a
  End If
Else
  UF管理.CB1.clear
  If WsCheck("保存用") Then
    k = ThisWorkbook.Sheets("保存用").Cells(Rows.Count, 2).End(xlUp).row
    For a = 2 To k
        UF管理.CB1.AddItem ThisWorkbook.Sheets("保存用").Cells(a, 2)
    Next a
  End If
  Me.L9 = ""
  Me.L10 = ""
End If
End Sub

Private Sub UserForm_Activate()
Call FormSetting
End Sub

Private Sub UserForm_Initialize()
Call GetZoom(Me)
Dim Temp As Variant:  Temp = InitialSetting(Me)
Dim a As Long
ThisWorkbook.Activate
UF管理.CB1.clear
If WsCheck("保存用") Then
  Dim k As Long: k = ThisWorkbook.Sheets("保存用").Cells(Rows.Count, 2).End(xlUp).row
  For a = 2 To k
      UF管理.CB1.AddItem ThisWorkbook.Sheets("保存用").Cells(a, 2)
  Next a
End If
End Sub
