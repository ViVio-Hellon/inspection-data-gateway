Attribute VB_Name = "LabelCls"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'--------UFﾄｯﾌﾟ記述
'Private LS(1 To 14) As New LabelCls
'Private LU(1 To 14) As New LabelCls
'Private LC(1 To 14) As New LabelCls
'Private LD(1 To 14) As New LabelCls
'Private LE(1 To 14) As New LabelCls
'--------Initialize
'Dim i As Long, a As Long
'For i = 1 To 14
'  For a = 1 To 5
'    Select Case a
'    Case 1: LS(i).EvS UF1("S" & i), i
'    Case 2: LU(i).EvU UF1("U" & i), i
'    Case 3: LC(i).EvC UF1("C" & i), i
'    Case 4: LD(i).EvD UF1("D" & i), i
'    Case 5: LE(i).EvE UF1("E" & i), i
'    End Select
'  Next a
'Next i
'*************************************************UF1_Label
Private WithEvents ObjS As MSForms.Label
Attribute ObjS.VB_VarHelpID = -1
Private WithEvents ObjU As MSForms.Label
Attribute ObjU.VB_VarHelpID = -1
Private WithEvents ObjC As MSForms.Label
Attribute ObjC.VB_VarHelpID = -1
Private WithEvents ObjD As MSForms.Label
Attribute ObjD.VB_VarHelpID = -1
Private WithEvents ObjE As MSForms.Label
Attribute ObjE.VB_VarHelpID = -1
Private WithEvents ObjF As MSForms.Label
Attribute ObjF.VB_VarHelpID = -1
'*************************************************UF1_ComboBox
Private WithEvents ObjCbS As MSForms.ComboBox
Attribute ObjCbS.VB_VarHelpID = -1
Private WithEvents ObjCbU As MSForms.ComboBox
Attribute ObjCbU.VB_VarHelpID = -1
Private WithEvents ObjCbC As MSForms.ComboBox
Attribute ObjCbC.VB_VarHelpID = -1
Private WithEvents ObjCbD As MSForms.ComboBox
Attribute ObjCbD.VB_VarHelpID = -1
Private WithEvents ObjCbE As MSForms.ComboBox
Attribute ObjCbE.VB_VarHelpID = -1
Private WithEvents ObjCbF As MSForms.ComboBox
Attribute ObjCbF.VB_VarHelpID = -1
'*************************************************UF_PackNo_Label
Private WithEvents ObjHiki As MSForms.Label
Attribute ObjHiki.VB_VarHelpID = -1
Private WithEvents ObjOrde As MSForms.Label
Attribute ObjOrde.VB_VarHelpID = -1
Private WithEvents ObjPack As MSForms.Label
Attribute ObjPack.VB_VarHelpID = -1
Private Cont As Long, PacCont As Long
' UF_PackNo_Label
Public Sub EvHiki(ByVal a As MSForms.Label, ByVal i As Long)
Set ObjHiki = a
PacCont = i
End Sub
Public Sub EvOrde(ByVal a As MSForms.Label, ByVal i As Long)
Set ObjOrde = a
PacCont = i
End Sub
Public Sub EvPack(ByVal a As MSForms.Label, ByVal i As Long)
Set ObjPack = a
PacCont = i
End Sub

Public Sub ObjHiki_Click()
Call Loop_Label(ObjHiki)
Call JPGLoad
End Sub
Public Sub ObjOrde_Click()
Call Loop_Label(ObjOrde)
Call JPGLoad
End Sub
Public Sub ObjPack_Click()
Call Loop_Label(ObjPack)
Call JPGLoad
End Sub
Public Sub ObjHiki_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call Loop_Label(ObjHiki)
Call JPGLoad
End Sub
Public Sub ObjOrde_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call Loop_Label(ObjOrde)
Call JPGLoad
End Sub
Public Sub ObjPack_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call Loop_Label(ObjPack)
Call JPGLoad
End Sub
'UF1_Label
Public Sub EvS(ByVal a As MSForms.Label, ByVal i As Long)
Set ObjS = a
Cont = i
End Sub
Public Sub EvU(ByVal a As MSForms.Label, ByVal i As Long)
Set ObjU = a
Cont = i
End Sub
Public Sub EvC(ByVal a As MSForms.Label, ByVal i As Long)
Set ObjC = a
Cont = i
End Sub
Public Sub EvD(ByVal a As MSForms.Label, ByVal i As Long)
Set ObjD = a
Cont = i
End Sub
Public Sub EvE(ByVal a As MSForms.Label, ByVal i As Long)
Set ObjE = a
Cont = i
End Sub
Public Sub EvF(ByVal a As MSForms.Label, ByVal i As Long)
Set ObjF = a
Cont = i
End Sub

Public Sub ObjS_Click()
Call TextTag(ObjS)
UF1.Manual.SetFocus
UF1.StartTB.SetFocus
'MsgBox "S"
End Sub
Public Sub ObjU_Click()
Call TextTag(ObjU)
UF1.Manual.SetFocus
UF1.StartTB.SetFocus
'MsgBox "U"
End Sub
Public Sub ObjC_Click()
Call TextTag(ObjC)
UF1.Manual.SetFocus
UF1.StartTB.SetFocus
'MsgBox "2"
End Sub
Public Sub ObjD_Click()
Call TextTag(ObjD)
UF1.Manual.SetFocus
UF1.StartTB.SetFocus
'MsgBox "D"
End Sub
Public Sub ObjE_Click()
Call TextTag(ObjE)
UF1.Manual.SetFocus
UF1.StartTB.SetFocus
'MsgBox "E"
End Sub
Public Sub ObjF_Click()
Call TextTag(ObjF)
UF1.Manual.SetFocus
UF1.StartTB.SetFocus
'MsgBox "F"
End Sub

'座標の保存とカウント**2023.4.24_追加
Private Sub ObjS_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim st As String:   st = UF1.INP.Tag
Dim Temp As String: Temp = UF1(st)

Dim NowTim As Date: NowTim = Format$(Now, "yyyy/mm/dd hh:mm:ss")
Call Time_Posting(st, NowTim, "") '測定時間と作業者の転写クリア
Call DiagonalErase(ObjS) '対角差ののﾘｾｯﾄ
ObjS.Caption = "" '消す
ObjS.Tag = "" 'NgPassのﾘｾｯﾄ

Call EraseCount(ObjS) '座標の保存とカウント

ObjS.BackColor = &H80000005 '背景色ﾘｾｯﾄ
Call NGErase(ObjS) 'NGの有無ﾌﾗｸﾞﾘｾｯﾄ

UF1.Manual.SetFocus
UF1.StartTB.SetFocus

Dim I_count As Long:  I_count = Len(st)
Dim pos As String:    pos = Left(st, 1)
Dim num As Long
Select Case I_count
Case 2: num = Right(st, 1)
Case 3: num = Right(st, 2)
End Select
'Q-S-QA-DIM-174_寸法検査結果の取り扱い基準*3回を超えて規格外れとなる場合は直ちに異常報告を行うこと_2023.7.19
'ここでは操業基準より厳しい2回で管理を行う_2023.8.15
'2024.2.6_1回目で理由入力必要と品証指示あり
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If (Temp <> "") And (val(UF1.DeleteC.Caption) >= 1) Then Call ConfiReason_Open(pos, num, Lname, Temp, val(UF1.DeleteC.Caption)) '空白でないかつ同じ座標を2回連続で消した
End Sub
Private Sub ObjU_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim st As String:   st = UF1.INP.Tag
Dim Temp As String: Temp = UF1(st)

Dim NowTim As Date: NowTim = Format$(Now, "yyyy/mm/dd hh:mm:ss")
Call Time_Posting(st, NowTim, "")
Call DiagonalErase(ObjU)
ObjU.Caption = ""
ObjU.Tag = ""

Call EraseCount(ObjU) '座標の保存とカウント

ObjU.BackColor = &H80000005
Call NGErase(ObjU)

UF1.Manual.SetFocus
UF1.StartTB.SetFocus

Dim I_count As Long:  I_count = Len(st)
Dim pos As String:    pos = Left(st, 1)
Dim num As Long
Select Case I_count
Case 2: num = Right(st, 1)
Case 3: num = Right(st, 2)
End Select

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If (Temp <> "") And (val(UF1.DeleteC.Caption) >= 1) Then Call ConfiReason_Open(pos, num, Lname, Temp, val(UF1.DeleteC.Caption))
End Sub
Private Sub ObjC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim st As String:   st = UF1.INP.Tag
Dim Temp As String: Temp = UF1(st)

Dim NowTim As Date: NowTim = Format$(Now, "yyyy/mm/dd hh:mm:ss")
Call Time_Posting(st, NowTim, "")
Call DiagonalErase(ObjC)
ObjC.Caption = ""
ObjC.Tag = ""

Call EraseCount(ObjC) '座標の保存とカウント

ObjC.BackColor = &H80000005
Call NGErase(ObjC)

UF1.Manual.SetFocus
UF1.StartTB.SetFocus

Dim I_count As Long:  I_count = Len(st)
Dim pos As String:    pos = Left(st, 1)
Dim num As Long
Select Case I_count
Case 2: num = Right(st, 1)
Case 3: num = Right(st, 2)
End Select

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If (Temp <> "") And (val(UF1.DeleteC.Caption) >= 1) Then Call ConfiReason_Open(pos, num, Lname, Temp, val(UF1.DeleteC.Caption))
End Sub
Private Sub ObjD_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim st As String:   st = UF1.INP.Tag
Dim Temp As String: Temp = UF1(st)

Dim NowTim As Date: NowTim = Format$(Now, "yyyy/mm/dd hh:mm:ss")
Call Time_Posting(st, NowTim, "")
Call DiagonalErase(ObjD)
ObjD.Caption = ""
ObjD.Tag = ""

Call EraseCount(ObjD) '座標の保存とカウント

ObjD.BackColor = &H80000005
Call NGErase(ObjD)

UF1.Manual.SetFocus
UF1.StartTB.SetFocus

Dim I_count As Long:  I_count = Len(st)
Dim pos As String:    pos = Left(st, 1)
Dim num As Long
Select Case I_count
Case 2: num = Right(st, 1)
Case 3: num = Right(st, 2)
End Select

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If (Temp <> "") And (val(UF1.DeleteC.Caption) >= 1) Then Call ConfiReason_Open(pos, num, Lname, Temp, val(UF1.DeleteC.Caption))
End Sub
Private Sub ObjE_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim st As String:   st = UF1.INP.Tag
Dim Temp As String: Temp = UF1(st)

Dim NowTim As Date: NowTim = Format$(Now, "yyyy/mm/dd hh:mm:ss")
Call Time_Posting(st, NowTim, "")
Call DiagonalErase(ObjE)
ObjE.Caption = ""
ObjE.Tag = ""

Call EraseCount(ObjE) '座標の保存とカウント

ObjE.BackColor = &H80000005
Call NGErase(ObjE)

UF1.Manual.SetFocus
UF1.StartTB.SetFocus

Dim I_count As Long:  I_count = Len(st)
Dim pos As String:    pos = Left(st, 1)
Dim num As Long
Select Case I_count
Case 2: num = Right(st, 1)
Case 3: num = Right(st, 2)
End Select

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If (Temp <> "") And (val(UF1.DeleteC.Caption) >= 1) Then Call ConfiReason_Open(pos, num, Lname, Temp, val(UF1.DeleteC.Caption))
End Sub
Private Sub ObjF_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim st As String:   st = UF1.INP.Tag
Dim Temp As String: Temp = UF1(st)

Dim NowTim As Date: NowTim = Format$(Now, "yyyy/mm/dd hh:mm:ss")
Call Time_Posting(st, NowTim, "")
Call DiagonalErase(ObjF)
ObjF.Caption = ""
ObjF.Tag = ""

Call EraseCount(ObjF) '座標の保存とカウント

ObjF.BackColor = &H80000005
Call NGErase(ObjF)

UF1.Manual.SetFocus
UF1.StartTB.SetFocus

Dim I_count As Long:  I_count = Len(st)
Dim pos As String:    pos = Left(st, 1)
Dim num As Long
Select Case I_count
Case 2: num = Right(st, 1)
Case 3: num = Right(st, 2)
End Select

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If (Temp <> "") And (val(UF1.DeleteC.Caption) >= 1) Then Call ConfiReason_Open(pos, num, Lname, Temp, val(UF1.DeleteC.Caption))
End Sub

'ドラッグ アンド ドロップ操作
'品証許可出ず_2022.04.25
Private Sub ObjS_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
If UF1.DBflag = "" Then Exit Sub '使用制限***2022.05.23
Dim D As DataObject
If Button <> 1 Then Exit Sub 'マウス左ボタンのドラッグ時に対応
Set D = New DataObject 'データオブジェクトに現在の選択地を格納
If ObjS.Caption <> "" Then
  ObjS.Tag = ObjS.Caption
  Dim n As String:  n = ObjS.Tag
  Dim TB As String: TB = ObjMoveText(ObjS)
  UF1.MoveFlag.Text = TB
  UF1.LabeTag.Text = n

  D.SetText ObjS.Caption 'オブジェクトにテキスト格納
  D.StartDrag 'ドラッグ開始
'  ObjS.Caption = "" '移動後は消す***ドロップまでMouseMove処理が続く
End If
End Sub
Private Sub ObjU_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
If UF1.DBflag = "" Then Exit Sub '使用制限***2022.05.23
Dim D As DataObject
If Button <> 1 Then Exit Sub
Set D = New DataObject
If ObjU.Caption <> "" Then
  ObjU.Tag = ObjU.Caption
  Dim n As String:  n = ObjU.Tag
  Dim TB As String: TB = ObjMoveText(ObjU)
  UF1.MoveFlag.Text = TB
  UF1.LabeTag.Text = n
  D.SetText ObjU.Caption
  D.StartDrag
End If
End Sub
Private Sub ObjC_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
If UF1.DBflag = "" Then Exit Sub '使用制限***2022.05.23
Dim D As DataObject
If Button <> 1 Then Exit Sub
Set D = New DataObject
If ObjC.Caption <> "" Then
  ObjC.Tag = ObjC.Caption
  Dim n As String:  n = ObjC.Tag
  Dim TB As String: TB = ObjMoveText(ObjC)
  UF1.MoveFlag.Text = TB
  UF1.LabeTag.Text = n
  D.SetText ObjC.Caption
  D.StartDrag
End If
End Sub
Private Sub ObjD_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
If UF1.DBflag = "" Then Exit Sub '使用制限***2022.05.23
Dim D As DataObject
If Button <> 1 Then Exit Sub
Set D = New DataObject
If ObjD.Caption <> "" Then
  ObjD.Tag = ObjD.Caption
  Dim n As String:  n = ObjD.Tag
  Dim TB As String: TB = ObjMoveText(ObjD)
  UF1.MoveFlag.Text = TB
  UF1.LabeTag.Text = n
  D.SetText ObjD.Caption
  D.StartDrag
End If
End Sub
Private Sub ObjE_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
If UF1.DBflag = "" Then Exit Sub '使用制限***2022.05.23
Dim D As DataObject
If Button <> 1 Then Exit Sub
Set D = New DataObject
If ObjE.Caption <> "" Then
  ObjE.Tag = ObjE.Caption
  Dim n As String:  n = ObjE.Tag
  Dim TB As String: TB = ObjMoveText(ObjE)
  UF1.MoveFlag.Text = TB
  UF1.LabeTag.Text = n
  D.SetText ObjE.Caption
  D.StartDrag
End If
End Sub
'○ドラッグ（In)
Private Sub ObjS_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal data As MSForms.DataObject, ByVal x As Single, ByVal y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
'マウスが入った時のイベント
'Cancel=TrueでDrag&Drop継続
If UF1.DBflag = "" Then Exit Sub '使用制限***2022.05.23
Dim Temp: Temp = data.GetText 'ドラッグ中のデータ
Dim st As String: st = UF1.MoveFlag.Text '転記したオブジェクト名
If ObjS.Caption = Temp Then
  Cancel = False '運ぶ先のデータが同じときドロップさせない
End If
If ObjS.Caption <> Temp Then '運ぶ先のデータが違う時
  Dim path As String: path = HistoryPath
  Select Case True
  Case UF1.L1Work:   Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "L-1\L-1_Log")
  Case UF1.LVC:      Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "LVC\LVC_Log")
  Case UF1.HVC:      Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "HVC\HVC_Log")
  Case UF1.Bala:     Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "バランサー\バランサー_Log")
  Case UF1.Tott:     Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "トット\トット_Log")
  Case UF1.Test:     Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "Test\Test_Log")
  Case Else:         Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "丸徳\丸徳_Log")
  End Select
  Cancel = True
  UF1(st).Caption = "" 'ドラッグ元データを消す
End If
End Sub
Private Sub ObjU_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal data As MSForms.DataObject, ByVal x As Single, ByVal y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
If UF1.DBflag = "" Then Exit Sub '使用制限***2022.05.23
Dim Temp: Temp = data.GetText
Dim st As String: st = UF1.MoveFlag.Text
If ObjU.Caption = Temp Then Cancel = False
If ObjU.Caption <> Temp Then
  Dim path As String: path = HistoryPath
  Select Case True
  Case UF1.L1Work:   Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "L-1\L-1_Log")
  Case UF1.LVC:      Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "LVC\LVC_Log")
  Case UF1.HVC:      Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "HVC\HVC_Log")
  Case UF1.Bala:     Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "バランサー\バランサー_Log")
  Case UF1.Tott:     Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "トット\トット_Log")
  Case UF1.Test:     Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "Test\Test_Log")
  Case Else:         Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "丸徳\丸徳_Log")
  End Select
  Cancel = True
  UF1(st).Caption = ""
End If
End Sub
Private Sub ObjC_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal data As MSForms.DataObject, ByVal x As Single, ByVal y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
If UF1.DBflag = "" Then Exit Sub '使用制限***2022.05.23
Dim Temp: Temp = data.GetText
Dim st As String: st = UF1.MoveFlag.Text
If ObjC.Caption = Temp Then Cancel = False
If ObjC.Caption <> Temp Then
  Dim path As String: path = HistoryPath
  Select Case True
  Case UF1.L1Work:   Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "L-1\L-1_Log")
  Case UF1.LVC:      Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "LVC\LVC_Log")
  Case UF1.HVC:      Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "HVC\HVC_Log")
  Case UF1.Bala:     Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "バランサー\バランサー_Log")
  Case UF1.Tott:     Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "トット\トット_Log")
  Case UF1.Test:     Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "Test\Test_Log")
  Case Else:         Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "丸徳\丸徳_Log")
  End Select
  Cancel = True
  UF1(st).Caption = ""
End If
End Sub
Private Sub ObjD_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal data As MSForms.DataObject, ByVal x As Single, ByVal y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
If UF1.DBflag = "" Then Exit Sub '使用制限***2022.05.23
Dim Temp: Temp = data.GetText
Dim st As String: st = UF1.MoveFlag.Text
If ObjD.Caption = Temp Then Cancel = False
If ObjD.Caption <> Temp Then
  Dim path As String: path = HistoryPath
  Select Case True
  Case UF1.L1Work:   Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "L-1\L-1_Log")
  Case UF1.LVC:      Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "LVC\LVC_Log")
  Case UF1.HVC:      Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "HVC\HVC_Log")
  Case UF1.Bala:     Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "バランサー\バランサー_Log")
  Case UF1.Tott:     Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "トット\トット_Log")
  Case UF1.Test:     Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "Test\Test_Log")
  Case Else:         Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "丸徳\丸徳_Log")
  End Select
  Cancel = True
  UF1(st).Caption = ""
End If
End Sub
Private Sub ObjE_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal data As MSForms.DataObject, ByVal x As Single, ByVal y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
If UF1.DBflag = "" Then Exit Sub '使用制限***2022.05.23
Dim Temp: Temp = data.GetText
Dim st As String: st = UF1.MoveFlag.Text
If ObjE.Caption = Temp Then Cancel = False
If ObjE.Caption <> Temp Then
  Dim path As String: path = HistoryPath
  Select Case True
  Case UF1.L1Work:   Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "L-1\L-1_Log")
  Case UF1.LVC:      Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "LVC\LVC_Log")
  Case UF1.HVC:      Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "HVC\HVC_Log")
  Case UF1.Bala:     Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "バランサー\バランサー_Log")
  Case UF1.Tott:     Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "トット\トット_Log")
  Case UF1.Test:     Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "Test\Test_Log")
  Case Else:         Call WriteLog("*UFKen:" & "DragOver:" & "座標_" & st & " Data_" & Temp, "", path & "丸徳\丸徳_Log")
  End Select
  Cancel = True
  UF1(st).Caption = ""
End If
End Sub
'○ドロップ
Private Sub ObjS_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal data As MSForms.DataObject, ByVal x As Single, ByVal y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
'ドラッグ時のみドラッグされたデータを追加
If UF1.DBflag = "" Then Exit Sub '使用制限***2022.05.23
If Action = fmActionDragDrop Then
  If MsgBox("Dataをドロップしますか？", vbYesNo + vbQuestion, "確認") = vbYes Then
    Dim path As String: path = HistoryPath
    Select Case True
    Case UF1.L1Work:   Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjS.Name & ":元data_" & ObjS.Caption & " 後data_" & data.GetText, "", path & "L-1\L-1_Log")
    Case UF1.LVC:      Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjS.Name & ":元data_" & ObjS.Caption & " 後data_" & data.GetText, "", path & "LVC\LVC_Log")
    Case UF1.HVC:      Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjS.Name & ":元data_" & ObjS.Caption & " 後data_" & data.GetText, "", path & "HVC\HVC_Log")
    Case UF1.Bala:     Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjS.Name & ":元data_" & ObjS.Caption & " 後data_" & data.GetText, "", path & "バランサー\バランサー_Log")
    Case UF1.Tott:     Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjS.Name & ":元data_" & ObjS.Caption & " 後data_" & data.GetText, "", path & "トット\トット_Log")
    Case UF1.Test:     Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjS.Name & ":元data_" & ObjS.Caption & " 後data_" & data.GetText, "", path & "Test\Test_Log")
    Case Else:         Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjS.Name & ":元data_" & ObjS.Caption & " 後data_" & data.GetText, "", path & "丸徳\丸徳_Log")
    End Select
    ObjS.Caption = data.GetText 'Log出力後実行
  End If
End If
data.clear 'DataObjectのデータクリア
End Sub
Private Sub ObjU_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal data As MSForms.DataObject, ByVal x As Single, ByVal y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
If UF1.DBflag = "" Then Exit Sub '使用制限***2022.05.23
If Action = fmActionDragDrop Then
  If MsgBox("Dataをドロップしますか？", vbYesNo + vbQuestion, "確認") = vbYes Then
    Dim path As String: path = HistoryPath
    Select Case True
    Case UF1.L1Work:   Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjU.Name & ":元data_" & ObjU.Caption & " 後data_" & data.GetText, "", path & "L-1\L-1_Log")
    Case UF1.LVC:      Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjU.Name & ":元data_" & ObjU.Caption & " 後data_" & data.GetText, "", path & "LVC\LVC_Log")
    Case UF1.HVC:      Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjU.Name & ":元data_" & ObjU.Caption & " 後data_" & data.GetText, "", path & "HVC\HVC_Log")
    Case UF1.Bala:     Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjU.Name & ":元data_" & ObjU.Caption & " 後data_" & data.GetText, "", path & "バランサー\バランサー_Log")
    Case UF1.Tott:     Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjU.Name & ":元data_" & ObjU.Caption & " 後data_" & data.GetText, "", path & "トット\トット_Log")
    Case UF1.Test:     Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjU.Name & ":元data_" & ObjU.Caption & " 後data_" & data.GetText, "", path & "Test\Test_Log")
    Case Else:         Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjU.Name & ":元data_" & ObjU.Caption & " 後data_" & data.GetText, "", path & "丸徳\丸徳_Log")
    End Select
    ObjU.Caption = data.GetText
  End If
End If
data.clear
End Sub
Private Sub ObjC_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal data As MSForms.DataObject, ByVal x As Single, ByVal y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
If UF1.DBflag = "" Then Exit Sub '使用制限***2022.05.23
If Action = fmActionDragDrop Then
  If MsgBox("Dataをドロップしますか？", vbYesNo + vbQuestion, "確認") = vbYes Then
    Dim path As String: path = HistoryPath
    Select Case True
    Case UF1.L1Work:   Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjC.Name & ":元data_" & ObjC.Caption & " 後data_" & data.GetText, "", path & "L-1\L-1_Log")
    Case UF1.LVC:      Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjC.Name & ":元data_" & ObjC.Caption & " 後data_" & data.GetText, "", path & "LVC\LVC_Log")
    Case UF1.HVC:      Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjC.Name & ":元data_" & ObjC.Caption & " 後data_" & data.GetText, "", path & "HVC\HVC_Log")
    Case UF1.Bala:     Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjC.Name & ":元data_" & ObjC.Caption & " 後data_" & data.GetText, "", path & "バランサー\バランサー_Log")
    Case UF1.Tott:     Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjC.Name & ":元data_" & ObjC.Caption & " 後data_" & data.GetText, "", path & "トット\トット_Log")
    Case UF1.Test:     Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjC.Name & ":元data_" & ObjC.Caption & " 後data_" & data.GetText, "", path & "Test\Test_Log")
    Case Else:         Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjC.Name & ":元data_" & ObjC.Caption & " 後data_" & data.GetText, "", path & "丸徳\丸徳_Log")
    End Select
    ObjC.Caption = data.GetText
  End If
End If
data.clear
End Sub
Private Sub ObjD_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal data As MSForms.DataObject, ByVal x As Single, ByVal y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
If UF1.DBflag = "" Then Exit Sub '使用制限***2022.05.23
If Action = fmActionDragDrop Then
  If MsgBox("Dataをドロップしますか？", vbYesNo + vbQuestion, "確認") = vbYes Then
    Dim path As String: path = HistoryPath
    Select Case True
    Case UF1.L1Work:   Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjD.Name & ":元data_" & ObjD.Caption & " 後data_" & data.GetText, "", path & "L-1\L-1_Log")
    Case UF1.LVC:      Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjD.Name & ":元data_" & ObjD.Caption & " 後data_" & data.GetText, "", path & "LVC\LVC_Log")
    Case UF1.HVC:      Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjD.Name & ":元data_" & ObjD.Caption & " 後data_" & data.GetText, "", path & "HVC\HVC_Log")
    Case UF1.Bala:     Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjD.Name & ":元data_" & ObjD.Caption & " 後data_" & data.GetText, "", path & "バランサー\バランサー_Log")
    Case UF1.Tott:     Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjD.Name & ":元data_" & ObjD.Caption & " 後data_" & data.GetText, "", path & "トット\トット_Log")
    Case UF1.Test:     Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjD.Name & ":元data_" & ObjD.Caption & " 後data_" & data.GetText, "", path & "Test\Test_Log")
    Case Else:         Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjD.Name & ":元data_" & ObjD.Caption & " 後data_" & data.GetText, "", path & "丸徳\丸徳_Log")
    End Select
    ObjD.Caption = data.GetText
  End If
End If
data.clear
End Sub
Private Sub ObjE_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal data As MSForms.DataObject, ByVal x As Single, ByVal y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
If UF1.DBflag = "" Then Exit Sub '使用制限***2022.05.23
If Action = fmActionDragDrop Then
  If MsgBox("Dataをドロップしますか？", vbYesNo + vbQuestion, "確認") = vbYes Then
    Dim path As String: path = HistoryPath
    Select Case True
    Case UF1.L1Work:   Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjE.Name & ":元data_" & ObjE.Caption & " 後data_" & data.GetText, "", path & "L-1\L-1_Log")
    Case UF1.LVC:      Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjE.Name & ":元data_" & ObjE.Caption & " 後data_" & data.GetText, "", path & "LVC\LVC_Log")
    Case UF1.HVC:      Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjE.Name & ":元data_" & ObjE.Caption & " 後data_" & data.GetText, "", path & "HVC\HVC_Log")
    Case UF1.Bala:     Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjE.Name & ":元data_" & ObjE.Caption & " 後data_" & data.GetText, "", path & "バランサー\バランサー_Log")
    Case UF1.Tott:     Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjE.Name & ":元data_" & ObjE.Caption & " 後data_" & data.GetText, "", path & "トット\トット_Log")
    Case UF1.Test:     Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjE.Name & ":元data_" & ObjE.Caption & " 後data_" & data.GetText, "", path & "Test\Test_Log")
    Case Else:         Call WriteLog("*UFKen:" & "DropOrPaste:" & "移動先_" & ObjE.Name & ":元data_" & ObjE.Caption & " 後data_" & data.GetText, "", path & "丸徳\丸徳_Log")
    End Select
    ObjE.Caption = data.GetText
  End If
End If
data.clear
End Sub

'ComboBox ObjCbS
Public Sub EvComS(ByVal a As MSForms.ComboBox, ByVal i As Long)
Set ObjCbS = a
Cont = i
End Sub
Public Sub EvComU(ByVal a As MSForms.ComboBox, ByVal i As Long)
Set ObjCbU = a
Cont = i
End Sub
Public Sub EvComC(ByVal a As MSForms.ComboBox, ByVal i As Long)
Set ObjCbC = a
Cont = i
End Sub
Public Sub EvComD(ByVal a As MSForms.ComboBox, ByVal i As Long)
Set ObjCbD = a
Cont = i
End Sub
Public Sub EvComE(ByVal a As MSForms.ComboBox, ByVal i As Long)
Set ObjCbE = a
Cont = i
End Sub
Public Sub EvComF(ByVal a As MSForms.ComboBox, ByVal i As Long)
Set ObjCbF = a
Cont = i
End Sub

'ドロップダウン リストを展開でoff　Exitでonにしている
Private Sub ObjCbS_DropButtonClick()
UF1.Load.Caption = "off"
End Sub
Private Sub ObjCbU_DropButtonClick()
UF1.Load.Caption = "off"
End Sub
Private Sub ObjCbC_DropButtonClick()
UF1.Load.Caption = "off"
End Sub
Private Sub ObjCbD_DropButtonClick()
UF1.Load.Caption = "off"
End Sub
Private Sub ObjCbE_DropButtonClick()
UF1.Load.Caption = "off"
End Sub
Private Sub ObjCbF_DropButtonClick()
UF1.Load.Caption = "off"
End Sub

Private Sub ObjCbS_Change()
If UF1.Loading.Tag = "Loading" Then Exit Sub '履歴呼び出しフラグ_印刷転送で使用
If UF1.FlagLabel = 0 Then Exit Sub

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim fname As String: fname = "Measured.log"
Dim path As String:  path = HistoryPath
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

Dim Temp As String: Temp = Right(UF1.AorB.Caption, 1)
If ObjCbS.Text <> "" Then
  Call WriteLog("*UFKen:" & "[" & Temp & "] " & ObjCbS.Name & ":" & ObjCbS.Text & ":" & Lname, "", Lpath, fname)
  Call WriteLog("*UFKen:" & "[" & Temp & "] " & ObjCbS.Name & ":" & ObjCbS.Text & ":" & Lname, "")
End If
End Sub
Private Sub ObjCbU_Change()
If UF1.Loading.Tag = "Loading" Then Exit Sub '履歴呼び出しフラグ_印刷転送で使用
If UF1.FlagLabel = 0 Then Exit Sub

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim fname As String: fname = "Measured.log"
Dim path As String:  path = HistoryPath
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

Dim Temp As String: Temp = Right(UF1.AorB.Caption, 1)
If ObjCbU.Text <> "" Then
  Call WriteLog("*UFKen:" & "[" & Temp & "] " & ObjCbU.Name & ":" & ObjCbU.Text & ":" & Lname, "", Lpath, fname)
  Call WriteLog("*UFKen:" & "[" & Temp & "] " & ObjCbU.Name & ":" & ObjCbU.Text & ":" & Lname, "")
End If
End Sub
Private Sub ObjCbC_Change()
If UF1.Loading.Tag = "Loading" Then Exit Sub '履歴呼び出しフラグ_印刷転送で使用
If UF1.FlagLabel = 0 Then Exit Sub

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim fname As String: fname = "Measured.log"
Dim path As String:  path = HistoryPath
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

Dim Temp As String: Temp = Right(UF1.AorB.Caption, 1)
If ObjCbC.Text <> "" Then
  Call WriteLog("*UFKen:" & "[" & Temp & "] " & ObjCbC.Name & ":" & ObjCbC.Text & ":" & Lname, "", Lpath, fname)
  Call WriteLog("*UFKen:" & "[" & Temp & "] " & ObjCbC.Name & ":" & ObjCbC.Text & ":" & Lname, "")
End If
End Sub
Private Sub ObjCbD_Change()
If UF1.Loading.Tag = "Loading" Then Exit Sub '履歴呼び出しフラグ_印刷転送で使用
If UF1.FlagLabel = 0 Then Exit Sub

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim fname As String: fname = "Measured.log"
Dim path As String:  path = HistoryPath
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

Dim Temp As String: Temp = Right(UF1.AorB.Caption, 1)
If ObjCbD.Text <> "" Then
  Call WriteLog("*UFKen:" & "[" & Temp & "] " & ObjCbD.Name & ":" & ObjCbD.Text & ":" & Lname, "", Lpath, fname)
  Call WriteLog("*UFKen:" & "[" & Temp & "] " & ObjCbD.Name & ":" & ObjCbD.Text & ":" & Lname, "")
End If
End Sub
Private Sub ObjCbE_Change()
If UF1.Loading.Tag = "Loading" Then Exit Sub '履歴呼び出しフラグ_印刷転送で使用
If UF1.FlagLabel = 0 Then Exit Sub

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim fname As String: fname = "Measured.log"
Dim path As String:  path = HistoryPath
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

Dim Temp As String: Temp = Right(UF1.AorB.Caption, 1)
If ObjCbE.Text <> "" Then
  Call WriteLog("*UFKen:" & "[" & Temp & "] " & ObjCbE.Name & ":" & ObjCbE.Text & ":" & Lname, "", Lpath, fname)
  Call WriteLog("*UFKen:" & "[" & Temp & "] " & ObjCbE.Name & ":" & ObjCbE.Text & ":" & Lname, "")
End If
End Sub
Private Sub ObjCbF_Change()
If UF1.Loading.Tag = "Loading" Then Exit Sub '履歴呼び出しフラグ_印刷転送で使用
If UF1.FlagLabel = 0 Then Exit Sub

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim fname As String: fname = "Measured.log"
Dim path As String:  path = HistoryPath
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

Dim Temp As String: Temp = Right(UF1.AorB.Caption, 1)
If ObjCbF.Text <> "" Then
  Call WriteLog("*UFKen:" & "[" & Temp & "] " & ObjCbF.Name & ":" & ObjCbF.Text & ":" & Lname, "", Lpath, fname)
  Call WriteLog("*UFKen:" & "[" & Temp & "] " & ObjCbF.Name & ":" & ObjCbF.Text & ":" & Lname, "")
End If
End Sub

