Attribute VB_Name = "UF1"
Attribute VB_Base = "0{1821599B-7B7B-4038-9BC8-0933FABADFD0}{00516896-C01D-4742-81EC-4C3097293802}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Old_CBS1 As String
Private Old_CBU1 As String
Private Old_CBC1 As String
Private Old_CBD1 As String
Private Old_CBE1 As String
Private Old_CBF1 As String

Private Old_CBS2 As String
Private Old_CBU2 As String
Private Old_CBC2 As String
Private Old_CBD2 As String
Private Old_CBE2 As String
Private Old_CBF2 As String

Private Old_CBS3 As String
Private Old_CBU3 As String
Private Old_CBC3 As String
Private Old_CBD3 As String
Private Old_CBE3 As String
Private Old_CBF3 As String

Private Old_CBS4 As String
Private Old_CBU4 As String
Private Old_CBC4 As String
Private Old_CBD4 As String
Private Old_CBE4 As String
Private Old_CBF4 As String

' Windows API定数定義
Private Const CF_BITMAP = 2
Private Const CF_PALETTE = 9
Private Const CF_ENHMETAFILE = 14

' Windows API関数宣言
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PictDesc, riid As GUID, ByVal fOwn As Long, ipic As IPicture) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PictDesc, riid As GUID, ByVal fOwn As Long, ipic As IPicture) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' 構造体定義
#If VBA7 Then
    Private Type PictDesc
        cbSizeofStruct As Long
        picType As Long
        hImage As LongPtr
        Option1 As LongPtr
    End Type
#Else
    Private Type PictDesc
        cbSizeofStruct As Long
        picType As Long
        hImage As Long
        Option1 As Long
    End Type
#End If

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


'ﾗﾍﾞﾙ
Private LS(1 To 14) As New LabelCls '耳付き丈3,4追加分+2_2022.06
Private LU(1 To 14) As New LabelCls
Private LC(1 To 14) As New LabelCls
Private LD(1 To 14) As New LabelCls
Private LE(1 To 14) As New LabelCls
Private LF(1 To 12) As New LabelCls
' ｺﾝﾎﾞﾎﾞｯｸｽ
Private ComS(1 To 4) As New LabelCls
Private ComU(1 To 4) As New LabelCls
Private ComC(1 To 4) As New LabelCls
Private ComD(1 To 4) As New LabelCls
Private ComE(1 To 4) As New LabelCls
Private ComF(1 To 4) As New LabelCls
'点滅用
Private BlinkLabel1   As clsBlinkLabel
' Excelウィンドウ状態退避用
Private g_sngTopSV As Single ' 上端位置
Private g_sngLeftSV As Single ' 左端位置
Private g_sngWidthSV As Single ' 幅
Private g_sngHeightSV As Single ' 高さ
' ウィンドウ状態
Private g_lngWindowState As XlWindowState
Dim Lname As String
Private SikaAtuAll As Variant
Private SikaSetAtu As Single
Private NowSizeH As Long
Private NowSizeW As Long
Private NowSizeMP As Long

Private Sub blindfm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call frameblindfold("通常")
End Sub


Private Sub CommandButton18_Click()
Dim st As String:     st = UF1.INP.Tag
Dim LeftSt As String: LeftSt = Left(st, 1)

Select Case st
Case "LateralS": LeftSt = "S"
Case "LateralU": LeftSt = "U"
Case "LateralC": LeftSt = "C"
Case "LateralD": LeftSt = "D"
Case "LateralE": LeftSt = "E"
Case "LateralF": LeftSt = "F"
'直角度
Case "DiagSKou": LeftSt = "S"
Case "DiagUKou": LeftSt = "U"
Case "DiagCKou": LeftSt = "C"
Case "DiagDKou": LeftSt = "D"
Case "DiagEKou": LeftSt = "E"
Case "DiagFKou": LeftSt = "F"
'歪A
Case "ZS": LeftSt = "S"
Case "ZU": LeftSt = "U"
Case "ZC": LeftSt = "C"
Case "ZD": LeftSt = "D"
Case "ZE": LeftSt = "E"
Case "ZF": LeftSt = "F"
'歪ｹﾞｰｼﾞ
Case "GS": LeftSt = "S"
Case "GU": LeftSt = "U"
Case "GC": LeftSt = "C"
Case "GD": LeftSt = "D"
Case "GE": LeftSt = "E"
Case "GF": LeftSt = "F"
End Select

If LeftSt = "" Then
  MsgBox "保留の根拠位置をクリックしてください。"
  Exit Sub
End If

Dim Temp As String
Select Case LeftSt
Case "S": Temp = UF1.S.Caption
Case "U": Temp = UF1.U.Caption
Case "C": Temp = UF1.C.Caption
Case "D": Temp = UF1.D.Caption
Case "E": Temp = UF1.E.Caption
Case "F": Temp = UF1.F.Caption
End Select

If MsgBox(Temp & "で保留処理を行いますか？", vbYesNo) = vbYes Then Call NG_Pass(LeftSt, "保留")
End Sub


Private Sub UserForm_Initialize()
If ThisWorkbook.Sheets("全ヵ所検査").Visible Or ThisWorkbook.Sheets("キャスク検査").Visible Then
  If Not ((UFCask.Visible = True) Or (UFMitsubishi.Visible = True)) Then MsgBox "Loadを失敗しました。" & vbCrLf & "フォームを閉じて起動しなおしてください。", vbExclamation
  ThisWorkbook.Sheets("全ヵ所検査").Visible = False
  ThisWorkbook.Sheets("キャスク検査").Visible = False
  Exit Sub
End If

' 印刷履歴の初期化
Call InitializePrintFlags

' 30日以上前の古い履歴を削除
Call CleanOldHistory
    
'2024.1.19_ﾃﾞﾌｫﾙﾄｻｲｽ変更
UF1.Height = 545 'Heightﾃﾞﾌｫﾙﾄｻｲｽﾞ_変更時はDefaultFormPosition,VC_Clear,UF1_Clickを修正の事
UF1.Width = 887.25 '2021.06.15_サイズ変更

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
UF1.Auto = "on" 'ﾃﾞﾌｫ
UF1.Tag = "STOP" '音楽ｽﾄｯﾌﾟ
UF1.UFTag = UF1.Tag '目視用
UF1.Load = "off" 'ﾌｫｰﾑｲﾍﾞﾝﾄ停止
UF1.FlagLabel = 1 'Changeｲﾍﾞﾝﾄ用
UF1.INP.Tag = ""
UF1.PicFM.Tag = 0 '2023.6.27**LoopJudg用カウント
UF1.CB.Tag = ""

NowSizeH = UF1.MPEX.Height '2024.5.30_追加
NowSizeW = UF1.MPEX.Width
NowSizeMP = UF1.MP1.Width

UFProgress.Show vbModeless
DoEvents

UFProgress.CommandButton1.Caption = "ClsData_Loading...."
UFProgress.Repaint
Call ProgressBar_count(1, 580)

'ｵｰﾄﾒｰｼｮﾝｴﾗｰ対策,Activateで実行とする_2022.06.23

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>検査ライン設定_2022.11.14
Dim tmpFaci As String: tmpFaci = GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")
Dim tmpMode As String: tmpMode = GetSetting("UF検査判定システム", "起動設定", "モード", "存在なし")
If tmpMode = "オフライン" Then UF1.Caption = UF1.Caption & tmpFaci & "  " & "「オフライン」"
If tmpMode = "オンライン" Then UF1.Caption = UF1.Caption & tmpFaci & "  " & "「オンライン」"

If tmpFaci = "存在なし" Then
 MsgBox "ライン設定がありません" & vbCrLf & "担当者に連絡して下さい", vbExclamation
 End
End If

Call Set_Shift '時間セット_2024.3.29
Call Visual_Add '外観ﾘｽﾄ'---------------先Loadで使う
Call GetZoom(Me) '2024.5.29

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'ﾗｲﾝ選択 Load_Text,WorkerLoad,LoopJudg '作業保存読み込み,検査員読み込み,検査判定
'2022.11.14_OPをtrueにすることでLine_AutoSelectionが実行されるのでここではOPの操作のみ行う
Select Case tmpFaci
Case "L-1":        Me.L1Work = True
Case "LVC":        Me.LVC = True
Case "HVC":        Me.HVC = True
Case "バランサー": Me.Bala = True
Case "トット":     Me.Tott = True
Case "Test":       Me.Test = True
Case "丸徳":       Me.Maru = True
End Select
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'先にLoad_Textをしよう(フォームに数値反映)
'オーダー厚のセット_2024.5.14
Dim LotNo As String: LotNo = UF1.Lot.Text

'印刷履歴の確認と印刷ボタン色の設定
Call CheckAndUpdateButtonColor

If SikaSetAtu = 0 Then
  SikaAtuAll = SikaAtuCreate(LotNo)
  If Not IsEmpty(SikaAtuAll) Then SikaSetAtu = SikaAtuAll(LBound(SikaAtuAll))
End If
If SikaSetAtu <> 0 Then
  If UF1.blindfold.Tag <> "Set" Then 'まだ追加されていない状態
    Call CBAdd(SikaSetAtu) '歪ｺﾝﾎﾞ
  Else 'Lselctでフラグ すでに追加済み　combに先入れがあったらスルー_2024.4.22
    UF1.blindfold.Tag = ""
  End If
Else
  Call CBAdd
End If

Call ProgressBar_count(1, 633)


'Call Line_AutoSelection 'ﾗｲﾝ選択

UF1.Tag = "" '音楽ｽﾀｰﾄ
UF1.UFTag = UF1.Tag
'ラベル反映 クラス
Dim i As Long, a As Long
For i = 1 To 14 '丈追加分+2_2022.06
  LS(i).EvS UF1("S" & i), i
  LU(i).EvU UF1("U" & i), i
  LC(i).EvC UF1("C" & i), i
  LD(i).EvD UF1("D" & i), i
  LE(i).EvE UF1("E" & i), i
Next i

For i = 1 To 10 '2024.5.21
  LF(i).EvF UF1("F" & i), i
Next i

Call ProgressBar_count(633, 780)
'ｺﾝﾎﾞﾎﾞｯｸｽ反映　ｸﾗｽ
For i = 1 To 4
  For a = 1 To 6
    Select Case a
    Case 1: ComS(i).EvComS UF1("CBS" & i), i
    Case 2: ComU(i).EvComU UF1("CBU" & i), i
    Case 3: ComC(i).EvComC UF1("CBC" & i), i
    Case 4: ComD(i).EvComD UF1("CBD" & i), i
    Case 5: ComE(i).EvComE UF1("CBE" & i), i
    Case 6: ComF(i).EvComF UF1("CBF" & i), i '2024.5.21
    End Select
  Next a
Next i
' 現在のExcelウィンドウ状態を保管
'With Application
'  .WindowState = xlNormal
'  g_sngTopSV = .Top
'  g_sngLeftSV = .Left
'  g_sngWidthSV = .Width
'  g_sngHeightSV = .Height
'End With

Call ProgressBar_count(780, 920)


Dim Comb As Variant
Me.Yama.clear
Comb = Array("1", "2", "3", "4", "5", "6", "7")
Me.Yama.List = Comb

DoEvents

UF1.Worker.BackColor = &H80000005 'ﾃﾞﾌｫ
Set BlinkLabel1 = New clsBlinkLabel '点滅ラベルの登録
BlinkLabel1.cmd = Me.AtButt


Call ProgressBar_count(920, 1000)
DoEvents

Dim Temp As Variant:  Temp = InitialSetting(UF1, "リスタート")

UF1.Load = "on"
Unload UFProgress
End Sub
Private Sub NGS_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If MsgBox("まとめて削除しますか?", vbYesNo) = vbYes Then
  Lname = UF1.LineName.Caption & UF1.LineNo.Caption
  Call ConfiReason_Open("S", 0, Lname, "S板1枚", val(UF1.DeleteC.Caption)) 'Logには位置0_:理由選択:板廃棄:検査員名で出力
End If
End Sub
Private Sub NGU_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If MsgBox("まとめて削除しますか?", vbYesNo) = vbYes Then
  Lname = UF1.LineName.Caption & UF1.LineNo.Caption
  Call ConfiReason_Open("U", 0, Lname, "U板1枚", val(UF1.DeleteC.Caption))
End If
End Sub
Private Sub NGC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If MsgBox("まとめて削除しますか?", vbYesNo) = vbYes Then
  Lname = UF1.LineName.Caption & UF1.LineNo.Caption
  Call ConfiReason_Open("C", 0, Lname, "C板1枚", val(UF1.DeleteC.Caption))
End If
End Sub
Private Sub NGD_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If MsgBox("まとめて削除しますか?", vbYesNo) = vbYes Then
  Lname = UF1.LineName.Caption & UF1.LineNo.Caption
  Call ConfiReason_Open("D", 0, Lname, "D板1枚", val(UF1.DeleteC.Caption))
End If
End Sub
Private Sub NGE_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If MsgBox("まとめて削除しますか?", vbYesNo) = vbYes Then
   Lname = UF1.LineName.Caption & UF1.LineNo.Caption
  Call ConfiReason_Open("E", 0, Lname, "E板1枚", val(UF1.DeleteC.Caption))
End If
End Sub
Private Sub NGF_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If MsgBox("まとめて削除しますか?", vbYesNo) = vbYes Then
   Lname = UF1.LineName.Caption & UF1.LineNo.Caption
  Call ConfiReason_Open("F", 0, Lname, "F板1枚", val(UF1.DeleteC.Caption))
End If
End Sub

Private Sub A1_Click()
Select Case True
Case UF1.Maru: Call Line_AutoSelection("表示") '丸徳の時のみ表示
Case Else:     Call Line_AutoSelection
End Select
End Sub
Private Sub A2_Click()
Select Case True
Case UF1.Maru: Call Line_AutoSelection("表示") '丸徳の時のみ表示
Case Else:     Call Line_AutoSelection
End Select
End Sub
Private Sub A3_Click()
Select Case True
Case UF1.Maru: Call Line_AutoSelection("表示") '丸徳の時のみ表示
Case Else:     Call Line_AutoSelection
End Select
End Sub
Private Sub A4_Click()
Select Case True
Case UF1.Maru: Call Line_AutoSelection("表示") '丸徳の時のみ表示
Case Else:     Call Line_AutoSelection
End Select
End Sub
Private Sub A5_Click()
Select Case True
Case UF1.Maru: Call Line_AutoSelection("表示") '丸徳の時のみ表示
Case Else:     Call Line_AutoSelection
End Select
End Sub
Private Sub A6_Click()
Select Case True
Case UF1.Maru: Call Line_AutoSelection("表示") '丸徳の時のみ表示
Case Else:     Call Line_AutoSelection
End Select
End Sub
Private Sub A7_Click()
Select Case True
Case UF1.Maru: Call Line_AutoSelection("表示") '丸徳の時のみ表示
Case Else:     Call Line_AutoSelection
End Select
End Sub
Private Sub A8_Click()
Select Case True
Case UF1.Maru: Call Line_AutoSelection("表示") '丸徳の時のみ表示
Case Else:     Call Line_AutoSelection
End Select
End Sub
Private Sub A9_Click()
Select Case True
Case UF1.Maru: Call Line_AutoSelection("表示") '丸徳の時のみ表示
Case Else:     Call Line_AutoSelection
End Select
End Sub
Private Sub A10_Click()
Select Case True
Case UF1.Maru: Call Line_AutoSelection("表示") '丸徳の時のみ表示
Case Else:     Call Line_AutoSelection
End Select
End Sub
Private Sub A11_Click()
Select Case True
Case UF1.Maru: Call Line_AutoSelection("表示") '丸徳の時のみ表示
Case Else:     Call Line_AutoSelection
End Select
End Sub

Private Sub Abno_Click()
On Error Resume Next
Select Case UF1.Abno.Caption
Case "異常値on": UF1.Abno.Caption = "異常値off"
Case "異常値off": UF1.Abno.Caption = "異常値on"
End Select
End Sub

Private Sub Am_Change()
If IsNumeric(UF1.Am) = False Then UF1.Am = ""
Call Form_Formula
End Sub

Private Sub AMode_Click()
Dim i As Long, FL As Boolean, KenM As String
For i = 1 To 4
  If UF1("Inspection" & i) = True Then
    KenM = UF1("Inspection" & i).Caption
    FL = True
    Exit For
  End If
Next i
If FL = False Then
  MsgBox "検査法を選択してください", vbCritical
  Exit Sub
End If
If UF1.Load.Caption = "off" Then Exit Sub
Select Case True
Case UF1.Test '無条件起動
  Call Ainput_Startup("UF1", KenM)
Case UF1.DBflag <> "" '2021.03.11_無条件起動
  Call Ainput_Startup("UF1", KenM)
Case Else
  With UF1
    Select Case True
    Case (.AMin.Caption <> "" And .AMax.Caption <> "" And .HMin.Caption <> "" And .HMax.Caption <> "" And .TMin.Caption <> "" And .TMax.Caption <> "")
      Call Ainput_Startup("UF1", KenM) '起動ﾌｫｰﾑ、検査法
    Case Else
    MsgBox "MinMax処理を行ってください", vbInformation
    End Select
  End With
End Select
End Sub
Private Sub AorB_Click()
If UF1.blindfold.Visible Then Exit Sub

UF1.Loading.Tag = "Loading" 'Access.log停止
'AorB.Tag**ﾃﾞﾌｫﾙﾄA
'A,B面の判定
If Me.AorB.Caption = "Now_A" Then
  Me.AorB.Caption = "Now_B"
  Me.AorB.BackColor = RGB(205, 0, 0)
  Me.MPEX.BackColor = RGB(0, 0, 255) ' 青色
  MsgBox "B面歪 入力画面です"
ElseIf Me.AorB.Caption = "Now_B" Then
  Me.AorB.Caption = "Now_A"
  Me.AorB.BackColor = &H8000000F
  Me.MPEX.BackColor = &H8000000F
End If

'表示非表示
Dim flag(1 To 6) As Boolean
UF1.Loading.Visible = True
UF1.HizCome.Visible = False

UF1.blindfold.tOP = 0
UF1.blindfold.Left = 0
UF1.blindfold.Height = 215
UF1.blindfold.Width = 388
UF1.blindfold.Visible = True

If UF1.FMS.Visible Then flag(1) = True
If UF1.FMS.Visible Then UF1.FMS.Visible = False
If UF1.FMU.Visible Then flag(2) = True
If UF1.FMU.Visible Then UF1.FMU.Visible = False
If UF1.FMC.Visible Then flag(3) = True
If UF1.FMC.Visible Then UF1.FMC.Visible = False
If UF1.FMD.Visible Then flag(4) = True
If UF1.FMD.Visible Then UF1.FMD.Visible = False
If UF1.FME.Visible Then flag(5) = True
If UF1.FME.Visible Then UF1.FME.Visible = False
If UF1.FMF.Visible Then flag(6) = True
If UF1.FMF.Visible Then UF1.FMF.Visible = False

UF1.Repaint
Sleep 1000

Dim n As String: n = Right(Me.AorB.Caption, 1)
Call ReflectStrain(n)

If flag(1) Then UF1.FMS.Visible = True
If flag(2) Then UF1.FMU.Visible = True
If flag(3) Then UF1.FMC.Visible = True
If flag(4) Then UF1.FMD.Visible = True
If flag(5) Then UF1.FME.Visible = True
If flag(6) Then UF1.FMF.Visible = True

UF1.HizCome.Visible = True
UF1.blindfold.Visible = False
UF1.Loading.Visible = False

UF1.Tag = "STOP" '音楽ｽﾄｯﾌﾟ
UF1.UFTag = UF1.Tag '目視用

Call SetSpecificTextBoxesBackColorToWhite '背景色リセット

Call HizLoopJudg(0)
UF1.Tag = ""
UF1.UFTag = UF1.Tag
UF1.Repaint
UF1.Loading.Tag = ""
End Sub

Private Sub Ap_Change()
If IsNumeric(UF1.Ap) = False Then UF1.Ap = ""
Call Form_Formula
End Sub
Private Sub Atu_Change()
If IsNumeric(UF1.ATU) = False Then UF1.ATU = ""
End Sub
Private Sub Byousya_AfterUpdate()
On Error Resume Next
UF1.厚板 = False
UF1.薄板 = False
UF1.少枚数 = False
UF1.極厚 = False

Call Six_Display("on")
Call LSelectLateral("強制", "6回目も表示")
Call RightAngleDisplay("強制", "6回目も表示")
Dim Temp As Variant:  Temp = InitialSetting(UF1)

End Sub
Private Sub Cle_Click()
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If (WsCheck("不良要因印刷" & "_" & Lname) = False) Then Exit Sub

Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname)
Dim bRange As Range: Set bRange = sh.Range(sh.Cells(3, 3), sh.Cells(8, 24))
Dim aRange As Range: Set aRange = sh.Range(sh.Cells(10, 3), sh.Cells(10, 24))
Dim cRange As Range: Set cRange = sh.Cells(11, 3)
bRange.ClearContents
aRange.ClearContents
cRange.ClearContents
End Sub
Private Sub CommandButton15_Click()
Me.Loading.Visible = True
Me.Repaint

Call AccesslogLoad(Me.LoadLot, Me.LoadWorker, Me.LoadTextName, Me.Reason)
Call ReDiag '対角再計算(線差)

Me.Loading.Visible = False
Me.Repaint
End Sub


Private Sub Conductiv_Click()
If UF1.Conductiv.BackColor = RGB(255, 0, 0) Then
  If MsgBox("導電率測定用紙を印刷しますか？", vbYesNo) = vbYes Then
    Call ConductivLoad
    Application.PrintCommunication = False
    Call DisplaySheets
    ThisWorkbook.Sheets("導電率測定").Visible = True
    ThisWorkbook.Sheets("導電率測定").PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
    Call HiddenSheets
  End If
End If
End Sub

Private Sub H494B_Click()
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If Register_Check(Format$(Now, "hh:mm:ss"), "") Then '検査員ﾁｪｯｸ *1.ｼｰﾄがない2.3行(Day1,Day2,Day3)に当日以外の登録がある3.その時間に登録が無い
  Call MsgError("検査員登録をして下さい", Lname, 2)
  Exit Sub
End If
Application.Visible = True

Me.Loading.Visible = True
Me.Repaint
UF1.Yama.Text = ""

Application.ScreenUpdating = False
Call HiddenSheets

Call BookSaClose(ThisWorkbook.path, "日報＆資材計算.xlsm") 'ﾊﾟｽ,閉じたいbook名

Call OpenMitubisi("H494")

If UFWorkerName.Visible = True Then Unload UFWorkerName

Me.Loading.Visible = False
Me.Repaint
UF1.Hide
End Sub

Private Sub Image1_Click()
Dim i As Long
With UF1
  .Image1.Height = 1
  .Image1.Width = 1
  .Lot.Visible = True
  .MPEX.Visible = True
  .MP1.Visible = True
  
  .MPEX.Left = 66
  .MP1.Left = 468
  
  .LZTKlabel.Visible = True
  
  .FS.Visible = True
  .FU.Visible = True
  .FC.Visible = True
  .FD.Visible = True
  .FE.Visible = True
  If .Yama.Text = "4" Then .FF.Visible = True

  'ものによっては再描写できなくなるので停止_2024.5.15
''If UF1.Sumisei.Caption <> "" Then
'    UF1.L11.Visible = True
'    UF1.LateralS.Visible = True
'    UF1.LateralU.Visible = True
'    UF1.LateralC.Visible = True
'    UF1.LateralD.Visible = True
'    UF1.LateralE.Visible = True
''  If UF1.Youto = "R115" Then 'QA-S-QA-DIM-154_対角線差及び直角度の測定基準
'    UF1.L10.Visible = True
'    UF1.DiagSKou.Visible = True
'    UF1.DiagUKou.Visible = True
'    UF1.DiagCKou.Visible = True
'    UF1.DiagDKou.Visible = True
'    UF1.DiagEKou.Visible = True
''  End If
''End If
  
  For i = 1 To 8
    UF1("Frm" & i).Visible = True
  Next i
  
    Dim TempArry As Variant: TempArry = GetLotInfo("Lot情報履歴", UF1.Lot.Text) 'ヒットしたロットの配列を作成
    If IsEmpty(TempArry) Then Exit Sub

    '2026.10.14_引当混在対策：共通判定
    Dim 検番 As String
    Dim NeedLateral As Boolean, NeedRight As Boolean, NeedFlat As Boolean
    Call JudgeLotAllocations(TempArry, 検番, NeedLateral, NeedRight, NeedFlat)
    If NeedLateral Then Call LSelectLateral(検番)
    If NeedRight Then Call RightAngleDisplay(検番)
  
  .Image1.Visible = False
  .Repaint
End With

'追加修正_20251.10
Select Case UF1.Additional.Caption
Case "対角ﾗﾃﾗﾙﾎﾞ-追記公差あり": Call LSelectLateral(検番)
Case "ﾗﾃﾗﾙﾎﾞ-追記公差あり":     Call LSelectLateral(検番)
End Select

Dim Temp As Variant:  Temp = InitialSetting(UF1)
End Sub

Private Sub KenHBox1_AfterUpdate()
UF1.KenHBox1 = Format((UF1.KenHBox1), "0.0")
Call LogOut(KenHBox1)
Call CalculateSheets(UF1, val(UF1.ATU))
End Sub
Private Sub KenHBox2_AfterUpdate()
UF1.KenHBox2 = Format((UF1.KenHBox2), "0.0")
Call LogOut(KenHBox2)
Call CalculateSheets(UF1, val(UF1.ATU))
'MsgBox Format(Round(CalculateSheets(UF1,UF1.Atu), 1), "0")
End Sub
Private Sub KenHBox3_AfterUpdate()
UF1.KenHBox3 = Format((UF1.KenHBox3), "0.0")
Call LogOut(KenHBox3)
Call CalculateSheets(UF1, val(UF1.ATU))
End Sub
Private Sub KenHBox4_AfterUpdate()
UF1.KenHBox4 = Format((UF1.KenHBox4), "0.0")
Call LogOut(KenHBox4)
Call CalculateSheets(UF1, val(UF1.ATU))
End Sub
Private Sub KenHBox5_AfterUpdate()
UF1.KenHBox5 = Format((UF1.KenHBox5), "0.0")
Call LogOut(KenHBox5)
Call CalculateSheets(UF1, val(UF1.ATU))
End Sub
Private Sub KenHBox6_AfterUpdate()
UF1.KenHBox6 = Format((UF1.KenHBox6), "0.0")
Call LogOut(KenHBox6)
Call CalculateSheets(UF1, val(UF1.ATU))
End Sub
Private Sub KenHBox7_AfterUpdate()
UF1.KenHBox7 = Format((UF1.KenHBox7), "0")
Call LogOut(KenHBox7)

'少枚数用の枚数を確定させる_2023.8.24
Dim th As Long
If UF1.ATU <> "" Then
  If UF1.KenHBox7 = "" Then
    th = Format(Round(CalculateSheets(UF1, UF1.ATU), 1), "0")
  End If
  If th >= 5 Then th = 5 '少枚数は5枚まで
End If

'If Kno = "H" And flag Then Call SmallNumberCount(th)
End Sub
Private Sub KenHBox7_Change()
If Not IsNumeric(Me.KenHBox7) Then Me.KenHBox7 = ""
End Sub
Private Sub KenCount_AfterUpdate()
UF1.KenCount = Format((UF1.KenCount), "0")
Call LogOut(KenCount)

'少枚数かつ枚数入力かつ検番頭H
Dim Kno As String: Kno = UF1.厚板.Caption
Dim flag As Boolean: If UF1.少枚数 Then flag = True

'少枚数用の枚数を確定させる_2023.8.24
Dim th As Long
If UF1.ATU <> "" Then
  If UF1.KenHBox7 = "" Then
    th = Format(Round(CalculateSheets(UF1, UF1.ATU), 1), "0")
  ElseIf UF1.KenCount <> "" Then
    th = UF1.KenCount
  End If
  If th >= 5 Then th = 5 '少枚数は5枚まで
End If

If Kno = "H" And flag Then Call SmallNumberCount(th)
End Sub
Private Sub KenCount_Change()
If Not IsNumeric(Me.KenCount) Then Me.KenCount = ""
End Sub


Private Sub LateralS_Change()
If IsNumeric(UF1.LateralS) = False Then
  If UF1.LateralS = "-" Then
  Else
    UF1.LateralS = ""
  End If
End If

If Left(Me.LateralS.Text, 1) = "-" Then Me.LateralS.maxLength = 4
If Left(Me.LateralS.Text, 1) <> "-" Then Me.LateralS.maxLength = 3
End Sub
Private Sub LateralU_Change()
If IsNumeric(UF1.LateralU) = False Then
  If UF1.LateralU = "-" Then
  Else
    UF1.LateralU = ""
  End If
End If

If Left(Me.LateralU.Text, 1) = "-" Then Me.LateralU.maxLength = 4
If Left(Me.LateralU.Text, 1) <> "-" Then Me.LateralU.maxLength = 3
End Sub
Private Sub LateralC_Change()
If IsNumeric(UF1.LateralC) = False Then
  If UF1.LateralC = "-" Then
  Else
    UF1.LateralC = ""
  End If
End If

If Left(Me.LateralC.Text, 1) = "-" Then Me.LateralC.maxLength = 4
If Left(Me.LateralC.Text, 1) <> "-" Then Me.LateralC.maxLength = 3
End Sub
Private Sub LateralD_Change()
If IsNumeric(UF1.LateralD) = False Then
  If UF1.LateralD = "-" Then
  Else
    UF1.LateralD = ""
  End If
End If

If Left(Me.LateralD.Text, 1) = "-" Then Me.LateralD.maxLength = 4
If Left(Me.LateralD.Text, 1) <> "-" Then Me.LateralD.maxLength = 3
End Sub
Private Sub LateralE_Change()
If IsNumeric(UF1.LateralE) = False Then
  If UF1.LateralE = "-" Then
  Else
    UF1.LateralE = ""
  End If
End If

If Left(Me.LateralE.Text, 1) = "-" Then Me.LateralE.maxLength = 4
If Left(Me.LateralE.Text, 1) <> "-" Then Me.LateralE.maxLength = 3
End Sub
Private Sub LateralF_Change()
If IsNumeric(UF1.LateralF) = False Then
  If UF1.LateralF = "-" Then
  Else
    UF1.LateralF = ""
  End If
End If

If Left(Me.LateralF.Text, 1) = "-" Then Me.LateralF.maxLength = 4
If Left(Me.LateralF.Text, 1) <> "-" Then Me.LateralF.maxLength = 3
End Sub

Private Sub LateralS_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.MPEX.Tag = "Click"
UF1.Manual.SetFocus
UF1.LateralS.SetFocus
'UF1.MPEX.Visible = False
'UF1.MP1.Visible = False
End Sub
Private Sub LateralU_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.MPEX.Tag = "Click"
UF1.Manual.SetFocus
UF1.LateralU.SetFocus
'UF1.MPEX.Visible = False
'UF1.MP1.Visible = False
End Sub
Private Sub LateralC_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.MPEX.Tag = "Click"
UF1.Manual.SetFocus
UF1.LateralC.SetFocus
'UF1.MPEX.Visible = False
'UF1.MP1.Visible = False
End Sub
Private Sub LateralD_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.MPEX.Tag = "Click"
UF1.Manual.SetFocus
UF1.LateralD.SetFocus
'UF1.MPEX.Visible = False
'UF1.MP1.Visible = False
End Sub
Private Sub LateralE_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.MPEX.Tag = "Click"
UF1.Manual.SetFocus
UF1.LateralE.SetFocus
'UF1.MPEX.Visible = False
'UF1.MP1.Visible = False
End Sub
Private Sub LateralF_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.MPEX.Tag = "Click"
UF1.Manual.SetFocus
UF1.LateralF.SetFocus
'UF1.MPEX.Visible = False
'UF1.MP1.Visible = False
End Sub

Private Sub DiagSKou_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.MPEX.Tag = "Click"
UF1.Manual.SetFocus
UF1.DiagSKou.SetFocus
End Sub
Private Sub DiagUKou_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.MPEX.Tag = "Click"
UF1.Manual.SetFocus
UF1.DiagUKou.SetFocus
End Sub
Private Sub DiagCKou_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.MPEX.Tag = "Click"
UF1.Manual.SetFocus
UF1.DiagCKou.SetFocus
End Sub
Private Sub DiagDKou_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.MPEX.Tag = "Click"
UF1.Manual.SetFocus
UF1.DiagDKou.SetFocus
End Sub
Private Sub DiagEKou_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.MPEX.Tag = "Click"
UF1.Manual.SetFocus
UF1.DiagEKou.SetFocus
End Sub
Private Sub DiagFKou_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.MPEX.Tag = "Click"
UF1.Manual.SetFocus
UF1.DiagFKou.SetFocus
End Sub

Private Sub DiagSKou_Change()
If IsNumeric(UF1.DiagSKou) = False Then
  If UF1.DiagSKou = "-" Then
  Else
    UF1.DiagSKou = ""
  End If
End If

If Left(Me.DiagSKou.Text, 1) = "-" Then Me.DiagSKou.maxLength = 4
If Left(Me.DiagSKou.Text, 1) <> "-" Then Me.DiagSKou.maxLength = 3
End Sub
Private Sub DiagUKou_Change()
If IsNumeric(UF1.DiagUKou) = False Then
  If UF1.DiagUKou = "-" Then
  Else
    UF1.DiagUKou = ""
  End If
End If

If Left(Me.DiagUKou.Text, 1) = "-" Then Me.DiagUKou.maxLength = 4
If Left(Me.DiagUKou.Text, 1) <> "-" Then Me.DiagUKou.maxLength = 3
End Sub
Private Sub DiagCKou_Change()
If IsNumeric(UF1.DiagCKou) = False Then
  If UF1.DiagCKou = "-" Then
  Else
    UF1.DiagCKou = ""
  End If
End If

If Left(Me.DiagCKou.Text, 1) = "-" Then Me.DiagCKou.maxLength = 4
If Left(Me.DiagCKou.Text, 1) <> "-" Then Me.DiagCKou.maxLength = 3
End Sub
Private Sub DiagDKou_Change()
If IsNumeric(UF1.DiagDKou) = False Then
  If UF1.DiagDKou = "-" Then
  Else
    UF1.DiagDKou = ""
  End If
End If

If Left(Me.DiagDKou.Text, 1) = "-" Then Me.DiagDKou.maxLength = 4
If Left(Me.DiagDKou.Text, 1) <> "-" Then Me.DiagDKou.maxLength = 3
End Sub
Private Sub DiagEKou_Change()
If IsNumeric(UF1.DiagEKou) = False Then
  If UF1.DiagEKou = "-" Then
  Else
    UF1.DiagEKou = ""
  End If
End If

If Left(Me.DiagEKou.Text, 1) = "-" Then Me.DiagEKou.maxLength = 4
If Left(Me.DiagEKou.Text, 1) <> "-" Then Me.DiagEKou.maxLength = 3
End Sub
Private Sub DiagFKou_Change()
If IsNumeric(UF1.DiagFKou) = False Then
  If UF1.DiagFKou = "-" Then
  Else
    UF1.DiagFKou = ""
  End If
End If

If Left(Me.DiagFKou.Text, 1) = "-" Then Me.DiagFKou.maxLength = 4
If Left(Me.DiagFKou.Text, 1) <> "-" Then Me.DiagFKou.maxLength = 3
End Sub

Private Sub LateralS_Enter()
Call TagSet(Me.LateralS)
ChangeBackColor LateralS
End Sub
Private Sub LateralU_Enter()
Call TagSet(Me.LateralU)
ChangeBackColor LateralU
End Sub
Private Sub LateralC_Enter()
Call TagSet(Me.LateralC)
ChangeBackColor LateralC
End Sub
Private Sub LateralD_Enter()
Call TagSet(Me.LateralD)
ChangeBackColor LateralD
End Sub
Private Sub LateralE_Enter()
Call TagSet(Me.LateralE)
ChangeBackColor LateralE
End Sub
Private Sub LateralF_Enter()
Call TagSet(Me.LateralF)
ChangeBackColor LateralF
End Sub
Private Sub LateralS_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub LateralU_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub LateralC_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub LateralD_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub LateralE_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub LateralF_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub


Private Sub DiagSKou_Enter()
Call TagSet(Me.DiagSKou)
ChangeBackColor DiagSKou
End Sub
Private Sub DiagUKou_Enter()
Call TagSet(Me.DiagUKou)
ChangeBackColor DiagUKou
End Sub
Private Sub DiagCKou_Enter()
Call TagSet(Me.DiagCKou)
ChangeBackColor DiagCKou
End Sub
Private Sub DiagDKou_Enter()
Call TagSet(Me.DiagDKou)
ChangeBackColor DiagDKou
End Sub
Private Sub DiagEKou_Enter()
Call TagSet(Me.DiagEKou)
ChangeBackColor DiagEKou
End Sub
Private Sub DiagFKou_Enter()
Call TagSet(Me.DiagFKou)
ChangeBackColor DiagFKou
End Sub

Private Sub DiagSKou_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub DiagUKou_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub DiagCKou_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub DiagDKou_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub DiagEKou_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub DiagFKou_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub

Private Sub DiagSKou_AfterUpdate()
Call CheckLabelsByCategory("S") '2025.1.23
End Sub
Private Sub DiagUKou_AfterUpdate()
Call CheckLabelsByCategory("U") '2025.1.23
End Sub
Private Sub DiagCKou_AfterUpdate()
Call CheckLabelsByCategory("C") '2025.1.23
End Sub
Private Sub DiagDKou_AfterUpdate()
Call CheckLabelsByCategory("D") '2025.1.23
End Sub
Private Sub DiagEKou_AfterUpdate()
Call CheckLabelsByCategory("E") '2025.1.23
End Sub
Private Sub DiagFKou_AfterUpdate()
Call CheckLabelsByCategory("F") '2025.1.23
End Sub

Private Sub DiagSKou_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
UF1.Loading.Visible = True

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Call WriteLog("*UFKen:" & "DiagSKou_BeforeUpdate:" & Me.DiagSKou.Text & ":" & Lname, "", "")

If Not IsNumeric(UF1.DiagSKou) Then
'  UF1.LateralS = ""
ElseIf UF1.DiagSKou = "-" Then '残す　測定不要のパターン
Else

  UF1.DiagSKou = UnitRecord(val(UF1.DiagSKou)) '測定値を0.5ピッチに変換
  UF1.DiagSKou = Format((UF1.DiagSKou), "0.0")
  
  UF1.Loading.Visible = False
  UF1.Repaint
  '判定
  Call RightAngleJudgColor(UF1.DiagSKou)
  Call Posting
End If

If Me.DiagSKou = "" Then Me.DiagSKou.BackColor = RGB(255, 255, 255)

UF1.Loading.Visible = False
UF1.Repaint

Call Out_Tolerance(UF1.INP.Tag)
End Sub
Private Sub DiagUKou_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
UF1.Loading.Visible = True

Call TagSet(Me.DiagUKou)

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Call WriteLog("*UFKen:" & "DiagUKou_BeforeUpdate:" & Me.DiagUKou.Text & ":" & Lname, "", "")

If Not IsNumeric(UF1.DiagUKou) Then
'  UF1.LateralS = ""
ElseIf UF1.DiagUKou = "-" Then '残す　測定不要のパターン
Else

  UF1.DiagUKou = UnitRecord(val(UF1.DiagUKou)) '測定値を0.5ピッチに変換
  UF1.DiagUKou = Format((UF1.DiagUKou), "0.0")
  
  UF1.Loading.Visible = False
  UF1.Repaint
  '判定
  Call RightAngleJudgColor(UF1.DiagUKou)
  Call Posting
End If

If Me.DiagUKou = "" Then Me.DiagUKou.BackColor = RGB(255, 255, 255)

UF1.Loading.Visible = False
UF1.Repaint

Call Out_Tolerance(UF1.INP.Tag)
End Sub
Private Sub DiagCKou_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
UF1.Loading.Visible = True

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Call WriteLog("*UFKen:" & "DiagCKou_BeforeUpdate:" & Me.DiagCKou.Text & ":" & Lname, "", "")

If Not IsNumeric(UF1.DiagCKou) Then
'  UF1.LateralS = ""
ElseIf UF1.DiagSKou = "-" Then '残す　測定不要のパターン
Else

  UF1.DiagCKou = UnitRecord(val(UF1.DiagCKou)) '測定値を0.5ピッチに変換
  UF1.DiagCKou = Format((UF1.DiagCKou), "0.0")
  
  UF1.Loading.Visible = False
  UF1.Repaint
  '判定
  Call RightAngleJudgColor(UF1.DiagCKou)
  Call Posting
End If

If Me.DiagCKou = "" Then Me.DiagCKou.BackColor = RGB(255, 255, 255)

UF1.Loading.Visible = False
UF1.Repaint

Call Out_Tolerance(UF1.INP.Tag)
End Sub
Private Sub DiagDKou_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
UF1.Loading.Visible = True

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Call WriteLog("*UFKen:" & "DiagDKou_BeforeUpdate:" & Me.DiagDKou.Text & ":" & Lname, "", "")

If Not IsNumeric(UF1.DiagDKou) Then
'  UF1.LateralS = ""
ElseIf UF1.DiagDKou = "-" Then '残す　測定不要のパターン
Else

  UF1.DiagSKou = UnitRecord(val(UF1.DiagDKou)) '測定値を0.5ピッチに変換
  UF1.DiagSKou = Format((UF1.DiagDKou), "0.0")
  
  UF1.Loading.Visible = False
  UF1.Repaint
  '判定
  Call RightAngleJudgColor(UF1.DiagDKou)
  Call Posting
End If

If Me.DiagDKou = "" Then Me.DiagDKou.BackColor = RGB(255, 255, 255)

UF1.Loading.Visible = False
UF1.Repaint

Call Out_Tolerance(UF1.INP.Tag)
End Sub
Private Sub DiagEKou_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
UF1.Loading.Visible = True

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Call WriteLog("*UFKen:" & "DiagEKou_BeforeUpdate:" & Me.DiagEKou.Text & ":" & Lname, "", "")

If Not IsNumeric(UF1.DiagEKou) Then
'  UF1.LateralS = ""
ElseIf UF1.DiagEKou = "-" Then '残す　測定不要のパターン
Else

  UF1.DiagEKou = UnitRecord(val(UF1.DiagEKou)) '測定値を0.5ピッチに変換
  UF1.DiagEKou = Format((UF1.DiagEKou), "0.0")
  
  UF1.Loading.Visible = False
  UF1.Repaint
  '判定
  Call RightAngleJudgColor(UF1.DiagEKou)
  Call Posting
End If

If Me.DiagEKou = "" Then Me.DiagEKou.BackColor = RGB(255, 255, 255)

UF1.Loading.Visible = False
UF1.Repaint

Call Out_Tolerance(UF1.INP.Tag)
End Sub
Private Sub DiagFKou_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
UF1.Loading.Visible = True

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Call WriteLog("*UFKen:" & "DiagFKou_BeforeUpdate:" & Me.DiagFKou.Text & ":" & Lname, "", "")

If Not IsNumeric(UF1.DiagFKou) Then
'  UF1.LateralS = ""
ElseIf UF1.DiagFKou = "-" Then '残す　測定不要のパターン
Else

  UF1.DiagFKou = UnitRecord(val(UF1.DiagFKou)) '測定値を0.5ピッチに変換
  UF1.DiagFKou = Format((UF1.DiagFKou), "0.0")
  
  UF1.Loading.Visible = False
  UF1.Repaint
  '判定
  Call RightAngleJudgColor(UF1.DiagFKou)
  Call Posting
End If

If Me.DiagEKou = "" Then Me.DiagEKou.BackColor = RGB(255, 255, 255)

UF1.Loading.Visible = False
UF1.Repaint

Call Out_Tolerance(UF1.INP.Tag)
End Sub
Private Sub LateralS_AfterUpdate()
Call CheckLabelsByCategory("S") '2025.1.23
End Sub
Private Sub LateralU_AfterUpdate()
Call CheckLabelsByCategory("U") '2025.1.23
End Sub
Private Sub LateralC_AfterUpdate()
Call CheckLabelsByCategory("C") '2025.1.23
End Sub
Private Sub LateralD_AfterUpdate()
Call CheckLabelsByCategory("D") '2025.1.23
End Sub
Private Sub LateralE_AfterUpdate()
Call CheckLabelsByCategory("E") '2025.1.23
End Sub
Private Sub LateralF_AfterUpdate()
Call CheckLabelsByCategory("F") '2025.1.23
End Sub

Private Sub LateralS_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
UF1.Loading.Visible = True

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Call WriteLog("*UFKen:" & "LateralS_BeforeUpdate:" & Me.LateralS.Text & ":" & Lname, "", "")

If Not IsNumeric(UF1.LateralS) Then
'  UF1.LateralS = ""
ElseIf UF1.LateralS = "-" Then '残す　測定不要のパターン
Else
'  If IsNumeric(UF1.LateralS) Then UF1.LateralS = Application.WorksheetFunction.RoundDown((UF1.LateralS), 1) '下1桁で切り捨て

  UF1.LateralS = UnitRecord(val(UF1.LateralS)) 'ラテラルボー測定値を0.5ピッチに変換
  UF1.LateralS = Format((UF1.LateralS), "0.0")
  
  UF1.Loading.Visible = False
  UF1.Repaint
  
  Call LateralbowJudgColor(UF1.LateralS)
  Call Posting
End If

If Me.LateralS = "" Then Me.LateralS.BackColor = RGB(255, 255, 255)

UF1.Loading.Visible = False
UF1.Repaint

Call Out_Tolerance(UF1.INP.Tag)
End Sub

Private Sub LateralU_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
UF1.Loading.Visible = True
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Call WriteLog("*UFKen:" & "LateralU_BeforeUpdate:" & Me.LateralU.Text & ":" & Lname, "", "")

If Not IsNumeric(UF1.LateralU) Then
'  UF1.LateralS = ""
ElseIf UF1.LateralU = "-" Then '残す　測定不要のパターン
Else
'  If IsNumeric(UF1.LateralU) Then UF1.LateralU = Application.WorksheetFunction.RoundDown((UF1.LateralU), 1) '下1桁で切り捨て

  UF1.LateralU = UnitRecord(val(UF1.LateralU)) 'ラテラルボー測定値を0.5ピッチに変換
  UF1.LateralU = Format((UF1.LateralU), "0.0")
  
  UF1.Loading.Visible = False
  UF1.Repaint
  
  Call LateralbowJudgColor(UF1.LateralU)
  Call Posting
End If

If Me.LateralU = "" Then Me.LateralU.BackColor = RGB(255, 255, 255)

UF1.Loading.Visible = False
UF1.Repaint

Call Out_Tolerance(UF1.INP.Tag)
End Sub
Private Sub LateralC_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
UF1.Loading.Visible = True
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Call WriteLog("*UFKen:" & "LateralC_BeforeUpdate:" & Me.LateralC.Text & ":" & Lname, "", "")

If Not IsNumeric(UF1.LateralC) Then
'  UF1.LateralS = ""
ElseIf UF1.LateralC = "-" Then '残す　測定不要のパターン
Else
'  If IsNumeric(UF1.LateralC) Then UF1.LateralC = Application.WorksheetFunction.RoundDown((UF1.LateralC), 1) '下1桁で切り捨て

  UF1.LateralC = UnitRecord(val(UF1.LateralC)) 'ラテラルボー測定値を0.5ピッチに変換
  UF1.LateralC = Format((UF1.LateralC), "0.0")
  
  UF1.Loading.Visible = False
  UF1.Repaint
  
  Call LateralbowJudgColor(UF1.LateralC)
  Call Posting
End If

If Me.LateralC = "" Then Me.LateralC.BackColor = RGB(255, 255, 255)

UF1.Loading.Visible = False
UF1.Repaint

Call Out_Tolerance(UF1.INP.Tag)
End Sub
Private Sub LateralD_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
UF1.Loading.Visible = True
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Call WriteLog("*UFKen:" & "LateralD_BeforeUpdate:" & Me.LateralD.Text & ":" & Lname, "", "")

If Not IsNumeric(UF1.LateralD) Then
'  UF1.LateralS = ""
ElseIf UF1.LateralD = "-" Then '残す　測定不要のパターン
Else
'  If IsNumeric(UF1.LateralD) Then UF1.LateralD = Application.WorksheetFunction.RoundDown((UF1.LateralD), 1) '下1桁で切り捨て

  UF1.LateralD = UnitRecord(val(UF1.LateralD)) 'ラテラルボー測定値を0.5ピッチに変換
  UF1.LateralD = Format((UF1.LateralD), "0.0")
  
  UF1.Loading.Visible = False
  UF1.Repaint
  
  Call LateralbowJudgColor(UF1.LateralD)
  Call Posting
End If

If Me.LateralD = "" Then Me.LateralD.BackColor = RGB(255, 255, 255)

UF1.Loading.Visible = False
UF1.Repaint

Call Out_Tolerance(UF1.INP.Tag)
End Sub
Private Sub LateralE_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
UF1.Loading.Visible = True
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Call WriteLog("*UFKen:" & "LateralE_BeforeUpdate:" & Me.LateralE.Text & ":" & Lname, "", "")

If Not IsNumeric(UF1.LateralE) Then
'  UF1.LateralS = ""
ElseIf UF1.LateralE = "-" Then '残す　測定不要のパターン
Else
'  If IsNumeric(UF1.LateralE) Then UF1.LateralE = Application.WorksheetFunction.RoundDown((UF1.LateralE), 1) '下1桁で切り捨て

  UF1.LateralE = UnitRecord(val(UF1.LateralE)) 'ラテラルボー測定値を0.5ピッチに変換
  UF1.LateralE = Format((UF1.LateralE), "0.0")
  
  UF1.Loading.Visible = False
  UF1.Repaint
  
  Call LateralbowJudgColor(UF1.LateralE)
  Call Posting
End If

If Me.LateralE = "" Then Me.LateralE.BackColor = RGB(255, 255, 255)

UF1.Loading.Visible = False
UF1.Repaint

Call Out_Tolerance(UF1.INP.Tag)
End Sub
Private Sub LateralF_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
UF1.Loading.Visible = True
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Call WriteLog("*UFKen:" & "LateralF_BeforeUpdate:" & Me.LateralF.Text & ":" & Lname, "", "")

If Not IsNumeric(UF1.LateralF) Then
'  UF1.LateralS = ""
ElseIf UF1.LateralF = "-" Then '残す　測定不要のパターン
Else
'  If IsNumeric(UF1.LateralE) Then UF1.LateralE = Application.WorksheetFunction.RoundDown((UF1.LateralE), 1) '下1桁で切り捨て

  UF1.LateralF = UnitRecord(val(UF1.LateralF)) 'ラテラルボー測定値を0.5ピッチに変換
  UF1.LateralF = Format((UF1.LateralF), "0.0")
  
  UF1.Loading.Visible = False
  UF1.Repaint
  
  Call LateralbowJudgColor(UF1.LateralF)
  Call Posting
End If

If Me.LateralF = "" Then Me.LateralF.BackColor = RGB(255, 255, 255)

UF1.Loading.Visible = False
UF1.Repaint

Call Out_Tolerance(UF1.INP.Tag)
End Sub

Private Sub LC1_Click()
Call LcColor(LC1, 1)
Call CheckLabelsByCategory("S") '2025.1.23
End Sub
Private Sub LC2_Click()
Call LcColor(LC2, 1)
Call CheckLabelsByCategory("U")
End Sub
Private Sub LC3_Click()
Call LcColor(LC3, 1)
Call CheckLabelsByCategory("C")
End Sub
Private Sub LC4_Click()
Call LcColor(LC4, 1)
Call CheckLabelsByCategory("D")
End Sub
Private Sub LC5_Click()
Call LcColor(LC5, 1)
Call CheckLabelsByCategory("E")
End Sub
Private Sub LC6_Click()
Call LcColor(LC6, 1)
Call CheckLabelsByCategory("F")
End Sub
Private Sub LC7_Click()
Call LcColor(LC7, 7)
Call CheckLabelsByCategory("S")
End Sub
Private Sub LC8_Click()
Call LcColor(LC8, 7)
Call CheckLabelsByCategory("U")
End Sub
Private Sub LC9_Click()
Call LcColor(LC9, 7)
Call CheckLabelsByCategory("C")
End Sub
Private Sub LC10_Click()
Call LcColor(LC10, 7)
Call CheckLabelsByCategory("D")
End Sub
Private Sub LC11_Click()
Call LcColor(LC11, 7)
Call CheckLabelsByCategory("E")
End Sub
Private Sub LC12_Click()
Call LcColor(LC12, 7)
Call CheckLabelsByCategory("F")
End Sub

Private Sub LC1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UF1.LC1.BackColor = RGB(255, 255, 255) ' 白色
Call btn_allClea(UF1.LBVisual)
End Sub
Private Sub LC2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UF1.LC2.BackColor = RGB(255, 255, 255)
Call btn_allClea(UF1.LBVisual)
End Sub
Private Sub LC3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UF1.LC3.BackColor = RGB(255, 255, 255)
Call btn_allClea(UF1.LBVisual)
End Sub
Private Sub LC4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UF1.LC4.BackColor = RGB(255, 255, 255)
Call btn_allClea(UF1.LBVisual)
End Sub
Private Sub LC5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UF1.LC5.BackColor = RGB(255, 255, 255)
Call btn_allClea(UF1.LBVisual)
End Sub
Private Sub LC6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UF1.LC6.BackColor = RGB(255, 255, 255)
Call btn_allClea(UF1.LBVisual)
End Sub
Private Sub LC7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UF1.LC7.BackColor = RGB(255, 255, 255)
Call btn_allClea(UF1.LBVisual)
End Sub
Private Sub LC8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UF1.LC8.BackColor = RGB(255, 255, 255)
Call btn_allClea(UF1.LBVisual)
End Sub
Private Sub LC9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UF1.LC9.BackColor = RGB(255, 255, 255)
Call btn_allClea(UF1.LBVisual)
End Sub
Private Sub LC10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UF1.LC10.BackColor = RGB(255, 255, 255)
Call btn_allClea(UF1.LBVisual)
End Sub

Private Sub ListReflect_Click()
Dim Temps As String: Temps = GetSelectedItems(UF1.DefectFactorLB)

If UF1.DefectFactorLB.ListIndex = -1 Then
  MsgBox "不良要因が選択されていません", vbCritical
  Exit Sub
End If

Dim FL As Boolean: If UF1("SBad") = "" And UF1("UBad") = "" And UF1("CBad") = "" And UF1("DBad") = "" And UF1("EBad") = "" And UF1("FBad") = "" Then FL = True
If FL Then
  MsgBox "枚数入力がありません", vbCritical
  Exit Sub
End If

Dim S As Long: If IsNumeric(UF1("SBad")) Then S = val(UF1("SBad"))
Dim U As Long: If IsNumeric(UF1("UBad")) Then U = val(UF1("UBad"))
Dim C As Long: If IsNumeric(UF1("CBad")) Then C = val(UF1("CBad"))
Dim D As Long: If IsNumeric(UF1("DBad")) Then D = val(UF1("DBad"))
Dim E As Long: If IsNumeric(UF1("EBad")) Then E = val(UF1("EBad"))
Dim F As Long: If IsNumeric(UF1("FBad")) Then F = val(UF1("FBad"))

'スポット処理は囲い文字処理なし_2024.1.10
If Temps = "ｽﾎﾟｯﾄ処理" Then GoTo SkipPrint
If S > 50 Or U > 50 Or C > 50 Or D > 50 Or E > 50 Or F > 50 Then
  MsgBox "すみません、50枚以上は分けて入力してください" '囲い文字処理の関係で50まで
  Exit Sub
End If

SkipPrint:
Dim n As String: n = TecReturn
Call SelectionInput(UF1(n)) '出力

'リセット
Dim i As Long
For i = 1 To 6
  Select Case i
  Case 1
    UF1("SBad") = ""
    UF1("SBad").BackColor = RGB(255, 255, 255)
  Case 2
    UF1("UBad") = ""
    UF1("UBad").BackColor = RGB(255, 255, 255)
  Case 3
    UF1("CBad") = ""
    UF1("CBad").BackColor = RGB(255, 255, 255)
  Case 4
    UF1("DBad") = ""
    UF1("DBad").BackColor = RGB(255, 255, 255)
  Case 5
    UF1("EBad") = ""
    UF1("EBad").BackColor = RGB(255, 255, 255)
  Case 6
    UF1("FBad") = ""
    UF1("FBad").BackColor = RGB(255, 255, 255)
  End Select
Next i

End Sub

Private Sub LoadLot_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'入力制限
Select Case KeyAscii.Value
Case 48 To 57 '0~9
Case 65 To 90 'A~Z
Case 97 To 122 'a~z
KeyAscii.Value = KeyAscii.Value - 32 '大文字強制変換
Case Else
KeyAscii.Value = 0
End Select
End Sub
Private Sub LoadWorker_Change()
If UF1.FlagLabel <> 1 Then Exit Sub
If Len(Me.LoadWorker.Text) <> 13 Then  'ｶｰﾄﾞ16桁　-4000番
  Exit Sub
Else
  UF1.FlagLabel = 0
  If NameConv(UF1.LoadWorker.Value) = "" Then
    UF1.LoadWorker = ""
    UF1.FlagLabel = 1
    Exit Sub
  End If
   UF1.LoadWorker = NameConv(UF1.LoadWorker.Value)
End If
UF1.FlagLabel = 1
End Sub

Private Sub MP1_Click(ByVal Index As Long)
UF1.MPEX.Tag = "Click"
End Sub

Private Sub MP1_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UF1.MPEX.Tag = "Click"
End Sub

Private Sub MPEX_Change()
UF1.Load.Caption = "on" 'イベント再開　DropButtonClickイベントでoffになったままになる
If UF1.ATU = "" Then Exit Sub
If UF1.MPEX.Value = 0 Then Call DrawListBoxContent(UF1, UF1.ATU)

If UF1.MPEX.Value = 2 Then '歪ページ
  If Me.AorB.Caption = "Now_A" Then UF1.MPEX.BackColor = &H8000000F 'デフォ色に戻す
  If Me.AorB.Caption = "Now_B" Then Me.MPEX.BackColor = RGB(0, 0, 255) ' 青色
  
  Dim TempW As Double: TempW = Weight(val(UF1.ATU), val(UF1.Hab), val(UF1.Tak))
  If TempW = 0 Then Exit Sub '重量なし処理停止
  
  Dim 歪AB As String:    歪AB = WithorWithout_Gauge '2023.11.20

  If 歪AB <> "" Then
    If 35 >= TempW Then MsgBox "ゲージ不要 * 両面定盤歪測定必要"
    If 35 < TempW Then MsgBox "ゲージによる測定必要"
  End If
  
  If Me.AorB.Caption = "Now_B" And TempW > 35 Then MsgBox "B面歪 入力画面です"
  
  Call 音楽を流す("歪の公差.wav")
  
  UF1.Tag = "STOP" '音楽ｽﾄｯﾌﾟ
  UF1.UFTag = UF1.Tag '目視用
  
  Call HizLoopJudg(0)
  
  UF1.Tag = ""
  UF1.UFTag = UF1.Tag
Else
  UF1.MPEX.BackColor = &H8000000F 'デフォ色に戻す
End If
End Sub

Private Sub MPEX_Click(ByVal Index As Long)
On Error Resume Next

UF1.MPEX.Tag = "Click" '2023.8.18_LateralbowJudgColorで致命的なエラーになる対策**Tagに文字列のある時は処理を行わない

Call btn_allClea(UF1.LBVisual)
Dim i As Long
For i = 1 To 10
  UF1("LC" & i).BackColor = RGB(255, 255, 255) ' 白色' クリックされたラベル以外は白色にする
Next i

Call Fm_Clear("Click") '歪ﾌﾚｰﾑﾘｾｯﾄ_2021.03.29

If UF1.ATU <> "" Then Call CalculateSheets(UF1, val(UF1.ATU))

Dim Kno As String: Kno = UF1.厚板.Caption
Dim flag As Boolean: If UF1.少枚数 Then flag = True

'少枚数用の枚数を確定させる_2023.8.24
'2024.10.30_検査回数と検入枚数をわける　KenHBox7→KenCount
Dim th As Long
If UF1.ATU <> "" Then
  If UF1.KenHBox7 = "" Then
    th = Format(Round(CalculateSheets(UF1, UF1.ATU), 1), "0")
  ElseIf UF1.KenCount <> "" Then
    th = UF1.KenCount
  End If
  If th >= 5 Then th = 5 '少枚数は5枚まで
End If

If Kno = "H" And flag Then Call SmallNumberCount(th)

UF1.Manual.SetFocus
UF1.StartTB.SetFocus
End Sub

Private Sub MPEX_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UF1.MPEX.Tag = "Click"
End Sub

Private Sub NoL_Click()
Call OrdNoOpen(Me.Lot) 'オーダーをコンボに入れる
End Sub

Private Sub OrdNoCB_Change()
'---------------------Lot情報履歴検索
If WsCheck("Lot情報履歴") = False Then Exit Sub  '-------------------------ｼｰﾄなし処理停止
If UF1.Lot = "" Then Exit Sub
Dim HisSH As Worksheet:   Set HisSH = ThisWorkbook.Sheets("Lot情報履歴")
Dim HisRan As Range:      Set HisRan = HisSH.Range(HisSH.Cells(1, 1), HisSH.Cells(UR(HisSH, 1), UC(HisSH, 1)))
Dim Hist(): Hist = HisRan
Dim i As Long, a As Long, n As Long: n = 0
Dim R_Ord As Long, R_Packing As Long, R_Lot As Long
For i = 1 To UBound(Hist, 2) 'ﾌｨｰﾙﾄﾞ名
  Select Case Hist(1, i)
  Case "ﾛｯﾄNo":        R_Lot = i
  Case "ｵｰﾀﾞｰNo":      R_Ord = i
  Case "包装仕様NO":   R_Packing = i
  End Select
Next i
Dim FL As Boolean: FL = False
Dim TempHiki()
For i = LBound(Hist, 1) To UBound(Hist, 1) '1次元ﾁｪｯｸ
  If Hist(i, R_Lot) = UF1.Lot Then a = a + 1
Next i

If a = 0 Then Exit Sub

ReDim TempHiki(1 To a, 1 To 2)
For i = LBound(Hist, 1) To UBound(Hist, 1)
  If Hist(i, R_Lot) = UF1.Lot Then '一致の時
    n = n + 1
    TempHiki(n, 1) = Hist(i, R_Ord)
    TempHiki(n, 2) = Hist(i, R_Packing)
  End If
Next i
For i = LBound(TempHiki, 1) To UBound(TempHiki, 1)
  If UF1.OrdNoCB.Text = TempHiki(i, 1) Then '一致の時
    UF1.PackingNo.Caption = TempHiki(i, 2)
  End If
Next i
If UF1.PackingNo = "" Then Exit Sub
'仕様書ｺﾒﾝﾄ表示
Dim TableName As String: TableName = "PkgSpecData"
Dim FPath As String
Select Case True
Case UF1.Test:  FPath = PackingPath
Case Else:      FPath = PackingPath
End Select
'ﾌｧｲﾙ名
Dim target As String:   target = "伝送用ファイル.accdb"
'ﾌｧｲﾙﾊﾟｽ
Dim FilePath As String: FilePath = FPath & target
  If ReadCheck(FilePath, 1) = False Then Exit Sub '1：ファイル,2：フォルダ
'検索項目
Dim SQLfilter As String: SQLfilter = "[包装仕様書№]='" & UF1.PackingNo & "'" '仕様書No検索
Dim Temp As Variant: Temp = GetRecordsArr(FilePath, TableName, SQLfilter) 'All,fields,record
'ﾌｨｰﾙﾄﾞ名ﾁｪｯｸ　*更新に対応
Dim fieldsfilter As String: fieldsfilter = ""
Dim Fields As Variant:  Fields = GetFieldsArr(FilePath, TableName, fieldsfilter)
If IsEmpty(Fields) Then Exit Sub 'emptyの場合停止
Dim C_梱包形態 As Long, C_タイトル As Long
For i = LBound(Fields, 2) To UBound(Fields, 2)
  Select Case Fields(1, i)
  Case "梱包形態":    C_梱包形態 = i
  Case "タイトル":    C_タイトル = i
  End Select
Next i

If IsEmpty(Temp) = False Then UF1.PackingForm.Caption = Temp(1, C_梱包形態)
If IsEmpty(Temp) = False Then UF1.PackingComment.Caption = Temp(1, C_タイトル)
Call VC_Picture(UF1.OrdNoCB.Text) 'VCリロード
Call MaterialsMsgBox(UF1.OrdNoCB.Text) '注意喚起リロード
End Sub
Private Sub SBad_Change()
If IsNumeric(UF1.SBad) = False Then UF1.SBad = ""
End Sub
Private Sub UBad_Change()
If IsNumeric(UF1.UBad) = False Then UF1.UBad = ""
End Sub
Private Sub CBad_Change()
If IsNumeric(UF1.CBad) = False Then UF1.CBad = ""
End Sub
Private Sub DBad_Change()
If IsNumeric(UF1.DBad) = False Then UF1.DBad = ""
End Sub
Private Sub EBad_Change()
If IsNumeric(UF1.EBad) = False Then UF1.EBad = ""
End Sub
Private Sub FBad_Change()
If IsNumeric(UF1.FBad) = False Then UF1.FBad = ""
End Sub

Private Sub SBad_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UF1.SBad = Format((UF1.SBad), "0")
Call TeColor(SBad)
End Sub
Private Sub UBad_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UF1.UBad = Format((UF1.UBad), "0")
Call TeColor(UBad)
End Sub
Private Sub CBad_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UF1.CBad = Format((UF1.CBad), "0")
Call TeColor(CBad)
End Sub
Private Sub DBad_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UF1.DBad = Format((UF1.DBad), "0")
Call TeColor(DBad)
End Sub
Private Sub EBad_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UF1.EBad = Format((UF1.EBad), "0")
Call TeColor(EBad)
End Sub
Private Sub FBad_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UF1.EBad = Format((UF1.FBad), "0")
Call TeColor(FBad)
End Sub

Private Sub SBH_Change()
If IsNumeric(UF1.SBH) = False Then UF1.SBH = ""
End Sub
Private Sub UBH_Change()
If IsNumeric(UF1.UBH) = False Then UF1.UBH = ""
End Sub
Private Sub CBH_Change()
If IsNumeric(UF1.CBH) = False Then UF1.CBH = ""
End Sub
Private Sub DBH_Change()
If IsNumeric(UF1.DBH) = False Then UF1.DBH = ""
End Sub
Private Sub EBH_Change()
If IsNumeric(UF1.EBH) = False Then UF1.EBH = ""
End Sub

Private Sub SBH_AfterUpdate()
UF1.SBH = Format((UF1.SBH), "0")
Call LogOut(SBH)
Call simplyjudge(SBH, BhJudg(SBH, Me.Zai, Me.Tyu))
End Sub
Private Sub UBH_AfterUpdate()
UF1.UBH = Format((UF1.UBH), "0")
Call LogOut(UBH)
Call simplyjudge(UBH, BhJudg(UBH, Me.Zai, Me.Tyu))
End Sub
Private Sub CBH_AfterUpdate()
UF1.CBH = Format((UF1.CBH), "0")
Call LogOut(CBH)
Call simplyjudge(CBH, BhJudg(CBH, Me.Zai, Me.Tyu))
End Sub
Private Sub DBH_AfterUpdate()
UF1.DBH = Format((UF1.DBH), "0")
Call LogOut(DBH)
Call simplyjudge(DBH, BhJudg(DBH, Me.Zai, Me.Tyu))
End Sub
Private Sub EBH_AfterUpdate()
UF1.EBH = Format((UF1.EBH), "0")
Call LogOut(EBH)
Call simplyjudge(EBH, BhJudg(EBH, Me.Zai, Me.Tyu))
End Sub
Private Sub TestButton_Click()

'Call 音楽を流す("確認作業.wav")

'Call MainDataSave("保存用") '耳
'
'ThisWorkbook.Sheets("保存用").Visible = True '保存用　印刷用_Test’Accesslog
'Call 不良要因印刷範囲("Test")
'ThisWorkbook.Sheets("異常ロット用_Test").Visible = True
'
Call 印刷転送
ThisWorkbook.Sheets("印刷用_Test").Visible = True '保存用　印刷用_Test 不良要因印刷_Test
Application.Visible = True 'Excel表示
'Select Case True
'Case UF1.L1Work: LN = "L-1"
'Case UF1.LVC:    LN = "LVC"
'Case UF1.HVC:    LN = "HVC"
'Case UF1.Bala:   LN = "大板"
'Case UF1.Tott:   LN = "大板"
'Case UF1.Maru:   LN = "中板"
'Case Else:       LN = "Test"
'End Select
'
'ThisWorkbook.Sheets("印刷履歴管理").Visible = True

'Call 履歴保存

'Dim SaveSh As Worksheet: Set SaveSh = ThisWorkbook.Sheets("印刷用_Test")
'Dim f As String: f = ThisWorkbook.path
'Call PDFCREATE(f, SaveSh, "L-1") 'ﾊﾟｽ,保存ｼｰﾄ,保存ﾌｧｲﾙ名

'Dim a As String: a = Visual_Check
'MsgBox a
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'
'Dim 仮, i As Long, n As Long
'Dim データ()
'仮 = UF_Array(1, 20, 1, 5)
'Dim Rcount As Long: Rcount = UBound(仮, 1)
'Dim Ccount As Long: Ccount = UBound(検査)
'
'ReDim データ(1 To Rcount + 1, 1 To Ccount + 1)
'For i = 1 To Ccount + 1
'  データ(1, i) = 検査(i - 1)
'Next i
''Stop
'For i = 2 To Rcount + 1
'  For n = 1 To UBound(仮, 2)
'    データ(i, n) = 仮(i - 1, n)
'  Next n
'Next i
''Stop
'
'Application.Visible = True 'Excel表示
End Sub
Private Sub ViewCB_Click()
'Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
'
''入力なし、シートなしの場合からの確認時対策
'If (WsCheck("不良要因印刷" & "_" & Lname) = False) Then
'  Call AddWS("不良要因印刷" & "_" & Lname)
'  Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
'ElseIf ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Cells(1, 2) <> UF1.LOT Then
'  Call AddWS("不良要因印刷" & "_" & Lname)
'  Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
'End If
'Dim WS As Worksheet: Set WS = ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname)
''UF1.Lot.TextBoxの裏にImage1を隠してある
'Dim i As Long
'WS.Range(WS.Cells(1, 2), WS.Cells(11, 24)).Copy
'
'With UF1
'  .Image1.Visible = True
'  .Image1.AutoSize = True
'  .Image1.PictureSizeMode = 1
'  .LOT.Visible = False
'  .MPEX.Visible = False
'  .MP1.Visible = False
'
'  .LZTKlabel.Visible = False
'
'  .FS.Visible = False
'  .FU.Visible = False
'  .FC.Visible = False
'  .FD.Visible = False
'  .FE.Visible = False
'  .FF.Visible = False
'
'  .L10.Visible = False
'  .DiagSKou.Visible = False
'  .DiagUKou.Visible = False
'  .DiagCKou.Visible = False
'  .DiagDKou.Visible = False
'  .DiagEKou.Visible = False
'
'  .L11.Visible = False
'  .LateralS.Visible = False
'  .LateralU.Visible = False
'  .LateralC.Visible = False
'  .LateralD.Visible = False
'  .LateralE.Visible = False
'
'  For i = 1 To 8
'    UF1("Frm" & i).Visible = False
'  Next i
''  .Image1.Height = Image1.Height * 0.9
''  .Image1.width = Image1.width * 0.9
'
'  Set .Image1.Picture = CreatePictureFromClipboard
'  Sleep 8
'  .Repaint
'  Set .Image1.Picture = CreatePictureFromClipboard
'  Sleep 8
'  .Repaint
'End With
    Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
    Dim ws As Worksheet
    Dim i As Long
    Dim tempPicture As StdPicture
    Dim retryCount As Integer
    #If VBA7 Then
        Dim originalWidth As LongPtr, originalHeight As LongPtr
        Dim targetWidth As LongPtr, targetHeight As LongPtr
    #Else
        Dim originalWidth As Long, originalHeight As Long
        Dim targetWidth As Long, targetHeight As Long
    #End If
    
    ' プログレス表示開始
    Call ShowProgress("データ準備中...")
    
    '入力なし、シートなしの場合からの確認時対策
    If (WsCheck("不良要因印刷" & "_" & Lname) = False) Then
        Call AddWS("不良要因印刷" & "_" & Lname)
        Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
    ElseIf ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Cells(1, 2) <> UF1.Lot Then
        Call AddWS("不良要因印刷" & "_" & Lname)
        Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
    End If
    
    Set ws = ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname)
    
    Call ShowProgress("画像データをコピー中...")
    
    ' クリップボードをクリア
    Application.CutCopyMode = False
    DoEvents
    Sleep 300
    
    ' セル範囲をコピー（複数回試行）
    retryCount = 0
    Do While retryCount < 3
    On Error Resume Next
        Call CopyRangeAsHighResolutionImage(ws, ws.Range(ws.Cells(1, 2), ws.Cells(11, 24)))
        If Err.Number <> 0 Then
            ' 高解像度コピー関数がない場合は通常のコピー
            Err.clear
            ws.Range(ws.Cells(1, 2), ws.Cells(11, 24)).Copy
        End If
        On Error GoTo 0
        DoEvents
        Sleep 500 ' コピー完了を確実に待つ
        
        ' クリップボードにデータがあるか確認
        If IsClipboardFormatAvailable(CF_BITMAP) <> 0 Or IsClipboardFormatAvailable(CF_ENHMETAFILE) <> 0 Then
            Exit Do
        End If
        
        retryCount = retryCount + 1
        Call ShowProgress("画像データをコピー中... (再試行 " & retryCount & "/3)")
    Loop
    
    Call ShowProgress("画像を取得中...")
    
    With UF1
        ' 画像取得の試行
        retryCount = 0
        Do While retryCount < 8
            Set tempPicture = CreatePictureFromClipboard
            If Not tempPicture Is Nothing Then Exit Do
            
            retryCount = retryCount + 1
            Call ShowProgress("画像を取得中... (試行 " & retryCount & "/8)")
            DoEvents
            Sleep 300
            
            ' 3回目以降は再コピーも試行
            If retryCount >= 3 Then
            On Error Resume Next
                Call CopyRangeAsHighResolutionImage(ws, ws.Range(ws.Cells(1, 2), ws.Cells(11, 24)))
                If Err.Number <> 0 Then
                    Err.clear
                    ws.Range(ws.Cells(1, 2), ws.Cells(11, 24)).Copy
                End If
                On Error GoTo 0
                DoEvents
                Sleep 300
            End If
        Loop
        
        ' 画像取得失敗時のエラーハンドリング
        If tempPicture Is Nothing Then
            Call HideProgress
            MsgBox "画像の取得に失敗しました。再度お試しください。", vbExclamation
            Exit Sub
        End If
        
        Call ShowProgress("画面を更新中...")
        
        ' Image1の初期設定（サイズ固定化のため）
        .Image1.Visible = False ' 一旦非表示にして設定
        .Image1.AutoSize = False ' AutoSizeを無効化
        .Image1.PictureSizeMode = 1 ' fmPictureSizeModeStretch
        
        ' 固定サイズを設定（必要に応じて調整）
        targetWidth = 850  ' 目標幅（ポイント単位）
        targetHeight = 450 ' 目標高さ（ポイント単位）
        
        .Image1.Width = targetWidth
        .Image1.Height = targetHeight
        
        ' UI要素の非表示処理（エラーハンドリング強化）
        On Error Resume Next
        .Lot.Visible = False
        If Err.Number <> 0 Then Debug.Print "LOT非表示エラー: " & Err.Description: Err.clear
        
        
        .MPEX.Visible = False
        .MPEX.Left = -1000
        If Err.Number <> 0 Then Debug.Print "MPEX非表示エラー: " & Err.Description: Err.clear
        
        .MP1.Visible = False
        .MP1.Left = -1000
        If Err.Number <> 0 Then Debug.Print "MP1非表示エラー: " & Err.Description: Err.clear
        
        .LZTKlabel.Visible = False
        If Err.Number <> 0 Then Debug.Print "LZTKlabel非表示エラー: " & Err.Description: Err.clear
        
        .FS.Visible = False
        If Err.Number <> 0 Then Debug.Print "FS非表示エラー: " & Err.Description: Err.clear
        
        .FU.Visible = False
        If Err.Number <> 0 Then Debug.Print "FU非表示エラー: " & Err.Description: Err.clear
        
        .FC.Visible = False
        If Err.Number <> 0 Then Debug.Print "FC非表示エラー: " & Err.Description: Err.clear
        
        .FD.Visible = False
        If Err.Number <> 0 Then Debug.Print "FD非表示エラー: " & Err.Description: Err.clear
        
        .FE.Visible = False
        If Err.Number <> 0 Then Debug.Print "FE非表示エラー: " & Err.Description: Err.clear
        
        .FF.Visible = False
        If Err.Number <> 0 Then Debug.Print "FF非表示エラー: " & Err.Description: Err.clear
        
        .L10.Visible = False
        If Err.Number <> 0 Then Debug.Print "L10非表示エラー: " & Err.Description: Err.clear
        
        .DiagSKou.Visible = False
        If Err.Number <> 0 Then Debug.Print "DiagSKou非表示エラー: " & Err.Description: Err.clear
        
        .DiagUKou.Visible = False
        If Err.Number <> 0 Then Debug.Print "DiagUKou非表示エラー: " & Err.Description: Err.clear
        
        .DiagCKou.Visible = False
        If Err.Number <> 0 Then Debug.Print "DiagCKou非表示エラー: " & Err.Description: Err.clear
        
        .DiagDKou.Visible = False
        If Err.Number <> 0 Then Debug.Print "DiagDKou非表示エラー: " & Err.Description: Err.clear
        
        .DiagEKou.Visible = False
        If Err.Number <> 0 Then Debug.Print "DiagEKou非表示エラー: " & Err.Description: Err.clear
        
        .L11.Visible = False
        If Err.Number <> 0 Then Debug.Print "L11非表示エラー: " & Err.Description: Err.clear
        
        .LateralS.Visible = False
        If Err.Number <> 0 Then Debug.Print "LateralS非表示エラー: " & Err.Description: Err.clear
        
        .LateralU.Visible = False
        If Err.Number <> 0 Then Debug.Print "LateralU非表示エラー: " & Err.Description: Err.clear
        
        .LateralC.Visible = False
        If Err.Number <> 0 Then Debug.Print "LateralC非表示エラー: " & Err.Description: Err.clear
        
        .LateralD.Visible = False
        If Err.Number <> 0 Then Debug.Print "LateralD非表示エラー: " & Err.Description: Err.clear
        
        .LateralE.Visible = False
        If Err.Number <> 0 Then Debug.Print "LateralE非表示エラー: " & Err.Description: Err.clear
        On Error GoTo 0
        
        ' フレーム要素の非表示処理
        For i = 1 To 8
            On Error Resume Next
            .Controls("Frm" & i).Visible = False
            If Err.Number <> 0 Then Debug.Print "Frm" & i & "非表示エラー: " & Err.Description: Err.clear
            On Error GoTo 0
        Next i
        
        ' 処理完了を確認
        DoEvents
        Sleep 100
        
        Call ShowProgress("画像を設定中...")
        
        ' 画像の設定
        Set .Image1.Picture = tempPicture
        DoEvents
        Sleep 200
        
        ' Image1を表示
        .Image1.Visible = True
        
        ' フォーム更新
        .Repaint
        DoEvents
        Sleep 200
        
        ' 最終確認と修正
        If .Image1.Picture Is Nothing Then
            Call ShowProgress("画像を再設定中...")
            
            ' 最終リトライ
            ws.Range(ws.Cells(1, 2), ws.Cells(11, 24)).Copy
            DoEvents
            Sleep 500
            
            Set tempPicture = CreatePictureFromClipboard
            If Not tempPicture Is Nothing Then
                Set .Image1.Picture = tempPicture
                .Repaint
                DoEvents
            Else
                Call HideProgress
                MsgBox "画像の表示に失敗しました。", vbExclamation
                Exit Sub
            End If
        End If
        
        ' サイズの最終調整（画像が設定された後）
        .Image1.Width = targetWidth
        .Image1.Height = targetHeight
        .Repaint
        
    End With
    
    ' 完了表示
    Call ShowProgress("完了しました")
    Sleep 1000
    Call HideProgress
    
    ' クリーンアップ
    Application.CutCopyMode = False
    Set tempPicture = Nothing
    Set ws = Nothing
End Sub

Private Sub Yama_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'入力制限
Select Case KeyAscii.Value
Case 48 To 57 '0~9
Case 65 To 90 'A~Z
Case 97 To 122 'a~z
KeyAscii.Value = KeyAscii.Value - 32 '大文字強制変換
Case Else
KeyAscii.Value = 0
End Select
End Sub
Private Sub Yama_Change()

If UF1.Load.Caption = "off" Then Exit Sub
Dim UfHeight As Long: UfHeight = UF1.Height
Dim SizeUf As Variant

If IsNumeric(UF1.Yama.Text) Then UF1.Yama = Format((UF1.Yama.Text), "0")

Dim i As Long
Select Case UF1.Yama.Text
Case 1, 2, 3, 4
  UF1.Yama.Tag = "UF1" '2026.07.27_追加
Case ""
  UF1.Yama.Tag = "UF1" '2026.07.27_追加
  Exit Sub
Case Else
  UF1.Yama.Tag = "別フォーム" '2026.07.27_追加
  MsgBox "別フォームで検査してください", vbCritical
  Exit Sub
End Select

'＜＜＜＜＜検査箇所回数変更の場合InspectionSelect,AtCountSelectの修正も必要＞＞＞＞＞＞＞＞

Me.Loading.Visible = True
Me.Repaint
  
If UF1.Yama.Text = "4" Then
  Call Six_Display("on")
  If UF1.Sumisei.Caption <> "" Then Call LSelect("住精")
  SizeUf = InitialSetting(UF1)
  Exit Sub
Else
  Call Six_Display("off")
End If
  
  If UF1.InspeButt.Visible And UF1.InspeButt.Caption = "検入≧26" Then
    Select Case UF1.Yama.Text '2022.03.22**追加Load時うまく読み込めない対策
    Case 1
      UF1.S = "S"
      UF1.C = ""
      UF1.E = "E"
      UF1.FS = "S"
      UF1.FC = ""
      UF1.FE = "E"
      UF1.OPS1.Caption = "S" '2021.09.06
      UF1.OPS3.Caption = ""
      UF1.OPS5.Caption = "E"
      UF1.OPE1.Caption = "S"
      UF1.OPE3.Caption = ""
      UF1.OPE5.Caption = "E"
      UF1.U = ""
      UF1.FU = ""
      UF1.D = ""
      UF1.FD = ""
    Case 2
      UF1.S = "2-2-S"
      UF1.U = ""
      UF1.C = ""
      UF1.D = ""
      UF1.E = "2-1-E"
      UF1.FS = "2-2-S"
      UF1.FU = ""
      UF1.FC = ""
      UF1.FD = ""
      UF1.FE = "2-1-E"
      UF1.OPS1.Caption = "2-2-S"
      UF1.OPS2.Caption = ""
      UF1.OPS3.Caption = ""
      UF1.OPS4.Caption = ""
      UF1.OPS5.Caption = "2-1-E"
      UF1.OPE1.Caption = "2-2-S"
      UF1.OPE2.Caption = ""
      UF1.OPE3.Caption = ""
      UF1.OPE4.Caption = ""
      UF1.OPE5.Caption = "2-1-E"
      UF1.U = ""
      UF1.FU = ""
      UF1.D = ""
      UF1.FD = ""
    End Select
  ElseIf UF1.InspeButt.Visible And UF1.InspeButt.Caption = "検入≦25" Then
    Select Case UF1.Yama.Text
    Case 1
      UF1.S = "S"
      UF1.C = "2"
      UF1.E = "E"
      UF1.FS = "S"
      UF1.FC = "2"
      UF1.FE = "E"
      UF1.OPS1.Caption = "S" '2021.09.06
      UF1.OPS3.Caption = "2"
      UF1.OPS5.Caption = "E"
      UF1.OPE1.Caption = "S"
      UF1.OPE3.Caption = "2"
      UF1.OPE5.Caption = "E"
      UF1.U = ""
      UF1.FU = ""
      UF1.D = ""
      UF1.FD = ""
    Case 2
      UF1.S = "2-2-S"
      UF1.U = ""
      UF1.C = "2-2-E"
      UF1.D = ""
      UF1.E = "2-1-E"
      UF1.FS = "2-2-S"
      UF1.FU = ""
      UF1.FC = "2-2-E"
      UF1.FD = ""
      UF1.FE = "2-1-E"
      UF1.OPS1.Caption = "2-2-S"
      UF1.OPS2.Caption = ""
      UF1.OPS3.Caption = "2-2-E"
      UF1.OPS4.Caption = ""
      UF1.OPS5.Caption = "2-1-E"
      UF1.OPE1.Caption = "2-2-S"
      UF1.OPE2.Caption = ""
      UF1.OPE3.Caption = "2-2-E"
      UF1.OPE4.Caption = ""
      UF1.OPE5.Caption = "2-1-E"
      UF1.U = ""
      UF1.FU = ""
      UF1.D = ""
      UF1.FD = ""
    End Select
  End If
  
  '2022.11.08_住精の描写が先だと5回検査の描写にならないLselectの作り直しが本来はいる.......
  '2023.6.20_もう限界だったのでいろいろ修正Lselectやや作り直した
  If UF1.Sumisei.Caption <> "" Then 'ｽﾐﾄﾓﾌﾗｸﾞ 2021.07.13_追加**2022.8.17*****ﾗﾃﾗﾙﾎﾞｰ測定の納入先が追加される為"住友精密"から<>に変更
    Call LSelect("住精")
    GoTo skipSmisei
  End If

  If Not UF1.Load.Caption = "off" Then Call LSelect("薄板")

If UF1.薄板.Caption = "T" And UF1.Yama.Text = "1" Then
  UF1.OPS1.Caption = "S"
  UF1.OPS2.Caption = ""
  UF1.OPS3.Caption = "2"
  UF1.OPS4.Caption = ""
  UF1.OPS5.Caption = "E"
  UF1.OPE1.Caption = "S"
  UF1.OPE2.Caption = ""
  UF1.OPE3.Caption = "2"
  UF1.OPE4.Caption = ""
  UF1.OPE5.Caption = "E"
ElseIf UF1.薄板.Caption = "T" And UF1.Yama.Text = "2" Then
  UF1.OPS1.Caption = "2-2-S"
  UF1.OPS2.Caption = ""
  UF1.OPS3.Caption = "2-2-E"
  UF1.OPS4.Caption = ""
  UF1.OPS5.Caption = "2-1-E"
  UF1.OPE1.Caption = "2-2-S"
  UF1.OPE2.Caption = ""
  UF1.OPE3.Caption = "2-2-E"
  UF1.OPE4.Caption = ""
  UF1.OPE5.Caption = "2-1-E"
End If


For i = 3 To 4
  Select Case True
  Case UF1.Yama.Text = "3" '先に実行
    UF1("C" & i).Enabled = True
    UF1("C" & i).Visible = True
    UF1.DiagU.Visible = True
    UF1.DiagU.Enabled = True
    UF1.DiagE.Visible = True
    UF1.DiagE.Enabled = True
  End Select
Next i
'厚検査箇所2,4の選択
'Loadtext時実行され、冷延なしの場合2か所になり戻せなくなる_2022.01.11_AtCountSelect修正
If Not UF1.Load.Caption = "off" Then
  Select Case UF1.AtButt.Caption
  Case "厚2ヵ所": Call AtCountSelect(4)
  Case "厚4ヵ所": Call AtCountSelect(2)
  End Select

'  If InStr(UF1.Reien.Caption, "2HI") > 0 Then Call AtCountSelect(4) '2023.3.16*2HIコースで板厚4回の場合

End If
'検査回数2,3の選択
If Not UF1.Load.Caption = "off" Then Call InspectionSelect("表示") 'チェンジ後は必ず表示させる

''2022.11.08_住精の描写が先だと5回検査の描写にならないLselectの作り直しが本来はいる.......
'If UF1.Sumisei.Caption <> "" Then 'ｽﾐﾄﾓﾌﾗｸﾞ 2021.07.13_追加**2022.8.17*****ﾗﾃﾗﾙﾎﾞｰ測定の納入先が追加される為"住友精密"から<>に変更
'  Call LSelect("住精")
'End If

skipSmisei:


SizeUf = InitialSetting(UF1)
Me.Loading.Visible = False
Me.Repaint


Call UserForm_Click 'フォーカス対策

If UF1.Load.Caption <> "off" Then
  Call Posting
  Dim path As String: path = HistoryPath
  Dim Temp As String: Temp = "*UFKen:山数ﾁｪﾝｼﾞ:" & UF1.Yama.Text
  Select Case True
  Case UF1.L1Work:   Call WriteLog(Temp, "", path & "L-1\L-1_Log")
  Case UF1.LVC:      Call WriteLog(Temp, "", path & "LVC\LVC_Log")
  Case UF1.HVC:      Call WriteLog(Temp, "", path & "HVC\HVC_Log")
  Case UF1.Bala:     Call WriteLog(Temp, "", path & "バランサー\バランサー_Log")
  Case UF1.Tott:     Call WriteLog(Temp, "", path & "トット\トット_Log")
  Case UF1.Test:     Call WriteLog(Temp, "", path & "Test\Test_Log")
  Case Else:         Call WriteLog(Temp, "", path & "丸徳\丸徳_Log")
  End Select
End If
End Sub
Private Sub Yama_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'山数選択時
With UF1
  .PicFM.Visible = True
  .PicFM.tOP = 120
  .PicFM.Left = 12
  .PicFM.Width = 624
  .PicFM.Height = 186
  .OrdNoCB.Visible = False
  .NoL.Visible = False
  .PackingForm.Visible = False
  .PackingComment.Visible = False
  .MPEX.Visible = False
End With
Me.Repaint
End Sub
Private Sub SlectFM_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'マウスポイントが山数選択から外れた時
With UF1
  .PicFM.tOP = 120
  .PicFM.Left = 12
  .PicFM.Width = 0
  .PicFM.Height = 0
  .PicFM.Visible = False
  .OrdNoCB.Visible = True
  .NoL.Visible = True
  .PackingForm.Visible = True
  .PackingComment.Visible = True
  .MPEX.Visible = True
End With

'On Error Resume Next
'UF1.Manual.SetFocus
'UF1.StartTB.SetFocus
End Sub
Private Sub ErFM_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
UF1.PicFM.tOP = 120
UF1.PicFM.Left = 12
UF1.PicFM.Width = 0
UF1.PicFM.Height = 0
UF1.PicFM.Visible = False
UF1.OrdNoCB.Visible = True
UF1.NoL.Visible = True
UF1.PackingForm.Visible = True
UF1.PackingComment.Visible = True
End Sub
Private Sub YamaLB_Click()
If UF1.ATU = "" Then Exit Sub
Call DrawListBoxContent(UF1, UF1.ATU)
End Sub

Private Sub ZS_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub ZU_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub ZC_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub ZD_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub ZE_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub ZF_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub

Private Sub ZS_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMS)
Call TagSet(Me.ZS)
ChangeBackColor ZS
End Sub
Private Sub ZU_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMU)
Call TagSet(Me.ZU)
ChangeBackColor ZU
End Sub
Private Sub ZC_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMC)
Call TagSet(Me.ZC)
ChangeBackColor ZC
End Sub
Private Sub ZD_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMD)
Call TagSet(Me.ZD)
ChangeBackColor ZD
End Sub
Private Sub ZE_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FME)
Call TagSet(Me.ZE)
ChangeBackColor ZE
End Sub
Private Sub ZF_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMF)
Call TagSet(Me.ZF)
ChangeBackColor ZF
End Sub

Private Sub GS_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub GU_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub GC_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub GD_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub GE_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub
Private Sub GF_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ResetBackColor
End Sub

Private Sub GS_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMS)
Call TagSet(Me.GS) '2024.2.14
ChangeBackColor GS
End Sub
Private Sub GU_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMU)
Call TagSet(Me.GU)
ChangeBackColor GU
End Sub
Private Sub GC_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMC)
Call TagSet(Me.GC)
ChangeBackColor GC
End Sub
Private Sub GD_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMD)
Call TagSet(Me.GD)
ChangeBackColor GD
End Sub
Private Sub GE_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FME)
Call TagSet(Me.GE)
ChangeBackColor GE
End Sub
Private Sub GF_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMF)
Call TagSet(Me.GF)
ChangeBackColor GF
End Sub

Private Sub CBS1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    ' CBS1の値が前回の値と異なる場合のみLogOutComboを呼び出す
    If CBS1.Value <> Old_CBS1 Then Call LogOutCombo(CBS1)
    ' 現在の値を次回の比較のために保存する
    Old_CBS1 = CBS1.Value
End Sub
Private Sub CBU1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBU1.Value <> Old_CBU1 Then Call LogOutCombo(CBU1)
    Old_CBU1 = CBU1.Value
End Sub
Private Sub CBC1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBC1.Value <> Old_CBC1 Then Call LogOutCombo(CBC1)
    Old_CBC1 = CBC1.Value
End Sub
Private Sub CBD1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBD1.Value <> Old_CBD1 Then Call LogOutCombo(CBD1)
    Old_CBD1 = CBD1.Value
End Sub
Private Sub CBE1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBE1.Value <> Old_CBE1 Then Call LogOutCombo(CBE1)
    Old_CBE1 = CBE1.Value
End Sub
Private Sub CBF1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBF1.Value <> Old_CBF1 Then Call LogOutCombo(CBF1)
    Old_CBF1 = CBF1.Value
End Sub

Private Sub CBS2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBS2.Value <> Old_CBS2 Then Call LogOutCombo(CBS2)
    Old_CBS2 = CBS2.Value
End Sub
Private Sub CBU2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBU2.Value <> Old_CBU2 Then Call LogOutCombo(CBU2)
    Old_CBU2 = CBU2.Value
End Sub
Private Sub CBC2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBC2.Value <> Old_CBC2 Then Call LogOutCombo(CBC2)
    Old_CBC2 = CBC2.Value
End Sub
Private Sub CBD2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBD2.Value <> Old_CBD2 Then Call LogOutCombo(CBD2)
    Old_CBD2 = CBD2.Value
End Sub
Private Sub CBE2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBE2.Value <> Old_CBE2 Then Call LogOutCombo(CBE2)
    Old_CBE2 = CBE2.Value
End Sub
Private Sub CBF2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBF2.Value <> Old_CBF2 Then Call LogOutCombo(CBF2)
    Old_CBF2 = CBF2.Value
End Sub

Private Sub CBS3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBS3.Value <> Old_CBS3 Then Call LogOutCombo(CBS3)
    Old_CBS3 = CBS3.Value
End Sub
Private Sub CBU3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBU3.Value <> Old_CBU3 Then Call LogOutCombo(CBU3)
    Old_CBU3 = CBU3.Value
End Sub
Private Sub CBC3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBC3.Value <> Old_CBC3 Then Call LogOutCombo(CBC3)
    Old_CBC3 = CBC3.Value
End Sub
Private Sub CBD3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBD3.Value <> Old_CBD3 Then Call LogOutCombo(CBD3)
    Old_CBD3 = CBD3.Value
End Sub
Private Sub CBE3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBE3.Value <> Old_CBE3 Then Call LogOutCombo(CBE3)
    Old_CBE3 = CBE3.Value
End Sub
Private Sub CBF3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBF3.Value <> Old_CBF3 Then Call LogOutCombo(CBF3)
    Old_CBF3 = CBF3.Value
End Sub

Private Sub CBS4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBS4.Value <> Old_CBS4 Then Call LogOutCombo(CBS4)
    Old_CBS4 = CBS4.Value
End Sub
Private Sub CBU4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBU4.Value <> Old_CBU4 Then Call LogOutCombo(CBU4)
    Old_CBU4 = CBU4.Value
End Sub
Private Sub CBC4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBC4.Value <> Old_CBC4 Then Call LogOutCombo(CBC4)
    Old_CBC4 = CBC4.Value
End Sub
Private Sub CBD4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBD4.Value <> Old_CBD4 Then Call LogOutCombo(CBD4)
    Old_CBD4 = CBD4.Value
End Sub
Private Sub CBE4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBE4.Value <> Old_CBE4 Then Call LogOutCombo(CBE4)
    Old_CBE4 = CBE4.Value
End Sub
Private Sub CBF4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.Load = "on"
    ResetBackColor
    If CBF4.Value <> Old_CBF4 Then Call LogOutCombo(CBF4)
    Old_CBF4 = CBF4.Value
End Sub

Private Sub CBS1_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMS)
ChangeBackColor CBS1
End Sub
Private Sub CBU1_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMU)
ChangeBackColor CBU1
End Sub
Private Sub CBC1_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMC)
ChangeBackColor CBC1
End Sub
Private Sub CBD1_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMD)
ChangeBackColor CBD1
End Sub
Private Sub CBE1_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FME)
ChangeBackColor CBE1
End Sub
Private Sub CBF1_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMF)
ChangeBackColor CBF1
End Sub

Private Sub CBS2_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMS)
ChangeBackColor CBS2
End Sub
Private Sub CBU2_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMU)
ChangeBackColor CBU2
End Sub
Private Sub CBC2_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMC)
ChangeBackColor CBC2
End Sub
Private Sub CBD2_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMD)
ChangeBackColor CBD2
End Sub
Private Sub CBE2_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FME)
ChangeBackColor CBE2
End Sub
Private Sub CBF2_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMF)
ChangeBackColor CBF2
End Sub

Private Sub CBS3_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMS)
ChangeBackColor CBS3
End Sub
Private Sub CBU3_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMU)
ChangeBackColor CBU3
End Sub
Private Sub CBC3_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMC)
ChangeBackColor CBC3
End Sub
Private Sub CBD3_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMD)
ChangeBackColor CBD3
End Sub
Private Sub CBE3_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FME)
ChangeBackColor CBE3
End Sub
Private Sub CBF3_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMF)
ChangeBackColor CBF3
End Sub

Private Sub CBS4_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMS)
ChangeBackColor CBS4
End Sub
Private Sub CBU4_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMU)
ChangeBackColor CBU4
End Sub
Private Sub CBC4_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMC)
ChangeBackColor CBC4
End Sub
Private Sub CBD4_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMD)
ChangeBackColor CBD4
End Sub
Private Sub CBE4_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FME)
ChangeBackColor CBE4
End Sub
Private Sub CBF4_Enter()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMF)
ChangeBackColor CBF4
End Sub

Private Sub CBS1_Change()
Call CB_UpDownAdd(CBS1)
End Sub
Private Sub CBU1_Change()
Call CB_UpDownAdd(CBU1)
End Sub
Private Sub CBC1_Change()
Call CB_UpDownAdd(CBC1)
End Sub
Private Sub CBD1_Change()
Call CB_UpDownAdd(CBD1)
End Sub
Private Sub CBE1_Change()
Call CB_UpDownAdd(CBE1)
End Sub
Private Sub CBF1_Change()
Call CB_UpDownAdd(CBF1)
End Sub

Private Sub CBS3_Change()
Call CBG_UpDownAdd(CBS3)
End Sub
Private Sub CBU3_Change()
Call CBG_UpDownAdd(CBU3)
End Sub
Private Sub CBC3_Change()
Call CBG_UpDownAdd(CBC3)
End Sub
Private Sub CBD3_Change()
Call CBG_UpDownAdd(CBD3)
End Sub
Private Sub CBE3_Change()
Call CBG_UpDownAdd(CBE3)
End Sub
Private Sub CBF3_Change()
Call CBG_UpDownAdd(CBF3)
End Sub

Private Sub CommandButton10_Click()
On Error Resume Next
''-----2021.03.04_Focus対策
'UF1.Worker.Enabled = True
'UF1.INP.Enabled = True
'UF1.Manual.Enabled = True
'UF1.StartTB.Enabled = True
'
Select Case True
Case UF1.Bala: GoTo line
Case Else:
  If MsgBox("Lot登録しますか", vbYesNo + vbQuestion, "確認") = vbYes Then
    If Me.Lot <> "" Then
      If MsgBox("ロット分割品ですか？", vbYesNo + vbQuestion, "確認") = vbYes Then
        If MsgBox("間違いなくロット分割ですか？？", vbYesNo + vbQuestion, "確認") = vbYes Then
          Me.SplitFlag.Caption = "Lot分割" 'Lot分割フラグ_2023.2.27
        Else
        GoTo line 'いいえ　で消す
        End If
      Else
      GoTo line 'いいえ　で消す
      End If
    Else
    GoTo line
    End If
  Else
    Exit Sub
  End If
End Select

line:
  BlinkLabel1.BlinkStop '点滅ｽﾄｯﾌﾟ
'  UF1.Repaint
  '-----2021.03.04_Focus対策
'  UF1.Worker.Enabled = False
'  UF1.INP.Enabled = False
'  UF1.Manual.Enabled = False
'  UF1.StartTB.Enabled = False
'
  UF1.Lot.BackColor = &HC0FFFF
  UF1.Lot.Locked = False
  UF1.Lot.Enabled = True
  UF1.Lot = ""
  
  UF1.VC_L_A.Tag = ""
  UF1.VC_L_B.Tag = ""
  
  UF1.Yama.Text = ""
  UF1.KenNo.Caption = ""
  UF1.HizCome.Caption = ""
  UF1.HizCome.Tag = ""
  UF1.CB = ""
  UF1.QF.Caption = ""
  UF1.LastHot.Caption = ""
  UF1.flatnesText = ""
  UF1.dataLoding.Caption = "" '*****履歴呼び出しがあった場合のリセット
  UF1.DebugTaikaku = ""
  UF1.DebugLate = ""
  
  AppActivate Application.Caption
  DoEvents
  UF1.Machin = ""
  
  UF1.CommandButton10.SetFocus
  UF1.Lot.SetFocus
End Sub

Private Sub CommandButton13_Click()
UF1.PasswordTex.Visible = True
If UF1.PasswordTex = "weekend99" Then
 If MsgBox("隠しｵﾌﾞｼﾞｪｸﾄを表示しますか？", vbYesNo) = vbYes Then
  UF1.HiddenFrame.Visible = False
  UF1.MP1.Test1.Visible = True
  UF1.MP1.Test2.Visible = True
  UF1.MP1.Test3.Visible = True
  
  SF3.Show vbModeless
  Call WindowMenuButtonOn(SF3)
 End If
ElseIf UF1.PasswordTex = "jump" Then '校正ｽｷｯﾌﾟﾌﾗｸﾞ
  UF1.WorkButton.Tag = "jump"
Else
 UF1.PasswordTex.SetFocus
End If
End Sub


Private Sub CommandButton14_Click()
Application.Visible = True 'Excel表示
Call KenSheets
End Sub

Private Sub CommandButton6_Click()
On Error Resume Next
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim Worker1: Worker1 = Split(UF1.T1, "_") '複数人の登録があった場合後の登録者を記録する
Dim Worker2: Worker2 = Split(UF1.T2, "_")
Dim Worker3: Worker3 = Split(UF1.T3, "_")
Dim Day1 As Date, Day2 As Date, Day3 As Date
If UF1.Day1 <> "" Then Day1 = Replace((UF1.Day1), ".", "/") '最終日付ﾁｪｯｸ
If UF1.Day2 <> "" Then Day2 = Replace((UF1.Day2), ".", "/")
If UF1.Day3 <> "" Then Day3 = Replace((UF1.Day3), ".", "/")
Select Case Application.WorksheetFunction.Max(Day1, Day2, Day3) '登録時間が一番遅い
Case Day1: UF校正.Kname = Worker1(UBound(Worker1))
Case Day2: UF校正.Kname = Worker2(UBound(Worker2))
Case Day3: UF校正.Kname = Worker3(UBound(Worker3))
End Select
UF校正.LineName = Lname
UF校正.Show vbModeless
End Sub

Private Sub CommandButton7_Click()
Call UFsizeSave(UF1) '2024.8.5
UF管理.Show vbModeless
End Sub

Private Sub CommandButton8_Click()
Dim lottemp As String: lottemp = UF1.Lot
Dim KenTemp As String: KenTemp = UF1.KenNo.Caption
Dim etcTemp As String: etcTemp = UF1.DBflag.Text

UF品質基準.LineName = UF1.LineName.Caption & UF1.LineNo.Caption
UF品質基準.T1 = UF1.Lot
UF品質基準.FlagAtu = UF1.ATU
UF品質基準.MP.Page1.Visible = True
UF品質基準.MP.Page2.Visible = True
'厚板,薄板の判定を行う
'Page1=0　Page2=1
If UF品質基準.FlagAtu = "" Then Exit Sub

  '<<<<<<<<<<<<<<<<<<<旧バージョン>>>>>>>>>>>>>>>>>>>>>>>>>
'2023.4.12_品証管理がしやすいようDB関数作成
'If UF品質基準.FlagAtu <= 4 Then
'  UF品質基準.MP.value = 1
'  UF品質基準.MP.Page1.Visible = False
'  Call 表面品質基準(UF品質基準.T1)
'  If IsNumeric(UF1.Atu) Then Call 平坦度(Val(UF1.Atu), Val(UF1.Hab), Val(UF1.Tak), UF1.Zai, UF1.Tyu, UF1.Youto) '2023.1.31_引数追加
'Else
'  UF品質基準.MP.value = 0
'  UF品質基準.MP.Page2.Visible = False
'  Call 品質グレード
'  If IsNumeric(UF1.Atu) Then Call 平坦度(Val(UF1.Atu), Val(UF1.Hab), Val(UF1.Tak), UF1.Zai, UF1.Tyu, UF1.Youto)
'End If
  '---------------------------------------------------------------------------------------
'2023.4.12_品証管理がしやすいようDB関数作成
If UF品質基準.FlagAtu >= 4 Then
  UF品質基準.MP.Value = 0
Else
  UF品質基準.MP.Value = 1
End If

''----------UF品質基準InitializeでDBdisplay外観_H実行に変更**2024.4.15
'
''Call DBdisplay歪(UF1.Lot, UF1.DBflag.text) '旧バージョン

Dim SetTable As String: SetTable = AddPagesForSubFolders(UF品質基準.MPH) 'ページ追加
UF品質基準.TableTB = "「追加ページ一覧」" & vbCrLf & vbCrLf & SetTable

Call LoadImagesIntoMultiPageFromTemp(UF品質基準.MPH) '写真追加

UF品質基準.Show vbModeless '歪検索より先に表示

UF1.Loading.Visible = True
UF1.Repaint

Dim QpArry As Variant:         If UF1.Lot <> "" Then QpArry = DistArray(lottemp, KenTemp, etcTemp, "")  'ゲージ、幅歪、丈歪、中、耳、コメント、参照元、番号、フラグ
'Dim QpArryG As Variant:        If UF1.Lot <> "" Then QpArryG = DistArray(lottemp, KenTemp, "取引", etcTemp, "")
'Dim UpdatedQpArryD As Variant: If UF1.Lot <> "" Then UpdatedQpArryD = UpdateQpArryD(QpArry, QpArryG)

If IsEmptyArray(QpArry) = False Then Call FormReflectDist(QpArry, "") 'フォーム反映
UF1.Loading.Visible = False
UF1.Repaint

If UF品質基準.FlagAtu >= 8 Then GradeArray = DBdisplay外観_H(UF1.Lot, UF1.DBflag.Text)
End Sub

Private Sub CommandButton9_Click()
'On Error Resume Next’様子見停止_2022.06.23
'Call HiddenSheets
'Application.Visible = True
Unload Me
End Sub

Private Sub Day1_Change()
If UF校正.Visible = True Then Exit Sub
If Not WorkeDayMax = "1" Then Exit Sub
End Sub
Private Sub Day2_Change()
If UF校正.Visible = True Then Exit Sub
If Not WorkeDayMax = "2" Then Exit Sub
End Sub
Private Sub Day3_Change()
If UF校正.Visible = True Then Exit Sub
If Not WorkeDayMax = "3" Then Exit Sub
End Sub

'Private Sub FMS_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'If Not UF1.Load = "off" Then
'On Error Resume Next
'UF1.Manual.SetFocus
'UF1.StartTB.SetFocus
'End If
'End Sub
'Private Sub FMU_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'If Not UF1.Load = "off" Then
'On Error Resume Next
'UF1.Manual.SetFocus
'UF1.StartTB.SetFocus
'End If
'End Sub
'Private Sub FMC_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'If Not UF1.Load = "off" Then
'On Error Resume Next
'UF1.Manual.SetFocus
'UF1.StartTB.SetFocus
'End If
'End Sub
'Private Sub FMD_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'If Not UF1.Load = "off" Then
'On Error Resume Next
'UF1.Manual.SetFocus
'UF1.StartTB.SetFocus
'End If
'End Sub
'Private Sub FME_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'If Not UF1.Load = "off" Then
'On Error Resume Next
'UF1.Manual.SetFocus
'UF1.StartTB.SetFocus
'End If
'End Sub
Private Sub FMS_Click()
UF1.MPEX.Tag = "Click"
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMS)
End Sub
Private Sub FMU_Click()
UF1.MPEX.Tag = "Click"
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMU)
End Sub
Private Sub FMC_Click()
UF1.MPEX.Tag = "Click"
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMC)
End Sub
Private Sub FMD_Click()
UF1.MPEX.Tag = "Click"
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMD)
End Sub
Private Sub FME_Click()
UF1.MPEX.Tag = "Click"
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FME)
End Sub
Private Sub FMF_Click()
UF1.MPEX.Tag = "Click"
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMF)
End Sub

Private Sub Hab_Change()
If IsNumeric(UF1.Hab) = False Then UF1.Hab = ""
End Sub

Private Sub Hm_Change()
If IsNumeric(UF1.Hm) = False Then UF1.Hm = ""
Call Form_Formula
End Sub

Private Sub Hp_Change()
If IsNumeric(UF1.Hp) = False Then UF1.Hp = ""
Call Form_Formula
End Sub

Private Sub IzyouFl_Click()
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("異常ロット用" & "_" & Lname) Then
  If ThisWorkbook.Sheets("異常ロット用" & "_" & Lname).Cells(1, 1) = UF1.Lot Then
    UF1.IzyouFl.BackColor = RGB(255, 0, 0)
    If MsgBox("異常内容印刷フラグが立っています,取り消しますか？", vbYesNo) = vbYes Then
      If MsgBox("本当に取り消しますか？", vbYesNo) = vbYes Then
        Application.DisplayAlerts = False
          If WsCheck("異常ロット用" & "_" & Lname) Then ThisWorkbook.Sheets("異常ロット用" & "_" & Lname).Delete
          UF1.IzyouFl.BackColor = &H8000000F
        Application.DisplayAlerts = True
      End If
    End If
  End If
End If
End Sub

Private Sub KYB_Click()
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If Register_Check(Format$(Now, "hh:mm:ss"), "") Then '検査員ﾁｪｯｸ *1.ｼｰﾄがない2.3行(Day1,Day2,Day3)に当日以外の登録がある3.その時間に登録が無い
  Call MsgError("検査員登録をして下さい", Lname, 2)
  Exit Sub
End If
Application.Visible = True
DoEvents
UF1.Yama.Text = ""
Call HiddenSheets
Call BookSaClose(ThisWorkbook.path, "日報＆資材計算.xlsm") 'ﾊﾟｽ,閉じたいbook名
UFCask.Show vbModeless
UFCask.LineName.Caption = UF1.LineName.Caption
If UFWorkerName.Visible = True Then Unload UFWorkerName
UF1.Hide
End Sub
Private Sub LZS_Click()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMS)
End Sub
Private Sub LZU_Click()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMU)
End Sub
Private Sub LZC_Click()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMC)
End Sub
Private Sub LZD_Click()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FMD)
End Sub
Private Sub LZE_Click()
If UF1.Load.Caption = "off" Then Exit Sub
Call FMCheck(FME)
End Sub

Private Sub L1Work_Click()
Dim n As String
If UF1.Load.Caption = "off" Then
  n = ""
Else
  n = "表示"
End If
Call Line_AutoSelection(n)
End Sub

Private Sub LVC_Click()
Dim n As String
If UF1.Load.Caption = "off" Then
  n = ""
Else
  n = "表示"
End If
Call Line_AutoSelection(n)
End Sub
Private Sub HVC_Click()
Dim n As String
If UF1.Load.Caption = "off" Then
  n = ""
Else
  n = "表示"
End If
Call Line_AutoSelection(n)
End Sub
Private Sub StartTB_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) '2023.6.21
If (KeyCode = vbKeyTab) Or (KeyCode = vbKeyReturn) Then 'tabかEnter
  If Me.FlagLabel <> 1 Then Exit Sub
  If UF1.StartTB = "" Then Exit Sub
  Me.INP = Me.StartTB
  Me.StartTB = ""
  UF1.Manual.SetFocus
  UF1.StartTB.SetFocus
  KeyCode = 0 'tab無効化
End If
End Sub

Private Sub StartTB_AfterUpdate()
'If Me.FlagLabel <> 1 Then Exit Sub
'If UF1.StartTB = "" Then Exit Sub
'Me.INP = Me.StartTB
'Me.StartTB = ""
'UF1.Manual.SetFocus
'UF1.StartTB.SetFocus
End Sub
Private Sub Tott_Click()
Dim n As String
If UF1.Load.Caption = "off" Then
  n = ""
Else
  n = "表示"
End If
Call Line_AutoSelection(n)
End Sub
Private Sub Bala_Click()
Dim n As String
If UF1.Load.Caption = "off" Then
  n = ""
Else
  n = "表示"
End If
Call Line_AutoSelection(n)
End Sub
Private Sub Test_Click()
Dim n As String
If UF1.Load.Caption = "off" Then
  n = ""
Else
  n = "表示"
End If
Call Line_AutoSelection(n)
End Sub
Private Sub Maru_Click()
Dim n As String
If UF1.Load.Caption = "off" Then
  n = ""
Else
  n = "表示"
End If
Call Line_AutoSelection(n)
End Sub
Private Sub Manual_Change()
If IsNumeric(UF1.Manual) = False Then UF1.Manual = ""
End Sub
Private Sub MP1_Change()
If UF1.MP1.Value = 5 Then Call IM_Select
If UF1.MP1.Value = 1 Then UF1.PicFM.Tag = 0 'LoopJudg用カウントリセット'MinMax
End Sub

Private Sub NewKen_Click()
Application.DisplayAlerts = False
ThisWorkbook.Save '上書き保存
Application.DisplayAlerts = True
DoEvents

BlinkLabel1.BlinkStop '点滅ｽﾄｯﾌﾟ
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If UF1.Lot = "" Then GoTo line
  Dim sh As String: sh = "保存用"
  Dim k As Long: k = ThisWorkbook.Sheets(sh).Cells(Rows.Count, 2).End(xlUp).row
  Dim myRange As Range: Set myRange = ThisWorkbook.Sheets(sh).Range(ThisWorkbook.Sheets(sh).Cells(2, 2), ThisWorkbook.Sheets(sh).Cells(k, 2))
  Dim MyObj As Range, Keyword As String
'  If Lotflag("SIKALOTNOW.accdb", "LOT一覧") Then Call AccUpdate '引当ﾁｪｯｸ
'  If Lotflag("SIKAODRNOW.accdb", "VC一覧") Then Call AccUpdate_Etc("SIKAODRNOW.accdb", "VC一覧")
  If MsgBox("新規検査作業を開始しますか", vbYesNo + vbQuestion, "確認") = vbYes Then
    Keyword = UF1.Lot
    Set MyObj = myRange.Find(Keyword, LookAt:=xlWhole) '部分一致　xlPart　完全一致　xlWhole
    If MyObj Is Nothing Then 'LotNoが保存用シートにない
        sh = "耳用" '耳付き用で再検索_2022.06.10*修正
        k = ThisWorkbook.Sheets(sh).Cells(Rows.Count, 2).End(xlUp).row
        Set myRange = ThisWorkbook.Sheets(sh).Range(ThisWorkbook.Sheets(sh).Cells(2, 2), ThisWorkbook.Sheets(sh).Cells(k, 2))
        Set MyObj = myRange.Find(Keyword, LookAt:=xlWhole)
        If MyObj Is Nothing Then
          Call MsgError("該当LotNoが保存されていません", Lname, 2, "END") '耳付き用でも作業をしていない
        Else '存在した場合
          GoTo line '消去までスキップ
        End If
    Else
line:
    Select Case True
    Case UF1.L1Work:   Call WriteLog("UFKen" & "*新規検査実行", "", HistoryPath & "L-1\L-1_Log")
    Case UF1.LVC:      Call WriteLog("UFKen" & "*新規検査実行", "", HistoryPath & "LVC\LVC_Log")
    Case UF1.HVC:      Call WriteLog("UFKen" & "*新規検査実行", "", HistoryPath & "HVC\HVC_Log")
    Case UF1.Bala:     Call WriteLog("UFKen" & "*新規検査実行", "", HistoryPath & "バランサー\バランサー_Log")
    Case UF1.Tott:     Call WriteLog("UFKen" & "*新規検査実行", "", HistoryPath & "トット\トット_Log")
    Case UF1.Maru:     Call WriteLog("UFKen" & "*新規検査実行", "", HistoryPath & "丸徳\丸徳_Log")
    Case Else:         Call WriteLog("UFKen" & "*新規検査実行", "") '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\検査履歴\Test\Test_Log"
    End Select
    
    UFProgress.Show vbModeless
    UFProgress.CommandButton1.Caption = "NewDataCreate...."
    DoEvents
    Call ProgressBar_count(1, 500)
    
    Call AddWS("作業保存" & "_" & Lname) 'ｼｰﾄ新規
    Call AddWS("WorkerTime" & "_" & Lname)
    
    On Error Resume Next
    Application.DisplayAlerts = False
      If WsCheck("印刷用" & "_" & Lname) Then ThisWorkbook.Sheets("印刷用" & "_" & Lname).Delete '不要ｼｰﾄ削除
      If WsCheck("異常ロット用" & "_" & Lname) Then ThisWorkbook.Sheets("異常ロット用" & "_" & Lname).Delete
      If WsCheck("不良要因印刷" & "_" & Lname) Then ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
     
    If Register_Check(Format$(Now, "hh:mm:ss"), "新規") Then Call AddWS("検査員" & "_" & Lname) '当日以外はｸﾘｱ
    Call ClearLSelectCache 'キャッシュクリア
    
    Application.EnableEvents = False
      UF1("LZS").Caption = "定盤:"
      UF1("LZU").Caption = "定盤:"
      UF1("LZC").Caption = "定盤:"
      UF1("LZD").Caption = "定盤:"
      UF1("LZE").Caption = "定盤:"
      UF1("LZF").Caption = "定盤:"
      
      UF1("LGS").Caption = "ｹﾞｰｼﾞ"
      UF1("LGU").Caption = "ｹﾞｰｼﾞ"
      UF1("LGC").Caption = "ｹﾞｰｼﾞ"
      UF1("LGD").Caption = "ｹﾞｰｼﾞ"
      UF1("LGE").Caption = "ｹﾞｰｼﾞ"
      UF1("LGF").Caption = "ｹﾞｰｼﾞ" 'Fのリセットがないせいで歪ｺﾝﾎﾞのリセットが行われないことがあった_2026.1.12
      Call CBAdd '歪ｺﾝﾎﾞ
      
      UF1.AMin.Caption = ""
      UF1.AMax.Caption = ""
      UF1.HMin.Caption = ""
      UF1.HMax.Caption = ""
      UF1.TMin.Caption = ""
      UF1.TMax.Caption = ""
      
      UF1.flatnesText = ""
      UF1.flatnes = ""
      UF1.ConductivText.Caption = ""
      UF1.Comment.Caption = ""
      UF1.Reien.Caption = ""
      UF1.QF.Caption = ""
      UF1.LastHot.Caption = ""
      UF1.dataLoding.Caption = "" '*****履歴呼び出しがあった場合ののリセット
      
      UF1.VC_L_A.Tag = ""
      UF1.VC_L_B.Tag = ""
      
      UF1.KenNo.Caption = ""
      UF1.KenH.Caption = ""
      UF1.HizCome.Caption = ""
      UF1.HizCome.Tag = ""
      UF1.Yama.Tag = ""
      
      UF1.Additional.Caption = ""
      UF1.TaikText.BackColor = &H80000006
      UF1.LateralText.BackColor = &H80000006
      UF1.Conductiv.BackColor = &H8000000F
      UF1.IzyouFl.BackColor = &H8000000F
            
      UF1.TSS.Visible = True
      UF1.TSU.Visible = True
      UF1.TSC.Visible = True
      UF1.TSD.Visible = True
      UF1.TSE.Visible = True
      UF1.TES.Visible = True
      UF1.TEU.Visible = True
      UF1.TEC.Visible = True
      UF1.TED.Visible = True
      UF1.TEE.Visible = True
            
      UF1.Height = 545 'Heightﾃﾞﾌｫﾙﾄｻｲｽﾞ_変更時はDefaultFormPosition,VC_Clear,UF1_Clickを修正の事
      
      UF1.Width = 887.25 '2024.05.28_サイズ変更
      
      '再読み込み
      '歪用_2023.8.21_Bのtagが残ってしまっていた
      'tagデータ消去
      UF1.Loading.Tag = "Loading" 'AccessLog停止_2023.8.25
      Call Split_Z_Tag_0("S")
      Call Split_Z_Tag_0("U")
      Call Split_Z_Tag_0("C")
      Call Split_Z_Tag_0("D")
      Call Split_Z_Tag_0("E")
      Call Split_Z_Tag_0("F")
      UF1.Loading.Tag = "" 'AccessLog再開
      
      If Me.AorB.Caption = "Now_B" Then
        Me.AorB.Caption = "Now_A"
        Me.AorB.BackColor = &H8000000F
        Me.MPEX.BackColor = &H8000000F
      End If
      
      Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("作業保存" & "_" & Lname)
      Dim A_range As Range: Set A_range = ws.Range(ws.Cells(2, 15), ws.Cells(7, 16)) 'A面
      Dim B_range As Range: Set B_range = ws.Range(ws.Cells(8, 15), ws.Cells(13, 16)) 'B面
      Call LosdStrain("A", A_range, B_range) 'tagに数値落とし込み
      
      Call Load_Text("ReLoad") 'ﾌｫｰﾑ描写更新
      Call WorkerLoad
      Call VC_Clear
      Call AtCountSelect(4)
      Call InspectionSelect("表示")
      
      Call RightAngleDisplay("新規")
      Call LSelectLateral("新規")
      UF1.Load.Caption = "on"
      UF1.Lot.Locked = True 'ﾛｯﾄ登録前に新規を押した場合解除状態になる_2021.02.09
      DoEvents
      Call WorkerName_Open
      UF1.Conductiv.BackColor = &H8000000F
    Application.EnableEvents = True
    End If
    
    Call ProgressBar_count(500, 1000)
    
    Dim SizeUf As Variant:   SizeUf = InitialSetting(UF1, "リセット")
    Dim NewTemp As Variant:  NewTemp = InitialSetting(UF1)
    
    DoEvents
    Unload UFProgress
    
  End If
'-----2021.03.04_Focus対策
UF1.Worker.Enabled = True
UF1.INP.Enabled = True
UF1.Manual.Enabled = True
UF1.StartTB.Enabled = True
End Sub

Private Sub OPS1_Change()
Call TB_BackColor(OPS1, "S")
End Sub
Private Sub OPS2_Change()
Call TB_BackColor(OPS2, "S")
End Sub
Private Sub OPS3_Change()
Call TB_BackColor(OPS3, "S")
End Sub
Private Sub OPS4_Change()
Call TB_BackColor(OPS4, "S")
End Sub
Private Sub OPS5_Change()
Call TB_BackColor(OPS5, "S")
End Sub
Private Sub OPS6_Change()
Call TB_BackColor(OPS6, "S")
End Sub

Private Sub OPE1_Change()
Call TB_BackColor(OPE1, "E")
End Sub
Private Sub OPE2_Change()
Call TB_BackColor(OPE2, "E")
End Sub
Private Sub OPE3_Change()
Call TB_BackColor(OPE3, "E")
End Sub
Private Sub OPE4_Change()
Call TB_BackColor(OPE4, "E")
End Sub
Private Sub OPE5_Change()
Call TB_BackColor(OPE5, "E")
End Sub
Private Sub OPE6_Change()
Call TB_BackColor(OPE6, "E")
End Sub

' ========== 印刷のみ実行するサブルーチン ==========
Private Sub PrintOnly(Lname As String, LotNo As String)
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("印刷用" & "_" & Lname)
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname)

Call MusicTrue("印刷.wav")
If MsgBox("このLOTは印刷済みです。再印刷しますか？", vbYesNo) = vbYes Then
    Application.PrintCommunication = False
    '再印刷
    sh.Cells(2, 4) = "再印刷"
    sh.Visible = True
    sh.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
    ws.Visible = True
    ws.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
    sh.Visible = False
    
    ' 再印刷の日時を更新
    Call SetPrinted(LotNo, Lname)
    
    Call MusicTrue("完了.wav")
    MsgBox "再印刷が完了しました", vbInformation
End If

Call frameblindfold("通常")
End Sub

Private Sub PrintBreset_Click()
Dim targetColor As Long: targetColor = RGB(255, 165, 0) 'オレンジ色
Dim LotNoTmp As String:  LotNoTmp = Me.Lot.Text

If Me.PrintB.BackColor = targetColor Then
  Dim rc As VbMsgBoxResult: rc = MsgBox("このロットの印刷履歴を消しますか？" & vbCrLf & _
              "消せば印刷実行時にもう一度保存処理されます。", _
              vbQuestion + vbYesNo, "履歴削除の確認")
  
  If rc = vbYes Then
      Call ClearPrinted(LotNoTmp) '履歴削除
      Call UpdatePrintButtonColor(LotNoTmp) 'ボタン色
      
      MsgBox "履歴を削除しました。", vbInformation
  Else
      MsgBox "履歴削除処理を中断しました。", vbExclamation
  End If
End If
End Sub

Private Sub PrintB_Click()
' 2度押し対策：すでに無効化されている、またはフラグが off なら処理を抜ける
If Not UF1.PrintB.Enabled Or UF1.Load.Caption = "off" Then Exit Sub

' ボタンを無効化し、フラグを off に（.Caption を明示）
UF1.PrintB.Enabled = False
UF1.Load.Caption = "off"

UF1.PrintB.Tag = ""
ThisWorkbook.Activate

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim LotNo As String: LotNo = UF1.Lot.Text
Dim LotNoBackup As String: LotNoBackup = LotNo
'印刷済みチェック - 印刷のみ実行
If IsPrinted(LotNo) Then
    Call PrintOnly(Lname, LotNo)
    UF1.Load.Caption = "on"
    UF1.PrintB.Enabled = True
    Exit Sub
End If

Call EndTime_Register

Dim flag As Boolean: flag = AdvanceCheck(LotNo)
Dim 先行 As String: 先行 = "先行データ票はありますか？"

If flag Then
    Call 音楽を流す("ありますか.wav")
    If MsgBox(先行, vbYesNo) = vbYes Then
        GoTo SkipPrint
    Else
        UF1.Load.Caption = "on"
        UF1.PrintB.Enabled = True
        Exit Sub
    End If
End If

SkipPrint:

Dim 定盤 As String: 定盤 = Distortion_Check("Z")
Dim ゲージ As String: ゲージ = Distortion_Check("G")
Dim ラテ As String: ラテ = Lateralbow_Check
Dim 外観 As String: 外観 = Visual_Check
Dim KenH As String: KenH = H_Check
Dim 歪AB As String: 歪AB = WithorWithout_Gauge
Dim 直角 As String: 直角 = Angle_Check

' バリデーションチェック
''許容はスルー　保留は止める**2024.4.16
''印刷に関しては保留許容ともに可とする(データ保存のため)
''MinMaxでNGrange_Checkにて保留を通過させない→ｱｲｶｽｲﾝﾌﾟｯﾄに行けない
Select Case True
Case 定盤 <> ""
    Call MsgError(定盤 & "入力なし", Lname, 1)
    UF1.Load.Caption = "on"
    UF1.PrintB.Enabled = True
    Exit Sub
Case ゲージ <> ""
    Call MsgError(ゲージ & "入力なし", Lname, 1)
    UF1.Load.Caption = "on"
    UF1.PrintB.Enabled = True
    Exit Sub
Case 歪AB <> ""
    Call MsgError(歪AB, Lname, 1)
    UF1.Load.Caption = "on"
    UF1.PrintB.Enabled = True
    Exit Sub
Case ラテ <> "" And (UF1.Sumisei.Caption <> """" And InStr(UF1.Additional.Caption, "ﾗﾃﾗﾙﾎﾞｰ検査不要") = 0)
    Call MsgError(ラテ & ":ﾗﾃﾗﾙﾎﾞｰ入力なし", Lname, 1)
    UF1.Load.Caption = "on"
    UF1.PrintB.Enabled = True
    Exit Sub
Case 直角 <> "" And UF1.Youto = "R115"
    Call MsgError(直角 & ":直角度入力なし", Lname, 1)
    UF1.Load.Caption = "on"
    UF1.PrintB.Enabled = True
    Exit Sub
Case 外観 <> ""
    Call MsgError(外観 & ":外観入力なし", Lname, 1)
    UF1.Load.Caption = "on"
    UF1.PrintB.Enabled = True
    Exit Sub
Case KenH = "入力なし"
    Call MsgError("検入高さor枚数　入力なし", Lname, 1)
    UF1.Load.Caption = "on"
    UF1.PrintB.Enabled = True
    Exit Sub
Case UF1.Yama.Visible And UF1.Yama.Text = ""
    Call MsgError("山数の入力がありません", Lname, 1)
    UF1.Load.Caption = "on"
    UF1.PrintB.Enabled = True
    Exit Sub
End Select

UFProgress.Show vbModeless
Call MusicTrue("おまちください.wav")
Call frameblindfold("目隠し")

Call Name_Match
Call Number_Count(UF1.Lot)
If UF1.PrintB.Tag = "stop" Then
    UF1.Load.Caption = "on"
    UF1.PrintB.Enabled = True
    Call frameblindfold("通常")
    Exit Sub
End If

Call ProgressBar_count(0, 250)
DoEvents
Call 印刷転送
If UF1.PrintB.Tag = "stop" Then
    Unload UFProgress
    MsgBox "ｷｬﾝｾﾙ処理", vbInformation
    UF1.Load.Caption = "on"
    UF1.PrintB.Enabled = True
    Call frameblindfold("通常")
    Exit Sub
End If

Call ProgressBar_count(250, 375)
DoEvents

Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("印刷用" & "_" & Lname)
If (WsCheck("不良要因印刷" & "_" & Lname) = False) Then
    Call AddWS("不良要因印刷" & "_" & Lname)
    Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname)
ElseIf ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Cells(1, 2) <> UF1.Lot Then
    Call AddWS("不良要因印刷" & "_" & Lname)
    Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname)
End If
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname)

Call MusicTrue("印刷.wav")
If MsgBox("印刷する？", vbYesNo) = vbYes Then
    Select Case True
    Case UF1.L1Work: Call WriteLog("*UFKen:" & "印刷" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "L-1\L-1_Log")
    Case UF1.LVC: Call WriteLog("*UFKen:" & "印刷" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "LVC\LVC_Log")
    Case UF1.HVC: Call WriteLog("*UFKen:" & "印刷" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "HVC\HVC_Log")
    Case UF1.Bala: Call WriteLog("*UFKen:" & "印刷" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "バランサー\バランサー_Log")
    Case UF1.Tott: Call WriteLog("*UFKen:" & "印刷" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "トット\トット_Log")
    Case UF1.Maru: Call WriteLog("*UFKen:" & "印刷" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "丸徳\丸徳_Log")
    Case Else: Call WriteLog("*UFKen:" & "印刷" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "Test\Test_Log")
    End Select
    
    Call ProgressBar_count(375, 550)
    DoEvents
    
    '印刷実行
    Application.PrintCommunication = False
    sh.Visible = True
    sh.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
    ws.Visible = True
    ws.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
    
    UF1.dataLoding.Caption = ""
    
    Call MainDataSave("保存用")
    Call ProgressBar_count(550, 650)
    DoEvents
    Call 履歴保存
    Call ProgressBar_count(650, 880)
    DoEvents
    
    ' CSV保存処理（既存コード）
    Dim saveSh As Worksheet
    Dim FileNewName As String
    If Not UF1.極厚 Then
        Set saveSh = ThisWorkbook.Sheets("保存用")
        Call ExpCSV(saveSh, Lname & "_履歴CSV")
        FileNewName = Lname & "_File"
    Else
        Set saveSh = ThisWorkbook.Sheets("耳用")
        Call ExpCSV(saveSh, Lname & "_耳用履歴CSV")
        FileNewName = Lname & "_耳用File"
    End If
    
    Dim Nitiina As String: Nitiina = HistoryPath
    Select Case True
    Case UF1.L1Work: If IsBookOpened(Nitiina & "L-1\L-1_履歴" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(saveSh, FileNewName, Nitiina & "L-1\L-1_履歴")
    Case UF1.LVC: If IsBookOpened(Nitiina & "LVC\LVC_履歴" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(saveSh, FileNewName, Nitiina & "LVC\LVC_履歴")
    Case UF1.HVC: If IsBookOpened(Nitiina & "HVC\HVC_履歴" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(saveSh, FileNewName, Nitiina & "HVC\HVC_履歴")
    Case UF1.Bala: If IsBookOpened(Nitiina & "バランサー\バランサー_履歴" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(saveSh, FileNewName, Nitiina & "バランサー\バランサー_履歴")
    Case UF1.Tott: If IsBookOpened(Nitiina & "トット\トット_履歴" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(saveSh, FileNewName, Nitiina & "トット\トット_履歴")
    Case UF1.Maru: If IsBookOpened(Nitiina & "丸徳\丸徳_履歴" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(saveSh, FileNewName, Nitiina & "丸徳\丸徳_履歴")
    Case Else: If IsBookOpened(Nitiina & "Test\Test_履歴" & "\" & "Test_File.csv") = False Then Call ExpCSV(saveSh, FileNewName, Nitiina & "Test\Test_履歴")
    End Select
    
    Dim Kouzyou As String: Kouzyou = FactoryyPath
    Select Case True
    Case UF1.L1Work: If IsBookOpened(Kouzyou & "【梱包】L-1" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(saveSh, FileNewName, Kouzyou & "【梱包】L-1")
    Case UF1.LVC: If IsBookOpened(Kouzyou & "【梱包】LVC" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(saveSh, FileNewName, Kouzyou & "【梱包】LVC")
    Case UF1.HVC: If IsBookOpened(Kouzyou & "【梱包】HVC" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(saveSh, FileNewName, Kouzyou & "【梱包】HVC")
    Case UF1.Bala: If IsBookOpened(Kouzyou & "【梱包】大板" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(saveSh, FileNewName, Kouzyou & "【梱包】大板")
    Case UF1.Tott: If IsBookOpened(Kouzyou & "【梱包】大板" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(saveSh, FileNewName, Kouzyou & "【梱包】大板")
    Case UF1.Maru: If IsBookOpened(Kouzyou & "【梱包】中板" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(saveSh, FileNewName, Kouzyou & "【梱包】中板")
    End Select
    
    sh.Visible = False
    
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    
    DoEvents
    Call ProgressBar_count(880, 1000)
    DoEvents
    Unload UFProgress
    
    If LotNo = "" Then LotNo = LotNoBackup '復元
    
    '印刷済みフラグを設定
    Call SetPrinted(LotNo, Lname)
    'ボタンの変色
    Call CheckAndUpdateButtonColor
    
    UF1.Load.Caption = "on"
    UF1.PrintB.Enabled = True
    Call MusicTrue("完了.wav")
    Call frameblindfold("通常")
Else
    Unload UFProgress
    MsgBox "ｷｬﾝｾﾙ処理", vbInformation
    UF1.Load.Caption = "on"
    UF1.PrintB.Enabled = True
    Call frameblindfold("通常")
    Exit Sub
End If
Application.Visible = False
End Sub

'Private Sub PrintB_Click()
'If UF1.Load.Caption = "off" Then Exit Sub
'UF1.PrintB.Tag = ""
'ThisWorkbook.Activate
'UF1.Load = "off" 'フォーム関係イベントを止める　Exit Sub
'Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
'Call EndTime_Register '終了時間をﾌｫｰﾑに書き込む
'
'Dim LotNo As String: LotNo = UF1.LOT.Text
'Dim flag As Boolean: flag = AdvanceCheck(LotNo)
'Dim 先行 As String:  先行 = "先行データ票はありますか？"
''2023.12.22***追加
'If flag Then
'  Call 音楽を流す("ありますか.wav")
'  If MsgBox(先行, vbYesNo) = vbYes Then
'    GoTo SkipPrint
'  Else
'    UF1.Load.Caption = "on"
'    Exit Sub
'  End If
'End If
'
'SkipPrint:
'
'Dim 定盤 As String:    定盤 = Distortion_Check("Z")
'Dim ゲージ As String:  ゲージ = Distortion_Check("G")
'Dim ラテ As String:    ラテ = Lateralbow_Check '2021.07.14
'Dim 外観 As String:    外観 = Visual_Check '2023.5.22
''Dim bh As String:      bh = BH_Check(Left(UF1.Lot, 1)) '2023.6.9
'Dim KenH As String:    KenH = H_Check '2023.6.9
'Dim 歪AB As String:    歪AB = WithorWithout_Gauge '2023.11.20
'Dim 直角 As String:    直角 = Angle_Check '2023.11.30
'
'
''許容はスルー　保留は止める**2024.4.16
''印刷に関しては保留許容ともに可とする(データ保存のため)
''MinMaxでNGrange_Checkにて保留を通過させない→ｱｲｶｽｲﾝﾌﾟｯﾄに行けない
'
'Select Case True
''Case NGrange_Check(1, 5)
''  Call MsgError("許容でない公差外が存在します", Lname, 1)
''  UF1.Load.Caption = "on"
''  Exit Sub
'Case Not CBdateCheck
'  Call MsgError("歪測定の入力が一つもありません", Lname, 1)
'  UF1.Load.Caption = "on"
'  Exit Sub
'Case 定盤 <> ""
'  Call MsgError(定盤 & "入力なし", Lname, 1)
'  UF1.Load.Caption = "on"
'  Exit Sub
'Case ゲージ <> ""
'  Call MsgError(ゲージ & "入力なし", Lname, 1)
'  UF1.Load.Caption = "on"
'  Exit Sub
'Case 歪AB <> ""
'  Call MsgError(歪AB, Lname, 1)
'  UF1.Load.Caption = "on"
'  Exit Sub
'Case ラテ <> "" And UF1.Sumisei.Caption <> "" '2021.07.14**2022.8.17*****ﾗﾃﾗﾙﾎﾞｰ測定の納入先が追加される為"住友精密"から<>に変更
'  Call MsgError(ラテ & ":ﾗﾃﾗﾙﾎﾞｰ入力なし", Lname, 1)
'  UF1.Load.Caption = "on"
'  Exit Sub
'Case 直角 <> "" And UF1.Youto = "R115" '2023.11.30
'  Call MsgError(直角 & ":直角度入力なし", Lname, 1)
'  UF1.Load.Caption = "on"
'  Exit Sub
'Case 外観 <> ""
'  Call MsgError(外観 & ":外観入力なし", Lname, 1)
'  UF1.Load.Caption = "on"
'  Exit Sub
''Case bh <> ""
''  Call MsgError(bh & ":BH入力なし", Lname, 1)
''  UF1.Load.Caption = "on"
''  Exit Sub
'Case KenH = "入力なし"
'  Call MsgError("検入高さor枚数　入力なし", Lname, 1)
'  UF1.Load.Caption = "on"
'  Exit Sub
'Case UF1.Yama.Visible And UF1.Yama.Text = "" '2022.01.11**山の表示があるときは冷延コース有6mm以下　か冷延なしで8mm未満
'  Call MsgError("山数の入力がありません", Lname, 1)
'  UF1.Load = "on"
'  Exit Sub
'End Select
'
'UFProgress.Show vbModeless
''Application.Speech.Speak Text:="少々お待ちください"
'Call MusicTrue("おまちください.wav")
'Call frameblindfold("目隠し")
'
'Call Name_Match '登録検査員と実績検査員の一致を確認　不一致の時は停止処理起動
'Call Number_Count(UF1.LOT) '梱包枚数ﾁｪｯｸ
'If UF1.PrintB.Tag = "stop" Then
'   UF1.Load.Caption = "on"
'   Call frameblindfold("通常")
'   Exit Sub
'End If
'Call ProgressBar_count(0, 250)
'DoEvents
'Call 印刷転送
'If UF1.PrintB.Tag = "stop" Then
'    Unload UFProgress
'    MsgBox "ｷｬﾝｾﾙ処理", vbInformation
'    UF1.Load.Caption = "on"
'    Call frameblindfold("通常")
'    Exit Sub
'End If
'
'Call ProgressBar_count(250, 375)
'DoEvents
'Dim SH As Worksheet: Set SH = ThisWorkbook.Sheets("印刷用" & "_" & Lname)
'If (WsCheck("不良要因印刷" & "_" & Lname) = False) Then 'シートなし
'  Call AddWS("不良要因印刷" & "_" & Lname)
'  Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname) '(原紙名、付ける名前)
'ElseIf ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname).Cells(1, 2) <> UF1.LOT Then '不良の入力がなければこの状態
'  Call AddWS("不良要因印刷" & "_" & Lname)
'  Call Pesute("不良要因印刷原紙", "不良要因印刷" & "_" & Lname)
'End If
'Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("不良要因印刷" & "_" & Lname)
'
''Application.Speech.Speak Text:="印刷しますか？"
'Call MusicTrue("印刷.wav")
'If MsgBox("印刷する？", vbYesNo) = vbYes Then
'  Select Case True
'  Case UF1.L1Work:   Call WriteLog("*UFKen:" & "印刷" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "L-1\L-1_Log")
'  Case UF1.LVC:      Call WriteLog("*UFKen:" & "印刷" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "LVC\LVC_Log")
'  Case UF1.HVC:      Call WriteLog("*UFKen:" & "印刷" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "HVC\HVC_Log")
'  Case UF1.Bala:     Call WriteLog("*UFKen:" & "印刷" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "バランサー\バランサー_Log")
'  Case UF1.Tott:     Call WriteLog("*UFKen:" & "印刷" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "トット\トット_Log")
'  Case UF1.Maru:     Call WriteLog("*UFKen:" & "印刷" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "丸徳\丸徳_Log")
'  Case Else:         Call WriteLog("*UFKen:" & "印刷" & " T1:" & UF1.T1 & " T2:" & UF1.T2 & " T3:" & UF1.T3, "", HistoryPath & "Test\Test_Log")
'  End Select
'  Call ProgressBar_count(375, 550)
'  DoEvents
'  Application.PrintCommunication = False
'  SH.Visible = True
'  SH.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
'
'  ws.Visible = True
'  ws.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
'
'  UF1.dataLoding.Caption = "" 'リセット*****履歴呼び出しがあった場合のリセット
'
'  Call MainDataSave("保存用") '印刷時にロット完了と判断
'  Call ProgressBar_count(550, 650)
'  DoEvents
'  Call 履歴保存 '2016.06.26
'  Call ProgressBar_count(650, 880)
'  DoEvents
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>2022.06***保存ファイル名修正
'  Dim SaveSh As Worksheet
'  Dim FileNewName As String
'  If Not UF1.極厚 Then '耳用と分けて保存が必要_2022.06.10
'    Set SaveSh = ThisWorkbook.Sheets("保存用")
'    Call ExpCSV(SaveSh, Lname & "_履歴CSV")
'    FileNewName = Lname & "_File" 'サーバー上ファイル名定義
'  Else
'    Set SaveSh = ThisWorkbook.Sheets("耳用")
'    Call ExpCSV(SaveSh, Lname & "_耳用履歴CSV")
'    FileNewName = Lname & "_耳用File"
'  End If
'
'  Dim Nitiina As String: Nitiina = HistoryPath
'  Select Case True '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\検査履歴\L-1 履歴"
'  Case UF1.L1Work: If IsBookOpened(Nitiina & "L-1\L-1_履歴" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(SaveSh, FileNewName, Nitiina & "L-1\L-1_履歴")
'  Case UF1.LVC:    If IsBookOpened(Nitiina & "LVC\LVC_履歴" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(SaveSh, FileNewName, Nitiina & "LVC\LVC_履歴")
'  Case UF1.HVC:    If IsBookOpened(Nitiina & "HVC\HVC_履歴" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(SaveSh, FileNewName, Nitiina & "HVC\HVC_履歴")
'  Case UF1.Bala:   If IsBookOpened(Nitiina & "バランサー\バランサー_履歴" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(SaveSh, FileNewName, Nitiina & "バランサー\バランサー_履歴")
'  Case UF1.Tott:   If IsBookOpened(Nitiina & "トット\トット_履歴" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(SaveSh, FileNewName, Nitiina & "トット\トット_履歴")
'  Case UF1.Maru:   If IsBookOpened(Nitiina & "丸徳\丸徳_履歴" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(SaveSh, FileNewName, Nitiina & "丸徳\丸徳_履歴")
'  Case Else:       If IsBookOpened(Nitiina & "Test\Test_履歴" & "\" & "Test_File.csv") = False Then Call ExpCSV(SaveSh, FileNewName, Nitiina & "Test\Test_履歴")
'  End Select
'  '2021.01.22_工場共有
'  Dim Kouzyou As String: Kouzyou = FactoryyPath
'  Select Case True '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\検査履歴\L-1 履歴"
'  Case UF1.L1Work: If IsBookOpened(Kouzyou & "【梱包】L-1" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(SaveSh, FileNewName, Kouzyou & "【梱包】L-1")
'  Case UF1.LVC:    If IsBookOpened(Kouzyou & "【梱包】LVC" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(SaveSh, FileNewName, Kouzyou & "【梱包】LVC")
'  Case UF1.HVC:    If IsBookOpened(Kouzyou & "【梱包】HVC" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(SaveSh, FileNewName, Kouzyou & "【梱包】HVC")
'  Case UF1.Bala:   If IsBookOpened(Kouzyou & "【梱包】大板" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(SaveSh, FileNewName, Kouzyou & "【梱包】大板")
'  Case UF1.Tott:   If IsBookOpened(Kouzyou & "【梱包】大板" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(SaveSh, FileNewName, Kouzyou & "【梱包】大板")
'  Case UF1.Maru:   If IsBookOpened(Kouzyou & "【梱包】中板" & "\" & Lname & "_File.csv") = False Then Call ExpCSV(SaveSh, FileNewName, Kouzyou & "【梱包】中板")
'  End Select
'  SH.Visible = False
'
'  Application.DisplayAlerts = False
'  ThisWorkbook.Save '上書き保存
'  Application.DisplayAlerts = True
'
'  DoEvents
'  Call ProgressBar_count(880, 1000)
'  DoEvents
'  Unload UFProgress
'  UF1.Load.Caption = "on"
'  Call MusicTrue("完了.wav")
'
'  Call frameblindfold("通常")
''  Application.Speech.Speak Text:="完了しました"
'Else
'Unload UFProgress
'MsgBox "ｷｬﾝｾﾙ処理", vbInformation
'  UF1.Load.Caption = "on"
'  Call frameblindfold("通常")
'  Exit Sub
'End If
'Application.Visible = False 'Excel非表示
'End Sub

Private Sub ReSize_Click()
On Error Resume Next
'-----2021.03.04_Focus対策
UF1.Worker.Enabled = True
UF1.INP.Enabled = True
UF1.Manual.Enabled = True
UF1.StartTB.Enabled = True
If MsgBox("Size、公差の編集を行いますか？", vbQuestion + vbYesNo) = vbYes Then
  With UF1
    .ReSize.Tag = "有"
    .ResizText.Caption = .ReSize.Tag
    .ATU.Locked = False '厚　製造サイズとする
    .Hab.Locked = False '巾
    .Tak.Locked = False '丈
    .Ap.Locked = False '+
    .Am.Locked = False '-
    .Hp.Locked = False '+
    .Hm.Locked = False '-
    .Tp.Locked = False '+
    .Tm.Locked = False '-
    '****追加_2021.0011.02******対角ラテラルボ-修正_2022.07.01
    .TAIK.Locked = False '対角
    .Lateral.Locked = False 'ﾗﾃﾗﾙﾎﾞｰ
    .TaikText.Locked = False '対角
    .LateralText.Locked = False 'ﾗﾃﾗﾙﾎﾞｰ
    .TAIK.BackColor = &HC0FFFF '対角
    .Lateral.BackColor = &HC0FFFF 'ﾗﾃﾗﾙﾎﾞｰ
    '***********
    .ATU.BackColor = &HC0FFFF
    .Hab.BackColor = &HC0FFFF '巾
    .Tak.BackColor = &HC0FFFF '丈
    .Ap.BackColor = &HC0FFFF '+
    .Am.BackColor = &HC0FFFF '-
    .Hp.BackColor = &HC0FFFF '+
    .Hm.BackColor = &HC0FFFF '-
    .Tp.BackColor = &HC0FFFF '+
    .Tm.BackColor = &HC0FFFF '-
  End With
End If
Call Posting '履歴に残す
End Sub

Private Sub T1_Change()
If UF1.Load.Caption = "off" Then Exit Sub 'ﾌｫｰﾑｲﾍﾞﾝﾄ停止
If UF1.T1 = "" Then Exit Sub
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim Worker1: Worker1 = Split(UF1.T1, "_") '複数人の登録があった場合後の登録者を記録する
Dim Wor1 As String: Wor1 = Worker1(UBound(Worker1))
Call WorkerSave
Call WorkerName_Open
On Error Resume Next
'If Not WorkeDayMax = "1" Then Exit Sub '2021.01.18_夜勤0時がMaxとならない為停止
If Mid((UF1.Day1), 1, 10) = YeaDayrCheck(Format$(Now, "hh:mm")) Then '当日の登録がある
  If 校正(Format$(Now, "hh:mm")) Then '当日の校正記録がない
    UF1.Load.Caption = "off"
    Application.WindowState = xlMinimized
    DoEvents
    Call 校正_Open(Wor1, Lname) '校正実行
    'SendKeys "{TAB}" '.INPG.SetFocusにしてある,SendKeys "{TAB}"でINP1000へ移動
    'Call Numlock_onoff
    Call 音楽検査員登録 '2021.04.14
  End If
End If
End Sub
Private Sub T2_Change()
If UF1.Load.Caption = "off" Then Exit Sub
If UF1.T2 = "" Then Exit Sub
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim Worker2: Worker2 = Split(UF1.T2, "_")
Dim Wor2 As String: Wor2 = Worker2(UBound(Worker2))
Call WorkerSave
Call WorkerName_Open
'If Not WorkeDayMax = "2" Then Exit Sub '2021.01.18_夜勤0時がMaxとならない為停止
On Error Resume Next
If Mid((UF1.Day2), 1, 10) = YeaDayrCheck(Format$(Now, "hh:mm")) Then '当日
  If 校正(Format$(Now, "hh:mm")) Then
    UF1.Load.Caption = "off"
    Application.WindowState = xlMinimized
    DoEvents
    Call 校正_Open(Wor2, Lname)
     Call 音楽検査員登録 '2021.04.14
  End If
End If
End Sub
Private Sub T3_Change()
If UF1.Load.Caption = "off" Then Exit Sub
If UF1.T3 = "" Then Exit Sub
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Dim Worker3: Worker3 = Split(UF1.T3, "_")
Dim Wor3 As String: Wor3 = Worker3(UBound(Worker3))
Call WorkerSave
Call WorkerName_Open
'If Not WorkeDayMax = "3" Then Exit Sub '2021.01.18_夜勤0時がMaxとならない為停止
On Error Resume Next
If Mid((UF1.Day3), 1, 10) = YeaDayrCheck(Format$(Now, "hh:mm")) Then '当日
  If 校正(Format$(Now, "hh:mm")) Then
    UF1.Load.Caption = "off"
    Application.WindowState = xlMinimized
    DoEvents
    Call 校正_Open(Wor3, Lname)
     Call 音楽検査員登録 '2021.04.14
  End If
End If
End Sub
Private Sub Tak_Change()
If IsNumeric(UF1.Tak) = False Then UF1.Tak = ""
End Sub
Private Sub Tm_Change()
If IsNumeric(UF1.Tm) = False Then UF1.Tm = ""
Call Form_Formula
End Sub

Private Sub Tp_Change()
If IsNumeric(UF1.Tp) = False Then UF1.Tp = ""
Call Form_Formula
End Sub
Private Sub Taik_Change()
With UF1
 .TaikText = ""
  If .Hab <> "" And .Tak <> "" Then
      Select Case .TAIK
      Case "4"
       .DebugTaikaku = Format(Diagonal(.Hab, .Tak), "0.0")  '4級(厳しい)
      Case "3"
       .DebugTaikaku = Format(ReDiagonal(.Hab, .Tak), "0.0")
      End Select
  End If
End With
End Sub
Private Sub Lateral_Change()
'等級で見るときは下1桁***2022.09.20
With UF1
.LateralText = ""
  If .Hab <> "" And .Tak <> "" Then
    Select Case .Lateral
    Case "4"
     .DebugLate = Format(Lateralbow_J(.Hab, .Tak, .ATU), "0")  '4級(厳しい)
     If .DebugLate = 0 Then .DebugLate = ""
    Case "3"
     .DebugLate = Format(ReLateralbow(.Hab, .Tak, .ATU), "0.0")
     If .DebugLate = 0 Then .DebugLate = ""
    End Select
  End If
End With
End Sub
Private Sub UserForm_Resize()
'Dim iHeight
'Dim iWidth
'On Error Resume Next
''// Heightの+36は目視調整した。
'If Application.WindowState = xlNormal Then ' Excelウィンドウ位置を同期させる
'  iWidth = UF1.InsideWidth - Application.Left * 2
'  iHeight = UF1.InsideHeight - Application.Top * 2 + 36
'End If
''// WidthとHeightには0以下は設定不可のためエラーになるので０チェック
'If (iWidth > 0 And iHeight > 0) Then
'  Application.Width = iWidth
'  Application.Height = iHeight
'End If

End Sub
Private Sub UserForm_Activate()
If UFCask.Visible = True Then Exit Sub 'ｷｬｽｸ中はUF1を干渉させない_2021.07.09
If WsCheck("全ヵ所検査") And WsCheck("キャスク検査") Then
  If ThisWorkbook.Sheets("全ヵ所検査").Visible Or ThisWorkbook.Sheets("キャスク検査").Visible Then Exit Sub
End If
If UF1.Load.Caption = "off" Then Exit Sub

Call WindowMenuButtonOn(Me) '2021.03.25
'Call IMEOFF

If Register_Check(Format$(Now, "hh:mm:ss")) Then UF1.WorkButton.BackColor = &H8080FF  '検査員登録なし
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("WorkerTime" & "_" & Lname) Then '検査員一致判定外れﾌﾗｸﾞ
  If ThisWorkbook.Sheets("WorkerTime" & "_" & Lname).Cells(1, 18) = "OUT" Then
    MsgBox "Excelを閉じても解除はできません" & vbCrLf & "隠蔽工作Log出力", vbExclamation
    Dim path As String: path = HistoryPath
    Dim Temp As String: Temp = "*UFKen:UFNameList起動隠蔽対策:Excel再起動の形跡"
    Select Case True
    Case UF1.L1Work:   Call WriteLog(Temp, "", path & "L-1\L-1_Log")
    Case UF1.LVC:      Call WriteLog(Temp, "", path & "LVC\LVC_Log")
    Case UF1.HVC:      Call WriteLog(Temp, "", path & "HVC\HVC_Log")
    Case UF1.Bala:     Call WriteLog(Temp, "", path & "バランサー\バランサー_Log")
    Case UF1.Tott:     Call WriteLog(Temp, "", path & "トット\トット_Log")
    Case UF1.Test:     Call WriteLog(Temp, "", path & "Test\Test_Log")
    Case Else:         Call WriteLog(Temp, "", path & "丸徳\丸徳_Log")
    End Select
    Call NameList_Open(1) 'ﾌｫｰﾑ起動
    Exit Sub
  End If
End If

'Application.Visible = False 'Excel非表示
On Error Resume Next 'アクティブが移りどこにフォーカスが移るか不明のため苦肉対策**2022.06.23
Application.ScreenUpdating = False
Call VisibleSheets 'シート表示
Call HiddenSheets '隠す
Application.ScreenUpdating = True
'With Application ' フォームの後ろにExcelウィンドウを隠す
'  .WindowState = xlNormal
'  .Top = Me.Top - 2
'  .Left = Me.Left - 2
'  .Width = Me.Width - 10
'  .Height = Me.Height - 10
'End With
If UF1.AtButt.Visible = True Then
  If UF1.AtCount = "厚2ヵ所検査中" Then
    AtButt.Caption = "前回設定2ヵ所"    '点滅させる内容をセット
  Else
    AtButt.Caption = "前回設定4ヵ所"    '点滅させる内容をセット
  End If
  
'  BlinkLabel1.Blink 2      '点滅回数登録
  
  If UF1.AtCount = "厚2ヵ所検査中" Then
    AtButt.Caption = "厚4ヵ所"
  Else
    AtButt.Caption = "厚2ヵ所"
  End If
End If

'Dim n As Long, a As Long
'If UF1.薄板 Then
'  Call Look_S_Blank(1, 10, KENNO)
'ElseIf UF1.極厚 Then
'  Call Look_Blank(1, 14)
'Else
'  Call Look_Blank(1, 10)
'End If
'If UF1.WorkButton.BackColor = &H8000000F Then Call WorkerName_Open '検査員表示_2021.02.15

On Error Resume Next
UF1.Lot.Locked = True
UF1.Lot.Enabled = False

Me.Manual.SetFocus 'ﾌｫｰﾑｸﾘｯｸでｱｸﾃｨﾌﾞになるのでLot登録時に処理が被る
Me.StartTB.SetFocus
Call FormSetting ' フォームをリサイズ可能にするための設定

If 校正(Format$(Now, "hh:mm")) = False Then UF1.ProofreadFlag = "ok" '校正済の時_2021.07.15追加
UF1.Auto = "on"

End Sub
Private Sub UserForm_Click()
UF1.MPEX.Tag = ""
UF1.MoveFlag = ""
UF1.MPEX.Visible = True

UF1.MPEX.Height = NowSizeH '2024.5.30_追加
UF1.MPEX.Width = NowSizeW
UF1.MP1.Width = NowSizeMP

UF1.Repaint

'2023.6.27_適当にtagに数字追加　カウントをとる**UF1.PicFM.Tag
Dim Lcount As Long: Lcount = 1
UF1.PicFM.Tag = val(UF1.PicFM.Tag) + Lcount

'2023.12.22_先行用追加
Dim LotNo As String: LotNo = UF1.Lot.Text
Dim FL As Boolean:   FL = AdvanceCheck(LotNo)
Dim 先行 As String:  先行 = "先行データ票が必要です。"
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption

If UF1.Load.Caption = "off" Then Exit Sub

'MsgBox UF1.PicFM.Tag
If FL And UF1.PicFM.Tag = 1 Then
  Call MsgError(先行, Lname, 4)
  Call 音楽を流す("必要です.wav")
End If

Call LoopJudg(val(UF1.PicFM.Tag))
Call HizLoopJudg(val(UF1.PicFM.Tag))

'2025.1.23
Call CheckLabelsByCategory("S")
Call CheckLabelsByCategory("U")
Call CheckLabelsByCategory("C")
Call CheckLabelsByCategory("D")
Call CheckLabelsByCategory("E")
Call CheckLabelsByCategory("F")

Call DefaultFormPosition 'UF1の位置

Call AllClose("日報＆資材計算.xlsm") '指定したブックとアクティブブック以外を閉じる

If Register_Check(Format$(Now, "hh:mm:ss")) Then UF1.WorkButton.BackColor = &H8080FF  '検査員登録なし
' Excelウィンドウ位置を同期させる
'With Application
'  If .WindowState = xlNormal Then
'    .Top = Me.Top + 1
'    .Left = Me.Left + 1
'  End If
'End With
UF1.Worker.BackColor = &H80000005
If Me.Lot <> "" Then
   BlinkLabel1.BlinkStop '点滅ｽﾄｯﾌﾟ
End If


Call Fm_Clear("Click") '歪ﾌﾚｰﾑﾘｾｯﾄ_2021.03.29
UF1.Auto = "on"

Dim i As Long
Dim KenNo As String: KenNo = Left(UF1.KenNo.Caption, 1) '検査No'頭1桁

Call LookingYellow '黄色リセット

Dim n As Long, a As Long
If UF1.薄板 Then
  Call Look_S_Blank(1, 10, KenNo)
ElseIf UF1.極厚 Then
  Call Look_Blank(1, 14)
ElseIf UF1.少枚数 And UF1.KenCount <> "" Then
  Call Smallnumber_Blank(1, 10, val(UF1.KenCount))
Else
  Call Look_Blank(1, 10)
End If

If Register_Check(Format$(Now, "hh:mm:ss")) Then '検査員登録なし
  UF1.WorkButton.BackColor = &H8080FF 'ない場合赤
  If UFWorkerName.Visible Then Unload UFWorkerName
Else
  UF1.WorkButton.BackColor = &H8000000F
  Call WorkerName_Open '検査員表示_2021.02.15
End If
'山数入力を警告_2021.12.21
If UF1.薄板 And UF1.Yama.Visible And UF1.Yama.Text = "" Then
  MsgBox "山数の入力がありません" & vbCrLf & "選択してください", vbExclamation
  Exit Sub
End If

Dim Kno As String:   Kno = UF1.厚板.Caption
Dim flag As Boolean: If UF1.少枚数 Then flag = True

'少枚数用の枚数を確定させる_2023.8.24
Dim th As Long
If UF1.ATU <> "" Then
  If UF1.KenCount = "" Then
    th = Format(Round(CalculateSheets(UF1, UF1.ATU), 1), "0")
  ElseIf UF1.KenCount <> "" Then
    th = UF1.KenCount
  End If
  If th >= 5 Then th = 5 '少枚数は5枚まで
End If

If Kno = "H" And flag Then
 If UF1.PicFM.Tag = 1 Then Call SmallNumberCount(th)  '少枚数設定後に検査回数で表示を変える 検番Hのときのみ
End If

On Error Resume Next

Call ResetBackColor("黄色リセットなし")

UF1.FlagLabel = 1
UF1.Manual.SetFocus
UF1.StartTB.SetFocus
'Call Pos_Check
End Sub
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UF1.Load.Caption = "on" 'イベント再開
End Sub

Private Sub UserForm_Layout()
' Excelウィンドウ位置を同期させる
On Error Resume Next
Application.Visible = False 'Excel非表示
'With Application
'  If .WindowState = xlNormal Then
'    .Top = Me.Top + 1
'    .Left = Me.Left + 5
'  End If
'End With
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
On Error Resume Next

UF1.Tag = "STOP" '判定音停止

BlinkLabel1.BlinkStop '点滅ｽﾄｯﾌﾟ

Call Posting
If UF管理.Visible = True Then Unload UF管理
If UFWorkerName.Visible = True Then Unload UFWorkerName

Call HiddenSheets("xlVeryHidden") '_2021.03.17

'2025.5.14
' 保管情報を元にExcelウィンドウを元に戻す
Application.Visible = True
Application.ScreenUpdating = False
Application.WindowState = xlMaximized '最大化
Application.ScreenUpdating = True

Application.DisplayAlerts = False
ThisWorkbook.Save '上書き保存
Application.DisplayAlerts = True
End Sub
Private Sub UserForm_Terminate()
Set BlinkLabel1 = Nothing
End Sub
Private Sub CommandButton4_Click()
UF1.Load.Caption = "off" 'フォーム関係イベントを止める　Exit Sub

'MinMaxはNGがあった場合は実行しないので再判定
Call LoopJudg(val(UF1.PicFM.Tag))
Call HizLoopJudg(val(UF1.PicFM.Tag))

Call MinMax '"許容"ﾌﾗｸﾞなしのNGがあるとき停止
UF1.Load.Caption = "on"
End Sub
Private Sub AtButt_Click()
On Error Resume Next
'-----2021.03.04_Focus対策
UF1.INP.Enabled = True
UF1.Manual.Enabled = True
UF1.StartTB.Enabled = True
UF1.StartTB.SetFocus
Select Case UF1.AtButt.Caption
Case "厚2ヵ所": Call AtCountSelect(2)
Case "厚4ヵ所": Call AtCountSelect(4)
End Select
End Sub
Private Sub InspeButt_Click()
On Error Resume Next
'-----2021.11.29_Focus対策
UF1.INP.Enabled = True
UF1.Manual.Enabled = True
UF1.StartTB.Enabled = True
UF1.StartTB.SetFocus
Select Case UF1.InspeButt.Caption
Case "検入≧26":  Call InspectionSelect("表示")
Case "検入≦25":  Call InspectionSelect("非表示")
End Select
End Sub

Private Sub CB_Change()
If UF1.Load.Caption = "off" Then Exit Sub 'ﾌｫｰﾑｲﾍﾞﾝﾄ停止
UF1.Loading.Visible = True
UF1.Repaint

'2024.11.6
Dim HistoryLot As Worksheet: Set HistoryLot = ThisWorkbook.Sheets("保存用")
Dim LotNo As String:         LotNo = Me.Lot
Dim flag As Boolean:         flag = CheckForLot(HistoryLot, LotNo, 2)

If flag Then
  If MsgBox("直近5Lot内に同じLotNoがあります" & vbCrLf & "このまま作業しますか？", vbYesNo + vbQuestion, "確認") = vbNo Then
  Exit Sub
  End If
End If

'引き当て取り込み
Call AccReflec

'ボタン変色
Call CheckAndUpdateButtonColor

If Me.CB.Tag = "end" Then Exit Sub

Call OrdNoOpen(Me.Lot) 'ｵｰﾀﾞｰNOをコンボボックスに入れる
Dim KonSiyouN As String:            KonSiyouN = PackingNoCheck(UF1.Lot)
If KonSiyouN <> "" Then Call UFPack_Open(KonSiyouN, Me.Lot)  '包装仕様書Noチェック

Sleep 5
Call frameblindfold("目隠し")

Dim QpArry As Variant:         If UF1.Lot <> "" Then QpArry = DistArray(UF1.Lot, UF1.KenNo.Caption, UF1.DBflag, "")  'ゲージ、幅歪、丈歪、中、耳、コメント、参照元、番号、フラグ
'Dim QpArryG As Variant:        If UF1.Lot <> "" Then QpArryG = DistArray(UF1.Lot, UF1.KenNo.Caption, "取引", UF1.DBflag, "")
'Dim UpdatedQpArryD As Variant: If UF1.Lot <> "" Then UpdatedQpArryD = UpdateQpArryD(QpArry, QpArryG)

Call frameblindfold("通常")

Call OpForcedSelection 'サイズをコンパクトにし作業選択を強調する_2024.5.10

If IsEmptyArray(QpArry) = False Then Call FormReflectDist(QpArry, "") 'フォーム反映
UF1.Loading.Visible = False
UF1.Repaint
End Sub
Private Sub CommandButton2_Click()
'Dim i As Long
'For i = 1 To 10
' UF1(S & i) = i
'Next i
UF1.Auto = "on"
Call LoopJudg
Call HizLoopJudg

If Not UF1.極厚 Then Call Look_Blank(1, 10)
If UF1.極厚 Then Call Look_Blank(1, 14)

UF1.Manual.SetFocus
UF1.StartTB.SetFocus
End Sub
Private Sub CommandButton3_Click()
If Len(UF1.INP) >= 4 Then
  UF1.INP = "#〆" & UF1.Manual
ElseIf Len(UF1.INP) <= 3 Then
  UF1.INP = "#〆" & Format((UF1.Manual), "0.000")
End If
End Sub
Private Sub Frame1_Click()
UF1.StartTB.SetFocus
End Sub
Private Sub Frame7_Click()
UF1.StartTB.SetFocus
End Sub
Private Sub INP_Change() 'ﾌﾗｸﾞが1でないときﾙｰﾌﾟを抜ける
If Me.FlagLabel <> 1 Then Exit Sub
If UF1.INP = "" Then Exit Sub
'2024.6.5_追加
If UF1.Yama.Visible And UF1.Yama.Text = "" Then
  MsgBox "山数の入力がありません" & vbCrLf & "選択してください", vbExclamation
  Exit Sub
End If
'2026.07.27_追加
If UF1.Yama.Tag = "別フォーム" Then
  MsgBox "別フォームで検査してください", vbCritical
  Exit Sub
End If

Me.FlagLabel = 0
Call Fm_Clear("All") '歪ﾌﾚｰﾑﾘｾｯﾄ-------------------------歪測定フレームに判定画像が接触すると描写が止まる_2020.12.17
UF1.PakeD = Maker(UF1.INP) 'ﾏﾙﾁﾍﾟｰｼﾞTest1

'-----------------------------------------------------------------------------------------------------------------------------------
Dim i As Long
Dim KenNo As String: KenNo = Left(UF1.Reien.Caption, 1)

'2023.5.10_引数追加
Dim YouC As String
If Me.Youto <> "" Then YouC = Me.Youto
If UF1.Auto = "on" Then Call PucketMove(UF1.PakeD, KenNo, YouC) 'ﾏｲｸﾛｺﾝﾍﾞｯｸｽ判定

'-----------------------------------------------------------------------------------------------------------------------------------
If UF1.INP.Tag = "" Then Exit Sub
Dim st As String: st = UF1.INP.Tag
Select Case Maker(UF1.INP) 'コンバートの有無
Case "DT", "#L", "#〆" 'DT10002+0000021.712M'＃L395.8
  UF1(st) = val(Conv(UF1.INP))
End Select

Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
Select Case True
Case Register_Check(Format$(Now, "hh:mm:ss"), "")  '検査員ﾁｪｯｸ *1.ｼｰﾄがない2.3行(Day1,Day2,Day3)に当日以外の登録がある3.その時間に登録が無い
  Call MsgError("検査員登録をして下さい", Lname, 2)
  UF1(st) = ""
  Me.INP = ""
  Me.FlagLabel = 1
  Exit Sub
Case UF1.厚板 = False And UF1.極厚 = False And UF1.薄板 = False And UF1.少枚数 = False And UF1.スラブ = False
  Call MsgError("作業選択してください", Lname, 1)
  UF1(st) = ""
  Me.INP = ""
  Me.FlagLabel = 1
  Exit Sub
Case UF1.ProofreadFlag <> "ok" '校正(Format$(Now, "hh:mm")) 'trueで未校正_2021.07.15
  Select Case True
  Case UF1.WorkButton.Tag = "jump"
  Case UF1.DBflag <> ""
  Case Else
    Call Proofread_Load
    UF1(st) = ""
    Me.INP = ""
    Me.FlagLabel = 1
    Exit Sub
  End Select
End Select

Application.ScreenUpdating = False
Call LoopText(st)
Call Posting
Call Out_Tolerance '公差外ﾁｪｯｸ
Application.ScreenUpdating = True

UF1.Auto = "on"
UF1.Load = "on"

Dim n As Long, a As Long
If UF1.薄板 Then
  Call Look_S_Blank(1, 10, KenNo)
ElseIf UF1.極厚 Then
  Call Look_Blank(1, 14)
ElseIf UF1.少枚数 And UF1.KenCount <> "" Then
  Call Smallnumber_Blank(1, 10, val(UF1.KenCount))
Else
  Call Look_Blank(1, 10)
End If

Call Fm_Clear("Click") '歪ﾌﾚｰﾑﾘｾｯﾄ

UF1.INP = ""
UF1.FlagLabel = 1 'フラグをもどす
End Sub
Private Sub Lot_Exit(ByVal Cancel As MSForms.ReturnBoolean) 'フレーム内にBoxを入れると発生しません
UF1.JudgBox2.BackColor = RGB(0, 0, 0)
'UF1.Lot.Locked = True
'UF1.Lot.Enabled = False
End Sub
Private Sub Lot_Enter()
UF1.JudgBox2.BackColor = RGB(0, 255, 0)
End Sub
Private Sub Lot_Change()
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If Len(Me.Lot.Text) <> 7 Then
 Exit Sub
Else
  Me.CB = Me.Lot

  
  '***イベント停止フラグ作りすぎてしまった...校正、MinmaxHistoryLoadは.Load.Captionで他はTag****
'  If UF1.Load.Caption = "off" Or UF1.Loading.Tag = "HistoryLoad" Then Exit Sub 'いらないかな...一旦停止_2024.4.23
  
  If UF1.Tag <> "STOP" Then  'Load中ではなく登録時のみ_2021.01.29
  
    Call WriteLog("Lot登録|" & Me.Lot.Text & "|" & Lname, "") 'ローカル保存
  
    Dim path As String: path = HistoryPath
    Select Case True
    Case UF1.L1Work:   Call WriteLog("Lot登録:" & Me.Lot.Text & ":" & Lname, "", path & "L-1\L-1_Log")
    Case UF1.LVC:      Call WriteLog("Lot登録:" & Me.Lot.Text & ":" & Lname, "", path & "LVC\LVC_Log")
    Case UF1.HVC:      Call WriteLog("Lot登録:" & Me.Lot.Text & ":" & Lname, "", path & "HVC\HVC_Log")
    Case UF1.Bala:     Call WriteLog("Lot登録:" & Me.Lot.Text & ":" & Lname, "", path & "バランサー\バランサー_Log")
    Case UF1.Tott:     Call WriteLog("Lot登録:" & Me.Lot.Text & ":" & Lname, "", path & "トット\トット_Log")
    Case UF1.Test:     Call WriteLog("Lot登録:" & Me.Lot.Text & ":" & Lname, "", path & "Test\Test_Log")
    Case Else:         Call WriteLog("Lot登録:" & Me.Lot.Text & ":" & Lname, "", path & "丸徳\丸徳_Log")
    End Select
    UF1.Machin = ""
  End If
  
  Dim fname As String: fname = "Measured.log"
  If UF1.Tag <> "STOP" Then
    Select Case True
    Case UF1.L1Work:   Call WriteLog("Lot登録:" & Me.Lot.Text & ":" & Lname, "", path & "L-1\L-1_Log", fname)
    Case UF1.LVC:      Call WriteLog("Lot登録:" & Me.Lot.Text & ":" & Lname, "", path & "LVC\LVC_Log", fname)
    Case UF1.HVC:      Call WriteLog("Lot登録:" & Me.Lot.Text & ":" & Lname, "", path & "HVC\HVC_Log", fname)
    Case UF1.Bala:     Call WriteLog("Lot登録:" & Me.Lot.Text & ":" & Lname, "", path & "バランサー\バランサー_Log", fname)
    Case UF1.Tott:     Call WriteLog("Lot登録:" & Me.Lot.Text & ":" & Lname, "", path & "トット\トット_Log", fname)
    Case UF1.Test:     Call WriteLog("Lot登録:" & Me.Lot.Text & ":" & Lname, "", path & "Test\Test_Log", fname)
    Case Else:         Call WriteLog("Lot登録:" & Me.Lot.Text & ":" & Lname, "", path & "丸徳\丸徳_Log", fname)
    End Select
  End If
  
  Select Case True '2021.11.24_導電率チェック
  Case Conductiv_Check: UF1.Conductiv.BackColor = RGB(255, 0, 0)
  Case Else:            UF1.Conductiv.BackColor = &H8000000F
  End Select
  
End If
End Sub
Private Sub Lot_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'入力制限
Select Case KeyAscii.Value
Case 48 To 57 '0~9
Case 65 To 90 'A~Z
Case 97 To 122 'a~z
KeyAscii.Value = KeyAscii.Value - 32 '大文字強制変換
Case Else
KeyAscii.Value = 0
End Select
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
If UF1.Load.Caption = "off" Then Exit Sub

With UF1
  If .Image1.Visible = False Then
    .MPEX.Tag = "" 'ラテラルボー判定表示対策
    .MPEX.Visible = True
    .MP1.Visible = True
    '.ErFM.Visible = True
  
    .PicFM.tOP = 120 '山数選択検査頻度表示
    .PicFM.Left = 12
    .PicFM.Width = 0
    .PicFM.Height = 0
    .OrdNoCB.Visible = True
    .NoL.Visible = True
    .PackingForm.Visible = True
    .PackingComment.Visible = True
    .MPEX.Visible = True
  End If
End With

End Sub

Private Sub WorkButton_Click()
On Error Resume Next
'-----2021.03.04_Focus対策
UF1.Worker.Enabled = True
UF1.INP.Enabled = True
UF1.Manual.Enabled = True
UF1.StartTB.Enabled = True

Select Case True
Case UF1.Bala: GoTo line
Case Else
  If MsgBox("検査員登録しますか", vbYesNo + vbQuestion, "確認") = vbYes Then
    GoTo line
  Else
    Exit Sub
  End If
End Select

line:
  If Me.Lot <> "" Then
     BlinkLabel1.BlinkStop '点滅ｽﾄｯﾌﾟ
  End If
  If Register_Check(Format$(Now, "hh:mm:ss")) Then Call Daily_Create '日報作成
  '-----2021.03.04_Focus対策
  UF1.INP.Enabled = False
  UF1.Manual.Enabled = False
  UF1.StartTB.Enabled = False
  
  If ReadCheck("C:\Windows\System32\ctfmon.exe", 1) Then '2021.04.27_ﾊﾞﾗﾝｻｰ50320対策
    With CreateObject("Wscript.Shell")
      Dim ret As Long: ret = .Run("C:\Windows\System32\ctfmon.exe", 5, True)
    End With
  End If
  '2021.05.25_ctfmon.exeを動かしてからﾌｫｰｶｽｾｯﾄ
  UF1.Worker.BackColor = &HC0FFFF
  UF1.Worker = ""
  AppActivate Application.Caption
  DoEvents
  UF1.INP.SetFocus
  UF1.Worker.SetFocus
  
  UF1.Worker.IMEMode = fmIMEModeOff '2021.04.19_50320PCのみﾌｫｰｶｽが移るとIMEがOffになる
  If ret = 0 Then Exit Sub
End Sub

Private Sub Worker_Change()
If Me.FlagLabel <> 1 Then Exit Sub
UF1.T1.Tag = "" '2021.04.14
UF1.T2.Tag = ""
UF1.T3.Tag = ""
Dim Temp As String
If Len(Me.Worker.Text) <> 13 Then  'ｶｰﾄﾞ16桁　-4000番
Exit Sub
Else
Dim WorName As String: WorName = NameConv(UF1.Worker.Value) '2021.04.16_13桁の時だけ変換
  UF1.FlagLabel = 0
  If WorName = "" Then
    UF1.Worker = ""
    UF1.Worker.BackColor = &H80000005
    UF1.FlagLabel = 1
    UF1.Worker.SetFocus
    Exit Sub
  End If
  Select Case TimeCheck(Format$(Now, "hh:mm:ss"))
  Case "日勤"
    UF1.Day1 = DayCheck(Format$(Now, "hh:mm"))
    Select Case True '2021.04.14_同名登録時は=処理とする**2021.04.15追記しないとchangeeventが起きない(校正)別の関数で短縮する_WorkerSave
    Case UF1.T1 = "":  UF1.T1 = WorName
    Case UF1.T1 <> "": UF1.T1 = UF1.T1 & "_" & WorName
    End Select
    Temp = "*UFKen:検査員登録_" & UF1.Day1 & ":" & UF1.T1
  Case "1直"
    UF1.Day1 = DayCheck(Format$(Now, "hh:mm"))
    Select Case True
    Case UF1.T1 = "":  UF1.T1 = WorName
    Case UF1.T1 <> "": UF1.T1 = UF1.T1 & "_" & WorName
    End Select
    Temp = "*UFKen:検査員登録_" & UF1.Day1 & ":" & UF1.T1
  Case "2直"
    UF1.Day2 = DayCheck(Format$(Now, "hh:mm"))
    Select Case True
    Case UF1.T2 = "":  UF1.T2 = WorName
    Case UF1.T2 <> "": UF1.T2 = UF1.T2 & "_" & WorName
    End Select
    Temp = "*UFKen:検査員登録_" & UF1.Day2 & ":" & UF1.T2
  Case "3直"
    UF1.Day3 = DayCheck(Format$(Now, "hh:mm"))
    Select Case True
    Case UF1.T3 = "":  UF1.T3 = WorName
    Case UF1.T3 <> "": UF1.T3 = UF1.T3 & "_" & WorName
    End Select
    Temp = "*UFKen:検査員登録_" & UF1.Day3 & ":" & UF1.T3
  End Select
Dim path As String: path = HistoryPath
Select Case True
Case UF1.L1Work:   Call WriteLog(Temp, "", path & "L-1\L-1_Log")
Case UF1.LVC:      Call WriteLog(Temp, "", path & "LVC\LVC_Log")
Case UF1.HVC:      Call WriteLog(Temp, "", path & "HVC\HVC_Log")
Case UF1.Bala:     Call WriteLog(Temp, "", path & "バランサー\バランサー_Log")
Case UF1.Tott:     Call WriteLog(Temp, "", path & "トット\トット_Log")
Case UF1.Test:     Call WriteLog(Temp, "") 'path & "Test\Test_Log"
Case Else:         Call WriteLog(Temp, "", path & "丸徳\丸徳_Log")
End Select

UF1.PicFM.Tag = 0 '2023.6.27**LoopJudg用カウント--------------検査員登録時もリセット

Dim LotNo As String: LotNo = UF1.Lot.Text
Dim flag As Boolean: flag = AdvanceCheck(LotNo)
Dim 先行 As String:  先行 = "先行データ票が必要です。"
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
'2023.12.22***追加
If flag Then
  If UF校正.Visible = False Then '校正表示中だと音声が混ざる
    Call MsgError(先行, Lname, 4)
    Call 音楽を流す("必要です.wav")
  End If
End If

Call Daily_Report(1) '日報情報_検査員,ライン名
End If
UF1.Worker = ""
UF1.Worker.BackColor = &H80000005
UF1.FlagLabel = 1
End Sub
Private Sub Worker_Exit(ByVal Cancel As MSForms.ReturnBoolean) 'フレーム内にBoxを入れると発生しません
Me.JudgBox1.BackColor = RGB(0, 0, 0)
End Sub
Private Sub Worker_Enter()
Me.JudgBox1.BackColor = RGB(0, 255, 0)
End Sub

Private Sub ZEB_Click()
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If Register_Check(Format$(Now, "hh:mm:ss"), "") Then '検査員ﾁｪｯｸ *1.ｼｰﾄがない2.3行(Day1,Day2,Day3)に当日以外の登録がある3.その時間に登録が無い
  Call MsgError("検査員登録をして下さい", Lname, 2)
  Exit Sub
End If
Application.Visible = True

Me.Loading.Visible = True
Me.Repaint
UF1.Yama = ""

Application.ScreenUpdating = False
Call HiddenSheets

Call BookSaClose(ThisWorkbook.path, "日報＆資材計算.xlsm") 'ﾊﾟｽ,閉じたいbook名

Call OpenMitubisi("全数")

If UFWorkerName.Visible = True Then Unload UFWorkerName


Me.Loading.Visible = False
Me.Repaint
UF1.Hide
End Sub

Private Sub ZS_Change()
Select Case UF1("CBS1").Text
Case "耳", "中"
    Me.ZS.maxLength = 5
Case Else
    If IsNumeric(UF1.ZS) = False Then UF1.ZS = ""
    Me.ZS.maxLength = 4
End Select
End Sub

Private Sub ZU_Change()
Select Case UF1("CBU1").Text
Case "耳", "中"
    Me.ZU.maxLength = 5
Case Else
    If IsNumeric(UF1.ZU) = False Then UF1.ZU = ""
    Me.ZU.maxLength = 4
End Select
End Sub
Private Sub ZC_Change()
Select Case UF1("CBC1").Text
Case "耳", "中"
    Me.ZC.maxLength = 5
Case Else
    If IsNumeric(UF1.ZC) = False Then UF1.ZC = ""
    Me.ZC.maxLength = 4
End Select
End Sub
Private Sub ZD_Change()
Select Case UF1("CBD1").Text
Case "耳", "中"
    Me.ZD.maxLength = 5
Case Else
    If IsNumeric(UF1.ZD) = False Then UF1.ZD = ""
    Me.ZD.maxLength = 4
End Select
End Sub
Private Sub ZE_Change()
Select Case UF1("CBE1").Text
Case "耳", "中"
    Me.ZE.maxLength = 5
Case Else
    If IsNumeric(UF1.ZE) = False Then UF1.ZE = ""
    Me.ZE.maxLength = 4
End Select
End Sub
Private Sub ZF_Change()
Select Case UF1("CBF1").Text
Case "耳", "中"
    Me.ZF.maxLength = 5
Case Else
    If IsNumeric(UF1.ZF) = False Then UF1.ZF = ""
    Me.ZF.maxLength = 4
End Select
End Sub
'なんかダメだった_2024.2.7
'Private Sub ZS_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'Static PreviousValue As String, CurrentValue As String
'CurrentValue = Me.ZS.value ' 現在のTextBoxの値を取得
'If PreviousValue <> "" And CurrentValue <> "" And PreviousValue <> CurrentValue Then
'  Call HizumiJudg(ZS)
'  Call LogOut(ZS, "AorB")
'End If
'PreviousValue = CurrentValue ' 現在の値を前回の値として保持
'
''Call HizumiJudg(ZS)
''Call LogOut(ZS, "AorB")
'End Sub
'Private Sub ZU_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'Call HizumiJudg(ZU)
'Call LogOut(ZU, "AorB")
'End Sub
'Private Sub ZC_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'Call HizumiJudg(ZC)
'Call LogOut(ZC, "AorB")
'End Sub
'Private Sub ZD_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'Call HizumiJudg(ZD)
'Call LogOut(ZD, "AorB")
'End Sub
'Private Sub ZE_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'Call HizumiJudg(ZE)
'Call LogOut(ZE, "AorB")
'End Sub
Private Sub ZS_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.ZS.SetFocus
UF1.MoveFlag = "on"
End Sub
Private Sub ZU_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.ZU.SetFocus
UF1.MoveFlag = "on"
End Sub
Private Sub ZC_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.ZC.SetFocus
UF1.MoveFlag = "on"
End Sub
Private Sub ZD_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.ZD.SetFocus
UF1.MoveFlag = "on"
End Sub
Private Sub ZE_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.ZE.SetFocus
UF1.MoveFlag = "on"
End Sub
Private Sub ZF_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.ZF.SetFocus
UF1.MoveFlag = "on"
End Sub

Private Sub GS_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.GS.SetFocus
UF1.MoveFlag = "on"
End Sub
Private Sub GU_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.GU.SetFocus
UF1.MoveFlag = "on"
End Sub
Private Sub GC_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.GC.SetFocus
UF1.MoveFlag = "on"
End Sub
Private Sub GD_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.GD.SetFocus
UF1.MoveFlag = "on"
End Sub
Private Sub GE_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.GE.SetFocus
UF1.MoveFlag = "on"
End Sub
Private Sub GF_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next
UF1.GF.SetFocus
UF1.MoveFlag = "on"
End Sub

Private Sub ZS_AfterUpdate()
Call HizumiJudg(ZS)
Call LogOut(ZS, "AorB")
Call Out_Tolerance(UF1.INP.Tag)
Call CheckLabelsByCategory("S") '2025.1.23
End Sub
Private Sub ZU_AfterUpdate()
Call HizumiJudg(ZU)
Call LogOut(ZU, "AorB")
Call Out_Tolerance(UF1.INP.Tag)
Call CheckLabelsByCategory("U")
End Sub
Private Sub ZC_AfterUpdate()
Call HizumiJudg(ZC)
Call LogOut(ZC, "AorB")
Call Out_Tolerance(UF1.INP.Tag)
Call CheckLabelsByCategory("C")
End Sub
Private Sub ZD_AfterUpdate()
Call HizumiJudg(ZD)
Call LogOut(ZD, "AorB")
Call Out_Tolerance(UF1.INP.Tag)
Call CheckLabelsByCategory("D")
End Sub
Private Sub ZE_AfterUpdate()
Call HizumiJudg(ZE)
Call LogOut(ZE, "AorB")
Call Out_Tolerance(UF1.INP.Tag)
Call CheckLabelsByCategory("E")
End Sub
Private Sub ZF_AfterUpdate()
Call HizumiJudg(ZF)
Call LogOut(ZF, "AorB")
Call Out_Tolerance(UF1.INP.Tag)
Call CheckLabelsByCategory("F")
End Sub

Private Sub ZS_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean) '2021.01.26
If IsNumeric(UF1.ZS) Then UF1.ZS = Format((UF1.ZS), "0.00")
Dim n As String: n = Right(Me.AorB.Caption, 1)
Call ReflectStrain(n, "保存")
End Sub
Private Sub ZU_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If IsNumeric(UF1.ZU) Then UF1.ZU = Format((UF1.ZU), "0.00")
Dim n As String: n = Right(Me.AorB.Caption, 1)
Call ReflectStrain(n, "保存")
End Sub
Private Sub ZC_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If IsNumeric(UF1.ZC) Then UF1.ZC = Format((UF1.ZC), "0.00")
Dim n As String: n = Right(Me.AorB.Caption, 1)
Call ReflectStrain(n, "保存")
End Sub
Private Sub ZD_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If IsNumeric(UF1.ZD) Then UF1.ZD = Format((UF1.ZD), "0.00")
Dim n As String: n = Right(Me.AorB.Caption, 1)
Call ReflectStrain(n, "保存")
End Sub
Private Sub ZE_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If IsNumeric(UF1.ZE) Then UF1.ZE = Format((UF1.ZE), "0.00")
Dim n As String: n = Right(Me.AorB.Caption, 1)
Call ReflectStrain(n, "保存")
End Sub
Private Sub ZF_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If IsNumeric(UF1.ZF) Then UF1.ZF = Format((UF1.ZF), "0.00")
Dim n As String: n = Right(Me.AorB.Caption, 1)
Call ReflectStrain(n, "保存")
End Sub

Private Sub GS_Change()
If IsNumeric(UF1.GS) = False Then UF1.GS = ""
Me.GS.maxLength = 4
End Sub
Private Sub GU_Change()
If IsNumeric(UF1.GU) = False Then UF1.GU = ""
Me.GU.maxLength = 4
End Sub
Private Sub GC_Change()
If IsNumeric(UF1.GC) = False Then UF1.GC = ""
Me.GC.maxLength = 4
End Sub
Private Sub GD_Change()
If IsNumeric(UF1.GD) = False Then UF1.GD = ""
Me.GD.maxLength = 4
End Sub
Private Sub GE_Change()
If IsNumeric(UF1.GE) = False Then UF1.GE = ""
Me.GE.maxLength = 4
End Sub

Private Sub GS_AfterUpdate()
Call HizumiJudg(GS)
Call LogOut(GS, "AorB")
Call Out_Tolerance(UF1.INP.Tag) '2024.2.14
Call CheckLabelsByCategory("S") '2025.1.23
End Sub
Private Sub GU_AfterUpdate()
Call HizumiJudg(GU)
Call LogOut(GU, "AorB")
Call Out_Tolerance(UF1.INP.Tag)
Call CheckLabelsByCategory("U")
End Sub
Private Sub GC_AfterUpdate()
Call HizumiJudg(GC)
Call LogOut(GC, "AorB")
Call Out_Tolerance(UF1.INP.Tag)
Call CheckLabelsByCategory("C")
End Sub
Private Sub GD_AfterUpdate()
Call HizumiJudg(GD)
Call LogOut(GD, "AorB")
Call Out_Tolerance(UF1.INP.Tag)
Call CheckLabelsByCategory("D")
End Sub
Private Sub GE_AfterUpdate()
Call HizumiJudg(GE)
Call LogOut(GE, "AorB")
Call Out_Tolerance(UF1.INP.Tag)
Call CheckLabelsByCategory("E")
End Sub
Private Sub GF_AfterUpdate()
Call HizumiJudg(GF)
Call LogOut(GF, "AorB")
Call Out_Tolerance(UF1.INP.Tag)
Call CheckLabelsByCategory("F")
End Sub

Private Sub GS_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean) '2021.01.26
If IsNumeric(UF1.GS) Then UF1.GS = Format((UF1.GS), "0.00")
Dim n As String: n = Right(Me.AorB.Caption, 1)
Call ReflectStrain(n, "保存")
End Sub
Private Sub GU_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If IsNumeric(UF1.GU) Then UF1.GU = Format((UF1.GU), "0.00")
Dim n As String: n = Right(Me.AorB.Caption, 1)
Call ReflectStrain(n, "保存")
End Sub
Private Sub GC_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If IsNumeric(UF1.GC) Then UF1.GC = Format((UF1.GC), "0.00")
Dim n As String: n = Right(Me.AorB.Caption, 1)
Call ReflectStrain(n, "保存")
End Sub
Private Sub GD_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If IsNumeric(UF1.GD) Then UF1.GD = Format((UF1.GD), "0.00")
Dim n As String: n = Right(Me.AorB.Caption, 1)
Call ReflectStrain(n, "保存")
End Sub
Private Sub GE_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If IsNumeric(UF1.GE) Then UF1.GE = Format((UF1.GE), "0.00")
Dim n As String: n = Right(Me.AorB.Caption, 1)
Call ReflectStrain(n, "保存")
End Sub
Private Sub GF_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If IsNumeric(UF1.GF) Then UF1.GF = Format((UF1.GF), "0.00")
Dim n As String: n = Right(Me.AorB.Caption, 1)
Call ReflectStrain(n, "保存")
End Sub

Private Sub 厚板_Click()
'Call LSelect("厚板") '2024.9.26_リセット目的で追加
'Me.Loading.Visible = True
'Me.Repaint
''Call LSelect("厚板")
'If UF1.Sumisei <> "" Then
'  Call LSelect("住精")
'Else
'  Call LSelect("厚板")
'End If
'
'Call OpForcedSelection 'サイズをコンパクトにし作業選択を強調する_2024.5.10
'
'On Error Resume Next
'UF1.Manual.SetFocus
'UF1.StartTB.SetFocus
'Me.Loading.Visible = False
'Me.Repaint

'1回目は描写リセットのみ実行（DB/納入チェックはスキップ）
Dim bSkipDBCheck As Boolean: bSkipDBCheck = True
Call LSelect("厚板", bSkipDBCheck)
Me.Loading.Visible = True
Me.Repaint

'2回目は通常フル処理
If UF1.Sumisei <> "" Then
  Call LSelect("住精")
Else
  Call LSelect("厚板")
End If

Call OpForcedSelection
On Error Resume Next
UF1.Manual.SetFocus
UF1.StartTB.SetFocus
Me.Loading.Visible = False
Me.Repaint

End Sub
Private Sub 少枚数_Click()
Call LSelect("少枚数")

'少枚数用の枚数を確定させる_2023.8.24
Dim th As Long
If UF1.ATU <> "" Then
  If UF1.KenHBox7 = "" Then
    th = Format(Round(CalculateSheets(UF1, UF1.ATU), 1), "0")
  ElseIf UF1.KenCount <> "" Then
    th = UF1.KenCount
  End If
  If th >= 5 Then th = 5 '少枚数は5枚まで
End If

Dim Kno As String: Kno = UF1.厚板.Caption
Dim flag As Boolean: If UF1.少枚数 Then flag = True
If Kno = "H" And flag Then
  Call SmallNumberCount(th)
  Call OpForcedSelection
  Exit Sub
End If

If UF1.Sumisei <> "" Then
  Call LSelect("住精")
Else
  Call LSelect("少枚数")
End If

Call OpForcedSelection
On Error Resume Next

UF1.Manual.SetFocus
UF1.StartTB.SetFocus
End Sub
Private Sub 薄板_Click()
'Me.Loading.Visible = True
'Me.Repaint
''薄板フォーマット描写後、ラテラルボー描写をさせる_2022.09.20
'Call LSelect("薄板")
'If UF1.Sumisei <> "" Then
'  Call LSelect("住精")
'End If
'
'Call OpForcedSelection
'On Error Resume Next
'
'UF1.Manual.SetFocus
'UF1.StartTB.SetFocus
'Me.Loading.Visible = False
'Me.Repaint
Me.Loading.Visible = True
Me.Repaint
Call LSelect("薄板", True)  '1回目：スキップ
If UF1.Sumisei <> "" Then
  Call LSelect("住精")       '2回目：フル処理
End If

End Sub
Private Sub 極厚_Click()
Call LSelect("極厚")

Call OpForcedSelection
On Error Resume Next

UF1.Manual.SetFocus
UF1.StartTB.SetFocus
End Sub

Private Sub スラブ_Click()

If UF1.Sumisei <> "" Then
  Call LSelect("住精")
Else
  Call LSelect("スラブ")
End If

Call OpForcedSelection
On Error Resume Next

UF1.Manual.SetFocus
UF1.StartTB.SetFocus
Me.Loading.Visible = False
Me.Repaint
End Sub

Private Sub 表示_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Application.Visible = True 'Excel表示
Application.ScreenUpdating = True
Application.EnableEvents = True
Call VisibleSheets
Dim wbName As String: wbName = ThisWorkbook.Name
Windows(wbName).Visible = True
End Sub
