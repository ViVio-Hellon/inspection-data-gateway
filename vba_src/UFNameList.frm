Attribute VB_Name = "UFNameList"
Attribute VB_Base = "0{7F261C45-E281-4E3D-A74B-D8D17E42BE54}{7F5B7532-4A89-4A1D-A43D-555ABB1ECAD2}"
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
Private Const CST_INT_FSIZE = 10  'フォントサイズ10の場合は↓の係数0.5
Private Const CST_INT_FCFCT = 0.5 'フォントサイズの係数（行数から全行の幅を取得するため）
       
Public GLB_INT_COL As Integer '列番
Public GLB_INT_ROW As Integer '行番

Private Type POINTAPI
        x As Long
        y As Long
End Type

'GetCursorPosはピクセル取得
'1ピクセル=0.75ポイント
'エクセルではポイントで座標管理してる。
#If Win64 Then
  Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#Else
  Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#End If

Private Sub OnOff_Click()
Select Case UFNameList.OnOff.Caption
Case "VBA停止"
   UFNameList.PasswordTex.Visible = True
   If UFNameList.PasswordTex = "weekend99" Then
    If MsgBox("VBAを停止しますか？", vbYesNo) = vbYes Then
      UFNameList.Content = "VBA停止"
      Application.Visible = True 'Excel表示
      End
    End If
   Else
    UFNameList.PasswordTex.SetFocus
   End If
Case "編集起動"
   Unload Me
   Call NameList_Open(2)
End Select
End Sub

Private Sub UserForm_Initialize()
' 現在のExcelウィンドウ状態を保管
'With Application
'  .WindowState = xlNormal
'  g_sngTopSV = .Top
'  g_sngLeftSV = .Left
'  g_sngWidthSV = .Width
'  g_sngHeightSV = .Height
'End With
    With Frame2
        'プロパティ
        .BorderStyle = fmBorderStyleSingle
        .Caption = ""
        .Height = CST_INT_FSIZE + 9
        .SpecialEffect = fmSpecialEffectFlat
        .Visible = False '最初は非表示
    End With

    With TextBox1
        'プロパティ
        .Font.Size = CST_INT_FSIZE
        .BorderStyle = fmBorderStyleNone
        .Text = ""
        .Height = CST_INT_FSIZE + 9
        .Width = Frame2.Width
        .tOP = 0
        .Left = 0

    End With
'    With LV1
'        'プロパティ
'        .Font.Size = CST_INT_FSIZE
'        .View = lvwReport           '表示
'        .LabelEdit = lvwManual      'ラベルの編集
'        .HideSelection = False      '選択の自動解除
'        .AllowColumnReorder = False '列の移動禁止
'        .FullRowSelect = True       '行全体を選択
'        .Gridlines = True           'グリッド線
'        .FlatScrollBar = True       'trueは無し
'        .Top = 0
'        .Left = 0
'
'        '列見出し
'        .ColumnHeaders.Add , "col1", "1"
'        .ColumnHeaders.Add , "col2", "2"
'        .ColumnHeaders.Add , "col3", "3"
'        .ColumnHeaders.Add , "col4", "4"
'        .ColumnHeaders.Add , "col5", "5"
'        .ColumnHeaders.Add , "col6", "6"
'        '.Height は　subLvwIniSetでやる
'        '.Width は subFrmSclSetでやる
'    End With

   
    subLvwIniSet
   
    With Frame1
        'プロパティ
'        .ScrollBars = fmScrollBarsBoth
        .Caption = ""
        .ScrollHeight = LV1.Height
        '.ScrollWidth は subFrmSclSetでやる
    End With
   
    subFrmSclSet
End Sub
'入力ボックスの入力イベント
Private Sub TextBox1_Change()
    If GLB_INT_COL = 0 Then
        LV1.ListItems(GLB_INT_ROW).Text = TextBox1.Text
    Else
        LV1.ListItems(GLB_INT_ROW).SubItems(GLB_INT_COL) = TextBox1.Text
    End If
End Sub
'スクロールバーを一度消さないと、なぜかバーが初期値になる（LV1_Clickで解除してる）
Private Sub LV1_Enter()
'    Frame1.ScrollBars = fmScrollBarsNone
End Sub
Private Sub LV1_Click()
If UFNameList.Content = "日報閲覧" Then
Dim pntPos As POINTAPI
Dim sglPosX, sglPosY As Single
Dim intColWid As Integer
Dim intColCnt As Integer
'sglLVposL Frame1の左端位置
'sglLVposR Frame1の右端位置
'sglLVposA セル列の左端
'sglLVposB セル列の右端
Dim sglLVposL, sglLVposR, sglLVposA, sglLVposB As Single
Dim i As Integer

'スクロールバーを一度消さないと、なぜかバーが初期値になる←を解除
'    Frame1.ScrollBars = fmScrollBarsBoth
Call GetCursorPos(pntPos)
'ピクセルからポイントへ変換
sglPosX = pntPos.x * 0.75
sglPosY = pntPos.y * 0.75
intColCnt = LV1.ColumnHeaders.Count
'Frame1内の左端位置（最後の +5 は枠幅）
sglLVposL = UFNameList.Left + Frame1.Left + LV1.Left + 5
'Frame1内の右端位置（最後の -5 は枠幅 -13はスクロールバー）
sglLVposR = sglLVposL + Frame1.Width - 5 - 13
'1列目のセル左端セット（スクロールに対応）
sglLVposA = sglLVposL - Frame1.ScrollLeft
'列幅合計用
intColWid = 0
For i = 1 To intColCnt
  sglLVposB = sglLVposA + LV1.ColumnHeaders(i).Width
  'マウス座標が列の幅内か判断
  If (sglLVposA < sglPosX) And (sglLVposB > sglPosX) Then
      If sglLVposA < sglLVposL Then
          'スクロールでクリックセルが左端よりでてる
          Frame1.ScrollLeft = Frame1.ScrollLeft - (sglLVposL - sglLVposA)
      ElseIf (sglLVposB > sglLVposR) And ((sglLVposB - sglLVposA) < (sglLVposR - sglLVposL)) Then
          'スクロールでクリックセルが右端よりでて、列幅がフレーム幅より小さい
          Frame1.ScrollLeft = Frame1.ScrollLeft + (sglLVposB - sglLVposR)
      End If
      '入力用テキストボックスサイズ
      TextBox1.Width = sglLVposB - sglLVposA + 85
      Frame2.Width = sglLVposB - sglLVposA + 86
      'セル番地をグローバル変数に格納
      GLB_INT_COL = i - 1
      GLB_INT_ROW = LV1.SelectedItem.Index
      '列の座標（最後の +2 は微調整）
      Frame2.Left = 10 'intColWid
      '行の座標　見出厚+3
      Frame2.tOP = 2 '+ (LV1.SelectedItem.Index * (CST_INT_FSIZE + CST_INT_FCFCT))
      '入力用ボックスの値セット
      If GLB_INT_COL = 0 Then
          TextBox1.Text = CStr(LV1.ListItems(GLB_INT_ROW).Text)
      Else
          TextBox1.Text = CStr(LV1.ListItems(GLB_INT_ROW).SubItems(GLB_INT_COL))
      End If
      '入力用ボックス表示
      Frame2.Visible = True
  End If
  sglLVposA = sglLVposB
  intColWid = intColWid + LV1.ColumnHeaders(i).Width
Next i
'列幅、行追加対応
subFrmSclSet
End If
End Sub
'ListViewの全列幅、行追加でサイズ変更
Private Sub subFrmSclSet()
'    Dim i As Integer
'    Dim sglWid As Single
'    Dim sglHig As Single
'
'    '横幅
'    sglWid = 0
'    For i = 1 To LV1.ColumnHeaders.Count
'        sglWid = sglWid + LV1.ColumnHeaders(i).Width
'    Next
'
'    LV1.Width = sglWid
'    Frame1.ScrollWidth = sglWid
'
'    'スクロールバーの修正を利用して列縮小時の右端表示対応
'    If Frame1.ScrollLeft > 1 Then
'        Frame1.ScrollLeft = Frame1.ScrollLeft - 1
'        Frame1.ScrollLeft = Frame1.ScrollLeft + 1
'    End If
'
'    '縦幅
'    '見出厚+3　　行数（+2は見出と1行余裕をみる）
'    sglHig = 3 + ((LV1.ListItems.Count + 2) * (CST_INT_FSIZE + CST_INT_FCFCT))
'    LV1.Height = sglHig
'    Frame1.ScrollHeight = sglHig
'
'    'スクロールバーの修正を利用して行変化に対応
'    If Frame1.ScrollTop > 1 Then
'        Frame1.ScrollTop = Frame1.ScrollTop - 1
'        Frame1.ScrollTop = Frame1.ScrollTop + 1
'    End If
End Sub
Private Sub subLvwIniSet()
'    Dim i As Integer
'
'    For i = 1 To 100
'        '1列目
'        LV1.ListItems.Add Text:=CStr(i)
'        '2列目
'        LV1.ListItems(i).SubItems(1) = "aaa"
'        '3列目
'        LV1.ListItems(i).SubItems(2) = "あああ"
'
'        '文字に色付ける場合
'        'LV1.ListItems(i).ForeColor = "&H00FF00"
'        'LV1.ListItems(i).ListSubItems(1).ForeColor = "&H00FF00"
'        'LV1.ListItems(i).ListSubItems(2).ForeColor = "&H00FF00"
'    Next
'
'    '見出厚+3　　行数（+2は見出と1行余裕をみる）
'    LV1.Height = 3 + ((LV1.ListItems.Count + 2) * (CST_INT_FSIZE + CST_INT_FCFCT))
End Sub
Private Sub LV1_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
  If KeyCode = 113 Then LV1.StartLabelEdit
End Sub
Private Sub CommandButton1_Click()
With LV1
  If MsgBox("現在選択されている " & .SelectedItem & " を削除しますか？", 36) = vbYes Then
    .ListItems.Remove .SelectedItem.Index
  End If
End With
End Sub
Private Sub CommandButton2_Click()
'UFNameList.PasswordTex.Visible = True
'If UFNameList.PasswordTex = "weekend99" Then'2021.03.29*作業者に任せる
 If MsgBox("日報ﾃﾞｰﾀを修正しますか？", vbYesNo) = vbYes Then
  Call ListPosting
  Dim sh As Worksheet, Temp1() As Variant, TemP2() As Variant
  Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
  Set sh = ThisWorkbook.Sheets("Daily" & "_" & Lname)
  Temp1 = Range(sh.Cells(2, 1), sh.Cells(2, 25))
  TemP2 = Range(sh.Cells(3, 1), sh.Cells(UR(sh, 6), 25))
  Call CreateField(UFNameList.LV1, Temp1)
  Call CreateList(UFNameList.LV1, TemP2)
 End If
'Else
' UFNameList.PasswordTex.SetFocus
'End If
End Sub
Private Sub CommandButton3_Click()
ThisWorkbook.Sheets("Daily_Test").Activate
Dim i As Long, a As Long
Dim Cont As Long: Cont = UFNameList.LV1.ListItems.Count 'ListView内のアイテム数
Dim ColCont As Long: ColCont = UC(ThisWorkbook.Sheets("Daily_Test"), 2) '*ListView上で表示している列に収める
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Daily_Test")
'ｸﾘｱ
Dim ClRan As Range: Set ClRan = Range(sh.Cells(3, 1), sh.Cells(20 + Cont, ColCont))
ClRan.ClearContents
'ｱｲﾃﾑの1行目～最終行までｾﾙに転記
'Stop
For i = 1 To Cont
 For a = 1 To ColCont
  sh.Cells(i + 2, a) = Split(UFNameList.LV1.ListItems.Item(i).SubItems(a), "") '文字列変換
 Next a
Next i
End Sub
Private Sub UserForm_Layout()
' Excelウィンドウ位置を同期させる
'With Application
'  If .WindowState = xlNormal Then
'    .Top = Me.Top + 1
'    .Left = Me.Left + 1
'  End If
'End With
End Sub
Private Sub UserForm_Activate()
'Application.Visible = True 'Excel表示
'Call KenSheets
'Sheets("Daily_Test").Activate
Application.Visible = False 'Excel非表示
Dim path As String: path = HistoryPath
Dim Temp As String: Temp = "*UFKen:UFNameList起動:" & UFNameList.Content
Select Case True
Case UF1.L1Work:   Call WriteLog(Temp, "", path & "L-1\L-1_Log")
Case UF1.LVC:      Call WriteLog(Temp, "", path & "LVC\LVC_Log")
Case UF1.HVC:      Call WriteLog(Temp, "", path & "HVC\HVC_Log")
Case UF1.Bala:     Call WriteLog(Temp, "", path & "バランサー\バランサー_Log")
Case UF1.Tott:     Call WriteLog(Temp, "", path & "トット\トット_Log")
Case UF1.Test:     Call WriteLog(Temp, "", path & "Test\Test_Log")
Case Else:         Call WriteLog(Temp, "", path & "丸徳\丸徳_Log")
End Select
'With Application ' フォームの後ろにExcelウィンドウを隠す
'  .Top = Me.Top + 1
'  .Left = Me.Left + 1
'  .Width = Me.Width - 2
'  .Height = Me.Height - 2
'End With
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' 保管情報を元にExcelウィンドウを元に戻す
'With Application
'  .Top = g_sngTopSV
'  .Left = g_sngLeftSV
'  .Width = g_sngWidthSV
'  .Height = g_sngHeightSV
'End With
Select Case UFNameList.Content
Case "日報閲覧"
'2021.06.02_作業者編集可なのでend処理停止
'   Application.Visible = True 'Excel表示
'   End
Case "検査員不一致"
   UF1.Show vbModeless
   DoEvents
Case "VBA停止"
   Application.Visible = True 'Excel表示
   End
End Select
End Sub

Private Sub CB1_Click()
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("WorkerTime" & "_" & Lname) = False Then Exit Sub '-------------------------ｼｰﾄなし
If Not UFNameList.Content = "検査員不一致" Then Exit Sub
On Error Resume Next
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("WorkerTime" & "_" & Lname)
Dim x As Long: x = UR(sh, 15)
Dim y As Long: y = UR(sh, 16)
Dim z As Long: z = UR(sh, 17)
Dim xyz As Long: xyz = Application.WorksheetFunction.Max(x, y, z)
'Stop
Dim Temp1() As Variant, TemP2() As Variant
Select Case True
Case CB1.Caption = "不一致ﾘｽﾄ"
  Temp1 = Range(sh.Cells(1, 15), sh.Cells(1, 17)) '重複除外ﾘｽﾄ
  TemP2 = Range(sh.Cells(2, 15), sh.Cells(xyz, 17)) '重複除外ﾘｽﾄ'最終行の一番大きい列
  Call CreateField(UFNameList.LV1, Temp1)
  Call CreateList(UFNameList.LV1, TemP2)
  UFNameList.LV1.ColumnHeaders.Item(2).Width = 90
  UFNameList.LV1.ColumnHeaders.Item(3).Width = 90
  UFNameList.LV1.ColumnHeaders.Item(4).Width = 90
  CB1.Caption = "全ﾘｽﾄ"
Case Else
  Temp1 = Range(sh.Cells(2, 1), sh.Cells(2, 12))
  TemP2 = Range(sh.Cells(3, 1), sh.Cells(UR(sh, 1), 12))
  Call CreateField(UFNameList.LV1, Temp1)
  Call CreateList(UFNameList.LV1, TemP2)
  CB1.Caption = "不一致ﾘｽﾄ"
End Select
End Sub
Private Sub 表示_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Application.Visible = True 'Excel表示
Application.ScreenUpdating = True
Application.EnableEvents = True
End Sub

