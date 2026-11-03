Attribute VB_Name = "UF_PackNo"
Attribute VB_Base = "0{4E0B35F2-5856-4BDF-8861-0C0E50245169}{45FD7BC9-DAED-43EC-A4A0-745D7C6D64BF}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'--------UFﾄｯﾌﾟ記述
Private PacHiki(1 To 5) As New LabelCls
Private PacOrde(1 To 5) As New LabelCls
Private Packing(1 To 5) As New LabelCls

'Private Sub Hiki1_Click() 'Clsモジュールに記載
'Call Loop_Label(Me.Hiki1)
'End Sub
Private Sub CommandButton1_Click()
'Call IeSiteSetTextData("")
Dim n As String: n = UF_PackNo("L1").Caption
If n <> "" Then Call IeSiteSetTextData(n)
'Call EdgeSiteSetTextData(n) 'NIS都合　edge使用不可**2022.02末
End Sub
Private Sub CommandButton2_Click()
On Error Resume Next
'Dim AppName As String: AppName = "msedgedriver_97.0.1072.69.exe"
'Dim Path_webDriver As String: Path_webDriver = "\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\【0】_VBA\【00】VBAｴｸｽﾎﾟｰﾄ\Web\詰め合わせパック" & "\" & AppName 'ThisWorkbook.path
'Dim Path_target As String: Path_target = "http://nlmfangydev2:9083/PkgSpecWeb/#/hososhiyosho/genba/etsuran"

If UF_PackNo.CB1 Then
'  Driver.Shutdown '【UFKen_18】IE'NIS都合　edge使用不可**2022.02末
  Call ie_kill
  Unload UF_PackNo
  
  If UFMitsubishi.Visible Or UFCask.Visible Then Exit Sub
  
  If UFMsgBox.Visible = False Then
    UF1.Show vbModeless '2023.10.10
    DoEvents
  End If
  
  'VBAのエディタを開いている状態でform2をunloadするとform1も一緒に閉じてしまう現象有
Else
  MsgBox "確認後に押してください。", vbCritical
End If
End Sub



Private Sub SpinButton1_SpinDown()
Dim lngImageHeight As Long
Dim lngImageWidth  As Long
Dim lngX           As Long
Dim lngY           As Long

With UF_PackNo
'画像の読込
  If lngImageHeight = 0 Then           '初めてクリックした時
      lngImageHeight = .PacIM.Height   '変数に元の画像サイズを記録する
      lngImageWidth = .PacIM.Width
  End If
'  .PacIM.Stretch = True
  .PacIM.Height = lngImageHeight
  .PacIM.Width = lngImageWidth
    '縮小のMAXの設定
    If lngImageHeight * 0.3 >= .PacIM.Height Then Exit Sub
    'イメージのサイズが変更できるように
'    .PacIM.Stretch = True
    'クリックの都度１０％づつ縮小する
    .PacIM.Height = .PacIM.Height * 0.9
    .PacIM.Width = .PacIM.Width * 0.9
End With
End Sub

Private Sub SpinButton1_SpinUp()
Dim lngImageHeight As Long
Dim lngImageWidth  As Long
Dim lngX           As Long
Dim lngY           As Long

With UF_PackNo
'画像の読込
  If lngImageHeight = 0 Then           '初めてクリックした時
      lngImageHeight = .PacIM.Height   '変数に元の画像サイズを記録する
      lngImageWidth = .PacIM.Width
  End If
'  .PacIM.Stretch = True
  .PacIM.Height = lngImageHeight
  .PacIM.Width = lngImageWidth

  '拡大のMAXの設定
  If lngImageHeight * 2 <= .PacIM.Height Then Exit Sub
  'イメージのサイズが変更できるように
'  .PacIM.Stretch = True
  'クリックの都度１０％づつ拡大する
  .PacIM.Height = .PacIM.Height * 1.1
  .PacIM.Width = .PacIM.Width * 1.1
End With
End Sub

Private Sub UserForm_Initialize()
Dim i As Long, a As Long
For i = 1 To 5
  For a = 1 To 3
    Select Case a
    Case 1: PacHiki(i).EvHiki UF_PackNo("Hiki" & i), i
    Case 2: PacOrde(i).EvOrde UF_PackNo("Ord" & i), i
    Case 3: Packing(i).EvPack UF_PackNo("Pac" & i), i
    End Select
  Next a
Next i

Call GetZoom(Me)
Dim Temp As Variant:  Temp = InitialSetting(Me)
Call DebugLog("UF_PackNo Zoom=" & UF_PackNo.Zoom & " Width=" & UF_PackNo.Width & " Height=" & UF_PackNo.Height)
End Sub
Private Sub UserForm_Activate()
Call FormSetting

Call WindowMenuButtonDel(UF_PackNo)
On Error Resume Next

If WsCheck("Lot情報履歴") = False Then Exit Sub  '-------------------------ｼｰﾄなし処理停止
If UF_PackNo.TBLoad = "" Then Exit Sub
Dim LotNo As String: LotNo = PcLot
If LotNo = "" Then Exit Sub
Dim HisSH As Worksheet:   Set HisSH = ThisWorkbook.Sheets("Lot情報履歴")
Dim HisRan As Range:      Set HisRan = Range(HisSH.Cells(1, 1), HisSH.Cells(UR(HisSH, 1), UC(HisSH, 1)))
Dim Hist(): Hist = HisRan
Dim i As Long, a As Long, n As Long: n = 0
Dim R_Ord As Long, R_Packing As Long, R_Lot As Long, R_Hik As Long
For i = 1 To UBound(Hist, 2) 'ﾌｨｰﾙﾄﾞ名
  Select Case Hist(1, i)
  Case "ﾛｯﾄNo":        R_Lot = i
  Case "ｵｰﾀﾞｰNo":      R_Ord = i
  Case "引当NO":       R_Hik = i
  Case "包装仕様NO":   R_Packing = i
  End Select
Next i
Dim FL As Boolean: FL = False
Dim TempHiki()
For i = LBound(Hist, 1) To UBound(Hist, 1) '1次元ﾁｪｯｸ
  If Hist(i, R_Lot) = LotNo Then a = a + 1
Next i

If a = 0 Then Exit Sub

If a >= 6 Then a = 5

ReDim TempHiki(1 To a, 1 To 4)
For i = LBound(Hist, 1) To UBound(Hist, 1)
  If Hist(i, R_Lot) = LotNo Then '一致の時
    FL = True
    n = n + 1
    TempHiki(n, 1) = Hist(i, R_Ord)
    TempHiki(n, 2) = Hist(i, R_Packing)
    TempHiki(n, 3) = (Hist(i, R_Hik))
    TempHiki(n, 4) = (Hist(i, R_Lot))
  End If
Next i
Call merge_sort2(TempHiki, 3) '配列内の並び替え_2022.02.01***引当No基準ｿｰﾄ
If IsEmpty(TempHiki) Then Exit Sub
UF_PackNo.Caption = TempHiki(1, 4)
For i = 1 To a
  UF_PackNo("Hiki" & i) = TempHiki(i, 3)
  UF_PackNo("Ord" & i) = TempHiki(i, 1)
  UF_PackNo("Pac" & i) = TempHiki(i, 2)
Next i

Call PackSelect(UF_PackNo("Pac" & 1))

Dim countone As Long:   countone = LBound(TempHiki, 1)
Dim tempOder As String: tempOder = TempHiki(countone, 1)

Dim result1 As String: result1 = MaterialsUpperPart(tempOder)         ' ザラ板/アングル
Dim result2 As String: result2 = MaterialsUpperPart(tempOder, True)   ' 当板有無

If result1 <> "" And result1 <> "一致なし" Then
  MsgBox result1
End If
If result1 = "アングル" Then
    If result2 <> "" And result2 <> "一致なし" Then MsgBox result2
End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then
  Cancel = True
End If
End Sub

