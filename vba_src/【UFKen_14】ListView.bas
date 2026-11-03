Attribute VB_Name = "【UFKen_14】ListView"
Option Explicit
'文字幅を測定するAPI利用>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Type RECT
    Left As Long
    tOP As Long
    Right As Long
    Bottom As Long
End Type

#If Win64 Then
  Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
  Private Declare PtrSafe Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hgdiobj As Long) As Long
  Private Declare PtrSafe Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
  
  Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, _
      ByVal nWidth As Long, _
      ByVal nEscapement As Long, _
      ByVal nOrientation As Long, _
      ByVal fnWeight As Long, _
      ByVal IfdwItalic As Long, _
      ByVal fdwUnderline As Long, _
      ByVal fdwStrikeOut As Long, _
      ByVal fdwCharSet As Long, _
      ByVal fdwOutputPrecision As Long, _
      ByVal fdwClipPrecision As Long, _
      ByVal fdwQuality As Long, _
      ByVal fdwPitchAndFamily As Long, _
      ByVal lpszFace As String) As Long
      
  Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
      ByVal lpStr As String, _
      ByVal nCount As Long, _
      lpRect As RECT, _
      ByVal wFormat As Long) As Long
#Else
  Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hgdiobj As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
  
  Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, _
      ByVal nWidth As Long, _
      ByVal nEscapement As Long, _
      ByVal nOrientation As Long, _
      ByVal fnWeight As Long, _
      ByVal IfdwItalic As Long, _
      ByVal fdwUnderline As Long, _
      ByVal fdwStrikeOut As Long, _
      ByVal fdwCharSet As Long, _
      ByVal fdwOutputPrecision As Long, _
      ByVal fdwClipPrecision As Long, _
      ByVal fdwQuality As Long, _
      ByVal fdwPitchAndFamily As Long, _
      ByVal lpszFace As String) As Long
      
  Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
      ByVal lpStr As String, _
      ByVal nCount As Long, _
      lpRect As RECT, _
      ByVal wFormat As Long) As Long
#End If

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_SCRIPT = 64
Private Const DT_CALCRECT = &H400
'文字幅測定関数
Public Function MeasureTextWidth(target_text As String, FONT_NAME As String, Optional font_height As Long = 10) As Long
    Dim hWholeScreenDC As Long: hWholeScreenDC = GetDC(0&)
    Dim hVirtualDC As Long: hVirtualDC = CreateCompatibleDC(hWholeScreenDC)
    Dim hFont As Long: hFont _
        = CreateFont(font_height, 0, 0, 0, FW_NORMAL, _
            0, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, _
            CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, _
            DEFAULT_PITCH Or FF_SCRIPT, FONT_NAME)
    Call SelectObject(hVirtualDC, hFont)
    Dim DrawAreaRectangle As RECT
    Call DrawText(hVirtualDC, target_text, -1, DrawAreaRectangle, DT_CALCRECT)
    Call DeleteObject(hFont)
    Call DeleteObject(hVirtualDC)
    Call ReleaseDC(0&, hWholeScreenDC)
    MeasureTextWidth = DrawAreaRectangle.Right - DrawAreaRectangle.Left
End Function
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'ListViewにフィールド名を設定
Public Sub LvField(Lv As Control, FieldData As Variant)

'.ListItems.Count 'Item数
'.ColumnHeaders.Count 'SubItem数

Dim j As Long
Dim n As Long: n = 0
Dim Lj As Long: Lj = UBound(FieldData, 2)

With Lv
  .Sorted = False
  .Enabled = False            '更新完了までﾛｯｸ
  .ColumnHeaders.clear        'ﾍｯﾀﾞｰをｸﾘｱ
  .View = lvwReport           '表示(ここでエラーが出る場合は、ListViewを貼り直し必須)
  .LabelEdit = lvwManual      'ラベルの編集
  .HideSelection = False      '選択の自動解除
  .HideColumnHeaders = False  '1行目ヘッダーの非表示
  .AllowColumnReorder = True  '列位置変更を許可
  .Appearance = ccFlat        '外観：平面的
  .FullRowSelect = True       '行全体を選択
  .Gridlines = True           'グリッド線
  .ColumnHeaders.Add Text:="", Width:=0 '空で挿入

  For j = 1 To Lj
    'ColumnHeaders.Add Index, Key, Text, Width, Alignment, Icon'列見出し
    If FieldData(1, j) <> "" Then '列見出しがあるもののみ追加
      n = n + 1
      .ColumnHeaders.Add , "Item" & j, FieldData(1, j)
    End If
  Next j
  
  .Enabled = True '更新完了で解除
  
End With

End Sub
Public Sub CreateAll(Lv As Control, FieldData As Variant, TargetData As Variant, Optional TrimColFlag As Boolean = False)
'サイズをゼロにして非表示にする方法
Lv.Visible = False '←この方法は表示位置がずれる
Dim LvHeight As Long: LvHeight = Lv.Height: Lv.Height = 0
Dim LvWidth As Long: LvWidth = Lv.Width: Lv.Width = 0
Call LvField(Lv, FieldData)
Call LvRec(Lv, TargetData, TrimColFlag)
Lv.Height = LvHeight
Lv.Width = LvWidth
Lv.Visible = True
Call ListPatch(Lv) '←位置ｽﾞﾚ対策で対応
End Sub
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'ListViewにフィールド名を設定
Public Sub CreateField(Lv As Control, FieldData As Variant)

'.ListItems.Count 'Item数
'.ColumnHeaders.Count 'SubItem数

If IsEmpty(FieldData) Then Exit Sub

Dim j As Long
Dim n As Long: n = 0
Dim Lj As Long: Lj = UBound(FieldData, 2)

Lv.Visible = False '←この方法でもいいが位置がずれる｡＋位置ｽﾞﾚ対策で対応
'ListView高速処理方法(ｺﾝﾄﾛｰﾙの高さを一旦0にして処理後に戻す。)
Dim LvHeight As Long: LvHeight = Lv.Height
Dim LvWidth As Long: LvWidth = Lv.Width
Lv.Height = 0
Lv.Width = 0

With Lv

  .Enabled = False            '更新完了までﾛｯｸ
  .ColumnHeaders.clear        'ﾍｯﾀﾞｰをｸﾘｱ
  .View = lvwReport           '表示(ここでエラーが出る場合は、ListViewを貼り直し必須)
  .LabelEdit = lvwManual      'ラベルの編集
  .HideSelection = False      '選択の自動解除
  .HideColumnHeaders = False  '1行目ヘッダーの非表示
  .AllowColumnReorder = True  '列位置変更を許可
  .Appearance = ccFlat        '外観：平面的
  .FullRowSelect = True       '行全体を選択
  .Gridlines = True           'グリッド線
  .ColumnHeaders.Add Text:="", Width:=0 '空で挿入

  For j = 1 To Lj
    'ColumnHeaders.Add Index, Key, Text, Width, Alignment, Icon'列見出し
    If FieldData(1, j) <> "" Then '列見出しがあるもののみ追加
      n = n + 1
      .ColumnHeaders.Add , "Item" & j, FieldData(1, j)
    End If
  Next j
  
  .Enabled = True '更新完了で解除
  
End With

DoEvents '内容処理促進
Lv.Height = LvHeight
Lv.Width = LvWidth
Lv.Visible = True
Lv.tOP = Lv.tOP + 1
Lv.tOP = Lv.tOP - 1

End Sub
'ListViewにデータをセット
Public Sub LvRec(Lv As Control, TargetData As Variant, Optional TrimColFlag As Boolean = False)
Dim sngTimer As Single: sngTimer = Timer
'.ListItems.Count 'Item数
'.ColumnHeaders.Count 'SubItem数
Lv.Tag = "" '初期化

Dim i As Long
Dim j As Long
Dim Li As Long: Li = UBound(TargetData, 1)
Dim Lj As Long: Lj = UBound(TargetData, 2)
Dim x() As Long: ReDim x(1 To Lj) As Long '文字幅計算用配列(文字幅から列幅を作る)

Const MAXWIDTH As Long = 30 '上限設定
Const MINWIDTH As Long = 6 '下限設定
Const ALIGNSET As Long = 12 '左寄せにする下限値
Const RATIO As Single = 7.95 '係数

Dim Ck() As Boolean: ReDim Ck(1 To Lj) As Boolean '空白ﾁｪｯｸ用（ﾀｲﾄﾙを除きすべて空白だったらTrue）

With Lv
  .Sorted = False
  .Enabled = False            '更新完了までﾛｯｸ
  .ListItems.clear            'ﾘｽﾄｱｲﾃﾑをｸﾘｱ
  .View = lvwReport           '表示(ここでエラーが出る場合は、ListViewを貼り直し必須)
  .LabelEdit = lvwManual      'ラベルの編集
  .HideSelection = False      '選択の自動解除
  .HideColumnHeaders = False  '1行目ヘッダーの非表示
  .AllowColumnReorder = True  '列位置変更を許可
  .Appearance = ccFlat        '外観：平面的
  .FullRowSelect = True       '行全体を選択
  .Gridlines = True           'グリッド線
  
  Dim buf As Variant
  Dim n As Long
  For i = 1 To Li
    .ListItems.Add.Text = "" '不要列(ﾀｲﾄﾙ表示位置制御できない)を空白で埋めておく
    For j = 1 To Lj
      
      buf = TargetData(i, j)
      
      If IsNull(buf) Then 'Nullチェック
        GoTo skipj:
      End If
      
      .ListItems(i).SubItems(j) = buf
      
      If i = 1 Then
        Ck(j) = (buf = "") '空白列の確認 → ﾃﾞｰﾀ無し:true
        x(j) = MINWIDTH '下限設定
      Else
        If buf <> "" Then Ck(j) = False
        'Maxの文字数の内容を探す。
        'n = GetWidth(buf)’自作幅測定
        n = MeasureTextWidth(CStr(buf), "メイリオ", 1)
        x(j) = IIf(x(j) < n, n, x(j)) 'ネットAPI
        If x(j) > MAXWIDTH Then x(j) = MAXWIDTH   '上限設定
      End If
      
skipj:
    Next j
  Next i
  
  For j = 1 To Lj '列方向題目の数
    If TrimColFlag And Ck(j) Then
      .ColumnHeaders.Item(j).Width = 0
    Else
      .ColumnHeaders.Item(j + 1).Width = x(j) * RATIO '係数
      .ColumnHeaders.Item(j + 1).Alignment = lvwColumnCenter 'lvwColumnRight'列の題目表示位置設定
      If x(j) > ALIGNSET Then
        .ColumnHeaders.Item(j + 1).Alignment = lvwColumnLeft
      End If
    End If
  Next j
  
  .Enabled = True '更新完了で解除
  
End With

'情報を格納(ListViewのListItem数 & ListVeiwのSubItem数
Lv.Tag = "," & Li & "," & Lj & "," & Format$(Timer - sngTimer, "#,0.00") 'ﾃﾞﾊﾞｯｸﾞ用(時間測定)

'LV.ListItem = False 'LV.SelectedItem.SubItems(2)
End Sub
'ListViewにデータをセット
Public Sub CreateList(Lv As Control, TargetData As Variant, Optional TrimColFlag As Boolean = False)
Dim sngTimer As Single: sngTimer = Timer
'.ListItems.Count 'Item数
'.ColumnHeaders.Count 'SubItem数
Lv.Tag = "" '初期化
If IsEmpty(TargetData) Then Exit Sub

On Error Resume Next

Dim i As Long
Dim j As Long
Dim Li As Long: Li = UBound(TargetData, 1)
Dim Lj As Long: Lj = UBound(TargetData, 2)
Dim x() As Long: ReDim x(1 To Lj) As Long '文字幅計算用配列(文字幅から列幅を作る)

Const MAXWIDTH As Long = 30 '上限設定
Const MINWIDTH As Long = 6 '下限設定
Const ALIGNSET As Long = 12 '左寄せにする下限値
Const RATIO As Single = 7.95 '係数

Dim Ck() As Boolean: ReDim Ck(1 To Lj) As Boolean '空白ﾁｪｯｸ用（ﾀｲﾄﾙを除きすべて空白だったらTrue）

Lv.Visible = False '←この方法でもいいが位置がずれる｡＋位置ｽﾞﾚ対策で対応
'ListView高速処理方法(ｺﾝﾄﾛｰﾙの高さを一旦0にして処理後に戻す。)
Dim LvHeight As Long: LvHeight = Lv.Height
Dim LvWidth As Long: LvWidth = Lv.Width
Lv.Height = 0
Lv.Width = 0

With Lv
  .Enabled = False            '更新完了までﾛｯｸ
  .ListItems.clear            'ﾘｽﾄｱｲﾃﾑをｸﾘｱ
  .View = lvwReport           '表示(ここでエラーが出る場合は、ListViewを貼り直し必須)
  .LabelEdit = lvwManual      'ラベルの編集
  .HideSelection = False      '選択の自動解除
  .HideColumnHeaders = False  '1行目ヘッダーの非表示
  .AllowColumnReorder = True  '列位置変更を許可
  .Appearance = ccFlat        '外観：平面的
  .FullRowSelect = True       '行全体を選択
  .Gridlines = True           'グリッド線
  
  Dim buf As Variant
  Dim n As Long
  For i = 1 To Li
    .ListItems.Add.Text = "" '不要列(ﾀｲﾄﾙ表示位置制御できない)を空白で埋めておく
    For j = 1 To Lj
      
      buf = TargetData(i, j)
      'Nullチェック
      If IsNull(buf) Then
        GoTo skipj:
      End If
      
      .ListItems(i).SubItems(j) = buf
      
      If i = 1 Then
        Ck(j) = (buf = "") '空白列の確認 → ﾃﾞｰﾀ無し:true
        x(j) = MINWIDTH '下限設定
      Else
        If buf <> "" Then Ck(j) = False
        'Maxの文字数の内容を探す。
        'n = GetWidth(buf)’自作幅測定
        n = MeasureTextWidth(CStr(buf), "メイリオ", 1)
        x(j) = IIf(x(j) < n, n, x(j)) 'ネットAPI
        If x(j) > MAXWIDTH Then x(j) = MAXWIDTH   '上限設定
      End If
      
skipj:
    Next j
  Next i
  
  For j = 1 To Lj '列方向題目の数
    If TrimColFlag And Ck(j) Then
      .ColumnHeaders.Item(j).Width = 0
    Else
      .ColumnHeaders.Item(j + 1).Width = x(j) * RATIO '係数
      .ColumnHeaders.Item(j + 1).Alignment = lvwColumnCenter 'lvwColumnRight'列の題目表示位置設定
      If x(j) > ALIGNSET Then
        .ColumnHeaders.Item(j + 1).Alignment = lvwColumnLeft
      End If
    End If
  Next j
  
  .Enabled = True '更新完了で解除
  
End With

'情報を格納(ListViewのListItem数 & ListVeiwのSubItem数
Lv.Tag = "," & Lv.ListItems.Count & "," & Lj & "," & Format$(Timer - sngTimer, "#,0.00") 'ﾃﾞﾊﾞｯｸﾞ用(時間測定)

DoEvents '内容処理促進
Lv.Height = LvHeight
Lv.Width = LvWidth
Lv.Visible = True
Lv.tOP = Lv.tOP + 1
Lv.tOP = Lv.tOP - 1

'LV.ListItem = False 'LV.SelectedItem.SubItems(2)
End Sub
'不要な表示を消す(幅ゼロで見えなくする)
Public Sub LvHide(Lv As Control, iArr As Variant)
Dim i As Long
For i = 1 To UBound(iArr)
  Lv.ColumnHeaders.Item(iArr(i)).Width = 0
Next i
End Sub
'ListView複数選択時にすべての選択対象のｲﾝﾃﾞｯｸｽ番号を取得
Public Function ViewSelNumArr(UF As UserForm) As Variant
Dim itmLvw As ListItem
Dim i As Long
Dim Temp As Variant: Temp = ""
For Each itmLvw In UF.LV1.ListItems
  If itmLvw.Selected Then
    Temp = Temp & "," & itmLvw.Index
  End If
Next itmLvw
ViewSelNumArr = Split(Temp, ",")
End Function

'Public Sub ならべかえ()
'With LV1.ColumnHeaders
'  .Item(Doj("行") + 1).Position = 1
'  .Item(Doj("形番") + 1).Position = 2
'  .Item(Doj("色種") + 1).Position = 3
'  .Item(Doj("寸法") + 1).Position = 4
'  .Item(Doj("数量") + 1).Position = 5
'  .Item(Doj("梱包") + 1).Position = 6
'End With
'End Sub

'文字列の半角換算幅を得る
'***(https://qiita.com/yokenzan/items/1f23dfbac5f4c8ed80f1)***
Private Function GetWidth(ByVal str As String) As Long
  Dim length As Single: length = 0
  Dim i As Long
  For i = 1 To Len(str) '文字数回繰り返す。
    ' 一文字ごとに調べ、半角であれば1、全角は2をret_lengthに加える。
    length = length + IIf(IsHalf(Mid$(str, i, 1)), 1, 2)
  Next i
  GetWidth = length
End Function
'半角文字か判定する
Private Function IsHalf(ByVal str As String) As Boolean
  If str = StrConv(str, vbNarrow) Then
    IsHalf = True '半角
  Else
    IsHalf = False '全角
  End If
End Function
'ListViewデータの一括取得
Public Function GetLvData(Lv As Control) As Variant

Dim i As Long
Dim j As Long
Const SV1 As String = "!;!#"  '複雑な文字列を設定
Const SV2 As String = "#;#?"
Dim n As Long: n = 0
Dim buf As String: buf = SV1
Dim Cnt As Variant: Cnt = Split(Lv.Tag, ",")
Dim Temp() As Variant

'Debug.Print LV.ListItems.Count

For i = 1 To Lv.ListItems.Count
  If Lv.ListItems(i).Selected Then
  'Debug.Print LV.ListItems(i).Selected
    For j = 1 To Cnt(2)
      '競合を避ける記号（できるだけ特殊な組み合わせ）で分割
      buf = buf & SV2 & Lv.ListItems(i).SubItems(j) 'ListViewの選択ｱｲﾃﾑ
    Next j
    buf = buf & SV1
    n = n + 1
  End If
Next i

Dim ibuf As Variant: ibuf = Split(buf, SV1)
Dim jbuf As Variant

ReDim Temp(1 To n, 1 To Cnt(2)) As Variant

For i = 1 To n
  jbuf = Split(ibuf(i), SV2)
  For j = 1 To Cnt(2)
    Temp(i, j) = jbuf(j)
  Next j
  Erase jbuf
Next i

GetLvData = Temp

End Function
Public Function GetList(Optional path As String, Optional Attr As Long = vbDirectory) As Variant
'vbNormal     0  標準ファイル
'vbReadOnly   1  読み取り専用ファイル
'vbHidden     2  隠しファイル
'vbSystem     4  システムファイル(windowsのみ)
'vbVolume     8  ボリュームラベル(windowsのみ)
'vbDirectory 16  フォルダ
'vbAlias     64  エイリアスファイル(windows不可)
If path = "" Then path = ThisWorkbook.path
If Not PathCheck(path) Then Exit Function
Dim List As Variant: List = ""
ChDir path
Dim Temp As String: Temp = Dir("*", Attr)
Do While Len(Temp) > 0
  Select Case Attr
    Case vbNormal:    If InStr(Temp, ".") > 1 Then List = List & "," & Temp
    Case vbDirectory: If InStr(Temp, ".") = 0 Then List = List & "," & Temp
    Case 3: If InStr(Temp, ".") <> 1 Then List = List & "," & Temp
  End Select
  Temp = Dir()
Loop
GetList = Split(List, ",")
End Function
'ListView位置ｽﾞﾚ対策
Public Sub ListPatch(Lv As Control)
Application.ScreenUpdating = False
Lv.tOP = Lv.tOP + 1
Lv.tOP = Lv.tOP - 1
Application.ScreenUpdating = True
End Sub

Public Sub ListPosting()  '再転記
Dim Lname As String: Lname = UF1.LineName.Caption & UF1.LineNo.Caption
If WsCheck("Daily" & "_" & Lname) = False Then Exit Sub '-------------------------ｼｰﾄなし
Dim i As Long, a As Long
Dim Cont As Long: Cont = UFNameList.LV1.ListItems.Count 'ListView内の行数
Dim ColCont As Long: ColCont = UC(ThisWorkbook.Sheets("Daily" & "_" & Lname), 2) '*ListView上で表示している列に収める
Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Daily" & "_" & Lname)
'ｸﾘｱ
Dim ClRan As Range: Set ClRan = Range(sh.Cells(3, 1), sh.Cells(20 + Cont, ColCont)) '2021.03.29*消す範囲が小さかった為修正
ClRan.ClearContents
'ｱｲﾃﾑの1行目～最終行までｾﾙに転記
'Stop
For i = 1 To Cont
 For a = 1 To ColCont
  sh.Cells(i + 2, a) = Split(UFNameList.LV1.ListItems.Item(i).SubItems(a), "") '文字列変換
 Next a
Next i
End Sub
