Attribute VB_Name = "TabletTest"
Option Explicit
Private Const fmTabStyleFixedSize As Long = 3
Public VCA有 As Boolean
Public VCB有 As Boolean
Public 新規 As Boolean
Public Sub VC描写有無()
If UF1.VC_L_A.Tag <> "" Then VCA有 = True
If UF1.VC_L_B.Tag <> "" Then VCB有 = True
If UF1.NewKen.Tag = "新規" Then 新規 = True
End Sub

Public Sub AdjustFormToScreenSize(ByRef form As Object)
Dim screenWidth As Long, screenHeight As Long
Dim formWidth As Long, formHeight As Long
Dim widthRatio As Double, heightRatio As Double
Dim zoomFactor As Double
Dim ctrl As Control

' スクリーンサイズを取得（Twips単位）
screenWidth = Application.Width
screenHeight = Application.Height

' フォームの元のサイズを保存
formWidth = form.Width
formHeight = form.Height

If VCA有 Or VCB有 Then formHeight = formHeight + 158
If 新規 Then formHeight = 545

' 幅と高さの比率を計算
widthRatio = screenWidth / formWidth
heightRatio = screenHeight / formHeight

' ズーム係数を決定（小さい方の比率を使用）
zoomFactor = IIf(widthRatio < heightRatio, widthRatio, heightRatio)

' フォームがスクリーンより小さい場合、ズームインする
If zoomFactor > 1 Then
    ' ズーム係数を適用（最大2倍まで）
    zoomFactor = IIf(zoomFactor > 2, 2, zoomFactor)
    form.Zoom = zoomFactor * 100 ' Zoomは％単位で設定

    ' フォームサイズを調整
    form.Width = formWidth * zoomFactor
    form.Height = formHeight * zoomFactor
Else
    ' フォームがスクリーンより大きい場合、ズームアウトする
    form.Zoom = zoomFactor * 100 ' Zoomは％単位で設定

    ' フォームサイズをスクリーンサイズに合わせる
    form.Width = screenWidth
    form.Height = screenHeight

    ' フォームを中央に配置
    form.Left = (screenWidth - form.Width) / 2
    form.tOP = (screenHeight - form.Height) / 2
End If

' フォーム内のすべてのコントロールに対してサイズ調整を行う
For Each ctrl In form.Controls
    If TypeName(ctrl) = "Frame" Then
        AdjustFrameSize ctrl, zoomFactor
    End If
Next ctrl
End Sub

Private Sub AdjustFrameSize(ByRef frame As MSForms.frame, ByVal zoomFactor As Double)
' Frameのサイズを調整
frame.Width = frame.Width * zoomFactor
frame.Height = frame.Height * zoomFactor

' Frameの位置を調整
frame.Left = frame.Left * zoomFactor
frame.tOP = frame.tOP * zoomFactor

' Frame内のコントロールに対してもサイズ調整を行う
Dim ctrl As Control
For Each ctrl In frame.Controls
    AdjustControlSize ctrl, zoomFactor
Next ctrl
End Sub

Private Sub AdjustControlSize(ByRef ctrl As Control, ByVal zoomFactor As Double)
' コントロールのサイズを調整
ctrl.Width = ctrl.Width * zoomFactor
ctrl.Height = ctrl.Height * zoomFactor

' コントロールの位置を調整
ctrl.Left = ctrl.Left * zoomFactor
ctrl.tOP = ctrl.tOP * zoomFactor

' フォントサイズを調整（オプション）
If ctrl.Font.Size > 0 Then
    ctrl.Font.Size = ctrl.Font.Size * zoomFactor
End If
End Sub
' AdjustFormToScreenSize サブルーチンは変更なし
'Public Sub AdjustFormToScreenSize(ByRef form As Object)
'Dim screenWidth As Long, screenHeight As Long
'Dim formWidth As Long, formHeight As Long
'Dim widthRatio As Double, heightRatio As Double
'Dim zoomFactor As Double
'
'' スクリーンサイズを取得（Twips単位）
'screenWidth = Application.Width
'screenHeight = Application.Height
'
'' フォームの元のサイズを保存
'formWidth = form.Width
'formHeight = form.Height
'
'' 幅と高さの比率を計算
'widthRatio = screenWidth / formWidth
'heightRatio = screenHeight / formHeight
'
'' ズーム係数を決定（小さい方の比率を使用）
'zoomFactor = IIf(widthRatio < heightRatio, widthRatio, heightRatio)
'
'' フォームがスクリーンより小さい場合、ズームインする
'If zoomFactor > 1 Then
'  ' ズーム係数を適用（最大2倍まで）
'  zoomFactor = IIf(zoomFactor > 2, 2, zoomFactor)
'  form.Zoom = zoomFactor * 100 ' Zoomは％単位で設定
'
'  ' フォームサイズを調整
'  form.Width = formWidth * zoomFactor
'  form.Height = formHeight * zoomFactor
'Else
'  ' フォームがスクリーンより大きい場合、ズームアウトする
'  form.Zoom = zoomFactor * 100 ' Zoomは％単位で設定
'
'  ' フォームサイズをスクリーンサイズに合わせる
'  form.Width = screenWidth
'  form.Height = screenHeight
'
'  ' フォームを中央に配置
'  form.Left = (screenWidth - form.Width) / 2
'  form.Top = (screenHeight - form.Height) / 2
'
'End If
'End Sub
Public Sub SaveOriginalControlInfo(ByRef form As Object, ByRef controlsInfo As Collection)
Dim ctrl As Control
Dim ctrlInfo As ControlInfo
For Each ctrl In form.Controls
  Set ctrlInfo = New ControlInfo
  ctrlInfo.InitFromControl ctrl
  If TypeName(ctrl) = "MultiPage" Then
      ctrlInfo.IsMultiPage = True
      ctrlInfo.PageCount = ctrl.Pages.Count
  ElseIf TypeName(ctrl) = "TabStrip" Then
      ctrlInfo.IsMultiPage = True
      ctrlInfo.PageCount = ctrl.Tabs.Count
  End If
  controlsInfo.Add ctrlInfo, ctrl.Name
Next ctrl
End Sub
Public Sub ResizeAndRepositionControls(ByRef form As Object, ByRef controlsInfo As Collection, ByVal widthRatio As Double, ByVal heightRatio As Double)
Dim ctrl As Control
Dim ctrlInfo As ControlInfo
Dim i As Long
Const fmTabStyleDefault As Long = 0
For Each ctrl In form.Controls
Set ctrlInfo = controlsInfo(ctrl.Name)

ctrl.Left = ctrlInfo.Left * widthRatio
ctrl.tOP = ctrlInfo.tOP * heightRatio
ctrl.Width = ctrlInfo.Width * widthRatio
ctrl.Height = ctrlInfo.Height * heightRatio

On Error Resume Next
If ctrlInfo.FontSize > 0 Then
    ctrl.Font.Size = ctrlInfo.FontSize * (widthRatio + heightRatio) / 3
End If

If ctrlInfo.IsMultiPage Then
    If TypeName(ctrl) = "MultiPage" Then
        ' マルチページの各ページのcaptionフォントサイズを調整
        For i = 0 To ctrl.Pages.Count - 1
            ctrl.Pages(i).Font.Size = ctrlInfo.FontSize * (widthRatio + heightRatio) / 15

            ' 各ページのサイズを調整
            ctrl.Pages(i).Width = ctrl.Width - 10 ' 10はマージン
            ctrl.Pages(i).Height = ctrl.Height - 25 ' 25はタブの高さ（おおよその値）
        Next i

        ' マルチページのページ数が変更されていた場合の処理
        If ctrl.Pages.Count > ctrlInfo.PageCount Then
            For i = ctrlInfo.PageCount To ctrl.Pages.Count - 1
                ctrl.Pages(i).Font.Size = ctrlInfo.FontSize * (widthRatio + heightRatio) / 2
                ctrl.Pages(i).Width = ctrl.Width - 10
                ctrl.Pages(i).Height = ctrl.Height - 25
            Next i
        End If

        ' タブの幅と高さを調整
        Dim tabWidth As Integer
        Dim tabHeight As Integer
        tabWidth = (ctrl.Width - 10) \ ctrl.Pages.Count  ' 10はマージン
        tabHeight = ctrl.Font.Size * 2.1  ' フォントサイズの2.1倍をタブの高さとする
        ctrl.TabFixedHeight = tabHeight

        ' タブのレイアウトスタイルを調整
        ctrl.TabOrientation = fmTabOrientationTop
        ctrl.TabStyle = fmTabStyleDefault

        ' 最小タブ幅を設定（必要に応じて調整）
        Const MIN_TAB_WIDTH As Integer = 80
        If tabWidth < MIN_TAB_WIDTH Then
            tabWidth = MIN_TAB_WIDTH
        End If

        ' タブの幅を設定
        ctrl.TabFixedWidth = tabWidth

        For i = 0 To ctrl.Pages.Count - 1
            ctrl.Pages(i).Caption = PadWithSpaces(ctrl.Pages(i).Caption, tabWidth \ 15) ' 7は文字幅の概算
        Next i

    ElseIf TypeName(ctrl) = "TabStrip" Then
        ' TabStripのタブ幅と高さを調整
        ctrl.Style = fmTabStyleFixedSize  ' タブの固定サイズスタイルを設定

        Dim tsTabWidth As Integer
        Dim tsTabHeight As Integer
        tsTabWidth = ctrl.Width \ ctrl.Tabs.Count
        tsTabHeight = ctrl.Font.Size * 1.5 ' フォントサイズの1.5倍をタブの高さとする
        ctrl.TabFixedHeight = tsTabHeight

        For i = 1 To ctrl.Tabs.Count
            ctrl.Tabs(i).Caption = PadWithSpaces(ctrl.Tabs(i).Caption, tsTabWidth \ 7)
        Next i
    End If
End If
On Error GoTo 0
Next
End Sub


' 文字列を指定の長さになるようにスペースで埋める関数
Private Function PadWithSpaces(ByVal str As String, ByVal length As Integer) As String
If Len(str) >= length Then
  PadWithSpaces = Left$(str, length)
Else
  PadWithSpaces = str & Space$(length - Len(str))
End If
End Function
