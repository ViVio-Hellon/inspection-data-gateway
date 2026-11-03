Attribute VB_Name = "clsBlinkLabel"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private WithEvents mctlLabel As MSForms.CommandButton   '点滅させるラベル
Attribute mctlLabel.VB_VarHelpID = -1

'-------------------------------------------------------
' ミリ秒単位でWait
#If VBA7 And Win64 Then
  Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
  Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'Wave音を鳴らす関数
#If VBA7 And Win64 Then
  Private Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
        (ByVal lpszName As String, ByVal hModule As LongPtr, ByVal dwFlags As Long) As Long
#Else
  Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
        (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
#End If
                     
'Windowsディレクトリのパスを取得する
#If VBA7 And Win64 Then
  Private Declare PtrSafe Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
        (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#Else
  Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
        (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#End If

Private Const MAX_PATH = 260
'PlaySoundの演奏方法
Private Const SND_ASYNC = &H1&           '非同期で再生
Private Const SND_FILENAME = &H20000     'lpszNameはファイル名

'-------------------------------------------------------
Private lngSaveLabelForeColor As Long
Private lngBlinkForeColor As Long
Private lngBlinkBackColor As Long

Private blnBlinkStop As Boolean

'---( cmd プロパティ)------------------------------------------------
Public Property Let cmd(ctlNewLabel As MSForms.CommandButton)
    Set mctlLabel = ctlNewLabel
End Property

'---( Blink メソッド)------------------------------------------------
' BlinkColor1:省略時、Labelの文字色
' BlinkColor2:省略時、Labelの背景色
' Buzzer     :点滅に合わせて音を鳴らす場合は[True]
' Sound      :0(Beep音),1(Ding.wav),2(Chord.wav),3(Notify.wav)
'             Ding.wav/Chord.wav/Notify.wavは
'            「Windowsフォルダ\Media」内のファイル
Public Sub Blink(ByVal BlinkCount As Long, _
                 Optional ByVal BlinkColor1 As Variant, _
                 Optional ByVal BlinkColor2 As Variant, _
                 Optional ByVal Buzzer As Boolean = False, _
                 Optional ByVal Sound As Integer = 0)
                 
Dim strBuffer As String * MAX_PATH
Dim strPath As String
Dim strWaveFile As String
Dim lngResult As Long
Dim i As Long
On Error Resume Next
#If VBA7 And Win64 Then
  Const cst_hModule0 As LongPtr = 0
#Else
  Const cst_hModule0 As Long = 0
#End If

'Blink数
    If (BlinkCount <= 1) Then
        Exit Sub
    End If
    
'BlinkColorのチェック
    If IsMissing(BlinkColor1) Then
        lngBlinkForeColor = mctlLabel.ForeColor
    Else
        Select Case VarType(BlinkColor1)
          Case vbInteger, vbLong
            Select Case BlinkColor1
              'vbScrollBars(&H80000000)～vbInfoBackground(&H80000018)
              Case vbScrollBars To vbInfoBackground
                lngBlinkForeColor = BlinkColor1
              'RGB
              Case &H0& To &HFFFFFF
                lngBlinkForeColor = BlinkColor1
              Case Else
                lngBlinkForeColor = mctlLabel.ForeColor
            End Select
          Case Else
            lngBlinkForeColor = mctlLabel.ForeColor
        End Select
    End If
    
    If IsMissing(BlinkColor2) Then
        lngBlinkBackColor = mctlLabel.BackColor
    Else
        Select Case VarType(BlinkColor2)
          Case vbInteger, vbLong
            Select Case BlinkColor2
              'vbScrollBars(&H80000000)～vbInfoBackground(&H80000018)
              Case vbScrollBars To vbInfoBackground
                lngBlinkBackColor = BlinkColor2
              'RGB
              Case &H0& To &HFFFFFF
                lngBlinkBackColor = BlinkColor2
              Case Else
                lngBlinkBackColor = mctlLabel.BackColor
            End Select
          Case Else
            lngBlinkBackColor = mctlLabel.BackColor
        End Select
    End If

' Buzzer指定
    If (Buzzer = True) Then
        Select Case Sound
          Case 1, 2, 3
            'Windowsディレクトリのパス名を取得
            lngResult = GetWindowsDirectory(strBuffer, Len(strBuffer))
            strPath = Left(strBuffer, InStr(strBuffer, vbNullChar) - 1)
            Select Case Sound
              Case 1
                strWaveFile = strPath & "\Media\Ding.wav"    '一般の警告音
              Case 2
                strWaveFile = strPath & "\Media\Chord.wav"   'メッセージ(警告,情報,問合せ)
              Case 3
                strWaveFile = strPath & "\Media\Notify.wav"  'メール着信音
            End Select
          Case Else
            'Beep音を使う
        End Select
    End If

'Blinkループ
    lngSaveLabelForeColor = mctlLabel.ForeColor
    blnBlinkStop = False
    i = 0
    Do While (blnBlinkStop = False)
        mctlLabel.ForeColor = lngBlinkBackColor
        DoEvents
        Sleep 500
        mctlLabel.ForeColor = lngBlinkForeColor
        If (Buzzer = True) Then
            Select Case Sound
              Case 1, 2, 3
                'Wave音を鳴らす
                lngResult = PlaySound(strWaveFile, cst_hModule0, _
                                SND_ASYNC Or SND_FILENAME)
              Case Else
                Beep
            End Select
        Else
            'Beepなし
        End If
        DoEvents
        Sleep 500
        i = i + 1
        If (i >= BlinkCount) Then
            blnBlinkStop = True
        End If
    Loop
'Labelの文字色を元に戻す
    mctlLabel.ForeColor = lngSaveLabelForeColor
End Sub

'---( BlinkStop メソッド)------------------------------------------------
Public Sub BlinkStop()
    '点滅中止フラグをオンにする
    blnBlinkStop = True
End Sub

'---( Label Click イベント クラスモジュール内のみ )------------------------
'ラベルのクリックで点滅中止
Private Sub mctlLabel_Click()
    Me.BlinkStop
End Sub
'
'=====[AddinBox(K.Tsunoda) CopyRight(C) 2004 Allrights Reserved.]========



