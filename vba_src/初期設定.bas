Attribute VB_Name = "初期設定"
Option Explicit
#If VBA7 And Win64 Then
  '64ビット版
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
  '32ビット版
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public MainPath As String
Public NextPath As String
Public QualityPath As String
Public PackingPath As String
Public HistoryPath As String
Public FactoryyPath As String
Public NisPath As String
Public Sub Setting_Path()
Dim tmpMode As String: tmpMode = GetSetting("UF検査判定システム", "起動設定", "モード", "存在なし")
If tmpMode = "オンライン" Then
  MainPath = "\\nlmsrvngy03\Read\【New】仕掛\台帳\"
  NextPath = "\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\仕掛\"
  QualityPath = "\\nlmsrvngy03\工場内共有\検査データ\品証-梱包やり取りフォルダ\Accdb\"
  PackingPath = "\\nlmfangyshrd\各課共有\0130_日軽稲沢\梱包課\AIM\【■】_参照用ファイル\" '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\検査\"
  HistoryPath = "\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\検査履歴\"
  FactoryyPath = "\\nlmsrvngy03\工場内共有\検査データ\"
  NisPath = "C:\Tools\測定データ電子化"
ElseIf tmpMode = "オフライン" Then
  MainPath = ThisWorkbook.path & "\"
  NextPath = ThisWorkbook.path & "\"
  QualityPath = ThisWorkbook.path & "\"
  PackingPath = ThisWorkbook.path & "\"
  HistoryPath = ThisWorkbook.path & "\"
  FactoryyPath = ThisWorkbook.path & "\"
End If
'fPath =                                      "\\nlmsrvngy03\Read\【New】仕掛\台帳\"
'fPath =                                      "\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\仕掛\"
'
'【UFKen_09】Input Kouf = SaveDateLocate(     "\\nlmsrvngy03\工場内共有\検査データ", "【梱包】" & LN)

'fPath =                                      "\\nlmsrvngy03\工場内共有\検査データ\品証-梱包やり取りフォルダ\Accdb\"
'【UFKen_02】FormRelat PackingNoCheck         "\\nlmsrvngy03\工場内共有\検査データ\包装仕様書"

'fPath =                                      "\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\検査\"
'【UFKen_03】SaveLoad VC_Picture  VCPath =    "\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\VC種類画像\"
'【UFKen_04】Judg  Dim path As String: path = "\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\検査履歴\"
'  Case Else:      fPath =                    "\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\梱包課共有\仕掛\"
End Sub
Public Sub AIM()
SaveSetting "UF検査判定システム", "起動設定", "設備名", "Test"
SaveSetting "UF検査判定システム", "起動設定", "モード", "オンライン"
End Sub
Public Sub HVC()
SaveSetting "UF検査判定システム", "起動設定", "設備名", "HVC"
End Sub
Public Sub test取得()
Debug.Print GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")
End Sub
Public Sub DelSet()
DeleteSetting "UF検査判定システム", "起動設定", "設備名"
DeleteSetting "UF検査判定システム", "起動設定", "モード"
'Debug.Print GetSetting("起動設定", "モード", "オフライン")
'Debug.Print GetSetting("UF検査判定システム", "モード", "存在なし")
End Sub

Public Sub 初期設定値確認(Optional flag As Boolean = False)
Call Setting_Path
Dim tmpFaci As String: tmpFaci = GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")
Dim tmpMode As String: tmpMode = GetSetting("UF検査判定システム", "起動設定", "モード", "存在なし")
Select Case True
Case tmpFaci = "存在なし"
  Sleep 500
  SF3.Show
End Select
End Sub

