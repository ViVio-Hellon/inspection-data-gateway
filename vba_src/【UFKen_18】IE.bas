Attribute VB_Name = "【UFKen_18】IE"
Dim objShell As Object
Dim objExec As Object

#If VBA7 And Win64 Then
  '64ビット版
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
  '32ビット版
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Public AppName As String
Public Path_webDriver As String
Public Path_target As String
Public Driver As New WebDriver

Public Sub EdgeSiteSetTextData(PackingNo As String)
'Dim AppName As String: AppName = "msedgedriver_97.0.1072.69.exe"
'Dim Path_webDriver As String: Path_webDriver = "C:\Users\xfa-ngykonpo\Desktop\詰め合わせパック" & "\" & AppName 'ThisWorkbook.path
'Dim Path_target As String: Path_target = "http://nlmfangydev2:9083/PkgSpecWeb/#/hososhiyosho/genba/etsuran"
On Error Resume Next
AppName = "msedgedriver_97.0.1072.69.exe"
Path_webDriver = "\\nlmfangyshrd\各課共有\0130_日軽稲沢\梱包課\AIM\検査\詰め合わせパック" & "\" & AppName 'ThisWorkbook.path
Path_target = "http://nlmfangyweb1a:9084/PkgSpecWeb/#/hososhiyosho/genba/etsuran" '"http://nlmfangydev2:9083/PkgSpecWeb/#/hososhiyosho/genba/etsuran"
'最後topで起動前が開ける
Sleep 3000
'Dim Driver As New WebDriver
Driver.Edge Path_webDriver
Driver.OpenBrowser
Driver.Navigate Path_target

Sleep 1000
Dim searchInput
Set searchInput = Driver.FindElement(By.Name, "pkgSpecNo")

searchInput.SetValue PackingNo '"5C1129"

Sleep 1000

'DoEvents
Driver.FindElement(By.ClassName, "btn btn-info c-btn-size")(1).Click 'By.Name, By.ID, By.ClassName
On Error GoTo 0
'Sleep 3000
'Driver.Execute Driver.CMD_REFRESH ' 再読み込み

Dim dProcessId  As Double: dProcessId = Shell(Path_webDriver) ' 再読み込みの代わり

'Sleep 5000
'Driver.Shutdown

'  Dim batName As String: batName = "Kill.bat"
'  Dim batPath As String: batPath = ThisWorkbook.path & "\" & batName
'  Dim dProcessId  As Double: dProcessId = Shell(batPath & " " & AppName)
End Sub

Public Sub IeSiteSetTextData(PackingNo As String)   '検索バーにテキストを入力する

On Error Resume Next
Dim ie As InternetExplorer: Set ie = New InternetExplorerMedium 'CreateObject("InternetExplorer.Application") 'IEオブジェクトを作成
Dim strUrl As String: strUrl = "http://nlmfangyweb1a:9084/PkgSpecWeb/#/hososhiyosho/genba/etsuran" '開きたいURLを変数に代入
Dim winshell As Object: Set winshell = CreateObject("Shell.Application")
ie.Visible = True 'IEを表示（見えるようにする）
ie.Navigate strUrl '指定したURLをIEで開く
    
Dim win As Object

'■シェルオブジェクトを使用し、IEを再度掴みなおす。
For Each win In winshell.Windows
  If win.Name = "Internet Explorer" Then
    Set objIE = win
    Exit For
  End If
Next
Sleep 3000

Dim ieDoc As Object: Set ieDoc = ie.document 'HTMLコレクションを変数へ格納
'ieDoc.getElementsByName("pkgSpecNo")(0).value = PackingNo '検索ワードを入力

Do While objIE.Busy = True
  DoEvents
Loop


Set evt = objIE.document.createEvent("HTMLEvents")
evt.initEvent "change", True, False

ieDoc.getElementById("pkgSpecNo").Value = PackingNo '検索ワードを入力
ieDoc.getElementById("pkgSpecNo").dispatchEvent evt 'JavaScriptを強制発火_2022.03.11

ieDoc.getElementsByClassName("btn btn-info c-btn-size")(0).Click '検索をクリックする_btn btn-info

End Sub
Public Sub TextDataLoad(objIE As InternetExplorer, PackingNo As String)  '検索バーにテキストを入力する
Dim ieDoc As Object: Set ieDoc = ie.document 'HTMLコレクションを変数へ格納
ieDoc.getElementsByName("pkgSpecNo")(0).Value = PackingNo '検索ワードを入力
Sleep 1000
ieDoc.getElementsByClassName("btn btn-info")(0).Click '検索をクリックする
End Sub

Public Sub ieCheck(objIE As InternetExplorer) 'IEが完全表示されるまで待機
Dim timeout As Date
timeout = Now + TimeSerial(0, 0, 20)
Do While objIE.Busy = True Or objIE.readyState <> 4
  DoEvents
  Sleep 1
  If Now > timeout Then
    objIE.Refresh
    timeout = Now + TimeSerial(0, 0, 20)
  End If
Loop

timeout = Now + TimeSerial(0, 0, 20)

Do While objIE.document.readyState <> "complete"
  DoEvents
  Sleep 1
  If Now > timeout Then
    objIE.Refresh
    timeout = Now + TimeSerial(0, 0, 20)
  End If
Loop
End Sub

Function WaitIE(objIE As InternetExplorer) 'IEの読み込み待ち関数
Do While objIE.Busy = True Or objIE.readyState < 4
  DoEvents
Loop
End Function

'■Internet Explorerの読込待ち処理
Public Function Call_IE_WaitTime()
    '■objIEがBusy(処理中なら)DoEventsで待機
    Do While objIE.Busy = True
        DoEvents
    Loop
    '■objIEがREADYSTATE_COMPLETE(全データ読込完了になるまで)DoEventsで待機
    Do While objIE.readyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
End Function
'IEのプロセスを終了させる
Public Sub ie_kill()
If IsExecutingIE() Then
'    MsgBox "IE起動中"
  Set objShell = CreateObject("WScript.Shell")
  Set objExec = objShell.Exec("taskkill.exe /F /IM iexplore.exe")
Else
'    MsgBox "IEを起動していない"
End If
End Sub
Private Function IsExecutingIE() As Boolean
    Dim ws As Object
    For Each ws In CreateObject("Shell.Application").Windows
        If TypeName(ws.document) = "HTMLDocument" Then
            IsExecutingIE = True
            Exit For
        End If
    Next
End Function
