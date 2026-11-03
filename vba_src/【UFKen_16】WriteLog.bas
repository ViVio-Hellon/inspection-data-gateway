Attribute VB_Name = "【UFKen_16】WriteLog"
Option Explicit
Public Const AccessLogTitle As String = "管理ﾌｧｲﾙ使用状況"
'Public Const AccessLogFileName As String = "Access.log"
' 標準はｶﾝﾏｾﾊﾟﾚｰﾄ(CSV)でﾛｸﾞ作成（超高速化版）
Public Sub WriteLog(msg1 As String, Optional Sepl As String = ",", Optional TargetPath As String = "", Optional TargetFile As String = "")
    Dim AccessLogFileName As String: AccessLogFileName = "Access.log"
    If TargetFile <> "" Then AccessLogFileName = TargetFile

    On Error Resume Next

    Dim SaveFile As String
    If TargetPath = "" Then
        SaveFile = ThisWorkbook.path & "\" & AccessLogFileName
    Else
        Call CreateFolder(TargetPath) ' ﾌｫﾙﾀﾞを作成
        SaveFile = TargetPath & "\" & AccessLogFileName
    End If

    ' 【改善点①】FSOオブジェクトをStaticで保持（毎回CreateObjectしない）
    Static fso As Object
    If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")

    Dim LOG As Object
    If fso.fileExists(SaveFile) Then
        Set LOG = fso.OpenTextFile(SaveFile, 8) ' 追記ﾓｰﾄﾞ
    Else
        fso.CreateTextFile SaveFile ' ログファイルがなければ作る
        Set LOG = fso.OpenTextFile(SaveFile, 8) ' 追記ﾓｰﾄﾞ
        LOG.WriteLine AccessLogTitle
    End If

    ' 【改善点②】WMIやWshNetworkを使わず、Windowsの環境変数からPC名を一瞬で取得
    Dim compName As String: compName = Environ("COMPUTERNAME")

    ' ログの書き込み（不要なスペースやOsInfoをトリム）
    LOG.WriteLine Format(Now, "YYYY/MM/DD hh:mm:ss") & " " & Sepl & compName & " " & Sepl & " " & msg1 & " "

    Set LOG = Nothing ' テキストファイルオブジェクトのみ解放（fsoは保持）
End Sub
