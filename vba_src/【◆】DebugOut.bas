Attribute VB_Name = "【◆】DebugOut"
Option Explicit
' ========================================
' デバッグログ出力モジュール
' ========================================

' グローバル変数でログファイルパスを保持
Private g_logFilePath As String

' ログ出力用関数（引数追加版）
Public Sub DebugLog(ByVal message As String, Optional ByVal useWorkbookPath As Boolean = False)
    On Error Resume Next
    
    ' ログファイルパスが初期化されていない場合は初期化
    If g_logFilePath = "" Then
        Call InitializeDebugLog(useWorkbookPath)
    End If
    
    ' FileSystemObjectを使用してファイルに追記
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim logFile As Object
    
    If g_logFilePath <> "" Then
        If fso.fileExists(g_logFilePath) Then
            Set logFile = fso.OpenTextFile(g_logFilePath, 8) ' 8=追記モード
        Else
            ' ファイルが存在しない場合は作成
            fso.CreateTextFile g_logFilePath
            Set logFile = fso.OpenTextFile(g_logFilePath, 8)
        End If
        
        ' タイムスタンプ付きでログ出力
        logFile.WriteLine Format(Now, "yyyy/mm/dd hh:mm:ss") & " | " & message
        logFile.Close
        Set logFile = Nothing
    End If
    
    Set fso = Nothing
    
    ' イミディエイトウィンドウにも出力（開発時用）
    Debug.Print message
    
    On Error GoTo 0
End Sub

' ログファイルの初期化（引数追加版）
Public Sub InitializeDebugLog(Optional ByVal useWorkbookPath As Boolean = False)
    On Error GoTo ErrorHandler
    
    ' すでに初期化済みの場合はスキップ
    If g_logFilePath <> "" Then
        Exit Sub
    End If
    Dim tmpFaci As String: tmpFaci = GetSetting("UF検査判定システム", "起動設定", "設備名", "存在なし")
    Dim targetLineName As String: targetLineName = tmpFaci
    Dim basePath As String
    Dim logFolderPath As String
    Dim logFileName As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim logFile As Object
    
    ' ベースパスの決定
    basePath = ThisWorkbook.path & "\DebugPrint"
    
    ' ログフォルダパス
    logFolderPath = basePath & "\" & targetLineName

    
    ' DebugPrintフォルダの有無をチェック・作成
    If Not fso.FolderExists(basePath) Then
        fso.CreateFolder basePath
    End If
    
    ' ライン名フォルダの有無をチェック・作成
    If Not fso.FolderExists(logFolderPath) Then
        fso.CreateFolder logFolderPath
    End If
    
    ' ログファイル名（日時付き）
    logFileName = "DebugLog_" & Format(Now, "yyyymmdd_hhmmss") & ".txt"
    g_logFilePath = logFolderPath & "\" & logFileName
    
    ' ヘッダー出力
    fso.CreateTextFile g_logFilePath
    Set logFile = fso.OpenTextFile(g_logFilePath, 8) ' 8=追記モード
    logFile.WriteLine "========================================="
    logFile.WriteLine "デバッグログ開始"
    logFile.WriteLine "実行日時: " & Format(Now, "yyyy/mm/dd hh:mm:ss")
    If useWorkbookPath Then
        logFile.WriteLine "ログパス: ThisWorkbook.Path"
        logFile.WriteLine "対象: MaterialForm"
    Else
        logFile.WriteLine "ライン名: " & targetLineName
    End If
    logFile.WriteLine "ログファイル: " & g_logFilePath
    logFile.WriteLine "========================================="
    logFile.WriteLine ""
    logFile.Close
    
    Set logFile = Nothing
    Set fso = Nothing
    
    Exit Sub
    
ErrorHandler:
    ' エラー時はイミディエイトウィンドウに出力
    Debug.Print "ログファイル初期化エラー: " & Err.Description
    g_logFilePath = ""
End Sub

' ログファイルを閉じる
Public Sub CloseDebugLog()
    On Error Resume Next
    
    If g_logFilePath = "" Then Exit Sub
    
    ' FileSystemObjectを使用して終了メッセージを追記
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim logFile As Object
    
    If fso.fileExists(g_logFilePath) Then
        Set logFile = fso.OpenTextFile(g_logFilePath, 8) ' 8=追記モード
        logFile.WriteLine ""
        logFile.WriteLine "========================================="
        logFile.WriteLine "デバッグログ終了"
        logFile.WriteLine "終了日時: " & Format(Now, "yyyy/mm/dd hh:mm:ss")
        logFile.WriteLine "========================================="
        logFile.Close
        Set logFile = Nothing
    End If
    
    Set fso = Nothing
    
    ' グローバル変数をリセット
    g_logFilePath = ""
    
    On Error GoTo 0
End Sub

' ログファイルパスをリセット（新しいログファイルを開始する場合）
Public Sub ResetDebugLog()
    g_logFilePath = ""
End Sub
