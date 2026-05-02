Attribute VB_Name = "Logger"
Option Explicit

' Log モジュール: セッション内でログをバッファし、まとめて1ファイルに出力する
' 使用方法:
'   Call Log.Init()  ' 初期化（オプション）
'   Call Log.Info("メッセージ")
'   Call Log.Warn("警告メッセージ")
'   Call Log.Error("エラーメッセージ")
'   Call Log.Flush() ' ファイルに書き出す

Private Type LogEntry
    Level As String
    TimeStamp As String
    Message As String
End Type

Private logBuffer As Collection
Private logFilePath As String
Private logInitialized As Boolean

Public Sub Init()
    On Error Resume Next
    If logInitialized Then Exit Sub
    Set logBuffer = New Collection
    logFilePath = GetDefaultLogPath()
    logInitialized = True
End Sub

Public Sub Info(msg As String)
    AddEntry "INFO", msg
End Sub

Public Sub Warn(msg As String)
    AddEntry "WARN", msg
End Sub

Public Sub [Error](msg As String)
    AddEntry "ERROR", msg
End Sub

Private Sub AddEntry(level As String, msg As String)
    On Error Resume Next
    If Not logInitialized Then Call Init()
    Dim e As LogEntry
    e.Level = level
    e.TimeStamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
    e.Message = msg
    logBuffer.Add e
End Sub

Public Sub Flush()
    On Error GoTo ErrHandler
    If Not logInitialized Then Exit Sub
    If logBuffer Is Nothing Then Exit Sub

    Dim folder As String
    folder = GetLogFolder()
    If Dir(folder, vbDirectory) = "" Then MkDir folder

    Dim fnum As Integer: fnum = FreeFile
    Open logFilePath For Append As #fnum
    Dim i As Long
    For i = 1 To logBuffer.Count
        Dim e As LogEntry
        e = logBuffer(i)
        Print #fnum, e.TimeStamp & " [" & e.Level & "] " & e.Message
    Next i
    Close #fnum

    ' Clear buffer after flush
    Set logBuffer = New Collection

    ' Purge old logs asynchronously by attempting now (best-effort)
    Call PurgeOldLogs(folder)
    Exit Sub
ErrHandler:
    On Error Resume Next
    Close #1
End Sub

Private Function GetLogFolder() As String
    GetLogFolder = Environ$("HOMEDRIVE") & Environ$("HOMEPATH") & "\Rakubim"
End Function

Private Function GetDefaultLogPath() As String
    Dim folder As String: folder = GetLogFolder()
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    Dim fname As String
    fname = Format(Now, "yyyy-mm-dd_hhnnss") & ".log"
    GetDefaultLogPath = folder & "\" & fname
End Function

Private Sub PurgeOldLogs(folder As String)
    On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fld As Object, f As Object
    Set fld = fso.GetFolder(folder)
    For Each f In fld.Files
        If DateDiff("d", f.DateCreated, Now) > 30 Then
            f.Delete True
        End If
    Next f
    Exit Sub
ErrHandler:
    On Error Resume Next
End Sub
