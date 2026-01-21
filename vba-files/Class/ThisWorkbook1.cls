VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Const UpdateCheckTargetPath = "\\sv01f02\本社_技術本部共通\◆01　照査室\05_設計ツール\■09_成果品ファイル一括整理ツール"

Sub Workbook_Open()
  Call IsExecutable
  Call SetPulldown

  SplashScreen.Show

  If (RequireUpdate() = True) Then
    Dim res As Long
    res = MsgBox("「らくびむ」の最新版がリリースされました" & vbCrLf & _ 
            "以下のフォルダから最新版をコピーしてご利用ください" & vbCrLf & vbCrLf & _ 
            UpdateCheckTargetPath & vbCrLf & vbCrLf & _ 
            "上記のフォルダを開きますか？",vbInformation + vbYesNo + vbDefaultButton2, "らくびむ - アップデート -")
    If (res = vbYes) Then
      Shell "C:\windows\explorer.exe " & UpdateCheckTargetPath & "\", vbNormalFocus
    End If
  End If
End Sub

Private Sub SetPulldown()
  Dim matcher As Paper2FileMatcher
  Set matcher = GetMatcher()
  If (matcher Is Nothing) Then Exit Sub

  Dim rowNum As Integer
  rowNum = getLastPaperRows()

  Call SetFilePulldownMenu(matcher, rowNum - LIMITED_SPACE_COUNTER)
End Sub

Private Function RequireUpdate() As Boolean
  Dim objFSO As FileSystemObject
  Set objFSO = New FileSystemObject
  ' アップデートの保存先が見つからない場合はアップデートチェックを行わない
  If Not objFSO.FolderExists(UpdateCheckTargetPath) Then
    RequireUpdate = False
    Exit Function
  End If

  ' アップデートの保存先に自分自身と同じ名前のファイルがないときには，新しいバージョンが来たと判断する
  If Not objFSO.FileExists(UpdateCheckTargetPath & "\" & ActiveWorkbook.name) Then
    RequireUpdate = True
  Else
    RequireUpdate = False
  End If
End Function

Private Function IsExecutable()
  ' フォルダの開き方によって「M:/」だったり「//sv01f02/」になったりするため、パスの末尾３階層分のみをチェックする
  Const checkLetterCount = 35
  If (Right(ActiveWorkbook.path, checkLetterCount) = Right(UpdateCheckTargetPath, checkLetterCount)) Then
    Dim res As Long
    res = MsgBox("「らくびむ」を配布元のフォルダで実行することはできません" & vbCrLf & _
                  "作業中のフォルダやサーバーにコピーしてご利用ください", vbCritical, "らくびむ - 実行できません -")
    
    ' 保存確認のメッセージを表示せずにExcelを閉じる
    Application.DisplayAlerts = False 
    Application.Quit
  End If
End Function