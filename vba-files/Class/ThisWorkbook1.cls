VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Const UpdateCheckTargetPath = "L:\○鋼構造設計部共通フォルダ\業務改善ＷＳ\2023年度\■リリース版"

Sub Workbook_Open()
  Call IsExecutable

  SplashScreen.Show

  If (CheckUpdate() = True) Then
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

Private Function CheckUpdate() As Boolean
  If (Dir(UpdateCheckTargetPath & "\" & ActiveWorkBook.Name) = "") Then
    CheckUpdate = True
  Else
    CheckUpdate = False
  End If
End Function

  Private Function IsExecutable() As Boolean
    If (ActiveWorkbook.path = UpdateCheckTargetPath) Then
      Dim res As Long
      res = MsgBox("「らくびむ」を配布元のフォルダで実行することはできません" & vbCrLf & _
                    "作業中のフォルダやサーバーにコピーしてご利用ください", vbCritical, "らくびむ - 実行できません -")
      
      ' 保存確認のメッセージを表示せずにExcelを閉じる
      Application.DisplayAlerts = False 
      Application.Quit
    End If
  End Function