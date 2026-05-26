VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ファイル名変更タブを作成
Sub GetFileNamesBtn_Click()
  Dim r As Routes
  Set r = New Routes

  Dim suggestResult As ObjResult
  Set suggestResult = r.SuggestNewFileNames()

  If (Not suggestResult.isOK) Then
    Call r.ShowErrorMsg(suggestResult.err)
  End If
End Sub

' 最終的なファイル名にリネーム
Sub ChangeFileNamesBtn_Click()
  Dim r As Routes
  Set r = New Routes

  Dim renameResult As ObjResult
  Set renameResult = r.RenameFiles()

  If (Not renameResult.isOK) Then
    Call r.ShowErrorMsg(renameResult.err)
  End If
End Sub

' PDFの新旧名称を一致させる
Sub NoRename4PDF_Click()
  Dim r As Routes
  Set r = New Routes

  Dim noRenameRes As ObjResult
  Set noRenameRes = r.NoRenamePdf()

  If (Not noRenameRes.isOK) Then
    Call r.ShowErrorMsg(noRenameRes.err)
  End If
End Sub

' CADの新旧名称を一致させる
Sub NoRename4CAD_Click()
  Dim r As Routes
  Set r = New Routes
  
  Dim noRenameRes As ObjResult
  Set noRenameRes = r.NoRenameCad()
  
  If (Not noRenameRes.isOK) Then
    Call r.ShowErrorMsg(noRenameRes.err)
  End If
End Sub

' 入力内容のリセット
Private Sub ResetBtn_Click()
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  Call ResetData(sheets.RenamePage, "A" & INIT_ROWNUM_MOKUROKU & ":B10000")
End Sub