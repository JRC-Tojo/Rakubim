VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 目録入力シートにおいて実際のファイル名称を補完する
' 実際のファイル名を読み込み、目録の図面名称と対応させる
Sub ReadFileNamesBtn_Click()
  Dim r As Routes
  Set r = New Routes

  Dim matchedResult As ObjResult
  Set matchedResult = r.GetMatchingLists()

  ' 目録入力シートにプルダウンを設置
  Call r.SetPulldownForFiles()

  If (Not matchedResult.isOK) Then
    Call r.ShowErrorMsg(matchedResult.err)
  End If
End Sub


' 目録入力シートのファイル名称を削除する（左）
Sub ResetMokuroku_Left_Click()
  Dim r As Routes
  Set r = New Routes

  Dim resetResult As ObjResult
  Set resetResult = r.ResetMokuroku()

  If (Not resetResult.isOK) Then
    Call r.ShowErrorMsg(resetResult.err)
  End If
End Sub

' 目録入力シートのファイル名称を削除する（右）
Sub ResetMokuroku_Right_Click()
  Dim r As Routes
  Set r = New Routes

  Dim resetResult As ObjResult
  Set resetResult = r.ResetMatchedFiles()

  If (Not resetResult.isOK) Then
    Call r.ShowErrorMsg(resetResult.err)
  End If
End Sub