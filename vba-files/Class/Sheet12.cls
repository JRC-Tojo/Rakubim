VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' 最終的な成果物であるデータファイル説明書の作成
Sub GenerateTableBtn_Click()
  Dim r As Routes
  Set r = New Routes

  Dim genTableResult As ObjResult
  Set genTableResult = r.GenerateRelationalTable()

  If (Not genTableResult.isOK) Then
    Call r.ShowErrorMsg(genTableResult.err)
  End If
End Sub


' 生成したデータファイル説明書を別のExcelに吐き出す
Private Sub OutputData_Click()
  Dim r As Routes
  Set r = New Routes

  Dim exportResult As ObjResult
  Set exportResult = r.ExportRelationalTable()

  If (Not exportResult.isOK) Then
    Call r.ShowErrorMsg(exportResult.err)
  End If
End Sub


' 入力内容のリセット
Private Sub ResetBtn_Click()
  Dim r As Routes
  Set r = New Routes
  
  Dim resetRes As ObjResult
  Set resetRes = r.ResetRelationalTable()
  
  If (Not resetRes.isOK) Then
    Call r.ShowErrorMsg(resetRes.err)
    Exit Sub
  End If
End Sub

