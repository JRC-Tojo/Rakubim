Attribute VB_Name = "FileProcess"
' ファイルに関連する処理()

' 目録入力シートにおけるデータが始まる行数
Public Const INIT_ROWNUM_MOKUROKU = 2

' 目録の記入データをリセット
Function ResetMokuroku(leftRight As String)
  If leftRight = "left" Then
    Call ResetData("目録入力", "A" & INIT_ROWNUM_MOKUROKU & ":E10000")
  Else
    Call ResetData("目録入力", "G" & INIT_ROWNUM_MOKUROKU & ":H10000")
  End If
End Function