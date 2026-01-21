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
  ' name2oldPは0番目がPDFのデータファイル説明書、1番目がDWGのデータファイル説明書となっている
  Dim name2oldP As Variant
  Dim name2ID As Object
  Dim oldP2newP As Object
  Dim systemIds As SystemIdObj
  Set systemIds = New SystemIdObj
  
  Dim contractID As String: contractID = systemIds.CompanyID & "-" & systemIds.ProjectID
  
  If checkFill4File() And checkFill4Mokuroku() Then
    ' 前のデータをリセット
    Call ResetData("データファイル説明書", "A2:F" & getTableCols() + 2)
    
    name2oldP = getRelationName2OldPath()
    Set name2ID = getRelationName2ID()
    Set oldP2newP = getRelationOld2NewPath()
    
    writeLinkTable contractID, systemIds.DesignID, name2oldP, name2ID, oldP2newP
  Else
    MsgBox _
      "目録入力シート または ファイル名変更シート に記入漏れが存在します" & vbCrLf & _
      "シート内にて色が塗られているセルにデータを入力してください"
  End If
End Sub

' 生成したデータファイル説明書を別のExcelに吐き出す
Private Sub OutputData_Click()
  ' ①変数定義
  Dim tmpRng As Range
  Dim basePath As String
  Dim systemIds As SystemIdObj
  Set systemIds = New SystemIdObj
  basePath = systemIds.TargetPath
  
  ' ②範囲指定
  sheets("データファイル説明書").Range("A:F").Select
  Set tmpRng = Selection
  
  ' ③ファイルの新規作成
  Dim book As Workbook
  Set book = Workbooks.Add
  
  ' ④選択範囲をコピー
  tmpRng.Copy
  
  ' ⑤選択範囲を貼り付け
  ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
  ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
  ActiveSheet.name = "データファイル説明書"
  
  ' ⑥名前を付けて保存
  book.SaveAs fileName:=basePath & "\CAD\" & systemIds.CompanyID & "-" & systemIds.ProjectID & ".xls", FileFormat:=xlWorkbookNormal
  
  ' ⑦ファイルを閉じる
  ActiveWorkbook.Close
  sheets("データファイル説明書").Range("A1").Select
End Sub

