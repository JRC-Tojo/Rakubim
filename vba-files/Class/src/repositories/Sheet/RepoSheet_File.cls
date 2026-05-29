VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepoSheet_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit


' 選択範囲を別ファイルにエクスポート
Public Function ExportToExcel(wb As Workbook, sheetName As String, targetPath As ObjPath, rangeAddress As String, newSheetName As String) As ObjResult
  On Error GoTo ErrHandler

  Dim targetNewSheetName As String
  targetNewSheetName = IIf(newSheetName = "", sheetName, newSheetName)

  ' ファイルの新規作成
  Dim book As Workbook
  Set book = Workbooks.Add

  ' 選択範囲をコピー
  wb.Sheets(sheetName).Range(rangeAddress).Copy

  ' 選択範囲を貼り付け
  book.Sheets(1).Range(rangeAddress).PasteSpecial Paste:=xlPasteColumnWidths
  book.Sheets(1).Range(rangeAddress).PasteSpecial Paste:=xlPasteAll
  book.Sheets(1).name = targetNewSheetName

  ' 名前を付けて保存
  book.SaveAs fileName:=targetPath.FullPath, FileFormat:=xlWorkbookNormal
  book.Close SaveChanges:=False

  Set ExportToExcel = Success()
ErrHandler:
  ' book.Close でErr.Number = 0が生じるため，ガードを入れる
  If (Err.Number = 0) Then
    Set ExportToExcel = Success()
  Else
    Set ExportToExcel = Failure("repo.sheet.export.error", targetPath.FullPath, Err.Number, Err.Description)
  End If
End Function


' 別ファイルから選択範囲をインポート
Public Function ImportFromExcel(wb As Workbook, sheetName As String, sourcePath As ObjPath, rangeAddress As String, sourceSheetName As String) As ObjResult
  On Error GoTo ErrHandler

  Dim targetSourceSheetName As String
  targetSourceSheetName = IIf(sourceSheetName = "", sheetName, sourceSheetName)

  ' ファイルを開く
  Dim book As Workbook
  Set book = Workbooks.Open(sourcePath.FullPath)

  ' 選択範囲をコピー
  book.Sheets(targetSourceSheetName).Range(rangeAddress).Copy

  ' 選択範囲を貼り付け
  wb.Sheets(sheetName).Range(rangeAddress).PasteSpecial Paste:=xlPasteColumnWidths
  wb.Sheets(sheetName).Range(rangeAddress).PasteSpecial Paste:=xlPasteAll

  ' ファイルを閉じる
  book.Close SaveChanges:=False

  Set ImportFromExcel = Success()
ErrHandler:
  ' book.Close でErr.Number = 0が生じるため，ガードを入れる
  If (Err.Number = 0) Then
    Set ImportFromExcel = Success()
  Else
    Set ImportFromExcel = Failure("repo.sheet.import.error", sourcePath.FullPath, Err.Number, Err.Description)
  End If
End Function