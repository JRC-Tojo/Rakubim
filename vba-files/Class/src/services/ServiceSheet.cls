VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServiceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private inner_SheetRepo As RepoSheet

Private Const MatchingSheetName As String = "目録入力"
Private Const MatchingPaperNumberStartCell As String = "A2"
Private Const MatchingPaperNameStartCell As String = "B2"
Private Const MatchingPdfStartCell As String = "G2"
Private Const MatchingCadStartCell As String = "H2"

Private Sub class_initialize()
  Set inner_SheetRepo = New RepoSheet
End Sub

' 処理対象のWorkbookを変更する
Public Function SetWorkbook(wb As Workbook) As RepoSheet
  Call inner_SheetRepo.SetWorkbook(wb)
End Function

' 目録入力シートの図面番号列を読み取る
Public Function ReadPaperNumbersInMatching() As ObjResult
  Set ReadPaperNumbersInMatching = inner_SheetRepo.ReadColumn(MatchingSheetName, MatchingPaperNumberStartCell)
End Function

' 目録入力シートの図面名称列を読み取る
Public Function ReadPaperNamesInMatching() As ObjResult
  Set ReadPaperNamesInMatching = inner_SheetRepo.ReadColumn(MatchingSheetName, MatchingPaperNameStartCell)
End Function

' 与えられたPDFファイル名群とCADファイル名群を目録入力シートのPDF列とCAD列にそれぞれ書き込む
Public Function WriteMatchedFileNames(pdfNames As ObjList, cadNames As ObjList) As ObjResult
  ' シートに書き込み
  Dim wrotePdfRes As ObjResult
  Dim wroteCadRes As ObjResult
  Set wrotePdfRes = inner_SheetRepo.WriteColFromCell(MatchingSheetName, MatchingPdfStartCell, pdfNames)
  Set wroteCadRes = inner_SheetRepo.WriteColFromCell(MatchingSheetName, MatchingCadStartCell, cadNames)

  Set WriteMatchedFileNames = PickResultErr(Nothing, wrotePdfRes, wroteCadRes)
End Function

' 指定した列の値をファイル名として`basePath/readName`のlist[Path]を作成
Private Function ReadColAsPath(basePath As ObjPath, sheetName As String, startCell As String) As ObjResult
  Dim colItems As ObjResult
  Set colItems = inner_SheetRepo.ReadColumn(sheetName, startCell)
  If (Not colItems.isOK) Then
    Set ReadColAsPath = colItems
    Exit Function
  End If

  Dim fileName As Variant
  Dim paths As ObjList
  Set paths = New ObjList
  For Each fileName In colItems.value
    Call paths.Append(basePath.Child(fileName))
  Next

  Set ReadColAsPath = Success(paths)
End Function

' 目録入力シートのPDF列をファイルパスとして読み込む
Public Function ReadPdfColAsPath(basePath As ObjPath) As ObjResult
  Set ReadPdfColAsPath = ReadColAsPath(basePath, MatchingSheetName, MatchingPdfStartCell)
End Function

' 目録入力シートのCAD列をファイルパスとして読み込む
Public Function ReadCadColAsPath(basePath As ObjPath) As ObjResult
  Set ReadCadColAsPath = ReadColAsPath(basePath, MatchingSheetName, MatchingCadStartCell)
End Function

' 目録入力シートのPDF列とCAD列にプルダウンメニューを追加する
Public Function SetFilesPulldown(pdfFilePaths As ObjList, cadFilePaths As ObjList, errLang As String) As ObjResult
  Dim endRowNum As ObjResult
  Set endRowNum = inner_SheetRepo.GetLastRow(MatchingSheetName, MatchingPaperNameStartCell)
  If (Not endRowNum.isOK) Then
    Set SetFilesPulldown = endRowNum
  End If

  Dim pdfAddress As String
  pdfAddress = Range(MatchingPdfStartCell).Resize(endRowNum.value).Address
  Dim cadAddress As String
  cadAddress = Range(MatchingCadStartCell).Resize(endRowNum.value).Address

  Dim errMsgRepo As RepoErrMsgs
  Set errMsgRepo = New RepoErrMsgs
  Dim errTitle As String: errTitle = errMsgRepo.GetTitle("service.sheet.pulldown.files")
  Dim errMsg As String: errMsg = errMsgRepo.GetMsg("service.sheet.pulldown.files")
  
  Call inner_SheetRepo.Pulldown(MatchingSheetName, pdfAddress, pdfFilePaths, errTitle, errMsg)
  Call inner_SheetRepo.Pulldown(MatchingSheetName, cadAddress, cadFilePaths, errTitle, errMsg)

  Set SetFilesPulldown = Success()
End Function


' 目録入力シートの左側をすべて削除する
Public Function ClearMokuroku() As ObjResult


End Function


' 目録入力シートの右側をすべて削除する
Public Function ClearMatchedFileNames() As ObjResult


End Function