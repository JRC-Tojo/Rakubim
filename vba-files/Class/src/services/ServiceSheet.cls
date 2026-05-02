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
  Dim sheetRepo As RepoSheet
  Set sheetRepo = New RepoSheet

  Set ReadPaperNumbersInMatching = sheetRepo.ReadColumn(MatchingSheetName, MatchingPaperNumberStartCell)
End Function

' 目録入力シートの図面名称列を読み取る
Public Function ReadPaperNamesInMatching() As ObjResult
  Dim sheetRepo As RepoSheet
  Set sheetRepo = New RepoSheet

  Set ReadPaperNamesInMatching = sheetRepo.ReadColumn(MatchingSheetName, MatchingPaperNameStartCell)
End Function

' 与えられたPDFファイル名群とCADファイル名群を目録入力シートのPDF列とCAD列にそれぞれ書き込む
Public Function WriteMatchedFileNames(pdfNames As ObjList, cadNames As ObjList) As ObjResult
  Dim sheetRepo As RepoSheet
  Set sheetRepo = New RepoSheet

  ' シートに書き込み
  Dim wrotePdfRes As ObjResult
  Dim wroteCadRes As ObjResult
  Set wrotePdfRes = sheetRepo.WriteColFromCell(MatchingSheetName, MatchingPdfStartCell, pdfNames)
  Set wroteCadRes = sheetRepo.WriteColFromCell(MatchingSheetName, MatchingCadStartCell, cadNames)

  Set WriteMatchedFileNames = PickResultErr(Nothing, wrotePdfRes, wroteCadRes)
End Function


' 目録入力シートのPDF列とCAD列にプルダウンメニューを追加する
Public Function SetFilesPulldown(pdfFilePaths As ObjList, cadFilePaths As ObjList, errLang As String) As ObjResult
  Dim sheetRepo As RepoSheet
  Set sheetRepo = New RepoSheet

  Dim endRowNum As ObjResult
  Set endRowNum = sheetRepo.GetLastRow(MatchingSheetName, MatchingPaperNameStartCell)
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
  
  Call sheetRepo.Pulldown(MatchingSheetName, pdfAddress, pdfFilePaths, errTitle, errMsg)
  Call sheetRepo.Pulldown(MatchingSheetName, cadAddress, cadFilePaths, errTitle, errMsg)

  Set SetFilesPulldown = Success()
End Function


' 目録入力シートの左側をすべて削除する
Public Function ClearMokuroku() As ObjResult


End Function


' 目録入力シートの右側をすべて削除する
Public Function ClearMatchedFileNames() As ObjResult


End Function