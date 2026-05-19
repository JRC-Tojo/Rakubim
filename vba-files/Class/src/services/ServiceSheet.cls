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

Private Const RenameSheetName As String = "ファイル名変更"
Private Const RenameOldNamesStartCell As String = "A2"
Private Const RenameNewNamesStartCell As String = "B2"

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
Private Function ReadColAsPath(basePath As ObjPath, targetExts As ModelExts, sheetName As String, startCell As String) As ObjResult
  Dim colItems As ObjResult
  Set colItems = inner_SheetRepo.ReadColumn(sheetName, startCell)
  If (Not colItems.isOK) Then
    Set ReadColAsPath = colItems
    Exit Function
  End If

  Dim fileName As Variant
  Dim paths As ObjList
  Set paths = New ObjList
  Dim tmpPath As ObjPath
  For Each fileName In colItems.value
    Set tmpPath = basePath.Child(fileName)
    ' TODO: 勝手に処理対象のファイルを落とすため，数のずれによるバグが生じる可能性あり
    ' エラーとして返すようにすべきか？
    If targetExts.ExtSet.Contains(tmpPath.Extension) Then
      Call paths.Append(tmpPath)
    End If
  Next

  Set ReadColAsPath = Success(paths)
End Function

' 目録入力シートのPDF列をファイルパスとして読み込む
Public Function ReadPdfColAsPath(basePath As ObjPath, targetExts As ModelExts) As ObjResult
  Set ReadPdfColAsPath = ReadColAsPath(basePath, targetExts, MatchingSheetName, MatchingPdfStartCell)
End Function

' 目録入力シートのCAD列をファイルパスとして読み込む
Public Function ReadCadColAsPath(basePath As ObjPath, targetExts As ModelExts) As ObjResult
  Set ReadCadColAsPath = ReadColAsPath(basePath, targetExts, MatchingSheetName, MatchingCadStartCell)
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


' ファイル名変更シートにおいて，各ファイルの変更先名称の提案をシートに書き込む
Public Function WriteSuggestRenames(pdfOldToNewDict As ObjDict, cadOldToNewDict As ObjDict) As ObjResult
  Dim oldNames As ObjList
  Dim newNames As ObjList
  Set oldNames = pdfOldToNewDict.Keys().Combine(cadOldToNewDict.Keys())
  Set newNames = pdfOldToNewDict.Values().Combine(cadOldToNewDict.Values())

  ' シートに書き込み
  Dim wroteOldRes As ObjResult
  Dim wroteNewRes As ObjResult
  Set wroteOldRes = inner_SheetRepo.WriteColFromCell(RenameSheetName, RenameOldNamesStartCell, oldNames)
  Set wroteNewRes = inner_SheetRepo.WriteColFromCell(RenameSheetName, RenameNewNamesStartCell, newNames)

  Set WriteSuggestRenames = PickResultErr(Nothing, wroteOldRes, wroteNewRes)
End Function


' 名前変更シートの指定したセルの列に記載された各ファイル名をファイルパス一覧として返す
Private Function readRenameSheetFilePaths(pdfBasePath As ObjPath, pdfExts As ModelExts, cadBasePath As ObjPath, cadExts As ModelExts, startCell As String) As ObjResult
  Dim pdfPaths As ObjResult
  Dim cadPaths As ObjResult

  Set pdfPaths = ReadColAsPath(pdfBasePath, pdfExts, RenameSheetName, startCell)
  Set cadPaths = ReadColAsPath(cadBasePath, cadExts, RenameSheetName, startCell)
  If (Not IsResultAllOK(pdfPaths, cadPaths)) Then
    Set readRenameSheetFilePaths = PickResultErr(Nothing, pdfPaths, cadPaths)
    Exit Function
  End If

  Set readRenameSheetFilePaths = Success(pdfPaths.value.Combine(cadPaths.value))
End Function


' 名前変更シートにおける旧ファイルパス一覧を返す
Public Function ReadOldFilePaths(pdfBasePath As ObjPath, pdfExts As ModelExts, cadBasePath As ObjPath, cadExts As ModelExts) As ObjResult
  Set ReadOldFilePaths = readRenameSheetFilePaths(pdfBasePath, pdfExts, cadBasePath, cadExts, RenameOldNamesStartCell)
End Function


' 名前変更シートにおける新ファイルパス一覧を返す
Public Function ReadNewFilePaths(pdfBasePath As ObjPath, pdfExts As ModelExts, cadBasePath As ObjPath, cadExts As ModelExts) As ObjResult
  Dim newPaths As ObjResult
  Set newPaths = readRenameSheetFilePaths(pdfBasePath, pdfExts, cadBasePath, cadExts, RenameNewNamesStartCell)
  If (Not newPaths.isOK) Then
    Set ReadNewFilePaths = newPaths
    Exit Function
  End If

  ' 新ファイル名が重複していないか確認
  Dim newPath As ObjPath
  Dim newPathSet As ObjSet
  Set newPathSet = New ObjSet
  For Each newPath In newPaths.value
    If newPathSet.Contains(newPath.FullPath) Then
      Set ReadNewFilePaths = Failure("service.sheet.newname.duplicated", newPath.Name)
      Exit Function
    Else
      Call newPathSet.Append(newPath.FullPath)
    End If
  Next newPath

  Set ReadNewFilePaths = newPaths
End Function