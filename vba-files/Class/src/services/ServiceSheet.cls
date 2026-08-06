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

Private Const FirstInputSheetName As String = "入力"
Private Const FirstInputCodeCell As String = "C6"

Private Const CodeSheetName As String = "業務機関コード"
Private Const CodeShowingStartCell As String = "H3"

Private Const MatchingSheetName As String = "目録入力"
Private Const MatchingPaperNumberStartCell As String = "A2"
Private Const MatchingPaperNameStartCell As String = "B2"
Private Const MatchingPdfStartCell As String = "G2"
Private Const MatchingCadStartCell As String = "H2"

Private Const RenameSheetName As String = "ファイル名変更"
Private Const RenameOldNamesStartCell As String = "A2"
Private Const RenameNewNamesStartCell As String = "B2"
Private Const RenameNoRenamePdfMsgCell As String = "G6"
Private Const RenameNoRenameCadMsgCell As String = "G8"

Private Const RelationalTableSheetName As String = "データファイル説明書"
Private Const RelationalKeiyakuStartCell As String = "A2"
Private Const RelationalPaperNumberStartCell As String = "B2"
Private Const RelationalDesignNumberStartCell As String = "C2"
Private Const RelationalPaperNameStartCell As String = "D2"
Private Const RelationalFileNameStartCell As String = "E2"
Private Const RelationalGridStartCell As String = "A1"
Private Const RelationalGridColCount As Long = 6

Private Sub class_initialize()
  Set inner_SheetRepo = New RepoSheet
End Sub

' 処理対象のWorkbookを変更する
Public Function SetWorkbook(wb As Workbook) As RepoSheet
  Call inner_SheetRepo.SetWorkbook(wb)
End Function

' 外部からシートをインポートする
Public Function ImportCodeSheet() As ObjResult
  Dim secretRepo As RepoSecret
  Set secretRepo = New RepoSecret
  Dim targetPath As ObjPath
  Set targetPath = New ObjPath
  Dim targetSheetName As String
  Dim decryptRes As ObjResult
  Dim existsSheet As ObjResult
  
  Const targetEncryptedPath = "mTJLK+YoxA9+x/e4foAR+cxDzb+sJVx/z4R0k6O1n9D/cxlqLLM7l7Mr3x+I/5U+WqQj/DfBQ2Um3q6Zvoo50qLqrPqrw0SF902hbPeMNCS13RBhyeLi71Y2HzfgWFZUS/XsQECS4xzhxcgmGOAYoh8nQOtRFtz43EUM4y9DBVqvvbAjGH7mBHj8sIKTVXwLX2vrHF8ozE+3epABMd8mzkTx+HZlTNFvwMBEgPfHLyosCYZ9jASZcw=="
  Set decryptRes = secretRepo.Decrypt(targetEncryptedPath)
  If (Not decryptRes.isOK) Then
    Set ImportCodeSheet = decryptRes
    Exit Function
  End If

  Call targetPath.Init(decryptRes.value)
  targetSheetName = targetPath.Stem

  Set existsSheet = inner_SheetRepo.IsExistsSheet(targetSheetName)
  If (Not existsSheet.isOK) Then
    Set ImportCodeSheet = existsSheet
    Exit Function
  End If
  
  If (existsSheet.value) Then
    Set ImportCodeSheet = Success()
  Else
    Set ImportCodeSheet = inner_SheetRepo.CopySheetFromOuter(targetPath, targetSheetName, visible:=False)
  End If
End Function

' コード一覧を入力シートにプルダウンで登録する
Public Function SetCodesPulldown() As ObjResult
  ' コード一覧の登録処理
  Dim codeNameList As ObjResult
  Set codeNameList = inner_SheetRepo.ReadColumn(CodeSheetName, CodeShowingStartCell)
  If (Not codeNameList.isOK) Then
    Set SetCodesPulldown = codeNameList
    Exit Function
  End If
  
  ' プルダウンを設置
  Dim errMsgRepo As RepoErrMsgs
  Set errMsgRepo = New RepoErrMsgs
  Dim errTitle As String: errTitle = errMsgRepo.GetTitle("service.sheet.pulldown.codes")
  Dim errMsg As String: errMsg = errMsgRepo.GetMsg("service.sheet.pulldown.codes")
  Call inner_SheetRepo.Pulldown(FirstInputSheetName, FirstInputCodeCell, codeNameList.value, errTitle, errMsg)

  ' 初期値を設定
  Call inner_SheetRepo.WriteCell(FirstInputSheetName, FirstInputCodeCell, codeNameList.value.At(-1))

  Set SetCodesPulldown = Success() 
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


' パス一覧のリストをファイル名のリストに変換する
Private Function ConvertPathToFileNames(paths As ObjList) As ObjList
  Dim fileNames As ObjList
  Set fileNames = New ObjList
  
  Dim path As ObjPath
  For Each path In paths
    Call fileNames.Append(path.Name)
  Next

  Set ConvertPathToFileNames = fileNames
End Function

' 指定したセルから行番号までの範囲におけるRangeオブジェクトを返す
Private Function RangeToRowIndex(startCell As String, rowIndex As Long) As Range
  Dim startRange As Range
  Set startRange = Range(startCell)
  Set RangeToRowIndex = Range(startRange, Cells(rowIndex, startRange.Column))
End Function

' 目録入力シートのPDF列とCAD列にプルダウンメニューを追加する
Public Function SetFilesPulldown(pdfFilePaths As ObjList, cadFilePaths As ObjList, errLang As String) As ObjResult
  Dim endRowNum As ObjResult
  Set endRowNum = inner_SheetRepo.GetLastRow(MatchingSheetName, MatchingPaperNameStartCell)
  If (Not endRowNum.isOK) Then
    Set SetFilesPulldown = endRowNum
    Exit Function
  End If

  Dim pdfAddress As String
  pdfAddress = RangeToRowIndex(MatchingPdfStartCell, endRowNum.value).Address
  Dim cadAddress As String
  cadAddress = RangeToRowIndex(MatchingCadStartCell, endRowNum.value).Address

  Dim pdfFileNames As ObjList
  Set pdfFileNames = ConvertPathToFileNames(pdfFilePaths)
  Dim cadFileNames As ObjList
  Set cadFileNames = ConvertPathToFileNames(cadFilePaths)

  Dim clearPulldownAddress As String
  clearPulldownAddress = Union(Range(MatchingPdfStartCell), Range(MatchingCadStartCell)).EntireColumn.Address
  Call inner_SheetRepo.ClearValidation(MatchingSheetName, clearPulldownAddress)

  Dim errMsgRepo As RepoErrMsgs
  Set errMsgRepo = New RepoErrMsgs
  Dim errTitle As String: errTitle = errMsgRepo.GetTitle("service.sheet.pulldown.files")
  Dim errMsg As String: errMsg = errMsgRepo.GetMsg("service.sheet.pulldown.files")
  
  Call inner_SheetRepo.Pulldown(MatchingSheetName, pdfAddress, pdfFileNames, errTitle, errMsg)
  Call inner_SheetRepo.Pulldown(MatchingSheetName, cadAddress, cadFileNames, errTitle, errMsg)

  Set SetFilesPulldown = Success()
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


' 名前変更シートにおける旧ファイル名と新ファイル名の対応関係を表す辞書を返す
Public Function ReadOldToNewFileNameDict(pdfBasePath As ObjPath, pdfExts As ModelExts, cadBasePath As ObjPath, cadExts As ModelExts) As ObjResult
  Dim oldPaths As ObjResult
  Dim newPaths As ObjResult
  Set oldPaths = ReadOldFilePaths(pdfBasePath, pdfExts, cadBasePath, cadExts)
  Set newPaths = ReadNewFilePaths(pdfBasePath, pdfExts, cadBasePath, cadExts)
  If (Not IsResultAllOK(oldPaths, newPaths)) Then
    Set ReadOldToNewFileNameDict = PickResultErr(Nothing, oldPaths, newPaths)
    Exit Function
  End If

  If (oldPaths.value.Count <> newPaths.value.Count) Then
    Set ReadOldToNewFileNameDict = Failure("service.sheet.old2new.missmatch")
    Exit Function
  End If

  Dim oldNew As ObjDict
  Set oldNew = New ObjDict
  Dim i As Long
  For i = 0 To oldPaths.value.Count - 1
    Call oldNew.Append(oldPaths.value.At(i).Name, newPaths.value.At(i).Name)
  Next i

  Set ReadOldToNewFileNameDict = Success(oldNew)
End Function


' 指定した拡張子の新ファイル名称を旧ファイル名称に戻す
Private Function RevertNewFileNames(targetExts As ModelExts) As ObjResult
  Dim oldFileNames As ObjResult
  Dim newFileNames As ObjResult
  Set newFileNames = inner_SheetRepo.ReadColumn(RenameSheetName, RenameNewNamesStartCell)
  Set oldFileNames = inner_SheetRepo.ReadColumn(RenameSheetName, RenameOldNamesStartCell)
  If (Not IsResultAllOK(newFileNames, oldFileNames)) Then
    Set RevertNewFileNames = PickResultErr(Nothing, newFileNames, oldFileNames)
    Exit Function
  End If

  ' 新ファイル名を旧ファイル名と同じにする
  Dim i As Long
  Dim newFileName As String
  Dim newFilePath As ObjPath
  Dim oldFileName As String
  For i = 0 To newFileNames.value.Count - 1
    oldFileName = oldFileNames.value.At(i)
    newFileName = newFileNames.value.At(i)
    Set newFilePath = New ObjPath
    Call newFilePath.Init(newFileName)
    If (targetExts.ExtSet.Contains(newFilePath.Extension)) Then
      Call newFileNames.value.Update(i, oldFileName)
    End If
  Next i

  ' 変更済み新ファイル名一覧をシートに書き込む
  Dim wroteRes As ObjResult
  Set wroteRes = inner_SheetRepo.WriteColFromCell(RenameSheetName, RenameNewNamesStartCell, newFileNames.value)
  If (Not wroteRes.isOK) Then
    Set RevertNewFileNames = wroteRes
    Exit Function
  End If

  Set RevertNewFileNames = Success()
End Function


' 名前変更シートの新旧ファイル名称を同じにする標準処理
Private Function ShowNoRenameMsg(targetExts As ModelExts, msgCell As String) As ObjResult
  Dim revertRes As ObjResult
  Set revertRes = RevertNewFileNames(targetExts)
  If (Not revertRes.isOK) Then
    Set ShowNoRenameMsg = revertRes
    Exit Function
  End If

  Set ShowNoRenameMsg = inner_SheetRepo.RegistCellMsg(RenameSheetName, msgCell, "設定OK")
End Function


' ファイル名変更シートにおいて，新PDFファイル名を旧ファイル名と同じにする
Public Function ShowNoRename4PDFMsg(pdfExts As ModelExts) As ObjResult
  Set ShowNoRename4PDFMsg = ShowNoRenameMsg(pdfExts, RenameNoRenamePdfMsgCell)
End Function


' ファイル名変更シートにおいて，新PDFファイル名を旧ファイル名と同じにする
Public Function ShowNoRename4CADMsg(cadExts As ModelExts) As ObjResult
  Set ShowNoRename4CADMsg = ShowNoRenameMsg(cadExts, RenameNoRenameCadMsgCell)
End Function


' データファイル説明書を書き込む
' 引数: `keiyakuId`: 契約番号
'       `designNum`: 設計番号
'       `papers`[List[ModelPaper]]: 目録入力シートから読み取った図面情報のリスト
'       `oldPdfPaths`[List[ObjPath]]: 名前変更シートから読み取った旧PDFファイルパスのリスト
'       `oldCadPaths`[List[ObjPath]]: 名前変更シートから読み取った旧CADファイルパスのリスト
'       `old2newPaths`[Dict[旧ファイルパス, 新ファイルパス]]: 名前変更シートから読み取った旧ファイルパスと新ファイルパスの対応関係を表す辞書
Public Function WriteRelationalTable(keiyakuId As String, designNum As String, papers As ObjList, oldPdfPaths As ObjList, oldCadPaths As ObjList, old2newPaths As ObjDict) As ObjResult
  If (papers.Count <> oldPdfPaths.Count Or papers.Count <> oldCadPaths.Count) Then
    Set WriteRelationalTable = Failure("service.sheet.relationaltable.missmatch")
    Exit Function
  End If

  Dim keiyakuIds As ObjList: Set keiyakuIds = New ObjList
  Dim paperIdxs As ObjList : Set paperIdxs = New ObjList
  Dim designNums As ObjList: Set designNums = New ObjList
  Dim paperNames As ObjList: Set paperNames = New ObjList
  Dim fileNames As ObjList : Set fileNames = New ObjList

  Dim i As Long
  Dim pdfFileName As String
  Dim cadFileName As String
  For i = 0 To papers.Count - 1
    pdfFileName = old2newPaths.At(oldPdfPaths.At(i).Name, defValue:="")
    cadFileName = old2newPaths.At(oldCadPaths.At(i).Name, defValue:="")
    If (pdfFileName = "") Then
      Set WriteRelationalTable = Failure("service.sheet.relationaltable.lostrename", oldPdfPaths.At(i).Name)
      Exit Function
    End If
    If (cadFileName = "") Then
      Set WriteRelationalTable = Failure("service.sheet.relationaltable.lostrename", oldCadPaths.At(i).Name)
      Exit Function
    End If

    ' PDF
    Call keiyakuIds.Append(keiyakuId)
    Call paperIdxs.Append(papers.At(i).PaperIndex)
    Call designNums.Append(designNum)
    Call paperNames.Append(papers.At(i).SourceName)
    Call fileNames.Append(pdfFileName)

    ' CAD
    Call keiyakuIds.Append(keiyakuId)
    Call paperIdxs.Append(papers.At(i).PaperIndex)
    Call designNums.Append(designNum)
    Call paperNames.Append(papers.At(i).SourceName)
    Call fileNames.Append(cadFileName)
  Next i

  ' シートに書き込み
  Dim wroteKeiyakuRes As ObjResult
  Dim wrotePaperIdxRes As ObjResult
  Dim wroteDesignNumRes As ObjResult
  Dim wrotePaperNameRes As ObjResult
  Dim wroteFileNameRes As ObjResult
  Dim wroteRes As ObjResult
  Set wroteKeiyakuRes = inner_SheetRepo.WriteColFromCell(RelationalTableSheetName, RelationalKeiyakuStartCell, keiyakuIds)
  Set wrotePaperIdxRes = inner_SheetRepo.WriteColFromCell(RelationalTableSheetName, RelationalPaperNumberStartCell, paperIdxs)
  Set wroteDesignNumRes = inner_SheetRepo.WriteColFromCell(RelationalTableSheetName, RelationalDesignNumberStartCell, designNums)
  Set wrotePaperNameRes = inner_SheetRepo.WriteColFromCell(RelationalTableSheetName, RelationalPaperNameStartCell, paperNames)
  Set wroteFileNameRes = inner_SheetRepo.WriteColFromCell(RelationalTableSheetName, RelationalFileNameStartCell, fileNames)
  Set wroteRes = PickResultErr(Nothing, wroteKeiyakuRes, wrotePaperIdxRes, wroteDesignNumRes, wrotePaperNameRes, wroteFileNameRes)
  If (Not wroteRes.isOK) Then
    Set WriteRelationalTable = wroteRes
    Exit Function
  End If

  ' 罫線を描画
  Dim endCell As String
  endCell = Range(RelationalGridStartCell).Offset(papers.Count * 2, RelationalGridColCount - 1).Address
  Call inner_SheetRepo.MakeGrid(RelationalTableSheetName, RelationalGridStartCell, endCell)

  ' 印刷範囲を設定
  Dim printAreaRes As ObjResult
  Set printAreaRes = inner_SheetRepo.RegistPrintArea(RelationalTableSheetName, RelationalGridStartCell, endCell)
  If (Not printAreaRes.isOK) Then
    Set WriteRelationalTable = printAreaRes
    Exit Function
  End If
  
  Set WriteRelationalTable = Success()
End Function


' データファイル説明書をエクスポートする
Public Function ExportRelationalTable(outputPath As ObjPath) As ObjResult
  ' データファイル説明書の最終行を取得
  Dim lastCellRow As ObjResult
  Set lastCellRow = inner_SheetRepo.GetLastRow(RelationalTableSheetName, RelationalGridStartCell)
  If (Not lastCellRow.isOK) Then
    Set ExportRelationalTable = lastCellRow
    Exit Function
  End If

  ' 出力範囲を確定
  Dim tableRangeAddress As String
  tableRangeAddress = Range(RelationalGridStartCell).Resize(lastCellRow.Value + 1, RelationalGridColCount).Address

  ' 外部Excelに出力
  Dim exportRes As ObjResult
  Set exportRes = inner_SheetRepo.ExportToExcel(RelationalTableSheetName, outputPath, tableRangeAddress)
  If Not exportRes.isOK Then
    Set ExportRelationalTable = exportRes
    Exit Function
  End If

  Set ExportRelationalTable = Success()
End Function


' 指定したシートのセル列のデータをリセットする
Private Function ResetCols(sheetName As String, ParamArray startCells() As Variant) As ObjResult
  Dim res As ObjResult
  Dim startCell As Variant
  For Each startCell In startCells
    Set res = inner_SheetRepo.ClearColFromCell(sheetName, startCell)
    If (Not res.isOK) Then
      Set ResetCols = res
      Exit Function
    End If
  Next startCell

  Set ResetCols = Success()
End Function


' 目録入力シートの図面目録部分をリセットする
Public Function ResetMokuroku() As ObjResult
  Set ResetMokuroku = ResetCols(MatchingSheetName, MatchingPaperNumberStartCell, MatchingPaperNameStartCell)
End Function


' 目録入力シートのファイル名記入部分をリセットする
Public Function ResetMatchedFileNames() As ObjResult
  Set ResetMatchedFileNames = ResetCols(MatchingSheetName, MatchingPdfStartCell, MatchingCadStartCell)
End Function


' ファイル名変更シートをリセットする
Public Function ResetRenameSheet() As ObjResult
  Set ResetRenameSheet = ResetCols(RenameSheetName, RenameOldNamesStartCell, RenameNewNamesStartCell)
End Function


' データファイル説明書をリセットする
Public Function ResetRelationalTable() As ObjResult
  Call inner_SheetRepo.ClearGrid(RelationalTableSheetName)
  Call inner_SheetRepo.ClearPrintArea(RelationalTableSheetName)
  Set ResetRelationalTable = ResetCols( _
    RelationalTableSheetName, _
    RelationalKeiyakuStartCell, _
    RelationalPaperNumberStartCell, _
    RelationalDesignNumberStartCell, _
    RelationalPaperNameStartCell, _
    RelationalFileNameStartCell _
  )
End Function