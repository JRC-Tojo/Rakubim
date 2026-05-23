VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControllerSheet10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function SuggestNewName() As ObjResult
  Dim matchingController As ControllerMatching
  Set matchingController = New ControllerMatching
  Dim fileService As ServiceFile
  Set fileService = New ServiceFile
  Dim sheetService As ServiceSheet
  Set sheetService = New ServiceSheet
  Dim config As ServiceSystemConfig
  Set config = New ServiceSystemConfig

  ' システム設定の読み込み
  Dim resConfig As ObjResult
  Set resConfig = config.LoadUserInput()
  If Not resConfig.isOK Then
    Set SuggestNewName = resConfig
    Exit Function
  End If

  ' 図面情報を取得
  Dim papers As ObjResult
  Set papers = matchingController.LoadPapers(sheetService)
  If Not papers.isOK Then
    Set SuggestNewName = papers
    Exit Function
  End If

  ' 入力済みのファイル群を読み込み
  Dim loadedPdfPaths As ObjResult
  Dim loadedCadPaths As ObjResult
  Set loadedPdfPaths = sheetService.ReadPdfColAsPath(config.PdfFolderPath, config.PdfExtensions)
  Set loadedCadPaths = sheetService.ReadCadColAsPath(config.CadFolderPath, config.CadExtensions)
  If (Not IsResultAllOK(loadedPdfPaths, loadedCadPaths)) Then
    Set SuggestNewName = PickResultErr(Nothing, loadedPdfPaths, loadedCadPaths)
    Exit Function
  End If

  ' 旧ファイル名→新ファイル名のペアを取得
  Dim oldToNewPdf As ObjResult
  Dim oldToNewCad As ObjResult
  Set oldToNewPdf = fileService.NewNames(papers.value, loadedPdfPaths.value, config.CompanyId, config.ProjectId, config.StructureId)
  Set oldToNewCad = fileService.NewNames(papers.value, loadedCadPaths.value, config.CompanyId, config.ProjectId, config.StructureId)
  If (Not IsResultAllOK(oldToNewPdf, oldToNewCad)) Then
    Set SuggestNewName = PickResultErr(Nothing, oldToNewPdf, oldToNewCad)
    Exit Function
  End If

  ' シートへ書き込み
  Dim wroteRes As ObjResult
  Set wroteRes = sheetService.WriteSuggestRenames(oldToNewPdf.value, oldToNewCad.value)
  If (Not wroteRes.isOK) Then
    Set SuggestNewName = wroteRes
    Exit Function
  End If

  Set SuggestNewName = Success()
End Function


Public Function RenameFiles() As ObjResult
  Dim fileService As ServiceFile
  Set fileService = New ServiceFile
  Dim sheetService As ServiceSheet
  Set sheetService = New ServiceSheet
  Dim config As ServiceSystemConfig
  Set config = New ServiceSystemConfig

  ' システム設定の読み込み
  Dim resConfig As ObjResult
  Set resConfig = config.LoadUserInput()
  If Not resConfig.isOK Then
    Set RenameFiles = resConfig
    Exit Function
  End If

  ' ファイルパス一覧を読み込み
  Dim oldPaths As ObjResult
  Dim newPaths As ObjResult
  Set oldPaths = sheetService.ReadOldFilePaths(config.PdfFolderPath, config.PdfExtensions, config.CadFolderPath, config.CadExtensions)
  Set newPaths = sheetService.ReadNewFilePaths(config.PdfFolderPath, config.PdfExtensions, config.CadFolderPath, config.CadExtensions)
  If (Not IsResultAllOK(oldPaths, newPaths)) Then
    Set RenameFiles = PickResultErr(Nothing, oldPaths, newPaths)
    Exit Function
  End If

  ' ファイルのリネーム
  Dim renameRes As ObjResult
  Set renameRes = fileService.RenameFiles(oldPaths.value, newPaths.value)
  If Not renameRes.isOK Then
    Set RenameFiles = renameRes
    Exit Function
  End If

  Set RenameFiles = Success()
End Function


Public Function NoRenamePdf() As ObjResult
  Dim sheetService As ServiceSheet
  Set sheetService = New ServiceSheet
  Dim config As ServiceSystemConfig
  Set config = New ServiceSystemConfig

  ' システム設定の読み込み
  Dim resConfig As ObjResult
  Set resConfig = config.LoadUserInput()
  If Not resConfig.isOK Then
    Set NoRenamePdf = resConfig
    Exit Function
  End If

  Set NoRenamePdf = sheetService.ShowNoRename4PDFMsg(config.PdfExtensions)
End Function


Public Function NoRenameCad() As ObjResult
  Dim sheetService As ServiceSheet
  Set sheetService = New ServiceSheet
  Dim config As ServiceSystemConfig
  Set config = New ServiceSystemConfig

  ' システム設定の読み込み
  Dim resConfig As ObjResult
  Set resConfig = config.LoadUserInput()
  If Not resConfig.isOK Then
    Set NoRenameCad = resConfig
    Exit Function
  End If

  Set NoRenameCad = sheetService.ShowNoRename4CADMsg(config.CadExtensions)
End Function