VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControllerSheet12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function GenerateRelationalTable() As ObjResult
  Dim matchingController As ControllerMatching
  Set matchingController = New ControllerMatching
  Dim sheetService As ServiceSheet
  Set sheetService = New ServiceSheet
  Dim config As ServiceSystemConfig
  Set config = New ServiceSystemConfig

  ' システム設定の読み込み
  Dim resConfig As ObjResult
  Set resConfig = config.LoadUserInput()
  If Not resConfig.isOK Then
    Set GenerateRelationalTable = resConfig
    Exit Function
  End If

  ' 図面情報を取得
  Dim papers As ObjResult
  Set papers = matchingController.LoadPapers(sheetService)
  If Not papers.isOK Then
    Set GenerateRelationalTable = papers
    Exit Function
  End If

  ' 入力済みのファイル群を読み込み
  Dim loadedPdfPaths As ObjResult
  Dim loadedCadPaths As ObjResult
  Set loadedPdfPaths = sheetService.ReadPdfColAsPath(config.PdfFolderPath, config.PdfExtensions)
  Set loadedCadPaths = sheetService.ReadCadColAsPath(config.CadFolderPath, config.CadExtensions)
  If (Not IsResultAllOK(loadedPdfPaths, loadedCadPaths)) Then
    Set GenerateRelationalTable = PickResultErr(Nothing, loadedPdfPaths, loadedCadPaths)
    Exit Function
  End If

  ' 旧ファイル名→新ファイル名のペアを取得
  Dim oldToNewFileNames As ObjResult
  Set oldToNewFileNames = sheetService.ReadOldToNewFileNameDict(config.PdfFolderPath, config.PdfExtensions, config.CadFolderPath, config.CadExtensions)
  If (Not oldToNewFileNames.isOK) Then
    Set GenerateRelationalTable = oldToNewFileNames
    Exit Function
  End If

  ' 記入済みの情報を削除する
  Dim clearRes As ObjResult
  Set clearRes = sheetService.ResetRelationalTable()
  If Not clearRes.isOK Then
    Set GenerateRelationalTable = clearRes
    Exit Function
  End If

  ' データファイル説明書を作成
  Dim wroteRes As ObjResult
  Dim keiyakuId As String
  keiyakuId = config.CompanyId & "-" & config.ProjectId
  Set wroteRes = sheetService.WriteRelationalTable(keiyakuId, config.DesignId, papers.value, loadedPdfPaths.value, loadedCadPaths.value, oldToNewFileNames.value)
  If Not wroteRes.isOK Then
    Set GenerateRelationalTable = wroteRes
    Exit Function
  End If

  Set GenerateRelationalTable = Success()
End Function


Public Function ExportRelationalTable() As ObjResult
  Dim sheetService As ServiceSheet
  Set sheetService = New ServiceSheet
  Dim config As ServiceSystemConfig
  Set config = New ServiceSystemConfig

  ' システム設定の読み込み
  Dim resConfig As ObjResult
  Set resConfig = config.LoadUserInput()
  If Not resConfig.isOK Then
    Set GenerateRelationalTable = resConfig
    Exit Function
  End If

  ' データファイル説明書を外部ファイルに出力する
  Dim exportRes As ObjResult
  Set exportRes = sheetService.ExportRelationalTable(config.RelationalTableOutputPath)
  If Not exportRes.isOK Then
    Set ExportRelationalTable = exportRes
    Exit Function
  End If

  Set ExportRelationalTable = Success()
End Function