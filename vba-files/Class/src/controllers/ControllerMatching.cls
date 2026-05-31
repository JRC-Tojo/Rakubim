VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControllerMatching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ControllerMatching.Run: 指定されたシートから図面名を読み取り、ファイル群とマッチングして結果を返す
' 引数: em (ModelEvents) - プログレス等イベントを受け取るオブジェクト
'       sheetName (String) - 図面名が記載されたシート名
' 戻り値: ObjResult - 成功時は MatchingResults を返す
Public Function Run() As ObjResult
  Dim sheetService As ServiceSheet
  Set sheetService = New ServiceSheet
  Dim config As ServiceSystemConfig
  Set config = New ServiceSystemConfig

  ' システム設定の読み込み
  Dim resConfig As ObjResult
  Set resConfig = config.LoadUserInput()
  If Not resConfig.isOK Then
    Set Run = resConfig
    Exit Function
  End If

  ' 図面情報を取得
  Dim papers As ObjResult
  Set papers = LoadPapers(sheetService)
  If Not papers.isOK Then
    Set Run = papers
    Exit Function
  End If
  
  ' PDF と CAD を個別にマッチング
  Dim pdfRes As ObjResult
  Dim cadRes As ObjResult
  Set pdfRes = completeMatchedFilesWithPapers( _
    papers.value, config.PdfFolderPath, config.PdfExtensions, config.UseSequenceForPdf _
  )
  Set cadRes = completeMatchedFilesWithPapers( _
    papers.value, config.CadFolderPath, config.CadExtensions, config.UseSequenceForCad _
  )
  If Not IsResultAllOK(pdfRes, cadRes) Then
    Set Run = PickResultErr(Nothing, pdfRes, cadRes)
    Exit Function
  End If

  ' 記入済みの情報を削除する
  Dim clearRes As ObjResult
  Set clearRes = sheetService.ClearMatchedFileNames()
  If Not clearRes.isOK Then
    Set Run = clearRes
    Exit Function
  End If

  ' 取り出した情報をもとにシートへの書きこみを行う
  Dim wroteRes As ObjResult
  Set wroteRes = sheetService.WriteMatchedFileNames(pdfRes.value, cadRes.value)
  If Not wroteRes.isOK Then
    Set Run = wroteRes
    Exit Function
  End If

  Set Run = Success()
End Function


Public Function LoadPapers(sheetService As ServiceSheet) As ObjResult
  Dim analyzerService As ServiceNameAnalyzer
  Set analyzerService = New ServiceNameAnalyzer
  Dim cleanerService As ServiceNameCleaner
  Set cleanerService = New ServiceNameCleaner
  Dim paperService As ServicePaper
  Set paperService = New ServicePaper

  ' 図面情報を取得
  Dim resPapers As ObjResult
  Set resPapers = paperService.LoadPapers(analyzerService, sheetService)
  If Not resPapers.isOK Then
    Set LoadPapers = resPapers
    Exit Function
  End If

  ' マッチング処理に向けて頻出語彙を除去する（各Modelの中のAnalyzedNameを書き換える）
  Dim resNewPapers As ObjResult
  Set resNewPapers = cleanerService.Clean(resPapers.value)
  ' 失敗した場合はClean前のデータで続行する
  If Not resNewPapers.isOK Then Set resNewPapers = resPapers

  Set LoadPapers = resNewPapers
End Function


Public Function LoadFiles(filePaths As ObjList, useSeq As Boolean) As ObjResult
  Dim analyzerService As ServiceNameAnalyzer
  Set analyzerService = New ServiceNameAnalyzer
  Dim cleanerService As ServiceNameCleaner
  Set cleanerService = New ServiceNameCleaner
  Dim fileService As ServiceFile
  Set fileService = New ServiceFile

  ' ファイル情報をファイルパス一覧から構築する
  Dim loadedFiles As ObjResult
  Set loadedFiles = fileService.LoadFiles(analyzerService, filePaths, useSeq)
  If (Not loadedFiles.isOK) Then
    Set LoadFiles = loadedFiles
    Exit Function
  End If

  ' 読み取ったファイル情報をクリーニング
  Dim loadedCleanedFiles As ObjResult
  Set loadedCleanedFiles = cleanerService.Clean(loadedFiles.value)
  If (Not loadedCleanedFiles.isOK) Then
    ' 失敗した場合はClean前のデータで続行する
    Set loadedCleanedFiles = loadedFiles
  End If

  Set LoadFiles = loadedCleanedFiles
End Function


Private Function completeMatchedFilesWithPapers( _
  papers As ObjList, _
  folderPath As ObjPath, _
  exts As ModelExts, _
  useSeq As Boolean, _
  Optional searchRecursive As Boolean = False _
)
  Dim fileService As ServiceFile
  Set fileService = New ServiceFile
  Dim matchingService As ServiceMatching
  Set matchingService = New ServiceMatching

  ' 指定したフォルダからファイルパス一覧を取得する
  Dim loadedPaths As ObjResult
  Set loadedPaths = fileService.LoadPaths(folderPath, exts, searchRecursive)
  If (Not loadedPaths.isOK) Then
    Set completeMatchedFilesWithPapers = loadedPaths
    Exit Function
  End If

  ' ファイル情報をファイルパス一覧から構築する
  Dim loadedFiles As ObjResult
  Set loadedFiles = LoadFiles(loadedPaths.value, useSeq)
  If (Not loadedFiles.isOK) Then
    Set completeMatchedFilesWithPapers = loadedFiles
    Exit Function
  End If

  ' 各図面とファイルをマッチング
  Dim matchedFileNames As ObjResult
  Set matchedFileNames = matchingService.Match(papers, loadedFiles.value, useSeq)

  Set completeMatchedFilesWithPapers = matchedFileNames
End Function