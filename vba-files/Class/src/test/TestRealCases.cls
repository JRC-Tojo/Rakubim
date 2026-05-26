VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestRealCases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' テストケースを格納するフォルダ名
Private Const CASES_FOLDER_NAME = "testCases"


Public Function Run()
  Dim Ts As TestSuite
  Set Ts = New TestSuite
  Ts.Description = "TestRealCases"

  ' テストケースファイルを開くためのExcel App
  Dim ExcelApp As Application
  Set ExcelApp = New Application
  ExcelApp.Visible = False
  ExcelApp.DisplayAlerts = False

  Dim testFolder As ObjPath
  Set testFolder = New ObjPath
  Set testFolder = testFolder.Init(ActiveWorkbook.Path).Child(CASES_FOLDER_NAME)

  If Not testFolder.Exists() Then
    Set Run = Ts
    Exit Function
  End If

  Dim caseFile As ObjPath
  For Each caseFile In testFolder.Colldir()
    Call runCase(Ts, ExcelApp, caseFile)
  Next

  ' close Excel App
  ExcelApp.Quit

  Set Run = Ts
End Function

Private Function runCase(ByRef Ts As TestSuite, ExcelApp As Application, wbPath As ObjPath)
  Dim wb As Workbook
  Set wb = LoadTestWorkbook(ExcelApp, wbPath)
  If (wb Is Nothing) Then Exit Function

  Dim config As ServiceSystemConfig
  Set config = LoadConfigs(wb)
  If (config Is Nothing) Then Exit Function

  Dim matchingController As ControllerMatching
  Set matchingController = New ControllerMatching
  Dim sheetService As ServiceSheet
  Set sheetService = New ServiceSheet
  Call sheetService.SetWorkbook(wb)

  Dim papers As ObjResult
  Set papers = matchingController.LoadPapers(sheetService)
  If (Not papers.isOK) Then Exit Function

  ' マッチ結果をダミーファイル群として読み込み
  Dim loadedPdfPaths As ObjResult
  Dim loadedCadPaths As ObjResult
  Set loadedPdfPaths = sheetService.ReadPdfColAsPath(config.PdfFolderPath, config.PdfExtensions)
  Set loadedCadPaths = sheetService.ReadCadColAsPath(config.CadFolderPath, config.CadExtensions)
  If (Not IsResultAllOK(loadedPdfPaths, loadedCadPaths)) Then Exit Function
    
  ' ダミーファイル群に対してマッチ結果を抽出
  Dim matchedPdfNames As ObjResult
  Dim matchedCadNames As ObjResult
  Set matchedPdfNames = CalcMatchedFiles(matchingController, papers.value, loadedPdfPaths.value, config.UseSequenceForPdf)
  Set matchedCadNames = CalcMatchedFiles(matchingController, papers.value, loadedCadPaths.value, config.UseSequenceForCad)
  If (Not IsResultAllOK(matchedPdfNames, matchedCadNames)) Then Exit Function

  With Ts.Test("Matching (Case: " & wbPath.Stem & ")")
    ' Check Result
    Dim i As Long
    For i = 0 To loadedPdfPaths.value.Count - 1
      .IsEqual loadedPdfPaths.value.At(i).Name, matchedPdfNames.value.At(i)  ' PDF
      .IsEqual loadedCadPaths.value.At(i).Name, matchedCadNames.value.At(i)  ' CAD
    Next i
  End With
  
  With Ts.Test("New Name (Case: " & wbPath.Stem & ")")
    ' Run target process

    ' Check Result
    ' Dim i As Long
    ' For i = 0 To count
    '   ' .IsEqual <処理結果>[i] <想定解>[i]  ' PDF
    '   ' .IsEqual <処理結果>[i] <想定解>[i]  ' CAD
    ' Next i
  End With

  ' close workbook
  wb.Close SaveChanges:=False
End Function


Private Function LoadTestWorkbook(ExcelApp As Application, targetPath As ObjPath) As Workbook
  If (targetPath.Extension <> "xlsx") Then Exit Function
  If (InStr(targetPath.Stem, "skip")) Then Exit Function
  Set LoadTestWorkbook = ExcelApp.Workbooks.Open(targetPath.FullPath, , True)
End Function


Private Function LoadConfigs(wb As Workbook) As ServiceSystemConfig
  Dim config As ServiceSystemConfig
  Set config = New ServiceSystemConfig

  ' システム設定の読み込み
  Dim resConfig As ObjResult
  Set resConfig = config.LoadUserInput(wb)
  If Not resConfig.isOK Then Exit Function
  
  Set LoadConfigs = config
End Function


Private Function CalcMatchedFiles( _
  matchingController As ControllerMatching, _
  papers As ObjList, _
  targetPaths As ObjList, _
  useSeq As Boolean _
) As ObjResult
  ' Path -> FileObj
  Dim loadedFiles As ObjResult
  Set loadedFiles = matchingController.LoadFiles(targetPaths, useSeq)
  If (Not loadedFiles.isOK) Then
    Set CalcMatchedFiles = loadedFiles
    Exit Function
  End If

  ' 各図面とファイルをマッチング
  Dim matchingService As ServiceMatching
  Set matchingService = New ServiceMatching
  Set CalcMatchedFiles = matchingService.Match(papers, loadedFiles.value, useSeq)
End Function


' UIからの実テスト用ファイル群を生成する
Public Function GenerateRealFiles()
  ' テストケースファイルを開くためのExcel App
  Dim ExcelApp As Application
  Set ExcelApp = New Application
  ExcelApp.Visible = False
  ExcelApp.DisplayAlerts = False

  Dim testFolder As ObjPath
  Set testFolder = New ObjPath
  Set testFolder = testFolder.Init(ActiveWorkbook.Path).Child(CASES_FOLDER_NAME)

  Dim sheetService As ServiceSheet
  Set sheetService = New ServiceSheet
  Dim wb As Workbook
  Dim config As ServiceSystemConfig
  Dim caseFile As ObjPath
  For Each caseFile In testFolder.Colldir()
    Dim rc As VbMsgBoxResult
    rc = MsgBox("テストケース「" & caseFile.Stem & "」のモックファイルを生成しますか？", vbYesNo) 

    If (rc = vbYes) Then
      Set wb = LoadTestWorkbook(ExcelApp, caseFile)
      Call sheetService.SetWorkbook(wb)
      Set config = LoadConfigs(wb)
  
      ' マッチ結果をダミーファイル群として読み込み
      Dim loadedPdfPaths As ObjResult
      Dim loadedCadPaths As ObjResult
      Set loadedPdfPaths = sheetService.ReadPdfColAsPath(config.PdfFolderPath, config.PdfExtensions)
      Set loadedCadPaths = sheetService.ReadCadColAsPath(config.CadFolderPath, config.CadExtensions)
      If (Not IsResultAllOK(loadedPdfPaths, loadedCadPaths)) Then
        Dim errController As ControllerError
        Set errController = New ControllerError
        Call errController.ShowErrorMsg(PickResultErr(Nothing, loadedPdfPaths, loadedCadPaths).err) 
        Exit Function
      End If
      
      ' 旧ファイルを削除して新ファイルを生成
      Call config.PdfFolderPath.Unlink(missingOk:=True)
      Call config.CadFolderPath.Unlink(missingOk:=True)
      Dim path As ObjPath
      For Each path In loadedPdfPaths.value.Combine(loadedCadPaths.value)
        Call path.Touch()
      Next

      ' １ケース生成すれば後続の処理は不要
      wb.Close SaveChanges:=False
      MsgBox "生成が完了しました" 
      Exit Function
    End If
  Next
End Function