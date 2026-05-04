VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestFullMatching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' テストケースを格納するフォルダ名
Private Const CASES_FOLDER_NAME = "testCases"

' テストケースExcelのうち，設定項目シートの名称
Private Const BASE_INFO_SHEET_NAME = "基本情報"

' テストケースExcelのうち，テスト対象データを格納したシートの名称
Private Const TEST_DATA_SHEET_NAME = "図面ファイル対応一覧"


Public Function Run()
  Dim Ts As TestSuite
  Set Ts = New TestSuite
  Ts.Description = "TestFullMatching"

  Dim testFolder As ObjPath
  Set testFolder = New ObjPath
  Set testFolder = testFolder.Init(ActiveWorkbook.path).Child("testCases")

  If Not testFolder.Exists() Then
    Set Run = Ts
    Exit Function
  End If

  Dim caseFile As ObjPath
  For Each caseFile In testFolder.Colldir()
    Call _Run(Ts, caseFile)
  Next

  Set Run = Ts
End Function


Private Function _Run(ByRef Ts As TestSuite, wbPath As ObjPath)
  ' load case
  Dim wb As Workbook
  Set wb = LoadTestWorkbook(wbPath)
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

  ' TODO: loadedPathsをダミーファイル群に置換
  ' TODO: PDF, CADで処理を分岐
  Dim loadedFiles As ObjResult
  Set loadedFiles = matchingController.LoadFiles(loadedPaths.value, useSeq)

  With Ts.Test("Matching (Case: " & wbPath.Stem & ")")
    ' Run target process

    ' Check Result
    ' Dim i As Long
    ' For i = 0 To count
    '   ' .IsEqual <処理結果>[i] <想定解>[i]  ' PDF
    '   ' .IsEqual <処理結果>[i] <想定解>[i]  ' CAD
    ' Next i
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
End Function


Private Function LoadTestWorkbook(targetPath As ObjPath) As Workbook
  If (targetPath.Extension <> "xlsb") Then Exit Function

  Dim ExcelApp As Application
  Set ExcelApp = New Application

  'エクセルを不可視で開く
  ExcelApp.Visible = False      'エクセル可視/不可視設定
  ExcelApp.DisplayAlerts = False      '警告メッセージをオフ
  Set LoadTestWorkbook = ExcelApp.Workbooks.Open(targetPath, , True)
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


' 読み込んだテストケースのうち，マッチング後のファイル一覧を元のファイル一覧としてlist[Path]を返す
Private Function LoadDummyPaths(wb As Workbook) As ObjList

End Function


' TestCaseObjの初期化
' Public Function InitTestCase(testCaseFileName As String)
'   inner_testCaseFileName = testCaseFileName
  
'   ' テストケースを取得
'   Call SetInnerVals()
'   ' テストケースを書き出し
'   Call CreateTestCaseFiles()
' End Function


' ファイル名一覧や図名一覧を取得
Private Function SetInnerVals()
  Dim ExcelApp As New Application
  Dim wb As Workbook

  'エクセルを不可視で開く
  ExcelApp.Visible = False      'エクセル可視/不可視設定
  ExcelApp.DisplayAlerts = False      '警告メッセージをオフ
  Set wb = ExcelApp.Workbooks.Open(ActiveWorkbook.path & "\testCases\" & inner_testCaseFileName, , True)

  ' utils
  Dim sUtill As SheetUtils
  Set sUtill = New SheetUtils

  ' 基本情報を取得
  With wb.Worksheets("基本情報")
    innerInfo_companyID = .Range("B1")
    innerInfo_projectID = .Range("B2")
    innerInfo_structureID = .Range("B3")
    innerInfo_useContinuousNumber = .Range("B4")
  End With

  ' 図面とファイルの対応関係を記録しているシート
  Const targetSheetName = "図面ファイル対応一覧"

  ' 旧名称のファイル一覧（確実にすべてのデータが埋まっているため，先行して取得する）
  Set inner_oldPdfNames = sUtill.Range2Collection(wb.Worksheets(targetSheetName), "C2:C10000")
  Set inner_oldDwgNames = sUtill.Range2Collection(wb.Worksheets(targetSheetName), "E2:E" & inner_oldPdfNames.Count + 1, False)
  
  ' 新名称のファイル一覧
  Set inner_newPdfNames = sUtill.Range2Collection(wb.Worksheets(targetSheetName), "D2:D" & inner_oldPdfNames.Count + 1, False)
  Set inner_newDwgNames = sUtill.Range2Collection(wb.Worksheets(targetSheetName), "F2:F" & inner_oldPdfNames.Count + 1, False)

  ' 図名一覧
  Set inner_paperNumbers = sUtill.Range2Collection(wb.Worksheets(targetSheetName), "A2:A" & inner_oldPdfNames.Count + 1, False)
  Set inner_paperNames = sUtill.Range2Collection(wb.Worksheets(targetSheetName), "B2:B" & inner_oldPdfNames.Count + 1, False)
End Sub

' テストケース一覧から実際のファイルを生成
' 生成時に既存のファイルがある場合はこれらを削除
Private Sub CreateTestCaseFiles()
  Dim objFSO As FileSystemObject
  Set objFSO = New FileSystemObject

  ' 既存のテストケースを削除
  If (objFSO.FolderExists(ActiveWorkbook.path & "\PDF")) Then
    objFSO.DeleteFolder(ActiveWorkbook.path & "\PDF")
  End If
  objFSO.CreateFolder(ActiveWorkbook.path & "\PDF")

  If (objFSO.FolderExists(ActiveWorkbook.path & "\CAD")) Then
    objFSO.DeleteFolder(ActiveWorkbook.path & "\CAD")
  End If
  objFSO.CreateFolder(ActiveWorkbook.path & "\CAD")

  ' ファイルを新規作成
  Dim fileName As Variant
  For Each fileName In inner_oldPdfNames
    objFSO.CreateTextFile(ActiveWorkbook.path & "\PDF\" & fileName)
  Next fileName 
  For Each fileName In inner_oldDwgNames
    objFSO.CreateTextFile(ActiveWorkbook.path & "\CAD\" & fileName)
  Next fileName 
End Sub



' 図面名と実際のファイルが正しくマッチするか
Private Sub MatchPaper2File( _
  matcher As Paper2FileMatcher, _
  paper As PaperObj, _
  correctPdfName As String, _
  correctDwgName As String _
)
  ' マッチング結果を取得
  Dim pdfRes As String
  Dim dwgRes As String
  pdfRes = matcher.GetBestFile(paper, "PDF", innerInfo_useContinuousNumber)
  dwgRes = matcher.GetBestFile(paper, "CAD", False)

  ' マッチング結果が間違っていた場合は当該ログを表示
  If (pdfRes <> correctPdfName) Then
    Debug.Print "Invalid Matching (" & paper.SourceName & ") : Expected -> " & correctPdfName & " | Received -> " & pdfRes
  End If
  If (dwgRes <> correctDwgName) Then
    Debug.Print "Invalid Matching (" & paper.SourceName & ") : Expected -> " & correctDwgName & " | Received -> " & dwgRes
  End If
End Sub

  
' リネームのためのオブジェクトにデータを登録
Private Function SetDataForCreater( _
  creater As NewNameCreater, _
  paper As PaperObj, _
  oldPdfName As String, _
  oldDwgName As String _
)
  ' 図面を登録
  Call creater.Papers.Add(paper.Id, paper)
  
  ' ファイルを登録
  Dim cadFile As FileObj
  Dim pdfFile As FileObj
  Set pdfFile = New FileObj
  Set cadFile = New FileObj
  Call pdfFile.InitFileObj(oldPdfName)
  Call cadFile.InitFileObj(oldDwgName)

  Call creater.SetPDFFile(pdfFile, paper.Id)
  Call creater.SetCADFile(cadFile, paper.Id)
End Function


' ファイルが正しくリネームされるか
Private Sub CorrectRenameFiles(creater As NewNameCreater)
  Dim i As Integer
  Dim old2newNames As Scripting.Dictionary
  Set old2newNames = New Scripting.Dictionary
  Dim oldNames As Collection
  Set oldNames = creater.OldNames
  Dim newNames As Collection
  Set newNames = creater.NewNames

  ' 旧ファイル名 -> 新ファイル名 の辞書を作成
  For i = 1 To oldNames.Count
    Call old2newNames.Add(oldNames.Item(i), newNames.Item(i))
  Next i
  
  ' PDFのリネームチェック
  For i = 1 To inner_oldPdfNames.Count
    If (inner_newPdfNames.Item(i) <> old2newNames(CStr(inner_oldPdfNames.Item(i)))) Then
      Debug.Print "Invalid Rename (" & inner_newPdfNames.Item(i) & ") : Expected -> " & inner_newPdfNames.Item(i) & " | Received -> " & old2newNames(CStr(inner_oldPdfNames.Item(i)))
    End If
  Next i
  ' CADのリネームチェック
  For i = 1 To inner_oldDwgNames.Count
    If (inner_newDwgNames.Item(i) <> old2newNames(CStr(inner_oldDwgNames.Item(i)))) Then
      Debug.Print "Invalid Rename (" & inner_newDwgNames.Item(i) & ") : Expected -> " & inner_newDwgNames.Item(i) & " | Received -> " & old2newNames(CStr(inner_oldDwgNames.Item(i)))
    End If
  Next i
End Sub