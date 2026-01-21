VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCaseObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' 基本情報
Private innerInfo_companyID As String
Private innerInfo_projectID As String
Private innerInfo_structureID As String
Private innerInfo_designID As String
Private innerInfo_useContinuousNumber As Boolean
' テストケース名
Private inner_testCaseFileName As String
' 旧名称のファイル一覧
Private inner_oldPdfNames As Collection
Private inner_oldDwgNames As Collection
' 新名称のファイル一覧
Private inner_newPdfNames As Collection
Private inner_newDwgNames As Collection
' 図名一覧
Private inner_paperNumbers As Collection
Private inner_paperNames As Collection

' TestCaseObjの初期化
Public Function InitTestCase(testCaseFileName As String)
  inner_testCaseFileName = testCaseFileName
  
  ' テストケースを取得
  Call SetInnerVals()
  ' テストケースを書き出し
  Call CreateTestCaseFiles()
End Function

' テスト結果を表示
Public Function GetSummary()
  Dim matcher As Paper2FileMatcher
  Set matcher = New Paper2FileMatcher
  Dim creater As NewNameCreater
  Set creater = New NewNameCreater

  ' 図面一覧をMatcherに登録
  Call matcher.BaseInitMatcher(ActiveWorkbook.path)
  Call matcher.InitMatcher(inner_paperNames)

  ' createrに必要な情報を登録
  Call creater.InitCreater(innerInfo_companyID, innerInfo_projectID, innerInfo_structureID)

  Dim i As Integer
  Dim paper As PaperObj
  For i = 1 To inner_paperNames.Count
    ' 無効な図面以外をテスト
    If (inner_paperNames.Item(i) <> "") Then
      ' 図面データを宣言
      Set paper = New PaperObj
      Call paper.InitPaperObj( _
        inner_paperNames.Item(i), _
        i, _
        inner_paperNumbers.Item(i) _
      )
    
      ' マッチング
      Call MatchPaper2File( _
        matcher, _
        paper, _
        inner_oldPdfNames.Item(i), _
        inner_oldDwgNames.Item(i) _
      )
    
      ' リネームのためのCreaterを定義
      Call SetDataForCreater( _
        creater, _
        paper, _
        inner_oldPdfNames.Item(i), _
        inner_oldDwgNames.Item(i) _
      )
    End If
  Next i

  ' リネーム結果の確認
  Call CorrectRenameFiles(creater)

  Debug.Print "All test cases in " & inner_testCaseFileName & " are done"
End Function


' ファイル名一覧や図名一覧を取得
Private Sub SetInnerVals()
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