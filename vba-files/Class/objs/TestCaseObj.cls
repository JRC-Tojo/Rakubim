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

' テストケース名
Private inner_testCaseFileName As String
' 旧名称のファイル一覧
Private inner_oldPdfNames As Collection
Private inner_oldDwgNames As Collection
' 新名称のファイル一覧
Private inner_newPdfNames As Collection
Private inner_newDwgNames As Collection
' 図名一覧
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
Private  Function GetSummary()
  ' マッチング
  Call MatchPaper2File()
  ' リネーム
  Call CorrectRenameFiles()
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

  ' 図面とファイルの対応関係を記録しているシート
  Const targetSheetName = "図面ファイル対応一覧"

  ' 旧名称のファイル一覧（確実にすべてのデータが埋まっているため，先行して取得する）
  Set inner_oldPdfNames = sUtill.Range2Collection(wb.Worksheets(targetSheetName), "B2:B10000")
  Set inner_oldDwgNames = sUtill.Range2Collection(wb.Worksheets(targetSheetName), "D2:D" & inner_oldPdfNames.Count + 1, False)
  
  ' 新名称のファイル一覧
  Set inner_newPdfNames = sUtill.Range2Collection(wb.Worksheets(targetSheetName), "C2:C" & inner_oldPdfNames.Count + 1, False)
  Set inner_newDwgNames = sUtill.Range2Collection(wb.Worksheets(targetSheetName), "E2:E" & inner_oldPdfNames.Count + 1, False)

  ' 図名一覧
  Set inner_paperNames = sUtill.Range2Collection(wb.Worksheets(targetSheetName), "A2:A" & inner_oldPdfNames.Count + 1, False)
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
Private Sub MatchPaper2File(matcher As Paper2FileMatcher, paper As PaperObj, correctFile As String)
  ' TODO: 実装
End Sub

  
' ファイルが正しくリネームされるか
Private Sub CorrectRenameFiles()
  ' TODO: 実装
End Sub