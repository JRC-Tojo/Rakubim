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
Private inner_oldFileNames As Collection
' 新名称のファイル一覧
Private inner_newFileNames As Collection
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

  ' 旧名称のファイル一覧
  Set inner_oldFileNames = sUtill.Range2Collection(wb.Worksheets("sheet1"), "A2:A10000")

  ' 図名一覧
  Set inner_paperNames = sUtill.Range2Collection(wb.Worksheets("sheet1"), "B2:B" & inner_oldFileNames.Count + 1, False)
  
  ' 新名称のファイル一覧
  Set inner_newFileNames = sUtill.Range2Collection(wb.Worksheets("sheet1"), "C2:C" & inner_oldFileNames.Count + 1, False)
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
  For Each fileName In inner_oldFileNames
    Dim fileExt As String
    fileExt = StrConv(Right(fileName, 3), vbLowerCase)

    Select Case fileExt 
      Case "pdf" 
        objFSO.CreateTextFile(ActiveWorkbook.path & "\PDF\" & fileName)
      Case "dwg"
        objFSO.CreateTextFile(ActiveWorkbook.path & "\CAD\" & fileName)
      Case Else 
        ' どこからも使わないであろうファイルはPDF，CADの両方に記載しておいて，
        ' 「マッチングしないこと」を確認する
        objFSO.CreateTextFile(ActiveWorkbook.path & "\PDF\" & fileName)
        objFSO.CreateTextFile(ActiveWorkbook.path & "\CAD\" & fileName)
    End Select
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