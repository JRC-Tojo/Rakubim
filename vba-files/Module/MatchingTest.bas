Attribute VB_Name = "MatchingTest"

' 登録されているすべてのテストケースを実行し，結果を出力する
Public Sub SummaryMatchingTests()
  Dim objFSO As FileSystemObject
　Set objFSO = New FileSystemObject
  
  ' 登録されているすべてのテストケースに対してテストを実行
  For Each objFile In objFSO.GetFolder(ActiveWorkbook.path & "\testCases").Files
    Dim testObj As TestCaseObj
    Set testObj = New TestCaseObj

    ' テストの準備
    testObj.InitTestCase(objFile.Name)

    ' テスト結果をDebug.Printで表示する
    Call testObj.GetSummary
  Next
End Sub
