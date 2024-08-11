Attribute VB_Name = "MatchingTest"

' 登録されているすべてのテストケースを実行し，結果を出力する
Public Sub SummaryMatchingTests()
  Dim objFSO As FileSystemObject
　Set objFSO = New FileSystemObject
  
  ' 登録されているすべてのテストケースに対してテストを実行
  For Each objFile In objFSO.GetFolder(ActiveWorkbook.path & "\testCases").Files
    ' ファイル名の先頭に「skip」と書かれたテストケースは除外する
    If (Left(objFile.Name, 4) <> "skip") Then
      Dim testObj As TestCaseObj
      Set testObj = New TestCaseObj
  
      ' テストの準備
      Call testObj.InitTestCase(objFile.Name)
  
      ' テスト結果をDebug.Printで表示する
      Call testObj.GetSummary
    Else
      Debug.Print "Skipped the testCase in " & objFile.Name
    End If
  Next
End Sub
