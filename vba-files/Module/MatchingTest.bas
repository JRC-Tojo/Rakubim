Attribute VB_Name = "MatchingTest"

' テストをスキップするための接頭辞
Private Const SKIP_PREFIX As String = "skip"

' 登録されているすべてのテストケースを実行し，結果を出力する
Public Sub SummaryMatchingTests()
  Dim objFSO As FileSystemObject
  Set objFSO = New FileSystemObject
  
  Dim testCasesPath As String
  testCasesPath = objFSO.BuildPath(ActiveWorkbook.path, "testCases")
  
  If Not objFSO.FolderExists(testCasesPath) Then
    Debug.Print "Test case folder not found: " & testCasesPath
    Exit Sub
  End If
  
  ' 登録されているすべてのテストケースに対してテストを実行
  Dim objFile As File
  For Each objFile In objFSO.GetFolder(testCasesPath).Files
    ' ファイル名の先頭に「skip」と書かれたテストケースは除外する
    If LCase(Left(objFile.Name, Len(SKIP_PREFIX))) <> SKIP_PREFIX Then
      Dim testObj As TestCaseObj
      Set testObj = New TestCaseObj
  
      ' テストの準備
      Call testObj.InitTestCase(objFile.Name)
  
      ' テスト結果をDebug.Printで表示する
      Call testObj.GetSummary
      ' オブジェクトの解放
      Set testObj = Nothing
    Else
      Debug.Print "Skipped the testCase in " & objFile.Name
    End If
  Next
  
  ' オブジェクトの解放
  Set objFSO = Nothing
End Sub
