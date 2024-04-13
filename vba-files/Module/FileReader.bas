Attribute VB_Name = "FileReader"
' ディレクトリ内のファイル名一覧を取得
Private Function getAllFileNames(folderName As String)
  Dim i As Integer: i = 0
  Dim returnVals() As String: ReDim returnVals(1)

  ' 先頭のファイル名の取得
  Dim strFileName As String: strFileName = Dir(ActiveWorkbook.path & "\" & folderName & "\", vbNormal)
  
  ' ファイルが見つからなくなるまで繰り返す
  Do While strFileName <> ""
    ' リストに値を追加
    ReDim Preserve returnVals(i + 1)
    returnVals(i) = strFileName
    
    ' 次のファイル名を取得
    strFileName = Dir()
    
    i = i + 1
  Loop
  
  getAllFileNames = returnVals
End Function

' ファイルの一覧オブジェクトを取得
Public Function GetFileCollection(folderName As String) As Collection
  Dim filePaths As Variant
  filePaths = getAllFileNames(folderName)
  Dim file As FileObj
  Dim returns As Collection
  Set returns = New Collection

  If (UBound(filePaths) = 0) Then
    Set GetFileCollection = returns
    Exit Function
  End If
  
  For Each filePath In filePaths
    If filePath <> "" Then
      Set file = New FileObj
      Call file.InitFileObj(filePath)
      If (file.FileType <> "OTHER") Then
        returns.Add file
      End If
    End If
  Next filePath
  
  Set GetFileCollection = DropDuplicateName(returns)
End Function

' 与えられたFileObjのCollectionからnameの重複を削除したCollectionを返す
Private Function DropDuplicateName(fileObjCollection As Collection)
  Dim tmpDict As Scripting.Dictionary
  Set tmpDict = New Scripting.Dictionary
  Dim file As FileObj
  
  For Each file In fileObjCollection
    If tmpDict.Exists(file.SnippedName) = True Then
      For Each tmpKey In dictKeys(file.PaperPaths)
        Set tmpDict(file.SnippedName).PaperPaths = JointPaperPaths(tmpDict(file.SnippedName).PaperPaths, tmpKey, file.PaperPaths(tmpKey))
      Next tmpKey
      For Each num In file.PaperNumbers
        tmpDict(file.SnippedName).PaperNumbers.Add num
      Next num
    Else
      tmpDict.Add file.SnippedName, file
    End If
  Next file
  
  Set DropDuplicateName = dictValues(tmpDict)
End Function

' 重複したFileObjをまとめる際に，PaperPathsを統合する
Private Function JointPaperPaths(PaperPaths As Scripting.Dictionary, ByVal newPaperKey As String, newPaperPath As String, Optional suffix As Integer = 0)
  ' 無効なパスを登録しない
  If newPaperPath = "" Then
    Set JointPaperPaths = PaperPaths
    Exit Function
  End If
  
  Dim tmpKey As String
  tmpKey = IIf(suffix = 0, newPaperKey, newPaperKey & "-" & suffix)

  If PaperPaths.Exists(tmpKey) = True Then
    Set JointPaperPaths = JointPaperPaths(PaperPaths, newPaperNumber, newPaperPath, suffix + 1)
  Else
    PaperPaths.Add tmpKey, newPaperPath
    Set JointPaperPaths = PaperPaths
  End If
End Function
