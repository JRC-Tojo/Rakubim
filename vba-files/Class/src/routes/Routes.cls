VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Routes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ResultErr型に基づいてエラーメッセージを表示する
Public Function ShowErrorMsg(err As ObjResultErr)
  Dim errController As ControllerError
  Set errController = New ControllerError
  Call errController.ShowErrorMsg(err)
End Function


' 目録入力シートに記載された図面群のマッチング結果リストを取得する
Public Function GetMatchingLists() As ObjResult
  Dim matchingController As ControllerMatching
  Set matchingController = New ControllerMatching
  Set GetMatchingLists = matchingController.Run()
End Function


' 目録入力シートのPDF列，CAD列にファイル名のプルダウンを設置する
Public Function SetPulldownForFiles() As ObjResult
  Dim sheetController As ControllerSheet8
  Set sheetController = New ControllerSheet8
  Set SetPulldownForFiles = sheetController.SetFilesPulldown()
End Function


' 図面とマッチングしたファイルの新名称を提案する
Public Function SuggestNewFileNames() As ObjResult
  Dim sheetController As ControllerSheet10
  Set sheetController = New ControllerSheet10
  Set SuggestNewFileNames = sheetController.SuggestNewName()
End Function


' ファイル名変更シートの旧ファイル名列と新ファイル名列を読み取り、ファイルのリネームを行う
Public Function RenameFiles() As ObjResult
  Dim sheetController As ControllerSheet10
  Set sheetController = New ControllerSheet10
  Set RenameFiles = sheetController.RenameFiles()
End Function


' PDFファイルの名前変更を行わない
Public Function NoRenamePdf() As ObjResult
  Dim sheetController As ControllerSheet10
  Set sheetController = New ControllerSheet10
  Set NoRenamePdf = sheetController.NoRenamePdf()
End Function


' CADファイルの名前変更を行わない
Public Function NoRenameCad() As ObjResult
  Dim sheetController As ControllerSheet10
  Set sheetController = New ControllerSheet10
  Set NoRenameCad = sheetController.NoRenameCad()
End Function


' データファイル説明書を作成する
Public Function GenerateRelationalTable() As ObjResult
  Dim sheetController As ControllerSheet12
  Set sheetController = New ControllerSheet12
  Set GenerateRelationalTable = sheetController.GenerateRelationalTable()
End Function


' データファイル説明書を外部ファイルに出力する
Public Function ExportRelationalTable() As ObjResult
  Dim sheetController As ControllerSheet12
  Set sheetController = New ControllerSheet12
  Set ExportRelationalTable = sheetController.ExportRelationalTable()
End Function


' 目録入力シートの図面目録部分をリセットする
Public Function ResetMokuroku() As ObjResult
  Dim sheetController As ControllerSheet8
  Set sheetController = New ControllerSheet8
  Set ResetMokuroku = sheetController.ResetMokuroku()
End Function


' 目録入力シートのマッチングしたファイル部分をリセットする
Public Function ResetMatchedFiles() As ObjResult
  Dim sheetController As ControllerSheet8
  Set sheetController = New ControllerSheet8
  Set ResetMatchedFiles = sheetController.ResetMatchedFiles()
End Function


' ファイル名変更シートの提案結果をリセットする
Public Function ResetSuggestion() As ObjResult
  Dim sheetController As ControllerSheet10
  Set sheetController = New ControllerSheet10
  Set ResetSuggestion = sheetController.ResetSuggestion()
End Function


' データファイル説明書の内容をリセットする
Public Function ResetRelationalTable() As ObjResult
  Dim sheetController As ControllerSheet12
  Set sheetController = New ControllerSheet12
  Set ResetRelationalTable = sheetController.ResetRelationalTable()
End Function