Attribute VB_Name = "SheetProcess"
' WorkSheetに対する処理を記載
' WorkSheetから読み込む処理をすべてここに記載することで、Sheetに変更が入った場合に対応しやすくする

' 目録入力シートでスキップしてよい空白行の数
Public Const LIMITED_SPACE_COUNTER = 3

' 指定した行番号の図面名称が空欄か否かをチェックする
Function checkSpaceRow(rowNum As Integer)
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.MokurokuPage)
    checkSpaceRow = .Cells(rowNum, "B").Value = ""
  End With
End Function

' 目録入力シートのPDFファイルとDWGファイルがすべて埋まっていることを確認
Function checkFill4Mokuroku()
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.MokurokuPage)
    Dim paperCounts As Integer: paperCounts = WorksheetFunction.CountA(.Range("B:B"))
    Dim pdfCounts As Integer: pdfCounts = WorksheetFunction.CountA(.Range("G:G"))
    Dim dwgCounts As Integer: dwgCounts = WorksheetFunction.CountA(.Range("H:H"))
  End With
  
  Dim systemObj As SystemIdObj
  Set systemObj = New SystemIdObj
  If systemObj.SelectedFileType = "ALL" Then
    checkFill4Mokuroku = Equal(paperCounts, pdfCounts, dwgCounts)
  ElseIf systemObj.SelectedFileType = "PDF" Then
    checkFill4Mokuroku = Equal(paperCounts, pdfCounts)
  ElseIf systemObj.SelectedFileType = "CAD" Then
    checkFill4Mokuroku = Equal(paperCounts, dwgCounts)
  End If
End Function

' ファイル名変更シートの旧ファイル名と新ファイル名がすべて埋まっていることを確認
Function checkFill4File()
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.RenamePage)
    Dim oldCounts As Integer: oldCounts = WorksheetFunction.CountA(.Range("A:A"))
    Dim newCounts As Integer: newCounts = WorksheetFunction.CountA(.Range("B:B"))
  End With
  
  checkFill4File = Equal(oldCounts, newCounts)
End Function

' 入力済みデータを削除する
Function ResetData(sheetName As String, deleteRange As String)
  With Worksheets(sheetName).Range(deleteRange)
    .Value = ""
    .Validation.Delete
  End With
End Function

' 目録から行番号として指定された行に記載された情報を返す
Public Function getPaperInfo(rowNum As Integer)
  Dim paper As PaperObj
  Set paper = New PaperObj
  
  ' 情報の取得
  ' TODO: 枝番が入ったときのエラー処理
  Call paper.GetPaper(rowNum)
  
  Set getPaperInfo = paper
End Function

' 目録入力の既存ファイル名群に記入
Function WriteOldP4Mokuroku( _
  rowIdx As Integer, _
  paperGlobalNumber As String, _
  ByVal oldPDF As String, _
  ByVal oldDWG As String _
)
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.MokurokuPage)
    .Cells(rowIdx, "A").Value = paperGlobalNumber
    .Cells(rowIdx, "G").Value = oldPDF
    .Cells(rowIdx, "H").Value = oldDWG
  End With
End Function

' ファイル名変更シートに古いパスと新しいパスを書き込む
' 引数のpathListはFinalPaperのGenerateIndex2PathList()によって生成される図面名称とファイル名の対応
' 引数のlastIdxは最後に使用した連番（これをもとに書き込む位置を決定）
Function writeName2PathList(oldPath As String, newPath As String, lastIdx As Integer)
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  
  Worksheets(sheets.RenamePage).Cells(lastIdx + 1, "A").Value = oldPath
  Worksheets(sheets.RenamePage).Cells(lastIdx + 1, "B").Value = newPath
End Function

' 指定された行番号のデータを読み取り、リネーム処理を行う
Function Renamer(ByVal rowNum As Integer, ByRef oldFileName As String, ByRef newFileName As String)
  Dim folderName As String
  Dim extension As String
  Dim filePath As String
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  Dim objFSO As FileSystemObject
  Set objFSO = New FileSystemObject
  Dim systemIds As SystemIdObj
  Set systemIds = New SystemIdObj
  
  With Worksheets(sheets.RenamePage)
    oldFileName = .Cells(rowNum, "A")
    If oldFileName = "" Then
      Renamer = False
    Else
      'Debug.Print oldFileName & "|" & StrConv(Right(oldFileName, 3), vbUpperCase) & "|" & .Cells(rowNum, "B")
      
      extension = StrConv(Right(oldFileName, 3), vbLowerCase)
      folderName = IIf(extension = "pdf", "PDF", IIf(extension = "dwg", "CAD", ""))
      filePath = systemIds.TargetPath & "\" & folderName & "\" & oldFileName
      If objFSO.FileExists(filePath) Then
        newFileName = .Cells(rowNum, "B")
        newPath = systemIds.TargetPath & "\" & folderName & "\" & newFileName
        
        ' リネームでエラーが出た（既に開かれていた）際にはダイアログを出す
        On Error Resume Next
        Name filePath As newPath
        If Err.Number > 0 Then
          Debug.Print oldFileName & "はファイルが開かれていたため、リネームに失敗しました。　エラーコード：" & Err.Number
          newFileName = ""
        End If
      Else
        Debug.Print oldFileName & "は存在しないため、リネームできませんでした。"
      End If
    End If
  End With
  
  Renamer = True
End Function

' 「目録入力」シートの一番下の要素の行番号を取得する
Function getLastPaperRows()
  Dim rowNum As Integer
  Dim paperCounter As Integer
  
  ' 何回空欄の行が連続したのか
  Dim spaceCounter As Integer: spaceCounter = 0

  rowNum = INIT_ROWNUM_MOKUROKU
  Do
    ' 目録データの取得
    If checkSpaceRow(rowNum) Then
      ' 将来的に
      ' ・目録一行空け：次の列から目録を再開
      ' ・目録二行空け：次のページから目録を再開
      ' という規則にする予定のため、spaceCount = 3までは容認
      spaceCounter = spaceCounter + 1
      If spaceCounter = LIMITED_SPACE_COUNTER Then Exit Do
    Else
      spaceCounter = 0
    End If
       
    ' 次処理に向けて値を更新
    rowNum = rowNum + 1
    paperCounter = paperCounter + 1
  Loop

  getLastPaperRows = rowNum
End Function

' リネーム対象のファイル数を返す
Function getRanameFileCounts()
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.RenamePage)
    getRanameFileCounts = WorksheetFunction.CountA(.Range("A:A")) - 1
  End With
End Function

' データファイル説明書の行数を返す
Function getTableCols()
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.DataTablePage)
    getTableCols = WorksheetFunction.CountA(.Range("A:A")) - 1
  End With
End Function

' マッチング器の初期化
Public Function GetMatcher() As Paper2FileMatcher
  Dim matcher As Paper2FileMatcher
  Set matcher = New Paper2FileMatcher
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  Dim sUtill As SheetUtils
  Set sUtill = New SheetUtils
  Dim systemIds As SystemIdObj
  Set systemIds = New SystemIdObj

  ' マッチング機構の初期化
  Call matcher.BaseInitMatcher(systemIds.TargetPath)

  ' ファイルが存在しない場合はmatcherを返さずに終了
  If (Not matcher.ExistFiles) Then Exit Function

  ' Matcherに必要な図面名称の一覧を登録する
  Call matcher.InitMatcher(sUtill.Range2Collection(Worksheets(sheets.MokurokuPage), "B2:B10000"))

  Set GetMatcher = matcher
End Function

' 各セルに実在ファイルのプルダウンを設定する
Public Sub SetFilePulldownMenu(matcher As Paper2FileMatcher, lastRowNum As Integer)
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.MokurokuPage)
    Call SetPulldownMenu(.Range("G2:G" & lastRowNum), matcher.loadedPdfPaths, 1)
    Call SetPulldownMenu(.Range("H2:H" & lastRowNum), matcher.loadedDwgPaths, 2)
  End With
End Sub


' 指定した配列の要素をプルダウンで選択できるようにする
' @param: targetCells -> プルダウンを設定するセル範囲
' @param: files -> プルダウンに記載するファイル名のリスト
' @param: z -> プルダウンのID（Pulldownsシートに書き込む際の列番号になる）
Private  Function SetPulldownMenu(targetCells As Range, fileNames As Collection, z As Integer)
  Dim sUtils As SheetUtils
  Set sUtils = New SheetUtils
  Dim targetColName As String
  targetColName = sUtils.ColumnIdx2Name(z)

  ' プルダウンシートにプルダウンで表示する内容の一覧を書き込む
  Call sUtils.Collection2Cell(fileNames, "Pulldowns", targetColName & "1")

  With targetCells.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
    xlBetween, Formula1:="=Pulldowns!$" & targetColName & "$1:$" & targetColName & "$" & fileNames.Count
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = "らくびむ  - 入力エラー -"
    .InputMessage = ""
    .ErrorMessage = "フォルダ内に存在しないファイル名は入力できません。"
    .IMEMode = xlIMEModeNoControl
    .ShowInput = True
    .ShowError = True
  End With
End Function

' 図面名と図番の関係性を返す
Function getRelationName2ID()
  Dim returnDic As Object
  Set returnDic = CreateObject("Scripting.Dictionary")
  Dim rowNum As Integer
  Dim spaceCounter As Integer: spaceCounter = 0
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.MokurokuPage)
    rowNum = 2
    Do
      ' 目録データの取得
      If checkSpaceRow(rowNum) Then
        ' 将来的に
        ' ・目録一行空け：次の列から目録を再開
        ' ・目録二行空け：次のページから目録を再開
        ' という規則にする予定のため、spaceCount = 3までは容認
        spaceCounter = spaceCounter + 1
        If spaceCounter = 3 Then Exit Do
      Else
        spaceCounter = 0
        returnDic.Add .Cells(rowNum, "B").Value, .Cells(rowNum, "A").Value
      End If
      
      rowNum = rowNum + 1
    Loop
  End With
  
  Set getRelationName2ID = returnDic

End Function

' 図面名と以前のファイル名の関係性を返す
Function getRelationName2OldPath()
  Dim pdfDic As Object
  Dim dwgDic As Object
  Set pdfDic = CreateObject("Scripting.Dictionary")
  Set dwgDic = CreateObject("Scripting.Dictionary")
  Dim rowNum As Integer
  Dim spaceCounter As Integer: spaceCounter = 0
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  
  With Worksheets(sheets.MokurokuPage)
    rowNum = 2
    Do
      ' 目録データの取得
      If checkSpaceRow(rowNum) Then
        ' 将来的に
        ' ・目録一行空け：次の列から目録を再開
        ' ・目録二行空け：次のページから目録を再開
        ' という規則にする予定のため、spaceCount = 3までは容認
        spaceCounter = spaceCounter + 1
        If spaceCounter = 3 Then Exit Do
      Else
        spaceCounter = 0
        pdfDic.Add .Cells(rowNum, "B").Value, .Cells(rowNum, "G").Value
        dwgDic.Add .Cells(rowNum, "B").Value, .Cells(rowNum, "H").Value
      End If
      
      rowNum = rowNum + 1
    Loop
  End With
  
  getRelationName2OldPath = Array(pdfDic, dwgDic)
  
End Function

' 以前のファイル名と新しいファイル名の関係性を返す
Function getRelationOld2NewPath()
  Dim returnDic As Object
  Set returnDic = CreateObject("Scripting.Dictionary")
  Dim rowNum As Integer
  Dim spaceCounter As Integer: spaceCounter = 0
  Dim test As Integer: test = 1
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  
  With Worksheets(sheets.RenamePage)
    rowNum = 2
    Do While .Cells(rowNum, "A").Value <> 0
      returnDic.Add .Cells(rowNum, "A").Value, .Cells(rowNum, "B").Value
      rowNum = rowNum + 1
    Loop
  End With
  
  Set getRelationOld2NewPath = returnDic
  
End Function

' 図面名とファイル名の関連性を出力する
Function writeLinkTable( _
  contractID As String, _
  DesignID As String, _
  ByVal name2oldP As Variant, _
  name2ID As Object, _
  oldP2newP As Object _
)
  Dim writeRowNum As Integer: writeRowNum = 1
  Dim pdfPath As String
  Dim dwgPath As String
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  
  With Worksheets(sheets.DataTablePage)
    ' データの書き出し
    For Each paperName In name2oldP(0)
      .Cells(2 * writeRowNum, 1).Value = contractID
      .Cells(2 * writeRowNum + 1, 1).Value = contractID
      .Cells(2 * writeRowNum, 2).Value = name2ID(paperName)
      .Cells(2 * writeRowNum + 1, 2).Value = name2ID(paperName)
      .Cells(2 * writeRowNum, 3).Value = DesignID
      .Cells(2 * writeRowNum + 1, 3).Value = DesignID
      .Cells(2 * writeRowNum, 4).Value = paperName
      .Cells(2 * writeRowNum + 1, 4).Value = paperName
      pdfPath = oldP2newP(name2oldP(0)(paperName))
      dwgPath = oldP2newP(name2oldP(1)(paperName))
      .Cells(2 * writeRowNum, 5).Value = pdfPath
      .Cells(2 * writeRowNum + 1, 5).Value = dwgPath
      writeRowNum = writeRowNum + 1
    Next
  
    ' 印刷範囲の設定
    .PageSetup.PrintArea = "$A$1:$F$" & getTableCols() + 1
  End With
  
  ' 罫線の描画
  Call ResetLines(sheets.DataTablePage, "A1", "F1048576")
  endRow = 2 * (writeRowNum - 1) + 1
  Call MakeLattice(sheets.DataTablePage, "A1", "F" & endRow)
  
End Function

' スプラッシュ画面を閉じる
Private Sub KillForm()
    Unload SplashScreen
End Sub


