Attribute VB_Name = "SheetProcess"
' WorkSheetに対する処理を記載
' WorkSheetから読み込む処理をすべてここに記載することで、Sheetに変更が入った場合に対応しやすくする

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

' データファイル説明書の行数を返す
Function getTableCols()
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.DataTablePage)
    getTableCols = WorksheetFunction.CountA(.Range("A:A")) - 1
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


