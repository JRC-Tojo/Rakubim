Attribute VB_Name = "SheetProcess"
' WorkSheetに対する処理を記載
' WorkSheetから読み込む処理をすべてここに記載することで、Sheetに変更が入った場合に対応しやすくする

' 指定した行番号の図面名称が空欄か否かをチェックする
Function checkSpaceRow(rowNum As Integer)
  With Worksheets("目録入力")
    checkSpaceRow = .Cells(rowNum, "B").Value = ""
  End With
End Function

' 目録入力シートのPDFファイルとDWGファイルがすべて埋まっていることを確認
Function checkFill4Mokuroku()
  With Worksheets("目録入力")
    Dim paperCounts As Integer: paperCounts = WorksheetFunction.CountA(.Range("B:B"))
    Dim pdfCounts As Integer: pdfCounts = WorksheetFunction.CountA(.Range("G:G"))
    Dim dwgCounts As Integer: dwgCounts = WorksheetFunction.CountA(.Range("H:H"))
  End With
  
  If Sheets("入力").AllFilesBtn.Value = True Then
    checkFill4Mokuroku = Equal(paperCounts, pdfCounts, dwgCounts)
  ElseIf Sheets("入力").OnlyCADBtn.Value = True Then
    checkFill4Mokuroku = Equal(paperCounts, dwgCounts)
  ElseIf Sheets("入力").OnlyPDFBtn.Value = True Then
    checkFill4Mokuroku = Equal(paperCounts, pdfCounts)
  End If
End Function

' ファイル名変更シートの旧ファイル名と新ファイル名がすべて埋まっていることを確認
Function checkFill4File()
  With Worksheets("ファイル名変更")
    Dim oldCounts As Integer: oldCounts = WorksheetFunction.CountA(.Range("A:A"))
    Dim newCounts As Integer: newCounts = WorksheetFunction.CountA(.Range("B:B"))
  End With
  
  checkFill4File = Equal(oldCounts, newCounts)
End Function

' 入力済みデータを削除する
Function ResetData(sheetName As String, deleteRange As String)
  With Worksheets(sheetName)
    .Range(deleteRange).Value = ""
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
  paperID As String, _
  ByVal oldPDF As String, _
  ByVal oldDWG As String _
)
  With Worksheets("目録入力")
    .Cells(rowIdx, "A").Value = paperID
    .Cells(rowIdx, "G").Value = oldPDF
    .Cells(rowIdx, "H").Value = oldDWG
  End With
End Function

' ファイル名変更シートに古いパスと新しいパスを書き込む
' 引数のpathListはFinalPaperのGenerateIndex2PathList()によって生成される図面名称とファイル名の対応
' 引数のlastIdxは最後に使用した連番（これをもとに書き込む位置を決定）
Function writeName2PathList(oldPath As String, newPath As String, lastIdx As Integer)
  Worksheets("ファイル名変更").Cells(lastIdx + 1, "A").Value = oldPath
  Worksheets("ファイル名変更").Cells(lastIdx + 1, "B").Value = newPath
End Function

' 指定された行番号のデータを読み取り、リネーム処理を行う
Function Renamer(ByVal rowNum As Integer)
  Dim folderName As String
  Dim fileName As String
  Dim newPath As String
  Dim extension As String
  Dim filePath As String
  
  With Worksheets("ファイル名変更")
    fileName = .Cells(rowNum, "A")
    If fileName = "" Then
      Renamer = False
    Else
      'Debug.Print fileName & "|" & StrConv(Right(fileName, 3), vbUpperCase) & "|" & .Cells(rowNum, "B")
      
      extension = Right(fileName, 3)
      folderName = IIf(extension = "pdf", "PDF", IIf(extension = "dwg", "CAD", ""))
      filePath = ActiveWorkbook.path & "\" & folderName & "\" & .Cells(rowNum, "A")
      If Dir(filePath) <> "" Then
        newPath = ActiveWorkbook.path & "\" & folderName & "\" & .Cells(rowNum, "B")
        If Dir(newPath) = "" Then
          Name filePath As newPath
        Else
          Debug.Print newPath & "は既に存在しています"
        End If
      End If
    End If
  End With
  
  Renamer = True
End Function

' リネーム対象のファイル数を返す
Function getRanameFileCounts()
  With Worksheets("ファイル名変更")
    getRanameFileCounts = WorksheetFunction.CountA(.Range("A:A")) - 1
  End With
End Function

' リンク表の行数を返す
Function getTableCols()
  With Worksheets("リンク表")
    getTableCols = WorksheetFunction.CountA(.Range("A:A")) - 1
  End With
End Function

' 図面名と図番の関係性を返す
Function getRelationName2ID()
  Dim returnDic As Object
  Set returnDic = CreateObject("Scripting.Dictionary")
  Dim rowNum As Integer
  Dim spaceCounter As Integer: spaceCounter = 0
  
  With Worksheets("目録入力")
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
  
  With Worksheets("目録入力")
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
  
  With Worksheets("ファイル名変更")
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
  
  With Worksheets("リンク表")
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
  Call ResetLines("リンク表", "A1", "F1048576")
  endRow = 2 * (writeRowNum - 1) + 1
  Call MakeLattice("リンク表", "A1", "F" & endRow)
  
End Function
