Attribute VB_Name = "FileProcess"
' ファイルに関連する処理()

' 目録入力シートにおけるデータが始まる行数
Public Const INIT_ROWNUM_MOKUROKU = 2

' 目録の記入データをリセット
Function ResetMokuroku(leftRight As String)
  If leftRight = "left" Then
    Call ResetData("目録入力", "A" & INIT_ROWNUM_MOKUROKU & ":E10000")
  Else
    Call ResetData("目録入力", "G" & INIT_ROWNUM_MOKUROKU & ":H10000")
  End If
End Function

' ファイル名を解析して，対応する図面名の箇所にファイル名を入力
Function FileNames2drawingNames(pdfObjs As Collection, dwgObjs As Collection)
  Dim rowNum As Integer
  Dim paperCounter As Integer
  Dim mostFileObj As FileObj
  Dim paperInfo As PaperObj
  Dim pdfPath As String
  Dim dwgPath As String
  Dim systemIds As SystemIdObj
  Set systemIds = New SystemIdObj
  Dim cleaner As FileNameCleaner
  Set cleaner = New FileNameCleaner
  
  ' 何回空欄の行が連続したのか
  Dim spaceCounter As Integer: spaceCounter = 0
  
  ' 図面目録で重複しているワードを除去する
  Call cleaner.AnalyzeNames(Range2Collection("目録入力", "B2:B1000"))
    
  ' 目録を読み込み始める行番号
  rowNum = INIT_ROWNUM_MOKUROKU
  paperCounter = 1
  If pdfObjs.Count <> 0 Then
    Set mostFileObj = SearchMostNumbersFile(pdfObjs)
  End If
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
      
      ' 図面情報を目録の名称より取得
      Set paperInfo = getPaperInfo(rowNum)
      
      If pdfObjs.Count <> 0 Then
        pdfPath = IIf(systemIds.UseContinuousNumber, _
                    ApplyContinuousFile(mostFileObj, paperCounter), _
                    getMatchPath(pdfObjs, paperInfo, cleaner))
      End If
      If dwgObjs.Count <> 0 Then
        dwgPath = getMatchPath(dwgObjs, paperInfo, cleaner)
      End If
      
      ' 探索結果の書き込み
      WriteOldP4Mokuroku rowNum, paperInfo.GlobalNumber, pdfPath, dwgPath
      paperCounter = paperCounter + 1
    End If
       
    ' 次処理に向けて値を更新
    rowNum = rowNum + 1
  Loop
End Function

' 指定したpaperInfoに最もマッチするFileObjをobjsに指定したCollectionから探索する
Public Function getMatchFile(objs As Collection, paperInfo As PaperObj, cleaner As FileNameCleaner) As FileObj
  Dim ratios As Collection
  Set ratios = New Collection
  Dim seq As SequenceMatcher
  Dim seg1 As String, seg2 As String
  Dim paperName As String
  Dim fileName As String
  
  ' すべてのファイルに対してマッチ率を算出する
  For Each file In objs
    Set seq = New SequenceMatcher
    ' 目録名称から除外すべき名称とスペースを除去
    paperName = cleaner.NameCleaner(paperInfo.paperName)
    fileName = cleaner.NameCleaner(StrConv(file.paperName, vbWide))
    Call seq.set_seqs(paperName, fileName)
    
    ' マッチ率の記録
    ratios.Add seq.ratio
  Next file
  
  ' argmaxを取る
  Dim bestIdx As Integer
  bestIdx = ArgMax(ratios)(1)

  ' もっともマッチしたファイルObjを取得
  Set getMatchFile = objs.item(bestIdx)
End Function

' 指定したpaperInfoに最もマッチするファイル名を返す
Private Function getMatchPath(objs As Collection, paperInfo As PaperObj, cleaner As FileNameCleaner)
  Dim resultFile As FileObj
  Set resultFile = getMatchFile(objs, paperInfo, cleaner)
  
  ' 元のファイル名を返す
  If resultFile.PaperPaths.Exists("idx" & paperInfo.PaperNumber) = True Then
    getMatchPath = resultFile.PaperPaths("idx" & paperInfo.PaperNumber)
  ElseIf resultFile.PaperPaths.Exists("idx-1") = True Then
    getMatchPath = resultFile.PaperPaths("idx-1")
  Else
    getMatchPath = ""
  End If
End Function

' PDFの連番処理用
' 最も番号を保持しているファイルオブジェクトを抽出することで、連番化されたファイル名称を導く
Private Function SearchMostNumbersFile(objs As Collection)
  Dim fileCounts As Collection
  Set fileCounts = New Collection
  Dim file As FileObj
  
  ' すべてのファイルに対してファイル数を算出
  For Each file In objs
    fileCounts.Add file.PaperNumbers.Count
  Next file
  
  ' 最もファイル数を持つファイルオブジェクトを返す
  Dim bestIdx As Integer
  bestIdx = ArgMax(fileCounts)(1)
  Set SearchMostNumbersFile = objs.item(bestIdx)
End Function

' PDFの連番処理用
' 指定されたファイルオブジェクトから番号を抜き出して割り当てていく
Private Function ApplyContinuousFile(file As FileObj, indexCounter As Integer)
  If file.PaperPaths.Exists("idx" & indexCounter) = True Then
    ApplyContinuousFile = file.PaperPaths("idx" & indexCounter)
  ElseIf file.PaperPaths.Exists("idx-1") = True Then
    ApplyContinuousFile = file.PaperPaths("idx-1")
  Else
    ApplyContinuousFile = ""
  End If
End Function
