Attribute VB_Name = "FileRenamer"
' PDFのリネームを行う
Public Function pdfRenamer()
  Dim readIdx As Integer
  Dim writeIdx As Integer
  
  Dim paperInfo As PaperObj
  Dim papernum As Integer
  Dim newPath As String
  
  Dim paperDic As Scripting.Dictionary
  Set paperDic = New Scripting.Dictionary
  Dim systemIds As SystemIdObj
  Set systemIds = New SystemIdObj
  
  Dim optionalPaperNumber As Integer: optionalPaperNumber = 0
  Dim spaceCounter As Integer: spaceCounter = 0
  
  ' 目録を読み込み始める行番号
  readIdx = INIT_ROWNUM_MOKUROKU
  writeIdx = 1
  Do
    If checkSpaceRow(readIdx) Then
      ' 将来的に
      ' 目録一行空け：次の列から目録を再開
      ' 目録二行空け：次のページから目録を再開
      ' という規則にする予定のため、spaceCount = 3までは容認
      spaceCounter = spaceCounter + 1
      If spaceCounter = 3 Then Exit Do
    Else
      spaceCounter = 0
      
      ' 目録データの取得
      Set paperInfo = getPaperInfo(readIdx)
      
      ' 書き出し
      newPath = GenerateIndex2NewPDFPath(paperInfo, systemIds)
      Call writeName2PathList(paperInfo.pdfPath, newPath, writeIdx)
      
      ' 次の連番に更新
      writeIdx = writeIdx + 1
    
    End If
    
    ' 次の行を読み込む
    readIdx = readIdx + 1
  Loop
  
  ' dwgRenamerがリネーム候補を書き出す際にどの行から書き込み始めるかを提示
  pdfRenamer = writeIdx
  
End Function

' DWGのリネームを行う
' TODO: DWGのリネームプロセスを完成させる
Public Function dwgRenamer(writeIdx As Integer)
  Dim readIdx As Integer
  Dim dwgObjs As Collection
  Dim beforePath As String: beforePath = ""
  Dim continueCounter As Integer: continueCounter = -1
  Dim paperInfo As PaperObj
  Dim newPath As String
  Dim systemIds As SystemIdObj
  Set systemIds = New SystemIdObj
  Dim cleaner As FileNameCleaner
  Set cleaner = New FileNameCleaner
  
  ' 何回空欄の行が連続したのか
  Dim spaceCounter As Integer: spaceCounter = 0
  
  ' 実際のファイルを読み込み、図面とパスのDictを形成
  Set dwgObjs = GetDWGCollection()
  
  ' 図面目録で重複しているワードを除去する
  Call cleaner.AnalyzeNames(Range2Collection(Worksheets("目録入力"), "B2:B1000"))
  
  ' 目録を読み込み始める行番号
  readIdx = INIT_ROWNUM_MOKUROKU
  Dim tmpFileObj As FileObj
  Do
    If checkSpaceRow(readIdx) Then
      ' 将来的に
      ' 目録一行空け：次の列から目録を再開
      ' 目録二行空け：次のページから目録を再開
      ' という規則にする予定のため、spaceCount = 3までは容認
      spaceCounter = spaceCounter + 1
      If spaceCounter = 3 Then Exit Do
    Else
      spaceCounter = 0
      
      ' 目録データの取得
      Set paperInfo = getPaperInfo(readIdx)
      
      ' すでにリネーム候補を提示したDWGファイルは処理しない
      If beforePath <> paperInfo.dwgPath Then
        ' 対応するファイルオブジェクトを生成
        Set tmpFileObj = getMatchFile(dwgObjs, paperInfo, cleaner)
      
        ' パスに基づいて正しい名前の候補を導出
        newPath = GenerateIndex2NewDWGPath(paperInfo, tmpFileObj, systemIds)
        
        ' 書き出し
        Call writeName2PathList(paperInfo.dwgPath, newPath, writeIdx)
        
        ' 次の連番に更新
        writeIdx = writeIdx + 1
        beforePath = paperInfo.dwgPath
        continueCounter = 0
      Else
        continueCounter = continueCounter + 1
      End If
    End If
    
    ' 次の行を読み込む
    readIdx = readIdx + 1
  Loop

End Function


' 目録を基に各図面名称とファイル名称のペアを作成する
' ここではファイル名称は最終的な名称であり、実際にそのファイルが存在しているかを考慮しない
' 引数のpaperInfoはSheetProcessモジュールのgetPaperInfo()によって作成されたDict
' 引数のlastIdxは図面全体の通し番号を付けるときに最後に使った番号を指定しておく（この処理で使用した最後の番号を戻り値で返す）
Private Function GenerateIndex2NewPDFPath( _
  paperInfo As PaperObj, _
  systemIds As SystemIdObj _
)
  Dim returnPath As String
  Dim paperID As String
  paperID = IIf(Len(paperInfo.GlobalNumber) > 4, paperInfo.GlobalNumber, Right("0000" & paperInfo.GlobalNumber, 4))

  GenerateIndex2NewPDFPath = systemIds.CompanyID & "-" & _
                               systemIds.ProjectID & "-" & _
                               systemIds.StructureID & "-" & _
                               paperID & "_" & _
                               paperInfo.SourceName & ".pdf"
End Function

' dwgファイルにおける旧ファイル名と新ファイル名のGenerater
Private Function GenerateIndex2NewDWGPath( _
  paperInfo As PaperObj, _
  dwgFileObj As FileObj, _
  systemIds As SystemIdObj _
)
  ' 連番部分
  Dim startNum As String
  startNum = PaperIDFormatter(paperInfo.GlobalNumber)
  Dim endNum As String
  endNum = PaperIDFormatter(CStr(CInt(paperInfo.GlobalNumber) + dwgFileObj.PaperNumbers.Count - 1))
  Dim rangePart1 As String
  rangePart1 = IIf(startNum = endNum, startNum, startNum & "-" & endNum)
  
  ' 名称部分
  Dim reg As RegExp
  Set reg = New RegExp
  With reg
    .Pattern = "（その[０-９]+）"
    .Global = True
  End With
  Dim name As String
  name = reg.Replace(paperInfo.paperName, "")
  
  ' （その〇～〇）部分
  Dim rangePart2 As String
  Dim matches As MatchCollection
  Set matches = reg.Execute(paperInfo.paperName)
  Dim singleNumber As String
  If matches.Count > 0 Then
    singleNumber = matches(0)
  Else
    singleNumber = ""
  End If
  
  rangePart2 = IIf( _
    startNum = endNum, _
    singleNumber, _
    "（その" & dwgFileObj.PaperNumbers(1) & "～" & dwgFileObj.PaperNumbers(1) + dwgFileObj.PaperNumbers.Count - 1 & "）")
  
  GenerateIndex2NewDWGPath = systemIds.CompanyID & "-" & _
                              systemIds.ProjectID & "-" & _
                              systemIds.StructureID & "-" & _
                              rangePart1 & "_" & _
                              name & rangePart2 & ".dwg"
End Function

' 連番（PaperID）のフォーマットを整える
Private Function PaperIDFormatter(paperID As String)
  PaperIDFormatter = IIf(Len(paperID) > 4, paperID, Right("0000" & paperID, 4))
End Function
