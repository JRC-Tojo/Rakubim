VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServiceMatching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private inner_cachedPaperAnalyzedNameToExtractedFiles As ObjDict

Private Sub Class_Initialize()
  Set inner_cachedPaperAnalyzedNameToExtractedFiles = New ObjDict
End Sub

' ServiceMatching.Match: 図面情報とファイル一覧を受け取り、簡易マッチング結果を返す
' 引数: papers (ObjList of ModelPaper), files (ObjList of ModelFile)
'       Optional useSeq (Boolean) - 連番照合モードを使用するか
' 戻り値: ObjResult - papersで指定された順にマッチングしたファイル名のObjListを返す
Public Function Match(papers As ObjList, files As ObjList, Optional useSeq As Boolean = False) As ObjResult
  On Error GoTo ErrHandler

  If (files.Count = 0) Then
    Set Match = Failure("service.match.noFiles")
    Exit Function
  End If
  
  ' AnalyzedNameでFilesを集約
  Dim name2fileList As ObjDict
  Set name2fileList = groupFilesByAnalyzedName(files)

  ' 連番(Sequence)照合モードか否かでマッチング機構を使い分ける
  Dim matchedFileNames As New ObjList
  If useSeq Then
    Set matchedFileNames = matchWithSequence(papers, name2fileList)
  Else
    Set matchedFileNames = matchWithNLP(papers, name2fileList)
  End If

  Set Match = Success(matchedFileNames)
  Exit Function
ErrHandler:
  Set Match = Failure("service.match.error", Err.Number, Err.Description)
End Function

' AnalyzedNameをKeyとしてModelFileをListに集約したDictを返す
Private Function groupFilesByAnalyzedName(files As ObjList) As ObjDict
  Dim nameTofileList As ObjDict
  Set nameTofileList = New ObjDict

  Dim fileList As ObjList
  Dim file As ModelFile
  For Each file In files
    If nameTofileList.Exists(file.AnalyzedName) Then
      Call nameTofileList.At(file.AnalyzedName).Append(file)
    Else
      Set fileList = New ObjList
      Call fileList.InitItems(file)
      Call nameTofileList.Append(file.AnalyzedName, fileList)
    End If
  Next file

  Set groupFilesByAnalyzedName = nameTofileList
End Function

' Private helper: 連番モードで最も登録ファイルが多いファイルオブジェクトの中からi番目に登録されているファイル情報を返す．
Private Function matchWithSequence(papers As ObjList, name2fileList As ObjDict) As ObjList
  ' 番号とファイル名の対応データ
  Dim paperNumber2FileSourceName As ObjDict
  Set paperNumber2FileSourceName = buildPaperNumberToSourceNameDict(name2fileList)

  ' file.PaperNumberの最も若い番号
  Dim lowestFileNumber As Long
  lowestFileNumber = 2 ^ 31 - 1
  Dim fNum As Variant
  For Each fNum In paperNumber2FileSourceName.Keys
    If (lowestFileNumber > fNum) Then lowestFileNumber = fNum
  Next

  ' それぞれの番号に対応するファイル名を取得
  Dim matchedFileNames As ObjList
  Set matchedFileNames = New ObjList
  Dim p As ModelPaper
  For Each p In papers
    Call matchedFileNames.Append(paperNumber2FileSourceName.At(p.PaperIndex + (lowestFileNumber - 1), defValue:=""))
  Next p

  Set matchWithSequence = matchedFileNames
End Function

' Private helper: 自然言語処理に基づいたマッチング結果を返す
Private Function matchWithNLP(papers As ObjList, name2fileList As ObjDict) As ObjList
  Dim matchedFileSourceNames As ObjList
  Set matchedFileSourceNames = New ObjList
  
  Dim p As ModelPaper
  Dim extractedFile As ModelFile
  Dim matchedFileSourceName As String
  For Each p In papers
    Set extractedFile = extractFileWithPaper(p, name2fileList)
    If Not (extractedFile Is Nothing) Then matchedFileSourceName = extractedFile.SourceName
    Call matchedFileSourceNames.Append(matchedFileSourceName)
    matchedFileSourceName = ""
  Next p

  Set matchWithNLP = matchedFileSourceNames
End Function


' PaperNumberに対応するSourceNameの辞書を返す
' ModelFile.AnalyzedNameをKeyとしてModelFileを集約し，集約数が最も多いAnalyzedNameから辞書を作成する
Private Function buildPaperNumberToSourceNameDict(name2fileList As ObjDict) As ObjDict
  ' 集約数が最も多いAnalyzedNameのList[ModelFile]を抽出する
  Dim maxFileList As ObjList
  Set maxFileList = New ObjList
  Dim fileList As ObjList
  For Each fileList In name2fileList.Values
    If fileList.Count > maxFileList.Count Then Set maxFileList = fileList
  Next fileList

  ' PaperNumberをキー，SourceNameをValueとする辞書を作成する
  Dim num2SourceName As ObjDict
  Set num2SourceName = New ObjDict
  Dim file As ModelFile
  For Each file In maxFileList
    Call num2SourceName.Append(file.PaperNumber, file.SourceName)
  Next file

  Set buildPaperNumberToSourceNameDict = num2SourceName
End Function

' SequenceMatcherによって特定のPaperに最もマッチするModelFileを抽出する
Private Function extractFileWithPaper(paper As ModelPaper, name2fileList As ObjDict) As ModelFile
  Dim seq As SequenceMatcher
  Set seq = New SequenceMatcher
  
  ' 最もPaper.AnalyzedNameとマッチするFile.AnalyzedNameを抽出
  Dim bestFiles As ObjList
  If inner_cachedPaperAnalyzedNameToExtractedFiles.Exists(paper.AnalyzedName) Then
    ' キャッシュからbestFilesを取得することで，同一AnalyzedNameに対する探索コストを省略する
    Set bestFiles = inner_cachedPaperAnalyzedNameToExtractedFiles.At(paper.AnalyzedName)
  Else
    Dim n2fs As ModelKeyValue
    Dim bestRatio As Double: bestRatio = -1
    For Each n2fs In name2fileList
      Call seq.set_seqs(paper.AnalyzedName, n2fs.Key)
      If seq.ratio > bestRatio Then
        Set bestFiles = n2fs.Value
        bestRatio = seq.ratio
      End If
    Next n2fs

    Call inner_cachedPaperAnalyzedNameToExtractedFiles.Append(paper.AnalyzedName, bestFiles)
  End If

  ' 抽出したFilesから番号情報がマッチするModelFileを抽出
  ' -> 抽出したFileが値を持っていない場合はすべての番号を有していると考える
  Dim bestFile As ModelFile
  Dim f As ModelFile
  For Each f In bestFiles
    If (f.PaperNumbers.Count = 0) Then
      Set bestFile = f
    ElseIf (bestFiles.Count = 1) Then
      Set bestFile = f
    ElseIf (f.PaperNumbers.Contains(paper.PaperNumber)) Then
      Set bestFile = f
    End If
  Next f

  Set extractFileWithPaper = bestFile
End Function