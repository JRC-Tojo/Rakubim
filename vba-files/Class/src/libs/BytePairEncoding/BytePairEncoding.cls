VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BytePairEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' 指定した文字列の各文字の隙間にスペースを挿入する
Private Function SplitIntoChars(s As String) As String
  Dim list As ObjList
  Set list = New ObjList

  Dim i As Long
  For i = 1 To Len(s)
    list.Append Mid$(s, i, 1)
  Next i
  
  SplitIntoChars = list.Join(" ")
End Function

' トークンペア頻度と単語頻度を計算する関数
' 引数: corpusDict (Scripting.Dictionary) - キー: 文(トークンが空白で区切られた文字列)、値: Long(頻度)
' 戻り値: pairCounts (Scripting.Dictionary), tokenCounts (Scripting.Dictionary)
Private Sub GetStats(ByVal corpusDict As Object, ByRef pairCounts As Object, ByRef tokenCounts As Object)
  Dim sentence As Variant
  Dim freq As Long

  For Each sentence In corpusDict.Keys
    freq = CLng(corpusDict(sentence))
    Dim tokens() As String
    tokens = Split(CStr(sentence), " ")
    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
      If tokenCounts.Exists(tokens(i)) Then
        tokenCounts(tokens(i)) = tokenCounts(tokens(i)) + freq
      Else
        tokenCounts(tokens(i)) = freq
      End If
      If i < UBound(tokens) Then
        Dim pairKey As String
        pairKey = tokens(i) & "|" & tokens(i + 1)
        If pairCounts.Exists(pairKey) Then
          pairCounts(pairKey) = pairCounts(pairKey) + freq
        Else
          pairCounts(pairKey) = freq
        End If
      End If
    Next i
  Next sentence
End Sub

' 最良ペアをPMIベースで計算して返す
' 引数: pairCounts, tokenCounts (両方 Scripting.Dictionary)、minFreq (Long)
' 戻り値: bestPair (2要素のVariant配列) または Nothing、bestScore (Double)
Private Function CalculateBestPMIPair(ByVal pairCounts As Object, ByVal tokenCounts As Object, Optional ByVal minFreq As Long = 2) As Variant
  Dim totalPairs As Double
  totalPairs = 0
  Dim bestPair As Variant
  bestPair = Empty
  Dim maxPMI As Double: maxPMI = -1

  ' 全ペアの総和（確率計算用）
  Dim k As Variant
  For Each k In pairCounts.Keys
    totalPairs = totalPairs + CDbl(pairCounts(k))
  Next k

  For Each k In pairCounts.Keys
    Dim countAB As Double: countAB = CDbl(pairCounts(k))
    ' 低頻度すぎるペアは無視（PMIのバイアス対策）
    If countAB < minFreq Then GoTo NextPair

    Dim parts() As String
    parts = Split(CStr(k), "|")
    Dim a As String: a = parts(0)
    Dim b As String: b = parts(1)

    If Not tokenCounts.Exists(a) Or Not tokenCounts.Exists(b) Then GoTo NextPair

    Dim countA As Double: countA = CDbl(tokenCounts(a))
    Dim countB As Double: countB = CDbl(tokenCounts(b))

    If countA <= 0 Or countB <= 0 Then GoTo NextPair

    ' P(a, b) / (P(a) * P(b))
    ' 実際には log2( (count_ab * total) / (count_a * count_b) )
    Dim pmi As Double
    pmi = (Log((countAB * totalPairs) / (countA * countB))) / Log(2#)
    
    ' PPMI (Positive PMI): 負の値は0とする
    Dim ppmi As Double
    If pmi < 0 Then
      ppmi = 0
    Else
      ppmi = pmi * (Log(countAB) / Log(2#))
    End If

    If ppmi > maxPMI Then
      maxPMI = ppmi
      Dim pairArr(1) As String
      pairArr(0) = a
      pairArr(1) = b
      bestPair = pairArr
    End If
NextPair:
  Next k

  Dim ret(1) As Variant
  ret(0) = bestPair
  ret(1) = maxPMI
  CalculateBestPMIPair = ret
End Function

' コーパス内で指定ペアを統合する関数
' 引数: pair (2要素の配列)、corpusDict (Scripting.Dictionary)
' 戻り値: newCorpus (Scripting.Dictionary)
Private Function MergeVocab(ByVal pair As Variant, ByVal corpusDict As Object) As Object
  Dim newCorpus As Dictionary
  Set newCorpus = New Dictionary

  Dim bigram As String
  bigram = pair(0) & " " & pair(1)
  Dim replacement As String
  replacement = pair(0) & pair(1)

  Dim s As Variant
  For Each s In corpusDict.Keys
    Dim newSentence As String
    newSentence = Replace(CStr(s), bigram, replacement)
    If newCorpus.Exists(newSentence) Then
      newCorpus(newSentence) = newCorpus(newSentence) + CLng(corpusDict(s))
    Else
      newCorpus(newSentence) = CLng(corpusDict(s))
    End If
  Next s

  Set MergeVocab = newCorpus
End Function

' TrainVocab: コーパス文字列配列を受け取り学習を行う
' 引数: sentencesArr (1次元のVariant配列またはCollection)、numMerges (Long)、minFreq (Long)
' 戻り値: finalCorpus (Scripting.Dictionary) - 学習後の文->頻度辞書
Public Function TrainVocab(ByVal sentencesArr As Variant, Optional ByVal numMerges As Long = 50, Optional ByVal minFreq As Long = 2) As Dictionary
  Dim corpus As Dictionary
  Set corpus = New Dictionary

  Dim i As Long
  ' 初期コーパス集計
  For i = LBound(sentencesArr) To UBound(sentencesArr)
    Dim s As String: s = CStr(sentencesArr(i))
    corpus(SplitIntoChars(s)) = 1
  Next i

  Dim merges As Collection
  Set merges = New Collection

  ' 学習ループ
  Dim iter As Long
  For iter = 1 To numMerges
    Dim pairCounts As Dictionary, tokenCounts As Dictionary
    Set pairCounts = New Dictionary
    Set tokenCounts = New Dictionary
    Call GetStats(corpus, pairCounts, tokenCounts)

    Dim res As Variant
    res = CalculateBestPMIPair(pairCounts, tokenCounts, minFreq)
    Dim bestPair As Variant: bestPair = res(0)
    Dim score As Double: score = CDbl(res(1))

    If IsEmpty(bestPair) Then
      Exit For
    End If

    merges.Add bestPair
    Dim newCorpus As Object
    Set newCorpus = MergeVocab(bestPair, corpus)

    ' 変化がなければ終了
    ' TODO: 本来のアルゴリズムでは終了したいが，Dictionaryの一致判定が難しいため保留
    ' If newCorpus.Count = corpus.Count Then
    '   Dim same As Boolean: same = True
    '   Dim key As Variant
    '   For Each key In corpus.Keys
    '     If Not newCorpus.Exists(key) Or newCorpus(key) <> corpus(key) Then
    '       same = False
    '       Exit For
    '     End If
    '   Next key
    '   If same Then Exit For
    ' End If

    Set corpus = newCorpus
  Next iter

  Set TrainVocab = corpus
End Function

' 学習済みコーパスから各トークンの頻度辞書を抽出して頻度順降順で返す
' 引数: corpusDict (Scripting.Dictionary) - 学習後の文->頻度辞書
' 戻り値: sortedTokenDict (Scripting.Dictionary) - 頻度順に挿入された辞書
Public Function ExtractSortedTokens(ByVal corpusDict As Dictionary) As Dictionary
  Dim tokenCounts As Dictionary
  Set tokenCounts = New Dictionary

  ' 学習した単語単位に分割し，それぞれの単語の個数をカウント
  Dim s As Variant
  For Each s In corpusDict.Keys
    Dim freq As Long: freq = CLng(corpusDict(s))
    Dim tokens() As String
    tokens = Split(CStr(s), " ")
    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
      If tokenCounts.Exists(tokens(i)) Then
        tokenCounts(tokens(i)) = tokenCounts(tokens(i)) + freq
      Else
        tokenCounts(tokens(i)) = freq
      End If
    Next i
  Next s

  ' ソート用の配列を作成
  Dim arr() As Variant
  ReDim arr(0 To tokenCounts.Count - 1) As Variant
  Dim idx As Long: idx = 0
  Dim k As Variant
  For Each k In tokenCounts.Keys
    Dim pair(1) As Variant
    pair(0) = k
    pair(1) = CLng(tokenCounts(k))
    arr(idx) = pair
    idx = idx + 1
  Next k

  ' 降順ソート (単純なクイックソート)
  If UBound(arr) > 0 Then QuickSortPairs arr, 0, UBound(arr)

  Dim sortedDict As Dictionary
  Set sortedDict = New Dictionary
  Dim j As Long
  For j = 0 To UBound(arr)
    sortedDict(arr(j)(0)) = arr(j)(1)
  Next j

  Set ExtractSortedTokens = sortedDict
End Function

' QuickSort ヘルパー: arr の各要素は二要素配列 (token, freq)
Private Sub QuickSortPairs(ByRef arr As Variant, ByVal low As Long, ByVal high As Long)
  Dim i As Long, j As Long
  Dim pivot As Long
  i = low: j = high
  pivot = arr((low + high) \ 2)(1)
  Dim tmp As Variant
  Do While i <= j
    Do While arr(i)(1) > pivot
      i = i + 1
    Loop
    Do While arr(j)(1) < pivot
      j = j - 1
    Loop
    If i <= j Then
      tmp = arr(i)
      arr(i) = arr(j)
      arr(j) = tmp
      i = i + 1
      j = j - 1
    End If
  Loop
  If low < j Then QuickSortPairs arr, low, j
  If i < high Then QuickSortPairs arr, i, high
End Sub