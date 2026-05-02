VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServiceNameCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ServiceNameCleaner: 既存の `FileNameCleaner` ロジックを流用し
' 図面名・ファイル名から頻出語彙を除去する

' 引数: pfs (ObjList of ModelPaper/ModelFile)
' 戻り値: ObjResult - 成功時は ObjList(ModelPaper/ModelFile) を返す
Public Function Clean(pfs As ObjList) As ObjResult
  On Error GoTo ErrHandler

  Dim cleanWords As ObjSet
  Set cleanWords = New ObjSet
  Dim outColl As ObjList
  Set outColl = New ObjList

  ' 除去単語リストを構築する
  Set cleanWords = buildCleanWords(extractAnalyzedNames(pfs))

  ' 除去リストの単語を削除する
  Dim model As Object
  For Each model In pfs
    model.AnalyzedName = cleanAnalyzedName(cleanWords, model.AnalyzedName)
    Call outColl.Append(model)
  Next

  Set Clean = Success(outColl)
  Exit Function
ErrHandler:
  Set Clean = Failure("service.namecleaner.clean.error", Err.Number, Err.Description)
End Function

' Private helper: ModelFile コレクションから AnalyzedName を抽出して配列で返す
Private Function extractAnalyzedNames(pfs As ObjList) As ObjSet
  Dim aNames As ObjSet
  Set aNames = New ObjSet

  Dim mf As Object
  For Each mf In pfs
    Call aNames.Append(mf.AnalyzedName)
  Next mf

  Set extractAnalyzedNames = aNames
End Function

' 除去リスト一覧を構築する
Private Function buildCleanWords(analyzedNames As ObjSet) As ObjSet
  ' 単語の抽出
  Dim bpe As BytePairEncoding
  Dim vocab As Dictionary
  Dim tokenToFreqs As Dictionary
  Set bpe = New BytePairEncoding
  Set vocab = bpe.TrainVocab(analyzedNames.ToArray())
  Set tokenToFreqs = bpe.ExtractSortedTokens(vocab)

  ' 頻出語（除去単語）とその他の語の境界を特定
  Dim otu As Otu
  Dim splitIndex As Long
  Set otu = New Otu
  splitIndex = otu.ArgSplit(tokenToFreqs.Items)

  ' 頻出語を除去単語リストに登録する
  Dim cleanWords As ObjSet
  Set cleanWords = New ObjSet
  Dim i As Long
  For i = 0 To splitIndex
    Call cleanWords.Append(tokenToFreqs.Keys(i))
  Next i

  ' 除去単語リスト内で重複部分がある語彙を除去
  Dim dropedCleanWords As ObjSet
  Set dropedCleanWords = New ObjSet
  Dim isOverlapped As Boolean
  Dim cleanWord As Variant
  For Each cleanWord In SortArrayByLength(cleanWords.ToList())
    isOverlapped = False

    ' すでに「残す」と決めたリストと比較
    Dim keptItem As Variant
    For Each keptItem In dropedCleanWords
      If HasCommonCharacter(CStr(cleanWord), CStr(keptItem)) Then
        isOverlapped = True
        Exit For
      End If
    Next keptItem
    
    ' 重複がなければ結果リストに追加
    If Not isOverlapped Then
      Call dropedCleanWords.Append(CStr(cleanWord))
    End If
  Next cleanWord
  
  Set buildCleanWords = dropedCleanWords
End Function


' 除去リストに存在する部分を削除する方法でAnalyzedNameを修正する
Private Function cleanAnalyzedName(cleanWords As ObjSet, aName As String) As String
  Dim cleanedName As String
  cleanedName = aName

  Dim arr As Variant
  arr = cleanWords.ToArray()
  Dim i As Long
  For i = LBound(arr) To UBound(arr)
    cleanedName = Replace(cleanedName, CStr(arr(i)), "")
  Next i
  
  cleanAnalyzedName = cleanedName
End Function

' 2つの文字列に共通の文字があるか判定する関数
Private Function HasCommonCharacter(str1 As String, str2 As String) As Boolean
  Dim k As Long
  For k = 1 To Len(str1)
    ' str1の各文字がstr2に含まれているかチェック
    If InStr(1, str2, Mid(str1, k, 1)) > 0 Then
      HasCommonCharacter = True
      Exit Function
    End If
  Next k
  HasCommonCharacter = False
End Function

' 配列を文字列の長さで降順ソートするサブ関数
Private Function SortArrayByLength(arr As ObjList) As ObjList
  Dim i As Long, j As Long
  Dim temp As String
  For i = 0 To arr.Count
    For j = i + 1 To arr.Count - 1
      If Len(arr.At(i)) < Len(arr.At(j)) Then
        temp = arr.At(i)
        Call arr.Update(i, arr.At(j))
        Call arr.Update(j, temp)
      End If
    Next j
  Next i
  Set SortArrayByLength = arr
End Function