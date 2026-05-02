VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServiceNameAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ServiceNameAnalyzer: 既存の `Analyzer` ロジックを流用し
' 図面名・ファイル名を解析して `AnalyzeName` / `SnippedName` / `PaperNumbers` を返すラッパ

' 与えられた名称から得られる解析結果を返す
' 引数: stemName (String) - ファイル名は拡張子抜きの名称，図面名称はそのままの名称
'       useContinuousNumber (Boolean) - 連番を含めるかどうかのフラグ
' 戻り値: ObjResult - 成功時は ModelAnalyzed オブジェクトを返す
Public Function AnalyzeName(stemName As String, useContinuousNumber As Boolean) As ObjResult
	On Error GoTo ErrHandler

	Dim analyzedName As String
	Dim snippedName As String
	Dim paperNumbers As ObjSet

  ' ファイル名から図面名称にあたる部分を抽出
	snippedName = LoadSnippedName(stemName)
  If (snippedName = "") Then
    Set AnalyzeName = Failure("repo.nameanalyzer.snippedname.empty", stemName)
    Exit Function
  End If

  ' 抽出した図面名称をマッチング用に正規化
	analyzedName = LoadAnalyzedName(snippedName)
  ' ファイル名から図面の連番情報を抽出
	Set paperNumbers = LoadPaperNumber(stemName, useContinuousNumber)

	' 解析結果を登録
	Dim ma As New ModelAnalyzed
	Call ma.InitModel(analyzedName, snippedName, paperNumbers)

	Set AnalyzeName = Success(ma)
	Exit Function
ErrHandler:
	Set AnalyzeName = Failure("repo.nameanalyzer.analyzename.error", stemName, Err.Number, Err.Description)
End Function


' Private helper: 指定名称から図面名称を取得
Private Function LoadSnippedName(name As String) As String
  Dim snippedName As String: snippedName = name
  Dim reg As ObjRegex: Set reg = New ObjRegex
  Dim m As Match

  ' 失敗時の戻り値を定義
  LoadSnippedName = ""

  ' 冒頭部分のマッチング状況を取得
  Dim headMatchs As ObjList
  Set headMatchs = calcHeadNumberMatch(name)
  ' 後ろから順に削除することでMatchのインデックスのズレを防止する
  For Each m In headMatchs.Reverse()
    snippedName = reg.DropMatchedPosition(m, snippedName)
  Next

  ' 末尾の図面番号部分のマッチング状況を取得
  Dim lastNumberMatches As ObjList
  Set lastNumberMatches = calcEndNumberMatch(snippedName)
  For Each m In lastNumberMatches.Reverse()
    snippedName = reg.DropMatchedPosition(m, snippedName)
  Next

  ' 図面名称の前後から連結文字を削除
  snippedName = trimJointChrs(snippedName)

  LoadSnippedName = snippedName
End Function


' Private helper: 指定名称から正規化した解析用名称を取得
Private Function LoadAnalyzedName(snippedName As String) As String
	Dim normSnippedName As String
  
  ' 大文字・全角に正規化
  normSnippedName = StrConv(StrConv(snippedName, vbUpperCase), vbWide)
  ' 連結文字をすべて削除
  normSnippedName = trimJointChrs(normSnippedName, False)
  
  LoadAnalyzedName = normSnippedName
End Function


' Private helper: 指定名称から図面の連番情報を取得
' 引数: name (String) - 解析対象の名称
'       useContinuousNumber (Boolean) - 連番を含めるかどうかのフラグ
' 戻り値: ObjSet - 図面の連番情報を格納したコレクション
Private Function LoadPaperNumber(name As String, useContinuousNumber As Boolean) As ObjSet
  ' 「末尾の数字」に該当する範囲を返す
  Dim matchedNumbers As ObjList
  Set matchedNumbers = calcEndNumberMatch(name)

  ' マッチが存在しない場合は空のコレクションを返す
  Set LoadPaperNumber = New ObjSet
  If (matchedNumbers Is Nothing) Then Exit Function
  If (matchedNumbers.Count = 0) Then Exit Function

  ' 抽出した範囲の文字列から数値配列を解析する
  Dim interpretedNumbers As ObjSet: Set interpretedNumbers = New ObjSet
  If (useContinuousNumber) Then
    Set interpretedNumbers = interpretedNumbers.Union(interpretStringWithNumbers(matchedNumbers.At(-1)))
  Else
    Dim matchedNumber As Match
    For Each matchedNumber In matchedNumbers
      Set interpretedNumbers = interpretedNumbers.Union(interpretStringWithNumbers(matchedNumber))
    Next
  End If

  Set LoadPaperNumber = interpretedNumbers
End Function


' Private helper: 指定名称から冒頭の連番マッチを取得
Private Function calcHeadNumberMatch(ByVal name As String) As ObjList
  Dim reg As ObjRegex: Set reg = New ObjRegex

  ' 半角文字に正規化
  Dim normName As String
  normName = StrConv(name, vbNarrow)
  
  ' 冒頭の連番を抽出
  Dim mc As ObjList
  Set mc = reg.GetMatches("^([0-9]+[_,\s~-]*)+", normName)
  
  Set calcHeadNumberMatch = mc
End Function


' Private helper: 指定名称から末尾の番号（その〇）を取得
' サポートしている末尾の番号のフォーマットはテスト`TestServiceNameAnalyzer.bas`を参照
Private Function calcEndNumberMatch(ByVal name As String) As ObjList
  Dim reg As ObjRegex: Set reg = New ObjRegex

  ' 全角文字に正規化
  Dim normName As String
  normName = StrConv(name, vbWide)
  
  ' 末尾の連番を抽出
  ' --- 1. 基本パーツの定義 ---
  Dim innerDigit As String
  Dim innerSep As String
  innerDigit = "０-９"    ' 数字（全角）
  innerSep = "～－＿，、　"  ' 区切り記号と空白

  ' --- 2. 各構成要素の組み立て ---
  Dim cDigit As String:  cDigit = "[" & innerDigit & "]"
  Dim cSep As String:    cSep = "[" & innerSep & "]"
  ' 接頭辞：数字・記号・カッコ「以外」の文字を0～5文字許容
  ' [^ ... ] の形式を動的に作成
  Dim cPrefix As String
  cPrefix = "[^" & innerDigit & innerSep & "（）" & "]{0,5}"
  
  ' --- 3. 全体パターンの構築 ---
  ' カッコ内の基本構造： (接頭辞+数字) + (記号+接頭辞+数字)の繰り返し
  Dim unitPattern As String
  unitPattern = cPrefix & cDigit & "+"
  ' カッコ内の全体構造：上記の基本構造 + それが繰り返されるパターン
  Dim sequencePattern As String
  sequencePattern = unitPattern & "(?:" & cSep & "+" & unitPattern & ")*"
  ' 単独数字のパターン：数字のみ
  Dim onlyDigitPattern As String
  onlyDigitPattern = cDigit & "+"
  ' 最終形：「（内容）」「（内容）～（内容）」「（単独数字）」をサポート
  Dim finalPattern As String
  finalPattern = "（" & sequencePattern & "）" & _
                 "(?:" & cSep & "（" & sequencePattern & "）" & ")?" & "|" & _
                 onlyDigitPattern & "$"
  
  Dim mc As ObjList
  Set mc = reg.GetMatches(finalPattern, normName)
  
  Set calcEndNumberMatch = mc
End Function


' 入力文字列の前後端部の連結文字を削除
' 引数: name (String) - 連結文字を削除したい名称
'       onlyFirstAndLast (Boolean) - 先頭と末尾の連結文字のみ削除するかどうか（デフォルトTrue）
' 戻り値: String - 連結文字を削除した名称
Private Function trimJointChrs(ByVal name As String, Optional onlyFirstAndLast As Boolean = True) As String
  Dim reg As ObjRegex: Set reg = New ObjRegex

  ' 半角文字に正規化
  Dim normName As String
  normName = StrConv(name, vbNarrow)

  ' 連結文字の位置を抽出
  Dim trimTargets As ObjList
  Set trimTargets = reg.GetMatches("[,_\s-]+", normName)
  
  ' 抽出した位置の文字を元の名称から削除
  Dim i As Long
  Dim targetChr As Match
  Dim trimedName As String: trimedName = name
  Dim trimHitAll As Boolean: trimHitAll = (Not onlyFirstAndLast)
  For Each targetChr In trimTargets.Reverse()
    ' 連結文字が先頭または末尾に存在する場合は削除する
    Dim isFirstChr As Boolean: isFirstChr = (targetChr.FirstIndex = 1)
    Dim isLastChr As Boolean: isLastChr = (targetChr.FirstIndex + targetChr.Length = Len(trimedName))
    If trimHitAll Or isFirstChr Or isLastChr Then
      trimedName = reg.DropMatchedPosition(targetChr, trimedName)
    End If
  Next targetChr

  trimJointChrs = trimedName
End Function


' 指定された数値入り名称の意味を解釈してその名称が持つ数字配列を返す
' ex) "その１～３" -> ObjSet(1,2,3)
'     "１-３" -> ObjSet(1,2,3)
'     "1, 3, 5" -> ObjSet(1,3,5)
Private Function interpretStringWithNumbers(ByVal matchedItem As Match) As ObjSet
  Dim reg As ObjRegex: Set reg = New ObjRegex

  ' 半角文字に正規化
  Dim normName As String
  normName = StrConv(matchedItem.Value, vbNarrow)
  ' 処理に関係ない文字はすべて落とす
  normName = reg.Replace("[^0-9,~\-]", normName, "")

  ' 数字部分を抽出
  Dim matchedNums As ObjList
  Set matchedNums = reg.GetMatches("\d+", normName)

  ' ５桁より大きい数字は260101のような日付の可能性が高いため除外する
  Dim i As Long
  Dim matchedNum As Match
  For i = 0 To matchedNums.Count - 1
    Set matchedNum = matchedNums.At(i)
    If matchedNum.Length > 5 Then Call matchedNums.Pop(i)
  Next

  ' 抽出した数字の意味を解釈して配列に格納
  ' 数字がない場合は空配列を返すため処理不要
  Dim interprettedNums As New ObjSet
  If matchedNums.Count = 1 Then
    ' 数字が一つだけの場合はその数字を返す
    Call interprettedNums.Append(CLng(matchedNums.At(0).Value))
  ElseIf matchedNums.Count > 1 Then
    ' 数字が複数ある場合は範囲指定されていると考えて処理する
    Set interprettedNums = interpretRangeString(normName)
    ' 範囲指定の数字が見えない場合は個別の数字をすべて返す
    If interprettedNums.Count = 0 Then
      For Each matchedNum In matchedNums
        Call interprettedNums.Append(CLng(matchedNum.Value))
      Next matchedNum
    End If
  End If

  Set interpretStringWithNumbers = interprettedNums
End Function


' 指定された文字列から番号範囲を示す文字列の意味を解釈して数値配列(ObjList)を返す
Private Function interpretRangeString(ByVal rangeStr As String) As ObjSet
  Dim reg As ObjRegex: Set reg = New ObjRegex
  Dim interprettedNums As ObjSet: Set interprettedNums = New ObjSet
  Dim matchedRange As Match
  Dim matchedRanges As ObjList
  Dim rangeParts As ObjList

  Dim i As Long
  Dim startNum As Long
  Dim endNum As Long
  
  Set matchedRanges = reg.GetMatches("[0-9]+[~-][0-9]+", rangeStr)
  For Each matchedRange In matchedRanges
    ' 範囲指定がある場合は範囲内の数字をすべて返す
    Set rangeParts = reg.GetMatches("[0-9]+", matchedRange.Value)
    startNum = CLng(rangeParts.At(0).Value)
    endNum = CLng(rangeParts.At(1).Value)
    For i = startNum To endNum
      Call interprettedNums.Append(i)
    Next i
  Next

  Set interpretRangeString = interprettedNums
End Function