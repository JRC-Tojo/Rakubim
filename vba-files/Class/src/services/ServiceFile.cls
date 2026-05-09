VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServiceFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ServiceFile.LoadFiles: 指定したファイル一覧を FileObj の ObjList で返す
' 引数: analyzer (ServiceNameAnalyzer) - ファイル名解析サービス
'       filePaths (ObjList(ObjPath)) - ファイルオブジェクトを構築するファイルパス一覧
'       useSeq (Boolean) - ファイル名の解析を連番モードで行うか
' 戻り値: ObjResult - 成功時は ObjList(FileObj) を返す
Public Function LoadFiles(analyzer As ServiceNameAnalyzer, filePaths As ObjList, useSeq As Boolean) As ObjResult
  ' ObjPath の情報を ServiceNameAnalyzer を使って解析し， ModelFile に変換する
  Dim fileColl As ObjList
  Set fileColl = buildModelFiles(analyzer, filePaths, useSeq)
  
  Set LoadFiles = Success(fileColl)
End Function

' 指定フォルダから指定拡張子のファイルパス一覧を返す
' 引数: folderPath (String) - 検索するフォルダのパス
'       extensions (ObjSet) - 拡張子文字列の配列（例: ObjSet("pdf","dwg")）
'       Optional recursive (Boolean) - サブフォルダを再帰的に検索するかどうか
' 戻り値: ObjResult - 成功時は ObjList(ObjPath) を返す
Public Function LoadPaths(folderPath As ObjPath, extensions As ModelExts, Optional recursive As Boolean = False) As ObjResult
  ' フォルダの存在チェック
  If (Not folderPath.Exists() Or folderPath.IsFile) Then
    Set LoadPaths = Failure("service.file.load.folder_not_found", folderPath)
    Exit Function
  End If

  ' フォルダ内のファイルを収集し，ObjList(ObjPath) を返す
  Dim pathColl As ObjList
  Set pathColl = collectFilesInFolder(folderPath, extensions, recursive)
  If (pathColl.Count = 0) Then
    Set LoadPaths = Failure("service.file.load.empty_folder", folderPath)
    Exit Function
  End If

  Set LoadPaths = Success(pathColl)
End Function

' Private helper: 指定フォルダ内のファイルを収集して ObjPath ObjList を返す
Private Function collectFilesInFolder(folder As ObjPath, exts As ModelExts, recursive As Boolean) As ObjList
  Dim readFiles As ObjList
  Set readFiles = New ObjList

  Dim f As ObjPath
  For Each f In folder.Colldir()
    If (f.IsFile And exts.ExtSet.Contains(f.Extension)) Then
      Call readFiles.Append(f)
    ElseIf (recursive And Not f.IsFile) Then
      ' サブフォルダの場合、再帰的に収集
      Dim subFiles As ObjList
      Set subFiles = collectFilesInFolder(f, exts, True)
      Dim sf As ObjPath
      For Each sf In subFiles
        Call readFiles.Append(sf)
      Next sf
    End If
  Next f

  Set collectFilesInFolder = readFiles
End Function

' Private helper: パス情報からファイル名を解析して ModelFile を生成する
' 引数:  analyzer (ServiceNameAnalyzer) - ファイル名解析サービス
'        ps (ObjList of ObjPath) - 解析対象のファイルパス情報コレクション
' 戻り値: ObjList of ModelFile
Private Function buildModelFiles(analyzer As ServiceNameAnalyzer, ps As ObjList, useSeq As Boolean) As ObjList
  Dim loadedFiles As ObjList
  Set loadedFiles = New ObjList

  ' ObjPath -> ModelFile に変換
  Dim p As ObjPath
  Dim res As ObjResult
  Dim info As ModelAnalyzed
  Dim mf As ModelFile
  For Each p In ps
    Set res = analyzer.AnalyzeName(p.Stem, useSeq)
    ' 解析成功時のみ ModelFile を生成
    If res.isOK Then
      Set info = res.value
      Set mf = New ModelFile
      Call mf.InitModel(p, info.AnalyzedName, info.SnippedName, info.PaperNumbers)
      Call loadedFiles.Append(mf)
    End If
  Next p

  Set buildModelFiles = loadedFiles
End Function

' ============================================

' ファイル群に対する新規名称を返す
' 引数：papers (ObjList[ModelPaper])
'       files (ObjList[ObjPath])
' 戻り値：ObjResult[Dict[oldName: NewName]]
Public Function NewNames(papers As ObjList, files As ObjList, companyId As String, projectId As String, structureId As String) As ObjResult
  If (papers.Count <> files.Count) Then
    Set NewNames = Failure("service.file.newname.missmatch")
    Exit Function
  End If
  If (companyId = "") Then
    Set NewNames = Failure("service.file.newname.company")
    Exit Function
  End If
  If (projectId = "") Then
    Set NewNames = Failure("service.file.newname.project")
    Exit Function
  End If
  If (structureId = "") Then
    Set NewNames = Failure("service.file.newname.structure")
    Exit Function
  End If

  Dim fileToPapers As ObjDict
  Set fileToPapers = New ObjDict
  Dim oldToNew As ObjDict
  Set oldToNew = New ObjDict

  ' ファイル名別に図面名のリストを作成
  Dim fName As ObjPath
  Dim paper As ModelPaper
  Dim emptyPapers As ObjList
  Dim i As Long
  For i = 0 To papers.Count
    fName = files.At(i).Name
    paper = papers.At(i)
    If (fileToPapers.Exists(fName)) Then 
      Call fileToPapers.At(fName).Append(paper)
    Else
      Set emptyPapers = New ObjList
      Call emptyPapers.InitItems(paper)
      Call fileToPapers.Append(fName, emptyPapers)
    End If
  Next i

  ' 定数宣言
  Dim newNameHeader As String
  newNameHeader = companyId & "-" & projectId & "-" & structureId
  Dim fileExt As String
  fileExt = files.At(0).Extension

  ' ファイル別に新ファイル名を作成
  Dim fpsPair As ModelKeyValue
  For Each fpsPair In fileToPapers
    Call oldToNew.Append(fpsPair.Key, CreateFileNameFromPapers(fpsPair.Value, newNameHeader, fileExt))
  Next

  Set NewNames = Success(oldToNew)
End Function


' 新規ファイル名を作成
Private Function CreateFileNameFromPapers(papers As ObjList, header As String, fileExt As String) As String
  Dim paperIdxs As String
  Dim fileTitle As String
  Dim paperNums As String
  paperIdxs = FormattedPaperIdxsFromPapers(papers)
  fileTitle = FormattedTitleFromPapers(papers)
  paperNums = FormattedPaperNumsFromPapers(papers)

  Dim newFileName As String
  newFileName = header & "-" & paperIdxs & "_" & fileTitle & paperNums & "." & fileExt

  CreateFileNameFromPapers = newFileName
End Function


Private Function FormattedTitleFromPapers(papers As ObjList) As String
  ' 図面名の重複を除外しながらユニークな図面名のリストを作成
  Dim uniquePaperNames As ObjSet
  Set uniquePaperNames = New ObjSet
  Dim paper As ModelPaper
  For Each paper In papers
    Call uniquePaperNames.Append(TitleFormatter(paper.SnippedName))  
  Next

  FormattedTitleFromPapers = uniquePaperNames.Join(", ")
End Function


' 新規図面に入るタイトル文字のフォーマット
Private Function TitleFormatter(title As String)
  Dim i As Integer
  For i = 1 To Len(title)
    If Mid(title, i, 1) Like "[ａ-ｚＡ-Ｚ０-９　]" Then
      Mid(title, i, 1) = StrConv(Mid(title, i, 1), vbNarrow)
    End If
  Next
  TitleFormatter = title
End Function


Private Function FormattedPaperIdxsFromPapers(papers As ObjList) As String
  Dim paperIdxs As ObjList
  Set paperIdxs = New ObjList
  Dim paper As ModelPaper
  For Each paper In papers
    Call paperIdxs.Append(paper.PaperIndex)
  Next

  ' PaperIndexが非対応の値の場合は最終表現から除外する
  Dim ignoreIdxs As ObjSet
  Set ignoreIdxs = New ObjSet
  Call ignoreIdxs.InitItems(-1)

  ' 数値の連続性でグループ分け
  Dim groupedIdxs As ObjList
  Set groupedIdxs = groupConsecutiveNumbers(paperIdxs, ignoreIdxs)

  ' 数値が存在しない場合は無効な値として返す
  If (groupedIdxs.Count = 0) Then
    FormattedPaperIdxsFromPapers = "XXXX"
    Exit Function
  End If

  ' 各数値を４桁パディングでフォーマット
  Dim formattedGroupIdx As ObjList
  Dim formattedGroupIdxs As ObjList
  Set formattedGroupIdxs = New ObjList
  Dim groupedIdx As ObjList
  Dim pIdx As Long
  For Each groupedIdx In groupedIdxs
    Set formattedGroupIdx = New ObjList
    For Each pIdx In groupedIdx
      Call formattedGroupIdx.Append(ToPaddedString(pIdx, 4))
    Next
    Call formattedGroupIdxs.Append(formattedGroupIdx)
  Next

  ' 連続/不連続状態を略記で表現
  Dim formattedPaperIdxs As String
  formattedPaperIdxs = formatGroupedConsecutiveNumbers(formattedGroupIdxs, ",", "-")

  FormattedPaperIdxsFromPapers = formattedPaperIdxs
End Function


Private Function FormattedPaperNumsFromPapers(papers As ObjList) As String
  Dim paperNums As ObjList
  Set paperNums = New ObjList
  Dim paper As ModelPaper
  For Each paper In papers
    Call paperNums.Append(paper.PaperNumber)
  Next

  ' PaperNumberが非対応の値の場合は最終表現から除外する
  Dim ignoreNums As ObjSet
  Set ignoreNums = New ObjSet
  Call ignoreNums.InitItems(-1)

  ' 数値の連続性でグループ分け
  Dim groupedNums As ObjList
  Set groupedNums = groupConsecutiveNumbers(paperNums, ignoreNums)

  ' 数値が存在しない場合は表示すべき値がないものとして返す
  If (groupedNums.Count = 0) Then
    FormattedPaperNumsFromPapers = ""
    Exit Function
  End If

  ' 各数値を全角にフォーマット
  Dim formattedGroupNum As ObjList
  Dim formattedGroupNums As ObjList
  Set formattedGroupNums = New ObjList
  Dim groupedNum As ObjList
  Dim pNum As Long
  For Each groupedNum In groupedNums
    Set formattedGroupNum = New ObjList
    For Each pNum In groupedNum
      Call formattedGroupNum.Append(StrConv(CStr(pNum), vbWide))
    Next
    Call formattedGroupNums.Append(formattedGroupNum)
  Next

  ' 連続/不連続状態を略記で表現
  Dim formattedPaperNums As String
  formattedPaperNums = formatGroupedConsecutiveNumbers(formattedGroupIdxs, "、その", "～その")

  FormattedPaperNumsFromPapers = "（その" & formattedPaperNums & "）"
End Function


' 数値配列のうち，連続する部分を配列でくくることにより，連続/非連続の別による２次元配列を作成する
' 引数：targets (ObjList) 数値配列
'       ignores (ObjSet) 無視する数字を定義可能
' 戻り値：ObjList (ex: [1, 2, 3, 6, 8] -> [[1, 2, 3], [6], [8]])
Private Function groupConsecutiveNumbers(targets As ObjList, Optional ignores As ObjSet = Nothing) As ObjList
  If (targets.Count = 0) Then
    Set groupConsecutiveNumbers = targets
    Exit Function
  End If

  Dim thisIgnores As ObjSet
  If (thisIgnores Is Nothing) Then 
    Set thisIgnores = New ObjSet
  Else
    Set thisIgnores = ignores
  End If

  Dim splitted As ObjList
  Set splitted = New ObjList
  Dim currentGroup As ObjList
  Set currentGroup = New ObjList

  Dim target As Long
  For Each target In targets
    ' 無視する数字の場合
    If (thisIgnores.Contains(target)) Then
      ' 現在のグループに中身があれば結果に追加してリセット
      If (currentGroup.Count > 0) Then
        Call splitted.Append(currentGroup)
        Set currentGroup = New ObjList
      End If
    Else
      ' グループが空なら現在の数字を追加
      If (currentGroup.Count = 0) Then
        Call currentGroup.Append(target)
      Else
        ' 1つ前の数字と連続しているかチェック
        If (target = currentGroup.At(-1)) Then
          Call currentGroup.Append(target)
        Else
          ' 連続していない場合、今のグループを結果に入れて新しいグループを開始
          Call splitted.Append(currentGroup)
          Set currentGroup = New ObjList
          Call currentGroup.InitItems(target)
        End If
      End If
    End If
  Next

  ' 最後に残ったグループがあれば追加
  If (currentGroup.Count > 0) Then
    Call splitted.Append(currentGroup)
  End If

  Set groupConsecutiveNumbers = splitted
End Function


' 連続状態でグループ分けした２次元配列をフォーマット
' 引数：groupedNums (ObjList[String]) 連続状態で２次元化された配列．内部の値はフォーマット済みの文字列とする．
'       divTxt (String) 不連続な箇所に挿入する文字列
'       conTxt (String) 連続な箇所に挿入する文字列
' 戻り値：(String) フォーマット済みの文字列 (ex: ([[1, 2, 3], [6], [8]], "|", "-") -> "1-3|6|8")
Private Function formatGroupedConsecutiveNumbers(groupedNums As ObjList, divTxt As String, conTxt As String) As String
  Dim formattedCons As ObjList
  Set formattedCons = New ObjList

  Dim innerList As ObjList
  For Each innerList In groupedNums
    If (innerList.Count = 1) Then
      Call formattedCons.Append(innerList.At(0))
    ElseIf (innerList.Count > 1) Then
      Call formattedCons.Append(innerList.At(0) & conTxt & innerList.At(-1))
    End If
  Next

  formatGroupedConsecutiveNumbers = formattedCons.Join(divTxt)
End Function


' 数値を指定された桁数の文字列に変換する（ゼロ埋め）
' @param num    対象の数値
' @param digits 指定する桁数
' @return       フォーマットされた文字列
Private Function ToPaddedString(ByVal num As Variant, ByVal digits As Integer) As String
  ToPaddedString = Format(num, String(digits, "0"))
End Function