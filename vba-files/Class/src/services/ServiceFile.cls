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
  ' TODO: 実装
End Function