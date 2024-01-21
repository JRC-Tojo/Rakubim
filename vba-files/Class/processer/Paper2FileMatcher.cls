VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Paper2FileMatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private inner_cleaner As FileNameCleaner
Private inner_SystemIds As SystemIdObj
Private inner_ContinuousPDF As FileObj
Private inner_PDFs As Collection
Private inner_DWGs As Collection

Private Sub class_initialize()
  Set inner_SystemIds = New SystemIdObj

  ' 図面目録で重複しているワードを除去するためのクリーナーを定義
  Dim sheet As SheetObj
  sheet = New SheetObj
  Call inner_cleaner.AnalyzeNames(Range2Collection(sheet.MokurokuPage, "B2:B10000"))

  ' ファイル群を取得
  Set inner_PDFs = GetFileCollection("PDF")
  Set inner_DWGs = GetFileCollection("CAD")

  ' PDF連番処理の場合の準備
  If (inner_SystemIds.UseContinuousNumber And inner_PDFs.Count > 0) Then
    Set inner_ContinuousPDF = SearchMostNumbersFile(inner_PDFs)
  End If
End Sub


' 引数で受けた図面に最も適当なファイルを探索する
Public Function GetBestFile(paper As PaperObj, fileType As String)
  if (systemIds.UseContinuousNumber And fileType = "PDF") Then
    GetBestFile = ApplyContinuousFile(paper.ContinuousNumber)
  Else
    Set files = IIf(fileType = "PDF", inner_PDFs, inner_DWGs)
    GetBestFile = getMatchPath(files, paper)
  End if
End Function


' 指定したpaperInfoに最もマッチするFileObjをobjsに指定したCollectionから探索する
Private Function getMatchFile(objs As Collection, paperInfo As PaperObj, cleaner As FileNameCleaner) As FileObj
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
Private Function getMatchPath(files As Collection, paperInfo As PaperObj, cleaner As FileNameCleaner)
  If (files.Count = 0) Then
    getMatchPath = ""
  End If
  
  Dim resultFile As FileObj
  Set resultFile = getMatchFile(files, paperInfo, cleaner)
  
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
Private Function ApplyContinuousFile(indexCounter As Integer)
  If inner_ContinuousPDF.PaperPaths.Exists("idx" & indexCounter) = True Then
    ApplyContinuousFile = inner_ContinuousPDF.PaperPaths("idx" & indexCounter)
  ElseIf inner_ContinuousPDF.PaperPaths.Exists("idx-1") = True Then
    ApplyContinuousFile = inner_ContinuousPDF.PaperPaths("idx-1")
  Else
    ApplyContinuousFile = ""
  End If
End Function
