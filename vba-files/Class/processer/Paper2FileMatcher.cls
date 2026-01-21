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

Private inner_Cleaner As FileNameCleaner
Private inner_SystemIds As SystemIdObj
Private inner_ContinuousPDF As FileObj
Private inner_PDFs As Collection
Private inner_DWGs As Collection

Public Function BaseInitMatcher(targetPath As String)
  Set inner_SystemIds = New SystemIdObj

  ' 処理対象とするファイル群を取得
  Set inner_PDFs = IIf(inner_SystemIds.SelectedFileType <> "CAD", GetFileCollection(targetPath, "PDF"), New Collection)
  Set inner_DWGs = IIf(inner_SystemIds.SelectedFileType <> "PDF", GetFileCollection(targetPath, "CAD"), New Collection)

  ' PDF連番処理の場合の準備
  If (inner_SystemIds.UseContinuousNumber And inner_PDFs.Count > 0) Then
    Set inner_ContinuousPDF = SearchMostNumbersFile()
  End If
End Function

' フォルダ内のPDF一覧を返す
Public Property Get loadedPdfPaths() As Collection
  Dim tmpCol As Collection
  Set tmpCol = New Collection
  
  ' FileObj -> File name
  Dim file As FileObj
  For Each file In inner_PDFs
    tmpCol.Add file.SourceName
  Next file

  ' 連番としてまとめられたファイル群も連結する
  If Not (inner_ContinuousPDF Is Nothing) Then
    Set tmpCol = CombineCols(tmpCol, dictValues(inner_ContinuousPDF.PaperPaths))
  End If
  
  Set loadedPdfPaths = tmpCol
End Property

' フォルダ内のDWG一覧を返す
Public Property Get loadedDwgPaths() As Collection
  Dim tmpCol As Collection
  Set tmpCol = New Collection
  
  ' FileObj -> File name
  Dim file As FileObj
  For Each file In inner_DWGs
    tmpCol.Add file.SourceName
  Next file

  Set loadedDwgPaths = tmpCol
End Property

' Matcherが初期化で必要とする情報を登録
Public Function InitMatcher(paperNames As Collection)
  Set inner_Cleaner = New FileNameCleaner

  ' 図面目録で重複しているワードを除去するためのクリーナーを定義
  Call inner_Cleaner.AnalyzeNames(paperNames)
End Function


' 読み込むべきフォルダが存在していることを確認
Public Property Get ExistFiles() As Boolean
  Dim pdfMode As Boolean: pdfMode = (inner_SystemIds.SelectedFileType = "PDF")
  Dim cadMode As Boolean: cadMode = (inner_SystemIds.SelectedFileType = "CAD")
  Dim hasPDFs As Boolean: hasPDFs = (inner_PDFs.Count > 0)
  Dim hasDWGs As Boolean: hasDWGs = (inner_DWGs.Count > 0)
  ExistFiles = (hasPDFs And hasDWGs) _
                Or (hasDWGs And cadMode) _
                Or (hasPDFs And pdfMode)
End Property


' 引数で受けた図面に最も適当なファイルを探索する
Public Function GetBestFile(paper As PaperObj, fileType As String, useContinuousNumber As Boolean)
  If (useContinuousNumber And fileType = "PDF") Then
    GetBestFile = ApplyContinuousFile(paper.continuousNumber)
  Else
    Dim files As Collection
    Set files = IIf(fileType = "PDF", inner_PDFs, inner_DWGs)
    GetBestFile = getMatchPath(files, paper)
  End If
End Function


' 指定したpaperに最もマッチするFileObjをobjsに指定したCollectionから探索する
Private Function getMatchFile(objs As Collection, paper As PaperObj) As FileObj
  Dim ratios As Collection
  Set ratios = New Collection
  Dim seq As SequenceMatcher
  Dim seg1 As String, seg2 As String
  Dim paperName As String
  Dim fileName As String
  
  ' すべてのファイルに対してマッチ率を算出する
  Dim file As FileObj
  For Each file In objs
    Set seq = New SequenceMatcher
    ' 目録名称から除外すべき名称とスペースを除去
    paperName = inner_Cleaner.NameCleaner(paper.AnalyzeName)
    fileName = inner_Cleaner.NameCleaner(file.AnalyzeName)
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

' 指定したpaperに最もマッチするファイル名を返す
Private Function getMatchPath(files As Collection, paper As PaperObj) As String
  If (files.Count = 0) Then
    getMatchPath = ""
    Exit Function
  End If
  
  Dim resultFile As FileObj
  Set resultFile = getMatchFile(files, paper)
  
  ' 元のファイル名を返す
  If resultFile.PaperPaths.Exists("idx" & paper.SpotNumber) = True Then
    getMatchPath = resultFile.PaperPaths("idx" & paper.SpotNumber)
  ElseIf resultFile.PaperPaths.Exists("idx-1") = True Then
    getMatchPath = resultFile.PaperPaths("idx-1")
  Else
    getMatchPath = ""
  End If
End Function



' PDFの連番処理用
' 最も番号を保持しているファイルオブジェクトを抽出することで、連番化されたファイル名称を導く
Private Function SearchMostNumbersFile()
  Dim fileCounts As Collection
  Set fileCounts = New Collection
  Dim file As FileObj
  
  ' すべてのファイルに対してファイル数を算出
  For Each file In inner_PDFs
    fileCounts.Add file.PaperNumbers.Count
  Next file
  
  ' 最もファイル数を持つファイルオブジェクトを返す
  Dim bestIdx As Integer
  bestIdx = ArgMax(fileCounts)(1)
  Set SearchMostNumbersFile = inner_PDFs.item(bestIdx)
End Function

' PDFの連番処理用
' 指定されたファイルオブジェクトから番号を抜き出して割り当てていく
Private Function ApplyContinuousFile(indexCounter As Integer)
  If (inner_ContinuousPDF Is Nothing) Then
    ApplyContinuousFile = ""
    Exit Function
  End If

  If inner_ContinuousPDF.PaperPaths.Exists("idx" & indexCounter) = True Then
    ApplyContinuousFile = inner_ContinuousPDF.PaperPaths("idx" & indexCounter)
  ElseIf inner_ContinuousPDF.PaperPaths.Exists("idx-1") = True Then
    ApplyContinuousFile = inner_ContinuousPDF.PaperPaths("idx-1")
  Else
    ApplyContinuousFile = ""
  End If
End Function
