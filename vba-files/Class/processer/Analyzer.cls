VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Analyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public inner_PaperNumbers As Collection
Public inner_SnippedName As String
Private inner_PaperPaths As Scripting.Dictionary
Private Const inner_NumberPattern = "[[（その|（][０-９]+）|[（その|（][０-９]+[～|－][０-９]+）|（その[０-９]+）[～|－]（その[０-９]+）]|[０-９]+$"
Private Const inner_RemoveNumbersPattern = "(((（その)|（)[０-９]+）|((（その)|（)[０-９]+[～|－][０-９]+）|（その[０-９]+）[～|－]（その[０-９]+）)|[０-９]+$"


Public Property Get PaperNumbers() As Collection
  Set PaperNumbers = inner_PaperNumbers
End Property

Public Property Get SnippedName() As String
  SnippedName = inner_SnippedName
End Property

Public Sub SetPaperInfo(paper As PaperObj)
  Call analyze(paper.SourceName)
End Sub

Public Sub SetFileInfo(file As FileObj)
  ' 拡張子を除外
  Dim stem As String
  stem = Left(file.SourceName, Len(file.SourceName) - 4)
  
  Call analyze(stem, file.FileType)
End Sub

' 名称を分析してクラスを作成
Private Sub analyze(ByVal path As String, Optional fileType As String = "")
  Set inner_PaperNumbers = getPaperNumber(path, fileType)
  inner_SnippedName = getBluePrintName(path)
End Sub

' ファイル名から図面の連番（その○○）を取得
' （図面番号、終了番号(only DWG)）を戻り値として返す
' 「図面番号」は連番や（その〇），（その〇～〇）の最初の数字，が入る
' 全角数字一つしか検知されない場合：その数字が図面の番号
' 全角数字が複数検知された場合：その数字は図面そのものに対して意味を持たない
' 半角数字が検知された場合：末尾の数字は連続図面の連番を示す（連番と「その〇」の対応関係は考慮しない）
Private Function getPaperNumber(ByVal fileName As String, fileType As String) As Collection
  Dim reg As RegExp
  Set reg = New RegExp
  Dim hasDateNum As Integer
  hasDateNum = 0
  Dim numMatches As MatchCollection
  Dim numCount As Integer
  Dim startNum As Integer: startNum = -1
  Dim endNum As Integer: endNum = -1
  Dim targetName As String
  ' 連番処理用（連番は「末尾数字」or「両括弧の中の数字」のみ）
  Const continuousNumbersPattern = "([０-９]+$)|(（[０-９]+）$)"
  Dim systemIds As SystemIdObj
  Set systemIds = New SystemIdObj
  
  ' 数字の抽出（ただし、図面名より後ろにある数字）
  Dim numReg As RegExp
  Set numReg = New RegExp
  Dim tmpMatches As MatchCollection
  With numReg
    .Pattern = IIf(systemIds.UseContinuousNumber And fileType = "PDF", continuousNumbersPattern, inner_NumberPattern)
    .Global = True
  End With
  ' ファイル名称の日本語以降を抜き出して情報抽出
  targetName = StrConv(fileName, vbWide)
  
  Set tmpMatches = numReg.Execute(targetName)
  ' 数値の候補を適切に抽出した後に、抽出文字列から数値を取得
  If tmpMatches.Count > 0 Then
    With reg
      .Pattern = "[０-９]+"
      .Global = True
    End With
    Set numMatches = reg.Execute(tmpMatches.item(0))
  
    ' 末尾が日付データの時は除外
    If CLng(numMatches.item(numMatches.Count - 1)) > 10000 Then
      hasDateNum = 1
    End If

    ' 数値の個数を記録
    numCount = numMatches.Count - hasDateNum
    
  ' 上記で抽出されなかった部分に入っている数値は見ない
  Else
    numCount = 0
  End If
  
  Dim sameNum As Integer
  ' 数値が何もない場合は何も返さない or 数値があっても「（その」がない場合はその数値に意味がないと考える
  If numCount = 0 Then
    Set getPaperNumber = RangeCollection(-1, -1)
  ' 数字が一つ見えた場合はその数字が図面そのものの番号である
  ElseIf numCount = 1 Then
    sameNum = CInt(numMatches.item(0))
    Set getPaperNumber = RangeCollection(sameNum, sameNum)
  ' PDFデータにおいて，数字が2つ見える場合は末尾の「1」が省略されたと考える
  ElseIf fileType <> "CAD" And numCount = 2 Then
    Set getPaperNumber = RangeCollection(-1, -1)
  ' PDFデータにおいて，数字が複数見える場合は末尾の数字が図面そのものの番号である
  ElseIf fileType <> "CAD" And numCount > 2 Then
    sameNum = CInt(numMatches.item(numMatches.Count - hasDateNum - 1))
    Set getPaperNumber = RangeCollection(sameNum, sameNum)
  ' CADデータにおいて，数字が複数見えた場合はその範囲すべてに対するファイルであると考える
  ElseIf fileType = "CAD" And numCount > 1 Then
    Dim lastIdx As Integer
    lastIdx = numMatches.Count - hasDateNum - 1
    Set getPaperNumber = RangeCollection(CInt(numMatches.item(lastIdx - 1)), CInt(numMatches.item(lastIdx)))
  ' このほかのパターンは存在しないはずだが，念のため記述
  Else
    sameNum = CInt(numMatches.item(numMatches.Count - hasDateNum - 1))
    Set getPaperNumber = RangeCollection(sameNum, sameNum)
  End If
End Function

Private Function getBluePrintName(ByVal fileName As String) As String
  Dim reg As RegExp
  Set reg = New RegExp
  Dim matches As MatchCollection
  
  ' 文字区切り
  With reg
    .Pattern = "[＿　][^＿　]+"
    .Global = True
  End With
  Set matches = reg.Execute(StrConv(fileName, vbWide))
  
  ' 図名候補を抽出
  Dim tmpName As String
  If matches.Count > 0 Then
    Dim i As Integer
    For i = 0 To matches.Count - 1
      If Not IsNumeric(matches(i)) Then
        tmpName = tmpName & matches(i)
      End If
    Next i
  Else
    tmpName = StrConv(fileName, vbWide)
  End If
  
  ' 数値を除外
  Dim numReg As RegExp
  Set numReg = New RegExp
  With numReg
    .Pattern = ".+(?=" & inner_RemoveNumbersPattern & ")"
    .Global = True
  End With
  Set matches = numReg.Execute(tmpName)
  
  ' 図名を確定
  If matches.Count > 0 Then
    getBluePrintName = matches.item(0)
  Else
    getBluePrintName = tmpName
  End If
End Function
