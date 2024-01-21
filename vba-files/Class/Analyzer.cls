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

Public Sub SetPaperInfo(paper as PaperObj)
  
End Sub

' 名称を分析してクラスを作成
Private Sub Analyze(ByVal path As String)
  ' この順序で実行すること
  Call GetFileType(path)
  Call GetPaperNumber(path)
  Call GetBluePrintName(path)
  Call SetKeyPathRelation(path)
End Sub

' ファイルの種類を特定
Private Sub GetFileType(ByVal fileName As String)
  Select Case StrConv(Right(fileName, 3), vbLowerCase)
    Case "pdf"
      FileType = "PDF"
    Case "dwg"
      FileType = "CAD"
    Case Else
      FileType = "OTHER"
  End Select
End Sub

' ファイル名から図面名称を抽出
Private Sub GetBluePrintName_old(ByVal fileName As String)
  Dim reg As RegExp
  Set reg = New RegExp
  Dim matches As MatchCollection
  
  With reg
    .Pattern = "[^\x01-\x7E].+（その"
    .Global = True
  End With
  Set matches = reg.Execute(fileName)
  
  If matches.Count <> 0 Then
    paperName = Left(matches.item(0), Len(matches.item(0)) - 3)
  Else
    With reg
      .Pattern = "[^\x01-\x7E]+"
      .Global = True
    End With
    
    ' 日本語が複数ある場合は末尾が「図」となっているものを返す
    ' （すべてに「図」がない場合は0番目を返す）
    Dim returnCount As Integer
    Set matches = reg.Execute(fileName)
    If matches.Count = 1 Then
      paperName = matches.item(0)
    ElseIf matches.Count > 0 Then
      returnCount = 0
      Dim i As Integer
      For i = 0 To matches.Count - 1
        If InStr(matches.item(i), "図") <> 0 Then
          returnCount = i
        End If
      Next i
      paperName = matches.item(returnCount)
    Else
      paperName = fileName
    End If
  End If

End Sub

' ファイル名から図面名称を抽出
' １．＿や全角スペースで区切る
' ２．区切った文字列のうち「図」を含むものを抽出
' ３．いずれにも「図」がない場合は、一番最後の区切りを採用
' ４．上記で抽出した文字列からGetPaperNumberに準じて数値の前の文字列を抽出
' ５．抽出結果が空の時は、最後の拡張子を削るのみしたうえで図名を確定
Private Sub GetBluePrintName_back(ByVal fileName As String)
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
    Dim returnIdx As Integer
    returnIdx = matches.Count - 1
    For i = 0 To matches.Count - 1
      If InStr(matches.item(i), "図") <> 0 Then
        returnIdx = i
      End If
    Next i
    ' 区切り文字が先頭に入っているため除外
    tmpName = Mid(matches.item(returnIdx), 2)
  Else
    tmpName = StrConv(fileName, vbWide)
  End If
  
  ' 数値を除外
  Dim numReg As RegExp
  Set numReg = New RegExp
  With numReg
    .Pattern = ".+(?=" & NumberPattern & ")"
    .Global = True
  End With
  Set matches = numReg.Execute(tmpName)
  
  ' 図名を確定
  If matches.Count > 0 Then
    paperName = matches.item(0)
  Else
    paperName = Left(tmpName, Len(tmpName) - 3)
  End If
End Sub
Private Sub GetBluePrintName(ByVal fileName As String)
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
  
  ' 拡張子を除外
  If FileType <> "OTHER" Then
    tmpName = Left(tmpName, Len(tmpName) - 4)
  End If
  
  ' 数値を除外
  Dim numReg As RegExp
  Set numReg = New RegExp
  With numReg
    .Pattern = ".+(?=" & NumberPattern & ")"
    .Global = True
  End With
  Set matches = numReg.Execute(tmpName)
  
  ' 図名を確定
  If matches.Count > 0 Then
    paperName = matches.item(0)
  Else
    paperName = tmpName
  End If
End Sub

' ファイル名から図面の連番（その○○）を取得
' （図面番号、終了番号(only DWG)）を戻り値として返す
' 「図面番号」は連番や（その〇），（その〇～〇）の最初の数字，が入る
' 全角数字一つしか検知されない場合：その数字が図面の番号
' 全角数字が複数検知された場合：その数字は図面そのものに対して意味を持たない
' 半角数字が検知された場合：末尾の数字は連続図面の連番を示す（連番と「その〇」の対応関係は考慮しない）
Private Sub GetPaperNumber(ByVal fileName As String)
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
    .Pattern = IIf( _
                  systemIds.UseContinuousNumber And FileType = "PDF", _
                  continuousNumbersPattern, _
                  NumberPattern)
    .Global = True
  End With
  ' ファイル名称の日本語以降を抜き出して情報抽出
  targetName = StrConv(fileName, vbWide)
  
  ' 拡張子を除外
  If FileType <> "OTHER" Then
    targetName = Left(targetName, Len(targetName) - 4)
  End If
  
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
    Set PaperNumbers = RangeCollection(-1, -1)
  ' 数字が一つ見えた場合はその数字が図面そのものの番号である
  ElseIf numCount = 1 Then
    sameNum = CInt(numMatches.item(0))
    Set PaperNumbers = RangeCollection(sameNum, sameNum)
  ' PDFデータにおいて，数字が2つ見える場合は末尾の「1」が省略されたと考える
  ElseIf FileType <> "CAD" And numCount = 2 Then
    Set PaperNumbers = RangeCollection(-1, -1)
  ' PDFデータにおいて，数字が複数見える場合は末尾の数字が図面そのものの番号である
  ElseIf FileType <> "CAD" And numCount > 2 Then
    sameNum = CInt(numMatches.item(numMatches.Count - hasDateNum - 1))
    Set PaperNumbers = RangeCollection(sameNum, sameNum)
  ' CADデータにおいて，数字が複数見えた場合はその範囲すべてに対するファイルであると考える
  ElseIf FileType = "CAD" And numCount > 1 Then
    Dim lastIdx As Integer
    lastIdx = numMatches.Count - hasDateNum - 1
    Set PaperNumbers = RangeCollection(CInt(numMatches.item(lastIdx - 1)), CInt(numMatches.item(lastIdx)))
  ' このほかのパターンは存在しないはずだが，念のため記述
  Else
    sameNum = CInt(numMatches.item(numMatches.Count - hasDateNum - 1))
    Set PaperNumbers = RangeCollection(sameNum, sameNum)
  End If
End Sub

' 「何番目」の情報をキーとして、それぞれのファイルパスを対応させたDictを作成
' このオブジェクトの時点で「どの図面か」は特定されているため、「図面内の何番目はどのパスか」が分かればよい
Private Sub SetKeyPathRelation(ByVal fileName As String)
  ' PaperNumbersの数だけpathも与えておく
  Dim i As Integer
  For i = 1 To PaperNumbers.Count
    PaperPaths.Add "idx" & PaperNumbers(i), fileName
  Next i
End Sub

  