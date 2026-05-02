VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ObjRegex: 正規表現を扱うユーティリティオブジェクト

' 指定パターンが入力文字列にマッチするかどうかを判定
' 引数: pattern (String) - 正規表現パターン
'       inputStr (String) - 判定対象の文字列
' 戻り値: Boolean - マッチする場合は True、しない場合は False
Public Function IsMatch(pattern As String, inputStr As String) As Boolean
  Dim reg As New RegExp
  
  With reg
    .Pattern = pattern
    .Global = True
  End With
  
  IsMatch = reg.Test(inputStr)
End Function

' 指定パターンにマッチする部分を置換した文字列を返す
' 引数: pattern (String) - 正規表現パターン
'       inputStr (String) - 置換対象の文字列
'       replaceStr (String) - 置換後の文字列
' 戻り値: String - 置換後の文字列
Public Function Replace(pattern As String, inputStr As String, replaceStr As String) As String
  Dim reg As New RegExp
  
  With reg
    .Pattern = pattern
    .Global = True
  End With
  
  Replace = reg.Replace(inputStr, replaceStr)
End Function

' 指定パターンにマッチする部分をすべて取得してコレクションで返す
' 引数: pattern (String) - 正規表現パターン
'       inputStr (String) - 検索対象の文字列
' 戻り値: ObjList - Matchオブジェクトのリスト
Public Function GetMatches(pattern As String, inputStr As String) As ObjList
  Dim reg As New RegExp
  Dim matches As Object
  Dim match As Object
  Dim matchedStrings As ObjList: Set matchedStrings = New ObjList
  
  With reg
    .Pattern = pattern
    .Global = True
  End With
  
  Set matches = reg.Execute(inputStr)
  
  For Each match In matches
    Call matchedStrings.Append(match)
  Next match
  
  Set GetMatches = matchedStrings
End Function

' Matchオブジェクトが示す位置の文字列を入力文字列から除外した文字列を返す
' 引数: match (Match) - Matchオブジェクト
'       inputStr (String) - 元の文字列
' 戻り値: String - 除外後の文字列
Public Function DropMatchedPosition(match As Match, inputStr As String) As String
  Dim result As String
  Dim start As Long, endPos As Long
  
  start = match.FirstIndex
  endPos = start + match.Length
  
  result = Left(inputStr, start) & Mid(inputStr, endPos + 1)
  
  DropMatchedPosition = result
End Function

' Matchオブジェクトが示す位置の文字列を入力文字列から抽出して返す
' 引数: match (Match) - Matchオブジェクト
'       inputStr (String) - 元の文字列
' 戻り値: String - 抽出した部分文字列
Public Function FocusMatchedPosition(match As Match, inputStr As String) As String
  FocusMatchedPosition = Mid(inputStr, match.FirstIndex + 1, match.Length)
End Function