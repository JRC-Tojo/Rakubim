VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileNameCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 全てのファイル名から共通語句を抽出し、それらを除去したファイル名一覧を返す

Option Explicit

' 文字列分離に使うパターン
Private Const SplitPattern = "[^　＿]+(?=[　＿])"
' 抽出した文字列とその個数
Private WordBundle As Scripting.Dictionary
Private RemoveWords As Collection

Private Sub class_initialize()
  ' 初期化
  Set WordBundle = New Scripting.Dictionary
  Set RemoveWords = New Collection

  ' 極めて汎用的な語彙は除外対象とする
  Call RemoveWords.Add("一般図")
  Call RemoveWords.Add("詳細図")
End Sub

' 名称一覧から除去すべき単語を学習する
Public Function AnalyzeNames(names As Collection)
  Dim reg As RegExp
  Dim matches As MatchCollection
  Set reg = New RegExp
  With reg
    .Pattern = SplitPattern
    .Global = True
  End With
  
  ' 名称をローラー
  Dim i As Integer
  For i = 1 To names.Count
    Call RegistWord(reg.Execute(StrConv(names.item(i), vbWide)))
  Next i
  
  ' 除去するワードを特定（現状は１つのみ抽出するが、必要であれば複数抽出もOK）
  ' 除去ワードが全体の80%以上に含まれていることを条件とする
  Dim maxInfo As Variant
  For i = 1 To dictValues(WordBundle).Count
    If dictValues(WordBundle).item(i) > names.Count * 0.8 Then
      RemoveWords.Add dictKeys(WordBundle).item(i)
    End If
  Next i
End Function

' 単発の名称から除去すべき語句を除去
Public Function NameCleaner(name As String)
  Dim returnName As String
  returnName = StrConv(name, vbWide)
  
  Dim i As Integer
  For i = 1 To RemoveWords.Count
    returnName = Replace(returnName, RemoveWords.item(i), "")
  Next i
  
  NameCleaner = returnName
End Function

' 取得したパターンを登録する
Private Sub RegistWord(matches As MatchCollection)
  If matches.Count = 0 Then
    Exit Sub
  End If
  
  Dim i As Integer
  Dim tmpWord As String
  For i = 0 To matches.Count - 1
    tmpWord = CStr(matches.item(i))
    If WordBundle.Exists(tmpWord) Then
      WordBundle(tmpWord) = WordBundle(tmpWord) + 1
    Else
      WordBundle.Add tmpWord, 1
    End If
  Next i
End Sub
