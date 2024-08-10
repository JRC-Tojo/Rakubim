VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NewNameCreater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private inner_RenamePDFFiles As Scripting.Dictionary
Private inner_RenameCADFiles As Scripting.Dictionary
Private inner_Papers As Scripting.Dictionary
Private inner_FixedHeadName As String

Private Sub class_initialize()
  Set inner_RenamePDFFiles = New Scripting.Dictionary
  Set inner_RenameCADFiles = New Scripting.Dictionary
  Set inner_Papers = New Scripting.Dictionary
End Sub

' Createrが初期化で必要とする情報を登録
Public Function InitCreater(companyId As String, projectId As String, structureId As String)
  inner_FixedHeadName = companyId & "-" & projectId & "-" & structureId
End Function

Public Property Get Papers() As Scripting.Dictionary
  Set Papers = inner_Papers
End Property

Public Property Get NamesCount() As Integer
  ' TODO: 個数が正しいか要確認
  NamesCount = UBound(inner_RenamePDFFiles.Keys) + UBound(inner_RenameCADFiles.Keys) + 2
End Property
Public Property Get OldNames() As Collection
  Set OldNames = CombineCols(dictKeys(inner_RenamePDFFiles), dictKeys(inner_RenameCADFiles))
End Property
Public Property Get NewNames() As Collection
  Dim returnCol As Collection
  Set returnCol = New Collection

  Dim file As FileObj
  For Each file In dictValues(inner_RenamePDFFiles)
    Call returnCol.Add(GetNewFileName(file))
  Next file
  For Each file In dictValues(inner_RenameCADFiles)
    Call returnCol.Add(GetNewFileName(file))
  Next file

  Set NewNames = returnCol
End Property


' 図面とファイルの対応関係を記録
Public Sub SetPDFFile(file As FileObj, paperID As String)
  Call RegistPaperID(inner_RenamePDFFiles, file, paperID)
End Sub
Public Sub SetCADFile(file As FileObj, paperID As String)
  Call RegistPaperID(inner_RenameCADFiles, file, paperID)
End Sub

' ファイルに対応する図面IDを登録
Private Sub RegistPaperID( _
  files As Scripting.Dictionary, _
  file As FileObj, _
  paperID As String _
)
  ' 無効なファイル名は除外する
  If (file.SourceName = "") Then
    Exit Sub
  End If

  ' リネーム対象のファイル一覧にファイルがない場合は追加
  If (Not files.Exists(file.SourceName)) Then
    Call files.Add(file.SourceName, file)
  End If

  ' リネーム対象のファイルと対応する図面IDを登録
  Call files.Item(file.SourceName).MatchedPaperIds.Add(paperID)
End Sub


' 指定されたFileに対する新名称を返す
Private Function GetNewFileName(file As FileObj) As String
  Dim title As String
  title = SetPaperNameText(file)
  GetNewFileName = inner_FixedHeadName & "-" & _
                    SetPaperNumberText(file) & "_" & _
                    title & _
                    SetPaperSpotNumbers(file, title) & _
                    "." & file.Extension
End Function

' 新名称に含める番号の表現を作成
Private Function SetPaperNumberText(file As FileObj)
  Dim returnText As String: returnText = ""
  Dim number As String
  Dim FormattedNumber As String

  ' 「0000, 」を削除して連番「-0000」を作成するため，連番処理を実行するための最低文字数
  Const MIN_TEXT_LEN As Integer = 6

  ' 連番のひとつ前の値を保持する
  ' TODO: 01-03,05-08のようなとび番があるケースに対応
  Dim beforeNumber As Integer: beforeNumber = -1
  Dim i As Integer
  For i = 1 To file.MatchedPaperIds.Count
    number = inner_Papers.Item(file.MatchedPaperIds(i)).GlobalNumber

    ' 数値の時に連続化検討処理を行う
    FormattedNumber = IIf(Len(number) > 4, number, Right("0000" & number, 4))
    If (IsNumeric(number)) Then
      If (CInt(number) - beforeNumber = 1 And Len(returnText) > MIN_TEXT_LEN) Then
        ' すでについている「0000, 」を削除して「-」を挿入することで「01-03」の表記を作成
        returnText = Left(returnText, Len(returnText) - MIN_TEXT_LEN) & "-"
      End If
      beforeNumber = CInt(number)
    End If

    returnText = returnText & FormattedNumber & ","
  Next i

  ' 最後の「, 」を除外して返す
  SetPaperNumberText = Left(returnText, Len(returnText) - 1)
End Function

' 新名称に含める図面名の表現を作成
Private Function SetPaperNameText(file As FileObj)
  Dim paperName As String
  Dim texts As Collection
  Set texts = New Collection

  ' 図面名の重複を除外しながらユニークな図面名のリストを作成
  Dim i As Integer
  For i = 1 To file.MatchedPaperIds.Count
    paperName = inner_Papers.Item(file.MatchedPaperIds(i)).SnippedName
    paperName = TitleFormatter(paperName)
    If (Not HasItem(texts, paperName)) Then
      Call texts.Add(paperName)
    End If
  Next i 

  ' カンマ区切りで図面名を結合
  Dim returnText As String
  Dim nowText As String
  Dim beforeText As String
  For i = 1 To texts.Count
    nowText = Trim(texts(i))
    If (beforeText <> nowText) Then
      returnText = returnText & nowText & ", "
    End If
    beforeText = nowText
  Next i

  ' 最後の「, 」を除外して返す
  SetPaperNameText = Left(returnText, Len(returnText) - 2)
End Function

' （その〇～その〇）を追加
Private Function SetPaperSpotNumbers(file As FileObj, title As String)
  Dim paperIds As Collection
  Set paperIds = file.MatchedPaperIds
  Dim firstNumber As Integer
  Dim lastNumber As Integer
  firstNumber = inner_Papers.Item(paperIds(1)).SpotNumber
  lastNumber = inner_Papers.Item(paperIds(paperIds.Count)).SpotNumber

  If (InStr(title, ", ") = 0 And firstNumber <> -1) Then
    If (firstNumber <> lastNumber) Then
      SetPaperSpotNumbers = "（その" & StrConv(CStr(firstNumber), vbWide) & "～その" & StrConv(CStr(lastNumber), vbWide) & "）"
    Else
      SetPaperSpotNumbers = "（その" & StrConv(CStr(firstNumber), vbWide) & "）"
    End If
  Else
    SetPaperSpotNumbers = ""
  End If
End Function

' 新規図面に入るタイトル文字のフォーマット
Private  Function TitleFormatter(title As String)
  Dim i As Integer
  For i = 1 To Len(title)
    If Mid(title, i, 1) Like "[ａ-ｚＡ-Ｚ０-９　]" Then
      Mid(title, i, 1) = StrConv(Mid(title, i, 1), vbNarrow)
    End If
  Next
  TitleFormatter = title
End Function