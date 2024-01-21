VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' ファイルの種類（PDF｜CAD）
Private inner_FileType As String
' 拡張子
Private inner_Extension As String
' ファイルの元の名前
Private inner_SourceName As String
' マッチング用の処理済みのファイル名
Private inner_SnippedName As String
' 同一ファイルに格納している（その〇）に該当する番号
Private inner_PaperNumbers As Collection
' PaperNumbersを使って図面データと突き合わせる用の辞書
Private inner_PaperPaths As Scripting.Dictionary
' このファイルがサポートしている図面のID一覧
Private inner_MatchedPaperIds As Collection

Private Sub class_initialize()
  ' 初期化
  Set inner_PaperNumbers = New Collection
  Set inner_PaperPaths = New Scripting.Dictionary
  Set inner_MatchedPaperIds = New Collection
End Sub

Public Property Get FileType() As String
  FileType = inner_FileType
End Property

Public Property Get Extension() As String
  Extension = inner_Extension
End Property

Public Property Get SourceName() As String
  SourceName = inner_SourceName
End Property

Public Property Let SnippedName(name As String)
  inner_SnippedName = name
End Property
Public Property Get SnippedName() As String
  SnippedName = inner_SnippedName
End Property

Public Property Get MatchedPaperIds() As Collection
  Set MatchedPaperIds = inner_MatchedPaperIds
End Property

' {"idx1": path}という形式の辞書を返す
' idxは図面が持つ番号を示しており，これを突き合わせることで（その〇）のような番号が整合してマッチングさせる
' 番号を持たない場合は"idx-1"がキーに登録される
Public Property Get PaperPaths() As Scripting.Dictionary
  Set PaperPaths = inner_PaperPaths
End Property
Public Property Set PaperPaths(paths As Scripting.Dictionary)
  Set inner_PaperPaths = paths
End Property

Public Property Get PaperNumbers() As Collection
  Set PaperNumbers = inner_PaperNumbers
End Property

Public Sub InitFileObj(ByVal fileName As String)
  ' 基本情報を登録
  inner_SourceName = fileName
  Call GetFileType()

  ' ファイル名が無効の時は処理を行わない
  If (fileName = "") Then
    Exit Sub
  End If
  
  ' 名称の分析結果を登録
  Dim analyzerObj As Analyzer
  Set analyzerObj = New Analyzer
  Call analyzerObj.SetFileInfo(Me)

  inner_SnippedName = analyzerObj.SnippedName
  Set inner_PaperNumbers = analyzerObj.PaperNumbers
  ' 登録された番号の情報をもとに図面との突き合わせ用データを作成
  Call SetKeyPathRelation()
End Sub

' ファイルの種類を特定
Private Sub GetFileType()
  inner_Extension = Right(SourceName, 3)
  Select Case StrConv(inner_Extension, vbLowerCase)
    Case "pdf"
      inner_FileType = "PDF"
    Case "dwg"
      inner_FileType = "CAD"
    Case Else
      inner_FileType = "OTHER"
  End Select
End Sub

' 「何番目」の情報をキーとして、それぞれのファイルパスを対応させたDictを作成
' このオブジェクトの時点で「どの図面か」は特定されているため、「図面内の何番目はどのパスか」が分かればよい
Private Sub SetKeyPathRelation()
  ' PaperNumbersの数だけpathも与えておく
  Set inner_PaperPaths = New Scripting.Dictionary

  Dim i As Integer
  For i = 1 To inner_PaperNumbers.Count
    inner_PaperPaths.Add "idx" & inner_PaperNumbers(i), SourceName
  Next i
End Sub