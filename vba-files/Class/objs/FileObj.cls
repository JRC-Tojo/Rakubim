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

' クラス変数
Private inner_FileType As String
Private inner_SourceName As String
Private inner_SnippedName As String
Private inner_PaperNumbers As Collection
Private inner_PaperPaths As Scripting.Dictionary

Private Sub class_initialize()
  ' 初期化
  Set inner_PaperNumbers = New Collection
  Set inner_PaperPaths = New Scripting.Dictionary
End Sub

Public Property Get SourceName() As String
  Set SourceName = inner_SourceName
End Property

Public Property Let SnippedName(name As String)
  inner_SnippedName = name
End Property
Public Property Get SnippedName() As String
  SnippedName = inner_SnippedName
End Property

' {"idx1": path}という形式の辞書を返す
' idxは図面が持つ番号を示しており，これを突き合わせることで（その〇）のような番号が整合してマッチングさせる
' 番号を持たない場合は"idx-1"がキーに登録される
Public Property Get PaperPaths() As Scripting.Dictionary
  Set PaperPaths = inner_PaperPaths
End Property

Public Property Set PaperNumbers(numbers As Collection)
  inner_PaperNumbers = numbers
  Call SetKeyPathRelation()
End Property
Public Property Get PaperNumbers() As RakuCollection
  Set PaperNumbers = inner_PaperNumbers
End Property

Public Property Get Id() As String
  dim calcUtil As New CalcUtils
  Id = calcUtil.HASH_SHA256(SourceName)
End Sub

Public Sub InitFileObj(ByVal fileName As String)
  inner_SourceName = fileName
  Call GetFileType()
End Sub

' ファイルの種類を特定
Private Sub GetFileType()
  Select Case StrConv(Right(SourceName, 3), vbLowerCase)
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

