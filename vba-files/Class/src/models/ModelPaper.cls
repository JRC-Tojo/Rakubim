VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelPaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private inner_PaperIndex As String
Private inner_SourceName As String
Private inner_AnalyzedName As String
Private inner_SnippedName As String
Private inner_PaperNumbers As ObjSet

' InitModel: ModelPaper にデータを登録する
' 引数: paperIndex (String), paperNumber (String), sourceName (String), Optional analyzeName (String), Optional snippedName (String)
' 戻り値: なし（内部にデータを保持）
Public Function InitModel(paperIndex As String, sourceName As String, analyzeName As String, snippedName As String, paperNumbers As ObjSet)
  inner_PaperIndex = paperIndex
  inner_SourceName = sourceName
  inner_AnalyzedName = analyzeName
  inner_SnippedName = snippedName
  Set inner_PaperNumbers = paperNumbers
End Function

' 図面番号を返す
Public Property Get PaperIndex() As Long
  On Error GoTo ConvertError
  PaperIndex = CLng(inner_PaperIndex)
  Exit Property
ConvertError:
  PaperIndex = -1
End Property

' 図面名称に含まれる番号（その○○）を返す
Public Property Get PaperNumber() As Long
  If inner_PaperNumbers.Count = 1 Then
    PaperNumber = inner_PaperNumbers.ToList().At(0)
  Else
    PaperNumber = -1
  End If
End Property

' 図面名称そのものを返す
Public Property Get SourceName() As String
  SourceName = inner_SourceName
End Property

' 図面名称からマッチングに使用できるように整形した名称を返す
Public Property Get AnalyzedName() As String
  AnalyzedName = inner_AnalyzedName
End Property
Public Property Let AnalyzedName(value As String)
  inner_AnalyzedName = value
End Property

' 図面名称から名称本体の部分（番号などの付属物を除く）を返す
Public Property Get SnippedName() As String
  SnippedName = inner_SnippedName
End Property