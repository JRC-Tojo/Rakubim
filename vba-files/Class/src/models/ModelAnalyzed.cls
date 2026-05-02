VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelAnalyzed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' マッチング処理での使用を想定した，正規化された解析結果
Private inner_AnalyzedName As String
' リネーム処理での使用を想定した，SourceNameの名称部分の解析結果
Private inner_SnippedName As String
' 図面の連番情報（その○○）を格納
Private inner_PaperNumbers As ObjSet

' 解析結果を返すためのモデル
Public Function InitModel(analyzedName As String, snippedName As String, paperNumbers As ObjSet)
  inner_AnalyzedName = analyzedName
  inner_SnippedName = snippedName
  Set inner_PaperNumbers = paperNumbers
End Function

Public Property Get AnalyzedName() As String
  AnalyzedName = inner_AnalyzedName
End Property

Public Property Get SnippedName() As String
  SnippedName = inner_SnippedName
End Property

Public Property Get PaperNumbers() As ObjSet
  Set PaperNumbers = inner_PaperNumbers
End Property