VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PaperObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' 入力された図面名称そのまま
Private inner_SourceName As String
' 処理済みでマッチングに使う名称
Private inner_AnalyzeName As String
' 元の名称から図面名のみを切り出した名称
Private inner_SnippedName As String
' 図面の連番
Private inner_ContinuousNumber As Integer
' 図面番号
Private inner_GlobalNumber As String
' （その〇）にあたる番号
Private inner_SpotNumber As Integer

Private Sub class_initialize()

End Sub

Public Property Get Id() As String
  Dim calcUtil As CalcUtils
  Set calcUtil = New CalcUtils
  Id = calcUtil.HASH_SHA256(inner_SourceName)
End Sub

Public Property Get SourceName() As String
  SourceName = inner_SourceName
End Property

Public Property Get AnalyzeName() As String
  AnalyzeName = inner_AnalyzeName
End Property

Public Property Get SnippedName() As String
  SnippedName = inner_SnippedName
End Property

Public Property Get ContinuousNumber() As Integer
  ContinuousNumber = inner_ContinuousNumber
End Property

' 図面番号を取得する
Public Property Get GlobalNumber() As String
  GlobalNumber = inner_GlobalNumber
End Property

' 同一図面種類内での数値＝（その〇）を取得する
' 単一図面の場合は-1を返す
Public Property Get SpotNumber() As Integer
  SpotNumber = inner_SpotNumber
End Property

' 指定した行数にあるデータから図面Objectを作成する
Public Sub InitPaperObj(paperName As String, continuousNumber As Integer, globalNumber As String)
  Dim analyzerObj As Analyzer
  Set analyzerObj = New Analyzer

  inner_SourceName = paperName
  inner_ContinuousNumber = continuousNumber  
  inner_GlobalNumber = globalNumber
  
  Call analyzerObj.SetPaperInfo(Me)
  inner_AnalyzeName = analyzerObj.AnalyzeName
  inner_SnippedName = analyzerObj.SnippedName
  inner_SpotNumber = analyzerObj.PaperNumbers(1)
End Sub
