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
Private inner_SnippedName As String
' 図面の連番
Private inner_ContinuousNumber As Integer
' 図面番号
Private inner_GlobalNumber As String
' （その〇）にあたる番号
Private inner_SpotNumber As Integer

Private Sub class_initialize()

End Sub

Public Property Get SourceName() As String
  Set SourceName = inner_SourceName
End Property

Public Property Get SnippedName() As String
  Set SnippedName = inner_SnippedName
End Property
Public Property Let SnippedName(name As String)
  inner_SnippedName = name
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
Public Sub InitPaperObj(ByVal rowNum As Integer, continuousNumber As Integer)
  Dim analyzerObj as Analyzer
  analyzerObj = New Analyzer
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  inner_ContinuousNumber = continuousNumber
  
  With Worksheets(sheets.MokurokuPage)
    inner_SourceName = .Cells(rowNum, "B").Value
    analyzerObj.SetPaperInfo(Me)

    inner_GlobalNumber = CompleteGlobalPaperNumber(rowNum)
    inner_SnippedName = Replace(StrConv(inner_SourceName, vbWide), "　", "")
    PaperNumber = analyzerObj.PaperNumbers(1)
    pdfPath = .Cells(rowNum, "G").Value
    dwgPath = .Cells(rowNum, "H").Value
  End With
End Sub

' 図面番号を補完する
' 補完した番号を戻り値として返す
Private Function CompleteGlobalPaperNumber(rowNum As Integer)
  Dim nowId As Variant
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.MokurokuPage)
    nowId = .Cells(rowNum, "A").Value
    If nowId = 0 Then
      nowId = .Cells(rowNum - 1, "A").Value + 1
    End If
  End With
  
  CompleteGlobalPaperNumber = CStr(nowId)
End Function
