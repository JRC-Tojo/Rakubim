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

' 目録入力シートに入力された図面のObject

Public paperID As String
Public SourceName As String
Public paperName As String
Public PaperNumber As Integer
Public pdfPath As String
Public dwgPath As String

Private Sub class_initialize()

End Sub

' 指定した行数にあるデータから図面Objectを作成する
Public Sub GetPaper(ByVal rowNum As Integer)
  Dim paperData As FileObj
  Dim sourcePaperName As String
  
  With Worksheets("目録入力")
    sourcePaperName = .Cells(rowNum, "B").Value
    Set paperData = AnalyzeName(sourcePaperName)
    paperID = CompletePaperID(rowNum)
    SourceName = .Cells(rowNum, "B").Value
    paperName = Replace(StrConv(sourcePaperName, vbWide), "　", "")
    PaperNumber = paperData.PaperNumbers(1)
    pdfPath = .Cells(rowNum, "G").Value
    dwgPath = .Cells(rowNum, "H").Value
  End With
End Sub

' 図面番号を補完する
' 補完した番号を戻り値として返す
Private Function CompletePaperID(rowNum As Integer)
  Dim nowId As Variant

  With Worksheets("目録入力")
    nowId = .Cells(rowNum, "A").Value
    If nowId = 0 Then
      nowId = .Cells(rowNum - 1, "A").Value + 1
    End If
  End With
  
  CompletePaperID = CStr(nowId)
End Function
