VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SystemIdObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public CompanyID As String
Public ProjectID As String
Public StructureID As String
Public DesignID As String
Public SelectedFileType As String
Public UseContinuousNumber As Boolean

Private Sub class_initialize()
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.FirstInputPage)
    CompanyID = getCompanyID(.Range("C3"))
    ProjectID = .Range("C4")
    StructureID = .Range("C5")
    DesignID = .Range("C6")
    SelectedFileType = SetFileType()
    UseContinuousNumber = CBool(.useContinuousNumberCheck.Value)
  End With
End Sub

' 業務機関コードから情報を取得する
Private Function getCompanyID(name As String)
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.TeamCodePage)
    getCompanyID = WorksheetFunction.VLookup(name, .Range("H3:I21"), 2, False)
  End With
End Function

Private Function SetFileType() As String
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.FirstInputPage)
    If (.AllFilesBtn.Value = True) Then
      SetFileType = "ALL"
    ElseIf (.OnlyPDFBtn.Value = True) Then
      SetFileType = "PDF"
    ElseIf (.OnlyCADBtn.Value = True) Then
      SetFileType = "CAD"
    Else
      SetFileType = "undefined"
    End If
  End With
End Function