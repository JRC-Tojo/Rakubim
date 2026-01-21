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

Private inner_TargetPath As String
Private inner_CompanyID As String
Private inner_ProjectID As String
Private inner_StructureID As String
Private inner_DesignID As String
Private inner_SelectedFileType As String
Private inner_UseContinuousNumber As Boolean

Private Sub class_initialize()
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.FirstInputPage)
    inner_TargetPath = .Range("E3")
    inner_CompanyID = getCompanyID(.Range("C6"))
    inner_ProjectID = .Range("C7")
    inner_StructureID = .Range("C8")
    inner_DesignID = .Range("C9")
    inner_UseContinuousNumber = CBool(.useContinuousNumberCheck.Value)
    ' 下記は「処理するデータの種類」という項目が入力シートにあった際に利用していた変数
    ' 現在は使っていないため，すべての種類を処理対象とする「ALL」を指定している
    inner_SelectedFileType = "ALL"
  End With
End Sub

Public Property Get TargetPath() As String
  TargetPath = inner_TargetPath
End Property

Public Property Get CompanyID() As String
  CompanyID = inner_CompanyID
End Property

Public Property Get ProjectID() As String
  ProjectID = inner_ProjectID
End Property

Public Property Get StructureID() As String
  StructureID = inner_StructureID
End Property

Public Property Get DesignID() As String
  DesignID = inner_DesignID
End Property

Public Property Get SelectedFileType() As String
  SelectedFileType = inner_SelectedFileType
End Property

Public Property Get UseContinuousNumber() As Boolean
  UseContinuousNumber = inner_UseContinuousNumber
End Property



' 業務機関コードから情報を取得する
Private Function getCompanyID(name As String)
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.TeamCodePage)
    getCompanyID = WorksheetFunction.VLookup(name, .Range("H3:I21"), 2, False)
  End With
End Function
