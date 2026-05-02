VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepoSheet_Body"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' 指定したシートの存在確認
Public Function IsExistsSheet(wb As Workbook, sheetName As String) As Boolean
  Dim sht As Worksheet
  For Each sht In wb.Worksheets
    If sht.Name = sheetName Then
      IsExistsSheet = True
      Exit Function
    End If
  Next sht
  IsExistsSheet = False
End Function


' シートの新規作成
Public Function Create(wb As Workbook, sheetName As String, Optional visible As Boolean = True) As Worksheet
  If IsExistsSheet(wb, sheetName) Then Exit Function
  
  Dim ws As Worksheet
  Set ws = wb.Sheets.Add(After:=Sheets(Sheets.Count))
　ws.Name = sheetName
  ws.Visible = visible

  Set Create = ws
End Function