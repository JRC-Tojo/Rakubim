Attribute VB_Name = "OnTimeHelper"

Option Explicit

' OnTimeで指定するプロシージャを定義する

Public Sub ResetCell(sheetName As String, cellAddress As String)
  Worksheets(sheetName).Range(cellAddress).Value = ""
End Sub