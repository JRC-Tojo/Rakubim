Attribute VB_Name = "OnTimeHelper"

Option Explicit

' OnTimeで指定するプロシージャを定義する

' 指定したセルの値をリセットする
Public Sub ResetCell(sheetName As String, cellAddress As String)
  Worksheets(sheetName).Range(cellAddress).Value = ""
End Sub
