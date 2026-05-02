VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepoSheet_Object"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit


' チェックボックスの値など，シート上に配置したオブジェクトの値を読み取る
Public Function ReadObjValue(wb As Workbook, sheetName As String, objName As String) As Variant
  ReadObjValue = wb.Worksheets(sheetName).OLEObjects(objName).Object.Value
End Function