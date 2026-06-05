VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepoSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private inner_Workbook As Workbook

Private inner_Body As RepoSheet_Body
Private inner_File As RepoSheet_File
Private inner_IO As RepoSheet_IO
Private inner_Object As RepoSheet_Object
Private inner_Lattice As RepoSheet_Lattice
Private inner_Validation As RepoSheet_Validation

Private Sub class_initialize()
  Set inner_Workbook = Application.ThisWorkbook
  Set inner_Body = New RepoSheet_Body
  Set inner_File = New RepoSheet_File
  Set inner_IO = New RepoSheet_IO
  Set inner_Object = New RepoSheet_Object
  Set inner_Lattice = New RepoSheet_Lattice
  Set inner_Validation = New RepoSheet_Validation
End Sub

Public Function SetWorkbook(wb As Workbook) As RepoSheet
  Set inner_Workbook = wb
End Function

' ========================= Body =========================

Public Function CreateSheet(sheetName As String, Optional visible As Boolean = True) As Worksheet
  Set CreateSheet = inner_Body.Create(inner_Workbook, sheetName, visible)
End Function

Public Function IsExistsSheet(sheetName As String) As ObjResult
  Set IsExistsSheet = inner_Body.IsExistsSheet(inner_Workbook, sheetName)
End Function

Public Function CopySheetFromOuter(outerPath As ObjPath, sheetName As String, Optional visible As Boolean = True) As ObjResult
  Set CopySheetFromOuter = inner_IO.CopySheetFromOuter(inner_Workbook, outerPath, sheetName, visible)
End Function

' ========================= IO =========================

Public Function ReadColumn(ByVal sheetName As String, ByVal startCell As String) As ObjResult
  Set ReadColumn = inner_IO.ReadColumn(inner_Workbook, sheetName, startCell)
End Function

Public Function ReadCell(sheetName As String, cellAddress As String, prohibitEmpty As Boolean) As ObjResult
  Set ReadCell = inner_IO.ReadCell(inner_Workbook, sheetName, cellAddress, prohibitEmpty)
End Function

Public Function WriteColFromCell(sheetName As String, startCell As String, values As ObjList) As ObjResult
  Set WriteColFromCell = inner_IO.WriteColFromCell(inner_Workbook, sheetName, startCell, values)
End Function

Public Function GetLastRow(sheetName As String, startCell As String, Optional mustContinue As Boolean = False) As ObjResult
  Set GetLastRow = inner_IO.GetLastRow(inner_Workbook, sheetName, startCell, mustContinue)
End Function

Public Function ClearColFromCell(sheetName As String, ByVal startCell As String) As ObjResult
  Set ClearColFromCell = inner_IO.ClearColFromCell(inner_Workbook, sheetName, startCell)
End Function

Public Function RegistCellMsg(sheetName As String, cellAddress As String, msg As String) As ObjResult
  Set RegistCellMsg = inner_IO.RegistCellMsg(inner_Workbook, sheetName, cellAddress, msg)
End Function

Public Function RegistPrintArea(sheetName As String, Optional startCell As String = "", Optional endCell As String = "") As ObjResult
  Set RegistPrintArea = inner_IO.RegistPrintArea(inner_Workbook, sheetName, startCell, endCell)
End Function

Public Function ClearPrintArea(sheetName As String) As ObjResult
  Set ClearPrintArea = inner_IO.ClearPrintArea(inner_Workbook, sheetName)
End Function

' ========================= IO =========================

Public Function ReadObjValue(sheetName As String, objName As String) As Variant
  ReadObjValue = inner_Object.ReadObjValue(inner_Workbook, sheetName, objName)
End Function

' ========================= Lattice =========================

Public Function MakeGrid(workSheetName As String, Optional startCell As String = "", Optional endCell As String = "")
  Call inner_Lattice.MakeGrid(inner_Workbook, workSheetName, startCell, endCell)
End Function

Public Function MakeBoldOutline(workSheetName As String, Optional startCell As String = "", Optional endCell As String = "")
  Call inner_Lattice.MakeBoldOutline(inner_Workbook, workSheetName, startCell, endCell)
End Function

Public Function ClearGrid(workSheetName As String, Optional startCell As String = "", Optional endCell As String = "")
  Call inner_Lattice.ClearGrid(inner_Workbook, workSheetName, startCell, endCell)
End Function

' ========================= Validation =========================

Public Function Pulldown(sheetName As String, cellAddress As String, elems As ObjList, Optional ErrTitle As String = "", Optional ErrMessage As String = "")
  Call inner_Validation.Pulldown(inner_Workbook, sheetName, cellAddress, elems, ErrTitle, ErrMessage)
End Function

' ========================= File =========================

Public Function ExportToExcel(sheetName As String, targetPath As ObjPath, rangeAddress As String, Optional newSheetName As String = "") As ObjResult
  Set ExportToExcel = inner_File.ExportToExcel(inner_Workbook, sheetName, targetPath, rangeAddress, newSheetName)
End Function