VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepoSheet_IO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit


' RepoSheet.ReadColumn: 指定シートの指定セルから縦方向に値を ObjList として返す
' 引数: sheetName (String), startCell (String)
' 戻り値: ObjResult - 成功時は ObjList(String) を返す
Public Function ReadColumn(wb As Workbook, sheetName As String, startCell As String, Optional ignoreEmptyCell As Boolean = True) As ObjResult
  On Error GoTo ErrHandler

  Dim list As ObjList
  Set list = New ObjList
  
  ' 配列一括読み込み（最初の10000行）
  Dim arrValues As Variant
  With wb.Worksheets(sheetName)
    arrValues = .Range(startCell).Resize(10000, 1).Value
  End With

  ' 配列をループ処理
  Dim i As Long
  For i = 1 To UBound(arrValues)
    If arrValues(i, 1) <> "" Then
      list.Append arrValues(i, 1)
    ElseIf Not ignoreEmptyCell Then
      list.Append ""
    End If
  Next i

  Set ReadColumn = Success(list)
  Exit Function
ErrHandler:
  Set ReadColumn = Failure("repo.sheet.readcol.error", sheetName, startCell, Err.Number, Err.Description)
End Function

' 特定のセルの値を取得する
Public Function ReadCell(wb As Workbook, sheetName As String, cellAddress As String, prohibitEmpty As Boolean) As ObjResult
  On Error GoTo ErrHandler

  Dim cellValue As String
  With wb.Worksheets(sheetName)
    cellValue = .Range(cellAddress).Value
  End With

  ' 空値禁止チェック
  If prohibitEmpty And IsEmpty(cellValue) Then
    Set ReadCell = Failure("repo.sheet.readcell.empty", sheetName, cellAddress)
    Exit Function
  End If

  Set ReadCell = Success(cellValue)
  Exit Function
ErrHandler:
  Set ReadCell = Failure("repo.sheet.readcell.error", sheetName, cellAddress, Err.Number, Err.Description)
End Function

' RepoSheet.WriteColFromCell: 指定シートの指定セルから縦方向に ObjList の値を書き込む
' 引数: sheetName (String), startCell (String), values (ObjList of String)
' 戻り値: ObjResult - 成功時は True を返す
Public Function WriteColFromCell(wb As Workbook, sheetName As String, startCell As String, values As ObjList) As ObjResult
  On Error GoTo ErrHandler
  ' Variants配列はRangeで一括書き込みが可能
  With wb.Worksheets(sheetName)
    .Range(startCell, .Range(startCell).Offset(values.Count - 1, 0)).Value = WorksheetFunction.Transpose(values.ToArray())
  End With

  Set WriteColFromCell = Success(True)
  Exit Function
ErrHandler:
  Set WriteColFromCell = Failure("repo.sheet.writecol.error", sheetName, startCell, Err.Number, Err.Description)
End Function

' 指定したセルを起点として一番下に値が入っているセルの行番号を返す
' `mustContinue`: Trueの時は連続するセル群の一番下を返す
Public Function GetLastRow(wb As Workbook, sheetName As String, startCell As String, Optional mustContinue As Boolean = False) As ObjResult
  On Error GoTo ErrHandler
  Dim lastRow As Long
  With wb.Worksheets(sheetName)
    If mustContinue Then
      lastRow = .Range(startCell).End(xlDown).Row
    Else
      lastRow = .Cells(Rows.Count, .Range(startCell).Column).End(xlUp).Row
    End If
  End With
  Set GetLastRow = Success(lastRow)
  Exit Function
ErrHandler:
  Set GetLastRow = Failure("repo.sheet.getlastrow.error", sheetName, startCell, Err.Number, Err.Description)
End Function

' 指定した範囲のセル内容をすべて削除する
Public Function ClearCells(wb As Workbook, sheetName As String, cellAddress As String) As ObjResult
  On Error GoTo ErrHandler
  With wb.Worksheets(sheetName).Range(cellAddress)
    .Value = ""
    .Validation.Delete
  End With
  Set ClearCells = Success()
  Exit Function
ErrHandler:
  Set ClearCells = Failure("repo.sheet.clearcells.error", sheetName, cellAddress, Err.Number, Err.Description)
End Function


' 指定セルにN秒で消えるメッセージを登録する
Public Function RegistCellMsg(wb As Workbook, sheetName As String, cellAddress As String, msg As String, Optional duration As Long = 5) As ObjResult
  On Error GoTo ErrHandler
  Dim targetCell As Range
  Set targetCell = wb.Worksheets(sheetName).Range(cellAddress)
  
  ' メッセージをセルに表示
  targetCell.Value = msg
  
  ' `duration`秒後にメッセージを消す
  Application.OnTime Now + TimeValue("00:00:" & Right("0" & duration, 2)), "'OnTimeHelper.ResetCell """ & sheetName & """,""" & cellAddress & """'"
  
  Set RegistCellMsg = Success()
  Exit Function
ErrHandler:
  Set RegistCellMsg = Failure("repo.sheet.registcellmsg.error", sheetName, cellAddress, Err.Number, Err.Description)
End Function


' 印刷範囲を設定する
Public Function RegistPrintArea(wb As Workbook, sheetName As String, startCell As String, endCell As String) As ObjResult
  On Error GoTo ErrHandler
  Dim ws As Worksheet
  Set ws = wb.Worksheets(sheetName)
  
  Dim targetStartCell As String
  Dim targetEndCell As String
  targetStartCell = IIf(startCell = "", "A1", startCell)
  targetEndCell = IIf(endCell = "", ws.Cells(ws.Rows.Count, ws.Columns.Count).Address, endCell)
  
  ws.PageSetup.PrintArea = targetStartCell & ":" & targetEndCell
  
  Set RegistPrintArea = Success()
  Exit Function
ErrHandler:
  Set RegistPrintArea = Failure("repo.sheet.registprintarea.error", sheetName, Err.Number, Err.Description)
End Function