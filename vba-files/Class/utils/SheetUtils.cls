VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' 範囲指定したセルに格納された情報をCollectionで返す
Public Function Range2Collection(ws As Worksheet, rangeArg As String, Optional ignoreEmptyCell As Boolean = True) As Collection
  Dim coll As Collection
  Set coll = New Collection

  Dim item As Variant
  With ws
    For Each item In .Range(rangeArg)
      If item <> "" Then
        coll.Add item
      ElseIf Not ignoreEmptyCell Then
        coll.Add ""
      End If
    Next
  End With

  Set Range2Collection = coll
End Function

' 指定した配列の値を指定したセルを１つ目として下方向に記述する
Public Sub Collection2Cell(col As Collection, sheetName As String, startCell As String)
  Dim arr As Variant: ReDim arr(1 To col.Count + 2)
  Dim i As Long: i = 1
  
  Dim colVal As Variant
  For Each colVal In col
    arr(i) = colVal
    i = i + 1
  Next colVal
  
  Call Variants2Cell(arr, sheetName, startCell)
End Sub

' 指定した配列の値を指定したセルを１つ目として下方向に記述する
Public Sub Variants2Cell(arr As Variant, sheetName As String, startCell As String)
  with Worksheets(sheetName)
    .Range(startCell, .Range(startCell).Offset(UBound(arr) - 1, 0)).Value = WorksheetFunction.Transpose(arr)
  End With
End Sub

' 列番号→列名
Public Function ColumnIdx2Name(ByVal colNum As Long) As String
  ColumnIdx2Name = Split(Columns(colNum).address, "$")(2)
End Function

' 列名→列番号
Public Function ColumnName2Idx(ByVal colName As String) As Long
  ColumnName2Idx = Columns(colName).column
End Function

' 指定した範囲に格子を描画する
Public Function MakeLattice(workSheetName As String, startCell As String, endCell As String)
  With Worksheets(workSheetName).Range(startCell & ":" & endCell)
    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
  End With
End Function

' 指定した範囲の外枠に太めの罫線を描画する
Public Function MakeBoldLattice(workSheetName As String, startCell As String, endCell As String)
  With Worksheets(workSheetName).Range(startCell & ":" & endCell)
    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
  End With
End Function

' 指定した範囲の罫線をすべて削除する
Public Function ResetLines(workSheetName As String, startCell As String, endCell As String)
  With Worksheets(workSheetName).Range(startCell & ":" & endCell)
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone
    .Borders(xlEdgeLeft).LineStyle = xlNone
    .Borders(xlEdgeTop).LineStyle = xlNone
    .Borders(xlEdgeBottom).LineStyle = xlNone
    .Borders(xlEdgeRight).LineStyle = xlNone
    .Borders(xlInsideVertical).LineStyle = xlNone
    .Borders(xlInsideHorizontal).LineStyle = xlNone
  End With
End Function
