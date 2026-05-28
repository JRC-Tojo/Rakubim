VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepoSheet_Lattice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' Œrü•`‰æ‚Ì”Ä—pŠÖ”
Private Function setRuledLines( _
  targetRange As Range, _
  directions As ObjList, _
  Optional lStyle As Long = xlContinuous, _
  Optional lColor As Long = 0, _
  Optional lWeight As long = xlThin _
)
  Dim i As Long
  For i = 0 To directions.Count - 1
    With targetRange.Borders(directions.At(i))
      .LineStyle = lStyle
      If (lStyle <> xlNone) Then
        .ColorIndex = lColor
        .TintAndShade = 0
        .Weight = lWeight
      End If
    End With
  Next i
End Function


' ƒZƒ‹”ÍˆÍî•ñ‚ğ\’z‚·‚é
Private Function buildRange(wb As Workbook, workSheetName As String, startCell As String, endCell As String)
  Dim ws As Worksheet
  Set ws = wb.Worksheets(workSheetName)

  Dim targetStartCell As String
  Dim targetEndCell As String
  targetStartCell = IIf(startCell = "", "A1", startCell)
  targetEndCell = IIf(endCell = "", ws.Cells(ws.Rows.Count, ws.Columns.Count).Address, endCell)
  
  Set buildRange = ws.Range(targetStartCell & ":" & targetEndCell)
End Function


' w’è‚µ‚½”ÍˆÍ‚ÉŠiq‚ğ•`‰æ‚·‚é
Public Function MakeGrid(wb As Workbook, workSheetName As String, startCell As String, endCell As String)
  Dim dirs As ObjList
  Set dirs = New ObjList
  Call dirs.InitItems(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
  Call setRuledLines(buildRange(wb, workSheetName, startCell, endCell), dirs)
End Function


' w’è‚µ‚½”ÍˆÍ‚ÌŠO˜g‚É‘¾‚ß‚ÌŒrü‚ğ•`‰æ‚·‚é
Public Function MakeBoldOutline(wb As Workbook, workSheetName As String, startCell As String, endCell As String)
  Dim dirs As ObjList
  Set dirs = New ObjList
  Call dirs.InitItems(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
  Call setRuledLines(buildRange(wb, workSheetName, startCell, endCell), dirs, lWeight:=xlThick)
End Function


' w’è‚µ‚½”ÍˆÍ‚ÌŒrü‚ğ‚·‚×‚Äíœ‚·‚é
Public Function ClearGrid(wb As Workbook, workSheetName As String, startCell As String, endCell As String)
  Dim dirs As ObjList
  Set dirs = New ObjList
  Call dirs.InitItems(xlDiagonalDown, xlDiagonalUp, xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
  Call setRuledLines(buildRange(wb, workSheetName, startCell, endCell), dirs, lStyle:=xlNone)
End Function