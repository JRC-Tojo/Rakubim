VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilsSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' îÕàÕéwíËÇµÇΩÉZÉãÇ…äiî[Ç≥ÇÍÇΩèÓïÒÇCollectionÇ≈ï‘Ç∑
Public Function Range2Collection(sheetName As String, rangeArg As String) As Collection
  Dim coll As Collection
  Set coll = New Collection

  Dim item As Variant
  With Worksheets(sheetName)
    For Each item In .Range(rangeArg)
      If item <> "" Then
        coll.Add item
      End If
    Next
  End With

  Set Range2Collection = coll
End Function


' éwíËÇµÇΩîÕàÕÇ…äiéqÇï`âÊÇ∑ÇÈ
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

' éwíËÇµÇΩîÕàÕÇÃäOògÇ…ëæÇﬂÇÃårê¸Çï`âÊÇ∑ÇÈ
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

' éwíËÇµÇΩîÕàÕÇÃårê¸ÇÇ∑Ç◊ÇƒçÌèúÇ∑ÇÈ
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
