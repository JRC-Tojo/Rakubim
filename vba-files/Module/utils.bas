Attribute VB_Name = "utils"
' Dictに値を追加する際にすでにKeyが存在していないかを確認し、存在した場合は追加をスキップする
Function checkDictAppend(dict As Object, key As String, ByVal appendValue As String)
  If dict.Exists(key) = False Then
    dict.Add key, appendValue
  End If
  
  Set checkDictAppend = dict
End Function

' 引数で指定したすべての値が同じ値であることを確認する
Function Equal(ParamArray arr() As Variant) As Boolean
  For Each x In arr
    If x <> arr(0) Then
        Exit Function
    End If
  Next
  
  Equal = True
End Function

' 範囲指定したセルに格納された情報をCollectionで返す
Function Range2Collection(sheetName As String, rangeArg As String) As Collection
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


' 指定した範囲に格子を描画する
Function MakeLattice(workSheetName As String, startCell As String, endCell As String)
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
Function MakeBoldLattice(workSheetName As String, startCell As String, endCell As String)
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
Function ResetLines(workSheetName As String, startCell As String, endCell As String)
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
