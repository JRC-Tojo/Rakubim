Attribute VB_Name = "CollectionUtils"
' ˜A”Ô‚Ì”z—ñ‚ğ¶¬‚·‚é
Function RangeCollection(startNum As Integer, endNum As Integer)
  Dim returns As Collection
  Set returns = New Collection
  
  For i = 0 To endNum - startNum
    returns.Add (startNum + i)
  Next i
  
  Set RangeCollection = returns
End Function