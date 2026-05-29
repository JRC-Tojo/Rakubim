Attribute VB_Name = "SheetProcess"

' “ü—ÍÏ‚İƒf[ƒ^‚ğíœ‚·‚é
Function ResetData(sheetName As String, deleteRange As String)
  With Worksheets(sheetName).Range(deleteRange)
    .Value = ""
    .Validation.Delete
  End With
End Function
