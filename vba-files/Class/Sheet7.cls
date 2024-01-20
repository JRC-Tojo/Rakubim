VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 入力シートにおける「業務機関名」「契約番号」を自動で補完
Private Sub AutoIDInputBtn_Click()
  Dim matches As Object
  Dim reg As Object
  Set reg = CreateObject("VBScript.RegExp")
  Dim inputName As String
  Dim companyCode As String
  companyCode = ""
  
  With Worksheets("入力")
    inputName = .Range("E4").Value
    
    With Worksheets("業務機関コード")
      companyCode = WorksheetFunction.VLookup(Left(inputName, 2), .Range("G3:H21"), 2, False)
    End With
    
    With reg
      .Pattern = "[０-９|0-9]+"
      .Global = True
    End With
    Set matches = reg.Execute(inputName)
    
    .Range("C3").Value = companyCode
    .Range("C4").Value = Format(Wareki2Seireki(matches.item(0)), "00") & Format(matches.item(1), "000")
  End With
End Sub

' 和暦から西暦の下２桁を取得する
Private Function Wareki2Seireki(year As String)
  Dim intYear As Integer
  intYear = CInt(year)
  
  If HeiseiBtn.Value = True Then
    Wareki2Seireki = IIf(intYear - 12 >= 0, intYear - 12, 100 + intYear - 12)
  ElseIf ReiwaBtn.Value = True Then
    Wareki2Seireki = intYear + 18
  ElseIf SeirekiBtn.Value = True Then
    Wareki2Seireki = Right(year, 2)
  Else
    Wareki2Seireki = 0
  End If
End Function
