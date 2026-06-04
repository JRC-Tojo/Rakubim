VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestYearConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function Run()
  Dim Ts As TestSuite
  Set Ts = New TestSuite
  Ts.Description = "TestYearConverter"

  With Ts.Test("seireki")
    Dim converter As YearConverter
    Set converter = New YearConverter
    
    ' 西暦が渡された場合はそのまま返す
    .IsEqual converter.GetSeireki(2026), 2026
    .IsEqual converter.GetSeireki(2020), 2020
    
    ' 和暦のテスト (令和)
    .IsEqual converter.GetSeireki(8), 2026  ' 令和8年
    .IsEqual converter.GetSeireki(1), 2019  ' 令和1年
    
    ' 和暦のテスト (平成)
    .IsEqual converter.GetSeireki(31), 2019  ' 平成31年
    .IsEqual converter.GetSeireki(30), 2018  ' 平成30年
    
    ' 存在しない年号
    .IsEqual converter.GetSeireki(100), -1
  End With

  Set Run = Ts
End Function