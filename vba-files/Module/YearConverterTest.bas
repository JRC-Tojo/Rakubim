Attribute VB_Name = "YearConverterTest"

' GetSeireki関数の単体テスト
Public Sub TestGetSeireki()
  Dim converter As YearConverter
  Set converter = New YearConverter
  
  ' 西暦が渡された場合はそのまま返す
  Debug.Assert converter.GetSeireki(2026) = 2026
  Debug.Assert converter.GetSeireki(2020) = 2020
  
  ' 和暦のテスト (令和)
  Debug.Assert converter.GetSeireki(8) = 2026  ' 令和8年
  Debug.Assert converter.GetSeireki(1) = 2019  ' 令和1年
  
  ' 和暦のテスト (平成)
  Debug.Assert converter.GetSeireki(31) = 2019  ' 平成31年
  Debug.Assert converter.GetSeireki(30) = 2018  ' 平成30年
  
  ' 存在しない年号
  Debug.Assert converter.GetSeireki(100) = -1
  
  MsgBox "All tests passed!"
End Sub