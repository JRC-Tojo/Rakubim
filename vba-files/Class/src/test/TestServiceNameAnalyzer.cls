VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestServiceNameAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function Run()
  Dim Ts As TestSuite: Set Ts = New TestSuite
  Ts.Description = "TestServiceNameAnalyzer"

  Dim cases As New Collection

  Dim c1_1 As New Dictionary: Set c1_1 = initNoLastNumber(c1_1)
  c1_1("name") = "exists-headNumber_no-lastNumber_pdf"
  c1_1("stemName") = "123-45678-0001_\‘¢ˆê”Ê}"
  cases.Add c1_1

  Dim c1_2 As New Dictionary: Set c1_2 = initNoLastNumber(c1_2)
  c1_2("name") = "no-headNumber_no-lastNumber_pdf"
  c1_2("stemName") = "\‘¢ˆê”Ê}"
  cases.Add c1_2

  Dim c1_3 As New Dictionary: Set c1_3 = initNoLastNumber(c1_3)
  c1_3("name") = "no-headNumber_no-lastNumber_with-date_pdf"
  c1_3("stemName") = "\‘¢ˆê”Ê}_260101"
  cases.Add c1_3

  Dim c2_1 As New Dictionary: Set c2_1 = initExistsLastNumberPdf(c2_1)
  c2_1("name") = "exists-headNumber_exists-lastNumber_pdf"
  c2_1("stemName") = "123-45678-0003_XXXXÚ×}i‚»‚Ì‚Pj"
  cases.Add c2_1

  Dim c2_2 As New Dictionary: Set c2_2 = initExistsLastNumberPdf(c2_2)
  c2_2("name") = "no-headNumber_exists-lastNumber_pdf"
  c2_2("stemName") = "XXXXÚ×}i‚»‚Ì1j"
  cases.Add c2_2

  Dim c2_3 As New Dictionary: Set c2_3 = initExistsLastNumberPdf(c2_3)
  c2_3("name") = "no-headNumber_exists-lastNumber_with-date_pdf"
  c2_3("stemName") = "XXXXÚ×}i‚»‚Ì1j_260101"
  cases.Add c2_3

  Dim c3_1 As New Dictionary: Set c3_1 = initExistsLastNumbersCad(c3_1)
  c3_1("name") = "exists-headNumber_exists-lastNumbers1_cad"
  c3_1("stemName") = "123-45678-0003-0005_XXXXÚ×}i‚»‚Ì‚P`‚Rj"
  cases.Add c3_1

  Dim c3_2 As New Dictionary: Set c3_2 = initExistsLastNumbersCad(c3_2)
  c3_2("name") = "exists-headNumber_exists-lastNumbers2_cad"
  c3_2("stemName") = "123-45678-0003-0005_XXXXÚ×}i‚»‚Ì‚Pj`i‚»‚Ì‚Rj"
  cases.Add c3_2

  Dim c3_3 As New Dictionary: Set c3_3 = initExistsLastNumbersCad(c3_3)
  c3_3("name") = "exists-headNumber_exists-lastNumbers3_cad"
  c3_3("stemName") = "123-45678-0003-0005_XXXXÚ×}i‚P`‚Rj"
  cases.Add c3_3

  Dim c3_4 As New Dictionary: Set c3_4 = initExistsLastNumbersCad(c3_4)
  c3_4("name") = "exists-headNumber_exists-lastNumbers4_cad"
  c3_4("stemName") = "123-45678-0003-0005_XXXXÚ×}i‚Pj`i‚Rj"
  cases.Add c3_4

  Dim c3_5 As New Dictionary: Set c3_5 = initExistsLastNumbersCad(c3_5)
  c3_5("name") = "exists-headNumber_exists-lastNumbers5_cad"
  c3_5("stemName") = "123-45678-0003-0005_XXXXÚ×}i1-3j"
  cases.Add c3_5

  Dim c3_6 As New Dictionary: Set c3_6 = initExistsLastNumbersCad(c3_6)
  c3_6("name") = "exists-headNumber_exists-lastNumbers6_cad"
  c3_6("stemName") = "123-45678-0003-0005_XXXXÚ×}i‚»‚Ì1, 2, 3j"
  cases.Add c3_6

  Dim c4_1 As New Dictionary: Set c4_1 = initContinuousNumberPdf(c4_1)
  c4_1("name") = "exists-headNumber_continuous1_pdf"
  c4_1("stemName") = "123-45678-0000_İŒv}ˆê® (1)"
  cases.Add c4_1

  Dim c4_2 As New Dictionary: Set c4_2 = initContinuousNumberPdf(c4_2)
  c4_2("name") = "exists-headNumber_continuous2_pdf"
  c4_2("stemName") = "123-45678-0000_İŒv}ˆê® 1"
  cases.Add c4_2

  Dim c4_3 As New Dictionary: Set c4_3 = initContinuousNumberPdf(c4_3)
  c4_3("name") = "exists-headNumber_continuous3_pdf"
  c4_3("stemName") = "123-45678-0000_İŒv}ˆê®‚P"
  cases.Add c4_3

  Dim c4_4 As New Dictionary: Set c4_4 = initContinuousNumberPdf(c4_4)
  c4_4("name") = "exists-headNumber_continuous4_pdf"
  c4_4("stemName") = "123-45678-0000_İŒv}ˆê® Part1"
  c4_4("expectedSnippedName") = "İŒv}ˆê® Part"
  c4_4("expectedAnalyzedName") = "İŒv}ˆê®‚o‚`‚q‚s"
  cases.Add c4_4

  Dim c100 As New Dictionary
  Dim c100Nums As ObjSet: Set c100Nums = New ObjSet
  c100("name") = "special_includeMultiple_dwg"
  c100("stemName") = "01-05_Œv‰æŠî€}i‚»‚Ì1-3jEˆê”Ê}i‚»‚Ì1-2jEİŒvğŒ“™"
  c100("useContinuousNumber") = False
  c100("expectedSnippedName") = "Œv‰æŠî€}Eˆê”Ê}EİŒvğŒ“™"
  c100("expectedAnalyzedName") = "Œv‰æŠî€}Eˆê”Ê}EİŒvğŒ“™"
  Set c100("expectedPaperNumbers") = c100Nums.InitItems(1, 2, 3)
  cases.Add c100

  Dim c101 As New Dictionary
  Dim c101Nums As ObjSet: Set c101Nums = New ObjSet
  c101("name") = "special_includeTrimChrInName_dwg"
  c101("stemName") = "01-05_P3-P4_\‘¢ˆê”Ê},Œv‰æŠî€}"
  c101("useContinuousNumber") = False
  c101("expectedSnippedName") = "P3-P4_\‘¢ˆê”Ê},Œv‰æŠî€}"
  c101("expectedAnalyzedName") = "‚o‚R‚o‚S\‘¢ˆê”Ê}Œv‰æŠî€}"
  Set c101("expectedPaperNumbers") = c101Nums
  cases.Add c101

  Dim i As Long
  Dim cName As String
  For i = 1 To cases.Count
    cName = cases(i)("name")
    Call runCaseNameAnalyzer(cases(i), Ts.Test(cName))
  Next i

  Set Run = Ts
End Function

Private Sub runCaseNameAnalyzer(tc As Dictionary, t As Object)
  Dim nameAnalyzer As ServiceNameAnalyzer: Set nameAnalyzer = New ServiceNameAnalyzer
  Dim result As ObjResult: Set result = nameAnalyzer.AnalyzeName(tc("stemName"), IIf(IsNull(tc("useContinuousNumber")), False, tc("useContinuousNumber")))
  t.IsOk result.isOK
  If result.isOK Then
    Dim analyzed As ModelAnalyzed: Set analyzed = result.value
    t.IsEqual analyzed.SnippedName, tc("expectedSnippedName")
    t.IsEqual analyzed.AnalyzedName, tc("expectedAnalyzedName")
    Dim expectedNumbers As ObjSet: Set expectedNumbers = tc("expectedPaperNumbers")
    Dim actualNumbers As ObjSet: Set actualNumbers = analyzed.PaperNumbers
    t.IsOk expectedNumbers.IsEqualTo(actualNumbers)
  End If
End Sub

Private Function initNoLastNumber(dict As Dictionary) As Dictionary
  Dim expextedNums As ObjSet: Set expextedNums = New ObjSet
  dict("useContinuousNumber") = False
  dict("expectedSnippedName") = "\‘¢ˆê”Ê}"
  dict("expectedAnalyzedName") = "\‘¢ˆê”Ê}"
  Set dict("expectedPaperNumbers") = expextedNums
  Set initNoLastNumber = dict
End Function

Private Function initExistsLastNumberPdf(dict As Dictionary) As Dictionary
  Dim expextedNums As ObjSet: Set expextedNums = New ObjSet
  dict("useContinuousNumber") = False
  dict("expectedSnippedName") = "XXXXÚ×}"
  dict("expectedAnalyzedName") = "‚w‚w‚w‚wÚ×}"
  Set dict("expectedPaperNumbers") = expextedNums.InitItems(1)
  Set initExistsLastNumberPdf = dict
End Function

Private Function initExistsLastNumbersCad(dict As Dictionary) As Dictionary
  Dim expextedNums As ObjSet: Set expextedNums = New ObjSet
  dict("useContinuousNumber") = False
  dict("expectedSnippedName") = "XXXXÚ×}"
  dict("expectedAnalyzedName") = "‚w‚w‚w‚wÚ×}"
  Set dict("expectedPaperNumbers") = expextedNums.InitItems(1, 2, 3)
  Set initExistsLastNumbersCad = dict
End Function

Private Function initContinuousNumberPdf(dict As Dictionary) As Dictionary
  Dim expextedNums As ObjSet: Set expextedNums = New ObjSet
  dict("useContinuousNumber") = True
  dict("expectedSnippedName") = "İŒv}ˆê®"
  dict("expectedAnalyzedName") = "İŒv}ˆê®"
  Set dict("expectedPaperNumbers") = expextedNums.InitItems(1)
  Set initContinuousNumberPdf = dict
End Function
