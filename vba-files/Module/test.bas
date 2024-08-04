Attribute VB_Name = "test"
Sub RegTest()
  Dim numReg As RegExp
  Set numReg = New RegExp
  Dim tmpMatches As MatchCollection
  With numReg
    .Pattern = ".+(?=[[i‚»‚Ì|i][‚O-‚X]+j|[i‚»‚Ì|i][‚O-‚X]+[`||][‚O-‚X]+j|i‚»‚Ì[‚O-‚X]+j[`||]i‚»‚Ì[‚O-‚X]+j]|[‚O-‚X]+$)"
    .Global = True
  End With

  Const name = "471-20011_İŒv}iP3-P4ƒ‰[ƒƒ“j 1.pdf"
  Set tmpMatches = numReg.Execute(StrConv(name, vbWide))
  Debug.Print tmpMatches.item(0)

End Sub

Sub RegTest2()
  Dim numReg As RegExp
  Set numReg = New RegExp
  Dim tmpMatches As MatchCollection
  With numReg
    .Pattern = "^([‚O-‚X]+|*)+"
    .Global = True
  End With

  Const name = "P3-P4ƒ‰[ƒƒ“\‘¢ˆê”Ê}i‚»‚Ì‚Pj"
  Debug.Print numReg.Replace(StrConv(name, vbWide), "")

End Sub

Sub HasItemTest()
  Dim list As Collection
  Set list = New Collection

  list.Add 1, "A"
  list.Add 2, "B"
  list.Add 3, "C"

  Debug.Print HasItem(list, "D")

End Sub

Sub TestArgMax()
  Dim arr As Collection
  Set arr = New Collection
  arr.Add 500
  arr.Add 100
  arr.Add 200

  Debug.Print ArgMax(arr)(1)

  Debug.Print arr.item(ArgMax(arr)(1))

End Sub

Sub PrintHash()
  Dim hash As String
  Dim calcUtil As New CalcUtils
  hash = calcUtil.HASH_SHA256("test")
  MsgBox hash, , "test‚ÌƒnƒbƒVƒ…’l"
End Sub

Sub numberTest()

  Const dummyName = "21033-0011~0021_•iì‰wŠX‹æ–k“’n‰º‹ë‘Ìˆê”Ê}i‚»‚Ì‚U~‚P‚Ujc’f} 230131 Ú²±³Ä1 (5)(1)"

  Dim data As Variant

  data = getPaperNumber(dummyName)

  Debug.Print data(0)
  Debug.Print data(1)
End Sub

Sub dictTest()
  Dim returnDic As Object
  Set returnDic = CreateObject("Scripting.Dictionary")

  returnDic.Add "test", 1
  returnDic.Add "ƒeƒXƒg", 2

  Debug.Print returnDic.item("test")
End Sub

Sub printKeyValues(dict As Object)
  For Each dkey In dict
    Debug.Print dkey & " : " & dict(dkey)
    Next
End Sub

Sub emptyTest()
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.RenamePage)
    Debug.Print .Cells(1, "E") = ""
    Debug.Print .Cells(1, "E") = 0
  End With
End Sub

' w’è‚µ‚½‚Q‚Â‚Ì•¶š—ñ‚Ìˆê’vŠ„‡‚ğ•Ô‚·
Sub SequentialMatching()
  Dim seq As SequenceMatcher
  Dim seg1 As String, seg2 As String

  Set seq = New SequenceMatcher
  seg1 = "‰¡—À‚o‚SÚ×}"
  seg2 = "|»’n•¢Ú×}"
  'seg2 = "‰¡—À‚o‚S"

  Call seq.set_seqs(seg1, seg2)

  Debug.Print "---RATIO---"
  Debug.Print seq.ratio
  'Debug.Print "---QUICK_RATIO---"
  'Debug.Print seq.quick_ratio
  'Debug.Print "---REAL_QUICK_RATIO---"
  'Debug.Print seq.real_quick_ratio
End Sub

Sub fileObjTest()
  Dim file As FileObj
  Set file = New FileObj

  file.analyze ("471-20011_İŒv}iP3-P4ƒ‰[ƒƒ“j 0–Ú˜^1.pdf")
  Debug.Print file.paperName
  Call printCollectionItems(file.PaperNumbers)
End Sub

Sub CleanerTest()
  Dim cleaner As FileNameCleaner
  Set cleaner = New FileNameCleaner
  Call cleaner.AnalyzeNames(Range2Collection(Worksheets("–Ú˜^“ü—Í"), "B2:B1000"))

  Debug.Print cleaner.NameCleaner("P3-P4ƒ‰[ƒƒ“ P4‹´‹r ’†‘w—ÀÚ×}i‚»‚Ì‚Pj")
End Sub
