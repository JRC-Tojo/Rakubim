VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestServiceMatching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function Run()
  Dim Ts As TestSuite: Set Ts = New TestSuite
  Ts.Description = "TestServiceMatching"

  Dim cases As New Collection

  ' ƒP[ƒX1: ˜A‘±PDF‚Ìƒ}ƒbƒ`ƒ“ƒO
  cases.Add buildCase( _
    "continuous pdf", _
    setPapers( _
      setPaper("1", "}–ÊA", "}–Ê‚`", Array()), _
      setPaper("2", "}–ÊB", "}–Ê‚a", Array()) _
    ), _
    setFiles( _
      setFile("}–Êˆê® Part1.pdf", "}–Êˆê®", Array(1)), _
      setFile("}–Êˆê® Part2.pdf", "}–Êˆê®", Array(2)) _
    ), _
    setExpects("}–Êˆê® Part1.pdf", "}–Êˆê® Part2.pdf"), _
    useSeq:=True _
  )

  ' ƒP[ƒX2: }–Ê“à‚Ì”Ô†‚Å‚Ìƒ}ƒbƒ`ƒ“ƒO
  cases.Add buildCase( _
    "match dwg based on paperNumber", _
    setPapers( _
      setPaper("1", "”CˆÓ}–Êi‚»‚Ì‚P)", "”CˆÓ}–Ê", Array(1)) _
    ), _
    setFiles( _
      setFile("”CˆÓ}–Ê1.dwg", "”CˆÓ}–Ê", Array(1)), _
      setFile("”CˆÓ}–Ê2.dwg", "”CˆÓ}–Ê", Array(2)) _
    ), _
    setExpects("”CˆÓ}–Ê1.dwg") _
  )

  ' ƒP[ƒX3: ’Z‚¢–¼‘OA”Ô†–³‚µCAD–¼‚Ìƒ}ƒbƒ`ƒ“ƒO
  cases.Add buildCase( _
    "match short dwg names and no number dwg names", _
    setPapers( _
      setPaper("10", "å\i‚»‚Ì‚Pj", "å\", Array(1)), _
      setPaper("21", "\‘¢ˆê”Ê}i‚»‚Ì‚Qj", "\‘¢ˆê”Ê}", Array(2)) _
    ), _
    setFiles( _
      setFile("åŒ….dwg", "åŒ…", Array()), _
      setFile("\‘¢ˆê”Ê}.dwg", "\‘¢ˆê”Ê}", Array()), _
      setFile("å\.dwg", "å\", Array()) _
    ), _
    setExpects("å\.dwg", "\‘¢ˆê”Ê}.dwg") _
  )

  ' ƒP[ƒX4: ˜A‘±PDFi“r’†”Ô†j‚Ìƒ}ƒbƒ`ƒ“ƒO
  cases.Add buildCase( _
    "continuous pdf with dropped head numbers", _
    setPapers( _
      setPaper("1", "}–ÊA", "}–Ê‚`", Array()), _
      setPaper("2", "}–ÊB", "}–Ê‚a", Array()) _
    ), _
    setFiles( _
      setFile("}–Êˆê® Part6.pdf", "}–Êˆê®", Array(6)), _
      setFile("}–Êˆê® Part7.pdf", "}–Êˆê®", Array(7)) _
    ), _
    setExpects("}–Êˆê® Part6.pdf", "}–Êˆê® Part7.pdf"), _
    useSeq:=True _
  )

  Dim i As Long
  Dim cName As String
  For i = 1 To cases.Count
    cName = cases(i)("name")
    Call runCaseAssert(cases(i), Ts.Test(cName))
  Next i

  Set Run = Ts
End Function

Private Sub runCaseAssert(tc As Dictionary, t As TestCase)
  Dim svc As New ServiceMatching
  Dim matchedFileNames As ObjResult
  Set matchedFileNames = svc.Match(tc("papers"), tc("files"), tc("useSeq"))
  
  t.IsOk matchedFileNames.isOK
  If Not matchedFileNames.isOK Then Exit Sub
  
  Dim getMatched As ObjList
  Set getMatched = matchedFileNames.value
  Dim expects As ObjList
  Set expects = tc("expects")

  ' ŒÂ”‚ÌŠm”F
  Call t.IsEqual(getMatched.Count, expects.Count)
  If getMatched.Count <> expects.Count Then Exit Sub
  
  ' ŒÂ•Ê‚Ì—v‘f‚ÌŠm”F
  Dim i As Long
  For i = 0 To expects.Count - 1
    Call t.IsEqual(getMatched.At(i), expects.At(i))
  Next i
End Sub


Private Function buildCase( _
  name As String, _
  papers As ObjList, _
  files As ObjList, _
  expects As ObjList, _
  Optional useSeq As Boolean = False _
)
  Dim d As Dictionary
  Set d = New Dictionary
  
  d("name") = name
  Set d("papers") = papers
  Set d("files") = files
  Set d("expects") = expects
  d("useSeq") = useSeq

  Set buildCase = d
End Function


' papers: Array(ModelPaper) -> List[ModelPaper]
Private Function setPapers(ParamArray papers() As Variant) As ObjList
  Dim pL As ObjList
  Set pL = New ObjList

  Call pL.InitItemsFromArray(papers)

  Set setPapers = pL
End Function


Private Function setPaper(pIdx As String, pName As String, aName As String, numStrs As Variant) As ModelPaper
  Dim paper As ModelPaper
  Set paper = New ModelPaper

  Dim numStrSet As ObjSet
  Set numStrSet = New ObjSet

  Call numStrSet.InitItemsFromArray(numStrs)
  Call paper.InitModel(pIdx, pName, aName, aName, numStrSet)

  Set setPaper = paper
End Function


' files: Array(ModelFile) -> List[ModelFile]
Private Function setFiles(ParamArray files() As Variant) As ObjList
  Dim fL As ObjList
  Set fL = New ObjList

  Call fL.InitItemsFromArray(files)

  Set setFiles = fL
End Function


Private Function setFile(fName As String, aName As String, numStrs As Variant) As ModelFile
  Dim file As ModelFile
  Set file = New ModelFile

  Dim numStrSet As ObjSet
  Set numStrSet = New ObjSet
  Call numStrSet.InitItemsFromArray(numStrs)

  Dim fPath As ObjPath
  Set fPath = New ObjPath
  Call fPath.Init(fName)

  Call file.InitModel(fPath, aName, aName, numStrSet)

  Set setFile = file
End Function


Private Function setExpects(ParamArray fNames() As Variant) As ObjList
  Dim fStrL As ObjList
  Set fStrL = New ObjList

  Call fStrL.InitItemsFromArray(fNames)

  Set setExpects = fStrL
End Function