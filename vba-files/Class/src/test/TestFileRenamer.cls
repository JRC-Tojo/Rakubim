VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestFileRenamer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' 事業者コードのテスト用値
Private Const TestCid = "000"

' プロジェクト番号のテスト用値
Private Const TestPid = "00000"

' 構造物通し番号のテスト用値
Private Const TestSid = "00"

' テスト用の標準拡張子
Private Const TestExt = ".tst"


Public Function Run()
  Dim Ts As TestSuite
  Set Ts = New TestSuite
  Ts.Description = "TestFileRenamer"

  Dim scenarios As ObjList
  Set scenarios = New ObjList

  ' ケース1: 基本のナンバリング
  scenarios.Append buildCase( _
    "default rename and numbers", _
    setPapers( _
      setPaper("1", "図面A", "図面Ａ", Array()), _
      setPaper("5", "図面B", "図面Ｂ", Array()), _
      setPaper("10", "図面C", "図面Ｃ", Array()), _
      setPaper("100", "図面D", "図面Ｄ", Array()) _
    ), _
    Array(0, 1, 2, 3), _
    Array("0001_図面A", "0005_図面B", "0010_図面C", "0100_図面D") _
  )

  ' ケース2: 複数図面のナンバリング
  scenarios.Append buildCase( _
    "multiple paper numbers", _
    setPapers( _
      setPaper("1", "図面A（その１）", "図面Ａ", Array(1)), _
      setPaper("2", "図面A（その２）", "図面Ａ", Array(2)), _
      setPaper("3", "図面A（その３）", "図面Ａ", Array(3)), _
      setPaper("10", "図面B（その１）", "図面Ｂ", Array(1)), _
      setPaper("11", "図面B（その２）", "図面Ｂ", Array(2)), _
      setPaper("12", "図面B（その３）", "図面Ｂ", Array(3)), _
      setPaper("15", "図面B（その６）", "図面Ｂ", Array(6)) _
    ), _
    Array(0, 0, 0, 1, 1, 1, 1), _  
    Array("0001-0003_図面A（その１～その３）", "0010-0012,0015_図面B（その１～その３、その６）") _
  )

  ' ケース3: 不明な値を含むリネーム
  scenarios.Append buildCase( _
    "invalid paper number", _
    setPapers( _
      setPaper("10", "図面A（その１）", "図面Ａ", Array(1)), _
      setPaper("10-1","図面A（その１-1）", "図面Ａ", Array(-1)), _
      setPaper("11", "図面A（その２）", "図面Ａ", Array(2)), _
      setPaper("20", "図面B（その１）", "図面Ｂ", Array(1)), _
      setPaper("21","図面B（その２）", "図面Ｂ", Array(-1)), _
      setPaper("21-1", "図面B（その２-1）", "図面Ｂ", Array(2)), _
      setPaper("A1", "図面C", "図面Ｃ", Array()), _
      setPaper("A5", "図面D（その５）", "図面Ｄ", Array(5)) _
    ), _
    Array(0, 0, 0, 1, 1, 1, 2, 3), _
    Array("0010-0011_図面A（その１～その２）", "0020-XXXX_図面B（その１～その２）", "XXXX_図面C", "XXXX_図面D（その５）") _
  )
  
  ' ケース4: 複数の図面名を含むリネーム
  scenarios.Append buildCase( _
    "multiple paper name", _
    setPapers( _
      setPaper("1", "図面A", "図面Ａ", Array()), _
      setPaper("2", "図面B", "図面Ｂ", Array()), _
      setPaper("3", "図面C", "図面Ｃ", Array()), _
      setPaper("10", "図面D（その１）", "図面Ｄ", Array(1)), _
      setPaper("11", "図面D（その２）", "図面Ｄ", Array(2)), _
      setPaper("12", "図面E（その１）", "図面Ｅ", Array(1)), _
      setPaper("13", "図面E（その２）", "図面Ｅ", Array(2)), _
      setPaper("14", "図面E（その３）", "図面Ｅ", Array(3)) _
    ), _
    Array(0, 0, 0, 1, 1, 1, 1, 1), _
    Array("0001-0003_図面A、図面B、図面C", "0010-0014_図面D、図面E（その１～その３）") _
  )

  Dim scenario As ObjDict
  For Each scenario In scenarios
    Call runCase(Ts, scenario)
  Next

  Set Run = Ts
End Function


' ファイル先頭に付与されるテスト用ヘッダー
Private Property Get TestHeader() As String
  TestHeader = TestCid & "-" & TestPid & "-" & TestSid
End Property


' テスト用の仮ファイル名を取得する
Private Function GetTestFileName(idx As Long) As String
  GetTestFileName = CStr(idx) & TestExt
End Function


Private Function runCase(Ts As TestSuite, scenario As ObjDict)
  Dim fileService As ServiceFile
  Set fileService = New ServiceFile
  Dim oldToNewNames As ObjResult
  Set oldToNewNames = fileService.NewNames( _
    scenario.At("papers"), scenario.At("files"), TestCid, TestPid, TestSid _
  )

  With Ts.Test(scenario.At("title"))
    .IsOk oldToNewNames.isOK
    If Not oldToNewNames.isOK Then Exit Function

    ' 個数を確認
    .IsEqual oldToNewNames.value.Count, scenario.At("expects").Count
    If oldToNewNames.value.Count <> scenario.At("expects").Count Then Exit Function

    ' それぞれの新ファイル名を確認
    Dim oldName As String
    Dim gotNewName As String
    Dim expNewName As String
    Dim i As Long
    For i = 0 To scenario.At("expects").Count
      oldName = GetTestFileName(i)
      gotNewName = oldToNewNames.value.At(oldName)
      expNewName = scenario.At("expects").At(i)
      .IsEqual gotNewName, expNewName
    Next i
  End With
End Function


' テストケースの生成関数
' 引数：titleName (String) テストケースのタイトル
'       papers (List[Paper]) Paperオブジェクトの一覧
'       fileIdxs (Long[]) 番号をファイル名と見立てたファイル一覧．複数図面を同一のファイルに集約する場合は同じ番号を記入する
'       newFileNames (String[]) 上記のファイル番号をインデックスとして，該当する位置に想定される変更後のファイル名を記入する
' 戻り値：(Dict) テストケースの各値を格納済みの辞書
Private Function buildCase( _
  titleName As String, _
  papers As ObjList, _
  fileIdxs As Variant, _
  newFileNames As Variant _
) As ObjDict
  Dim scenario As ObjDict
  Set scenario = New ObjDict
  Dim i As Long

  Call scenario.Append("title", titleName)
  Call scenario.Append("papers", papers)

  ' 新ファイル名一覧の対応インデックス番号をファイル名と見立てて，ファイル一覧を作成
  Dim filePaths As ObjList
  Set filePaths = New ObjList
  Call filePaths.InitItemsFromArray(fileIdxs)
  Dim maxFileIdx As Long: maxFileIdx = 0
  For i = 0 To filePaths.Count - 1
    If maxFileIdx < filePaths.At(i) Then maxFileIdx = filePaths.At(i)
    Call filePaths.Update(i, GetTestFileName(filePaths.At(i)))
  Next i
  Call scenario.Append("files", filePaths)

  ' 出力結果として想定されるファイル名一覧を作成
  Dim newNames As ObjList
  Set newNames = New ObjList
  Call newNames.InitItemsFromArray(newFileNames)
  For i = 0 To newNames.Count - 1
    Call newNames.Update(i, TestHeader & "-" & newNames.At(i) & TestExt)
  Next i
  Call scenario.Append("expects", newNames)

  Debug.Assert filePaths.Count = papers.Count   ' 各図面に対応するファイル番号の指定に過不足アリ
  Debug.Assert maxFileIdx + 1 = newNames.Count  ' インデックス番号と想定される新ファイル名が対応していない

  Set buildCase = scenario
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