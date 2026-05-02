VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServicePaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ServicePaper.LoadPapers: 図面名を読み込み、ModelPaper のコレクションを返す
' 引数: analyzer (ServiceNameAnalyzer) - 図面名解析サービス
'       sheetService (ServiceSheet) - シート操作サービス
' 戻り値: ObjResult - 成功時は ObjList(ModelPaper) を返す
Public Function LoadPapers(analyzer As ServiceNameAnalyzer, sheetService As ServiceSheet) As ObjResult
  ' 図番列と図名列の２列を読み取り
  Dim numRes As ObjResult
  Dim nameRes As ObjResult
  Set numRes = sheetService.ReadPaperNumbersInMatching()
  Set nameRes = sheetService.ReadPaperNamesInMatching()
  If Not IsResultAllOK(numRes, nameRes) Then
    Set LoadPapers = PickResultErr(Nothing, numRes, nameRes)
    Exit Function
  End If

  ' 図面の番号と名称の生値を ModelPaper コレクションに変換
  Dim outColl As ObjResult
  Set outColl = buildModelPapers(analyzer, numRes.value, nameRes.value)

  Set LoadPapers = outColl
End Function

' Private helper: 生データ ObjList を受け取り ModelPaper の ObjResult of ObjList を返す
' 役割を小さい関数に分割することでテストと可読性を改善する
Private Function buildModelPapers(analyzer As ServiceNameAnalyzer, numbers As ObjList, names As ObjList) As ObjResult
  Dim paperObjs As ObjList
  Set paperObjs = New ObjList
  
  ' numbers と names の要素数を確認
  If (numbers.Count = 0 Or names.Count = 0) Then
    Set buildModelPapers = Failure("service.paper.build.emptySheet")
    Exit Function
  End If
  If numbers.Count <> names.Count Then
    Set buildModelPapers = Failure("service.paper.build.mismatch")
    Exit Function
  End If
  
  ' 各要素を ModelPaper に変換してコレクションに追加
  Dim i As Long
  Dim res As ObjResult
  Dim info As ModelAnalyzed
  Dim pm As ModelPaper
  For i = 0 To numbers.Count - 1
    Set res = analyzer.AnalyzeName(names.At(i), False)
    If res.isOK Then
      Set info = res.value
      Set pm = New ModelPaper
      Call pm.InitModel(numbers.At(i), names.At(i), info.AnalyzedName, info.SnippedName, info.PaperNumbers)
      Call paperObjs.Append(pm)
    End If
  Next i

  Set buildModelPapers = Success(paperObjs)
End Function
