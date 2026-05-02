VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestRepoSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function Run()
  Dim Ts As TestSuite: Set Ts = New TestSuite
  Ts.Description = "TestRepoSheet"

  ' RepoSheet の IO メソッドのテスト
  With Ts.Test("Write and Read Column")
    ' 準備: テスト用シート名
    Dim sheetRepo As RepoSheet: Set sheetRepo = New RepoSheet
    Dim sheetName As String: sheetName = "test_repo_io"

    ' 既存シートがあれば削除（事前クリーン）
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' テスト用シートを作成
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = sheetName

    ' 値を書き込むための ObjList を作成
    Dim vals As ObjList: Set vals = New ObjList
    Call vals.InitItems("alpha", "beta", "gamma")

    Dim r As ObjResult
    ' 書き込み
    Set r = sheetRepo.WriteColFromCell(sheetName, "A1", vals)
    .IsOk r.isOK

    ' 単一セル読み取りの検証
    Set r = sheetRepo.ReadCell(sheetName, "A1", True)
    .IsOk r.isOK
    .IsEqual r.value, "alpha"

    ' 列読み取りの検証
    Set r = sheetRepo.ReadColumn(sheetName, "A1")
    .IsOk r.isOK
    Dim readList As ObjList: Set readList = r.value
    .IsEqual readList.Count, 3
    .IsEqual readList.At(0), "alpha"
    .IsEqual readList.At(1), "beta"
    .IsEqual readList.At(2), "gamma"

    ' 後片付け: テスト用シートを削除
    Application.DisplayAlerts = False
    Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
  End With

  Set Run = Ts
End Function
