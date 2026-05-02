VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestObjDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function Run()
  Dim Ts As TestSuite
  Set Ts = New TestSuite
  Ts.Description = "TestObjDict"

  Dim d As ObjDict

  ' 新規作成と Count の確認
  With Ts.Test("InitAndCount")
    Set d = New ObjDict
    .IsEqual d.Count, 0
  End With

  ' Append と At（Variant 値）の基本動作
  With Ts.Test("Append_At_Variant")
    Set d = New ObjDict
    .IsOk d.Append("a", 123)
    .IsEqual d.Count, 1
    .IsOk d.Exists("a")
    .IsEqual d.At("a"), 123
  End With

  ' 同一 Key の追加は失敗すること
  With Ts.Test("Append_Duplicate")
    Set d = New ObjDict
    .IsOk d.Append("k", "v")
    .NotOk d.Append("k", "v2")
    .IsEqual d.Count, 1
    .IsEqual d.At("k"), "v"
  End With

  ' オブジェクト値が参照のまま返ることを確認
  With Ts.Test("Append_At_Object")
    Set d = New ObjDict
    Dim c1 As Collection: Set c1 = New Collection
    c1.Add "inner"
    .IsOk d.Append("o", c1)
    .IsOk (d.At("o") Is c1)
  End With
  
  ' Keys / Values / For Each (ModelKeyValue を介して) の確認
  With Ts.Test("Keys_Values_ForEach")
    Set d = New ObjDict
    Dim expKeys As Variant: expKeys = Array("k1", "k2", "k3")
    Dim expVals As Variant: expVals = Array(10, 20, 30)
    Call d.Append(expKeys(0), expVals(0))
    Call d.Append(expKeys(1), expVals(1))
    Call d.Append(expKeys(2), expVals(2))

    .IsEqual d.Keys.Count, 3
    .IsEqual d.Keys.At(0), expKeys(0)

    .IsEqual d.Values.Count, 3
    .IsEqual d.Values.At(1), expVals(1)

    ' For Each で ModelKeyValue を得ることを確認
    Dim i As Long: i = 0
    Dim kv As ModelKeyValue
    For Each kv In d
      .IsEqual expKeys(i), kv.Key
      .IsEqual expVals(i), kv.Value
      i = i + 1
    Next kv
  End With

  ' Remove のテスト
  With Ts.Test("Remove")
    Set d = New ObjDict
    Call d.Append("r", 1)
    .IsEqual d.Count, 1
    .IsOk d.Remove("r")
    .IsEqual d.Count, 0
    .NotOk d.Exists("r")
  End With

  ' At のデフォルト値の挙動（Variant と Object）
  With Ts.Test("At_Default")
    Set d = New ObjDict
    Dim defVal As Variant: defVal = "def"
    .IsEqual d.At("no", defVal), "def"

    Dim defObject As Collection: Set defObject = New Collection
    defObject.Add "x"
    .IsOk (d.At("noobj", defObject) Is defObject)
  End With

  ' Update のテスト（修正後、値が更新されること）
  With Ts.Test("Update")
    Set d = New ObjDict
    Call d.Append("u", "old")
    .IsOk d.Update("u", "new")
    .IsEqual d.At("u"), "new"
  End With

  Set Run = Ts
End Function