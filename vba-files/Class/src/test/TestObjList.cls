VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestObjList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function Run()
  Dim Ts As TestSuite
  Set Ts = New TestSuite
  Ts.Description = "TestObjList"

  ' 初期化と Count のテスト
  With Ts.Test("InitAndCount")
    Dim l1 As ObjList: Set l1 = New ObjList
    Call l1.InitItems("a", "b", "c")
    .IsEqual l1.Count, 3

    Dim arr As Variant: arr = Array(10, 20, 30)
    Dim l2 As ObjList: Set l2 = New ObjList
    Call l2.InitItemsFromArray(arr)
    .IsEqual l2.Count, 3
    .IsEqual l2.At(1), 20

    Dim coll As Collection: Set coll = New Collection
    coll.Add "x"
    coll.Add "y"
    Dim l3 As ObjList: Set l3 = New ObjList
    Call l3.InitItemsFromCollection(coll)
    .IsEqual l3.Count, 2
    .IsEqual l3.At(1), "y"
  End With

  ' Append と At のテスト（正順・逆順インデックス）
  With Ts.Test("Append_At_Variant")
    Dim l As ObjList: Set l = New ObjList
    l.InitItems "x", "y"
    l.Append "z"
    .IsEqual l.Count, 3
    .IsEqual l.At(0), "x"
    .IsEqual l.At(-1), "z"
    .IsEqual l.At(10, ""), ""
  End With

  ' Pop と Update のテスト
  With Ts.Test("Pop_Update")
    Dim lp As ObjList: Set lp = New ObjList
    lp.InitItems 1, 2, 3
    lp.Pop 1
    .IsEqual lp.Count, 2
    lp.Update 0, 99
    .IsEqual lp.At(0), 99
    .IsEqual lp.At(1), 3
  End With

  ' Contains のテスト
  With Ts.Test("Contains")
    Dim lc As ObjList: Set lc = New ObjList
    lc.InitItems "aa", "bb", "cc"
    .IsOk lc.Contains("bb")
    .NotOk lc.Contains("dd")
  End With

  ' ToArray と ToString のテスト
  With Ts.Test("ToArray_ToString")
    Dim lt As ObjList: Set lt = New ObjList
    lt.InitItems "A", "B"
    Dim outArr As Variant: outArr = lt.ToArray()
    .IsEqual outArr(0), "A"
    .IsEqual lt.ToString(), "[A, B]"
  End With

  ' At のテスト（オブジェクト参照がそのまま返ること）
  With Ts.Test("Append_At_Object")
    Dim lo As ObjList: Set lo = New ObjList
    Dim c As Collection: Set c = New Collection
    Dim d As Collection: Set d = New Collection
    c.Add "inner"
    d.Add "inner2"
    lo.Append c
    lo.Append d
    Dim o As Object: Set o = lo.At(-1)
    .IsOk (o Is d)
  End With

  ' For Each ループのテスト（NewEnum が正しく機能するか）
  With Ts.Test("ForEach")
    Dim lf As ObjList: Set lf = New ObjList
    lf.InitItems "first", "second", "third"
    Dim expected As Variant: expected = Array("first", "second", "third")
    Dim i As Long: i = 0
    Dim item As Variant
    For Each item In lf
      .IsEqual item, expected(i)
      i = i + 1
    Next item
  End With

  ' 別のObjListを結合するテスト
  With Ts.Test("Combine")
    Dim a As ObjList: Set a = New ObjList
    Dim b As ObjList: Set b = New ObjList
    a.InitItems 1, 2
    b.InitItems 3
    Call a.Combine(b)
    .IsEqual a.Count, 3
    .IsEqual a.At(2), 3
  End With

  ' 逆順 (Reverse) のテスト
  With Ts.Test("Reverse")
    Dim lr As ObjList: Set lr = New ObjList
    lr.InitItems "a", "b", "c"
    Dim rev As ObjList: Set rev = lr.Reverse(False)
    .IsEqual rev.Count, 3
    .IsEqual rev.At(0), "c"

    Dim rev2 As ObjList: Set rev2 = lr.Reverse(True)
    .IsEqual lr.At(0), "c"
    .IsEqual rev2.At(0), "c"

    Dim lo2 As ObjList: Set lo2 = New ObjList
    Dim o1 As Collection: Set o1 = New Collection
    Dim o2 As Collection: Set o2 = New Collection
    o1.Add "one"
    o2.Add "two"
    lo2.Append o1
    lo2.Append o2
    Dim revObj As ObjList: Set revObj = lo2.Reverse(False)
    Dim outObj As Object: Set outObj = revObj.At(0)
    .IsOk (outObj Is o2)

    Dim revObj2 As ObjList: Set revObj2 = lo2.Reverse(True)
    .IsOk (lo2.At(0) Is o2)
    .IsOk (revObj2.At(0) Is o2)
  End With

  ' Clear と ToCollection のテスト
  With Ts.Test("Clear_ToCollection")
    Dim lc2 As ObjList: Set lc2 = New ObjList
    lc2.InitItems 1, 2
    lc2.Clear
    .IsEqual lc2.Count, 0
    Dim collOut As Collection: Set collOut = lc2.ToCollection()
    .IsEqual collOut.Count, 0
  End With

  Set Run = Ts
End Function
