VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestObjSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function Run()
  Dim Ts As TestSuite: Set Ts = New TestSuite
  Ts.Description = "TestObjSet"

  With Ts.Test("InitAndCount")
    Dim l As ObjSet: Set l = New ObjSet
    Call l.InitItems("a", "b", "c")
    .IsEqual l.Count, 3
  End With

  With Ts.Test("Append_Variant")
    Dim l2 As ObjSet: Set l2 = New ObjSet
    l2.InitItems "x", "y"
    l2.Append "z"
    .IsEqual l2.Count, 3
  End With

  With Ts.Test("Append_Object")
    Dim s As ObjSet: Set s = New ObjSet
    Dim c As Collection: Set c = New Collection
    Dim d As Collection: Set d = New Collection
    c.Add "inner"
    d.Add "inner2"
    s.Append c
    s.Append d
    .IsEqual s.Count, 2
  End With

  With Ts.Test("Pop")
    Dim lp As ObjSet: Set lp = New ObjSet
    lp.InitItems 5, 8, 10
    lp.Pop 8
    .IsEqual lp.Count, 2
  End With

  With Ts.Test("Contains")
    Dim lc As ObjSet: Set lc = New ObjSet
    lc.InitItems "aa", "bb", "cc"
    .IsOk lc.Contains("bb")
    .NotOk lc.Contains("dd")
  End With

  With Ts.Test("ToArray_ToString")
    Dim lt As ObjSet: Set lt = New ObjSet
    lt.InitItems "A", "B"
    Dim arr As Variant: arr = lt.ToArray()
    .IsEqual arr(0), "A"
    .IsEqual lt.ToString(), "(A, B)"
  End With

  With Ts.Test("ForEach")
    Dim s1 As ObjSet: Set s1 = New ObjSet
    s1.InitItems "first", "second", "third"
    Dim expected As Variant: expected = Array("first", "second", "third")
    Dim i As Long: i = 0
    Dim item As Variant
    For Each item In s1
      .IsEqual item, expected(i)
      i = i + 1
    Next item
  End With

  With Ts.Test("Union")
    Dim a As ObjSet: Set a = New ObjSet
    Dim b As ObjSet: Set b = New ObjSet
    a.InitItems 1, 2
    b.InitItems 3
    Dim union As ObjSet: Set union = a.Union(b)
    .IsEqual union.Count, 3
  End With

  With Ts.Test("Clear_ToCollection")
    Dim s2 As ObjSet: Set s2 = New ObjSet
    s2.InitItems 1, 2
    s2.Clear
    .IsEqual s2.Count, 0
    Dim coll As Collection: Set coll = s2.ToCollection()
    .IsEqual coll.Count, 0
  End With

  With Ts.Test("Intersection")
    Dim a2 As ObjSet: Set a2 = New ObjSet
    Dim b2 As ObjSet: Set b2 = New ObjSet
    a2.InitItems 1, 2, 3
    b2.InitItems 2, 4
    Dim inter As ObjSet: Set inter = a2.Intersection(b2)
    .IsEqual inter.Count, 1
    .IsOk inter.Contains(2)
  End With

  With Ts.Test("Difference")
    Dim a3 As ObjSet: Set a3 = New ObjSet
    Dim b3 As ObjSet: Set b3 = New ObjSet
    a3.InitItems 1, 2, 3
    b3.InitItems 2
    Dim diff As ObjSet: Set diff = a3.Difference(b3)
    .IsEqual diff.Count, 2
    .IsOk diff.Contains(1)
    .IsOk diff.Contains(3)
  End With

  With Ts.Test("IsSubsetIsSuperset")
    Dim a4 As ObjSet: Set a4 = New ObjSet
    Dim b4 As ObjSet: Set b4 = New ObjSet
    a4.InitItems 1, 2
    b4.InitItems 1, 2, 3
    .IsOk a4.IsSubsetOf(b4)
    .IsOk b4.IsSupersetOf(a4)
  End With

  With Ts.Test("IsDisjoint")
    Dim a5 As ObjSet: Set a5 = New ObjSet
    Dim b5 As ObjSet: Set b5 = New ObjSet
    a5.InitItems 1, 2
    b5.InitItems 3, 4
    .IsOk a5.IsDisjointWith(b5)
    b5.Append 2
    .NotOk a5.IsDisjointWith(b5)
  End With

  Set Run = Ts
End Function
