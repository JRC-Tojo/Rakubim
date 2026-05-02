VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestObjRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function Run()
  Dim Ts As TestSuite: Set Ts = New TestSuite
  Ts.Description = "TestObjRegex"

  With Ts.Test("IsMatch")
    Dim r As ObjRegex: Set r = New ObjRegex
    .IsOk r.IsMatch("\d+", "abc123")
    .NotOk r.IsMatch("xyz", "abcdef")
  End With

  With Ts.Test("Replace")
    Dim r2 As ObjRegex: Set r2 = New ObjRegex
    Dim rep As String: rep = r2.Replace("a+", "aaabca", "Z")
    .IsOk (InStr(rep, "Z") > 0)
  End With

  With Ts.Test("GetMatches_Focus_Drop")
    Dim r3 As ObjRegex: Set r3 = New ObjRegex
    Dim coll As ObjList: Set coll = r3.GetMatches("(\d+)", "a12b34")
    .IsEqual coll.Count, 2
    If coll.Count >= 1 Then
      Dim m As Match: Set m = coll.At(0)
      .IsEqual r3.FocusMatchedPosition(m, "a12b34"), "12"
      .IsEqual r3.DropMatchedPosition(m, "a12b34"), "ab34"
    End If
  End With

  Set Run = Ts
End Function
