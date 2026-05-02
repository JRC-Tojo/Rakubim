VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCluster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function Run()
  Dim Ts As TestSuite
  Set Ts = New TestSuite
  Ts.Description = "TestCluster"

  Dim sortedTargets As ObjList
  Set sortedTargets = New ObjList
  ' 小さい値（10前後）と大きい値（100前後）の混合（ソート済みの前提）
  Call sortedTargets.InitItems(8, 9, 10, 11, 12, 95, 98, 102, 105, 110)
  
  With Ts.Test("Otu")
    Dim otu As Otu
    Set otu = New Otu

    Dim threshold As Double
    threshold = otu.SplitThreshold(sortedTargets.ToArray())
    Dim splitIdx As Long
    splitIdx = otu.ArgSplit(sortedTargets.ToArray())

    .IsEqual threshold, 53.5
    .IsEqual splitIdx, 4
  End With

  Set Run = Ts
End Function