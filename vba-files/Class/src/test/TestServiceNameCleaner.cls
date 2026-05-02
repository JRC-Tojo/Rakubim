VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestServiceNameCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function Run()
  Dim Ts As TestSuite: Set Ts = New TestSuite
  Ts.Description = "TestServiceNameCleaner"

  With Ts.Test("Placeholder_CleanExamples")
    .IsOk True
  End With

  Set Run = Ts
End Function
