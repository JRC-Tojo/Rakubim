VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' Simple Progress object: default no-op implementations
Public Sub Start()
    ' default no-op
End Sub

Public Sub Update(pct As Double)
    ' default no-op
End Sub

Public Sub Close()
    ' default no-op
End Sub
