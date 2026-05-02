VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepoEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public EM As ModelEvents

Public Function Init() As RepoEvents
    If EM Is Nothing Then
        Set EM = New ModelEvents
        Call EM.InitModel
    End If
    Set Init = Me
End Function
