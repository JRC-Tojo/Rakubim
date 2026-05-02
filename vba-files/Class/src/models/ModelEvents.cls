VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Progress As Progress

Public Function InitModel() As ModelEvents
    If Progress Is Nothing Then
        Set Progress = New Progress
    End If
    Set InitModel = Me
End Function
