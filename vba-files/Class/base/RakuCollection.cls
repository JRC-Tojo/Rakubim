VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RakuCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private InnerCol As Collection

Private Sub class_initialize()
  ' èâä˙âª
  Set InnerCol = New Collection
End Sub

Public Sub Add(item As Object)
  InnerCol.Add(item)
End Sub

Public Function Item(idx As Integer)
  Item = InnerCol.Item(idx)
End Function

Public Property Get Count()
  Count = InnerCol.Count
End Property

