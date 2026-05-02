VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelKeyValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private inner_Key As Variant
Private inner_Value As Variant

Public Function InitModel(key As Variant, value As Variant)
  If IsObject(key) Then
    Set inner_Key = key
  Else
    inner_Key = key
  End If

  If IsObject(value) Then
    Set inner_Value = value
  Else
    inner_Value = value
  End If

  Set InitModel = Me
End Function

Public Property Get Key() As Variant
  If IsObject(inner_Key) Then
    Set Key = inner_Key
  Else
    Key = inner_Key
  End If
End Property

Public Property Get Value() As Variant
  If IsObject(inner_Value) Then
    Set Value = inner_Value
  Else
    Value = inner_Value
  End If
End Property