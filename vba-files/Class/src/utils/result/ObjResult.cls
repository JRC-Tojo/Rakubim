VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ObjResult: wrapper for API return values
' Properties:
'  - isOK As Boolean
'  - value As Variant
'  - err As ObjResultErr

Private p_isOK As Boolean
Private p_value As Variant
Private p_err As ObjResultErr

Public Property Get isOK() As Boolean
  isOK = p_isOK
End Property

Public Property Get value()
  If IsObject(p_value) Then
    Set value = p_value
  Else
    value = p_value
  End If
End Property

Public Property Get err() As ObjResultErr
  Set err = p_err
End Property

Public Function InitSuccess(val As Variant) As ObjResult
  p_isOK = True
  If IsObject(val) Then
    Set p_value = val
  Else
    p_value = val
  End If
  Set InitSuccess = Me
End Function

Public Function InitFailure(e As ObjResultErr) As ObjResult
  p_isOK = False
  Set p_value = Nothing
  Set p_err = e
  Set InitFailure = Me
End Function