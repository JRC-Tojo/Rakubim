VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjResultErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ObjResultErr: structured error information
Private p_key As String
Private p_message As String
Private p_args As Variant

Public Property Get Key() As String
    Key = p_key
End Property

Public Property Get Message() As String
    Message = p_message
End Property

Public Property Get Args() As Variant
    Args = p_args
End Property

Public Function Init(key As String, Optional ByVal args As Variant) As ObjResultErr
    p_key = key
    p_message = message
    If IsMissing(args) Then
        p_args = Empty
    Else
        p_args = args
    End If
    Set Init = Me
End Function
