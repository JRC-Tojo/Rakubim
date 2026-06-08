VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' 和暦（平成X年の「平成」の部分）
Private inner_Wareki As String

' 年号（平成X年の「X」の部分）
Private inner_year As Integer


Public Function SetYear(wareki As String, year As Integer)
  inner_Wareki = wareki
  inner_year = year
End Function

Public Property Get Wareki() As String
  Wareki = inner_Wareki
End Property

Public Property Get year() As Integer
  year = inner_year
End Property

