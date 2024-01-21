VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalcUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Function HASH_SHA256(str As String) As String
  Dim sha256m As Object
  Dim utf8 As Object
  Dim bytes() As Byte
  Dim hash() As Byte
  Dim i As Integer
  Dim res As String

  Set utf8 = CreateObject("System.Text.UTF8Encoding")
  bytes = utf8.GetBytes_4(str)

  Set sha256m = CreateObject("System.Security.Cryptography.SHA256Managed")
  hash = sha256m.ComputeHash_2((bytes))

  For i = LBound(hash) To UBound(hash)
      res = res & LCase(Right("0" & Hex(hash(i)), 2))
  Next i

  HASH_SHA256 = LCase(res)
End Function