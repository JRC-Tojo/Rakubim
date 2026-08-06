VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function Run()
  Dim Ts As TestSuite
  Set Ts = New TestSuite
  Ts.Description = "TestEncrypt"

  With Ts.Test("Encrypt/Decrypt")
    Const target = "‚±‚ê‚ÍƒeƒXƒg"

    Dim secretRepo As RepoSecret
    Set secretRepo = New RepoSecret
    Dim encryptedResult As ObjResult
    Dim decryptedResult As ObjResult

    Set encryptedResult = secretRepo.Encrypt(target)
    .IsOk encryptedResult.isOK
    If Not encryptedResult.isOK Then Exit Function

    Set decryptedResult = secretRepo.Decrypt(encryptedResult.value)
    .IsOk decryptedResult.isOK
    If Not decryptedResult.isOK Then Exit Function
  
    .IsEqual decryptedResult.value, target
  End With

  Set Run = Ts
End Function