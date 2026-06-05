VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepoSecret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' 引数で与えらえた文字列を暗号化して返す
Public Function Encrypt(targetText As String) As ObjResult
  On Error GoTo ErrHandler
  Dim encrypter As EncryptDes
  Set encrypter = New EncryptDes
  
  Dim encryptedTxt As String
  encryptedTxt = encrypter.EncryptStringTripleDES(targetText, getEncryptKey())
  
  Set Encrypt = Success(encryptedTxt)
  Exit Function
ErrHandler:
  Set Encrypt = Failure("repo.secret.encrypt.error", Err.Number, Err.Description)
End Function


' 引数で与えられた文字列を復号して返す
Public Function Decrypt(targetEncrypted As String) As ObjResult
  On Error GoTo ErrHandler
  Dim encrypter As EncryptDes
  Set encrypter = New EncryptDes
  
  Dim decryptedTxt As String
  decryptedTxt = encrypter.DecryptStringTripleDES(targetEncrypted, getEncryptKey())

  Set Decrypt = Success(decryptedTxt)
  Exit Function
ErrHandler:
  Set Decrypt = Failure("repo.secret.decrypt.error", Err.Number, Err.Description)
End Function


' ユーザー名の取得
Private Function getUserName() As String
  Dim wsh As Object
  Dim execObj As Object
  Set wsh = CreateObject("WScript.Shell")

  Const command = "whoami"
  Set execObj = wsh.exec("%ComSpec% /c " & command)
  getUserName = Trim(execObj.stdOut.ReadAll)
End Function


' 暗号化に用いるKeyを取得
Private Function getEncryptKey() As String
  getEncryptKey = Split(getUserName(), "\")(0)
End Function