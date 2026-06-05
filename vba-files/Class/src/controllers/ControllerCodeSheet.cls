VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControllerCodeSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function Run() As ObjResult
  Dim sheetService As ServiceSheet
  Set sheetService = New ServiceSheet

  ' コードシートをインポート
  Dim importRes As ObjResult
  Set importRes = sheetService.ImportCodeSheet()
  If (Not importRes.isOK) Then
    Set Run = importRes
    Exit Function
  End If

  ' コードのプルダウンを設置
  Dim pulldownRes As ObjResult
  Set pulldownRes = sheetService.SetCodesPulldown()
  If (Not pulldownRes.isOK) Then
    Set Run = pulldownRes
    Exit Function
  End If
  
  Set Run = Success()
End Function