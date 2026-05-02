VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControllerError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function ShowErrorMsg(err As ObjResultErr)
  Dim config As ServiceSystemConfig
  Set config = New ServiceSystemConfig

  ' システム設定の読み込み
  Dim resConfig As ObjResult
  Set resConfig = config.LoadUserInput()
  If Not resConfig.isOK Then
    Set Run = resConfig
    Exit Function
  End If

  ' エラーメッセージを表示
  Dim errService As ServiceError
  Set errService = New ServiceError
  Call errService.ShowErrMsg(err, config.MsgLanguage)
End Function