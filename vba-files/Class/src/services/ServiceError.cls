VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServiceError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ResultErr型の情報に基づいてエラーメッセージを表示する
Public Function ShowErrMsg(err As ObjResultErr, lang As String)
  Dim repo As RepoErrMsgs
  Set repo = New RepoErrMsgs
  Call repo.SetLanguage(lang)

  MsgBox GetErrMsg(repo, err), GetMsgStyle(repo, err), GetErrTitle(repo, err)
End Function


' エラーに対応するタイトルを返す
Private Function GetErrTitle(repo As RepoErrMsgs, err As ObjResultErr)
  GetErrTitle = repo.GetTitle(err.Key, err.Args)
End Function


' エラーに対応する説明文を返す
Private Function GetErrMsg(repo As RepoErrMsgs, err As ObjResultErr)
  GetErrMsg = repo.GetMsg(err.Key, err.Args)
End Function


' メッセージボックスの体裁を決定する
Private Function GetMsgStyle(repo As RepoErrMsgs, err As ObjResultErr)
  GetMsgStyle = vbOKOnly + repo.GetIcon(err.Key) + vbMsgBoxSetForeground
End Function