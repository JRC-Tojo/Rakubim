VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepoErrMsgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' エラーの表示言語
Private inner_Lang As String

' エラーキーに対応するアイコン群（エラーマーク以外を設定する場合）
Public inner_Icons As ObjDict

Private Sub class_initialize()
  inner_Lang = "JaJp"
  
  Set inner_Icons = New ObjDict
  With inner_Icons
    ' 特別にアイコンを設定する場合は以下に列挙
    .Append "HogeHoge", vbExclamation
  End With
End Sub

Private Function GetContainer()
  Select Case inner_Lang
    Case "JaJp": Set GetContainer = New RepoErrMsgs_JaJp
    Case Else: Debug.Assert False  ' 非対応の言語が指定された
  End Select
End Function

Public Function SetLanguage(lang As String)
  inner_Lang = lang
End Function

Public Function GetTitle(key As String, Optional ByVal args As Variant = Empty) As String
  Dim container As Object
  Set container = GetContainer()

  Dim title As String
  title = container.MsgboxTitle & "　- " & container.inner_Msgs.At(key)(0) & " -"
  GetTitle = CraftedTxts(title, args)
End Function

Public Function GetMsg(key As String, Optional ByVal args As Variant = Empty) As String
  Dim container As Object
  Set container = GetContainer()

  Dim msg As String
  msg = container.inner_Msgs.At(key)(1)
  GetMsg = CraftedTxts(msg, args)
End Function

Public Function GetIcon(key As String) As Long
  GetIcon = inner_Icons.At(key, defValue:=vbCritical)
End Function


' 各言語の表示候補メッセージを加工して表示可能データとして返す
Private Function CraftedTxts(sourceMsg As String, ByVal args As Variant) As String
  Dim crafted As String: crafted = sourceMsg
  crafted = EmbedArgs(crafted, args)
  crafted = ReplaceBrake(crafted)
  CraftedTxts = crafted
End Function


' メッセージ中に変数を埋め込んで返す
' 埋め込み箇所の特殊文字は `{}` とする
Private Function EmbedArgs(msgStr As String, ByVal args As Variant) As String
  EmbedArgs = msgStr
  If IsEmpty(args) Then Exit Function
  
  Dim divedMsg As Variant
  Dim divedCount As Long
  Dim argsCount As Long
  
  divedMsg = Split(msgStr, "{}")
  divedCount = Ubound(divedMsg) - Lbound(divedMsg) + 1
  argsCount = Ubound(args) - Lbound(args) + 1
  
  If (divedCount = 0) Then Exit Function
  if (divedCount - 1 <> argsCount) Then Exit Function

  Dim i As Long: i = 0
  Dim embedMsg As String: embedMsg = divedMsg(0)
  For i = 0 To argsCount - 1
    embedMsg = embedMsg & args(i) & divedMsg(i + 1)
  Next i

  EmbedArgs = embedMsg 
End Function


' 改行文字\nをvbCrLfに置換する
Private Function ReplaceBrake(msgStr As String) As String
  Dim splitedMsgs As Variant
  splitedMsgs = Split(msgStr, "\n")

  Dim SplitedMsgList As ObjList
  Set SplitedMsgList = New ObjList
  Call SplitedMsgList.InitItemsFromArray(splitedMsgs)

  ReplaceBrake = SplitedMsgList.Join(vbCrLf)
End Function