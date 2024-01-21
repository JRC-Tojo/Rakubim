VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Property Get FirstInputPage() As String
  FirstInputPage = "入力"
End Property

Public Property Get MokurokuPage() As String
  MokurokuPage = "目録入力"
End Property

Public Property Get RenamePage() As String
  RenamePage = "ファイル名変更"
End Property

Public Property Get DataTablePage() As String
  DataTablePage = "リンク表"
End Property

Public Property Get TeamCodePage() As String
  TeamCodePage = "業務機関コード"
End Property

