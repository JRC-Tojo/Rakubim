VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 目録入力シートにおいて実際のファイル名称を補完する
' 実際のファイル名を読み込み、目録の図面名称と対応させる
Sub ReadFileNamesBtn_Click()
  Dim pdfObjs As New Collection
  Dim dwgObjs As New Collection
  
  ' 実際のファイルを読み込み、図面とパスのDictを形成
  ' TODO: ラジオボタンの選択によって処理を変えるコードのリファクタリング
  If Sheets("入力").OnlyCADBtn.Value <> True Then
    Set pdfObjs = GetPDFCollection()
  End If
  If Sheets("入力").OnlyPDFBtn.Value <> True Then
    Set dwgObjs = GetDWGCollection()
  End If

  Call FileNames2drawingNames(pdfObjs, dwgObjs)
End Sub

' 目録入力シートのファイル名称を削除する（左）
Sub ResetMokuroku_Left_Click()
  Call ResetMokuroku("left")
End Sub

' 目録入力シートのファイル名称を削除する（右）
Sub ResetMokuroku_Right_Click()
  Call ResetMokuroku("right")
End Sub
