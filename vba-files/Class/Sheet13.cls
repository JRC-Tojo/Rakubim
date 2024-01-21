VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 目録入力シートにおいて実際のファイル名称を補完する
' 実際のファイル名を読み込み、目録の図面名称と対応させる
Sub ReadFileNamesBtn_Click()
  Dim rowNum As Integer
  Dim matcher As Paper2FileMatcher
  Set matcher = New Paper2FileMatcher
  Dim paper As PaperObj
  Dim pdfPath As String
  Dim dwgPath As String
  
  ' 何回空欄の行が連続したのか
  Dim spaceCounter As Integer: spaceCounter = 0
    
  ' 目録を読み込み始める行番号
  rowNum = INIT_ROWNUM_MOKUROKU
  Do
    ' 目録データの取得
    If checkSpaceRow(rowNum) Then
      ' 将来的に
      ' ・目録一行空け：次の列から目録を再開
      ' ・目録二行空け：次のページから目録を再開
      ' という規則にする予定のため、spaceCount = 3までは容認
      spaceCounter = spaceCounter + 1
      If spaceCounter = 3 Then Exit Do
    Else
      spaceCounter = 0
      
      ' PaperObjを宣言
      Call paper.InitPaperObj(rowNum)
      pdfPath = matcher.GetBestFile(paper, "PDF")
      dwgPath = matcher.GetBestFile(paper, "CAD")
      
      ' 探索結果の書き込み
      WriteOldP4Mokuroku rowNum, paper.GlobalNumber, pdfPath, dwgPath
    End If
       
    ' 次処理に向けて値を更新
    rowNum = rowNum + 1
  Loop
End Sub

' 目録入力シートのファイル名称を削除する（左）
Sub ResetMokuroku_Left_Click()
  Call ResetMokuroku("left")
End Sub

' 目録入力シートのファイル名称を削除する（右）
Sub ResetMokuroku_Right_Click()
  Call ResetMokuroku("right")
End Sub
