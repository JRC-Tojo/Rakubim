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
  Dim rowNum As Integer
  Dim paperCounter As Integer
  Dim matcher As Paper2FileMatcher
  Set matcher = New Paper2FileMatcher
  Dim paper As PaperObj
  Dim pdfPath As String
  Dim dwgPath As String
  
  ' 何回空欄の行が連続したのか
  Dim spaceCounter As Integer: spaceCounter = 0

  ' ファイルが存在しない場合はエラーを表示して終了
  If (Not matcher.ExistFiles) Then
    MsgBox "読み込むファイルが見つかりません．" & vbCrLf & _
              "以下のポイントを確認したうえで再度ボタンを押してください．" & vbCrLf & vbCrLf & _
              "1. 【半角英文字で】「CAD」「PDF」という名前のフォルダを作成する" & vbCrLf & _
              "2. 上記フォルダにファイルが格納されている" & vbCrLf & _
              "3. 「入力」シートで設定した処理するファイルの種類に対応したファイルを格納している" _
              ,,"らくびむ　- 探索エラー -"
    Exit Sub
  End If
    
  ' 目録を読み込み始める行番号
  rowNum = INIT_ROWNUM_MOKUROKU
  paperCounter = 1
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
      Set paper = New PaperObj
      Call paper.InitPaperObj(rowNum, paperCounter)
      pdfPath = matcher.GetBestFile(paper, "PDF")
      dwgPath = matcher.GetBestFile(paper, "CAD")
      
      ' 探索結果の書き込み
      WriteOldP4Mokuroku rowNum, paper.GlobalNumber, pdfPath, dwgPath
    End If
       
    ' 次処理に向けて値を更新
    rowNum = rowNum + 1
    paperCounter = paperCounter + 1
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

' 目録の記入データをリセット
Private Sub ResetMokuroku(leftRight As String)
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  If leftRight = "left" Then
    Call ResetData(sheets.MokurokuPage, "A" & INIT_ROWNUM_MOKUROKU & ":E10000")
  Else
    Call ResetData(sheets.MokurokuPage, "G" & INIT_ROWNUM_MOKUROKU & ":H10000")
  End If
End Sub
