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
  Set matcher = GetMatcher()
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  Dim sUtill As SheetUtils
  Set sUtill = New SheetUtils
  Dim systemSettings As SystemIdObj
  Set systemSettings = New SystemIdObj
  Dim paper As PaperObj
  Dim paperNumber As String
  Dim pdfPath As String
  Dim dwgPath As String
  
  ' 何回空欄の行が連続したのか
  Dim spaceCounter As Integer: spaceCounter = 0

  ' ファイルが存在しない場合はエラーを表示して終了
  If (matcher Is Nothing) Then
    MsgBox "読み込むファイルが見つかりません．" & vbCrLf & _
              "以下のポイントを確認したうえで再度ボタンを押してください．" & vbCrLf & vbCrLf & _
              "1. 【半角英文字で】「CAD」「PDF」という名前のフォルダを作成する" & vbCrLf & _
              "2. 上記フォルダにファイルが格納されている" _
              ,,"らくびむ　- 探索エラー -"
    Exit Sub
  End If

  ' Matcherに必要な図面名称の一覧を登録する
  Call matcher.InitMatcher(sUtill.Range2Collection(Worksheets(sheets.MokurokuPage), "B2:B10000"))

  ' 図面番号が全く入っていないときには，一番最初を１番で補完する
  If (Worksheets(sheets.MokurokuPage).Range("A2").Value = "") Then
    Worksheets(sheets.MokurokuPage).Range("A2").Value = 1
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
      If spaceCounter = LIMITED_SPACE_COUNTER Then Exit Do
    Else
      spaceCounter = 0

      ' 図面番号を補完
      paperNumber = CompleteGlobalPaperNumber(rowNum)
      
      ' PaperObjを宣言
      Set paper = New PaperObj
      Call paper.InitPaperObj(Worksheets(sheets.MokurokuPage).Cells(rowNum, "B").Value, paperCounter, paperNumber)
      pdfPath = matcher.GetBestFile(paper, "PDF", systemSettings.UseContinuousNumber)
      dwgPath = matcher.GetBestFile(paper, "CAD", False)
      
      ' 探索結果の書き込み
      WriteOldP4Mokuroku rowNum, paper.GlobalNumber, pdfPath, dwgPath
    End If
       
    ' 次処理に向けて値を更新
    rowNum = rowNum + 1
    paperCounter = paperCounter + 1
  Loop

  ' 格納されたファイル名一覧をプルダウンで選択できるようにする
  Call SetFilePulldownMenu(matcher, rowNum - LIMITED_SPACE_COUNTER)

  If (paper.GlobalNumber = "") Then
    MsgBox "図面番号の補完に失敗しました．" & vbCrLf & _ 
              "【番号をすべて入力】した後，「ファイル名変更」シートへ移動してください" _ 
              ,,"らくびむ  - 補完エラー -"
    End
  End If
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


' 図面番号を補完する
' 補完した番号を戻り値として返す
Private Function CompleteGlobalPaperNumber(rowNum As Integer) As String
  Dim backId As Variant
  Dim nowId As Variant
  Dim sheets As SheetObj
  Set sheets = New SheetObj

  With Worksheets(sheets.MokurokuPage)
    backId = .Cells(rowNum - 1, "A").Value
    nowId = .Cells(rowNum, "A").Value
    If nowId = "" Then
      If IsNumeric(backId) And backId <> "" Then
        nowId = backId + 1  
      End If
    End If
  End With
  
  CompleteGlobalPaperNumber = CStr(nowId)
End Function