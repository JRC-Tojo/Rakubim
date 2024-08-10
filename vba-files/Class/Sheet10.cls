VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ファイル名変更タブを作成
Sub GetFileNamesBtn_Click()
  Dim systemSettings As SystemIdObj
  Set systemSettings = New SystemIdObj
  Dim lastIdx As Integer: lastIdx = 1
  
  If checkFill4Mokuroku() Then
    ' 前のデータをリセット
    Call ResetData("ファイル名変更", "A2:B10000")

    ' 新ファイル名称の生成器を宣言
    Dim creater As NewNameCreater
    Set creater = New NewNameCreater
    Call creater.InitCreater(systemSettings.CompanyID, systemSettings.ProjectID, systemSettings.StructureID)

    ' 入力情報の読み取り
    Call ReadMokurokuInput(creater)

    ' 書き出し
    Call WriteOld2NewNames(creater)
  Else
    MsgBox _
      "図面名とPDF、DWGファイルの対応がすべての図面に対して定義されていません" & vbCrLf & _
      "目録入力シートのPDFファイルとDWGファイルの列をすべて埋めてください"
  End If
End Sub

' 目録入力で入力された情報でデータを構築
Private Sub ReadMokurokuInput(creater As NewNameCreater)
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  ' 何回空欄の行が連続したのか
  Dim spaceCounter As Integer: spaceCounter = 0
  ' 目録を読み込み始める行番号
  Dim rowNum As Integer: rowNum = INIT_ROWNUM_MOKUROKU
  
  Dim paper As PaperObj
  Dim paperCounter As Integer: paperCounter = 1
  Dim cadFile As FileObj
  Dim pdfFile As FileObj
  With Worksheets(sheets.MokurokuPage)
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
        Call paper.InitPaperObj(.Cells(rowNum, "B").Value, paperCounter, .Cells(rowNum, "A").Value)
        Call creater.Papers.Add(paper.Id, paper)
        ' FileObjを構築
        Set pdfFile = New FileObj
        Call pdfFile.InitFileObj(.Cells(rowNum, "G").Value)
        Set cadFile = New FileObj
        Call cadFile.InitFileObj(.Cells(rowNum, "H").Value)

        ' 図面とファイルの対応関係を記録
        Call creater.SetPDFFile(pdfFile, paper.Id)
        Call creater.SetCADFile(cadFile, paper.Id)
      End If
        
      ' 次処理に向けて値を更新
      rowNum = rowNum + 1
      paperCounter = paperCounter + 1
    Loop
  End With
End Sub

Private Sub WriteOld2NewNames(creater As NewNameCreater)
  Dim sheets As SheetObj
  Dim sUtils As SheetUtils
  Set sheets = New SheetObj
  Set sUtils = New SheetUtils

  Call sUtils.Collection2Cell(creater.OldNames, sheets.RenamePage, "A2")
  Call sUtils.Collection2Cell(creater.NewNames, sheets.RenamePage, "B2")
End Sub

' 最終的なファイル名にリネーム
Sub ChangeFileNamesBtn_Click()
  If checkFill4File() And checkFill4Mokuroku() Then
    For rowNum = 2 To getRanameFileCounts() + 1
      Call Renamer(rowNum)
    Next rowNum
  Else
    MsgBox _
      "旧ファイル名に対応する新ファイル名がすべての項目に対して定義されていません" & vbCrLf & _
      "ファイル名変更シートの色が塗られている箇所に新ファイル名を定義してください"
  End If
End Sub

' PDFの新旧名称を一致させる
Sub NoRename4PDF_Click()
  Call NoRename("PDF")
  
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  Worksheets(sheets.RenamePage).Range("G6").Value = "設定OK"
  Application.OnTime Earliesttime:=Now + TimeValue("00:00:05"), Procedure:="sheet10.ResetPDFMessage"
End Sub

' CADの新旧名称を一致させる
Sub NoRename4CAD_Click()
  Call NoRename("CAD")
  
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  Worksheets(sheets.RenamePage).Range("G8").Value = "設定OK"
  Application.OnTime Earliesttime:=Now + TimeValue("00:00:05"), Procedure:="sheet10.ResetCADMessage"
End Sub

Private Sub NoRename(fType As String)
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  ' 目録を読み込み始める行番号
  Dim rowNum As Integer: rowNum = 2
  
  Dim file As FileObj
  With Worksheets(sheets.RenamePage)
    Do
      ' 空白が来たら終了
      If .Cells(rowNum, "A").Value = "" Then
        Exit Do
      End If

      ' FileObjを構築
      Set file = New FileObj
      Call file.InitFileObj(.Cells(rowNum, "A").Value)
      If (file.FileType = fType) Then
        .Cells(rowNum, "B").Value = .Cells(rowNum, "A").Value
      End If

      ' 次処理に向けて値を更新
      rowNum = rowNum + 1
    Loop
  End With
End Sub

' 入力内容のリセット
Private Sub ResetBtn_Click()
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  Call ResetData(sheets.RenamePage, "A" & INIT_ROWNUM_MOKUROKU & ":B10000")
End Sub

' 表示したメッセージを消す
Sub ResetPDFMessage()
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  Worksheets(sheets.RenamePage).Range("G6").Value = ""
End Sub
Sub ResetCADMessage()
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  Worksheets(sheets.RenamePage).Range("G8").Value = ""
End Sub