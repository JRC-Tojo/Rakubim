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
  Dim r As Routes
  Set r = New Routes

  Dim suggestResult As ObjResult
  Set suggestResult = r.SuggestNewFileNames()

  If (Not suggestResult.isOK) Then
    Call r.ShowErrorMsg(suggestResult.err)
  End If
End Sub

' 最終的なファイル名にリネーム
Sub ChangeFileNamesBtn_Click()
  If checkFill4File() And checkFill4Mokuroku() Then
    Dim duplicatedFileName As String
    duplicatedFileName = checkDuplicate4File()
    If (duplicatedFileName <> "") Then
      MsgBox _
        "変更後ファイル名「" & duplicatedFileName & "」が重複していたため，名前変更を行いませんでした．" & vbCrLf & _
        "変更後ファイル名称を修正して再度お試しください．" _
        , vbExclamation, "らくびむ - 変更後ファイル名の重複 -"
      Exit Sub
    End If

    Dim skippedPaths As Collection
    Set skippedPaths = New Collection
    Dim oldFileName As String
    Dim newFileName As String
    For rowNum = 2 To getRanameFileCounts() + 1
      oldFileName = ""
      newFileName = ""
      Call Renamer(rowNum, oldFileName, newFileName)
      If (newFileName = "") Then
        skippedPaths.Add(oldFileName)
      End If
    Next rowNum

    If (skippedPaths.Count > 0) Then
      Dim pathTxts As String
      For Each p In skippedPaths
        pathTxts = pathTxts & "・" & p & vbCrLf
      Next p 

      MsgBox _
        "以下のファイルは存在しなかった、または既に開かれていたため、リネーム処理が行われませんでした。" & vbCrLf & _
        "成果品フォルダが正しく設定されていることを確認し、問題が解決しない場合はご自身のデスクトップ上に成果品ファイルをダウンロードしてお試しください。" & vbCrLf & pathTxts _
        , vbExclamation, "らくびむ - 存在しないパス -" 
    End If
  Else
    MsgBox _
      "旧ファイル名に対応する新ファイル名がすべての項目に対して定義されていません" & vbCrLf & _
      "ファイル名変更シートの色が塗られている箇所に新ファイル名を定義してください" _
      , vbExclamation, "らくびむ - 新ファイル定義エラー -"
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