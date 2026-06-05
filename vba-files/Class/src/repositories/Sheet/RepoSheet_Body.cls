VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepoSheet_Body"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' 指定したシートの存在確認
Public Function IsExistsSheet(wb As Workbook, sheetName As String) As ObjResult
  Dim sht As Worksheet
  For Each sht In wb.Worksheets
    If sht.Name = sheetName Then
      Set IsExistsSheet = Success(True)
      Exit Function
    End If
  Next sht
  Set IsExistsSheet = Success(False)
End Function


' シートの新規作成
Public Function Create(wb As Workbook, sheetName As String, Optional visible As Boolean = True) As Worksheet
  If IsExistsSheet(wb, sheetName).value Then Exit Function
  
  Dim ws As Worksheet
  Set ws = wb.Sheets.Add(After:=Sheets(Sheets.Count))
  ws.Name = sheetName
  ws.Visible = visible

  Set Create = ws
End Function


' 外部ファイルから特定のシートを取り込む
Public Function CopySheetFromOuter(wb As Workbook, outerPath As ObjPath, sheetName As String, Optional visible As Boolean = True) As ObjResult
  On Error GoTo ErrHandler
  Dim sourceBook As Workbook
  Dim filePath As Variant
  Dim copiedSheet As Worksheet

  ' シートコピーの前後で選択位置が変わってしまう問題があるため、処理前の選択位置を保存しておく
  Dim originalSheet As Worksheet
  Dim originalRange As Range
  Set originalSheet = ActiveSheet
  Set originalRange = Selection
  
  ' 画面更新を一時停止
  Application.ScreenUpdating = False
  
  ' 外部ブックを開く
  filePath = outerPath.FullPath
  Set sourceBook = Workbooks.Open(filePath, ReadOnly:=True)
    
  ' 名前を指定して末尾にコピー
  sourceBook.Sheets(sheetName).Copy After:=wb.Sheets(wb.Sheets.Count)
  
  ' コピーしたシートを非表示にする
  If (Not visible) Then
    Set copiedSheet = wb.Sheets(wb.Sheets.Count)
    copiedSheet.Visible = xlSheetHidden
  End If
  
  ' 外部ブックを閉じる
  sourceBook.Close SaveChanges:=False
  
  ' 選択位置を復元する
  originalSheet.Select
  originalRange.Select
  
  ' 画面更新を再開
  Application.ScreenUpdating = True

  Set CopySheetFromOuter = Success()
  Exit Function
ErrHandler:
  Set CopySheetFromOuter = Failure("repo.sheet.copysheetouter.error", outerPath.FullPath, sheetName, Err.Number, Err.Description)
End Function