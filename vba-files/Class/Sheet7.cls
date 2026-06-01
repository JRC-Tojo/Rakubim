VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' 入力シートにおける処理対象のフォルダを選択する
Private Sub FolderSelectBtn_Click()
  Dim folderPath As Variant
  With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "らくびむ　- 成果品フォルダを選択してください -"
    .AllowMultiSelect = False

    ' ダイアログを表示し，キャンセルされた場合は処理を終了
    If (.Show <> True) Then Exit Sub
    
    folderPath = .SelectedItems(1)
  End With

  With WorkSheets("入力")
    .Range("E3").Value = folderPath 
  End With
End Sub

' 入力シートにおける「業務機関名」「契約番号」を自動で補完
Private Sub AutoIDInputBtn_Click()
  Dim matches As Object
  Dim reg As Object
  Set reg = CreateObject("VBScript.RegExp")
  Dim inputName As String
  Dim companyCode As String
  companyCode = ""
  
  With Worksheets("入力")
    inputName = .Range("E7").Value
    
    With Worksheets("業務機関コード")
      On Error Resume Next
      companyCode = WorksheetFunction.VLookup(Left(inputName, 2), .Range("G3:H22"), 2, False)
      If (Err.Number > 0) Then
        MsgBox "業務機関名称の取得に失敗しました．" & vbCrLf & _
                "業務機関コードのセルを選択し，契約先の業務機関をプルダウンメニューから選択してください．" _
                ,,"らくびむ　- 業務機関取得エラー -"
      End If
    End With
    
    With reg
      .Pattern = "[０-９|0-9]+"
      .Global = True
    End With
    Set matches = reg.Execute(inputName)
    
    .Range("C6").Value = companyCode
    .Range("C7").Value = Format(ClipSeireki(matches.item(0)), "00") & Format(matches.item(1), "000")
  End With
End Sub

' 和暦から西暦の下２桁を取得する
Private Function ClipSeireki(year As String)
  Dim intYear As Integer
  intYear = CInt(year)

  Dim yearConverter As YearConverter
  Set yearConverter = New YearConverter

  Dim resultSeireki As Integer
  resultSeireki = yearConverter.GetSeireki(intYear)

  If (resultSeireki = -1) Then
    MsgBox "年号の読み込みに失敗しました．" & vbCrLf & _
              "契約番号の上２桁は「00」で出力されるため，適切な番号を設定してください．" _
              ,,"らくびむ　- 年号取得エラー -"
    ClipSeireki = 0
  Else
    ClipSeireki = resultSeireki Mod 100
  End If
End Function
