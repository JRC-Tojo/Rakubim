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
  Dim lastIdx As Integer: lastIdx = 1
  
  If checkFill4Mokuroku() Then
    ' 前のデータをリセット
    Call ResetData("ファイル名変更", "A2:B10000")
    
    If Sheets("入力").OnlyCADBtn.Value <> True Then
      lastIdx = pdfRenamer()
    End If
    If Sheets("入力").OnlyPDFBtn.Value <> True Then
      Call dwgRenamer(lastIdx)
    End If
  Else
    MsgBox _
      "図面名とPDF、DWGファイルの対応がすべての図面に対して定義されていません" & vbCrLf & _
      "目録入力シートのPDFファイルとDWGファイルの列をすべて埋めてください"
  End If
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

' 入力内容のリセット
Private Sub ResetBtn_Click()
  Dim sheets As SheetObj
  Set sheets = New SheetObj
  Call ResetData(sheets.RenamePage, "A" & INIT_ROWNUM_MOKUROKU & ":B10000")
End Sub