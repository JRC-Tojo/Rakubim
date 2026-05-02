VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepoSheet_Validation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' 対象のシート名と開始点位置のセル座標からPulldownsシートにおける書き込み列名を返す
Private Function GetPulldownInfo2ColName(sheetName As String, cellAddress As String) As String
  Dim pulldownID As String
  pulldownID = sheetName & "/" & Range(cellAddress).Item(1).Address()
  Select Case pulldownID
    Case "目録入力/$G$2"
      GetPulldownInfo2ColName = "A"
    Case "目録入力/$H$2"
      GetPulldownInfo2ColName = "B"
    Case Else
      Debug.Assert False  ' 登録されていない箇所でPulldownが呼び出された
  End Select
End Function

' プルダウンメニューを指定したRangeに設置する
Public Function Pulldown(wb As Workbook, ByVal sheetName As String, ByVal cellAddress As String, elems As ObjList, Optional ErrTitle As String = "", Optional ErrMessage As String = "")
  Dim repoSheetIo As RepoSheet_IO
  Set repoSheetIo = New RepoSheet_IO

  ' pulldown要素の列挙用シートに列挙内容を書き込み
  ' (直接Formula1に各要素を指定すると，要素文字列内にカンマが含まれる場合に誤作動する)
  Dim targetColName As String
  targetColName = GetPulldownInfo2ColName(sheetName, cellAddress)
  Call repoSheetIo.WriteColFromCell(wb, "Pulldowns", targetColName & "1", elems)

  With wb.Worksheets(sheetName).Range(cellAddress).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
      Formula1:="=Pulldowns!$" & targetColName & "$1:$" & targetColName & "$" & elems.Count
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ErrTitle
    .InputMessage = ""
    .ErrorMessage = ErrMessage
    .IMEMode = xlIMEModeNoControl
    .ShowInput = True
    .ShowError = ErrTitle <> ErrMessage
  End With
End Function