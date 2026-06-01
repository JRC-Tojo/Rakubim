VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControllerSheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ファイル一覧のプルダウンメニューをマッチング結果セルに設置する
' 戻り値: ObjResult[Nothing]
Public Function SetFilesPulldown() As ObjResult
  Dim fileService As ServiceFile
  Set fileService = New ServiceFile
  Dim sheetService As ServiceSheet
  Set sheetService = New ServiceSheet
  Dim config As ServiceSystemConfig
  Set config = New ServiceSystemConfig

  ' システム設定の読み込み
  Dim resConfig As ObjResult
  Set resConfig = config.LoadUserInput()
  If Not resConfig.isOK Then
    Set SetFilesPulldown = resConfig
    Exit Function
  End If

  ' ファイル群の読み込み
  Dim pdfFilePaths As ObjResult
  Set pdfFilePaths = fileService.LoadPaths(config.PdfFolderPath, config.PdfExtensions)
  Dim cadFilePaths As ObjResult
  Set cadFilePaths = fileService.LoadPaths(config.CadFolderPath, config.CadExtensions)
  If (Not IsResultAllOK(pdfFilePaths, cadFilePaths)) Then
    Set SetFilesPulldown = PickResultErr(Nothing, pdfFilePaths, cadFilePaths)
    Exit Function
  End If

  ' プルダウンメニューのセット
  Set SetFilesPulldown = sheetService.SetFilesPulldown(pdfFilePaths.value, cadFilePaths.value, config.MsgLanguage)
End Function


' 図面目録部分をリセットする
' 戻り値: ObjResult[Nothing]
Public Function ResetMokuroku() As ObjResult
  Dim sheetService As ServiceSheet
  Set sheetService = New ServiceSheet
  Set ResetMokuroku = sheetService.ResetMokuroku()
End Function


' ファイル部分をリセットする
' 戻り値: ObjResult[Nothing]
Public Function ResetMatchedFiles() As ObjResult
  Dim sheetService As ServiceSheet
  Set sheetService = New ServiceSheet
  Set ResetMatchedFiles = sheetService.ResetMatchedFileNames()
End Function