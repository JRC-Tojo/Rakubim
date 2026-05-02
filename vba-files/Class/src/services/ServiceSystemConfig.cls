VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServiceSystemConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' .LoadUserInput()で読み取った値を保存しておく
Private cachedUserInput As Dictionary

' ユーザー情報として読み取る情報のKeyとその位置を定義
Private userInputPositions As Dictionary

' ユーザー入力シートの名前
Private Const USER_INPUT_SHEET_NAME As String = "入力"

Private Sub Class_Initialize()
  ' ユーザー入力キャッシュの初期化
  Set cachedUserInput = New Dictionary

  ' ユーザー入力位置の定義
  Set userInputPositions = New Dictionary
  userInputPositions.Add "TargetFolder", "E3"
  userInputPositions.Add "CompanyID", "C6"
  userInputPositions.Add "ProjectID", "C7"
  userInputPositions.Add "StructureID", "C8"
  userInputPositions.Add "DesignID", "C9"
End Sub

' ユーザーがExcelに入力した内容を取得する
' 取得した情報は内部にキャッシュしておき、以降の呼び出しではキャッシュを返すようにする
Public Function LoadUserInput(Optional wb As Workbook = Nothing) As ObjResult
  Dim sheetRepo As RepoSheet
  Set sheetRepo = New RepoSheet

  ' 読み取り対象のブックを指定
  If (Not (wb Is Nothing)) Then
    Call sheetRepo.SetWorkbook(wb)
  End If
  
  ' セル位置から情報を読み取ってキャッシュに格納
  Dim key As Variant
  For Each key In userInputPositions.Keys
    Dim cellAddress As String
    cellAddress = userInputPositions(key)
    Dim res As ObjResult
    Set res = sheetRepo.ReadCell(USER_INPUT_SHEET_NAME, cellAddress, True)
    If Not res.isOK Then
      Set LoadUserInput = res.err
      Exit Function
    End If
    cachedUserInput.Add key, res.value
  Next key

  ' 特殊オブジェクトから値を取得する
  cachedUserInput.Add "UseContinuousNumber", CBool(sheetRepo.ReadObjValue(USER_INPUT_SHEET_NAME, "useContinuousNumberCheck"))

  Set LoadUserInput = Success(Nothing)
End Function


' キャッシュの読み取り
Private Function GetCachedVals(targetKey As String) As Variant
  Debug.Assert IsLoadedCache  ' 読み取り前にMe.LoadUserInput()を呼び出す必要がある
  GetCachedVals = cachedUserInput(targetKey)
End Function

' メッセージの表示言語
Public Property Get MsgLanguage() As String
  MsgLanguage = "JaJp"
End Property

Public Property Get IsLoadedCache() As Boolean
  IsLoadedCache = cachedUserInput.Count > 0
End Property

Public Property Get PdfFolderPath() As ObjPath
  Dim path As ObjPath
  Set path = New ObjPath
  Set PdfFolderPath = path.Init(GetCachedVals("TargetFolder")).Child("PDF")
End Property

Public Property Get CadFolderPath() As ObjPath
  Dim path As ObjPath
  Set path = New ObjPath
  Set CadFolderPath = path.Init(GetCachedVals("TargetFolder")).Child("CAD")
End Property

Public Property Get PdfExtensions() As Variant
  Dim exts As ModelExts: Set exts = New ModelExts
  Set PdfExtensions = exts.InitModel("pdf")
End Property

Public Property Get CadExtensions() As Variant
  Dim exts As ModelExts: Set exts = New ModelExts
  Set CadExtensions = exts.InitModel("dwg", "dxf")
End Property

Public Property Get UseSequenceForPdf() As Boolean
  UseSequenceForPdf = GetCachedVals("UseContinuousNumber")
End Property

Public Property Get UseSequenceForCad() As Boolean
  UseSequenceForCad = False
End Property
