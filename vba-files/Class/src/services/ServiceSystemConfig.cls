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
Private cachedUserInput As ObjDict

' ユーザー情報として読み取る情報のKeyとその位置を定義
Private userInputPositions As ObjDict

' ユーザー入力シートの名前
Private Const USER_INPUT_SHEET_NAME As String = "入力"

Private Sub Class_Initialize()
  ' ユーザー入力キャッシュの初期化
  Set cachedUserInput = New ObjDict

  ' ユーザー入力位置の定義
  Set userInputPositions = New ObjDict
  Call userInputPositions.Append("TargetFolder", "E3")
  Call userInputPositions.Append("CompanyID", "C6")
  Call userInputPositions.Append("ProjectID", "C7")
  Call userInputPositions.Append("StructureID", "C8")
  Call userInputPositions.Append("DesignID", "C9")
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
  Dim keyVal As ModelKeyValue
  For Each keyVal In userInputPositions
    Dim res As ObjResult
    Set res = sheetRepo.ReadCell(USER_INPUT_SHEET_NAME, keyVal.Value, True)
    If Not res.isOK Then
      Set LoadUserInput = res.err
      Exit Function
    End If
    Call cachedUserInput.Append(keyVal.Key, res.value)
  Next

  ' 特殊オブジェクトから値を取得する
  Call cachedUserInput.Append("UseContinuousNumber", CBool(sheetRepo.ReadObjValue(USER_INPUT_SHEET_NAME, "useContinuousNumberCheck")))
  Call cachedUserInput.Update("CompanyID", getCompanyID(cachedUserInput.At("CompanyID")))

  Set LoadUserInput = Success(Nothing)
End Function


' 選択中の業務機関表示から番号を取得する
Private Function getCompanyID(companyName As String)
  ' 不明な場合は000を返す
  getCompanyID = "000"
  
  ' `XXX (000)`の形式のため、()内の数字をコードとして返す
  Dim regEx As ObjRegex
  Set regEx = New ObjRegex
  Dim matchedNums As ObjList
  Set matchedNums = regEx.GetMatches("\((\d+)\)", companyName)
  If (matchedNums.Count <> 1) Then Exit Function

  ' キャプチャグループから括弧内の数字を抽出
  Dim match As Match
  Set match = matchedNums.At(0)
  getCompanyID = Format(CLng(match.SubMatches(0)), "000")
End Function


' キャッシュの読み取り
Private Function GetCachedVals(targetKey As String) As Variant
  Debug.Assert IsLoadedCache  ' 読み取り前にMe.LoadUserInput()を呼び出す必要がある
  GetCachedVals = cachedUserInput.At(targetKey)
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

Public Property Get PdfExtensions() As ModelExts
  Dim exts As ModelExts: Set exts = New ModelExts
  Set PdfExtensions = exts.InitModel("pdf")
End Property

Public Property Get CadExtensions() As ModelExts
  Dim exts As ModelExts: Set exts = New ModelExts
  Set CadExtensions = exts.InitModel("dwg", "dxf")
End Property

Public Property Get UseSequenceForPdf() As Boolean
  UseSequenceForPdf = GetCachedVals("UseContinuousNumber")
End Property

Public Property Get UseSequenceForCad() As Boolean
  UseSequenceForCad = False
End Property

Public Property Get CompanyId() As String
  CompanyId = GetCachedVals("CompanyID")
End Property

Public Property Get ProjectId() As String
  ProjectId = GetCachedVals("ProjectID")
End Property

Public Property Get StructureId() As String
  StructureId = GetCachedVals("StructureID")
End Property

Public Property Get DesignId() As String
  DesignId = GetCachedVals("DesignID")
End Property

Public Property Get RelationalTableOutputPath() As ObjPath
  Set RelationalTableOutputPath =  Me.CadFolderPath.Child(Me.CompanyId & "-" & Me.ProjectId & ".xls")
End Property