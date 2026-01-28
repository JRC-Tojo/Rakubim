VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Const ReleaseFilePath = "https://jrcjregroup.app.box.com/folder/348645297166"
Private Const ThisVersion = "v1.1.0"

Sub Workbook_Open()
  Call SetPulldown

  SplashScreen.Show

  If (RequireUpdate() = True) Then
    Dim res As Long
    res = MsgBox("「らくびむ」の最新版がリリースされました" & vbCrLf & _ 
            "以下の配布先から最新版をコピーしてご利用ください" & vbCrLf & vbCrLf & _ 
            ReleaseFilePath & vbCrLf & vbCrLf & _ 
            "上記の配布先をブラウザで開きますか？",vbInformation + vbYesNo + vbDefaultButton2, "らくびむ - アップデート -")
    If (res = vbYes) Then
      ActiveWorkbook.FollowHyperlink Address:=ReleaseFilePath, SubAddress:="", NewWindow:=True
    End If
  End If
End Sub

Private Sub SetPulldown()
  Dim matcher As Paper2FileMatcher
  Set matcher = GetMatcher()
  If (matcher Is Nothing) Then Exit Sub

  Dim rowNum As Integer
  rowNum = getLastPaperRows()

  Call SetFilePulldownMenu(matcher, rowNum - LIMITED_SPACE_COUNTER)
End Sub

Private Function RequireUpdate() As Boolean
  Dim latestVersion As String
  On Error GoTo ErrHandler
  latestVersion = LoadLatestGhReleaseVersion("JRC-Chorus", "RakubimPublic")
  If (latestVersion <> ThisVersion) Then
    RequireUpdate = True
  Else
    RequireUpdate = False
  End If
  Exit Function
ErrHandler:
  ' エラーが発生した場合はアップデート不要とみなす
  RequireUpdate = False
End Function

' GitHub ReleasesのAPIにアクセスし、最新のリリースバージョンを取得する
' 引数:
'   accountName (String) - GitHub アカウント名（または組織名）
'   repoName    (String) - リポジトリ名
' 戻り値:
'   String - tag_name（例: "v1.0.2"）
' 例外:
'   HTTP エラーや tag_name が見つからない場合はエラーを発生させる
Private  Function LoadLatestGhReleaseVersion(accountName As String, repoName As String) As String
  Dim http As Object
  Dim resp As String
  Dim url As String
  On Error GoTo ErrHandler

  If Len(Trim$(accountName)) = 0 Or Len(Trim$(repoName)) = 0 Then
    Err.Raise vbObjectError + 540, "LoadLatestGhReleaseVersion", "accountName または repoName が空です。"
  End If

  url = "https://api.github.com/repos/" & accountName & "/" & repoName & "/releases/latest"

  Set http = CreateObject("MSXML2.XMLHTTP")
  http.Open "GET", url, False
  http.setRequestHeader "Accept", "application/vnd.github.v3+json"
  ' http.setRequestHeader "User-Agent", "VBA-JsonUtils"
  http.send

  If http.Status < 200 Or http.Status >= 300 Then
    Err.Raise vbObjectError + 541, "LoadLatestGhReleaseVersion", "HTTP error: " & http.Status & " - " & http.statusText
  End If

  resp = http.responseText
  Set http = Nothing

  ' 正規表現で "tag_name": "値" を抽出（簡易）
  Dim re As Object
  Dim matches As Object
  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = """tag_name""\s*:\s*""([^""]+)"""
  re.IgnoreCase = True
  re.Global = False

  Set matches = re.Execute(resp)
  If matches.Count > 0 Then
    LoadLatestGhReleaseVersion = matches(0).SubMatches(0)
  Else
    Err.Raise vbObjectError + 542, "LoadLatestGhReleaseVersion", "tag_name がレスポンスに見つかりませんでした。"
  End If

  Exit Function
ErrHandler:
  Err.Raise Err.Number, "LoadLatestGhReleaseVersion", Err.Description
End Function