VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelExts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' 正規化された拡張子群
Private inner_NormalizedExtensions As ObjSet

' ModelExts の初期化
' 引数: extensions (ObjSet of String) - 正規化された拡張子のセット
' 戻り値: なし（内部にデータを保持）
Public Function InitModel(ParamArray extensions() As Variant) As ModelExts
  Dim exts As ObjSet: Set exts = New ObjSet
  Call exts.InitItemsFromArray(extensions)
  Set inner_NormalizedExtensions = normalizeExtensions(exts)
  Set InitModel = Me
End Function

' 正規化された拡張子のセットを返す
Public Property Get ExtSet() As ObjSet
  Set ExtSet = inner_NormalizedExtensions
End Property

' 与えられた集合の文字をすべて正規化する
' 引数: extensions (ObjSet of String) - 正規化する前の拡張子のセット
' 戻り値: ObjSet of String - 正規化された拡張子のセット
Private Function normalizeExtensions(extensions As ObjSet) As ObjSet
  Dim normalized As ObjSet
  Set normalized = New ObjSet
  
  Dim ext As Variant
  Dim extStr As String
  For Each ext In extensions
    On Error Resume Next
    extStr = CStr(ext)
    If Err.Number = 0 Then
      extStr = normExtFromString(extStr)
      If extStr <> "" Then Call normalized.Append(extStr)
    End If
  Next ext
  
  Set normalizeExtensions = normalized
End Function


' 与えられた文字列を拡張子として正しい形式に変換して返す
' 変換に失敗した場合は空文字列を返す
Private Function normExtFromString(targetExt As String) As String
  ' 失敗時の戻り値を定義
  normExtFromString = ""
  
  ' 失敗ケースを列挙
  If Len(targetExt) = 0 Then Exit Function
  
  ' 正規化
  Dim normedExt As String
  normedExt = LCase(Trim(targetExt))
  If Left(normedExt, 1) = "." Then normedExt = Mid(normedExt, 2)

  normExtFromString = normedExt
End Function